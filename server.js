'use strict';

/**
 * Maiershirts – Lexware/Lexoffice Angebots-Tool
 * Features (wie vorher, NICHT abgespeckt):
 * - public/ statisch + index.html auf /
 * - templates/ statisch + stabiler Alias /templates/lexware_template.xlsx
 * - /api/ping
 * - /api/test-excel (nur validieren + Summary)
 * - /api/create-offer (Angebot erstellen + finalize)
 * - /api/download-pdf (PDF via /v1/quotations/{id}/file)
 * - /api/articles (Artikel-Lookup für Template/Debug – mit Cache + Paging)
 * - Passwortschutz: TOOL_PASSWORD (Body) ODER APP_USER/APP_PASS (Basic Auth)
 * - RateLimit: 2 req/s -> TokenBucket + Retry bei 429 (Backoff + Jitter)
 * - Idempotency: verhindert 3x Erstellung bei Doppelklick/Retry
 * - IMMER technical im Fehlerfall (httpStatus + raw + meta + request/traceId)
 */

require('dotenv').config();

const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const express = require('express');
const axios = require('axios');
const XLSX = require('xlsx');

const app = express();
app.use(express.json({ limit: '20mb' }));

// -----------------------------
// ENV
// -----------------------------
const API_KEY = process.env.LEXOFFICE_API_KEY || process.env.LEXWARE_API_KEY || '';
const API_BASE_URL = (process.env.LEXOFFICE_API_BASE_URL || process.env.LEXWARE_API_BASE_URL || 'https://api.lexware.io')
  .replace(/\/+$/, '');

const FINALIZE_DEFAULT = (process.env.FINALIZE_DEFAULT || 'true').toLowerCase() === 'true';

// Passwortschutz (2 Varianten, beide bleiben drin):
// A) TOOL_PASSWORD: wird im Body mitgeschickt (Frontend Passwortfeld)
// B) APP_USER/APP_PASS: HTTP Basic Auth (Render Env: APP_USER/APP_PASS)
const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';
const APP_USER = process.env.APP_USER || '';
const APP_PASS = process.env.APP_PASS || '';
const ALLOW_REMOTE = (process.env.ALLOW_REMOTE || 'true').toLowerCase() === 'true';

const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE_DEFAULT || process.env.ALLOW_PRICE_OVERRIDE || 'false').toLowerCase() === 'true';

// Optional: zusätzliches Mindestintervall (falls gewünscht)
const MIN_INTERVAL_MS = Number(process.env.LEXWARE_MIN_INTERVAL_MS || '0');

// -----------------------------
// Static files
// -----------------------------
app.use(express.static(path.join(__dirname, 'public')));

// Variante B: /templates aus Repo-Ordner "templates"
app.use('/templates', express.static(path.join(__dirname, 'templates')));

// Stabiler Alias-Link (Button nutzt IMMER diese URL):
// /templates/lexware_template.xlsx -> templates/Lexware_Template.xlsx
app.get('/templates/lexware_template.xlsx', (req, res) => {
  const filePath = path.join(__dirname, 'templates', 'Lexware_Template.xlsx');
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('Template-Datei nicht gefunden. Erwartet: templates/Lexware_Template.xlsx');
  }
  res.sendFile(filePath);
});

// Root sicher bedienen (gegen "Cannot GET /")
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// -----------------------------
// Helpers: technical payload (bombensicher)
// -----------------------------
function safeJson(v) {
  try {
    if (v === undefined) return null;
    return JSON.parse(JSON.stringify(v));
  } catch {
    try {
      return String(v);
    } catch {
      return null;
    }
  }
}

function extractMeta(raw) {
  const r = raw || {};
  return {
    timestamp: r.timestamp || null,
    status: r.status || null,
    error: r.error || null,
    message: r.message || null,
    path: r.path || null,
    traceId: r.traceId || null,
    requestId: r.requestId || null,
    details: Array.isArray(r.details) ? r.details : null
  };
}

function buildTechnical({ httpStatus, raw, err }) {
  const rawSafe = safeJson(raw);
  return {
    httpStatus: httpStatus ?? null,
    raw: rawSafe,
    meta: extractMeta(rawSafe),
    errorMessage: err?.message || null
  };
}

function ok(res, payload) {
  return res.json({ ok: true, ...payload });
}

function fail(res, payload) {
  // garantieren, dass "technical" existiert, wenn httpStatus/err/raw da sind
  if (!payload.technical && (payload.httpStatus || payload.raw || payload.err)) {
    payload.technical = buildTechnical({ httpStatus: payload.httpStatus, raw: payload.raw, err: payload.err });
    delete payload.httpStatus;
    delete payload.raw;
    delete payload.err;
  }
  return res.json({ ok: false, ...payload });
}

// -----------------------------
// Auth: Basic Auth (APP_USER/APP_PASS) optional
// -----------------------------
function parseBasicAuth(header) {
  if (!header || !header.startsWith('Basic ')) return null;
  const b64 = header.slice(6);
  try {
    const decoded = Buffer.from(b64, 'base64').toString('utf8');
    const idx = decoded.indexOf(':');
    if (idx < 0) return null;
    return { user: decoded.slice(0, idx), pass: decoded.slice(idx + 1) };
  } catch {
    return null;
  }
}

function basicAuthMiddleware(req, res, next) {
  // Wenn APP_USER/PASS gesetzt sind, schützen wir (wenn ALLOW_REMOTE nicht explizit "false" – sonst streng)
  if (!APP_USER || !APP_PASS) return next();

  const auth = parseBasicAuth(req.headers.authorization);
  if (auth && auth.user === APP_USER && auth.pass === APP_PASS) return next();

  res.setHeader('WWW-Authenticate', 'Basic realm="Maiershirts Tool"');
  return res.status(401).send('Auth required');
}

// TOOL_PASSWORD Middleware (Body password)
function toolPasswordMiddleware(req, res, next) {
  if (!TOOL_PASSWORD) return next();
  const supplied = req.body?.password || req.query?.password || req.headers['x-tool-password'];
  if (supplied === TOOL_PASSWORD) return next();

  return fail(res, {
    stage: 'auth',
    status: 'UNAUTHORIZED',
    message: 'Passwort ungültig oder fehlt.',
    technical: buildTechnical({ httpStatus: 401, raw: { message: 'UNAUTHORIZED' } })
  });
}

// Kombinierter Schutz:
// - Wenn Basic Auth aktiv: reicht Basic Auth (Frontend Passwortfeld darf dann leer bleiben).
// - Wenn kein Basic Auth aktiv: TOOL_PASSWORD greift (falls gesetzt).
function authMiddleware(req, res, next) {
  if (APP_USER && APP_PASS) return basicAuthMiddleware(req, res, next);
  return toolPasswordMiddleware(req, res, next);
}

// -----------------------------
// Rate Limit: TokenBucket (2 req/s) + optional MIN_INTERVAL_MS
// -----------------------------
class TokenBucket {
  constructor({ capacity, refillPerSec }) {
    this.capacity = capacity;
    this.refillPerSec = refillPerSec;
    this.tokens = capacity;
    this.lastRefill = Date.now();
    this.queue = [];
    this.timer = null;
  }

  _refill() {
    const now = Date.now();
    const elapsed = (now - this.lastRefill) / 1000;
    if (elapsed <= 0) return;
    const add = elapsed * this.refillPerSec;
    this.tokens = Math.min(this.capacity, this.tokens + add);
    this.lastRefill = now;
  }

  async acquire() {
    return new Promise((resolve) => {
      this.queue.push(resolve);
      this._drain();
    });
  }

  _drain() {
    if (this.timer) {
      clearTimeout(this.timer);
      this.timer = null;
    }

    this._refill();

    while (this.tokens >= 1 && this.queue.length) {
      this.tokens -= 1;
      const resolve = this.queue.shift();
      resolve();
    }

    if (this.queue.length) {
      // nächster Versuch in ~100ms
      this.timer = setTimeout(() => this._drain(), 100);
    }
  }
}

const bucket = new TokenBucket({ capacity: 2, refillPerSec: 2 }); // 2 req/s

let lastCallTs = 0;
async function enforceMinInterval() {
  if (!MIN_INTERVAL_MS) return;
  const now = Date.now();
  const diff = now - lastCallTs;
  if (diff < MIN_INTERVAL_MS) {
    await new Promise(r => setTimeout(r, MIN_INTERVAL_MS - diff));
  }
  lastCallTs = Date.now();
}

async function lexwareRequest({ method, url, headers, data, responseType, accept }) {
  // Retry bei 429 empfohlen: exponential backoff + jitter
  const maxRetries = 5;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    await bucket.acquire();
    await enforceMinInterval();

    const res = await axios({
      method,
      url,
      headers: {
        Authorization: `Bearer ${API_KEY}`,
        Accept: accept || 'application/json',
        ...(headers || {})
      },
      data,
      responseType: responseType || 'json',
      validateStatus: () => true
    });

    if (res.status !== 429) return res;

    // 429: Backoff
    const base = 800 * Math.pow(2, attempt);
    const jitter = Math.floor(Math.random() * 250);
    const wait = Math.min(12000, base + jitter);
    await new Promise(r => setTimeout(r, wait));
  }

  // falls nach Retries immer noch 429: letzten Versuch (künstlich) zurück
  return { status: 429, data: { message: 'Rate limit exceeded (client retries exhausted)' } };
}

// -----------------------------
// Excel parsing helpers (wie vorher: vertikal/KeyValue + horizontal)
// -----------------------------
function sheetToJson(wb, name) {
  const sh = wb.Sheets[name];
  return sh ? XLSX.utils.sheet_to_json(sh, { defval: '' }) : null;
}

function sheetRowsToKeyValueObject(rows) {
  if (!rows || !rows.length) return null;

  const first = rows[0];
  const fieldKeys = ['Feld', 'feld', 'Field', 'field'];
  const valueKeys = ['Wert', 'wert', 'Value', 'value', 'val'];

  const fieldCol = fieldKeys.find(k => k in first);
  const valueCol = valueKeys.find(k => k in first);
  if (!fieldCol || !valueCol) return null;

  const obj = {};
  for (const r of rows) {
    const key = String(r[fieldCol] || '').trim();
    if (!key) continue;
    obj[key] = r[valueCol];
  }
  return obj;
}

function numOrNull(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function toLowerTrim(v) {
  return String(v || '').trim().toLowerCase();
}

// -----------------------------
// Artikel Cache (für material/service: Preis + Titel aus Stammdaten)
// -----------------------------
const articleCache = {
  byId: new Map(),
  list: null,
  fetchedAt: 0,
  ttlMs: 10 * 60 * 1000
};

async function getArticleById(articleId) {
  if (!articleId) return null;
  const cached = articleCache.byId.get(articleId);
  if (cached && (Date.now() - cached._ts) < articleCache.ttlMs) return cached.data;

  const res = await lexwareRequest({
    method: 'GET',
    url: `${API_BASE_URL}/v1/articles/${encodeURIComponent(articleId)}`
  });

  if (res.status >= 200 && res.status < 300) {
    articleCache.byId.set(articleId, { _ts: Date.now(), data: res.data });
    return res.data;
  }
  return null;
}

async function listAllArticlesCached() {
  if (articleCache.list && (Date.now() - articleCache.fetchedAt) < articleCache.ttlMs) {
    return articleCache.list;
  }

  const all = [];
  let page = 0;
  const size = 250;

  while (true) {
    const res = await lexwareRequest({
      method: 'GET',
      url: `${API_BASE_URL}/v1/articles?page=${page}&size=${size}`
    });

    if (!(res.status >= 200 && res.status < 300) || !res.data) break;

    const content = Array.isArray(res.data.content) ? res.data.content : [];
    all.push(...content);

    if (res.data.last === true) break;
    if (content.length === 0) break;
    page++;
    if (page > 100) break; // Safety
  }

  articleCache.list = all;
  articleCache.fetchedAt = Date.now();
  return all;
}

// -----------------------------
// Build Quotation Payload (wie vorher, inkl. Regeln)
// -----------------------------
async function parseExcelAndBuildQuotationPayload(excelBase64, { allowPriceOverride }) {
  const errors = [];
  const warnings = [];
  const autoNamedLineItems = [];
  const byType = {};

  const wb = XLSX.read(Buffer.from(excelBase64, 'base64'), { type: 'buffer' });

  const angebotRows = sheetToJson(wb, 'Angebot');
  const kundeRows = sheetToJson(wb, 'Kunde');
  const posRows = sheetToJson(wb, 'Positionen');

  // Pflicht-Sheets
  if (!angebotRows) errors.push({ sheet: 'Angebot', message: 'Sheet „Angebot“ fehlt.' });
  if (!kundeRows) errors.push({ sheet: 'Kunde', message: 'Sheet „Kunde“ fehlt.' });
  if (!posRows) errors.push({ sheet: 'Positionen', message: 'Sheet „Positionen“ fehlt.' });

  if (errors.length) {
    return { ok: false, payload: null, summary: { errors, warnings } };
  }

  // Angebot: vertikal oder horizontal
  const angebot = sheetRowsToKeyValueObject(angebotRows) || angebotRows[0] || {};
  const taxType = String(angebot.taxType || angebot.TaxType || angebot.TAXTYPE || '').trim();

  if (!taxType) {
    errors.push({ sheet: 'Angebot', row: 2, field: 'taxType', message: 'taxType ist Pflicht (z. B. „gross“ oder „net“).' });
  }

  // Kunde: vertikal oder horizontal
  const kunde = sheetRowsToKeyValueObject(kundeRows) || kundeRows[0] || {};
  const customerName = String(kunde.name || kunde.Name || '').trim();

  if (!customerName) {
    errors.push({ sheet: 'Kunde', row: 2, field: 'name', message: 'Kundenname ist Pflicht.' });
  }

  // Address (KontaktId optional – wie vorher: wenn Felder gesetzt, wird address.name genutzt)
  // Hinweis: Lexware akzeptiert address.contactId auch – wir lassen das als optionales Feld zu:
  const address = {
    name: customerName || undefined,
    contactId: String(kunde.contactId || '').trim() || undefined,
    street: String(kunde.street || '').trim() || undefined,
    zip: String(kunde.zip || '').trim() || undefined,
    city: String(kunde.city || '').trim() || undefined,
    countryCode: String(kunde.countryCode || 'DE').trim() || 'DE',
    contactPerson: String(kunde.contactPerson || '').trim() || undefined,
    email: String(kunde.email || '').trim() || undefined,
    phone: String(kunde.phone || '').trim() || undefined,
  };

  // voucherDate muss ISO datetime sein
  const voucherDate = new Date().toISOString();
  const expirationDate = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString();

  // taxConditions Pflicht (nicht null)
  const taxConditions = taxType ? { taxType } : null;
  if (!taxConditions) {
    errors.push({ sheet: 'Angebot', row: 2, field: 'taxConditions', message: 'taxConditions.taxType ist Pflicht.' });
  }

  // Positionen -> lineItems
  const lineItems = [];
  for (let i = 0; i < posRows.length; i++) {
    const row = posRows[i];
    const excelRow = i + 2;

    const type = toLowerTrim(row.type);
    const articleId = String(row.articleId || row.articleID || '').trim();
    let name = String(row.name || '').trim();
    const description = String(row.description || '').trim();

    const qty = numOrNull(row.quantity ?? row.qty ?? row.Qty ?? row.Menge ?? row.menge);
    const unitName = String(row.unitName || row.unit || '').trim();

    const unitPriceAmount = numOrNull(row.unitPriceAmount ?? row.price ?? row.Preis ?? row.unitPrice);
    const taxRatePercentage = numOrNull(row.taxRatePercentage ?? row.taxRate ?? row.tax);
    const discountPercent = numOrNull(row.discountPercent ?? row.discount);

    // Leere Zeilen überspringen
    const hasAny = type || articleId || name || description || qty || unitName || unitPriceAmount || taxRatePercentage || discountPercent;
    if (!hasAny) continue;

    // Pflicht: type
    if (!type) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'type', message: 'type ist Pflicht.' });
      continue;
    }
    byType[type] = (byType[type] || 0) + 1;

    // text: kein qty nötig
    if (type === 'text') {
      const txtName = name || description || `Hinweis ${excelRow}`;
      lineItems.push({
        type: 'text',
        name: txtName,
        description: description || undefined
      });
      continue;
    }

    // Pflicht: qty > 0
    if (!(qty > 0)) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'qty', message: 'qty muss größer als 0 sein.' });
      continue;
    }

    // unitName Pflicht (für nicht-text)
    if (!unitName) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'unitName', message: 'unitName ist Pflicht.' });
      continue;
    }

    // material: articleId Pflicht, Preis IMMER aus Lexware, Excel-Preis ignorieren, Name aus Artikel wenn leer
    // service: ähnlich (wenn du Services als Artikel pflegst)
    if ((type === 'material' || type === 'service') && !articleId) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'articleId', message: 'articleId ist Pflicht bei type=material/service.' });
      continue;
    }

    // Name automatisch ergänzen, wenn leer + articleId gesetzt
    if (!name && articleId) {
      // Vorher: Name aus Lexware ergänzen (wir machen das wieder)
      const art = await getArticleById(articleId);
      const autoName = art?.title || `Artikel ${articleId}`;
      name = autoName;
      autoNamedLineItems.push({ row: excelRow, articleId, name: autoName });
      warnings.push({ sheet: 'Positionen', row: excelRow, message: `Name war leer → automatisch aus Artikel ergänzt: "${autoName}".` });
    }

    if (!name) {
      name = `Position ${excelRow}`;
      autoNamedLineItems.push({ row: excelRow, articleId: articleId || null, name });
      warnings.push({ sheet: 'Positionen', row: excelRow, message: `Name war leer → automatisch gesetzt: "${name}".` });
    }

    const item = {
      type,              // material | service | custom
      name,
      description: description || undefined,
      quantity: qty,
      unitName
    };

    // WICHTIG: material/service erwartet "id" = article uuid
    if (articleId && (type === 'material' || type === 'service')) {
      item.id = articleId;
      // Preis NICHT setzen (kommt aus Artikelstamm)
    } else {
      // custom ohne articleId: Preis Pflicht
      // Preis nur Pflicht, wenn KEINE articleId gesetzt ist
      if (!articleId && unitPriceAmount === null) {
        errors.push({ sheet: 'Positionen', row: excelRow, field: 'unitPriceAmount', message: 'Preis ist Pflicht, wenn keine articleId gesetzt ist.' });
        continue;
      }

      // Wenn articleId gesetzt und allowPriceOverride aktiv: Preis darf überschrieben werden (optional)
      // (wie vorher: Schalter)
      const shouldSendPrice =
        (!articleId) ||
        (allowPriceOverride === true);

      if (shouldSendPrice && unitPriceAmount !== null) {
        const rate = taxRatePercentage !== null ? taxRatePercentage : 19;
        if (taxType === 'gross') {
          item.unitPrice = { currency: 'EUR', grossAmount: unitPriceAmount, taxRatePercentage: rate };
        } else {
          item.unitPrice = { currency: 'EUR', netAmount: unitPriceAmount, taxRatePercentage: rate };
        }
      }
    }

    if (discountPercent !== null) {
      item.discountPercentage = discountPercent;
    }

    lineItems.push(item);
  }

  if (!lineItems.length) {
    errors.push({ sheet: 'Positionen', message: 'Keine Positionen gefunden.' });
  }

  const summary = {
    errors,
    warnings,
    byType,
    autoNamedLineItems,
    allowPriceOverrideUsed: !!allowPriceOverride,
    voucherDate,
    taxType
  };

  if (errors.length) {
    return { ok: false, payload: null, summary };
  }

  // Wie vorher: totalPrice + shippingConditions mitgeben (robust)
  const payload = {
    voucherDate,
    expirationDate,
    address,
    lineItems,
    totalPrice: { currency: 'EUR' },
    taxConditions,
    shippingConditions: { shippingType: 'none' }
  };

  return { ok: true, payload, summary };
}

// -----------------------------
// Idempotency (gegen 3x Erstellung)
// -----------------------------
const inFlight = new Map();

function hashRequest({ excelData, allowPriceOverride, finalize }) {
  return crypto
    .createHash('sha256')
    .update(String(finalize ? '1' : '0'))
    .update('|')
    .update(String(allowPriceOverride ? '1' : '0'))
    .update('|')
    .update(excelData)
    .digest('hex');
}

// -----------------------------
// API endpoints
// -----------------------------
app.get('/api/ping', (req, res) => {
  ok(res, {
    status: 'OK',
    passwordProtected: !!(TOOL_PASSWORD || (APP_USER && APP_PASS)),
    passwordMode: (APP_USER && APP_PASS) ? 'basic' : (TOOL_PASSWORD ? 'toolPassword' : 'none'),
    allowPriceOverrideDefault: ALLOW_PRICE_OVERRIDE_DEFAULT,
    minIntervalMs: MIN_INTERVAL_MS,
    apiBaseUrl: API_BASE_URL,
    finalizeDefault: FINALIZE_DEFAULT,
    allowRemote: ALLOW_REMOTE
  });
});

// Artikel-Lookup (für Template/Debug)
app.get('/api/articles', authMiddleware, async (req, res) => {
  try {
    if (!API_KEY) {
      return fail(res, { stage: 'config', status: 'CONFIG_ERROR', message: 'API Key fehlt.', technical: buildTechnical({ httpStatus: 500, raw: { message: 'NO_API_KEY' } }) });
    }
    const list = await listAllArticlesCached();
    ok(res, { status: 'SUCCESS', data: { count: list.length, articles: list } });
  } catch (err) {
    fail(res, { stage: 'articles', status: 'ERROR', message: err.message, technical: buildTechnical({ httpStatus: 500, raw: { message: 'ARTICLES_EXCEPTION' }, err }) });
  }
});

// Excel Test
app.post('/api/test-excel', authMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body || {};
    const allow = typeof allowPriceOverride === 'boolean' ? allowPriceOverride : ALLOW_PRICE_OVERRIDE_DEFAULT;

    if (!excelData) {
      return fail(res, { stage: 'input', status: 'VALIDATION_ERROR', message: 'Keine Excel-Daten übergeben.', technical: buildTechnical({ httpStatus: 400, raw: { message: 'NO_EXCEL' } }) });
    }

    const parsed = await parseExcelAndBuildQuotationPayload(excelData, { allowPriceOverride: allow });
    if (!parsed.ok) {
      return fail(res, {
        stage: 'validation',
        status: 'VALIDATION_ERROR',
        message: 'Excel enthält Validierungsfehler. Details siehe errors.',
        data: { summary: parsed.summary }
      });
    }

    ok(res, {
      stage: 'test',
      status: 'SUCCESS',
      message: 'Test erfolgreich — keine kritischen Fehler.',
      data: { summary: parsed.summary }
    });
  } catch (err) {
    fail(res, { stage: 'test', status: 'ERROR', message: err.message, technical: buildTechnical({ httpStatus: 500, raw: { message: 'TEST_EXCEPTION' }, err }) });
  }
});

// Angebot erstellen (+ finalize)
app.post('/api/create-offer', authMiddleware, async (req, res) => {
  const startedAt = Date.now();
  try {
    const { excelData, allowPriceOverride, finalize } = req.body || {};
    const allow = typeof allowPriceOverride === 'boolean' ? allowPriceOverride : ALLOW_PRICE_OVERRIDE_DEFAULT;
    const doFinalize = typeof finalize === 'boolean' ? finalize : FINALIZE_DEFAULT;

    if (!excelData) {
      return fail(res, { stage: 'input', status: 'VALIDATION_ERROR', message: 'Keine Excel-Daten übergeben.', technical: buildTechnical({ httpStatus: 400, raw: { message: 'NO_EXCEL' } }) });
    }
    if (!API_KEY) {
      return fail(res, { stage: 'config', status: 'CONFIG_ERROR', message: 'API Key fehlt.', technical: buildTechnical({ httpStatus: 500, raw: { message: 'NO_API_KEY' } }) });
    }

    const key = hashRequest({ excelData, allowPriceOverride: allow, finalize: doFinalize });
    if (inFlight.has(key)) {
      const cached = await inFlight.get(key);
      return res.json(cached);
    }

    const promise = (async () => {
      const parsed = await parseExcelAndBuildQuotationPayload(excelData, { allowPriceOverride: allow });
      if (!parsed.ok) {
        return {
          ok: false,
          stage: 'validation',
          status: 'VALIDATION_ERROR',
          message: 'Excel enthält Validierungsfehler. Details siehe errors.',
          data: { summary: parsed.summary }
        };
      }

      const url = `${API_BASE_URL}/v1/quotations${doFinalize ? '?finalize=true' : ''}`;

      const apiRes = await lexwareRequest({
        method: 'POST',
        url,
        headers: { 'Content-Type': 'application/json' },
        data: parsed.payload,
        accept: 'application/json'
      });

      // Fehlerfälle IMMER mit technical
      if (apiRes.status < 200 || apiRes.status >= 300) {
        return {
          ok: false,
          stage: 'lexware-create',
          status: apiRes.status === 429 ? 'RATE_LIMIT' : 'ERROR',
          message: apiRes.status === 429 ? 'Rate limit exceeded' : 'Lexware API Fehler',
          technical: buildTechnical({ httpStatus: apiRes.status, raw: apiRes.data }),
          data: { summary: parsed.summary }
        };
      }

      return {
        ok: true,
        stage: 'lexware-create',
        status: 'SUCCESS',
        message: 'Angebot erstellt.',
        data: {
          quotationId: apiRes.data?.id || null,
          summary: parsed.summary,
          ms: Date.now() - startedAt
        }
      };
    })();

    inFlight.set(key, promise);
    const result = await promise;
    // Cache kurz halten (gegen Doppelklick)
    setTimeout(() => inFlight.delete(key), 15000);

    return res.json(result);
  } catch (err) {
    return fail(res, { stage: 'lexware-create', status: 'ERROR', message: err.message || 'Unerwarteter Fehler', technical: buildTechnical({ httpStatus: 500, raw: { message: 'UNHANDLED_EXCEPTION' }, err }) });
  }
});

// PDF Download (wie vorher) – direkt /v1/quotations/{id}/file
app.get('/api/download-pdf', authMiddleware, async (req, res) => {
  try {
    const quotationId = String(req.query.id || '').trim();
    if (!quotationId) {
      return res.status(400).send('Missing id');
    }
    if (!API_KEY) {
      return res.status(500).send('API Key fehlt');
    }

    const apiRes = await lexwareRequest({
      method: 'GET',
      url: `${API_BASE_URL}/v1/quotations/${encodeURIComponent(quotationId)}/file`,
      responseType: 'arraybuffer',
      accept: '*/*'
    });

    if (apiRes.status < 200 || apiRes.status >= 300) {
      // als JSON Fehler zurückgeben (damit UI es protokollieren kann)
      return res.status(200).json({
        ok: false,
        stage: 'lexware-pdf',
        status: apiRes.status === 429 ? 'RATE_LIMIT' : 'ERROR',
        message: apiRes.status === 429 ? 'Rate limit exceeded' : 'Lexware PDF Fehler',
        technical: buildTechnical({ httpStatus: apiRes.status, raw: apiRes.data })
      });
    }

    const contentType = apiRes.headers['content-type'] || 'application/pdf';
    const disposition = apiRes.headers['content-disposition'] || 'attachment; filename="quotation.pdf"';

    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Disposition', disposition);
    return res.status(200).send(Buffer.from(apiRes.data));
  } catch (err) {
    return res.status(200).json({
      ok: false,
      stage: 'lexware-pdf',
      status: 'ERROR',
      message: err.message,
      technical: buildTechnical({ httpStatus: 500, raw: { message: 'PDF_EXCEPTION' }, err })
    });
  }
});

// -----------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server läuft auf Port', PORT));
