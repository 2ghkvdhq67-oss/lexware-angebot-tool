'use strict';

require('dotenv').config();

const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const express = require('express');
const axios = require('axios');
const XLSX = require('xlsx');

// ✅ SAFEST FIX: Parser optional laden (sonst crasht Render beim Start)
let parseLines = null;
try {
  // wichtig: Pfad + Groß/Kleinschreibung muss exakt stimmen auf Linux (Render)
  ({ parseLines } = require('./lib/parseText'));
} catch (e) {
  console.error('[BOOT] Optional parser module ./lib/parseText konnte nicht geladen werden.');
  console.error('[BOOT] Ursache:', e?.message || e);
  // Server läuft weiter, nur Phase-1 Endpoints liefern CONFIG_ERROR
  parseLines = null;
}

const app = express();
app.use(express.json({ limit: '25mb' }));

// ------------------------------------------------------------
// ENV
// ------------------------------------------------------------
const API_KEY = process.env.LEXOFFICE_API_KEY || process.env.LEXWARE_API_KEY || '';
const API_BASE_URL = (process.env.LEXOFFICE_API_BASE_URL || process.env.LEXWARE_API_BASE_URL || 'https://api.lexware.io')
  .replace(/\/+$/, '');

const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';
const APP_USER = process.env.APP_USER || '';
const APP_PASS = process.env.APP_PASS || '';

const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE_DEFAULT || process.env.ALLOW_PRICE_OVERRIDE || 'false').toLowerCase() === 'true';

const FINALIZE_DEFAULT = (process.env.FINALIZE_DEFAULT || 'true').toLowerCase() === 'true';

// optional extra buffer interval in ms (zusätzlich zum TokenBucket)
const MIN_INTERVAL_MS = Number(process.env.LEXWARE_MIN_INTERVAL_MS || '0');

// TTL für dynamisches Template: 10 Minuten
const TEMPLATE_TTL_MS = 10 * 60 * 1000;

// ------------------------------------------------------------
// Static: public + templates (statisch bleibt erhalten!)
// ------------------------------------------------------------
app.use(express.static(path.join(__dirname, 'public')));
app.use('/templates', express.static(path.join(__dirname, 'templates')));

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// Stabiler Alias: /templates/lexware_template.xlsx (statisch)
app.get('/templates/lexware_template.xlsx', (req, res) => {
  const templatesDir = path.join(__dirname, 'templates');
  const preferred = [
    'Lexware_Template.xlsx',
    'lexware_template.xlsx',
    'Lexware_Template.XLSX',
    'lexware_template.XLSX'
  ].map(n => path.join(templatesDir, n));

  let filePath = preferred.find(p => fs.existsSync(p));

  if (!filePath) {
    try {
      const candidates = fs.existsSync(templatesDir)
        ? fs.readdirSync(templatesDir).filter(f => f.toLowerCase().endsWith('.xlsx'))
        : [];
      if (candidates.length) filePath = path.join(templatesDir, candidates[0]);
    } catch { /* ignore */ }
  }

  if (!filePath || !fs.existsSync(filePath)) {
    return res.status(404).send('Template-Datei nicht gefunden. Lege eine .xlsx in /templates ab.');
  }

  return res.sendFile(filePath);
});

// ------------------------------------------------------------
// Helpers: technical payload (bombensicher)
// ------------------------------------------------------------
function safeJson(v) {
  try {
    if (v === undefined) return null;
    return JSON.parse(JSON.stringify(v));
  } catch {
    try { return String(v); } catch { return null; }
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
  if (!payload.technical && (payload.httpStatus || payload.raw || payload.err)) {
    payload.technical = buildTechnical({ httpStatus: payload.httpStatus, raw: payload.raw, err: payload.err });
    delete payload.httpStatus;
    delete payload.raw;
    delete payload.err;
  }
  return res.json({ ok: false, ...payload });
}

// ------------------------------------------------------------
// Auth: Basic Auth optional ODER TOOL_PASSWORD im Body/Query/Header
// ------------------------------------------------------------
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
  if (!APP_USER || !APP_PASS) return next();
  const auth = parseBasicAuth(req.headers.authorization);
  if (auth && auth.user === APP_USER && auth.pass === APP_PASS) return next();
  res.setHeader('WWW-Authenticate', 'Basic realm="Maiershirts Tool"');
  return res.status(401).send('Auth required');
}

function toolPasswordMiddleware(req, res, next) {
  if (!TOOL_PASSWORD) return next();

  // wichtig: GET Downloads können kein Body-Passwort schicken -> Header/Query erlaubt
  const supplied =
    req.body?.password ||
    req.query?.password ||
    req.headers['x-tool-password'];

  if (supplied === TOOL_PASSWORD) return next();

  return fail(res, {
    stage: 'auth',
    status: 'UNAUTHORIZED',
    message: 'Passwort ungültig oder fehlt.',
    technical: buildTechnical({ httpStatus: 401, raw: { message: 'UNAUTHORIZED' } })
  });
}

function authMiddleware(req, res, next) {
  if (APP_USER && APP_PASS) return basicAuthMiddleware(req, res, next);
  return toolPasswordMiddleware(req, res, next);
}

// ------------------------------------------------------------
// Rate Limit: TokenBucket 2 req/s + optional MIN_INTERVAL_MS + retry 429
// ------------------------------------------------------------
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
    if (this.timer) clearTimeout(this.timer);
    this.timer = null;

    this._refill();

    while (this.tokens >= 1 && this.queue.length) {
      this.tokens -= 1;
      const resolve = this.queue.shift();
      resolve();
    }

    if (this.queue.length) {
      this.timer = setTimeout(() => this._drain(), 100);
    }
  }
}

const bucket = new TokenBucket({ capacity: 2, refillPerSec: 2 });
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

    const base = 800 * Math.pow(2, attempt);
    const jitter = Math.floor(Math.random() * 250);
    const wait = Math.min(12000, base + jitter);
    await new Promise(r => setTimeout(r, wait));
  }

  return { status: 429, data: { message: 'Rate limit exceeded (client retries exhausted)' } };
}

// ------------------------------------------------------------
// Excel helper
// ------------------------------------------------------------
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

function round2(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

// ------------------------------------------------------------
// Artikel Cache + Artikel API
// ------------------------------------------------------------
const articleCache = {
  byId: new Map(),
  list: null,
  fetchedAt: 0,
  ttlMs: TEMPLATE_TTL_MS // 10 Minuten
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
  if (articleCache.list && (Date.now() - articleCache.fetchedAt) < articleCache.ttlMs) return articleCache.list;

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
    if (page > 100) break;
  }

  articleCache.list = all;
  articleCache.fetchedAt = Date.now();
  return all;
}

// ------------------------------------------------------------
// UnitPrice builder (AUTO)
// ------------------------------------------------------------
function buildUnitPriceFromNetGross({ net, gross, taxRate }) {
  const tr = Number(taxRate ?? 19);
  let netAmount = net != null ? Number(net) : null;
  let grossAmount = gross != null ? Number(gross) : null;

  if (netAmount == null && grossAmount != null) netAmount = grossAmount / (1 + tr / 100);
  if (grossAmount == null && netAmount != null) grossAmount = netAmount * (1 + tr / 100);

  if (netAmount == null || grossAmount == null) return null;

  return {
    currency: 'EUR',
    netAmount: round2(netAmount),
    grossAmount: round2(grossAmount),
    taxRatePercentage: tr
  };
}

function buildUnitPriceFromExcel({ taxType, amount, taxRate }) {
  const tr = Number(taxRate ?? 19);
  const a = Number(amount);
  if (!Number.isFinite(a)) return null;

  if (String(taxType).trim() === 'gross') {
    return buildUnitPriceFromNetGross({ gross: a, net: null, taxRate: tr });
  }
  return buildUnitPriceFromNetGross({ net: a, gross: null, taxRate: tr });
}

function buildUnitPriceFromArticle(articleObj) {
  const p = articleObj?.price;
  if (!p) return null;

  const taxRate = Number(p.taxRate ?? 19);
  const net = p.netPrice != null ? Number(p.netPrice) : null;
  const gross = p.grossPrice != null ? Number(p.grossPrice) : null;

  return buildUnitPriceFromNetGross({ net, gross, taxRate });
}

// ------------------------------------------------------------
// Excel -> Quotation Payload
// ------------------------------------------------------------
async function parseExcelAndBuildQuotationPayload(excelBase64, { allowPriceOverride }) {
  const errors = [];
  const warnings = [];
  const autoNamedLineItems = [];
  const byType = {};

  const wb = XLSX.read(Buffer.from(excelBase64, 'base64'), { type: 'buffer' });

  const angebotRows = sheetToJson(wb, 'Angebot');
  const kundeRows = sheetToJson(wb, 'Kunde');
  const posRows = sheetToJson(wb, 'Positionen');

  if (!angebotRows) errors.push({ sheet: 'Angebot', message: 'Sheet „Angebot“ fehlt.' });
  if (!kundeRows) errors.push({ sheet: 'Kunde', message: 'Sheet „Kunde“ fehlt.' });
  if (!posRows) errors.push({ sheet: 'Positionen', message: 'Sheet „Positionen“ fehlt.' });
  if (errors.length) return { ok: false, payload: null, summary: { errors, warnings } };

  const angebot = sheetRowsToKeyValueObject(angebotRows) || angebotRows[0] || {};
  const taxType = String(angebot.taxType || angebot.TaxType || angebot.TAXTYPE || '').trim();
  if (!taxType) errors.push({ sheet: 'Angebot', row: 2, field: 'taxType', message: 'taxType ist Pflicht (z. B. „gross“ oder „net“).' });

  const kunde = sheetRowsToKeyValueObject(kundeRows) || kundeRows[0] || {};
  const customerName = String(kunde.name || kunde.Name || '').trim();
  if (!customerName) errors.push({ sheet: 'Kunde', row: 2, field: 'name', message: 'Kundenname ist Pflicht.' });

  const taxConditions = taxType ? { taxType } : null;
  if (!taxConditions) errors.push({ sheet: 'Angebot', row: 2, field: 'taxConditions', message: 'taxConditions.taxType ist Pflicht.' });

  const voucherDate = new Date().toISOString();
  const expirationDate = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString();

  const address = {
    name: customerName || undefined,
    contactId: String(kunde.contactId || '').trim() || undefined,
    street: String(kunde.street || '').trim() || undefined,
    zip: String(kunde.zip || '').trim() || undefined,
    city: String(kunde.city || '').trim() || undefined,
    countryCode: String(kunde.countryCode || 'DE').trim() || 'DE',
    contactPerson: String(kunde.contactPerson || '').trim() || undefined,
    email: String(kunde.email || '').trim() || undefined,
    phone: String(kunde.phone || '').trim() || undefined
  };

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

    const unitPriceAmount = numOrNull(row.unitPriceAmount ?? row.unitPrice ?? row.price ?? row.Preis);
    const taxRatePercentage = numOrNull(row.taxRatePercentage ?? row.taxRate ?? row.tax);
    const discountPercent = numOrNull(row.discountPercent ?? row.discount);

    const hasAny = type || articleId || name || description || qty || unitName || unitPriceAmount || taxRatePercentage || discountPercent;
    if (!hasAny) continue;

    if (!type) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'type', message: 'type ist Pflicht.' });
      continue;
    }

    byType[type] = (byType[type] || 0) + 1;

    if (type === 'text') {
      const txtName = name || description || `Hinweis ${excelRow}`;
      lineItems.push({ type: 'text', name: txtName, description: description || undefined });
      continue;
    }

    if (!(qty > 0)) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'qty', message: 'qty muss größer als 0 sein.' });
      continue;
    }

    if (!unitName) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'unitName', message: 'unitName ist Pflicht.' });
      continue;
    }

    if ((type === 'material' || type === 'service') && !articleId) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'articleId', message: 'articleId ist Pflicht bei type=material/service.' });
      continue;
    }

    let articleObj = null;
    if (articleId) {
      articleObj = await getArticleById(articleId);
      if (!articleObj) {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'articleId',
          message: `Artikel konnte nicht geladen werden (articleId=${articleId}).`
        });
        continue;
      }
    }

    if (!name && articleId) {
      const autoName = articleObj?.title || `Artikel ${articleId}`;
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
      type,
      name,
      description: description || undefined,
      quantity: qty,
      unitName
    };

    if (articleId && (type === 'material' || type === 'service')) {
      item.id = articleId;
    }

    // --- AUTO unitPrice ---
    const canUseExcelPrice =
      unitPriceAmount !== null &&
      !(type === 'material' || type === 'service') &&
      !(type === 'custom' && articleId && allowPriceOverride !== true);

    if (canUseExcelPrice) {
      const rate = taxRatePercentage !== null ? taxRatePercentage : 19;
      const up = buildUnitPriceFromExcel({ taxType, amount: unitPriceAmount, taxRate: rate });
      if (!up) {
        errors.push({ sheet: 'Positionen', row: excelRow, field: 'unitPriceAmount', message: 'unitPrice konnte aus Excel nicht gebaut werden.' });
        continue;
      }
      item.unitPrice = up;
    } else {
      if (articleId) {
        const up = buildUnitPriceFromArticle(articleObj);
        if (!up) {
          errors.push({
            sheet: 'Positionen',
            row: excelRow,
            field: 'unitPriceAmount',
            message: `unitPrice fehlt/ist unvollständig im Artikelstamm (articleId=${articleId}).`
          });
          continue;
        }
        item.unitPrice = up;

        if (unitPriceAmount === null || type === 'material' || type === 'service') {
          warnings.push({
            sheet: 'Positionen',
            row: excelRow,
            message: `Preis automatisch aus Artikelstamm gesetzt (articleId=${articleId}).`
          });
        }
      } else {
        if (unitPriceAmount === null) {
          errors.push({
            sheet: 'Positionen',
            row: excelRow,
            field: 'unitPriceAmount',
            message: 'Preis ist Pflicht, wenn keine articleId gesetzt ist.'
          });
          continue;
        }
        const rate = taxRatePercentage !== null ? taxRatePercentage : 19;
        const up = buildUnitPriceFromExcel({ taxType, amount: unitPriceAmount, taxRate: rate });
        if (!up) {
          errors.push({ sheet: 'Positionen', row: excelRow, field: 'unitPriceAmount', message: 'unitPrice konnte aus Excel nicht gebaut werden.' });
          continue;
        }
        item.unitPrice = up;
      }
    }

    if (discountPercent !== null) item.discountPercentage = discountPercent;

    if (!item.unitPrice) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'unitPrice', message: 'unitPrice fehlt (würde Lexware 406 auslösen).' });
      continue;
    }

    lineItems.push(item);
  }

  if (!lineItems.length) errors.push({ sheet: 'Positionen', message: 'Keine Positionen gefunden.' });

  const summary = {
    errors,
    warnings,
    byType,
    autoNamedLineItems,
    allowPriceOverrideUsed: !!allowPriceOverride,
    voucherDate,
    taxType
  };

  if (errors.length) return { ok: false, payload: null, summary };

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

// ------------------------------------------------------------
// Idempotency (gegen 3x Erstellung)
// ------------------------------------------------------------
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

// ------------------------------------------------------------
// ✅ Dynamisches Template (10 Minuten TTL)
// ------------------------------------------------------------
const templateCache = {
  buffer: null,
  createdAt: 0
};

function buildTemplateWorkbook(articles) {
  const wb = XLSX.utils.book_new();

  const rows = (articles || []).map(a => ({
    id: a.id || '',
    title: a.title || '',
    articleNumber: a.articleNumber || '',
    type: a.type || '',
    unitName: a.unitName || '',
    netPrice: a.price?.netPrice ?? '',
    grossPrice: a.price?.grossPrice ?? '',
    leadingPrice: a.price?.leadingPrice ?? '',
    taxRate: a.price?.taxRate ?? '',
    archived: a.archived ?? '',
    version: a.version ?? ''
  }));
  const shArticles = XLSX.utils.json_to_sheet(rows, {
    header: ['id','title','articleNumber','type','unitName','netPrice','grossPrice','leadingPrice','taxRate','archived','version']
  });
  XLSX.utils.book_append_sheet(wb, shArticles, 'Artikel-Lookup');

  const angebot = [
    { Feld: 'taxType', Wert: 'net', Hinweis: 'Pflicht: net oder gross (entspricht Lexware taxConditions.taxType)' }
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(angebot), 'Angebot');

  const kunde = [
    { Feld: 'kind', Wert: 'company', Hinweis: 'company oder person' },
    { Feld: 'name', Wert: '', Hinweis: 'Pflicht: Firmenname (company) oder Vollname (person)' },
    { Feld: 'email', Wert: '', Hinweis: 'Optional, aber hilfreich fürs Matching' },
    { Feld: 'contactPerson', Wert: '', Hinweis: 'Optional' },
    { Feld: 'street', Wert: '', Hinweis: 'Optional' },
    { Feld: 'zip', Wert: '', Hinweis: 'Optional' },
    { Feld: 'city', Wert: '', Hinweis: 'Optional' },
    { Feld: 'countryCode', Wert: 'DE', Hinweis: 'ISO 3166-1 alpha-2, z.B. DE' },
    { Feld: 'phone', Wert: '', Hinweis: 'Optional' },
    { Feld: 'contactId', Wert: '', Hinweis: 'Optional: Lexware Contact ID, wenn du sie kennst' }
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(kunde), 'Kunde');

  const posHeader = [
    'pos','type','articleId','name','description','quantity','unitName','unitPriceAmount','taxRatePercentage','discountPercent'
  ];
  const posExample = [
    {
      pos: 1, type: 'custom', articleId: '', name: 'DTF Druck Brust 10×8 cm',
      description: '1-farbig, inkl. Positionierung', quantity: 25, unitName: 'Stk',
      unitPriceAmount: 6.9, taxRatePercentage: 19, discountPercent: ''
    },
    {
      pos: 2, type: 'text', articleId: '', name: 'Lieferzeit / Hinweis',
      description: 'Druckfreigabe erforderlich. Lieferzeit ca. 7–10 Werktage.', quantity: '', unitName: '',
      unitPriceAmount: '', taxRatePercentage: '', discountPercent: ''
    },
    {
      pos: 3, type: 'material', articleId: '(aus Artikel-Lookup id kopieren)', name: '',
      description: '', quantity: 25, unitName: 'Stk', unitPriceAmount: '',
      taxRatePercentage: '', discountPercent: ''
    }
  ];
  const shPos = XLSX.utils.json_to_sheet(posExample, { header: posHeader });
  XLSX.utils.book_append_sheet(wb, shPos, 'Positionen');

  const help = [
    { Schritt: 1, Hinweis: 'Artikel-Lookup: gewünschte articleId kopieren' },
    { Schritt: 2, Hinweis: 'Angebot: taxType setzen (net/gross)' },
    { Schritt: 3, Hinweis: 'Kunde: name ist Pflicht' },
    { Schritt: 4, Hinweis: 'Positionen: type + quantity > 0 + unitName Pflicht (außer text)' },
    { Schritt: 5, Hinweis: 'material/service: articleId Pflicht, Preis wird automatisch aus Artikelstamm gesetzt' },
    { Schritt: 6, Hinweis: 'custom ohne articleId: unitPriceAmount Pflicht' }
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(help), 'Anleitung');

  return wb;
}

// ✅ NEU (Phase 1): Text → Excel Workbook im gleichen Format
function buildWorkbookFromParsedText({ items, taxType }) {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { Feld: 'taxType', Wert: taxType || 'net', Hinweis: 'Pflicht: net oder gross' }
  ]), 'Angebot');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { Feld: 'kind', Wert: 'company', Hinweis: 'company oder person' },
    { Feld: 'name', Wert: '', Hinweis: 'Pflicht: Firmenname (company) oder Vollname (person)' },
    { Feld: 'email', Wert: '', Hinweis: 'Optional' },
    { Feld: 'contactPerson', Wert: '', Hinweis: 'Optional' },
    { Feld: 'street', Wert: '', Hinweis: 'Optional' },
    { Feld: 'zip', Wert: '', Hinweis: 'Optional' },
    { Feld: 'city', Wert: '', Hinweis: 'Optional' },
    { Feld: 'countryCode', Wert: 'DE', Hinweis: 'ISO 3166-1 alpha-2, z.B. DE' },
    { Feld: 'phone', Wert: '', Hinweis: 'Optional' }
  ]), 'Kunde');

  const header = ['pos','type','articleId','name','description','quantity','unitName','unitPriceAmount','taxRatePercentage','discountPercent'];
  const rows = (items || []).map((it, idx) => ({
    pos: idx + 1,
    type: it.type || 'custom',
    articleId: it.articleId || '',
    name: it.name || '',
    description: it.description || (it.articleNumber ? `ArtNr/SKU: ${it.articleNumber}` : ''),
    quantity: it.type === 'text' ? '' : (it.quantity ?? ''),
    unitName: it.type === 'text' ? '' : (it.unitName || 'Stk'),
    unitPriceAmount: '',
    taxRatePercentage: '',
    discountPercent: ''
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows, { header }), 'Positionen');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { Schritt: 1, Hinweis: 'Diese Datei wurde aus "Text einfügen" erzeugt.' },
    { Schritt: 2, Hinweis: 'Bitte Kunde.name + Angebot.taxType prüfen/ausfüllen.' },
    { Schritt: 3, Hinweis: 'custom ohne articleId: unitPriceAmount ist Pflicht.' },
    { Schritt: 4, Hinweis: 'material/service: articleId setzen (aus Artikel-Lookup) — Preis kommt dann automatisch.' }
  ]), 'Anleitung');

  return wb;
}

// ------------------------------------------------------------
// API
// ------------------------------------------------------------
app.get('/api/ping', (req, res) => {
  ok(res, {
    status: 'OK',
    passwordProtected: !!(TOOL_PASSWORD || (APP_USER && APP_PASS)),
    passwordMode: (APP_USER && APP_PASS) ? 'basic' : (TOOL_PASSWORD ? 'toolPassword' : 'none'),
    allowPriceOverrideDefault: ALLOW_PRICE_OVERRIDE_DEFAULT,
    minIntervalMs: MIN_INTERVAL_MS,
    apiBaseUrl: API_BASE_URL,
    finalizeDefault: FINALIZE_DEFAULT,
    templateTtlMs: TEMPLATE_TTL_MS,
    parserAvailable: !!parseLines
  });
});

app.get('/api/articles', authMiddleware, async (req, res) => {
  try {
    if (!API_KEY) {
      return fail(res, {
        stage: 'config',
        status: 'CONFIG_ERROR',
        message: 'API Key fehlt.',
        technical: buildTechnical({ httpStatus: 500, raw: { message: 'NO_API_KEY' } })
      });
    }
    const list = await listAllArticlesCached();
    ok(res, { status: 'SUCCESS', data: { count: list.length, articles: list } });
  } catch (err) {
    fail(res, {
      stage: 'articles',
      status: 'ERROR',
      message: err.message,
      technical: buildTechnical({ httpStatus: 500, raw: { message: 'ARTICLES_EXCEPTION' }, err })
    });
  }
});

// ✅ Dynamisches Template (auth + TTL 10min)
app.get('/api/template.xlsx', authMiddleware, async (req, res) => {
  try {
    if (!API_KEY) {
      return res.status(200).json({
        ok: false,
        stage: 'template',
        status: 'CONFIG_ERROR',
        message: 'API Key fehlt.',
        technical: buildTechnical({ httpStatus: 500, raw: { message: 'NO_API_KEY' } })
      });
    }

    const now = Date.now();
    if (templateCache.buffer && (now - templateCache.createdAt) < TEMPLATE_TTL_MS) {
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="lexware_template.xlsx"');
      return res.status(200).send(templateCache.buffer);
    }

    const articles = await listAllArticlesCached();
    const wb = buildTemplateWorkbook(articles);
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    templateCache.buffer = buffer;
    templateCache.createdAt = now;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="lexware_template.xlsx"');
    return res.status(200).send(buffer);
  } catch (err) {
    return res.status(200).json({
      ok: false,
      stage: 'template',
      status: 'ERROR',
      message: err.message,
      technical: buildTechnical({ httpStatus: 500, raw: { message: 'TEMPLATE_EXCEPTION' }, err })
    });
  }
});

// ✅ Phase 1: Text Parser → JSON (falls Parser fehlt: CONFIG_ERROR statt Crash)
app.post('/api/parse-text', authMiddleware, async (req, res) => {
  try {
    if (!parseLines) {
      return fail(res, {
        stage: 'parse-text',
        status: 'CONFIG_ERROR',
        message: 'Parser-Modul fehlt (./lib/parseText.js). Bitte Datei hinzufügen/committen.',
        technical: buildTechnical({ httpStatus: 500, raw: { message: 'PARSER_NOT_AVAILABLE' } })
      });
    }

    const text = String(req.body?.text || '').trim();
    if (!text) {
      return fail(res, {
        stage: 'parse-text',
        status: 'VALIDATION_ERROR',
        message: 'Kein Text übergeben.',
        technical: buildTechnical({ httpStatus: 400, raw: { message: 'NO_TEXT' } })
      });
    }

    const { items, warnings } = parseLines(text);
    return ok(res, {
      stage: 'parse-text',
      status: 'SUCCESS',
      data: { count: items.length, warnings, items }
    });
  } catch (err) {
    return fail(res, {
      stage: 'parse-text',
      status: 'ERROR',
      message: err.message,
      technical: buildTechnical({ httpStatus: 500, raw: { message: 'PARSE_EXCEPTION' }, err })
    });
  }
});

// ✅ Phase 1: Text → Excel Download
app.get('/api/text-to-excel', authMiddleware, async (req, res) => {
  try {
    if (!parseLines) {
      return res.status(200).json({
        ok: false,
        stage: 'text-to-excel',
        status: 'CONFIG_ERROR',
        message: 'Parser-Modul fehlt (./lib/parseText.js). Bitte Datei hinzufügen/committen.',
        technical: buildTechnical({ httpStatus: 500, raw: { message: 'PARSER_NOT_AVAILABLE' } })
      });
    }

    const text = String(req.query?.text || '').trim();
    const taxType = String(req.query?.taxType || 'net').trim();
    if (!text) return res.status(400).send('Missing text');

    const { items } = parseLines(text);
    const wb = buildWorkbookFromParsedText({ items, taxType });
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="from_text.xlsx"');
    return res.status(200).send(buffer);
  } catch (err) {
    return res.status(200).json({
      ok: false,
      stage: 'text-to-excel',
      status: 'ERROR',
      message: err.message,
      technical: buildTechnical({ httpStatus: 500, raw: { message: 'TEXT_TO_EXCEL_EXCEPTION' }, err })
    });
  }
});

app.post('/api/test-excel', authMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body || {};
    const allow = typeof allowPriceOverride === 'boolean' ? allowPriceOverride : ALLOW_PRICE_OVERRIDE_DEFAULT;

    if (!excelData) {
      return fail(res, {
        stage: 'input',
        status: 'VALIDATION_ERROR',
        message: 'Keine Excel-Daten übergeben.',
        technical: buildTechnical({ httpStatus: 400, raw: { message: 'NO_EXCEL' } })
      });
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
    fail(res, {
      stage: 'test',
      status: 'ERROR',
      message: err.message,
      technical: buildTechnical({ httpStatus: 500, raw: { message: 'TEST_EXCEPTION' }, err })
    });
  }
});

app.post('/api/create-offer', authMiddleware, async (req, res) => {
  const startedAt = Date.now();

  try {
    const { excelData, allowPriceOverride, finalize } = req.body || {};
    const allow = typeof allowPriceOverride === 'boolean' ? allowPriceOverride : ALLOW_PRICE_OVERRIDE_DEFAULT;
    const doFinalize = typeof finalize === 'boolean' ? finalize : FINALIZE_DEFAULT;

    if (!excelData) {
      return fail(res, {
        stage: 'input',
        status: 'VALIDATION_ERROR',
        message: 'Keine Excel-Daten übergeben.',
        technical: buildTechnical({ httpStatus: 400, raw: { message: 'NO_EXCEL' } })
      });
    }

    if (!API_KEY) {
      return fail(res, {
        stage: 'config',
        status: 'CONFIG_ERROR',
        message: 'API Key fehlt.',
        technical: buildTechnical({ httpStatus: 500, raw: { message: 'NO_API_KEY' } })
      });
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
    setTimeout(() => inFlight.delete(key), 15000);

    return res.json(result);
  } catch (err) {
    return fail(res, {
      stage: 'lexware-create',
      status: 'ERROR',
      message: err.message || 'Unerwarteter Fehler',
      technical: buildTechnical({ httpStatus: 500, raw: { message: 'UNHANDLED_EXCEPTION' }, err })
    });
  }
});

app.get('/api/download-pdf', authMiddleware, async (req, res) => {
  try {
    const quotationId = String(req.query.id || '').trim();
    if (!quotationId) return res.status(400).send('Missing id');
    if (!API_KEY) return res.status(500).send('API Key fehlt');

    const apiRes = await lexwareRequest({
      method: 'GET',
      url: `${API_BASE_URL}/v1/quotations/${encodeURIComponent(quotationId)}/file`,
      responseType: 'arraybuffer',
      accept: '*/*'
    });

    if (apiRes.status < 200 || apiRes.status >= 300) {
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

// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server läuft auf Port', PORT));
