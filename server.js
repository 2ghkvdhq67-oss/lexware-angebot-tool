'use strict';

/**
 * Maiershirts — Lexoffice Angebots-Tool Backend
 * --------------------------------------------
 * Funktionen:
 *  ✔ Passwortschutz (optional via TOOL_PASSWORD)
 *  ✔ Testmodus / Validierung
 *  ✔ Angebot erstellen
 *  ✔ Preis-Override Schalter (optional)
 *  ✔ Automatisches Ergänzen von Artikelnamen
 *  ✔ Rate-Limiter (mind. 600ms Abstand zu API Calls)
 *  ✔ Strukturierte Fehlermeldungen
 *  ✔ Statisches Frontend aus /public
 */

const path = require('path');
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const dotenv = require('dotenv');
const XLSX = require('xlsx');

dotenv.config();

const app = express();
app.use(bodyParser.json({ limit: '10mb' }));

// 👉 Frontend aus /public Ordner bereitstellen
app.use(express.static(path.join(__dirname, 'public')));

// ----------------------------------------------------
// ENV Variablen
// ----------------------------------------------------

const LEXOFFICE_API_KEY = process.env.LEXOFFICE_API_KEY || '';
const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';

const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE || 'false').toLowerCase() === 'true';

const MIN_INTERVAL_MS = parseInt(process.env.LEXWARE_MIN_INTERVAL_MS || '600', 10);

console.log('[INFO] Rate-Limiter aktiv:', MIN_INTERVAL_MS, 'ms');

// ----------------------------------------------------
// Rate Limiter
// ----------------------------------------------------

let lastCall = 0;

async function wait(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function callLexoffice(config) {
  const diff = Date.now() - lastCall;
  if (diff < MIN_INTERVAL_MS) {
    await wait(MIN_INTERVAL_MS - diff);
  }

  const res = await axios(config);
  lastCall = Date.now();
  return res;
}

// ----------------------------------------------------
// Passwortschutz (optional)
// ----------------------------------------------------

function passwordMiddleware(req, res, next) {
  if (!TOOL_PASSWORD) return next();

  const supplied =
    req.body?.password ||
    req.headers['x-tool-password'] ||
    req.query?.password;

  if (supplied !== TOOL_PASSWORD) {
    return res.status(401).json({
      ok: false,
      stage: 'auth',
      status: 'UNAUTHORIZED',
      message: 'Passwort ungültig oder fehlt'
    });
  }

  next();
}

// ----------------------------------------------------
// Excel Parser & Validierung
// ----------------------------------------------------

async function parseExcel(excelData, options = {}) {
  const { allowPriceOverride } = options;

  let workbook;

  if (typeof excelData === 'string') {
    const buf = Buffer.from(excelData, 'base64');
    workbook = XLSX.read(buf, { type: 'buffer' });
  } else if (Buffer.isBuffer(excelData)) {
    workbook = XLSX.read(excelData, { type: 'buffer' });
  } else {
    throw new Error('excelData Format ungültig');
  }

  function readSheet(name) {
    const sh = workbook.Sheets[name];
    return sh ? XLSX.utils.sheet_to_json(sh, { defval: '' }) : null;
  }

  const angebot = readSheet('Angebot');
  const kunde = readSheet('Kunde');
  const positionen = readSheet('Positionen');

  const errors = [];
  const warnings = [];
  const autoNamed = [];
  const byType = {};

  if (!angebot) errors.push({ sheet: 'Angebot', message: 'Sheet fehlt' });
  if (!kunde) errors.push({ sheet: 'Kunde', message: 'Sheet fehlt' });
  if (!positionen) errors.push({ sheet: 'Positionen', message: 'Sheet fehlt' });

  if (errors.length) {
    return { quotation: null, summary: { errors, warnings, autoNamed, byType } };
  }

  const angebotRow = angebot[0] || {};
  const taxType = angebotRow.taxType || angebotRow.TAXTYPE || '';

  if (!taxType) {
    errors.push({ sheet: 'Angebot', field: 'taxType', message: 'taxType ist Pflicht' });
  }

  const kundeRow = kunde[0] || {};
  const customerName = kundeRow.name || kundeRow.Name || '';

  if (!customerName) {
    errors.push({ sheet: 'Kunde', field: 'name', message: 'Kundenname ist Pflicht' });
  }

  const address = {
    name: customerName,
    street: kundeRow.street || '',
    zip: kundeRow.zip || '',
    city: kundeRow.city || '',
    countryCode: kundeRow.countryCode || 'DE'
  };

  const lineItems = [];

  positionen.forEach((row, i) => {
    const excelRow = i + 2;

    const type = (row.type || '').toString().trim();
    const qty = Number(row.qty || 0);
    const articleId = (row.articleId || '').toString().trim();
    const priceExcel = row.price || null;
    const name = (row.name || '').toString().trim();

    if (!type) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'type', message: 'type ist Pflicht' });
      return;
    }

    if (!(qty > 0)) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'qty', message: 'qty > 0 erforderlich' });
      return;
    }

    if (!byType[type]) byType[type] = 0;
    byType[type]++;

    let usePrice = false;

    if (!articleId) {
      if (!priceExcel) {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'price',
          message: 'Preis erforderlich wenn keine articleId'
        });
        return;
      }
      usePrice = true;
    } else {
      if (type === 'material') usePrice = false;
      else if (allowPriceOverride) usePrice = true;
      else usePrice = false;
    }

    let finalName = name;

    if (!finalName && articleId) {
      finalName = `(Artikel ${articleId})`;
      autoNamed.push({ row: excelRow, articleId, name: finalName });
      warnings.push({ row: excelRow, message: 'Name automatisch ergänzt' });
    }

    lineItems.push({
      type: 'custom',
      articleId: articleId || null,
      name: finalName || `Position ${excelRow}`,
      quantity: qty,
      price: usePrice ? Number(priceExcel) : null
    });
  });

  if (errors.length) {
    return {
      quotation: null,
      summary: { errors, warnings, autoNamed, byType }
    };
  }

  const quotation = {
    voucherDate: new Date().toISOString().substring(0, 10),
    taxType,
    address,
    lineItems,
    finalized: true
  };

  return { quotation, summary: { errors, warnings, autoNamed, byType } };
}

// ----------------------------------------------------
// Lexoffice — Angebot erstellen
// ----------------------------------------------------

async function createQuotation(payload) {
  const res = await callLexoffice({
    method: 'POST',
    url: 'https://api.lexoffice.io/v1/quotations',
    headers: {
      Authorization: `Bearer ${LEXOFFICE_API_KEY}`,
      'Content-Type': 'application/json'
    },
    data: payload,
    validateStatus: () => true
  });

  if (res.status === 429) {
    const e = new Error('Rate Limit');
    e.status = 429;
    throw e;
  }

  if (res.status < 200 || res.status >= 300) {
    const e = new Error('Lexoffice Fehler');
    e.status = res.status;
    e.raw = res.data;
    throw e;
  }

  return res.data;
}

// ----------------------------------------------------
// API Routes
// ----------------------------------------------------

app.get('/api/ping', (req, res) => {
  res.json({
    ok: true,
    passwordProtected: !!TOOL_PASSWORD,
    allowPriceOverrideDefault: ALLOW_PRICE_OVERRIDE_DEFAULT,
    minIntervalMs: MIN_INTERVAL_MS
  });
});

// Testmodus
app.post('/api/test-excel', passwordMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body;

    const parsed = await parseExcel(excelData, {
      allowPriceOverride:
        typeof allowPriceOverride === 'boolean'
          ? allowPriceOverride
          : ALLOW_PRICE_OVERRIDE_DEFAULT
    });

    const hasErrors = parsed.summary.errors.length > 0;

    res.json({
      ok: !hasErrors,
      stage: 'test',
      status: hasErrors ? 'VALIDATION_ERROR' : 'SUCCESS',
      data: parsed.summary
    });

  } catch (err) {
    res.json({
      ok: false,
      stage: 'test',
      status: 'ERROR',
      message: err.message
    });
  }
});

// Angebot erstellen
app.post('/api/create-offer', passwordMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body;

    const parsed = await parseExcel(excelData, {
      allowPriceOverride:
        typeof allowPriceOverride === 'boolean'
          ? allowPriceOverride
          : ALLOW_PRICE_OVERRIDE_DEFAULT
    });

    if (!parsed.quotation) {
      return res.json({
        ok: false,
        stage: 'validation',
        status: 'VALIDATION_ERROR',
        data: parsed.summary
      });
    }

    const result = await createQuotation(parsed.quotation);

    res.json({
      ok: true,
      stage: 'lexoffice-create',
      status: 'SUCCESS',
      data: {
        quotationId: result.id,
        summary: parsed.summary
      }
    });

  } catch (err) {

    if (err.status === 429) {
      return res.json({
        ok: false,
        status: 'RATE_LIMIT',
        stage: 'lexoffice-create',
        message: 'Rate Limit — später erneut versuchen'
      });
    }

    res.json({
      ok: false,
      status: 'ERROR',
      stage: 'lexoffice-create',
      message: err.message
    });
  }
});

// ----------------------------------------------------

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server läuft auf Port', PORT));
