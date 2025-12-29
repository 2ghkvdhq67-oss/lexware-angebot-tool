'use strict';

/**
 * Maiershirts — Lexoffice / Lexware Angebots-Tool Backend
 * ------------------------------------------------------
 * Features:
 *  - Passwortschutz (optional via TOOL_PASSWORD)
 *  - Schalter: allowPriceOverride
 *  - Testmodus / Validierung
 *  - Angebot erstellen & finalisieren
 *  - Zentraler Rate-Limiter (mind. 600ms Abstand)
 *  - Saubere 429-Behandlung (RATE_LIMIT → kein Retry)
 *  - Excel-Upload (Base64 oder Buffer)
 *  - Liefert deinen vorhandenen public/ Ordner aus (index.html)
 */

const path = require('path');          // ✅ NEU: für public-Ordner
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const dotenv = require('dotenv');
const XLSX = require('xlsx');

dotenv.config();

const app = express();
app.use(bodyParser.json({ limit: '10mb' }));

// ✅ NEU: Statische Dateien aus public/ ausliefern (index.html, JS, CSS, …)
app.use(express.static(path.join(__dirname, 'public')));

// ------------------------------------------------------
// ENV Konfiguration
// ------------------------------------------------------

const LEXOFFICE_API_KEY = process.env.LEXOFFICE_API_KEY || '';
const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';
const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE || 'false').toLowerCase() === 'true';

const MIN_INTERVAL_MS = parseInt(process.env.LEXWARE_MIN_INTERVAL_MS || '600', 10);

if (!LEXOFFICE_API_KEY) {
  console.warn('[WARNUNG] LEXOFFICE_API_KEY ist nicht gesetzt.');
}

console.log('[INFO] Rate-Limit Mindestabstand:', MIN_INTERVAL_MS, 'ms');

// ------------------------------------------------------
// Zentraler Rate-Limiter
// ------------------------------------------------------

let lastLexwareCallTime = 0;

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function callLexwareApi(axiosConfig) {

  const now = Date.now();
  const elapsed = now - lastLexwareCallTime;

  if (elapsed < MIN_INTERVAL_MS) {
    await sleep(MIN_INTERVAL_MS - elapsed);
  }

  const res = await axios(axiosConfig);

  lastLexwareCallTime = Date.now();
  return res;
}

// ------------------------------------------------------
// Passwortschutz (optional)
// ------------------------------------------------------

function passwordMiddleware(req, res, next) {

  if (!TOOL_PASSWORD)
    return next();

  const provided =
    req.headers['x-tool-password'] ||
    req.body?.password ||
    req.query?.password;

  if (provided !== TOOL_PASSWORD) {
    return res.status(401).json({
      ok: false,
      status: 'UNAUTHORIZED',
      stage: 'auth',
      message: 'Zugriff verweigert. Passwort ungültig oder fehlt.'
    });
  }

  next();
}

// ------------------------------------------------------
// Excel → Payload Mapping & Validierung
// ------------------------------------------------------

async function parseExcelAndBuildQuotationPayload(excelData, options) {

  const { allowPriceOverride, mode } = options || {};

  // Excel einlesen
  let workbook;

  if (Buffer.isBuffer(excelData)) {
    workbook = XLSX.read(excelData, { type: 'buffer' });
  } else if (typeof excelData === 'string') {
    const buf = Buffer.from(excelData, 'base64');
    workbook = XLSX.read(buf, { type: 'buffer' });
  } else {
    const err = new Error('excelData Format unbekannt (erwartet Buffer oder Base64-String)');
    err.status = 400;
    throw err;
  }

  function sheetJson(name, required = true) {
    const sheet = workbook.Sheets[name];
    if (!sheet) return { rows: [], missing: required };
    return { rows: XLSX.utils.sheet_to_json(sheet, { defval: '' }), missing: false };
  }

  const angebot = sheetJson('Angebot', true);
  const kunde = sheetJson('Kunde', true);
  const positionen = sheetJson('Positionen', true);

  const errors = [];
  const warnings = [];
  const autoNamedLineItems = [];
  const byType = {};

  // Pflicht-Sheets
  if (angebot.missing)
    errors.push({ sheet: 'Angebot', message: 'Tab „Angebot“ fehlt.' });

  if (kunde.missing)
    errors.push({ sheet: 'Kunde', message: 'Tab „Kunde“ fehlt.' });

  if (positionen.missing)
    errors.push({ sheet: 'Positionen', message: 'Tab „Positionen“ fehlt.' });

  if (errors.length)
    return { quotationPayload: null, summary: { byType, warnings, errors, autoNamedLineItems } };

  // Angebot
  const angebotRow = angebot.rows[0] || {};
  const taxType =
    angebotRow.taxType ||
    angebotRow.TAXTYPE ||
    angebotRow.tax ||
    '';

  if (!taxType) {
    errors.push({
      sheet: 'Angebot',
      row: 2,
      field: 'taxType',
      message: 'taxType ist Pflicht.'
    });
  }

  // Kunde
  const kundeRow = kunde.rows[0] || {};
  const customerName = kundeRow.name || kundeRow.Name || '';

  if (!customerName) {
    errors.push({
      sheet: 'Kunde',
      row: 2,
      field: 'name',
      message: 'Kundenname ist Pflicht.'
    });
  }

  const address = {
    name: customerName,
    street: kundeRow.street || kundeRow.Straße || '',
    zip: kundeRow.zip || kundeRow.PLZ || '',
    city: kundeRow.city || kundeRow.Ort || '',
    countryCode: kundeRow.countryCode || 'DE'
  };

  // Positionen
  const lineItems = [];

  positionen.rows.forEach((row, idx) => {

    const excelRow = idx + 2;

    const type = (row.type || row.Typ || '').toString().trim();
    const qty = Number(row.qty || row.Menge || 0);
    const articleId = (row.articleId || row.articleID || '').toString().trim();
    const priceExcel = row.price || row.Preis || null;
    const name = (row.name || row.Bezeichnung || '').toString().trim();

    if (!type) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'type', message: 'type ist Pflicht.' });
      return;
    }

    if (!byType[type]) byType[type] = 0;
    byType[type]++;

    if (!(qty > 0)) {
      errors.push({ sheet: 'Positionen', row: excelRow, field: 'qty', message: 'qty muss > 0 sein.' });
      return;
    }

    let useExcelPrice = false;

    if (!articleId) {

      if (priceExcel == null || priceExcel === '') {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'price',
          message: 'Preis ist Pflicht, wenn keine articleId gesetzt ist.'
        });
        return;
      }

      useExcelPrice = true;

    } else {

      if (type === 'material')
        useExcelPrice = false;

      else if (allowPriceOverride)
        useExcelPrice = true;

      else
        useExcelPrice = false;
    }

    let finalName = name;

    if (!finalName && articleId) {
      finalName = `(Name aus Artikel ${articleId})`;
      autoNamedLineItems.push({ row: excelRow, articleId, name: finalName });
      warnings.push({ row: excelRow, message: `Name automatisch aus Artikel ${articleId}` });
    }

    lineItems.push({
      type: 'custom',
      articleId: articleId || null,
      name: finalName || `Position ${excelRow}`,
      quantity: qty,
      price: useExcelPrice ? Number(priceExcel) : null,
      meta: {
        excelRow,
        sourcePrice: useExcelPrice ? 'excel' : 'article'
      }
    });
  });

  // Wenn Fehler → kein Payload erzeugen
  if (errors.length) {
    return {
      quotationPayload: null,
      summary: { byType, warnings, errors, autoNamedLineItems }
    };
  }

  // Angebots-Payload (Schema ggf. erweitern)
  const quotationPayload = {
    voucherDate: new Date().toISOString().substring(0, 10),
    address,
    taxType,
    lineItems,
    finalized: true
  };

  return {
    quotationPayload,
    summary: { byType, warnings, errors, autoNamedLineItems }
  };
}

// ------------------------------------------------------
// Lexoffice Angebot erstellen
// ------------------------------------------------------

async function createLexofficeQuotationInApi(payload) {

  if (!LEXOFFICE_API_KEY) {
    const err = new Error('API Key fehlt');
    err.status = 500;
    throw err;
  }

  const res = await callLexwareApi({
    method: 'POST',
    url: 'https://api.lexoffice.io/v1/quotations',
    data: payload,
    headers: {
      Authorization: `Bearer ${LEXOFFICE_API_KEY}`,
      'Content-Type': 'application/json',
      Accept: 'application/json'
    },
    validateStatus: () => true
  });

  if (res.status === 429) {
    const err = new Error('Rate limit exceeded');
    err.status = 429;
    err.raw = res.data;
    throw err;
  }

  if (res.status < 200 || res.status >= 300) {
    const err = new Error('Lexoffice Error');
    err.status = res.status;
    err.raw = res.data;
    throw err;
  }

  return res.data;
}

// ------------------------------------------------------
// API ROUTES
// ------------------------------------------------------

app.get('/api/ping', (req, res) => {
  res.json({
    ok: true,
    status: 'OK',
    passwordProtected: !!TOOL_PASSWORD,
    allowPriceOverrideDefault: ALLOW_PRICE_OVERRIDE_DEFAULT,
    minIntervalMs: MIN_INTERVAL_MS
  });
});

// ---------------- Testmodus ----------------

app.post('/api/test-excel', passwordMiddleware, async (req, res) => {

  try {

    const { excelData, allowPriceOverride } = req.body || {};

    if (!excelData)
      return res.json({
        ok: false,
        status: 'VALIDATION_ERROR',
        message: 'Excel fehlt',
        httpStatus: 400
      });

    const result = await parseExcelAndBuildQuotationPayload(excelData, {
      allowPriceOverride: typeof allowPriceOverride === 'boolean'
        ? allowPriceOverride
        : ALLOW_PRICE_OVERRIDE_DEFAULT,
      mode: 'test'
    });

    const hasErrors = result.summary.errors.length > 0;

    return res.json({
      ok: !hasErrors,
      status: hasErrors ? 'VALIDATION_ERROR' : 'SUCCESS',
      stage: 'test',
      data: result.summary
    });

  } catch (err) {

    console.error('[test-excel]', err);

    res.json({
      ok: false,
      status: 'ERROR',
      stage: 'test',
      message: err.message,
      httpStatus: err.status || 500
    });
  }
});

// --------------- Angebot erstellen ---------------

app.post('/api/create-offer', passwordMiddleware, async (req, res) => {

  try {

    const { excelData, allowPriceOverride } = req.body || {};

    if (!excelData)
      return res.json({
        ok: false,
        status: 'VALIDATION_ERROR',
        message: 'Excel fehlt',
        httpStatus: 400
      });

    const { quotationPayload, summary } =
      await parseExcelAndBuildQuotationPayload(excelData, {
        allowPriceOverride: typeof allowPriceOverride === 'boolean'
          ? allowPriceOverride
          : ALLOW_PRICE_OVERRIDE_DEFAULT,
        mode: 'create'
      });

    if (!quotationPayload)
      return res.json({
        ok: false,
        status: 'VALIDATION_ERROR',
        stage: 'validation',
        message: 'Excel enthält Fehler',
        data: summary
      });

    const result = await createLexofficeQuotationInApi(quotationPayload);

    return res.json({
      ok: true,
      status: 'SUCCESS',
      stage: 'lexoffice-create',
      data: {
        quotationId: result.id || null,
        summary
      }
    });

  } catch (err) {

    console.error('[create-offer]', err);

    if (err.status === 429)
      return res.json({
        ok: false,
        status: 'RATE_LIMIT',
        stage: 'lexoffice-create',
        message: 'Rate-Limit — später erneut versuchen',
        httpStatus: 429
      });

    res.json({
      ok: false,
      status: 'ERROR',
      stage: 'lexoffice-create',
      message: err.message,
      httpStatus: err.status || 500
    });
  }
});

// ------------------------------------------------------

const PORT = process.env.PORT || 3000;

app.listen(PORT, () =>
  console.log('Backend läuft auf Port', PORT)
);
