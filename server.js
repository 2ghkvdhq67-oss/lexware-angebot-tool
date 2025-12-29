'use strict';

const path = require('path');
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const dotenv = require('dotenv');
const XLSX = require('xlsx');

dotenv.config();

const app = express();
app.use(bodyParser.json({ limit: '10mb' }));

// ----------------------------------------------------
// Static assets
// ----------------------------------------------------

// Frontend aus /public (index.html, CSS, JS …)
app.use(express.static(path.join(__dirname, 'public')));

// Templates-Ordner explizit freigeben (Variante B)
app.use('/templates', express.static(path.join(__dirname, 'templates')));

// ----------------------------------------------------
// ENV Variablen
// ----------------------------------------------------

const LEXOFFICE_API_KEY = process.env.LEXOFFICE_API_KEY || '';
const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';
const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE || 'false').toLowerCase() === 'true';
const MIN_INTERVAL_MS = parseInt(process.env.LEXWARE_MIN_INTERVAL_MS || '600', 10);

console.log('[INFO] Rate-Limiter:', MIN_INTERVAL_MS, 'ms');

// ----------------------------------------------------
// Rate Limiter
// ----------------------------------------------------

let lastCallTs = 0;
const sleep = ms => new Promise(r => setTimeout(r, ms));

async function callLexoffice(config) {
  const now = Date.now();
  const diff = now - lastCallTs;
  if (diff < MIN_INTERVAL_MS) {
    await sleep(MIN_INTERVAL_MS - diff);
  }

  const res = await axios(config);
  lastCallTs = Date.now();
  return res;
}

// ----------------------------------------------------
// Passwort-Middleware
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
      message: 'Passwort ungültig oder fehlt.'
    });
  }

  next();
}

// ----------------------------------------------------
// Helper: vertikales Sheet (Feld / Wert / Hinweis)
// ----------------------------------------------------

function sheetRowsToKeyValueObject(rows) {
  if (!rows || !rows.length) return null;

  const first = rows[0];
  const fieldKeys = ['Feld', 'feld', 'Field', 'field'];
  const valueKeys = ['Wert', 'wert', 'Value', 'value', 'val'];

  const fieldCol = fieldKeys.find(k => k in first);
  if (!fieldCol) return null;

  const valueCol = valueKeys.find(k => k in first);
  if (!valueCol) return null;

  const obj = {};
  rows.forEach(r => {
    const key = (r[fieldCol] || '').toString().trim();
    if (!key) return;
    obj[key] = r[valueCol];
  });

  return obj;
}

// ----------------------------------------------------
// Excel-Parser & Validierung
// ----------------------------------------------------

async function parseExcel(excelData, { allowPriceOverride }) {
  if (!excelData) {
    const e = new Error('Keine Excel-Daten übergeben.');
    e.status = 400;
    throw e;
  }

  let workbook;

  if (typeof excelData === 'string') {
    workbook = XLSX.read(Buffer.from(excelData, 'base64'), { type: 'buffer' });
  } else if (Buffer.isBuffer(excelData)) {
    workbook = XLSX.read(excelData, { type: 'buffer' });
  } else {
    const e = new Error('excelData Format ungültig');
    e.status = 400;
    throw e;
  }

  const readSheet = name => {
    const sh = workbook.Sheets[name];
    return sh ? XLSX.utils.sheet_to_json(sh, { defval: '' }) : null;
  };

  const angebotRows = readSheet('Angebot');
  const kundeRows = readSheet('Kunde');
  const positions = readSheet('Positionen');

  const errors = [];
  const warnings = [];
  const autoNamed = [];
  const byType = {};

  // Pflicht-Sheets
  if (!angebotRows) errors.push({ sheet: 'Angebot', message: 'Sheet „Angebot“ fehlt.' });
  if (!kundeRows) errors.push({ sheet: 'Kunde', message: 'Sheet „Kunde“ fehlt.' });
  if (!positions) errors.push({ sheet: 'Positionen', message: 'Sheet „Positionen“ fehlt.' });

  if (errors.length) {
    return { quotation: null, summary: { errors, warnings, autoNamed, byType } };
  }

  // -------- Angebot: horizontal ODER vertikal --------
  let angebot = sheetRowsToKeyValueObject(angebotRows) || angebotRows[0] || {};
  const taxType =
    angebot.taxType ||
    angebot.TAXTYPE ||
    angebot.tax ||
    angebot.TaxType ||
    '';

  if (!taxType) {
    errors.push({
      sheet: 'Angebot',
      row: 2,
      field: 'taxType',
      message: 'taxType ist Pflicht (z. B. „net“ oder „gross“).'
    });
  }

  // -------- Kunde: horizontal ODER vertikal --------
  let kunde = sheetRowsToKeyValueObject(kundeRows) || kundeRows[0] || {};

  const customerName =
    kunde.name ||
    kunde.Name ||
    kunde.company ||
    kunde.companyName ||
    '';

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
    street: kunde.street || kunde.Straße || '',
    zip: kunde.zip || kunde.PLZ || '',
    city: kunde.city || kunde.Ort || '',
    countryCode: kunde.countryCode || 'DE',
    email: kunde.email || '',
    contactPerson: kunde.contactPerson || ''
  };

  // -------- Positionen --------
  const lineItems = [];

  positions.forEach((row, idx) => {
    const excelRow = idx + 2; // erste Datenzeile

    const type = (row.type || '').toString().trim();      // custom / text / material
    const articleId = (row.articleId || row.articleID || '').toString().trim();
    const name = (row.name || '').toString().trim();
    const description = (row.description || '').toString().trim();

    const qtyRaw =
      row.quantity ??
      row.qty ??
      row.Qty ??
      row.Menge ??
      row.menge ??
      0;

    const qty = Number(qtyRaw);

    const unitName = (row.unitName || row.unit || '').toString().trim();

    const priceRaw =
      row.unitPriceAmount ??
      row.price ??
      row.Preis ??
      null;

    const priceExcel =
      priceRaw !== '' && priceRaw !== null ? Number(priceRaw) : null;

    const taxRate = row.taxRatePercentage ?? row.taxRate ?? null;
    const discountPercent = row.discountPercent ?? row.discount ?? null;

    const hasAny =
      type || articleId || name || description || qtyRaw || priceRaw;
    if (!hasAny) return;

    if (!type) {
      errors.push({
        sheet: 'Positionen',
        row: excelRow,
        field: 'type',
        message: 'type ist Pflicht.'
      });
      return;
    }

    if (!byType[type]) byType[type] = 0;
    byType[type]++;

    // text-Position → keine Menge nötig
    if (type === 'text') {
      lineItems.push({
        type: 'text',
        name: name || description || `Textzeile ${excelRow}`,
        description: description || null,
        quantity: null,
        unitName: null,
        price: null,
        taxRate,
        discountPercent
      });
      return;
    }

    // alle anderen Typen → qty > 0
    if (!(qty > 0)) {
      errors.push({
        sheet: 'Positionen',
        row: excelRow,
        field: 'qty',
        message: 'quantity / qty muss größer als 0 sein.'
      });
      return;
    }

    let useExcelPrice = false;

    if (!articleId) {
      // keine articleId → Preis MUSS aus Excel kommen
      if (priceExcel === null || Number.isNaN(priceExcel)) {
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
      // articleId vorhanden
      if (type === 'material') {
        // Material → immer Preis aus Artikelstamm
        useExcelPrice = false;
      } else if (allowPriceOverride) {
        // Override erlaubt → Excel-Preis darf überschreiben (wenn vorhanden)
        useExcelPrice = priceExcel != null && !Number.isNaN(priceExcel);
      } else {
        // sonst Preis aus Lexoffice
        useExcelPrice = false;
      }
    }

    let finalName = name;
    if (!finalName && articleId) {
      finalName = `(Artikel ${articleId})`;
      autoNamed.push({ row: excelRow, articleId, name: finalName });
      warnings.push({
        sheet: 'Positionen',
        row: excelRow,
        message: `Name war leer und wurde aus Artikel ${articleId} ergänzt.`
      });
    }

    lineItems.push({
      type: type === 'material' ? 'material' : 'custom',
      articleId: articleId || null,
      name: finalName || `Position ${excelRow}`,
      description: description || null,
      quantity: qty,
      unitName: unitName || null,
      price: useExcelPrice ? priceExcel : null,
      taxRate,
      discountPercent
    });
  });

  if (errors.length) {
    return { quotation: null, summary: { errors, warnings, autoNamed, byType } };
  }

  // WICHTIG: voucherDate jetzt als vollständiger ISO-String (Lexoffice-Fix)
  const quotation = {
    voucherDate: new Date().toISOString(),
    taxType,
    address,
    lineItems,
    finalized: true
  };

  return { quotation, summary: { errors, warnings, autoNamed, byType } };
}

// ----------------------------------------------------
// Lexoffice-Anfrage
// ----------------------------------------------------

async function createQuotation(payload) {
  if (!LEXOFFICE_API_KEY) {
    const e = new Error('LEXOFFICE_API_KEY ist nicht gesetzt.');
    e.status = 500;
    throw e;
  }

  const res = await callLexoffice({
    method: 'POST',
    url: 'https://api.lexoffice.io/v1/quotations',
    headers: {
      Authorization: `Bearer ${LEXOFFICE_API_KEY}`,
      'Content-Type': 'application/json',
      Accept: 'application/json'
    },
    data: payload,
    validateStatus: () => true
  });

  if (res.status === 429) {
    const e = new Error('Rate limit exceeded');
    e.status = 429;
    e.raw = res.data;
    throw e;
  }

  if (res.status < 200 || res.status >= 300) {
    const e = new Error('Lexoffice API Fehler');
    e.status = res.status;
    e.raw = res.data;
    throw e;
  }

  return res.data;
}

// ----------------------------------------------------
// API-Routen
// ----------------------------------------------------

app.get('/api/ping', (req, res) => {
  res.json({
    ok: true,
    passwordProtected: !!TOOL_PASSWORD,
    allowPriceOverrideDefault: ALLOW_PRICE_OVERRIDE_DEFAULT,
    minIntervalMs: MIN_INTERVAL_MS
  });
});

app.post('/api/test-excel', passwordMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body || {};

    if (!excelData) {
      return res.json({
        ok: false,
        stage: 'input',
        status: 'VALIDATION_ERROR',
        message: 'Es wurden keine Excel-Daten übergeben.',
        data: { summary: { errors: [], warnings: [], autoNamed: [], byType: {} } }
      });
    }

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
      message: hasErrors
        ? 'Excel enthält Validierungsfehler.'
        : 'Test erfolgreich.',
      data: { summary: parsed.summary }
    });
  } catch (err) {
    res.json({
      ok: false,
      stage: 'test',
      status: 'ERROR',
      message: err.message,
      technical: { status: err.status || null }
    });
  }
});

app.post('/api/create-offer', passwordMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body || {};

    if (!excelData) {
      return res.json({
        ok: false,
        stage: 'input',
        status: 'VALIDATION_ERROR',
        message: 'Es wurden keine Excel-Daten übergeben.',
        data: { summary: { errors: [], warnings: [], autoNamed: [], byType: {} } }
      });
    }

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
        message: 'Excel enthält Validierungsfehler.',
        data: { summary: parsed.summary }
      });
    }

    const result = await createQuotation(parsed.quotation);

    res.json({
      ok: true,
      stage: 'lexoffice-create',
      status: 'SUCCESS',
      message: 'Angebot erstellt.',
      data: {
        quotationId: result.id,
        summary: parsed.summary
      }
    });
  } catch (err) {
    res.json({
      ok: false,
      stage: 'lexoffice-create',
      status: err.status === 429 ? 'RATE_LIMIT' : 'ERROR',
      message: err.message,
      technical: {
        httpStatus: err.status || null,
        raw: err.raw || null
      }
    });
  }
});

// ----------------------------------------------------

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server läuft auf Port', PORT));
