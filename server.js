'use strict';

/**
 * Maiershirts — Lexoffice Angebots-Tool Backend
 * --------------------------------------------
 * - Frontend aus /public
 * - Passwortschutz (TOOL_PASSWORD, optional)
 * - Testmodus (/api/test-excel)
 * - Angebot erstellen (/api/create-offer)
 * - Preis-Override-Schalter (ALLOW_PRICE_OVERRIDE)
 * - Rate-Limiter (LEXWARE_MIN_INTERVAL_MS, Standard 600ms)
 * - Unterstützt Excel-Columns wie im Screenshot:
 *   pos, type, articleId, name, description, quantity, unitName,
 *   unitPriceAmount, taxRatePercentage, discountPercent
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

// Frontend aus /public
app.use(express.static(path.join(__dirname, 'public')));

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

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

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
      message: 'Passwort ungültig oder fehlt.'
    });
  }

  next();
}

// ----------------------------------------------------
// Excel Parsing & Validierung
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
    const err = new Error('excelData Format ungültig (erwartet Base64-String oder Buffer)');
    err.status = 400;
    throw err;
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

  // Pflicht-Sheets
  if (!angebot) errors.push({ sheet: 'Angebot', message: 'Sheet „Angebot“ fehlt.' });
  if (!kunde) errors.push({ sheet: 'Kunde', message: 'Sheet „Kunde“ fehlt.' });
  if (!positionen) errors.push({ sheet: 'Positionen', message: 'Sheet „Positionen“ fehlt.' });

  if (errors.length) {
    return { quotation: null, summary: { errors, warnings, autoNamed, byType } };
  }

  // ---------------- Angebot ----------------
  const angebotRow = angebot[0] || {};
  const taxType =
    angebotRow.taxType ||
    angebotRow.TAXTYPE ||
    angebotRow.tax ||
    ''; // z.B. "net" oder "gross"

  if (!taxType) {
    errors.push({
      sheet: 'Angebot',
      row: 2,
      field: 'taxType',
      message: 'taxType ist Pflicht (z. B. „net“ oder „gross“).'
    });
  }

  // ---------------- Kunde ----------------
  const kundeRow = kunde[0] || {};
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

  // ---------------- Positionen ----------------
  const lineItems = [];

  positionen.forEach((row, idx) => {
    const excelRow = idx + 2; // erste Datenzeile ist Excel-Zeile 2

    const type = (row.type || '').toString().trim(); // custom / text / material
    const articleId = (row.articleId || row.articleID || '').toString().trim();
    const name = (row.name || '').toString().trim();
    const description = (row.description || '').toString().trim();

    // WICHTIG: quantity aus mehreren möglichen Spalten lesen
    const qtyRaw =
      row.quantity ??
      row.qty ??
      row.Qty ??
      row.Menge ??
      row.menge ??
      0;
    const qty = Number(qtyRaw);

    const unitName = (row.unitName || row.unit || '').toString().trim();

    // Preis aus `unitPriceAmount` ODER `price` / `Preis`
    const unitPriceRaw =
      row.unitPriceAmount ??
      row.price ??
      row.Preis ??
      null;
    const priceExcel = unitPriceRaw !== '' && unitPriceRaw !== null
      ? Number(unitPriceRaw)
      : null;

    const taxRate = row.taxRatePercentage ?? row.taxRate ?? null;
    const discountPercent = row.discountPercent ?? row.discount ?? null;

    // Leere Zeile komplett überspringen
    const hasAnyValue =
      type || articleId || name || description || qtyRaw || unitPriceRaw;
    if (!hasAnyValue) return;

    // Typ prüfen
    if (!type) {
      errors.push({
        sheet: 'Positionen',
        row: excelRow,
        field: 'type',
        message: 'type ist Pflicht (z. B. „custom“, „text“, „material“).'
      });
      return;
    }

    if (!byType[type]) byType[type] = 0;
    byType[type]++;

    // Sonderfall: text-Position
    if (type === 'text') {
      // Für Textzeilen ist qty optional, kein Preis nötig
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

    // Für alle anderen Typen ist qty > 0 Pflicht
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
      // Keine articleId → Preis MUSS aus Excel kommen
      if (priceExcel === null || Number.isNaN(priceExcel)) {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'price',
          message: 'Preis (unitPriceAmount / price) ist Pflicht, wenn keine articleId gesetzt ist.'
        });
        return;
      }
      useExcelPrice = true;

    } else {
      // Es gibt eine articleId
      if (type === 'material') {
        // Material: Preis immer aus Artikelstamm
        useExcelPrice = false;
      } else if (allowPriceOverride) {
        // Override erlaubt: Excel-Preis darf überschreiben
        useExcelPrice = priceExcel !== null && !Number.isNaN(priceExcel);
      } else {
        // Standard: Preis aus Lexoffice
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
        message: `Name war leer und wurde aus Artikel ${articleId} (Platzhalter) ergänzt.`
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

  // Minimal-Quotation für Lexoffice (die API errechnet Preise/Steuern anhand Artikel)
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
    const msg = hasErrors
      ? 'Excel enthält Validierungsfehler. Details siehe errors.'
      : 'Test erfolgreich. Keine kritischen Fehler.';

    res.json({
      ok: !hasErrors,
      stage: 'test',
      status: hasErrors ? 'VALIDATION_ERROR' : 'SUCCESS',
      message: msg,
      data: { summary: parsed.summary }
    });

  } catch (err) {
    console.error('[test-excel] Fehler:', err);

    res.json({
      ok: false,
      stage: 'test',
      status: 'ERROR',
      message: err.message || 'Unbekannter Fehler im Testmodus.',
      technical: {
        status: err.status || null
      }
    });
  }
});

// Angebot erstellen
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
        message: 'Excel enthält Validierungsfehler. Angebot wurde NICHT erstellt.',
        data: { summary: parsed.summary }
      });
    }

    const result = await createQuotation(parsed.quotation);

    res.json({
      ok: true,
      stage: 'lexoffice-create',
      status: 'SUCCESS',
      message: 'Angebot erfolgreich in Lexoffice erstellt.',
      data: {
        quotationId: result.id,
        summary: parsed.summary
      }
    });

  } catch (err) {
    console.error('[create-offer] Fehler:', err);

    if (err.status === 429) {
      return res.json({
        ok: false,
        stage: 'lexoffice-create',
        status: 'RATE_LIMIT',
        message: 'Lexoffice Rate-Limit überschritten. Bitte später erneut versuchen.',
        technical: { status: 429 }
      });
    }

    res.json({
      ok: false,
      stage: 'lexoffice-create',
      status: 'ERROR',
      message: err.message || 'Unbekannter Fehler bei der Angebotserstellung.',
      technical: {
        status: err.status || null
      }
    });
  }
});

// ----------------------------------------------------

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server läuft auf Port', PORT));
