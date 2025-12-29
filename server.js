'use strict';

/**
 * Maiershirts — Lexoffice Angebots-Tool Backend
 * --------------------------------------------
 * Funktionen:
 *  - Passwortschutz (TOOL_PASSWORD, optional)
 *  - Testmodus / Validierung (/api/test-excel)
 *  - Angebot erstellen (/api/create-offer)
 *  - Preis-Override-Schalter (ALLOW_PRICE_OVERRIDE, Checkbox)
 *  - Rate-Limiter (LEXWARE_MIN_INTERVAL_MS, Standard 600ms)
 *  - Detaillierte Fehlermeldungen (errors mit Sheet/Row/Field)
 *  - Liefert Frontend aus /public
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

// Frontend aus /public ausliefern (index.html, CSS, JS, …)
app.use(express.static(path.join(__dirname, 'public')));

// ----------------------------------------------------
// ENV Variablen
// ----------------------------------------------------

const LEXOFFICE_API_KEY = process.env.LEXOFFICE_API_KEY || '';
const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';
const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE || 'false').toLowerCase() === 'true';
const MIN_INTERVAL_MS = parseInt(process.env.LEXWARE_MIN_INTERVAL_MS || '600', 10);

console.log('[INFO] Rate-Limiter aktiv mit', MIN_INTERVAL_MS, 'ms Abstand');

// ----------------------------------------------------
// Rate Limiter
// ----------------------------------------------------

let lastCall = 0;

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function callLexoffice(config) {
  const diff = Date.now() - lastCall;
  if (diff < MIN_INTERVAL_MS) {
    await sleep(MIN_INTERVAL_MS - diff);
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

  // excelData kann Base64 oder Buffer sein
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
    return {
      quotation: null,
      summary: { errors, warnings, autoNamed, byType }
    };
  }

  // Angebot
  const angebotRow = angebot[0] || {};
  const taxType = angebotRow.taxType || angebotRow.TAXTYPE || angebotRow.tax || '';

  if (!taxType) {
    errors.push({
      sheet: 'Angebot',
      row: 2,
      field: 'taxType',
      message: 'taxType ist Pflicht (z. B. „gross“ oder „net“).'
    });
  }

  // Kunde
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

  // Positionen
  const lineItems = [];

  positionen.forEach((row, idx) => {
    const excelRow = idx + 2;

    const type = (row.type || row.Typ || '').toString().trim();
    const qty = Number(row.qty || row.Menge || 0);
    const articleId = (row.articleId || row.articleID || '').toString().trim();
    const priceExcel = row.price || row.Preis || null;
    const name = (row.name || row.Bezeichnung || '').toString().trim();

    if (!type) {
      errors.push({
        sheet: 'Positionen',
        row: excelRow,
        field: 'type',
        message: 'type ist Pflicht (z. B. „material“ oder „service“).'
      });
      return;
    }

    if (!(qty > 0)) {
      errors.push({
        sheet: 'Positionen',
        row: excelRow,
        field: 'qty',
        message: 'qty muss größer als 0 sein.'
      });
      return;
    }

    if (!byType[type]) byType[type] = 0;
    byType[type]++;

    let usePrice = false;

    if (!articleId) {
      if (!priceExcel && priceExcel !== 0) {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'price',
          message: 'Preis ist Pflicht, wenn keine articleId gesetzt ist.'
        });
        return;
      }
      usePrice = true;
    } else {
      if (type === 'material') {
        usePrice = false; // Preis kommt aus Lexoffice-Artikel
      } else if (allowPriceOverride) {
        usePrice = true;  // Excel-Preis darf überschreiben
      } else {
        usePrice = false; // Standard: Preis aus Lexoffice
      }
    }

    let finalName = name;

    if (!finalName && articleId) {
      finalName = `(Artikel ${articleId})`;
      autoNamed.push({ row: excelRow, articleId, name: finalName });
      warnings.push({
        row: excelRow,
        message: `Name war leer und wurde aus Artikel ${articleId} ergänzt (Platzhalter).`
      });
    }

    lineItems.push({
      type: 'custom',
      articleId: articleId || null,
      name: finalName || `Position ${excelRow}`,
      quantity: qty,
      price: usePrice ? Number(priceExcel) : null
      // Preis aus Artikelstamm wird später in Lexoffice gezogen
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

  return {
    quotation,
    summary: { errors, warnings, autoNamed, byType }
  };
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
