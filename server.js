// server.js
// Maiershirts — Lexware Angebots-Tool

require('dotenv').config();
const express = require('express');
const path = require('path');
const fs = require('fs');
const axios = require('axios');
const XLSX = require('xlsx');

const app = express();

// ---------- ENV / Konfiguration ----------
const PORT = process.env.PORT || 3000;

const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';
const LEXWARE_API_KEY = process.env.LEXWARE_API_KEY || process.env.LEXOFFICE_API_KEY || '';
const LEXWARE_API_BASE_URL =
  process.env.LEXWARE_API_BASE_URL ||
  process.env.LEXOFFICE_API_BASE_URL ||
  'https://api.lexware.io';

const MIN_CALL_INTERVAL_MS = Number(process.env.MIN_CALL_INTERVAL_MS || '600');
const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE_DEFAULT || '').toLowerCase() === 'true';

// einfacher Zeitstempel für Rate-Limit
let lastLexwareCallTs = 0;

// ---------- Middleware ----------
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// statische Dateien: dein vorhandener public-Ordner + Template-Download
app.use(express.static(path.join(__dirname, 'public')));

// Template explizit mit korrektem Dateinamen ausliefern
// URL in der UI: /templates/lexware_template.xlsx
app.get('/templates/lexware_template.xlsx', (req, res) => {
  const filePath = path.join(__dirname, 'templates', 'Lexware_Template.xlsx');
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('Template-Datei nicht gefunden (Lexware_Template.xlsx).');
  }
  res.sendFile(filePath);
});

// Passwort-Middleware (optional, je nach .env)
function checkPassword(req, res, next) {
  if (!TOOL_PASSWORD) return next(); // kein Passwort konfiguriert

  const pw = req.body.password || req.query.password || '';
  if (pw !== TOOL_PASSWORD) {
    return res.status(401).json({
      ok: false,
      status: 'UNAUTHORIZED',
      message: 'Passwort ist falsch oder fehlt.',
    });
  }
  next();
}

// ---------- Helper: Excel lesen ----------

function readFieldValueSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;

  // Annahme: Zeile 1: Überschriften: "Feld" | "Wert" | "Hinweis"
  // Wir lesen ab Zeile 2 (range:1)
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: ['field', 'value', 'hint'],
    range: 1,
    defval: '',
  });

  const map = {};
  rows.forEach((row) => {
    const key = (row.field || '').toString().trim();
    if (!key) return;
    map[key] = (row.value || '').toString().trim();
  });

  return map;
}

function parseVoucherDate(excelValue) {
  if (!excelValue) return null;

  // Falls Datum schon RFC3339 mit "T" ist, einfach durchreichen
  if (typeof excelValue === 'string' && excelValue.includes('T')) {
    return excelValue;
  }

  // Falls es ein String "YYYY-MM-DD" ist
  if (typeof excelValue === 'string') {
    const m = excelValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      const [_, y, mo, d] = m;
      return `${y}-${mo}-${d}T00:00:00.000+01:00`;
    }
  }

  // Falls es ein Zahlwert (Excel-Datum) ist, versuchen wir, ihn zu einem JS-Datum umzuwandeln
  if (typeof excelValue === 'number') {
    // sehr einfache Excel-Konvertierung (ohne 1900-Leap-Bug-Behandlung, hier völlig ausreichend)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
    const ms = excelValue * 24 * 60 * 60 * 1000;
    const d = new Date(excelEpoch.getTime() + ms);

    const year = d.getUTCFullYear();
    const month = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${year}-${month}-${day}T00:00:00.000+01:00`;
  }

  // Notfall: Heute
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}T00:00:00.000+01:00`;
}

function calcExpirationDateFromVoucherDate(voucherDateRFC3339) {
  try {
    // "YYYY-MM-DDT..." -> wir parsen nur Datumsteil
    const baseStr = voucherDateRFC3339.substring(0, 10);
    const [y, m, d] = baseStr.split('-').map((x) => parseInt(x, 10));
    const dt = new Date(y, m - 1, d);
    dt.setDate(dt.getDate() + 30); // +30 Tage

    const year = dt.getFullYear();
    const month = String(dt.getMonth() + 1).padStart(2, '0');
    const day = String(dt.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}T00:00:00.000+01:00`;
  } catch (e) {
    return voucherDateRFC3339; // Fallback: gleiches Datum
  }
}

function parseExcelAndBuildQuotationPayload(excelBase64, options) {
  const errors = [];
  const warnings = [];

  const buf = Buffer.from(excelBase64, 'base64');
  let workbook;
  try {
    workbook = XLSX.read(buf, { type: 'buffer' });
  } catch (e) {
    throw new Error('Excel-Datei konnte nicht gelesen werden.');
  }

  // Pflicht-Sheets checken
  const requiredSheets = ['Angebot', 'Kunde', 'Positionen'];
  requiredSheets.forEach((name) => {
    if (!workbook.Sheets[name]) {
      errors.push({
        sheet: name,
        row: null,
        field: null,
        message: `Pflicht-Tabelle "${name}" fehlt.`,
      });
    }
  });

  if (errors.length) {
    return {
      summary: { errors, warnings },
      readyForCreate: false,
      quotation: null,
    };
  }

  // --- Angebot ---
  const angebotMap = readFieldValueSheet(workbook, 'Angebot') || {};
  const rawVoucherDate = angebotMap.voucherDate || angebotMap.datum || '';
  const rawTaxType = (angebotMap.taxType || '').trim();

  if (!rawTaxType) {
    errors.push({
      sheet: 'Angebot',
      row: 2,
      field: 'taxType',
      message: 'taxType ist Pflicht (z. B. "net" oder "gross").',
    });
  }

  let voucherDate = parseVoucherDate(rawVoucherDate || null);
  if (!rawVoucherDate) {
    warnings.push({
      sheet: 'Angebot',
      row: 2,
      field: 'voucherDate',
      message: 'voucherDate war leer. Es wurde automatisch ein Datum gesetzt.',
    });
  }

  const expirationDate =
    angebotMap.expirationDate && angebotMap.expirationDate.trim()
      ? parseVoucherDate(angebotMap.expirationDate.trim())
      : calcExpirationDateFromVoucherDate(voucherDate);

  const taxConditions = rawTaxType
    ? { taxType: rawTaxType }
    : null;

  if (!taxConditions) {
    // falls oben noch nicht als Fehler behandelt
    errors.push({
      sheet: 'Angebot',
      row: 2,
      field: 'taxConditions',
      message: 'taxConditions.taxType ist Pflicht.',
    });
  }

  // optionale Texte
  const title = angebotMap.title || 'Angebot';
  const introduction = angebotMap.introduction || 'Gerne bieten wir Ihnen an:';
  const remark =
    angebotMap.remark ||
    'Wir freuen uns auf Ihre Auftragserteilung und sichern eine einwandfreie Ausführung zu.';

  // --- Kunde ---
  const kundeMap = readFieldValueSheet(workbook, 'Kunde') || {};
  const customerName = kundeMap.name || '';
  if (!customerName) {
    errors.push({
      sheet: 'Kunde',
      row: 2,
      field: 'name',
      message: 'Kundenname ist Pflicht.',
    });
  }

  const address = {};
  const contactId = (kundeMap.contactId || '').trim();
  if (contactId) {
    address.contactId = contactId;
  } else {
    address.name = customerName;
    address.street = kundeMap.street || undefined;
    address.city = kundeMap.city || undefined;
    address.zip = kundeMap.zip || undefined;
    address.countryCode = (kundeMap.countryCode || 'DE').trim() || 'DE';
    if (!address.countryCode) {
      errors.push({
        sheet: 'Kunde',
        row: 2,
        field: 'countryCode',
        message: 'countryCode ist Pflicht, z. B. "DE".',
      });
    }
  }

  // --- Positionen ---
  const posSheet = workbook.Sheets['Positionen'];
  const posRows = XLSX.utils.sheet_to_json(posSheet, { defval: '' });

  const lineItems = [];
  const allowPriceOverride = !!options.allowPriceOverride;

  posRows.forEach((row, i) => {
    const excelRow = i + 2; // Zeile 1 = Header

    const typeRaw = (row.type || '').toString().trim().toLowerCase();
    const articleId = (row.articleId || '').toString().trim() || null;
    const nameRaw = (row.name || '').toString().trim();
    const descRaw = (row.description || '').toString().trim();
    const unitNameRaw = (row.unitName || '').toString().trim();

    const qty = row.quantity !== undefined && row.quantity !== ''
      ? Number(row.quantity)
      : (row.qty !== undefined && row.qty !== '' ? Number(row.qty) : 0);

    const priceAmount =
      row.unitPriceAmount !== undefined && row.unitPriceAmount !== ''
        ? Number(row.unitPriceAmount)
        : null;

    const taxRate =
      row.taxRatePercentage !== undefined && row.taxRatePercentage !== ''
        ? Number(row.taxRatePercentage)
        : null;

    const discount =
      row.discountPercent !== undefined && row.discountPercent !== ''
        ? Number(row.discountPercent)
        : 0;

    // komplett leere Zeile -> ignorieren
    if (!typeRaw && !articleId && !nameRaw && !descRaw && !qty && !unitNameRaw) {
      return;
    }

    const typeAllowed = ['material', 'service', 'custom', 'text'];
    if (!typeAllowed.includes(typeRaw)) {
      errors.push({
        sheet: 'Positionen',
        row: excelRow,
        field: 'type',
        message: 'Ungültiger type. Erlaubt: material, service, custom, text.',
      });
      return;
    }

    // Validierung Menge / Einheit / Preis
    if (typeRaw !== 'text') {
      if (!qty || isNaN(qty) || qty <= 0) {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'quantity',
          message: 'qty muss größer als 0 sein.',
        });
      }
      if (!unitNameRaw) {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'unitName',
          message: 'unitName ist Pflicht für material/service/custom.',
        });
      }

      // Lexware fordert unitPrice *immer* für material/service/custom
      if (!priceAmount && !articleId) {
        errors.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'unitPriceAmount',
          message:
            'unitPriceAmount ist Pflicht, wenn kein articleId gesetzt ist.',
        });
      }

      if (!priceAmount && articleId && !allowPriceOverride) {
        // Hier können wir entweder Fehler werfen oder warnen.
        // Ich werfe eine *Warnung* und lasse die Position zu, um kompatibel zu bleiben.
        warnings.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'unitPriceAmount',
          message:
            'unitPriceAmount ist leer, articleId ist gesetzt. Lexware berechnet den Preis aus dem Artikelstamm.',
        });
      }
    }

    let finalName = nameRaw;
    if (!finalName) {
      if (articleId) {
        finalName = `Artikel ${articleId}`;
        warnings.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'name',
          message:
            'Name war leer. Es wurde automatisch "Artikel {articleId}" gesetzt.',
        });
      } else {
        finalName = `Position ${excelRow}`;
        warnings.push({
          sheet: 'Positionen',
          row: excelRow,
          field: 'name',
          message:
            'Name war leer. Es wurde automatisch "Position {Zeile}" gesetzt.',
        });
      }
    }

    const lineItem = {
      type: typeRaw, // material / service / custom / text
      name: finalName,
    };

    if (descRaw) lineItem.description = descRaw;

    if (typeRaw !== 'text') {
      lineItem.quantity = qty;
      lineItem.unitName = unitNameRaw;

      // unitPrice nur senden, wenn wir einen Wert haben (oder Preis-Override erlaubt ist)
      if (priceAmount || allowPriceOverride) {
        const taxRatePercentage =
          taxRate !== null && !isNaN(taxRate) ? taxRate : 19;

        // je nach taxConditions.taxType net oder gross befüllen
        if (taxConditions && taxConditions.taxType === 'gross') {
          lineItem.unitPrice = {
            currency: 'EUR',
            grossAmount: priceAmount || 0,
            taxRatePercentage,
          };
        } else {
          lineItem.unitPrice = {
            currency: 'EUR',
            netAmount: priceAmount || 0,
            taxRatePercentage,
          };
        }
      }

      // Discount
      if (discount && !isNaN(discount)) {
        lineItem.discountPercentage = discount;
      }
    }

    // WICHTIG: Artikel-Referenz -> id (nicht articleId)
    if (articleId && (typeRaw === 'material' || typeRaw === 'service')) {
      lineItem.id = articleId;
    }

    lineItems.push(lineItem);
  });

  if (!lineItems.length) {
    errors.push({
      sheet: 'Positionen',
      row: null,
      field: null,
      message: 'Es wurden keine gültigen Positionen gefunden.',
    });
  }

  const summary = { errors, warnings };

  if (errors.length) {
    return {
      summary,
      readyForCreate: false,
      quotation: null,
    };
  }

  // Quotation-Payload nach Lexware-Doku
  // https://developers.lexware.io/docs/ (Quotations Endpoint)
  const quotation = {
    voucherDate,
    expirationDate,
    address,
    lineItems,
    totalPrice: {
      currency: 'EUR',
    },
    taxConditions,
    title,
    introduction,
    remark,
  };

  return {
    summary,
    readyForCreate: true,
    quotation,
  };
}

// ---------- Helper: Lexware API Call ----------

async function createLexwareQuotation(quotationPayload) {
  const url = `${LEXWARE_API_BASE_URL.replace(/\/+$/, '')}/v1/quotations?finalize=true`;

  const now = Date.now();
  const diff = now - lastLexwareCallTs;
  if (diff < MIN_CALL_INTERVAL_MS) {
    const waitMs = MIN_CALL_INTERVAL_MS - diff;
    await new Promise((res) => setTimeout(res, waitMs));
  }

  lastLexwareCallTs = Date.now();

  try {
    const res = await axios.post(url, quotationPayload, {
      headers: {
        Authorization: `Bearer ${LEXWARE_API_KEY}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      validateStatus: () => true, // wir behandeln Fehler manuell
    });

    return {
      httpStatus: res.status,
      data: res.data,
    };
  } catch (e) {
    if (e.response) {
      return {
        httpStatus: e.response.status,
        data: e.response.data,
      };
    }
    throw e;
  }
}

// ---------- API-Routen ----------

// Systemstatus
app.get('/api/ping', (req, res) => {
  res.json({
    ok: true,
    status: 'OK',
    passwordProtected: !!TOOL_PASSWORD,
    allowPriceOverrideDefault: ALLOW_PRICE_OVERRIDE_DEFAULT,
    minIntervalMs: MIN_CALL_INTERVAL_MS,
  });
});

// Testmodus: Excel nur prüfen, nichts in Lexware anlegen
app.post('/api/test-excel', checkPassword, (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body || {};
    if (!excelData) {
      return res.status(400).json({
        ok: false,
        status: 'NO_EXCEL',
        message: 'Es wurde keine Excel-Datei übertragen.',
      });
    }

    const result = parseExcelAndBuildQuotationPayload(excelData, {
      allowPriceOverride,
    });

    if (!result.readyForCreate) {
      return res.status(200).json({
        ok: false,
        status: 'VALIDATION_ERROR',
        message: 'Excel enthält Validierungsfehler. Details siehe errors.',
        data: {
          summary: result.summary,
        },
      });
    }

    // Test erfolgreich
    return res.status(200).json({
      ok: true,
      status: 'OK',
      message: 'Test erfolgreich — keine kritischen Fehler.',
      data: {
        summary: result.summary,
        quotationPreview: result.quotation,
      },
    });
  } catch (e) {
    console.error('Fehler im Testmodus:', e);
    return res.status(500).json({
      ok: false,
      status: 'ERROR',
      message: 'Interner Serverfehler im Testmodus.',
    });
  }
});

// Angebot in Lexware anlegen
app.post('/api/create-offer', checkPassword, async (req, res) => {
  try {
    const { excelData, allowPriceOverride } = req.body || {};
    if (!excelData) {
      return res.status(400).json({
        ok: false,
        status: 'NO_EXCEL',
        message: 'Es wurde keine Excel-Datei übertragen.',
      });
    }

    const parsed = parseExcelAndBuildQuotationPayload(excelData, {
      allowPriceOverride,
    });

    if (!parsed.readyForCreate) {
      return res.status(200).json({
        ok: false,
        status: 'VALIDATION_ERROR',
        message: 'Excel enthält Validierungsfehler. Details siehe errors.',
        data: {
          summary: parsed.summary,
        },
      });
    }

    // API-Key prüfen
    if (!LEXWARE_API_KEY) {
      return res.status(500).json({
        ok: false,
        status: 'CONFIG_ERROR',
        stage: 'lexware-create',
        message: 'LEXWARE_API_KEY ist nicht konfiguriert.',
      });
    }

    const apiResult = await createLexwareQuotation(parsed.quotation);

    if (apiResult.httpStatus < 200 || apiResult.httpStatus >= 300) {
      // Lexware-Fehler durchreichen
      return res.status(200).json({
        ok: false,
        status: 'ERROR',
        stage: 'lexware-create',
        message: 'Lexware API Fehler',
        httpStatus: apiResult.httpStatus,
        technical: apiResult.data || null,
      });
    }

    // Erfolgreich: Lexware gibt id + resourceUri etc. zurück
    const quotationId = apiResult.data && apiResult.data.id;

    return res.status(200).json({
      ok: true,
      status: 'OK',
      message: 'Angebot erfolgreich in Lexware erstellt.',
      data: {
        quotationId,
        summary: parsed.summary,
        technical: apiResult.data,
      },
    });
  } catch (e) {
    console.error('Fehler beim Erstellen des Angebots:', e);
    return res.status(500).json({
      ok: false,
      status: 'ERROR',
      stage: 'lexware-create',
      message: 'Unerwarteter Fehler beim Erstellen des Angebots.',
    });
  }
});

// Fallback-Route für "Cannot GET /"
app.get('/', (req, res) => {
  // Wichtig: deine vorhandene public/index.html wird durch express.static ausgeliefert.
  // Hier nur kurzer Hinweis, falls etwas schief geht.
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ---------- Start ----------
app.listen(PORT, () => {
  console.log(`Maiershirts Lexware-Tool läuft auf Port ${PORT}`);
});
