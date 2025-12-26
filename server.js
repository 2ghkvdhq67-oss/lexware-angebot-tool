// server.js
// ------------------------------------------------------
// Lexoffice / Lexware Angebots-Tool Backend (Maiershirts)
// - Passwortschutz (optional via ENV TOOL_PASSWORD)
// - Schalter: allowPriceOverride (ENV + Request-Body)
// - Endpoints: /api/test-excel, /api/create-offer, /api/ping
// - Zentraler Rate-Limiter für Lexware-API (min. Abstand)
// - Saubere Rate-Limit-Behandlung (HTTP 429)
// - Garantiert nur EIN Create-Call pro Anfrage
// ------------------------------------------------------

'use strict';

const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const dotenv = require('dotenv');

dotenv.config();

const app = express();

// Body Parser für JSON (Frontend sendet JSON)
app.use(bodyParser.json({ limit: '10mb' }));

// ------------------------------------------------------
// Konfiguration über ENV Variablen
// ------------------------------------------------------
//
// LEXOFFICE_API_KEY         = dein Lexware/Lexoffice API-Key
// TOOL_PASSWORD             = optionales Passwort für UI (leer = kein Schutz)
// ALLOW_PRICE_OVERRIDE      = 'true' oder 'false' (Default für Schalter)
// LEXWARE_MIN_INTERVAL_MS   = min. Abstand zwischen API-Calls in ms (Default: 600)
// PORT                      = Port für Express (z. B. 3000)
//
// Beispiel .env:
//
// LEXOFFICE_API_KEY=xxxx
// TOOL_PASSWORD=meinGeheimesPasswort
// ALLOW_PRICE_OVERRIDE=false
// LEXWARE_MIN_INTERVAL_MS=600
// PORT=3000
// ------------------------------------------------------

const LEXOFFICE_API_KEY = process.env.LEXOFFICE_API_KEY || '';
const TOOL_PASSWORD = process.env.TOOL_PASSWORD || '';
const ALLOW_PRICE_OVERRIDE_DEFAULT =
  (process.env.ALLOW_PRICE_OVERRIDE || 'false').toLowerCase() === 'true';

const MIN_INTERVAL_MS = parseInt(process.env.LEXWARE_MIN_INTERVAL_MS || '600', 10);

if (!LEXOFFICE_API_KEY) {
  console.warn(
    '[WARNUNG] LEXOFFICE_API_KEY ist nicht gesetzt. Lexware/Lexoffice-Aufrufe werden fehlschlagen.'
  );
}

console.log('[INFO] MIN_INTERVAL_MS (Rate-Limit-Puffer):', MIN_INTERVAL_MS, 'ms');

// ------------------------------------------------------
// Zentraler Rate-Limiter für alle Lexware/Lexoffice API-Calls
// ------------------------------------------------------
//
// Lexware erlaubt 2 Requests/Sekunde. Durch einen Mindestabstand von
// z.B. 600–700 ms bleiben wir sicher darunter (Puffer für Netzwerk-Jitter).
//
// Alle API-Calls Richtung Lexware/Lexoffice SOLLEN über diese Funktion laufen.
// ------------------------------------------------------

let lastLexwareCallTime = 0;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Wrapper für alle Lexware/Lexoffice-Requests.
 * Stellt sicher, dass zwischen zwei Aufrufen mindestens MIN_INTERVAL_MS vergeht.
 *
 * @param {Object} axiosConfig - Config für axios (method, url, headers, data, etc.)
 * @returns axios Response
 */
async function callLexwareApi(axiosConfig) {
  const now = Date.now();
  const elapsed = now - lastLexwareCallTime;

  if (elapsed < MIN_INTERVAL_MS) {
    const waitMs = MIN_INTERVAL_MS - elapsed;
    // Optionale Logging-Ausgabe:
    // console.log(`[RateLimiter] Warte ${waitMs} ms vor dem nächsten Lexware-Call`);
    await sleep(waitMs);
  }

  // Jetzt API-Call ausführen
  const res = await axios(axiosConfig);

  // Zeitstempel aktualisieren
  lastLexwareCallTime = Date.now();

  return res;
}

// ------------------------------------------------------
// Middleware: einfacher Passwortschutz (optional)
// ------------------------------------------------------

function passwordMiddleware(req, res, next) {
  if (!TOOL_PASSWORD) {
    // Kein Passwort konfiguriert -> kein Schutz aktiv
    return next();
  }

  // Passwort kann z.B. im Header oder Body mitgeschickt werden
  const provided =
    req.headers['x-tool-password'] ||
    (req.body && req.body.password) ||
    (req.query && req.query.password);

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
// Hilfsfunktion: Excel-Daten validieren & Payload bauen
// ------------------------------------------------------
//
// HIER musst du deine bestehende Excel-Verarbeitung integrieren.
// Aktuell ist das nur ein Platzhalter mit Struktur.
// excelData kann z.B. ein Base64-String oder bereits geparstes JSON sein.
// ------------------------------------------------------

async function parseExcelAndBuildQuotationPayload(excelData, options) {
  const { allowPriceOverride, mode } = options || {};

  // TODO: Deine echte Excel-Logik einbauen:
  // - Template-Struktur prüfen (Sheets: Angebot / Kunde / Positionen)
  // - Pflichtfelder prüfen (taxType, Kunde.name, type, qty, etc.)
  // - Lexware/Artikel lookup für articleId
  // - Name auto ergänzen, wenn leer
  // - Preise ggf. aus Artikeln ziehen (Standard: immer aus Stammdaten)
  // - allowPriceOverride beachten:
  //   * false: Excel-Preis ignorieren, immer Artikelpreis verwenden
  //   * true: Excel-Preis darf Stammpreis überschreiben (bewusst!)

  // Für jetzt: Beispiel-Struktur zurückgeben
  const fakeQuotationPayload = {
    // HIER ECHTEN Lexware/Lexoffice-Payload einfügen (z.B. Quotation)
    voucherDate: new Date().toISOString().substring(0, 10),
    address: {
      name: 'Demo Kunde GmbH'
    },
    lineItems: [
      // Beispiel-Positionen
      // {
      //   type: 'PRODUCT',
      //   name: 'Beispielartikel',
      //   quantity: 10,
      //   // Preis je nach API-Schema
      // }
    ],
    taxType: 'gross', // Beispiel
    finalized: true // direkt finalisieren (falls API das unterstützt)
  };

  const summary = {
    byType: {
      material: 1,
      service: 0
    },
    autoNamedLineItems: [
      // { position: 1, row: 5, articleName: 'Beispielartikel' }
    ],
    warnings: [],
    errors: [] // falls im Testmodus Fehler gefunden werden
  };

  // Beispiel: im Testmodus könnten hier valide Fehler reingeschrieben werden
  if (mode === 'test') {
    // summary.errors.push({ sheet: 'Kunde', row: 2, field: 'name', message: 'Pflichtfeld fehlt' });
  }

  return {
    quotationPayload: fakeQuotationPayload,
    summary
  };
}

// ------------------------------------------------------
// Hilfsfunktion: Lexware/Lexoffice Angebot erstellen
// ------------------------------------------------------
//
// ACHTUNG: Nur EIN API-Call hier drin!
// Keine Schleifen, keine zusätzlichen Retries – das machen wir oben.
// ------------------------------------------------------

async function createLexofficeQuotationInApi(quotationPayload) {
  if (!LEXOFFICE_API_KEY) {
    const err = new Error('LEXOFFICE_API_KEY ist nicht gesetzt');
    err.status = 500;
    throw err;
  }

  const url = 'https://api.lexoffice.io/v1/quotations'; // ggf. Lexware-URL anpassen

  const res = await callLexwareApi({
    method: 'POST',
    url,
    data: quotationPayload,
    headers: {
      Authorization: `Bearer ${LEXOFFICE_API_KEY}`,
      'Content-Type': 'application/json',
      Accept: 'application/json'
    },
    validateStatus: () => true // wir behandeln Fehler selbst
  });

  // Rate-Limit
  if (res.status === 429) {
    const err = new Error('Rate limit exceeded');
    err.status = 429;
    err.raw = res.data;
    throw err;
  }

  // Sonstige Fehler
  if (res.status < 200 || res.status >= 300) {
    const err = new Error('Lexware/Lexoffice Error');
    err.status = res.status;
    err.raw = res.data;
    throw err;
  }

  return res.data; // sollte u.a. id enthalten
}

// ------------------------------------------------------
// Endpoint: System-Check / Ping
// ------------------------------------------------------

app.get('/api/ping', (req, res) => {
  return res.json({
    ok: true,
    status: 'OK',
    message: 'Lexoffice/Lexware Angebots-Tool Backend läuft.',
    allowPriceOverrideDefault: ALLOW_PRICE_OVERRIDE_DEFAULT,
    passwordProtected: !!TOOL_PASSWORD,
    minIntervalMs: MIN_INTERVAL_MS
  });
});

// ------------------------------------------------------
// Endpoint: Testmodus (Excel nur prüfen, nicht in Lexware/Lexoffice schreiben)
// ------------------------------------------------------

app.post('/api/test-excel', passwordMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride: allowPriceOverrideClient } = req.body || {};

    if (!excelData) {
      return res.status(200).json({
        ok: false,
        status: 'VALIDATION_ERROR',
        stage: 'input',
        message: 'Es wurden keine Excel-Daten übermittelt.',
        httpStatus: 400
      });
    }

    const allowPriceOverride =
      typeof allowPriceOverrideClient === 'boolean'
        ? allowPriceOverrideClient
        : ALLOW_PRICE_OVERRIDE_DEFAULT;

    const { quotationPayload, summary } = await parseExcelAndBuildQuotationPayload(excelData, {
      allowPriceOverride,
      mode: 'test'
    });

    const hasErrors = Array.isArray(summary.errors) && summary.errors.length > 0;

    return res.status(200).json({
      ok: !hasErrors,
      status: hasErrors ? 'VALIDATION_ERROR' : 'SUCCESS',
      stage: 'test',
      message: hasErrors
        ? 'Die Excel-Datei enthält Fehler. Bitte prüfen und korrigieren.'
        : 'Test erfolgreich. Angebot kann erstellt werden.',
      httpStatus: 200,
      data: {
        summary,
        allowPriceOverride
        // optional zur Kontrolle:
        // quotationPayload
      }
    });
  } catch (err) {
    console.error('[test-excel] Fehler:', err);
    return res.status(200).json({
      ok: false,
      status: 'ERROR',
      stage: 'test',
      message: 'Unerwarteter Fehler beim Testmodus.',
      httpStatus: err.status || 500,
      technical: {
        error: err.message || String(err)
      }
    });
  }
});

// ------------------------------------------------------
// Endpoint: Angebot erstellen & finalisieren
// ------------------------------------------------------
//
// WICHTIG:
// - Nur EIN Aufruf von createLexofficeQuotationInApi()
// - Rate-Limit (429) wird sauber als RATE_LIMIT zurückgegeben
// - Kein automatisches Retry bei POST (kein Risiko für Duplikate)
// ------------------------------------------------------

app.post('/api/create-offer', passwordMiddleware, async (req, res) => {
  try {
    const { excelData, allowPriceOverride: allowPriceOverrideClient } = req.body || {};

    if (!excelData) {
      return res.status(200).json({
        ok: false,
        status: 'VALIDATION_ERROR',
        stage: 'input',
        message: 'Es wurden keine Excel-Daten übermittelt.',
        httpStatus: 400
      });
    }

    const allowPriceOverride =
      typeof allowPriceOverrideClient === 'boolean'
        ? allowPriceOverrideClient
        : ALLOW_PRICE_OVERRIDE_DEFAULT;

    // 1) Excel prüfen & Lexware/Lexoffice-Payload bauen
    const { quotationPayload, summary } = await parseExcelAndBuildQuotationPayload(excelData, {
      allowPriceOverride,
      mode: 'create'
    });

    const hasErrors = Array.isArray(summary.errors) && summary.errors.length > 0;
    if (hasErrors) {
      return res.status(200).json({
        ok: false,
        status: 'VALIDATION_ERROR',
        stage: 'validation',
        message: 'Die Excel-Datei enthält Fehler. Angebot wurde nicht erstellt.',
        httpStatus: 200,
        data: {
          summary,
          allowPriceOverride
        }
      });
    }

    // 2) EINMALIGER Lexware/Lexoffice-Erstellungs-Call
    const lexRes = await createLexofficeQuotationInApi(quotationPayload);

    const quotationId = lexRes.id || null;

    return res.status(200).json({
      ok: true,
      status: 'SUCCESS',
      stage: 'lexoffice-create',
      message: 'Angebot erfolgreich in Lexware/Lexoffice erstellt.',
      httpStatus: 200,
      data: {
        quotationId,
        lexofficeResponse: lexRes,
        summary,
        allowPriceOverride
      }
    });
  } catch (err) {
    console.error('[create-offer] Fehler:', err);

    const status = err.status || err.response?.status || 500;
    const raw = err.raw || err.response?.data || null;
    const msg = err.message || 'Fehler beim Erstellen des Angebots.';

    // Spezieller Fall: Rate-Limit 429
    if (status === 429 || msg === 'Rate limit exceeded') {
      return res.status(200).json({
        ok: false,
        status: 'RATE_LIMIT',
        stage: 'lexoffice-create',
        message: 'Lexware/Lexoffice-Rate-Limit überschritten.',
        userMessage:
          'Lexware/Lexoffice meldet aktuell ein Rate-Limit. Das Angebot konnte nicht erstellt bzw. das PDF nicht geladen werden. ' +
          'Bitte später erneut versuchen. Deine Excel-Daten sind grundsätzlich in Ordnung.',
        httpStatus: 429,
        technical: {
          raw
        }
      });
    }

    // Allgemeiner Fehler
    return res.status(200).json({
      ok: false,
      status: 'ERROR',
      stage: 'lexoffice-create',
      message: 'Fehler beim Erstellen des Angebots in Lexware/Lexoffice.',
      httpStatus: status,
      technical: {
        raw
      }
    });
  }
});

// ------------------------------------------------------
// Start Server
// ------------------------------------------------------

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Lexware/Lexoffice Angebots-Tool Backend läuft auf Port ${PORT}`);
});
