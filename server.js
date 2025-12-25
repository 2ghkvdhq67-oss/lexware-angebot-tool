import "dotenv/config";
import express from "express";
import path from "path";
import fs from "fs";
import multer from "multer";
import xlsx from "xlsx";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

const BASE_URL = "https://api.lexware.io";

app.use(express.static(path.resolve("./public")));

// Health
app.get("/health", (_req, res) => res.json({ ok: true }));

// --------- Helper: Zahlen & Fehler ---------
function validationError(message, sheet, row, field) {
  return {
    ok: false,
    status: "ERROR",
    stage: "validation",
    message,
    sheet,
    row,
    field
  };
}

// Zahl aus Excel robust parsen (inkl. "6,9")
function parseNumberValue(value) {
  if (value === undefined || value === null || value === "") return NaN;
  if (typeof value === "string") {
    value = value.replace(",", ".").trim();
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

// Einfache OK-Antwort
function successResponse(action, extra = {}) {
  return {
    ok: true,
    status: "OK",
    action,
    ...extra
  };
}

// --------- Helper: alle Artikel holen (Paging) ---------
async function fetchAllArticles() {
  if (!process.env.LEXWARE_API_KEY) {
    return [];
  }

  const articles = [];
  let page = 0;
  const pageSize = 100;

  while (true) {
    const url = `${BASE_URL}/v1/articles?page=${page}&size=${pageSize}`;
    const res = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${process.env.LEXWARE_API_KEY}`,
        Accept: "application/json"
      }
    });

    if (!res.ok) {
      throw new Error(`Artikel konnten nicht geladen werden (HTTP ${res.status})`);
    }

    const data = await res.json();
    if (Array.isArray(data.content)) {
      articles.push(...data.content);
    }

    const isLast = data.last === true;
    const totalPages = typeof data.totalPages === "number" ? data.totalPages : null;

    if (isLast) break;
    if (totalPages !== null && page >= totalPages - 1) break;

    page += 1;
    if (page > 50) break; // Sicherheitslimit
  }

  return articles;
}

// --------- Template Download (mit Artikel-Lookup + Anleitung) ---------
app.get("/download-template-with-articles", async (_req, res) => {
  const p = path.resolve("./templates/Lexware_Template.xlsx");
  if (!fs.existsSync(p)) {
    return res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "server",
      message: "Template fehlt auf dem Server"
    });
  }

  // Basis-Template laden
  let wb;
  try {
    const fileBuf = fs.readFileSync(p);
    wb = xlsx.read(fileBuf, { type: "buffer" });
  } catch (e) {
    console.error("Fehler beim Lesen des Template-Excels:", e);
    return res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "server",
      message: "Template-Datei kann nicht gelesen werden"
    });
  }

  // Artikel-Lookup füllen
  try {
    const articles = await fetchAllArticles();

    if (articles.length > 0) {
      const sheetName = "Artikel-Lookup";
      const header = [
        "id",
        "title",
        "articleNumber",
        "unitName",
        "netPrice",
        "grossPrice",
        "taxRate"
      ];

      const rows = articles.map((a) => [
        a.id || "",
        a.title || "",
        a.articleNumber || "",
        a.unitName || "",
        a.price && typeof a.price.netPrice === "number" ? a.price.netPrice : "",
        a.price && typeof a.price.grossPrice === "number" ? a.price.grossPrice : "",
        a.price && typeof a.price.taxRate === "number" ? a.price.taxRate : ""
      ]);

      const data = [header, ...rows];
      const ws = xlsx.utils.aoa_to_sheet(data);

      wb.Sheets[sheetName] = ws;
      if (!wb.SheetNames.includes(sheetName)) {
        wb.SheetNames.push(sheetName);
      }
    } else {
      console.log("Keine Artikel aus Lexoffice erhalten – Artikel-Lookup bleibt leer.");
    }
  } catch (e) {
    console.error("Fehler beim Laden der Artikel für das Template:", e);
  }

  // Anleitung-Tab (ausführlich)
  try {
    const helpSheetName = "Anleitung";

    const helpData = [
      ["Bereich", "Feld", "Erklärung"],

      [
        "Positionen",
        "type = custom",
        "Individuelle Position mit eigenem Text, Preis und Menge. Kein Artikel aus Lexoffice. "
        + "Pflicht: quantity / qty > 0 und unitPriceAmount / price >= 0. Beispiel: Sonderdruck, Einmalkosten."
      ],
      [
        "Positionen",
        "type = service",
        "Dienstleistung oder Arbeitszeit mit Preis und Menge. Optional articleId, wenn die Leistung in Lexoffice als Artikel existiert. "
        + "Pflicht: quantity / qty > 0 und unitPriceAmount / price >= 0. Beispiel: Gestaltung, Montage."
      ],
      [
        "Positionen",
        "type = material",
        "Ware / Artikel aus Lexoffice. articleId ist Pflicht (aus Tab 'Artikel-Lookup'). "
        + "name kann leer bleiben – dann kommt der Name aus Lexoffice. Pflicht: articleId, quantity / qty > 0, unitPriceAmount / price >= 0."
      ],
      [
        "Positionen",
        "type = text",
        "Reine Infozeile ohne Preis und ohne Menge. Wird nur als Text im Angebot angezeigt (z.B. Lieferzeit-Hinweise). "
        + "In der Regel quantity / qty und unitPriceAmount / price leer lassen."
      ],

      [
        "Positionen",
        "articleId",
        "Nur bei type = material (oder optional service) verwenden. ID aus 'Artikel-Lookup' kopieren. "
        + "Wenn type = material und articleId fehlt, wird die Datei abgelehnt."
      ],
      [
        "Positionen",
        "quantity / qty",
        "Menge der Position. Muss größer als 0 sein. Ganze Zahlen oder Dezimalzahlen (z.B. 1,5 Stunden) sind erlaubt."
      ],
      [
        "Positionen",
        "unitPriceAmount / price",
        "Einzelpreis pro Stück/Einheit. Muss 0 oder größer sein. Komma oder Punkt erlaubt (z.B. 6,9 oder 6.9). "
        + "Ob Netto oder Brutto hängt von Angebot.taxType ab."
      ],
      [
        "Positionen",
        "name & description",
        "name = kurze Bezeichnung der Position. description = optionaler längerer Beschreibungstext. "
        + "Bei material darf name leer sein (dann kommt der Name aus Lexoffice). Bei custom/service/text sollte name ausgefüllt werden."
      ],

      [
        "Angebot",
        "taxType",
        "Pflichtfeld. Steuert, ob Preise als Netto oder Brutto interpretiert werden. Üblich: 'net' (Netto) oder 'gross' (Brutto)."
      ],
      [
        "Angebot",
        "currency",
        "Währung des Angebots. Standard: EUR."
      ],
      [
        "Kunde",
        "name",
        "Pflichtfeld. Name des Kunden (Firma oder Person), an den das Angebot geht."
      ],
      [
        "Kunde",
        "contactId",
        "Optional. Wenn ausgefüllt, wird direkt dieser Kontakt aus Lexoffice verwendet."
      ],
      [
        "Allgemein",
        "Struktur",
        "Bitte Sheet-Namen und Spaltenüberschriften nicht ändern. Neue Zeilen für weitere Positionen sind erlaubt."
      ],
      [
        "Allgemein",
        "Tipp",
        "Im Tool zuerst 'Nur prüfen' verwenden. Wenn keine Fehlermeldung kommt, anschließend 'Angebot erstellen'. "
        + "Fehlermeldungen nennen immer Tabelle, Zeile und Feld."
      ]
    ];

    const wsHelp = xlsx.utils.aoa_to_sheet(helpData);
    wb.Sheets[helpSheetName] = wsHelp;
    if (!wb.SheetNames.includes(helpSheetName)) {
      wb.SheetNames.push(helpSheetName);
    }
  } catch (e) {
    console.error("Fehler beim Erzeugen der Anleitung-Tabelle:", e);
  }

  // Workbook zurück an den Browser
  try {
    const outBuf = xlsx.write(wb, { bookType: "xlsx", type: "buffer" });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="Lexoffice_Template_mit_Artikeln.xlsx"'
    );
    return res.send(outBuf);
  } catch (e) {
    console.error("Fehler beim Schreiben des Template-Excels:", e);
    return res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "server",
      message: "Template konnte nicht erzeugt werden"
    });
  }
});

// --------- API Test ---------
app.get("/api-test", async (_req, res) => {
  try {
    const r = await fetch(`${BASE_URL}/v1/profile`, {
      headers: {
        Authorization: `Bearer ${process.env.LEXWARE_API_KEY}`,
        Accept: "application/json"
      }
    });
    if (!r.ok) {
      return res.status(500).json({
        ok: false,
        status: "ERROR",
        stage: "api-test",
        message: "API-Verbindung fehlgeschlagen",
        httpStatus: r.status
      });
    }
    const d = await r.json();
    res.json({
      ok: true,
      status: "OK",
      stage: "api-test",
      organization: d.organizationName || "OK"
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "api-test",
      message: "API-Test fehlgeschlagen",
      technical: String(e)
    });
  }
});

// --------- Gemeinsame Funktion: Excel einlesen & validieren ---------
function parseAndValidateExcel(buffer) {
  let wb;
  try {
    wb = xlsx.read(buffer, { type: "buffer" });
  } catch {
    return {
      ok: false,
      statusCode: 400,
      body: {
        ok: false,
        status: "ERROR",
        stage: "validation",
        message: "Excel-Datei kann nicht gelesen werden"
      }
    };
  }

  // Pflicht-Sheets
  for (const s of ["Angebot", "Kunde", "Positionen"]) {
    if (!wb.Sheets[s]) {
      return {
        ok: false,
        statusCode: 422,
        body: {
          ok: false,
          status: "ERROR",
          stage: "validation",
          message: `Tabelle fehlt: ${s}`,
          sheet: s
        }
      };
    }
  }

  // Angebot
  const angebot = Object.fromEntries(
    xlsx.utils
      .sheet_to_json(wb.Sheets["Angebot"], { header: 1, defval: "" })
      .slice(1)
      .filter((r) => r[0])
      .map((r) => [r[0], r[1]])
  );
  if (!angebot.taxType) {
    return {
      ok: false,
      statusCode: 422,
      body: {
        ok: false,
        status: "ERROR",
        stage: "validation",
        message: "Feld fehlt: Angebot.taxType",
        sheet: "Angebot",
        field: "taxType"
      }
    };
  }

  // Kunde
  const kunde = Object.fromEntries(
    xlsx.utils
      .sheet_to_json(wb.Sheets["Kunde"], { header: 1, defval: "" })
      .slice(1)
      .filter((r) => r[0])
      .map((r) => [r[0], r[1]])
  );
  if (!kunde.name) {
    return {
      ok: false,
      statusCode: 422,
      body: {
        ok: false,
        status: "ERROR",
        stage: "validation",
        message: "Feld fehlt: Kunde.name",
        sheet: "Kunde",
        field: "name"
      }
    };
  }

  // Positionen
  const pos = xlsx.utils.sheet_to_json(wb.Sheets["Positionen"], { defval: "" });
  if (!pos.length) {
    return {
      ok: false,
      statusCode: 422,
      body: {
        ok: false,
        status: "ERROR",
        stage: "validation",
        message: "Keine Positionen vorhanden",
        sheet: "Positionen"
      }
    };
  }

  const byType = {};
  for (let i = 0; i < pos.length; i++) {
    const r = pos[i];
    const row = i + 2;

    const typeRaw = (r.type || "").toString().trim();
    const nameRaw = (r.name || "").toString().trim();

    // type immer Pflicht
    if (!typeRaw) {
      return {
        ok: false,
        statusCode: 422,
        body: validationError("Typ (type) fehlt.", "Positionen", row, "type")
      };
    }

    const typeLower = typeRaw.toLowerCase();

    // name-Pflicht nur für nicht-material
    if (typeLower !== "material" && typeLower !== "text" && !nameRaw) {
      return {
        ok: false,
        statusCode: 422,
        body: validationError("Positionsname (name) fehlt.", "Positionen", row, "name")
      };
    }

    // text-Zeilen dürfen keinen Preis und keine Menge haben (optional, wir erlauben Menge/Preis aber aktuell)
    // hier NICHT hart prüfen, damit es flexibel bleibt

    // Menge: qty oder quantity
    const rawQty = r.qty ?? r.quantity;
    const qty = parseNumberValue(rawQty);
    if (!Number.isFinite(qty) || qty <= 0) {
      return {
        ok: false,
        statusCode: 422,
        body: validationError(
          "Menge muss größer als 0 sein.",
          "Positionen",
          row,
          r.qty !== undefined ? "qty" : "quantity"
        )
      };
    }

    // Preis: price oder unitPriceAmount
    const rawPrice = r.price ?? r.unitPriceAmount;
    const price = parseNumberValue(rawPrice);
    if (!Number.isFinite(price) || price < 0) {
      return {
        ok: false,
        statusCode: 422,
        body: validationError(
          "Preis muss 0 oder größer sein.",
          "Positionen",
          row,
          r.price !== undefined ? "price" : "unitPriceAmount"
        )
      };
    }

    // articleId-Pflicht nur für material
    if (typeLower === "material" && !r.articleId) {
      return {
        ok: false,
        statusCode: 422,
        body: validationError(
          "articleId ist für type = material erforderlich.",
          "Positionen",
          row,
          "articleId"
        )
      };
    }

    byType[typeRaw] = (byType[typeRaw] || 0) + 1;
    r.qty = qty;
    r.price = price;
    r.type = typeLower;   // normalisiert
    r.name = nameRaw;     // getrimmt (kann leer sein)
  }

  return {
    ok: true,
    angebot,
    kunde,
    pos,
    byType
  };
}

// --------- TESTMODUS: nur prüfen ---------
app.post("/validate-excel", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      status: "ERROR",
      stage: "validation",
      message: "Keine Datei hochgeladen"
    });
  }

  const result = parseAndValidateExcel(req.file.buffer);
  if (!result.ok) {
    return res.status(result.statusCode).json(result.body);
  }

  const { angebot, kunde, pos, byType } = result;

  return res.json(
    successResponse("validate", {
      message: "Excel erfolgreich geprüft",
      customer: kunde.name,
      taxType: angebot.taxType,
      positions: pos.length,
      byType
    })
  );
});

// --------- LIVE: Angebot erzeugen ---------
app.post("/create-quote-from-excel", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      status: "ERROR",
      stage: "validation",
      message: "Keine Datei hochgeladen"
    });
  }

  const result = parseAndValidateExcel(req.file.buffer);
  if (!result.ok) {
    return res.status(result.statusCode).json(result.body);
  }

  if (!process.env.LEXWARE_API_KEY) {
    return res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "config",
      message: "LEXWARE_API_KEY ist nicht gesetzt"
    });
  }

  const { angebot, kunde, pos, byType } = result;

  const now = new Date();
  const voucherDate = angebot.voucherDate ? new Date(angebot.voucherDate) : now;
  const expirationDate = angebot.expirationDate
    ? new Date(angebot.expirationDate)
    : new Date(voucherDate.getTime() + 14 * 24 * 60 * 60 * 1000);
  const shippingDate = angebot.shippingDate ? new Date(angebot.shippingDate) : voucherDate;

  const currency = angebot.currency || "EUR";
  const taxType = angebot.taxType;
  const defaultTaxRate =
    angebot.taxRateDefault !== undefined && angebot.taxRateDefault !== ""
      ? parseNumberValue(angebot.taxRateDefault)
      : 19;

  const address = {};
  if (kunde.contactId) {
    address.contactId = String(kunde.contactId).trim();
  } else {
    address.name = kunde.name;
    address.countryCode = (kunde.countryCode || "DE").toString().trim();
  }

  const lineItems = pos.map((r) => {
    const type = r.type; // schon lower-case
    const qty = r.qty;
    const price = r.price;
    const rawTaxRate = r.taxRate ?? r.taxRatePercentage ?? defaultTaxRate;
    const taxRate = parseNumberValue(rawTaxRate);

    const unitPrice = {
      currency,
      taxRatePercentage: Number.isFinite(taxRate) ? taxRate : defaultTaxRate
    };

    if (taxType === "gross") {
      unitPrice.grossAmount = price;
    } else {
      unitPrice.netAmount = price;
    }

    const item = {
      type,
      quantity: qty,
      unitName: r.unitName || "Stk",
      unitPrice
    };

    // name nur setzen, wenn vorhanden (bei material darf leer sein)
    if (r.name) {
      item.name = r.name;
    }

    if (r.description) item.description = r.description;
    if ((type === "material" || type === "service") && r.articleId) {
      item.id = r.articleId;
    }

    return item;
  });

  const quotationPayload = {
    voucherDate: voucherDate.toISOString(),
    expirationDate: expirationDate.toISOString(),
    address,
    lineItems,
    totalPrice: { currency },
    taxConditions: { taxType },
    shippingConditions: {
      shippingType: angebot.shippingType || "service",
      shippingDate: shippingDate.toISOString()
    },
    ...(angebot.title ? { title: angebot.title } : {}),
    ...(angebot.introduction ? { introduction: angebot.introduction } : {}),
    ...(angebot.remark ? { remark: angebot.remark } : {})
  };

  try {
    const createRes = await fetch(`${BASE_URL}/v1/quotations?finalize=true`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.LEXWARE_API_KEY}`,
        "Content-Type": "application/json",
        Accept: "application/json"
      },
      body: JSON.stringify(quotationPayload)
    });

    if (!createRes.ok) {
      let errorBody = null;
      try {
        errorBody = await createRes.json();
      } catch {
        // ignore
      }
      return res.status(502).json({
        ok: false,
        status: "ERROR",
        stage: "lexoffice-create",
        message: "Fehler beim Erstellen des Angebots in Lexoffice",
        httpStatus: createRes.status,
        technical: errorBody
      });
    }

    const quotation = await createRes.json();
    const quotationId = quotation.id;

    return res.json(
      successResponse("create-quote", {
        message: "Angebot in Lexoffice erstellt",
        quotationId,
        customer: kunde.name,
        taxType: angebot.taxType,
        positions: pos.length,
        byType
      })
    );
  } catch (e) {
    console.error("Fehler Lexoffice-API:", e);
    return res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "lexoffice-create",
      message: "Unerwarteter Fehler bei der Lexoffice-API",
      technical: String(e)
    });
  }
});

// --------- PDF-Download ---------
app.get("/download-quote-pdf", async (req, res) => {
  const quotationId = req.query.id;
  if (!quotationId) {
    return res.status(400).json({
      ok: false,
      status: "ERROR",
      stage: "pdf",
      message: "quotationId fehlt"
    });
  }

  if (!process.env.LEXWARE_API_KEY) {
    return res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "config",
      message: "LEXWARE_API_KEY ist nicht gesetzt"
    });
  }

  try {
    const fileRes = await fetch(
      `${BASE_URL}/v1/quotations/${encodeURIComponent(quotationId)}/file`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${process.env.LEXWARE_API_KEY}`,
          Accept: "application/pdf"
        }
      }
    );

    if (!fileRes.ok) {
      let errorText = null;
      try {
        errorText = await fileRes.text();
      } catch {
        // ignore
      }
      const isRateLimit = fileRes.status === 429;

      return res.status(fileRes.status).json({
        ok: false,
        status: "ERROR",
        stage: "pdf",
        message: isRateLimit
          ? "PDF kann aktuell nicht geladen werden (Lexoffice-Rate-Limit 429). Bitte später erneut versuchen oder das Angebot direkt in Lexoffice öffnen."
          : "PDF konnte nicht geladen werden.",
        httpStatus: fileRes.status,
        technical: errorText
      });
    }

    const arrayBuffer = await fileRes.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const contentType = fileRes.headers.get("content-type") || "application/pdf";
    const contentDispositionHeader = fileRes.headers.get("content-disposition");
    const fallbackFilename = `Angebot-${quotationId}.pdf`;

    res.setHeader("Content-Type", contentType);
    if (contentDispositionHeader && contentDispositionHeader.includes("filename=")) {
      res.setHeader("Content-Disposition", contentDispositionHeader);
    } else {
      res.setHeader("Content-Disposition", `attachment; filename="${fallbackFilename}"`);
    }

    return res.send(buffer);
  } catch (e) {
    console.error("Fehler beim PDF-Download:", e);
    return res.status(500).json({
      ok: false,
      status: "ERROR",
      stage: "pdf",
      message: "Unerwarteter Fehler beim PDF-Download",
      technical: String(e)
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server läuft");
});
