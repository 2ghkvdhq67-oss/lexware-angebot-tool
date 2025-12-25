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
    // Ohne API-Key können wir keine Artikel holen
    return [];
  }

  const articles = [];
  let page = 0;
  const pageSize = 100;

  // Paging-Schleife mit Sicherheitslimit
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
    if (page > 50) break; // hartes Sicherheitslimit (~5.000 Artikel)
  }

  return articles;
}

// --------- Template Download (mit Artikel-Lookup + ausführlicher Anleitung) ---------
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

  // Versuchen, Artikel aus Lexoffice zu holen – bei Fehler Template trotzdem liefern
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
    // Fallback: trotzdem Template ausliefern
  }

  // NEU: ausführlicher Anleitung-Tab
  try {
    const helpSheetName = "Anleitung";

    const helpData = [
      ["Bereich", "Feld", "Erklärung"],

      // ==== TYPE: ausführlich, je Typ eine Zeile ====

      [
        "Positionen",
        "type = custom",
        "Individuelle Position mit eigenem Text, Preis und Menge. "
        + "Kein Artikel aus Lexoffice. articleId bleibt leer. "
        + "Pflicht: quantity / qty > 0 und unitPriceAmount / price >= 0. "
        + "Beispiele: Sonderdruck, Einmalkosten, Layout-Pauschale."
      ],

      [
        "Positionen",
        "type = service",
        "Dienstleistung oder Arbeitszeit mit Preis und Menge. "
        + "Optional kann eine articleId hinterlegt werden (falls die Dienstleistung als Artikel in Lexoffice existiert). "
        + "Pflicht: quantity / qty > 0 und unitPriceAmount / price >= 0. "
        + "Beispiele: Gestaltung, Einrichtung, Montage, Arbeitsstunden."
      ],

      [
        "Positionen",
        "type = material",
        "Ware / Artikel aus Lexoffice. articleId ist Pflicht und muss zu einem Eintrag im Tab 'Artikel-Lookup' passen. "
        + "Menge und Preis kommen aus der Excel (oder orientieren sich am Artikel). "
        + "Pflicht: articleId ausgefüllt, quantity / qty > 0, unitPriceAmount / price >= 0. "
        + "Beispiele: T-Shirts, Textilien, Zubehör, Standardartikel."
      ],

      [
        "Positionen",
        "type = text",
        "Reine Infozeile ohne Preis und ohne Menge. "
        + "Wird nur als Text im Angebot angezeigt (z.B. Hinweise, Lieferzeit, Trennzeilen). "
        + "In der Regel keine quantity / qty und kein unitPriceAmount / price setzen."
      ],

      // ==== weitere Felder in Positionen ====

      [
        "Positionen",
        "articleId",
        "Nur verwenden, wenn type = material (oder optional service). "
        + "Die articleId ist die interne Lexoffice-Artikel-ID. "
        + "Sie kann aus dem Tab 'Artikel-Lookup' kopiert werden. "
        + "Wenn type = material und articleId fehlt, wird das Angebot nicht akzeptiert."
      ],

      [
        "Positionen",
        "quantity / qty",
        "Menge der Position. Muss größer als 0 sein. "
        + "Zulässig sind ganze Zahlen oder Dezimalzahlen (z.B. 1,5 für Arbeitsstunden). "
        + "Bei type = text wird quantity normalerweise leer gelassen."
      ],

      [
        "Positionen",
        "unitPriceAmount / price",
        "Einzelpreis pro Stück / Einheit. Muss 0 oder größer sein. "
        + "Dezimaltrennzeichen: Komma oder Punkt sind erlaubt (z.B. 6,9 oder 6.9). "
        + "Bei type = text bleibt dieses Feld normalerweise leer. "
        + "Ob es sich um Netto- oder Bruttopreis handelt, wird über Angebot.taxType gesteuert."
      ],

      [
        "Positionen",
        "name und description",
        "name = Kurzbezeichnung, die in der Angebotszeile angezeigt wird. "
        + "description = optionaler längerer Text (z.B. Details, Zusatzinfos). "
        + "Beides ist besonders wichtig bei custom, service und text, damit das Angebot verständlich ist."
      ],

      // ==== Angebot & Kunde ====

      [
        "Angebot",
        "taxType",
        "Pflichtfeld. Steuert, ob Preise als Netto oder Brutto interpretiert werden. "
        + "Übliche Werte: 'net' für Nettopreise (zzgl. MwSt.) oder 'gross' für Bruttopreise (inkl. MwSt.). "
        + "Muss zur Preislogik in den Positionen passen."
      ],

      [
        "Angebot",
        "currency",
        "Währung des Angebots, Standard ist EUR. Nur ändern, wenn in Lexoffice passende Währungskonten eingerichtet sind."
      ],

      [
        "Kunde",
        "name",
        "Pflichtfeld. Name des Kunden (Firma oder Privatperson), an den das Angebot gesendet wird. "
        + "Kann entweder zu einem bestehenden Kontakt in Lexoffice passen oder für einen neuen Kontakt verwendet werden."
      ],

      [
        "Kunde",
        "contactId",
        "Optional. Wenn vorhanden, verweist dieser Wert direkt auf einen bestehenden Kontakt in Lexoffice. "
        + "Dann werden Adresse und Firmendaten aus Lexoffice übernommen. "
        + "Wenn contactId leer ist, wird anhand von name und weiteren Feldern (Straße, PLZ, Ort) gearbeitet."
      ],

      // ==== Allgemeine Hinweise ====

      [
        "Allgemein",
        "Struktur",
        "Bitte die Sheet-Namen ('Angebot', 'Kunde', 'Positionen', 'Artikel-Lookup', 'Anleitung') "
        + "sowie die Spaltenüberschriften nicht ändern. "
        + "Neue Zeilen für weitere Positionen sind erlaubt."
      ],

      [
        "Allgemein",
        "Prüfung",
        "Empfehlung: Zuerst immer 'Nur prüfen' im Tool verwenden. "
        + "Wenn keine Fehlermeldung kommt, anschließend 'Angebot erstellen'. "
        + "Fehlermeldungen nennen immer Tabelle, Zeile und Feld, die korrigiert werden müssen."
      ]
    ];

    const wsHelp = xlsx.utils.aoa_to_sheet(helpData);
    wb.Sheets[helpSheetName] = wsHelp;
    if (!wb.SheetNames.includes(helpSheetName)) {
      wb.SheetNames.push(helpSheetName);
    }
  } catch (e) {
    console.error("Fehler beim Erzeugen der Anleitung-Tabelle:", e);
    // Template trotzdem ausliefern
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

    if (!r.type || !r.name) {
      return {
        ok: false,
        statusCode: 422,
        body: validationError("Typ oder Positionsname fehlt.", "Positionen", row, "type/name")
      };
    }

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

    if (String(r.type).toLowerCase() === "material" && !r.articleId) {
      return {
        ok: false,
        statusCode: 422,
        body: validationError(
          "articleId ist für Material erforderlich.",
          "Positionen",
          row,
          "articleId"
        )
      };
    }

    byType[r.type] = (byType[r.type] || 0) + 1;
    r.qty = qty;
    r.price = price;
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
    const type = String(r.type).toLowerCase();
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
      name: r.name,
      quantity: qty,
      unitName: r.unitName || "Stk",
      unitPrice
    };

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
