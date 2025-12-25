import "dotenv/config";
import express from "express";
import path from "path";
import fs from "fs";
import multer from "multer";
import xlsx from "xlsx";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static(path.resolve("./public")));

// Health
app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

// Template Download
app.get("/download-template-with-articles", (_req, res) => {
  const p = path.resolve("./templates/Lexware_Template.xlsx");
  if (!fs.existsSync(p)) {
    return res.status(500).json({
      ok: false,
      message: "Template-Datei fehlt auf dem Server"
    });
  }
  res.download(p, "Lexware_Template.xlsx");
});

// API Test
app.get("/api-test", async (_req, res) => {
  try {
    const r = await fetch("https://api.lexware.io/v1/profile", {
      headers: {
        Authorization: `Bearer ${process.env.LEXWARE_API_KEY}`,
        Accept: "application/json"
      }
    });

    if (!r.ok) {
      return res.status(500).json({
        ok: false,
        message: "API-Verbindung fehlgeschlagen",
        status: r.status
      });
    }

    const d = await r.json();
    res.json({
      ok: true,
      organization: d.organizationName || "API-Verbindung erfolgreich"
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      message: "API-Test fehlgeschlagen",
      error: String(e)
    });
  }
});

// ------------------------------------------------------------------
// Helper für klare Fehlermeldungen
// ------------------------------------------------------------------
function validationError(message, sheet, row, field) {
  return {
    ok: false,
    message: `${message} (Tabelle: ${sheet}, Zeile ${row})`,
    details: { sheet, row, field }
  };
}

// ------------------------------------------------------------------
// Helper: Excel einlesen & validieren (gemeinsam für Test & Live)
// ------------------------------------------------------------------
function parseAndValidateExcel(buffer) {
  let wb;
  try {
    wb = xlsx.read(buffer, { type: "buffer" });
  } catch {
    return {
      ok: false,
      status: 400,
      error: {
        ok: false,
        message: "Excel-Datei kann nicht gelesen werden"
      }
    };
  }

  // Pflicht-Sheets
  for (const s of ["Angebot", "Kunde", "Positionen"]) {
    if (!wb.Sheets[s]) {
      return {
        ok: false,
        status: 422,
        error: {
          ok: false,
          message: `Tabelle fehlt: ${s}`,
          details: { sheet: s }
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
      status: 422,
      error: {
        ok: false,
        message: "Feld fehlt: Angebot.taxType",
        details: { sheet: "Angebot", field: "taxType" }
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
      status: 422,
      error: {
        ok: false,
        message: "Feld fehlt: Kunde.name",
        details: { sheet: "Kunde", field: "name" }
      }
    };
  }

  // Positionen
  const pos = xlsx.utils.sheet_to_json(wb.Sheets["Positionen"], {
    defval: ""
  });

  if (!pos.length) {
    return {
      ok: false,
      status: 422,
      error: {
        ok: false,
        message: "Keine Positionen vorhanden",
        details: { sheet: "Positionen" }
      }
    };
  }

  const byType = {};

  for (let i = 0; i < pos.length; i++) {
    const r = pos[i];
    const row = i + 2; // Header + Excel-Index

    if (!r.type || !r.name) {
      return {
        ok: false,
        status: 422,
        error: validationError(
          "Typ oder Positionsname fehlt",
          "Positionen",
          row,
          "type/name"
        )
      };
    }

    const qty = Number(r.qty);
    if (!Number.isFinite(qty) || qty <= 0) {
      return {
        ok: false,
        status: 422,
        error: validationError(
          "Menge muss größer als 0 sein",
          "Positionen",
          row,
          "qty"
        )
      };
    }

    const price = Number(r.price);
    if (!Number.isFinite(price) || price < 0) {
      return {
        ok: false,
        status: 422,
        error: validationError(
          "Preis muss 0 oder größer sein",
          "Positionen",
          row,
          "price"
        )
      };
    }

    if (String(r.type).toLowerCase() === "material" && !r.articleId) {
      return {
        ok: false,
        status: 422,
        error: validationError(
          "articleId ist für Material erforderlich",
          "Positionen",
          row,
          "articleId"
        )
      };
    }

    byType[r.type] = (byType[r.type] || 0) + 1;
  }

  return {
    ok: true,
    angebot,
    kunde,
    pos,
    byType
  };
}

// ------------------------------------------------------------------
// Helper: Angebots-Payload für Lexware bauen
// ------------------------------------------------------------------
function buildQuotationPayload(angebot, kunde, pos) {
  const now = new Date();

  const voucherDate = angebot.voucherDate
    ? new Date(angebot.voucherDate)
    : now;

  const expirationDate = angebot.expirationDate
    ? new Date(angebot.expirationDate)
    : new Date(voucherDate.getTime() + 14 * 24 * 60 * 60 * 1000);

  const shippingDate = angebot.shippingDate
    ? new Date(angebot.shippingDate)
    : voucherDate;

  const currency = angebot.currency || "EUR";
  const taxType = angebot.taxType;
  const defaultTaxRate =
    angebot.taxRateDefault !== undefined && angebot.taxRateDefault !== ""
      ? Number(angebot.taxRateDefault)
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
    const qty = Number(r.qty);
    const price = Number(r.price);
    const taxRate =
      r.taxRate !== undefined && r.taxRate !== ""
        ? Number(r.taxRate)
        : defaultTaxRate;

    const unitPrice = {
      currency,
      taxRatePercentage: taxRate
    };

    if (taxType === "gross") {
      unitPrice.grossAmount = price;
    } else {
      unitPrice.netAmount = price;
    }

    const item = {
      type, // z.B. custom | material | service | text
      name: r.name,
      quantity: qty,
      unitName: r.unitName || "Stück",
      unitPrice
    };

    if (r.description) {
      item.description = r.description;
    }
    if ((type === "material" || type === "service") && r.articleId) {
      item.id = r.articleId;
    }

    return item;
  });

  return {
    voucherDate: voucherDate.toISOString(),
    expirationDate: expirationDate.toISOString(),
    address,
    lineItems,
    totalPrice: {
      currency
    },
    taxConditions: {
      taxType
    },
    shippingConditions: {
      shippingType: angebot.shippingType || "service",
      shippingDate: shippingDate.toISOString()
    },
    ...(angebot.title ? { title: angebot.title } : {}),
    ...(angebot.introduction ? { introduction: angebot.introduction } : {}),
    ...(angebot.remark ? { remark: angebot.remark } : {})
  };
}

// ------------------------------------------------------------------
// TESTMODUS — Excel prüfen
// ------------------------------------------------------------------
app.post("/validate-excel", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      message: "Keine Datei hochgeladen"
    });
  }

  const result = parseAndValidateExcel(req.file.buffer);
  if (!result.ok) {
    return res.status(result.status).json(result.error);
  }

  const { kunde, angebot, pos, byType } = result;

  return res.json({
    ok: true,
    summary: {
      customer: kunde.name,
      taxType: angebot.taxType,
      positions: pos.length,
      byType
    }
  });
});

// ------------------------------------------------------------------
// LIVE — Angebot erstellen & PDF zurückgeben
// ------------------------------------------------------------------
app.post("/create-quote-from-excel", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      message: "Keine Datei hochgeladen"
    });
  }

  const parsed = parseAndValidateExcel(req.file.buffer);
  if (!parsed.ok) {
    return res.status(parsed.status).json(parsed.error);
  }

  if (!process.env.LEXWARE_API_KEY) {
    return res.status(500).json({
      ok: false,
      message: "LEXWARE_API_KEY ist nicht gesetzt"
    });
  }

  const { angebot, kunde, pos, byType } = parsed;
  const quotationPayload = buildQuotationPayload(angebot, kunde, pos);
  const baseUrl = "https://api.lexware.io";

  try {
    // 1) Angebot erstellen
    const createRes = await fetch(`${baseUrl}/v1/quotations?finalize=true`, {
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
        // ignorieren
      }
      return res.status(502).json({
        ok: false,
        message: "Fehler beim Erstellen des Angebots in Lexware",
        status: createRes.status,
        error: errorBody
      });
    }

    const quotation = await createRes.json();
    const quotationId = quotation.id;

    // 2) PDF laden
    const fileRes = await fetch(`${baseUrl}/v1/quotations/${quotationId}/file`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${process.env.LEXWARE_API_KEY}`,
        Accept: "application/pdf"
      }
    });

    if (!fileRes.ok) {
      let errorText = null;
      try {
        errorText = await fileRes.text();
      } catch {
        // ignorieren
      }
      return res.status(502).json({
        ok: false,
        message: "Angebot erstellt, aber PDF konnte nicht geladen werden",
        status: fileRes.status,
        error: errorText
      });
    }

    const arrayBuffer = await fileRes.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const contentType =
      fileRes.headers.get("content-type") || "application/pdf";
    const contentDispositionHeader = fileRes.headers.get(
      "content-disposition"
    );
    const fallbackFilename = `Angebot-${quotationId}.pdf`;

    res.setHeader("Content-Type", contentType);
    if (contentDispositionHeader && contentDispositionHeader.includes("filename=")) {
      res.setHeader("Content-Disposition", contentDispositionHeader);
    } else {
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${fallbackFilename}"`
      );
    }

    // einfache Summary als Header
    res.setHeader(
      "X-Lexware-Quote-Summary",
      encodeURIComponent(
        JSON.stringify({
          customer: kunde.name,
          taxType: angebot.taxType,
          positions: pos.length,
          byType
        })
      )
    );

    return res.send(buffer);
  } catch (e) {
    console.error("Fehler Lexware-API:", e);
    return res.status(500).json({
      ok: false,
      message: "Unerwarteter Fehler bei der Lexware-API",
      error: String(e)
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server läuft");
});
