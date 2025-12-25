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
app.get("/health", (_req, res) => res.json({ ok: true }));

// Template Download
app.get("/download-template-with-articles", (_req, res) => {
  const p = path.resolve("./templates/Lexware_Template.xlsx");
  if (!fs.existsSync(p))
    return res.status(500).json({ ok:false, message:"Template fehlt" });
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
    if (!r.ok) return res.status(500).json({ ok:false, status:r.status });
    const d = await r.json();
    res.json({ ok:true, org:d.organizationName || "OK" });
  } catch (e) {
    res.status(500).json({ ok:false, error:String(e) });
  }
});

// helper für lesbare Fehlermeldungen
function validationError(message, sheet, row, field) {
  return {
    ok:false,
    message:`${message} Tabelle: ${sheet}, Zeile ${row}`,
    details:{ sheet, row, field }
  };
}

// helper: Zahl aus Excel parsen, inkl. "6,9"
function parseNumberValue(value) {
  if (value === undefined || value === null || value === "") return NaN;
  if (typeof value === "string") {
    value = value.replace(",", ".").trim();
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

// ------------------ TESTMODUS: nur prüfen ------------------
app.post("/validate-excel", upload.single("file"), (req, res) => {
  if (!req.file)
    return res.status(400).json({ ok:false, message:"Keine Datei hochgeladen" });

  let wb;
  try {
    wb = xlsx.read(req.file.buffer, { type:"buffer" });
  } catch {
    return res.status(400).json({
      ok:false,
      message:"Excel-Datei kann nicht gelesen werden"
    });
  }

  // Pflicht-Sheets
  for (const s of ["Angebot","Kunde","Positionen"]) {
    if (!wb.Sheets[s])
      return res.status(422).json({
        ok:false,
        message:`Tabelle fehlt: ${s}`,
        details:{ sheet:s }
      });
  }

  // Angebot
  const angebot = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Angebot"], { header:1, defval:"" })
      .slice(1).filter(r=>r[0]).map(r=>[r[0], r[1]])
  );
  if (!angebot.taxType)
    return res.status(422).json({
      ok:false,
      message:"Feld fehlt: Angebot.taxType",
      details:{sheet:"Angebot", field:"taxType"}
    });

  // Kunde
  const kunde = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Kunde"], { header:1, defval:"" })
      .slice(1).filter(r=>r[0]).map(r=>[r[0], r[1]])
  );
  if (!kunde.name)
    return res.status(422).json({
      ok:false,
      message:"Feld fehlt: Kunde.name",
      details:{sheet:"Kunde", field:"name"}
    });

  // Positionen
  const pos = xlsx.utils.sheet_to_json(wb.Sheets["Positionen"], { defval:"" });
  if (!pos.length)
    return res.status(422).json({
      ok:false,
      message:"Keine Positionen vorhanden",
      details:{sheet:"Positionen"}
    });

  const byType = {};
  for (let i=0;i<pos.length;i++){
    const r = pos[i];
    const row = i+2;

    if (!r.type || !r.name)
      return res.status(422).json(
        validationError("Typ oder Positionsname fehlt.", "Positionen", row, "type/name")
      );

    // Menge: qty oder quantity
    const rawQty = r.qty ?? r.quantity;
    const qty = parseNumberValue(rawQty);
    if (!Number.isFinite(qty) || qty<=0)
      return res.status(422).json(
        validationError(
          "Menge muss größer als 0 sein.",
          "Positionen",
          row,
          r.qty !== undefined ? "qty" : "quantity"
        )
      );

    // Preis: price oder unitPriceAmount
    const rawPrice = r.price ?? r.unitPriceAmount;
    const price = parseNumberValue(rawPrice);
    if (!Number.isFinite(price) || price<0)
      return res.status(422).json(
        validationError(
          "Preis muss 0 oder größer sein.",
          "Positionen",
          row,
          r.price !== undefined ? "price" : "unitPriceAmount"
        )
      );

    if (String(r.type).toLowerCase()==="material" && !r.articleId)
      return res.status(422).json(
        validationError(
          "articleId ist für Material erforderlich.",
          "Positionen",
          row,
          "articleId"
        )
      );

    byType[r.type] = (byType[r.type] || 0) + 1;
    // normalisierte Werte merken
    r.qty = qty;
    r.price = price;
  }

  res.json({
    ok:true,
    summary:{
      customer:kunde.name,
      taxType:angebot.taxType,
      positions:pos.length,
      byType
    }
  });
});

// ------------------ LIVE: Angebot + PDF ------------------
app.post("/create-quote-from-excel", upload.single("file"), async (req, res) => {
  if (!req.file)
    return res.status(400).json({ ok:false, message:"Keine Datei hochgeladen" });

  // gleiche Logik wie im Testmodus nutzen
  let wb;
  try {
    wb = xlsx.read(req.file.buffer, { type:"buffer" });
  } catch {
    return res.status(400).json({
      ok:false,
      message:"Excel-Datei kann nicht gelesen werden"
    });
  }

  for (const s of ["Angebot","Kunde","Positionen"]) {
    if (!wb.Sheets[s])
      return res.status(422).json({
        ok:false,
        message:`Tabelle fehlt: ${s}`,
        details:{ sheet:s }
      });
  }

  const angebot = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Angebot"], { header:1, defval:"" })
      .slice(1).filter(r=>r[0]).map(r=>[r[0], r[1]])
  );
  if (!angebot.taxType)
    return res.status(422).json({
      ok:false,
      message:"Feld fehlt: Angebot.taxType",
      details:{sheet:"Angebot", field:"taxType"}
    });

  const kunde = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Kunde"], { header:1, defval:"" })
      .slice(1).filter(r=>r[0]).map(r=>[r[0], r[1]])
  );
  if (!kunde.name)
    return res.status(422).json({
      ok:false,
      message:"Feld fehlt: Kunde.name",
      details:{sheet:"Kunde", field:"name"}
    });

  const pos = xlsx.utils.sheet_to_json(wb.Sheets["Positionen"], { defval:"" });
  if (!pos.length)
    return res.status(422).json({
      ok:false,
      message:"Keine Positionen vorhanden",
      details:{sheet:"Positionen"}
    });

  const byType = {};
  for (let i=0;i<pos.length;i++){
    const r = pos[i];
    const row = i+2;

    if (!r.type || !r.name)
      return res.status(422).json(
        validationError("Typ oder Positionsname fehlt.", "Positionen", row, "type/name")
      );

    const rawQty = r.qty ?? r.quantity;
    const qty = parseNumberValue(rawQty);
    if (!Number.isFinite(qty) || qty<=0)
      return res.status(422).json(
        validationError(
          "Menge muss größer als 0 sein.",
          "Positionen",
          row,
          r.qty !== undefined ? "qty" : "quantity"
        )
      );

    const rawPrice = r.price ?? r.unitPriceAmount;
    const price = parseNumberValue(rawPrice);
    if (!Number.isFinite(price) || price<0)
      return res.status(422).json(
        validationError(
          "Preis muss 0 oder größer sein.",
          "Positionen",
          row,
          r.price !== undefined ? "price" : "unitPriceAmount"
        )
      );

    if (String(r.type).toLowerCase()==="material" && !r.articleId)
      return res.status(422).json(
        validationError(
          "articleId ist für Material erforderlich.",
          "Positionen",
          row,
          "articleId"
        )
      );

    byType[r.type] = (byType[r.type] || 0) + 1;
    r.qty = qty;
    r.price = price;
  }

  if (!process.env.LEXWARE_API_KEY) {
    return res.status(500).json({ ok:false, message:"LEXWARE_API_KEY ist nicht gesetzt" });
  }

  const baseUrl = "https://api.lexware.io";

  // Angebots-Payload bauen
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
      try { errorBody = await createRes.json(); } catch {}
      return res.status(502).json({
        ok:false,
        message:"Fehler beim Erstellen des Angebots in Lexware",
        status:createRes.status,
        error:errorBody
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
      try { errorText = await fileRes.text(); } catch {}

      const isRateLimit = fileRes.status === 429;

      return res.status(fileRes.status).json({
        ok:false,
        message: isRateLimit
          ? "Angebot wurde erstellt, aber das Lexware-API-Limit ist erreicht (429). Das PDF kann aktuell nicht geladen werden. Bitte öffne das Angebot direkt in Lexware oder starte den PDF-Download später erneut."
          : "Angebot erstellt, aber PDF konnte nicht geladen werden",
        status:fileRes.status,
        quotationId,
        error:errorText
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

    // kleine Zusammenfassung im Header
    res.setHeader(
      "X-Lexware-Quote-Summary",
      encodeURIComponent(JSON.stringify({
        customer:kunde.name,
        taxType:angebot.taxType,
        positions:pos.length,
        byType
      }))
    );

    return res.send(buffer);
  } catch (e) {
    console.error("Fehler Lexware-API:", e);
    return res.status(500).json({
      ok:false,
      message:"Unerwarteter Fehler bei der Lexware-API",
      error:String(e)
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, ()=>console.log("Server läuft"));
