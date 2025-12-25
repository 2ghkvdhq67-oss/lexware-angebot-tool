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

// Testmodus: Excel lesen und validieren
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

// Platzhalter für später
app.post("/create-quote-from-excel", upload.single("file"), (_req, res) =>
  res.status(501).json({ ok:false, message:"Angebot/PDF folgt" })
);

const PORT = process.env.PORT || 3000;
app.listen(PORT, ()=>console.log("Server läuft"));
