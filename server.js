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
  if (!fs.existsSync(p)) return res.status(500).json({ ok:false, message:"Template fehlt" });
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

// Testmodus: Excel lesen & validieren
app.post("/validate-excel", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ ok:false, message:"Keine Datei" });

  let wb;
  try {
    wb = xlsx.read(req.file.buffer, { type:"buffer" });
  } catch {
    return res.status(400).json({ ok:false, message:"Excel nicht lesbar" });
  }

  for (const s of ["Angebot","Kunde","Positionen"]) {
    if (!wb.Sheets[s]) return res.status(422).json({ ok:false, message:`Sheet fehlt: ${s}` });
  }

  const angebot = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Angebot"], { header:1, defval:"" })
      .slice(1).filter(r=>r[0]).map(r=>[r[0], r[1]])
  );
  if (!angebot.taxType)
    return res.status(422).json({ ok:false, message:"Angebot.taxType fehlt", details:{sheet:"Angebot", field:"taxType"} });

  const kunde = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Kunde"], { header:1, defval:"" })
      .slice(1).filter(r=>r[0]).map(r=>[r[0], r[1]])
  );
  if (!kunde.name)
    return res.status(422).json({ ok:false, message:"Kunde.name fehlt", details:{sheet:"Kunde", field:"name"} });

  const pos = xlsx.utils.sheet_to_json(wb.Sheets["Positionen"], { defval:"" });
  if (!pos.length)
    return res.status(422).json({ ok:false, message:"Keine Positionen", details:{sheet:"Positionen"} });

  const byType = {};
  for (let i=0;i<pos.length;i++){
    const r = pos[i];
    const row = i+2;

    if (!r.type || !r.name)
      return res.status(422).json({ ok:false, message:"type/name fehlt", details:{sheet:"Positionen", row, field:"type/name"} });

    const qty = Number(r.qty);
    if (!Number.isFinite(qty) || qty<=0)
      return res.status(422).json({ ok:false, message:"qty > 0", details:{sheet:"Positionen", row, field:"qty"} });

    const price = Number(r.price);
    if (!Number.isFinite(price) || price<0)
      return res.status(422).json({ ok:false, message:"price >= 0", details:{sheet:"Positionen", row, field:"price"} });

    if (String(r.type).toLowerCase()==="material" && !r.articleId)
      return res.status(422).json({ ok:false, message:"articleId fehlt", details:{sheet:"Positionen", row, field:"articleId"} });

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
