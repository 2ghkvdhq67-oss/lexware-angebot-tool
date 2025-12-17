import "dotenv/config";
import express from "express";
import path from "path";
import fs from "fs";
import multer from "multer";
import xlsx from "xlsx";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static(path.resolve("./public")));

// ---------- Health ----------
app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

// ---------- Template ----------
app.get("/download-template-with-articles", (_req, res) => {
  const p = path.resolve("./templates/Lexware_Template.xlsx");
  if (!fs.existsSync(p)) return res.status(500).send("Template fehlt");
  res.download(p, "Lexware_Template.xlsx");
});

// ---------- API Test ----------
app.get("/api-test", async (_req, res) => {
  try {
    const r = await fetch("https://api.lexware.io/v1/profile", {
      headers: {
        Authorization: `Bearer ${process.env.LEXWARE_API_KEY}`,
        Accept: "application/json"
      }
    });
    if (!r.ok) return res.status(500).json({ ok: false, status: r.status });
    const d = await r.json();
    res.json({ ok: true, org: d.organizationName || "OK" });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e) });
  }
});

// ---------- Excel VALIDIEREN ----------
app.post("/validate-excel", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, message: "Keine Datei" });

  const wb = xlsx.read(req.file.buffer, { type: "buffer" });
  for (const s of ["Angebot", "Kunde", "Positionen"]) {
    if (!wb.Sheets[s]) return res.status(422).json({ ok:false, message:`Sheet fehlt: ${s}` });
  }

  const angebot = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Angebot"], { header:1 })
      .slice(1).filter(r=>r[0]).map(r=>[r[0],r[1]])
  );
  if (!angebot.taxType) return res.status(422).json({ ok:false, message:"Angebot.taxType fehlt" });

  const kunde = Object.fromEntries(
    xlsx.utils.sheet_to_json(wb.Sheets["Kunde"], { header:1 })
      .slice(1).filter(r=>r[0]).map(r=>[r[0],r[1]])
  );
  if (!kunde.name) return res.status(422).json({ ok:false, message:"Kunde.name fehlt" });

  const pos = xlsx.utils.sheet_to_json(wb.Sheets["Positionen"], { defval:"" });
  if (!pos.length) return res.status(422).json({ ok:false, message:"Keine Positionen" });

  const byType = {};
  for (let i=0;i<pos.length;i++){
    const r = pos[i];
    const row = i+2;

    if (!r.type || !r.name)
      return res.status(422).json({ ok:false, message:"type/name fehlt", details:{row} });

    const qty = Number(r.qty);
    if (!Number.isFinite(qty) || qty<=0)
      return res.status(422).json({ ok:false, message:"qty > 0", details:{row} });

    const price = Number(r.price);
    if (!Number.isFinite(price) || price<0)
      return res.status(422).json({ ok:false, message:"price >= 0", details:{row} });

    if (String(r.type).toLowerCase()==="material" && !r.articleId)
      return res.status(422).json({ ok:false, message:"articleId fehlt", details:{row} });

    byType[r.type]=(byType[r.type]||0)+1;
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

// ---------- ANGEBOT + PDF ----------
app.post("/create-quote-from-excel", upload.single("file"), async (req, res) => {
  // bewusst simpel: erst validieren
  const r = await fetch("http://localhost/health"); // Platzhalter, Logik steht
  return res.status(501).json({ ok:false, message:"Angebotserstellung folgt im nächsten Schritt" });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, ()=>console.log("Server läuft"));
