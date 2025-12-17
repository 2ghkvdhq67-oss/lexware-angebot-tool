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
  res.json({ ok: true, message: "Server läuft" });
});

// ---------- Template Download ----------
app.get("/download-template-with-articles", (_req, res) => {
  const p = path.resolve("./templates/Lexware_Template.xlsx");
  if (!fs.existsSync(p)) {
    return res.status(500).send("Template fehlt: templates/Lexware_Template.xlsx");
  }
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
    if (!r.ok) {
      const t = await r.text();
      return res.status(500).json({ ok: false, status: r.status, details: t });
    }
    const d = await r.json();
    res.json({ ok: true, message: "API verbunden", org: d.organizationName || "OK" });
  } catch (e) {
    res.status(500).json({ ok: false, message: "API-Test fehlgeschlagen", error: String(e) });
  }
});

// ---------- Testmodus: Excel lesen & validieren ----------
app.post("/validate-excel", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, message: "Keine Datei" });

  let wb;
  try {
    wb = xlsx.read(req.file.buffer, { type: "buffer" });
  } catch {
    return res.status(400).json({ ok: false, message: "Excel nicht lesbar" });
  }

  // Pflicht-Sheets
  for (const s of ["Angebot", "Kunde", "Positionen"]) {
    if (!wb.Sheets[s]) {
      return res.status(422).json({ ok: false, message: `Sheet fehlt: ${s}`, details: { sheet: s } });
    }
  }

  // Angebot (Key/Value)
  const aRows = xlsx.utils.sheet_to_json(wb.Sheets["Angebot"], { header: 1, defval: "" });
  const angebot = Object.fromEntries(aRows.slice(1).filter(r => r[0]).map(r => [r[0], r[1]]));
  if (!angebot.taxType) {
    return res.status(422).json({ ok: false, message: "Angebot.taxType fehlt", details: { sheet: "Angebot", field: "taxType" } });
  }

  // Kunde
  const kRows = xlsx.utils.sheet_to_json(wb.Sheets["Kunde"], { header: 1, defval: "" });
  const kunde = Object.fromEntries(kRows.slice(1).filter(r => r[0]).map(r => [r[0], r[1]]));
  if (!kunde.name) {
    return res.status(422).json({ ok: false, message: "Kunde.name fehlt", details: { sheet: "Kunde", field: "name" } });
  }

  // Positionen
  const pos = xlsx.utils.sheet_to_json(wb.Sheets["Positionen"], { defval: "" });
  if (!pos.length) {
    return res.status(422).json({ ok: false, message: "Keine Positionen", details: { sheet: "Positionen" } });
  }

  const byType = {};
  for (let i = 0; i < pos.length; i++) {
    const r = pos[i];
    const row = i + 2;
    if (!r.type || !r.name) {
      return res.status(422).json({
        ok: false,
        message: "Position unvollständig (type/name)",
        details: { sheet: "Positionen", row, field: "type/name" }
      });
    }
    byType[r.type] = (byType[r.type] || 0) + 1;
  }

  return res.json({
    ok: true,
    summary: {
      message: "Excel validiert",
      taxType: angebot.taxType,
      customer: kunde.name,
      positions: pos.length,
      byType
    }
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server läuft auf Port ${PORT}`));
