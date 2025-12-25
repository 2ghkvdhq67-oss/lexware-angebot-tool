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
    return res.status(500).json({ ok:false, message:"Template-Datei fehlt auf dem Server" });

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

    if (!r.ok)
      return res.status(500).json({
        ok:false,
        message:"API-Verbindung fehlgeschlagen",
        status:r.status
      });

    const d = await r.json();
    res.json({
      ok:true,
      organization: d.organizationName || "API-Verbindung erfolgreich"
    });
  } catch (e) {
    res.status(500).json({
      ok:false,
      message:"API-Test konnte nicht ausgeführt werden",
      error:String(e)
    });
  }
});


// ---------------------------------------------------------
// Helper für Klartext-Fehlermeldungen
// ---------------------------------------------------------
function validationError(message, sheet, row, field) {
  return {
    ok:false,
    message:`${message} (Tabelle: ${sheet}, Zeile ${row})`,
    details:{ sheet, row, field }
  };
}


// ---------------------------------------------------------
// TESTMODUS — Excel prüfen & verständliche Fehlermeldungen
// ---------------------------------------------------------
app.post("/validate-excel", upload.single("file"), (req, res) => {

  if (!req.file)
    return res.status(400).json({
      ok:false,
      message:"Es wurde keine Datei hochgeladen"
    });

  let wb;
  try {
    wb = xlsx.read(req.file.buffer, { type:"buffer" });
  } catch {
    return res.status(400).json({
      ok:false,
      message:"Die Excel-Datei konnte nicht gelesen werden"
    });
  }

  // Pflicht-Sheets
  for (const s of ["Angebot","Kunde","Positionen"]) {
    if (!wb.Sheets[s])
      return res.status(422).json({
        ok:false,
        message:`Die Tabelle "${s}" fehlt in der Excel-Datei`,
        details:{ sheet:s }
      });
  }

  // Angebot
  const angebot =
