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

// Zahl aus Excel robust parsen (inkl. "6,9")
function parseNumberValue(value) {
  if (value === undefined || value === null || value === "") return NaN;
  if (typeof value === "string") {
    // Komma in Punkt umwandeln, Leerzeichen entfernen
    value = value.replace(",", ".").trim();
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
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

  //
