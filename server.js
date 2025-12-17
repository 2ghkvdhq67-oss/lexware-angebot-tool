import "dotenv/config";
import express from "express";
import path from "path";
import fs from "fs";
import multer from "multer";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Statische Dateien (index.html)
app.use(express.static(path.resolve("./public")));

// ========== HEALTH CHECK ==========
app.get("/health", (req, res) => {
  res.json({ ok: true, message: "Server läuft" });
});

// ========== TEMPLATE DOWNLOAD ==========
app.get("/download-template-with-articles", (req, res) => {
  const templatePath = path.resolve("./templates/Lexware_Template.xlsx");

  if (!fs.existsSync(templatePath)) {
    return res.status(500).send("Template fehlt: templates/Lexware_Template.xlsx");
  }

  res.download(templatePath, "Lexware_Template.xlsx");
});

// ========== TESTMODUS (WICHTIG!) ==========
app.post("/validate-excel", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      ok: false,
      message: "Keine Excel-Datei hochgeladen"
    });
  }

  // KEINE Excel-Logik, KEINE API – nur schneller Test
  return res.json({
    ok: true,
    summary: {
      message: "✅ Testmodus erfolgreich",
      fileName: req.file.originalname,
      fileSizeKB: Math.round(req.file.size / 1024)
    }
  });
});

// ========== PLATZHALTER FÜR SPÄTER ==========
app.post("/create-quote-from-excel", upload.single("file"), async (req, res) => {
  res.status(501).json({
    ok: false,
    message: "Noch nicht implementiert"
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server läuft auf Port ${PORT}`);
});
