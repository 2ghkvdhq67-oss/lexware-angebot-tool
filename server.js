import "dotenv/config";
import express from "express";
import path from "path";
import fs from "fs";

const app = express();

// statische Dateien (index.html)
app.use(express.static(path.resolve("./public")));

// 🔽 DIESER ENDPOINT FEHLTE BISHER
app.get("/download-template-with-articles", (req, res) => {
  const templatePath = path.resolve("./templates/Lexware_Template.xlsx");

  if (!fs.existsSync(templatePath)) {
    return res.status(500).send("Template fehlt: templates/Lexware_Template.xlsx");
  }

  res.download(templatePath, "Lexware_Template.xlsx");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server läuft auf Port ${PORT}`);
});
