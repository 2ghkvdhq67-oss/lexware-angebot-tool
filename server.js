// Starter server.js
import express from "express";
const app = express();
app.use(express.static("public"));
app.listen(3000, ()=>console.log("Server läuft auf http://localhost:3000"));
