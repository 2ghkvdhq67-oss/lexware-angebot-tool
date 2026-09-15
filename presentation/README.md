# Maiershirts – Präsentationsmaster

PowerPoint-Vorlage (Master) im Maiershirts-Look mit Logo, passend zum
Farbschema des Lexware-Angebots-Tools (dunkles Anthrazit `#0B0F14`, Akzent-Grün `#22C55E`).

## Dateien

| Datei | Zweck |
|---|---|
| `dist/Maiershirts_Master.pptx` | Beispieldeck mit allen Layouts (10 Folien) – zum Kopieren und Anpassen |
| `dist/Maiershirts_Master.potx` | Vorlage: Doppelklick öffnet eine neue Präsentation mit den Mastern |
| `assets/logo-dark.png` | Logo für helle Folien (dunkle Schrift) |
| `assets/logo-light.png` | Logo für dunkle Folien (helle Schrift) |
| `build.js` | Erzeugt beide Dateien neu (`npm run build`) |
| `make-logo.js` | Erzeugt die Platzhalter-Wortmarke (`npm run logo`) |

## Master / Layouts

| Master | Verwendung |
|---|---|
| `MS_TITEL` | Titelfolie, dunkel, großes Logo |
| `MS_ABSCHNITT` | Abschnittstrenner, dunkel, Nummer + Titel |
| `MS_INHALT` | Inhaltsfolie, weiß, Logo oben rechts, Fußzeile mit Seitenzahl |
| `MS_ABSCHLUSS` | Abschlussfolie, dunkel, Logo zentriert, Kontakt |

Das Beispieldeck zeigt pro Layout ein Muster: Agenda mit Icon-Zeilen,
Zwei-Spalten-Folie mit Karten, Kennzahlen, Prozess (4 Schritte),
Angebotstabelle, natives Diagramm und eine leere Standardfolie.

## Echtes Logo einsetzen

Die Wortmarke in `assets/` ist ein **Platzhalter**. Sobald das Original-Logo vorliegt:

1. `assets/logo-dark.png` durch das Logo für helle Hintergründe ersetzen,
   `assets/logo-light.png` durch die Variante für dunkle Hintergründe
   (PNG mit transparentem Hintergrund, gern breit, z. B. 1100 × 200 px).
   Liegt nur eine Datei vor, wird sie für beide Fälle verwendet.
2. `npm install` (einmalig) und `npm run build` ausführen.

Das Logo wird proportional in die vorgesehene Fläche eingepasst, das
Seitenverhältnis bleibt erhalten.

## Anpassen

- Texte, Kennzahlen, Tabellenwerte und Kontaktdaten im Beispieldeck sind
  Beispielinhalte (siehe Notizen der jeweiligen Folie) und müssen vor
  Verwendung ersetzt werden.
- Farben stehen am Anfang von `build.js` im Objekt `C`.
- Schrift: Calibri (Standard in Office).
