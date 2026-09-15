# Maiershirts – Präsentationsmaster

PowerPoint-Vorlage (Master) im Maiershirts-Look mit Logo, nach den
Brand Colors (Stand Mai 2026): Schwarz `#0f0f0f`, Asparagus `#819E72`,
Warm Sand `#E8DFD0`, Weiß. Schrift: Inter (Fallback in PowerPoint: Arial/Calibri,
falls Inter nicht installiert ist – Inter gibt es kostenlos bei Google Fonts).

## Dateien

| Datei | Zweck |
|---|---|
| `dist/Maiershirts_Master.pptx` | Beispieldeck mit allen Layouts (10 Folien) – zum Kopieren und Anpassen |
| `dist/Maiershirts_Master.potx` | Vorlage: Doppelklick öffnet eine neue Präsentation mit den Mastern |
| `assets/logo.svg` | Logo als Vektor (Bergmarke + Wortmarke), nachgebaut nach der Originalvorlage |
| `assets/logo-dark.png` | Logo schwarz für helle Folien |
| `assets/logo-light.png` | Logo weiß für dunkle Folien |
| `build.js` | Erzeugt beide Dateien neu (`npm run build`) |
| `make-logo.js` | Erzeugt Logo-SVG und die beiden PNGs (`npm run logo`) |

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

## Original-Logodatei einsetzen

Das Logo in `assets/` ist ein Vektor-Nachbau der Originalvorlage. Um stattdessen
die Originaldatei (z. B. aus `Downloads/firmenlogo`) zu verwenden:

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
