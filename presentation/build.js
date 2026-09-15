/**
 * Maiershirts – PowerPoint-Master (Folienvorlage) mit Logo
 *
 * Erzeugt dist/Maiershirts_Master.pptx (Beispieldeck mit allen Layouts)
 * und dist/Maiershirts_Master.potx (Vorlage zum Öffnen als "Neue Präsentation").
 *
 * Logo: assets/logo-dark.png (für helle Folien) und assets/logo-light.png
 * (für dunkle Folien). Beide Dateien können durch das Original-Logo ersetzt
 * werden; das Seitenverhältnis wird automatisch beibehalten ("contain").
 *
 * Aufruf: npm run build
 */
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const Fa = require('react-icons/fa');
const JSZip = require('jszip');

// ---------- Farben: Maiershirts Brand Colors (Stand Mai 2026) ----------
const C = {
  dark: '0F0F0F',      // Schwarz – Text / Struktur / dunkle Blöcke
  green: '819E72',     // Asparagus – Akzent / Buttons / Highlights
  greenDark: '677E5B', // Asparagus Dark – Hover / Akzent-Text
  greenLight: 'A3C095',// Asparagus Light – Badges / leichte Akzente
  greenTint: 'E3ECDD', // sehr helle Asparagus-Fläche (Kreise auf Weiß)
  beige: 'E8DFD0',     // Warm Sand – Hintergrund-Sektionen / Karten
  beigeDark: 'D9D2C5', // Warm Sand dunkel – Linien
  cream: 'FAF7F2',     // Cream – Body-Hintergrund
  white: 'FFFFFF',
  text: '0F0F0F',      // Text auf Weiß
  textMuted: '5C5C58', // gedämpfter Text auf Weiß
  light: 'FFFFFF',     // Text auf Schwarz
  muted: 'CFC8BB',     // gedämpfter Text auf Schwarz (Warm Sand abgedunkelt)
  grayBg: 'E8DFD0',    // Kartenfläche auf Weiß = Warm Sand
  line: 'D9D2C5',      // Tabellen-/Trennlinien
};
const FONT = 'Inter';
const DIST = path.join(__dirname, 'dist');
const ASSETS = path.join(__dirname, 'assets');

// ---------- Logo laden ----------
function logoPath(kind) {
  const preferred = path.join(ASSETS, `logo-${kind}.png`);
  const fallback = path.join(ASSETS, kind === 'dark' ? 'logo-light.png' : 'logo-dark.png');
  if (fs.existsSync(preferred)) return preferred;
  if (fs.existsSync(fallback)) return fallback;
  throw new Error('Kein Logo gefunden. Bitte assets/logo-dark.png und assets/logo-light.png anlegen (npm run logo erzeugt Platzhalter).');
}
const LOGO_DARK = logoPath('dark');   // auf hellem Grund
const LOGO_LIGHT = logoPath('light'); // auf dunklem Grund

// Logo als Bild-Objekt: Box (x, y, w, h), Logo wird proportional eingepasst.
function logo(kind, x, y, w, h) {
  return {
    image: {
      path: kind === 'dark' ? LOGO_DARK : LOGO_LIGHT,
      x, y, w, h,
      sizing: { type: 'contain', w, h },
    },
  };
}

// ---------- Icons (react-icons -> PNG) ----------
async function iconPng(Icon, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color: '#' + color, size: 256 })
  );
  const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

// Icon in farbigem Kreis (das wiederkehrende Motiv des Masters)
function iconCircle(slide, data, x, y, d, circleColor) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: circleColor }, line: { color: circleColor } });
  const pad = d * 0.27;
  slide.addImage({ data, x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}

// Hintergrund für dunkle Master: Grundfarbe plus weiche grüne Kreise (das Motiv)
async function darkBackground(name, circles) {
  const svgCircles = circles.map(([cx, cy, r, op]) => `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#${C.green}" fill-opacity="${op}"/>`).join('');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080" viewBox="0 0 1920 1080"><rect width="1920" height="1080" fill="#${C.dark}"/>${svgCircles}</svg>`;
  const out = path.join(require('os').tmpdir(), `maiershirts-bg-${name}.png`);
  await sharp(Buffer.from(svg)).png().toFile(out);
  return out;
}

async function main() {
  fs.mkdirSync(DIST, { recursive: true });
  // Koordinaten in px auf 1920x1080 (192 px = 1 Zoll)
  const BG_TITEL = await darkBackground('titel', [[1824, 192, 500, 0.22], [1805, 1000, 250, 0.32]]);
  const BG_ABSCHNITT = await darkBackground('abschnitt', [[1786, 538, 326, 0.24]]);
  const BG_ABSCHLUSS = await darkBackground('abschluss', [[1880, 40, 300, 0.22]]);

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9'; // 10" x 5.625"
  pres.author = 'Maiershirts';
  pres.company = 'Maiershirts';
  pres.title = 'Maiershirts – Präsentationsmaster';
  pres.lang = 'de-DE';

  // ---------- Master 1: Titelfolie (dunkel) ----------
  pres.defineSlideMaster({
    title: 'MS_TITEL',
    background: { path: BG_TITEL },
    objects: [
      // dezenter grüner Kreis als Motiv, rechts angeschnitten
      logo('light', 0.6, 0.5, 2.2, 1.2),
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 2.05, w: 7.2, h: 1.3, fontFace: FONT, fontSize: 40, bold: true, color: C.white, valign: 'bottom', margin: 0 }, text: 'Titel der Präsentation' } },
      { placeholder: { options: { name: 'sub', type: 'body', x: 0.6, y: 3.45, w: 7.2, h: 0.7, fontFace: FONT, fontSize: 18, color: C.muted, valign: 'top', margin: 0 }, text: 'Untertitel · Datum' } },
      { text: { text: 'Maiershirts', options: { x: 0.6, y: 5.05, w: 4, h: 0.3, fontFace: FONT, fontSize: 10, color: C.muted, margin: 0 } } },
    ],
  });

  // ---------- Master 2: Abschnittsfolie (dunkel) ----------
  pres.defineSlideMaster({
    title: 'MS_ABSCHNITT',
    background: { path: BG_ABSCHNITT },
    objects: [
      logo('light', 0.6, 0.45, 1.45, 0.8),
      { placeholder: { options: { name: 'num', type: 'body', x: 0.6, y: 1.7, w: 3, h: 0.9, fontFace: FONT, fontSize: 54, bold: true, color: C.green, valign: 'bottom', margin: 0 }, text: '01' } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 2.65, w: 7, h: 1.0, fontFace: FONT, fontSize: 36, bold: true, color: C.white, valign: 'top', margin: 0 }, text: 'Abschnittstitel' } },
      { placeholder: { options: { name: 'sub', type: 'body', x: 0.6, y: 3.7, w: 7, h: 0.6, fontFace: FONT, fontSize: 16, color: C.muted, valign: 'top', margin: 0 }, text: 'Kurzbeschreibung des Abschnitts' } },
    ],
  });

  // ---------- Master 3: Inhaltsfolie (hell) ----------
  pres.defineSlideMaster({
    title: 'MS_INHALT',
    background: { color: C.white },
    objects: [
      logo('dark', 8.25, 0.3, 1.25, 0.68),
      { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.35, w: 7.5, h: 0.75, fontFace: FONT, fontSize: 28, bold: true, color: C.text, valign: 'middle', margin: 0 }, text: 'Folientitel' } },
      { text: { text: 'Maiershirts', options: { x: 0.5, y: 5.15, w: 3, h: 0.3, fontFace: FONT, fontSize: 9, color: C.textMuted, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.15, w: 0.5, h: 0.3, fontFace: FONT, fontSize: 9, color: C.textMuted, align: 'right', margin: 0 },
  });

  // ---------- Master 4: Abschlussfolie (dunkel) ----------
  pres.defineSlideMaster({
    title: 'MS_ABSCHLUSS',
    background: { path: BG_ABSCHLUSS },
    objects: [
      logo('light', 3.7, 0.7, 2.6, 1.42),
      { placeholder: { options: { name: 'title', type: 'title', x: 1, y: 2.25, w: 8, h: 0.9, align: 'center', fontFace: FONT, fontSize: 36, bold: true, color: C.white, valign: 'middle', margin: 0 }, text: 'Vielen Dank' } },
      { placeholder: { options: { name: 'sub', type: 'body', x: 1, y: 3.2, w: 8, h: 0.9, align: 'center', fontFace: FONT, fontSize: 14, color: C.muted, valign: 'top', margin: 0 }, text: 'Kontakt' } },
    ],
  });

  // ---------- Icons vorbereiten ----------
  const ic = {
    users: await iconPng(Fa.FaUsers, C.greenDark),
    shirt: await iconPng(Fa.FaTshirt, C.greenDark),
    steps: await iconPng(Fa.FaClipboardList, C.greenDark),
    invoice: await iconPng(Fa.FaFileInvoiceDollar, C.greenDark),
    print: await iconPng(Fa.FaPrint, C.greenDark),
    palette: await iconPng(Fa.FaPalette, C.greenDark),
    truck: await iconPng(Fa.FaTruck, C.greenDark),
    check: await iconPng(Fa.FaCheckCircle, C.greenDark),
    bolt: await iconPng(Fa.FaBolt, C.greenDark),
    // helle Varianten für dunkle Folien
    mailL: await iconPng(Fa.FaEnvelope, C.dark),
    phoneL: await iconPng(Fa.FaPhone, C.dark),
    globeL: await iconPng(Fa.FaGlobe, C.dark),
  };

  // ======================================================================
  // Folie 1 – Titel
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_TITEL' });
    s.addText('Angebotspräsentation', { placeholder: 'title', isTextBox: true });
    s.addText('Individuell bedruckte Textilien für Teams, Vereine und Unternehmen · September 2026', { placeholder: 'sub', isTextBox: true });
    s.addNotes('Titelfolie (Master „MS_TITEL“). Titel und Untertitel sind Platzhalter – einfach überschreiben.');
  }

  // ======================================================================
  // Folie 2 – Agenda (Icon-Zeilen)
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    s.addText('Agenda', { placeholder: 'title', isTextBox: true });
    const items = [
      ['01', 'Über uns', 'Wer wir sind und wofür wir stehen', ic.users],
      ['02', 'Leistungen', 'Druckverfahren, Textilien und Veredelung', ic.shirt],
      ['03', 'Ablauf', 'Von der Anfrage bis zur Lieferung', ic.steps],
      ['04', 'Angebot', 'Positionen, Preise und nächste Schritte', ic.invoice],
    ];
    items.forEach(([num, head, desc, icon], i) => {
      const y = 1.35 + i * 0.95;
      iconCircle(s, icon, 0.5, y, 0.62, C.greenTint);
      s.addText(num, { x: 1.35, y, w: 0.6, h: 0.62, fontFace: FONT, fontSize: 20, bold: true, color: C.greenDark, valign: 'middle', margin: 0, isTextBox: true });
      s.addText(head, { x: 2.0, y: y - 0.02, w: 6.5, h: 0.34, fontFace: FONT, fontSize: 18, bold: true, color: C.text, valign: 'bottom', margin: 0, isTextBox: true });
      s.addText(desc, { x: 2.0, y: y + 0.32, w: 6.5, h: 0.3, fontFace: FONT, fontSize: 12, color: C.textMuted, valign: 'top', margin: 0, isTextBox: true });
    });
    s.addNotes('Agenda-Layout: Icon im grünen Kreis, Nummer, Überschrift und Kurztext pro Zeile.');
  }

  // ======================================================================
  // Folie 3 – Abschnitt
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_ABSCHNITT' });
    s.addText('01', { placeholder: 'num', isTextBox: true });
    s.addText('Über uns', { placeholder: 'title', isTextBox: true });
    s.addText('Textilveredelung aus einer Hand – persönlich, schnell und in gleichbleibender Qualität', { placeholder: 'sub', isTextBox: true });
    s.addNotes('Abschnittsfolie (Master „MS_ABSCHNITT“): Nummer, Titel, Kurzbeschreibung.');
  }

  // ======================================================================
  // Folie 4 – Zwei Spalten: Text links, Feature-Karten rechts
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    s.addText('Was uns ausmacht', { placeholder: 'title', isTextBox: true });
    s.addText(
      'Maiershirts veredelt Textilien für Unternehmen, Vereine und Events. Vom einzelnen Shirt bis zur kompletten Teamausstattung begleiten wir jedes Projekt persönlich – von der Motividee bis zum fertigen Paket.',
      { x: 0.5, y: 1.35, w: 4.1, h: 1.5, fontFace: FONT, fontSize: 14, color: C.text, valign: 'top', margin: 0, isTextBox: true }
    );
    s.addText([
      { text: 'Beratung zu Textil, Verfahren und Motiv', options: { bullet: true, breakLine: true } },
      { text: 'Druck und Stickerei im eigenen Haus', options: { bullet: true, breakLine: true } },
      { text: 'Kleine Auflagen ab 1 Stück', options: { bullet: true, breakLine: true } },
      { text: 'Verlässliche Liefertermine', options: { bullet: true } },
    ], { x: 0.5, y: 2.95, w: 4.1, h: 1.8, fontFace: FONT, fontSize: 14, color: C.text, valign: 'top', margin: 0, paraSpaceAfter: 6, isTextBox: true });

    const cards = [
      ['Siebdruck', 'Kräftige Farben, ideal für größere Auflagen', ic.print],
      ['Stickerei', 'Hochwertig und langlebig für Workwear', ic.palette],
      ['Express', 'Kurze Produktionszeiten auf Anfrage', ic.bolt],
      ['Qualität', 'Geprüfte Textilien namhafter Hersteller', ic.check],
    ];
    cards.forEach(([head, desc, icon], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 5.0 + col * 2.3, y = 1.35 + row * 1.85;
      s.addShape('roundRect', { x, y, w: 2.15, h: 1.7, fill: { color: C.grayBg }, line: { color: C.grayBg }, rectRadius: 0.1 });
      iconCircle(s, icon, x + 0.2, y + 0.2, 0.5, C.white);
      s.addText(head, { x: x + 0.2, y: y + 0.8, w: 1.8, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: C.text, margin: 0, isTextBox: true });
      s.addText(desc, { x: x + 0.2, y: y + 1.1, w: 1.8, h: 0.5, fontFace: FONT, fontSize: 10.5, color: C.textMuted, valign: 'top', margin: 0, isTextBox: true });
    });
    s.addNotes('Zwei-Spalten-Layout: Fließtext und Aufzählung links, 2x2-Kartenraster mit Icons rechts. Beispieltext – bitte anpassen.');
  }

  // ======================================================================
  // Folie 5 – Kennzahlen
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    s.addText('Zahlen, die für uns sprechen', { placeholder: 'title', isTextBox: true });
    const stats = [
      ['1.200+', 'Projekte pro Jahr', 'Beispielwert'],
      ['48 h', 'Angebot in Arbeitstagen', 'Beispielwert'],
      ['98 %', 'Wiederkehrende Kunden', 'Beispielwert'],
    ];
    stats.forEach(([big, label, note], i) => {
      const x = 0.5 + i * 3.05;
      s.addShape('roundRect', { x, y: 1.5, w: 2.9, h: 2.6, fill: { color: C.grayBg }, line: { color: C.grayBg }, rectRadius: 0.12 });
      s.addText(big, { x: x + 0.25, y: 1.75, w: 2.4, h: 1.1, fontFace: FONT, fontSize: 44, bold: true, color: C.greenDark, valign: 'middle', margin: 0, isTextBox: true });
      s.addText(label, { x: x + 0.25, y: 2.9, w: 2.4, h: 0.65, fontFace: FONT, fontSize: 14, valign: 'top', bold: true, color: C.text, margin: 0, isTextBox: true });
      s.addText(note, { x: x + 0.25, y: 3.55, w: 2.4, h: 0.35, fontFace: FONT, fontSize: 11, italic: true, color: C.textMuted, margin: 0, isTextBox: true });
    });
    s.addText('Alle Werte sind Platzhalter und werden vor Verwendung durch echte Kennzahlen ersetzt.', { x: 0.5, y: 4.4, w: 9, h: 0.35, fontFace: FONT, fontSize: 11, color: C.textMuted, margin: 0, isTextBox: true });
    s.addNotes('Kennzahlen-Layout: drei große Zahlen mit Beschriftung. Werte sind Beispiele.');
  }

  // ======================================================================
  // Folie 6 – Ablauf (Prozess mit 4 Schritten)
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    s.addText('So läuft ein Auftrag ab', { placeholder: 'title', isTextBox: true });
    const steps = [
      ['Anfrage', 'Textil, Menge und Motiv per Mail oder Telefon', ic.users],
      ['Angebot', 'Angebot aus Lexware innerhalb von 48 h', ic.invoice],
      ['Produktion', 'Druck oder Stickerei nach Freigabe', ic.print],
      ['Lieferung', 'Versand oder Abholung zum Wunschtermin', ic.truck],
    ];
    const d = 1.0, gap = 2.3, x0 = 0.6, yC = 1.75;
    steps.forEach(([head, desc, icon], i) => {
      const x = x0 + i * gap;
      if (i < steps.length - 1) {
        s.addShape('line', { x: x + d + 0.15, y: yC + d / 2, w: gap - d - 0.3, h: 0, line: { color: C.greenLight, width: 2, endArrowType: 'triangle' } });
      }
      iconCircle(s, icon, x, yC, d, C.greenTint);
      s.addText(String(i + 1), { x: x + d - 0.3, y: yC - 0.1, w: 0.36, h: 0.36, fontFace: FONT, fontSize: 11, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0, fill: { color: C.greenDark }, shape: 'ellipse', isTextBox: true });
      s.addText(head, { x: x - 0.55, y: yC + d + 0.25, w: d + 1.1, h: 0.35, fontFace: FONT, fontSize: 16, bold: true, color: C.text, align: 'center', margin: 0, isTextBox: true });
      s.addText(desc, { x: x - 0.55, y: yC + d + 0.62, w: d + 1.1, h: 0.9, fontFace: FONT, fontSize: 11.5, color: C.textMuted, align: 'center', valign: 'top', margin: 0, isTextBox: true });
    });
    s.addNotes('Prozess-Layout: vier Schritte mit Icons und Pfeilen.');
  }

  // ======================================================================
  // Folie 7 – Angebotstabelle
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    s.addText('Angebotsübersicht', { placeholder: 'title', isTextBox: true });
    const hdr = (t, align = 'left') => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.dark }, align, fontSize: 12 } });
    const cell = (t, align = 'left', opts = {}) => ({ text: t, options: { align, fontSize: 12, color: C.text, ...opts } });
    const rows = [
      [hdr('Pos.'), hdr('Artikel'), hdr('Menge', 'right'), hdr('Einzelpreis', 'right'), hdr('Gesamt', 'right')],
      [cell('1'), cell('T-Shirt Bio-Baumwolle, Brustdruck 1-farbig'), cell('50', 'right'), cell('12,90 €', 'right'), cell('645,00 €', 'right')],
      [cell('2'), cell('Hoodie, Rückendruck 2-farbig'), cell('25', 'right'), cell('34,50 €', 'right'), cell('862,50 €', 'right')],
      [cell('3'), cell('Polo-Shirt, Logo-Stickerei'), cell('20', 'right'), cell('26,00 €', 'right'), cell('520,00 €', 'right')],
      [cell('4'), cell('Einrichtungskosten Siebdruck'), cell('1', 'right'), cell('45,00 €', 'right'), cell('45,00 €', 'right')],
      [cell(''), cell('Netto gesamt', 'left', { bold: true }), cell(''), cell(''), cell('2.072,50 €', 'right', { bold: true, color: C.greenDark })],
    ];
    s.addTable(rows, {
      x: 0.5, y: 1.35, w: 9.0, colW: [0.6, 4.6, 1.0, 1.4, 1.4],
      fontFace: FONT, rowH: 0.42, border: { type: 'solid', color: C.line, pt: 0.75 },
      fill: { color: C.white }, valign: 'middle', margin: [0.04, 0.1, 0.04, 0.1],
    });
    s.addText('Beispielpositionen · Preise netto zzgl. MwSt. · Gültig 30 Tage', { x: 0.5, y: 4.15, w: 9, h: 0.3, fontFace: FONT, fontSize: 10.5, italic: true, color: C.textMuted, margin: 0, isTextBox: true });
    s.addNotes('Tabellen-Layout für Angebotspositionen – passend zu den Angeboten aus dem Lexware-Tool. Beispielwerte.');
  }

  // ======================================================================
  // Folie 8 – Diagramm (nativ)
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    s.addText('Auftragsvolumen nach Produktgruppe', { placeholder: 'title', isTextBox: true });
    s.addChart(pres.charts.BAR, [
      { name: 'Aufträge', labels: ['T-Shirts', 'Hoodies', 'Polos', 'Workwear', 'Caps'], values: [420, 260, 180, 150, 90] },
    ], {
      x: 0.5, y: 1.3, w: 5.8, h: 3.7,
      barDir: 'col', chartColors: [C.green],
      showTitle: false, showLegend: false,
      showValue: true, dataLabelPosition: 'outEnd', dataLabelFontFace: FONT, dataLabelFontSize: 11, dataLabelColor: C.text,
      catAxisLabelFontFace: FONT, catAxisLabelFontSize: 11, catAxisLabelColor: C.textMuted,
      valAxisLabelFontFace: FONT, valAxisLabelFontSize: 10, valAxisLabelColor: C.textMuted,
      valGridLine: { color: C.line, size: 0.5 }, catGridLine: { style: 'none' },
      valAxisLineShow: false, catAxisLineShow: false,
    });
    s.addShape('roundRect', { x: 6.6, y: 1.3, w: 2.9, h: 3.7, fill: { color: C.grayBg }, line: { color: C.grayBg }, rectRadius: 0.12 });
    s.addText('Beispieldaten', { x: 6.85, y: 1.5, w: 2.4, h: 0.3, fontFace: FONT, fontSize: 11, italic: true, color: C.textMuted, margin: 0, isTextBox: true });
    s.addText('T-Shirts machen den größten Anteil aus', { x: 6.85, y: 1.85, w: 2.4, h: 0.7, fontFace: FONT, fontSize: 16, bold: true, color: C.text, margin: 0, isTextBox: true });
    s.addText('Das Diagramm ist ein natives PowerPoint-Diagramm: Werte per Rechtsklick → „Daten bearbeiten“ ändern.', { x: 6.85, y: 2.65, w: 2.4, h: 1.2, fontFace: FONT, fontSize: 12, color: C.textMuted, valign: 'top', margin: 0, isTextBox: true });
    s.addNotes('Diagramm-Layout: natives Säulendiagramm links, Kernaussage rechts. Beispieldaten.');
  }

  // ======================================================================
  // Folie 9 – Leere Inhaltsfolie (Layout-Muster)
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    s.addText('Folientitel', { placeholder: 'title', isTextBox: true });
    s.addText([
      { text: 'Erste Aussage der Folie', options: { bullet: true, breakLine: true } },
      { text: 'Zweite Aussage mit einer kurzen Erläuterung', options: { bullet: true, breakLine: true } },
      { text: 'Dritte Aussage', options: { bullet: true, breakLine: true } },
      { text: 'Unterpunkt zur dritten Aussage', options: { bullet: true, indentLevel: 1, breakLine: true } },
      { text: 'Weiterer Unterpunkt', options: { bullet: true, indentLevel: 1 } },
    ], { x: 0.5, y: 1.35, w: 5.4, h: 3.4, fontFace: FONT, fontSize: 16, color: C.text, valign: 'top', margin: 0, paraSpaceAfter: 8, isTextBox: true });
    s.addShape('roundRect', { x: 6.3, y: 1.35, w: 3.2, h: 3.4, fill: { color: C.grayBg }, line: { color: C.grayBg }, rectRadius: 0.12 });
    iconCircle(s, ic.shirt, 7.5, 2.15, 0.8, C.white);
    s.addText('Platz für Bild, Grafik oder Zitat', { x: 6.5, y: 3.15, w: 2.8, h: 0.6, fontFace: FONT, fontSize: 12, color: C.textMuted, align: 'center', margin: 0, isTextBox: true });
    s.addNotes('Standard-Inhaltsfolie: Aufzählung links, Bild-/Grafikfläche rechts. Der graue Kasten kann gelöscht und durch ein Bild ersetzt werden.');
  }

  // ======================================================================
  // Folie 10 – Abschluss
  // ======================================================================
  {
    const s = pres.addSlide({ masterName: 'MS_ABSCHLUSS' });
    s.addText('Vielen Dank für Ihr Interesse', { placeholder: 'title', isTextBox: true });
    s.addText('Wir freuen uns auf Ihre Anfrage.', { placeholder: 'sub', isTextBox: true });
    const contact = [
      [ic.mailL, 'info@maiershirts.de'],
      [ic.phoneL, '+49 (0) 000 000000'],
      [ic.globeL, 'www.maiershirts.de'],
    ];
    const totalW = 2.9 * contact.length, x0 = (10 - totalW) / 2;
    contact.forEach(([icon, label], i) => {
      const x = x0 + i * 2.9;
      iconCircle(s, icon, x, 4.0, 0.42, C.green);
      s.addText(label, { x: x + 0.52, y: 4.0, w: 2.3, h: 0.42, fontFace: FONT, fontSize: 12, color: C.light, valign: 'middle', margin: 0, isTextBox: true });
    });
    s.addNotes('Abschlussfolie (Master „MS_ABSCHLUSS“). Kontaktdaten sind Platzhalter – bitte durch die echten Daten ersetzen.');
  }

  // ---------- Schreiben ----------
  const pptxPath = path.join(DIST, 'Maiershirts_Master.pptx');
  await pres.writeFile({ fileName: pptxPath });
  console.log('geschrieben:', pptxPath);

  // .potx: identisches Paket, nur der Content-Type der Präsentation ist "template"
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'
  ));
  const potxPath = path.join(DIST, 'Maiershirts_Master.potx');
  fs.writeFileSync(potxPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('geschrieben:', potxPath);
}

main().catch((e) => { console.error(e); process.exit(1); });
