'use strict';

/**
 * Parse-Fokus: Kundenanfragen + Lieferschein-Text, der aus Mail/PDF kopiert wird.
 * Ziel: möglichst oft Menge + Artikeltext erkennen, robust gegen Tabs/Spaces/Bullets.
 *
 * Output items:
 *  - type: custom | material | service | text
 *  - quantity (number|null)
 *  - unitName (string|null)
 *  - name (string)
 *  - description (string|null)
 *  - articleId (uuid|null)  // falls direkt enthalten
 *  - articleNumber (string|null) // SKU/ArtNr falls erkannt
 *  - size/color (optional hints)
 */

const UUID_RE =
  /\b[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}\b/i;

const SIZE_RE =
  /\b(XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|XXXXL|5XL|4XL|3XL|2XL|\d{2,3})\b/i;

const MONEY_RE =
  /(\d{1,4}(?:[.,]\d{1,2})?)\s?(€|eur)\b/i;

const ARTNR_RE =
  /\b(?:art(?:ikel)?\.?\s?(?:nr\.?|no\.?|number)?|sku|ean|gtin|#)\s*[:=]?\s*([A-Z0-9][A-Z0-9\-_.\/]{2,})\b/i;

const QTY_RE_LIST = [
  // "25 Stk", "25 Stück", "25 pcs", "25x"
  /^\s*(\d+(?:[.,]\d+)?)\s*(x|stk|stck|stück|stueck|pcs|pc|pieces|piece|units|unit)\b/i,
  // "Menge: 25"
  /\b(?:menge|qty|quantity|anzahl)\s*[:=]\s*(\d+(?:[.,]\d+)?)\b/i,
  // "... 25 Stk" (am Ende)
  /\b(\d+(?:[.,]\d+)?)\s*(stk|stück|stueck|pcs|pc|x)\s*$/i,
  // Tabellen: "25\tE190\t..."
  /^\s*(\d+(?:[.,]\d+)?)\s*\t/i
];

function cleanLine(line) {
  let s = String(line || '').trim();
  // remove common bullets/prefixes
  s = s.replace(/^[\-\*\u2022•\u25CF]+\s*/, '');
  // collapse spaces (but keep tabs for table detection)
  s = s.replace(/[ ]{2,}/g, ' ');
  return s;
}

function parseNumber(n) {
  if (n == null) return null;
  const s = String(n).trim().replace(',', '.');
  const val = Number(s);
  return Number.isFinite(val) ? val : null;
}

function guessType(line) {
  const l = line.toLowerCase();

  // hint/text lines
  if (/(lieferzeit|hinweis|bemerkung|note|info|achtung|bitte|wichtig|druckfreigabe|freigabe|required)/i.test(l)) {
    return 'text';
  }

  // material hints (rohware)
  if (/(t-?shirt|shirt|hoodie|sweat|zip|jacke|softshell|cap|mütze|beanie|hose|textil|rohware|polo|b&c|gildan|fruit|stanley|stella)/i.test(l)) {
    return 'material';
  }

  // service hints
  if (/(design|grafik|layout|setup|einrichtung|digitalisierung|stickprogramm|vektor|freistellung|versand|porto|shipping)/i.test(l)) {
    return 'service';
  }

  // default
  return 'custom';
}

function extractQtyAndUnit(line) {
  let qty = null;
  let unit = null;

  for (const re of QTY_RE_LIST) {
    const m = line.match(re);
    if (!m) continue;

    // first group should be qty
    qty = parseNumber(m[1]);
    if (qty == null) continue;

    // unit
    if (m[2]) {
      const u = String(m[2]).toLowerCase();
      unit =
        u === 'x' ? 'Stk' :
        (u === 'pcs' || u === 'pc' || u === 'piece' || u === 'pieces') ? 'Stk' :
        (u === 'stk' || u === 'stck' || u === 'stück' || u === 'stueck') ? 'Stk' :
        'Stk';
    } else {
      unit = 'Stk';
    }

    // remove matched qty/unit from line (only for certain patterns)
    if (re.source.startsWith('^')) {
      line = line.replace(re, '').trim();
    }
    return { qty, unit, rest: line.trim() };
  }

  return { qty: null, unit: null, rest: line.trim() };
}

function extractArticleId(line) {
  const m = line.match(UUID_RE);
  return m ? m[0] : null;
}

function extractArticleNumber(line) {
  const m = line.match(ARTNR_RE);
  return m ? m[1] : null;
}

function extractSize(line) {
  const m = line.match(SIZE_RE);
  return m ? m[1] : null;
}

function extractPriceHint(line) {
  const m = line.match(MONEY_RE);
  if (!m) return null;
  return m[1].replace(',', '.');
}

function stripNoiseForName(line) {
  let s = line;

  // strip price mentions
  s = s.replace(MONEY_RE, '').trim();

  // strip obvious labels
  s = s.replace(/\b(?:menge|qty|quantity|anzahl|art(?:ikel)?\.?\s?(?:nr\.?|no\.?|number)?|sku|ean|gtin)\b\s*[:=]?\s*/ig, '');

  // collapse spaces
  s = s.replace(/[ ]{2,}/g, ' ').trim();
  return s;
}

function parseLines(text) {
  const raw = String(text || '');
  const lines = raw
    .split(/\r?\n/)
    .map(cleanLine)
    .filter(Boolean);

  const items = [];
  const warnings = [];

  for (let idx = 0; idx < lines.length; idx++) {
    const original = lines[idx];

    // Try tabular split (common from PDFs/Excel copy)
    // e.g. "25\tE190\tweiß\tL\t..."
    const tabParts = original.split('\t').map(s => s.trim()).filter(Boolean);

    let work = original;
    if (tabParts.length >= 2) {
      // If first cell is qty, rebuild line: "qty rest..."
      const q = parseNumber(tabParts[0]);
      if (q != null) {
        work = `${tabParts[0]} Stk ${tabParts.slice(1).join(' ')}`;
      }
    }

    const articleId = extractArticleId(work);
    const articleNumber = extractArticleNumber(work);
    const size = extractSize(work);
    const priceHint = extractPriceHint(work);

    const type = guessType(work);

    const { qty, unit, rest } = extractQtyAndUnit(work);

    let name = stripNoiseForName(rest || work);
    name = name || original;

    // If we detected "text" type, keep entire line as description if name is too short
    let description = null;

    if (type === 'text') {
      // prefer a short title
      if (name.length > 40) {
        description = name;
        name = 'Hinweis';
      } else {
        description = null;
      }
    } else {
      // For non-text: if price hint exists, keep it in description (but don't parse into unitPrice here)
      const hints = [];
      if (articleNumber) hints.push(`ArtNr/SKU: ${articleNumber}`);
      if (size) hints.push(`Größe: ${size}`);
      if (priceHint) hints.push(`Preis-Hinweis: ${priceHint} EUR`);
      if (hints.length) description = hints.join(' • ');
    }

    // If qty missing for non-text -> warn and skip OR treat as text?
    if ((type !== 'text') && (qty == null)) {
      warnings.push({
        line: idx + 1,
        message: `Keine Menge erkannt → als Hinweiszeile übernommen: "${original}"`
      });
      items.push({
        type: 'text',
        quantity: null,
        unitName: null,
        name: 'Hinweis',
        description: original,
        articleId: articleId || null,
        articleNumber: articleNumber || null
      });
      continue;
    }

    items.push({
      type,
      quantity: qty,
      unitName: type === 'text' ? null : (unit || 'Stk'),
      name: name,
      description,
      articleId: articleId || null,
      articleNumber: articleNumber || null
    });
  }

  return { items, warnings };
}

module.exports = { parseLines };
