/**
 * Maiershirts-Logo als Vektor (Bergmarke + Wortmarke), nachgebaut nach der
 * Originalvorlage. Erzeugt:
 *   assets/logo.svg         Vektorquelle (schwarz)
 *   assets/logo-dark.png    schwarze Version für helle Hintergründe
 *   assets/logo-light.png   weiße Version für dunkle Hintergründe
 *
 * Liegt das Original-Logo als Datei vor, einfach logo-dark.png / logo-light.png
 * überschreiben und `npm run build` ausführen.
 */
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const BLACK = '#0F0F0F';
const WHITE = '#FFFFFF';

// Ein "Λ" als Polygon: Spitze bei (ax, top), Füße auf y = base, Halbspannweite s,
// horizontale Schenkelstärke t. Die innere Spitze liegt um t*(Höhe/s) tiefer.
function peak(ax, top, base, s, t) {
  const innerTop = top + t * ((base - top) / s);
  return [
    [ax - s, base], [ax, top], [ax + s, base],
    [ax + s - t, base], [ax, innerTop], [ax - s + t, base],
  ].map((p) => p.join(',')).join(' ');
}
// Rechter Schenkel des linken Λ (für die Aussparung am Kreuzungspunkt)
function rightLeg(ax, top, base, s, t) {
  const innerTop = top + t * ((base - top) / s);
  return [[ax, top], [ax + s, base], [ax + s - t, base], [ax, innerTop]].map((p) => p.join(',')).join(' ');
}

function logoSvg(color) {
  const top = 6, base = 208, s = 98, t = 36;
  const L = peak(246, top, base, s, t);
  const R = peak(346, top, base, s, t);
  const leg = rightLeg(246, top, base, s, t);
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="600" height="330" viewBox="0 0 600 330">
  <defs>
    <mask id="cut">
      <rect width="600" height="330" fill="white"/>
      <polygon points="${leg}" fill="black" stroke="black" stroke-width="14" stroke-linejoin="round"/>
    </mask>
  </defs>
  <polygon points="${R}" fill="${color}" mask="url(#cut)"/>
  <polygon points="${L}" fill="${color}"/>
  <text x="300" y="312" text-anchor="middle" font-family="Montserrat, Arial, sans-serif"
        font-weight="700" font-size="76" letter-spacing="3" fill="${color}">MAIERSHIRTS</text>
</svg>`;
}

async function write(name, color) {
  const out = path.join(__dirname, 'assets', name);
  await sharp(Buffer.from(logoSvg(color)), { density: 288 }).png().toFile(out);
  console.log('geschrieben:', out);
}

(async () => {
  fs.writeFileSync(path.join(__dirname, 'assets', 'logo.svg'), logoSvg(BLACK));
  await write('logo-dark.png', BLACK);
  await write('logo-light.png', WHITE);
})();
