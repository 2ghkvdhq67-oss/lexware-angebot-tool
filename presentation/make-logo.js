/**
 * Erzeugt eine Platzhalter-Wortmarke "MAIERSHIRTS" als PNG.
 *
 * Sobald das echte Logo vorliegt, einfach die Dateien
 *   assets/logo-dark.png   (Logo für helle Hintergründe)
 *   assets/logo-light.png  (Logo für dunkle Hintergründe)
 * durch die Originaldateien ersetzen und `npm run build` ausführen.
 */
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const { FaTshirt } = require('react-icons/fa');

const GREEN = '#22C55E';
const DARK = '#0B0F14';
const LIGHT = '#FFFFFF';

function logoSvg(textColor) {
  const icon = ReactDOMServer.renderToStaticMarkup(
    React.createElement(FaTshirt, { color: GREEN, size: 150 })
  );
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1100" height="200" viewBox="0 0 1100 200">
  <g transform="translate(20 25)">${icon}</g>
  <text x="200" y="128" font-family="Liberation Sans, Arial, Helvetica, sans-serif"
        font-size="118" font-weight="700" letter-spacing="4" fill="${textColor}">MAIER<tspan fill="${GREEN}">SHIRTS</tspan></text>
</svg>`;
}

async function write(name, textColor) {
  const out = path.join(__dirname, 'assets', name);
  await sharp(Buffer.from(logoSvg(textColor))).png().toFile(out);
  console.log('geschrieben:', out);
}

(async () => {
  await write('logo-dark.png', DARK);
  await write('logo-light.png', LIGHT);
})();
