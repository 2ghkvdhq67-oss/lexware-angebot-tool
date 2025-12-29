// --- Preislogik (AUTOFILL aus Artikelstamm, falls Excel leer) ---

// Helper: unitPrice aus Artikel ziehen und passend zu taxType bauen
async function buildUnitPriceFromArticle(articleId, taxType) {
  const art = await getArticleById(articleId);
  const p = art?.price || null;
  if (!p) return null;

  // Lexware Artikelpreis kann netPrice/grossPrice/leadingPrice/taxRate enthalten
  const taxRate = Number.isFinite(Number(p.taxRate)) ? Number(p.taxRate) : 19;

  if (taxType === 'gross') {
    const gross = Number.isFinite(Number(p.grossPrice))
      ? Number(p.grossPrice)
      : (Number.isFinite(Number(p.netPrice)) ? Number(p.netPrice) * (1 + taxRate / 100) : null);

    if (gross == null) return null;

    return { currency: 'EUR', grossAmount: Number(gross.toFixed(2)), taxRatePercentage: taxRate };
  }

  // net
  const net = Number.isFinite(Number(p.netPrice))
    ? Number(p.netPrice)
    : (Number.isFinite(Number(p.grossPrice)) ? Number(p.grossPrice) / (1 + taxRate / 100) : null);

  if (net == null) return null;

  return { currency: 'EUR', netAmount: Number(net.toFixed(2)), taxRatePercentage: taxRate };
}

// ... innerhalb deiner Positions-Schleife, nachdem item erstellt ist:

// 1) Wenn Excel Preis vorhanden -> wie bisher setzen (net/gross)
if (unitPriceAmount !== null) {
  const rate = taxRatePercentage !== null ? taxRatePercentage : 19;

  if (taxType === 'gross') {
    item.unitPrice = { currency: 'EUR', grossAmount: unitPriceAmount, taxRatePercentage: rate };
  } else {
    item.unitPrice = { currency: 'EUR', netAmount: unitPriceAmount, taxRatePercentage: rate };
  }
} else {
  // 2) Excel Preis leer -> AUTOFILL
  if (articleId) {
    // Preis aus Artikelstamm holen (egal ob material/service/custom)
    const autoPrice = await buildUnitPriceFromArticle(articleId, taxType);

    if (autoPrice) {
      item.unitPrice = autoPrice;
      warnings.push({
        sheet: 'Positionen',
        row: excelRow,
        message: `Preis war leer → automatisch aus Artikelstamm ergänzt (articleId=${articleId}).`
      });
    } else {
      // Artikel hat keinen Preis -> Fehler (sonst wieder 406)
      errors.push({
        sheet: 'Positionen',
        row: excelRow,
        field: 'unitPriceAmount',
        message: `Preis fehlt in Excel und konnte nicht aus Artikelstamm ermittelt werden (articleId=${articleId}).`
      });
      continue;
    }
  } else {
    // kein articleId und kein Preis -> wie bisher Fehler
    errors.push({
      sheet: 'Positionen',
      row: excelRow,
      field: 'unitPriceAmount',
      message: 'Preis ist Pflicht, wenn keine articleId gesetzt ist.'
    });
    continue;
  }
}
