# /pages/fuer-firmen — Hub-and-Spoke Split (C-lite)

Vier HTML-Bodies für Shopify Pages. **Noch nicht deployed** — Freigabe abwarten.

## Dateien → Shopify Pages

| Datei | Handle | Page-ID | Aktion |
|---|---|---|---|
| `fuer-firmen.html` | `fuer-firmen` | `gid://shopify/Page/165736775944` | pageUpdate (entkernt → Hub) |
| `arbeitskleidung-bedrucken.html` | `arbeitskleidung-bedrucken` | `gid://shopify/Page/167150715144` | pageUpdate (Vollausbau) |
| `corporate-fashion.html` | `corporate-fashion` | — | pageCreate |
| `event-merch.html` | `event-merch` | — | pageCreate |

## Titles & Meta

### fuer-firmen (Hub)
- **title:** Textildruck für Firmen — Workwear, Corporate Fashion & Event-Merch
- **title_tag:** `Textildruck für Firmen – Workwear, Corporate Fashion & Merch | Maiershirts`
- **description_tag:** `Firmenkleidung mit Logo aus Ammerbuch bei Tübingen: Arbeitskleidung, Corporate Fashion und Event-Merch. Ab 1 Stück, Kauf auf Rechnung, fester Ansprechpartner.`
- ⚠️ **Wichtig:** Der bisherige title_tag lautete „Arbeitskleidung & Workwear bedrucken – ab 1 Stück | Maiershirts Tübingen" — exakt das Hauptkeyword der Spoke-Seite. Diese Kannibalisierung ist der Hauptgrund für Pos 38 (Hub) / Pos 41 (Spoke). Der neue title_tag muss zwingend mitdeployed werden, sonst verpufft der Split.

### arbeitskleidung-bedrucken (Spoke 1, Priorität)
- **title:** Arbeitskleidung bedrucken lassen — Workwear & Warnwesten mit Logo
- **title_tag:** `Arbeitskleidung bedrucken lassen – Workwear mit Logo | Maiershirts Tübingen`
- **description_tag:** `Arbeitskleidung bedrucken und beschriften lassen: Logo auf Softshell, Bundhose, Kasack & Warnweste. Bis 90 °C und chemische Reinigung beständig. Ab 1 Stück, aus Ammerbuch bei Tübingen.`

### corporate-fashion (Spoke 2, neu)
- **title:** Corporate Fashion & Teamwear mit Logo
- **title_tag:** `Corporate Fashion mit Logo – Firmenkleidung & Teamwear | Maiershirts`
- **description_tag:** `Corporate Fashion in Ihren CI-Farben: Polos, Hemden, Softshells & Strick mit Logo. Musterteile vorab, Damen- & Herrenschnitte, Größen bis 5XL, Nachbestellung jederzeit.`

### event-merch (Spoke 3, neu, minimal)
- **title:** Event-Merch mit Logo — Shirts, Caps & Taschen
- **title_tag:** `Event-Merch bedrucken – Shirts, Caps & Taschen mit Logo | Maiershirts`
- **description_tag:** `Merch für Firmenjubiläum, Messe, Betriebsausflug und Vereinsfest. Ab 1 Stück, verbindlicher Liefertermin zum Eventdatum. Produziert in Ammerbuch bei Tübingen.`

## Getroffene Entscheidungen (Freigabe Stefan)
1. **Referenzen:** Branche + Region, keine Namen/Logos. Echte Kundennamen in Schritt 2 nachziehen.
2. **Fotos:** Detail-Ausschnitte. Für Arbeitskleidung existieren noch keine eigenen Aufnahmen → 2 Bildplätze als HTML-Kommentar markiert (`<!-- BILDPLATZ n -->`), text-first live, Fotos später nachsetzen. Corporate Fashion nutzt vorhandene Produktbilder aus Shopify Files.
3. **Corporate Fashion:** Business-Casual (Polos, Hemden, Softshell, Fleece, Sweats, Strick/Half-Zip, Caps). Kein Business-Formal — explizit als bewusste Abgrenzung auf der Seite formuliert.
4. **Preise:** keine Zahlen. FAQ beantwortet die Preisfrage über den Beratungsweg, statt sie zu ignorieren (wichtig für LLM-Antworten).
5. **CTA:** durchgehend Kontaktformular `/pages/contact`, kein WhatsApp/Calendly.

## Compliance-Checks (bestanden)
- SEO-Policy: kein DTF, Siebdruck, Stickerei, Flock/Flex. Nur Textildruck + Transferveredelung (Supacolour Industrial Wash, Avery Dennison).
- Keine Preisangaben in allen 4 Bodies.
- Alle 8 JSON-LD-Blöcke (4× FAQPage, 4× LocalBusiness) valides JSON.
- Interne Verlinkung: Hub ↔ alle 3 Spokes, Spokes untereinander, plus /pages/textildruck, /pages/textilien, /pages/textildruck-tuebingen.

## Deploy (nach Freigabe)

Bestehende Seiten:
```graphql
mutation UpdatePage($id: ID!, $page: PageUpdateInput!) {
  pageUpdate(id: $id, page: $page) { page { id handle title } userErrors { field message } }
}
```
Neue Seiten:
```graphql
mutation CreatePage($page: PageCreateInput!) {
  pageCreate(page: $page) { page { id handle title } userErrors { field message } }
}
```
Metafields (`global.title_tag`, `global.description_tag`, Typ `single_line_text_field`) im selben Input mitgeben.

Danach Live-Verify per curl auf allen 4 URLs (Statuscode + Stichprobe auf H2-Überschriften und title_tag).

## Offene Punkte für Schritt 2
- 2 Handy-Nahaufnahmen für die Bildplätze auf `/pages/arbeitskleidung-bedrucken` (Logo auf Softshell, Rückenbeschriftung auf Arbeitsjacke).
- Echte Kundenreferenzen mit Freigabe für Namen/Logo.
- Eigene lokale Landings `/pages/arbeitskleidung-tuebingen` und `/pages/arbeitskleidung-reutlingen` — beide Queries stehen bereits auf Pos 7,5 bzw. 9,8. Im aktuellen Sprint als Regionalabschnitt auf der Spoke abgebildet; separate Landings erst bauen, wenn der Split gewirkt hat, sonst kannibalisiert es erneut.
