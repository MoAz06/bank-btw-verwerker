# bank-btw-verwerker

Verwerkt een bank-csv-export (ING) naar een overzichtelijk, **aangifte-klaar** Excel-bestand: btw per rubriek, winst, en de controle of elke uitgave een bon heeft. Alles draait **lokaal en gratis** — geen cloud, geen abonnement, geen AI-kosten. Je financiële gegevens blijven op je eigen pc.

## Wat het doet

- Leest een bank-csv en bouwt een Excel met tabbladen: **Transacties, Overzicht, BTW-aangifte, Uitgaven, Inkomsten, Investeringen, Leveranciers** (+ **Bonnencontrole** als je bonnen vergelijkt).
- **BTW-aangifte-tab**: de cijfers die je letterlijk overtikt in Mijn Belastingdienst Zakelijk — rubriek 1a / 5a / 5b / 5c, live gekoppeld aan je btw-percentages. Plus een "zet dit opzij voor de aangifte"-potje met deadline.
- Rekent btw per transactie (standaard 21%, per rij aan te passen), splitst incl./excl./btw-bedrag, en berekent winst + btw-saldo.
- Herkent **echte klantfacturen** (jaarmaand-factuurnummer binnen de periode) vs. **kredietfacturen** (terugbetaling → telt als uitgave).
- Sluit privérekeningen, de eigen spaarrekening, de Belastingdienst en **leningen/aflossingen** automatisch uit.
- **Verlegde btw** (onderaanneming bouw): aparte rubriek 1e/4b met een controleer-notitie.
- **Investeringen**: grote uitgaven worden als kandidaat apart gezet (bus, machine, gereedschap) met KIA-signaal.
- **Leveranciers-overzicht**: uitgaven per genormaliseerde leverancier + een lijst opvallende bedragen.
- **Leveranciers-geheugen** (`leveranciers.json`): onthoudt je btw-keuze per leverancier voor volgende kwartalen.
- **Bonnen-matching** (PDF's én foto's, met OCR): koppelt bonnen aan bankuitgaven en laat zien wat wel/niet matcht — inclusief een **btw-risico** (voorbelasting zonder factuur).
- **Boekhouder-export**: een kale import-CSV (`..._boekhouder_import.csv`) die je boekhouder in zijn pakket kan inlezen.

## Gebruiken

Start het venster:

```
python gui.py
```

Kies je bank-csv → je krijgt een **stoplicht-eindscherm** (winst, btw, en wat je nog moet controleren). Klik daarna eventueel op "Vergelijk met bonnen" en kies de map met je bon-PDF's of foto's.

Ook als los command-line script:

```
python ing_btw_verwerker.py mijn_export.csv
```

## Installeren

```
pip install -r requirements.txt
```

Voor OCR op gescande/foto-bonnen heb je Tesseract nodig (eenmalig):

```
winget install --id UB-Mannheim.TesseractOCR -e
```

Zonder Tesseract werkt alles gewoon; alleen niet-digitale bonnen worden dan overgeslagen met een melding. Voor betere Nederlandse OCR kan je `nld.traineddata` in de `tessdata`-map van Tesseract zetten. iPhone-foto's (HEIC)? Dan ook `pip install pillow-heif`.

## Instellingen

Alles in `config.json` (zie `config.example.json` voor uitleg per sleutel). O.a.: eigen (privé)rekeningen, 0%-btw-partijen, verlegde btw, uitsluit-woorden (lening/aflossing), investering-drempel, leverancier-aliassen, en de KOR-schakelaar. Ontbrekende sleutels krijgen een verstandige standaard.

Handmatige bon-correcties zet je in `overrides.json` in de bonnenmap (`{"bon.pdf": {"genegeerd": true}}` of `{"forceer_nummer": "12345"}`).

## Let op

De bank-csv, de Excel-output, `config.json` en `leveranciers.json` bevatten echte financiële gegevens. Die staan in `.gitignore` en horen niet in git. Verlegde btw, investeringen/afschrijving en KOR zijn fiscaal maatwerk — de app geeft een goede eerste schatting, maar controleer de bijzondere gevallen met je boekhouder.
