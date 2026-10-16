# Kwartaal

Een **app** die je bank-csv-export (ING) verwerkt tot een **aangifte-klaar** overzicht: btw per rubriek, winst, en de controle of elke uitgave een bon heeft. Alles gebeurt **in de app zelf** — tabs voor Transacties, Overzicht, BTW-aangifte enzovoort, waar je de btw per transactie live kan aanpassen. De Excel is nog maar één knop: **Download Excel**. Alles draait **lokaal en gratis** — geen cloud, geen abonnement, geen AI-kosten. Je financiële gegevens blijven op je eigen pc.

## Wat het doet

- Leest een bank-csv en toont alles **in de app** in tabs: **Transacties, Overzicht, BTW-aangifte, Uitgaven, Inkomsten, Investeringen, Leveranciers, Bonnencontrole**.
- **Live bewerken**: pas de btw van een transactie aan via een keuzelijstje; alle andere tabs en de aangifte rekenen meteen mee. Correcties worden onthouden (`overrides.json`), eventueel ook per leverancier voor volgende kwartalen.
- **Toegankelijk gebouwd**: grote leesbare tekst (knoppen A / A+ / A++), licht/donker, uitleg van de vaktermen, en volledige toetsenbord-/schermlezer-ondersteuning.
- **Download Excel**: dezelfde cijfers als vaste momentopname in een Excel met tabbladen (voor je boekhouder of je archief).
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

Start de app (dubbelklik `start.bat`, of via de terminal):

```
py -3 app.py
```

> Let op: gebruik de Python waar de vereisten op staan. Heb je meerdere Python-versies, dan draait `python app.py` mogelijk de verkeerde (zonder `pywebview`) — `py -3 app.py` of `start.bat` gebruikt de standaard Python 3 van de py-launcher; controleer met `py -0` welke dat is.

Klik **Kies bankbestand…** en selecteer je ING-csv-export. Je ziet meteen de kerncijfers en alle tabs. Pas eventueel btw-percentages aan (ze rekenen live door), controleer je bonnen via de tab **Bonnencontrole**, en klik **Download Excel** als je een bestand wil bewaren.

Als losse `.exe` (dubbelklikken, geen Python nodig) bouw je met:

```
python -m PyInstaller --noconfirm bank_btw_app.spec
```

Het resultaat staat in `dist/Kwartaal.exe`.

Ook nog beschikbaar als los command-line script dat direct een Excel schrijft:

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

Je btw-correcties uit de app worden opgeslagen in `overrides.json` **naast de app** (per transactie). Dat is een ander bestand dan de handmatige **bon**-correcties, die in `overrides.json` **in de bonnenmap** staan (`{"bon.pdf": {"genegeerd": true}}` of `{"forceer_nummer": "12345"}`).

## Let op

De bank-csv, de Excel-output, `config.json` en `leveranciers.json` bevatten echte financiële gegevens. Die staan in `.gitignore` en horen niet in git. Verlegde btw, investeringen/afschrijving en KOR zijn fiscaal maatwerk — de app geeft een goede eerste schatting, maar controleer de bijzondere gevallen met je boekhouder.
