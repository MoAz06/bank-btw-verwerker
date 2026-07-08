# bank-btw-verwerker

Verwerkt een ING csv-export naar een overzichtelijk Excel-bestand met btw, inkomsten/uitgaven, en winst. Herkent facturen automatisch, filtert privé-transacties eruit, en kan je bonnen (PDF's) matchen tegen je bankmutaties zodat je ziet welke uitgaven nog geen bonnetje hebben en andersom.

## Wat het doet

- Leest een ING csv-export en bouwt een Excel met tabbladen: Transacties, Overzicht, Uitgaven, Inkomsten
- Rekent per transactie de btw uit (standaard 21%, zelf aan te passen per rij)
- Herkent of een inkomst een echte factuur is of een kredietfactuur (terugbetaling)
- Sluit privérekeningen en de eigen spaarrekening automatisch uit
- Vergelijkt een map met bonnen (PDF's, ook gescande foto's via OCR) tegen je uitgaven en laat zien wat wel/niet matcht

## Gebruiken

Start het venster:

```
python gui.py
```

Kies je ING csv-bestand, klik daarna eventueel op "Vergelijk met bonnenmap..." om je bonnen erbij te pakken.

Er is ook een los command-line script (`ing_btw_verwerker.py`) als je liever geen venster wil:

```
python ing_btw_verwerker.py mijn_export.csv
```

## Installeren

```
pip install -r requirements.txt
```

Voor OCR op gescande bonnen heb je daarnaast Tesseract nodig:

```
winget install --id UB-Mannheim.TesseractOCR -e
```

Zonder Tesseract werkt alles gewoon, alleen worden gescande (niet-digitale) bonnen dan overgeslagen met een melding.

## Instellingen

In `config.json` staan de dingen die je zelf kan aanpassen:

- `eigen_rekeningen`: IBAN's die als privé gelden (worden nooit meegeteld)
- `eigen_naam_patronen`: naamdelen die ook als privé gelden, bijvoorbeeld je eigen spaarrekening
- `nul_btw_patronen`: namen van partijen waar standaard 0% btw op zit (bijvoorbeeld de belastingdienst of bankkosten)

## Let op

De bank-csv en de Excel-output bevatten echte financiële gegevens. Zet die niet in git (staat al in `.gitignore`).
