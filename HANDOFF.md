# Handoff — Kwartaal (bank-btw-verwerker)

Doel van dit document: een volgende Claude-sessie kan het project hiermee volledig overnemen zonder de vorige gesprekgeschiedenis. Peildatum van deze notitie: augustus 2026. Werkmap: `C:\boekhouding` (voorheen `D:\silly`).

## 1. Wat is dit

**Kwartaal** is een lokale, gratis Windows-app voor een niet-technische Nederlandse installateur (Ostade Installatietechniek) die zelf de kwartaal-btw doet. Je leest een ING-bank-csv-export in; de app splitst btw/inkomsten/uitgaven, bouwt de aangifte-rubrieken (1a/1e/4b/5a/5b/5c), matcht bonnen aan uitgaven (OCR) en laat alles in de app zelf zien in tabs. Excel is nog maar één ding: een **Download**-knop (momentopname).

Harde eisen (blijven gelden):
- **Lokaal + gratis**: geen cloud, geen abonnement, geen AI-kosten. Data blijft op de pc.
- **Toegankelijkheid + leesbaarheid eerst**: dit is voor een echt persoon, niet voor techneuten. Grote tekst (A/A+/A++), licht/donker, uitleg van vaktermen, volledige toetsenbord-/schermlezer-ondersteuning.
- **Levering als 1 dubbelklik-.exe.**

Werkstijl-afspraken van de gebruiker (belangrijk):
- **De gebruiker commit zelf.** Geef alleen de git-commando's, voer ze niet uit.
- Commit-berichten: door mensen geschreven Nederlands, **geen** AI-tekst, **geen** `Co-Authored-By`, **geen** gedachtestreepjes (em-dash).
- `config.json` en `leveranciers.json` zijn lokaal en staan in `.gitignore`.
- Caveman-mode staat aan via een hook (compacte stijl); code/commits/security schrijf je gewoon normaal.

## 2. Snelstart

**Stand per 17 oktober 2026:** het project staat nu in `C:\boekhouding` (branch `reframe`); `D:\silly` bestaat niet meer. Op deze pc staat alleen **Python 3.13** (`C:\Users\moham\AppData\Local\Programs\Python\Python313`), de oude 3.11 is weg. Alle dependencies uit `requirements.txt` zijn op 3.13 geinstalleerd; de regressietest en de app draaien daarop.

```
py -3 app.py
```

Of dubbelklik `start.bat` (gebruikt `pyw -3`, geen zwart terminalvenster).

Exe bouwen (1 bestand, PyInstaller eerst installeren met `py -3 -m pip install pyinstaller`):
```
py -3 -m PyInstaller --noconfirm bank_btw_app.spec
```
Resultaat: `dist\Kwartaal.exe` (~54 MB, WebView2-backend). Vereist `webui\app-icon.ico` (bestaat).

Testbestanden (gitignored, buiten git):
- Q2: `C:\boekhouding\NL68INGB0675668204_01-04-2026_04-07-2026.csv` (staat in de projectmap)
- Q1: de csv `NL68INGB0675668204_01-01-2026_31-03-2026.csv` is op deze pc niet meer aanwezig; de golden ervoor staat nog in `_golden/`
- Bonnen (115 pdf's/foto's): stonden in `C:\Users\moham\Downloads\Bonnen test\Bonnen`, niet opnieuw gecontroleerd

## 3. Architectuur (3 lagen + app)

Bewust ontkoppeld zodat de cijfers in de app en in de Excel per definitie gelijk zijn:

1. **Rekenkern** `berekening.py` — pure functies, geen openpyxl/GUI. `excl_van`, `bepaal_btw`, `bepaal_aangifte_info`, `totalen_per_richting`. Constanten `STANDAARD_BTW=21`, `KOR_GRENS=20000`.
2. **Datamodel** `boekhouding.py` — de ENE bron van waarheid. `class Boekhouding` leest ruwe csv -> transacties, past overrides toe, herberekent alles bij een wijziging. `bereken_model(csv_pad, config=None, geheugen=None, overrides=None)`. Laadt zelf config/geheugen/overrides als je ze niet meegeeft.
3. **Presentatie/export** `ing_btw_verwerker.py` — bouwt de Excel (openpyxl). Sinds Fase 3 is de Excel een **momentopname met uitgerekende waarden** (geen formules, geen bladbeveiliging). `exporteer_excel(model, out, aangifte_override=None)`, `exporteer_csv(model, out)`, en de losse `verwerk_bestand(csv, out, overrides=None)` (command-line pad + regressietest).
4. **App** `app.py` (pywebview-venster) + `webmodel.py` (serializer model -> JSON) + `webui/` (index.html, style.css, app.js). `webmodel` importeert alleen uit `berekening` (niet uit openpyxl).

Overige modules: `leveranciers.py` (naam-normalisatie + btw-geheugen per leverancier), `bonnen_matcher.py` (OCR + matching), `boekhoud_export.py` (kale import-csv voor de boekhouder), `gui.py` (OUDE tkinter-stoplicht-GUI, nog aanwezig als terugval, niet meer de hoofdweg).

### Btw-precedentie (in `Boekhouding.bereken`)
`override (overrides.json) > verlegde btw > automatische 0% (config) > onthouden keuze (leveranciers.json) > standaard 21%`. Een handmatige btw-override zet `is_verlegd=False` en `btw_bron="override"`.

### Persistentie
- `overrides.json` **naast de app** = handmatige correcties per transactie (btw%, richting, investering), sleutel = `tx_id` (sha1 van datum+bedrag+tegenrekening+mededeling, 12 tekens, stabiel over kwartalen). Gitignored.
- `leveranciers.json` = onthouden btw per leverancier (met "onthoud voor deze leverancier" overschrijven). Gitignored.
- `overrides.json` **in de bonnenmap** = iets heel anders: handmatige bon-correcties (`genegeerd`/`forceer_nummer`), gelezen door `bonnen_matcher`.

## 4. De JS<->Python-brug (contract)

pywebview stelt de `Api`-methoden 1-op-1 bloot als `window.pywebview.api.<naam>()` (snake_case, geen camelCase). Alles retourneert `{ok: bool, ...}`; bij succes meestal `{ok:true, data:<model-of-bonnen-dict>}`, bij annuleren `{ok:false, geannuleerd:true}`, bij fout `{ok:false, fout:"nl-tekst", detail:"traceback"}`.

Door `webui/app.js` aangeroepen: `kies_csv`, `bevestig_csv`, `annuleer_csv`, `huidige_data`, `zet_btw`, `herstel_btw`, `zet_richting`, `kies_bonnen_map`, `download_excel`.

Import verloopt in twee stappen (preview -> bevestigen): `kies_csv`/`bereid_csv_voor` berekent een preview zonder het actieve kwartaal te vervangen; `bevestig_csv` maakt het actief (en schrijft dan pas het leveranciers-geheugen); `annuleer_csv` gooit de preview weg. `laad(pad)` = preview + bevestig ineen (voor een start-csv van de commandoregel).

`model_naar_dict(model)` levert o.a.: `transacties[]` (met `id, datum(iso), naam, leverancier, factuur, richting, btw_pct, btw_bron, bedrag, excl, btw, is_investering, is_verlegd, is_kredietfactuur, auto_nul`), `inkomsten[]`, `uitgaven[]`, `investeringen[]`, `leveranciers{rijen,afwijkingen}`, `totalen{...}`, `winst`, `btw_saldo`, `aangifte{...rubrieken...}`, `buiten_periode[]`, `tellingen{...}`. Alles JSON-serialiseerbaar (datums als "YYYY-MM-DD", geen datetime/set/Path). `bonnen_naar_dict(analyse)` levert `totaal_bonnen, btw_risico, tellingen, zeker[], onzeker[], geen_match[], onleesbaar[], zonder_bon[]`.

## 5. Fases (roadmap) — allemaal gedaan

- **0** golden regressie-vangnet (`regressietest.py`)
- **1** pure rekenkern `berekening.py`
- **2** datamodel `boekhouding.py`, `bereken_model`
- **3** Excel wordt momentopname met waarden (formules/beveiliging weg)
- **4** overrides.json-persistentie + geheugen-overschrijven, precedentie
- **5** pywebview-app-schil, 8 tabs, Download Excel, toegankelijk
- **6** live btw% bewerken -> alle tabs rekenen mee (herstel-knop, onthoud-per-leverancier)
- **7** Bonnencontrole-tab (OCR), uitleg-termen, licht/donker, A/A+/A++, plain errors, exe

Sinds fase 7 heeft de gebruiker zelf uitgebreid: app hernoemd naar **Kwartaal** (+ icoon `webui/app-icon.*`), import-preview/bevestig-flow, `zet_richting` (richting per transactie omzetten in de UI), bonnen-analyse wordt in de gedownloade Excel meegeschreven (`bouw_bonnencontrole_sheet`), en `exporteer_excel(..., aangifte_override=...)` zodat de app-teksten (periode/deadline) in de Excel komen.

## 6. Verifiëren (doe dit na elke wijziging)

Regressie (vergelijkt door-Python-berekende cijfers met de golden in `_golden/`, gitignored):
```
py -3 regressietest.py check "NL68INGB0675668204_01-04-2026_04-07-2026.csv"
```
Verwacht: Q2 116 tx, winst 1346.19, saldo -5978.07 (bevestigd op Python 3.13, 17-10-2026). Q1 (120 tx, winst 62147.13, saldo 3819.36) kan pas weer als de Q1-csv terug is.

**Let op:** de regressietest verwijdert `leveranciers.json` in de projectmap (voor een schone run). Bestaat dat bestand met echte keuzes, maak dan eerst een kopie. Bij bewuste logica-wijziging herijk je met `regressietest.py opslaan ...`.

Let op: de regressietest leest **modelgetallen**, niet Excel-celwaarden. Na een wijziging aan de Excel-bouwers apart controleren dat celwaarden == model (zie scratchpad-scripts hieronder).

Losse test-/preview-scripts staan in de scratchpad van de vorige sessie (niet in de repo, mogelijk al opgeruimd):
een scratchpad van een vorige sessie op de oude pc (niet meer aanwezig)
— `spotcheck_excel.py` (Excel-cel == model), `fase4_test.py` (overrides/persistentie/precedentie), `fase6_test.py` (live bewerken via Api), `fase7_bonnen.py` (bonnen end-to-end). Zijn makkelijk te herschrijven als ze weg zijn; het patroon is: `sys.path.insert(0,"C:/boekhouding")`, model bouwen met `bereken_model(csv, overrides={})`, assert op de uitkomsten.

UI kan je niet met de ingebouwde Browser via `file://` bekijken (geblokkeerd) en screenshots time-outen in deze omgeving. Werkwijze die wél werkte: `py -3 -m http.server 8731` in `webui/` draaien, een `_preview.html` genereren (index.html met `window.__DEMO_MODEL__ = <json>` geïnjecteerd vóór `app.js`), en dan met de Browser-tools `read_page`/`javascript_tool` de DOM inspecteren. `_preview.html` is gitignored. app.js heeft een demo-fallback: zonder `window.pywebview` toont het `window.__DEMO_MODEL__`.

De exe smoke-testen (headless): starten en na ~9s checken dat het proces nog leeft = imports + venster-init geslaagd.

## 7. Bekende valkuilen

- **Verkeerde Python.** Vroeger stond 3.14 zonder packages naast 3.11. Nu is er alleen 3.13 met alle packages; `py -3` / `start.bat` volstaat. Komt er weer een tweede Python bij, check dan eerst `py -0`.
- **pywebview file-filter-format.** `create_file_dialog(file_types=...)` labels mogen vóór de haakjes ALLEEN `[\w ]` zijn. Een koppelteken (bv. `"ING CSV-export (*.csv)"`) gooit `ValueError: ... is not a valid file filter` en de knop lijkt stuk. Gebruik `"ING CSV bestand (*.csv)"`, `"Excel bestand (*.xlsx)"`, `"Alle bestanden (*.*)"`.
- **Dialoog-constanten** `webview.OPEN_DIALOG/SAVE_DIALOG/FOLDER_DIALOG` (10/30/20) werken maar zijn deprecated in pywebview 6.2.1 (nieuw: `FileDialog.OPEN/SAVE/FOLDER`). Waarschuwingen zijn onschuldig en onzichtbaar via `pyw`/de exe.
- **Exe-naam mismatch.** Huidige `bank_btw_app.spec` bouwt `dist\Kwartaal.exe`, maar in `dist\` staat nog de oude `Bank-btw-verwerker.exe` (van vóór de rename). Even opnieuw bouwen om `dist\Kwartaal.exe` te krijgen.
- **HEIC-bonnen** (iPhone) werken alleen met `pillow-heif`; die is NIET geïnstalleerd en zit niet in de exe. JPG/PNG/PDF gaan gewoon. De import is lazy in `bonnen_matcher`.
- **Bonnen-snapshot veroudert.** Na een btw-wijziging wordt `bonnen_analyse` in `app.py` op `None` gezet (het btw-risico klopt anders niet meer); de gebruiker moet de bonnen opnieuw controleren. Bewust.
- **pdfplumber FontBBox-waarschuwingen** bij het lezen van sommige pdf's zijn onschuldig.

## 8. Git-staat + voorgestelde commits

Alles staat nog **ongecommit** (fases 0 t/m 7 zaten nooit in git; laatste commit is `a204aa1`). Gewijzigd: `.gitignore`, `README.md`, `config.example.json`, `ing_btw_verwerker.py`, `leveranciers.py`, `requirements.txt`. Nieuw: `app.py`, `webmodel.py`, `berekening.py`, `boekhouding.py`, `regressietest.py`, `start.bat`, `bank_btw_app.spec`, `webui/`, en dit `HANDOFF.md`. Gitignored blijven: `config.json`, `leveranciers.json`, `overrides.json`, `_golden/`, `dist/`, `build/`, `*.exe`, `webui/_preview.html`, `*.csv/*.xlsx/*.pdf`. De spec is expliciet wél getrackt via `!bank_btw_app.spec`.

De gebruiker commit zelf. Zinvolle groepering (hele bestanden, geen hunk-splitsing nodig):
```
git add berekening.py boekhouding.py leveranciers.py regressietest.py
git commit -m "rekenkern, datamodel en overrides-persistentie als aparte lagen met regressie-vangnet"

git add ing_btw_verwerker.py
git commit -m "excel wordt een vaste momentopname met uitgerekende waarden in plaats van formules"

git add app.py webmodel.py webui start.bat bank_btw_app.spec
git commit -m "kwartaal-app met tabs in webview: csv inlezen met controle, btw en richting live bewerken, bonnencontrole en excel-download"

git add .gitignore README.md requirements.txt config.example.json HANDOFF.md
git commit -m "readme, requirements, gitignore en handoff bijgewerkt voor de app en de exe-build"
```

## 9. Mogelijke vervolgstappen

- `dist\Kwartaal.exe` opnieuw bouwen met de huidige spec (oude exe-naam ligt er nog).
- Beslissen of `gui.py` (oude tkinter-versie) weg mag of als terugval blijft.
- Verlegde-btw-precedentie is in code geregeld maar niet met echte verlegde transacties in de testkwartalen bevestigd (Q1/Q2 hebben er geen); als er een kwartaal mét verlegde btw komt, even natrekken.
- Overweeg de dialoog-constanten naar `webview.FileDialog.*` te tillen als je pywebview vastzet op 6.x.
- Golden-referentie staat lokaal in `_golden/`; op een andere pc moet je die eerst met `regressietest.py opslaan` opnieuw genereren vanaf vertrouwde csv's.

## 10. Persistent geheugen

In `C:\Users\moham\.claude\projects\C--boekhouding\memory\` staat `MEMORY.md` (index). De oude `D--silly`-map met `pywebview-file-filter-format.md` is weg; de valkuil staat hierboven in paragraaf 7.
