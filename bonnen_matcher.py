#!/usr/bin/env python3
"""Vergelijkt bonnen (PDF's) met de uitgaven uit een verwerkte ING csv."""
import difflib
import json
import re
import shutil
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path

import pdfplumber

MAAND_NAAM_NAAR_NUMMER = {
    "januari": 1, "februari": 2, "maart": 3, "april": 4, "mei": 5, "juni": 6,
    "juli": 7, "augustus": 8, "september": 9, "oktober": 10, "november": 11, "december": 12,
}

LABEL_NUMMER_PATROON = re.compile(
    r"(?:Factuurnummer|Factuurnr\.?|Factuur\s*nr\.?|Contractnummer|Aanslagnummer|"
    r"Bestelnummer|Bestelling\s*#|Ordernummer|Order\s*nr\.?)\s*[:\s#]*\s*([A-Za-z0-9][A-Za-z0-9\-\.]{3,20})",
    re.IGNORECASE,
)
KAAL_NUMMER_PATROON = re.compile(r"\b\d{5,15}\b")

# Voorkeur voor een bedrag dat expliciet "incl. btw" vermeldt (het echte totaalbedrag) boven
# een kale "Totaal"-regel, want sommige bonnen tonen ook een subtotaal-per-btw-tarief.
BEDRAG_INCL_PATROON = re.compile(
    r"Totaal\s*(?:bedrag)?\s*incl\.?\s*(?:BTW)?\s*[:\s]*(?:EUR|€)?\s*([0-9][0-9\.]*,\d{2})", re.IGNORECASE
)
TOTAAL_BEDRAG_PATROON = re.compile(
    r"Totaal\s*(?:incl\.?\s*BTW)?\s*[:\s]*(?:EUR|€)?\s*([0-9][0-9\.]*,\d{2})", re.IGNORECASE
)
ELK_BEDRAG_PATROON = re.compile(r"(?:€|EUR)\s*([0-9][0-9\.]*,\d{2})", re.IGNORECASE)

DATUM_LABEL_PATROON = re.compile(
    r"(?:Factuurdatum|Datum|Dagtekening|Documentdatum)\s*[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", re.IGNORECASE
)
DATUM_KAAL_PATROON = re.compile(r"\b(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})\b")
DATUM_TEKST_PATROON = re.compile(
    r"\b(\d{1,2})\s+(januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december)\s+(\d{4})\b",
    re.IGNORECASE,
)


def vind_tesseract() -> str:
    if getattr(sys, "frozen", False):
        gebundeld = Path(sys._MEIPASS) / "tesseract" / "tesseract.exe"
        if gebundeld.exists():
            return str(gebundeld)
    standaard = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if standaard.exists():
        return str(standaard)
    gevonden = shutil.which("tesseract")
    return gevonden or ""


def _heeft_nederlandse_taaldata(tesseract_cmd: str) -> bool:
    tessdata = Path(tesseract_cmd).parent / "tessdata"
    return (tessdata / "nld.traineddata").exists()


def _ocr_pagina(pagina_afbeelding, tesseract_cmd: str, taal: str) -> str:
    import pytesseract

    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    try:
        osd = pytesseract.image_to_osd(pagina_afbeelding)
        rotatie = int(re.search(r"Rotate:\s*(\d+)", osd).group(1))
        if rotatie:
            pagina_afbeelding = pagina_afbeelding.rotate(360 - rotatie, expand=True)
    except Exception:
        pass
    return pytesseract.image_to_string(pagina_afbeelding, lang=taal)


PDF_EXTENSIES = {".pdf"}
BEELD_EXTENSIES = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".heif"}


def _registreer_heic():
    """Probeer HEIC/HEIF-ondersteuning aan te zetten (telefoon-foto's). Stil falen als niet aanwezig."""
    try:
        import pillow_heif

        pillow_heif.register_heif_opener()
    except Exception:
        pass


def lees_pdf_tekst(pdf_pad: Path) -> tuple:
    """Retourneert (tekst, ocr_gebruikt)."""
    tekst_delen = []
    try:
        with pdfplumber.open(pdf_pad) as pdf:
            for pagina in pdf.pages:
                tekst_delen.append(pagina.extract_text() or "")
    except Exception:
        pass
    tekst = "\n".join(tekst_delen)
    if tekst.strip():
        return tekst, False

    tesseract_cmd = vind_tesseract()
    if not tesseract_cmd:
        return "", False

    taal = "nld+eng" if _heeft_nederlandse_taaldata(tesseract_cmd) else "eng"

    try:
        import fitz
        from PIL import Image

        doc = fitz.open(pdf_pad)
        ocr_delen = []
        for pagina in doc:
            pix = pagina.get_pixmap(dpi=250)
            afbeelding = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ocr_delen.append(_ocr_pagina(afbeelding, tesseract_cmd, taal))
        return "\n".join(ocr_delen), True
    except Exception:
        return "", False


def lees_afbeelding_tekst(beeld_pad: Path) -> tuple:
    """OCR direct op een losse foto/afbeelding (JPG/PNG/HEIC). Retourneert (tekst, ocr_gebruikt)."""
    tesseract_cmd = vind_tesseract()
    if not tesseract_cmd:
        return "", False
    taal = "nld+eng" if _heeft_nederlandse_taaldata(tesseract_cmd) else "eng"
    try:
        from PIL import Image

        if beeld_pad.suffix.lower() in (".heic", ".heif"):
            _registreer_heic()
        with Image.open(beeld_pad) as afbeelding:
            afbeelding = afbeelding.convert("RGB")
            tekst = _ocr_pagina(afbeelding, tesseract_cmd, taal)
        return tekst, True
    except Exception:
        return "", False


def vind_kandidaat_nummers(tekst: str) -> list:
    gelabeld = [m.group(1) for m in LABEL_NUMMER_PATROON.finditer(tekst)]
    kaal = KAAL_NUMMER_PATROON.findall(tekst)
    gezien = set()
    resultaat = []
    for nummer in gelabeld + kaal:
        genormaliseerd = nummer.strip()
        if genormaliseerd not in gezien:
            gezien.add(genormaliseerd)
            resultaat.append(genormaliseerd)
    return resultaat


def vind_bedrag(tekst: str):
    match = BEDRAG_INCL_PATROON.findall(tekst) or TOTAAL_BEDRAG_PATROON.findall(tekst)
    bron = match if match else ELK_BEDRAG_PATROON.findall(tekst)
    if not bron:
        return None
    bedragen = [float(b.replace(".", "").replace(",", ".")) for b in bron]
    return max(bedragen)


def vind_datum(tekst: str):
    for patroon in (DATUM_LABEL_PATROON, DATUM_KAAL_PATROON):
        match = patroon.search(tekst)
        if match:
            try:
                if patroon is DATUM_LABEL_PATROON:
                    return _parse_datum_string(match.group(1))
                dag, maand, jaar = match.groups()
                jaar = int(jaar) if len(jaar) == 4 else 2000 + int(jaar)
                return datetime(jaar, int(maand), int(dag))
            except ValueError:
                continue
    match = DATUM_TEKST_PATROON.search(tekst)
    if match:
        dag, maand_naam, jaar = match.groups()
        try:
            return datetime(int(jaar), MAAND_NAAM_NAAR_NUMMER[maand_naam.lower()], int(dag))
        except ValueError:
            return None
    return None


def _parse_datum_string(tekst: str):
    for scheider in ("-", "/"):
        delen = tekst.split(scheider)
        if len(delen) == 3:
            dag, maand, jaar = delen
            jaar = int(jaar) if len(jaar) == 4 else 2000 + int(jaar)
            return datetime(jaar, int(maand), int(dag))
    return None


def lees_bon(bon_pad: Path) -> dict:
    if bon_pad.suffix.lower() in BEELD_EXTENSIES:
        tekst, ocr_gebruikt = lees_afbeelding_tekst(bon_pad)
    else:
        tekst, ocr_gebruikt = lees_pdf_tekst(bon_pad)
    return {
        "bestand": bon_pad.name,
        "tekst": tekst,
        "tekst_gevonden": bool(tekst.strip()),
        "ocr_gebruikt": ocr_gebruikt,
        "kandidaat_nummers": vind_kandidaat_nummers(tekst) if tekst else [],
        "bedrag": vind_bedrag(tekst) if tekst else None,
        "datum": vind_datum(tekst) if tekst else None,
    }


def lees_bonnen_parallel(pdf_paden: list, voortgang_callback=None) -> list:
    """Leest meerdere bonnen tegelijk (OCR draait als los proces per bon, dus prima
    parallelliseerbaar via threads - de Tesseract-aanroep zelf is toch al een subprocess."""
    totaal = len(pdf_paden)
    resultaten = {}
    afgerond = 0
    with ThreadPoolExecutor(max_workers=min(8, (len(pdf_paden) or 1))) as executor:
        toekomsten = {executor.submit(lees_bon, pad): pad for pad in pdf_paden}
        for toekomst in as_completed(toekomsten):
            pad = toekomsten[toekomst]
            resultaten[pad] = toekomst.result()
            afgerond += 1
            if voortgang_callback:
                voortgang_callback(afgerond, totaal, pad.name)
    return [resultaten[pad] for pad in pdf_paden]


def _naam_gelijkenis(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()


_KORTE_WOORDEN = {"b.v.", "bv", "n.v.", "nv", "the", "van", "de", "der", "een", "en", "via"}


def _naam_past_bij_tekst(transactie_naam: str, bon_tekst: str) -> float:
    """Fractie van de betekenisvolle woorden uit de transactienaam die in de bontekst voorkomen."""
    bon_tekst_lower = bon_tekst.lower()
    woorden = [w.strip(".,()") for w in transactie_naam.lower().split()]
    woorden = [w for w in woorden if len(w) > 2 and w not in _KORTE_WOORDEN]
    if not woorden:
        return 0.0
    gevonden = sum(1 for w in woorden if w in bon_tekst_lower)
    return gevonden / len(woorden)


def laad_overrides(bonnen_map: Path) -> dict:
    """Handmatige correcties in overrides.json (in de bonnenmap zelf), bijvoorbeeld:
    {"rare_bon.pdf": {"forceer_nummer": "12345"}} of {"dubbele_bon.pdf": {"genegeerd": true}}."""
    pad = bonnen_map / "overrides.json"
    if not pad.exists():
        return {}
    try:
        with open(pad, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


STRAK_DAGEN_VENSTER = 5
NAAM_DREMPEL = 0.3


def match_bon(bon: dict, uitgaven: list, dagen_venster: int = 45):
    """Retourneert (transacties_lijst, zekerheid, gebruikt_nummer).

    Zoekt eerst op factuurnummer, uitsluitend in het Omschrijving-deel van elke transactie
    (tx["factuur"]) - niet de volledige Mededelingen-tekst, anders geven toevallige
    cijferoverlappen in gedeelde velden (IBAN, Machtiging ID, Incassant ID zijn identiek bij
    elke incasso van dezelfde crediteur) valse matches over alle transacties van die partij heen.

    Geen nummer gevonden? Dan bedrag+datum: een uniek bedrag (maar 1 kandidaat past) dat ook nog
    eens dicht bij de bondatum ligt (binnen STRAK_DAGEN_VENSTER) en waarvan de leveranciersnaam
    ook echt in de bontekst staat, is net zo overtuigend als een factuurnummer -> "zeker". Past
    het bedrag wel maar zit er twijfel in (meerdere kandidaten, of ver in tijd, of naam niet
    herkend), dan blijft het "onzeker" - zelf controleren.
    """
    for nummer in bon["kandidaat_nummers"]:
        treffers = [t for t in uitgaven if t["factuur"] and nummer in t["factuur"]]
        if treffers and len(treffers) <= 2:
            return treffers, "zeker", nummer
        # Matcht een nummer meer dan 2 transacties, is het waarschijnlijk een klant/
        # debiteurnummer (komt bij elke betaling van die partij terug), geen factuurnummer.

    if bon["bedrag"] is None:
        return [], None, None

    kandidaten = []
    for t in uitgaven:
        if abs(t["bedrag"] - bon["bedrag"]) > 0.02:
            continue
        if bon["datum"] is not None:
            verschil_dagen = (t["datum"] - bon["datum"]).days
            if not (-5 <= verschil_dagen <= dagen_venster):
                continue
        kandidaten.append(t)

    if not kandidaten:
        return [], None, None
    if len(kandidaten) == 1:
        kandidaat = kandidaten[0]
        dagen_verschil = abs((kandidaat["datum"] - bon["datum"]).days) if bon["datum"] else None
        naam_score = _naam_past_bij_tekst(kandidaat["naam"], bon["tekst"]) if bon["tekst"] else 0.0
        if dagen_verschil is not None and dagen_verschil <= STRAK_DAGEN_VENSTER and naam_score >= NAAM_DREMPEL:
            return kandidaten, "zeker", None
        return kandidaten, "onzeker", None

    def score(t):
        naam_score = _naam_past_bij_tekst(t["naam"], bon["tekst"]) if bon["tekst"] else 0.0
        dagen_score = abs((t["datum"] - bon["datum"]).days) if bon["datum"] else 0
        return (-naam_score, dagen_score)

    kandidaten.sort(key=score)
    return kandidaten[:1], "onzeker", None


def analyseer_bonnenmap(bonnen_map: Path, transacties: list, voortgang_callback=None) -> dict:
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]
    overrides = laad_overrides(bonnen_map)

    geldige_ext = PDF_EXTENSIES | BEELD_EXTENSIES
    bon_paden = sorted({p.resolve() for p in bonnen_map.iterdir() if p.suffix.lower() in geldige_ext})
    bonnen = lees_bonnen_parallel(bon_paden, voortgang_callback)

    zeker, onzeker, geen_match, onleesbaar = [], [], [], []
    gematchte_transacties = set()

    for bon in bonnen:
        override = overrides.get(bon["bestand"])
        if override and override.get("genegeerd"):
            continue
        if override and override.get("forceer_nummer"):
            bon["kandidaat_nummers"] = [override["forceer_nummer"]] + bon["kandidaat_nummers"]

        if not bon["tekst_gevonden"] and not bon["kandidaat_nummers"]:
            onleesbaar.append(bon)
            continue
        treffers, zekerheid, gebruikt_nummer = match_bon(bon, uitgaven)
        for t in treffers:
            gematchte_transacties.add(id(t))
        if zekerheid == "zeker":
            zeker.append((bon, treffers, gebruikt_nummer))
        elif zekerheid == "onzeker":
            onzeker.append((bon, treffers))
        else:
            geen_match.append(bon)

    # Kredietfacturen (terugontvangen geld) hebben per definitie geen inkoopbon en horen
    # niet in "uitgaven zonder bon" of het btw-risico.
    uitgaven_zonder_bon = [
        t for t in uitgaven
        if id(t) not in gematchte_transacties and not t.get("is_kredietfactuur")
    ]

    # Btw op uitgaven zonder bon = het bedrag dat je NIET veilig kan terugvragen (geen factuur
    # om de aftrek te onderbouwen). Bij een controle schrapt de Belastingdienst dit.
    def _btw_bedrag(t):
        pct = t.get("btw_pct", 21)
        bedrag = t.get("bedrag_effect", t["bedrag"])
        return bedrag - bedrag / (1 + pct / 100)

    btw_risico = round(sum(_btw_bedrag(t) for t in uitgaven_zonder_bon), 2)

    return {
        "zeker": zeker,
        "onzeker": onzeker,
        "geen_match": geen_match,
        "onleesbaar": onleesbaar,
        "uitgaven_zonder_bon": uitgaven_zonder_bon,
        "totaal_bonnen": len(bonnen),
        "btw_risico": btw_risico,
    }
