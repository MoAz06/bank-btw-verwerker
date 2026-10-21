#!/usr/bin/env python3
"""Vergelijkt bonnen (PDF's) met de uitgaven uit een verwerkte ING csv."""
import difflib
import json
import re
import shutil
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import pdfplumber

MAAND_NAAM_NAAR_NUMMER = {
    "januari": 1, "februari": 2, "maart": 3, "april": 4, "mei": 5, "juni": 6,
    "juli": 7, "augustus": 8, "september": 9, "oktober": 10, "november": 11, "december": 12,
}

LABEL_NUMMER_PATROON = re.compile(
    r"(?:Factuurnummer|Factuurnr\.?|Factuur\s*nr\.?|Factuur|Contractnummer|Aanslagnummer|Kwitantienummer|"
    r"Bestelnummer|Bestelling\s*#|Ordernummer|Order\s*nr\.?|Referentie|Kenmerk)(?![A-Za-z])\s*[:\s#]*\s*([A-Za-z0-9][A-Za-z0-9\-\.]{3,20})",
    re.IGNORECASE,
)
KAAL_NUMMER_PATROON = re.compile(r"\b\d{5,15}\b")
# Factuurnummers met streepjes, zoals "8A5-2026-04", "2026-0103", "012-2026", "26-0126091".
CODE_NUMMER_PATROON = re.compile(r"\b(?=(?:[A-Z0-9\-]*\d){5})[A-Z0-9]{2,6}-\d{2,6}(?:-\d{2,6})?\b")

# Het totaalbedrag: het grootste bedrag dat bij een "Totaal"/"Te betalen"-label staat (incl. is
# altijd >= excl., en een subtotaal per btw-tarief is kleiner). Lukt dat niet, dan het grootste
# bedrag met een euroteken. Een bon met alleen "Totaal 0,00" levert geen bedrag op.
# Bedragen in Nederlandse notatie (1.234,56) én internationale webshop-notatie (1,234.56 / 56.88).
_BEDRAG = r"([0-9][0-9\.,]*[\.,]\d{2})(?!\d)"
TOTAAL_LABEL_PATROON = re.compile(
    r"(?:Tota(?:al|le)|Te\s+betalen|Eindbedrag|Totaalbedrag)[^0-9]{0,45}?" + _BEDRAG, re.IGNORECASE
)
ELK_BEDRAG_PATROON = re.compile(r"(?:€|EUR)\s*" + _BEDRAG, re.IGNORECASE)

# Datumlabels op volgorde van betrouwbaarheid: de factuurdatum gaat vóór een kale "Datum"
# (die vaak de afdruk- of betaaldatum is). "Vervaldatum"/"Leverdatum" tellen niet mee.
_DATUM_LABELS = ["Factuurdatum", "Documentdatum", r"Datum\s+bestelling", "Dagtekening", "Datum"]
_MAANDEN = "januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december"
DATUM_LABEL_PATRONEN = [
    (
        re.compile(r"(?<![A-Za-z])" + label + r"(?![A-Za-z])\s*[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", re.IGNORECASE),
        re.compile(r"(?<![A-Za-z])" + label + r"(?![A-Za-z])\s*[:\s]*(\d{4})-(\d{2})-(\d{2})", re.IGNORECASE),
        re.compile(r"(?<![A-Za-z])" + label + r"(?![A-Za-z])\s*[:\s]*(\d{1,2})\s+(" + _MAANDEN + r")\s+(\d{4})", re.IGNORECASE),
    )
    for label in _DATUM_LABELS
]
DATUM_KAAL_PATROON = re.compile(r"\b(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})\b")
DATUM_KASSA_PATROON = re.compile(r"\b(\d{4})-([a-z]{3})-(\d{1,2})\b", re.IGNORECASE)
MAAND_AFKORTING = {"jan": 1, "feb": 2, "mrt": 3, "maa": 3, "mar": 3, "apr": 4, "mei": 5, "may": 5, "jun": 6,
                   "jul": 7, "aug": 8, "sep": 9, "okt": 10, "oct": 10, "nov": 11, "dec": 12}
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
    tekst = pytesseract.image_to_string(pagina_afbeelding, lang=taal)
    if len(tekst.strip()) >= 40:
        return tekst
    # Bijna niets herkend: de scan ligt waarschijnlijk gedraaid. Probeer de andere richtingen
    # en houd de lezing met de meeste tekst.
    beste = tekst
    for graden in (90, 180, 270):
        poging = pytesseract.image_to_string(pagina_afbeelding.rotate(graden, expand=True), lang=taal)
        if len(poging.strip()) > len(beste.strip()):
            beste = poging
    return beste


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
    codes = CODE_NUMMER_PATROON.findall(tekst)
    kaal = KAAL_NUMMER_PATROON.findall(tekst)
    gezien = set()
    resultaat = []
    for nummer in gelabeld + codes + kaal:
        genormaliseerd = nummer.strip().rstrip(".")
        # Een datum ("06-03-2026") is geen factuurnummer.
        if DATUM_KAAL_PATROON.fullmatch(genormaliseerd):
            continue
        if genormaliseerd not in gezien:
            gezien.add(genormaliseerd)
            resultaat.append(genormaliseerd)
    return resultaat


def _naar_bedrag(tekst: str) -> float:
    """'1.234,56' -> 1234.56, '1,234.56' -> 1234.56, '56.88' -> 56.88: het laatste scheidingsteken is de komma."""
    laatste = max(tekst.rfind(","), tekst.rfind("."))
    heel, decimaal = tekst[:laatste], tekst[laatste + 1:]
    heel = re.sub(r"[\.,]", "", heel)
    return float(f"{heel or '0'}.{decimaal}")


def vind_bedrag(tekst: str):
    gelabeld = []
    for m in TOTAAL_LABEL_PATROON.finditer(tekst):
        if re.search(r"korting|bespaar|retour", m.group(0), re.IGNORECASE):
            continue  # "Totaal korting € 30,60" is geen totaalbedrag
        gelabeld.append(_naar_bedrag(m.group(1)))
    gelabeld = [b for b in gelabeld if b > 0]
    # Bedragen met een euroteken erbij tellen ook mee: op een gescande bon staat het totaal
    # soms los van het label ("Totaal incl. btw" op de ene regel, "EUR 69,65" verderop).
    # Een losse regel is nooit groter dan het totaal, dus het maximum blijft het totaal.
    los = [_naar_bedrag(b) for b in ELK_BEDRAG_PATROON.findall(tekst)]
    los = [b for b in los if b > 0]
    alles = gelabeld + los
    return max(alles) if alles else None


def vind_datum(tekst: str):
    for numeriek, iso, met_tekst in DATUM_LABEL_PATRONEN:
        match = numeriek.search(tekst)
        if match:
            try:
                return _parse_datum_string(match.group(1))
            except ValueError:
                pass
        match = iso.search(tekst)
        if match:
            try:
                jaar, maand, dag = match.groups()
                return datetime(int(jaar), int(maand), int(dag))
            except ValueError:
                pass
        match = met_tekst.search(tekst)
        if match:
            dag, maand_naam, jaar = match.groups()
            try:
                return datetime(int(jaar), MAAND_NAAM_NAAR_NUMMER[maand_naam.lower()], int(dag))
            except ValueError:
                pass
    match = DATUM_KAAL_PATROON.search(tekst)
    if match:
        try:
            dag, maand, jaar = match.groups()
            jaar = int(jaar) if len(jaar) == 4 else 2000 + int(jaar)
            return datetime(jaar, int(maand), int(dag))
        except ValueError:
            pass
    match = DATUM_KASSA_PATROON.search(tekst)  # "2026-feb-13" (Praxis-kassabon)
    if match:
        jaar, maand_afk, dag = match.groups()
        maand = MAAND_AFKORTING.get(maand_afk.lower())
        if maand:
            try:
                return datetime(int(jaar), maand, int(dag))
            except ValueError:
                pass
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
    """Handmatige correcties in overrides.json (in de bonnenmap zelf). Per bon een dict met:
    - "tx_id": "abc123def456" (of een lijst): vaste koppeling aan die transactie(s), wint van alles
    - "forceer_nummer": "12345": dit nummer vooraan in de kandidaat-nummers zetten
    - "genegeerd": true: bon overslaan (hoort bij een ander kwartaal, dubbel, geen uitgave)
    - "op_naam": true/false: handmatig vastleggen of het een factuur op jouw naam is
    - "reden": vrije tekst, wordt genegeerd door de code (notitie voor jezelf)."""
    pad = bonnen_map / "overrides.json"
    if not pad.exists():
        return {}
    try:
        with open(pad, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


# Een incasso of weekfactuur wordt doorgaans 7 tot 14 dagen na de factuurdatum afgeschreven.
STRAK_DAGEN_VENSTER = 14
NAAM_DREMPEL = 0.3


def _nummer_in(nummer: str, factuur: str) -> bool:
    """Nummer als los token in de omschrijving (grenzen: geen letter/cijfer eromheen), ook in de
    variant zonder spaties omdat ING lange nummers afbreekt ("573 0839")."""
    patroon = re.compile(r"(?<![A-Za-z0-9])" + re.escape(nummer) + r"(?![A-Za-z0-9])", re.IGNORECASE)
    return bool(patroon.search(factuur) or patroon.search(factuur.replace(" ", "")))


def _naam_score(t: dict, bon: dict) -> float:
    """Naam-overeenkomst op de banknaam én op de genormaliseerde leverancier (de banknaam is
    vaak een betaalprovider: "... via Stichting Mollie Payments")."""
    if not bon.get("tekst"):
        return 0.0
    scores = [_naam_past_bij_tekst(t["naam"], bon["tekst"])]
    if t.get("leverancier"):
        scores.append(_naam_past_bij_tekst(t["leverancier"], bon["tekst"]))
    return max(scores)


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
        # ING breekt lange omschrijvingen af met een spatie middenin een nummer ("573 0839"),
        # dus vergelijk ook met de omschrijving zonder spaties.
        # Als heel token, niet als deel van een ander nummer: "064519" (een telefoonfragment op een
        # KPN-factuur) mag niet matchen met Travelcard-factuur "N06451933".
        treffers = [t for t in uitgaven if t["factuur"] and _nummer_in(nummer, t["factuur"])]
        if treffers and len(treffers) <= 2:
            return treffers, "zeker", nummer
        # Matcht een nummer meer dan 2 transacties, dan is het waarschijnlijk een klant- of
        # debiteurnummer (komt bij elke betaling van die partij terug). Dat is nog steeds
        # bruikbaar: binnen die partij kiest bedrag + datum de juiste betaling.
        if len(treffers) > 2 and bon["bedrag"] is not None:
            passend = [t for t in treffers if abs(t["bedrag"] - bon["bedrag"]) <= 0.02]
            if bon["datum"] is not None:
                passend = [t for t in passend if -5 <= (t["datum"] - bon["datum"]).days <= dagen_venster]
                passend.sort(key=lambda t: abs((t["datum"] - bon["datum"]).days))
            if len(passend) == 1:
                return passend, "zeker", nummer
            if len(passend) > 1 and bon["datum"] is not None:
                eerste = abs((passend[0]["datum"] - bon["datum"]).days)
                tweede = abs((passend[1]["datum"] - bon["datum"]).days)
                if eerste <= STRAK_DAGEN_VENSTER and tweede - eerste >= STRAK_DAGEN_VENSTER:
                    return passend[:1], "zeker", nummer
                return passend[:1], "onzeker", nummer

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
        if dagen_verschil is not None and dagen_verschil <= STRAK_DAGEN_VENSTER and _naam_score(kandidaat, bon) >= NAAM_DREMPEL:
            return kandidaten, "zeker", None
        return kandidaten, "onzeker", None

    def score(t):
        dagen_score = abs((t["datum"] - bon["datum"]).days) if bon["datum"] else 0
        return (-_naam_score(t, bon), dagen_score)

    kandidaten.sort(key=score)
    return kandidaten[:1], "onzeker", None


# Boven dit bedrag (incl. btw) eist de Belastingdienst een factuur met jouw naam en adres
# voor aftrek van voorbelasting; daaronder volstaat een "vereenvoudigde factuur" (kassabon).
FACTUUR_GRENS = 100.0

BEWIJS_TEKST = {
    "factuur_op_naam": "Factuur op naam",
    "bon_voldoende": "Kassabon (voldoende onder EUR 100)",
    "bon_zonder_naam": "Kassabon, factuur op naam nodig",
    "onbekend": "Niet gecontroleerd",
}


def _plat(tekst: str) -> str:
    """Kleine letters zonder spaties/punten/streepjes, zodat 'NL 0012.34.567.B01' ook matcht."""
    return re.sub(r"[\s\.\-/]", "", (tekst or "").lower())


def staat_op_naam(bon_tekst: str, kenmerken: list):
    """True/False als er kenmerken zijn ingesteld, None als er niets te controleren valt."""
    kenmerken = [k for k in (kenmerken or []) if k and str(k).strip()]
    if not kenmerken or not bon_tekst:
        return None if not kenmerken else False
    plat = _plat(bon_tekst)
    return any(_plat(str(k)) in plat for k in kenmerken)


def bewijs_status(bon: dict, treffers: list) -> str:
    bedrag = max((abs(t.get("bedrag_effect", t["bedrag"])) for t in treffers), default=0.0)
    if bon.get("op_naam") is True:
        return "factuur_op_naam"
    if bedrag < FACTUUR_GRENS:
        return "bon_voldoende"
    if bon.get("op_naam") is None:
        return "onbekend"
    return "bon_zonder_naam"


def analyseer_bonnenmap(bonnen_map: Path, transacties: list, voortgang_callback=None, kenmerken: list = None,
                        gebruik_overrides: bool = True) -> dict:
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]
    overrides = laad_overrides(bonnen_map) if gebruik_overrides else {}
    kenmerken = [k for k in (kenmerken or []) if k and str(k).strip()]

    geldige_ext = PDF_EXTENSIES | BEELD_EXTENSIES
    bon_paden = sorted({p.resolve() for p in bonnen_map.iterdir() if p.suffix.lower() in geldige_ext})
    bonnen = lees_bonnen_parallel(bon_paden, voortgang_callback)

    zeker, onzeker, geen_match, onleesbaar, zonder_naam = [], [], [], [], []
    gematchte_transacties = set()

    for bon in bonnen:
        override = overrides.get(bon["bestand"])
        if not isinstance(override, dict):
            override = None
        if override and override.get("genegeerd"):
            continue
        if override and override.get("forceer_nummer"):
            bon["kandidaat_nummers"] = [str(override["forceer_nummer"])] + bon["kandidaat_nummers"]
        bon["op_naam"] = staat_op_naam(bon["tekst"], kenmerken)
        if override and "op_naam" in override:  # handmatig: {"bon.pdf": {"op_naam": true}}
            bon["op_naam"] = bool(override["op_naam"])

        vaste_ids = override.get("tx_id") if override else None
        if vaste_ids:
            # Vaste koppeling (handmatig of eerder door Kwartaal vastgelegd): geen heuristiek nodig,
            # ook niet voor een scan zonder leesbare tekst.
            if not isinstance(vaste_ids, list):
                vaste_ids = [vaste_ids]
            alle_uitgaven = [t for t in transacties if t["richting"] == "uitgave"]  # incl. kredietfacturen
            treffers = [t for t in alle_uitgaven if t.get("tx_id") in vaste_ids]
            treffers_uitgaven = [t for t in treffers if not t.get("is_kredietfactuur")]
            zekerheid, gebruikt_nummer = ("zeker", "vaste koppeling") if treffers else (None, None)
            if treffers:
                for t in treffers:
                    gematchte_transacties.add(id(t))
                bon["bewijs"] = bewijs_status(bon, treffers_uitgaven or treffers)
                if bon["bewijs"] == "bon_zonder_naam":
                    zonder_naam.append((bon, treffers))
                zeker.append((bon, treffers, gebruikt_nummer))
            else:
                geen_match.append(bon)
            continue

        if not bon["tekst_gevonden"] and not bon["kandidaat_nummers"]:
            onleesbaar.append(bon)
            continue
        treffers, zekerheid, gebruikt_nummer = match_bon(bon, uitgaven)
        for t in treffers:
            gematchte_transacties.add(id(t))
        if zekerheid in ("zeker", "onzeker"):
            bon["bewijs"] = bewijs_status(bon, treffers)
            if bon["bewijs"] == "bon_zonder_naam":
                zonder_naam.append((bon, treffers))
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
    # Btw op gematchte bonnen boven de factuurgrens waar jouw naam niet op staat: bij een
    # controle is dat net zo kwetsbaar als een ontbrekende bon.
    naam_txs = {id(t): t for _, treffers in zonder_naam for t in treffers}
    btw_risico_naam = round(sum(_btw_bedrag(t) for t in naam_txs.values()), 2)

    return {
        "zeker": zeker,
        "onzeker": onzeker,
        "geen_match": geen_match,
        "onleesbaar": onleesbaar,
        "uitgaven_zonder_bon": uitgaven_zonder_bon,
        "zonder_naam": zonder_naam,
        "totaal_bonnen": len(bonnen),
        "btw_risico": btw_risico,
        "btw_risico_naam": btw_risico_naam,
        "kenmerken_ingesteld": bool(kenmerken),
    }
