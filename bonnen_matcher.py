#!/usr/bin/env python3
"""Vergelijkt bonnen (PDF's) met de uitgaven uit een verwerkte ING csv."""
import difflib
import re
import shutil
import sys
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


def _ocr_pagina(pagina_afbeelding, tesseract_cmd: str) -> str:
    import pytesseract

    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    try:
        osd = pytesseract.image_to_osd(pagina_afbeelding)
        rotatie = int(re.search(r"Rotate:\s*(\d+)", osd).group(1))
        if rotatie:
            pagina_afbeelding = pagina_afbeelding.rotate(360 - rotatie, expand=True)
    except Exception:
        pass
    return pytesseract.image_to_string(pagina_afbeelding, lang="eng")


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

    try:
        import fitz
        from PIL import Image

        doc = fitz.open(pdf_pad)
        ocr_delen = []
        for pagina in doc:
            pix = pagina.get_pixmap(dpi=250)
            afbeelding = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ocr_delen.append(_ocr_pagina(afbeelding, tesseract_cmd))
        return "\n".join(ocr_delen), True
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
    match = TOTAAL_BEDRAG_PATROON.findall(tekst)
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


def lees_bon(pdf_pad: Path) -> dict:
    tekst, ocr_gebruikt = lees_pdf_tekst(pdf_pad)
    return {
        "bestand": pdf_pad.name,
        "tekst_gevonden": bool(tekst.strip()),
        "ocr_gebruikt": ocr_gebruikt,
        "kandidaat_nummers": vind_kandidaat_nummers(tekst) if tekst else [],
        "bedrag": vind_bedrag(tekst) if tekst else None,
        "datum": vind_datum(tekst) if tekst else None,
    }


def _naam_gelijkenis(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()


def match_bon(bon: dict, uitgaven: list, dagen_venster: int = 45):
    """Retourneert (transacties_lijst, zekerheid, gebruikt_nummer).

    Zoekt uitsluitend in het Omschrijving-deel van elke transactie (tx["factuur"]),
    niet de volledige Mededelingen-tekst - anders geven toevallige cijferoverlappen in
    gedeelde velden (IBAN, Machtiging ID, Incassant ID zijn identiek bij elke incasso
    van dezelfde crediteur) valse matches over alle transacties van die partij heen.
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
        return kandidaten, "onzeker", None
    kandidaten.sort(key=lambda t: abs((t["datum"] - bon["datum"]).days) if bon["datum"] else 0)
    return kandidaten[:1], "onzeker", None


def analyseer_bonnenmap(bonnen_map: Path, transacties: list, voortgang_callback=None) -> dict:
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]

    pdf_paden = sorted({p.resolve() for p in bonnen_map.iterdir() if p.suffix.lower() == ".pdf"})
    totaal = len(pdf_paden)
    bonnen = []
    for i, pdf_pad in enumerate(pdf_paden, start=1):
        if voortgang_callback:
            voortgang_callback(i, totaal, pdf_pad.name)
        bonnen.append(lees_bon(pdf_pad))

    zeker, onzeker, geen_match, onleesbaar = [], [], [], []
    gematchte_transacties = set()

    for bon in bonnen:
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

    uitgaven_zonder_bon = [t for t in uitgaven if id(t) not in gematchte_transacties]

    return {
        "zeker": zeker,
        "onzeker": onzeker,
        "geen_match": geen_match,
        "onleesbaar": onleesbaar,
        "uitgaven_zonder_bon": uitgaven_zonder_bon,
        "totaal_bonnen": len(bonnen),
    }
