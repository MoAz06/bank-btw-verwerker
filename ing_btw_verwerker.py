#!/usr/bin/env python3
"""ING csv-export -> btw/inkomsten-uitgaven Excel overzicht.

Gebruik:
    python ing_btw_verwerker.py <ing_export.csv> [output.xlsx]
"""
import csv
import itertools
import json
import statistics
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

import re

from leveranciers import (
    laad_leveranciers_geheugen,
    normaliseer_leverancier,
    schrijf_leveranciers_geheugen,
    vul_geheugen_aan,
)

DEFAULT_CONFIG = {
    "eigen_rekeningen": {"NL76INGB0003298592", "NL29RABO0366729276"},
    "eigen_naam_patronen": ["zakelijke oranje spaarrekening"],
    "nul_btw_patronen": [
        "kosten zakelijk betalingsverkeer", "ontvangsten hiltermann l",
        "flexkracht", "unive het groene hart bemiddeling", "vwpfs",
    ],
    "verlegd_patronen": [],
    "uitsluit_patronen": ["lening", "aflossing"],
    "investering_drempel": 450.0,
    # Leveranciers die (ook bij grote bedragen) nooit een investering zijn: materiaal-
    # groothandels, betaaldiensten, lease, verzekering, telecom, parkeren.
    "geen_investering_leveranciers": [
        "Wasco", "Technische Unie", "Bouwmaat", "Hornbach", "Toolstation",
        "Proforto", "Sanitino", "De Ned. Afvalcontainer Centrale",
        "Verploegen Verwarming", "Mollie Payments", "MultiSafepay", "Stripe", "PayPal",
        "VWPFS (lease)", "Hiltermann Lease", "Travelcard", "KPN",
        "Unive Het Groene Hart Bemiddeling B.V.", "Riverty (parkeren)",
    ],
    "leverancier_aliassen": {},
    "kor_actief": False,
    "jaaromzet_reeds": 0.0,
    "boekhoud_export": True,
}
BELASTINGDIENST_PATROON = "belastingdienst"
# Wordt in verwerk_bestand() overschreven met de lijst uit config.json.
NUL_BTW_PATRONEN = list(DEFAULT_CONFIG["nul_btw_patronen"])
STANDAARD_BTW = 21
KOR_GRENS = 20000.0
GELD_FORMAAT = '€ #,##0.00'
MAAND_NAMEN = ["januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november", "december"]

KLEUR_HEADER_BG = "1F4E78"
KLEUR_UITGAVE = "C00000"
KLEUR_INKOMST = "006100"
KLEUR_KREDIETFACTUUR = "002060"
KLEUR_WAARSCHUWING_BG = "FFC000"
KLEUR_SUBTOTAAL_BG = "DDEBF7"
KLEUR_BAND_BG = "F2F2F2"
KLEUR_GOED_BG, KLEUR_GOED_FONT = "C6EFCE", "006100"
KLEUR_SLECHT_BG, KLEUR_SLECHT_FONT = "FFC7CE", "9C0006"

REQUIRED_COLUMNS = [
    "Datum", "Naam / Omschrijving", "Rekening", "Tegenrekening",
    "Code", "Af Bij", "Bedrag (EUR)", "Mutatiesoort", "Mededelingen",
]

# Een "echte" factuur van jouw klant heeft een omschrijving met jaarmaand+cijfers (bv. 2026050012).
# Inkomsten met een andere soort verwijzing/nummer zijn kredietfacturen (terugbetaling op een
# eerdere uitgave) en tellen mee als uitgave. Inkomsten zonder enige verwijzing vervallen.
# Valt de jaarmaand van de factuur buiten de periode van de bestandsnaam (bv. een maart-factuur
# in een bestand dat in april begint), dan telt die factuur nergens in mee en wordt hij apart
# gemarkeerd zodat je 'm handmatig kan controleren.
FACTUUR_PATROON = re.compile(r"\b(\d{4})(0[1-9]|1[0-2])\d{2,6}\b")
CIJFER_PATROON = re.compile(r"\d{2,}")
_STOP_LABELS = r"(?:Naam:|Omschrijving:|IBAN:|Kenmerk:|Machtiging ID:|Incassant ID:|Datum/Tijd:|Valutadatum:)"
OMSCHRIJVING_PATROON = re.compile(r"Omschrijving:\s*(.*?)\s*(?=" + _STOP_LABELS + r"|$)")
BESTANDSNAAM_DATUM_PATROON = re.compile(r"(\d{2})-(\d{2})-(\d{4})_(\d{2})-(\d{2})-(\d{4})")


def bepaal_geldige_jaarmaanden(bestandsnaam: str):
    """(Jaar, maand)-paren die in de periode van de bestandsnaam vallen (dd-mm-jjjj_dd-mm-jjjj).

    Retourneert None als de bestandsnaam niet dat patroon volgt (dan geldt geen periodefilter).
    """
    match = BESTANDSNAAM_DATUM_PATROON.search(bestandsnaam)
    if not match:
        return None
    _, maand1, jaar1, _, maand2, jaar2 = match.groups()
    start = (int(jaar1), int(maand1))
    eind = (int(jaar2), int(maand2))
    if eind < start:
        start, eind = eind, start
    geldig = set()
    jaar, maand = start
    while (jaar, maand) <= eind:
        geldig.add((jaar, maand))
        maand += 1
        if maand > 12:
            maand = 1
            jaar += 1
    return geldig


def programma_map() -> Path:
    """Map waar config.json hoort te staan - naast de .exe (bevroren) of naast dit script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def laad_config(script_dir: Path) -> dict:
    pad = script_dir / "config.json"
    if not pad.exists():
        return DEFAULT_CONFIG
    with open(pad, encoding="utf-8") as f:
        data = json.load(f)
    return {
        "eigen_rekeningen": {iban.replace(" ", "").upper() for iban in data.get("eigen_rekeningen", [])},
        "eigen_naam_patronen": [p.lower() for p in data.get("eigen_naam_patronen", [])],
        "nul_btw_patronen": [p.lower() for p in data.get("nul_btw_patronen", DEFAULT_CONFIG["nul_btw_patronen"])],
        "verlegd_patronen": [p.lower() for p in data.get("verlegd_patronen", [])],
        "uitsluit_patronen": [p.lower() for p in data.get("uitsluit_patronen", DEFAULT_CONFIG["uitsluit_patronen"])],
        "investering_drempel": float(data.get("investering_drempel", DEFAULT_CONFIG["investering_drempel"])),
        "geen_investering_leveranciers": data.get("geen_investering_leveranciers", DEFAULT_CONFIG["geen_investering_leveranciers"]),
        "leverancier_aliassen": {k.lower(): v for k, v in data.get("leverancier_aliassen", {}).items()},
        "kor_actief": bool(data.get("kor_actief", False)),
        "jaaromzet_reeds": float(data.get("jaaromzet_reeds", 0.0)),
        "boekhoud_export": bool(data.get("boekhoud_export", True)),
    }


def parse_bedrag(tekst: str) -> float:
    return float(tekst.replace(".", "").replace(",", "."))


def parse_datum(tekst: str) -> datetime:
    return datetime.strptime(tekst, "%Y%m%d")


def lees_csv(pad: Path) -> list:
    with open(pad, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        ontbrekend = [c for c in REQUIRED_COLUMNS if c not in (reader.fieldnames or [])]
        if ontbrekend:
            raise SystemExit(
                f"Dit lijkt geen ING csv-export: kolommen {ontbrekend} ontbreken.\n"
                f"Gevonden kolommen: {reader.fieldnames}"
            )
        transacties = []
        for i, rij in enumerate(reader, start=2):
            try:
                transacties.append({
                    "datum": parse_datum(rij["Datum"]),
                    "naam": rij["Naam / Omschrijving"].strip(),
                    "tegenrekening": rij["Tegenrekening"].strip(),
                    "af_bij": rij["Af Bij"].strip(),
                    "bedrag": parse_bedrag(rij["Bedrag (EUR)"]),
                    "mededelingen": rij["Mededelingen"].strip(),
                })
            except (ValueError, KeyError) as e:
                print(f"Waarschuwing: rij {i} overgeslagen ({e})", file=sys.stderr)
        return transacties


def is_prive(transactie: dict, config: dict) -> bool:
    tegenrekening = transactie["tegenrekening"].replace(" ", "").upper()
    if tegenrekening and tegenrekening in config["eigen_rekeningen"]:
        return True
    naam_lower = transactie["naam"].lower()
    return any(patroon in naam_lower for patroon in config["eigen_naam_patronen"])


def is_belastingdienst(transactie: dict) -> bool:
    return BELASTINGDIENST_PATROON in transactie["naam"].lower()


def is_uitgesloten(transactie: dict, config: dict) -> bool:
    """Leningen/aflossingen e.d. zijn geen kosten of omzet (balans-mutaties) - eruit.

    Hele-woord-match: "lening" mag NIET matchen op "Dienstverlening".
    """
    tekst = (transactie["naam"] + " " + transactie["mededelingen"]).lower()
    for patroon in config.get("uitsluit_patronen", []):
        if re.search(r"\b" + re.escape(patroon) + r"\b", tekst):
            return True
    return False


def bepaal_omschrijving_tekst(mededelingen: str) -> str:
    match = OMSCHRIJVING_PATROON.search(mededelingen)
    return match.group(1) if match else ""


def classificeer_richting(transactie: dict, omschrijving: str, geldige_jaarmaanden):
    """Richting van een transactie: (richting, is_kredietfactuur, buiten_periode_jaarmaand).

    richting is een van "uitgave", "inkomst", "buiten_periode" of None (negeren).
    Uitgaven (Af) tellen altijd mee. Inkomsten (Bij) tellen alleen mee als er een
    herkenbare factuur/verwijzing in de omschrijving staat: jaarmaand+cijfers binnen de
    periode van het bestand is een echte factuur (inkomst); dezelfde vorm maar buiten de
    periode wordt apart gemarkeerd en telt nergens in mee; een andere verwijzing is een
    kredietfactuur (uitgave); geen verwijzing betekent negeren.
    """
    if transactie["af_bij"] == "Af":
        return "uitgave", False, None

    buiten_periode = None
    for match in FACTUUR_PATROON.finditer(omschrijving):
        jaarmaand = (int(match.group(1)), int(match.group(2)))
        if geldige_jaarmaanden is None or jaarmaand in geldige_jaarmaanden:
            return "inkomst", False, None
        if buiten_periode is None:
            buiten_periode = jaarmaand
    if buiten_periode is not None:
        return "buiten_periode", False, buiten_periode

    if CIJFER_PATROON.search(omschrijving):
        return "uitgave", True, None
    return None, False, None


def standaard_btw_percentage(transactie: dict):
    naam_lower = transactie["naam"].lower()
    if any(patroon in naam_lower for patroon in NUL_BTW_PATRONEN):
        return 0, True
    return STANDAARD_BTW, False


def bepaal_btw(transactie: dict, leverancier: str, config: dict, geheugen: dict):
    """Bepaalt het btw-percentage + herkomst voor een transactie.

    Volgorde (wettelijke/config-regels gaan VOOR een onthouden gok):
    verlegde btw (0%) -> automatische 0%-partij (config) -> onthouden keuze
    (leveranciers.json, alleen een plausibel 0-100 tarief) -> standaard 21%.
    Retourneert (btw_pct, bron, is_verlegd, is_auto_nul).
    """
    naam_lower = transactie["naam"].lower()

    if any(patroon in naam_lower for patroon in config.get("verlegd_patronen", [])):
        return 0, "verlegd", True, False

    if any(patroon in naam_lower for patroon in config.get("nul_btw_patronen", NUL_BTW_PATRONEN)):
        return 0, "nul-auto", False, True

    onthouden = geheugen.get(leverancier)
    if onthouden and isinstance(onthouden.get("btw"), (int, float)):
        onthouden_btw = int(onthouden["btw"])
        if 0 <= onthouden_btw <= 100:
            return onthouden_btw, "geheugen", False, onthouden_btw == 0

    return STANDAARD_BTW, "standaard", False, False


def _excl_van(bedrag: float, btw_pct: int) -> float:
    noemer = 1 + btw_pct / 100
    if noemer == 0:  # onmogelijk btw-tarief (-100%); val terug op het bedrag zelf
        return bedrag
    return bedrag / noemer


def _kleur_voor(tx: dict) -> str:
    if tx.get("is_kredietfactuur"):
        return KLEUR_KREDIETFACTUUR
    return KLEUR_UITGAVE if tx["richting"] == "uitgave" else KLEUR_INKOMST


def _naam_weergave(tx: dict) -> str:
    naam = tx["naam"]
    if tx.get("is_kredietfactuur"):
        naam += " (kredietfactuur)"
    if tx.get("is_verlegd"):
        naam += " (btw verlegd)"
    if tx.get("is_investering"):
        naam += " (investering)"
    return naam


def _formateer_euro(bedrag: float) -> str:
    return f"€ {bedrag:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _stijl_header_rij(ws, rij_nummer, aantal_kolommen):
    fill = PatternFill("solid", start_color=KLEUR_HEADER_BG, end_color=KLEUR_HEADER_BG)
    for col in range(1, aantal_kolommen + 1):
        cell = ws.cell(row=rij_nummer, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill


def bouw_transacties_sheet(wb: Workbook, transacties: list):
    ws = wb.active
    ws.title = "Transacties"

    headers = ["Datum", "Naam / Omschrijving", "Factuur", "Uitgave Incl. Btw", "Inkomst Incl. Btw", "Btw %", "Excl. Btw", "Btw-bedrag", "Type"]
    ws.append(headers)
    _stijl_header_rij(ws, 1, len(headers))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:I1"
    ws.sheet_properties.outlinePr.summaryBelow = True

    legende = ws.cell(row=1, column=11, value="Rood = uitgave  |  Groen = inkomst  |  Oranje = kredietfactuur")
    legende.font = Font(italic=True, color="808080")

    geel = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
    band_fill = PatternFill("solid", start_color=KLEUR_BAND_BG, end_color=KLEUR_BAND_BG)
    subtotaal_fill = PatternFill("solid", start_color=KLEUR_SUBTOTAAL_BG, end_color=KLEUR_SUBTOTAAL_BG)
    boven_rand = Border(top=Side(style="thin"))

    gesorteerd = sorted(transacties, key=lambda t: t["datum"])
    band_index = 0

    for (jaar, maand), groep in itertools.groupby(gesorteerd, key=lambda t: (t["datum"].year, t["datum"].month)):
        maand_start_rij = ws.max_row + 1
        for tx in groep:
            rij = ws.max_row + 1
            btw_pct = tx["btw_pct"]
            is_auto_nul = tx.get("btw_auto_nul", False)
            effect = tx.get("bedrag_effect", tx["bedrag"])
            uitgave = effect if tx["richting"] == "uitgave" else None
            inkomst = effect if tx["richting"] == "inkomst" else None
            kleur = _kleur_voor(tx)

            ws.cell(row=rij, column=1, value=tx["datum"]).number_format = "dd-mm-yyyy"
            ws.cell(row=rij, column=2, value=_naam_weergave(tx))
            ws.cell(row=rij, column=3, value=tx["factuur"])
            ws.cell(row=rij, column=4, value=uitgave).number_format = GELD_FORMAAT
            ws.cell(row=rij, column=5, value=inkomst).number_format = GELD_FORMAAT
            btw_cell = ws.cell(row=rij, column=6, value=btw_pct)
            excl_cell = ws.cell(row=rij, column=7, value=f"=(D{rij}+E{rij})/(1+F{rij}/100)")
            excl_cell.number_format = GELD_FORMAAT
            btw_bedrag_cell = ws.cell(row=rij, column=8, value=f"=(D{rij}+E{rij})-G{rij}")
            btw_bedrag_cell.number_format = GELD_FORMAAT
            ws.cell(row=rij, column=9, value="Detail")

            for col in (4, 7, 8):
                ws.cell(row=rij, column=col).font = Font(color=kleur)
            if band_index % 2 == 1:
                for col in range(1, 9):
                    ws.cell(row=rij, column=col).fill = band_fill
            if is_auto_nul:
                btw_cell.fill = geel
            ws.row_dimensions[rij].outline_level = 1
            band_index += 1

        maand_eind_rij = ws.max_row
        subtotaal_rij = ws.max_row + 1
        ws.cell(row=subtotaal_rij, column=2, value=f"Totaal {MAAND_NAMEN[maand - 1]} {jaar}")
        ws.cell(row=subtotaal_rij, column=4, value=f"=SUM(D{maand_start_rij}:D{maand_eind_rij})").number_format = GELD_FORMAAT
        ws.cell(row=subtotaal_rij, column=5, value=f"=SUM(E{maand_start_rij}:E{maand_eind_rij})").number_format = GELD_FORMAAT
        ws.cell(row=subtotaal_rij, column=7, value=f"=SUM(G{maand_start_rij}:G{maand_eind_rij})").number_format = GELD_FORMAAT
        ws.cell(row=subtotaal_rij, column=8, value=f"=SUM(H{maand_start_rij}:H{maand_eind_rij})").number_format = GELD_FORMAAT
        ws.cell(row=subtotaal_rij, column=9, value="Subtotaal")
        for col in range(1, 9):
            cell = ws.cell(row=subtotaal_rij, column=col)
            cell.fill = subtotaal_fill
            cell.font = Font(bold=True)
            cell.border = boven_rand
        band_index = 0

    for col, breedte in zip("ABCDEFGHI", [12, 40, 26, 16, 16, 8, 14, 14, 10]):
        ws.column_dimensions[col].width = breedte
    ws.column_dimensions["I"].hidden = True
    ws.column_dimensions["K"].width = 45

    return ws


def bouw_overzicht_sheet(wb: Workbook, buiten_periode: list):
    ws = wb.create_sheet("Overzicht")

    def banner(rij, tekst, kleur, kolommen=2, grootte=14):
        ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=kolommen)
        cell = ws.cell(row=rij, column=1, value=tekst)
        cell.font = Font(bold=True, color="FFFFFF", size=grootte)
        cell.fill = PatternFill("solid", start_color=kleur, end_color=kleur)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[rij].height = 22 if grootte > 14 else 18

    banner(1, "Btw & Inkomsten/Uitgaven overzicht", KLEUR_HEADER_BG, kolommen=5, grootte=16)

    banner(3, "UITGAVEN", KLEUR_UITGAVE)
    ws.cell(row=4, column=1, value="Totaal Uitgaven Incl. Btw").font = Font(bold=True)
    ws.cell(row=4, column=2, value='=SUMIFS(Transacties!D:D,Transacties!I:I,"Detail")').number_format = GELD_FORMAAT
    ws.cell(row=5, column=1, value="Totaal Uitgaven Excl. Btw").font = Font(bold=True)
    ws.cell(row=5, column=2, value='=SUMIFS(Transacties!G:G,Transacties!I:I,"Detail",Transacties!D:D,"<>")').number_format = GELD_FORMAAT
    ws.cell(row=6, column=1, value="Totaal Uitgaven Btw").font = Font(bold=True)
    ws.cell(row=6, column=2, value='=SUMIFS(Transacties!H:H,Transacties!I:I,"Detail",Transacties!D:D,"<>")').number_format = GELD_FORMAAT

    banner(8, "INKOMSTEN", KLEUR_INKOMST)
    ws.cell(row=9, column=1, value="Totaal Inkomsten Incl. Btw").font = Font(bold=True)
    ws.cell(row=9, column=2, value='=SUMIFS(Transacties!E:E,Transacties!I:I,"Detail")').number_format = GELD_FORMAAT
    ws.cell(row=10, column=1, value="Totaal Inkomsten Excl. Btw").font = Font(bold=True)
    ws.cell(row=10, column=2, value='=SUMIFS(Transacties!G:G,Transacties!I:I,"Detail",Transacties!E:E,"<>")').number_format = GELD_FORMAAT
    ws.cell(row=11, column=1, value="Totaal Inkomsten Btw").font = Font(bold=True)
    ws.cell(row=11, column=2, value='=SUMIFS(Transacties!H:H,Transacties!I:I,"Detail",Transacties!E:E,"<>")').number_format = GELD_FORMAAT

    banner(13, "RESULTAAT", KLEUR_HEADER_BG)
    ws.cell(row=14, column=1, value="Winst (Excl. Btw)").font = Font(bold=True)
    winst_cel = ws.cell(row=14, column=2, value="=B10-B5")
    winst_cel.number_format = GELD_FORMAAT
    ws.cell(row=15, column=1, value="Btw-saldo (+ afdragen / - terugvragen)").font = Font(bold=True)
    ws.cell(row=15, column=2, value="=B11-B6").number_format = GELD_FORMAAT
    ws.cell(row=16, column=1, value="Btw dat van jou blijft (Uitgaven Btw - Inkomsten Btw)").font = Font(bold=True)
    ws.cell(row=16, column=2, value="=B6-B11").number_format = GELD_FORMAAT

    ws.conditional_formatting.add(
        "B14",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"],
                   fill=PatternFill("solid", start_color=KLEUR_GOED_BG, end_color=KLEUR_GOED_BG),
                   font=Font(color=KLEUR_GOED_FONT, bold=True)),
    )
    ws.conditional_formatting.add(
        "B14",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", start_color=KLEUR_SLECHT_BG, end_color=KLEUR_SLECHT_BG),
                   font=Font(color=KLEUR_SLECHT_FONT, bold=True)),
    )

    ws.cell(row=3, column=4, value="Categorie").font = Font(bold=True)
    ws.cell(row=3, column=5, value="Bedrag (Excl. Btw)").font = Font(bold=True)
    ws.cell(row=4, column=4, value="Inkomsten")
    ws.cell(row=4, column=5, value="=B10").number_format = GELD_FORMAAT
    ws.cell(row=5, column=4, value="Uitgaven")
    ws.cell(row=5, column=5, value="=B5").number_format = GELD_FORMAAT

    grafiek = BarChart()
    grafiek.type = "col"
    grafiek.title = "Inkomsten vs Uitgaven (Excl. Btw)"
    grafiek.y_axis.title = "EUR"
    grafiek.legend = None
    data = Reference(ws, min_col=5, min_row=3, max_row=5)
    categorieen = Reference(ws, min_col=4, min_row=4, max_row=5)
    grafiek.add_data(data, titles_from_data=True)
    grafiek.set_categories(categorieen)

    punt_inkomst = DataPoint(idx=0)
    punt_inkomst.graphicalProperties.solidFill = KLEUR_INKOMST
    punt_uitgave = DataPoint(idx=1)
    punt_uitgave.graphicalProperties.solidFill = KLEUR_UITGAVE
    grafiek.series[0].data_points = [punt_inkomst, punt_uitgave]

    grafiek.width, grafiek.height = 10, 7
    ws.add_chart(grafiek, "D7")

    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 18

    if buiten_periode:
        start_rij = 22
        ws.merge_cells(start_row=start_rij, start_column=1, end_row=start_rij, end_column=5)
        banner_cel = ws.cell(row=start_rij, column=1, value="LET OP - FACTUUR BUITEN PERIODE VAN DIT BESTAND")
        banner_cel.font = Font(bold=True, color="FFFFFF", size=12)
        banner_cel.fill = PatternFill("solid", start_color=KLEUR_WAARSCHUWING_BG, end_color=KLEUR_WAARSCHUWING_BG)
        banner_cel.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[start_rij].height = 20

        rij = start_rij + 1
        for tx in sorted(buiten_periode, key=lambda t: t["datum"]):
            jaar, maand = tx["buiten_jaarmaand"]
            ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=5)
            tekst = (
                f"{tx['datum'].strftime('%d-%m-%Y')}   {tx['naam']}   "
                f"Factuur: {tx['factuur']}  ({MAAND_NAMEN[maand - 1]} {jaar})   "
                f"Bedrag: {_formateer_euro(tx['bedrag'])}"
            )
            cel = ws.cell(row=rij, column=1, value=tekst)
            cel.alignment = Alignment(horizontal="left")
            ws.row_dimensions[rij].height = 16
            rij += 1

    return ws


def bouw_richting_sheet(wb: Workbook, titel: str, kleur: str, transacties: list):
    ws = wb.create_sheet(titel)
    kop = ["Datum", "Naam / Omschrijving", "Factuur", "Bedrag Incl. Btw", "Btw %", "Excl. Btw", "Btw-bedrag"]
    geel = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(kop))
    banner_cel = ws.cell(row=1, column=1, value=titel)
    banner_cel.font = Font(bold=True, color="FFFFFF", size=14)
    banner_cel.fill = PatternFill("solid", start_color=kleur, end_color=kleur)
    banner_cel.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    for i, kopnaam in enumerate(kop, start=1):
        c = ws.cell(row=2, column=i, value=kopnaam)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=KLEUR_HEADER_BG, end_color=KLEUR_HEADER_BG)

    rij = 3
    for tx in sorted(transacties, key=lambda t: t["datum"]):
        font_kleur = _kleur_voor(tx)
        btw_pct = tx["btw_pct"]
        is_auto_nul = tx.get("btw_auto_nul", False)

        ws.cell(row=rij, column=1, value=tx["datum"]).number_format = "dd-mm-yyyy"
        ws.cell(row=rij, column=2, value=_naam_weergave(tx))
        ws.cell(row=rij, column=3, value=tx["factuur"])
        bedrag_cel = ws.cell(row=rij, column=4, value=tx.get("bedrag_effect", tx["bedrag"]))
        bedrag_cel.number_format = GELD_FORMAAT
        bedrag_cel.font = Font(color=font_kleur)
        btw_cel = ws.cell(row=rij, column=5, value=btw_pct)
        if is_auto_nul:
            btw_cel.fill = geel
        excl_cel = ws.cell(row=rij, column=6, value=f"=D{rij}/(1+E{rij}/100)")
        excl_cel.number_format = GELD_FORMAAT
        excl_cel.font = Font(color=font_kleur)
        btwbedrag_cel = ws.cell(row=rij, column=7, value=f"=D{rij}-F{rij}")
        btwbedrag_cel.number_format = GELD_FORMAAT
        btwbedrag_cel.font = Font(color=font_kleur)
        rij += 1

    laatste_data_rij = rij - 1
    if laatste_data_rij >= 3:
        totaal_rij = rij
        ws.cell(row=totaal_rij, column=2, value=f"Totaal {titel.lower()}").font = Font(bold=True)
        ws.cell(row=totaal_rij, column=4, value=f"=SUM(D3:D{laatste_data_rij})").number_format = GELD_FORMAAT
        ws.cell(row=totaal_rij, column=6, value=f"=SUM(F3:F{laatste_data_rij})").number_format = GELD_FORMAAT
        ws.cell(row=totaal_rij, column=7, value=f"=SUM(G3:G{laatste_data_rij})").number_format = GELD_FORMAAT
        for col in range(1, len(kop) + 1):
            c = ws.cell(row=totaal_rij, column=col)
            c.font = Font(bold=True)
            c.border = Border(top=Side(style="thin"))

    for col, breedte in zip("ABCDEFG", [12, 45, 30, 16, 8, 14, 14]):
        ws.column_dimensions[col].width = breedte
    ws.freeze_panes = "A3"
    return ws


def bouw_bonnencontrole_sheet(wb: Workbook, analyse: dict):
    KLEUR_ZEKER = KLEUR_INKOMST
    KLEUR_ONZEKER = KLEUR_WAARSCHUWING_BG
    KLEUR_GEEN = KLEUR_UITGAVE
    KLEUR_GRIJS = "808080"
    KLEUR_ZEKER_RIJ = "E2EFDA"
    KLEUR_ONZEKER_RIJ = "FFF2CC"
    KLEUR_GEEN_RIJ = "FCE4E4"
    KLEUR_GRIJS_RIJ = "EDEDED"

    if "Bonnencontrole" in wb.sheetnames:
        del wb["Bonnencontrole"]
    ws = wb.create_sheet("Bonnencontrole")

    def banner(rij, tekst, kleur, kolommen=5):
        ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=kolommen)
        cel = ws.cell(row=rij, column=1, value=tekst)
        cel.font = Font(bold=True, color="FFFFFF", size=12)
        cel.fill = PatternFill("solid", start_color=kleur, end_color=kleur)
        cel.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[rij].height = 20

    def tabel_sectie(start_rij, titel, kleur, kleur_rij, kolomnamen, databronnen, tabelnaam):
        """Bouwt 1 sectie: banner, kolomkoppen, gekleurde datarijen, als losse Excel-tabel (filter+structuur)."""
        banner(start_rij, f"{titel} ({len(databronnen)})", kleur)
        kop_rij = start_rij + 1
        for i, naam in enumerate(kolomnamen, start=1):
            c = ws.cell(row=kop_rij, column=i, value=naam)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color=KLEUR_HEADER_BG, end_color=KLEUR_HEADER_BG)

        rij = kop_rij
        for waarden in databronnen:
            rij += 1
            for i, waarde in enumerate(waarden, start=1):
                cel = ws.cell(row=rij, column=i, value=waarde)
                cel.fill = PatternFill("solid", start_color=kleur_rij, end_color=kleur_rij)
                if isinstance(waarde, datetime):
                    cel.number_format = "dd-mm-yyyy"
                elif isinstance(waarde, (int, float)) and not isinstance(waarde, bool):
                    cel.number_format = GELD_FORMAAT

        if rij > kop_rij:
            tabel = Table(displayName=tabelnaam, ref=f"A{kop_rij}:{chr(64 + len(kolomnamen))}{rij}")
            tabel.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
            ws.add_table(tabel)
        else:
            ws.cell(row=rij + 1, column=1, value="(geen)")
            rij += 1
        return rij

    rij = 1
    banner(rij, f"Bonnencontrole - {analyse['totaal_bonnen']} bonnen verwerkt", KLEUR_HEADER_BG)
    rij += 2

    ws.cell(row=rij, column=1, value="Categorie").font = Font(bold=True)
    ws.cell(row=rij, column=2, value="Aantal").font = Font(bold=True)
    kern_start = rij + 1
    kerncijfers = [
        ("Zeker gematcht", len(analyse["zeker"])),
        ("Controleer (onzeker)", len(analyse["onzeker"])),
        ("Geen match", len(analyse["geen_match"])),
        ("Onleesbaar", len(analyse["onleesbaar"])),
        ("Uitgaven zonder bon", len(analyse["uitgaven_zonder_bon"])),
    ]
    for i, (naam, aantal) in enumerate(kerncijfers):
        ws.cell(row=kern_start + i, column=1, value=naam)
        ws.cell(row=kern_start + i, column=2, value=aantal)
    kern_eind = kern_start + len(kerncijfers) - 1

    risico = analyse.get("btw_risico", 0)
    risico_rij = kern_eind + 2
    banner(risico_rij, "BTW-RISICO: uitgaven zonder bon", KLEUR_UITGAVE if risico > 0 else KLEUR_INKOMST)
    ws.cell(row=risico_rij + 1, column=1, value="Btw die je nu terugvraagt zonder factuur eronder:").font = Font(bold=True)
    rc = ws.cell(row=risico_rij + 1, column=3, value=risico)
    rc.number_format = GELD_FORMAAT
    rc.font = Font(bold=True, size=12, color=(KLEUR_SLECHT_FONT if risico > 0 else KLEUR_GOED_FONT))
    ws.cell(row=risico_rij + 2, column=1, value=(
        "Dit is voorbelasting die de Belastingdienst bij controle kan schrappen. Zoek de bon op, "
        "of houd er rekening mee bij je 5b."
    )).font = Font(italic=True, color="595959")
    ws.merge_cells(start_row=risico_rij + 2, start_column=1, end_row=risico_rij + 2, end_column=5)

    rij = risico_rij + 4

    zeker_rijen = [
        (bon["bestand"], gebruikt_nummer or "(bedrag + datum uniek)", t["datum"], t["naam"], t["bedrag"])
        for bon, treffers, gebruikt_nummer in analyse["zeker"]
        for t in treffers
    ]
    rij = tabel_sectie(
        rij, "ZEKER GEMATCHT", KLEUR_ZEKER, KLEUR_ZEKER_RIJ,
        ["Bon", "Herkend nummer", "Transactiedatum", "Transactienaam", "Bedrag"], zeker_rijen, "BonZeker",
    )
    rij += 2

    onzeker_rijen = [
        (bon["bestand"], bon["bedrag"], bon["datum"], t["datum"], t["naam"], t["bedrag"])
        for bon, treffers in analyse["onzeker"]
        for t in treffers
    ]
    rij = tabel_sectie(
        rij, "CONTROLEER - ONZEKERE MATCH", KLEUR_ONZEKER, KLEUR_ONZEKER_RIJ,
        ["Bon", "Bedrag (bon)", "Datum (bon)", "Mogelijke transactiedatum", "Mogelijke naam", "Mogelijk bedrag"],
        onzeker_rijen, "BonOnzeker",
    )
    rij += 2

    geen_match_rijen = [
        (bon["bestand"], ", ".join(bon["kandidaat_nummers"][:3]), bon["bedrag"], bon["datum"])
        for bon in analyse["geen_match"]
    ]
    rij = tabel_sectie(
        rij, "GEEN MATCH GEVONDEN VOOR DEZE BON", KLEUR_GEEN, KLEUR_GEEN_RIJ,
        ["Bon", "Kandidaat-nummers", "Bedrag (bon)", "Datum (bon)"], geen_match_rijen, "BonGeenMatch",
    )
    rij += 2

    onleesbaar_rijen = [(bon["bestand"],) for bon in analyse["onleesbaar"]]
    rij = tabel_sectie(
        rij, "ONLEESBAAR - geen tekst/OCR gelukt", KLEUR_GRIJS, KLEUR_GRIJS_RIJ,
        ["Bon"], onleesbaar_rijen, "BonOnleesbaar",
    )
    rij += 2

    zonder_bon_rijen = [
        (t["datum"], t["naam"], t["bedrag"])
        for t in sorted(analyse["uitgaven_zonder_bon"], key=lambda t: t["datum"])
    ]
    tabel_sectie(
        rij, "UITGAVEN ZONDER BON", KLEUR_GEEN, KLEUR_GEEN_RIJ,
        ["Datum", "Naam", "Bedrag"], zonder_bon_rijen, "BonUitgavenZonderBon",
    )

    for col, breedte in zip("ABCDEF", [38, 22, 16, 30, 16, 14]):
        ws.column_dimensions[col].width = breedte
    return ws


def _blok_banner(ws, rij, tekst, kleur, kolommen, grootte=13):
    ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=kolommen)
    cel = ws.cell(row=rij, column=1, value=tekst)
    cel.font = Font(bold=True, color="FFFFFF", size=grootte)
    cel.fill = PatternFill("solid", start_color=kleur, end_color=kleur)
    cel.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[rij].height = 22 if grootte >= 14 else 18


def bouw_aangifte_sheet(wb: Workbook, aangifte: dict):
    """De cijfers die je letterlijk overtikt in Mijn Belastingdienst Zakelijk.

    Kern-rubrieken (1a/5a/5b/5c) zijn live gekoppeld aan het Overzicht, zodat ze
    meebewegen als je een btw% in Transacties aanpast. Verlegde btw en KOR worden
    apart en met een controleer-notitie getoond (fiscaal maatwerk).
    """
    ws = wb.create_sheet("BTW-aangifte")
    vv_excl = aangifte["verlegd_verkoop_excl"]
    vi_excl = aangifte["verlegd_inkoop_excl"]
    vi_btw = round(vi_excl * STANDAARD_BTW / 100, 2)

    _blok_banner(ws, 1, "BTW-AANGIFTE  —  over te tikken in Mijn Belastingdienst Zakelijk", KLEUR_HEADER_BG, 3, grootte=15)
    ws.cell(row=2, column=1, value=aangifte["periode_tekst"]).font = Font(italic=True, color="595959")
    ws.cell(row=3, column=1, value=f"Aangifte + betaling {aangifte['deadline_tekst']}").font = Font(italic=True, color="595959")

    if aangifte["kor_actief"]:
        _blok_banner(ws, 5, "KLEineOndernemersRegeling (KOR) ACTIEF", KLEUR_WAARSCHUWING_BG, 3, grootte=13)
        ws.cell(row=6, column=1, value=(
            "Onder de KOR doe je GEEN btw-aangifte: geen btw op je verkopen, geen voorbelasting terugvragen. "
            "De rubrieken hieronder zijn puur informatief."
        )).alignment = Alignment(wrap_text=True)
        ws.merge_cells("A6:C6")
        ws.row_dimensions[6].height = 30
        kop_rij = 8
    else:
        kop_rij = 5

    ws.cell(row=kop_rij, column=1, value="Rubriek").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=kop_rij, column=2, value="Omzet (excl. btw)").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=kop_rij, column=3, value="Btw").font = Font(bold=True, color="FFFFFF")
    for c in range(1, 4):
        ws.cell(row=kop_rij, column=c).fill = PatternFill("solid", start_color=KLEUR_HEADER_BG, end_color=KLEUR_HEADER_BG)

    rij = kop_rij + 1

    def rubriek(label, omzet_formule, btw_formule, vet=False):
        nonlocal rij
        ws.cell(row=rij, column=1, value=label).font = Font(bold=vet)
        if omzet_formule is not None:
            c = ws.cell(row=rij, column=2, value=omzet_formule)
            c.number_format = GELD_FORMAAT
        if btw_formule is not None:
            c = ws.cell(row=rij, column=3, value=btw_formule)
            c.number_format = GELD_FORMAAT
            c.font = Font(bold=vet)
        huidige = rij
        rij += 1
        return huidige

    # 1a omzet = alleen de 21%-belaste inkomsten (live SUMIFS), 1a btw = totale inkomsten-btw
    # (0%-inkomsten dragen 0 btw bij, dus B11 is exact de 21%-btw).
    omzet_1a = '=SUMIFS(Transacties!G:G,Transacties!I:I,"Detail",Transacties!E:E,"<>",Transacties!F:F,21)'
    rij_1a = rubriek("1a  Leveringen/diensten belast met 21%", omzet_1a, "=Overzicht!B11")
    rij_4b = None
    if vv_excl > 0:
        rubriek("1e  Omzet waarbij btw naar afnemer is verlegd (0%)", vv_excl, 0)
    if aangifte.get("overige_omzet_excl", 0) > 0:
        rubriek("Overige omzet 0% / anders (controleer rubriek)", f"=Overzicht!B10-C{rij_1a}-{vv_excl}", 0)
    if vi_excl > 0:
        rij_4b = rubriek("4b  Btw naar mij verlegd (inkoop onderaanneming)", vi_excl, vi_btw)

    btw_4b = f"+C{rij_4b}" if rij_4b else ""
    rij_5a = rubriek("5a  Verschuldigde omzetbelasting", None, f"=C{rij_1a}{btw_4b}", vet=True)
    rij_5b = rubriek("5b  Voorbelasting (terug te vragen)", None, f"=Overzicht!B6{('+' + str(vi_btw)) if vi_btw else ''}", vet=True)
    rij_5c = rubriek("5c  Te betalen (+) / terug te vragen (-)", None, f"=C{rij_5a}-C{rij_5b}", vet=True)

    ws.conditional_formatting.add(
        f"C{rij_5c}",
        CellIsRule(operator="greaterThan", formula=["0"],
                   fill=PatternFill("solid", start_color=KLEUR_SLECHT_BG, end_color=KLEUR_SLECHT_BG),
                   font=Font(color=KLEUR_SLECHT_FONT, bold=True)))
    ws.conditional_formatting.add(
        f"C{rij_5c}",
        CellIsRule(operator="lessThanOrEqual", formula=["0"],
                   fill=PatternFill("solid", start_color=KLEUR_GOED_BG, end_color=KLEUR_GOED_BG),
                   font=Font(color=KLEUR_GOED_FONT, bold=True)))

    if vv_excl > 0 or vi_excl > 0:
        ws.cell(row=rij + 1, column=1, value=(
            "Let op: verlegde btw is fiscaal maatwerk — controleer 1e/4b met je boekhouder."
        )).font = Font(italic=True, color=KLEUR_SLECHT_FONT)
        rij += 1

    # Btw-potje
    potje_rij = rij + 2
    _blok_banner(ws, potje_rij, "ZET DIT OPZIJ VOOR DE AANGIFTE", KLEUR_INKOMST, 3, grootte=13)
    ws.cell(row=potje_rij + 1, column=1, value="Reserveer nu op je spaarrekening:").font = Font(bold=True)
    reserve = ws.cell(row=potje_rij + 1, column=3, value=f"=MAX(0,C{rij_5c})")
    reserve.number_format = GELD_FORMAAT
    reserve.font = Font(bold=True, size=12, color=KLEUR_INKOMST)
    ws.cell(row=potje_rij + 2, column=1, value=(
        f"De btw die je klanten betaalden is niet van jou — {aangifte['deadline_tekst']} moet het naar de Belastingdienst."
    )).font = Font(italic=True, color="595959")

    # Jaaromzet / KOR-grens — alleen tonen als de KOR uberhaupt relevant is (in de buurt van
    # de grens, of KOR staat aan). Voor een ondernemer die ruim boven € 20.000 draait is dit ruis.
    jaaromzet = aangifte["jaaromzet_tot_nu"]
    if aangifte["kor_actief"] or jaaromzet <= 1.5 * KOR_GRENS:
        jaar_rij = potje_rij + 4
        _blok_banner(ws, jaar_rij, "JAAROMZET (KOR-grens € 20.000)", KLEUR_HEADER_BG, 3, grootte=12)
        ws.cell(row=jaar_rij + 1, column=1, value="Omzet excl. btw dit jaar tot nu:").font = Font(bold=True)
        jc = ws.cell(row=jaar_rij + 1, column=3, value=round(jaaromzet, 2))
        jc.number_format = GELD_FORMAAT
        if not aangifte["kor_actief"] and 0.9 * KOR_GRENS <= jaaromzet <= KOR_GRENS:
            ws.cell(row=jaar_rij + 2, column=1, value="Je nadert de KOR-grens — normaal geen probleem als je btw-plichtig bent.").font = Font(italic=True, color=KLEUR_SLECHT_FONT)
        if aangifte["kor_actief"] and jaaromzet > KOR_GRENS:
            ws.cell(row=jaar_rij + 2, column=1, value="LET OP: boven € 20.000 — je moet uit de KOR stappen en btw gaan afdragen.").font = Font(bold=True, color=KLEUR_SLECHT_FONT)

    for col, breedte in zip("ABC", [46, 20, 18]):
        ws.column_dimensions[col].width = breedte
    return ws


def bouw_investeringen_sheet(wb: Workbook, investeringen: list, drempel: float):
    ws = wb.create_sheet("Investeringen")
    _blok_banner(ws, 1, f"GROTE UITGAVEN — mogelijk investering (≥ {_formateer_euro(drempel)} excl.)", KLEUR_HEADER_BG, 6, grootte=14)
    ws.cell(row=2, column=1, value=(
        "Dit zijn kandidaten om te controleren, geen zekerheid. Is het een bedrijfsmiddel dat jaren "
        "meegaat (bus, machine, gereedschap)? Dan is de btw direct terug (zit al in 5b), maar het bedrag "
        "excl. btw schrijf je af over meerdere jaren i.p.v. in één keer als kosten. Zijn het gewoon "
        "materialen of een dienst, dan is het normale kosten. Overleg met je boekhouder."
    )).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:F3")
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 24

    kop = ["Datum", "Leverancier", "Factuur", "Bedrag incl.", "Excl. btw", "Btw"]
    kop_rij = 4
    for i, naam in enumerate(kop, start=1):
        c = ws.cell(row=kop_rij, column=i, value=naam)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=KLEUR_HEADER_BG, end_color=KLEUR_HEADER_BG)

    rij = kop_rij
    totaal_excl = 0.0
    for tx in sorted(investeringen, key=lambda t: t["datum"]):
        rij += 1
        excl = _excl_van(tx["bedrag"], tx["btw_pct"])
        totaal_excl += excl
        ws.cell(row=rij, column=1, value=tx["datum"]).number_format = "dd-mm-yyyy"
        ws.cell(row=rij, column=2, value=tx["leverancier"])
        ws.cell(row=rij, column=3, value=tx["factuur"])
        ws.cell(row=rij, column=4, value=tx["bedrag"]).number_format = GELD_FORMAAT
        ws.cell(row=rij, column=5, value=round(excl, 2)).number_format = GELD_FORMAAT
        ws.cell(row=rij, column=6, value=round(tx["bedrag"] - excl, 2)).number_format = GELD_FORMAAT

    if rij > kop_rij:
        tabel = Table(displayName="Investeringen", ref=f"A{kop_rij}:F{rij}")
        tabel.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tabel)
        totaal_rij = rij + 1
        ws.cell(row=totaal_rij, column=2, value="Totaal investeringen (excl. btw)").font = Font(bold=True)
        ws.cell(row=totaal_rij, column=5, value=round(totaal_excl, 2)).number_format = GELD_FORMAAT
        ws.cell(row=totaal_rij, column=5).font = Font(bold=True)
        if totaal_excl >= 2900:
            ws.cell(row=totaal_rij + 2, column=1, value=(
                "Mogelijk recht op Kleinschaligheidsinvesteringsaftrek (KIA) — vraag je boekhouder. "
                "(Indicatief; KIA geldt per kalenderjaar, dit is alleen deze periode.)"
            )).font = Font(bold=True, color=KLEUR_INKOMST)
    else:
        ws.cell(row=kop_rij + 1, column=1, value="(geen investeringen boven de drempel in deze periode)")

    for col, breedte in zip("ABCDEF", [12, 30, 26, 16, 16, 14]):
        ws.column_dimensions[col].width = breedte
    return ws


def bouw_leveranciers_sheet(wb: Workbook, uitgaven: list):
    ws = wb.create_sheet("Leveranciers")
    _blok_banner(ws, 1, "UITGAVEN PER LEVERANCIER", KLEUR_HEADER_BG, 6, grootte=14)

    per_leverancier = {}
    for tx in uitgaven:
        per_leverancier.setdefault(tx["leverancier"], []).append(tx)

    kop = ["Leverancier", "Aantal", "Totaal incl.", "Gemiddeld", "Eerste", "Laatste"]
    kop_rij = 3
    for i, naam in enumerate(kop, start=1):
        c = ws.cell(row=kop_rij, column=i, value=naam)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=KLEUR_HEADER_BG, end_color=KLEUR_HEADER_BG)

    rij = kop_rij
    for leverancier, txs in sorted(per_leverancier.items(), key=lambda kv: -sum(t.get("bedrag_effect", t["bedrag"]) for t in kv[1])):
        rij += 1
        bedragen = [t.get("bedrag_effect", t["bedrag"]) for t in txs]
        datums = [t["datum"] for t in txs]
        ws.cell(row=rij, column=1, value=leverancier)
        ws.cell(row=rij, column=2, value=len(txs))
        ws.cell(row=rij, column=3, value=round(sum(bedragen), 2)).number_format = GELD_FORMAAT
        ws.cell(row=rij, column=4, value=round(sum(bedragen) / len(bedragen), 2)).number_format = GELD_FORMAAT
        ws.cell(row=rij, column=5, value=min(datums)).number_format = "dd-mm-yyyy"
        ws.cell(row=rij, column=6, value=max(datums)).number_format = "dd-mm-yyyy"

    if rij > kop_rij:
        tabel = Table(displayName="Leveranciers", ref=f"A{kop_rij}:F{rij}")
        tabel.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabel)

    # Opvallende bedragen: transactie die sterk afwijkt van het typische bedrag van die leverancier.
    afwijkingen = []
    for leverancier, txs in per_leverancier.items():
        if len(txs) < 4:
            continue
        bedragen = [t.get("bedrag_effect", t["bedrag"]) for t in txs]
        mediaan = statistics.median(bedragen)
        for t in txs:
            bedrag = t.get("bedrag_effect", t["bedrag"])
            if mediaan > 0 and bedrag >= 3 * mediaan and bedrag - mediaan >= 100:
                afwijkingen.append((t, mediaan, bedrag))

    afw_rij = rij + 3
    _blok_banner(ws, afw_rij, f"OPVALLENDE BEDRAGEN (controleer) ({len(afwijkingen)})", KLEUR_WAARSCHUWING_BG, 6, grootte=12)
    afw_kop = afw_rij + 1
    for i, naam in enumerate(["Datum", "Leverancier", "Bedrag", "Typisch (mediaan)"], start=1):
        c = ws.cell(row=afw_kop, column=i, value=naam)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=KLEUR_HEADER_BG, end_color=KLEUR_HEADER_BG)
    r = afw_kop
    for t, mediaan, bedrag in sorted(afwijkingen, key=lambda x: -x[2]):
        r += 1
        ws.cell(row=r, column=1, value=t["datum"]).number_format = "dd-mm-yyyy"
        ws.cell(row=r, column=2, value=t["leverancier"])
        ws.cell(row=r, column=3, value=round(bedrag, 2)).number_format = GELD_FORMAAT
        ws.cell(row=r, column=4, value=round(mediaan, 2)).number_format = GELD_FORMAAT
    if not afwijkingen:
        ws.cell(row=afw_kop + 1, column=1, value="(geen opvallende bedragen gevonden)")

    for col, breedte in zip("ABCDEF", [30, 10, 16, 16, 12, 12]):
        ws.column_dimensions[col].width = breedte
    return ws


def bouw_workbook(transacties: list, buiten_periode: list, aangifte: dict, config: dict) -> Workbook:
    wb = Workbook()
    bouw_transacties_sheet(wb, transacties)
    bouw_overzicht_sheet(wb, buiten_periode)
    bouw_aangifte_sheet(wb, aangifte)
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]
    inkomsten = [t for t in transacties if t["richting"] == "inkomst"]
    bouw_richting_sheet(wb, "Uitgaven", KLEUR_UITGAVE, uitgaven)
    bouw_richting_sheet(wb, "Inkomsten", KLEUR_INKOMST, inkomsten)
    investeringen = [t for t in uitgaven if t.get("is_investering")]
    bouw_investeringen_sheet(wb, investeringen, config["investering_drempel"])
    bouw_leveranciers_sheet(wb, uitgaven)
    return wb


def _bepaal_aangifte_info(verwerkt: list, config: dict) -> dict:
    inkomsten = [t for t in verwerkt if t["richting"] == "inkomst"]
    uitgaven = [t for t in verwerkt if t["richting"] == "uitgave"]

    def excl(t):
        return _excl_van(t.get("bedrag_effect", t["bedrag"]), t["btw_pct"])

    def btw(t):
        return t.get("bedrag_effect", t["bedrag"]) - excl(t)

    inkomsten_excl = sum(excl(t) for t in inkomsten)
    inkomsten_btw = sum(btw(t) for t in inkomsten)
    uitgaven_excl = sum(excl(t) for t in uitgaven)
    uitgaven_btw = sum(btw(t) for t in uitgaven)
    verlegd_verkoop_excl = round(sum(excl(t) for t in inkomsten if t.get("is_verlegd")), 2)
    verlegd_inkoop_excl = round(sum(excl(t) for t in uitgaven if t.get("is_verlegd")), 2)
    # Omzet die niet 21% en niet verlegd is (bv. 0%-inkomsten) hoort NIET in rubriek 1a.
    overige_omzet_excl = round(
        sum(excl(t) for t in inkomsten if not t.get("is_verlegd") and t["btw_pct"] != 21), 2
    )

    datums = [t["datum"] for t in verwerkt] or [datetime.now()]
    periode_start, periode_eind = min(datums), max(datums)
    volgende_maand = periode_eind.month % 12 + 1
    volgende_jaar = periode_eind.year + (1 if periode_eind.month == 12 else 0)
    deadline_tekst = f"uiterlijk eind {MAAND_NAMEN[volgende_maand - 1]} {volgende_jaar} (1 maand na kwartaaleinde)"
    periode_tekst = (
        f"Periode {periode_start.strftime('%d-%m-%Y')} t/m {periode_eind.strftime('%d-%m-%Y')}"
    )

    btw_saldo = inkomsten_btw - uitgaven_btw
    return {
        "verlegd_verkoop_excl": verlegd_verkoop_excl,
        "verlegd_inkoop_excl": verlegd_inkoop_excl,
        "overige_omzet_excl": overige_omzet_excl,
        "jaaromzet_tot_nu": config["jaaromzet_reeds"] + inkomsten_excl,
        "kor_actief": config["kor_actief"],
        "deadline_tekst": deadline_tekst,
        "periode_tekst": periode_tekst,
        "winst": round(inkomsten_excl - uitgaven_excl, 2),
        "btw_saldo": round(btw_saldo, 2),
        "btw_te_reserveren": round(max(0.0, btw_saldo), 2),
    }


def verwerk_bestand(csv_pad: Path, output_pad: Path = None) -> dict:
    """Verwerkt een ING csv-export tot een Excel-overzicht. Retourneert een resultaten-dict.

    Gooit SystemExit met een duidelijke Nederlandse melding bij een ongeldig bestand.
    """
    if not csv_pad.exists():
        raise SystemExit(f"Bestand niet gevonden: {csv_pad}")
    output_pad = output_pad or csv_pad.with_name(f"{csv_pad.stem}_btw_overzicht.xlsx")

    global NUL_BTW_PATRONEN
    map_pad = programma_map()
    config = laad_config(map_pad)
    NUL_BTW_PATRONEN = config["nul_btw_patronen"]
    geheugen = laad_leveranciers_geheugen(map_pad)
    aliassen = config["leverancier_aliassen"]

    transacties = lees_csv(csv_pad)
    if not transacties:
        raise SystemExit("Geen (geldige) transacties gevonden in dit bestand.")

    geldige_jaarmaanden = bepaal_geldige_jaarmaanden(csv_pad.name)

    aantal_prive = sum(1 for t in transacties if is_prive(t, config))
    aantal_belastingdienst = sum(
        1 for t in transacties if not is_prive(t, config) and is_belastingdienst(t)
    )
    aantal_uitgesloten = sum(
        1 for t in transacties
        if not is_prive(t, config) and not is_belastingdienst(t) and is_uitgesloten(t, config)
    )
    overig = [
        t for t in transacties
        if not is_prive(t, config) and not is_belastingdienst(t) and not is_uitgesloten(t, config)
    ]

    verwerkt = []
    buiten_periode = []
    aantal_geen_factuur = 0
    aantal_kredietfactuur = 0
    nieuwe_leveranciers = 0
    for t in overig:
        omschrijving = bepaal_omschrijving_tekst(t["mededelingen"])
        richting, is_krediet, buiten_jaarmaand = classificeer_richting(t, omschrijving, geldige_jaarmaanden)
        if richting is None:
            aantal_geen_factuur += 1
            continue
        if richting == "buiten_periode":
            buiten_periode.append({**t, "factuur": omschrijving, "buiten_jaarmaand": buiten_jaarmaand})
            continue
        if is_krediet:
            aantal_kredietfactuur += 1

        leverancier = normaliseer_leverancier(t["naam"], aliassen)
        btw_pct, btw_bron, is_verlegd, is_auto_nul = bepaal_btw(t, leverancier, config, geheugen)
        excl = _excl_van(t["bedrag"], btw_pct)
        is_investering = (
            richting == "uitgave" and not is_krediet
            and excl >= config["investering_drempel"]
            and leverancier not in config.get("geen_investering_leveranciers", [])
        )

        tx = {
            **t,
            "richting": richting,
            "is_kredietfactuur": is_krediet,
            "factuur": omschrijving,
            "leverancier": leverancier,
            "btw_pct": btw_pct,
            "btw_bron": btw_bron,
            "btw_auto_nul": is_auto_nul,
            "is_verlegd": is_verlegd,
            "is_investering": is_investering,
            # Effectief bedrag voor alle geld-optellingen: een kredietfactuur (terugontvangen
            # geld op een eerdere inkoop) telt NEGATIEF, zodat het de uitgaven en voorbelasting
            # verlaagt i.p.v. verhoogt. "bedrag" blijft het ruwe positieve bedrag (voor bon-match).
            "bedrag_effect": -t["bedrag"] if is_krediet else t["bedrag"],
        }
        verwerkt.append(tx)
        if vul_geheugen_aan(geheugen, leverancier, btw_pct, richting == "uitgave"):
            nieuwe_leveranciers += 1

    schrijf_leveranciers_geheugen(map_pad, geheugen)

    aangifte = _bepaal_aangifte_info(verwerkt, config)
    wb = bouw_workbook(verwerkt, buiten_periode, aangifte, config)
    wb.save(output_pad)

    export_pad = None
    if config.get("boekhoud_export", True) and verwerkt:
        try:
            import boekhoud_export
            export_pad = output_pad.with_name(f"{csv_pad.stem}_boekhouder_import.csv")
            boekhoud_export.schrijf_export(verwerkt, export_pad)
        except Exception as fout:
            print(f"Waarschuwing: boekhoud-export mislukt ({fout})", file=sys.stderr)
            export_pad = None

    return {
        "output_pad": output_pad,
        "export_pad": export_pad,
        "aantal_verwerkt": len(verwerkt),
        "aantal_prive": aantal_prive,
        "aantal_belastingdienst": aantal_belastingdienst,
        "aantal_uitgesloten": aantal_uitgesloten,
        "aantal_geen_factuur": aantal_geen_factuur,
        "aantal_kredietfactuur": aantal_kredietfactuur,
        "aantal_buiten_periode": len(buiten_periode),
        "aantal_investeringen": sum(1 for t in verwerkt if t.get("is_investering")),
        "aantal_verlegd": sum(1 for t in verwerkt if t.get("is_verlegd")),
        "nieuwe_leveranciers": nieuwe_leveranciers,
        "geen_periode_gevonden": geldige_jaarmaanden is None,
        "kor_actief": config["kor_actief"],
        "winst": aangifte["winst"],
        "btw_saldo": aangifte["btw_saldo"],
        "btw_te_reserveren": aangifte["btw_te_reserveren"],
        "deadline_tekst": aangifte["deadline_tekst"],
        "transacties": verwerkt,
        "workbook": wb,
    }


def print_resultaat(resultaat: dict):
    if resultaat["geen_periode_gevonden"]:
        print(
            "Waarschuwing: kon geen periode uit de bestandsnaam halen - "
            "factuurherkenning accepteert nu elke jaarmaand.",
            file=sys.stderr,
        )
    if resultaat["aantal_verwerkt"] == 0:
        print("Waarschuwing: geen transacties over na filtering - controleer config.json.", file=sys.stderr)
    print(f"Klaar: {resultaat['aantal_verwerkt']} transacties verwerkt.")
    print(
        f"Uitgesloten: {resultaat['aantal_prive']} prive/spaar, {resultaat['aantal_belastingdienst']} belastingdienst, "
        f"{resultaat.get('aantal_uitgesloten', 0)} uitsluit-patroon (lening/aflossing e.d.), "
        f"{resultaat['aantal_geen_factuur']} inkomsten zonder herkenbare factuur."
    )
    print(f"Waarvan {resultaat['aantal_kredietfactuur']} kredietfacturen omgezet naar uitgave.")
    if resultaat["aantal_buiten_periode"]:
        print(f"Let op: {resultaat['aantal_buiten_periode']} factuur/facturen buiten de periode van dit bestand - zie Overzicht-tab.")
    print(f"Bestand geschreven naar: {resultaat['output_pad']}")


def main():
    if len(sys.argv) < 2:
        print("Gebruik: python ing_btw_verwerker.py <ing_export.csv> [output.xlsx]")
        sys.exit(1)

    csv_pad = Path(sys.argv[1])
    output_pad = Path(sys.argv[2]) if len(sys.argv) >= 3 else None
    resultaat = verwerk_bestand(csv_pad, output_pad)
    print_resultaat(resultaat)


def _voer_uit():
    try:
        main()
    except SystemExit as fout:
        if isinstance(fout.code, str):
            print(f"Fout: {fout.code}", file=sys.stderr)
        if getattr(sys, "frozen", False):
            input("\nDruk op Enter om af te sluiten...")
        sys.exit(fout.code if isinstance(fout.code, int) else 1)
    else:
        if getattr(sys, "frozen", False):
            input("\nDruk op Enter om af te sluiten...")


if __name__ == "__main__":
    _voer_uit()
