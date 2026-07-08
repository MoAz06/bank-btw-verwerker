#!/usr/bin/env python3
"""ING csv-export -> btw/inkomsten-uitgaven Excel overzicht.

Gebruik:
    python ing_btw_verwerker.py <ing_export.csv> [output.xlsx]
"""
import csv
import itertools
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DEFAULT_CONFIG = {
    "eigen_rekeningen": {"NL76INGB0003298592", "NL29RABO0366729276"},
    "eigen_naam_patronen": ["zakelijke oranje spaarrekening"],
    "nul_btw_patronen": [
        "kosten zakelijk betalingsverkeer", "ontvangsten hiltermann l",
        "flexkracht", "unive het groene hart bemiddeling", "vwpfs",
    ],
}
BELASTINGDIENST_PATROON = "belastingdienst"
# Wordt in main() overschreven met de lijst uit config.json.
NUL_BTW_PATRONEN = list(DEFAULT_CONFIG["nul_btw_patronen"])
STANDAARD_BTW = 21
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


def _kleur_voor(tx: dict) -> str:
    if tx["is_kredietfactuur"]:
        return KLEUR_KREDIETFACTUUR
    return KLEUR_UITGAVE if tx["richting"] == "uitgave" else KLEUR_INKOMST


def _naam_weergave(tx: dict) -> str:
    return tx["naam"] + (" (kredietfactuur)" if tx["is_kredietfactuur"] else "")


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
    ws.auto_filter.ref = "A1:H1"
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
            btw_pct, is_auto_nul = standaard_btw_percentage(tx)
            uitgave = tx["bedrag"] if tx["richting"] == "uitgave" else None
            inkomst = tx["bedrag"] if tx["richting"] == "inkomst" else None
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
    ws.cell(row=4, column=4, value="Uitgaven")
    ws.cell(row=4, column=5, value="=B5").number_format = GELD_FORMAAT
    ws.cell(row=5, column=4, value="Inkomsten")
    ws.cell(row=5, column=5, value="=B10").number_format = GELD_FORMAAT

    grafiek = BarChart()
    grafiek.type = "col"
    grafiek.title = "Uitgaven vs Inkomsten (Excl. Btw)"
    grafiek.y_axis.title = "EUR"
    grafiek.legend = None
    data = Reference(ws, min_col=5, min_row=3, max_row=5)
    categorieen = Reference(ws, min_col=4, min_row=4, max_row=5)
    grafiek.add_data(data, titles_from_data=True)
    grafiek.set_categories(categorieen)
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
        btw_pct, is_auto_nul = standaard_btw_percentage(tx)

        ws.cell(row=rij, column=1, value=tx["datum"]).number_format = "dd-mm-yyyy"
        ws.cell(row=rij, column=2, value=_naam_weergave(tx))
        ws.cell(row=rij, column=3, value=tx["factuur"])
        bedrag_cel = ws.cell(row=rij, column=4, value=tx["bedrag"])
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

    rij = 1
    banner(rij, f"Bonnencontrole - {analyse['totaal_bonnen']} bonnen verwerkt", KLEUR_HEADER_BG)
    rij += 2

    banner(rij, f"ZEKER GEMATCHT ({len(analyse['zeker'])})", KLEUR_ZEKER)
    rij += 1
    for bon, treffers, gebruikt_nummer in analyse["zeker"]:
        transactie_tekst = "; ".join(
            f"{t['datum'].strftime('%d-%m-%Y')} {t['naam']} {_formateer_euro(t['bedrag'])}" for t in treffers
        )
        ws.cell(row=rij, column=1, value=bon["bestand"])
        ws.cell(row=rij, column=2, value=gebruikt_nummer)
        ws.cell(row=rij, column=3, value=transactie_tekst)
        rij += 1
    rij += 1

    banner(rij, f"CONTROLEER - ONZEKERE MATCH ({len(analyse['onzeker'])})", KLEUR_ONZEKER)
    rij += 1
    for bon, treffers in analyse["onzeker"]:
        transactie_tekst = "; ".join(
            f"{t['datum'].strftime('%d-%m-%Y')} {t['naam']} {_formateer_euro(t['bedrag'])}" for t in treffers
        )
        bon_info = f"bedrag {_formateer_euro(bon['bedrag'])}" if bon["bedrag"] else "bedrag onbekend"
        if bon["datum"]:
            bon_info += f", datum {bon['datum'].strftime('%d-%m-%Y')}"
        ws.cell(row=rij, column=1, value=bon["bestand"])
        ws.cell(row=rij, column=2, value=bon_info)
        ws.cell(row=rij, column=3, value=f"mogelijk: {transactie_tekst}")
        rij += 1
    rij += 1

    banner(rij, f"GEEN MATCH GEVONDEN VOOR DEZE BON ({len(analyse['geen_match'])})", KLEUR_GEEN)
    rij += 1
    for bon in analyse["geen_match"]:
        info = []
        if bon["bedrag"]:
            info.append(_formateer_euro(bon["bedrag"]))
        if bon["datum"]:
            info.append(bon["datum"].strftime("%d-%m-%Y"))
        ws.cell(row=rij, column=1, value=bon["bestand"])
        ws.cell(row=rij, column=2, value=", ".join(bon["kandidaat_nummers"][:3]))
        ws.cell(row=rij, column=3, value=" / ".join(info) if info else "(geen gegevens herkend)")
        rij += 1
    rij += 1

    banner(rij, f"ONLEESBAAR - geen tekst/OCR gelukt ({len(analyse['onleesbaar'])})", KLEUR_GRIJS)
    rij += 1
    for bon in analyse["onleesbaar"]:
        ws.cell(row=rij, column=1, value=bon["bestand"])
        rij += 1
    rij += 1

    banner(rij, f"UITGAVEN ZONDER BON ({len(analyse['uitgaven_zonder_bon'])})", KLEUR_GEEN)
    rij += 1
    for t in sorted(analyse["uitgaven_zonder_bon"], key=lambda t: t["datum"]):
        ws.cell(row=rij, column=1, value=t["datum"].strftime("%d-%m-%Y"))
        ws.cell(row=rij, column=2, value=t["naam"])
        ws.cell(row=rij, column=3, value=_formateer_euro(t["bedrag"]))
        rij += 1

    for col, breedte in zip("ABCDE", [35, 30, 55, 14, 14]):
        ws.column_dimensions[col].width = breedte
    return ws


def bouw_workbook(transacties: list, buiten_periode: list) -> Workbook:
    wb = Workbook()
    bouw_transacties_sheet(wb, transacties)
    bouw_overzicht_sheet(wb, buiten_periode)
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]
    inkomsten = [t for t in transacties if t["richting"] == "inkomst"]
    bouw_richting_sheet(wb, "Uitgaven", KLEUR_UITGAVE, uitgaven)
    bouw_richting_sheet(wb, "Inkomsten", KLEUR_INKOMST, inkomsten)
    return wb


def verwerk_bestand(csv_pad: Path, output_pad: Path = None) -> dict:
    """Verwerkt een ING csv-export tot een Excel-overzicht. Retourneert een resultaten-dict.

    Gooit SystemExit met een duidelijke Nederlandse melding bij een ongeldig bestand.
    """
    if not csv_pad.exists():
        raise SystemExit(f"Bestand niet gevonden: {csv_pad}")
    output_pad = output_pad or csv_pad.with_name(f"{csv_pad.stem}_btw_overzicht.xlsx")

    global NUL_BTW_PATRONEN
    config = laad_config(programma_map())
    NUL_BTW_PATRONEN = config["nul_btw_patronen"]
    transacties = lees_csv(csv_pad)
    if not transacties:
        raise SystemExit("Geen (geldige) transacties gevonden in dit bestand.")

    geldige_jaarmaanden = bepaal_geldige_jaarmaanden(csv_pad.name)

    aantal_prive = sum(1 for t in transacties if is_prive(t, config))
    aantal_belastingdienst = sum(
        1 for t in transacties if not is_prive(t, config) and is_belastingdienst(t)
    )
    overig = [t for t in transacties if not is_prive(t, config) and not is_belastingdienst(t)]

    verwerkt = []
    buiten_periode = []
    aantal_geen_factuur = 0
    aantal_kredietfactuur = 0
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
        verwerkt.append({**t, "richting": richting, "is_kredietfactuur": is_krediet, "factuur": omschrijving})

    wb = bouw_workbook(verwerkt, buiten_periode)
    wb.save(output_pad)

    return {
        "output_pad": output_pad,
        "aantal_verwerkt": len(verwerkt),
        "aantal_prive": aantal_prive,
        "aantal_belastingdienst": aantal_belastingdienst,
        "aantal_geen_factuur": aantal_geen_factuur,
        "aantal_kredietfactuur": aantal_kredietfactuur,
        "aantal_buiten_periode": len(buiten_periode),
        "geen_periode_gevonden": geldige_jaarmaanden is None,
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
