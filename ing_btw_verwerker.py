#!/usr/bin/env python3
"""ING csv-export -> btw/inkomsten-uitgaven Excel overzicht.

Gebruik:
    python ing_btw_verwerker.py <ing_export.csv> [output.xlsx]
"""
import itertools
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

from berekening import (
    KOR_GRENS,
    MAAND_NAMEN,
    bepaal_aangifte_info,
    bepaal_rubrieken,
    excl_van,
    totalen_per_richting,
)
from leveranciers import schrijf_leveranciers_geheugen
from boekhouding import (
    Boekhouding,
    bereken_model,
    programma_map,
)

# Compat-aliassen (interne aanroepen gebruiken deze namen nog).
_excl_van = excl_van
_bepaal_aangifte_info = bepaal_aangifte_info

GELD_FORMAAT = '€ #,##0.00'

KLEUR_HEADER_BG = "1F4E78"
KLEUR_UITGAVE = "C00000"
KLEUR_INKOMST = "006100"
KLEUR_KREDIETFACTUUR = "002060"
KLEUR_WAARSCHUWING_BG = "FFC000"
KLEUR_SUBTOTAAL_BG = "DDEBF7"
KLEUR_BAND_BG = "F2F2F2"
KLEUR_GOED_BG, KLEUR_GOED_FONT = "C6EFCE", "006100"
KLEUR_SLECHT_BG, KLEUR_SLECHT_FONT = "FFC7CE", "9C0006"

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


# De incl/excl/btw-aggregatie leeft in de rekenkern (berekening.py); hier alleen een alias
# zodat bestaande aanroepen (_totalen) blijven werken en app + Excel dezelfde bron delen.
_totalen = totalen_per_richting


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
        maand_som = {4: 0.0, 5: 0.0, 7: 0.0, 8: 0.0}
        for tx in groep:
            rij = ws.max_row + 1
            btw_pct = tx["btw_pct"]
            is_auto_nul = tx.get("btw_auto_nul", False)
            effect = tx.get("bedrag_effect", tx["bedrag"])
            uitgave = effect if tx["richting"] == "uitgave" else None
            inkomst = effect if tx["richting"] == "inkomst" else None
            excl = round(_excl_van(effect, btw_pct), 2)
            btw_bedrag = round(effect - excl, 2)
            kleur = _kleur_voor(tx)

            ws.cell(row=rij, column=1, value=tx["datum"]).number_format = "dd-mm-yyyy"
            ws.cell(row=rij, column=2, value=_naam_weergave(tx))
            ws.cell(row=rij, column=3, value=tx["factuur"])
            ws.cell(row=rij, column=4, value=uitgave).number_format = GELD_FORMAAT
            ws.cell(row=rij, column=5, value=inkomst).number_format = GELD_FORMAAT
            btw_cell = ws.cell(row=rij, column=6, value=btw_pct)
            ws.cell(row=rij, column=7, value=excl).number_format = GELD_FORMAAT
            ws.cell(row=rij, column=8, value=btw_bedrag).number_format = GELD_FORMAAT
            ws.cell(row=rij, column=9, value="Detail")

            maand_som[4] += uitgave or 0.0
            maand_som[5] += inkomst or 0.0
            maand_som[7] += excl
            maand_som[8] += btw_bedrag

            for col in (4, 7, 8):
                ws.cell(row=rij, column=col).font = Font(color=kleur)
            if band_index % 2 == 1:
                for col in range(1, 9):
                    ws.cell(row=rij, column=col).fill = band_fill
            if is_auto_nul:
                btw_cell.fill = geel
            ws.row_dimensions[rij].outline_level = 1
            band_index += 1

        subtotaal_rij = ws.max_row + 1
        ws.cell(row=subtotaal_rij, column=2, value=f"Totaal {MAAND_NAMEN[maand - 1]} {jaar}")
        for col in (4, 5, 7, 8):
            ws.cell(row=subtotaal_rij, column=col, value=round(maand_som[col], 2)).number_format = GELD_FORMAAT
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


def bouw_overzicht_sheet(wb: Workbook, buiten_periode: list, totalen: dict, buiten_kwartaal: list = None):
    ws = wb.create_sheet("Overzicht")
    buiten_kwartaal = buiten_kwartaal or []

    def banner(rij, tekst, kleur, kolommen=2, grootte=14):
        ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=kolommen)
        cell = ws.cell(row=rij, column=1, value=tekst)
        cell.font = Font(bold=True, color="FFFFFF", size=grootte)
        cell.fill = PatternFill("solid", start_color=kleur, end_color=kleur)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[rij].height = 22 if grootte > 14 else 18

    def geld(rij, label, waarde):
        ws.cell(row=rij, column=1, value=label).font = Font(bold=True)
        ws.cell(row=rij, column=2, value=round(waarde, 2)).number_format = GELD_FORMAAT

    banner(1, "Btw & Inkomsten/Uitgaven overzicht", KLEUR_HEADER_BG, kolommen=5, grootte=16)

    banner(3, "UITGAVEN", KLEUR_UITGAVE)
    geld(4, "Totaal Uitgaven Incl. Btw", totalen["uitgaven_incl"])
    geld(5, "Totaal Uitgaven Excl. Btw", totalen["uitgaven_excl"])
    geld(6, "Totaal Uitgaven Btw", totalen["uitgaven_btw"])

    banner(8, "INKOMSTEN", KLEUR_INKOMST)
    geld(9, "Totaal Inkomsten Incl. Btw", totalen["inkomsten_incl"])
    geld(10, "Totaal Inkomsten Excl. Btw", totalen["inkomsten_excl"])
    geld(11, "Totaal Inkomsten Btw", totalen["inkomsten_btw"])

    banner(13, "RESULTAAT", KLEUR_HEADER_BG)
    geld(14, "Winst (Excl. Btw)", totalen["inkomsten_excl"] - totalen["uitgaven_excl"])
    geld(15, "Btw-saldo (+ afdragen / - terugvragen)", totalen["inkomsten_btw"] - totalen["uitgaven_btw"])
    geld(16, "Btw dat van jou blijft (Uitgaven Btw - Inkomsten Btw)", totalen["uitgaven_btw"] - totalen["inkomsten_btw"])

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
    ws.cell(row=4, column=5, value=round(totalen["inkomsten_excl"], 2)).number_format = GELD_FORMAAT
    ws.cell(row=5, column=4, value="Uitgaven")
    ws.cell(row=5, column=5, value=round(totalen["uitgaven_excl"], 2)).number_format = GELD_FORMAAT

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

    rij = 22

    def waarschuwing_banner(tekst):
        nonlocal rij
        ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=5)
        banner_cel = ws.cell(row=rij, column=1, value=tekst)
        banner_cel.font = Font(bold=True, color="FFFFFF", size=12)
        banner_cel.fill = PatternFill("solid", start_color=KLEUR_WAARSCHUWING_BG, end_color=KLEUR_WAARSCHUWING_BG)
        banner_cel.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[rij].height = 20
        rij += 1

    def tekst_regel(tekst):
        nonlocal rij
        ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=5)
        cel = ws.cell(row=rij, column=1, value=tekst)
        cel.alignment = Alignment(horizontal="left")
        ws.row_dimensions[rij].height = 16
        rij += 1

    if buiten_kwartaal:
        waarschuwing_banner("NIET MEEGETELD - BOEKINGEN NA HET KWARTAALEINDE (HOREN IN DE VOLGENDE AANGIFTE)")
        for tx in sorted(buiten_kwartaal, key=lambda t: t["datum"]):
            bedrag = -tx["bedrag"] if tx["af_bij"] == "Af" else tx["bedrag"]
            tekst_regel(
                f"{tx['datum'].strftime('%d-%m-%Y')}   {tx['naam']}   "
                f"{'Af' if tx['af_bij'] == 'Af' else 'Bij'}: {_formateer_euro(bedrag)}"
            )
        rij += 1

    if buiten_periode:
        waarschuwing_banner("LET OP - FACTUUR BUITEN PERIODE VAN DIT BESTAND")
        for tx in sorted(buiten_periode, key=lambda t: t["datum"]):
            jaar, maand = tx["buiten_jaarmaand"]
            tekst_regel(
                f"{tx['datum'].strftime('%d-%m-%Y')}   {tx['naam']}   "
                f"Factuur: {tx['factuur']}  ({MAAND_NAMEN[maand - 1]} {jaar})   "
                f"Bedrag: {_formateer_euro(tx['bedrag'])}"
            )

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

    som_incl = som_excl = som_btw = 0.0
    rij = 3
    for tx in sorted(transacties, key=lambda t: t["datum"]):
        font_kleur = _kleur_voor(tx)
        is_auto_nul = tx.get("btw_auto_nul", False)
        effect = tx.get("bedrag_effect", tx["bedrag"])
        excl = round(_excl_van(effect, tx["btw_pct"]), 2)
        btw_bedrag = round(effect - excl, 2)
        som_incl += effect
        som_excl += excl
        som_btw += btw_bedrag

        ws.cell(row=rij, column=1, value=tx["datum"]).number_format = "dd-mm-yyyy"
        ws.cell(row=rij, column=2, value=_naam_weergave(tx))
        ws.cell(row=rij, column=3, value=tx["factuur"])
        bedrag_cel = ws.cell(row=rij, column=4, value=effect)
        btw_cel = ws.cell(row=rij, column=5, value=tx["btw_pct"])
        excl_cel = ws.cell(row=rij, column=6, value=excl)
        btwbedrag_cel = ws.cell(row=rij, column=7, value=btw_bedrag)
        bedrag_cel.number_format = GELD_FORMAAT
        bedrag_cel.font = Font(color=font_kleur)
        if is_auto_nul:
            btw_cel.fill = geel
        excl_cel.number_format = GELD_FORMAAT
        excl_cel.font = Font(color=font_kleur)
        btwbedrag_cel.number_format = GELD_FORMAAT
        btwbedrag_cel.font = Font(color=font_kleur)
        rij += 1

    laatste_data_rij = rij - 1
    if laatste_data_rij >= 3:
        totaal_rij = rij
        ws.cell(row=totaal_rij, column=2, value=f"Totaal {titel.lower()}").font = Font(bold=True)
        ws.cell(row=totaal_rij, column=4, value=round(som_incl, 2)).number_format = GELD_FORMAAT
        ws.cell(row=totaal_rij, column=6, value=round(som_excl, 2)).number_format = GELD_FORMAAT
        ws.cell(row=totaal_rij, column=7, value=round(som_btw, 2)).number_format = GELD_FORMAAT
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
        ("Bon zonder je naam, factuur nodig (vanaf EUR 100)", len(analyse.get("zonder_naam", []))),
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
    risico_naam = analyse.get("btw_risico_naam", 0)
    ws.cell(row=risico_rij + 3, column=1, value="Btw op bonnen zonder jouw naam (vanaf EUR 100 is een factuur op naam vereist):").font = Font(bold=True)
    rn = ws.cell(row=risico_rij + 3, column=3, value=risico_naam)
    rn.number_format = GELD_FORMAAT
    rn.font = Font(bold=True, color=(KLEUR_SLECHT_FONT if risico_naam > 0 else KLEUR_GOED_FONT))

    rij = risico_rij + 5

    from bonnen_matcher import BEWIJS_TEKST
    zeker_rijen = [
        (bon["bestand"], gebruikt_nummer or "(bedrag + datum uniek)", t["datum"], t["naam"], t["bedrag"],
         BEWIJS_TEKST.get(bon.get("bewijs", "onbekend"), "Niet gecontroleerd"))
        for bon, treffers, gebruikt_nummer in analyse["zeker"]
        for t in treffers
    ]
    rij = tabel_sectie(
        rij, "ZEKER GEMATCHT", KLEUR_ZEKER, KLEUR_ZEKER_RIJ,
        ["Bon", "Herkend nummer", "Transactiedatum", "Transactienaam", "Bedrag", "Bewijs"], zeker_rijen, "BonZeker",
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

    for col, breedte in zip("ABCDEF", [38, 22, 16, 30, 16, 30]):
        ws.column_dimensions[col].width = breedte
    return ws


def _blok_banner(ws, rij, tekst, kleur, kolommen, grootte=13):
    ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=kolommen)
    cel = ws.cell(row=rij, column=1, value=tekst)
    cel.font = Font(bold=True, color="FFFFFF", size=grootte)
    cel.fill = PatternFill("solid", start_color=kleur, end_color=kleur)
    cel.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[rij].height = 22 if grootte >= 14 else 18


def bouw_aangifte_sheet(wb: Workbook, aangifte: dict, totalen: dict, transacties: list):
    """De cijfers die je letterlijk overtikt in Mijn Belastingdienst Zakelijk.

    Momentopname: de kern-rubrieken (1a/5a/5b/5c) zijn vooraf uitgerekende waarden
    uit het model, geen Excel-formules meer. Bewerken doe je in de app; de Excel is
    een vaste snapshot. Verlegde btw en KOR worden apart en met een controleer-notitie
    getoond (fiscaal maatwerk).
    """
    ws = wb.create_sheet("BTW-aangifte")
    # Dezelfde afleiding als de app-tab (berekening.bepaal_rubrieken): vooraf uitgerekende
    # waarden, geen formules.
    rub = bepaal_rubrieken(transacties)
    vv_excl = rub["verlegd_verkoop_excl"]
    vi_excl = rub["verlegd_inkoop_excl"]
    vi_btw = rub["verlegd_inkoop_btw"]
    btw_5a, btw_5b, btw_5c = rub["btw_5a"], rub["btw_5b"], rub["btw_5c"]
    deadline_tekst = aangifte["deadline_tekst"]
    deadline_is_controle = deadline_tekst.lower().startswith("controleer")

    _blok_banner(ws, 1, "BTW-AANGIFTE  —  over te tikken in Mijn Belastingdienst Zakelijk", KLEUR_HEADER_BG, 3, grootte=15)
    ws.cell(row=2, column=1, value=aangifte["periode_tekst"]).font = Font(italic=True, color="595959")
    deadline_regel = (
        deadline_tekst.capitalize()
        if deadline_is_controle
        else f"Aangifte{' + betaling' if btw_5c > 0 else ''} {deadline_tekst}"
    )
    ws.cell(row=3, column=1, value=deadline_regel).font = Font(italic=True, color="595959")

    if aangifte["kor_actief"]:
        _blok_banner(ws, 5, "KLEINEONDERNEMERSREGELING (KOR) ACTIEF", KLEUR_WAARSCHUWING_BG, 3, grootte=13)
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

    def rubriek(label, omzet_waarde, btw_waarde, vet=False):
        nonlocal rij
        ws.cell(row=rij, column=1, value=label).font = Font(bold=vet)
        if omzet_waarde is not None:
            c = ws.cell(row=rij, column=2, value=omzet_waarde)
            c.number_format = GELD_FORMAAT
        if btw_waarde is not None:
            c = ws.cell(row=rij, column=3, value=btw_waarde)
            c.number_format = GELD_FORMAAT
            c.font = Font(bold=vet)
        huidige = rij
        rij += 1
        return huidige

    rubriek("1a  Leveringen/diensten belast met 21%", rub["omzet_1a"], rub["btw_1a"])
    if rub["heeft_1b"]:
        rubriek("1b  Leveringen/diensten belast met 9%", rub["omzet_1b"], rub["btw_1b"])
    if rub["heeft_1e"]:
        rubriek("1e  Omzet waarbij btw naar afnemer is verlegd (0%)", vv_excl, 0)
    if rub["heeft_overige"]:
        rubriek("Overige omzet 0% / anders (controleer rubriek)", rub["overige_excl"], rub["overige_btw"])
    # Naar jou verlegde btw per rubriek (Toelichting btw-aangifte 2026): binnenland 2a, buiten EU 4a, EU 4b.
    if rub["heeft_2a"]:
        rubriek("2a  Btw naar mij verlegd (binnenland, onderaanneming)", rub["verlegd_2a_excl"], rub["verlegd_2a_btw"])
    if rub["heeft_4a"]:
        rubriek("4a  Inkoop van buiten de EU (btw verlegd)", rub["verlegd_4a_excl"], rub["verlegd_4a_btw"])
    if rub["heeft_4b"]:
        rubriek("4b  Inkoop uit andere EU-landen (btw verlegd)", rub["verlegd_4b_excl"], rub["verlegd_4b_btw"])

    rubriek("5a  Verschuldigde omzetbelasting", None, btw_5a, vet=True)
    rubriek("5b  Voorbelasting (terug te vragen)", None, btw_5b, vet=True)
    rij_5c = rubriek("5c  Te betalen (+) / terug te vragen (-)", None, btw_5c, vet=True)
    if rub["btw_niet_onderbouwd"]:
        ws.cell(row=rij, column=1, value=(
            f"Niet in 5b: {_formateer_euro(rub['btw_niet_onderbouwd'])} btw op uitgaven zonder gekoppelde bon of factuur "
            "(art. 15 Wet OB: aftrek alleen met factuur). Koppel de bonnen in de app om dit mee te tellen."
        )).font = Font(italic=True, color=KLEUR_SLECHT_FONT)
        ws.merge_cells(start_row=rij, start_column=1, end_row=rij, end_column=3)
        rij += 1

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

    if rub["heeft_1e"] or rub["heeft_2a"] or rub["heeft_4a"] or rub["heeft_4b"]:
        ws.cell(row=rij + 1, column=1, value=(
            "Let op: verlegde btw is fiscaal maatwerk — controleer 1e/2a/4a/4b met je boekhouder."
        )).font = Font(italic=True, color=KLEUR_SLECHT_FONT)
        rij += 1

    # Betaal- of teruggaafpositie.
    potje_rij = rij + 2
    terug_te_vragen = btw_5c <= 0
    _blok_banner(
        ws,
        potje_rij,
        "TERUG TE VRAGEN BTW" if terug_te_vragen else "TE RESERVEREN VOOR DE AANGIFTE",
        KLEUR_HEADER_BG if terug_te_vragen else KLEUR_INKOMST,
        3,
        grootte=13,
    )
    ws.cell(
        row=potje_rij + 1,
        column=1,
        value="Saldo volgens dit bestand:" if terug_te_vragen else "Reserveer voor betaling:",
    ).font = Font(bold=True)
    reserve = ws.cell(row=potje_rij + 1, column=3, value=round(abs(btw_5c), 2))
    reserve.number_format = GELD_FORMAAT
    reserve.font = Font(
        bold=True,
        size=12,
        color=KLEUR_HEADER_BG if terug_te_vragen else KLEUR_INKOMST,
    )
    if deadline_is_controle:
        potje_uitleg = deadline_tekst.capitalize()
    elif terug_te_vragen:
        potje_uitleg = f"Dien de aangifte {deadline_tekst} in."
    else:
        potje_uitleg = f"Dien de aangifte in en betaal {deadline_tekst}."
    ws.cell(row=potje_rij + 2, column=1, value=potje_uitleg).font = Font(italic=True, color="595959")

    # Het model bevat alleen de omzet uit het geladen bankbestand. Presenteer die niet
    # als volledige jaaromzet; de KOR-grens geldt voor het hele kalenderjaar.
    jaaromzet = aangifte["jaaromzet_tot_nu"]
    if aangifte["kor_actief"] or jaaromzet <= 1.5 * KOR_GRENS:
        jaar_rij = potje_rij + 4
        _blok_banner(ws, jaar_rij, "OMZET IN DIT BANKBESTAND · KOR-GRENS € 20.000", KLEUR_HEADER_BG, 3, grootte=12)
        ws.cell(row=jaar_rij + 1, column=1, value="Omzet excl. btw in dit bestand:").font = Font(bold=True)
        jc = ws.cell(row=jaar_rij + 1, column=3, value=round(jaaromzet, 2))
        jc.number_format = GELD_FORMAAT
        ws.cell(
            row=jaar_rij + 2,
            column=1,
            value="De KOR-grens geldt voor de totale omzet per kalenderjaar; tel ook omzet buiten dit bestand mee.",
        ).font = Font(italic=True, color="595959")
        if not aangifte["kor_actief"] and 0.9 * KOR_GRENS <= jaaromzet <= KOR_GRENS:
            ws.cell(row=jaar_rij + 3, column=1, value="Dit bestand ligt dicht bij de jaargrens; controleer de volledige kalenderjaaromzet.").font = Font(italic=True, color=KLEUR_SLECHT_FONT)
        if aangifte["kor_actief"] and jaaromzet > KOR_GRENS:
            ws.cell(row=jaar_rij + 3, column=1, value="Controleer direct of de KOR-grens in het kalenderjaar is overschreden.").font = Font(bold=True, color=KLEUR_SLECHT_FONT)

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


def bouw_workbook(transacties: list, buiten_periode: list, aangifte: dict, config: dict,
                  buiten_kwartaal: list = None) -> Workbook:
    wb = Workbook()
    # Vaste momentopname: alle bedragen staan als uitgerekende waarden in de cellen
    # (geen Excel-formules). Bewerken doe je in de app, niet in de Excel.
    totalen = _totalen(transacties)
    bouw_transacties_sheet(wb, transacties)
    bouw_overzicht_sheet(wb, buiten_periode, totalen, buiten_kwartaal)
    bouw_aangifte_sheet(wb, aangifte, totalen, transacties)
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]
    inkomsten = [t for t in transacties if t["richting"] == "inkomst"]
    bouw_richting_sheet(wb, "Uitgaven", KLEUR_UITGAVE, uitgaven)
    bouw_richting_sheet(wb, "Inkomsten", KLEUR_INKOMST, inkomsten)
    investeringen = [t for t in uitgaven if t.get("is_investering")]
    bouw_investeringen_sheet(wb, investeringen, config["investering_drempel"])
    bouw_leveranciers_sheet(wb, uitgaven)
    return wb


def exporteer_excel(model: Boekhouding, output_pad: Path, aangifte_override: dict = None) -> Workbook:
    """Bouwt de Excel uit het model (pure consument) en slaat 'm op."""
    aangifte = aangifte_override if aangifte_override is not None else model.aangifte
    wb = bouw_workbook(model.transacties, model.buiten_periode, aangifte, model.config,
                       buiten_kwartaal=getattr(model, "buiten_kwartaal", None))
    wb.save(output_pad)
    return wb


def exporteer_csv(model: Boekhouding, output_pad: Path):
    """Boekhouder-import-CSV uit het model. Retourneert het pad, of None bij niets/fout."""
    if not (model.config.get("boekhoud_export", True) and model.transacties):
        return None
    try:
        import boekhoud_export
        export_pad = output_pad.with_name(output_pad.stem.replace("_btw_overzicht", "") + "_boekhouder_import.csv")
        boekhoud_export.schrijf_export(model.transacties, export_pad)
        return export_pad
    except Exception as fout:
        print(f"Waarschuwing: boekhoud-export mislukt ({fout})", file=sys.stderr)
        return None


def verwerk_bestand(csv_pad: Path, output_pad: Path = None, overrides: dict = None) -> dict:
    """Berekent het model en exporteert Excel + boekhouder-CSV. Retourneert een resultaten-dict.

    overrides=None laadt overrides.json; geef {} mee voor een schone run (regressietest).
    Gooit SystemExit met een duidelijke Nederlandse melding bij een ongeldig bestand.
    """
    if not csv_pad.exists():
        raise SystemExit(f"Bestand niet gevonden: {csv_pad}")
    output_pad = output_pad or csv_pad.with_name(f"{csv_pad.stem}_btw_overzicht.xlsx")

    model = bereken_model(csv_pad, overrides=overrides)
    schrijf_leveranciers_geheugen(programma_map(), model.geheugen)

    wb = exporteer_excel(model, output_pad)
    export_pad = exporteer_csv(model, output_pad)

    resultaat = model.resultaat()
    resultaat["output_pad"] = output_pad
    resultaat["export_pad"] = export_pad
    resultaat["workbook"] = wb
    resultaat["model"] = model
    return resultaat


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
