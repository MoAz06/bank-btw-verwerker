#!/usr/bin/env python3
"""Genereert een neutraal import-bestand voor de boekhouder.

Een kale, gestructureerde CSV (puntkomma-gescheiden, Nederlandse komma-decimalen)
die de boekhouder in zijn pakket kan inlezen — datum, tegenpartij, bedragen, btw-code.
Zo bespaar je hem inboekwerk zonder zelf een boekhoudpakket na te bouwen.
"""
import csv
from pathlib import Path


def _euro(bedrag: float) -> str:
    return f"{bedrag:.2f}".replace(".", ",")


def _btw_code(tx: dict) -> str:
    if tx.get("is_verlegd"):
        return "VERLEGD"
    if tx["btw_pct"] == 21:
        return "HOOG_21"
    if tx["btw_pct"] == 9:
        return "LAAG_9"
    if tx["btw_pct"] == 0:
        return "GEEN_0"
    return f"OVERIG_{tx['btw_pct']}"


KOLOMMEN = [
    "Datum", "Type", "Tegenpartij", "Factuurnummer",
    "Bedrag incl", "Btw %", "Bedrag excl", "Btw-bedrag", "Btw-code",
    "Tegenrekening", "Investering",
]


def schrijf_export(transacties: list, pad: Path):
    with open(pad, "w", newline="", encoding="utf-8-sig") as f:
        schrijver = csv.writer(f, delimiter=";")
        schrijver.writerow(KOLOMMEN)
        for tx in sorted(transacties, key=lambda t: t["datum"]):
            bedrag = tx.get("bedrag_effect", tx["bedrag"])
            noemer = 1 + tx["btw_pct"] / 100
            excl = bedrag / noemer if noemer else bedrag
            btw_bedrag = bedrag - excl
            soort = "Inkomst" if tx["richting"] == "inkomst" else ("Creditnota" if tx.get("is_kredietfactuur") else "Uitgave")
            schrijver.writerow([
                tx["datum"].strftime("%d-%m-%Y"),
                soort,
                tx.get("leverancier", tx["naam"]),
                tx.get("factuur", ""),
                _euro(bedrag),
                tx["btw_pct"],
                _euro(excl),
                _euro(btw_bedrag),
                _btw_code(tx),
                tx.get("tegenrekening", ""),
                "Ja" if tx.get("is_investering") else "",
            ])
