#!/usr/bin/env python3
"""Leverancier-normalisatie en een lokaal geheugen van je btw-keuzes per leverancier.

De bank-namen zijn rommelig ("CCV*VERPLOEGEN VERWARMING NLD", "Bouwmaat Den Haag
SGRAVENHAGE"). Normalisatie brengt varianten van dezelfde leverancier terug tot een
nette canonieke naam, zodat je per leverancier kan optellen en de app je btw-keuzes
kan onthouden voor volgende kwartalen (leveranciers.json).
"""
import json
import re
from pathlib import Path

# Betaalautomaat/terminal-prefixen die de ING voor de echte naam plakt.
_TERMINAL_PREFIXEN = re.compile(r"^(?:CCV\*|BCK\*|C&M\*|ZTL\*|CCV|BEA|GEA)\s*", re.IGNORECASE)

# Plaats-/land-achtige staartwoorden die niets over de leverancier zeggen.
_PLAATS_STAART = {
    "nld", "sgravenhage", "s gravenhage", "gravenhage", "den haag", "wateringen",
    "naaldwijk", "moordrecht", "monster", "helmond", "huizen", "rijswijk",
    "zoetermeer", "delft", "westland", "ypenburg", "binckhorst", "amstelveen",
    "deventer", "twello", "almere", "amsterdam", "rotterdam", "nederland",
}

# Standaard-aliassen: als dit patroon (kleine letters) in de banknaam voorkomt,
# wordt het de canonieke leveranciersnaam. Uitbreidbaar via config (leverancier_aliassen).
STANDAARD_ALIASSEN = {
    "bouwmaat": "Bouwmaat",
    "hornbach": "Hornbach",
    "technische unie": "Technische Unie",
    "wasco": "Wasco",
    "verploegen": "Verploegen Verwarming",
    "travelcard": "Travelcard",
    "kpn": "KPN",
    "toolstation": "Toolstation",
    "proforto": "Proforto",
    "hbm machines": "HBM Machines",
    "afvalcontainer": "De Ned. Afvalcontainer Centrale",
    "amazon": "Amazon",
    "praxis": "Praxis",
    "shell": "Shell",
    "esso": "Esso",
    "sanitino": "Sanitino",
    "mollie": "Mollie Payments",
    "multisafepay": "MultiSafepay",
    "stripe": "Stripe",
    "paypal": "PayPal",
    "riverty": "Riverty (parkeren)",
    "easypark": "Riverty (parkeren)",
    "hostnet": "Hostnet",
    "microsoft": "Microsoft",
    "vwpfs": "VWPFS (lease)",
    "hiltermann": "Hiltermann Lease",
    "wsdhy": "Welkoop Ypenburg",
}


def _titel(tekst: str) -> str:
    kleine = {"van", "de", "der", "den", "en", "via", "b.v.", "bv"}
    woorden = tekst.split()
    return " ".join(w if w.lower() in kleine else w.capitalize() for w in woorden)


def normaliseer_leverancier(naam: str, extra_aliassen: dict = None) -> str:
    """Brengt een rommelige banknaam terug tot een nette canonieke leveranciersnaam."""
    if not naam:
        return "Onbekend"
    aliassen = {**STANDAARD_ALIASSEN, **(extra_aliassen or {})}
    naam_lower = naam.lower()
    # Woordgrens-match: "esso"/"shell"/"amazon" mogen NIET binnen andere woorden matchen.
    for patroon, canoniek in aliassen.items():
        if re.search(r"(?<![a-z0-9])" + re.escape(patroon) + r"(?![a-z0-9])", naam_lower):
            return canoniek

    schoon = _TERMINAL_PREFIXEN.sub("", naam).strip()
    # Staartwoorden (plaats/land) achteraan wegknippen.
    while True:
        woorden = schoon.split()
        if len(woorden) <= 1:
            break
        laatste_een = woorden[-1].lower().strip(".,")
        laatste_twee = " ".join(woorden[-2:]).lower().strip(".,")
        if laatste_twee in _PLAATS_STAART:
            schoon = " ".join(woorden[:-2])
        elif laatste_een in _PLAATS_STAART:
            schoon = " ".join(woorden[:-1])
        else:
            break
    schoon = re.sub(r"\s+", " ", schoon).strip()
    return _titel(schoon) if schoon else naam.strip()


def geheugen_pad(map_pad: Path) -> Path:
    return map_pad / "leveranciers.json"


def laad_leveranciers_geheugen(map_pad: Path) -> dict:
    """Retourneert {leverancier: {"btw": int, "bon_verwacht": bool, "categorie": str}}."""
    pad = geheugen_pad(map_pad)
    if not pad.exists():
        return {}
    try:
        with open(pad, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def schrijf_leveranciers_geheugen(map_pad: Path, geheugen: dict):
    pad = geheugen_pad(map_pad)
    try:
        with open(pad, "w", encoding="utf-8") as f:
            json.dump(geheugen, f, ensure_ascii=False, indent=2, sort_keys=True)
    except OSError:
        pass


def vul_geheugen_aan(geheugen: dict, leverancier: str, btw_pct: int, bon_verwacht: bool) -> bool:
    """Voegt een leverancier toe aan het geheugen als die er nog niet in staat.

    Retourneert True als het een nieuwe (nog onbekende) leverancier was.
    """
    if leverancier in geheugen:
        return False
    geheugen[leverancier] = {"btw": btw_pct, "bon_verwacht": bon_verwacht}
    return True


def onthoud_btw(geheugen: dict, leverancier: str, btw_pct: int):
    """Overschrijft de onthouden btw voor een leverancier (handmatige keuze wint).

    Anders dan vul_geheugen_aan (die alleen gaten vult) legt dit een expliciete
    correctie vast, zodat dezelfde leverancier voortaan dat tarief krijgt.
    """
    if not leverancier:
        return
    bestaand = geheugen.get(leverancier, {})
    geheugen[leverancier] = {**bestaand, "btw": int(btw_pct)}
