#!/usr/bin/env python3
"""Pure rekenkern: btw, bedragen en aangifte-cijfers.

Geen openpyxl, geen Excel, geen GUI — alleen getallen in, getallen uit. Zo is de logica
testbaar en kunnen zowel de Excel-export als (straks) de app-tabs uit dezelfde bron putten.
"""
from datetime import datetime

STANDAARD_BTW = 21
KOR_GRENS = 20000.0
MAAND_NAMEN = [
    "januari", "februari", "maart", "april", "mei", "juni",
    "juli", "augustus", "september", "oktober", "november", "december",
]


def excl_van(bedrag: float, btw_pct: int) -> float:
    """Bedrag exclusief btw. Vangt een onmogelijk tarief (-100%) af i.p.v. te crashen."""
    noemer = 1 + btw_pct / 100
    if noemer == 0:
        return bedrag
    return bedrag / noemer


# Compat-alias voor bestaande aanroepen.
_excl_van = excl_van


def totalen_per_richting(transacties: list) -> dict:
    """Incl/excl/btw opgeteld per richting, als kale getallen (voor Excel-snapshot en app-tabs)."""
    def incl(t):
        return t.get("bedrag_effect", t["bedrag"])

    def excl(t):
        return excl_van(incl(t), t["btw_pct"])

    ink = [t for t in transacties if t["richting"] == "inkomst"]
    uit = [t for t in transacties if t["richting"] == "uitgave"]
    return {
        "inkomsten_incl": round(sum(incl(t) for t in ink), 2),
        "inkomsten_excl": round(sum(excl(t) for t in ink), 2),
        "inkomsten_btw": round(sum(incl(t) - excl(t) for t in ink), 2),
        "uitgaven_incl": round(sum(incl(t) for t in uit), 2),
        "uitgaven_excl": round(sum(excl(t) for t in uit), 2),
        "uitgaven_btw": round(sum(incl(t) - excl(t) for t in uit), 2),
    }


def bepaal_btw(transactie: dict, leverancier: str, config: dict, geheugen: dict):
    """Btw-percentage + herkomst voor een transactie.

    Volgorde (config/wettelijke regels gaan VOOR een onthouden gok):
    verlegde btw (0%) -> automatische 0%-partij (config) -> onthouden keuze
    (leveranciers.json, alleen een plausibel 0-100 tarief) -> standaard 21%.
    Retourneert (btw_pct, bron, is_verlegd, is_auto_nul).
    """
    naam_lower = transactie["naam"].lower()

    if any(patroon in naam_lower for patroon in config.get("verlegd_patronen", [])):
        return 0, "verlegd", True, False

    if any(patroon in naam_lower for patroon in config.get("nul_btw_patronen", [])):
        return 0, "nul-auto", False, True

    onthouden = geheugen.get(leverancier)
    if onthouden and isinstance(onthouden.get("btw"), (int, float)):
        onthouden_btw = int(onthouden["btw"])
        if 0 <= onthouden_btw <= 100:
            return onthouden_btw, "geheugen", False, onthouden_btw == 0

    return STANDAARD_BTW, "standaard", False, False


def bepaal_aangifte_info(verwerkt: list, config: dict) -> dict:
    """Alle geaggregeerde aangifte-cijfers uit de verwerkte transacties."""
    inkomsten = [t for t in verwerkt if t["richting"] == "inkomst"]
    uitgaven = [t for t in verwerkt if t["richting"] == "uitgave"]

    def excl(t):
        return excl_van(t.get("bedrag_effect", t["bedrag"]), t["btw_pct"])

    def btw(t):
        return t.get("bedrag_effect", t["bedrag"]) - excl(t)

    inkomsten_excl = sum(excl(t) for t in inkomsten)
    inkomsten_btw = sum(btw(t) for t in inkomsten)
    uitgaven_excl = sum(excl(t) for t in uitgaven)
    uitgaven_btw = sum(btw(t) for t in uitgaven)
    verlegd_verkoop_excl = round(sum(excl(t) for t in inkomsten if t.get("is_verlegd")), 2)
    verlegd_inkoop_excl = round(sum(excl(t) for t in uitgaven if t.get("is_verlegd")), 2)
    overige_omzet_excl = round(
        sum(excl(t) for t in inkomsten if not t.get("is_verlegd") and t["btw_pct"] != 21), 2
    )

    datums = [t["datum"] for t in verwerkt] or [datetime.now()]
    periode_start, periode_eind = min(datums), max(datums)
    volgende_maand = periode_eind.month % 12 + 1
    volgende_jaar = periode_eind.year + (1 if periode_eind.month == 12 else 0)
    deadline_tekst = f"uiterlijk eind {MAAND_NAMEN[volgende_maand - 1]} {volgende_jaar} (1 maand na kwartaaleinde)"
    periode_tekst = f"Periode {periode_start.strftime('%d-%m-%Y')} t/m {periode_eind.strftime('%d-%m-%Y')}"

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


# Compat-alias voor bestaande aanroepen.
_bepaal_aangifte_info = bepaal_aangifte_info
