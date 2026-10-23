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


# Soorten verlegging en de aangifterubriek waarin de naar jou verlegde btw hoort
# (Toelichting btw-aangifte 2026): binnenland -> 2a, EU-inkoop -> 4b, buiten de EU -> 4a.
VERLEGGING_SOORTEN = ("binnenland", "eu", "buiten_eu")
VERLEGGING_RUBRIEK = {"binnenland": "2a", "eu": "4b", "buiten_eu": "4a"}


def bepaal_verlegging(transactie: dict, config: dict):
    """Welke verlegging op een transactie van toepassing is volgens de config-patronen, of None."""
    naam_lower = transactie["naam"].lower()
    for soort, sleutel in (("binnenland", "verlegd_patronen"), ("eu", "eu_verlegd_patronen"),
                           ("buiten_eu", "buiten_eu_verlegd_patronen")):
        if any(patroon in naam_lower for patroon in config.get(sleutel, [])):
            return soort
    return None


def bepaal_btw(transactie: dict, leverancier: str, config: dict, geheugen: dict):
    """Btw-percentage + herkomst voor een transactie.

    Volgorde (config/wettelijke regels gaan VOOR een onthouden gok):
    verlegde btw (0%) -> automatische 0%-partij (config) -> onthouden keuze
    (leveranciers.json, alleen een plausibel 0-100 tarief) -> standaard 21%.
    Retourneert (btw_pct, bron, is_verlegd, is_auto_nul).
    """
    naam_lower = transactie["naam"].lower()

    if bepaal_verlegging(transactie, config):
        return 0, "verlegd", True, False

    if any(patroon in naam_lower for patroon in config.get("nul_btw_patronen", [])):
        return 0, "nul-auto", False, True

    onthouden = geheugen.get(leverancier)
    if onthouden and isinstance(onthouden.get("btw"), (int, float)):
        onthouden_btw = int(onthouden["btw"])
        if 0 <= onthouden_btw <= 100:
            return onthouden_btw, "geheugen", False, onthouden_btw == 0

    return STANDAARD_BTW, "standaard", False, False


def bepaal_rubrieken(transacties: list) -> dict:
    """De aangifte-rubrieken (1a, 1b, 1e, 4b, 5a, 5b, 5c) als kale getallen.

    De ENE afleiding die zowel de app-tab als de Excel-tab gebruiken, zodat die nooit
    uiteen kunnen lopen. Alles op basis van bedrag_effect (kredietfacturen negatief).
    """
    tot = totalen_per_richting(transacties)

    def incl(t):
        return t.get("bedrag_effect", t["bedrag"])

    def excl(t):
        return excl_van(incl(t), t["btw_pct"])

    inkomsten = [t for t in transacties if t["richting"] == "inkomst"]
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]
    gewoon = [t for t in inkomsten if not t.get("is_verlegd")]

    omzet_1a = round(sum(excl(t) for t in gewoon if t["btw_pct"] == 21), 2)
    btw_1a = round(sum(incl(t) - excl(t) for t in gewoon if t["btw_pct"] == 21), 2)
    omzet_1b = round(sum(excl(t) for t in gewoon if t["btw_pct"] == 9), 2)
    btw_1b = round(sum(incl(t) - excl(t) for t in gewoon if t["btw_pct"] == 9), 2)
    verlegd_verkoop_excl = round(sum(excl(t) for t in inkomsten if t.get("is_verlegd")), 2)

    # Naar jou verlegde btw per rubriek: binnenland (onderaannemer) 2a, EU-inkoop 4b, buiten EU 4a.
    # Je geeft de btw aan in die rubriek en trekt 'm in dezelfde aangifte in 5b weer af.
    verlegd = {}
    for soort, rubriek in VERLEGGING_RUBRIEK.items():
        excl_som = round(sum(excl(t) for t in uitgaven if t.get("is_verlegd") and t.get("verlegging", "binnenland") == soort), 2)
        verlegd[rubriek] = {"excl": excl_som, "btw": round(excl_som * STANDAARD_BTW / 100, 2)}
    verlegd_inkoop_excl = round(sum(v["excl"] for v in verlegd.values()), 2)
    verlegd_inkoop_btw = round(sum(v["btw"] for v in verlegd.values()), 2)

    # Rest van de omzet (0% of een afwijkend tarief) incl. de btw die daar eventueel op zit.
    overige_excl = round(tot["inkomsten_excl"] - omzet_1a - omzet_1b - verlegd_verkoop_excl, 2)
    overige_btw = round(tot["inkomsten_btw"] - btw_1a - btw_1b, 2)

    # Voorbelasting (5b) alleen voor uitgaven waarvan de factuur of bon er is (art. 15 Wet OB):
    # 'aftrekbaar' wordt door het model gezet. Kredietfacturen verlagen de aftrek altijd.
    def aftrekbaar(t):
        return t.get("aftrekbaar", True) or t.get("is_kredietfactuur") or incl(t) < 0

    btw_onderbouwd = round(sum(incl(t) - excl(t) for t in uitgaven if aftrekbaar(t)), 2)
    btw_niet_onderbouwd = round(sum(incl(t) - excl(t) for t in uitgaven if not aftrekbaar(t)), 2)

    btw_5a = round(tot["inkomsten_btw"] + verlegd_inkoop_btw, 2)
    btw_5b = round(btw_onderbouwd + verlegd_inkoop_btw, 2)
    btw_5c = round(btw_5a - btw_5b, 2)
    return {
        "omzet_1a": omzet_1a,
        "btw_1a": btw_1a,
        "omzet_1b": omzet_1b,
        "btw_1b": btw_1b,
        "heeft_1b": omzet_1b != 0 or btw_1b != 0,
        "verlegd_verkoop_excl": verlegd_verkoop_excl,
        "heeft_1e": verlegd_verkoop_excl != 0,
        "overige_excl": overige_excl,
        "overige_btw": overige_btw,
        "heeft_overige": overige_excl != 0 or overige_btw != 0,
        "verlegd_inkoop_excl": verlegd_inkoop_excl,
        "verlegd_inkoop_btw": verlegd_inkoop_btw,
        "verlegd_2a_excl": verlegd["2a"]["excl"],
        "verlegd_2a_btw": verlegd["2a"]["btw"],
        "heeft_2a": verlegd["2a"]["excl"] != 0,
        "verlegd_4a_excl": verlegd["4a"]["excl"],
        "verlegd_4a_btw": verlegd["4a"]["btw"],
        "heeft_4a": verlegd["4a"]["excl"] != 0,
        "verlegd_4b_excl": verlegd["4b"]["excl"],
        "verlegd_4b_btw": verlegd["4b"]["btw"],
        "heeft_4b": verlegd["4b"]["excl"] != 0,
        "btw_onderbouwd": btw_onderbouwd,
        "btw_niet_onderbouwd": btw_niet_onderbouwd,
        "btw_5a": btw_5a,
        "btw_5b": btw_5b,
        "btw_5c": btw_5c,
        "reserveren": round(max(0.0, btw_5c), 2),
    }


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
