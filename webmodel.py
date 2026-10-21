#!/usr/bin/env python3
"""Zet een Boekhouding-model om naar kale, JSON-klare data voor de app-tabs.

Alleen getallen, strings en lijsten — geen datetime, geen openpyxl. De front-end (app.js)
verzorgt alle opmaak (euro's, datums, kleuren) zodat opmaak op één plek zit en toegankelijk
te sturen is. Deze module deelt de rekenkern (berekening.py) met de Excel-export, dus de
cijfers in de app en in de gedownloade Excel zijn per definitie gelijk.
"""
import re
import statistics
from calendar import monthrange
from datetime import date, timedelta

from berekening import bepaal_rubrieken, excl_van, totalen_per_richting
from boekhouding import bepaal_kwartaal_venster


def _iso(d):
    return d.strftime("%Y-%m-%d")


def tx_dict(t: dict) -> dict:
    """Eén transactie als platte dict met vooruitberekende excl/btw."""
    effect = t.get("bedrag_effect", t["bedrag"])
    excl = round(excl_van(effect, t["btw_pct"]), 2)
    return {
        "id": t.get("tx_id"),
        "datum": _iso(t["datum"]),
        "naam": t["naam"],
        "leverancier": t.get("leverancier"),
        "factuur": t.get("factuur", ""),
        "tegenrekening": t.get("tegenrekening", ""),
        "mededelingen": t.get("mededelingen", ""),
        "richting": t["richting"],
        "btw_pct": t["btw_pct"],
        "btw_bron": t.get("btw_bron"),
        "bedrag": round(effect, 2),
        "excl": excl,
        "btw": round(effect - excl, 2),
        "is_investering": bool(t.get("is_investering")),
        "is_verlegd": bool(t.get("is_verlegd")),
        "is_kredietfactuur": bool(t.get("is_kredietfactuur")),
        "auto_nul": bool(t.get("btw_auto_nul")),
        "nieuw": bool(t.get("is_nieuwe_leverancier")),
        # Handmatig bewerkte velden met hun oorspronkelijke waarde (voor markering + herstel).
        "bewerkt": {k: (round(v, 2) if isinstance(v, float) else v) for k, v in (t.get("bewerkt") or {}).items()},
    }


def _leveranciers_overzicht(uitgaven: list) -> dict:
    per = {}
    for t in uitgaven:
        per.setdefault(t["leverancier"], []).append(t)

    rijen = []
    for leverancier, txs in per.items():
        bedragen = [t.get("bedrag_effect", t["bedrag"]) for t in txs]
        datums = [t["datum"] for t in txs]
        rijen.append({
            "leverancier": leverancier,
            "aantal": len(txs),
            "totaal": round(sum(bedragen), 2),
            "gemiddeld": round(sum(bedragen) / len(bedragen), 2),
            "eerste": _iso(min(datums)),
            "laatste": _iso(max(datums)),
        })
    rijen.sort(key=lambda r: -r["totaal"])

    # Opvallende bedragen: sterk boven de mediaan van die leverancier (zelfde regel als Excel).
    afwijkingen = []
    for leverancier, txs in per.items():
        if len(txs) < 4:
            continue
        bedragen = [t.get("bedrag_effect", t["bedrag"]) for t in txs]
        mediaan = statistics.median(bedragen)
        for t in txs:
            bedrag = t.get("bedrag_effect", t["bedrag"])
            if mediaan > 0 and bedrag >= 3 * mediaan and bedrag - mediaan >= 100:
                afwijkingen.append({
                    "id": t.get("tx_id"),
                    "datum": _iso(t["datum"]),
                    "leverancier": leverancier,
                    "bedrag": round(bedrag, 2),
                    "mediaan": round(mediaan, 2),
                })
    afwijkingen.sort(key=lambda a: -a["bedrag"])
    return {"rijen": rijen, "afwijkingen": afwijkingen}


_BESTANDSPERIODE = re.compile(
    r"(?<!\d)(\d{2})-(\d{2})-(\d{4})_(\d{2})-(\d{2})-(\d{4})(?!\d)"
)


def _periode_dict(model) -> dict:
    """Exacte bronperiode en alleen bij voldoende zekerheid een kwartaallabel."""
    match = _BESTANDSPERIODE.search(model.bestandsnaam or "")
    bron_herkend = False
    if match:
        try:
            d1, m1, j1, d2, m2, j2 = map(int, match.groups())
            start, eind = date(j1, m1, d1), date(j2, m2, d2)
            bron_herkend = True
        except ValueError:
            match = None
    if not bron_herkend:
        datums = [t["datum"].date() if hasattr(t["datum"], "date") else t["datum"] for t in model.ruwe]
        start = min(datums)
        eind = max(datums)
    if eind < start:
        start, eind = eind, start
    bron_start, bron_eind = start, eind

    kwartaal = (start.month - 1) // 3 + 1
    kwartaal_start = date(start.year, 1 + (kwartaal - 1) * 3, 1)
    if kwartaal == 4:
        volgende_start = date(start.year + 1, 1, 1)
    else:
        volgende_start = date(start.year, kwartaal_start.month + 3, 1)
    kwartaal_eind = volgende_start - timedelta(days=1)

    # Dezelfde regel als het model gebruikt om boekingen na het kwartaaleinde apart te
    # zetten (boekhouding.bepaal_kwartaal_venster), zodat label en cijfers altijd kloppen.
    venster = bepaal_kwartaal_venster(model.bestandsnaam or "") if bron_herkend else None
    volledig_kwartaal = venster is not None
    if volledig_kwartaal:
        kwartaal_start, kwartaal_eind = venster[0].date(), venster[1].date()
        kwartaal = (kwartaal_start.month - 1) // 3 + 1
        volgende_start = kwartaal_eind + timedelta(days=1)
    deadline = None
    kwartaal_label = None
    if volledig_kwartaal:
        kwartaal_label = f"Q{kwartaal} {start.year}"
        deadline_maand = volgende_start.month
        deadline_jaar = volgende_start.year
        deadline = date(
            deadline_jaar,
            deadline_maand,
            monthrange(deadline_jaar, deadline_maand)[1],
        )

    getoonde_start = kwartaal_start if volledig_kwartaal else bron_start
    getoonde_eind = kwartaal_eind if volledig_kwartaal else bron_eind
    return {
        "start": getoonde_start.isoformat(),
        "eind": getoonde_eind.isoformat(),
        "bron_start": bron_start.isoformat(),
        "bron_eind": bron_eind.isoformat(),
        "bron_herkend": bron_herkend,
        "volledig_kwartaal": volledig_kwartaal,
        "kwartaal": kwartaal if volledig_kwartaal else None,
        "jaar": start.year if volledig_kwartaal else None,
        "label": kwartaal_label or "Aangepaste periode",
        "deadline": deadline.isoformat() if deadline else None,
    }


def _aangifte_dict(model, periode: dict) -> dict:
    """Aangifte-rubrieken als kale getallen (zelfde afleiding als de Excel-tab)."""
    a = model.aangifte
    rubrieken = bepaal_rubrieken(model.transacties)

    return {
        **rubrieken,
        "kor_actief": a["kor_actief"],
        "periode_tekst": (
            "Periode "
            f"{date.fromisoformat(periode['start']).strftime('%d-%m-%Y')} "
            "t/m "
            f"{date.fromisoformat(periode['eind']).strftime('%d-%m-%Y')}"
        ),
        "deadline_tekst": (
            f"uiterlijk {date.fromisoformat(periode['deadline']).strftime('%d-%m-%Y')}"
            if periode["deadline"]
            else "controleer de uiterste datum in Mijn Belastingdienst Zakelijk"
        ),
        "deadline_betrouwbaar": bool(periode["deadline"]),
        "jaaromzet_tot_nu": round(a["jaaromzet_tot_nu"], 2),
    }


def _bon_datum(d):
    return _iso(d) if d else ""


def bonnen_naar_dict(analyse: dict) -> dict:
    """Zet de bonnen-analyse (OCR-matching) om naar JSON-klare tabeldata voor de app-tab."""
    zeker = [{
        "bon": bon["bestand"],
        "herkend": gebruikt_nummer or "(bedrag + datum uniek)",
        "tx_datum": _iso(t["datum"]),
        "tx_naam": t["naam"],
        "bedrag": round(t.get("bedrag_effect", t["bedrag"]), 2),
        "bewijs": bon.get("bewijs", "onbekend"),
    } for bon, treffers, gebruikt_nummer in analyse["zeker"] for t in treffers]

    onzeker = [{
        "bon": bon["bestand"],
        "bon_bedrag": bon.get("bedrag"),
        "bon_datum": _bon_datum(bon.get("datum")),
        "tx_datum": _iso(t["datum"]),
        "tx_naam": t["naam"],
        "tx_bedrag": round(t.get("bedrag_effect", t["bedrag"]), 2),
        "bewijs": bon.get("bewijs", "onbekend"),
    } for bon, treffers in analyse["onzeker"] for t in treffers]

    zonder_naam = [{
        "bon": bon["bestand"],
        "tx_datum": _iso(t["datum"]),
        "tx_naam": t["naam"],
        "bedrag": round(t.get("bedrag_effect", t["bedrag"]), 2),
    } for bon, treffers in analyse.get("zonder_naam", []) for t in treffers]

    geen_match = [{
        "bon": bon["bestand"],
        "kandidaten": ", ".join(bon["kandidaat_nummers"][:3]),
        "bedrag": bon.get("bedrag"),
        "datum": _bon_datum(bon.get("datum")),
    } for bon in analyse["geen_match"]]

    onleesbaar = [{"bon": bon["bestand"]} for bon in analyse["onleesbaar"]]

    zonder_bon = [{
        "datum": _iso(t["datum"]),
        "naam": t["naam"],
        "bedrag": round(t.get("bedrag_effect", t["bedrag"]), 2),
    } for t in sorted(analyse["uitgaven_zonder_bon"], key=lambda t: t["datum"])]

    return {
        "totaal_bonnen": analyse["totaal_bonnen"],
        "btw_risico": analyse["btw_risico"],
        "btw_risico_naam": analyse.get("btw_risico_naam", 0.0),
        "kenmerken_ingesteld": bool(analyse.get("kenmerken_ingesteld")),
        "tellingen": {
            "zeker": len(analyse["zeker"]),
            "onzeker": len(analyse["onzeker"]),
            "geen_match": len(analyse["geen_match"]),
            "onleesbaar": len(analyse["onleesbaar"]),
            "zonder_bon": len(analyse["uitgaven_zonder_bon"]),
            "zonder_naam": len(analyse.get("zonder_naam", [])),
        },
        "zeker": zeker,
        "onzeker": onzeker,
        "geen_match": geen_match,
        "onleesbaar": onleesbaar,
        "zonder_bon": zonder_bon,
        "zonder_naam": zonder_naam,
    }


def model_naar_dict(model) -> dict:
    """Alles wat de app-tabs nodig hebben, in één JSON-klaar pakket."""
    transacties = [tx_dict(t) for t in sorted(model.transacties, key=lambda t: t["datum"])]
    uitgaven_ruw = [t for t in model.transacties if t["richting"] == "uitgave"]
    inkomsten = [t for t in transacties if t["richting"] == "inkomst"]
    uitgaven = [t for t in transacties if t["richting"] == "uitgave"]
    investeringen = [t for t in uitgaven if t["is_investering"]]
    tot = totalen_per_richting(model.transacties)

    buiten = [{
        "id": t.get("tx_id"),
        "datum": _iso(t["datum"]),
        "naam": t["naam"],
        "factuur": t.get("factuur", ""),
        "tegenrekening": t.get("tegenrekening", ""),
        "mededelingen": t.get("mededelingen", ""),
        "bedrag": round(t["bedrag"], 2),
        "jaarmaand": list(t["buiten_jaarmaand"]),
    } for t in sorted(model.buiten_periode, key=lambda t: t["datum"])]

    buiten_kwartaal = [{
        "id": t.get("tx_id"),
        "datum": _iso(t["datum"]),
        "naam": t["naam"],
        "factuur": t.get("factuur", ""),
        "af_bij": t["af_bij"],
        "bedrag": round(-t["bedrag"] if t["af_bij"] == "Af" else t["bedrag"], 2),
    } for t in sorted(model.buiten_kwartaal, key=lambda t: t["datum"])]

    niet_ingedeeld = [{
        "id": t.get("tx_id"),
        "datum": _iso(t["datum"]),
        "naam": t["naam"],
        "factuur": t.get("factuur", ""),
        "tegenrekening": t.get("tegenrekening", ""),
        "mededelingen": t.get("mededelingen", ""),
        "bedrag": round(t["bedrag"], 2),
    } for t in sorted(model.niet_ingedeeld, key=lambda t: t["datum"])]

    r = model.resultaat()
    periode = _periode_dict(model)
    # Nieuwe leveranciers dit kwartaal: de eerste controlestap (btw is daar een gok).
    nieuw_per = {}
    for t in model.transacties:
        if t.get("is_nieuwe_leverancier"):
            nieuw_per.setdefault(t["leverancier"], []).append(t)
    nieuwe_leveranciers = sorted(([{
        "leverancier": lev,
        "aantal": len(txs),
        "totaal": round(sum(x.get("bedrag_effect", x["bedrag"]) for x in txs), 2),
        "richting": "inkomst" if all(x["richting"] == "inkomst" for x in txs) else "uitgave",
        "btw_pct": sorted({x["btw_pct"] for x in txs}),
        # Bevestigd = bewust gekozen (override of "onthoud voor leverancier") of door een regel bepaald.
        "bevestigd": all(x["btw_bron"] in ("override", "nul-auto", "verlegd") for x in txs)
                     or bool((model.geheugen.get(lev) or {}).get("bevestigd")),
        "eerste_id": txs[0].get("tx_id"),
    } for lev, txs in nieuw_per.items()]), key=lambda n: -abs(n["totaal"]))
    buiten_beschouwing = (
        r["aantal_prive"] + r["aantal_belastingdienst"] + r["aantal_uitgesloten"]
        + r["aantal_buiten_kwartaal"]
    )
    aandachtspunten = (
        len(niet_ingedeeld) + len(buiten) + (0 if periode["volledig_kwartaal"] else 1)
    )
    return {
        "bestandsnaam": model.bestandsnaam,
        "bedrijfsnaam": model.config.get("bedrijfsnaam") or None,
        "periode": periode,
        "importcontrole": {
            "bronregels": len(model.ruwe),
            "verwerkt": r["aantal_verwerkt"],
            "buiten_beschouwing": buiten_beschouwing,
            "aandachtspunten": aandachtspunten,
        },
        "transacties": transacties,
        "niet_ingedeeld": niet_ingedeeld,
        "inkomsten": inkomsten,
        "uitgaven": uitgaven,
        "investeringen": investeringen,
        "investering_drempel": round(model.config["investering_drempel"], 2),
        "leveranciers": _leveranciers_overzicht(uitgaven_ruw),
        "nieuwe_leveranciers": nieuwe_leveranciers,
        "totalen": tot,
        "winst": round(tot["inkomsten_excl"] - tot["uitgaven_excl"], 2),
        "btw_saldo": round(tot["inkomsten_btw"] - tot["uitgaven_btw"], 2),
        "aangifte": _aangifte_dict(model, periode),
        "buiten_periode": buiten,
        "buiten_kwartaal": buiten_kwartaal,
        "tellingen": {
            "verwerkt": r["aantal_verwerkt"],
            "prive": r["aantal_prive"],
            "belastingdienst": r["aantal_belastingdienst"],
            "uitgesloten": r["aantal_uitgesloten"],
            "geen_factuur": r["aantal_geen_factuur"],
            "kredietfactuur": r["aantal_kredietfactuur"],
            "buiten_periode": r["aantal_buiten_periode"],
            "buiten_kwartaal": r["aantal_buiten_kwartaal"],
            "investeringen": r["aantal_investeringen"],
            "verlegd": r["aantal_verlegd"],
            "nieuwe_leveranciers": r["aantal_nieuwe_leveranciers"],
        },
        "geen_periode_gevonden": r["geen_periode_gevonden"],
    }
