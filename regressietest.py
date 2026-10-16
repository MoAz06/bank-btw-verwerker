#!/usr/bin/env python3
"""Regressie-vangnet voor de refactor naar een app.

Legt de door de app berekende cijfers vast als 'gouden referentie' en vergelijkt latere
runs daartegen. Zo weet je bij elke refactorstap zeker dat de getallen identiek blijven.

Gebruik:
    python regressietest.py opslaan <csv> [<csv> ...]   # maak/overschrijf de gouden referentie
    python regressietest.py check   <csv> [<csv> ...]   # vergelijk huidige uitkomst met de referentie
"""
import json
import sys
import tempfile
from pathlib import Path

from ing_btw_verwerker import _bepaal_aangifte_info, programma_map, verwerk_bestand

GOLDEN_MAP = Path(__file__).resolve().parent / "_golden"


def _r(x):
    return round(float(x), 2) if isinstance(x, (int, float)) else x


def snapshot(csv_pad: Path) -> dict:
    """Berekent een deterministische samenvatting van wat de app produceert voor deze csv.

    Verse leveranciers-geheugen (config-gestuurde btw) voor reproduceerbaarheid; schrijft naar
    een tijdelijke output zodat er geen bestanden blijven slingeren.
    """
    geheugen_pad = programma_map() / "leveranciers.json"
    if geheugen_pad.exists():
        geheugen_pad.unlink()

    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.xlsx"
        r = verwerk_bestand(csv_pad, out, overrides={})

    if geheugen_pad.exists():
        geheugen_pad.unlink()

    aangifte = _bepaal_aangifte_info(r["transacties"], _config())

    txs = []
    for t in sorted(r["transacties"], key=lambda t: (t["datum"], t["naam"], t["bedrag"])):
        txs.append({
            "datum": t["datum"].strftime("%Y-%m-%d"),
            "naam": t["naam"],
            "leverancier": t.get("leverancier"),
            "richting": t["richting"],
            "btw_pct": t["btw_pct"],
            "bedrag_effect": _r(t.get("bedrag_effect", t["bedrag"])),
            "is_investering": bool(t.get("is_investering")),
            "is_verlegd": bool(t.get("is_verlegd")),
            "is_kredietfactuur": bool(t.get("is_kredietfactuur")),
        })

    return {
        "aantal_verwerkt": r["aantal_verwerkt"],
        "winst": _r(r["winst"]),
        "btw_saldo": _r(r["btw_saldo"]),
        "btw_te_reserveren": _r(r["btw_te_reserveren"]),
        "aantal_investeringen": r["aantal_investeringen"],
        "aantal_kredietfactuur": r["aantal_kredietfactuur"],
        "aantal_verlegd": r["aantal_verlegd"],
        "aantal_uitgesloten": r.get("aantal_uitgesloten", 0),
        "aantal_prive": r["aantal_prive"],
        "aantal_belastingdienst": r["aantal_belastingdienst"],
        "aantal_geen_factuur": r["aantal_geen_factuur"],
        "aantal_buiten_periode": r["aantal_buiten_periode"],
        "aangifte": {k: _r(v) for k, v in aangifte.items() if isinstance(v, (int, float))},
        "transacties": txs,
    }


def _config():
    from ing_btw_verwerker import laad_config
    return laad_config(programma_map())


def _verschillen(oud, nieuw, pad=""):
    fouten = []
    if isinstance(oud, dict) and isinstance(nieuw, dict):
        for k in set(oud) | set(nieuw):
            fouten += _verschillen(oud.get(k, "<ontbreekt>"), nieuw.get(k, "<ontbreekt>"), f"{pad}.{k}")
    elif isinstance(oud, list) and isinstance(nieuw, list):
        if len(oud) != len(nieuw):
            fouten.append(f"{pad}: lengte {len(oud)} -> {len(nieuw)}")
        for i, (a, b) in enumerate(zip(oud, nieuw)):
            fouten += _verschillen(a, b, f"{pad}[{i}]")
    elif oud != nieuw:
        fouten.append(f"{pad}: {oud!r} -> {nieuw!r}")
    return fouten


def main():
    if len(sys.argv) < 3 or sys.argv[1] not in ("opslaan", "check"):
        print(__doc__)
        sys.exit(1)
    modus, csv_paden = sys.argv[1], [Path(p) for p in sys.argv[2:]]
    GOLDEN_MAP.mkdir(exist_ok=True)

    alles_ok = True
    for csv_pad in csv_paden:
        snap = snapshot(csv_pad)
        golden_pad = GOLDEN_MAP / f"{csv_pad.stem}.json"
        if modus == "opslaan":
            with open(golden_pad, "w", encoding="utf-8") as f:
                json.dump(snap, f, ensure_ascii=False, indent=1)
            print(f"[opgeslagen] {csv_pad.name}: {snap['aantal_verwerkt']} transacties, "
                  f"winst {snap['winst']}, saldo {snap['btw_saldo']}")
        else:
            if not golden_pad.exists():
                print(f"[GEEN REFERENTIE] {csv_pad.name} - draai eerst 'opslaan'")
                alles_ok = False
                continue
            with open(golden_pad, encoding="utf-8") as f:
                oud = json.load(f)
            fouten = _verschillen(oud, snap)
            if fouten:
                alles_ok = False
                print(f"[VERSCHIL] {csv_pad.name}: {len(fouten)} afwijking(en):")
                for fout in fouten[:25]:
                    print("   ", fout)
                if len(fouten) > 25:
                    print(f"    ... en {len(fouten) - 25} meer")
            else:
                print(f"[OK] {csv_pad.name}: identiek aan referentie "
                      f"({snap['aantal_verwerkt']} transacties, winst {snap['winst']}, saldo {snap['btw_saldo']})")

    sys.exit(0 if alles_ok else 1)


if __name__ == "__main__":
    main()
