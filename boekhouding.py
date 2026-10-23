#!/usr/bin/env python3
"""Het datamodel: van ruwe bank-csv naar een in-memory Boekhouding.

Dit is de ENE bron van waarheid. De Excel-export en (straks) de app-tabs lezen allebei
uit ditzelfde model, zodat ze per definitie niet uiteen kunnen lopen. Geen openpyxl hier.
"""
import csv
import hashlib
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

from berekening import VERLEGGING_SOORTEN, bepaal_aangifte_info, bepaal_btw, bepaal_verlegging, excl_van
from leveranciers import (
    laad_leveranciers_geheugen,
    normaliseer_leverancier,
    onthoud_btw,
    schrijf_leveranciers_geheugen,
    vul_geheugen_aan,
)

DEFAULT_CONFIG = {
    "bedrijfsnaam": "",
    "factuur_kenmerken": [],
    "eigen_rekeningen": {"NL76INGB0003298592", "NL29RABO0366729276"},
    "eigen_naam_patronen": ["zakelijke oranje spaarrekening"],
    "nul_btw_patronen": [
        "kosten zakelijk betalingsverkeer", "ontvangsten hiltermann l",
        "flexkracht", "unive het groene hart bemiddeling", "vwpfs",
    ],
    "verlegd_patronen": [],
    "eu_verlegd_patronen": [],
    "buiten_eu_verlegd_patronen": [],
    "voorbelasting_vereist_bon": True,
    "uitsluit_patronen": ["lening", "aflossing"],
    "investering_drempel": 450.0,
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

REQUIRED_COLUMNS = [
    "Datum", "Naam / Omschrijving", "Rekening", "Tegenrekening",
    "Code", "Af Bij", "Bedrag (EUR)", "Mutatiesoort", "Mededelingen",
]

FACTUUR_PATROON = re.compile(r"\b(\d{4})(0[1-9]|1[0-2])\d{2,6}\b")
CIJFER_PATROON = re.compile(r"\d{2,}")
_STOP_LABELS = r"(?:Naam:|Omschrijving:|IBAN:|Kenmerk:|Machtiging ID:|Incassant ID:|Datum/Tijd:|Valutadatum:)"
OMSCHRIJVING_PATROON = re.compile(r"Omschrijving:\s*(.*?)\s*(?=" + _STOP_LABELS + r"|$)")
BESTANDSNAAM_DATUM_PATROON = re.compile(r"(\d{2})-(\d{2})-(\d{4})_(\d{2})-(\d{2})-(\d{4})")


# ---------- configuratie ----------
def programma_map() -> Path:
    """Map waar config.json / leveranciers.json horen te staan - naast de .exe of dit script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def overrides_pad(map_pad: Path = None) -> Path:
    """Bestand met handmatige correcties per transactie (btw%, richting, investering)."""
    return (map_pad or programma_map()) / "overrides.json"


def laad_overrides(map_pad: Path = None) -> dict:
    """Retourneert {tx_id: {veld: waarde}}. Leeg dict als er nog niets gecorrigeerd is."""
    pad = overrides_pad(map_pad)
    if not pad.exists():
        return {}
    try:
        with open(pad, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def schrijf_overrides(overrides: dict, map_pad: Path = None):
    """Slaat de correcties op. Verwijdert het bestand als er geen correcties (meer) zijn."""
    pad = overrides_pad(map_pad)
    try:
        if not overrides:
            if pad.exists():
                pad.unlink()
            return
        with open(pad, "w", encoding="utf-8") as f:
            json.dump(overrides, f, ensure_ascii=False, indent=2, sort_keys=True)
    except OSError:
        pass


def laad_config(script_dir: Path = None) -> dict:
    script_dir = script_dir or programma_map()
    pad = script_dir / "config.json"
    if not pad.exists():
        return DEFAULT_CONFIG
    with open(pad, encoding="utf-8") as f:
        data = json.load(f)
    return {
        "bedrijfsnaam": str(data.get("bedrijfsnaam", "") or "").strip(),
        # Teksten waaraan je herkent dat een factuur op jouw naam staat (bedrijfsnaam, btw-id, kvk, adres).
        "factuur_kenmerken": [str(p).strip() for p in data.get("factuur_kenmerken", []) if str(p).strip()],
        "eigen_rekeningen": {iban.replace(" ", "").upper() for iban in data.get("eigen_rekeningen", [])},
        "eigen_naam_patronen": [p.lower() for p in data.get("eigen_naam_patronen", [])],
        "nul_btw_patronen": [p.lower() for p in data.get("nul_btw_patronen", DEFAULT_CONFIG["nul_btw_patronen"])],
        "verlegd_patronen": [p.lower() for p in data.get("verlegd_patronen", [])],
        # Naar jou verlegde btw uit het buitenland: EU-inkoop (rubriek 4b) en buiten de EU (4a).
        "eu_verlegd_patronen": [p.lower() for p in data.get("eu_verlegd_patronen", [])],
        "buiten_eu_verlegd_patronen": [p.lower() for p in data.get("buiten_eu_verlegd_patronen", [])],
        # Art. 15 Wet OB: aftrek alleen met een factuur. Uit = oude gedrag (elke uitgave telt mee).
        "voorbelasting_vereist_bon": bool(data.get("voorbelasting_vereist_bon", True)),
        "uitsluit_patronen": [p.lower() for p in data.get("uitsluit_patronen", DEFAULT_CONFIG["uitsluit_patronen"])],
        "investering_drempel": float(data.get("investering_drempel", DEFAULT_CONFIG["investering_drempel"])),
        "geen_investering_leveranciers": data.get("geen_investering_leveranciers", DEFAULT_CONFIG["geen_investering_leveranciers"]),
        "leverancier_aliassen": {k.lower(): v for k, v in data.get("leverancier_aliassen", {}).items()},
        "kor_actief": bool(data.get("kor_actief", False)),
        "jaaromzet_reeds": float(data.get("jaaromzet_reeds", 0.0)),
        "boekhoud_export": bool(data.get("boekhoud_export", True)),
    }


# ---------- csv inlezen ----------
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


def tx_id(t: dict) -> str:
    """Stabiel id per transactie (boekdatum+bedrag+tegenrekening+mededeling), voor overrides.

    Altijd op de bankgegevens: een handmatige factuurdatum vervangt 'datum', maar het id
    blijft op de boekdatum, anders raakt de correctie zichzelf kwijt.
    """
    boekdatum = t.get("boekdatum") or t["datum"]
    sleutel = f"{boekdatum.strftime('%Y%m%d')}|{t['bedrag']:.2f}|{t['tegenrekening']}|{t['mededelingen']}"
    return hashlib.sha1(sleutel.encode("utf-8")).hexdigest()[:12]


# ---------- classificatie ----------
def bepaal_geldige_jaarmaanden(bestandsnaam: str):
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


def bepaal_kwartaal_venster(bestandsnaam: str):
    """(start, eind) van het kwartaal als de bestandsnaam een volledig kwartaal beschrijft, anders None.

    Regel (dezelfde als het periodelabel in de app): de export begint exact op dag 1 van een
    kwartaal en eindigt op of maximaal 7 dagen na het kwartaaleinde (ING-exports lopen in de
    praktijk soms een paar dagen door). Boekingen na het kwartaaleinde horen dan in het
    volgende kwartaal, anders tel je ze twee keer.
    """
    match = BESTANDSNAAM_DATUM_PATROON.search(bestandsnaam or "")
    if not match:
        return None
    try:
        d1, m1, j1, d2, m2, j2 = map(int, match.groups())
        start, eind = datetime(j1, m1, d1), datetime(j2, m2, d2)
    except ValueError:
        return None
    if eind < start:
        start, eind = eind, start
    kwartaal = (start.month - 1) // 3 + 1
    kwartaal_start = datetime(start.year, 1 + (kwartaal - 1) * 3, 1)
    if kwartaal == 4:
        volgende_start = datetime(start.year + 1, 1, 1)
    else:
        volgende_start = datetime(start.year, kwartaal_start.month + 3, 1)
    kwartaal_eind = volgende_start - timedelta(days=1)
    if start == kwartaal_start and kwartaal_eind <= eind <= kwartaal_eind + timedelta(days=7):
        return kwartaal_start, kwartaal_eind
    return None


def is_prive(transactie: dict, config: dict) -> bool:
    tegenrekening = transactie["tegenrekening"].replace(" ", "").upper()
    if tegenrekening and tegenrekening in config["eigen_rekeningen"]:
        return True
    naam_lower = transactie["naam"].lower()
    return any(patroon in naam_lower for patroon in config["eigen_naam_patronen"])


def is_belastingdienst(transactie: dict) -> bool:
    return BELASTINGDIENST_PATROON in transactie["naam"].lower()


def is_uitgesloten(transactie: dict, config: dict) -> bool:
    """Leningen/aflossingen e.d. zijn geen kosten of omzet. Hele-woord-match."""
    tekst = (transactie["naam"] + " " + transactie["mededelingen"]).lower()
    for patroon in config.get("uitsluit_patronen", []):
        if re.search(r"\b" + re.escape(patroon) + r"\b", tekst):
            return True
    return False


def bepaal_omschrijving_tekst(mededelingen: str) -> str:
    match = OMSCHRIJVING_PATROON.search(mededelingen)
    return match.group(1) if match else ""


def classificeer_richting(transactie: dict, omschrijving: str, geldige_jaarmaanden):
    """(richting, is_kredietfactuur, buiten_periode_jaarmaand). richting: uitgave|inkomst|buiten_periode|None."""
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


# Velden die per transactie handmatig te overschrijven zijn (overrides.json).
OVERRIDE_VELDEN = ("btw_pct", "richting", "is_investering", "uitgesloten", "bedrag", "naam", "leverancier", "factuur",
                   "factuurdatum", "verlegging", "bon_aanwezig")


def parse_factuurdatum(waarde):
    """'2026-03-28' of '28-03-2026' -> datetime, anders None."""
    if isinstance(waarde, datetime):
        return waarde
    tekst = str(waarde or "").strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(tekst, fmt)
        except ValueError:
            continue
    return None
BTW_VELDEN = OVERRIDE_VELDEN  # compat-alias


class Boekhouding:
    """In-memory model: transacties + config + geheugen + handmatige overrides.

    Roep bereken() aan na een wijziging; dat draait alle pure functies opnieuw over de
    volledige lijst (ruim snel genoeg) zodat elke afgeleide waarde altijd klopt.
    """

    def __init__(self, ruwe_transacties, bestandsnaam, config, geheugen, overrides=None, map_pad=None):
        self.ruwe = ruwe_transacties
        self.bestandsnaam = bestandsnaam
        self.config = config
        self.geheugen = geheugen
        self.overrides = overrides or {}
        self.map_pad = map_pad or programma_map()
        self.geldige_jaarmaanden = bepaal_geldige_jaarmaanden(bestandsnaam or "")
        self.kwartaal_venster = bepaal_kwartaal_venster(bestandsnaam or "")
        self.nieuwe_leveranciers = 0
        # Leveranciers die vóór dit bestand al bekend waren. Alles wat daar niet in zit is
        # "nieuw dit kwartaal" en verdient een btw-controle (standaard 21% is een gok).
        self.bekende_leveranciers = set(geheugen.keys())
        # tx_id's waaraan een bon of factuur is gekoppeld (uit de bonnencontrole); bepaalt samen
        # met de override 'bon_aanwezig' of de voorbelasting in 5b meetelt.
        self.bonnen_ids = set()
        self.bereken()

    def zet_bonnen(self, tx_ids):
        self.bonnen_ids = set(tx_ids or [])
        self.bereken()

    def _binnen_kwartaal(self, t) -> bool:
        if self.kwartaal_venster is None:
            return True
        start, eind = self.kwartaal_venster
        return start <= t["datum"] <= eind

    def bereken(self):
        config, geheugen = self.config, self.geheugen
        aliassen = config["leverancier_aliassen"]

        # Factuurstelsel (art. 13 en 15 Wet OB): het tijdvak volgt de factuurdatum. De boekdatum
        # van de bank is alleen de standaardwaarde; een handmatige factuurdatum vervangt 'm
        # voor alles wat volgt (kwartaalvenster, sortering, bonnenmatching, Excel).
        ruwe_met_datum = []
        for t in self.ruwe:
            ovr = self.overrides.get(tx_id(t), {})
            factuurdatum = parse_factuurdatum(ovr.get("factuurdatum")) if ovr.get("factuurdatum") else None
            if factuurdatum:
                ruwe_met_datum.append({**t, "boekdatum": t["datum"], "datum": factuurdatum, "factuurdatum": factuurdatum})
            else:
                ruwe_met_datum.append({**t, "boekdatum": t["datum"], "factuurdatum": None})

        # Boekingen buiten het herkende kwartaal (meestal de eerste dagen van het volgende
        # kwartaal die ING nog meeneemt) tellen nergens mee; ze horen in de volgende aangifte.
        # Staat de factuur wél in dit kwartaal gedateerd, zet dan de factuurdatum: dan telt de regel mee.
        self.buiten_kwartaal = [
            {**t, "tx_id": tx_id(t), "factuur": bepaal_omschrijving_tekst(t["mededelingen"]),
             "bewerkt": ({"factuurdatum": t["boekdatum"]} if t.get("factuurdatum") else {})}
            for t in ruwe_met_datum if not self._binnen_kwartaal(t)
        ]
        binnen = [t for t in ruwe_met_datum if self._binnen_kwartaal(t)]
        vereist_bon = config.get("voorbelasting_vereist_bon", True)

        self.aantal_prive = sum(1 for t in binnen if is_prive(t, config))
        self.aantal_belastingdienst = sum(
            1 for t in binnen if not is_prive(t, config) and is_belastingdienst(t)
        )
        self.aantal_uitgesloten = sum(
            1 for t in binnen
            if not is_prive(t, config) and not is_belastingdienst(t)
            and is_uitgesloten(t, config) and not self._override_richting(t)
        )
        overig = [
            t for t in binnen
            if not is_prive(t, config) and not is_belastingdienst(t)
            and not (is_uitgesloten(t, config) and not self._override_richting(t))
        ]

        verwerkt, buiten_periode, niet_ingedeeld = [], [], []
        self.aantal_geen_factuur = 0
        self.aantal_kredietfactuur = 0
        self.nieuwe_leveranciers = 0
        for t in overig:
            tid = tx_id(t)  # altijd op de ruwe bankregel, zodat correcties stabiel blijven
            ovr = self.overrides.get(tid, {})
            bewerkt = {}  # veld -> oorspronkelijke (automatische) waarde, voor herstellen in de UI
            if t.get("factuurdatum"):
                bewerkt["factuurdatum"] = t["boekdatum"]

            # Handmatig bedrag of naam: vervangt de bankwaarde voor alles wat volgt.
            if isinstance(ovr.get("bedrag"), (int, float)) and float(ovr["bedrag"]) > 0:
                bewerkt["bedrag"] = t["bedrag"]
                t = {**t, "bedrag": round(float(ovr["bedrag"]), 2)}
            if isinstance(ovr.get("naam"), str) and ovr["naam"].strip():
                bewerkt["naam"] = t["naam"]
                t = {**t, "naam": ovr["naam"].strip()}
            omschrijving = bepaal_omschrijving_tekst(t["mededelingen"])
            if isinstance(ovr.get("factuur"), str) and ovr["factuur"].strip():
                bewerkt["factuur"] = omschrijving
                omschrijving = ovr["factuur"].strip()
            richting, is_krediet, buiten_jaarmaand = classificeer_richting(t, omschrijving, self.geldige_jaarmaanden)

            if ovr.get("richting"):  # handmatige richting-correctie wint
                bewerkt["richting"] = richting
                richting, buiten_jaarmaand = ovr["richting"], None
                # Het teken volgt de geldstroom: geld dat BINNENKWAM maar als uitgave wordt
                # geboekt is een creditnota (negatieve kosten); geld dat WEGGING maar als
                # inkomst wordt geboekt is een terugbetaling aan een klant (negatieve omzet).
                is_krediet = (richting == "uitgave" and t["af_bij"] == "Bij") or (
                    richting == "inkomst" and t["af_bij"] == "Af"
                )
            if richting is None:
                self.aantal_geen_factuur += 1
                niet_ingedeeld.append({
                    **t,
                    "tx_id": tid,
                    "factuur": omschrijving,
                })
                continue
            if richting == "buiten_periode":
                buiten_periode.append({**t, "tx_id": tid, "factuur": omschrijving, "buiten_jaarmaand": buiten_jaarmaand})
                continue
            if is_krediet:
                self.aantal_kredietfactuur += 1

            leverancier = normaliseer_leverancier(t["naam"], aliassen)
            if isinstance(ovr.get("leverancier"), str) and ovr["leverancier"].strip():
                bewerkt["leverancier"] = leverancier
                leverancier = ovr["leverancier"].strip()
            btw_pct, btw_bron, is_verlegd, is_auto_nul = bepaal_btw(t, leverancier, config, geheugen)
            verlegging = bepaal_verlegging(t, config) if is_verlegd else None
            if isinstance(ovr.get("btw_pct"), (int, float)):  # handmatige btw-correctie wint van alles
                bewerkt["btw_pct"] = btw_pct
                btw_pct, btw_bron, is_auto_nul = int(ovr["btw_pct"]), "override", int(ovr["btw_pct"]) == 0
                is_verlegd, verlegging = False, None
            if ovr.get("verlegging") in VERLEGGING_SOORTEN:  # handmatige verlegging: 0% en verlegd
                bewerkt["verlegging"] = verlegging
                verlegging, is_verlegd, btw_pct, btw_bron, is_auto_nul = ovr["verlegging"], True, 0, "verlegd", False
            excl = excl_van(t["bedrag"], btw_pct)
            is_investering = (
                richting == "uitgave" and not is_krediet
                and excl >= config["investering_drempel"]
                and leverancier not in config.get("geen_investering_leveranciers", [])
            )
            if "is_investering" in ovr:
                bewerkt["is_investering"] = is_investering
                is_investering = bool(ovr["is_investering"])

            # Bewijs voor de aftrek (art. 15 lid 1 onder a Wet OB): gekoppelde bon uit de
            # bonnencontrole, of handmatig bevestigd "factuur aanwezig". Zonder bewijs telt de
            # btw niet mee in 5b (tenzij de config dat uitzet).
            if tid in self.bonnen_ids:
                bon_status = "gekoppeld"
            elif ovr.get("bon_aanwezig") is True:
                bon_status = "bevestigd"
                bewerkt["bon_aanwezig"] = False
            else:
                bon_status = "ontbreekt"
            aftrekbaar = (
                richting != "uitgave" or is_krediet or btw_pct == 0 or not vereist_bon
                or bon_status != "ontbreekt"
            )

            verwerkt.append({
                **t,
                "tx_id": tid,
                "bewerkt": bewerkt,
                "verlegging": verlegging,
                "bon_status": bon_status,
                "aftrekbaar": aftrekbaar,
                "is_nieuwe_leverancier": leverancier not in self.bekende_leveranciers,
                "richting": richting,
                "is_kredietfactuur": is_krediet,
                "factuur": omschrijving,
                "leverancier": leverancier,
                "btw_pct": btw_pct,
                "btw_bron": btw_bron,
                "btw_auto_nul": is_auto_nul,
                "is_verlegd": is_verlegd,
                "is_investering": is_investering,
                "bedrag_effect": -t["bedrag"] if is_krediet else t["bedrag"],
            })
            if vul_geheugen_aan(geheugen, leverancier, btw_pct, richting == "uitgave"):
                self.nieuwe_leveranciers += 1

        self.transacties = verwerkt
        self.buiten_periode = buiten_periode
        self.niet_ingedeeld = niet_ingedeeld
        self.aangifte = bepaal_aangifte_info(verwerkt, config)

    def _override_richting(self, t):
        return self.overrides.get(tx_id(t), {}).get("richting")

    # ---- bewerken (Fase 4) ----
    def zet_override(self, tid: str, veld: str, waarde, bewaar: bool = True):
        if veld not in OVERRIDE_VELDEN:
            raise ValueError(f"onbekend override-veld: {veld}")
        self.overrides.setdefault(tid, {})[veld] = waarde
        self.bereken()
        if bewaar:
            self.bewaar_overrides()

    def wis_override(self, tid: str, veld: str = None, bewaar: bool = True):
        """Wist één veld of (veld=None) de hele correctie voor deze transactie."""
        if veld is None:
            self.overrides.pop(tid, None)
        elif tid in self.overrides:
            self.overrides[tid].pop(veld, None)
            if not self.overrides[tid]:
                self.overrides.pop(tid, None)
        self.bereken()
        if bewaar:
            self.bewaar_overrides()

    def zet_btw(self, tid: str, btw_pct: int, onthoud_voor_leverancier: bool = False, bewaar: bool = True):
        """Handmatige btw-correctie voor één transactie; optioneel onthouden voor de leverancier."""
        if onthoud_voor_leverancier:
            tx = next((t for t in self.transacties if t.get("tx_id") == tid), None)
            if tx:
                onthoud_btw(self.geheugen, tx["leverancier"], btw_pct)
                schrijf_leveranciers_geheugen(self.map_pad, self.geheugen)
        self.zet_override(tid, "btw_pct", int(btw_pct), bewaar=bewaar)

    def bewaar_overrides(self):
        schrijf_overrides(self.overrides, self.map_pad)

    # ---- samenvatting voor GUI / compat ----
    def resultaat(self) -> dict:
        a = self.aangifte
        return {
            "aantal_verwerkt": len(self.transacties),
            "aantal_prive": self.aantal_prive,
            "aantal_belastingdienst": self.aantal_belastingdienst,
            "aantal_uitgesloten": self.aantal_uitgesloten,
            "aantal_geen_factuur": self.aantal_geen_factuur,
            "aantal_kredietfactuur": self.aantal_kredietfactuur,
            "aantal_buiten_periode": len(self.buiten_periode),
            "aantal_buiten_kwartaal": len(self.buiten_kwartaal),
            "aantal_niet_onderbouwd": sum(
                1 for t in self.transacties if t["richting"] == "uitgave" and not t.get("aftrekbaar", True)),
            "aantal_investeringen": sum(1 for t in self.transacties if t.get("is_investering")),
            "aantal_verlegd": sum(1 for t in self.transacties if t.get("is_verlegd")),
            "nieuwe_leveranciers": self.nieuwe_leveranciers,
            "aantal_nieuwe_leveranciers": len({t["leverancier"] for t in self.transacties if t.get("is_nieuwe_leverancier")}),
            "geen_periode_gevonden": self.geldige_jaarmaanden is None,
            "kor_actief": self.config["kor_actief"],
            "winst": a["winst"],
            "btw_saldo": a["btw_saldo"],
            "btw_te_reserveren": a["btw_te_reserveren"],
            "deadline_tekst": a["deadline_tekst"],
            "transacties": self.transacties,
            "buiten_periode": self.buiten_periode,
            "aangifte": a,
        }


def bereken_model(csv_pad: Path, config: dict = None, geheugen: dict = None, overrides: dict = None) -> Boekhouding:
    """Leest de csv en bouwt het model. Laadt config/geheugen/overrides zelf als je ze niet meegeeft."""
    map_pad = programma_map()
    config = config if config is not None else laad_config(map_pad)
    geheugen = geheugen if geheugen is not None else laad_leveranciers_geheugen(map_pad)
    overrides = overrides if overrides is not None else laad_overrides(map_pad)
    ruwe = lees_csv(csv_pad)
    if not ruwe:
        raise SystemExit("Geen (geldige) transacties gevonden in dit bestand.")
    return Boekhouding(ruwe, csv_pad.name, config, geheugen, overrides, map_pad=map_pad)
