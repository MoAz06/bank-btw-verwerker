#!/usr/bin/env python3
"""Kwartaal — lokale btw-verwerking vanuit een bankbestand.

Eén venster met tabs (Transacties, Overzicht, BTW-aangifte, Uitgaven, Inkomsten,
Investeringen, Leveranciers, Bonnencontrole). Alles gebeurt in de app; de Excel is
alleen nog een download-knop. Draait lokaal via WebView2, geen internet nodig.
"""
import sys
import traceback
from pathlib import Path

import webview

from boekhouding import bereken_model, programma_map, schrijf_leveranciers_geheugen
from ing_btw_verwerker import bouw_bonnencontrole_sheet, exporteer_csv, exporteer_excel
from webmodel import bonnen_naar_dict, model_naar_dict

WEBUI = Path(__file__).resolve().parent / "webui" / "index.html"


def _eerste(resultaat):
    """create_file_dialog geeft soms een tuple, soms een string, soms None."""
    if not resultaat:
        return None
    if isinstance(resultaat, (list, tuple)):
        return resultaat[0]
    return resultaat


class Api:
    """De brug tussen de web-UI en het Python-model. JS roept deze methoden aan."""

    def __init__(self):
        self.model = None
        self.csv_pad = None
        self.pending_model = None
        self.pending_csv_pad = None
        self.bonnen_map = None
        self.bonnen_analyse = None

    # ---- laden ----
    def kies_csv(self):
        keuze = _eerste(webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=("ING CSV bestand (*.csv)", "Alle bestanden (*.*)"),
        ))
        if not keuze:
            return {"ok": False, "geannuleerd": True}
        return self.bereid_csv_voor(keuze)

    def bereid_csv_voor(self, pad):
        """Bereken een importpreview zonder het huidige kwartaal te vervangen."""
        try:
            nieuw_pad = Path(pad)
            nieuw_model = bereken_model(nieuw_pad)
            self.pending_csv_pad = nieuw_pad
            self.pending_model = nieuw_model
            return {"ok": True, "data": model_naar_dict(nieuw_model)}
        except SystemExit as fout:
            return {"ok": False, "fout": str(fout.code if isinstance(fout.code, str) else fout)}
        except Exception:
            return {"ok": False, "fout": "Onverwachte fout bij inlezen.", "detail": traceback.format_exc()}

    def bevestig_csv(self):
        """Maak de gecontroleerde import actief."""
        if not self.pending_model or not self.pending_csv_pad:
            return {"ok": False, "fout": "Geen gecontroleerd bestand klaar om te openen."}
        self.model = self.pending_model
        self.csv_pad = self.pending_csv_pad
        self.pending_model = None
        self.pending_csv_pad = None
        self.bonnen_map = None
        self.bonnen_analyse = None
        schrijf_leveranciers_geheugen(programma_map(), self.model.geheugen)
        return {"ok": True, "data": model_naar_dict(self.model)}

    def annuleer_csv(self):
        self.pending_model = None
        self.pending_csv_pad = None
        return {"ok": True}

    def laad(self, pad):
        """Direct laden voor een expliciet startbestand vanaf de commandoregel."""
        resultaat = self.bereid_csv_voor(pad)
        if not resultaat.get("ok"):
            return resultaat
        return self.bevestig_csv()

    def huidige_data(self):
        if not self.model:
            return {"ok": False, "fout": "Nog geen bestand geladen."}
        return {"ok": True, "data": model_naar_dict(self.model)}

    # ---- bewerken (Fase 6) ----
    def zet_btw(self, tid, pct, onthoud=False):
        if not self.model:
            return {"ok": False, "fout": "Nog geen bestand geladen."}
        try:
            self.model.zet_btw(tid, int(pct), onthoud_voor_leverancier=bool(onthoud))
            self.bonnen_analyse = None
            return {"ok": True, "data": model_naar_dict(self.model)}
        except Exception:
            return {"ok": False, "fout": "Bewerken mislukt.", "detail": traceback.format_exc()}

    def herstel_btw(self, tid):
        if not self.model:
            return {"ok": False, "fout": "Nog geen bestand geladen."}
        try:
            self.model.wis_override(tid, "btw_pct")
            self.bonnen_analyse = None
            return {"ok": True, "data": model_naar_dict(self.model)}
        except Exception:
            return {"ok": False, "fout": "Herstellen mislukt.", "detail": traceback.format_exc()}

    def zet_richting(self, tid, richting):
        if not self.model:
            return {"ok": False, "fout": "Nog geen bestand geladen."}
        if richting not in {"inkomst", "uitgave"}:
            return {"ok": False, "fout": "Kies inkomsten of uitgaven."}
        try:
            self.model.zet_override(tid, "richting", richting)
            self.bonnen_analyse = None
            return {"ok": True, "data": model_naar_dict(self.model)}
        except Exception:
            return {"ok": False, "fout": "Indelen mislukt.", "detail": traceback.format_exc()}

    # ---- alles bewerkbaar: één generieke ingang ----
    BEWERKBAAR = {"bedrag", "naam", "leverancier", "factuur", "btw_pct", "richting", "is_investering"}

    @staticmethod
    def _parse_bedrag(waarde):
        """'1.234,56', '1234,56', '1234.56' of een getal -> float. ValueError bij onzin."""
        if isinstance(waarde, (int, float)):
            return float(waarde)
        s = str(waarde).replace("€", "").replace(" ", "").replace(" ", "").strip().lstrip("+-")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif s.count(".") == 1 and len(s.split(".")[1]) == 3:
            s = s.replace(".", "")  # "1.234" is een duizendtal, geen decimaal
        return float(s)

    def zet_veld(self, tid, veld, waarde):
        """Zet een handmatige correctie op één veld; het hele model rekent opnieuw door."""
        if not self.model:
            return {"ok": False, "fout": "Nog geen bestand geladen."}
        if veld not in self.BEWERKBAAR:
            return {"ok": False, "fout": "Dit veld is niet te bewerken."}
        try:
            if veld == "bedrag":
                bedrag = self._parse_bedrag(waarde)
                if not bedrag > 0:
                    return {"ok": False, "fout": "Vul een bedrag groter dan 0 in."}
                nieuw = round(bedrag, 2)
            elif veld == "btw_pct":
                nieuw = int(waarde)
                if not 0 <= nieuw <= 100:
                    return {"ok": False, "fout": "Btw moet tussen 0 en 100 procent liggen."}
            elif veld == "richting":
                if waarde not in {"inkomst", "uitgave"}:
                    return {"ok": False, "fout": "Kies inkomsten of uitgaven."}
                nieuw = waarde
            elif veld == "is_investering":
                nieuw = bool(waarde) if not isinstance(waarde, str) else waarde.lower() in {"1", "true", "ja"}
            else:  # naam, leverancier, factuur
                nieuw = str(waarde or "").strip()[:200]
                if not nieuw:
                    return self.herstel_veld(tid, veld)
            self.model.zet_override(tid, veld, nieuw)
            self.bonnen_analyse = None
            return {"ok": True, "data": model_naar_dict(self.model)}
        except ValueError:
            return {"ok": False, "fout": "Dat is geen geldig bedrag. Gebruik bijvoorbeeld 1234,56."}
        except Exception:
            return {"ok": False, "fout": "Bewerken mislukt.", "detail": traceback.format_exc()}

    def herstel_veld(self, tid, veld):
        """Haalt de handmatige correctie van één veld weg; de bankwaarde geldt weer."""
        if not self.model:
            return {"ok": False, "fout": "Nog geen bestand geladen."}
        if veld not in self.BEWERKBAAR:
            return {"ok": False, "fout": "Dit veld is niet te bewerken."}
        try:
            self.model.wis_override(tid, veld)
            self.bonnen_analyse = None
            return {"ok": True, "data": model_naar_dict(self.model)}
        except Exception:
            return {"ok": False, "fout": "Herstellen mislukt.", "detail": traceback.format_exc()}

    # ---- bonnencontrole (Fase 7) ----
    def kies_bonnen_map(self):
        if not self.model:
            return {"ok": False, "fout": "Laad eerst je bankbestand."}
        keuze = _eerste(webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG))
        if not keuze:
            return {"ok": False, "geannuleerd": True}
        return self.controleer_bonnen(keuze)

    def controleer_bonnen(self, map_pad):
        if not self.model:
            return {"ok": False, "fout": "Laad eerst je bankbestand."}
        try:
            from bonnen_matcher import analyseer_bonnenmap
            self.bonnen_map = Path(map_pad)
            analyse = analyseer_bonnenmap(self.bonnen_map, self.model.transacties)
            self.bonnen_analyse = analyse
            return {"ok": True, "data": bonnen_naar_dict(analyse)}
        except Exception:
            return {"ok": False, "fout": "Bonnen controleren mislukt. Staat er wel een map met pdf's of foto's?",
                    "detail": traceback.format_exc()}

    # ---- downloaden ----
    def download_excel(self):
        if not self.model:
            return {"ok": False, "fout": "Nog geen bestand geladen."}
        standaard = self.csv_pad.with_name(f"{self.csv_pad.stem}_btw_overzicht.xlsx").name
        keuze = _eerste(webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG, save_filename=standaard,
            file_types=("Excel bestand (*.xlsx)",),
        ))
        if not keuze:
            return {"ok": False, "geannuleerd": True}
        try:
            out = Path(keuze)
            if out.suffix.lower() != ".xlsx":
                out = out.with_suffix(".xlsx")
            aangifte_export = dict(self.model.aangifte)
            web_aangifte = model_naar_dict(self.model)["aangifte"]
            aangifte_export["deadline_tekst"] = web_aangifte["deadline_tekst"]
            aangifte_export["periode_tekst"] = web_aangifte["periode_tekst"]
            werkboek = exporteer_excel(self.model, out, aangifte_override=aangifte_export)
            if self.bonnen_analyse:
                bouw_bonnencontrole_sheet(werkboek, self.bonnen_analyse)
                werkboek.save(out)
            csv_pad = exporteer_csv(self.model, out)
            return {"ok": True, "pad": str(out), "csv": str(csv_pad) if csv_pad else None}
        except Exception:
            return {"ok": False, "fout": "Opslaan mislukt.", "detail": traceback.format_exc()}


def main():
    api = Api()
    # Een start-CSV kan als argument meegegeven worden (handig bij testen).
    if len(sys.argv) > 1:
        api.laad(sys.argv[1])
    webview.create_window(
        "Kwartaal",
        str(WEBUI),
        js_api=api,
        width=1180,
        height=820,
        min_size=(900, 640),
    )
    webview.start()


if __name__ == "__main__":
    main()
