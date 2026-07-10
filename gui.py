#!/usr/bin/env python3
"""Venster om een bank-csv te verwerken, met een stoplicht-eindscherm: in één oogopslag
zie je of je klaar bent of nog iets moet controleren."""
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ing_btw_verwerker import _formateer_euro, bouw_bonnencontrole_sheet, verwerk_bestand

VENSTER_TITEL = "Bank Btw Verwerker"

GROEN = "#2E7D32"
ORANJE = "#E67E22"
ROOD = "#C0392B"
GRIJS = "#7F8C8D"
BLAUW = "#1F4E78"


class App(tk.Tk):
    def __init__(self, startbestand: str = None):
        super().__init__()
        self.title(VENSTER_TITEL)
        self.geometry("620x640")
        self.configure(bg="white")
        self.laatste_output_pad = None
        self.laatste_export_pad = None
        self.laatste_transacties = None
        self.laatste_wb = None
        self.laatste_resultaat = None
        self.bonnen_frame = None
        self._wachtrij = queue.Queue()

        tk.Label(
            self, text="Bank-csv → btw / inkomsten-uitgaven, aangifte-klaar.",
            font=("Segoe UI", 12, "bold"), bg="white", fg=BLAUW,
        ).pack(padx=16, pady=(14, 4), anchor="w")

        knoprij = tk.Frame(self, bg="white")
        knoprij.pack(padx=16, pady=4, anchor="w", fill="x")
        self.csv_knop = tk.Button(
            knoprij, text="Kies bank-csv…", command=self.kies_bestand,
            font=("Segoe UI", 10), padx=10, pady=6,
        )
        self.csv_knop.pack(side="left")
        self.bonnen_knop = tk.Button(
            knoprij, text="Vergelijk met bonnen (map)…", command=self.kies_bonnenmap,
            font=("Segoe UI", 10), padx=10, pady=6, state="disabled",
        )
        self.bonnen_knop.pack(side="left", padx=8)

        # Resultaatpaneel (stoplicht) — wordt telkens opnieuw gevuld.
        self.paneel = tk.Frame(self, bg="white")
        self.paneel.pack(padx=16, pady=8, anchor="w", fill="both", expand=True)

        self.voortgang_label = tk.Label(self, text="", font=("Segoe UI", 9), bg="white", fg=GRIJS)
        self.voortgang_balk = ttk.Progressbar(self, length=580, mode="determinate")

        openrij = tk.Frame(self, bg="white")
        openrij.pack(padx=16, pady=(0, 14), anchor="w")
        self.open_knop = tk.Button(
            openrij, text="Open Excel", command=self.open_excel,
            font=("Segoe UI", 10), padx=10, pady=6, state="disabled",
        )
        self.open_knop.pack(side="left")
        self.export_knop = tk.Button(
            openrij, text="Open boekhouder-CSV", command=self.open_export,
            font=("Segoe UI", 10), padx=10, pady=6, state="disabled",
        )
        self.export_knop.pack(side="left", padx=8)

        self._toon_placeholder()
        if startbestand:
            self.verwerk(Path(startbestand))

    # ---- paneel-helpers ----
    def _leeg_paneel(self):
        for w in self.paneel.winfo_children():
            w.destroy()

    def _toon_placeholder(self):
        self._leeg_paneel()
        tk.Label(
            self.paneel, text="Kies een bank-csv om te beginnen.",
            font=("Segoe UI", 10), bg="white", fg=GRIJS,
        ).pack(anchor="w")

    def _kop(self, tekst, ouder=None):
        tk.Label(ouder or self.paneel, text=tekst, font=("Segoe UI", 11, "bold"), bg="white", fg=BLAUW).pack(anchor="w", pady=(8, 2))

    def _geldregel(self, label, bedrag, kleur=None, dik=False, ouder=None):
        rij = tk.Frame(ouder or self.paneel, bg="white")
        rij.pack(anchor="w", fill="x")
        tk.Label(rij, text=label, font=("Segoe UI", 10), bg="white", width=34, anchor="w").pack(side="left")
        tk.Label(
            rij, text=_formateer_euro(bedrag), bg="white",
            font=("Segoe UI", 11 if dik else 10, "bold" if dik else "normal"),
            fg=kleur or "black",
        ).pack(side="left")

    def _stoplicht(self, kleur, tekst, ouder=None):
        rij = tk.Frame(ouder or self.paneel, bg="white")
        rij.pack(anchor="w", fill="x", pady=1)
        tk.Canvas(rij, width=14, height=14, bg="white", highlightthickness=0).pack(side="left", padx=(0, 8))
        stip = rij.winfo_children()[0]
        stip.create_oval(2, 2, 12, 12, fill=kleur, outline=kleur)
        tk.Label(rij, text=tekst, font=("Segoe UI", 10), bg="white", anchor="w", justify="left", wraplength=520).pack(side="left")

    # ---- acties ----
    def kies_bestand(self):
        pad = filedialog.askopenfilename(
            title="Kies bank-csv", filetypes=[("CSV bestanden", "*.csv"), ("Alle bestanden", "*.*")],
        )
        if pad:
            self.verwerk(Path(pad))

    def _fout_melding(self, fout) -> str:
        tekst = str(fout)
        low = tekst.lower()
        if "permission denied" in low or "being used" in low:
            return ("Het Excel-bestand staat waarschijnlijk nog open. Sluit het in Excel "
                    "en probeer opnieuw.")
        if "geen ing csv" in low or "kolommen" in low:
            return ("Dit lijkt geen bank-csv-export. Exporteer in je bank je transacties als CSV "
                    "(bij ING: Zoeken → Downloaden → CSV) en kies dat bestand.")
        return tekst

    def verwerk(self, csv_pad: Path):
        self.open_knop.configure(state="disabled")
        self.export_knop.configure(state="disabled")
        self.bonnen_knop.configure(state="disabled")
        self.laatste_transacties = None
        self._leeg_paneel()
        tk.Label(self.paneel, text="Bezig met verwerken…", font=("Segoe UI", 10), bg="white", fg=GRIJS).pack(anchor="w")
        self.update_idletasks()
        try:
            resultaat = verwerk_bestand(csv_pad)
        except SystemExit as fout:
            self._toon_fout(self._fout_melding(fout))
            return
        except Exception as fout:
            self._toon_fout(self._fout_melding(fout))
            return

        self.laatste_output_pad = resultaat["output_pad"]
        self.laatste_export_pad = resultaat.get("export_pad")
        self.laatste_transacties = resultaat["transacties"]
        self.laatste_wb = resultaat["workbook"]
        self.laatste_resultaat = resultaat
        self._toon_resultaat(resultaat)
        self.open_knop.configure(state="normal")
        self.bonnen_knop.configure(state="normal")
        self.export_knop.configure(state="normal" if self.laatste_export_pad else "disabled")

    def _toon_fout(self, tekst):
        self._leeg_paneel()
        tk.Label(self.paneel, text="Mislukt", font=("Segoe UI", 11, "bold"), bg="white", fg=ROOD).pack(anchor="w")
        tk.Label(self.paneel, text=tekst, font=("Segoe UI", 10), bg="white", fg="black",
                 wraplength=560, justify="left").pack(anchor="w", pady=4)

    def _toon_resultaat(self, r):
        self._leeg_paneel()
        if r["aantal_verwerkt"] == 0:
            tk.Label(self.paneel, text="Geen bruikbare transacties gevonden",
                     font=("Segoe UI", 12, "bold"), bg="white", fg=ORANJE).pack(anchor="w")
            tk.Label(self.paneel, text="Alles is uitgesloten (privé/belasting/lening) of er zaten geen "
                     "herkenbare transacties in dit bestand. Controleer je config.json of de csv.",
                     font=("Segoe UI", 10), bg="white", fg="black", wraplength=560, justify="left").pack(anchor="w", pady=4)
            self.bonnen_frame = tk.Frame(self.paneel, bg="white")
            self.bonnen_frame.pack(anchor="w", fill="x")
            return
        tk.Label(self.paneel, text=f"Klaar — {r['aantal_verwerkt']} transacties verwerkt",
                 font=("Segoe UI", 12, "bold"), bg="white", fg=GROEN).pack(anchor="w")

        self._kop("Resultaat")
        self._geldregel("Winst (excl. btw):", r["winst"], GROEN if r["winst"] >= 0 else ROOD, dik=True)
        saldo = r["btw_saldo"]
        if saldo > 0:
            self._geldregel("Btw te betalen:", saldo, ROOD, dik=True)
            self._stoplicht(GROEN, f"Zet {_formateer_euro(r['btw_te_reserveren'])} opzij — {r['deadline_tekst']}.")
        else:
            self._geldregel("Btw terug te vragen:", -saldo, GROEN, dik=True)

        self._kop("Te controleren")
        if r["aantal_geen_factuur"]:
            self._stoplicht(ORANJE, f"{r['aantal_geen_factuur']} inkomsten zonder herkenbare factuur (niet meegeteld — check Overzicht).")
        if r["aantal_buiten_periode"]:
            self._stoplicht(ORANJE, f"{r['aantal_buiten_periode']} factuur/facturen buiten deze periode — zie Overzicht.")
        if r.get("aantal_investeringen"):
            self._stoplicht(GRIJS, f"{r['aantal_investeringen']} grote uitgaven om te checken (investering of kosten?) — zie tab Investeringen.")
        if r.get("aantal_verlegd"):
            self._stoplicht(ORANJE, f"{r['aantal_verlegd']} transacties met verlegde btw — controleer met je boekhouder.")
        if r.get("aantal_uitgesloten"):
            self._stoplicht(GRIJS, f"{r['aantal_uitgesloten']} lening/aflossing uitgesloten (geen kosten/omzet).")
        if r.get("nieuwe_leveranciers"):
            self._stoplicht(GRIJS, f"{r['nieuwe_leveranciers']} nieuwe leveranciers onthouden (leveranciers.json) — check of de btw% klopt.")
        aandacht = any([
            r["aantal_geen_factuur"], r["aantal_buiten_periode"], r.get("aantal_verlegd"),
            r.get("aantal_investeringen"), r.get("aantal_uitgesloten"), r.get("nieuwe_leveranciers"),
        ])
        if not aandacht:
            self._stoplicht(GROEN, "Niets bijzonders om te controleren aan de transactiekant.")

        self._stoplicht(GRIJS, "Vergelijk met je bonnen (knop bovenaan) om te zien of elke uitgave een bon heeft.")

        # Eigen sub-frame voor de bonnenuitkomst (wordt bij elke bonnenrun vervangen, niet gestapeld).
        self.bonnen_frame = tk.Frame(self.paneel, bg="white")
        self.bonnen_frame.pack(anchor="w", fill="x")

    def _toon_bonnen(self, analyse):
        f = self.bonnen_frame
        for w in f.winfo_children():
            w.destroy()
        self._kop("Bonnencontrole", ouder=f)
        if analyse.get("totaal_bonnen", 0) == 0:
            self._stoplicht(GRIJS, "Geen bonnen (PDF's of foto's) in die map gevonden.", ouder=f)
            return
        if len(analyse["uitgaven_zonder_bon"]):
            self._stoplicht(ROOD, f"{len(analyse['uitgaven_zonder_bon'])} uitgaven zonder bon.", ouder=f)
        if analyse.get("btw_risico", 0) > 0:
            self._stoplicht(ROOD, f"Btw-risico {_formateer_euro(analyse['btw_risico'])}: voorbelasting zonder factuur eronder (kan bij controle geschrapt worden).", ouder=f)
        if len(analyse["onzeker"]):
            self._stoplicht(ORANJE, f"{len(analyse['onzeker'])} bonnen met onzekere match — controleer in tab Bonnencontrole.", ouder=f)
        if len(analyse["geen_match"]):
            self._stoplicht(ORANJE, f"{len(analyse['geen_match'])} bonnen zonder match in de bank.", ouder=f)
        if len(analyse["onleesbaar"]):
            hint = ""
            try:
                from bonnen_matcher import vind_tesseract
                if not vind_tesseract():
                    hint = " Installeer Tesseract-OCR voor gescande/foto-bonnen."
            except Exception:
                pass
            self._stoplicht(GRIJS, f"{len(analyse['onleesbaar'])} bonnen onleesbaar (scan/foto).{hint}", ouder=f)
        self._stoplicht(GROEN, f"{len(analyse['zeker'])} bonnen zeker gematcht.", ouder=f)
        if not analyse["uitgaven_zonder_bon"] and not analyse["onzeker"] and not analyse["geen_match"]:
            self._stoplicht(GROEN, "Alle bonnen matchen — administratie compleet.", ouder=f)

    def open_excel(self):
        if self.laatste_output_pad:
            try:
                os.startfile(self.laatste_output_pad)
            except OSError as fout:
                messagebox.showerror(VENSTER_TITEL, self._fout_melding(fout))

    def open_export(self):
        if self.laatste_export_pad:
            os.startfile(self.laatste_export_pad)

    # ---- bonnen (achtergrond) ----
    def kies_bonnenmap(self):
        map_pad = filedialog.askdirectory(title="Kies map met bonnen (PDF's en/of foto's)")
        if not map_pad:
            return
        self.bonnen_knop.configure(state="disabled")
        self.csv_knop.configure(state="disabled")
        self.voortgang_balk["value"] = 0
        self.voortgang_label.configure(text="Bezig…")
        self.voortgang_label.pack(padx=16, pady=(0, 0), anchor="w")
        self.voortgang_balk.pack(padx=16, pady=(2, 8))

        def voortgang(huidig, totaal, bestandsnaam):
            self._wachtrij.put(("voortgang", huidig, totaal, bestandsnaam))

        def werk():
            from bonnen_matcher import analyseer_bonnenmap
            try:
                analyse = analyseer_bonnenmap(Path(map_pad), self.laatste_transacties, voortgang)
                self._wachtrij.put(("ok", analyse, None, None))
            except Exception as fout:
                self._wachtrij.put(("fout", fout, None, None))

        threading.Thread(target=werk, daemon=True).start()
        self.after(100, self._verwerk_wachtrij)

    def _verwerk_wachtrij(self):
        try:
            while True:
                bericht = self._wachtrij.get_nowait()
                status = bericht[0]
                if status == "voortgang":
                    _, huidig, totaal, naam = bericht
                    self.voortgang_balk["maximum"] = totaal
                    self.voortgang_balk["value"] = huidig
                    self.voortgang_label.configure(text=f"Bezig: {huidig}/{totaal} — {naam}")
                    continue
                self._bonnen_klaar(status, bericht[1])
                return
        except queue.Empty:
            self.after(100, self._verwerk_wachtrij)

    def _bonnen_klaar(self, status, resultaat):
        self.voortgang_label.pack_forget()
        self.voortgang_balk.pack_forget()
        self.csv_knop.configure(state="normal")
        self.bonnen_knop.configure(state="normal")

        if status == "fout":
            self._toon_fout(f"Bonnen vergelijken mislukt:\n{resultaat}")
            return

        analyse = resultaat
        bouw_bonnencontrole_sheet(self.laatste_wb, analyse)
        try:
            self.laatste_wb.save(self.laatste_output_pad)
        except OSError as fout:
            messagebox.showerror(VENSTER_TITEL, self._fout_melding(fout))
            return
        self._toon_bonnen(analyse)


if __name__ == "__main__":
    App(sys.argv[1] if len(sys.argv) > 1 else None).mainloop()
