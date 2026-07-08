#!/usr/bin/env python3
"""Simpel venster om een ING csv-export te verwerken. Bouwt voort op ing_btw_verwerker.py."""
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ing_btw_verwerker import bouw_bonnencontrole_sheet, verwerk_bestand

VENSTER_TITEL = "ING Btw Verwerker"


class App(tk.Tk):
    def __init__(self, startbestand: str = None):
        super().__init__()
        self.title(VENSTER_TITEL)
        self.geometry("560x460")
        self.resizable(False, False)
        self.laatste_output_pad = None
        self.laatste_transacties = None
        self.laatste_wb = None
        self._wachtrij = queue.Queue()

        tk.Label(
            self, text="Verwerk een ING csv-export naar een btw/inkomsten-uitgaven Excel-overzicht.",
            wraplength=520, justify="left", font=("Segoe UI", 10),
        ).pack(padx=16, pady=(16, 8), anchor="w")

        self.csv_knop = tk.Button(
            self, text="Kies ING csv-bestand...", command=self.kies_bestand,
            font=("Segoe UI", 10), padx=10, pady=6,
        )
        self.csv_knop.pack(padx=16, pady=4, anchor="w")

        self.resultaat_tekst = tk.Text(
            self, height=13, width=68, font=("Consolas", 9), state="disabled",
            bg="#f4f4f4", relief="flat",
        )
        self.resultaat_tekst.pack(padx=16, pady=8)

        self.voortgang_label = tk.Label(self, text="", font=("Segoe UI", 9), fg="#555555")
        self.voortgang_balk = ttk.Progressbar(self, length=520, mode="determinate")

        knoppen_frame = tk.Frame(self)
        knoppen_frame.pack(padx=16, pady=(0, 16), anchor="w")

        self.open_knop = tk.Button(
            knoppen_frame, text="Open Excel-bestand", command=self.open_resultaat,
            font=("Segoe UI", 10), padx=10, pady=6, state="disabled",
        )
        self.open_knop.grid(row=0, column=0, padx=(0, 8))

        self.bonnen_knop = tk.Button(
            knoppen_frame, text="Vergelijk met bonnenmap...", command=self.kies_bonnenmap,
            font=("Segoe UI", 10), padx=10, pady=6, state="disabled",
        )
        self.bonnen_knop.grid(row=0, column=1)

        if startbestand:
            self.verwerk(Path(startbestand))

    def _schrijf(self, tekst: str):
        self.resultaat_tekst.configure(state="normal")
        self.resultaat_tekst.delete("1.0", "end")
        self.resultaat_tekst.insert("1.0", tekst)
        self.resultaat_tekst.configure(state="disabled")

    def _voeg_toe(self, tekst: str):
        self.resultaat_tekst.configure(state="normal")
        self.resultaat_tekst.insert("end", tekst)
        self.resultaat_tekst.see("end")
        self.resultaat_tekst.configure(state="disabled")

    def kies_bestand(self):
        pad = filedialog.askopenfilename(
            title="Kies ING csv-bestand", filetypes=[("CSV bestanden", "*.csv"), ("Alle bestanden", "*.*")],
        )
        if pad:
            self.verwerk(Path(pad))

    def verwerk(self, csv_pad: Path):
        self.open_knop.configure(state="disabled")
        self.bonnen_knop.configure(state="disabled")
        self.laatste_output_pad = None
        self.laatste_transacties = None
        self.laatste_wb = None
        try:
            resultaat = verwerk_bestand(csv_pad)
        except SystemExit as fout:
            messagebox.showerror(VENSTER_TITEL, str(fout))
            self._schrijf(f"Mislukt: {fout}")
            return
        except Exception as fout:
            messagebox.showerror(VENSTER_TITEL, f"Onverwachte fout:\n{fout}")
            self._schrijf(f"Mislukt: {fout}")
            return

        regels = [
            f"Klaar: {resultaat['aantal_verwerkt']} transacties verwerkt.",
            "",
            f"Uitgesloten: {resultaat['aantal_prive']} prive/spaar, "
            f"{resultaat['aantal_belastingdienst']} belastingdienst,",
            f"             {resultaat['aantal_geen_factuur']} inkomsten zonder herkenbare factuur.",
            f"Waarvan {resultaat['aantal_kredietfactuur']} kredietfacturen omgezet naar uitgave.",
        ]
        if resultaat["aantal_buiten_periode"]:
            regels.append(
                f"\nLet op: {resultaat['aantal_buiten_periode']} factuur/facturen buiten de "
                f"periode van dit bestand - zie Overzicht-tab."
            )
        if resultaat["geen_periode_gevonden"]:
            regels.append(
                "\nWaarschuwing: kon geen periode uit de bestandsnaam halen - "
                "factuurherkenning accepteert nu elke jaarmaand."
            )
        regels.append(f"\nOpgeslagen als:\n{resultaat['output_pad']}")
        self._schrijf("\n".join(regels))

        self.laatste_output_pad = resultaat["output_pad"]
        self.laatste_transacties = resultaat["transacties"]
        self.laatste_wb = resultaat["workbook"]
        self.open_knop.configure(state="normal")
        self.bonnen_knop.configure(state="normal")

    def open_resultaat(self):
        if self.laatste_output_pad:
            os.startfile(self.laatste_output_pad)

    def kies_bonnenmap(self):
        map_pad = filedialog.askdirectory(title="Kies map met bonnen (PDF's)")
        if not map_pad:
            return
        self.bonnen_knop.configure(state="disabled")
        self.csv_knop.configure(state="disabled")
        self._voeg_toe("\n\nBonnen vergelijken...")

        self.voortgang_balk["value"] = 0
        self.voortgang_label.configure(text="Bezig...")
        self.voortgang_label.pack(padx=16, pady=(4, 0), anchor="w")
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
                    _, huidig, totaal, bestandsnaam = bericht
                    self.voortgang_balk["maximum"] = totaal
                    self.voortgang_balk["value"] = huidig
                    self.voortgang_label.configure(text=f"Bezig: {huidig}/{totaal} - {bestandsnaam}")
                    continue

                self._bonnen_klaar(status, bericht[1])
                return
        except queue.Empty:
            self.after(100, self._verwerk_wachtrij)

    def _bonnen_klaar(self, status: str, resultaat):
        self.voortgang_label.pack_forget()
        self.voortgang_balk.pack_forget()
        self.csv_knop.configure(state="normal")
        self.bonnen_knop.configure(state="normal")

        if status == "fout":
            messagebox.showerror(VENSTER_TITEL, f"Bonnen vergelijken mislukt:\n{resultaat}")
            self._voeg_toe(f"\nMislukt: {resultaat}")
            return

        analyse = resultaat
        bouw_bonnencontrole_sheet(self.laatste_wb, analyse)
        self.laatste_wb.save(self.laatste_output_pad)

        self._voeg_toe(
            f"\n\nBonnencontrole klaar ({analyse['totaal_bonnen']} bonnen):\n"
            f"  {len(analyse['zeker'])} zeker gematcht\n"
            f"  {len(analyse['onzeker'])} onzeker - controleer zelf\n"
            f"  {len(analyse['geen_match'])} geen match gevonden\n"
            f"  {len(analyse['onleesbaar'])} onleesbaar\n"
            f"  {len(analyse['uitgaven_zonder_bon'])} uitgaven zonder bon\n"
            f"Zie tabblad 'Bonnencontrole' in het Excel-bestand."
        )


if __name__ == "__main__":
    App(sys.argv[1] if len(sys.argv) > 1 else None).mainloop()
