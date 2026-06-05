#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modulo: gui_freeze
GUI stabile — versione 1.4.2 corretta (log in basso, CSV1–3)
Autore: Ignazio Rusconi-Clerici — macOS
"""

# --- IRC shared bootstrap ---
# Rende disponibili i moduli in Python/shared/ senza dipendere da PYTHONPATH.
# Saltato se eseguito da bundle PyInstaller (sys.frozen=True): in quel caso
# i moduli sono gia' inclusi nel bundle.
import sys as _sys
from pathlib import Path as _Path
if not getattr(_sys, 'frozen', False):
    _shared = _Path.home() / "Library/CloudStorage/Dropbox/Documenti_IRC/Python/shared"
    if str(_shared) not in _sys.path:
        _sys.path.insert(0, str(_shared))
# --- end IRC shared bootstrap ---


import tkinter as tk
from tkinter import ttk
from path_widgets import PathLabel

# =========================================================
# 🧊 FUNZIONE CONGELATA — GUI 1.4.2 (corretta)
# =========================================================
def crea_gui_1_4_2(app):
    """
    Crea l'interfaccia grafica base per CaricaTemperature.
    Include:
      • Selettore Excel
      • Tre slot CSV (CSV 1, CSV 2, CSV 3)
      • Checkbox "Attiva log"
      • Barra di avanzamento
      • Pulsanti "Avvia caricamento" e "Esci"
      • Finestra log in fondo alla finestra
    """
    pad = {"padx": 12, "pady": 6}

    # --- Selettore Excel ---
    frm_excel = ttk.Frame(app)
    frm_excel.pack(fill="x", **pad)
    ttk.Label(frm_excel, text="File Excel di destinazione:").pack(anchor="w")
    row_excel = ttk.Frame(frm_excel)
    row_excel.pack(fill="x")
    app.lbl_excel = PathLabel(row_excel, text="(non selezionato)", foreground="#666")
    app.lbl_excel.pack(side="left", fill="x", expand=True)
    ttk.Button(row_excel, text="Scegli Excel", command=app.choose_excel).pack(side="right")

    ttk.Separator(app).pack(fill="x", padx=12, pady=4)

    # --- Selettori CSV ---
    frm_csv = ttk.Frame(app)
    frm_csv.pack(fill="x", **pad)

    def make_row(parent, key, label):
        f = ttk.Frame(parent)
        f.pack(fill="x", pady=2)
        ttk.Label(f, text=f"{label}:").pack(side="left")
        val = PathLabel(f, text="(non selezionato)", foreground="#666")
        val.pack(side="left", fill="x", expand=True)
        ttk.Button(f, text="Scegli CSV", command=lambda: app.choose_csv(key, val)).pack(side="right")
        return val

    app.lbl_csv1 = make_row(frm_csv, "CSV1", "CSV 1")
    app.lbl_csv2 = make_row(frm_csv, "CSV2", "CSV 2")
    app.lbl_csv3 = make_row(frm_csv, "CSV3", "CSV 3")

    ttk.Separator(app).pack(fill="x", padx=12, pady=4)

    # --- Checkbox log ---
    frm_opts = ttk.Frame(app)
    frm_opts.pack(fill="x", **pad)
    app.log_active = tk.BooleanVar(value=True)
    ttk.Checkbutton(frm_opts, text="Attiva log", variable=app.log_active).pack(anchor="w")

    # --- Barra di avanzamento ---
    frm_status = ttk.Frame(app)
    frm_status.pack(fill="x", **pad)
    app.status_var = tk.StringVar(value="In attesa…")
    ttk.Label(frm_status, textvariable=app.status_var).pack(anchor="w")
    app.pbar = ttk.Progressbar(frm_status, mode="determinate", maximum=100, value=0)
    app.pbar.pack(fill="x")

   # --- Area log (in fondo, sopra i bottoni) ---
    ttk.Separator(app).pack(fill="x", padx=12, pady=4)
    frm_log = ttk.Frame(app)
    frm_log.pack(fill="both", expand=True, **pad)
    ttk.Label(frm_log, text="Log messaggi:").pack(anchor="w")
    app.txt_log = tk.Text(frm_log, height=8, wrap="word", bg="#f9f9f9")
    app.txt_log.pack(fill="both", expand=True)

    # --- Bottoni finali ---
    ttk.Separator(app).pack(fill="x", padx=12, pady=4)
    frm_run = ttk.Frame(app)
    frm_run.pack(fill="x", **pad)
    app.btn_run = ttk.Button(frm_run, text="Avvia caricamento", command=app.run)
    app.btn_run.pack(side="left", padx=(0, 8))
    ttk.Button(frm_run, text="Esci", command=app.destroy).pack(side="right")


    
    return app