#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modulo: lettura_csv_freeze
Contiene le funzioni stabili (freeze) per la lettura dei file CSV
Versione congelata: leggi_csv_4_0 — testata e stabile
"""

import pandas as pd
from datetime import timedelta

# =========================================================
# 🧊 FUNZIONE CONGELATA — Lettura CSV (v4.0)
# =========================================================
def leggi_csv_4_0(path):
    """Lettura CSV con offset dinamico — versione stabile congelata"""
    with open(path, 'r', encoding='utf-8') as f:
        first = f.readline()
    sep = '\t' if '\t' in first else ','

    df = pd.read_csv(
        path,
        sep=sep,
        skiprows=1,
        usecols=[0, 1],
        names=["timestamp", "temp"],
        engine="python"
    )

    df["timestamp"] = pd.to_datetime(
        df["timestamp"],
        errors="coerce",
        format="%Y-%m-%d %H:%M:%S"
    )
    df = df.dropna(subset=["timestamp", "temp"])

    dati = {}
    for _, r in df.iterrows():
        ts = r["timestamp"]
        # offset dinamico: se il file parte dalle 02:01 → sposta -1 ora
        if ts.hour == 2 and ts.minute == 1:
            ts -= timedelta(hours=1)
        dati[ts] = float(str(r["temp"]).replace(",", "."))
    return dati