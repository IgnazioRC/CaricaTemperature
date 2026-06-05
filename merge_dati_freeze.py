#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modulo: merge_dati_freeze
Funzioni stabili per il merge dei dati CSV in DataFrame unificato
"""

import pandas as pd

# =========================================================
# 🧊 FUNZIONE CONGELATA — Merge dati CSV 1.4
# =========================================================
def merge_dati_1_4(lista_dizionari):
    """Unisce più dizionari {timestamp: valore} in un DataFrame unico."""
    df = pd.DataFrame(lista_dizionari).fillna(pd.NA)
    df = df.sort_index()
    df = df.reset_index().rename(columns={"index": "Data"})
    return df