# Carica Temperature

Modulo principale del sistema **Caricamento Temperature**, versione 1.4.2 stabile.

**Autore:** Ignazio Rusconi-Clerici  
**Versione:** 1.4.2

---

## Descrizione generale

Gestisce l'interfaccia utente, la selezione dei file di input/output e l'esecuzione coordinata dei moduli congelati:

| Modulo | Versione | Funzione |
|--------|----------|----------|
| `gui_freeze.py` | 1.4.2 | Interfaccia Tkinter stabile |
| `lettura_csv_freeze.py` | 4.0 | Lettura robusta CSV |
| `merge_dati_freeze.py` | 1.0 | Fusione automatica dataset |
| `scrittura_excel_freeze.py` | 1.4 | Scrittura su Excel con grafici |

---

## Configurazione

```
~/Library/CloudStorage/Dropbox/Documenti_IRC/Python/_config/apps/CaricaTemperature/config.json
```

Contiene gli ultimi percorsi usati per Excel e CSV.

---

## Flusso operativo

1. Selezione file Excel di destinazione
2. Lettura automatica delle intestazioni dalle celle **B1–D1** del primo foglio
3. Aggiornamento etichette dei pulsanti CSV nella GUI
4. Selezione fino a tre file CSV
5. Avvio elaborazione in thread separato: Lettura CSV → Merge → Scrittura Excel
6. Log automatico su GUI e su file `.log` accanto all'Excel

---

## Parametri CLI

```bash
python3 CaricaTemperature_main.py
python3 CaricaTemperature_main.py --base_path /percorso/cartella
```

---

## Dipendenze

```bash
pip install openpyxl
```

---

## Novità versione 1.4.2

- Lettura automatica intestazioni Excel (celle B1–D1)
- Aggiornamento dinamico etichette CSV nella GUI
- Reset path CSV quando si seleziona un nuovo Excel
- Log su file accanto all'Excel selezionato
