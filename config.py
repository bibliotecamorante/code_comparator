"""
Configurazione globale dell'applicazione
"""
from pathlib import Path

# ============================================================
# PERCORSI
# ============================================================
try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

CONFIG_FILE = str(BASE_DIR / "presets.json")

# ============================================================
# COSTANTI GLOBALI
# ============================================================
MAX_COLUMNS = 7
DEFAULT_EXCEL_MIN_ROWS = 100
DEFAULT_ZOOM_LEVEL = 110
MIN_COLUMNS_PARSE_THRESHOLD = 0.5

# ============================================================
# DIMENSIONI UI
# ============================================================
BUTTON_HEIGHT = 35
COMPARE_BUTTON_HEIGHT = 60
SETTINGS_BUTTON_WIDTH = 140
SETTINGS_BUTTON_HEIGHT = 40
DEFAULT_ROW_HEIGHT = 30
HEADER_ROW_HEIGHT = 19

# ============================================================
# COLORI
# ============================================================
HEADER_BG_COLOR = "#ecf0f1"
HEADER_TEXT_COLOR = "#2c3e50"
HEADER_BORDER_COLOR = "#bdc3c7"
HIGHLIGHT_COLOR = "#fff3cd"
MATCH_COLOR = "#e8f8f5"
MISMATCH_COLOR = "#fdedec"
WHITE_COLOR = "#ffffff"
DUPLICATE_HIGHLIGHT = "#ffeaa7"
DUPLICATE_KEY_COLOR = "#ff7675"

# ============================================================
# EXCEL EXPORT - LARGHEZZE COLONNE
# ============================================================
EXCEL_COLUMN_WIDTHS = {
    65: ["TITOLO", "TITLE", "DESCRIZIONE", "DESCRIPTION", "ISBD"],
    25: ["NOTE", "NOTES", "LEGAMI", "LINKS", "COMMENTI", "COMMENTS"],
    23: ["AUTORE", "AUTHOR", "EDITORE", "PUBLISHER", "COLLOCAZIONE", "LOCATION"],
    15: ["ISBN", "CODICE", "CODE", "INVENTARIO", "INVENTORY", "ID"],
    12: ["ANNO", "YEAR", "PREZZO", "PRICE", "SEQUENZA", "SEZIONE", "SECTION", "SPECIFICAZIONE"]
}

# ============================================================
# PRESET PREDEFINITI (PROTETTI)
# ============================================================
DEFAULT_PRESETS = {
    "ACQUISTI (ISBN, Titolo, Autore, Editore, Anno, Prezzo, Note)": {
        "key_column": "ISBN",
        "columns": ["ISBN", "TITOLO", "AUTORE", "EDITORE", "ANNO", "PREZZO", "NOTE"]
    },
    "ACQUISTI 2 (ISBN, Titolo)": {
        "key_column": "ISBN",
        "columns": ["ISBN", "TITOLO"]
    },
    "STAMPA REGISTRI 1 (Inventario, Sezione, Collocazione, Specificazione, Sequenza, Descrizione ISBD, Legami)": {
        "key_column": "Inventario",
        "columns": ["Inventario", "Sezione", "Collocazione", "Specificazione", 
                   "Sequenza", "Descrizione ISBD", "Legami"]
    },
    "STAMPA REGISTRI 2 (Inventario, Descrizione ISBD)": {
        "key_column": "Inventario",
        "columns": ["Inventario", "Descrizione ISBD"]
    }
}

# Set di preset che non possono essere eliminati
PROTECTED_PRESETS = set(DEFAULT_PRESETS.keys())