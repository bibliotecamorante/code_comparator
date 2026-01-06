import sys
import os
import pandas as pd
import io
import csv  # ← NUOVO: necessario per gestire le virgolette
import json
import shutil
import logging
from pathlib import Path


from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QTableWidget, QTableWidgetItem, 
                             QPushButton, QLabel, QHeaderView, QTabWidget,
                             QMessageBox, QFileDialog, QMenu, QDialog, 
                             QLineEdit, QComboBox, QFormLayout, QDialogButtonBox,
                             QGroupBox, QScrollArea)
from PyQt6.QtCore import Qt, QStandardPaths
from PyQt6.QtGui import QColor, QKeySequence, QFont
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

# Log configuration (errors only, visible in console)
logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


# ============================================================
# GLOBAL EXCEPTION HANDLER
# ============================================================
def global_exception_handler(exctype, value, tb):
    """Gestore di eccezioni globale per errori non gestiti"""
    logging.exception("Errore non gestito:", exc_info=(exctype, value, tb))
    
    # Mostra dialog solo se QApplication è attiva
    if QApplication.instance():
        QMessageBox.critical(
            None, 
            "❌ Errore Critico",
            f"Si è verificato un errore imprevisto:\n\n{value}\n\n"
            "L'applicazione potrebbe non funzionare correttamente.\n"
            "Controlla il log per maggiori dettagli."
        )
    
    # Chiama il gestore predefinito per mantenere il traceback
    sys.__excepthook__(exctype, value, tb)

sys.excepthook = global_exception_handler


# ============================================================
# GLOBAL CONSTANTS
# ============================================================
# Colors
HEADER_BG_COLOR = "#ecf0f1"
HEADER_TEXT_COLOR = "#2c3e50"
HEADER_BORDER_COLOR = "#bdc3c7"
HIGHLIGHT_COLOR = "#fff3cd"
MATCH_COLOR = "#e8f8f5"
MISMATCH_COLOR = "#fdedec"
WHITE_COLOR = "#ffffff"
DUPLICATE_HIGHLIGHT = "#ffeaa7"
DUPLICATE_KEY_COLOR = "#ff7675"

# Dimensions
BUTTON_HEIGHT = 35
COMPARE_BUTTON_HEIGHT = 60
SETTINGS_BUTTON_WIDTH = 140
SETTINGS_BUTTON_HEIGHT = 40
DEFAULT_ROW_HEIGHT = 30
HEADER_ROW_HEIGHT = 19

# Excel Export Column Widths (keyword-based mapping)
EXCEL_COLUMN_WIDTHS = {
    65: ["TITOLO", "TITLE", "DESCRIZIONE", "DESCRIPTION", "ISBD"],
    25: ["NOTE", "NOTES", "LEGAMI", "LINKS", "COMMENTI", "COMMENTS"],
    23: ["AUTORE", "AUTHOR", "EDITORE", "PUBLISHER", "COLLOCAZIONE", "LOCATION"],
    15: ["ISBN", "CODICE", "CODE", "INVENTARIO", "INVENTORY", "ID"],
    12: ["ANNO", "YEAR", "PREZZO", "PRICE", "SEQUENZA", "SEZIONE", "SECTION", "SPECIFICAZIONE"]
}

MAX_COLUMNS = 7

# ============================================================
# GLOBAL CSS STYLES
# ============================================================
APP_STYLESHEET = """
    QMainWindow {
        background-color: #f5f6fa;
    }
    
    /* STILE BASE PULSANTI */
    QPushButton {
        background-color: #3498db;
        color: white;
        border: none;
        padding: 10px 20px;
        font-weight: bold;
        border-radius: 5px;
        font-size: 13px;
    }
    QPushButton:hover {
        background-color: #2980b9;
    }
    QPushButton:pressed {
        background-color: #21618c;
    }
    QPushButton:disabled {
        background-color: #bdc3c7;
        color: #7f8c8d;
    }
    
    /* PULSANTI SPECIALIZZATI (solo colori diversi) */
    QPushButton#btn_compare {
        background-color: #9b59b6;
        color: white;
        font-size: 16px;
        border: 3px solid #6c3483;
    }
    QPushButton#btn_compare:hover {
        background-color: #8e44ad;
        border: 3px solid #5b2c6f;
    }

    QPushButton#btn_export_matches {
        background-color: #27ae60;
        color: white;
    }
    QPushButton#btn_export_matches:hover {
        background-color: #229954;
    }

    QPushButton#btn_export_mismatches {
        background-color: #e67e22;
        color: white;
    }
    QPushButton#btn_export_mismatches:hover {
        background-color: #d35400;
    }

    QPushButton#btn_clear_results {
        background-color: #95a5a6;
        color: white;
    }
    QPushButton#btn_clear_results:hover {
        background-color: #7f8c8d;
    }

    QPushButton#btn_settings {
        background-color: #34495e;
        color: white;
    }
    QPushButton#btn_settings:hover {
        background-color: #2c3e50;
    }

    QPushButton#btn_swap {
        background-color: #f39c12;
        color: white;
    }
    QPushButton#btn_swap:hover {
        background-color: #e67e22;
    }
    
    /* TABELLE */
    QTableWidget {
        border: 1px solid #bdc3c7;
        border-radius: 5px;
        background-color: white;
        gridline-color: #ecf0f1;
        selection-background-color: #cce5ff;
        selection-color: #000000;
    }
    
    QHeaderView::section {
        background-color: #ecf0f1;
        color: #2c3e50;
        padding: 10px;
        border: 1px solid #bdc3c7;
        font-weight: bold;
        font-size: 13px;
    }
    QHeaderView::section:hover {
        background-color: #d5dbdb;
    }
    
    /* TABS */
    QTabWidget::pane {
        border: 1px solid #bdc3c7;
        border-radius: 5px;
        background-color: white;
    }
    QTabBar::tab {
        background-color: #ecf0f1;
        color: #2c3e50;
        padding: 10px 20px;
        border-top-left-radius: 5px;
        border-top-right-radius: 5px;
        margin-right: 2px;
        font-weight: bold;
    }
    QTabBar::tab:selected {
        background-color: #3498db;
        color: white;
    }
    QTabBar::tab:hover {
        background-color: #bdc3c7;
    }
"""


class SettingsDialog(QDialog):
    """Finestra di dialogo per gestire i preset di configurazione"""
    def __init__(self, presets, current_preset, parent=None):
        super().__init__(parent)
        self.presets = presets
        self.current_preset = current_preset
        self.result = None
        
        self.setWindowTitle("⚙️ Impostazioni Comparatore")
        self.resize(350, 700)  
        
        self.init_ui()
        
    @staticmethod
    def validate_preset_data(preset_name, key_column, columns):
        """
        Validazione unificata per i dati del preset.
        Returns: (is_valid, error_message)
        """
        if not preset_name or not preset_name.strip():
            return False, "Il nome del preset è obbligatorio"
        
        if not key_column or not key_column.strip():
            return False, "Il nome della colonna chiave è obbligatorio"
        
        if not isinstance(columns, list) or len(columns) < 1:
            return False, "Deve esserci almeno una colonna"
        
        if len(columns) > MAX_COLUMNS:
            return False, f"Massimo {MAX_COLUMNS} colonne consentite"
        
        if key_column not in columns:
            return False, f"La colonna chiave '{key_column}' deve essere presente nelle colonne"
        
        return True, ""
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        # --- SEZIONE: SELEZIONA PRESET ---
        group_select = QGroupBox("📋 Seleziona Preset Esistente")
        group_select_layout = QVBoxLayout()
        
        self.combo_presets = QComboBox()
        self.combo_presets.addItems(self.presets.keys())
        self.combo_presets.setCurrentText(self.current_preset)
        self.combo_presets.currentTextChanged.connect(self.on_preset_changed)
        
        # Label per mostrare dettagli preset
        self.label_preset_details = QLabel()
        self.label_preset_details.setStyleSheet("padding: 10px; background-color: #ecf0f1; border-radius: 5px;")
        self.label_preset_details.setWordWrap(True)
        self.update_preset_details()
        
        btn_select_preset = QPushButton("✓ Usa Questo Preset")
        btn_select_preset.clicked.connect(self.select_existing_preset)
        
        group_select_layout.addWidget(QLabel("Preset disponibili:"))
        group_select_layout.addWidget(self.combo_presets)
        group_select_layout.addWidget(self.label_preset_details)
        group_select_layout.addWidget(btn_select_preset)
        group_select.setLayout(group_select_layout)
        
        layout.addWidget(group_select)
        
        # --- SEPARATORE ---
        separator = QLabel("━" * 80)
        separator.setAlignment(Qt.AlignmentFlag.AlignCenter)
        separator.setStyleSheet("color: #bdc3c7;")
        layout.addWidget(separator)
        
        # --- SEZIONE: CREA NUOVO PRESET ---
        group_create = QGroupBox("➕ Crea Nuovo Preset")
        group_create_layout = QFormLayout()

        self.input_preset_name = QLineEdit()
        self.input_preset_name.setPlaceholderText("Es: Stampa Registri, Gestione Inventari")

        # 7 campi per le colonne (il primo è la colonna chiave)
        # Campo Nome Preset
        preset_label = QLabel("Nome Preset*:")
        preset_label.setStyleSheet("font-weight: bold;")
        group_create_layout.addRow(preset_label, self.input_preset_name)
        
        # Campi per la colonna chiave (affiancati con uguale larghezza)
        key_widget = QWidget()
        key_layout = QHBoxLayout(key_widget)
        key_layout.setContentsMargins(0, 0, 0, 0)
        key_layout.setSpacing(10)
        
        # Layout verticale per Nome colonna chiave
        name_layout = QVBoxLayout()
        name_layout.setSpacing(5)
        name_label = QLabel("Nome colonna chiave*:")
        name_label.setStyleSheet("font-weight: bold;")
        self.input_key_name = QLineEdit()
        self.input_key_name.setPlaceholderText("Es: ISBN, Inventario...")
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.input_key_name)
        
        # Layout verticale per Numero colonna chiave
        number_layout = QVBoxLayout()
        number_layout.setSpacing(5)
        number_label = QLabel("Numero colonna chiave*:")
        number_label.setStyleSheet("font-weight: bold;")
        self.input_key_number = QLineEdit()
        self.input_key_number.setPlaceholderText(f"Da 1 a {MAX_COLUMNS}")
        number_layout.addWidget(number_label)
        number_layout.addWidget(self.input_key_number)
        
        # Aggiungi i due layout affiancati con uguale larghezza
        key_layout.addLayout(name_layout, 1)
        key_layout.addLayout(number_layout, 1)
        
        # Connetti i campi chiave per auto-valorizzazione
        self.input_key_name.textChanged.connect(self.update_key_column)
        self.input_key_number.textChanged.connect(self.update_key_column)
        
        group_create_layout.addRow("", key_widget)
        
        # Separatore visivo
        separator_key = QLabel("─" * 60)
        separator_key.setAlignment(Qt.AlignmentFlag.AlignCenter)
        separator_key.setStyleSheet("color: #bdc3c7;")
        group_create_layout.addRow("", separator_key)
        
        # 7 campi per le colonne (tutte opzionali)
        self.column_inputs = []
        for i in range(MAX_COLUMNS):
            line_edit = QLineEdit()
            line_edit.setPlaceholderText("Opzionale")
            label_text = f"Colonna {i+1}:"
            
            self.column_inputs.append(line_edit)
            group_create_layout.addRow(label_text, line_edit)
        
        btn_create_preset = QPushButton("✓ Crea e Usa Nuovo Preset")
        btn_create_preset.clicked.connect(self.create_new_preset)
        
        group_create.setLayout(group_create_layout)
        
        # Scroll area per il form
        scroll = QScrollArea()
        scroll.setWidget(group_create)
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(400)
        
        layout.addWidget(scroll)
        layout.addWidget(btn_create_preset)
        
        # --- PULSANTE CHIUDI ---
        btn_close = QPushButton("✕ Chiudi")
        btn_close.clicked.connect(self.reject)
        layout.addWidget(btn_close)
        
    def update_key_column(self):
        """Auto-valorizza il campo colonna in base a Nome e Numero colonna chiave"""
        key_name = self.input_key_name.text().strip()
        key_number_text = self.input_key_number.text().strip()
        
        # Cancella il valore precedente se esisteva
        if hasattr(self, '_last_key_number') and self._last_key_number is not None:
            if 1 <= self._last_key_number <= 7:
                # Cancella solo se il valore corrente corrisponde al nome chiave precedente
                old_value = self.column_inputs[self._last_key_number - 1].text().strip()
                if hasattr(self, '_last_key_name') and old_value == self._last_key_name:
                    self.column_inputs[self._last_key_number - 1].clear()
        
        # Valida il numero colonna
        if not key_number_text or not key_name:
            return
        
        try:
            key_number = int(key_number_text)
            if 1 <= key_number <= MAX_COLUMNS:
                # Valorizza automaticamente il campo corrispondente
                self.column_inputs[key_number - 1].setText(key_name)
                # Memorizza l'ultima posizione e nome
                self._last_key_number = key_number
                self._last_key_name = key_name
        except ValueError:
            pass
    
    def on_preset_changed(self, preset_name):
        """Aggiorna i dettagli quando si cambia preset"""
        self.update_preset_details()
    
    def update_preset_details(self):
        """Mostra i dettagli del preset selezionato"""
        preset_name = self.combo_presets.currentText()
        preset_data = self.presets[preset_name]
        
        columns_str = ", ".join(preset_data["columns"])
        details = f"<b>Colonna Chiave:</b> {preset_data['key_column']}<br>"
        details += f"<b>Colonne:</b> {columns_str}"
        
        self.label_preset_details.setText(details)
    
    def select_existing_preset(self):
        """Seleziona un preset esistente"""
        self.result = {
            'action': 'select',
            'preset_name': self.combo_presets.currentText()
        }
        self.accept()
    
    def create_new_preset(self):
        """Crea un nuovo preset personalizzato"""
        preset_name = self.input_preset_name.text().strip()
        key_column = self.input_key_name.text().strip()
        key_number_text = self.input_key_number.text().strip()
        
        # Valida Numero colonna chiave
        if not key_number_text:
            QMessageBox.warning(self, "Campo Obbligatorio", 
                              "Il Numero colonna chiave è obbligatorio!")
            return
        
        try:
            key_number = int(key_number_text)
            if not (1 <= key_number <= MAX_COLUMNS):
                QMessageBox.warning(
                    self,
                    "Valore Non Valido",
                    f"Il Numero colonna chiave deve essere compreso tra 1 e {MAX_COLUMNS}!"
                )
                return

        except ValueError:
            QMessageBox.warning(self, "Valore Non Valido", 
                              "Il Numero colonna chiave deve essere un numero!")
            return
        
        # Raccogli tutte le colonne valorizzate
        columns = []
        for i in range(MAX_COLUMNS):
            col_name = self.column_inputs[i].text().strip()
            if col_name:
                columns.append(col_name)
        
        # Usa validazione unificata
        is_valid, error_msg = self.validate_preset_data(preset_name, key_column, columns)
        if not is_valid:
            QMessageBox.warning(self, "Validazione Fallita", error_msg)
            return
        
        # Controlla se esiste già
        if preset_name in self.presets:
            reply = QMessageBox.question(
                self,
                "Preset Esistente",
                f"Il preset '{preset_name}' esiste già. Vuoi sovrascriverlo?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
        
        self.result = {
            'action': 'create',
            'preset_name': preset_name,
            'key_column': key_column,
            'columns': columns
        }
        self.accept()
    
    def get_result(self):
        """Restituisce il risultato della dialog"""
        return self.result


class ComparisonEngine:
    """Motore di confronto logico separato dalla UI"""
    
    @staticmethod
    def count_duplicates(df: pd.DataFrame, key_column: str = 'KEY_NORMALIZED') -> tuple[int, int]:
        """
        Conta duplicati in un DataFrame.
        
        Args:
            df: DataFrame con colonna chiave
            key_column: Nome colonna chiave (default: 'KEY_NORMALIZED')
        
        Returns:
            tuple: (unique_keys_duplicated, total_rows_duplicated)
        """
        if df.empty or key_column not in df.columns:
            return 0, 0
        
        # Filtra righe con chiave valida
        valid_mask = df[key_column].notna() & (df[key_column].astype(str).str.strip() != "")
        df_valid = df[valid_mask]
        
        if df_valid.empty:
            return 0, 0
        
        duplicates_mask = df_valid[key_column].duplicated(keep=False)
        unique_keys = df_valid[duplicates_mask][key_column].nunique()
        total_rows = duplicates_mask.sum()
        
        return unique_keys, total_rows
    
    @staticmethod
    def compare(df1, df2, key_column='KEY_NORMALIZED'):
        """
        Confronta due DataFrame basandosi sulla colonna chiave normalizzata.
        
        Args:
            df1: DataFrame principale (worklist)
            df2: DataFrame di confronto
            key_column: Nome della colonna normalizzata (default: 'KEY_NORMALIZED')
        
        Returns:
            tuple: (matches_df, mismatches_df, stats_dict)
        """
        # Estrai le chiavi normalizzate
        keys_df2 = set(df2[key_column].dropna())
        
        # Identifica corrispondenze e mancanti
        matches_mask = df1[key_column].isin(keys_df2)
        matches = df1[matches_mask].copy()
        mismatches = df1[~matches_mask].copy()
        
        # Calcola statistiche usando il metodo centralizzato
        duplicates_df1, _ = ComparisonEngine.count_duplicates(df1, key_column)
        duplicates_df2, _ = ComparisonEngine.count_duplicates(df2, key_column)
        
        stats = {
            'total': len(df1),
            'matches': len(matches),
            'mismatches': len(mismatches),
            'duplicates_df1': duplicates_df1,
            'duplicates_df2': duplicates_df2
        }
        
        return matches, mismatches, stats


class TableManager:
    """Gestore centralizzato per la configurazione e manipolazione delle tabelle"""
    
    def __init__(self, columns=None):
        self.columns = columns if columns is not None else []
        self._sort_column = None
        self._sort_order = None
    
    def configure_table_widget(self, table):
        """Configurazione centralizzata dello stile e del comportamento delle tabelle"""
        table.setColumnCount(len(self.columns))
        table.setHorizontalHeaderLabels(self.columns)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(False)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        
        # Imposta altezza minima righe per evitare tagli
        table.verticalHeader().setDefaultSectionSize(35)
        table.verticalHeader().setMinimumSectionSize(35)
    
    def enable_manual_sorting(self, table):
        """Abilita sorting manuale con indicatori Unicode visibili"""
        header = table.horizontalHeader()
        header.setSortIndicatorShown(False)
        header.setSectionsClickable(True)
        
        def on_header_clicked(logical_index):
            # Determina la direzione dell'ordinamento
            if self._sort_column == logical_index:
                self._sort_order = Qt.SortOrder.DescendingOrder if self._sort_order == Qt.SortOrder.AscendingOrder else Qt.SortOrder.AscendingOrder
            else:
                self._sort_column = logical_index
                self._sort_order = Qt.SortOrder.AscendingOrder
            
            # Ordina i dati
            table.setSortingEnabled(True)
            table.sortItems(logical_index, self._sort_order)
            table.setSortingEnabled(False)
            
            # Aggiorna le intestazioni con frecce Unicode
            for col in range(table.columnCount()):
                original_text = self.columns[col]
                if col == logical_index:
                    arrow = " ▲" if self._sort_order == Qt.SortOrder.AscendingOrder else " ▼"
                    table.horizontalHeaderItem(col).setText(original_text + arrow)
                else:
                    table.horizontalHeaderItem(col).setText(original_text)
        
        header.sectionClicked.connect(on_header_clicked)

    def reset_sorting(self, table):
        """Rimuove gli indicatori di sorting dalla tabella"""
        self._sort_column = None
        self._sort_order = None
        
        # Ripristina le intestazioni originali senza frecce
        for col in range(table.columnCount()):
            if col < len(self.columns):
                table.horizontalHeaderItem(col).setText(self.columns[col])

class CodeComparator(QMainWindow):
    """Applicazione principale per confrontare liste basate su colonne chiave"""
    
    def __init__(self):
        super().__init__()
        QApplication.setStyle('Fusion')
        
        self.setWindowTitle("Code Comparator - Confronto Liste")
        self.resize(1400, 900)

        # ============================================================
        # CONFIGURAZIONE PRESET
        # ============================================================
        self.presets = {
            "PREDEFINITO (ISBN, Titolo, Autore Editore, Anno, Prezzo, Note)": {
                "key_column": "ISBN",
                "columns": ["ISBN", "TITOLO", "AUTORE", "EDITORE", "ANNO", "PREZZO", "NOTE"]
            },
            "STAMPA REGISTRI (Inventario, Sezione, Collocazione, Specificazione, Sequenza, Descrizione ISBD, Legami)": {
                "key_column": "Inventario",
                "columns": ["Inventario", "Sezione", "Collocazione", "Specificazione", 
                           "Sequenza", "Descrizione ISBD", "Legami"]
            }
        }
        
        # --- File di configurazione nella cartella dello script ---
        try:
            base_dir = Path(__file__).resolve().parent
        except NameError:
            # Fallback se __file__ non è definito (es. in alcuni contesti di esecuzione)
            base_dir = Path.cwd()
        self.config_file = str(base_dir / "presets.json")
        
        # --- Carica preset salvati ---
        self.load_presets()
        
        self.current_preset = "PREDEFINITO (ISBN, Titolo, Autore Editore, Anno, Prezzo, Note)"
        self.key_column = self.presets[self.current_preset]["key_column"]
        self.columns = self.presets[self.current_preset]["columns"].copy()
        
        # ============================================================
        # TABLE MANAGER
        # ============================================================
        self.table_manager = TableManager(self.columns)
        
        # ============================================================
        # CENTRALIZZAZIONE: Dizionario unico per gestire le liste
        # ============================================================
        self.lists = {
            1: {
                "df": pd.DataFrame(columns=self.columns),
                "table": None,
                "label": None,
                "duplicates_label": None,  # ✅ AGGIUNGI QUESTO
                "deleted": [],
                "btn_restore": None
            },
            2: {
                "df": pd.DataFrame(columns=self.columns),
                "table": None,
                "label": None,
                "duplicates_label": None,  # ✅ AGGIUNGI QUESTO
                "deleted": [],
                "btn_restore": None
            }
        }
                
        

        # Risultati del confronto
        self.matches_df = None
        self.mismatches_df = None

        self.init_ui()
        self.apply_styles()
    
    
    
    
    # ============================================================
    # PERSISTENZA PRESET
    # ============================================================
    
    def save_presets(self):
        """
        Salva i preset attuali in presets.json nella cartella dello script.
        Crea un backup prima di sovrascrivere.
        """
        try:
            # Backup del file esistente
            if os.path.exists(self.config_file):
                backup_file = self.config_file + ".bak"
                try:
                    shutil.copy2(self.config_file, backup_file)
                except Exception as e:
                    logging.warning(f"Impossibile creare backup: {e}")  # warning invece di error
            
            # Salva i preset
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.presets, f, indent=4, ensure_ascii=False)
            
            logging.info(f"Preset salvati in: {self.config_file}")  # ✅
            
        except PermissionError:
            QMessageBox.warning(
                self,
                "Errore Permessi",
                f"Impossibile scrivere su:\n{self.config_file}\n\n"
                "Verifica i permessi della cartella."
            )
            logging.error(f"PermissionError salvando preset in: {self.config_file}")
        except Exception as e:
            QMessageBox.warning(
                self, 
                "Errore Salvataggio",
                f"Impossibile salvare i preset:\n{str(e)}\n\nI preset saranno persi alla chiusura."
            )
            logging.exception(f"Errore nel salvataggio dei preset: {e}")

    def load_presets(self):
        """
        Carica i preset da presets.json se esiste.
        Valida i dati caricati per evitare corruzione.
        """
        if not os.path.exists(self.config_file):
            logging.info("Nessun file presets.json trovato. Uso preset predefiniti.")  
            return
        
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                loaded_presets = json.load(f)
            
            # Validazione: ogni preset deve avere "key_column" e "columns"
            valid_presets = {}
            for preset_name, preset_data in loaded_presets.items():
                if self._validate_preset(preset_name, preset_data):
                    valid_presets[preset_name] = preset_data
                else:
                    logging.warning(f"Preset '{preset_name}' non valido, ignorato.")  # ✅ (usa warning per problemi)
            
            # Unisci con i preset predefiniti (i custom sovrascrivono i default se hanno lo stesso nome)
            self.presets.update(valid_presets)
            
            logging.info(f"Caricati {len(valid_presets)} preset da: {self.config_file}") 
            
        except json.JSONDecodeError as e:
            QMessageBox.warning(...)
            logging.error(f"Errore parsing JSON: {e}")  # ✅ CORRETTO
            self._restore_backup()
            
        except Exception as e:
            logging.error(f"Errore nel caricamento dei preset: {e}")  # ✅ CORRETTO

    def _validate_preset(self, preset_name, preset_data):
        """
        Valida la struttura di un preset usando la validazione unificata.
        
        Args:
            preset_name: Nome del preset
            preset_data: Dizionario con "key_column" e "columns"
        
        Returns:
            True se valido, False altrimenti
        """
        if not isinstance(preset_data, dict):
            return False
        
        if "key_column" not in preset_data or "columns" not in preset_data:
            return False
        
        # Usa la validazione unificata di SettingsDialog
        is_valid, _ = SettingsDialog.validate_preset_data(
            preset_name,
            preset_data["key_column"],
            preset_data["columns"]
        )
        
        return is_valid

    def _restore_backup(self):
        """Tenta di ripristinare il backup in caso di file corrotto"""
        backup_file = self.config_file + ".bak"
        
        if os.path.exists(backup_file):
            try:
                shutil.copy2(backup_file, self.config_file)
                logging.info(f"Backup ripristinato da: {backup_file}")
                
                # Riprova a caricare
                self.load_presets()
            except Exception as e:
                logging.exception(f"Impossibile ripristinare il backup: {e}")
    
    # ============================================================
    # GESTIONE EVENTI TASTIERA
    # ============================================================
    
    def _get_active_list_id(self):
        """Restituisce l'ID della lista con focus, o None"""
        focused = QApplication.focusWidget()
        for list_id, data in self.lists.items():
            if focused == data["table"] or data["table"].isAncestorOf(focused):
                return list_id
        return None

    def keyPressEvent(self, event):
        """Gestisce CTRL+Z per ripristinare righe eliminate e CTRL+C per copiare"""
        list_id = self._get_active_list_id()
        
        if list_id is None:
            super().keyPressEvent(event)
            return
        
        # CTRL+Z: Ripristina righe eliminate
        if event.matches(QKeySequence.StandardKey.Undo):
            if self.lists[list_id]["deleted"]:
                self.restore_rows(list_id)
            event.accept()
        
        # CTRL+C: Copia righe selezionate
        elif event.matches(QKeySequence.StandardKey.Copy):
            if self.lists[list_id]["table"].selectedItems():
                self.copy_selected_rows(list_id)
            event.accept()
        
        else:
            super().keyPressEvent(event)
    

    # ============================================================
    # INTERFACCIA UTENTE
    # ============================================================

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 10, 15, 15)

        # --- HEADER: TITOLO E IMPOSTAZIONI ---
        header_layout = QHBoxLayout()
        
        title_label = QLabel("📚 Code Comparator - Confronto Liste")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50; padding: 5px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        btn_settings = QPushButton("⚙️ Impostazioni")
        btn_settings.setFixedSize(140, 40)  # Aumentato da 130 a 150
        btn_settings.clicked.connect(self.open_settings)
        btn_settings.setObjectName("btn_settings")
        
        header_layout.addStretch()
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(btn_settings)
        
        main_layout.addLayout(header_layout)

        # --- SEZIONE INPUT: TABELLE LISTE ---
        input_layout = QHBoxLayout()
        
        self.lists[1]["table"] = self.create_table_group(input_layout, "📋 Lista 1 (Worklist)", 1)
        self.lists[2]["table"] = self.create_table_group(input_layout, "📂 Lista 2 (Confronto)", 2)
        
        main_layout.addLayout(input_layout)

        # --- BOTTONI AZIONI PRINCIPALI ---
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(10)

        # Pulsante Inverti Liste
        self.btn_swap = QPushButton("🔄 INVERTI LISTE")
        self.btn_swap.setFixedHeight(60)
        self.btn_swap.clicked.connect(self.swap_lists)
        self.btn_swap.setToolTip("Scambia Lista 1 con Lista 2")

        self.btn_compare = QPushButton("🔍 CONFRONTA LISTE")
        self.btn_compare.setFixedHeight(60)
        self.btn_compare.clicked.connect(self.compare_data)

        self.btn_export_matches = QPushButton("💾 ESPORTA CORRISPONDENZE")
        self.btn_export_matches.setFixedHeight(60)
        self.btn_export_matches.clicked.connect(lambda: self.export_results("matches"))
        self.btn_export_matches.setEnabled(False)

        self.btn_export_mismatches = QPushButton("💾 ESPORTA MANCANTI")
        self.btn_export_mismatches.setFixedHeight(60)
        self.btn_export_mismatches.clicked.connect(lambda: self.export_results("mismatches"))
        self.btn_export_mismatches.setEnabled(False)

        self.btn_clear_results = QPushButton("🗑️ CANCELLA RISULTATI")
        self.btn_clear_results.setFixedHeight(60)
        self.btn_clear_results.clicked.connect(self.clear_results)
        self.btn_clear_results.setEnabled(False)

        buttons_layout.addWidget(self.btn_swap)
        buttons_layout.addWidget(self.btn_compare)
        buttons_layout.addWidget(self.btn_export_matches)
        buttons_layout.addWidget(self.btn_export_mismatches)
        buttons_layout.addWidget(self.btn_clear_results)
        
        main_layout.addLayout(buttons_layout)

        # --- SEZIONE RISULTATI (TABS) ---
        self.tabs_results = QTabWidget()
        
        self.res_matches = self._create_result_table()
        self.res_mismatches = self._create_result_table()
        
        
        
        self.tabs_results.addTab(self.res_matches, "✓ Corrispondenze (0)")
        self.tabs_results.addTab(self.res_mismatches, "✗ Mancanti (0)")
        
        main_layout.addWidget(self.tabs_results)

        # --- BARRA INFORMAZIONI ---
        self.info_label = QLabel("Pronto per il confronto. Incolla i dati da Excel/Google Sheets.")
        self.info_label.setStyleSheet("padding: 8px; background-color: #ecf0f1; border-radius: 5px;")
        main_layout.addWidget(self.info_label)

    def create_table_group(self, layout, title, list_id):
        """
        Crea un gruppo tabella con controlli per una lista specifica.
        CENTRALIZZATO: Usa list_id per identificare la lista.
        """
        container = QWidget()
        v_layout = QVBoxLayout(container)
        v_layout.setSpacing(10)
        
        # Header con titolo e contatore
        header_layout = QHBoxLayout()
        label = QLabel(title)
        label.setStyleSheet("font-weight: bold; font-size: 14px; color: #34495e;")

        count_label = QLabel("(0 righe)")
        count_label.setStyleSheet("color: #7f8c8d; font-size: 12px;")

        # Label duplicati cliccabile
        duplicates_label = QLabel("")
        duplicates_label.setStyleSheet(
            "color: #e74c3c; font-size: 12px; text-decoration: underline;"  # ✅ Rimosso "cursor: pointer"
        )
        duplicates_label.setCursor(Qt.CursorShape.PointingHandCursor)  # ✅ Questo è il modo corretto in Qt
        duplicates_label.mousePressEvent = lambda event: self.show_duplicates_dialog(list_id)

        header_layout.addWidget(label)
        header_layout.addWidget(count_label)
        header_layout.addWidget(duplicates_label)
        header_layout.addStretch()
        
        # Bottoni azione (4 pulsanti) - CENTRALIZZATI con lambda
        btn_layout = QHBoxLayout()
        
        btn_paste = QPushButton("📋 Incolla da Excel")
        btn_paste.clicked.connect(lambda: self.paste_data(list_id))
        btn_paste.setFixedHeight(35)
        
        btn_clear = QPushButton("🗑️ Cancella lista")
        btn_clear.clicked.connect(lambda: self.clear_data(list_id))
        btn_clear.setFixedHeight(35)
        
        btn_delete = QPushButton("❌ Elimina righe")
        btn_delete.clicked.connect(lambda: self.delete_rows(list_id))
        btn_delete.setFixedHeight(35)
        
        btn_restore = QPushButton("↩️ Ripristina righe")
        btn_restore.clicked.connect(lambda: self.restore_rows(list_id))
        btn_restore.setFixedHeight(35)
        btn_restore.setEnabled(False)
        
        btn_layout.addWidget(btn_paste)
        btn_layout.addWidget(btn_clear)
        btn_layout.addWidget(btn_delete)
        btn_layout.addWidget(btn_restore)
        
        # Tabella con menu contestuale
        table = QTableWidget(0, len(self.columns))
        self.table_manager.configure_table_widget(table)
        self.table_manager.enable_manual_sorting(table)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        
        table.customContextMenuRequested.connect(lambda pos: self.show_context_menu(pos, list_id))
        
        v_layout.addLayout(header_layout)
        v_layout.addLayout(btn_layout)
        v_layout.addWidget(table)
        
        layout.addWidget(container)
        
        # Salva riferimenti nel dizionario centralizzato
        self.lists[list_id]["table"] = table
        self.lists[list_id]["label"] = count_label
        self.lists[list_id]["duplicates_label"] = duplicates_label  # *** NUOVO ***
        self.lists[list_id]["btn_restore"] = btn_restore

        return table

    

    def _create_result_table(self):
        """Crea e configura una tabella risultato"""
        table = QTableWidget()
        self.table_manager.configure_table_widget(table)
        self.table_manager.enable_manual_sorting(table)
        return table


    # ============================================================
    # METODI UNIFICATI PER GESTIONE LISTE
    # ============================================================
    
    def _update_counters(self, list_id, force: bool = False):
        """
        Aggiorna contatori righe e duplicati per una lista.
        
        Args:
            list_id: Identificativo della lista (1 o 2)
            force: Se True, aggiorna sempre anche se il conteggio non è cambiato
        """
        store = self.lists[list_id]
        df = store["df"]
        
        # ✅ OTTIMIZZAZIONE: Salta se il DataFrame è vuoto e non forzato
        if df.empty and not force:
            store["label"].setText("(0 righe)")
            store["duplicates_label"].setText("")
            return
        
        # Filtra righe con chiave vuota o None prima di contare duplicati
        df_with_keys = self._get_df_with_valid_keys(df)
        
        if len(df_with_keys) == 0:
            store["duplicates_label"].setText("")
            return
        
        duplicates_unique, duplicates_count = ComparisonEngine.count_duplicates(df_with_keys)

        if duplicates_count == 0:
            store["duplicates_label"].setText("")
            return
        store["duplicates_label"].setText(
            f"⚠️ {duplicates_unique} duplicati ({duplicates_count} righe)"
        )
    

    def paste_data(self, list_id: int) -> None:
        """Metodo UNIFICATO per incollare dati in qualsiasi lista"""
        new_df = self.get_clipboard_data()
        if new_df is None:
            return
        
        store = self.lists[list_id]
        old_row_count = len(store["df"])
        
        # Aggiungi i nuovi dati (gestisci numerazione progressiva)
        if not store["df"].empty:
            # ✅ Continua la numerazione delle righe originali
            last_row = store["df"]['__ORIGINAL_ROW__'].max()
            new_df['__ORIGINAL_ROW__'] = new_df['__ORIGINAL_ROW__'] + last_row
            store["df"] = pd.concat([store["df"], new_df], ignore_index=True)
        else:
            # ✅ Prima volta: usa il DataFrame così com'è
            store["df"] = new_df
        
        self.update_table_display(store["table"], store["df"], highlight_from=old_row_count)
        self._update_counters(list_id)
        self.table_manager.reset_sorting(store["table"])

        new_rows = len(new_df)
        if old_row_count > 0:
            self.info_label.setText(f"✅ Lista {list_id}: aggiunte {new_rows} righe (totale: {len(store['df'])})")
            logging.info(f"Lista {list_id}: incollate {new_rows} righe (totale: {len(store['df'])})")
        else:
            self.info_label.setText(f"✅ Lista {list_id} caricata: {len(store['df'])} righe")
            logging.info(f"Lista {list_id}: caricata con {len(store['df'])} righe")

    def clear_data(self, list_id: int) -> None:
        """
        Cancella tutti i dati di una lista specifica.
        
        Args:
            list_id: Identificativo della lista (1 o 2)
        
        Note:
            Questa operazione resetta anche lo storico delle eliminazioni.
        """
        store = self.lists[list_id]
        store["df"] = pd.DataFrame(columns=self.columns + ['__ORIGINAL_ROW__'])
        store["table"].setRowCount(0)
        store["deleted"].clear()
        store["btn_restore"].setEnabled(False)
        
        self._update_counters(list_id)
        self.info_label.setText(f"Lista {list_id} cancellata")

    def delete_rows(self, list_id: int) -> None:
        """Metodo UNIFICATO per eliminare righe selezionate"""
        store = self.lists[list_id]
        table = store["table"]
        df = store["df"]
        
        selected_rows_ui = sorted(set(item.row() for item in table.selectedItems()))
        
        if not selected_rows_ui:
            QMessageBox.warning(self, "Nessuna Selezione", 
                              "Seleziona almeno una riga da eliminare!")
            return
        
        reply = QMessageBox.question(
            self, 
            "Conferma Eliminazione",
            f"Vuoi eliminare {len(selected_rows_ui)} riga/e selezionata/e?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        indices_to_delete = []
        deleted_rows_data = []
        
        for row_idx in selected_rows_ui:
            # ✅ OTTIMIZZATO: Usa sempre la prima colonna per UserRole (più affidabile)
            item = table.item(row_idx, 0)
            if not item:
                logging.warning(f"Riga {row_idx}: nessun item trovato nella prima colonna")
                continue
            
            real_index = item.data(Qt.ItemDataRole.UserRole)
            if real_index is None:
                logging.warning(f"Riga {row_idx}: UserRole non impostato")
                continue
            
            if real_index is not None and isinstance(real_index, int) and 0 <= real_index < len(df):
                indices_to_delete.append(real_index)
                try:
                    deleted_rows_data.append({
                        'index': real_index,
                        'data': df.iloc[real_index].to_dict()
                    })
                except Exception as e:
                    logging.error(f"Errore nel salvataggio riga {real_index}: {e}")
        
        if not indices_to_delete:
            QMessageBox.warning(self, "Errore", 
                "Impossibile identificare le righe da eliminare!\n"
                "Prova a ricaricare i dati.")
            return
        
        store["deleted"].append(deleted_rows_data)
        df_updated = df.drop(indices_to_delete).reset_index(drop=True)
        store["df"] = df_updated
        
        self.update_table_display(store["table"], df_updated)
        self._update_counters(list_id)
        store["btn_restore"].setEnabled(True)

        logging.info(f"Lista {list_id}: eliminate {len(indices_to_delete)} righe")
        self.info_label.setText(f"✅ Eliminate {len(indices_to_delete)} righe dalla Lista {list_id}")

    
    def restore_rows(self, list_id: int) -> None:
        """Metodo UNIFICATO per ripristinare righe eliminate"""
        store = self.lists[list_id]
        
        if not store["deleted"]:
            QMessageBox.information(self, "Nessuna Operazione", 
                                  "Non ci sono eliminazioni da ripristinare!")
            return
        
        last_deleted = store["deleted"].pop()
        df = store["df"]
        
        try:
            # Ordina per indice decrescente per evitare shift
            sorted_deleted = sorted(last_deleted, key=lambda x: x['index'], reverse=True)
            
            # ✅ VALIDAZIONE: Verifica che le colonne siano ancora compatibili
            if last_deleted:
                sample_row = last_deleted[0]['data']
                missing_cols = set(sample_row.keys()) - set(df.columns)
                if missing_cols:
                    raise ValueError(
                        f"Impossibile ripristinare: colonne mancanti nel DataFrame corrente: {missing_cols}"
                    )
            
            # Reinserisci le righe nelle posizioni originali
            for item in sorted_deleted:
                original_index = item['index']
                row_data = pd.Series(item['data'])
                
                # Inserisci la riga nella posizione originale
                if original_index <= len(df):
                    df = pd.concat([
                        df.iloc[:original_index],
                        pd.DataFrame([row_data]),
                        df.iloc[original_index:]
                    ], ignore_index=True)
                else:
                    df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
            
            store["df"] = df
            
            # Trova le posizioni delle righe ripristinate per evidenziarle
            restored_positions = sorted([item['index'] for item in last_deleted])
            min_position = min(restored_positions) if restored_positions else 0
            
            self.update_table_display(store["table"], store["df"], highlight_from=min_position)
            self._update_counters(list_id)
            self.table_manager.reset_sorting(store["table"])  # AGGIUNGI QUESTA RIGA

            if not store["deleted"]:
                store["btn_restore"].setEnabled(False)

            self.info_label.setText(f"✅ Ripristinate {len(last_deleted)} righe nella Lista {list_id}")
            
        except Exception as e:
            # Ripristina lo stato precedente in caso di errore
            store["deleted"].append(last_deleted)
            logging.exception(f"Errore nel ripristino righe: {e}")
            QMessageBox.critical(
                self,
                "❌ Errore Ripristino",
                f"Impossibile ripristinare le righe:\n{str(e)}\n\n"
                "Le righe sono state mantenute nello storico eliminazioni."
            )
    
    

    def copy_selected_rows(self, list_id: int) -> None:
        """Copia le righe selezionate negli appunti (formato Excel-compatibile)"""
        store = self.lists[list_id]
        table = store["table"]
        
        selected_rows = sorted(set(item.row() for item in table.selectedItems()))
        
        if not selected_rows:
            QMessageBox.warning(self, "Nessuna Selezione", 
                              "Seleziona almeno una riga da copiare!")
            return
        
        copied_data = []
        num_cols = table.columnCount()  # ✅ Usa il numero di colonne EFFETTIVE della tabella
        for row_idx in selected_rows:
            row_data = []
            for col_idx in range(num_cols):  
                item = table.item(row_idx, col_idx)
                cell_value = item.text() if item else ""
                
                # ✅ SOLUZIONE: Racchiudi tra virgolette se contiene newline o tab
                if '\n' in cell_value or '\t' in cell_value or '"' in cell_value:
                    # Escape virgolette doppie (come fa Excel)
                    cell_value = cell_value.replace('"', '""')
                    cell_value = f'"{cell_value}"'
                
                row_data.append(cell_value)
            copied_data.append("\t".join(row_data))
        
        clipboard_text = "\n".join(copied_data)
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text)
        
        self.info_label.setText(f"✓ Copiate {len(selected_rows)} righe dalla Lista {list_id} negli appunti")


    def show_duplicates_dialog(self, list_id):
        """Mostra una finestra con le righe duplicate"""
        store = self.lists[list_id]
        df = store["df"]
        
        # Filtra righe con chiave vuota PRIMA di cercare duplicati
        df_with_keys = self._get_df_with_valid_keys(df)
        
        if df_with_keys.empty:
            QMessageBox.information(self, "Nessun Dato", 
                                   f"La Lista {list_id} non contiene righe con chiave valida!")
            return
        
        # Trova le righe duplicate (ora senza le righe vuote)
        duplicated_mask = df_with_keys['KEY_NORMALIZED'].duplicated(keep=False)
        duplicates_df = df_with_keys[duplicated_mask].copy()
        
        if duplicates_df.empty:
            QMessageBox.information(self, "Nessun Duplicato", 
                                   f"La Lista {list_id} non contiene duplicati!")
            return
            
        # Crea finestra di dialogo
        dialog = QDialog(self)
        dialog.setWindowTitle(f"🔍 Duplicati - Lista {list_id}")
        dialog.resize(1000, 600)
        
        layout = QVBoxLayout(dialog)
        
        # Informazioni
        duplicates_unique = duplicates_df['KEY_NORMALIZED'].nunique()
        duplicates_count = len(duplicates_df)
        
        # Conta quante righe con chiave vuota sono state escluse
        empty_keys_count = len(df) - len(df_with_keys)
        
        info_text = (
            f"<b>Chiavi duplicate:</b> {duplicates_unique}<br>"
            f"<b>Righe totali duplicate:</b> {duplicates_count}"
        )
        
        if empty_keys_count > 0:
            info_text += f"<br><b>⚠️ Righe con chiave vuota (escluse):</b> {empty_keys_count}"
        
        info_label = QLabel(info_text)
        info_label.setStyleSheet("padding: 10px; background-color: #fff3cd; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # Tabella con righe duplicate
        table = QTableWidget()
        table.setColumnCount(len(self.columns) + 1)  # +1 per la colonna "Riga"
        table.setHorizontalHeaderLabels(["Riga Visibile"] + self.columns)
        
        # Configura tabella
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(True)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        
        # ✅ SOLUZIONE: Costruisci una mappa chiave -> posizione visibile nella tabella UI
        table_widget = store["table"]
        key_to_visual_row = {}
        
        for visual_row in range(table_widget.rowCount()):
            # Trova l'item della colonna chiave (con protezione da cambio preset)
            try:
                key_col_index = self.columns.index(self.key_column)
            except ValueError:
                QMessageBox.warning(
                    self, 
                    "⚠️ Errore Configurazione",
                    f"La colonna chiave '{self.key_column}' non esiste più.\n"
                    "Probabilmente hai cambiato preset dopo aver caricato i dati."
                )
                return
            item = table_widget.item(visual_row, key_col_index)
            
            if item:
                key_normalized = self.normalize_key(item.text())
                if key_normalized:
                    # Se la chiave ha duplicati, salva tutte le posizioni
                    if key_normalized not in key_to_visual_row:
                        key_to_visual_row[key_normalized] = []
                    key_to_visual_row[key_normalized].append(visual_row + 1)  # +1 per numerazione umana
        
        # Popola tabella (ordina per chiave normalizzata per raggruppare i duplicati)
        duplicates_sorted = duplicates_df.sort_values('KEY_NORMALIZED').reset_index(drop=False)
        table.setRowCount(len(duplicates_sorted))
        
        for i, row in duplicates_sorted.iterrows():
            key = row['KEY_NORMALIZED']
            
            # ✅ Usa la mappa per trovare le posizioni visibili
            visual_rows = key_to_visual_row.get(key, [])
            
            if visual_rows:
                # Se ci sono più righe con la stessa chiave, mostra tutte
                row_text = ", ".join(map(str, visual_rows))
                row_item = QTableWidgetItem(row_text)
                row_item.setBackground(QColor("#ffeaa7"))
                table.setItem(i, 0, row_item)
            else:
                # Fallback: la chiave non è più presente nella tabella UI
                # (può succedere se i dati sono stati modificati dopo il confronto)
                row_item = QTableWidgetItem("N/D")
                row_item.setBackground(QColor("#ffcccc"))
                row_item.setToolTip("Riga non trovata nella tabella corrente")
                table.setItem(i, 0, row_item)
            
            # Altre colonne
            for j, col_name in enumerate(self.columns):
                value = row[col_name] if col_name in row else ""
                item = QTableWidgetItem(str(value) if not pd.isna(value) else "")
                
                # Evidenzia la colonna chiave
                if col_name == self.key_column:
                    item.setBackground(QColor("#ff7675"))
                    font = QFont()
                    font.setBold(True)
                    item.setFont(font)
                
                table.setItem(i, j + 1, item)
        
        layout.addWidget(table)
        
        # Pulsante chiudi
        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(dialog.accept)
        layout.addWidget(btn_close)
        
        dialog.exec()
        
        
    # ============================================================
    # MENU CONTESTUALE
    # ============================================================

    def show_context_menu(self, pos, list_id):
        """Menu contestuale UNIFICATO"""
        store = self.lists[list_id]
        menu = QMenu(self)
        
        # Azioni del menu
        paste_action = menu.addAction("📋 Incolla da Excel")
        copy_action = menu.addAction("📄 Copia righe selezionate")  # *** AGGIUNTO ***
        delete_action = menu.addAction("❌ Elimina righe selezionate")
        menu.addSeparator()
        restore_action = menu.addAction("↩️ Ripristina ultima eliminazione")
        
        # Disabilita "Copia" ed "Elimina" se non ci sono selezioni
        has_selection = bool(store["table"].selectedItems())
        copy_action.setEnabled(has_selection)  # *** AGGIUNTO ***
        delete_action.setEnabled(has_selection)
        
        # Disabilita "Ripristina" se non ci sono eliminazioni da ripristinare
        restore_action.setEnabled(len(store["deleted"]) > 0)
        
        # Connetti le azioni (CENTRALIZZATO)
        paste_action.triggered.connect(lambda: self.paste_data(list_id))
        copy_action.triggered.connect(lambda: self.copy_selected_rows(list_id))  # *** AGGIUNTO ***
        delete_action.triggered.connect(lambda: self.delete_rows(list_id))
        restore_action.triggered.connect(lambda: self.restore_rows(list_id))
        
        # Mostra il menu nella posizione del cursore
        menu.exec(store["table"].viewport().mapToGlobal(pos))


    def _get_df_with_valid_keys(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filtra il DataFrame escludendo righe con chiavi vuote o None.
        
        Args:
            df: DataFrame con colonna KEY_NORMALIZED
        
        Returns:
            DataFrame filtrato con solo chiavi valide
        
        Example:
            >>> df = pd.DataFrame({'KEY_NORMALIZED': ['ABC123', None, '', 'XYZ789']})
            >>> valid_df = self._get_df_with_valid_keys(df)
            >>> len(valid_df)  # Restituisce 2 (solo ABC123 e XYZ789)
        """
        if df.empty or 'KEY_NORMALIZED' not in df.columns:
            return pd.DataFrame()
        
        return df[
            df['KEY_NORMALIZED'].notna() & 
            (df['KEY_NORMALIZED'].astype(str).str.strip() != "")
        ].copy()


    # ============================================================
    # UTILITÀ
    # ============================================================

    def normalize_key(self, key_value) -> str:
        """
        Normalizza la colonna chiave rimuovendo caratteri non alfanumerici.
        
        Args:
            key_value: Valore da normalizzare
        
        Returns:
            Stringa normalizzata (solo caratteri alfanumerici in maiuscolo)
            Restituisce stringa vuota se il valore è None/NaN o vuoto dopo la normalizzazione
        """
        if pd.isna(key_value):
            return ""
        
        key_str = str(key_value).strip()
        
        # ✅ MIGLIORAMENTO: Gestisci esplicitamente stringhe vuote
        if not key_str:
            return ""
        
        normalized = ''.join(c for c in key_str if c.isalnum()).upper()
        
        # ✅ MIGLIORAMENTO: Log per debug se la normalizzazione rimuove tutto
        if not normalized and key_str:
            logging.debug(f"Chiave '{key_str}' normalizzata in stringa vuota")
        
        return normalized

    
    def _get_display_columns(self) -> list:
        """Restituisce solo le colonne visibili (esclude colonne tecniche)"""
        return [col for col in self.columns if not col.startswith('__')]
        
    def _parse_clipboard_text(self, text):
        """
        Trasforma il testo della clipboard in DataFrame grezzo.
        Gestisce sia celle con virgolette (Excel) sia senza (fallback).
        
        Args:
            text: Testo copiato dalla clipboard
        
        Returns:
            DataFrame grezzo o None in caso di errore
        """
        try:
            # ✅ TENTATIVO 1: Usa csv.reader (gestisce virgolette)
            rows = []
            reader = csv.reader(
                io.StringIO(text), 
                delimiter='\t',
                quotechar='"',
                doublequote=True,
                skipinitialspace=False
            )
            
            for row in reader:
                rows.append(row)
            
            if not rows:
                raise pd.errors.EmptyDataError("Nessuna riga trovata")
            
            # ✅ VERIFICA: Controlla se ci sono righe con meno colonne del previsto
            expected_cols = len(self.columns)
            # ✅ COSTANTE: Soglia minima per considerare una riga come incompleta
            MIN_COLUMNS_THRESHOLD = max(1, expected_cols // 2)  # Almeno metà delle colonne
            processed_rows = []
            
            for row in rows:
                num_cols = len(row)
                
                # Riga completa
                if num_cols >= expected_cols:
                    processed_rows.append(row[:expected_cols])
                
                # Riga incompleta: uniscila alla precedente (fallback per righe senza virgolette)
                elif processed_rows:
                    # Unisci all'ultima cella dell'ultima riga
                    processed_rows[-1][-1] += '\n' + '\t'.join(row)
                else:
                    # Prima riga malformata: ignorala
                    logging.warning(f"Prima riga malformata ignorata: {row}")
            
            if not processed_rows:
                raise pd.errors.EmptyDataError("Nessuna riga valida trovata")
            
            # Converti in DataFrame
            df = pd.DataFrame(processed_rows, dtype=str)
            
            # ✅ Traccia il numero di riga originale
            df['__ORIGINAL_ROW__'] = range(1, len(df) + 1)
            
            # Assegna nomi colonne
            df.columns = self.columns + ['__ORIGINAL_ROW__']
            
            return df
            
        except pd.errors.EmptyDataError:
            QMessageBox.warning(self, "Dati Vuoti", 
                               "I dati copiati sono vuoti o non validi.")
            return None
            
        except Exception as e:
            logging.exception("Errore parsing clipboard")
            QMessageBox.critical(self, "Errore", 
                               f"Errore nel parsing dei dati:\n{str(e)}")
            return None
        
    
    def _handle_header_detection(self, df: pd.DataFrame) -> pd.DataFrame | None:
        """
        Gestisce il rilevamento automatico dell'intestazione.
        
        Args:
            df: DataFrame da validare
        
        Returns:
            DataFrame con intestazione applicata, o None se l'utente annulla
        
        Note:
            Questo metodo:
            1. Controlla se la prima riga sembra un'intestazione
            2. Chiede conferma all'utente
            3. Applica l'intestazione se confermato
        """
        if df.empty:
            return df
        
        first_row_key = df.iloc[0][self.key_column]
        first_row_key_norm = self.normalize_key(first_row_key)
        
        # Determina se la prima riga è un'intestazione
        is_header = (
            not first_row_key_norm or
            first_row_key_norm.isalpha() or
            first_row_key_norm == self.normalize_key(self.key_column)
        )
        
        if not is_header:
            return df  # Non è un'intestazione, continua
        
        # Chiedi conferma all'utente
        reply = QMessageBox.question(
            self,
            "📋 Intestazione Rilevata",
            f"La prima riga sembra essere un'intestazione:\n"
            f"Valore colonna chiave: '{first_row_key}'\n\n"
            f"Vuoi usare la prima riga come intestazione ed eliminarla dai dati?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.No:
            return None  # Utente ha annullato
        
        # Applica intestazione
        return self._apply_header_from_first_row(df)
        
        
    def _apply_header_from_first_row(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Applica l'intestazione dalla prima riga del DataFrame.
        
        Args:
            df: DataFrame con intestazione nella prima riga
        
        Returns:
            DataFrame con intestazione applicata e prima riga rimossa
        
        Note:
            Questo metodo:
            1. Estrae i nomi colonna dalla prima riga
            2. Aggiorna tutte le strutture dati (preset, tabelle, TableManager)
            3. Salva le modifiche su disco
        """
        # Estrai nomi colonne dalla prima riga
        new_columns = [
            str(df.iloc[0][col]).strip() 
            if not pd.isna(df.iloc[0][col]) and str(df.iloc[0][col]).strip()
            else col
            for col in self.columns
        ]
        
        # Rimuovi prima riga e aggiorna colonne
        df = df.iloc[1:].reset_index(drop=True)
        df.columns = new_columns + ['__ORIGINAL_ROW__']
        
        # Aggiorna colonna chiave
        try:
            old_key_idx = list(self.presets[self.current_preset]["columns"]).index(self.key_column)
            self.key_column = new_columns[old_key_idx]
        except (ValueError, IndexError):
            self.key_column = new_columns[0] if new_columns else self.key_column
        
        self.columns = new_columns
        self.table_manager.columns = new_columns
        
                
        # Aggiorna tutte le tabelle
        for list_id in [1, 2]:
            table = self.lists[list_id]["table"]
            table.setColumnCount(len(self.columns))
            table.setHorizontalHeaderLabels(self.columns)
        
        for result_table in [self.res_matches, self.res_mismatches]:
            result_table.setColumnCount(len(self.columns))
            result_table.setHorizontalHeaderLabels(self.columns)
        
        self.info_label.setText(
            f"✅ Intestazione applicata: colonna chiave = '{self.key_column}'"
        )
        
        return df
        
    def _validate_imported_df(self, df: pd.DataFrame) -> pd.DataFrame | None:
        """
        Valida il DataFrame importato e normalizza la colonna chiave.
        Gestisce automaticamente la presenza di intestazioni nella prima riga.
        
        Args:
            df: DataFrame da validare
        
        Returns:
            DataFrame validato con KEY_NORMALIZED, o None se validazione fallisce
        """
        # Verifica che la colonna chiave esista
        if self.key_column not in df.columns:
            QMessageBox.warning(
                self,
                "Colonna Chiave Mancante",
                f"I dati incollati non contengono la colonna chiave '{self.key_column}'.\n"
                f"Verifica di aver selezionato il preset corretto nelle Impostazioni."
            )
            return None
        
        # Gestione intestazione
        df = self._handle_header_detection(df)
        if df is None:
            return None  # Utente ha annullato
        
        # Normalizza colonna chiave per tutte le righe
        df['KEY_NORMALIZED'] = df[self.key_column].apply(self.normalize_key)
        
        # Verifica chiavi vuote
        empty_mask = df['KEY_NORMALIZED'].isna() | (df['KEY_NORMALIZED'] == "")
        if empty_mask.any():
            empty_count = empty_mask.sum()
            reply = QMessageBox.question(
                self,
                "⚠️ Chiavi Vuote Rilevate",
                f"Attenzione: {empty_count} righe hanno la colonna chiave '{self.key_column}' vuota.\n\n"
                f"Queste righe potrebbero causare problemi nel confronto.\n"
                f"Vuoi continuare comunque?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return None
        
        return df
    
    
    
    
    def get_clipboard_data(self) -> pd.DataFrame | None:
        """Legge la clipboard e la trasforma in DataFrame validato"""
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        
        if not text or not text.strip():
            QMessageBox.warning(self, "Clipboard Vuota", 
                              "La clipboard è vuota. Copia prima i dati da Excel.")
            return None
        
        # Parsing separato dalla validazione
        df = self._parse_clipboard_text(text)
        if df is None:
            return None
        
        # Validazione separata dal parsing
        df = self._validate_imported_df(df)
        
        return df

    def update_table_display(self, table_widget: QTableWidget, df: pd.DataFrame, 
                            highlight_from: int | None = None) -> None:
        """Popola la QTableWidget con i dati del DataFrame"""
        display_columns = self._get_display_columns()
        display_df = df[display_columns] if all(col in df.columns for col in display_columns) else df
        
        show_cursor = len(display_df) > 1000
        if show_cursor:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        
        # Salva stato sorting e disabilitalo temporaneamente
        was_sorting_enabled = table_widget.isSortingEnabled()
        table_widget.setSortingEnabled(False)
        table_widget.setUpdatesEnabled(False)
        
        try:
            table_widget.setRowCount(len(display_df))
            
            is_res = table_widget in [self.res_matches, self.res_mismatches]
            bg_color = QColor(MATCH_COLOR) if table_widget == self.res_matches else QColor(MISMATCH_COLOR)
            color_highlight = QColor(HIGHLIGHT_COLOR)
            color_white = QColor(WHITE_COLOR)
            
            # ✅ OTTIMIZZAZIONE: Pre-alloca tutti gli item in una volta sola
            for i in range(len(display_df)):
                for j in range(len(display_columns)):
                    if not table_widget.item(i, j):
                        table_widget.setItem(i, j, QTableWidgetItem())

            # ✅ OTTIMIZZAZIONE: Popola i dati senza ricreare gli item
            for i, row in enumerate(display_df.itertuples(index=True)):
                real_index = row.Index  # ✅ Usa l'indice esplicito del DataFrame
                for j, col_name in enumerate(display_columns):
                    value = row[j + 1] if j + 1 < len(row) else ""
                    
                    item = table_widget.item(i, j)
                    val_str = str(value) if not pd.isna(value) else ""
                    item.setText(val_str)
                    item.setToolTip(val_str if len(val_str) > 30 else "")
                    item.setData(Qt.ItemDataRole.UserRole, real_index)  # ✅ CORRETTO
                    
                    if highlight_from is not None and i >= highlight_from:
                        item.setBackground(color_highlight)
                    elif is_res:
                        item.setBackground(bg_color)
                    else:
                        item.setBackground(color_white)
        
        except Exception as e:
            logging.exception(f"Errore in update_table_display: {e}")
            QMessageBox.critical(
                self,
                "❌ Errore Visualizzazione",
                f"Impossibile visualizzare i dati:\n{str(e)}"
            )
        
        finally:
            table_widget.setUpdatesEnabled(True)
            # Ripristina stato sorting originale
            table_widget.setSortingEnabled(was_sorting_enabled)
            
            if show_cursor:
                QApplication.restoreOverrideCursor()

    # ============================================================
    # CONFRONTO E RISULTATI
    # ============================================================

    def compare_data(self) -> None:
        """Confronta le due liste"""
        if self.lists[1]["df"].empty or self.lists[2]["df"].empty:
            QMessageBox.warning(self, "Dati Mancanti", 
                              "Devi caricare entrambe le liste prima di confrontarle!")
            return

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        self.info_label.setText("⏳ Confronto in corso...")
        QApplication.processEvents()
        
        logging.info(f"Inizio confronto: Lista1={len(self.lists[1]['df'])} righe, Lista2={len(self.lists[2]['df'])} righe")
        
        try:
            matches, mismatches, stats = ComparisonEngine.compare(
                self.lists[1]["df"],
                self.lists[2]["df"]
            )
            
            # Verifica duplicati con conferma utente
            if stats['duplicates_df1'] > 0:
                reply = QMessageBox.question(
                    self, 
                    "⚠️ Duplicati Rilevati",
                    f"Attenzione: ci sono {stats['duplicates_df1']} chiavi duplicate nella Lista 1.\n"
                    f"Questo potrebbe alterare il conteggio dei mancanti.\n\n"
                    f"Vuoi continuare comunque?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    return
            
            if stats['duplicates_df2'] > 0:
                reply = QMessageBox.question(
                    self,
                    "⚠️ Duplicati Rilevati",
                    f"Attenzione: ci sono {stats['duplicates_df2']} chiavi duplicate nella Lista 2.\n"
                    f"Questo potrebbe causare conteggi errati nelle corrispondenze.\n\n"
                    f"Vuoi continuare comunque?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    return

            self.matches_df = matches
            self.mismatches_df = mismatches

            self.update_table_display(self.res_matches, matches)
            self.update_table_display(self.res_mismatches, mismatches)
            
            self.tabs_results.setTabText(0, f"✓ Corrispondenze ({len(matches)})")
            self.tabs_results.setTabText(1, f"✗ Mancanti ({len(mismatches)})")
            
            self.btn_export_matches.setEnabled(len(matches) > 0)
            self.btn_export_mismatches.setEnabled(len(mismatches) > 0)
            self.btn_clear_results.setEnabled(True)
            
            self.info_label.setText(
                f"✓ Confronto completato: {len(matches)} corrispondenze, "
                f"{len(mismatches)} mancanti su {len(self.lists[1]['df'])} totali"
            )
            
            QMessageBox.information(self, "Confronto Completato", 
                f"Risultati:\n\n"
                f"✓ Corrispondenze: {len(matches)}\n"
                f"✗ Mancanti: {len(mismatches)}\n"
                f"📊 Totale Lista 1: {len(self.lists[1]['df'])}"
            )
            
        except KeyError as e:
            logging.exception(f"Errore chiave mancante: {e}")
            QMessageBox.critical(
                self, 
                "❌ Errore Configurazione", 
                f"Colonna mancante nel confronto:\n{str(e)}\n\n"
                f"Verifica che la configurazione corrisponda ai dati caricati."
            )
        except Exception as e:
            logging.exception(f"Errore in compare_data: {e}")
            QMessageBox.critical(
                self, 
                "❌ Errore", 
                f"Errore durante il confronto:\n{str(e)}"
            )
        finally:
            QApplication.restoreOverrideCursor()

    def clear_results(self, confirm: bool = True) -> None:
        """Cancella i risultati del confronto
        
        Args:
            confirm: Se True, chiede conferma all'utente (default: True)
        """
        if confirm:
            reply = QMessageBox.question(
                self, 
                "Conferma Cancellazione",
                "Vuoi cancellare i risultati del confronto?\n\n"
                "Le liste originali (Lista 1 e Lista 2) non verranno modificate.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply != QMessageBox.StandardButton.Yes:
                return
        
        self.res_matches.setRowCount(0)
        self.res_mismatches.setRowCount(0)
        
        self.tabs_results.setTabText(0, "✓ Corrispondenze (0)")
        self.tabs_results.setTabText(1, "✗ Mancanti (0)")
        
        self.btn_export_matches.setEnabled(False)
        self.btn_export_mismatches.setEnabled(False)
        self.btn_clear_results.setEnabled(False)
        
        self.matches_df = None
        self.mismatches_df = None
        
        self.info_label.setText("✓ Risultati cancellati. Pronto per un nuovo confronto.")
            
            
    def swap_lists(self) -> None:
        """Inverte Lista 1 con Lista 2"""
        # Verifica se almeno una lista ha dati
        if self.lists[1]["df"].empty and self.lists[2]["df"].empty:
            QMessageBox.information(self, "Nessun Dato", 
                                  "Entrambe le liste sono vuote. Nulla da invertire!")
            return
        
        # Chiedi conferma
        reply = QMessageBox.question(
            self,
            "🔄 Conferma Inversione",
            "Vuoi invertire Lista 1 con Lista 2?\n\n"
            "• Lista 1 (Worklist) diventerà Lista 2 (Confronto)\n"
            "• Lista 2 (Confronto) diventerà Lista 1 (Worklist)\n\n"
            "I risultati del confronto verranno cancellati.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.No:
            return
        
        # Scambia i DataFrame
        temp_df = self.lists[1]["df"].copy()
        self.lists[1]["df"] = self.lists[2]["df"].copy()
        self.lists[2]["df"] = temp_df
        
        # Scambia lo storico eliminazioni
        temp_deleted = self.lists[1]["deleted"].copy()
        self.lists[1]["deleted"] = self.lists[2]["deleted"].copy()
        self.lists[2]["deleted"] = temp_deleted
        
        # Aggiorna le tabelle
        for list_id in [1, 2]:
            store = self.lists[list_id]
            self.update_table_display(store["table"], store["df"])
            self._update_counters(list_id)
            store["btn_restore"].setEnabled(len(store["deleted"]) > 0)
        
        # Cancella i risultati
        self.clear_results(confirm=False)
        
        self.info_label.setText("✓ Liste invertite con successo")
        
        QMessageBox.information(self, "✓ Inversione Completata",
            "Le liste sono state invertite:\n\n"
            f"• Lista 1: {len(self.lists[1]['df'])} righe\n"
            f"• Lista 2: {len(self.lists[2]['df'])} righe")
    

    # ============================================================
    # ESPORTAZIONE UNIFICATA
    # ============================================================

    def export_results(self, result_type: str) -> None:
        """
        Metodo UNIFICATO per esportare corrispondenze o mancanti.
        
        Args:
            result_type: "matches" o "mismatches"
        """
        df = self.matches_df if result_type == "matches" else self.mismatches_df
        prefix = f"{self.key_column}_Corrispondenze" if result_type == "matches" else f"{self.key_column}_Mancanti"
        sheet_name = "Corrispondenze" if result_type == "matches" else "Mancanti"
        
        if df is None or df.empty:
            QMessageBox.warning(self, "Nessun Dato", 
                              f"Non ci sono {sheet_name.lower()} da esportare!")
            return
        
        desktop = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DesktopLocation)
        file_path = self._get_unique_filepath(desktop, prefix, "xlsx")
        
        try:
            self._export_to_formatted_excel(
                df[self.columns], 
                file_path,
                sheet_name
            )
            
            QMessageBox.information(self, "Esportazione Completata", 
                f"File salvato con successo sul Desktop:\n{os.path.basename(file_path)}")
            self.info_label.setText(f"✓ Esportati {len(df)} {sheet_name.lower()}: {os.path.basename(file_path)}")
            
        except Exception as e:
            logging.error(f"Errore in export_results: {e}")
            QMessageBox.critical(self, "Errore", 
                               f"Errore durante l'esportazione:\n{str(e)}")

    def _get_unique_filepath(self, directory, base_name, extension):
        """
        Genera un percorso file univoco con numerazione progressiva.
        
        Args:
            directory: Cartella di destinazione
            base_name: Nome base del file (senza estensione)
            extension: Estensione del file (senza punto)
        
        Returns:
            Percorso completo del file con numerazione progressiva se necessario
        """
        # Prova prima senza numero
        filepath = os.path.join(directory, f"{base_name}.{extension}")
        
        if not os.path.exists(filepath):
            return filepath
        
        # Se esiste, aggiungi numerazione progressiva
        counter = 1
        while True:
            filepath = os.path.join(directory, f"{base_name}_{counter}.{extension}")
            if not os.path.exists(filepath):
                return filepath
            counter += 1

    @staticmethod
    def _calculate_excel_column_width(column_name):
        """
        Calcola la larghezza ottimale per una colonna Excel basandosi sul nome.
        La ricerca è case-insensitive.
        
        Args:
            column_name: Nome della colonna (es. "ISBN", "Titolo", "autore")
        
        Returns:
            int: Larghezza in unità Excel (default 20)
        """
        col_name_upper = column_name.upper()
        
        for target_width, keywords in EXCEL_COLUMN_WIDTHS.items():
            if any(kw in col_name_upper for kw in keywords):
                return target_width
        
        return 20  # Default


    def _export_to_formatted_excel(self, df: pd.DataFrame, filepath: str, sheet_name: str):
        """
        Esporta DataFrame in Excel con formattazione professionale.
        
        Specifiche di formattazione:
        - Header azzurro (#8DBEE3) con testo in grassetto
        - Larghezze colonne ottimizzate dinamicamente
        - Altezza righe uniforme (30px)
        - Allineamento verticale top, orizzontale left
        - Wrap text abilitato
        - Layout stampa ottimizzato (A4 orizzontale)
        - Freeze panes sulla prima riga
        - Zoom 110%
        
        Args:
            df: DataFrame da esportare
            filepath: Percorso file di output
            sheet_name: Nome del foglio Excel
        """
        # Crea workbook e worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        
        # ====================================================================
        # CONFIGURAZIONE LAYOUT E STAMPA
        # ====================================================================
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.freeze_panes = 'A2'
        ws.sheet_view.zoomScale = 110
        
        # ====================================================================
        # STILI
        # ====================================================================
        header_fill = PatternFill(start_color='8DBEE3', end_color='8DBEE3', fill_type='solid')
        header_font = Font(bold=True)
        alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        
        # ====================================================================
        # SCRITTURA HEADER CON LARGHEZZE DINAMICHE
        # ====================================================================
        for col_idx, column_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=column_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = alignment
            
            # Calcolo larghezza intelligente basato sul nome colonna
            width = self._calculate_excel_column_width(column_name)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
            
        
        # Altezza riga header
        ws.row_dimensions[1].height = 19
        
    
        # ====================================================================
        # SCRITTURA DATI CON ALTEZZA UNIFORME (CICLO UNIFICATO)
        # ====================================================================
        # Calcola il numero totale di righe da formattare (minimo 100 per consistenza visiva)
        total_rows = max(len(df) + 1, 100)

        # Imposta l'altezza per TUTTE le righe (anche quelle vuote)
        for row_idx in range(2, total_rows + 1):
            ws.row_dimensions[row_idx].height = 30

        # Scrivi i dati nelle righe valorizzate
        for row_idx, row_data in enumerate(df.itertuples(index=False), 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value if not pd.isna(value) else ""
                cell.alignment = alignment
        
        # ====================================================================
        # SALVATAGGIO
        # ====================================================================
        wb.save(filepath)

    # ============================================================
    # IMPOSTAZIONI E PRESET
    # ============================================================

    def open_settings(self):
        """Apre la finestra di dialogo delle impostazioni"""
        # Controlla se ci sono dati caricati
        has_data = not self.lists[1]["df"].empty or not self.lists[2]["df"].empty
        has_results = self.matches_df is not None or self.mismatches_df is not None
        
        if has_data or has_results:
            reply = QMessageBox.warning(
                self,
                "⚠️ Attenzione",
                "Cambiare configurazione cancellerà TUTTI i dati caricati e i risultati.\n\n"
                "Vuoi continuare?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply != QMessageBox.StandardButton.Yes:
                return
        
        dialog = SettingsDialog(self.presets, self.current_preset, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            result = dialog.get_result()
            
            if result['action'] == 'select':
                # Carica preset esistente
                self.current_preset = result['preset_name']
                self.key_column = self.presets[self.current_preset]["key_column"]
                self.columns = self.presets[self.current_preset]["columns"].copy()
                
            elif result['action'] == 'create':
                # Crea nuovo preset
                preset_name = result['preset_name']
                self.presets[preset_name] = {
                    "key_column": result['key_column'],
                    "columns": result['columns']
                }
                self.current_preset = preset_name
                self.key_column = result['key_column']
                self.columns = result['columns']
                # --- SALVA AUTOMATICAMENTE SU DISCO ---
                self.save_presets()
            
            # Aggiorna solo le colonne del TableManager esistente
            self.table_manager.columns = self.columns
            
            # Resetta le liste
            self.reset_all_data()
            
            # Aggiorna il titolo della finestra
            self.setWindowTitle(f"Code Comparator - {self.current_preset}")
            
            self.info_label.setText(f"✓ Configurazione cambiata: {self.current_preset} (Colonna chiave: {self.key_column})")

    def reset_all_data(self):
        """Reset completo di tutti i dati quando si cambia configurazione"""
        # Reset dataframe nelle liste
        for list_id in [1, 2]:
            store = self.lists[list_id]
            store["df"] = pd.DataFrame(columns=self.columns)
            store["deleted"].clear()
            
            # Reset tabelle input con riconfigurare completa
            self.table_manager.configure_table_widget(store["table"])
            self.table_manager.enable_manual_sorting(store["table"])
            self.table_manager.reset_sorting(store["table"])
            store["table"].setRowCount(0)
            
            # Reset contatori
            store["label"].setText("(0 righe)")
            
            # Reset pulsanti ripristina
            store["btn_restore"].setEnabled(False)
        
        # Reset risultati
        self.table_manager.reset_sorting(self.res_matches)
        self.res_matches.setRowCount(0)
        self.table_manager.reset_sorting(self.res_mismatches)
        self.res_mismatches.setRowCount(0)
        
        # Reset tab
        self.tabs_results.setTabText(0, "✓ Corrispondenze (0)")
        self.tabs_results.setTabText(1, "✗ Mancanti (0)")
        
        # Disabilita pulsanti export
        self.btn_export_matches.setEnabled(False)
        self.btn_export_mismatches.setEnabled(False)
        self.btn_clear_results.setEnabled(False)
        
        # Rimuovi dataframe risultati
        self.matches_df = None
        self.mismatches_df = None

    

    def apply_styles(self):
        """Applica gli stili CSS all'applicazione"""
        self.setStyleSheet(APP_STYLESHEET)
        
        # Imposta gli object names per i pulsanti
        self.btn_swap.setObjectName("btn_swap")
        self.btn_compare.setObjectName("btn_compare")
        self.btn_export_matches.setObjectName("btn_export_matches")
        self.btn_export_mismatches.setObjectName("btn_export_mismatches")
        self.btn_clear_results.setObjectName("btn_clear_results")

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CodeComparator()
    window.show()
    sys.exit(app.exec())
