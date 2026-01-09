"""
Finestra principale dell'applicazione - PARTE 1: Inizializzazione e Gestione Dati
"""
import sys
import os
import logging
import io
import csv
import pandas as pd

from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QTableWidget, QTableWidgetItem, QPushButton, QLabel, 
                             QTabWidget, QMessageBox, QMenu, QDialog, QApplication)
from PyQt6.QtCore import Qt, QStandardPaths
from PyQt6.QtGui import QColor, QKeySequence, QFont

from config import (MAX_COLUMNS, DUPLICATE_KEY_COLOR, DUPLICATE_HIGHLIGHT)
from models import ComparisonEngine, PresetManager
from utils import normalize_key, export_to_formatted_excel, get_unique_filepath
from .table_manager import TableManager
from .settings_dialog import SettingsDialog
from .styles import APP_STYLESHEET


class CodeComparator(QMainWindow):
    """Applicazione principale per confrontare liste basate su colonne chiave"""
    
    @property
    def columns(self):
        """Restituisce le colonne correnti"""
        return self._columns
    
    @columns.setter
    def columns(self, nuovo_valore):
        """Aggiorna colonne e sincronizza con table_manager"""
        if not isinstance(nuovo_valore, list):
            raise ValueError("columns deve essere una lista")
        
        self._columns = nuovo_valore
        
        if hasattr(self, 'table_manager'):
            self.table_manager.columns = nuovo_valore
    
    def __init__(self):
        super().__init__()
        QApplication.setStyle('Fusion')
        
        self.setWindowTitle("Code Comparator - Confronto Liste")
        self.resize(1200, 900)
        self.setMinimumSize(800, 600)

        # Inizializza PresetManager
        self.preset_manager = PresetManager()
        
        # Usa il preset predefinito se esiste
        if self.preset_manager.default_preset and self.preset_manager.default_preset in self.preset_manager.presets:
            self.current_preset = self.preset_manager.default_preset
        else:
            self.current_preset = list(self.preset_manager.presets.keys())[0]
        
        self.key_column = self.preset_manager.presets[self.current_preset]["key_column"]
        self._columns = self.preset_manager.presets[self.current_preset]["columns"].copy()
        
        # Inizializza TableManager
        self.table_manager = TableManager(self._columns)
        
        # Dizionario unico per gestire le liste
        self.lists = {
            1: {
                "df": pd.DataFrame(columns=self.columns),
                "table": None,
                "label": None,
                "duplicates_label": None,
                "deleted": [],
                "btn_restore": None
            },
            2: {
                "df": pd.DataFrame(columns=self.columns),
                "table": None,
                "label": None,
                "duplicates_label": None,
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
        
        if event.matches(QKeySequence.StandardKey.Undo):
            if self.lists[list_id]["deleted"]:
                self.restore_rows(list_id)
            event.accept()
        
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
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 10, 15, 15)

        # HEADER
        header_layout = QHBoxLayout()
        
        title_label = QLabel("📚 Code Comparator - Confronto Liste")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50; padding: 5px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        btn_settings = QPushButton("⚙️ Impostazioni")
        btn_settings.setMinimumSize(140, 40)
        btn_settings.clicked.connect(self.open_settings)
        btn_settings.setObjectName("btn_settings")
        
        header_layout.addStretch()
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(btn_settings)
        
        main_layout.addLayout(header_layout)

        # SEZIONE INPUT: TABELLE LISTE
        input_layout = QHBoxLayout()
        
        self.lists[1]["table"] = self.create_table_group(input_layout, "📋 Lista 1 (Worklist)", 1)
        self.lists[2]["table"] = self.create_table_group(input_layout, "📂 Lista 2 (Confronto)", 2)
        
        main_layout.addLayout(input_layout)

        # BOTTONI AZIONI PRINCIPALI
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(10)

        self.btn_swap = QPushButton("🔄 INVERTI LISTE")
        self.btn_swap.setFixedHeight(60)
        self.btn_swap.clicked.connect(self.swap_lists)
        self.btn_swap.setToolTip("Scambia Lista 1 con Lista 2")
        self.btn_swap.setObjectName("btn_swap")

        self.btn_compare = QPushButton("🔍 CONFRONTA LISTE")
        self.btn_compare.setFixedHeight(60)
        self.btn_compare.clicked.connect(self.compare_data)
        self.btn_compare.setObjectName("btn_compare")

        self.btn_export_matches = QPushButton("💾 ESPORTA CORRISPONDENZE")
        self.btn_export_matches.setFixedHeight(60)
        self.btn_export_matches.clicked.connect(lambda: self.export_results("matches"))
        self.btn_export_matches.setEnabled(False)
        self.btn_export_matches.setObjectName("btn_export_matches")

        self.btn_export_mismatches = QPushButton("💾 ESPORTA MANCANTI")
        self.btn_export_mismatches.setFixedHeight(60)
        self.btn_export_mismatches.clicked.connect(lambda: self.export_results("mismatches"))
        self.btn_export_mismatches.setEnabled(False)
        self.btn_export_mismatches.setObjectName("btn_export_mismatches")

        self.btn_clear_results = QPushButton("🗑️ CANCELLA RISULTATI")
        self.btn_clear_results.setFixedHeight(60)
        self.btn_clear_results.clicked.connect(self.clear_results)
        self.btn_clear_results.setEnabled(False)
        self.btn_clear_results.setObjectName("btn_clear_results")

        buttons_layout.addWidget(self.btn_swap)
        buttons_layout.addWidget(self.btn_compare)
        buttons_layout.addWidget(self.btn_export_matches)
        buttons_layout.addWidget(self.btn_export_mismatches)
        buttons_layout.addWidget(self.btn_clear_results)
        
        main_layout.addLayout(buttons_layout)

        # SEZIONE RISULTATI (TABS)
        self.tabs_results = QTabWidget()
        
        self.res_matches = self._create_result_table()
        self.res_mismatches = self._create_result_table()
        
        self.tabs_results.addTab(self.res_matches, "✓ Corrispondenze (0)")
        self.tabs_results.addTab(self.res_mismatches, "✗ Mancanti (0)")
        
        main_layout.addWidget(self.tabs_results)

        # BARRA INFORMAZIONI
        self.info_label = QLabel("Pronto per il confronto. Incolla i dati da Excel/Google Sheets.")
        self.info_label.setStyleSheet("padding: 8px; background-color: #ecf0f1; border-radius: 5px;")
        main_layout.addWidget(self.info_label)

    def create_table_group(self, layout, title, list_id):
        """Crea un gruppo tabella con controlli per una lista specifica"""
        container = QWidget()
        v_layout = QVBoxLayout(container)
        v_layout.setSpacing(10)
        
        # Header con titolo e contatore
        header_layout = QHBoxLayout()
        label = QLabel(title)
        label.setStyleSheet("font-weight: bold; font-size: 14px; color: #34495e;")

        count_label = QLabel("(0 righe)")
        count_label.setStyleSheet("color: #7f8c8d; font-size: 12px;")

        duplicates_label = QLabel("")
        duplicates_label.setStyleSheet(
            "color: #e74c3c; font-size: 12px; text-decoration: underline;"
        )
        duplicates_label.setCursor(Qt.CursorShape.PointingHandCursor)
        duplicates_label.mousePressEvent = lambda event: self.show_duplicates_dialog(list_id)

        header_layout.addWidget(label)
        header_layout.addWidget(count_label)
        header_layout.addWidget(duplicates_label)
        header_layout.addStretch()
        
        # Bottoni azione
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
        
        # Salva riferimenti
        store = self.lists[list_id]
        store["table"] = table
        store["label"] = count_label
        store["duplicates_label"] = duplicates_label
        store["btn_restore"] = btn_restore

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
    
    def _update_counters(self, list_id):
        """Aggiorna contatori righe e duplicati per una lista"""
        store = self.lists[list_id]
        df = store["df"]
        
        store["label"].setText(f"({len(df)} righe)")
        
        if df.empty:
            store["duplicates_label"].setText("")
            return
        
        df_with_keys = self._get_df_with_valid_keys(df)
        
        if df_with_keys.empty:
            store["duplicates_label"].setText("")
            return
        
        duplicates_unique, duplicates_count = ComparisonEngine.count_duplicates(df_with_keys)
        
        if duplicates_count > 0:
            store["duplicates_label"].setText(
                f"⚠️ {duplicates_unique} duplicati ({duplicates_count} righe)"
            )
        else:
            store["duplicates_label"].setText("")
    
    def paste_data(self, list_id: int) -> None:
        """Metodo UNIFICATO per incollare dati in qualsiasi lista"""
        new_df = self.get_clipboard_data()
        if new_df is None:
            return
        
        store = self.lists[list_id]
        old_row_count = len(store["df"])
        
        if not store["df"].empty:
            last_row = store["df"]['__ORIGINAL_ROW__'].max()
            new_df['__ORIGINAL_ROW__'] = new_df['__ORIGINAL_ROW__'] + last_row
            store["df"] = pd.concat([store["df"], new_df], ignore_index=True)
        else:
            store["df"] = new_df
        
        self.table_manager.update_table_display(store["table"], store["df"], highlight_from=old_row_count)
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
        """Cancella tutti i dati di una lista specifica"""
        self._reset_single_list(list_id)
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
            item = None
            for col in range(table.columnCount()):
                item = table.item(row_idx, col)
                if item:
                    break
            
            if not item:
                logging.warning(f"Riga {row_idx}: nessun item trovato in alcuna colonna")
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
        
        self.table_manager.update_table_display(store["table"], df_updated)
        self._update_counters(list_id)
        self.table_manager.reset_sorting(store["table"]) 
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
            sorted_deleted = sorted(last_deleted, key=lambda x: x['index'], reverse=True)
            
            if last_deleted:
                sample_row = last_deleted[0]['data']
                missing_cols = set(sample_row.keys()) - set(df.columns)
                if missing_cols:
                    raise ValueError(
                        f"Impossibile ripristinare: colonne mancanti nel DataFrame corrente: {missing_cols}"
                    )
            
            for item in sorted_deleted:
                original_index = item['index']
                row_data = pd.Series(item['data'])
                
                if original_index <= len(df):
                    df = pd.concat([
                        df.iloc[:original_index],
                        pd.DataFrame([row_data]),
                        df.iloc[original_index:]
                    ], ignore_index=True)
                else:
                    df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
            
            store["df"] = df
            
            restored_positions = sorted([item['index'] for item in last_deleted])
            min_position = min(restored_positions) if restored_positions else 0
            
            self.table_manager.update_table_display(store["table"], store["df"], highlight_from=min_position)
            self._update_counters(list_id)
            self.table_manager.reset_sorting(store["table"])

            if not store["deleted"]:
                store["btn_restore"].setEnabled(False)

            self.info_label.setText(f"✅ Ripristinate {len(last_deleted)} righe nella Lista {list_id}")
            
        except Exception as e:
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
        num_cols = table.columnCount()
        for row_idx in selected_rows:
            row_data = []
            for col_idx in range(num_cols):  
                item = table.item(row_idx, col_idx)
                cell_value = item.text() if item else ""
                
                if '\n' in cell_value or '\t' in cell_value or '"' in cell_value:
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
        
        df_with_keys = self._get_df_with_valid_keys(df)
        
        if df_with_keys.empty:
            QMessageBox.information(self, "Nessun Dato", 
                                   f"La Lista {list_id} non contiene righe con chiave valida!")
            return
        
        duplicated_mask = df_with_keys['KEY_NORMALIZED'].duplicated(keep=False)
        duplicates_df = df_with_keys[duplicated_mask].copy()
        
        if duplicates_df.empty:
            QMessageBox.information(self, "Nessun Duplicato", 
                                   f"La Lista {list_id} non contiene duplicati!")
            return
            
        dialog = QDialog(self)
        dialog.setWindowTitle(f"🔍 Duplicati - Lista {list_id}")
        dialog.resize(1000, 600)
        
        layout = QVBoxLayout(dialog)
        
        duplicates_unique = duplicates_df['KEY_NORMALIZED'].nunique()
        duplicates_count = len(duplicates_df)
        
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
        
        table = QTableWidget()
        table.setColumnCount(len(self.columns) + 1)
        table.setHorizontalHeaderLabels(["Riga Visibile"] + self.columns)
        
        from PyQt6.QtWidgets import QHeaderView
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(True)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        
        table_widget = store["table"]
        key_to_visual_row = {}
        
        for visual_row in range(table_widget.rowCount()):
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
                key_normalized = normalize_key(item.text())
                if key_normalized:
                    if key_normalized not in key_to_visual_row:
                        key_to_visual_row[key_normalized] = []
                    key_to_visual_row[key_normalized].append(visual_row + 1)
        
        duplicates_sorted = duplicates_df.sort_values('KEY_NORMALIZED').reset_index(drop=False)
        table.setRowCount(len(duplicates_sorted))
        
        for i, row in duplicates_sorted.iterrows():
            key = row['KEY_NORMALIZED']
            
            visual_rows = key_to_visual_row.get(key, [])
            
            if visual_rows:
                row_text = ", ".join(map(str, visual_rows))
                row_item = QTableWidgetItem(row_text)
                row_item.setBackground(QColor(DUPLICATE_HIGHLIGHT))
                table.setItem(i, 0, row_item)
            else:
                row_item = QTableWidgetItem("N/D")
                row_item.setBackground(QColor("#ffcccc"))
                row_item.setToolTip("Riga non trovata nella tabella corrente")
                table.setItem(i, 0, row_item)
            
            for j, col_name in enumerate(self.columns):
                value = row[col_name] if col_name in row else ""
                item = QTableWidgetItem(str(value) if not pd.isna(value) else "")
                
                if col_name == self.key_column:
                    item.setBackground(QColor(DUPLICATE_KEY_COLOR))
                    font = QFont()
                    font.setBold(True)
                    item.setFont(font)
                
                table.setItem(i, j + 1, item)
        
        layout.addWidget(table)
        
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
        
        paste_action = menu.addAction("📋 Incolla da Excel")
        copy_action = menu.addAction("📄 Copia righe selezionate")
        delete_action = menu.addAction("❌ Elimina righe selezionate")
        menu.addSeparator()
        restore_action = menu.addAction("↩️ Ripristina ultima eliminazione")
        
        has_selection = bool(store["table"].selectedItems())
        copy_action.setEnabled(has_selection)
        delete_action.setEnabled(has_selection)
        
        restore_action.setEnabled(len(store["deleted"]) > 0)
        
        paste_action.triggered.connect(lambda: self.paste_data(list_id))
        copy_action.triggered.connect(lambda: self.copy_selected_rows(list_id))
        delete_action.triggered.connect(lambda: self.delete_rows(list_id))
        restore_action.triggered.connect(lambda: self.restore_rows(list_id))
        
        menu.exec(store["table"].viewport().mapToGlobal(pos))

    def _get_df_with_valid_keys(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filtra il DataFrame escludendo righe con chiavi vuote o None"""
        if df.empty or 'KEY_NORMALIZED' not in df.columns:
            return pd.DataFrame()
        
        return df[
            df['KEY_NORMALIZED'].notna() & 
            (df['KEY_NORMALIZED'].astype(str).str.strip() != "")
        ].copy()
        

    # ============================================================
    # UTILITÀ E PARSING
    # ============================================================
        
    def _get_display_columns(self) -> list:
        """Restituisce solo le colonne visibili (esclude colonne tecniche)"""
        return [col for col in self.columns if not col.startswith('__')]
        
    def _parse_clipboard_text(self, text):
        """Trasforma il testo della clipboard in DataFrame grezzo"""
        try:
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
            
            max_cols = max(len(row) for row in rows)
            
            processed_rows = []
            for row in rows:
                normalized_row = row + [''] * (max_cols - len(row))
                processed_rows.append(normalized_row)
            
            if not processed_rows:
                raise pd.errors.EmptyDataError("Nessuna riga valida trovata")
            
            df = pd.DataFrame(processed_rows, dtype=str)
            
            df['__ORIGINAL_ROW__'] = range(1, len(df) + 1)
            
            if max_cols == len(self.columns):
                column_names = self.columns + ['__ORIGINAL_ROW__']
            elif max_cols < len(self.columns):
                column_names = self.columns[:max_cols] + ['__ORIGINAL_ROW__']
            else:
                extra_cols = [f"Colonna_{i+1}" for i in range(len(self.columns), max_cols)]
                column_names = self.columns + extra_cols + ['__ORIGINAL_ROW__']
            
            df.columns = column_names
            
            return df
            
        except pd.errors.EmptyDataError:
            QMessageBox.warning(self, "Dati Vuoti", 
                               "I dati copiati sono vuoti o non validi.")
            return None
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "❌ Errore Critico - Parsing",
                f"Errore nel parsing dei dati:\n\n{str(e)}"
            )
            logging.exception(f"Errore parsing clipboard: {e}")
            return None
    
    def _handle_header_detection(self, df: pd.DataFrame) -> pd.DataFrame | None:
        """Gestisce il rilevamento automatico dell'intestazione"""
        if df.empty:
            return df
        
        first_row_key = df.iloc[0][self.key_column]
        first_row_key_norm = normalize_key(first_row_key)
        
        is_header = (
            not first_row_key_norm or
            first_row_key_norm.isalpha() or
            first_row_key_norm == normalize_key(self.key_column)
        )
        
        if not is_header:
            return df
        
        reply = QMessageBox.question(
            self,
            "📋 Intestazione Rilevata",
            f"La prima riga sembra essere un'intestazione:\n"
            f"Valore colonna chiave: '{first_row_key}'\n\n"
            f"Vuoi usare la prima riga come intestazione ed eliminarla dai dati?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.No:
            return None
        
        return self._apply_header_from_first_row(df)
        
    def _apply_header_from_first_row(self, df: pd.DataFrame) -> pd.DataFrame:
        """Applica l'intestazione dalla prima riga del DataFrame"""
        data_columns = [col for col in df.columns if col != '__ORIGINAL_ROW__']
        
        new_columns = []
        for i, col_name in enumerate(data_columns):
            value = df.iloc[0][col_name]
            if not pd.isna(value) and str(value).strip():
                new_columns.append(str(value).strip())
            else:
                new_columns.append(col_name)
        
        df = df.iloc[1:].reset_index(drop=True)
        df.columns = new_columns + ['__ORIGINAL_ROW__']
        
        try:
            old_key_idx = list(self.preset_manager.presets[self.current_preset]["columns"]).index(self.key_column)
            if old_key_idx < len(new_columns):
                self.key_column = new_columns[old_key_idx]
            else:
                self.key_column = new_columns[0] if new_columns else self.key_column
        except (ValueError, IndexError):
            self.key_column = new_columns[0] if new_columns else self.key_column
        
        self.columns = new_columns
        self.table_manager.columns = new_columns
        
        for list_id in [1, 2]:
            table = self.lists[list_id]["table"]
            table.setColumnCount(len(self.columns))
            table.setHorizontalHeaderLabels(self.columns)
            self.table_manager.enable_manual_sorting(table)
        
        for result_table in [self.res_matches, self.res_mismatches]:
            result_table.setColumnCount(len(self.columns))
            result_table.setHorizontalHeaderLabels(self.columns)
            self.table_manager.enable_manual_sorting(result_table)
        
        if self.matches_df is not None or self.mismatches_df is not None:
            self.clear_results(confirm=False)
        
        self.info_label.setText(
            f"✅ Intestazione applicata: colonna chiave = '{self.key_column}'"
        )
        
        return df
        
    def _validate_imported_df(self, df: pd.DataFrame) -> pd.DataFrame | None:
        """Valida il DataFrame importato e normalizza la colonna chiave"""
        if self.key_column not in df.columns:
            QMessageBox.warning(
                self,
                "Colonna Chiave Mancante",
                f"I dati incollati non contengono la colonna chiave '{self.key_column}'.\n"
                f"Verifica di aver selezionato il preset corretto nelle Impostazioni."
            )
            return None
        
        df = self._handle_header_detection(df)
        if df is None:
            return None
        
        df['KEY_NORMALIZED'] = df[self.key_column].apply(normalize_key)
        
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
        
        df = self._parse_clipboard_text(text)
        if df is None:
            return None
        
        df = self._validate_imported_df(df)
        
        return df

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

            self.table_manager.update_table_display(self.res_matches, matches, is_match_result=True)
            self.table_manager.update_table_display(self.res_mismatches, mismatches, is_mismatch_result=True)
            
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
            QMessageBox.critical(
                self, 
                "⌠Errore Critico - Configurazione", 
                f"Colonna mancante nel confronto:\n\n{str(e)}\n\n"
                f"Verifica che la configurazione corrisponda ai dati caricati."
            )
            logging.exception(f"Errore chiave mancante: {e}")
        except Exception as e:
            QMessageBox.critical(
                self, 
                "⌠Errore Critico - Confronto", 
                f"Errore durante il confronto:\n\n{str(e)}"
            )
            logging.exception(f"Errore in compare_data: {e}")
        finally:
            QApplication.restoreOverrideCursor()

    def clear_results(self, confirm: bool = True) -> None:
        """Cancella i risultati del confronto"""
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
        if self.lists[1]["df"].empty and self.lists[2]["df"].empty:
            QMessageBox.information(self, "Nessun Dato", 
                                  "Entrambe le liste sono vuote. Nulla da invertire!")
            return
        
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
        
        temp_df = self.lists[1]["df"].copy()
        self.lists[1]["df"] = self.lists[2]["df"].copy()
        self.lists[2]["df"] = temp_df
        
        temp_deleted = self.lists[1]["deleted"].copy()
        self.lists[1]["deleted"] = self.lists[2]["deleted"].copy()
        self.lists[2]["deleted"] = temp_deleted
        
        for list_id in [1, 2]:
            store = self.lists[list_id]
            self.table_manager.update_table_display(store["table"], store["df"])
            self._update_counters(list_id)
            store["btn_restore"].setEnabled(len(store["deleted"]) > 0)
        
        self.clear_results(confirm=False)
        
        self.info_label.setText("✓ Liste invertite con successo")
        
        QMessageBox.information(self, "✓ Inversione Completata",
            "Le liste sono state invertite:\n\n"
            f"• Lista 1: {len(self.lists[1]['df'])} righe\n"
            f"• Lista 2: {len(self.lists[2]['df'])} righe")

    # ============================================================
    # ESPORTAZIONE
    # ============================================================

    def export_results(self, result_type: str) -> None:
        """Esporta corrispondenze o mancanti"""
        df = self.matches_df if result_type == "matches" else self.mismatches_df
        prefix = f"{self.key_column}_Corrispondenze" if result_type == "matches" else f"{self.key_column}_Mancanti"
        sheet_name = "Corrispondenze" if result_type == "matches" else "Mancanti"
        
        if df is None or df.empty:
            QMessageBox.warning(self, "Nessun Dato", 
                              f"Non ci sono {sheet_name.lower()} da esportare!")
            return
        
        desktop = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DesktopLocation)
        file_path = get_unique_filepath(desktop, prefix, "xlsx")
        
        try:
            export_to_formatted_excel(
                df[self.columns], 
                file_path,
                sheet_name
            )
            
            QMessageBox.information(self, "Esportazione Completata", 
                f"File salvato con successo sul Desktop:\n{os.path.basename(file_path)}")
            self.info_label.setText(f"✓ Esportati {len(df)} {sheet_name.lower()}: {os.path.basename(file_path)}")
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "⌠Errore Critico - Esportazione",
                f"Errore durante l'esportazione:\n\n{str(e)}"
            )
            logging.exception(f"Errore in export_results: {e}")

    # ============================================================
    # IMPOSTAZIONI E PRESET
    # ============================================================

    def open_settings(self):
        """Apre la finestra di dialogo delle impostazioni"""
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
        
        dialog = SettingsDialog(
            self.preset_manager.presets, 
            self.current_preset, 
            self.preset_manager.default_preset, 
            self.preset_manager.protected_presets, 
            self
        )
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            result = dialog.get_result()
            
            if result['action'] == 'select':
                self.current_preset = result['preset_name']
                self.key_column = self.preset_manager.presets[self.current_preset]["key_column"]
                self.columns = self.preset_manager.presets[self.current_preset]["columns"].copy()
                self.table_manager.columns = self.columns
                
            elif result['action'] == 'create':
                preset_name = result['preset_name']
                self.preset_manager.presets[preset_name] = {
                    "key_column": result['key_column'],
                    "columns": result['columns']
                }
                self.current_preset = preset_name
                self.key_column = result['key_column']
                self.columns = result['columns']
                self.table_manager.columns = self.columns
                
                if result.get('set_as_default', False):
                    self.preset_manager.default_preset = preset_name
                
                self.preset_manager.save_presets()
            
            elif result['action'] == 'set_default':
                self.preset_manager.default_preset = result['preset_name']
                self.preset_manager.save_presets()
                QMessageBox.information(
                    self,
                    "✓ Preset Predefinito Impostato",
                    f"Il preset '{result['preset_name']}' è ora il predefinito.\n\n"
                    "Verrà caricato automaticamente all'avvio dell'applicazione."
                )
                return
            
            self.preset_manager.presets = dialog.presets.copy()
            self.preset_manager.default_preset = dialog.default_preset
            
            self.reset_all_data()
            
            self.setWindowTitle(f"Code Comparator - {self.current_preset}")
            
            self.info_label.setText(f"✓ Configurazione cambiata: {self.current_preset} (Colonna chiave: {self.key_column})")

    def _reset_table_widget(self, table_widget):
        """Resetta una tabella"""
        table_widget.setColumnCount(len(self.columns))
        table_widget.setHorizontalHeaderLabels(self.columns)
        self.table_manager.reset_sorting(table_widget)
        table_widget.setRowCount(0)
    
    def _reset_single_list(self, list_id: int) -> None:
        """Reset di una singola lista"""
        store = self.lists[list_id]
        store["df"] = pd.DataFrame(columns=self.columns)
        store["deleted"].clear()
        
        self._reset_table_widget(store["table"])
        
        store["label"].setText("(0 righe)")
        store["duplicates_label"].setText("")
        
        store["btn_restore"].setEnabled(False)

    def reset_all_data(self):
        """Reset completo di tutti i dati quando si cambia configurazione"""
        for list_id in [1, 2]:
            self._reset_single_list(list_id)
        
        for table in [self.res_matches, self.res_mismatches]:
            self._reset_table_widget(table)
        
        self.tabs_results.setTabText(0, "✅ Corrispondenze (0)")
        self.tabs_results.setTabText(1, "✗ Mancanti (0)")
            
        self.btn_export_matches.setEnabled(False)
        self.btn_export_mismatches.setEnabled(False)
        self.btn_clear_results.setEnabled(False)
        
        self.matches_df = None
        self.mismatches_df = None

    def apply_styles(self):
        """Applica gli stili CSS all'applicazione"""
        self.setStyleSheet(APP_STYLESHEET)