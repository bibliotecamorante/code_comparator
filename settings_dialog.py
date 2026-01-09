"""
Dialog per la gestione dei preset di configurazione
"""
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QLineEdit, QFormLayout, QScrollArea,
                             QWidget, QListWidget, QListWidgetItem, 
                             QRadioButton, QButtonGroup, QMessageBox)
from PyQt6.QtCore import Qt

from config import MAX_COLUMNS
from models import PresetManager


class SettingsDialog(QDialog):
    """Interfaccia Master-Detail con Radio Button per colonna chiave"""
    
    def __init__(self, presets, current_preset, default_preset, protected_presets, parent=None):
        super().__init__(parent)
        self.presets = presets.copy()
        self.current_selected_name = current_preset
        self.default_preset = default_preset
        self.protected_presets = protected_presets
        self.result = None
        self.is_new_preset_mode = False
        
        self.setWindowTitle("⚙️ Gestione Preset di Configurazione")
        self.resize(950, 650)
        self.setMinimumSize(900, 600)
        
        self.init_ui()
        self.load_preset_list()
        self.select_preset_in_list(current_preset)

    @staticmethod
    def validate_preset_data(preset_name, key_column, columns):
        """Validazione unificata"""
        return PresetManager.validate_preset_data(preset_name, key_column, columns)

    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # LATO SINISTRO: LISTA
        left_widget = QWidget()
        left_widget.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa; 
                border-right: 2px solid #dee2e6;
            }
        """)
        left_widget.setMinimumWidth(280)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(15, 15, 15, 15)
        left_layout.setSpacing(10)
        
        title_list = QLabel("📋 I TUOI PRESET")
        title_list.setStyleSheet("""
            font-weight: bold; 
            font-size: 13px;
            color: #2c3e50; 
            padding: 8px;
            background-color: #e9ecef;
            border-radius: 4px;
        """)
        left_layout.addWidget(title_list)

        self.label_preset_count = QLabel()
        self.label_preset_count.setStyleSheet("color: #6c757d; font-size: 11px; padding: 2px 8px;")
        left_layout.addWidget(self.label_preset_count)

        self.list_presets = QListWidget()
        self.list_presets.setStyleSheet("""
            QListWidget {
                border: 2px solid #ced4da;
                border-radius: 5px;
                background-color: white;
                font-size: 12px;
            }
            QListWidget::item {
                padding: 14px 10px;
                border-bottom: 1px solid #f1f3f5;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
            QListWidget::item:selected {
                background-color: #e7f1ff;
                color: #084298;
                font-weight: bold;
                border-left: 4px solid #0d6efd;
            }
        """)
        self.list_presets.currentItemChanged.connect(self.on_selection_changed)
        left_layout.addWidget(self.list_presets)

        self.label_selection_info = QLabel()
        self.label_selection_info.setStyleSheet("""
            color: #6c757d; 
            font-size: 10px; 
            padding: 5px 8px;
            background-color: #f8f9fa;
            border-radius: 3px;
        """)
        self.label_selection_info.setWordWrap(True)
        left_layout.addWidget(self.label_selection_info)

        # Pulsanti di gestione
        left_buttons_top = QHBoxLayout()
        left_buttons_top.setSpacing(8)
        
        self.btn_add_new = QPushButton("➕ Nuovo")
        self.btn_add_new.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 10px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        self.btn_add_new.clicked.connect(self.prepare_new_preset)
        
        self.btn_delete = QPushButton("🗑️ Elimina")
        self.btn_delete.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                padding: 10px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.btn_delete.clicked.connect(self.delete_selected_preset)
        
        left_buttons_top.addWidget(self.btn_add_new)
        left_buttons_top.addWidget(self.btn_delete)
        left_layout.addLayout(left_buttons_top)
        
        # Pulsanti di ordinamento
        left_buttons_bottom = QHBoxLayout()
        left_buttons_bottom.setSpacing(8)
        
        self.btn_move_up = QPushButton("⬆️ Sposta Su")
        self.btn_move_up.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 10px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.btn_move_up.clicked.connect(self.move_preset_up)
        self.btn_move_up.setEnabled(False)
        
        self.btn_move_down = QPushButton("⬇️ Sposta Giù")
        self.btn_move_down.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 10px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.btn_move_down.clicked.connect(self.move_preset_down)
        self.btn_move_down.setEnabled(False)
        
        left_buttons_bottom.addWidget(self.btn_move_up)
        left_buttons_bottom.addWidget(self.btn_move_down)
        left_layout.addLayout(left_buttons_bottom)
        
        main_layout.addWidget(left_widget)

        # LATO DESTRO: EDITOR
        right_widget = QWidget()
        right_widget.setStyleSheet("background-color: white;")
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(25, 20, 25, 20)
        right_layout.setSpacing(15)

        self.editor_header = QLabel("✏️ Modifica Configurazione")
        self.editor_header.setStyleSheet("""
            font-size: 16px;
            font-weight: bold;
            color: #2c3e50;
            padding: 10px;
            border-bottom: 2px solid #3498db;
        """)
        right_layout.addWidget(self.editor_header)

        self.label_error = QLabel()
        self.label_error.setStyleSheet("""
            color: #dc3545; 
            background-color: #f8d7da;
            border: 1px solid #f5c2c7;
            border-radius: 4px;
            padding: 8px;
            font-size: 11px;
        """)
        self.label_error.setWordWrap(True)
        self.label_error.hide()
        right_layout.addWidget(self.label_error)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background-color: white; }")
        
        form_container = QWidget()
        form_layout = QFormLayout(form_container)
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form_layout.setSpacing(15)
        form_layout.setContentsMargins(10, 10, 10, 10)

        # Nome Preset
        name_label = QLabel("Nome Preset:")
        name_label.setStyleSheet("font-weight: bold; color: #495057;")
        
        self.input_name = QLineEdit()
        self.input_name.setPlaceholderText("Nome identificativo del preset")
        self.input_name.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 2px solid #ced4da;
                border-radius: 5px;
                font-size: 13px;
            }
            QLineEdit:focus {
                border: 2px solid #3498db;
            }
            QLineEdit:disabled {
                background-color: #e9ecef;
            }
        """)
        self.input_name.textChanged.connect(self.clear_error)
        self.input_name.setToolTip("Nome che identifica questo preset (es: 'Acquisti 2025', 'Stampa Registri')")
        
        form_layout.addRow(name_label, self.input_name)

        # Separatore
        separator = QLabel()
        separator.setStyleSheet("""
            background-color: #dee2e6; 
            min-height: 2px; 
            max-height: 2px;
            margin: 10px 0px;
        """)
        form_layout.addRow(separator)

        # Intestazione colonne
        columns_header_widget = QWidget()
        columns_header_layout = QVBoxLayout(columns_header_widget)
        columns_header_layout.setSpacing(5)
        columns_header_layout.setContentsMargins(0, 0, 0, 0)
        
        columns_title = QLabel("📊 Definizione Colonne")
        columns_title.setStyleSheet("""
            font-weight: bold; 
            font-size: 13px;
            color: #2c3e50;
        """)
        
        columns_instructions = QLabel(
            "🔑 <b>Seleziona la colonna chiave</b> cliccando sul radio button.<br>"
            "Il radio button si attiva automaticamente quando compili il campo.<br>"
            "La colonna chiave è quella usata per confrontare le liste (es: ISBN, Codice Inventario)"
        )
        columns_instructions.setStyleSheet("""
            color: #6c757d; 
            font-size: 11px;
            background-color: #fff3cd;
            border: 1px solid #ffc107;
            border-radius: 4px;
            padding: 8px;
        """)
        columns_instructions.setWordWrap(True)
        
        columns_header_layout.addWidget(columns_title)
        columns_header_layout.addWidget(columns_instructions)
        form_layout.addRow(columns_header_widget)

        # Colonne con radio button
        self.key_group = QButtonGroup(self)
        self.key_group.setExclusive(True)
        self.column_rows = []

        for i in range(MAX_COLUMNS):
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(10)

            radio = QRadioButton()
            radio.setEnabled(False)
            radio.setStyleSheet("""
                QRadioButton::indicator {
                    width: 20px;
                    height: 20px;
                }
                QRadioButton::indicator:checked {
                    background-color: #e74c3c;
                    border: 2px solid #c0392b;
                    border-radius: 10px;
                }
                QRadioButton::indicator:disabled {
                    background-color: #e9ecef;
                    border: 2px solid #ced4da;
                }
            """)
            radio.setToolTip("Seleziona questa colonna come chiave")
            self.key_group.addButton(radio, i)

            edit = QLineEdit()
            edit.setPlaceholderText(f"Colonna {i+1} (Opzionale)")
            edit.setStyleSheet("""
                QLineEdit {
                    padding: 10px;
                    border: 2px solid #ced4da;
                    border-radius: 4px;
                    font-size: 12px;
                }
                QLineEdit:focus {
                    border: 2px solid #3498db;
                }
            """)
            edit.textChanged.connect(self.clear_error)
            edit.textChanged.connect(lambda text, r=radio: r.setEnabled(bool(text.strip())))

            def make_highlight_callback(edit_widget, radio_widget):
                def on_toggle(checked):
                    if checked:
                        edit_widget.setStyleSheet("""
                            QLineEdit {
                                padding: 10px;
                                border: 3px solid #e74c3c;
                                border-radius: 4px;
                                font-size: 12px;
                                background-color: #fff3cd;
                                font-weight: bold;
                            }
                            QLineEdit:focus {
                                border: 3px solid #c0392b;
                            }
                        """)
                    else:
                        edit_widget.setStyleSheet("""
                            QLineEdit {
                                padding: 10px;
                                border: 2px solid #ced4da;
                                border-radius: 4px;
                                font-size: 12px;
                            }
                            QLineEdit:focus {
                                border: 2px solid #3498db;
                            }
                        """)
                return on_toggle
            
            radio.toggled.connect(make_highlight_callback(edit, radio))

            row_layout.addWidget(radio)
            row_layout.addWidget(edit)

            self.column_rows.append((radio, edit))
            
            label = QLabel(f"Colonna {i+1}:")
            label.setStyleSheet("color: #6c757d; font-size: 12px;")
            form_layout.addRow(label, row_widget)

        scroll.setWidget(form_container)
        right_layout.addWidget(scroll)

        # Azioni Editor
        editor_actions = QHBoxLayout()
        editor_actions.setSpacing(10)
        
        self.btn_set_default = QPushButton("⭐ Imposta Predefinito")
        self.btn_set_default.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                padding: 12px 20px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        self.btn_set_default.clicked.connect(self.action_set_default)
        
        self.btn_save_changes = QPushButton("💾 Salva modifiche")
        self.btn_save_changes.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 12px 20px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.btn_save_changes.clicked.connect(self.action_save_preset)

        editor_actions.addWidget(self.btn_set_default)
        editor_actions.addStretch()
        editor_actions.addWidget(self.btn_save_changes)
        right_layout.addLayout(editor_actions)

        # Footer
        right_layout.addStretch()
        
        footer_separator = QLabel()
        footer_separator.setStyleSheet("""
            background-color: #dee2e6; 
            min-height: 1px; 
            max-height: 1px;
        """)
        right_layout.addWidget(footer_separator)

        final_buttons = QHBoxLayout()
        final_buttons.setSpacing(15)
        
        btn_cancel = QPushButton("✕ Chiudi")
        btn_cancel.setFixedWidth(120)
        btn_cancel.setFixedHeight(40)
        btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        btn_cancel.clicked.connect(self.reject)
        
        self.btn_use = QPushButton("🚀 USA QUESTO PRESET")
        self.btn_use.setFixedHeight(45)
        self.btn_use.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 0px 30px;
                font-size: 15px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        self.btn_use.clicked.connect(self.action_use_selected)
        
        final_buttons.addWidget(btn_cancel)
        final_buttons.addStretch()
        final_buttons.addWidget(self.btn_use)
        right_layout.addLayout(final_buttons)

        main_layout.addWidget(right_widget, 2)

    def load_preset_list(self):
        """Popola la lista dei preset con icone di stato"""
        self.list_presets.clear()
        
        if hasattr(self.parent(), 'preset_manager'):
            saved_order = self.parent().preset_manager.preset_order
        else:
            saved_order = []
        
        ordered_names = []
        remaining_names = set(self.presets.keys())
        
        for name in saved_order:
            if name in self.presets:
                ordered_names.append(name)
                remaining_names.discard(name)
        
        ordered_names.extend(sorted(remaining_names))
        
        for name in ordered_names:
            display_name = name
            
            # Prima aggiungi l'icona predefinito (se applicabile)
            if name == self.default_preset:
                display_name = f"⭐ {display_name}"
            
            # Poi aggiungi l'icona protetto/personalizzato
            if name in self.protected_presets:
                display_name = f"🔒 {display_name}"
            else:
                display_name = f"🟢 {display_name}"  # VERDE - preset personalizzati
            
            item = QListWidgetItem(display_name)
            item.setData(Qt.ItemDataRole.UserRole, name)
            self.list_presets.addItem(item)
        
        total = len(ordered_names)
        protected = len([n for n in ordered_names if n in self.protected_presets])
        custom = total - protected
        self.label_preset_count.setText(
            f"Totale: {total} preset ({custom} personalizzati, {protected} di sistema)"
        )

    def select_preset_in_list(self, name):
        """Sincronizza selezione grafica con nome logico"""
        for i in range(self.list_presets.count()):
            item = self.list_presets.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == name:
                self.list_presets.setCurrentRow(i)
                break

    def on_selection_changed(self, current_item, previous_item):
        """Carica preset usando radio button"""
        if not current_item:
            return
        
        clean_name = current_item.data(Qt.ItemDataRole.UserRole)
        self.current_selected_name = clean_name
        
        preset_data = self.presets.get(clean_name)
        if not preset_data:
            return

        self.is_new_preset_mode = False
        self.update_editor_mode()

        self.input_name.setText(clean_name)
        
        for radio, _ in self.column_rows:
            radio.setChecked(False)
        
        cols = preset_data['columns']
        key = preset_data['key_column']

        for i, (radio, edit) in enumerate(self.column_rows):
            if i < len(cols):
                edit.setText(cols[i])
                radio.setEnabled(True)
                if cols[i] == key:
                    radio.setChecked(True)
            else:
                edit.clear()
                radio.setEnabled(False)
        
        is_default = (clean_name == self.default_preset)
        is_protected = clean_name in self.protected_presets
        
        info_parts = []
        if is_default:
            info_parts.append("⭐ Preset predefinito")
        if is_protected:
            info_parts.append("🔒 Preset protetto")
        
        if info_parts:
            self.label_selection_info.setText(" • ".join(info_parts))
        else:
            self.label_selection_info.setText("Preset personalizzato")
        
        self.btn_delete.setEnabled(not is_protected)
        self.input_name.setEnabled(not is_protected)
        
        current_row = self.list_presets.currentRow()
        self.btn_move_up.setEnabled(current_row > 0)
        self.btn_move_down.setEnabled(current_row < self.list_presets.count() - 1)
        
        self.clear_error()

    def update_editor_mode(self):
        """Aggiorna UI in base alla modalità"""
        if self.is_new_preset_mode:
            self.editor_header.setText("🆕 Nuovo Preset")
            self.editor_header.setStyleSheet("""
                font-size: 16px;
                font-weight: bold;
                color: #27ae60;
                padding: 10px;
                border-bottom: 3px solid #27ae60;
            """)
            self.btn_save_changes.setText("➕ Crea preset")
            self.btn_set_default.setEnabled(False)
            self.btn_use.setEnabled(False)
            self.label_selection_info.setText(
                "✏️ Compila le colonne e clicca sul radio button per selezionare la chiave"
            )
        else:
            self.editor_header.setText("✏️ Modifica Configurazione")
            self.editor_header.setStyleSheet("""
                font-size: 16px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px;
                border-bottom: 2px solid #3498db;
            """)
            self.btn_save_changes.setText("💾 Salva modifiche")
            self.btn_set_default.setEnabled(True)
            self.btn_use.setEnabled(True)

    def show_error(self, message):
        """Mostra errore inline"""
        self.label_error.setText(f"⚠️ {message}")
        self.label_error.show()

    def clear_error(self):
        """Nascondi errore"""
        self.label_error.hide()

    def prepare_new_preset(self):
        """Pulisce form per nuovo preset"""
        self.list_presets.clearSelection()
        self.current_selected_name = None
        self.is_new_preset_mode = True
        
        self.input_name.clear()
        
        for radio, edit in self.column_rows:
            radio.setChecked(False)
            edit.clear()
        
        self.input_name.setEnabled(True)
        self.input_name.setFocus()
        
        self.update_editor_mode()
        self.clear_error()

    def action_save_preset(self):
        """Salva preset con validazione"""
        name = self.input_name.text().strip()
        
        columns = []
        key_column = None
        
        for i, (radio, edit) in enumerate(self.column_rows):
            text = edit.text().strip()
            if text:
                columns.append(text)
                if radio.isChecked() and text:
                    key_column = text
        
        if not key_column:
            self.show_error(
                "Seleziona una colonna chiave cliccando sul radio button. "
                "Il radio button si abilita automaticamente quando compili il campo."
            )
            return
        
        is_valid, error_msg = self.validate_preset_data(name, key_column, columns)
        if not is_valid:
            self.show_error(error_msg)
            return
        
        # CASO 1: Stiamo modificando un preset E cambiamo il nome
        if not self.is_new_preset_mode and self.current_selected_name and name != self.current_selected_name:
            # Controlla se il nuovo nome esiste già
            if name in self.presets:
                reply = QMessageBox.question(
                    self,
                    "⚠️ Preset Esistente",
                    f"Il preset '<b>{name}</b>' esiste già.\n\n"
                    "Vuoi sovrascriverlo con questa nuova configurazione?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    return
            
            # Elimina il vecchio preset
            if self.current_selected_name in self.presets:
                del self.presets[self.current_selected_name]
                
                parent = self.parent()
                if hasattr(parent, 'preset_manager'):
                    if self.current_selected_name in parent.preset_manager.preset_order:
                        parent.preset_manager.preset_order.remove(self.current_selected_name)
        
        # CASO 2: Stiamo creando un nuovo preset E il nome esiste già
        elif self.is_new_preset_mode and name in self.presets:
            reply = QMessageBox.question(
                self,
                "⚠️ Preset Esistente",
                f"Il preset '<b>{name}</b>' esiste già.\n\n"
                "Vuoi sovrascriverlo con questa nuova configurazione?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
        
        # CASO 3: Stiamo modificando con lo stesso nome -> nessun controllo, sovrascrivi direttamente
        
        # Salva il preset (nuovo o aggiornato)
        self.result = {
            'action': 'create',
            'preset_name': name,
            'key_column': key_column,
            'columns': columns,
            'set_as_default': False
        }
        
        self.presets[name] = {
            'key_column': key_column,
            'columns': columns
        }
        
        parent = self.parent()
        if hasattr(parent, 'preset_manager'):
            if name not in parent.preset_manager.preset_order:
                parent.preset_manager.preset_order.append(name)
            
            # IMPORTANTE: Sincronizza i preset con il parent e salva su disco
            parent.preset_manager.presets = self.presets.copy()
            parent.preset_manager.save_presets()
        
        self.load_preset_list()
        self.select_preset_in_list(name)
        
        self.is_new_preset_mode = False
        self.current_selected_name = name
        self.update_editor_mode()
        
        self.label_selection_info.setText(
            f"✅ Preset '{name}' salvato con successo! "
            f"Colonna chiave: {key_column}\n"
            f"Premi '🚀 USA QUESTO PRESET' per applicarlo"
        )
    
    
    
    def action_set_default(self):
        """Imposta preset come predefinito"""
        if not self.current_selected_name:
            self.show_error("Seleziona un preset prima di impostarlo come predefinito")
            return
        
        if self.current_selected_name == self.default_preset:
            QMessageBox.information(
                self,
                "Già Predefinito",
                f"Il preset '<b>{self.current_selected_name}</b>' è già impostato come predefinito!"
            )
            return
        
        self.result = {
            'action': 'set_default', 
            'preset_name': self.current_selected_name
        }
        self.accept()

    def action_use_selected(self):
        """Usa il preset selezionato"""
        if not self.current_selected_name:
            self.show_error("Seleziona un preset prima di usarlo")
            return
        
        if (self.result and 
            self.result.get('action') == 'create' and 
            self.result.get('preset_name') == self.current_selected_name):
            self.accept()
        else:
            self.result = {
                'action': 'select', 
                'preset_name': self.current_selected_name
            }
            self.accept()

    def delete_selected_preset(self):
        """Elimina preset selezionato"""
        if not self.current_selected_name:
            self.show_error("Seleziona un preset prima di eliminarlo")
            return
        
        if self.current_selected_name in self.protected_presets:
            self.show_error(
                f"Il preset '{self.current_selected_name}' è protetto e non può essere eliminato"
            )
            return
        
        confirm = QMessageBox.question(
            self, 
            "🗑️ Conferma Eliminazione", 
            f"Sei sicuro di voler eliminare definitivamente il preset "
            f"'<b>{self.current_selected_name}</b>'?\n\n"
            "Questa azione non può essere annullata.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if confirm == QMessageBox.StandardButton.Yes:
            preset_to_delete = self.current_selected_name
            
            del self.presets[preset_to_delete]
            
            parent = self.parent()
            if hasattr(parent, 'preset_manager'):
                if preset_to_delete in parent.preset_manager.preset_order:
                    parent.preset_manager.preset_order.remove(preset_to_delete)
                
                if self.default_preset == preset_to_delete:
                    self.default_preset = None
                
                parent.preset_manager.presets = self.presets.copy()
                parent.preset_manager.default_preset = self.default_preset
                parent.preset_manager.save_presets()
            
            self.load_preset_list()
            if self.presets:
                first_preset_name = list(self.presets.keys())[0]
                self.select_preset_in_list(first_preset_name)
            else:
                self.prepare_new_preset()
            
            self.label_selection_info.setText(
                f"✅ Preset '{preset_to_delete}' eliminato con successo"
            )

    def move_preset_up(self):
        """Sposta il preset selezionato verso l'alto nella lista"""
        current_row = self.list_presets.currentRow()
        
        if current_row <= 0:
            return
        
        current_item = self.list_presets.currentItem()
        preset_name = current_item.data(Qt.ItemDataRole.UserRole)
        
        self._update_preset_order()
        parent = self.parent()
        
        if hasattr(parent, 'preset_manager'):
            preset_order = parent.preset_manager.preset_order
            if preset_name in preset_order:
                current_index = preset_order.index(preset_name)
                if current_index > 0:
                    preset_order[current_index], preset_order[current_index - 1] = \
                        preset_order[current_index - 1], preset_order[current_index]
        
        self.load_preset_list()
        self.list_presets.setCurrentRow(current_row - 1)
        
        if hasattr(parent, 'preset_manager'):
            parent.preset_manager.save_presets()
        
        self.label_selection_info.setText("✅ Preset spostato verso l'alto")

    def move_preset_down(self):
        """Sposta il preset selezionato verso il basso nella lista"""
        current_row = self.list_presets.currentRow()
        
        if current_row >= self.list_presets.count() - 1:
            return
        
        current_item = self.list_presets.currentItem()
        preset_name = current_item.data(Qt.ItemDataRole.UserRole)
        
        self._update_preset_order()
        parent = self.parent()
        
        if hasattr(parent, 'preset_manager'):
            preset_order = parent.preset_manager.preset_order
            if preset_name in preset_order:
                current_index = preset_order.index(preset_name)
                if current_index < len(preset_order) - 1:
                    preset_order[current_index], preset_order[current_index + 1] = \
                        preset_order[current_index + 1], preset_order[current_index]
        
        self.load_preset_list()
        self.list_presets.setCurrentRow(current_row + 1)
        
        if hasattr(parent, 'preset_manager'):
            parent.preset_manager.save_presets()
        
        self.label_selection_info.setText("✅ Preset spostato verso il basso")

    def _update_preset_order(self):
        """Sincronizza l'ordine corrente della lista con preset_order del parent"""
        parent = self.parent()
        if not hasattr(parent, 'preset_manager'):
            return
        
        current_order = []
        
        for i in range(self.list_presets.count()):
            item = self.list_presets.item(i)
            preset_name = item.data(Qt.ItemDataRole.UserRole)
            current_order.append(preset_name)
        
        parent.preset_manager.preset_order = current_order

    def get_result(self):
        """Restituisce il risultato della dialog"""
        return self.result