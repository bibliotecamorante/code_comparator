"""
Gestione centralizzata delle tabelle
"""
import pandas as pd
from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView, QApplication
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from config import HIGHLIGHT_COLOR, MATCH_COLOR, MISMATCH_COLOR, WHITE_COLOR


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
        
        table.verticalHeader().setDefaultSectionSize(35)
        table.verticalHeader().setMinimumSectionSize(35)
    
    def enable_manual_sorting(self, table):
        """Abilita sorting manuale con indicatori Unicode visibili"""
        header = table.horizontalHeader()
        header.setSortIndicatorShown(False)
        header.setSectionsClickable(True)
        
        if not hasattr(self, '_sort_handlers'):
            self._sort_handlers = {}
        
        if table in self._sort_handlers:
            try:
                header.sectionClicked.disconnect(self._sort_handlers[table])
            except TypeError:
                pass
        
        def on_header_clicked(logical_index):
            if self._sort_column == logical_index:
                self._sort_order = Qt.SortOrder.DescendingOrder if self._sort_order == Qt.SortOrder.AscendingOrder else Qt.SortOrder.AscendingOrder
            else:
                self._sort_column = logical_index
                self._sort_order = Qt.SortOrder.AscendingOrder
            
            table.setSortingEnabled(True)
            table.sortItems(logical_index, self._sort_order)
            table.setSortingEnabled(False)
            
            for col in range(table.columnCount()):
                if col < len(self.columns):
                    original_text = self.columns[col]
                    if col == logical_index:
                        arrow = " ▲" if self._sort_order == Qt.SortOrder.AscendingOrder else " ▼"
                        table.horizontalHeaderItem(col).setText(original_text + arrow)
                    else:
                        table.horizontalHeaderItem(col).setText(original_text)
        
        self._sort_handlers[table] = on_header_clicked
        header.sectionClicked.connect(on_header_clicked)

    def reset_sorting(self, table):
        """Rimuove gli indicatori di sorting dalla tabella"""
        self._sort_column = None
        self._sort_order = None
        
        for col in range(table.columnCount()):
            if col < len(self.columns):
                table.horizontalHeaderItem(col).setText(self.columns[col])
    
    def update_table_display(self, table_widget: QTableWidget, df: pd.DataFrame, 
                            highlight_from: int | None = None,
                            is_match_result: bool = False,
                            is_mismatch_result: bool = False) -> None:
        """Popola la QTableWidget con i dati del DataFrame"""
        display_columns = [col for col in self.columns if col in df.columns and not col.startswith('__')]
        display_df = df[display_columns]
        
        show_cursor = len(display_df) > 1000
        if show_cursor:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        
        was_sorting_enabled = table_widget.isSortingEnabled()
        table_widget.setSortingEnabled(False)
        table_widget.setUpdatesEnabled(False)
        
        try:
            table_widget.setRowCount(len(display_df))
            
            is_res = is_match_result or is_mismatch_result
            bg_color = QColor(MATCH_COLOR) if is_match_result else QColor(MISMATCH_COLOR)
            color_highlight = QColor(HIGHLIGHT_COLOR)
            color_white = QColor(WHITE_COLOR)
            
            for i in range(len(display_df)):
                for j in range(len(display_columns)):
                    if not table_widget.item(i, j):
                        table_widget.setItem(i, j, QTableWidgetItem())

            for i, row in enumerate(display_df.itertuples(index=True)):
                real_index = row.Index
                for j, col_name in enumerate(display_columns):
                    value = row[j + 1] if j + 1 < len(row) else ""
                    
                    item = table_widget.item(i, j)
                    val_str = str(value) if not pd.isna(value) else ""
                    item.setText(val_str)
                    item.setToolTip(val_str if len(val_str) > 30 else "")
                    item.setData(Qt.ItemDataRole.UserRole, real_index)
                    
                    if highlight_from is not None and i >= highlight_from:
                        item.setBackground(color_highlight)
                    elif is_res:
                        item.setBackground(bg_color)
                    else:
                        item.setBackground(color_white)
        
        finally:
            table_widget.setUpdatesEnabled(True)
            table_widget.setSortingEnabled(was_sorting_enabled)
            
            if show_cursor:
                QApplication.restoreOverrideCursor()