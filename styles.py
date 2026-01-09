"""
Stili CSS per l'interfaccia grafica
"""

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
    
    /* PULSANTI SPECIALIZZATI */
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