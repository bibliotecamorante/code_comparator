"""
Code Comparator - Applicazione per confronto liste
Entry point principale
"""
import sys
import logging
from PyQt6.QtWidgets import QApplication, QMessageBox

from gui import CodeComparator

# Configurazione logging
logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


def global_exception_handler(exctype, value, tb):
    """Gestore di eccezioni globale per errori non gestiti"""
    logging.exception("Errore non gestito:", exc_info=(exctype, value, tb))
    
    if QApplication.instance():
        QMessageBox.critical(
            None, 
            "⌠Errore Critico",
            f"Si è verificato un errore imprevisto:\n\n{str(value)}\n\n"
            "L'applicazione potrebbe non funzionare correttamente."
        )
    
    sys.__excepthook__(exctype, value, tb)


if __name__ == "__main__":
    sys.excepthook = global_exception_handler
    
    app = QApplication(sys.argv)
    window = CodeComparator()
    window.show()
    sys.exit(app.exec())