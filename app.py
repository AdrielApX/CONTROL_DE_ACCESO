import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QPalette, QColor

# CONECTAMOS CON LA BASE DE DATOS Y LA VENTANA PRINCIPAL
import base_datos
from ventana_principal import SistemaBiblioteca

# =====================================================================
# INICIAR APLICACIÓN
# =====================================================================
if __name__ == "__main__":
    # Arrancamos primero la base de datos
    base_datos.inicializar_bd()
    
    app = QApplication(sys.argv)
    
    paleta_clara = QPalette()
    paleta_clara.setColor(QPalette.ColorRole.Window, QColor("#F8FAFC"))
    paleta_clara.setColor(QPalette.ColorRole.WindowText, QColor("#1E293B"))
    paleta_clara.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
    paleta_clara.setColor(QPalette.ColorRole.AlternateBase, QColor("#F1F5F9"))
    paleta_clara.setColor(QPalette.ColorRole.Text, QColor("#1E293B"))
    paleta_clara.setColor(QPalette.ColorRole.Button, QColor("#E2E8F0"))
    paleta_clara.setColor(QPalette.ColorRole.ButtonText, QColor("#1E293B"))
    paleta_clara.setColor(QPalette.ColorRole.Highlight, QColor("#0EA5E9"))
    paleta_clara.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(paleta_clara)
    
    app.setStyleSheet("""
        QMainWindow, QDialog, QMessageBox { background-color: #F8FAFC; }
        QLabel { font-family: 'Segoe UI'; color: #1E293B; }
        QLineEdit { font-family: 'Segoe UI'; color: #1E293B; background-color: white; border: 1px solid #CBD5E1; border-radius: 6px; padding: 5px; }
        QTableWidget { background-color: white; color: #1E293B; gridline-color: #E2E8F0; border: 1px solid #CBD5E1; font-size: 14px; }
        QHeaderView::section { background-color: #F1F5F9; color: #1E293B; font-weight: bold; border: 1px solid #E2E8F0; font-size: 14px; padding: 5px; }
        QComboBox { background-color: white; color: #1E293B; border: 1px solid #CBD5E1; border-radius: 6px; padding: 5px; min-height: 30px; }
        QComboBox QAbstractItemView { background-color: white; color: #1E293B; selection-background-color: #0EA5E9; selection-color: white; }
        QPushButton { font-family: 'Segoe UI'; font-size: 14px; border-radius: 6px; padding: 8px; }
    """)
    
    ventana = SistemaBiblioteca()
    ventana.show()
    sys.exit(app.exec())