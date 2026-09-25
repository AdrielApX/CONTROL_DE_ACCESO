import os
import subprocess
from PyQt6.QtWidgets import QLineEdit

# =====================================================================
# COMPONENTES TÁCTILES PERSONALIZADOS
# =====================================================================
def abrir_teclado_virtual():
    try:
        tabtip = r"C:\Program Files\Common Files\microsoft shared\ink\TabTip.exe"
        if os.path.exists(tabtip):
            subprocess.Popen(tabtip)
    except Exception:
        pass

class LineEditTactil(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QLineEdit {
                font-size: 18px; 
                min-height: 50px; 
                border: 2px solid #CBD5E1; 
                border-radius: 8px; 
                padding: 5px 10px;
                background-color: white;
                color: #1E293B;
            }
            QLineEdit:focus { border: 2px solid #0EA5E9; }
        """)
    
    def focusInEvent(self, event):
        super().focusInEvent(event)
        abrir_teclado_virtual()