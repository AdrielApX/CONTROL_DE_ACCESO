import sys
import os
import sqlite3
import subprocess
from datetime import datetime, timedelta

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QComboBox, QTableWidget, 
    QTableWidgetItem, QStackedWidget, QFormLayout, QHeaderView, 
    QMessageBox, QDialog, QGridLayout, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QPalette, QColor  # Se agregaron QPalette y QColor para el control de tema
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# =====================================================================
# CONFIGURACIÓN GENERAL
# =====================================================================
DB_NAME = "biblioteca_utesa.db"

# =====================================================================
# 1. BASE DE DATOS E INICIALIZACIÓN
# =====================================================================
def inicializar_bd():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        identificacion TEXT UNIQUE NOT NULL,
        nombre_completo TEXT NOT NULL,
        tipo_usuario TEXT NOT NULL,
        carrera TEXT
    );
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre_pc TEXT UNIQUE NOT NULL,
        activo INTEGER DEFAULT 1
    );
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS accesos_salon (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario_id INTEGER,
        fecha_hora TEXT DEFAULT (datetime('now', 'localtime')),
        FOREIGN KEY(usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
    );
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS prestamos_pc (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario_id INTEGER,
        equipo_id INTEGER,
        fecha_prestamo TEXT DEFAULT (date('now', 'localtime')),
        FOREIGN KEY(usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE,
        FOREIGN KEY(equipo_id) REFERENCES equipos(id) ON DELETE CASCADE
    );
    """)
    
    cursor.execute("SELECT COUNT(*) FROM equipos")
    if cursor.fetchone()[0] == 0:
        for i in range(1, 5):
            cursor.execute("INSERT INTO equipos (nombre_pc, activo) VALUES (?, 1)", (f"PC {i}",))
            
    conn.commit()
    conn.close()

def abrir_teclado_virtual():
    try:
        tabtip = r"C:\Program Files\Common Files\microsoft shared\ink\TabTip.exe"
        if os.path.exists(tabtip):
            subprocess.Popen(tabtip)
    except Exception:
        pass

# =====================================================================
# 2. COMPONENTES TÁCTILES PERSONALIZADOS
# =====================================================================
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

# =====================================================================
# 3. INTERFAZ PRINCIPAL
# =====================================================================
class SistemaBiblioteca(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("UTESA Plus - Control de Biblioteca")
        self.resize(1200, 750)
        
        self.usuario_actual_prestamo = None
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        
        self.vista_formulario_principal = QWidget()
        self.vista_kiosco_acceso = QWidget()
        self.vista_metricas = QWidget()
        
        self.stacked_widget.addWidget(self.vista_formulario_principal)
        self.stacked_widget.addWidget(self.vista_kiosco_acceso)
        self.stacked_widget.addWidget(self.vista_metricas)
        
        self.crear_vista_formulario_principal()
        self.crear_vista_kiosco_acceso()
        self.crear_vista_metricas()
        
        self.stacked_widget.setCurrentIndex(0)

    def cargar_imagen_branding(self, size=80):
        lbl_img = QLabel()
        pixmap = QPixmap("mi_foto.png")
        if not pixmap.isNull():
            lbl_img.setPixmap(pixmap.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            lbl_img.setText("👤")
            lbl_img.setStyleSheet(f"font-size: {size//2}px; color: #94A3B8;")
        lbl_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        return lbl_img

    def barra_navegacion(self, index_actual):
        layout = QHBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        
        btn1 = QPushButton("Formulario de Asignación de PCs")
        btn2 = QPushButton("Acceso Táctil")
        btn3 = QPushButton("Métricas y Reportes")
        
        for btn in [btn1, btn2, btn3]:
            btn.setMinimumHeight(40)
            btn.setStyleSheet("background-color: #E2E8F0; color: #1E293B; font-weight: 500;")
            
        btn1.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(0))
        btn2.clicked.connect(self.activar_modo_kiosco)
        btn3.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(2))
        
        if index_actual == 0: btn1.setStyleSheet("background-color: #0EA5E9; color: white; font-weight: bold;")
        if index_actual == 2: btn3.setStyleSheet("background-color: #0EA5E9; color: white; font-weight: bold;")
        
        layout.addWidget(btn1)
        layout.addWidget(btn2)
        layout.addWidget(btn3)
        return layout

    # -----------------------------------------------------------------
    # VISTA 1: FORMULARIO PRINCIPAL Y EDICIÓN DE PC'S
    # -----------------------------------------------------------------
    def crear_vista_formulario_principal(self):
        layout_global = QVBoxLayout(self.vista_formulario_principal)
        layout_global.addLayout(self.barra_navegacion(0))
        
        layout_contenido = QHBoxLayout()
        
        sec_formulario = QVBoxLayout()
        lbl_f_titulo = QLabel("Asignación de Computadoras")
        lbl_f_titulo.setStyleSheet("font-size: 18px; font-weight: bold; color: #0F172A; margin-bottom: 10px;")
        sec_formulario.addWidget(lbl_f_titulo)
        
        form_layout = QFormLayout()
        self.txt_form_buscar = QLineEdit()
        self.txt_form_buscar.setPlaceholderText("Ingrese Cédula o Matrícula y presione Enter")
        self.txt_form_buscar.setMinimumHeight(35)
        self.txt_form_buscar.returnPressed.connect(self.buscar_usuario_formulario)
        
        self.lbl_form_datos = QLabel("Usuario: No seleccionado\nTipo: -\nCarrera: -")
        self.lbl_form_datos.setStyleSheet("background-color: #F1F5F9; color: #1E293B; padding: 10px; border-radius: 6px; font-size: 14px;")
        
        self.txt_form_fecha = QLineEdit(datetime.now().strftime("%Y-%m-%d"))
        self.txt_form_fecha.setReadOnly(True)
        self.txt_form_fecha.setMinimumHeight(35)
        
        self.cmb_form_pcs = QComboBox()
        self.cmb_form_pcs.setMinimumHeight(35)
        self.actualizar_combo_pcs()
        
        form_layout.addRow("Identificación:", self.txt_form_buscar)
        form_layout.addRow("Datos del Usuario:", self.lbl_form_datos)
        form_layout.addRow("Día (Actual):", self.txt_form_fecha)
        form_layout.addRow("PC a Asignar:", self.cmb_form_pcs)
        sec_formulario.addLayout(form_layout)
        
        btn_guardar_p = QPushButton("Guardar Asignación")
        btn_guardar_p.setStyleSheet("background-color: #10B981; color: white; font-weight: bold; min-height: 45px; font-size: 15px;")
        btn_guardar_p.clicked.connect(self.guardar_asignacion_pc)
        sec_formulario.addWidget(btn_guardar_p)
        sec_formulario.addStretch()
        
        sec_edicion = QVBoxLayout()
        lbl_e_titulo = QLabel("Editor de Formulario (PCs)")
        lbl_e_titulo.setStyleSheet("font-size: 18px; font-weight: bold; color: #0F172A; margin-bottom: 10px;")
        sec_edicion.addWidget(lbl_e_titulo)
        
        self.table_pcs = QTableWidget(0, 2)
        self.table_pcs.setHorizontalHeaderLabels(["Nombre PC", "Estado"])
        self.table_pcs.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.actualizar_tabla_pcs()
        sec_edicion.addWidget(self.table_pcs)
        
        layout_agregar_pc = QHBoxLayout()
        self.txt_nueva_pc = QLineEdit()
        self.txt_nueva_pc.setPlaceholderText("Ej: PC 5")
        self.txt_nueva_pc.setMinimumHeight(35)
        btn_agregar_pc = QPushButton("Agregar PC")
        btn_agregar_pc.setStyleSheet("background-color: #0F172A; color: white; font-weight: 500; min-height: 35px;")
        btn_agregar_pc.clicked.connect(self.agregar_nueva_pc)
        
        layout_agregar_pc.addWidget(self.txt_nueva_pc)
        layout_agregar_pc.addWidget(btn_agregar_pc)
        sec_edicion.addLayout(layout_agregar_pc)
        
        btn_cambiar_estado = QPushButton("Alternar Activo / Inactivo")
        btn_cambiar_estado.setStyleSheet("background-color: #EF4444; color: white; min-height: 35px;")
        btn_cambiar_estado.clicked.connect(self.alternar_estado_pc)
        sec_edicion.addWidget(btn_cambiar_estado)
        
        frame_izq = QFrame()
        frame_izq.setLayout(sec_formulario)
        frame_der = QFrame()
        frame_der.setLayout(sec_edicion)
        
        layout_contenido.addWidget(frame_izq, 1)
        layout_contenido.addWidget(frame_der, 1)
        layout_global.addLayout(layout_contenido)

    def actualizar_combo_pcs(self):
        self.cmb_form_pcs.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre_pc FROM equipos WHERE activo = 1")
        for row in cursor.fetchall():
            self.cmb_form_pcs.addItem(row[1], row[0])
        conn.close()

    def actualizar_tabla_pcs(self):
        self.table_pcs.setRowCount(0)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre_pc, activo FROM equipos")
        for row_idx, row in enumerate(cursor.fetchall()):
            self.table_pcs.insertRow(row_idx)
            item_nombre = QTableWidgetItem(row[1])
            item_nombre.setData(Qt.ItemDataRole.UserRole, row[0])
            item_estado = QTableWidgetItem("Activo" if row[2] == 1 else "Inactivo")
            
            item_nombre.setForeground(Qt.GlobalColor.black)
            item_estado.setForeground(Qt.GlobalColor.black)
            
            self.table_pcs.setItem(row_idx, 0, item_nombre)
            self.table_pcs.setItem(row_idx, 1, item_estado)
        conn.close()

    def buscar_usuario_formulario(self):
        ident = self.txt_form_buscar.text().strip()
        if not ident: return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre_completo, tipo_usuario, carrera FROM usuarios WHERE identificacion = ?", (ident,))
        res = cursor.fetchone()
        conn.close()
        
        if res:
            self.usuario_actual_prestamo = res[0]
            self.lbl_form_datos.setText(f"Usuario: {res[1]}\nTipo: {res[2]}\nCarrera: {res[3] if res[3] else 'N/A'}")
        else:
            self.usuario_actual_prestamo = None
            self.lbl_form_datos.setText("Usuario no registrado.")
            QMessageBox.warning(self, "Aviso", "El usuario no existe. Debe registrarse en el sistema.")

    def agregar_nueva_pc(self):
        nombre = self.txt_nueva_pc.text().strip()
        if not nombre: return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO equipos (nombre_pc, activo) VALUES (?, 1)", (nombre,))
            conn.commit()
            self.txt_nueva_pc.clear()
            self.actualizar_tabla_pcs()
            self.actualizar_combo_pcs()
        except sqlite3.IntegrityError:
            QMessageBox.critical(self, "Error", "Ese nombre de PC ya existe.")
        finally:
            conn.close()

    def alternar_estado_pc(self):
        fila_sel = self.table_pcs.currentRow()
        if fila_sel < 0: return
        id_pc = self.table_pcs.item(fila_sel, 0).data(Qt.ItemDataRole.UserRole)
        estado_actual = self.table_pcs.item(fila_sel, 1).text()
        nuevo_estado = 0 if estado_actual == "Activo" else 1
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("UPDATE equipos SET activo = ? WHERE id = ?", (nuevo_estado, id_pc))
        conn.commit()
        conn.close()
        self.actualizar_tabla_pcs()
        self.actualizar_combo_pcs()

    def guardar_asignacion_pc(self):
        if not self.usuario_actual_prestamo: return
        id_pc = self.cmb_form_pcs.currentData()
        if not id_pc: return
            
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO prestamos_pc (usuario_id, equipo_id) VALUES (?, ?)", (self.usuario_actual_prestamo, id_pc))
        conn.commit()
        conn.close()
        
        QMessageBox.information(self, "Éxito", "Asignación guardada correctamente.")
        self.txt_form_buscar.clear()
        self.lbl_form_datos.setText("Usuario: No seleccionado\nTipo: -\nCarrera: -")
        self.usuario_actual_prestamo = None

    # -----------------------------------------------------------------
    # VISTA 2: KIOSCO TÁCTIL
    # -----------------------------------------------------------------
    def crear_vista_kiosco_acceso(self):
        self.layout_kiosco_master = QVBoxLayout(self.vista_kiosco_acceso)
        self.layout_kiosco_master.setContentsMargins(40, 40, 40, 40)
        self.kiosco_stack = QStackedWidget()
        self.layout_kiosco_master.addWidget(self.kiosco_stack)
        
        # 2.1 INICIO KIOSCO
        self.p_kiosco_inicio = QWidget()
        lay_ini = QVBoxLayout(self.p_kiosco_inicio)
        lay_ini.setSpacing(30)
        
        lbl_k_tit = QLabel("SISTEMA DE ACCESO - BIBLIOTECA")
        lbl_k_tit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_k_tit.setStyleSheet("font-size: 32px; font-weight: bold; color: #0F172A; margin-bottom: 20px;")
        lay_ini.addWidget(lbl_k_tit)
        
        btn_k_ingresar = QPushButton("INGRESAR AL SALÓN")
        btn_k_ingresar.setStyleSheet("background-color: #0EA5E9; color: white; font-size: 24px; font-weight: bold; border-radius: 15px; min-height: 140px;")
        btn_k_ingresar.clicked.connect(lambda: self.kiosco_stack.setCurrentIndex(1))
        
        btn_k_registrar = QPushButton("REGISTRARME COMO NUEVO USUARIO")
        btn_k_registrar.setStyleSheet("background-color: #0F172A; color: white; font-size: 24px; font-weight: bold; border-radius: 15px; min-height: 140px;")
        btn_k_registrar.clicked.connect(lambda: self.kiosco_stack.setCurrentIndex(2))
        
        btn_k_salir = QPushButton("⬅SALIR")
        btn_k_salir.setStyleSheet("background-color: transparent; border: none; font-size: 16px; color: #64748B;")
        btn_k_salir.clicked.connect(self.salir_modo_kiosco)
        
        lay_ini.addWidget(btn_k_ingresar)
        lay_ini.addWidget(btn_k_registrar)
        lay_ini.addWidget(btn_k_salir, 0, Qt.AlignmentFlag.AlignLeft)
        self.kiosco_stack.addWidget(self.p_kiosco_inicio)
        
        # 2.2 KEYPAD DE INGRESO
        self.p_kiosco_ingreso = QWidget()
        lay_ing = QHBoxLayout(self.p_kiosco_ingreso)
        
        lay_ing_izq = QVBoxLayout()
        lbl_ing_tit = QLabel("Digite su Cédula o Matrícula")
        lbl_ing_tit.setStyleSheet("font-size: 24px; font-weight: bold; color: #1E293B;")
        
        self.txt_ingreso_id = QLineEdit()
        self.txt_ingreso_id.setStyleSheet("font-size: 28px; min-height: 60px; border: 3px solid #0EA5E9; border-radius: 8px; padding: 5px; background-color: white; color: #1E293B;")
        
        btn_ing_aceptar = QPushButton("CONFIRMAR ENTRADA")
        btn_ing_aceptar.setStyleSheet("background-color: #10B981; color: white; font-size: 20px; font-weight: bold; min-height: 60px; border-radius: 8px;")
        btn_ing_aceptar.clicked.connect(self.procesar_ingreso_salon)
        
        btn_ing_volver = QPushButton("Volver Atrás")
        btn_ing_volver.setStyleSheet("background-color: #64748B; color: white; font-size: 16px; min-height: 45px;")
        btn_ing_volver.clicked.connect(self.limpiar_volver_kiosco)
        
        lay_ing_izq.addWidget(lbl_ing_tit)
        lay_ing_izq.addWidget(self.txt_ingreso_id)
        lay_ing_izq.addWidget(btn_ing_aceptar)
        lay_ing_izq.addStretch()
        lay_ing_izq.addWidget(btn_ing_volver)
        
        grid_keypad = QGridLayout()
        grid_keypad.setSpacing(10)
        botones_pad = ['1','2','3','4','5','6','7','8','9','Borrar','0','Limpiar']
        row, col = 0, 0
        for boton in botones_pad:
            btn_pad = QPushButton(boton)
            btn_pad.setMinimumSize(85, 85)
            btn_pad.setStyleSheet("background-color: white; color: #1E293B; border: 2px solid #CBD5E1; font-size: 22px; font-weight: bold; border-radius: 10px;")
            btn_pad.clicked.connect(self.tecla_pad_presionada)
            grid_keypad.addWidget(btn_pad, row, col)
            col += 1
            if col > 2:
                col = 0
                row += 1
                
        lay_ing.addLayout(lay_ing_izq, 4)
        lay_ing.addLayout(grid_keypad, 3)
        self.kiosco_stack.addWidget(self.p_kiosco_ingreso)
        
        # 2.3 REGISTRO TÁCTIL
        self.p_kiosco_registro = QWidget()
        lay_reg = QVBoxLayout(self.p_kiosco_registro)
        
        lay_header_reg = QHBoxLayout()
        lbl_reg_tit = QLabel("Formulario de Registro")
        lbl_reg_tit.setStyleSheet("font-size: 24px; font-weight: bold; color: #0F172A;")
        
        lbl_branding = self.cargar_imagen_branding(size=70)
        lay_header_reg.addWidget(lbl_reg_tit)
        lay_header_reg.addStretch()
        lay_header_reg.addWidget(lbl_branding)
        lay_reg.addLayout(lay_header_reg)
        
        form_reg = QFormLayout()
        form_reg.setSpacing(15)
        
        self.txt_reg_id = LineEditTactil()
        self.txt_reg_nombre = LineEditTactil()
        
        self.cmb_reg_tipo = QComboBox()
        self.cmb_reg_tipo.addItems(["Estudiante", "Profesor", "Empleado", "Invitado"])
        self.cmb_reg_tipo.setStyleSheet("background-color: white; color: #1E293B; font-size: 18px; min-height: 50px; border: 1px solid #CBD5E1;")
        self.cmb_reg_tipo.currentTextChanged.connect(self.evaluar_obligatoriedad_carrera)
        
        self.txt_reg_carrera = LineEditTactil()
        self.lbl_alerta_carrera = QLabel("* Obligatorio para Estudiantes")
        self.lbl_alerta_carrera.setStyleSheet("color: #EF4444; font-weight: bold; font-size: 13px;")
        
        form_reg.addRow(QLabel("Cédula / Matrícula:"), self.txt_reg_id)
        form_reg.addRow(QLabel("Nombre Completo:"), self.txt_reg_nombre)
        form_reg.addRow(QLabel("Tipo de Usuario:"), self.cmb_reg_tipo)
        form_reg.addRow(QLabel("Carrera:"), self.txt_reg_carrera)
        form_reg.addRow("", self.lbl_alerta_carrera)
        lay_reg.addLayout(form_reg)
        
        btn_reg_guardar = QPushButton("GUARDAR REGISTRO E INGRESAR")
        btn_reg_guardar.setStyleSheet("background-color: #10B981; color: white; font-size: 20px; font-weight: bold; min-height: 60px; border-radius: 8px;")
        btn_reg_guardar.clicked.connect(self.procesar_registro_usuario)
        
        btn_reg_volver = QPushButton("Volver Atrás")
        btn_reg_volver.setStyleSheet("background-color: #64748B; color: white; font-size: 16px; min-height: 45px;")
        btn_reg_volver.clicked.connect(self.limpiar_volver_kiosco)
        
        lay_reg.addWidget(btn_reg_guardar)
        lay_reg.addStretch()
        lay_reg.addWidget(btn_reg_volver)
        self.kiosco_stack.addWidget(self.p_kiosco_registro)

    def activar_modo_kiosco(self):
        self.stacked_widget.setCurrentIndex(1)
        self.kiosco_stack.setCurrentIndex(0)
        self.showFullScreen()

    def salir_modo_kiosco(self):
        dialogo = QDialog(self)
        dialogo.setWindowTitle("Seguridad")
        lay_d = QVBoxLayout(dialogo)
        lay_d.addWidget(QLabel("Contraseña administrativa:"))
        txt_pass = QLineEdit()
        txt_pass.setEchoMode(QLineEdit.EchoMode.Password)
        lay_d.addWidget(txt_pass)
        btn_d = QPushButton("Validar")
        btn_d.setStyleSheet("background-color: #0F172A; color: white; min-height: 35px;")
        btn_d.clicked.connect(dialogo.accept)
        lay_d.addWidget(btn_d)
        
        if dialogo.exec() == QDialog.DialogCode.Accepted:
            if txt_pass.text() == "1234":
                self.showNormal()
                self.stacked_widget.setCurrentIndex(0)
            else:
                QMessageBox.critical(self, "Error", "Clave incorrecta")

    def tecla_pad_presionada(self):
        btn = self.sender()
        texto_tecla = btn.text()
        texto_actual = self.txt_ingreso_id.text()
        if texto_tecla == "Limpiar": self.txt_ingreso_id.clear()
        elif texto_tecla == "Borrar": self.txt_ingreso_id.setText(texto_actual[:-1])
        else: self.txt_ingreso_id.setText(texto_actual + texto_tecla)

    def evaluar_obligatoriedad_carrera(self, tipo):
        if tipo == "Estudiante":
            self.lbl_alerta_carrera.setText("* Obligatorio")
            self.lbl_alerta_carrera.setStyleSheet("color: #EF4444; font-weight: bold;")
        else:
            self.lbl_alerta_carrera.setText("Opcional")
            self.lbl_alerta_carrera.setStyleSheet("color: #64748B;")

    def limpiar_volver_kiosco(self):
        self.txt_ingreso_id.clear()
        self.txt_reg_id.clear()
        self.txt_reg_nombre.clear()
        self.txt_reg_carrera.clear()
        self.cmb_reg_tipo.setCurrentIndex(0)
        self.kiosco_stack.setCurrentIndex(0)

    def procesar_ingreso_salon(self):
        ident = self.txt_ingreso_id.text().strip()
        if not ident: return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre_completo FROM usuarios WHERE identificacion = ?", (ident,))
        res = cursor.fetchone()
        if res:
            cursor.execute("INSERT INTO accesos_salon (usuario_id) VALUES (?)", (res[0],))
            conn.commit()
            QMessageBox.information(self, "Adelante", f"Bienvenido/a {res[1]}")
            self.limpiar_volver_kiosco()
        else:
            QMessageBox.warning(self, "Error", "ID no registrada. Vaya a Nuevo Usuario.")
        conn.close()

    def procesar_registro_usuario(self):
        ident = self.txt_reg_id.text().strip()
        nombre = self.txt_reg_nombre.text().strip()
        tipo = self.cmb_reg_tipo.currentText()
        carrera = self.txt_reg_carrera.text().strip()
        
        if not ident or not nombre: return
        if tipo == "Estudiante" and not carrera:
            QMessageBox.warning(self, "Aviso", "Indique su carrera.")
            return
            
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO usuarios (identificacion, nombre_completo, tipo_usuario, carrera) VALUES (?, ?, ?, ?)",
                (ident, nombre, tipo, carrera if tipo == "Estudiante" else None)
            )
            nuevo_id = cursor.lastrowid
            cursor.execute("INSERT INTO accesos_salon (usuario_id) VALUES (?)", (nuevo_id,))
            conn.commit()
            QMessageBox.information(self, "Éxito", f"¡Registro completado {nombre}!")
            self.limpiar_volver_kiosco()
        except sqlite3.IntegrityError:
            QMessageBox.critical(self, "Error", "Esa Cédula o Matrícula ya existe.")
        finally:
            conn.close()

    # -----------------------------------------------------------------
    # VISTA 3: MÉTRICAS Y EXCEL
    # -----------------------------------------------------------------
    def crear_vista_metricas(self):
        layout_global = QVBoxLayout(self.vista_metricas)
        layout_global.addLayout(self.barra_navegacion(2))
        
        lbl_m_titulo = QLabel("Panel de Estadísticas")
        lbl_m_titulo.setStyleSheet("font-size: 18px; font-weight: bold; color: #0F172A; margin-bottom: 10px;")
        layout_global.addWidget(lbl_m_titulo)
        
        form_m = QFormLayout()
        self.cmb_periodos = QComboBox()
        self.cmb_periodos.addItems(["Día actual", "Semana", "Último mes", "Último trimestre", "Cuatrimestre", "Último año"])
        self.cmb_periodos.setMinimumHeight(35)
        form_m.addRow("Seleccionar Rango:", self.cmb_periodos)
        layout_global.addLayout(form_m)
        
        btn_generar_excel = QPushButton("Generar Estadística en Excel")
        btn_generar_excel.setStyleSheet("background-color: #1E293B; color: white; font-size: 16px; font-weight: bold; min-height: 55px; border-radius: 8px;")
        btn_generar_excel.clicked.connect(self.exportar_metricas_excel)
        layout_global.addWidget(btn_generar_excel)
        layout_global.addStretch()

    def calcular_fechas(self):
        p = self.cmb_periodos.currentText()
        hoy = datetime.now()
        fin = hoy.strftime("%Y-%m-%d 23:59:59")
        
        if p == "Día actual": ini = hoy
        elif p == "Semana": ini = hoy - timedelta(days=7)
        elif p == "Último mes": ini = hoy - timedelta(days=30)
        elif p == "Último trimestre": ini = hoy - timedelta(days=90)
        elif p == "Cuatrimestre": ini = hoy - timedelta(days=120)
        elif p == "Último año": ini = hoy - timedelta(days=365)
        else: ini = hoy
            
        return ini.strftime("%Y-%m-%d 00:00:00"), fin

    def exportar_metricas_excel(self):
        f_inicio, f_fin = self.calcular_fechas()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        cursor.execute("SELECT nombre_pc FROM equipos ORDER BY id ASC")
        lista_pcs = [row[0] for row in cursor.fetchall()]
        
        frag_sql = [f"SUM(CASE WHEN e.nombre_pc = '{pc}' THEN 1 ELSE 0 END) AS [{pc}]" for pc in lista_pcs]
        cols_pcs = ", " + ", ".join(frag_sql) if frag_sql else ""
        
        query = f"""
        SELECT 
            u.nombre_completo AS [Nombre],
            u.identificacion AS [ID],
            u.tipo_usuario AS [Tipo],
            COALESCE(u.carrera, 'N/A') AS [Carrera],
            (SELECT COUNT(*) FROM accesos_salon a WHERE a.usuario_id = u.id AND datetime(a.fecha_hora) BETWEEN datetime(?) AND datetime(?)) AS [Veces Salón]
            {cols_pcs}
        FROM usuarios u
        LEFT JOIN prestamos_pc p ON u.id = p.usuario_id AND date(p.fecha_prestamo) BETWEEN date(?) AND date(?)
        LEFT JOIN equipos e ON p.equipo_id = e.id
        GROUP BY u.id
        ORDER BY u.nombre_completo ASC;
        """
        
        try:
            cursor.execute(query, (f_inicio, f_fin, f_inicio, f_fin))
            headers = [desc[0] for desc in cursor.description]
            datos = cursor.fetchall()
        finally:
            conn.close()
            
        wb = Workbook()
        ws = wb.active
        ws.title = "Métricas"
        
        f_head = Font(bold=True, color="FFFFFF")
        fill_head = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=text)
            cell.font = f_head
            cell.fill = fill_head
            
        for r_idx, fila in enumerate(datos, 2):
            for c_idx, val in enumerate(fila, 1):
                ws.cell(row=r_idx, column=c_idx, value=val)
                
        for col in ws.columns:
            l_max = max(len(str(c.value or '')) for c in col)
            ws.column_dimensions[col[0].column_letter].width = max(l_max + 2, 10)
            
        nombre = f"Reporte_Biblioteca_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(nombre)
        QMessageBox.information(self, "Excel", f"Guardado: {nombre}")

# =====================================================================
# INICIAR APLICACIÓN
# =====================================================================
if __name__ == "__main__":
    inicializar_bd()
    app = QApplication(sys.argv)
    
    # -----------------------------------------------------------------
    # SOLUCIÓN DE CONTEXTO CLARO ABSOLUTO (INMUNIDAD MODO OSCURO)
    # -----------------------------------------------------------------
    # Forzamos una paleta de colores clara nativa del software a la aplicación
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
    
    # Aplicamos las Hojas de Estilo (QSS) globales a nivel de la 'app'
    # Esto asegura que afecte también a sub-ventanas flotantes como QMessageBox, QDialog y vistas de ComboBox.
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