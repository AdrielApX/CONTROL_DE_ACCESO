import sqlite3

DB_NAME = "biblioteca_utesa.db"

# =====================================================================
# 1. BASE DE DATOS E INICIALIZACIÓN
# =====================================================================
def inicializar_bd():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    
    cursor.execute("CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, cedula TEXT UNIQUE NOT NULL, matricula TEXT UNIQUE, nombre_completo TEXT NOT NULL, tipo_usuario TEXT NOT NULL, carrera TEXT);")
    cursor.execute("CREATE TABLE IF NOT EXISTS equipos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre_pc TEXT UNIQUE NOT NULL, activo INTEGER DEFAULT 1);")
    cursor.execute("CREATE TABLE IF NOT EXISTS accesos_salon (id INTEGER PRIMARY KEY AUTOINCREMENT, usuario_id INTEGER, fecha_hora TEXT DEFAULT (datetime('now', 'localtime')), FOREIGN KEY(usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE);")
    cursor.execute("CREATE TABLE IF NOT EXISTS prestamos_pc (id INTEGER PRIMARY KEY AUTOINCREMENT, usuario_id INTEGER, equipo_id INTEGER, fecha_prestamo TEXT DEFAULT (date('now', 'localtime')), FOREIGN KEY(usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE, FOREIGN KEY(equipo_id) REFERENCES equipos(id) ON DELETE CASCADE);")
    
    cursor.execute("SELECT COUNT(*) FROM equipos")
    if cursor.fetchone()[0] == 0:
        for i in range(1, 5):
            cursor.execute("INSERT INTO equipos (nombre_pc, activo) VALUES (?, 1)", (f"PC {i}",))
            
    conn.commit()
    conn.close()

# ---- FUNCIONES QUE ALIMENTAN A LA INTERFAZ ----

def obtener_pcs_activas():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre_pc FROM equipos WHERE activo = 1")
    res = cursor.fetchall()
    conn.close()
    return res

def obtener_todas_pcs():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre_pc, activo FROM equipos")
    res = cursor.fetchall()
    conn.close()
    return res

def buscar_usuario(ident):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre_completo, tipo_usuario, carrera FROM usuarios WHERE cedula = ? OR matricula = ?", (ident, ident))
    res = cursor.fetchone()
    conn.close()
    return res

def agregar_equipo(nombre):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO equipos (nombre_pc, activo) VALUES (?, 1)", (nombre,))
    conn.commit()
    conn.close()

def actualizar_estado_equipo(id_pc, nuevo_estado):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("UPDATE equipos SET activo = ? WHERE id = ?", (nuevo_estado, id_pc))
    conn.commit()
    conn.close()

def guardar_asignacion(usuario_id, id_pc):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO prestamos_pc (usuario_id, equipo_id) VALUES (?, ?)", (usuario_id, id_pc))
    conn.commit()
    conn.close()

def registrar_acceso(usuario_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO accesos_salon (usuario_id) VALUES (?)", (usuario_id,))
    conn.commit()
    conn.close()

def registrar_nuevo_usuario(cedula, matricula, nombre, tipo, carrera):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO usuarios (cedula, matricula, nombre_completo, tipo_usuario, carrera) VALUES (?, ?, ?, ?, ?)",
        (cedula, matricula if matricula else None, nombre, tipo, carrera if tipo == "Estudiante" else None)
    )
    nuevo_id = cursor.lastrowid
    cursor.execute("INSERT INTO accesos_salon (usuario_id) VALUES (?)", (nuevo_id,))
    conn.commit()
    conn.close()

def obtener_datos_metricas(f_inicio, f_fin):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute("SELECT nombre_pc FROM equipos ORDER BY id ASC")
    lista_pcs = [row[0] for row in cursor.fetchall()]
    frag_sql = [f"SUM(CASE WHEN e.nombre_pc = '{pc}' THEN 1 ELSE 0 END) AS [{pc}]" for pc in lista_pcs]
    cols_pcs = ", " + ", ".join(frag_sql) if frag_sql else ""
    
    query = f"""
    SELECT 
        u.nombre_completo AS [Nombre],
        u.cedula AS [Cédula],
        COALESCE(u.matricula, 'N/A') AS [Matrícula],
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
    cursor.execute(query, (f_inicio, f_fin, f_inicio, f_fin))
    headers = [desc[0] for desc in cursor.description]
    datos = cursor.fetchall()
    conn.close()
    return headers, datos