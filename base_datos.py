import psycopg2

# Tu conexión real al servidor de Neon.tech
DB_URL = "postgresql://neondb_owner:npg_4owufargSh5c@ep-sparkling-cloud-ate9kxrl.c-9.us-east-1.aws.neon.tech/neondb?sslmode=require"

def obtener_conexion():
    return psycopg2.connect(DB_URL)

def inicializar_bd():
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Permitir que la cedula sea NULL para soportar estudiantes solo con matrícula
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id SERIAL PRIMARY KEY,
        cedula VARCHAR(30) UNIQUE,
        matricula VARCHAR(30) UNIQUE,
        nombre_completo VARCHAR(150) NOT NULL,
        tipo_usuario VARCHAR(50) NOT NULL,
        carrera VARCHAR(100)
    );
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipos (
        id SERIAL PRIMARY KEY,
        nombre_pc VARCHAR(50) UNIQUE NOT NULL,
        activo INTEGER DEFAULT 1
    );
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS accesos_salon (
        id SERIAL PRIMARY KEY,
        usuario_id INTEGER REFERENCES usuarios(id) ON DELETE CASCADE,
        fecha_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        biblioteca VARCHAR(100)
    );
    """)
    
    # Parches de estructura
    cursor.execute("ALTER TABLE accesos_salon ADD COLUMN IF NOT EXISTS biblioteca VARCHAR(100);")
    cursor.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS bloqueado BOOLEAN DEFAULT FALSE;")
    
    # Asegurar que cedula no sea NOT NULL si la tabla fue creada anteriormente
    cursor.execute("ALTER TABLE usuarios ALTER COLUMN cedula DROP NOT NULL;")
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS prestamos_pc (
        id SERIAL PRIMARY KEY,
        usuario_id INTEGER REFERENCES usuarios(id) ON DELETE CASCADE,
        equipo_id INTEGER REFERENCES equipos(id) ON DELETE CASCADE,
        fecha_prestamo DATE DEFAULT CURRENT_DATE
    );
    """)
    
    cursor.execute("SELECT COUNT(*) FROM equipos")
    if cursor.fetchone()[0] == 0:
        for i in range(1, 5):
            cursor.execute("INSERT INTO equipos (nombre_pc, activo) VALUES (%s, 1)", (f"PC {i}",))
            
    conn.commit()
    conn.close()

def obtener_pcs_activas():
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre_pc FROM equipos WHERE activo = 1")
    res = cursor.fetchall()
    conn.close()
    return res

def obtener_todas_pcs():
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre_pc, activo FROM equipos")
    res = cursor.fetchall()
    conn.close()
    return res

def buscar_usuario(ident):
    if not ident or not str(ident).strip():
        return None
    ident = str(ident).strip()
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, nombre_completo, tipo_usuario, carrera, COALESCE(bloqueado, FALSE) 
        FROM usuarios 
        WHERE (cedula IS NOT NULL AND cedula = %s) 
           OR (matricula IS NOT NULL AND matricula = %s)
    """, (ident, ident))
    res = cursor.fetchone()
    conn.close()
    return res

def agregar_equipo(nombre):
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO equipos (nombre_pc, activo) VALUES (%s, 1)", (nombre,))
    conn.commit()
    conn.close()

def actualizar_estado_equipo(id_pc, nuevo_estado):
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("UPDATE equipos SET activo = %s WHERE id = %s", (nuevo_estado, id_pc))
    conn.commit()
    conn.close()

def guardar_asignacion(usuario_id, id_pc):
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO prestamos_pc (usuario_id, equipo_id, fecha_prestamo) 
        VALUES (%s, %s, (CURRENT_TIMESTAMP AT TIME ZONE 'America/Santo_Domingo')::date)
    """, (usuario_id, id_pc))
    conn.commit()
    conn.close()

def registrar_acceso(usuario_id, biblioteca):
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO accesos_salon (usuario_id, fecha_hora, biblioteca) 
        VALUES (%s, CURRENT_TIMESTAMP AT TIME ZONE 'America/Santo_Domingo', %s)
    """, (usuario_id, biblioteca))
    conn.commit()
    conn.close()

def registrar_nuevo_usuario(cedula, matricula, nombre, tipo, carrera, biblioteca):
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Sanitizar cadenas vacías a NULL
    val_cedula = cedula.strip() if cedula and cedula.strip() else None
    val_matricula = matricula.strip() if matricula and matricula.strip() else None
    val_carrera = carrera.strip() if carrera and carrera.strip() and tipo == "Estudiante" else None

    # Validaciones previas de existencia
    if val_cedula:
        cursor.execute("SELECT id FROM usuarios WHERE cedula = %s", (val_cedula,))
        if cursor.fetchone():
            conn.close()
            raise Exception("Ya existe un usuario registrado con esta Cédula o ID.")
            
    if val_matricula:
        cursor.execute("SELECT id FROM usuarios WHERE matricula = %s", (val_matricula,))
        if cursor.fetchone():
            conn.close()
            raise Exception("Ya existe un usuario registrado con esta Matrícula.")

    cursor.execute(
        """INSERT INTO usuarios (cedula, matricula, nombre_completo, tipo_usuario, carrera) 
           VALUES (%s, %s, %s, %s, %s) RETURNING id""",
        (val_cedula, val_matricula, nombre.strip(), tipo, val_carrera)
    )
    nuevo_id = cursor.fetchone()[0]
    
    cursor.execute("""
        INSERT INTO accesos_salon (usuario_id, fecha_hora, biblioteca) 
        VALUES (%s, CURRENT_TIMESTAMP AT TIME ZONE 'America/Santo_Domingo', %s)
    """, (nuevo_id, biblioteca))
    conn.commit()
    conn.close()

def obtener_datos_metricas(f_inicio, f_fin):
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    cursor.execute("SELECT nombre_pc FROM equipos ORDER BY id ASC")
    lista_pcs = [row[0] for row in cursor.fetchall()]
    frag_sql = [f"SUM(CASE WHEN e.nombre_pc = '{pc}' THEN 1 ELSE 0 END) AS \"{pc}\"" for pc in lista_pcs]
    cols_pcs = ", " + ", ".join(frag_sql) if frag_sql else ""
    
    query = f"""
    SELECT 
        u.nombre_completo AS "Nombre",
        COALESCE(u.cedula, 'N/A') AS "Cédula",
        COALESCE(u.matricula, 'N/A') AS "Matrícula",
        u.tipo_usuario AS "Tipo",
        COALESCE(u.carrera, 'N/A') AS "Carrera",
        TO_CHAR(MAX(a.fecha_hora), 'YYYY-MM-DD') AS "Última Fecha",
        TO_CHAR(MAX(a.fecha_hora), 'HH12:MI:SS AM') AS "Última Hora",
        COALESCE(MAX(a.biblioteca), 'N/A') AS "Última Biblioteca",
        COUNT(a.id) AS "Veces Salón"
        {cols_pcs}
    FROM usuarios u
    LEFT JOIN accesos_salon a ON u.id = a.usuario_id AND a.fecha_hora BETWEEN %s::timestamp AND %s::timestamp
    LEFT JOIN prestamos_pc p ON u.id = p.usuario_id AND p.fecha_prestamo BETWEEN %s::date AND %s::date
    LEFT JOIN equipos e ON p.equipo_id = e.id
    GROUP BY u.id
    ORDER BY u.nombre_completo ASC;
    """
    
    cursor.execute(query, (f_inicio, f_fin, f_inicio, f_fin))
    headers = [desc[0] for desc in cursor.description]
    datos = cursor.fetchall()
    conn.close()
    return headers, datos

# ==========================================
# GESTIÓN DE USUARIOS
# ==========================================

def obtener_usuarios_gestion(filtro=""):
    conexion = obtener_conexion()
    cursor = conexion.cursor()
    if filtro and filtro.strip():
        termino = f"%{filtro.strip()}%"
        # Búsqueda por Cédula, Matrícula, Nombre y/o Apellidos de forma flexible
        query = """
            SELECT id, COALESCE(cedula, ''), COALESCE(matricula, ''), nombre_completo, tipo_usuario, COALESCE(carrera, ''), COALESCE(bloqueado, FALSE) 
            FROM usuarios 
            WHERE cedula ILIKE %s 
               OR matricula ILIKE %s 
               OR nombre_completo ILIKE %s
            ORDER BY id DESC
        """
        cursor.execute(query, (termino, termino, termino))
    else:
        query = "SELECT id, COALESCE(cedula, ''), COALESCE(matricula, ''), nombre_completo, tipo_usuario, COALESCE(carrera, ''), COALESCE(bloqueado, FALSE) FROM usuarios ORDER BY id DESC"
        cursor.execute(query)
    res = cursor.fetchall()
    conexion.close()
    return res

def eliminar_usuario(id_usuario):
    conexion = obtener_conexion()
    cursor = conexion.cursor()
    cursor.execute("DELETE FROM usuarios WHERE id = %s", (id_usuario,))
    conexion.commit()
    conexion.close()

def actualizar_usuario(id_usuario, cedula, matricula, nombre, tipo, carrera):
    conexion = obtener_conexion()
    cursor = conexion.cursor()
    
    # Mapear strings vacíos a NULL para evitar colisiones UNIQUE
    val_cedula = cedula.strip() if cedula and cedula.strip() else None
    val_matricula = matricula.strip() if matricula and matricula.strip() else None
    val_carrera = carrera.strip() if carrera and carrera.strip() else None

    # Validar duplicados omitiendo los valores nulos
    if val_cedula:
        cursor.execute("SELECT id FROM usuarios WHERE cedula = %s AND id != %s", (val_cedula, id_usuario))
        if cursor.fetchone():
            conexion.close()
            raise Exception("Esta Cédula o ID ya pertenece a otro usuario.")

    if val_matricula:
        cursor.execute("SELECT id FROM usuarios WHERE matricula = %s AND id != %s", (val_matricula, id_usuario))
        if cursor.fetchone():
            conexion.close()
            raise Exception("Esta Matrícula ya pertenece a otro usuario.")

    cursor.execute("""
        UPDATE usuarios 
        SET cedula = %s, matricula = %s, nombre_completo = %s, tipo_usuario = %s, carrera = %s 
        WHERE id = %s
    """, (val_cedula, val_matricula, nombre.strip(), tipo, val_carrera, id_usuario))
    
    conexion.commit()
    conexion.close()

def alternar_bloqueo_usuario(id_usuario, estado_actual):
    conexion = obtener_conexion()
    cursor = conexion.cursor()
    nuevo_estado = not estado_actual
    cursor.execute("UPDATE usuarios SET bloqueado = %s WHERE id = %s", (nuevo_estado, id_usuario))
    conexion.commit()
    conexion.close()