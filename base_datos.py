import psycopg2

# Tu conexión real al servidor de Neon.tech
DB_URL = "postgresql://neondb_owner:npg_4owufargSh5c@ep-sparkling-cloud-ate9kxrl.c-9.us-east-1.aws.neon.tech/neondb?sslmode=require"

def obtener_conexion():
    return psycopg2.connect(DB_URL)

def inicializar_bd():
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id SERIAL PRIMARY KEY,
        cedula VARCHAR(20) UNIQUE NOT NULL,
        matricula VARCHAR(20) UNIQUE,
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
        fecha_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    
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
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre_completo, tipo_usuario, carrera FROM usuarios WHERE cedula = %s OR matricula = %s", (ident, ident))
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
    # Forzamos la zona horaria de RD para la fecha
    cursor.execute("""
        INSERT INTO prestamos_pc (usuario_id, equipo_id, fecha_prestamo) 
        VALUES (%s, %s, (CURRENT_TIMESTAMP AT TIME ZONE 'America/Santo_Domingo')::date)
    """, (usuario_id, id_pc))
    conn.commit()
    conn.close()

def registrar_acceso(usuario_id):
    conn = obtener_conexion()
    cursor = conn.cursor()
    # Forzamos la zona horaria de RD para la fecha y hora
    cursor.execute("""
        INSERT INTO accesos_salon (usuario_id, fecha_hora) 
        VALUES (%s, CURRENT_TIMESTAMP AT TIME ZONE 'America/Santo_Domingo')
    """, (usuario_id,))
    conn.commit()
    conn.close()

def registrar_nuevo_usuario(cedula, matricula, nombre, tipo, carrera):
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute(
        """INSERT INTO usuarios (cedula, matricula, nombre_completo, tipo_usuario, carrera) 
           VALUES (%s, %s, %s, %s, %s) RETURNING id""",
        (cedula, matricula if matricula else None, nombre, tipo, carrera if tipo == "Estudiante" else None)
    )
    nuevo_id = cursor.fetchone()[0]
    
    # Forzamos la zona horaria de RD al crear el primer acceso
    cursor.execute("""
        INSERT INTO accesos_salon (usuario_id, fecha_hora) 
        VALUES (%s, CURRENT_TIMESTAMP AT TIME ZONE 'America/Santo_Domingo')
    """, (nuevo_id,))
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
        u.cedula AS "Cédula",
        COALESCE(u.matricula, 'N/A') AS "Matrícula",
        u.tipo_usuario AS "Tipo",
        COALESCE(u.carrera, 'N/A') AS "Carrera",
        TO_CHAR(MAX(a.fecha_hora), 'YYYY-MM-DD') AS "Última Fecha",
        TO_CHAR(MAX(a.fecha_hora), 'HH12:MI:SS AM') AS "Última Hora",
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