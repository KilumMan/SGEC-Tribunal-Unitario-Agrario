import sqlite3
from datetime import datetime

def conectar():
    """Conecta a la base de datos."""
    conn = sqlite3.connect('inventario.db')
    return conn

def crear_tablas():
    """Crea la tabla de equipos y movimientos si no existen."""
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Tabla de equipos (sin cambios)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS equipos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                asignado_a TEXT NOT NULL,
                tipo_equipo TEXT NOT NULL,
                num_inventario TEXT NOT NULL,
                marca TEXT,
                modelo TEXT,
                num_serie TEXT UNIQUE,
                estado TEXT NOT NULL,
                descripcion_estado TEXT,
                ultima_actualizacion TIMESTAMP
            )
        """)
        # Nueva tabla para registrar todos los movimientos
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS movimientos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                equipo_info TEXT NOT NULL,
                accion TEXT NOT NULL,
                fecha TIMESTAMP
            )
        """)
        conn.commit()
    except sqlite3.Error as e:
        print(f"Error al crear tablas: {e}")
    finally:
        if conn:
            conn.close()

def registrar_movimiento(equipo_info, accion):
    """Inserta un nuevo registro en la tabla de movimientos."""
    try:
        conn = conectar()
        cursor = conn.cursor()
        sql = "INSERT INTO movimientos (equipo_info, accion, fecha) VALUES (?, ?, ?)"
        cursor.execute(sql, (equipo_info, accion, datetime.now()))
        conn.commit()
    except sqlite3.Error as e:
        print(f"Error al registrar movimiento: {e}")
    finally:
        if conn:
            conn.close()

def agregar_equipo(datos):
    """Agrega un nuevo equipo y registra el movimiento."""
    datos.append(datetime.now())
    try:
        conn = conectar()
        cursor = conn.cursor()
        sql = """
            INSERT INTO equipos (asignado_a, tipo_equipo, num_inventario, marca, modelo, num_serie, estado, descripcion_estado, ultima_actualizacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        cursor.execute(sql, datos)
        conn.commit()
        # Registrar el movimiento
        info = f"Equipo '{datos[1]}' asignado a '{datos[0]}'"
        registrar_movimiento(info, "Agregado")
    except sqlite3.Error as e:
        print(f"Error al agregar equipo: {e}")
        return False
    finally:
        if conn:
            conn.close()
    return True

def actualizar_equipo(id_equipo, datos):
    """Actualiza un equipo y registra el movimiento."""
    datos.append(datetime.now())
    datos.append(id_equipo)
    try:
        conn = conectar()
        cursor = conn.cursor()
        sql = """
            UPDATE equipos
            SET asignado_a = ?, tipo_equipo = ?, num_inventario = ?, marca = ?, modelo = ?,
                num_serie = ?, estado = ?, descripcion_estado = ?, ultima_actualizacion = ?
            WHERE id = ?
        """
        cursor.execute(sql, datos)
        conn.commit()
        # Registrar el movimiento
        info = f"Equipo '{datos[1]}' asignado a '{datos[0]}'"
        registrar_movimiento(info, "Actualizado")
    except sqlite3.Error as e:
        print(f"Error al actualizar equipo: {e}")
        return False
    finally:
        if conn:
            conn.close()
    return True

def eliminar_equipo(id_equipo):
    """Elimina un equipo y registra el movimiento."""
    try:
        conn = conectar()
        cursor = conn.cursor()
        # 1. Obtener la info del equipo ANTES de borrarlo
        cursor.execute("SELECT tipo_equipo, asignado_a FROM equipos WHERE id = ?", (id_equipo,))
        equipo_a_borrar = cursor.fetchone()
        
        if equipo_a_borrar:
            # 2. Eliminar el equipo
            cursor.execute("DELETE FROM equipos WHERE id = ?", (id_equipo,))
            conn.commit()
            
            # 3. Registrar el movimiento de eliminación
            info = f"Equipo '{equipo_a_borrar[0]}' que estaba asignado a '{equipo_a_borrar[1]}'"
            registrar_movimiento(info, "Eliminado")
            return True
        else:
            return False # No se encontró el equipo
            
    except sqlite3.Error as e:
        print(f"Error al eliminar equipo: {e}")
        return False
    finally:
        if conn:
            conn.close()

def obtener_todos_los_equipos():
    """Obtiene todos los equipos de la base de datos."""
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, asignado_a, tipo_equipo, num_inventario, marca, modelo, num_serie, estado FROM equipos ORDER BY id DESC")
        return cursor.fetchall()
    except sqlite3.Error as e:
        print(f"Error al obtener equipos: {e}")
        return []
    finally:
        if conn:
            conn.close()

def obtener_movimientos_recientes(limite=5):
    """Obtiene los últimos movimientos registrados (agregar, actualizar, eliminar)."""
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT equipo_info, accion, fecha FROM movimientos ORDER BY fecha DESC LIMIT ?", (limite,))
        return cursor.fetchall()
    except sqlite3.Error as e:
        print(f"Error al obtener movimientos recientes: {e}")
        return []
    finally:
        if conn:
            conn.close()

def obtener_equipo_por_id(id_equipo):
    """Obtiene toda la información de un equipo por su ID."""
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM equipos WHERE id = ?", (id_equipo,))
        return cursor.fetchone()
    except sqlite3.Error as e:
        print(f"Error al obtener equipo por ID: {e}")
        return None
    finally:
        if conn:
            conn.close()

# Al iniciar, se asegura que las tablas estén creadas
crear_tablas()