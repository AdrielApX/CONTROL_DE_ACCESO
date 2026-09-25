# Sistema de Control de Acceso

Un sistema de control de acceso y gestión desarrollado para optimizar los procesos de una biblioteca. Este proyecto está diseñado para funcionar como una aplicación de escritorio conectada a una infraestructura de base de datos en la nube, garantizando la seguridad, persistencia y disponibilidad de los datos.

## Tecnologías y Arquitectura

### Lenguaje Principal
* **Python 3.13:** Lenguaje núcleo del proyecto, elegido por su versatilidad, eficiencia en el manejo de datos y amplio ecosistema de herramientas.

### Base de Datos en la Nube
* **PostgreSQL:** Sistema de gestión de bases de datos relacional utilizado para almacenar de forma estructurada los registros de accesos a usuarios de la biblioteca.
* **Alojamiento Cloud:** La base de datos está desplegada en la nube, lo que permite la centralización de la información y el acceso en tiempo real sin depender de un servidor local de base de datos.

### Librerías y Dependencias Clave
* **Controladores de Base de Datos (ej. `psycopg2` o `SQLAlchemy`):** Utilizados para establecer la conexión segura y ejecutar transacciones SQL entre la aplicación de Python y PostgreSQL en la nube.
* **Interfaz Gráfica (GUI):** Implementada para proporcionar una experiencia de usuario intuitiva al personal de la biblioteca.
* **PyInstaller:** Herramienta de empaquetado empleada para compilar el código fuente y sus dependencias en un único archivo ejecutable (`.exe`), aislando al usuario final de la necesidad de instalar Python o configurar el entorno.

## Entorno de Desarrollo

El sistema fue construido y optimizado en el siguiente entorno:
* **Sistema Operativo:** Windows.
* **IDE (Entorno de Desarrollo Integrado):** Visual Studio Code.
* **Control de Versiones y Repositorio:** Git y GitHub, manejando un flujo de trabajo estructurado mediante ramas (`main`, `qa`, `dev`).
* **Gestión de Entorno:** Entorno virtual nativo de Python (`venv`) para el aislamiento de dependencias (`pip`).

## Despliegue y Ejecución

El proyecto está diseñado para distribuirse como un ejecutable independiente. La comunicación con la base de datos se realiza a través de credenciales de red, por lo que el equipo cliente únicamente requiere:
1. Conexión a internet estable (para la sincronización con PostgreSQL).
2. Ejecutar el archivo generado en la carpeta `dist/` tras la compilación.

Para entornos de desarrollo, el levantamiento requiere activar el entorno virtual y asegurar la configuración de las variables de entorno de la base de datos antes de compilar con:
```bash
pyinstaller --clean app.spec
