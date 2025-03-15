# Analizador de Eventos de Windows (GUI)

Herramienta en Python con interfaz gráfica (Tkinter) que permite analizar y filtrar eventos del Visor de Eventos de Windows de forma sencilla y exportarlos en formato CSV.

## 🚀 Características

- **Lectura de eventos** desde los registros locales de Windows (Application, System, Security).
- **Filtrado por niveles**: Error, Warning, Critical.
- **Filtrado por palabras clave**.
- **Filtrado por ID de evento específico o múltiples IDs**.
- **Exportación a CSV** para análisis externo.
- **Interfaz gráfica mejorada**, intuitiva y lista para uso operativo.

## ✅ Ejemplo de uso

1. Selecciona el tipo de log a analizar (Application, System, Security).
2. Opcional: Introduce palabras clave separadas por coma.
3. Opcional: Introduce IDs de eventos específicos separados por coma.
4. Selecciona los niveles de eventos que deseas consultar.
5. Selecciona o deja la carpeta de salida por defecto.
6. Haz clic en **"Analizar Eventos"** para generar el informe CSV.

## 📦 Requerimientos

- **Sistema operativo**: Windows 10/11, Windows Server (requiere acceso a los registros del sistema).
- **Python 3.8+**
- **Dependencias**:
  - `pywin32`
  - `tkinter` (incluido por defecto en la mayoría de distribuciones Python para Windows)
