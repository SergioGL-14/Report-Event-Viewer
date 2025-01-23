# **Analizador de Logs de Windows**

## Descripción
El **Analizador de Logs de Windows** es una herramienta desarrollada en Python que permite analizar los eventos del Visor de Eventos de Windows tanto de manera local como remota. Incluye múltiples opciones de filtrado y genera informes detallados en formato CSV.

## Características
- **Soporte para equipos locales y remotos**: 
  Especifica el nombre de un equipo en la red para analizar sus logs, siempre que tengas los permisos necesarios.
- **Filtros avanzados**:
  - Niveles de severidad: `ERROR`, `WARNING`, `CRITICAL`.
  - Filtrado por palabras clave.
  - Filtrado por ID específico de evento.
- **Tipos de logs compatibles**:
  - `Application`
  - `System`
  - `Security`
- **Generación de reportes**:
  - Los resultados se exportan automáticamente en un archivo `.csv`.

## Requisitos
1. **Python 3.9 o superior** instalado en tu sistema.
2. **Sistema Operativo**: Windows.
3. **Dependencias de Python**:
   Instala la biblioteca necesaria ejecutando:
   ```bash
   pip install pywin32
