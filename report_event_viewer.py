# Importar Módulos Necesarios
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Checkbutton
import csv
import os
import win32evtlog

# Definir funciones
def read_event_logs(server_name="localhost", log_type="Application", levels=None, keywords=None):
    """
    Lee los logs de eventos del Visor de Eventos de un equipo (local o remoto).
    
    :param server_name: Nombre del equipo (por defecto 'localhost').
    :param log_type: Tipo de log ('Application', 'System', 'Security', etc.).
    :param levels: Lista de niveles de filtro (ERROR, WARNING, CRITICAL).
    :param keywords: Lista de palabras clave para filtrar mensajes.
    :return: Lista de eventos.
    """
    try:
        handle = win32evtlog.OpenEventLog(server_name, log_type)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo conectar al equipo {server_name}. Verifica el nombre del equipo y permisos.\n{str(e)}")
        return []

    flags = win32evtlog.EVENTLOG_FORWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    events = []
    levels_map = {"ERROR": 2, "WARNING": 3, "CRITICAL": 1}  # Mapa de niveles

    while True:
        records = win32evtlog.ReadEventLog(handle, flags, 0)
        if not records:
            break
        for record in records:
            # Filtro por niveles
            if levels and record.EventType not in [levels_map[lvl] for lvl in levels]:
                continue
            # Filtro por palabras clave
            message = record.StringInserts
            if keywords and not any(kw.lower() in str(message).lower() for kw in keywords):
                continue
            # Formateo del evento para el informe
            events.append({
                "DateTime": record.TimeGenerated.Format(),
                "Source": record.SourceName,
                "EventID": record.EventID,
                "Category": record.EventCategory,
                "Message": " | ".join(message) if message else "No message"
            })
    win32evtlog.CloseEventLog(handle)
    return events

# Generar Informe CSV
def generate_csv_report(events, output_path):
    with open(output_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=["DateTime", "Source", "EventID", "Category", "Message"])
        writer.writeheader()  # Escribir la cabecera del archivo
        for event in events:
            writer.writerow(event)  # Escribir cada evento

# Seleccionar Carpeta
def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_path_var.set(folder)  # Guardar la ruta seleccionada

# Añadir Lógica Análisis
def analyze_logs():
    server_name = server_name_entry.get().strip() or "localhost"
    log_type = log_type_var.get()
    keywords = keywords_entry.get().split(",") if keywords_entry.get() else None
    output_folder = folder_path_var.get()
    levels = [level for level, var in level_vars.items() if var.get()]

    # Validaciones
    if not server_name:
        messagebox.showerror("Error", "Especifica un nombre de equipo.")
        return
    if not log_type:
        messagebox.showerror("Error", "Selecciona un tipo de log.")
        return
    if not output_folder:  # Si no carpeta, usar Documentos
        output_folder = os.path.join(os.path.expanduser("~"), "Documents")
        folder_path_var.set(output_folder)
    if not levels:
        messagebox.showerror("Error", "Selecciona al menos un nivel de filtro.")
        return

    try:
        events = read_event_logs(server_name, log_type, levels, keywords)
        if not events:
            messagebox.showinfo("Resultado", f"No se encontraron eventos en el equipo '{server_name}' con los filtros seleccionados.")
            return
        output_path = os.path.join(output_folder, f"event_report_{server_name}.csv")
        generate_csv_report(events, output_path)
        messagebox.showinfo("Completado", f"Informe generado exitosamente en:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

# Crear Interfaz
root = tk.Tk()
root.title("Analizador de Eventos de Windows")

# Definir Variables Controles:
log_type_var = tk.StringVar()  # Tipo de log
folder_path_var = tk.StringVar()  # Carpeta de salida
folder_path_var.set(os.path.join(os.path.expanduser("~"), "Documents"))  # Carpeta predeterminada
level_vars = {  # Niveles de filtro
    "ERROR": tk.BooleanVar(),
    "WARNING": tk.BooleanVar(),
    "CRITICAL": tk.BooleanVar()
}

# Añadir Controles Ventana:
# Nombre de equipo
tk.Label(root, text="Nombre de Equipo:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
server_name_entry = tk.Entry(root, width=40)
server_name_entry.grid(row=0, column=1, columnspan=2, padx=10, pady=5, sticky="w")
server_name_entry.insert(0, "localhost")  # Valor predeterminado

# Selección del tipo de log
tk.Label(root, text="Tipo de Log:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
log_type_dropdown = Combobox(root, textvariable=log_type_var, values=["Application", "System", "Security"])
log_type_dropdown.grid(row=1, column=1, padx=10, pady=5, sticky="w")

# Filtros de nivel
tk.Label(root, text="Niveles de Filtro:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
filter_frame = tk.Frame(root)
filter_frame.grid(row=2, column=1, columnspan=2, padx=10, pady=5, sticky="w")
for level, var in level_vars.items():
    Checkbutton(filter_frame, text=level, variable=var).pack(side="left", padx=5)

# Palabras clave
tk.Label(root, text="Palabras Clave (opcional):").grid(row=3, column=0, padx=10, pady=5, sticky="w")
keywords_entry = tk.Entry(root, width=40)
keywords_entry.grid(row=3, column=1, columnspan=2, padx=10, pady=5, sticky="w")

# Carpeta de salida
tk.Label(root, text="Carpeta de Salida:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
folder_entry = tk.Entry(root, textvariable=folder_path_var, width=40)
folder_entry.grid(row=4, column=1, padx=10, pady=5, sticky="w")
tk.Button(root, text="Seleccionar", command=select_folder).grid(row=4, column=2, padx=10, pady=5, sticky="w")

# Botón para analizar eventos
tk.Button(root, text="Analizar Eventos", command=analyze_logs, bg="blue", fg="white").grid(row=5, column=0, columnspan=3, pady=10)

log_type_var.set("Application")

root.mainloop()
