import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Checkbutton, Separator
import csv
import os
import win32evtlog


# Función para leer logs de eventos con filtro por ID de evento
def read_event_logs(server_name="localhost", log_type="Application", levels=None, keywords=None, event_ids=None):
    try:
        handle = win32evtlog.OpenEventLog(server_name, log_type)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo conectar al equipo '{server_name}' o al log '{log_type}'.\nDetalles: {str(e)}")
        return []

    flags = win32evtlog.EVENTLOG_FORWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    events = []
    levels_map = {"ERROR": 2, "WARNING": 3, "CRITICAL": 1}

    while True:
        records = win32evtlog.ReadEventLog(handle, flags, 0)
        if not records:
            break
        for record in records:
            if levels and record.EventType not in [levels_map[lvl] for lvl in levels]:
                continue

            # Filtrar por ID de evento
            event_id = record.EventID & 0xFFFF
            if event_ids and event_id not in event_ids:
                continue

            # Filtrar por palabras clave
            message = record.StringInserts
            combined_message = " ".join(message) if message else ""
            if keywords and not any(kw.lower() in combined_message.lower() for kw in keywords):
                continue

            events.append({
                "DateTime": record.TimeGenerated.Format(),
                "Source": record.SourceName,
                "EventID": event_id,
                "Category": record.EventCategory,
                "Message": combined_message if combined_message else "No message"
            })
    win32evtlog.CloseEventLog(handle)
    return events


# Generar CSV
def generate_csv_report(events, output_path):
    with open(output_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=["DateTime", "Source", "EventID", "Category", "Message"])
        writer.writeheader()
        for event in events:
            writer.writerow(event)


# Seleccionar carpeta
def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_path_var.set(folder)


# Análisis completo
def analyze_logs():
    server_name = server_name_entry.get().strip() or "localhost"
    log_type = log_type_var.get()
    keywords = [kw.strip() for kw in keywords_entry.get().split(",") if kw.strip()]
    output_folder = folder_path_var.get()
    levels = [level for level, var in level_vars.items() if var.get()]
    event_ids = [int(eid.strip()) for eid in event_id_entry.get().split(",") if eid.strip().isdigit()]

    if not log_type:
        messagebox.showerror("Error", "Selecciona un tipo de log.")
        return
    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Documents")
        folder_path_var.set(output_folder)
    if not levels and not event_ids:
        messagebox.showerror("Error", "Selecciona al menos un nivel de filtro o un ID de evento.")
        return

    events = read_event_logs(server_name, log_type, levels, keywords, event_ids)
    if not events:
        messagebox.showinfo("Resultado", f"No se encontraron eventos con los filtros seleccionados en '{server_name}'.")
        return

    output_path = os.path.join(output_folder, f"event_report_{server_name}_{log_type}.csv")
    generate_csv_report(events, output_path)
    messagebox.showinfo("Completado", f"Informe generado exitosamente en:\n{output_path}")


# ---------- GUI Mejorada ---------- #
root = tk.Tk()
root.title("Analizador Avanzado de Eventos de Windows")
root.geometry("600x400")  # Ajuste de tamaño

# Variables
log_type_var = tk.StringVar(value="Application")
folder_path_var = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Documents"))
level_vars = {level: tk.BooleanVar() for level in ["ERROR", "WARNING", "CRITICAL"]}

# Estilo títulos de sección
def section_title(text, row):
    tk.Label(root, text=text, font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=3, sticky="w", padx=10, pady=5)

# Sección: Conexión
section_title("Conexión al Equipo:", 0)
tk.Label(root, text="Nombre de Equipo:").grid(row=1, column=0, padx=10, sticky="w")
server_name_entry = tk.Entry(root, width=50)
server_name_entry.grid(row=1, column=1, columnspan=2, padx=10, sticky="w")
server_name_entry.insert(0, "localhost")

# Sección: Parámetros de Búsqueda
Separator(root, orient='horizontal').grid(row=2, column=0, columnspan=3, sticky="ew", pady=5)
section_title("Parámetros de Búsqueda:", 3)

tk.Label(root, text="Tipo de Log:").grid(row=4, column=0, padx=10, sticky="w")
Combobox(root, textvariable=log_type_var, values=["Application", "System", "Security"]).grid(row=4, column=1, padx=10, sticky="w")

tk.Label(root, text="IDs de Evento (separados por coma):").grid(row=5, column=0, padx=10, sticky="w")
event_id_entry = tk.Entry(root, width=50)
event_id_entry.grid(row=5, column=1, columnspan=2, padx=10, sticky="w")

tk.Label(root, text="Palabras Clave (separadas por coma):").grid(row=6, column=0, padx=10, sticky="w")
keywords_entry = tk.Entry(root, width=50)
keywords_entry.grid(row=6, column=1, columnspan=2, padx=10, sticky="w")

# Filtros de Nivel
tk.Label(root, text="Niveles:").grid(row=7, column=0, padx=10, sticky="w")
filter_frame = tk.Frame(root)
filter_frame.grid(row=7, column=1, columnspan=2, padx=10, pady=5, sticky="w")
for level, var in level_vars.items():
    Checkbutton(filter_frame, text=level, variable=var).pack(side="left", padx=5)

# Sección: Salida
Separator(root, orient='horizontal').grid(row=8, column=0, columnspan=3, sticky="ew", pady=5)
section_title("Carpeta de Salida:", 9)

tk.Entry(root, textvariable=folder_path_var, width=50).grid(row=10, column=0, columnspan=2, padx=10, pady=5, sticky="w")
tk.Button(root, text="Seleccionar", command=select_folder).grid(row=10, column=2, padx=10, pady=5, sticky="w")

# Botón principal
Separator(root, orient='horizontal').grid(row=11, column=0, columnspan=3, sticky="ew", pady=10)
tk.Button(root, text="Analizar Eventos", command=analyze_logs, bg="green", fg="white", font=("Arial", 11, "bold")).grid(row=12, column=0, columnspan=3, pady=15)

root.mainloop()
