import os
import re
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk

try:
    import win32evtlog
except ImportError:
    messagebox.showerror("Error de Dependencia", "pywin32 es requerido para acceder al Visor de Eventos de Windows. Instala 'pywin32'.")
    raise

# =========================
# Funciones de utilidad
# =========================

def sanitize_filename(name: str) -> str:
    """
    Sustituye caracteres inválidos en nombres de archivo por guiones bajos.
    """
    return re.sub(r'[\\/*?"<>|:]+', '_', name)


def read_event_logs(server: str = "localhost",
                    log_type: str = "Application",
                    levels: list = None,
                    keywords: list = None,
                    event_ids: list = None) -> list:
    """
    Lee eventos del Visor de Eventos de Windows con filtros opcionales.
    :returns: Lista de dicts con campos DateTime, Source, EventID, Category, Message
    """
    try:
        handle = win32evtlog.OpenEventLog(server, log_type)
    except Exception as e:
        raise RuntimeError(f"No se pudo abrir log '{log_type}' en '{server}': {e}")

    flags = win32evtlog.EVENTLOG_FORWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    levels_map = {"ERROR": win32evtlog.EVENTLOG_ERROR_TYPE,
                  "WARNING": win32evtlog.EVENTLOG_WARNING_TYPE,
                  "INFORMATION": win32evtlog.EVENTLOG_INFORMATION_TYPE}
    events = []

    try:
        while True:
            records = win32evtlog.ReadEventLog(handle, flags, 0)
            if not records:
                break
            for rec in records:
                evt_type = rec.EventType
                if levels and evt_type not in [levels_map[l] for l in levels if l in levels_map]:
                    continue

                eid = rec.EventID & 0xFFFF
                if event_ids and eid not in event_ids:
                    continue

                inserts = rec.StringInserts or []
                message = " ".join(inserts) or "No message"
                if keywords and not any(kw.lower() in message.lower() for kw in keywords):
                    continue

                events.append({
                    "DateTime": rec.TimeGenerated.Format(),
                    "Source": rec.SourceName,
                    "EventID": eid,
                    "Category": rec.EventCategory,
                    "Message": message
                })
    except Exception as e:
        raise RuntimeError(f"Error leyendo registros: {e}")
    finally:
        win32evtlog.CloseEventLog(handle)

    return events


def generate_csv_report(events: list, output_path: str):
    """
    Exporta los eventos a CSV en la ruta indicada.
    """
    if not events:
        raise ValueError("No hay eventos para exportar.")

    fieldnames = ["DateTime", "Source", "EventID", "Category", "Message"]
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for evt in events:
                writer.writerow(evt)
    except Exception as e:
        raise IOError(f"Error exportando CSV: {e}")

# =========================
# Clase de la Aplicación
# =========================

class EventAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Analizador de Eventos de Windows")
        self.geometry("800x700")
        self.configure(bg="#F5F5F5")
        self._setup_style()
        self._create_widgets()

    def _setup_style(self):
        style = ttk.Style(self)
        style.theme_use('default')
        # General Labels & Frames
        style.configure('TLabel', background='#F5F5F5', foreground='#333333', font=('Segoe UI', 9))
        style.configure('TFrame', background='#F5F5F5')
        style.configure('TLabelframe', background='#F5F5F5', foreground='#333333')
        style.configure('TLabelframe.Label', background='#F5F5F5', foreground='#333333')
        # Treeview light style
        style.configure('Treeview',
                        background='#FFFFFF',
                        fieldbackground='#FFFFFF',
                        foreground='#333333',
                        rowheight=25)
        style.map('Treeview', background=[('selected', '#0078D7')], foreground=[('selected', '#FFFFFF')])
        style.configure('Treeview.Heading', background='#E1E1E1', foreground='#333333', font=('Segoe UI', 9, 'bold'))

    def _create_widgets(self):
        # ------------------- Configuración -------------------
        cfg = ttk.LabelFrame(self, text="Parámetros de Búsqueda", padding=(10,8))
        cfg.pack(fill='x', padx=15, pady=10)

        # Servidor y Tipo de Log
        ttk.Label(cfg, text="Equipo:").grid(row=0, column=0, sticky='w')
        self.server_entry = ttk.Entry(cfg)
        self.server_entry.insert(0, 'localhost')
        self.server_entry.grid(row=0, column=1, sticky='w', padx=5)

        ttk.Label(cfg, text="Tipo de Log:").grid(row=0, column=2, sticky='w', padx=(20,0))
        self.log_type = ttk.Combobox(cfg, values=["Application", "System", "Security"], state='readonly', width=15)
        self.log_type.current(0)
        self.log_type.grid(row=0, column=3, sticky='w', padx=5)

        # IDs y Keywords
        ttk.Label(cfg, text="IDs (coma):").grid(row=1, column=0, sticky='w', pady=8)
        self.event_id_entry = ttk.Entry(cfg, width=30)
        self.event_id_entry.grid(row=1, column=1, sticky='w')

        ttk.Label(cfg, text="Keywords (coma):").grid(row=1, column=2, sticky='w', padx=(20,0))
        self.keywords_entry = ttk.Entry(cfg, width=30)
        self.keywords_entry.grid(row=1, column=3, sticky='w')

        # Niveles
        ttk.Label(cfg, text="Niveles:").grid(row=2, column=0, sticky='w')
        levels = ["ERROR", "WARNING", "INFORMATION"]
        self.level_vars = {}
        lv_frame = ttk.Frame(cfg)
        lv_frame.grid(row=2, column=1, columnspan=3, sticky='w', pady=8)
        for lvl in levels:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(lv_frame, text=lvl, variable=var, bg="#F5F5F5", fg="#333333",
                                selectcolor="#E1E1E1", activebackground="#F5F5F5")
            cb.pack(side='left', padx=10)
            self.level_vars[lvl] = var

        # ------------------- Vista Previa -------------------
        prev_frame = ttk.LabelFrame(self, text="Vista Previa (máx 50)", padding=(10,8))
        prev_frame.pack(fill='both', expand=True, padx=15, pady=10)

        cols = ["DateTime", "Source", "EventID", "Message"]
        self.tree = ttk.Treeview(prev_frame, columns=cols, show='headings', height=10)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150, anchor='w')
        vsb = ttk.Scrollbar(prev_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)

        # ------------------- Logs -------------------
        log_frame = ttk.LabelFrame(self, text="Logs", padding=(10,8))
        log_frame.pack(fill='x', padx=15, pady=10)
        self.log_box = ScrolledText(log_frame, height=5, state='disabled', bg='#FFFFFF', fg='#333333', font=('Segoe UI', 9))
        self.log_box.pack(fill='x')

        # ------------------- Botones -------------------
        btn_frame = ttk.Frame(self, padding=(15,0))
        btn_frame.pack(fill='x', pady=10)

        self.load_btn = tk.Button(btn_frame, text="Cargar Eventos",
                                  bg="#0078D7", fg="#FFFFFF", font=('Segoe UI', 10, 'bold'),
                                  activebackground="#005A9E", command=self.load_events)
        self.load_btn.pack(side='left')

        self.export_btn = tk.Button(btn_frame, text="Exportar CSV",
                                    bg="#0078D7", fg="#FFFFFF", font=('Segoe UI', 10, 'bold'),
                                    activebackground="#005A9E", state='disabled', command=self.save_csv)
        self.export_btn.pack(side='left', padx=10)

    def log(self, msg: str):
        self.log_box['state'] = 'normal'
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box['state'] = 'disabled'

    def load_events(self):
        server = self.server_entry.get().strip() or 'localhost'
        log_type = self.log_type.get()
        keywords = [k.strip() for k in self.keywords_entry.get().split(',') if k.strip()]
        ids = [int(i) for i in self.event_id_entry.get().split(',') if i.strip().isdigit()]
        levels = [lvl for lvl, var in self.level_vars.items() if var.get()]

        self.log(f"Cargando eventos de '{server}' - {log_type}...")
        try:
            self.events = read_event_logs(server, log_type, levels, keywords, ids)
            self.log(f"{len(self.events)} eventos cargados.")
            for i in self.tree.get_children(): self.tree.delete(i)
            for evt in self.events[:50]:
                self.tree.insert('', 'end', values=(evt['DateTime'], evt['Source'], evt['EventID'], evt['Message']))
            self.export_btn['state'] = 'normal'
        except Exception as e:
            messagebox.showerror("Error leyendo eventos", str(e))
            self.log(f"ERROR: {e}")

    def save_csv(self):
        if not hasattr(self, 'events') or not self.events:
            messagebox.showwarning("Sin datos", "No hay eventos para exportar.")
            return

        default = sanitize_filename(f"events_{self.server_entry.get()}_{self.log_type.get()}.csv")
        path = filedialog.asksaveasfilename(defaultextension='.csv', initialfile=default,
                                            filetypes=[('CSV','*.csv'),('Todos','*.*')])
        if not path:
            return
        self.log(f"Exportando a '{path}'...")
        try:
            generate_csv_report(self.events, path)
            self.log("Exportación finalizada correctamente.")
            messagebox.showinfo("Listo", f"CSV generado:\n{path}")
        except Exception as e:
            messagebox.showerror("Error exportando", str(e))
            self.log(f"ERROR: {e}")

# =========================
# Lógica principal
# =========================
if __name__ == '__main__':
    app = EventAnalyzerApp()
    app.mainloop()
