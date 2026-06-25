import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import platform
import subprocess

# CORRECCIÓN DEL FALSO POSITIVO: Validación estricta del módulo 'docx'
try:
    from moodle_logic import SmartQuizParser, read_docx
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    if "No module named 'docx'" in str(e):
        messagebox.showerror("Error de Dependencia", "Falta la librería 'python-docx'.\nInstálala usando: pip install python-docx")
    else:
        # Si el error es por otra cosa (como desincronización de funciones), nos mostrará el motivo real
        messagebox.showerror("Error de Importación de Lógica", f"No se pudo cargar el archivo moodle_logic.py:\n{e}")
    sys.exit()

class MoodleConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Cuestionarios a Moodle XML")

        # ------------------------------------------------------------------
        # TAMAÑO Y POSICIÓN ADAPTATIVOS A CUALQUIER MONITOR
        # ------------------------------------------------------------------
        # En vez de un geometry() fijo (p.ej. 1000x750) que en monitores
        # pequeños o con escalado de Windows distinto deja botones fuera de
        # la pantalla, calculamos un tamaño relativo a la resolución real
        # disponible, con un mínimo razonable y centrado en pantalla.
        self.root.update_idletasks()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()

        # Usamos un porcentaje de la pantalla disponible, con topes sensatos
        win_w = min(1000, int(screen_w * 0.90))
        win_h = min(750, int(screen_h * 0.85))

        # Tamaño mínimo absoluto MUY bajo, para que en monitores pequeños
        # (ej. laptops 1366x768 con barra de tareas, o pantallas 720p)
        # la ventana siga siendo usable y nunca tape los botones inferiores.
        min_w = min(700, screen_w - 40)
        min_h = min(520, screen_h - 80)

        x = max(0, (screen_w - win_w) // 2)
        y = max(0, (screen_h - win_h) // 2)

        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(min_w, min_h)

        # Permite maximizar la ventana fácilmente si el usuario lo necesita
        try:
            if platform.system() == "Windows":
                self.root.state("zoomed")
            # En Linux/Mac no existe "zoomed" de forma fiable; dejamos el
            # tamaño calculado, que ya es responsivo.
        except tk.TclError:
            pass

        self.root.configure(padx=15, pady=15)

        self.selected_files = []
        self.parser = SmartQuizParser()
        self.current_source_file = None

        self._setup_ui()

    def _setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#004d99")
        style.configure("Total.TLabel", font=("Segoe UI", 12, "bold"), foreground="#2e7d32")

        # ------------------------------------------------------------------
        # CONTENEDOR RAÍZ CON GRID (en vez de pack puro)
        # ------------------------------------------------------------------
        # La clave para que los botones NUNCA desaparezcan en monitores
        # pequeños es usar grid con pesos: la fila del notebook (central)
        # se expande/encoge, pero las filas de cabecera, progreso y botones
        # inferiores tienen tamaño fijo y SIEMPRE son visibles porque Tkinter
        # las respeta antes de repartir el espacio restante.
        self.root.grid_rowconfigure(2, weight=1)   # fila del notebook crece/decrece
        self.root.grid_columnconfigure(0, weight=1)

        # --- CABECERA ---
        frame_header = ttk.Frame(self.root)
        frame_header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(frame_header, text="Conversor Inteligente Multiformato", style="Header.TLabel").pack(anchor="w")
        ttk.Label(frame_header, text="Procesa Word en lote, audita preguntas y exporta un XML purificado para Moodle.",
                  foreground="#555", wraplength=900).pack(anchor="w", pady=(2, 0))

        # --- PANEL SUPERIOR (selección / análisis) ---
        frame_top = ttk.Frame(self.root)
        frame_top.grid(row=1, column=0, sticky="ew", pady=5)

        self.btn_select = ttk.Button(frame_top, text="📂 Añadir Archivos Word (.docx)", command=self.select_files)
        self.btn_select.pack(side="left")

        self.btn_clear = ttk.Button(frame_top, text="🗑️ Limpiar", command=self.clear_all, state="disabled")
        self.btn_clear.pack(side="left", padx=(5, 0))

        self.lbl_files_count = ttk.Label(frame_top, text="Ningún archivo seleccionado.", font=("Segoe UI", 10, "italic"))
        self.lbl_files_count.pack(side="left", padx=15)

        self.btn_analyze = ttk.Button(frame_top, text="🔍 1. Analizar y Auditar", command=self.start_analysis, state="disabled")
        self.btn_analyze.pack(side="right")

        # --- NOTEBOOK (zona que se adapta: crece o se encoge según la pantalla) ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=2, column=0, sticky="nsew", pady=10)

        self.tab_preview = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_preview, text="👁️ Auditoría de Preguntas")
        self._setup_preview_tab()

        self.tab_log = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_log, text="⚙️ Consola de Procesamiento")
        self._setup_log_tab()

        # --- BARRA DE PROGRESO (siempre visible, tamaño fijo) ---
        self.progress_var = tk.DoubleVar()
        self.progressbar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progressbar.grid(row=3, column=0, sticky="ew", pady=5)

        # --- PANEL INFERIOR (totales + exportar) — SIEMPRE VISIBLE ---
        frame_bottom = ttk.Frame(self.root)
        frame_bottom.grid(row=4, column=0, sticky="ew", pady=(5, 0))

        self.lbl_total = ttk.Label(frame_bottom, text="Total: 0 | ✅ Válidas: 0 | ⚠️ Rechazadas: 0", style="Total.TLabel")
        self.lbl_total.pack(side="left")

        self.btn_export = ttk.Button(frame_bottom, text="💾 2. Exportar XML Válido", command=self.export_xml, state="disabled")
        self.btn_export.pack(side="right")

    def _setup_preview_tab(self):
        self.tab_preview.grid_rowconfigure(0, weight=1)
        self.tab_preview.grid_columnconfigure(0, weight=1)

        self.paned = ttk.PanedWindow(self.tab_preview, orient="vertical")
        self.paned.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # --- PANEL TABLA ---
        frame_tabla = ttk.Frame(self.paned)
        self.paned.add(frame_tabla, weight=1)

        # CHECKBOX DE FILTRO
        frame_filtro = ttk.Frame(frame_tabla)
        frame_filtro.pack(fill="x", pady=(0, 5))

        self.var_filtro_rechazadas = tk.BooleanVar(value=False)
        self.chk_filtro = ttk.Checkbutton(
            frame_filtro,
            text="⚠️ Mostrar SOLO preguntas RECHAZADAS (Sin respuesta)",
            variable=self.var_filtro_rechazadas,
            command=self._populate_preview,
            state="disabled"
        )
        self.chk_filtro.pack(side="left")

        # TABLA
        columns = ("numero", "estado", "archivo", "enunciado")
        self.tree = ttk.Treeview(frame_tabla, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("numero", text="N° en doc.")
        self.tree.heading("estado", text="Estado")
        self.tree.heading("archivo", text="Archivo Origen")
        self.tree.heading("enunciado", text="Enunciado")

        # Columnas con stretch para que se adapten al ancho real de la ventana
        self.tree.column("numero", width=40, minwidth=30, anchor="center", stretch=False)
        self.tree.column("estado", width=100, minwidth=80, anchor="center", stretch=False)
        self.tree.column("archivo", width=150, minwidth=100, anchor="w", stretch=True)
        self.tree.column("enunciado", width=300, minwidth=150, stretch=True)

        self.tree.bind("<<TreeviewSelect>>", self._show_question_detail)

        # Tag para resaltar en rojo las preguntas rechazadas (sin respuesta)
        self.tree.tag_configure("rechazada", foreground="#c62828")
        self.tree.tag_configure("valida", foreground="#000000")

        scrollbar_tree = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_tree.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_tree.pack(side="right", fill="y")

        # --- PANEL DETALLE ---
        frame_detalle = ttk.LabelFrame(self.paned, text="Previsualización Estructurada y Corrección")
        self.paned.add(frame_detalle, weight=1)
        frame_detalle.grid_rowconfigure(0, weight=1)
        frame_detalle.grid_columnconfigure(0, weight=1)

        self.txt_preview = tk.Text(frame_detalle, bg="#f0f4fa", font=("Segoe UI", 11), wrap="word", state="disabled", padx=15, pady=15)
        self.txt_preview.tag_configure("alerta_rechazo", foreground="#c62828", font=("Segoe UI", 11, "bold"))
        scrollbar_txt = ttk.Scrollbar(frame_detalle, orient="vertical", command=self.txt_preview.yview)
        self.txt_preview.configure(yscrollcommand=scrollbar_txt.set)

        self.txt_preview.grid(row=0, column=0, sticky="nsew")
        scrollbar_txt.grid(row=0, column=1, sticky="ns")

        # BOTÓN ABRIR ARCHIVO — fila propia, siempre visible bajo el texto
        self.btn_open_file = ttk.Button(frame_detalle, text="📝 Abrir Documento Word Original para Corregir", command=self.open_source_file, state="disabled")
        self.btn_open_file.grid(row=1, column=0, columnspan=2, sticky="ew", padx=15, pady=10)

    def _setup_log_tab(self):
        self.tab_log.grid_rowconfigure(0, weight=1)
        self.tab_log.grid_columnconfigure(0, weight=1)

        self.txt_log = tk.Text(self.tab_log, bg="#1e1e1e", fg="#4af626", font=("Consolas", 9), state="disabled", wrap="word")
        scrollbar = ttk.Scrollbar(self.tab_log, command=self.txt_log.yview)
        self.txt_log.config(yscrollcommand=scrollbar.set)
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

    def log(self, message):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", message + "\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        self.root.update_idletasks()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar documentos Word",
            filetypes=(("Documentos de Word", "*.docx"), ("Todos los archivos", "*.*"))
        )
        if files:
            nuevos_archivos = sum(1 for f in files if f not in self.selected_files)
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)

            if nuevos_archivos > 0:
                self.lbl_files_count.config(text=f"{len(self.selected_files)} documento(s) en cola.")
                self.btn_analyze.config(state="normal")
                self.btn_clear.config(state="normal")
                self.btn_export.config(state="disabled")
                self.chk_filtro.config(state="disabled")
                self.btn_open_file.config(state="disabled")

                self.txt_log.config(state="normal")
                self.txt_log.delete(1.0, "end")
                self.txt_log.config(state="disabled")

                for item in self.tree.get_children(): self.tree.delete(item)

                self.txt_preview.config(state="normal")
                self.txt_preview.delete(1.0, "end")
                self.txt_preview.config(state="disabled")

                self.lbl_total.config(text="Total: 0 | ✅ Válidas: 0 | ⚠️ Rechazadas: 0")
                self.progress_var.set(0)
                self.log(f"Listo para analizar {len(self.selected_files)} archivo(s).")

    def clear_all(self):
        self.selected_files = []
        self.parser = SmartQuizParser()
        self.current_source_file = None

        self.lbl_files_count.config(text="Ningún archivo seleccionado.")
        self.lbl_total.config(text="Total: 0 | ✅ Válidas: 0 | ⚠️ Rechazadas: 0")
        self.progress_var.set(0)
        self.var_filtro_rechazadas.set(False)
        self.chk_filtro.config(state="disabled")
        self.btn_open_file.config(state="disabled")

        self.txt_log.config(state="normal")
        self.txt_log.delete(1.0, "end")
        self.txt_log.config(state="disabled")

        for item in self.tree.get_children(): self.tree.delete(item)

        self.txt_preview.config(state="normal")
        self.txt_preview.delete(1.0, "end")
        self.txt_preview.config(state="disabled")

        self.btn_analyze.config(state="disabled")
        self.btn_export.config(state="disabled")
        self.btn_clear.config(state="disabled")

    def start_analysis(self):
        self.btn_select.config(state="disabled")
        self.btn_clear.config(state="disabled")
        self.btn_analyze.config(state="disabled")
        self.btn_export.config(state="disabled")
        self.chk_filtro.config(state="disabled")
        self.parser = SmartQuizParser()

        self.notebook.select(self.tab_log)
        threading.Thread(target=self._analysis_thread, daemon=True).start()

    def _analysis_thread(self):
        total_files = len(self.selected_files)
        try:
            for i, file_path in enumerate(self.selected_files):
                filename = os.path.basename(file_path)
                self.log(f"\n[{i+1}/{total_files}] Analizando: {filename}...")

                texto = read_docx(file_path)
                self.parser.parse_document(texto, file_path=file_path, log_callback=self.log)

                self.progress_var.set(((i + 1) / total_files) * 100)

            self.log("\n[✓] Análisis completado. Auditando validación de respuestas...")
            self.root.after(0, self._populate_preview)

        except Exception as e:
            self.log(f"\n[X] EXCEPCIÓN DETECTADA: {str(e)}")
            messagebox.showerror("Error", f"Fallo en motor:\n{str(e)}")
            self.root.after(0, self._restore_buttons)

    def _populate_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.parser.validate_questions()
        solo_rechazadas = self.var_filtro_rechazadas.get()

        count_validas = 0
        count_rechazadas = 0

        # Contador de pregunta POR ARCHIVO: se reinicia cada vez que cambia
        # el archivo de origen, siguiendo el orden en que aparecen.
        contador_por_archivo = {}

        for index, q in enumerate(self.parser.questions):
            is_valid = q.get('is_valid', False)

            if is_valid: count_validas += 1
            else: count_rechazadas += 1

            file_path = q.get('file_path', 'Desconocido')
            contador_por_archivo[file_path] = contador_por_archivo.get(file_path, 0) + 1
            numero_en_archivo = contador_por_archivo[file_path]
            # Lo guardamos en la pregunta para poder reutilizarlo en el detalle
            q['numero_en_archivo'] = numero_en_archivo

            if solo_rechazadas and is_valid:
                continue

            estado = "✅ Válida" if is_valid else "⚠️ Rechazada"
            archivo_nombre = os.path.basename(file_path)
            clean_text = q['text'].replace('\n', ' ').strip()

            tag = "valida" if is_valid else "rechazada"

            self.tree.insert(
                "", "end", iid=str(index),
                values=(numero_en_archivo, estado, archivo_nombre, clean_text),
                tags=(tag,)
            )

        total_q = len(self.parser.questions)
        self.lbl_total.config(text=f"Total: {total_q}  |  ✅ Válidas: {count_validas}  |  ⚠️ Rechazadas: {count_rechazadas}")

        self.notebook.select(self.tab_preview)
        self._restore_buttons()

        if total_q > 0:
            self.chk_filtro.config(state="normal")

        if count_validas > 0:
            self.btn_export.config(state="normal")

        if count_rechazadas > 0 and not solo_rechazadas:
            messagebox.showwarning("Atención: Preguntas Incompletas", f"Se detectaron {count_rechazadas} preguntas sin respuesta marcada.\n\nUtiliza la tabla para auditarlas y abrir el Word para corregirlas. Estas preguntas NO se exportarán.")

    def _show_question_detail(self, event):
        selected_item = self.tree.selection()
        if not selected_item: return

        idx = int(selected_item[0])
        q = self.parser.questions[idx]

        self.current_source_file = q.get('file_path')

        self.txt_preview.config(state="normal")
        self.txt_preview.delete(1.0, "end")

        archivo_nombre = os.path.basename(q.get('file_path', 'Desconocido'))
        numero_en_archivo = q.get('numero_en_archivo', idx + 1)

        if not q.get('is_valid', False):
            encabezado_alerta = (
                f"🛑 PREGUNTA {numero_en_archivo} DEL ARCHIVO \"{archivo_nombre}\" "
                f"NO TIENE RESPUESTA Y NO SERÁ EXPORTADA.\n"
                f"Utiliza el botón inferior para abrir el documento original y corregirla.\n\n"
            )
            self.txt_preview.insert("end", encabezado_alerta, "alerta_rechazo")
            self.btn_open_file.config(state="normal")
        else:
            self.btn_open_file.config(state="disabled")

        texto_moodle = f"Pregunta {numero_en_archivo} de \"{archivo_nombre}\"\n\n"
        texto_moodle += f"{q['text'].strip()}\n\nSeleccione una:\n"

        for opt in q['options']:
            es_correcta = opt.get('is_correct') or opt['letter'] == q.get('correct_letter')
            marcador = "( • )" if es_correcta else "(   )"
            texto_moodle += f"  {marcador} {opt['letter']}.  {opt['text'].strip()}\n\n"

        self.txt_preview.insert("end", texto_moodle)
        self.txt_preview.config(state="disabled")

    def open_source_file(self):
        if self.current_source_file and os.path.exists(self.current_source_file):
            try:
                if platform.system() == 'Darwin':       # macOS
                    subprocess.call(('open', self.current_source_file))
                elif platform.system() == 'Windows':    # Windows
                    os.startfile(self.current_source_file)
                else:                                   # Linux
                    subprocess.call(('xdg-open', self.current_source_file))
            except Exception as e:
                messagebox.showerror("Error de Sistema", f"No se pudo abrir el archivo:\n{e}")

    def _restore_buttons(self):
        self.btn_select.config(state="normal")
        self.btn_clear.config(state="normal")
        self.btn_analyze.config(state="normal")

    def export_xml(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xml",
            initialfile="Banco_Maestro_Moodle.xml",
            title="Guardar archivo consolidado XML como...",
            filetypes=[("Moodle XML", "*.xml")]
        )

        if not save_path: return

        try:
            xml_final = self.parser.generate_valid_xml(log_callback=self.log)
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(xml_final)

            self.log(f"\n[✓] XML GENERADO EXITOSAMENTE.\nRuta: {save_path}")
            messagebox.showinfo("Exportación Limpia", "¡Archivo XML guardado!\n\nNota: Solo se incluyeron las preguntas válidas.")

        except Exception as e:
            messagebox.showerror("Error al Guardar", f"Hubo un problema escribiendo el XML:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MoodleConverterApp(root)
    root.mainloop()