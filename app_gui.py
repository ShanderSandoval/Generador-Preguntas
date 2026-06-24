import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys

# Importamos la lógica (El archivo moodle_logic.py debe estar en la misma carpeta)
try:
    from moodle_logic import SmartQuizParser, read_docx
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    if "docx" in str(e):
        messagebox.showerror("Error de Dependencia", "Falta la librería 'python-docx'.\nInstálala usando: pip install python-docx")
    else:
        messagebox.showerror("Error de Importación", f"No se pudo cargar la lógica:\n{e}")
    sys.exit()

class MoodleConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Cuestionarios a Moodle XML")
        self.root.geometry("900x700") 
        self.root.configure(padx=20, pady=20)
        self.root.minsize(800, 600)

        self.selected_files = []
        self.parser = SmartQuizParser()

        self._setup_ui()

    def _setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#004d99")
        style.configure("Total.TLabel", font=("Segoe UI", 14, "bold"), foreground="#2e7d32")

        ttk.Label(self.root, text="Conversor Inteligente Multiformato", style="Header.TLabel").pack(anchor="w", pady=(0, 5))
        ttk.Label(self.root, text="Procesa archivos Word en lote, audita las preguntas detectadas y genera tu XML.", foreground="#555").pack(anchor="w", pady=(0, 15))

        # --- PANEL DE SELECCIÓN ---
        frame_top = ttk.Frame(self.root)
        frame_top.pack(fill="x", pady=5)

        self.btn_select = ttk.Button(frame_top, text="📂 Añadir Archivos Word (.docx)", command=self.select_files)
        self.btn_select.pack(side="left")

        # BOTÓN LIMPIAR
        self.btn_clear = ttk.Button(frame_top, text="🗑️ Limpiar", command=self.clear_all, state="disabled")
        self.btn_clear.pack(side="left", padx=(5, 0))

        self.lbl_files_count = ttk.Label(frame_top, text="Ningún archivo seleccionado.", font=("Segoe UI", 10, "italic"))
        self.lbl_files_count.pack(side="left", padx=15)

        self.btn_analyze = ttk.Button(frame_top, text="🔍 1. Analizar y Previsualizar", command=self.start_analysis, state="disabled")
        self.btn_analyze.pack(side="right")

        # --- PESTAÑAS (NOTEBOOK) ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, pady=15)

        # Pestaña 1: Vista Previa (Maestro-Detalle)
        self.tab_preview = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_preview, text="👁️ Vista Previa de Preguntas")
        self._setup_preview_tab()

        # Pestaña 2: Consola / Log
        self.tab_log = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_log, text="⚙️ Consola de Procesamiento")
        self._setup_log_tab()

        # --- BARRA DE PROGRESO ---
        self.progress_var = tk.DoubleVar()
        self.progressbar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progressbar.pack(fill="x", pady=5)

        # --- PANEL INFERIOR (EXPORTAR) ---
        frame_bottom = ttk.Frame(self.root)
        frame_bottom.pack(fill="x", pady=5)

        self.lbl_total = ttk.Label(frame_bottom, text="Total Extraídas: 0", style="Total.TLabel")
        self.lbl_total.pack(side="left")

        self.btn_export = ttk.Button(frame_bottom, text="💾 2. Generar XML Final", command=self.export_xml, state="disabled")
        self.btn_export.pack(side="right")

    def _setup_preview_tab(self):
        # PanedWindow para dividir la pantalla horizontalmente
        self.paned = ttk.PanedWindow(self.tab_preview, orient="vertical")
        self.paned.pack(fill="both", expand=True, padx=5, pady=5)

        # --- PANEL SUPERIOR: LA TABLA ---
        frame_tabla = ttk.Frame(self.paned)
        self.paned.add(frame_tabla, weight=1)

        columns = ("numero", "tipo", "enunciado")
        self.tree = ttk.Treeview(frame_tabla, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("numero", text="#")
        self.tree.heading("tipo", text="Tipo")
        self.tree.heading("enunciado", text="Enunciado de la Pregunta (Clic para ver detalle abajo)")
        
        self.tree.column("numero", width=40, anchor="center")
        self.tree.column("tipo", width=80, anchor="center")
        self.tree.column("enunciado", width=610)

        # Evento para mostrar detalle al hacer clic
        self.tree.bind("<<TreeviewSelect>>", self._show_question_detail)

        scrollbar_tree = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_tree.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_tree.pack(side="right", fill="y")

        # --- PANEL INFERIOR: PREVISUALIZACIÓN MOODLE ---
        frame_detalle = ttk.LabelFrame(self.paned, text="Previsualización estilo Moodle")
        self.paned.add(frame_detalle, weight=1)

        self.txt_preview = tk.Text(frame_detalle, bg="#f0f4fa", font=("Segoe UI", 11), wrap="word", state="disabled", padx=15, pady=15)
        scrollbar_txt = ttk.Scrollbar(frame_detalle, orient="vertical", command=self.txt_preview.yview)
        self.txt_preview.configure(yscrollcommand=scrollbar_txt.set)

        self.txt_preview.pack(side="left", fill="both", expand=True)
        scrollbar_txt.pack(side="right", fill="y")

    def _setup_log_tab(self):
        self.txt_log = tk.Text(self.tab_log, bg="#1e1e1e", fg="#4af626", font=("Consolas", 9), state="disabled", wrap="word")
        scrollbar = ttk.Scrollbar(self.tab_log, command=self.txt_log.yview)
        self.txt_log.config(yscrollcommand=scrollbar.set)
        
        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def log(self, message):
        """Añade texto a la consola de log de forma segura."""
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", message + "\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        self.root.update_idletasks()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar documentos de Word",
            filetypes=(("Documentos de Word", "*.docx"), ("Todos los archivos", "*.*"))
        )
        if files:
            nuevos_archivos = 0
            # Añadimos a la lista solo los que no se hayan seleccionado antes
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    nuevos_archivos += 1
            
            if nuevos_archivos > 0:
                self.lbl_files_count.config(text=f"{len(self.selected_files)} documento(s) en cola.")
                self.btn_analyze.config(state="normal")
                self.btn_clear.config(state="normal")
                self.btn_export.config(state="disabled") # Bloqueamos exportar hasta auditar los nuevos
                
                # Limpiamos visuales para forzar un nuevo análisis
                self.txt_log.config(state="normal")
                self.txt_log.delete(1.0, "end")
                self.txt_log.config(state="disabled")
                
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                self.txt_preview.config(state="normal")
                self.txt_preview.delete(1.0, "end")
                self.txt_preview.config(state="disabled")
                
                self.lbl_total.config(text="Total Extraídas: 0")
                self.progress_var.set(0)
                self.log(f"Listo para analizar {len(self.selected_files)} archivo(s) acumulados en total.")

    def clear_all(self):
        """Reinicia la aplicación por completo."""
        self.selected_files = []
        self.parser = SmartQuizParser()
        
        self.lbl_files_count.config(text="Ningún archivo seleccionado.")
        self.lbl_total.config(text="Total Extraídas: 0")
        self.progress_var.set(0)
        
        self.txt_log.config(state="normal")
        self.txt_log.delete(1.0, "end")
        self.txt_log.config(state="disabled")
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
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
        self.parser = SmartQuizParser() 
        
        # Saltamos a la consola para ver el avance
        self.notebook.select(self.tab_log)
        threading.Thread(target=self._analysis_thread, daemon=True).start()

    def _analysis_thread(self):
        total_files = len(self.selected_files)
        
        try:
            for i, file_path in enumerate(self.selected_files):
                filename = os.path.basename(file_path)
                self.log(f"\n[{i+1}/{total_files}] Analizando: {filename}...")
                
                texto = read_docx(file_path)
                self.parser.parse_document(texto, log_callback=self.log)
                
                # Actualizar barra de progreso general por archivo
                progress = ((i + 1) / total_files) * 100
                self.progress_var.set(progress)
                
            self.log("\n[✓] Análisis completado. Poblando tabla de vista previa...")
            self.root.after(0, self._populate_preview)

        except Exception as e:
            self.log(f"\n[X] EXCEPCIÓN DETECTADA: {str(e)}")
            messagebox.showerror("Error de Análisis", f"El motor encontró un error:\n{str(e)}")
            self.root.after(0, self._restore_buttons)

    def _populate_preview(self):
        """Llena el Treeview con las preguntas detectadas."""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for index, q in enumerate(self.parser.questions):
            q_type = "V/F" if len(q['options']) == 2 else "Múltiple"
            clean_text = q['text'].replace('\n', ' ').strip()
                
            # Guardamos el index original en el iid para luego identificar la pregunta al hacer clic
            self.tree.insert("", "end", iid=str(index), values=(index + 1, q_type, clean_text))

        total_q = len(self.parser.questions)
        self.lbl_total.config(text=f"Total Extraídas: {total_q}")
        
        self.notebook.select(self.tab_preview)
        
        self._restore_buttons()
        if total_q > 0:
            self.btn_export.config(state="normal") 
            messagebox.showinfo("Auditoría Lista", "Análisis terminado.\n\nHaz clic en cualquier pregunta de la tabla superior para auditar su estructura en el panel inferior.")

    def _show_question_detail(self, event):
        """Dibuja las opciones con formato al seleccionar una pregunta en la tabla."""
        selected_item = self.tree.selection()
        if not selected_item:
            return
        
        idx = int(selected_item[0])
        q = self.parser.questions[idx]

        self.txt_preview.config(state="normal")
        self.txt_preview.delete(1.0, "end")

        texto_moodle = f"{q['text'].strip()}\n\nSeleccione una:\n"
        
        for opt in q['options']:
            es_correcta = opt.get('is_correct') or opt['letter'] == q.get('correct_letter')
            marcador = "( • )" if es_correcta else "(   )"
            texto_moodle += f"  {marcador} {opt['letter']}.  {opt['text'].strip()}\n\n"

        self.txt_preview.insert("end", texto_moodle)
        self.txt_preview.config(state="disabled")

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
            xml_final = self.parser.generate_all_xml(log_callback=self.log)
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(xml_final)
                
            self.log(f"\n[✓] XML GENERADO EXITOSAMENTE.\nRuta: {save_path}")
            messagebox.showinfo("Exportación Exitosa", "¡El archivo XML se ha guardado correctamente y está listo para importarse a Moodle!")
            
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"Hubo un problema escribiendo el archivo XML:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MoodleConverterApp(root)
    root.mainloop()