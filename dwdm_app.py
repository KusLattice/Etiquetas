import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import sys
import threading
from tkinterdnd2 import TkinterDnD, DND_FILES
import re
from datetime import datetime
from core_parser import parse_vsdx_connections, generate_excel

class TkinterDnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class App(TkinterDnDCTk):
    def __init__(self):
        super().__init__()

        # Configuración de ventana
        self.title("Proyecto Merge DWDM - Extractor Multicliente")
        self.geometry("650x600")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # Header Frame
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.header_frame.grid_columnconfigure(0, weight=1)
        
        self.header_label = ctk.CTkLabel(self.header_frame, text="⚡ DWDM VSDX Parser", font=ctk.CTkFont(size=24, weight="bold"))
        self.header_label.grid(row=0, column=0, sticky="w")
        
        # Selector de Telco
        self.telco_var = ctk.StringVar(value="ClaroVTR")
        self.telco_menu = ctk.CTkOptionMenu(
            self.header_frame, 
            values=["ClaroVTR", "Movistar", "WOM", "Entel"],
            variable=self.telco_var,
            font=ctk.CTkFont(weight="bold")
        )
        self.telco_menu.grid(row=0, column=1, sticky="e")
        
        self.subtitle_label = ctk.CTkLabel(self, text="Selecciona la Telco, arrastra el Visio y visualiza tu Excel", text_color="gray")
        self.subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="w")

        # Drag and Drop Frame
        self.file_frame = ctk.CTkFrame(self, height=120, fg_color="#1e1e1e", border_width=2, border_color="#3a7ebf")
        self.file_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)
        self.file_frame.grid_propagate(False)
        
        self.file_path_var = ctk.StringVar(value="📥 Arrastra y suelta el archivo .vsdx aquí\no haz click para buscar manualmente")
        self.file_label = ctk.CTkLabel(self.file_frame, textvariable=self.file_path_var, text_color="lightgray", anchor="center", font=ctk.CTkFont(size=14))
        self.file_label.grid(row=0, column=0, padx=10, pady=40, sticky="nsew")
        
        # Configurar Drag and Drop - Registrar ambos para que no haya zonas muertas
        self.file_frame.drop_target_register(DND_FILES)
        self.file_label.drop_target_register(DND_FILES)
        
        self.file_frame.dnd_bind('<<Drop>>', self.on_drop)
        self.file_label.dnd_bind('<<Drop>>', self.on_drop)
        
        self.file_frame.bind("<Button-1>", lambda e: self.select_file())
        self.file_label.bind("<Button-1>", lambda e: self.select_file())

        # Consola de log
        self.console = ctk.CTkTextbox(self, state="disabled")
        self.console.grid(row=4, column=0, padx=20, pady=10, sticky="nsew")

        # Contenedor de Botones
        self.bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_frame.grid(row=5, column=0, padx=20, pady=(10, 20), sticky="ew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)
        self.bottom_frame.grid_columnconfigure(1, weight=1)

        # Botón Acción
        self.btn_generate = ctk.CTkButton(self.bottom_frame, text="⚙️ Generar Excel", height=40, font=ctk.CTkFont(size=15, weight="bold"), command=self.start_generation, state="disabled")
        self.btn_generate.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        # Botón Visualizar Excel
        self.btn_open_excel = ctk.CTkButton(self.bottom_frame, text="📊 Ver Excel", height=40, fg_color="#2ecc71", hover_color="#27ae60", font=ctk.CTkFont(size=15, weight="bold"), command=self.open_excel, state="disabled")
        self.btn_open_excel.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        
        self.selected_file = None
        self.last_output_file = None

    def log(self, message):
        self.console.configure(state="normal")
        self.console.insert("end", f"> {message}\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    def on_drop(self, event):
        filepath = event.data
        if filepath.startswith('{') and filepath.endswith('}'):
            filepath = filepath[1:-1]
            
        if filepath.lower().endswith('.vsdx'):
            self.set_file(filepath)
        else:
            self.log("⚠️ Archivo ignorado. Solo se admiten archivos .vsdx")
            messagebox.showwarning("Archivo Inválido", "Por favor, arrastra un archivo de Visio (.vsdx)")

    def select_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecciona el diagrama VSDX",
            filetypes=[("Visio Files", "*.vsdx"), ("All Files", "*.*")]
        )
        if filepath:
            self.set_file(filepath)

    def set_file(self, filepath):
        self.selected_file = filepath
        self.file_path_var.set(f"📄 {os.path.basename(filepath)}")
        self.file_frame.configure(border_color="#2ecc71")
        self.btn_generate.configure(state="normal", fg_color=["#3a7ebf", "#1f538d"])
        self.btn_open_excel.configure(state="disabled")
        self.log(f"Archivo cargado listo para procesar: {filepath}")

    def start_generation(self):
        if not self.selected_file:
            return
            
        self.btn_generate.configure(state="disabled", text="Procesando...")
        self.btn_open_excel.configure(state="disabled")
        telco = self.telco_var.get()
        self.log(f"Iniciando extracción DWDM usando formato de {telco}...")
        
        threading.Thread(target=self.process_file, args=(telco,), daemon=True).start()

    def process_file(self, telco):
        try:
            # Generar nombre de salida con timestamp para evitar errores de permiso
            timestamp = datetime.now().strftime("%H%M%S")
            output_file = self.selected_file.replace(".vsdx", f"_{telco.upper()}_FIBRAS_{timestamp}.xlsx")
            
            # Pasamos el telco_name al parser core
            conns = parse_vsdx_connections(self.selected_file, telco_name=telco)
            self.log(f"Se extrajeron {len(conns)} conexiones físicas en crudo.")
            
            if not conns:
                self.log("⚠️ No se encontraron conexiones válidas. Revisa el Visio.")
                messagebox.showwarning("Sin Resultados", "No se encontraron conexiones para procesar.\n\nVerifica que las líneas estén cerca de los textos y que no sean metadatos filtrados.")
                return

            # Gen Excel
            count = generate_excel(conns, output_file)
            self.last_output_file = output_file
            
            self.log(f"¡Éxito! Se guardaron {count} conexiones únicas en Excel.")
            self.log(f"📍 Ruta: {output_file}")
            
            # Habilitar el botón de visualizar
            self.btn_open_excel.configure(state="normal")
            
        except Exception as e:
            self.log(f"❌ ERROR CRÍTICO: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error procesando el archivo:\n{str(e)}")
            
        finally:
            self.btn_generate.configure(state="normal", text="⚙️ Generar de Nuevo")

    def open_excel(self):
        if self.last_output_file and os.path.exists(self.last_output_file):
            self.log(f"Abriendo {os.path.basename(self.last_output_file)}...")
            if sys.platform == 'win32':
                os.startfile(self.last_output_file)
            else:
                import subprocess
                subprocess.call(['open', self.last_output_file])
        else:
            self.log("⚠️ El archivo Excel no existe o no se pudo generar.")
            messagebox.showwarning("Archivo no encontrado", "No se pudo encontrar el archivo Excel generado.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
