import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import sys
import threading
from tkinterdnd2 import TkinterDnD, DND_FILES
import re
from datetime import datetime
from PIL import Image
import math
import time
from tkinter import font as tkfont
from core_parser import parse_vsdx_connections, generate_excel
from pdf_parser import parse_pdf_connections

class TkinterDnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class App(TkinterDnDCTk):
    def __init__(self):
        super().__init__()

        # Configuración de ventana
        self.title("Generador de Etiquetas")
        self.geometry("680x650")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        # Temas Disponibles
        self.themes = {
            "Dark": {
                "bg": "#0a0a0a", "header": "#1a1a1a", "btn_acc": "#2c3e50", "btn_ok": "#27ae60", 
                "text": "#e0e0e0", "btn_text": "white", "accent": "#3498db", "font": "Inter",
                "decor": "🖥️ 🛰️ 🏢 📶 💠"
            },
            "Matrix": {
                "bg": "#000000", "header": "#000000", "btn_acc": "#003b00", "btn_ok": "#008f11", 
                "text": "#00ff41", "btn_text": "#00ff41", "accent": "#00ff41", "font": "Consolas",
                "decor": "0101 1010 0110 1101"
            },
            "El Trauco": {
                "bg": "#142114", "header": "#0a120a", "btn_acc": "#2d5a27", "btn_ok": "#3e7a3a", 
                "text": "#d4e1d1", "btn_text": "white", "accent": "#4a7c44", "font": "Georgia",
                "decor": "🌲 🍄 👺 🌿 🪵"
            },
            "La Pincoya": {
                "bg": "#003d4d", "header": "#002b36", "btn_acc": "#007c91", "btn_ok": "#4db6ac", 
                "text": "#e0f2f1", "btn_text": "white", "accent": "#26c6da", "font": "Trebuchet MS",
                "decor": "🧜‍♀️ 🐚 🌊 🏝️ 🐠"
            },
            "El Caleuche": {
                "bg": "#0d1a1a", "header": "#050f0f", "btn_acc": "#1a4a4a", "btn_ok": "#26a69a", 
                "text": "#a7ffeb", "btn_text": "black", "accent": "#64ffda", "font": "Courier New",
                "decor": "🚢 👻 🌫️ ⚓ 🌑"
            }
        }
        self.current_theme_name = "Dark"

        # Header Frame
        self.header_frame = ctk.CTkFrame(self, fg_color="#1a1a2e", corner_radius=12)
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="ew")
        self.header_frame.grid_columnconfigure(0, weight=1)

        self.header_label = ctk.CTkLabel(
            self.header_frame,
            text="Generador de Etiquetas",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.header_label.grid(row=0, column=0, padx=15, pady=(20, 5), sticky="w")

        # Decoración temática
        self.decor_label = ctk.CTkLabel(
            self.header_frame,
            text=self.themes[self.current_theme_name]["decor"],
            font=ctk.CTkFont(size=18)
        )
        self.decor_label.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="w")

        # Configuración de Proyectos por Empresa
        self.projects_map = {
            "Movistar": ["Core", "Genérico"],
            "ClaroVTR": ["Merge", "Genérico"],
            "WOM": ["Core", "Genérico"],
            "Entel": ["Core", "Genérico"],
            "Tigo": ["Core", "Genérico"]
        }

        # Cargar Logos
        try:
            self.logos = {
                "ClaroVTR": ctk.CTkImage(light_image=Image.open("resources/claro.png"), dark_image=Image.open("resources/claro.png"), size=(100, 50)),
                "Movistar": ctk.CTkImage(light_image=Image.open("resources/movistar.png"), dark_image=Image.open("resources/movistar.png"), size=(80, 80)),
                "WOM": ctk.CTkImage(light_image=Image.open("resources/wom.png"), dark_image=Image.open("resources/wom.png"), size=(100, 50)),
                "Entel": ctk.CTkImage(light_image=Image.open("resources/entel.png"), dark_image=Image.open("resources/entel.png"), size=(100, 50)),
                "Tigo": ctk.CTkImage(light_image=Image.open("resources/tigo.png"), dark_image=Image.open("resources/tigo.png"), size=(100, 50))
            }
        except Exception:
            self.logos = {}

        # Frame de Selectores
        self.selectors_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.selectors_frame.grid(row=0, column=1, padx=15, pady=10, sticky="e")

        # Contenedor para el logo flotante (Rectangular)
        self.logo_container = ctk.CTkFrame(self.selectors_frame, width=140, height=80, fg_color="transparent")
        self.logo_container.grid(row=0, column=0, rowspan=2, padx=(0, 15))
        self.logo_container.grid_propagate(False)

        # Logo de la Operadora actual
        self.logo_display = ctk.CTkLabel(self.logo_container, text="", image=self.logos.get("Movistar"))
        self.logo_display.place(relx=0.5, rely=0.5, anchor="center")

        # Operadora
        self.telco_var = ctk.StringVar(value="Movistar")
        self.telco_menu = ctk.CTkOptionMenu(
            self.selectors_frame,
            values=list(self.projects_map.keys()),
            variable=self.telco_var,
            command=self.update_projects,
            font=ctk.CTkFont(weight="bold"),
            width=140,
            fg_color="#2c3e50",
            button_color="#34495e"
        )
        self.telco_menu.grid(row=0, column=1, padx=5, pady=5)

        # Proyecto (Sub-menú)
        self.project_var = ctk.StringVar(value="Core")
        self.project_menu = ctk.CTkOptionMenu(
            self.selectors_frame,
            values=self.projects_map["Movistar"],
            variable=self.project_var,
            font=ctk.CTkFont(weight="bold"),
            width=140,
            fg_color="#34495e",
            button_color="#2c3e50"
        )
        self.project_menu.grid(row=1, column=1, padx=5, pady=5)

        # Selector de Temas
        self.theme_var = ctk.StringVar(value="Dark")
        self.theme_menu = ctk.CTkOptionMenu(
            self.selectors_frame,
            values=list(self.themes.keys()),
            variable=self.theme_var,
            command=self.change_theme,
            font=ctk.CTkFont(size=11),
            width=140,
            fg_color="#34495e",
            button_color="#2c3e50"
        )
        self.theme_menu.grid(row=2, column=1, padx=5, pady=5)
        
        self.theme_label = ctk.CTkLabel(self.selectors_frame, text="Diseño:", font=ctk.CTkFont(size=12))
        self.theme_label.grid(row=2, column=0, padx=(10, 5), sticky="e")

        self.subtitle_label = ctk.CTkLabel(
            self,
            text="Arrastra un archivo .vsdx o .pdf y genera el Excel de etiquetas P-touch",
            text_color="gray"
        )
        self.subtitle_label.grid(row=1, column=0, padx=20, pady=(8, 10), sticky="w")

        # Drag and Drop Frame
        self.file_frame = ctk.CTkFrame(self, height=120, fg_color="#1e1e1e", border_width=2, border_color="#3a7ebf")
        self.file_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)
        self.file_frame.grid_propagate(False)
        
        self.file_path_var = ctk.StringVar(value="📥 Arrastra y suelta el archivo .vsdx o .pdf aquí\no haz click para buscar manualmente")
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
        self.btn_generate = ctk.CTkButton(
            self.bottom_frame, 
            text="⚙️ Analizar y crear", 
            height=40, 
            font=ctk.CTkFont(size=15, weight="bold"), 
            command=self.start_generation, 
            state="disabled",
            fg_color="#3498db",
            hover_color="#2980b9"
        )
        self.btn_generate.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        # Botón Visualizar Excel
        self.btn_open_excel = ctk.CTkButton(
            self.bottom_frame, 
            text="📊 Abrir Excel", 
            height=40, 
            fg_color="#2ecc71", 
            hover_color="#27ae60", 
            font=ctk.CTkFont(size=15, weight="bold"), 
            command=self.open_excel, 
            state="disabled"
        )
        self.btn_open_excel.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        
        self.selected_file = None
        self.last_output_file = None

        # Iniciar Animaciones
        self.animate_logo()

    def update_projects(self, selected_telco):
        """Actualiza el sub-menú de proyectos basado en la operadora."""
        # Actualizar Logo
        if selected_telco in self.logos:
            self.logo_display.configure(image=self.logos[selected_telco])

        new_projects = self.projects_map.get(selected_telco, ["Core"])
        self.project_menu.configure(values=new_projects)
        self.project_var.set(new_projects[0])
        self.log(f"Cambiado a {selected_telco} — Proyecto: {new_projects[0]}")

    def change_theme(self, theme_name):
        """Aplica cambios visuales según el tema seleccionado."""
        self.current_theme_name = theme_name
        theme = self.themes[theme_name]
        
        # Appearance Mode
        if theme_name == "Minimalist":
            ctk.set_appearance_mode("light")
        else:
            ctk.set_appearance_mode("dark")

        # UI Updates
        self.configure(fg_color=theme["bg"])
        self.header_frame.configure(fg_color=theme["header"])
        self.header_label.configure(
            text="Generador de Etiquetas",
            text_color=theme["text"], 
            font=ctk.CTkFont(family=theme["font"], size=26, weight="bold")
        )
        self.subtitle_label.configure(
            text="Arrastra un archivo .vsdx o .pdf y genera el Excel de etiquetas P-touch",
            text_color=theme["text"], 
            font=ctk.CTkFont(family=theme["font"], size=13)
        )
        self.file_frame.configure(fg_color=theme["bg"] if theme_name != "Minimalist" else "white", border_color=theme["accent"])
        self.file_label.configure(text_color=theme["text"])
        self.theme_label.configure(text_color=theme["text"])
        
        self.btn_generate.configure(
            text="⚙️ Analizar y crear",
            fg_color=theme["btn_acc"], 
            hover_color=theme["accent"],
            text_color=theme["btn_text"]
        )
        self.btn_open_excel.configure(
            fg_color=theme["btn_ok"], 
            text_color=theme["btn_text"]
        )
        
        # Actualizar Decoración
        self.decor_label.configure(text=theme["decor"], text_color=theme["text"])
        
        # Consola y detalles específicos
        if theme_name == "Matrix":
            console_bg = "#000000"
            console_text = "#00ff41"
            font_family = "Consolas"
        elif theme_name == "El Trauco":
            console_bg = "#0d140d"
            console_text = "#a3b1a3"
            font_family = "Georgia"
        elif theme_name == "La Pincoya":
            console_bg = "#002b36"
            console_text = "#b2dfdb"
            font_family = "Trebuchet MS"
        elif theme_name == "El Caleuche":
            console_bg = "#050f0f"
            console_text = "#64ffda"
            font_family = "Courier New"
        else:
            console_bg = "#1a1a1a"
            console_text = "#e0e0e0"
            font_family = "Inter"

        self.console.configure(fg_color=console_bg, text_color=console_text, font=ctk.CTkFont(family=font_family, size=12))

        self.log(f"Tema aplicado: {theme_name}")

    def animate_logo(self):
        """Efecto de flotación suave para el logo."""
        try:
            t = time.time() * 2.5  # Velocidad de oscilación
            y_offset = math.sin(t) * 0.1  # Amplitud del movimiento
            self.logo_display.place(relx=0.5, rely=0.5 + y_offset, anchor="center")
            self.after(30, self.animate_logo)
        except Exception:
            pass  # Evitar errores al cerrar la app

    def log(self, message):
        self.console.configure(state="normal")
        self.console.insert("end", f"> {message}\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    def on_drop(self, event):
        filepath = event.data
        if filepath.startswith('{') and filepath.endswith('}'):
            filepath = filepath[1:-1]
            
        if filepath.lower().endswith('.vsdx') or filepath.lower().endswith('.pdf'):
            self.set_file(filepath)
        else:
            self.log("⚠️ Archivo ignorado. Solo se admiten archivos .vsdx o .pdf")
            messagebox.showwarning("Archivo Inválido", "Por favor, arrastra un archivo Visio (.vsdx) o PDF (.pdf)")

    def select_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecciona el diagrama DWDM",
            filetypes=[("Diagramas DWDM", "*.vsdx *.pdf"), ("Visio", "*.vsdx"), ("PDF", "*.pdf"), ("All Files", "*.*")]
        )
        if filepath:
            self.set_file(filepath)

    def set_file(self, filepath):
        # Limpiar consola al cargar nuevo archivo
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

        self.selected_file = filepath
        self.file_path_var.set(f"📄 {os.path.basename(filepath)}")
        self.file_frame.configure(border_color="#2ecc71")
        self.btn_generate.configure(state="normal")
        self.btn_open_excel.configure(state="disabled")
        self.log(f"Archivo cargado listo para procesar: {filepath}")

    def start_generation(self):
        if not self.selected_file:
            return
            
        self.btn_generate.configure(state="disabled", text="Procesando...")
        self.btn_open_excel.configure(state="disabled")
        telco = self.telco_var.get()
        project = self.project_var.get()
        self.log(f"Iniciando extracción DWDM para {telco} — {project}...")
        
        threading.Thread(target=self.process_file, args=(telco, project), daemon=True).start()

    def process_file(self, telco, project):
        try:
            # Nombre de salida con timestamp
            timestamp = datetime.now().strftime("%H%M%S")
            ext = os.path.splitext(self.selected_file)[1].lower()
            base = re.sub(r'\.(vsdx|pdf)$', '', self.selected_file, flags=re.IGNORECASE)
            output_file = f"{base}_{telco.upper()}_{project.upper()}_{timestamp}.xlsx"

            # Despachar al parser correcto según extensión
            if ext == '.pdf':
                self.log("[PDF] Parseando diagrama vectorial PDF...")
                conns = parse_pdf_connections(self.selected_file)
            else:
                self.log("[VSDX] Parseando diagrama Visio...")
                conns = parse_vsdx_connections(self.selected_file, telco_name=telco, project_name=project)

            self.log(f"Se extrajeron {len(conns)} conexiones físicas.")
            
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
