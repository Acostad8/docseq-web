import threading
import os
import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox

from core import numerar, insertar_seq, renumerar, verificar


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")


SECCION = "#2b2b2b" if ctk.get_appearance_mode() == "Dark" else "#e8e8e8"


class App(ctk.CTk):
    PAD = 12

    def __init__(self):
        super().__init__()
        self.title("DocSeq — Numeración automática para Word")
        self.geometry("840x680")
        self.minsize(700, 560)

        self.ruta_entrada = ctk.StringVar()
        self.ruta_salida = None
        self._procesando = False

        self._build_ui()
        self._log("Selecciona un archivo .docx para comenzar.")
        self._bind_drop()
        self.bind("<Control-o>", lambda e: self._browse_file())
        self.bind("<Control-O>", lambda e: self._browse_file())

    def _ts(self):
        return datetime.datetime.now().strftime("%H:%M:%S")

    def _format_size(self, path):
        size = os.path.getsize(path)
        if size < 1024: return f"{size} B"
        if size < 1048576: return f"{size/1024:.1f} KB"
        return f"{size/1048576:.1f} MB"

    def _bind_drop(self):
        self.drop_target_register = None
        try:
            self.tk.eval("package require tkdnd 2.0")
            self.tk.eval(f"tkdnd::drop_target register {self._w} {['{TKNDD:DND,FILES}']}")
            self.tk.createcommand("::tkdnd::Drop", self._on_drop)
        except:
            pass

    def _on_drop(self, *args):
        if args and args[0]:
            ruta = args[0][0]
            if ruta.endswith(".docx"):
                self.ruta_entrada.set(ruta)
                self._log(f"Archivo seleccionado: {ruta}")
                self._log(f"  Tamaño: {self._format_size(ruta)}")

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # ── Header ────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color=("gray90", "gray20"), corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            hdr,             text="DocSeq",
            font=ctk.CTkFont(size=22, weight="bold")
        ).grid(row=0, column=0, padx=self.PAD, pady=(14, 2), sticky="w")

        ctk.CTkLabel(
            hdr,
            text="Numeración de ilustraciones y tablas mediante campos SEQ",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60"),
        ).grid(row=1, column=0, padx=self.PAD, pady=(0, 14), sticky="w")

        # ── Card: Archivo ──────────────────────────────────────────
        card_file = ctk.CTkFrame(self, corner_radius=8)
        card_file.grid(row=1, column=0, padx=self.PAD, pady=(self.PAD, 0),
                       sticky="ew")
        card_file.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            card_file, text="Documento",
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, padx=self.PAD, pady=(self.PAD, 2), sticky="w",
               columnspan=3)

        ctk.CTkLabel(
            card_file, text="Ruta:",
            font=ctk.CTkFont(size=12)
        ).grid(row=1, column=0, padx=self.PAD, pady=4, sticky="w")

        self.entry_file = ctk.CTkEntry(
            card_file, textvariable=self.ruta_entrada,
            placeholder_text="Selecciona un archivo .docx...",
        )
        self.entry_file.grid(row=1, column=1, padx=4, pady=4, sticky="ew")

        self.btn_browse = ctk.CTkButton(
            card_file, text="Examinar", width=100,
            command=self._browse_file
        )
        self.btn_browse.grid(row=1, column=2, padx=(4, self.PAD), pady=4)

        self.file_size_label = ctk.CTkLabel(
            card_file, text="", font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray50")
        )
        self.file_size_label.grid(row=2, column=0, columnspan=3,
                                  padx=self.PAD, pady=(0, self.PAD), sticky="w")
        self.file_size_label.grid_remove()

        # ── Card: Acciones ─────────────────────────────────────────
        card_actions = ctk.CTkFrame(self, corner_radius=8)
        card_actions.grid(row=2, column=0, padx=self.PAD, pady=self.PAD,
                          sticky="nsew")
        card_actions.grid_columnconfigure(0, weight=1)
        card_actions.grid_rowconfigure(3, weight=1)

        # ── Encabezado Acciones ─────────────────────────────────────
        acciones_header = ctk.CTkFrame(card_actions, fg_color="transparent")
        acciones_header.grid(row=0, column=0, padx=self.PAD, pady=(self.PAD, 0), sticky="ew")
        acciones_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            acciones_header, text="Acciones",
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        # ── Botones principales ────────────────────────────────────
        btn_frame = ctk.CTkFrame(card_actions, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=self.PAD, pady=(8, 0), sticky="ew")
        btn_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.btn_paso1 = ctk.CTkButton(
            btn_frame, text="1  Numerar",
            font=ctk.CTkFont(size=13),
            command=lambda: self._procesar("numerar")
        )
        self.btn_paso1.grid(row=0, column=0, padx=4, pady=6, sticky="ew")

        self.btn_paso2 = ctk.CTkButton(
            btn_frame, text="2  Insertar SEQ",
            font=ctk.CTkFont(size=13),
            command=lambda: self._procesar("insertar_seq")
        )
        self.btn_paso2.grid(row=0, column=1, padx=4, pady=6, sticky="ew")

        self.btn_renumerar = ctk.CTkButton(
            btn_frame, text="Renumerar",
            font=ctk.CTkFont(size=13),
            fg_color=("gray50", "gray30"),
            hover_color=("gray40", "gray20"),
            command=lambda: self._procesar("renumerar")
        )
        self.btn_renumerar.grid(row=0, column=2, padx=4, pady=6, sticky="ew")

        self.btn_verificar = ctk.CTkButton(
            btn_frame, text="Verificar",
            font=ctk.CTkFont(size=13),
            fg_color=("gray50", "gray30"),
            hover_color=("gray40", "gray20"),
            command=lambda: self._procesar("verificar")
        )
        self.btn_verificar.grid(row=0, column=3, padx=4, pady=6, sticky="ew")

        # ── Botón Procesar Todo ────────────────────────────────────
        sep = ctk.CTkFrame(card_actions, height=1, fg_color=("gray70", "gray40"))
        sep.grid(row=2, column=0, padx=self.PAD * 2, pady=(8, 0), sticky="ew")

        self.btn_todo = ctk.CTkButton(
            card_actions, text="Procesar todo  (Paso 1  →  Paso 2)",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            fg_color="#1a6b1a",
            hover_color="#145214",
            command=lambda: self._procesar("todo")
        )
        self.btn_todo.grid(row=3, column=0, padx=self.PAD * 2, pady=(10, 4),
                           sticky="ew")

        # ── Log ────────────────────────────────────────────────────
        log_header = ctk.CTkFrame(card_actions, fg_color="transparent")
        log_header.grid(row=4, column=0, padx=self.PAD, pady=(8, 2), sticky="ew")
        log_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            log_header, text="Resultados",
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        self.btn_clear = ctk.CTkButton(
            log_header, text="Limpiar", width=70,
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            text_color=("gray50", "gray60"),
            hover=False,
            command=self._limpiar_log
        )
        self.btn_clear.grid(row=0, column=1, padx=(4, 0))

        self.log_text = ctk.CTkTextbox(
            card_actions, wrap="word",
            font=ctk.CTkFont(size=12, family="Consolas"),
            corner_radius=6,
        )
        self.log_text.grid(row=5, column=0, padx=self.PAD, pady=(0, self.PAD),
                           sticky="nsew")

        self.log_text.tag_config("error", foreground="#f87171")
        self.log_text.tag_config("warning", foreground="#fbbf24")
        self.log_text.tag_config("info", foreground="#4ade80")
        self.log_text.tag_config("dim", foreground="#555555")

        # ── Footer ─────────────────────────────────────────────────
        footer = ctk.CTkFrame(self, corner_radius=0, height=48,
                              fg_color=("gray90", "gray20"))
        footer.grid(row=3, column=0, sticky="ew")
        footer.grid_columnconfigure(2, weight=1)

        self.status_dot = ctk.CTkLabel(footer, text="●",
                                       font=ctk.CTkFont(size=16),
                                       text_color="#4ade80")
        self.status_dot.grid(row=0, column=0, padx=(self.PAD, 2), pady=10)

        self.status_text = ctk.CTkLabel(footer, text="Listo",
                                        font=ctk.CTkFont(size=12))
        self.status_text.grid(row=0, column=1, padx=(0, 8), pady=10)

        self.progress = ctk.CTkProgressBar(footer, mode="indeterminate",
                                           width=160, height=10)
        self.progress.grid(row=0, column=2, padx=8, pady=10, sticky="ew")
        self.progress.set(0)
        self.progress.grid_remove()

        self.btn_open = ctk.CTkButton(
            footer, text="Abrir carpeta", width=110,
            command=self._abrir_carpeta
        )
        self.btn_open.grid(row=0, column=3, padx=4, pady=10)

        self.btn_salir = ctk.CTkButton(
            footer, text="Salir", width=70,
            fg_color=("gray50", "gray30"),
            hover_color=("gray40", "gray20"),
            command=self.destroy
        )
        self.btn_salir.grid(row=0, column=4, padx=(4, self.PAD), pady=10)

    # ── Helpers ────────────────────────────────────────────────────

    def _log(self, msg, tag=None):
        ts = self._ts()
        if tag is None:
            if msg.startswith("ERROR") or msg.startswith("  [ERROR]") or msg.startswith("  ERROR"):
                tag = "error"
            elif msg.startswith("  [ADVERTENCIA]"):
                tag = "warning"
            elif msg.startswith("OK") or msg.startswith("  Resumen") or msg.startswith("Archivo guardado"):
                tag = "info"
        self.log_text.insert("end", f"{ts} ", "dim")
        self.log_text.insert("end", msg + "\n", tag) if tag else self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def _limpiar_log(self):
        self.log_text.delete("1.0", "end")
        self._log("Log limpiado.")

    def _browse_file(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar documento Word",
            filetypes=[("Documentos Word", "*.docx"), ("Todos", "*.*")])
        if ruta:
            self.ruta_entrada.set(ruta)
            self._log(f"Archivo seleccionado: {ruta}")
            try:
                self.file_size_label.configure(text=f"Tamaño: {self._format_size(ruta)}")
                self.file_size_label.grid()
            except:
                pass

    def _archivo_valido(self):
        ruta = self.ruta_entrada.get().strip()
        if not ruta:
            messagebox.showwarning("Sin archivo",
                                   "Selecciona un archivo .docx primero.")
            return None
        if not os.path.exists(ruta):
            messagebox.showerror("Archivo no encontrado",
                                 f"No se encuentra:\n{ruta}")
            return None
        return ruta

    def _set_processing(self, active):
        self._procesando = active
        estado = "disabled" if active else "normal"
        for btn in [self.btn_browse, self.btn_paso1, self.btn_paso2,
                    self.btn_renumerar, self.btn_verificar, self.btn_todo]:
            btn.configure(state=estado)
        self.entry_file.configure(state=estado)
        if active:
            self.status_dot.configure(text_color="#fbbf24")
            self.status_text.configure(text="Procesando...")
            self.progress.grid()
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.grid_remove()
            self.status_dot.configure(text_color="#4ade80")
            self.status_text.configure(text="Listo")

    def _set_error_status(self):
        self.status_dot.configure(text_color="#f87171")

    # ── Procesamiento ──────────────────────────────────────────────

    def _procesar(self, accion):
        if self._procesando:
            return
        ruta = self._archivo_valido()
        if not ruta:
            return
        self._set_processing(True)
        self.ruta_salida = None

        def target():
            try:
                safe_log = lambda m: self.after(0, lambda: self._log(m))
                if accion == "numerar":
                    safe_log("─" * 48)
                    safe_log("Paso 1 — Numerando ilustraciones y tablas...")
                    res = numerar.procesar(ruta, callback=safe_log)
                    self.after(0, lambda: self._resultado_numerar(res))
                elif accion == "insertar_seq":
                    safe_log("─" * 48)
                    safe_log("Paso 2 — Insertando campos SEQ...")
                    res = insertar_seq.procesar(ruta, callback=safe_log)
                    self.after(0, lambda: self._resultado_default(res))
                elif accion == "renumerar":
                    safe_log("─" * 48)
                    safe_log("Renumerando...")
                    res = renumerar.renumerar(ruta, callback=safe_log)
                    self.after(0, lambda: self._resultado_default(res))
                elif accion == "verificar":
                    safe_log("─" * 48)
                    safe_log("Verificando...")
                    res = verificar.verificar(ruta, callback=safe_log)
                    self.after(0, lambda: self._resultado_verificar(res))
                elif accion == "todo":
                    safe_log("═" * 48)
                    safe_log("Proceso completo:  Paso 1  →  Paso 2")
                    safe_log("")
                    safe_log("▶ Paso 1 — Numerar")
                    res1 = numerar.procesar(ruta, callback=safe_log)
                    if not res1.get("success"):
                        safe_log("")
                        safe_log("Paso 1 falló. Se cancela el proceso.")
                        self.after(0, lambda: self._resultado_numerar(res1))
                        return
                    self.after(0, lambda: self._resultado_numerar(res1, False))
                    safe_log("")
                    safe_log("▶ Paso 2 — Insertar SEQ")
                    ruta1 = res1["ruta_salida"]
                    res2 = insertar_seq.procesar(ruta1, callback=safe_log)
                    self.after(0, lambda: self._resultado_default(res2))
                    safe_log("")
                    safe_log("Proceso completo finalizado.")
                    safe_log("Abre el archivo en Word y presiona Ctrl+A → F9.")
                    safe_log("═" * 48)
            except Exception as e:
                self.after(0, lambda: self._on_error(str(e)))
            finally:
                self.after(0, lambda: self._set_processing(False))

        threading.Thread(target=target, daemon=True).start()

    def _resultado_numerar(self, res, mostrar_resumen=True):
        if not res.get("success"):
            self._log(f"  ERROR: {res.get('error', 'Error desconocido')}", "error")
            self.after(0, lambda: self._set_error_status())
            return
        c = res["contadores"]
        if mostrar_resumen:
            self._log("")
            self._log("  Resumen:", "info")
            for tipo, cuenta in c.items():
                self._log(f"    {tipo}s:  {cuenta}")
            self._log(f"    Modificados:  {res['modificados']}")
            if res["errores"]:
                self._log(f"    Errores:      {res['errores']}", "error")
        self.ruta_salida = res["ruta_salida"]

    def _resultado_default(self, res):
        if not res.get("success"):
            self._log(f"  ERROR: {res.get('error', 'Error desconocido')}", "error")
            self.after(0, lambda: self._set_error_status())
            return
        self.ruta_salida = res.get("ruta_salida")
        if self.ruta_salida:
            self._log(f"  Archivo guardado: {self.ruta_salida}", "info")

    def _resultado_verificar(self, res):
        if res["success"]:
            self._log("  Documento OK — sin errores.", "info")
        else:
            self._log(f"  {res['errores_totales']} problema(s) encontrado(s).", "error")
            self.after(0, lambda: self._set_error_status())
        self.ruta_salida = None

    def _on_error(self, msg):
        self._log(f"ERROR: {msg}", "error")
        self._set_error_status()
        messagebox.showerror("Error", msg)

    def _abrir_carpeta(self):
        ruta = self.ruta_salida
        if ruta and os.path.exists(ruta):
            os.startfile(os.path.dirname(os.path.abspath(ruta)))
            return
        ruta_base = self.ruta_entrada.get().strip()
        if ruta_base:
            carpeta = os.path.dirname(os.path.abspath(ruta_base))
            if os.path.exists(carpeta):
                os.startfile(carpeta)


def launch():
    ctk.set_appearance_mode("System")
    app = App()
    app.mainloop()
