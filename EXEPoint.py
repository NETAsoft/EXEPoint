import os
import sys
import shutil
import subprocess
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

BTN_TEXT = "CONVERT TO .EXE"


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def pyinstaller_command():
    """Rodando como script usa 'python -m PyInstaller'.
    Rodando como .exe depende de um 'pyinstaller' instalado no PATH."""
    if getattr(sys, 'frozen', False):
        return ["pyinstaller"]
    return [sys.executable, "-m", "PyInstaller"]


class EXEPointApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EXEPoint - PowerPoint to EXE Converter")
        self.root.geometry("550x470")
        self.root.resizable(False, False)

        # Cores estilo PowerPoint
        self.main_color = "#B7472A"
        self.bg_color = "#F3F2F1"
        self.card_color = "#FFFFFF"
        self.text_color = "#323130"

        self.root.configure(bg=self.bg_color)
        self.ppsx_path = ""
        self.icon_path = ""
        self.output_path = ""
        self.output_custom = False

        # --- BARRA DE TOPO ---
        self.top_bar = tk.Frame(self.root, bg=self.main_color, height=55)
        self.top_bar.pack(fill=tk.X)
        self.top_bar.pack_propagate(False)

        # Carrega o logo.png na interface
        self.logo_image = None
        logo_file = resource_path("logo.png")
        if os.path.exists(logo_file):
            try:
                img = Image.open(logo_file).resize((32, 32), Image.Resampling.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(img)
                self.lbl_logo = tk.Label(self.top_bar, image=self.logo_image, bg=self.main_color)
                self.lbl_logo.pack(side=tk.LEFT, padx=(15, 0))
            except Exception:
                pass

        self.lbl_title = tk.Label(
            self.top_bar, text="EXEPoint", fg="white", bg=self.main_color,
            font=("Segoe UI", 16, "bold"), padx=10
        )
        self.lbl_title.pack(side=tk.LEFT)

        self.lbl_subtitle = tk.Label(
            self.top_bar, text="Kiosk Mode (Native Animations)", fg="#F1C4B9", bg=self.main_color,
            font=("Segoe UI", 9, "italic")
        )
        # CORRIGIDO: anchor='b' não existe no Tkinter, o equivalente é 's' (south)
        self.lbl_subtitle.pack(side=tk.LEFT, anchor='s', pady=14)

        # --- ÁREA CENTRAL ---
        self.card = tk.Frame(self.root, bg=self.card_color, bd=1, relief=tk.SOLID, highlightbackground="#D2D0CE")
        self.card.pack(fill=tk.X, padx=30, pady=20)

        # Seleção da Apresentação
        tk.Label(self.card, text="Select presentation (.ppsx or .pptx):", fg=self.text_color,
                 bg=self.card_color, font=("Segoe UI", 10, "bold")).pack(pady=(10, 2))
        self.btn_select = tk.Button(self.card, text="Browse Presentation", command=self.select_file,
                                    bg=self.main_color, fg="white", relief=tk.FLAT,
                                    font=("Segoe UI", 9), cursor="hand2")
        self.btn_select.pack()
        self.lbl_file = tk.Label(self.card, text="No file selected", fg="#797775", bg=self.card_color,
                                 font=("Segoe UI", 9), wraplength=450)
        self.lbl_file.pack(pady=(2, 10))

        # Seleção do Ícone (.ico)
        tk.Label(self.card, text="Select Custom Icon (.ico) - Optional:", fg=self.text_color,
                 bg=self.card_color, font=("Segoe UI", 10, "bold")).pack(pady=(5, 2))
        self.btn_icon = tk.Button(self.card, text="Browse Icon (.ico)", command=self.select_icon,
                                  bg="#505050", fg="white", relief=tk.FLAT,
                                  font=("Segoe UI", 9), cursor="hand2")
        self.btn_icon.pack()
        self.lbl_icon_status = tk.Label(self.card, text="Default PyInstaller icon will be used",
                                        fg="#797775", bg=self.card_color,
                                        font=("Segoe UI", 9), wraplength=450)
        self.lbl_icon_status.pack(pady=(2, 10))

        # Destino do .exe
        tk.Label(self.card, text="Save .exe as:", fg=self.text_color, bg=self.card_color,
                 font=("Segoe UI", 10, "bold")).pack(pady=(5, 2))
        self.btn_output = tk.Button(self.card, text="Choose Save Location", command=self.select_output,
                                    bg="#505050", fg="white", relief=tk.FLAT,
                                    font=("Segoe UI", 9), cursor="hand2")
        self.btn_output.pack()
        self.lbl_output = tk.Label(self.card, text="Select a presentation first", fg="#797775",
                                   bg=self.card_color, font=("Segoe UI", 9), wraplength=450)
        self.lbl_output.pack(pady=(2, 10))

        # --- BOTÃO CONVERTER ---
        self.btn_generate = tk.Button(
            self.root, text=BTN_TEXT, command=self.generate_executable,
            bg="#323130", fg="white", relief=tk.FLAT, state=tk.DISABLED,
            font=("Segoe UI", 12, "bold"), activebackground="#201F1E", activeforeground="white",
            padx=30, pady=8, cursor="hand2"
        )
        self.btn_generate.pack(pady=(5, 5))

        # Rodapé de status
        self.lbl_status = tk.Label(
            self.root, text="Ready", fg="#797775", bg=self.bg_color, font=("Segoe UI", 9)
        )
        self.lbl_status.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="EXEPoint - Select Presentation",
            filetypes=[("PowerPoint Files", "*.ppsx *.pptx")]
        )
        if path:
            self.ppsx_path = os.path.abspath(path)
            self.lbl_file.config(text=os.path.basename(path), fg=self.text_color)
            self.btn_generate.config(state=tk.NORMAL, bg=self.main_color)
            if not self.output_custom:
                base = os.path.splitext(os.path.basename(self.ppsx_path))[0]
                self.set_output(os.path.join(os.path.dirname(self.ppsx_path), f"{base}.exe"))
            self.update_status(f"Presentation loaded: {os.path.basename(path)}")

    def set_output(self, path):
        self.output_path = os.path.abspath(path)
        self.lbl_output.config(text=self.output_path, fg=self.text_color)

    def select_output(self):
        initial_name = "presentation.exe"
        initial_dir = None
        if self.output_path:
            initial_name = os.path.basename(self.output_path)
            initial_dir = os.path.dirname(self.output_path)
        elif self.ppsx_path:
            initial_name = os.path.splitext(os.path.basename(self.ppsx_path))[0] + ".exe"
            initial_dir = os.path.dirname(self.ppsx_path)
        path = filedialog.asksaveasfilename(
            title="EXEPoint - Save Executable As",
            defaultextension=".exe",
            initialfile=initial_name,
            initialdir=initial_dir,
            filetypes=[("Executable", "*.exe")]
        )
        if path:
            self.set_output(path)
            self.output_custom = True
            self.update_status(f"Output set: {os.path.basename(path)}")

    def select_icon(self):
        path = filedialog.askopenfilename(
            title="EXEPoint - Select App Icon",
            filetypes=[("Icon Files", "*.ico")]
        )
        if path:
            self.icon_path = os.path.abspath(path)
            self.lbl_icon_status.config(text=os.path.basename(path), fg=self.text_color)
            self.update_status(f"Custom icon loaded: {os.path.basename(path)}")

    def update_status(self, text, color="#797775"):
        self.lbl_status.config(text=text, fg=color)
        self.root.update()

    def generate_executable(self):
        if not self.ppsx_path:
            return
        if not self.output_path:
            self.select_output()
            if not self.output_path:
                return

        self.btn_generate.config(state=tk.DISABLED, bg="#A19F9D", text="Processing...")
        self.update_status("Starting presentation embedding...")

        # CORRIGIDO: splitext retorna (nome, extensão), é preciso pegar o índice [1]
        ext_original = os.path.splitext(self.ppsx_path)[1].lower()
        base_name = os.path.splitext(os.path.basename(self.ppsx_path))[0]
        # Tudo é construído numa pasta temporária, apagada no final.
        project_dir = tempfile.mkdtemp(prefix="exepoint_")

        try:

            target_presentation_name = f"presentation{ext_original}"
            shutil.copy(self.ppsx_path, os.path.join(project_dir, target_presentation_name))

            # Escreve o player interno que vai rodar dentro da apresentação final
            self.update_status("Compiling script with native transitions and window hooks...")
            player_script = os.path.join(project_dir, "player.py")
            with open(player_script, "w", encoding="utf-8") as f:
                f.write(f"""
import os
import sys
import time
import tkinter as tk
import pythoncom
import win32api
import win32gui
import win32com.client

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

if __name__ == "__main__":
    # Escudo Invisível Anti-Alt+F4 e Bloqueio de Inputs
    root = tk.Tk()
    root.title("EXEPoint Shield")
    root.attributes('-alpha', 0.0)

    root.bind("<Alt-F4>", lambda e: "break")
    root.bind("<Key>", lambda e: "break")
    root.bind("<Button-1>", lambda e: "break")
    root.protocol("WM_DELETE_WINDOW", lambda: None)

    # CORRIGIDO: inicializadas antes do try para o finally não dar NameError
    ppt_app = None
    presentation = None

    try:
        pythoncom.CoInitialize()
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True

        pres_path = resource_path("{target_presentation_name}")
        presentation = ppt_app.Presentations.Open(pres_path, ReadOnly=True, WithWindow=True)

        slide_show = presentation.SlideShowSettings
        slide_show.ShowType = 3  # Quiosque: tela cheia, sem barra de botoes inferior
        slide_show.LoopUntilStopped = False
        window = slide_show.Run()

        # No modo quiosque o PowerPoint ignora clique/teclado, entao a navegacao
        # e feita aqui chamando View.Next()/Previous() (animacoes nativas continuam).
        NEXT_KEYS = [0x01, 0x27, 0x28, 0x22, 0x20, 0x0D]  # clique, setas dir/baixo, PgDn, espaco, enter
        PREV_KEYS = [0x02, 0x25, 0x26, 0x21, 0x08]        # clique direito, setas esq/cima, PgUp, backspace
        held = set()

        def just_pressed(vk):
            down = (win32api.GetAsyncKeyState(vk) & 0x8000) != 0
            was = vk in held
            if down:
                held.add(vk)
            else:
                held.discard(vk)
            return down and not was

        while True:
            root.update()
            time.sleep(0.03)
            try:
                view = window.View
            except Exception:
                break

            go_next = False
            go_prev = False
            for vk in NEXT_KEYS:
                if just_pressed(vk):
                    go_next = True
            for vk in PREV_KEYS:
                if just_pressed(vk):
                    go_prev = True

            try:
                if win32gui.GetForegroundWindow() != window.HWND:
                    continue
                if go_next:
                    view.Next()
                elif go_prev:
                    view.Previous()
            except Exception:
                pass

    except Exception:
        pass
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        try:
            if ppt_app is not None:
                ppt_app.Quit()
        except Exception:
            pass
        try:
            root.destroy()
        except Exception:
            pass
""")

            # Configura e executa o PyInstaller por baixo dos panos
            self.update_status("Assembling executable bundle via PyInstaller...")
            pyinstaller_cmd = pyinstaller_command() + [
                "--onefile",
                "--windowed",
                f"--name={base_name}",
                f"--distpath={os.path.join(project_dir, 'dist')}",
                f"--workpath={os.path.join(project_dir, 'work')}",
                f"--specpath={project_dir}",
                f"--add-data={target_presentation_name}{os.pathsep}.",
                # CORRIGIDO: garante que o pywin32 seja incluído no player
                "--hidden-import=win32com.client",
                "--hidden-import=pythoncom",
                "--hidden-import=pywintypes",
                "--hidden-import=win32api",
                "--hidden-import=win32gui",
                player_script,
            ]

            if self.icon_path and os.path.exists(self.icon_path):
                pyinstaller_cmd.append(f"--icon={self.icon_path}")

            # CREATE_NO_WINDOW evita abrir uma janela de console durante a compilação
            subprocess.run(
                pyinstaller_cmd,
                cwd=project_dir,
                check=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )

            # Move apenas o .exe para o local escolhido pelo usuário
            generated_exe = os.path.join(project_dir, "dist", f"{base_name}.exe")
            if os.path.exists(generated_exe):
                final_exe = self.output_path
                os.makedirs(os.path.dirname(final_exe), exist_ok=True)
                if os.path.exists(final_exe):
                    os.remove(final_exe)
                shutil.move(generated_exe, final_exe)
                self.update_status("Conversion completed successfully!", "#4CAF50")
                messagebox.showinfo(
                    "EXEPoint",
                    f"Success! Executable saved at:\n{final_exe}"
                )
            else:
                raise FileNotFoundError("PyInstaller finished but the .exe was not found.")

        except Exception as e:
            self.update_status("Conversion failed", "#B7472A")
            messagebox.showerror("EXEPoint Error", f"A problem occurred during file generation:\n{str(e)}")

        finally:
            shutil.rmtree(project_dir, ignore_errors=True)
            # CORRIGIDO: texto igual ao do botão original
            self.btn_generate.config(state=tk.NORMAL, bg=self.main_color, text=BTN_TEXT)


if __name__ == "__main__":
    root = tk.Tk()
    app = EXEPointApp(root)
    root.mainloop()