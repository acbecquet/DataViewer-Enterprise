"""
Translation Tool - GUI Version
Supports Excel, PowerPoint, PDF, JPG, and PNG files.
Translates between any of 20 supported languages using Claude AI.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import base64
import json
from pathlib import Path
from anthropic import Anthropic

# ── Module imports ─────────────────────────────────────────────────────────────
_import_errors = []

try:
    from translate_excel import (
        translate_excel,
        generate_output_filename as generate_excel_output,
    )
except ImportError as e:
    _import_errors.append(f"translate_excel: {e}")

try:
    from translate_powerpoint import (
        translate_powerpoint,
        generate_output_filename as generate_ppt_output,
    )
except ImportError as e:
    _import_errors.append(f"translate_powerpoint: {e}")

try:
    from translate_image import translate_image_file, generate_output_path as generate_img_output
    IMAGE_MODULE_AVAILABLE = True
except ImportError as e:
    IMAGE_MODULE_AVAILABLE = False
    _import_errors.append(f"translate_image: {e}")

MODULES_AVAILABLE = len(_import_errors) == 0
import_error = "; ".join(_import_errors)

# ── Language list ──────────────────────────────────────────────────────────────
LANGUAGES = [
    "Chinese (Simplified)",
    "Chinese (Traditional)",
    "English",
    "Arabic",
    "Dutch",
    "French",
    "German",
    "Hindi",
    "Indonesian",
    "Italian",
    "Japanese",
    "Korean",
    "Polish",
    "Portuguese (Brazilian)",
    "Russian",
    "Spanish",
    "Swedish",
    "Thai",
    "Turkish",
    "Vietnamese",
]

# File-type groups
_OFFICE_EXTS  = {'.xlsx', '.xls', '.pptx', '.ppt'}
_IMAGE_EXTS   = {'.jpg', '.jpeg', '.png'}
_PDF_EXTS     = {'.pdf'}
_ALL_EXTS     = _OFFICE_EXTS | _IMAGE_EXTS | _PDF_EXTS


def _file_label(ext: str) -> str:
    if ext in ('.xlsx', '.xls'):
        return "[Excel]"
    if ext in ('.pptx', '.ppt'):
        return "[PPT]"
    if ext == '.pdf':
        return "[PDF]"
    if ext in ('.jpg', '.jpeg', '.png'):
        return "[IMG]"
    return "[?]"


# ── Helpers ────────────────────────────────────────────────────────────────────

class TextRedirector:
    def __init__(self, log_callback):
        self.log_callback = log_callback

    def write(self, text):
        if text.strip():
            self.log_callback(text.strip(), "info")

    def flush(self):
        pass


class APIKeyManager:
    def __init__(self):
        self.config_dir  = Path.home() / '.document_translator'
        self.config_file = self.config_dir / 'config.dat'
        self.config_dir.mkdir(exist_ok=True)

    def save_api_key(self, api_key: str) -> bool:
        try:
            encoded = base64.b64encode(api_key.encode()).decode()
            with open(self.config_file, 'w') as f:
                json.dump({'api_key': encoded}, f)
            return True
        except Exception as e:
            print(f"Error saving API key: {e}")
            return False

    def load_api_key(self) -> str | None:
        try:
            if not self.config_file.exists():
                return None
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            encoded = config.get('api_key')
            if encoded:
                return base64.b64decode(encoded.encode()).decode()
            return None
        except Exception as e:
            print(f"Error loading API key: {e}")
            return None


# ── Main GUI ───────────────────────────────────────────────────────────────────

class TranslatorGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Document Translator")

        try:
            self.root.iconbitmap(self.get_resource_path('resources/ccell_icon.ico'))
        except Exception:
            pass

        self.api_key_manager = APIKeyManager()
        self.api_key          = None
        self.client           = None
        self.selected_files   = []
        self.source_language  = tk.StringVar(value="Chinese (Simplified)")
        self.target_language  = tk.StringVar(value="English")
        self.is_translating   = False

        window_width, window_height = 700, 660
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(False, False)

        saved_key = self.api_key_manager.load_api_key()
        if saved_key:
            self.api_key = saved_key
            try:
                self.client = Anthropic(api_key=self.api_key)
            except Exception:
                self.api_key = None

        self.create_widgets()
        self.root.update_idletasks()
        self.center_window(window_width, window_height)

        if not MODULES_AVAILABLE:
            self.log_message(f"ERROR: Required modules not found: {import_error}", "error")
            self.translate_button.config(state="disabled")
        elif not self.client:
            self.root.after(100, self.show_api_key_dialog)

    # ── Utility ──────────────────────────────────────────────────────────────

    def get_resource_path(self, relative_path: str) -> str:
        try:
            base = sys._MEIPASS
        except AttributeError:
            base = os.path.abspath(".")
        return os.path.join(base, relative_path)

    def center_window(self, width: int, height: int, window: tk.Toplevel | None = None):
        w = window or self.root
        sw, sh = w.winfo_screenwidth(), w.winfo_screenheight()
        w.geometry(f"{width}x{height}+{(sw - width) // 2}+{(sh - height) // 2}")

    # ── Widget construction ───────────────────────────────────────────────────

    def create_widgets(self):
        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Change API Key", command=self.show_api_key_dialog)
        settings_menu.add_separator()
        settings_menu.add_command(label="Exit", command=self.root.quit)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

        # Header
        header = tk.Frame(self.root, bg="#2c3e50", height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="Document Translator", font=("Arial", 24, "bold"),
                 bg="#2c3e50", fg="white").pack(pady=(16, 2))
        tk.Label(header,
                 text="Translate Excel, PowerPoint, PDF, and image files between any supported language",
                 font=("Arial", 9), bg="#2c3e50", fg="#ecf0f1").pack()

        # Content frame
        cf = tk.Frame(self.root, padx=20, pady=15)
        cf.pack(fill="both", expand=True)

        # Language selectors
        lang_frame = tk.LabelFrame(cf, text="Language", padx=10, pady=10)
        lang_frame.pack(fill="x", pady=(0, 10))
        lang_frame.columnconfigure(0, weight=1)
        lang_frame.columnconfigure(1, weight=0)
        lang_frame.columnconfigure(2, weight=1)

        src_col = tk.Frame(lang_frame)
        src_col.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        tk.Label(src_col, text="Source Language", font=("Arial", 10, "bold")).pack(anchor="w")
        self.source_combo = ttk.Combobox(src_col, textvariable=self.source_language,
                                          values=LANGUAGES, state="readonly",
                                          font=("Arial", 11), width=26)
        self.source_combo.pack(fill="x", pady=(4, 0))

        tk.Label(lang_frame, text="→", font=("Arial", 18), fg="#2c3e50").grid(
            row=0, column=1, padx=8, pady=(16, 0))

        tgt_col = tk.Frame(lang_frame)
        tgt_col.grid(row=0, column=2, sticky="ew", padx=(5, 0))
        tk.Label(tgt_col, text="Target Language", font=("Arial", 10, "bold")).pack(anchor="w")
        self.target_combo = ttk.Combobox(tgt_col, textvariable=self.target_language,
                                          values=LANGUAGES, state="readonly",
                                          font=("Arial", 11), width=26)
        self.target_combo.pack(fill="x", pady=(4, 0))

        # File list
        file_frame = tk.LabelFrame(cf, text="Selected Files", padx=10, pady=10)
        file_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.file_listbox = tk.Listbox(file_frame, height=7, font=("Arial", 10))
        self.file_listbox.pack(fill="both", expand=True, side="left")
        sb = tk.Scrollbar(file_frame, command=self.file_listbox.yview)
        sb.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=sb.set)

        # Buttons
        btn_frame = tk.Frame(cf)
        btn_frame.pack(fill="x", pady=(0, 10))

        self.select_button = tk.Button(btn_frame, text="Select Files",
                                        command=self.select_files,
                                        bg="#3498db", fg="white",
                                        font=("Arial", 11, "bold"),
                                        padx=20, pady=10, cursor="hand2")
        self.select_button.pack(side="left", padx=(0, 10))

        self.clear_button = tk.Button(btn_frame, text="Clear",
                                       command=self.clear_files,
                                       bg="#95a5a6", fg="white",
                                       font=("Arial", 11),
                                       padx=20, pady=10, cursor="hand2")
        self.clear_button.pack(side="left")

        self.translate_button = tk.Button(btn_frame, text="Translate",
                                           command=self.start_translation,
                                           bg="#27ae60", fg="white",
                                           font=("Arial", 11, "bold"),
                                           padx=20, pady=10, cursor="hand2",
                                           state="disabled")
        self.translate_button.pack(side="right")

        # Progress log
        log_frame = tk.LabelFrame(cf, text="Progress", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8,
                                                   font=("Consolas", 9),
                                                   wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config("error",   foreground="red")
        self.log_text.tag_config("success", foreground="green")
        self.log_text.tag_config("info",    foreground="blue")

        # Status bar
        self.status_label = tk.Label(self.root, text="Ready",
                                      bd=1, relief="sunken", anchor="w",
                                      font=("Arial", 9))
        self.status_label.pack(side="bottom", fill="x")

    # ── Dialogs ───────────────────────────────────────────────────────────────

    def show_about(self):
        messagebox.showinfo(
            "About Document Translator",
            "Document Translator v2.1.0\n\n"
            "Supported formats:\n"
            "  • Excel (.xlsx, .xls)\n"
            "  • PowerPoint (.pptx, .ppt) — incl. tables\n"
            "  • PDF (.pdf)\n"
            "  • Images (.jpg, .png)\n\n"
            "Powered by Claude AI\n\n"
            "\u00a9 2025 SDR"
        )

    def show_api_key_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("API Key Required")
        try:
            dialog.iconbitmap(self.get_resource_path('resources/ccell_icon.ico'))
        except Exception:
            pass

        dw, dh = 500, 280
        dialog.geometry(f"{dw}x{dh}")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_window(dw, dh, dialog)

        tk.Label(dialog, text="Anthropic API Key Required",
                 font=("Arial", 14, "bold")).pack(pady=(20, 10))
        tk.Label(dialog, text="Please enter your Anthropic API key to use the translator.",
                 font=("Arial", 10)).pack()

        link = tk.Label(dialog, text="Get your API key from: https://console.anthropic.com",
                        font=("Arial", 9), fg="blue", cursor="hand2")
        link.pack(pady=(5, 20))
        link.bind("<Button-1>", lambda e: __import__('webbrowser').open("https://console.anthropic.com"))

        kf = tk.Frame(dialog)
        kf.pack(pady=10)
        tk.Label(kf, text="API Key:", font=("Arial", 10)).pack(side="left", padx=(0, 10))
        key_entry = tk.Entry(kf, width=40, font=("Arial", 10), show="*")
        key_entry.pack(side="left")
        key_entry.focus()

        show_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dialog, text="Show API key", variable=show_var,
                       command=lambda: key_entry.config(show="" if show_var.get() else "*"),
                       font=("Arial", 9)).pack()

        def save_key():
            key = key_entry.get().strip()
            if not key:
                messagebox.showerror("Error", "Please enter an API key", parent=dialog)
                return
            if not key.startswith("sk-ant-"):
                messagebox.showerror("Error",
                    "Invalid API key format. Must start with 'sk-ant-'", parent=dialog)
                return
            try:
                self.client  = Anthropic(api_key=key)
                self.api_key = key
                if self.api_key_manager.save_api_key(key):
                    messagebox.showinfo("Success", "API key saved successfully!", parent=dialog)
                else:
                    messagebox.showwarning("Warning",
                        "API key is valid but could not be saved to disk.", parent=dialog)
                dialog.destroy()
                self.log_message("API key configured successfully", "success")
            except Exception as e:
                messagebox.showerror("Error", f"Invalid API key: {e}", parent=dialog)

        key_entry.bind("<Return>", lambda e: save_key())

        bf = tk.Frame(dialog)
        bf.pack(pady=20)
        tk.Button(bf, text="Save", command=save_key,
                  bg="#27ae60", fg="white", font=("Arial", 10, "bold"),
                  padx=20, pady=5).pack(side="left", padx=5)
        tk.Button(bf, text="Cancel", command=dialog.destroy,
                  bg="#95a5a6", fg="white", font=("Arial", 10),
                  padx=20, pady=5).pack(side="left", padx=5)

    # ── File management ───────────────────────────────────────────────────────

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select files to translate",
            filetypes=[
                ("All supported files", "*.xlsx *.xls *.pptx *.ppt *.pdf *.jpg *.jpeg *.png"),
                ("Office files",        "*.xlsx *.xls *.pptx *.ppt"),
                ("Excel files",         "*.xlsx *.xls"),
                ("PowerPoint files",    "*.pptx *.ppt"),
                ("PDF files",           "*.pdf"),
                ("Image files",         "*.jpg *.jpeg *.png"),
                ("All files",           "*.*"),
            ]
        )
        if file_paths:
            self.selected_files = list(file_paths)
            self.update_file_list()
            self.log_message(f"Selected {len(file_paths)} file(s)", "info")
            if self.client and not _import_errors:
                self.translate_button.config(state="normal")

    def clear_files(self):
        self.selected_files = []
        self.update_file_list()
        self.translate_button.config(state="disabled")
        self.log_message("File list cleared", "info")

    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for path in self.selected_files:
            ext  = os.path.splitext(path)[1].lower()
            name = os.path.basename(path)
            self.file_listbox.insert(tk.END, f"{_file_label(ext)} {name}")

    def log_message(self, message: str, tag: str = "info"):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update()

    # ── Translation ───────────────────────────────────────────────────────────

    def start_translation(self):
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select files to translate")
            return
        if not self.client:
            messagebox.showerror("Error", "API key not configured")
            self.show_api_key_dialog()
            return

        src = self.source_language.get()
        tgt = self.target_language.get()
        if src == tgt:
            messagebox.showwarning("Same Language",
                "Source and target languages are the same. Please choose different languages.")
            return

        if not messagebox.askyesno("Confirm Translation",
                                   f"Translate {len(self.selected_files)} file(s)\n"
                                   f"from {src} to {tgt}?"):
            return

        self.select_button.config(state="disabled")
        self.clear_button.config(state="disabled")
        self.translate_button.config(state="disabled")
        self.source_combo.config(state="disabled")
        self.target_combo.config(state="disabled")
        self.status_label.config(text="Translating...")

        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

        self.is_translating = True
        threading.Thread(target=self.translate_files, daemon=True).start()

    def translate_files(self):
        src   = self.source_language.get()
        tgt   = self.target_language.get()
        total = len(self.selected_files)
        ok, fail = 0, 0

        self.log_message(f"Starting translation of {total} file(s): {src} → {tgt}\n", "info")
        os.environ["ANTHROPIC_API_KEY"] = self.api_key

        old_stdout, sys.stdout = sys.stdout, TextRedirector(self.log_message)

        try:
            for idx, file_path in enumerate(self.selected_files, 1):
                name = os.path.basename(file_path)
                ext  = os.path.splitext(file_path)[1].lower()
                self.log_message(f"[{idx}/{total}] Processing: {name}", "info")

                try:
                    if ext in ('.xlsx', '.xls'):
                        out = generate_excel_output(file_path, self.client, src, tgt)
                        self.log_message(f"  Output: {os.path.basename(out)}", "info")
                        translate_excel(file_path, out, src, tgt, self.client)

                    elif ext in ('.pptx', '.ppt'):
                        out = generate_ppt_output(file_path, self.client, src, tgt)
                        self.log_message(f"  Output: {os.path.basename(out)}", "info")
                        translate_powerpoint(file_path, out, src, tgt, self.client)

                    elif ext in ('.pdf', '.jpg', '.jpeg', '.png'):
                        if not IMAGE_MODULE_AVAILABLE:
                            self.log_message(
                                "  ERROR: translate_image module not available", "error")
                            fail += 1
                            continue
                        out = generate_img_output(file_path)
                        self.log_message(f"  Output: {os.path.basename(out)}", "info")
                        translate_image_file(file_path, out, src, tgt, self.client)

                    else:
                        self.log_message("  ERROR: Unsupported file type", "error")
                        fail += 1
                        continue

                    self.log_message("  \u2713 SUCCESS\n", "success")
                    ok += 1

                except Exception as e:
                    self.log_message(f"  \u2717 FAILED: {e}\n", "error")
                    fail += 1

        finally:
            sys.stdout = old_stdout

        self.log_message("=" * 50, "info")
        self.log_message("TRANSLATION COMPLETE", "info")
        self.log_message(f"Total: {total} | Success: {ok} | Failed: {fail}", "info")
        self.log_message("=" * 50, "info")

        self.root.after(0, self.translation_complete, ok, fail)

    def translation_complete(self, ok: int, fail: int):
        self.select_button.config(state="normal")
        self.clear_button.config(state="normal")
        self.translate_button.config(state="normal")
        self.source_combo.config(state="readonly")
        self.target_combo.config(state="readonly")
        self.status_label.config(text="Ready")
        self.is_translating = False

        if fail == 0:
            messagebox.showinfo("Success", f"All {ok} file(s) translated successfully!")
        else:
            messagebox.showwarning("Completed with Errors",
                                   f"Translation completed:\n{ok} succeeded, {fail} failed")


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    TranslatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
