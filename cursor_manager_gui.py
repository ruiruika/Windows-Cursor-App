import os
import ctypes
import winreg
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Windows API Constants
SPI_SETCURSORS = 0x0057
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDCHANGE = 0x02
REG_CURSOR_PATH = r"Control Panel\Cursors"

# Global Mapping: Supports both Japanese and common English filename keywords
CURSOR_REG_MAPPING = {
    "Arrow": ["通常", "1_通常", "通常.ani", "通常.cur", "Normal", "Default", "Arrow"],
    "Help": ["ヘルプの選択", "2_ヘルプの選択", "ヘルプ.ani", "ヘルプ.cur", "Help", "Question"],
    "AppStarting": ["バックグラウンドで作業中", "3_バックグラウンドで作業中", "バックグラウンド.ani", "バックグラウンド.cur", "Working", "Background"],
    "Wait": ["待ち状態", "4_待ち状態", "待ち状態.ani", "待ち状態.cur", "Wait", "Busy"],
    "Crosshair": ["領域選択", "5_領域選択", "領域選択.ani", "領域選択.cur", "Cross", "Precision"],
    "IBeam": ["テキスト選択", "6_テキスト選択", "テキスト選択.ani", "テキスト選択.cur", "Text", "IBeam"],
    "Handwriting": ["手書き", "7_手書き", "手書き.ani", "手書き.cur", "Handwriting", "Pen"],
    "No": ["利用不可", "8_利用不可", "利用不可.ani", "利用不可.cur", "Unavailable", "No", "Denied"],
    "SizeNS": ["上下に拡大,縮小", "9_上下に拡大,縮小", "上下.ani", "上下.cur", "上下", "NS", "NorthSouth", "Vertical"],
    "SizeWE": ["左右に拡大,縮小", "10_左右に拡大,縮小", "左右.ani", "左右.cur", "左右", "WE", "WestEast", "Horizontal"],
    "SizeNWSE": ["斜めに拡大,縮小1", "11_斜めに拡大,縮小1", "斜め.ani", "斜め.cur", "斜め", "NWSE", "Diagonal1"],
    "SizeNESW": ["斜めに拡大,縮小2", "12_斜めに拡大,縮小2", "斜め2.ani", "斜め2.cur", "斜め2", "NESW", "Diagonal2"],
    "SizeAll": ["移動", "13_移動", "移動.ani", "移動.cur", "Move", "SizeAll"],
    "UpArrow": ["代替選択", "代替選択.ani", "代替選択.cur", "Alternate", "UpArrow"],
    "Hand": ["リンクの選択", "15_リンクの選択", "リンクの選択.ani", "リンクの選択.cur", "リンク", "Link", "Hand"],
    "LocationSelect": ["位置選択", "Location", "Pin", "GPS", "Point"],
    "PersonSelect": ["ユーザー選択", "Person", "User", "People"],
}

LANGUAGES = {
    "JP": {
        "title": "マウスカーソル一括変更ツール (Global対応版)",
        "btn_apply": "カーソルセットを適用する",
        "btn_reset": "標準に戻す",
        "btn_settings": "言語設定",
        "msg_success": "適用完了: {} 個のカーソルを更新しました。",
        "msg_reset": "Windows標準に戻しました。",
        "confirm_reset": "すべてのカーソルをWindows標準に戻しますか？",
        "lang_name": "日本語 (JP)",
        "setting_title": "設定",
        "select_lang": "言語選択:",
        "current_path": "現在のパス:",
    },
    "EN": {
        "title": "Mouse Cursor Bulk Applier (Global Version)",
        "btn_apply": "Apply Cursor Set",
        "btn_reset": "Restore Defaults",
        "btn_settings": "Language Settings",
        "msg_success": "Success: Applied {} cursor types from {}.",
        "msg_reset": "Restored to system defaults.",
        "confirm_reset": "Reset all cursors to Windows default values?",
        "lang_name": "English (EN)",
        "setting_title": "Settings",
        "select_lang": "Language:",
        "current_path": "Current Path:",
    }
}

class CursorManagerApp:
    def __init__(self, root):
        self.root = root
        self.settings_file = "settings.json"
        
        # UI configuration
        self.root.geometry("600x300")
        self.load_settings()
        self.lang = LANGUAGES.get(self.current_lang, LANGUAGES["JP"])
        self.root.title(self.lang["title"])
        
        # Initial language setup if first time
        if not os.path.exists(self.settings_file):
            self.first_time_setup()
        
        self.setup_ui()

    def first_time_setup(self):
        self.root.withdraw()
        lang_win = tk.Toplevel()
        lang_win.title("Language / 言語選択")
        lang_win.geometry("300x150")
        lang_win.grab_set()

        # Exit if window is closed without selection
        self.cancelled = False
        def on_close():
            self.cancelled = True
            self.root.destroy()
            exit()
        lang_win.protocol("WM_DELETE_WINDOW", on_close)
        
        label = ttk.Label(lang_win, text="Please select a language\n言語を選択してください", padding=10)
        label.pack()
        
        lang_var = tk.StringVar(value="JP")
        combo = ttk.Combobox(lang_win, textvariable=lang_var, values=["JP", "EN"], state="readonly")
        combo.pack(pady=10)
        
        def set_lang():
            self.current_lang = lang_var.get()
            self.lang = LANGUAGES[self.current_lang]
            self.save_settings()
            lang_win.destroy()
            
        ttk.Button(lang_win, text="OK", command=set_lang).pack()
        self.root.wait_window(lang_win)
        
        # Check if root still exists and we didn't cancel
        try:
            if not getattr(self, "cancelled", False) and self.root.winfo_exists():
                self.root.deiconify()
        except (tk.TclError, AttributeError):
            pass

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.current_lang = data.get("lang", "JP")
                    self.applied_path = data.get("last_path", "-")
            except:
                self.current_lang = "JP"
                self.applied_path = "-"
        else:
            self.current_lang = "JP"
            self.applied_path = "-"

    def save_settings(self):
        data = {"lang": self.current_lang, "last_path": self.applied_path}
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def setup_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        # Responsive Layout Configuration
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Header Frame (Path and Language Toggle)
        header_frame = ttk.Frame(self.root, padding="15")
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.columnconfigure(0, weight=1)

        self.path_label = ttk.Label(header_frame, text=f"{self.lang['current_path']} {self.applied_path}", wraplength=450)
        self.path_label.grid(row=0, column=0, sticky="w")
        
        ttk.Button(header_frame, text=self.lang["btn_settings"], command=self.open_settings).grid(row=0, column=1, sticky="e", padx=(10, 0))

        # Main Button Area
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=1, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Big Apply Button
        apply_btn = tk.Button(main_frame, text=self.lang["btn_apply"], font=("Meiryo", 14, "bold"), 
                              bg="#4a90e2", fg="white", activebackground="#357abd", relief="flat",
                              command=self.select_and_apply)
        apply_btn.grid(row=0, column=0, sticky="nsew", pady=10)

        # Footer Area (Reset)
        footer_frame = ttk.Frame(self.root, padding="15")
        footer_frame.grid(row=2, column=0, sticky="ew")
        footer_frame.columnconfigure(0, weight=1)

        ttk.Button(footer_frame, text=self.lang["btn_reset"], command=self.reset_to_default).grid(row=0, column=0, sticky="e")

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title(self.lang["setting_title"])
        win.geometry("300x150")
        win.grab_set()
        
        frame = ttk.Frame(win, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=self.lang["select_lang"]).grid(row=0, column=0, sticky=tk.W, pady=5)
        lang_var = tk.StringVar(value=self.current_lang)
        lang_combo = ttk.Combobox(frame, textvariable=lang_var, values=["JP", "EN"], state="readonly")
        lang_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=10)
        
        def save_and_close():
            self.current_lang = lang_var.get()
            self.lang = LANGUAGES[self.current_lang]
            self.save_settings()
            self.setup_ui()
            win.destroy()
            
        ttk.Button(frame, text="OK", command=save_and_close).grid(row=1, column=0, columnspan=2, pady=20)

    def select_and_apply(self):
        file_path = filedialog.askopenfilename(
            title=self.lang["btn_apply"],
            filetypes=[("Cursor files", "*.ani;*.cur")]
        )
        if not file_path:
            return

        folder_path = os.path.dirname(file_path)
        self.applied_path = folder_path
        self.apply_folder(folder_path)
        self.save_settings()
        self.path_label.config(text=f"{self.lang['current_path']} {self.applied_path}")

    def update_registry(self, name, value):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_CURSOR_PATH, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
            winreg.CloseKey(key)
            return True
        except: return False

    def trigger_refresh(self):
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETCURSORS, 0, 0, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)

    def apply_folder(self, folder_path):
        all_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.ani', '.cur'))]
        
        applied_mapping = {}
        used_files = set()

        # Step 1: Exact matches (priority)
        for reg_name, keywords in CURSOR_REG_MAPPING.items():
            for filename in all_files:
                if filename in used_files: continue
                basename = os.path.splitext(filename)[0].lower()
                if any(kw.lower() == basename for kw in keywords):
                    applied_mapping[reg_name] = os.path.join(folder_path, filename)
                    used_files.add(filename)
                    break

        # Step 2: Partial matches for remaining
        for reg_name, keywords in CURSOR_REG_MAPPING.items():
            if reg_name in applied_mapping: continue
            for filename in all_files:
                if filename in used_files: continue
                basename = os.path.splitext(filename)[0].lower()
                # Check for "contains" match, prioritize longer keywords
                sorted_kws = sorted(keywords, key=len, reverse=True)
                if any(kw.lower() in basename for kw in sorted_kws):
                    applied_mapping[reg_name] = os.path.join(folder_path, filename)
                    used_files.add(filename)
                    break
        
        applied_count = 0
        for reg_name in CURSOR_REG_MAPPING.keys():
            val = applied_mapping.get(reg_name, "")
            if self.update_registry(reg_name, val):
                if val: applied_count += 1
        
        self.trigger_refresh()
        messagebox.showinfo(self.lang["setting_title"], self.lang["msg_success"].format(applied_count, os.path.basename(folder_path)))

    def reset_to_default(self):
        if messagebox.askyesno(self.lang["title"], self.lang["confirm_reset"]):
            for reg_name in set(CURSOR_REG_MAPPING.keys()):
                self.update_registry(reg_name, "")
            self.trigger_refresh()
            self.applied_path = "-"
            self.save_settings()
            self.path_label.config(text=f"{self.lang['current_path']} {self.applied_path}")
            messagebox.showinfo(self.lang["title"], self.lang["msg_reset"])

if __name__ == "__main__":
    root = tk.Tk()
    app = CursorManagerApp(root)
    root.mainloop()
