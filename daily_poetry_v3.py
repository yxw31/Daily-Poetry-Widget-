import tkinter as tk
from tkinter import Menu, messagebox, colorchooser
import requests
import os
import sys
import winreg as reg
import threading
import time
import json
from datetime import datetime

# ================== pywin32 for 点击穿透 ==================
try:
    import win32gui
    import win32con
except ImportError:
    print("正在安装 pywin32...")
    os.system("pip install pywin32 -i https://pypi.tuna.tsinghua.edu.cn/simple")
    import win32gui
    import win32con

CONFIG_FILE = "daily_poetry_config.json"
TOKEN_URL = "https://v2.jinrishici.com/token"
POETRY_URL = "https://v2.jinrishici.com/one.json"
REFRESH_INTERVAL = 3600

class PoetryWidget:
    def __init__(self):
        self.load_config()
        self.locked = False
        self.token = None
        self.hwnd = None

        self.root = tk.Tk()
        self.root.title("每日诗词")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=self.config["bg_color"])
        self.root.attributes("-alpha", self.config["bg_alpha"])   # 背景透明度

        self.width = self.config.get("width", 460)
        self.height = self.config.get("height", 240)
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        self.root.geometry(f"{self.width}x{self.height}+{screen_w - self.width - 60}+{screen_h - self.height - 120}")

        # Canvas 让文字背景透明
        self.canvas = tk.Canvas(self.root, width=self.width, height=self.height,
                                bg=self.config["bg_color"], highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.text_content = self.canvas.create_text(self.width//2, 75, text="正在加载诗词...",
                                                    font=("微软雅黑", self.config["font_size_content"], "bold"),
                                                    fill=self.config["fg_content"], width=self.width-70, justify="center")

        self.text_author = self.canvas.create_text(self.width//2, 150, text="",
                                                   font=("微软雅黑", self.config["font_size_author"]),
                                                   fill=self.config["fg_author"])

        self.text_date = self.canvas.create_text(self.width//2, self.height-35, text="",
                                                 font=("微软雅黑", 10), fill=self.config["fg_date"])

        self.lock_btn = None

        self.root.bind("<Button-3>", self.show_menu)
        self.root.bind("<Button-1>", self.start_drag)
        self.root.bind("<B1-Motion>", self.do_drag)

        self.root.after(300, self.init_hwnd)
        self.refresh_poetry()
        threading.Thread(target=self.auto_refresh, daemon=True).start()

        self.root.mainloop()

    def load_config(self):
        default = {
            "bg_alpha": 0.78,                    # 背景透明度（可调节）
            "bg_color": "#1e1e2e",
            "fg_content": "#e0e0ff",
            "fg_author": "#a0a0cc",
            "fg_date": "#666688",
            "width": 460, "height": 240,
            "display_mode": "full",
            "font_size_content": 22,
            "font_size_author": 15,
            "theme": "dark"
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                self.config = {**default, **loaded}
            except:
                self.config = default
        else:
            self.config = default
        self.save_config()

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except:
            pass

    def init_hwnd(self):
        try:
            self.hwnd = win32gui.GetParent(self.root.winfo_id())
            self.apply_click_through()
        except:
            pass

    def apply_click_through(self):
        if not self.hwnd:
            return
        try:
            exstyle = win32gui.GetWindowLong(self.hwnd, win32con.GWL_EXSTYLE)
            if self.locked:
                exstyle |= (win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED | win32con.WS_EX_NOACTIVATE)
            else:
                exstyle &= ~(win32con.WS_EX_TRANSPARENT | win32con.WS_EX_NOACTIVATE)
            win32gui.SetWindowLong(self.hwnd, win32con.GWL_EXSTYLE, exstyle)
        except:
            pass

    # ================== 诗词获取 ==================
    def get_token(self):
        try:
            r = requests.get(TOKEN_URL, timeout=8)
            if r.status_code == 200:
                self.token = r.json().get("data")
        except:
            pass

    def fetch_poetry(self):
        headers = {"X-User-Token": self.token} if self.token else {}
        try:
            r = requests.get(POETRY_URL, headers=headers, timeout=10)
            if r.status_code == 200 and r.json().get("status") == "success":
                p = r.json()["data"]
                return p.get("content", ""), p.get("origin", {}).get("title", ""), p.get("origin", {}).get("dynasty", ""), p.get("origin", {}).get("author", "")
        except:
            pass
        import random
        fb = [("人生若只如初见，何事秋风悲画扇。", "纳兰性德", "清代", ""), ("明月松间照，清泉石上流。", "王维", "唐代", "")]
        return random.choice(fb)

    def refresh_poetry(self):
        if not self.token:
            self.get_token()
        content, title, dynasty, author = self.fetch_poetry()
        today = datetime.now().strftime("%Y年%m月%d日")

        mode = self.config["display_mode"]
        author_str = f"{title} 〔{dynasty}〕{author}".strip() if mode != "content_only" else ""

        self.canvas.itemconfig(self.text_content, text=content)
        self.canvas.itemconfig(self.text_author, text=author_str)
        self.canvas.itemconfig(self.text_date, text=f"—— {today}  每日诗词")

    def auto_refresh(self):
        while True:
            time.sleep(REFRESH_INTERVAL)
            self.root.after(0, self.refresh_poetry)

    # ================== 菜单 & 小锁解锁 ==================
    def show_menu(self, event):
        if self.locked:
            return
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label="🔄 刷新诗句", command=self.refresh_poetry)
        menu.add_command(label="⚙️ 设置", command=self.open_settings)
        menu.add_command(label="🔒 锁定界面", command=self.lock)
        menu.add_command(label="切换位置", command=self.toggle_position)
        menu.add_separator()
        self.autostart_var = tk.BooleanVar(value=self.is_autostart_enabled())
        menu.add_checkbutton(label="开机自启", variable=self.autostart_var, command=self.toggle_autostart)
        menu.add_separator()
        menu.add_command(label="退出", command=self.quit_app)
        menu.post(event.x_root, event.y_root)

    def create_lock_button(self):
        if self.lock_btn:
            return
        self.lock_btn = tk.Button(self.canvas, text="🔒", font=("微软雅黑", 18, "bold"), bg="#333333", fg="#ffdd00",
                                  relief="flat", bd=0, command=self.unlock, cursor="hand2", width=2, height=1)
        self.lock_btn.place(x=self.width - 55, y=15)

    def remove_lock_button(self):
        if self.lock_btn:
            self.lock_btn.destroy()
            self.lock_btn = None

    def lock(self):
        self.locked = True
        self.apply_click_through()
        self.create_lock_button()

    def unlock(self):
        self.locked = False
        self.apply_click_through()
        self.remove_lock_button()

    def start_drag(self, event):
        if self.locked:
            return
        self._drag_x = event.x
        self._drag_y = event.y

    def do_drag(self, event):
        if self.locked:
            return
        x = self.root.winfo_x() + event.x - self._drag_x
        y = self.root.winfo_y() + event.y - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    def toggle_position(self):
        if self.locked:
            return
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        if self.root.winfo_x() > screen_w // 2:
            self.root.geometry(f"+50+50")
        else:
            self.root.geometry(f"+{screen_w - self.width - 60}+{screen_h - self.height - 120}")

    # ================== 开机自启 ==================
    def get_reg_key(self):
        return reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_SET_VALUE)

    def is_autostart_enabled(self):
        try:
            reg.QueryValueEx(self.get_reg_key(), "DailyPoetryWidget")
            return True
        except:
            return False

    def toggle_autostart(self):
        try:
            key = self.get_reg_key()
            exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
            if self.autostart_var.get():
                reg.SetValueEx(key, "DailyPoetryWidget", 0, reg.REG_SZ, f'"{exe_path}"')
            else:
                reg.DeleteValue(key, "DailyPoetryWidget")
        except:
            pass

    def quit_app(self):
        if messagebox.askokcancel("退出", "确定退出每日诗词吗？"):
            self.root.destroy()

    # ================== 设置窗口 ==================
    def open_settings(self):
        if self.locked:
            return
        win = tk.Toplevel(self.root)
        win.title("设置 - 每日诗词")
        win.geometry("480x680")
        win.grab_set()

        tk.Label(win, text="主题风格", font=("微软雅黑", 11)).pack(pady=(15,5))
        tk.Button(win, text="🌙 深色主题", width=15, command=lambda: self.apply_theme("dark")).pack(pady=5)
        tk.Button(win, text="☀️ 浅色主题", width=15, command=lambda: self.apply_theme("light")).pack(pady=5)

        tk.Label(win, text="背景透明度（0.1~1.0）").pack(pady=(15,5))
        bg_scale = tk.Scale(win, from_=0.1, to=1.0, resolution=0.01, orient="horizontal", length=380,
                            command=lambda v: [self.config.update({"bg_alpha": float(v)}), self.root.attributes("-alpha", float(v))])
        bg_scale.set(self.config["bg_alpha"])
        bg_scale.pack()

        tk.Label(win, text="诗句字体大小").pack(pady=(15,5))
        fs_content = tk.Scale(win, from_=14, to=32, resolution=1, orient="horizontal", length=380,
                              command=lambda v: [self.config.update({"font_size_content": int(v)}), self.canvas.itemconfig(self.text_content, font=("微软雅黑", int(v), "bold"))])
        fs_content.set(self.config["font_size_content"])
        fs_content.pack()

        tk.Label(win, text="作者字体大小").pack(pady=(10,5))
        fs_author = tk.Scale(win, from_=10, to=22, resolution=1, orient="horizontal", length=380,
                             command=lambda v: [self.config.update({"font_size_author": int(v)}), self.canvas.itemconfig(self.text_author, font=("微软雅黑", int(v)))])
        fs_author.set(self.config["font_size_author"])
        fs_author.pack()

        tk.Label(win, text="显示内容").pack(pady=(15,8))
        mode_var = tk.StringVar(value=self.config["display_mode"])
        tk.Radiobutton(win, text="仅显示诗句", variable=mode_var, value="content_only").pack(anchor="w", padx=80)
        tk.Radiobutton(win, text="诗句 + 作者", variable=mode_var, value="content_author").pack(anchor="w", padx=80)
        tk.Radiobutton(win, text="诗句 + 作者 + 标题", variable=mode_var, value="full").pack(anchor="w", padx=80)

        tk.Label(win, text="颜色微调", font=("微软雅黑", 11)).pack(pady=(20,8))
        color_frame = tk.Frame(win)
        color_frame.pack()
        def choose(key):
            c = colorchooser.askcolor(title=f"选择 {key}")[1]
            if c:
                self.config[key] = c
                if key == "bg_color":
                    self.canvas.configure(bg=c)
                    self.root.configure(bg=c)
                else:
                    fill_key = {"fg_content": self.text_content, "fg_author": self.text_author, "fg_date": self.text_date}
                    self.canvas.itemconfig(fill_key[key], fill=c)

        tk.Button(color_frame, text="背景色", command=lambda: choose("bg_color")).grid(row=0, column=0, padx=8, pady=5)
        tk.Button(color_frame, text="诗句色", command=lambda: choose("fg_content")).grid(row=0, column=1, padx=8, pady=5)
        tk.Button(color_frame, text="作者色", command=lambda: choose("fg_author")).grid(row=0, column=2, padx=8, pady=5)
        tk.Button(color_frame, text="日期色", command=lambda: choose("fg_date")).grid(row=0, column=3, padx=8, pady=5)

        def save_all():
            self.config["display_mode"] = mode_var.get()
            self.save_config()
            self.refresh_poetry()
            messagebox.showinfo("保存成功", "设置已保存并生效！")
            win.destroy()

        tk.Button(win, text="💾 保存并应用所有设置", font=("微软雅黑", 12, "bold"), bg="#4a90e2", fg="white", height=2, command=save_all).pack(pady=30)

    def apply_theme(self, theme):
        self.config["theme"] = theme
        if theme == "dark":
            self.config.update({"bg_color": "#1e1e2e", "fg_content": "#e0e0ff", "fg_author": "#a0a0cc", "fg_date": "#666688"})
        else:
            self.config.update({"bg_color": "#f0f0f5", "fg_content": "#2c2c2c", "fg_author": "#444444", "fg_date": "#777777"})
        self.canvas.configure(bg=self.config["bg_color"])
        self.root.configure(bg=self.config["bg_color"])
        self.canvas.itemconfig(self.text_content, fill=self.config["fg_content"])
        self.canvas.itemconfig(self.text_author, fill=self.config["fg_author"])
        self.canvas.itemconfig(self.text_date, fill=self.config["fg_date"])

# ================== 启动 ==================
if __name__ == "__main__":
    try:
        import requests
    except ImportError:
        os.system("pip install requests -i https://pypi.tuna.tsinghua.edu.cn/simple")
        import requests
    PoetryWidget()