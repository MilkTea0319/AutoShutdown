import tkinter as tk
from tkinter import messagebox
import time
import math
import datetime
import os
import sys
import json
import socket
import threading
import ctypes
from ctypes import wintypes
import pystray
from PIL import Image, ImageDraw

# --- 設定區 ---
APP_NAME = "自動關機"
PORT_ID = 53117
CONFIG_DIR = os.path.join(os.environ['APPDATA'], "AutoShutdown")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
ICON_FILENAME = "image/icon.ico"

# 深色模式配色表
COLORS = {
    "bg": "#202020",
    "fg": "#e0e0e0",
    "entry_bg": "#3c3f41",
    "entry_fg": "#ffffff",
    "entry_focus": "#00bcd4",
    "border_active": "#00bcd4",
    "border_inactive": "#424242",
    "btn_start": "#2e7d32",
    "btn_cancel": "#c62828",
    "status_bg": "#004d40",
    "status_fg": "#80cbc4",
    "clock_face": "#1e1e1e",
    "clock_border": "#00bcd4",
    "clock_tick": "#808080",
    "hand_target": "#ff1744",
    "hand_hour": "#ffffff",
    "hand_min": "#ffffff",
    "hand_sec": "#2979ff"
}

FONT_FAMILY = "微軟正黑體"
FONT_INPUT = "Arial"


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def check_single_instance():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(('localhost', PORT_ID))
        return s
    except socket.error:
        user32 = ctypes.windll.user32
        h_wnd = user32.FindWindowW(None, APP_NAME)
        if h_wnd:
            user32.ShowWindow(h_wnd, 9)
            user32.SetForegroundWindow(h_wnd)
        sys.exit(0)


def load_config():
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {"mode": 1, "cd_h": "0", "cd_m": "0", "sp_h": "0", "sp_m": "0", "skip_warning": False}


def save_config(data):
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def enable_high_dpi():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class ShutdownApp:
    def __init__(self, root, socket_obj):
        self.root = root
        self.socket_obj = socket_obj

        current_dpi = self.root.winfo_fpixels('1i')
        self.scale = current_dpi / 96.0
        if self.scale < 1.0: self.scale = 1.0

        self.root.title(APP_NAME)

        base_w, base_h = 450, 720
        scaled_w = int(base_w * self.scale)
        scaled_h = int(base_h * self.scale)
        self.root.geometry(f"{scaled_w}x{scaled_h}")
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS["bg"])

        try:
            self.root.iconbitmap(resource_path(ICON_FILENAME))
        except:
            pass

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Unmap>", self.on_window_minimize)
        self.config = load_config()
        self.is_running = False
        self.target_time = None
        self.preview_time = None
        self.mode_var = tk.IntVar(value=self.config.get("mode", 1))
        self.tray_icon = None

        self.fonts = {
            "ui": (FONT_FAMILY, int(12 * self.scale)),
            "ui_bold": (FONT_FAMILY, int(12 * self.scale), "bold"),
            "input": (FONT_INPUT, int(13 * self.scale), "bold"),
            "status": (FONT_FAMILY, int(13 * self.scale), "bold"),
            "title_frame": (FONT_FAMILY, int(11 * self.scale)),
            "btn_big": (FONT_FAMILY, int(14 * self.scale), "bold"),
            "warning_title": (FONT_FAMILY, int(16 * self.scale), "bold"),
            "warning_text": (FONT_FAMILY, int(12 * self.scale))
        }

        # --- UI 佈局 ---

        clock_size = int(280 * self.scale)
        self.canvas = tk.Canvas(root, width=clock_size, height=clock_size, bg=COLORS["bg"], highlightthickness=0)
        self.canvas.pack(pady=(int(20 * self.scale), int(10 * self.scale)))

        self.frame_status = tk.Frame(root, bg=COLORS["status_bg"], bd=0)
        self.frame_status.pack(fill="x", padx=int(25 * self.scale), pady=int(10 * self.scale))

        self.lbl_status = tk.Label(self.frame_status, text="準備就緒",
                                   fg=COLORS["status_fg"], bg=COLORS["status_bg"],
                                   font=self.fonts["status"])
        self.lbl_status.pack(pady=int(10 * self.scale))

        # 3. 控制區
        pad_x_outer = int(20 * self.scale)
        pad_y_inner = int(15 * self.scale)
        border_pad = int(3 * self.scale)

        # --- 倒數模式 ---
        self.wrap_cd = tk.Frame(root, bg=COLORS["border_inactive"], padx=border_pad, pady=border_pad)
        self.wrap_cd.pack(fill="x", padx=pad_x_outer, pady=int(5 * self.scale))

        self.group_cd = tk.LabelFrame(self.wrap_cd, text=" ⏳ 倒數計時 ",
                                      bg=COLORS["bg"], fg=COLORS["fg"], bd=0, font=self.fonts["title_frame"])
        self.group_cd.pack(fill="both", expand=True)

        frame_cd_inner = tk.Frame(self.group_cd, bg=COLORS["bg"])
        frame_cd_inner.pack(pady=pad_y_inner)

        tk.Label(frame_cd_inner, text="經過", bg=COLORS["bg"], fg=COLORS["fg"], font=self.fonts["ui"]).pack(side="left")

        entry_width = 4  # Entry 的 width 是字元數，不需要乘 scale
        self.entry_cd_h = tk.Entry(frame_cd_inner, width=entry_width, bg=COLORS["entry_bg"], fg=COLORS["entry_fg"],
                                   font=self.fonts["input"], justify="center", insertbackground="white")
        self.entry_cd_h.insert(0, self.config.get("cd_h", "0"))
        self.entry_cd_h.pack(side="left", padx=5)

        tk.Label(frame_cd_inner, text="小時", bg=COLORS["bg"], fg=COLORS["fg"], font=self.fonts["ui"]).pack(side="left")

        self.entry_cd_m = tk.Entry(frame_cd_inner, width=entry_width, bg=COLORS["entry_bg"], fg=COLORS["entry_fg"],
                                   font=self.fonts["input"], justify="center", insertbackground="white")
        self.entry_cd_m.insert(0, self.config.get("cd_m", "0"))
        self.entry_cd_m.pack(side="left", padx=5)
        tk.Label(frame_cd_inner, text="分後關機", bg=COLORS["bg"], fg=COLORS["fg"], font=self.fonts["ui"]).pack(
            side="left")

        # --- 指定時間 ---
        self.wrap_sp = tk.Frame(root, bg=COLORS["border_inactive"], padx=border_pad, pady=border_pad)
        self.wrap_sp.pack(fill="x", padx=pad_x_outer, pady=int(10 * self.scale))

        self.group_sp = tk.LabelFrame(self.wrap_sp, text=" ⏰ 指定時間 ",
                                      bg=COLORS["bg"], fg=COLORS["fg"], bd=0, font=self.fonts["title_frame"])
        self.group_sp.pack(fill="both", expand=True)

        frame_sp_inner = tk.Frame(self.group_sp, bg=COLORS["bg"])
        frame_sp_inner.pack(pady=pad_y_inner)

        tk.Label(frame_sp_inner, text="設定於", bg=COLORS["bg"], fg=COLORS["fg"], font=self.fonts["ui"]).pack(
            side="left")

        self.entry_sp_h = tk.Entry(frame_sp_inner, width=entry_width, bg=COLORS["entry_bg"], fg=COLORS["entry_fg"],
                                   font=self.fonts["input"], justify="center", insertbackground="white")
        val_h = self.config.get("sp_h", "0")
        self.entry_sp_h.insert(0, val_h if val_h else "0")
        self.entry_sp_h.pack(side="left", padx=5)

        tk.Label(frame_sp_inner, text="點", bg=COLORS["bg"], fg=COLORS["fg"], font=self.fonts["ui"]).pack(side="left")

        self.entry_sp_m = tk.Entry(frame_sp_inner, width=entry_width, bg=COLORS["entry_bg"], fg=COLORS["entry_fg"],
                                   font=self.fonts["input"], justify="center", insertbackground="white")
        val_m = self.config.get("sp_m", "0")
        self.entry_sp_m.insert(0, val_m if val_m else "0")
        self.entry_sp_m.pack(side="left", padx=5)
        tk.Label(frame_sp_inner, text="分關機", bg=COLORS["bg"], fg=COLORS["fg"], font=self.fonts["ui"]).pack(
            side="left")

        # --- 按鈕 ---
        self.btn_toggle = tk.Button(root, text="開始倒數", bg=COLORS["btn_start"], fg="white",
                                    font=self.fonts["btn_big"], activebackground="#1b5e20", activeforeground="white",
                                    command=self.toggle_schedule, relief="flat", cursor="hand2")
        self.btn_toggle.pack(pady=(int(20 * self.scale), int(10 * self.scale)), ipadx=int(30 * self.scale),
                             ipady=int(5 * self.scale))

        # --- 綁定 ---
        self.entry_cd_h.bind("<FocusIn>", lambda e: self.set_mode(1))
        self.entry_cd_m.bind("<FocusIn>", lambda e: self.set_mode(1))
        self.entry_sp_h.bind("<FocusIn>", lambda e: self.set_mode(2))
        self.entry_sp_m.bind("<FocusIn>", lambda e: self.set_mode(2))

        for entry in [self.entry_cd_h, self.entry_cd_m, self.entry_sp_h, self.entry_sp_m]:
            entry.bind("<KeyRelease>", self.update_preview)

        self.on_mode_change()
        self.update_preview()
        self.update_clock()

    def on_window_minimize(self, event):
        if event.widget == self.root:
            if self.root.state() == 'iconic':
                self.root.withdraw()
                if self.tray_icon is None:
                    self.minimize_to_tray()

    def toggle_schedule(self):
        if not self.is_running:
            self.start_process()
        else:
            self.stop_process()

    def set_mode(self, mode):
        if self.mode_var.get() != mode:
            self.mode_var.set(mode)
            self.on_mode_change()
            self.update_preview()

    def on_mode_change(self):
        mode = self.mode_var.get()
        if mode == 1:
            self.wrap_cd.config(bg=COLORS["border_active"])
            self.wrap_sp.config(bg=COLORS["border_inactive"])
            self.group_cd.config(fg=COLORS["border_active"])
            self.group_sp.config(fg="gray")
        else:
            self.wrap_cd.config(bg=COLORS["border_inactive"])
            self.wrap_sp.config(bg=COLORS["border_active"])
            self.group_cd.config(fg="gray")
            self.group_sp.config(fg=COLORS["border_active"])

    def calculate_target_time(self):
        now = datetime.datetime.now().replace(microsecond=0)
        try:
            if self.mode_var.get() == 1:
                h = int(self.entry_cd_h.get() or 0)
                m = int(self.entry_cd_m.get() or 0)
                if h == 0 and m == 0: return now, None
                return now + datetime.timedelta(hours=h, minutes=m), None
            else:
                h = int(self.entry_sp_h.get() or 0)
                m = int(self.entry_sp_m.get() or 0)
                if not (0 <= h <= 23 and 0 <= m <= 59): return None, "時間格式錯誤"
                target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                if target <= now: target += datetime.timedelta(days=1)
                return target, None
        except ValueError:
            return None, "格式錯誤"

    def update_preview(self, event=None):
        if self.is_running: return
        target, _ = self.calculate_target_time()
        self.preview_time = target
        if target:
            self.lbl_status.config(text=f"預計於 {target.strftime('%H:%M')} 關機", fg=COLORS["status_fg"])
        else:
            self.lbl_status.config(text="等待設定...", fg="gray")

    def draw_hand(self, center, length, angle, color, width=2):
        scaled_width = max(1, int(width * self.scale))
        angle_rad = math.radians(angle - 90)
        x = center + length * math.cos(angle_rad)
        y = center + length * math.sin(angle_rad)
        self.canvas.create_line(center, center, x, y, width=scaled_width, fill=color, capstyle=tk.ROUND)

    def update_clock(self):
        self.canvas.delete("all")
        center = 140 * self.scale
        radius = 120 * self.scale

        border_w = int(3 * self.scale)

        self.canvas.create_oval(center - radius, center - radius, center + radius, center + radius, width=border_w,
                                outline=COLORS["clock_border"], fill=COLORS["clock_face"])

        tick_len_long = 20 * self.scale
        tick_len_short = 10 * self.scale
        tick_width = max(1, int(2 * self.scale))

        for i in range(12):
            angle = i * 30
            rad = math.radians(angle - 90)
            x1 = center + (radius - tick_len_short) * math.cos(rad)
            y1 = center + (radius - tick_len_short) * math.sin(rad)
            x2 = center + (radius - tick_len_long) * math.cos(rad)
            y2 = center + (radius - tick_len_long) * math.sin(rad)
            self.canvas.create_line(x1, y1, x2, y2, width=tick_width, fill=COLORS["clock_tick"])

        now = datetime.datetime.now()
        display_target = self.target_time if self.is_running else self.preview_time

        if display_target:
            t_h, t_m = display_target.hour, display_target.minute
            self.draw_hand(center, radius * 0.55, (t_h % 12 + t_m / 60) * 30, COLORS["hand_target"], width=5)
            self.draw_hand(center, radius * 0.75, t_m * 6, COLORS["hand_target"], width=3)

            if self.is_running:
                remaining = self.target_time - now
                if remaining.total_seconds() <= 0:
                    self.execute_shutdown()
                else:
                    rem_sec = int(remaining.total_seconds())
                    rem_h, rem_r = divmod(rem_sec, 3600)
                    rem_m, rem_s = divmod(rem_r, 60)
                    time_str = f"{rem_h}:{rem_m:02}:{rem_s:02}"

                    self.lbl_status.config(
                        text=f"將於 {self.target_time.strftime('%H:%M')} 關機\n剩餘 {time_str}",
                        fg="#ff5252")

        # 指針
        self.draw_hand(center, radius * 0.5, (now.hour % 12) * 30 + now.minute * 0.5, COLORS["hand_hour"], width=6)
        self.draw_hand(center, radius * 0.8, now.minute * 6, COLORS["hand_min"], width=4)
        self.draw_hand(center, radius * 0.9, now.second * 6, COLORS["hand_sec"], width=2)

        # 中心點
        center_dot = 5 * self.scale
        self.canvas.create_oval(center - center_dot, center - center_dot, center + center_dot, center + center_dot,
                                fill=COLORS["hand_sec"])

        self.root.after(100, self.update_clock)

    def start_process(self):
        target, err = self.calculate_target_time()
        if self.mode_var.get() == 1:
            try:
                if int(self.entry_cd_h.get() or 0) == 0 and int(self.entry_cd_m.get() or 0) == 0:
                    messagebox.showwarning("警告", "倒數時間不能為 0")
                    return
            except:
                pass

        if err:
            messagebox.showerror("設定錯誤", err)
            return

        self.target_time = target
        self.config["mode"] = self.mode_var.get()
        self.config["cd_h"] = self.entry_cd_h.get()
        self.config["cd_m"] = self.entry_cd_m.get()
        self.config["sp_h"] = self.entry_sp_h.get()
        self.config["sp_m"] = self.entry_sp_m.get()
        save_config(self.config)

        self.is_running = True
        self.update_ui_state(locked=True)

        if not self.config.get("skip_warning", False):
            self.show_warning_dialog()
        else:
            self.minimize_to_tray()

    def stop_process(self):
        self.is_running = False
        self.target_time = None
        self.preview_time = None
        self.lbl_status.config(text="排程已取消", fg="#cfcfcf")
        self.update_ui_state(locked=False)
        self.update_preview()

    def update_ui_state(self, locked):
        state = "disabled" if locked else "normal"
        bg_color = "#2b2b2b" if locked else COLORS["entry_bg"]

        if locked:
            self.btn_toggle.config(text="取消設定", bg=COLORS["btn_cancel"], activebackground="#b71c1c")
        else:
            self.btn_toggle.config(text="開始倒數", bg=COLORS["btn_start"], activebackground="#1b5e20")

        for entry in [self.entry_cd_h, self.entry_cd_m, self.entry_sp_h, self.entry_sp_m]:
            entry.config(state=state, disabledbackground=bg_color)

    def show_warning_dialog(self):
        top = tk.Toplevel(self.root)
        top.title("提示")

        w, h = int(340 * self.scale), int(240 * self.scale)
        top.geometry(f"{w}x{h}")
        top.configure(bg="#2b2b2b")
        top.resizable(False, False)

        x = self.root.winfo_x() + 50
        y = self.root.winfo_y() + 100
        top.geometry(f"+{x}+{y}")

        tk.Label(top, text="⚠️ 縮小提示", font=self.fonts["warning_title"], fg="#ffb74d", bg="#2b2b2b").pack(
            pady=int(15 * self.scale))
        tk.Label(top, text="程式將縮小至右下角\n在背景執行自動關機", font=self.fonts["warning_text"], fg="#e0e0e0",
                 bg="#2b2b2b").pack(pady=int(5 * self.scale))

        chk_var = tk.BooleanVar()
        chk = tk.Checkbutton(top, text="不再提示", variable=chk_var,
                             bg="#2b2b2b", fg="#4fc3f7", font=self.fonts["ui_bold"],
                             selectcolor="#3c3f41", activebackground="#2b2b2b", activeforeground="#4fc3f7")
        chk.pack(pady=int(15 * self.scale))

        def on_confirm():
            if chk_var.get():
                self.config["skip_warning"] = True
                save_config(self.config)
            top.destroy()
            self.minimize_to_tray()

        tk.Button(top, text="好，我知道了", command=on_confirm, width=15, bg="#1976d2", fg="white",
                  font=self.fonts["title_frame"], relief="flat").pack(pady=5)

    def load_tray_icon(self):
        try:
            icon_path = resource_path(ICON_FILENAME)
            if os.path.exists(icon_path): return Image.open(icon_path)
        except:
            pass
        img = Image.new('RGB', (64, 64), (32, 32, 32))
        ImageDraw.Draw(img).ellipse((8, 8, 56, 56), fill='#ff1744')
        return img

    def minimize_to_tray(self):
        try:
            alpha = 1.0
            while alpha > 0:
                alpha -= 0.1
                self.root.attributes('-alpha', alpha)
                self.root.update()
                time.sleep(0.015)
        except:
            pass

        self.root.withdraw()
        self.root.attributes('-alpha', 1.0)

        if self.target_time:
            title_text = f"自動關機 ({self.target_time.strftime('%H:%M')})"
        else:
            title_text = "自動關機"

        menu = pystray.Menu(
            pystray.MenuItem("顯示主視窗", self.restore_from_tray, default=True),
            pystray.MenuItem("取消關機並顯示視窗", self.cancel_and_restore)
        )
        self.tray_icon = pystray.Icon("MapleTimer", self.load_tray_icon(),
                                      title_text, menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def restore_from_tray(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
        self.root.attributes('-alpha', 0.0)
        self.root.deiconify()
        self.root.state('normal')

        try:
            alpha = 0.0
            while alpha < 1.0:
                alpha += 0.1
                self.root.attributes('-alpha', alpha)
                self.root.update()
                time.sleep(0.015)
            self.root.attributes('-alpha', 1.0)
        except:
            self.root.attributes('-alpha', 1.0)

    def cancel_and_restore(self, icon=None, item=None):
        self.stop_process()
        self.restore_from_tray()

    def execute_shutdown(self):
        self.is_running = False
        self.lbl_status.config(text="執行關機中...", fg="#ff5252")
        os.system("shutdown /s /t 0")

    def on_close(self):
        if self.is_running:
            self.minimize_to_tray()
            messagebox.showinfo("提示", "倒數計時中，程式已縮小至系統列")
        else:
            if self.tray_icon: self.tray_icon.stop()
            self.root.destroy()
            sys.exit()


if __name__ == "__main__":
    enable_high_dpi()
    single_socket = check_single_instance()
    root = tk.Tk()
    app = ShutdownApp(root, single_socket)
    root.mainloop()
