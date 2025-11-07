import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import time
import threading
# Cần cài đặt: pip install pandas
try:
    import pandas as pd
except ImportError:
    messagebox.showerror("Lỗi Cài Đặt", "Vui lòng chạy 'pip install pandas' để cài đặt thư viện xử lý CSV.")
    exit()

import subprocess
import ctypes
import json
from decimal import Decimal
from enum import Enum, auto

# Cần cài đặt: pip install ttkthemes
try:
    from ttkthemes import ThemedTk
except ImportError:
    messagebox.showerror("Lỗi Cài Đặt", "Vui lòng chạy 'pip install ttkthemes' để cài đặt thư viện giao diện.")
    exit()

# Cần cài đặt: pip install Pillow
try:
    from PIL import Image, ImageTk, ImageDraw
    import io
    import base64
except ImportError:
    messagebox.showerror("Lỗi Cài Đặt", "Vui lòng chạy 'pip install Pillow' để cài đặt thư viện xử lý ảnh.")
    exit()

# Cần cài đặt: pip install pynput
try:
    from pynput import mouse, keyboard
    from pynput.keyboard import Key
except ImportError:
    messagebox.showerror("Lỗi Cài Đặt", "Vui lòng chạy 'pip install pynput' để cài đặt thư viện ghi chuột/phím.")
    exit()

# Cần cài đặt: pip install pywin32
try:
    import win32gui
    import win32con
    import win32api
    import win32process
    from win32process import GetWindowThreadProcessId
except ImportError:
    win32gui = win32con = win32api = win32process = None
    messagebox.showerror("Lỗi Cài Đặt", "Vui lòng chạy 'pip install pywin32' để sử dụng chức năng cửa sổ Windows.")
    exit()

# =========================================================================
# ----------------------------- WinAPI Helpers ----------------------------
# =========================================================================
user32 = ctypes.windll.user32


def get_dpi_scale_factor(hwnd):
    """
    Lấy tỷ lệ DPI (ví dụ: 1.5 cho 150% scale) cho cửa sổ hiện tại,
    sử dụng WinAPI để đảm bảo chính xác.
    """
    if win32api is None:
        return 1.0
    try:
        # Lấy handle của monitor chứa cửa sổ
        monitor_handle = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)

        # Lấy DPI (ví dụ: 144 cho 150%)
        dpi_x = win32api.GetDpiForMonitor(monitor_handle, 0)

        # Tỷ lệ scale: DPI / 96 (DPI chuẩn 100%)
        scale_factor = Decimal(dpi_x) / Decimal(96)
        return float(scale_factor)
    except Exception:
        return 1.0


def hwnd_from_title(title_substring):
    """Tìm handle cửa sổ có tiêu đề chứa chuỗi con."""
    result = []

    def enumf(hwnd, arg):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd) or ""
            if title_substring.lower() in title.lower():
                result.append(hwnd)
        return True

    win32gui.EnumWindows(enumf, None)
    return result[0] if result else None


def get_window_rect(hwnd):
    """Lấy (left, top, right, bottom) tọa độ tuyệt đối của cửa sổ (LOGICAL PIXEL)."""
    try:
        rect = win32gui.GetWindowRect(hwnd)
        return rect
    except Exception:
        return None


def bring_to_front(hwnd):
    """
    Đưa cửa sổ lên foreground.
    Chỉ gọi SW_RESTORE nếu cửa sổ đang bị Minimize, giữ nguyên trạng thái Maximize.
    """
    try:
        placement = win32gui.GetWindowPlacement(hwnd)

        if placement[1] == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        elif placement[1] == win32con.SW_SHOWMAXIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

        user32.SetForegroundWindow(hwnd)
        time.sleep(0.1)
        return True
    except Exception:
        return False


# =========================================================================
# ----------------------------- Mouse/Keyboard Sender ---------------------
# =========================================================================

mouse_controller = mouse.Controller()
keyboard_controller = keyboard.Controller()


def send_char_to_hwnd(hwnd, char):
    """Gửi ký tự thông thường (cần focus)."""
    if bring_to_front(hwnd):
        keyboard_controller.type(char)
        time.sleep(0.01)
        return True
    return False


def send_key_to_hwnd(hwnd, key_name):
    """Gửi phím đặc biệt (cần focus)."""
    if bring_to_front(hwnd):
        try:
            key_name = key_name.lower().replace("key.", "")
            key = getattr(Key, key_name, None)

            if key:
                keyboard_controller.press(key)
                keyboard_controller.release(key)
                time.sleep(0.01)
                return True
            else:
                keyboard_controller.press(key_name)
                keyboard_controller.release(key_name)
                time.sleep(0.01)
                return True
        except Exception:
            return False
    return False


def send_combo_to_hwnd(hwnd, combo_string):
    """Gửi tổ hợp phím (ví dụ: 'ctrl+a', 'alt+f4') (cần focus)."""
    if bring_to_front(hwnd):
        try:
            parts = combo_string.lower().split('+')
            keys = []

            for part in parts:
                key = getattr(Key, part, None)
                if key:
                    keys.append(key)
                else:
                    keys.append(part)

            for key in keys[:-1]:
                keyboard_controller.press(key)

            keyboard_controller.press(keys[-1])
            keyboard_controller.release(keys[-1])

            for key in reversed(keys[:-1]):
                keyboard_controller.release(key)

            time.sleep(0.01)
            return True
        except Exception:
            return False
    return False


def send_mouse_click(hwnd, x_offset_logical, y_offset_logical, button_type, recorded_scale):
    """
    Thực hiện click chuột tại vị trí OFFSET PIXEL cố định so với góc trên bên trái
    của cửa sổ mục tiêu (hwnd), có xử lý DPI Scaling.
    """
    if not bring_to_front(hwnd):
        return False

    rect = get_window_rect(hwnd)
    if not rect:
        return False
    left, top, right, bottom = rect

    # --- XỬ LÝ DPI SCALING ---
    current_scale = get_dpi_scale_factor(hwnd)

    # Tính Normalized Offset (Offset chuẩn 100% scale)
    recorded_scale = recorded_scale if recorded_scale > 0 else 1.0
    x_offset_normalized = x_offset_logical / recorded_scale
    y_offset_normalized = y_offset_logical / recorded_scale

    # Tính Offset Logical hiện tại (Offset chuẩn * Scale hiện tại)
    # Đây là vị trí offset LOGICAL cần tìm trong cửa sổ hiện tại
    x_offset_current_logical = x_offset_normalized * current_scale
    y_offset_current_logical = y_offset_normalized * current_scale

    # VỊ TRÍ TUYỆT ĐỐI CUỐI CÙNG (Logical Pixel)
    x_abs_logical = left + x_offset_current_logical
    y_abs_logical = top + y_offset_current_logical

    # Di chuyển chuột đến vị trí tuyệt đối (Logical Pixel)
    mouse_controller.position = (int(x_abs_logical), int(y_abs_logical))
    time.sleep(0.01)

    button = mouse.Button.left
    click_count = 1

    if "right" in button_type:
        button = mouse.Button.right

    if "double" in button_type:
        click_count = 2

    # Thực hiện click
    mouse_controller.click(button, click_count)

    time.sleep(0.01)
    return True


# =========================================================================
# ------------------------------ Macro Model ------------------------------
# =========================================================================

class MacroStepType(Enum):
    """Defines the types of steps in a macro for better type safety."""
    COLUMN_DATA = "col"
    MOUSE_CLICK = "mouse"
    KEY_PRESS = "key"
    KEY_COMBO = "combo"
    END_OF_ROW = "end"


class MacroStep:
    def __init__(
        self,
        typ,
        key_value=None,
        col_index=None,
        delay_after=0.01,
        x_offset=None,
        y_offset=None,
        dpi_scale=1.0,
    ):
        self.typ = typ
        self.key_value = key_value
        self.col_index = col_index
        self.delay_after = delay_after
        # Tọa độ Offset ghi nhận lúc ghi (Logical Pixel)
        self.x_offset_logical = x_offset
        self.y_offset_logical = y_offset
        self.dpi_scale = dpi_scale
        self.item_id = None
        self.item_idx = -1

    def __repr__(self):
        delay_ms = int(self.delay_after * 1000)
        delay_str = f"(Chờ: {delay_ms}ms)"

        if self.typ == MacroStepType.COLUMN_DATA.value:
            col_display = self.col_index + 1 if self.col_index is not None else "N/A"
            return f"[COL] Gửi giá trị cột {col_display} {delay_str}"
        elif self.typ == MacroStepType.MOUSE_CLICK.value:
            # HIỂN THỊ DƯỚNG DẠNG PIXEL OFFSET CHUẨN (100% SCALE)

            scale = self.dpi_scale if self.dpi_scale > 0 else 1.0

            # Tính toán Offset Chuẩn hóa (Normalized Offset)
            x_norm = int(self.x_offset_logical / scale) if self.x_offset_logical is not None else "N/A"
            y_norm = int(self.y_offset_logical / scale) if self.y_offset_logical is not None else "N/A"

            click_type = self.key_value.replace("_click", "").capitalize()
            return f"[MOUSE] {click_type} Click tại Offset Chuẩn ({x_norm}px, {y_norm}px) (Scale Ghi: {int(self.dpi_scale * 100)}%) {delay_str}"
        elif self.typ == MacroStepType.KEY_PRESS.value:
            return f"[KEY] Gửi phím: '{self.key_value.upper()}' {delay_str}"
        elif self.typ == MacroStepType.KEY_COMBO.value:
            return f"[COMBO] Gửi tổ hợp phím: '{self.key_value.upper()}' {delay_str}"
        elif self.typ == MacroStepType.END_OF_ROW.value:
            return f"[END] Kết thúc dòng {delay_str}"
        return f"<{self.typ}>"

    def to_dict(self):
        return {
            'typ': self.typ,
            'key_value': self.key_value,
            'col_index': self.col_index,
            'delay_after': self.delay_after,
            'x_offset_logical': self.x_offset_logical,
            'y_offset_logical': self.y_offset_logical,
            'dpi_scale': self.dpi_scale
        }

    @staticmethod
    def from_dict(data):
        x_val = data.get("x_offset_logical", data.get("x_offset"))
        y_val = data.get("y_offset_logical", data.get("y_offset"))

        return MacroStep(
            typ=data["typ"],
            key_value=data.get('key_value'),
            col_index=data.get('col_index'),
            delay_after=data.get('delay_after', 0.01),
            x_offset=x_val,
            y_offset=y_val,
            dpi_scale=data.get('dpi_scale', 1.0)
        )


# =========================================================================
# ------------------------------ HUD Window -------------------------------
# =========================================================================

class RecordingHUD(tk.Toplevel):
    PAUSED_COLOR = "#FFD700" # Vàng
    """
    Một cửa sổ HUD nhỏ, luôn ở trên cùng, để hiển thị trạng thái ghi/phát
    và cung cấp nút Stop.
    """
    def __init__(self, parent, stop_callback):
        super().__init__(parent)
        self.stop_callback = stop_callback

        self.pause_event = threading.Event()
        self.is_paused = False

        # Biến để di chuyển cửa sổ
        self._offset_x = 0
        self._offset_y = 0

        # Thiết lập cửa sổ HUD
        self.overrideredirect(True)  # Bỏ viền và thanh tiêu đề
        self.attributes('-topmost', True)  # Luôn ở trên cùng
        self.attributes('-alpha', 0.9)  # Hơi trong suốt
        self.withdraw() # Ẩn cửa sổ ban đầu để tránh nhấp nháy
        # Tạo style cho các widget trong HUD
        style = ttk.Style(self)
        style.configure("HUD.TFrame", background="#282c34")
        style.configure("HUD.TLabel", background="#282c34", foreground="white", font=("Arial", 10))
        style.configure("HUD.TButton", font=("Arial", 9, "bold"), foreground="white")
        style.map("HUD.TButton",
                  background=[('active', '#e06c75'), ('!active', '#d15660')],
                  foreground=[('active', 'white')])
        style.configure("Pause.TButton", font=("Arial", 9, "bold"))
        style.map("Pause.TButton",
                  background=[('active', '#61afef'), ('!active', '#5699d6')],
                  foreground=[('active', 'white')])

        # Frame chính của HUD
        main_frame = ttk.Frame(self, style="HUD.TFrame", padding=(10, 5))
        main_frame.pack()

        # Label hiển thị trạng thái
        self.status_label = ttk.Label(main_frame, text="Chuẩn bị...", style="HUD.TLabel")
        self.status_label.pack(side="left", padx=(0, 10))

        # Nút Pause/Resume
        self.pause_button = ttk.Button(main_frame, text="❚❚ PAUSE", style="Pause.TButton", command=self.toggle_pause)
        self.pause_button.pack(side="left", padx=(0, 10))
        self.pause_button.pack_forget() # Ẩn nút pause ban đầu

        # Nút Stop
        stop_button = ttk.Button(main_frame, text="■ STOP", style="HUD.TButton", command=self.stop_callback)
        stop_button.pack(side="left")

        # Gán sự kiện để di chuyển HUD
        main_frame.bind("<Button-1>", self._on_mouse_press)
        main_frame.bind("<B1-Motion>", self._on_mouse_drag)
        self.status_label.bind("<Button-1>", self._on_mouse_press)
        self.status_label.bind("<B1-Motion>", self._on_mouse_drag)

        # Căn giữa HUD ở cạnh trên màn hình
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        window_width = self.winfo_width()
        x = (screen_width // 2) - (window_width // 2)
        self.geometry(f"+{x}+20") # 20px từ cạnh trên
        self.deiconify() # Hiện cửa sổ ở đúng vị trí

    def _on_mouse_press(self, event):
        self._offset_x = event.x
        self._offset_y = event.y

    def _on_mouse_drag(self, event):
        x = self.winfo_pointerx() - self._offset_x
        y = self.winfo_pointery() - self._offset_y
        self.geometry(f"+{x}+{y}")

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_event.clear() # Chặn thread
            self.pause_button.config(text="▶ RESUME")
            self.update_status("⏸ TẠM DỪNG", self.PAUSED_COLOR)
        else:
            self.pause_event.set() # Cho phép thread chạy tiếp
            self.pause_button.config(text="❚❚ PAUSE")
            # Trạng thái sẽ được cập nhật lại bởi vòng lặp chính

    def update_status(self, text, color="white"):
        """Cập nhật văn bản và màu sắc của label trạng thái."""
        if self.is_paused: # Nếu đang pause thì không cập nhật status từ bên ngoài
            return
        self.status_label.config(text=text, foreground=color)

    def close(self):
        self.destroy()


# =========================================================================
# ------------------------------ Tkinter App (Themed) ---------------------
# =========================================================================

LOGO_PNG_BASE64 = b"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"  # Placeholder


class MacroApp(ThemedTk):
    def __init__(self):
        super().__init__(theme="arc")

        self.option_add("*font", "Arial 9") # SỬA: Tải logo từ chuỗi Base64
        original_image = None
        try:
            logo_data = base64.b64decode(LOGO_PNG_BASE64)
            original_image = Image.open(io.BytesIO(logo_data))
        except Exception as e:
            messagebox.showerror("Lỗi Logo", f"Không thể tải logo từ Base64: {e}")
            # Tạo ảnh placeholder nếu không tìm thấy logo
            original_image = Image.new('RGB', (60, 20), color='#1e90ff')
            d = ImageDraw.Draw(original_image)
            d.text((5, 5), "VT", fill=(255, 255, 255))

        icon_image = original_image.copy()
        icon_image.thumbnail((64, 64), Image.Resampling.LANCZOS)
        self.app_icon = ImageTk.PhotoImage(icon_image)
        self.iconphoto(True, self.app_icon)

        header_image = original_image.copy()
        # SỬA: Tăng kích thước logo để cân đối với tiêu đề
        header_image.thumbnail((36, 36), Image.Resampling.LANCZOS)
        self.header_logo = ImageTk.PhotoImage(header_image)

        self.title("Việt Tín Auto Sender V2025.04 (Fix & Layout Update)")
        self.geometry("1000x800")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.macro_steps = []
        self.recording = False
        self.current_col_index = 0
        self.last_key_time = 0.0
        self.csv_path = None
        self.acsoft_path = None
        self.cancel_flag = threading.Event()
        self.target_window_title = ""
        self.df_csv = pd.DataFrame()

        self.speed_mode = tk.IntVar(value=1)
        self.spin_speed_val = tk.IntVar(value=500)
        self.spin_between_val = tk.IntVar(value=2)
        self.show_realtime_status = tk.BooleanVar(value=False) # SỬA: Tắt theo mặc định để giảm lag

        self.txt_delimiter = None
        self.txt_acpath = None
        self.txt_csv = None
        self.lbl_realtime_status = None

        self.mouse_listener = None
        self.keyboard_listener = None
        self.current_modifiers = set()

        self.pause_event = None
        self.hud_window = None
        self.realtime_status_frame = None

        self.setup_ui()

        self.style = ttk.Style()
        self.style.map("Macro.Treeview", background=[("selected", "#1e90ff"), ("active", "#e1e1e1")])
        self.style.configure("Macro.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", font=("Arial", 9))
        self.style.map("Highlight.Treeview", background=[("selected", "#FFA07A"), ("active", "#e1e1e1")])
        self.style.configure('Accent.TButton', font=('Arial', 9, 'bold'))

        self.protocol("WM_DELETE_WINDOW", self.on_app_close)
        self.resizable(False, False) # SỬA: Vô hiệu hóa thay đổi kích thước cửa sổ

        # BẮT ĐẦU VÒNG LẶP CẬP NHẬT TRẠNG THÁI REAL-TIME
        self._update_status_bar_info()

    def on_app_close(self):
        self.stop_listeners()
        self.destroy()

    def setup_ui(self):
        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=10)
        main_frame.grid_rowconfigure(2, weight=1)  # Dòng chứa CSV/Macro (g_data_macro)
        main_frame.grid_columnconfigure(0, weight=1)

        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        if self.header_logo:
            tk.Label(header_frame, image=self.header_logo).pack(side="left", padx=(0, 5))
        tk.Label(header_frame, text="Việt Tín Auto Sender V2025.04", font=("Arial", 18, "bold"), fg="#1e90ff").pack(
            side="left", padx=(0, 10)
        )
        tk.Label(header_frame, text="Ghi macro (Phím, Chuột & Dữ liệu) & replay vào ACSOFT", font=("Arial", 10)).pack(
            side="left", padx=10
        )

        top_controls_frame = ttk.Frame(main_frame)
        top_controls_frame.grid(row=1, column=0, sticky="ew", pady=5)
        top_controls_frame.grid_columnconfigure(0, weight=1)
        top_controls_frame.grid_columnconfigure(1, weight=2)
        top_controls_frame.grid_columnconfigure(2, weight=1)

        g1 = ttk.LabelFrame(top_controls_frame, text="1) Đường dẫn file chạy ACSOFT")
        g1.grid(row=0, column=0, sticky="nwe", padx=(0, 10), pady=5)
        g1.grid_columnconfigure(0, weight=1)
        g1l = ttk.Frame(g1)
        g1l.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.txt_acpath = ttk.Entry(g1l, width=10)
        self.txt_acpath.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(g1l, text="Browse", command=self.browse_ac, width=8).pack(side="left", padx=(0, 5))
        ttk.Button(g1l, text="Mở ACSOFT", command=self.open_ac, width=10).pack(side="left")

        g2 = ttk.LabelFrame(top_controls_frame, text="2) File CSV chứa dữ liệu / Cửa sổ mục tiêu")
        g2.grid(row=0, column=1, sticky="nwe", padx=(0, 10), pady=5)
        g2.grid_columnconfigure(0, weight=1)

        g2l_csv = ttk.Frame(g2)
        g2l_csv.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 0))
        tk.Label(g2l_csv, text="File CSV:", width=8, anchor="w").pack(side="left")
        self.txt_csv = ttk.Entry(g2l_csv, width=10)
        self.txt_csv.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(g2l_csv, text="Browse CSV", command=self.browse_csv).pack(side="left")

        g2l_window = ttk.Frame(g2)
        g2l_window.grid(row=1, column=0, sticky="ew", padx=5, pady=(5, 5))

        # --- SỬA LỖI LAYOUT: Dọn dẹp và chỉ sử dụng Grid cho g2l_window ---
        # Cấu hình để chỉ cột 3 (chứa Combobox) được co giãn
        g2l_window.grid_columnconfigure(3, weight=1)

        # Cột 0: Label "Delimiter"
        tk.Label(g2l_window, text="Delimiter:", width=8, anchor="w").grid(row=0, column=0, sticky="w")
        # Cột 1: Ô nhập Delimiter
        self.txt_delimiter = ttk.Entry(g2l_window, width=3)
        self.txt_delimiter.insert(0, ";")
        self.txt_delimiter.grid(row=0, column=1, sticky="w", padx=(0, 10))

        # Cột 2: Label "Cửa sổ"
        tk.Label(g2l_window, text="Cửa sổ:", width=6, anchor="w").grid(row=0, column=2, sticky="w")
        # Cột 3: Combobox (co giãn)
        self.combo_windows = ttk.Combobox(g2l_window, state="readonly", width=30)
        self.combo_windows.grid(row=0, column=3, sticky="ew", padx=(0, 5))
        self.combo_windows.bind("<<ComboboxSelected>>", self.on_window_select)
        # Cột 4: Nút "Làm mới"
        ttk.Button(g2l_window, text="Làm mới", command=self.refresh_windows, width=8).grid(row=0, column=4, sticky="e")
        # --- KẾT THÚC SỬA ---

        self.refresh_windows()

        g4 = ttk.LabelFrame(top_controls_frame, text="3) Tùy chọn chạy")
        g4.grid(row=0, column=2, sticky="nwe", pady=5)
        g4.grid_columnconfigure(0, weight=1)

        g4l_speed = ttk.Frame(g4)
        g4l_speed.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.speed_mode = tk.IntVar(value=1)
        ttk.Radiobutton(g4l_speed, text="Tốc độ đã ghi", variable=self.speed_mode, value=1).pack(side="left", padx=5)
        self.rb_custom = ttk.Radiobutton(g4l_speed, text="Tốc độ cố định:", variable=self.speed_mode, value=2)
        self.rb_custom.pack(side="left", padx=10)

        self.spin_speed = ttk.Spinbox(g4l_speed, from_=100, to=10000, increment=100, textvariable=self.spin_speed_val,
                                      width=5)
        self.spin_speed.pack(side="left", padx=5)
        tk.Label(g4l_speed, text="ms").pack(side="left")
        self.speed_mode.trace_add('write', lambda *args: self.spin_speed.config(
            state='normal' if self.speed_mode.get() == 2 else 'disabled'))
        self.spin_speed.config(state='disabled')

        g4l_delay = ttk.Frame(g4)
        g4l_delay.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
        tk.Label(g4l_delay, text="Đợi giữa 2 dòng (1-20 giây):").pack(side="left")
        self.spin_between = ttk.Spinbox(g4l_delay, from_=1, to=20, textvariable=self.spin_between_val, width=5)
        self.spin_between.pack(side="left", padx=5)


        # ====================================================================
        # KHUNG CHỨA DỮ LIỆU VÀ MACRO (SIDE-BY-SIDE)
        # ====================================================================
        g_data_macro = ttk.Frame(main_frame)
        g_data_macro.grid(row=2, column=0, sticky="nsew", pady=5)
        g_data_macro.grid_rowconfigure(0, weight=1)

        # Cấu hình 2 cột: Cột 0 (CSV) và Cột 1 (Macro) đều co giãn bằng nhau
        g_data_macro.grid_columnconfigure(0, weight=1, uniform="group1")
        g_data_macro.grid_columnconfigure(1, weight=1, uniform="group1")

        # --- CỘT TRÁI: DỮ LIỆU CSV (G2) - Ô MÀU VÀNG ---
        g2_table_frame = ttk.LabelFrame(g_data_macro, text="Dữ liệu CSV (Chỉ hiển thị 10 dòng đầu)")
        g2_table_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        g2_table_frame.grid_rowconfigure(0, weight=1)
        g2_table_frame.grid_columnconfigure(0, weight=1)

        # SỬA: CHỈ NHẬN TREEVIEW VÀ BỎ GRID CONTAINER FRAME KHÔNG CẦN THIẾT
        self.tree_csv = self._create_treeview(g2_table_frame, max_cols=20)
        # csv_container_frame.grid(row=0, column=0, sticky='nsew', padx=5, pady=5) # Dòng này đã bị loại bỏ

        # --- CỘT PHẢI: GHI MACRO (G3) - Ô MÀU XANH/ĐỎ ---
        g3 = ttk.LabelFrame(g_data_macro, text="4) Ghi Macro & Điều chỉnh")
        g3.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        # Cấu hình Hàng 1 (Treeview) là hàng DUY NHẤT để co giãn lấp đầy không gian còn lại
        g3.grid_rowconfigure(1, weight=1)
        g3.grid_columnconfigure(0, weight=1)

        # G3: Hàng 0 - Nút Ghi và Tải/Lưu
        g3_controls_record = ttk.Frame(g3)
        g3_controls_record.grid(row=0, column=0, sticky='ew', padx=5, pady=(5, 5))

        self.btn_record = ttk.Button(g3_controls_record, text="Record Macro (5s chuẩn bị)", command=self.record_macro,
                                     style='Accent.TButton')
        self.btn_record.pack(side="left", padx=(0, 10))

        ttk.Button(g3_controls_record, text="Lưu Macro", command=self.save_macro).pack(side="left", padx=5)
        ttk.Button(g3_controls_record, text="Mở Macro", command=self.load_macro).pack(side="left", padx=5)
        ttk.Button(g3_controls_record, text="Clear Macro", command=self.clear_macro).pack(side="right", padx=5)

        # G3: Hàng 1 - Treeview Macro (Bảng Macro co giãn nằm ở đây)
        # --- SỬA LỖI LAYOUT: Xóa code thừa, chỉ tạo 1 Treeview trong 1 Frame chứa ---
        macro_container_frame = ttk.Frame(g3) # Tạo một frame chứa Treeview và scrollbar
        macro_container_frame.grid(row=1, column=0, sticky='nsew', padx=5, pady=(5, 5))
        self.tree_macro = self._create_treeview_macro(macro_container_frame)

        # G3: Hàng 2 - CÁC NÚT THÊM BƯỚC THỦ CÔNG ([+])
        g3_controls_add = ttk.Frame(g3)
        g3_controls_add.grid(row=2, column=0, sticky="ew", padx=5, pady=(5, 0))

        ttk.Button(g3_controls_add, text="[+] Cột Dữ Liệu", command=lambda: self.add_manual_step("col")).pack(
            side="left", padx=(0, 5)
        )
        ttk.Button(g3_controls_add, text="[+] Phím (Key)", command=lambda: self.add_manual_step("key")).pack(
            side="left", padx=5
        )
        ttk.Button(g3_controls_add, text="[+] Tổ Hợp (Combo)", command=lambda: self.add_manual_step("combo")).pack(
            side="left", padx=5
        )
        ttk.Button(g3_controls_add, text="[+] Click Chuột", command=lambda: self.add_manual_step("mouse")).pack(
            side="left", padx=5
        )

        # G3: Hàng 3 - CÁC NÚT SỬA/XÓA/END
        g3_controls_edit = ttk.Frame(g3)
        g3_controls_edit.grid(row=3, column=0, sticky="w", padx=5, pady=(0, 0))

        ttk.Button(g3_controls_edit, text="Sửa Dòng", command=self.edit_macro_step).pack(side="left", padx=(0, 5))
        ttk.Button(g3_controls_edit, text="Xóa Dòng", command=self.delete_macro_step).pack(side="left", padx=5)
        ttk.Button(g3_controls_edit, text="[+] Kết Thúc Dòng", command=lambda: self.add_manual_step("end")).pack(
            side="left", padx=5
        )

        # G3: Hàng 4 - Ghi chú cho Record
        tk.Label(g3, text="Ghi: Insert->cột | Phím/Chuột->thao tác | ESC->kết thúc",
                 font=("Arial", 9, "italic"), fg="gray").grid(row=4, column=0, sticky="w", padx=5, pady=(5, 5))
        # ====================================================================

        # RUN BUTTONS
        g5 = ttk.Frame(main_frame)
        g5.grid(row=3, column=0, sticky="ew", pady=(5, 0))

        tk.Label(g5, text="Chọn Chế độ Chạy:", font=("Arial", 9, "bold")).pack(side="left", padx=(0, 10))

        self.btn_test = ttk.Button(g5, text="CHẠY THỬ (1 DÒNG)", command=self.on_test, style='Accent.TButton')
        self.btn_test.pack(side="left", padx=10)

        self.btn_runall = ttk.Button(g5, text="CHẠY TẤT CẢ", command=self.on_run_all, style='Accent.TButton')
        self.btn_runall.pack(side="left", padx=10)

        self.btn_stop = ttk.Button(g5, text="STOP (ESC)", command=self.on_cancel, state='disabled')
        self.btn_stop.pack(side="left", padx=10)

        self.lbl_status = tk.Label(g5, text="Chờ...", fg="#1e90ff", font=("Arial", 10, "bold"))
        self.lbl_status.pack(side="left", padx=(20, 0))

        # ------------------------ REAL-TIME STATUS FRAME ------------------------
        self.realtime_status_frame = ttk.LabelFrame(main_frame, text="Thông tin Tọa độ (Real-time)")
        self.realtime_status_frame.grid(row=4, column=0, sticky="ew", pady=(5, 0))
        self.realtime_status_frame.grid_columnconfigure(0, weight=1)

        status_controls = ttk.Frame(self.realtime_status_frame)
        status_controls.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        # --- SỬA LỖI LAYOUT: Dọn dẹp và cấu hình lại phần status real-time ---
        # Cấu hình để label co giãn, nút check giữ nguyên
        status_controls.grid_columnconfigure(0, weight=1)

        # Label hiển thị thông tin, đặt ở cột 0
        self.lbl_realtime_status = tk.Label(status_controls, text="...", justify="left", anchor="w", fg="gray")
        self.lbl_realtime_status.grid(row=0, column=0, sticky="ew")
        # Gán sự kiện để cập nhật wraplength khi kích thước label thay đổi
        self.lbl_realtime_status.bind('<Configure>',
                                      lambda e: self.lbl_realtime_status.config(wraplength=self.lbl_realtime_status.winfo_width() - 10))

        # Nút Checkbox, đặt ở cột 1
        ttk.Checkbutton(status_controls, text="Hiện/Ẩn", variable=self.show_realtime_status,
                        command=self._toggle_realtime_status).grid(row=0, column=1, sticky="e", padx=(10, 0))

        # Disclaimer (now row 5)
        tk.Label(main_frame,
                 text="Lưu ý: Ứng dụng BẮT BUỘC đưa ACSOFT lên foreground (phải focus). Nhấn phím ESC để hủy quá trình chạy.",
                 wraplength=900, justify="left", fg="gray", font=("Arial", 8)).grid(row=5, column=0, sticky="w",
                                                                                    pady=(5, 0))

        # SỬA: Áp dụng trạng thái ẩn/hiện ban đầu
        # --- KẾT THÚC SỬA ---
        self._toggle_realtime_status()

        # -------------------------- Real-time Status Update --------------------------

    def _toggle_realtime_status(self):
        """Toggle the visibility of the real-time status frame."""
        # Nếu checkbox được bật, bắt đầu vòng lặp cập nhật.
        # Nếu checkbox bị tắt, vòng lặp sẽ tự dừng ở lần chạy tiếp theo.
        if self.show_realtime_status.get():
            self.lbl_realtime_status.config(text="Đang tải thông tin...")
            self._update_status_bar_info()
        else:
            # Xóa văn bản khi bị ẩn đi
            self.lbl_realtime_status.config(text="...")

    def _update_status_bar_info(self):
        # Chỉ thực hiện công việc và lên lịch lại nếu checkbox được bật
        if self.show_realtime_status.get():
            try:
                hwnd = hwnd_from_title(self.target_window_title)
                # 1. Screen Dimensions (Logical Pixels)
                screen_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                screen_h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

                # 2. Cursor Position (Absolute, Logical Pixels)
                cursor_x, cursor_y = win32api.GetCursorPos()

                status_parts = [f"Màn hình: **{screen_w}x{screen_h}px**"]

                if hwnd and self.target_window_title:
                    rect = get_window_rect(hwnd)
                    if rect:
                        left, top, right, bottom = rect
                        width = right - left
                        height = bottom - top

                        x_offset_logical = cursor_x - left
                        y_offset_logical = cursor_y - top

                        dpi_scale = get_dpi_scale_factor(hwnd)

                        scale = dpi_scale if dpi_scale > 0 else 1.0 # Fix: Ensure scale is not zero
                        if scale == 0: scale = 1.0 # Defensive check
                        x_norm = int(x_offset_logical / scale)
                        y_norm = int(y_offset_logical / scale)

                        status_parts.extend([
                            f"**DPI Scale:** {int(dpi_scale * 100)}%",
                            f"**Cửa sổ ({width}x{height}px):** L{left} T{top} (Logical)",
                            f"**Chuột Tuyệt đối:** X{cursor_x} Y{cursor_y}",
                            f"**Chuột Offset Chuẩn (100%):** Xo{x_norm} Yo{y_norm}",
                        ])
                    else:
                        status_parts.append(f"Cửa sổ '{self.target_window_title}' không tìm thấy hoặc bị ẩn.")
                else:
                    status_parts.append(f"Vui lòng chọn **Cửa sổ mục tiêu**.")

                status_text = " | ".join(status_parts)
                self.lbl_realtime_status.config(text=status_text)

            except Exception as e:
                self.lbl_realtime_status.config(text=f"Lỗi cập nhật trạng thái: {e}")

            # Lên lịch cho lần cập nhật tiếp theo
            self.after(200, self._update_status_bar_info)

    # -------------------------- Recording Listeners --------------------------

    def stop_listeners(self):
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener = None
        if self.keyboard_listener:
            self.keyboard_listener.stop()
            self.keyboard_listener = None

    def _on_mouse_click(self, x, y, button, pressed):
        if not self.recording or not pressed:
            return

        hwnd = hwnd_from_title(self.target_window_title)
        if not hwnd: return

        rect = get_window_rect(hwnd)
        if not rect: return
        left, top, right, bottom = rect

        # --- FIX: SỬ DỤNG WIN32API ĐỂ LẤY TỌA ĐỘ CHUỘT TUYỆT ĐỐI CHÍNH XÁC ---
        cursor_x, cursor_y = win32api.GetCursorPos()

        # Chỉ ghi nhận click trong phạm vi cửa sổ LOGICAL (đã scale)
        if not (left <= cursor_x < right and top <= cursor_y < bottom):
            return

            # 1. TÍNH TOÁN OFFSET LOGICAL PIXEL
        x_offset_logical = cursor_x - left
        y_offset_logical = cursor_y - top

        # 2. TÍNH TOÁN ĐỘ TRỄ
        current_time = time.time()

        if self.last_key_time == 0.0:
            delay = 0.0
        else:
            delay = current_time - self.last_key_time

        self.last_key_time = current_time

        # 3. TẠO STEP VÀ THÊM VÀO MACRO (Gồm DPI Scale)
        click_type = 'left_click'
        if button == mouse.Button.right:
            click_type = 'right_click'

        current_scale = get_dpi_scale_factor(hwnd)

        step = MacroStep('mouse',
                         key_value=click_type,
                         delay_after=delay,
                         x_offset=x_offset_logical,
                         y_offset=y_offset_logical,
                         dpi_scale=current_scale)

        self.after(0, self.add_macro_step, step)

    def _on_key_press(self, key):
        if not self.recording:
            # Cho phép ESC dừng ngay cả khi chưa vào run worker
            if key == Key.esc:
                self.cancel_run()
            return

        # SỬA: DỪNG GHI NGAY LẬP TỨC VỚI PHÍM ESC
        if key == Key.esc:
            # Ghi bước END cuối cùng với độ trễ
            current_time = time.time()
            if self.last_key_time == 0.0:
                delay = 0.0
            else:
                delay = current_time - self.last_key_time
            step = MacroStep('end', delay_after=delay)
            self.after(0, self.add_macro_step, step)

            self.after(0, self.stop_recording)
            return

        # Xử lý các phím bổ trợ (Ctrl, Alt, Shift)
        if key in [Key.ctrl_l, Key.ctrl_r, Key.alt_l, Key.alt_r, Key.shift_l, Key.shift_r]:
            self.current_modifiers.add(key)
            return

        # 1. TÍNH TOÁN ĐỘ TRỄ
        current_time = time.time()

        if self.last_key_time == 0.0:
            delay = 0.0
        else:
            delay = current_time - self.last_key_time

        self.last_key_time = current_time

        key_name = ""
        typ = ""
        col_index = None

        # 2. XỬ LÝ PHÍM
        if self.current_modifiers:
            typ = 'combo'
            modifier_names = sorted([str(m).replace("Key.", "") for m in self.current_modifiers])

            if hasattr(key, 'char') and key.char is not None:
                main_key = key.char.lower()
            else:
                main_key = str(key).replace("Key.", "")

            key_name = "+".join(modifier_names) + "+" + main_key

        else:
            if key == Key.insert:
                typ = 'col'
                col_index = self.current_col_index
                self.current_col_index += 1
                key_name = None

            elif hasattr(key, 'char') and key.char is not None:
                typ = 'key'
                key_name = key.char

            else:
                typ = 'key'
                key_name = str(key).replace("Key.", "")

        # 3. THÊM STEP
        if typ and typ != 'end':
            step = MacroStep(typ, key_value=key_name, col_index=col_index if typ == 'col' else None, delay_after=delay)
            self.after(0, self.add_macro_step, step)

    def _on_key_release(self, key):
        if key in [Key.ctrl_l, Key.ctrl_r, Key.alt_l, Key.alt_r, Key.shift_l, Key.shift_r]:
            if key in self.current_modifiers:
                self.current_modifiers.remove(key)

    def add_macro_step(self, step):
        step.item_idx = len(self.macro_steps)
        self.macro_steps.append(step)

        description = repr(step)
        item_id = self.tree_macro.insert("", tk.END, text=str(len(self.macro_steps)),
                                         values=(step.typ.upper(), description))
        step.item_id = item_id
        self.tree_macro.yview_moveto(1)

    def clear_macro(self):
        self.macro_steps.clear()
        self.populate_macro_tree()

    # -------------------------- Recording & Countdown --------------------------

    def record_macro(self):
        if self.recording: return
        if not self.target_window_title:
            messagebox.showwarning("Lỗi", "Vui lòng chọn cửa sổ mục tiêu ACSOFT trước.")
            return

        self.clear_macro()
        self.current_col_index = 0
        self.recording = True

        self.last_key_time = 0.0

        self.cancel_flag.clear()
        threading.Thread(target=self._countdown_and_record, daemon=True).start()

    def _on_escape_press(self, key):
        print(f"Key pressed during countdown: {key}")
        if key == keyboard.Key.esc:
            self.cancel_flag.set()
            self.cancel_run()

    def _countdown_and_record(self):
        self.stop_listeners()
        self.current_modifiers.clear()

        # SỬA: Khởi động listener cho ESC ngay từ đầu đếm ngược
        countdown_escape_listener = keyboard.Listener(on_press=self._on_escape_press)
        countdown_escape_listener.start()

        # SỬA: Sử dụng HUD thay cho cửa sổ đếm ngược
        self.hud_window = RecordingHUD(self, self.cancel_run)

        try:
            for i in range(5, 0, -1):
                if self.cancel_flag.is_set():
                    if self.hud_window: self.hud_window.close()
                    self.after(0, self.stop_recording)
                    return
                # Cập nhật HUD
                self.after(0, self.hud_window.update_status, f"Bắt đầu ghi sau: {i}s", "#87CEEB")
                self.update_idletasks()
                
                # SỬA: Thay thế time.sleep(1) bằng vòng lặp không chặn để ESC hoạt động ngay
                delay_start_time = time.time()
                while time.time() - delay_start_time < 1.0:
                    if self.cancel_flag.is_set():
                        break
                    time.sleep(0.05)

            if not self.cancel_flag.is_set():
                # Focus cửa sổ mục tiêu
                hwnd = hwnd_from_title(self.target_window_title)
                if hwnd:
                    bring_to_front(hwnd)

                # Cập nhật HUD sang trạng thái đang ghi
                self.after(0, self.hud_window.update_status, "🔴 ĐANG GHI... (Nhấn ESC để dừng)", "#FF4500")
                self.update_idletasks()
                self.after(100, self._start_listeners)
            else:
                if self.hud_window: self.hud_window.close()
        finally:
            # Đảm bảo listener tạm thời được dừng
            countdown_escape_listener.stop()

    def _start_listeners(self):
        if not self.recording: return

        self.mouse_listener = mouse.Listener(on_click=self._on_mouse_click, on_move=lambda x, y: None)
        self.mouse_listener.start()

        # SỬA: Listener chính cho các phím khác sẽ được khởi động ở đây.
        self.keyboard_listener = keyboard.Listener(on_press=self._on_key_press, on_release=self._on_key_release)
        self.keyboard_listener.start()

    def stop_recording(self):
        # Hàm này được gọi khi ESC được nhấn hoặc khi cancel_run được gọi lúc đang ghi
        if not self.recording: return

        self.stop_listeners()
        self.recording = False

        # Đóng HUD nếu nó tồn tại
        if self.hud_window: self.hud_window.close()

        # Chỉ hiển thị thông báo nếu không bị hủy bởi ESC (dấu hiệu của việc chạy)
        if not self.cancel_flag.is_set():
            messagebox.showinfo("Hoàn thành", f"Đã ghi xong macro với {len(self.macro_steps)} bước.")

    def on_test(self):
        self._run_macro(test_mode=True)

    def on_run_all(self):
        self._run_macro(test_mode=False)

    def on_cancel(self):
        self.cancel_run()

    def cancel_run(self):
        # Dừng cả Running và Recording
        is_running = self.btn_stop['state'] == 'normal'
        if is_running or self.recording:
            self.cancel_flag.set()
            if self.recording:
                self.after(0, self.stop_recording)
                self.after(0, self._reset_buttons)

    def _run_macro(self, test_mode):
        if not self.target_window_title:
            messagebox.showwarning("Lỗi", "Vui lòng chọn cửa sổ mục tiêu hợp lệ trước.")
            return

        hwnd = hwnd_from_title(self.target_window_title)
        if not hwnd:
            messagebox.showwarning("Lỗi",
                                   f"Không tìm thấy cửa sổ: '{self.target_window_title}'. Vui lòng làm mới và chọn lại.")
            return

        if self.df_csv.empty:
            messagebox.showwarning("Lỗi", "Vui lòng load file CSV có dữ liệu.")
            return
        if not self.macro_steps:
            messagebox.showwarning("Lỗi", "Chưa có bước macro nào được ghi.")
            return

        self.cancel_flag.clear()
        self.btn_test.config(state='disabled')
        self.btn_runall.config(state='disabled')
        self.btn_stop.config(state='normal')

        # CHỈ FOCUS CỬA SỔ, KHÔNG MAXIMIZE
        if not bring_to_front(hwnd):
            messagebox.showwarning("Cảnh báo", "Không thể đưa cửa sổ lên foreground. Macro vẫn tiếp tục chạy.")

        threading.Thread(target=self._countdown_and_run_worker, args=(test_mode,), daemon=True).start()

    def _countdown_and_run_worker(self, test_mode):
        # SỬA: Sử dụng HUD thay cho cửa sổ đếm ngược
        # Gán pause_event từ HUD cho luồng chạy macro
        self.pause_event = threading.Event()
        self.pause_event.set() # Mặc định là không pause

        # SỬA: Khởi động listener cho ESC ngay từ đầu đếm ngược
        countdown_escape_listener = keyboard.Listener(on_press=self._on_escape_press)
        countdown_escape_listener.start()

        self.hud_window = RecordingHUD(self, self.cancel_run)

        try:
            for i in range(5, 0, -1):
                if self.cancel_flag.is_set():
                    if self.hud_window: self.hud_window.close()
                    self.after(0, self._reset_buttons)
                    return
                # Cập nhật HUD
                self.after(0, self.hud_window.update_status, f"Bắt đầu chạy sau: {i}s", "#87CEEB")
                self.update_idletasks()

                # SỬA: Thay thế time.sleep(1) bằng vòng lặp không chặn để ESC hoạt động ngay
                delay_start_time = time.time()
                while time.time() - delay_start_time < 1.0:
                    if self.cancel_flag.is_set():
                        break
                    time.sleep(0.05)

            if not self.cancel_flag.is_set():
                # Cập nhật HUD sang trạng thái đang chạy
                self.hud_window.pause_event = self.pause_event # Liên kết event
                # SỬA LỖI: Sử dụng lambda để gọi pack() với các đối số từ khóa một cách chính xác
                self.after(0, lambda: self.hud_window.pause_button.pack(side="left", padx=(0, 10)))

                self.after(0, self.hud_window.update_status, "▶️ ĐANG CHẠY, ẤN ESC ĐỂ DỪNG...", "#98FB98")
                self._macro_run_worker(test_mode)
            else:
                # Nếu bị hủy trong lúc đếm ngược, chỉ cần reset các nút
                self.after(0, self._reset_buttons)
        finally:
            # SỬA: Đảm bảo listener chỉ được dừng sau khi _macro_run_worker đã chạy xong
            countdown_escape_listener.stop()

    def _macro_run_worker(self, test_mode):
        try:
            hwnd = hwnd_from_title(self.target_window_title)
            if not hwnd: return

            use_recorded_speed = self.speed_mode.get() == 1
            custom_delay_s = self.spin_speed_val.get() / 1000.0
            row_delay = self.spin_between_val.get()

            rows_to_run = self.df_csv.iloc[:1] if test_mode else self.df_csv

            for row_index, row_data in rows_to_run.iterrows():
                if self.cancel_flag.is_set():
                    break

                # SỬA: Thêm logic kiểm tra Pause
                if self.pause_event:
                    self.pause_event.wait() # Thread sẽ dừng ở đây nếu event bị clear()

                csv_item_id = f"csv_{row_index}"
                self.after(0, self._highlight_csv_row, csv_item_id)
                self.after(0, self.tree_csv.focus, csv_item_id)
                self.after(0, self.tree_csv.see, csv_item_id)

                self.after(0, self.lbl_status.config,
                           {'text': f"Đang chạy dòng CSV số: {row_index + 1}/{len(rows_to_run)}..."})

                self._run_macro_for_row(hwnd, row_data.tolist(), use_recorded_speed, custom_delay_s)

                if test_mode: break

                if self.cancel_flag.is_set():
                    break

                self.after(0, self._unhighlight_csv_row, csv_item_id)

                # SỬA: Thay thế time.sleep() bằng vòng lặp không chặn để ESC hoạt động ngay lập tức
                delay_start_time = time.time()
                while time.time() - delay_start_time < row_delay:
                    if self.cancel_flag.is_set():
                        break
                    # Chờ một khoảng ngắn và kiểm tra lại, thay vì ngủ một giấc dài
                    time.sleep(0.05)

            if not self.cancel_flag.is_set():
                self.after(0, self.lbl_status.config, {'text': "Hoàn thành!"})
                messagebox.showinfo("Hoàn thành", "Chạy macro hoàn tất.")
            else:
                self.after(0, self.lbl_status.config, {'text': "Đã hủy bởi người dùng."})
                messagebox.showinfo("Đã hủy", "Quá trình chạy đã bị hủy bởi người dùng.")

        except Exception as e:
            self.after(0, self.lbl_status.config, {'text': f"LỖI: {str(e)}"})
            messagebox.showerror("Lỗi khi chạy", str(e))
        finally:
            self.after(0, self._reset_buttons)
            if self.hud_window: self.after(0, self.hud_window.close)
            self.after(0, self._clear_macro_highlights)
            self.cancel_flag.clear()

    def _run_macro_for_row(self, hwnd, row_data, use_recorded_speed, custom_delay_s):
        for step in self.macro_steps:
            if self.cancel_flag.is_set():
                return

            # SỬA: Thêm logic kiểm tra Pause
            if self.pause_event:
                self.pause_event.wait()

            self.after(0, self._highlight_macro_step, step)

            if step.typ == MacroStepType.COLUMN_DATA.value:
                col_index = step.col_index
                if col_index is not None and col_index < len(row_data):
                    value = row_data[col_index]
                    self.after(
                        0, self.lbl_status.config, {"text": f"ĐANG GỬI CỘT #{col_index + 1}: '{value}'"}
                    )
                    send_char_to_hwnd(hwnd, str(value))

            elif step.typ == MacroStepType.KEY_PRESS.value:
                self.after(0, self.lbl_status.config, {'text': f"ĐANG GỬI PHÍM: {step.key_value.upper()}"})
                send_key_to_hwnd(hwnd, step.key_value)

            elif step.typ == MacroStepType.KEY_COMBO.value:
                self.after(0, self.lbl_status.config, {'text': f"ĐANG GỬI TỔ HỢP PHÍM: {step.key_value.upper()}"})
                send_combo_to_hwnd(hwnd, step.key_value)

            elif step.typ == MacroStepType.MOUSE_CLICK.value:
                # SỬ DỤNG TỌA ĐỘ CHUẨN ĐỂ HIỂN THỊ TRẠNG THÁI
                scale = step.dpi_scale if step.dpi_scale > 0 else 1.0
                if scale == 0: scale = 1.0 # Defensive check

                x_norm = int(step.x_offset_logical / scale)
                y_norm = int(step.y_offset_logical / scale)

                self.after(0, self.lbl_status.config,
                           {'text': f"ĐANG CLICK: {step.key_value.upper()} tại Offset Chuẩn ({x_norm}px, {y_norm}px)"})
                send_mouse_click(hwnd, step.x_offset_logical, step.y_offset_logical, step.key_value, step.dpi_scale)

            elif step.typ == 'end':
                self.after(0, self.lbl_status.config, {'text': "Đã kết thúc dòng, đang chờ chuyển dòng..."})

            delay = step.delay_after if use_recorded_speed else custom_delay_s

            if delay > 0:
                # SỬA: Thay thế time.sleep() bằng vòng lặp không chặn để ESC hoạt động ngay lập tức
                delay_start_time = time.time()
                while time.time() - delay_start_time < delay:
                    if self.cancel_flag.is_set():
                        break
                    # Chờ một khoảng rất ngắn (50ms) và kiểm tra lại
                    time.sleep(0.05)

    # -------------------------- UI Helpers (Đã Sửa - KHẮC PHỤC LỖI LAYOUT) --------------------------

    def _create_treeview(self, container_frame, max_cols=20):
        """Tạo Treeview cho CSV Data."""
        columns = [f"Cột {i + 1}" for i in range(max_cols)]
        tree = ttk.Treeview(container_frame, columns=columns, show='headings')

        # Thêm Scrollbar
        vsb = ttk.Scrollbar(container_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Đặt layout
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        # Cấu hình container_frame để Treeview chiếm hết không gian
        container_frame.grid_rowconfigure(0, weight=1)
        container_frame.grid_columnconfigure(0, weight=1)

        # SỬA: CHỈ TRẢ VỀ TREEVIEW
        return tree

    def _create_treeview_macro(self, container_frame):
        """Tạo Treeview cho Macro Steps."""
        tree = ttk.Treeview(container_frame, columns=('type', 'description'), show='tree headings',
                            style='Macro.Treeview')

        # Thẻ để highlight dòng đang chạy
        tree.tag_configure('highlight', background='#FFA07A', font=('Arial', 9, 'bold'))

        tree.heading("#0", text="STT", anchor='center')
        tree.column("#0", width=50, stretch=tk.NO, anchor='center')

        tree.heading('type', text="Loại", anchor='w')
        tree.column('type', width=80, stretch=tk.NO, anchor='w')

        tree.heading('description', text="Mô tả chi tiết", anchor='w')
        tree.column('description', width=300, stretch=tk.YES, anchor='w')

        # Thêm Scrollbar
        vsb = ttk.Scrollbar(container_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Đặt layout
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        # Cấu hình container_frame để Treeview chiếm hết không gian
        container_frame.grid_rowconfigure(0, weight=1)
        container_frame.grid_columnconfigure(0, weight=1)

        # Bắt sự kiện double click để sửa macro
        tree.bind('<Double-1>', lambda e: self.edit_macro_step())

        # SỬA: CHỈ TRẢ VỀ TREEVIEW
        return tree

    # -------------------------- Manual Macro Functions --------------------------

    def add_manual_step(self, step_type):
        """Thêm bước macro thủ công với giá trị mặc định và mở cửa sổ sửa."""
        # Dựa vào số bước hiện tại để gợi ý cột tiếp theo
        next_col_index = sum(1 for step in self.macro_steps if step.typ == 'col')
        current_scale = get_dpi_scale_factor(None)  # Lấy DPI hệ thống

        if step_type == 'col':
            initial_step = MacroStep(step_type, col_index=next_col_index)
        elif step_type == 'key':
            initial_step = MacroStep(step_type, key_value='ENTER')
        elif step_type == 'combo':
            initial_step = MacroStep(step_type, key_value='CTRL+C')
        elif step_type == 'mouse':
            # Gợi ý tọa độ (100, 100) được scale theo DPI hiện tại để người dùng dễ sửa
            initial_x = 100.0 * current_scale
            initial_y = 100.0 * current_scale
            initial_step = MacroStep(step_type, key_value='left_click', x_offset=initial_x, y_offset=initial_y,
                                     dpi_scale=current_scale)
        elif step_type == 'end':
            initial_step = MacroStep('end')
        else:
            return

        # Thêm vào list và treeview
        self.add_macro_step(initial_step)

        # Tự động mở cửa sổ sửa để người dùng điều chỉnh
        self.after(50, lambda: self.edit_macro_step(initial_step.item_id))

    def edit_macro_step(self, item_id=None):
        selected_item = item_id if item_id else self.tree_macro.focus()
        if not selected_item:
            messagebox.showwarning("Lỗi", "Vui lòng chọn một bước macro để sửa.")
            return

        selected_step = next((step for step in self.macro_steps if step.item_id == selected_item), None)
        if not selected_step: return

        edit_win = tk.Toplevel(self)
        edit_win.title(f"Sửa Bước Macro #{selected_step.item_idx + 1}")
        edit_win.transient(self)
        edit_win.grab_set()

        frame = ttk.Frame(edit_win, padding="10")
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Loại Bước:").grid(row=0, column=0, sticky='w', pady=5)
        typ_var = tk.StringVar(value=selected_step.typ)
        combo_typ = ttk.Combobox(frame, textvariable=typ_var, state="readonly",
                                 values=['col', 'key', 'combo', 'mouse', 'end'])
        combo_typ.grid(row=0, column=1, sticky='ew', padx=5, pady=5)

        # Giá trị cho key/click
        ttk.Label(frame, text="Giá trị (Key/Click Type):").grid(row=1, column=0, sticky='w', pady=5)
        key_val_var = tk.StringVar(value=str(selected_step.key_value) if selected_step.key_value is not None else "")
        entry_key_val = ttk.Entry(frame, textvariable=key_val_var)
        entry_key_val.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        # Chỉ số cột
        ttk.Label(frame, text="Chỉ số Cột (0, 1, 2...):").grid(row=2, column=0, sticky='w', pady=5)
        col_index_var = tk.StringVar(value=str(selected_step.col_index) if selected_step.col_index is not None else "")
        entry_col_index = ttk.Entry(frame, textvariable=col_index_var)
        entry_col_index.grid(row=2, column=1, sticky='ew', padx=5, pady=5)

        # Tọa độ Offset X (Logical Pixel)
        ttk.Label(frame, text="Offset X (Logical Pixel):").grid(row=3, column=0, sticky='w', pady=5)
        x_offset_var = tk.StringVar(
            value=f"{selected_step.x_offset_logical:.2f}" if selected_step.x_offset_logical is not None else "")
        entry_x_offset = ttk.Entry(frame, textvariable=x_offset_var)
        entry_x_offset.grid(row=3, column=1, sticky='ew', padx=5, pady=5)

        # Tọa độ Offset Y (Logical Pixel)
        ttk.Label(frame, text="Offset Y (Logical Pixel):").grid(row=4, column=0, sticky='w', pady=5)
        y_offset_var = tk.StringVar(
            value=f"{selected_step.y_offset_logical:.2f}" if selected_step.y_offset_logical is not None else "")
        entry_y_offset = ttk.Entry(frame, textvariable=y_offset_var)
        entry_y_offset.grid(row=4, column=1, sticky='ew', padx=5, pady=5)

        # DPI Scale (chỉ để xem, không sửa)
        ttk.Label(frame, text=f"DPI Scale Ghi (%):").grid(row=5, column=0, sticky='w', pady=5)
        ttk.Label(frame, text=f"{int(selected_step.dpi_scale * 100)}%").grid(row=5, column=1, sticky='w', padx=5,
                                                                             pady=5)

        # Độ trễ
        ttk.Label(frame, text="Độ trễ sau (ms - 10-10000):").grid(row=6, column=0, sticky='w', pady=5)
        delay_ms = int(selected_step.delay_after * 1000)
        delay_var = tk.IntVar(value=delay_ms)
        spin_delay = ttk.Spinbox(frame, from_=10, to=10000, increment=10,
                                 textvariable=delay_var, width=10)
        spin_delay.grid(row=6, column=1, sticky='ew', padx=5, pady=5)

        def save_changes():
            try:
                new_delay_ms = delay_var.get()
                if new_delay_ms < 10 or new_delay_ms > 10000:
                    raise ValueError("Độ trễ phải nằm trong khoảng 10ms - 10000ms.")

                selected_step.typ = typ_var.get().lower()
                selected_step.delay_after = new_delay_ms / 1000.0

                # Reset các giá trị phụ thuộc vào type
                selected_step.col_index = None
                selected_step.key_value = None
                selected_step.x_offset_logical = None
                selected_step.y_offset_logical = None

                if selected_step.typ == 'col':
                    selected_step.col_index = int(col_index_var.get())
                elif selected_step.typ in ['key', 'combo']:
                    selected_step.key_value = key_val_var.get().upper().strip()
                elif selected_step.typ == 'mouse':
                    selected_step.key_value = key_val_var.get().lower().strip()
                    selected_step.x_offset_logical = float(x_offset_var.get())
                    selected_step.y_offset_logical = float(y_offset_var.get())

                # Cập nhật hiển thị
                self.tree_macro.item(selected_item, values=(selected_step.typ.upper(), repr(selected_step)))

                edit_win.destroy()
                messagebox.showinfo("Thành công", "Đã sửa bước macro thành công.")
            except ValueError as e:
                messagebox.showerror("Lỗi dữ liệu", f"Lỗi: {e}\nKiểm tra định dạng (số nguyên/số thực) và phạm vi.")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu: {e}")

        btn_save = ttk.Button(edit_win, text="Lưu Thay Đổi", command=save_changes, style='Accent.TButton')
        btn_save.pack(pady=10)

        edit_win.protocol("WM_DELETE_WINDOW", edit_win.destroy)
        self.wait_window(edit_win)

    def delete_macro_step(self):
        selected_item = self.tree_macro.focus()
        if not selected_item:
            messagebox.showwarning("Lỗi", "Vui lòng chọn một bước macro để xóa.")
            return

        if messagebox.askyesno("Xác nhận Xóa", "Bạn có chắc chắn muốn xóa bước macro này?"):
            try:
                selected_step = next((step for step in self.macro_steps if step.item_id == selected_item), None)
                if selected_step:
                    self.macro_steps.remove(selected_step)

                self.tree_macro.delete(selected_item)

                self.populate_macro_tree()

            except Exception as e:
                messagebox.showerror("Lỗi Xóa", f"Không thể xóa bước macro. Lỗi: {e}")

    def populate_macro_tree(self):
        self.tree_macro.delete(*self.tree_macro.get_children())
        # Cần reset current_col_index khi populate
        self.current_col_index = 0
        for idx, step in enumerate(self.macro_steps):
            step.item_idx = idx
            if step.typ == 'col':
                step.col_index = self.current_col_index
                self.current_col_index += 1

            description = repr(step)
            item_id = self.tree_macro.insert("", tk.END, text=str(idx + 1), values=(step.typ.upper(), description))
            step.item_id = item_id

    def on_window_select(self, event):
        self.target_window_title = self.combo_windows.get()

    def refresh_windows(self):
        titles = []

        def enum_handler(hwnd, titles):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title not in titles:
                    titles.append(title)
            return True

        win32gui.EnumWindows(enum_handler, titles)

        titles.sort()
        self.combo_windows['values'] = titles
        if titles:
            initial_select = None
            for title in titles:
                if "acsoft" in title.lower() or "kế toán" in title.lower() or "việt tín" in title.lower():
                    initial_select = title
                    break
            if initial_select:
                self.combo_windows.set(initial_select)
                self.target_window_title = initial_select
            elif titles:
                self.combo_windows.set(titles[0])
                self.target_window_title = titles[0]
        else:
            self.combo_windows.set("Không tìm thấy cửa sổ")
            self.target_window_title = ""

    def browse_ac(self):
        p = filedialog.askopenfilename(title="Chọn file chạy ACSOFT (.exe)",
                                       filetypes=(("Executable files", "*.exe"), ("All files", "*.*")))
        if p:
            self.txt_acpath.delete(0, tk.END)
            self.txt_acpath.insert(0, p)
            self.acsoft_path = p

    def open_ac(self):
        p = self.txt_acpath.get().strip()
        if not p or not os.path.isfile(p):
            messagebox.showwarning("Lỗi", "Chưa chọn file exe hợp lệ.")
            return
        try:
            subprocess.Popen([p])
            messagebox.showinfo("Đã mở", "Đã mở ACSOFT (chờ phần mềm khởi động).")
            self.after(2000, self.refresh_windows)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Mở không thành công: {e}")

    def browse_csv(self):
        p = filedialog.askopenfilename(title="Chọn file CSV chứa dữ liệu",
                                       filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        if p:
            self.txt_csv.delete(0, tk.END)
            self.txt_csv.insert(0, p)
            self.csv_path = p
            self.load_csv_data(p)

    def load_csv_data(self, path):
        delimiter = self.txt_delimiter.get().strip()
        if not delimiter:
            messagebox.showwarning("Lỗi", "Vui lòng nhập ký tự phân cách (Delimiter).")
            return

        try:
            self.df_csv = pd.read_csv(path,
                                      header=None,
                                      dtype=str,
                                      keep_default_na=False,
                                      sep=delimiter)

            self.tree_csv.delete(*self.tree_csv.get_children())

            num_cols = len(self.df_csv.columns)
            columns = [f"Cột {i + 1}" for i in range(num_cols)]
            self.tree_csv.config(columns=columns)

            self.tree_csv.heading("#0", text="STT", anchor='center')
            self.tree_csv.column("#0", width=50, stretch=tk.NO, anchor='center')

            for i, col_name in enumerate(columns):
                self.tree_csv.heading(col_name, text=col_name, anchor='w')
                self.tree_csv.column(col_name, width=120, stretch=tk.NO)

            rows_to_display = self.df_csv.head(10)
            for idx, row in rows_to_display.iterrows():
                values = [row[c] for c in row.index]
                self.tree_csv.insert("", tk.END, text=str(idx + 1), values=values, iid=f"csv_{idx}")

        except Exception as e:
            self.df_csv = pd.DataFrame()
            messagebox.showerror("Lỗi đọc CSV", f"Không thể đọc file CSV bằng delimiter '{delimiter}'. Lỗi: {e}")
            self.tree_csv.delete(*self.tree_csv.get_children())

    def _reset_buttons(self):
        self.btn_test.config(state='normal')
        self.btn_runall.config(state='normal')
        self.btn_stop.config(state='disabled')
        self.lbl_status.config(text="Chờ...")
        if self.hud_window:
            self.hud_window.close()
            self.hud_window = None
        self.pause_event = None

    def _highlight_macro_step(self, step):
        for item_id in self.tree_macro.get_children():
            current_tags = list(self.tree_macro.item(item_id, 'tags'))
            if 'highlight' in current_tags:
                current_tags.remove('highlight')
                self.tree_macro.item(item_id, tags=tuple(current_tags))

        current_tags = list(self.tree_macro.item(step.item_id, 'tags'))
        if 'highlight' not in current_tags:
            current_tags.append('highlight')
        self.tree_macro.item(step.item_id, tags=tuple(current_tags))
        self.tree_macro.see(step.item_id)

    def _clear_macro_highlights(self):
        for item_id in self.tree_macro.get_children():
            current_tags = list(self.tree_macro.item(item_id, 'tags'))
            if 'highlight' in current_tags:
                current_tags.remove('highlight')
                self.tree_macro.item(item_id, tags=tuple(current_tags))

    def _highlight_csv_row(self, item_id):
        for iid in self.tree_csv.get_children():
            current_tags = list(self.tree_csv.item(iid, 'tags'))
            if 'highlight' in current_tags:
                current_tags.remove('highlight')
                self.tree_csv.item(iid, tags=tuple(current_tags))

        current_tags = list(self.tree_csv.item(item_id, 'tags'))
        if 'highlight' not in current_tags:
            current_tags.append('highlight')
        self.tree_csv.item(item_id, tags=tuple(current_tags))

    def _unhighlight_csv_row(self, item_id):
        current_tags = list(self.tree_csv.item(item_id, 'tags'))
        if 'highlight' in current_tags:
            current_tags.remove('highlight')
            self.tree_csv.item(item_id, tags=tuple(current_tags))

    def save_macro(self):
        if not self.macro_steps:
            messagebox.showwarning("Lỗi", "Chưa có macro nào để lưu.")
            return

        p = filedialog.asksaveasfilename(defaultextension=".json",
                                         filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
                                         title="Lưu file Macro & Cài đặt")
        if p:
            try:
                macro_data = [step.to_dict() for step in self.macro_steps]
                settings_data = self._collect_app_settings()

                full_data = {
                    "app_settings": settings_data,
                    "macro_steps": macro_data
                }

                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(full_data, f, indent=4)
                messagebox.showinfo("Thành công", f"Đã lưu macro và cài đặt thành công vào:\n{p}")
            except Exception as e:
                messagebox.showerror("Lỗi Lưu", f"Không thể lưu file macro. Lỗi: {e}")

    def load_macro(self):
        p = filedialog.askopenfilename(filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
                                       title="Mở file Macro & Cài đặt")
        if p:
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    full_data = json.load(f)

                macro_data = full_data.get("macro_steps", [])
                new_steps = [MacroStep.from_dict(d) for d in macro_data]
                self.macro_steps = new_steps
                self.populate_macro_tree()

                settings_data = full_data.get("app_settings", {})
                self._apply_app_settings(settings_data)

                messagebox.showinfo("Thành công", f"Đã tải macro và cài đặt thành công từ:\n{p}")
            except Exception as e:
                messagebox.showerror("Lỗi Mở", f"Không thể mở file macro. Lỗi: {e}")

    def _collect_app_settings(self):
        return {
            "acsoft_path": self.txt_acpath.get() if self.txt_acpath else "",
            "csv_path": self.txt_csv.get() if self.txt_csv else "",
            "delimiter": self.txt_delimiter.get() if self.txt_delimiter else ";",
            "target_window_title": self.target_window_title,
            "speed_mode": self.speed_mode.get(),
            "custom_speed_ms": self.spin_speed_val.get(),
            "delay_between_rows_s": self.spin_between_val.get(),
            "show_realtime_status": self.show_realtime_status.get(),
        }

    def _apply_app_settings(self, settings):
        try:
            self.txt_acpath.delete(0, tk.END);
            self.txt_acpath.insert(0, settings.get("acsoft_path", ""))
            self.txt_csv.delete(0, tk.END);
            self.txt_csv.insert(0, settings.get("csv_path", ""))
            self.txt_delimiter.delete(0, tk.END);
            self.txt_delimiter.insert(0, settings.get("delimiter", ";"))
            target_title = settings.get("target_window_title", "")
            if target_title in self.combo_windows['values']:
                self.combo_windows.set(target_title)
                self.target_window_title = target_title
            self.speed_mode.set(settings.get("speed_mode", 1))
            self.spin_speed_val.set(settings.get("custom_speed_ms", 500))
            self.spin_between_val.set(settings.get("delay_between_rows_s", 2))

            self.show_realtime_status.set(settings.get("show_realtime_status", True))
            self._toggle_realtime_status()

            csv_path = settings.get("csv_path")
            if csv_path and os.path.exists(csv_path):
                self.load_csv_data(csv_path)
        except Exception as e:
            print(f"Lỗi khi áp dụng settings: {e}")


if __name__ == "__main__":
    app = MacroApp()
    app.mainloop()