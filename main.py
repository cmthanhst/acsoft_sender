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
from tkinter import font as tkFont
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
        style.configure("HUD.TLabel", background="#282c34", foreground="white", font=("Courier", 10))
        style.configure("HUD.TButton", font=("Courier", 9, "bold"), foreground="white")
        style.map("HUD.TButton",
                  background=[('active', '#e06c75'), ('!active', '#d15660')],
                  foreground=[('active', 'white')])
        style.configure("Pause.TButton", font=("Courier", 9, "bold"))
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

is_collapsed = True
class MacroApp(ThemedTk):
    def __init__(self):
        super().__init__(theme="arc")

        # SỬA: Giữ cho cửa sổ ứng dụng luôn ở trên cùng
        self.attributes('-topmost', True)

        # SỬA: Bỏ thanh tiêu đề mặc định của Windows
        self.overrideredirect(True)
        self._offset_x = 0
        self._offset_y = 0
        
        # SỬA: Thêm biến trạng thái cho việc thu gọn cửa sổ
        self.is_collapsed = False
        self.normal_geometry = ""

        # Determine the font family
        self.preferred_font_family = "Courier"
        if self.preferred_font_family not in tkFont.families():
            self.preferred_font_family = "Courier"

        self.option_add("*font", (self.preferred_font_family, 9)) # SỬA: Tải logo từ chuỗi Base64
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

        # SỬA: Đặt lại title để hiển thị trên taskbar
        self.title("Việt Tín Auto Sender V2025.04")
        self._offset_x = 0
        self._offset_y = 0
        # self.title("...") # Không cần title nữa
        self._offset_x = 0 # Biến cho việc kéo thả cửa sổ
        self._offset_y = 0 # Biến cho việc kéo thả cửa sổ
        self.geometry("1200x800")
        
        # SỬA: Tạm thời ẩn cửa sổ để áp dụng style một cách an toàn
        self.withdraw()

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
        self.dark_mode_var = tk.BooleanVar(value=True) # SỬA: Mặc định là giao diện tối
        self.show_realtime_status = tk.BooleanVar(value=False) # SỬA: Tắt theo mặc định để giảm lag

        self.txt_delimiter = None
        self.txt_acpath = None
        self.txt_csv = None
        self.lbl_realtime_status = None
        self.dark_mode_btn = None # SỬA: Thêm biến cho nút dark mode

        self.mouse_listener = None
        self.keyboard_listener = None
        self.current_modifiers = set()

        self.pause_event = None
        self.hud_window = None
        self.realtime_status_frame = None

        # SỬA: Khởi tạo self.style TRƯỚC khi gọi setup_ui() để tránh lỗi AttributeError
        self.style = ttk.Style()
        self.style.map("Macro.Treeview", background=[("selected", "#1e90ff"), ("active", "#e1e1e1")])
        self.style.configure("Macro.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", font=("Courier", 9))

        # SỬA: Áp dụng font cho Macro.Treeview
        self.style.configure("Macro.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", font=(self.preferred_font_family, 9))
        self.style.map("Highlight.Treeview", background=[("selected", "#FFA07A"), ("active", "#e1e1e1")]) # This style is not used in the provided code
        
        # SỬA: Cấu hình font cho các widget ttk chung
        self.style.configure('TButton', font=(self.preferred_font_family, 9))
        self.style.configure('Accent.TButton', font=(self.preferred_font_family, 9, 'bold'))
        self.style.configure('TLabelframe.Label', font=(self.preferred_font_family, 9, 'bold')) # Tiêu đề LabelFrame
        self.style.configure('TEntry', font=(self.preferred_font_family, 9))
        self.style.configure('TCombobox', font=(self.preferred_font_family, 9))
        self.style.configure('TSpinbox', font=(self.preferred_font_family, 9))
        self.style.configure('TRadiobutton', font=(self.preferred_font_family, 9))
        self.style.configure('TCheckbutton', font=(self.preferred_font_family, 9))
        self.style.configure('TLabel', font=(self.preferred_font_family, 9)) # Cho ttk.Label

        self.setup_ui() # Call setup_ui after preferred_font_family is set

        self.protocol("WM_DELETE_WINDOW", self.on_app_close)
        # self.resizable(False, False) # SỬA: Xóa bỏ dòng này để tránh xung đột
        self.resizable(False, False) # SỬA: Vô hiệu hóa thay đổi kích thước cửa sổ

        # BẮT ĐẦU VÒNG LẶP CẬP NHẬT TRẠNG THÁI REAL-TIME
        self._update_status_bar_info()

        # SỬA: Áp dụng font cho tất cả các widget hiện có và widget được tạo sau
        self._apply_font_recursively(self)

        # SỬA: Áp dụng theme ban đầu.
        # Vì toggle_dark_mode() sẽ đảo ngược trạng thái, chúng ta cần đặt giá trị ban đầu
        # là ngược lại với giá trị mong muốn (muốn tối -> đặt là sáng) rồi mới gọi hàm.
        self.dark_mode_var.set(False) # Đặt là 'sáng' để lần gọi đầu tiên sẽ chuyển thành 'tối'.
        self.toggle_dark_mode()
        # SỬA: Gán sự kiện để bo góc cửa sổ khi kích thước thay đổi (chạy 1 lần lúc đầu)
        self.bind("<Configure>", self._apply_rounding_region)
        
        # SỬA: Áp dụng việc xóa title bar ngay lập tức và sau đó hiển thị lại cửa sổ
        # self._remove_title_bar()
        self.update_idletasks()
        self.deiconify()


        # SỬA: Hiển thị icon trên taskbar để có thể minimize/restore
        # self.after(10, lambda: self._set_appwindow())
        # SỬA: Hợp nhất các lệnh tùy chỉnh cửa sổ và gọi sau khi mainloop bắt đầu
        self.after(10, self._setup_custom_window)
    #viết 1 hàm minimize được gọi khi ấn vào nút minimize. sẽ thu cửa sổ nhỏ lại chỉ còn kích thước vừa logo, tiêu đề app và hàng nút thu nhỏ, đóng app.
    def _set_appwindow(self):
        
        """
        SỬA: Thu nhỏ cửa sổ về kích thước chỉ còn thanh tiêu đề tùy chỉnh.
        """
        if not self.is_collapsed:
            self.normal_geometry = self.geometry() # Lưu lại kích thước ban đầu
            # Lấy chiều cao của thanh tiêu đề
            self.update_idletasks()
            header_height = self.header_frame.winfo_height()+20
            
            # --- SỬA: TÍNH TOÁN VỊ TRÍ GÓC DƯỚI BÊN PHẢI ---
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            window_width = 300
            # Đặt vị trí x: (chiều rộng màn hình - chiều rộng cửa sổ - padding)
            x_pos = screen_width - window_width - 10
            # Đặt vị trí y: (chiều cao màn hình - chiều cao cửa sổ - chiều cao taskbar ước tính)
            y_pos = screen_height - header_height - 40
            self.geometry(f"{window_width}x{header_height}+{x_pos}+{y_pos}") # Kích thước nhỏ, đặt ở góc dưới bên phải
            self.is_collapsed = True
            # Ẩn các phần còn lại của ứng dụng
            self.top_controls_frame.grid_remove()
            self.g_data_macro.grid_remove()
            self.g5.grid_remove()
            self.realtime_status_frame.grid_remove()
            self.minimize_btn.config(text=" □ ") # Đổi nút thành restore
        else:
            self.geometry(self.normal_geometry) # Khôi phục kích thước ban đầu
            self.is_collapsed = False
            # Hiện lại các phần của ứng dụng
            self.top_controls_frame.grid()
            self.g_data_macro.grid()
            self.g5.grid()
            self.realtime_status_frame.grid()
            self.minimize_btn.config(text=" _ ") # Đổi nút thành minimize
            
    def _on_title_bar_press(self, event):
        self._offset_x = event.x
        self._offset_y = event.y

    def _on_title_bar_drag(self, event):
        x = self.winfo_pointerx() - self._offset_x
        y = self.winfo_pointery() - self._offset_y
        self.geometry(f"+{x}+{y}")

    def minimize_window(self):
        self._set_appwindow()

    def _apply_rounding_region(self, event):
        """SỬA: Áp dụng bo góc bằng phương pháp SetWindowRgn, đáng tin cậy hơn."""
        try:
            hwnd = self.winfo_id()
            width, height = self.winfo_width(), self.winfo_height()
            # Tạo một vùng hình chữ nhật bo góc với bán kính 20x20
            rgn = win32gui.CreateRoundRectRgn(0, 0, width, height, 20, 20)
            win32gui.SetWindowRgn(hwnd, rgn, True)
        except Exception:
            pass

    def _setup_custom_window(self):
        """SỬA: Hợp nhất các hàm tùy chỉnh cửa sổ (xóa title bar, hiện trên taskbar)."""
        try:
            hwnd = self.winfo_id()            
            # SỬA: Xóa bỏ triệt để các style liên quan đến title bar và viền
            hwnd = self.winfo_id()
            # 1. Xóa bỏ triệt để các style liên quan đến title bar và viền
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            style &= ~(win32con.WS_CAPTION | win32con.WS_THICKFRAME | win32con.WS_MINIMIZEBOX | win32con.WS_MAXIMIZEBOX)
            win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
            # SỬA: Buộc cửa sổ vẽ lại với style mới ngay lập tức

            # 2. Ép cửa sổ hiển thị trên taskbar
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            ex_style |= win32con.WS_EX_APPWINDOW
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)

            # 3. Buộc cửa sổ vẽ lại với style mới ngay lập tức
            win32gui.SetWindowPos(hwnd, None, 0, 0, 0, 0, win32con.SWP_FRAMECHANGED | win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOOWNERZORDER)
            hwnd = self.winfo_id()
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            style |= win32con.WS_EX_APPWINDOW
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
        except Exception:
            pass
    # --- KẾT THÚC THÊM HÀM ---

    def on_app_close(self):
        self.stop_listeners()
        self.destroy()

    def setup_ui(self):
        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=10)
        main_frame.grid_rowconfigure(2, weight=1)  # Dòng chứa CSV/Macro (g_data_macro)
        main_frame.grid_columnconfigure(0, weight=1)

        # SỬA: header_frame giờ là thanh tiêu đề tùy chỉnh
        header_frame = tk.Frame(main_frame) # Dùng tk.Frame để dễ tùy chỉnh màu nền
        self.header_frame = header_frame # Lưu tham chiếu
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header_frame.grid_columnconfigure(1, weight=1) # Cột chứa tiêu đề sẽ co giãn

        if self.header_logo:
            logo_label = tk.Label(header_frame, image=self.header_logo)
            logo_label.grid(row=0, column=0, sticky='w', padx=5, pady=5)
            logo_label.bind("<Button-1>", self._on_title_bar_press)
            logo_label.bind("<B1-Motion>", self._on_title_bar_drag)

        title_label = tk.Label(header_frame, text="Việt Tín Auto Sender V2025.04", font=(self.preferred_font_family, 10, "bold"))
        title_label.grid(row=0, column=1, sticky='w', padx=10)
        title_label.bind("<Button-1>", self._on_title_bar_press)
        title_label.bind("<B1-Motion>", self._on_title_bar_drag)

        # SỬA: Thay Checkbutton bằng Button để chuyển chế độ tối
        initial_dark_mode_text = "☀" # Sẽ được cập nhật ngay khi toggle_dark_mode được gọi
        self.dark_mode_btn = ttk.Button(header_frame, text=initial_dark_mode_text, command=self.toggle_dark_mode,
                                        width=3, style='DarkMode.TButton')
        self.dark_mode_btn.grid(row=0, column=2, sticky='e', padx=5)

        # SỬA: Frame chứa các nút điều khiển cửa sổ
        window_controls_frame = tk.Frame(header_frame)
        window_controls_frame.grid(row=0, column=3, sticky='e')
        
        self.minimize_btn = tk.Label(window_controls_frame, text=" _ ", font=(self.preferred_font_family, 10, "bold"))
        self.minimize_btn.pack(side="left")
        self.minimize_btn.bind("<Button-1>", lambda e: self.minimize_window())
        
        close_btn = tk.Label(window_controls_frame, text=" X ", font=(self.preferred_font_family, 10, "bold"))
        close_btn.pack(side="left", padx=(0, 5))
        close_btn.bind("<Button-1>", lambda e: self.on_app_close())

        # SỬA: Gán sự kiện kéo thả cho các phần nền của thanh tiêu đề, TRỪ các nút điều khiển
        for widget in [header_frame, window_controls_frame, logo_label, title_label]:
            widget.bind("<Button-1>", self._on_title_bar_press)
            widget.bind("<B1-Motion>", self._on_title_bar_drag)

        # Gán sự kiện hover cho các nút
        self.minimize_btn.bind("<Enter>", lambda e: e.widget.config(background="#6c757d"))
        self.minimize_btn.bind("<Leave>", lambda e: self._update_widget_colors(header_frame, *self.get_current_colors()))
        close_btn.bind("<Enter>", lambda e: e.widget.config(background="#dc3545"))
        close_btn.bind("<Leave>", lambda e: self._update_widget_colors(header_frame, *self.get_current_colors()))

        top_controls_frame = ttk.Frame(main_frame)
        top_controls_frame.grid(row=1, column=0, sticky="ew", pady=5)
        top_controls_frame.grid_columnconfigure(0, weight=1)
        top_controls_frame.grid_columnconfigure(1, weight=2)
        top_controls_frame.grid_columnconfigure(2, weight=1)
        self.top_controls_frame = top_controls_frame

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
        self.g_data_macro = g_data_macro

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

        self.btn_record = ttk.Button(g3_controls_record, text="● Record Macro (5s chuẩn bị)", command=self.record_macro,
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
        tk.Label(g3, text="Ghi: Insert->cột | Phím/Chuột->thao tác | ESC->kết thúc", font=(self.preferred_font_family, 9, "italic"), fg="gray").grid(row=4, column=0, sticky="w", padx=5, pady=(5, 5))
        # ====================================================================

        # RUN BUTTONS
        g5 = ttk.Frame(main_frame)
        g5.grid(row=3, column=0, sticky="ew", pady=(5, 0))
        
        self.g5 = g5

        tk.Label(g5, text="Chọn Chế độ Chạy:", font=("Courier", 9, "bold")).pack(side="left", padx=(0, 10))

        self.btn_test = ttk.Button(g5, text="▶️ CHẠY THỬ (1 DÒNG)", command=self.on_test, style='Accent.TButton')
        self.btn_test.pack(side="left", padx=10)

        self.btn_runall = ttk.Button(g5, text="▶️ CHẠY TẤT CẢ", command=self.on_run_all, style='Accent.TButton')
        self.btn_runall.pack(side="left", padx=10)

        self.btn_stop = ttk.Button(g5, text="STOP (ESC)", command=self.on_cancel, state='disabled')
        self.btn_stop.pack(side="left", padx=10)

        self.lbl_status = tk.Label(g5, text="Chờ...", fg="#1e90ff", font=(self.preferred_font_family, 10, "bold"))
        self.lbl_status.pack(side="left", padx=(20, 0))

        # ------------------------ REAL-TIME STATUS FRAME ------------------------
        self.realtime_status_frame = ttk.LabelFrame(main_frame, text="Thông tin Tọa độ (Real-time)", style="Dark.TLabelframe" if self.dark_mode_var.get() else "TLabelframe")
        self.realtime_status_frame.grid(row=4, column=0, sticky="ew", pady=(5, 0))
        self.realtime_status_frame.grid_columnconfigure(0, weight=1)

        status_controls = ttk.Frame(self.realtime_status_frame)
        status_controls.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        # --- SỬA LỖI LAYOUT: Dọn dẹp và cấu hình lại phần status real-time ---
        # Cấu hình để label co giãn, nút check giữ nguyên
        status_controls.grid_columnconfigure(0, weight=1)

        # Label hiển thị thông tin, đặt ở cột 0
        self.lbl_realtime_status = tk.Label(status_controls, text="...", justify="left", anchor="w")
        self.lbl_realtime_status.grid(row=0, column=0, sticky="ew")
        # Gán sự kiện để cập nhật wraplength khi kích thước label thay đổi
        self.lbl_realtime_status.bind('<Configure>',
                                      lambda e: self.lbl_realtime_status.config(wraplength=self.lbl_realtime_status.winfo_width() - 10))

        # Nút Checkbox, đặt ở cột 1
        ttk.Checkbutton(status_controls, text="Hiện/Ẩn", variable=self.show_realtime_status,
                        command=self._toggle_realtime_status).grid(row=0, column=1, sticky="e", padx=(10, 0))

        # Disclaimer (now row 5)
        tk.Label(main_frame,
                 text="Lưu ý: Ứng dụng BẮT BUỘC đưa ACSOFT lên foreground (phải focus). Nhấn phím ESC để hủy quá trình chạy.").grid(row=5, column=0, sticky="w",
                                                                                    pady=(5, 0))

        # SỬA: Áp dụng trạng thái ẩn/hiện ban đầu
        # --- KẾT THÚC SỬA ---
        self._toggle_realtime_status()

    def _apply_font_recursively(self, parent):
        """
        Áp dụng font chữ cho tất cả các widget con của một widget cha.
        """
        for child in parent.winfo_children():
            try:
                child.config(font=(self.preferred_font_family, 9))
            except tk.TclError:
                # Widget này không hỗ trợ cấu hình font
                pass

            # Gọi đệ quy cho các widget con của widget này
            self._apply_font_recursively(child)


    # -------------------------- UI Helpers (Theme, Status, etc.) --------------------------    
    def get_current_colors(self):
        """Trả về bộ màu (bg, fg, special_fg) cho theme hiện tại."""
        if self.dark_mode_var.get():
            return "#464646", "white", "#a9b7c6"
        else:
            return "#f0f0f0", "black", "gray"

    def toggle_dark_mode(self):
        """Chuyển đổi giữa theme sáng và tối."""
        # SỬA: Đảo ngược trạng thái khi nút được nhấn
        # Lấy trạng thái hiện tại và đảo ngược nó
        current_state = self.dark_mode_var.get()
        self.dark_mode_var.set(not current_state)

        is_dark = self.dark_mode_var.get()
        theme_name = "equilux" if is_dark else "arc"
        self.set_theme(theme_name)

        bg_color, fg_color, special_fg_color = self.get_current_colors()

        # --- SỬA: CẤU HÌNH LẠI STYLE CHO NÚT DARK MODE SAU KHI ĐỔI THEME ---
        # Việc này đảm bảo font tùy chỉnh không bị theme mới ghi đè.
        self.style.configure('DarkMode.TButton', font=('Courier', 12))
        self.style.configure('DarkMode.TButton', font=(self.preferred_font_family, 12))
        # SỬA: Cập nhật ký tự trên nút
        new_text = "◐" if is_dark else "☀"
        if self.dark_mode_btn:
            self.dark_mode_btn.config(text=new_text)

        # SỬA: Áp dụng màu nền cho cửa sổ chính
        self.config(background=bg_color)

        # Cập nhật màu cho các label có màu tùy chỉnh
        for widget in self.winfo_children():
            self._update_widget_colors(widget, bg_color, fg_color, special_fg_color)

        # SỬA: Gọi lại hàm áp dụng font đệ quy sau khi đổi theme
        self._apply_font_recursively(self)

    def _update_widget_colors(self, parent_widget, bg, fg, special_fg):
        """Đệ quy cập nhật màu cho các widget con."""
        for child in parent_widget.winfo_children():
            try:
                # SỬA: Áp dụng màu cho các widget tk thông thường
                widget_class = child.winfo_class()
                # Bỏ qua các nút ttk vì chúng đã có style riêng
                #if widget_class in ['Label', 'TFrame', 'Frame'] and not isinstance(child, ttk.Checkbutton):
                child.config(background=bg)
                # Chỉ đổi màu chữ cho Label, không đổi cho Frame
                if widget_class == 'Label':
                    # Nếu là label đặc biệt (status, italic) thì dùng màu chữ phụ
                    if child == self.lbl_realtime_status or "italic" in str(child.cget("font")):
                        child.config(fg=special_fg)
                    else:
                        child.config(fg=fg)

            except tk.TclError:
                # Bỏ qua các widget không có thuộc tính 'font' (ví dụ: Frame, Scrollbar)
                pass
            # Tiếp tục đệ quy cho các widget con
            self._update_widget_colors(child, bg, fg, special_fg)

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
                self.after(0, self.hud_window.update_status, "● ĐANG GHI... (Nhấn ESC để dừng)", "#FF4500")
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
        tree.tag_configure('highlight', background='#FFA07A', font=(self.preferred_font_family, 9, 'bold'))

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
            "dark_mode": self.dark_mode_var.get(), # SỬA: Lưu trạng thái dark mode
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

            self.show_realtime_status.set(settings.get("show_realtime_status", False))
            self._toggle_realtime_status()

            # SỬA: Tải và áp dụng trạng thái dark mode
            # Đặt trạng thái mong muốn, sau đó gọi toggle_dark_mode để nó "đảo ngược" lại đúng trạng thái đó
            # Ví dụ: muốn dark mode (True), ta set var thành False rồi gọi toggle, nó sẽ đảo thành True
            should_be_dark = settings.get("dark_mode", True)
            if self.dark_mode_var.get() != should_be_dark:
                self.dark_mode_var.set(not should_be_dark)
                self.toggle_dark_mode()

            csv_path = settings.get("csv_path")
            if csv_path and os.path.exists(csv_path):
                self.load_csv_data(csv_path)
        except Exception as e:
            print(f"Lỗi khi áp dụng settings: {e}")


if __name__ == "__main__":
    app = MacroApp()
    app.mainloop()