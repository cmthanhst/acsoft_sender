import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit, QFileDialog, QMessageBox, QFrame, QVBoxLayout, QHBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QSpinBox, QComboBox, QDialog, QDialogButtonBox, QGroupBox, QRadioButton, QScrollBar, QStyleFactory
from PySide6.QtGui import QIcon, QPixmap, QImage, QFont, QColor, QCursor, QKeyEvent, QMouseEvent, QPainter, QRegion, QBrush
from PySide6.QtCore import Qt, QTimer, QEvent, QRect, Signal, QIODevice, QBuffer, QThread, QObject
import os
import time
import threading
try:
    import pandas as pd
except ImportError:
    QMessageBox.critical(None, "Lỗi Cài Đặt", "Vui lòng chạy 'pip install pandas' để cài\nđặt thư viện xử lý CSV.")
    sys.exit()

import subprocess
import ctypes
import json
from decimal import Decimal
from enum import Enum, auto

try:
    from PIL import Image, ImageDraw
    import io
    import base64
except ImportError:
    QMessageBox.critical(None, "Lỗi Cài Đặt", "Vui lòng chạy 'pip install Pillow' để cài\nđặt thư viện xử lý ảnh.")
    sys.exit()

try:
    from pynput import mouse, keyboard
    from pynput.keyboard import Key
except ImportError:
    QMessageBox.critical(None, "Lỗi Cài Đặt", "Vui lòng chạy 'pip install pynput' để cài\nđặt thư viện ghi chuột/phím.")
    sys.exit()

try:
    import win32gui
    import win32con
    import win32api
    import win32process
    from win32process import GetWindowThreadProcessId
except ImportError:
    win32gui = win32con = win32api = win32process = None
    QMessageBox.critical(None, "Lỗi Cài Đặt", "Vui lòng chạy 'pip install pywin32' để sử\ndụng chức năng cửa sổ Windows.")
    sys.exit()

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
        monitor_handle = win32api.MonitorFromWindow(hwnd, 
win32con.MONITOR_DEFAULTTONEAREST)

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
    """Lấy (left, top, right, bottom) tọa độ tuyệt đối của cửa sổ (LOGICAL 
PIXEL)."""
    try:
        rect = win32gui.GetWindowRect(hwnd)
        return rect
    except Exception:
        return None

def bring_to_front(hwnd):
    """
    Đưa cửa sổ lên foreground.
    Chỉ gọi SW_RESTORE nếu cửa sổ đang bị Minimize, giữ nguyên trạng thái 
Maximize.
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

def send_mouse_click(hwnd, x_offset_logical, y_offset_logical, button_type, 
recorded_scale):
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
            col_display = self.col_index + 1 if self.col_index is not None else 
"N/A"
            return f"[COL] Gửi giá trị cột {col_display} {delay_str}"
        elif self.typ == MacroStepType.MOUSE_CLICK.value:
            # HIỂN THỊ DƯỚNG DẠNG PIXEL OFFSET CHUẨN (100% SCALE)

            scale = self.dpi_scale if self.dpi_scale > 0 else 1.0

            # Tính toán Offset Chuẩn hóa (Normalized Offset)
            x_norm = int(self.x_offset_logical / scale) if self.x_offset_logical 
is not None else "N/A"
            y_norm = int(self.y_offset_logical / scale) if self.y_offset_logical 
is not None else "N/A"

            click_type = self.key_value.replace("_click", "").capitalize()
            return f"[MOUSE] {click_type} Click tại Offset Chuẩn ({x_norm}px, 
{y_norm}px) (Scale Ghi: {int(self.dpi_scale * 100)}%) {delay_str}"
        elif self.typ == MacroStepType.KEY_PRESS.value:
            return f"[KEY] Gửi phím: '{self.key_value.upper()}' {delay_str}"
        elif self.typ == MacroStepType.KEY_COMBO.value:
            return f"[COMBO] Gửi tổ hợp phím: '{self.key_value.upper()}' 
{delay_str}"
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

class RecordingHUD(QDialog):
    PAUSED_COLOR = QColor("#FFD700")  # Vàng
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
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFixedSize(300, 50)  # Adjust size as needed

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)

        self.status_label = QLabel("Chuẩn bị...")
        self.status_label.setFont(QFont("Courier", 10))
        layout.addWidget(self.status_label)

        self.pause_button = QPushButton(" PAUSE")
        self.pause_button.clicked.connect(self.toggle_pause)
        layout.addWidget(self.pause_button)
        self.pause_button.hide()  # Ẩn nút pause ban đầu

        stop_button = QPushButton("■ STOP")
        stop_button.clicked.connect(self.stop_callback)
        layout.addWidget(stop_button)

        # Gán sự kiện để di chuyển HUD
        self.setMouseTracking(True)

        # Căn giữa HUD ở cạnh trên màn hình
        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width() // 2) - (self.width() // 2), 20)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._offset_x = event.x()
            self._offset_y = event.y()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            x = self.mapToGlobal(event.pos()).x() - self._offset_x
            y = self.mapToGlobal(event.pos()).y() - self._offset_y
            self.move(x, y)

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_event.clear()  # Chặn thread
            self.pause_button.setText("  RESUME")
            self.update_status("  TẠM DỪNG", self.PAUSED_COLOR.name())
        else:
            self.pause_event.set()  # Cho phép thread chạy tiếp
            self.pause_button.setText(" PAUSE")
            # Trạng thái sẽ được cập nhật lại bởi vòng lặp chính

    def update_status(self, text, color="white"):
        """Cập nhật văn bản và màu sắc của label trạng thái."""
        if self.is_paused:  # Nếu đang pause thì không cập nhật status từ bên ngoài
            return
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"color: {color};")

    def closeEvent(self, event):
        self.destroy()

# =========================================================================
# ------------------------------ Qt App ---------------------
# =========================================================================

LOGO_PNG_BASE64 = \
b"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