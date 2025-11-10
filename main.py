import sys
import os
import base64
import json
import threading # SỬA: Thêm threading
import subprocess # SỬA: Thêm thư viện subprocess để chạy file .exe
import time
import ctypes # SỬA: Thêm thư viện ctypes để tương tác với Windows API
from PySide6.QtCore import Qt, QPoint, QSize, QTimer, QThread, Signal, QObject
from PySide6.QtWidgets import (    
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit,
    QFileDialog, QMessageBox, QFrame, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QSpinBox,    
    QComboBox, QDialog, QGroupBox, QRadioButton, QButtonGroup, QDialogButtonBox, QAbstractItemView
)
from PySide6.QtGui import QIcon, QPixmap, QFont, QColor, QCursor, QPainter, QBrush, QRegion

# SỬA: Thêm các thư viện cần thiết cho việc lấy danh sách cửa sổ
try:
    import win32gui
    import win32con
    import win32api
    import win32process
    from win32process import GetWindowThreadProcessId

    # SỬA: Thêm pynput và xử lý lỗi
    from pynput import mouse, keyboard
    from pynput.keyboard import Key

except ImportError:
    app = QApplication(sys.argv)
    QMessageBox.critical(None, "Lỗi Cài Đặt", "Vui lòng chạy 'pip install pywin32 pynput' để sử dụng đầy đủ chức năng.")
    sys.exit()

# SỬA: Thêm import pandas và xử lý lỗi nếu chưa cài đặt
try:
    import pandas as pd
except ImportError:
    # Dùng một QApplication tạm thời để có thể hiển thị QMessageBox
    app = QApplication(sys.argv)
    QMessageBox.critical(None, "Lỗi Cài Đặt", "Vui lòng chạy 'pip install pandas' để cài đặt thư viện xử lý CSV.")
    sys.exit()

# Đọc nội dung Base64 từ file
try:
    with open('logo_base64.txt', 'r') as f:
        LOGO_PNG_BASE64 = f.read().strip()
except FileNotFoundError:
    print("Lỗi: Không tìm thấy file 'logo_base64.txt'. Sẽ sử dụng ảnh placeholder.")
    LOGO_PNG_BASE64 = "" # Để trống nếu không tìm thấy

# =========================================================================
# SỬA: DI CHUYỂN CÁC HÀM TIỆN ÍCH RA NGOÀI CLASS (GLOBAL SCOPE)
# =========================================================================

# ----------------------------- WinAPI Helpers ----------------------------
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

# ----------------------------- Mouse/Keyboard Sender ---------------------
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
# -------------------- HUD Window (PySide6) -------------------------------
# =========================================================================
class RecordingHUD(QDialog):
    """
    Giao diện cho cửa sổ HUD (Heads-Up Display) được xây dựng bằng PySide6.
    Cửa sổ này không viền, trong suốt và có thể kéo thả.
    """
    def __init__(self, parent=None):
        super().__init__(parent)

        # SỬA: Thêm các thuộc tính cho chức năng Pause/Resume
        self.pause_event = None # Sẽ là một threading.Event() được truyền từ bên ngoài
        self.is_paused = False

        # Thiết lập cửa sổ không viền, luôn ở trên và trong suốt
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool) # Qt.Tool để không hiện trên taskbar
        self.setAttribute(Qt.WA_TranslucentBackground, True) # Nền trong suốt

        # Frame chính với nền tối và bo góc (style được định nghĩa trong dark.qss)
        self.main_frame = QFrame(self) # Tên object "mainFrame" để stylesheet có thể target
        self.main_frame.setObjectName("mainFrame") 

        # Layout cho các thành phần trong HUD
        layout = QHBoxLayout(self.main_frame)
        layout.setContentsMargins(10, 5, 10, 5)

        self.status_label = QLabel("Chuẩn bị...") # Font được định nghĩa trong dark.qss
        self.status_label.setObjectName("hudStatusLabel") # Để có thể target trong QSS nếu cần
        layout.addWidget(self.status_label) # Style cho QLabel được định nghĩa trong dark.qss

        # SỬA: Thêm ghi chú cho phím F5
        self.f5_label = QLabel("(F5)")
        self.f5_label.setStyleSheet("color: #aaa; font-size: 8pt;") # Màu xám nhạt, chữ nhỏ
        self.f5_label.hide() # Ẩn ban đầu

        self.pause_button = QPushButton("❚❚ PAUSE") # Font được định nghĩa trong dark.qss
        self.pause_button.setObjectName("hudPauseButton")
        self.pause_button.hide() # Ẩn ban đầu
        layout.addWidget(self.pause_button)

        self.stop_button = QPushButton("■ STOP") # Font được định nghĩa trong dark.qss
        self.stop_button.setObjectName("hudStopButton")
        layout.addWidget(self.f5_label) # SỬA: Thêm label F5 vào layout
        layout.addWidget(self.stop_button)

        # Layout chính của Dialog
        dialog_layout = QVBoxLayout(self)
        dialog_layout.setContentsMargins(0, 0, 0, 0)
        dialog_layout.addWidget(self.main_frame)

        # Biến để kéo thả
        self._drag_pos = None

        # Căn giữa màn hình
        self.adjustSize()
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        self.move(x, 20)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self._drag_pos:
            diff = event.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + diff)
            self._drag_pos = event.globalPosition().toPoint()

    # SỬA: Thêm hàm để xử lý việc tạm dừng/tiếp tục
    def toggle_pause(self):
        """Xử lý sự kiện khi nút Pause/Resume được nhấn."""
        if not self.pause_event:
            return

        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_event.clear()  # Chặn luồng worker
            self.pause_button.setText("▶ RESUME")
            self.update_status("❚❚ TẠM DỪNG", "#FFD700") # Màu vàng
        else:
            self.pause_event.set()  # Cho phép luồng worker chạy tiếp
            self.pause_button.setText("❚❚ PAUSE")
            # Trạng thái sẽ được cập nhật lại bởi worker, không cần set ở đây

    def update_status(self, text, color="white"):
        """Cập nhật văn bản và màu sắc của label trạng thái."""
        # SỬA: Nếu đang tạm dừng, không cho worker cập nhật đè lên chữ "TẠM DỪNG"
        if self.is_paused:
            # Chỉ cho phép cập nhật nếu đó là thông điệp TẠM DỪNG
            if "TẠM DỪNG" in text:
                self.status_label.setText(text)
                self.status_label.setStyleSheet(f"color: {color};")
            return
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"color: {color};")

# =========================================================================
# SỬA: THÊM ĐỊNH NGHĨA MACROSTEP VÀO TRƯỚC KHI SỬ DỤNG
# =========================================================================
from enum import Enum

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
            scale = self.dpi_scale if self.dpi_scale > 0 else 1.0
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
        return {k: v for k, v in self.__dict__.items() if k != 'item_id'}

    @staticmethod
    def from_dict(data):
        return MacroStep(
            typ=data["typ"],
            key_value=data.get('key_value'),
            col_index=data.get('col_index'),
            delay_after=data.get('delay_after', 0.01),
            x_offset=data.get("x_offset_logical", data.get("x_offset")),
            y_offset=data.get("y_offset_logical", data.get("y_offset")),
            dpi_scale=data.get('dpi_scale', 1.0)
        )
# =========================================================================
# -------------------- MacroStep Edit Dialog ------------------------------
# =========================================================================
from PySide6.QtGui import QDoubleValidator # Import QDoubleValidator

class MacroStepEditDialog(QDialog):
    def __init__(self, parent, step: MacroStep):
        super().__init__(parent)
        self.setWindowTitle(f"Sửa Bước Macro #{step.item_idx + 1}")
        self.setModal(True) # Make it a modal dialog
        self.step = step # Reference to the original MacroStep object

        self.setup_ui()
        self.load_step_data()

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        form_layout = QGridLayout()

        # Type
        form_layout.addWidget(QLabel("Loại Bước:"), 0, 0)
        self.type_combo = QComboBox()
        self.type_combo.addItems(['col', 'key', 'combo', 'mouse', 'end'])
        form_layout.addWidget(self.type_combo, 0, 1)

        # Key/Click Value
        form_layout.addWidget(QLabel("Giá trị (Key/Click Type):"), 1, 0)
        self.key_value_edit = QLineEdit()
        form_layout.addWidget(self.key_value_edit, 1, 1)

        # Column Index
        form_layout.addWidget(QLabel("Chỉ số Cột (0, 1, 2...):"), 2, 0)
        self.col_index_spin = QSpinBox()
        self.col_index_spin.setRange(0, 999) # Assuming max 1000 columns
        form_layout.addWidget(self.col_index_spin, 2, 1)

        # X Offset (Logical Pixel)
        form_layout.addWidget(QLabel("Offset X (Logical Pixel):"), 3, 0)
        self.x_offset_edit = QLineEdit()
        self.x_offset_edit.setValidator(QDoubleValidator())
        form_layout.addWidget(self.x_offset_edit, 3, 1)

        # Y Offset (Logical Pixel)
        form_layout.addWidget(QLabel("Offset Y (Logical Pixel):"), 4, 0)
        self.y_offset_edit = QLineEdit()
        self.y_offset_edit.setValidator(QDoubleValidator())
        form_layout.addWidget(self.y_offset_edit, 4, 1)

        # DPI Scale (Read-only)
        form_layout.addWidget(QLabel("DPI Scale Ghi (%):"), 5, 0)
        self.dpi_scale_label = QLabel()
        form_layout.addWidget(self.dpi_scale_label, 5, 1)

        # Delay After (ms)
        form_layout.addWidget(QLabel("Độ trễ sau (ms - 10-10000):"), 6, 0)
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(10, 10000)
        self.delay_spin.setSingleStep(10)
        form_layout.addWidget(self.delay_spin, 6, 1)

        main_layout.addLayout(form_layout)

        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)

        # Connect type_combo to update visibility of other fields
        self.type_combo.currentTextChanged.connect(self.update_field_visibility)

    def load_step_data(self):
        self.type_combo.setCurrentText(self.step.typ)
        self.key_value_edit.setText(str(self.step.key_value) if self.step.key_value is not None else "")
        self.col_index_spin.setValue(self.step.col_index if self.step.col_index is not None else 0)
        self.x_offset_edit.setText(f"{self.step.x_offset_logical:.2f}" if self.step.x_offset_logical is not None else "")
        self.y_offset_edit.setText(f"{self.step.y_offset_logical:.2f}" if self.step.y_offset_logical is not None else "")
        self.dpi_scale_label.setText(f"{int(self.step.dpi_scale * 100)}%")
        self.delay_spin.setValue(int(self.step.delay_after * 1000))
        self.update_field_visibility(self.step.typ)

    def update_field_visibility(self, step_type):
        # Hide/show fields based on step_type
        is_col = (step_type == 'col')
        is_key_combo = (step_type == 'key' or step_type == 'combo')
        is_mouse = (step_type == 'mouse')

        self.key_value_edit.setVisible(is_key_combo or is_mouse)
        self.col_index_spin.setVisible(is_col)
        self.x_offset_edit.setVisible(is_mouse)
        self.y_offset_edit.setVisible(is_mouse)
        # Hide/show the label and its value for DPI scale
        self.dpi_scale_label.parentWidget().setVisible(is_mouse)
        self.dpi_scale_label.setVisible(is_mouse)

    def get_edited_data(self):
        # This method will be called if dialog is accepted
        pass
        # Validation is handled here
        # ... (logic for getting edited data and validation, see below)

# =========================================================================
# -------------------- Recording Worker (Thread) --------------------------
# =========================================================================
class RecordingWorker(QObject):
    """
    Worker chạy trong một thread riêng để lắng nghe sự kiện chuột và phím
    mà không làm đóng băng giao diện chính.
    """
    update_hud_signal = Signal(str, str)
    add_step_signal = Signal(object)
    recording_finished_signal = Signal(bool) # Gửi True nếu hoàn thành, False nếu bị hủy

    def __init__(self, target_window_title, parent_app):
        super().__init__()
        self.target_window_title = target_window_title
        self.app = parent_app # Tham chiếu đến ứng dụng chính
        self.cancel_flag = threading.Event()

    def run(self):
        """Hàm chính của worker, bao gồm đếm ngược và bắt đầu lắng nghe."""
        self.app.stop_listeners() # Dừng listener cũ nếu có
        self.app.current_modifiers.clear()

        # Listener tạm thời cho phím ESC trong lúc đếm ngược
        countdown_escape_listener = keyboard.Listener(on_press=self._on_escape_press)
        countdown_escape_listener.start()

        try:
            # Đếm ngược 5 giây
            for i in range(5, 0, -1):
                if self.cancel_flag.is_set():
                    return # Chỉ cần return, finally sẽ xử lý việc emit tín hiệu
                self.update_hud_signal.emit(f"Bắt đầu ghi sau: {i}s", "#87CEEB")
                time.sleep(1) # Dừng 1 giây

            if self.cancel_flag.is_set():
                return

            # Bắt đầu ghi
            hwnd = hwnd_from_title(self.target_window_title)
            if hwnd:
                bring_to_front(hwnd)

            self.update_hud_signal.emit("● ĐANG GHI... (Nhấn ESC để dừng)", "#FF4500")
            self.app.last_key_time = time.time() # Bắt đầu tính thời gian từ đây
            self._start_main_listeners()

            # Giữ worker sống trong khi listener đang chạy
            self.cancel_flag.wait() # Chờ cho đến khi cờ cancel được set

        finally:
            countdown_escape_listener.stop()
            self.app.stop_listeners()
            # Gửi tín hiệu hoàn thành (ngay cả khi bị hủy) để dọn dẹp giao diện
            self.recording_finished_signal.emit(not self.cancel_flag.is_set())

    def _on_escape_press(self, key):
        """Chỉ lắng nghe phím ESC."""
        if key == keyboard.Key.esc:
            self.cancel_run()
            self.stop()

    def _start_main_listeners(self):
        """Khởi động các listener chính cho việc ghi macro."""
        # Lắng nghe sự kiện chuột và phím từ pynput
        # Các hàm _on_mouse_click, _on_key_press, _on_key_release nằm trong class MacroApp
        # để dễ dàng truy cập các thuộc tính của app.
        self.app.mouse_listener = mouse.Listener(on_click=self.app._on_mouse_click)
        self.app.mouse_listener.start()

        self.app.keyboard_listener = keyboard.Listener(
            on_press=self.app._on_key_press,
            on_release=self.app._on_key_release
        )
        self.app.keyboard_listener.start()

    def stop(self):
        """Dừng worker và các listener."""
        # Ghi bước END cuối cùng trước khi dừng
        current_time = time.time()
        delay = current_time - self.app.last_key_time if self.app.last_key_time != 0.0 else 0.0
        end_step = MacroStep('end', delay_after=delay)
        self.add_step_signal.emit(end_step)

        # Đặt cờ để dừng vòng lặp chờ trong hàm run()
        self.cancel_flag.set()


# =========================================================================
# -------------------- Main Application Window (PySide6) ------------------
# =========================================================================
from PySide6.QtCore import Signal # SỬA: Import Signal

class MacroApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Thiết lập cửa sổ không viền
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowTitle("Việt Tín Auto Sender V2025.04")
        # SỬA: Xóa kích thước cố định, chỉ đặt chiều rộng và để chiều cao tự động
        # self.resize(1200, 800)
        # Load logo và set icon cho taskbar
        try:
            with open('logo_base64.txt', 'r') as f:
                logo_base64 = f.read().strip()
            pixmap = QPixmap()
            pixmap.loadFromData(base64.b64decode(logo_base64))
            self.setWindowIcon(QIcon(pixmap))
        except FileNotFoundError:
            print("Lỗi: Không tìm thấy file 'logo_base64.txt'. Không set icon.")
        self.setFixedWidth(1200)

        # Frame chính với bo góc
        self.main_frame = QFrame()
        self.main_frame.setObjectName("mainFrame")
        self.setCentralWidget(self.main_frame)

        # Layout chính
        self.main_layout = QVBoxLayout(self.main_frame)
        self.main_layout.setContentsMargins(1, 1, 1, 1) # Viền mỏng để thấy bo góc
        self.main_layout.setSpacing(0)

        # Biến cho việc kéo thả và thu gọn
        self._drag_pos = None
        self.is_collapsed = False
        self.normal_geometry = self.geometry()

        self.is_dark_mode = False # SỬA: Đặt trạng thái ban đầu là False (sáng)
        # SỬA: Thêm các thuộc tính cần thiết cho Group 1
        self.txt_acpath = None
        self.acsoft_path = None

        # SỬA: Thêm các thuộc tính cần thiết cho Group 2
        self.txt_csv = None
        self.txt_delimiter = None
        self.combo_windows = None
        self.target_window_title = ""

        # SỬA: Thêm các thuộc tính cho việc ghi macro
        self.macro_steps = []
        self.recording = False
        self.current_col_index = 0
        self.last_key_time = 0.0
        self.mouse_listener = None
        self.keyboard_listener = None
        self.current_modifiers = set()
        self.hud_window = None
        self.recording_thread = None
        self.recording_worker = None
        self._recording_finished_processed = False # New flag to prevent double processing
        # SỬA: Thêm thuộc tính cho việc chạy macro
        self.run_thread = None
        self.run_worker = None
        self.pause_run_flag = threading.Event() # SỬA: Thêm cờ cho việc Pause/Resume
        self.cancel_run_flag = threading.Event()
        # SỬA: Thêm thuộc tính cho việc tô sáng
        self.highlight_color = QColor("#FFA07A22") # Màu cam nhạt để tô sáng
        self.default_bg_color = QColor("#3c4049") # Màu nền mặc định của bảng (dark mode)
        self.last_highlighted_csv_row = -1
        self.last_highlighted_macro_row = -1
        # SỬA: Thêm thuộc tính cho các nút chạy
        self.record_btn = None # Thêm thuộc tính cho nút record
        self.run_test_btn = None
        self.run_all_btn = None
        self.stop_btn = None
        self.ontop_btn = None # SỬA: Thêm thuộc tính cho nút Always on Top
        self.edit_step_btn = None # SỬA: Thêm thuộc tính cho nút Sửa Dòng
        # SỬA: Thêm các thuộc tính cho status bar
        self.lbl_realtime_status = None
        self.chk_show_realtime_status = None
        self.realtime_status_timer = QTimer(self)
        self.realtime_status_timer.timeout.connect(self._update_status_bar_info)


        self.df_csv = pd.DataFrame() # Khởi tạo dataframe rỗng

        # --- Tạo các thành phần giao diện ---
        self._create_header_bar() 
        self._create_top_controls()
        self._create_data_macro_section()
        self._create_run_buttons()
        self._create_realtime_status_bar()
        self._create_disclaimer()

        # Áp dụng theme tối mặc định
        self.toggle_dark_mode() # SỬA: Gọi hàm để nó tự chuyển sang chế độ tối

        # SỬA: Tải danh sách cửa sổ lần đầu khi khởi động
        self.refresh_windows()

        # SỬA: Khởi động listener ESC toàn cục
        self._start_global_hotkey_listener()

        # SỬA: Yêu cầu cửa sổ tự điều chỉnh kích thước sau khi đã tạo xong mọi thứ
        self.adjustSize()

    # SỬA: Xóa bỏ paintEvent và resizeEvent vì việc bo góc giờ đã được xử lý
    # hoàn toàn bằng stylesheet trong file dark.qss, giúp code sạch hơn và
    # hiệu quả hơn.

    def _create_header_bar(self):
        """Tạo thanh tiêu đề tùy chỉnh."""
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(10, 5, 5, 5)

        # Logo
        logo_label = QLabel()
        if LOGO_PNG_BASE64:
            pixmap = QPixmap()
            pixmap.loadFromData(base64.b64decode(LOGO_PNG_BASE64))
            logo_label.setPixmap(pixmap.scaled(36, 36, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        header_layout.addWidget(logo_label)

        # Tiêu đề
        title_label = QLabel("Việt Tín Auto Sender V2025.04") # Font được định nghĩa trong dark.qss
        title_label.setObjectName("titleLabel") # Gán objectName để QSS có thể target
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Nút Dark Mode
        self.dark_mode_btn = QPushButton("◐")
        self.dark_mode_btn.setFixedSize(30, 30)
        self.dark_mode_btn.clicked.connect(self.toggle_dark_mode) # SỬA: Kết nối sự kiện click
        header_layout.addWidget(self.dark_mode_btn)

        # SỬA: Thêm nút Always on Top
        self.ontop_btn = QPushButton("◰")
        self.ontop_btn.setFixedSize(30, 30)
        self.ontop_btn.setCheckable(True) # Cho phép nút có trạng thái bật/tắt
        self.ontop_btn.clicked.connect(self._toggle_always_on_top)
        header_layout.addWidget(self.ontop_btn)

        # Nút Minimize
        self.minimize_btn = QPushButton("_")
        self.minimize_btn.setFixedSize(30, 30)
        self.minimize_btn.clicked.connect(self.showMinimized) # SỬA: Kết nối nút minimize để thu nhỏ cửa sổ
        header_layout.addWidget(self.minimize_btn)

        # Nút Close
        close_btn = QPushButton("X")
        close_btn.setFixedSize(30, 30)
        close_btn.setObjectName("closeButton")
        close_btn.clicked.connect(self.close)
        header_layout.addWidget(close_btn)

        self.main_layout.addWidget(header_widget)

        # Gán sự kiện kéo thả
        header_widget.mousePressEvent = self._on_title_bar_press
        header_widget.mouseMoveEvent = self._on_title_bar_drag

    def _create_top_controls(self):
        """Tạo 3 group box điều khiển ở trên cùng."""
        top_controls_widget = QWidget()
        top_controls_layout = QHBoxLayout(top_controls_widget)
        top_controls_layout.setContentsMargins(10, 10, 10, 10)

        # Group 1: ACSOFT Path
        g1 = QGroupBox("1) Đường dẫn file chạy ACSOFT")
        # SỬA: Chuyển sang QVBoxLayout để có 2 dòng
        g1_layout = QVBoxLayout(g1)

        # Dòng 1: LineEdit
        self.txt_acpath = QLineEdit() # SỬA: Gán LineEdit vào thuộc tính self.txt_acpath
        g1_layout.addWidget(self.txt_acpath)

        # Dòng 2: Các nút
        button_container = QWidget() # Tạo container cho các nút
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(0, 0, 0, 0) # Xóa margin thừa
        button_layout.addStretch() # Đẩy các nút về bên phải
        button_layout.addWidget(QPushButton("Browse", clicked=self.browse_ac)) # SỬA: Kết nối nút Browse
        button_layout.addWidget(QPushButton("Chạy ACSOFT", clicked=self.open_ac)) # SỬA: Kết nối nút Mở ACSOFT
        g1_layout.addWidget(button_container)

        # Group 2: CSV and Target Window
        g2 = QGroupBox("2) File CSV chứa dữ liệu / Cửa sổ mục tiêu")
        g2_layout = QVBoxLayout(g2)
        
        csv_layout = QHBoxLayout()
        # SỬA: Thêm LineEdit và nút Browse CSV
        csv_layout.addWidget(QLabel("File CSV:"))
        self.txt_csv = QLineEdit()
        csv_layout.addWidget(self.txt_csv)
        csv_layout.addWidget(QPushButton("Browse CSV", clicked=self.browse_csv))
        g2_layout.addLayout(csv_layout)

        window_layout = QHBoxLayout()
        window_layout.addWidget(QLabel("Delimiter:"))
        # SỬA: Thêm LineEdit và ComboBox
        # SỬA: Đặt chiều rộng cố định cho LineEdit của Delimiter
        delimiter_edit = QLineEdit(";") # Tạo LineEdit
        self.txt_delimiter = delimiter_edit # Gán cho self.txt_delimiter

        delimiter_edit.setFixedWidth(40)
        window_layout.addWidget(delimiter_edit)
        window_layout.addWidget(QLabel("Cửa sổ:"))
        # SỬA: Thêm stretch factor để ComboBox lấp đầy không gian còn trống
        self.combo_windows = QComboBox()
        self.combo_windows.currentTextChanged.connect(self.on_window_select) # Kết nối sự kiện chọn
        window_layout.addWidget(self.combo_windows, 1) # Stretch factor = 1
        window_layout.addWidget(QPushButton("Làm mới", clicked=self.refresh_windows)) # Kết nối nút
        g2_layout.addLayout(window_layout)

        # Group 3: Run Options
        g4 = QGroupBox("3) Tùy chọn chạy")
        g4_layout = QVBoxLayout(g4)
        
        speed_layout = QHBoxLayout()
        
        # SỬA: Tạo QButtonGroup để quản lý các RadioButton
        self.speed_mode_group = QButtonGroup(self)

        self.radio_recorded_speed = QRadioButton("Tốc độ đã ghi")
        self.radio_fixed_speed = QRadioButton("Tốc độ cố định:")
        
        # SỬA: Đặt mặc định là "Tốc độ cố định"
        self.radio_fixed_speed.setChecked(True)

        self.speed_mode_group.addButton(self.radio_recorded_speed, 1) # ID 1 cho tốc độ đã ghi
        self.speed_mode_group.addButton(self.radio_fixed_speed, 2)    # ID 2 cho tốc độ cố định

        speed_layout.addWidget(self.radio_recorded_speed)
        speed_layout.addWidget(self.radio_fixed_speed)
        
        # SỬA: Cấu hình QSpinBox cho tốc độ cố định
        self.spin_fixed_speed = QSpinBox()
        self.spin_fixed_speed.setRange(100, 10000) # Giới hạn từ 100-10000
        self.spin_fixed_speed.setValue(1000)      # Mặc định là 1000
        speed_layout.addWidget(self.spin_fixed_speed)
        speed_layout.addWidget(QLabel("ms"))
        speed_layout.addStretch()
        g4_layout.addLayout(speed_layout)

        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("Đợi giữa 2 dòng (1-20 giây):"))
        # SỬA: Cấu hình QSpinBox cho độ trễ giữa các dòng
        self.spin_delay_between_rows = QSpinBox()
        self.spin_delay_between_rows.setRange(1, 20) # Giới hạn từ 1-20
        self.spin_delay_between_rows.setValue(2)    # Mặc định là 2
        delay_layout.addWidget(self.spin_delay_between_rows, 1) # Thêm stretch factor = 1
        g4_layout.addLayout(delay_layout)

        top_controls_layout.addWidget(g1, 1)
        top_controls_layout.addWidget(g2, 2)
        top_controls_layout.addWidget(g4, 1)

        self.main_layout.addWidget(top_controls_widget)

    def _create_data_macro_section(self):
        """Tạo khu vực hiển thị dữ liệu CSV và các bước Macro."""
        data_macro_widget = QWidget()
        data_macro_layout = QHBoxLayout(data_macro_widget)
        data_macro_layout.setContentsMargins(10, 0, 10, 10)

        # CSV Data Table
        csv_group = QGroupBox("Dữ liệu CSV (Chỉ hiển thị 10 dòng đầu)")
        csv_layout = QVBoxLayout(csv_group)
        self.tree_csv = QTableWidget(10, 20) # 10 dòng, 20 cột
        self.tree_csv.setHorizontalHeaderLabels([f"Cột {i+1}" for i in range(20)])
        self.tree_csv.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        csv_layout.addWidget(self.tree_csv)

        # Macro Steps
        macro_group = QGroupBox("4) Ghi Macro & Điều chỉnh")
        macro_layout = QVBoxLayout(macro_group)

        # -- Macro Controls --
        record_controls = QHBoxLayout()
        self.record_btn = QPushButton("● Record Macro (5s chuẩn bị)")
        self.record_btn.clicked.connect(self.record_macro) # SỬA: Kết nối nút Record
        self.record_btn.setObjectName("recordButton") # Gán ID cho nút Record
        record_controls.addWidget(self.record_btn)
        record_controls.addStretch() # Giữ nguyên stretch để đẩy các nút còn lại sang phải
        record_controls.addWidget(QPushButton("Lưu Macro", clicked=self.save_macro)) # SỬA: Kết nối nút Lưu Macro
        record_controls.addWidget(QPushButton("Mở Macro", clicked=self.load_macro))
        record_controls.addWidget(QPushButton("Clear Macro", clicked=self.clear_macro)) # SỬA: Kết nối nút Clear Macro
        macro_layout.addLayout(record_controls)        

        # -- Macro Table --
        self.tree_macro = QTableWidget()
        self.tree_macro.setColumnCount(3)
        self.tree_macro.setHorizontalHeaderLabels(["STT", "Loại", "Mô tả chi tiết"])
        self.tree_macro.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree_macro.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tree_macro.setSelectionBehavior(QAbstractItemView.SelectRows) # Đảm bảo chọn toàn bộ hàng
        self.tree_macro.itemSelectionChanged.connect(self._update_edit_button_text) # SỬA: Kết nối sự kiện chọn dòng
        self.tree_macro.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        macro_layout.addWidget(self.tree_macro)


        # -- Manual Add Buttons --
        add_controls = QHBoxLayout()
        add_controls.addWidget(QPushButton("[+] Cột Dữ Liệu", clicked=lambda: self.add_manual_step("col")))
        add_controls.addWidget(QPushButton("[+] Phím (Key)", clicked=lambda: self.add_manual_step("key")))
        add_controls.addWidget(QPushButton("[+] Tổ Hợp (Combo)", clicked=lambda: self.add_manual_step("combo")))
        add_controls.addWidget(QPushButton("[+] Click Chuột", clicked=lambda: self.add_manual_step("mouse")))
        add_controls.addStretch()
        macro_layout.addLayout(add_controls)


        # -- Edit/Delete Buttons --
        edit_controls = QHBoxLayout()
        self.edit_step_btn = QPushButton("Sửa Dòng", clicked=self.edit_macro_step) # SỬA: Gán nút vào thuộc tính self
        edit_controls.addWidget(self.edit_step_btn)
        edit_controls.addWidget(QPushButton("Xóa Dòng", clicked=self.delete_macro_step))
        edit_controls.addWidget(QPushButton("[+] Kết Thúc Dòng", clicked=lambda: self.add_manual_step("end")))
        edit_controls.addStretch()
        macro_layout.addLayout(edit_controls)


        # -- Note --
        macro_layout.addWidget(QLabel("Ghi: Insert->cột | Phím/Chuột->thao tác | ESC->kết thúc"))
        data_macro_layout.addWidget(csv_group, 1)
        data_macro_layout.addWidget(macro_group, 1)

        self.main_layout.addWidget(data_macro_widget)

    def _create_run_buttons(self):
        """Tạo các nút chạy ở dưới cùng."""
        run_widget = QWidget()
        run_layout = QHBoxLayout(run_widget)
        run_layout.setContentsMargins(10, 0, 10, 10)

        run_layout.addWidget(QLabel("Chọn Chế độ Chạy:"))
        self.run_test_btn = QPushButton("► CHẠY THỬ (1 DÒNG)", clicked=self.on_test)
        self.run_test_btn.setObjectName("runTestButton")
        run_layout.addWidget(self.run_test_btn)
        self.run_all_btn = QPushButton("► CHẠY TẤT CẢ", clicked=self.on_run_all)
        self.run_all_btn.setObjectName("runAllButton")
        run_layout.addWidget(self.run_all_btn)
        self.stop_btn = QPushButton("STOP (ESC)", clicked=self.cancel_run) # SỬA: Kết nối nút STOP
        self.stop_btn.setObjectName("stopButton")
        self.stop_btn.setEnabled(False) # SỬA: Vô hiệu hóa ban đầu
        run_layout.addWidget(self.stop_btn)
        run_layout.addStretch()
        self.lbl_status = QLabel("Chờ...") # Font được định nghĩa trong dark.qss
        self.lbl_status.setObjectName("statusLabel") # Gán objectName để QSS có thể target
        run_layout.addWidget(self.lbl_status)

        self.main_layout.addWidget(run_widget)

    def _create_realtime_status_bar(self):
        """Tạo thanh trạng thái tọa độ real-time."""
        status_group = QGroupBox("Thông tin Tọa độ (Real-time)")
        status_layout = QHBoxLayout(status_group)

        self.lbl_realtime_status = QLabel("...") # SỬA: Gán vào thuộc tính self
        self.lbl_realtime_status.setWordWrap(True) # Cho phép xuống dòng
        status_layout.addWidget(self.lbl_realtime_status, 1) # Thêm stretch factor

        self.chk_show_realtime_status = QCheckBox("Hiện/Ẩn") # SỬA: Gán vào thuộc tính self
        self.chk_show_realtime_status.stateChanged.connect(self._toggle_realtime_status) # SỬA: Kết nối sự kiện
        # SỬA: Thêm stylesheet để đổi màu dấu check thành xanh lá
        self.chk_show_realtime_status.setStyleSheet("""
            QCheckBox::indicator:checked {
                background-color: #28a745;
                border: 1px solid #218838;
            }
        """)

        status_layout.addWidget(self.chk_show_realtime_status)
        # Thêm vào layout chính với margin
        frame = QFrame()
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(10, 0, 10, 10)
        frame_layout.addWidget(status_group)
        self.main_layout.addWidget(frame)

    def _create_disclaimer(self):
        """Tạo dòng ghi chú cuối cùng."""
        disclaimer_label = QLabel("Lưu ý: Ứng dụng BẮT BUỘC đưa ACSOFT lên foreground (phải focus). Nhấn phím ESC để hủy quá trình chạy.") # Font được định nghĩa trong dark.qss
        disclaimer_label.setObjectName("disclaimerLabel") # Gán objectName để QSS có thể target
        disclaimer_label.setWordWrap(True)
        
        frame = QFrame()
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(10, 0, 10, 10)
        frame_layout.addWidget(disclaimer_label)
        self.main_layout.addWidget(frame)
        # SỬA: Xóa addStretch() để cửa sổ có thể co lại vừa với nội dung
        # self.main_layout.addStretch()

    # --- Window Dragging and Themeing ---

    def _on_title_bar_press(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPosition().toPoint()

    def _on_title_bar_drag(self, event):
        if self._drag_pos:
            diff = event.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + diff)
            self._drag_pos = event.globalPosition().toPoint()

    def toggle_dark_mode(self):
        """Chuyển đổi giao diện sáng/tối."""
        self.is_dark_mode = not self.is_dark_mode # Đảo ngược trạng thái

        stylesheet = ""
        if self.is_dark_mode:
            # Giao diện tối
            try:
                with open('dark.qss', 'r', encoding='utf-8') as f:
                    stylesheet = f.read()
                self.dark_mode_btn.setText("☀")
            except FileNotFoundError:
                print("Lỗi: Không tìm thấy file 'dark.qss'. Sử dụng stylesheet mặc định.")
        else:
            # Giao diện sáng
            try:
                with open('light.qss', 'r', encoding='utf-8') as f:
                    stylesheet = f.read()
                self.dark_mode_btn.setText("◐")
            except FileNotFoundError:
                print("Lỗi: Không tìm thấy file 'light.qss'. Sử dụng stylesheet mặc định.")
        
        # SỬA: Thêm style cho các nút bị vô hiệu hóa (disabled)
        # Style này sẽ được nối vào cuối stylesheet đã tải
        disabled_button_style = """
            QPushButton:disabled {
                background-color: #555;
                color: #888;
                border: 1px solid #666;
            }
        """
        if not self.is_dark_mode:
            # Style cho giao diện sáng
            disabled_button_style = """
                QPushButton:disabled {
                    background-color: #dcdcdc;
                    color: #a0a0a0;
                    border: 1px solid #c0c0c0;
                }
            """
        stylesheet += disabled_button_style

        # Cập nhật màu nền cho các hàng được tô sáng
        if self.is_dark_mode:
            self.highlight_color = QColor("#FFA07A22") # Cam nhạt cho nền tối
            self.default_bg_color = QColor("#3c4049")
        else:
            self.highlight_color = QColor("#FFDAB9") # Màu peachpuff cho nền sáng
            self.default_bg_color = QColor("#ffffff")

        # Xóa tô sáng cũ để áp dụng màu nền mới
        self.clear_all_highlights()

        self.setStyleSheet(stylesheet)
        # self.update() # Không cần thiết nữa khi không dùng paintEvent

    def _toggle_always_on_top(self, checked):
        """SỬA: Bật/tắt chế độ 'Luôn ở trên cùng'."""
        # Lấy các cờ cửa sổ hiện tại
        flags = self.windowFlags()

        if checked:
            # Thêm cờ "Always on Top"
            flags |= Qt.WindowStaysOnTopHint
        else:
            # Xóa cờ "Always on Top"
            flags &= ~Qt.WindowStaysOnTopHint

        # Áp dụng lại các cờ mới
        self.setWindowFlags(flags)
        self.show() # Phải gọi show() lại sau khi thay đổi cờ cửa sổ

    def _update_edit_button_text(self):
        """SỬA: Cập nhật văn bản của nút 'Sửa Dòng' dựa trên lựa chọn hiện tại."""
        selected_items = self.tree_macro.selectedItems()
        if selected_items:
            # Lấy chỉ số hàng từ item đầu tiên được chọn
            row = selected_items[0].row()
            # Số thứ tự hiển thị là chỉ số hàng + 1
            step_number = row + 1
            self.edit_step_btn.setText(f"Sửa Dòng #{step_number}")
        else:
            self.edit_step_btn.setText("Sửa Dòng")

    # SỬA: Thêm các hàm xử lý cho status bar
    def _toggle_realtime_status(self, state):
        """Bật/tắt bộ đếm thời gian cập nhật trạng thái real-time."""
        if state == Qt.Checked.value:
            # Cập nhật ngay lần đầu, sau đó bắt đầu timer
            self._update_status_bar_info()
            self.realtime_status_timer.start(200) # Cập nhật mỗi 200ms
        else:
            self.realtime_status_timer.stop()
            self.lbl_realtime_status.setText("...") # Reset text khi tắt

    def _update_status_bar_info(self):
        """Lấy thông tin và cập nhật vào label trạng thái."""
        if not self.chk_show_realtime_status.isChecked():
            return

        try:
            hwnd = hwnd_from_title(self.target_window_title)
            # 1. Kích thước màn hình (Logical Pixels)
            screen_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

            # 2. Vị trí con trỏ (Tuyệt đối, Logical Pixels)
            cursor_x, cursor_y = win32api.GetCursorPos()

            status_parts = [f"Màn hình: <b>{screen_w}x{screen_h}px</b>"]

            if hwnd and self.target_window_title:
                rect = get_window_rect(hwnd)
                if rect:
                    left, top, right, bottom = rect
                    width, height = right - left, bottom - top
                    x_offset_logical, y_offset_logical = cursor_x - left, cursor_y - top
                    dpi_scale = get_dpi_scale_factor(hwnd)
                    scale = dpi_scale if dpi_scale > 0 else 1.0
                    x_norm = int(x_offset_logical / scale)
                    y_norm = int(y_offset_logical / scale)

                    status_parts.extend([
                        f"<b>DPI Scale:</b> {int(dpi_scale * 100)}%",
                        f"<b>Cửa sổ ({width}x{height}px):</b> L{left} T{top}",
                        f"<b>Chuột Tuyệt đối:</b> X{cursor_x} Y{cursor_y}",
                        f"<b>Chuột Offset Chuẩn (100%):</b> Xo{x_norm} Yo{y_norm}",
                    ])
                else:
                    status_parts.append(f"Cửa sổ '{self.target_window_title}' không tìm thấy hoặc bị ẩn.")
            else:
                status_parts.append("Vui lòng chọn <b>Cửa sổ mục tiêu</b>.")

            self.lbl_realtime_status.setText(" | ".join(status_parts))
        except Exception as e:
            self.lbl_realtime_status.setText(f"Lỗi cập nhật trạng thái: {e}")
    # SỬA: Thêm logic cho Group 1
    def browse_ac(self):
        """Mở hộp thoại chọn file .exe cho ACSOFT."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file chạy ACSOFT (.exe)",
            "", # Thư mục mặc định
            "Executable files (*.exe);;All files (*.*)"
        )
        if file_path:
            self.txt_acpath.setText(file_path)
            self.acsoft_path = file_path

    def open_ac(self):
        """Mở ứng dụng ACSOFT."""
        p = self.txt_acpath.text().strip()
        if not p or not os.path.isfile(p):
            QMessageBox.warning(self, "Lỗi", "Chưa chọn file exe hợp lệ.")
            return
        try:
            subprocess.Popen([p])
            QMessageBox.information(self, "Đã mở", "Đã mở ACSOFT (chờ phần mềm khởi động).")
            # SỬA: Giả định có hàm refresh_windows() sẽ được thêm sau
            # QTimer.singleShot(2000, self.refresh_windows) 
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Mở không thành công: {e}")

    def browse_csv(self):
        """Mở hộp thoại chọn file CSV."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file CSV chứa dữ liệu",
            "", # Thư mục mặc định
            "CSV files (*.csv);;All files (*.*)"
        )
        if file_path:
            self.txt_csv.setText(file_path)
            self.csv_path = file_path
            self.load_csv_data(file_path)

    def load_csv_data(self, path):
        """Đọc dữ liệu từ file CSV và hiển thị lên bảng."""
        delimiter = self.txt_delimiter.text().strip()
        if not delimiter:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập ký tự phân cách (Delimiter).")
            return

        try:
            # SỬA: Gán dữ liệu vào self.df_csv thay vì biến cục bộ
            self.df_csv = pd.read_csv(path,
                                      header=None,
                                      dtype=str,
                                      keep_default_na=False,
                                      sep=delimiter)

            self.tree_csv.clearContents() # Xóa nội dung cũ, giữ lại header

            num_cols = len(self.df_csv.columns)
            columns = [f"Cột {i + 1}" for i in range(num_cols)]
            self.tree_csv.setColumnCount(num_cols)
            self.tree_csv.setHorizontalHeaderLabels(columns)
            self.tree_csv.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

            rows_to_display = self.df_csv.head(10)
            self.tree_csv.setRowCount(len(rows_to_display))
            for row_index, row_data in rows_to_display.iterrows():
                for col_index, value in enumerate(row_data):
                    item = QTableWidgetItem(str(value))
                    self.tree_csv.setItem(row_index, col_index, item)

        except Exception as e:
            self.df_csv = pd.DataFrame() # SỬA: Đảm bảo df_csv rỗng nếu có lỗi
            QMessageBox.critical(self, "Lỗi đọc CSV", f"Không thể đọc file CSV bằng delimiter '{delimiter}'. Lỗi: {e}")
            self.tree_csv.clearContents()
            self.tree_csv.setRowCount(0)

    # SỬA: Thêm logic cho việc làm mới và chọn cửa sổ
    def refresh_windows(self):
        """Lấy danh sách các cửa sổ đang mở và cập nhật vào ComboBox."""
        titles = []

        def enum_handler(hwnd, titles_list):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title not in titles_list:
                    titles_list.append(title)
            return True

        win32gui.EnumWindows(enum_handler, titles)
        titles.sort()

        self.combo_windows.clear()
        self.combo_windows.addItems(titles)

        if titles:
            initial_select = None
            for title in titles:
                if "acsoft" in title.lower() or "kế toán" in title.lower() or "việt tín" in title.lower():
                    initial_select = title
                    break
            
            if initial_select:
                self.combo_windows.setCurrentText(initial_select)
            else:
                self.combo_windows.setCurrentIndex(0) # Chọn mục đầu tiên nếu không tìm thấy
        else:
            self.combo_windows.addItem("Không tìm thấy cửa sổ")

    def on_window_select(self, text):
        """Lưu lại tiêu đề cửa sổ khi người dùng chọn từ ComboBox."""
        self.target_window_title = text

    # --- SỬA: THÊM LOGIC GHI MACRO ---

    def record_macro(self):
        """Bắt đầu quá trình ghi macro."""
        if self.recording:
            return
        if not self.target_window_title:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn cửa sổ mục tiêu ACSOFT trước.")
            return

        # Dọn dẹp macro cũ
        self.macro_steps.clear()
        self.tree_macro.setRowCount(0)
        self.current_col_index = 0
        self.recording = True
        self._recording_finished_processed = False # Reset flag for new recording session
        self._set_buttons_for_running(True) # SỬA: Cập nhật trạng thái các nút
        
        # Hiển thị HUD
        self.hud_window = RecordingHUD(self)
        self.hud_window.stop_button.clicked.connect(self.stop_recording)
        self.hud_window.show()
 
        # Tạo và chạy worker trong thread mới
        self.recording_thread = QThread()
        self.recording_worker = RecordingWorker(self.target_window_title, self)
        self.recording_worker.moveToThread(self.recording_thread)

        # Kết nối tín hiệu từ worker đến các slot trong GUI thread
        self.recording_worker.update_hud_signal.connect(self.update_hud_status)
        self.recording_worker.add_step_signal.connect(self.add_macro_step_to_gui)
        self.recording_worker.recording_finished_signal.connect(self.on_recording_finished)

        # Bắt đầu thread
        self.recording_thread.started.connect(self.recording_worker.run)
        self.recording_thread.start()

    def stop_recording(self):
        """Dừng quá trình ghi macro (được gọi từ nút Stop trên HUD hoặc phím ESC)."""
        if self.recording_worker:
            self.recording_worker.stop()

    def on_recording_finished(self, completed):
        """Dọn dẹp sau khi worker kết thúc."""
        self.recording = False

        # Dọn dẹp thread
        if self.recording_thread:
            self.recording_thread.quit()
            self.recording_thread.wait()
            self.recording_thread = None
            self.recording_worker = None

        # Đóng HUD
        if self.hud_window:
            self.hud_window.close()
            self.hud_window = None
            
        self._set_buttons_for_running(False) # SỬA: Khôi phục trạng thái các nút

        # Hiển thị thông báo
        if completed:
            QMessageBox.information(self, "Hoàn thành", f"Đã ghi xong macro với {len(self.macro_steps)} bước.")
        else:
            QMessageBox.information(self, "Đã hủy", "Quá trình ghi macro đã được hủy.")

    def update_hud_status(self, text, color):
        """Slot để cập nhật HUD từ thread worker."""
        if self.hud_window:
            self.hud_window.update_status(text, color)

    def add_macro_step_to_gui(self, step):
        """Thêm một bước macro vào danh sách và bảng hiển thị."""
        step.item_idx = len(self.macro_steps)
        self.macro_steps.append(step)

        row_position = self.tree_macro.rowCount()
        self.tree_macro.insertRow(row_position)

        self.tree_macro.setItem(row_position, 0, QTableWidgetItem(str(step.item_idx + 1)))
        self.tree_macro.setItem(row_position, 1, QTableWidgetItem(step.typ.upper()))
        
        description_item = QTableWidgetItem(repr(step))
        description_item.setData(Qt.UserRole, step) # Lưu đối tượng MacroStep vào item
        self.tree_macro.setItem(row_position, 2, description_item)
        
        self.tree_macro.scrollToBottom()

    def stop_listeners(self):
        """Dừng các listener pynput."""
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener = None
        if self.keyboard_listener:
            self.keyboard_listener.stop()
            self.keyboard_listener = None

    # --- Các hàm xử lý sự kiện từ pynput (chạy trong thread của pynput) ---

    def _on_mouse_click(self, x, y, button, pressed):
        if not self.recording or not pressed:
            return

        hwnd = hwnd_from_title(self.target_window_title)
        if not hwnd or win32gui.GetForegroundWindow() != hwnd:
            return # Chỉ ghi khi cửa sổ mục tiêu đang được focus

        rect = get_window_rect(hwnd)
        if not rect: return
        left, top, _, _ = rect

        cursor_x, cursor_y = win32api.GetCursorPos()
        x_offset_logical = cursor_x - left
        y_offset_logical = cursor_y - top

        current_time = time.time()
        delay = current_time - self.last_key_time
        self.last_key_time = current_time

        click_type = 'right_click' if button == mouse.Button.right else 'left_click'
        current_scale = get_dpi_scale_factor(hwnd)

        step = MacroStep('mouse', key_value=click_type, delay_after=delay,
                         x_offset=x_offset_logical, y_offset=y_offset_logical, dpi_scale=current_scale)
        self.recording_worker.add_step_signal.emit(step)

    def _on_key_press(self, key):
        #nếu đang chạy phát macro thì gọi hàm self.cancel_run()
        if key == keyboard.Key.esc and self.run_worker:
            self.cancel_run()
            return
        if not self.recording: return

        if key in [Key.ctrl_l, Key.ctrl_r, Key.alt_l, Key.alt_r, Key.shift_l, Key.shift_r]:
            self.current_modifiers.add(key)
            return

        current_time = time.time()
        delay = current_time - self.last_key_time
        self.last_key_time = current_time

        typ, key_name, col_index = "", "", None

        if self.current_modifiers:
            typ = 'combo'
            modifier_names = sorted([str(m).replace("Key.", "") for m in self.current_modifiers])
            main_key = key.char.lower() if hasattr(key, 'char') and key.char else str(key).replace("Key.", "")
            key_name = "+".join(modifier_names) + "+" + main_key
        elif key == Key.insert:
            typ, col_index = 'col', self.current_col_index
            self.current_col_index += 1
        else:
            typ = 'key'
            key_name = key.char if hasattr(key, 'char') and key.char else str(key).replace("Key.", "")

        if typ:
            step = MacroStep(typ, key_value=key_name, col_index=col_index, delay_after=delay)
            self.recording_worker.add_step_signal.emit(step)

    def _on_key_release(self, key):
        if key in self.current_modifiers:
            self.current_modifiers.discard(key)

    def clear_macro(self):
        """Xóa tất cả các bước macro khỏi danh sách và bảng hiển thị."""
        reply = QMessageBox.question(self, 'Xác nhận xóa', 'Bạn có chắc chắn muốn xóa toàn bộ macro không?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.macro_steps.clear()
            self.tree_macro.setRowCount(0)
            self.current_col_index = 0 # Đặt lại chỉ số cột

    def save_macro(self):
        """Lưu macro vào file JSON."""
        if not self.macro_steps:
            QMessageBox.warning(self, "Lỗi", "Chưa có macro nào để lưu.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu file Macro",
            "",
            "JSON files (*.json);;All files (*.*)"
        )

        if file_path:
            try:
                macro_data = [step.to_dict() for step in self.macro_steps]
                settings_data = self._collect_app_settings()

                full_data = {
                    "app_settings": settings_data,
                    "macro_steps": macro_data
                }

                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(full_data, f, indent=4)

                QMessageBox.information(self, "Thành công", f"Đã lưu macro vào:\n{file_path}")

            except Exception as e:
                QMessageBox.critical(self, "Lỗi Lưu", f"Không thể lưu file macro. Lỗi: {e}")

    def _collect_app_settings(self):
        """Thu thập các cài đặt của ứng dụng để lưu."""
        settings = {}

        # Group 1
        settings['acsoft_path'] = self.txt_acpath.text() if self.txt_acpath else ""

        # Group 2
        settings['csv_path'] = self.txt_csv.text() if self.txt_csv else ""
        settings['delimiter'] = self.txt_delimiter.text() if self.txt_delimiter else ";"
        settings['target_window_title'] = self.target_window_title

        # Group 3
        settings['speed_mode'] = 1 if self.radio_recorded_speed.isChecked() else 2
        settings['custom_speed_ms'] = self.spin_fixed_speed.value()
        settings['delay_between_rows_s'] = self.spin_delay_between_rows.value()

        return settings

    def load_macro(self):
        """Tải macro từ file JSON."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Mở file Macro",
            "",
            "JSON files (*.json);;All files (*.*)"
        )

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    full_data = json.load(f)

                macro_data = full_data.get("macro_steps", [])
                self.macro_steps = [MacroStep.from_dict(data) for data in macro_data]
                self.populate_macro_table()

                settings_data = full_data.get("app_settings", {})
                self._apply_app_settings(settings_data)

                # SỬA: Tự động load CSV nếu đường dẫn có trong settings
                csv_path_from_settings = settings_data.get('csv_path')
                if csv_path_from_settings and os.path.exists(csv_path_from_settings):
                    self.load_csv_data(csv_path_from_settings)

                QMessageBox.information(self, "Thành công", f"Đã tải macro từ:\n{file_path}")

            except Exception as e:
                QMessageBox.critical(self, "Lỗi Mở", f"Không thể mở file macro. Lỗi: {e}")

    def add_manual_step(self, step_type):
        """Thêm bước macro thủ công với giá trị mặc định và mở cửa sổ sửa."""
        # Dựa vào số bước hiện tại để gợi ý cột tiếp theo
        next_col_index = sum(1 for step in self.macro_steps if step.typ == 'col')
        current_scale = get_dpi_scale_factor(None)  # Lấy DPI hệ thống

        initial_step = None
        if step_type == 'col':
            initial_step = MacroStep(step_type, col_index=next_col_index)
        elif step_type == 'key':
            initial_step = MacroStep(step_type, key_value='ENTER')
        elif step_type == 'combo':
            initial_step = MacroStep(step_type, key_value='CTRL+C')
        elif step_type == 'mouse':
            # Gợi ý tọa độ (100, 100) được scale theo DPI hiện tại để người dùng dễ sửa
            # PySide6 uses logical pixels directly, so no need to scale initial values here
            initial_x = 100.0
            initial_y = 100.0
            initial_step = MacroStep(step_type, key_value='left_click', x_offset=initial_x, y_offset=initial_y,
                                     dpi_scale=current_scale)
        elif step_type == 'end':
            initial_step = MacroStep('end')
        else:
            return

        # Thêm vào list và tableview
        self.add_macro_step_to_gui(initial_step)

        # Tự động mở cửa sổ sửa để người dùng điều chỉnh
        # Pass the step object itself, or its index
        self.edit_macro_step(initial_step.item_idx)

    def edit_macro_step(self, step_index):
        """Mở cửa sổ chỉnh sửa cho bước macro được chọn. step_index dùng cho add_manual_step."""
        selected_step = None
        # if step_index is not None:
        #     # If an index is provided, find the step by index
        #     selected_step = next((step for step in self.macro_steps if step.item_idx == step_index), None)
        # else:
        #     # If no index, get the currently selected row onin the QTableWidget
        row = self.tree_macro.currentRow() # SỬA: Dùng currentRow() để lấy hàng đang chọn
        if row < 0: # currentRow() trả về -1 nếu không có hàng nào được chọn
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn một bước macro để sửa.")
            return
        
        # Lấy đối tượng MacroStep từ QTableWidgetItem ở cột 2 (Mô tả chi tiết)
        item = self.tree_macro.item(row, 2)
        if item:
            selected_step = item.data(Qt.UserRole) # Lấy đối tượng MacroStep đã lưu

        if not selected_step:
            QMessageBox.warning(self, "Lỗi", "Không tìm thấy bước macro để sửa.")
            return

        dialog = MacroStepEditDialog(self, selected_step)
        if dialog.exec() == QDialog.Accepted:
            # get_edited_data also performs validation
            new_typ = dialog.type_combo.currentText()
            new_delay_after = dialog.delay_spin.value() / 1000.0

            # Reset values first
            selected_step.col_index = None
            selected_step.key_value = None
            selected_step.x_offset_logical = None
            selected_step.y_offset_logical = None

            try:
                if new_typ == 'col':
                    selected_step.col_index = dialog.col_index_spin.value()
                elif new_typ in ['key', 'combo']:
                    selected_step.key_value = dialog.key_value_edit.text().upper().strip()
                elif new_typ == 'mouse':
                    selected_step.key_value = dialog.key_value_edit.text().lower().strip()
                    selected_step.x_offset_logical = float(dialog.x_offset_edit.text())
                    selected_step.y_offset_logical = float(dialog.y_offset_edit.text())

                selected_step.typ = new_typ
                selected_step.delay_after = new_delay_after

                # Update the QTableWidget row
                # Tìm lại hàng trong bảng dựa trên item_idx (nếu cần) hoặc sử dụng row đã chọn
                current_row_in_table = self.macro_steps.index(selected_step) # Lấy vị trí hiện tại của step trong list
                self.tree_macro.setItem(current_row_in_table, 0, QTableWidgetItem(str(current_row_in_table + 1)))
                self.tree_macro.setItem(current_row_in_table, 1, QTableWidgetItem(selected_step.typ.upper()))
                
                description_item = QTableWidgetItem(repr(selected_step))
                description_item.setData(Qt.UserRole, selected_step) # Cập nhật lại đối tượng MacroStep đã lưu
                self.tree_macro.setItem(current_row_in_table, 2, description_item)
                QMessageBox.information(self, "Thành công", "Đã sửa bước macro thành công.")
            except ValueError as e:
                QMessageBox.warning(self, "Lỗi", f"Dữ liệu không hợp lệ: {e}. Vui lòng kiểm tra lại.")
            except Exception as e:
                QMessageBox.critical(self, "Lỗi", f"Không thể lưu thay đổi: {e}")
        else:
            QMessageBox.information(self, "Hủy bỏ", "Không có thay đổi nào được lưu.")

    def delete_macro_step(self):
        """Xóa bước macro được chọn."""
        # ... (logic for deleting macro step, see below)
        selected_rows = self.tree_macro.selectedIndexes()
        if not selected_rows:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn một bước macro để xóa.")
            return

        row_to_delete = selected_rows[0].row()

        reply = QMessageBox.question(self, "Xác nhận Xóa", "Bạn có chắc chắn muốn xóa bước macro này?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                # Remove from the list
                del self.macro_steps[row_to_delete]
                # Remove from the QTableWidget
                self.tree_macro.removeRow(row_to_delete)
                # Re-populate to update STT and item_idx
                self.populate_macro_table()
                QMessageBox.information(self, "Thành công", "Đã xóa bước macro.")
            except Exception as e:
                QMessageBox.critical(self, "Lỗi Xóa", f"Không thể xóa bước macro. Lỗi: {e}")

    def populate_macro_table(self):
        """
        Xóa và điền lại bảng macro (self.tree_macro) từ danh sách self.macro_steps.
        Hàm này cũng sẽ tính toán lại chỉ số cho các bước 'col'.
        """
        self.tree_macro.setRowCount(0)
        self.current_col_index = 0 # Reset lại bộ đếm cột

        for idx, step in enumerate(self.macro_steps):
            step.item_idx = idx
            if step.typ == 'col':
                step.col_index = self.current_col_index
                self.current_col_index += 1

            row_position = self.tree_macro.rowCount()
            self.tree_macro.insertRow(row_position)
            self.tree_macro.setItem(row_position, 0, QTableWidgetItem(str(idx + 1))) # STT
            self.tree_macro.setItem(row_position, 1, QTableWidgetItem(step.typ.upper()))
            
            description_item = QTableWidgetItem(repr(step))
            description_item.setData(Qt.UserRole, step) # Lưu đối tượng MacroStep vào item
            self.tree_macro.setItem(row_position, 2, description_item)

    def _apply_app_settings(self, settings):
        """Áp dụng các cài đặt đã lưu."""
        # Group 1
        if 'acsoft_path' in settings and self.txt_acpath:
            self.txt_acpath.setText(settings['acsoft_path'])

        # Group 2
        if 'csv_path' in settings and self.txt_csv:
            self.txt_csv.setText(settings['csv_path'])
        if 'delimiter' in settings and self.txt_delimiter:
            self.txt_delimiter.setText(settings['delimiter'])
        if 'target_window_title' in settings and self.combo_windows:
            index = self.combo_windows.findText(settings['target_window_title'])
            if index >= 0:
                self.combo_windows.setCurrentIndex(index)

        # Group 3
        if 'speed_mode' in settings:
            if settings['speed_mode'] == 1:
                self.radio_recorded_speed.setChecked(True)
            else:
                self.radio_fixed_speed.setChecked(True)
        if 'custom_speed_ms' in settings and self.spin_fixed_speed:
            self.spin_fixed_speed.setValue(settings['custom_speed_ms'])


    def on_test(self):
        """Chạy macro thử nghiệm (1 dòng)."""
        self._run_macro(test_mode=True)

    def on_run_all(self):
        """Chạy macro cho tất cả các dòng trong file CSV."""
        self._run_macro(test_mode=False)
        
    # --- SỬA: THÊM LOGIC CHẠY MACRO ---
    def _run_macro(self, test_mode):
        """Chuẩn bị và bắt đầu chạy macro trong một thread riêng."""
        if not self.target_window_title:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn cửa sổ mục tiêu hợp lệ trước.")
            return

        hwnd = hwnd_from_title(self.target_window_title)
        if not hwnd:
            QMessageBox.warning(self, "Lỗi", f"Không tìm thấy cửa sổ: '{self.target_window_title}'. Vui lòng làm mới và chọn lại.")
            return

        if self.df_csv.empty:
            QMessageBox.warning(self, "Lỗi", "Vui lòng load file CSV có dữ liệu.")
            return
        if not self.macro_steps:
            QMessageBox.warning(self, "Lỗi", "Chưa có bước macro nào được ghi.")
            return

        # Reset cờ hủy và cập nhật UI
        self.cancel_run_flag.clear()
        self.pause_run_flag.set() # SỬA: Đặt cờ pause về trạng thái "chạy" (set) ban đầu
        self._set_buttons_for_running(True) # SỬA: Cập nhật trạng thái các nút

        # Hiển thị HUD
        self.hud_window = RecordingHUD(self)
        self.hud_window.stop_button.clicked.connect(self.cancel_run)
        self.hud_window.show()

        # SỬA: Hiển thị và kết nối nút Pause trên HUD
        self.hud_window.pause_button.show()
        self.hud_window.f5_label.show() # SỬA: Hiện cả ghi chú F5
        self.hud_window.pause_event = self.pause_run_flag # Truyền cờ pause vào HUD
        self.hud_window.pause_button.clicked.connect(self.hud_window.toggle_pause)

        # SỬA: Kết nối signal với slot. Đây là cách giao tiếp an toàn và đáng tin cậy nhất.
        # Khi signal được emit từ bất kỳ luồng nào, slot sẽ được thực thi trên luồng chính của GUI.
        self.toggle_pause_signal.connect(self.hud_window.toggle_pause)

        # Lấy các cài đặt hiện tại
        use_recorded_speed = self.radio_recorded_speed.isChecked()
        custom_delay_s = self.spin_fixed_speed.value() / 1000.0
        row_delay_s = self.spin_delay_between_rows.value()

        # Tạo và chạy worker trong thread mới
        self.run_thread = QThread()
        self.run_worker = MacroRunnerWorker(
            target_window_title=self.target_window_title,
            macro_steps=self.macro_steps,
            df_csv=self.df_csv,
            test_mode=test_mode,
            use_recorded_speed=use_recorded_speed,
            custom_delay_s=custom_delay_s,
            row_delay_s=row_delay_s,
            cancel_flag=self.cancel_run_flag,
            pause_flag=self.pause_run_flag # SỬA: Truyền cờ pause vào worker
        )
        self.run_worker.moveToThread(self.run_thread)

        # Kết nối tín hiệu
        self.run_worker.update_hud_signal.connect(self.update_hud_status)
        self.run_worker.update_status_signal.connect(self.lbl_status.setText)
        self.run_worker.run_finished_signal.connect(self.on_run_finished)
        self.run_worker.highlight_csv_row_signal.connect(self.highlight_csv_row)
        self.run_worker.highlight_macro_step_signal.connect(self.highlight_macro_step)

        self.run_thread.started.connect(self.run_worker.run)
        self.run_thread.start()

    def cancel_run(self):
        """Đặt cờ để hủy quá trình chạy macro."""
        self.cancel_run_flag.set()
        if self.recording_worker: # Nếu đang ghi thì cũng hủy
            self.stop_recording()

    def on_run_finished(self, completed, message):
        """Dọn dẹp sau khi chạy macro xong."""
        self._set_buttons_for_running(False) # SỬA: Khôi phục trạng thái các nút

        self.lbl_status.setText(message)
        QMessageBox.information(self, "Thông báo", message)

        # Dọn dẹp thread
        if self.run_thread:
            self.run_thread.quit()
            self.run_thread.wait()
            self.run_thread = None
            self.run_worker = None

        # Đóng HUD
        if self.hud_window:
            self.hud_window.close()
            self.hud_window = None
        
        # Xóa tô sáng khi kết thúc
        self.clear_all_highlights()

    def _set_buttons_for_running(self, is_running):
        """Hàm tiện ích để bật/tắt các nút điều khiển khi chạy/dừng."""
        self.record_btn.setEnabled(not is_running)
        self.run_test_btn.setEnabled(not is_running)
        self.run_all_btn.setEnabled(not is_running)
        self.stop_btn.setEnabled(is_running)
        self.lbl_status.setText("Đang chạy..." if is_running else "Chờ...")
        # SỬA: Ẩn ghi chú F5 khi không chạy
        if not is_running and self.hud_window:
            self.hud_window.f5_label.hide()

    def _start_global_hotkey_listener(self):
        """SỬA: Khởi động listener bàn phím toàn cục để bắt phím nóng (ESC, F5)."""
        def on_press_global(key):
            if key == Key.esc:
                # SỬA: Gọi trực tiếp self.cancel_run(). Hàm này đã được thiết kế thread-safe.
                self.cancel_run()
            
            # SỬA: Thêm logic cho phím F5
            elif key == Key.f5:
                # Chỉ hoạt động khi đang chạy macro và HUD đang hiển thị
                if self.run_worker and self.hud_window:
                    # SỬA: Không dùng QTimer.singleShot hoặc gọi trực tiếp.
                    # Thay vào đó, phát một tín hiệu (signal) đã được kết nối sẵn.
                    self.toggle_pause_signal.emit()

        self.global_hotkey_listener = keyboard.Listener(on_press=on_press_global, daemon=True)
        self.global_hotkey_listener.start()

    def closeEvent(self, event):
        """Xử lý sự kiện đóng cửa sổ."""
        if self.global_esc_listener:
            self.global_esc_listener.stop()
        # Gọi hàm closeEvent của lớp cha để đảm bảo các dọn dẹp khác được thực hiện
        super().closeEvent(event)

    # SỬA: Thêm signal để giao tiếp an toàn giữa các luồng
    toggle_pause_signal = Signal()

    def highlight_csv_row(self, row_index):
        """Tô sáng một hàng trong bảng CSV."""
        # Chỉ tô sáng nếu hàng đó nằm trong 10 hàng đang hiển thị
        if row_index < self.tree_csv.rowCount():
            self._highlight_row(self.tree_csv, row_index, 'csv')

    def highlight_macro_step(self, step_index):
        """Tô sáng một bước trong bảng Macro."""
        self._highlight_row(self.tree_macro, step_index, 'macro')

    def _highlight_row(self, table_widget, row_index, table_type):
        """Hàm chung để tô sáng một hàng trong QTableWidget."""
        last_highlighted_attr = f'last_highlighted_{table_type}_row'
        last_row = getattr(self, last_highlighted_attr, -1)

        # Bỏ tô sáng hàng cũ
        if last_row != -1 and last_row < table_widget.rowCount():
            for col in range(table_widget.columnCount()):
                item = table_widget.item(last_row, col)
                if item:
                    item.setBackground(self.default_bg_color)

        # Tô sáng hàng mới
        if row_index < table_widget.rowCount():
            for col in range(table_widget.columnCount()):
                item = table_widget.item(row_index, col)
                if item:
                    item.setBackground(self.highlight_color)
            table_widget.scrollToItem(table_widget.item(row_index, 0), QAbstractItemView.ScrollHint.PositionAtCenter)

        setattr(self, last_highlighted_attr, row_index)

    def clear_all_highlights(self):
        """Xóa tất cả các tô sáng trên cả hai bảng."""
        self._highlight_row(self.tree_csv, -1, 'csv')
        self._highlight_row(self.tree_macro, -1, 'macro')


# =========================================================================
# -------------------- Macro Runner Worker (Thread) -----------------------
# =========================================================================
class MacroRunnerWorker(QObject):
    update_hud_signal = Signal(str, str)
    update_status_signal = Signal(str)
    run_finished_signal = Signal(bool, str)
    highlight_csv_row_signal = Signal(int)
    highlight_macro_step_signal = Signal(int)

    def __init__(self, target_window_title, macro_steps, df_csv, test_mode, use_recorded_speed, custom_delay_s, row_delay_s, cancel_flag, pause_flag):
        super().__init__()
        self.target_window_title = target_window_title
        self.macro_steps = macro_steps
        self.df_csv = df_csv
        self.test_mode = test_mode
        self.use_recorded_speed = use_recorded_speed
        self.custom_delay_s = custom_delay_s
        self.row_delay_s = row_delay_s
        self.cancel_flag = cancel_flag
        self.pause_flag = pause_flag # SỬA: Lưu lại cờ pause

    def run(self):
        """Hàm chính của worker, thực thi macro."""
        try:
            hwnd = hwnd_from_title(self.target_window_title)
            if not hwnd:
                self.run_finished_signal.emit(False, "Lỗi: Không tìm thấy cửa sổ mục tiêu.")
                return

            # Đếm ngược
            for i in range(5, 0, -1):
                if self.cancel_flag.is_set():
                    self.run_finished_signal.emit(False, "Đã hủy bởi người dùng.")
                    return
                self.update_hud_signal.emit(f"Bắt đầu chạy sau: {i}s", "#87CEEB")
                # SỬA: Thay thế time.sleep(1) bằng vòng lặp không chặn để ESC hoạt động ngay
                delay_start_time = time.time()
                while time.time() - delay_start_time < 1.0:
                    if self.cancel_flag.is_set():
                        break
                    # Chờ một khoảng rất ngắn và kiểm tra lại
                    time.sleep(0.05)

            self.update_hud_signal.emit("► ĐANG CHẠY...", "#98FB98")

            rows_to_run = self.df_csv.head(1) if self.test_mode else self.df_csv

            for row_index, row_data in rows_to_run.iterrows():
                if self.cancel_flag.is_set(): break
                self.highlight_csv_row_signal.emit(row_index) # Gửi tín hiệu tô sáng dòng CSV
                self.update_status_signal.emit(f"Đang chạy dòng CSV số: {row_index + 1}/{len(rows_to_run)}...")

                # SỬA: Kiểm tra trạng thái pause trước khi xử lý mỗi dòng
                # SỬA: Thay thế pause_flag.wait() bằng vòng lặp để có thể kiểm tra cancel_flag
                while not self.pause_flag.is_set():
                    if self.cancel_flag.is_set():
                        break # Thoát khỏi vòng lặp pause nếu bị hủy
                    time.sleep(0.1) # Chờ một chút trước khi kiểm tra lại
                if self.cancel_flag.is_set(): break # Nếu bị hủy trong lúc pause, thoát khỏi vòng lặp dòng

                for step in self.macro_steps:
                    if self.cancel_flag.is_set(): break
                    
                    self.highlight_macro_step_signal.emit(step.item_idx) # Gửi tín hiệu tô sáng bước macro
                    # Thực thi bước macro
                    if step.typ == 'col':
                        if step.col_index is not None and step.col_index < len(row_data):
                            send_char_to_hwnd(hwnd, str(row_data.iloc[step.col_index]))
                    elif step.typ == 'key':
                        send_key_to_hwnd(hwnd, step.key_value)
                    elif step.typ == 'combo':
                        send_combo_to_hwnd(hwnd, step.key_value)
                    elif step.typ == 'mouse':
                        send_mouse_click(hwnd, step.x_offset_logical, step.y_offset_logical, step.key_value, step.dpi_scale)

                    # SỬA: Kiểm tra trạng thái pause sau mỗi bước
                    # SỬA: Thay thế pause_flag.wait() bằng vòng lặp để có thể kiểm tra cancel_flag
                    while not self.pause_flag.is_set():
                        if self.cancel_flag.is_set():
                            break # Thoát khỏi vòng lặp pause nếu bị hủy
                        time.sleep(0.1) # Chờ một chút trước khi kiểm tra lại
                    if self.cancel_flag.is_set(): break # Nếu bị hủy trong lúc pause, thoát khỏi vòng lặp bước

                    # Chờ theo độ trễ đã cấu hình
                    # SỬA: Thay thế time.sleep() bằng vòng lặp không chặn để ESC hoạt động ngay
                    delay = step.delay_after if self.use_recorded_speed else self.custom_delay_s
                    delay_start_time = time.time()
                    while time.time() - delay_start_time < delay:
                        if self.cancel_flag.is_set():
                            break
                        time.sleep(0.05) # Chờ một khoảng rất ngắn và kiểm tra lại

                if self.cancel_flag.is_set(): break

                # SỬA: Thay thế time.sleep() giữa các dòng bằng vòng lặp không chặn
                delay_start_time = time.time()
                while time.time() - delay_start_time < self.row_delay_s:
                    if self.cancel_flag.is_set():
                        break
                    time.sleep(0.05)

            if self.cancel_flag.is_set():
                self.run_finished_signal.emit(False, "Đã hủy bởi người dùng.")
            else:
                self.run_finished_signal.emit(True, "Hoàn thành!")

        except Exception as e:
            self.run_finished_signal.emit(False, f"Lỗi khi chạy: {e}")


if __name__ == "__main__":
    # SỬA: Đặt AppUserModelID để Windows hiển thị đúng icon trên taskbar
    # Đây là bước quan trọng để Windows không nhóm ứng dụng của bạn với Python.
    myappid = 'viettin.autosoft.sender.2025' # Chuỗi ID duy nhất cho ứng dụng
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication(sys.argv)
    
    # Hiển thị HUD để kiểm tra
    # hud = RecordingHUD()
    # hud.show()

    # Hiển thị cửa sổ chính
    main_win = MacroApp()
    main_win.show()

    sys.exit(app.exec())