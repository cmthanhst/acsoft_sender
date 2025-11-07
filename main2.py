import sys
import os
import base64
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit,
    QFileDialog, QMessageBox, QFrame, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QSpinBox,
    QComboBox, QDialog, QGroupBox, QRadioButton
)
from PySide6.QtGui import QIcon, QPixmap, QFont, QColor, QCursor, QPainter, QBrush, QRegion
from PySide6.QtCore import Qt, QPoint, QSize

# Đọc nội dung Base64 từ file
try:
    with open('logo_base64.txt', 'r') as f:
        LOGO_PNG_BASE64 = f.read().strip()
except FileNotFoundError:
    print("Lỗi: Không tìm thấy file 'logo_base64.txt'. Sẽ sử dụng ảnh placeholder.")
    LOGO_PNG_BASE64 = "" # Để trống nếu không tìm thấy

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

        # Thiết lập cửa sổ không viền, luôn ở trên và trong suốt
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool) # Qt.Tool để không hiện trên taskbar
        self.setAttribute(Qt.WA_TranslucentBackground, True) # Nền trong suốt

        # Frame chính với nền tối và bo góc (style được định nghĩa trong dark.qss)
        self.main_frame = QFrame(self) # Tên object "mainFrame" để stylesheet có thể target
        self.main_frame.setObjectName("mainFrame") 

        # Layout cho các thành phần trong HUD
        layout = QHBoxLayout(self.main_frame)
        layout.setContentsMargins(10, 5, 10, 5)

        self.status_label = QLabel("Chuẩn bị...")
        self.status_label.setFont(QFont("Courier", 10))
        layout.addWidget(self.status_label) # Style cho QLabel được định nghĩa trong dark.qss

        self.pause_button = QPushButton("❚❚ PAUSE")
        self.pause_button.setFont(QFont("Courier", 9, QFont.Bold))
        self.pause_button.hide() # Ẩn ban đầu
        layout.addWidget(self.pause_button)

        self.stop_button = QPushButton("■ STOP")
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

# =========================================================================
# -------------------- Main Application Window (PySide6) ------------------
# =========================================================================
class MacroApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Thiết lập cửa sổ không viền
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowTitle("Việt Tín Auto Sender V2025.04")
        # SỬA: Xóa kích thước cố định, chỉ đặt chiều rộng và để chiều cao tự động
        # self.resize(1200, 800)
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

        # --- Tạo các thành phần giao diện ---
        self._create_header_bar() 
        self._create_top_controls()
        self._create_data_macro_section()
        self._create_run_buttons()
        self._create_realtime_status_bar()
        self._create_disclaimer()

        # Áp dụng theme tối mặc định
        self.toggle_dark_mode(is_dark=True) # Gọi để áp dụng dark theme từ file qss

        # SỬA: Yêu cầu cửa sổ tự điều chỉnh kích thước sau khi đã tạo xong mọi thứ
        self.adjustSize()

    def paintEvent(self, event):
        """Vẽ nền bo góc cho cửa sổ chính."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Màu nền dựa trên theme
        bg_color = QColor("#282c34") if self.is_dark_mode else QColor("#f0f0f0")
        
        # Vẽ hình chữ nhật bo góc
        painter.setBrush(QBrush(bg_color))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(self.rect(), 15, 15)

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
        title_label = QLabel("Việt Tín Auto Sender V2025.04")
        title_label.setFont(QFont("Courier", 10, QFont.Bold))
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Nút Dark Mode
        self.dark_mode_btn = QPushButton("◐")
        self.dark_mode_btn.setFixedSize(30, 30)
        header_layout.addWidget(self.dark_mode_btn)

        # Nút Minimize
        self.minimize_btn = QPushButton("_")
        self.minimize_btn.setFixedSize(30, 30)
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
        g1_layout = QHBoxLayout(g1)
        g1_layout.addWidget(QLineEdit())
        g1_layout.addWidget(QPushButton("Browse"))
        g1_layout.addWidget(QPushButton("Mở ACSOFT"))

        # Group 2: CSV and Target Window
        g2 = QGroupBox("2) File CSV chứa dữ liệu / Cửa sổ mục tiêu")
        g2_layout = QVBoxLayout(g2)
        
        csv_layout = QHBoxLayout()
        csv_layout.addWidget(QLabel("File CSV:"))
        csv_layout.addWidget(QLineEdit())
        csv_layout.addWidget(QPushButton("Browse CSV"))
        g2_layout.addLayout(csv_layout)

        window_layout = QHBoxLayout()
        window_layout.addWidget(QLabel("Delimiter:"))
        window_layout.addWidget(QLineEdit(";"))
        window_layout.addWidget(QLabel("Cửa sổ:"))
        window_layout.addWidget(QComboBox())
        window_layout.addWidget(QPushButton("Làm mới"))
        g2_layout.addLayout(window_layout)

        # Group 3: Run Options
        g4 = QGroupBox("3) Tùy chọn chạy")
        g4_layout = QVBoxLayout(g4)
        
        speed_layout = QHBoxLayout()
        speed_layout.addWidget(QRadioButton("Tốc độ đã ghi"))
        speed_layout.addWidget(QRadioButton("Tốc độ cố định:"))
        speed_layout.addWidget(QSpinBox())
        speed_layout.addWidget(QLabel("ms"))
        speed_layout.addStretch()
        g4_layout.addLayout(speed_layout)

        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("Đợi giữa 2 dòng (1-20 giây):"))
        delay_layout.addWidget(QSpinBox())
        delay_layout.addStretch()
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
        record_controls.addWidget(QPushButton("⚫ Record Macro (5s chuẩn bị)"))
        record_controls.addStretch()
        record_controls.addWidget(QPushButton("Lưu Macro"))
        record_controls.addWidget(QPushButton("Mở Macro"))
        record_controls.addWidget(QPushButton("Clear Macro"))
        macro_layout.addLayout(record_controls)

        # -- Macro Table --
        self.tree_macro = QTableWidget()
        self.tree_macro.setColumnCount(3)
        self.tree_macro.setHorizontalHeaderLabels(["STT", "Loại", "Mô tả chi tiết"])
        self.tree_macro.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree_macro.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tree_macro.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        macro_layout.addWidget(self.tree_macro)

        # -- Manual Add Buttons --
        add_controls = QHBoxLayout()
        add_controls.addWidget(QPushButton("[+] Cột Dữ Liệu"))
        add_controls.addWidget(QPushButton("[+] Phím (Key)"))
        add_controls.addWidget(QPushButton("[+] Tổ Hợp (Combo)"))
        add_controls.addWidget(QPushButton("[+] Click Chuột"))
        add_controls.addStretch()
        macro_layout.addLayout(add_controls)

        # -- Edit/Delete Buttons --
        edit_controls = QHBoxLayout()
        edit_controls.addWidget(QPushButton("Sửa Dòng"))
        edit_controls.addWidget(QPushButton("Xóa Dòng"))
        edit_controls.addWidget(QPushButton("[+] Kết Thúc Dòng"))
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
        run_layout.addWidget(QPushButton("►CHẠY THỬ (1 DÒNG)"))
        run_layout.addWidget(QPushButton("►CHẠY TẤT CẢ"))
        run_layout.addWidget(QPushButton("STOP (ESC)"))
        run_layout.addStretch()
        self.lbl_status = QLabel("Chờ...")
        self.lbl_status.setFont(QFont("Courier", 10, QFont.Bold))
        run_layout.addWidget(self.lbl_status)

        self.main_layout.addWidget(run_widget)

    def _create_realtime_status_bar(self):
        """Tạo thanh trạng thái tọa độ real-time."""
        status_group = QGroupBox("Thông tin Tọa độ (Real-time)")
        status_layout = QHBoxLayout(status_group)

        self.lbl_realtime_status = QLabel("...")
        status_layout.addWidget(self.lbl_realtime_status)
        status_layout.addStretch()
        status_layout.addWidget(QCheckBox("Hiện/Ẩn"))

        # Thêm vào layout chính với margin
        frame = QFrame()
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(10, 0, 10, 10)
        frame_layout.addWidget(status_group)
        self.main_layout.addWidget(frame)

    def _create_disclaimer(self):
        """Tạo dòng ghi chú cuối cùng."""
        disclaimer_label = QLabel("Lưu ý: Ứng dụng BẮT BUỘC đưa ACSOFT lên foreground (phải focus). Nhấn phím ESC để hủy quá trình chạy.")
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

    def toggle_dark_mode(self, is_dark):
        """Chuyển đổi giao diện sáng/tối."""
        self.is_dark_mode = is_dark
        if is_dark:
            # Giao diện tối
            try:
                with open('dark.qss', 'r', encoding='utf-8') as f:
                    stylesheet = f.read()
            except FileNotFoundError:
                print("Lỗi: Không tìm thấy file 'dark.qss'. Sử dụng stylesheet mặc định.")
                stylesheet = ""
            self.dark_mode_btn.setText("☀")
        else:
            # Giao diện sáng: Xóa stylesheet hiện tại (hoặc có thể load light.qss nếu muốn)
            stylesheet = "" 
            self.dark_mode_btn.setText("◐")
        
        self.setStyleSheet(stylesheet)
        self.update() # Yêu cầu vẽ lại cửa sổ để áp dụng paintEvent

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Hiển thị HUD để kiểm tra
    # hud = RecordingHUD()
    # hud.show()

    # Hiển thị cửa sổ chính
    main_win = MacroApp()
    main_win.show()

    sys.exit(app.exec())