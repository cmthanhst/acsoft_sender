import sys
import os
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QCheckBox, QRadioButton,
    QComboBox, QSlider, QProgressBar, QGroupBox, QListWidget, QSpinBox
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon

class TestApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PySide6 Test App 🧪")
        self.setMinimumSize(QSize(600, 400))
        
        # 0. Theo dõi chủ đề hiện tại
        self.current_theme = "light"
        
        # Thiết lập widget trung tâm và layout chính
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 1. Nút Chuyển Chủ đề và Thanh Tiêu đề
        header_layout = QHBoxLayout()
        
        self.theme_button = QPushButton("Chuyển sang Chủ đề Tối 🌙")
        self.theme_button.clicked.connect(self.toggle_theme)
        header_layout.addWidget(self.theme_button)
        
        main_layout.addLayout(header_layout)
        
        # 2. Tạo GroupBox chứa tất cả các phần tử
        self.create_widgets(main_layout)
        
        # 3. Áp dụng chủ đề ban đầu
        self.apply_theme("light")

    def create_widgets(self, main_layout):
        """Tạo và sắp xếp các widget test vào layout chính."""
        
        # --- Phần 1: Các Widget Input Cơ bản ---
        input_group = QGroupBox("1. Input và Nút bấm")
        input_layout = QVBoxLayout(input_group)
        
        # QLabel
        label = QLabel("QLabel: Đây là văn bản tĩnh.")
        input_layout.addWidget(label)
        
        # QLineEdit
        line_edit = QLineEdit("QLineEdit: Nhập văn bản một dòng...")
        input_layout.addWidget(line_edit)
        
        # QTextEdit
        text_edit = QTextEdit()
        text_edit.setPlaceholderText("QTextEdit: Văn bản nhiều dòng...")
        text_edit.setMaximumHeight(80)
        input_layout.addWidget(text_edit)
        
        # QPushButton
        button_layout = QHBoxLayout()
        button1 = QPushButton("Nút Bấm 1")
        button2 = QPushButton("Nút Tắt (Disabled)")
        button2.setEnabled(False)
        button_layout.addWidget(button1)
        button_layout.addWidget(button2)
        input_layout.addLayout(button_layout)
        
        main_layout.addWidget(input_group)
        
        # --- Phần 2: Các Widget Lựa chọn và Điều khiển ---
        control_group = QGroupBox("2. Lựa chọn và Điều khiển")
        control_layout = QHBoxLayout(control_group)
        
        # QCheckBox và QRadioButton
        check_radio_layout = QVBoxLayout()
        checkbox = QCheckBox("QCheckBox")
        radio1 = QRadioButton("QRadioButton 1")
        radio2 = QRadioButton("QRadioButton 2 (Checked)")
        radio2.setChecked(True)
        check_radio_layout.addWidget(checkbox)
        check_radio_layout.addWidget(radio1)
        check_radio_layout.addWidget(radio2)
        control_layout.addLayout(check_radio_layout)
        
        # QComboBox
        combo_layout = QVBoxLayout()
        combo_label = QLabel("QComboBox:")
        combobox = QComboBox()
        combobox.addItems(["Mục 1", "Mục 2", "Mục 3 dài hơn"])
        combo_layout.addWidget(combo_label)
        combo_layout.addWidget(combobox)
        control_layout.addLayout(combo_layout)

        # QSpinBox
        spin_layout = QVBoxLayout()
        spin_label = QLabel("QSpinBox:")
        spinbox = QSpinBox()
        spinbox.setRange(0, 100)
        spinbox.setValue(42)
        spin_layout.addWidget(spin_label)
        spin_layout.addWidget(spinbox)
        control_layout.addLayout(spin_layout)

        main_layout.addWidget(control_group)

        # --- Phần 3: QSlider, QProgressBar và QListWidget ---
        extra_group = QGroupBox("3. Thanh trượt và Danh sách")
        extra_layout = QVBoxLayout(extra_group)
        
        # QSlider
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(0, 100)
        slider.setValue(60)
        extra_layout.addWidget(QLabel("QSlider (60/100):"))
        extra_layout.addWidget(slider)
        
        # QProgressBar
        progress_bar = QProgressBar()
        progress_bar.setValue(75)
        extra_layout.addWidget(QLabel("QProgressBar (75%):"))
        extra_layout.addWidget(progress_bar)
        
        # QListWidget
        list_widget = QListWidget()
        list_widget.addItems(["Mục Danh sách 1", "Mục Danh sách 2 (Đã chọn)", "Mục Danh sách 3"])
        list_widget.setCurrentRow(1)
        list_widget.setMaximumHeight(80)
        extra_layout.addWidget(QLabel("QListWidget:"))
        extra_layout.addWidget(list_widget)
        
        main_layout.addWidget(extra_group)
        main_layout.addStretch() # Đẩy các widget lên trên

    def load_qss(self, filename):
        """Đọc và trả về nội dung tập tin QSS."""
        # SỬA: Sử dụng os.path.join để tạo đường dẫn an toàn và linh hoạt hơn
        # Điều này đảm bảo nó hoạt động đúng trên các hệ điều hành khác nhau.
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, filename)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read()
        except FileNotFoundError:
            print(f"Lỗi: Không tìm thấy tập tin style '{file_path}'.")
            return ""

    def apply_theme(self, theme):
        """Áp dụng chủ đề (light/dark) bằng cách tải QSS."""
        style = self.load_qss(f"{theme}.qss")
        if style:
            app.setStyleSheet(style)
        self.current_theme = theme
        
        # Cập nhật văn bản nút
        if theme == "light":
            self.theme_button.setText("Chuyển sang Chủ đề Tối 🌙")
        else:
            self.theme_button.setText("Chuyển sang Chủ đề Sáng ☀️")

    def toggle_theme(self):
        """Chuyển đổi giữa chủ đề sáng và tối."""
        if self.current_theme == "light":
            self.apply_theme("dark")
        else:
            self.apply_theme("light")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TestApp()
    window.show()
    sys.exit(app.exec())