# liquid_gloss_ripple_adaptive.py
# Requirements: PySide6, Pillow
# Auto-detect: screen size, Windows theme (dark/light), wallpaper
# Features: Mica/Acrylic/None toggle, Gloss Swipe (wallpaper hue), Intensity/Blur sliders,
# Chrome-like ripple on click, auto-tuned sizes based on detected screen.

import sys
import ctypes
from ctypes import wintypes, byref, c_int
import time
import colorsys
import winreg
from PIL import Image
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QHBoxLayout,
    QComboBox, QSlider, QGridLayout
)
from PySide6.QtGui import QColor, QPainter, QFont, QEnterEvent, QMouseEvent, QRadialGradient
from PySide6.QtCore import Qt, QTimer, QPointF

# ---------------- Win32 helpers / DWM ----------------
user32 = ctypes.windll.user32
dwmapi = ctypes.windll.dwmapi

SPI_GETDESKWALLPAPER = 0x0073
MAX_PATH = 260

DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_SYSTEMBACKDROP_TYPE = 38
DWMSBT_NONE = 1
DWMSBT_MAINWINDOW = 2

class ACCENTPOLICY(ctypes.Structure):
    _fields_ = [
        ("AccentState", ctypes.c_int),
        ("AccentFlags", ctypes.c_int),
        ("GradientColor", ctypes.c_uint),
        ("AnimationId", ctypes.c_int)
    ]

class WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
    _fields_ = [
        ("Attribute", ctypes.c_int),
        ("Data", ctypes.c_void_p),
        ("SizeOfData", ctypes.c_size_t)
    ]

try:
    SetWindowCompositionAttribute = user32.SetWindowCompositionAttribute
except AttributeError:
    SetWindowCompositionAttribute = None

def _argb(a, r, g, b):
    return (a << 24) | (r << 16) | (g << 8) | b

def enable_mica(hwnd: int, prefer_dark: bool) -> bool:
    try:
        dark = c_int(1 if prefer_dark else 0)
        dwmapi.DwmSetWindowAttribute(wintypes.HWND(hwnd),
                                     wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE),
                                     byref(dark),
                                     ctypes.sizeof(dark))
    except Exception:
        pass
    try:
        backdrop = c_int(DWMSBT_MAINWINDOW)
        res = dwmapi.DwmSetWindowAttribute(wintypes.HWND(hwnd),
                                           wintypes.DWORD(DWMWA_SYSTEMBACKDROP_TYPE),
                                           byref(backdrop),
                                           ctypes.sizeof(backdrop))
        return res == 0
    except Exception:
        return False

def disable_mica(hwnd: int):
    try:
        backdrop = c_int(DWMSBT_NONE)
        dwmapi.DwmSetWindowAttribute(wintypes.HWND(hwnd),
                                     wintypes.DWORD(DWMWA_SYSTEMBACKDROP_TYPE),
                                     byref(backdrop),
                                     ctypes.sizeof(backdrop))
    except Exception:
        pass

def enable_acrylic(hwnd: int, tint=(200,255,255,255)) -> bool:
    if SetWindowCompositionAttribute is None:
        return False
    try:
        accent = ACCENTPOLICY()
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
        accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.AccentFlags = 2
        a,r,g,b = tint
        accent.GradientColor = _argb(a, r, g, b)
        accent.AnimationId = 0
        data = WINDOWCOMPOSITIONATTRIBDATA()
        data.Attribute = 19  # WCA_ACCENT_POLICY
        data.SizeOfData = ctypes.sizeof(accent)
        data.Data = ctypes.cast(ctypes.pointer(accent), ctypes.c_void_p)
        res = SetWindowCompositionAttribute(wintypes.HWND(hwnd), ctypes.byref(data))
        return bool(res)
    except Exception:
        return False

def disable_acrylic(hwnd: int):
    if SetWindowCompositionAttribute is None:
        return
    try:
        accent = ACCENTPOLICY()
        accent.AccentState = 0
        accent.AccentFlags = 0
        accent.GradientColor = 0
        accent.AnimationId = 0
        data = WINDOWCOMPOSITIONATTRIBDATA()
        data.Attribute = 19
        data.SizeOfData = ctypes.sizeof(accent)
        data.Data = ctypes.cast(ctypes.pointer(accent), ctypes.c_void_p)
        SetWindowCompositionAttribute(wintypes.HWND(hwnd), ctypes.byref(data))
    except Exception:
        pass

# ---------------- Wallpaper sampling ----------------
def get_wallpaper_path():
    buf = ctypes.create_unicode_buffer(MAX_PATH)
    res = user32.SystemParametersInfoW(SPI_GETDESKWALLPAPER, MAX_PATH, buf, 0)
    if res:
        path = buf.value
        return path if path else None
    return None

def average_image_color(path, thumb=(128,128)):
    try:
        img = Image.open(path).convert("RGB")
        img.thumbnail(thumb)
        pixels = list(img.getdata())
        if not pixels:
            return (180,180,180)
        r = sum(p[0] for p in pixels)//len(pixels)
        g = sum(p[1] for p in pixels)//len(pixels)
        b = sum(p[2] for p in pixels)//len(pixels)
        return (r,g,b)
    except Exception:
        return (180,180,180)

def rgb_to_hsv(r,g,b):
    return colorsys.rgb_to_hsv(r/255.0,g/255.0,b/255.0)

def hsv_to_rgb(h,s,v):
    rr,gg,bb = colorsys.hsv_to_rgb(h,s,v)
    return int(rr*255), int(gg*255), int(bb*255)

# ---------------- Windows theme detection ----------------
def is_windows_dark_mode():
    try:
        # Registry key: HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        # 0 = dark, 1 = light
        return val == 0
    except Exception:
        # fallback assume dark
        return True

# ---------------- Ripple class ----------------
class Ripple:
    def __init__(self, center: QPointF, max_radius: float, duration: float = 0.65, color=(255,255,255)):
        self.center = QPointF(center)
        self.max_radius = max_radius
        self.duration = duration
        self.start_time = time.time()
        self.color = color
    def progress(self):
        t = (time.time() - self.start_time) / self.duration
        return max(0.0, min(1.0, t))
    def alive(self):
        return self.progress() < 1.0

# ---------------- UI ----------------
class LiquidWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Liquid Glass — Adaptive (auto-res, theme, wallpaper)")
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        # We'll choose a default size relative to primary screen
        screen = QApplication.primaryScreen()
        s = screen.size()
        # window initial size: 60% of smaller screen dimension
        base_w = int(s.width() * 0.6)
        base_h = int(s.height() * 0.55)
        self.resize(max(800, base_w), max(520, base_h))

        main = QVBoxLayout(self)
        main.setContentsMargins(18,18,18,18)

        # panel
        panel = QWidget(self)
        panel.setObjectName("panel")
        p_layout = QVBoxLayout(panel)
        p_layout.setContentsMargins(22,18,22,18)
        p_layout.setSpacing(10)

        self._is_dark = is_windows_dark_mode()

        # Title / subtitle adapt to theme
        title = QLabel("Liquid Glass · Gloss Swipe + Ripple (adaptive)")
        title.setFont(QFont("Segoe UI Variable", 14, QFont.Bold))
        title.setStyleSheet("color: rgba(255,255,255,0.95);" if self._is_dark else "color: rgba(0,0,0,0.95);")

        subtitle = QLabel("Auto-detect resolution & theme. Click to create Chrome-like ripple.")
        subtitle.setFont(QFont("Segoe UI", 10))
        subtitle.setStyleSheet("color: rgba(255,255,255,0.8);" if self._is_dark else "color: rgba(0,0,0,0.7);")

        # controls
        ctrl_row = QHBoxLayout()
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Mica", "Acrylic", "None"])
        self.mode_combo.setFixedWidth(140)
        combo_bg = "rgba(255,255,255,10%)" if self._is_dark else "rgba(0,0,0,6%)"
        combo_border = "rgba(255,255,255,14%)" if self._is_dark else "rgba(0,0,0,8%)"
        self.mode_combo.setStyleSheet(f"QComboBox{{padding:6px;border-radius:8px;background:{combo_bg};color:#fff;border:1px solid {combo_border};}}")

        refresh_btn = QPushButton("Refresh Wallpaper")
        refresh_btn.setFixedHeight(34)
        refresh_btn.setCursor(Qt.PointingHandCursor)
        refresh_btn.setStyleSheet(f"QPushButton{{border-radius:8px;padding:6px;background:{combo_bg};color:inherit;border:1px solid {combo_border};}} QPushButton:hover{{opacity:0.9}}")

        ctrl_row.addWidget(self.mode_combo)
        ctrl_row.addWidget(refresh_btn)
        ctrl_row.addStretch()

        # sliders
        sliders = QWidget()
        sgrid = QGridLayout(sliders)
        sgrid.setContentsMargins(0,0,0,0)
        sgrid.setHorizontalSpacing(12)
        sgrid.setVerticalSpacing(6)

        text_color = "rgba(255,255,255,0.9)" if self._is_dark else "rgba(0,0,0,0.85)"
        self.int_label = QLabel("Intensity: 70")
        self.int_label.setStyleSheet(f"color: {text_color}")
        self.int_slider = QSlider(Qt.Horizontal)
        self.int_slider.setRange(0,100)
        self.int_slider.setValue(70)

        self.blur_label = QLabel("Blur Radius (sim): 18")
        self.blur_label.setStyleSheet(f"color: {text_color}")
        self.blur_slider = QSlider(Qt.Horizontal)
        self.blur_slider.setRange(0,50)
        self.blur_slider.setValue(18)

        self.ripple_label = QLabel("Ripple Power: 1.0")
        self.ripple_label.setStyleSheet(f"color: {text_color}")
        self.ripple_slider = QSlider(Qt.Horizontal)
        self.ripple_slider.setRange(50,200)
        self.ripple_slider.setValue(100)

        sgrid.addWidget(self.int_label, 0, 0)
        sgrid.addWidget(self.int_slider, 0, 1)
        sgrid.addWidget(self.blur_label, 1, 0)
        sgrid.addWidget(self.blur_slider, 1, 1)
        sgrid.addWidget(self.ripple_label, 2, 0)
        sgrid.addWidget(self.ripple_slider, 2, 1)

        # buttons
        btn_row = QHBoxLayout()
        b1 = QPushButton("Primary")
        b2 = QPushButton("Secondary")
        for b in (b1,b2):
            b.setCursor(Qt.PointingHandCursor)
            b.setFixedHeight(42)
            btn_bg = "rgba(255,255,255,10%)" if self._is_dark else "rgba(0,0,0,6%)"
            b.setStyleSheet(f"QPushButton{{background:{btn_bg};color:inherit;border-radius:10px;padding:8px 14px;border:1px solid {combo_border};}} QPushButton:hover{{opacity:0.95}}")

        btn_row.addWidget(b1)
        btn_row.addWidget(b2)
        btn_row.addStretch()

        p_layout.addWidget(title)
        p_layout.addWidget(subtitle)
        p_layout.addLayout(ctrl_row)
        p_layout.addWidget(sliders)
        p_layout.addSpacing(6)
        p_layout.addLayout(btn_row)
        p_layout.addStretch()

        main.addWidget(panel)

        # panel style - dark/light adapt
        if self._is_dark:
            panel_bg = "qlineargradient(x1:0 y1:0 x2:1 y2:1, stop:0 rgba(20,20,20,88%), stop:1 rgba(18,18,18,82%))"
            border_col = "rgba(255,255,255,6%)"
        else:
            panel_bg = "qlineargradient(x1:0 y1:0 x2:1 y2:1, stop:0 rgba(255,255,255,86%), stop:1 rgba(255,255,255,82%))"
            border_col = "rgba(0,0,0,6%)"

        self.setStyleSheet(f"""
            #panel {{
                background: {panel_bg};
                border-radius: 18px;
                border: 1px solid {border_col};
            }}
        """)

        # gloss state
        self._mouse_target = QPointF(self.width()/2, self.height()/3)
        self._mouse_pos = QPointF(self._mouse_target)
        self._gloss_timer = QTimer(self)
        self._gloss_timer.setInterval(16)
        self._gloss_timer.timeout.connect(self._update_gloss_pos)
        self._gloss_timer.start()

        # ripples list
        self._ripples = []

        # wallpaper color
        self._wallpaper_rgb = (180,180,180)
        self._wallpaper_hsv = (0.0, 0.0, 0.6)
        self._read_wallpaper_color()

        # connect signals
        self.mode_combo.currentTextChanged.connect(self._on_mode_changed)
        refresh_btn.clicked.connect(self._on_refresh_wallpaper)
        self.int_slider.valueChanged.connect(self._on_int_changed)
        self.blur_slider.valueChanged.connect(self._on_blur_changed)
        self.ripple_slider.valueChanged.connect(self._on_ripple_changed)

        self._last_mode = None
        self._drag_pos = None

    # wallpaper
    def _read_wallpaper_color(self):
        path = get_wallpaper_path()
        if not path:
            self._wallpaper_rgb = (180,180,180)
        else:
            self._wallpaper_rgb = average_image_color(path)
        h,s,v = rgb_to_hsv(*self._wallpaper_rgb)
        s = min(1.0, max(0.05, s * (1.08 if self._is_dark else 1.02)))
        v = min(1.0, max(0.18 if self._is_dark else 0.25, v * (1.05 if self._is_dark else 1.02)))
        self._wallpaper_hsv = (h,s,v)
        self.update()

    def _on_refresh_wallpaper(self):
        self._read_wallpaper_color()

    def _on_int_changed(self, v):
        self.int_label.setText(f"Intensity: {v}")
        if self._last_mode == "Acrylic":
            a = int(v * 2.55)
            r,g,b = self._wallpaper_rgb
            enable_acrylic(int(self.winId()), tint=(a, r, g, b))

    def _on_blur_changed(self, v):
        self.blur_label.setText(f"Blur Radius (sim): {v}")
        self.update()

    def _on_ripple_changed(self, v):
        self.ripple_label.setText(f"Ripple Power: {v/100.0:.2f}")

    def showEvent(self, ev):
        super().showEvent(ev)
        self._apply_backdrop(self.mode_combo.currentText())

    def _on_mode_changed(self, text):
        self._apply_backdrop(text)

    def _apply_backdrop(self, text):
        hwnd = int(self.winId())
        disable_acrylic(hwnd)
        disable_mica(hwnd)
        if text == "Mica":
            ok = enable_mica(hwnd, prefer_dark=self._is_dark)
            if not ok:
                a = int(self.int_slider.value() * 2.55)
                enable_acrylic(hwnd, tint=(a, *self._wallpaper_rgb))
        elif text == "Acrylic":
            a = int(self.int_slider.value() * 2.55)
            enable_acrylic(hwnd, tint=(a, *self._wallpaper_rgb))
        else:
            pass
        self._last_mode = text
        self.update()

    def _update_gloss_pos(self):
        dx = self._mouse_target.x() - self._mouse_pos.x()
        dy = self._mouse_target.y() - self._mouse_pos.y()
        # easing factor slightly higher on larger screens
        screen = QApplication.primaryScreen().size()
        scale_factor = (screen.width() * screen.height()) / (1920*1080)
        lerp = 0.14 + min(0.18, 0.02 * (scale_factor ** 0.5))
        self._mouse_pos.setX(self._mouse_pos.x() + dx * lerp)
        self._mouse_pos.setY(self._mouse_pos.y() + dy * lerp)
        # prune ripples
        self._ripples = [r for r in self._ripples if r.alive()]
        self.update()

    def mouseMoveEvent(self, ev: QMouseEvent):
        if getattr(self, "_drag_pos", None) and ev.buttons() & Qt.LeftButton:
            self.move(ev.globalPosition().toPoint() - self._drag_pos)
            ev.accept()
            return
        pos = ev.position()
        self._mouse_target = QPointF(pos.x(), pos.y())
        super().mouseMoveEvent(ev)

    def mousePressEvent(self, ev: QMouseEvent):
        if ev.button() == Qt.LeftButton:
            pos = ev.position()
            # max radius depends on window diagonal and screen scale
            screen = QApplication.primaryScreen().size()
            diag = (self.width()**2 + self.height()**2) ** 0.5
            base_max = diag * 0.55
            power = self.ripple_slider.value() / 100.0
            max_r = base_max * (0.55 + 0.9 * power)
            h,s,v = self._wallpaper_hsv
            # adjust ripple tint for dark/light
            if self._is_dark:
                r_c,g_c,b_c = hsv_to_rgb(h, min(1.0, s*0.9), min(1.0, v*1.05))
            else:
                r_c,g_c,b_c = hsv_to_rgb(h, min(1.0, s*0.6), min(1.0, max(0.6, v*0.95)))
            ripple = Ripple(QPointF(pos.x(), pos.y()), max_radius=max_r, duration=0.72, color=(r_c,g_c,b_c))
            self._ripples.append(ripple)
            self._mouse_target = QPointF(pos.x(), pos.y())
            # allow dragging
            self._drag_pos = ev.globalPosition().toPoint() - self.frameGeometry().topLeft()
            ev.accept()
        super().mousePressEvent(ev)

    def mouseReleaseEvent(self, ev):
        self._drag_pos = None
        ev.accept()

    @staticmethod
    def ease_out_cubic(t):
        t -= 1
        return t*t*t + 1

    def paintEvent(self, ev):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Outer tint depends on theme
        if self._is_dark:
            outer = QColor(12,12,12,220)
            border_col = QColor(255,255,255,10)
        else:
            outer = QColor(255,255,255,240)
            border_col = QColor(0,0,0,10)

        painter.setBrush(outer)
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(self.rect(), 20, 20)

        # Gloss swipe
        mx = self._mouse_pos.x()
        my = self._mouse_pos.y()
        w,h = self.width(), self.height()

        # size scales with screen/window size
        screen = QApplication.primaryScreen().size()
        screen_scale = (screen.width()*screen.height()) / (1920*1080)
        base_gloss_w = max(200, w * (0.5 + 0.12 * (screen_scale ** 0.5)))
        base_gloss_h = max(40, h * (0.12 + 0.02 * (screen_scale ** 0.5)))
        blur_val = self.blur_slider.value()
        gloss_w = base_gloss_w * (1.0 + blur_val * 0.008)
        gloss_h = base_gloss_h * (0.45 + blur_val * 0.008)

        intensity = self.int_slider.value() / 100.0
        center_alpha = int(255 * (0.5 * intensity + 0.15))
        mid_alpha = int(255 * (0.32 * intensity + 0.10))
        outer_alpha = int(255 * (0.12 * intensity + 0.02))

        h_w, s_w, v_w = self._wallpaper_hsv
        s_n = min(1.0, s_w * (1.05 if self._is_dark else 1.02))
        v_n = min(1.0, v_w * (1.06 if self._is_dark else 1.02))
        r_c,g_c,b_c = hsv_to_rgb(h_w, s_n, v_n)

        painter.save()
        painter.setCompositionMode(QPainter.CompositionMode_Plus)
        radial = QRadialGradient(QPointF(mx, my), gloss_w * 0.9)
        radial.setColorAt(0.0, QColor(r_c, g_c, b_c, center_alpha))
        radial.setColorAt(0.22, QColor(r_c, g_c, b_c, mid_alpha))
        radial.setColorAt(0.6, QColor(r_c, g_c, b_c, outer_alpha))
        radial.setColorAt(1.0, QColor(r_c, g_c, b_c, 0))
        painter.setBrush(radial)
        painter.setPen(Qt.NoPen)
        painter.translate(mx, my)
        dx = (self._mouse_target.x() - mx)
        angle = max(-30, min(30, dx / max(1, w) * 80))
        painter.rotate(angle)
        sx = 1.9 + blur_val * 0.01
        sy = 0.36 + blur_val * 0.003
        painter.scale(sx, sy)
        painter.drawEllipse(QPointF(0,0), gloss_w*0.5, gloss_h*0.5)
        painter.restore()

        # Ripples
        for r in list(self._ripples):
            p = r.progress()
            if p >= 1.0:
                continue
            eased = self.ease_out_cubic(p)
            radius = r.max_radius * eased
            alpha = int(220 * (1.0 - eased) * intensity)
            rc, gc, bc = r.color
            grad = QRadialGradient(r.center, radius)
            grad.setColorAt(0.0, QColor(255,255,255, int(alpha*0.9)))
            grad.setColorAt(0.2, QColor(rc, gc, bc, int(alpha*0.7)))
            grad.setColorAt(0.75, QColor(rc, gc, bc, int(alpha*0.18)))
            grad.setColorAt(1.0, QColor(rc, gc, bc, 0))
            painter.setCompositionMode(QPainter.CompositionMode_SourceOver)
            painter.setBrush(grad)
            painter.setPen(Qt.NoPen)
            painter.drawEllipse(r.center, radius, radius * 0.9)

        # subtle inner border
        painter.setPen(border_col)
        painter.setBrush(Qt.NoBrush)
        rect = self.rect().adjusted(1,1,-1,-1)
        painter.drawRoundedRect(rect, 18, 18)

# ---------------- Run ----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = LiquidWindow()
    win.show()
    sys.exit(app.exec())
