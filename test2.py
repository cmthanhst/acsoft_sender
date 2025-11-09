#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Liquid Glass main window using PySide6
Works cross-platform; on macOS it looks especially "glass-like".
"""

from PySide6.QtCore import (
    Qt, QTimer, QPointF, QRectF
)
from PySide6.QtGui import (
    QPainter, QColor, QBrush, QPen, QPainterPath,
    QLinearGradient, QRadialGradient, QFont
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QGraphicsDropShadowEffect
)
import sys
import math
import time


class LiquidGlassWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        # Enable mouse tracking for dynamic highlight
        self.setMouseTracking(True)
        self.cursor_pos = QPointF(self.width() / 2, self.height() / 2)
        self.animation_phase = 0.0

        # Soft shadow effect (outside)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(40)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(0, 0, 0, 120))
        self.setGraphicsEffect(shadow)

        # For smooth repainting animation
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.on_animate)
        self.timer.start(16)  # ~60 FPS

    def sizeHint(self):
        return self.parent().size()

    def on_animate(self):
        # advance gloss animation phase
        self.animation_phase += 0.02
        if self.animation_phase > math.pi * 2:
            self.animation_phase -= math.pi * 2
        self.update()

    def mouseMoveEvent(self, event):
        self.cursor_pos = event.position()
        self.update()

    def enterEvent(self, event):
        # when cursor enters, make sure we track; keep default

        super().enterEvent(event)

    def paintEvent(self, event):
        w = self.width()
        h = self.height()
        radius = 18  # corner radius

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Fill transparent background (the window itself is translucent)
        painter.fillRect(self.rect(), Qt.transparent)

        # Rounded rect path
        r = QRectF(8, 8, w - 16, h - 16)
        path = QPainterPath()
        path.addRoundedRect(r, radius, radius)

        # base glass: subtle frosted gradient (semi-transparent)
        grad = QLinearGradient(r.topLeft(), r.bottomRight())
        grad.setColorAt(0.0, QColor.fromRgbF(1.0, 1.0, 1.0, 0.12))
        grad.setColorAt(0.5, QColor.fromRgbF(0.95, 0.97, 1.0, 0.08))
        grad.setColorAt(1.0, QColor.fromRgbF(0.9, 0.92, 1.0, 0.06))
        painter.setBrush(QBrush(grad))
        painter.setPen(Qt.NoPen)
        painter.drawPath(path)

        # Inner soft shadow (to simulate depth)
        inner_shadow = QLinearGradient(r.topLeft(), r.bottomLeft())
        inner_shadow.setColorAt(0.0, QColor(255, 255, 255, 40))
        inner_shadow.setColorAt(1.0, QColor(0, 0, 0, 20))
        painter.setBrush(QBrush(inner_shadow))
        painter.setPen(Qt.NoPen)
        # draw slightly inset to create inner rim
        inner_r = QRectF(r.adjusted(2, 2, -2, -2))
        inner_path = QPainterPath()
        inner_path.addRoundedRect(inner_r, max(0, radius-2), max(0, radius-2))
        painter.drawPath(inner_path)

        # Frosted blur mimic: draw several semi-transparent white shapes
        # (gives perception of blur without real background blur)
        for i in range(3):
            alpha = int(8 + i * 6)
            fg = QRadialGradient(
                r.center().x() - (i*20), r.top() + 40 + i*10, 200 + i*40
            )
            fg.setColorAt(0.0, QColor(255, 255, 255, alpha))
            fg.setColorAt(1.0, QColor(255, 255, 255, 0))
            painter.setBrush(QBrush(fg))
            painter.drawPath(path)

        # Dynamic highlight / reflection following cursor (radial gradient)
        cx = self.cursor_pos.x()
        cy = self.cursor_pos.y()
        # limit to window rect
        cx = max(r.left()+20, min(r.right()-20, cx))
        cy = max(r.top()+20, min(r.bottom()-20, cy))

        # moving spotlight radius varies with animation phase
        pulse = 1.0 + 0.12 * math.sin(self.animation_phase * 2.0)
        spot_rad = min(w, h) * 0.4 * pulse

        spotlight = QRadialGradient(QPointF(cx, cy), spot_rad)
        spotlight.setColorAt(0.0, QColor(255, 255, 255, 90))
        spotlight.setColorAt(0.25, QColor(255, 255, 255, 42))
        spotlight.setColorAt(1.0, QColor(255, 255, 255, 0))
        painter.setBrush(spotlight)
        painter.setPen(Qt.NoPen)
        # Clip to rounded rect so highlight doesn't leak outside
        painter.save()
        painter.setClipPath(path)
        painter.drawRect(r)
        painter.restore()

        # Gloss strip (chrome-like) at top-left curved shape
        gloss = QLinearGradient(r.topLeft(), QPointF(r.right(), r.top() + 40))
        gloss.setColorAt(0.0, QColor(255, 255, 255, 180))
        gloss.setColorAt(0.6, QColor(255, 255, 255, 60))
        gloss.setColorAt(1.0, QColor(255, 255, 255, 0))
        painter.setBrush(QBrush(gloss))
        gloss_path = QPainterPath()
        gloss_rect = QRectF(r.left()+6, r.top()+6, r.width()-12, min(60, r.height()/3))
        gloss_path.addRoundedRect(gloss_rect, radius/1.5, radius/1.5)
        # intersect with main path
        painter.save()
        painter.setClipPath(path)
        painter.drawPath(gloss_path)
        painter.restore()

        # faint border
        pen = QPen(QColor(255, 255, 255, 25))
        pen.setWidthF(1.0)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawPath(path)

        # Optionally draw a title (subtle)
        painter.setPen(QColor(255, 255, 255, 150))
        f = QFont("Helvetica Neue", 12)
        f.setWeight(QFont.Weight.Light)
        painter.setFont(f)
        painter.drawText(r.left() + 20, r.top() + 32, "Liquid Glass — Main Window")

        painter.end()


class LiquidMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Liquid Glass Window")
        self.setWindowFlag(Qt.FramelessWindowHint)  # no native titlebar
        self.setAttribute(Qt.WA_TranslucentBackground)  # allow transparent parts
        self.setMouseTracking(True)

        # Size and center
        self.resize(860, 520)
        self.center_on_screen()

        # Main glass widget
        content = LiquidGlassWidget(self)
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        content_container = QWidget()
        content_container.setLayout(layout)
        layout.addWidget(content)
        self.setCentralWidget(content_container)

        # Allow moving the window by dragging anywhere on the glass
        self._mouse_drag_pos = None

    def center_on_screen(self):
        screen = QApplication.primaryScreen().availableGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._mouse_drag_pos = event.globalPosition() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self._mouse_drag_pos is not None and event.buttons() & Qt.LeftButton:
            self.move((event.globalPosition() - self._mouse_drag_pos).toPoint())
            event.accept()

    def mouseReleaseEvent(self, event):
        self._mouse_drag_pos = None
        super().mouseReleaseEvent(event)


def main():
    app = QApplication(sys.argv)
    # On macOS it's nice to allow the window to appear with a slight vibrancy style in titlebar area:
    # but we are using a frameless translucent window, so the glass painting handles the look.
    w = LiquidMainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
