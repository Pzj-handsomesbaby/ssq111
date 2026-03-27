# -*- coding: utf-8 -*-
"""UI baseline helpers for consistent fonts and window sizes."""

from PyQt5.QtGui import QFont

DEFAULT_FONT_FAMILY = "Microsoft YaHei"
FALLBACK_FONT_FAMILY = "SimHei"
DEFAULT_FONT_SIZE = 10

QT_MAIN_SIZE = (1600, 950)
QT_CHILD_SIZE = (1280, 820)
QT_MIN_SIZE = (1080, 720)

TK_MAIN_GEOMETRY = "1280x820"
TK_MIN_SIZE = (1080, 720)


def apply_qt_app_style(app):
    """Set a conservative global font and lightweight widget spacing."""
    app.setFont(QFont(DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE))
    app.setStyleSheet(
        "QWidget { font-family: 'Microsoft YaHei', 'SimHei', sans-serif; font-size: 10pt; }"
    )


def apply_qt_window_baseline(window, size=QT_CHILD_SIZE, min_size=QT_MIN_SIZE):
    """Apply a consistent default/minimum size to Qt windows."""
    window.resize(*size)
    window.setMinimumSize(*min_size)


def apply_tk_window_baseline(root, geometry=TK_MAIN_GEOMETRY, min_size=TK_MIN_SIZE):
    """Apply a consistent default/minimum size to Tk windows."""
    root.geometry(geometry)
    root.minsize(*min_size)

