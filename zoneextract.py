"""
ZoneExtract — extract data from OCR'd PDF forms into a table, by zones and templates.
ZoneExtract — извлечение данных из распознанных PDF-бланков в таблицу по зонам и шаблонам.

https://github.com/demyanenkopavel-wq/ZoneExtract  ·  MIT License

Программа читает текстовый слой PDF внутри прямоугольных зон и раскладывает
текст по столбцам таблицы. Набор полей задаёт пользователь в шаблоне:
  • поля, заполняемые вручную (общие для всей пачки);
  • поля из PDF (у каждого — своя зона на странице, обработка, очистка и проверка);
  • вычисляемые поля (формула из других полей);
  • от 1 до 8 страниц на документ;
  • опциональное сопоставление значений со словарём.
Интерфейс: русский и английский (Вид → Язык).

Запуск:
    pip install PySide6 PyMuPDF pandas openpyxl
    python zoneextract.py
"""

from __future__ import annotations

import ast
import importlib.util
import os
import re
import sys
import json
import tempfile
import subprocess
from collections import Counter
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from difflib import SequenceMatcher
from functools import lru_cache
from typing import Optional

# Понятное сообщение вместо ImportError, если не установлены зависимости.
_MISSING = [pkg for mods, pkg in ((("PySide6",), "PySide6"), (("pymupdf", "fitz"), "PyMuPDF"),
                                  (("pandas",), "pandas"), (("openpyxl",), "openpyxl"))
            if not any(importlib.util.find_spec(m) for m in mods)]
if _MISSING:
    print("ZoneExtract: missing packages / не установлены пакеты: " + ", ".join(_MISSING))
    print("pip install " + " ".join(_MISSING))
    if sys.stdin is not None and sys.stdin.isatty():
        input("Enter…")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
# PySide6
# ─────────────────────────────────────────────────────────────────────────────

from PySide6.QtCore import (
    Qt, QThread, Signal, QObject, QRectF, QPointF, QTimer, QSize, QEvent,
)
from PySide6.QtGui import (
    QColor, QFont, QFontMetrics, QPainter, QPen, QBrush, QPixmap, QImage,
    QKeySequence, QCursor, QAction, QActionGroup, QIcon,
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QTabWidget,
    QLabel, QPushButton, QLineEdit, QComboBox, QCheckBox,
    QTreeWidget, QTreeWidgetItem, QHeaderView, QScrollArea,
    QFileDialog, QMessageBox, QDialog, QDialogButtonBox, QProgressDialog,
    QProgressBar, QStatusBar, QFrame, QGroupBox, QMenu,
    QAbstractItemView, QSpinBox, QDoubleSpinBox, QSlider,
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QGraphicsRectItem, QGraphicsItem, QListWidget, QListWidgetItem,
    QPlainTextEdit, QTextBrowser, QSplitter,
)

try:
    import pymupdf as fitz   # PyMuPDF ≥ 1.24
except ImportError:          # старые версии PyMuPDF
    import fitz
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# Язык интерфейса (русский / английский)
# ─────────────────────────────────────────────────────────────────────────────
# Исходные тексты в коде — русские. tr() возвращает английский перевод из
# словаря _EN (в конце файла), если выбран английский язык. Язык читается при
# запуске из ui_settings.json; если он не задан — берётся язык системы.
# Смена языка применяется после перезапуска программы.

LANGUAGES = {"ru": "Русский", "en": "English"}

# Папка программы: здесь лежат шаблоны и настройки. Для собранного exe
# (PyInstaller) — папка exe, иначе — папка скрипта.
APP_DIR = (os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, "frozen", False)
           else os.path.dirname(os.path.abspath(__file__)))


def _detect_lang() -> str:
    cfg = os.path.join(APP_DIR, "ui_settings.json")
    try:
        with open(cfg, "r", encoding="utf-8") as f:
            lang = json.load(f).get("lang", "")
        if lang in LANGUAGES:
            return lang
    except Exception:
        pass
    try:
        from PySide6.QtCore import QLocale
        return "ru" if QLocale.system().name().lower().startswith(("ru", "be", "uk", "kk")) else "en"
    except Exception:
        return "ru"


LANG = _detect_lang()

# Разделитель дроби в извлечённых числах: «12,5» (рус.) или «12.5» (англ.).
DEC_SEP = "," if LANG == "ru" else "."


def tr(_text: str, **kw) -> str:
    """Перевод строки интерфейса. Именованные аргументы подставляются в {скобки}."""
    s = _EN.get(_text, _text) if LANG == "en" else _text
    return s.format(**kw) if kw else s


# ─────────────────────────────────────────────────────────────────────────────
# Совместимость энумов PySide6 5.x / 6.x
# ─────────────────────────────────────────────────────────────────────────────

def _qenum(base, scoped: str, member: str):
    """Возвращает член перечисления Qt, работая и с «scoped»-энумами 6.x
    (base.<scoped>.<member>), и с «плоскими» энумами 5.x (base.<member>).

    Приоритет отдаётся scoped-варианту; проверка идёт через ``is not None``,
    поэтому члены со значением 0 (например ``NoFocus``) не теряются, как это
    было бы при связывании через ``or``.
    """
    cls = getattr(base, scoped, None)
    if cls is not None:
        v = getattr(cls, member, None)
        if v is not None:
            return v
    v = getattr(base, member, None)
    if v is not None:
        return v
    raise AttributeError(f"{getattr(base, '__name__', base)}.{scoped}.{member} не найден")


# Кнопки диалогов
_MB_YES    = _qenum(QMessageBox, "StandardButton", "Yes")
_MB_NO     = _qenum(QMessageBox, "StandardButton", "No")
_DB_OK     = _qenum(QDialogButtonBox, "StandardButton", "Ok")
_DB_CANCEL = _qenum(QDialogButtonBox, "StandardButton", "Cancel")
_DB_SAVE   = _qenum(QDialogButtonBox, "StandardButton", "Save")
_DLG_OK    = _qenum(QDialog, "DialogCode", "Accepted")
_WIN_MODAL = _qenum(Qt, "WindowModality", "WindowModal")

# Флаги Qt
_ALIGN_RIGHT   = _qenum(Qt, "AlignmentFlag", "AlignRight")
_ALIGN_CENTER  = _qenum(Qt, "AlignmentFlag", "AlignCenter")
_ALIGN_TOP     = _qenum(Qt, "AlignmentFlag", "AlignTop")
_ALIGN_LEFT    = _qenum(Qt, "AlignmentFlag", "AlignLeft")
_ALIGN_VCENTER = _qenum(Qt, "AlignmentFlag", "AlignVCenter")
_NO_FRAME      = _qenum(QFrame, "Shape", "NoFrame")
_HLINE         = _qenum(QFrame, "Shape", "HLine")
_VLINE         = _qenum(QFrame, "Shape", "VLine")
_ITEM_EDITABLE = _qenum(Qt, "ItemFlag", "ItemIsEditable")
_NO_FOCUS      = _qenum(Qt, "FocusPolicy", "NoFocus")
_CUSTOM_CTX    = _qenum(Qt, "ContextMenuPolicy", "CustomContextMenu")
_EXT_SEL       = _qenum(QAbstractItemView, "SelectionMode", "ExtendedSelection")
_SINGLE_SEL    = _qenum(QAbstractItemView, "SelectionMode", "SingleSelection")
_DND_INTERNAL  = _qenum(QAbstractItemView, "DragDropMode", "InternalMove")
_MOVE_ACTION   = _qenum(Qt, "DropAction", "MoveAction")
_USER_ROLE     = _qenum(Qt, "ItemDataRole", "UserRole")
_NO_EDIT       = _qenum(QAbstractItemView, "EditTrigger", "NoEditTriggers")
_SCROLL_CENTER = _qenum(QAbstractItemView, "ScrollHint", "PositionAtCenter")
_RESIZE_CONT   = _qenum(QHeaderView, "ResizeMode", "ResizeToContents")
_SCROLL_DRAG   = _qenum(QGraphicsView, "DragMode", "ScrollHandDrag")
_ANCHOR_MOUSE  = _qenum(QGraphicsView, "ViewportAnchor", "AnchorUnderMouse")
_ANTIALIASING  = _qenum(QPainter, "RenderHint", "Antialiasing")
_IMG_RGB888    = _qenum(QImage, "Format", "Format_RGB888")
_HORIZONTAL    = _qenum(Qt, "Orientation", "Horizontal")
_LEFT_BTN      = _qenum(Qt, "MouseButton", "LeftButton")
_BOLD          = _qenum(QFont, "Weight", "Bold")
_WHITE         = _qenum(Qt, "GlobalColor", "white")
_ARROW_CUR     = _qenum(Qt, "CursorShape", "ArrowCursor")
_SIZEALL_CUR   = _qenum(Qt, "CursorShape", "SizeAllCursor")
_SMOOTH        = _qenum(Qt, "TransformationMode", "SmoothTransformation")
_NO_PEN        = _qenum(Qt, "PenStyle", "NoPen")
_GI_MOVABLE    = _qenum(QGraphicsItem, "GraphicsItemFlag", "ItemIsMovable")
_GI_SELECTABLE = _qenum(QGraphicsItem, "GraphicsItemFlag", "ItemIsSelectable")
_GI_GEOMCHG    = _qenum(QGraphicsItem, "GraphicsItemFlag", "ItemSendsGeometryChanges")

# Клавиши (используются в редакторе автодополнения)
_KEY_DOWN = _qenum(Qt, "Key", "Key_Down")
_KEY_UP   = _qenum(Qt, "Key", "Key_Up")
_KEY_ESC  = _qenum(Qt, "Key", "Key_Escape")

# Стили пера для зон
_DASH_PEN  = _qenum(Qt, "PenStyle", "DashLine")
_SOLID_PEN = _qenum(Qt, "PenStyle", "SolidLine")


def _cursor(shape: str):
    return _qenum(Qt, "CursorShape", shape)


_CURSOR_MAP = {
    "tl": _cursor("SizeFDiagCursor"),
    "br": _cursor("SizeFDiagCursor"),
    "tr": _cursor("SizeBDiagCursor"),
    "bl": _cursor("SizeBDiagCursor"),
    "l":  _cursor("SizeHorCursor"),
    "r":  _cursor("SizeHorCursor"),
    "t":  _cursor("SizeVerCursor"),
    "b":  _cursor("SizeVerCursor"),
}


# ─────────────────────────────────────────────────────────────────────────────
# Защита от случайной смены значений колёсиком мыши
# ─────────────────────────────────────────────────────────────────────────────
# Списки и счётчики НЕ меняют значение колёсиком. Событие игнорируется и
# уходит родителю, поэтому окно (например, конструктор шаблона) прокручивается.

class NoWheelComboBox(QComboBox):
    def wheelEvent(self, event):
        event.ignore()


class NoWheelSpinBox(QSpinBox):
    def wheelEvent(self, event):
        event.ignore()


class NoWheelDoubleSpinBox(QDoubleSpinBox):
    def wheelEvent(self, event):
        event.ignore()

# ─────────────────────────────────────────────────────────────────────────────
# Константы
# ─────────────────────────────────────────────────────────────────────────────

VERSION = "2.0.0"

# Максимальное число страниц на один документ (для многостраничных бланков).
MAX_DOC_PAGES = 8

# Порог автозамены при сопоставлении по словарю (0.0–1.0).
AUTO_MATCH_THRESHOLD = 0.85

# Файл с пользовательскими шаблонами (рядом со скриптом)
_TEMPLATES_CFG = os.path.join(APP_DIR, "templates_config.json")

# ─────────────────────────────────────────────────────────────────────────────
# Темы оформления
# ─────────────────────────────────────────────────────────────────────────────
#
# Единый источник правды по цветам — датакласс Theme. Все стили (глобальный QSS
# и инлайн-стили нестандартных виджетов) читают активную тему T, поэтому смена
# темы применяется «на лету», без перезапуска.
#
# Тело документа (таблица, поля ввода, карточки) во всех темах остаётся светлым —
# это айдентика приложения и лучший фон для работы с бумагами. Меняются «хром»
# (тёмные панели, вкладки, меню, тулбары), акцентный цвет и шапка.

@dataclass(frozen=True)
class Theme:
    name: str
    # Светлое тело: окно, карточки, поля, таблица
    app_bg:         str   # фон окна и содержимого вкладок
    surface:        str   # карточки, поля ввода, списки, чётная строка таблицы
    surface_alt:    str   # нечётная строка таблицы (чередование)
    content_border: str   # рамки полей/карточек и линии сетки таблицы
    text:           str   # основной текст на светлом
    text_muted:     str   # приглушённый текст/подписи на светлом
    # «Хром»: панели, вкладки, меню, тулбары, статус-бар
    chrome:         str   # основной фон панелей
    chrome_deep:    str   # самый глубокий слой (меню, активная вкладка, полоса координат)
    chrome_raised:  str   # приподнятые контролы на хроме (кнопки тулбаров)
    chrome_border:  str   # рамки на хроме
    on_chrome:      str   # текст на хроме
    on_chrome_dim:  str   # приглушённый текст на хроме (подсказки, предпросмотр)
    # Акцент
    accent:         str
    accent_hi:      str   # ховер/нажатие
    accent_soft:    str   # мягкий фон выделения на светлом
    accent_bright:  str   # яркий акцент-текст (активная вкладка, выделение)
    # Шапка приложения
    header_from:        str
    header_to:          str
    header_ctl:         str  # фон комбобокса/кнопок в шапке
    header_ctl_border:  str
    header_ctl_hi:      str  # ховер кнопок шапки
    header_ctl_off:     str  # неактивная кнопка шапки
    header_ctl_off_fg:  str
    header_dim:         str  # подписи/версия в шапке
    # Семантика статусов
    success:            str
    success_hi:         str
    warning:            str
    danger:             str
    danger_hi:          str
    danger_deep:        str  # фон приглушённой опасной кнопки
    danger_deep_fg:     str
    danger_deep_border: str
    # Цвета предпросмотра текста зон
    ok_text:            str
    err_text:           str
    warn_text:          str
    # Состояния ячеек таблицы
    err_bg:             str
    err_fg:             str
    empty_bg:           str
    empty_fg:           str
    # Редактор зон
    canvas_bg:          str  # фон холста QGraphicsView
    preview_bg:         str  # фон строки предпросмотра
    placeholder:        str  # текст-заглушка сцены и разделитель страниц
    # Повторяющееся значение ключевого поля
    dup_bg:             str = "#ede9fe"
    dup_fg:             str = "#5b21b6"


THEMES: dict[str, Theme] = {
    "Индиго": Theme(
        name="Индиго",
        app_bg="#f8fafc", surface="#ffffff", surface_alt="#f1f5f9",
        content_border="#cbd5e1", text="#0f172a", text_muted="#64748b",
        chrome="#1e293b", chrome_deep="#0f172a", chrome_raised="#334155",
        chrome_border="#475569", on_chrome="#e2e8f0", on_chrome_dim="#cbd5e1",
        accent="#3b82f6", accent_hi="#2563eb", accent_soft="#dbeafe", accent_bright="#2563eb",
        header_from="#1e3a8a", header_to="#1d4ed8", header_ctl="#1e40af",
        header_ctl_border="#60a5fa", header_ctl_hi="#3b82f6",
        header_ctl_off="#1e3a8a", header_ctl_off_fg="#3b5bb5", header_dim="#bfdbfe",
        success="#059669", success_hi="#047857", warning="#d97706",
        danger="#dc2626", danger_hi="#b91c1c",
        danger_deep="#7f1d1d", danger_deep_fg="#fca5a5", danger_deep_border="#991b1b",
        ok_text="#4ade80", err_text="#f87171", warn_text="#fbbf24",
        err_bg="#fee2e2", err_fg="#991b1b", empty_bg="#fde68a", empty_fg="#92400e",
        canvas_bg="#e2e8f0", preview_bg="#0b1220", placeholder="#64748b",
    ),
    "Графит": Theme(
        name="Графит",
        app_bg="#fafafa", surface="#ffffff", surface_alt="#f4f4f5",
        content_border="#d4d4d8", text="#18181b", text_muted="#57575e",
        chrome="#27272a", chrome_deep="#18181b", chrome_raised="#3f3f46",
        chrome_border="#52525b", on_chrome="#e4e4e7", on_chrome_dim="#d4d4d8",
        accent="#0d9488", accent_hi="#0f766e", accent_soft="#ccfbf1", accent_bright="#0f766e",
        header_from="#134e4a", header_to="#0f766e", header_ctl="#115e59",
        header_ctl_border="#2dd4bf", header_ctl_hi="#14b8a6",
        header_ctl_off="#134e4a", header_ctl_off_fg="#4b8f89", header_dim="#99f6e4",
        success="#059669", success_hi="#047857", warning="#d97706",
        danger="#dc2626", danger_hi="#b91c1c",
        danger_deep="#7f1d1d", danger_deep_fg="#fca5a5", danger_deep_border="#991b1b",
        ok_text="#34d399", err_text="#f87171", warn_text="#fbbf24",
        err_bg="#fee2e2", err_fg="#991b1b", empty_bg="#fde68a", empty_fg="#92400e",
        canvas_bg="#e4e4e7", preview_bg="#0c0a09", placeholder="#71717a",
    ),
    "Изумруд": Theme(
        name="Изумруд",
        app_bg="#f6faf8", surface="#ffffff", surface_alt="#eef5f1",
        content_border="#cbd5cf", text="#0f231b", text_muted="#5b6f66",
        chrome="#14302a", chrome_deep="#0b201b", chrome_raised="#245247",
        chrome_border="#356b5d", on_chrome="#e6f2ee", on_chrome_dim="#c4ddd4",
        accent="#10b981", accent_hi="#059669", accent_soft="#d1fae5", accent_bright="#047857",
        header_from="#064e3b", header_to="#059669", header_ctl="#065f46",
        header_ctl_border="#34d399", header_ctl_hi="#10b981",
        header_ctl_off="#064e3b", header_ctl_off_fg="#3f7a68", header_dim="#a7f3d0",
        success="#059669", success_hi="#047857", warning="#d97706",
        danger="#dc2626", danger_hi="#b91c1c",
        danger_deep="#7f1d1d", danger_deep_fg="#fca5a5", danger_deep_border="#991b1b",
        ok_text="#4ade80", err_text="#f87171", warn_text="#fbbf24",
        err_bg="#fee2e2", err_fg="#991b1b", empty_bg="#fde68a", empty_fg="#92400e",
        canvas_bg="#dfe9e4", preview_bg="#07160f", placeholder="#5b6f66",
    ),
    "Аметист": Theme(
        name="Аметист",
        app_bg="#faf8fd", surface="#ffffff", surface_alt="#f4f0fb",
        content_border="#d6cee6", text="#1e1633", text_muted="#655a7d",
        chrome="#2a2140", chrome_deep="#1b1530", chrome_raised="#443661",
        chrome_border="#584878", on_chrome="#ece7f5", on_chrome_dim="#cfc4e4",
        accent="#8b5cf6", accent_hi="#7c3aed", accent_soft="#ede9fe", accent_bright="#7c3aed",
        header_from="#4c1d95", header_to="#7c3aed", header_ctl="#5b21b6",
        header_ctl_border="#a78bfa", header_ctl_hi="#8b5cf6",
        header_ctl_off="#4c1d95", header_ctl_off_fg="#6d5a9c", header_dim="#ddd6fe",
        success="#059669", success_hi="#047857", warning="#d97706",
        danger="#dc2626", danger_hi="#b91c1c",
        danger_deep="#7f1d1d", danger_deep_fg="#fca5a5", danger_deep_border="#991b1b",
        ok_text="#4ade80", err_text="#f87171", warn_text="#fbbf24",
        err_bg="#fee2e2", err_fg="#991b1b", empty_bg="#fde68a", empty_fg="#92400e",
        canvas_bg="#e6e0f0", preview_bg="#130d24", placeholder="#655a7d",
    ),
    "Светлая": Theme(
        name="Светлая",
        app_bg="#f1f5f9", surface="#ffffff", surface_alt="#f8fafc",
        content_border="#cbd5e1", text="#0f172a", text_muted="#64748b",
        chrome="#e2e8f0", chrome_deep="#f8fafc", chrome_raised="#f1f5f9",
        chrome_border="#cbd5e1", on_chrome="#334155", on_chrome_dim="#64748b",
        accent="#2563eb", accent_hi="#1d4ed8", accent_soft="#dbeafe", accent_bright="#1d4ed8",
        header_from="#1e40af", header_to="#2563eb", header_ctl="#1d4ed8",
        header_ctl_border="#60a5fa", header_ctl_hi="#3b82f6",
        header_ctl_off="#1e3a8a", header_ctl_off_fg="#93b4f0", header_dim="#bfdbfe",
        success="#059669", success_hi="#047857", warning="#d97706",
        danger="#dc2626", danger_hi="#b91c1c",
        danger_deep="#fecaca", danger_deep_fg="#b91c1c", danger_deep_border="#fca5a5",
        ok_text="#16a34a", err_text="#dc2626", warn_text="#d97706",
        err_bg="#fee2e2", err_fg="#991b1b", empty_bg="#fef3c7", empty_fg="#92400e",
        canvas_bg="#e2e8f0", preview_bg="#f8fafc", placeholder="#94a3b8",
    ),
}

_THEME_ORDER = list(THEMES.keys())
_DEFAULT_THEME = "Индиго"

# Файл пользовательских настроек интерфейса (рядом со скриптом)
_UI_SETTINGS_CFG = os.path.join(APP_DIR, "ui_settings.json")


def load_ui_settings() -> dict:
    try:
        with open(_UI_SETTINGS_CFG, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except Exception as e:
        _log_error(tr("не удалось прочитать настройки интерфейса"), e)
        return {}


def save_ui_settings(data: dict) -> bool:
    try:
        with open(_UI_SETTINGS_CFG, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        _log_error(tr("не удалось сохранить настройки интерфейса"), e)
        return False


# Активная тема. Переопределяется _apply_theme_tokens ниже.
T: Theme = THEMES[_DEFAULT_THEME]

# Устаревшие имена C_* сохранены для совместимости с кодом, читающим их
# «на месте» (например, при заливке ячеек таблицы). Значения выставляет
# _apply_theme_tokens при каждой смене темы.
C_BG = C_BG2 = C_BG3 = C_CARD = C_ACCENT = C_ACCENT2 = C_ACCENT_L = ""
C_SUCCESS = C_WARNING = C_DANGER = C_TEXT = C_TEXT2 = ""
C_ERR_BG = C_ERR_FG = C_EMPTY_BG = C_EMPTY_FG = C_ALT_ROW = ""


def _apply_theme_tokens(theme: Theme):
    """Делает тему активной и синхронизирует устаревшие константы C_*."""
    global T
    global C_BG, C_BG2, C_BG3, C_CARD, C_ACCENT, C_ACCENT2, C_ACCENT_L
    global C_SUCCESS, C_WARNING, C_DANGER, C_TEXT, C_TEXT2
    global C_ERR_BG, C_ERR_FG, C_EMPTY_BG, C_EMPTY_FG, C_ALT_ROW
    T          = theme
    C_BG       = theme.app_bg
    C_BG2      = theme.chrome
    C_BG3      = theme.content_border
    C_CARD     = theme.surface
    C_ACCENT   = theme.accent
    C_ACCENT2  = theme.accent_hi
    C_ACCENT_L = theme.accent_soft
    C_SUCCESS  = theme.success
    C_WARNING  = theme.warning
    C_DANGER   = theme.danger
    C_TEXT     = theme.text
    C_TEXT2    = theme.text_muted
    C_ERR_BG   = theme.err_bg
    C_ERR_FG   = theme.err_fg
    C_EMPTY_BG = theme.empty_bg
    C_EMPTY_FG = theme.empty_fg
    C_ALT_ROW  = theme.surface_alt


_apply_theme_tokens(THEMES[_DEFAULT_THEME])

# Видимый маркер «значение не извлечено». Ставится только на экран —
# в Excel и буфер обмена не попадает (см. _cell_text).
def missing_mark() -> str:
    return tr("∅ нет данных")

# Палитра для зон / столбцов (циклическая)
_ZONE_PALETTE = [
    (59, 130, 246), (16, 185, 129), (245, 158, 11), (239, 68, 68),
    (168, 85, 247), (236, 72, 153), (20, 184, 166), (234, 179, 8),
    (99, 102, 241), (217, 70, 239), (14, 165, 233), (132, 204, 22),
]


def _palette_color(i: int, alpha: int = 120) -> QColor:
    r, g, b = _ZONE_PALETTE[i % len(_ZONE_PALETTE)]
    return QColor(r, g, b, alpha)


def _default_zone(i: int) -> tuple:
    """Зона по умолчанию для i-го извлекаемого поля (в координатах при 300 DPI)."""
    y = 300 + (i % 9) * 300
    return (400, y, 1400, y + 200)


def _default_fallback_zone(primary: tuple) -> tuple:
    """Резервная зона по умолчанию: под основной, той же ширины и высоты."""
    x0, y0, x1, y1 = primary
    h = max(120, y1 - y0)
    ny0 = y1 + 60
    return (x0, ny0, x1, ny0 + h)

# ─────────────────────────────────────────────────────────────────────────────
# Глобальный стиль (QSS) — строится из активной темы
# ─────────────────────────────────────────────────────────────────────────────

def build_qss(t: "Theme" = None) -> str:
    """Собирает глобальную таблицу стилей для указанной темы (по умолчанию — активной)."""
    if t is None:
        t = T
    return f"""
QMainWindow, QDialog {{
    background-color: {t.app_bg};
}}
QWidget {{
    background-color: {t.app_bg};
    color: {t.text};
    font-family: 'Segoe UI', 'Inter', sans-serif;
    font-size: 10pt;
}}
QToolTip {{
    background: {t.chrome_deep};
    color: {t.on_chrome};
    border: 1px solid {t.chrome_border};
    border-radius: 4px;
    padding: 4px 8px;
}}
QTabWidget::pane {{
    border: 1px solid {t.chrome_border};
    border-top: none;
    background: {t.app_bg};
}}
QTabBar {{ background: {t.chrome}; }}
QTabBar::tab {{
    background: {t.chrome};
    color: {t.on_chrome_dim};
    padding: 10px 22px;
    border: none;
    border-bottom: 3px solid transparent;
    font-size: 10pt;
    font-weight: 500;
    min-width: 130px;
}}
QTabBar::tab:selected {{
    background: {t.chrome_deep};
    color: {t.accent_bright};
    border-bottom: 3px solid {t.accent};
    font-weight: bold;
}}
QTabBar::tab:hover:!selected {{
    background: {t.chrome_raised};
    color: {t.on_chrome};
}}
QPushButton {{
    background-color: {t.chrome_raised};
    color: {t.on_chrome};
    border: 1px solid {t.chrome_border};
    padding: 7px 16px;
    border-radius: 6px;
    font-size: 9pt;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {t.accent};
    color: white;
    border-color: {t.accent};
}}
QPushButton:pressed  {{ background-color: {t.accent_hi}; color: white; }}
QPushButton:disabled {{ background-color: {t.chrome}; color: {t.chrome_border}; border-color: {t.chrome_border}; }}
QPushButton#primary  {{ background-color: {t.accent}; color: white; font-weight: bold; font-size: 10pt; padding: 9px 20px; border: none; }}
QPushButton#primary:hover {{ background-color: {t.accent_hi}; }}
QPushButton#success  {{ background-color: {t.success}; color: white; font-weight: bold; font-size: 11pt; padding: 10px 24px; border: none; }}
QPushButton#success:hover {{ background-color: {t.success_hi}; }}
QPushButton#danger   {{ background-color: {t.danger}; color: white; border: none; font-weight: bold; }}
QPushButton#danger:hover {{ background-color: {t.danger_hi}; color: white; }}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
    background-color: {t.surface};
    color: {t.text};
    border: 1px solid {t.content_border};
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 10pt;
    selection-background-color: {t.accent};
    selection-color: white;
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {{ border: 1px solid {t.accent}; }}
QLineEdit:hover, QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover {{ border-color: {t.accent}; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{
    background: {t.surface};
    selection-background-color: {t.accent_soft};
    selection-color: {t.accent_bright};
    border: 1px solid {t.content_border};
    outline: none;
}}
QCheckBox {{ color: {t.text}; spacing: 8px; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border: 1px solid {t.content_border}; border-radius: 4px; background: {t.surface};
}}
QCheckBox::indicator:checked {{ background: {t.accent}; border-color: {t.accent}; }}
QTreeWidget {{
    background-color: {t.surface};
    alternate-background-color: {t.app_bg};
    color: {t.text};
    border: none;
    gridline-color: {t.content_border};
    outline: none;
}}
QTreeWidget::item {{ padding: 2px 0; }}
QTreeWidget::item:selected {{ background-color: {t.accent_soft}; color: {t.accent_bright}; }}
QHeaderView::section {{
    background-color: {t.chrome};
    color: {t.on_chrome_dim};
    padding: 6px 8px;
    border: none;
    border-right: 1px solid {t.chrome_border};
    border-bottom: 1px solid {t.chrome_border};
    font-weight: bold;
    font-size: 9pt;
}}
QScrollBar:vertical {{ background: {t.chrome}; width: 9px; border-radius: 4px; margin: 0; }}
QScrollBar::handle:vertical {{ background: {t.chrome_border}; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: {t.accent}; }}
QScrollBar:horizontal {{ background: {t.chrome}; height: 9px; border-radius: 4px; }}
QScrollBar::handle:horizontal {{ background: {t.chrome_border}; border-radius: 4px; min-width: 30px; }}
QScrollBar::handle:horizontal:hover {{ background: {t.accent}; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: none; }}
QProgressBar {{
    background: {t.content_border}; border: none; border-radius: 4px; height: 6px; text-align: center;
}}
QProgressBar::chunk {{ background: {t.accent}; border-radius: 4px; }}
QGroupBox {{
    color: {t.accent}; border: 1px solid {t.content_border}; border-radius: 8px;
    margin-top: 12px; padding: 8px; font-weight: bold; background: {t.surface};
}}
QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 6px; background: {t.surface}; }}
QStatusBar {{
    background: {t.chrome}; color: {t.on_chrome_dim};
    border-top: 2px solid {t.accent}; font-size: 9pt; padding: 2px 8px;
}}
QMenuBar {{ background: {t.chrome_deep}; color: {t.on_chrome}; border-bottom: 1px solid {t.chrome_border}; }}
QMenuBar::item {{ padding: 5px 12px; background: transparent; }}
QMenuBar::item:selected {{ background: {t.chrome_raised}; color: {t.accent_bright}; }}
QMenu {{
    background: {t.chrome}; color: {t.on_chrome};
    border: 1px solid {t.chrome_border}; border-radius: 6px; padding: 4px;
}}
QMenu::item {{ padding: 6px 22px; border-radius: 4px; }}
QMenu::item:selected {{ background: {t.chrome_raised}; color: {t.accent_bright}; }}
QMenu::item:checked {{ color: {t.accent_bright}; }}
QMenu::separator {{ height: 1px; background: {t.chrome_border}; margin: 4px 8px; }}
QFrame[frameShape="4"], QFrame[frameShape="5"] {{ color: {t.content_border}; }}
QListWidget {{
    background: {t.surface}; color: {t.text};
    border: 1px solid {t.content_border}; border-radius: 6px;
    outline: none;
}}
QListWidget::item:selected {{ background: {t.accent_soft}; color: {t.accent_bright}; }}
QSpinBox {{ padding-right: 24px; }}
QSpinBox::up-button {{
    subcontrol-origin: border; subcontrol-position: top right;
    width: 22px; height: 15px; border-left: 1px solid {t.content_border};
    border-top-right-radius: 5px; background: {t.accent};
}}
QSpinBox::down-button {{
    subcontrol-origin: border; subcontrol-position: bottom right;
    width: 22px; height: 15px; border-left: 1px solid {t.content_border};
    border-bottom-right-radius: 5px; background: {t.accent};
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {t.accent_hi}; }}
QSpinBox::up-button:pressed, QSpinBox::down-button:pressed {{ background: {t.accent_hi}; }}
QSpinBox::up-arrow {{
    image: none; width: 0; height: 0;
    border-left: 5px solid transparent; border-right: 5px solid transparent;
    border-bottom: 7px solid white;
}}
QSpinBox::down-arrow {{
    image: none; width: 0; height: 0;
    border-left: 5px solid transparent; border-right: 5px solid transparent;
    border-top: 7px solid white;
}}
QSpinBox::up-arrow:disabled {{ border-bottom-color: {t.on_chrome_dim}; }}
QSpinBox::down-arrow:disabled {{ border-top-color: {t.on_chrome_dim}; }}
"""


def current_qss() -> str:
    """Стиль для активной темы (используется диалогами при открытии)."""
    return build_qss(T)


# ── Переключение темы «на лету» ──────────────────────────────────────────────
# Виджеты с собственными инлайн-стилями (шапка, редактор зон, тулбар таблицы)
# регистрируют колбэк перекраски. Глобальный QSS обновляется одним вызовом.

_retheme_callbacks: list = []


def register_retheme(callback):
    """Регистрирует функцию, вызываемую при каждой смене темы."""
    if callback not in _retheme_callbacks:
        _retheme_callbacks.append(callback)


def unregister_retheme(callback):
    try:
        _retheme_callbacks.remove(callback)
    except ValueError:
        pass


def available_themes() -> list:
    return list(_THEME_ORDER)


def current_theme_name() -> str:
    return T.name


def set_theme(name: str, *, persist: bool = True):
    """Делает тему активной, обновляет весь интерфейс и (опц.) сохраняет выбор."""
    theme = THEMES.get(name)
    if theme is None:
        return
    _apply_theme_tokens(theme)
    app = QApplication.instance()
    if app is not None:
        app.setStyleSheet(build_qss(theme))
    for cb in list(_retheme_callbacks):
        try:
            cb()
        except Exception as e:
            _log_error(tr("перекраска виджета при смене темы"), e)
    if persist:
        settings = load_ui_settings()
        settings["theme"] = name
        save_ui_settings(settings)

# ─────────────────────────────────────────────────────────────────────────────
# Утилиты
# ─────────────────────────────────────────────────────────────────────────────

def _log_error(context: str, exc: Exception):
    """Единая точка для неблокирующих ошибок: пишем в stderr, но не роняем UI."""
    print(f"[ZoneExtract] {context}: {exc}", file=sys.stderr)

# ── Единые обёртки диалогов сообщений ────────────────────────────────────────
def _warn(parent, text: str, title: str = ""):
    title = title or tr("Предупреждение")
    QMessageBox.warning(parent, title, text)


def _info(parent, text: str, title: str = ""):
    title = title or tr("Готово")
    QMessageBox.information(parent, title, text)


def _error(parent, text: str, title: str = ""):
    title = title or tr("Ошибка")
    QMessageBox.critical(parent, title, text)


def _ask(parent, text: str, title: str = "") -> bool:
    """Вопрос Да/Нет; по умолчанию выбрано «Нет». True — если выбрано «Да»."""
    title = title or tr("Подтверждение")
    return QMessageBox.question(parent, title, text, _MB_YES | _MB_NO, _MB_NO) == _MB_YES



def _open_folder(path: str):
    try:
        if sys.platform == "win32":
            subprocess.Popen(f'explorer "{path}"')
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        _log_error(tr("не удалось открыть папку"), e)


def _pages_for_row(row: int, ppd: int, total: int) -> list[int]:
    """Номера страниц PDF, относящиеся к строке таблицы row.

    Каждому документу соответствует ppd страниц подряд (ppd = страниц на
    документ): [row*ppd, row*ppd+1, ...]. Отбрасываются страницы за
    пределами [0, total).
    """
    ppd = max(1, int(ppd))
    start = row * ppd
    return [start + k for k in range(ppd) if 0 <= start + k < total]


def _open_row_document(pdf_path: str, row: int, ppd: int, total: int,
                       start: Optional[int] = None) -> list[int]:
    """Извлекает страницы строки во временный PDF и открывает его внешне.

    Возвращает список открытых страниц. Бросает ValueError, если строке не
    соответствует ни одна страница (например, строк больше, чем документов).
    """
    if start is None:
        pages = _pages_for_row(row, ppd, total)
    else:
        pages = [start + k for k in range(max(1, int(ppd))) if 0 <= start + k < total]
    if not pages:
        raise ValueError(tr("Страница не существует."))
    temp = _extract_pages_to_temp(pdf_path, pages, tr("строка_{v}", v=row + 1))
    _open_pdf_externally(temp)
    return pages


def _safe_name(s: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", str(s)).strip() or tr("файл")


def _open_pdf_externally(pdf_path: str):
    """Открывает PDF во внешней программе."""
    try:
        if sys.platform == "win32":
            os.startfile(pdf_path)
        elif sys.platform == "darwin":
            subprocess.run(["open", pdf_path])
        else:
            subprocess.run(["xdg-open", pdf_path])
    except Exception as e:
        raise RuntimeError(str(e)) from e


def _extract_pages_to_temp(pdf_path: str, pages: list[int], tag: str) -> str:
    """Извлекает страницы из PDF во временный файл, возвращает его путь."""
    pages_str = "_".join(str(p + 1) for p in pages)
    fname     = tr("temp_{tag}_стр_{pages_str}.pdf", tag=tag, pages_str=pages_str)
    temp_path = os.path.join(tempfile.gettempdir(), _safe_name(fname))
    src = fitz.open(pdf_path)
    dst = fitz.open()
    for pg in pages:
        dst.insert_pdf(src, from_page=pg, to_page=pg)
    dst.save(temp_path)
    dst.close()
    src.close()
    return temp_path


# ── Папка с PDF как один источник ────────────────────────────────────────────

def _natural_key(s: str):
    """Ключ «естественной» сортировки: «2.pdf» < «10.pdf»."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def _list_pdfs(folder: str, recursive: bool = False) -> list[str]:
    """PDF-файлы папки в естественном порядке (с подпапками — по относительному пути)."""
    out = []
    if recursive:
        for root, dirs, files in os.walk(folder):
            dirs.sort(key=_natural_key)
            for f in files:
                if f.lower().endswith(".pdf"):
                    out.append(os.path.join(root, f))
    else:
        for f in os.listdir(folder):
            fp = os.path.join(folder, f)
            if f.lower().endswith(".pdf") and os.path.isfile(fp):
                out.append(fp)
    out.sort(key=lambda fp: _natural_key(os.path.relpath(fp, folder)))
    return out


def _has_pdfs_in_subfolders(folder: str) -> bool:
    for root, _dirs, files in os.walk(folder):
        if root != folder and any(f.lower().endswith(".pdf") for f in files):
            return True
    return False


def _merge_pdfs(files: list[str], ppd: int, progress=None) -> dict:
    """Склеивает PDF-файлы во временный файл (текстовый слой сохраняется).

    Если в файле число страниц не кратно ppd (страниц на документ), он
    добивается пустыми страницами — чтобы следующий файл начинался с нового
    документа и нарезка пачки не «съезжала».

    Возвращает словарь:
        path     — путь к склеенному временному PDF;
        page_src — для каждой страницы: (имя файла, номер стр. в файле) или
                   None для пустой страницы-добивки;
        padded   — [(имя файла, сколько добавлено пустых страниц)];
        failed   — [(имя файла, текст ошибки)].
    progress(i, total, name) -> bool: False = отмена.
    """
    ppd = max(1, int(ppd))
    out = fitz.open()
    page_src, padded, failed = [], [], []
    total = len(files)
    try:
        for i, fp in enumerate(files):
            name = os.path.basename(fp)
            if progress is not None and progress(i, total, name) is False:
                raise InterruptedError(tr("Отменено пользователем"))
            try:
                src = fitz.open(fp)
            except Exception as e:
                failed.append((name, str(e)))
                continue
            try:
                n = len(src)
                if n == 0:
                    failed.append((name, tr("нет страниц")))
                    continue
                out.insert_pdf(src)
                page_src.extend((name, k + 1) for k in range(n))
                rem = n % ppd
                if rem:
                    add = ppd - rem
                    r = src[n - 1].rect
                    for _ in range(add):
                        out.new_page(width=r.width, height=r.height)
                        page_src.append(None)
                    padded.append((name, add))
            finally:
                src.close()
        if len(out) == 0:
            raise ValueError(tr("Ни один PDF из папки не удалось прочитать."))
        fd, path = tempfile.mkstemp(prefix=tr("папка_"), suffix=".pdf")
        os.close(fd)
        out.save(path, garbage=1, deflate=True)
    finally:
        out.close()
    return {"path": path, "page_src": page_src, "padded": padded, "failed": failed}


def _clear_layout(lay):
    """Рекурсивно удаляет все виджеты и вложенные layout'ы из layout'а."""
    if lay is None:
        return
    while lay.count():
        it = lay.takeAt(0)
        w = it.widget()
        if w is not None:
            w.setParent(None)
            w.deleteLater()
        else:
            sub = it.layout()
            if sub is not None:
                _clear_layout(sub)


# ── Сопоставление по словарю (общая нечёткая логика) ─────────────────────────

def _norm_match(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _similarity(a: str, b: str) -> float:
    na, nb = _norm_match(a), _norm_match(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def _best_match(value: str, candidates: list[str]) -> tuple[str, float]:
    best, best_score = "", 0.0
    for c in candidates:
        score = _similarity(value, c)
        if score > best_score:
            best_score, best = score, c
    return best, best_score


# ── Проверка значения поля по регулярному выражению ──────────────────────────

def _field_valid(value: str, pattern: str) -> bool:
    """True, если значение проходит проверку.

    Пустой шаблон = проверки нет (всегда True). Проверяется полное
    совпадение (re.fullmatch). Некорректный шаблон трактуется как «без
    проверки», чтобы ошибка правила не ломала работу с таблицей.
    """
    if not pattern:
        return True
    try:
        return re.fullmatch(pattern, str(value).strip()) is not None
    except re.error:
        return True


def _clean_value(value: str, pattern: str) -> str:
    """Извлекает нужную часть значения по регулярному выражению.

    Применяется сразу после чтения зоны, ДО проверки. Логика:
      • пустой шаблон — вернуть как есть;
      • ищем совпадение (re.search) в строке;
      • если в шаблоне есть группы — вернуть их, склеенные подряд (без
        разделителей), иначе всё совпадение;
      • если совпадения нет или шаблон некорректен — вернуть исходный текст.
    Пример: шаблон  0*(\\d+)  из «ОРДЕР № 00000012345» вернёт «12345».
    Пример: шаблон  [Пп][Кк]\\s*(\\d+)\\s*\\+\\s*(\\d+)  из «ПК 400+30(право)»
    вернёт «40030» (группы склеиваются).
    """
    if not pattern:
        return value
    try:
        m = re.search(pattern, str(value))
    except re.error:
        return value
    if not m:
        return value
    groups = [g for g in m.groups() if g is not None] if m.lastindex else []
    out = "".join(groups) if groups else m.group(0)
    return out.strip()


# ─────────────────────────────────────────────────────────────────────────────
# OCR-обработка полей
# ─────────────────────────────────────────────────────────────────────────────
# «Обработка» выполняется ДО «Очистки (regex)»: сырой текст зоны → обработка →
# очистка → проверка. Каждый обработчик: str → str (пусто = не распознано).

_LAT_TO_CYR = str.maketrans(
    "ABEKMHOPCTYXabcehkmoprtxy",
    "АВЕКМНОРСТУХавсенкмоертху",
)

# Буквы, похожие на цифры, в числовой зоне → цифры.
_VOL_CHAR_MAP = str.maketrans({
    "J": "3", "j": "3", "З": "3", "з": "3",
    "О": "0", "о": "0", "O": "0", "o": "0",
    "I": "1", "l": "1", "|": "1", "i": "1", "t": "1", "!": "1",
    "S": "5", "s": "5", "Ц": "4", "В": "8", "в": "8", "б": "6",
})


def _post_digits(raw: str) -> str:
    return re.sub(r"\D", "", str(raw))


def _post_number(raw: str) -> str:
    """Первое число строки; разделитель дроби — запятая."""
    s = str(raw).replace("\u00a0", "").replace(" ", "").replace(",", ".")
    m = re.search(r"\d+\.?\d*", s)
    return m.group(0).rstrip(".").replace(".", DEC_SEP) if m else ""


_OCR_TOKEN = re.compile(r"[0-9OОоoЗзIl|!SsВвбЦJjti.,]+")


def _ocr_digits_in_numbers(s: str) -> str:
    """Буквы-двойники → цифры только внутри чисел (рядом с настоящей цифрой).
    «Вес груза 4О,5» → «Вес груза 40,5»: слова не трогаются."""
    return _OCR_TOKEN.sub(
        lambda m: m.group(0).translate(_VOL_CHAR_MAP) if re.search(r"\d", m.group(0)) else m.group(0),
        str(s))


def _post_number_ocr(raw: str) -> str:
    """Как «Число», но буквы-двойники внутри числа (О, З, l, S, В…) → цифры."""
    s = _ocr_digits_in_numbers(raw)
    s = re.sub(r"(?<=\d)\s+(?=\d)", "", s)
    m = re.search(r"\d+(?:[.,]\d+)?", s)
    return m.group(0).replace(",", ".").replace(".", DEC_SEP) if m else ""


def _post_plate_ru(raw: str) -> str:
    """Гос. номер РФ: латиница→кириллица, О↔0 и З↔3 по позициям."""
    clean = "".join(
        ch.translate(_LAT_TO_CYR)
        for ch in str(raw).upper()
        if re.match(r"[А-ЯЁ0-9]", ch.translate(_LAT_TO_CYR))
    )
    pat = r"([АВЕКМНОРСТУХ0])(\d|[ОЗ]){3}([АВЕКМНОРСТУХ0]{2})(\d|[ОЗ]){2,3}"
    match = re.search(pat, clean)
    if not match:
        return re.sub(r"[^А-ЯЁ0-9]", "", clean)
    chars = list(match.group(0))
    if chars[0] == "0":
        chars[0] = "О"
    for i in range(1, 4):
        if chars[i] == "О": chars[i] = "0"
        if chars[i] == "З": chars[i] = "3"
    for i in range(4, 6):
        if chars[i] == "0": chars[i] = "О"
    for i in range(6, len(chars)):
        if chars[i] == "О": chars[i] = "0"
        if chars[i] == "З": chars[i] = "3"
    return "".join(chars)


def _post_plate_slash(raw: str) -> str:
    """Гос. номер до «/» (тягач/прицеп записаны через слэш)."""
    return _post_plate_ru(str(raw).split("/", 1)[0])


_FIO_STOP = frozenset(["вод", "удост", "реквизит", "идентиф", "паспорт", "пасп",
                       "телефон", "тел", "серия", "права", "лицен", "номер", "свид"])


def _post_fio_short(raw: str) -> str:
    """ФИО водителя → «Фамилия И.О.»; хвост с реквизитами и OCR-мусор отбрасываются."""
    raw = str(raw)
    m = re.search(r",", raw)
    if m:
        raw = raw[:m.start()]
    m = re.search(r"[\s]+(?:вод[.\s]|удост|реквизит|идентиф|позвол|паспорт|пасп|"
                  r"телефон|тел[.\s]|серия|права|лицен|свид)", raw, re.I)
    if m:
        raw = raw[:m.start()]
    lines = [ln for ln in raw.split("\n") if re.search(r"[А-ЯЁа-яё]", ln)]
    text = " ".join(lines) if lines else raw
    m = re.search(r"[А-ЯЁ]", text)
    if m:
        text = text[m.start():]
    text = re.sub(r"[•·]\s*\S*", " ", text)
    text = re.sub(r"(?<=[А-ЯЁа-яё])[''\u2019`\-](?=[А-ЯЁа-яё])", "", text)
    for _ in range(3):
        text = re.sub(r"(?:^|(?<=\s))([А-ЯЁ])\s+([А-ЯЁа-яё]{3,})", r"\1\2", text)
        text = re.sub(r"([А-ЯЁа-яё]{3,})\s+([а-яё])(?=\s|$)", r"\1\2", text)
    text = re.sub(r"\s+", " ", text).strip()
    words = []
    for w in re.sub(r"[^А-ЯЁа-яё\s]", " ", text).split():
        if any(s in w.lower() for s in _FIO_STOP):
            break
        if len(w) >= 2:
            words.append(w)
    if not words:
        return ""
    surname = words[0].capitalize()
    initials = "".join(w[0].upper() + "." for w in words[1:3] if w)
    return f"{surname} {initials}".strip() if initials else surname


_MONTHS = {"январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма": 5, "июн": 6,
           "июл": 7, "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
           "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
           "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}


def _month_of(word: str) -> int:
    word = word.lower()
    for stem in sorted(_MONTHS, key=len, reverse=True):
        if word.startswith(stem):
            return _MONTHS[stem]
    return 0


def _post_date_words(raw: str) -> str:
    """«"5" марта 2025 г.», «5 March 2025», «March 5, 2025» → «05.03.2025»."""
    s = str(raw).replace("\n", " ")
    day = month = 0
    year = ""
    m = re.search(r"(\d{1,2})\D{0,3}?\s*([А-ЯЁа-яёA-Za-z]{3,})\.?,?\s*(\d{4})", s)
    if m:
        day, month, year = int(m.group(1)), _month_of(m.group(2)), m.group(3)
    if not month:
        m = re.search(r"([A-Za-zА-ЯЁа-яё]{3,})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s*(\d{4})", s)
        if m:
            day, month, year = int(m.group(2)), _month_of(m.group(1)), m.group(3)
    if not month or not 1 <= day <= 31:
        return ""
    return f"{day:02d}.{month:02d}.{year}"


_RE_NETTO = re.compile(r"(?:нетто|net(?:\s*weight)?)\s*[:\-\u2013]?\s*([\d\u00a0.,\s]+?)\s*(?:кг|kg)",
                       re.IGNORECASE)


def _post_netto_t(raw: str) -> str:
    """«Нетто: 51 470 кг» / «Net weight: 51470 kg» → вес в тоннах «51,47»."""
    m = _RE_NETTO.search(str(raw).replace("\n", " "))
    if not m:
        return ""
    num = m.group(1).replace("\u00a0", "").replace(" ", "").strip(" .,")
    if not num:
        return ""
    try:
        kg = Decimal(num.replace(",", "."))
    except InvalidOperation:
        return ""
    t = format((kg / 1000).normalize(), "f")
    return t.replace(".", DEC_SEP)


# (ключ, подпись, функция). Ключ хранится в шаблоне: поле «post».
# Подпись переводится при показе (tr).
POST_PROCESSORS = [
    ("",               "— нет —",                      None),
    ("number",         "Число",                         _post_number),
    ("number_ocr",     "Число + OCR (О→0, З→3…)",       _post_number_ocr),
    ("digits",         "Только цифры",                  _post_digits),
    ("plate_ru",       "Гос. номер РФ (OCR)",           _post_plate_ru),
    ("plate_ru_slash", "Гос. номер РФ до «/»",          _post_plate_slash),
    ("fio_short",      "ФИО → Фамилия И.О.",            _post_fio_short),
    ("netto_t",        "Нетто, кг → тонны",             _post_netto_t),
    ("date_words",     "Дата словами → ДД.ММ.ГГГГ",     _post_date_words),
]
POST_MAP = {k: (label, fn) for (k, label, fn) in POST_PROCESSORS}


def process_value(raw: str, post: str = "", clean: str = "") -> str:
    """Сырой текст зоны → обработка (post) → очистка (regex)."""
    v = raw or ""
    fn = POST_MAP.get(post or "", ("", None))[1]
    if fn is not None:
        try:
            v = fn(v)
        except Exception as e:
            _log_error(tr("обработчик «{post}»", post=post), e)
    return _clean_value(v, clean)


# ─────────────────────────────────────────────────────────────────────────────
# Вычисляемые поля (формулы)
# ─────────────────────────────────────────────────────────────────────────────
# Формула: арифметика над полями в квадратных скобках. Пример: [Тонн] / 1,6
# Разрешено: + - * / ( ) и числа (дробь через точку или запятую).
# Счёт в Decimal, округление математическое (ROUND_HALF_UP): 28,375 → 28,38.

_FORMULA_REF = re.compile(r"\[([^\[\]]+)\]")
_FORMULA_NODES = (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Name,
                  ast.Load, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd)


def formula_refs(formula: str) -> list:
    """Имена полей, на которые ссылается формула (без повторов, по порядку)."""
    out = []
    for n in _FORMULA_REF.findall(str(formula)):
        n = n.strip()
        if n not in out:
            out.append(n)
    return out


@lru_cache(maxsize=256)
def compile_formula(formula: str):
    """Разбирает формулу. Возвращает (узел AST, список имён полей).
    Бросает ValueError/SyntaxError, если формула некорректна."""
    names = formula_refs(formula)
    expr = _FORMULA_REF.sub(lambda m: f"__v{names.index(m.group(1).strip())}", formula)
    expr = expr.replace(",", ".")
    tree = ast.parse(expr.strip(), mode="eval")
    for node in ast.walk(tree):
        if not isinstance(node, _FORMULA_NODES):
            raise ValueError(tr("недопустимая операция: {name__}", name__=type(node).__name__))
        if isinstance(node, ast.Constant) and not isinstance(node.value, (int, float)):
            raise ValueError(tr("допустимы только числа"))
        if isinstance(node, ast.Name) and not re.fullmatch(r"__v\d+", node.id):
            raise ValueError(tr("неизвестное имя «{id}» — поля пишутся в [скобках]", id=node.id))
    return tree.body, tuple(names)


def _to_decimal(s):
    t = str(s).replace("\u00a0", "").replace(" ", "").replace(",", ".").strip()
    if not t:
        return None
    try:
        return Decimal(t)
    except InvalidOperation:
        return None


def _eval_node(node, vals):
    if isinstance(node, ast.Constant):
        return Decimal(str(node.value))
    if isinstance(node, ast.Name):
        return vals[int(node.id[3:])]
    if isinstance(node, ast.UnaryOp):
        v = _eval_node(node.operand, vals)
        return -v if isinstance(node.op, ast.USub) else v
    a, b = _eval_node(node.left, vals), _eval_node(node.right, vals)
    if isinstance(node.op, ast.Add):
        return a + b
    if isinstance(node.op, ast.Sub):
        return a - b
    if isinstance(node.op, ast.Mult):
        return a * b
    return a / b


def eval_formula(formula: str, values: dict, decimals: int = 2) -> str:
    """Считает формулу по значениям полей. Пусто — если поля нет/не число."""
    try:
        node, names = compile_formula(str(formula))
    except Exception:
        return ""
    vals = []
    for n in names:
        d = _to_decimal(values.get(n, ""))
        if d is None:
            return ""
        vals.append(d)
    try:
        res = _eval_node(node, vals)
        res = res.quantize(Decimal(1).scaleb(-int(decimals)), rounding=ROUND_HALF_UP)
    except (ArithmeticError, IndexError):
        return ""
    return format(res, "f").replace(".", DEC_SEP)


# ─────────────────────────────────────────────────────────────────────────────
# Отчёт (колонки для вставки в Google Таблицы)
# ─────────────────────────────────────────────────────────────────────────────
# Строка описания = одна колонка:
#   [Имя поля]          — значение поля строки таблицы
#   Подпись = значение  — постоянный текст (правится в диалоге перед копированием)
#   -                   — пустая колонка
#   любой другой текст  — постоянный текст без правки

def parse_report_line(line: str) -> tuple:
    s = str(line).strip()
    if s in ("", "-", "—", "–"):
        return ("empty", "", "")
    m = re.fullmatch(r"\[(.+)\]", s)
    if m:
        return ("field", m.group(1).strip(), "")
    if "=" in s:
        a, b = s.split("=", 1)
        return ("text", a.strip(), b.strip())
    return ("const", "", s)


# ── Пресеты типов извлекаемых полей ──────────────────────────────────────────
# Каждый элемент: (ключ, подпись, шаблон проверки, шаблон очистки, пример).
#   • «шаблон проверки» — re.fullmatch по УЖЕ очищенному значению;
#   • «шаблон очистки»  — подсказка для колонки «Очистка» (re.search, группа
#     вытаскивает значение из «шумного» текста зоны). Подставляется как
#     умолчание, поле остаётся редактируемым.
# Ключ ""       — свободный текст (без проверки);
# ключ "custom" — ручной регекс из поля «Проверка (regex)».
FIELD_TYPES = [
    ("",           "Свободный текст (без проверки)", "",                             "",                                   ""),
    ("int",        "Число — целое",                  r"\d+",                         r"(\d+)",                             "12345"),
    ("decimal",    "Число — дробное",                r"\d+([.,]\d+)?",               r"(\d+(?:[.,]\d+)?)",                 "12,5"),
    ("money",      "Сумма (деньги)",                 r"\d{1,3}([ .]?\d{3})*([.,]\d{1,2})?", r"(\d[\d ]*(?:[.,]\d{1,2})?)", "1 234,56"),
    ("qty",        "Количество",                     r"\d+([.,]\d+)?",               r"(\d+(?:[.,]\d+)?)",                 "10"),
    ("date_full",  "Дата ДД.ММ.ГГГГ",                r"\d{2}\.\d{2}\.\d{4}",         r"(\d{2}\.\d{2}\.\d{4})",             "31.12.2025"),
    ("date_short", "Дата ДД.ММ.ГГ",                  r"\d{2}\.\d{2}\.\d{2}",         r"(\d{2}\.\d{2}\.\d{2})",             "31.12.25"),
    ("date_any",   "Дата (. / -)",                   r"\d{2}[./\-]\d{2}[./\-]\d{2,4}", r"(\d{2}[./\-]\d{2}[./\-]\d{2,4})",  "31/12/2025"),
    ("doc_no",     "Номер документа",                r"[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9 №/\-]*", "",                  "ТН-000123"),
    ("article",    "Артикул / код",                  r"[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9/_.\-]*", "",                  "ART-10-45B"),
    ("inn",        "ИНН (10 или 12 цифр)",           r"\d{10}|\d{12}",               r"(\d{12}|\d{10})",                   "7701234567"),
    ("kpp",        "КПП (9 цифр)",                   r"\d{9}",                       r"(\d{9})",                           "770101001"),
    ("ean13",      "Штрихкод EAN-13",                r"\d{13}",                      r"(\d{13})",                          "4601234567890"),
    ("ean8",       "Штрихкод EAN-8",                 r"\d{8}",                       r"(\d{8})",                           "46012345"),
    ("plate_ru",   "Гос. номер авто (РФ)",           r"[АВЕКМНОРСТУХ]\d{3}[АВЕКМНОРСТУХ]{2}\d{2,3}", r"([АВЕКМНОРСТУХ]\d{3}[АВЕКМНОРСТУХ]{2}\d{2,3})", "А123ВС77"),
    ("custom",     "Своё (регекс)",                  None,                           None,                                 ""),
]
FIELD_TYPE_MAP = {k: (label, pat, clean, ex) for (k, label, pat, clean, ex) in FIELD_TYPES}


def resolve_check_pattern(field: dict) -> str:
    """Итоговый шаблон проверки поля с учётом выбранного типа.

    Пресеты возвращают свой шаблон; тип «custom» и старые шаблоны без типа —
    ручной регекс из ключа «regex»; свободный текст — пусто (без проверки).
    """
    t = field.get("type")
    if t is None:                       # старый шаблон: тип не задан
        return field.get("regex", "")
    if t == "custom":
        return field.get("regex", "")
    return FIELD_TYPE_MAP.get(t, ("", "", "", ""))[1] or ""


def type_clean_suggestion(type_key) -> str:
    """Рекомендуемый шаблон очистки для типа (для custom/неизвестного — пусто)."""
    if not type_key or type_key == "custom":
        return ""
    return FIELD_TYPE_MAP.get(type_key, ("", "", "", ""))[2] or ""



_DOC_CODE_CLEAN = r"(?:№|(?i:\bno\b)\.?|#)\s*([A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9/\-]*)"

# ─────────────────────────────────────────────────────────────────────────────
# «Что это за поле» — готовые наборы правил (настройка поля по примеру)
# ─────────────────────────────────────────────────────────────────────────────
# Каждый вид поля = обработка (post) + очистка (clean) + проверка (type/regex).
# Пользователь выбирает вид или вставляет пример текста зоны и жмёт «🔮».
# Ручная настройка regex доступна в «⚙» и не нужна в обычной работе.

# key, подпись, post, clean, type, regex, пример текста зоны, пояснение.
# Подпись, пример и пояснение переводятся при показе (tr).
FIELD_KINDS = [
    ("text",       "Текст как есть",                  "",               "",                                   "",          "",
     "ООО «Ромашка»",                     "Без изменений и без проверки."),
    ("number",     "Число: вес, объём, сумма",        "number_ocr",     "",                                   "decimal",   "",
     "Масса груза 44,1 т",                "Первое число в зоне. Буквы-двойники внутри числа исправляются: О→0, З→3."),
    ("num_nozero", "Номер без нулей впереди",         "",               r"0*(\d+)",                           "int",       "",
     "ОРДЕР № 00000012345",               "Цифры номера, нули в начале убираются."),
    ("digits_num", "Номер из цифр (как есть)",        "",               r"(\d+)",                             "int",       "",
     "Накладная № 004512",                "Первая группа цифр, нули сохраняются."),
    ("doc_code",   "Номер с буквами (после «№»)",     "",               _DOC_CODE_CLEAN,                      "doc_no",    "",
     "Накладная № ТН-000123 от 05.03",    "Берётся код после «№», «No» или «#»: буквы, цифры, «-» и «/»."),
    ("date_full",  "Дата 05.03.2025",                 "",               r"(\d{2}\.\d{2}\.\d{4})",             "date_full", "",
     "от 05.03.2025 г.",                  "Дата в формате ДД.ММ.ГГГГ."),
    ("date_words", "Дата «5 марта 2025» → 05.03.2025", "date_words",     "",                                   "date_full", "",
     "«5» марта 2025 г.",                 "Месяц словом переводится в число."),
    ("plate_ru",   "Гос. номер РФ",                   "plate_ru",       "",                                   "plate_ru",  "",
     "T 123 OP 178",                      "Латиница → кириллица, О↔0 и З↔3 по позициям."),
    ("plate_slash","Гос. номер тягача (до «/»)",      "plate_ru_slash", "",                                   "plate_ru",  "",
     "Т123ОР178 / АВ1234 78",             "Берётся номер до «/», номер прицепа отбрасывается."),
    ("fio",        "ФИО → Фамилия И.О.",              "fio_short",      "",                                   "",          "",
     "Иванов Иван Петрович, паспорт 4000",
     "Сокращает ФИО, отрезает паспорт, телефон и мусор."),
    ("netto_t",    "Нетто, кг → тонны",               "netto_t",        "",                                   "decimal",   "",
     "Нетто: 51 470 кг",                  "Килограммы делятся на 1000."),
    ("inn",        "ИНН",                             "",               r"(?<!\d)(\d{12}|\d{10})(?!\d)",       "inn",       "",
     "ИНН 7701234567",                    "10 или 12 цифр."),
]
KIND_CUSTOM = "custom"
KIND_MAP = {k[0]: {"label": k[1], "post": k[2], "clean": k[3], "type": k[4],
                   "regex": k[5], "example": k[6], "hint": k[7]} for k in FIELD_KINDS}

# Порядок подбора: от самых узнаваемых видов к общим.
_KIND_RANK = ["plate_slash", "plate_ru", "inn", "date_full", "date_words", "netto_t",
              "number", "num_nozero", "digits_num", "doc_code", "fio"]
# Слова в названии поля, которые подсказывают вид (русские и английские).
_KIND_HINTS = {
    "date_full": ("дата", "date"), "date_words": ("дата", "date"),
    "fio": ("фио", "водител", "получ", "сдал", "принял", "отпуст", "name", "driver"),
    "plate_ru": ("гос", "тс", "авто", "машин", "plate", "vehicle", "truck"),
    "plate_slash": ("гос", "тс", "авто", "plate", "truck"),
    "netto_t": ("вес", "тонн", "нетто", "масс", "net", "weight"),
    "number": ("вес", "тонн", "объ", "сумм", "кол", "масс", "цена",
               "weight", "qty", "quantity", "amount", "total", "price", "volume", "sum"),
    "doc_code": ("№", "номер", "наклад", "invoice", "number", "no", "doc"),
    "inn": ("инн", "inn", "tax"),
}


def kind_state(key: str) -> dict:
    """Правила поля для вида key: post, clean, type, regex."""
    k = KIND_MAP.get(key) or KIND_MAP["text"]
    return {"post": k["post"], "clean": k["clean"], "type": k["type"], "regex": k["regex"]}


def match_kind(state: dict) -> str:
    """Вид поля, которому соответствуют правила; иначе «custom»."""
    t = state.get("type") or ""
    for key in KIND_MAP:
        k = KIND_MAP[key]
        if ((state.get("post") or "") == k["post"] and (state.get("clean") or "") == k["clean"]
                and t == k["type"] and (t != "custom" or (state.get("regex") or "") == k["regex"])):
            return key
    return KIND_CUSTOM


def rule_result(example: str, state: dict) -> tuple:
    """Прогоняет пример через правила. Возвращает (текст, статус ok/err/neutral)."""
    if not str(example).strip():
        return ("", "neutral")
    pattern = resolve_check_pattern(state)
    for label, pat in ((tr("очистки"), state.get("clean", "")), (tr("проверки"), pattern)):
        if pat:
            try:
                re.compile(pat)
            except re.error as e:
                return (tr("ошибка regex {label}: {e}", label=label, e=e), "err")
    res = process_value(example, state.get("post", ""), state.get("clean", "")).strip()
    if not res:
        return (tr("→ пусто: правило ничего не нашло"), "err")
    if not pattern:
        return (f"→ «{res}»", "neutral")
    if _field_valid(res, pattern):
        return (f"→ «{res}»  ✓", "ok")
    return (tr("→ «{res}»  ✗ не прошло проверку", res=res), "err")


def suggest_kinds(example: str, field_name: str = "") -> list:
    """Подбирает виды поля по примеру текста зоны.
    Возвращает [(key, результат)] — лучший вариант первым."""
    ex = str(example or "")
    if not ex.strip():
        return []
    name = str(field_name or "").lower()
    found = []
    has_digit = re.search(r"\d", ex) is not None
    for key in _KIND_RANK:
        st = kind_state(key)
        if not has_digit and key not in ("fio", "date_words"):
            continue
        try:
            res = process_value(ex, st["post"], st["clean"]).strip()
        except Exception:
            continue
        if not res:
            continue
        if key == "fio":
            words = re.findall(r"[А-ЯЁа-яё]{2,}", ex)[:3]
            caps = sum(1 for w in words if w[0].isupper() and not w.isupper()) >= 2
            if re.search(r"[«»\"]|\b(?:ООО|ОАО|ЗАО|ПАО|АО|ИП|ГУП|МУП)\b", ex):
                caps = False
            hinted = any(h in name for h in _KIND_HINTS["fio"])
            ok = ((caps or hinted)
                  and re.fullmatch(r"[А-ЯЁ][а-яё]+ [А-ЯЁ]\.(?:[А-ЯЁ]\.)?", res) is not None)
        elif key in ("plate_ru", "plate_slash"):
            flat = re.sub(r"\s+", "", ex.upper().translate(_LAT_TO_CYR))
            looks = re.search(r"[АВЕКМНОРСТУХ]\d{3}[АВЕКМНОРСТУХ]{2}\d{2,3}", flat) is not None
            hinted = any(h in name for h in _KIND_HINTS.get(key, ()))
            ok = ((looks or hinted) and _field_valid(res, resolve_check_pattern(st))
                  and (key == "plate_ru" or "/" in ex))
        elif key == "doc_code":
            # Код документа ищем только после явной метки «№», «No» или «#».
            ok = (re.search(r"№|(?i:\bno\b)|#", ex) is not None
                  and _field_valid(res, resolve_check_pattern(st)))
        else:
            pat = resolve_check_pattern(st)
            ok = bool(pat) and _field_valid(res, pat)
        if ok:
            found.append((key, res))

    def _is_date(m):
        d, mo = int(m.group(1)), int(m.group(2))
        return 1 <= d <= 31 and 1 <= mo <= 12

    # Дроби в тексте, не считая дат вида 05.03 и 05.03.2025
    no_dates = re.sub(r"(?<!\d)(\d{1,2})[./](\d{1,2})(?:[./]\d{2,4})?(?!\d)",
                      lambda m: " " if _is_date(m) else m.group(0), ex)
    has_dec = re.search(r"\d[.,]\d", _ocr_digits_in_numbers(no_dates)) is not None

    def score(item):
        key = item[0]
        s = _KIND_RANK.index(key)
        hinted = any(h in name for h in _KIND_HINTS.get(key, ()))
        if hinted:
            s -= 100
        if has_dec and key in ("number", "netto_t"):
            s -= 50
        if key == "doc_code" and re.search(r"[A-Za-zА-Яа-яЁё]", item[1]):
            s -= 10
        if key == "fio":
            s -= 20
        if key == "number" and not has_dec and not hinted:
            s += 60
        return s
    found.sort(key=score)
    return found or [("text", ex.strip())]


# Виды обработки, после которых значение — число (для Excel и итогов).
_NUMERIC_POSTS = {"number", "number_ocr", "netto_t"}
_NUMERIC_TYPES = {"int", "decimal", "money", "qty"}
_SUM_TYPES = {"decimal", "money", "qty"}


# ─────────────────────────────────────────────────────────────────────────────
# Модель шаблона
# ─────────────────────────────────────────────────────────────────────────────

class Template:
    """Пользовательский шаблон извлечения.

    Атрибуты:
        name           — название шаблона.
        pages          — сколько страниц на один документ (1..MAX_DOC_PAGES).
                         Пачка PDF нарезается на документы по pages страниц.
        manual_fields  — список названий полей, заполняемых вручную (общие для пачки).
        extract_fields — список полей извлечения; каждое — словарь:
                           {"name": str, "page": int (0..pages-1),
                            "dict": bool, "dict_path": str,
                            "regex": str, "clean": str, "fallback": bool}
        zones          — {name: (x0, y0, x1, y1)} основных зон (координаты 300 DPI).
        fallback_zones — {name: (x0, y0, x1, y1)} резервных зон для полей с
                         fallback=True. Читается, если основная зона пуста.
        name_field     — имя извлекаемого поля, по значению которого называются
                         файлы при разбивке пачки. Пусто = первое извлекаемое поле.
        column_order   — порядок столбцов таблицы (имена ручных и извлекаемых
                         полей вперемешку). Пусто = сначала ручные, потом извлекаемые.
    """

    def __init__(self, name="", pages=1, manual_fields=None,
                 extract_fields=None, zones=None, fallback_zones=None,
                 name_field="", column_order=None, computed_fields=None,
                 report_lines=None, auto_number=None, auto_start=None):
        self.name           = name
        self.pages          = max(1, min(MAX_DOC_PAGES, int(pages)))
        self.manual_fields  = list(manual_fields or [])
        self.extract_fields = [dict(f) for f in (extract_fields or [])]
        self.zones          = {k: tuple(v) for k, v in (zones or {}).items()}
        self.fallback_zones = {k: tuple(v) for k, v in (fallback_zones or {}).items()}
        self.name_field     = str(name_field or "")
        self.column_order   = [str(x) for x in (column_order or [])]
        self.computed_fields = [dict(c) for c in (computed_fields or [])]
        self.report_lines   = [str(x) for x in (report_lines or [])]
        self.auto_number    = [str(x) for x in (auto_number or [])]
        self.auto_start     = dict(auto_start or {})
        self._normalize()

    def _normalize(self):
        """Приводит поля извлечения к единому виду и достраивает зоны по умолчанию."""
        for i, f in enumerate(self.extract_fields):
            f["name"]      = str(f.get("name", "")).strip()
            page           = int(f.get("page", 0))
            f["page"]      = max(0, min(self.pages - 1, page))
            f["dict"]      = bool(f.get("dict", False))
            f["dict_path"] = str(f.get("dict_path", ""))
            f["regex"]     = str(f.get("regex", ""))
            f["clean"]     = str(f.get("clean", ""))
            f["fallback"]  = bool(f.get("fallback", False))
            post           = str(f.get("post", "") or "")
            f["post"]      = post if post in POST_MAP else ""
            f["example"]   = str(f.get("example", ""))
            # Тип поля (пресет проверки). Старые шаблоны без типа: если задан
            # ручной regex — считаем тип "custom", иначе — свободный текст.
            t = f.get("type")
            if t is None or t not in FIELD_TYPE_MAP:
                t = "custom" if f["regex"].strip() else ""
            f["type"] = t
            f["kind"] = match_kind(f)
        # Гарантируем основную зону для каждого извлекаемого поля
        for i, f in enumerate(self.extract_fields):
            if f["name"] and f["name"] not in self.zones:
                self.zones[f["name"]] = _default_zone(i)
        # Гарантируем резервную зону для полей, где включён fallback
        for f in self.extract_fields:
            if f["name"] and f["fallback"] and f["name"] not in self.fallback_zones:
                self.fallback_zones[f["name"]] = _default_fallback_zone(self.zones[f["name"]])
        # Убираем зоны для несуществующих полей
        valid = {f["name"] for f in self.extract_fields}
        self.zones = {k: tuple(v) for k, v in self.zones.items() if k in valid}
        self.fallback_zones = {k: tuple(v) for k, v in self.fallback_zones.items() if k in valid}
        # Поле для имени файла при разбивке: допускается только извлекаемое поле.
        # Пусто = авто (первое извлекаемое поле).
        if self.name_field and self.name_field not in valid:
            self.name_field = ""
        # Порядок столбцов: оставляем существующие поля (без дублей),
        # недостающие дописываем в конец в порядке «ручные → извлекаемые».
        cf_out = []
        for c in self.computed_fields:
            try:
                dec = int(c.get("decimals", 2))
            except (TypeError, ValueError):
                dec = 2
            cf_out.append({"name": str(c.get("name", "")).strip(),
                           "formula": str(c.get("formula", "")),
                           "decimals": max(0, min(6, dec))})
        self.computed_fields = cf_out
        self.auto_number = [n for n in dict.fromkeys(self.auto_number) if n in self.manual_fields]
        starts = {}
        for n in self.auto_number:
            try:
                starts[n] = max(0, int(self.auto_start.get(n, 1)))
            except (TypeError, ValueError):
                starts[n] = 1
        self.auto_start = starts
        all_names = (list(self.manual_fields) + [f["name"] for f in self.extract_fields]
                     + [c["name"] for c in self.computed_fields])
        known = set(all_names)
        order, seen = [], set()
        for nm in self.column_order:
            if nm in known and nm not in seen:
                order.append(nm)
                seen.add(nm)
        order += [nm for nm in all_names if nm not in seen]
        self.column_order = order

    @property
    def duplex(self) -> bool:
        # Сохранено для совместимости; означает «более одной страницы на документ».
        return self.pages >= 2

    @property
    def extract_names(self) -> list[str]:
        return [f["name"] for f in self.extract_fields]

    @property
    def key_field(self) -> str:
        """Ключевое поле: по нему ищутся дубли и называются файлы при разбивке."""
        names = self.extract_names
        if self.name_field in names:
            return self.name_field
        return names[0] if names else ""

    def numeric_names(self) -> set:
        """Поля, которые в Excel пишутся числом."""
        out = {f["name"] for f in self.extract_fields
               if f.get("type") in _NUMERIC_TYPES or f.get("post") in _NUMERIC_POSTS}
        out |= set(self.computed_names) | set(self.auto_number)
        return out

    def sum_names(self) -> list:
        """Поля, по которым считается итог (вес, объём, суммы)."""
        out = [f["name"] for f in self.extract_fields
               if f.get("type") in _SUM_TYPES or f.get("post") in _NUMERIC_POSTS]
        out += self.computed_names
        return [n for n in self.column_order if n in out]

    @property
    def computed_names(self) -> list[str]:
        return [c["name"] for c in self.computed_fields]

    @property
    def columns(self) -> list[str]:
        return list(self.column_order)

    def column_index(self) -> dict:
        """{имя поля: индекс столбца в таблице}."""
        return {nm: i for i, nm in enumerate(self.column_order)}

    @property
    def dict_fields(self) -> list[dict]:
        return [f for f in self.extract_fields if f["dict"]]

    @property
    def fallback_fields(self) -> list[dict]:
        return [f for f in self.extract_fields if f.get("fallback")]

    def colors(self) -> dict:
        return {f["name"]: _palette_color(i) for i, f in enumerate(self.extract_fields)}

    def default_zones(self) -> dict:
        return {f["name"]: _default_zone(i) for i, f in enumerate(self.extract_fields)}

    def default_fallback_zones(self) -> dict:
        return {f["name"]: _default_fallback_zone(_default_zone(i))
                for i, f in enumerate(self.extract_fields) if f.get("fallback")}

    def to_dict(self) -> dict:
        return {
            "name":           self.name,
            "pages":          self.pages,
            "manual_fields":  list(self.manual_fields),
            "extract_fields": [dict(f) for f in self.extract_fields],
            "zones":          {k: list(v) for k, v in self.zones.items()},
            "fallback_zones": {k: list(v) for k, v in self.fallback_zones.items()},
            "name_field":     self.name_field,
            "column_order":   list(self.column_order),
            "computed_fields": [dict(c) for c in self.computed_fields],
            "report_lines":   list(self.report_lines),
            "auto_number":    list(self.auto_number),
            "auto_start":     dict(self.auto_start),
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Template":
        return cls(
            name=d.get("name", ""),
            pages=d.get("pages", 1),
            manual_fields=d.get("manual_fields", []),
            extract_fields=d.get("extract_fields", []),
            zones=d.get("zones", {}),
            fallback_zones=d.get("fallback_zones", {}),
            name_field=d.get("name_field", ""),
            column_order=d.get("column_order", []),
            computed_fields=d.get("computed_fields", []),
            report_lines=d.get("report_lines", []),
            auto_number=d.get("auto_number", []),
            auto_start=d.get("auto_start", {}),
        )


def load_templates() -> tuple[list, str]:
    """Загружает список шаблонов и имя выбранного из файла конфигурации."""
    try:
        with open(_TEMPLATES_CFG, "r", encoding="utf-8") as f:
            data = json.load(f)
        tpls = [Template.from_dict(t) for t in data.get("templates", [])]
        return tpls, data.get("selected", "")
    except FileNotFoundError:
        return [], ""
    except Exception as e:
        _log_error(tr("не удалось прочитать конфигурацию шаблонов"), e)
        return [], ""


def save_templates(templates: list, selected: str = "") -> bool:
    """Сохраняет шаблоны в конфиг. Возвращает True при успехе."""
    try:
        data = {"templates": [t.to_dict() for t in templates], "selected": selected}
        with open(_TEMPLATES_CFG, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        _log_error(tr("не удалось сохранить конфигурацию шаблонов"), e)
        return False

# ─────────────────────────────────────────────────────────────────────────────
# Парсер PDF
# ─────────────────────────────────────────────────────────────────────────────

class PdfParser:
    """Читает текст из текстового слоя PDF по прямоугольным зонам."""

    BASE_DPI = 300

    def __init__(self, pdf_path: str, dpi: int = 300, zones: Optional[dict] = None):
        self.pdf_path = pdf_path
        self.dpi      = dpi
        self._scale   = dpi / self.BASE_DPI
        self.zones    = {k: tuple(v) for k, v in (zones or {}).items()}
        self._render_cache: dict = {}
        self.doc      = fitz.open(pdf_path)

    def set_dpi(self, dpi: int):
        self.dpi    = dpi
        self._scale = dpi / self.BASE_DPI

    def set_zones(self, zones: dict):
        self.zones = {k: tuple(v) for k, v in zones.items()}

    def _scaled(self, coords: tuple) -> tuple:
        return tuple(int(v * self._scale) for v in coords)

    def _read_rect(self, page_num: int, coords: tuple) -> str:
        try:
            page = self.doc[page_num]
            pts  = [(v / self.dpi) * 72 for v in coords]
            text = page.get_text("text", clip=fitz.Rect(*pts))
            return " ".join(text.split("\n")).strip()
        except Exception:
            return ""

    def read_zone(self, page_num, coords) -> str:
        """Читает произвольную зону (координаты в системе 300 DPI) без сохранения
        её в self.zones — нужно для резервных зон и живого предпросмотра."""
        if page_num is None or coords is None:
            return ""
        try:
            if not (0 <= page_num < len(self.doc)):
                return ""
        except Exception:
            return ""
        return self._read_rect(page_num, self._scaled(tuple(coords)))

    def read_field(self, page_num: int, key: str) -> str:
        """Возвращает сырой текст зоны поля key со страницы page_num."""
        return self.read_zone(page_num, self.zones.get(key))

    def render_page(self, page_num: int, dpi: int = 150) -> QPixmap:
        """Рендерит страницу. Кэширует последние 6 рендеров."""
        key = (page_num, dpi)
        if key in self._render_cache:
            return self._render_cache[key]
        page = self.doc[page_num]
        mat  = fitz.Matrix(dpi / 72, dpi / 72)
        pix  = page.get_pixmap(matrix=mat)
        img  = QImage(pix.samples, pix.width, pix.height, pix.stride, _IMG_RGB888)
        pm   = QPixmap.fromImage(img)
        if len(self._render_cache) >= 16:
            del self._render_cache[next(iter(self._render_cache))]
        self._render_cache[key] = pm
        return pm

    def page_count(self) -> int:
        return len(self.doc)

    def close(self):
        try:
            self.doc.close()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# Поток извлечения
# ─────────────────────────────────────────────────────────────────────────────

class ExtractionWorker(QObject):
    progress = Signal(int, int)
    finished = Signal(list)
    error    = Signal(str)

    def __init__(self, parser: PdfParser, fields: list, pages: int,
                 fallback_zones: Optional[dict] = None):
        super().__init__()
        self.parser = parser
        self.fields = fields            # список dict {"name","page","clean","fallback",...}
        self.pages  = max(1, int(pages))  # страниц на документ
        self.fallback_zones = dict(fallback_zones or {})
        self._stop  = False

    def run(self):
        try:
            n      = self.parser.page_count()
            groups = list(range(0, n, self.pages))
            total  = len(groups)
            result = []
            for done, start in enumerate(groups, 1):
                if self._stop:
                    break
                row = {"_page": start}
                for f in self.fields:
                    name = f["name"]
                    src  = start + int(f.get("page", 0))
                    if src >= n:
                        row[name] = ""
                        continue
                    clean = f.get("clean", "")
                    post = f.get("post", "")
                    val = process_value(self.parser.read_field(src, name), post, clean)
                    # Резервная зона: если по основной ничего не извлеклось
                    if (not val.strip() and f.get("fallback")
                            and name in self.fallback_zones):
                        raw2 = self.parser.read_zone(src, self.fallback_zones[name])
                        val2 = process_value(raw2, post, clean)
                        if val2.strip():
                            val = val2
                    row[name] = val
                result.append(row)
                self.progress.emit(done, total)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))

    def stop(self):
        self._stop = True

# ─────────────────────────────────────────────────────────────────────────────
# ZoneHandle — перетаскиваемый прямоугольник зоны
# ─────────────────────────────────────────────────────────────────────────────

class ZoneHandle(QGraphicsRectItem):

    def __init__(self, zone_key: str, color: QColor, on_change=None,
                 label: str = "", dashed: bool = False):
        super().__init__()
        self.zone_key   = zone_key
        self._color     = color
        self._on_change = on_change
        self._label     = label or zone_key
        self._dashed    = dashed
        self.setFlags(_GI_MOVABLE | _GI_SELECTABLE | _GI_GEOMCHG)
        self.setAcceptHoverEvents(True)
        self._resizing    = False
        self._resize_edge: Optional[str] = None
        self._drag_start  = None
        self._rect_start  = None
        self._min_size    = 20

    def paint(self, painter: QPainter, option, widget=None):
        r = self.rect()
        c = QColor(self._color)
        painter.fillRect(r, QColor(c.red(), c.green(), c.blue(), 26 if self._dashed else 40))
        pen = QPen(c, 2)
        pen.setStyle(_DASH_PEN if self._dashed else _SOLID_PEN)
        painter.setPen(pen)
        painter.drawRect(int(r.left()), int(r.top()), int(r.width()), int(r.height()))
        hs = 8
        painter.setBrush(QBrush(c))
        painter.setPen(QPen(_NO_PEN))
        for x, y in [
            (r.left(), r.top()), (r.right() - hs, r.top()),
            (r.left(), r.bottom() - hs), (r.right() - hs, r.bottom() - hs),
        ]:
            painter.drawRect(int(x), int(y), hs, hs)
        # Подпись — ярлык над рамкой, чтобы не закрывать текст в зоне.
        self._tag_width()
        painter.setFont(self._TAG_FONT)
        tag = QRectF(r.left(), r.top() - self._TAG_H, self._tag_width(), self._TAG_H)
        painter.setPen(QPen(_NO_PEN))
        painter.setBrush(QBrush(QColor(c.red(), c.green(), c.blue(), 235)))
        painter.drawRect(tag)
        painter.setPen(QPen(_WHITE))
        painter.drawText(tag, _ALIGN_CENTER, self._label)

    _TAG_H = 18
    _TAG_FONT = None

    def _tag_width(self) -> float:
        if ZoneHandle._TAG_FONT is None:
            ZoneHandle._TAG_FONT = QFont("Segoe UI", 9, _BOLD)
        return QFontMetrics(self._TAG_FONT).horizontalAdvance(self._label) + 12

    def boundingRect(self):
        r = self.rect()
        extra = max(0.0, self._tag_width() - r.width())
        return r.adjusted(-2, -self._TAG_H - 2, extra + 2, 2)

    def _edge_at(self, pos: QPointF) -> Optional[str]:
        r, mg = self.rect(), 10
        on_l = abs(pos.x() - r.left())   < mg
        on_r = abs(pos.x() - r.right())  < mg
        on_t = abs(pos.y() - r.top())    < mg
        on_b = abs(pos.y() - r.bottom()) < mg
        if on_l and on_t: return "tl"
        if on_r and on_t: return "tr"
        if on_l and on_b: return "bl"
        if on_r and on_b: return "br"
        if on_r: return "r"
        if on_l: return "l"
        if on_t: return "t"
        if on_b: return "b"
        return None

    def hoverMoveEvent(self, event):
        edge = self._edge_at(event.pos())
        self.setCursor(QCursor(_CURSOR_MAP.get(edge, _ARROW_CUR) if edge else _SIZEALL_CUR))

    def mousePressEvent(self, event):
        if event.button() == _LEFT_BTN:
            edge = self._edge_at(event.pos())
            if edge:
                self._resizing    = True
                self._resize_edge = edge
                self._drag_start  = event.scenePos()
                self._rect_start  = self.rect()
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._resizing and self._resize_edge:
            delta = event.scenePos() - self._drag_start
            r     = QRectF(self._rect_start) if self._rect_start is not None else self.rect()
            e     = self._resize_edge
            if "l" in e: r.setLeft(r.left()   + delta.x())
            if "r" in e: r.setRight(r.right()  + delta.x())
            if "t" in e: r.setTop(r.top()      + delta.y())
            if "b" in e: r.setBottom(r.bottom() + delta.y())
            if r.width() > self._min_size and r.height() > self._min_size:
                self.setRect(r.normalized())
                if self._on_change:
                    self._on_change(self.zone_key)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._resizing:
            self._resizing = False
            if self._on_change:
                self._on_change(self.zone_key)
            event.accept()
            return
        super().mouseReleaseEvent(event)
        if self._on_change:
            self._on_change(self.zone_key)

# ─────────────────────────────────────────────────────────────────────────────
# ZoneEditorWidget — динамический редактор зон
# ─────────────────────────────────────────────────────────────────────────────

# Инлайн-стили тулбара редактора зон строятся из активной темы, чтобы редактор
# перекрашивался вместе с остальным интерфейсом при смене темы.

def _tb_btn_ss() -> str:
    return (
        f"QPushButton{{background:{T.chrome_raised};color:{T.on_chrome};border:1px solid {T.chrome_border};"
        "border-radius:4px;padding:3px 10px;font-size:9pt;}"
        f"QPushButton:hover{{background:{T.accent};color:#fff;border-color:{T.accent};}}"
        f"QPushButton:disabled{{background:{T.chrome};color:{T.chrome_border};}}"
    )


def _tb_lbl_ss() -> str:
    return f"color:{T.on_chrome}; background:{T.chrome}; font-size:9pt;"


def _tb_accent_lbl_ss() -> str:
    return f"color:{T.accent_bright}; background:{T.chrome}; font-size:9pt; font-weight:bold;"


def _slider_ss() -> str:
    return (
        f"QSlider::groove:horizontal{{background:{T.chrome_raised};height:4px;border-radius:2px;}}"
        f"QSlider::handle:horizontal{{background:{T.accent};width:14px;height:14px;margin:-5px 0;border-radius:7px;}}"
        f"QSlider::handle:horizontal:hover{{background:{T.accent_hi};}}"
        f"QSlider::sub-page:horizontal{{background:{T.accent};border-radius:2px;}}"
    )


def _cb_zone_ss(hex_c: str) -> str:
    return (
        f"QCheckBox{{color:{hex_c};background:{T.chrome};font-weight:bold;margin:0 6px;font-size:9pt;}}"
        f"QCheckBox::indicator{{width:14px;height:14px;border:2px solid {hex_c};border-radius:3px;background:{T.chrome_deep};}}"
        f"QCheckBox::indicator:checked{{background:{hex_c};border-color:{hex_c};}}"
    )


class ZoneEditorWidget(QWidget):
    # payload: {"zones": {name: coords}, "fallback": {name: coords}}
    zones_changed = Signal(dict)

    _SEP_H = 10  # высота разделителя между страницами в стеке

    def __init__(self, parent=None):
        super().__init__(parent)
        self._parser: Optional[PdfParser] = None
        self._render_dpi       = 150
        self._pages            = 1          # страниц на документ (ppd)
        self._group_start      = 0          # индекс первой страницы текущего документа
        self._fields: list     = []
        self._colors: dict     = {}
        self._default_zones: dict = {}
        self._entries: list    = []         # записи зон (основные + резервные)
        self._entry_by_key: dict = {}
        self._zone_coords: dict = {}        # {key: (x0,y0,x1,y1)} в 300 DPI — все зоны
        self._pix_scale        = 1.0
        self._page_offsets: list = []       # y-смещение каждой страницы группы (или None)
        self._zone_items: dict = {}
        self._zone_checks: dict = {}
        self._coord_labels: dict = {}
        self._preview_key: Optional[str] = None
        self._setup_ui()

    # ── UI ────────────────────────────────────────────────────────────────
    def _setup_ui(self):
        # Ссылки на элементы «хрома», перекрашиваемые при смене темы.
        self._tb_labels: list = []   # обычные подписи тулбара
        self._tb_buttons: list = []  # кнопки тулбара
        self._tb_seps: list = []     # вертикальные разделители тулбара

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._tb_top = top = QWidget()
        top.setStyleSheet(f"background:{T.chrome}; border-bottom:2px solid {T.accent};")
        top_lay = QHBoxLayout(top)
        top_lay.setContentsMargins(6, 4, 6, 4)
        top_lay.setSpacing(4)

        def _lbl(text):
            l = QLabel(text)
            l.setStyleSheet(_tb_lbl_ss())
            top_lay.addWidget(l)
            self._tb_labels.append(l)
            return l

        def _btn(text, slot, w=None):
            b = QPushButton(text)
            b.setStyleSheet(_tb_btn_ss())
            b.setCursor(QCursor(_ARROW_CUR))
            if w:
                b.setFixedSize(w, 28)
            b.clicked.connect(slot)
            top_lay.addWidget(b)
            self._tb_buttons.append(b)
            return b

        def _sep():
            s = QFrame(); s.setFrameShape(_VLINE)
            s.setStyleSheet(f"color:{T.chrome_border};")
            top_lay.addWidget(s)
            self._tb_seps.append(s)
            return s

        _lbl(tr("  Документ: "))
        self._page_btn_prev = _btn("◀", self._page_prev, 28)
        self._page_lbl = QLabel("1 / 1")
        self._page_lbl.setFixedWidth(56)
        self._page_lbl.setAlignment(_ALIGN_CENTER)
        self._page_lbl.setStyleSheet(_tb_accent_lbl_ss())
        top_lay.addWidget(self._page_lbl)
        self._page_btn_next = _btn("▶", self._page_next, 28)

        _sep()
        _lbl(tr("  Масштаб: "))
        self._zoom_slider = QSlider(_HORIZONTAL)
        self._zoom_slider.setRange(50, 300)
        self._zoom_slider.setValue(100)
        self._zoom_slider.setFixedWidth(120)
        self._zoom_slider.setStyleSheet(_slider_ss())
        self._zoom_slider.valueChanged.connect(self._on_zoom)
        top_lay.addWidget(self._zoom_slider)
        self._zoom_label = QLabel("100%")
        self._zoom_label.setStyleSheet(_tb_accent_lbl_ss())
        self._zoom_label.setFixedWidth(42)
        top_lay.addWidget(self._zoom_label)

        _sep()

        # Динамический контейнер чекбоксов зон (в прокрутке по горизонтали)
        self._cb_scroll = QScrollArea()
        self._cb_scroll.setWidgetResizable(True)
        self._cb_scroll.setFrameShape(_NO_FRAME)
        self._cb_scroll.setFixedHeight(34)
        self._cb_scroll.setStyleSheet(f"background:{T.chrome}; border:none;")
        self._cb_host = QWidget()
        self._cb_host.setStyleSheet(f"background:{T.chrome};")
        self._cb_lay = QHBoxLayout(self._cb_host)
        self._cb_lay.setContentsMargins(0, 0, 0, 0)
        self._cb_lay.setSpacing(0)
        self._cb_scroll.setWidget(self._cb_host)
        top_lay.addWidget(self._cb_scroll, 1)

        self._btn_preview_all = QPushButton(tr("👁 Текст всех зон"))
        self._btn_preview_all.setFixedHeight(28)
        self._btn_preview_all.setStyleSheet(_tb_btn_ss())
        self._btn_preview_all.clicked.connect(self._preview_all)
        self._tb_buttons.append(self._btn_preview_all)
        top_lay.addWidget(self._btn_preview_all)

        btn_reset = QPushButton(tr("↺ Сбросить зоны"))
        btn_reset.setFixedHeight(28)
        btn_reset.setStyleSheet(_tb_btn_ss())
        btn_reset.clicked.connect(self._reset_zones)
        self._tb_buttons.append(btn_reset)
        top_lay.addWidget(btn_reset)
        layout.addWidget(top)

        self._hint = hint = QLabel(
            tr("  🖱  Перетащите прямоугольник — перемещение.  Тяните за угол/край — размер.  Колёсико — масштаб.  Пунктирная рамка — резервная зона.  Внизу — живой предпросмотр текста активной зоны.")
        )
        hint.setStyleSheet(self._hint_ss())
        layout.addWidget(hint)

        self._scene = QGraphicsScene(self)
        self._view  = QGraphicsView(self._scene)
        self._view.setRenderHint(_ANTIALIASING)
        self._view.setDragMode(_SCROLL_DRAG)
        self._view.setTransformationAnchor(_ANCHOR_MOUSE)
        self._view.setResizeAnchor(_ANCHOR_MOUSE)
        self._view.setStyleSheet(f"background: {T.canvas_bg}; border: none;")
        self._view.wheelEvent = self._wheel_event
        layout.addWidget(self._view, 1)

        # Живой предпросмотр текста активной зоны
        self._preview_lbl = QLabel(tr("Наведите/сдвиньте зону, чтобы увидеть распознанный текст."))
        self._preview_lbl.setWordWrap(True)
        self._preview_lbl.setMinimumHeight(38)
        self._set_preview(tr("Наведите/сдвиньте зону, чтобы увидеть распознанный текст."))
        layout.addWidget(self._preview_lbl)

        # Таймер антидребезга предпросмотра
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(90)
        self._preview_timer.timeout.connect(self._do_preview)

        self._coord_host = QWidget()
        self._coord_host.setStyleSheet(f"background:{T.chrome_deep}; border-top:1px solid {T.chrome_border};")
        self._coord_outer = QHBoxLayout(self._coord_host)
        self._coord_outer.setContentsMargins(0, 0, 0, 0)
        self._coord_outer.setSpacing(0)
        self._coord_scroll = coord_scroll = QScrollArea()
        coord_scroll.setWidgetResizable(True)
        coord_scroll.setFrameShape(_NO_FRAME)
        coord_scroll.setFixedHeight(30)
        coord_scroll.setStyleSheet(f"background:{T.chrome_deep}; border:none;")
        self._coord_inner = inner = QWidget()
        inner.setStyleSheet(f"background:{T.chrome_deep};")
        self._coord_lay = QHBoxLayout(inner)
        self._coord_lay.setContentsMargins(10, 4, 10, 4)
        self._coord_lay.setSpacing(16)
        coord_scroll.setWidget(inner)
        self._coord_outer.addWidget(coord_scroll)
        layout.addWidget(self._coord_host)

        register_retheme(self.retheme)

    def _hint_ss(self) -> str:
        return (
            f"background:{T.chrome}; color:{T.on_chrome_dim}; padding:5px 10px;"
            f" font-size:9pt; border-bottom:1px solid {T.chrome_raised};"
        )

    def retheme(self):
        """Перекрашивает редактор зон под активную тему."""
        self._tb_top.setStyleSheet(f"background:{T.chrome}; border-bottom:2px solid {T.accent};")
        for l in self._tb_labels:
            l.setStyleSheet(_tb_lbl_ss())
        for b in self._tb_buttons:
            b.setStyleSheet(_tb_btn_ss())
        for s in self._tb_seps:
            s.setStyleSheet(f"color:{T.chrome_border};")
        self._page_lbl.setStyleSheet(_tb_accent_lbl_ss())
        self._zoom_label.setStyleSheet(_tb_accent_lbl_ss())
        self._zoom_slider.setStyleSheet(_slider_ss())
        self._cb_scroll.setStyleSheet(f"background:{T.chrome}; border:none;")
        self._cb_host.setStyleSheet(f"background:{T.chrome};")
        self._hint.setStyleSheet(self._hint_ss())
        self._view.setStyleSheet(f"background: {T.canvas_bg}; border: none;")
        self._coord_host.setStyleSheet(f"background:{T.chrome_deep}; border-top:1px solid {T.chrome_border};")
        self._coord_scroll.setStyleSheet(f"background:{T.chrome_deep}; border:none;")
        self._coord_inner.setStyleSheet(f"background:{T.chrome_deep};")
        # Чекбоксы зон, подписи координат, заглушка и разделители перерисовываются.
        self._rebuild_controls()
        self._render_page()

    # ── Загрузка / перестроение ──────────────────────────────────────────────
    def load(self, parser, fields, zones, fallback_zones, colors,
             default_zones, pages, page=0):
        self._parser        = parser
        self._fields        = [dict(f) for f in fields]
        self._colors        = dict(colors)
        self._default_zones = dict(default_zones)
        self._pages         = max(1, int(pages))
        self._group_start   = (int(page) // self._pages) * self._pages
        self._build_entries(dict(zones), dict(fallback_zones or {}))
        if parser is not None:
            parser.set_zones(self._primary_map())
        self._rebuild_controls()
        self._update_page_label()
        self._render_page()

    def _build_entries(self, zones: dict, fallback_zones: dict):
        """Строит список записей зон (основные + резервные) и их координаты."""
        self._entries = []
        self._entry_by_key = {}
        self._zone_coords = {}
        for i, f in enumerate(self._fields):
            name = f["name"]
            page = int(f.get("page", 0))
            primary_coords = tuple(zones.get(name) or self._default_zones.get(name)
                                   or _default_zone(i))
            p_entry = {
                "key": name, "field": name, "label": name, "page": page,
                "is_fallback": False, "clean": f.get("clean", ""), "post": f.get("post", ""),
                "regex": f.get("regex", ""), "type": f.get("type"),
            }
            self._entries.append(p_entry)
            self._entry_by_key[name] = p_entry
            self._zone_coords[name] = primary_coords

            if f.get("fallback"):
                fkey = name + "\x00fb"
                fcoords = tuple(fallback_zones.get(name)
                                or _default_fallback_zone(primary_coords))
                f_entry = {
                    "key": fkey, "field": name, "label": tr("{name} · резерв", name=name),
                    "page": page, "is_fallback": True,
                    "clean": f.get("clean", ""), "post": f.get("post", ""), "regex": f.get("regex", ""),
                    "type": f.get("type"),
                }
                self._entries.append(f_entry)
                self._entry_by_key[fkey] = f_entry
                self._zone_coords[fkey] = fcoords

    def _primary_map(self) -> dict:
        return {e["field"]: self._zone_coords[e["key"]]
                for e in self._entries if not e["is_fallback"] and e["key"] in self._zone_coords}

    def _fallback_map(self) -> dict:
        return {e["field"]: self._zone_coords[e["key"]]
                for e in self._entries if e["is_fallback"] and e["key"] in self._zone_coords}

    def _entry_color(self, entry) -> QColor:
        return self._colors.get(entry["field"], QColor(100, 100, 255, 120))

    def _rebuild_controls(self):
        """Пересоздаёт чекбоксы зон и метки координат под текущий набор записей."""
        _clear_layout(self._cb_lay)
        _clear_layout(self._coord_lay)
        self._zone_checks = {}
        self._coord_labels = {}
        for entry in self._entries:
            key = entry["key"]
            color = self._entry_color(entry)
            hex_c = f"#{color.red():02x}{color.green():02x}{color.blue():02x}"
            cb_text = (entry["field"] + tr("·рез")) if entry["is_fallback"] else entry["field"]
            cb = QCheckBox(cb_text)
            cb.setChecked(True)
            cb.setStyleSheet(_cb_zone_ss(hex_c))
            cb.toggled.connect(lambda checked, k=key: self._toggle_zone(k, checked))
            self._cb_lay.addWidget(cb)
            self._zone_checks[key] = cb

            marker = "◌" if entry["is_fallback"] else "●"
            lbl = QLabel(f"{marker} {entry['label']}: —")
            lbl.setStyleSheet(f"color:{hex_c}; font-size:9pt; font-weight:bold; background:{T.chrome_deep};")
            self._coord_lay.addWidget(lbl)
            self._coord_labels[key] = lbl
        self._cb_lay.addStretch()
        self._coord_lay.addStretch()

    # ── Рендеринг ────────────────────────────────────────────────────────────
    def _render_page(self):
        self._scene.clear()
        self._zone_items = {}
        self._page_offsets = []
        if self._parser is None:
            t = self._scene.addText(tr("Откройте PDF-файл, чтобы разместить зоны."))
            t.setDefaultTextColor(QColor(T.placeholder))
            t.setPos(20, 20)
            self._set_preview(tr("Наведите/сдвиньте зону, чтобы увидеть распознанный текст."))
            return
        zoom = self._zoom_slider.value() / 100.0
        dpi  = max(40, min(600, int(self._render_dpi * zoom)))
        self._pix_scale = dpi / self._parser.BASE_DPI
        n = self._parser.page_count()

        y = 0
        max_w = 1
        first = True
        for k in range(self._pages):
            pnum = self._group_start + k
            if pnum >= n:
                self._page_offsets.append(None)
                continue
            pix = self._parser.render_page(pnum, dpi=dpi)
            if not first:
                self._scene.addRect(
                    QRectF(0, y, max(max_w, pix.width()), self._SEP_H),
                    QPen(_NO_PEN), QBrush(QColor(T.placeholder)),
                )
                y += self._SEP_H
            item = QGraphicsPixmapItem(pix)
            item.setPos(0, y)
            self._scene.addItem(item)
            self._page_offsets.append(y)
            y += pix.height()
            max_w = max(max_w, pix.width())
            first = False

        if y <= 0:
            # ни одной страницы для этой группы
            t = self._scene.addText(tr("Для этого документа нет страниц."))
            t.setDefaultTextColor(QColor(T.placeholder))
            t.setPos(20, 20)
            return
        self._scene.setSceneRect(QRectF(0, 0, max_w, y))
        self._draw_zones()
        self._set_preview(tr("Наведите/сдвиньте зону, чтобы увидеть распознанный текст."))

    def _offset_for(self, entry):
        p = entry["page"]
        if 0 <= p < len(self._page_offsets):
            return self._page_offsets[p]
        return None

    def _draw_zones(self):
        if self._parser is None:
            return
        for entry in self._entries:
            key = entry["key"]
            off = self._offset_for(entry)
            coords = self._zone_coords.get(key)
            if off is None or coords is None:
                continue
            x0, y0, x1, y1 = [v * self._pix_scale for v in coords]
            y0 += off
            y1 += off
            item = ZoneHandle(key, self._entry_color(entry),
                              on_change=self._on_zone_moved, label=entry["label"],
                              dashed=entry["is_fallback"])
            item.setRect(QRectF(x0, y0, x1 - x0, y1 - y0))
            self._scene.addItem(item)
            self._zone_items[key] = item
            if key in self._zone_checks:
                item.setVisible(self._zone_checks[key].isChecked())
        self._update_coord_labels()

    def _on_zone_moved(self, key: str):
        self._update_coord_labels()
        self._preview_key = key
        self._preview_timer.start()
        self._emit_zones()

    def _item_coords(self, key):
        """Координаты зоны key в системе 300 DPI (с учётом смещения её страницы)."""
        item = self._zone_items.get(key)
        entry = self._entry_by_key.get(key)
        if item is None or entry is None:
            return None
        off = self._offset_for(entry) or 0
        r, p, ps = item.rect(), item.pos(), self._pix_scale
        return (
            int((r.left()   + p.x()) / ps),
            int((r.top()    + p.y() - off) / ps),
            int((r.right()  + p.x()) / ps),
            int((r.bottom() + p.y() - off) / ps),
        )

    def _update_coord_labels(self):
        for key, item in self._zone_items.items():
            c = self._item_coords(key)
            if c is None:
                continue
            self._zone_coords[key] = c  # держим карту в актуальном виде
            entry = self._entry_by_key.get(key)
            marker = "◌" if (entry and entry["is_fallback"]) else "●"
            label = entry["label"] if entry else key
            if key in self._coord_labels:
                self._coord_labels[key].setText(
                    f"{marker} {label}: ({c[0]},{c[1]})–({c[2]},{c[3]})")

    def _collect_zones(self) -> dict:
        out = {}
        for key in self._zone_items:
            c = self._item_coords(key)
            if c is not None:
                out[key] = c
        return out

    def _emit_zones(self):
        collected = self._collect_zones()
        self._zone_coords.update(collected)
        if self._parser is not None:
            self._parser.set_zones(self._primary_map())
        self.zones_changed.emit({"zones": self._primary_map(),
                                 "fallback": self._fallback_map()})

    def _toggle_zone(self, key: str, visible: bool):
        if key in self._zone_items:
            self._zone_items[key].setVisible(visible)

    def _reset_zones(self):
        if self._parser is None or not self._default_zones:
            return
        base_primary = dict(self._default_zones)
        base_fallback = {e["field"]: _default_fallback_zone(base_primary.get(e["field"],
                         _default_zone(0))) for e in self._entries if e["is_fallback"]}
        self._build_entries(base_primary, base_fallback)
        self._parser.set_zones(self._primary_map())
        self._render_page()
        self.zones_changed.emit({"zones": self._primary_map(),
                                 "fallback": self._fallback_map()})

    # ── Живой предпросмотр ────────────────────────────────────────────────────
    def _set_preview(self, text: str, color: str = None):
        if color is None:
            color = T.on_chrome_dim
        self._preview_lbl.setStyleSheet(
            f"background:{T.preview_bg}; border-top:2px solid {T.accent}; padding:6px 12px;"
            f" font-family:'Consolas','Courier New',monospace; font-size:9pt; color:{color};"
        )
        self._preview_lbl.setText(text)

    def _read_entry(self, entry):
        """Возвращает (raw, cleaned, ok) для зоны записи на её странице группы."""
        coords = self._zone_coords.get(entry["key"])
        pnum = self._group_start + entry["page"]
        raw = self._parser.read_zone(pnum, coords) if self._parser else ""
        has_proc = bool(entry["clean"] or entry.get("post"))
        cleaned = process_value(raw, entry.get("post", ""), entry["clean"]) if has_proc else raw
        ok = _field_valid(cleaned, resolve_check_pattern(entry))
        return raw, cleaned, ok

    def _do_preview(self):
        key = self._preview_key
        entry = self._entry_by_key.get(key) if key else None
        if entry is None or self._parser is None:
            return
        raw, cleaned, ok = self._read_entry(entry)
        shown_raw = raw if raw.strip() else "—"
        text = tr("▸ {label} · стр {v}:  «{shown_raw}»", label=entry['label'], v=entry['page'] + 1, shown_raw=shown_raw)
        if entry["clean"] or entry.get("post"):
            text += f"   →   «{cleaned or '—'}»"
        color = T.on_chrome_dim
        if resolve_check_pattern(entry):
            if ok and cleaned.strip():
                text += tr("   ✓ проверка пройдена")
                color = T.ok_text
            else:
                text += tr("   ✗ не совпало с проверкой")
                color = T.err_text
        elif not raw.strip():
            color = T.warn_text
        self._set_preview(text, color)

    def _preview_all(self):
        if self._parser is None or not self._entries:
            self._set_preview(tr("Нет открытого PDF или зон для предпросмотра."), T.warn_text)
            return
        lines = []
        for entry in self._entries:
            if self._offset_for(entry) is None:
                continue
            raw, cleaned, ok = self._read_entry(entry)
            val = cleaned if (entry["clean"] or entry.get("post")) else raw
            val = val.strip() or "—"
            if len(val) > 70:
                val = val[:67] + "…"
            mark = "◌" if entry["is_fallback"] else "●"
            status = ""
            if resolve_check_pattern(entry):
                status = "  ✓" if (ok and cleaned.strip()) else "  ✗"
            lines.append(tr("{mark} {label} (стр {v}): «{val}»{status}", mark=mark, label=entry['label'], v=entry['page'] + 1, val=val, status=status))
        self._preview_lbl.setWordWrap(True)
        self._set_preview("\n".join(lines) if lines else tr("Для этого документа нет зон."))

    # ── Навигация ────────────────────────────────────────────────────────────
    def _page_prev(self):
        if self._group_start >= self._pages:
            self._group_start -= self._pages
            self._update_page_label()
            self._render_page()

    def _page_next(self):
        if not self._parser:
            return
        total = self._parser.page_count()
        if self._group_start + self._pages < total:
            self._group_start += self._pages
            self._update_page_label()
            self._render_page()

    def _update_page_label(self):
        total = self._parser.page_count() if self._parser else 1
        cnt = (total + self._pages - 1) // self._pages
        num = self._group_start // self._pages + 1
        self._page_lbl.setText(f"{num} / {max(1, cnt)}")
        self._page_btn_prev.setEnabled(self._group_start >= self._pages)
        self._page_btn_next.setEnabled(self._group_start + self._pages < total)

    def _on_zoom(self, val: int):
        self._zoom_label.setText(f"{val}%")
        self._render_page()

    def _wheel_event(self, event):
        factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
        self._view.scale(factor, factor)

# ─────────────────────────────────────────────────────────────────────────────
# Автодополнение в ячейке (для диалога сопоставления)
# ─────────────────────────────────────────────────────────────────────────────

class _AcEditor(QWidget):
    committed = Signal(str)
    cancelled = Signal()

    def __init__(self, candidates, initial_text, parent=None):
        super().__init__(parent)
        self._candidates = candidates
        self._committing = False

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        self._ed = QLineEdit(initial_text)
        self._ed.setStyleSheet(
            f"QLineEdit{{background:{T.chrome_deep};color:{T.on_chrome};border:2px solid {T.accent};font-size:10pt;padding:2px 4px;}}"
        )
        self._ed.selectAll()
        lay.addWidget(self._ed)

        self._list = QListWidget()
        self._list.setStyleSheet(
            f"QListWidget{{background:{T.chrome};color:{T.on_chrome};border:2px solid {T.accent};border-top:none;font-size:10pt;}}"
            "QListWidget::item{padding:3px 6px;}"
            f"QListWidget::item:hover{{background:{T.chrome_raised};}}"
            f"QListWidget::item:selected{{background:{T.accent};color:white;}}"
        )
        self._list.setFocusPolicy(_NO_FOCUS)
        self._list.setFixedHeight(0)
        lay.addWidget(self._list)

        self._ed.textChanged.connect(self._fill_list)
        self._ed.returnPressed.connect(self._on_enter)
        self._ed.installEventFilter(self)
        self._list.itemClicked.connect(self._on_pick)
        self.setFocusProxy(self._ed)
        self._fill_list(initial_text)

    def _fill_list(self, text):
        self._list.clear()
        lo = text.lower()
        matches = [c for c in self._candidates if lo in c.lower()] if lo else self._candidates[:40]
        if not matches:
            self._list.setFixedHeight(0)
            return
        for m in matches[:50]:
            self._list.addItem(m)
        row_h = self._list.sizeHintForRow(0) + 2
        self._list.setFixedHeight(min(len(matches), 8) * row_h + 4)
        self.adjustSize()
        if self.parent():
            self.parent().update()

    def _on_pick(self, list_item):
        self._ed.setText(list_item.text())
        self._commit()

    def _on_enter(self):
        sel = self._list.currentItem()
        if sel and self._list.height() > 0:
            self._ed.setText(sel.text())
        self._commit()

    def eventFilter(self, obj, event):
        if obj is self._ed:
            t = event.type()
            if t == QEvent.Type.KeyPress:
                key = event.key()
                if key == _KEY_DOWN:
                    self._list.setCurrentRow(min(self._list.currentRow() + 1, self._list.count() - 1))
                    return True
                if key == _KEY_UP:
                    self._list.setCurrentRow(max(self._list.currentRow() - 1, 0))
                    return True
                if key == _KEY_ESC:
                    self.cancelled.emit()
                    return True
            if t == QEvent.Type.FocusOut:
                QTimer.singleShot(0, self._commit)
        return super().eventFilter(obj, event)

    def _commit(self):
        if self._committing:
            return
        self._committing = True
        self.committed.emit(self._ed.text().strip())

    def current_text(self):
        return self._ed.text().strip()


class _AutoCompleteDelegate(QObject):
    def __init__(self, tree, candidates, col, parent=None):
        super().__init__(parent)
        self._tree       = tree
        self._candidates = candidates
        self._col        = col
        self._editor: Optional[_AcEditor] = None
        self._item       = None
        tree.itemDoubleClicked.connect(self._on_double_click)

    def _on_double_click(self, item, col):
        if col != self._col:
            return
        self._close_editor()
        self._item = item
        rect   = self._tree.visualItemRect(item)
        header = self._tree.header()
        x = sum(header.sectionSize(c) for c in range(col))
        w = header.sectionSize(col)
        ed = _AcEditor(self._candidates, item.text(col), parent=self._tree.viewport())
        ed.move(x, rect.top())
        ed.setFixedWidth(w)
        ed.adjustSize()
        ed.show()
        ed.setFocus()
        self._editor = ed
        ed.committed.connect(self._on_committed)
        ed.cancelled.connect(self._close_editor)

    def _on_committed(self, val):
        item = self._item
        if item is not None and val:
            item.setText(self._col, val)
            dlg = self.parent()
            if dlg is not None and hasattr(dlg, "_sync_result"):
                dlg._sync_result(item, val)
        self._close_editor()

    def _close_editor(self):
        if self._editor is not None:
            self._editor.hide()
            self._editor.deleteLater()
            self._editor = None
        self._item = None

# ─────────────────────────────────────────────────────────────────────────────
# MatchDialog — сопоставление извлечённых значений со словарём
# ─────────────────────────────────────────────────────────────────────────────

class MatchDialog(QDialog):
    """Нечёткое сопоставление значений одного поля с загруженным словарём."""

    new_entry = Signal(str)

    def __init__(self, waybills, field_name, candidates, pages=1,
                 pdf_path="", parser=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Сопоставление: {field_name}", field_name=field_name))
        self.resize(920, 600)
        self.setStyleSheet(current_qss())
        self._waybills   = waybills
        self._field      = field_name
        self._candidates = list(candidates)
        self._pages      = max(1, int(pages))
        self._pdf_path   = pdf_path
        self._parser     = parser
        self._result: list[str] = []
        self._setup_ui()
        self._populate()
        self._delegate = _AutoCompleteDelegate(self._tree, self._candidates, col=2, parent=self)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        self._tree = QTreeWidget()
        self._tree.setColumnCount(5)
        self._tree.setHeaderLabels(["#", tr("Распознано"), tr("Лучшее совпадение"), tr("Сходство"), tr("Использовать")])
        self._tree.header().setSectionResizeMode(_RESIZE_CONT)
        self._tree.setAlternatingRowColors(True)
        self._tree.setSelectionMode(_EXT_SEL)
        self._tree.setEditTriggers(_NO_EDIT)
        layout.addWidget(self._tree, 1)

        ctrl = QHBoxLayout()
        ctrl.setSpacing(8)
        ctrl.addWidget(QLabel(tr("Порог автозамены (%):")))
        self._threshold = NoWheelSpinBox()
        self._threshold.setRange(0, 100)
        self._threshold.setValue(int(AUTO_MATCH_THRESHOLD * 100))
        self._threshold.setFixedWidth(70)
        ctrl.addWidget(self._threshold)

        btn_all = QPushButton(tr("✔ Применить ко всем"))
        btn_all.setObjectName("primary")
        btn_all.clicked.connect(self._apply_all)
        ctrl.addWidget(btn_all)

        btn_sel = QPushButton(tr("✔ Применить к выделенным"))
        btn_sel.clicked.connect(self._apply_selected)
        ctrl.addWidget(btn_sel)

        ctrl.addStretch()

        btn_view = QPushButton(tr("👁 Просмотреть документ"))
        btn_view.clicked.connect(self._view_waybill)
        ctrl.addWidget(btn_view)

        self._btn_add = QPushButton(tr("➕ Добавить в список"))
        self._btn_add.setToolTip(tr("Добавить выделенное значение в словарь"))
        self._btn_add.clicked.connect(self._add_selected_to_list)
        ctrl.addWidget(self._btn_add)

        layout.addLayout(ctrl)

        bb = QDialogButtonBox(_DB_OK | _DB_CANCEL)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def _populate(self):
        self._tree.clear()
        self._result = []
        for i, wb in enumerate(self._waybills):
            value = wb.get(self._field, "")
            best, score = _best_match(value, self._candidates)
            pct  = int(score * 100)
            use  = best if score >= AUTO_MATCH_THRESHOLD else value
            item = QTreeWidgetItem([str(i + 1), value, best, f"{pct}%", use])
            item.setFlags(item.flags() | _ITEM_EDITABLE)
            self._tree.addTopLevelItem(item)
            self._result.append(use)

    def _sync_result(self, item, val):
        idx = self._tree.indexOfTopLevelItem(item)
        if 0 <= idx < len(self._result):
            self._result[idx] = val
            item.setText(4, val)
        is_new = val and val not in self._candidates
        self._btn_add.setEnabled(bool(is_new))

    def _add_selected_to_list(self):
        sel = self._tree.selectedItems()
        items = sel if sel else [
            self._tree.topLevelItem(i)
            for i in range(self._tree.topLevelItemCount())
            if self._tree.topLevelItem(i) is not None
        ]
        new_vals = [
            it.text(4).strip() for it in items
            if it.text(4).strip() and it.text(4).strip() not in self._candidates
        ]
        if not new_vals:
            _info(self, tr("Все значения из колонки «Использовать» уже в словаре."), tr("Нет новых"))
            return
        msg = tr("Добавить в словарь:\n") + "\n".join(f"  • {v}" for v in new_vals)
        if not _ask(self, msg, tr("Подтвердите")):
            return
        for v in new_vals:
            self._candidates.append(v)
            self._delegate._candidates = self._candidates
            self.new_entry.emit(v)
        self._btn_add.setEnabled(False)

    def _apply_all(self):
        thr = self._threshold.value() / 100.0
        for i in range(self._tree.topLevelItemCount()):
            item = self._tree.topLevelItem(i)
            if item is None:
                continue
            try:
                score = int(item.text(3).rstrip("%")) / 100.0
            except ValueError:
                score = 0.0
            if score >= thr:
                best = item.text(2)
                item.setText(4, best)
                self._result[i] = best

    def _apply_selected(self):
        selected = self._tree.selectedItems()
        if not selected:
            _info(self, tr("Выделите строки для применения."), tr("Подсказка"))
            return
        thr = self._threshold.value() / 100.0
        for item in selected:
            i = self._tree.indexOfTopLevelItem(item)
            if i < 0:
                continue
            try:
                score = int(item.text(3).rstrip("%")) / 100.0
            except ValueError:
                score = 0.0
            if score >= thr:
                best = item.text(2)
                item.setText(4, best)
                if i < len(self._result):
                    self._result[i] = best

    def _view_waybill(self):
        if not self._pdf_path:
            _warn(self, tr("PDF не открыт."))
            return
        sel = self._tree.selectedItems()
        if not sel:
            _warn(self, tr("Выберите строку."))
            return
        row = self._tree.indexOfTopLevelItem(sel[0])
        total = self._parser.page_count() if self._parser else 0
        try:
            start = (self._waybills[row].get("_page")
                     if 0 <= row < len(self._waybills) else None)
            _open_row_document(self._pdf_path, row, self._pages, total, start)
        except ValueError:
            _warn(self, tr("Страница не существует."))
        except Exception as e:
            _error(self, tr("Не удалось открыть документ:\n{e}", e=e))

    def get_results(self) -> list[str]:
        return [
            self._tree.topLevelItem(i).text(4)
            for i in range(self._tree.topLevelItemCount())
            if self._tree.topLevelItem(i) is not None
        ]

# ─────────────────────────────────────────────────────────────────────────────
# EditCellDialog / CompressDialog / Card
# ─────────────────────────────────────────────────────────────────────────────

class EditCellDialog(QDialog):
    def __init__(self, col_name: str, current_value: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Редактирование"))
        self.setFixedSize(400, 160)
        self.setStyleSheet(current_qss())

        layout = QVBoxLayout(self)
        hdr = QFrame()
        hdr.setFixedHeight(42)
        hdr.setStyleSheet(f"background:{C_ACCENT}; border-radius:0px;")
        hdr_lay = QHBoxLayout(hdr)
        lbl = QLabel(f"✏  {col_name}")
        lbl.setStyleSheet("color:white; font-weight:bold; font-size:10pt;")
        hdr_lay.addWidget(lbl)
        layout.addWidget(hdr)

        body  = QWidget()
        b_lay = QFormLayout(body)
        b_lay.setContentsMargins(16, 12, 16, 8)
        self._entry = QLineEdit(current_value)
        self._entry.selectAll()
        b_lay.addRow(tr("Новое значение:"), self._entry)
        layout.addWidget(body)

        bb = QDialogButtonBox(_DB_SAVE | _DB_CANCEL)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        self._entry.returnPressed.connect(self.accept)
        self._entry.setFocus()

    def value(self) -> str:
        return self._entry.text().strip()


class CompressDialog(QDialog):
    def __init__(self, title: str, with_limit: bool = False, with_split: bool = False, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self._with_limit = with_limit
        self._with_split = with_split

        h = 240 + (50 if with_limit else 0) + (40 if with_split else 0)
        self.setFixedSize(420, h)
        self.setStyleSheet(current_qss())

        layout = QFormLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)

        self._dpi = NoWheelComboBox()
        self._dpi.addItems(["72", "96", "120", "150", "200"])
        self._dpi.setCurrentText("150")
        layout.addRow(tr("DPI для сжатия:"), self._dpi)

        self._quality = NoWheelSpinBox()
        self._quality.setRange(1, 100)
        self._quality.setValue(85)
        layout.addRow(tr("Качество JPEG (1–100):"), self._quality)

        self._grayscale = QCheckBox(tr("Оттенки серого (меньше файл, для ч/б документов)"))
        self._grayscale.setChecked(True)
        layout.addRow(tr("Режим:"), self._grayscale)

        if with_limit:
            self._max_mb = NoWheelDoubleSpinBox()
            self._max_mb.setRange(0, 9999)
            self._max_mb.setValue(50)
            self._max_mb.setSuffix(tr(" МБ"))
            layout.addRow(tr("Макс. размер части:"), self._max_mb)
            hint = QLabel(tr("(0 = не разбивать)"))
            hint.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
            layout.addRow("", hint)

        if with_split:
            self._do_compress = QCheckBox(tr("Уменьшить DPI при разбивке"))
            self._do_compress.setChecked(True)
            layout.addRow(tr("Сжатие:"), self._do_compress)

        bb = QDialogButtonBox(_DB_OK | _DB_CANCEL)
        bb.button(_DB_OK).setText(tr("▶ Начать"))
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addRow(bb)

    def get_config(self) -> dict:
        cfg = {
            "dpi":       int(self._dpi.currentText()),
            "quality":   self._quality.value(),
            "grayscale": self._grayscale.isChecked(),
        }
        if self._with_limit:
            cfg["max_mb"] = float(self._max_mb.value())
        if self._with_split:
            cfg["do_compress"] = self._do_compress.isChecked()
        return cfg


class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(
            f"QFrame {{ background:{C_CARD}; border:1px solid {C_BG3}; border-radius:8px; }}"
        )

# ─────────────────────────────────────────────────────────────────────────────
# TemplateBuilderDialog — конструктор шаблона
# ─────────────────────────────────────────────────────────────────────────────

class _FieldPickCombo(NoWheelComboBox):
    """Выбор поля; список вариантов обновляется под текущие извлекаемые поля.

    provider() должен возвращать актуальный список имён извлекаемых полей.
    Если auto_label задан — первым пунктом идёт он со значением "" (авто).
    """
    def __init__(self, provider, parent=None, auto_label="__auto__"):
        super().__init__(parent)
        if auto_label == "__auto__":
            auto_label = tr("Авто (первое извлекаемое поле)")
        self._provider = provider
        self._auto = auto_label
        if auto_label is not None:
            self.addItem(auto_label, "")

    def refresh(self, select_name=None):
        keep = self.currentData() if select_name is None else select_name
        self.blockSignals(True)
        self.clear()
        if self._auto is not None:
            self.addItem(self._auto, "")
        for nm in self._provider():
            if nm:
                self.addItem(nm, nm)
        i = self.findData(keep) if keep else 0
        self.setCurrentIndex(i if i >= 0 else 0)
        self.blockSignals(False)

    def showPopup(self):
        self.refresh()
        super().showPopup()


_RES_STYLE = {
    "ok":      "color:#15803d; font-weight:bold;",
    "err":     "color:#dc2626; font-weight:bold;",
    "neutral": "color:#475569;",
}

def _regex_cheatsheet() -> str:
    return tr("Шпаргалка regex:  \\d — цифра;  \\d+ — одна или больше цифр;  \\d{6} — ровно 6 цифр;  [.,] — точка или запятая;  \\s — пробел;  .* — любой текст;  (…) — в «Очистке» это часть, которую надо взять;  ? — необязательно.\nПример очистки: 0*(\\d+) из «ОРДЕР № 00000012345» → «12345». Несколько групп склеиваются: [Пп][Кк]\\s*(\\d+)\\s*\\+\\s*(\\d+) из «ПК 400+30» → «40030».")


class FieldRulesDialog(QDialog):
    """Ручная настройка правил одного поля: обработка → очистка → проверка."""

    def __init__(self, state: dict, example: str, field_name: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Правила поля «{field_name}»", field_name=field_name) if field_name else tr("Правила поля"))
        self.setMinimumWidth(720)
        self.setStyleSheet(current_qss())
        self._custom_regex = state.get("regex", "") if state.get("type") == "custom" else ""

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 14)
        lay.setSpacing(10)
        intro = QLabel(
            tr("Текст зоны проходит три шага: <b>1. обработка</b> (готовые исправления распознавания) → <b>2. очистка</b> (что вырезать) → <b>3. проверка</b> (правильный ли формат). Обычно хватает выбора в колонке «Что это за поле».")
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color:{C_TEXT2};")
        lay.addWidget(intro)

        form = QFormLayout()
        form.setSpacing(8)
        form.setLabelAlignment(_ALIGN_RIGHT)
        self._post = NoWheelComboBox()
        for key, label, _fn in POST_PROCESSORS:
            self._post.addItem(tr(label), key)
        i = self._post.findData(state.get("post", ""))
        self._post.setCurrentIndex(max(0, i))
        form.addRow(tr("1. Обработка (OCR):"), self._post)

        self._clean = QLineEdit(state.get("clean", ""))
        self._clean.setPlaceholderText(tr("пусто — без очистки;  напр. 0*(\\d+)"))
        form.addRow(tr("2. Очистка (regex):"), self._clean)

        self._type = NoWheelComboBox()
        for key, label, _pat, _cl, _ex in FIELD_TYPES:
            self._type.addItem(tr(label), key)
        i = self._type.findData(state.get("type", ""))
        self._type.setCurrentIndex(max(0, i))
        form.addRow(tr("3. Проверка формата:"), self._type)

        self._regex = QLineEdit(self._custom_regex)
        self._regex.setPlaceholderText(tr("своё правило проверки, напр. \\d{6}"))
        form.addRow(tr("    свой regex проверки:"), self._regex)
        lay.addLayout(form)

        ex_row = QHBoxLayout()
        ex_row.addWidget(QLabel(tr("Пример текста зоны:")))
        self._ex = QLineEdit(example or "")
        self._ex.setPlaceholderText(tr("вставьте распознанный текст зоны"))
        ex_row.addWidget(self._ex, 1)
        lay.addLayout(ex_row)
        self._res = QLabel("")
        self._res.setWordWrap(True)
        lay.addWidget(self._res)

        cheat = QLabel(_regex_cheatsheet())
        cheat.setWordWrap(True)
        cheat.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt; background:{C_BG};"
                            f" border:1px solid {C_BG3}; border-radius:6px; padding:8px;")
        lay.addWidget(cheat)

        bb = QDialogButtonBox(_DB_OK | _DB_CANCEL)
        bb.button(_DB_OK).setText(tr("Готово"))
        bb.button(_DB_CANCEL).setText(tr("Отмена"))
        bb.accepted.connect(self._on_ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

        self._post.currentIndexChanged.connect(self._update)
        self._type.currentIndexChanged.connect(self._on_type)
        for w in (self._clean, self._regex, self._ex):
            w.textChanged.connect(self._update)
        self._on_type()

    def _on_type(self, *_):
        key = self._type.currentData()
        self._regex.blockSignals(True)
        if key == "custom":
            self._regex.setText(self._custom_regex)
            self._regex.setEnabled(True)
        else:
            if self._regex.isEnabled():
                self._custom_regex = self._regex.text()
            self._regex.setText(FIELD_TYPE_MAP.get(key, ("", "", "", ""))[1] or "")
            self._regex.setEnabled(False)
        self._regex.blockSignals(False)
        self._update()

    def state(self) -> dict:
        t = self._type.currentData() or ""
        return {"post": self._post.currentData() or "", "clean": self._clean.text().strip(),
                "type": t, "regex": self._regex.text().strip() if t == "custom" else ""}

    def example(self) -> str:
        return self._ex.text()

    def _update(self, *_):
        text, kind = rule_result(self._ex.text(), self.state())
        self._res.setText(text or tr("Вставьте пример — здесь появится результат."))
        self._res.setStyleSheet(_RES_STYLE[kind] + " font-size:10pt; padding:2px;")

    def _on_ok(self):
        st = self.state()
        for label, pat in ((tr("очистки"), st["clean"]), (tr("проверки"), st["regex"])):
            if pat:
                try:
                    re.compile(pat)
                except re.error as e:
                    _warn(self, tr("Некорректный regex {label}:\n{e}", label=label, e=e), tr("Проверка"))
                    return
        self.accept()


class TemplateBuilderDialog(QDialog):
    """Конструктор шаблона: поля, правила по примеру, формулы, порядок, отчёт."""

    def __init__(self, template: Optional[Template] = None, existing_names=None, parent=None,
                 copy_of: str = "", sample_reader=None):
        super().__init__(parent)
        # sample_reader(имя поля, страница) -> текст зоны из открытого PDF (или "")
        self._sample_reader = sample_reader
        self._editing = template is not None
        self._copy_of = copy_of
        self._orig_name = template.name if template else ""
        self._existing_names = set(existing_names or [])
        self._existing_zones = dict(template.zones) if template else {}
        self._existing_fallback = dict(template.fallback_zones) if template else {}
        self._existing_dictpaths = (
            {f["name"]: f.get("dict_path", "") for f in template.extract_fields}
            if template else {}
        )
        self._existing_auto_start = dict(template.auto_start) if template else {}
        self._manual_rows: list = []    # [(QLineEdit, QCheckBox)]
        self._extract_rows: list = []   # [dict: виджеты + правила]
        self._computed_rows: list = []  # [(QLineEdit, QLineEdit, QSpinBox)]
        self._uid_seq = 0

        self.setWindowTitle(self._title_text())
        self.resize(1280, 860)
        self.setStyleSheet(current_qss())
        self._setup_ui()
        if template:
            self._prefill(template)
        else:
            self._extract_count.setValue(1)
        if copy_of:
            self._name.setToolTip(tr("Новое название. Исходный шаблон не меняется."))
            self._name.setFocus()
            self._name.selectAll()

    def _title_text(self) -> str:
        if self._copy_of:
            return tr("Копия шаблона «{copy_of}»", copy_of=self._copy_of)
        return tr("Изменение шаблона") if self._editing else tr("Создание шаблона")

    # ── UI ────────────────────────────────────────────────────────────────
    @staticmethod
    def _hint(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setWordWrap(True)
        lbl.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
        return lbl

    @staticmethod
    def _head(text: str, width: int = 0, tip: str = "") -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt; font-weight:bold;")
        if width:
            lbl.setFixedWidth(width)
        if tip:
            lbl.setToolTip(tip)
        return lbl

    def _count_row(self, parent_lay, maximum: int, slot) -> QSpinBox:
        row = QHBoxLayout()
        row.addWidget(QLabel(tr("Количество полей:")))
        sp = NoWheelSpinBox()
        sp.setRange(0, maximum)
        sp.setFixedWidth(80)
        sp.valueChanged.connect(lambda *_: slot())
        row.addWidget(sp)
        row.addStretch()
        parent_lay.addLayout(row)
        return sp

    @staticmethod
    def _host(parent_lay):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 2, 0, 0)
        lay.setSpacing(6)
        parent_lay.addWidget(w)
        return lay

    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        hdr = QFrame()
        hdr.setFixedHeight(46)
        hdr.setStyleSheet(f"background:{C_ACCENT};")
        hlay = QHBoxLayout(hdr)
        hlay.setContentsMargins(18, 0, 18, 0)
        title = QLabel("⚙  " + self._title_text())
        title.setStyleSheet("color:white; font-weight:bold; font-size:12pt;")
        hlay.addWidget(title)
        root.addWidget(hdr)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(_NO_FRAME)
        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(14)

        # ① Основное
        gb0 = QGroupBox(tr("1. Основное"))
        f0 = QFormLayout(gb0)
        f0.setSpacing(10)
        f0.setLabelAlignment(_ALIGN_RIGHT)
        self._name = QLineEdit()
        self._name.setPlaceholderText(tr("Например: Товарные накладные"))
        f0.addRow(tr("Название шаблона:"), self._name)
        self._pages = NoWheelSpinBox()
        self._pages.setRange(1, MAX_DOC_PAGES)
        self._pages.setFixedWidth(80)
        self._pages.valueChanged.connect(lambda *_: self._rebuild_extract_rows())
        pr = QHBoxLayout()
        pr.addWidget(self._pages)
        pr.addWidget(self._hint(tr("сколько страниц скана — один документ (1–{MAX_DOC_PAGES}). Двусторонний скан накладной = 2.", MAX_DOC_PAGES=MAX_DOC_PAGES)), 1)
        f0.addRow(tr("Страниц на документ:"), pr)
        lay.addWidget(gb0)

        # ② Ручные поля
        gb_m = QGroupBox(tr("2. Поля, заполняемые вручную"))
        m_lay = QVBoxLayout(gb_m)
        m_lay.addWidget(self._hint(
            tr("Одинаковы для всей пачки: год, дата, перевозчик, материал… Заполняются на вкладке «1. Общие данные». Галочка «№ по порядку» — номер строки ставится сам (1, 2, 3…) и продолжается после сохранения в Excel.")))
        self._manual_count = self._count_row(m_lay, 30, self._rebuild_manual_rows)
        self._manual_host_lay = self._host(m_lay)
        lay.addWidget(gb_m)

        # ③ Поля из PDF
        gb_e = QGroupBox(tr("3. Поля из PDF (у каждого — своя зона на странице)"))
        e_lay = QVBoxLayout(gb_e)
        e_lay.addWidget(self._hint(
            tr("Как настроить поле: нажмите 🔮 — программа возьмёт текст из рамки поля в открытом PDF "
               "и сама подберёт вид поля. Справа виден результат: ✓ — правило работает, ✗ — нет. "
               "Если PDF не открыт — вставьте пример текста вручную. Ручные правила (regex) — кнопка ⚙.")))
        self._extract_count = self._count_row(e_lay, 40, self._rebuild_extract_rows)
        btn_all = QPushButton(tr("🔮 Подобрать вид для всех полей"))
        btn_all.setToolTip(tr("Для каждого поля: пример из открытого PDF → подбор вида"))
        btn_all.clicked.connect(self._suggest_all)
        e_lay.addWidget(btn_all, alignment=_ALIGN_LEFT)
        h = QHBoxLayout()
        h.setSpacing(6)
        h.addWidget(self._head(tr("Название столбца")), 2)
        h.addWidget(self._head(tr("Что это за поле"), 250))
        h.addWidget(self._head(tr("Пример текста зоны")), 3)
        h.addWidget(self._head("", 34))
        h.addWidget(self._head(tr("Результат"), 230))
        h.addWidget(self._head(tr("Стр."), 92, tr("На какой странице документа поле")))
        h.addWidget(self._head(tr("Слов."), 46, tr("Сопоставлять со справочником (словарём)")))
        h.addWidget(self._head(tr("Резерв"), 56, tr("Запасная зона: читается, если основная пуста")))
        h.addWidget(self._head("", 34))
        e_lay.addLayout(h)
        self._extract_host_lay = self._host(e_lay)
        lay.addWidget(gb_e)

        # ④ Вычисляемые поля
        gb_c = QGroupBox(tr("4. Вычисляемые поля (формула из других полей)"))
        c_lay = QVBoxLayout(gb_c)
        c_lay.addWidget(self._hint(
            tr("Поле в формуле — в квадратных скобках: [Тонн] / 1,6. Можно + − * / и скобки. Округление математическое (28,375 → 28,38). Пустое или нечисловое исходное поле даёт пустой результат.")))
        self._computed_count = self._count_row(c_lay, 20, self._rebuild_computed_rows)
        hc = QHBoxLayout()
        hc.setSpacing(6)
        hc.addWidget(self._head(tr("Название столбца")), 1)
        hc.addWidget(self._head(tr("Формула")), 2)
        hc.addWidget(self._head(tr("Знаков после «,»"), 120))
        c_lay.addLayout(hc)
        self._computed_host_lay = self._host(c_lay)
        lay.addWidget(gb_c)

        # ⑤ Ключевое поле
        gb_k = QGroupBox(tr("5. Ключевое поле"))
        k_lay = QHBoxLayout(gb_k)
        k_lay.addWidget(QLabel(tr("Поле:")))
        self._name_combo = _FieldPickCombo(self._extract_names_now)
        self._name_combo.setMinimumWidth(260)
        k_lay.addWidget(self._name_combo)
        k_lay.addWidget(self._hint(
            tr("Обычно № документа. По нему ищутся повторы (в пачке и в Excel) и называются файлы при «Разбить на файлы».")), 1)
        lay.addWidget(gb_k)

        # ⑥ Порядок столбцов
        gb_o = QGroupBox(tr("6. Порядок столбцов в таблице и Excel (перетаскивайте мышью)"))
        o_lay = QHBoxLayout(gb_o)
        o_lay.setSpacing(8)
        self._order_list = QListWidget()
        self._order_list.setDragDropMode(_DND_INTERNAL)
        self._order_list.setDefaultDropAction(_MOVE_ACTION)
        self._order_list.setSelectionMode(_SINGLE_SEL)
        self._order_list.setStyleSheet("QListWidget::item { padding: 5px 8px; }")
        o_lay.addWidget(self._order_list, 1)
        ob = QVBoxLayout()
        ob.setSpacing(6)
        for txt, d, tip in (("▲", -1, tr("Выше")), ("▼", 1, tr("Ниже"))):
            b = QPushButton(txt)
            b.setFixedWidth(44)
            b.setToolTip(tip)
            b.clicked.connect(lambda _=False, dd=d: self._move_order(dd))
            ob.addWidget(b)
        br = QPushButton("↺")
        br.setFixedWidth(44)
        br.setToolTip(tr("Сбросить: ручные → из PDF → вычисляемые"))
        br.clicked.connect(self._reset_order)
        ob.addWidget(br)
        ob.addStretch()
        o_lay.addLayout(ob)
        lay.addWidget(gb_o)

        # ⑦ Отчёт
        gb_r = QGroupBox(tr("7. Отчёт — кнопка «📊 Отчёт» на вкладке «Таблица»"))
        r_lay = QVBoxLayout(gb_r)
        r_lay.addWidget(self._hint(
            tr("Одна строка = одна колонка отчёта.   [Имя поля] — значение поля.   Подпись = значение — постоянный текст (его можно поправить перед копированием).   - — пустая колонка.   Пусто — отчёт не настроен.")))
        self._report_edit = QPlainTextEdit()
        self._report_edit.setPlaceholderText(
            tr("-\n[Дата]\nПоставщик = ООО «Ромашка»\n[№ документа]\n[Вес]"))
        self._report_edit.setFixedHeight(180)
        self._report_edit.setStyleSheet(
            f"QPlainTextEdit{{background:{C_CARD}; color:{C_TEXT}; border:1px solid {C_BG3};"
            " border-radius:6px; font-family:'Consolas','Courier New',monospace; font-size:10pt;}")
        r_lay.addWidget(self._report_edit)
        lay.addWidget(gb_r)
        lay.addStretch()

        scroll.setWidget(content)
        root.addWidget(scroll, 1)

        bb = QDialogButtonBox(_DB_OK | _DB_CANCEL)
        bb.button(_DB_OK).setText(tr("💾 Сохранить шаблон"))
        bb.button(_DB_CANCEL).setText(tr("Отмена"))
        bb.accepted.connect(self._on_accept)
        bb.rejected.connect(self.reject)
        bar = QHBoxLayout()
        bar.setContentsMargins(18, 8, 18, 12)
        bar.addStretch()
        bar.addWidget(bb)
        root.addLayout(bar)

    def _new_uid(self) -> int:
        self._uid_seq += 1
        return self._uid_seq

    def _extract_names_now(self) -> list:
        return [r["ne"].text().strip() for r in self._extract_rows if r["ne"].text().strip()]

    # ── Ручные поля ──────────────────────────────────────────────────────────
    def _rebuild_manual_rows(self, data: Optional[list] = None):
        if data is None:
            data = [(e.text(), cb.isChecked(), e._uid) for (e, cb) in self._manual_rows]
        _clear_layout(self._manual_host_lay)
        self._manual_rows = []
        for i in range(self._manual_count.value()):
            text, auto, uid = data[i] if i < len(data) else ("", False, self._new_uid())
            row = QHBoxLayout()
            ed = QLineEdit(text)
            ed.setPlaceholderText(tr("Название поля {v}", v=i + 1))
            ed._uid = uid
            ed.textChanged.connect(self._sync_order_list)
            row.addWidget(ed, 1)
            cb = QCheckBox(tr("№ по порядку"))
            cb.setChecked(auto)
            cb.setToolTip(tr("Номер строки ставится автоматически: начало задаётся на вкладке «1. Общие данные», после сохранения в Excel продолжается."))
            row.addWidget(cb)
            self._manual_host_lay.addLayout(row)
            self._manual_rows.append((ed, cb))
        self._sync_order_list()

    # ── Поля из PDF ──────────────────────────────────────────────────────────
    def _read_extract(self) -> list:
        return [{"name": r["ne"].text(), "page": r["pc"].currentIndex(),
                 "dict": r["dc"].isChecked(), "fallback": r["fbc"].isChecked(),
                 "example": r["ex"].text(), "uid": r["uid"], "orig": r["orig"], **r["state"]}
                for r in self._extract_rows]

    def _rebuild_extract_rows(self, data: Optional[list] = None):
        if not hasattr(self, "_extract_host_lay"):
            return
        if data is None:
            data = self._read_extract()
        _clear_layout(self._extract_host_lay)
        self._extract_rows = []
        ppd = self._pages.value()
        for i in range(self._extract_count.value()):
            d = data[i] if i < len(data) else {"uid": self._new_uid(), **kind_state("text")}
            self._extract_rows.append(self._make_extract_row(i, d, ppd))
        self._sync_order_list()

    def _make_extract_row(self, i: int, d: dict, ppd: int) -> dict:
        row = QHBoxLayout()
        row.setSpacing(6)
        r = {"uid": d.get("uid") or self._new_uid(), "orig": d.get("orig", ""),
             "state": {k: d.get(k, "") for k in ("post", "clean", "type", "regex")}}

        ne = QLineEdit(d.get("name", ""))
        ne.setPlaceholderText(tr("Поле {v}", v=i + 1))
        ne._uid = r["uid"]
        ne.textChanged.connect(self._sync_order_list)
        row.addWidget(ne, 2)

        kc = NoWheelComboBox()
        for key in KIND_MAP:
            kc.addItem(tr(KIND_MAP[key]["label"]), key)
            kc.setItemData(kc.count() - 1, tr(KIND_MAP[key]["hint"]), _qenum(Qt, "ItemDataRole", "ToolTipRole"))
        kc.addItem(tr("⚙ Своё правило (вручную)"), KIND_CUSTOM)
        kc.setFixedWidth(250)
        row.addWidget(kc)

        ex = QLineEdit(d.get("example", ""))
        ex.setPlaceholderText(tr("напр. ОРДЕР № 00000012345"))
        ex.setToolTip(tr("Пример распознанного текста зоны. Нужен только для проверки правила."))
        row.addWidget(ex, 3)

        bt = QPushButton("🔮")
        bt.setFixedSize(34, 32)
        bt.setStyleSheet("QPushButton{padding:0; font-size:12pt;}")
        bt.setToolTip(tr("Подобрать вид поля по примеру. Если пример пуст — текст берётся из открытого PDF."))
        row.addWidget(bt)

        res = QLabel("")
        res.setFixedWidth(230)
        row.addWidget(res)

        pc = NoWheelComboBox()
        pc.addItems([tr("Стр. {v}", v=k + 1) for k in range(ppd)])
        pc.setFixedWidth(92)
        pc.setCurrentIndex(max(0, min(int(d.get("page", 0) or 0), ppd - 1)))
        pc.setEnabled(ppd > 1)
        row.addWidget(pc)

        dc = QCheckBox()
        dc.setFixedWidth(46)
        dc.setStyleSheet("QCheckBox{margin-left:14px;}")
        dc.setChecked(bool(d.get("dict")))
        row.addWidget(dc)

        fbc = QCheckBox()
        fbc.setFixedWidth(56)
        fbc.setStyleSheet("QCheckBox{margin-left:18px;}")
        fbc.setChecked(bool(d.get("fallback")))
        row.addWidget(fbc)

        adv = QPushButton("⚙")
        adv.setFixedSize(34, 32)
        adv.setStyleSheet("QPushButton{padding:0; font-size:12pt;}")
        adv.setToolTip(tr("Ручные правила: обработка, очистка, проверка (regex)"))
        row.addWidget(adv)

        r.update(ne=ne, kc=kc, ex=ex, res=res, pc=pc, dc=dc, fbc=fbc)
        key = match_kind(r["state"])
        kc.setCurrentIndex(max(0, kc.findData(key)))
        r["kind"] = key

        kc.currentIndexChanged.connect(lambda *_, rr=r: self._on_kind_changed(rr))
        ex.textChanged.connect(lambda *_, rr=r: self._update_result(rr))
        bt.clicked.connect(lambda _=False, rr=r: self._suggest(rr))
        adv.clicked.connect(lambda _=False, rr=r: self._open_rules(rr))

        self._extract_host_lay.addLayout(row)
        self._update_result(r)
        return r

    def _set_kind(self, r: dict, key: str):
        r["kc"].blockSignals(True)
        r["kc"].setCurrentIndex(max(0, r["kc"].findData(key)))
        r["kc"].blockSignals(False)
        r["kind"] = key

    def _on_kind_changed(self, r: dict):
        key = r["kc"].currentData()
        if key == KIND_CUSTOM:
            if not self._open_rules(r):
                self._set_kind(r, r["kind"])
            return
        r["state"] = kind_state(key)
        r["kind"] = key
        if not r["ex"].text().strip():
            r["ex"].setPlaceholderText(tr("напр. ") + tr(KIND_MAP[key]["example"]))
        self._update_result(r)

    def _open_rules(self, r: dict) -> bool:
        dlg = FieldRulesDialog(r["state"], r["ex"].text(), r["ne"].text().strip(), self)
        if dlg.exec() != _DLG_OK:
            return False
        r["state"] = dlg.state()
        r["ex"].setText(dlg.example())
        self._set_kind(r, match_kind(r["state"]))
        self._update_result(r)
        return True

    def _read_sample(self, r: dict) -> str:
        """Текст зоны поля из открытого PDF (первый непустой документ)."""
        if self._sample_reader is None:
            return ""
        name = r["orig"] or r["ne"].text().strip()
        try:
            return (self._sample_reader(name, r["pc"].currentIndex()) or "").strip()
        except Exception as e:
            _log_error("чтение примера из PDF", e)
            return ""

    def _suggest(self, r: dict, quiet: bool = False) -> bool:
        ex = r["ex"].text()
        if not ex.strip():
            ex = self._read_sample(r)
            if ex:
                r["ex"].setText(ex)
        if not ex.strip():
            if not quiet:
                _info(self, tr("Нет примера текста для этого поля.\n\n"
                               "Вариант 1: откройте PDF (вкладка «2. Распознавание»), поставьте рамку поля "
                               "на нужное место (вкладка «3. Редактор зон») и снова нажмите 🔮 — "
                               "пример возьмётся из PDF сам.\n\n"
                               "Вариант 2: вставьте в «Пример» текст вручную."), tr("Подбор вида поля"))
            return False
        found = suggest_kinds(ex, r["ne"].text())
        if not found:
            r["res"].setText(tr("не подобрать — выберите вручную"))
            r["res"].setStyleSheet(_RES_STYLE["err"])
            return False
        key = found[0][0]
        r["state"] = kind_state(key)
        self._set_kind(r, key)
        self._update_result(r)
        others = [f"{tr(KIND_MAP[k]['label'])} → «{v}»" for k, v in found[1:4]]
        if others:
            r["res"].setToolTip(tr("Подходит также:\n") + "\n".join(others))
        return True

    def _suggest_all(self):
        """🔮 для всех полей: пример берётся из PDF, вид подбирается сам."""
        done = sum(1 for r in self._extract_rows if self._suggest(r, quiet=True))
        miss = len(self._extract_rows) - done
        if miss:
            _info(self, tr("Подобрано полей: {done}. Без примера: {miss}.\n\n"
                           "Для полей без примера поставьте рамку на текст (вкладка «3. Редактор зон») "
                           "или вставьте пример вручную.", done=done, miss=miss), tr("Подбор вида поля"))

    def _update_result(self, r: dict):
        text, kind = rule_result(r["ex"].text(), r["state"])
        if not text:
            text = tr(KIND_MAP[r["kind"]]["hint"]) if r["kind"] in KIND_MAP else tr("своё правило")
            kind = "neutral"
        r["res"].setText(text)
        r["res"].setToolTip(text)
        r["res"].setStyleSheet(_RES_STYLE[kind] + " font-size:9pt;")

    # ── Вычисляемые поля ─────────────────────────────────────────────────────
    def _rebuild_computed_rows(self, data: Optional[list] = None):
        if data is None:
            data = [(n.text(), f.text(), d.value(), n._uid) for (n, f, d) in self._computed_rows]
        _clear_layout(self._computed_host_lay)
        self._computed_rows = []
        for i in range(self._computed_count.value()):
            name, formula, dec, uid = (data[i] if i < len(data)
                                       else ("", "", 2, self._new_uid()))
            row = QHBoxLayout()
            row.setSpacing(6)
            ne = QLineEdit(name)
            ne.setPlaceholderText(tr("Вычисляемое поле {v}", v=i + 1))
            ne._uid = uid
            ne.textChanged.connect(self._sync_order_list)
            row.addWidget(ne, 1)
            fe = QLineEdit(formula)
            fe.setPlaceholderText(tr("напр. [Тонн] / 1,6"))
            row.addWidget(fe, 2)
            ds = NoWheelSpinBox()
            ds.setRange(0, 6)
            ds.setValue(int(dec))
            ds.setFixedWidth(120)
            row.addWidget(ds)
            self._computed_host_lay.addLayout(row)
            self._computed_rows.append((ne, fe, ds))
        self._sync_order_list()

    # ── Порядок столбцов ─────────────────────────────────────────────────────
    def _order_entries(self) -> dict:
        out = {}
        for i, (e, _cb) in enumerate(self._manual_rows):
            out[e._uid] = ("m", e.text().strip() or tr("(ручное поле {v})", v=i + 1))
        for i, r in enumerate(self._extract_rows):
            out[r["uid"]] = ("e", r["ne"].text().strip() or tr("(поле из PDF {v})", v=i + 1))
        for i, (n, _f, _d) in enumerate(self._computed_rows):
            out[n._uid] = ("c", n.text().strip() or tr("(вычисляемое поле {v})", v=i + 1))
        return out

    @staticmethod
    def _order_label(kind: str, text: str) -> str:
        return {"m": tr("✍  {text}      — вручную", text=text),
                "c": tr("🧮  {text}      — формула", text=text)}.get(kind, tr("📄  {text}      — из PDF", text=text))

    def _sync_order_list(self, *_):
        lst = getattr(self, "_order_list", None)
        if lst is None:
            return
        entries = self._order_entries()
        cur = [lst.item(i).data(_USER_ROLE) for i in range(lst.count())]
        order = [u for u in cur if u in entries] + [u for u in entries if u not in cur]
        if order == cur:
            for i, u in enumerate(order):
                lst.item(i).setText(self._order_label(*entries[u]))
        else:
            lst.clear()
            for u in order:
                it = QListWidgetItem(self._order_label(*entries[u]))
                it.setData(_USER_ROLE, u)
                lst.addItem(it)
        lst.setFixedHeight(max(3, min(lst.count(), 12)) * 30 + 8)

    def _name_to_uid(self) -> dict:
        m = {e.text().strip(): e._uid for (e, _cb) in self._manual_rows}
        m.update({r["ne"].text().strip(): r["uid"] for r in self._extract_rows})
        m.update({n.text().strip(): n._uid for (n, _f, _d) in self._computed_rows})
        return m

    def _set_order_by_names(self, names: list):
        n2u = self._name_to_uid()
        self._order_list.clear()
        for nm in names:
            if nm in n2u:
                it = QListWidgetItem(nm)
                it.setData(_USER_ROLE, n2u[nm])
                self._order_list.addItem(it)
        self._sync_order_list()

    def _move_order(self, d: int):
        lst = self._order_list
        row = lst.currentRow()
        nr = row + d
        if row < 0 or not (0 <= nr < lst.count()):
            return
        lst.insertItem(nr, lst.takeItem(row))
        lst.setCurrentRow(nr)

    def _reset_order(self):
        self._order_list.clear()
        self._sync_order_list()

    def _order_names(self) -> list:
        u2n = {u: n for n, u in self._name_to_uid().items()}
        return [u2n[self._order_list.item(i).data(_USER_ROLE)]
                for i in range(self._order_list.count())
                if u2n.get(self._order_list.item(i).data(_USER_ROLE))]

    # ── Заполнение / сохранение ──────────────────────────────────────────────
    def _prefill(self, tpl: Template):
        self._name.setText(tpl.name)
        self._pages.blockSignals(True)
        self._pages.setValue(tpl.pages)
        self._pages.blockSignals(False)

        self._manual_count.blockSignals(True)
        self._manual_count.setValue(len(tpl.manual_fields))
        self._manual_count.blockSignals(False)
        self._rebuild_manual_rows([(n, n in tpl.auto_number, self._new_uid())
                                   for n in tpl.manual_fields])

        self._extract_count.blockSignals(True)
        self._extract_count.setValue(len(tpl.extract_fields))
        self._extract_count.blockSignals(False)
        self._rebuild_extract_rows([dict(f, uid=self._new_uid(), orig=f["name"])
                                    for f in tpl.extract_fields])

        self._computed_count.blockSignals(True)
        self._computed_count.setValue(len(tpl.computed_fields))
        self._computed_count.blockSignals(False)
        self._rebuild_computed_rows([(c["name"], c["formula"], c["decimals"], self._new_uid())
                                     for c in tpl.computed_fields])

        self._report_edit.setPlainText("\n".join(tpl.report_lines))
        self._name_combo.refresh(tpl.name_field)
        self._set_order_by_names(tpl.column_order)

    def _apply_renames(self):
        """Переименованные поля из PDF: [Старое] → [Новое] в формулах и отчёте."""
        ren = {r["orig"]: r["ne"].text().strip() for r in self._extract_rows
               if r["orig"] and r["ne"].text().strip() and r["orig"] != r["ne"].text().strip()}
        if not ren:
            return
        def sub(text):
            return _FORMULA_REF.sub(lambda m: f"[{ren.get(m.group(1).strip(), m.group(1))}]", text)
        for (_n, fe, _d) in self._computed_rows:
            fe.setText(sub(fe.text()))
        self._report_edit.setPlainText(sub(self._report_edit.toPlainText()))
        if self._name_combo.currentData() in ren:
            self._name_combo.refresh(ren[self._name_combo.currentData()])
        for r in self._extract_rows:
            if r["orig"] in ren:
                r["orig_saved"] = r["orig"]

    def _report_lines_now(self) -> list:
        lines = [ln.rstrip() for ln in self._report_edit.toPlainText().splitlines()]
        while lines and not lines[-1].strip():
            lines.pop()
        return lines

    def _on_accept(self):
        name = self._name.text().strip()
        if not name:
            _warn(self, tr("Укажите название шаблона."), tr("Проверка"))
            return
        if name != self._orig_name and name in self._existing_names:
            _warn(self, tr("Шаблон «{name}» уже существует.", name=name), tr("Проверка"))
            return
        self._apply_renames()
        manual = [e.text().strip() for (e, _cb) in self._manual_rows]
        extract = [r["ne"].text().strip() for r in self._extract_rows]
        computed = [n.text().strip() for (n, _f, _d) in self._computed_rows]
        all_names = manual + extract + computed
        if any(not nm for nm in all_names):
            _warn(self, tr("Все названия полей должны быть заполнены."), tr("Проверка"))
            return
        dup = [n for n in dict.fromkeys(all_names) if all_names.count(n) > 1]
        if dup:
            _warn(self, tr("Названия полей повторяются: ") + ", ".join(dup), tr("Проверка"))
            return
        for r in self._extract_rows:
            st = r["state"]
            for label, pat in ((tr("очистки"), st.get("clean", "")),
                               (tr("проверки"), st.get("regex", "") if st.get("type") == "custom" else "")):
                if pat:
                    try:
                        re.compile(pat)
                    except re.error as e:
                        _warn(self, tr("Некорректный regex {label} у поля «{v}»:\n{e}", label=label, v=r['ne'].text().strip(), e=e), tr("Проверка"))
                        return
        known = manual + extract
        for (ne, fe, _d) in self._computed_rows:
            nm, formula = ne.text().strip(), fe.text().strip()
            if not formula:
                _warn(self, tr("Укажите формулу для поля «{nm}».", nm=nm), tr("Проверка"))
                return
            try:
                compile_formula(formula)
            except Exception as e:
                _warn(self, tr("Ошибка в формуле поля «{nm}»:\n{formula}\n\n{e}", nm=nm, formula=formula, e=e), tr("Проверка"))
                return
            for ref in formula_refs(formula):
                if ref not in known:
                    _warn(self, tr("В формуле поля «{nm}» поле [{ref}] не найдено.\nМожно ссылаться на ручные поля, поля из PDF и вычисляемые поля выше по списку.", nm=nm, ref=ref), tr("Проверка"))
                    return
            known.append(nm)
        for ln in self._report_lines_now():
            kind, a, _b = parse_report_line(ln)
            if kind == "field" and a not in all_names:
                _warn(self, tr("В отчёте указано поле [{a}], которого нет в шаблоне.", a=a), tr("Проверка"))
                return
        self._result = self._build_template(name, manual)
        self.accept()

    def _build_template(self, name: str, manual: list) -> Template:
        pages = self._pages.value()
        extract = []
        src_of = {}
        for r in self._extract_rows:
            nm = r["ne"].text().strip()
            src_of[nm] = r["orig"] or nm
            st = r["state"]
            extract.append({
                "name":      nm,
                "page":      max(0, min(r["pc"].currentIndex(), pages - 1)) if pages > 1 else 0,
                "dict":      r["dc"].isChecked(),
                "dict_path": self._existing_dictpaths.get(src_of[nm], ""),
                "kind":      match_kind(st),
                "type":      st.get("type", ""),
                "regex":     st.get("regex", "") if st.get("type") == "custom" else "",
                "clean":     st.get("clean", ""),
                "post":      st.get("post", ""),
                "fallback":  r["fbc"].isChecked(),
                "example":   r["ex"].text(),
            })
        zones = {f["name"]: self._existing_zones[src_of[f["name"]]]
                 for f in extract if src_of[f["name"]] in self._existing_zones}
        fallback_zones = {f["name"]: self._existing_fallback[src_of[f["name"]]]
                          for f in extract
                          if f["fallback"] and src_of[f["name"]] in self._existing_fallback}
        auto = [e.text().strip() for (e, cb) in self._manual_rows if cb.isChecked()]
        return Template(
            name=name, pages=pages, manual_fields=manual, extract_fields=extract,
            zones=zones, fallback_zones=fallback_zones,
            name_field=self._name_combo.currentData() or "",
            column_order=self._order_names(),
            computed_fields=[{"name": n.text().strip(), "formula": f.text().strip(),
                              "decimals": d.value()} for (n, f, d) in self._computed_rows],
            report_lines=self._report_lines_now(),
            auto_number=auto,
            auto_start={n: self._existing_auto_start.get(n, 1) for n in auto},
        )

    def get_template(self) -> Optional[Template]:
        return getattr(self, "_result", None)


# ─────────────────────────────────────────────────────────────────────────────
# Значок программы (рисуется кодом, без файлов)
# ─────────────────────────────────────────────────────────────────────────────

def app_icon_pixmap(size: int = 64) -> QPixmap:
    """Значок: лист документа с пунктирной зоной и строками текста."""
    pm = QPixmap(size, size)
    pm.fill(QColor(0, 0, 0, 0))
    p = QPainter(pm)
    p.setRenderHint(_ANTIALIASING)
    k = size / 64.0
    p.setPen(QPen(QColor("#1e3a8a"), 2.5 * k))
    p.setBrush(QBrush(QColor("#ffffff")))
    p.drawRoundedRect(QRectF(12 * k, 6 * k, 40 * k, 52 * k), 5 * k, 5 * k)
    p.setPen(QPen(QColor("#cbd5e1"), 3 * k))
    for y in (16, 22, 46, 52):
        p.drawLine(QPointF(19 * k, y * k), QPointF(45 * k, y * k))
    zone = QPen(QColor("#3b82f6"), 2.5 * k)
    zone.setStyle(_DASH_PEN)
    p.setPen(zone)
    p.setBrush(QBrush(QColor(59, 130, 246, 50)))
    p.drawRect(QRectF(17 * k, 28 * k, 30 * k, 12 * k))
    p.setPen(QPen(QColor("#2563eb"), 3 * k))
    p.drawLine(QPointF(21 * k, 34 * k), QPointF(37 * k, 34 * k))
    p.end()
    return pm


# ─────────────────────────────────────────────────────────────────────────────
# Главное окно
# ─────────────────────────────────────────────────────────────────────────────

class _ChromeStyleMixin:
    """Стили и перекраска «хрома» (шапка, тулбары, разделители)."""

    def _retheme(self):
        """Перекрашивает элементы главного окна с собственными инлайн-стилями."""
        self._header.setStyleSheet(self._header_bg_ss())
        self._tmpl_lbl.setStyleSheet(self._header_dim_ss())
        self._ver_lbl.setStyleSheet(self._header_dim_ss())
        self._tmpl_combo.setStyleSheet(self._header_combo_ss())
        for b in self._hdr_btns:
            b.setStyleSheet(self._header_btn_ss())
        self._stripe.setStyleSheet(f"background:{T.accent};")
        if getattr(self, "_zones_desc", None) is not None:
            self._zones_desc.setStyleSheet(self._zones_desc_ss())
        self._retheme_table_toolbar()
        if getattr(self, "_pdf_info", None) is not None:
            self._pdf_info.setStyleSheet(f"color:{T.accent}; font-size:9pt;")
        if getattr(self, "_totals_lbl", None) is not None:
            self._totals_lbl.setStyleSheet(self._totals_ss())
        # Перекрасить уже заполненную таблицу под новую тему.
        if self.template is not None and self.waybills:
            self._refresh()

    @staticmethod
    def _totals_ss() -> str:
        return (f"background:{T.surface}; color:{T.text}; padding:6px 12px;"
                f" border-top:1px solid {T.content_border}; font-size:10pt;")

    @staticmethod
    def _zones_desc_ss() -> str:
        return (
            f"QFrame#zonesBar{{background:{T.chrome}; border-bottom:1px solid {T.chrome_raised};}}"
            f"QFrame#zonesBar QLabel{{color:{T.on_chrome_dim}; font-size:9pt; background:transparent;}}"
        )

    def _retheme_table_toolbar(self):
        tb = getattr(self, "_tbl_toolbar", None)
        if tb is None:
            return
        tb.setStyleSheet(f"background:{T.chrome}; border-bottom:2px solid {T.accent};")
        for s in self._tbl_seps:
            s.setStyleSheet(f"color:{T.chrome_border}; background:{T.chrome_border}; max-width:1px;")
        for b, danger in self._tbl_btns:
            b.setStyleSheet(self._tbl_btn_danger_ss() if danger else self._tbl_btn_ss())

    # ═════════════════════════════════════════════════════════════════════════
    # UI
    # ═════════════════════════════════════════════════════════════════════════

    # ── Стили шапки (строятся из активной темы) ──────────────────────────────
    @staticmethod
    def _header_bg_ss() -> str:
        return (
            f"background: qlineargradient(x1:0,y1:0,x2:1,y2:0,"
            f"stop:0 {T.header_from},stop:1 {T.header_to});"
        )

    @staticmethod
    def _header_dim_ss() -> str:
        return f"color:{T.header_dim}; font-size:9pt; background:transparent;"

    @staticmethod
    def _header_combo_ss() -> str:
        return (
            f"QComboBox {{ background:{T.header_ctl}; color:white; border:1px solid {T.header_ctl_border}; "
            "border-radius:5px; padding:4px 8px; font-size:9pt; font-weight:bold; }"
            "QComboBox::drop-down { border:none; width:18px; }"
            f"QComboBox QAbstractItemView {{ background:{T.chrome}; color:white; "
            f"selection-background-color:{T.header_ctl_hi}; }}"
        )

    @staticmethod
    def _header_btn_ss() -> str:
        return (
            f"QPushButton{{background:{T.header_ctl};color:white;border:1px solid {T.header_ctl_border};"
            "border-radius:5px;font-size:9pt;font-weight:bold;padding:4px 12px;}"
            f"QPushButton:hover{{background:{T.header_ctl_hi};}}"
            f"QPushButton:disabled{{background:{T.header_ctl_off};color:{T.header_ctl_off_fg};border-color:{T.header_ctl_off_fg};}}"
        )


class _TemplateDictMixin:
    """Шаблоны, словари и сопоставление полей."""

    # ═════════════════════════════════════════════════════════════════════════
    # Управление шаблонами
    # ═════════════════════════════════════════════════════════════════════════

    def _refresh_template_combo(self):
        self._tmpl_combo.blockSignals(True)
        self._tmpl_combo.clear()
        for t in self.templates:
            self._tmpl_combo.addItem(t.name)
        if self.template is not None:
            idx = self._tmpl_combo.findText(self.template.name)
            if idx >= 0:
                self._tmpl_combo.setCurrentIndex(idx)
        self._tmpl_combo.blockSignals(False)
        has = self.template is not None
        self._btn_edit_tmpl.setEnabled(has)
        self._btn_copy_tmpl.setEnabled(has)
        self._btn_del_tmpl.setEnabled(has)

    def _select_template_by_name(self, name: str):
        for t in self.templates:
            if t.name == name:
                self._apply_template(t)
                self._refresh_template_combo()
                return

    def _on_template_selected(self, idx: int):
        if idx < 0 or idx >= len(self.templates):
            return
        self._apply_template(self.templates[idx])

    def _new_template(self):
        dlg = TemplateBuilderDialog(
            template=None,
            existing_names=[t.name for t in self.templates],
            parent=self,
        )
        if dlg.exec() != _DLG_OK:
            return
        tpl = dlg.get_template()
        if tpl is None:
            return
        self.templates.append(tpl)
        save_templates(self.templates, tpl.name)
        self._apply_template(tpl)
        self._refresh_template_combo()
        self.status.showMessage(tr("Шаблон «{name}» создан.", name=tpl.name))

    def _sample_text(self, field_name: str, page_idx: int) -> str:
        """Текст зоны поля в открытом PDF: первый непустой из первых 10 документов."""
        if self.parser is None or self.template is None:
            return ""
        coords = self.template.zones.get(field_name)
        if coords is None:
            return ""
        ppd = self.template.pages
        n = self.parser.page_count()
        for d in range(min(10, (n + ppd - 1) // ppd)):
            text = self.parser.read_zone(d * ppd + page_idx, coords)
            if text.strip():
                return text
        return ""

    def _unique_copy_name(self, base: str) -> str:
        names = {t.name for t in self.templates}
        name = tr("{base} (копия)", base=base)
        n = 2
        while name in names:
            name = tr("{base} (копия {n})", base=base, n=n)
            n += 1
        return name

    def _copy_template(self):
        """Копия текущего шаблона: все поля, зоны, правила, формулы, отчёт.
        Открывается в конструкторе — переименуйте и поправьте. Отмена = копия не создаётся."""
        if self.template is None:
            _info(self, tr("Сначала выберите шаблон, который нужно скопировать."), tr("Шаблоны"))
            return
        src_name = self.template.name
        copy = Template.from_dict(self.template.to_dict())
        copy.name = self._unique_copy_name(src_name)
        dlg = TemplateBuilderDialog(template=copy,
                                    existing_names=[t.name for t in self.templates],
                                    parent=self, copy_of=src_name,
                                    sample_reader=self._sample_text)
        if dlg.exec() != _DLG_OK:
            return
        tpl = dlg.get_template()
        if tpl is None:
            return
        self.templates.append(tpl)
        save_templates(self.templates, tpl.name)
        self._apply_template(tpl)
        self._refresh_template_combo()
        self.status.showMessage(tr("Создан шаблон «{name}» — копия «{src_name}».", name=tpl.name, src_name=src_name))

    def _edit_template(self):
        if self.template is None:
            _info(self, tr("Сначала выберите или создайте шаблон."), tr("Шаблоны"))
            return
        others = [t.name for t in self.templates if t is not self.template]
        dlg = TemplateBuilderDialog(template=self.template, existing_names=others, parent=self,
                                    sample_reader=self._sample_text)
        if dlg.exec() != _DLG_OK:
            return
        tpl = dlg.get_template()
        if tpl is None:
            return
        i = self.templates.index(self.template)
        self.templates[i] = tpl
        save_templates(self.templates, tpl.name)
        self._apply_template(tpl)
        self._refresh_template_combo()
        self.status.showMessage(tr("Шаблон «{name}» обновлён.", name=tpl.name))

    def _delete_template(self):
        if self.template is None:
            return
        name = self.template.name
        if not _ask(self, tr("Удалить шаблон «{name}»?", name=name), tr("Удаление")):
            return
        self.templates = [t for t in self.templates if t is not self.template]
        nxt = self.templates[0] if self.templates else None
        save_templates(self.templates, nxt.name if nxt else "")
        self._apply_template(nxt)
        self._refresh_template_combo()
        self.status.showMessage(tr("Шаблон «{name}» удалён.", name=name))

    # ═════════════════════════════════════════════════════════════════════════
    # Применение шаблона / перестроение динамики
    # ═════════════════════════════════════════════════════════════════════════

    def _apply_template(self, tpl: Optional[Template]):
        self.template = tpl
        self.waybills = []
        self.tree.clear()
        # Папка склеивалась под другое число страниц на документ — пересобрать,
        # иначе выравнивание пустыми страницами не совпадёт с шаблоном.
        if (tpl is not None and self._src_folder
                and tpl.pages != self._merged_ppd):
            self._load_folder(self._src_folder, self._src_recursive, quiet=True)
        self._dictionaries = {}
        self._dict_paths = {}

        self._rebuild_manual_form()
        self._rebuild_dict_cards()
        self._rebuild_match_area()
        self._rebuild_table_columns()
        self._reload_zone_editor()
        self._update_mode_label()
        self._update_pdf_info()

        # Автозагрузка словарей по сохранённым путям
        if tpl is not None:
            for f in tpl.dict_fields:
                path = f.get("dict_path", "")
                if path and os.path.exists(path):
                    self._load_dict_file(f["name"], path, announce=False)

        save_templates(self.templates, tpl.name if tpl else "")
        if tpl is not None:
            self.status.showMessage(tr("Шаблон: {name}", name=tpl.name))

    def _require_template(self) -> bool:
        if self.template is None:
            _info(self, tr("Сначала создайте или выберите шаблон."), tr("Шаблон"))
            return False
        return True

    def _rebuild_manual_form(self):
        _clear_layout(self._manual_form)
        self._manual_entries = {}
        fields = self.template.manual_fields if self.template else []
        if not fields:
            hint = QLabel(tr("В текущем шаблоне нет полей для ручного ввода.")
                          if self.template else tr("Шаблон не выбран."))
            hint.setStyleSheet(f"color:{C_TEXT2}; font-size:10pt;")
            self._manual_form.addRow(hint)
            return
        auto = set(self.template.auto_number)
        for name in fields:
            lbl = QLabel(name + ":")
            lbl.setStyleSheet(f"color:{C_TEXT2};")
            if name in auto:
                sp = NoWheelSpinBox()
                sp.setRange(0, 99999999)
                sp.setValue(int(self.template.auto_start.get(name, 1)))
                sp.setFixedWidth(160)
                sp.valueChanged.connect(lambda v, n=name: self._on_auto_start(n, v))
                row = QHBoxLayout()
                row.addWidget(sp)
                hint = QLabel(tr("№ по порядку: первая строка получит этот номер, дальше +1. После сохранения в Excel число увеличится само."))
                hint.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
                row.addWidget(hint, 1)
                self._manual_form.addRow(lbl, row)
                self._manual_entries[name] = sp
            else:
                ed = QLineEdit()
                ed.setMinimumWidth(380)
                ed.editingFinished.connect(self._on_manual_changed)
                self._manual_form.addRow(lbl, ed)
                self._manual_entries[name] = ed

    def _on_auto_start(self, name: str, value: int):
        if self.template is None:
            return
        self.template.auto_start[name] = int(value)
        save_templates(self.templates, self.template.name)
        if self.waybills:
            self._refresh()

    def _on_manual_changed(self):
        if self.waybills:
            self._refresh()

    def _rebuild_dict_cards(self):
        _clear_layout(self._dict_card_lay)
        self._dict_btns = {}
        self._dict_lbls = {}
        dict_fields = self.template.dict_fields if self.template else []
        title = QLabel(tr("🗂  Словари для сопоставления"))
        title.setStyleSheet(f"color:{C_TEXT2}; font-weight:bold;")
        self._dict_card_lay.addWidget(title)
        if not dict_fields:
            hint = QLabel(tr("В шаблоне нет полей со словарём."))
            hint.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
            self._dict_card_lay.addWidget(hint)
            return
        for f in dict_fields:
            name = f["name"]
            row = QHBoxLayout()
            row.setSpacing(10)
            lbl = QLabel(f"{name}:")
            lbl.setStyleSheet(f"color:{C_TEXT}; font-weight:bold;")
            lbl.setMinimumWidth(140)
            row.addWidget(lbl)
            btn = QPushButton(tr("Загрузить словарь…"))
            btn.clicked.connect(lambda _=False, n=name: self._load_dict(n))
            row.addWidget(btn)
            self._dict_btns[name] = btn
            info = QLabel(tr("Не загружен"))
            info.setStyleSheet(f"color:{C_TEXT2};")
            row.addWidget(info, 1)
            self._dict_lbls[name] = info
            self._dict_card_lay.addLayout(row)

    def _rebuild_match_area(self):
        _clear_layout(self._tools_host_lay)
        self._match_btns = {}
        dict_fields = self.template.dict_fields if self.template else []
        if not dict_fields:
            hint = QLabel(tr("Нет полей со словарём. Отметьте «Словарь» у поля при создании шаблона."))
            hint.setStyleSheet(f"color:{C_TEXT2}; font-size:10pt;")
            self._tools_host_lay.addWidget(hint)
            return
        for f in dict_fields:
            name = f["name"]
            card = Card()
            c_lay = QHBoxLayout(card)
            c_lay.setContentsMargins(20, 12, 20, 12)
            lbl = QLabel(f"👤  {name}")
            lbl.setStyleSheet(f"color:{C_TEXT2}; font-weight:bold;")
            lbl.setMinimumWidth(200)
            c_lay.addWidget(lbl)
            btn = QPushButton(tr("🔍 Сопоставить…"))
            btn.clicked.connect(lambda _=False, n=name: self._match_field(n))
            c_lay.addWidget(btn)
            c_lay.addStretch()
            self._match_btns[name] = btn
            self._tools_host_lay.addWidget(card)
        self._update_match_buttons()

    def _rebuild_table_columns(self):
        cols = self.template.columns if self.template else []
        if not cols:
            self.tree.setColumnCount(1)
            self.tree.setHeaderLabels(["—"])
            return
        self.tree.setColumnCount(len(cols))
        self.tree.setHeaderLabels(cols)
        for i, name in enumerate(cols):
            self.tree.setColumnWidth(i, max(90, min(220, 40 + len(name) * 11)))

    def _reload_zone_editor(self):
        if self.template is None:
            self._zone_editor.load(None, [], {}, {}, {}, {}, 1)
            return
        fields = [
            {"name": f["name"], "page": f["page"],
             "clean": f.get("clean", ""), "post": f.get("post", ""), "regex": f.get("regex", ""),
             "type": f.get("type"), "fallback": f.get("fallback", False)}
            for f in self.template.extract_fields
        ]
        self._zone_editor.load(
            self.parser, fields,
            dict(self.template.zones), dict(self.template.fallback_zones),
            self.template.colors(), self.template.default_zones(),
            self.template.pages, page=0,
        )

    def _update_mode_label(self):
        if self.template is None:
            self._mode_lbl.setText("—")
        else:
            p = self.template.pages
            word = tr("страница") if p == 1 else (tr("страницы") if 2 <= p <= 4 else tr("страниц"))
            self._mode_lbl.setText(tr("{p} {word} на документ", p=p, word=word))

    def _update_pdf_info(self):
        if not self.parser:
            return
        n = self.parser.page_count()
        ppd = self.template.pages if self.template else 1
        cnt = (n + ppd - 1) // ppd
        mode = tr(" (по {ppd} стр.)", ppd=ppd) if ppd > 1 else ""
        txt = tr("Страниц: {n}  |  Документов{mode}: {cnt}", n=n, mode=mode, cnt=cnt)
        if self._src_folder:
            blanks = sum(1 for x in self._page_src if x is None)
            n_files = len({x[0] for x in self._page_src if x is not None})
            txt = tr("Файлов: {n_files}  |  ", n_files=n_files) + txt
            if blanks:
                txt += tr("  |  добавлено пустых страниц: {blanks}", blanks=blanks)
        self._pdf_info.setText(txt)

    # ═════════════════════════════════════════════════════════════════════════
    # Словари и сопоставление
    # ═════════════════════════════════════════════════════════════════════════

    def _load_dict(self, name: str):
        p, _ = QFileDialog.getOpenFileName(
            self, tr("Словарь для «{name}»", name=name), "", tr("Текст (*.txt);;Все (*.*)"))
        if not p:
            return
        self._load_dict_file(name, p, announce=True)
        # Запоминаем путь в шаблоне
        if self.template is not None:
            for f in self.template.extract_fields:
                if f["name"] == name:
                    f["dict_path"] = p
            save_templates(self.templates, self.template.name)

    def _load_dict_file(self, name: str, path: str, announce: bool = True):
        try:
            with open(path, encoding="utf-8") as f:
                lines = [ln.strip() for ln in f if ln.strip()]
        except Exception as e:
            if announce:
                _error(self, tr("Не удалось загрузить словарь:\n{e}", e=e))
            return
        self._dictionaries[name] = lines
        self._dict_paths[name] = path
        if name in self._dict_lbls:
            self._dict_lbls[name].setText(tr("{v} ({n} записей)", v=os.path.basename(path), n=len(lines)))
        self._update_match_buttons()
        if announce:
            self.status.showMessage(tr("Словарь «{name}»: загружено {n} записей.", name=name, n=len(lines)))

    def _update_match_buttons(self):
        for name, btn in self._match_btns.items():
            has = bool(self._dictionaries.get(name))
            btn.setEnabled(has)
            btn.setToolTip("" if has else tr("Сначала загрузите словарь на вкладке «Распознавание»."))

    def _match_field(self, name: str):
        if not self.waybills:
            _warn(self, tr("Сначала извлеките данные."))
            return
        candidates = self._dictionaries.get(name, [])
        if not candidates:
            _warn(self, tr("Загрузите словарь для «{name}» на вкладке «Распознавание».", name=name))
            return
        dlg = MatchDialog(
            self.waybills, name, candidates,
            pages=self.template.pages if self.template else 1,
            pdf_path=self.pdf_path or "", parser=self.parser, parent=self,
        )
        dlg.new_entry.connect(lambda v, n=name: self._on_new_dict_entry(n, v))
        if dlg.exec() == _DLG_OK:
            for i, val in enumerate(dlg.get_results()):
                if i < len(self.waybills):
                    self.waybills[i][name] = val
            self._refresh()
            self.status.showMessage(tr("Поле «{name}» обновлено по словарю.", name=name))

    def _auto_match_field(self, name: str):
        candidates = self._dictionaries.get(name, [])
        if not candidates or not self.waybills:
            return 0
        replaced = 0
        for wb in self.waybills:
            best, score = _best_match(wb.get(name, ""), candidates)
            if score >= AUTO_MATCH_THRESHOLD:
                wb[name] = best
                replaced += 1
        return replaced

    def _on_new_dict_entry(self, name: str, value: str):
        lst = self._dictionaries.setdefault(name, [])
        if value in lst:
            return
        lst.append(value)
        path = self._dict_paths.get(name)
        if path:
            try:
                with open(path, "a", encoding="utf-8") as f:
                    f.write("\n" + value)
            except Exception as e:
                _warn(self, tr("Не удалось дописать словарь:\n{e}", e=e), tr("Ошибка записи"))
            if name in self._dict_lbls:
                self._dict_lbls[name].setText(tr("{v} ({n} записей)", v=os.path.basename(path), n=len(lst)))
        self.status.showMessage(tr("В словарь «{name}» добавлено: {value}", name=name, value=value))

    # ═════════════════════════════════════════════════════════════════════════
    # Зоны
    # ═════════════════════════════════════════════════════════════════════════

    def _on_zones_changed(self, payload: dict):
        if self.template is None:
            return
        zones = payload.get("zones", {})
        fallback = payload.get("fallback", {})
        valid = {f["name"] for f in self.template.extract_fields}
        # Обновляем зоны только для существующих полей
        for k, v in zones.items():
            if k in valid:
                self.template.zones[k] = tuple(v)
        for k, v in fallback.items():
            if k in valid:
                self.template.fallback_zones[k] = tuple(v)
        if self.parser:
            self.parser.set_zones(self.template.zones)
        save_templates(self.templates, self.template.name)
        self.status.showMessage(tr("Зоны обновлены и сохранены."))


class _ExtractMixin:
    """Открытие PDF/Excel и фоновое извлечение данных."""

    # ═════════════════════════════════════════════════════════════════════════
    # PDF / Excel
    # ═════════════════════════════════════════════════════════════════════════

    def _open_pdf(self):
        if not self._require_template():
            return
        p, _ = QFileDialog.getOpenFileName(self, tr("Выберите PDF"), "", tr("PDF (*.pdf);;Все (*.*)"))
        if not p:
            return
        self._load_pdf_path(p)

    def _load_pdf_path(self, p: str, check_text: bool = True) -> bool:
        """Открывает один PDF-файл как источник. True — если открылся."""
        try:
            self._set_pdf_source(p)
            self._drop_merged_temp()
            self._src_folder = None
            self._src_files = []
            self._page_src = []
            self._pdf_lbl.setText(os.path.basename(p))
            self._pdf_lbl.setStyleSheet(f"color:{C_SUCCESS};")
            self._pdf_lbl.setToolTip(p)
            self._update_pdf_info()
            self.status.showMessage(tr("Открыт: {p}", p=p))
            self._reload_zone_editor()
            if check_text:
                self._warn_no_text()
            return True
        except Exception as e:
            _error(self, tr("Не удалось открыть PDF:\n{e}", e=e))
            return False

    def _set_pdf_source(self, path: str):
        """Делает path текущим PDF: закрывает старый парсер, открывает новый."""
        if self.parser:
            self.parser.close()
            self.parser = None
        dpi = PdfParser.BASE_DPI
        self.parser = PdfParser(path, dpi=dpi, zones=dict(self.template.zones))
        self.pdf_path = path

    def _drop_merged_temp(self):
        """Удаляет прошлый склеенный временный PDF (если он больше не открыт)."""
        t = self._merged_temp
        if t and t != self.pdf_path:
            try:
                os.remove(t)
            except OSError:
                pass
        if t != self.pdf_path:
            self._merged_temp = None

    def _open_folder(self):
        if not self._require_template():
            return
        folder = QFileDialog.getExistingDirectory(self, tr("Папка с PDF-файлами"))
        if not folder:
            return
        recursive = False
        if _has_pdfs_in_subfolders(folder):
            recursive = _ask(self, tr("В подпапках тоже есть PDF.\nВключить их в обработку?"),
                             tr("Подпапки"))
        self._load_folder(folder, recursive)

    def _load_folder(self, folder: str, recursive: bool = False, quiet: bool = False):
        """Склеивает все PDF папки в один временный файл и открывает его."""
        files = _list_pdfs(folder, recursive)
        if not files:
            if not quiet:
                _warn(self, tr("В выбранной папке нет PDF-файлов."))
            return
        ppd = self.template.pages if self.template else 1

        dlg = QProgressDialog(tr("Склеиваю PDF…"), tr("Отмена"), 0, len(files), self)
        dlg.setWindowTitle(tr("Папка с PDF"))
        dlg.setWindowModality(_WIN_MODAL)
        dlg.setMinimumDuration(300)

        def _progress(i, total, name):
            dlg.setValue(i)
            dlg.setLabelText(tr("Файл {v} из {total}:\n{name}", v=i + 1, total=total, name=name))
            QApplication.processEvents()
            return not dlg.wasCanceled()

        try:
            res = _merge_pdfs(files, ppd, _progress)
        except InterruptedError:
            dlg.close()
            self.status.showMessage(tr("Загрузка папки отменена."))
            return
        except Exception as e:
            dlg.close()
            _error(self, tr("Не удалось склеить PDF из папки:\n{e}", e=e))
            return
        dlg.setValue(len(files))
        dlg.close()

        old_temp = self._merged_temp
        try:
            self._set_pdf_source(res["path"])
        except Exception as e:
            _error(self, tr("Не удалось открыть склеенный PDF:\n{e}", e=e))
            return
        if old_temp and old_temp != res["path"]:
            try:
                os.remove(old_temp)
            except OSError:
                pass
        self._merged_temp = res["path"]
        self._src_folder = folder
        self._src_recursive = recursive
        self._src_files = files
        self._merged_ppd = ppd
        self._page_src = res["page_src"]

        ok_files = len(files) - len(res["failed"])
        self._pdf_lbl.setText(tr("📁 {v}  (файлов: {ok_files})", v=os.path.basename(folder) or folder, ok_files=ok_files))
        self._pdf_lbl.setStyleSheet(f"color:{C_SUCCESS};")
        self._pdf_lbl.setToolTip(folder)
        self._update_pdf_info()
        self._reload_zone_editor()
        self.status.showMessage(
            tr("Папка: {folder}  —  файлов {ok_files}, страниц {v}", folder=folder, ok_files=ok_files, v=self.parser.page_count()))

        if quiet:
            return
        self._warn_no_text()
        notes = []
        if res["padded"]:
            lst = "\n".join(f"  • {n}: +{k}" for n, k in res["padded"][:15])
            more = tr("\n  … и ещё {v}", v=len(res['padded']) - 15) if len(res["padded"]) > 15 else ""
            notes.append(tr("Число страниц не кратно {ppd} — добавлены пустые страницы, чтобы каждый файл начинался с нового документа:\n{lst}{more}", ppd=ppd, lst=lst, more=more))
        if res["failed"]:
            lst = "\n".join(f"  • {n}: {e}" for n, e in res["failed"][:15])
            notes.append(tr("Пропущены (не открылись):\n{lst}", lst=lst))
        if notes:
            _warn(self, "\n\n".join(notes), tr("Папка загружена с замечаниями"))

    def _warn_no_text(self):
        """Предупреждает, если в PDF нет распознанного текста (скан без OCR)."""
        if self.parser is None:
            return
        doc = self.parser.doc
        n = len(doc)
        try:
            has = any(doc[i].get_text("text").strip() for i in sorted({0, n // 2, n - 1}) if n)
        except Exception:
            return
        if not has:
            _warn(self, tr("В PDF не найден распознанный текст.\n\nПрограмма читает только текстовый слой. Распознайте скан в ABBYY FineReader («Сохранить как → PDF с возможностью поиска») и откройте файл снова."),
                  tr("Нет текстового слоя"))

    def _row_source(self, row: int) -> str:
        """«файл.pdf, стр. N» для строки таблицы (только для источника-папки)."""
        if not self._page_src or self.template is None:
            return ""
        for pg in self._row_pages(row):
            if pg >= len(self._page_src):
                continue
            src = self._page_src[pg]
            if src is not None:
                return tr("{v}, стр. {v2}", v=src[0], v2=src[1])
        return ""

    def closeEvent(self, event):
        if self.parser:
            self.parser.close()
            self.parser = None
        if self._merged_temp:
            try:
                os.remove(self._merged_temp)
            except OSError:
                pass
        super().closeEvent(event)

    def _open_excel(self):
        p, _ = QFileDialog.getOpenFileName(self, tr("Выберите Excel"), "", tr("Excel (*.xlsx *.xls);;Все (*.*)"))
        if not p:
            return
        self.excel_path = p
        try:
            df = pd.read_excel(p)
            self.tree.clear()
            self.tree.setColumnCount(len(df.columns))
            self.tree.setHeaderLabels([str(c) for c in df.columns])
            for _, row in df.iterrows():
                self.tree.addTopLevelItem(QTreeWidgetItem([str(v) for v in row]))
            self.status.showMessage(tr("Загружено {n} записей из {p}", n=len(df), p=p))
        except Exception as e:
            _error(self, tr("Не удалось открыть Excel:\n{e}", e=e))


    # ═════════════════════════════════════════════════════════════════════════
    # Извлечение
    # ═════════════════════════════════════════════════════════════════════════

    def _start_extract(self):
        if not self._require_template():
            return
        if not self.parser:
            _warn(self, tr("Сначала выберите PDF файл."))
            return
        self.parser.set_zones(self.template.zones)
        self._run_btn.setEnabled(False)
        self._prog.setValue(0)

        self._worker = ExtractionWorker(
            self.parser, list(self.template.extract_fields),
            self.template.pages, dict(self.template.fallback_zones))
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_extract_done)
        self._worker.error.connect(self._on_extract_error)
        self._thread.start()

    def _on_progress(self, done: int, total: int):
        self._prog.setMaximum(max(1, total))
        self._prog.setValue(done)
        self._prog_lbl.setText(f"{done} / {total}")
        self.status.showMessage(tr("Обработано: {done} / {total}", done=done, total=total))

    def _on_extract_done(self, result: list):
        self.waybills = result
        self._teardown_thread()
        self._run_btn.setEnabled(True)

        # Автосопоставление по загруженным словарям
        for f in (self.template.dict_fields if self.template else []):
            if self._dictionaries.get(f["name"]):
                self._auto_match_field(f["name"])

        self._refresh()
        self.tabs.setCurrentIndex(4)
        probs = len(self._problem_rows())
        if probs:
            _info(self, tr("Извлечено записей: {n}.\nПроблемных строк: {probs}.\n\nКурсор стоит на первой проблеме. Справа — участок скана.\nВпишите правильное значение и нажмите Enter — откроется следующая.\nF3 — следующая проблема без правки.", n=len(self.waybills), probs=probs))
            self._jump_next_problem(quiet=True)
        else:
            _info(self, tr("Извлечено записей: {n}.\nПроблем не найдено.", n=len(self.waybills)))

    def _on_extract_error(self, msg: str):
        self._teardown_thread()
        self._run_btn.setEnabled(True)
        _error(self, msg, tr("Ошибка извлечения"))

    def _teardown_thread(self):
        """Останавливает поток извлечения и освобождает объекты worker/thread."""
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait()
            self._thread.deleteLater()
            self._thread = None
        if self._worker is not None:
            self._worker.deleteLater()
            self._worker = None


def _norm_key(v) -> str:
    """Ключ для поиска дублей: без пробелов, без учёта регистра, латиница→кириллица."""
    return re.sub(r"\s+", "", str(v if v is not None else "")).upper().translate(_LAT_TO_CYR)


def _fmt_total(d: Decimal) -> str:
    """1234.5 → «1 234,5» (до 3 знаков после запятой)."""
    s = format(d.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP), ",f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if DEC_SEP == ".":
        return s
    return s.replace(",", "\u00a0").replace(".", ",")


def _choose(parent, title: str, text: str, buttons: list) -> int:
    """Вопрос с несколькими вариантами. Последняя кнопка — «отмена».
    Возвращает индекс нажатой кнопки."""
    box = QMessageBox(parent)
    box.setWindowTitle(title)
    box.setText(text)
    box.setIcon(_qenum(QMessageBox, "Icon", "Question"))
    accept = _qenum(QMessageBox, "ButtonRole", "AcceptRole")
    reject = _qenum(QMessageBox, "ButtonRole", "RejectRole")
    btns = [box.addButton(b, reject if i == len(buttons) - 1 else accept)
            for i, b in enumerate(buttons)]
    box.setDefaultButton(btns[0])
    box.exec()
    clicked = box.clickedButton()
    return btns.index(clicked) if clicked in btns else len(buttons) - 1


def _excel_value(v: str, numeric: bool):
    """Значение для ячейки Excel: числовые поля — числом, остальное — текстом."""
    if not numeric:
        return v
    d = _to_decimal(v)
    if d is None:
        return v
    return int(d) if d == d.to_integral_value() else float(d)


class _TableMixin:
    """Таблица результатов: проверка, дубли, итоги, панель скана, правка, экспорт."""

    # ═════════════════════════════════════════════════════════════════════════
    # Значения строк
    # ═════════════════════════════════════════════════════════════════════════

    def _manual_vals(self, idx: Optional[int] = None) -> dict:
        """Значения ручных полей. Для «№ по порядку» — номер строки idx."""
        auto = set(self.template.auto_number) if self.template else set()
        out = {}
        for k, w in self._manual_entries.items():
            if k in auto:
                out[k] = str(w.value() + idx) if idx is not None else ""
            else:
                out[k] = w.text()
        return out

    @staticmethod
    def _cell_text(item: QTreeWidgetItem, col: int) -> str:
        """Реальное значение ячейки (экранный маркер «∅ нет данных» = пусто)."""
        t = item.text(col)
        return "" if t == missing_mark() else t

    def _row_values(self, idx: int) -> dict:
        """Все значения строки idx: ручные (с правками строки), из PDF, формулы."""
        wb = self.waybills[idx]
        manual = self.template.manual_fields
        mv = self._manual_vals(idx)
        vals = {c: (wb[c] if c in wb else mv.get(c, "")) for c in manual}
        vals.update({k: v for k, v in wb.items() if k not in vals})
        for c in self.template.computed_fields:
            vals[c["name"]] = eval_formula(c["formula"], vals, c["decimals"])
        return vals

    def _row_pages(self, row: int) -> list:
        """Страницы PDF документа строки (с учётом удалённых строк)."""
        ppd = self.template.pages if self.template else 1
        total = self.parser.page_count() if self.parser else 0
        if 0 <= row < len(self.waybills) and "_page" in self.waybills[row]:
            start = int(self.waybills[row]["_page"])
            return [start + k for k in range(ppd) if 0 <= start + k < total]
        return _pages_for_row(row, ppd, total)

    def _checked_fields(self) -> list:
        """(поле, шаблон проверки) — поля с подсветкой: из PDF и формулы."""
        if self.template is None:
            return []
        out = [(f["name"], resolve_check_pattern(f)) for f in self.template.extract_fields]
        out += [(c["name"], "") for c in self.template.computed_fields]
        return out

    def _current_cell(self):
        it = self.tree.currentItem()
        if it is None:
            return None
        return (self.tree.indexOfTopLevelItem(it), max(0, self.tree.currentColumn()))

    # ═════════════════════════════════════════════════════════════════════════
    # Отрисовка таблицы
    # ═════════════════════════════════════════════════════════════════════════

    def _apply_validation(self, item: QTreeWidgetItem, alt_bg: QColor) -> tuple:
        """Жёлтый ∅ — пусто; красный — не прошло проверку. Возвращает (пустых, ошибок)."""
        if self.template is None:
            return (0, 0)
        cmap = self.template.column_index()
        missing = errors = 0
        for fname, pattern in self._checked_fields():
            col = cmap.get(fname, -1)
            if col < 0:
                continue
            value = self._cell_text(item, col)
            fnt = item.font(col)
            item.setTextAlignment(col, _ALIGN_LEFT | _ALIGN_VCENTER)
            if value.strip() == "":
                item.setText(col, missing_mark())
                item.setBackground(col, QBrush(QColor(C_EMPTY_BG)))
                item.setForeground(col, QBrush(QColor(C_EMPTY_FG)))
                fnt.setBold(True)
                missing += 1
            elif _field_valid(value, pattern):
                item.setBackground(col, QBrush(alt_bg))
                item.setForeground(col, QBrush())
                fnt.setBold(False)
            else:
                item.setBackground(col, QBrush(QColor(C_ERR_BG)))
                item.setForeground(col, QBrush(QColor(C_ERR_FG)))
                fnt.setBold(True)
                errors += 1
            item.setFont(col, fnt)
        return (missing, errors)

    def _mark_duplicates(self) -> int:
        """Подсвечивает повторы ключевого поля в пачке. Возвращает число строк-повторов."""
        self._dup_rows = {}
        if self.template is None or not self.template.key_field:
            return 0
        col = self.template.column_index().get(self.template.key_field)
        if col is None:
            return 0
        groups: dict = {}
        for i in range(self.tree.topLevelItemCount()):
            k = _norm_key(self._cell_text(self.tree.topLevelItem(i), col))
            if k:
                groups.setdefault(k, []).append(i)
        for rows in groups.values():
            if len(rows) < 2:
                continue
            tip = tr("Повтор в пачке: строки ") + ", ".join(str(r + 1) for r in rows)
            for r in rows:
                it = self.tree.topLevelItem(r)
                it.setBackground(col, QBrush(QColor(T.dup_bg)))
                it.setForeground(col, QBrush(QColor(T.dup_fg)))
                fnt = it.font(col)
                fnt.setBold(True)
                it.setFont(col, fnt)
                it.setToolTip(col, tip)
                self._dup_rows[r] = col
        return len(self._dup_rows)

    def _refresh(self, keep=None):
        """Перестраивает таблицу из данных. keep=(строка, столбец) — куда вернуть курсор."""
        if keep is None:
            keep = self._current_cell()
        self.tree.clear()
        self._dup_rows = {}
        if self.template is None:
            self._update_totals()
            self._update_scan_panel()
            return
        cols = self.template.columns
        manual = set(self.template.manual_fields)
        total_missing = total_err = 0
        for idx, wb in enumerate(self.waybills):
            vals = self._row_values(idx)
            item = QTreeWidgetItem([str(vals.get(c, "")) for c in cols])
            item.setSizeHint(0, QSize(0, 26))
            src = self._row_source(idx)
            alt_bg = QColor(C_CARD) if idx % 2 == 0 else QColor(C_ALT_ROW)
            for c, name in enumerate(cols):
                item.setBackground(c, QBrush(alt_bg))
                tip = tr("Источник: {src}", src=src) if src else ""
                if name in manual and name in wb:
                    tip = (tip + "\n" if tip else "") + tr("Значение изменено только для этой строки")
                if tip:
                    item.setToolTip(c, tip)
            m, e = self._apply_validation(item, alt_bg)
            total_missing += m
            total_err += e
            self.tree.addTopLevelItem(item)
        dups = self._mark_duplicates()

        parts = [tr("Всего записей: {n}", n=len(self.waybills))]
        if total_missing:
            parts.append(tr("без значения: {total_missing} (жёлтые ∅)", total_missing=total_missing))
        if total_err:
            parts.append(tr("ошибок проверки: {total_err} (красные)", total_err=total_err))
        if dups:
            parts.append(tr("повторов «{key_field}»: {dups} (фиолетовые)", key_field=self.template.key_field, dups=dups))
        self.status.showMessage("  |  ".join(parts))
        self._update_totals()

        if keep and 0 <= keep[0] < self.tree.topLevelItemCount():
            self.tree.setCurrentItem(self.tree.topLevelItem(keep[0]), keep[1])
            self.tree.scrollToItem(self.tree.topLevelItem(keep[0]))
        self._update_scan_panel()

    def _update_totals(self):
        lbl = getattr(self, "_totals_lbl", None)
        if lbl is None:
            return
        n = self.tree.topLevelItemCount()
        if self.template is None or n == 0:
            lbl.setText("")
            return
        hmap = {self.tree.headerItem().text(c): c for c in range(self.tree.columnCount())}
        parts = [tr("Записей: <b>{n}</b>", n=n)]
        for name in self.template.sum_names():
            c = hmap.get(name)
            if c is None:
                continue
            total, cnt = Decimal(0), 0
            for i in range(n):
                d = _to_decimal(self._cell_text(self.tree.topLevelItem(i), c))
                if d is not None:
                    total += d
                    cnt += 1
            gap = tr(" <span style='color:{C_WARNING}'>(без {v})</span>", C_WARNING=C_WARNING, v=n - cnt) if cnt < n else ""
            parts.append(f"Σ {name}: <b>{_fmt_total(total)}</b>{gap}")
        problems = len(self._problem_rows())
        if problems:
            parts.append(tr("<span style='color:{C_DANGER}'>⚠ проблемных строк: {problems} — F3</span>", C_DANGER=C_DANGER, problems=problems))
        else:
            parts.append(tr("<span style='color:{C_SUCCESS}'>✓ проблем нет</span>", C_SUCCESS=C_SUCCESS))
        lbl.setText("&nbsp;&nbsp;•&nbsp;&nbsp;".join(parts))

    # ═════════════════════════════════════════════════════════════════════════
    # Проблемные строки
    # ═════════════════════════════════════════════════════════════════════════

    def _problem_rows(self) -> list:
        """[(строка, столбец, тип)]: missing — пусто, error — ошибка, dup — повтор."""
        out = []
        if self.template is None:
            return out
        cmap = self.template.column_index()
        dups = getattr(self, "_dup_rows", {})
        for i in range(self.tree.topLevelItemCount()):
            it = self.tree.topLevelItem(i)
            found = None
            for fname, pattern in self._checked_fields():
                col = cmap.get(fname, -1)
                if col < 0:
                    continue
                value = self._cell_text(it, col)
                if value.strip() == "":
                    found = (i, col, "missing")
                    break
                if not _field_valid(value, pattern):
                    found = (i, col, "error")
                    break
            if found is None and i in dups:
                found = (i, dups[i], "dup")
            if found:
                out.append(found)
        return out

    def _jump_next_problem(self, quiet: bool = False, focus_tree: bool = True):
        """Переход к следующей проблемной строке (по кругу)."""
        if self.tree.topLevelItemCount() == 0:
            self.status.showMessage(tr("Таблица пуста."))
            return
        problems = self._problem_rows()
        if not problems:
            self.status.showMessage(tr("Проблем нет: все ячейки заполнены, проверены, повторов нет."))
            if not quiet:
                _info(self, tr("Все ячейки заполнены, проходят проверку, повторов нет."), tr("Проблем нет"))
            return
        cur = self._current_cell()
        cur_row = cur[0] if cur else -1
        target = next((p for p in problems if p[0] > cur_row), problems[0])
        idx, col, kind = target
        item = self.tree.topLevelItem(idx)
        self.tree.setCurrentItem(item, col)
        self.tree.scrollToItem(item, _SCROLL_CENTER)
        if focus_tree:
            self.tree.setFocus()
        kind_txt = {"missing": tr("нет данных"), "error": tr("ошибка проверки"),
                    "dup": tr("повтор номера")}[kind]
        self.status.showMessage(tr("Проблема {v} из {n}  —  строка {v2}  ({kind_txt})", v=problems.index(target) + 1, n=len(problems), v2=idx + 1, kind_txt=kind_txt))

    # ═════════════════════════════════════════════════════════════════════════
    # Правка значений
    # ═════════════════════════════════════════════════════════════════════════

    def _computed_formula(self, name: str) -> Optional[str]:
        if self.template is None:
            return None
        for c in self.template.computed_fields:
            if c["name"] == name:
                return c["formula"]
        return None

    def _set_value(self, row: int, col: int, value: str):
        """Записывает значение ячейки в данные и перерисовывает таблицу."""
        name = self.tree.headerItem().text(col)
        if self.template is not None and 0 <= row < len(self.waybills):
            self.waybills[row][name] = value
            self._refresh(keep=(row, col))
        else:
            it = self.tree.topLevelItem(row)
            if it is not None:
                it.setText(col, value)

    def _edit_cell(self, item: QTreeWidgetItem, column: int = None):
        if column is None:
            column = 0
        col_name = self.tree.headerItem().text(column)
        formula = self._computed_formula(col_name)
        if formula is not None:
            _info(self, tr("«{col_name}» считается по формуле:\n{formula}\n\nИзмените исходное поле — значение пересчитается.", col_name=col_name, formula=formula), tr("Вычисляемое поле"))
            return
        dlg = EditCellDialog(col_name, self._cell_text(item, column), self)
        if dlg.exec() != _DLG_OK:
            return
        self._set_value(self.tree.indexOfTopLevelItem(item), column, dlg.value())
        self.status.showMessage(tr("Значение изменено."))

    def _del_row(self):
        sel = self.tree.selectedItems()
        if not sel:
            _warn(self, tr("Выберите строку."))
            return
        indices = sorted({self.tree.indexOfTopLevelItem(it) for it in sel}, reverse=True)
        if self.template is not None and self.waybills:
            for idx in indices:
                if 0 <= idx < len(self.waybills):
                    del self.waybills[idx]
            self._refresh(keep=(min(indices), 0))
        else:
            for idx in indices:
                self.tree.takeTopLevelItem(idx)
        self.status.showMessage(tr("Удалено строк: {n}.", n=len(indices)))

    def _clear_all(self):
        if _ask(self, tr("Очистить все данные?")):
            self.tree.clear()
            self.waybills.clear()
            self._update_totals()
            self._update_scan_panel()
            self.status.showMessage(tr("Данные очищены."))

    def _context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        menu.addAction(tr("👁 Просмотреть документ"), self._view_selected)
        menu.addSeparator()
        col = self.tree.currentColumn()
        if col >= 0:
            cname = self.tree.headerItem().text(col)
            menu.addAction(tr("✏ Редактировать «{cname}»", cname=cname), lambda: self._edit_cell(item, col))
        menu.addSeparator()
        menu.addAction(tr("🗑 Удалить строку"), self._del_row)
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    def _view_selected(self):
        if not self.pdf_path:
            _warn(self, tr("Сначала откройте PDF файл."))
            return
        cur = self._current_cell()
        if cur is None:
            _warn(self, tr("Выберите строку."))
            return
        row = cur[0]
        pages = self._row_pages(row)
        if not pages:
            _warn(self, tr("Страница не существует."))
            return
        try:
            temp = _extract_pages_to_temp(self.pdf_path, pages, tr("строка_{v}", v=row + 1))
            _open_pdf_externally(temp)
            src = self._row_source(row)
            self.status.showMessage(tr("Открыт документ строки {v}", v=row + 1)
                                    + (f"  —  {src}" if src else ""))
        except Exception as e:
            _error(self, tr("Не удалось открыть документ:\n{e}", e=e))

    # ═════════════════════════════════════════════════════════════════════════
    # Панель скана: вырезка зоны рядом с таблицей
    # ═════════════════════════════════════════════════════════════════════════

    def _build_scan_panel(self) -> QWidget:
        panel = QFrame()
        panel.setMinimumWidth(280)
        panel.setStyleSheet(f"QFrame#scanPanel{{background:{C_CARD}; border-left:1px solid {C_BG3};}}")
        panel.setObjectName("scanPanel")
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(8)
        self._scan_title = QLabel(tr("Скан ячейки"))
        self._scan_title.setWordWrap(True)
        self._scan_title.setStyleSheet(f"color:{C_TEXT}; font-weight:bold; font-size:10pt;")
        lay.addWidget(self._scan_title)
        self._scan_img = QLabel()
        self._scan_img.setAlignment(_ALIGN_CENTER)
        self._scan_img.setMinimumHeight(90)
        self._scan_img.setStyleSheet(f"background:{C_BG}; border:1px solid {C_BG3}; border-radius:6px;")
        lay.addWidget(self._scan_img)
        self._scan_raw = QLabel("")
        self._scan_raw.setWordWrap(True)
        self._scan_raw.setTextInteractionFlags(_qenum(Qt, "TextInteractionFlag", "TextSelectableByMouse"))
        self._scan_raw.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
        lay.addWidget(self._scan_raw)
        lay.addWidget(QLabel(tr("Значение в таблице:")))
        self._scan_edit = QLineEdit()
        self._scan_edit.setPlaceholderText(tr("впишите правильное значение и нажмите Enter"))
        self._scan_edit.returnPressed.connect(self._scan_apply)
        lay.addWidget(self._scan_edit)
        btn = QPushButton(tr("✔ Применить → следующая проблема"))
        btn.setObjectName("primary")
        btn.clicked.connect(self._scan_apply)
        lay.addWidget(btn)
        hint = QLabel(tr("Enter — сохранить и перейти к следующей проблеме.\nF3 — следующая проблема без правки.\nСиняя рамка — основная зона, оранжевая пунктирная — резервная."))
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
        lay.addWidget(hint)
        lay.addStretch()
        self._scan_pix: Optional[QPixmap] = None
        self._update_scan_panel()
        return panel

    def _zone_crop(self, page_num: int, coords_list: list) -> Optional[QPixmap]:
        """Вырезка страницы вокруг зон (координаты 300 DPI) с рамками зон."""
        if self.parser is None or not (0 <= page_num < self.parser.page_count()):
            return None
        page = self.parser.doc[page_num]
        rects = [fitz.Rect(*[v / PdfParser.BASE_DPI * 72 for v in c]) for c in coords_list if c]
        if not rects:
            return None
        union = fitz.Rect(rects[0])
        for r in rects[1:]:
            union |= r
        m = 16
        clip = fitz.Rect(union.x0 - m, union.y0 - m, union.x1 + m, union.y1 + m) & page.rect
        if clip.is_empty:
            return None
        zoom = 2.5
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, _IMG_RGB888).copy()
        pm = QPixmap.fromImage(img)
        p = QPainter(pm)
        for i, r in enumerate(rects):
            pen = QPen(QColor(T.accent if i == 0 else T.warning), 2)
            pen.setStyle(_SOLID_PEN if i == 0 else _DASH_PEN)
            p.setPen(pen)
            p.drawRect(QRectF((r.x0 - clip.x0) * zoom, (r.y0 - clip.y0) * zoom,
                              r.width * zoom, r.height * zoom))
        p.end()
        return pm

    def _show_scan_pix(self):
        if self._scan_pix is None:
            self._scan_img.clear()
            return
        w = max(120, self._scan_img.width() - 8)
        pm = self._scan_pix
        if pm.width() > w:
            pm = pm.scaledToWidth(w, _SMOOTH)
        self._scan_img.setPixmap(pm)

    def _update_scan_panel(self, *_):
        if getattr(self, "_scan_title", None) is None:
            return
        self._scan_pix = None
        cur = self._current_cell()
        if self.template is None or cur is None or self.tree.topLevelItemCount() == 0:
            self._scan_title.setText(tr("Скан ячейки"))
            self._scan_raw.setText(tr("Выберите ячейку в таблице — здесь появится этот участок скана и распознанный текст."))
            self._scan_edit.clear()
            self._scan_edit.setEnabled(False)
            self._show_scan_pix()
            return
        row, col = cur
        item = self.tree.topLevelItem(row)
        name = self.tree.headerItem().text(col)
        self._scan_title.setText(tr("Строка {v}  •  {name}", v=row + 1, name=name))
        self._scan_edit.setEnabled(True)
        self._scan_edit.setText(self._cell_text(item, col))

        formula = self._computed_formula(name)
        field = next((f for f in self.template.extract_fields if f["name"] == name), None)
        if formula is not None:
            self._scan_edit.setEnabled(False)
            self._scan_raw.setText(tr("Считается по формуле: {formula}\nИсправьте исходное поле — значение пересчитается.", formula=formula))
        elif field is None:
            self._scan_raw.setText(tr("Поле заполняется вручную (вкладка «1. Общие данные»). Здесь можно задать значение только для этой строки."))
        elif not self.parser or row >= len(self.waybills):
            self._scan_raw.setText(tr("PDF не открыт — скан недоступен."))
        else:
            pages = self._row_pages(row)
            pg = (pages[0] + field["page"]) if pages else -1
            coords = [self.template.zones.get(name)]
            fb = self.template.fallback_zones.get(name) if field.get("fallback") else None
            if fb:
                coords.append(fb)
            try:
                self._scan_pix = self._zone_crop(pg, coords)
                raw = self.parser.read_zone(pg, coords[0])
                txt = tr("Распознанный текст зоны (стр. {v}):\n«{v2}»", v=pg + 1, v2=raw or '—')
                if fb:
                    txt += tr("\nРезервная зона: «{v}»", v=self.parser.read_zone(pg, fb) or '—')
            except Exception as e:
                txt = tr("Не удалось показать скан: {e}", e=e)
            self._scan_raw.setText(txt)
        self._show_scan_pix()

    def _scan_apply(self):
        cur = self._current_cell()
        if cur is None or not self._scan_edit.isEnabled():
            return
        row, col = cur
        self._set_value(row, col, self._scan_edit.text().strip())
        if self._problem_rows():
            self._jump_next_problem(quiet=True, focus_tree=False)
        self._scan_edit.setFocus()
        self._scan_edit.selectAll()

    def _toggle_scan_panel(self, on: bool):
        self._scan_panel.setVisible(on)
        if on:
            self._update_scan_panel()

    # ═════════════════════════════════════════════════════════════════════════
    # Экспорт
    # ═════════════════════════════════════════════════════════════════════════

    def _table_data(self) -> tuple:
        cols = [self.tree.headerItem().text(c) for c in range(self.tree.columnCount())]
        data = [[self._cell_text(it, c) for c in range(len(cols))]
                for i in range(self.tree.topLevelItemCount())
                if (it := self.tree.topLevelItem(i)) is not None]
        return cols, data

    def _save_excel(self):
        """Дописывает строки в Excel. Старые данные и оформление файла сохраняются."""
        import openpyxl
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter

        if self.tree.topLevelItemCount() == 0:
            _warn(self, tr("Нет данных для сохранения."))
            return
        if not self.excel_path:
            p, _ = QFileDialog.getSaveFileName(self, tr("Сохранить Excel"), tr("документы.xlsx"), "Excel (*.xlsx)")
            if not p:
                return
            self.excel_path = p
        path = self.excel_path
        cols, data = self._table_data()
        numeric = self.template.numeric_names() if self.template else set()
        key = self.template.key_field if self.template else ""
        kidx = cols.index(key) if key in cols else -1

        try:
            is_new = not os.path.exists(path)
            if is_new:
                book = openpyxl.Workbook()
                ws = book.active
                ws.title = tr("Данные")
            else:
                book = openpyxl.load_workbook(path)
                ws = book.active
                header = [("" if c.value is None else str(c.value)).strip() for c in ws[1]]
                while header and not header[-1]:
                    header.pop()
                if not header:
                    is_new = True
                elif header != cols:
                    ch = _choose(self, tr("Столбцы не совпадают"),
                                 tr("В файле «{v}» другие столбцы:\n  в файле:   {v2}\n  в таблице: {v3}\n\nДописать нельзя — данные съедут по столбцам.", v=os.path.basename(path), v2=', '.join(header), v3=', '.join(cols)),
                                 [tr("Сохранить в другой файл…"), tr("Отмена")])
                    if ch != 0:
                        return
                    p, _ = QFileDialog.getSaveFileName(self, tr("Сохранить Excel"), tr("документы.xlsx"),
                                                       "Excel (*.xlsx)")
                    if p:
                        self.excel_path = p
                        self._save_excel()
                    return
                elif kidx >= 0:
                    existing = {_norm_key(ws.cell(r, kidx + 1).value) for r in range(2, ws.max_row + 1)}
                    existing.discard("")
                    rep_rows = [i for i, row in enumerate(data) if _norm_key(row[kidx]) in existing]
                    if rep_rows:
                        sample = ", ".join(data[i][kidx] for i in rep_rows[:10])
                        more = " …" if len(rep_rows) > 10 else ""
                        ch = _choose(self, tr("Повторы в файле"),
                                     tr("{n} из {n2} строк уже есть в файле (по полю «{key}»):\n{sample}{more}", n=len(rep_rows), n2=len(data), key=key, sample=sample, more=more),
                                     [tr("Пропустить повторы"), tr("Сохранить все"), tr("Отмена")])
                        if ch == 2:
                            return
                        if ch == 0:
                            skip = set(rep_rows)
                            data = [row for i, row in enumerate(data) if i not in skip]
            if is_new:
                ws.append(cols)
                for c in ws[1]:
                    c.font = Font(bold=True)
                ws.freeze_panes = "A2"
            for row in data:
                ws.append([_excel_value(v, cols[j] in numeric) for j, v in enumerate(row)])
            if is_new:
                for j, name in enumerate(cols):
                    width = max([len(name)] + [len(str(r[j])) for r in data[:300]]) + 3
                    ws.column_dimensions[get_column_letter(j + 1)].width = max(9, min(45, width))
            book.save(path)
        except PermissionError:
            _error(self, tr("Файл занят:\n{path}\n\nЗакройте его в Excel и нажмите «В Excel» ещё раз.", path=path))
            return
        except Exception as e:
            _error(self, tr("Не удалось сохранить:\n{e}", e=e))
            return

        msg = tr("Сохранено строк: {n}\n{path}", n=len(data), path=path)
        nxt = self._advance_auto_numbers()
        if nxt:
            msg += tr("\n\nСледующая пачка начнётся с: ") + ", ".join(f"{k} = {v}" for k, v in nxt.items())
        self.status.showMessage(tr("Сохранено в Excel: {n} строк → {path}", n=len(data), path=path))
        _info(self, msg, tr("Сохранено"))

    def _advance_auto_numbers(self) -> dict:
        """После сохранения: номера строк закрепляются, начало следующей пачки сдвигается."""
        if self.template is None or not self.template.auto_number or not self.waybills:
            return {}
        out = {}
        for name in self.template.auto_number:
            w = self._manual_entries.get(name)
            if w is None:
                continue
            start = w.value()
            for i, wb in enumerate(self.waybills):
                wb.setdefault(name, str(start + i))
            new_start = start + len(self.waybills)
            w.blockSignals(True)
            w.setValue(new_start)
            w.blockSignals(False)
            self.template.auto_start[name] = new_start
            out[name] = new_start
        save_templates(self.templates, self.template.name)
        return out

    def _copy_clip(self):
        _cols, data = self._table_data()
        if not data:
            _warn(self, tr("Нет данных."))
            return
        QApplication.clipboard().setText("\n".join("\t".join(r) for r in data))
        self.status.showMessage(tr("Скопировано {n} строк → Ctrl+V", n=len(data)))
        _info(self, tr("Скопировано {n} строк.", n=len(data)))

    def _copy_report(self):
        """Копирует колонки отчёта (из шаблона) в буфер для Google Таблиц."""
        if self.template is None or self.tree.topLevelItemCount() == 0:
            _warn(self, tr("Нет данных."))
            return
        lines = list(self.template.report_lines)
        if not any(ln.strip() for ln in lines):
            _info(self, tr("В шаблоне не настроен отчёт.\nШаблоны → Изменить шаблон… → блок «7. Отчёт»."), tr("Отчёт"))
            return
        parsed = [parse_report_line(ln) for ln in lines]
        texts = [(i, p) for i, p in enumerate(parsed) if p[0] == "text"]
        if texts:
            dlg = QDialog(self)
            dlg.setWindowTitle(tr("Параметры отчёта"))
            dlg.setMinimumWidth(520)
            dlg.setStyleSheet(current_qss())
            v = QVBoxLayout(dlg)
            form = QFormLayout()
            form.setSpacing(8)
            edits = []
            for _i, (_k, label, val) in texts:
                ed = QLineEdit(val)
                ed.setMinimumWidth(340)
                form.addRow(QLabel(f"{label}:"), ed)
                edits.append(ed)
            v.addLayout(form)
            note = QLabel(tr("Значения применятся ко всем строкам и запомнятся в шаблоне."))
            note.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
            v.addWidget(note)
            bb = QDialogButtonBox(_DB_OK | _DB_CANCEL)
            bb.button(_DB_OK).setText(tr("Скопировать"))
            bb.accepted.connect(dlg.accept)
            bb.rejected.connect(dlg.reject)
            v.addWidget(bb)
            if dlg.exec() != _DLG_OK:
                return
            changed = False
            for (i, (_k, label, _old)), ed in zip(texts, edits):
                val = ed.text().strip()
                parsed[i] = ("text", label, val)
                line = f"{label} = {val}"
                if self.template.report_lines[i] != line:
                    self.template.report_lines[i] = line
                    changed = True
            if changed:
                save_templates(self.templates, self.template.name)

        hmap = {self.tree.headerItem().text(c): c for c in range(self.tree.columnCount())}
        rows = []
        for r in range(self.tree.topLevelItemCount()):
            it = self.tree.topLevelItem(r)
            cells = []
            for kind, a, b in parsed:
                if kind == "field":
                    c = hmap.get(a)
                    cells.append(self._cell_text(it, c) if c is not None else "")
                elif kind in ("text", "const"):
                    cells.append(b)
                else:
                    cells.append("")
            rows.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(rows))
        self.status.showMessage(tr("Отчёт: скопировано {n} строк → Ctrl+V в Google Таблицы", n=len(rows)))
        _info(self, tr("Отчёт скопирован: {n} строк.", n=len(rows)))


class _PdfBatchMixin:
    """Пакетные операции с PDF: сжатие и разбивка."""

    # ═════════════════════════════════════════════════════════════════════════
    # PDF-операции (сжатие / разбивка)
    # ═════════════════════════════════════════════════════════════════════════

    @staticmethod
    def _jpeg_page(src_doc, pg: int, dpi: int, quality: int, grayscale: bool = True):
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        cs  = fitz.csGRAY if grayscale else fitz.csRGB
        pix = src_doc[pg].get_pixmap(matrix=mat, colorspace=cs)
        return pix.tobytes("jpeg", jpg_quality=quality), pix.width, pix.height

    @staticmethod
    def _add_page(out_doc, img: bytes, w: int, h: int):
        p = out_doc.new_page(width=w, height=h)
        p.insert_image(fitz.Rect(0, 0, w, h), stream=img)

    def _ask_compress_cfg(self, title: str, **kwargs) -> Optional[dict]:
        dlg = CompressDialog(title, parent=self, **kwargs)
        return dlg.get_config() if dlg.exec() == _DLG_OK else None

    def _compress_all(self):
        if not self.pdf_path:
            _warn(self, tr("Сначала откройте PDF файл."))
            return
        cfg = self._ask_compress_cfg(tr("Сжатие всей пачки PDF"), with_limit=True)
        if cfg is None:
            return
        out_dir = QFileDialog.getExistingDirectory(self, tr("Папка для сохранения"))
        if not out_dir:
            return

        src_name  = (_safe_name(os.path.basename(self._src_folder.rstrip("/\\")))
                     if self._src_folder
                     else os.path.splitext(os.path.basename(self.pdf_path))[0])
        src_doc   = fitz.open(self.pdf_path)
        total     = len(src_doc)
        max_bytes = cfg["max_mb"] * 1024 ** 2 if cfg["max_mb"] > 0 else float("inf")
        saved: list = []
        part      = 1
        out_doc   = fitz.open()
        cur_size  = 0

        def _flush(doc, num):
            suf   = f"_{cfg['dpi']}dpi" + (tr("_часть{num}", num=num) if max_bytes < float("inf") else "")
            fname = os.path.join(out_dir, f"{src_name}{suf}.pdf")
            doc.save(fname, garbage=4, deflate=True, deflate_images=True, use_objstms=1)
            doc.close()
            saved.append((fname, os.path.getsize(fname)))
            return fitz.open()

        for idx in range(total):
            img, w, h = self._jpeg_page(src_doc, idx, cfg["dpi"], cfg["quality"], cfg.get("grayscale", True))
            self._add_page(out_doc, img, w, h)
            cur_size += len(img)
            QApplication.processEvents()
            if cur_size >= max_bytes and idx < total - 1:
                out_doc  = _flush(out_doc, part)
                part    += 1
                cur_size = 0
        if len(out_doc):
            _flush(out_doc, part)
        src_doc.close()

        details = "\n".join(tr("  {v}: {v2:.1f} МБ", v=os.path.basename(f), v2=sz/1024**2) for f, sz in saved)
        _info(self, tr("Готово! Файлов: {n}\nDPI: {dpi},  JPEG: {quality}%\n\n{details}\n\nПапка: {out_dir}", n=len(saved), dpi=cfg['dpi'], quality=cfg['quality'], details=details, out_dir=out_dir), tr("Сжатие завершено"))
        if _ask(self, tr("Открыть папку?"), tr("Открыть?")):
            _open_folder(out_dir)

    def _split_docs(self):
        if not self.pdf_path:
            _warn(self, tr("Сначала откройте PDF файл."))
            return
        if not self.tree.topLevelItemCount():
            _warn(self, tr("Таблица пустая."))
            return

        # Колонка для имён файлов: выбранное в шаблоне поле, иначе — первое
        # извлекаемое поле (поведение по умолчанию).
        name_col = None
        if self.template and self.template.extract_names:
            ex = self.template.extract_names
            nf = self.template.name_field if self.template.name_field in ex else ex[0]
            name_col = self.template.column_index().get(nf)
        names = []
        for i in range(self.tree.topLevelItemCount()):
            it = self.tree.topLevelItem(i)
            if it is None:
                continue
            names.append(self._cell_text(it, name_col) if name_col is not None else "")

        cfg = self._ask_compress_cfg(tr("Разбивка PDF на файлы"), with_split=True)
        if cfg is None:
            return
        out_dir = QFileDialog.getExistingDirectory(self, tr("Папка для сохранения"))
        if not out_dir:
            return

        src_doc  = fitz.open(self.pdf_path)
        total_pg = len(src_doc)
        ppd      = self.template.pages if self.template else 1
        expected = (total_pg + ppd - 1) // ppd
        if len(names) != expected:
            mode = tr("по {ppd} стр.", ppd=ppd) if ppd > 1 else tr("по 1 стр.")
            if not _ask(self, tr("Строк: {n}, ожидается: {expected} ({mode})\nПродолжить?", n=len(names), expected=expected, mode=mode), tr("Внимание")):
                src_doc.close()
                return

        name_cnt  = Counter(names)
        name_seen: dict = {}
        saved = errors = 0
        for ri, val in enumerate(names):
            pages = self._row_pages(ri)
            if not pages:
                continue
            base = val if val and val not in ("", "None") else tr("документ_{v}", v=ri + 1)
            if name_cnt[val] > 1:
                seen = name_seen.get(val, 0) + 1
                name_seen[val] = seen
                base = f"{base}_{seen}"
            out_path = os.path.join(out_dir, f"{_safe_name(base)}.pdf")
            try:
                nd = fitz.open()
                for pg in pages:
                    if cfg.get("do_compress", False):
                        img, w, h = self._jpeg_page(src_doc, pg, cfg["dpi"], cfg["quality"], cfg.get("grayscale", True))
                        self._add_page(nd, img, w, h)
                    else:
                        nd.insert_pdf(src_doc, from_page=pg, to_page=pg)
                nd.save(out_path, garbage=4, deflate=True, deflate_images=True, use_objstms=1)
                nd.close()
                saved += 1
            except Exception:
                errors += 1
            QApplication.processEvents()
        src_doc.close()

        nc  = tr(", сжатие {dpi} DPI", dpi=cfg['dpi']) if cfg.get("do_compress") else ""
        npg = tr("{ppd} стр/файл", ppd=ppd) if ppd > 1 else tr("1 стр/файл")
        msg = tr("Готово!\nФайлов: {saved} ({npg}{nc})\nПапка: {out_dir}", saved=saved, npg=npg, nc=nc, out_dir=out_dir)
        if errors:
            msg += tr("\nОшибок: {errors}", errors=errors)
        _info(self, msg, tr("Разбивка завершена"))
        if saved and _ask(self, tr("Открыть папку?"), tr("Открыть?")):
            _open_folder(out_dir)


# ─────────────────────────────────────────────────────────────────────────────
# Демо-пример: PDF с текстовым слоем + готовый шаблон
# ─────────────────────────────────────────────────────────────────────────────
# Демо создаётся на лету (файлы не нужны). В пачке 6 накладных; в трёх из них
# специально есть проблемы, чтобы показать подсветку:
#   • буква «О» вместо нуля в массе → исправляется обработкой OCR;
#   • нет массы → жёлтая ячейка «∅»;
#   • буква «О» в дате → красная ячейка (не прошла проверку);
#   • повтор номера документа → фиолетовая ячейка.

_DEMO_PT = 300 / 72  # точки PDF → координаты зон (300 DPI)

# Прямоугольники значений на странице A4, в точках PDF: (x0, y0, x1, y1)
_DEMO_RECTS = {
    "doc":    (200, 140, 470, 162),
    "date":   (200, 185, 470, 207),
    "supp":   (200, 230, 470, 252),
    "weight": (200, 275, 470, 297),
}


def _demo_docs() -> list:
    """(номер, дата, поставщик, масса) для 6 демо-документов."""
    if LANG == "ru":
        return [
            ("Накладная № ТН-0001", "от 03.03.2025 г.", "ООО «Ромашка»", "24,35"),
            ("Накладная № ТН-0002", "от 03.03.2025 г.", "ООО «Лютик»",   "2О,15"),
            ("Накладная № ТН-0003", "от 04.03.2025 г.", "АО «Василёк»",  ""),
            ("Накладная № ТН-0004", "от О4.03.2025 г.", "ООО «Ромашка»", "22,90"),
            ("Накладная № ТН-0003", "от 05.03.2025 г.", "ООО «Лютик»",   "25,10"),
            ("Накладная № ТН-0006", "от 05.03.2025 г.", "АО «Василёк»",  "23,80"),
        ]
    return [
        ("Delivery note No. DN-0001", "dated 03.03.2025", "Acme Ltd.",    "24.35"),
        ("Delivery note No. DN-0002", "dated 03.03.2025", "Globex Inc.",  "2O.15"),
        ("Delivery note No. DN-0003", "dated 04.03.2025", "Initech LLC",  ""),
        ("Delivery note No. DN-0004", "dated O4.03.2025", "Acme Ltd.",    "22.90"),
        ("Delivery note No. DN-0003", "dated 05.03.2025", "Globex Inc.",  "25.10"),
        ("Delivery note No. DN-0006", "dated 05.03.2025", "Initech LLC",  "23.80"),
    ]


def _demo_html(page, rect, html: str, size: int = 11, color: str = "#111827", bold=False):
    css = (f"* {{font-family: sans-serif; font-size: {size}px; color: {color};"
           f" font-weight: {'bold' if bold else 'normal'};}}")
    page.insert_htmlbox(fitz.Rect(*rect), html, css=css)


def make_demo_pdf(path: str) -> str:
    """Создаёт демо-PDF (6 одностраничных накладных) с текстовым слоем."""
    if not hasattr(fitz.Page, "insert_htmlbox"):
        raise RuntimeError(tr("Для демо нужен PyMuPDF 1.23 или новее:\npip install -U PyMuPDF"))
    goods = ([("Щебень фр. 20-40", "т"), ("Песок строительный", "т"), ("Доставка", "рейс")]
             if LANG == "ru" else
             [("Crushed stone 20-40", "t"), ("Building sand", "t"), ("Delivery", "trip")])
    labels = ([tr("Документ:"), tr("Дата:"), tr("Поставщик:"), tr("Масса нетто, т:")])
    doc = fitz.open()
    for i, (no, date, supp, weight) in enumerate(_demo_docs()):
        page = doc.new_page(width=595, height=842)
        _demo_html(page, (50, 50, 545, 80), tr("ТОВАРНАЯ НАКЛАДНАЯ"), size=18, bold=True)
        _demo_html(page, (50, 82, 545, 100),
                   tr("Демо-документ ZoneExtract. Все данные вымышленные."), size=9, color="#6b7280")
        for label, key in zip(labels, ("doc", "date", "supp", "weight")):
            x0, y0, x1, y1 = _DEMO_RECTS[key]
            _demo_html(page, (50, y0 + 3, 195, y1), label, color="#374151")
            page.draw_rect(fitz.Rect(x0 - 4, y0 - 2, x1, y1 + 2), color=(0.8, 0.83, 0.87), width=0.8)
            value = {"doc": no, "date": date, "supp": supp, "weight": weight}[key]
            if value:
                _demo_html(page, (x0, y0 + 3, x1 - 4, y1), value, size=12)
        # Таблица товаров — для правдоподобия, программа её не читает
        y = 345
        page.draw_line(fitz.Point(50, y - 4), fitz.Point(545, y - 4), color=(0.4, 0.4, 0.4), width=0.8)
        for k, (name, unit) in enumerate(goods):
            _demo_html(page, (50, y, 80, y + 18), str(k + 1), size=10)
            _demo_html(page, (85, y, 380, y + 18), name, size=10)
            _demo_html(page, (390, y, 450, y + 18), unit, size=10)
            _demo_html(page, (460, y, 545, y + 18), str(1 + (i + k) % 3), size=10)
            y += 22
        page.draw_line(fitz.Point(50, y), fitz.Point(545, y), color=(0.4, 0.4, 0.4), width=0.8)
        _demo_html(page, (50, 700, 290, 720), tr("Отпустил: ____________"), size=10)
        _demo_html(page, (310, 700, 545, 720), tr("Принял: ____________"), size=10)
        _demo_html(page, (50, 800, 545, 815), tr("Лист {n} из {total}", n=i + 1, total=6),
                   size=8, color="#9ca3af")
    doc.save(path, garbage=1, deflate=True)
    doc.close()
    return path


def _demo_zone(key: str, pad: float = 3) -> tuple:
    x0, y0, x1, y1 = _DEMO_RECTS[key]
    return tuple(int(round(v * _DEMO_PT)) for v in (x0 - pad, y0 - pad, x1 + pad, y1 + pad))


def demo_template() -> Template:
    """Готовый шаблон для демо-PDF."""
    f_no, f_date, f_supp, f_w = tr("№ документа"), tr("Дата"), tr("Поставщик"), tr("Масса, т")
    f_auto, f_mat, f_vol = tr("№ п/п"), tr("Материал"), tr("Объём, м³")
    docs = _demo_docs()

    def fld(name, kind, example):
        return dict(name=name, page=0, example=example, **kind_state(kind))

    return Template(
        name=tr("Демо: накладные"),
        pages=1,
        manual_fields=[f_auto, f_mat],
        extract_fields=[
            fld(f_no, "doc_code", docs[0][0]),
            fld(f_date, "date_full", docs[0][1]),
            fld(f_supp, "text", docs[0][2]),
            fld(f_w, "number", docs[1][3]),
        ],
        zones={f_no: _demo_zone("doc"), f_date: _demo_zone("date"),
               f_supp: _demo_zone("supp"), f_w: _demo_zone("weight")},
        name_field=f_no,
        column_order=[f_auto, f_no, f_date, f_supp, f_mat, f_w, f_vol],
        computed_fields=[{"name": f_vol, "formula": f"[{f_w}] / 1{DEC_SEP}6", "decimals": 2}],
        report_lines=[f"[{f_date}]", f"[{f_no}]", f"[{f_supp}]",
                      tr("Склад") + " = " + tr("Основной"), f"[{f_w}]", f"[{f_vol}]"],
        auto_number=[f_auto],
        auto_start={f_auto: 1},
    )


# ─────────────────────────────────────────────────────────────────────────────
# Мастер «Первый шаблон»
# ─────────────────────────────────────────────────────────────────────────────

class FirstTemplateWizard(QDialog):
    """Три шага: образец PDF → список полей → что делать дальше."""

    def __init__(self, existing_names=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Мастер: первый шаблон"))
        self.resize(720, 560)
        self.setStyleSheet(current_qss())
        self._existing = set(existing_names or [])
        self._pdf_path = ""
        self._pages_in_pdf = 0
        self._result: Optional[Template] = None

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        hdr = QFrame()
        hdr.setStyleSheet(f"background:{C_ACCENT};")
        h = QHBoxLayout(hdr)
        h.setContentsMargins(18, 10, 18, 10)
        self._step_lbl = QLabel()
        self._step_lbl.setStyleSheet("color:white; font-weight:bold; font-size:12pt; background:transparent;")
        h.addWidget(self._step_lbl)
        root.addWidget(hdr)

        from PySide6.QtWidgets import QStackedWidget
        self._stack = QStackedWidget()
        root.addWidget(self._stack, 1)
        self._stack.addWidget(self._page_sample())
        self._stack.addWidget(self._page_fields())
        self._stack.addWidget(self._page_next())

        bar = QHBoxLayout()
        bar.setContentsMargins(18, 8, 18, 14)
        self._btn_back = QPushButton(tr("← Назад"))
        self._btn_back.clicked.connect(lambda: self._go(-1))
        bar.addWidget(self._btn_back)
        bar.addStretch()
        btn_cancel = QPushButton(tr("Отмена"))
        btn_cancel.clicked.connect(self.reject)
        bar.addWidget(btn_cancel)
        self._btn_next = QPushButton()
        self._btn_next.setObjectName("primary")
        self._btn_next.clicked.connect(lambda: self._go(1))
        bar.addWidget(self._btn_next)
        root.addLayout(bar)
        self._update_buttons()

    # ── Страницы ─────────────────────────────────────────────────────────────
    @staticmethod
    def _text(html: str) -> QLabel:
        lbl = QLabel(html)
        lbl.setWordWrap(True)
        lbl.setTextFormat(_qenum(Qt, "TextFormat", "RichText"))
        return lbl

    def _page_sample(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(22, 18, 22, 10)
        lay.setSpacing(12)
        lay.addWidget(self._text(tr(
            "Шаблон — это настройка под <b>один вид бланка</b>: где на странице лежит каждое "
            "значение и как его проверить. Настраивается один раз, потом пачки "
            "обрабатываются в пару кликов.<br><br>"
            "Возьмите PDF с несколькими такими бланками. PDF должен быть "
            "<b>с текстовым слоем</b> (после распознавания, например в FineReader: "
            "«PDF с возможностью поиска»).")))
        row = QHBoxLayout()
        btn = QPushButton(tr("📄 Выбрать PDF-образец…"))
        btn.setObjectName("primary")
        btn.clicked.connect(self._pick_pdf)
        row.addWidget(btn)
        self._pdf_lbl = QLabel(tr("не выбран (можно выбрать позже)"))
        self._pdf_lbl.setStyleSheet(f"color:{C_TEXT2};")
        row.addWidget(self._pdf_lbl, 1)
        lay.addLayout(row)
        form = QFormLayout()
        form.setSpacing(10)
        self._name = QLineEdit()
        self._name.setPlaceholderText(tr("Например: Товарные накладные"))
        form.addRow(tr("Название шаблона:"), self._name)
        self._pages = NoWheelSpinBox()
        self._pages.setRange(1, MAX_DOC_PAGES)
        self._pages.setFixedWidth(80)
        prow = QHBoxLayout()
        prow.addWidget(self._pages)
        prow.addWidget(self._text(tr("Обычно 1. Если бланк двусторонний и сканируется "
                                     "с обеих сторон — 2.")), 1)
        form.addRow(tr("Страниц на один документ:"), prow)
        lay.addLayout(form)
        lay.addStretch()
        return w

    def _page_fields(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(22, 18, 22, 10)
        lay.setSpacing(8)
        lay.addWidget(self._text(tr(
            "<b>Какие значения взять с бланка?</b> По одному на строку. Это будут столбцы "
            "таблицы. Каждому полю на следующем шаге вы поставите рамку на странице.")))
        self._extract = QPlainTextEdit()
        self._extract.setPlaceholderText(tr("№ документа\nДата\nПоставщик\nМасса, т"))
        self._extract.setFixedHeight(150)
        lay.addWidget(self._extract)
        lay.addWidget(self._text(tr(
            "<b>Общие значения для всей пачки</b> (необязательно) — их нет на бланке, "
            "вы впишете их один раз: например, материал или объект.")))
        self._manual = QPlainTextEdit()
        self._manual.setPlaceholderText(tr("Материал"))
        self._manual.setFixedHeight(80)
        lay.addWidget(self._manual)
        self._auto = QCheckBox(tr("Добавить столбец «№ п/п» — номер строки по порядку"))
        self._auto.setChecked(True)
        lay.addWidget(self._auto)
        lay.addStretch()
        return w

    def _page_next(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(22, 18, 22, 10)
        lay.addWidget(self._text(tr(
            "<b>Шаблон почти готов. Осталось три действия:</b><br><br>"
            "<b>1. Расставить рамки.</b> Откроется вкладка «3. Редактор зон». На странице — "
            "цветные рамки, по одной на поле. Перетащите каждую на своё значение. Внизу "
            "виден текст, который программа прочитала из рамки.<br><br>"
            "<b>2. Настроить поля.</b> Нажмите «🔮 Дальше: настроить поля…», затем "
            "«🔮 Подобрать вид для всех полей». Программа сама поймёт, где число, где дата, "
            "где номер. Сохраните шаблон.<br><br>"
            "<b>3. Извлечь.</b> Вкладка «2. Распознавание» → «▶ Извлечь данные». "
            "Результат — на вкладке «5. Таблица».<br><br>"
            "Подробно — «Справка → Инструкция».")))
        lay.addStretch()
        return w

    # ── Логика ───────────────────────────────────────────────────────────────
    def _pick_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, tr("PDF-образец"), "", tr("PDF (*.pdf);;Все (*.*)"))
        if not p:
            return
        try:
            d = fitz.open(p)
            self._pages_in_pdf = len(d)
            d.close()
        except Exception as e:
            _error(self, tr("Не удалось открыть PDF:\n{e}", e=e))
            return
        self._pdf_path = p
        self._pdf_lbl.setText(tr("{name}  ({n} стр.)", name=os.path.basename(p), n=self._pages_in_pdf))
        self._pdf_lbl.setStyleSheet(f"color:{C_SUCCESS};")
        if not self._name.text().strip():
            self._name.setText(os.path.splitext(os.path.basename(p))[0][:40])

    @staticmethod
    def _lines(edit: QPlainTextEdit) -> list:
        return [ln.strip() for ln in edit.toPlainText().splitlines() if ln.strip()]

    def _validate(self, step: int) -> bool:
        if step == 0:
            name = self._name.text().strip()
            if not name:
                _warn(self, tr("Укажите название шаблона."), tr("Проверка"))
                return False
            if name in self._existing:
                _warn(self, tr("Шаблон «{name}» уже существует.", name=name), tr("Проверка"))
                return False
        if step == 1:
            extract = self._lines(self._extract)
            if not extract:
                _warn(self, tr("Укажите хотя бы одно поле, которое нужно взять с бланка."),
                      tr("Проверка"))
                return False
            names = extract + self._lines(self._manual) + ([tr("№ п/п")] if self._auto.isChecked() else [])
            dup = [n for n in dict.fromkeys(names) if names.count(n) > 1]
            if dup:
                _warn(self, tr("Названия полей повторяются: ") + ", ".join(dup), tr("Проверка"))
                return False
            bad = [n for n in names if "[" in n or "]" in n]
            if bad:
                _warn(self, tr("В названиях полей нельзя использовать [ и ]: ") + ", ".join(bad),
                      tr("Проверка"))
                return False
        return True

    def _go(self, d: int):
        i = self._stack.currentIndex()
        if d > 0 and not self._validate(i):
            return
        if d > 0 and i == self._stack.count() - 1:
            self._finish()
            return
        self._stack.setCurrentIndex(max(0, min(self._stack.count() - 1, i + d)))
        self._update_buttons()

    def _update_buttons(self):
        i = self._stack.currentIndex()
        titles = [tr("Шаг 1 из 3 — образец документа"), tr("Шаг 2 из 3 — какие данные нужны"),
                  tr("Шаг 3 из 3 — что дальше")]
        self._step_lbl.setText("🧙  " + titles[i])
        self._btn_back.setEnabled(i > 0)
        self._btn_next.setText(tr("Создать шаблон") if i == self._stack.count() - 1 else tr("Далее →"))

    def _finish(self):
        extract = self._lines(self._extract)
        manual = self._lines(self._manual)
        auto = [tr("№ п/п")] if self._auto.isChecked() else []
        manual = auto + manual
        self._result = Template(
            name=self._name.text().strip(),
            pages=self._pages.value(),
            manual_fields=manual,
            extract_fields=[dict(name=n, page=0, **kind_state("text")) for n in extract],
            name_field=extract[0],
            auto_number=auto,
            auto_start={n: 1 for n in auto},
        )
        self.accept()

    def result_template(self) -> Optional[Template]:
        return self._result

    def pdf_path(self) -> str:
        return self._pdf_path


# ─────────────────────────────────────────────────────────────────────────────
# Справка
# ─────────────────────────────────────────────────────────────────────────────

HOMEPAGE = "https://github.com/demyanenkopavel-wq/ZoneExtract"


def _help_css() -> str:
    return f"""
<style>
 body {{ font-family:'Segoe UI',sans-serif; font-size:10pt; color:{T.text}; }}
 h1 {{ font-size:15pt; color:{T.accent_bright}; margin:4px 0 8px 0; }}
 h2 {{ font-size:12pt; color:{T.accent_bright}; margin:18px 0 6px 0; }}
 pre {{ background:{T.chrome_deep}; color:{T.on_chrome}; padding:10px;
        font-family:Consolas,'Courier New',monospace; font-size:9.5pt; }}
 code {{ font-family:Consolas,'Courier New',monospace; background:{T.surface_alt}; }}
 table {{ border-collapse:collapse; }}
 td, th {{ border:1px solid {T.content_border}; padding:5px 8px; vertical-align:top; }}
 th {{ background:{T.chrome}; color:{T.on_chrome}; text-align:left; }}
 .k {{ font-family:Consolas,monospace; background:{T.surface_alt}; padding:1px 5px; }}
</style>"""


def _kind_rows_html() -> str:
    from html import escape as esc
    rows = ""
    for key, k in KIND_MAP.items():
        example = tr(k["example"])
        res, st = rule_result(example, kind_state(key))
        color = {"ok": T.success, "err": T.danger}.get(st, T.text)
        rows += (f"<tr><td><b>{esc(tr(k['label']))}</b></td>"
                 f"<td><code>{esc(example)}</code></td>"
                 f"<td style='color:{color}'><code>{esc(res)}</code></td>"
                 f"<td>{esc(tr(k['hint']))}</td></tr>")
    return rows


def _colors_rows_html(lang: str) -> str:
    m = missing_mark()
    if lang == "ru":
        rows = [(T.empty_bg, T.empty_fg, m, "В рамке нет текста.",
                 "Посмотрите скан справа и впишите значение. Пусто во всех строках — сдвиньте рамку."),
                (T.err_bg, T.err_fg, "красный", "Значение есть, но формат не тот.",
                 "Обычно ошибка распознавания. Сверьте со сканом и исправьте."),
                (T.dup_bg, T.dup_fg, "фиолетовый", "Номер документа повторяется в пачке.",
                 "Обычно лист отсканирован дважды. Удалите лишнюю строку.")]
        head = "<tr><th>Цвет</th><th>Что значит</th><th>Что делать</th></tr>"
    else:
        rows = [(T.empty_bg, T.empty_fg, m, "The frame has no text.",
                 "Look at the scan on the right and type the value. Empty in all rows — move the frame."),
                (T.err_bg, T.err_fg, "red", "There is a value, but the format is wrong.",
                 "Usually an OCR error. Compare with the scan and fix it."),
                (T.dup_bg, T.dup_fg, "purple", "The document number repeats in the batch.",
                 "Usually a page was scanned twice. Delete the extra row.")]
        head = "<tr><th>Color</th><th>Meaning</th><th>What to do</th></tr>"
    return head + "".join(
        f"<tr><td style='background:{bg}; color:{fg}'><b>{name}</b></td><td>{a}</td><td>{b}</td></tr>"
        for bg, fg, name, a, b in rows)


def help_html() -> str:
    """Инструкция на текущем языке."""
    lang = LANG
    if lang == "ru":
        body = f"""
<h1>ZoneExtract — инструкция</h1>
<p>Программа переносит данные с одинаковых бланков (накладные, акты, путевые листы…) в таблицу.
Она читает <b>текстовый слой PDF</b> внутри прямоугольных <b>рамок</b>, которые вы ставите
на странице один раз.</p>
<pre>
 Скан ──► Распознавание (OCR) ──► PDF с текстом ──► ZoneExtract ──► Таблица ──► Excel / буфер
           FineReader и т.п.                        рамка → поле    проверка
</pre>

<h2>Самый быстрый способ понять — демо</h2>
<p><b>Справка → Демо-пример</b>. Программа создаст тестовый PDF и готовый шаблон.
Нажмите «▶ Извлечь данные» — и посмотрите, как подсвечиваются ошибки.</p>

<h2>Первый шаблон — 5 шагов</h2>
<ol>
 <li><b>Шаблоны → Мастер: первый шаблон.</b> Выберите PDF-образец, впишите названия нужных
     значений (по одному на строку).</li>
 <li><b>Вкладка «3. Редактор зон».</b> Перетащите каждую цветную рамку на своё значение.
     Внизу — текст, который видит программа. Рамка с запасом — это нормально.</li>
 <li><b>«🔮 Дальше: настроить поля…»</b> → <b>«🔮 Подобрать вид для всех полей»</b>.
     Программа возьмёт текст из рамок и сама поймёт: число, дата, номер… Справа ✓ — всё
     в порядке. Нажмите «💾 Сохранить шаблон».</li>
 <li><b>Вкладка «2. Распознавание»</b> → «▶ Извлечь данные».</li>
 <li><b>Вкладка «5. Таблица».</b> Исправьте подсвеченные ячейки и сохраните результат.</li>
</ol>
<p>Дальше для новых пачек того же бланка: выбрать шаблон → открыть PDF → «Извлечь».</p>

<h2>Правка ошибок</h2>
<table>{_colors_rows_html(lang)}</table>
<p><span class=k>F3</span> — следующая проблема. Справа — кусок скана из рамки: впишите
правильное значение и нажмите <span class=k>Enter</span>, программа перейдёт к следующей.
Внизу таблицы — число строк и суммы числовых столбцов.</p>

<h2>Виды полей</h2>
<p>Вид поля решает, что взять из текста рамки и как проверить результат.
Выберите вид в колонке «Что это за поле» или нажмите 🔮.</p>
<table><tr><th>Вид</th><th>Текст в рамке</th><th>Результат</th><th>Что делает</th></tr>
{_kind_rows_html()}</table>

<h2>Другие поля шаблона</h2>
<ul>
 <li><b>Вручную</b> — одно значение на всю пачку (материал, объект). Вписывается на вкладке
     «1. Общие данные». Галочка «№ по порядку» нумерует строки; после сохранения в Excel
     нумерация продолжается.</li>
 <li><b>Вычисляемые</b> — формула из других полей: <code>[Масса, т] / 1,6</code>.
     Можно + − * / и скобки.</li>
 <li><b>Ключевое поле</b> — обычно № документа. По нему ищутся повторы и называются файлы
     при «Разбить на файлы».</li>
 <li><b>Резерв</b> — вторая рамка. Читается, если в первой пусто (бланки чуть разные).</li>
 <li><b>Словарь</b> — список правильных значений (txt, по одному на строку). Распознанное
     значение заменяется на самое похожее из списка.</li>
</ul>

<h2>Результат</h2>
<ul>
 <li><b>💾 В Excel</b> — строки дописываются в конец файла, старые данные не меняются.
     Числа сохраняются числами. Повторы номеров программа предложит пропустить.</li>
 <li><b>📋 В буфер</b> — вся таблица, вставка в Excel или Google Таблицы через Ctrl+V.</li>
 <li><b>📊 Отчёт</b> — свои колонки (блок «7. Отчёт» в шаблоне): <code>[Поле]</code> — значение,
     <code>Склад = Основной</code> — постоянный текст, <code>-</code> — пустая колонка.</li>
 <li><b>✂️ Разбить на файлы</b> — каждый документ пачки в свой PDF, имя — по ключевому полю.</li>
</ul>

<h2>Свои правила (⚙)</h2>
<p>Нужны редко. Текст рамки проходит три шага:</p>
<pre>
 текст рамки ──► 1. Обработка ──► 2. Очистка (regex) ──► 3. Проверка ──► ячейка
                (исправления OCR)   (что вырезать)        (не прошло = красный)
</pre>
<p>Пример: очистка <code>0*(\\d+)</code> из «ОРДЕР № 00000012345» даёт «12345».</p>

<h2>Если что-то не так</h2>
<table><tr><th>Проблема</th><th>Что делать</th></tr>
 <tr><td>«Нет текстового слоя»</td><td>Распознайте скан (FineReader, Acrobat, OCRmyPDF) и откройте снова.</td></tr>
 <tr><td>Поле пустое во всех строках</td><td>Рамка не на месте. Вкладка 3: сдвиньте её на текст.</td></tr>
 <tr><td>Строки «съехали» относительно документов</td><td>Проверьте «Страниц на документ» и нет ли в PDF лишних страниц.</td></tr>
 <tr><td>Много красных ячеек</td><td>Плохой скан (нужно 300 DPI) или не тот вид поля.</td></tr>
 <tr><td>Excel не сохраняется</td><td>Файл открыт в Excel — закройте его.</td></tr>
</table>

<h2>Клавиши</h2>
<table>
 <tr><td class=k>F3</td><td>Следующая проблема</td></tr>
 <tr><td class=k>Enter</td><td>В панели скана: сохранить и перейти дальше</td></tr>
 <tr><td class=k>Двойной клик</td><td>Изменить ячейку</td></tr>
 <tr><td class=k>Ctrl+O / Ctrl+Shift+O</td><td>Открыть PDF / папку с PDF</td></tr>
</table>
<p>Шаблоны хранятся в файле <code>templates_config.json</code> рядом с программой —
его можно скопировать на другой компьютер.</p>
"""
    else:
        body = f"""
<h1>ZoneExtract — user guide</h1>
<p>ZoneExtract moves data from same-layout forms (delivery notes, invoices, waybills…) into a
table. It reads the <b>PDF text layer</b> inside rectangular <b>frames</b> that you place on the
page once.</p>
<pre>
 Scan ──► OCR ──────────────► PDF with text ──► ZoneExtract ──► Table ──► Excel / clipboard
          FineReader, etc.                       frame → field   checks
</pre>

<h2>Fastest way to learn — the demo</h2>
<p><b>Help → Demo</b>. The program creates a test PDF and a ready template.
Click “▶ Extract data” and see how problems are highlighted.</p>

<h2>Your first template — 5 steps</h2>
<ol>
 <li><b>Templates → First template wizard.</b> Choose a sample PDF, type the names of the
     values you need (one per line).</li>
 <li><b>Tab “3. Zone editor”.</b> Drag each colored frame onto its value. The bottom line
     shows the text the program sees. A frame with some margin is fine.</li>
 <li><b>“🔮 Next: set up fields…”</b> → <b>“🔮 Detect type for all fields”</b>. The program
     takes the text from the frames and detects: number, date, document number… A ✓ on the
     right means it works. Click “💾 Save template”.</li>
 <li><b>Tab “2. Recognition”</b> → “▶ Extract data”.</li>
 <li><b>Tab “5. Table”.</b> Fix the highlighted cells and save the result.</li>
</ol>
<p>For new batches of the same form: choose the template → open the PDF → “Extract”.</p>

<h2>Fixing problems</h2>
<table>{_colors_rows_html(lang)}</table>
<p><span class=k>F3</span> — next problem. The right panel shows the scan inside the frame:
type the correct value and press <span class=k>Enter</span> to go to the next problem.
The bar under the table shows the row count and sums of numeric columns.</p>

<h2>Field types</h2>
<p>The field type decides what to take from the frame text and how to check the result.
Choose it in the “What is this field” column or click 🔮.</p>
<table><tr><th>Type</th><th>Text in frame</th><th>Result</th><th>What it does</th></tr>
{_kind_rows_html()}</table>

<h2>Other template fields</h2>
<ul>
 <li><b>Manual</b> — one value for the whole batch (material, site). Type it on tab
     “1. Common data”. The “Row number” checkbox numbers the rows; after saving to Excel
     the numbering continues.</li>
 <li><b>Calculated</b> — a formula from other fields: <code>[Weight, t] / 1.6</code>.
     You can use + − * / and brackets.</li>
 <li><b>Key field</b> — usually the document number. Used to find duplicates and to name
     files in “Split into files”.</li>
 <li><b>Fallback</b> — a second frame. Read when the first one is empty (forms vary slightly).</li>
 <li><b>Dictionary</b> — a list of correct values (txt, one per line). The recognized value is
     replaced with the most similar one from the list.</li>
</ul>

<h2>Output</h2>
<ul>
 <li><b>💾 To Excel</b> — rows are appended to the end of the file; old data is not changed.
     Numbers are saved as numbers. The program offers to skip duplicate numbers.</li>
 <li><b>📋 To clipboard</b> — the whole table; paste into Excel or Google Sheets with Ctrl+V.</li>
 <li><b>📊 Report</b> — your own columns (block “7. Report” in the template): <code>[Field]</code> —
     value, <code>Warehouse = Main</code> — fixed text, <code>-</code> — empty column.</li>
 <li><b>✂️ Split into files</b> — each document of the batch into its own PDF, named by the key field.</li>
</ul>

<h2>Custom rules (⚙)</h2>
<p>Rarely needed. The frame text goes through three steps:</p>
<pre>
 frame text ──► 1. Processing ──► 2. Cleaning (regex) ──► 3. Check ──► cell
               (OCR fixes)        (what to cut out)       (fails = red)
</pre>
<p>Example: cleaning <code>0*(\\d+)</code> on “ORDER No. 00000012345” gives “12345”.</p>

<h2>Troubleshooting</h2>
<table><tr><th>Problem</th><th>What to do</th></tr>
 <tr><td>“No text layer”</td><td>Run OCR on the scan (FineReader, Acrobat, OCRmyPDF) and open it again.</td></tr>
 <tr><td>A field is empty in all rows</td><td>The frame is in the wrong place. Tab 3: move it onto the text.</td></tr>
 <tr><td>Rows are shifted against documents</td><td>Check “Pages per document” and extra pages in the PDF.</td></tr>
 <tr><td>Many red cells</td><td>Poor scan (use 300 DPI) or wrong field type.</td></tr>
 <tr><td>Excel file is not saved</td><td>The file is open in Excel — close it.</td></tr>
</table>

<h2>Keys</h2>
<table>
 <tr><td class=k>F3</td><td>Next problem</td></tr>
 <tr><td class=k>Enter</td><td>In the scan panel: save and go on</td></tr>
 <tr><td class=k>Double click</td><td>Edit a cell</td></tr>
 <tr><td class=k>Ctrl+O / Ctrl+Shift+O</td><td>Open PDF / folder with PDFs</td></tr>
</table>
<p>Templates are stored in <code>templates_config.json</code> next to the program —
you can copy it to another computer.</p>
"""
    return _help_css() + body


class _OnboardingMixin:
    """Демо, мастер, окно приветствия, язык, справка."""

    # ── Демо ────────────────────────────────────────────────────────────────
    def _open_demo(self):
        path = os.path.join(tempfile.gettempdir(), f"ZoneExtract_demo_{LANG}.pdf")
        try:
            make_demo_pdf(path)
        except Exception as e:
            _error(self, tr("Не удалось создать демо-PDF:\n{e}", e=e))
            return
        tpl = demo_template()
        self.templates = [t for t in self.templates if t.name != tpl.name] + [tpl]
        save_templates(self.templates, tpl.name)
        self._apply_template(tpl)
        self._refresh_template_combo()
        mat = self._manual_entries.get(tr("Материал"))
        if isinstance(mat, QLineEdit):
            mat.setText(tr("Щебень"))
        if not self._load_pdf_path(path, check_text=False):
            return
        self.tabs.setCurrentIndex(2)
        _info(self, tr(
            "Демо загружено: 6 накладных и шаблон «{name}».\n\n"
            "Сейчас открыт «Редактор зон»: цветные рамки показывают, откуда берутся значения.\n\n"
            "Дальше: вкладка «2. Распознавание» → «▶ Извлечь данные».\n"
            "В таблице будут три проблемы — пустая ячейка, ошибка формата и повтор номера. "
            "Нажимайте F3 и исправляйте их по скану справа.", name=tpl.name), tr("Демо-пример"))

    # ── Мастер ──────────────────────────────────────────────────────────────
    def _run_wizard(self):
        dlg = FirstTemplateWizard([t.name for t in self.templates], self)
        if dlg.exec() != _DLG_OK:
            return
        tpl = dlg.result_template()
        if tpl is None:
            return
        self.templates.append(tpl)
        save_templates(self.templates, tpl.name)
        self._apply_template(tpl)
        self._refresh_template_combo()
        if dlg.pdf_path():
            self._load_pdf_path(dlg.pdf_path())
        self.tabs.setCurrentIndex(2)
        msg = tr("Шаблон «{name}» создан.", name=tpl.name) + "\n\n"
        if self.parser is None:
            msg += tr("Откройте PDF-образец: вкладка «2. Распознавание» → «Выбрать файл…».") + "\n\n"
        msg += tr("Теперь перетащите каждую цветную рамку на своё значение. "
                  "Затем нажмите «🔮 Дальше: настроить поля…» вверху.")
        _info(self, msg, tr("Мастер: первый шаблон"))

    def _maybe_welcome(self):
        if self.templates:
            return
        ch = _choose(self, tr("Добро пожаловать"), tr(
            "ZoneExtract переносит данные из PDF-документов в таблицу.\n\n"
            "Вы один раз показываете, где на бланке лежит каждое значение, "
            "дальше программа сама обрабатывает пачки таких бланков.\n\n"
            "С чего начать?"),
            [tr("▶ Демо-пример"), tr("🧙 Мастер: первый шаблон"), tr("Позже")])
        if ch == 0:
            self._open_demo()
        elif ch == 1:
            self._run_wizard()

    # ── Язык ────────────────────────────────────────────────────────────────
    def _choose_language(self, lang: str):
        for code, act in self._lang_actions.items():
            act.setChecked(code == LANG)
        if lang == LANG:
            return
        settings = load_ui_settings()
        settings["lang"] = lang
        save_ui_settings(settings)
        # Сообщение — на выбранном языке, чтобы его понял тот, кто язык выбрал.
        text = ("Язык изменится после перезапуска. Перезапустить сейчас?" if lang == "ru"
                else "The language will change after restart. Restart now?")
        title = "Язык" if lang == "ru" else "Language"
        if QMessageBox.question(self, title, text, _MB_YES | _MB_NO, _MB_YES) == _MB_YES:
            self._restart()

    def _restart(self):
        if getattr(sys, "frozen", False):
            args = [sys.executable] + sys.argv[1:]
        else:
            args = [sys.executable, os.path.abspath(sys.argv[0])] + sys.argv[1:]
        try:
            subprocess.Popen(args)
        except Exception as e:
            _error(self, str(e))
            return
        self.close()
        QApplication.instance().quit()

    # ── Справка ─────────────────────────────────────────────────────────────
    def _show_help(self):
        dlg = QDialog(self)
        dlg.setWindowTitle(tr("Инструкция"))
        dlg.resize(980, 760)
        dlg.setStyleSheet(current_qss())
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(0, 0, 0, 12)
        tb = QTextBrowser()
        tb.setOpenExternalLinks(True)
        tb.setStyleSheet(f"QTextBrowser {{ background:{T.surface}; border:none; padding:12px; }}")
        tb.setHtml(help_html())
        lay.addWidget(tb)
        close_btn = QPushButton(tr("Закрыть"))
        close_btn.setObjectName("primary")
        close_btn.setFixedWidth(140)
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn, alignment=_ALIGN_CENTER)
        dlg.exec()

    def _show_about(self):
        box = QMessageBox(self)
        box.setWindowTitle(tr("О программе"))
        box.setIconPixmap(app_icon_pixmap(64))
        box.setTextFormat(_qenum(Qt, "TextFormat", "RichText"))
        box.setText(tr(
            "<h3>ZoneExtract {version}</h3>"
            "<p>Извлечение данных из распознанных PDF по зонам и шаблонам.</p>"
            "<p>Открытый исходный код, лицензия MIT.<br>"
            "<a href='{url}'>{url}</a></p>"
            "<p>Зависимости: PySide6, PyMuPDF, pandas, openpyxl.</p>",
            version=VERSION, url=HOMEPAGE))
        box.exec()


class MainWindow(_ChromeStyleMixin, _TemplateDictMixin, _ExtractMixin, _TableMixin, _PdfBatchMixin,
                 _OnboardingMixin, QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ZoneExtract")
        self.setWindowIcon(QIcon(app_icon_pixmap(64)))
        self.resize(1440, 860)
        self.setMinimumSize(1100, 680)

        # Состояние
        self.templates, selected = load_templates()
        self.template: Optional[Template] = None
        self.pdf_path: Optional[str]      = None
        self.parser: Optional[PdfParser]  = None
        self.waybills: list[dict]         = []
        self.excel_path: Optional[str]    = None
        self._dictionaries: dict[str, list] = {}
        self._dict_paths: dict[str, str]    = {}
        self._dict_btns: dict[str, QPushButton] = {}
        self._dict_lbls: dict[str, QLabel]      = {}
        self._match_btns: dict[str, QPushButton] = {}
        self._manual_entries: dict[str, QLineEdit] = {}
        self._worker: Optional[ExtractionWorker] = None
        # Источник «папка»: склеенный временный PDF и откуда каждая страница
        self._src_folder: Optional[str] = None
        self._src_recursive = False
        self._src_files: list = []
        self._merged_ppd = 1
        self._merged_temp: Optional[str] = None
        self._page_src: list = []
        self._thread: Optional[QThread] = None

        self._setup_ui()
        self._setup_menu()

        # Выбрать сохранённый шаблон (или первый доступный)
        start = None
        for t in self.templates:
            if t.name == selected:
                start = t
                break
        if start is None and self.templates:
            start = self.templates[0]
        self._refresh_template_combo()
        if start is not None:
            self._select_template_by_name(start.name)
        else:
            self._apply_template(None)

        register_retheme(self._retheme)
        QTimer.singleShot(400, self._maybe_welcome)

    def _setup_ui(self):
        self._header = header = QWidget()
        header.setFixedHeight(58)
        header.setStyleSheet(self._header_bg_ss())
        h_lay = QHBoxLayout(header)
        h_lay.setContentsMargins(12, 0, 20, 0)
        h_lay.setSpacing(10)

        _logo_lbl = QLabel()
        _logo_lbl.setPixmap(app_icon_pixmap(44))
        _logo_lbl.setStyleSheet("background:transparent;")
        h_lay.addWidget(_logo_lbl)

        self._title_lbl = title = QLabel("ZONEEXTRACT")
        title.setStyleSheet("color:white; font-size:14pt; font-weight:900; letter-spacing:1px; background:transparent;")
        h_lay.addWidget(title)
        h_lay.addStretch()

        self._tmpl_lbl = QLabel(tr("Шаблон:"))
        self._tmpl_lbl.setStyleSheet(self._header_dim_ss())
        h_lay.addWidget(self._tmpl_lbl)

        self._tmpl_combo = NoWheelComboBox()
        self._tmpl_combo.setFixedWidth(180)
        self._tmpl_combo.setStyleSheet(self._header_combo_ss())
        self._tmpl_combo.currentIndexChanged.connect(self._on_template_selected)
        h_lay.addWidget(self._tmpl_combo)

        self._hdr_btns: list = []

        def _hbtn(text, slot, tip):
            b = QPushButton(text)
            b.setFixedHeight(30)
            b.setCursor(QCursor(_ARROW_CUR))
            b.setToolTip(tip)
            b.setStyleSheet(self._header_btn_ss())
            b.clicked.connect(slot)
            h_lay.addWidget(b)
            self._hdr_btns.append(b)
            return b

        self._btn_new_tmpl  = _hbtn(tr("Создать"),  self._new_template,    tr("Создать новый шаблон"))
        self._btn_edit_tmpl = _hbtn(tr("Изменить"), self._edit_template,   tr("Изменить текущий шаблон"))
        self._btn_copy_tmpl = _hbtn(tr("Копировать"), self._copy_template,
                                    tr("Создать новый шаблон на основе текущего: зоны, поля и правила копируются"))
        self._btn_del_tmpl  = _hbtn(tr("Удалить"),  self._delete_template, tr("Удалить текущий шаблон"))

        h_lay.addSpacing(8)
        self._ver_lbl = QLabel(f"v{VERSION}")
        self._ver_lbl.setStyleSheet(self._header_dim_ss())
        h_lay.addWidget(self._ver_lbl)

        self._stripe = stripe = QFrame()
        stripe.setFixedHeight(3)
        stripe.setStyleSheet(f"background:{T.accent};")

        central  = QWidget()
        main_lay = QVBoxLayout(central)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.setSpacing(0)
        main_lay.addWidget(header)
        main_lay.addWidget(stripe)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        main_lay.addWidget(self.tabs, 1)

        self.status = QStatusBar()
        self.status.showMessage(tr("  Готов к работе"))
        self.setStatusBar(self.status)
        self.setCentralWidget(central)

        self._build_tab_common()
        self._build_tab_extract()
        self._build_tab_zones()
        self._build_tab_tools()
        self._build_tab_table()

    def _setup_menu(self):
        mb = self.menuBar()
        fm = mb.addMenu(tr("Файл"))
        fm.addAction(tr("Открыть Excel…"), self._open_excel, QKeySequence("Ctrl+E"))
        fm.addAction(tr("Открыть PDF…"),   self._open_pdf,   QKeySequence("Ctrl+O"))
        fm.addAction(tr("Открыть папку с PDF…"), self._open_folder, QKeySequence("Ctrl+Shift+O"))
        fm.addSeparator()
        fm.addAction(tr("Выход"), self.close, QKeySequence("Ctrl+Q"))

        tm = mb.addMenu(tr("Шаблоны"))
        tm.addAction(tr("Создать шаблон…"),  self._new_template)
        tm.addAction(tr("Изменить шаблон…"), self._edit_template)
        tm.addAction(tr("Копировать шаблон…"), self._copy_template)
        tm.addAction(tr("Удалить шаблон"),   self._delete_template)
        tm.addSeparator()
        tm.addAction(tr("🧙 Мастер: первый шаблон…"), self._run_wizard)

        vm = mb.addMenu(tr("Вид"))
        theme_menu = vm.addMenu(tr("Тема"))
        self._theme_group = QActionGroup(self)
        self._theme_group.setExclusive(True)
        self._theme_actions: dict = {}
        active = current_theme_name()
        for name in available_themes():
            act = QAction(tr(name), self, checkable=True)
            act.setChecked(name == active)
            act.triggered.connect(lambda _=False, n=name: self._choose_theme(n))
            self._theme_group.addAction(act)
            theme_menu.addAction(act)
            self._theme_actions[name] = act

        lang_menu = vm.addMenu(tr("Язык / Language"))
        self._lang_group = QActionGroup(self)
        self._lang_group.setExclusive(True)
        self._lang_actions: dict = {}
        for code, title in LANGUAGES.items():
            act = QAction(title, self, checkable=True)
            act.setChecked(code == LANG)
            act.triggered.connect(lambda _=False, c=code: self._choose_language(c))
            self._lang_group.addAction(act)
            lang_menu.addAction(act)
            self._lang_actions[code] = act

        hm = mb.addMenu(tr("Справка"))
        hm.addAction(tr("Инструкция"),  self._show_help, QKeySequence("F1"))
        hm.addAction(tr("▶ Демо-пример"), self._open_demo)
        hm.addAction(tr("🧙 Мастер: первый шаблон…"), self._run_wizard)
        hm.addSeparator()
        hm.addAction(tr("О программе"), self._show_about)

    def resizeEvent(self, event):
        """В узком окне название в шапке прячется, чтобы не обрезались кнопки."""
        super().resizeEvent(event)
        lbl = getattr(self, "_title_lbl", None)
        if lbl is not None:
            lbl.setVisible(self.width() >= 1280)

    def _choose_theme(self, name: str):
        set_theme(name)
        act = self._theme_actions.get(name)
        if act is not None and not act.isChecked():
            act.setChecked(True)
        self.status.showMessage(tr("  Тема: {name}", name=tr(name)), 3000)

    # ── Вкладка 1: Общие данные ──────────────────────────────────────────────
    def _build_tab_common(self):
        page = QWidget()
        self.tabs.addTab(page, tr("1. Общие данные"))
        outer = QVBoxLayout(page)
        outer.setContentsMargins(20, 16, 20, 16)

        h_row = QHBoxLayout()
        t = QLabel(tr("Общие данные пачки"))
        t.setStyleSheet(f"color:{C_TEXT}; font-size:13pt; font-weight:bold;")
        h_row.addWidget(t)
        s = QLabel(tr("Заполняются вручную и подставляются в каждую строку"))
        s.setStyleSheet(f"color:{C_TEXT2}; font-size:10pt;")
        h_row.addWidget(s)
        h_row.addStretch()
        outer.addLayout(h_row)
        outer.addSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(_NO_FRAME)
        self._manual_host = QWidget()
        self._manual_form = QFormLayout(self._manual_host)
        self._manual_form.setContentsMargins(16, 16, 16, 16)
        self._manual_form.setSpacing(12)
        self._manual_form.setLabelAlignment(_ALIGN_RIGHT)
        scroll.setWidget(self._manual_host)
        outer.addWidget(scroll, 1)

    # ── Вкладка 2: Распознавание ─────────────────────────────────────────────
    def _build_tab_extract(self):
        page = QWidget()
        self.tabs.addTab(page, tr("2. Распознавание"))
        outer = QVBoxLayout(page)
        outer.setContentsMargins(20, 16, 20, 16)
        outer.setSpacing(12)

        t = QLabel(tr("Распознавание документов"))
        t.setStyleSheet(f"color:{C_TEXT}; font-size:13pt; font-weight:bold;")
        outer.addWidget(t)

        c1 = Card()
        c1_lay = QVBoxLayout(c1)
        c1_lay.setContentsMargins(20, 16, 20, 16)
        lbl_pdf = QLabel(tr("📄  PDF файл или папка с PDF"))
        lbl_pdf.setStyleSheet(f"color:{C_TEXT2}; font-weight:bold;")
        c1_lay.addWidget(lbl_pdf)

        pdf_row = QHBoxLayout()
        btn_pdf = QPushButton(tr("Выбрать файл…"))
        btn_pdf.setObjectName("primary")
        btn_pdf.clicked.connect(self._open_pdf)
        pdf_row.addWidget(btn_pdf)
        btn_dir = QPushButton(tr("📁 Выбрать папку…"))
        btn_dir.setObjectName("primary")
        btn_dir.setToolTip(tr("Все PDF из папки обрабатываются как один файл (в порядке имён: 1, 2, 10…)"))
        btn_dir.clicked.connect(self._open_folder)
        pdf_row.addWidget(btn_dir)
        self._pdf_lbl = QLabel(tr("Файл не выбран"))
        self._pdf_lbl.setStyleSheet(f"color:{C_TEXT2};")
        pdf_row.addWidget(self._pdf_lbl, 1)
        c1_lay.addLayout(pdf_row)

        self._pdf_info = QLabel("")
        self._pdf_info.setStyleSheet(f"color:{C_ACCENT}; font-size:9pt;")
        c1_lay.addWidget(self._pdf_info)

        sep = QFrame(); sep.setFrameShape(_HLINE)
        c1_lay.addWidget(sep)

        settings_row = QHBoxLayout()
        settings_row.setSpacing(32)

        mode_grp   = QGroupBox(tr("Режим (из шаблона)"))
        mode_inner = QHBoxLayout(mode_grp)
        self._mode_lbl = QLabel("—")
        self._mode_lbl.setStyleSheet(f"color:{C_TEXT};")
        mode_inner.addWidget(self._mode_lbl)
        settings_row.addWidget(mode_grp)
        settings_row.addStretch()
        c1_lay.addLayout(settings_row)
        outer.addWidget(c1)

        # Карточка словарей (динамическая)
        self._dict_card = Card()
        self._dict_card_lay = QVBoxLayout(self._dict_card)
        self._dict_card_lay.setContentsMargins(20, 12, 20, 12)
        self._dict_card_lay.setSpacing(8)
        outer.addWidget(self._dict_card)

        # Карточка запуска
        c2 = Card()
        c2_lay = QHBoxLayout(c2)
        c2_lay.setContentsMargins(20, 16, 20, 16)
        self._run_btn = QPushButton(tr("▶  Извлечь данные"))
        self._run_btn.setObjectName("success")
        self._run_btn.setFixedHeight(46)
        self._run_btn.clicked.connect(self._start_extract)
        c2_lay.addWidget(self._run_btn)

        prog_wrap = QWidget()
        pw_lay    = QVBoxLayout(prog_wrap)
        pw_lay.setContentsMargins(16, 0, 0, 0)
        self._prog = QProgressBar()
        self._prog.setTextVisible(False)
        self._prog.setFixedHeight(6)
        pw_lay.addWidget(self._prog)
        self._prog_lbl = QLabel("")
        self._prog_lbl.setStyleSheet(f"color:{C_TEXT2}; font-size:9pt;")
        self._prog_lbl.setAlignment(_ALIGN_RIGHT)
        pw_lay.addWidget(self._prog_lbl)
        c2_lay.addWidget(prog_wrap, 1)
        outer.addWidget(c2)
        outer.addStretch()

    # ── Вкладка 3: Редактор зон ──────────────────────────────────────────────
    def _build_tab_zones(self):
        page = QWidget()
        self.tabs.addTab(page, tr("3. Редактор зон"))
        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        self._zones_desc = bar = QFrame()
        bar.setObjectName("zonesBar")
        bar.setStyleSheet(self._zones_desc_ss())
        bar_lay = QHBoxLayout(bar)
        bar_lay.setContentsMargins(12, 6, 12, 6)
        desc = QLabel(tr("Поставьте каждую цветную рамку на своё значение. Зоны сохраняются сразу. "
                         "Когда рамки на местах — следующий шаг:"))
        desc.setWordWrap(True)
        bar_lay.addWidget(desc, 1)
        btn_next = QPushButton(tr("🔮 Дальше: настроить поля…"))
        btn_next.setObjectName("primary")
        btn_next.setToolTip(tr("Открыть шаблон и подобрать вид каждого поля по тексту из рамок"))
        btn_next.clicked.connect(self._edit_template)
        bar_lay.addWidget(btn_next)
        outer.addWidget(bar)
        self._zone_editor = ZoneEditorWidget()
        self._zone_editor.zones_changed.connect(self._on_zones_changed)
        outer.addWidget(self._zone_editor, 1)

    # ── Вкладка 4: Инструменты ───────────────────────────────────────────────
    def _build_tab_tools(self):
        page = QWidget()
        self.tabs.addTab(page, tr("4. Инструменты"))
        outer = QVBoxLayout(page)
        outer.setContentsMargins(20, 16, 20, 16)
        outer.setSpacing(12)
        t = QLabel(tr("Сопоставление по словарям"))
        t.setStyleSheet(f"color:{C_TEXT}; font-size:13pt; font-weight:bold;")
        outer.addWidget(t)
        s = QLabel(tr("Для полей, отмеченных «Словарь» в шаблоне, здесь появляются кнопки сопоставления."))
        s.setStyleSheet(f"color:{C_TEXT2}; font-size:10pt;")
        outer.addWidget(s)

        self._tools_host = QWidget()
        self._tools_host_lay = QVBoxLayout(self._tools_host)
        self._tools_host_lay.setContentsMargins(0, 8, 0, 0)
        self._tools_host_lay.setSpacing(10)
        outer.addWidget(self._tools_host)
        outer.addStretch()

    # ── Вкладка 5: Таблица ───────────────────────────────────────────────────
    @staticmethod
    def _tbl_btn_ss() -> str:
        return (
            f"QPushButton {{ background-color: {T.chrome_raised}; color: {T.on_chrome}; "
            f"border: 1px solid {T.chrome_border}; padding: 7px 14px; border-radius: 6px; "
            "font-size: 9pt; font-weight: 500; }"
            f"QPushButton:hover {{ background-color: {T.accent}; color: white; border-color: {T.accent}; }}"
            f"QPushButton:pressed {{ background-color: {T.accent_hi}; color: white; }}"
            f"QPushButton:checked {{ background-color: {T.accent}; color: white; border-color: {T.accent}; }}"
        )

    @staticmethod
    def _tbl_btn_danger_ss() -> str:
        return (
            f"QPushButton {{ background-color: {T.danger_deep}; color: {T.danger_deep_fg}; "
            f"border: 1px solid {T.danger_deep_border}; padding: 7px 14px; border-radius: 6px; "
            "font-size: 9pt; font-weight: 500; }"
            f"QPushButton:hover {{ background-color: {T.danger}; color: white; border-color: {T.danger}; }}"
        )

    def _build_tab_table(self):
        page = QWidget()
        self.tabs.addTab(page, tr("5. Таблица"))
        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._tbl_toolbar = tb_frame = QFrame()
        tb_frame.setStyleSheet(f"background:{T.chrome}; border-bottom:2px solid {T.accent};")
        tb_lay = QHBoxLayout(tb_frame)
        tb_lay.setContentsMargins(10, 8, 10, 8)
        tb_lay.setSpacing(6)

        self._tbl_seps: list = []
        self._tbl_btns: list = []   # (кнопка, danger)

        def _sep():
            f = QFrame()
            f.setFrameShape(_VLINE)
            f.setStyleSheet(f"color:{T.chrome_border}; background:{T.chrome_border}; max-width:1px;")
            tb_lay.addWidget(f)
            self._tbl_seps.append(f)

        def _btn(text, slot, danger=False):
            b = QPushButton(text)
            b.setStyleSheet(self._tbl_btn_danger_ss() if danger else self._tbl_btn_ss())
            b.clicked.connect(slot)
            tb_lay.addWidget(b)
            self._tbl_btns.append((b, danger))
            return b

        _btn(tr("💾  В Excel"), self._save_excel)
        _btn(tr("📋  В буфер"), self._copy_clip)
        _btn(tr("📊  Отчёт"), self._copy_report)
        _sep()
        _btn(tr("✂️  Разбить на файлы"), self._split_docs)
        _btn(tr("🗜  Сжать"), self._compress_all)
        _btn(tr("👁  Просмотреть"), self._view_selected)
        _sep()
        jump_btn = _btn(tr("⚠  К проблеме"), self._jump_next_problem)
        jump_btn.setShortcut("F3")
        jump_btn.setToolTip(tr("Следующая строка: нет данных, ошибка или повтор номера (F3)"))
        scan_btn = _btn(tr("🖼  Скан"), self._toggle_scan_panel)
        scan_btn.setCheckable(True)
        scan_btn.setChecked(True)
        scan_btn.setToolTip(tr("Показать/скрыть участок скана для выбранной ячейки"))
        _sep()
        _btn(tr("🗑  Удалить"), self._del_row, danger=True)
        _btn(tr("✖  Очистить"), self._clear_all, danger=True)
        tb_lay.addStretch()
        outer.addWidget(tb_frame)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(1)
        self.tree.setHeaderLabels(["—"])
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionMode(_EXT_SEL)
        self.tree.setEditTriggers(_NO_EDIT)
        self.tree.itemDoubleClicked.connect(self._edit_cell)
        self.tree.setContextMenuPolicy(_CUSTOM_CTX)
        self.tree.customContextMenuRequested.connect(self._context_menu)
        self.tree.header().setSortIndicatorShown(True)

        self._scan_panel = self._build_scan_panel()
        split = QSplitter(_HORIZONTAL)
        split.addWidget(self.tree)
        split.addWidget(self._scan_panel)
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 0)
        split.setSizes([1040, 360])
        split.splitterMoved.connect(lambda *_: self._show_scan_pix())
        outer.addWidget(split, 1)

        self._totals_lbl = QLabel("")
        self._totals_lbl.setTextFormat(_qenum(Qt, "TextFormat", "RichText"))
        self._totals_lbl.setStyleSheet(self._totals_ss())
        outer.addWidget(self._totals_lbl)
        self.tree.selectionModel().currentChanged.connect(self._update_scan_panel)

    # ═════════════════════════════════════════════════════════════════════════
    # Справка
    # ═════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────────────────
# Английский перевод интерфейса: русский текст → английский текст.
# Чтобы добавить язык, сделайте такой же словарь и добавьте код в LANGUAGES и tr().
# ─────────────────────────────────────────────────────────────────────────────

_EN: dict = {
    "∅ нет данных":
        "∅ no data",
    "temp_{tag}_стр_{pages_str}.pdf":
        "temp_{tag}_p_{pages_str}.pdf",
    "Шпаргалка regex:  \\d — цифра;  \\d+ — одна или больше цифр;  \\d{6} — ровно 6 цифр;  [.,] — точка или запятая;  \\s — пробел;  .* — любой текст;  (…) — в «Очистке» это часть, которую надо взять;  ? — необязательно.\nПример очистки: 0*(\\d+) из «ОРДЕР № 00000012345» → «12345». Несколько групп склеиваются: [Пп][Кк]\\s*(\\d+)\\s*\\+\\s*(\\d+) из «ПК 400+30» → «40030».":
        "Regex cheat sheet:  \\d — digit;  \\d+ — one or more digits;  \\d{6} — exactly 6 digits;  [.,] — dot or comma;  \\s — space;  .* — any text;  (…) — in “Cleaning” this is the part to keep;  ? — optional.\nCleaning example: 0*(\\d+) on «ORDER No. 00000012345» → «12345». Several groups are joined: [Kk][Mm]\\s*(\\d+)\\s*\\+\\s*(\\d+) on «KM 400+30» → «40030».",
    "Предупреждение":
        "Warning",
    "Готово":
        "Done",
    "Ошибка":
        "Error",
    "Подтверждение":
        "Confirmation",
    "строка_{v}":
        "row_{v}",
    "файл":
        "file",
    "→ «{res}»  ✗ не прошло проверку":
        "→ «{res}»  ✗ failed the check",
    "▸ {label} · стр {v}:  «{shown_raw}»":
        "▸ {label} · p. {v}:  «{shown_raw}»",
    "{base} (копия)":
        "{base} (copy)",
    "Страниц: {n}  |  Документов{mode}: {cnt}":
        "Pages: {n}  |  Documents{mode}: {cnt}",
    "Сохранено строк: {n}\n{path}":
        "Rows saved: {n}\n{path}",
    "Готово!\nФайлов: {saved} ({npg}{nc})\nПапка: {out_dir}":
        "Done!\nFiles: {saved} ({npg}{nc})\nFolder: {out_dir}",
    "Документ:":
        "Document:",
    "Дата:":
        "Date:",
    "Поставщик:":
        "Supplier:",
    "Масса нетто, т:":
        "Net weight, t:",
    "№ документа":
        "Document No.",
    "Дата":
        "Date",
    "Поставщик":
        "Supplier",
    "Масса, т":
        "Weight, t",
    "№ п/п":
        "Row No.",
    "Материал":
        "Material",
    "Объём, м³":
        "Volume, m³",
    "Теперь перетащите каждую цветную рамку на своё значение. Затем нажмите «🔮 Дальше: настроить поля…» вверху.":
        "Now drag each colored frame onto its value. Then click “🔮 Next: set up fields…” at the top.",
    "Страница не существует.":
        "The page does not exist.",
    "очистки":
        "cleaning",
    "проверки":
        "check",
    "→ пусто: правило ничего не нашло":
        "→ empty: the rule found nothing",
    "  Документ: ":
        "  Document: ",
    "  Масштаб: ":
        "  Zoom: ",
    "👁 Текст всех зон":
        "👁 Text of all zones",
    "↺ Сбросить зоны":
        "↺ Reset zones",
    "  🖱  Перетащите прямоугольник — перемещение.  Тяните за угол/край — размер.  Колёсико — масштаб.  Пунктирная рамка — резервная зона.  Внизу — живой предпросмотр текста активной зоны.":
        "  🖱  Drag a rectangle to move it.  Drag a corner/edge to resize.  Mouse wheel — zoom.  Dashed frame — fallback zone.  Bottom line — live preview of the active zone text.",
    "Наведите/сдвиньте зону, чтобы увидеть распознанный текст.":
        "Hover over or move a zone to see the recognized text.",
    "Сопоставление: {field_name}":
        "Matching: {field_name}",
    "✔ Применить ко всем":
        "✔ Apply to all",
    "✔ Применить к выделенным":
        "✔ Apply to selected",
    "👁 Просмотреть документ":
        "👁 View document",
    "➕ Добавить в список":
        "➕ Add to list",
    "Добавить выделенное значение в словарь":
        "Add the selected value to the dictionary",
    "Добавить в словарь:\n":
        "Add to dictionary:\n",
    "Редактирование":
        "Edit",
    "Новое значение:":
        "New value:",
    "DPI для сжатия:":
        "DPI for compression:",
    "Качество JPEG (1–100):":
        "JPEG quality (1–100):",
    "Оттенки серого (меньше файл, для ч/б документов)":
        "Grayscale (smaller file, for black-and-white documents)",
    "Режим:":
        "Mode:",
    "▶ Начать":
        "▶ Start",
    "Авто (первое извлекаемое поле)":
        "Auto (first extracted field)",
    "Текст зоны проходит три шага: <b>1. обработка</b> (готовые исправления распознавания) → <b>2. очистка</b> (что вырезать) → <b>3. проверка</b> (правильный ли формат). Обычно хватает выбора в колонке «Что это за поле».":
        "The zone text goes through three steps: <b>1. processing</b> (ready-made OCR fixes) → <b>2. cleaning</b> (what to cut out) → <b>3. check</b> (is the format right). Usually choosing a type in the “What is this field” column is enough.",
    "1. Обработка (OCR):":
        "1. Processing (OCR):",
    "пусто — без очистки;  напр. 0*(\\d+)":
        "empty — no cleaning;  e.g. 0*(\\d+)",
    "2. Очистка (regex):":
        "2. Cleaning (regex):",
    "3. Проверка формата:":
        "3. Format check:",
    "своё правило проверки, напр. \\d{6}":
        "your own check rule, e.g. \\d{6}",
    "    свой regex проверки:":
        "    custom check regex:",
    "вставьте распознанный текст зоны":
        "paste the recognized zone text",
    "Отмена":
        "Cancel",
    "Копия шаблона «{copy_of}»":
        "Copy of template «{copy_of}»",
    "Изменение шаблона":
        "Edit template",
    "Создание шаблона":
        "New template",
    "1. Основное":
        "1. Basics",
    "Например: Товарные накладные":
        "For example: Delivery notes",
    "Название шаблона:":
        "Template name:",
    "Страниц на документ:":
        "Pages per document:",
    "2. Поля, заполняемые вручную":
        "2. Fields entered manually",
    "3. Поля из PDF (у каждого — своя зона на странице)":
        "3. Fields from PDF (each has its own zone on the page)",
    "🔮 Подобрать вид для всех полей":
        "🔮 Detect type for all fields",
    "Для каждого поля: пример из открытого PDF → подбор вида":
        "For each field: sample from the open PDF → type detection",
    "4. Вычисляемые поля (формула из других полей)":
        "4. Calculated fields (formula from other fields)",
    "5. Ключевое поле":
        "5. Key field",
    "6. Порядок столбцов в таблице и Excel (перетаскивайте мышью)":
        "6. Column order in the table and Excel (drag with the mouse)",
    "Сбросить: ручные → из PDF → вычисляемые":
        "Reset: manual → from PDF → calculated",
    "7. Отчёт — кнопка «📊 Отчёт» на вкладке «Таблица»":
        "7. Report — the “📊 Report” button on the “Table” tab",
    "-\n[Дата]\nПоставщик = ООО «Ромашка»\n[№ документа]\n[Вес]":
        "-\n[Date]\nSupplier = Acme Ltd.\n[Document No.]\n[Weight]",
    "💾 Сохранить шаблон":
        "💾 Save template",
    "Поле {v}":
        "Field {v}",
    "⚙ Своё правило (вручную)":
        "⚙ Custom rule (manual)",
    "напр. ОРДЕР № 00000012345":
        "e.g. ORDER No. 00000012345",
    "Пример распознанного текста зоны. Нужен только для проверки правила.":
        "A sample of recognized zone text. Used only to test the rule.",
    "Подобрать вид поля по примеру. Если пример пуст — текст берётся из открытого PDF.":
        "Detect the field type from the sample. If the sample is empty, the text is taken from the open PDF.",
    "Ручные правила: обработка, очистка, проверка (regex)":
        "Manual rules: processing, cleaning, check (regex)",
    "📄  {text}      — из PDF":
        "📄  {text}      — from PDF",
    "Шаблон «{name}» создан.":
        "Template «{name}» created.",
    "{base} (копия {n})":
        "{base} (copy {n})",
    "Создан шаблон «{name}» — копия «{src_name}».":
        "Template «{name}» created — a copy of «{src_name}».",
    "Шаблон «{name}» обновлён.":
        "Template «{name}» updated.",
    "Шаблон «{name}» удалён.":
        "Template «{name}» deleted.",
    "🗂  Словари для сопоставления":
        "🗂  Dictionaries for matching",
    " (по {ppd} стр.)":
        " ({ppd} p. each)",
    "Словарь для «{name}»":
        "Dictionary for «{name}»",
    "Текст (*.txt);;Все (*.*)":
        "Text (*.txt);;All (*.*)",
    "В словарь «{name}» добавлено: {value}":
        "Added to dictionary «{name}»: {value}",
    "Зоны обновлены и сохранены.":
        "Zones updated and saved.",
    "Выберите PDF":
        "Choose a PDF",
    "PDF (*.pdf);;Все (*.*)":
        "PDF (*.pdf);;All (*.*)",
    "Папка с PDF-файлами":
        "Folder with PDF files",
    "Склеиваю PDF…":
        "Merging PDFs…",
    "Папка с PDF":
        "Folder with PDFs",
    "📁 {v}  (файлов: {ok_files})":
        "📁 {v}  (files: {ok_files})",
    "Папка: {folder}  —  файлов {ok_files}, страниц {v}":
        "Folder: {folder}  —  {ok_files} files, {v} pages",
    "Выберите Excel":
        "Choose an Excel file",
    "Excel (*.xlsx *.xls);;Все (*.*)":
        "Excel (*.xlsx *.xls);;All (*.*)",
    "Обработано: {done} / {total}":
        "Processed: {done} / {total}",
    "Ошибка извлечения":
        "Extraction error",
    "Всего записей: {n}":
        "Total rows: {n}",
    "Записей: <b>{n}</b>":
        "Rows: <b>{n}</b>",
    "Проблема {v} из {n}  —  строка {v2}  ({kind_txt})":
        "Problem {v} of {n}  —  row {v2}  ({kind_txt})",
    "Значение изменено.":
        "Value changed.",
    "Удалено строк: {n}.":
        "Rows deleted: {n}.",
    "Очистить все данные?":
        "Clear all data?",
    "🗑 Удалить строку":
        "🗑 Delete row",
    "Скан ячейки":
        "Cell scan",
    "впишите правильное значение и нажмите Enter":
        "type the correct value and press Enter",
    "✔ Применить → следующая проблема":
        "✔ Apply → next problem",
    "Enter — сохранить и перейти к следующей проблеме.\nF3 — следующая проблема без правки.\nСиняя рамка — основная зона, оранжевая пунктирная — резервная.":
        "Enter — save and go to the next problem.\nF3 — next problem without editing.\nBlue frame — main zone, orange dashed — fallback zone.",
    "Строка {v}  •  {name}":
        "Row {v}  •  {name}",
    "Сохранено в Excel: {n} строк → {path}":
        "Saved to Excel: {n} rows → {path}",
    "Сохранено":
        "Saved",
    "Скопировано {n} строк → Ctrl+V":
        "Copied {n} rows → Ctrl+V",
    "Скопировано {n} строк.":
        "Copied {n} rows.",
    "Отчёт: скопировано {n} строк → Ctrl+V в Google Таблицы":
        "Report: copied {n} rows → Ctrl+V into Google Sheets",
    "Отчёт скопирован: {n} строк.":
        "Report copied: {n} rows.",
    "Сжатие всей пачки PDF":
        "Compress the whole PDF batch",
    "Папка для сохранения":
        "Folder to save to",
    "Готово! Файлов: {n}\nDPI: {dpi},  JPEG: {quality}%\n\n{details}\n\nПапка: {out_dir}":
        "Done! Files: {n}\nDPI: {dpi},  JPEG: {quality}%\n\n{details}\n\nFolder: {out_dir}",
    "Сжатие завершено":
        "Compression finished",
    "Открыть папку?":
        "Open the folder?",
    "Открыть?":
        "Open?",
    "Разбивка PDF на файлы":
        "Split PDF into files",
    ", сжатие {dpi} DPI":
        ", compressed to {dpi} DPI",
    "{ppd} стр/файл":
        "{ppd} p./file",
    "1 стр/файл":
        "1 p./file",
    "\nОшибок: {errors}":
        "\nErrors: {errors}",
    "Разбивка завершена":
        "Split finished",
    "Для демо нужен PyMuPDF 1.23 или новее:\npip install -U PyMuPDF":
        "The demo needs PyMuPDF 1.23 or newer:\npip install -U PyMuPDF",
    "ТОВАРНАЯ НАКЛАДНАЯ":
        "DELIVERY NOTE",
    "Демо-документ ZoneExtract. Все данные вымышленные.":
        "ZoneExtract demo document. All data is fictional.",
    "Отпустил: ____________":
        "Released by: ____________",
    "Принял: ____________":
        "Received by: ____________",
    "Лист {n} из {total}":
        "Sheet {n} of {total}",
    "Демо: накладные":
        "Demo: delivery notes",
    "Мастер: первый шаблон":
        "First template wizard",
    "← Назад":
        "← Back",
    "📄 Выбрать PDF-образец…":
        "📄 Choose a sample PDF…",
    "не выбран (можно выбрать позже)":
        "not chosen (you can choose it later)",
    "Страниц на один документ:":
        "Pages per document:",
    "№ документа\nДата\nПоставщик\nМасса, т":
        "Document No.\nDate\nSupplier\nWeight, t",
    "Добавить столбец «№ п/п» — номер строки по порядку":
        "Add a “Row No.” column — row number in order",
    "PDF-образец":
        "Sample PDF",
    "{name}  ({n} стр.)":
        "{name}  ({n} p.)",
    "Шаг 1 из 3 — образец документа":
        "Step 1 of 3 — sample document",
    "Шаг 2 из 3 — какие данные нужны":
        "Step 2 of 3 — what data you need",
    "Шаг 3 из 3 — что дальше":
        "Step 3 of 3 — what next",
    "Демо загружено: 6 накладных и шаблон «{name}».\n\nСейчас открыт «Редактор зон»: цветные рамки показывают, откуда берутся значения.\n\nДальше: вкладка «2. Распознавание» → «▶ Извлечь данные».\nВ таблице будут три проблемы — пустая ячейка, ошибка формата и повтор номера. Нажимайте F3 и исправляйте их по скану справа.":
        "Demo loaded: 6 delivery notes and the template «{name}».\n\nThe “Zone editor” is open now: the colored frames show where the values come from.\n\nNext: tab “2. Recognition” → “▶ Extract data”.\nThe table will have three problems — an empty cell, a format error and a repeated number. Press F3 and fix them using the scan on the right.",
    "Демо-пример":
        "Demo",
    "Добро пожаловать":
        "Welcome",
    "ZoneExtract переносит данные из PDF-документов в таблицу.\n\nВы один раз показываете, где на бланке лежит каждое значение, дальше программа сама обрабатывает пачки таких бланков.\n\nС чего начать?":
        "ZoneExtract moves data from PDF documents into a table.\n\nYou show once where each value is on the form, and then the program processes batches of such forms by itself.\n\nWhere do you want to start?",
    "Инструкция":
        "User guide",
    "Закрыть":
        "Close",
    "О программе":
        "About",
    "<h3>ZoneExtract {version}</h3><p>Извлечение данных из распознанных PDF по зонам и шаблонам.</p><p>Открытый исходный код, лицензия MIT.<br><a href='{url}'>{url}</a></p><p>Зависимости: PySide6, PyMuPDF, pandas, openpyxl.</p>":
        "<h3>ZoneExtract {version}</h3><p>Data extraction from OCR'd PDFs by zones and templates.</p><p>Open source, MIT license.<br><a href='{url}'>{url}</a></p><p>Dependencies: PySide6, PyMuPDF, pandas, openpyxl.</p>",
    "Шаблон:":
        "Template:",
    "Создать":
        "New",
    "Создать новый шаблон":
        "Create a new template",
    "Изменить":
        "Edit",
    "Изменить текущий шаблон":
        "Edit the current template",
    "Копировать":
        "Copy",
    "Создать новый шаблон на основе текущего: зоны, поля и правила копируются":
        "Create a new template based on the current one: zones, fields and rules are copied",
    "Удалить":
        "Delete",
    "Удалить текущий шаблон":
        "Delete the current template",
    "  Готов к работе":
        "  Ready",
    "Файл":
        "File",
    "Открыть Excel…":
        "Open Excel…",
    "Открыть PDF…":
        "Open PDF…",
    "Открыть папку с PDF…":
        "Open folder with PDFs…",
    "Выход":
        "Exit",
    "Шаблоны":
        "Templates",
    "Создать шаблон…":
        "New template…",
    "Изменить шаблон…":
        "Edit template…",
    "Копировать шаблон…":
        "Copy template…",
    "Удалить шаблон":
        "Delete template",
    "🧙 Мастер: первый шаблон…":
        "🧙 First template wizard…",
    "Вид":
        "View",
    "Тема":
        "Theme",
    "Язык / Language":
        "Язык / Language",
    "Справка":
        "Help",
    "▶ Демо-пример":
        "▶ Demo",
    "  Тема: {name}":
        "  Theme: {name}",
    "1. Общие данные":
        "1. Common data",
    "Общие данные пачки":
        "Common data for the batch",
    "Заполняются вручную и подставляются в каждую строку":
        "Entered manually and put into every row",
    "2. Распознавание":
        "2. Recognition",
    "Распознавание документов":
        "Document recognition",
    "📄  PDF файл или папка с PDF":
        "📄  PDF file or folder with PDFs",
    "Выбрать файл…":
        "Choose file…",
    "📁 Выбрать папку…":
        "📁 Choose folder…",
    "Все PDF из папки обрабатываются как один файл (в порядке имён: 1, 2, 10…)":
        "All PDFs in the folder are processed as one file (sorted by name: 1, 2, 10…)",
    "Файл не выбран":
        "No file chosen",
    "Режим (из шаблона)":
        "Mode (from template)",
    "▶  Извлечь данные":
        "▶  Extract data",
    "3. Редактор зон":
        "3. Zone editor",
    "Поставьте каждую цветную рамку на своё значение. Зоны сохраняются сразу. Когда рамки на местах — следующий шаг:":
        "Put each colored frame onto its value. Zones are saved immediately. When the frames are in place — next step:",
    "🔮 Дальше: настроить поля…":
        "🔮 Next: set up fields…",
    "Открыть шаблон и подобрать вид каждого поля по тексту из рамок":
        "Open the template and detect the type of each field from the text in the frames",
    "4. Инструменты":
        "4. Tools",
    "Сопоставление по словарям":
        "Dictionary matching",
    "Для полей, отмеченных «Словарь» в шаблоне, здесь появляются кнопки сопоставления.":
        "Fields marked “Dictionary” in the template get matching buttons here.",
    "5. Таблица":
        "5. Table",
    "💾  В Excel":
        "💾  To Excel",
    "📋  В буфер":
        "📋  To clipboard",
    "📊  Отчёт":
        "📊  Report",
    "✂️  Разбить на файлы":
        "✂️  Split into files",
    "🗜  Сжать":
        "🗜  Compress",
    "👁  Просмотреть":
        "👁  View",
    "⚠  К проблеме":
        "⚠  Next problem",
    "Следующая строка: нет данных, ошибка или повтор номера (F3)":
        "Next row with no data, an error or a repeated number (F3)",
    "🖼  Скан":
        "🖼  Scan",
    "Показать/скрыть участок скана для выбранной ячейки":
        "Show/hide the scan fragment for the selected cell",
    "🗑  Удалить":
        "🗑  Delete",
    "✖  Очистить":
        "✖  Clear",
    "не удалось прочитать настройки интерфейса":
        "could not read UI settings",
    "не удалось сохранить настройки интерфейса":
        "could not save UI settings",
    "не удалось открыть папку":
        "could not open folder",
    "Ни один PDF из папки не удалось прочитать.":
        "None of the PDFs in the folder could be read.",
    "папка_":
        "folder_",
    "недопустимая операция: {name__}":
        "operation not allowed: {name__}",
    "допустимы только числа":
        "only numbers are allowed",
    "неизвестное имя «{id}» — поля пишутся в [скобках]":
        "unknown name «{id}» — write fields in [brackets]",
    "не удалось прочитать конфигурацию шаблонов":
        "could not read template configuration",
    "не удалось сохранить конфигурацию шаблонов":
        "could not save template configuration",
    "Откройте PDF-файл, чтобы разместить зоны.":
        "Open a PDF file to place zones.",
    "Для этого документа нет страниц.":
        "This document has no pages.",
    "   ✓ проверка пройдена":
        "   ✓ check passed",
    "   ✗ не совпало с проверкой":
        "   ✗ does not match the check",
    "Нет открытого PDF или зон для предпросмотра.":
        "No open PDF or zones to preview.",
    "{mark} {label} (стр {v}): «{val}»{status}":
        "{mark} {label} (p. {v}): «{val}»{status}",
    "Для этого документа нет зон.":
        "This document has no zones.",
    "Распознано":
        "Recognized",
    "Лучшее совпадение":
        "Best match",
    "Сходство":
        "Similarity",
    "Использовать":
        "Use",
    "Порог автозамены (%):":
        "Auto-replace threshold (%):",
    "Все значения из колонки «Использовать» уже в словаре.":
        "All values in the “Use” column are already in the dictionary.",
    "Нет новых":
        "Nothing new",
    "Подтвердите":
        "Confirm",
    "Выделите строки для применения.":
        "Select the rows to apply to.",
    "Подсказка":
        "Hint",
    "PDF не открыт.":
        "No PDF is open.",
    "Выберите строку.":
        "Select a row.",
    " МБ":
        " MB",
    "Макс. размер части:":
        "Max. part size:",
    "(0 = не разбивать)":
        "(0 = do not split)",
    "Уменьшить DPI при разбивке":
        "Reduce DPI when splitting",
    "Сжатие:":
        "Compression:",
    "Правила поля «{field_name}»":
        "Field rules «{field_name}»",
    "Правила поля":
        "Field rules",
    "Пример текста зоны:":
        "Zone text sample:",
    "Вставьте пример — здесь появится результат.":
        "Paste a sample — the result will appear here.",
    "Новое название. Исходный шаблон не меняется.":
        "New name. The original template is not changed.",
    "Количество полей:":
        "Number of fields:",
    "сколько страниц скана — один документ (1–{MAX_DOC_PAGES}). Двусторонний скан накладной = 2.":
        "how many scanned pages make one document (1–{MAX_DOC_PAGES}). A double-sided scan = 2.",
    "Одинаковы для всей пачки: год, дата, перевозчик, материал… Заполняются на вкладке «1. Общие данные». Галочка «№ по порядку» — номер строки ставится сам (1, 2, 3…) и продолжается после сохранения в Excel.":
        "Same for the whole batch: year, date, carrier, material… Entered on tab “1. Common data”. The “Row number” checkbox numbers rows automatically (1, 2, 3…) and continues after saving to Excel.",
    "Как настроить поле: нажмите 🔮 — программа возьмёт текст из рамки поля в открытом PDF и сама подберёт вид поля. Справа виден результат: ✓ — правило работает, ✗ — нет. Если PDF не открыт — вставьте пример текста вручную. Ручные правила (regex) — кнопка ⚙.":
        "How to set up a field: click 🔮 — the program takes the text from the field frame in the open PDF and detects the field type. The result is shown on the right: ✓ — the rule works, ✗ — it does not. If no PDF is open, paste a text sample manually. Manual rules (regex) — the ⚙ button.",
    "Название столбца":
        "Column name",
    "Что это за поле":
        "What is this field",
    "Пример текста зоны":
        "Zone text sample",
    "Результат":
        "Result",
    "Стр.":
        "Page",
    "На какой странице документа поле":
        "Which page of the document the field is on",
    "Слов.":
        "Dict.",
    "Сопоставлять со справочником (словарём)":
        "Match against a reference list (dictionary)",
    "Резерв":
        "Fallback",
    "Запасная зона: читается, если основная пуста":
        "Fallback zone: read if the main zone is empty",
    "Поле в формуле — в квадратных скобках: [Тонн] / 1,6. Можно + − * / и скобки. Округление математическое (28,375 → 28,38). Пустое или нечисловое исходное поле даёт пустой результат.":
        "A field in a formula goes in square brackets: [Tonnes] / 1.6. You can use + − * / and brackets. Rounding is half-up (28.375 → 28.38). An empty or non-numeric source field gives an empty result.",
    "Формула":
        "Formula",
    "Знаков после «,»":
        "Decimal places",
    "Поле:":
        "Field:",
    "Обычно № документа. По нему ищутся повторы (в пачке и в Excel) и называются файлы при «Разбить на файлы».":
        "Usually the document number. Used to find duplicates (in the batch and in Excel) and to name files in “Split into files”.",
    "Выше":
        "Up",
    "Ниже":
        "Down",
    "Одна строка = одна колонка отчёта.   [Имя поля] — значение поля.   Подпись = значение — постоянный текст (его можно поправить перед копированием).   - — пустая колонка.   Пусто — отчёт не настроен.":
        "One line = one report column.   [Field name] — field value.   Label = value — fixed text (you can edit it before copying).   - — empty column.   Empty — no report.",
    "Название поля {v}":
        "Field name {v}",
    "№ по порядку":
        "Row number",
    "Номер строки ставится автоматически: начало задаётся на вкладке «1. Общие данные», после сохранения в Excel продолжается.":
        "The row number is set automatically: the start is set on tab “1. Common data”; it continues after saving to Excel.",
    "Стр. {v}":
        "Page {v}",
    "не подобрать — выберите вручную":
        "cannot detect — choose manually",
    "Подобрано полей: {done}. Без примера: {miss}.\n\nДля полей без примера поставьте рамку на текст (вкладка «3. Редактор зон») или вставьте пример вручную.":
        "Fields detected: {done}. Without a sample: {miss}.\n\nFor fields without a sample, put the frame on the text (tab “3. Zone editor”) or paste a sample manually.",
    "Подбор вида поля":
        "Field type detection",
    "своё правило":
        "custom rule",
    "Вычисляемое поле {v}":
        "Calculated field {v}",
    "напр. [Тонн] / 1,6":
        "e.g. [Tonnes] / 1.6",
    "Укажите название шаблона.":
        "Enter a template name.",
    "Проверка":
        "Check",
    "Шаблон «{name}» уже существует.":
        "Template «{name}» already exists.",
    "Все названия полей должны быть заполнены.":
        "All field names must be filled in.",
    "Сначала выберите шаблон, который нужно скопировать.":
        "First choose the template you want to copy.",
    "Сначала выберите или создайте шаблон.":
        "First choose or create a template.",
    "Удалить шаблон «{name}»?":
        "Delete template «{name}»?",
    "Удаление":
        "Delete",
    "Шаблон: {name}":
        "Template: {name}",
    "Сначала создайте или выберите шаблон.":
        "First create or choose a template.",
    "Шаблон":
        "Template",
    "В шаблоне нет полей со словарём.":
        "The template has no dictionary fields.",
    "Загрузить словарь…":
        "Load dictionary…",
    "Не загружен":
        "Not loaded",
    "Нет полей со словарём. Отметьте «Словарь» у поля при создании шаблона.":
        "No dictionary fields. Tick “Dict.” for a field when you create the template.",
    "🔍 Сопоставить…":
        "🔍 Match…",
    "страница":
        "page",
    "{p} {word} на документ":
        "{p} {word} per document",
    "Файлов: {n_files}  |  ":
        "Files: {n_files}  |  ",
    "  |  добавлено пустых страниц: {blanks}":
        "  |  blank pages added: {blanks}",
    "{v} ({n} записей)":
        "{v} ({n} entries)",
    "Словарь «{name}»: загружено {n} записей.":
        "Dictionary «{name}»: {n} entries loaded.",
    "Сначала извлеките данные.":
        "Extract the data first.",
    "Загрузите словарь для «{name}» на вкладке «Распознавание».":
        "Load a dictionary for «{name}» on the “Recognition” tab.",
    "Поле «{name}» обновлено по словарю.":
        "Field «{name}» updated from the dictionary.",
    "Открыт: {p}":
        "Opened: {p}",
    "В подпапках тоже есть PDF.\nВключить их в обработку?":
        "Subfolders also contain PDFs.\nInclude them?",
    "Подпапки":
        "Subfolders",
    "Файл {v} из {total}:\n{name}":
        "File {v} of {total}:\n{name}",
    "\n  … и ещё {v}":
        "\n  … and {v} more",
    "Число страниц не кратно {ppd} — добавлены пустые страницы, чтобы каждый файл начинался с нового документа:\n{lst}{more}":
        "The page count is not a multiple of {ppd} — blank pages were added so that each file starts a new document:\n{lst}{more}",
    "Пропущены (не открылись):\n{lst}":
        "Skipped (could not open):\n{lst}",
    "Папка загружена с замечаниями":
        "Folder loaded with notes",
    "В PDF не найден распознанный текст.\n\nПрограмма читает только текстовый слой. Распознайте скан в ABBYY FineReader («Сохранить как → PDF с возможностью поиска») и откройте файл снова.":
        "No recognized text found in the PDF.\n\nThe program reads only the text layer. Run OCR on the scan (for example ABBYY FineReader: “Save as → Searchable PDF”, or the free OCRmyPDF) and open the file again.",
    "Нет текстового слоя":
        "No text layer",
    "{v}, стр. {v2}":
        "{v}, p. {v2}",
    "Загружено {n} записей из {p}":
        "Loaded {n} rows from {p}",
    "Сначала выберите PDF файл.":
        "Choose a PDF file first.",
    "Извлечено записей: {n}.\nПроблемных строк: {probs}.\n\nКурсор стоит на первой проблеме. Справа — участок скана.\nВпишите правильное значение и нажмите Enter — откроется следующая.\nF3 — следующая проблема без правки.":
        "Rows extracted: {n}.\nProblem rows: {probs}.\n\nThe cursor is on the first problem. The scan fragment is on the right.\nType the correct value and press Enter — the next one opens.\nF3 — next problem without editing.",
    "Извлечено записей: {n}.\nПроблем не найдено.":
        "Rows extracted: {n}.\nNo problems found.",
    "Повтор в пачке: строки ":
        "Repeated in the batch: rows ",
    "без значения: {total_missing} (жёлтые ∅)":
        "no value: {total_missing} (yellow ∅)",
    "ошибок проверки: {total_err} (красные)":
        "check errors: {total_err} (red)",
    "повторов «{key_field}»: {dups} (фиолетовые)":
        "repeated «{key_field}»: {dups} (purple)",
    " <span style='color:{C_WARNING}'>(без {v})</span>":
        " <span style='color:{C_WARNING}'>(missing {v})</span>",
    "<span style='color:{C_DANGER}'>⚠ проблемных строк: {problems} — F3</span>":
        "<span style='color:{C_DANGER}'>⚠ problem rows: {problems} — F3</span>",
    "<span style='color:{C_SUCCESS}'>✓ проблем нет</span>":
        "<span style='color:{C_SUCCESS}'>✓ no problems</span>",
    "Таблица пуста.":
        "The table is empty.",
    "Проблем нет: все ячейки заполнены, проверены, повторов нет.":
        "No problems: all cells are filled, checked, no duplicates.",
    "нет данных":
        "no data",
    "ошибка проверки":
        "check error",
    "повтор номера":
        "repeated number",
    "«{col_name}» считается по формуле:\n{formula}\n\nИзмените исходное поле — значение пересчитается.":
        "«{col_name}» is calculated by a formula:\n{formula}\n\nChange the source field — the value will be recalculated.",
    "Вычисляемое поле":
        "Calculated field",
    "Данные очищены.":
        "Data cleared.",
    "✏ Редактировать «{cname}»":
        "✏ Edit «{cname}»",
    "Сначала откройте PDF файл.":
        "Open a PDF file first.",
    "Значение в таблице:":
        "Value in the table:",
    "Выберите ячейку в таблице — здесь появится этот участок скана и распознанный текст.":
        "Select a cell in the table — the scan fragment and the recognized text will appear here.",
    "Считается по формуле: {formula}\nИсправьте исходное поле — значение пересчитается.":
        "Calculated by a formula: {formula}\nFix the source field — the value will be recalculated.",
    "Нет данных для сохранения.":
        "No data to save.",
    "Сохранить Excel":
        "Save Excel",
    "документы.xlsx":
        "documents.xlsx",
    "Данные":
        "Data",
    "\n\nСледующая пачка начнётся с: ":
        "\n\nThe next batch starts with: ",
    "Нет данных.":
        "No data.",
    "В шаблоне не настроен отчёт.\nШаблоны → Изменить шаблон… → блок «7. Отчёт».":
        "The template has no report.\nTemplates → Edit template… → block “7. Report”.",
    "Отчёт":
        "Report",
    "Параметры отчёта":
        "Report settings",
    "Значения применятся ко всем строкам и запомнятся в шаблоне.":
        "The values apply to all rows and are saved in the template.",
    "Скопировать":
        "Copy",
    "  {v}: {v2:.1f} МБ":
        "  {v}: {v2:.1f} MB",
    "Таблица пустая.":
        "The table is empty.",
    "по {ppd} стр.":
        "{ppd} p. each",
    "по 1 стр.":
        "1 p. each",
    "документ_{v}":
        "document_{v}",
    "Шаблон — это настройка под <b>один вид бланка</b>: где на странице лежит каждое значение и как его проверить. Настраивается один раз, потом пачки обрабатываются в пару кликов.<br><br>Возьмите PDF с несколькими такими бланками. PDF должен быть <b>с текстовым слоем</b> (после распознавания, например в FineReader: «PDF с возможностью поиска»).":
        "A template is a setup for <b>one kind of form</b>: where each value is on the page and how to check it. You set it up once, then batches are processed in a couple of clicks.<br><br>Take a PDF with several such forms. The PDF must have a <b>text layer</b> (after OCR, for example ABBYY FineReader “Searchable PDF” or the free OCRmyPDF).",
    "Обычно 1. Если бланк двусторонний и сканируется с обеих сторон — 2.":
        "Usually 1. If the form is double-sided and both sides are scanned — 2.",
    "<b>Какие значения взять с бланка?</b> По одному на строку. Это будут столбцы таблицы. Каждому полю на следующем шаге вы поставите рамку на странице.":
        "<b>Which values do you need from the form?</b> One per line. They become the table columns. On the next step you put a frame on the page for each field.",
    "<b>Общие значения для всей пачки</b> (необязательно) — их нет на бланке, вы впишете их один раз: например, материал или объект.":
        "<b>Common values for the whole batch</b> (optional) — they are not on the form; you type them once: for example, material or site.",
    "<b>Шаблон почти готов. Осталось три действия:</b><br><br><b>1. Расставить рамки.</b> Откроется вкладка «3. Редактор зон». На странице — цветные рамки, по одной на поле. Перетащите каждую на своё значение. Внизу виден текст, который программа прочитала из рамки.<br><br><b>2. Настроить поля.</b> Нажмите «🔮 Дальше: настроить поля…», затем «🔮 Подобрать вид для всех полей». Программа сама поймёт, где число, где дата, где номер. Сохраните шаблон.<br><br><b>3. Извлечь.</b> Вкладка «2. Распознавание» → «▶ Извлечь данные». Результат — на вкладке «5. Таблица».<br><br>Подробно — «Справка → Инструкция».":
        "<b>The template is almost ready. Three actions are left:</b><br><br><b>1. Place the frames.</b> Tab “3. Zone editor” opens. There are colored frames on the page, one per field. Drag each onto its value. The bottom line shows the text the program read from the frame.<br><br><b>2. Set up the fields.</b> Click “🔮 Next: set up fields…”, then “🔮 Detect type for all fields”. The program detects where the number, the date and the document number are. Save the template.<br><br><b>3. Extract.</b> Tab “2. Recognition” → “▶ Extract data”. The result is on tab “5. Table”.<br><br>Details — “Help → User guide”.",
    "Создать шаблон":
        "Create template",
    "Далее →":
        "Next →",
    "Щебень":
        "Crushed stone",
    "Откройте PDF-образец: вкладка «2. Распознавание» → «Выбрать файл…».":
        "Open the sample PDF: tab “2. Recognition” → “Choose file…”.",
    "🧙 Мастер: первый шаблон":
        "🧙 First template wizard",
    "Позже":
        "Later",
    "перекраска виджета при смене темы":
        "widget repaint on theme change",
    "Отменено пользователем":
        "Cancelled by user",
    "обработчик «{post}»":
        "processor «{post}»",
    "{name} · резерв":
        "{name} · fallback",
    "·рез":
        "·fb",
    "Не удалось открыть документ:\n{e}":
        "Could not open the document:\n{e}",
    "напр. ":
        "e.g. ",
    "Нет примера текста для этого поля.\n\nВариант 1: откройте PDF (вкладка «2. Распознавание»), поставьте рамку поля на нужное место (вкладка «3. Редактор зон») и снова нажмите 🔮 — пример возьмётся из PDF сам.\n\nВариант 2: вставьте в «Пример» текст вручную.":
        "No text sample for this field.\n\nOption 1: open a PDF (tab “2. Recognition”), put the field frame in the right place (tab “3. Zone editor”) and click 🔮 again — the sample is taken from the PDF.\n\nOption 2: paste the text into “Sample” manually.",
    "Подходит также:\n":
        "Also fits:\n",
    "(ручное поле {v})":
        "(manual field {v})",
    "(поле из PDF {v})":
        "(PDF field {v})",
    "(вычисляемое поле {v})":
        "(calculated field {v})",
    "✍  {text}      — вручную":
        "✍  {text}      — manual",
    "🧮  {text}      — формула":
        "🧮  {text}      — formula",
    "Названия полей повторяются: ":
        "Field names repeat: ",
    "Укажите формулу для поля «{nm}».":
        "Enter a formula for field «{nm}».",
    "В отчёте указано поле [{a}], которого нет в шаблоне.":
        "The report uses field [{a}], which is not in the template.",
    "В текущем шаблоне нет полей для ручного ввода.":
        "The current template has no manual fields.",
    "Шаблон не выбран.":
        "No template selected.",
    "№ по порядку: первая строка получит этот номер, дальше +1. После сохранения в Excel число увеличится само.":
        "Row number: the first row gets this number, then +1. After saving to Excel the number increases automatically.",
    "страницы":
        "pages",
    "страниц":
        "pages",
    "Сначала загрузите словарь на вкладке «Распознавание».":
        "Load a dictionary on the “Recognition” tab first.",
    "Не удалось открыть PDF:\n{e}":
        "Could not open the PDF:\n{e}",
    "В выбранной папке нет PDF-файлов.":
        "The selected folder has no PDF files.",
    "Загрузка папки отменена.":
        "Folder loading cancelled.",
    "Не удалось склеить PDF из папки:\n{e}":
        "Could not merge the PDFs from the folder:\n{e}",
    "Не удалось открыть склеенный PDF:\n{e}":
        "Could not open the merged PDF:\n{e}",
    "Не удалось открыть Excel:\n{e}":
        "Could not open Excel:\n{e}",
    "Источник: {src}":
        "Source: {src}",
    "Все ячейки заполнены, проходят проверку, повторов нет.":
        "All cells are filled, pass the check, no duplicates.",
    "Проблем нет":
        "No problems",
    "Открыт документ строки {v}":
        "Opened the document of row {v}",
    "Поле заполняется вручную (вкладка «1. Общие данные»). Здесь можно задать значение только для этой строки.":
        "This field is entered manually (tab “1. Common data”). Here you can set a value for this row only.",
    "Файл занят:\n{path}\n\nЗакройте его в Excel и нажмите «В Excel» ещё раз.":
        "The file is busy:\n{path}\n\nClose it in Excel and click “To Excel” again.",
    "Не удалось сохранить:\n{e}":
        "Could not save:\n{e}",
    "_часть{num}":
        "_part{num}",
    "Строк: {n}, ожидается: {expected} ({mode})\nПродолжить?":
        "Rows: {n}, expected: {expected} ({mode})\nContinue?",
    "Внимание":
        "Attention",
    "Основной":
        "Main",
    "Укажите хотя бы одно поле, которое нужно взять с бланка.":
        "Enter at least one field to take from the form.",
    "Не удалось создать демо-PDF:\n{e}":
        "Could not create the demo PDF:\n{e}",
    "ошибка regex {label}: {e}":
        "regex error ({label}): {e}",
    "Ошибка в формуле поля «{nm}»:\n{formula}\n\n{e}":
        "Error in the formula of field «{nm}»:\n{formula}\n\n{e}",
    "В формуле поля «{nm}» поле [{ref}] не найдено.\nМожно ссылаться на ручные поля, поля из PDF и вычисляемые поля выше по списку.":
        "The formula of field «{nm}» uses [{ref}], which is not found.\nYou can refer to manual fields, PDF fields and calculated fields above in the list.",
    "Не удалось загрузить словарь:\n{e}":
        "Could not load the dictionary:\n{e}",
    "Не удалось дописать словарь:\n{e}":
        "Could not append to the dictionary:\n{e}",
    "Ошибка записи":
        "Write error",
    "Значение изменено только для этой строки":
        "Value changed for this row only",
    "PDF не открыт — скан недоступен.":
        "No PDF is open — the scan is not available.",
    "Распознанный текст зоны (стр. {v}):\n«{v2}»":
        "Recognized zone text (p. {v}):\n«{v2}»",
    "Склад":
        "Warehouse",
    "В названиях полей нельзя использовать [ и ]: ":
        "Field names cannot contain [ or ]: ",
    "нет страниц":
        "no pages",
    "Некорректный regex {label}:\n{e}":
        "Invalid {label} regex:\n{e}",
    "\nРезервная зона: «{v}»":
        "\nFallback zone: «{v}»",
    "Не удалось показать скан: {e}":
        "Could not show the scan: {e}",
    "Столбцы не совпадают":
        "Columns do not match",
    "В файле «{v}» другие столбцы:\n  в файле:   {v2}\n  в таблице: {v3}\n\nДописать нельзя — данные съедут по столбцам.":
        "The file «{v}» has different columns:\n  in the file:  {v2}\n  in the table: {v3}\n\nCannot append — the data would shift between columns.",
    "Некорректный regex {label} у поля «{v}»:\n{e}":
        "Invalid {label} regex in field «{v}»:\n{e}",
    "Сохранить в другой файл…":
        "Save to another file…",
    "Повторы в файле":
        "Duplicates in the file",
    "{n} из {n2} строк уже есть в файле (по полю «{key}»):\n{sample}{more}":
        "{n} of {n2} rows are already in the file (by field «{key}»):\n{sample}{more}",
    "Пропустить повторы":
        "Skip duplicates",
    "Сохранить все":
        "Save all",
    "— нет —":
        "— none —",
    "Число":
        "Number",
    "Число + OCR (О→0, З→3…)":
        "Number + OCR (O→0, l→1…)",
    "Только цифры":
        "Digits only",
    "Гос. номер РФ (OCR)":
        "Russian license plate (OCR)",
    "Гос. номер РФ до «/»":
        "Russian license plate before “/”",
    "ФИО → Фамилия И.О.":
        "Russian full name → Surname I.P.",
    "Нетто, кг → тонны":
        "Net, kg → tonnes",
    "Дата словами → ДД.ММ.ГГГГ":
        "Date in words → DD.MM.YYYY",
    "Свободный текст (без проверки)":
        "Free text (no check)",
    "Число — целое":
        "Number — integer",
    "Число — дробное":
        "Number — decimal",
    "Сумма (деньги)":
        "Amount (money)",
    "Количество":
        "Quantity",
    "Дата ДД.ММ.ГГГГ":
        "Date DD.MM.YYYY",
    "Дата ДД.ММ.ГГ":
        "Date DD.MM.YY",
    "Дата (. / -)":
        "Date (. / -)",
    "Номер документа":
        "Document number",
    "Артикул / код":
        "Article / code",
    "ИНН (10 или 12 цифр)":
        "INN (10 or 12 digits)",
    "КПП (9 цифр)":
        "KPP (9 digits)",
    "Штрихкод EAN-13":
        "Barcode EAN-13",
    "Штрихкод EAN-8":
        "Barcode EAN-8",
    "Гос. номер авто (РФ)":
        "Russian license plate",
    "Своё (регекс)":
        "Custom (regex)",
    "Текст как есть":
        "Text as is",
    "ООО «Ромашка»":
        "Acme Ltd.",
    "Без изменений и без проверки.":
        "No changes and no check.",
    "Число: вес, объём, сумма":
        "Number: weight, volume, amount",
    "Масса груза 44,1 т":
        "Gross weight 44.1 t",
    "Первое число в зоне. Буквы-двойники внутри числа исправляются: О→0, З→3.":
        "The first number in the zone. Look-alike letters inside the number are fixed: O→0, l→1.",
    "Номер без нулей впереди":
        "Number without leading zeros",
    "ОРДЕР № 00000012345":
        "ORDER No. 00000012345",
    "Цифры номера, нули в начале убираются.":
        "The digits of the number; leading zeros are removed.",
    "Номер из цифр (как есть)":
        "Number of digits (as is)",
    "Накладная № 004512":
        "Delivery note No. 004512",
    "Первая группа цифр, нули сохраняются.":
        "The first group of digits; zeros are kept.",
    "Номер с буквами (после «№»)":
        "Number with letters (after “No.”)",
    "Накладная № ТН-000123 от 05.03":
        "Invoice No. INV-000123 dated 05.03",
    "Берётся код после «№», «No» или «#»: буквы, цифры, «-» и «/».":
        "The code after “№”, “No” or “#”: letters, digits, “-” and “/”.",
    "Дата 05.03.2025":
        "Date 05.03.2025",
    "от 05.03.2025 г.":
        "dated 05.03.2025",
    "Дата в формате ДД.ММ.ГГГГ.":
        "Date in DD.MM.YYYY format.",
    "Дата «5 марта 2025» → 05.03.2025":
        "Date “5 March 2025” → 05.03.2025",
    "«5» марта 2025 г.":
        "5 March 2025",
    "Месяц словом переводится в число.":
        "The month name is turned into a number (English and Russian).",
    "Гос. номер РФ":
        "Russian license plate",
    "Латиница → кириллица, О↔0 и З↔3 по позициям.":
        "Latin → Cyrillic, O↔0 and З↔3 by position.",
    "Гос. номер тягача (до «/»)":
        "Truck plate (before “/”)",
    "Т123ОР178 / АВ1234 78":
        "Т123ОР178 / АВ1234 78",
    "Берётся номер до «/», номер прицепа отбрасывается.":
        "Takes the plate before “/”, the trailer plate is dropped.",
    "Иванов Иван Петрович, паспорт 4000":
        "Иванов Иван Петрович, паспорт 4000",
    "Сокращает ФИО, отрезает паспорт, телефон и мусор.":
        "Shortens a Russian full name, cuts off passport, phone and noise.",
    "Нетто: 51 470 кг":
        "Net weight: 51470 kg",
    "Килограммы делятся на 1000.":
        "Kilograms are divided by 1000.",
    "ИНН":
        "INN (Russian tax ID)",
    "ИНН 7701234567":
        "INN 7701234567",
    "10 или 12 цифр.":
        "10 or 12 digits.",
}


# ─────────────────────────────────────────────────────────────────────────────
# Точка входа
# ─────────────────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    # Применяем сохранённую тему (или тему по умолчанию) до создания окна.
    saved = load_ui_settings().get("theme", _DEFAULT_THEME)
    if saved not in THEMES:
        saved = _DEFAULT_THEME
    _apply_theme_tokens(THEMES[saved])
    app.setStyleSheet(build_qss())
    app.setApplicationName("ZoneExtract")
    app.setWindowIcon(QIcon(app_icon_pixmap(64)))
    app.setApplicationVersion(VERSION)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
