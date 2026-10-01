#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mosque Management Professional - PyQt6
Version 4.0 Enhanced, en-US localized edition.

Highlights:
- First-release current database schema with user-created first administrator account.
- PBKDF2-SHA256 password storage.
- Consistent admin/operator/viewer permissions.
- Dashboard, date filters, search, pagination, and cash-flow chart.
- Recurring transactions, activity reminders, and asset maintenance schedules.
- Online SQLite backup, integrity verification, safe restore, and backup retention.
- CSV/Excel/PDF import and export restricted to approved tables.
- Mosque profile, permanent dark interface, activity log, and password management.
- en-US text/date localization and configurable display currency formatting.

Required dependency:
    pip install PyQt6
Optional dependencies:
    pip install matplotlib openpyxl reportlab
"""

from __future__ import annotations

import calendar
import configparser
import csv
import datetime as dt
import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import sqlite3
import sys
import traceback
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from PyQt6.QtCore import (
        QByteArray, QDate, QDateTime, QLocale, QPoint, QPointF, QRectF,
        QTimer, Qt, QUrl, pyqtSignal,
    )
    from PyQt6.QtGui import (
        QAction, QBrush, QColor, QDesktopServices, QFont, QIcon, QPainter,
        QPen, QPixmap, QPalette,
    )
    from PyQt6.QtWidgets import (
        QAbstractItemView,
        QApplication,
        QCheckBox,
        QComboBox,
        QDateEdit,
        QDateTimeEdit,
        QDialog,
        QDialogButtonBox,
        QDoubleSpinBox,
        QFileDialog,
        QFormLayout,
        QFrame,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QListWidget,
        QMainWindow,
        QMenu,
        QMessageBox,
        QPlainTextEdit,
        QProgressBar,
        QPushButton,
        QSpinBox,
        QSplitter,
        QStatusBar,
        QTabWidget,
        QTableWidget,
        QTableWidgetItem,
        QTextEdit,
        QToolBar,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:  # pragma: no cover - user-facing dependency message
    raise SystemExit(
        "PyQt6 is not installed. Run: pip install PyQt6\n"
        f"Details: {exc}"
    )

try:
    import matplotlib

    matplotlib.use("QtAgg")
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure

    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import (
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


APP_NAME = "Mosque Management Professional"
APP_VERSION = "4.0"
DATABASE_SCHEMA_VERSION = "4.0"
ICON_FILENAMES = ("mosque_icon.ico", "mosque_icon.png")
DATE_FMT = "%Y-%m-%d"  # database storage format; do not localize
DATETIME_FMT = "%Y-%m-%d %H:%M:%S"  # database storage format
APP_LOCALE_NAME = "en_US"
DATE_DISPLAY_FMT = "MMM d, yyyy"
DATETIME_DISPLAY_FMT = "MMM d, yyyy h:mm AP"
CURRENCY_PROFILES = {
    "USD": {"symbol": "$", "decimals": 2, "step": 100.0},
    "IDR": {"symbol": "Rp", "decimals": 0, "step": 50_000.0},
    "EUR": {"symbol": "€", "decimals": 2, "step": 100.0},
    "GBP": {"symbol": "£", "decimals": 2, "step": 100.0},
    "JPY": {"symbol": "¥", "decimals": 0, "step": 10_000.0},
    "SGD": {"symbol": "S$", "decimals": 2, "step": 100.0},
    "MYR": {"symbol": "RM", "decimals": 2, "step": 100.0},
}
_CURRENT_CURRENCY = "USD"
_APP_LOCALE = QLocale(APP_LOCALE_NAME)
PASSWORD_MAX_LENGTH = 6
PASSWORD_PATTERN = re.compile(r"[A-Za-z0-9]{1,6}")
USERNAME_PATTERN = re.compile(r"[A-Za-z0-9_.-]{3,40}")


def validate_username(username: str) -> str:
    value = (username or "").strip()
    if not USERNAME_PATTERN.fullmatch(value):
        raise ValueError(
            "Username must be 3–40 characters and may contain letters, numbers, "
            "periods, underscores, or hyphens."
        )
    return value


def validate_password(password: str) -> str:
    value = password or ""
    if not PASSWORD_PATTERN.fullmatch(value):
        raise ValueError(
            "Password must contain 1–6 characters using letters or numbers only."
        )
    return value


ENUM_LABELS = {
    "transaction_type": {"semua": "All", "pemasukan": "Income", "pengeluaran": "Expense"},
    "status_filter": {"semua": "All", "aktif": "Active", "nonaktif": "Inactive"},
    "donor_category": {
        "individu": "Individual", "keluarga": "Family", "perusahaan": "Company",
        "yayasan": "Foundation", "komunitas": "Community",
    },
    "activity_status": {
        "semua": "All", "rencana": "Planned", "berjalan": "In Progress",
        "selesai": "Completed", "batal": "Canceled",
    },
    "priority": {"rendah": "Low", "normal": "Normal", "tinggi": "High", "mendesak": "Urgent"},
    "asset_category": {
        "semua": "All", "bangunan": "Building", "tanah": "Land", "kendaraan": "Vehicle",
        "elektronik": "Electronics", "peralatan": "Equipment", "furnitur": "Furniture", "lainnya": "Other",
    },
    "asset_condition": {
        "semua": "All", "baik": "Good", "perlu_perawatan": "Needs Maintenance",
        "rusak_ringan": "Minor Damage", "rusak_berat": "Major Damage", "hilang": "Missing",
    },
    "interval_unit": {"hari": "Day(s)", "minggu": "Week(s)", "bulan": "Month(s)", "tahun": "Year(s)"},
}

def resource_path(filename: str) -> str:
    """Return a resource path that also works in PyInstaller bundles."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base / filename)


def create_builtin_mosque_icon(size: int = 256) -> QIcon:
    """Create a scalable mosque icon when no external icon file is present."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    scale = size / 256.0

    def rect(x: float, y: float, width: float, height: float) -> QRectF:
        return QRectF(x * scale, y * scale, width * scale, height * scale)

    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QBrush(QColor("#176b52")))
    painter.drawRoundedRect(rect(6, 6, 244, 244), 48 * scale, 48 * scale)

    # Crescent moon.
    painter.setBrush(QBrush(QColor("#f6c453")))
    painter.drawEllipse(rect(178, 28, 42, 42))
    painter.setBrush(QBrush(QColor("#176b52")))
    painter.drawEllipse(rect(188, 23, 42, 42))

    painter.setBrush(QBrush(QColor("#ffffff")))
    # Main prayer hall and central dome.
    painter.drawRect(rect(55, 132, 146, 66))
    painter.drawPie(rect(76, 76, 104, 104), 0, 180 * 16)
    painter.drawRect(rect(124, 65, 8, 24))
    painter.drawEllipse(rect(121, 56, 14, 14))

    # Minarets.
    painter.drawRect(rect(31, 98, 24, 100))
    painter.drawPie(rect(28, 75, 30, 38), 0, 180 * 16)
    painter.drawRect(rect(201, 98, 24, 100))
    painter.drawPie(rect(198, 75, 30, 38), 0, 180 * 16)

    # Doors and base line.
    painter.setBrush(QBrush(QColor("#176b52")))
    painter.drawPie(rect(112, 148, 32, 50), 0, 180 * 16)
    painter.drawRect(rect(112, 173, 32, 25))
    painter.setPen(QPen(QColor("#f6c453"), max(2.0, 7.0 * scale)))
    painter.drawLine(QPointF(25 * scale, 207 * scale), QPointF(231 * scale, 207 * scale))

    painter.end()
    return QIcon(pixmap)


def load_mosque_icon() -> QIcon:
    """Load the bundled mosque icon, with a built-in vector fallback."""
    for filename in ICON_FILENAMES:
        candidate = resource_path(filename)
        if os.path.isfile(candidate):
            icon = QIcon(candidate)
            if not icon.isNull():
                return icon
    return create_builtin_mosque_icon()


def configure_localization(settings: Optional["SettingsManager"] = None) -> None:
    """Configure the en-US locale and selected display currency.

    Currency selection changes formatting only; stored amounts are never converted.
    """
    global _CURRENT_CURRENCY, _APP_LOCALE
    _APP_LOCALE = QLocale(APP_LOCALE_NAME)
    QLocale.setDefault(_APP_LOCALE)
    requested = settings.get("Localization", "currency", "USD") if settings else "USD"
    _CURRENT_CURRENCY = requested if requested in CURRENCY_PROFILES else "USD"

def enum_label(group: str, value: Any) -> str:
    text = str(value or "")
    return ENUM_LABELS.get(group, {}).get(text, text.replace("_", " ").title())

def add_enum_items(combo: QComboBox, group: str, values: Sequence[str]) -> None:
    for value in values:
        combo.addItem(enum_label(group, value), value)

def set_combo_data(combo: QComboBox, value: Any) -> None:
    index = combo.findData(value)
    if index >= 0:
        combo.setCurrentIndex(index)
    elif combo.isEditable():
        combo.setEditText(str(value or ""))

def combo_data_or_text(combo: QComboBox) -> str:
    text = combo.currentText().strip()
    index = combo.currentIndex()
    if combo.isEditable() and (index < 0 or text != combo.itemText(index)):
        return text
    data = combo.currentData()
    return str(data) if data is not None else text

def fmt_date(value: Any) -> str:
    text = str(value or "").strip()[:10]
    date = QDate.fromString(text, "yyyy-MM-dd")
    return _APP_LOCALE.toString(date, DATE_DISPLAY_FMT) if date.isValid() else text

def fmt_datetime(value: Any) -> str:
    text = str(value or "").strip()[:19]
    parsed = QDateTime.fromString(text, "yyyy-MM-dd HH:mm:ss")
    return _APP_LOCALE.toString(parsed, DATETIME_DISPLAY_FMT) if parsed.isValid() else text
ALLOWED_DATA_TABLES = {
    "pengurus",
    "donatur",
    "transaksi",
    "kegiatan",
    "aset",
    "transaksi_rutin",
}
CURRENT_SCHEMA_COLUMNS = {
    "pengurus": {"id", "nama", "jabatan", "no_hp", "alamat", "aktif", "tanggal_bergabung", "email", "catatan"},
    "donatur": {"id", "nama", "no_hp", "alamat", "email", "kategori", "aktif", "tanggal_daftar", "catatan"},
    "transaksi": {"id", "tanggal", "jenis", "kategori", "deskripsi", "jumlah", "donatur_id", "bukti", "created_by", "updated_at"},
    "kegiatan": {"id", "nama", "tanggal_mulai", "tanggal_selesai", "deskripsi", "penanggung_jawab", "anggaran", "realisasi", "status", "lokasi", "prioritas", "pengingat_hari"},
    "aset": {"id", "nama", "kategori", "nilai", "tahun_perolehan", "kondisi", "keterangan", "lokasi", "tanggal_perolehan", "tanggal_perawatan_terakhir", "tanggal_perawatan_berikut", "aktif"},
    "transaksi_rutin": {"id", "nama", "jenis", "kategori", "deskripsi", "jumlah", "donatur_id", "interval_nilai", "interval_unit", "tanggal_berikut", "aktif", "terakhir_diproses"},
    "users": {"username", "password", "role", "nama_lengkap", "aktif", "dibuat_pada", "login_terakhir", "wajib_ganti_password"},
    "log_aktivitas": {"id", "waktu", "user", "aksi", "detail"},
    "id_counter": {"entity_type", "counter"},
    "app_meta": {"key", "value"},
}
ROLE_PERMISSIONS = {
    "admin": {"view", "add", "edit", "delete", "export", "import", "backup", "restore", "users", "settings"},
    "operator": {"view", "add", "edit", "export", "backup"},
    "user": {"view", "add", "edit", "export", "backup"},  # kompatibilitas role lama
    "viewer": {"view", "export"},
}


DARK_STYLESHEET = """
    QMainWindow, QDialog, QWidget {
        background: #20252b;
        color: #eef2f5;
    }
    QMenuBar, QMenu, QToolBar, QStatusBar {
        background: #293038;
        color: #eef2f5;
    }
    QMenuBar::item, QMenu::item {
        background: transparent;
        padding: 6px 10px;
    }
    QMenuBar::item:selected, QMenu::item:selected {
        background: #3b4752;
    }
    QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QSpinBox,
    QDoubleSpinBox, QDateEdit, QDateTimeEdit, QListWidget, QTableWidget {
        background: #2d353d;
        color: #f7fafc;
        border: 1px solid #52606d;
        border-radius: 4px;
        padding: 4px;
        selection-background-color: #2d8f68;
        selection-color: #ffffff;
    }
    QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled,
    QComboBox:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled,
    QDateEdit:disabled, QDateTimeEdit:disabled {
        background: #272d33;
        color: #7f8a94;
    }
    QComboBox QAbstractItemView {
        background: #2d353d;
        color: #f7fafc;
        selection-background-color: #2d8f68;
    }
    QHeaderView::section {
        background: #36414b;
        color: #ffffff;
        padding: 6px;
        border: 0;
        border-right: 1px solid #52606d;
        border-bottom: 1px solid #52606d;
    }
    QTableWidget {
        gridline-color: #46515b;
        alternate-background-color: #252c32;
    }
    QPushButton {
        background: #36414b;
        color: #ffffff;
        border: 1px solid #607080;
        border-radius: 5px;
        padding: 6px 11px;
    }
    QPushButton:hover { background: #465563; }
    QPushButton:pressed { background: #2d8f68; }
    QPushButton:disabled {
        color: #7c8792;
        background: #2a3036;
        border-color: #424b54;
    }
    QGroupBox {
        border: 1px solid #465563;
        border-radius: 7px;
        margin-top: 8px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 5px;
    }
    QTabWidget::pane { border: 1px solid #465563; }
    QTabBar::tab {
        background: #293038;
        color: #dce3e8;
        padding: 8px 14px;
        border: 1px solid #3e4852;
        border-bottom: 0;
    }
    QTabBar::tab:selected {
        background: #3b4752;
        color: #ffffff;
    }
    QTabBar::tab:hover { background: #333d46; }
    QToolTip {
        background: #11161a;
        color: #ffffff;
        border: 1px solid #66717c;
    }
    QScrollBar:vertical {
        background: #252c32;
        width: 12px;
        margin: 0;
    }
    QScrollBar::handle:vertical {
        background: #52606d;
        min-height: 24px;
        border-radius: 6px;
    }
    QScrollBar:horizontal {
        background: #252c32;
        height: 12px;
        margin: 0;
    }
    QScrollBar::handle:horizontal {
        background: #52606d;
        min-width: 24px;
        border-radius: 6px;
    }
    QScrollBar::add-line, QScrollBar::sub-line {
        width: 0;
        height: 0;
    }
    QProgressBar {
        background: #252c32;
        color: #ffffff;
        border: 1px solid #52606d;
        border-radius: 4px;
        text-align: center;
    }
    QProgressBar::chunk {
        background: #2d8f68;
        border-radius: 3px;
    }
"""


def apply_dark_theme(app: QApplication) -> None:
    """Terapkan antarmuka gelap permanen ke seluruh aplikasi, termasuk dialog login."""
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor("#20252b"))
    palette.setColor(QPalette.ColorRole.WindowText, QColor("#eef2f5"))
    palette.setColor(QPalette.ColorRole.Base, QColor("#2d353d"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#252c32"))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#11161a"))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#f7fafc"))
    palette.setColor(QPalette.ColorRole.Button, QColor("#36414b"))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.BrightText, QColor("#ff6b6b"))
    palette.setColor(QPalette.ColorRole.Link, QColor("#63b3ed"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#2d8f68"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    app.setPalette(palette)
    app.setStyleSheet(DARK_STYLESHEET)


def now_iso() -> str:
    return dt.datetime.now().replace(microsecond=0).isoformat(sep=" ")


def today_iso() -> str:
    return dt.date.today().isoformat()


def fmt_currency(value: Any) -> str:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        number = 0.0
    profile = CURRENCY_PROFILES[_CURRENT_CURRENCY]
    decimals = int(profile["decimals"])
    formatted = _APP_LOCALE.toString(abs(number), "f", decimals)
    sign = "-" if number < 0 else ""
    return f"{sign}{profile['symbol']}{formatted}"


def safe_float(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    original = str(value).strip()
    legacy_idr = "Rp" in original or "IDR" in original.upper()
    text = original
    for token in ("USD", "IDR", "EUR", "GBP", "JPY", "SGD", "MYR", "Rp", "S$", "RM", "$", "€", "£", "¥"):
        text = text.replace(token, "")
    text = text.replace(" ", "")
    if not text:
        return default
    if "," in text and "." in text:
        if text.rfind(".") > text.rfind(","):
            text = text.replace(",", "")  # en-US: 1,234.56
        else:
            text = text.replace(".", "").replace(",", ".")  # id-ID: 1.234,56
    elif "," in text:
        tail = text.rsplit(",", 1)[-1]
        text = text.replace(",", ".") if len(tail) in {1, 2} else text.replace(",", "")
    elif text.count(".") > 1 or (legacy_idr and "." in text):
        text = text.replace(".", "")
    try:
        return float(text)
    except ValueError:
        return default

def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def normalize_phone(value: str) -> str:
    text = re.sub(r"[^0-9+]", "", (value or "").strip())
    if text.startswith("+62"):
        return "0" + text[3:]
    if text.startswith("62") and not text.startswith("0"):
        return "0" + text[2:]
    return text


def valid_email(value: str) -> bool:
    if not value:
        return True
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", value.strip()))


def month_bounds(year: int, month: int) -> Tuple[str, str]:
    last_day = calendar.monthrange(year, month)[1]
    return dt.date(year, month, 1).isoformat(), dt.date(year, month, last_day).isoformat()


def shift_month(date_value: dt.date, months: int) -> dt.date:
    month_index = date_value.year * 12 + date_value.month - 1 + months
    year, month_zero = divmod(month_index, 12)
    month = month_zero + 1
    day = min(date_value.day, calendar.monthrange(year, month)[1])
    return dt.date(year, month, day)


def parse_date(value: Any, fallback: Optional[dt.date] = None) -> dt.date:
    fallback = fallback or dt.date.today()
    if isinstance(value, dt.date):
        return value
    text = str(value or "").strip()[:10]
    try:
        return dt.date.fromisoformat(text)
    except ValueError:
        return fallback


def elide(text: Any, length: int = 45) -> str:
    value = str(text or "")
    return value if len(value) <= length else value[: length - 1] + "…"


def unique_path(path: str) -> str:
    candidate = Path(path)
    if not candidate.exists():
        return str(candidate)
    stem, suffix = candidate.stem, candidate.suffix
    for i in range(1, 10000):
        other = candidate.with_name(f"{stem}_{i}{suffix}")
        if not other.exists():
            return str(other)
    raise RuntimeError('Unable to create a unique file name.')


class PasswordHasher:
    """PBKDF2 password storage dengan dukungan verifikasi hash SHA256 lama."""

    ALGORITHM = "pbkdf2_sha256"
    ITERATIONS = 260_000

    @classmethod
    def hash(cls, password: str) -> str:
        password = validate_password(password)
        salt = secrets.token_hex(16)
        digest = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), bytes.fromhex(salt), cls.ITERATIONS
        ).hex()
        return f"{cls.ALGORITHM}${cls.ITERATIONS}${salt}${digest}"

    @classmethod
    def verify(cls, password: str, stored: str) -> Tuple[bool, bool]:
        """Return (valid, needs_rehash)."""
        stored = stored or ""
        if stored.startswith(f"{cls.ALGORITHM}$"):
            try:
                _, iterations, salt, expected = stored.split("$", 3)
                digest = hashlib.pbkdf2_hmac(
                    "sha256",
                    password.encode("utf-8"),
                    bytes.fromhex(salt),
                    int(iterations),
                ).hex()
                valid = hmac.compare_digest(digest, expected)
                return valid, valid and int(iterations) < cls.ITERATIONS
            except (ValueError, TypeError):
                return False, False
        return False, False


class SettingsManager:
    def __init__(self, config_file: str = "mosque-settings.ini"):
        self.config_file = os.path.abspath(config_file)
        self.config = configparser.ConfigParser()
        self.defaults = {
            "Window": {
                "width": "1280",
                "height": "760",
                "x": "",
                "y": "",
                "state": "normal",
                "layout_state": "",
            },
            "Database": {
                "path": "mosque-management.db",
                "auto_backup": "True",
                "backup_interval_days": "1",
                "last_backup_date": "",
            },
            "Backup": {"backup_dir": "backup", "max_backups": "15"},
            "User": {"last_username": "", "remember_username": "False"},
            "Masjid": {
                "nama": "Mosque",
                "alamat": "",
                "telepon": "",
                "email": "",
                "ketua_dkm": "",
            },
            "Security": {"session_timeout_minutes": "0"},
            "Localization": {"locale": "en_US", "currency": "USD"},
        }
        self.load()

    def load(self) -> None:
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding="utf-8")
        self.set_defaults()
        # Remove obsolete appearance settings because the application uses one permanent theme.
        if self.config.has_section("Appearance"):
            self.config.remove_section("Appearance")
        self.save()

    def set_defaults(self) -> None:
        for section, values in self.defaults.items():
            if not self.config.has_section(section):
                self.config.add_section(section)
            for key, value in values.items():
                if not self.config.has_option(section, key):
                    self.config.set(section, key, value)

    def save(self) -> None:
        Path(self.config_file).parent.mkdir(parents=True, exist_ok=True)
        with open(self.config_file, "w", encoding="utf-8") as handle:
            self.config.write(handle)

    def get(self, section: str, key: str, fallback: Any = None) -> Any:
        try:
            return self.config.get(section, key, fallback=fallback)
        except (configparser.Error, ValueError):
            return fallback

    def get_bool(self, section: str, key: str, fallback: bool = False) -> bool:
        try:
            return self.config.getboolean(section, key, fallback=fallback)
        except (configparser.Error, ValueError):
            return fallback

    def get_int(self, section: str, key: str, fallback: int = 0) -> int:
        try:
            return self.config.getint(section, key, fallback=fallback)
        except (configparser.Error, ValueError):
            return fallback

    def set(self, section: str, key: str, value: Any, save: bool = True) -> None:
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))
        if save:
            self.save()

    def set_many(self, values: Dict[str, Dict[str, Any]]) -> None:
        for section, items in values.items():
            for key, value in items.items():
                self.set(section, key, value, save=False)
        self.save()

    def get_db_path(self) -> str:
        path = self.get("Database", "path", "mosque-management.db")
        if not os.path.isabs(path):
            path = os.path.join(os.path.dirname(self.config_file), path)
        return os.path.abspath(path)

    def get_backup_dir(self) -> str:
        path = self.get("Backup", "backup_dir", "backup")
        if not os.path.isabs(path):
            path = os.path.join(os.path.dirname(self.config_file), path)
        return os.path.abspath(path)

    def window_geometry(self) -> Tuple[int, int, Optional[int], Optional[int], str]:
        width = max(1000, self.get_int("Window", "width", 1280))
        height = max(650, self.get_int("Window", "height", 760))
        x_text = self.get("Window", "x", "")
        y_text = self.get("Window", "y", "")
        x = safe_int(x_text) if str(x_text).strip() else None
        y = safe_int(y_text) if str(y_text).strip() else None
        return width, height, x, y, self.get("Window", "state", "normal")


class Database:
    def __init__(self, db_path: str):
        self.db_path = os.path.abspath(db_path)
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        self.connection: sqlite3.Connection
        self.connect()
        self.create_schema()
        self.initialize_schema_metadata()
        self.validate_current_schema()
        self.seed_defaults()

    def connect(self) -> None:
        self.connection = sqlite3.connect(self.db_path, timeout=20)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.execute("PRAGMA journal_mode = WAL")
        self.connection.execute("PRAGMA synchronous = NORMAL")
        self.connection.execute("PRAGMA busy_timeout = 10000")

    def close(self) -> None:
        if getattr(self, "connection", None):
            self.connection.close()

    @contextmanager
    def transaction(self):
        try:
            yield self.connection
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            raise

    def execute(self, query: str, params: Sequence[Any] = ()) -> sqlite3.Cursor:
        try:
            with self.transaction() as connection:
                return connection.execute(query, tuple(params))
        except sqlite3.Error as exc:
            raise RuntimeError(f"Database error: {exc}") from exc

    def executemany(self, query: str, rows: Iterable[Sequence[Any]]) -> None:
        try:
            with self.transaction() as connection:
                connection.executemany(query, rows)
        except sqlite3.Error as exc:
            raise RuntimeError(f"Database error: {exc}") from exc

    def fetch_one(self, query: str, params: Sequence[Any] = ()) -> Optional[sqlite3.Row]:
        return self.connection.execute(query, tuple(params)).fetchone()

    def fetch_all(self, query: str, params: Sequence[Any] = ()) -> List[sqlite3.Row]:
        return self.connection.execute(query, tuple(params)).fetchall()

    def scalar(self, query: str, params: Sequence[Any] = (), default: Any = 0) -> Any:
        row = self.fetch_one(query, params)
        return row[0] if row is not None else default

    def create_schema(self) -> None:
        with self.transaction() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS pengurus (
                    id TEXT PRIMARY KEY,
                    nama TEXT NOT NULL,
                    jabatan TEXT NOT NULL,
                    no_hp TEXT NOT NULL,
                    alamat TEXT NOT NULL,
                    aktif INTEGER NOT NULL DEFAULT 1,
                    tanggal_bergabung TEXT NOT NULL,
                    email TEXT DEFAULT '',
                    catatan TEXT DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS donatur (
                    id TEXT PRIMARY KEY,
                    nama TEXT NOT NULL,
                    no_hp TEXT NOT NULL,
                    alamat TEXT NOT NULL,
                    email TEXT DEFAULT '',
                    kategori TEXT NOT NULL DEFAULT 'individu',
                    aktif INTEGER NOT NULL DEFAULT 1,
                    tanggal_daftar TEXT NOT NULL DEFAULT '',
                    catatan TEXT DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS transaksi (
                    id TEXT PRIMARY KEY,
                    tanggal TEXT NOT NULL,
                    jenis TEXT NOT NULL CHECK(jenis IN ('pemasukan','pengeluaran')),
                    kategori TEXT NOT NULL,
                    deskripsi TEXT NOT NULL,
                    jumlah REAL NOT NULL CHECK(jumlah >= 0),
                    donatur_id TEXT,
                    bukti TEXT,
                    created_by TEXT DEFAULT '',
                    updated_at TEXT DEFAULT '',
                    FOREIGN KEY (donatur_id) REFERENCES donatur(id) ON DELETE SET NULL
                );

                CREATE TABLE IF NOT EXISTS kegiatan (
                    id TEXT PRIMARY KEY,
                    nama TEXT NOT NULL,
                    tanggal_mulai TEXT NOT NULL,
                    tanggal_selesai TEXT NOT NULL,
                    deskripsi TEXT NOT NULL,
                    penanggung_jawab TEXT NOT NULL,
                    anggaran REAL NOT NULL DEFAULT 0,
                    realisasi REAL NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'rencana',
                    lokasi TEXT DEFAULT '',
                    prioritas TEXT DEFAULT 'normal',
                    pengingat_hari INTEGER NOT NULL DEFAULT 3
                );

                CREATE TABLE IF NOT EXISTS aset (
                    id TEXT PRIMARY KEY,
                    nama TEXT NOT NULL,
                    kategori TEXT NOT NULL,
                    nilai REAL NOT NULL DEFAULT 0,
                    tahun_perolehan INTEGER NOT NULL,
                    kondisi TEXT NOT NULL,
                    keterangan TEXT DEFAULT '',
                    lokasi TEXT DEFAULT '',
                    tanggal_perolehan TEXT DEFAULT '',
                    tanggal_perawatan_terakhir TEXT DEFAULT '',
                    tanggal_perawatan_berikut TEXT DEFAULT '',
                    aktif INTEGER NOT NULL DEFAULT 1
                );

                CREATE TABLE IF NOT EXISTS transaksi_rutin (
                    id TEXT PRIMARY KEY,
                    nama TEXT NOT NULL,
                    jenis TEXT NOT NULL CHECK(jenis IN ('pemasukan','pengeluaran')),
                    kategori TEXT NOT NULL,
                    deskripsi TEXT NOT NULL,
                    jumlah REAL NOT NULL CHECK(jumlah > 0),
                    donatur_id TEXT,
                    interval_nilai INTEGER NOT NULL DEFAULT 1,
                    interval_unit TEXT NOT NULL DEFAULT 'bulan',
                    tanggal_berikut TEXT NOT NULL,
                    aktif INTEGER NOT NULL DEFAULT 1,
                    terakhir_diproses TEXT DEFAULT '',
                    FOREIGN KEY (donatur_id) REFERENCES donatur(id) ON DELETE SET NULL
                );

                CREATE TABLE IF NOT EXISTS users (
                    username TEXT PRIMARY KEY,
                    password TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'viewer',
                    nama_lengkap TEXT DEFAULT '',
                    aktif INTEGER NOT NULL DEFAULT 1,
                    dibuat_pada TEXT DEFAULT '',
                    login_terakhir TEXT DEFAULT '',
                    wajib_ganti_password INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS log_aktivitas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    waktu TEXT NOT NULL,
                    user TEXT NOT NULL,
                    aksi TEXT NOT NULL,
                    detail TEXT DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS id_counter (
                    entity_type TEXT PRIMARY KEY,
                    counter INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS app_meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_transaksi_tanggal ON transaksi(tanggal);
                CREATE INDEX IF NOT EXISTS idx_transaksi_jenis ON transaksi(jenis);
                CREATE INDEX IF NOT EXISTS idx_transaksi_kategori ON transaksi(kategori);
                CREATE INDEX IF NOT EXISTS idx_kegiatan_tanggal ON kegiatan(tanggal_mulai);
                CREATE INDEX IF NOT EXISTS idx_aset_perawatan ON aset(tanggal_perawatan_berikut);
                CREATE INDEX IF NOT EXISTS idx_log_waktu ON log_aktivitas(waktu);
                """
            )

    def initialize_schema_metadata(self) -> None:
        """Mark this database as the current first-release schema baseline."""
        row = self.fetch_one("SELECT value FROM app_meta WHERE key='schema_version'")
        if row is not None and str(row[0]) != DATABASE_SCHEMA_VERSION:
            raise RuntimeError(
                "This database uses an unsupported schema version "
                f"({row[0]}). Expected {DATABASE_SCHEMA_VERSION}. "
                "Automatic migration is intentionally disabled for this baseline release."
            )
        self.execute(
            "INSERT OR REPLACE INTO app_meta(key, value) VALUES('schema_version', ?)",
            (DATABASE_SCHEMA_VERSION,),
        )
        self.execute(
            "INSERT OR REPLACE INTO app_meta(key, value) VALUES('edition', ?)",
            ("en-US",),
        )

    def validate_current_schema(self) -> None:
        """Reject incomplete/older schemas instead of silently migrating them."""
        problems: List[str] = []
        for table, required_columns in CURRENT_SCHEMA_COLUMNS.items():
            table_exists = self.fetch_one(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
                (table,),
            )
            if table_exists is None:
                problems.append(f"missing table: {table}")
                continue
            actual_columns = {
                str(row[1]) for row in self.fetch_all(f"PRAGMA table_info({table})")
            }
            missing = sorted(required_columns - actual_columns)
            if missing:
                problems.append(f"{table} missing columns: {', '.join(missing)}")
        if problems:
            raise RuntimeError(
                "The selected database is not compatible with the current baseline schema:\n- "
                + "\n- ".join(problems)
            )

    def seed_defaults(self) -> None:
        entities = {
            "pengurus": "P",
            "donatur": "D",
            "transaksi": "T",
            "kegiatan": "K",
            "aset": "A",
            "transaksi_rutin": "R",
        }
        for entity in entities:
            self.execute(
                "INSERT OR IGNORE INTO id_counter(entity_type, counter) VALUES(?, 0)",
                (entity,),
            )
            max_id = self.scalar(
                f"SELECT MAX(CAST(SUBSTR(id, 2) AS INTEGER)) FROM {entity}", default=0
            ) or 0
            self.execute(
                "UPDATE id_counter SET counter=MAX(counter, ?) WHERE entity_type=?",
                (int(max_id), entity),
            )

    def user_count(self) -> int:
        return int(self.scalar("SELECT COUNT(*) FROM users", default=0))

    def create_initial_admin(self, username: str, password: str) -> None:
        """Create the administrator selected by the user on the first launch."""
        if self.user_count() != 0:
            raise RuntimeError("Initial setup is only available when no users exist.")
        username = validate_username(username)
        password = validate_password(password)
        self.execute(
            """INSERT INTO users
               (username,password,role,nama_lengkap,aktif,dibuat_pada,wajib_ganti_password)
               VALUES(?,?,'admin','Administrator',1,?,0)""",
            (username, PasswordHasher.hash(password), now_iso()),
        )
        self.log(username, "Initial setup", "Primary administrator account created")

    def next_id(self, entity: str) -> str:
        prefixes = {
            "pengurus": "P",
            "donatur": "D",
            "transaksi": "T",
            "kegiatan": "K",
            "aset": "A",
            "transaksi_rutin": "R",
        }
        if entity not in prefixes:
            raise ValueError('Unknown ID entity.')
        with self.transaction() as connection:
            connection.execute(
                "UPDATE id_counter SET counter=counter+1 WHERE entity_type=?", (entity,)
            )
            counter = connection.execute(
                "SELECT counter FROM id_counter WHERE entity_type=?", (entity,)
            ).fetchone()[0]
        return f"{prefixes[entity]}{counter:05d}"

    def log(self, user: str, action: str, detail: str = "") -> None:
        self.execute(
            "INSERT INTO log_aktivitas(waktu,user,aksi,detail) VALUES(?,?,?,?)",
            (now_iso(), user or "system", action, detail[:1000]),
        )

    def integrity_check(self) -> Tuple[bool, str]:
        result = self.scalar("PRAGMA integrity_check", default="unknown")
        return result == "ok", str(result)

    def backup(self, backup_dir: str, label: str = "backup") -> str:
        Path(backup_dir).mkdir(parents=True, exist_ok=True)
        timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = unique_path(os.path.join(backup_dir, f"{label}_{timestamp}.db"))
        destination = sqlite3.connect(path)
        try:
            self.connection.backup(destination)
            check = destination.execute("PRAGMA integrity_check").fetchone()[0]
            if check != "ok":
                raise RuntimeError(f"Backup verification failed: {check}")
        finally:
            destination.close()
        return path

    @staticmethod
    def verify_database(path: str) -> Tuple[bool, str]:
        if not os.path.isfile(path):
            return False, 'File not found.'
        connection = None
        try:
            connection = sqlite3.connect(f"file:{Path(path).as_posix()}?mode=ro", uri=True)
            result = connection.execute("PRAGMA integrity_check").fetchone()[0]
            return result == "ok", str(result)
        except sqlite3.Error as exc:
            return False, str(exc)
        finally:
            if connection:
                connection.close()

    def restore(self, source: str, backup_dir: str) -> str:
        ok, message = self.verify_database(source)
        if not ok:
            raise RuntimeError(f"The restore file is invalid: {message}")
        safety_backup = self.backup(backup_dir, "before_restore")
        self.close()
        try:
            shutil.copy2(source, self.db_path)
            self.connect()
            ok, message = self.integrity_check()
            if not ok:
                raise RuntimeError(message)
            self.create_schema()
            self.initialize_schema_metadata()
            self.validate_current_schema()
            self.seed_defaults()
        except Exception:
            if getattr(self, "connection", None):
                try:
                    self.close()
                except Exception:
                    pass
            shutil.copy2(safety_backup, self.db_path)
            self.connect()
            raise
        return safety_backup


class MasjidManager:
    def __init__(self, db: Database):
        self.db = db
        self.current_user = "system"
        self.current_role = "viewer"

    def set_current_user(self, username: str, role: str) -> None:
        self.current_user = username
        self.current_role = role if role in ROLE_PERMISSIONS else "viewer"

    def can(self, action: str) -> bool:
        return action in ROLE_PERMISSIONS.get(self.current_role, {"view"})

    def require(self, action: str) -> None:
        if not self.can(action):
            raise PermissionError(f"Role {self.current_role} does not have '{action}' permission.")

    def authenticate(self, username: str, password: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one("SELECT * FROM users WHERE username=?", (username.strip(),))
        if not row or not row["aktif"]:
            return None
        valid, rehash = PasswordHasher.verify(password, row["password"])
        if not valid:
            self.db.log(username or "unknown", 'Login failed', 'Incorrect password')
            return None
        if rehash:
            self.db.execute(
                "UPDATE users SET password=? WHERE username=?",
                (PasswordHasher.hash(password), row["username"]),
            )
        self.db.execute(
            "UPDATE users SET login_terakhir=? WHERE username=?",
            (now_iso(), row["username"]),
        )
        self.db.log(row["username"], "Login", 'Login successful')
        result = dict(row)
        result["role"] = "operator" if result["role"] == "user" else result["role"]
        return result

    def change_password(self, username: str, old_password: str, new_password: str) -> None:
        row = self.db.fetch_one("SELECT password FROM users WHERE username=?", (username,))
        if not row or not PasswordHasher.verify(old_password, row["password"])[0]:
            raise ValueError('The current password is incorrect.')
        self.db.execute(
            "UPDATE users SET password=?, wajib_ganti_password=0 WHERE username=?",
            (PasswordHasher.hash(new_password), username),
        )
        self.db.log(username, 'Change password', 'Password updated successfully')

    @staticmethod
    def _validate_required(data: Dict[str, Any], fields: Sequence[str]) -> None:
        missing = [field for field in fields if not str(data.get(field, "")).strip()]
        if missing:
            raise ValueError('Required fields are missing: ' + ", ".join(missing))

    def _page(self, query: str, count_query: str, params: Sequence[Any], limit: int, offset: int) -> Tuple[List[Dict[str, Any]], int]:
        rows = [dict(row) for row in self.db.fetch_all(query + " LIMIT ? OFFSET ?", (*params, limit, offset))]
        total = int(self.db.scalar(count_query, params, 0))
        return rows, total

    # ---------------- Pengurus ----------------
    def save_pengurus(self, data: Dict[str, Any], item_id: Optional[str] = None) -> str:
        self.require("edit" if item_id else "add")
        self._validate_required(data, ["nama", "jabatan", "no_hp", "alamat"])
        if not valid_email(data.get("email", "")):
            raise ValueError('Invalid email format.')
        values = (
            data["nama"].strip(),
            data["jabatan"].strip(),
            normalize_phone(data["no_hp"]),
            data["alamat"].strip(),
            data.get("email", "").strip(),
            data.get("catatan", "").strip(),
            int(bool(data.get("aktif", True))),
        )
        if item_id:
            self.db.execute(
                """UPDATE pengurus SET nama=?,jabatan=?,no_hp=?,alamat=?,email=?,catatan=?,aktif=?
                   WHERE id=?""",
                (*values, item_id),
            )
            self.db.log(self.current_user, 'Edit committee member', item_id)
            return item_id
        item_id = self.db.next_id("pengurus")
        self.db.execute(
            """INSERT INTO pengurus
               (id,nama,jabatan,no_hp,alamat,email,catatan,aktif,tanggal_bergabung)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (item_id, *values, today_iso()),
        )
        self.db.log(self.current_user, 'Add committee member', f"{item_id} - {data['nama']}")
        return item_id

    def get_pengurus(self, item_id: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one("SELECT * FROM pengurus WHERE id=?", (item_id,))
        return dict(row) if row else None

    def page_pengurus(self, search: str, status: str, limit: int, offset: int) -> Tuple[List[Dict[str, Any]], int]:
        where = ["1=1"]
        params: List[Any] = []
        if search:
            where.append("(id LIKE ? OR nama LIKE ? OR jabatan LIKE ? OR no_hp LIKE ?)")
            token = f"%{search}%"
            params.extend([token] * 4)
        if status == "aktif":
            where.append("aktif=1")
        elif status == "nonaktif":
            where.append("aktif=0")
        clause = " WHERE " + " AND ".join(where)
        return self._page(
            "SELECT * FROM pengurus" + clause + " ORDER BY aktif DESC,nama",
            "SELECT COUNT(*) FROM pengurus" + clause,
            params,
            limit,
            offset,
        )

    def delete_pengurus(self, item_id: str, permanent: bool = False) -> None:
        self.require("delete")
        if permanent:
            self.db.execute("DELETE FROM pengurus WHERE id=?", (item_id,))
            action = 'Permanently delete committee member'
        else:
            self.db.execute("UPDATE pengurus SET aktif=0 WHERE id=?", (item_id,))
            action = 'Deactivate committee member'
        self.db.log(self.current_user, action, item_id)

    # ---------------- Donatur ----------------
    def save_donatur(self, data: Dict[str, Any], item_id: Optional[str] = None) -> str:
        self.require("edit" if item_id else "add")
        self._validate_required(data, ["nama", "no_hp", "alamat"])
        if not valid_email(data.get("email", "")):
            raise ValueError('Invalid email format.')
        values = (
            data["nama"].strip(),
            normalize_phone(data["no_hp"]),
            data["alamat"].strip(),
            data.get("email", "").strip(),
            data.get("kategori", "individu"),
            int(bool(data.get("aktif", True))),
            data.get("catatan", "").strip(),
        )
        if item_id:
            self.db.execute(
                """UPDATE donatur SET nama=?,no_hp=?,alamat=?,email=?,kategori=?,aktif=?,catatan=?
                   WHERE id=?""",
                (*values, item_id),
            )
            self.db.log(self.current_user, 'Edit donor', item_id)
            return item_id
        item_id = self.db.next_id("donatur")
        self.db.execute(
            """INSERT INTO donatur
               (id,nama,no_hp,alamat,email,kategori,aktif,catatan,tanggal_daftar)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (item_id, *values, today_iso()),
        )
        self.db.log(self.current_user, 'Add donor', f"{item_id} - {data['nama']}")
        return item_id

    def get_donatur(self, item_id: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one("SELECT * FROM donatur WHERE id=?", (item_id,))
        return dict(row) if row else None

    def list_donatur(self, active_only: bool = True) -> List[Dict[str, Any]]:
        query = "SELECT * FROM donatur"
        if active_only:
            query += " WHERE aktif=1"
        query += " ORDER BY nama"
        return [dict(row) for row in self.db.fetch_all(query)]

    def page_donatur(self, search: str, status: str, limit: int, offset: int) -> Tuple[List[Dict[str, Any]], int]:
        where = ["1=1"]
        params: List[Any] = []
        if search:
            where.append("(d.id LIKE ? OR d.nama LIKE ? OR d.no_hp LIKE ? OR d.email LIKE ?)")
            token = f"%{search}%"
            params.extend([token] * 4)
        if status == "aktif":
            where.append("d.aktif=1")
        elif status == "nonaktif":
            where.append("d.aktif=0")
        clause = " WHERE " + " AND ".join(where)
        select = """SELECT d.*,
                    COALESCE(SUM(CASE WHEN t.jenis='pemasukan' THEN t.jumlah ELSE 0 END),0) total_donasi,
                    MAX(CASE WHEN t.jenis='pemasukan' THEN t.tanggal END) donasi_terakhir
                    FROM donatur d LEFT JOIN transaksi t ON t.donatur_id=d.id"""
        group_order = " GROUP BY d.id ORDER BY d.aktif DESC,total_donasi DESC,d.nama"
        return self._page(
            select + clause + group_order,
            "SELECT COUNT(*) FROM donatur d" + clause,
            params,
            limit,
            offset,
        )

    def delete_donatur(self, item_id: str, permanent: bool = False) -> None:
        self.require("delete")
        if permanent:
            self.db.execute("DELETE FROM donatur WHERE id=?", (item_id,))
            action = 'Permanently delete donor'
        else:
            self.db.execute("UPDATE donatur SET aktif=0 WHERE id=?", (item_id,))
            action = 'Deactivate donor'
        self.db.log(self.current_user, action, item_id)

    # ---------------- Transaksi ----------------
    def save_transaksi(self, data: Dict[str, Any], item_id: Optional[str] = None) -> str:
        self.require("edit" if item_id else "add")
        self._validate_required(data, ["tanggal", "jenis", "kategori", "deskripsi"])
        amount = safe_float(data.get("jumlah"))
        if amount <= 0:
            raise ValueError('The transaction amount must be greater than zero.')
        if data["jenis"] not in {"pemasukan", "pengeluaran"}:
            raise ValueError('Invalid transaction type.')
        values = (
            str(data["tanggal"])[:19],
            data["jenis"],
            data["kategori"].strip(),
            data["deskripsi"].strip(),
            amount,
            data.get("donatur_id") or None,
            data.get("bukti", "").strip() or None,
            now_iso(),
        )
        if item_id:
            self.db.execute(
                """UPDATE transaksi SET tanggal=?,jenis=?,kategori=?,deskripsi=?,jumlah=?,
                   donatur_id=?,bukti=?,updated_at=? WHERE id=?""",
                (*values, item_id),
            )
            self.db.log(self.current_user, 'Edit transaction', item_id)
            return item_id
        item_id = self.db.next_id("transaksi")
        self.db.execute(
            """INSERT INTO transaksi
               (id,tanggal,jenis,kategori,deskripsi,jumlah,donatur_id,bukti,created_by,updated_at)
               VALUES(?,?,?,?,?,?,?,?,?,?)""",
            (item_id, *values[:-1], self.current_user, values[-1]),
        )
        self.db.log(self.current_user, 'Add transaction', f"{item_id} - {data['jenis']} {amount}")
        return item_id

    def get_transaksi(self, item_id: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one("SELECT * FROM transaksi WHERE id=?", (item_id,))
        return dict(row) if row else None

    def transaction_categories(self, jenis: Optional[str] = None) -> List[str]:
        params: List[Any] = []
        query = "SELECT DISTINCT kategori FROM transaksi WHERE TRIM(kategori)<>''"
        if jenis:
            query += " AND jenis=?"
            params.append(jenis)
        query += " ORDER BY kategori"
        return [str(row[0]) for row in self.db.fetch_all(query, params)]

    def page_transaksi(
        self,
        search: str,
        jenis: str,
        start_date: str,
        end_date: str,
        limit: int,
        offset: int,
    ) -> Tuple[List[Dict[str, Any]], int]:
        where = ["date(t.tanggal) BETWEEN date(?) AND date(?)"]
        params: List[Any] = [start_date, end_date]
        if search:
            where.append("(t.id LIKE ? OR t.kategori LIKE ? OR t.deskripsi LIKE ? OR d.nama LIKE ?)")
            token = f"%{search}%"
            params.extend([token] * 4)
        if jenis in {"pemasukan", "pengeluaran"}:
            where.append("t.jenis=?")
            params.append(jenis)
        clause = " WHERE " + " AND ".join(where)
        select = """SELECT t.*, COALESCE(d.nama,'') donatur_nama
                    FROM transaksi t LEFT JOIN donatur d ON d.id=t.donatur_id"""
        return self._page(
            select + clause + " ORDER BY t.tanggal DESC,t.id DESC",
            "SELECT COUNT(*) FROM transaksi t LEFT JOIN donatur d ON d.id=t.donatur_id" + clause,
            params,
            limit,
            offset,
        )

    def delete_transaksi(self, item_id: str) -> None:
        self.require("delete")
        self.db.execute("DELETE FROM transaksi WHERE id=?", (item_id,))
        self.db.log(self.current_user, 'Delete transaction', item_id)

    def finance_totals(self, start_date: Optional[str] = None, end_date: Optional[str] = None) -> Dict[str, float]:
        where = []
        params: List[Any] = []
        if start_date:
            where.append("date(tanggal)>=date(?)")
            params.append(start_date)
        if end_date:
            where.append("date(tanggal)<=date(?)")
            params.append(end_date)
        clause = " WHERE " + " AND ".join(where) if where else ""
        row = self.db.fetch_one(
            """SELECT
               COALESCE(SUM(CASE WHEN jenis='pemasukan' THEN jumlah ELSE 0 END),0),
               COALESCE(SUM(CASE WHEN jenis='pengeluaran' THEN jumlah ELSE 0 END),0)
               FROM transaksi""" + clause,
            params,
        )
        income, expense = float(row[0]), float(row[1])
        return {"pemasukan": income, "pengeluaran": expense, "saldo": income - expense}

    # ---------------- Kegiatan ----------------
    def save_kegiatan(self, data: Dict[str, Any], item_id: Optional[str] = None) -> str:
        self.require("edit" if item_id else "add")
        self._validate_required(data, ["nama", "tanggal_mulai", "tanggal_selesai", "deskripsi", "penanggung_jawab"])
        if parse_date(data["tanggal_selesai"]) < parse_date(data["tanggal_mulai"]):
            raise ValueError('The end date cannot be earlier than the start date.')
        values = (
            data["nama"].strip(),
            str(data["tanggal_mulai"])[:10],
            str(data["tanggal_selesai"])[:10],
            data["deskripsi"].strip(),
            data["penanggung_jawab"].strip(),
            max(0, safe_float(data.get("anggaran"))),
            max(0, safe_float(data.get("realisasi"))),
            data.get("status", "rencana"),
            data.get("lokasi", "").strip(),
            data.get("prioritas", "normal"),
            max(0, safe_int(data.get("pengingat_hari"), 3)),
        )
        if item_id:
            self.db.execute(
                """UPDATE kegiatan SET nama=?,tanggal_mulai=?,tanggal_selesai=?,deskripsi=?,
                   penanggung_jawab=?,anggaran=?,realisasi=?,status=?,lokasi=?,prioritas=?,pengingat_hari=?
                   WHERE id=?""",
                (*values, item_id),
            )
            self.db.log(self.current_user, 'Edit activity', item_id)
            return item_id
        item_id = self.db.next_id("kegiatan")
        self.db.execute(
            """INSERT INTO kegiatan
               (id,nama,tanggal_mulai,tanggal_selesai,deskripsi,penanggung_jawab,anggaran,
                realisasi,status,lokasi,prioritas,pengingat_hari)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (item_id, *values),
        )
        self.db.log(self.current_user, 'Add activity', f"{item_id} - {data['nama']}")
        return item_id

    def get_kegiatan(self, item_id: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one("SELECT * FROM kegiatan WHERE id=?", (item_id,))
        return dict(row) if row else None

    def page_kegiatan(self, search: str, status: str, limit: int, offset: int) -> Tuple[List[Dict[str, Any]], int]:
        where = ["1=1"]
        params: List[Any] = []
        if search:
            where.append("(id LIKE ? OR nama LIKE ? OR penanggung_jawab LIKE ? OR lokasi LIKE ?)")
            token = f"%{search}%"
            params.extend([token] * 4)
        if status and status != "semua":
            where.append("status=?")
            params.append(status)
        clause = " WHERE " + " AND ".join(where)
        return self._page(
            "SELECT * FROM kegiatan" + clause + " ORDER BY tanggal_mulai DESC",
            "SELECT COUNT(*) FROM kegiatan" + clause,
            params,
            limit,
            offset,
        )

    def upcoming_activities(
        self,
        days: Optional[int] = 30,
        limit: Optional[int] = 10,
        newest_first: bool = False,
    ) -> List[Dict[str, Any]]:
        where = [
            "status NOT IN ('selesai','batal')",
            "date(tanggal_mulai)>=date(?)",
        ]
        params: List[Any] = [today_iso()]
        if days is not None:
            end = (dt.date.today() + dt.timedelta(days=days)).isoformat()
            where.append("date(tanggal_mulai)<=date(?)")
            params.append(end)

        direction = "DESC" if newest_first else "ASC"
        query = (
            "SELECT * FROM kegiatan WHERE "
            + " AND ".join(where)
            + f" ORDER BY tanggal_mulai {direction}, id {direction}"
        )
        if limit is not None:
            query += " LIMIT ?"
            params.append(max(0, int(limit)))
        return [dict(row) for row in self.db.fetch_all(query, params)]

    def delete_kegiatan(self, item_id: str) -> None:
        self.require("delete")
        self.db.execute("DELETE FROM kegiatan WHERE id=?", (item_id,))
        self.db.log(self.current_user, 'Delete activity', item_id)

    # ---------------- Aset ----------------
    def save_aset(self, data: Dict[str, Any], item_id: Optional[str] = None) -> str:
        self.require("edit" if item_id else "add")
        self._validate_required(data, ["nama", "kategori", "kondisi"])
        acquisition = str(data.get("tanggal_perolehan") or today_iso())[:10]
        values = (
            data["nama"].strip(),
            data["kategori"],
            max(0, safe_float(data.get("nilai"))),
            parse_date(acquisition).year,
            data["kondisi"],
            data.get("keterangan", "").strip(),
            data.get("lokasi", "").strip(),
            acquisition,
            str(data.get("tanggal_perawatan_terakhir") or "")[:10],
            str(data.get("tanggal_perawatan_berikut") or "")[:10],
            int(bool(data.get("aktif", True))),
        )
        if item_id:
            self.db.execute(
                """UPDATE aset SET nama=?,kategori=?,nilai=?,tahun_perolehan=?,kondisi=?,
                   keterangan=?,lokasi=?,tanggal_perolehan=?,tanggal_perawatan_terakhir=?,
                   tanggal_perawatan_berikut=?,aktif=? WHERE id=?""",
                (*values, item_id),
            )
            self.db.log(self.current_user, 'Edit asset', item_id)
            return item_id
        item_id = self.db.next_id("aset")
        self.db.execute(
            """INSERT INTO aset
               (id,nama,kategori,nilai,tahun_perolehan,kondisi,keterangan,lokasi,
                tanggal_perolehan,tanggal_perawatan_terakhir,tanggal_perawatan_berikut,aktif)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (item_id, *values),
        )
        self.db.log(self.current_user, 'Add asset', f"{item_id} - {data['nama']}")
        return item_id

    def get_aset(self, item_id: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one("SELECT * FROM aset WHERE id=?", (item_id,))
        return dict(row) if row else None

    def page_aset(self, search: str, category: str, condition: str, limit: int, offset: int) -> Tuple[List[Dict[str, Any]], int]:
        where = ["1=1"]
        params: List[Any] = []
        if search:
            where.append("(id LIKE ? OR nama LIKE ? OR lokasi LIKE ? OR keterangan LIKE ?)")
            token = f"%{search}%"
            params.extend([token] * 4)
        if category and category != "semua":
            where.append("kategori=?")
            params.append(category)
        if condition and condition != "semua":
            where.append("kondisi=?")
            params.append(condition)
        clause = " WHERE " + " AND ".join(where)
        return self._page(
            "SELECT * FROM aset" + clause + " ORDER BY aktif DESC,nama",
            "SELECT COUNT(*) FROM aset" + clause,
            params,
            limit,
            offset,
        )

    def maintenance_due(
        self,
        days: Optional[int] = 30,
        limit: Optional[int] = 10,
        newest_first: bool = False,
    ) -> List[Dict[str, Any]]:
        where = [
            "aktif=1",
            "COALESCE(tanggal_perawatan_berikut,'')<>''",
        ]
        params: List[Any] = []
        if days is not None:
            end = (dt.date.today() + dt.timedelta(days=days)).isoformat()
            where.append("date(tanggal_perawatan_berikut)<=date(?)")
            params.append(end)

        direction = "DESC" if newest_first else "ASC"
        query = (
            "SELECT * FROM aset WHERE "
            + " AND ".join(where)
            + f" ORDER BY tanggal_perawatan_berikut {direction}, id {direction}"
        )
        if limit is not None:
            query += " LIMIT ?"
            params.append(max(0, int(limit)))
        return [dict(row) for row in self.db.fetch_all(query, params)]

    def delete_aset(self, item_id: str, permanent: bool = False) -> None:
        self.require("delete")
        if permanent:
            self.db.execute("DELETE FROM aset WHERE id=?", (item_id,))
            action = 'Permanently delete asset'
        else:
            self.db.execute("UPDATE aset SET aktif=0 WHERE id=?", (item_id,))
            action = 'Deactivate asset'
        self.db.log(self.current_user, action, item_id)

    # ---------------- Transaksi Rutin ----------------
    def save_recurring(self, data: Dict[str, Any], item_id: Optional[str] = None) -> str:
        self.require("edit" if item_id else "add")
        self._validate_required(data, ["nama", "jenis", "kategori", "deskripsi", "tanggal_berikut"])
        amount = safe_float(data.get("jumlah"))
        if amount <= 0:
            raise ValueError('The recurring transaction amount must be greater than zero.')
        values = (
            data["nama"].strip(),
            data["jenis"],
            data["kategori"].strip(),
            data["deskripsi"].strip(),
            amount,
            data.get("donatur_id") or None,
            max(1, safe_int(data.get("interval_nilai"), 1)),
            data.get("interval_unit", "bulan"),
            str(data["tanggal_berikut"])[:10],
            int(bool(data.get("aktif", True))),
        )
        if item_id:
            self.db.execute(
                """UPDATE transaksi_rutin SET nama=?,jenis=?,kategori=?,deskripsi=?,jumlah=?,
                   donatur_id=?,interval_nilai=?,interval_unit=?,tanggal_berikut=?,aktif=? WHERE id=?""",
                (*values, item_id),
            )
            self.db.log(self.current_user, 'Edit recurring transaction', item_id)
            return item_id
        item_id = self.db.next_id("transaksi_rutin")
        self.db.execute(
            """INSERT INTO transaksi_rutin
               (id,nama,jenis,kategori,deskripsi,jumlah,donatur_id,interval_nilai,
                interval_unit,tanggal_berikut,aktif)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            (item_id, *values),
        )
        self.db.log(self.current_user, 'Add recurring transaction', f"{item_id} - {data['nama']}")
        return item_id

    def get_recurring(self, item_id: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one("SELECT * FROM transaksi_rutin WHERE id=?", (item_id,))
        return dict(row) if row else None

    def page_recurring(self, search: str, active: str, limit: int, offset: int) -> Tuple[List[Dict[str, Any]], int]:
        where = ["1=1"]
        params: List[Any] = []
        if search:
            where.append("(r.id LIKE ? OR r.nama LIKE ? OR r.kategori LIKE ? OR r.deskripsi LIKE ?)")
            token = f"%{search}%"
            params.extend([token] * 4)
        if active == "aktif":
            where.append("r.aktif=1")
        elif active == "nonaktif":
            where.append("r.aktif=0")
        clause = " WHERE " + " AND ".join(where)
        select = """SELECT r.*,COALESCE(d.nama,'') donatur_nama
                    FROM transaksi_rutin r LEFT JOIN donatur d ON d.id=r.donatur_id"""
        return self._page(
            select + clause + " ORDER BY r.aktif DESC,r.tanggal_berikut",
            "SELECT COUNT(*) FROM transaksi_rutin r" + clause,
            params,
            limit,
            offset,
        )

    def post_due_recurring(self, through_date: Optional[str] = None) -> int:
        self.require("add")
        through = parse_date(through_date or today_iso())
        rows = self.db.fetch_all(
            "SELECT * FROM transaksi_rutin WHERE aktif=1 AND date(tanggal_berikut)<=date(?) ORDER BY tanggal_berikut",
            (through.isoformat(),),
        )
        posted = 0
        for row in rows:
            current = parse_date(row["tanggal_berikut"])
            safety = 0
            while current <= through and safety < 120:
                self.save_transaksi(
                    {
                        "tanggal": current.isoformat() + " 08:00:00",
                        "jenis": row["jenis"],
                        "kategori": row["kategori"],
                        "deskripsi": f"{row['deskripsi']} [Recurring {row['id']}]",
                        "jumlah": row["jumlah"],
                        "donatur_id": row["donatur_id"],
                        "bukti": "",
                    }
                )
                posted += 1
                unit = row["interval_unit"]
                value = max(1, int(row["interval_nilai"]))
                if unit == "hari":
                    current += dt.timedelta(days=value)
                elif unit == "minggu":
                    current += dt.timedelta(weeks=value)
                elif unit == "tahun":
                    current = shift_month(current, 12 * value)
                else:
                    current = shift_month(current, value)
                safety += 1
            self.db.execute(
                "UPDATE transaksi_rutin SET tanggal_berikut=?,terakhir_diproses=? WHERE id=?",
                (current.isoformat(), now_iso(), row["id"]),
            )
        if posted:
            self.db.log(self.current_user, 'Process recurring transactions', f"{posted} transactions created")
        return posted

    def delete_recurring(self, item_id: str) -> None:
        self.require("delete")
        self.db.execute("DELETE FROM transaksi_rutin WHERE id=?", (item_id,))
        self.db.log(self.current_user, 'Delete recurring transaction', item_id)

    # ---------------- Dashboard & laporan ----------------
    def dashboard_summary(self) -> Dict[str, Any]:
        totals = self.finance_totals()
        month_start = dt.date.today().replace(day=1).isoformat()
        month_totals = self.finance_totals(month_start, today_iso())
        return {
            **totals,
            "pemasukan_bulan": month_totals["pemasukan"],
            "pengeluaran_bulan": month_totals["pengeluaran"],
            "pengurus_aktif": int(self.db.scalar("SELECT COUNT(*) FROM pengurus WHERE aktif=1")),
            "donatur_aktif": int(self.db.scalar("SELECT COUNT(*) FROM donatur WHERE aktif=1")),
            "kegiatan_mendatang": int(
                self.db.scalar(
                    "SELECT COUNT(*) FROM kegiatan WHERE status NOT IN ('selesai','batal') AND date(tanggal_mulai)>=date(?)",
                    (today_iso(),),
                )
            ),
            "nilai_aset": float(self.db.scalar("SELECT COALESCE(SUM(nilai),0) FROM aset WHERE aktif=1")),
            "perawatan_jatuh_tempo": len(self.maintenance_due(30, 1000)),
        }

    def monthly_cashflow(self, months: int = 12) -> List[Dict[str, Any]]:
        result: List[Dict[str, Any]] = []
        first = dt.date.today().replace(day=1)
        for delta in range(months - 1, -1, -1):
            month_date = shift_month(first, -delta)
            start, end = month_bounds(month_date.year, month_date.month)
            totals = self.finance_totals(start, end)
            result.append(
                {
                    "label": _APP_LOCALE.toString(QDate(month_date.year, month_date.month, 1), "MMM yyyy"),
                    "start": start,
                    "end": end,
                    **totals,
                }
            )
        return result

    def recent_transactions(self, limit: Optional[int] = 8) -> List[Dict[str, Any]]:
        query = "SELECT * FROM transaksi ORDER BY tanggal DESC, id DESC"
        params: List[Any] = []
        if limit is not None:
            query += " LIMIT ?"
            params.append(max(0, int(limit)))
        return [dict(row) for row in self.db.fetch_all(query, params)]

    def financial_report(self, start_date: str, end_date: str) -> Dict[str, Any]:
        totals = self.finance_totals(start_date, end_date)
        categories = [
            dict(row)
            for row in self.db.fetch_all(
                """SELECT jenis,kategori,COUNT(*) jumlah_transaksi,SUM(jumlah) total
                   FROM transaksi WHERE date(tanggal) BETWEEN date(?) AND date(?)
                   GROUP BY jenis,kategori ORDER BY jenis,kategori""",
                (start_date, end_date),
            )
        ]
        transactions = [
            dict(row)
            for row in self.db.fetch_all(
                """SELECT t.id,t.tanggal,t.jenis,t.kategori,t.deskripsi,t.jumlah,
                   COALESCE(d.nama,'') donatur
                   FROM transaksi t LEFT JOIN donatur d ON d.id=t.donatur_id
                   WHERE date(t.tanggal) BETWEEN date(?) AND date(?)
                   ORDER BY t.tanggal,t.id""",
                (start_date, end_date),
            )
        ]
        return {
            "start_date": start_date,
            "end_date": end_date,
            **totals,
            "categories": categories,
            "transactions": transactions,
        }

    # ---------------- Export / Import ----------------
    def export_rows_csv(self, rows: List[Dict[str, Any]], filename: str) -> str:
        if not rows:
            raise ValueError('There is no data to export.')
        with open(filename, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            writer.writerows(rows)
        return filename

    def export_rows_excel(self, rows: List[Dict[str, Any]], filename: str, sheet_name: str = "Data") -> str:
        if not HAS_OPENPYXL:
            raise RuntimeError('OpenPyXL is not installed: pip install openpyxl')
        if not rows:
            raise ValueError('There is no data to export.')
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = re.sub(r"[\\/*?:\[\]]", "_", sheet_name)[:31]
        headers = list(rows[0].keys())
        worksheet.append(headers)
        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2D6A4F")
            cell.alignment = Alignment(horizontal="center")
        for row in rows:
            worksheet.append([row.get(header, "") for header in headers])
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for index, header in enumerate(headers, 1):
            max_len = max(len(str(header)), *(len(str(row.get(header, ""))) for row in rows))
            worksheet.column_dimensions[get_column_letter(index)].width = min(max_len + 2, 45)
        workbook.save(filename)
        return filename

    def export_rows_pdf(self, rows: List[Dict[str, Any]], filename: str, title: str) -> str:
        if not HAS_REPORTLAB:
            raise RuntimeError('ReportLab is not installed: pip install reportlab')
        if not rows:
            raise ValueError('There is no data to export.')
        styles = getSampleStyleSheet()
        small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=7, leading=8)
        document = SimpleDocTemplate(
            filename,
            pagesize=landscape(A4),
            rightMargin=24,
            leftMargin=24,
            topMargin=24,
            bottomMargin=24,
        )
        story = [Paragraph(title, styles["Title"]), Spacer(1, 8)]
        story.append(Paragraph(f"Dibuat: {now_iso()}", styles["Normal"]))
        story.append(Spacer(1, 10))
        headers = list(rows[0].keys())
        data = [[Paragraph(str(header), small) for header in headers]]
        for row in rows:
            data.append([Paragraph(elide(row.get(header, ""), 80), small) for header in headers])
        table = Table(data, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2D6A4F")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F7F4")]),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(table)
        document.build(story)
        return filename

    def table_rows(self, table: str) -> List[Dict[str, Any]]:
        if table not in ALLOWED_DATA_TABLES:
            raise ValueError('Table is not allowed.')
        return [dict(row) for row in self.db.fetch_all(f"SELECT * FROM {table} ORDER BY 1")]

    def import_file(self, table: str, filename: str) -> Dict[str, Any]:
        self.require("import")
        if table not in ALLOWED_DATA_TABLES:
            raise ValueError('Table is not allowed.')
        extension = Path(filename).suffix.lower()
        records: List[Dict[str, Any]] = []
        if extension == ".csv":
            with open(filename, "r", encoding="utf-8-sig", newline="") as handle:
                records = [dict(row) for row in csv.DictReader(handle)]
        elif extension in {".xlsx", ".xlsm"}:
            if not HAS_OPENPYXL:
                raise RuntimeError('OpenPyXL is not installed.')
            worksheet = load_workbook(filename, read_only=True, data_only=True).active
            iterator = worksheet.iter_rows(values_only=True)
            headers = [str(value).strip() if value is not None else "" for value in next(iterator)]
            records = [dict(zip(headers, row)) for row in iterator]
        else:
            raise ValueError('The import format must be CSV or XLSX.')
        columns = self.db.table_columns(table)
        inserted = updated = skipped = 0
        errors: List[str] = []
        for row_number, record in enumerate(records, 2):
            clean = {key: value for key, value in record.items() if key in columns and key}
            if not clean:
                skipped += 1
                continue
            item_id = str(clean.get("id", "")).strip()
            if not item_id:
                clean["id"] = self.db.next_id(table)
            keys = list(clean.keys())
            placeholders = ",".join("?" for _ in keys)
            update_columns = [key for key in keys if key != "id"]
            update_clause = ",".join(f"{key}=excluded.{key}" for key in update_columns)
            sql = f"INSERT INTO {table} ({','.join(keys)}) VALUES ({placeholders})"
            if update_clause:
                sql += f" ON CONFLICT(id) DO UPDATE SET {update_clause}"
            existed = bool(self.db.fetch_one(f"SELECT 1 FROM {table} WHERE id=?", (clean["id"],)))
            try:
                self.db.execute(sql, tuple(clean[key] for key in keys))
                updated += int(existed)
                inserted += int(not existed)
            except Exception as exc:
                errors.append(f"Row {row_number}: {exc}")
        self.db.log(
            self.current_user,
            'Import data',
            f"{table}: inserted={inserted}, updated={updated}, failed={len(errors)}",
        )
        return {"inserted": inserted, "updated": updated, "skipped": skipped, "errors": errors}

    # ---------------- Users, logs, maintenance ----------------
    def list_users(self) -> List[Dict[str, Any]]:
        self.require("users")
        return [
            dict(row)
            for row in self.db.fetch_all(
                "SELECT username,role,nama_lengkap,aktif,dibuat_pada,login_terakhir,wajib_ganti_password FROM users ORDER BY username"
            )
        ]

    def save_user(self, data: Dict[str, Any], original_username: Optional[str] = None) -> str:
        self.require("users")
        username = validate_username(original_username or data.get("username", ""))
        role = data.get("role", "viewer")
        if original_username == self.current_user and not data.get("aktif", True):
            raise ValueError('The currently signed-in user cannot be deactivated.')
        if role not in {"admin", "operator", "viewer"}:
            raise ValueError('Invalid role.')
        if original_username:
            existing = self.db.fetch_one(
                "SELECT role,aktif FROM users WHERE username=?", (original_username,)
            )
            if not existing:
                raise ValueError("The selected user no longer exists.")
            removing_last_admin = (
                existing["role"] == "admin"
                and bool(existing["aktif"])
                and (role != "admin" or not bool(data.get("aktif", True)))
                and int(
                    self.db.scalar(
                        "SELECT COUNT(*) FROM users WHERE role='admin' AND aktif=1",
                        default=0,
                    )
                ) <= 1
            )
            if removing_last_admin:
                raise ValueError("At least one active administrator account is required.")
            fields = ["role=?", "nama_lengkap=?", "aktif=?", "wajib_ganti_password=?"]
            params: List[Any] = [
                role,
                data.get("nama_lengkap", "").strip(),
                int(bool(data.get("aktif", True))),
                int(bool(data.get("wajib_ganti_password", False))),
            ]
            if data.get("password"):
                fields.append("password=?")
                params.append(PasswordHasher.hash(data["password"]))
            params.append(original_username)
            self.db.execute(f"UPDATE users SET {','.join(fields)} WHERE username=?", params)
            self.db.log(self.current_user, 'Edit user', original_username)
            return original_username
        password = data.get("password", "")
        if not password:
            raise ValueError('A password is required for a new user.')
        self.db.execute(
            """INSERT INTO users
               (username,password,role,nama_lengkap,aktif,dibuat_pada,wajib_ganti_password)
               VALUES(?,?,?,?,?,?,?)""",
            (
                username,
                PasswordHasher.hash(password),
                role,
                data.get("nama_lengkap", "").strip(),
                int(bool(data.get("aktif", True))),
                now_iso(),
                int(bool(data.get("wajib_ganti_password", True))),
            ),
        )
        self.db.log(self.current_user, 'Add user', username)
        return username

    def delete_user(self, username: str) -> None:
        self.require("users")
        if username == self.current_user:
            raise ValueError('The currently signed-in user cannot be deleted.')
        existing = self.db.fetch_one(
            "SELECT role,aktif FROM users WHERE username=?", (username,)
        )
        if not existing:
            raise ValueError("The selected user no longer exists.")
        if (
            existing["role"] == "admin"
            and bool(existing["aktif"])
            and int(
                self.db.scalar(
                    "SELECT COUNT(*) FROM users WHERE role='admin' AND aktif=1",
                    default=0,
                )
            ) <= 1
        ):
            raise ValueError("The last active administrator account cannot be deleted.")
        self.db.execute("DELETE FROM users WHERE username=?", (username,))
        self.db.log(self.current_user, 'Delete user', username)

    def logs(self, search: str = "", limit: int = 500) -> List[Dict[str, Any]]:
        if search:
            token = f"%{search}%"
            rows = self.db.fetch_all(
                """SELECT * FROM log_aktivitas
                   WHERE user LIKE ? OR aksi LIKE ? OR detail LIKE ?
                   ORDER BY id DESC LIMIT ?""",
                (token, token, token, limit),
            )
        else:
            rows = self.db.fetch_all(
                "SELECT * FROM log_aktivitas ORDER BY id DESC LIMIT ?", (limit,)
            )
        return [dict(row) for row in rows]

    def cleanup_logs(self, older_than_days: int) -> int:
        self.require("delete")
        cutoff = (dt.datetime.now() - dt.timedelta(days=older_than_days)).strftime(DATETIME_FMT)
        cursor = self.db.execute("DELETE FROM log_aktivitas WHERE waktu<?", (cutoff,))
        count = cursor.rowcount
        self.db.log(self.current_user, 'Clean logs', f"{count} rows before {cutoff}")
        return count

# ============================== UI HELPERS ==============================


def make_money_spin(value: float = 0.0) -> QDoubleSpinBox:
    profile = CURRENCY_PROFILES[_CURRENT_CURRENCY]
    widget = QDoubleSpinBox()
    widget.setRange(0, 999_999_999_999_999)
    widget.setDecimals(int(profile["decimals"]))
    widget.setSingleStep(float(profile["step"]))
    widget.setGroupSeparatorShown(True)
    widget.setPrefix(f"{profile['symbol']} ")
    widget.setValue(float(value or 0))
    return widget


def qdate(value: Any, fallback: Optional[dt.date] = None) -> QDate:
    date_value = parse_date(value, fallback)
    return QDate(date_value.year, date_value.month, date_value.day)


def selected_id(table: QTableWidget) -> Optional[str]:
    row = table.currentRow()
    if row < 0 or not table.item(row, 0):
        return None
    return table.item(row, 0).text()


def set_table_item(table: QTableWidget, row: int, column: int, value: Any, user_data: Any = None) -> None:
    item = QTableWidgetItem(str(value if value is not None else ""))
    if user_data is not None:
        item.setData(Qt.ItemDataRole.UserRole, user_data)
    table.setItem(row, column, item)


class OptionalDateWidget(QWidget):
    def __init__(self, value: str = "", parent: Optional[QWidget] = None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.enabled_box = QCheckBox('Set date')
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat(DATE_DISPLAY_FMT)
        self.date_edit.setDate(qdate(value))
        self.enabled_box.setChecked(bool(value))
        self.date_edit.setEnabled(bool(value))
        self.enabled_box.toggled.connect(self.date_edit.setEnabled)
        layout.addWidget(self.enabled_box)
        layout.addWidget(self.date_edit, 1)

    def value(self) -> str:
        if not self.enabled_box.isChecked():
            return ""
        return self.date_edit.date().toString("yyyy-MM-dd")


class BaseDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        app = QApplication.instance()
        if app is not None:
            self.setWindowIcon(app.windowIcon())
        self.setWindowModality(Qt.WindowModality.WindowModal)

    def showEvent(self, event) -> None:  # noqa: N802 - Qt naming convention
        super().showEvent(event)
        parent = self.parentWidget()
        if parent:
            frame = self.frameGeometry()
            frame.moveCenter(parent.frameGeometry().center())
            self.move(frame.topLeft())

    def add_buttons(self, layout: QFormLayout | QVBoxLayout) -> QDialogButtonBox:
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        if isinstance(layout, QFormLayout):
            layout.addRow(buttons)
        else:
            layout.addWidget(buttons)
        return buttons


class FirstRunSetupDialog(BaseDialog):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(f"First-Time Setup - {APP_NAME}")
        self.setFixedSize(460, 410)
        layout = QVBoxLayout(self)

        app = QApplication.instance()
        icon_label = QLabel()
        if app is not None:
            icon_label.setPixmap(app.windowIcon().pixmap(72, 72))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)

        title = QLabel("Create the First Administrator")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Segoe UI", 17, QFont.Weight.Bold))
        subtitle = QLabel(
            "No user account exists yet. Choose the username and password "
            "for the first administrator."
        )
        subtitle.setWordWrap(True)
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        form = QFormLayout()
        self.username = QLineEdit()
        self.username.setMaxLength(40)
        self.password = QLineEdit()
        self.confirm_password = QLineEdit()
        for widget in (self.password, self.confirm_password):
            widget.setEchoMode(QLineEdit.EchoMode.Password)
            widget.setMaxLength(PASSWORD_MAX_LENGTH)
        self.confirm_password.returnPressed.connect(self.accept)
        self.show_password = QCheckBox("Show password")
        self.show_password.toggled.connect(self._toggle_password_visibility)
        self.remember_username = QCheckBox("Remember username")
        self.remember_username.setChecked(True)

        form.addRow("Username", self.username)
        form.addRow("Password", self.password)
        form.addRow("Confirm password", self.confirm_password)
        form.addRow("", self.show_password)
        form.addRow("", self.remember_username)
        layout.addLayout(form)

        policy = QLabel(
            "Password policy: maximum 6 characters; letters and numbers only."
        )
        policy.setWordWrap(True)
        policy.setAlignment(Qt.AlignmentFlag.AlignCenter)
        policy.setStyleSheet("color: #b7c1c9; font-size: 9pt;")
        layout.addWidget(policy)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.username.setFocus()

    def _toggle_password_visibility(self, checked: bool) -> None:
        mode = (
            QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
        )
        self.password.setEchoMode(mode)
        self.confirm_password.setEchoMode(mode)

    def accept(self) -> None:
        try:
            validate_username(self.username.text())
            validate_password(self.password.text())
        except ValueError as exc:
            QMessageBox.warning(self, "Validation", str(exc))
            return
        if self.password.text() != self.confirm_password.text():
            QMessageBox.warning(
                self, "Validation", "The password confirmation does not match."
            )
            return
        super().accept()

    def data(self) -> Tuple[str, str, bool]:
        return (
            self.username.text().strip(),
            self.password.text(),
            self.remember_username.isChecked(),
        )


class LoginDialog(BaseDialog):
    def __init__(self, settings: SettingsManager, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle(f"Sign In - {APP_NAME}")
        self.setFixedSize(420, 365)
        layout = QVBoxLayout(self)
        app = QApplication.instance()
        icon_label = QLabel()
        if app is not None:
            icon_label.setPixmap(app.windowIcon().pixmap(72, 72))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)
        title = QLabel('Mosque Management')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        subtitle = QLabel('Sign in with a registered account')
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        form = QFormLayout()
        remember = settings.get_bool("User", "remember_username", False)
        remembered_name = settings.get("User", "last_username", "") if remember else ""
        self.username = QLineEdit(remembered_name)
        self.username.setMaxLength(40)
        self.password = QLineEdit()
        self.password.setMaxLength(PASSWORD_MAX_LENGTH)
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        self.password.returnPressed.connect(self.accept)
        self.show_password = QCheckBox('Show password')
        self.show_password.toggled.connect(
            lambda checked: self.password.setEchoMode(
                QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
            )
        )
        self.remember_username = QCheckBox('Remember username')
        self.remember_username.setChecked(remember)
        form.addRow("Username", self.username)
        form.addRow("Password", self.password)
        form.addRow("", self.show_password)
        form.addRow("", self.remember_username)
        layout.addLayout(form)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        note = QLabel('Password: maximum 6 characters; letters and numbers only.')
        note.setWordWrap(True)
        note.setAlignment(Qt.AlignmentFlag.AlignCenter)
        note.setStyleSheet("color: #b7c1c9; font-size: 9pt;")
        layout.addWidget(note)
        if remembered_name:
            self.password.setFocus()
        else:
            self.username.setFocus()

    def credentials(self) -> Tuple[str, str, bool]:
        return (
            self.username.text().strip(),
            self.password.text(),
            self.remember_username.isChecked(),
        )


class PasswordDialog(BaseDialog):
    def __init__(self, parent: Optional[QWidget] = None, mandatory: bool = False):
        super().__init__(parent)
        self.mandatory = mandatory
        self.setWindowTitle('Change Password')
        self.setFixedSize(420, 260)
        form = QFormLayout(self)
        self.old_password = QLineEdit()
        self.new_password = QLineEdit()
        self.confirm_password = QLineEdit()
        for widget in (self.old_password, self.new_password, self.confirm_password):
            widget.setEchoMode(QLineEdit.EchoMode.Password)
            widget.setMaxLength(PASSWORD_MAX_LENGTH)
        form.addRow('Current password', self.old_password)
        form.addRow('New password', self.new_password)
        form.addRow('Confirm password', self.confirm_password)
        info = QLabel('Use 1–6 characters. Letters and numbers only.')
        info.setWordWrap(True)
        form.addRow(info)
        buttons = self.add_buttons(form)
        if mandatory:
            buttons.button(QDialogButtonBox.StandardButton.Cancel).setEnabled(False)

    def accept(self) -> None:
        try:
            validate_password(self.new_password.text())
        except ValueError as exc:
            QMessageBox.warning(self, 'Validation', str(exc))
            return
        if self.new_password.text() != self.confirm_password.text():
            QMessageBox.warning(self, 'Validation', 'The password confirmation does not match.')
            return
        super().accept()

    def reject(self) -> None:
        if self.mandatory:
            QMessageBox.warning(self, 'Security', 'You must change the password before continuing.')
            return
        super().reject()

    def data(self) -> Tuple[str, str]:
        return self.old_password.text(), self.new_password.text()


class PengurusDialog(BaseDialog):
    def __init__(self, data: Optional[Dict[str, Any]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle('Add Committee Member' if not data else 'Edit Committee Member')
        self.resize(500, 430)
        form = QFormLayout(self)
        self.nama = QLineEdit(data.get("nama", "") if data else "")
        self.jabatan = QLineEdit(data.get("jabatan", "") if data else "")
        self.no_hp = QLineEdit(data.get("no_hp", "") if data else "")
        self.email = QLineEdit(data.get("email", "") if data else "")
        self.alamat = QTextEdit(data.get("alamat", "") if data else "")
        self.alamat.setMaximumHeight(80)
        self.catatan = QTextEdit(data.get("catatan", "") if data else "")
        self.catatan.setMaximumHeight(70)
        self.aktif = QCheckBox('Active')
        self.aktif.setChecked(bool(data.get("aktif", 1)) if data else True)
        form.addRow('Name *', self.nama)
        form.addRow('Position *', self.jabatan)
        form.addRow('Phone *', self.no_hp)
        form.addRow("Email", self.email)
        form.addRow('Address *', self.alamat)
        form.addRow('Notes', self.catatan)
        form.addRow("Status", self.aktif)
        self.add_buttons(form)

    def data(self) -> Dict[str, Any]:
        return {
            "nama": self.nama.text(),
            "jabatan": self.jabatan.text(),
            "no_hp": self.no_hp.text(),
            "email": self.email.text(),
            "alamat": self.alamat.toPlainText(),
            "catatan": self.catatan.toPlainText(),
            "aktif": self.aktif.isChecked(),
        }


class DonaturDialog(BaseDialog):
    def __init__(self, data: Optional[Dict[str, Any]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle('Add Donor' if not data else 'Edit Donor')
        self.resize(500, 450)
        form = QFormLayout(self)
        self.nama = QLineEdit(data.get("nama", "") if data else "")
        self.no_hp = QLineEdit(data.get("no_hp", "") if data else "")
        self.email = QLineEdit(data.get("email", "") if data else "")
        self.kategori = QComboBox()
        add_enum_items(self.kategori, "donor_category", ["individu", "keluarga", "perusahaan", "yayasan", "komunitas"])
        self.kategori.setEditable(True)
        if data:
            set_combo_data(self.kategori, data.get("kategori", "individu"))
        self.alamat = QTextEdit(data.get("alamat", "") if data else "")
        self.alamat.setMaximumHeight(80)
        self.catatan = QTextEdit(data.get("catatan", "") if data else "")
        self.catatan.setMaximumHeight(70)
        self.aktif = QCheckBox('Active')
        self.aktif.setChecked(bool(data.get("aktif", 1)) if data else True)
        form.addRow('Name *', self.nama)
        form.addRow('Phone *', self.no_hp)
        form.addRow("Email", self.email)
        form.addRow('Category', self.kategori)
        form.addRow('Address *', self.alamat)
        form.addRow('Notes', self.catatan)
        form.addRow("Status", self.aktif)
        self.add_buttons(form)

    def data(self) -> Dict[str, Any]:
        return {
            "nama": self.nama.text(),
            "no_hp": self.no_hp.text(),
            "email": self.email.text(),
            "kategori": combo_data_or_text(self.kategori),
            "alamat": self.alamat.toPlainText(),
            "catatan": self.catatan.toPlainText(),
            "aktif": self.aktif.isChecked(),
        }


class TransaksiDialog(BaseDialog):
    def __init__(
        self,
        donors: List[Dict[str, Any]],
        categories: List[str],
        data: Optional[Dict[str, Any]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle('Add Transaction' if not data else 'Edit Transaction')
        self.resize(560, 470)
        form = QFormLayout(self)
        self.tanggal = QDateTimeEdit()
        self.tanggal.setCalendarPopup(True)
        self.tanggal.setDisplayFormat(DATETIME_DISPLAY_FMT)
        if data:
            parsed = QDateTime.fromString(str(data.get("tanggal", ""))[:19], "yyyy-MM-dd HH:mm:ss")
            self.tanggal.setDateTime(parsed if parsed.isValid() else QDateTime.currentDateTime())
        else:
            self.tanggal.setDateTime(QDateTime.currentDateTime())
        self.jenis = QComboBox()
        add_enum_items(self.jenis, "transaction_type", ["pemasukan", "pengeluaran"])
        set_combo_data(self.jenis, data.get("jenis", "pemasukan") if data else "pemasukan")
        self.kategori = QComboBox()
        self.kategori.setEditable(True)
        defaults = [
            'Donation',
            'Charity',
            "Zakat",
            'Activity Donation',
            'Operations',
            'Electricity & Water',
            'Maintenance',
            'Honorarium',
            'Social Services',
            'Other',
        ]
        self.kategori.addItems(list(dict.fromkeys(defaults + categories)))
        if data:
            self.kategori.setCurrentText(data.get("kategori", ""))
        self.deskripsi = QTextEdit(data.get("deskripsi", "") if data else "")
        self.deskripsi.setMaximumHeight(85)
        self.jumlah = make_money_spin(data.get("jumlah", 0) if data else 0)
        self.donatur = QComboBox()
        self.donatur.addItem('— No linked donor —', None)
        for donor in donors:
            self.donatur.addItem(f"{donor['nama']} ({donor['id']})", donor["id"])
        if data and data.get("donatur_id"):
            index = self.donatur.findData(data["donatur_id"])
            if index >= 0:
                self.donatur.setCurrentIndex(index)
        self.bukti = QLineEdit(data.get("bukti", "") if data else "")
        browse = QPushButton('Browse…')
        browse.clicked.connect(self.choose_file)
        evidence_row = QWidget()
        evidence_layout = QHBoxLayout(evidence_row)
        evidence_layout.setContentsMargins(0, 0, 0, 0)
        evidence_layout.addWidget(self.bukti, 1)
        evidence_layout.addWidget(browse)
        form.addRow('Date *', self.tanggal)
        form.addRow('Type *', self.jenis)
        form.addRow('Category *', self.kategori)
        form.addRow('Description *', self.deskripsi)
        form.addRow('Amount *', self.jumlah)
        form.addRow('Donor', self.donatur)
        form.addRow('Evidence', evidence_row)
        self.add_buttons(form)

    def choose_file(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            'Select Transaction Evidence',
            "",
            'Documents/Images (*.pdf *.png *.jpg *.jpeg *.webp);;All Files (*.*)',
        )
        if filename:
            self.bukti.setText(filename)

    def data(self) -> Dict[str, Any]:
        return {
            "tanggal": self.tanggal.dateTime().toString("yyyy-MM-dd HH:mm:ss"),
            "jenis": combo_data_or_text(self.jenis),
            "kategori": combo_data_or_text(self.kategori),
            "deskripsi": self.deskripsi.toPlainText(),
            "jumlah": self.jumlah.value(),
            "donatur_id": self.donatur.currentData(),
            "bukti": self.bukti.text(),
        }


class KegiatanDialog(BaseDialog):
    def __init__(self, data: Optional[Dict[str, Any]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle('Add Activity' if not data else 'Edit Activity')
        self.resize(560, 610)
        form = QFormLayout(self)
        self.nama = QLineEdit(data.get("nama", "") if data else "")
        self.mulai = QDateEdit()
        self.mulai.setCalendarPopup(True)
        self.mulai.setDisplayFormat(DATE_DISPLAY_FMT)
        self.mulai.setDate(qdate(data.get("tanggal_mulai", "") if data else today_iso()))
        self.selesai = QDateEdit()
        self.selesai.setCalendarPopup(True)
        self.selesai.setDisplayFormat(DATE_DISPLAY_FMT)
        self.selesai.setDate(
            qdate(data.get("tanggal_selesai", "") if data else (dt.date.today() + dt.timedelta(days=1)))
        )
        self.lokasi = QLineEdit(data.get("lokasi", "") if data else "")
        self.penanggung = QLineEdit(data.get("penanggung_jawab", "") if data else "")
        self.deskripsi = QTextEdit(data.get("deskripsi", "") if data else "")
        self.deskripsi.setMaximumHeight(90)
        self.anggaran = make_money_spin(data.get("anggaran", 0) if data else 0)
        self.realisasi = make_money_spin(data.get("realisasi", 0) if data else 0)
        self.status = QComboBox()
        add_enum_items(self.status, "activity_status", ["rencana", "berjalan", "selesai", "batal"])
        set_combo_data(self.status, data.get("status", "rencana") if data else "rencana")
        self.prioritas = QComboBox()
        add_enum_items(self.prioritas, "priority", ["rendah", "normal", "tinggi", "mendesak"])
        set_combo_data(self.prioritas, data.get("prioritas", "normal") if data else "normal")
        self.pengingat = QSpinBox()
        self.pengingat.setRange(0, 365)
        self.pengingat.setSuffix(' days before')
        self.pengingat.setValue(safe_int(data.get("pengingat_hari", 3), 3) if data else 3)
        form.addRow('Name *', self.nama)
        form.addRow('Start date *', self.mulai)
        form.addRow('End date *', self.selesai)
        form.addRow('Location', self.lokasi)
        form.addRow('Person in charge *', self.penanggung)
        form.addRow('Description *', self.deskripsi)
        form.addRow('Budget', self.anggaran)
        form.addRow('Actual', self.realisasi)
        form.addRow("Status", self.status)
        form.addRow('Priority', self.prioritas)
        form.addRow('Reminder', self.pengingat)
        self.add_buttons(form)

    def data(self) -> Dict[str, Any]:
        return {
            "nama": self.nama.text(),
            "tanggal_mulai": self.mulai.date().toString("yyyy-MM-dd"),
            "tanggal_selesai": self.selesai.date().toString("yyyy-MM-dd"),
            "lokasi": self.lokasi.text(),
            "penanggung_jawab": self.penanggung.text(),
            "deskripsi": self.deskripsi.toPlainText(),
            "anggaran": self.anggaran.value(),
            "realisasi": self.realisasi.value(),
            "status": combo_data_or_text(self.status),
            "prioritas": combo_data_or_text(self.prioritas),
            "pengingat_hari": self.pengingat.value(),
        }


class AsetDialog(BaseDialog):
    def __init__(self, data: Optional[Dict[str, Any]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle('Add Asset' if not data else 'Edit Asset')
        self.resize(560, 570)
        form = QFormLayout(self)
        self.nama = QLineEdit(data.get("nama", "") if data else "")
        self.kategori = QComboBox()
        self.kategori.setEditable(True)
        add_enum_items(self.kategori, "asset_category", ["bangunan", "tanah", "kendaraan", "elektronik", "peralatan", "furnitur", "lainnya"])
        if data:
            set_combo_data(self.kategori, data.get("kategori", "peralatan"))
        self.nilai = make_money_spin(data.get("nilai", 0) if data else 0)
        self.tanggal_perolehan = QDateEdit()
        self.tanggal_perolehan.setCalendarPopup(True)
        self.tanggal_perolehan.setDisplayFormat(DATE_DISPLAY_FMT)
        self.tanggal_perolehan.setDate(qdate(data.get("tanggal_perolehan", "") if data else today_iso()))
        self.kondisi = QComboBox()
        add_enum_items(self.kondisi, "asset_condition", ["baik", "perlu_perawatan", "rusak_ringan", "rusak_berat", "hilang"])
        set_combo_data(self.kondisi, data.get("kondisi", "baik") if data else "baik")
        self.lokasi = QLineEdit(data.get("lokasi", "") if data else "")
        self.last_maintenance = OptionalDateWidget(data.get("tanggal_perawatan_terakhir", "") if data else "")
        self.next_maintenance = OptionalDateWidget(data.get("tanggal_perawatan_berikut", "") if data else "")
        self.keterangan = QTextEdit(data.get("keterangan", "") if data else "")
        self.keterangan.setMaximumHeight(80)
        self.aktif = QCheckBox('Active / still owned')
        self.aktif.setChecked(bool(data.get("aktif", 1)) if data else True)
        form.addRow('Name *', self.nama)
        form.addRow('Category *', self.kategori)
        form.addRow('Value', self.nilai)
        form.addRow('Acquisition date', self.tanggal_perolehan)
        form.addRow('Condition *', self.kondisi)
        form.addRow('Location', self.lokasi)
        form.addRow('Last maintenance', self.last_maintenance)
        form.addRow('Next maintenance', self.next_maintenance)
        form.addRow('Description', self.keterangan)
        form.addRow("Status", self.aktif)
        self.add_buttons(form)

    def data(self) -> Dict[str, Any]:
        return {
            "nama": self.nama.text(),
            "kategori": combo_data_or_text(self.kategori),
            "nilai": self.nilai.value(),
            "tanggal_perolehan": self.tanggal_perolehan.date().toString("yyyy-MM-dd"),
            "kondisi": combo_data_or_text(self.kondisi),
            "lokasi": self.lokasi.text(),
            "tanggal_perawatan_terakhir": self.last_maintenance.value(),
            "tanggal_perawatan_berikut": self.next_maintenance.value(),
            "keterangan": self.keterangan.toPlainText(),
            "aktif": self.aktif.isChecked(),
        }


class RecurringDialog(BaseDialog):
    def __init__(
        self,
        donors: List[Dict[str, Any]],
        categories: List[str],
        data: Optional[Dict[str, Any]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle('Add Recurring Transaction' if not data else 'Edit Recurring Transaction')
        self.resize(560, 520)
        form = QFormLayout(self)
        self.nama = QLineEdit(data.get("nama", "") if data else "")
        self.jenis = QComboBox()
        add_enum_items(self.jenis, "transaction_type", ["pemasukan", "pengeluaran"])
        set_combo_data(self.jenis, data.get("jenis", "pengeluaran") if data else "pengeluaran")
        self.kategori = QComboBox()
        self.kategori.setEditable(True)
        self.kategori.addItems(categories or ['Operations', 'Electricity & Water', 'Donation'])
        if data:
            self.kategori.setCurrentText(data.get("kategori", ""))
        self.deskripsi = QTextEdit(data.get("deskripsi", "") if data else "")
        self.deskripsi.setMaximumHeight(80)
        self.jumlah = make_money_spin(data.get("jumlah", 0) if data else 0)
        self.donatur = QComboBox()
        self.donatur.addItem('— No linked donor —', None)
        for donor in donors:
            self.donatur.addItem(f"{donor['nama']} ({donor['id']})", donor["id"])
        if data and data.get("donatur_id"):
            index = self.donatur.findData(data["donatur_id"])
            if index >= 0:
                self.donatur.setCurrentIndex(index)
        interval_row = QWidget()
        interval_layout = QHBoxLayout(interval_row)
        interval_layout.setContentsMargins(0, 0, 0, 0)
        self.interval_value = QSpinBox()
        self.interval_value.setRange(1, 365)
        self.interval_value.setValue(safe_int(data.get("interval_nilai", 1), 1) if data else 1)
        self.interval_unit = QComboBox()
        add_enum_items(self.interval_unit, "interval_unit", ["hari", "minggu", "bulan", "tahun"])
        set_combo_data(self.interval_unit, data.get("interval_unit", "bulan") if data else "bulan")
        interval_layout.addWidget(QLabel('Every'))
        interval_layout.addWidget(self.interval_value)
        interval_layout.addWidget(self.interval_unit, 1)
        self.next_date = QDateEdit()
        self.next_date.setCalendarPopup(True)
        self.next_date.setDisplayFormat(DATE_DISPLAY_FMT)
        self.next_date.setDate(qdate(data.get("tanggal_berikut", "") if data else today_iso()))
        self.aktif = QCheckBox('Active')
        self.aktif.setChecked(bool(data.get("aktif", 1)) if data else True)
        form.addRow('Name *', self.nama)
        form.addRow('Type *', self.jenis)
        form.addRow('Category *', self.kategori)
        form.addRow('Description *', self.deskripsi)
        form.addRow('Amount *', self.jumlah)
        form.addRow('Donor', self.donatur)
        form.addRow('Interval', interval_row)
        form.addRow('Next date *', self.next_date)
        form.addRow("Status", self.aktif)
        self.add_buttons(form)

    def data(self) -> Dict[str, Any]:
        return {
            "nama": self.nama.text(),
            "jenis": combo_data_or_text(self.jenis),
            "kategori": combo_data_or_text(self.kategori),
            "deskripsi": self.deskripsi.toPlainText(),
            "jumlah": self.jumlah.value(),
            "donatur_id": self.donatur.currentData(),
            "interval_nilai": self.interval_value.value(),
            "interval_unit": combo_data_or_text(self.interval_unit),
            "tanggal_berikut": self.next_date.date().toString("yyyy-MM-dd"),
            "aktif": self.aktif.isChecked(),
        }


class UserDialog(BaseDialog):
    def __init__(self, data: Optional[Dict[str, Any]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle('Add User' if not data else 'Edit User')
        self.setFixedSize(470, 370)
        form = QFormLayout(self)
        self.username = QLineEdit(data.get("username", "") if data else "")
        self.username.setMaxLength(40)
        self.username.setEnabled(not bool(data))
        self.full_name = QLineEdit(data.get("nama_lengkap", "") if data else "")
        self.password = QLineEdit()
        self.password.setMaxLength(PASSWORD_MAX_LENGTH)
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        if data:
            self.password.setPlaceholderText('Leave blank to keep unchanged')
        self.role = QComboBox()
        self.role.addItems(["viewer", "operator", "admin"])
        self.role.setCurrentText(data.get("role", "viewer") if data else "viewer")
        self.aktif = QCheckBox('Active')
        self.aktif.setChecked(bool(data.get("aktif", 1)) if data else True)
        self.must_change = QCheckBox('Require password change at sign-in')
        self.must_change.setChecked(bool(data.get("wajib_ganti_password", 1)) if data else True)
        form.addRow("Username *", self.username)
        form.addRow('Full name', self.full_name)
        form.addRow("Password" if data else "Password *", self.password)
        form.addRow("Role", self.role)
        form.addRow("Status", self.aktif)
        form.addRow('Security', self.must_change)
        info = QLabel('Password: 1–6 letters/numbers only. • viewer: view/export • operator: add/edit • admin: full access')
        info.setWordWrap(True)
        form.addRow(info)
        self.add_buttons(form)

    def data(self) -> Dict[str, Any]:
        return {
            "username": self.username.text(),
            "nama_lengkap": self.full_name.text(),
            "password": self.password.text(),
            "role": self.role.currentText(),
            "aktif": self.aktif.isChecked(),
            "wajib_ganti_password": self.must_change.isChecked(),
        }


class SettingsDialog(BaseDialog):
    def __init__(self, settings: SettingsManager, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle('Application Settings')
        self.resize(620, 560)
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)

        profile_tab = QWidget()
        profile_form = QFormLayout(profile_tab)
        self.masjid_name = QLineEdit(settings.get("Masjid", "nama", "Mosque"))
        self.masjid_address = QTextEdit(settings.get("Masjid", "alamat", ""))
        self.masjid_address.setMaximumHeight(90)
        self.masjid_phone = QLineEdit(settings.get("Masjid", "telepon", ""))
        self.masjid_email = QLineEdit(settings.get("Masjid", "email", ""))
        self.chairman = QLineEdit(settings.get("Masjid", "ketua_dkm", ""))
        profile_form.addRow('Mosque name', self.masjid_name)
        profile_form.addRow('Address', self.masjid_address)
        profile_form.addRow('Phone', self.masjid_phone)
        profile_form.addRow("Email", self.masjid_email)
        profile_form.addRow('Board chair', self.chairman)
        tabs.addTab(profile_tab, 'Profile')

        system_tab = QWidget()
        system_form = QFormLayout(system_tab)
        self.db_path = QLineEdit(settings.get_db_path())
        db_button = QPushButton('Browse…')
        db_button.clicked.connect(self.choose_db)
        db_row = QWidget()
        db_layout = QHBoxLayout(db_row)
        db_layout.setContentsMargins(0, 0, 0, 0)
        db_layout.addWidget(self.db_path, 1)
        db_layout.addWidget(db_button)
        self.backup_dir = QLineEdit(settings.get_backup_dir())
        backup_button = QPushButton('Browse…')
        backup_button.clicked.connect(self.choose_backup)
        backup_row = QWidget()
        backup_layout = QHBoxLayout(backup_row)
        backup_layout.setContentsMargins(0, 0, 0, 0)
        backup_layout.addWidget(self.backup_dir, 1)
        backup_layout.addWidget(backup_button)
        self.auto_backup = QCheckBox('Active')
        self.auto_backup.setChecked(settings.get_bool("Database", "auto_backup", True))
        self.backup_interval = QSpinBox()
        self.backup_interval.setRange(1, 365)
        self.backup_interval.setSuffix(' days')
        self.backup_interval.setValue(settings.get_int("Database", "backup_interval_days", 1))
        self.max_backups = QSpinBox()
        self.max_backups.setRange(1, 500)
        self.max_backups.setValue(settings.get_int("Backup", "max_backups", 15))
        system_form.addRow("Database", db_row)
        system_form.addRow('Backup folder', backup_row)
        system_form.addRow('Automatic backup', self.auto_backup)
        system_form.addRow('Backup interval', self.backup_interval)
        system_form.addRow('Maximum backups', self.max_backups)
        warning = QLabel('A database location change takes effect after restarting the application.')
        warning.setWordWrap(True)
        warning.setStyleSheet("color: #f0b44d;")
        system_form.addRow(warning)
        tabs.addTab(system_tab, "System")

        localization_tab = QWidget()
        localization_form = QFormLayout(localization_tab)
        self.locale_name = QComboBox()
        self.locale_name.addItem("English (United States)", "en_US")
        self.currency = QComboBox()
        for code in CURRENCY_PROFILES:
            profile = CURRENCY_PROFILES[code]
            self.currency.addItem(f"{code} — {profile['symbol']}", code)
        set_combo_data(self.currency, settings.get("Localization", "currency", "USD"))
        sample = QLabel("Example: $1,234.56 • Aug 4, 2026")
        sample.setWordWrap(True)
        note = QLabel("Currency selection changes formatting only. Existing stored amounts are not converted.")
        note.setWordWrap(True)
        note.setStyleSheet("color: #f0b44d;")
        localization_form.addRow("Language / region", self.locale_name)
        localization_form.addRow("Display currency", self.currency)
        localization_form.addRow("en-US format", sample)
        localization_form.addRow(note)
        tabs.addTab(localization_tab, "Localization")

        self.add_buttons(layout)

    def choose_db(self) -> None:
        filename, _ = QFileDialog.getSaveFileName(
            self, 'Select Database', self.db_path.text(), "SQLite Database (*.db)"
        )
        if filename:
            self.db_path.setText(filename)

    def choose_backup(self) -> None:
        directory = QFileDialog.getExistingDirectory(
            self, 'Select Backup Folder', self.backup_dir.text()
        )
        if directory:
            self.backup_dir.setText(directory)

    def data(self) -> Dict[str, Dict[str, Any]]:
        return {
            "Masjid": {
                "nama": self.masjid_name.text().strip() or "Mosque",
                "alamat": self.masjid_address.toPlainText().strip(),
                "telepon": self.masjid_phone.text().strip(),
                "email": self.masjid_email.text().strip(),
                "ketua_dkm": self.chairman.text().strip(),
            },
            "Database": {
                "path": self.db_path.text().strip(),
                "auto_backup": self.auto_backup.isChecked(),
                "backup_interval_days": self.backup_interval.value(),
            },
            "Backup": {
                "backup_dir": self.backup_dir.text().strip(),
                "max_backups": self.max_backups.value(),
            },
            "Localization": {
                "locale": "en_US",
                "currency": combo_data_or_text(self.currency),
            },
        }

# ============================== MAIN WINDOW ==============================


class MainWindow(QMainWindow):
    def __init__(
        self,
        settings: SettingsManager,
        db: Database,
        auth: Dict[str, Any],
    ):
        super().__init__()
        self.settings = settings
        self.db = db
        self.manager = MasjidManager(db)
        self.manager.set_current_user(auth["username"], auth.get("role", "viewer"))
        self.auth = auth
        self.current_user = auth["username"]
        self.current_role = self.manager.current_role
        self.pages: Dict[str, int] = {
            "pengurus": 0,
            "donatur": 0,
            "transaksi": 0,
            "kegiatan": 0,
            "aset": 0,
            "rutin": 0,
        }
        self.page_sizes: Dict[str, int] = {key: 25 for key in self.pages}
        self.page_widgets: Dict[str, Dict[str, Any]] = {}
        self._closing = False
        app = QApplication.instance()
        if app is not None:
            self.setWindowIcon(app.windowIcon())
        self.setWindowTitle(
            f"{APP_NAME} {APP_VERSION} — {self.settings.get('Masjid', 'nama', 'Mosque')}"
        )
        self.setMinimumSize(1100, 700)
        self.restore_geometry()
        self.create_actions()
        self.create_menu()
        self.create_toolbar()
        self.create_tabs()
        self.create_status_bar()
        self.restore_layout()
        self.apply_theme()
        self.apply_permissions()
        self.refresh_all()
        QTimer.singleShot(500, self.run_startup_tasks)

    # ---------------- General UI ----------------
    def restore_geometry(self) -> None:
        """Restore geometry; center the window on first use."""
        width, height, x, y, state = self.settings.window_geometry()

        primary_screen = QApplication.primaryScreen()
        if primary_screen is None:
            self.resize(width, height)
            return

        if x is None or y is None:
            available = primary_screen.availableGeometry()
            width = min(width, available.width())
            height = min(height, available.height())
            x = available.x() + max(0, (available.width() - width) // 2)
            y = available.y() + max(0, (available.height() - height) // 2)
        else:
            # Use the screen containing the saved window center. If that monitor
            # is no longer connected, move the window to the primary screen.
            screen = QApplication.screenAt(QPoint(x + width // 2, y + height // 2))
            available = (screen or primary_screen).availableGeometry()
            width = min(width, available.width())
            height = min(height, available.height())
            max_x = available.x() + available.width() - width
            max_y = available.y() + available.height() - height
            x = min(max(x, available.x()), max_x)
            y = min(max(y, available.y()), max_y)

        self.setGeometry(x, y, width, height)
        if state == "maximized":
            QTimer.singleShot(0, self.showMaximized)

    def restore_layout(self) -> None:
        """Restore the QMainWindow layout after menus and toolbars exist."""
        encoded = str(self.settings.get("Window", "layout_state", "") or "").strip()
        if not encoded:
            return
        try:
            state = QByteArray.fromHex(encoded.encode("ascii"))
            self.restoreState(state, 1)
        except (ValueError, TypeError):
            # Konfigurasi layout yang rusak tidak boleh mencegah aplikasi terbuka.
            pass

    def create_actions(self) -> None:
        self.actions: Dict[str, QAction] = {}

        def add(name: str, text: str, slot, shortcut: Optional[str] = None) -> QAction:
            action = QAction(text, self)
            action.triggered.connect(slot)
            if shortcut:
                action.setShortcut(shortcut)
            self.actions[name] = action
            return action

        add("backup", "Backup Database", self.backup_database, "Ctrl+B")
        add("restore", "Restore Database", self.restore_database)
        add("export", "Export Data", self.export_data, "Ctrl+E")
        add("import", "Import Data", self.import_data)
        add("settings", 'Settings', self.open_settings)
        add("exit", 'Exit', self.close, "Ctrl+Q")
        add("add_transaction", 'Add Transaction', self.add_transaksi, "Ctrl+T")
        add("add_donor", 'Add Donor', self.add_donatur, "Ctrl+D")
        add("add_activity", 'Add Activity', self.add_kegiatan, "Ctrl+K")
        add("process_recurring", 'Process Recurring Transactions', self.process_recurring)
        add("integrity", 'Check Database Integrity', self.check_integrity)
        add("cleanup_logs", 'Clean Old Logs', self.cleanup_logs)
        add("users", 'Manage Users', self.manage_users)
        add("password", 'Change Password', self.change_password)
        add("about", "About", self.show_about)
        app = QApplication.instance()
        if app is not None:
            self.actions["about"].setIcon(app.windowIcon())
        add("refresh", 'Refresh All', self.refresh_all, "F5")

    def create_menu(self) -> None:
        file_menu = self.menuBar().addMenu("&File")
        file_menu.addAction(self.actions["backup"])
        file_menu.addAction(self.actions["restore"])
        file_menu.addSeparator()
        file_menu.addAction(self.actions["export"])
        file_menu.addAction(self.actions["import"])
        file_menu.addSeparator()
        file_menu.addAction(self.actions["settings"])
        file_menu.addSeparator()
        file_menu.addAction(self.actions["exit"])

        data_menu = self.menuBar().addMenu("&Data")
        data_menu.addAction(self.actions["add_transaction"])
        data_menu.addAction(self.actions["add_donor"])
        data_menu.addAction(self.actions["add_activity"])
        data_menu.addSeparator()
        data_menu.addAction(self.actions["refresh"])

        tools_menu = self.menuBar().addMenu("&Tools")
        tools_menu.addAction(self.actions["process_recurring"])
        tools_menu.addAction(self.actions["integrity"])
        tools_menu.addAction(self.actions["cleanup_logs"])

        account_menu = self.menuBar().addMenu('&Account')
        account_menu.addAction(self.actions["password"])
        account_menu.addAction(self.actions["users"])

        help_menu = self.menuBar().addMenu('&Help')
        help_menu.addAction(self.actions["about"])

    def create_toolbar(self) -> None:
        toolbar = QToolBar('Quick Access', self)
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        toolbar.addAction(self.actions["add_transaction"])
        toolbar.addAction(self.actions["add_donor"])
        toolbar.addAction(self.actions["add_activity"])
        toolbar.addSeparator()
        toolbar.addAction(self.actions["backup"])
        toolbar.addAction(self.actions["refresh"])
        self.addToolBar(toolbar)

    def create_status_bar(self) -> None:
        self.user_status = QLabel(f"👤 {self.current_user} ({self.current_role})")
        self.notification_status = QLabel("")
        self.statusBar().addPermanentWidget(self.notification_status)
        self.statusBar().addPermanentWidget(self.user_status)
        self.statusBar().showMessage('Ready')

    def create_tabs(self) -> None:
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        # Connect the signal only after all tabs are created. The first addTab()
        # can otherwise trigger on_tab_changed() before Dashboard exists.
        self.setCentralWidget(self.tabs)

        self.dashboard_tab = QWidget()
        self.transaction_tab = QWidget()
        self.donor_tab = QWidget()
        self.committee_tab = QWidget()
        self.activity_tab = QWidget()
        self.asset_tab = QWidget()
        self.recurring_tab = QWidget()
        self.report_tab = QWidget()
        self.log_tab = QWidget()

        self.tabs.addTab(self.dashboard_tab, "Dashboard")
        self.tabs.addTab(self.transaction_tab, 'Finance')
        self.tabs.addTab(self.donor_tab, 'Donor')
        self.tabs.addTab(self.committee_tab, 'Committee')
        self.tabs.addTab(self.activity_tab, 'Activities')
        self.tabs.addTab(self.asset_tab, 'Assets')
        self.tabs.addTab(self.recurring_tab, 'Recurring Transactions')
        self.tabs.addTab(self.report_tab, 'Reports')
        self.tabs.addTab(self.log_tab, "Log")

        self.init_dashboard_tab()
        self.init_transaction_tab()
        self.init_donor_tab()
        self.init_committee_tab()
        self.init_activity_tab()
        self.init_asset_tab()
        self.init_recurring_tab()
        self.init_report_tab()
        self.init_log_tab()

        # Aman dipasang setelah semua atribut tab (dashboard_cards, tabel, dll.) tersedia.
        self.tabs.currentChanged.connect(self.on_tab_changed)

    def apply_permissions(self) -> None:
        self.actions["backup"].setEnabled(self.manager.can("backup"))
        self.actions["export"].setEnabled(self.manager.can("export"))
        self.actions["restore"].setEnabled(self.manager.can("restore"))
        self.actions["import"].setEnabled(self.manager.can("import"))
        self.actions["settings"].setEnabled(self.manager.can("settings"))
        self.actions["users"].setEnabled(self.manager.can("users"))
        self.actions["cleanup_logs"].setEnabled(self.manager.can("delete"))
        self.actions["add_transaction"].setEnabled(self.manager.can("add"))
        self.actions["add_donor"].setEnabled(self.manager.can("add"))
        self.actions["add_activity"].setEnabled(self.manager.can("add"))
        self.actions["process_recurring"].setEnabled(self.manager.can("add"))
        for button in getattr(self, "write_buttons", []):
            permission = button.property("permission") or "add"
            button.setEnabled(self.manager.can(permission))

    def apply_theme(self) -> None:
        """Terapkan dark theme permanen; tidak tersedia mode terang."""
        app = QApplication.instance()
        if app is not None:
            apply_dark_theme(app)
        if HAS_MATPLOTLIB and hasattr(self, "dashboard_figure"):
            self.refresh_dashboard()

    def configure_table(self, table: QTableWidget, headers: Sequence[str], stretch_column: Optional[int] = None) -> None:
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(list(headers))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(False)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        if stretch_column is not None:
            table.horizontalHeader().setSectionResizeMode(stretch_column, QHeaderView.ResizeMode.Stretch)

    def create_pager(self, key: str, refresh_callback, layout: QHBoxLayout) -> None:
        previous = QPushButton('‹ Previous')
        next_button = QPushButton('Next ›')
        label = QLabel('Page 1')
        size = QComboBox()
        size.addItems(["10", "25", "50", "100"])
        size.setCurrentText(str(self.page_sizes[key]))
        previous.clicked.connect(lambda: self.change_page(key, -1, refresh_callback))
        next_button.clicked.connect(lambda: self.change_page(key, 1, refresh_callback))
        size.currentTextChanged.connect(
            lambda value: self.change_page_size(key, int(value), refresh_callback)
        )
        layout.addStretch()
        layout.addWidget(QLabel('Rows:'))
        layout.addWidget(size)
        layout.addWidget(previous)
        layout.addWidget(label)
        layout.addWidget(next_button)
        self.page_widgets[key] = {
            "prev": previous,
            "next": next_button,
            "label": label,
            "size": size,
        }

    def change_page(self, key: str, delta: int, refresh_callback) -> None:
        self.pages[key] = max(0, self.pages[key] + delta)
        refresh_callback()

    def change_page_size(self, key: str, size: int, refresh_callback) -> None:
        self.page_sizes[key] = size
        self.pages[key] = 0
        refresh_callback()

    def update_pager(self, key: str, total: int) -> None:
        size = self.page_sizes[key]
        max_page = max(0, (total - 1) // size)
        self.pages[key] = min(self.pages[key], max_page)
        current = self.pages[key]
        widgets = self.page_widgets[key]
        widgets["prev"].setEnabled(current > 0)
        widgets["next"].setEnabled(current < max_page)
        widgets["label"].setText(f"Page {current + 1}/{max_page + 1} • {total} records")

    def reset_page_and_refresh(self, key: str, callback) -> None:
        self.pages[key] = 0
        callback()

    def make_button(self, text: str, slot, permission: Optional[str] = None) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(slot)
        if permission:
            button.setProperty("permission", permission)
            if not hasattr(self, "write_buttons"):
                self.write_buttons: List[QPushButton] = []
            self.write_buttons.append(button)
        return button

    def on_tab_changed(self, index: int) -> None:
        widget = self.tabs.widget(index)
        callbacks = {
            self.dashboard_tab: self.refresh_dashboard,
            self.transaction_tab: self.refresh_transaksi,
            self.donor_tab: self.refresh_donatur,
            self.committee_tab: self.refresh_pengurus,
            self.activity_tab: self.refresh_kegiatan,
            self.asset_tab: self.refresh_aset,
            self.recurring_tab: self.refresh_recurring,
            self.report_tab: self.refresh_report,
            self.log_tab: self.refresh_logs,
        }
        callback = callbacks.get(widget)
        if callback:
            callback()

    def handle_error(self, title: str, exc: Exception) -> None:
        QMessageBox.critical(self, title, str(exc))
        self.statusBar().showMessage(f"Failed: {exc}", 7000)

    def confirm_delete(self, label: str) -> bool:
        return (
            QMessageBox.question(
                self,
                'Confirmation',
                f"Are you sure you want to delete/deactivate {label}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            == QMessageBox.StandardButton.Yes
        )

    # ---------------- Dashboard ----------------
    def init_dashboard_tab(self) -> None:
        layout = QVBoxLayout(self.dashboard_tab)
        header = QHBoxLayout()
        self.dashboard_title = QLabel(
            f"🏛️ {self.settings.get('Masjid', 'nama', 'Mosque')}"
        )
        self.dashboard_title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        header.addWidget(self.dashboard_title)
        header.addStretch()
        refresh = QPushButton('Refresh Dashboard')
        refresh.clicked.connect(self.refresh_dashboard)
        header.addWidget(refresh)
        layout.addLayout(header)

        cards = QGridLayout()
        self.dashboard_cards: Dict[str, QLabel] = {}
        card_specs = [
            ("saldo", 'Overall Balance'),
            ("pemasukan_bulan", 'Income This Month'),
            ("pengeluaran_bulan", 'Expenses This Month'),
            ("donatur_aktif", 'Active Donors'),
            ("pengurus_aktif", 'Active Committee Members'),
            ("kegiatan_mendatang", 'Upcoming Activities'),
            ("nilai_aset", 'Active Asset Value'),
            ("perawatan_jatuh_tempo", 'Maintenance Due ≤30 Days'),
        ]
        for index, (key, title) in enumerate(card_specs):
            box = QGroupBox(title)
            box_layout = QVBoxLayout(box)
            value = QLabel("—")
            value.setAlignment(Qt.AlignmentFlag.AlignCenter)
            value.setFont(QFont("Segoe UI", 15, QFont.Weight.Bold))
            box_layout.addWidget(value)
            cards.addWidget(box, index // 4, index % 4)
            self.dashboard_cards[key] = value
        layout.addLayout(cards)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        chart_container = QGroupBox('12-Month Cash Flow')
        chart_layout = QVBoxLayout(chart_container)
        if HAS_MATPLOTLIB:
            self.dashboard_figure = Figure(figsize=(8, 4), tight_layout=True)
            self.dashboard_canvas = FigureCanvas(self.dashboard_figure)
            chart_layout.addWidget(self.dashboard_canvas)
        else:
            chart_layout.addWidget(QLabel('Install matplotlib to display the chart.'))
        splitter.addWidget(chart_container)

        info_tabs = QTabWidget()
        recent_page = QWidget()
        recent_layout = QVBoxLayout(recent_page)
        self.recent_table = QTableWidget()
        self.configure_table(
            self.recent_table,
            ["ID", 'Date', 'Type', 'Description', 'Amount'],
            stretch_column=3,
        )
        recent_layout.addWidget(self.recent_table)
        info_tabs.addTab(recent_page, 'Recent Transactions')

        upcoming_page = QWidget()
        upcoming_layout = QVBoxLayout(upcoming_page)
        self.upcoming_table = QTableWidget()
        self.configure_table(
            self.upcoming_table,
            ['Date', 'Activities', 'Location', 'Priority'],
            stretch_column=1,
        )
        upcoming_layout.addWidget(self.upcoming_table)
        info_tabs.addTab(upcoming_page, 'Upcoming Agenda')

        maintenance_page = QWidget()
        maintenance_layout = QVBoxLayout(maintenance_page)
        self.maintenance_table = QTableWidget()
        self.configure_table(
            self.maintenance_table,
            ['Schedule', 'Asset', 'Location', "Condition"],
            stretch_column=1,
        )
        maintenance_layout.addWidget(self.maintenance_table)
        info_tabs.addTab(maintenance_page, 'Asset Maintenance')
        splitter.addWidget(info_tabs)
        splitter.setSizes([700, 500])
        layout.addWidget(splitter, 1)

    def refresh_dashboard(self) -> None:
        try:
            summary = self.manager.dashboard_summary()
            currency_keys = {"saldo", "pemasukan_bulan", "pengeluaran_bulan", "nilai_aset"}
            for key, label in self.dashboard_cards.items():
                label.setText(fmt_currency(summary[key]) if key in currency_keys else str(summary[key]))
            recent = self.manager.recent_transactions(limit=None)
            self.recent_table.setRowCount(len(recent))
            for row_index, row in enumerate(recent):
                values = [
                    row["id"],
                    fmt_datetime(row["tanggal"]),
                    enum_label("transaction_type", row["jenis"]),
                    elide(row["deskripsi"], 50),
                    fmt_currency(row["jumlah"]),
                ]
                for column, value in enumerate(values):
                    set_table_item(self.recent_table, row_index, column, value)

            upcoming = self.manager.upcoming_activities(days=None, limit=None, newest_first=True)
            self.upcoming_table.setRowCount(len(upcoming))
            for row_index, row in enumerate(upcoming):
                values = [fmt_date(row["tanggal_mulai"]), row["nama"], row["lokasi"], enum_label("priority", row["prioritas"])]
                for column, value in enumerate(values):
                    set_table_item(self.upcoming_table, row_index, column, value)

            maintenance = self.manager.maintenance_due(days=None, limit=None, newest_first=True)
            self.maintenance_table.setRowCount(len(maintenance))
            for row_index, row in enumerate(maintenance):
                values = [
                    fmt_date(row["tanggal_perawatan_berikut"]),
                    row["nama"],
                    row["lokasi"],
                    enum_label("asset_condition", row["kondisi"]),
                ]
                for column, value in enumerate(values):
                    set_table_item(self.maintenance_table, row_index, column, value)

            if HAS_MATPLOTLIB:
                monthly = self.manager.monthly_cashflow(12)
                self.dashboard_figure.clear()
                axis = self.dashboard_figure.add_subplot(111)
                labels = [item["label"] for item in monthly]
                income = [item["pemasukan"] for item in monthly]
                expense = [item["pengeluaran"] for item in monthly]
                positions = list(range(len(labels)))
                width = 0.38
                axis.bar([x - width / 2 for x in positions], income, width, label="Income")
                axis.bar([x + width / 2 for x in positions], expense, width, label="Expenses")
                axis.set_xticks(positions)
                axis.set_xticklabels(labels, rotation=35, ha="right")
                axis.set_ylabel(_CURRENT_CURRENCY)
                axis.legend()
                axis.grid(axis="y", alpha=0.2)
                self.dashboard_figure.patch.set_facecolor("#20252b")
                axis.set_facecolor("#20252b")
                axis.tick_params(colors="#eef2f5")
                axis.xaxis.label.set_color("#eef2f5")
                axis.yaxis.label.set_color("#eef2f5")
                for spine in axis.spines.values():
                    spine.set_color("#87929d")
                legend = axis.get_legend()
                if legend:
                    legend.get_frame().set_facecolor("#293038")
                    legend.get_frame().set_edgecolor("#52606d")
                    for legend_text in legend.get_texts():
                        legend_text.set_color("#eef2f5")
                self.dashboard_canvas.draw_idle()
            self.notification_status.setText(
                f"🔔 {len(upcoming)} agenda items • {len(maintenance)} maintenance items"
            )
            self.statusBar().showMessage('Dashboard updated', 3000)
        except Exception as exc:
            self.handle_error('Failed to load dashboard', exc)

    # ---------------- Keuangan ----------------
    def init_transaction_tab(self) -> None:
        layout = QVBoxLayout(self.transaction_tab)
        controls = QHBoxLayout()
        controls.addWidget(self.make_button('Add', self.add_transaksi, "add"))
        controls.addWidget(self.make_button('Edit', self.edit_transaksi, "edit"))
        controls.addWidget(self.make_button('Delete', self.delete_transaksi, "delete"))
        open_button = QPushButton('Open Evidence')
        open_button.clicked.connect(self.open_transaction_evidence)
        controls.addWidget(open_button)
        controls.addWidget(QLabel('Search:'))
        self.transaction_search = QLineEdit()
        self.transaction_search.setPlaceholderText('ID, category, description, donor…')
        self.transaction_search.setClearButtonEnabled(True)
        self.transaction_search.textChanged.connect(
            lambda: self.reset_page_and_refresh("transaksi", self.refresh_transaksi)
        )
        controls.addWidget(self.transaction_search, 1)
        self.transaction_type = QComboBox()
        add_enum_items(self.transaction_type, "transaction_type", ["semua", "pemasukan", "pengeluaran"])
        self.transaction_type.setItemText(0, "All")
        self.transaction_type.currentTextChanged.connect(
            lambda: self.reset_page_and_refresh("transaksi", self.refresh_transaksi)
        )
        controls.addWidget(self.transaction_type)
        layout.addLayout(controls)

        date_controls = QHBoxLayout()
        date_controls.addWidget(QLabel('Period:'))
        self.transaction_start = QDateEdit()
        self.transaction_start.setCalendarPopup(True)
        self.transaction_start.setDisplayFormat(DATE_DISPLAY_FMT)
        self.transaction_start.setDate(QDate(dt.date.today().year, 1, 1))
        self.transaction_end = QDateEdit()
        self.transaction_end.setCalendarPopup(True)
        self.transaction_end.setDisplayFormat(DATE_DISPLAY_FMT)
        self.transaction_end.setDate(QDate.currentDate())
        self.transaction_start.dateChanged.connect(
            lambda: self.reset_page_and_refresh("transaksi", self.refresh_transaksi)
        )
        self.transaction_end.dateChanged.connect(
            lambda: self.reset_page_and_refresh("transaksi", self.refresh_transaksi)
        )
        date_controls.addWidget(self.transaction_start)
        date_controls.addWidget(QLabel('to'))
        date_controls.addWidget(self.transaction_end)
        month_button = QPushButton('This Month')
        month_button.clicked.connect(self.set_transaction_current_month)
        year_button = QPushButton('This Year')
        year_button.clicked.connect(self.set_transaction_current_year)
        date_controls.addWidget(month_button)
        date_controls.addWidget(year_button)
        self.transaction_summary = QLabel("")
        self.transaction_summary.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        date_controls.addStretch()
        date_controls.addWidget(self.transaction_summary)
        layout.addLayout(date_controls)

        self.transaction_table = QTableWidget()
        self.configure_table(
            self.transaction_table,
            ["ID", 'Date', 'Type', 'Category', 'Description', 'Donor', 'Amount', 'Evidence', 'Created By'],
            stretch_column=4,
        )
        self.transaction_table.doubleClicked.connect(self.edit_transaksi)
        self.transaction_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.transaction_table.customContextMenuRequested.connect(self.transaction_context_menu)
        layout.addWidget(self.transaction_table, 1)
        pager = QHBoxLayout()
        self.create_pager("transaksi", self.refresh_transaksi, pager)
        layout.addLayout(pager)

    def set_transaction_current_month(self) -> None:
        today = dt.date.today()
        start, end = month_bounds(today.year, today.month)
        self.transaction_start.setDate(qdate(start))
        self.transaction_end.setDate(qdate(end))

    def set_transaction_current_year(self) -> None:
        today = dt.date.today()
        self.transaction_start.setDate(QDate(today.year, 1, 1))
        self.transaction_end.setDate(QDate(today.year, 12, 31))

    def transaction_context_menu(self, point) -> None:
        menu = QMenu(self)
        edit_action = menu.addAction('Edit')
        evidence_action = menu.addAction('Open Evidence')
        delete_action = menu.addAction('Delete')
        edit_action.setEnabled(self.manager.can("edit"))
        delete_action.setEnabled(self.manager.can("delete"))
        selected = menu.exec(self.transaction_table.viewport().mapToGlobal(point))
        if selected == edit_action:
            self.edit_transaksi()
        elif selected == evidence_action:
            self.open_transaction_evidence()
        elif selected == delete_action:
            self.delete_transaksi()

    def refresh_transaksi(self) -> None:
        try:
            start = self.transaction_start.date().toString("yyyy-MM-dd")
            end = self.transaction_end.date().toString("yyyy-MM-dd")
            if start > end:
                return
            size = self.page_sizes["transaksi"]
            offset = self.pages["transaksi"] * size
            rows, total = self.manager.page_transaksi(
                self.transaction_search.text().strip(),
                combo_data_or_text(self.transaction_type),
                start,
                end,
                size,
                offset,
            )
            if not rows and total and self.pages["transaksi"] > 0:
                self.pages["transaksi"] -= 1
                return self.refresh_transaksi()
            self.transaction_table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                values = [
                    row["id"],
                    fmt_datetime(row["tanggal"]),
                    enum_label("transaction_type", row["jenis"]),
                    row["kategori"],
                    elide(row["deskripsi"], 60),
                    row["donatur_nama"],
                    fmt_currency(row["jumlah"]),
                    os.path.basename(row["bukti"]) if row["bukti"] else "",
                    row["created_by"],
                ]
                for column, value in enumerate(values):
                    set_table_item(self.transaction_table, row_index, column, value)
            totals = self.manager.finance_totals(start, end)
            self.transaction_summary.setText(
                f"Income {fmt_currency(totals['pemasukan'])}  |  Expenses {fmt_currency(totals['pengeluaran'])}  |  Net {fmt_currency(totals['saldo'])}"
            )
            self.update_pager("transaksi", total)
        except Exception as exc:
            self.handle_error('Failed to load transactions', exc)

    def add_transaksi(self) -> None:
        try:
            self.manager.require("add")
            dialog = TransaksiDialog(
                self.manager.list_donatur(),
                self.manager.transaction_categories(),
                parent=self,
            )
            if dialog.exec() == QDialog.DialogCode.Accepted:
                item_id = self.manager.save_transaksi(dialog.data())
                QMessageBox.information(self, 'Success', f"Transaction {item_id} was added.")
                self.refresh_transaksi()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to add transaction', exc)

    def edit_transaksi(self) -> None:
        item_id = selected_id(self.transaction_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a transaction to edit.')
            return
        try:
            self.manager.require("edit")
            data = self.manager.get_transaksi(item_id)
            if not data:
                raise ValueError('Transaction not found.')
            dialog = TransaksiDialog(
                self.manager.list_donatur(False),
                self.manager.transaction_categories(),
                data,
                self,
            )
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.manager.save_transaksi(dialog.data(), item_id)
                self.refresh_transaksi()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to edit transaction', exc)

    def delete_transaksi(self) -> None:
        item_id = selected_id(self.transaction_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a transaction to delete.')
            return
        if not self.confirm_delete(f"transaction {item_id}"):
            return
        try:
            self.manager.delete_transaksi(item_id)
            self.refresh_transaksi()
            self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to delete transaction', exc)

    def open_transaction_evidence(self) -> None:
        item_id = selected_id(self.transaction_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a transaction first.')
            return
        data = self.manager.get_transaksi(item_id)
        path = data.get("bukti") if data else ""
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, 'Evidence', 'The evidence file was not found or has not been provided.')
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(path)))

    # ---------------- Donatur ----------------
    def init_donor_tab(self) -> None:
        layout = QVBoxLayout(self.donor_tab)
        controls = QHBoxLayout()
        controls.addWidget(self.make_button('Add', self.add_donatur, "add"))
        controls.addWidget(self.make_button('Edit', self.edit_donatur, "edit"))
        controls.addWidget(self.make_button('Deactivate', self.delete_donatur, "delete"))
        controls.addWidget(QLabel('Search:'))
        self.donor_search = QLineEdit()
        self.donor_search.setPlaceholderText('ID, name, phone, email…')
        self.donor_search.setClearButtonEnabled(True)
        self.donor_search.textChanged.connect(
            lambda: self.reset_page_and_refresh("donatur", self.refresh_donatur)
        )
        controls.addWidget(self.donor_search, 1)
        self.donor_status = QComboBox()
        add_enum_items(self.donor_status, "status_filter", ["semua", "aktif", "nonaktif"])
        set_combo_data(self.donor_status, "aktif")
        self.donor_status.currentTextChanged.connect(
            lambda: self.reset_page_and_refresh("donatur", self.refresh_donatur)
        )
        controls.addWidget(self.donor_status)
        layout.addLayout(controls)
        self.donor_table = QTableWidget()
        self.configure_table(
            self.donor_table,
            ["ID", 'Name', 'Phone', "Email", 'Category', "Status", 'Total Donations', 'Last Donation'],
            stretch_column=1,
        )
        self.donor_table.doubleClicked.connect(self.edit_donatur)
        layout.addWidget(self.donor_table, 1)
        pager = QHBoxLayout()
        self.create_pager("donatur", self.refresh_donatur, pager)
        layout.addLayout(pager)

    def refresh_donatur(self) -> None:
        try:
            size = self.page_sizes["donatur"]
            rows, total = self.manager.page_donatur(
                self.donor_search.text().strip(),
                combo_data_or_text(self.donor_status),
                size,
                self.pages["donatur"] * size,
            )
            if not rows and total and self.pages["donatur"] > 0:
                self.pages["donatur"] -= 1
                return self.refresh_donatur()
            self.donor_table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                values = [
                    row["id"],
                    row["nama"],
                    row["no_hp"],
                    row["email"],
                    enum_label("donor_category", row["kategori"]),
                    "Active" if row["aktif"] else "Inactive",
                    fmt_currency(row["total_donasi"]),
                    fmt_date(row["donasi_terakhir"] or ""),
                ]
                for column, value in enumerate(values):
                    set_table_item(self.donor_table, row_index, column, value)
            self.update_pager("donatur", total)
        except Exception as exc:
            self.handle_error('Failed to load donors', exc)

    def add_donatur(self) -> None:
        try:
            dialog = DonaturDialog(parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                item_id = self.manager.save_donatur(dialog.data())
                QMessageBox.information(self, 'Success', f"Donor {item_id} was added.")
                self.refresh_donatur()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to add donor', exc)

    def edit_donatur(self) -> None:
        item_id = selected_id(self.donor_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a donor to edit.')
            return
        try:
            data = self.manager.get_donatur(item_id)
            dialog = DonaturDialog(data, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.manager.save_donatur(dialog.data(), item_id)
                self.refresh_donatur()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to edit donor', exc)

    def delete_donatur(self) -> None:
        item_id = selected_id(self.donor_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a donor to deactivate.')
            return
        if not self.confirm_delete(f"donor {item_id}"):
            return
        try:
            self.manager.delete_donatur(item_id)
            self.refresh_donatur()
            self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to deactivate donor', exc)

    # ---------------- Pengurus ----------------
    def init_committee_tab(self) -> None:
        layout = QVBoxLayout(self.committee_tab)
        controls = QHBoxLayout()
        controls.addWidget(self.make_button('Add', self.add_pengurus, "add"))
        controls.addWidget(self.make_button('Edit', self.edit_pengurus, "edit"))
        controls.addWidget(self.make_button('Deactivate', self.delete_pengurus, "delete"))
        controls.addWidget(QLabel('Search:'))
        self.committee_search = QLineEdit()
        self.committee_search.setPlaceholderText('ID, name, position, phone…')
        self.committee_search.setClearButtonEnabled(True)
        self.committee_search.textChanged.connect(
            lambda: self.reset_page_and_refresh("pengurus", self.refresh_pengurus)
        )
        controls.addWidget(self.committee_search, 1)
        self.committee_status = QComboBox()
        add_enum_items(self.committee_status, "status_filter", ["semua", "aktif", "nonaktif"])
        set_combo_data(self.committee_status, "aktif")
        self.committee_status.currentTextChanged.connect(
            lambda: self.reset_page_and_refresh("pengurus", self.refresh_pengurus)
        )
        controls.addWidget(self.committee_status)
        layout.addLayout(controls)
        self.committee_table = QTableWidget()
        self.configure_table(
            self.committee_table,
            ["ID", 'Name', 'Position', 'Phone', "Email", 'Address', "Status", 'Joined'],
            stretch_column=5,
        )
        self.committee_table.doubleClicked.connect(self.edit_pengurus)
        layout.addWidget(self.committee_table, 1)
        pager = QHBoxLayout()
        self.create_pager("pengurus", self.refresh_pengurus, pager)
        layout.addLayout(pager)

    def refresh_pengurus(self) -> None:
        try:
            size = self.page_sizes["pengurus"]
            rows, total = self.manager.page_pengurus(
                self.committee_search.text().strip(),
                combo_data_or_text(self.committee_status),
                size,
                self.pages["pengurus"] * size,
            )
            if not rows and total and self.pages["pengurus"] > 0:
                self.pages["pengurus"] -= 1
                return self.refresh_pengurus()
            self.committee_table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                values = [
                    row["id"],
                    row["nama"],
                    row["jabatan"],
                    row["no_hp"],
                    row["email"],
                    elide(row["alamat"], 55),
                    'Active' if row["aktif"] else "Inactive",
                    fmt_date(row["tanggal_bergabung"]),
                ]
                for column, value in enumerate(values):
                    set_table_item(self.committee_table, row_index, column, value)
            self.update_pager("pengurus", total)
        except Exception as exc:
            self.handle_error('Failed to load committee members', exc)

    def add_pengurus(self) -> None:
        try:
            dialog = PengurusDialog(parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                item_id = self.manager.save_pengurus(dialog.data())
                QMessageBox.information(self, 'Success', f"Committee member {item_id} was added.")
                self.refresh_pengurus()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to add committee member', exc)

    def edit_pengurus(self) -> None:
        item_id = selected_id(self.committee_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a committee member to edit.')
            return
        try:
            dialog = PengurusDialog(self.manager.get_pengurus(item_id), self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.manager.save_pengurus(dialog.data(), item_id)
                self.refresh_pengurus()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to edit committee member', exc)

    def delete_pengurus(self) -> None:
        item_id = selected_id(self.committee_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a committee member to deactivate.')
            return
        if not self.confirm_delete(f"committee member {item_id}"):
            return
        try:
            self.manager.delete_pengurus(item_id)
            self.refresh_pengurus()
            self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to deactivate committee member', exc)

    # ---------------- Kegiatan ----------------
    def init_activity_tab(self) -> None:
        layout = QVBoxLayout(self.activity_tab)
        controls = QHBoxLayout()
        controls.addWidget(self.make_button('Add', self.add_kegiatan, "add"))
        controls.addWidget(self.make_button('Edit', self.edit_kegiatan, "edit"))
        controls.addWidget(self.make_button('Delete', self.delete_kegiatan, "delete"))
        controls.addWidget(QLabel('Search:'))
        self.activity_search = QLineEdit()
        self.activity_search.setPlaceholderText('ID, name, person in charge, location…')
        self.activity_search.setClearButtonEnabled(True)
        self.activity_search.textChanged.connect(
            lambda: self.reset_page_and_refresh("kegiatan", self.refresh_kegiatan)
        )
        controls.addWidget(self.activity_search, 1)
        self.activity_status = QComboBox()
        add_enum_items(self.activity_status, "activity_status", ["semua", "rencana", "berjalan", "selesai", "batal"])
        self.activity_status.currentTextChanged.connect(
            lambda: self.reset_page_and_refresh("kegiatan", self.refresh_kegiatan)
        )
        controls.addWidget(self.activity_status)
        layout.addLayout(controls)
        self.activity_table = QTableWidget()
        self.configure_table(
            self.activity_table,
            [
                "ID",
                'Name',
                'Start',
                'End',
                'Location',
                'Person in Charge',
                "Status",
                'Priority',
                'Budget',
                'Actual',
                'Variance',
            ],
            stretch_column=1,
        )
        self.activity_table.doubleClicked.connect(self.edit_kegiatan)
        layout.addWidget(self.activity_table, 1)
        pager = QHBoxLayout()
        self.create_pager("kegiatan", self.refresh_kegiatan, pager)
        layout.addLayout(pager)

    def refresh_kegiatan(self) -> None:
        try:
            size = self.page_sizes["kegiatan"]
            rows, total = self.manager.page_kegiatan(
                self.activity_search.text().strip(),
                combo_data_or_text(self.activity_status),
                size,
                self.pages["kegiatan"] * size,
            )
            if not rows and total and self.pages["kegiatan"] > 0:
                self.pages["kegiatan"] -= 1
                return self.refresh_kegiatan()
            self.activity_table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                values = [
                    row["id"],
                    row["nama"],
                    fmt_date(row["tanggal_mulai"]),
                    fmt_date(row["tanggal_selesai"]),
                    row["lokasi"],
                    row["penanggung_jawab"],
                    enum_label("activity_status", row["status"]),
                    enum_label("priority", row["prioritas"]),
                    fmt_currency(row["anggaran"]),
                    fmt_currency(row["realisasi"]),
                    fmt_currency(float(row["anggaran"]) - float(row["realisasi"])),
                ]
                for column, value in enumerate(values):
                    set_table_item(self.activity_table, row_index, column, value)
            self.update_pager("kegiatan", total)
        except Exception as exc:
            self.handle_error('Failed to load activities', exc)

    def add_kegiatan(self) -> None:
        try:
            dialog = KegiatanDialog(parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                item_id = self.manager.save_kegiatan(dialog.data())
                QMessageBox.information(self, 'Success', f"Activity {item_id} was added.")
                self.refresh_kegiatan()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to add activity', exc)

    def edit_kegiatan(self) -> None:
        item_id = selected_id(self.activity_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select an activity to edit.')
            return
        try:
            dialog = KegiatanDialog(self.manager.get_kegiatan(item_id), self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.manager.save_kegiatan(dialog.data(), item_id)
                self.refresh_kegiatan()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to edit activity', exc)

    def delete_kegiatan(self) -> None:
        item_id = selected_id(self.activity_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select an activity to delete.')
            return
        if not self.confirm_delete(f"activity {item_id}"):
            return
        try:
            self.manager.delete_kegiatan(item_id)
            self.refresh_kegiatan()
            self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to delete activity', exc)

    # ---------------- Aset ----------------
    def init_asset_tab(self) -> None:
        layout = QVBoxLayout(self.asset_tab)
        controls = QHBoxLayout()
        controls.addWidget(self.make_button('Add', self.add_aset, "add"))
        controls.addWidget(self.make_button('Edit', self.edit_aset, "edit"))
        controls.addWidget(self.make_button('Deactivate', self.delete_aset, "delete"))
        controls.addWidget(QLabel('Search:'))
        self.asset_search = QLineEdit()
        self.asset_search.setPlaceholderText('ID, name, location, description…')
        self.asset_search.setClearButtonEnabled(True)
        self.asset_search.textChanged.connect(
            lambda: self.reset_page_and_refresh("aset", self.refresh_aset)
        )
        controls.addWidget(self.asset_search, 1)
        self.asset_category = QComboBox()
        add_enum_items(self.asset_category, "asset_category", ["semua", "bangunan", "tanah", "kendaraan", "elektronik", "peralatan", "furnitur", "lainnya"])
        self.asset_category.currentTextChanged.connect(
            lambda: self.reset_page_and_refresh("aset", self.refresh_aset)
        )
        self.asset_condition = QComboBox()
        add_enum_items(self.asset_condition, "asset_condition", ["semua", "baik", "perlu_perawatan", "rusak_ringan", "rusak_berat", "hilang"])
        self.asset_condition.currentTextChanged.connect(
            lambda: self.reset_page_and_refresh("aset", self.refresh_aset)
        )
        controls.addWidget(self.asset_category)
        controls.addWidget(self.asset_condition)
        layout.addLayout(controls)
        self.asset_table = QTableWidget()
        self.configure_table(
            self.asset_table,
            ["ID", 'Name', 'Category', 'Value', 'Acquired', "Condition", 'Location', 'Next Maintenance', "Status"],
            stretch_column=1,
        )
        self.asset_table.doubleClicked.connect(self.edit_aset)
        layout.addWidget(self.asset_table, 1)
        pager = QHBoxLayout()
        self.create_pager("aset", self.refresh_aset, pager)
        layout.addLayout(pager)

    def refresh_aset(self) -> None:
        try:
            size = self.page_sizes["aset"]
            rows, total = self.manager.page_aset(
                self.asset_search.text().strip(),
                combo_data_or_text(self.asset_category),
                combo_data_or_text(self.asset_condition),
                size,
                self.pages["aset"] * size,
            )
            if not rows and total and self.pages["aset"] > 0:
                self.pages["aset"] -= 1
                return self.refresh_aset()
            self.asset_table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                values = [
                    row["id"],
                    row["nama"],
                    enum_label("asset_category", row["kategori"]),
                    fmt_currency(row["nilai"]),
                    fmt_date(row["tanggal_perolehan"]),
                    enum_label("asset_condition", row["kondisi"]),
                    row["lokasi"],
                    fmt_date(row["tanggal_perawatan_berikut"]),
                    "Active" if row["aktif"] else "Inactive",
                ]
                for column, value in enumerate(values):
                    set_table_item(self.asset_table, row_index, column, value)
            self.update_pager("aset", total)
        except Exception as exc:
            self.handle_error('Failed to load assets', exc)

    def add_aset(self) -> None:
        try:
            dialog = AsetDialog(parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                item_id = self.manager.save_aset(dialog.data())
                QMessageBox.information(self, 'Success', f"Asset {item_id} was added.")
                self.refresh_aset()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to add asset', exc)

    def edit_aset(self) -> None:
        item_id = selected_id(self.asset_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select an asset to edit.')
            return
        try:
            dialog = AsetDialog(self.manager.get_aset(item_id), self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.manager.save_aset(dialog.data(), item_id)
                self.refresh_aset()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to edit asset', exc)

    def delete_aset(self) -> None:
        item_id = selected_id(self.asset_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select an asset to deactivate.')
            return
        if not self.confirm_delete(f"asset {item_id}"):
            return
        try:
            self.manager.delete_aset(item_id)
            self.refresh_aset()
            self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to deactivate asset', exc)

    # ---------------- Transaksi rutin ----------------
    def init_recurring_tab(self) -> None:
        layout = QVBoxLayout(self.recurring_tab)
        info = QLabel(
            "Recurring schedules create actual transactions when Process is used. "
            "The system catches up missed occurrences through the processing date."
        )
        info.setWordWrap(True)
        layout.addWidget(info)
        controls = QHBoxLayout()
        controls.addWidget(self.make_button('Add', self.add_recurring, "add"))
        controls.addWidget(self.make_button('Edit', self.edit_recurring, "edit"))
        controls.addWidget(self.make_button('Delete', self.delete_recurring, "delete"))
        controls.addWidget(self.make_button('Process Due Items', self.process_recurring, "add"))
        controls.addWidget(QLabel('Search:'))
        self.recurring_search = QLineEdit()
        self.recurring_search.setPlaceholderText('ID, name, category, description…')
        self.recurring_search.setClearButtonEnabled(True)
        self.recurring_search.textChanged.connect(
            lambda: self.reset_page_and_refresh("rutin", self.refresh_recurring)
        )
        controls.addWidget(self.recurring_search, 1)
        self.recurring_status = QComboBox()
        add_enum_items(self.recurring_status, "status_filter", ["semua", "aktif", "nonaktif"])
        set_combo_data(self.recurring_status, "aktif")
        self.recurring_status.currentTextChanged.connect(
            lambda: self.reset_page_and_refresh("rutin", self.refresh_recurring)
        )
        controls.addWidget(self.recurring_status)
        layout.addLayout(controls)
        self.recurring_table = QTableWidget()
        self.configure_table(
            self.recurring_table,
            ["ID", 'Name', 'Type', 'Category', 'Amount', 'Donor', 'Interval', 'Next Date', "Status"],
            stretch_column=1,
        )
        self.recurring_table.doubleClicked.connect(self.edit_recurring)
        layout.addWidget(self.recurring_table, 1)
        pager = QHBoxLayout()
        self.create_pager("rutin", self.refresh_recurring, pager)
        layout.addLayout(pager)

    def refresh_recurring(self) -> None:
        try:
            size = self.page_sizes["rutin"]
            rows, total = self.manager.page_recurring(
                self.recurring_search.text().strip(),
                combo_data_or_text(self.recurring_status),
                size,
                self.pages["rutin"] * size,
            )
            if not rows and total and self.pages["rutin"] > 0:
                self.pages["rutin"] -= 1
                return self.refresh_recurring()
            self.recurring_table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                values = [
                    row["id"],
                    row["nama"],
                    enum_label("transaction_type", row["jenis"]),
                    row["kategori"],
                    fmt_currency(row["jumlah"]),
                    row["donatur_nama"],
                    f"{row['interval_nilai']} {enum_label('interval_unit', row['interval_unit'])}",
                    fmt_date(row["tanggal_berikut"]),
                    "Active" if row["aktif"] else "Inactive",
                ]
                for column, value in enumerate(values):
                    set_table_item(self.recurring_table, row_index, column, value)
            self.update_pager("rutin", total)
        except Exception as exc:
            self.handle_error('Failed to load recurring transactions', exc)

    def add_recurring(self) -> None:
        try:
            dialog = RecurringDialog(
                self.manager.list_donatur(),
                self.manager.transaction_categories(),
                parent=self,
            )
            if dialog.exec() == QDialog.DialogCode.Accepted:
                item_id = self.manager.save_recurring(dialog.data())
                QMessageBox.information(self, 'Success', f"Recurring transaction {item_id} was added.")
                self.refresh_recurring()
        except Exception as exc:
            self.handle_error('Failed to add recurring transaction', exc)

    def edit_recurring(self) -> None:
        item_id = selected_id(self.recurring_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a recurring transaction to edit.')
            return
        try:
            dialog = RecurringDialog(
                self.manager.list_donatur(False),
                self.manager.transaction_categories(),
                self.manager.get_recurring(item_id),
                self,
            )
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.manager.save_recurring(dialog.data(), item_id)
                self.refresh_recurring()
        except Exception as exc:
            self.handle_error('Failed to edit recurring transaction', exc)

    def delete_recurring(self) -> None:
        item_id = selected_id(self.recurring_table)
        if not item_id:
            QMessageBox.warning(self, 'Select Data', 'Select a recurring transaction to delete.')
            return
        if not self.confirm_delete(f"recurring transaction {item_id}"):
            return
        try:
            self.manager.delete_recurring(item_id)
            self.refresh_recurring()
        except Exception as exc:
            self.handle_error('Failed to delete recurring transaction', exc)

    def process_recurring(self) -> None:
        try:
            self.manager.require("add")
            date_dialog = BaseDialog(self)
            date_dialog.setWindowTitle('Process Recurring Transactions')
            date_dialog.setFixedSize(430, 180)
            form = QFormLayout(date_dialog)
            date_edit = QDateEdit(QDate.currentDate())
            date_edit.setCalendarPopup(True)
            date_edit.setDisplayFormat(DATE_DISPLAY_FMT)
            form.addRow('Process through date', date_edit)
            warning = QLabel('All active schedules due by this date will be posted as actual transactions.')
            warning.setWordWrap(True)
            form.addRow(warning)
            date_dialog.add_buttons(form)
            if date_dialog.exec() == QDialog.DialogCode.Accepted:
                count = self.manager.post_due_recurring(date_edit.date().toString("yyyy-MM-dd"))
                QMessageBox.information(self, 'End', f"{count} transactions were created.")
                self.refresh_recurring()
                self.refresh_transaksi()
                self.refresh_dashboard()
        except Exception as exc:
            self.handle_error('Failed to process recurring transactions', exc)

    # ---------------- Laporan ----------------
    def init_report_tab(self) -> None:
        layout = QVBoxLayout(self.report_tab)
        controls = QHBoxLayout()
        controls.addWidget(QLabel('Period:'))
        self.report_start = QDateEdit()
        self.report_start.setCalendarPopup(True)
        self.report_start.setDisplayFormat(DATE_DISPLAY_FMT)
        self.report_start.setDate(QDate(dt.date.today().year, 1, 1))
        self.report_end = QDateEdit()
        self.report_end.setCalendarPopup(True)
        self.report_end.setDisplayFormat(DATE_DISPLAY_FMT)
        self.report_end.setDate(QDate.currentDate())
        controls.addWidget(self.report_start)
        controls.addWidget(QLabel('to'))
        controls.addWidget(self.report_end)
        generate = QPushButton('Generate')
        generate.clicked.connect(self.refresh_report)
        export_excel = QPushButton("Export Excel")
        export_excel.clicked.connect(lambda: self.export_report("xlsx"))
        export_pdf = QPushButton("Export PDF")
        export_pdf.clicked.connect(lambda: self.export_report("pdf"))
        controls.addWidget(generate)
        controls.addWidget(export_excel)
        controls.addWidget(export_pdf)
        controls.addStretch()
        layout.addLayout(controls)

        summary_layout = QGridLayout()
        self.report_summary: Dict[str, QLabel] = {}
        for index, (key, title) in enumerate(
            [
                ("pemasukan", 'Total Income'),
                ("pengeluaran", 'Total Expenses'),
                ("saldo", 'Surplus / Deficit'),
                ("count", 'Transaction Count'),
            ]
        ):
            box = QGroupBox(title)
            box_layout = QVBoxLayout(box)
            label = QLabel("—")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
            box_layout.addWidget(label)
            summary_layout.addWidget(box, 0, index)
            self.report_summary[key] = label
        layout.addLayout(summary_layout)

        splitter = QSplitter(Qt.Orientation.Vertical)
        category_box = QGroupBox('Summary by Category')
        category_layout = QVBoxLayout(category_box)
        self.report_category_table = QTableWidget()
        self.configure_table(
            self.report_category_table,
            ['Type', 'Category', 'Transaction Count', "Total"],
            stretch_column=1,
        )
        category_layout.addWidget(self.report_category_table)
        splitter.addWidget(category_box)

        transaction_box = QGroupBox('Transaction Details')
        transaction_layout = QVBoxLayout(transaction_box)
        self.report_transaction_table = QTableWidget()
        self.configure_table(
            self.report_transaction_table,
            ["ID", 'Date', 'Type', 'Category', 'Description', 'Donor', 'Amount'],
            stretch_column=4,
        )
        transaction_layout.addWidget(self.report_transaction_table)
        splitter.addWidget(transaction_box)
        splitter.setSizes([220, 420])
        layout.addWidget(splitter, 1)
        self.current_report: Dict[str, Any] = {}

    def refresh_report(self) -> None:
        try:
            start = self.report_start.date().toString("yyyy-MM-dd")
            end = self.report_end.date().toString("yyyy-MM-dd")
            if start > end:
                raise ValueError('The start date cannot be later than the end date.')
            report = self.manager.financial_report(start, end)
            self.current_report = report
            self.report_summary["pemasukan"].setText(fmt_currency(report["pemasukan"]))
            self.report_summary["pengeluaran"].setText(fmt_currency(report["pengeluaran"]))
            self.report_summary["saldo"].setText(fmt_currency(report["saldo"]))
            self.report_summary["count"].setText(str(len(report["transactions"])))
            categories = report["categories"]
            self.report_category_table.setRowCount(len(categories))
            for row_index, row in enumerate(categories):
                values = [
                    enum_label("transaction_type", row["jenis"]),
                    row["kategori"],
                    row["jumlah_transaksi"],
                    fmt_currency(row["total"]),
                ]
                for column, value in enumerate(values):
                    set_table_item(self.report_category_table, row_index, column, value)
            transactions = report["transactions"]
            self.report_transaction_table.setRowCount(len(transactions))
            for row_index, row in enumerate(transactions):
                values = [
                    row["id"],
                    fmt_datetime(row["tanggal"]),
                    enum_label("transaction_type", row["jenis"]),
                    row["kategori"],
                    elide(row["deskripsi"], 70),
                    row["donatur"],
                    fmt_currency(row["jumlah"]),
                ]
                for column, value in enumerate(values):
                    set_table_item(self.report_transaction_table, row_index, column, value)
            self.statusBar().showMessage('Report updated', 3000)
        except Exception as exc:
            self.handle_error('Failed to generate report', exc)

    def export_report(self, extension: str) -> None:
        try:
            self.manager.require("export")
            self.refresh_report()
            if not self.current_report.get("transactions"):
                raise ValueError('There are no transactions in the report period.')
            start = self.current_report["start_date"]
            end = self.current_report["end_date"]
            export_rows = [
                {
                    "ID": row["id"],
                    "Date": fmt_datetime(row["tanggal"]),
                    "Type": enum_label("transaction_type", row["jenis"]),
                    "Category": row["kategori"],
                    "Description": row["deskripsi"],
                    "Donor": row["donatur"],
                    f"Amount ({_CURRENT_CURRENCY})": fmt_currency(row["jumlah"]),
                }
                for row in self.current_report["transactions"]
            ]
            default = f"financial_report_{start}_{end}.{extension}"
            if extension == "xlsx":
                filename, _ = QFileDialog.getSaveFileName(
                    self, 'Save Excel Report', default, "Excel (*.xlsx)"
                )
                if filename:
                    if not filename.lower().endswith(".xlsx"):
                        filename += ".xlsx"
                    self.manager.export_rows_excel(
                        export_rows, filename, 'Financial Report'
                    )
            else:
                filename, _ = QFileDialog.getSaveFileName(
                    self, 'Save PDF Report', default, "PDF (*.pdf)"
                )
                if filename:
                    if not filename.lower().endswith(".pdf"):
                        filename += ".pdf"
                    title = (
                        f"Financial Report {fmt_date(start)} to {fmt_date(end)} — "
                        f"Income {fmt_currency(self.current_report['pemasukan'])}, "
                        f"Expenses {fmt_currency(self.current_report['pengeluaran'])}, "
                        f"Balance {fmt_currency(self.current_report['saldo'])}"
                    )
                    self.manager.export_rows_pdf(
                        export_rows, filename, title
                    )
            if filename:
                QMessageBox.information(self, 'Success', f"Report saved:\n{filename}")
        except Exception as exc:
            self.handle_error('Failed to export report', exc)

    # ---------------- Log ----------------
    def init_log_tab(self) -> None:
        layout = QVBoxLayout(self.log_tab)
        controls = QHBoxLayout()
        controls.addWidget(QLabel('Search:'))
        self.log_search = QLineEdit()
        self.log_search.setPlaceholderText('User, action, details…')
        self.log_search.setClearButtonEnabled(True)
        self.log_search.returnPressed.connect(self.refresh_logs)
        controls.addWidget(self.log_search, 1)
        refresh = QPushButton('Refresh')
        refresh.clicked.connect(self.refresh_logs)
        controls.addWidget(refresh)
        controls.addWidget(self.make_button('Clean Old Logs', self.cleanup_logs, "delete"))
        layout.addLayout(controls)
        self.log_table = QTableWidget()
        self.configure_table(
            self.log_table,
            ["Waktu", "User", "Aksi", "Detail"],
            stretch_column=3,
        )
        layout.addWidget(self.log_table, 1)

    def refresh_logs(self) -> None:
        try:
            rows = self.manager.logs(self.log_search.text().strip(), 1000)
            self.log_table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                values = [fmt_datetime(row["waktu"]), row["user"], row["aksi"], row["detail"]]
                for column, value in enumerate(values):
                    set_table_item(self.log_table, row_index, column, value)
        except Exception as exc:
            self.handle_error('Failed to load logs', exc)

    def cleanup_logs(self) -> None:
        try:
            dialog = BaseDialog(self)
            dialog.setWindowTitle('Clean Old Logs')
            dialog.setFixedSize(420, 180)
            form = QFormLayout(dialog)
            days = QSpinBox()
            days.setRange(30, 3650)
            days.setValue(365)
            days.setSuffix(' days')
            form.addRow('Delete logs older than', days)
            note = QLabel('Transactions and primary records will not be deleted.')
            note.setWordWrap(True)
            form.addRow(note)
            dialog.add_buttons(form)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                if self.confirm_delete(f"logs older than {days.value()} days"):
                    count = self.manager.cleanup_logs(days.value())
                    QMessageBox.information(self, "Complete", f"{count} log rows were deleted.")
                    self.refresh_logs()
        except Exception as exc:
            self.handle_error('Failed to clean logs', exc)

    # ---------------- File & database operations ----------------
    def choose_table(self, title: str = 'Select Table') -> Optional[str]:
        dialog = BaseDialog(self)
        dialog.setWindowTitle(title)
        dialog.setFixedSize(380, 180)
        form = QFormLayout(dialog)
        combo = QComboBox()
        display = {
            "pengurus": 'Committee',
            "donatur": 'Donor',
            "transaksi": "Transactions",
            "kegiatan": 'Activities',
            "aset": 'Assets',
            "transaksi_rutin": 'Recurring Transactions',
        }
        for key, label in display.items():
            combo.addItem(label, key)
        form.addRow("Table", combo)
        dialog.add_buttons(form)
        return combo.currentData() if dialog.exec() == QDialog.DialogCode.Accepted else None

    def backup_database(self, quiet: bool = False) -> Optional[str]:
        try:
            self.manager.require("backup")
            path = self.db.backup(self.settings.get_backup_dir())
            self.db.log(self.current_user, 'Back Up Database', path)
            self.prune_backups()
            self.settings.set("Database", "last_backup_date", today_iso())
            if not quiet:
                QMessageBox.information(self, 'Backup Successful', f"Database saved to:\n{path}")
            return path
        except Exception as exc:
            if not quiet:
                self.handle_error('Backup failed', exc)
            return None

    def prune_backups(self) -> None:
        directory = Path(self.settings.get_backup_dir())
        if not directory.exists():
            return
        max_files = max(1, self.settings.get_int("Backup", "max_backups", 15))
        backups = sorted(
            [path for path in directory.glob("backup_*.db") if path.is_file()],
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
        for path in backups[max_files:]:
            try:
                path.unlink()
            except OSError:
                pass

    def restore_database(self) -> None:
        try:
            self.manager.require("restore")
            filename, _ = QFileDialog.getOpenFileName(
                self,
                'Select Backup Database',
                self.settings.get_backup_dir(),
                'SQLite Database (*.db);;All Files (*.*)',
            )
            if not filename:
                return
            ok, message = Database.verify_database(filename)
            if not ok:
                raise RuntimeError(f"Invalid database: {message}")
            answer = QMessageBox.warning(
                self,
                'Confirm Restore',
                'Restoring will replace the active database. A safety backup will be created automatically. Continue?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            safety = self.db.restore(filename, self.settings.get_backup_dir())
            self.db.log(self.current_user, 'Restore database', f"{filename}; safety={safety}")
            self.refresh_all()
            QMessageBox.information(
                self,
                'Restore Successful',
                f"Database restored successfully.\nSafety backup: {safety}",
            )
        except Exception as exc:
            self.handle_error('Restore failed', exc)

    def export_data(self) -> None:
        try:
            self.manager.require("export")
            table = self.choose_table("Export Data")
            if not table:
                return
            rows = self.manager.table_rows(table)
            if not rows:
                raise ValueError('The table has no data.')
            filename, selected_filter = QFileDialog.getSaveFileName(
                self,
                'Save Data',
                f"{table}_{dt.datetime.now():%Y%m%d}",
                "Excel (*.xlsx);;CSV (*.csv);;PDF (*.pdf)",
            )
            if not filename:
                return
            if "Excel" in selected_filter:
                if not filename.lower().endswith(".xlsx"):
                    filename += ".xlsx"
                self.manager.export_rows_excel(rows, filename, table)
            elif "CSV" in selected_filter:
                if not filename.lower().endswith(".csv"):
                    filename += ".csv"
                self.manager.export_rows_csv(rows, filename)
            else:
                if not filename.lower().endswith(".pdf"):
                    filename += ".pdf"
                self.manager.export_rows_pdf(rows, filename, f"{table.title()} Data")
            self.db.log(self.current_user, "Export data", f"{table} -> {filename}")
            QMessageBox.information(self, 'Success', f"Data saved:\n{filename}")
        except Exception as exc:
            self.handle_error('Export failed', exc)

    def import_data(self) -> None:
        try:
            self.manager.require("import")
            table = self.choose_table("Import Data")
            if not table:
                return
            filename, _ = QFileDialog.getOpenFileName(
                self,
                'Select Import File',
                "",
                "CSV/Excel (*.csv *.xlsx *.xlsm)",
            )
            if not filename:
                return
            answer = QMessageBox.question(
                self,
                'Confirm Import',
                'Records with matching IDs will be updated. Creating a backup first is recommended. Continue?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            self.backup_database(quiet=True)
            result = self.manager.import_file(table, filename)
            error_text = ""
            if result["errors"]:
                error_text = '\n\nSample errors:\n' + "\n".join(result["errors"][:8])
            QMessageBox.information(
                self,
                'Import Complete',
                f"Inserted: {result['inserted']}\nUpdated: {result['updated']}\n"
                f"Skipped: {result['skipped']}\nFailed: {len(result['errors'])}{error_text}",
            )
            self.refresh_all()
        except Exception as exc:
            self.handle_error('Import failed', exc)

    def check_integrity(self) -> None:
        try:
            ok, message = self.db.integrity_check()
            if ok:
                QMessageBox.information(self, 'Database Integrity', 'The database is healthy (integrity_check: ok).')
            else:
                QMessageBox.critical(self, 'Database Integrity', f"Problem found:\n{message}")
        except Exception as exc:
            self.handle_error('Check failed', exc)

    # ---------------- Settings & users ----------------
    def open_settings(self) -> None:
        try:
            self.manager.require("settings")
            old_db = self.settings.get_db_path()
            dialog = SettingsDialog(self.settings, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                values = dialog.data()
                self.settings.set_many(values)
                configure_localization(self.settings)
                self.dashboard_title.setText(f"🏛️ {values['Masjid']['nama']}")
                self.setWindowTitle(f"{APP_NAME} {APP_VERSION} — {values['Masjid']['nama']}")
                self.apply_theme()
                self.refresh_all()
                if os.path.abspath(values["Database"]["path"]) != old_db:
                    QMessageBox.information(
                        self,
                        'Settings Saved',
                        'The database location changed. Restart the application to use it.',
                    )
                else:
                    QMessageBox.information(self, 'Settings Saved', 'Settings saved successfully.')
        except Exception as exc:
            self.handle_error('Failed to save settings', exc)

    def change_password(self, mandatory: bool = False) -> bool:
        dialog = PasswordDialog(self, mandatory)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return False
        try:
            old_password, new_password = dialog.data()
            self.manager.change_password(self.current_user, old_password, new_password)
            self.auth["wajib_ganti_password"] = 0
            QMessageBox.information(self, 'Success', 'Password updated successfully.')
            return True
        except Exception as exc:
            self.handle_error('Failed to change password', exc)
            if mandatory:
                QTimer.singleShot(100, lambda: self.change_password(True))
            return False

    def manage_users(self) -> None:
        try:
            self.manager.require("users")
            dialog = BaseDialog(self)
            dialog.setWindowTitle('Manage Users')
            dialog.resize(850, 520)
            layout = QVBoxLayout(dialog)
            table = QTableWidget()
            self.configure_table(
                table,
                ["Username", 'Full Name', "Role", "Status", 'Last Login', 'Password Change Required'],
                stretch_column=1,
            )
            layout.addWidget(table, 1)
            buttons = QHBoxLayout()
            add_button = QPushButton('Add')
            edit_button = QPushButton('Edit')
            delete_button = QPushButton('Delete')
            close_button = QPushButton('Close')
            buttons.addWidget(add_button)
            buttons.addWidget(edit_button)
            buttons.addWidget(delete_button)
            buttons.addStretch()
            buttons.addWidget(close_button)
            layout.addLayout(buttons)

            def refresh() -> None:
                users = self.manager.list_users()
                table.setRowCount(len(users))
                for row_index, row in enumerate(users):
                    values = [
                        row["username"],
                        row["nama_lengkap"],
                        row["role"],
                        'Active' if row["aktif"] else "Inactive",
                        fmt_datetime(row["login_terakhir"]),
                        'Yes' if row["wajib_ganti_password"] else 'No',
                    ]
                    for column, value in enumerate(values):
                        set_table_item(table, row_index, column, value)

            def add_user() -> None:
                user_dialog = UserDialog(parent=dialog)
                if user_dialog.exec() == QDialog.DialogCode.Accepted:
                    try:
                        self.manager.save_user(user_dialog.data())
                        refresh()
                    except Exception as exc:
                        self.handle_error('Failed to add user', exc)

            def edit_user() -> None:
                username = selected_id(table)
                if not username:
                    QMessageBox.warning(dialog, 'Select Data', 'Select a user to edit.')
                    return
                row = self.db.fetch_one(
                    "SELECT username,role,nama_lengkap,aktif,wajib_ganti_password FROM users WHERE username=?",
                    (username,),
                )
                user_dialog = UserDialog(dict(row), dialog)
                if user_dialog.exec() == QDialog.DialogCode.Accepted:
                    try:
                        self.manager.save_user(user_dialog.data(), username)
                        refresh()
                    except Exception as exc:
                        self.handle_error('Failed to edit user', exc)

            def delete_user() -> None:
                username = selected_id(table)
                if not username:
                    QMessageBox.warning(dialog, 'Select Data', 'Select a user to delete.')
                    return
                if self.confirm_delete(f"user {username}"):
                    try:
                        self.manager.delete_user(username)
                        refresh()
                    except Exception as exc:
                        self.handle_error('Failed to delete user', exc)

            add_button.clicked.connect(add_user)
            edit_button.clicked.connect(edit_user)
            delete_button.clicked.connect(delete_user)
            table.doubleClicked.connect(edit_user)
            close_button.clicked.connect(dialog.accept)
            refresh()
            dialog.exec()
        except Exception as exc:
            self.handle_error('Failed to open user management', exc)

    # ---------------- Startup, refresh, about, close ----------------
    def run_startup_tasks(self) -> None:
        try:
            self.auto_backup_if_due()
            upcoming = self.manager.upcoming_activities(14, 100)
            maintenance = self.manager.maintenance_due(14, 100)
            due_recurring = int(
                self.db.scalar(
                    "SELECT COUNT(*) FROM transaksi_rutin WHERE aktif=1 AND date(tanggal_berikut)<=date(?)",
                    (today_iso(),),
                )
            )
            messages = []
            if upcoming:
                messages.append(f"{len(upcoming)} activities in the next 14 days")
            if maintenance:
                messages.append(f"{len(maintenance)} assets need maintenance within 14 days")
            if due_recurring:
                messages.append(f"{due_recurring} recurring schedules are due")
            if messages:
                self.notification_status.setToolTip("\n".join(messages))
                self.statusBar().showMessage(" • ".join(messages), 10000)
            if bool(self.auth.get("wajib_ganti_password")):
                QMessageBox.warning(
                    self,
                    'Account Security',
                    'This account is using its default password or is required to change it.',
                )
                self.change_password(True)
        except Exception as exc:
            self.statusBar().showMessage(f"Startup task failed: {exc}", 8000)

    def auto_backup_if_due(self) -> None:
        if not self.settings.get_bool("Database", "auto_backup", True):
            return
        interval = max(1, self.settings.get_int("Database", "backup_interval_days", 1))
        last_text = self.settings.get("Database", "last_backup_date", "")
        last = parse_date(last_text, dt.date(2000, 1, 1))
        if (dt.date.today() - last).days >= interval:
            try:
                path = self.db.backup(self.settings.get_backup_dir())
                self.db.log(self.current_user, 'Automatic backup', path)
                self.prune_backups()
                self.settings.set("Database", "last_backup_date", today_iso())
            except Exception as exc:
                self.statusBar().showMessage(f"Automatic backup failed: {exc}", 8000)

    def refresh_all(self) -> None:
        callbacks = [
            self.refresh_dashboard,
            self.refresh_transaksi,
            self.refresh_donatur,
            self.refresh_pengurus,
            self.refresh_kegiatan,
            self.refresh_aset,
            self.refresh_recurring,
            self.refresh_report,
            self.refresh_logs,
        ]
        for callback in callbacks:
            callback()
        self.statusBar().showMessage('All data refreshed', 3000)

    def show_about(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"About {APP_NAME}")
        dialog.setWindowIcon(self.windowIcon())
        dialog.setModal(True)
        dialog.setMinimumWidth(580)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(28, 24, 28, 22)
        layout.setSpacing(12)

        icon_label = QLabel()
        icon_label.setPixmap(self.windowIcon().pixmap(88, 88))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)

        about_label = QLabel(
            f"""
            <div style="text-align:center; line-height:1.45;">
              <h2 style="margin:4px 0 2px 0;">{APP_NAME}</h2>
              <div><b>Version {APP_VERSION}</b></div>
              <p>
                Created by:<br>
                <b>rahfie27</b>
              </p>
              <p>
                <b>E-COMPUTER</b><br>
                SERVICE KOMPUTER PANGGILAN BOGOR
              </p>
              <p>Copyright © ECOMTECH 2026 - All Right Reserved</p>
              <p>
                <b>Contact:</b><br>
                <a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> /
                <a href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>
              </p>
              <p>
                <b>Donation:</b><br>
                <a href="https://paypal.me/rahfie">paypal.me/rahfie</a>
              </p>
              <p>
                <b>WARNING!</b><br>
                This software is provided as-is without warranty.
              </p>
              <p>
                <b>THANKS TO:</b><br>
                NSANE FORUM ADMIN, STAFF, MOD, MEMBER, AND VISITOR.<br>
                <a href="https://nsaneforums.com">https://nsaneforums.com</a>
              </p>
            </div>
            """
        )
        about_label.setTextFormat(Qt.TextFormat.RichText)
        about_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        about_label.setWordWrap(True)
        about_label.setOpenExternalLinks(True)
        about_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        layout.addWidget(about_label)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.setCenterButtons(True)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)

        dialog.adjustSize()
        dialog.move(self.frameGeometry().center() - dialog.rect().center())
        dialog.exec()

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt naming convention
        if self._closing:
            event.accept()
            return
        self._closing = True
        try:
            # Save normal geometry even when maximized, together with the
            # QMainWindow and toolbar layout state.
            geometry = self.normalGeometry() if self.isMaximized() else self.geometry()
            layout_state = bytes(self.saveState(1).toHex()).decode("ascii")
            self.settings.set_many(
                {
                    "Window": {
                        "width": geometry.width(),
                        "height": geometry.height(),
                        "x": geometry.x(),
                        "y": geometry.y(),
                        "state": "maximized" if self.isMaximized() else "normal",
                        "layout_state": layout_state,
                    }
                }
            )
            if self.settings.get_bool("Database", "auto_backup", True):
                self.auto_backup_if_due()
            self.db.log(self.current_user, 'Exit', 'Application closed')
            self.db.close()
        except Exception:
            pass
        event.accept()


# ============================== ENTRY POINT ==============================


def install_exception_hook() -> None:
    def handler(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        detail = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        try:
            QMessageBox.critical(
                None,
                'Unexpected Error',
                f"An unhandled error occurred:\n\n{exc_value}\n\n"
                "Details were written to error.log.",
            )
        finally:
            with open("error.log", "a", encoding="utf-8") as handle:
                handle.write(f"\n[{now_iso()}]\n{detail}\n")

    sys.excepthook = handler


def main() -> int:
    install_exception_hook()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("E-COMPUTER")
    app.setStyle("Fusion")
    app.setWindowIcon(load_mosque_icon())
    apply_dark_theme(app)

    settings = SettingsManager()
    configure_localization(settings)
    db: Optional[Database] = None
    try:
        db = Database(settings.get_db_path())
    except Exception as exc:
        QMessageBox.critical(None, 'Database Could Not Be Opened', str(exc))
        return 1

    if db.user_count() == 0:
        setup = FirstRunSetupDialog()
        if setup.exec() != QDialog.DialogCode.Accepted:
            db.close()
            return 0
        username, password, remember_username = setup.data()
        try:
            db.create_initial_admin(username, password)
            settings.set_many(
                {
                    "User": {
                        "remember_username": str(remember_username),
                        "last_username": username if remember_username else "",
                    }
                }
            )
            QMessageBox.information(
                None,
                "Setup Complete",
                "The first administrator account was created successfully. "
                "Sign in to continue.",
            )
        except Exception as exc:
            QMessageBox.critical(None, "Setup Failed", str(exc))
            db.close()
            return 1

    while True:
        login = LoginDialog(settings)
        if login.exec() != QDialog.DialogCode.Accepted:
            db.close()
            return 0
        username, password, remember_username = login.credentials()
        if not username or not password:
            QMessageBox.warning(None, "Login", 'Username and password are required.')
            continue
        manager = MasjidManager(db)
        auth = manager.authenticate(username, password)
        if not auth:
            QMessageBox.critical(None, 'Login Failed', 'Incorrect username/password or the account is inactive.')
            continue
        settings.set_many(
            {
                "User": {
                    "remember_username": str(remember_username),
                    "last_username": username if remember_username else "",
                }
            }
        )
        window = MainWindow(settings, db, auth)
        window.show()
        return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
