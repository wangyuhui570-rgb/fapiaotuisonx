import ctypes
import concurrent.futures
import json
import os
import shutil
import sys
import traceback
import urllib.error
import urllib.request

from PySide6.QtCore import QPoint, QThread, Qt, Signal, Slot, QTimer, QSize, QUrl, QObject, QSharedMemory, QEvent
from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont, QGuiApplication, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QStyle,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

import invoice_request_generator as generator
import store_mapping
import template_generator
import wecom_delivery


def app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.getcwd())


def resource_base_dir():
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.abspath(os.getcwd())


APP_DIR = app_base_dir()
RESOURCE_DIR = resource_base_dir()
ASSETS_DIR = os.path.join(RESOURCE_DIR, "assets")
ICONS_DIR = os.path.join(ASSETS_DIR, "icons")
APP_ICON_PATH = os.path.join(ASSETS_DIR, "app_icon.ico")
APP_ICON_PNG_PATH = os.path.join(ASSETS_DIR, "app_icon.png")
GUIDE_PATH = os.path.join(APP_DIR, "APPLICATION_TEMPLATE_GUIDE.txt")
STORE_MAPPING_PATH = os.path.join(APP_DIR, "店铺发票对应公司.txt")
UI_STATE_PATH = os.path.join(APP_DIR, "application_template_state.json")
ERROR_LOG_PATH = os.path.join(APP_DIR, "application_template_error.log")
SINGLE_INSTANCE_KEY = "Mayn.InvoiceRequestTemplateTool.SingleInstance"

DEFAULT_CSV_DIR = APP_DIR
DEFAULT_WECOM_WEBHOOK_URL = ""
DEFAULT_WECOM_WEBHOOK_NOTE = ""
MIN_WINDOW_WIDTH = 820
MIN_WINDOW_HEIGHT = 412
LOG_OPEN_HEIGHT = 780
DEV_WINDOW_SUFFIX = "（开发版）" if not getattr(sys, "frozen", False) else ""


def load_icon(name, fallback=None):
    path = os.path.join(ICONS_DIR, f"{name}.png")
    if os.path.exists(path):
        return QIcon(path)
    return fallback or QIcon()


def load_app_icon():
    if os.path.exists(APP_ICON_PATH):
        return QIcon(APP_ICON_PATH)
    if os.path.exists(APP_ICON_PNG_PATH):
        return QIcon(APP_ICON_PNG_PATH)
    return load_icon("download")


def set_windows_app_id():
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Mayn.InvoiceRequestTemplateTool")
    except Exception:
        return


def apply_soft_shadow(widget, blur=28, y_offset=8, alpha=24):
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setOffset(0, y_offset)
    shadow.setColor(QColor(15, 23, 42, alpha))
    widget.setGraphicsEffect(shadow)


def load_guide_text():
    if os.path.exists(GUIDE_PATH):
        with open(GUIDE_PATH, "r", encoding="utf-8") as file_obj:
            return file_obj.read()
    return (
        "使用说明\n"
        "1. 选择要处理的表格文件或表格文件夹。\n"
        "2. 确认店铺名、所属店铺、货物名称、申请人等参数。\n"
        "3. 点击“开始生成申请模板”，结果会显示在下方，可直接复制。\n"
    )


def create_editable_combo():
    return DoubleClickEditableComboBox()


def create_readonly_line_edit():
    line_edit = QLineEdit()
    line_edit.setReadOnly(True)
    line_edit.setFocusPolicy(Qt.NoFocus)
    return line_edit


class DoubleClickEditableLineEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.setReadOnly(False)
            self.setFocus(Qt.MouseFocusReason)
            self.selectAll()
        super().mouseDoubleClickEvent(event)

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self.setReadOnly(True)


class DoubleClickEditableComboBox(QComboBox):
    def __init__(self):
        super().__init__()
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.double_click_callback = None
        self._popup_timer = QTimer(self)
        self._popup_timer.setSingleShot(True)
        self._popup_timer.timeout.connect(self._show_popup_if_still_readonly)
        if self.lineEdit() is not None:
            self.lineEdit().setReadOnly(True)
            self.lineEdit().installEventFilter(self)

    def _double_click_interval(self):
        app = QApplication.instance()
        if app is None:
            return 250
        return max(180, int(app.doubleClickInterval()))

    def _show_popup_if_still_readonly(self):
        if self.lineEdit() is not None and self.lineEdit().isReadOnly():
            self.showPopup()

    def eventFilter(self, watched, event):
        if watched is self.lineEdit():
            if event.type() == QEvent.MouseButtonDblClick and event.button() == Qt.LeftButton:
                self._popup_timer.stop()
                self.lineEdit().setReadOnly(False)
                self.lineEdit().setFocus(Qt.MouseFocusReason)
                self.lineEdit().selectAll()
                if callable(self.double_click_callback):
                    self.double_click_callback()
                event.accept()
                return True
            if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                if self.lineEdit().isReadOnly():
                    self._popup_timer.start(self._double_click_interval())
                    event.accept()
                    return True
            if event.type() == QEvent.FocusOut:
                self._popup_timer.stop()
                self.lineEdit().setReadOnly(True)
        return super().eventFilter(watched, event)

    def focusOutEvent(self, event):
        self._popup_timer.stop()
        super().focusOutEvent(event)
        if self.lineEdit() is not None:
            self.lineEdit().setReadOnly(True)


class Worker(QObject):
    log = Signal(str)
    finished = Signal(bool, str, object)

    def __init__(self, action_name, task):
        super().__init__()
        self.action_name = action_name
        self.task = task

    @Slot()
    def run(self):
        try:
            result = self.task(self.log.emit)
            self.finished.emit(True, f"{self.action_name}已完成。", result)
        except Exception:
            self.log.emit(traceback.format_exc())
            self.finished.emit(False, f"{self.action_name}失败。", None)


class SingleInstanceGuard:
    def __init__(self, key):
        self.shared_memory = QSharedMemory(key)
        self._owns_lock = False

    def acquire(self):
        if self.shared_memory.create(1):
            self._owns_lock = True
            return True

        if self.shared_memory.error() == QSharedMemory.AlreadyExists:
            return False

        if self.shared_memory.attach():
            self.shared_memory.detach()
            if self.shared_memory.create(1):
                self._owns_lock = True
                return True

        return False

    def release(self):
        if self._owns_lock and self.shared_memory.isAttached():
            self.shared_memory.detach()
        self._owns_lock = False


class TitleBar(QFrame):
    def __init__(self, window):
        super().__init__(window)
        self._drag_active = False
        self._drag_offset = QPoint()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_active = True
            self._drag_offset = event.globalPosition().toPoint() - self.window().frameGeometry().topLeft()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_active and event.buttons() & Qt.LeftButton:
            self.window().move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_active = False
            event.accept()
            return
        super().mouseReleaseEvent(event)


class PathField(QFrame):
    def __init__(self, label_text, line_edit, browse_handler, extra_label=None):
        super().__init__()
        self.setObjectName("PathField")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(6)

        label = QLabel(label_text)
        label.setObjectName("FieldLabel")
        header.addWidget(label, 0, Qt.AlignVCenter)
        header.addStretch(1)

        if extra_label is not None:
            extra_label.setObjectName("InlineFieldHint")
            header.addWidget(extra_label, 0, Qt.AlignVCenter)

        layout.addLayout(header)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)

        line_edit.setObjectName("PathInput")
        row.addWidget(line_edit, 1)

        browse_button = QPushButton("更改")
        browse_button.setProperty("role", "subtle")
        browse_button.setObjectName("CompactButton")
        browse_button.setFixedWidth(68)
        browse_button.clicked.connect(browse_handler)
        row.addWidget(browse_button)

        layout.addLayout(row)


class ActionPathField(QFrame):
    def __init__(self, label_text, line_edit, extra_label=None, buttons=None):
        super().__init__()
        self.setObjectName("PathField")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(6)

        label = QLabel(label_text)
        label.setObjectName("FieldLabel")
        header.addWidget(label, 0, Qt.AlignVCenter)
        header.addStretch(1)

        if extra_label is not None:
            extra_label.setObjectName("InlineFieldHint")
            header.addWidget(extra_label, 0, Qt.AlignVCenter)

        layout.addLayout(header)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)

        line_edit.setObjectName("PathInput")
        row.addWidget(line_edit, 1)

        for button_spec in buttons or []:
            button = QPushButton(button_spec["text"])
            button.setProperty("role", button_spec.get("role", "subtle"))
            button.setObjectName("CompactButton")
            button.setFixedWidth(button_spec.get("width", 68))
            button.clicked.connect(button_spec["handler"])
            row.addWidget(button)

        layout.addLayout(row)


class ConfigField(QFrame):
    def __init__(self, label_text, line_edit):
        super().__init__()
        self.setObjectName("ConfigField")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        label = QLabel(label_text)
        label.setObjectName("FieldLabel")
        label.setFixedWidth(60)
        layout.addWidget(label)

        line_edit.setObjectName("PathInput")
        layout.addWidget(line_edit, 1)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.worker = None
        self.current_action_name = ""
        self.latest_ticker_message = ""
        self.busy_step = 0
        self.log_panel_expanded = False
        self.help_panel_expanded = False
        self.advanced_panel_expanded = False
        self.last_summary = None
        self.current_result_index = 0
        self.selected_csv_paths = []
        self.store_bindings = []
        self._updating_store_widgets = False
        self._store_binding_edit_mode = False
        self._binding_mode_original_store = ""
        self._binding_mode_original_owned = ""
        self.help_window = None
        self.help_text_box = None
        self.log_window = None
        self.log_view_box = None
        self.wecom_window = None
        self.wecom_mode_combo = None
        self.wecom_webhook_edit = None
        self.wecom_webhook_note_edit = None
        self.wecom_smart_bot_id_edit = None
        self.wecom_smart_bot_secret_edit = None
        self.wecom_smart_chat_id_combo = None
        self.wecom_webhook_panel = None
        self.wecom_smart_panel = None
        self.wecom_interval_spin = None
        self.auto_send_timer = QTimer(self)
        self.auto_send_timer.setSingleShot(True)
        self.auto_send_timer.timeout.connect(self._advance_auto_send_wecom)
        self.auto_send_running = False
        self.smart_bot_sender = wecom_delivery.SmartBotSender()
        self.ui_state = self._load_state()

        self.busy_timer = QTimer(self)
        self.busy_timer.setInterval(360)
        self.busy_timer.timeout.connect(self._advance_busy_state)

        self.feedback_timer = QTimer(self)
        self.feedback_timer.setSingleShot(True)
        self.feedback_timer.timeout.connect(self._hide_feedback)

        self.window_icon = load_app_icon()
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setWindowTitle(f"发票申请模版产出工具{DEV_WINDOW_SUFFIX}")
        self.setWindowIcon(self.window_icon)
        self.setMinimumSize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)
        self.resize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)
        self.setAcceptDrops(True)

        self.csv_dir_edit = QLineEdit()
        self.store_name_edit = create_editable_combo()
        self.owned_store_edit = create_readonly_line_edit()
        self.goods_name_edit = DoubleClickEditableLineEdit()
        self.applicant_edit = DoubleClickEditableLineEdit()
        self.result_box = QPlainTextEdit()
        self.result_box.setObjectName("ResultBox")
        self.result_box.setReadOnly(True)
        self.result_box.setFont(QFont("Microsoft YaHei UI", 10))
        self.log_box = QPlainTextEdit()
        self.log_box.setObjectName("LogBox")
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 9))
        self.log_box.hide()

        self._load_settings()
        self._build_ui()
        self._apply_styles()
        self._bind_live_preview()
        self._restore_panel_state()
        self._seed_ticker()
        self._sync_status_labels()
        self.refresh_preview(push_message=False)

    def _load_state(self):
        if not os.path.exists(UI_STATE_PATH):
            return {}
        try:
            with open(UI_STATE_PATH, "r", encoding="utf-8") as file_obj:
                return json.load(file_obj)
        except (OSError, json.JSONDecodeError):
            return {}

    def _load_settings(self):
        self.store_bindings = store_mapping.load_bindings(STORE_MAPPING_PATH)
        self._refresh_store_mapping_widgets()
        csv_dir = self.ui_state.get("csv_dir") or DEFAULT_CSV_DIR
        if not os.path.isdir(csv_dir):
            csv_dir = DEFAULT_CSV_DIR
        self.csv_dir_edit.setText(csv_dir)
        self.store_name_edit.setCurrentText(self.ui_state.get("store_name") or generator.DEFAULT_STORE_NAME)
        self._sync_owned_store_for_selected_store(preferred_owned=self.ui_state.get("owned_store") or generator.DEFAULT_OWNED_STORE)
        self.goods_name_edit.setText(self.ui_state.get("goods_name") or generator.DEFAULT_GOODS_NAME)
        self.applicant_edit.setText(self.ui_state.get("applicant") or generator.DEFAULT_APPLICANT)
        self.wecom_webhook_url = (self.ui_state.get("wecom_webhook_url") or DEFAULT_WECOM_WEBHOOK_URL).strip()
        self.wecom_webhook_note = (self.ui_state.get("wecom_webhook_note") or DEFAULT_WECOM_WEBHOOK_NOTE).strip()
        self.wecom_transport_mode = (self.ui_state.get("wecom_transport_mode") or "webhook").strip() or "webhook"
        self.wecom_smart_bot_id = (self.ui_state.get("wecom_smart_bot_id") or "").strip()
        self.wecom_smart_bot_secret = (self.ui_state.get("wecom_smart_bot_secret") or "").strip()
        self.wecom_smart_chat_id = (self.ui_state.get("wecom_smart_chat_id") or "").strip()
        self.wecom_smart_chat_id_history = list(self.ui_state.get("wecom_smart_chat_id_history") or [])
        self.wecom_send_interval_seconds = float(self.ui_state.get("wecom_send_interval_seconds") or 1.0)
        self.action_mode = self.ui_state.get("action_mode") or "manual"
        self._lock_store_binding_inputs()

    def _save_settings(self):
        self.ui_state = {
            "csv_dir": self.csv_dir(),
            "store_name": self.store_name_edit.currentText().strip(),
            "owned_store": self.owned_store_edit.text().strip(),
            "goods_name": self.goods_name_edit.text().strip(),
            "applicant": self.applicant_edit.text().strip(),
            "wecom_webhook_url": self._wecom_webhook(),
            "wecom_webhook_note": self._wecom_webhook_note(),
            "wecom_transport_mode": getattr(self, "wecom_transport_mode", "webhook"),
            "wecom_smart_bot_id": (getattr(self, "wecom_smart_bot_id", "") or "").strip(),
            "wecom_smart_bot_secret": (getattr(self, "wecom_smart_bot_secret", "") or "").strip(),
            "wecom_smart_chat_id": (getattr(self, "wecom_smart_chat_id", "") or "").strip(),
            "wecom_smart_chat_id_history": list(getattr(self, "wecom_smart_chat_id_history", []) or []),
            "wecom_send_interval_seconds": float(getattr(self, "wecom_send_interval_seconds", 1.0) or 1.0),
            "action_mode": getattr(self, "action_mode", "manual"),
        }
        try:
            with open(UI_STATE_PATH, "w", encoding="utf-8") as file_obj:
                json.dump(self.ui_state, file_obj, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def _refresh_store_mapping_widgets(self):
        current_store = self.store_name_edit.currentText().strip()
        self.store_name_edit.blockSignals(True)
        self.store_name_edit.clear()
        self.store_name_edit.addItems(store_mapping.unique_store_names(self.store_bindings))
        self.store_name_edit.setCurrentText(current_store)
        self.store_name_edit.blockSignals(False)

    def _lock_store_binding_inputs(self):
        self._store_binding_edit_mode = False
        if self.store_name_edit.lineEdit() is not None:
            self.store_name_edit.lineEdit().setReadOnly(True)
        self.owned_store_edit.setReadOnly(True)
        self.owned_store_edit.setFocusPolicy(Qt.NoFocus)

    def _unlock_store_binding_inputs(self):
        self._store_binding_edit_mode = True
        if self.store_name_edit.lineEdit() is not None:
            self.store_name_edit.lineEdit().setReadOnly(False)
            self.store_name_edit.lineEdit().setFocus(Qt.MouseFocusReason)
            self.store_name_edit.lineEdit().selectAll()
        self.owned_store_edit.setReadOnly(False)
        self.owned_store_edit.setFocusPolicy(Qt.StrongFocus)

    def _sync_owned_store_for_selected_store(self, preferred_owned=None):
        store_name = self.store_name_edit.currentText().strip()
        current_owned = preferred_owned if preferred_owned is not None else self.owned_store_edit.text().strip()
        owned_options = store_mapping.owned_stores_for(self.store_bindings, store_name)
        self._updating_store_widgets = True
        self.owned_store_edit.blockSignals(True)
        if owned_options:
            next_owned = current_owned if current_owned in owned_options else owned_options[0]
            self.owned_store_edit.setText(next_owned)
        else:
            self.owned_store_edit.setText(current_owned)
        self.owned_store_edit.blockSignals(False)
        self._updating_store_widgets = False

    def _restore_panel_state(self):
        if hasattr(self, "advanced_panel"):
            self.advanced_panel.show()

    def _seed_ticker(self):
        self.latest_ticker_message = "当前状态：等待开始。"
        self._render_ticker()

    def _build_ui(self):
        root = QWidget()
        root.setObjectName("RootWindow")
        root.setAcceptDrops(True)
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(7, 7, 7, 7)
        outer.setSpacing(6)
        outer.addWidget(self._build_header())
        outer.setAlignment(Qt.AlignTop)
        outer.addWidget(self._build_main_card(), 0, Qt.AlignTop)
        self._install_drop_targets(root)

    def _bind_live_preview(self):
        self.csv_dir_edit.editingFinished.connect(lambda: self.refresh_preview(push_message=False))
        self.store_name_edit.currentTextChanged.connect(self._on_store_name_changed)
        self.store_name_edit.activated.connect(lambda _=0: self._save_settings_and_refresh())
        self.store_name_edit.double_click_callback = self.enter_store_binding_mode
        if self.store_name_edit.lineEdit() is not None:
            self.store_name_edit.lineEdit().editingFinished.connect(self._on_store_name_editing_finished)
        self.owned_store_edit.editingFinished.connect(self._on_owned_store_editing_finished)
        for widget in (self.goods_name_edit, self.applicant_edit):
            widget.editingFinished.connect(lambda w=widget: w.setReadOnly(True))
            widget.editingFinished.connect(lambda: self._save_settings_and_refresh())

    def _install_drop_targets(self, widget):
        widget.installEventFilter(self)
        for child in widget.findChildren(QWidget):
            child.installEventFilter(self)

    def _build_header(self):
        title_bar = TitleBar(self)
        title_bar.setObjectName("TitleBar")

        row = QHBoxLayout(title_bar)
        row.setContentsMargins(12, 0, 10, 0)
        row.setSpacing(8)

        title_icon = QLabel()
        title_icon.setObjectName("TitleIcon")
        title_icon.setPixmap(self.window_icon.pixmap(18, 18))
        row.addWidget(title_icon)

        title = QLabel(f"发票申请模版产出工具{DEV_WINDOW_SUFFIX}")
        title.setObjectName("WindowTitle")
        row.addWidget(title, 1)

        self.more_button = QToolButton()
        self.more_button.setObjectName("MoreButton")
        self.more_button.setText("")
        self.more_button.setIcon(load_icon("more", self.style().standardIcon(QStyle.SP_TitleBarUnshadeButton)))
        self.more_button.setIconSize(QSize(14, 14))
        self.more_button.setPopupMode(QToolButton.InstantPopup)
        self.more_button.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.more_button.setMenu(self._build_more_menu())
        row.addWidget(self.more_button)

        self.minimize_button = QToolButton()
        self.minimize_button.setObjectName("TitleBarButton")
        self.minimize_button.setText("-")
        self.minimize_button.clicked.connect(self.showMinimized)
        row.addWidget(self.minimize_button)

        self.close_button = QToolButton()
        self.close_button.setObjectName("CloseTitleBarButton")
        self.close_button.setText("x")
        self.close_button.clicked.connect(self.close)
        row.addWidget(self.close_button)
        return title_bar

    def _build_main_card(self):
        card = QFrame()
        self.main_card = card
        card.setObjectName("MainCard")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        apply_soft_shadow(card, blur=32, y_offset=10, alpha=22)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(6)

        self.feedback_banner = QLabel(card)
        self.feedback_banner.setObjectName("FeedbackBanner")
        self.feedback_banner.setWordWrap(True)
        self.feedback_banner.hide()
        self.feedback_banner.raise_()

        content = QWidget()
        content.setObjectName("MainContent")

        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(6)
        content_layout.setAlignment(Qt.AlignTop)

        self.csv_preview_label = QLabel()
        self.csv_preview_label.setWordWrap(False)

        self.csv_path_field = ActionPathField(
            "CSV 文件夹",
            self.csv_dir_edit,
            self.csv_preview_label,
            buttons=[
                {"text": "更改", "handler": self.choose_csv_dir, "width": 68},
                {"text": "清空CSV", "handler": self.clear_current_csv_files, "width": 82},
            ],
        )
        content_layout.addWidget(self.csv_path_field)

        self.generate_btn = QPushButton("开始生成")
        self.generate_btn.hide()

        self.advanced_panel = QFrame()
        self.advanced_panel.setObjectName("AdvancedPanel")
        self.advanced_panel.setFixedWidth(248)
        advanced_layout = QVBoxLayout(self.advanced_panel)
        advanced_layout.setContentsMargins(8, 8, 8, 8)
        advanced_layout.setSpacing(6)

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(0)
        grid.setVerticalSpacing(5)
        grid.addWidget(ConfigField("店铺名", self.store_name_edit), 0, 0)
        grid.addWidget(ConfigField("所属店铺", self.owned_store_edit), 1, 0)
        grid.addWidget(ConfigField("货物名称", self.goods_name_edit), 2, 0)
        grid.addWidget(ConfigField("申请人", self.applicant_edit), 3, 0)
        advanced_layout.addLayout(grid)

        result_tools = QHBoxLayout()
        result_tools.setContentsMargins(0, 0, 0, 0)
        result_tools.setSpacing(6)

        self.select_csv_files_btn = QPushButton("选择文件")
        self.select_csv_files_btn.setProperty("role", "subtle")
        self.select_csv_files_btn.setObjectName("CompactButton")
        self.select_csv_files_btn.setFixedWidth(86)
        self.select_csv_files_btn.clicked.connect(self.choose_csv_files)
        result_tools.addWidget(self.select_csv_files_btn)

        self.use_all_csv_btn = QPushButton("全部文件")
        self.use_all_csv_btn.setProperty("role", "subtle")
        self.use_all_csv_btn.setObjectName("CompactButton")
        self.use_all_csv_btn.setFixedWidth(86)
        self.use_all_csv_btn.clicked.connect(self.reset_csv_selection)
        result_tools.addWidget(self.use_all_csv_btn)

        self.csv_selection_label = QLabel("全部CSV（0）")
        self.csv_selection_label.setObjectName("HintLabel")
        self.csv_selection_label.setWordWrap(False)
        result_tools.addWidget(self.csv_selection_label)

        self.action_mode_combo = QComboBox()
        self.action_mode_combo.setObjectName("PathInput")
        self.action_mode_combo.setFixedWidth(104)
        self.action_mode_combo.addItem("手动模式", "manual")
        self.action_mode_combo.addItem("机器人模式", "robot")
        current_index = max(0, self.action_mode_combo.findData(self.action_mode))
        self.action_mode_combo.setCurrentIndex(current_index)
        self.action_mode_combo.currentIndexChanged.connect(self.on_action_mode_changed)
        result_tools.addWidget(self.action_mode_combo)

        result_tools.addStretch(1)

        self.prev_result_btn = QPushButton("上一条")
        self.prev_result_btn.setProperty("role", "subtle")
        self.prev_result_btn.setObjectName("CompactButton")
        self.prev_result_btn.setFixedWidth(76)
        self.prev_result_btn.clicked.connect(self.show_prev_result)
        result_tools.addWidget(self.prev_result_btn)

        self.next_result_btn = QPushButton("下一条")
        self.next_result_btn.setProperty("role", "subtle")
        self.next_result_btn.setObjectName("CompactButton")
        self.next_result_btn.setFixedWidth(76)
        self.next_result_btn.clicked.connect(self.show_next_result)
        result_tools.addWidget(self.next_result_btn)

        self.copy_result_btn = QPushButton("复制当前")
        self.copy_result_btn.setProperty("role", "subtle")
        self.copy_result_btn.setObjectName("CompactButton")
        self.copy_result_btn.setFixedWidth(92)
        self.copy_result_btn.clicked.connect(self.copy_result_text)
        result_tools.addWidget(self.copy_result_btn)

        self.copy_next_result_btn = QPushButton("复制并下一条")
        self.copy_next_result_btn.setProperty("role", "subtle")
        self.copy_next_result_btn.setObjectName("CompactButton")
        self.copy_next_result_btn.setFixedWidth(116)
        self.copy_next_result_btn.clicked.connect(self.copy_and_next_result)
        result_tools.addWidget(self.copy_next_result_btn)

        self.send_wecom_btn = QPushButton("发群")
        self.send_wecom_btn.setProperty("role", "subtle")
        self.send_wecom_btn.setObjectName("CompactButton")
        self.send_wecom_btn.setFixedWidth(70)
        self.send_wecom_btn.clicked.connect(self.send_current_to_wecom)
        result_tools.addWidget(self.send_wecom_btn)

        self.send_next_wecom_btn = QPushButton("发群并下一条")
        self.send_next_wecom_btn.setProperty("role", "subtle")
        self.send_next_wecom_btn.setObjectName("CompactButton")
        self.send_next_wecom_btn.setFixedWidth(116)
        self.send_next_wecom_btn.clicked.connect(self.send_current_to_wecom_and_next)
        result_tools.addWidget(self.send_next_wecom_btn)

        self.auto_send_wecom_btn = QPushButton("自动发群")
        self.auto_send_wecom_btn.setProperty("role", "subtle")
        self.auto_send_wecom_btn.setObjectName("CompactButton")
        self.auto_send_wecom_btn.setFixedWidth(92)
        self.auto_send_wecom_btn.clicked.connect(self.toggle_auto_send_wecom)
        result_tools.addWidget(self.auto_send_wecom_btn)
        content_layout.addLayout(result_tools)

        result_body = QHBoxLayout()
        result_body.setContentsMargins(0, 0, 0, 0)
        result_body.setSpacing(8)

        self.result_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.result_box.setMinimumHeight(220)
        self.result_box.setMaximumHeight(280)
        self.result_index_label = QLabel("0 / 0", self.result_box.viewport())
        self.result_index_label.setObjectName("ResultIndexLabel")
        self.result_index_label.setWordWrap(False)
        self.result_index_label.hide()
        self.result_index_label.raise_()

        result_body.addWidget(self.result_box, 1)
        result_body.addWidget(self.advanced_panel)
        content_layout.addLayout(result_body)

        self.ticker_bar = QLabel()
        self.ticker_bar.hide()

        self.inline_status_label = QLabel()
        self.inline_status_label.hide()

        self.inline_path_label = QLabel()
        self.inline_path_label.hide()

        layout.addWidget(content)
        layout.setAlignment(content, Qt.AlignTop)
        self._apply_action_mode()
        self._refresh_robot_button_labels()
        self._update_auto_send_button()
        return card

    def _build_help_panel(self):
        panel = QFrame()
        panel.setObjectName("HelpPanel")

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("使用帮助")
        title.setObjectName("SectionTitle")
        header.addWidget(title)
        header.addStretch(1)

        hide_btn = QPushButton("收起")
        hide_btn.setProperty("role", "subtle")
        hide_btn.setObjectName("CompactButton")
        hide_btn.setFixedWidth(82)
        hide_btn.clicked.connect(lambda: self.toggle_help_panel(force_visible=False))
        header.addWidget(hide_btn)
        layout.addLayout(header)

        text = QLabel(load_guide_text())
        text.setObjectName("HelpText")
        text.setTextFormat(Qt.PlainText)
        text.setWordWrap(True)
        text.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(text)
        return panel

    def _build_more_menu(self):
        menu = QMenu(self)

        help_action = QAction("使用帮助", self)
        help_action.setIcon(load_icon("help", self.style().standardIcon(QStyle.SP_DialogHelpButton)))
        help_action.triggered.connect(self.show_help_window)
        menu.addAction(help_action)

        log_action = QAction("查看日志", self)
        log_action.setIcon(load_icon("log", self.style().standardIcon(QStyle.SP_FileDialogDetailedView)))
        log_action.triggered.connect(self.show_log_window)
        menu.addAction(log_action)

        reset_action = QAction("恢复默认参数", self)
        reset_action.setIcon(load_icon("refresh", self.style().standardIcon(QStyle.SP_BrowserReload)))
        reset_action.triggered.connect(self.restore_default_settings)
        menu.addAction(reset_action)

        wecom_setting_action = QAction("企业微信机器人设置", self)
        wecom_setting_action.setIcon(load_icon("send", self.style().standardIcon(QStyle.SP_ArrowForward)))
        wecom_setting_action.triggered.connect(self.show_wecom_window)
        menu.addAction(wecom_setting_action)

        menu.addSeparator()

        open_csv_action = QAction("打开 CSV 文件夹", self)
        open_csv_action.setIcon(load_icon("folder", self.style().standardIcon(QStyle.SP_DirOpenIcon)))
        open_csv_action.triggered.connect(lambda: self.open_folder(self.csv_dir()))
        menu.addAction(open_csv_action)

        menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.setIcon(load_icon("exit", self.style().standardIcon(QStyle.SP_DialogCloseButton)))
        exit_action.triggered.connect(self.close)
        menu.addAction(exit_action)
        return menu

    def _apply_styles(self):
        self.setStyleSheet(
            """
            QWidget {
                background: #eef1f5;
                color: #1d1d1f;
                font-family: "Segoe UI Variable", "Segoe UI", "Microsoft YaHei UI", sans-serif;
                font-size: 13px;
            }
            QMainWindow { background: #e9edf3; }
            #RootWindow {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f6f8fb, stop:1 #edf1f5);
                border: none;
                border-radius: 16px;
            }
            #TitleBar {
                min-height: 40px;
                background: rgba(255,255,255,0.54);
                border: none;
                border-radius: 12px;
            }
            #TitleIcon { background: transparent; min-width: 18px; max-width: 18px; }
            #WindowTitle {
                font-size: 15px;
                font-weight: 600;
                color: #1d1d1f;
                background: transparent;
            }
            #MainCard {
                background: rgba(255,255,255,0.90);
                border: none;
                border-radius: 16px;
            }
            #PathField, #ConfigField, #AdvancedPanel {
                background: rgba(255,255,255,0.72);
                border: none;
                border-radius: 12px;
            }
            #FieldLabel, #SectionTitle {
                background: transparent;
                color: #4f5663;
                font-size: 11px;
                font-weight: 600;
                letter-spacing: 0.2px;
            }
            #InlineFieldHint {
                background: transparent;
                color: #64748b;
                font-size: 11px;
            }
            #ResultIndexLabel {
                background: transparent;
                color: #64748b;
                font-size: 11px;
            }
            #HintLabel, #HelpText {
                background: transparent;
                color: #6b7280;
                font-size: 12px;
                line-height: 1.5;
            }
            #PathInput, QComboBox#PathInput, #ResultBox, #LogBox {
                background: rgba(255,255,255,0.96);
                border: 1px solid rgba(148,163,184,0.34);
                border-radius: 10px;
                padding: 8px 10px;
                selection-background-color: rgba(37,99,235,0.18);
            }
            #PathInput:focus, QComboBox#PathInput:focus, #ResultBox:focus, #LogBox:focus {
                border: 1px solid rgba(37,99,235,0.65);
            }
            QComboBox#PathInput {
                padding-right: 8px;
            }
            QComboBox#PathInput::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 0px;
                border: none;
                background: transparent;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }
            QComboBox#PathInput::drop-down:hover {
                background: transparent;
            }
            QComboBox#PathInput::down-arrow {
                width: 0px;
                height: 0px;
            }
            #ResultBox, #LogBox {
                padding: 10px 12px;
                background: rgba(250,251,253,0.98);
            }
            QPushButton {
                min-height: 34px;
                padding: 0 14px;
                border-radius: 10px;
                border: 1px solid rgba(148,163,184,0.28);
                background: rgba(255,255,255,0.82);
                color: #334155;
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: rgba(255,255,255,0.98);
                border: 1px solid rgba(148,163,184,0.38);
            }
            QPushButton:pressed { background: rgba(241,245,249,0.98); }
            QPushButton[role="primary"] {
                background: #2563eb;
                border: 1px solid #2563eb;
                color: white;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton[role="primary"]:hover {
                background: #1d4ed8;
                border: 1px solid #1d4ed8;
            }
            QPushButton[role="primary"][busy="true"] {
                background: #4f7df0;
                border: 1px solid #4f7df0;
            }
            #CompactButton {
                min-height: 32px;
                padding: 0 12px;
                border-radius: 9px;
            }
            #AdvancedToggleButton {
                min-height: 32px;
                padding: 0 12px;
                border-radius: 10px;
                text-align: left;
            }
            #MoreButton, #TitleBarButton, #CloseTitleBarButton {
                min-width: 28px;
                max-width: 28px;
                min-height: 28px;
                max-height: 28px;
                padding: 0;
                border-radius: 14px;
                background: rgba(255,255,255,0.72);
                border: 1px solid rgba(148,163,184,0.22);
                color: #475569;
                font-size: 14px;
                font-weight: 600;
            }
            #MoreButton::menu-indicator { image: none; width: 0px; }
            #MoreButton:hover, #TitleBarButton:hover, #CloseTitleBarButton:hover {
                background: rgba(255,255,255,0.94);
            }
            #CloseTitleBarButton:hover {
                background: rgba(239,68,68,0.12);
                color: #dc2626;
                border: 1px solid rgba(239,68,68,0.18);
            }
            #TickerBar {
                background: rgba(248,250,252,0.94);
                border: 1px solid rgba(148,163,184,0.22);
                border-radius: 11px;
                padding: 10px 12px;
                color: #334155;
                font-size: 12px;
            }
            #InlineStatusMeta, #InlinePathMeta {
                background: transparent;
                color: #64748b;
                font-size: 12px;
            }
            #FeedbackBanner {
                background: rgba(37,99,235,0.10);
                border: 1px solid rgba(37,99,235,0.18);
                border-radius: 11px;
                color: #1d4ed8;
                padding: 10px 12px;
                font-size: 12px;
                font-weight: 600;
            }
            #FeedbackBanner[tone="success"] {
                background: rgba(34,197,94,0.10);
                border: 1px solid rgba(34,197,94,0.18);
                color: #15803d;
            }
            #FeedbackBanner[tone="error"] {
                background: rgba(239,68,68,0.10);
                border: 1px solid rgba(239,68,68,0.18);
                color: #dc2626;
            }
            QMenu {
                background: rgba(255,255,255,0.98);
                border: 1px solid rgba(148,163,184,0.20);
                border-radius: 10px;
                padding: 6px;
            }
            QMenu::item {
                padding: 8px 18px 8px 12px;
                border-radius: 8px;
                color: #334155;
            }
            QMenu::item:selected {
                background: rgba(37,99,235,0.08);
                color: #1d4ed8;
            }
            """
        )

    def _set_primary_busy(self, busy):
        if not hasattr(self, "generate_btn"):
            return
        self.generate_btn.setProperty("busy", "true" if busy else "false")
        self.generate_btn.style().unpolish(self.generate_btn)
        self.generate_btn.style().polish(self.generate_btn)
        self.generate_btn.update()

    def _push_ticker_message(self, message):
        text = " ".join(part.strip() for part in str(message).splitlines() if part.strip())
        if not text:
            return
        self.latest_ticker_message = text
        self._render_ticker()

    def _render_ticker(self):
        if hasattr(self, "ticker_bar"):
            self.ticker_bar.setText(self.latest_ticker_message or "当前状态：等待开始。")

    def _position_feedback_banner(self):
        margin_x = 14
        top = 10
        width = max(220, self.main_card.width() - margin_x * 2)
        self.feedback_banner.setFixedWidth(width)
        self.feedback_banner.adjustSize()
        height = max(40, self.feedback_banner.sizeHint().height() + 12)
        self.feedback_banner.setGeometry(margin_x, top, width, height)

    def _position_result_index_label(self):
        if not hasattr(self, "result_index_label") or not hasattr(self, "result_box"):
            return
        self.result_index_label.adjustSize()
        label_size = self.result_index_label.sizeHint()
        viewport = self.result_box.viewport()
        x = max(6, viewport.width() - label_size.width() - 12)
        y = max(6, viewport.height() - label_size.height() - 8)
        self.result_index_label.move(x, y)

    def _show_feedback(self, message, tone, timeout_ms=5000):
        self.feedback_banner.setText(message)
        self.feedback_banner.setProperty("tone", tone)
        self.feedback_banner.style().unpolish(self.feedback_banner)
        self.feedback_banner.style().polish(self.feedback_banner)
        self._position_feedback_banner()
        self.feedback_banner.show()
        self.feedback_banner.raise_()
        if timeout_ms > 0:
            self.feedback_timer.start(timeout_ms)
        else:
            self.feedback_timer.stop()

    def _hide_feedback(self):
        self.feedback_timer.stop()
        self.feedback_banner.hide()

    def csv_dir(self):
        return os.path.abspath(self.csv_dir_edit.text().strip() or DEFAULT_CSV_DIR)

    def selected_csv_paths_in_dir(self):
        normalized = []
        for path in self.selected_csv_paths:
            absolute = os.path.abspath(path)
            if not template_generator.is_supported_tabular_file(absolute):
                continue
            normalized.append(absolute)
        self.selected_csv_paths = normalized
        return normalized

    def current_csv_paths(self):
        selected = self.selected_csv_paths_in_dir()
        if selected:
            return selected
        return template_generator.list_csv_files(self.csv_dir())

    def settings_payload(self):
        return {
            "store_name": self.store_name_edit.currentText().strip() or generator.DEFAULT_STORE_NAME,
            "owned_store": self.owned_store_edit.text().strip() or generator.DEFAULT_OWNED_STORE,
            "goods_name": self.goods_name_edit.text().strip() or generator.DEFAULT_GOODS_NAME,
            "applicant": self.applicant_edit.text().strip() or generator.DEFAULT_APPLICANT,
        }

    def _sync_status_labels(self, state="等待开始", result="请先确认表格文件夹和申请参数。"):
        if hasattr(self, "inline_status_label"):
            self.inline_status_label.setText(f"当前状态：{state} | {result}")
        if hasattr(self, "inline_path_label"):
            self.inline_path_label.setText("结果位置：下方结果框")

    def _save_settings_and_refresh(self):
        if self._updating_store_widgets:
            return
        self._save_settings()
        self.refresh_preview(push_message=False)

    def _apply_action_mode(self):
        manual_mode = getattr(self, "action_mode", "manual") != "robot"
        total = len((self.last_summary or {}).get("texts") or [])
        if hasattr(self, "copy_result_btn"):
            self.copy_result_btn.setVisible(manual_mode)
        if hasattr(self, "copy_next_result_btn"):
            self.copy_next_result_btn.setVisible(manual_mode)
        if hasattr(self, "send_wecom_btn"):
            self.send_wecom_btn.setVisible(not manual_mode)
        if hasattr(self, "send_next_wecom_btn"):
            self.send_next_wecom_btn.setVisible(not manual_mode)
        if hasattr(self, "auto_send_wecom_btn"):
            self.auto_send_wecom_btn.setVisible(not manual_mode)
        if hasattr(self, "prev_result_btn"):
            self.prev_result_btn.setVisible(manual_mode and total > 1)
        if hasattr(self, "next_result_btn"):
            self.next_result_btn.setVisible(manual_mode and total > 1)

    def _refresh_robot_button_labels(self):
        if self._wecom_transport() == "smart_bot":
            single_text = "发送"
            next_text = "发送并下一条"
            auto_text = "自动发送"
            pause_text = "暂停发送"
        else:
            single_text = "发群"
            next_text = "发群并下一条"
            auto_text = "自动发群"
            pause_text = "暂停发群"
        if hasattr(self, "send_wecom_btn"):
            self.send_wecom_btn.setText(single_text)
        if hasattr(self, "send_next_wecom_btn"):
            self.send_next_wecom_btn.setText(next_text)
        if hasattr(self, "auto_send_wecom_btn") and not self.auto_send_running:
            self.auto_send_wecom_btn.setText(auto_text)
        self._auto_send_pause_text = pause_text
        self._auto_send_idle_text = auto_text

    def on_action_mode_changed(self, _index):
        if not hasattr(self, "action_mode_combo"):
            return
        self.action_mode = self.action_mode_combo.currentData() or "manual"
        self._apply_action_mode()
        self._save_settings()

    def _on_store_name_changed(self, _text):
        if self._updating_store_widgets:
            return
        if not self._store_binding_edit_mode:
            self._sync_owned_store_for_selected_store()
            return
        current_store = self.store_name_edit.currentText().strip()
        if current_store == self._binding_mode_original_store:
            self.owned_store_edit.setText(self._binding_mode_original_owned)
            return
        owned_options = store_mapping.owned_stores_for(self.store_bindings, current_store)
        if owned_options:
            self.owned_store_edit.setText(owned_options[0])
            return
        self.owned_store_edit.clear()

    def enter_store_binding_mode(self):
        self._binding_mode_original_store = self.store_name_edit.currentText().strip()
        self._binding_mode_original_owned = self.owned_store_edit.text().strip()
        self._unlock_store_binding_inputs()
        self._show_feedback("请先双击店铺名，再输入新的店铺名和所属店铺完成绑定。", "info", timeout_ms=3500)

    def _on_store_name_editing_finished(self):
        if self.store_name_edit.lineEdit() is not None and not self._store_binding_edit_mode:
            self.store_name_edit.lineEdit().setReadOnly(True)
        if not self._store_binding_edit_mode:
            self._save_settings_and_refresh()

    def _on_owned_store_editing_finished(self):
        if self._store_binding_edit_mode:
            store_name = self.store_name_edit.currentText().strip()
            owned_store = self.owned_store_edit.text().strip()
            if store_name and owned_store:
                updated_bindings = store_mapping.upsert_binding(self.store_bindings, store_name, owned_store)
                if updated_bindings != self.store_bindings:
                    store_mapping.save_bindings(STORE_MAPPING_PATH, updated_bindings)
                    self.store_bindings = updated_bindings
                    self._refresh_store_mapping_widgets()
                self.store_name_edit.setCurrentText(store_name)
                self._sync_owned_store_for_selected_store(preferred_owned=owned_store)
                self._lock_store_binding_inputs()
                self._save_settings_and_refresh()
                self._show_feedback("新绑定已保存。", "success", timeout_ms=2500)
                return
            self._lock_store_binding_inputs()
            self._show_feedback("保存店铺绑定失败，请检查店铺名和所属店铺。", "error", timeout_ms=3000)
            return
        self.owned_store_edit.setReadOnly(True)

    def choose_csv_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择表格文件夹", self.csv_dir())
        if path:
            self.csv_dir_edit.setText(path)
            self.selected_csv_paths = []
            self._save_settings()
            self.refresh_preview()

    def choose_csv_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择要处理的表格文件",
            self.csv_dir(),
            "表格文件 (*.csv *.xlsx *.xlsm)",
        )
        if file_paths:
            self.selected_csv_paths = [os.path.abspath(path) for path in file_paths]
            self.refresh_preview()

    def reset_csv_selection(self):
        self.selected_csv_paths = []
        self.refresh_preview()

    def _dropped_csv_paths(self, mime_data):
        if mime_data is None or not mime_data.hasUrls():
            return []
        csv_paths = []
        for url in mime_data.urls():
            if not url.isLocalFile():
                continue
            local_path = os.path.abspath(url.toLocalFile())
            if template_generator.is_supported_tabular_file(local_path):
                csv_paths.append(local_path)
        return csv_paths

    def clear_current_csv_files(self):
        csv_dir = self.csv_dir()
        if not os.path.isdir(csv_dir):
            self._show_feedback("表格文件夹不存在，请先重新选择。", "error", timeout_ms=3500)
            return

        removed_count = 0
        for name in os.listdir(csv_dir):
            path = os.path.join(csv_dir, name)
            if template_generator.is_supported_tabular_file(path):
                os.remove(path)
                removed_count += 1

        self.selected_csv_paths = []
        self._save_settings()
        self.refresh_preview(push_message=False)

        if removed_count:
            self._show_feedback(f"已清空当前表格文件夹，删除 {removed_count} 个表格文件。", "success", timeout_ms=3000)
            self._push_ticker_message(f"已清空 {removed_count} 个表格文件。")
        else:
            self._show_feedback("当前文件夹里没有可清空的表格文件。", "info", timeout_ms=2500)

    def import_dropped_csv_files(self, csv_paths):
        if not csv_paths:
            self._show_feedback("没有可导入的表格文件。", "info", timeout_ms=2500)
            return

        target_dir = self.csv_dir()
        if not os.path.isdir(target_dir):
            self._show_feedback("表格文件夹不存在，请先重新选择。", "error", timeout_ms=3500)
            return

        copied_count = 0
        skipped_count = 0
        imported_paths = []
        for source_path in csv_paths:
            file_name = os.path.basename(source_path)
            stem, ext = os.path.splitext(file_name)
            candidate = os.path.join(target_dir, file_name)
            suffix = 1
            while os.path.exists(candidate):
                candidate = os.path.join(target_dir, f"{stem}_{suffix}{ext}")
                suffix += 1
            try:
                shutil.copy2(source_path, candidate)
                imported_paths.append(candidate)
                copied_count += 1
            except OSError:
                skipped_count += 1

        if copied_count:
            self.selected_csv_paths = []
            self._save_settings()
            self.refresh_preview(push_message=False)
            message = f"已导入 {copied_count} 个表格文件到当前文件夹。"
            if skipped_count:
                message += f" 跳过 {skipped_count} 个。"
            self._show_feedback(message, "success", timeout_ms=3500)
            self._push_ticker_message(message)
            return

        self._show_feedback("没有导入任何表格文件。", "error", timeout_ms=3500)

    def refresh_preview(self, push_message=True):
        try:
            csv_paths = self.current_csv_paths()
            _, rows, _ = template_generator.read_csv_rows(csv_paths[0])
            selected_count = len(self.selected_csv_paths_in_dir())
            if selected_count:
                self.csv_selection_label.setText(f"已选 {selected_count} 个")
            else:
                self.csv_selection_label.setText(f"全部表格（{len(csv_paths)}）")
            self.use_all_csv_btn.setDisabled(selected_count == 0)
            self.csv_preview_label.setText(f"{len(csv_paths)}个表格文件，可生成{len(rows)}份。")
            if rows:
                self.current_result_index = 0
                self.last_summary = {"texts": [generator.build_output_text(rows[0], self.settings_payload()).rstrip()]}
                self._render_current_result()
            else:
                self.result_box.setPlainText("请选择表格文件夹。")
                self.result_index_label.setText("0 / 0")
            if push_message:
                self._push_ticker_message(f"已选择 {len(csv_paths)} 个表格文件，等待生成。")
            self._sync_status_labels(result=f"已选择 {len(csv_paths)} 个表格文件。")
        except Exception as exc:
            self.csv_selection_label.setText("读取表格失败")
            self.use_all_csv_btn.setDisabled(True)
            self.csv_preview_label.setText(str(exc))
            self.result_box.setPlainText("结果会在读取 CSV 后显示在这里。")
            self.result_index_label.setText("0 / 0")
            self.last_summary = None
            if push_message:
                self._push_ticker_message("当前目录里没有找到表格文件。")
            self._sync_status_labels(result="当前目录里没有找到表格文件。")

    def _copy_current_result_to_clipboard(self, feedback_message):
        text = self.result_box.toPlainText().strip()
        if not text:
            self._show_feedback("当前还没有可复制的模板。", "info", timeout_ms=2500)
            return False
        QGuiApplication.clipboard().setText(text)
        self._show_feedback(feedback_message, "success", timeout_ms=2500)
        self._push_ticker_message(feedback_message)
        return True

    def copy_result_text(self):
        self._copy_current_result_to_clipboard("结果已复制到剪贴板。")

    def copy_and_next_result(self):
        texts = (self.last_summary or {}).get("texts") or []
        if not self._copy_current_result_to_clipboard("结果已复制并切到下一条。"):
            return
        if not texts:
            return
        if self.current_result_index < len(texts) - 1:
            self.current_result_index += 1
            self._render_current_result()
            return
        self._show_feedback("结果已复制，当前已经是最后一条。", "success", timeout_ms=2500)
        self._push_ticker_message("结果已复制，当前已经是最后一条。")

    def current_result_text(self):
        return self.result_box.toPlainText().strip()

    def _wecom_transport(self):
        return (getattr(self, "wecom_transport_mode", "webhook") or "webhook").strip() or "webhook"

    def _wecom_webhook(self):
        return (getattr(self, "wecom_webhook_url", "") or DEFAULT_WECOM_WEBHOOK_URL).strip()

    def _wecom_webhook_note(self):
        return (getattr(self, "wecom_webhook_note", "") or DEFAULT_WECOM_WEBHOOK_NOTE).strip()

    def _smart_bot_config(self):
        return (
            (getattr(self, "wecom_smart_bot_id", "") or "").strip(),
            (getattr(self, "wecom_smart_bot_secret", "") or "").strip(),
            self._smart_chat_id_value(),
        )

    def _smart_chat_id_value(self):
        if getattr(self, "wecom_smart_chat_id_combo", None) is not None:
            return self.wecom_smart_chat_id_combo.currentText().strip()
        return (getattr(self, "wecom_smart_chat_id", "") or "").strip()

    def _refresh_smart_chat_id_combo(self):
        combo = getattr(self, "wecom_smart_chat_id_combo", None)
        if combo is None:
            return
        current_text = combo.currentText().strip() or (getattr(self, "wecom_smart_chat_id", "") or "").strip()
        combo.blockSignals(True)
        combo.clear()
        for item in getattr(self, "wecom_smart_chat_id_history", []) or []:
            combo.addItem(item)
        combo.setCurrentText(current_text)
        combo.blockSignals(False)

    def _remember_smart_chat_id(self, chat_id):
        chat_id = (chat_id or "").strip()
        if not chat_id:
            return
        history = [item for item in (getattr(self, "wecom_smart_chat_id_history", []) or []) if (item or "").strip()]
        history = [item for item in history if item != chat_id]
        history.insert(0, chat_id)
        self.wecom_smart_chat_id_history = history[:12]
        self.wecom_smart_chat_id = chat_id
        self._refresh_smart_chat_id_combo()
        self._save_settings()

    def _ensure_wecom_webhook(self):
        webhook = self._wecom_webhook()
        if webhook:
            return webhook
        self._show_feedback("请先在“更多”里设置企业微信群机器人的 Webhook。", "info", timeout_ms=3500)
        self.show_wecom_window()
        return ""

    def _ensure_smart_bot_config(self):
        bot_id, secret, chat_id = self._smart_bot_config()
        missing = wecom_delivery.missing_smart_bot_fields(bot_id, secret, chat_id)
        if not missing:
            return bot_id, secret, chat_id
        self._show_feedback(
            f"请先在“更多”里补全智能机器人设置：{", ".join(missing)}。",
            "info",
            timeout_ms=4000,
        )
        self.show_wecom_window()
        return None

    def _send_text_to_wecom(self, text):
        if self._wecom_transport() == "smart_bot":
            config = self._ensure_smart_bot_config()
            if not config:
                return False
            bot_id, secret, chat_id = config
            try:
                self.smart_bot_sender.send_markdown(bot_id, secret, chat_id, text)
                self._remember_smart_chat_id(chat_id)
                return True
            except concurrent.futures.TimeoutError:
                self._show_feedback("智能机器人发送超时，请检查网络或机器人配置。", "error", timeout_ms=5000)
                return False
            except Exception as exc:
                self._show_feedback(f"智能机器人发送失败：{exc}", "error", timeout_ms=5000)
                return False

        webhook = self._ensure_wecom_webhook()
        if not webhook:
            return False
        payload = json.dumps(
            {"msgtype": "text", "text": {"content": text}},
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            webhook,
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=10) as response:
                response_body = response.read().decode("utf-8", errors="ignore")
        except urllib.error.URLError as exc:
            self._show_feedback(f"发送到企业微信群机器人失败：{exc}", "error", timeout_ms=5000)
            return False

        try:
            result = json.loads(response_body or "{}")
        except json.JSONDecodeError:
            self._show_feedback("企业微信消息发送失败，请检查机器人配置或网络。", "error", timeout_ms=4000)
            return False

        if result.get("errcode") not in (0, "0", None):
            self._show_feedback(f"企业微信返回错误：{result.get('errmsg') or result}", "error", timeout_ms=5000)
            return False
        return True

    def send_current_to_wecom(self):
        text = self.current_result_text()
        if not text:
            self._show_feedback("当前还没有可发送的申请模板。", "info", timeout_ms=2500)
            return
        if self._send_text_to_wecom(text):
            if self._wecom_transport() == "smart_bot":
                message = "当前模板已发送到企业微信智能机器人目标会话。"
            else:
                message = "当前模板已发送到企业微信群机器人。"
            self._show_feedback(message, "success", timeout_ms=3000)
            self._push_ticker_message(message)

    def send_current_to_wecom_and_next(self):
        texts = (self.last_summary or {}).get("texts") or []
        text = self.current_result_text()
        if not text:
            self._show_feedback("当前还没有可发送的申请模板。", "info", timeout_ms=2500)
            return
        if not self._send_text_to_wecom(text):
            return
        if texts and self.current_result_index < len(texts) - 1:
            self.current_result_index += 1
            self._render_current_result()
            message = "当前模板已发送，并已切到下一条。"
            self._show_feedback(message, "success", timeout_ms=3000)
            self._push_ticker_message(message)
            return
        message = "当前模板已发送，当前已经是最后一条。"
        self._show_feedback(message, "success", timeout_ms=3000)
        self._push_ticker_message(message)

    def _update_auto_send_button(self):
        if not hasattr(self, "auto_send_wecom_btn"):
            return
        self._refresh_robot_button_labels()
        self.auto_send_wecom_btn.setText(
            getattr(self, "_auto_send_pause_text", "暂停发送")
            if self.auto_send_running
            else getattr(self, "_auto_send_idle_text", "自动发送")
        )

    def _stop_auto_send_wecom(self, message=None, tone="info"):
        self.auto_send_timer.stop()
        self.auto_send_running = False
        self._update_auto_send_button()
        if message:
            self._show_feedback(message, tone, timeout_ms=3000 if tone == "success" else 4000)
            self._push_ticker_message(message)

    def toggle_auto_send_wecom(self):
        if self.auto_send_running:
            self._stop_auto_send_wecom("已暂停自动发送。", "info")
            return
        texts = (self.last_summary or {}).get("texts") or []
        if not texts:
            self._show_feedback("当前还没有可发送的申请模板。", "info", timeout_ms=2500)
            return
        if not self._ensure_wecom_webhook():
            return
        self.auto_send_running = True
        self._update_auto_send_button()
        self._push_ticker_message("自动发送已开始。")
        self._advance_auto_send_wecom()

    def _advance_auto_send_wecom(self):
        if not self.auto_send_running:
            return
        texts = (self.last_summary or {}).get("texts") or []
        if not texts:
            self._stop_auto_send_wecom("当前没有可继续发送的模板。", "info")
            return
        text = self.current_result_text()
        if not text:
            self._stop_auto_send_wecom("当前没有可继续发送的模板。", "info")
            return
        if not self._send_text_to_wecom(text):
            self._stop_auto_send_wecom("自动发送过程中出现错误，请检查机器人配置。", "error")
            return
        if self.current_result_index >= len(texts) - 1:
            self._stop_auto_send_wecom("当前模板已经全部发送完成。", "success")
            return
        self.current_result_index += 1
        self._render_current_result()
        interval_ms = max(200, int(float(getattr(self, "wecom_send_interval_seconds", 1.0) or 1.0) * 1000))
        self.auto_send_timer.start(interval_ms)

    def _render_current_result(self):
        texts = (self.last_summary or {}).get("texts") or []
        total = len(texts)
        if not total:
            self.result_box.setPlainText("当前还没有可显示的模板。")
            self.result_index_label.setText("0 / 0")
            self.result_index_label.setVisible(False)
            self.prev_result_btn.setVisible(False)
            self.next_result_btn.setVisible(False)
            self.prev_result_btn.setDisabled(True)
            self.next_result_btn.setDisabled(True)
            self.copy_result_btn.setDisabled(True)
            self.copy_next_result_btn.setDisabled(True)
            self.send_wecom_btn.setDisabled(True)
            self.send_next_wecom_btn.setDisabled(True)
            self.auto_send_wecom_btn.setDisabled(True)
            return

        self.current_result_index = max(0, min(self.current_result_index, total - 1))
        self.result_box.setPlainText(texts[self.current_result_index])
        self.result_index_label.setText(f"{self.current_result_index + 1} / {total}")
        self.result_index_label.setVisible(total > 1)
        self._position_result_index_label()
        manual_mode = getattr(self, "action_mode", "manual") != "robot"
        self.prev_result_btn.setVisible(total > 1 and manual_mode)
        self.next_result_btn.setVisible(total > 1 and manual_mode)
        self.prev_result_btn.setDisabled(self.current_result_index <= 0)
        self.next_result_btn.setDisabled(self.current_result_index >= total - 1)
        self.copy_result_btn.setDisabled(False)
        self.copy_next_result_btn.setDisabled(False)
        self.send_wecom_btn.setDisabled(False)
        self.send_next_wecom_btn.setDisabled(False)
        self.auto_send_wecom_btn.setDisabled(False)

    def show_prev_result(self):
        texts = (self.last_summary or {}).get("texts") or []
        if not texts:
            return
        self.current_result_index = max(0, self.current_result_index - 1)
        self._render_current_result()
        self._push_ticker_message(f"当前查看第 {self.current_result_index + 1} / {len(texts)} 条申请模板。")

    def show_next_result(self):
        texts = (self.last_summary or {}).get("texts") or []
        if not texts:
            return
        self.current_result_index = min(len(texts) - 1, self.current_result_index + 1)
        self._render_current_result()
        self._push_ticker_message(f"当前查看第 {self.current_result_index + 1} / {len(texts)} 条申请模板。")

    def open_folder(self, path):
        os.makedirs(path, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def clear_log(self):
        self.log_box.clear()
        if self.log_view_box is not None:
            self.log_view_box.clear()
        self._show_feedback("已经是第一条。", "info", timeout_ms=2500)
        self._push_ticker_message("已经是第一条。")

    def toggle_help_panel(self, force_visible=None):
        self.show_help_window()

    def toggle_log_panel(self, force_visible=None):
        self.show_log_window()

    def _advance_busy_state(self):
        dots = "." * (self.busy_step % 3 + 1)
        if self.current_action_name and hasattr(self, "generate_btn"):
            self.generate_btn.setText(f"{self.current_action_name}{dots}")
        self.busy_step += 1

    def validate_before_generate(self):
        if not os.path.isdir(self.csv_dir()):
            self._show_feedback("表格文件夹不存在，请先重新选择。", "error", timeout_ms=4000)
            return False
        if self.worker_thread is not None:
            self._show_feedback("当前目录里没有可处理的表格文件，请先放入或选择表格。", "info", timeout_ms=3500)
            return False
        return True

    def set_running(self, running, message):
        self.more_button.setDisabled(running)
        self._set_primary_busy(running)
        self.select_csv_files_btn.setDisabled(running)
        self.use_all_csv_btn.setDisabled(running or len(self.selected_csv_paths_in_dir()) == 0)
        self.csv_dir_edit.setDisabled(running)
        self.store_name_edit.setDisabled(running)
        self.owned_store_edit.setDisabled(running)
        self.goods_name_edit.setDisabled(running)
        self.applicant_edit.setDisabled(running)
        if hasattr(self, "copy_result_btn"):
            self.copy_result_btn.setDisabled(running)
        if hasattr(self, "copy_next_result_btn"):
            self.copy_next_result_btn.setDisabled(running)
        if hasattr(self, "send_wecom_btn"):
            self.send_wecom_btn.setDisabled(running)
        if hasattr(self, "send_next_wecom_btn"):
            self.send_next_wecom_btn.setDisabled(running)
        if hasattr(self, "auto_send_wecom_btn"):
            self.auto_send_wecom_btn.setDisabled(running)
        if hasattr(self, "prev_result_btn"):
            self.prev_result_btn.setDisabled(running)
        if hasattr(self, "next_result_btn"):
            self.next_result_btn.setDisabled(running)
        if running:
            self.busy_step = 0
            self.busy_timer.start()
            self._advance_busy_state()
            self._sync_status_labels(state="失败", result=message)
            return
        self.busy_timer.stop()
        if hasattr(self, "generate_btn"):
            self.generate_btn.setText("开始生成申请模板")
        self._sync_status_labels(state="完成", result=message)

    def start_generate(self):
        if not self.validate_before_generate():
            return
        self._save_settings()
        self.start_worker("生成申请模板", self.run_generate_job)

    def start_worker(self, action_name, task):
        self.current_action_name = action_name
        self.log_box.appendPlainText(f"\n=== {action_name} ===\n")
        self._push_ticker_message(f"{action_name}已开始。")
        self.set_running(True, f"{action_name}中...")

        self.worker_thread = QThread(self)
        self.worker = Worker(action_name, task)
        self.worker.moveToThread(self.worker_thread)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(self._clear_worker_refs)
        self.worker_thread.start()

    def run_generate_job(self, logger):
        summary = generator.generate_request_texts(
            csv_dir=self.csv_dir(),
            settings=self.settings_payload(),
            logger=logger,
            csv_paths=self.selected_csv_paths_in_dir(),
        )
        logger(
            f"生成完成：共读取 {summary['csv_files']} 个表格文件，处理 {summary['rows']} 行，生成 {summary['generated']} 份文字模板。"
        )
        return summary

    @Slot(str)
    def append_log(self, text):
        if not text:
            return
        self.log_box.insertPlainText(text + ("" if text.endswith("\n") else "\n"))
        self.log_box.ensureCursorVisible()
        if self.log_view_box is not None:
            self.log_view_box.setPlainText(self.log_box.toPlainText())
            self.log_view_box.moveCursor(self.log_view_box.textCursor().End)
        for line in str(text).splitlines():
            if line.strip():
                self._push_ticker_message(line.strip())

    @Slot(bool, str, object)
    def on_worker_finished(self, success, message, result):
        self.current_action_name = ""
        self.busy_timer.stop()
        self._set_primary_busy(False)
        if hasattr(self, "generate_btn"):
            self.generate_btn.setText("开始生成申请模板")
        self.more_button.setDisabled(False)
        if hasattr(self, "copy_result_btn"):
            self.copy_result_btn.setDisabled(False)
        if hasattr(self, "copy_next_result_btn"):
            self.copy_next_result_btn.setDisabled(False)
        if hasattr(self, "send_wecom_btn"):
            self.send_wecom_btn.setDisabled(False)
        if hasattr(self, "send_next_wecom_btn"):
            self.send_next_wecom_btn.setDisabled(False)
        if hasattr(self, "auto_send_wecom_btn"):
            self.auto_send_wecom_btn.setDisabled(False)
        self.last_summary = result

        if success and result:
            self.current_result_index = 0
            self._render_current_result()
            summary_message = f"已生成 {result['generated']} 份申请模板，可直接复制。"
            self._sync_status_labels(state="已完成", result=summary_message)
            self._push_ticker_message(summary_message)
            self._show_feedback(summary_message, "success", timeout_ms=5000)
            return

        self._render_current_result()
        self._sync_status_labels(state="失败", result=message)
        self._push_ticker_message(message)
        self._show_feedback(f"{message} 你可以在“更多”里打开日志查看详细原因。", "error", timeout_ms=7000)

    @Slot()
    def _clear_worker_refs(self):
        self.worker = None
        self.worker_thread = None

    def closeEvent(self, event):
        self.auto_send_timer.stop()
        self.auto_send_running = False
        self.smart_bot_sender.close()
        self._save_settings()
        super().closeEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._position_feedback_banner()
        self._position_result_index_label()

    def dragEnterEvent(self, event):
        if self._dropped_csv_paths(event.mimeData()):
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if self._dropped_csv_paths(event.mimeData()):
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        csv_paths = self._dropped_csv_paths(event.mimeData())
        if csv_paths:
            event.acceptProposedAction()
            self.import_dropped_csv_files(csv_paths)
            return
        super().dropEvent(event)

    def eventFilter(self, watched, event):
        if event.type() in (QEvent.DragEnter, QEvent.DragMove, QEvent.Drop):
            csv_paths = self._dropped_csv_paths(event.mimeData())
            if csv_paths:
                event.acceptProposedAction()
                if event.type() == QEvent.Drop:
                    self.import_dropped_csv_files(csv_paths)
                return True
        return super().eventFilter(watched, event)

    def restore_default_settings(self):
        self.store_name_edit.setCurrentText(generator.DEFAULT_STORE_NAME)
        self._sync_owned_store_for_selected_store(preferred_owned=generator.DEFAULT_OWNED_STORE)
        self.goods_name_edit.setText(generator.DEFAULT_GOODS_NAME)
        self.applicant_edit.setText(generator.DEFAULT_APPLICANT)
        self._save_settings()
        self.refresh_preview(push_message=False)
        self._show_feedback("默认参数已恢复。", "success", timeout_ms=2500)

    def save_wecom_settings(self):
        if self.wecom_window is None:
            return
        if self.wecom_mode_combo is not None:
            self.wecom_transport_mode = self.wecom_mode_combo.currentData() or "webhook"
        if self.wecom_webhook_edit is not None:
            self.wecom_webhook_url = self.wecom_webhook_edit.text().strip() or DEFAULT_WECOM_WEBHOOK_URL
        if self.wecom_webhook_note_edit is not None:
            self.wecom_webhook_note = self.wecom_webhook_note_edit.text().strip() or DEFAULT_WECOM_WEBHOOK_NOTE
        if self.wecom_smart_bot_id_edit is not None:
            self.wecom_smart_bot_id = self.wecom_smart_bot_id_edit.text().strip()
        if self.wecom_smart_bot_secret_edit is not None:
            self.wecom_smart_bot_secret = self.wecom_smart_bot_secret_edit.text().strip()
        self.wecom_smart_chat_id = self._smart_chat_id_value()
        if self.wecom_interval_spin is not None:
            self.wecom_send_interval_seconds = float(self.wecom_interval_spin.value())
        self._save_settings()
        self._refresh_robot_button_labels()
        self._update_auto_send_button()
        if self._wecom_transport() == "smart_bot":
            self._show_feedback("企业微信智能机器人设置已保存。", "success", timeout_ms=2500)
        else:
            self._show_feedback("企业微信群机器人 Webhook 已保存。", "success", timeout_ms=2500)
    def send_wecom_test_message(self):
        self.save_wecom_settings()
        test_text = "发票申请模版产出工具测试消息\n如果你看到这条消息，说明机器人已接通。"
        if self._send_text_to_wecom(test_text):
            if self._wecom_transport() == "smart_bot":
                self._show_feedback("测试消息已发送到企业微信智能机器人目标会话。", "success", timeout_ms=3000)
            else:
                self._show_feedback("测试消息已发送到企业微信群。", "success", timeout_ms=3000)

    def on_wecom_mode_changed(self, _index):
        if self.wecom_mode_combo is None:
            return
        self.wecom_transport_mode = self.wecom_mode_combo.currentData() or "webhook"
        self._apply_wecom_setting_panel_mode()
        self._refresh_robot_button_labels()
        self._update_auto_send_button()

    def _apply_wecom_setting_panel_mode(self):
        smart_mode = self._wecom_transport() == "smart_bot"
        if self.wecom_webhook_panel is not None:
            self.wecom_webhook_panel.setVisible(not smart_mode)
        if self.wecom_smart_panel is not None:
            self.wecom_smart_panel.setVisible(smart_mode)
        if self.wecom_window is not None:
            self.wecom_window.resize(560, 286 if smart_mode else 220)

    def show_wecom_window(self):
        if self.wecom_window is None:
            window = QWidget(None, Qt.Window)
            window.setWindowTitle("企业微信机器人设置")
            window.setWindowIcon(self.window_icon)
            window.resize(560, 220)
            layout = QVBoxLayout(window)
            layout.setContentsMargins(12, 12, 12, 12)
            layout.setSpacing(8)

            title = QLabel("企业微信机器人")
            title.setObjectName("SectionTitle")
            layout.addWidget(title)

            mode_row = QHBoxLayout()
            mode_row.setContentsMargins(0, 0, 0, 0)
            mode_row.setSpacing(8)
            mode_label = QLabel("接入方式")
            mode_label.setObjectName("FieldLabel")
            mode_row.addWidget(mode_label)
            mode_combo = QComboBox()
            mode_combo.setObjectName("PathInput")
            mode_combo.addItem("群机器人 Webhook", "webhook")
            mode_combo.addItem("智能机器人 Bot ID + Secret", "smart_bot")
            mode_index = max(0, mode_combo.findData(self._wecom_transport()))
            mode_combo.setCurrentIndex(mode_index)
            mode_combo.currentIndexChanged.connect(self.on_wecom_mode_changed)
            mode_row.addWidget(mode_combo, 1)
            layout.addLayout(mode_row)

            hint = QLabel("Webhook 适合固定群推送；智能机器人模式需要填写 Bot ID、Secret 和目标会话 Chat ID。")
            hint.setObjectName("HintLabel")
            hint.setWordWrap(True)
            layout.addWidget(hint)

            webhook_panel = QWidget()
            webhook_layout = QVBoxLayout(webhook_panel)
            webhook_layout.setContentsMargins(0, 0, 0, 0)
            webhook_layout.setSpacing(6)

            webhook_edit = QLineEdit()
            webhook_edit.setObjectName("PathInput")
            webhook_edit.setPlaceholderText("https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=...")
            webhook_edit.setText(self._wecom_webhook())
            webhook_layout.addWidget(webhook_edit)

            webhook_note_edit = QLineEdit()
            webhook_note_edit.setObjectName("PathInput")
            webhook_note_edit.setPlaceholderText("机器人备注")
            webhook_note_edit.setText(self._wecom_webhook_note())
            webhook_layout.addWidget(webhook_note_edit)
            layout.addWidget(webhook_panel)

            smart_panel = QWidget()
            smart_layout = QVBoxLayout(smart_panel)
            smart_layout.setContentsMargins(0, 0, 0, 0)
            smart_layout.setSpacing(6)

            bot_id_edit = QLineEdit()
            bot_id_edit.setObjectName("PathInput")
            bot_id_edit.setPlaceholderText("Bot ID")
            bot_id_edit.setText((getattr(self, "wecom_smart_bot_id", "") or "").strip())
            smart_layout.addWidget(bot_id_edit)

            secret_edit = QLineEdit()
            secret_edit.setObjectName("PathInput")
            secret_edit.setPlaceholderText("Secret")
            secret_edit.setEchoMode(QLineEdit.Password)
            secret_edit.setText((getattr(self, "wecom_smart_bot_secret", "") or "").strip())
            smart_layout.addWidget(secret_edit)

            chat_id_combo = QComboBox()
            chat_id_combo.setObjectName("PathInput")
            chat_id_combo.setEditable(True)
            chat_id_combo.setInsertPolicy(QComboBox.NoInsert)
            chat_id_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            if chat_id_combo.lineEdit() is not None:
                chat_id_combo.lineEdit().setPlaceholderText("目标会话 Chat ID（单聊 userid / 群聊 chatid）")
            smart_layout.addWidget(chat_id_combo)

            smart_hint = QLabel("智能机器人主动发消息需要目标会话 Chat ID；如果你发给群，就填群聊 chatid。")
            smart_hint.setObjectName("HintLabel")
            smart_hint.setWordWrap(True)
            smart_layout.addWidget(smart_hint)

            if not wecom_delivery.smart_bot_sdk_available():
                sdk_hint = QLabel("当前环境缺少 wecom-aibot-sdk，智能机器人模式暂不可用。")
                sdk_hint.setObjectName("HintLabel")
                sdk_hint.setWordWrap(True)
                smart_layout.addWidget(sdk_hint)

            layout.addWidget(smart_panel)

            interval_row = QHBoxLayout()
            interval_row.setContentsMargins(0, 0, 0, 0)
            interval_row.setSpacing(8)

            interval_label = QLabel("自动发送间隔")
            interval_label.setObjectName("FieldLabel")
            interval_row.addWidget(interval_label)

            interval_spin = QDoubleSpinBox()
            interval_spin.setObjectName("PathInput")
            interval_spin.setDecimals(1)
            interval_spin.setMinimum(0.2)
            interval_spin.setMaximum(60.0)
            interval_spin.setSingleStep(0.2)
            interval_spin.setSuffix(" 秒")
            interval_spin.setValue(float(getattr(self, "wecom_send_interval_seconds", 1.0) or 1.0))
            interval_row.addWidget(interval_spin)
            interval_row.addStretch(1)
            layout.addLayout(interval_row)

            buttons = QHBoxLayout()
            buttons.setContentsMargins(0, 0, 0, 0)
            buttons.setSpacing(6)
            buttons.addStretch(1)

            save_btn = QPushButton("保存")
            save_btn.setProperty("role", "subtle")
            save_btn.setObjectName("CompactButton")
            save_btn.setFixedWidth(82)
            save_btn.clicked.connect(self.save_wecom_settings)
            buttons.addWidget(save_btn)

            test_btn = QPushButton("发送测试")
            test_btn.setProperty("role", "subtle")
            test_btn.setObjectName("CompactButton")
            test_btn.setFixedWidth(92)
            test_btn.clicked.connect(self.send_wecom_test_message)
            buttons.addWidget(test_btn)

            layout.addLayout(buttons)

            self.wecom_window = window
            self.wecom_mode_combo = mode_combo
            self.wecom_webhook_edit = webhook_edit
            self.wecom_webhook_note_edit = webhook_note_edit
            self.wecom_smart_bot_id_edit = bot_id_edit
            self.wecom_smart_bot_secret_edit = secret_edit
            self.wecom_smart_chat_id_combo = chat_id_combo
            self.wecom_webhook_panel = webhook_panel
            self.wecom_smart_panel = smart_panel
            self.wecom_interval_spin = interval_spin
        else:
            if self.wecom_mode_combo is not None:
                mode_index = max(0, self.wecom_mode_combo.findData(self._wecom_transport()))
                self.wecom_mode_combo.setCurrentIndex(mode_index)
            self.wecom_webhook_edit.setText(self._wecom_webhook())
            if self.wecom_smart_bot_id_edit is not None:
                self.wecom_smart_bot_id_edit.setText((getattr(self, "wecom_smart_bot_id", "") or "").strip())
            if self.wecom_smart_bot_secret_edit is not None:
                self.wecom_smart_bot_secret_edit.setText((getattr(self, "wecom_smart_bot_secret", "") or "").strip())
            if self.wecom_interval_spin is not None:
                self.wecom_interval_spin.setValue(float(getattr(self, "wecom_send_interval_seconds", 1.0) or 1.0))
        self._refresh_smart_chat_id_combo()
        self._apply_wecom_setting_panel_mode()
        self.wecom_window.show()
        self.wecom_window.raise_()
        self.wecom_window.activateWindow()

    def show_help_window(self):
        if self.help_window is None:
            window = QWidget(None, Qt.Window)
            window.setWindowTitle("使用帮助")
            window.setWindowIcon(self.window_icon)
            window.resize(520, 420)
            layout = QVBoxLayout(window)
            layout.setContentsMargins(12, 12, 12, 12)
            layout.setSpacing(8)
            title = QLabel("使用帮助")
            title.setObjectName("SectionTitle")
            layout.addWidget(title)
            text_box = QPlainTextEdit()
            text_box.setReadOnly(True)
            text_box.setObjectName("LogBox")
            text_box.setPlainText(load_guide_text())
            layout.addWidget(text_box, 1)
            self.help_window = window
            self.help_text_box = text_box
        else:
            self.help_text_box.setPlainText(load_guide_text())
        self.help_window.show()
        self.help_window.raise_()
        self.help_window.activateWindow()

    def show_log_window(self):
        if self.log_window is None:
            window = QWidget(None, Qt.Window)
            window.setWindowTitle("处理日志")
            window.setWindowIcon(self.window_icon)
            window.resize(680, 480)
            layout = QVBoxLayout(window)
            layout.setContentsMargins(12, 12, 12, 12)
            layout.setSpacing(8)
            header = QHBoxLayout()
            title = QLabel("处理日志")
            title.setObjectName("SectionTitle")
            header.addWidget(title)
            header.addStretch(1)
            clear_btn = QPushButton("清空")
            clear_btn.setProperty("role", "subtle")
            clear_btn.setObjectName("CompactButton")
            clear_btn.clicked.connect(self.clear_log)
            header.addWidget(clear_btn)
            layout.addLayout(header)
            log_view = QPlainTextEdit()
            log_view.setReadOnly(True)
            log_view.setObjectName("LogBox")
            layout.addWidget(log_view, 1)
            self.log_window = window
            self.log_view_box = log_view
        self.log_view_box.setPlainText(self.log_box.toPlainText())
        self.log_window.show()
        self.log_window.raise_()
        self.log_window.activateWindow()


def _main_window_refresh_preview(self, push_message=True):
    try:
        csv_paths = self.current_csv_paths()
        selected_count = len(self.selected_csv_paths_in_dir())
        if selected_count:
            self.csv_selection_label.setText(f"已选 {selected_count} 个")
        else:
            self.csv_selection_label.setText(f"全部表格（{len(csv_paths)}）")
        self.use_all_csv_btn.setDisabled(selected_count == 0)

        summary = generator.generate_request_texts(
            csv_dir=self.csv_dir(),
            settings=self.settings_payload(),
            csv_paths=self.selected_csv_paths_in_dir(),
        )
        self.csv_preview_label.setText(f"{summary['csv_files']}个表格文件，可生成{summary['generated']}份。")
        self.generate_btn.setDisabled(summary["generated"] == 0)
        if summary["texts"]:
            self.current_result_index = 0
            self.last_summary = summary
            self._render_current_result()
        else:
            self.last_summary = {"texts": []}
            self.result_box.setPlainText("请选择表格文件夹。")
            self.result_index_label.setText("0 / 0")
            self.result_index_label.setVisible(False)
            self.prev_result_btn.setVisible(False)
            self.next_result_btn.setVisible(False)
            self.prev_result_btn.setDisabled(True)
            self.next_result_btn.setDisabled(True)
            self.copy_result_btn.setDisabled(True)

        if push_message:
            self._push_ticker_message(
                f"{summary['csv_files']} 个表格文件，可生成 {summary['generated']} 份。"
            )
        self._sync_status_labels(
            result=f"{summary['csv_files']} 个表格文件，可生成 {summary['generated']} 份。"
        )
    except Exception as exc:
        self.csv_selection_label.setText("读取表格失败")
        self.use_all_csv_btn.setDisabled(True)
        self.generate_btn.setDisabled(True)
        self.csv_preview_label.setText(str(exc))
        self.result_box.setPlainText("结果会在读取 CSV 后显示在这里。")
        self.result_index_label.setText("0 / 0")
        self.result_index_label.setVisible(False)
        self.last_summary = None
        self.prev_result_btn.setVisible(False)
        self.next_result_btn.setVisible(False)
        self.prev_result_btn.setDisabled(True)
        self.next_result_btn.setDisabled(True)
        self.copy_result_btn.setDisabled(True)
        if push_message:
            self._push_ticker_message("当前目录里没有找到表格文件。")
        self._sync_status_labels(result="当前目录里没有找到表格文件。")


MainWindow.refresh_preview = _main_window_refresh_preview


def run_app():
    set_windows_app_id()
    app = QApplication(sys.argv)
    app.setWindowIcon(load_app_icon())
    instance_guard = SingleInstanceGuard(SINGLE_INSTANCE_KEY)
    if not instance_guard.acquire():
        return 0
    app.aboutToQuit.connect(instance_guard.release)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    try:
        sys.exit(run_app())
    except Exception:
        with open(ERROR_LOG_PATH, "a", encoding="utf-8") as file_obj:
            file_obj.write(traceback.format_exc())
            file_obj.write("\n")
        raise

