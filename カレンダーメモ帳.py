# -*- coding: utf-8 -*-
"""
カレンダー型メモ帳（連続日付ビュー・無限スクロール版）

主な変更点
- 月区切りなしの連続表示を、スクロール端で「前後に日付を動的生成」して擬似的に無限化
- 初期表示は約2ヶ月（62日）。上下端に近づくと自動で±60日ずつ拡張
- 既存機能は維持（設定の永続化/ダークテーマ/列管理/自動保存/URL強調＆ダブルクリック/祝日・週末配色/検索）
- 検索の「表示範囲」は現在テーブルに展開済みの範囲を指します。「すべて」は保存済み全日付

注: QTableWidget は真の“無限”は持てないため、必要に応じて前後へ行を追加して実質的に無限にスクロールできるようにしています。
"""

import json
import re
import sys
import uuid
import webbrowser
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Dict, List, Set

from PyQt6.QtCore import Qt, QDate, pyqtSignal, QTimer
from PyQt6.QtGui import (
    QAction, QFont, QSyntaxHighlighter,
    QTextCharFormat, QColor, QTextCursor, QCloseEvent, QGuiApplication
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QTextEdit, QPushButton, QLabel,
    QSpinBox, QLineEdit, QFileDialog, QListWidget, QListWidgetItem,
    QMessageBox, QAbstractItemView, QTextBrowser, QDialog, QFormLayout,
    QTabWidget, QGroupBox, QTreeWidget, QTreeWidgetItem, QCalendarWidget,
    QHBoxLayout, QRadioButton, QCheckBox
)

# --------- 定数/ユーティリティ ----------
JP_WEEK = ["日", "月", "火", "水", "木", "金", "土"]
URL_RE = re.compile(r"https?://\S+", re.IGNORECASE)
SETTINGS_PATH = Path.home() / ".calendar_notes_settings.json"
DEFAULT_AUTOSAVE = Path.home() / "calendar-notes_autosave.json"

def pad2(n: int) -> str:
    return str(n).zfill(2)

def parse_holidays_str(s: str) -> Set[str]:
    vals = set()
    if not s:
        return vals
    for token in re.split(r"[,\s]+", s.strip()):
        t = token.strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", t):
            vals.add(t)
    return vals

def date_key(dt: date) -> str:
    return dt.strftime("%Y-%m-%d")


# ---------- URLハイライター ----------
class UrlHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor("blue"))
        fmt.setFontUnderline(True)
        self.format = fmt

    def highlightBlock(self, text: str):
        for m in URL_RE.finditer(text):
            start, end = m.span()
            self.setFormat(start, end - start, self.format)


# ---------- 自動リサイズ付きテキストエディタ ----------
class AutoResizeTextEdit(QTextEdit):
    heightChanged = pyqtSignal(int)
    requestUrlList = pyqtSignal(list)  # URL一覧ダイアログ表示要求

    def __init__(self, *args, base_font_pt=11, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptRichText(False)
        self.setFont(QFont("Meiryo UI", base_font_pt))
        self._padding_px = 6
        self.highlighter = UrlHighlighter(self.document())
        self.textChanged.connect(self._auto_resize)
        self._auto_resize()

        # 右クリックメニュー
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.ActionsContextMenu)
        act_preview = QAction("Markdownプレビュー", self)
        act_openurl = QAction("最初のURLを開く", self)
        act_urllist = QAction("このセルのURL一覧…", self)
        act_preview.triggered.connect(self.preview_markdown)
        act_openurl.triggered.connect(self.open_first_url)
        act_urllist.triggered.connect(self.open_url_list)
        self.addAction(act_preview)
        self.addAction(act_openurl)
        self.addAction(act_urllist)

    def mouseDoubleClickEvent(self, ev):
        cursor = self.cursorForPosition(ev.pos())
        cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        text = cursor.selectedText()
        if text.startswith("http://") or text.startswith("https://"):
            webbrowser.open(text)
        else:
            super().mouseDoubleClickEvent(ev)

    def setPointSize(self, pt: int):
        f = self.font()
        f.setPointSize(pt)
        self.setFont(f)
        self._auto_resize()

    def _calc_height(self) -> int:
        doc = self.document()
        doc.setTextWidth(self.viewport().width())
        doc_h = int(doc.documentLayout().documentSize().height())
        return max(36, min(10000, doc_h + self._padding_px))

    def _auto_resize(self):
        h = self._calc_height()
        if self.height() != h:
            self.setFixedHeight(h)
        self.heightChanged.emit(h)

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        self._auto_resize()

    def preview_markdown(self):
        dlg = QTextBrowser()
        dlg.setWindowTitle("Markdownプレビュー")
        dlg.setMarkdown(self.toPlainText() or "_(内容なし)_")
        dlg.setMinimumSize(480, 360)
        dlg.show()
        self._preview_holder = dlg

    def _extract_urls(self) -> List[str]:
        return URL_RE.findall(self.toPlainText() or "")

    def open_first_url(self):
        urls = self._extract_urls()
        if urls:
            webbrowser.open(urls[0])
        else:
            QMessageBox.information(self, "URL", "URLが見つかりません。")

    def open_url_list(self):
        urls = self._extract_urls()
        self.requestUrlList.emit(urls)


# ---------- URL一覧ダイアログ ----------
class UrlListDialog(QDialog):
    def __init__(self, urls: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("URL一覧")
        self.setMinimumSize(560, 360)
        v = QVBoxLayout(self)
        self.list = QListWidget()
        for u in urls:
            self.list.addItem(QListWidgetItem(u))
        v.addWidget(self.list, 1)

        h = QHBoxLayout()
        btn_open = QPushButton("開く")
        btn_copy = QPushButton("コピー")
        btn_close = QPushButton("閉じる")
        h.addWidget(btn_open); h.addWidget(btn_copy); h.addWidget(btn_close)
        v.addLayout(h)

        btn_open.clicked.connect(self._open)
        btn_copy.clicked.connect(self._copy)
        btn_close.clicked.connect(self.accept)

    def _current_url(self) -> str | None:
        it = self.list.currentItem()
        return it.text() if it else None

    def _open(self):
        u = self._current_url()
        if u: webbrowser.open(u)

    def _copy(self):
        u = self._current_url()
        if u:
            QGuiApplication.clipboard().setText(u)
            QMessageBox.information(self, "コピー", "URLをクリップボードにコピーしました。")


# ---------- 設定ダイアログ（即時反映＋永続化） ----------
class SettingsDialog(QDialog):
    def __init__(self, app: "CalendarApp"):
        super().__init__(app)
        self.app = app
        self.setWindowTitle("設定")
        self.setMinimumSize(720, 520)

        tabs = QTabWidget(self)

        # 表示設定
        w_disp = QWidget()
        f_disp = QFormLayout(w_disp)
        self.spin_font = QSpinBox()
        self.spin_font.setRange(9, 28)
        self.spin_font.setValue(app.font_pt)
        self.spin_font.valueChanged.connect(self._on_font_change)
        f_disp.addRow("文字サイズ (pt)", self.spin_font)

        self.chk_dark = QCheckBox("ダークテーマを有効にする")
        self.chk_dark.setChecked(self.app.settings.get("theme", "light") == "dark")
        self.chk_dark.toggled.connect(self._on_theme_toggle)
        f_disp.addRow(self.chk_dark)

        self.ed_holidays = QLineEdit()
        self.ed_holidays.setPlaceholderText("例) 2025-01-01, 2025-05-03 ...")
        self.ed_holidays.setText(app.settings.get("holidays", ""))
        self.ed_holidays.editingFinished.connect(self._on_holidays_change)
        f_disp.addRow("祝日 (YYYY-MM-DD, カンマ区切り)", self.ed_holidays)
        tabs.addTab(w_disp, "表示設定")

        # 列管理
        w_cols = QWidget()
        v_cols = QVBoxLayout(w_cols)
        hl = QHBoxLayout()
        v_cols.addLayout(hl)
        self.list_cols = QListWidget()
        self.list_cols.setMinimumWidth(260)
        hl.addWidget(self.list_cols)

        op = QVBoxLayout()
        self.ed_col_title = QLineEdit()
        btn_set_title = QPushButton("名称変更")
        btn_add = QPushButton("+ 列を追加")
        btn_del = QPushButton("選択列を削除")
        self.spin_col_width = QSpinBox()
        self.spin_col_width.setRange(160, 800)
        btn_set_width = QPushButton("幅を反映")
        op.addWidget(QLabel("列名"))
        op.addWidget(self.ed_col_title)
        op.addWidget(btn_set_title)
        op.addSpacing(8)
        op.addWidget(btn_add)
        op.addWidget(btn_del)
        op.addSpacing(8)
        op.addWidget(QLabel("幅(px)"))
        op.addWidget(self.spin_col_width)
        op.addWidget(btn_set_width)
        op.addStretch()
        hl.addLayout(op)

        self.list_cols.currentRowChanged.connect(self._on_col_selected)
        btn_set_title.clicked.connect(self._rename_selected_col)
        btn_set_width.clicked.connect(self._resize_selected_col)
        btn_add.clicked.connect(self._add_column)
        btn_del.clicked.connect(self._remove_selected_col)

        tabs.addTab(w_cols, "列管理")

        # 自動保存
        w_auto = QWidget()
        f_auto = QFormLayout(w_auto)
        self.chk_autosave = QPushButton("自動保存を有効にする (ON/OFF切替)")
        self.chk_autosave.setCheckable(True)
        self.chk_autosave.setChecked(bool(app.settings.get("autosave_enabled", True)))
        self.chk_autosave.toggled.connect(self._on_autosave_toggle)
        f_auto.addRow(self.chk_autosave)

        self.spin_interval = QSpinBox()
        self.spin_interval.setRange(3, 600)
        self.spin_interval.setValue(int(app.settings.get("autosave_interval_sec", 10)))
        self.spin_interval.valueChanged.connect(self._on_interval_change)
        f_auto.addRow("自動保存間隔 (秒)", self.spin_interval)

        self.ed_autopath = QLineEdit(str(app.autosave_path))
        btn_browse = QPushButton("参照…")
        btn_browse.clicked.connect(self._browse_autopath)
        row = QHBoxLayout()
        row.addWidget(self.ed_autopath); row.addWidget(btn_browse)
        f_auto.addRow("自動保存ファイル", row)

        tabs.addTab(w_auto, "自動保存")

        lay = QVBoxLayout(self)
        lay.addWidget(tabs)

        self._refresh_col_list()

    # ----- 表示設定/祝日/フォント/テーマ -----
    def _on_font_change(self, v: int):
        self.app._apply_font_all(v)
        self.app.settings["font_pt"] = int(v)
        self.app._save_settings()

    def _on_theme_toggle(self, on: bool):
        theme = "dark" if on else "light"
        self.app._apply_theme(theme)
        self.app.settings["theme"] = theme
        self.app._save_settings()

    def _on_holidays_change(self):
        txt = self.ed_holidays.text().strip()
        self.app.settings["holidays"] = txt
        self.app.holidays = parse_holidays_str(txt)
        self.app.rebuild_all()  # 色の更新を反映
        self.app._save_settings()

    # ----- 列管理 -----
    def _refresh_col_list(self):
        self.list_cols.clear()
        for c in self.app.columns:
            item = QListWidgetItem(f"{c['title']} ({c['width']}px)")
            item.setData(Qt.ItemDataRole.UserRole, c["id"])
            self.list_cols.addItem(item)
        if self.app.columns:
            self.list_cols.setCurrentRow(0)

    def _on_col_selected(self, row: int):
        if 0 <= row < len(self.app.columns):
            c = self.app.columns[row]
            self.ed_col_title.setText(c["title"])
            self.spin_col_width.setValue(int(c["width"]))

    def _rename_selected_col(self):
        r = self.list_cols.currentRow()
        if r < 0: return
        self.app.columns[r]["title"] = self.ed_col_title.text() or f"項目{r+1}"
        self._refresh_col_list()
        self.app.rebuild_all()
        self.app._mark_dirty()
        self.app.settings["columns"] = self.app.columns
        self.app._save_settings()

    def _resize_selected_col(self):
        r = self.list_cols.currentRow()
        if r < 0: return
        self.app.columns[r]["width"] = int(self.spin_col_width.value())
        self._refresh_col_list()
        self.app.rebuild_all()
        self.app._mark_dirty()
        self.app.settings["columns"] = self.app.columns
        self.app._save_settings()

    def _add_column(self):
        new_id = f"col-{uuid.uuid4().hex[:8]}"
        self.app.columns.append({"id": new_id, "title": f"項目{len(self.app.columns)+1}", "width": 260})
        self._refresh_col_list()
        self.app.rebuild_all()
        self.app._mark_dirty()
        self.app.settings["columns"] = self.app.columns
        self.app._save_settings()

    def _remove_selected_col(self):
        r = self.list_cols.currentRow()
        if r < 0: return
        col_id = self.app.columns[r]["id"]
        for k in list(self.app.cells.keys()):
            if col_id in self.app.cells[k]:
                del self.app.cells[k][col_id]
        del self.app.columns[r]
        self._refresh_col_list()
        self.app.rebuild_all()
        self.app._mark_dirty()
        self.app.settings["columns"] = self.app.columns
        self.app._save_settings()

    # ----- 自動保存 -----
    def _on_autosave_toggle(self, on: bool):
        self.app.settings["autosave_enabled"] = bool(on)
        self.app._apply_autosave_settings()
        self.app._save_settings()

    def _on_interval_change(self, v: int):
        self.app.settings["autosave_interval_sec"] = int(v)
        self.app._apply_autosave_settings()
        self.app._save_settings()

    def _browse_autopath(self):
        path, _ = QFileDialog.getSaveFileName(self, "自動保存ファイル", str(self.app.autosave_path), "JSON (*.json)")
        if not path: return
        self.ed_autopath.setText(path)
        self.app.autosave_path = Path(path)
        self.app.settings["autosave_path"] = str(self.app.autosave_path)
        self.app._save_settings()


# ---------- 検索ダイアログ（表示範囲/全データ） ----------
class SearchDialog(QDialog):
    """全セルから文字列を検索し、ダブルクリックで該当日にスクロール"""
    def __init__(self, app: "CalendarApp"):
        super().__init__(app)
        self.app = app
        self.setWindowTitle("検索")
        self.setMinimumSize(760, 520)

        root = QVBoxLayout(self)

        g1 = QGroupBox("文字列検索")
        f1 = QFormLayout(g1)
        self.ed_query = QLineEdit()
        f1.addRow("検索語", self.ed_query)

        lay_scope = QHBoxLayout()
        self.rb_range = QRadioButton("表示範囲"); self.rb_all = QRadioButton("すべて")
        self.rb_range.setChecked(True)
        lay_scope.addWidget(self.rb_range); lay_scope.addWidget(self.rb_all)
        f1.addRow("検索範囲", lay_scope)

        btn_find = QPushButton("検索")
        f1.addRow(btn_find)
        root.addWidget(g1)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["日付", "列", "内容"])
        self.tree.setColumnWidth(0, 160)
        root.addWidget(self.tree, 1)

        btn_find.clicked.connect(self._do_search)
        self.tree.itemActivated.connect(self._jump_and_close)

    def _do_search(self):
        q = (self.ed_query.text() or "").strip().lower()
        self.tree.clear()
        if not q:
            return

        def add_result(dt: date, col_title: str, txt: str):
            label = f"{dt.month}月{pad2(dt.day)}日({JP_WEEK[dt.weekday()]})"
            item = QTreeWidgetItem([label, col_title, txt.replace("\n", " ")])
            item.setData(0, Qt.ItemDataRole.UserRole, (dt.year, dt.month, dt.day))
            self.tree.addTopLevelItem(item)

        if self.rb_all.isChecked():
            for key, row in self.app.cells.items():
                try:
                    y, m, d = map(int, key.split("-"))
                    dt = date(y, m, d)
                except:
                    continue
                for col in self.app.columns:
                    txt = (row.get(col["id"], "") or "")
                    if q in txt.lower():
                        add_result(dt, col["title"], txt)
        else:
            # テーブルに展開済みの範囲のみ
            cur = self.app.range_start
            while cur <= self.app.range_end:
                row = self.app.cells.get(date_key(cur), {})
                for col in self.app.columns:
                    txt = (row.get(col["id"], "") or "")
                    if q in txt.lower():
                        add_result(cur, col["title"], txt)
                cur += timedelta(days=1)

    def _jump_and_close(self, item: QTreeWidgetItem):
        meta = item.data(0, Qt.ItemDataRole.UserRole)
        if not meta:
            return
        y, m, d = meta
        target = date(y, m, d)
        # 表示範囲外なら、その日を含むよう前後に展開してからスクロール
        self.app.ensure_date_visible(target)
        self.app.scroll_to_date(target)
        self.accept()


# ---------- 月選択ダイアログ ----------
class MonthPickerDialog(QDialog):
    """QCalendarWidgetで日付選択 → 選択月の月初へ表示をジャンプ/展開"""
    def __init__(self, app: "CalendarApp"):
        super().__init__(app)
        self.app = app
        self.setWindowTitle("月選択")
        self.setMinimumSize(420, 340)

        v = QVBoxLayout(self)
        self.cal = QCalendarWidget()
        # 現在の表示開始月を基準に
        rs = self.app.range_start
        self.cal.setSelectedDate(QDate(rs.year, rs.month, 1))
        v.addWidget(self.cal)

        h = QHBoxLayout()
        btn_ok = QPushButton("この月へ移動")
        btn_close = QPushButton("閉じる")
        h.addWidget(btn_ok); h.addWidget(btn_close)
        v.addLayout(h)

        btn_ok.clicked.connect(self._apply_and_close)
        btn_close.clicked.connect(self.reject)

    def _apply_and_close(self):
        qd = self.cal.selectedDate()
        month_start = date(qd.year(), qd.month(), 1)
        # 月初から約2ヶ月分へ置き換え（従来と同規模）
        self.app.rebuild_range(month_start, month_start + timedelta(days=61), keep_scroll=False)
        self.accept()


# ---------- メインアプリ ----------
class CalendarApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("カレンダー型メモ帳")

        today = date.today()
        # 初期表示：今月1日から約2ヶ月分（62日）
        month_start = date(today.year, today.month, 1)
        self.range_start: date = month_start
        self.range_end: date = month_start + timedelta(days=61)

        # 既定（設定で上書き）
        self.columns = [
            {"id": "col-1", "title": "予定", "width": 260},
            {"id": "col-2", "title": "メモ", "width": 260},
        ]
        self.cells: Dict[str, Dict[str, str]] = {}
        self.font_pt = 11
        self.holidays: Set[str] = set()

        # 設定ロード & 反映
        self.settings = self._load_settings()
        self._apply_settings_boot()
        self._apply_theme(self.settings.get("theme", "light"))

        # 起動時：前回の自動保存 JSON を読み込む（存在すれば）
        self.autosave_path = Path(self.settings.get("autosave_path", str(DEFAULT_AUTOSAVE)))
        self._load_last_autosave()

        # UI 構築
        root = QWidget()
        root_layout = QVBoxLayout(root)
        self.setCentralWidget(root)

        # メニューバー
        menubar = self.menuBar()
        filemenu = menubar.addMenu("ファイル")
        act_save = QAction("JSONで保存", self)
        act_load = QAction("JSONを読み込み", self)
        act_settings = QAction("設定...", self)
        act_save.triggered.connect(self.save_json)
        act_load.triggered.connect(self.load_json)
        act_settings.triggered.connect(self.open_settings)
        filemenu.addAction(act_save)
        filemenu.addAction(act_load)
        filemenu.addAction(act_settings)

        act_search = menubar.addAction("検索")
        act_search.triggered.connect(self.open_search_dialog)

        act_month = menubar.addAction("月選択")
        act_month.triggered.connect(self.open_month_dialog)

        # テーブル
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        root_layout.addWidget(self.table, 1)

        # スクロール監視（端で自動拡張）
        self.table.verticalScrollBar().valueChanged.connect(self._on_scroll_extend_if_needed)

        # 初回描画
        self.rebuild_range(self.range_start, self.range_end, keep_scroll=False)

        # 自動保存 初期化
        self._dirty = False
        self._autosave_timer = QTimer(self)
        self._autosave_timer.timeout.connect(self._autosave_tick)
        self._apply_autosave_settings()

    # ---------- 設定 I/O ----------
    def _load_settings(self) -> dict:
        if SETTINGS_PATH.exists():
            try:
                with SETTINGS_PATH.open("r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "font_pt": 11,
            "columns": self.columns,
            "holidays": "",
            "autosave_enabled": True,
            "autosave_interval_sec": 10,
            "autosave_path": str(DEFAULT_AUTOSAVE),
            "theme": "light",
        }

    def _save_settings(self):
        self.settings["font_pt"] = self.font_pt
        self.settings["columns"] = self.columns
        self.settings.setdefault("holidays", "")
        self.settings.setdefault("autosave_enabled", True)
        self.settings.setdefault("autosave_interval_sec", 10)
        self.settings.setdefault("autosave_path", str(self.autosave_path))
        self.settings.setdefault("theme", self.settings.get("theme", "light"))
        try:
            with SETTINGS_PATH.open("w", encoding="utf-8") as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _apply_settings_boot(self):
        self.font_pt = int(self.settings.get("font_pt", 11))
        self.columns = list(self.settings.get("columns", self.columns))
        self.holidays = parse_holidays_str(self.settings.get("holidays", ""))

    # ---------- テーマ ----------
    def _apply_theme(self, theme: str):
        app = QApplication.instance()
        if not app:
            return
        if theme == "dark":
            css = """
            QWidget { background-color: #1f2430; color: #e6e6e6; }
            QTableWidget, QTreeWidget, QTextEdit, QLineEdit { background-color: #2a2f3a; selection-background-color: #3b4252; }
            QHeaderView::section { background-color: #2f3441; color: #e6e6e6; }
            QMenuBar { background-color: #2a2f3a; }
            QMenuBar::item:selected { background: #3b4252; }
            QMenu { background-color: #2a2f3a; color: #e6e6e6; }
            QMenu::item:selected { background: #3b4252; }
            QPushButton { background-color: #364051; border: 1px solid #475063; padding: 4px 8px; }
            QPushButton:hover { background-color: #3f4a5e; }
            QGroupBox { border: 1px solid #3b4252; margin-top: 8px; }
            QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 3px; }
            QTabBar::tab { background: #2a2f3a; border: 1px solid #475063; padding: 6px 10px; }
            QTabBar::tab:selected { background: #364051; }
            QToolTip { background-color: #2a2f3a; color: #e6e6e6; border: 1px solid #475063; }
            """
        else:
            css = ""  # ライト
        app.setStyleSheet(css)

    # ---------- 起動時：前回自動保存の読み込み ----------
    def _load_json_path(self, path: Path, silent: bool = True) -> bool:
        try:
            with open(path, "r", encoding="utf-8") as f:
                obj = json.load(f)
            self.columns = obj.get("columns", self.columns)
            self.cells = obj.get("cells", self.cells)
            # year/month があれば、その月初から2ヶ月を初期表示に
            y = obj.get("year", None)
            m = obj.get("month", None)
            if isinstance(y, int) and isinstance(m, int) and 1 <= m <= 12:
                ms = date(y, m, 1)
                self.range_start = ms
                self.range_end = ms + timedelta(days=61)
            self._dirty = False
            return True
        except Exception as e:
            if not silent:
                QMessageBox.warning(self, "読み込み失敗", str(e))
            return False

    def _load_last_autosave(self):
        candidates = []
        sp = self.settings.get("autosave_path")
        if sp:
            candidates.append(Path(sp))
        candidates.append(DEFAULT_AUTOSAVE)
        for p in candidates:
            if p.exists():
                if self._load_json_path(p, silent=True):
                    self.autosave_path = p
                    break

    # ---------- テーブル構築（範囲全面再構築） ----------
    def rebuild_range(self, start: date, end: date, keep_scroll: bool):
        """start..end を全面再構築。keep_scroll=True なら元の先頭行を維持"""
        # アンカー（トップの可視行）を記憶
        anchor_dt = None
        if keep_scroll and self.table.rowCount() > 0:
            top_row = self.table.rowAt(0)
            if top_row < 0: top_row = 0
            anchor_dt = self.range_start + timedelta(days=top_row)

        self.range_start, self.range_end = start, end
        total_days = (end - start).days + 1
        self.table.clear()
        headers = ["日付"] + [c["title"] for c in self.columns]
        self.table.setColumnCount(1 + len(self.columns))
        self.table.setRowCount(total_days)
        self.table.setHorizontalHeaderLabels(headers)

        # 各日付行を生成
        cur = start
        for r in range(total_days):
            self._build_row(r, cur)
            cur += timedelta(days=1)

        self._apply_font_all(self.font_pt)

        # アンカーへスクロール
        if anchor_dt:
            self.scroll_to_date(anchor_dt)

    def rebuild_all(self):
        """列名/幅や祝日変更などの際に、現在範囲を描画し直す"""
        self.rebuild_range(self.range_start, self.range_end, keep_scroll=True)

    def _build_row(self, row: int, dt: date):
        key = date_key(dt)
        label = f"{dt.month}月{pad2(dt.day)}日({JP_WEEK[dt.weekday()]})"

        # --- 日付セル（左端のみ色付け） ---
        it = QTableWidgetItem(label)
        it.setFlags(Qt.ItemFlag.ItemIsEnabled)

        # 曜日・祝日の色分け
        is_holiday = key in self.holidays
        is_sun = (dt.weekday() == 6)
        is_sat = (dt.weekday() == 5)

        if is_holiday or is_sun:
            it.setBackground(QColor("#E53935"))   # 赤
            it.setForeground(QColor("#FFFFFF"))   # 白文字
        elif is_sat:
            it.setBackground(QColor("#E3F2FD"))   # 水色
            it.setForeground(QColor("#0D47A1"))   # 青文字
        # 平日は色なし

        self.table.setItem(row, 0, it)

        # --- 入力セル（色は塗らない） ---
        for ci, col in enumerate(self.columns, start=1):
            editor = AutoResizeTextEdit(base_font_pt=self.font_pt)
            if key in self.cells and col["id"] in self.cells[key]:
                editor.setPlainText(self.cells[key][col["id"]])

            editor.requestUrlList.connect(self._show_url_list_dialog)

            def make_slot(k=key, col_id=col["id"], ed=editor):
                def _slot():
                    self.cells.setdefault(k, {})[col_id] = ed.toPlainText()
                    self._mark_dirty()
                return _slot
            editor.textChanged.connect(make_slot())

            # 行高調整
            def adjust_row_height(r=row):
                max_h = 36
                for cc in range(1, self.table.columnCount()):
                    w = self.table.cellWidget(r, cc)
                    if isinstance(w, AutoResizeTextEdit):
                        max_h = max(max_h, w.height())
                if self.table.rowHeight(r) != max_h:
                    self.table.setRowHeight(r, max_h)

            editor.heightChanged.connect(lambda _h, r=row: adjust_row_height(r))
            QTimer.singleShot(0, lambda r=row: adjust_row_height(r))

            self.table.setCellWidget(row, ci, editor)

            key = date_key(dt)
            label = f"{dt.month}月{pad2(dt.day)}日({JP_WEEK[dt.weekday()]})"
            it = QTableWidgetItem(label)
            it.setFlags(Qt.ItemFlag.ItemIsEnabled)
            # 週末/祝日
            if dt.weekday() == 5:  # 土
                it.setBackground(QColor("#E8F4FF"))
            elif dt.weekday() == 6:  # 日
                it.setBackground(QColor("#FFE8EE"))
            if key in self.holidays:
                it.setBackground(QColor("#FFF6C2"))
            self.table.setItem(row, 0, it)

            for ci, col in enumerate(self.columns, start=1):
                editor = AutoResizeTextEdit(base_font_pt=self.font_pt)
                if key in self.cells and col["id"] in self.cells[key]:
                    editor.setPlainText(self.cells[key][col["id"]])

                editor.requestUrlList.connect(self._show_url_list_dialog)

            def make_slot(k=key, col_id=col["id"], ed=editor):
                def _slot():
                    self.cells.setdefault(k, {})[col_id] = ed.toPlainText()
                    self._mark_dirty()
                return _slot
            editor.textChanged.connect(make_slot())

            # 行の高さ … 行内最大に
            def adjust_row_height(row=row):
                max_h = 36
                for cc in range(1, self.table.columnCount()):
                    w = self.table.cellWidget(row, cc)
                    if isinstance(w, AutoResizeTextEdit):
                        max_h = max(max_h, w.height())
                if self.table.rowHeight(row) != max_h:
                    self.table.setRowHeight(row, max_h)

            editor.heightChanged.connect(lambda _h, r=row: adjust_row_height(r))
            QTimer.singleShot(0, lambda r=row: adjust_row_height(r))

            self.table.setCellWidget(row, ci, editor)

    def _apply_font_all(self, pt: int):
        self.font_pt = int(pt)
        for r in range(self.table.rowCount()):
            for c in range(1, self.table.columnCount()):
                w = self.table.cellWidget(r, c)
                if isinstance(w, AutoResizeTextEdit):
                    w.setPointSize(self.font_pt)

    # ---------- スクロール端で動的拡張 ----------
    def _on_scroll_extend_if_needed(self, _value: int):
        bar = self.table.verticalScrollBar()
        val = bar.value()
        mx = bar.maximum()
        # 下端に近い？ → 後ろへ +60日
        if mx - val < 120:  # 閾値（ピクセル）。必要に応じて調整可
            self._append_days(60)
        # 上端に近い？ → 前へ -60日
        if val < 120:
            self._prepend_days(60)

    def _append_days(self, n: int):
        # 現在の末尾から n 日 追加
        start_row = self.table.rowCount()
        self.table.setRowCount(start_row + n)
        dt = self.range_end + timedelta(days=1)
        for i in range(n):
            self._build_row(start_row + i, dt)
            dt += timedelta(days=1)
        self.range_end += timedelta(days=n)

    def _prepend_days(self, n: int):
        # 先頭に n 行挿入。表示位置がずれないようアンカー復元
        top_row_before = self.table.rowAt(0)
        if top_row_before < 0: top_row_before = 0
        anchor_dt = self.range_start + timedelta(days=top_row_before)

        for i in range(n):
            self.table.insertRow(0)
        # 新しく挿入した n 行（先頭側）を古い開始日の前から埋める
        dt = self.range_start - timedelta(days=n)
        for r in range(n):
            self._build_row(r, dt)
            dt += timedelta(days=1)
        self.range_start -= timedelta(days=n)

        # アンカー行（元の先頭行に相当）をトップに復元
        self.scroll_to_date(anchor_dt)

    # ---------- ユーティリティ ----------
    def ensure_date_visible(self, target: date):
        """target が表示範囲に入るまで前後に拡張"""
        # 必要に応じて60日単位で伸ばす
        while target < self.range_start:
            self._prepend_days(60)
        while target > self.range_end:
            self._append_days(60)

    def scroll_to_date(self, dt: date):
        if dt < self.range_start or dt > self.range_end:
            return
        idx = (dt - self.range_start).days
        it = self.table.item(idx, 0)
        if it:
            self.table.scrollToItem(it, QAbstractItemView.ScrollHint.PositionAtCenter)

    # ---------- ダイアログ起動 ----------
    def open_search_dialog(self):
        SearchDialog(self).exec()

    def open_month_dialog(self):
        MonthPickerDialog(self).exec()

    def _show_url_list_dialog(self, urls: List[str]):
        if not urls:
            QMessageBox.information(self, "URL", "URLが見つかりません。")
            return
        UrlListDialog(urls, self).exec()

    # ---------- 保存/読込 ----------
    def _write_json(self, path: Path):
        # 互換のため year/month は現在の range_start の月を保存
        obj = {
            "year": self.range_start.year,
            "month": self.range_start.month,
            "columns": self.columns,
            "cells": self.cells,
            "savedAt": datetime.now().isoformat()
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)

    def save_json(self):
        path, _ = QFileDialog.getSaveFileName(self, "JSONで保存", "", "JSON (*.json)")
        if not path:
            return
        self._write_json(Path(path))
        self.autosave_path = Path(path)
        self.settings["autosave_path"] = str(self.autosave_path)
        self._save_settings()
        self._dirty = False

    def load_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "JSONを読み込み", "", "JSON (*.json)")
        if not path:
            return
        if self._load_json_path(Path(path), silent=False):
            # 読み込んだ year/month に基づく2ヶ月へ置換
            self.rebuild_range(self.range_start, self.range_start + timedelta(days=61), keep_scroll=False)
            self.autosave_path = Path(path)
            self.settings["autosave_path"] = str(self.autosave_path)
            self._save_settings()
            self._dirty = False

    def open_settings(self):
        SettingsDialog(self).exec()

    # ---------- 自動保存 ----------
    def _apply_autosave_settings(self):
        enabled = bool(self.settings.get("autosave_enabled", True))
        interval_sec = int(self.settings.get("autosave_interval_sec", 10))
        self.autosave_path = Path(self.settings.get("autosave_path", str(self.autosave_path)))
        if enabled:
            if interval_sec < 3:
                interval_sec = 3
            if getattr(self, "_autosave_timer", None) is not None:
                self._autosave_timer.setInterval(interval_sec * 1000)
                if not self._autosave_timer.isActive():
                    self._autosave_timer.start()
        else:
            if getattr(self, "_autosave_timer", None) is not None:
                self._autosave_timer.stop()

    def _mark_dirty(self):
        self._dirty = True

    def _autosave_tick(self):
        if self._dirty and self.autosave_path:
            try:
                self._write_json(self.autosave_path)
                self._dirty = False
            except Exception:
                pass

    def closeEvent(self, ev: QCloseEvent):
        try:
            self._save_settings()
            if getattr(self, "_dirty", False) and self.autosave_path:
                self._write_json(self.autosave_path)
        except Exception:
            pass
        super().closeEvent(ev)


def main():
    app = QApplication(sys.argv)
    w = CalendarApp()
    w.resize(1200, 800)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
