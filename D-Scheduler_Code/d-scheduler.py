import json,re,sys,uuid,webbrowser,math
from pathlib import Path
from datetime import date, timedelta
from typing import Dict, List, Set

from PyQt6.QtCore import Qt, QDate, QDateTime, pyqtSignal, QTimer
from PyQt6.QtGui import (
    QAction, QFont, QSyntaxHighlighter,
    QTextCharFormat, QColor, QTextCursor, QCloseEvent, QGuiApplication
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QTextEdit, QPushButton, QLabel,
    QSpinBox, QLineEdit, QFileDialog, QListWidget, QListWidgetItem,
    QMessageBox, QAbstractItemView, QTextBrowser, QDialog, QFormLayout,
    QGroupBox, QTreeWidget, QTreeWidgetItem, QCalendarWidget,
    QHBoxLayout, QRadioButton, QCheckBox
)

try:
    import jpholiday  # pip install jpholiday
    HAS_JPHOLIDAY = True
except Exception:
    HAS_JPHOLIDAY = False

# --------- 定数/ユーティリティ ----------
JP_WEEK = ["月", "火", "水", "木", "金", "土", "日"]
URL_RE = re.compile(r"https?://\S+", re.IGNORECASE)
SETTINGS_PATH = Path.home() / ".calendar_notes_settings.json"
DEFAULT_AUTOSAVE_DIR = Path.home() / "D-Schedule"
DEFAULT_AUTOSAVE = DEFAULT_AUTOSAVE_DIR / "calendar-notes_autosave.json"

def _strip_url_trailing_punct(u: str) -> str:
    """
    URL抽出後の末尾に付くことがある句読点や括弧類を除去する。
    例: "https://example.com)." → "https://example.com"
    """
    if not u:
        return u
    # 末尾の )]}.,;:!?、。 を連続で取り除く
    while u and u[-1] in ")]}.,;:!?、。":
        u = u[:-1]
    return u

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
        # URL検出後、末尾の句読点・括弧などを取り除き、実際にハイライトする長さを調整する。
        for m in URL_RE.finditer(text):
            start, end = m.span()
            frag = text[start:end]
            cleaned = _strip_url_trailing_punct(frag)
            self.setFormat(start, len(cleaned), self.format)


# ---------- 自動リサイズ付きテキストエディタ ----------
class AutoResizeTextEdit(QTextEdit):
    heightChanged = pyqtSignal(int)
    requestUrlList = pyqtSignal(list)

    def __init__(self, *args, base_font_pt=11, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptRichText(False)
        self.setFont(QFont("Meiryo UI", base_font_pt))
        self._padding_px = 6

        self.setContentsMargins(0, 0, 0, 0)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.document().setDocumentMargin(2)

        self.highlighter = UrlHighlighter(self.document())
        self.textChanged.connect(self._auto_resize)
        self._auto_resize()

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
        text = _strip_url_trailing_punct(cursor.selectedText() or "")
        if text.startswith("http://") or text.startswith("https://"):
            webbrowser.open(text)
        else:
            super().mouseDoubleClickEvent(ev)

    def setPointSize(self, pt: int):
        f = self.font()
        f.setPointSize(pt)
        self.setFont(f)
        self._auto_resize()

    def measureFullHeight(self) -> int:
        doc = self.document()
        vw = max(1, self.viewport().width() - 1)
        doc.setTextWidth(vw)
        size = doc.documentLayout().documentSize()

        h_doc = int(math.ceil(size.height()))
        fw = self.frameWidth()
        margin = int(doc.documentMargin())
        pad = self._padding_px
        return max(36, h_doc + 2 * fw + 2 * margin + pad) + 1

    def _auto_resize(self):
        # 1パス目
        h1 = self.measureFullHeight()
        if self.height() != h1:
            self.setFixedHeight(h1)
            self.heightChanged.emit(h1)
        # 2パス目（同期：再レイアウト直後にもう一度）
        h2 = self.measureFullHeight()
        if self.height() != h2:
            self.setFixedHeight(h2)
            self.heightChanged.emit(h2)

    def preview_markdown(self):
        dlg = QTextBrowser()
        dlg.setWindowTitle("Markdownプレビュー")
        dlg.setMarkdown(self.toPlainText() or "_(内容なし)_")
        dlg.setMinimumSize(480, 360)
        dlg.show()
        self._preview_holder = dlg

    def _extract_urls(self) -> list[str]:
        txt = self.toPlainText() or ""
        raw = URL_RE.findall(txt)
        return [_strip_url_trailing_punct(u) for u in raw]

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
    """
    縦列（カラム）の追加/削除・幅変更、テーマ切替、手動祝日の編集に対応した設定ダイアログ
    """
    applied = pyqtSignal(dict)

    def __init__(self, parent, settings: dict):
        super().__init__(parent)
        self.setWindowTitle("設定")
        self.setModal(True)
        self.settings = dict(settings)

        v = QVBoxLayout(self)

        # 列（カラム）設定
        box_cols = QGroupBox("列（カラム）設定")
        v_cols = QVBoxLayout(box_cols)

        # 列行を入れる専用コンテナ（この中に行を足す）
        self._col_rows = []  # (id, le_title, sp_width, row_layout)
        self._cols_container = QWidget()
        self._cols_v = QVBoxLayout(self._cols_container)
        self._cols_v.setContentsMargins(0, 0, 0, 0)
        self._cols_v.setSpacing(6)
        v_cols.addWidget(self._cols_container)

        # 既存列を追加（ボタンより上＝コンテナへ）
        for c in self.settings.get("columns", []):
            self._append_column_row(c.get("id", ""), c.get("title", ""), c.get("width", 240))

        # 「列を追加」ボタンはコンテナの“下”に置く
        hb_add = QHBoxLayout()
        btn_add = QPushButton("列を追加")
        btn_add.clicked.connect(lambda: self._append_column_row("", "", 240))
        hb_add.addStretch(1)
        hb_add.addWidget(btn_add)
        v_cols.addLayout(hb_add)

        v.addWidget(box_cols)

        # 表示 / テーマ
        box_view = QGroupBox("表示 / テーマ")
        v_view = QVBoxLayout(box_view)

        row_f = QHBoxLayout()
        row_f.addWidget(QLabel("本文フォント(pt):"))
        self.sp_font = QSpinBox()
        self.sp_font.setRange(8, 32)
        self.sp_font.setValue(int(self.settings.get("font_pt", 11)))
        row_f.addWidget(self.sp_font)
        v_view.addLayout(row_f)

        row_t = QHBoxLayout()
        row_t.addWidget(QLabel("テーマ:"))
        self.cb_theme_dark = QCheckBox("ダークモード")
        self.cb_theme_dark.setChecked(self.settings.get("theme", "light") == "dark")
        row_t.addWidget(self.cb_theme_dark)
        v_view.addLayout(row_t)

        row_a = QHBoxLayout()
        self.cb_today_top = QCheckBox("常に「今日」を最上段にする")
        self.cb_today_top.setChecked(bool(self.settings.get("always_today_top", True)))
        row_a.addWidget(self.cb_today_top)
        v_view.addLayout(row_a)

        v.addWidget(box_view)

        # その他
        box_misc = QGroupBox("その他")
        v_misc = QVBoxLayout(box_misc)

        row_s = QHBoxLayout()
        row_s.addWidget(QLabel("自動保存(秒):"))
        self.sp_autosave = QSpinBox()
        self.sp_autosave.setRange(5, 600)
        self.sp_autosave.setValue(int(self.settings.get("autosave_interval_sec", 10)))
        row_s.addWidget(self.sp_autosave)
        v_misc.addLayout(row_s)

        row_e = QHBoxLayout()
        row_e.addWidget(QLabel("自動拡張日数（片側）:"))
        self.sp_expand = QSpinBox()
        self.sp_expand.setRange(7, 120)
        self.sp_expand.setValue(int(self.settings.get("expand_days_each", 30)))
        row_e.addWidget(self.sp_expand)
        v_misc.addLayout(row_e)

        v.addWidget(box_misc)

        # 手動祝日
        box_h = QGroupBox("手動祝日（1行1日で YYYY-MM-DD を入力）")
        v_h = QVBoxLayout(box_h)
        self.ed_holidays = QTextEdit()
        self.ed_holidays.setPlaceholderText("例:\n2025-01-01\n2025-02-11\n…")
        self.ed_holidays.setFixedHeight(120)

        mh_list = self.settings.get("manual_holidays", []) or []
        hol_str = self.settings.get("holidays", "")
        if isinstance(mh_list, list) and mh_list:
            self.ed_holidays.setPlainText("\n".join(mh_list))
        elif isinstance(hol_str, str) and hol_str.strip():
            self.ed_holidays.setPlainText(hol_str)
        else:
            self.ed_holidays.setPlainText("")

        v_h.addWidget(self.ed_holidays)
        v.addWidget(box_h)

        # ボタン
        hb = QHBoxLayout()
        btn_apply = QPushButton("適用")
        btn_close = QPushButton("閉じる")
        btn_apply.clicked.connect(self._on_apply)
        btn_close.clicked.connect(self.reject)
        hb.addStretch(1)
        hb.addWidget(btn_apply)
        hb.addWidget(btn_close)
        v.addLayout(hb)

        self.setMinimumWidth(520)

    def _append_column_row(self, cid: str, title: str, width: int):
        row = QHBoxLayout()
        le = QLineEdit(title or "")
        sp = QSpinBox()
        sp.setRange(80, 1200)
        sp.setValue(max(80, min(int(width or 240), 1200)))
        btn_del = QPushButton("削除")

        def do_del():
            # 内部リストから削除
            for i, (rid, rle, rsp, rlay) in enumerate(list(self._col_rows)):
                if rle is le and rsp is sp:
                    self._col_rows.pop(i)
                    break
            # レイアウトから取り外し
            while row.count():
                w = row.takeAt(0).widget()
                if w:
                    w.setParent(None)

        btn_del.clicked.connect(do_del)

        row.addWidget(QLabel("タイトル:"))
        row.addWidget(le, 1)
        row.addWidget(QLabel("幅:"))
        row.addWidget(sp)
        row.addWidget(btn_del)

        # ★重要：追加先は self._cols_v（ボタンより上のコンテナ）
        self._cols_v.addLayout(row)

        self._col_rows.append((cid or "", le, sp, row))

    def _on_apply(self):
        cols = []
        seen = set()
        for cid, le, sp, _row in self._col_rows:
            title = (le.text() or "").strip()
            width = int(sp.value())
            if not cid:
                cid = (title or "col") + "-" + uuid.uuid4().hex[:4]
            if cid in seen:
                cid = cid + "-" + uuid.uuid4().hex[:2]
            seen.add(cid)
            cols.append({"id": cid, "title": title or cid, "width": width})

        mh_text = self.ed_holidays.toPlainText()
        mh_list = []
        rx = re.compile(r"^\d{4}-\d{2}-\d{2}$")
        for line in (mh_text or "").splitlines():
            s = line.strip()
            if s and rx.match(s):
                mh_list.append(s)
        mh_list = sorted(set(mh_list))

        new_settings = dict(self.settings)
        new_settings["columns"] = cols
        new_settings["font_pt"] = int(self.sp_font.value())
        new_settings["theme"] = "dark" if self.cb_theme_dark.isChecked() else "light"
        new_settings["always_today_top"] = bool(self.cb_today_top.isChecked())

        # ← ここが今回の修正ポイント（スピンボックス値を反映）
        new_settings["autosave_enabled"] = bool(new_settings.get("autosave_enabled", True))
        new_settings["autosave_interval_sec"] = int(self.sp_autosave.value())

        new_settings["expand_days_each"] = int(self.sp_expand.value())
        new_settings["holidays"] = "\n".join(mh_list)  # ストレージは文字列で保持（既存仕様に合わせる）

        self.applied.emit(new_settings)
        self.accept()


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
        # 選んだ日をアンカーに（＝最上行）
        self.app.set_view_anchor(target)
        self.accept()


# ---------- 日付選択ダイアログ ----------
class MonthPickerDialog(QDialog):
    """QCalendarWidgetで日付選択 → 選んだ日 or 今日 を基準（最上行）にして表示"""
    def __init__(self, app: "CalendarApp"):
        super().__init__(app)
        self.app = app
        self.setWindowTitle("日付選択")
        self.setMinimumSize(420, 360)

        v = QVBoxLayout(self)
        self.cal = QCalendarWidget()

        # 直近のアンカー日を初期選択（前回選んだ日を再表示）
        anchor = getattr(self.app, "current_anchor_date", None) or self.app.range_start
        self.cal.setSelectedDate(QDate(anchor.year, anchor.month, anchor.day))
        v.addWidget(self.cal)

        # ボタン群：今日へ / この日に合わせる（最上行） / 閉じる
        h = QHBoxLayout()
        btn_today = QPushButton("今日へ")
        btn_ok = QPushButton("この日に合わせる（最上行）")
        btn_close = QPushButton("閉じる")
        h.addWidget(btn_today)
        h.addWidget(btn_ok)
        h.addWidget(btn_close)
        v.addLayout(h)

        btn_today.clicked.connect(self._jump_today)
        btn_ok.clicked.connect(self._apply_and_close)
        btn_close.clicked.connect(self.reject)

    def _jump_today(self):
        """今日を基準にして、今日が最上行に来るよう表示を切り替える"""
        t = date.today()
        # カレンダー選択も今日に更新（次回開いたときの見た目にも自然）
        self.cal.setSelectedDate(QDate(t.year, t.month, t.day))
        # 表示は今日をアンカー（＝最上段）
        self.app.set_view_anchor(t)
        # 念のため可視位置も明示的に「最上段」へ（中央寄せを避ける）
        self.app.scroll_to_date(t)
        self.accept()


    def _apply_and_close(self):
        """カレンダーで選択した日を基準（最上行）にして表示"""
        qd = self.cal.selectedDate()
        chosen = date(qd.year(), qd.month(), qd.day())
        self.app.set_view_anchor(chosen)
        self.accept()


# ---------- メインアプリ ----------
class CalendarApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("D-Scheduler")

        today = date.today()
        # ▼ 初期表示は今日を含む範囲（今日を基準に前後31日ずつ＝約2か月）
        self.range_start: date = today - timedelta(days=31)
        self.range_end: date = today + timedelta(days=31)

        # 今日をアンカー（日付基準）にする
        self.current_anchor_date: date = today

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
        self._ensure_autosave_dir()

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

        act_month = menubar.addAction("日付選択")
        act_month.triggered.connect(self.open_month_dialog)

        # テーブル
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        root_layout.addWidget(self.table, 1)

        # 幅監視
        self.table.horizontalHeader().sectionResized.connect(self._on_section_resized)

        # スクロール監視（端で自動拡張）
        self.table.verticalScrollBar().valueChanged.connect(self._on_scroll_extend_if_needed)

        # 初回描画
        self.set_view_anchor(date.today())

        # 自動保存 初期化
        self._dirty = False
        self._autosave_timer = QTimer(self)
        self._autosave_timer.timeout.connect(self._autosave_tick)
        self._apply_autosave_settings()

    # --- スクロール／表示位置関連 ---
    def _ensure_autosave_dir(self):
        """autosave_path の親ディレクトリを必ず作成する。"""
        try:
            if getattr(self, "autosave_path", None):
                self.autosave_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass


    def set_view_anchor(self, anchor: date):
        """
        アンカー（基準日）を指定し、その日が「一番上」に来るように
        表示範囲を『anchor ～ anchor+61日』で再構築する。
        """
        self.current_anchor_date = anchor
        start = anchor
        end = anchor + timedelta(days=61)  # 約2ヶ月分
        # keep_scroll=False で先頭から描画＝アンカーが上端に来る
        self.rebuild_range(start, end, keep_scroll=False)

    def _get_top_visible_date(self) -> date:
        """現在のテーブルで最上段に見えている日付を返す。"""
        top_row = self.table.rowAt(0)
        if top_row < 0:
            top_row = 0
        return self.range_start + timedelta(days=top_row)

    def _begin_table_update(self):
        """テーブル全面更新の前に呼ぶ：スクロール信号を止め、再入を防止。"""
        self._ensure_flags()
        self._is_building = True
        try:
            bar = self.table.verticalScrollBar()
            self._sb_prev_block = bar.blockSignals(True)
        except Exception:
            self._sb_prev_block = False

    def _end_table_update(self):
        """テーブル全面更新の後に呼ぶ：スクロール信号を元に戻す。"""
        try:
            bar = self.table.verticalScrollBar()
            bar.blockSignals(self._sb_prev_block)
        except Exception:
            pass
        self._is_building = False

     # --- 列／サイズ関連 ---
    def _sync_row_height(self, row: int):
        """
        指定行の各セル(QTextEdit)の必要高さを計測し、
        1) 行全体の高さを最大に統一
        2) 各列のエディタの高さも最大に合わせる（＝文字枠も揃える）
        3) +1px の安全マージンで“隙間”発生を抑止
        """
        max_h = 36
        editors = []
        for ci in range(1, self.table.columnCount()):
            w = self.table.cellWidget(row, ci)
            if isinstance(w, AutoResizeTextEdit):
                h = w.measureFullHeight()
                max_h = max(max_h, h)
                editors.append(w)

        # 行全体を +1px で設定
        self.table.setRowHeight(row, max_h + 1)

        # 各エディタも最大に合わせる
        for ed in editors:
            ed.blockSignals(True)
            ed.setFixedHeight(max_h + 1)
            ed.blockSignals(False)
            
    def _on_section_resized(self, index: int, old_size: int, new_size: int):
        # 右端の項目列(=データ列)のみ保存（0列目は日付列）
        if index >= 1 and index - 1 < len(self.columns):
            self.columns[index - 1]["width"] = int(new_size)
            self._save_settings()
        # タイマーを発行せず、その場で可視行だけ再計算
        self._recalc_visible_rows()

    def _recalc_visible_rows(self):
        """現在のビューで見えている行だけ行高を同期（高速）。"""
        vp = self.table.viewport()
        top = self.table.rowAt(0)
        bot = self.table.rowAt(vp.height() - 1)

        if top < 0:
            top = 0
        if bot < 0:
            bot = self.table.rowCount() - 1

        for r in range(top, bot + 1):
            self._sync_row_height(r)
            
    def _recalc_all_row_heights(self):
        # 1周目
        for r in range(self.table.rowCount()):
            self._sync_row_height(r)
        # 2周目（レイアウト反映直後の微妙なズレを吸収）
        for r in range(self.table.rowCount()):
            self._sync_row_height(r)

           
    def get_current_column_widths(self) -> dict:
        """
        実テーブルの現在幅を読み出して {col_id: width} を返す。
        右側の項目列のみ（0列目=日付は対象外）。
        """
        header = self.table.horizontalHeader()
        widths = {}
        for idx, col in enumerate(self.columns, start=1):
            widths[col["id"]] = header.sectionSize(idx)
        return widths

    def apply_column_widths(self, widths: dict):
        """
        {col_id: width} を受け取り、再描画せずに列幅を反映。
        最上段の日付（表示位置）を保持し、可視範囲の行高を再計算する。
        """
        # いまの最上段日付を保持（巻き戻り防止）
        anchor = self._get_top_visible_date()

        header = self.table.horizontalHeader()
        prev_block = header.blockSignals(True)
        try:
            for idx, col in enumerate(self.columns, start=1):
                w = int(widths.get(col["id"], header.sectionSize(idx)))
                w = max(80, w)  # 最小幅の安全弁
                col["width"] = w
                self.table.setColumnWidth(idx, w)
        finally:
            header.blockSignals(prev_block)

        # 即時に可視範囲のみ行高再計算（singleShotを使わない）
        self._recalc_visible_rows()

        # 表示位置を維持
        self.scroll_to_date(anchor)

        # 設定へ保存
        self._save_settings()

    def _ensure_flags(self):
        """内部フラグの初期化"""
        if not hasattr(self, "_is_building"):
            self._is_building = False
        if not hasattr(self, "_is_extending"):
            self._is_extending = False
        if not hasattr(self, "_sb_prev_block"):
            self._sb_prev_block = False

    # ---------- 設定 I/O ----------
    def _load_settings(self) -> dict:
        if SETTINGS_PATH.exists():
            try:
                with SETTINGS_PATH.open("r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        cols_copy = [dict(c) for c in self.columns]
        return {
            "font_pt": 11,
            "columns": cols_copy,
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
        self.expand_each = int(self.settings.get("expand_days_each", 60))

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
        """
        JSONの columns / cells 構造を厳格にバリデーション。
        不正値はデフォルトにフォールバックし、壊れている場合は読み込みを中断する。
        """
        try:
            with open(path, "r", encoding="utf-8") as f:
                obj = json.load(f)

            # columns: [{id,title,width}]
            cols_raw = obj.get("columns", self.columns)
            cols = []
            if isinstance(cols_raw, list):
                seen = set()
                for c in cols_raw:
                    if not isinstance(c, dict):
                        continue
                    cid = str(c.get("id") or "").strip() or str(uuid.uuid4())
                    if cid in seen:
                        cid = cid + "-" + uuid.uuid4().hex[:4]
                    seen.add(cid)
                    title = str(c.get("title") or cid)
                    try:
                        width = int(c.get("width", 240))
                    except Exception:
                        width = 240
                    width = max(80, min(width, 1200))
                    cols.append({"id": cid, "title": title, "width": width})
            if not cols:
                raise ValueError("columns が不正です")

            # cells: {"YYYY-MM-DD": {col_id: str}}
            cells_raw = obj.get("cells", self.cells)
            cells = {}
            if isinstance(cells_raw, dict):
                for k, v in cells_raw.items():
                    if not (isinstance(k, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", k)):
                        continue
                    if not isinstance(v, dict):
                        continue
                    entry = {}
                    for cid, txt in v.items():
                        if not isinstance(cid, str):
                            continue
                        entry[cid] = "" if txt is None else str(txt)
                    cells[k] = entry

            if cells is None:
                cells = {}

            # 反映
            self.columns = cols
            self.cells = cells

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
                QMessageBox.warning(self, "読み込み失敗", f"JSONが壊れている可能性があります。\n{e}")
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
    def _is_holiday_jp(self, dt: date) -> bool:
        """
        日本の祝日を判定。jpholiday があればそちらを優先し、無い場合は
        設定で入力された self.holidays（"YYYY-MM-DD" の集合）を見る。
        """
        try:
            if HAS_JPHOLIDAY and jpholiday.is_holiday(dt):
                return True
        except Exception:
            pass
        # フォールバック（手入力祝日）
        return dt.strftime("%Y-%m-%d") in self.holidays

    
    def rebuild_range(self, start: date, end: date, keep_scroll: bool):
        """start..end を全面再構築。keep_scroll=True なら元の先頭行を維持"""
        # ★ 初回だけ、スクロールバーの actionTriggered を接続してユーザー操作を記録
        if not getattr(self, "_scroll_action_connected", False):
            try:
                self.table.verticalScrollBar().actionTriggered.connect(self._on_scroll_action)
                self._scroll_action_connected = True
            except Exception:
                self._scroll_action_connected = False

        anchor_dt = None
        if keep_scroll and self.table.rowCount() > 0:
            top_row = self.table.rowAt(0)
            if top_row < 0:
                top_row = 0
            anchor_dt = self.range_start + timedelta(days=top_row)

        self._begin_table_update()
        try:
            self.range_start, self.range_end = start, end
            total_days = (end - start).days + 1
            self.table.clear()
            headers = ["日付"] + [c["title"] for c in self.columns]
            self.table.setColumnCount(1 + len(self.columns))
            self.table.setRowCount(total_days)
            self.table.setHorizontalHeaderLabels(headers)

            cur = start
            for r in range(total_days):
                self._build_row(r, cur)
                cur += timedelta(days=1)

            self._apply_font_all(self.font_pt)
        finally:
            self._end_table_update()

        # 全行の高さを一括で同期
        self._recalc_all_row_heights()

        if anchor_dt:
            self.scroll_to_date(anchor_dt)

    def rebuild_all(self):
        """列名/幅や祝日変更などの際に、現在範囲を描画し直す"""
        self.rebuild_range(self.range_start, self.range_end, keep_scroll=True)

    def _build_row(self, row: int, dt: date):
        key = dt.strftime("%Y-%m-%d")
        w = dt.weekday()  # 月=0..日=6
        label = f"{dt.month}月{str(dt.day).zfill(2)}日({JP_WEEK[w]})"

        # --- 日付セル ---
        it = QTableWidgetItem(label)
        it.setFlags(Qt.ItemFlag.ItemIsEnabled)

        is_holiday = self._is_holiday_jp(dt)
        is_sun = (w == 6)
        is_sat = (w == 5)

        # 色付け（日祝ピンク・土曜水色）
        PINK_BG, PINK_FG = QColor("#FCE4EC"), QColor("#AD1457")
        BLUE_BG, BLUE_FG = QColor("#E3F2FD"), QColor("#0D47A1")

        if is_holiday or is_sun:
            it.setBackground(PINK_BG)
            it.setForeground(PINK_FG)
        elif is_sat:
            it.setBackground(BLUE_BG)
            it.setForeground(BLUE_FG)

        self.table.setItem(row, 0, it)

        # --- 入力セル ---
        editors = []
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

            self.table.setCellWidget(row, ci, editor)
            editors.append(editor)

        # 各エディタの高さ変化で都度再計算
        for ed in editors:
            ed.heightChanged.connect(lambda _h, r=row: self._sync_row_height(r))

        # ここで即時に一度だけ同期（singleShotを廃止）
        self._sync_row_height(row)

    
    def _apply_font_all(self, pt: int):
        self.font_pt = int(pt)
        for r in range(self.table.rowCount()):
            for c in range(1, self.table.columnCount()):
                w = self.table.cellWidget(r, c)
                if isinstance(w, AutoResizeTextEdit):
                    w.setPointSize(self.font_pt)

    # ---------- スクロール端で動的拡張 ----------
    def _on_scroll_action(self, action: int):
        """
        スクロールバーの actionTriggered シグナルを受けて、
        直近のユーザー操作があったことを記録する。
        """
        # ユーザー操作の有効期限（ミリ秒）
        self._user_scroll_expire_ms = int(QGuiApplication.primaryScreen().refreshRate() or 60) * 4
        self._last_user_scroll_ms = QDateTime.currentMSecsSinceEpoch()
    
    def _on_scroll_extend_if_needed(self, value: int):
        """
        スクロール端で日付範囲を自動拡張する。ただし
        ・直近ユーザー操作（actionTriggered）が一定時間内にあった場合のみ
        ・プログラムによるスクロール（scrollToItem/scrollToBottom 等）では拡張しない
        """
        bar = self.table.verticalScrollBar()

        # 直近ユーザー操作の有無（ホイール/キー/ドラッグに反応）
        now_ms = QDateTime.currentMSecsSinceEpoch()
        ok_user = False
        if hasattr(self, "_last_user_scroll_ms") and hasattr(self, "_user_scroll_expire_ms"):
            ok_user = (now_ms - self._last_user_scroll_ms) <= max(80, int(self._user_scroll_expire_ms))
        # 追加の安全弁：ドラッグ中のみ許可（ホイールも許可したい場合は or 条件を見直す）
        ok_drag = bar.isSliderDown()

        if not (ok_user or ok_drag):
            return

        # 端付近の閾値
        rng = bar.maximum() - bar.minimum()
        if rng <= 0:
            return
        near = max(8, rng // 20)

        # 下端に近い → 末尾へ拡張
        if bar.value() >= bar.maximum() - near:
            self._append_days_auto(self.expand_each if hasattr(self, "expand_each") else 60)
            return

        # 上端に近い → 先頭へ拡張
        if bar.value() <= bar.minimum() + near:
            self._prepend_days_auto(self.expand_each if hasattr(self, "expand_each") else 60)
            return
   
    def _append_days(self, n: int):
        """現在の末尾から n 日 追加"""
        if getattr(self, "_is_extending", False):
            return
        self._ensure_flags()
        self._is_extending = True
        try:
            # スクロール信号を止めて拡張（連鎖抑止）
            bar = self.table.verticalScrollBar()
            prev = bar.blockSignals(True)
            try:
                start_row = self.table.rowCount()
                self.table.setRowCount(start_row + n)
                dt = self.range_end + timedelta(days=1)
                for i in range(n):
                    self._build_row(start_row + i, dt)
                    dt += timedelta(days=1)
                self.range_end += timedelta(days=n)
            finally:
                bar.blockSignals(prev)
        finally:
            self._is_extending = False

    def _prepend_days(self, n: int):
        """先頭側へ n 日分挿入し、表示位置を保つ（最上段基準で復帰）"""
        if getattr(self, "_is_extending", False):
            return
        self._ensure_flags()
        self._is_extending = True
        try:
            bar = self.table.verticalScrollBar()
            prev = bar.blockSignals(True)
            try:
                # いまのトップ可視行に対応する日付を保存（後で戻す）
                top_row_before = self.table.rowAt(0)
                if top_row_before < 0:
                    top_row_before = 0
                anchor_dt = self.range_start + timedelta(days=top_row_before)

                for _ in range(n):
                    self.table.insertRow(0)

                dt = self.range_start - timedelta(days=n)
                for r in range(n):
                    self._build_row(r, dt)
                    dt += timedelta(days=1)
                self.range_start -= timedelta(days=n)

                # 元の位置に戻す（最上段へ）
                self.scroll_to_date(anchor_dt)
            finally:
                bar.blockSignals(prev)
        finally:
            self._is_extending = False

    def _append_days_auto(self, n: int):
        """
        末尾側に n 日自動拡張。拡張中はスクロールシグナルを一時停止して再帰発火を抑止。
        """
        bar = self.table.verticalScrollBar()
        prev = bar.blockSignals(True)
        try:
            end = self.range_end + timedelta(days=n)
            self.rebuild_range(self.range_start, end, keep_scroll=True)
        finally:
            bar.blockSignals(prev)
        self._recalc_visible_rows()

    def _prepend_days_auto(self, n: int):
        """
        先頭側に n 日自動拡張。復帰スクロールもシグナル停止して連鎖発火を抑止。
        """
        bar = self.table.verticalScrollBar()
        prev = bar.blockSignals(True)
        try:
            start = self.range_start - timedelta(days=n)
            # 先頭拡張時のアンカーは「拡張前の最上段」を維持
            self.rebuild_range(start, self.range_end, keep_scroll=True)
        finally:
            bar.blockSignals(prev)
        self._recalc_visible_rows()


    # ---------- ユーティリティ ----------
    def ensure_date_visible(self, target: date):
        """target が表示範囲に入るまで前後に拡張"""
        # 必要に応じて60日単位で伸ばす
        while target < self.range_start:
            self._prepend_days(60)
        while target > self.range_end:
            self._append_days(60)

    def scroll_to_date(self, dt: date):
        """
        指定日が表示範囲内にある前提で、その行を『最上段（PositionAtTop）』へスクロールする。
        ※従来は中央寄せだったため、拡張直後に視覚的な“位置ズレ”が発生しやすかった。
        """
        if dt < self.range_start or dt > self.range_end:
            return
        idx = (dt - self.range_start).days
        it = self.table.item(idx, 0)
        if it:
            self.table.scrollToItem(it, QAbstractItemView.ScrollHint.PositionAtTop)


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
        """
        安全なアトミック書き込みを行う。
        一時ファイルへの保存後、rename/replace で本体へ反映する。
        """
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            tmp = path.with_suffix(path.suffix + ".tmp")
            with tmp.open("w", encoding="utf-8") as f:
                json.dump({"columns": self.columns, "cells": self.cells}, f, ensure_ascii=False, indent=2)
            tmp.replace(path)
        except Exception as e:
            QMessageBox.warning(self, "保存失敗", str(e))

    def save_json(self):
        path, _ = QFileDialog.getSaveFileName(self, "スケジュールを保存", "", "JSON (*.json)")
        if not path:
            return
        self._write_json(Path(path))
        self.autosave_path = Path(path)
        self.settings["autosave_path"] = str(self.autosave_path)
        self._save_settings()
        self._dirty = False

    def load_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "スケジュールを読み込み", "", "JSON (*.json)")
        if not path:
            return
        if self._load_json_path(Path(path), silent=False):
            # 読み込んだ year/month に基づく2ヶ月へ置換
            self.rebuild_range(self.range_start, self.range_start + timedelta(days=61), keep_scroll=False)
            self.autosave_path = Path(path)
            self.settings["autosave_path"] = str(self.autosave_path)
            self._save_settings()
            self._dirty = False
            self.current_anchor_date = self.range_start

    def open_settings(self):
        dlg = SettingsDialog(self, self.settings)

        # 表示位置（最上段日付）を保持
        try:
            top_anchor = self._get_top_visible_date()
        except Exception:
            top_anchor = self.range_start

        def on_applied(s: dict):
            # 受け取った設定を反映
            self.settings = dict(s)

            # 列（増減/タイトル/幅）
            self.columns = list(self.settings.get("columns", self.columns))

            # フォント
            self.font_pt = int(self.settings.get("font_pt", self.font_pt))

            # テーマ
            theme = self.settings.get("theme", self.settings.get("theme", "light"))
            self._apply_theme(theme)

            # 手動祝日（文字列 → set へ）
            self.holidays = parse_holidays_str(self.settings.get("holidays", ""))

            # ★ 自動拡張日数の反映
            self.expand_each = int(self.settings.get("expand_days_each", getattr(self, "expand_each", 60)))

            # ヘッダー更新（列数・ラベル・幅）
            headers = ["日付"] + [c["title"] for c in self.columns]
            self.table.setColumnCount(1 + len(self.columns))
            self.table.setHorizontalHeaderLabels(headers)
            self.table.setColumnWidth(0, 120)
            for i, col in enumerate(self.columns, start=1):
                self.table.setColumnWidth(i, int(col.get("width", 240)))

            # 現在の範囲を再構築（祝日色・エディタ等を再作成）
            self.rebuild_all()

            # 全セルのフォントサイズを適用
            self._apply_font_all(self.font_pt)

            # 表示位置を復元（最上段に戻す）
            if isinstance(top_anchor, date):
                self.ensure_date_visible(top_anchor)
                self.scroll_to_date(top_anchor)

            # 自動保存設定を再適用（保存先なども反映）
            self._apply_autosave_settings()

            # 最後に設定を永続化
            self._save_settings()

        dlg.applied.connect(on_applied)
        dlg.exec()
D
    # ---------- 自動保存 ----------
    def _apply_autosave_settings(self):
        enabled = bool(self.settings.get("autosave_enabled", True))
        interval_sec = int(self.settings.get("autosave_interval_sec", 10))

        # 設定から保存先を反映
        apath = self.settings.get("autosave_path", str(getattr(self, "autosave_path", DEFAULT_AUTOSAVE)))
        self.autosave_path = Path(apath)

        # ▼フォルダ自動生成
        self._ensure_autosave_dir()

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
