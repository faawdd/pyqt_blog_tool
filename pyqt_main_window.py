import sys
from pathlib import Path
import re
from datetime import datetime

from PyQt6.QtCore import QEvent, QSize, Qt, QTimer
from PyQt6.QtGui import QAction, QFont, QIcon, QResizeEvent, QTextCharFormat, QTextCursor
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFontComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QSpinBox,
    QStatusBar,
    QStyle,
    QTextEdit,
    QToolButton,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from docx_converter import convert_docx_to_html_and_markdown, html_to_markdown
from git_publish import GitPushThread, RepoSyncThread, build_front_matter
from git import Repo


LOCAL_REPO_PATH = r"c:\Users\14434\my-blog"
REMOTE_REPO_URL = "https://github.com/faawdd/my-blog.git"
POSTS_RELATIVE_DIR = Path("content/posts")
GIT_BRANCH = "main"
GIT_USER_NAME = "faawdd"
GIT_USER_EMAIL = "1443469207@qq.com"


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("本地 Hugo 博客写作工具")
        self.resize(1200, 760)
        self.last_docx_path: str | None = None
        self.git_thread: GitPushThread | None = None
        self.repo_sync_thread: RepoSyncThread | None = None
        self._sync_guard = False
        self._icon_pixel_size = self._detect_icon_pixel_size()
        self._hover_icon_map: dict[object, tuple[QIcon, QIcon]] = {}
        self.current_article_path: Path | None = None
        self.current_front_matter = ""
        self.article_records: list[Path] = []
        self._metadata_updating = False
        self._compact_mode = False
        self._compact_actions: list[QAction] = []
        self._compact_widgets: list[QWidget] = []
        self._compact_measure_in_progress = False

        self._ensure_git_identity()
        self._setup_editors()
        self._setup_toolbar()
        self._setup_statusbar()
        self._apply_styles()

    def _ensure_git_identity(self) -> None:
        repo_git_dir = Path(LOCAL_REPO_PATH) / ".git"
        if not repo_git_dir.exists():
            return
        try:
            repo = Repo(LOCAL_REPO_PATH)
            with repo.config_writer() as config_writer:
                config_writer.set_value("user", "name", GIT_USER_NAME)
                config_writer.set_value("user", "email", GIT_USER_EMAIL)
        except Exception:
            # Keep UI available even if local git config write fails.
            pass

    def _setup_editors(self) -> None:
        left_panel = QWidget(self)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 8, 8, 8)
        left_layout.setSpacing(8)

        left_header_layout = QHBoxLayout()
        self.sync_repo_button = QPushButton("同步仓库", left_panel)
        self.sync_repo_button.clicked.connect(self.on_sync_repo)
        left_header_layout.addWidget(self.sync_repo_button)

        self.new_post_button = QPushButton("新建文章", left_panel)
        self.new_post_button.clicked.connect(self.on_new_article)
        left_header_layout.addWidget(self.new_post_button)
        left_layout.addLayout(left_header_layout)

        left_manage_layout = QHBoxLayout()
        self.rename_post_button = QPushButton("重命名", left_panel)
        self.rename_post_button.clicked.connect(self.on_rename_article)
        left_manage_layout.addWidget(self.rename_post_button)

        self.delete_post_button = QPushButton("删除", left_panel)
        self.delete_post_button.clicked.connect(self.on_delete_article)
        left_manage_layout.addWidget(self.delete_post_button)
        left_layout.addLayout(left_manage_layout)

        self.search_input = QLineEdit(left_panel)
        self.search_input.setPlaceholderText("搜索文章文件名...")
        self.search_input.textChanged.connect(self.on_filter_or_sort_changed)
        left_layout.addWidget(self.search_input)

        self.sort_combo = QComboBox(left_panel)
        self.sort_combo.addItems([
            "按修改时间（新到旧）",
            "按修改时间（旧到新）",
            "按文件名（A-Z）",
        ])
        self.sort_combo.currentIndexChanged.connect(self.on_filter_or_sort_changed)
        left_layout.addWidget(self.sort_combo)

        self.article_title_label = QLabel("当前文章: 未选择", left_panel)
        left_layout.addWidget(self.article_title_label)

        self.article_list = QListWidget(left_panel)
        self.article_list.itemClicked.connect(self.on_article_item_clicked)
        left_layout.addWidget(self.article_list)

        meta_group = QGroupBox("Front Matter", left_panel)
        meta_form = QFormLayout(meta_group)

        self.fm_title_input = QLineEdit(meta_group)
        self.fm_tags_input = QLineEdit(meta_group)
        self.fm_categories_input = QLineEdit(meta_group)
        self.fm_draft_checkbox = QCheckBox("draft", meta_group)
        self.fm_apply_button = QPushButton("应用元数据", meta_group)
        self.fm_apply_button.clicked.connect(self.on_apply_front_matter)

        self.fm_title_input.textChanged.connect(self.on_front_matter_field_changed)
        self.fm_tags_input.textChanged.connect(self.on_front_matter_field_changed)
        self.fm_categories_input.textChanged.connect(self.on_front_matter_field_changed)
        self.fm_draft_checkbox.toggled.connect(self.on_front_matter_field_changed)

        meta_form.addRow("title", self.fm_title_input)
        meta_form.addRow("tags(逗号分隔)", self.fm_tags_input)
        meta_form.addRow("categories(逗号分隔)", self.fm_categories_input)
        meta_form.addRow("", self.fm_draft_checkbox)
        meta_form.addRow("", self.fm_apply_button)
        left_layout.addWidget(meta_group)

        self.rich_editor = QTextEdit(self)
        self.rich_editor.setPlaceholderText("在这里编写富文本内容，或导入 Word 内容...")
        self.rich_editor.textChanged.connect(self.on_rich_text_changed)

        self.md_preview = QTextEdit(self)
        self.md_preview.setPlaceholderText("这里实时预览 Markdown 源码...")
        self.md_preview.setReadOnly(True)

        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.addWidget(left_panel)
        splitter.addWidget(self.rich_editor)
        splitter.addWidget(self.md_preview)
        splitter.setSizes([280, 560, 560])

        self.setCentralWidget(splitter)
        self.refresh_article_list()

    def _setup_toolbar(self) -> None:
        toolbar = QToolBar("编辑工具栏", self)
        self.toolbar = toolbar
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setAllowedAreas(Qt.ToolBarArea.TopToolBarArea)
        toolbar.setIconSize(QSize(self._icon_pixel_size, self._icon_pixel_size))
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, toolbar)

        def add_icon_action(
            text: str,
            tooltip: str,
            local_icon: str,
            fallback_icon: QStyle.StandardPixmap,
            handler,
        ) -> QAction:
            normal_icon = self._load_toolbar_icon(local_icon, "normal", fallback_icon)
            hover_icon = self._load_toolbar_icon(local_icon, "hover", fallback_icon)
            action = QAction(normal_icon, text, self)
            action.setToolTip(tooltip)
            action.setStatusTip(tooltip)
            action.triggered.connect(handler)
            toolbar.addAction(action)

            widget = toolbar.widgetForAction(action)
            if widget is not None:
                widget.setToolTip(tooltip)
                widget.installEventFilter(self)
                self._hover_icon_map[widget] = (normal_icon, hover_icon)
            return action

        self.bold_action = add_icon_action(
            "加粗",
            "加粗",
            "bold.svg",
            QStyle.StandardPixmap.SP_DialogApplyButton,
            self.on_bold,
        )
        self.italic_action = add_icon_action(
            "倾斜",
            "倾斜",
            "italic.svg",
            QStyle.StandardPixmap.SP_DialogResetButton,
            self.on_italic,
        )
        self.image_action = add_icon_action(
            "插入图片",
            "插入图片",
            "image.svg",
            QStyle.StandardPixmap.SP_FileIcon,
            self.on_insert_image,
        )
        self.front_matter_action = add_icon_action(
            "Front Matter",
            "插入 Front Matter 模板",
            "frontmatter.svg",
            QStyle.StandardPixmap.SP_MessageBoxInformation,
            self.on_insert_front_matter,
        )
        self.footnote_action = add_icon_action(
            "注释角标",
            "插入注释角标",
            "footnote.svg",
            QStyle.StandardPixmap.SP_MessageBoxInformation,
            self.on_insert_footnote,
        )
        self.h1_action = add_icon_action(
            "一级标题",
            "插入一级标题",
            "h1.svg",
            QStyle.StandardPixmap.SP_ArrowUp,
            self.on_heading_1,
        )
        self.h2_action = add_icon_action(
            "二级标题",
            "插入二级标题",
            "h2.svg",
            QStyle.StandardPixmap.SP_ArrowUp,
            self.on_heading_2,
        )
        self.h3_action = add_icon_action(
            "三级标题",
            "插入三级标题",
            "h3.svg",
            QStyle.StandardPixmap.SP_ArrowUp,
            self.on_heading_3,
        )
        self.quote_action = add_icon_action(
            "引用",
            "插入引用格式",
            "quote.svg",
            QStyle.StandardPixmap.SP_MessageBoxInformation,
            self.on_insert_quote,
        )
        self.table_action = add_icon_action(
            "插入表格",
            "插入 GFM 表格",
            "table.svg",
            QStyle.StandardPixmap.SP_DirIcon,
            self.on_insert_table,
        )
        self.task_action = add_icon_action(
            "任务列表",
            "插入任务列表",
            "task.svg",
            QStyle.StandardPixmap.SP_DialogApplyButton,
            self.on_insert_task_item,
        )
        self.formula_action = add_icon_action(
            "公式块",
            "插入 LaTeX 公式块",
            "formula.svg",
            QStyle.StandardPixmap.SP_ComputerIcon,
            self.on_insert_formula_block,
        )
        self.mermaid_action = add_icon_action(
            "Mermaid",
            "插入 Mermaid 图表块",
            "mermaid.svg",
            QStyle.StandardPixmap.SP_FileDialogDetailedView,
            self.on_insert_mermaid_block,
        )

        code_menu_button = QToolButton(self)
        code_normal_icon = self._load_toolbar_icon(
            "code.svg",
            "normal",
            QStyle.StandardPixmap.SP_FileDialogContentsView,
        )
        code_hover_icon = self._load_toolbar_icon(
            "code.svg",
            "hover",
            QStyle.StandardPixmap.SP_FileDialogContentsView,
        )
        code_menu_button.setIcon(code_normal_icon)
        code_menu_button.setToolTip("插入代码块（选择语言）")
        code_menu_button.setStatusTip("插入代码块（选择语言）")
        code_menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        code_menu_button.installEventFilter(self)
        self._hover_icon_map[code_menu_button] = (code_normal_icon, code_hover_icon)
        code_menu = QMenu(code_menu_button)
        for language in ["python", "html", "css", "javascript", "text"]:
            action = code_menu.addAction(language)
            action.triggered.connect(
                lambda _checked=False, lang=language: self.on_insert_code_block(lang)
            )
        code_menu_button.setMenu(code_menu)
        toolbar.addWidget(code_menu_button)
        self.code_menu_button = code_menu_button

        toolbar.addSeparator()

        font_controls_widget = QWidget(self)
        font_controls_layout = QHBoxLayout(font_controls_widget)
        font_controls_layout.setContentsMargins(0, 0, 0, 0)
        font_controls_layout.setSpacing(4)

        self.font_family_combo = QFontComboBox(font_controls_widget)
        self.font_family_combo.setMinimumWidth(150)
        self.font_family_combo.setMaximumWidth(190)
        self.font_family_combo.setToolTip("选择字体")
        self.font_family_combo.setCurrentFont(self.rich_editor.font())
        self.font_family_combo.currentFontChanged.connect(self.on_font_family_changed)
        font_controls_layout.addWidget(self.font_family_combo)

        self.font_size_spin = QSpinBox(font_controls_widget)
        self.font_size_spin.setRange(8, 72)
        self.font_size_spin.setFixedWidth(70)
        self.font_size_spin.setValue(max(8, int(self.rich_editor.fontPointSize() or 12)))
        self.font_size_spin.setToolTip("选择字号")
        self.font_size_spin.valueChanged.connect(self.on_font_size_changed)
        font_controls_layout.addWidget(self.font_size_spin)

        toolbar.addWidget(font_controls_widget)
        self.font_controls_widget = font_controls_widget

        toolbar.addSeparator()

        self.import_docx_action = add_icon_action(
            "导入 Docx",
            "导入 Docx",
            "import-docx.svg",
            QStyle.StandardPixmap.SP_DialogOpenButton,
            self.on_import_docx,
        )
        self.publish_action = add_icon_action(
            "一键生成 MD 并发布",
            "一键生成 MD 并发布",
            "publish.svg",
            QStyle.StandardPixmap.SP_DialogSaveButton,
            self.on_publish,
        )

        self._compact_actions = [
            self.front_matter_action,
            self.footnote_action,
            self.h2_action,
            self.h3_action,
            self.quote_action,
            self.table_action,
            self.task_action,
            self.formula_action,
            self.mermaid_action,
        ]
        self._compact_widgets = [self.code_menu_button]

        self.more_menu_button = QToolButton(self)
        more_normal_icon = self._load_toolbar_icon(
            "more.svg",
            "normal",
            QStyle.StandardPixmap.SP_TitleBarMenuButton,
        )
        more_hover_icon = self._load_toolbar_icon(
            "more.svg",
            "hover",
            QStyle.StandardPixmap.SP_TitleBarMenuButton,
        )
        self.more_menu_button.setIcon(more_normal_icon)
        self.more_menu_button.setToolTip("更多格式")
        self.more_menu_button.setStatusTip("更多格式")
        self.more_menu_button.setFixedWidth(34)
        self.more_menu_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        self.more_menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.more_menu = QMenu(self.more_menu_button)
        self.more_menu_button.setMenu(self.more_menu)
        self.more_menu_button.installEventFilter(self)
        self._hover_icon_map[self.more_menu_button] = (more_normal_icon, more_hover_icon)
        self.more_menu_button.setVisible(False)
        toolbar.addWidget(self.more_menu_button)
        self._rebuild_compact_menu()
        QTimer.singleShot(0, self._apply_narrow_window_strategy)

    def _load_toolbar_icon(
        self,
        icon_filename: str,
        state: str,
        fallback_icon: QStyle.StandardPixmap,
    ) -> QIcon:
        icon_path = (
            Path(__file__).resolve().parent
            / "assets"
            / "icons"
            / "theme"
            / str(self._icon_pixel_size)
            / state
            / icon_filename
        )
        if icon_path.exists():
            return QIcon(str(icon_path))
        return self.style().standardIcon(fallback_icon)

    def _detect_icon_pixel_size(self) -> int:
        app = QApplication.instance()
        if app is None or app.primaryScreen() is None:
            return 20
        screen = app.primaryScreen()
        dpr = screen.devicePixelRatio()
        dpi = screen.logicalDotsPerInch()
        return 24 if dpr >= 1.5 or dpi >= 120 else 20

    def eventFilter(self, watched, event):
        icon_pair = self._hover_icon_map.get(watched)
        if icon_pair and isinstance(watched, QToolButton):
            normal_icon, hover_icon = icon_pair
            if event.type() == QEvent.Type.Enter:
                watched.setIcon(hover_icon)
            elif event.type() == QEvent.Type.Leave:
                watched.setIcon(normal_icon)
        return super().eventFilter(watched, event)

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._apply_narrow_window_strategy()

    def _rebuild_compact_menu(self) -> None:
        self.more_menu.clear()
        compact_specs = [
            ("Front Matter", self.on_insert_front_matter),
            ("注释角标", self.on_insert_footnote),
            ("二级标题", self.on_heading_2),
            ("三级标题", self.on_heading_3),
            ("引用", self.on_insert_quote),
            ("插入表格", self.on_insert_table),
            ("任务列表", self.on_insert_task_item),
            ("公式块", self.on_insert_formula_block),
            ("Mermaid", self.on_insert_mermaid_block),
        ]
        for text, handler in compact_specs:
            action = self.more_menu.addAction(text)
            action.triggered.connect(handler)

        code_submenu = self.more_menu.addMenu("代码块")
        for language in ["python", "html", "css", "javascript", "text"]:
            action = code_submenu.addAction(language)
            action.triggered.connect(
                lambda _checked=False, lang=language: self.on_insert_code_block(lang)
            )

    def _apply_narrow_window_strategy(self) -> None:
        if not hasattr(self, "toolbar") or self._compact_measure_in_progress:
            return

        should_compact = self._should_use_compact_toolbar()
        if should_compact == self._compact_mode:
            self.more_menu_button.setVisible(should_compact)
            return

        self._compact_mode = should_compact
        for action in self._compact_actions:
            action.setVisible(not should_compact)
        for widget in self._compact_widgets:
            widget.setVisible(not should_compact)
        self.more_menu_button.setVisible(should_compact)

    def _should_use_compact_toolbar(self) -> bool:
        if not hasattr(self, "toolbar"):
            return False

        self._compact_measure_in_progress = True
        action_visibility = {action: action.isVisible() for action in self._compact_actions}
        widget_visibility = {widget: widget.isVisible() for widget in self._compact_widgets}
        more_visible = self.more_menu_button.isVisible()

        try:
            for action in self._compact_actions:
                action.setVisible(True)
            for widget in self._compact_widgets:
                widget.setVisible(True)
            self.more_menu_button.setVisible(False)

            self.toolbar.adjustSize()
            self.toolbar.updateGeometry()

            required_width = self.toolbar.sizeHint().width()
            available_width = max(self.toolbar.width(), self.width())
            return required_width > (available_width - 24)
        finally:
            for action, visible in action_visibility.items():
                action.setVisible(visible)
            for widget, visible in widget_visibility.items():
                widget.setVisible(visible)
            self.more_menu_button.setVisible(more_visible)
            self._compact_measure_in_progress = False

    def _setup_statusbar(self) -> None:
        status_bar = QStatusBar(self)
        self.setStatusBar(status_bar)
        self.statusBar().showMessage("准备就绪")

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background: #f3f4f6;
            }
            QToolBar {
                background: #ffffff;
                border-bottom: 1px solid #d1d5db;
                spacing: 6px;
                padding: 6px;
            }
            QToolBar QToolButton {
                background: #f9fafb;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 5px 10px;
            }
            QToolBar QToolButton:hover {
                background: #eef2ff;
            }
            QPushButton {
                background: #f8fafc;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 6px 10px;
            }
            QPushButton:hover {
                background: #e2e8f0;
            }
            QListWidget {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
            }
            QListWidget::item {
                padding: 6px;
            }
            QListWidget::item:selected {
                background: #dbeafe;
                color: #1f2937;
            }
            QTextEdit {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px;
                selection-background-color: #c7d2fe;
            }
            QStatusBar {
                background: #ffffff;
                border-top: 1px solid #d1d5db;
            }
            """
        )

    def _merge_format_on_selection(self, text_format: QTextCharFormat) -> None:
        cursor = self.rich_editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.mergeCharFormat(text_format)
        self.rich_editor.mergeCurrentCharFormat(text_format)

    def _insert_text_at_cursor(self, text: str) -> None:
        cursor = self.rich_editor.textCursor()
        cursor.insertText(text)
        self.rich_editor.setTextCursor(cursor)

    def _apply_editor_font(self, family: str | None = None, point_size: int | None = None) -> None:
        base_font = self.rich_editor.font()
        if family:
            base_font.setFamily(family)
        if point_size is not None:
            base_font.setPointSize(point_size)

        self.rich_editor.setFont(base_font)
        self.md_preview.setFont(base_font)

    def on_font_family_changed(self, font: QFont) -> None:
        self._apply_editor_font(family=font.family())
        self.statusBar().showMessage(f"字体已切换为: {font.family()}", 2000)

    def on_font_size_changed(self, size: int) -> None:
        self._apply_editor_font(point_size=size)
        self.statusBar().showMessage(f"字号已设置为: {size}", 2000)

    def _current_iso_datetime(self) -> str:
        # Use local timezone and keep seconds for Hugo front matter readability.
        return datetime.now().astimezone().isoformat(timespec="seconds")

    def on_bold(self) -> None:
        current_weight = self.rich_editor.fontWeight()
        target_weight = (
            QFont.Weight.Normal
            if current_weight == QFont.Weight.Bold
            else QFont.Weight.Bold
        )
        fmt = QTextCharFormat()
        fmt.setFontWeight(target_weight)
        self._merge_format_on_selection(fmt)
        self.statusBar().showMessage("已应用加粗样式", 2000)

    def on_italic(self) -> None:
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.rich_editor.fontItalic())
        self._merge_format_on_selection(fmt)
        self.statusBar().showMessage("已应用倾斜样式", 2000)

    def on_insert_image(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择图片",
            "",
            "Images (*.png *.jpg *.jpeg *.gif *.webp)",
        )
        if not file_path:
            self.statusBar().showMessage("已取消插入图片", 2000)
            return

        self.rich_editor.insertHtml(f'<img src="{file_path}" alt="image" />')
        self.statusBar().showMessage("已插入图片", 2000)

    def on_insert_front_matter(self) -> None:
        # Avoid duplicating front matter when users click the template button repeatedly.
        if self.current_front_matter.strip():
            QMessageBox.information(self, "提示", "文档顶部已存在 Front Matter。")
            return

        default_title = ""
        if self.current_article_path:
            default_title = self.current_article_path.stem

        self.current_front_matter = build_front_matter(
            title=default_title or "",
            tags=[],
            categories=[],
            draft=False,
            date_iso=self._current_iso_datetime(),
        )
        self._metadata_updating = True
        self.fm_title_input.setText(default_title)
        self.fm_tags_input.setText("")
        self.fm_categories_input.setText("")
        self.fm_draft_checkbox.setChecked(False)
        self._metadata_updating = False
        self.on_rich_text_changed()
        self.statusBar().showMessage("已插入 Front Matter 模板", 2500)

    def on_insert_footnote(self) -> None:
        note_text, ok = QInputDialog.getText(self, "注释内容", "请输入注释内容:")
        if not ok or not note_text.strip():
            return

        index = self._next_available_footnote_index()

        cursor = self.rich_editor.textCursor()
        cursor.insertText(f"[^{index}]")

        end_cursor = self.rich_editor.textCursor()
        end_cursor.movePosition(QTextCursor.MoveOperation.End)
        end_cursor.insertText(f"\n[^{index}]: {note_text}\n")

        self.statusBar().showMessage("已插入注释角标", 2000)

    def _insert_heading(self, level: int, fallback_text: str) -> None:
        cursor = self.rich_editor.textCursor()
        selected_text = cursor.selectedText().strip() if cursor.hasSelection() else ""
        heading_text = selected_text or fallback_text
        prefix = "#" * level
        cursor.insertText(f"{prefix} {heading_text}\n")

    def on_heading_1(self) -> None:
        self._insert_heading(1, "一级标题")
        self.statusBar().showMessage("已插入一级标题", 2000)

    def on_heading_2(self) -> None:
        self._insert_heading(2, "二级标题")
        self.statusBar().showMessage("已插入二级标题", 2000)

    def on_heading_3(self) -> None:
        self._insert_heading(3, "三级标题")
        self.statusBar().showMessage("已插入三级标题", 2000)

    def on_insert_quote(self) -> None:
        cursor = self.rich_editor.textCursor()
        selected_text = cursor.selectedText().strip() if cursor.hasSelection() else "引用内容"
        quote_lines = [line for line in selected_text.splitlines() if line.strip()]
        if not quote_lines:
            quote_lines = ["引用内容"]
        quote_block = "\n".join(f"> {line}" for line in quote_lines) + "\n"
        cursor.insertText(quote_block)
        self.statusBar().showMessage("已插入引用格式", 2000)

    def on_insert_table(self) -> None:
        rows, ok = QInputDialog.getInt(self, "插入表格", "行数:", value=3, min=1, max=50)
        if not ok:
            return
        cols, ok = QInputDialog.getInt(self, "插入表格", "列数:", value=3, min=1, max=20)
        if not ok:
            return

        # Emit pure GFM table syntax so Hugo/GitHub can render it consistently.
        headers = [f"标题{i + 1}" for i in range(cols)]
        header_line = "| " + " | ".join(headers) + " |"
        separator_line = "| " + " | ".join(["---"] * cols) + " |"
        body_lines = []
        for row_idx in range(rows):
            row_cells = [f"内容{row_idx + 1}-{col_idx + 1}" for col_idx in range(cols)]
            body_lines.append("| " + " | ".join(row_cells) + " |")
        table_md = "\n".join([header_line, separator_line, *body_lines]) + "\n"

        self._insert_text_at_cursor(table_md)
        self.statusBar().showMessage("已插入表格", 2000)

    def on_insert_task_item(self) -> None:
        self._insert_text_at_cursor("- [ ] ")
        self.statusBar().showMessage("已插入任务列表项", 2000)

    def on_insert_code_block(self, language: str) -> None:
        # Insert fenced code block and place caret inside for immediate typing.
        code_template = f"\n```{language}\n\n```\n"
        cursor = self.rich_editor.textCursor()
        cursor.insertText(code_template)
        cursor.movePosition(QTextCursor.MoveOperation.Up)
        self.rich_editor.setTextCursor(cursor)
        self.statusBar().showMessage(f"已插入 {language} 代码块", 2000)

    def on_insert_formula_block(self) -> None:
        self._insert_text_at_cursor("\n$$\n\n$$\n")
        self.statusBar().showMessage("已插入公式块", 2000)

    def on_insert_mermaid_block(self) -> None:
        mermaid_template = (
            "\n```mermaid\n"
            "graph TD\n"
            "    A[开始] --> B[结束]\n"
            "```\n"
        )
        self._insert_text_at_cursor(mermaid_template)
        self.statusBar().showMessage("已插入 Mermaid 图表块", 2000)

    def _next_available_footnote_index(self) -> int:
        text = self.rich_editor.toPlainText()
        used_numbers = {int(num) for num in re.findall(r"\[\^(\d+)\]", text)}
        candidate = 1
        while candidate in used_numbers:
            candidate += 1
        return candidate

    def on_sync_repo(self) -> None:
        if self.repo_sync_thread and self.repo_sync_thread.isRunning():
            return

        self.statusBar().showMessage("正在拉取仓库...")
        self.repo_sync_thread = RepoSyncThread(
            local_repo_path=LOCAL_REPO_PATH,
            remote_repo_url=REMOTE_REPO_URL,
            branch=GIT_BRANCH,
            git_user_name=GIT_USER_NAME,
            git_user_email=GIT_USER_EMAIL,
            parent=self,
        )
        self.repo_sync_thread.result_signal.connect(self.on_repo_sync_result)
        self.repo_sync_thread.start()

    def on_repo_sync_result(self, success: bool, message: str) -> None:
        if success:
            self.refresh_article_list()
            self.statusBar().showMessage("同步成功", 3000)
            QMessageBox.information(self, "同步完成", message)
        else:
            self.statusBar().showMessage("同步失败", 3000)
            QMessageBox.critical(self, "同步失败", message)
        QTimer.singleShot(3000, lambda: self.statusBar().showMessage("准备就绪"))

    def refresh_article_list(self) -> None:
        posts_dir = Path(LOCAL_REPO_PATH) / POSTS_RELATIVE_DIR
        posts_dir.mkdir(parents=True, exist_ok=True)

        selected_path = str(self.current_article_path) if self.current_article_path else ""
        self.article_records = sorted(posts_dir.glob("*.md"), key=lambda p: p.name.lower())

        search_text = self.search_input.text().strip().lower()
        filtered = [
            path for path in self.article_records if search_text in path.name.lower()
        ]

        sort_mode = self.sort_combo.currentText()
        if "新到旧" in sort_mode:
            filtered.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        elif "旧到新" in sort_mode:
            filtered.sort(key=lambda p: p.stat().st_mtime)
        else:
            filtered.sort(key=lambda p: p.name.lower())

        self.article_list.clear()
        for article_path in filtered:
            item = QListWidgetItem(article_path.name)
            item.setData(Qt.ItemDataRole.UserRole, str(article_path))
            item.setToolTip(
                f"最后修改: {datetime.fromtimestamp(article_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')}"
            )
            self.article_list.addItem(item)

        if selected_path:
            self.select_article_item(Path(selected_path))

    def on_filter_or_sort_changed(self) -> None:
        self.refresh_article_list()

    def on_article_item_clicked(self, item: QListWidgetItem) -> None:
        article_path_str = item.data(Qt.ItemDataRole.UserRole)
        if not article_path_str:
            return
        self.load_article(Path(article_path_str))

    def load_article(self, article_path: Path) -> None:
        try:
            raw_markdown = article_path.read_text(encoding="utf-8")
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", f"读取文章失败:\n{exc}")
            return

        front_matter, body_markdown = self._split_front_matter(raw_markdown)
        self.current_article_path = article_path
        self.current_front_matter = front_matter

        title = self._extract_title(front_matter) or article_path.stem
        self.article_title_label.setText(f"当前文章: {title}")
        self.setWindowTitle(f"本地 Hugo 博客写作工具 - {title}")

        self._metadata_updating = True
        self.fm_title_input.setText(title)
        self.fm_tags_input.setText(", ".join(self._extract_list_field(front_matter, "tags")))
        self.fm_categories_input.setText(
            ", ".join(self._extract_list_field(front_matter, "categories"))
        )
        self.fm_draft_checkbox.setChecked(self._extract_draft(front_matter))
        self._metadata_updating = False

        self.rich_editor.blockSignals(True)
        self.rich_editor.setMarkdown(body_markdown)
        self.rich_editor.blockSignals(False)
        self.on_rich_text_changed()
        self.statusBar().showMessage(f"已加载: {article_path.name}", 2500)

    def on_new_article(self) -> None:
        filename, ok = QInputDialog.getText(self, "新建文章", "请输入文章文件名:")
        if not ok or not filename.strip():
            return

        sanitized = re.sub(r'[^a-zA-Z0-9\-_.]+', "-", filename.strip()).strip(".-")
        if not sanitized:
            QMessageBox.warning(self, "无效名称", "文章文件名不能为空。")
            return

        posts_dir = Path(LOCAL_REPO_PATH) / POSTS_RELATIVE_DIR
        posts_dir.mkdir(parents=True, exist_ok=True)
        article_path = posts_dir / f"{sanitized}.md"
        if article_path.exists():
            QMessageBox.warning(self, "已存在", "同名文章已存在，请更换文件名。")
            return

        initial_front_matter = build_front_matter(
            sanitized.replace("-", " "),
            tags=[],
            categories=[],
            draft=False,
            date_iso=self._current_iso_datetime(),
        )
        article_path.write_text(initial_front_matter, encoding="utf-8")

        self.refresh_article_list()
        self.select_article_item(article_path)
        self.load_article(article_path)

    def on_rename_article(self) -> None:
        if not self.current_article_path:
            QMessageBox.warning(self, "未选择文章", "请先在列表中选择文章。")
            return

        new_name, ok = QInputDialog.getText(
            self,
            "重命名文章",
            "请输入新的文件名:",
            text=self.current_article_path.stem,
        )
        if not ok or not new_name.strip():
            return

        sanitized = re.sub(r'[^a-zA-Z0-9\-_.]+', "-", new_name.strip()).strip(".-")
        if not sanitized:
            QMessageBox.warning(self, "无效名称", "新的文件名无效。")
            return

        new_path = self.current_article_path.with_name(f"{sanitized}.md")
        if new_path.exists() and new_path != self.current_article_path:
            QMessageBox.warning(self, "已存在", "目标文件名已存在。")
            return

        try:
            self.current_article_path.rename(new_path)
        except Exception as exc:
            QMessageBox.critical(self, "重命名失败", f"重命名失败:\n{exc}")
            return

        self.current_article_path = new_path
        self.refresh_article_list()
        self.select_article_item(new_path)
        self.statusBar().showMessage("重命名完成", 2000)
        self.start_git_push("Rename post via PyQt6 Client")

    def on_delete_article(self) -> None:
        if not self.current_article_path:
            QMessageBox.warning(self, "未选择文章", "请先在列表中选择文章。")
            return

        article_name = self.current_article_path.name
        confirm = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除文章 {article_name} 吗？",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        try:
            self.current_article_path.unlink(missing_ok=False)
        except Exception as exc:
            QMessageBox.critical(self, "删除失败", f"删除失败:\n{exc}")
            return

        self.current_article_path = None
        self.current_front_matter = ""
        self.article_title_label.setText("当前文章: 未选择")
        self.rich_editor.clear()
        self.md_preview.clear()
        self.refresh_article_list()
        self.statusBar().showMessage("删除完成", 2000)
        self.start_git_push("Delete post via PyQt6 Client")

    def on_apply_front_matter(self) -> None:
        self.update_front_matter_from_fields()
        self.on_rich_text_changed()
        self.statusBar().showMessage("Front Matter 已更新", 2000)

    def on_front_matter_field_changed(self) -> None:
        if self._metadata_updating:
            return
        self.update_front_matter_from_fields()
        self.on_rich_text_changed()

    def update_front_matter_from_fields(self) -> None:
        title = self.fm_title_input.text().strip()
        if not title and self.current_article_path:
            title = self.current_article_path.stem

        tags = [x.strip() for x in self.fm_tags_input.text().split(",") if x.strip()]
        categories = [
            x.strip() for x in self.fm_categories_input.text().split(",") if x.strip()
        ]
        draft = self.fm_draft_checkbox.isChecked()
        date_iso = self._extract_date(self.current_front_matter) or self._current_iso_datetime()

        self.current_front_matter = build_front_matter(
            title=title,
            tags=tags,
            categories=categories,
            draft=draft,
            date_iso=date_iso,
        )
        self.article_title_label.setText(f"当前文章: {title or '未命名'}")

    def start_git_push(self, commit_message: str) -> None:
        if self.git_thread and self.git_thread.isRunning():
            QMessageBox.information(self, "提示", "当前已有 Git 推送任务正在执行。")
            return

        self.git_thread = GitPushThread(
            repo_path=LOCAL_REPO_PATH,
            branch=GIT_BRANCH,
            commit_message=commit_message,
            parent=self,
        )
        self.git_thread.started.connect(self.on_git_push_started)
        self.git_thread.result_signal.connect(self.on_git_push_result)
        self.git_thread.start()

    def select_article_item(self, article_path: Path) -> None:
        target = str(article_path)
        for index in range(self.article_list.count()):
            item = self.article_list.item(index)
            if item.data(Qt.ItemDataRole.UserRole) == target:
                self.article_list.setCurrentItem(item)
                break

    def _split_front_matter(self, text: str) -> tuple[str, str]:
        match = re.match(r"(?s)^---\n.*?\n---\n?", text)
        if not match:
            return "", text
        front = match.group(0).rstrip() + "\n\n"
        body = text[match.end():]
        return front, body

    def _extract_title(self, front_matter: str) -> str:
        title_match = re.search(r'(?m)^title:\s*"?(.*?)"?\s*$', front_matter)
        return title_match.group(1).strip() if title_match else ""

    def _extract_list_field(self, front_matter: str, field_name: str) -> list[str]:
        match = re.search(rf"(?m)^{field_name}:\s*\[(.*?)\]\s*$", front_matter)
        if not match:
            return []
        raw_items = [item.strip().strip('"\'') for item in match.group(1).split(",")]
        return [item for item in raw_items if item]

    def _extract_draft(self, front_matter: str) -> bool:
        match = re.search(r"(?m)^draft:\s*(true|false)\s*$", front_matter, re.IGNORECASE)
        return bool(match and match.group(1).lower() == "true")

    def _extract_date(self, front_matter: str) -> str:
        match = re.search(r"(?m)^date:\s*(.+?)\s*$", front_matter)
        return match.group(1).strip().strip('"') if match else ""

    def _editor_body_markdown(self) -> str:
        body_markdown = self.rich_editor.document().toMarkdown().strip()
        if not body_markdown:
            body_markdown = html_to_markdown(self.rich_editor.toHtml()).strip()
        return body_markdown + "\n" if body_markdown else ""

    def _compose_full_markdown(self) -> str:
        body_markdown = self._editor_body_markdown()
        front = self.current_front_matter.strip()
        if front:
            return front + "\n\n" + body_markdown
        return body_markdown

    def on_rich_text_changed(self) -> None:
        if self._sync_guard:
            return

        self._sync_guard = True
        try:
            markdown_text = self._compose_full_markdown()
            self.md_preview.setPlainText(markdown_text)
        finally:
            self._sync_guard = False

    def on_import_docx(self) -> None:
        docx_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Docx 文件",
            "",
            "Word Documents (*.docx)",
        )
        if not docx_path:
            self.statusBar().showMessage("已取消导入", 2000)
            return

        self.statusBar().showMessage("正在导入...", 1500)
        try:
            html_text, markdown_text = convert_docx_to_html_and_markdown(
                docx_path,
                "./static/images/",
            )
            self.last_docx_path = docx_path
            self.current_article_path = None
            self.current_front_matter = ""
            self.rich_editor.setHtml(html_text)
            self.md_preview.setPlainText(markdown_text)
            print(f"Docx 导入成功: {docx_path}")
            self.statusBar().showMessage("导入完成", 3000)
        except Exception as exc:
            print(f"Docx 导入失败: {exc}")
            QMessageBox.critical(self, "导入失败", f"导入 Docx 时出错:\n{exc}")
            self.statusBar().showMessage("导入失败", 3000)

    def on_publish(self) -> None:
        if self.git_thread and self.git_thread.isRunning():
            return

        if not self.current_article_path:
            QMessageBox.warning(self, "无法发布", "请先从左侧列表选择或新建文章。")
            return

        try:
            self.update_front_matter_from_fields()
            if not self.current_front_matter.strip():
                self.current_front_matter = build_front_matter(
                    title=self.current_article_path.stem,
                    tags=[],
                    categories=[],
                    draft=False,
                    date_iso=self._current_iso_datetime(),
                )
            full_markdown = self._compose_full_markdown().strip() + "\n"
            self.current_article_path.write_text(full_markdown, encoding="utf-8")
            print(f"文章已保存: {self.current_article_path}")
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", f"保存文章时出错:\n{exc}")
            self.statusBar().showMessage("保存失败", 3000)
            return

        self.start_git_push("Update post via PyQt6 Client")

    def on_git_push_started(self) -> None:
        self.statusBar().showMessage("正在推送至 GitHub...")

    def on_git_push_result(self, success: bool, message: str) -> None:
        if success:
            self.refresh_article_list()
            QMessageBox.information(self, "发布成功", message)
            self.statusBar().showMessage("同步成功", 3000)
        else:
            QMessageBox.critical(self, "发布失败", message)
            self.statusBar().showMessage("同步失败", 3000)
        QTimer.singleShot(3000, lambda: self.statusBar().showMessage("准备就绪"))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
