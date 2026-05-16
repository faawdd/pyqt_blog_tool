import sys
from pathlib import Path
import re
import json
from datetime import datetime

from PyQt6.QtCore import QEvent, QSize, Qt, QTimer, QStandardPaths
from PyQt6.QtGui import (
    QAction,
    QColor,
    QFont,
    QIcon,
    QPalette,
    QPainter,
    QResizeEvent,
    QSyntaxHighlighter,
    QTextCharFormat,
    QTextCursor,
)
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
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
    QSizePolicy,
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


class MarkdownHighlighter(QSyntaxHighlighter):
    """Markdown 语法高亮器，应用于 Markdown 预览面板。"""

    # 多行状态常量
    _STATE_NORMAL = 0
    _STATE_FRONT_MATTER = 1
    _STATE_CODE_FENCE = 2
    _STATE_MATH_BLOCK = 3

    def __init__(self, document, theme: str = "light") -> None:
        super().__init__(document)
        self._single_rules: list[tuple[re.Pattern, QTextCharFormat]] = []
        self.theme = theme
        self._setup_rules()

    @staticmethod
    def _fmt(
        color: str,
        bold: bool = False,
        italic: bool = False,
        bg: str | None = None,
    ) -> QTextCharFormat:
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(color))
        if bold:
            fmt.setFontWeight(QFont.Weight.Bold)
        if italic:
            fmt.setFontItalic(True)
        if bg:
            fmt.setBackground(QColor(bg))
        return fmt

    def _setup_rules(self) -> None:
        f = self._fmt
        
        if self.theme == "dark":
            # 暗色主题颜色
            rules: list[tuple[re.Pattern, QTextCharFormat]] = [
                # 标题 # 到 ######
                (re.compile(r'^#{1,6}(?!#) .+'), f('#60a5fa', bold=True)),
                # 加粗 **text** / __text__
                (re.compile(r'\*\*[^*\n]+\*\*'), f('#c084fc', bold=True)),
                (re.compile(r'__[^_\n]+__'), f('#c084fc', bold=True)),
                # 斜体 *text* / _text_
                (re.compile(r'(?<!\*)\*(?!\*)[^*\n]+(?<!\*)\*(?!\*)'), f('#22d3ee', italic=True)),
                (re.compile(r'(?<!_)_(?!_)[^_\n]+(?<!_)_(?!_)'), f('#22d3ee', italic=True)),
                # 行内代码 `code`
                (re.compile(r'`[^`\n]+`'), f('#86efac', bg='#1e3a1f')),
                # 行内公式 $...$
                (re.compile(r'\$[^$\n]+\$'), f('#fbbf24', bg='#2d2413')),
                # 引用块 > ...
                (re.compile(r'^>.*'), f('#a0a0a0', italic=True)),
                # 图片 ![alt](url)
                (re.compile(r'!\[[^\]]*\]\([^)]+\)'), f('#2dd4bf')),
                # 链接 [text](url)
                (re.compile(r'\[[^\]]+\]\([^)]+\)'), f('#38bdf8')),
                # 脚注 [^1]
                (re.compile(r'\[\^[0-9a-zA-Z]+\]'), f('#f43f5e', bold=True, bg='#3f1621')),
                # 任务列表
                (re.compile(r'^[\s]*[-*]\s+\[[ xX]\]'), f('#f59e0b', bold=True)),
                # 无序列表标记
                (re.compile(r'^[\s]*[-*]\s'), f('#818cf8')),
                # 有序列表标记
                (re.compile(r'^[\s]*\d+\.\s'), f('#818cf8')),
                # 表格分隔符
                (re.compile(r'\|'), f('#64748b')),
                # 分隔线
                (re.compile(r'^-{3,}\s*$'), f('#78716c')),
            ]
        else:
            # 亮色主题颜色
            rules: list[tuple[re.Pattern, QTextCharFormat]] = [
                # 标题 # 到 ######
                (re.compile(r'^#{1,6}(?!#) .+'), f('#1d4ed8', bold=True)),
                # 加粗 **text** / __text__
                (re.compile(r'\*\*[^*\n]+\*\*'), f('#7c3aed', bold=True)),
                (re.compile(r'__[^_\n]+__'), f('#7c3aed', bold=True)),
                # 斜体 *text* / _text_
                (re.compile(r'(?<!\*)\*(?!\*)[^*\n]+(?<!\*)\*(?!\*)'), f('#0891b2', italic=True)),
                (re.compile(r'(?<!_)_(?!_)[^_\n]+(?<!_)_(?!_)'), f('#0891b2', italic=True)),
                # 行内代码 `code`
                (re.compile(r'`[^`\n]+`'), f('#059669', bg='#f0fdf4')),
                # 行内公式 $...$
                (re.compile(r'\$[^$\n]+\$'), f('#b45309', bg='#fffbeb')),
                # 引用块 > ...
                (re.compile(r'^>.*'), f('#78716c', italic=True)),
                # 图片 ![alt](url)
                (re.compile(r'!\[[^\]]*\]\([^)]+\)'), f('#0d9488')),
                # 链接 [text](url)
                (re.compile(r'\[[^\]]+\]\([^)]+\)'), f('#0369a1')),
                # 脚注 [^1]
                (re.compile(r'\[\^[0-9a-zA-Z]+\]'), f('#e11d48', bold=True, bg='#fdf2f8')),
                # 任务列表
                (re.compile(r'^[\s]*[-*]\s+\[[ xX]\]'), f('#d97706', bold=True)),
                # 无序列表标记
                (re.compile(r'^[\s]*[-*]\s'), f('#6366f1')),
                # 有序列表标记
                (re.compile(r'^[\s]*\d+\.\s'), f('#6366f1')),
                # 表格分隔符
                (re.compile(r'\|'), f('#64748b')),
                # 分隔线
                (re.compile(r'^-{3,}\s*$'), f('#94a3b8')),
            ]
        
        self._single_rules = rules

    def _get_theme_colors(self) -> dict[str, str]:
        """返回当前主题的配色方案。"""
        if self.theme == "dark":
            return {
                "front_matter": "#c084fc",
                "front_matter_key": "#22d3ee",
                "front_matter_text": "#a0a0a0",
                "code_fence": "#86efac",
                "code_fence_bg": "#1e3a1f",
                "math_block": "#fbbf24",
                "math_block_bg": "#2d2413",
            }
        else:
            return {
                "front_matter": "#c026d3",
                "front_matter_key": "#0d9488",
                "front_matter_text": "#6b7280",
                "code_fence": "#059669",
                "code_fence_bg": "#f0fdf4",
                "math_block": "#b45309",
                "math_block_bg": "#fffbeb",
            }

    def highlightBlock(self, text: str) -> None:  # type: ignore[override]
        prev = self.previousBlockState()
        block_num = self.currentBlock().blockNumber()
        fm_delim = re.compile(r'^---\s*$')
        fence_re = re.compile(r'^```')
        math_re = re.compile(r'^\$\$')
        
        colors = self._get_theme_colors()
        f = self._fmt

        # ─── Front Matter（文件首个 --- 开头） ───
        if block_num == 0 and fm_delim.match(text):
            self.setCurrentBlockState(self._STATE_FRONT_MATTER)
            self.setFormat(0, len(text), f(colors["front_matter"], bold=True))
            return

        if prev == self._STATE_FRONT_MATTER:
            if fm_delim.match(text):
                self.setCurrentBlockState(self._STATE_NORMAL)
                self.setFormat(0, len(text), f(colors["front_matter"], bold=True))
            else:
                self.setCurrentBlockState(self._STATE_FRONT_MATTER)
                self.setFormat(0, len(text), f(colors["front_matter_text"]))
                m = re.match(r'^([a-zA-Z_][\w]*)(\s*:)', text)
                if m:
                    self.setFormat(0, m.end(1), f(colors["front_matter_key"], bold=True))
            return

        # ─── 代码块 ``` ───
        if prev == self._STATE_CODE_FENCE:
            self.setFormat(0, len(text), f(colors["code_fence"], bg=colors["code_fence_bg"]))
            if fence_re.match(text):
                self.setCurrentBlockState(self._STATE_NORMAL)
            else:
                self.setCurrentBlockState(self._STATE_CODE_FENCE)
            return

        if fence_re.match(text):
            self.setFormat(0, len(text), f(colors["code_fence"], bg=colors["code_fence_bg"]))
            self.setCurrentBlockState(self._STATE_CODE_FENCE)
            return

        # ─── 公式块 $$ ───
        if prev == self._STATE_MATH_BLOCK:
            self.setFormat(0, len(text), f(colors["math_block"], bg=colors["math_block_bg"]))
            if math_re.match(text):
                self.setCurrentBlockState(self._STATE_NORMAL)
            else:
                self.setCurrentBlockState(self._STATE_MATH_BLOCK)
            return

        if math_re.match(text):
            self.setFormat(0, len(text), f(colors["math_block"], bg=colors["math_block_bg"]))
            self.setCurrentBlockState(self._STATE_MATH_BLOCK)
            return

        # ─── 普通行：应用单行规则 ───
        self.setCurrentBlockState(self._STATE_NORMAL)
        for pattern, fmt in self._single_rules:
            for m in pattern.finditer(text):
                self.setFormat(m.start(), m.end() - m.start(), fmt)


class ElidedLabel(QLabel):
    """自动用省略号截断过长文本的 QLabel。

    重写 sizeHint / minimumSizeHint，使水平方向不随文本内容撑大布局。
    """

    def sizeHint(self):
        hint = super().sizeHint()
        hint.setWidth(50)
        return hint

    def minimumSizeHint(self):
        hint = super().minimumSizeHint()
        hint.setWidth(50)
        return hint

    def paintEvent(self, event) -> None:  # type: ignore[override]
        painter = QPainter(self)
        metrics = self.fontMetrics()
        elided = metrics.elidedText(self.text(), Qt.TextElideMode.ElideRight, self.width())
        painter.drawText(self.rect(), self.alignment() | Qt.AlignmentFlag.AlignVCenter, elided)



# 应用信息
__appname__ = "墨筑 (MoZu)"
__version__ = "1.0"

POSTS_RELATIVE_DIR = Path("content/post")
POSTS_RELATIVE_DIR_CANDIDATES = [Path("content/post"), Path("content/posts")]
DEFAULT_BRANCH = "main"
LEGACY_SETTINGS_FILE = Path(__file__).resolve().parent / "app_settings.json"


def get_settings_file_path() -> Path:
    config_base = QStandardPaths.writableLocation(
        QStandardPaths.StandardLocation.AppConfigLocation
    )
    if not config_base:
        # Fallback when Qt cannot resolve platform config dir.
        return LEGACY_SETTINGS_FILE

    settings_dir = Path(config_base) / "MoZu"
    settings_dir.mkdir(parents=True, exist_ok=True)
    return settings_dir / "app_settings.json"


SETTINGS_FILE = get_settings_file_path()


def get_resource_base_path() -> Path:
    # PyInstaller onefile 解包后会将资源放在 _MEIPASS 目录。
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent


def detect_system_theme() -> str:
    """检测系统主题偏好，返回 'dark' 或 'light'。"""
    try:
        # 检查 QPalette 背景颜色亮度
        app = QApplication.instance()
        if app is None:
            return "light"
        
        palette = app.palette()
        bg_color = palette.color(QPalette.ColorRole.Window)
        brightness = bg_color.lightness()
        
        return "dark" if brightness < 128 else "light"
    except Exception:
        return "light"


class RepoCommitConfigDialog(QDialog):
    def __init__(
        self,
        local_repo_path: str,
        remote_repo_url: str,
        branch: str,
        git_user_name: str,
        git_user_email: str,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("仓库与提交配置")
        self.resize(560, 260)

        layout = QVBoxLayout(self)

        repo_group = QGroupBox("仓库配置", self)
        repo_form = QFormLayout(repo_group)
        self.repo_path_input = QLineEdit(local_repo_path, self)
        self.remote_url_input = QLineEdit(remote_repo_url, self)
        self.branch_input = QLineEdit(branch or DEFAULT_BRANCH, self)

        commit_group = QGroupBox("提交配置", self)
        commit_form = QFormLayout(commit_group)
        self.user_name_input = QLineEdit(git_user_name, self)
        self.user_email_input = QLineEdit(git_user_email, self)

        browse_button = QPushButton("浏览...", self)
        browse_button.clicked.connect(self.on_browse_repo_path)
        repo_path_row = QHBoxLayout()
        repo_path_row.addWidget(self.repo_path_input)
        repo_path_row.addWidget(browse_button)

        repo_form.addRow("本地仓库路径", self._wrap_layout(repo_path_row))
        repo_form.addRow("远端仓库 URL", self.remote_url_input)
        repo_form.addRow("分支", self.branch_input)

        commit_form.addRow("提交用户名", self.user_name_input)
        commit_form.addRow("提交邮箱", self.user_email_input)

        layout.addWidget(repo_group)
        layout.addWidget(commit_group)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _wrap_layout(self, row_layout: QHBoxLayout) -> QWidget:
        wrapper = QWidget(self)
        wrapper.setLayout(row_layout)
        return wrapper

    def on_browse_repo_path(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择本地仓库目录")
        if selected:
            self.repo_path_input.setText(selected)

    def values(self) -> dict[str, str]:
        return {
            "local_repo_path": self.repo_path_input.text().strip(),
            "remote_repo_url": self.remote_url_input.text().strip(),
            "branch": self.branch_input.text().strip() or DEFAULT_BRANCH,
            "git_user_name": self.user_name_input.text().strip(),
            "git_user_email": self.user_email_input.text().strip(),
        }


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{__appname__} 博客工具 v{__version__}")
        # 设置应用图标
        icon_path = str(get_resource_base_path() / "assets" / "icons" / "mozu.svg")
        self.setWindowIcon(QIcon(icon_path))
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

        self.local_repo_path = ""
        self.remote_repo_url = ""
        self.git_branch = DEFAULT_BRANCH
        self.git_user_name = ""
        self.git_user_email = ""

        self._load_settings()

        self._ensure_git_identity()
        self._setup_editors()
        self._setup_toolbar()
        self._setup_statusbar()
        self._apply_styles()

    def _load_settings(self) -> None:
        if not SETTINGS_FILE.exists() and LEGACY_SETTINGS_FILE.exists():
            try:
                SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
                SETTINGS_FILE.write_text(
                    LEGACY_SETTINGS_FILE.read_text(encoding="utf-8"),
                    encoding="utf-8",
                )
            except Exception:
                # Ignore migration failure and continue with current state.
                pass

        if not SETTINGS_FILE.exists():
            return
        try:
            raw = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return

        self.local_repo_path = str(raw.get("local_repo_path", "")).strip()
        self.remote_repo_url = str(raw.get("remote_repo_url", "")).strip()
        self.git_branch = str(raw.get("branch", DEFAULT_BRANCH)).strip() or DEFAULT_BRANCH
        self.git_user_name = str(raw.get("git_user_name", "")).strip()
        self.git_user_email = str(raw.get("git_user_email", "")).strip()

    def _save_settings(self) -> None:
        data = {
            "local_repo_path": self.local_repo_path,
            "remote_repo_url": self.remote_repo_url,
            "branch": self.git_branch,
            "git_user_name": self.git_user_name,
            "git_user_email": self.git_user_email,
        }
        try:
            SETTINGS_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as exc:
            QMessageBox.warning(self, "保存失败", f"配置保存失败:\n{exc}")

    def _configured_posts_dir(
        self,
        show_warning: bool = True,
        create_if_missing: bool = False,
    ) -> Path | None:
        repo_path = self.local_repo_path.strip()
        if not repo_path:
            if show_warning:
                QMessageBox.warning(
                    self,
                    "未配置仓库",
                    "请先点击“仓库设置”配置本地仓库路径。",
                )
            return None

        repo_base = Path(repo_path)
        for relative_dir in POSTS_RELATIVE_DIR_CANDIDATES:
            candidate = repo_base / relative_dir
            if candidate.exists() and candidate.is_dir():
                return candidate

        fallback = repo_base / POSTS_RELATIVE_DIR
        if create_if_missing:
            fallback.mkdir(parents=True, exist_ok=True)
        return fallback

    def _collect_article_files(self) -> list[Path]:
        repo_path = self.local_repo_path.strip()
        if not repo_path:
            return []

        posts_dir = self._configured_posts_dir(show_warning=False)
        if posts_dir and posts_dir.exists():
            files = [p for p in posts_dir.rglob("*.md") if p.name != "_index.md"]
            if files:
                return sorted(files, key=lambda p: p.name.lower())

        # 回退策略：当仓库结构不含 content/post(s) 时，扫描 content 下全部文章。
        content_dir = Path(repo_path) / "content"
        if not content_dir.exists():
            return []

        files = [p for p in content_dir.rglob("*.md") if p.name != "_index.md"]
        return sorted(files, key=lambda p: p.name.lower())

    def _ensure_git_identity(self) -> None:
        repo_path = self.local_repo_path.strip()
        if not repo_path:
            return

        repo_git_dir = Path(repo_path) / ".git"
        if not repo_git_dir.exists():
            return
        if not (self.git_user_name and self.git_user_email):
            return

        try:
            repo = Repo(repo_path)
            with repo.config_writer() as config_writer:
                config_writer.set_value("user", "name", self.git_user_name)
                config_writer.set_value("user", "email", self.git_user_email)
        except Exception:
            # Keep UI available even if local git config write fails.
            pass

    def _setup_editors(self) -> None:

        left_panel = QWidget(self)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 8, 8, 8)
        left_layout.setSpacing(8)

        # 软件名称标签
        app_title = QLabel(f"{__appname__} 博客工具 v{__version__}", left_panel)
        font = app_title.font()
        font.setPointSize(16)
        font.setBold(True)
        app_title.setFont(font)
        app_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(app_title)

        left_header_layout = QHBoxLayout()
        self.sync_repo_button = QPushButton("同步仓库", left_panel)
        self.sync_repo_button.clicked.connect(self.on_sync_repo)
        left_header_layout.addWidget(self.sync_repo_button)

        self.repo_config_button = QPushButton("仓库设置", left_panel)
        self.repo_config_button.clicked.connect(self.on_open_repo_config)
        left_header_layout.addWidget(self.repo_config_button)
        left_layout.addLayout(left_header_layout)

        left_manage_layout = QHBoxLayout()
        self.new_post_button = QPushButton("新建文章", left_panel)
        self.new_post_button.clicked.connect(self.on_new_article)
        left_manage_layout.addWidget(self.new_post_button)

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
        self.sort_combo.setObjectName("sortCombo")
        self.sort_combo.setFixedHeight(30)
        self.sort_combo.addItems([
            "按修改时间（新到旧）",
            "按修改时间（旧到新）",
            "按文件名（A-Z）",
        ])
        self.sort_combo.currentIndexChanged.connect(self.on_filter_or_sort_changed)
        left_layout.addWidget(self.sort_combo)

        self.article_title_label = ElidedLabel("当前文章: 未选择", left_panel)
        self.article_title_label.setMinimumWidth(0)
        left_layout.addWidget(self.article_title_label)

        self.article_list = QListWidget(left_panel)
        self.article_list.itemClicked.connect(self.on_article_item_clicked)
        self.article_list.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.article_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.article_list.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding
        )
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
        self._md_highlighter = MarkdownHighlighter(self.md_preview.document(), detect_system_theme())

        # 同步滚动
        self._scroll_guard = False

        def _sync_scroll_to_preview(value: int) -> None:
            if self._scroll_guard:
                return
            preview_bar = self.md_preview.verticalScrollBar()
            editor_bar = self.rich_editor.verticalScrollBar()
            if editor_bar.maximum() == 0:
                return
            ratio = value / editor_bar.maximum()
            self._scroll_guard = True
            preview_bar.setValue(int(ratio * preview_bar.maximum()))
            self._scroll_guard = False

        def _sync_scroll_to_editor(value: int) -> None:
            if self._scroll_guard:
                return
            editor_bar = self.rich_editor.verticalScrollBar()
            preview_bar = self.md_preview.verticalScrollBar()
            if preview_bar.maximum() == 0:
                return
            ratio = value / preview_bar.maximum()
            self._scroll_guard = True
            editor_bar.setValue(int(ratio * editor_bar.maximum()))
            self._scroll_guard = False

        self.rich_editor.verticalScrollBar().valueChanged.connect(_sync_scroll_to_preview)
        self.md_preview.verticalScrollBar().valueChanged.connect(_sync_scroll_to_editor)

        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.addWidget(left_panel)
        splitter.addWidget(self.rich_editor)
        splitter.addWidget(self.md_preview)
        splitter.setSizes([280, 560, 560])
        splitter.setStretchFactor(0, 0)  # 左侧栏：不随窗口拉伸
        splitter.setStretchFactor(1, 1)  # 编辑区：按比例伸缩
        splitter.setStretchFactor(2, 1)  # 预览区：按比例伸缩
        left_panel.setMaximumWidth(400)
        left_panel.setSizePolicy(
            QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred
        )

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
        font_controls_widget.setObjectName("fontControls")
        font_controls_layout = QHBoxLayout(font_controls_widget)
        font_controls_layout.setContentsMargins(0, 0, 0, 0)
        font_controls_layout.setSpacing(6)

        toolbar_control_height = 30

        self.font_family_combo = QFontComboBox(font_controls_widget)
        self.font_family_combo.setMinimumWidth(150)
        self.font_family_combo.setMaximumWidth(190)
        self.font_family_combo.setFixedHeight(toolbar_control_height)
        self.font_family_combo.setToolTip("选择字体")
        self.font_family_combo.setCurrentFont(self.rich_editor.font())
        self.font_family_combo.currentFontChanged.connect(self.on_font_family_changed)
        font_controls_layout.addWidget(self.font_family_combo)

        self.font_size_spin = QSpinBox(font_controls_widget)
        self.font_size_spin.setRange(8, 72)
        self.font_size_spin.setFixedWidth(70)
        self.font_size_spin.setFixedHeight(toolbar_control_height)
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
        self.more_menu_button.setFixedSize(34, toolbar_control_height)
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
        # 根据当前主题选择图标目录
        current_theme = detect_system_theme()
        if current_theme == "dark":
            state_dir = f"dark_{state}"
        else:
            state_dir = state
        
        icon_path = (
            get_resource_base_path()
            / "assets"
            / "icons"
            / "theme"
            / str(self._icon_pixel_size)
            / state_dir
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

    def _get_stylesheet(self, theme: str) -> str:
        """返回指定主题的样式表。"""
        dark_arrow_path = (
            get_resource_base_path() / "assets" / "icons" / "down-arrow-dark.svg"
        ).resolve().as_posix()
        light_arrow_path = (
            get_resource_base_path() / "assets" / "icons" / "down-arrow-light.svg"
        ).resolve().as_posix()
        dark_up_arrow_path = (
            get_resource_base_path() / "assets" / "icons" / "up-arrow-dark.svg"
        ).resolve().as_posix()
        light_up_arrow_path = (
            get_resource_base_path() / "assets" / "icons" / "up-arrow-light.svg"
        ).resolve().as_posix()

        if theme == "dark":
            dark_stylesheet = """
            QMainWindow {
                background: #1e1e1e;
                color: #e0e0e0;
            }
            QToolBar {
                background: #2d2d2d;
                border-bottom: 1px solid #3d3d3d;
                spacing: 8px;
                padding: 6px;
            }
            QToolBar::separator {
                width: 1px;
                margin: 4px 6px;
                background: #444444;
            }
            QToolBar QToolButton {
                background: #3b3b3b;
                color: #e0e0e0;
                border: 1px solid #4b4b4b;
                border-radius: 6px;
                padding: 0;
                min-width: 32px;
                min-height: 30px;
            }
            QToolBar QToolButton:hover {
                background: #4a4a4a;
                border: 1px solid #636363;
            }
            QToolBar QToolButton:pressed {
                background: #565656;
                border: 1px solid #6d6d6d;
            }
            QToolBar QWidget#fontControls {
                background: transparent;
            }
            QToolBar QWidget#fontControls QComboBox,
            QToolBar QWidget#fontControls QSpinBox {
                background: #3b3b3b;
                color: #e0e0e0;
                border: 1px solid #4b4b4b;
                border-radius: 6px;
                padding: 0 8px;
                min-height: 30px;
            }
            QToolBar QWidget#fontControls QComboBox:hover,
            QToolBar QWidget#fontControls QSpinBox:hover {
                background: #4a4a4a;
                border: 1px solid #636363;
            }
            QToolBar QWidget#fontControls QComboBox::drop-down {
                width: 22px;
                border-left: 1px solid #525252;
                background: #343434;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            QToolBar QWidget#fontControls QComboBox::down-arrow {
                image: url(__DOWN_ARROW_DARK__);
                width: 12px;
                height: 12px;
            }
            QToolBar QWidget#fontControls QSpinBox::up-button,
            QToolBar QWidget#fontControls QSpinBox::down-button {
                width: 16px;
                border-left: 1px solid #525252;
                background: #343434;
            }
            QToolBar QWidget#fontControls QSpinBox::up-button:hover,
            QToolBar QWidget#fontControls QSpinBox::down-button:hover {
                background: #414141;
            }
            QToolBar QWidget#fontControls QSpinBox::up-arrow {
                image: url(__UP_ARROW_DARK__);
                width: 10px;
                height: 10px;
            }
            QToolBar QWidget#fontControls QSpinBox::down-arrow {
                image: url(__DOWN_ARROW_DARK__);
                width: 10px;
                height: 10px;
            }
            QPushButton {
                background: #3d3d3d;
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                border-radius: 6px;
                padding: 6px 10px;
            }
            QPushButton:hover {
                background: #4d4d4d;
                border: 1px solid #5d5d5d;
            }
            QGroupBox {
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                border-radius: 4px;
                padding: 8px;
                margin-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
            }
            QLineEdit {
                background: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                border-radius: 4px;
                padding: 4px;
            }
            QLineEdit:focus {
                border: 1px solid #5d7dd9;
            }
            QSpinBox {
                background: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                border-radius: 4px;
                padding: 2px;
            }
            QComboBox {
                background: #3d3d3d;
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                border-radius: 4px;
                padding: 4px;
            }
            QComboBox#sortCombo {
                background: #3b3b3b;
                color: #e0e0e0;
                border: 1px solid #4b4b4b;
                border-radius: 6px;
                padding: 0 8px;
                min-height: 30px;
            }
            QComboBox#sortCombo:hover {
                background: #4a4a4a;
                border: 1px solid #636363;
            }
            QComboBox#sortCombo:focus {
                border: 1px solid #5d7dd9;
            }
            QComboBox#sortCombo::drop-down {
                width: 22px;
                border-left: 1px solid #525252;
                background: #343434;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            QComboBox#sortCombo::down-arrow {
                image: url(__DOWN_ARROW_DARK__);
                width: 12px;
                height: 12px;
            }
            QListWidget {
                background: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                border-radius: 6px;
            }
            QListWidget::item {
                padding: 6px;
            }
            QListWidget::item:selected {
                background: #5d7dd9;
                color: #ffffff;
            }
            QTextEdit {
                background: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                border-radius: 6px;
                padding: 8px;
                selection-background-color: #5d7dd9;
            }
            QStatusBar {
                background: #2d2d2d;
                color: #e0e0e0;
                border-top: 1px solid #4d4d4d;
            }
            QToolButton::menu-indicator {
                image: none;
            }
            """
            dark_stylesheet = dark_stylesheet.replace(
                "__DOWN_ARROW_DARK__", f'"{dark_arrow_path}"'
            )
            return dark_stylesheet.replace("__UP_ARROW_DARK__", f'"{dark_up_arrow_path}"')
        else:  # light theme
            light_stylesheet = """
            QMainWindow {
                background: #f3f4f6;
                color: #1f2937;
            }
            QToolBar {
                background: #ffffff;
                border-bottom: 1px solid #d1d5db;
                spacing: 8px;
                padding: 6px;
            }
            QToolBar::separator {
                width: 1px;
                margin: 4px 6px;
                background: #d1d5db;
            }
            QToolBar QToolButton {
                background: #f9fafb;
                color: #1f2937;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 0;
                min-width: 32px;
                min-height: 30px;
            }
            QToolBar QToolButton:hover {
                background: #eef2ff;
                border: 1px solid #94a3b8;
            }
            QToolBar QToolButton:pressed {
                background: #dbeafe;
                border: 1px solid #93c5fd;
            }
            QToolBar QWidget#fontControls {
                background: transparent;
            }
            QToolBar QWidget#fontControls QComboBox,
            QToolBar QWidget#fontControls QSpinBox {
                background: #f9fafb;
                color: #1f2937;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 0 8px;
                min-height: 30px;
            }
            QToolBar QWidget#fontControls QComboBox:hover,
            QToolBar QWidget#fontControls QSpinBox:hover {
                background: #eef2ff;
                border: 1px solid #94a3b8;
            }
            QToolBar QWidget#fontControls QComboBox::drop-down {
                width: 22px;
                border-left: 1px solid #cbd5e1;
                background: #f1f5f9;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            QToolBar QWidget#fontControls QComboBox::down-arrow {
                image: url(__DOWN_ARROW_LIGHT__);
                width: 12px;
                height: 12px;
            }
            QToolBar QWidget#fontControls QSpinBox::up-button,
            QToolBar QWidget#fontControls QSpinBox::down-button {
                width: 16px;
                border-left: 1px solid #cbd5e1;
                background: #f1f5f9;
            }
            QToolBar QWidget#fontControls QSpinBox::up-button:hover,
            QToolBar QWidget#fontControls QSpinBox::down-button:hover {
                background: #e2e8f0;
            }
            QToolBar QWidget#fontControls QSpinBox::up-arrow {
                image: url(__UP_ARROW_LIGHT__);
                width: 10px;
                height: 10px;
            }
            QToolBar QWidget#fontControls QSpinBox::down-arrow {
                image: url(__DOWN_ARROW_LIGHT__);
                width: 10px;
                height: 10px;
            }
            QPushButton {
                background: #f8fafc;
                color: #1f2937;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 6px 10px;
            }
            QPushButton:hover {
                background: #e2e8f0;
            }
            QGroupBox {
                color: #1f2937;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 8px;
                margin-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
            }
            QLineEdit {
                background: #ffffff;
                color: #1f2937;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 4px;
            }
            QLineEdit:focus {
                border: 1px solid #3b82f6;
            }
            QSpinBox {
                background: #ffffff;
                color: #1f2937;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 2px;
            }
            QComboBox {
                background: #f9fafb;
                color: #1f2937;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 4px;
            }
            QComboBox#sortCombo {
                background: #f9fafb;
                color: #1f2937;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 0 8px;
                min-height: 30px;
            }
            QComboBox#sortCombo:hover {
                background: #eef2ff;
                border: 1px solid #94a3b8;
            }
            QComboBox#sortCombo:focus {
                border: 1px solid #3b82f6;
            }
            QComboBox#sortCombo::drop-down {
                width: 22px;
                border-left: 1px solid #cbd5e1;
                background: #f1f5f9;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            QComboBox#sortCombo::down-arrow {
                image: url(__DOWN_ARROW_LIGHT__);
                width: 12px;
                height: 12px;
            }
            QListWidget {
                background: #ffffff;
                color: #1f2937;
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
                color: #1f2937;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px;
                selection-background-color: #c7d2fe;
            }
            QStatusBar {
                background: #ffffff;
                color: #1f2937;
                border-top: 1px solid #d1d5db;
            }
            QToolButton::menu-indicator {
                image: none;
            }
            """
            light_stylesheet = light_stylesheet.replace(
                "__DOWN_ARROW_LIGHT__", f'"{light_arrow_path}"'
            )
            return light_stylesheet.replace(
                "__UP_ARROW_LIGHT__", f'"{light_up_arrow_path}"'
            )

    def _apply_styles(self) -> None:
        """根据系统主题自动应用相应样式。"""
        theme = detect_system_theme()
        self.setStyleSheet(self._get_stylesheet(theme))

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

    def on_open_repo_config(self) -> None:
        dialog = RepoCommitConfigDialog(
            local_repo_path=self.local_repo_path,
            remote_repo_url=self.remote_repo_url,
            branch=self.git_branch,
            git_user_name=self.git_user_name,
            git_user_email=self.git_user_email,
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        values = dialog.values()
        self.local_repo_path = values["local_repo_path"]
        self.remote_repo_url = values["remote_repo_url"]
        self.git_branch = values["branch"]
        self.git_user_name = values["git_user_name"]
        self.git_user_email = values["git_user_email"]

        self._save_settings()
        self._ensure_git_identity()
        self.refresh_article_list()
        self.statusBar().showMessage("仓库与提交配置已保存", 2500)

    def on_sync_repo(self) -> None:
        if self.repo_sync_thread and self.repo_sync_thread.isRunning():
            return

        if not self.local_repo_path:
            QMessageBox.warning(self, "未配置仓库", "请先在“仓库设置”中配置本地仓库路径。")
            return
        if not self.remote_repo_url:
            QMessageBox.warning(self, "未配置仓库", "请先在“仓库设置”中配置远端仓库 URL。")
            return

        self.statusBar().showMessage("正在拉取仓库...")
        self.repo_sync_thread = RepoSyncThread(
            local_repo_path=self.local_repo_path,
            remote_repo_url=self.remote_repo_url,
            branch=self.git_branch,
            git_user_name=self.git_user_name,
            git_user_email=self.git_user_email,
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
        if not self.local_repo_path.strip():
            self.article_list.clear()
            self.article_records = []
            return

        selected_path = str(self.current_article_path) if self.current_article_path else ""
        self.article_records = self._collect_article_files()

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
        self.article_title_label.setToolTip(title)
        self.setWindowTitle(f"{__appname__} 博客工具 v{__version__} - {title}")

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
        filename, ok = QInputDialog.getText(self, "新建文章", "请输入文章标题或文件名：")
        if not ok or not filename.strip():
            return

        # 允许用户输入中文标题，自动转为文件名
        title = filename.strip()
        # 文件名仅保留字母数字下划线和短横线，中文转拼音可选，这里直接用短横线替换
        sanitized = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5\-_]+', "-", title).strip(".-")
        # 去除多余的连续短横线
        sanitized = re.sub(r'-+', '-', sanitized)
        if not sanitized:
            QMessageBox.warning(self, "无效名称", "文章标题不能为空。")
            return

        # 自动补全.md后缀
        if not sanitized.lower().endswith('.md'):
            filename_md = sanitized + '.md'
        else:
            filename_md = sanitized

        posts_dir = self._configured_posts_dir(show_warning=True, create_if_missing=True)
        if posts_dir is None:
            return
        article_path = posts_dir / filename_md
        if article_path.exists():
            QMessageBox.warning(self, "已存在", "同名文章已存在，请更换标题或文件名。")
            return

        # front matter 的 title 字段用原始输入
        initial_front_matter = build_front_matter(
            title,
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
        self.article_title_label.setToolTip(title or '未命名')

    def start_git_push(self, commit_message: str) -> None:
        if self.git_thread and self.git_thread.isRunning():
            QMessageBox.information(self, "提示", "当前已有 Git 推送任务正在执行。")
            return

        if not self.local_repo_path:
            QMessageBox.warning(self, "未配置仓库", "请先在“仓库设置”中配置本地仓库路径。")
            return

        self.git_thread = GitPushThread(
            repo_path=self.local_repo_path,
            branch=self.git_branch,
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
