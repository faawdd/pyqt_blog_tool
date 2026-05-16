from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable

from git import GitCommandError, Repo
from PyQt6.QtCore import QThread, pyqtSignal


def _sanitize_filename(name: str) -> str:
    sanitized = re.sub(r'[\\/:*?"<>|]', "", name).strip()
    sanitized = re.sub(r"\s+", "-", sanitized)
    return sanitized or "untitled-post"


def build_front_matter(
    title: str,
    tags: Iterable[str] | None = None,
    categories: Iterable[str] | None = None,
    draft: bool = False,
    date_iso: str | None = None,
) -> str:
    tz = timezone(timedelta(hours=8))
    now_iso = date_iso or datetime.now(tz).isoformat(timespec="seconds")
    safe_title = title.replace('"', '\\"').strip() or "Untitled"
    tags_list = [tag.strip() for tag in (tags or []) if tag.strip()]
    categories_list = [cat.strip() for cat in (categories or []) if cat.strip()]
    tags_yaml = ", ".join(f'"{tag}"' for tag in tags_list)
    categories_yaml = ", ".join(f'"{cat}"' for cat in categories_list)
    return (
        "---\n"
        f'title: "{safe_title}"\n'
        f"date: {now_iso}\n"
        f"tags: [{tags_yaml}]\n"
        f"categories: [{categories_yaml}]\n"
        f"draft: {'true' if draft else 'false'}\n"
        "---\n\n"
    )


def save_hugo_post(markdown_body: str, title: str, content_posts_dir: str) -> Path:
    output_dir = Path(content_posts_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    filename = f"{_sanitize_filename(title)}.md"
    output_path = output_dir / filename

    full_text = build_front_matter(title) + markdown_body.strip() + "\n"
    output_path.write_text(full_text, encoding="utf-8")
    return output_path


class GitPushThread(QThread):
    result_signal = pyqtSignal(bool, str)

    def __init__(
        self,
        repo_path: str,
        branch: str = "main",
        commit_message: str = "Publish post",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.repo_path = repo_path
        self.branch = branch
        self.commit_message = commit_message

    def run(self) -> None:
        try:
            repo = Repo(self.repo_path)
            repo.git.add(".")

            try:
                repo.index.commit(self.commit_message)
            except GitCommandError as commit_error:
                # Allow continue if there is nothing new to commit.
                if "nothing to commit" not in str(commit_error).lower():
                    raise
            except Exception as commit_error:
                if "nothing to commit" not in str(commit_error).lower():
                    raise

            repo.remote(name="origin").push(self.branch)
            self.result_signal.emit(True, "已成功推送到 GitHub。")
        except Exception as exc:
            self.result_signal.emit(False, f"推送失败: {exc}")


class RepoSyncThread(QThread):
    result_signal = pyqtSignal(bool, str)

    def __init__(
        self,
        local_repo_path: str,
        remote_repo_url: str,
        branch: str = "main",
        git_user_name: str | None = None,
        git_user_email: str | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.local_repo_path = Path(local_repo_path)
        self.remote_repo_url = remote_repo_url
        self.branch = branch
        self.git_user_name = git_user_name
        self.git_user_email = git_user_email

    def _ensure_repo_identity(self, repo: Repo) -> None:
        if not (self.git_user_name and self.git_user_email):
            return
        with repo.config_writer() as config_writer:
            config_writer.set_value("user", "name", self.git_user_name)
            config_writer.set_value("user", "email", self.git_user_email)

    def run(self) -> None:
        try:
            git_dir = self.local_repo_path / ".git"

            if not git_dir.exists():
                if self.local_repo_path.exists() and any(self.local_repo_path.iterdir()):
                    # 目录非空但非 Git 仓库：就地 init 并关联远端
                    repo = Repo.init(str(self.local_repo_path))
                    self._ensure_repo_identity(repo)
                    origin = repo.create_remote("origin", self.remote_repo_url)
                    try:
                        origin.fetch()
                        remote_refs = [ref.name for ref in origin.refs]
                        tracking = f"origin/{self.branch}"
                        if tracking in remote_refs:
                            repo.git.checkout("-b", self.branch, tracking)
                    except GitCommandError:
                        # 远端为空或分支不存在，本地已 init 完成，忽略 fetch 错误
                        pass
                    self.result_signal.emit(
                        True,
                        "已在现有目录初始化 Git 仓库并关联远端，请确认本地文件后再推送。",
                    )
                    return
                self.local_repo_path.mkdir(parents=True, exist_ok=True)
                Repo.clone_from(
                    self.remote_repo_url,
                    str(self.local_repo_path),
                    branch=self.branch,
                )
                cloned_repo = Repo(str(self.local_repo_path))
                self._ensure_repo_identity(cloned_repo)
                self.result_signal.emit(True, "仓库克隆完成。")
                return

            repo = Repo(str(self.local_repo_path))
            self._ensure_repo_identity(repo)
            repo.remote(name="origin").pull(self.branch)
            self.result_signal.emit(True, "仓库拉取完成。")
        except Exception as exc:
            self.result_signal.emit(False, f"仓库同步失败: {exc}")
