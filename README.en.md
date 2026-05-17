<div align="center">

<h1>MoZu</h1>

<p>A local writing and publishing client focused on Hugo workflows.</p>

<p><a href="README.md">中文说明</a></p>

<p><img src="assets/icons/mozu.svg" alt="MoZu Icon" width="120"></p>

<p>
	<img src="https://img.shields.io/badge/version-1.0.7-0f766e" alt="Version">
	<img src="https://img.shields.io/badge/python-3.11+-2563eb" alt="Python">
	<img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-475569" alt="Platform">
	<img src="https://img.shields.io/badge/license-MIT-16a34a" alt="License">
</p>

</div>

## Product Positioning

MoZu is built for creators who want a local, focused writing-to-publishing flow, especially for Hugo blogs. It lets you edit content, maintain Front Matter, import documents, and publish via Git in one desktop app.

## Key Highlights

- Dual-pane writing with rich text editing and synchronized Markdown preview.
- Full post management: create, rename, delete, search, and sort.
- Visual Front Matter editing to reduce metadata mistakes.
- Docx import and conversion for faster content migration.
- Integrated Git publishing for a faster draft-to-release cycle.

## Feature Overview

### Editing And Preview

- Markdown highlighting, formatting toolbar, and quick insert blocks.

### Metadata Management

- Direct editing and write-back for title, tags, categories, and draft.

### Content Import

- Import Word documents and convert them into blog-ready Markdown.

### Publishing Workflow

- Works with local repositories and supports Git commit/push workflows.

## Screenshots

Add application screenshots here (main window, editor, publishing flow) for better product presentation.

## Quick Start

### 1. Requirements

- Python 3.11+
- Windows, macOS, or Linux

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run

```bash
python pyqt_main_window.py
```

## Download And Build

This repository includes a GitHub Actions workflow:

- Pushing to `main` triggers automatic builds.
- Pushing `v*` tags builds and publishes multi-platform binaries to GitHub Releases.

### Release Artifacts

- Windows amd64: `Mozhu_Blog_Tool_Windows_x86.zip` (contains onefile exe)
- macOS ARM64: `Mozhu_Blog_Tool_Mac_ARM64.zip` (.app)
- Linux amd64: `Mozhu_Blog_Tool_Linux_amd64.AppImage`

### Packaging Notes

- Windows and Linux use PyInstaller onefile mode (`-F`).
- macOS builds a windowed app bundle (`.app`).
- Toolbar and app icon visuals are unified based on `assets/icons/mozu.svg`.

## Version

- Current: 1.0.7
- Product Name: MoZu (墨筑)

## Roadmap

- Add drag-and-drop image upload with automatic path management.
- Improve multi-repository configuration and quick switching.
- Add template system and workflow presets.

## Contributing

Issues and pull requests are welcome.

## License

MIT
