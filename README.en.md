# MoZu

> A local writing and publishing client focused on Hugo workflows.

[中文说明](README.md)

![MoZu Icon](assets/icons/mozu.svg)

![Version](https://img.shields.io/badge/version-1.0-0f766e)
![Python](https://img.shields.io/badge/python-3.11+-2563eb)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-475569)
![License](https://img.shields.io/badge/license-MIT-16a34a)

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
- Windows or macOS

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run

```bash
python pyqt_main_window.py
```

## Download And Build

This repository includes a GitHub Actions workflow: pushing to `main` triggers builds, and pushing `v*` tags builds and publishes Windows/macOS zip packages.

## Version

- Current: 1.0
- Product Name: MoZu (墨筑)

## Roadmap

- Add drag-and-drop image upload with automatic path management.
- Improve multi-repository configuration and quick switching.
- Add template system and workflow presets.

## Contributing

Issues and pull requests are welcome.

## License

MIT
