<div align="center">

<h1>墨筑 MoZu</h1>

<p>一款专注 Hugo 工作流的本地写作与发布客户端。</p>

<p><a href="README.en.md">English README</a></p>

<p><img src="assets/icons/mozu.svg" alt="墨筑图标" width="120"></p>

<p>
	<img src="https://img.shields.io/badge/version-1.0-0f766e" alt="版本">
	<img src="https://img.shields.io/badge/python-3.11+-2563eb" alt="Python">
	<img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS-475569" alt="平台">
	<img src="https://img.shields.io/badge/license-MIT-16a34a" alt="协议">
</p>

</div>

## 产品定位

墨筑（MoZu）面向希望在本地完成“写作-整理-发布”闭环的博客作者，特别适合 Hugo 用户。你可以在一个桌面应用中完成文章编辑、Front Matter 维护、文档导入与 Git 发布，避免在多个工具之间频繁切换。

## 核心亮点

- 双栏写作体验，富文本编辑与 Markdown 源码预览同步。
- 文章管理完整，支持新建、重命名、删除、搜索与排序。
- Front Matter 可视化维护，降低手写元数据错误率。
- 支持 Docx 导入与转换，加速存量内容迁移。
- 集成 Git 发布能力，提升从草稿到上线的效率。

## 功能一览

### 编辑与预览

- Markdown 高亮、标题与格式工具栏、常用内容块插入。

### 元数据管理

- title、tags、categories、draft 字段可直接编辑并回写。

### 内容导入

- 导入 Word 文档并转换为适配博客的 Markdown。

### 发布流程

- 与本地仓库协作，支持 Git 提交与推送流程。

## 界面预览

你可以在此处放置应用截图（例如主界面、编辑器、发布流程）以获得更好的展示效果。

## 快速开始

### 1. 环境要求

- Python 3.11+
- Windows 或 macOS

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 运行应用

```bash
python pyqt_main_window.py
```

## 下载与构建

仓库已配置 GitHub Actions 工作流：推送到 `main` 分支会自动构建，推送 `v*` 标签会自动产出并发布 Windows 与 macOS 安装包（zip）。

## 版本信息

- 当前版本：1.0
- 产品名称：墨筑 (MoZu)

## 路线图

- 增加图文拖拽上传与自动路径管理。
- 增强多仓库配置与一键切换。
- 完善模板系统与写作工作流预设。

## 贡献

欢迎提交 Issue 与 PR，一起完善墨筑。

## 许可证

MIT
