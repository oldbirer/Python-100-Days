# 在 VS Code 中处理 Word（.docx）文件指南

> 语言：简体中文 | 适用仓库类型：写作、教程、技术文档

---

## 1. 为什么 VS Code 默认无法直接编辑 .docx？

`.docx` 是 **Office Open XML（OOXML）** 格式的文件，本质上是一个 **ZIP 压缩包**，内部包含 XML、图片、字体等多种二进制和文本资源，并非普通的纯文本文件。

VS Code 是一款**面向纯文本和代码**的编辑器，当它尝试直接打开 `.docx` 时，会以二进制方式读取压缩数据，因此你看到的是一堆乱码，而不是可读的正文内容。这是文件格式本身的限制，与 VS Code 的版本无关。

---

## 2. 可行方案

### 2.1 在 VS Code 中预览（只读）

推荐安装以下插件之一：

| 插件名称 | 插件 ID | 说明 |
|---|---|---|
| **Office Viewer** | `cweijan.vscode-office` | 支持 docx / xlsx / pptx 预览，免费 |
| **vscode-office** | `vscode-office.vscode-office` | 同类插件，界面简洁 |
| **Docs View** | `bierner.docs-view` | 轻量级文档预览 |

**安装步骤（以 Office Viewer 为例）：**

1. 打开 VS Code，按 `Ctrl+Shift+X`（macOS：`Cmd+Shift+X`）打开扩展面板。
2. 搜索 `Office Viewer` 或插件 ID `cweijan.vscode-office`。
3. 点击 **Install（安装）**，等待安装完成。
4. 在资源管理器中双击 `.docx` 文件，插件会自动接管并渲染预览。

> **注意**：预览插件通常为只读模式，无法在 VS Code 中修改内容后保存回 `.docx`。如需编辑，请参考下节方案。

---

### 2.2 编辑 .docx 文件

#### 方案 A：使用专用 Office 软件编辑

最简单、保真度最高的方式是使用支持 `.docx` 格式的桌面软件：

- **Microsoft Word**（Windows / macOS）
- **WPS Office**（Windows / macOS / Linux，免费）
- **LibreOffice Writer**（全平台，开源免费）
- **OnlyOffice Desktop**（全平台，开源免费，兼容性好）

这些软件能完整保留样式、表格、图片、页眉页脚等富文本元素。

---

#### 方案 B：docx ↔ Markdown 双向转换（Pandoc 工作流）

如果你希望在 VS Code 中用 Markdown 编写内容，再生成 `.docx`，可以使用 [**Pandoc**](https://pandoc.org/) 进行格式转换。

**安装 Pandoc：**

```bash
# macOS（Homebrew）
brew install pandoc

# Ubuntu / Debian
sudo apt install pandoc

# Windows（Scoop）
scoop install pandoc

# 或从官网下载安装包：https://pandoc.org/installing.html
```

**常用命令示例：**

```bash
# Markdown → Word（生成 .docx）
pandoc input.md -o output.docx

# Word → Markdown（从 .docx 提取文本）
pandoc input.docx -o output.md

# 使用自定义样式模板生成 Word（推荐，保持公司/学校样式）
pandoc input.md --reference-doc=my-template.docx -o output.docx

# 生成样式模板文件（首次使用时执行，再修改该模板中的样式）
pandoc --print-default-data-file reference.docx > my-template.docx
```

**注意事项：**

| 元素 | 转换保真度 | 备注 |
|---|---|---|
| 标题 / 段落 / 列表 | ✅ 良好 | 基本无损 |
| 表格 | ⚠️ 一般 | 复杂合并单元格可能丢失格式 |
| 图片 | ⚠️ 一般 | 本地图片可嵌入，网络图片需手动处理 |
| 公式（LaTeX） | ✅ 良好 | 需安装 `pandoc-crossref` 等过滤器 |
| 复杂样式 / 页眉页脚 | ❌ 较差 | 建议使用参考模板（`--reference-doc`）补救 |
| 批注 / 修订记录 | ❌ 不支持 | 双向转换时会丢失 |

---

### 2.3 备选方案

- **导出为 PDF 仅查看**：在 Word / WPS 中打开后另存为 PDF，然后用 VS Code 插件（如 vscode-pdf）或系统自带阅读器预览。
- **复制文本到 Markdown / 纯文本**：将 Word 中的正文粘贴到 `.md` 或 `.txt` 文件，在 VS Code 中编辑纯文本版本，格式须手动重排。
- **使用在线转换工具**：如 [Word to Markdown Converter](https://word2md.com/) 或 [CloudConvert](https://cloudconvert.com/)，上传后下载 Markdown 文件（注意隐私风险，不要上传敏感文档）。

---

## 3. 常见报错与现象排查

| 现象 / 报错 | 原因 | 解决方法 |
|---|---|---|
| 打开后显示二进制乱码（`PK\x03\x04...`） | VS Code 以纯文本读取了 ZIP 压缩数据 | 安装 Office Viewer 等预览插件，或用专用软件打开 |
| 提示"无法显示此文件类型" | 未安装对应预览插件 | 安装 `cweijan.vscode-office` 插件 |
| 插件已安装但无法预览 | 插件未激活或文件关联被覆盖 | 右键文件 → 选择打开方式 → 选择 Office Viewer；或禁用其他冲突插件 |
| 文件过大（>10 MB）导致预览卡顿/崩溃 | 插件渲染能力有限 | 改用桌面 Office 软件打开 |
| Pandoc 转换后样式丢失 | 未指定参考模板 | 使用 `--reference-doc` 参数指定已设置好样式的 `.docx` 模板 |
| Pandoc 命令找不到（`command not found`） | Pandoc 未安装或未加入 PATH | 按第 2.2 节重新安装，并确认 PATH 配置 |

---

## 4. 推荐路径

根据仓库用途，建议选择以下工作流：

### ✅ 推荐：以 Markdown 为主（写作 / 教程 / 技术文档类仓库）

```
编写 .md → Git 版本管理 → 渲染为 HTML / PDF
```

- Markdown 是纯文本，天然适合 Git diff / Review。
- VS Code 原生支持 Markdown 预览（`Ctrl+Shift+V`）。
- 可用 [MkDocs](https://www.mkdocs.org/)、[Docusaurus](https://docusaurus.io/) 等工具自动生成文档站点。
- 本仓库（Python-100-Days）即采用此方案，所有教程均为 `.md` 格式。

### 📄 必须交付 Word 时：Markdown → Pandoc → .docx

```
编写 .md → pandoc input.md --reference-doc=template.docx -o output.docx → 交付
```

1. 在 VS Code 中用 Markdown 撰写内容。
2. 准备一份符合要求样式的 Word 模板（`template.docx`）。
3. 执行 Pandoc 命令生成最终 `.docx`，交付后不再回编辑。
4. 后续修改仍在 `.md` 文件中进行，再重新生成。

> **不建议**将 `.docx` 文件纳入 Git 仓库进行版本管理，因为其二进制格式无法产生有意义的 diff，且文件体积较大。如必须归档，建议同时保留对应的 `.md` 源文件。
