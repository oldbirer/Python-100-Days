# 在 VS Code 中处理 Word（.docx）文档

## 1. 为什么 VS Code 无法原生编辑 .docx 文件？

`.docx` 文件本质上是一个经过 ZIP 压缩的 **OOXML（Office Open XML）** 包，内部包含 XML 文件、图片、样式表等多种资源。  
VS Code 是以**纯文本编辑**为核心的代码编辑器，它能够读取 `.docx` 压缩包中的原始二进制数据，但无法将其解析并呈现为带格式的文档（即 WYSIWYG，所见即所得）。因此，直接用 VS Code 打开 `.docx` 时，会看到乱码或收到"无法显示该文件"的提示。

---

## 2. 可行方案

### 2.1 仅预览：使用 VS Code 扩展

如果只需要在 VS Code 中**预览** Word 文档内容，可以安装以下扩展：

| 扩展名称 | 扩展 ID | 说明 |
|---|---|---|
| **Office Viewer (docx, xlsx...)** | `cweijan.vscode-office` | 支持预览 .docx、.xlsx、.pdf 等 Office 文件 |
| **vscode-office** | `HookyQR.office` | 轻量级 Office 文件查看器 |

**安装步骤：**

1. 打开 VS Code，按下 `Ctrl+Shift+X`（macOS：`Cmd+Shift+X`）打开扩展面板。
2. 在搜索框中输入扩展名称（如 `Office Viewer`）。
3. 点击 **Install** 安装。
4. 安装完成后，在资源管理器中右键单击 `.docx` 文件，选择 **Open With → Office Viewer** 即可预览。

> **注意**：扩展仅提供预览功能，**不支持编辑并保留原有格式**。

---

### 2.2 真正编辑并保留格式：使用专业 Office 套件

若需完整保留字体、段落、图表、页眉页脚等排版信息，必须使用支持 WYSIWYG 的专业软件：

| 软件 | 平台 | 费用 |
|---|---|---|
| **Microsoft Word** | Windows / macOS | 商业授权 |
| **LibreOffice Writer** | Windows / macOS / Linux | 免费开源 |
| **OnlyOffice Desktop** | Windows / macOS / Linux | 免费开源（社区版） |
| **WPS Office** | Windows / macOS / Linux | 免费（国内常用） |

VS Code 本身不适合进行 WYSIWYG 文档编辑，建议将上述工具作为 `.docx` 的主要编辑环境。

---

### 2.3 文本化工作流：使用 pandoc 转换

如果希望**在 VS Code 中编写内容，最终输出为 Word 格式**，可以借助 [pandoc](https://pandoc.org/) 在 Markdown 与 .docx 之间互相转换。

#### 安装 pandoc

- **Windows**：从 [pandoc 官网](https://pandoc.org/installing.html) 下载安装包，或使用 `winget`：
  ```bash
  winget install JohnMacFarlane.Pandoc
  ```
- **macOS**：
  ```bash
  brew install pandoc
  ```
- **Linux（Debian/Ubuntu）**：
  ```bash
  sudo apt install pandoc
  ```

#### 将 .docx 转换为 Markdown（在 VS Code 中编辑）

```bash
pandoc input.docx -o output.md
```

执行后，在 VS Code 中打开 `output.md` 即可进行编辑。

#### 将 Markdown 导出回 .docx

```bash
pandoc output.md -o final.docx
```

#### 使用自定义样式模板（保留公司/团队排版风格）

```bash
# 先生成默认模板
pandoc --print-default-data-file reference.docx > reference.docx

# 用 Word 编辑 reference.docx 中的样式，然后在导出时引用
pandoc output.md --reference-doc=reference.docx -o final.docx
```

> **说明**：pandoc 转换后的 Markdown 会保留大部分文字内容，但复杂表格、嵌套图文框等排版元素可能丢失，请在转换后仔细核对。

---

### 2.4 仅需提取文本内容

若只需要获取 `.docx` 中的纯文本，可使用以下方式：

- **另存为纯文本**：在 Word / LibreOffice 中选择 **另存为 → 纯文本（.txt）**，再在 VS Code 中打开。
- **另存为 PDF**：选择 **另存为 → PDF**，再用 VS Code 的 PDF 预览扩展查看。
- **复制粘贴**：在 Word / LibreOffice 中全选（`Ctrl+A`）并复制，粘贴到 VS Code 新文件中。
- **使用 Python 提取文本**（适合批量处理）：
  ```bash
  pip install python-docx
  ```
  ```python
  from docx import Document

  doc = Document("input.docx")
  for para in doc.paragraphs:
      print(para.text)
  ```

---

## 3. 常见问题排查

### 3.1 安装扩展后仍然无法预览

- **检查扩展是否已启用**：打开扩展面板，确认扩展状态为 *已启用（Enabled）*，而非已禁用或受工作区限制。
- **确认文件关联**：右键单击 `.docx` 文件 → **Open With**，选择对应的扩展，而非默认的文本编辑器。
- **重新加载窗口**：按 `Ctrl+Shift+P`，输入 `Reload Window` 并执行。
- **查看扩展输出日志**：`查看（View）→ 输出（Output）`，在下拉框中选择扩展名称，查看错误信息。

### 3.2 文件过大导致预览卡顿或失败

- `.docx` 文件若包含大量高分辨率图片，体积可能超过扩展的处理上限。
- 建议先用 Word / LibreOffice 压缩图片后再进行预览，或直接使用专业软件打开。

### 3.3 受限环境（企业策略 / 教育网络）

- 部分企业或学校会通过组策略禁止安装 VS Code 扩展。
- 此时可使用 **pandoc 命令行工具**在本地转换文件，无需额外扩展权限。

### 3.4 远程开发 / Dev Container 场景

- 在 **VS Code Remote SSH** 或 **Dev Container** 中，扩展需安装在**远程端**，而非本地。
- 打开远程窗口后，在扩展面板中点击 **Install in Remote/Container** 按钮进行安装。
- 若远程服务器无 GUI 环境，预览扩展可能无法正常渲染，建议将文件下载到本地后再用专业软件打开。

---

## 参考资源

- [pandoc 官方文档](https://pandoc.org/MANUAL.html)
- [python-docx 文档](https://python-docx.readthedocs.io/)
- [VS Code 扩展市场](https://marketplace.visualstudio.com/vscode)
- [LibreOffice 官网](https://www.libreoffice.org/)
- [OnlyOffice 官网](https://www.onlyoffice.com/)
