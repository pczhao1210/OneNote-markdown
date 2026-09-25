# OneNote Markdown Help

Version: `1.2.1-fix.1`. Install the package matching OneNote itself: `x86`, `x64`, or native `arm64`. The Arm64 package requires Windows 11 and .NET Framework 4.8.1.

> [中文说明](#中文说明)

## Preview workflow

- **Render Page** creates or updates one preview linked to all non-managed content on the page.
- **Render Selection** requires an actual complete selection and maintains a preview for that selected source region.
- **Render Text Box** maintains an independent preview for the text box containing the cursor.
- New previews default to the source's right side. Settings can place them below and control the gap and width.
- Moving a preview, resizing it, or editing/deleting its title is preserved on refresh. Use **Reset Layout** to reapply layout settings.
- Use **Apply Defaults** when you explicitly want to reapply the current theme, default title, and layout to the full-page preview.
- **Delete Preview** removes only the managed preview under the cursor.

Managed previews use internal metadata rather than title text or coordinates. Repeated `F5` updates the full-page preview instead of appending another copy. A legacy block is never claimed or deleted only because its title says `Markdown Render`.

## Refresh and conflicts

Manual refresh is the default. Enable delayed auto-refresh in **Settings** if desired, then explicitly enable **Live Mode**. Consecutive input is debounced; delayed work is bound to the page, text element, and source captured at trigger time. Changing pages, changing the source, typing again, or disabling Live Mode cancels stale work.

If the preview body was edited manually, automatic overwrite is paused and manual refresh asks before replacing it. Position, width, and title changes do not count as body conflicts.

## Rendering

Supported content includes headings, nested lists, task lists, block quotes, horizontal rules, variable-length backtick/tilde fences, syntax-highlighted code, GFM-style tables, HTTP/HTTPS/mail links, standalone images, and inline formatting.

- Relative image paths are resolved when importing a file. Local files are limited to 10 MB.
- Remote images are disabled by default and, when enabled, accept only HTTP/HTTPS with size and timeout limits.
- Block LaTeX supports one-line `$$...$$` and multiline blocks. A failed formula reports the reason and keeps the source.
- Common Mermaid `flowchart`/`graph` syntax is rendered locally as an image. Unsupported statements, other diagram languages, timeouts, and resource-limit failures keep the complete fenced source.
- Formula and diagram results use bounded in-memory caches that are cleared when relevant settings change.

## Import and export

Imported Markdown stores its original source and resolves relative images against the source file directory. Export omits managed preview bodies, restores stored Markdown, and excludes add-in titles. If an imported preview body was modified, export explicitly asks before using the saved original source.

Ordinary OneNote rich text is converted on a best-effort basis. Arbitrary OneNote objects are not guaranteed to round-trip losslessly.

## Settings

Settings are stored at `%AppData%\OneNoteMarkdown\settings\theme.ini`.

Key options include:

- `theme.preset=technical|study|minimal`
- `preview.title.show`, `preview.title.text`
- `preview.position=right|below`, `preview.gap`, `preview.width`
- `preview.autoRefresh`, `preview.autoRefresh.delayMs`
- `font.*`, `enable.latex.image`, `enable.code.lineNumber`
- `image.allowRemote`, `diagram.timeoutMs`
- `language=auto|zh|en`

## Shortcuts

| Shortcut | Action |
|---|---|
| `F5` | Create/update the full-page preview |
| `F8` | Copy exported Markdown |
| `Ctrl+\` | Toggle Live Mode |
| `Ctrl+Enter` | Render the current line in place |

Logs contain status, timing, and errors but not Markdown source text:

`%AppData%\OneNoteMarkdown\logs\onenotemarkdown.log`

---

<a name="中文说明"></a>

# OneNote Markdown 中文说明

版本：`1.2.1-fix.1`。安装包必须匹配 OneNote 本身的架构：`x86`、`x64` 或原生 `arm64`；Arm64 安装包要求 Windows 11 和 .NET Framework 4.8.1。

## 对照预览

- **渲染整页**：为页面全部非托管内容创建或更新一个关联预览。
- **渲染选区**：仅接受真实、完整的选区，并按所选来源区域维护预览。
- **渲染正文框**：为光标所在正文框维护独立预览。
- 新预览默认在来源右侧；可在设置中改为下方，并调整间距和宽度。
- 移动、调整宽度、修改或删除标题后，刷新会保留这些改动；需要恢复时使用 **重置布局**。
- 需要将当前主题、默认标题和布局显式重新应用到整页预览时，使用 **应用预设**。
- **删除预览**只删除光标所在的插件托管预览。

预览身份使用内部元数据，不依赖标题或坐标。连续按 `F5` 会更新整页预览，不会重复追加。旧内容不会仅因标题为 `Markdown Render` 就被认定或删除。

## 刷新与冲突

默认使用手动刷新。如需编辑后延迟自动刷新，请在设置中启用，并显式开启 **实时模式**。连续输入会合并任务；任务绑定触发时的页面、正文元素和源码。切页、源码变化、继续输入或关闭实时模式都会取消过期任务。

检测到预览正文被手动编辑时，自动覆盖会暂停，手动刷新会询问是否覆盖。位置、宽度和标题变化不算正文冲突。

## 渲染能力

支持标题、嵌套列表、任务列表、引用、分隔线、不同长度的反引号/波浪线围栏、代码高亮、GFM 风格表格、HTTP/HTTPS/邮件链接、独立图片以及常见行内格式。

- 导入文件时，相对图片路径按源文件目录解析；本地文件限制 10 MB。
- 远程图片默认关闭；启用后仅接受 HTTP/HTTPS，并限制大小和超时。
- 块级公式支持单行 `$$...$$` 和多行形式；失败会显示原因并保留公式源码。
- 常见 Mermaid `flowchart`/`graph` 在本地离线生成图片；不支持的语句、其他图表语言、超时或资源限制会回退为完整围栏源码。
- 公式和图表使用容量受限的内存缓存，相关设置变化时自动失效。

## 导入导出

导入时保存完整原始 Markdown，并按源文件目录解析相对图片。导出会排除托管预览正文、恢复保存的 Markdown，并排除插件标题。若导入预览正文被修改，导出前会明确询问是否使用保存的原始源码。

普通 OneNote 富文本采用尽力转换，不承诺任意 OneNote 对象完全无损往返。

## 设置

设置文件：`%AppData%\OneNoteMarkdown\settings\theme.ini`

主要选项：

- `theme.preset=technical|study|minimal`
- `preview.title.show`、`preview.title.text`
- `preview.position=right|below`、`preview.gap`、`preview.width`
- `preview.autoRefresh`、`preview.autoRefresh.delayMs`
- `font.*`、`enable.latex.image`、`enable.code.lineNumber`
- `image.allowRemote`、`diagram.timeoutMs`
- `language=auto|zh|en`

## 快捷键

| 快捷键 | 功能 |
|---|---|
| `F5` | 创建/更新整页预览 |
| `F8` | 复制导出的 Markdown |
| `Ctrl+\` | 开启/关闭实时模式 |
| `Ctrl+Enter` | 原地渲染当前行 |

日志只记录状态、耗时和错误，不记录 Markdown 源码：

`%AppData%\OneNoteMarkdown\logs\onenotemarkdown.log`
