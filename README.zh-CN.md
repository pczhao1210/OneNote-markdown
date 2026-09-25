[English](README.md) | **中文**

# OneNote Markdown

在 Microsoft OneNote 中编写和渲染 Markdown 的插件。支持托管对照预览、LaTeX 公式、离线 Mermaid 流程图、表格、图片、链接和代码高亮，界面支持中英文切换。

![OneNote Markdown 演示截图](docs/demo.png)

## 下载

请从仓库的 [Releases](https://github.com/oldding/OneNote-markdown/releases) 页面下载：

- `OneNoteMarkdownSetup-1.2.1-fix-x86.exe`：适用于 32 位 OneNote
- `OneNoteMarkdownSetup-1.2.1-fix-x64.exe`：适用于 x64 OneNote
- `OneNoteMarkdownSetup-1.2.1-fix-arm64.exe`：适用于 Windows 11 原生 Arm64 OneNote，要求 .NET Framework 4.8.1

> 安装包必须匹配 **OneNote 的位数**，不是 Windows 的位数。64 位 Windows 也可能安装了 32 位 OneNote。

## 功能

| 功能 | 说明 |
|------|------|
| **托管对照预览** | 为整页、选区或当前正文框创建右侧/下方预览，重复渲染原位更新 |
| **刷新与布局管理** | 默认手动刷新，可选防抖自动刷新，并支持冲突检测、重置布局、删除预览和区域级跳转源码 |
| **LaTeX 公式** | 块级公式本地渲染、容量受限缓存及保留源码的错误回退 |
| **Mermaid 流程图** | 常见 Mermaid 流程图在本地离线生成图片，不上传笔记内容 |
| **代码高亮** | 多种语言语法高亮，支持行号显示 |
| **导入/导出** | 保留导入文件源码和相对图片路径，导出时避免源码与托管预览重复 |
| **剪贴板支持** | 一键复制页面为 Markdown |
| **设置对话框** | 图形化配置样式预设、字体、标题、预览布局、刷新延迟、图片及界面语言 |

## 快捷键

| 快捷键 | 功能 |
|--------|------|
| `F5` | 渲染整页 |
| `F8` | 复制 Markdown 到剪贴板 |
| `Ctrl+\` | 开启/关闭实时模式 |

## 使用方法

1. 查看 OneNote 的位数：`文件 → 帐户 → 关于 OneNote`
2. 32 位 OneNote 安装 `OneNoteMarkdownSetup-1.2.1-fix-x86.exe`
3. x64 OneNote 安装 `OneNoteMarkdownSetup-1.2.1-fix-x64.exe`
4. 原生 Arm64 OneNote 安装 `OneNoteMarkdownSetup-1.2.1-fix-arm64.exe`
4. 安装后，OneNote 功能区会出现 "Markdown" 选项卡
5. 在页面中编写 Markdown 文本，点击"渲染整页"、"渲染正文框"或"渲染选区"
6. 对同一来源再次渲染会更新关联预览，不会重复追加

详细说明见 [HELP.md](HELP.md) 和 [PLUGIN_TEST_GUIDE.md](PLUGIN_TEST_GUIDE.md)。遇到加载问题时，先确认安装包与 OneNote 位数一致。

## 项目结构

```
src/
├── OneNoteMarkdown.AddIn/         # 插件主体
│   ├── AddIn/                     # 插件入口与功能区处理
│   ├── Features/                  # 功能命令（渲染、导入、导出等）
│   ├── Markdown/                  # Markdown 解析与渲染引擎
│   ├── OneNote/                   # OneNote API 交互
│   ├── Localization/              # 国际化（中文 + 英文）
│   ├── Rendering/                 # 公式与图表渲染
│   ├── Settings/                  # 主题与配置
│   └── UI/                        # UI 组件（设置、帮助对话框）
└── OneNoteMarkdown.Installer/     # 安装程序（Inno Setup）
```

## 技术栈

- C# / .NET Framework 4.8
- OneNote COM Interop API
- WPF / WinForms（UI 对话框）
- Inno Setup（安装程序）

Office Primary Interop Assemblies 位于 `src/OneNoteMarkdown.AddIn/ThirdParty/Office`，因此 CI 不需要预装 Microsoft Office。

## 构建与测试

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

构建完成后会生成：

- `src/OneNoteMarkdown.Installer/Output/OneNoteMarkdownSetup-1.2.1-fix-x86.exe`
- `src/OneNoteMarkdown.Installer/Output/OneNoteMarkdownSetup-1.2.1-fix-x64.exe`
- `src/OneNoteMarkdown.Installer/Output/OneNoteMarkdownSetup-1.2.1-fix-arm64.exe`

GitHub Actions 的每次构建也会上传同名的 x86/x64/Arm64 安装包。

## 许可证

MIT
