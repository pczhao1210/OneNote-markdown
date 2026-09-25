**English** | [中文](README.zh-CN.md)

# OneNote Markdown

A Microsoft OneNote add-in for writing and rendering Markdown directly in your notes. Supports managed side-by-side previews, LaTeX formulas, offline Mermaid flowcharts, tables, images, links, and syntax highlighting.

![OneNote Markdown Demo](docs/demo.png)

## Download

Download from the [Releases](https://github.com/oldding/OneNote-markdown/releases) page:

- `OneNoteMarkdownSetup-1.2.1-fix-x86.exe` — for 32-bit OneNote
- `OneNoteMarkdownSetup-1.2.1-fix-x64.exe` — for x64 OneNote
- `OneNoteMarkdownSetup-1.2.1-fix-arm64.exe` — for native Arm64 OneNote on Windows 11; requires .NET Framework 4.8.1

> The installer must match **OneNote's bitness**, not Windows'. A 64-bit Windows may still run 32-bit OneNote.

## Features

| Feature | Description |
|---------|-------------|
| **Managed Preview** | Update full-page, selection, or current-text-box previews in place, beside or below their source |
| **Refresh Control** | Manual refresh by default, optional debounced refresh, conflict detection, layout reset, deletion, and source-region navigation |
| **LaTeX Formulas** | Block math rendering with bounded local caching and source-preserving error fallback |
| **Mermaid Flowcharts** | Render common Mermaid flowcharts locally without sending note content to an online service |
| **Code Highlighting** | Syntax highlighting for multiple languages |
| **Import/Export** | Preserve imported Markdown, relative local images, and avoid exporting managed previews twice |
| **Clipboard Support** | One-click copy page as Markdown |
| **Settings Dialog** | Configure presets, fonts, preview title/layout, refresh delay, images, and UI language |

## Shortcuts

| Key | Action |
|-----|--------|
| `F5` | Render page |
| `F8` | Copy Markdown to clipboard |
| `Ctrl+\` | Toggle live preview |

## Getting Started

1. Check OneNote bitness: `File → Account → About OneNote`
2. Install the matching version (`x86`, `x64`, or native `arm64`) for OneNote
3. After installation, a "Markdown" tab appears in the OneNote ribbon
4. Write Markdown text on a page, then click "Render Page", "Render Text Box", or "Render Selection"
5. Re-running a command updates its linked preview instead of appending another copy

For detailed usage, see [HELP.md](HELP.md). For plugin testing guide, see [PLUGIN_TEST_GUIDE.md](PLUGIN_TEST_GUIDE.md).

## Project Structure

```
src/
├── OneNoteMarkdown.AddIn/         # Add-in core
│   ├── AddIn/                     # Entry point & ribbon handling
│   ├── Features/                  # Commands (render, import, export, etc.)
│   ├── Markdown/                  # Markdown parser & renderer
│   ├── OneNote/                   # OneNote API interaction
│   ├── Localization/              # i18n (Chinese + English)
│   ├── Rendering/                 # Formula & diagram rendering
│   ├── Settings/                  # Theme & configuration
│   └── UI/                        # UI components (Settings, Help dialogs)
└── OneNoteMarkdown.Installer/     # Inno Setup installer
```

## Tech Stack

- C# / .NET Framework 4.8
- OneNote COM Interop API
- WPF / WinForms (UI dialogs)
- Inno Setup (installer)

Office Primary Interop Assemblies are bundled at `src/OneNoteMarkdown.AddIn/ThirdParty/Office`, so CI does not require Microsoft Office pre-installed.

## Build & Test

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

Build output:

- `src/OneNoteMarkdown.Installer/Output/OneNoteMarkdownSetup-1.2.1-fix-x86.exe`
- `src/OneNoteMarkdown.Installer/Output/OneNoteMarkdownSetup-1.2.1-fix-x64.exe`
- `src/OneNoteMarkdown.Installer/Output/OneNoteMarkdownSetup-1.2.1-fix-arm64.exe`

GitHub Actions builds also upload both installers as artifacts.

## License

MIT
