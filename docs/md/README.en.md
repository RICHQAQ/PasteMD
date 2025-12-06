# PasteMD
<p align="center">
  <img src="../../assets/icons/logo.png" alt="PasteMD" width="160" height="160">
</p>

<p align="center">
  <a href="https://github.com/RICHQAQ/PasteMD/releases"><img src="https://img.shields.io/github/v/release/RICHQAQ/PasteMD?sort=semver&label=release"></a>
  <a href="https://github.com/RICHQAQ/PasteMD/releases"><img src="https://img.shields.io/github/downloads/RICHQAQ/PasteMD/total?label=downloads"></a>
  <a href="LICENSE"><img src="https://img.shields.io/github/license/RICHQAQ/PasteMD"></a>
  <img src="https://img.shields.io/badge/python-3.12%2B-3776ab">
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20Word%20%7C%20WPS-5e8d36">
</p>

<p align="center">
  <a href="README.en.md">English</a>
  | 
  <a href="../../README.md">简体中文</a> 
</p>

PasteMD is a lightweight tray app that watches your clipboard, converts Markdown or HTML-rich text to DOCX through Pandoc, and pastes the result straight into the caret position of Word or WPS. It understands Markdown tables and can paste them directly into Excel with formatting preserved, and it recognizes HTML rich text (except math) copied from web pages.

---

## Feature Highlights

### Demo Videos

#### Markdown → Word/WPS

<p align="center">
  <img src="../../docs/gif/demo.gif" alt="Markdown to Word demo" width="600">
</p>

#### Copy AI web reply → Word/WPS
<p align="center">
  <img src="../../docs/gif/demo-html.gif" alt="HTML rich text demo" width="600">
</p>

#### Markdown tables → Excel
<p align="center">
  <img src="../../docs/gif/demo-excel.gif" alt="Markdown table to Excel demo" width="600">
</p>

#### Apply formatting presets
<p align="center">
  <img src="../../docs/gif/demo-chage_format.gif" alt="Formatting demo" width="600">
</p>

### Workflow Boosters

- Global hotkey (default `Ctrl+B`) to paste the latest Markdown/HTML clipboard snapshot as DOCX.
- Automatically recognizes Markdown tables, converts them to spreadsheets, and pastes into Excel while keeping bold/italic/code formats.
- Recognizes HTML rich text (except math) copied from web pages and converts/pastes into Word/WPS.
- Detects the foreground target app (Word, WPS, or Excel) and opens the correct program when needed.
- Tray menu for toggling features, viewing logs, reloading config, and checking for updates.
- Optional toast notifications and background logging for every conversion.

---

## AI Website Compatibility

The following table summarizes how well popular AI chat sites work with PasteMD when copying Markdown or direct HTML content.

| AI Service | Copy Markdown (no formulas) | Copy Markdown (with formulas) | Copy page content (no formulas) | Copy page content (with formulas) |
|------------|----------------------------|-------------------------------|---------------------------------|-----------------------------------|
| Kimi | ✅ Perfect | ✅ Perfect | ✅ Perfect | ⚠️ Formulas missing |
| DeepSeek | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| Tongyi Qianwen | ✅ Perfect | ✅ Perfect | ✅ Perfect | ⚠️ Formulas missing |
| Doubao* | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| ChatGLM/Zhipu | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| ChatGPT | ✅ Perfect | ⚠️ Rendered as code | ✅ Perfect | ✅ Perfect |
| Gemini | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| Grok | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| Claude | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |

_*Doubao requires granting clipboard read permissions in the browser before copying HTML content with formulas (set it via the lock icon near the URL bar)._

Legend:
- ✅ **Perfect** — formatting, styles, and formulas are kept as-is.
- ⚠️ **Rendered as code** — math formulas appear as raw LaTeX and must be rebuilt inside Word/WPS.
- ⚠️ **Formulas missing** — math formulas are removed; rebuild them manually with the equation editor.

Test description:
1. **Copy Markdown** — use the “Copy” button provided beneath most AI responses (typically Markdown, sometimes HTML).
2. **Copy page content** — manually select the AI reply and copy (HTML rich text).

---

## Getting Started

1. Download an executable from the [Releases page](https://github.com/RICHQAQ/PasteMD/releases/):
   - **PasteMD_vx.x.x.exe** — portable build, requires Pandoc to be installed and accessible from `PATH`.
   - **PasteMD_pandoc-Setup.exe** — bundled installer that ships with Pandoc and works out of the box.
2. Open Word, WPS, or Excel and place the caret where you want to paste.
3. Copy Markdown or HTML-rich text, then press the global hotkey (`Ctrl+B` by default).
4. PasteMD will:
   - Send Markdown tables to Excel (when Excel is already open).
   - Convert regular Markdown/HTML to DOCX and insert it into Word/WPS.
5. A notification in the tray (and optional toast) confirms success or failure.

---

## Configuration

The first launch creates a `config.json` file. Edit it directly, then use the tray menu item **“Reload config/hotkey”** to apply changes instantly.

```json
{
  "hotkey": "<ctrl>+b",
  "pandoc_path": "pandoc",
  "reference_docx": null,
  "save_dir": "%USERPROFILE%\\Documents\\pastemd",
  "keep_file": false,
  "notify": true,
  "enable_excel": true,
  "excel_keep_format": true,
  "auto_open_on_no_app": true,
  "md_disable_first_para_indent": true,
  "html_disable_first_para_indent": true,
  "move_cursor_to_end": true,
  "language": "zh"
}
```

Key fields:

- `hotkey` — global shortcut syntax such as `<ctrl>+<alt>+v`.
- `pandoc_path` — executable name or absolute path for Pandoc.
- `reference_docx` — optional style template consumed by Pandoc.
- `save_dir` — directory used when generated DOCX files are kept.
- `keep_file` — store converted DOCX files to disk instead of deleting them.
- `notify` — show system notifications when conversions finish.
- `enable_excel` — detect Markdown tables and paste them into Excel automatically.
- `excel_keep_format` — attempt to preserve bold/italic/code styles inside Excel.
- `auto_open_on_no_app` — auto-create a document and open it with the default handler when no target app is detected.
- `md_disable_first_para_indent` / `html_disable_first_para_indent` — normalize the first paragraph style to body text.
- `move_cursor_to_end` — move the caret to the end of the inserted result.
- `language` — UI language, `en` or `zh`.

---

## Tray Menu

- Show the current global hotkey (read-only).
- Enable/disable the hotkey.
- Toggle notifications, automatic document creation when no target app is found, and cursor movement to the end after paste.
- Enable or disable Excel-specific features and formatting preservation.
- Toggle keeping generated DOCX files.
- Open save directory, view logs, edit configuration, or reload hotkeys.
- Check for updates and view installed version.
- Quit PasteMD.

---

## Build From Source

Recommended environment: Python 3.12 (64-bit).

```bash
pip install -r requirements.txt
python main.py
```

Packaged build (PyInstaller):

```bash
pyinstaller --clean -F -w -n PasteMD ^
  --icon assets\\icons\\logo.ico ^
  --add-data "assets\\icons;assets\\icons" ^
  --add-data "pastemd\\i18n\\locales;pastemd\\i18n\\locales" ^
  --hidden-import plyer.platforms.win.notification ^
  --hidden-import pastemd.i18n.locales.zh ^
  --hidden-import pastemd.i18n.locales.en ^
  main.py
```

The compiled executable will be placed in `dist/PasteMD.exe`.

---

## ⭐ Star

Every star helps — thank you for sharing PasteMD with more users.

[![Star History Chart](https://api.star-history.com/svg?repos=RICHQAQ/PasteMD&type=Date)](https://star-history.com/#RICHQAQ/PasteMD&Date)

---

## ☕ Support & Donation

If PasteMD saves you time, consider buying the author a coffee — your support helps prioritize fixes, enhancements, and new integrations.

| Alipay | WeChat |
| --- | --- |
| ![Alipay](../../docs/pay/Alipay.jpg) | ![WeChat](../../docs/pay/Weixinpay.png) |

---

## License

This project is released under the [MIT License](LICENSE).
