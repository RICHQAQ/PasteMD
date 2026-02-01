# PasteMD
<p align="center">
  <img src="../../assets/icons/logo.png" alt="PasteMD" width="160" height="160">
</p>

<p align="center">
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/v/release/RICHQAQ/PasteMD?sort=semver&label=Release&style=flat-square&logo=github" alt="Release">
  </a>
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/downloads/RICHQAQ/PasteMD/total?label=Downloads&style=flat-square&logo=github" alt="Downloads">
  </a>
  <a href="../../LICENSE">
    <img src="https://img.shields.io/github/license/RICHQAQ/PasteMD?style=flat-square" alt="License">
  </a>
  <img src="https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.12+">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Word%20%7C%20WPS-5e8d36?style=flat-square&logo=windows&logoColor=white" alt="Platform">
</p>

<p align="center">
  <a href="README.en.md">English</a>
  |
  <a href="../../README.md">简体中文</a> 
  |
  <a href="README.ja.md">日本語</a>
</p>

<p align="center">
<a href="https://hellogithub.com/repository/RICHQAQ/PasteMD" target="_blank"><img src="https://abroad.hellogithub.com/v1/widgets/recommend.svg?rid=7dfb1883330d441f9264d8e8945c75e2&claim_uid=RvDOqI1Satiwzh4&theme=neutral" alt="Featured｜HelloGitHub" style="width: 250px; height: 54px;" width="250" height="54" /></a>
</p>

> When writing papers or reports, do formulas copied from AI tools (like ChatGPT or DeepSeek) turn into garbled text in Word? Do Markdown tables fail to paste correctly into Excel? **PasteMD was built specifically to solve these problems.**
> 
> <img src="../../docs/gif/atri/igood.gif"
     alt="I am good"
     width="100">

A lightweight tray app:
Reads **Markdown from the clipboard**, converts it to DOCX with **Pandoc**, and inserts it at the caret position in **Word/WPS**.

**✨ Feature**: Smart Markdown table detection, one-click paste to **Excel**!

**✨ Feature**: Smart HTML rich text detection, easy one-click paste of AI replies from the web into **Word/WPS**!

**✨ New feature**: App extensions (HTML+Markdown/HTML/Markdown/LaTeX/File paste), match by app/window title (e.g., Yuque/QQ).

**✨ New feature**: Conversion enhancements: configure Pandoc Filters by conversion type; auto-fix some LaTeX syntax and standalone `$...$` formula blocks.

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


* Global hotkey (default `Ctrl+Shift+B`) to paste Markdown → DOCX with one key.
* **✨ Smart Markdown table detection** and paste into Excel.
* **✨ App extensions**: configure HTML+Markdown/HTML/Markdown/LaTeX/File paste modes per app, with window-title matching.
* **✨ Conversion enhancements**: add Pandoc Filters by conversion type, auto-fix some LaTeX syntax and standalone `$...$` formula blocks.
* Auto-detect the active app: Word or WPS.
* Smartly open the required app (Word/Excel).
* Tray menu: keep files, view logs/config, etc.
* System notifications.
* No console window, non-blocking, stable.

---

## 📊 AI Website Compatibility

The following table summarizes compatibility when copying content from major AI chat sites:

| AI Service | Copy Markdown (no formulas) | Copy Markdown (with formulas) | Copy page content (no formulas) | Copy page content (with formulas) |
|---------|:----------------------------:|:----------------------------:|:---------------------------:|:---------------------------:|
| **Kimi** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ⚠️ Formulas missing |
| **DeepSeek** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| **Tongyi Qianwen** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ⚠️ Formulas missing |
| **Doubao\*** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| **Zhipu Qingyan<br/>/ChatGLM** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| **ChatGPT** | ✅ Perfect | ⚠️ Rendered as code | ✅ Perfect | ✅ Perfect |
| **Gemini** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| **Grok** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |
| **Claude** | ✅ Perfect | ✅ Perfect | ✅ Perfect | ✅ Perfect |

**Legend:**
- ✅ **Perfect**: formatting, style, and formulas are displayed correctly
- ⚠️ **Rendered as code**: formulas appear as LaTeX code; use Word/WPS equation editor to fix
- ⚠️ **Formulas missing**: formulas are removed; re-enter them with the equation editor
- **Doubao**: before copying web content (with formulas), enable “Allow clipboard read” in the browser (via the icon left of the URL)

**Test method:**
1. **Copy Markdown**: click the “Copy” button in an AI reply (usually Markdown, some sites include HTML)
2. **Copy page content**: select the AI reply content directly and copy (HTML rich text)

---

## 🚀 Getting Started

1. Download an executable from the [Releases page](https://github.com/RICHQAQ/PasteMD/releases/):

   * ~~**PasteMD_vx.x.x.exe** — **portable build**; requires Pandoc installed and available in command line. If not installed, download from the [Pandoc website](https://pandoc.org/installing.html).~~ (No longer provided. Build from source if needed.)
   * **PasteMD_pandoc-Setup.exe** — **all-in-one installer** bundled with Pandoc.

2. Open Word, WPS, or Excel and place the caret where you want to insert.

3. Copy **Markdown** or **web content** into the clipboard, then press **Ctrl+Shift+B**.

4. The result will be inserted automatically:
   - **Markdown tables** → paste into Excel automatically (if Excel is open)
   - **Regular Markdown/web content** → convert to DOCX and insert into Word/WPS

5. A success/failure notification appears at the bottom right.

---

## ⚙️ Configuration

The first run creates `config.json` in the user data directory (Windows: `%APPDATA%\\PasteMD\\config.json`, MacOS: `~/Library/Application Support/PasteMD/config.json`). You can edit it manually:

```json
{
  "hotkey": "<ctrl>+<shift>+b",
  "pandoc_path": "pandoc",
  "reference_docx": null,
  "save_dir": "%USERPROFILE%\\Documents\\pastemd",
  "keep_file": false,
  "notify": true,
  "startup_notify": true,
  "enable_excel": true,
  "excel_keep_format": true,
  "paste_delay_s": 0.3,
  "no_app_action": "open",
  "md_disable_first_para_indent": true,
  "html_disable_first_para_indent": true,
  "html_formatting": {
    "strikethrough_to_del": true
  },
  "move_cursor_to_end": true,
  "Keep_original_formula": false,
  "enable_latex_replacements": true,
  "fix_single_dollar_block": true,
  "language": "zh-CN",
  "pandoc_request_headers": [
    "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
  ],
  "pandoc_filters": [],
  "pandoc_filters_by_conversion": {
    "md_to_docx": [],
    "html_to_docx": [],
    "html_to_md": [],
    "md_to_html": [],
    "md_to_rtf": [],
    "md_to_latex": []
  },
  "extensible_workflows": {
    "html": {
      "enabled": true,
      "apps": [],
      "keep_formula_latex": true
    },
    "md": {
      "enabled": true,
      "apps": [],
      "html_formatting": {
        "css_font_to_semantic": true,
        "bold_first_row_to_header": true
      }
    },
    "latex": {
      "enabled": true,
      "apps": []
    },
    "file": {
      "enabled": true,
      "apps": []
    }
  }
}
```

Fields:

* `hotkey`: global hotkey syntax, e.g. `<ctrl>+<alt>+v`.
* `pandoc_path`: Pandoc executable path.
* `reference_docx`: optional Pandoc reference template.
* `save_dir`: output directory when keeping files.
* `keep_file`: keep generated DOCX files or delete them.
* `notify`: show system notifications.
* `startup_notify`: show a notification on startup.
* **`enable_excel`**: enable smart Markdown table detection and paste into Excel (default true).
* **`excel_keep_format`**: keep Markdown formatting (bold/italic/code, etc.) when pasting into Excel (default true).
* `paste_delay_s`: delay in seconds before pasting (Windows may need a short delay for clipboard updates).
* **`no_app_action`**: default action when no target app (Word/Excel) is detected (default `"open"`). Options: `open`=auto open, `save`=save only, `clipboard`=copy file to clipboard, `none`=no action.
* **`md_disable_first_para_indent`**: disable special formatting for the first paragraph when converting Markdown (default true).
* **`html_formatting`**: formatting options for HTML rich text conversion.
  * **`strikethrough_to_del`**: convert ~~ strikethrough to `<del>` for correct rendering (default true).
* **`html_disable_first_para_indent`**: disable special formatting for the first paragraph when converting HTML rich text (default true).
* **`move_cursor_to_end`**: move the caret to the end after inserting (default true).
* **`Keep_original_formula`**: keep original math formulas (LaTeX code form).
* `enable_latex_replacements`: auto-fix some incompatible LaTeX syntax (e.g. replace `{\\kern 10pt}` with `\\qquad`).
* `fix_single_dollar_block`: auto-detect and fix standalone `$ ... $` formula blocks (convert to `$$ ... $$`).
* `language`: UI language. `zh-CN` (Simplified Chinese), `en-US` (English), `ja-JP` (Japanese).
* `pandoc_request_headers`: request headers for Pandoc when downloading remote resources (one `Header: Value` per line).
* **`pandoc_filters`**: custom Pandoc Filter list. Add `.lua` scripts or executable paths; filters run in list order. Extends conversion functions (custom formatting, special syntax transforms, etc.). Default empty. Example: `["%APPDATA%\\npm\\mermaid-filter.cmd"]` enables Mermaid diagrams.
* `pandoc_filters_by_conversion`: configure Filters per conversion type (e.g. `md_to_docx`, `html_to_md`, etc.).
* `extensible_workflows`: app extension settings (match by app/window title and choose paste mode). See below.

Apply changes via the tray menu **“Reload config/hotkey”**.

---

## 🔧 Advanced: Custom Pandoc Filters

### What are Pandoc Filters?

Pandoc Filters are plugin programs that process content during conversion. PasteMD supports configuring multiple filters that execute sequentially to extend functionality.

### Use Case Example: Mermaid Diagram Support

To use Mermaid diagrams in Markdown and convert them properly to Word, you can use [mermaid-filter](https://github.com/raghur/mermaid-filter).

**1. Install mermaid-filter**

```bash
npm install --global mermaid-filter
```

*Prerequisite: [Node.js](https://nodejs.org/) must be installed*

<details>
<summary>⚠️ <b>Troubleshooting: Chrome Download Failure</b></summary>

Installing mermaid-filter requires downloading Chromium. If automatic download fails, you can download it manually:

**Step 1: Find Required Chromium Version**

Check the file: `%APPDATA%\npm\node_modules\mermaid-filter\node_modules\puppeteer-core\lib\cjs\puppeteer\revisions.d.ts`

Find content like:
```typescript
chromium: "1108766";
```
Or in the error message, e.g.:
```bash
npm error Error: Download failed: server returned code 502. URL: https://npmmirror.com/mirrors/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```
Find version like `Win_x64/1108766`.

Note down this version number (e.g., `1108766`).

**Step 2: Download Chromium**

Based on the version number from Step 1, download the corresponding Chromium:

```
https://storage.googleapis.com/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```

(Replace `1108766` in the URL with your version number)

**Step 3: Extract to Designated Directory**

Extract the downloaded `chrome-win.zip` to:

```
%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win
```

(Replace `1108766` in the path with your version number)

After extraction, `chrome.exe` should be located at:  
`%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win\chrome.exe`

</details>

**2. Configure in PasteMD**

Option 1: Via Settings UI
- Open PasteMD Settings → Conversion Tab → Pandoc Filters
- Click "Add..." button
- Select filter file: `%APPDATA%\npm\mermaid-filter.cmd`
- Save settings

Option 2: Edit config file
```json
{
  "pandoc_filters": [
    "%APPDATA%\\npm\\mermaid-filter.cmd"
  ]
}
```

**3. Test It Out**

Copy the following Markdown and convert with PasteMD:

~~~markdown
```mermaid
graph LR
    A[Start] --> B[Process]
    B --> C[End]
```
~~~

The Mermaid diagram will be rendered as an image and inserted into Word.

### More Filter Resources

- [Official Pandoc Filters List](https://github.com/jgm/pandoc/wiki/Pandoc-Filters)
- [Lua Filters Documentation](https://pandoc.org/lua-filters.html)

---

## 🧩 App Extensions (Custom Paste Workflows)

In Settings → **App Extensions**, you can configure paste modes per app and match by window title (regex supported):

* **HTML** / **Markdown** / **LaTeX** / **File**: choose the most suitable paste mode for each target app
  - HTML/Markdown suit rich-text note apps (e.g., Yuque)
  - LaTeX suits academic sites like Overleaf
  - File suits apps like QQ/WeChat that accept file attachments

> Tip: for the same app, configure only one workflow to avoid conflicts; use “window title matching” when needed.

Example config (excerpt):

> On Windows, `id` is usually the app's exe path; on macOS it's the bundle id (recommended to add via the settings UI for auto-fill).

```json
{
  "extensible_workflows": {
    "html": {
      "enabled": true,
      "apps": [
        {
          "name": "Yuque",
          "id": "/path/yuque.exe",
          "window_patterns": []
        }
      ],
      "keep_formula_latex": true
    },
    "latex": {
      "enabled": true,
      "apps": [
        {
          "name": "chrome",
          "id": "/path/chrome.exe",
          "window_patterns": [
            ".*overleaf.*"
          ]
        }
      ]
    },
    "file": {
      "enabled": true,
      "apps": [
        {
          "name": "QQ",
          "id": "/path/qq.exe",
          "window_patterns": []
        }
      ]
    }
  }
}
```

---

## Tray Menu

* Quick view: current global hotkey (read-only).
* Enable hotkey: toggle global hotkey.
* Popup notifications: toggle system notifications.
* No app action: default action when Word/WPS/Excel is not detected (auto open / save only / copy to clipboard / none).
* Move cursor to end: move caret to end after insertion.
* HTML formatting: toggle **strikethrough ~~ to `<del>`** and other HTML cleanup for correct conversion.
* Set hotkey: record and save a new global hotkey in the UI (takes effect immediately).
* Keep generated files: when enabled, DOCX files are saved to `save_dir`.
* Open save directory, view logs, edit config, reload config/hotkey.
* Version: show current version; check for updates; when available, show item and open download page.
* Quit: exit the app.

---

## 📦 Run From Source / Build

Recommended: Python 3.12 (64-bit).

```bash
pip install -r requirements.txt
python main.py
```

Using PyInstaller:

```bash
pyinstaller --clean -F -w -n PasteMD
  --icon assets\icons\logo.ico
  --add-data "assets\icons;assets\icons"
  --add-data "pastemd\i18n\locales\*.json;pastemd\i18n\locales"
  --add-data "pastemd\lua;pastemd\lua"
  --hidden-import plyer.platforms.win.notification
  main.py
```

The executable will be generated at `dist/PasteMD.exe`.

---

## ⭐ Star

Thanks for every star — please share PasteMD with more users. I am aiming for 4096 stars and will keep working hard!

<img src="../../docs/gif/atri/likeyou.gif"
     alt="like you"
     width="150">

[![Star History Chart](https://api.star-history.com/svg?repos=RICHQAQ/PasteMD&type=date&legend=top-left)](https://www.star-history.com/#RICHQAQ/PasteMD&type=date&legend=top-left)

## ☕ Support & Donation

If you have ideas or suggestions, feel free to open an issue! 🤯🤯🤯


You are also welcome to join the **PasteMD user group** to chat with others:
<div align="center">
  <img src="../../docs/img/qrcode.jpg" alt="PasteMD QQ group QR code" width="200" />
  <br>
  <sub>Scan to join the PasteMD QQ group</sub>
</div>

If this tool helps you, you can buy the author a coffee ☕. Your support motivates ongoing fixes, enhancements, and more scenarios. Thanks for every bit of support!

<img src="../../docs/gif/atri/flower.gif"
     alt="flower"
     width="150">
     
| Alipay | WeChat |
| --- | --- |
| ![Alipay](../../docs/pay/Alipay.jpg) | ![WeChat](../../docs/pay/Weixinpay.png) |


---

## License

This project is released under the [MIT License](../../LICENSE).
Third-party licenses are listed in [THIRD_PARTY_NOTICES.md](../../THIRD_PARTY_NOTICES.md).
