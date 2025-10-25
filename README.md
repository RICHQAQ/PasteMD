# PasteMD
<p align="center">
  <img src="assets/icons/logo.png" alt="PasteMD" width="160" height="160">
</p>

一个常驻托盘的小工具：
从 **剪贴板读取 Markdown**，调用 **Pandoc** 转换为 DOCX，并自动插入到 **Word/WPS** 光标位置。

**✨ 新功能**：智能识别 Markdown 表格，一键粘贴到 **Excel**！

---

## 功能特点

### 演示效果

<p align="center">
  <img src="docs/demo.gif" alt="演示动图" width="480">
</p>

* 全局热键（默认 `Ctrl+B`）一键粘贴 Markdown → DOCX。
* **✨ 智能识别 Markdown 表格**，自动粘贴到 Excel。
* 自动识别当前前台应用：Word 或 WPS。
* 智能打开所需应用为Word/Excel。
* 托盘菜单，可保留文件、查看日志/配置等。
* 支持系统通知提醒。
* 无黑框，无阻塞，稳定运行。

---

## 🚀使用方法

1. 下载可执行文件（[Releases 页面](https://github.com/RICHQAQ/PasteMD/releases/)）：

   * **PasteMD\_vx.x.x.exe**：**便携版**，需要你本机已经安装好 **Pandoc** 并能在命令行运行。
   若未安装，请到 [Pandoc 官网](https://pandoc.org/installing.html) 下载安装即可。
   * **PasteMD\_pandoc-Setup.exe**：**一体化安装包**，自带 Pandoc，不需要另外配置环境。

2. 打开 Word、WPS 或 Excel，光标放在需要插入的位置。

3. 复制 Markdown 到剪贴板，按下热键 **Ctrl+B**。

4. 转换结果会自动插入到文档中：
   - **Markdown 表格** → 自动粘贴到 Excel（如果 Excel 已打开）
   - **普通 Markdown** → 转换为 DOCX 并插入 Word/WPS

5. 右下角会提示成功/失败。

---

## ⚙️配置

首次运行会生成 `config.json`，可手动编辑：

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
  "auto_open_on_no_app": true
}
```

字段说明：

* `hotkey`：全局热键，语法如 `<ctrl>+<alt>+v`。
* `pandoc_path`：Pandoc 可执行文件路径。
* `reference_docx`：Pandoc 参考模板（可选）。
* `save_dir`：保留文件时的保存目录。
* `keep_file`：是否保留生成的 DOCX 文件。
* `notify`：是否显示系统通知。
* **`enable_excel`**：**✨ 新功能** - 是否启用智能识别 Markdown 表格并粘贴到 Excel（默认 true）。
* **`excel_keep_format`**：**✨ 新功能** - Excel 粘贴时是否保留 Markdown 格式（粗体、斜体、代码等），默认 true。
* **`auto_open_on_no_app`**：**✨ 新功能** 当未检测到目标应用（如 Word/Excel）时，是否自动创建文件并用系统默认应用打开（默认 true）。

修改后可在托盘菜单选择 **“重载配置/热键”** 立即生效。

---

## 托盘菜单

* 快捷显示：当前全局热键（只读）。
* 启用热键：开/关全局热键。
* 弹窗通知：开/关系统通知。
* 无应用时自动打开：当未检测到 Word/Excel 时是否自动创建并用默认应用打开。
* 设置热键：通过图形界面录制并保存新的全局热键（即时生效）。
* 保留生成文件：勾选后生成的 DOCX 会保存在 `save_dir`。
* 启动插入 Excel：启用/禁用 Markdown 表格智能识别并粘贴至 Excel。
* 启动 Excel 解析特殊格式：粘贴到 Excel 时尽量保留粗体、斜体、代码等格式。
* 打开保存目录、查看日志、编辑配置、重载配置/热键。
* 版本：显示当前版本；可检查更新；若检测到新版本，会显示条目并可点击打开下载页面。
* 退出：退出程序。

---

## 📦从源码运行 / 打包

建议 Python 3.12 (64位)。

```bash
pip install -r requirements.txt
python main.py
```

使用 PyInstaller：

```bash
pyinstaller --clean -F -w -n PasteMD  --icon assets\icons\logo.ico  --add-data "assets\icons;assets\icons" --hidden-import plyer.platforms.win.notification  main.py
```

生成的程序在 `dist/PasteMD.exe`。

---

## 🍵支持与打赏

如果有什么想法和好建议，欢迎issue交流！🤯🤯🤯

希望这个小工具对你有帮助，欢迎请作者👻喝杯咖啡☕～你的支持会让我更有动力持续修复问题、完善功能、适配更多场景并保持长期维护。感谢每一份支持！

| 支付宝 | 微信 |
| --- | --- |
| ![支付宝打赏](docs/pay/Alipay.jpg) | ![微信打赏](docs/pay/Weixinpay.png) |


---

## License

This project is licensed under the [MIT License](LICENSE).
