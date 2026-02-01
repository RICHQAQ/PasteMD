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

> 論文やレポートを書いていると、ChatGPT / DeepSeek などの AI サイトからコピーした数式が Word で文字化けしませんか？Markdown の表を Excel に貼り付けると崩れませんか？**PasteMD はその問題を解決するためのツールです。**
> 
> <img src="../../docs/gif/atri/igood.gif"
     alt="I am good"
     width="100">

常駐トレイの小さなツール：
**クリップボードの Markdown** を読み取り、**Pandoc** で DOCX に変換し、**Word/WPS** のカーソル位置に自動挿入します。

**✨ 機能**：Markdown 表をスマートに認識し、**Excel** にワンクリック貼り付け！

**✨ 機能**：HTML リッチテキストをスマートに認識し、Web 上の AI 返信を **Word/WPS** にワンクリック貼り付け！

**✨ 新機能**：アプリ拡張（HTML+Markdown/HTML/Markdown/LaTeX/ファイル貼り付け）。アプリ/ウィンドウタイトルでマッチ可能（例：語雀/QQ）。

**✨ 新機能**：変換強化：変換タイプごとの Pandoc Filters 設定、LaTeX 構文と単独の `$...$` 数式ブロックの自動修正。

---

## 機能紹介

### デモ

#### Markdown → Word/WPS

<p align="center">
  <img src="../../docs/gif/demo.gif" alt="デモ動画" width="600">
</p>

#### Web の AI 返信 → Word/WPS
<p align="center">
  <img src="../../docs/gif/demo-html.gif" alt="HTMLデモ" width="600">
</p>

#### Markdown 表 → Excel
<p align="center">
  <img src="../../docs/gif/demo-excel.gif" alt="Excelデモ" width="600">
</p>

#### 設定画面
<p align="center">
  <img src="../../docs/gif/demo-chage_format.gif" alt="設定デモ" width="600">
</p>


* グローバルホットキー（既定 `Ctrl+Shift+B`）で Markdown → DOCX を一発変換。
* **✨ Markdown 表を自動認識**し、Excel に貼り付け。
* **✨ アプリ拡張**：アプリごとに HTML+Markdown/HTML/Markdown/LaTeX/ファイル貼り付けを設定し、ウィンドウタイトルでマッチ可能。
* **✨ 変換強化**：変換タイプ別に Pandoc Filters を追加し、一部 LaTeX 構文と単独 `$...$` 数式ブロックを自動修正。
* 前面アプリが Word/WPS かを自動判定。
* 必要に応じて Word/Excel を自動起動。
* トレイメニューから、ファイル保持やログ/設定の閲覧が可能。
* システム通知に対応。
* ブラックウィンドウなし、ノンブロッキングで安定稼働。

---

## 📊 AI サイト互換性テスト

主要 AI チャットサイトでのコピー＆貼り付け互換性テスト結果：

| AI サイト | Markdown のコピー<br/>（数式なし） | Markdown のコピー<br/>（数式あり） | Web 内容のコピー<br/>（数式なし） | Web 内容のコピー<br/>（数式あり） |
|---------|:----------------------------:|:----------------------------:|:---------------------------:|:---------------------------:|
| **Kimi** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ⚠️ 数式が表示不可 |
| **DeepSeek** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 |
| **通义千问** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ⚠️ 数式が表示不可 |
| **豆包\*** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 |
| **智谱清言<br/>/ChatGLM** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 |
| **ChatGPT** | ✅ 完全対応 | ⚠️ 数式がコード表示 | ✅ 完全対応 | ✅ 完全対応 |
| **Gemini** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 |
| **Grok** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 |
| **Claude** | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 | ✅ 完全対応 |

**凡例：**
- ✅ **完全対応**：書式/スタイル/数式が正しく表示
- ⚠️ **数式がコード表示**：数式が LaTeX コードとして表示されるため、Word/WPS の数式エディタで修正が必要
- ⚠️ **数式が表示不可**：数式が失われるため、Word/WPS の数式エディタで再入力が必要
- **豆包**：Web 内容（数式あり）をコピーする前に、ブラウザで「クリップボードの読み取り許可」を有効化してください（URL 左側のアイコンから設定可能）

**テスト方法：**
1. **Markdown のコピー**：AI 返信内の「コピー」ボタンを押す（Markdown 形式。サイトによっては HTML も含む）
2. **Web 内容のコピー**：AI 返信内容を直接選択してコピー（HTML リッチテキスト）

---

## 🚀 使い方

1. 実行ファイルをダウンロード（[Releases ページ](https://github.com/RICHQAQ/PasteMD/releases/)）：

   * ~~**PasteMD_vx.x.x.exe**：**ポータブル版**。事前に **Pandoc** をインストールし、コマンドラインで実行できる必要があります。  
   未インストールの場合は [Pandoc 公式サイト](https://pandoc.org/installing.html) から導入してください。~~（提供終了。必要なら自前でビルド）
   * **PasteMD_pandoc-Setup.exe**：**統合インストーラ**。Pandoc 同梱で追加設定不要。

2. Word/WPS/Excel を開き、挿入したい位置にカーソルを置く。

3. **Markdown** または **Web 内容** をクリップボードにコピーし、**Ctrl+Shift+B** を押す。

4. 変換結果が自動で挿入されます：
   - **Markdown 表** → Excel に自動貼り付け（Excel が開いている場合）
   - **通常 Markdown/Web 内容** → DOCX に変換して Word/WPS に挿入

5. 右下に成功/失敗の通知が表示されます。

---

## ⚙️ 設定

初回実行時にユーザーデータディレクトリへ `config.json` が生成されます（Windows：`%APPDATA%\\PasteMD\\config.json`、MacOS：`~/Library/Application Support/PasteMD/config.json`）。手動で編集可能です：

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

項目説明：

* `hotkey`：グローバルホットキー。例 `<ctrl>+<alt>+v`。
* `pandoc_path`：Pandoc 実行ファイルのパス。
* `reference_docx`：Pandoc 参照テンプレート（任意）。
* `save_dir`：生成ファイルを保持する場合の保存先。
* `keep_file`：生成した DOCX を保持するか。
* `notify`：システム通知を表示するか。
* `startup_notify`：起動時に通知を表示するか。
* **`enable_excel`**：Markdown 表の自動認識と Excel 貼り付けを有効化（既定 true）。
* **`excel_keep_format`**：Excel 貼り付け時に Markdown 書式（太字/斜体/コードなど）を保持（既定 true）。
* `paste_delay_s`：貼り付け前の遅延秒数（Windows ではクリップボード更新に時間がかかる場合があります）。
* **`no_app_action`**：対象アプリ（Word/Excel など）が検出されない場合の既定動作（既定 `"open"`）。選択肢：`open`=自動で開く、`save`=保存のみ、`clipboard`=クリップボードへコピー、`none`=何もしない。
* **`md_disable_first_para_indent`**：Markdown 変換時、先頭段落の特殊書式を無効化し正文スタイルに統一（既定 true）。
* **`html_formatting`**：HTML リッチテキスト変換時の整形オプション。
  * **`strikethrough_to_del`**：取り消し線 ~~ を `<del>` に変換して正しく表示（既定 true）。
* **`html_disable_first_para_indent`**：HTML 変換時、先頭段落の特殊書式を無効化（既定 true）。
* **`move_cursor_to_end`**：挿入後にカーソルを末尾へ移動（既定 true）。
* **`Keep_original_formula`**：元の数式（LaTeX コード）を保持。
* `enable_latex_replacements`：互換性のない LaTeX 構文を自動修正（例：`{\\kern 10pt}` を `\\qquad` へ置換）。
* `fix_single_dollar_block`：単独行の `$ ... $` 数式ブロックを自動認識し `$$ ... $$` に修正。
* `language`：表示言語。`zh-CN` は簡体字中国語、`en-US` は英語、`ja-JP` は日本語。
* `pandoc_request_headers`：Pandoc がリモート資源をダウンロードする際のリクエストヘッダー（1 行 1 ヘッダー）。
* **`pandoc_filters`**：カスタム Pandoc Filter のリスト。`.lua` スクリプトや実行ファイルを指定し、順番に実行されます。高度な変換に使用。既定は空。例：`["%APPDATA%\\npm\\mermaid-filter.cmd"]` で Mermaid 図をサポート。
* `pandoc_filters_by_conversion`：変換タイプ別の Filters 設定（例：`md_to_docx`、`html_to_md` など）。
* `extensible_workflows`：アプリ拡張設定（アプリ/ウィンドウタイトルでマッチして貼り付け方式を切替）。詳細は下記。

変更後はトレイメニューの **「設定/ホットキーを再読み込み」** で即反映されます。

---

## 🔧 高度な機能：Pandoc Filters のカスタマイズ

### Pandoc Filter とは？

Pandoc Filter は、変換処理中に内容をカスタマイズするためのプラグインです。PasteMD は複数の Filter を順に適用できます。

### 例：Mermaid 図のサポート

Markdown で Mermaid 図を使いたい場合は、[mermaid-filter](https://github.com/raghur/mermaid-filter) が便利です。

**1. mermaid-filter をインストール**

```bash
npm install --global mermaid-filter
```

*前提： [Node.js](https://nodejs.org/) が必要です*

<details>
<summary>⚠️ <b>トラブルシューティング：Chrome のダウンロード失敗</b></summary>

mermaid-filter のインストール時に Chromium をダウンロードします。失敗する場合は手動でダウンロードしてください：

**手順 1：必要な Chromium バージョンの確認**

次のファイルを開きます：`%APPDATA%\npm\node_modules\mermaid-filter\node_modules\puppeteer-core\lib\cjs\puppeteer\revisions.d.ts`

以下のような記述を探します：
```typescript
chromium: "1108766";
```

あるいはエラー出力の例：
```bash
npm error Error: Download failed: server returned code 502. URL: https://npmmirror.com/mirrors/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```
`Win_x64/1108766` のような番号がバージョンです。

この番号（例：`1108766`）をメモしてください。

**手順 2：Chromium をダウンロード**

上記バージョンに合わせてダウンロード：

```
https://storage.googleapis.com/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```

（URL 内の `1108766` を置き換えてください）

**手順 3：指定ディレクトリへ解凍**

ダウンロードした `chrome-win.zip` を以下へ解凍：

```
%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win
```

（ディレクトリ内の `1108766` も置き換え）

解凍後、次の場所に `chrome.exe` があることを確認：
`%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win\chrome.exe`

</details>

**2. PasteMD に設定**

方法 1：設定画面から
- PasteMD 設定 → 変換タブ → Pandoc Filters
- 「追加...」をクリック
- Filter を選択：`%APPDATA%\npm\mermaid-filter.cmd`
- 保存

方法 2：設定ファイルの編集
```json
{
  "pandoc_filters": [
    "%APPDATA%\\npm\\mermaid-filter.cmd"
  ]
}
```

**3. 動作確認**

以下の Markdown をコピーして PasteMD で変換：

~~~markdown
```mermaid
graph LR
    A[開始] --> B[処理]
    B --> C[終了]
```
~~~

Mermaid 図が画像として Word に挿入されます。

### その他の Filter 参考

- [Pandoc Filters 公式一覧](https://github.com/jgm/pandoc/wiki/Pandoc-Filters)
- [Lua Filters ドキュメント](https://pandoc.org/lua-filters.html)

---

## 🧩 アプリ拡張（カスタム貼り付けワークフロー）

設定 → **アプリ拡張** でアプリごとの貼り付け方式を設定し、ウィンドウタイトル（正規表現）でマッチ可能です：

* **HTML** / **Markdown** / **LaTeX** / **ファイル**：対象アプリに最適な貼り付け方式を選択
  - HTML/Markdown は語雀などのリッチテキストノートに適合
  - LaTeX は Overleaf などの学術サイトに適合
  - ファイルは QQ/WeChat など添付として貼り付けるアプリに適合

> ヒント：同一アプリには 1 つのワークフローのみ推奨（競合回避）。ウィンドウ名を分けたい場合は「ウィンドウ名マッチ」を使ってください。

設定例（抜粋）：

> Windows では `id` は通常 exe パス、macOS では bundle id（設定画面で追加すると自動入力されるので推奨）。

```json
{
  "extensible_workflows": {
    "html": {
      "enabled": true,
      "apps": [
        {
          "name": "语雀",
          "id": "/path/语雀.exe",
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

## トレイメニュー

* 快捷表示：現在のグローバルホットキー（読み取り専用）。
* ホットキーを有効化：グローバルホットキーのオン/オフ。
* ポップアップ通知：システム通知のオン/オフ。
* 無アプリ時の動作：Word/WPS/Excel が検出されない場合の既定動作（自動で開く/保存のみ/クリップボードへコピー/なし）。
* 挿入後にカーソルを末尾へ移動：挿入後のカーソル移動のオン/オフ。
* HTML 整形：**取り消し線 ~~ を `<del>` に変換** など、HTML 自動整形。
* ホットキー設定：UI で新しいホットキーを録音して保存（即時反映）。
* 生成ファイルを保持：オンにすると `save_dir` へ DOCX を保存。
* 保存先フォルダを開く、ログを見る、設定を編集、設定/ホットキーを再読み込み。
* バージョン：現在のバージョン表示、更新確認、更新時はダウンロードページを開く。
* 終了：アプリを終了。

---

## 📦 ソースから実行 / パッケージ化

Python 3.12（64bit）推奨。

```bash
pip install -r requirements.txt
python main.py
```

PyInstaller の場合：

```bash
pyinstaller --clean -F -w -n PasteMD
  --icon assets\icons\logo.ico
  --add-data "assets\icons;assets\icons"
  --add-data "pastemd\i18n\locales\*.json;pastemd\i18n\locales"
  --add-data "pastemd\lua;pastemd\lua"
  --hidden-import plyer.platforms.win.notification
  main.py
```

生成された実行ファイルは `dist/PasteMD.exe` に出力されます。

---

## ⭐ Star

Star してくれたみなさん、いつもありがとうございます！もっと多くの方に届くように頑張ります。4096 stars 目標です。

<img src="../../docs/gif/atri/likeyou.gif"
     alt="likeyou"
     width="150">

[![Star History Chart](https://api.star-history.com/svg?repos=RICHQAQ/PasteMD&type=date&legend=top-left)](https://www.star-history.com/#RICHQAQ/PasteMD&type=date&legend=top-left)

## 🍵 サポート/投げ銭

ご意見やアイデアがあれば issue でぜひ教えてください！🤯🤯🤯

**PasteMD ユーザー交流グループ**にも参加歓迎：
<div align="center">
  <img src="../../docs/img/qrcode.jpg" alt="PasteMD交流群二维码" width="200" />
  <br>
  <sub>PasteMD QQ 交流群に参加</sub>
</div>

このツールが役に立てば嬉しいです。作者にコーヒーを一杯ごちそうしてもらえると、さらに開発の励みになります。ありがとうございます！

<img src="../../docs/gif/atri/flower.gif"
     alt="flower"
     width="150">
     
| 支付宝 | 微信 |
| --- | --- |
| ![支付宝打赏](../../docs/pay/Alipay.jpg) | ![微信打赏](../../docs/pay/Weixinpay.png) |


---

## License

This project is licensed under the [MIT License](../../LICENSE).
Third-party licenses are listed in [THIRD_PARTY_NOTICES.md](../../THIRD_PARTY_NOTICES.md).
