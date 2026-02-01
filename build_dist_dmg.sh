#!/usr/bin/env bash
# build_dist_dmg.sh
# PasteMD 一键构建“可分发 DMG”（Nuitka 最小化打包 + Developer ID 签名 + Notarize + Staple）
#
# 目标：
# 1) 用 build_macos.sh 的“更精简”的 Nuitka 打包方式（减少显式 include-package，减小包体积）
# 2) 然后按 dist 流程：修 plist / 资源归位 / Mach-O 分层签名 / 生成 DMG / 公证 / staple
#
# 依赖：
# - Xcode Command Line Tools
# - python + nuitka:   pip install -U nuitka
# - Developer ID Application 证书已在钥匙串
# - notarytool profile 已创建：
#     xcrun notarytool store-credentials "PasteMDNotary" --apple-id "xxx" --team-id "TEAMID" --password "app-specific-password"
#
# 用法：
#   export CERT_IDENTITY_PASTEMD='Developer ID Application: Your Name (TEAMID)'
#   export NOTARY_PROFILE='PasteMDNotary'
#   ./build_dist_dmg.sh
#
# 可选：
#   STRIP_BINARIES=1 ./build_dist_dmg.sh          # 先 strip 再签名（可能减小体积；若出问题就关闭）
#   DMG_SIGN=0 ./build_dist_dmg.sh                # 不对 dmg 本体 codesign（默认签）
#   EXTRA_INCLUDES=1 ./build_dist_dmg.sh          # 额外显式 include 一些动态导入包（更稳但更大）

set -euo pipefail

echo "==> 开始构建 PasteMD（可分发 DMG）..."

############################
# 可配置项
############################
APP_NAME="PasteMD"
BUNDLE_ID="com.richqaq.pastemd"
ENTRY="PasteMD.py"
OUT_DIR="nuitka/macos"
DIST_DIR="dist_dmg"

CERT_IDENTITY="${CERT_IDENTITY_PASTEMD:-}"   # Developer ID Application: ... (TEAMID)
NOTARY_PROFILE="${NOTARY_PROFILE:-}"         # notarytool store-credentials 的 profile 名
PYTHON_BIN="${PYTHON_BIN:-python}"

# 可选开关
STRIP_BINARIES="${STRIP_BINARIES:-0}"
DMG_SIGN="${DMG_SIGN:-1}"
EXTRA_INCLUDES="${EXTRA_INCLUDES:-0}"

############################
# 小工具函数
############################
need_cmd() { command -v "$1" >/dev/null 2>&1 || { echo "错误：找不到命令：$1"; exit 1; }; }

run_noproxy() {
  env -u http_proxy -u https_proxy -u all_proxy -u HTTP_PROXY -u HTTPS_PROXY -u ALL_PROXY "$@"
}

has_nuitka_opt() {
  "$PYTHON_BIN" -m nuitka --help 2>/dev/null | grep -q -- "$1"
}

require_nuitka_opt() {
  if ! has_nuitka_opt "$1"; then
    echo "错误：当前 Nuitka 不支持 $1，请升级 Nuitka。"
    exit 1
  fi
}

############################
# 环境检查
############################
need_cmd "$PYTHON_BIN"
need_cmd codesign
need_cmd security
need_cmd xcrun
need_cmd hdiutil
need_cmd file
need_cmd xattr
need_cmd ditto
need_cmd find
need_cmd chmod
test -x /usr/libexec/PlistBuddy || { echo "错误：找不到 /usr/libexec/PlistBuddy"; exit 1; }

"$PYTHON_BIN" -m nuitka --version >/dev/null 2>&1 || {
  echo "错误：当前 python 环境中没有 Nuitka。请先安装：pip install -U nuitka"
  exit 1
}

if [[ -z "$CERT_IDENTITY" ]]; then
  echo "错误：未设置 CERT_IDENTITY_PASTEMD（Developer ID Application 证书名）"
  echo "示例：export CERT_IDENTITY_PASTEMD='Developer ID Application: Your Name (TEAMID)'"
  exit 1
fi

if [[ -z "$NOTARY_PROFILE" ]]; then
  echo "错误：未设置 NOTARY_PROFILE（notarytool store-credentials 的 profile 名）"
  echo "示例：export NOTARY_PROFILE='PasteMDNotary'"
  exit 1
fi

# 提示：证书是否存在
if ! security find-identity -v -p codesigning | grep -Fq "$CERT_IDENTITY"; then
  echo "警告：在钥匙串中没找到该签名身份：$CERT_IDENTITY"
  echo "     你可以运行：security find-identity -v -p codesigning 复制正确的名称"
fi

############################
# 版本号
############################
VERSION="$("$PYTHON_BIN" -c "import sys; sys.path.insert(0,'.'); from pastemd import __version__; print(__version__)")"
echo "==> 版本号：$VERSION"

############################
# 清理
############################
echo "==> 清理旧目录：$OUT_DIR / $DIST_DIR"
rm -rf "$OUT_DIR"
mkdir -p "$DIST_DIR"

############################
# 权限描述（TCC 提示文案）
############################
APPLE_EVENTS_DESC="PasteMD needs permission to identify the frontmost app window (Word, WPS, etc.) so it can insert content into the correct target."
SCREEN_CAPTURE_DESC="PasteMD needs Screen Recording permission to read window titles (macOS may treat window title access as Screen Recording)."
INPUT_MONITORING_DESC="PasteMD needs permission for global hotkey listening."

############################
# Nuitka 构建（尽量按 build_macos.sh 的“精简方式”）
############################
require_nuitka_opt "--macos-signed-app-name"

echo "==> Nuitka 构建 .app（精简配置）..."

NUITKA_CMD=(
  "$PYTHON_BIN" -m nuitka "$ENTRY"
  --standalone
  --macos-create-app-bundle
  --macos-app-name="$APP_NAME"
  --macos-app-icon=assets/icons/logo.icns
  --enable-plugin=tk-inter
  --output-dir="$OUT_DIR"
  --output-filename="$APP_NAME"

  # 资源打包（按你项目必需）
  --include-data-dir=assets/icons=assets/icons
  --include-data-dir=pastemd/lua=lua
  --include-data-files=pastemd/i18n/locales/*.json=i18n/locales/
  --include-data-dir=third_party/pandoc/macos=pandoc

  # 排除测试相关依赖（减小体积）
  --nofollow-import-to=pytest
  --nofollow-import-to=test
  --nofollow-import-to=tests
)

# 可选：写入版本号
if has_nuitka_opt "--macos-app-version"; then
  NUITKA_CMD+=( --macos-app-version="$VERSION" )
fi

# 让 bundle id 稳定
NUITKA_CMD+=( --macos-signed-app-name="$BUNDLE_ID" )

# 为“动态导入”提供少量必要显式 include（比你之前那一大串轻很多）
# 注：如果你某些模块确实运行时报缺失，再开 EXTRA_INCLUDES=1 或自己加
NUITKA_CMD+=(
  --include-package=pync
  --include-package-data=pync
  --include-package=plyer
  --include-package=plyer.platforms.macosx
  --include-module=plyer.platforms.macosx.notification
)

if [[ "$EXTRA_INCLUDES" == "1" ]]; then
  echo "==> EXTRA_INCLUDES=1：增加显式包含（更稳但更大）..."
  NUITKA_CMD+=(
    --include-package=pystray
    --include-package=pynput
    --include-package=pynput.keyboard
    --include-package=pynput._util
    --include-package=pynput._util.darwin
    --include-package=PIL
    --include-package=tkinter
    --include-package=Quartz
    --include-package=AppKit
    --include-package=Foundation
    --include-package=objc
    --include-package=Cocoa
  )
fi

"${NUITKA_CMD[@]}"

APP_PATH="$OUT_DIR/$APP_NAME.app"
PLIST_PATH="$APP_PATH/Contents/Info.plist"
[[ -d "$APP_PATH" ]] || { echo "错误：未找到构建产物：$APP_PATH"; exit 1; }

############################
# 修正 Info.plist（Bundle ID + UsageDescription + Icon）
############################
echo "==> 修正 Info.plist（Bundle ID）..."
CURRENT_ID="$(/usr/libexec/PlistBuddy -c "Print :CFBundleIdentifier" "$PLIST_PATH" 2>/dev/null || true)"
if [[ "$CURRENT_ID" != "$BUNDLE_ID" ]]; then
  /usr/libexec/PlistBuddy -c "Set :CFBundleIdentifier $BUNDLE_ID" "$PLIST_PATH" 2>/dev/null \
    || /usr/libexec/PlistBuddy -c "Add :CFBundleIdentifier string $BUNDLE_ID" "$PLIST_PATH"
fi

echo "==> 写入 Info.plist 权限描述..."
/usr/libexec/PlistBuddy -c "Set :NSAppleEventsUsageDescription $APPLE_EVENTS_DESC" "$PLIST_PATH" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :NSAppleEventsUsageDescription string $APPLE_EVENTS_DESC" "$PLIST_PATH"
/usr/libexec/PlistBuddy -c "Set :NSScreenCaptureUsageDescription $SCREEN_CAPTURE_DESC" "$PLIST_PATH" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :NSScreenCaptureUsageDescription string $SCREEN_CAPTURE_DESC" "$PLIST_PATH"
/usr/libexec/PlistBuddy -c "Set :NSInputMonitoringUsageDescription $INPUT_MONITORING_DESC" "$PLIST_PATH" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :NSInputMonitoringUsageDescription string $INPUT_MONITORING_DESC" "$PLIST_PATH"

INFO_PLIST_LPROJ_DIR="assets/macos/InfoPlist.strings"
if [[ -d "$INFO_PLIST_LPROJ_DIR" ]]; then
  echo "==> 复制 InfoPlist.strings 本地化文件..."
  mkdir -p "$APP_PATH/Contents/Resources"
  for lproj in "$INFO_PLIST_LPROJ_DIR"/*.lproj; do
    [[ -d "$lproj" ]] || continue
    cp -R "$lproj" "$APP_PATH/Contents/Resources/"
  done
else
  echo "警告：未找到 InfoPlist.strings 目录：$INFO_PLIST_LPROJ_DIR"
fi

echo "==> 确保应用图标写入 Resources..."
ICON_SRC="assets/icons/logo.icns"
ICON_DST="$APP_PATH/Contents/Resources/AppIcon.icns"
if [[ -f "$ICON_SRC" ]]; then
  mkdir -p "$(dirname "$ICON_DST")"
  cp -f "$ICON_SRC" "$ICON_DST"
  /usr/libexec/PlistBuddy -c "Set :CFBundleIconFile AppIcon.icns" "$PLIST_PATH" 2>/dev/null \
    || /usr/libexec/PlistBuddy -c "Add :CFBundleIconFile string AppIcon.icns" "$PLIST_PATH"
  /usr/libexec/PlistBuddy -c "Set :CFBundleIconName AppIcon" "$PLIST_PATH" 2>/dev/null \
    || /usr/libexec/PlistBuddy -c "Add :CFBundleIconName string AppIcon" "$PLIST_PATH"
else
  echo "警告：找不到图标文件：$ICON_SRC"
fi

############################
# 清理 quarantine
############################
echo "==> 清理 quarantine ..."
xattr -cr "$APP_PATH" || true

############################
# 资源归位：Resources + MacOS 相对 symlink（保持旧 resource_path 兼容）
############################
MACOS_DIR="$APP_PATH/Contents/MacOS"
RES_DIR="$APP_PATH/Contents/Resources"
mkdir -p "$RES_DIR"

move_dir_and_link_back() {
  local rel="$1"   # e.g. lua / i18n / assets / pandoc/share
  local src="$MACOS_DIR/$rel"
  local dst="$RES_DIR/$rel"

  if [[ -e "$src" && ! -L "$src" ]]; then
    echo "==> 迁移并回链：$rel"

    mkdir -p "$(dirname "$dst")"
    rm -rf "$dst" 2>/dev/null || true
    mv "$src" "$dst"

    mkdir -p "$(dirname "$src")"
    local src_parent
    src_parent="$(dirname "$src")"
    local rel_target
    rel_target="$("$PYTHON_BIN" - <<PY
import os
print(os.path.relpath(os.path.realpath("$dst"), os.path.realpath("$src_parent")))
PY
)"
    ln -s "$rel_target" "$src"
  fi
}

# 你的纯资源目录（保持旧路径：MacOS/<rel>）
move_dir_and_link_back "assets"
move_dir_and_link_back "lua"
move_dir_and_link_back "i18n"
move_dir_and_link_back "tcl-files"
move_dir_and_link_back "tk-files"
move_dir_and_link_back "docx/templates"

# Third-party notices: keep inside app bundle for compliance visibility
NOTICES_SRC="THIRD_PARTY_NOTICES.md"
NOTICES_DST="$RES_DIR/THIRD_PARTY_NOTICES.md"
if [[ -f "$NOTICES_SRC" ]]; then
  echo "==> 复制 Third-Party Notices 到 Resources..."
  cp -f "$NOTICES_SRC" "$NOTICES_DST"
else
  echo "警告：未找到 $NOTICES_SRC，跳过第三方声明文件打包。"
fi

# Pandoc：只把 share 挪走（man 在里面），bin 保持在 MacOS（更符合预期）
if [[ -d "$MACOS_DIR/pandoc/share" && ! -L "$MACOS_DIR/pandoc/share" ]]; then
  move_dir_and_link_back "pandoc/share"
fi

# 保险：去掉 MacOS 下非 Mach-O 文件的可执行位
echo "==> 去掉 Contents/MacOS 下非 Mach-O 文件的可执行位（保险）..."
if [[ -d "$MACOS_DIR" ]]; then
  while IFS= read -r -d '' f; do
    if [[ -x "$f" ]] && ! file "$f" | grep -q "Mach-O"; then
      chmod a-x "$f" || true
    fi
  done < <(find "$MACOS_DIR" -type f -perm -111 -print0)
fi

############################
# 可选：strip（必须在签名前）
############################
if [[ "$STRIP_BINARIES" == "1" ]]; then
  if command -v strip >/dev/null 2>&1; then
    echo "==> STRIP_BINARIES=1：strip Mach-O（签名前）..."
    while IFS= read -r -d '' f; do
      if file "$f" | grep -q "Mach-O"; then
        # -S: strip debug symbols（相对安全）
        strip -S "$f" 2>/dev/null || true
      fi
    done < <(find "$APP_PATH/Contents" -type f -print0)
  else
    echo "警告：找不到 strip，跳过 STRIP_BINARIES。"
  fi
fi

############################
# Developer ID 签名：先签 Mach-O，再签整个 .app（签名不使用 --deep）
############################
echo "==> Developer ID 签名：先签所有 Mach-O，再签整个 .app（签名不使用 --deep）..."
sign_one() { run_noproxy codesign --force --options runtime --timestamp --sign "$CERT_IDENTITY" "$1"; }

# 1) 签所有 Mach-O（.so/.dylib/可执行文件等）
while IFS= read -r -d '' f; do
  if file "$f" | grep -q "Mach-O"; then
    sign_one "$f"
  fi
done < <(find "$APP_PATH/Contents" -type f -print0)

# 2) 最后签 app 本体
sign_one "$APP_PATH"

echo "==> 校验签名（verify 才用 --deep）..."
run_noproxy codesign --verify --deep --strict --verbose=4 "$APP_PATH"

############################
# 生成 DMG
############################
echo "==> 生成 DMG ..."
STAGE_DIR="$(mktemp -d)"
trap 'rm -rf "$STAGE_DIR"' EXIT

# 用 ditto 保留 symlink/元数据，避免签名失效
ditto "$APP_PATH" "$STAGE_DIR/$APP_NAME.app"
ln -s /Applications "$STAGE_DIR/Applications" || true

# 额外在 DMG 根目录放一份第三方声明
if [[ -f "$NOTICES_SRC" ]]; then
  cp -f "$NOTICES_SRC" "$STAGE_DIR/THIRD_PARTY_NOTICES.md"
fi

DMG_PATH="$DIST_DIR/${APP_NAME}-${VERSION}.dmg"
rm -f "$DMG_PATH"

hdiutil create \
  -volname "$APP_NAME" \
  -srcfolder "$STAGE_DIR" \
  -format UDZO \
  -imagekey zlib-level=9 \
  "$DMG_PATH"

if [[ "$DMG_SIGN" == "1" ]]; then
  echo "==> （可选）签名 DMG ..."
  run_noproxy codesign --force --timestamp --sign "$CERT_IDENTITY" "$DMG_PATH" || true
else
  echo "==> DMG_SIGN=0：跳过 DMG codesign"
fi

############################
# 公证 + Staple
############################
echo "==> 提交公证（notarytool submit --wait）..."
run_noproxy xcrun notarytool submit "$DMG_PATH" \
  --keychain-profile "$NOTARY_PROFILE" \
  --wait

echo "==> stapler staple/validate ..."
xcrun stapler staple "$DMG_PATH"
xcrun stapler validate "$DMG_PATH"

echo ""
echo "✅ 全部完成！可分发 DMG："
echo "   $DMG_PATH"
