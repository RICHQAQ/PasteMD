#!/usr/bin/env bash
# build_macos.sh
# PasteMD macOS Nuitka 打包脚本（带签名 / 固定 Bundle ID / 固定安装路径）
# 目的：
# 1) 用“固定的证书身份”签名，减少每次构建都被 macOS 当成“新应用”的概率（权限更容易沿用）
# 2) 确保 CFBundleIdentifier（Bundle ID）稳定
# 3) 将 .app 复制到固定目录（/Applications），进一步提升权限复用稳定性

set -euo pipefail  # 遇到错误退出；未定义变量当错误；管道错误会传递

echo "开始构建 PasteMD (macOS)..."

############################
# 你需要修改/确认的配置项 #
############################

# 应用显示名称（.app 名称）
APP_NAME="PasteMD"

# 你的 Bundle ID（建议长期固定不变）
BUNDLE_ID="com.richqaq.pastemd"

# 你的代码签名证书身份（security find-identity -v -p codesigning 输出里的那条）
# 例：Apple Development: xxx@163.com (xxxxxxxx)
CERT_IDENTITY="${CERT_IDENTITY_PASTEMD}"

# Nuitka 输出目录
OUT_DIR="nuitka/macos"

# 入口脚本
ENTRY="PasteMD.py"

# 你希望把 app 安装到哪里（开发期建议固定路径，权限更稳）
INSTALL_DIR="${INSTALL_DIR:-/Applications}"

# 是否静默（默认 0：不静默，便于排错；设置 QUIET=1 ./build_macos.sh 可静默）
QUIET="${QUIET:-0}"

# Python 可执行文件（默认用当前环境 python；你也可 PYTHON_BIN=/path/to/python ./build_macos.sh）
PYTHON_BIN="${PYTHON_BIN:-python}"

####################
# 基础环境检查     #
####################
echo "检查 Bundle ID：$BUNDLE_ID"
if [[ -z "$BUNDLE_ID" ]]; then
  echo "错误：BUNDLE_ID 不能为空。"
  exit 1
fi
if [[ "$BUNDLE_ID" != *.* ]] || [[ ! "$BUNDLE_ID" =~ ^[A-Za-z0-9][A-Za-z0-9.-]*[A-Za-z0-9]$ ]]; then
  echo "错误：BUNDLE_ID 格式不合法，建议使用反向域名格式（例如 com.example.app）。"
  exit 1
fi

command -v "$PYTHON_BIN" >/dev/null 2>&1 || { echo "错误：找不到 python：$PYTHON_BIN"; exit 1; }
command -v codesign >/dev/null 2>&1 || { echo "错误：找不到 codesign（Xcode Command Line Tools 是否安装？）"; exit 1; }
test -x /usr/libexec/PlistBuddy || { echo "错误：找不到 /usr/libexec/PlistBuddy"; exit 1; }

# 检查 Nuitka 是否可用
"$PYTHON_BIN" -m nuitka --version >/dev/null 2>&1 || {
  echo "错误：当前 python 环境中没有 Nuitka。请先安装：pip install nuitka"
  exit 1
}

echo "使用签名身份：${CERT_IDENTITY:-<未设置>}"
if [[ -z "${CERT_IDENTITY:-}" ]]; then
  echo "警告：CERT_IDENTITY_PASTEMD 未设置，将导致 codesign 失败。"
fi

# 检查证书身份是否存在（只是提示，不强制退出）
if ! security find-identity -v -p codesigning | grep -Fq "$CERT_IDENTITY"; then
  echo "警告：在钥匙串中没找到该签名身份：$CERT_IDENTITY"
  echo "     你可以运行：security find-identity -v -p codesigning 复制正确的名称"
fi

####################
# 获取版本号       #
####################
VERSION="$("$PYTHON_BIN" -c "import sys; sys.path.insert(0, '.'); from pastemd import __version__; print(__version__)")"
echo "构建版本：$VERSION"

####################
# 清理旧构建       #
####################
echo "清理旧构建目录：$OUT_DIR"
rm -rf "$OUT_DIR"

############################
# 判断 Nuitka 是否支持某参数 #
############################
has_nuitka_opt() {
  # 用 --help 判断参数是否存在；不同 Nuitka 版本参数会变
  "$PYTHON_BIN" -m nuitka --help 2>/dev/null | grep -q -- "$1"
}

require_nuitka_opt() {
  if ! has_nuitka_opt "$1"; then
    echo "错误：当前 Nuitka 不支持 $1，请升级 Nuitka。"
    exit 1
  fi
}

####################
# 组装 Nuitka 命令 #
####################
echo "开始运行 Nuitka..."

APPLE_EVENTS_DESC="PasteMD needs permission to identify the frontmost app window (Word, WPS, etc.) so it can insert content into the correct target."
SCREEN_CAPTURE_DESC="PasteMD needs screen recording access to read window titles (macOS treats window title access as Screen Recording)."
INPUT_MONITORING_DESC="PasteMD needs permission for global hotkey listening."

require_nuitka_opt "--macos-signed-app-name"


# 用数组拼命令，避免空格/引号问题
NUITKA_CMD=(
  "$PYTHON_BIN" -m nuitka "$ENTRY"
  --standalone
  --macos-create-app-bundle
  --macos-app-name="$APP_NAME"
  --macos-app-icon=assets/icons/logo.icns
  --enable-plugin=tk-inter
  --include-package=truststore
  --output-dir="$OUT_DIR"
  --output-filename="$APP_NAME"

  # 资源打包
  --include-data-dir=assets/icons=assets/icons
  --include-data-dir=pastemd/lua=lua
  --include-data-files=pastemd/i18n/locales/*.json=i18n/locales/
  --include-data-dir=third_party/pandoc/macos=pandoc

  # 显式包含的包/模块（按你当前项目需要）
  --include-package=pync
  --include-package-data=pync
  --include-package=plyer
  --include-package=plyer.platforms.macosx
  --include-module=plyer.platforms.macosx.notification
  # --include-package=pynput
  # --include-package=pynput.keyboard
  # --include-package=pynput._util
  # --include-package=pynput._util.darwin

  # --include-package=Quartz
  # --include-package=AppKit
  # --include-package=Foundation
  # --include-package=objc
  # --include-package=Cocoa

  # --include-package=PIL
  # --include-package=tkinter
  # --include-package=pystray

  # 排除测试相关依赖
  --nofollow-import-to=pytest
  --nofollow-import-to=test
  --nofollow-import-to=tests
)

# 可选：设置版本号（若当前 Nuitka 支持）
if has_nuitka_opt "--macos-app-version"; then
  NUITKA_CMD+=( --macos-app-version="$VERSION" )
fi

# 可选：设置签名相关的“应用标识名”（不少版本用它来生成/写入 bundle id 或签名名）
# 注意：不同版本对这个参数含义略有差异，但通常我们希望它稳定为反向域名
if has_nuitka_opt "--macos-signed-app-name"; then
  NUITKA_CMD+=( --macos-signed-app-name="$BUNDLE_ID" )
fi

# 可选：让 Nuitka 自己签名（若当前 Nuitka 支持）
# 你已经确认有 Apple Development 证书身份，这里可以直接用
# if has_nuitka_opt "--macos-sign-identity"; then
#   NUITKA_CMD+=( --macos-sign-identity="$CERT_IDENTITY" )
# fi

# if has_nuitka_opt "--macos-sign-notarization" ; then
#   NUITKA_CMD+=( --macos-sign-notarization )
# fi

# 是否静默
if [[ "$QUIET" == "1" ]]; then
  NUITKA_CMD+=( --quiet )
fi

# 执行构建
"${NUITKA_CMD[@]}"

####################
# 构建产物路径     #
####################
APP_PATH="$OUT_DIR/$APP_NAME.app"
PLIST_PATH="$APP_PATH/Contents/Info.plist"

if [[ ! -d "$APP_PATH" ]]; then
  echo "错误：构建完成但未找到 app：$APP_PATH"
  exit 1
fi

echo "构建完成：$APP_PATH"

#############################################
# 写入 Info.plist 权限描述（Usage Descriptions）
#############################################
echo "写入 Info.plist 权限描述..."

/usr/libexec/PlistBuddy -c "Set :NSAppleEventsUsageDescription $APPLE_EVENTS_DESC" "$PLIST_PATH" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :NSAppleEventsUsageDescription string $APPLE_EVENTS_DESC" "$PLIST_PATH"
/usr/libexec/PlistBuddy -c "Set :NSScreenCaptureUsageDescription $SCREEN_CAPTURE_DESC" "$PLIST_PATH" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :NSScreenCaptureUsageDescription string $SCREEN_CAPTURE_DESC" "$PLIST_PATH"
/usr/libexec/PlistBuddy -c "Set :NSInputMonitoringUsageDescription $INPUT_MONITORING_DESC" "$PLIST_PATH" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :NSInputMonitoringUsageDescription string $INPUT_MONITORING_DESC" "$PLIST_PATH"

#############################################
# 复制 InfoPlist.strings 本地化文件          #
#############################################
INFO_PLIST_LPROJ_DIR="assets/macos/InfoPlist.strings"
if [[ -d "$INFO_PLIST_LPROJ_DIR" ]]; then
  echo "复制 InfoPlist.strings 本地化文件..."
  mkdir -p "$APP_PATH/Contents/Resources"
  for lproj in "$INFO_PLIST_LPROJ_DIR"/*.lproj; do
    [[ -d "$lproj" ]] || continue
    cp -R "$lproj" "$APP_PATH/Contents/Resources/"
  done
else
  echo "警告：未找到 InfoPlist.strings 目录：$INFO_PLIST_LPROJ_DIR"
fi

echo "InfoPlist.strings 本地化文件已复制。"

#############################################
# 检查 Bundle ID 是否符合预期
#############################################
echo "检查 Bundle ID..."
ACTUAL_BUNDLE_ID="$(/usr/libexec/PlistBuddy -c "Print :CFBundleIdentifier" "$PLIST_PATH" 2>/dev/null || true)"
if [[ -z "$ACTUAL_BUNDLE_ID" ]]; then
  echo "错误：未在 Info.plist 中找到 CFBundleIdentifier"
  exit 1
fi
if [[ "$ACTUAL_BUNDLE_ID" != "$BUNDLE_ID" ]]; then
  echo "错误：Bundle ID 不一致"
  echo "  期望：$BUNDLE_ID"
  echo "  实际：$ACTUAL_BUNDLE_ID"
  exit 1
fi
echo "Bundle ID 检查通过：$ACTUAL_BUNDLE_ID"
####################
# 重新签名（稳）   #
####################
# 即使 Nuitka 已签名，修 plist 后也必须重签。
# --deep：对内部嵌套框架/二进制一起签
# --force：覆盖旧签名
echo "对 app 进行 codesign 签名：$CERT_IDENTITY"
codesign --force --deep --options runtime --timestamp --sign "$CERT_IDENTITY" "$APP_PATH"

####################
# 验证签名信息     #
####################
echo "验证签名（若这里报错，说明签名不完整或有未签文件）..."
codesign --verify --deep --strict --verbose=2 "$APP_PATH" || true

echo "打印关键签名信息（Identifier/TeamIdentifier/Authority）..."
codesign -dv --verbose=4 "$APP_PATH" 2>&1 | egrep "Identifier|TeamIdentifier|Authority" || true

echo "确认 Info.plist 的 Bundle ID："
defaults read "$PLIST_PATH" CFBundleIdentifier || true

#################################
# 开发期：复制到固定路径再运行    #
#################################
# 原因：部分权限对“路径”也敏感。开发期反复构建如果每次路径/包结构变化太大，
# 系统更容易当成新应用。把 app 固定放到 ~/Applications 往往更省心。
echo "复制到固定安装目录：$INSTALL_DIR"
mkdir -p "$INSTALL_DIR"
rm -rf "$INSTALL_DIR/$APP_NAME.app"
cp -R "$APP_PATH" "$INSTALL_DIR/$APP_NAME.app"

echo "最终 App 位置：$INSTALL_DIR/$APP_NAME.app"
echo "启动应用..."
open "$INSTALL_DIR/$APP_NAME.app"

echo "全部完成 ✅"
