#!/bin/bash
# 打包 BulkRename-mac 为 macOS .app（PyInstaller），并清理不需要的文件控制体积。
# 用法：bash build_app.sh
set -euo pipefail
cd "$(dirname "$0")"

PY="${PY:-/opt/miniconda3/envs/iceberg/bin/python}"
APP_NAME="批量重命名工具"

# 控制体积：排除与本项目无关、却被连带分析进来的包。
# 已验证：屏蔽它们后 openpyxl 仍正常读写名单（lxml/PIL 是 openpyxl 的可选依赖，
# 会自动回退到标准库 ElementTree；numpy/charset_normalizer/tkinter 全程未使用）。
"$PY" -m PyInstaller --windowed --name "$APP_NAME" --clean --noconfirm \
    --exclude-module numpy \
    --exclude-module PIL \
    --exclude-module lxml \
    --exclude-module tkinter \
    --exclude-module charset_normalizer \
    main.py

# macOS 上 PyInstaller 会同时产出 onedir 目录（dist/批量重命名工具/），
# .app 已自包含，该目录是冗余的，删掉
rm -rf "dist/$APP_NAME"

APP="dist/$APP_NAME.app"
if [ ! -d "$APP" ]; then
    echo "打包失败：未找到 $APP" >&2
    exit 1
fi

# 控制体积：删除 Qt 自带语言文件（*.qm）。程序界面文案内置在代码里，
# 且 Qt 对话框按钮在无翻译时本就显示英文，删除后功能不受影响。
QM_SIZE=$(find "$APP" -name '*.qm' -print0 2>/dev/null | xargs -0 du -ck 2>/dev/null | tail -1 | cut -f1)
find "$APP" -name '*.qm' -delete
echo "已清理 Qt 语言文件：${QM_SIZE:-0} KB"

# 清理后重新 ad-hoc 签名，保证本机可正常打开
codesign --force --deep --sign - "$APP" >/dev/null 2>&1 || true

# 预设名单复制到 .app 旁边（打包版的「使用预设名单」在该位置查找）
cp -f 名单.xlsx dist/ 2>/dev/null || true

echo "打包完成：${APP}（$(du -sh "$APP" | cut -f1)）"
