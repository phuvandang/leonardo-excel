#!/usr/bin/env bash
# install.sh — cài skill leonardo-excel cho Claude Code trên máy này.
#   1. Kiểm tra python3
#   2. Chép skill vào ~/.claude/skills/leonardo-excel (bản cũ được đổi tên giữ lại, không xoá)
#   3. Tạo môi trường Python riêng với openpyxl và pillow
# Chạy lại lệnh này sau khi `git pull` để cập nhật.
set -euo pipefail
PKG_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEST="$HOME/.claude/skills/leonardo-excel"
VENV="$HOME/.local/share/leonardo-excel-venv"
ok() { echo "✓ $1"; }; err() { echo "✗ $1"; }

command -v python3 >/dev/null 2>&1 || { err "Chưa có python3. Trên macOS chạy: xcode-select --install"; exit 1; }
ok "python3: $(python3 --version)"

mkdir -p "$HOME/.claude/skills"
if [[ -d "$DEST" ]]; then
  BAK="$DEST.cu-$(date +%Y%m%d-%H%M%S)"
  mv "$DEST" "$BAK"
  ok "Bản cũ được giữ lại tại: $BAK (anh/chị tự xoá khi không cần)"
fi
cp -R "$PKG_DIR/skills/leonardo-excel" "$DEST"
ok "Đã chép skill vào $DEST"

if [[ ! -x "$VENV/bin/python" ]]; then
  python3 -m venv "$VENV"
fi
"$VENV/bin/pip" install -q --upgrade openpyxl pillow
ok "Môi trường Python: $VENV"

bash "$PKG_DIR/check-setup.sh"
echo ""
echo "Xong. Mở phiên Claude Code mới và thử: \"Tạo file Excel theo chuẩn Leonardo cho …\""
