#!/usr/bin/env bash
# install.sh — cài skill leonardo-office (chuẩn trình bày Excel và Word của Leonardo) cho Claude Code trên máy này.
#   1. Kiểm tra python3
#   2. Chép skill vào ~/.claude/skills/leonardo-office (bản cũ được chuyển sang ~/.claude/skills-backup, không xoá)
#   3. Tạo môi trường Python riêng với openpyxl, pillow và python-docx
# Chạy lại lệnh này sau khi `git pull` để cập nhật.
set -euo pipefail
PKG_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEST="$HOME/.claude/skills/leonardo-office"
VENV="$HOME/.local/share/leonardo-office-venv"
ok() { echo "✓ $1"; }; err() { echo "✗ $1"; }

command -v python3 >/dev/null 2>&1 || { err "Chưa có python3. Trên macOS chạy: xcode-select --install"; exit 1; }
ok "python3: $(python3 --version)"

mkdir -p "$HOME/.claude/skills" "$HOME/.claude/skills-backup"
STAMP="$(date +%Y%m%d-%H%M%S)"
# leonardo-excel và leonardo-word là tên cũ của skill này: chuyển đi để không còn skill trùng
for NAME in leonardo-office leonardo-excel leonardo-word; do
  D="$HOME/.claude/skills/$NAME"
  if [[ -d "$D" ]]; then
    BAK="$HOME/.claude/skills-backup/$NAME.cu-$STAMP"   # để ngoài thư mục skills, tránh bị nạp thành skill trùng
    mv "$D" "$BAK"
    ok "Bản cũ được giữ lại tại: $BAK (anh/chị tự xoá khi không cần)"
  fi
done
cp -R "$PKG_DIR/skills/leonardo-office" "$DEST"
ok "Đã chép skill vào $DEST"

if [[ ! -x "$VENV/bin/python" ]]; then
  python3 -m venv "$VENV"
fi
"$VENV/bin/pip" install -q --upgrade openpyxl pillow python-docx
ok "Môi trường Python: $VENV"

bash "$PKG_DIR/check-setup.sh"
echo ""
echo "Xong. Mở phiên Claude Code mới và thử: \"Tạo file Excel theo chuẩn Leonardo cho …\" hoặc \"Soạn file Word theo chuẩn Leonardo: …\""
