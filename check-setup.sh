#!/usr/bin/env bash
# check-setup.sh — kiểm tra skill leonardo-office đã sẵn sàng. Mã thoát 0 nếu ổn, 1 nếu thiếu.
set -u
DEST="$HOME/.claude/skills/leonardo-office"; VENV="$HOME/.local/share/leonardo-office-venv"; FAIL=0
ok() { echo "  ✓ $1"; }; err() { echo "  ✗ $1"; FAIL=1; }
echo "Kiểm tra skill leonardo-office"
[[ -f "$DEST/SKILL.md" ]] && ok "SKILL.md" || err "Chưa có $DEST/SKILL.md — chạy: bash install.sh"
[[ -f "$DEST/excel/GUIDE.md" ]] && ok "Phần Excel" || err "Thiếu $DEST/excel/GUIDE.md"
[[ -f "$DEST/word/GUIDE.md" ]] && ok "Phần Word" || err "Thiếu $DEST/word/GUIDE.md"
[[ -f "$DEST/excel/assets/logo-forest-480.png" ]] && ok "Logo" || err "Thiếu logo trong $DEST/excel/assets"
[[ -f "$DEST/excel/config.json" ]] && ok "config.json" || err "Thiếu config.json"
[[ -f "$DEST/excel/rules.json" ]] && ok "rules.json" || err "Thiếu rules.json"
for OLD in leonardo-excel leonardo-word; do
  [[ -d "$HOME/.claude/skills/$OLD" ]] && err "Còn skill tên cũ $OLD trong ~/.claude/skills — chạy lại: bash install.sh"
done
if [[ -x "$VENV/bin/python" ]] && "$VENV/bin/python" -c "import openpyxl, PIL, docx" 2>/dev/null; then ok "Python + openpyxl + pillow + python-docx"; else err "Thiếu môi trường Python — chạy: bash install.sh"; fi
command -v qlmanage >/dev/null 2>&1 && ok "Quick Look (xem trước ảnh)" || echo "  ⚠ Không có Quick Look: bước xem trước ảnh sẽ bị bỏ qua (chỉ có trên macOS)"
exit $FAIL
