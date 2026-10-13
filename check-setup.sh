#!/usr/bin/env bash
# check-setup.sh — kiểm tra skill leonardo-excel đã sẵn sàng. Mã thoát 0 nếu ổn, 1 nếu thiếu.
set -u
DEST="$HOME/.claude/skills/leonardo-excel"; VENV="$HOME/.local/share/leonardo-excel-venv"; FAIL=0
ok() { echo "  ✓ $1"; }; err() { echo "  ✗ $1"; FAIL=1; }
echo "Kiểm tra skill leonardo-excel và leonardo-word"
[[ -f "$DEST/SKILL.md" ]] && ok "SKILL.md" || err "Chưa có $DEST/SKILL.md — chạy: bash install.sh"
[[ -f "$DEST/assets/logo-forest-480.png" ]] && ok "Logo" || err "Thiếu logo trong $DEST/assets"
[[ -f "$DEST/config.json" ]] && ok "config.json" || err "Thiếu config.json"
[[ -f "$DEST/rules.json" ]] && ok "rules.json" || err "Thiếu rules.json"
[[ -f "$HOME/.claude/skills/leonardo-word/SKILL.md" ]] && ok "Skill leonardo-word" || err "Chưa có skill leonardo-word — chạy: bash install.sh"
if [[ -x "$VENV/bin/python" ]] && "$VENV/bin/python" -c "import openpyxl, PIL, docx" 2>/dev/null; then ok "Python + openpyxl + pillow + python-docx"; else err "Thiếu môi trường Python — chạy: bash install.sh"; fi
command -v qlmanage >/dev/null 2>&1 && ok "Quick Look (xem trước ảnh)" || echo "  ⚠ Không có Quick Look: bước xem trước ảnh sẽ bị bỏ qua (chỉ có trên macOS)"
exit $FAIL
