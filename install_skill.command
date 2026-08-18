#!/bin/bash
# install_skill.command — ติดตั้งสกิลไปยัง Claude Code และ Codex ในครั้งเดียว
#
# วิธีใช้: ดับเบิลคลิกไฟล์นี้ใน Finder (หรือ bash install_skill.command)
#
# ปลายทางที่ติดตั้งให้:
#   ~/.claude/skills/tulibs-thesis-docx-checker   (Claude Code / Claude Desktop personal skills)
#   ~/.codex/skills/tulibs-thesis-docx-checker    (Codex)
#
# สำหรับ Claude Desktop แบบ "สกิลในบัญชี" ให้ใช้ไฟล์ tulibs-thesis-docx-checker.skill
# แล้วกดปุ่ม Save skill ในแชทแทน — วิธีนั้นไม่ต้องแตะโฟลเดอร์ในเครื่อง

set -u
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
NAME="tulibs-thesis-docx-checker"

echo "======================================================================"
echo "ติดตั้ง $NAME"
echo "ต้นทาง: $SRC"
echo "======================================================================"

if [ ! -f "$SRC/SKILL.md" ] || [ ! -f "$SRC/scripts/check_deep.py" ]; then
  echo "✗ โฟลเดอร์นี้ไม่ใช่ต้นฉบับสกิล"
  read -r -p "กด Enter เพื่อปิด" _; exit 1
fi

for base in "$HOME/.claude/skills" "$HOME/.codex/skills"; do
  dest="$base/$NAME"
  echo
  echo "→ $dest"
  mkdir -p "$dest" 2>/dev/null || { echo "  ✗ สร้างโฟลเดอร์ไม่ได้"; continue; }
  bash "$SRC/sync_skill.command" "$dest" 2>&1 | sed -n 's/^  /    /p'
done

echo
echo "----------------------------------------------------------------------"
echo "เสร็จแล้ว — ปิดแล้วเปิด Claude/Codex ใหม่หนึ่งครั้งเพื่อให้เห็นสกิลรุ่นใหม่"
echo "ทดสอบว่าใช้ได้: สั่ง 'ตรวจไฟล์วิทยานิพนธ์ <ชื่อไฟล์>.docx'"
read -r -p "กด Enter เพื่อปิด" _
