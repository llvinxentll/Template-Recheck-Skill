#!/bin/bash
# sync_skill.command — คัดลอกสกิลรุ่นล่าสุดไปทับสำเนาที่ Codex/Claude ใช้อยู่
#
# วิธีใช้: ดับเบิลคลิกไฟล์นี้ใน Finder (หรือรัน bash sync_skill.command)
#   - ไม่ใส่อะไร  → ค้นหาสำเนาในเครื่องให้เอง แล้วถามก่อนทับ
#   - ใส่ path    → คัดลอกไปที่นั้นเลย  เช่น  bash sync_skill.command ~/.codex/skills/tulibs
#
# ต้นทางคือโฟลเดอร์ที่ไฟล์นี้อยู่ (ที่เราแก้โค้ดกัน)

set -u
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SRC" || exit 1

echo "======================================================================"
echo "TULIBS Thesis Checker — sync สกิลไปยังสำเนาที่ใช้งานจริง"
echo "ต้นทาง: $SRC"
echo "======================================================================"

if [ ! -f "$SRC/SKILL.md" ] || [ ! -f "$SRC/scripts/check_docx.py" ]; then
  echo "✗ โฟลเดอร์นี้ไม่ใช่ต้นฉบับสกิล (ไม่พบ SKILL.md หรือ scripts/check_docx.py)"
  read -r -p "กด Enter เพื่อปิด" _; exit 1
fi

copy_to() {
  local dest="$1"
  mkdir -p "$dest" || return 1
  if command -v rsync >/dev/null 2>&1; then
    # กันไฟล์วิทยานิพนธ์จริงและรายงานผลตรวจไม่ให้ตามไปอยู่ในโฟลเดอร์สกิล —
    # เป็นข้อมูลนักศึกษา (ชื่อ-นามสกุล รหัส เนื้อหาที่ยังไม่เผยแพร่) และหนักเป็นสิบ MB
    # สกิลต้องการแค่ SKILL.md + scripts + references เท่านั้น
    rsync -a --delete \
      --exclude '.DS_Store' --exclude '__pycache__' --exclude 'work/' \
      --exclude '*.pyc' --exclude 'sync_skill.command' \
      --exclude '.git/' --exclude 'Template-Recheck-Skill*/' \
      --exclude 'ไฟล์สำหรับทดสอบ/' --exclude 'ผลทดสอบ*/' \
      --exclude 'อ้างอิงคำเรียกประเภทงาน/' --exclude 'ตัวอย่างการเขียนและวางเนื้อหา/' \
      --exclude '*.docx' \
      "$SRC"/ "$dest"/
  else
    echo "  ⚠ ไม่มี rsync — คัดลอกทั้งโฟลเดอร์ ให้ลบโฟลเดอร์ไฟล์นักศึกษาที่ปลายทางเอง"
    cp -R "$SRC"/. "$dest"/
    rm -rf "$dest/ไฟล์สำหรับทดสอบ" "$dest"/ผลทดสอบ* "$dest/.git" \
           "$dest"/Template-Recheck-Skill* "$dest/อ้างอิงคำเรียกประเภทงาน" 2>/dev/null
  fi
}

verify() {
  local dest="$1" ok=1
  grep -q 'def humanize' "$dest/scripts/make_report.py" 2>/dev/null \
    || { echo "  ✗ make_report.py ยังไม่มีชั้นแปลผลเป็นภาษาที่แก้ตามได้ (รุ่นเก่า)"; ok=0; }
  grep -q 'def search_text' "$dest/scripts/make_report.py" 2>/dev/null \
    || { echo "  ✗ make_report.py ยังระบุจุดด้วยเลขแผ่นแทนคำค้น (รุ่นเก่า)"; ok=0; }
  grep -q 'def find_wrong_word' "$dest/scripts/make_report.py" 2>/dev/null \
    || { echo "  ✗ make_report.py ยังไม่ระบุคำผิด/ตำแหน่งช่องว่างซ้อนแบบเจาะจง (รุ่นเก่า)"; ok=0; }
  grep -q 'def uniquify_search' "$dest/scripts/make_report.py" 2>/dev/null \
    || { echo "  ✗ make_report.py ยังไม่ขยายคำค้นให้ไม่ซ้ำ/บอกจุดที่เท่าไร (รุ่นเก่า)"; ok=0; }
  grep -q 'engine="markers"' "$dest/scripts/check_docx.py" 2>/dev/null \
    || { echo "  ✗ check_docx.py ไม่ใช่รุ่นที่ใช้ markers เป็นค่าเริ่มต้น"; ok=0; }
  grep -q 'ไม่ต้อง render อะไรทั้งสิ้น' "$dest/SKILL.md" 2>/dev/null \
    || { echo "  ✗ SKILL.md ไม่ใช่รุ่นล่าสุด"; ok=0; }
  grep -q 'def display_location' "$dest/scripts/make_report.py" 2>/dev/null \
    || { echo "  ✗ make_report.py ยังแสดงเลขย่อหน้าในรายงาน (รุ่นเก่า)"; ok=0; }
  grep -q 'agent_plan' "$dest/scripts/split_docx.py" 2>/dev/null \
    || { echo "  ✗ split_docx.py ยังไม่มีแผน fan-out agent (รุ่นเก่า)"; ok=0; }
  grep -q '_PAGE_MATCH_WINDOW' "$dest/scripts/check_docx.py" 2>/dev/null \
    || { echo "  ✗ check_docx.py ยังเป็นรุ่นที่เลขแผ่นเพี้ยน (ไม่มีขอบเขตการจับคู่หน้า)"; ok=0; }
  grep -q '_NOT_A_SURNAME' "$dest/scripts/check_deep.py" 2>/dev/null \
    || { echo "  ✗ check_deep.py ยังเป็นรุ่นที่ฟ้อง 'Table (2019) ไม่มีในรายการอ้างอิง' (รุ่นเก่า)"; ok=0; }
  grep -rq 'opens in new window' "$dest/references/apa7-eng" 2>/dev/null \
    && { echo "  ✗ references/apa7-eng ยังมีขยะจากการ scrape ('opens in new window' ปนใน DOI)"; ok=0; }
  # ตัวตรวจ false positive: ถ้า sync แล้วสคริปต์เพี้ยน จะรู้ตรงนี้ก่อนเอาไปตรวจเล่มจริง
  if [ -f "$dest/scripts/test_false_positives.py" ]; then
    python3 "$dest/scripts/test_false_positives.py" >/dev/null 2>&1 \
      || { echo "  ✗ ชุดทดสอบ false positive ไม่ผ่าน — อย่าเพิ่งใช้รุ่นนี้ตรวจเล่มจริง"; ok=0; }
  else
    echo "  ✗ ไม่มี scripts/test_false_positives.py (รุ่นเก่า)"; ok=0
  fi
  [ "$ok" = 1 ] && echo "  ✓ ตรวจแล้วเป็นรุ่นล่าสุดครบ"
}

# --- ปลายทางที่ระบุมาเอง ---------------------------------------------------
if [ "$#" -ge 1 ]; then
  for dest in "$@"; do
    dest="${dest/#\~/$HOME}"
    echo; echo "→ คัดลอกไป: $dest"
    copy_to "$dest" && verify "$dest" || echo "  ✗ คัดลอกไม่สำเร็จ"
  done
  echo; echo "เสร็จแล้ว"
  exit 0
fi

# --- ค้นหาสำเนาในเครื่อง ---------------------------------------------------
echo
echo "กำลังค้นหาสำเนาสกิลในเครื่อง (อาจใช้เวลาสักครู่) ..."
FOUND=()
while IFS= read -r f; do
  d="$(cd "$(dirname "$f")/.." && pwd)"
  [ "$d" = "$SRC" ] && continue
  case " ${FOUND[*]-} " in *" $d "*) continue;; esac
  FOUND+=("$d")
done < <(find "$HOME/.codex" "$HOME/.claude" "$HOME/Documents" "$HOME/Desktop" \
              "$HOME/Library/Application Support/Claude" \
              -name check_docx.py -path '*/scripts/*' 2>/dev/null)

if [ "${#FOUND[@]}" -eq 0 ]; then
  echo
  echo "ไม่พบสำเนาอื่นในเครื่อง — แปลว่า Codex น่าจะอ่านจากโฟลเดอร์นี้โดยตรงอยู่แล้ว"
  echo "ถ้ารู้ว่าปลายทางอยู่ที่ไหน สั่งได้เลย:"
  echo "    bash \"$SRC/sync_skill.command\" <path ปลายทาง>"
  read -r -p "กด Enter เพื่อปิด" _; exit 0
fi

echo
echo "พบสำเนาเหล่านี้:"
for i in "${!FOUND[@]}"; do echo "  [$((i+1))] ${FOUND[$i]}"; done
echo "  [a] ทับทุกอัน"
echo "  [q] ยกเลิก"
echo
read -r -p "เลือกหมายเลขที่จะทับด้วยรุ่นล่าสุด: " ans

case "$ans" in
  q|Q) echo "ยกเลิก"; exit 0;;
  a|A) targets=("${FOUND[@]}");;
  *) idx=$((ans-1))
     if [ -z "${FOUND[$idx]-}" ]; then echo "หมายเลขไม่ถูกต้อง"; exit 1; fi
     targets=("${FOUND[$idx]}");;
esac

for dest in "${targets[@]}"; do
  echo; echo "→ ทับ: $dest"
  copy_to "$dest" && verify "$dest" || echo "  ✗ คัดลอกไม่สำเร็จ"
done

echo
echo "เสร็จแล้ว — สั่ง Codex ตรวจไฟล์ใหม่ได้เลย"
read -r -p "กด Enter เพื่อปิด" _
