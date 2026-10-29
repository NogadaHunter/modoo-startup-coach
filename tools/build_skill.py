#!/usr/bin/env python3
"""Claude.ai 업로드용 dist/modoo-startup-coach.skill 을 다시 만든다.

스킬 내용(plugins/modoo-startup-coach/skills/modoo-startup-coach/)을 고친 뒤 실행:
  python tools/build_skill.py
"""
import os
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NAME = "modoo-startup-coach"
SRC = os.path.join(ROOT, "plugins", NAME, "skills", NAME)
DST = os.path.join(ROOT, "dist", NAME + ".skill")
SKIP_DIRS = {"__pycache__", ".git"}
SKIP_FILES = {".DS_Store", "Thumbs.db"}

os.makedirs(os.path.dirname(DST), exist_ok=True)
count = 0
with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(SRC):
        dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS)
        for f in sorted(files):
            if f in SKIP_FILES or f.endswith(".pyc"):
                continue
            full = os.path.join(root, f)
            arc = NAME + "/" + os.path.relpath(full, SRC).replace(os.sep, "/")
            z.write(full, arc)
            count += 1
print(f"만들었어요: {os.path.relpath(DST, ROOT)} (파일 {count}개, {os.path.getsize(DST):,} bytes)")
