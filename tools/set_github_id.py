#!/usr/bin/env python3
"""README·plugin.json 등에 있는 YOUR-GITHUB-ID 를 실제 GitHub 아이디로 한 번에 바꾼다.

  python tools/set_github_id.py 내아이디
"""
import os
import sys

PLACEHOLDER = "YOUR-GITHUB-ID"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEXT_EXT = {".md", ".json", ".txt", ".py"}

if len(sys.argv) != 2 or not sys.argv[1].strip() or "/" in sys.argv[1]:
    sys.exit("사용법: python tools/set_github_id.py <GitHub아이디>")
gid = sys.argv[1].strip()
changed = []
for root, dirs, files in os.walk(ROOT):
    dirs[:] = [d for d in dirs if d not in {".git", "__pycache__", "dist", "tools"}]
    for f in files:
        p = os.path.join(root, f)
        if os.path.splitext(f)[1] not in TEXT_EXT:
            continue
        with open(p, encoding="utf-8") as fh:
            t = fh.read()
        if PLACEHOLDER in t:
            with open(p, "w", encoding="utf-8", newline="") as fh:
                fh.write(t.replace(PLACEHOLDER, gid))
            changed.append(os.path.relpath(p, ROOT))
print(f"'{PLACEHOLDER}' → '{gid}' : {len(changed)}개 파일")
for c in changed:
    print("  -", c)
