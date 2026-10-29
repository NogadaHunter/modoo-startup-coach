#!/usr/bin/env python3
"""올리기 전에 저장소 구조를 점검한다 (claude CLI가 없을 때 쓰는 간이 점검).
claude CLI가 있으면 `claude plugin validate .` 도 함께 돌린다.

  python tools/check_repo.py
"""
import json
import os
import re
import sys
import zipfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KEBAB = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")
RESERVED = re.compile(r"^(claude|anthropic|anthropics|cc-plugin)(-|$)|claude-code|official")
errors, warns = [], []


def load(path):
    try:
        with open(os.path.join(ROOT, path), encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        errors.append(f"{path}: 파일이 없어요")
    except json.JSONDecodeError as e:
        errors.append(f"{path}: JSON 문법 오류 ({e})")
    return None


mk = load(".claude-plugin/marketplace.json")
if mk:
    for k in ("name", "owner", "plugins"):
        if k not in mk:
            errors.append(f"marketplace.json: '{k}' 필드가 없어요")
    if mk.get("name") and (not KEBAB.match(mk["name"]) or RESERVED.search(mk["name"])):
        errors.append(f"marketplace.json: 이름 '{mk['name']}'이 규칙에 맞지 않아요")
    if not (mk.get("owner") or {}).get("name"):
        errors.append("marketplace.json: owner.name 이 없어요")
    for i, p in enumerate(mk.get("plugins", [])):
        src = p.get("source", "")
        if ".." in str(src):
            errors.append(f"plugins[{i}]: source 에 '..' 를 쓸 수 없어요")
        if isinstance(src, str):
            pdir = os.path.join(ROOT, src)
            if not os.path.isdir(pdir):
                errors.append(f"plugins[{i}]: source 폴더가 없어요 ({src})")
                continue
            man = load(os.path.join(src, ".claude-plugin", "plugin.json"))
            if man:
                if man.get("name") != p.get("name"):
                    errors.append(f"plugins[{i}]: marketplace 이름 '{p.get('name')}' 과 plugin.json 이름 '{man.get('name')}' 이 달라요")
                if not KEBAB.match(man.get("name", "")):
                    errors.append(f"plugin.json: 이름 '{man.get('name')}' 은 kebab-case 여야 해요")
                for k in ("version", "description", "author"):
                    if k not in man:
                        warns.append(f"plugin.json: '{k}' 가 없어요")
            sdir = os.path.join(pdir, "skills")
            for s in sorted(os.listdir(sdir)) if os.path.isdir(sdir) else []:
                sk = os.path.join(sdir, s, "SKILL.md")
                if not os.path.isfile(sk):
                    errors.append(f"skills/{s}: SKILL.md 가 없어요")
                    continue
                head = open(sk, encoding="utf-8").read().split("---")
                if len(head) < 3:
                    errors.append(f"skills/{s}/SKILL.md: 맨 위 --- 머리말(frontmatter)이 없어요")
                    continue
                fm = head[1]
                nm = re.search(r"^name:\s*(.+)$", fm, re.M)
                ds = re.search(r"^description:\s*(.+)$", fm, re.M)
                if not nm or nm.group(1).strip() != s:
                    errors.append(f"skills/{s}/SKILL.md: name 이 폴더 이름과 같아야 해요")
                if not ds:
                    errors.append(f"skills/{s}/SKILL.md: description 이 없어요")
                elif len(ds.group(1).strip()) > 1024:
                    warns.append(f"skills/{s}/SKILL.md: description 이 {len(ds.group(1).strip())}자예요 (Claude.ai 업로드는 1,024자 이하 권장)")
                for root, dirs, files in os.walk(os.path.join(sdir, s)):
                    if "__pycache__" in dirs:
                        warns.append(f"{os.path.relpath(root, ROOT)}/__pycache__ 는 지우고 올리세요")

dist = os.path.join(ROOT, "dist", "modoo-startup-coach.skill")
if os.path.isfile(dist):
    with zipfile.ZipFile(dist) as z:
        names = z.namelist()
    if "modoo-startup-coach/SKILL.md" not in names:
        errors.append("dist/*.skill 안에 modoo-startup-coach/SKILL.md 가 없어요")
    src_count = sum(len(f) for _, _, f in os.walk(os.path.join(ROOT, "plugins/modoo-startup-coach/skills/modoo-startup-coach")))
    if len(names) != src_count:
        warns.append(f"dist/*.skill 파일 수({len(names)})가 스킬 폴더({src_count})와 달라요 → python tools/build_skill.py")
else:
    errors.append("dist/modoo-startup-coach.skill 이 없어요 → python tools/build_skill.py")

for need in ("README.md", "LICENSE"):
    if not os.path.isfile(os.path.join(ROOT, need)):
        errors.append(f"{need} 가 없어요")

placeholder = []
for root, dirs, files in os.walk(ROOT):
    dirs[:] = [d for d in dirs if d not in {".git", "dist", "tools"}]
    for f in files:
        if f.endswith((".md", ".json")):
            if "YOUR-GITHUB-ID" in open(os.path.join(root, f), encoding="utf-8").read():
                placeholder.append(os.path.relpath(os.path.join(root, f), ROOT))
if placeholder:
    warns.append("YOUR-GITHUB-ID 자리표시가 남아 있어요 → python tools/set_github_id.py <아이디> : " + ", ".join(placeholder))

print("■ 저장소 점검")
for e in errors:
    print("  ✖", e)
for w in warns:
    print("  ⚠", w)
print("  ✔ 통과" if not errors else f"  실패 {len(errors)}건")
sys.exit(1 if errors else 0)
