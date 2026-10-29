#!/usr/bin/env python3
"""모두의 창업 원고 기계 점검.

원고(draft.txt)를 '## 문항 이름' 줄로 나눠 글자 수와 위험 표현 후보를 찾는다.
결과는 후보일 뿐이다. 문맥을 읽고 진짜 문제만 리포트에 넣는다.

사용법:
  python3 check_draft.py draft.txt [--track 일반|로컬] [--limit 2000] [--json check.json]
  --limit 0  : 분량 검사 끄기 (사업계획서 등)
"""
import argparse
import json
import re
import sys
from datetime import date

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

THIS_YEAR = date.today().year

# 문항별 분량 규칙: (최소, 최대, 메모)
LIMITS = {
    "Q1": (10, None, "최소 10자"),
    "Q2": (None, 2000, ""), "Q3-1": (None, 2000, ""), "Q3-2": (None, 2000, ""),
    "Q4-1": (None, 2000, ""), "Q4-2": (None, 2000, ""), "Q7-1": (None, 2000, ""),
    "Q8": (None, 2000, "분량은 신청 화면 확인"), "재도전": (None, 2000, "분량은 신청 화면 확인"),
}
LONG_ITEMS = {"Q2", "Q3-1", "Q3-2", "Q4-1", "Q4-2", "Q7-1", "Q8", "재도전"}
EXPERIENCE_ITEMS = {"Q2", "Q8", "재도전", "1-1"}

PLAN_WORDS = (r"(계획|목표|예정|추진|하겠|할 것|할것|확보할|진행할|검토|가설|기대|예상|추정|목표치|도전|"
              r"단계|도모|로드맵|확대|전환|런칭|출시|분기|향후|이후|시점|추후|전망|위해|하려|할\s|될\s)")
SOURCE_WORDS = (r"(출처|자료\s*:|에 따르면|에 의하면|통계|보고서|발표|백서|리포트|"
                r"\(\s*[^)]*(19|20)\d\d[^)]*\)|20\d\d\s*년\s*기준|유로모니터|Statista|통계청|국가데이터처)")
SELF_EVIDENCE = r"(시식|응답자|설문|인터뷰|테스트|실험|자체|직접|자사|본인|대표자|집행|지원금|자기자본)"
OTHERS = r"(사가|사는|사의|기업|업체|경쟁사|브랜드가|시장은|시장이|시장 전체)"

BIG_UNIT = r"\d[\d,\.]*\s*(억\s*원|억|조\s*원|조|만\s*명|천\s*명|달러|\$|엔|만\s*가구|만\s*톤|개국)(?![가-힣])"
PCT = r"\d[\d,\.]*\s*(%|퍼센트)"
PCT_CONTEXT = r"(시장|성장|증가|감소|비중|점유|인구|가구|소비자|규모|비율)"
CLAIM = (r"(매출|판매량|누적\s*판매|계약|협약|MOU|업무\s*협약|특허\s*(등록|보유)|특허를\s*(보유|등록)|"
         r"투자\s*(유치|받)|입점|수상|선정되|고객\s*\d|회원\s*\d|사용자\s*\d|다운로드|재구매율|체결|공급\s*계약|납품)")
FIXED_NUM = (r"((정확도|정밀도|단축|절감|향상|감소|개선|증가|만족도)\D{0,8}\d+(\.\d+)?\s*%|"
             r"\d+(\.\d+)?\s*%\s*(의\s*)?(정확도|단축|절감|향상|감소|개선|증가))")
HYPE = ["혁신적", "획기적", "압도적", "완벽", "최초", "유일", "세계 최고", "국내 최고", "최고의", "게임 체인저", "게임체인저",
        "패러다임", "블루오션", "독보적", "혁명", "무한한", "폭발적", "전무후무", "독점", "독창적인", "차세대", "초격차"]
AI_STYLE = [
    (r"[—–]", "긴 줄표(—)"),
    (r"단순(히|한)\s?[^\s,.]{1,12}\s?(이|가)\s?아니라", "'단순히 ~가 아니라'"),
    (r"[을를]\s?넘어(서)?\s", "'~를 넘어'"),
    (r"새로운\s?지평", "'새로운 지평'"),
    (r"시너지", "'시너지'"),
    (r"선순환", "'선순환'"),
    (r"극대화", "'극대화'"),
    (r"선도(하|할|적)", "'선도'"),
    (r"에\s?기여할\s?(것|수)", "'~에 기여할 것'"),
    (r"(이는|이를 통해|이를 바탕으로)", "'이는/이를 통해' 연결어"),
    (r"현대\s?사회", "'현대 사회' 일반론"),
]
VAGUE = ["다양한", "여러 가지", "효율적", "최적의", "스마트한", "맞춤형", "차별화된", "편리한", "고도화", "체계적",
         "효과적", "원활한", "극대화"]
EXPERIENCE = r"(저는|제가|나는|내가|저희|직접|경험|현장에서|만나|인터뷰|물어|여쭤|확인했|겪|봤|보았|들었|해 보|해봤|시도|창업자는|대표자는)"
# 자리표시: 예시 칸 이름이 남은 것만 ([사회적 배경] 같은 소제목은 제외)
PLACEHOLDER_IN = r"(명$|이름|입력|작성|예시|지역명|품종|수치|숫자|날짜|금액|기관명|회사명|제품명|확인 필요|OO|○○|xx|XX|^\s*$)"
PLACEHOLDER_TOKEN = r"(○○|△△|XXX|xxx|\{[^}]{1,30}\})"
POLICY_GRUMBLE = r"(제도|정책|법|규제)\S{0,3}\s?(개선|개정|완화)\S{0,3}\s?(이|을|가)?\s?(필요|촉구|요구|시급)"
LOCAL_WORDS = r"(지역|마을|군|읍|면|특산|원물|농가|어촌|어가|상권|주민|로컬|향토|전통시장|산지|생산자|관광)"
BULLET_START = r"^\s*([-•◦○〇●▶▷※■□◆◇·]|\d+[.)]\s|[①-⑳]|[가-하][.)]\s|\[)"

REG = {
    "식품": (r"식품|가공식품|소스|잼|음료|반찬|밀키트|도시락|스프레드|베이커리|디저트", "식품위생법 영업신고·표시광고·HACCP·유통기한"),
    "주류": (r"주류|과실주|막걸리|맥주|와인|양조|증류주|전통주", "주류 면허법(제조·판매 면허), 주세법, 통신판매 고시"),
    "건강기능": (r"건강기능|기능성|혈당|다이어트|체지방|면역", "건강기능식품법, 식품 표시광고법(기능성·질병 표현 제한)"),
    "의료": (r"의료|진단|치료|병원|환자|처방|재활|의료기기", "의료법, 의료기기법(SW 의료기기 해당 여부), 민감정보"),
    "화장품": (r"화장품|스킨케어|코스메틱", "화장품법(제조·책임판매업, 기능성 표시)"),
    "개인정보": (r"개인정보|위치정보|얼굴 인식|음성 인식|생체|CCTV|데이터\s?수집|회원\s?정보", "개인정보 보호법, 위치정보법"),
    "드론·로봇": (r"드론|자율주행|로봇", "항공안전법, 전파법, 규제 샌드박스"),
    "금융": (r"결제|송금|대출|보험|투자\s?중개|핀테크|코인|가상자산", "전자금융거래법·자본시장법·보험업법 (로컬은 금융·보험업 제외 업종)"),
    "반려동물": (r"반려동물|펫|강아지|고양이|동물병원", "동물보호법 영업 등록, 수의사법 (로컬은 수의업 제외)"),
    "교육": (r"학원|과외|교습", "학원법(학원·교습소 등록)"),
    "전기제품": (r"배터리|충전기|전기\s?제품|전자\s?기기|웨어러블|IoT", "전기용품 및 생활용품 안전관리법(KC), 전파법"),
    "숙박·관광": (r"숙박|민박|게스트하우스|캠핑|글램핑|체험\s?프로그램", "관광진흥법, 공중위생관리법, 농어촌민박"),
    "운송": (r"셔틀|운송업|배달|퀵|모빌리티|카풀", "여객·화물자동차 운수사업법"),
    "부동산": (r"부동산|중개업|공유\s?오피스|공유\s?주방", "공인중개사법·건축법 (로컬은 부동산업 대부분 제외)"),
}


def split_sections(text):
    sections = []
    cur_name, cur_lines, cur_start = None, [], 1
    for i, line in enumerate(text.splitlines(), 1):
        m = re.match(r"^\s*##\s+(.+?)\s*$", line)
        if m:
            if cur_name is not None or any(l.strip() for l in cur_lines):
                sections.append((cur_name or "앞부분", cur_lines, cur_start))
            cur_name, cur_lines, cur_start = m.group(1), [], i + 1
        else:
            cur_lines.append(line)
    sections.append((cur_name or "전체", cur_lines, cur_start))
    return [(n, "\n".join(ls).strip("\n"), s) for n, ls, s in sections if n != "앞부분" or "\n".join(ls).strip()]


def item_id(name):
    m = re.match(r"^\s*(Q\s?\d+(?:-\d+)?)", name, re.I)
    if m:
        return m.group(1).upper().replace(" ", "")
    if "재도전" in name or "1기" in name or "달라진" in name:
        return "재도전"
    m = re.match(r"^\s*(\d+-\d+)", name)
    if m:
        return m.group(1)
    return name.strip()


def units(text):
    """PDF·붙여넣기에서 끊긴 줄을 이어 붙인 뒤 문장(또는 글머리 항목) 단위로 나눈다."""
    blocks, cur = [], ""
    for line in text.splitlines():
        s = line.strip()
        if not s:
            if cur:
                blocks.append(cur)
            cur = ""
            continue
        if re.match(BULLET_START, s) or not cur:
            if cur:
                blocks.append(cur)
            cur = s
        else:
            cur += " " + s
    if cur:
        blocks.append(cur)
    out = []
    for b in blocks:
        for part in re.split(r"(?<=[.!?。])\s+(?=[가-힣A-Z0-9\"“(\[])", b):
            if part.strip():
                out.append(part.strip())
    return out


def snippet(s, n=80):
    s = re.sub(r"\s+", " ", s).strip()
    return s if len(s) <= n else s[: n - 1] + "…"


def is_stat(s):
    if len(s) < 18:  # 표 칸 숫자
        return False
    if re.search(BIG_UNIT, s):
        return True
    return bool(re.search(PCT, s) and re.search(PCT_CONTEXT, s))


def check(text, track="일반", default_limit=2000):
    report = {"track": track, "sections": [], "flags": [], "regulation": [], "totals": {}}
    secs = split_sections(text)
    reg_hits = {}
    for name, body, start in secs:
        iid = item_id(name)
        with_sp = len(body.replace("\n", ""))
        no_sp = len(re.sub(r"\s", "", body))
        mn, mx, memo = LIMITS.get(iid, (None, None, ""))
        if default_limit == 0:
            mn, mx, memo = None, None, ""
        elif mx is not None and default_limit and default_limit != 2000:
            mx = default_limit
        info = {"name": name, "id": iid, "chars": with_sp, "chars_no_space": no_sp, "limit": mx, "memo": memo, "line": start}
        report["sections"].append(info)

        def flag(cat, text_, hint):
            report["flags"].append({"section": name, "category": cat, "text": snippet(text_), "hint": hint})

        if mx and with_sp > mx:
            flag("분량 초과", f"{with_sp:,}자 (제한 {mx:,}자)", f"{with_sp - mx:,}자 줄이기. 공백 제외 {no_sp:,}자")
        if mn and with_sp < mn:
            flag("분량 미달", f"{with_sp}자", f"최소 {mn}자")
        if iid in LONG_ITEMS and default_limit != 0 and 0 < with_sp < 500:
            flag("너무 짧음", f"{with_sp}자", "1~2줄 답은 '성의 없음'으로 읽힘. 핵심 문항은 800자 이상 권장")
        if iid == "Q1" and with_sp > 150:
            flag("한 줄 소개가 김", f"{with_sp}자", "60~120자 안에서 고객·방식·가치만")

        for s in units(body):
            planish = re.search(PLAN_WORDS, s)
            selfish = re.search(SELF_EVIDENCE, s)
            if is_stat(s) and not re.search(SOURCE_WORDS, s) and not selfish:
                flag("출처 필요", s, "통계·시장 수치에 (기관, 자료명, 연도)를 붙이기")
            yrs = [int(y) for y in re.findall(r"(?<!\d)(20[01]\d|202\d)(?!\d)", s)]
            old = [y for y in yrs if y <= THIS_YEAR - 4]
            if old and is_stat(s) and max(yrs) <= THIS_YEAR - 4:
                flag("오래된 자료?", s, f"{max(yrs)}년까지의 자료. 최근 2~3년 자료가 있는지 확인")
            if re.search(CLAIM, s) and not planish and not re.search(OTHERS, s):
                flag("증빙 필요(실적)", s, "실제로 있는 실적인가? 증빙 가능한가? 아니면 '계획'으로")
            if re.search(FIXED_NUM, s) and not re.search(r"(목표|가설|예상|기대|추정|예측|시뮬레이션|이내|이상을)", s):
                flag("확정 수치?", s, "실증 전 수치면 '목표'나 '가설'로")
            if re.search(POLICY_GRUMBLE, s):
                flag("제도 개선 요구?", s, "단순 제도개선·불만은 심사 제외. 내가 만드는 해결책으로 바꾸기")
            for w in HYPE:
                if w in s:
                    flag("과장·단정", s, f"'{w}' → 근거(숫자·비교)로 바꾸거나 빼기")
                    break
        for m in re.finditer(r"\[([^\]]{0,40})\]", body):
            inner = m.group(1)
            if re.match(r"^\d+$", inner):
                continue  # 각주 번호 [1]
            if re.search(PLACEHOLDER_IN, inner):
                flag("자리표시 남음", m.group(0), "실제 내용으로 바꾸기 (제출 전 필수)")
        for m in re.finditer(PLACEHOLDER_TOKEN, body):
            flag("자리표시 남음", m.group(0), "실제 내용으로 바꾸기 (제출 전 필수)")
        style_hits = []
        for pat, label in AI_STYLE:
            n = len(re.findall(pat, body))
            if n:
                style_hits.append(f"{label} {n}회")
        can = len(re.findall(r"할\s?수\s?있", body))
        if can >= 5:
            style_hits.append(f"'~할 수 있' {can}회")
        eng = len(re.findall(r"\([A-Za-z][A-Za-z\s\-&\.]{2,}\)", body))
        if eng >= 4:
            style_hits.append(f"영어 괄호 병기 {eng}회")
        if style_hits:
            flag("AI 문체 신호", ", ".join(style_hits), "해당 문장을 신청자 말투로. 연결어·줄표 줄이기")
        vague = sorted({w for w in VAGUE if w in body})
        if len(vague) >= 3:
            flag("막연한 표현", ", ".join(vague), "무엇이 몇 개, 무엇보다 어떻게 나은지 구체적으로")
        if iid in EXPERIENCE_ITEMS and with_sp > 100 and not re.search(EXPERIENCE, body):
            flag("경험 문장 없음", name, "신청자가 직접 한 일·본 것·들은 말을 넣기 (진정성)")
        if track.startswith("로컬") and iid in {"Q1", "Q2", "Q3-1"} and with_sp > 0 and not re.search(LOCAL_WORDS, body):
            flag("지역 자원 언급 없음", name, "로컬 트랙: 어떤 지역 자원을 쓰는지, 지역에 무엇이 남는지")
        for field, (pat, law) in REG.items():
            found = sorted(set(re.findall(pat, body)))
            if found:
                reg_hits.setdefault(field, {"field": field, "keywords": set(), "check": law, "sections": set()})
                reg_hits[field]["keywords"].update(found)
                reg_hits[field]["sections"].add(name)
    for v in reg_hits.values():
        report["regulation"].append({"field": v["field"], "keywords": sorted(v["keywords"])[:8],
                                      "check": v["check"], "sections": sorted(v["sections"])})
    report["totals"] = {
        "sections": len(report["sections"]),
        "chars": sum(s["chars"] for s in report["sections"]),
        "flags": len(report["flags"]),
        "by_category": {},
    }
    for f in report["flags"]:
        report["totals"]["by_category"][f["category"]] = report["totals"]["by_category"].get(f["category"], 0) + 1
    return report


def print_report(r):
    print(f"■ 트랙: {r['track']} · 문항 {r['totals']['sections']}개 · 전체 {r['totals']['chars']:,}자 · 후보 {r['totals']['flags']}건")
    print("\n[글자 수] (공백 포함 / 공백 제외, 줄바꿈 제외. 최종 확인은 신청 화면에서)")
    for s in r["sections"]:
        lim = f" / 제한 {s['limit']:,}" if s["limit"] else ""
        warn = "  ← 초과" if s["limit"] and s["chars"] > s["limit"] else ""
        memo = f"  ({s['memo']})" if s["memo"] else ""
        print(f"  {s['name'][:28]:<28} {s['chars']:>6,} / {s['chars_no_space']:>6,}{lim}{warn}{memo}")
    if r["totals"]["by_category"]:
        print("\n[후보 수] " + " · ".join(f"{k} {v}" for k, v in r["totals"]["by_category"].items()))
    cur = None
    for f in r["flags"]:
        if f["section"] != cur:
            cur = f["section"]
            print(f"\n▶ {cur}")
        print(f"  - [{f['category']}] {f['text']}\n      → {f['hint']}")
    if r["regulation"]:
        print("\n[규제 확인 키워드]")
        for g in r["regulation"]:
            print(f"  - {g['field']}: {', '.join(g['keywords'])} → {g['check']}")
    print("\n※ 위 결과는 자동으로 찾은 '후보'예요. 문맥을 읽고 진짜 문제만 리포트에 넣으세요.")


def main():
    ap = argparse.ArgumentParser(description="모두의 창업 원고 기계 점검")
    ap.add_argument("draft")
    ap.add_argument("--track", default="일반", help="일반 | 로컬")
    ap.add_argument("--limit", type=int, default=2000, help="긴 문항 분량 제한 (0 = 검사 안 함)")
    ap.add_argument("--json", help="결과 JSON 저장 경로")
    a = ap.parse_args()
    with open(a.draft, encoding="utf-8-sig") as f:
        text = f.read()
    r = check(text, a.track, a.limit)
    print_report(r)
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, indent=2)
        print(f"\nJSON 저장: {a.json}")


if __name__ == "__main__":
    main()
