# 리포트(DOCX) 만드는 법

DOCX는 `scripts/build_report.py`로만 만든다. Claude는 `report.json`만 쓴다. 디자인(표지·색·표·점수 막대·쪽 번호·유의사항)은 스크립트가 매번 같게 넣는다.

```bash
python3 $S/build_report.py report.json -o "<아이디어명>_<문서종류>_v1.docx"
python3 $S/build_report.py --verify "<아이디어명>_<문서종류>_v1.docx"
python3 $S/build_report.py --example > example.json    # 모든 블록이 든 예시 JSON
```

## 1. JSON 뼈대

```json
{
  "doc_type": "review",
  "meta": {
    "title": "도전신청서 피드백 리포트",
    "idea": "아이디어명",
    "applicant": "신청자 이름 또는 닉네임",
    "track": "일반/기술",
    "stage": "1R 도전신청서 (2026년 2차)",
    "mentor_org": "선택한 멘토기관",
    "consultant": "컨설턴트 이름",
    "consultant_org": "소속",
    "date": "2026-10-30",
    "version": "v1"
  },
  "summary": {
    "grade": "B",
    "total": 72,
    "verdict": "총평 두세 문장",
    "scores": [
      {"name": "A 문제 정의·진정성", "score": 4, "max": 5, "weight": 20, "note": "근거 한 줄"}
    ],
    "org_scores": [{"name": "도전의 진정성", "score": 4, "max": 5, "note": ""}],
    "strengths": ["잘된 점"],
    "top_fixes": ["가장 먼저 고칠 점"],
    "red_flags": ["감점 신호"]
  },
  "sections": [
    {"title": "자격·요건 점검", "intro": "선택: 첫 문단", "blocks": []}
  ],
  "notes": ["리포트 끝 유의사항에 더할 말(선택)"]
}
```

- `doc_type`: `review`(도전신청서 피드백) · `compare`(다시 보기) · `plan_review`(사업계획서) · `ir_review`(발표) · `blueprint`(작성 설계도) · `interview`(인터뷰 질문지) · `eligibility`(자격 진단). 표지 색 띠와 유의사항 문구가 바뀐다.
- `meta`의 빈 값은 표지에 `-`로 나온다. 컨설턴트를 모르면 비워 둔다.
- `summary`는 선택. 넣으면 '한눈에 보기' 쪽이 생긴다(등급 배지, 점수 막대, 잘된 점·고칠 점 두 칸, 감점 신호).
- 장 번호(1, 2, 3…)와 목차는 스크립트가 붙인다. `title`에 번호를 쓰지 않는다.
- 글 안에서 `**굵게**`를 쓸 수 있다. 줄바꿈은 `\n`.

## 2. 블록 종류

| type | 쓰는 곳 | 키 |
|---|---|---|
| `p` | 문단 | `text` |
| `h` | 소제목 | `text` |
| `bullets` / `numbers` | 목록 | `items` (문자열 배열) |
| `table` | 표 | `headers`, `rows`, `widths`(cm, 선택). '통과/확인 필요/제외 위험/반영/일부 반영/미반영/일치/불일치/고칠 것' 칸은 색이 자동으로 칠해진다 |
| `callout` | 강조 상자 | `tone`(info·good·warn·bad), `title`, `text`, `items`(선택) |
| `scores` | 점수 막대 | `items`: `name`, `score`, `max`, `note` |
| `question` | 문항 카드 (③④⑤) | `no`, `name`, `chars`, `limit`, `score`, `max`, `looking_for`, `quote`, `good`[], `fix`[], `rewrite`{`before`,`after`,`why`}, `ask`[] |
| `rewrite` | 전→후 예시 단독 | `before`, `after`, `why` |
| `blueprint` | 작성 설계도 (②) | `no`, `name`, `looking_for`, `message`, `facts`[], `outline`[{`step`,`what`,`chars`}], `blanks`[], `evidence`[], `avoid`[], `length` |
| `interview` | 질문지 (②) | `groups`[{`title`, `questions`[]}], `lines`(답 칸 줄 수, 기본 4) |
| `qa` | 예상 질문 (⑥) | `items`[{`q`, `direction`, `prepare`}] |
| `checklist` | 체크리스트 (⑦) | `items`: 문자열 또는 {`text`, `status`} |
| `compare` | 점수 변화 (④) | `items`[{`name`, `before`, `after`, `reason`}] |
| `slide` | 장표 카드 (⑥) | `no`, `title`, `message`, `good`[], `fix`[], `retitle` |
| `pagebreak` | 쪽 나눔 | |

`question.rewrite.after`는 신청자 문장을 다듬은 **참고용 예시**다. 신청자가 말하지 않은 사실을 넣지 말고, 빈 곳은 `[신청자 작성: …]`으로 둔다. 스크립트가 '참고용, 그대로 쓰지 말고 본인 말로' 표시를 붙인다.

## 3. 모드별 장 구성

### ③ 도전신청서 피드백 (`review`)
1. 자격·요건 점검 — `table`(확인 항목 / 신청자 상황 / 판정 / 할 일)
2. 평가자 3인의 시선 — `callout` 3개(현장·기술 / 사업·투자 / 창업 교육)
3. 문항별 피드백 — 문항마다 `question`
4. 사실·근거 확인표 — `table`(주장 / 위치 / 근거 / 판정 / 할 일)
5. 본인 표현 점검 — `table`(표현 / 위치 / 바꾸는 방향) + `p`
6. 다음 코칭 숙제 — `numbers` + `checklist`(제출 전)

### ④ 다시 보기 (`compare`)
1. 점수 변화 — `compare`(관점별) + `compare`(문항별)
2. 지난 지적 반영 여부 — `table`(지난 지적 / 판정(반영·일부 반영·미반영) / 근거)
3. 새로 생긴 문제 — `bullets`
4. 문항별 피드백 — 바뀐 문항만 `question`
5. 다음 숙제

### ⑤ 사업계획서 (`plan_review`)
1. 1R 진전 점검 — `table`
2. 평가자 시선 — `callout`
3. 장·절별 피드백 — 절마다 `question`(`no`에 '1-1', `limit` 0)
4. 숫자 일치 점검 — `table`(항목 / 위치 / 적힌 값 / 다시 계산 / 판정 / 조치)
5. 집행계획 점검 — `table`
6. 사실·근거, 본인 표현, 다음 숙제

### ⑥ 발표 (`ir_review`)
1. 흐름 점검 — `table`(순서 / 있어야 할 장표 / 지금 / 판정)
2. 장표별 피드백 — 장마다 `slide`
3. 시간 배분 제안 — `table`
4. 예상 질문 — `qa`
5. 리허설 체크리스트 — `checklist`

### ② 질문지 (`interview`), 설계도 (`blueprint`)
- 질문지: 안내 `callout`(작성 요령) → `interview`.
- 설계도: 전체 흐름 `callout`(Q1 후보 2~3개, 문항을 잇는 한 줄) → 문항마다 `blueprint` → 확인할 것 `table` → 다음 숙제.

## 4. 문장 쓰는 법 (리포트 안)

- 신청자에게 건네는 문서다. 존댓말(해요체), 짧은 문장, 쉬운 말.
- 지적은 "무엇이 / 왜 / 어떻게"를 한 묶음으로. 예: "Q2 첫 문단이 일반론이에요. 평가자는 신청자가 문제를 직접 본 장면을 찾아요. 처음 그 문제를 본 날을 두세 문장으로 써 보세요."
- 칭찬은 구체적으로(어느 문장이 왜 좋은지).
- 리포트 문장 자체도 과장어·긴 줄표·영어 남발을 피한다.
- 점수는 "코칭용 추정"이라는 말을 요약 쪽에 한 번 쓴다(스크립트가 배지 아래에 자동 표기).
