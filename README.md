# 모두의 창업 신청서 코칭 에이전트 (modoo-startup-coach)

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/deed.ko)
![Claude Skill](https://img.shields.io/badge/Claude-Skill-D97757)
![version](https://img.shields.io/badge/version-1.0.1-2E7D52)

정부 「모두의 창업 프로젝트」(중소벤처기업부·창업진흥원·소상공인시장진흥공단, [modoo.or.kr](https://www.modoo.or.kr)) 참가자를 돕는 **컨설턴트용 Claude 스킬**이에요.
도전신청서·사업계획서·오디션 발표자료를 공고와 평가 기준에 맞춰 진단하고, 신청자에게 바로 건넬 수 있는 **코칭 리포트(DOCX)** 를 만들어요.

> 신청서 파일 하나를 올리면 → 공고 기준 자격 점검 + 평가자 시선의 채점 + 문항별 피드백 + 코칭 리포트까지 만들어져요.

> [!IMPORTANT]
> 이 스킬은 중소벤처기업부·창업진흥원·소상공인시장진흥공단의 공식 도구가 아니에요. 공고문과 공개 자료를 바탕으로 만든 **독립 코칭 도구**이고, 점수는 **코칭용 추정**이에요. 선정을 보장하지 않아요.

## 무엇을 하나요

| 모드 | 이럴 때 | 결과물 |
|---|---|---|
| ① 자격·트랙 진단 | 신청 전, 자격이나 트랙이 헷갈릴 때 | 진단표 (통과 / 확인 필요 / 제외 위험) |
| ② 인터뷰 → 작성 설계도 | 초안이 아직 없을 때 | 인터뷰 질문지 DOCX → 문항별 작성 설계도 DOCX |
| ③ 도전신청서 피드백 | 초안이 있을 때 (핵심) | 피드백 리포트 DOCX |
| ④ 고친 원고 다시 보기 | 피드백을 반영했을 때 | 점수 변화·반영 여부 리포트 DOCX |
| ⑤ 사업계획서 검토 | 1R 진출 뒤 지역 예선·오디션 서류 | 숫자 일치·집행계획 점검 리포트 DOCX |
| ⑥ 오디션 발표(IR) 검토 | 지역·권역·전국 오디션 | 장표별 피드백 + 예상 질문 리포트 DOCX |
| ⑦ 제출 직전 최종 점검 | 마감 1~3일 전 | 17개 항목 체크 결과 |

리포트에는 코칭 등급(A~D)과 5관점 점수 막대, 문항 카드(잘된 점·고칠 점·고쳐 쓰기 참고 예시·물어볼 것), 사실·근거 확인표, 본인 표현 점검, 다음 코칭 숙제가 들어가요. 예시는 [`examples/`](examples/)에 있어요.

### 대신 써 주지 않아요

공고문에는 "사업계획서 등을 **타인이 대신 작성하여 제출하는 경우**, 작성자(대필자)와 도전자 등 관련자 전원이 처벌될 수 있다"는 조항이 있고, 2기부터 AI 생성률·표절 검사도 강화됐어요.
그래서 이 스킬은 진단·질문·작성 설계도·피드백·짧은 참고 예시까지만 하고, 본문은 신청자가 자기 경험과 말로 쓰게 이끌어요.

## 설치

### A. Claude.ai (웹·데스크톱 앱)

1. [`dist/modoo-startup-coach.skill`](dist/modoo-startup-coach.skill) 을 내려받아요 (파일을 열고 **Download raw file**).
2. claude.ai → 왼쪽 아래 내 이름 → **설정 → 기능** → '코드 실행 및 파일 생성'의 **클라우드에서 코드 실행**을 켜요.
3. 왼쪽 메뉴 **사용자 지정 → Skills → + 추가 → Skill 업로드** → 받은 `.skill` 파일을 넣고 **업로드**. (압축은 풀지 않아요)
4. 목록에 modoo-startup-coach가 보이고 스위치가 켜져 있으면 끝이에요.

화면 캡처로 보는 등록 방법: [`docs/claude-ai-skill-register-guide.pptx`](docs/claude-ai-skill-register-guide.pptx)

### B. Claude Code (플러그인 마켓플레이스)

Claude Code 안에서:

```text
/plugin marketplace add NogadaHunter/modoo-startup-coach
/plugin install modoo-startup-coach@peopleware-skills
```

터미널에서:

```bash
claude plugin marketplace add NogadaHunter/modoo-startup-coach
claude plugin install modoo-startup-coach@peopleware-skills
```

플러그인 스킬은 `/modoo-startup-coach:modoo-startup-coach` 로 직접 부를 수 있고, "모두의 창업 신청서 검토해줘"처럼 말하면 자동으로 쓰여요.

### C. 직접 복사

[`plugins/modoo-startup-coach/skills/modoo-startup-coach/`](plugins/modoo-startup-coach/skills/modoo-startup-coach/) 폴더를 통째로 `~/.claude/skills/` 아래에 복사해요.

## 사용 예

```text
(파일 첨부: 신청자 도전신청서)
모두의 창업 도전신청서 검토해서 피드백 리포트 만들어줘.
신청자: [이름] / 트랙: [일반·기술 또는 로컬] / 멘토기관: [기관명]
컨설턴트: [내 이름], [소속]
```

```text
(파일 첨부: 발표자료 PDF)
모두의 창업 지역 오디션 발표자료야. 발표 7분, 질의응답 5분이야.
장표별로 고칠 점을 알려 주고, 심사위원 예상 질문 20개와 답변 방향을 정리해줘.
```

상황별 프롬프트 전체: [`docs/USAGE.md`](docs/USAGE.md) · 가이드북: [`docs/modoo-startup-coach-guidebook.pdf`](docs/modoo-startup-coach-guidebook.pdf)

## 필요한 것

- Claude.ai: **클라우드에서 코드 실행** 켜기 (리포트 DOCX를 만들 때 필요)
- Claude Code: Python 3 + `python-docx` (없으면 스킬이 `pip install python-docx` 를 시도해요). PDF 읽기에는 `pymupdf` 또는 `pypdf`

## 폴더 구조

```text
modoo-startup-coach/
├── .claude-plugin/marketplace.json        ← Claude Code 마켓플레이스 목록
├── plugins/modoo-startup-coach/
│   ├── .claude-plugin/plugin.json         ← 플러그인 정보 (이름·버전)
│   └── skills/modoo-startup-coach/        ← 스킬 본체
│       ├── SKILL.md                       ← 모드 7개, 원칙, 절차
│       ├── references/                    ← 공고 요약, 자격, 문항 가이드, 채점 기준, 인터뷰, 대필·사실 규칙, 라운드 점검표
│       └── scripts/                       ← read_doc.py · check_draft.py · build_report.py
├── dist/modoo-startup-coach.skill         ← Claude.ai 업로드용 (위 스킬 폴더를 묶은 zip)
├── docs/                                  ← 사용 프롬프트, 가이드북, 등록 방법, 올리는 법
├── examples/                              ← 가상 신청자 예시 리포트
└── tools/                                 ← build_skill.py · check_repo.py · set_github_id.py
```

## 기준과 업데이트

- 기준 공고: **중소벤처기업부 공고 제2026-511호 「모두의 창업 프로젝트」 통합 모집공고(2차), 2026.8.20.**
- 다른 차수·연도에는 새 공고문이나 신청 화면 캡처를 함께 첨부하면, 바뀐 문항·분량·자격부터 맞춰서 봐요.
- 스킬을 고친 뒤에는 `python tools/build_skill.py` 로 `dist/` 를 다시 만들고, `plugin.json` 의 `version` 을 올려 주세요. 변경 내역: [`CHANGELOG.md`](CHANGELOG.md)

## 참고 자료

- 중소벤처기업부, 「모두의 창업 프로젝트」 통합 모집공고(2차), 공고 제2026-511호, 2026.8.20.
- 모두의 창업 누리집, https://www.modoo.or.kr
- k-startup.ai, 「모두의 창업 2차 도전신청서 작성 가이드」, https://www.k-startup.ai/guides/modoo-challenge-application
- 멘토기관 안내자료, 「모두의 창업 2기 분석 및 신청서 작성 Tip」(2026) — 평가지표 예시, BAD & GOOD CASE, 제출 직전 체크 항목을 요약·인용

## 라이선스

[CC BY-NC-SA 4.0](LICENSE) — 출처를 밝히면 자유롭게 쓰고 고칠 수 있어요. 단, **영리 목적으로 쓸 수 없고**, 고친 것도 같은 조건으로 공유해야 해요.
공고문 등 공공 자료의 저작권은 각 기관에 있어요.

## 만든 곳

피플웨어교육컨설팅그룹 · AI프롬프트연구소

---

### English summary

A Claude skill (and Claude Code plugin) for consultants coaching applicants of Korea's government startup program **"모두의 창업 프로젝트" (Modoo Startup Project)**. It checks eligibility, scores application drafts against the official evaluation criteria, and produces Korean coaching reports (DOCX) for application forms, regional-round business plans and audition pitch decks. It never ghost-writes applications. Licensed under CC BY-NC-SA 4.0. Not an official government tool.
