# 부의금 정산 스킬 (Claude Code)

장례를 치른 뒤 **부의금·조화·장례 지출을 정리하고 가족별로 나누는 일**을 Claude 에게 맡기는 스킬입니다.

```
① 구글 시트 세팅 → ② 입력하면 자동 점검 → ③ 일차 소계 → ④ 지출·공금 → ⑤ 가족 정산 → ⑥ 공유용 정리 → ⑦ 한 분당 엑셀 + PDF
```

- 부의금 · 근조화환/근조기 · 지출을 **구글 시트 하나**에 적으면 요약·가족별·물품 통계가 바로 계산됩니다.
- 금액이 글자로 적혔거나, 받은 가족이 비었거나, 같은 줄이 두 번 적히면 노란색으로 알려 줍니다.
- 직접 적은 `1일차 소계`·`총합계` 줄과 시트 계산이 맞는지 검산합니다.
- 마지막에 가족 **한 분당 폴더 하나**(엑셀 1개 + PDF 요약 리포트 1개)로 나눠 줍니다.
- **사용자가 직접 적은 칸은 절대 고치지 않습니다.** 이상해 보이면 알리기만 합니다.

## 설치

Claude Code 에서:

```
/plugin marketplace add Ikek102030/bugeum-settlement-skill
/plugin install bugeum-settlement@bugeum-settlement
```

또는 `skills/bugeum-settlement` 폴더를 `~/.claude/skills/` 에 복사합니다.

## 필요한 것

- Claude 의 **Google Sheets / Google Drive 커넥터** (시트 읽기·쓰기, 엑셀로 내보내기)
- 마지막 단계(한 분당 파일): Python 3 + `openpyxl`, PDF 는 Google Chrome

## 쓰는 법

빈 구글 시트를 만들어 Claude 에게 링크를 주고 "부의금 정리할 시트 세팅해줘" 라고 하면 됩니다.
그 다음은 시트에 적고, 필요한 걸 말로 요청하면 됩니다. 예:

- "근조화환 겹치는 거 있는지 봐줘"
- "1일차, 2일차 소계 보고 싶어"
- "지출 정산 시트 세팅해줘"
- "시트 공유할 건데 예쁘게 정리해줘"
- "한 분당 엑셀이랑 PDF 로 나눠줘"

## 구성

```
skills/bugeum-settlement/
  SKILL.md                     작업 순서와 규칙
  references/sheet-formulas.md 탭별 수식 (실제로 검증한 것)
  scripts/deliver.py           내보낸 엑셀 → 한 분당 엑셀 + PDF
```

실제 장부 데이터(이름·금액)는 이 저장소에 들어 있지 않습니다.
