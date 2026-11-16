---
name: excel-report
description: 엑셀 원본 자료로 팀장 보고용 보고서(xlsx)를 만들 때 사용. 수식 참조, 검산 열, 서식, A4 인쇄, Windows Excel COM 작업 방법 포함
---

# 엑셀 보고서 작성 규칙

## 내용
- 지정한 자료만 사용한다. 자료에 없는 비교(전주 대비, 목표 대비 등)는 넣지 않는다.
- 표는 맨 위에, 설명 문장은 표 아래에 둔다. "표 위에 핵심 N줄"처럼 따로 요청하면 그 요청을 따른다.
- 표의 모든 값은 원본 셀을 참조하는 수식(SUMIFS, INDEX/MATCH 등)으로 넣는다.
- 핵심 요약 문장도 가능하면 `TEXT()`를 쓴 수식으로 만들어 원본이 바뀌면 함께 바뀌게 한다.
- 검산 열(예: 수량×단가)과 비중 합계 100%를 함께 보여 준다.
- 원본의 날짜(`10/5(월)`)와 마감(`10/7`)은 텍스트다. SUMIFS 조건은 이 텍스트와 똑같이 맞춘다.

## 서식
- 글꼴 `맑은 고딕`, 천 단위 쉼표, 비율 `0.0%`
- 표 머리글은 남색 배경에 흰 글씨, 합계 행은 굵게
- A4 한 장에 인쇄되도록 맞춘다.
- 파일 이름: `<보고서종류>_<MMDD-MMDD>.xlsx`

## 작업 환경 (Windows)
- Python이 없다. PowerShell에서 `New-Object -ComObject Excel.Application`을 쓰고, `finally`에서 `$x.Quit()`를 호출한다.
- 한글이 든 `.ps1`은 `Invoke-Expression (Get-Content -Raw -Encoding UTF8 <script.ps1>)`로 실행한다.
- `Range.Value2`에 배열을 바로 넣지 말고, 원본 범위는 `Range.Copy(대상Range)`로 복사한다.
- 새 통합문서 전에 `SheetsInNewWorkbook = 1`로 바꾸고 작업 후 원래 값으로 돌려놓는다.
- 저장은 `SaveAs(path, 51)`, 레이아웃 확인은 `ExportAsFixedFormat(0, pdf)`로 한다.
