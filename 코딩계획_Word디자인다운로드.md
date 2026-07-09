# 코딩 계획 — "Download for Word" 버튼 & 샘플 디자인 재현

## 목표
1. 리포트가 생성된 후, 화면 상단 우측(빨간 원 위치)에 **"Download for Word"** 버튼을 만든다.
2. 버튼을 누르면 폴더의 `Eng-Sample_SAT_Analysis_Report & Improvement Plan.docx` 와 **똑같은 디자인 포맷**의 Word 파일이 다운로드되게 한다.

## 샘플 문서 분석 결과 (재현할 디자인)
샘플은 단순 마크다운 변환이 아니라, 다음 요소로 구성된 디자인 레이아웃이다.

- **히어로 배너** (네이비 `0B2545` 박스): "ELITE PREP" + "SAT Analysis Report & Improvement Plan"
- **인포 바** (2열, 하단 금색 `C9A227` 선): STUDENT 이름 | 진단시험 개수·범위
- **Score Trajectory Overview** 섹션 + 3개 통계 카드 (SCORE RANGE / PEAK / LATEST, 연회색 `EEF2F7`)
- **Diagnostic Report Overview** (Student Name / Test Trajectory / Performance Trend)
- **Key Weaknesses & Curricular Solutions**: 약점별 심각도 배지
  (CRITICAL WEAKNESS = 빨강 `B3261E`, FOUNDATIONAL GAP = 파랑 `2E6BB8`) + 제목 + 설명 + Curriculum Chapters 목록
- **Tutoring Recommendations & Logistics** + 3개 통계 카드 (TARGET SCORE / TOTAL HOURS / FREQUENCY)
- **Week-by-Week Study Plan**: 3열 표(Week / Focus / Curriculum), 행 교차 음영(`EEF2F7`)
- **푸터**: 가운데 정렬 회색 텍스트

색상: 네이비 `0B2545`, 슬레이트 `14315B`, 파랑 `2E6BB8`, 본문 `22303F`,
회색 `5B6772`, 빨강 `B3261E`, 연회색 배경 `EEF2F7`, 표 격자 `D9DEE7`.

## 구현 단계
1. **파서 함수** `_parse_report(md_text)`: 생성된 마크다운 리포트에서
   학생명, 시험 궤적(점수), Performance Trend, 약점 목록(심각도·챕터),
   목표 점수·시간·빈도, 주차별 학습계획을 추출 (모든 항목 예외 처리·폴백).
2. **스타일 헬퍼**: 셀 음영(`_shade_cell`), 테이블 테두리 설정, 통계 카드/배지 생성 함수.
3. **빌더 함수** `convert_report_to_styled_docx(md_text)`: 위 데이터로 샘플과 동일한
   docx를 조립해 bytes 반환.
4. **UI 변경**: 상단 헤더를 좌(제목)/우(버튼) 2열로 배치하고,
   리포트 생성 후 우측에 "Download for Word" 버튼을 채운다.
   본문의 기존 단순 Word 버튼은 이 스타일 버전으로 교체, PDF 버튼은 유지.

## 추가 (2차) — 점수 추이 그래프 포함
샘플 docx의 "Score Trajectory Overview"에는 임베드된 **꺾은선 차트**가 있었음(문단3, 6.9"×3.2").
내 생성본에 이 그래프가 빠져 있어 추가한다.

- **데이터 확보**: Gemini 프롬프트 끝에 시험별 섹션 점수(Total/Math/RW) 구조화 블록
  ```scores [{"test","total","math","rw"}, ...]``` 을 추가로 출력하도록 지시.
- **파싱**: `_parse_report`에서 `scores` 배열 추출. 궤적(trajectory)이 없으면 scores로 대체.
- **차트**: matplotlib로 샘플과 동일한 스타일 렌더 —
  네이비 Total(다이아 마커, 데이터 라벨), 파랑 Math(원), 금색 RW(사각),
  최고점 "Peak" 표시, 하단 범례, y축 Score, 연회색 그리드 → PNG.
- **삽입 위치**: 제목 → 차트 이미지 → 이탤릭 부제 → 통계카드 순으로 재배치.
- **표시 정리**: 화면 리포트에서 ```scores 블록도 숨김 처리.
- scores 블록이 없으면(구버전 리포트) Total 한 줄만 그리는 폴백.

## 안전장치
- 파싱 실패 시 해당 섹션만 건너뛰고 문서 생성은 계속(앱이 죽지 않도록).
- 한글/이모지는 기존 `clean_text_for_export`로 제거해 Word 호환성 유지.
