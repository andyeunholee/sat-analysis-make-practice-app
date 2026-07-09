# 코딩 계획: 분석 리포트 Word 다운로드 버튼 추가

## 배경
- 현재 분석 리포트는 화면에만 표시됨
- 사용자가 Word 파일로 저장하고 싶음
- 다행히 `convert_markdown_to_docx()` 함수는 이미 존재 ([app.py:379](app.py#L379))

## 수정 위치
[app.py:882-889](app.py#L882-L889) "Display Report" 섹션

## 구현 계획

### 1. 다운로드 버튼 추가
"📄 View Full Analysis Report" expander 위에 다운로드 버튼 영역 추가:

- **Word 다운로드 버튼** (`📥 Download as Word`): 주 요청
- **PDF 다운로드 버튼** (`📥 Download as PDF`): 보너스 (이미 `convert_markdown_to_pdf()` 존재)
- 2개 컬럼으로 나란히 배치

### 2. 파일명
- 기본값: `SAT_Analysis_Report.docx` / `SAT_Analysis_Report.pdf`
- 가독성 위해 날짜 추가: `SAT_Analysis_Report_YYYYMMDD.docx`

### 3. 변환 처리
- JSON 블록 제거된 `clean_report` 사용 (이미 처리됨)
- `convert_markdown_to_docx(clean_report)` 호출 → bytes 반환
- `st.download_button`의 `data` 파라미터에 직접 전달

### 4. 사용자 경험
- 버튼 클릭 시 즉시 다운로드 시작
- 변환 실패 시 에러 메시지 표시

## 기대 효과
- 한 번의 클릭으로 보기 좋은 Word 문서 저장
- 학부모/학생에게 공유 가능한 포맷
