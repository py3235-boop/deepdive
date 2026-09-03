# deepdive

팀 딥다이브 — 2026 전남·광주 청년 AI 솔버톤

사내 업무를 Google Workspace 기반으로 자동화하는 프로젝트 모음입니다.

## 프로젝트

| 폴더 | 내용 | 상태 |
|---|---|---|
| [`balju-automation`](./balju-automation) | **발주서 자동화** — 고객사 발주서를 형식에 상관없이 읽어 딥다이브 품목코드로 변환하고 월별 시트에 기록 | 구현 완료 |
| [`production-planning`](./production-planning) | 생산 계획 | 준비 중 |
| [`outbound-plan`](./outbound-plan) | 출고 계획 | 준비 중 |

## 공통 스택

- Google 스프레드시트 + Apps Script (스프레드시트 바인딩)
- 배포·관리: [clasp](https://github.com/google/clasp)
- 외부 API·유료 서비스를 쓰지 않는다

각 폴더의 `README.md` 에 설치·사용법이 있습니다.
