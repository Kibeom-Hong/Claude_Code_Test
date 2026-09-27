# VLA 특론 4주차 — 모방학습과 세 개의 정책 (BeT · Diffusion Policy · ACT)

원래 4주차 강의안을 바탕으로, **140분 수업을 그대로 진행할 수 있는 자료 일체**입니다.

| 파일 | 내용 | 언제 보나 |
|---|---|---|
| [`01_강의안_상세.md`](01_강의안_상세.md) | 분 단위 대본: 말할 것 · 판서 · 슬라이드 번호 · 체크 질문, 자를 곳/늘릴 곳, 칠판 최종 모습 | **가장 먼저** |
| [`02_논문해설.md`](02_논문해설.md) | 3편 심화 해설: 구조 · 손실 · 하이퍼파라미터 · 실험 · ablation · 한계 · 후속, 비교표, 2·3·6주차 연결 문장 | 수업 전 정독 |
| [`03_예상질문_퀴즈_과제.md`](03_예상질문_퀴즈_과제.md) | 예상 질문 14개와 답, 수업 중 퀴즈 5개 + O/X, 리딩 리포트 양식과 예시, 데모 변형 과제 | 전날 연습 / 과제 공지 |
| [`slides/Week04_Imitation_Learning.pptx`](slides/Week04_Imitation_Learning.pptx) | **40장, 모든 장에 발표자 노트(대본 + 시간 표시)** | 수업용 |
| [`slides/Week04_Imitation_Learning.pdf`](slides/Week04_Imitation_Learning.pdf) | 배포용 PDF (대체 폰트로 렌더링) | 학생 공유 |
| [`demo/multimodality_demo.py`](demo/multimodality_demo.py) | MSE 회귀 vs BeT식 이산화 vs Diffusion을 같은 데이터로 비교 (numpy · matplotlib만, CPU 약 40초) | 슬라이드 14 / 과제 |
| [`demo/multimodality_demo.png`](demo/multimodality_demo.png) | 데모 결과 그림 — 장애물 구간에 떨어진 비율 MSE 100% · BeT 0% · Diffusion 약 3% | 슬라이드 14 |

## 수업 흐름 (140분)
| 시간 | 블록 | 슬라이드 |
|---|---|---|
| 0:00–0:10 | 복습 + 용어 정의 (정책 · 시범 · BC) | 1–5 |
| 0:10–0:35 | 적 1호: 누적 오차 (운전 교습, εT², DAgger) | 6–10 |
| 0:35–0:55 | 적 2호: 다봉성 + 시간 축 갈팡질팡 + 데모 + action chunking | 11–16 |
| 0:55–1:05 | 휴식 (퀴즈) | 17 |
| 1:05–1:25 | 해법 1 — BeT | 18–22 |
| 1:25–1:50 | 해법 2 — Diffusion Policy | 23–29 |
| 1:50–2:10 | 해법 3 — ACT | 30–36 |
| 2:10–2:20 | 비교표 · O/X · 다음 주 예고 | 37–40 |

## 수업 전 준비
- 영상 두 개: ALOHA/ACT (tonyzhaozh.github.io/aloha), Diffusion Policy (diffusion-policy.cs.columbia.edu)
- 칠판을 세 구역으로: 왼쪽 "두 개의 적" / 가운데 "세 해법" / 오른쪽 "action chunking 박스"
- 색 약속: BeT = 보라, Diffusion Policy = 주황, ACT = 청록, 적 = 빨강

## 재생성
```bash
cd demo && python multimodality_demo.py          # 그림 재생성
cd ../slides && npm install pptxgenjs && node build_week04.js
```
