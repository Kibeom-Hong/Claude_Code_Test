# Basics of VLA (Vision-Language-Action) — 2시간 강의 자료

| 파일 | 내용 | 언제 보나 |
|---|---|---|
| [`01_강의계획_및_준비사항.md`](01_강의계획_및_준비사항.md) | 분 단위 타임테이블, 오늘 밤 체크리스트, 청중별 조절, 운영 팁, 영상 목록 | **가장 먼저** |
| [`02_강의노트_개념편.md`](02_강의노트_개념편.md) | Part 1–2 개념 설명 (VLA 정의, 행동 공간, 모방학습 함정, 행동 출력 4가지 방식, 액션 청킹, FAST, 이중 시스템) + 용어 사전 | 강의 전 정독 |
| [`03_논문해설.md`](03_논문해설.md) | Part 3 논문 14편 쉬운 해설 (RT-1 → RT-2 → OXE → Octo → OpenVLA/OFT → π0 → FAST → π0.5/KI → Helix → GR00T → Gemini Robotics → π*0.6 → π0.7) + 비교표 | 강의 전 정독 (★ 4편 우선) |
| [`04_예상질문_퀴즈.md`](04_예상질문_퀴즈.md) | 예상 질문 Top 10 + 짧은 답 모음 + 강의 중 퀴즈 | 강의 전날 소리 내어 연습 |
| [`05_참고자료.md`](05_참고자료.md) | 논문·코드·영상 링크 (청중 공유용) | 강의 후 공유 |
| [`slides/VLA_Basics_Lecture.pptx`](slides/VLA_Basics_Lecture.pptx) | **56장 슬라이드, 모든 장에 발표자 노트(대본 + 시간표시)** | 발표용 |
| [`slides/VLA_Basics_Lecture.pdf`](slides/VLA_Basics_Lecture.pdf) | PDF 버전 (배포용, 폰트는 대체 폰트로 렌더링됨) | 공유용 |
| [`slides/build_slides.js`](slides/build_slides.js) | 슬라이드 생성 스크립트 (수정 후 `node build_slides.js`로 재생성) | 수정할 때 |

## 강의 구성 (120분)
1. 오프닝 (5분) — 로드맵, 학습목표, π0.5 영상 훅
2. **Part 1. 왜 VLA인가** (20분) — 전통 파이프라인 한계, 모라벡 역설, LLM 레시피, VLA 정의, 행동 공간, 데이터 격차, 모방학습의 두 함정
3. **Part 2. VLA의 구성요소** (25분) — 해부도, Vision Encoder, VLM 백본, **행동 출력 4가지 방식**, 액션 청킹, FAST, 학습 레시피, 단일 vs 이중 시스템
4. 휴식 (10분)
5. **Part 3. 논문으로 보는 VLA의 진화** (45분) — RT-1부터 π0.7(2026)까지
6. **Part 4. 평가·한계·미래** (10분) + 요약 & Q&A (5분)

## 사용 전 할 일
- 슬라이드 1장 표지의 `강의자 이름 · 소속`을 수정하세요.
- 발표할 때 PowerPoint **발표자 보기**를 켜면 각 슬라이드 노트(대본)가 보입니다.
- 폰트는 `Malgun Gothic`(맑은 고딕) 기준입니다. Mac에서는 자동으로 대체 폰트가 적용됩니다.

## 슬라이드 재생성
```bash
cd slides
npm install pptxgenjs
node build_slides.js
```
