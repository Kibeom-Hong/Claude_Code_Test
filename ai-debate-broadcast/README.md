# 독서토론대회 AI 활용 중계 시스템

옹호팀/비판팀이 AI에게 프롬프트를 실행하면, 청중 화면(프로젝터)에 각 팀의
프롬프트와 결과물이 깔끔하게 요약되어 실시간(WebSocket)으로 표시됩니다.

## 1. 설치

```bash
pip install -r requirements.txt
```

## 2. API 키 설정 (선택)

```bash
export ANTHROPIC_API_KEY="발급받은 키"   # Windows는 set ANTHROPIC_API_KEY=...
```

- 키가 있으면 Claude(claude-sonnet-4-6)가 팀 프롬프트에 실제로 답하고 요약도 생성합니다.
- 키가 없으면 목업 모드로 동작하여 키 없이도 전체 흐름을 리허설할 수 있습니다.

## 3. 실행

```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

## 4. 대회장 접속 구성

노트북(서버)과 팀 기기, 프로젝터 PC가 같은 와이파이에 있으면 됩니다.
노트북의 IP를 확인한 뒤(`ipconfig` 또는 `ifconfig`) 아래 주소로 접속합니다.

| 화면 | 주소 | 용도 |
|---|---|---|
| 참가팀 화면 | `http://<노트북IP>:8000/team` | 논제 설정, 팀 선택, 프롬프트 실행, 결과 전문 확인 |
| 청중 화면 | `http://<노트북IP>:8000/audience` | 프로젝터에 전체화면으로 표시 |

- 참가팀 화면은 옹호팀·비판팀 기기에서 각각 열고, 상단에서 자기 팀을 선택해 사용합니다.
- 논제는 어느 참가팀 화면에서든 설정하면 모든 화면에 즉시 반영됩니다.
- 청중 화면은 옹호팀(좌)·비판팀(우) 2열로 요약 카드가 최신순으로 쌓입니다.

## 5. 구조

```
main.py              FastAPI 서버 (라우팅, WebSocket 브로드캐스트, 백그라운드 요약)
ai.py                AI 호출 (결과물 생성 + 청중용 JSON 요약, 목업 폴백 포함)
templates/team.html      참가팀 화면
templates/audience.html  청중 화면
```

상태는 서버 메모리에만 저장되므로 서버를 재시작하면 기록이 초기화됩니다.
대회 기록을 남기려면 종료 전 `http://<노트북IP>:8000/api/state` 를 저장해 두세요.

## 6. 외부 인터넷 공개가 필요한 경우

같은 와이파이가 아닌 곳에서도 접속해야 한다면 다음 중 하나를 사용하세요.

- 임시 터널: `ngrok http 8000` 또는 `cloudflared tunnel --url http://localhost:8000`
- 클라우드 배포: Render, Railway, Cloud Run 등에 이 폴더를 그대로 배포
  (시작 명령 `uvicorn main:app --host 0.0.0.0 --port $PORT`)
