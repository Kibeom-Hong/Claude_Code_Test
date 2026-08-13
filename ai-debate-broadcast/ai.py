"""AI 호출 모듈 (Claude).

ANTHROPIC_API_KEY 환경변수가 있으면 Claude를 사용하고,
없으면 데모용 목업 응답으로 동작한다 (키 없이도 전체 흐름 시연 가능).
"""
import os
import json
import re

API_KEY = os.getenv("ANTHROPIC_API_KEY")
MOCK_MODE = not API_KEY

_client = None
if not MOCK_MODE:
    from anthropic import Anthropic
    _client = Anthropic()

MODEL = "claude-sonnet-4-6"

ANSWER_SYSTEM = (
    "당신은 독서토론대회 참가 팀의 리서치를 돕는 AI입니다. "
    "팀이 입력한 프롬프트에 정확하고 간결하게 답하세요. "
    "한국어로, 핵심 위주로 답하되 필요한 근거를 갖추세요."
)

SUMMARY_SYSTEM = (
    "당신은 독서토론대회 중계 화면을 위한 요약 담당자입니다. "
    "참가 팀이 AI에게 입력한 프롬프트와 AI의 결과물을 받아, "
    "청중이 한눈에 이해할 수 있도록 요약합니다. 과장이나 평가 없이 중립적이고 간결하게 작성하세요. "
    "반드시 아래 키를 가진 JSON 객체 하나만 출력하세요. 마크다운 백틱, 설명, 다른 텍스트를 절대 붙이지 마세요.\n"
    '{"prompt_summary": "팀이 AI에게 무엇을 어떻게 요청했는지 한 문장으로", '
    '"output_points": ["결과물의 핵심 내용 요점 (2~3개, 각각 한 문장)"]}'
)


def _gen(system: str, user: str) -> str:
    resp = _client.messages.create(
        model=MODEL,
        max_tokens=1024,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return "".join(b.text for b in resp.content if b.type == "text")


def generate_answer(topic: str, prompt: str) -> str:
    """팀 프롬프트에 대한 AI 결과물 생성."""
    if MOCK_MODE:
        return (
            f"[목업 응답 — ANTHROPIC_API_KEY를 설정하면 실제 Claude가 답합니다]\n\n"
            f"'{prompt[:40]}' 요청에 대한 예시 결과물입니다. "
            f"작품의 서사 구조와 인물의 선택을 근거로 두 가지 논점을 제시할 수 있습니다. "
            f"첫째, 인물의 행위는 사회적 규범과의 긴장 속에서 해석되어야 합니다. "
            f"둘째, 결말부의 상징은 해당 논점을 뒷받침하는 핵심 장치로 볼 수 있습니다."
        )
    topic_line = f'현재 토론 논제는 "{topic}"입니다.\n' if topic else ""
    return _gen(ANSWER_SYSTEM, f"{topic_line}[팀 프롬프트]\n{prompt}")


def summarize(team_label: str, prompt: str, output: str) -> dict:
    """청중 화면용 프롬프트·결과물 요약 (JSON)."""
    fallback = {
        "prompt_summary": f"{team_label}이 AI에게 논거 자료를 요청했습니다.",
        "output_points": ["요약 생성에 실패해 원문을 참고해 주세요."],
    }
    if MOCK_MODE:
        return {
            "prompt_summary": f"{team_label}이 '{prompt[:30]}...' 관점의 논거를 요청했습니다.",
            "output_points": [
                "인물의 행위를 사회적 규범과의 긴장 관계로 해석해야 한다는 논점을 제시했습니다.",
                "결말부의 상징을 핵심 근거 장치로 지목했습니다.",
            ],
        }
    try:
        raw = _gen(
            SUMMARY_SYSTEM,
            f"팀: {team_label}\n\n[팀이 입력한 프롬프트]\n{prompt}\n\n[AI 결과물]\n{output}\n\n청중용 요약 JSON을 작성하세요.",
        )
        clean = re.sub(r"```json|```", "", raw).strip()
        data = json.loads(clean)
        if "prompt_summary" in data and isinstance(data.get("output_points"), list):
            return data
        return fallback
    except Exception:
        return fallback
