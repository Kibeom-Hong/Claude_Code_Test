// VLA 기초 강의 슬라이드 생성 스크립트
// 실행: npm install pptxgenjs && node build_slides.js  → VLA_Basics_Lecture.pptx
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Basics of Vision-Language-Action Models";

// ---------- 디자인 토큰 ----------
const F = "Malgun Gothic";
const MONO = "Courier New";
const C = {
  ink: "14171C", ink2: "232932", ink3: "2E3540",
  text: "1A1D23", muted: "5F6773", faint: "9AA1AB",
  line: "D9DDE3", card: "F3F4F6", white: "FFFFFF",
  V: "0E9F9A", L: "7C5CFF", A: "F2651F",
  Vt: "E1F5F4", Lt: "EEEAFF", At: "FFEBDD",
  good: "1E9E5A", bad: "D64545",
};
const W = 13.333, H = 7.5;
let pageNo = 0;

// ---------- 헬퍼 ----------
function T(s, text, o) {
  s.addText(text, Object.assign({ fontFace: F, color: C.text, fontSize: 14, margin: 0, valign: "top", isTextBox: true }, o));
}
function box(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({
    x, y, w, h, fontFace: F, fontSize: 14, color: C.text, align: "center", valign: "middle",
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, fill: { color: C.card },
    margin: 0.08, isTextBox: true,
  }, o));
}
function arrow(s, x1, y1, x2, y2, color = C.faint, width = 2) {
  const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    line: { color, width, endArrowType: "triangle" } };
  if (x2 < x1) o.flipH = true;
  if (y2 < y1) o.flipV = true;
  s.addShape(pres.shapes.LINE, o);
}
function dot(s, x, y, d, color) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color } });
}
function vlaDots(s, x, y) { dot(s, x, y, 0.13, C.V); dot(s, x + 0.19, y, 0.13, C.L); dot(s, x + 0.38, y, 0.13, C.A); }
function numCircle(s, n, x, y, d, color, fs = 16) {
  s.addText(String(n), { x, y, w: d, h: d, shape: pres.shapes.OVAL, fill: { color }, line: { color }, color: C.white,
    fontFace: F, fontSize: fs, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function bullets(s, items, x, y, w, h, fs = 16, color = C.text) {
  const runs = items.map((it, i) => {
    const isObj = typeof it === "object" && !Array.isArray(it);
    const txt = isObj ? it.t : it;
    const opts = { bullet: isObj && it.sub ? { indent: 18 } : true, indentLevel: isObj && it.sub ? 1 : 0,
      fontSize: isObj && it.sub ? fs - 2 : fs, color: isObj && it.sub ? C.muted : color, paraSpaceAfter: 6 };
    if (isObj && it.b) opts.bold = true;
    if (i < items.length - 1) opts.breakLine = true;
    return { text: txt, options: opts };
  });
  s.addText(runs, { x, y, w, h, fontFace: F, valign: "top", margin: 0, isTextBox: true });
}
function content(title, section) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: C.white };
  vlaDots(s, 0.6, 0.42);
  T(s, section, { x: 1.2, y: 0.33, w: 8, h: 0.3, fontSize: 11, color: C.muted, bold: true, charSpacing: 1 });
  T(s, title, { x: 0.6, y: 0.68, w: 12.1, h: 0.75, fontSize: 28, bold: true, valign: "middle" });
  T(s, String(pageNo), { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.faint, align: "right" });
  return s;
}
function dark() {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: C.ink };
  return s;
}
function section(num, title, sub, minutes, notes) {
  const s = dark();
  T(s, `PART ${num}`, { x: 0.8, y: 2.1, w: 6, h: 0.5, fontSize: 18, color: C.A, bold: true, charSpacing: 3 });
  T(s, title, { x: 0.8, y: 2.65, w: 11.5, h: 1.2, fontSize: 44, color: C.white, bold: true, valign: "middle" });
  T(s, sub, { x: 0.8, y: 3.95, w: 11, h: 0.9, fontSize: 18, color: "B8C0CC" });
  vlaDots(s, 0.8, 5.3);
  T(s, minutes, { x: 1.4, y: 5.22, w: 6, h: 0.3, fontSize: 13, color: "8A93A0" });
  s.addNotes(notes);
  return s;
}
function card(s, x, y, w, h, head, body, accent, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: o.fill || C.card }, line: { color: o.fill || C.card } });
  T(s, head, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.45, fontSize: o.hs || 17, bold: true, color: accent || C.text, valign: "middle" });
  if (typeof body === "string") T(s, body, { x: x + 0.25, y: y + 0.72, w: w - 0.5, h: h - 0.9, fontSize: o.bs || 13.5, color: o.bc || C.text, lineSpacingMultiple: 1.15 });
  else bullets(s, body, x + 0.25, y + 0.72, w - 0.5, h - 0.9, o.bs || 13.5);
}
function stat(s, x, y, w, big, label, color) {
  T(s, big, { x, y, w, h: 0.9, fontSize: 40, bold: true, color, valign: "bottom" });
  T(s, label, { x, y: y + 0.95, w, h: 0.7, fontSize: 13, color: C.muted });
}
function paperTag(s, text, x = 9.3, y = 0.36, w = 3.45) {
  T(s, text, { x, y, w, h: 0.3, fontSize: 11, color: C.muted, align: "right" });
}
function code(s, text, x, y, w, h, fs = 15) {
  s.addText(text, { x, y, w, h, fontFace: MONO, fontSize: fs, color: "E8EAED", fill: { color: C.ink2 },
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, margin: 0.2, valign: "middle", isTextBox: true });
}
function tbl(s, rows, o) {
  const data = rows.map((r, ri) => r.map((cell) => ({
    text: String(cell),
    options: ri === 0 ? { bold: true, color: C.white, fill: { color: C.ink2 } } : { color: C.text, fill: { color: ri % 2 ? C.white : "F7F8FA" } },
  })));
  s.addTable(data, Object.assign({ fontFace: F, fontSize: 12, border: { type: "solid", pt: 0.5, color: C.line }, valign: "middle", margin: 0.06 }, o));
}

// =====================================================================
// 1. 표지
// =====================================================================
{
  const s = dark();
  vlaDots(s, 0.8, 1.2);
  T(s, "Basics of", { x: 0.8, y: 1.9, w: 10, h: 0.7, fontSize: 28, color: "B8C0CC" });
  s.addText([
    { text: "Vision", options: { color: C.V } }, { text: "-", options: { color: C.white } },
    { text: "Language", options: { color: C.L } }, { text: "-", options: { color: C.white } },
    { text: "Action", options: { color: C.A, breakLine: true } }, { text: "Models", options: { color: C.white } },
  ], { x: 0.8, y: 2.5, w: 12, h: 1.9, fontFace: F, fontSize: 48, bold: true, margin: 0, valign: "middle", isTextBox: true });
  T(s, "보고, 이해하고, 움직이는 로봇 AI — RT-2부터 π0.5, 그리고 그 너머까지", { x: 0.8, y: 4.6, w: 12, h: 0.6, fontSize: 20, color: "D7DCE3" });
  T(s, "강의자 이름  ·  소속  ·  2026", { x: 0.8, y: 5.9, w: 8, h: 0.4, fontSize: 14, color: "8A93A0" });
  s.addNotes(`[0:00–0:02] 인사 및 자기소개.
"안녕하세요, 오늘은 2시간 동안 Vision-Language-Action, 줄여서 VLA 모델의 기초를 다룹니다."
- 제목의 색을 가리키며: "청록은 Vision, 보라는 Language, 주황은 Action입니다. 오늘 모든 슬라이드에서 이 색 약속을 지킵니다."
- (준비) 표지의 '강의자 이름 · 소속'을 본인 정보로 바꾸세요.`);
}

// 2. 로드맵
{
  const s = content("오늘의 로드맵 (120분)", "OPENING");
  const parts = [
    ["PART 1", "왜 VLA인가", "20분", C.V],
    ["PART 2", "VLA의 구성요소", "25분", C.L],
    ["BREAK", "휴식", "10분", C.faint],
    ["PART 3", "논문으로 보는\nVLA의 진화", "45분", C.A],
    ["PART 4", "평가 · 한계 · 미래", "15분", C.ink3],
  ];
  const x0 = 0.6, gw = 2.3, gap = 0.155;
  parts.forEach((p, i) => {
    const x = x0 + i * (gw + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.0, w: gw, h: 3.2, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
    dot(s, x + 0.25, 2.3, 0.28, p[3]);
    T(s, p[0], { x: x + 0.25, y: 2.8, w: gw - 0.4, h: 0.35, fontSize: 12, bold: true, color: C.muted });
    T(s, p[1], { x: x + 0.25, y: 3.2, w: gw - 0.4, h: 1.0, fontSize: 18, bold: true });
    T(s, p[2], { x: x + 0.25, y: 4.45, w: gw - 0.4, h: 0.5, fontSize: 24, bold: true, color: p[3] === C.faint ? C.muted : p[3] });
  });
  T(s, "+ 마무리 요약 & Q&A 5분", { x: 0.6, y: 5.55, w: 8, h: 0.4, fontSize: 14, color: C.muted });
  bullets(s, [
    "Part 1–2: 개념과 원리 (무엇을, 왜)",
    "Part 3: 대표 논문 12편으로 흐름 잡기 (어떻게 발전했나)",
  ], 0.6, 6.0, 12, 0.9, 14, C.muted);
  s.addNotes(`[0:02–0:03] "오늘은 4개 파트로 진행합니다. 앞의 절반은 개념, 휴식 후 절반은 논문입니다."
- Part 1: 왜 로봇에 이런 모델이 필요한지
- Part 2: VLA가 어떤 부품으로 되어 있는지 — 특히 '행동을 어떻게 출력하느냐'가 핵심
- Part 3: RT-1부터 2026년 π0.7까지 논문 흐름
- Part 4: 어떻게 평가하고, 무엇이 남았는지
"질문은 언제든 환영하지만, 긴 질문은 휴식 시간이나 마지막에 받겠습니다."`);
}

// 3. 학습목표
{
  const s = content("오늘이 끝나면 여러분은…", "OPENING");
  const goals = [
    ["VLA가 무엇인지", "이미지 + 언어 + 로봇 상태 → 행동.\n한 문장으로 설명할 수 있다.", C.V],
    ["행동을 출력하는 4가지 방식", "이산 토큰 · 연속 회귀 · 디퓨전/플로우 · 압축 토큰(FAST)의 차이를 안다.", C.L],
    ["대표 논문의 계보", "RT-2 → OpenVLA → π0 → π0.5 → 이중 시스템 → RL까지 흐름을 설명할 수 있다.", C.A],
  ];
  goals.forEach((g, i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.0, w: 3.8, h: 3.9, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
    numCircle(s, i + 1, x + 0.3, 2.3, 0.7, g[2], 22);
    T(s, g[0], { x: x + 0.3, y: 3.25, w: 3.2, h: 0.9, fontSize: 20, bold: true });
    T(s, g[1], { x: x + 0.3, y: 4.2, w: 3.2, h: 1.5, fontSize: 14.5, color: C.muted, lineSpacingMultiple: 1.2 });
  });
  s.addNotes(`[0:03] 학습목표 3가지를 읽어줍니다.
"강의가 끝났을 때 이 3가지를 말할 수 있으면 성공입니다. 특히 2번, 행동 출력 방식은 VLA 논문을 읽을 때 가장 먼저 봐야 할 포인트입니다."`);
}

// 4. 훅
{
  const s = content("로봇에게 \"식탁 좀 치워줘\"라고 말하면?", "OPENING");
  T(s, "“식탁 좀 치워줘”", { x: 0.6, y: 1.9, w: 5.6, h: 1.0, fontSize: 36, bold: true, color: C.L });
  bullets(s, [
    "어떤 게 쓰레기고 어떤 게 그릇인지 알아야 하고 (Vision + 상식)",
    "\"치운다\"가 무슨 뜻인지 이해해야 하고 (Language)",
    "처음 보는 컵도 깨지지 않게 집어야 한다 (Action)",
  ], 0.6, 3.2, 5.7, 2.6, 17);
  const cx = [7.0, 9.0, 11.0];
  const lab = [["V", "보고", C.V], ["L", "이해하고", C.L], ["A", "움직인다", C.A]];
  lab.forEach((l, i) => {
    s.addText(l[0], { x: cx[i], y: 2.6, w: 1.5, h: 1.5, shape: pres.shapes.OVAL, fill: { color: l[2] }, line: { color: l[2] },
      color: C.white, fontFace: F, fontSize: 40, bold: true, align: "center", valign: "middle", isTextBox: true });
    T(s, l[1], { x: cx[i] - 0.2, y: 4.25, w: 1.9, h: 0.5, fontSize: 18, bold: true, align: "center" });
    if (i < 2) arrow(s, cx[i] + 1.55, 3.35, cx[i + 1] - 0.05, 3.35, C.faint, 2.5);
  });
  T(s, "이 세 가지를 하나의 모델로 → VLA", { x: 7.0, y: 5.3, w: 5.6, h: 0.5, fontSize: 18, bold: true, align: "center", color: C.text });
  T(s, "▶ 영상: π0.5 — 처음 보는 집 청소 (pi.website/blog/pi05)", { x: 0.6, y: 6.4, w: 10, h: 0.35, fontSize: 12, color: C.muted });
  s.addNotes(`[0:03–0:05] ★ 여기서 π0.5 영상을 30초 정도 틀어주세요 (pi.website/blog/pi05).
"이 로봇은 이 집에 처음 왔습니다. 이 집 데이터로 학습한 적이 없습니다. 그런데 '부엌 정리해'라는 말만 듣고 접시를 싱크대에 넣고, 쓰레기를 버립니다."
"이게 가능하려면 세 가지가 필요합니다. 보고(V), 이해하고(L), 움직이는(A) 것. 이걸 하나의 신경망으로 해내는 게 오늘의 주제, VLA입니다."`);
}

// =====================================================================
// PART 1
// =====================================================================
section(1, "왜 VLA인가", "LLM은 시를 쓰는데, 왜 로봇은 아직 빨래를 못 갤까?", "약 20분 · 슬라이드 6–13",
  `[0:05] Part 1 시작. 질문을 던지고 잠깐 멈추세요: "LLM은 시를 쓰는데, 왜 우리 집엔 빨래 개는 로봇이 없을까요?"`);

// 6. 전통 파이프라인
{
  const s = content("전통적인 로봇: 사람이 설계한 모듈의 파이프라인", "PART 1 · 왜 VLA인가");
  const mods = ["카메라", "인식\n(물체 검출)", "상태 추정\n(위치·지도)", "계획\n(경로·잡기)", "제어\n(모터)"];
  mods.forEach((m, i) => {
    const x = 0.6 + i * 2.5;
    box(s, m, x, 2.0, 2.0, 1.0, { fontSize: 14, bold: true, fill: { color: i === 0 ? C.Vt : C.card } });
    if (i < mods.length - 1) arrow(s, x + 2.02, 2.5, x + 2.48, 2.5);
  });
  T(s, "각 모듈을 사람이 설계 · 튜닝", { x: 0.6, y: 3.1, w: 12, h: 0.35, fontSize: 12, color: C.muted, align: "center" });
  card(s, 0.6, 3.8, 5.9, 2.3, "장점", [
    "공장처럼 고정된 환경에서 정확하고 빠름",
    "각 단계를 해석·디버깅하기 쉬움",
  ], C.good);
  card(s, 6.8, 3.8, 5.9, 2.3, "한계", [
    "새 물체·새 환경마다 다시 설계 (컵 → 비닐봉지? 수건?)",
    "모듈 간 오류 누적 (인식 1cm 오차 → 잡기 실패)",
    "\"식탁 치워줘\" 같은 애매한 언어를 이해 못 함",
  ], C.bad);
  s.addNotes(`[0:06] "고전적인 로봇은 이렇게 모듈을 사람이 이어 붙입니다. 공장에서는 아주 잘 됩니다. 같은 부품, 같은 위치니까요."
"그런데 집은 다릅니다. 매번 물건이 다르고, 수건처럼 모양이 계속 변하는 물체도 있죠. 새로운 상황마다 사람이 다시 설계해야 합니다."
"그래서 '데이터로부터 통째로(end-to-end) 배우자'는 흐름이 나왔고, 그 최신 형태가 VLA입니다."`);
}

// 7. 모라벡 역설
{
  const s = content("모라벡의 역설: 사람에게 쉬운 일이 AI에겐 어렵다", "PART 1 · 왜 VLA인가");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.9, w: 5.9, h: 4.6, rectRadius: 0.08, fill: { color: C.Lt }, line: { color: C.Lt } });
  T(s, "AI가 이미 해낸 것", { x: 0.9, y: 2.1, w: 5.3, h: 0.5, fontSize: 18, bold: true, color: C.L });
  [["1997", "체스 세계 챔피언 격파"], ["2016", "바둑 (AlphaGo)"], ["2022~", "시·코드·논문 작성 (LLM)"]].forEach((r, i) => {
    T(s, r[0], { x: 0.9, y: 2.85 + i * 1.1, w: 1.6, h: 0.8, fontSize: 30, bold: true, color: C.L, valign: "middle" });
    T(s, r[1], { x: 2.6, y: 2.85 + i * 1.1, w: 3.7, h: 0.8, fontSize: 16, valign: "middle" });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: 1.9, w: 5.9, h: 4.6, rectRadius: 0.08, fill: { color: C.At }, line: { color: C.At } });
  T(s, "아직 어려운 것", { x: 7.1, y: 2.1, w: 5.3, h: 0.5, fontSize: 18, bold: true, color: C.A });
  bullets(s, ["빨래 개기 (모양이 계속 변함)", "처음 보는 집에서 설거지", "계란을 깨지 않고 집기", "이 모든 걸 하나의 모델로"], 7.1, 2.85, 5.3, 2.2, 17);
  T(s, "이유: 물리 세계는 연속적 · 예측 어려움 · 데이터 비쌈 · 실수 비용 큼", { x: 7.1, y: 5.3, w: 5.3, h: 1.0, fontSize: 14, color: C.muted, bold: true });
  s.addNotes(`[0:08] 모라벡의 역설: "컴퓨터에게 어려운 일(체스)은 쉽고, 사람에게 쉬운 일(걷기, 집기)은 어렵다."
"AI는 1997년에 체스 챔피언을 이겼는데, 빨래 개는 로봇은 2024년이 되어서야 데모 수준으로 나왔습니다."
이유 4가지를 짚기: 연속적인 물리 세계, 예측 불가(천은 모양이 변함), 데이터가 비쌈, 실수하면 컵이 깨짐.`);
}

// 8. LLM 레시피
{
  const s = content("LLM의 성공 레시피를 로봇에도 쓸 수 있을까?", "PART 1 · 왜 VLA인가");
  const steps = [["대규모 사전학습", "인터넷 텍스트 수조 토큰"], ["사후학습", "지시 따르기 · 대화 · RL"], ["범용 모델", "하나의 모델로 수많은 작업"]];
  steps.forEach((st, i) => {
    const x = 0.6 + i * 4.2;
    box(s, [{ text: st[0], options: { bold: true, fontSize: 18, breakLine: true } }, { text: st[1], options: { fontSize: 13, color: C.muted } }], x, 1.95, 3.6, 1.3, { fill: { color: C.Lt } });
    if (i < 2) arrow(s, x + 3.65, 2.6, x + 4.15, 2.6, C.L);
  });
  T(s, "LLM", { x: 0.6, y: 1.6, w: 2, h: 0.3, fontSize: 12, bold: true, color: C.L });
  T(s, "로봇에 그대로 적용하려면 두 가지 벽이 있다", { x: 0.6, y: 3.65, w: 12, h: 0.5, fontSize: 18, bold: true });
  card(s, 0.6, 4.3, 5.9, 1.7, "벽 1 · 데이터", "로봇 데이터는 인터넷에 공짜로 쌓여 있지 않다. 사람이 직접 조종해서 모아야 한다.", C.A);
  card(s, 6.8, 4.3, 5.9, 1.7, "벽 2 · 출력", "로봇의 출력은 단어가 아니라 연속적인 모터 명령이다.", C.A);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.2, w: 12.1, h: 0.75, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "아이디어: 이미 인터넷으로 학습된 VLM을 가져와서, 로봇 행동만 추가로 가르치자!  →  VLA", { x: 0.85, y: 6.2, w: 11.7, h: 0.75, fontSize: 16, bold: true, color: C.white, valign: "middle" });
  s.addNotes(`[0:10] "LLM은 사전학습 → 사후학습 → 범용 모델이라는 레시피로 성공했습니다. 로봇 연구자들은 '이걸 로봇에도?'라고 생각했죠."
"벽이 두 개 있습니다. 첫째, 데이터. 인터넷 텍스트처럼 공짜로 쌓여있지 않습니다. 둘째, 출력. 로봇은 단어가 아니라 모터 명령을 내야 합니다."
"그래서 나온 아이디어: 인터넷으로 이미 세상을 배운 VLM을 가져와서, '행동'만 추가로 가르치자. 이게 VLA입니다."
비유: "책으로 세상을 다 배운 사람에게 운전 연수만 시키는 것."`);
}

// 9. VLA 정의
{
  const s = content("VLA의 정의", "PART 1 · 왜 VLA인가");
  box(s, "카메라 이미지\n(머리 · 손목 카메라)", 0.6, 1.9, 2.9, 1.0, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 14 });
  box(s, "언어 지시\n\"put the cup in the sink\"", 0.6, 3.1, 2.9, 1.0, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 14 });
  box(s, "로봇 상태\n(관절 각도 · 그리퍼)", 0.6, 4.3, 2.9, 1.0, { fill: { color: C.card }, color: C.muted, bold: true, fontSize: 14 });
  [2.4, 3.6, 4.8].forEach((y) => arrow(s, 3.55, y, 4.45, 3.6));
  box(s, "VLA\n(사전학습된 VLM\n+ 행동 출력부)", 4.5, 2.6, 3.0, 2.0, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 17 });
  arrow(s, 7.55, 3.6, 8.35, 3.6, C.A, 2.5);
  box(s, "로봇 행동\n(다음 H스텝의 움직임)", 8.4, 3.0, 2.9, 1.2, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 15 });
  code(s, "π(a[t:t+H] | 이미지, 언어, 상태)", 0.6, 5.65, 6.9, 0.7, 16);
  T(s, "용어의 기원: RT-2 (Google DeepMind, 2023)", { x: 7.9, y: 5.75, w: 4.9, h: 0.5, fontSize: 13, color: C.muted, valign: "middle" });
  T(s, "보통 사전학습된 VLM 위에 만든다 → 인터넷 지식을 로봇이 활용", { x: 0.6, y: 6.5, w: 12, h: 0.4, fontSize: 14, bold: true });
  s.addNotes(`[0:12] 정의: "카메라 이미지와 자연어 지시를 받아서, 로봇 행동을 직접 출력하는 end-to-end 모델. 보통 사전학습된 VLM 위에 만든다."
- 입력 3가지: 이미지(1~3대 카메라), 언어 지시, 로봇 상태(proprioception).
- 출력: 행동. H=1이면 한 스텝, H>1이면 '액션 청크' (Part 2에서 설명).
- π는 policy(정책). 로봇 분야에서 '관측 → 행동' 함수를 policy라고 부릅니다.
- 이름은 2023년 RT-2 논문에서 처음 썼습니다.
"한 문장으로: 눈과 언어를 가진 AI에 손을 달아준 것."`);
}

// 10. LLM/VLM/VLA 비교
{
  const s = content("LLM · VLM · VLA 무엇이 다른가", "PART 1 · 왜 VLA인가");
  tbl(s, [
    ["", "LLM", "VLM", "VLA"],
    ["입력", "텍스트", "이미지 + 텍스트", "이미지 + 텍스트 + 로봇 상태"],
    ["출력", "텍스트", "텍스트", "로봇 행동 (+ 때로 텍스트)"],
    ["예시", "GPT, Llama, Claude", "LLaVA, PaliGemma, Qwen-VL", "RT-2, OpenVLA, π0, GR00T"],
    ["틀리면?", "오답 문장", "오답 문장", "컵이 깨진다 (물리적 결과)"],
    ["학습 데이터", "수조 토큰", "이미지-텍스트 수십억 쌍", "로봇 시연 수십만~수백만 에피소드"],
    ["동작 방식", "한 번 답하면 끝", "한 번 답하면 끝", "닫힌 루프: 행동 → 세상 변화 → 새 관측"],
  ], { x: 0.6, y: 1.9, w: 12.1, colW: [1.8, 2.9, 3.3, 4.1], fontSize: 14, rowH: 0.62 });
  s.addNotes(`[0:14] 표를 한 줄씩 짚되, 마지막 두 줄을 강조:
"VLA는 틀리면 컵이 깨집니다. 그리고 닫힌 루프입니다. 내가 한 행동이 세상을 바꾸고, 바뀐 세상이 다음 입력이 됩니다. 챗봇은 한 번 답하면 끝이지만, 로봇은 자기 실수를 다음 순간에 다시 봐야 합니다."
이 '닫힌 루프' 특성이 뒤에 나올 '누적 오차' 문제로 이어집니다.`);
}

// 11. 행동이란
{
  const s = content("로봇의 \"행동(Action)\"이란?", "PART 1 · 왜 VLA인가");
  T(s, "가장 흔한 표현: 손끝(엔드 이펙터)의 7차원 변화량", { x: 0.6, y: 1.75, w: 6.5, h: 0.4, fontSize: 16, bold: true });
  const dims = [["Δx", "앞뒤"], ["Δy", "좌우"], ["Δz", "위아래"], ["Δroll", "회전"], ["Δpitch", "회전"], ["Δyaw", "회전"], ["grip", "열기/닫기"]];
  dims.forEach((d, i) => {
    const x = 0.6 + i * 0.93;
    box(s, [{ text: d[0], options: { bold: true, fontSize: 14, color: C.A, breakLine: true } }, { text: d[1], options: { fontSize: 10.5, color: C.muted } }], x, 2.3, 0.85, 1.0, { fill: { color: C.At }, margin: 0.02 });
  });
  bullets(s, [
    "다른 표현: 관절 각도(팔 7개 관절), 양팔은 14차원, 휴머노이드는 수십 차원",
    "Embodiment(몸의 형태)마다 행동 공간이 다르다 → 여러 로봇 데이터를 합치기 어려움",
    "제어 주파수(Hz): 1초에 몇 번 행동을 내보내나 — 정교한 작업일수록 높아야 함",
  ], 0.6, 3.65, 6.4, 3.0, 15);
  s.addChart(pres.charts.BAR, [{ name: "제어 주파수 (Hz)", labels: ["RT-1", "OpenVLA", "π0", "GR00T N1 (S1)", "Helix (S1)"], values: [3, 6, 50, 120, 200] }], {
    x: 7.3, y: 1.7, w: 5.4, h: 4.9, barDir: "bar", chartColors: [C.A], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFontSize: 12, dataLabelColor: C.text, catAxisLabelFontSize: 12, catAxisLabelColor: C.text, catAxisLabelFontFace: F,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "모델별 제어 주파수 (Hz, 대략값)", titleFontSize: 13, titleColor: C.muted, titleFontFace: F, barGapWidthPct: 60,
  });
  s.addNotes(`[0:16] "로봇의 행동은 보통 손끝을 얼마나 움직일지를 나타내는 7개 숫자입니다. x,y,z 이동, 3축 회전, 그리고 그리퍼 열고 닫기."
"양팔 로봇이면 14개, 휴머노이드면 수십 개가 됩니다. 로봇마다 이 숫자의 개수와 의미가 달라서, 여러 로봇 데이터를 합치기 어렵습니다. 이걸 embodiment 문제라고 합니다."
오른쪽 차트: "1초에 몇 번 행동하느냐. RT-1은 3번, π0는 50번, Helix의 빠른 시스템은 200번입니다. 빨래 개기처럼 섬세한 작업은 빨라야 합니다. 그런데 큰 모델은 느립니다. 이 긴장관계가 VLA 설계의 핵심입니다."`);
}

// 12. 데이터
{
  const s = content("로봇 데이터는 어디서 오나: 데이터 격차", "PART 1 · 왜 VLA인가");
  stat(s, 0.6, 1.8, 3.9, "~15조", "LLM 사전학습 텍스트 토큰\n(예: Llama 3) — 인터넷에서 사실상 공짜", C.L);
  stat(s, 4.7, 1.8, 3.9, "~100만", "공개 로봇 데이터 최대 규모 궤적 수\n(Open X-Embodiment, 2023)", C.A);
  stat(s, 8.8, 1.8, 3.9, "~1만 시간", "π0 자체 수집 로봇 데이터\n(업계 최대급, 비공개)", C.A);
  T(s, "로봇 데이터 = 사람이 직접 조종(Teleoperation)해서 한 에피소드씩", { x: 0.6, y: 4.0, w: 12, h: 0.45, fontSize: 17, bold: true });
  card(s, 0.6, 4.6, 3.9, 2.1, "리더–팔로워 팔", "ALOHA: 사람이 리더 팔을 움직이면 로봇 팔이 그대로 따라함. 저가·양팔", C.V);
  card(s, 4.7, 4.6, 3.9, 2.1, "VR · 스페이스마우스", "VR 컨트롤러나 3D 마우스로 손끝 위치를 직접 지정", C.V);
  card(s, 8.8, 4.6, 3.9, 2.1, "한 에피소드 =", "(이미지들, 로봇 상태들, 행동들, 언어 지시)의 시계열 기록", C.V);
  s.addNotes(`[0:18] 숫자로 격차를 보여주기. "LLM은 15조 토큰, 로봇은 공개 데이터를 다 모아도 100만 에피소드 정도입니다."
"로봇 데이터는 사람이 한 시간씩 직접 로봇을 조종해서 모아야 합니다. ALOHA 같은 리더-팔로워 시스템, VR 컨트롤러 등을 씁니다."
"그래서 VLA 연구의 큰 줄기는 '적은 로봇 데이터로 어떻게 최대한 일반화하냐'입니다. 답은 크게 세 가지: 인터넷 지식(VLM) 활용, 여러 로봇 데이터 합치기, 시뮬레이션·사람 영상 활용."`);
}

// 13. BC와 두 함정
{
  const s = content("모방학습(Behavior Cloning)과 두 가지 함정", "PART 1 · 왜 VLA인가");
  T(s, "\"이 장면에서 사람은 이렇게 움직였으니, 너도 그렇게 해\" — 시연 데이터로 지도학습", { x: 0.6, y: 1.7, w: 12, h: 0.45, fontSize: 16, bold: true, color: C.muted });
  // 함정 1
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.35, w: 5.9, h: 4.4, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "함정 1 · 누적 오차", { x: 0.85, y: 2.5, w: 5.4, h: 0.45, fontSize: 18, bold: true, color: C.bad });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 3.9, w: 4.9, h: 0, line: { color: C.good, width: 3 } });
  T(s, "사람 시연", { x: 4.9, y: 3.5, w: 1.2, h: 0.3, fontSize: 11, color: C.good, bold: true });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 3.9, w: 1.6, h: 0.15, line: { color: C.bad, width: 3 } });
  s.addShape(pres.shapes.LINE, { x: 2.6, y: 4.05, w: 1.5, h: 0.4, line: { color: C.bad, width: 3 } });
  s.addShape(pres.shapes.LINE, { x: 4.1, y: 4.45, w: 1.6, h: 0.75, line: { color: C.bad, width: 3, endArrowType: "triangle" } });
  T(s, "로봇", { x: 5.2, y: 5.05, w: 1.0, h: 0.3, fontSize: 11, color: C.bad, bold: true });
  T(s, "조금 틀림 → 처음 보는 상황 → 더 틀림 → 눈덩이\n(차선 중앙만 연습한 운전자)", { x: 0.85, y: 5.5, w: 5.4, h: 1.1, fontSize: 13.5, color: C.text });
  // 함정 2
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: 2.35, w: 5.9, h: 4.4, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "함정 2 · 다봉성 (정답이 여러 개)", { x: 7.05, y: 2.5, w: 5.4, h: 0.45, fontSize: 18, bold: true, color: C.bad });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.35, y: 3.3, w: 0.8, h: 0.8, fill: { color: C.ink3 }, line: { color: C.ink3 } });
  T(s, "장애물", { x: 9.15, y: 3.0, w: 1.2, h: 0.3, fontSize: 11, align: "center", color: C.muted });
  dot(s, 9.6, 5.1, 0.3, C.A);
  arrow(s, 9.65, 5.1, 8.6, 3.4, C.good, 2.5);
  arrow(s, 9.85, 5.1, 10.9, 3.4, C.good, 2.5);
  arrow(s, 9.75, 5.05, 9.75, 4.2, C.bad, 3);
  T(s, "왼쪽 50%", { x: 7.3, y: 3.6, w: 1.3, h: 0.3, fontSize: 11, color: C.good, bold: true });
  T(s, "오른쪽 50%", { x: 11.0, y: 3.6, w: 1.5, h: 0.3, fontSize: 11, color: C.good, bold: true });
  T(s, "평균 = 충돌!", { x: 10.1, y: 4.45, w: 1.6, h: 0.3, fontSize: 11, color: C.bad, bold: true });
  T(s, "MSE 회귀로 배우면 두 정답의 평균을 출력\n→ 분포를 표현하는 출력 방식이 필요 (Part 2)", { x: 7.05, y: 5.5, w: 5.4, h: 1.1, fontSize: 13.5 });
  s.addNotes(`[0:20–0:25] 대부분의 VLA는 모방학습(Behavior Cloning)으로 학습합니다. 사람 시연을 정답으로 하는 지도학습.
함정 1, 누적 오차: "차선 중앙만 연습한 운전자는 차선에서 조금 벗어나면 복귀를 못 합니다. 로봇도 조금 틀리면, 학습 때 본 적 없는 상황이 되어 더 틀립니다."
함정 2, 다봉성: ★ 퀴즈로 던지기 — "시연 절반은 왼쪽, 절반은 오른쪽으로 피했습니다. 평균을 내도록 학습하면?" (답: 정면 충돌)
"오늘 볼 논문들의 설계는 대부분 이 두 함정을 피하기 위한 것입니다. 기억해두세요."
→ Part 1 끝. 질문 1~2개 받기.`);
}

// =====================================================================
// PART 2
// =====================================================================
section(2, "VLA의 구성요소", "언어 모델은 '단어'를 출력하는데, 로봇 팔의 움직임은 어떻게 출력하지?", "약 25분 · 슬라이드 15–25",
  `[0:25] Part 2 시작. "이제 VLA를 분해해 봅시다. 핵심 질문은 하나입니다: 행동을 어떻게 출력할까?"`);

// 15. 해부도
{
  const s = content("VLA 해부도: VLM + 행동 출력부", "PART 2 · 구성요소");
  box(s, "이미지", 0.6, 2.0, 1.6, 0.8, { fill: { color: C.Vt }, color: C.V, bold: true });
  box(s, "언어 지시", 0.6, 3.3, 1.6, 0.8, { fill: { color: C.Lt }, color: C.L, bold: true });
  box(s, "로봇 상태", 0.6, 4.6, 1.6, 0.8, { fill: { color: C.card }, color: C.muted, bold: true });
  box(s, "Vision Encoder\nSigLIP · DINOv2", 2.7, 2.0, 2.3, 0.8, { fill: { color: C.V }, color: C.white, bold: true, fontSize: 13 });
  box(s, "Projector\n(MLP)", 5.4, 2.0, 1.5, 0.8, { fill: { color: C.card }, bold: true, fontSize: 13 });
  box(s, "Tokenizer", 2.7, 3.3, 2.3, 0.8, { fill: { color: C.L }, color: C.white, bold: true, fontSize: 13 });
  box(s, "선형층 / 토큰화", 2.7, 4.6, 2.3, 0.8, { fill: { color: C.card }, bold: true, fontSize: 13 });
  arrow(s, 2.2, 2.4, 2.7, 2.4); arrow(s, 2.2, 3.7, 2.7, 3.7); arrow(s, 2.2, 5.0, 2.7, 5.0); arrow(s, 5.0, 2.4, 5.4, 2.4);
  box(s, "LLM Backbone\nLlama · Gemma · Qwen", 7.4, 2.0, 2.3, 3.4, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 15 });
  arrow(s, 6.9, 2.4, 7.4, 2.7); arrow(s, 5.0, 3.7, 7.4, 3.7); arrow(s, 5.0, 5.0, 7.4, 4.7);
  box(s, "Action Head\n토큰 · 회귀 · 플로우", 10.2, 2.9, 2.5, 1.6, { fill: { color: C.A }, color: C.white, bold: true, fontSize: 15 });
  arrow(s, 9.7, 3.7, 10.2, 3.7, C.A, 2.5);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.55, y: 1.75, w: 7.35, h: 3.95, rectRadius: 0.1, fill: { color: C.white, transparency: 100 }, line: { color: C.L, width: 1.5, dashType: "dash" } });
  T(s, "= 사실상 VLM 그대로 (LLaVA 구조)", { x: 2.6, y: 5.8, w: 7.3, h: 0.35, fontSize: 13, color: C.L, bold: true, align: "center" });
  T(s, "VLA에서 새로 생긴 것", { x: 10.2, y: 4.65, w: 2.5, h: 0.35, fontSize: 12, color: C.A, bold: true, align: "center" });
  T(s, "설계 선택의 대부분은 \"Action Head를 어떻게 만들까?\"에 있다", { x: 0.6, y: 6.4, w: 12.1, h: 0.45, fontSize: 17, bold: true, align: "center" });
  s.addNotes(`[0:26] "VLA를 열어보면 이렇게 생겼습니다. 점선 안쪽은 사실 VLM, 예를 들어 LLaVA와 똑같습니다. 이미지는 비전 인코더로 토큰이 되고, 프로젝터로 LLM이 이해할 수 있는 벡터로 바뀌고, 언어 토큰과 함께 LLM에 들어갑니다."
"VLA에서 새로 생긴 건 두 가지: 로봇 상태 입력, 그리고 주황색 Action Head."
"그래서 VLM을 아는 분은 VLA를 이미 절반 아는 겁니다. 나머지 절반이 'Action Head를 어떻게 만드냐'입니다."`);
}

// 16. Vision encoder
{
  const s = content("Vision Encoder: 이미지를 토큰으로", "PART 2 · 구성요소");
  // image grid
  for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8 + c * 0.55, y: 2.0 + r * 0.55, w: 0.5, h: 0.5, fill: { color: (r + c) % 2 ? C.Vt : "C5ECE9" }, line: { color: C.white } });
  }
  T(s, "이미지를 패치로 자름\n(224×224, 14×14 패치 → 256개)", { x: 0.6, y: 4.3, w: 2.9, h: 0.8, fontSize: 12, color: C.muted, align: "center" });
  arrow(s, 3.2, 3.05, 4.1, 3.05, C.V, 2.5);
  box(s, "ViT", 4.2, 2.5, 1.2, 1.1, { fill: { color: C.V }, color: C.white, bold: true, fontSize: 20 });
  arrow(s, 5.5, 3.05, 6.3, 3.05, C.V, 2.5);
  for (let i = 0; i < 8; i++) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.4 + i * 0.4, y: 2.75, w: 0.3, h: 0.6, rectRadius: 0.05, fill: { color: C.V }, line: { color: C.V } });
  T(s, "이미지 토큰 = 단어처럼 LLM에 입력", { x: 6.3, y: 3.5, w: 3.5, h: 0.4, fontSize: 12, color: C.muted });
  card(s, 0.6, 5.0, 5.9, 1.9, "SigLIP (Google)", "이미지-텍스트 대조학습 → \"이게 뭔지\" 의미 파악에 강함", C.V, { fill: C.Vt });
  card(s, 6.8, 5.0, 5.9, 1.9, "DINOv2 (Meta)", "자기지도학습 → \"어디에 어떤 모양\" 공간·기하 정보에 강함.\nOpenVLA는 둘을 함께 사용", C.V, { fill: C.Vt });
  T(s, "카메라 여러 대(머리 + 손목)면 각각 인코딩해 토큰을 이어 붙임", { x: 10.0, y: 2.1, w: 2.8, h: 1.6, fontSize: 13, color: C.text });
  s.addNotes(`[0:28] "이미지는 ViT로 들어갑니다. 이미지를 작은 정사각형 패치로 자르고, 각 패치를 벡터 하나로 만듭니다. 224×224 이미지면 256개 토큰. LLM 입장에서는 단어 256개가 들어온 것과 같습니다."
"대표 인코더가 두 개 있는데, SigLIP은 '이게 컵이다'라는 의미에 강하고, DINOv2는 '컵이 여기 이런 모양으로 있다'는 공간 정보에 강합니다. 로봇은 둘 다 필요하죠. 그래서 OpenVLA는 둘을 합쳐 씁니다."`);
}

// 17. VLM backbone
{
  const s = content("왜 사전학습된 VLM을 쓰나: 인터넷 지식의 전이", "PART 2 · 구성요소");
  box(s, "\"멸종된 동물을 집어\"", 0.6, 2.0, 3.6, 0.9, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 17 });
  arrow(s, 4.3, 2.45, 5.1, 2.45, C.faint, 2.5);
  box(s, "VLM의 인터넷 지식:\n공룡 = 멸종된 동물", 5.2, 2.0, 3.3, 0.9, { fill: { color: C.card }, bold: true, fontSize: 14 });
  arrow(s, 8.6, 2.45, 9.4, 2.45, C.A, 2.5);
  box(s, "공룡 장난감을 집는다", 9.5, 2.0, 3.2, 0.9, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 16 });
  T(s, "로봇 데이터에 '멸종'이라는 단어는 한 번도 없었다 (RT-2 실험)", { x: 0.6, y: 3.05, w: 12, h: 0.4, fontSize: 13, color: C.muted });
  card(s, 0.6, 3.7, 5.9, 3.0, "주의: 파국적 망각", [
    "로봇 데이터로만 학습하면 VLM이 원래 알던 언어·상식을 잊어버림",
    "해결 ① Co-training: 웹 데이터(VQA·캡션)를 섞어서 함께 학습",
    "해결 ② Knowledge Insulation: 행동 헤드의 그래디언트가 VLM을 망가뜨리지 않게 차단",
  ], C.bad);
  card(s, 6.8, 3.7, 5.9, 3.0, "대표 백본", [
    "RT-2: PaLI-X (55B), PaLM-E (12B)",
    "OpenVLA: Llama 2 7B (Prismatic)",
    "π0 / π0.5: PaliGemma 3B",
    "GR00T: Eagle-2 → Cosmos-Reason",
    "Gemini Robotics: Gemini",
  ], C.L);
  s.addNotes(`[0:30] "왜 굳이 VLM을 쓸까요? 인터넷 지식 때문입니다."
★ 퀴즈: "로봇한테 '멸종된 동물을 집어'라고 했더니 무엇을 집었을까요? 테이블엔 사과, 컵, 공룡 장난감." → 공룡.
"로봇 데이터에는 '멸종'이라는 말이 한 번도 없었습니다. VLM이 인터넷에서 배운 지식이 행동으로 이어진 겁니다."
"그런데 주의할 점: 로봇 데이터로만 학습시키면 VLM이 원래 알던 걸 잊어버립니다. 파국적 망각. 그래서 웹 데이터를 섞어 학습하는 co-training이 중요합니다. RT-2와 π0.5에서 다시 나옵니다."`);
}

// 18. 4가지 방식
{
  const s = content("★ 핵심 질문: 행동을 어떻게 출력할까?", "PART 2 · 구성요소");
  const ways = [
    ["이산 토큰", "행동 값을 256구간으로 잘라 '단어'처럼 분류", "RT-1 · RT-2 · OpenVLA"],
    ["연속 회귀", "MLP로 실수 벡터를 바로 출력 (L1/MSE)", "OpenVLA-OFT"],
    ["디퓨전 / 플로우 매칭", "노이즈에서 시작해 행동을 점점 '그려냄'", "Octo · π0 · GR00T"],
    ["압축 토큰 (FAST)", "행동 시퀀스를 JPEG처럼 압축해 토큰화", "π0-FAST · π0.5 사전학습"],
  ];
  ways.forEach((w, i) => {
    const x = 0.6 + (i % 2) * 6.2, y = 1.85 + Math.floor(i / 2) * 2.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.9, h: 2.25, rectRadius: 0.08, fill: { color: C.At }, line: { color: C.At } });
    numCircle(s, i + 1, x + 0.3, y + 0.3, 0.6, C.A, 18);
    T(s, w[0], { x: x + 1.1, y: y + 0.3, w: 4.6, h: 0.6, fontSize: 20, bold: true, valign: "middle" });
    T(s, w[1], { x: x + 1.1, y: y + 1.0, w: 4.6, h: 0.6, fontSize: 14.5 });
    T(s, w[2], { x: x + 1.1, y: y + 1.65, w: 4.6, h: 0.4, fontSize: 12.5, color: C.A, bold: true });
  });
  T(s, "VLA 논문을 읽을 때 가장 먼저 확인할 것", { x: 0.6, y: 6.85, w: 12, h: 0.35, fontSize: 13, color: C.muted, bold: true });
  s.addNotes(`[0:32] ★ 이 강의의 이론적 핵심 슬라이드. 천천히.
"LLM은 다음 단어를 고릅니다. 그런데 로봇 행동은 0.0132 같은 연속된 실수입니다. 방법은 크게 4가지입니다."
1) 이산 토큰: 숫자를 256칸 눈금으로 잘라서 '단어'처럼 다룸. RT-2, OpenVLA.
2) 연속 회귀: 그냥 숫자를 바로 출력. 빠르지만 다봉성 문제.
3) 디퓨전/플로우: 이미지 생성 AI처럼 노이즈에서 동작을 그려냄. π0, GR00T.
4) FAST: 동작 시퀀스를 JPEG처럼 압축해서 토큰으로.
"VLA 논문을 읽으면 가장 먼저 '행동을 어떻게 출력했지?'를 보세요. 이 선택이 속도, 정밀도, 학습 효율을 다 결정합니다."`);
}

// 19. 이산 토큰
{
  const s = content("① 이산 토큰화: 연속 값을 256칸 눈금으로", "PART 2 · 구성요소");
  // number line
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.7, w: 11.5, h: 0, line: { color: C.text, width: 2 } });
  for (let i = 0; i <= 16; i++) s.addShape(pres.shapes.LINE, { x: 0.9 + i * (11.5 / 16), y: 2.55, w: 0, h: 0.3, line: { color: C.faint, width: 1 } });
  T(s, "-0.05 m", { x: 0.5, y: 2.95, w: 1.0, h: 0.3, fontSize: 12, color: C.muted, align: "center" });
  T(s, "+0.05 m", { x: 11.9, y: 2.95, w: 1.0, h: 0.3, fontSize: 12, color: C.muted, align: "center" });
  const px = 0.9 + 11.5 * (0.063 / 0.1);
  dot(s, px - 0.12, 2.58, 0.24, C.A);
  T(s, "Δx = +0.013 m", { x: px - 1.0, y: 1.9, w: 2.0, h: 0.35, fontSize: 14, bold: true, color: C.A, align: "center" });
  T(s, "→ bin 161 → 토큰 \"161\"", { x: px - 1.3, y: 3.0, w: 2.6, h: 0.35, fontSize: 14, bold: true, color: C.A, align: "center" });
  code(s, "행동 7차원 = 토큰 7개:  \"161  98  130  127  140  126  255\"", 0.6, 3.75, 12.1, 0.7, 16);
  card(s, 0.6, 4.75, 5.9, 2.0, "장점", ["VLM 구조를 전혀 바꿀 필요 없음 (그냥 '단어' 예측)", "분류 문제라 여러 정답(다봉성) 표현 가능"], C.good);
  card(s, 6.8, 4.75, 5.9, 2.0, "단점", ["해상도 한계 (256단계)", "토큰을 하나씩 순차 생성 → 느림 (고주파 제어 불리)"], C.bad);
  s.addNotes(`[0:34] "Δx가 -5cm에서 +5cm 사이라고 합시다. 이 구간을 256칸으로 나누면, +1.3cm는 161번 칸입니다. 그럼 로봇 행동은 '161'이라는 단어가 됩니다."
"7차원 행동이면 단어 7개. LLM은 문장 쓰듯 이 7개 단어를 씁니다."
세부(질문 대비): OpenVLA는 범위를 최소~최대가 아니라 1~99 백분위수로 잡아 이상치 영향을 줄이고, Llama 토크나이저에서 가장 안 쓰는 256개 토큰을 행동 토큰으로 덮어씁니다.
퀴즈(선택): "한 칸 크기는?" → 10cm/256 ≈ 0.39mm.
"장점은 VLM을 그대로 쓸 수 있다는 것. 단점은 느리고, 256단계라 거칠다는 것."`);
}

// 20. 디퓨전/플로우
{
  const s = content("③ 디퓨전 · 플로우 매칭: 노이즈에서 행동을 그려내기", "PART 2 · 구성요소");
  // noise → action
  const pts = [[1.0, 2.2], [1.6, 3.3], [1.2, 4.1], [2.1, 2.6], [1.8, 3.9], [2.3, 3.4]];
  pts.forEach((p) => dot(s, p[0], p[1], 0.18, C.faint));
  T(s, "노이즈", { x: 0.8, y: 4.5, w: 1.8, h: 0.3, fontSize: 13, bold: true, color: C.muted, align: "center" });
  for (let k = 0; k < 4; k++) {
    const x = 3.0 + k * 1.3;
    arrow(s, x, 3.3, x + 1.0, 3.3, C.A, 2);
    T(s, `step ${k + 1}`, { x: x - 0.1, y: 2.9, w: 1.2, h: 0.3, fontSize: 10, color: C.muted, align: "center" });
  }
  s.addShape(pres.shapes.LINE, { x: 8.3, y: 3.6, w: 1.0, h: 0.5, line: { color: C.A, width: 4 } });
  s.addShape(pres.shapes.LINE, { x: 9.3, y: 3.2, w: 1.0, h: 0.9, line: { color: C.A, width: 4 }, flipV: true });
  s.addShape(pres.shapes.LINE, { x: 10.3, y: 3.2, w: 0.9, h: 0.3, line: { color: C.A, width: 4 } });
  T(s, "행동 청크 (궤적)", { x: 8.3, y: 4.5, w: 3.0, h: 0.3, fontSize: 13, bold: true, color: C.A, align: "center" });
  card(s, 0.6, 5.0, 3.9, 1.9, "디퓨전", "노이즈를 조금씩 걷어내며 생성 (Diffusion Policy, Octo)", C.A);
  card(s, 4.7, 5.0, 3.9, 1.9, "플로우 매칭", "노이즈 → 정답으로 가는 '속도'를 학습. 직선 경로라 ~10스텝이면 충분 (π0, GR00T)", C.A);
  card(s, 8.8, 5.0, 3.9, 1.9, "왜 좋은가", "왼쪽/오른쪽 중 하나를 샘플 → 다봉성 해결, 연속 · 정밀", C.good);
  T(s, "학습: A_τ = τ·A + (1−τ)·ε   →   v_θ(A_τ) ≈ A − ε", { x: 0.6, y: 1.7, w: 12, h: 0.4, fontFace: MONO, fontSize: 14, color: C.muted });
  s.addNotes(`[0:37] "이미지 생성 AI, Stable Diffusion을 떠올려보세요. 노이즈에서 시작해 점점 선명한 그림을 그립니다. 로봇도 똑같이 노이즈에서 시작해 동작 궤적을 그려낼 수 있습니다."
플로우 매칭 직관 (π0에서 사용):
1) 정답 행동 A와 랜덤 노이즈 ε를 섞는다 (τ 비율).
2) 모델은 "지금 위치에서 정답 쪽으로 가는 방향(속도)"을 맞춘다. 정답 방향은 A−ε.
3) 추론 때는 노이즈에서 시작해서 예측한 방향으로 10걸음 걷는다.
비유: "안개 속에서 나침반 보며 10걸음 걸어 목적지에 도착."
"왜 좋냐? 장애물 예제에서, 이 모델은 왼쪽 또는 오른쪽 중 하나를 샘플합니다. 평균 내서 충돌하지 않죠."
(질문 대비) π0 논문·openpi 코드는 시간 방향을 반대로 잡아 u=ε−A로 표기 — 부호 관례 차이일 뿐.`);
}

// 21. 비교표
{
  const s = content("행동 출력 방식 한눈에 비교", "PART 2 · 구성요소");
  tbl(s, [
    ["방식", "대표 모델", "다봉성", "정밀도", "추론 속도", "VLM 변경"],
    ["① 이산 토큰 (binning)", "RT-2, OpenVLA", "○", "△ (256단계)", "느림 (순차)", "없음"],
    ["② 연속 회귀", "OpenVLA-OFT", "✕", "○", "매우 빠름 (병렬)", "헤드 추가"],
    ["③ 디퓨전 / 플로우", "Octo, π0, GR00T", "◎", "◎", "중간 (~10스텝 반복)", "행동 전문가 추가"],
    ["④ 압축 토큰 (FAST)", "π0-FAST", "○", "○", "느림 (순차)", "토크나이저만"],
  ], { x: 0.6, y: 1.9, w: 12.1, colW: [2.8, 2.4, 1.3, 1.8, 2.3, 1.5], fontSize: 14, rowH: 0.7 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.8, w: 12.1, h: 0.95, rectRadius: 0.08, fill: { color: C.At }, line: { color: C.At } });
  T(s, "정답은 하나가 아니다 → π0.5처럼 \"학습은 토큰(FAST)으로, 실행은 플로우로\" 섞어 쓰는 하이브리드가 등장", { x: 0.85, y: 5.8, w: 11.6, h: 0.95, fontSize: 15, bold: true, valign: "middle" });
  s.addNotes(`[0:39] 표를 가로로 한 줄씩. "토큰은 쉽지만 느리고, 회귀는 빠르지만 다봉성에 약하고, 플로우는 정교하지만 반복 계산이 필요합니다."
"재미있는 건 정답이 하나가 아니라는 것. π0.5는 학습할 때는 토큰으로 빠르게, 로봇을 움직일 때는 플로우로 정교하게 — 둘을 섞습니다. 또 OFT는 단순 회귀가 벤치마크에서 제일 좋기도 합니다."`);
}

// 22. 액션 청킹
{
  const s = content("액션 청킹: 한 스텝씩이 아니라 '다음 1초'를 한 번에", "PART 2 · 구성요소");
  for (let i = 0; i < 20; i++) {
    const inChunk = i >= 2 && i < 12;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8 + i * 0.58, y: 2.3, w: 0.5, h: 0.6, fill: { color: inChunk ? C.A : C.card }, line: { color: C.white } });
  }
  T(s, "t", { x: 0.8 + 2 * 0.58, y: 2.95, w: 0.5, h: 0.3, fontSize: 12, align: "center", color: C.muted });
  T(s, "t + H", { x: 0.8 + 11 * 0.58 - 0.2, y: 2.95, w: 0.9, h: 0.3, fontSize: 12, align: "center", color: C.muted });
  T(s, "한 번의 추론 → H개 행동 (π0: H=50, 50Hz → 1초 분량)", { x: 0.8 + 2 * 0.58, y: 1.8, w: 8, h: 0.4, fontSize: 14, bold: true, color: C.A });
  const ben = [
    ["누적 오차 감소", "결정 횟수가 1/H로 → 오류가 쌓일 기회가 줄어듦"],
    ["시간적 일관성", "사람 시연의 '잠깐 멈춤' 같은 행동도 자연스럽게"],
    ["속도", "무거운 모델을 1초에 1–2번만 돌려도 50Hz 제어"],
  ];
  ben.forEach((b, i) => card(s, 0.6 + i * 4.1, 3.6, 3.8, 1.9, b[0], b[1], C.A));
  T(s, "실행 요령: 청크 일부만 실행 후 재예측 (receding horizon) · 겹치는 예측 평균 (temporal ensembling, ACT) · 실행 중 다음 청크 미리 계산 (비동기 / Real-Time Chunking)", { x: 0.6, y: 5.8, w: 12.1, h: 0.9, fontSize: 13.5, color: C.muted });
  s.addNotes(`[0:41] "운전할 때 매 순간 핸들 각도를 새로 결정하지 않죠. '앞으로 1초는 이렇게 꺾자' 하고 계획합니다. 액션 청킹이 그겁니다."
"ALOHA/ACT 논문(2023)에서 대중화됐고, 지금은 거의 모든 VLA가 씁니다."
장점 3가지 → 누적 오차 감소(함정 1 해결!), 시간적 일관성, 속도.
단점 질문 대비: "청크 실행 중엔 새 관측을 반영 못 한다 → 일부만 실행하고 재예측, 또는 비동기 추론으로 해결."`);
}

// 23. FAST
{
  const s = content("④ FAST: 동작을 JPEG처럼 압축해서 토큰으로", "PART 2 · 구성요소");
  T(s, "문제: 50Hz 단순 binning → 1초에 700토큰, 이웃 토큰이 거의 같아 모델이 '복사'만 배움", { x: 0.6, y: 1.75, w: 12.1, h: 0.45, fontSize: 15, bold: true, color: C.bad });
  const st = [["정규화", "각 차원을\n[-1, 1]로"], ["DCT", "주파수 성분으로 변환\n(JPEG과 같은 원리)"], ["양자화", "반올림 →\n고주파 대부분 0"], ["BPE 압축", "LLM 토크나이저처럼\n압축 → 30~60토큰"]];
  st.forEach((x, i) => {
    const bx = 0.6 + i * 3.15;
    box(s, [{ text: x[0], options: { bold: true, fontSize: 18, color: C.A, breakLine: true } }, { text: x[1], options: { fontSize: 13, color: C.text } }], bx, 2.5, 2.75, 1.7, { fill: { color: C.At } });
    if (i < 3) arrow(s, bx + 2.78, 3.35, bx + 3.12, 3.35, C.A);
  });
  stat(s, 0.6, 4.5, 3.8, "≈5배", "학습 속도 향상 (π0-FAST vs π0)\n성능은 디퓨전 기반과 비슷", C.A);
  card(s, 4.7, 4.6, 8.0, 2.1, "정리", [
    "토큰(자기회귀) 방식도 고주파 · 정교한 작업이 가능해짐",
    "단점: 토큰을 순차 생성하므로 추론은 플로우보다 느림",
    "범용 토크나이저 FAST+ 공개 → π0.5 사전학습에 사용",
  ], C.A);
  s.addNotes(`[0:43] (시간 없으면 한 줄로: "동작을 JPEG처럼 압축해서 토큰 수를 줄였다" 하고 넘어가기)
"50Hz로 binning하면 1초에 수백 개 토큰이 나오는데, 이웃 토큰이 거의 같은 값이라 모델이 '앞 토큰 복사'만 배웁니다."
"FAST는 JPEG의 원리를 씁니다. 부드러운 동작은 저주파 몇 개로 표현되니까, DCT로 주파수로 바꾸고, 반올림해서 0을 많이 만들고, BPE로 압축합니다."
"결과: 학습이 5배 빨라지고, 성능은 디퓨전 방식과 비슷. 다만 추론은 느립니다."`);
}

// 24. 데이터 & 레시피
{
  const s = content("데이터셋과 학습 레시피: LLM과 똑같은 구조", "PART 2 · 구성요소");
  const st = [["VLM 사전학습", "인터넷 이미지-텍스트\n(이미 된 것을 가져옴)", C.L], ["로봇 사전학습", "여러 로봇 · 여러 작업\n→ 범용 기초 체력", C.A], ["사후학습", "목표 로봇 · 목표 작업\n고품질 데이터 → 숙련", C.A], ["RL · 경험 학습", "직접 해보고 개선\n(π*0.6, 2025~)", C.good]];
  st.forEach((x, i) => {
    const bx = 0.6 + i * 3.15;
    box(s, [{ text: x[0], options: { bold: true, fontSize: 17, color: x[2], breakLine: true } }, { text: x[1], options: { fontSize: 13 } }], bx, 1.85, 2.75, 1.6, { fill: { color: C.card } });
    if (i < 3) arrow(s, bx + 2.78, 2.65, bx + 3.12, 2.65);
  });
  T(s, "+ Co-training: 로봇 데이터와 웹 데이터를 섞어 상식 유지", { x: 0.6, y: 3.55, w: 12, h: 0.35, fontSize: 13, color: C.L, bold: true });
  tbl(s, [
    ["데이터셋", "연도", "규모", "특징"],
    ["Open X-Embodiment", "2023", "100만+ 궤적 · 22종 로봇 · 21개 기관", "\"로봇계의 ImageNet\""],
    ["DROID", "2024", "7.6만 궤적 · 350시간 · 564개 장면", "실제 다양한 환경의 Franka"],
    ["π 데이터셋", "2024", "~1만 시간 · 7종 로봇 · 68개 작업", "비공개 (π0 학습용)"],
    ["LeRobot 허브", "2024~", "수백 개 커뮤니티 데이터셋", "저가 SO-100/101 팔"],
  ], { x: 0.6, y: 4.1, w: 12.1, colW: [2.6, 1.0, 5.0, 3.5], fontSize: 13, rowH: 0.52 });
  s.addNotes(`[0:45] "학습 레시피는 LLM과 평행합니다. VLM 사전학습은 이미 된 걸 가져오고, 여러 로봇 데이터로 로봇 사전학습, 목표 작업으로 사후학습, 최근엔 RL까지."
"로봇마다 행동 차원이 다른 건? π0는 가장 큰 차원으로 0을 채워 맞추고(zero-padding), Octo·GR00T는 로봇별 입출력층을 둡니다."
데이터셋은 OXE만 확실히 기억시키기: "22종 로봇, 21개 기관, 100만 궤적. 로봇계의 ImageNet."`);
}

// 25. 단일 vs 이중
{
  const s = content("두 가지 아키텍처 계열: 단일 시스템 vs 이중 시스템", "PART 2 · 구성요소");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.8, w: 5.9, h: 4.9, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "단일 시스템 (Monolithic)", { x: 0.85, y: 1.95, w: 5.4, h: 0.45, fontSize: 18, bold: true });
  box(s, "관측 + 언어", 1.0, 2.7, 1.6, 0.8, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 13 });
  arrow(s, 2.65, 3.1, 3.05, 3.1);
  box(s, "하나의 VLA", 3.1, 2.7, 1.6, 0.8, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 13 });
  arrow(s, 4.75, 3.1, 5.15, 3.1, C.A);
  box(s, "행동", 5.2, 2.7, 1.0, 0.8, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 13 });
  bullets(s, ["예: RT-2, OpenVLA, π0", "장점: 단순, end-to-end 학습", "단점: 모델이 크면 느림"], 0.85, 4.0, 5.4, 2.4, 15);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: 1.8, w: 5.9, h: 4.9, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "이중 시스템 (Dual-System)", { x: 7.05, y: 1.95, w: 5.4, h: 0.45, fontSize: 18, bold: true });
  box(s, "System 2 · 느린 VLM\n\"생각\" 7–10Hz", 7.1, 2.55, 2.6, 0.9, { fill: { color: C.L }, color: C.white, bold: true, fontSize: 12.5 });
  arrow(s, 8.4, 3.5, 8.4, 3.85, C.L);
  box(s, "System 1 · 빠른 정책\n\"반사\" 120–200Hz", 7.1, 3.9, 2.6, 0.9, { fill: { color: C.A }, color: C.white, bold: true, fontSize: 12.5 });
  T(s, "천천히 생각하는 뇌 +\n반사적으로 움직이는 척수", { x: 9.9, y: 2.9, w: 2.6, h: 1.4, fontSize: 13, color: C.muted, valign: "middle" });
  bullets(s, ["예: Figure Helix, NVIDIA GR00T N1, Gemini Robotics", "장점: 큰 모델의 지능 + 빠른 제어", "단점: 두 시스템 인터페이스 설계"], 7.05, 5.0, 5.4, 1.6, 14);
  s.addNotes(`[0:47–0:50] "VLA 구조는 크게 두 계열입니다. 하나의 모델이 다 하는 단일 시스템, 그리고 느리게 생각하는 큰 모델과 빠르게 움직이는 작은 모델을 나눈 이중 시스템."
"이름은 카너먼의 『생각에 관한 생각』에서 왔습니다. System 1은 빠른 직관, System 2는 느린 숙고."
"π0는 중간 형태입니다. 한 모델 안에 큰 VLM과 작은 행동 전문가가 같이 있습니다."
→ Part 2 끝. "이제 10분 쉬겠습니다. 쉬는 동안 화면의 질문을 생각해보세요."`);
}

// 26. 휴식
{
  const s = dark();
  T(s, "☕  10분 휴식", { x: 0.8, y: 1.6, w: 11, h: 1.0, fontSize: 44, bold: true, color: C.white });
  T(s, "쉬는 동안 생각해보기", { x: 0.8, y: 3.0, w: 11, h: 0.5, fontSize: 18, color: C.A, bold: true });
  T(s, "로봇은 1초에 50번 행동해야 하는데,\n7B 모델 한 번 추론에 150ms가 걸린다면 어떻게 해야 할까?", { x: 0.8, y: 3.6, w: 11.5, h: 1.4, fontSize: 24, color: "D7DCE3" });
  vlaDots(s, 0.8, 5.6);
  s.addNotes(`[0:50–1:00] 휴식. 화면의 질문을 띄워두기.
돌아와서 답 받기 → 예시 답: 액션 청킹(한 번에 여러 스텝), 이중 시스템(빠른 작은 모델), 작은 모델/양자화, 비동기 추론.
"이 답들이 전부 오늘 볼 논문들의 설계입니다."`);
}

// =====================================================================
// PART 3
// =====================================================================
section(3, "논문으로 보는 VLA의 진화", "RT-1(2022)부터 π0.7(2026)까지 — 사람들은 이 문제를 어떻게 풀어왔나", "약 45분 · 슬라이드 28–49",
  `[1:00] Part 3 시작. 휴식 퀴즈 답 받기(1분). "이제 실제 논문들을 시간순으로 보겠습니다. RT-2, OpenVLA, π0, π0.5 네 편에 시간을 집중합니다."`);

// 28. 타임라인
{
  const s = content("VLA 계보 한 장 요약 (2022 → 2026)", "PART 3 · 논문");
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 3.9, w: 12.1, h: 0, line: { color: C.line, width: 3 } });
  const ev = [
    ["2022.12", "RT-1", "대규모 로봇 데이터\n+ Transformer", 0],
    ["2023.07", "RT-2", "행동 = 텍스트 토큰\nVLA의 탄생", 1],
    ["2023.10", "OXE / RT-X", "22종 로봇\n연합 데이터", 0],
    ["2024.06", "OpenVLA", "오픈소스 7B\nVLA", 1],
    ["2024.10", "π0", "VLM + 플로우\n50Hz", 0],
    ["2025.02–03", "Helix · GR00T\nGemini Rob.", "이중 시스템\n휴머노이드", 1],
    ["2025.04", "π0.5", "처음 보는 집\n일반화", 0],
    ["2025.09–11", "Gemini R. 1.5\nπ*0.6", "생각하는 VLA\nRL로 개선", 1],
    ["2026.04", "π0.7", "조종 가능한\n범용 모델", 0],
  ];
  ev.forEach((e, i) => {
    const x = 0.6 + i * 1.36;
    const up = e[3] === 0;
    const col = [C.V, C.L, C.A][i % 3];
    dot(s, x + 0.5, 3.78, 0.24, col);
    T(s, e[0], { x: x - 0.05, y: up ? 3.3 : 4.2, w: 1.3, h: 0.3, fontSize: 10.5, color: C.muted, align: "center", bold: true });
    T(s, e[1], { x: x - 0.1, y: up ? 2.35 : 4.55, w: 1.4, h: 0.75, fontSize: 12.5, bold: true, align: "center", valign: up ? "bottom" : "top" });
    T(s, e[2], { x: x - 0.1, y: up ? 1.6 : 5.45, w: 1.4, h: 0.75, fontSize: 10.5, color: C.muted, align: "center", valign: up ? "bottom" : "top" });
  });
  T(s, "큰 흐름: 전용 모델 → VLM 기반 → 정교한 연속 제어 → 일반화 → 추론 · 경험 학습", { x: 0.6, y: 6.5, w: 12.1, h: 0.4, fontSize: 15, bold: true, align: "center" });
  s.addNotes(`[1:01] 전체 지도를 먼저 보여주기. "앞으로 45분은 이 타임라인을 따라갑니다."
큰 흐름 요약: "처음엔 로봇 전용 모델(RT-1) → VLM에 행동을 가르침(RT-2, OpenVLA) → 정교한 연속 제어(π0) → 처음 보는 곳에서 일반화(π0.5) → 이중 시스템, 생각하는 VLA → 경험으로 스스로 개선(π*0.6) → 말로 코칭 가능한 모델(π0.7)."`);
}

// 29. 배경
{
  const s = content("배경: VLA 이전의 두 가지 핵심 기법 (2023)", "PART 3 · 논문");
  card(s, 0.6, 1.85, 5.9, 3.4, "ACT / ALOHA  (Zhao et al.)", [
    "약 2만 달러짜리 저가 양팔 원격조종 시스템",
    "Action Chunking with Transformers: 한 번에 100스텝 예측",
    "~50개 시연(10분)으로 배터리 끼우기 등 정밀 작업 80–90%",
    { t: "→ '액션 청킹'을 대중화", b: true },
  ], C.V);
  card(s, 6.8, 1.85, 5.9, 3.4, "Diffusion Policy  (Chi et al.)", [
    "이미지 생성용 디퓨전으로 행동 시퀀스 생성",
    "다봉성을 자연스럽게 표현",
    "여러 벤치마크에서 평균 ~47% 향상",
    { t: "→ '행동 = 생성 모델 문제' 관점", b: true },
  ], C.A);
  T(s, "한계: 작업마다 모델을 따로 학습하는 전문가(specialist) 정책, 언어 이해 없음 → VLA는 이를 하나의 범용 모델로", { x: 0.6, y: 5.6, w: 12.1, h: 0.8, fontSize: 15, bold: true, color: C.muted });
  s.addNotes(`[1:03] 짧게(2분). "VLA 이전에도 모방학습은 잘 되고 있었습니다. 두 논문만 기억하세요."
- ACT/ALOHA: 싸게 양팔 데이터를 모으는 방법 + 액션 청킹.
- Diffusion Policy: 디퓨전으로 행동 생성.
"둘 다 오늘 뒤에서 계속 등장합니다. 다만 작업마다 모델을 따로 학습해야 했고, 언어를 이해하지 못했습니다. VLA는 이걸 하나의 범용 모델로 만들려는 시도입니다."`);
}

// 30. RT-1
{
  const s = content("RT-1: 로봇 13대, 17개월, 13만 에피소드", "PART 3 · 논문");
  paperTag(s, "Google · 2022.12");
  box(s, "이미지 6장", 0.6, 2.0, 1.7, 0.8, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 13 });
  box(s, "언어 지시", 0.6, 3.1, 1.7, 0.8, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 13 });
  arrow(s, 2.35, 2.4, 2.75, 2.4);
  arrow(s, 2.35, 3.5, 2.75, 2.7, C.L);
  box(s, "EfficientNet\n+ FiLM", 2.8, 2.0, 1.7, 0.9, { fill: { color: C.V }, color: C.white, bold: true, fontSize: 13 });
  arrow(s, 4.55, 2.45, 4.95, 2.45);
  box(s, "TokenLearner\n81 → 8 토큰", 5.0, 2.0, 1.9, 0.9, { fill: { color: C.card }, bold: true, fontSize: 13 });
  arrow(s, 6.95, 2.45, 7.35, 2.45);
  box(s, "Transformer\n(8층)", 7.4, 2.0, 1.7, 0.9, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 13 });
  arrow(s, 9.15, 2.45, 9.55, 2.45, C.A);
  box(s, "행동 11차원\n× 256 bin", 9.6, 2.0, 1.9, 0.9, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 13 });
  T(s, "35M 파라미터 · 3Hz · VLM 아님", { x: 2.8, y: 3.1, w: 8.7, h: 0.4, fontSize: 13, color: C.muted });
  stat(s, 0.6, 4.2, 2.9, "97%", "학습한 작업 성공률", C.A);
  stat(s, 3.7, 4.2, 2.9, "76%", "처음 보는 지시", C.A);
  stat(s, 6.8, 4.2, 2.9, "700+", "수행 가능한 지시 수", C.V);
  card(s, 9.9, 4.2, 2.8, 2.5, "한계", "인터넷 지식 없음 → 로봇 데이터에 없는 개념은 모름", C.bad, { hs: 16 });
  T(s, "기억할 것: '행동을 256구간으로 나눈 분류' 아이디어의 원조", { x: 0.6, y: 6.35, w: 9, h: 0.4, fontSize: 14, bold: true });
  s.addNotes(`[1:05] "VLA의 할아버지, RT-1입니다. 구글이 로봇 13대로 17개월 동안 13만 개 시연을 모았습니다."
구조: 이미지를 EfficientNet으로 인코딩하는데, 언어로 FiLM 조건화 → '무엇을 볼지'를 언어가 조절. TokenLearner로 토큰 수를 줄여 3Hz 실시간.
"출력은 행동 11차원을 각각 256칸으로 나눈 분류. 이 아이디어가 RT-2, OpenVLA로 이어집니다."
결과: 본 작업 97%, 처음 보는 지시 76%. 한계: 인터넷 지식이 없어서, 로봇 데이터에 없는 개념은 모름 → RT-2의 동기.`);
}

// 31. RT-2 (1)
{
  const s = content("★ RT-2: \"행동도 그냥 언어다\"", "PART 3 · 논문");
  paperTag(s, "Google DeepMind · 2023.07");
  T(s, "질문", { x: 0.6, y: 1.85, w: 2, h: 0.3, fontSize: 12, bold: true, color: C.L });
  code(s, "\"What action should the robot take to pick up the apple?\"", 0.6, 2.2, 12.1, 0.65, 15);
  T(s, "VLM의 정답 (텍스트!)", { x: 0.6, y: 3.0, w: 4, h: 0.3, fontSize: 12, bold: true, color: C.A });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 3.35, w: 12.1, h: 1.0, rectRadius: 0.06, fill: { color: C.ink2 }, line: { color: C.ink2 } });
  const toks = [["1", "종료"], ["128", "Δx"], ["91", "Δy"], ["241", "Δz"], ["5", "Δroll"], ["101", "Δpitch"], ["127", "Δyaw"], ["217", "그리퍼"]];
  toks.forEach((t, i) => {
    T(s, t[0], { x: 0.8 + i * 1.47, y: 3.4, w: 1.3, h: 0.55, fontFace: MONO, fontSize: 20, bold: true, color: "FFB38A", align: "center", valign: "middle" });
    T(s, t[1], { x: 0.8 + i * 1.47, y: 3.93, w: 1.3, h: 0.35, fontSize: 11.5, color: "B8C0CC", align: "center" });
  });
  card(s, 0.6, 4.55, 5.9, 2.25, "Co-fine-tuning (핵심 비법)", "로봇 데이터만이 아니라 원래의 인터넷 VQA·캡션 데이터와 섞어서 학습 → 상식 유지 + 행동 학습", C.L);
  card(s, 6.8, 4.55, 5.9, 2.25, "모델", "PaLI-X (5B / 55B), PaLM-E (12B)\n55B는 클라우드 TPU에서 1–3Hz", C.muted);
  s.addNotes(`[1:07] ★ 중요 논문. "VLA라는 이름을 만든 논문입니다. 아이디어는 딱 하나: 행동도 그냥 언어다."
"VLM한테 '사과를 집으려면 로봇이 뭘 해야 해?'라고 물으면, 답으로 '1 128 91 241 ...'이라는 숫자 문자열을 말하도록 가르칩니다. 각 숫자는 RT-1처럼 256칸 중 몇 번 칸인지."
"VLM 입장에서는 캡션 달기, 질문 답하기와 똑같은 텍스트 생성 문제입니다. 구조를 하나도 안 바꿨습니다."
Co-fine-tuning 강조: "로봇 데이터만 넣으면 VLM이 원래 상식을 잊어버립니다. 그래서 인터넷 데이터와 섞어서 학습합니다. 실험에서 섞은 쪽이 확실히 좋았습니다."
(질문 대비) PaLI-X는 1000 이하 숫자 토큰이 이미 있어 그대로 쓰고, PaLM-E는 가장 안 쓰는 256개 토큰을 덮어씀. 추론 시 행동 토큰만 나오도록 제한.`);
}

// 32. RT-2 (2)
{
  const s = content("★ RT-2: 가르치지 않은 능력 (Emergent Capabilities)", "PART 3 · 논문");
  paperTag(s, "Google DeepMind · 2023.07");
  const ex = [["\"멸종된 동물을 집어\"", "→ 공룡 장난감"], ["\"즉석 망치로 쓸 물건을 집어\"", "→ 돌"], ["\"콜라를 테일러 스위프트에게\"", "→ 사진 위로"], ["\"바나나를 2+1의 합으로\"", "→ 숫자 3 위로"]];
  ex.forEach((e, i) => {
    const y = 1.85 + i * 1.12;
    box(s, e[0], 0.6, y, 4.3, 0.9, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 14.5 });
    T(s, e[1], { x: 5.05, y, w: 2.0, h: 0.9, fontSize: 15, bold: true, color: C.A, valign: "middle" });
  });
  s.addChart(pres.charts.BAR, [{ name: "처음 보는 물체·배경·환경 평균 성공률 (%)", labels: ["RT-1", "RT-2"], values: [32, 62] }], {
    x: 7.4, y: 1.7, w: 5.3, h: 3.6, barDir: "col", chartColors: [C.A], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14,
    dataLabelColor: C.text, catAxisLabelFontSize: 14, catAxisLabelColor: C.text, catAxisLabelFontFace: F, valAxisHidden: true, valAxisMaxVal: 80,
    valGridLine: { style: "none" }, showLegend: false, showTitle: true, title: "일반화 성공률 (%, 대략값)", titleFontSize: 13, titleColor: C.muted, titleFontFace: F, barGapWidthPct: 80,
  });
  card(s, 7.4, 5.5, 5.3, 1.35, "한계", "새로운 '동작'은 못 배움 — 인터넷 지식은 무엇을·어디로만 도와줌. 크고 느리고 비공개.", C.bad, { hs: 15, bs: 13 });
  T(s, "변형: 행동 전에 \"Plan: …\"을 텍스트로 먼저 생성 (Chain-of-thought)", { x: 0.6, y: 6.45, w: 6.6, h: 0.4, fontSize: 13, color: C.muted });
  s.addNotes(`[1:10] ★ 퀴즈로 진행: 왼쪽 예시를 하나씩 가리며 "로봇이 뭘 했을까요?"
- 멸종된 동물 → 공룡 장난감
- 즉석 망치 → 돌
- 테일러 스위프트 → 사진 인식
- 2+1 → 숫자 3
"이 중 어느 것도 로봇 데이터로 가르치지 않았습니다. 인터넷 지식이 로봇 행동으로 전이된 겁니다."
결과: 처음 보는 상황에서 RT-1 약 32% → RT-2 약 62%, 약 2배. (약 6,000회 실제 로봇 평가)
한계 꼭 짚기: "새로운 동작 자체는 못 배웁니다. 로봇 데이터에 없는 '뒤집기' 같은 동작은 인터넷 지식으로 생기지 않죠. 그리고 55B라 크고 느리고, 공개되지 않았습니다." → 다음 논문들의 동기.`);
}

// 33. OXE
{
  const s = content("Open X-Embodiment & RT-X: 로봇계의 ImageNet", "PART 3 · 논문");
  paperTag(s, "21개 기관 연합 · 2023.10");
  stat(s, 0.6, 1.75, 2.9, "22종", "로봇 형태 (embodiment)", C.V);
  stat(s, 3.7, 1.75, 2.9, "21개", "참여 기관 (34개 연구실)", C.V);
  stat(s, 6.8, 1.75, 2.9, "100만+", "궤적 (60개 데이터셋)", C.A);
  stat(s, 9.9, 1.75, 2.8, "527", "기술 (skills)", C.A);
  T(s, "질문: 다른 로봇의 데이터도 내 로봇에 도움이 될까? (cross-embodiment transfer)", { x: 0.6, y: 3.75, w: 12.1, h: 0.45, fontSize: 16, bold: true });
  card(s, 0.6, 4.4, 5.9, 1.8, "RT-1-X", "데이터가 적은 연구실에서, 각자의 원래 방법 대비 평균 +50%", C.A);
  card(s, 6.8, 4.4, 5.9, 1.8, "RT-2-X (55B)", "다른 로봇 데이터 덕분에 emergent skill 평가에서 RT-2 대비 약 3배", C.A);
  s.addNotes(`[1:13] "로봇 데이터 부족에 대한 커뮤니티의 답입니다. 전 세계 21개 기관이 데이터를 모아 하나의 포맷(RLDS)으로 합쳤습니다. 22종 로봇, 100만 궤적."
"질문은 '다른 로봇 데이터가 내 로봇에 도움이 되냐'. 답은 '된다'. RT-1-X는 평균 50% 향상, RT-2-X는 emergent skill이 3배."
"이후 Octo, OpenVLA, π0 모두 이 데이터를 사전학습에 씁니다."`);
}

// 34. Octo
{
  const s = content("Octo: 오픈소스 범용 정책 (디퓨전 헤드)", "PART 3 · 논문");
  paperTag(s, "UC Berkeley 외 · 2024.05");
  const toks = [["언어 (T5)", C.L, C.Lt], ["목표 이미지", C.V, C.Vt], ["관측 이미지", C.V, C.Vt], ["readout", C.A, C.At]];
  toks.forEach((t, i) => box(s, t[0], 0.6 + i * 2.2, 1.95, 2.0, 0.75, { fill: { color: t[2] }, color: t[1], bold: true, fontSize: 13 }));
  box(s, "Transformer (블록 단위 어텐션)", 0.6, 3.0, 8.6, 0.8, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 15 });
  arrow(s, 9.25, 3.4, 9.75, 3.4, C.A, 2.5);
  box(s, "디퓨전 헤드\n→ 행동 청크", 9.8, 2.95, 2.9, 0.9, { fill: { color: C.A }, color: C.white, bold: true, fontSize: 14 });
  bullets(s, [
    "OXE 80만 궤적으로 학습 · 27M / 93M 파라미터 (작음, VLM 아님)",
    "readout 토큰: BERT의 [CLS]처럼 정보를 모아 행동 헤드로",
    "새 카메라 · 새 행동 공간 → 토큰/헤드만 추가 → 파인튜닝 친화적",
    "한계: 언어 이해 · 상식이 약함 → OpenVLA로",
  ], 0.6, 4.2, 9.0, 2.6, 15);
  s.addNotes(`[1:15] (시간 없으면 스킵 가능) "Octo는 완전 오픈소스 범용 정책의 시작입니다. OXE 80만 궤적으로 학습했고, 크기는 1억 파라미터 이하로 작습니다."
"특징은 유연함: 언어 대신 목표 이미지로 작업을 줄 수도 있고, 새 카메라나 새 로봇이 생기면 토큰과 헤드만 추가해서 파인튜닝합니다. 행동은 디퓨전 헤드로 출력."
"다만 VLM 기반이 아니라 언어 이해가 약합니다. 그래서 다음, OpenVLA."`);
}

// 35. OpenVLA (1)
{
  const s = content("★ OpenVLA: RT-2 레시피를 오픈소스 7B로", "PART 3 · 논문");
  paperTag(s, "Stanford · UC Berkeley · TRI 외 · 2024.06");
  box(s, "이미지\n224×224", 0.6, 2.3, 1.6, 1.3, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 13 });
  box(s, "SigLIP\n(의미)", 2.7, 1.9, 1.7, 0.85, { fill: { color: C.V }, color: C.white, bold: true, fontSize: 13 });
  box(s, "DINOv2\n(공간)", 2.7, 3.1, 1.7, 0.85, { fill: { color: C.V }, color: C.white, bold: true, fontSize: 13 });
  arrow(s, 2.25, 2.95, 2.7, 2.3); arrow(s, 2.25, 2.95, 2.7, 3.5);
  box(s, "concat\n+ MLP", 4.9, 2.4, 1.4, 1.1, { fill: { color: C.card }, bold: true, fontSize: 13 });
  arrow(s, 4.45, 2.3, 4.9, 2.8); arrow(s, 4.45, 3.5, 4.9, 3.1);
  box(s, "언어 지시", 4.9, 4.0, 1.4, 0.8, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 13 });
  box(s, "Llama 2\n7B", 6.9, 2.4, 1.9, 2.4, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 18 });
  arrow(s, 6.35, 2.95, 6.9, 2.95); arrow(s, 6.35, 4.4, 6.9, 4.4, C.L);
  arrow(s, 8.85, 3.6, 9.35, 3.6, C.A, 2.5);
  box(s, "행동 토큰 7개\n(256 bin)", 9.4, 3.0, 2.3, 1.2, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 14 });
  T(s, "기반 VLM: Prismatic-7B", { x: 6.9, y: 4.9, w: 3, h: 0.3, fontSize: 12, color: C.muted });
  bullets(s, [
    "데이터: OXE에서 선별한 97만 궤적 · A100 64장으로 14일",
    "bin 범위 = 1~99 백분위수 (이상치 무시) · 안 쓰는 256개 토큰을 행동 토큰으로",
    "발견: VLM과 달리 비전 인코더도 같이 학습해야 성능↑ (로봇엔 세밀한 공간 정보가 중요)",
  ], 0.6, 5.35, 12.1, 1.6, 14.5);
  s.addNotes(`[1:17] ★ 중요 논문. "RT-2는 공개되지 않았고 너무 컸습니다. OpenVLA는 이 레시피를 연구자 누구나 쓸 수 있게 7B 오픈소스로 만들었습니다."
구조: "보시면 LLaVA와 거의 같습니다. 이미지 → SigLIP과 DINOv2 두 인코더 → 합쳐서 MLP → Llama 2. 출력은 RT-2처럼 행동 토큰 7개."
"SigLIP은 '이게 뭔지', DINOv2는 '어디에 있는지'. 로봇은 둘 다 필요하죠."
세부 두 가지만: 1~99 백분위수로 범위 잡기, 안 쓰는 토큰 덮어쓰기.
흥미로운 발견: "보통 VLM 학습에서는 비전 인코더를 얼리는데, 로봇에서는 같이 학습해야 잘 됐습니다. 로봇은 '컵이 3cm 왼쪽에 있다' 같은 세밀한 정보가 필요하니까요."`);
}

// 36. OpenVLA (2)
{
  const s = content("★ OpenVLA: 7배 작은데 더 잘한다, 그리고 누구나 쓸 수 있다", "PART 3 · 논문");
  paperTag(s, "Stanford · UC Berkeley · TRI 외 · 2024.06");
  stat(s, 0.6, 1.75, 3.9, "+16.5%p", "RT-2-X (55B) 대비 절대 성공률\n29개 작업 · 2종 로봇", C.A);
  stat(s, 4.7, 1.75, 3.9, "7배", "더 적은 파라미터 (7B vs 55B)", C.A);
  stat(s, 8.8, 1.75, 3.9, "+20.4%p", "새 로봇(Franka) 파인튜닝 시\nDiffusion Policy(처음부터) 대비", C.A);
  card(s, 0.6, 4.0, 5.9, 2.8, "실용성 = 연구의 민주화", [
    "LoRA: 파라미터 1.4%만 학습해도 전체 파인튜닝과 동등",
    "4-bit 양자화로 메모리 절감, 성능 유지",
    "코드 · 가중치 · 데이터 파이프라인 전부 공개",
  ], C.good);
  card(s, 6.8, 4.0, 5.9, 2.8, "한계", [
    "~6Hz: 토큰 7개 순차 생성, 청킹 없음 → 정밀 작업 불리",
    "이미지 1장 · 로봇 상태 입력 없음",
    "256 bin 해상도 한계",
  ], C.bad);
  s.addNotes(`[1:20] "결과: 55B RT-2-X보다 7배 작은데 성공률이 16.5%p 높습니다. 새 로봇에 파인튜닝해도 처음부터 학습한 Diffusion Policy보다 좋습니다. 특히 여러 물체와 언어 지시가 섞인 작업에서."
"진짜 의미는 실용성입니다. LoRA로 GPU 한 장이면 파인튜닝이 되고, 모든 게 공개됐습니다. 이후 수많은 VLA 연구의 베이스라인이 됐습니다."
"한계는 속도. 1초에 6번. 빨래 개기 같은 건 어렵습니다." → OFT와 π0로 연결.`);
}

// 37. OFT
{
  const s = content("OpenVLA-OFT: 행동 출력 방식만 바꿨더니", "PART 3 · 논문");
  paperTag(s, "Stanford (Kim, Finn, Liang) · 2025.02");
  const ch = [["병렬 디코딩", "토큰을 하나씩이 아니라\n한 번의 forward로"], ["액션 청킹", "여러 스텝을\n한 번에 예측"], ["연속 행동 + L1", "256 bin 대신\nMLP로 실수값 직접"]];
  ch.forEach((c, i) => {
    numCircle(s, i + 1, 0.6, 1.95 + i * 1.5, 0.55, C.A, 16);
    T(s, c[0], { x: 1.35, y: 1.9 + i * 1.5, w: 4.5, h: 0.45, fontSize: 18, bold: true });
    T(s, c[1], { x: 1.35, y: 2.35 + i * 1.5, w: 4.5, h: 0.8, fontSize: 13.5, color: C.muted });
  });
  s.addChart(pres.charts.BAR, [{ name: "LIBERO 평균 성공률 (%)", labels: ["OpenVLA", "OpenVLA-OFT"], values: [76.5, 97.1] }], {
    x: 6.3, y: 1.7, w: 6.4, h: 3.9, barDir: "col", chartColors: [C.A], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 15,
    dataLabelColor: C.text, catAxisLabelFontSize: 14, catAxisLabelColor: C.text, catAxisLabelFontFace: F, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 110,
    valGridLine: { style: "none" }, showLegend: false, showTitle: true, title: "LIBERO 평균 성공률 (%)", titleFontSize: 13, titleColor: C.muted, titleFontFace: F, barGapWidthPct: 80, dataLabelFormatCode: "0.0",
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.0, w: 12.1, h: 0.85, rectRadius: 0.08, fill: { color: C.At }, line: { color: C.At } });
  T(s, "행동 생성 처리량 ~26배 ↑  ·  교훈: 행동을 '토큰'으로 출력하는 게 꼭 최선은 아니다", { x: 0.85, y: 6.0, w: 11.6, h: 0.85, fontSize: 16, bold: true, valign: "middle" });
  s.addNotes(`[1:22] "같은 OpenVLA인데, 파인튜닝할 때 행동 출력 방식 세 가지만 바꿨습니다. 병렬 디코딩, 액션 청킹, 연속 값 회귀."
"시뮬레이션 벤치마크 LIBERO에서 76.5% → 97.1%, 속도는 26배."
"Part 2에서 본 비교표가 실제로 얼마나 중요한지 보여주는 예입니다. 행동을 어떻게 출력하느냐가 성능과 속도를 다 바꿉니다."`);
}

// 38. π0 (1)
{
  const s = content("★ π0: VLM + 플로우 매칭 '행동 전문가'", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2024.10");
  box(s, "이미지 2–3장", 0.6, 1.95, 1.9, 0.75, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 13 });
  box(s, "언어 지시", 0.6, 2.85, 1.9, 0.75, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 13 });
  box(s, "로봇 상태 q", 0.6, 4.05, 1.9, 0.75, { fill: { color: C.card }, color: C.muted, bold: true, fontSize: 13 });
  box(s, "노이즈 섞인\n행동 청크", 0.6, 4.95, 1.9, 0.85, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 13 });
  box(s, "PaliGemma VLM (3B)\n\"큰 두뇌\"", 3.1, 1.95, 4.0, 1.65, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 16 });
  box(s, "Action Expert (300M)\n\"행동 전문가\"", 3.1, 4.05, 4.0, 1.75, { fill: { color: C.A }, color: C.white, bold: true, fontSize: 16 });
  arrow(s, 2.55, 2.3, 3.1, 2.5); arrow(s, 2.55, 3.2, 3.1, 3.0);
  arrow(s, 2.55, 4.4, 3.1, 4.6); arrow(s, 2.55, 5.35, 3.1, 5.2);
  s.addShape(pres.shapes.LINE, { x: 5.1, y: 3.62, w: 0, h: 0.42, line: { color: C.L, width: 2.5, dashType: "dash", endArrowType: "triangle" } });
  T(s, "같은 어텐션으로 연결", { x: 5.25, y: 3.65, w: 2.2, h: 0.35, fontSize: 11, color: C.L, bold: true });
  arrow(s, 7.15, 4.9, 7.75, 4.9, C.A, 2.5);
  box(s, "속도 벡터 v\n× 10스텝 적분\n→ 50스텝 행동", 7.8, 4.05, 2.2, 1.75, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 13 });
  bullets(s, [
    "이미지·언어는 한 번만 계산 → 캐시",
    "행동 부분만 10번 반복 → 빠름",
    "행동 청크 H=50 · 최대 50Hz",
    "RTX 4090에서 ~73ms",
    "여러 로봇: 행동 벡터를 zero-padding",
  ], 10.3, 1.95, 2.5, 4.0, 13);
  s.addNotes(`[1:24] ★ 중요 논문. "RT-2, OpenVLA는 토큰 방식이라 느리고 거칠었습니다. π0는 VLM의 일반화 능력과 디퓨전 정책의 정교함을 한 모델에 합쳤습니다."
구조: "큰 두뇌 PaliGemma 3B와, 작은 행동 전문가 300M. 이미지·언어 토큰은 큰 두뇌 가중치로, 상태·행동 토큰은 행동 전문가 가중치로 처리하지만, 같은 어텐션 안에서 서로를 봅니다."
"행동 전문가는 플로우 매칭으로, 노이즈에서 10걸음 만에 50스텝짜리 행동 청크를 만듭니다. 이미지·언어 부분은 한 번만 계산해서 캐시하고 행동 부분만 10번 돌리니까 빠릅니다."
(질문 대비) 블록 단위 인과 어텐션: [이미지+언어] → [상태] → [행동], 앞 블록은 뒤를 못 봄. 시간 τ는 노이즈가 큰 쪽을 더 자주 뽑도록 Beta 분포.`);
}

// 39. π0 (2) 레시피
{
  const s = content("★ π0: LLM처럼 사전학습 → 사후학습", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2024.10");
  card(s, 0.6, 1.85, 5.9, 2.6, "사전학습 (Pre-training)", [
    "자체 데이터 ~1만 시간 · 7종 로봇 구성 · 68개 작업 + OXE",
    "다양하지만 품질은 섞인 데이터",
    { t: "→ 넓은 기초 능력 + 실수에서 회복하는 능력", b: true },
  ], C.V, { fill: C.Vt });
  arrow(s, 6.52, 3.15, 6.78, 3.15, C.A, 3);
  card(s, 6.8, 1.85, 5.9, 2.6, "사후학습 (Post-training)", [
    "목표 작업의 고품질 데이터 5~100시간",
    "모범 사례 위주",
    { t: "→ 유창하고 능숙한 동작", b: true },
  ], C.A, { fill: C.At });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "비유", { x: 0.85, y: 4.9, w: 2, h: 0.35, fontSize: 13, bold: true, color: C.muted });
  T(s, "사전학습만 하면 → 이것저것 할 줄 알지만 서툼\n사후학습만 하면 → 능숙하지만 한 번 실수하면 회복 못 함", { x: 0.85, y: 5.25, w: 11.6, h: 0.9, fontSize: 15 });
  s.addNotes(`[1:27] "π0의 두 번째 기여는 학습 레시피입니다. LLM과 똑같이 사전학습 → 사후학습."
"사전학습은 1만 시간, 7종 로봇, 68개 작업. 품질은 섞여 있어요. 실수하고 복구하는 장면도 있죠. 그래서 넓은 능력과 회복 능력을 배웁니다."
"사후학습은 목표 작업의 깔끔한 데이터. 능숙함을 배웁니다."
"둘 다 필요합니다. 사전학습만 하면 서툴고, 사후학습만 하면 한 번 실수하면 회복을 못 합니다."
로봇 7종(질문 대비): UR5e 단일/양팔, Franka, 양팔 Trossen(ALOHA), 양팔 ARX, 모바일 Trossen, 모바일 Fibocom.`);
}

// 40. π0 (3) 결과
{
  const s = content("★ π0: 처음으로 범용 모델이 해낸 정교하고 긴 작업", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2024.10");
  const tasks = [["빨래 개기", "건조기에서 꺼내 → 바구니 → 테이블 → 개기 (수십 분짜리 작업)"], ["테이블 치우기", "식기는 통에, 쓰레기는 휴지통에 분류"], ["상자 조립", "납작한 골판지를 접어 상자로"], ["계란 담기", "깨지지 않게 계란판에 하나씩"]];
  tasks.forEach((t, i) => {
    const x = 0.6 + (i % 2) * 6.2, y = 1.85 + Math.floor(i / 2) * 1.75;
    card(s, x, y, 5.9, 1.5, t[0], t[1], C.A, { hs: 18, bs: 14 });
  });
  bullets(s, [
    "사전학습 모델만으로(zero-shot)도 OpenVLA · Octo를 크게 앞섬",
    "사전학습을 거친 모델이 사후학습 효율도 훨씬 좋음",
    "2025.02 openpi로 코드 · 가중치 공개 → \"VLM + 행동 전문가\" 구조가 사실상 표준으로",
  ], 0.6, 5.5, 12.1, 1.4, 15);
  s.addNotes(`[1:29] ★ 여기서 π0 빨래 개기 영상 30초 (pi.website/blog/pi0).
"빨래를 건조기에서 꺼내서, 바구니에 담고, 테이블로 가져가서, 하나씩 갭니다. 천은 모양이 계속 바뀌니까 로봇에게 가장 어려운 작업 중 하나입니다."
"이후 GR00T, SmolVLA 등 많은 모델이 'VLM + 행동 전문가' 구조를 따릅니다."
"한계: 처음 보는 환경에서는 아직 약했습니다. → π0.5"`);
}

// 41. FAST
{
  const s = content("FAST · π0-FAST: 토큰 방식의 부활", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2025.01");
  const st = ["행동 청크", "DCT", "양자화", "BPE", "30–60 토큰"];
  st.forEach((x, i) => {
    box(s, x, 0.6 + i * 2.5, 2.0, 2.0, 0.9, { fill: { color: i === 4 ? C.A : C.At }, color: i === 4 ? C.white : C.A, bold: true, fontSize: 15 });
    if (i < 4) arrow(s, 2.62 + i * 2.5, 2.45, 3.08 + i * 2.5, 2.45, C.A);
  });
  card(s, 0.6, 3.3, 3.9, 2.4, "성능", "디퓨전 기반 π0과 비슷한 성능", C.good);
  card(s, 4.7, 3.3, 3.9, 2.4, "학습", "약 5배 빠른 학습. DROID로 학습한 모델이 새 환경에서 zero-shot 동작", C.good);
  card(s, 8.8, 3.3, 3.9, 2.4, "단점", "자기회귀 디코딩이라 추론이 플로우보다 느림", C.bad);
  T(s, "→ π0.5의 사전학습에 사용 (다음 슬라이드)", { x: 0.6, y: 6.1, w: 12, h: 0.4, fontSize: 15, bold: true, color: C.A });
  s.addNotes(`[1:31] 짧게(1분). "Part 2에서 본 FAST입니다. 동작을 JPEG처럼 압축해서 토큰으로 만들었더니, 토큰 방식도 π0 수준으로 잘 되고 학습은 5배 빨라졌습니다. 단점은 추론 속도."
"이 FAST가 π0.5에서 핵심 역할을 합니다."`);
}

// 42. π0.5 (1)
{
  const s = content("★ π0.5: 처음 보는 집에서 일하려면 어떤 데이터가 필요한가", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2025.04");
  const d = [
    ["MM", "모바일 로봇 데이터", "~100개 집 · ~400시간", C.A],
    ["ME", "다양한 환경의 고정형 팔", "환경 다양성", C.A],
    ["CE", "실험실 다른 로봇 데이터", "OXE 포함 · 다양한 기술", C.A],
    ["HL", "고수준 하위작업 라벨", "\"베개를 집어라\"", C.L],
    ["WD", "웹 데이터", "캡션 · VQA · 물체 위치", C.V],
    ["VI", "사람의 말 코칭", "단계별 언어 지시", C.L],
  ];
  d.forEach((x, i) => {
    const cx = 0.6 + (i % 3) * 4.1, cy = 1.85 + Math.floor(i / 3) * 1.85;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: 3.8, h: 1.6, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
    s.addText(x[0], { x: cx + 0.25, y: cy + 0.3, w: 0.95, h: 0.95, shape: pres.shapes.OVAL, fill: { color: x[3] }, line: { color: x[3] }, color: C.white, fontFace: F, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    T(s, x[1], { x: cx + 1.4, y: cy + 0.3, w: 2.3, h: 0.5, fontSize: 15, bold: true, valign: "middle" });
    T(s, x[2], { x: cx + 1.4, y: cy + 0.82, w: 2.3, h: 0.5, fontSize: 12.5, color: C.muted });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.75, w: 12.1, h: 1.1, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "목표 로봇 데이터(MM)는 일부일 뿐 — 나머지 이질적인 데이터가 일반화를 만든다\n웹 데이터를 빼면 처음 보는 물체 성능↓ · 학습한 집 ~100개면 테스트 집에서 직접 학습한 수준에 근접", { x: 0.85, y: 5.75, w: 11.6, h: 1.1, fontSize: 14.5, bold: true, color: C.white, valign: "middle" });
  s.addNotes(`[1:33] ★ 중요 논문. 처음 슬라이드에서 본 영상의 논문입니다.
"π0는 배운 환경에서 능숙했지만 처음 보는 집에서는 약했습니다. π0.5의 질문: 처음 보는 곳에서 일하려면 어떤 데이터가 필요한가?"
"답: 여러 종류의 데이터를 섞는 것(co-training). 목표 로봇 데이터는 약 100개 집, 400시간. 여기에 고정형 팔 데이터, 다른 로봇 데이터, 하위작업 라벨, 웹 데이터, 사람이 말로 코칭한 데이터를 섞습니다."
★ 퀴즈: "웹 데이터를 빼면 어떤 능력이 떨어질까요?" → 처음 보는 물체.
"그리고 학습한 집 수를 늘릴수록 좋아지는데, 약 100개 집이면 테스트 집에서 직접 학습한 모델과 비슷해집니다. '양보다 다양성'."`);
}

// 43. π0.5 (2)
{
  const s = content("★ π0.5: 먼저 말하고, 그다음 움직인다", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2025.04");
  box(s, "\"침실 정리해\"", 0.6, 2.0, 2.4, 0.9, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 15 });
  arrow(s, 3.05, 2.45, 3.55, 2.45);
  box(s, "π0.5\n(고수준 추론)", 3.6, 2.0, 2.0, 0.9, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 13 });
  arrow(s, 5.65, 2.45, 6.15, 2.45, C.L);
  box(s, "텍스트: \"베개를 집어\"", 6.2, 2.0, 2.6, 0.9, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 14 });
  arrow(s, 8.85, 2.45, 9.35, 2.45);
  box(s, "π0.5\n(저수준 행동)", 9.4, 2.0, 1.6, 0.9, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 13 });
  arrow(s, 11.05, 2.45, 11.45, 2.45, C.A);
  box(s, "행동\n청크", 11.5, 2.0, 1.2, 0.9, { fill: { color: C.A }, color: C.white, bold: true, fontSize: 13 });
  T(s, "같은 모델이 하위 작업을 텍스트로 예측한 뒤, 그 텍스트를 조건으로 행동 생성 (로봇판 chain-of-thought)", { x: 0.6, y: 3.05, w: 12.1, h: 0.4, fontSize: 13, color: C.muted });
  card(s, 0.6, 3.75, 5.9, 3.0, "2단계 학습", [
    "사전학습 (280k 스텝): 모든 것을 토큰으로 — 행동은 FAST 토큰 → 빠르고 효율적",
    "사후학습 (80k 스텝): Action Expert(플로우) 추가 → 실행은 빠르고 정교",
  ], C.A);
  card(s, 6.8, 3.75, 5.9, 3.0, "Knowledge Insulation (2025.05)", [
    "VLM은 FAST 토큰으로 학습 (좋은 표현)",
    "Action Expert의 그래디언트가 VLM으로 가지 않게 차단 (stop-gradient)",
    "→ VLM 지식 보존 · 학습 빠름 · 언어 지시 따르기↑",
  ], C.L);
  s.addNotes(`[1:36] "π0.5의 두 번째 아이디어: 한 모델이 먼저 '다음에 뭘 할지'를 텍스트로 말하고, 그 다음 행동을 만듭니다. '침실 정리해' → '베개를 집어' → 행동."
"사람이 머릿속으로 계획하고 몸을 움직이는 것과 같습니다. 로봇판 chain-of-thought."
"세 번째: 학습은 토큰(FAST)으로 효율적으로, 실행은 플로우로 정교하게. Part 2 비교표의 하이브리드가 이것입니다."
Knowledge Insulation: "행동 전문가가 처음에 서툴게 학습할 때 나오는 그래디언트가 VLM을 망가뜨리지 않도록 막아버립니다. VLM의 지식을 '단열'시키는 거죠."
결과: 학습에 없던 3개의 실제 집에서 부엌·침실 정리. 2025.09 openpi에 가중치 공개.`);
}

// 44. 이중 시스템
{
  const s = content("이중 시스템 휴머노이드: Figure Helix & NVIDIA GR00T N1", "PART 3 · 논문");
  paperTag(s, "Figure 2025.02 · NVIDIA 2025.03");
  const cols = [
    ["Helix (Figure)", [["System 2", "7B 오픈 VLM · 7–9Hz\n장면·언어 이해 → 잠재 벡터"], ["System 1", "80M Transformer · 200Hz\n손가락까지 상체 35 자유도"]],
      ["~500시간 원격조종 데이터", "VLM이 영상을 보고 지시문을 역생성해 라벨로", "두 로봇 협업, 로봇 탑재 GPU에서 실행"]],
    ["GR00T N1 (NVIDIA)", [["System 2", "Eagle-2 VLM (1.34B) · ~10Hz\nLLM 중간층 특징 사용"], ["System 1", "DiT + 플로우 매칭 · 120Hz\n16스텝 청크 (~64ms, L40)"]],
      ["전체 2.2B · 오픈 가중치", "로봇별 인코더/디코더, 몸통 공유", "후속 N1.5(VLM 고정) · N1.6(Cosmos-Reason, 32층 DiT)"]],
  ];
  cols.forEach((c, i) => {
    const x = 0.6 + i * 6.2;
    T(s, c[0], { x, y: 1.8, w: 5.9, h: 0.45, fontSize: 19, bold: true });
    c[1].forEach((sys, j) => {
      const col = j === 0 ? C.L : C.A;
      box(s, sys[0], x, 2.4 + j * 1.25, 1.5, 1.05, { fill: { color: col }, color: C.white, bold: true, fontSize: 14 });
      box(s, sys[1], x + 1.6, 2.4 + j * 1.25, 4.3, 1.05, { fill: { color: j === 0 ? C.Lt : C.At }, fontSize: 13, align: "left", margin: 0.15 });
    });
    arrow(s, x + 0.75, 3.47, x + 0.75, 3.63, C.faint);
    bullets(s, c[2], x, 5.05, 5.9, 1.8, 13.5);
  });
  s.addNotes(`[1:38] "Part 2에서 본 이중 시스템의 대표 사례입니다."
Helix: "생각하는 7B VLM이 초당 7~9번 상황을 이해해서 잠재 벡터로 요약하고, 80M짜리 작은 모델이 초당 200번 손가락까지 움직입니다. 500시간 데이터, 두 로봇이 처음 보는 식료품을 협업해서 냉장고에 넣는 데모가 유명합니다." (가능하면 영상)
GR00T N1: "NVIDIA의 오픈 휴머노이드 모델. Eagle-2 VLM + 플로우 매칭 DiT, 120Hz."
공통 메시지: "큰 모델의 지능과 빠른 제어를 동시에 얻는 절충."`);
}

// 45. GR00T 데이터 피라미드
{
  const s = content("GR00T N1의 데이터 피라미드: 부족한 로봇 데이터를 채우는 법", "PART 3 · 논문");
  paperTag(s, "NVIDIA · 2025.03");
  const layers = [
    ["실제 로봇 데이터", "적음 · 비쌈 · 정확", C.A, 3.0],
    ["합성 데이터", "시뮬레이션(DexMimicGen) · 영상 생성 모델로 만든 궤적", C.L, 5.0],
    ["웹 데이터 + 사람 1인칭 영상", "많음 · 공짜 · 행동 라벨 없음", C.V, 7.0],
  ];
  layers.forEach((l, i) => {
    const w = l[3], x = 0.6 + (7.0 - w) / 2, y = 1.9 + i * 1.5;
    s.addText(l[0], { x, y, w, h: 1.3, shape: pres.shapes.RECTANGLE, fill: { color: l[2] }, line: { color: C.white }, color: C.white, fontFace: F, fontSize: 15, bold: true, align: "center", valign: "middle", margin: 0.05, isTextBox: true });
    T(s, l[1], { x: 7.9, y, w: 4.8, h: 1.3, fontSize: 14, valign: "middle", color: C.text });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.5, w: 12.1, h: 0.5, rectRadius: 0.06, fill: { color: C.card }, line: { color: C.card } });
  T(s, "행동 라벨 없는 영상은? → 잠재 행동(latent action): 앞뒤 프레임 차이에서 '가상의 행동'을 학습해 라벨로 사용", { x: 0.8, y: 6.5, w: 11.8, h: 0.5, fontSize: 13, valign: "middle", bold: true });
  s.addNotes(`[1:40] (시간 없으면 말로만) "NVIDIA다운 해법입니다. 로봇 데이터가 부족하면 시뮬레이션과 생성 영상, 사람 영상으로 채웁니다."
"피라미드 아래로 갈수록 많고 싸지만 행동 라벨이 없습니다. 사람 영상에는 '로봇 관절 각도'가 없으니까요. 그래서 영상 앞뒤 프레임 차이로 '가상의 행동'을 학습해서 라벨로 씁니다 (LAPA 계열)."`);
}

// 46. Gemini Robotics
{
  const s = content("Gemini Robotics → 1.5: 생각하고 행동하는 에이전트", "PART 3 · 논문");
  paperTag(s, "Google DeepMind · 2025.03 / 2025.09");
  box(s, "Gemini Robotics-ER 1.5\n\"지휘자\" · 계획 · 도구 사용(웹 검색)", 0.6, 1.95, 4.6, 1.3, { fill: { color: C.L }, color: C.white, bold: true, fontSize: 14 });
  arrow(s, 2.9, 3.3, 2.9, 3.75, C.L);
  T(s, "단계별 자연어 지시", { x: 3.05, y: 3.35, w: 2.5, h: 0.35, fontSize: 11, color: C.L, bold: true });
  box(s, "Gemini Robotics 1.5 (VLA)\n생각(텍스트) → 행동", 0.6, 3.8, 4.6, 1.3, { fill: { color: C.A }, color: C.white, bold: true, fontSize: 14 });
  T(s, "예: \"우리 동네 분리수거 규칙대로 쓰레기를 분류해\" → ER이 규칙을 검색 → VLA가 분류", { x: 0.6, y: 5.35, w: 4.6, h: 1.3, fontSize: 13, color: C.muted });
  card(s, 5.6, 1.85, 3.4, 4.9, "Gemini Robotics", [
    "2025.03 · Gemini 2.0 기반",
    "ER: 가리키기, 3D 박스, 잡을 위치 등 공간 추론",
    "클라우드 백본 + 로컬 액션 디코더, 청크로 50Hz",
    "종이접기 · 도시락 싸기",
  ], C.muted, { bs: 13 });
  card(s, 9.3, 1.85, 3.4, 4.9, "Gemini Robotics 1.5", [
    "2025.09 · Think before acting: 행동 전에 자연어로 추론",
    "Motion Transfer: ALOHA · Franka · 휴머노이드 Apollo 사이 동작 전이",
    "2026.04 ER 1.6 (계기판 읽기 등)",
  ], C.A, { bs: 13 });
  s.addNotes(`[1:42] "Google DeepMind의 접근입니다. 두 모델이 짝을 이룹니다."
"ER은 '지휘자'. 전체 계획을 세우고, 필요하면 웹 검색 같은 도구도 씁니다. 예를 들어 '우리 동네 분리수거 규칙대로 분류해'라고 하면 규칙을 검색합니다."
"VLA는 '손'. 1.5 버전부터는 행동 전에 자연어로 생각합니다. LLM의 reasoning 모델처럼요."
"Motion Transfer: 양팔 로봇에서 배운 동작이 휴머노이드로 전이됩니다. 한 모델로 여러 몸."
(가능하면 Gemini Robotics 1.5 영상)`);
}

// 47. π*0.6 RECAP
{
  const s = content("π*0.6 · RECAP: 경험과 교정으로 스스로 나아지는 VLA", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2025.11");
  const loop = [["① 수집", "시연 + 자율 시도\n+ 사람의 교정"], ["② 가치 함수", "\"얼마나 잘 되고\n있나\" 예측"], ["③ 어드밴티지", "각 행동이 평균보다\n나았나/못했나"], ["④ 조건부 학습", "\"Advantage: +/−\"\n태그 붙여 전부 학습"]];
  loop.forEach((l, i) => {
    const x = 0.6 + i * 3.1;
    box(s, [{ text: l[0], options: { bold: true, fontSize: 16, color: C.good, breakLine: true } }, { text: l[1], options: { fontSize: 12.5 } }], x, 1.9, 2.7, 1.5, { fill: { color: "E6F4EC" } });
    if (i < 3) arrow(s, x + 2.72, 2.65, x + 3.08, 2.65, C.good);
  });
  T(s, "↺ 반복 · 실행할 땐 항상 \"positive\"로 조건 → 좋은 행동 쪽에서 샘플", { x: 0.6, y: 3.5, w: 12.1, h: 0.4, fontSize: 14, bold: true, color: C.good });
  stat(s, 0.6, 4.1, 3.9, "2배+", "가장 어려운 작업의 처리량\n(빨래: 시간당 ~4개 → 8개+)", C.good);
  stat(s, 4.7, 4.1, 3.9, "~½", "실패율 감소\n(빨래 성공률 ~40% → 75%+)", C.good);
  card(s, 8.8, 4.1, 3.9, 2.7, "의미", "모방(시연자 수준에 갇힘) → 경험(스스로 개선). LLM의 RLHF · reasoning RL과 평행. 에스프레소를 하루 종일 연속 제조.", C.good, { bs: 13 });
  s.addNotes(`[1:44] "모방학습은 시연자 수준을 넘을 수 없고, 실수에서 배우지 못합니다. LLM이 RL로 한 단계 더 올라갔듯, 로봇도 RL이 필요합니다."
RECAP 쉽게: "코치가 경기 영상을 보면서 '이 동작 좋았어', '이건 별로'라고 태그를 붙여줍니다. 선수는 좋은 것, 나쁜 것 모두 태그와 함께 배우고, 실전에서는 '좋았던 동작'을 떠올리며 합니다."
1) 시연 + 로봇이 직접 해본 시도 + 사람의 교정 수집, 2) 가치 함수로 '얼마나 잘 되고 있나' 예측, 3) 각 행동이 평균보다 나았는지 계산, 4) 'Advantage: positive/negative' 태그를 붙여 학습. 실행할 땐 항상 positive.
결과: 처리량 2배 이상, 실패율 절반. 에스프레소를 하루 종일. (영상: pi.website/blog/pistar06)`);
}

// 48. π0.7
{
  const s = content("π0.7 (2026): 말로 코칭할 수 있는 \"조종 가능한\" 범용 모델", "PART 3 · 논문");
  paperTag(s, "Physical Intelligence · 2026.04");
  T(s, "프롬프트에 언어 지시만이 아니라 다양한 '컨텍스트'를 함께 넣는다", { x: 0.6, y: 1.8, w: 12, h: 0.45, fontSize: 17, bold: true });
  const ctx = [["하위작업 지시", "\"먼저 뚜껑을 열어\"", C.L, C.Lt], ["하위목표 이미지", "\"이렇게 된 상태로\"", C.V, C.Vt], ["에피소드 메타데이터", "품질 · 속도 · 전략", C.A, C.At]];
  ctx.forEach((c, i) => {
    box(s, [{ text: c[0], options: { bold: true, fontSize: 17, color: c[2], breakLine: true } }, { text: c[1], options: { fontSize: 13.5 } }], 0.6 + i * 4.1, 2.45, 3.8, 1.35, { fill: { color: c[3] } });
  });
  bullets(s, [
    "서툰 시연 · 실패 · 이전 모델 궤적까지 \"이건 어떤 품질의 데이터\"라는 태그와 함께 학습 → 실행 시 원하는 품질 지정",
    "5B 파라미터 · 범용 모델로 전문가(파인튜닝) 모델 수준 성능",
    "조합적 일반화: 배운 기술을 재조합 — 말로 단계별 코칭받아 에어프라이어로 고구마 조리",
    "아직: 처음 보는 작업 60–80% vs 학습 분포 내 작업 >90%",
  ], 0.6, 4.15, 12.1, 2.7, 15);
  s.addNotes(`[1:46] "가장 최근, 2026년 4월 모델입니다. 핵심은 '조종 가능성(steerability)'."
"언어 지시뿐 아니라 하위작업 지시, 목표 상태 이미지, '이 데이터는 빠르고 깔끔한 시연이었다' 같은 메타데이터를 함께 넣어 학습합니다. 그러면 품질이 섞인 데이터도 버리지 않고 쓸 수 있고, 실행할 때 '깔끔하게'를 지정할 수 있습니다."
"결과: 처음 보는 작업도 사람이 말로 단계별로 코칭하면 해냅니다. 로봇에게도 프롬프트 엔지니어링이 가능한 시대."
주의: 처음 보는 작업은 아직 60–80% 수준.`);
}

// 49. 비교표
{
  const s = content("주요 VLA 한눈에 비교", "PART 3 · 논문");
  tbl(s, [
    ["모델", "연도", "크기", "백본", "행동 출력", "공개"],
    ["RT-1", "2022", "35M", "EfficientNet + Transformer", "이산 토큰", "일부"],
    ["RT-2", "2023", "55B", "PaLI-X / PaLM-E", "이산 토큰 (텍스트)", "✕"],
    ["Octo", "2024", "93M", "Transformer", "디퓨전", "○"],
    ["OpenVLA", "2024", "7B", "Llama 2 + SigLIP/DINOv2", "이산 토큰", "○"],
    ["π0", "2024", "3.3B", "PaliGemma + Action Expert", "플로우 매칭", "○"],
    ["OpenVLA-OFT", "2025", "7B", "OpenVLA", "연속 회귀 + 청킹", "○"],
    ["Helix", "2025", "7B + 80M", "이중 시스템", "연속 (S1)", "✕"],
    ["GR00T N1", "2025", "2.2B", "Eagle-2 + DiT", "플로우 매칭", "○"],
    ["π0.5", "2025", "~3B", "PaliGemma + AE", "FAST → 플로우", "○"],
    ["Gemini Robotics 1.5", "2025", "비공개", "Gemini", "비공개", "✕"],
    ["SmolVLA", "2025", "0.45B", "SmolVLM2 + AE", "플로우 매칭", "○"],
  ], { x: 0.6, y: 1.75, w: 12.1, colW: [2.4, 0.9, 1.3, 3.6, 2.8, 1.1], fontSize: 12.5, rowH: 0.42 });
  s.addNotes(`[1:48] Part 3 정리. "표를 다 읽을 필요는 없습니다. 두 가지 흐름만 보세요."
1) 행동 출력: 이산 토큰(RT-1/RT-2/OpenVLA) → 플로우 매칭·연속(π0, GR00T) → 하이브리드(π0.5).
2) 공개: 초기엔 비공개 → OpenVLA 이후 오픈소스가 활발 (openpi, GR00T, SmolVLA).
SmolVLA: Hugging Face의 450M 소형 모델, 노트북 GPU에서도 돌아감 — 실습 시작점으로 추천.`);
}

// =====================================================================
// PART 4
// =====================================================================
section(4, "평가 · 한계 · 미래", "그래서 이제 다 된 걸까?", "약 10분 · 슬라이드 51–54",
  `[1:48] Part 4. "그래서 이제 로봇 문제는 다 풀린 걸까요? 아닙니다. 어떻게 평가하고, 무엇이 남았는지 봅시다."`);

// 51. 평가
{
  const s = content("어떻게 평가하나: 벤치마크와 그 한계", "PART 4 · 평가와 미래");
  s.addChart(pres.charts.BAR, [{ name: "LIBERO 평균 성공률 (%)", labels: ["Diffusion Policy", "Octo", "OpenVLA", "π0-FAST", "π0", "OpenVLA-OFT"], values: [72.4, 75.1, 76.5, 85.5, 94.2, 97.1] }], {
    x: 0.5, y: 1.7, w: 6.6, h: 4.6, barDir: "bar", chartColors: [C.A], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12,
    dataLabelColor: C.text, catAxisLabelFontSize: 12, catAxisLabelColor: C.text, catAxisLabelFontFace: F, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 110,
    valGridLine: { style: "none" }, showLegend: false, showTitle: true, title: "LIBERO 평균 성공률 (%, OpenVLA-OFT 논문 기준)", titleFontSize: 12, titleColor: C.muted, titleFontFace: F, barGapWidthPct: 50, dataLabelFormatCode: "0.0",
  });
  tbl(s, [
    ["벤치마크", "특징"],
    ["LIBERO", "시뮬레이션, 4개 스위트 · 포화 상태"],
    ["SimplerEnv", "실제 로봇과 상관된 시뮬 평가"],
    ["CALVIN", "긴 연속 작업 (몇 개 연속 성공?)"],
    ["RoboTwin 2.0", "양팔 조작 시뮬"],
    ["RoboArena 등", "여러 기관 분산 실세계 평가"],
  ], { x: 7.4, y: 1.8, w: 5.3, colW: [1.8, 3.5], fontSize: 12.5, rowH: 0.5 });
  T(s, "시뮬 97%여도 실세계는 다르다 → 결국 실제 로봇 평가가 필요하지만 비싸고 재현이 어렵다", { x: 7.4, y: 5.1, w: 5.3, h: 1.2, fontSize: 14, bold: true, color: C.bad });
  s.addNotes(`[1:50] "대표 벤치마크 LIBERO입니다. 보시다시피 97%까지 올라와서 거의 포화됐습니다."
"문제는 시뮬레이션 97%가 실세계 97%가 아니라는 것. 실제 평가는 로봇으로 하나하나 해봐야 해서 비싸고, 연구실마다 환경이 달라 재현이 어렵습니다."
"그래서 SimplerEnv(실제와 상관된 시뮬), RoboArena(여러 기관이 분산해서 실제 평가) 같은 시도가 나옵니다. 평가 자체가 연구 주제입니다."`);
}

// 52. 남은 과제
{
  const s = content("아직 남은 과제", "PART 4 · 평가와 미래");
  const probs = [
    ["데이터", "여전히 인터넷보다 수만 배 적다. 사람 영상 · 시뮬 · 월드 모델로 확장 중", C.A],
    ["속도 · 효율", "큰 모델 = 느림. 엣지 배포, 소형화, 비동기 추론", C.A],
    ["신뢰성", "데모는 되지만 99.9%는 아직. 긴 작업에서 누적 실패", C.bad],
    ["안전", "VLA 자체엔 안전 보장 없음 → 외부 안전장치, 안전 추론 연구", C.bad],
    ["정교함 · 감각", "촉각 · 힘 제어, 손가락 조작은 초기 단계", C.L],
    ["평가", "재현 가능한 실세계 평가 방법 부재", C.L],
  ];
  probs.forEach((p, i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.85 + Math.floor(i / 3) * 2.5;
    card(s, x, y, 3.8, 2.2, p[0], p[1], p[2], { hs: 19, bs: 14 });
  });
  s.addNotes(`[1:52] 6가지를 빠르게.
- 데이터: 여전히 LLM 대비 턱없이 적음.
- 속도: 큰 모델은 느림.
- 신뢰성: 데모 영상은 잘 된 것만 보여줍니다. 공장·가정에서 쓰려면 99.9%가 필요한데 아직 멀었습니다.
- 안전: VLA 자체엔 안전 보장이 없어서 바깥에 전통적인 안전장치(힘 제한, 비상정지)를 둡니다.
- 촉각·힘: 대부분 카메라만 씁니다.
- 평가: 앞 슬라이드.
"여러분이 연구 주제를 찾는다면 이 6개 중 하나를 고르면 됩니다."`);
}

// 53. 트렌드
{
  const s = content("2025–26 트렌드: VLA는 어디로 가고 있나", "PART 4 · 평가와 미래");
  const tr = [
    ["모방 → 경험 (RL)", "로봇이 직접 해보며 개선 · π*0.6 RECAP, 실세계 RL", C.good],
    ["생각하는 VLA", "행동 전 추론 · ECoT, π0.5 하위작업 예측, Gemini Robotics 1.5", C.L],
    ["에이전트 + VLA", "고수준 VLM(계획·도구) + 저수준 VLA(동작) · Gemini ER, Hi Robot", C.L],
    ["데이터 확장", "사람 1인칭 영상 · 시뮬 · 영상 생성 월드 모델 · 잠재 행동", C.V],
    ["작고 빠르고 열린 모델", "SmolVLA(450M) · openpi · LeRobot · GR00T 오픈 가중치", C.A],
    ["조종 가능성 · 코칭", "다양한 컨텍스트로 조건화, 말로 가르치기 · π0.7", C.A],
  ];
  tr.forEach((t, i) => {
    const y = 1.8 + i * 0.85;
    dot(s, 0.65, y + 0.2, 0.32, t[2]);
    T(s, t[0], { x: 1.2, y, w: 3.6, h: 0.72, fontSize: 17, bold: true, valign: "middle" });
    T(s, t[1], { x: 4.9, y, w: 7.8, h: 0.72, fontSize: 14.5, color: C.muted, valign: "middle" });
  });
  s.addNotes(`[1:54] "LLM이 걸어온 길을 로봇이 따라가고 있습니다: 사전학습 → 지시 따르기 → 추론(reasoning) → RL → 에이전트."
하나씩 한 줄로. 강조: "모방에서 경험으로", "생각하고 행동하기", "작고 열린 모델로 누구나 시작 가능".`);
}

// 54. 직접 해보기
{
  const s = content("직접 해보기: 오늘 밤부터 시작할 수 있는 것", "PART 4 · 평가와 미래");
  const st = [
    ["LeRobot 설치", "Hugging Face의 로봇 학습 라이브러리.\n데이터셋 · 학습 · 평가 · 시각화"],
    ["사전학습 모델 받기", "SmolVLA (노트북 GPU) · π0/π0.5 (openpi) · OpenVLA · GR00T"],
    ["시뮬에서 파인튜닝", "LIBERO / SimplerEnv 데이터로 LoRA 파인튜닝"],
    ["실제 로봇으로", "SO-101 같은 저가 팔 + 원격조종으로 50개 시연 → 파인튜닝"],
  ];
  st.forEach((x, i) => {
    const y = 1.85 + i * 1.2;
    numCircle(s, i + 1, 0.6, y + 0.1, 0.7, C.A, 20);
    T(s, x[0], { x: 1.55, y, w: 3.6, h: 0.9, fontSize: 18, bold: true, valign: "middle" });
    T(s, x[1], { x: 5.2, y, w: 7.5, h: 0.9, fontSize: 14, color: C.muted, valign: "middle" });
  });
  T(s, "github.com/huggingface/lerobot  ·  github.com/Physical-Intelligence/openpi  ·  github.com/openvla/openvla", { x: 0.6, y: 6.8, w: 12.1, h: 0.35, fontSize: 12, color: C.muted, fontFace: MONO });
  s.addNotes(`[1:55] (시간 없으면 자료 공유로 대체) "관심 있으시면 오늘 밤부터 시작할 수 있습니다."
"LeRobot 설치 → SmolVLA 같은 작은 사전학습 모델 → 시뮬레이션 데이터로 파인튜닝 → 저가 로봇 팔로 실제 시연 50개."
"작업 하나 파인튜닝하는 데 필요한 건 수십~수백 개 시연, GPU 한 장입니다."`);
}

// 55. 요약
{
  const s = dark();
  T(s, "오늘의 핵심 5줄", { x: 0.8, y: 0.7, w: 11, h: 0.8, fontSize: 34, bold: true, color: C.white });
  const k = [
    ["VLA = VLM(눈+언어)에 로봇의 손(행동)을 붙인 모델", C.V],
    ["인터넷 지식이 로봇으로 전이된다 (RT-2: 멸종된 동물 → 공룡)", C.L],
    ["핵심 설계 = 행동 출력 방식: 토큰 · 회귀 · 디퓨전/플로우 · FAST", C.A],
    ["정교함은 액션 청킹 + 플로우(π0), 일반화는 데이터 다양성(π0.5)", C.A],
    ["다음 단계: 생각하는 VLA · 경험으로 배우는 VLA(RL) · 말로 코칭", C.good],
  ];
  k.forEach((x, i) => {
    const y = 1.8 + i * 0.95;
    numCircle(s, i + 1, 0.8, y + 0.05, 0.6, x[1], 18);
    T(s, x[0], { x: 1.65, y, w: 11, h: 0.7, fontSize: 20, color: C.white, valign: "middle" });
  });
  s.addNotes(`[1:56] 5줄을 천천히 읽기. "오늘 이 5줄만 기억하시면 됩니다."
처음 학습목표 슬라이드와 연결: "처음에 말씀드린 세 가지 — VLA가 뭔지, 행동 출력 4가지, 논문 계보 — 이제 설명하실 수 있을 겁니다."`);
}

// 56. Q&A
{
  const s = dark();
  T(s, "Q & A", { x: 0.8, y: 1.6, w: 11, h: 1.3, fontSize: 60, bold: true, color: C.white });
  vlaDots(s, 0.8, 3.1);
  T(s, "추천 읽을거리", { x: 0.8, y: 3.7, w: 11, h: 0.4, fontSize: 16, bold: true, color: C.A });
  T(s, "RT-2 (2307.15818) · OpenVLA (2406.09246) · π0 (2410.24164) · π0.5 (2504.16054)\n서베이: A Survey on VLA Models — An Action Tokenization Perspective (2507.01925)\n실습: LeRobot · openpi", { x: 0.8, y: 4.2, w: 11.8, h: 1.4, fontSize: 15, color: "D7DCE3", lineSpacingMultiple: 1.3 });
  T(s, "감사합니다", { x: 0.8, y: 6.3, w: 6, h: 0.5, fontSize: 18, color: "8A93A0" });
  s.addNotes(`[1:57–2:00] 질문 받기. 예상 질문과 답변은 04_예상질문_퀴즈.md 참고.
모르는 질문: "좋은 질문입니다. 원문을 확인해서 자료 공유할 때 같이 답변드리겠습니다." — 지어내지 않기.
자료 공유: 05_참고자료.md 링크 모음 + PPT PDF.`);
}

pres.writeFile({ fileName: __dirname + "/VLA_Basics_Lecture.pptx" }).then((f) => console.log("wrote", f, "slides:", pageNo));
