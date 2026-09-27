// VLA 특론 4주차 슬라이드 생성 스크립트
// 실행: npm install pptxgenjs && node build_week04.js → Week04_Imitation_Learning.pptx
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
// 4주차 슬라이드 본문
// =====================================================================
pres.title = "VLA 특론 4주차 — 모방학습과 세 개의 정책";
const path = require("path");
const DEMO_IMG = path.join(__dirname, "..", "demo", "multimodality_demo.png");
const BET = C.L, DP = C.A, ACT = C.V;           // 세 해법의 색 약속
const BETt = C.Lt, DPt = C.At, ACTt = C.Vt;
const S_ENEMY = "두 개의 적", S_SOL = "세 개의 해법";

// 1. 표지
{
  const s = dark();
  vlaDots(s, 0.8, 1.2);
  T(s, "VLA 특론 · 4주차", { x: 0.8, y: 1.7, w: 10, h: 0.6, fontSize: 22, color: "B8C0CC" });
  T(s, "모방학습과 세 개의 정책", { x: 0.8, y: 2.35, w: 12, h: 1.0, fontSize: 46, bold: true, color: C.white, valign: "middle" });
  s.addText([
    { text: "BeT", options: { color: BET } }, { text: "  ·  ", options: { color: "8A93A0" } },
    { text: "Diffusion Policy", options: { color: DP } }, { text: "  ·  ", options: { color: "8A93A0" } },
    { text: "ACT", options: { color: ACT } },
  ], { x: 0.8, y: 3.45, w: 12, h: 0.7, fontFace: F, fontSize: 28, bold: true, margin: 0, valign: "middle", isTextBox: true });
  T(s, "VLA 논문을 읽기 위한 문법 시간 — 누적 오차, 다봉성, action chunking", { x: 0.8, y: 4.35, w: 12, h: 0.5, fontSize: 18, color: "D7DCE3" });
  T(s, "140분 · 읽을 논문: arXiv 2206.11251 · 2303.04137 · 2304.13705", { x: 0.8, y: 6.0, w: 10, h: 0.4, fontSize: 13, color: "8A93A0" });
  s.addNotes(`[0:00] 오프닝.
"지난 2주 동안 그림 그리는 기술, 생성모델을 배웠습니다. 오늘부터 드디어 로봇입니다. 오늘 수업은 사실상 VLA 논문을 읽기 위한 '문법 시간'입니다."
색 약속: 오늘은 BeT=보라, Diffusion Policy=주황, ACT=청록, '적'은 빨강으로 통일했습니다.`);
}

// 2. 오늘의 지도
{
  const s = content("오늘의 지도 (140분)", "OPENING");
  const blocks = [
    ["0:00", "복습 · 용어", "10", C.faint],
    ["0:10", "적 1호\n누적 오차", "25", C.bad],
    ["0:35", "적 2호\n다봉성 + chunking", "20", C.bad],
    ["0:55", "휴식", "10", C.faint],
    ["1:05", "해법 1\nBeT", "20", BET],
    ["1:25", "해법 2\nDiffusion Policy", "25", DP],
    ["1:50", "해법 3\nACT", "20", ACT],
    ["2:10", "비교 · 예고", "10", C.ink3],
  ];
  const x0 = 0.6, gw = 1.42, gap = 0.1;
  blocks.forEach((b, i) => {
    const x = x0 + i * (gw + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.95, w: gw, h: 3.0, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
    dot(s, x + 0.2, 2.2, 0.26, b[3]);
    T(s, b[0], { x: x + 0.2, y: 2.62, w: gw - 0.3, h: 0.3, fontSize: 12, bold: true, color: C.muted });
    T(s, b[1], { x: x + 0.2, y: 3.0, w: gw - 0.3, h: 1.1, fontSize: 14.5, bold: true });
    T(s, b[2] + "분", { x: x + 0.2, y: 4.25, w: gw - 0.3, h: 0.5, fontSize: 20, bold: true, color: b[3] === C.faint ? C.muted : b[3] });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.3, w: 12.1, h: 1.3, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "오늘의 한 문장", { x: 0.85, y: 5.42, w: 4, h: 0.35, fontSize: 12, bold: true, color: C.A });
  T(s, "모방학습은 지도학습인데, 로봇에서는 누적 오차와 다봉성 때문에 실패한다.\nBeT · DP · ACT는 이산화 · 생성 · CVAE로 다봉성을 풀었고, 셋 다 action chunking을 썼다.", { x: 0.85, y: 5.78, w: 11.6, h: 0.8, fontSize: 15, color: C.white, bold: true });
  s.addNotes(`[0:01] 오늘의 흐름 소개. "앞 55분은 문제(두 개의 적), 휴식 후 65분은 해법(세 논문)입니다."
아래 검은 박스의 한 문장이 오늘의 결론입니다. 수업 끝에 다시 이 문장으로 돌아옵니다.`);
}

// 3. 네 단어
{
  const s = content("오늘의 문법: 네 단어", "OPENING");
  const words = [
    ["모방학습 (BC)", "사람 시범을 지도학습으로 따라 하기", C.ink3],
    ["누적 오차", "작은 실수 → 처음 보는 상태 → 더 큰 실수", C.bad],
    ["다봉성", "정답 행동이 여러 개 → 평균은 오답", C.bad],
    ["Action Chunking", "미래 행동 k개를 한 번에 예측", C.good],
  ];
  words.forEach((w, i) => {
    const x = 0.6 + i * 3.08;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.95, w: 2.85, h: 3.4, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
    numCircle(s, i + 1, x + 0.25, 2.2, 0.6, w[2], 18);
    T(s, w[0], { x: x + 0.25, y: 3.0, w: 2.4, h: 0.8, fontSize: 18, bold: true });
    T(s, w[1], { x: x + 0.25, y: 3.85, w: 2.4, h: 1.3, fontSize: 14, color: C.muted, lineSpacingMultiple: 1.2 });
  });
  T(s, "→ 앞으로 읽을 모든 VLA 논문의 서론과 방법에 이 네 단어가 나온다", { x: 0.6, y: 5.75, w: 12.1, h: 0.5, fontSize: 17, bold: true });
  s.addNotes(`[0:02] "오늘 배우는 네 단어 — 모방학습, 누적 오차, 다봉성, action chunking — 가 앞으로 읽을 모든 VLA 논문의 서론과 방법에 나옵니다. 문법 시간이라고 한 이유입니다."`);
}

// 4. 용어 정의
{
  const s = content("용어 정의: 정책 · 시범 · 모방학습", "PART A · 복습과 용어");
  const defs = [
    ["정책 π(a|s)", "관측 s (카메라 이미지, 관절 각도)를 받아 행동 a를 내는 함수 = 로봇의 두뇌"],
    ["시범 (demonstration)", "사람이 원격조종(teleoperation)으로 로봇을 움직이며 기록한 (s, a) 시퀀스"],
    ["모방학습 / BC", "시범으로 π를 지도학습 — \"s를 보면 사람이 했던 a를 내라\""],
  ];
  defs.forEach((d, i) => {
    const y = 1.9 + i * 1.0;
    box(s, d[0], 0.6, y, 3.3, 0.8, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 15 });
    T(s, d[1], { x: 4.15, y, w: 8.5, h: 0.8, fontSize: 16, valign: "middle" });
  });
  T(s, "원격조종으로 시범을 모으는 방법", { x: 0.6, y: 5.1, w: 8, h: 0.4, fontSize: 15, bold: true, color: C.muted });
  const tele = [["리더–팔로워 팔", "사람이 작은 팔을 움직이면 로봇 팔이 따라함 (ALOHA)"], ["VR 컨트롤러", "손 위치를 로봇 손끝으로"], ["스페이스마우스", "3D 마우스로 손끝 이동"]];
  tele.forEach((t, i) => card(s, 0.6 + i * 4.1, 5.55, 3.8, 1.35, t[0], t[1], C.V, { fill: C.Vt, hs: 15, bs: 12.5 }));
  s.addNotes(`[0:03–0:08] 칠판에 정의 3줄을 씁니다 (칠판 왼쪽 구역).
정책: 로봇의 두뇌. 시범: 사람이 원격조종하며 기록한 (관측, 행동) 쌍의 시퀀스. BC: 그걸 지도학습.
원격조종 방법 3가지를 짧게 — 특히 리더-팔로워(ALOHA)는 오늘 ACT에서 다시 나옵니다.`);
}

// 5. BC = 지도학습
{
  const s = content("BC는 그냥 지도학습이다… 그럼 끝난 것 아닌가?", "PART A · 복습과 용어");
  box(s, "고양이 사진", 0.6, 2.0, 2.4, 0.9, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 15 });
  arrow(s, 3.05, 2.45, 3.75, 2.45);
  box(s, "분류기", 3.8, 2.0, 1.8, 0.9, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 15 });
  arrow(s, 5.65, 2.45, 6.35, 2.45);
  box(s, "\"고양이\"", 6.4, 2.0, 2.0, 0.9, { fill: { color: C.card }, bold: true, fontSize: 15 });
  box(s, "카메라 화면\n+ 관절 각도", 0.6, 3.3, 2.4, 0.9, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 14 });
  arrow(s, 3.05, 3.75, 3.75, 3.75);
  box(s, "정책 π", 3.8, 3.3, 1.8, 0.9, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 15 });
  arrow(s, 5.65, 3.75, 6.35, 3.75, C.A);
  box(s, "사람이 그때\n움직인 방향", 6.4, 3.3, 2.0, 0.9, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 14 });
  T(s, "이미지 분류", { x: 8.7, y: 2.0, w: 3.5, h: 0.9, fontSize: 15, color: C.muted, valign: "middle" });
  T(s, "모방학습 (BC)", { x: 8.7, y: 3.3, w: 3.5, h: 0.9, fontSize: 15, color: C.muted, valign: "middle" });
  code(s, "min_θ  E_(s,a)~D [ −log π_θ(a | s) ]     또는     || π_θ(s) − a ||²", 0.6, 4.65, 12.1, 0.75, 16);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.7, w: 12.1, h: 1.0, rectRadius: 0.08, fill: { color: "FDECEC" }, line: { color: "FDECEC" } });
  T(s, "그런데 로봇에서는 이 '그냥 지도학습'이 두 가지 방식으로 처참하게 실패한다 → 두 개의 적", { x: 0.85, y: 5.7, w: 11.6, h: 1.0, fontSize: 17, bold: true, color: C.bad, valign: "middle" });
  s.addNotes(`[0:08–0:10] "이미지 분류와 똑같습니다. 입력이 사진 대신 카메라 화면, 정답이 '고양이' 대신 '사람이 그때 움직인 방향'일 뿐."
"그럼 끝난 것 아닌가요?" — 잠깐 멈추고 반응을 봅니다. 반박하는 학생이 있으면 "좋습니다, 그게 오늘 주제"라고 받기.`);
}

// 6. 섹션: 두 개의 적
section("①", "두 개의 적", "누적 오차(covariate shift)와 다봉성(multimodality)", "0:10 – 0:55 · 슬라이드 7–16",
  `[0:10] "지도학습이 로봇에서 실패하는 두 가지 방식을 보겠습니다."`);

// 7. 적 1: 대전제
{
  const s = content("적 1호 · 누적 오차: 로봇은 자기 시험 문제를 스스로 만든다", S_ENEMY);
  T(s, "지도학습의 대전제: 훈련 분포 = 테스트 분포", { x: 0.6, y: 1.75, w: 12, h: 0.45, fontSize: 17, bold: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.35, w: 7.6, h: 4.3, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 3.2, w: 6.6, h: 0, line: { color: C.good, width: 4 } });
  T(s, "사람 시범 궤적", { x: 5.8, y: 2.75, w: 1.9, h: 0.35, fontSize: 12, bold: true, color: C.good });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 3.2, w: 1.6, h: 0.25, line: { color: C.bad, width: 4 } });
  s.addShape(pres.shapes.LINE, { x: 2.6, y: 3.45, w: 1.7, h: 0.65, line: { color: C.bad, width: 4 } });
  s.addShape(pres.shapes.LINE, { x: 4.3, y: 4.1, w: 1.8, h: 1.35, line: { color: C.bad, width: 4, endArrowType: "triangle" } });
  T(s, "✗", { x: 6.05, y: 5.3, w: 0.5, h: 0.5, fontSize: 24, bold: true, color: C.bad });
  T(s, "작은 오차 ε", { x: 2.0, y: 3.55, w: 1.6, h: 0.3, fontSize: 11.5, color: C.bad });
  T(s, "처음 보는 상태", { x: 3.3, y: 4.3, w: 1.8, h: 0.3, fontSize: 11.5, color: C.bad });
  T(s, "더 처음 보는 상태", { x: 4.1, y: 5.35, w: 1.9, h: 0.3, fontSize: 11.5, color: C.bad });
  bullets(s, [
    "정책의 행동이 다음 관측을 결정한다",
    "조금 틀림 → 시범에 없던 상태 → 거기선 더 틀림 → 눈덩이",
    "시험 문제를 자기 답이 만든다",
  ], 8.5, 2.45, 4.3, 3.0, 15);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.5, y: 5.2, w: 4.2, h: 1.45, rectRadius: 0.08, fill: { color: "FDECEC" }, line: { color: "FDECEC" } });
  T(s, "이름: covariate shift\n(입력 분포가 학습 때와 달라짐)", { x: 8.7, y: 5.2, w: 3.9, h: 1.45, fontSize: 13.5, bold: true, color: C.bad, valign: "middle" });
  s.addNotes(`[0:10–0:17] "지도학습에는 숨은 대전제가 있습니다. 훈련 데이터와 테스트 데이터가 같은 분포. 로봇은 이걸 스스로 깹니다. 다음 시험 문제를 자기 답이 만들기 때문입니다."
칠판 왼쪽에 이 그림을 그리며: 첫 스텝의 작은 오차 → 시범에 없던 상태 → 그 상태에서는 오차가 더 큼 → …`);
}

// 8. 운전 비유 + T²
{
  const s = content("왜 이렇게 나쁜가: 복구하는 법이 데이터에 없다", S_ENEMY);
  card(s, 0.6, 1.85, 6.0, 3.3, "비유 · 운전 교습", [
    "교습 영상은 전부 차선 한가운데에서 찍혔다",
    "영상만 보고 배운 초보가 살짝 차선을 벗어나면?",
    "\"벗어났을 때 돌아오는 장면\"은 영상에 한 번도 없었다",
    { t: "→ 시범자는 실수를 안 하므로, 실수 복구법이 데이터에 없다", b: true },
  ], C.bad, { bs: 14.5 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.9, y: 1.85, w: 5.8, h: 3.3, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "이론 한 줄 (Ross & Bagnell, 2010)", { x: 7.15, y: 2.05, w: 5.3, h: 0.4, fontSize: 15, bold: true, color: C.muted });
  T(s, "한 스텝 실수 확률 ε, 에피소드 길이 T", { x: 7.15, y: 2.6, w: 5.3, h: 0.4, fontSize: 14 });
  stat(s, 7.15, 3.0, 2.5, "εT", "일반 지도학습", C.good);
  stat(s, 9.9, 3.0, 2.6, "εT²", "BC (최악의 경우)", C.bad);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.45, w: 12.1, h: 1.25, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "학생에게: \"그럼 어떻게 해결할까요?\"  →  \"실수 장면도 넣자\" · \"사람이 계속 고쳐주자\"", { x: 0.85, y: 5.45, w: 11.6, h: 1.25, fontSize: 16, bold: true, color: C.white, valign: "middle" });
  s.addNotes(`[0:17–0:22] 운전 교습 비유가 핵심입니다. "시범자는 실수를 안 하기 때문에, 실수에서 복구하는 법이 데이터에 없다."
이론은 말로만: "한 번 실수하면 이후 남은 스텝 전부에서 틀릴 수 있어서, 오차가 선형이 아니라 제곱으로 나빠진다는 고전 결과가 있습니다." 증명은 하지 않습니다.
그리고 질문을 던져 30초 답을 받습니다. 보통 나오는 두 답이 곧 다음 슬라이드의 두 해법입니다.`);
}

// 9. DAgger
{
  const s = content("고전적 해법: DAgger — 로봇이 간 곳에서 라벨을 받아라", S_ENEMY);
  paperTag(s, "Ross, Gordon, Bagnell · 2011");
  const steps = [["①", "시범으로\nπ 학습"], ["②", "π를 실제로 굴림\n→ 이상한 상태 도달"], ["③", "전문가에게 \"여기선\n뭘 해야 하나요?\""], ["④", "데이터에 추가\n→ 재학습"]];
  const pos = [[0.9, 2.0], [7.4, 2.0], [7.4, 4.0], [0.9, 4.0]];
  steps.forEach((st, i) => {
    box(s, [{ text: st[0] + " ", options: { bold: true, color: C.A, fontSize: 18 } }, { text: st[1], options: { fontSize: 15, bold: true } }], pos[i][0], pos[i][1], 4.9, 1.4, { fill: { color: C.card } });
  });
  arrow(s, 5.85, 2.7, 7.35, 2.7, C.A, 2.5);
  arrow(s, 9.85, 3.45, 9.85, 3.95, C.A, 2.5);
  arrow(s, 7.35, 4.7, 5.85, 4.7, C.A, 2.5);
  arrow(s, 3.35, 3.95, 3.35, 3.45, C.A, 2.5);
  card(s, 0.6, 5.7, 5.9, 1.2, "장점", "로봇이 실제로 가는 상태 분포에서 라벨 → 원리적으로 옳다", C.good, { hs: 15, bs: 13 });
  card(s, 6.8, 5.7, 5.9, 1.2, "단점", "전문가가 로봇 옆에 계속 대기해야 한다 → 비싸다", C.bad, { hs: 15, bs: 13 });
  s.addNotes(`[0:22–0:28] DAgger(Dataset Aggregation). 루프를 칠판에 그리며 설명.
"로봇이 실제로 가는 상태에서 라벨을 받으니 원리적으로 옳습니다. 문제는 로봇 한 대에 사람 한 명이 붙어 있어야 한다는 것."
14주차 연결: "이 '사람이 옆에서 교정' 정신은 HIL-SERL, π*0.6의 교정 데이터로 부활합니다."
(시간이 밀리면 이 슬라이드는 한 줄로 넘기기)`);
}

// 10. 현대적 처방
{
  const s = content("현대적 처방 두 가지", S_ENEMY);
  card(s, 0.6, 1.9, 5.9, 2.5, "처방 1 · 데이터를 다양하게", [
    "실수 · 복구 장면까지 포함해 많이 모은다",
    "여러 사람, 여러 환경, 여러 로봇",
    { t: "→ 9–11주차 데이터 스케일링의 확장판", b: true },
  ], C.V, { fill: C.Vt, bs: 15 });
  card(s, 6.8, 1.9, 5.9, 2.5, "처방 2 · 여러 스텝을 한 번에 예측", [
    "한 스텝씩 결정하면 오차가 끼어들 기회가 T번",
    "k스텝씩 결정하면 T/k번으로 줄어든다",
    { t: "→ 잠시 후 action chunking", b: true },
  ], C.good, { fill: "E6F4EC", bs: 15 });
  T(s, "체크: 왜 사람 시범 데이터에는 복구 장면이 없나요?  →  사람은 잘하니까", { x: 0.6, y: 4.8, w: 12.1, h: 0.5, fontSize: 16, bold: true, color: C.muted });
  s.addNotes(`[0:28–0:35] 두 처방을 짧게. 처방 2가 곧 action chunking이라는 것을 예고만 하고 넘어갑니다.
체크 질문을 던지고 "사람은 잘하니까"라는 답이 나오면 적 1호 통과.`);
}

// 11. 적 2: 장애물
{
  const s = content("적 2호 · 다봉성: 평균은 정답이 아니다", S_ENEMY);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.8, w: 6.4, h: 4.9, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  dot(s, 3.6, 2.1, 0.4, C.good);
  T(s, "목표", { x: 4.1, y: 2.1, w: 1, h: 0.4, fontSize: 13, bold: true, color: C.good, valign: "middle" });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: 3.4, w: 1.0, h: 1.0, fill: { color: C.ink3 }, line: { color: C.ink3 } });
  T(s, "장애물", { x: 3.3, y: 3.4, w: 1.0, h: 1.0, fontSize: 12, align: "center", valign: "middle", bold: true, color: C.white });
  dot(s, 3.6, 5.8, 0.4, C.A);
  T(s, "로봇", { x: 4.1, y: 5.8, w: 1, h: 0.4, fontSize: 13, bold: true, color: C.A, valign: "middle" });
  arrow(s, 3.7, 5.8, 1.9, 3.9, C.good, 3); arrow(s, 1.9, 3.8, 3.6, 2.55, C.good, 3);
  arrow(s, 3.9, 5.8, 5.7, 3.9, C.good, 3); arrow(s, 5.7, 3.8, 4.0, 2.55, C.good, 3);
  arrow(s, 3.8, 5.75, 3.8, 4.45, C.bad, 4);
  T(s, "왼쪽 50%", { x: 0.9, y: 4.4, w: 1.4, h: 0.3, fontSize: 12, bold: true, color: C.good });
  T(s, "오른쪽 50%", { x: 5.5, y: 4.4, w: 1.4, h: 0.3, fontSize: 12, bold: true, color: C.good });
  T(s, "평균 = 충돌!", { x: 4.3, y: 4.95, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: C.bad });
  T(s, "MSE 회귀가 배우는 것", { x: 7.4, y: 1.9, w: 5.3, h: 0.45, fontSize: 17, bold: true });
  code(s, "argmin_f E|| f(s) − a ||²  =  E[a | s]", 7.4, 2.45, 5.3, 0.7, 15);
  bullets(s, [
    "MSE의 최적해는 조건부 평균",
    "정답 분포의 봉우리가 둘이면, 평균은 두 봉우리 사이 골짜기",
    "골짜기 = 시범에 한 번도 없던 행동",
  ], 7.4, 3.45, 5.3, 2.2, 15);
  T(s, "지난주(2주차) 질문의 재등장", { x: 7.4, y: 6.1, w: 5.3, h: 0.4, fontSize: 13, color: C.muted });
  s.addNotes(`[0:35–0:39] 지난주 장애물 그림을 다시. "MSE로 배우면 평균, 정면 돌진."
오른쪽 수식: MSE를 최소화하는 함수는 조건부 평균이라는 사실 하나만 짚으면 됩니다. "봉우리가 둘이면 평균은 골짜기에 떨어집니다."`);
}

// 12. 다봉성 투성이
{
  const s = content("로봇 조작은 다봉성 투성이다", S_ENEMY);
  const ex = [
    ["잡는 방법", "컵을 위에서 · 옆에서 · 손잡이로 → 모두 정답"],
    ["여러 시범자", "사람마다 스타일이 다르다"],
    ["같은 사람도", "그날그날 다르게 한다 (손떨림, 속도)"],
    ["순서", "부엌 정리: 전자레인지 먼저? 주전자 먼저?\n(BeT의 Franka Kitchen 실험)"],
  ];
  ex.forEach((e, i) => {
    const x = 0.6 + (i % 2) * 6.2, y = 1.9 + Math.floor(i / 2) * 2.1;
    card(s, x, y, 5.9, 1.85, e[0], e[1], C.bad, { bs: 15 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 6.2, w: 12.1, h: 0.75, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "시범 데이터의 행동 분포는 거의 항상 봉우리가 여러 개다", { x: 0.85, y: 6.2, w: 11.6, h: 0.75, fontSize: 17, bold: true, color: C.white, valign: "middle" });
  s.addNotes(`[0:39–0:44] 예시 4개를 빠르게. 특히 '순서의 다봉성'은 BeT 실험(Franka Kitchen)에서 다시 나오니 기억시키기.`);
}

// 13. 시간 축
{
  const s = content("시간 축의 다봉성: 갈팡질팡 (dithering)", S_ENEMY);
  T(s, "매 스텝 독립적으로 샘플링하면…", { x: 0.6, y: 1.8, w: 12, h: 0.45, fontSize: 17, bold: true });
  const seq = ["왼쪽!", "오른쪽!", "왼쪽!", "오른쪽!", "왼쪽!"];
  seq.forEach((q, i) => {
    const x = 0.6 + i * 2.45;
    box(s, [{ text: `t = ${i + 1}`, options: { fontSize: 12, color: C.muted, breakLine: true } }, { text: q, options: { fontSize: 18, bold: true, color: i % 2 ? C.A : C.L } }], x, 2.45, 2.15, 1.2, { fill: { color: C.card } });
  });
  T(s, "→ 로봇은 제자리에서 덜덜덜", { x: 0.6, y: 3.85, w: 12, h: 0.45, fontSize: 17, bold: true, color: C.bad });
  card(s, 0.6, 4.55, 5.9, 2.1, "필요한 것 1 · 봉우리 고르기", "여러 정답 중 하나를 확실히 선택한다", C.L, { fill: C.Lt, bs: 15 });
  card(s, 6.8, 4.55, 5.9, 2.1, "필요한 것 2 · 고른 봉우리 유지하기", "한 번 고른 계획을 한동안 지킨다 (시간적 일관성)\n→ 세 논문 모두를 action chunking으로 이끈 관찰", C.good, { fill: "E6F4EC", bs: 15 });
  s.addNotes(`[0:44–0:48] 오늘 새로 추가하는 개념입니다.
"봉우리를 고르는 것만으로는 부족합니다. 매 스텝 따로 뽑으면 이번엔 왼쪽 계획, 다음엔 오른쪽 계획에서 뽑혀서 로봇이 갈팡질팡합니다."
이 관찰이 세 논문 모두를 action chunking으로 이끕니다.`);
}

// 14. 데모
{
  const s = content("데모: 같은 데이터, 같은 크기의 신경망, 출력 표현만 다르게", S_ENEMY);
  s.addImage({ path: DEMO_IMG, x: 0.6, y: 1.75, w: 12.1, h: 12.1 * 736 / 2400 });
  const r = [["MSE 회귀", "100%", C.bad], ["BeT식 이산화", "0%", BET], ["Diffusion", "3%", DP]];
  T(s, "장애물 구간(a ≈ 0)에 떨어진 행동 비율", { x: 0.6, y: 5.6, w: 5, h: 0.4, fontSize: 13, color: C.muted, bold: true });
  r.forEach((x, i) => {
    T(s, x[1], { x: 0.6 + i * 3.0, y: 5.95, w: 1.5, h: 0.7, fontSize: 30, bold: true, color: x[2], valign: "middle" });
    T(s, x[0], { x: 2.0 + i * 3.0, y: 5.95, w: 1.5, h: 0.7, fontSize: 13, valign: "middle" });
  });
  T(s, "demo/multimodality_demo.py (numpy만, CPU 40초)", { x: 9.3, y: 6.3, w: 3.4, h: 0.4, fontSize: 11, color: C.muted, align: "right" });
  s.addNotes(`[0:48–0:52] 데모 그림. 회색 점이 사람 시범, 빨간 띠가 장애물.
(1) MSE: 빨간 선이 정확히 장애물 한가운데 → 100%.
(2) BeT식 이산화: 두 줄기를 모두 따라감 → 0%. 군집 경계에서 살짝 끊기는 계단 현상 = 이산화의 흔적.
(3) Diffusion: 두 줄기 모두, 약간 흩어짐 → 약 3%. "천 번 대신 50번 denoising한 작은 모델이라 조금 퍼졌다."
"같은 데이터, 같은 크기의 신경망인데 출력 표현만 바꿨습니다."
시간이 남으면 코드를 화면 공유: MSE 버전과 BeT 버전의 차이는 손실 함수 몇 줄뿐.`);
}

// 15. chunking 정의
{
  const s = content("Action Chunking: 미래 k개를 한 번에", S_ENEMY);
  T(s, "한 스텝씩", { x: 0.6, y: 1.9, w: 2, h: 0.4, fontSize: 14, bold: true, color: C.muted });
  for (let i = 0; i < 16; i++) {
    s.addShape(pres.shapes.RECTANGLE, { x: 2.6 + i * 0.6, y: 1.9, w: 0.52, h: 0.45, fill: { color: C.card }, line: { color: C.white } });
    dot(s, 2.76 + i * 0.6, 2.5, 0.2, C.ink3);
  }
  T(s, "● = 모델 추론 (16번)", { x: 2.6, y: 2.8, w: 5, h: 0.3, fontSize: 11.5, color: C.muted });
  T(s, "k = 8씩", { x: 0.6, y: 3.45, w: 2, h: 0.4, fontSize: 14, bold: true, color: C.good });
  for (let i = 0; i < 16; i++) {
    s.addShape(pres.shapes.RECTANGLE, { x: 2.6 + i * 0.6, y: 3.45, w: 0.52, h: 0.45, fill: { color: i < 8 ? C.good : "8FD3AE" }, line: { color: C.white } });
  }
  dot(s, 2.76, 4.05, 0.2, C.ink3); dot(s, 2.76 + 8 * 0.6, 4.05, 0.2, C.ink3);
  T(s, "● = 모델 추론 (2번) · 한 뭉치 = 하나의 계획", { x: 2.6, y: 4.35, w: 7, h: 0.3, fontSize: 11.5, color: C.muted });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.0, w: 12.1, h: 1.7, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "Action Chunking", { x: 0.85, y: 5.15, w: 5, h: 0.4, fontSize: 14, bold: true, color: C.good });
  T(s, "현재 관측 s를 보고 행동 1개가 아니라 미래 행동 k개(예: k = 16 ~ 100)를 한 번에 예측하는 것", { x: 0.85, y: 5.6, w: 11.6, h: 0.9, fontSize: 18, bold: true, color: C.white });
  s.addNotes(`[0:52–0:55] 칠판 오른쪽 구역에 박스로 정의를 쓰고 수업 끝까지 지우지 않습니다.
비유: "운전하면서 매 순간 핸들을 새로 결정하는 게 아니라 '앞으로 2초는 이렇게 돌리자'고 계획하는 것."`);
}

// 16. 세 효과
{
  const s = content("Chunking의 세 가지 효과 — 두 적 모두에 듣는다", S_ENEMY);
  const ef = [
    ["효과 1", "누적 오차 감소", "결정 횟수 1/k → 오차가 끼어들 기회 감소", "적 1호 대응", C.bad],
    ["효과 2", "시간적 일관성", "한 뭉치 = 하나의 계획에서 나옴 → 갈팡질팡 방지", "적 2호 대응", C.bad],
    ["효과 3", "실시간 제어", "추론 1번으로 k스텝 커버 → 무거운 모델도 버팀", "후반부 복선 (VLA)", C.A],
  ];
  ef.forEach((e, i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.9, w: 3.8, h: 3.1, rectRadius: 0.08, fill: { color: "E6F4EC" }, line: { color: "E6F4EC" } });
    T(s, e[0], { x: x + 0.25, y: 2.1, w: 3.3, h: 0.35, fontSize: 12, bold: true, color: C.good });
    T(s, e[1], { x: x + 0.25, y: 2.5, w: 3.3, h: 0.55, fontSize: 21, bold: true });
    T(s, e[2], { x: x + 0.25, y: 3.15, w: 3.3, h: 1.1, fontSize: 14, lineSpacingMultiple: 1.15 });
    T(s, e[3], { x: x + 0.25, y: 4.4, w: 3.3, h: 0.4, fontSize: 12.5, bold: true, color: e[4] });
  });
  tbl(s, [
    ["", "Diffusion Policy", "ACT", "π0 (7주차)"],
    ["chunk 길이", "16 생성 / 8 실행", "100", "50"],
    ["제어 주파수", "~10Hz", "50Hz", "50Hz"],
    ["초 단위", "~0.8초 실행", "2초", "1초"],
  ], { x: 0.6, y: 5.3, w: 12.1, colW: [2.5, 3.2, 3.2, 3.2], fontSize: 13, rowH: 0.4 });
  s.addNotes(`[0:55 직전] 세 효과를 칠판 박스에 채웁니다. "효과 1은 적 1호, 효과 2는 적 2호, 효과 3은 학기 후반 VLA의 복선."
표: "논문을 읽을 때 chunk 길이와 제어 주파수를 같이 보고 초 단위로 환산하는 습관을 들이세요."
(DP의 제어 주파수는 과제·세팅마다 다름 — 대략값)
→ 휴식.`);
}

// 17. 휴식
{
  const s = dark();
  T(s, "☕  10분 휴식", { x: 0.8, y: 1.5, w: 11, h: 1.0, fontSize: 44, bold: true, color: C.white });
  T(s, "쉬는 동안 생각해보기", { x: 0.8, y: 2.9, w: 11, h: 0.5, fontSize: 18, color: C.A, bold: true });
  T(s, "k-means로 행동을 8개 군집으로 나누면,\n모델은 '8지선다 객관식'을 푸는 셈입니다.\n이 방법의 약점은 무엇일까요?", { x: 0.8, y: 3.5, w: 11.5, h: 1.9, fontSize: 26, color: "D7DCE3" });
  vlaDots(s, 0.8, 5.9);
  s.addNotes(`[0:55–1:05] 휴식. 돌아와서 답 받기 → 정밀도 손실(군집 중심값만 냄), k 선택에 민감. "BeT는 이 약점을 '오프셋'으로 보완합니다."`);
}

// 18. 섹션: 세 개의 해법
{
  const s = dark();
  T(s, "PART ②", { x: 0.8, y: 1.3, w: 6, h: 0.5, fontSize: 18, color: C.A, bold: true, charSpacing: 3 });
  T(s, "세 개의 해법", { x: 0.8, y: 1.85, w: 11.5, h: 1.0, fontSize: 44, color: C.white, bold: true, valign: "middle" });
  const sol = [["BeT", "행동을 양자화하라", "이산화 + 오프셋", BET], ["Diffusion Policy", "행동을 생성하라", "조건부 diffusion", DP], ["ACT", "행동을 뭉치로 예측하라", "chunk 회귀 + CVAE", ACT]];
  sol.forEach((x, i) => {
    const bx = 0.8 + i * 4.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: 3.3, w: 3.7, h: 2.6, rectRadius: 0.08, fill: { color: C.ink2 }, line: { color: C.ink2 } });
    dot(s, bx + 0.3, 3.6, 0.3, x[3]);
    T(s, x[0], { x: bx + 0.3, y: 4.05, w: 3.2, h: 0.5, fontSize: 21, bold: true, color: x[3] });
    T(s, x[1], { x: bx + 0.3, y: 4.6, w: 3.2, h: 0.5, fontSize: 16, color: C.white });
    T(s, x[2], { x: bx + 0.3, y: 5.2, w: 3.2, h: 0.4, fontSize: 13, color: "B8C0CC" });
  });
  T(s, "1:05 – 2:10 · 슬라이드 19–36", { x: 0.8, y: 6.3, w: 6, h: 0.3, fontSize: 13, color: "8A93A0" });
  s.addNotes(`[1:05] 휴식 퀴즈 답을 받고 시작. "이제 두 적을 이긴 세 논문을 봅니다. 색 약속: BeT 보라, DP 주황, ACT 청록."`);
}

// 19. BeT 문제 인식
{
  const s = content("BeT: 언어 모델은 다봉성을 어떻게 다루지?", "해법 1 · BeT");
  paperTag(s, "Shafiullah et al. · NeurIPS 2022");
  code(s, "\"나는 오늘 ___\"", 0.6, 1.95, 5.2, 0.8, 20);
  const cand = [["밥을", 0.35], ["학교에", 0.3], ["잠을", 0.2], ["…", 0.15]];
  cand.forEach((c, i) => {
    const y = 3.0 + i * 0.55;
    T(s, c[0], { x: 0.6, y, w: 1.2, h: 0.45, fontSize: 15, bold: true, valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 1.9, y: y + 0.08, w: c[1] * 9, h: 0.32, fill: { color: BET }, line: { color: BET } });
    T(s, `${Math.round(c[1] * 100)}%`, { x: 2.0 + c[1] * 9, y, w: 0.8, h: 0.45, fontSize: 12, color: C.muted, valign: "middle" });
  });
  T(s, "softmax: 여러 후보에 확률을 나눠 준다 → 전혀 곤란하지 않다", { x: 0.6, y: 5.3, w: 5.8, h: 0.8, fontSize: 14, bold: true, color: BET });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: 1.95, w: 5.9, h: 4.2, rectRadius: 0.08, fill: { color: BETt }, line: { color: BETt } });
  T(s, "BeT의 발상", { x: 7.05, y: 2.15, w: 5.4, h: 0.45, fontSize: 18, bold: true, color: BET });
  bullets(s, [
    "언어 모델이 다봉성에 강한 이유는 어휘가 이산적이라서",
    "그렇다면 연속 행동도 '어휘'로 만들자",
    "Transformer는 그대로, 출력만 분류로",
    { t: "→ \"k개의 모드를 돌 하나로 복제한다\"", b: true },
  ], 7.05, 2.75, 5.4, 3.2, 15.5);
  s.addNotes(`[1:05–1:09] "BeT 저자들의 질문: 언어 모델은 다봉성을 어떻게 다루죠? '나는 오늘 ___'의 빈칸에 여러 답이 가능하지만 GPT는 전혀 곤란해하지 않습니다. 어휘가 이산적이라 softmax로 확률을 나눠주면 끝이니까요."
"그럼 행동도 이산화하자!" 논문 부제가 'Cloning k modes with one stone'.`);
}

// 20. BeT 방법
{
  const s = content("BeT의 방법: 동네를 고르고, 번지수를 찍는다", "해법 1 · BeT");
  paperTag(s, "Shafiullah et al. · NeurIPS 2022");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.85, w: 5.6, h: 4.9, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  const clusters = [[1.6, 2.8], [3.8, 2.4], [2.3, 4.6], [4.6, 4.9]];
  const rr = (i) => Math.sin(i * 12.9898) * 0.5;
  clusters.forEach((c, ci) => {
    for (let j = 0; j < 9; j++) dot(s, c[0] + rr(ci * 10 + j) * 0.9, c[1] + rr(ci * 10 + j + 5) * 0.8, 0.1, C.faint);
    s.addShape(pres.shapes.OVAL, { x: c[0] - 0.55, y: c[1] - 0.5, w: 1.1, h: 1.0, fill: { color: C.white, transparency: 100 }, line: { color: BET, width: 1.5, dashType: "dash" } });
    dot(s, c[0] - 0.09, c[1] - 0.09, 0.18, BET);
  });
  arrow(s, 3.8, 2.4, 4.25, 2.05, C.A, 2.5);
  T(s, "오프셋", { x: 4.25, y: 1.95, w: 1, h: 0.3, fontSize: 11, bold: true, color: C.A });
  T(s, "● 군집 중심 = 행동 어휘   · 시범 행동", { x: 0.8, y: 6.3, w: 5.3, h: 0.3, fontSize: 11, color: C.muted });
  T(s, "① k-means 양자화", { x: 6.6, y: 1.9, w: 6, h: 0.45, fontSize: 18, bold: true, color: BET });
  T(s, "시범의 모든 행동 벡터 → k개 대표 행동(군집 중심) = \"행동 사전\"", { x: 6.6, y: 2.4, w: 6.1, h: 0.8, fontSize: 14.5 });
  T(s, "② Transformer가 관측 히스토리를 보고", { x: 6.6, y: 3.3, w: 6, h: 0.45, fontSize: 18, bold: true, color: BET });
  bullets(s, [
    "(a) 어느 군집인가? → softmax 분류 → 봉우리 고르기",
    "(b) 중심에서 얼마나? → 잔차 오프셋 회귀 → 정밀도 복원",
  ], 6.6, 3.8, 6.1, 1.3, 14.5);
  code(s, "행동 = 중심[j] + 오프셋[j],  j ~ softmax", 6.6, 5.2, 6.1, 0.65, 14);
  T(s, "비유: 지도에서 '어느 동네'인지 고르고(이산), 동네 안 번지수를 찍는다(연속)", { x: 6.6, y: 6.0, w: 6.1, h: 0.7, fontSize: 13.5, bold: true, color: C.muted });
  s.addNotes(`[1:09–1:17] 칠판 가운데 구역에 그림: 점 구름 → 동그라미(군집) → 한 동그라미 안에서 화살표(오프셋).
"이산 선택이 봉우리 고르기를, 오프셋이 정밀도를 담당하는 분업입니다."
휴식 퀴즈와 연결: "8지선다의 약점(정밀도)을 오프셋이 메웁니다."`);
}

// 21. BeT 손실/구조
{
  const s = content("BeT의 구조와 손실", "해법 1 · BeT");
  paperTag(s, "Shafiullah et al. · NeurIPS 2022");
  const obs = ["o(t−3)", "o(t−2)", "o(t−1)", "o(t)"];
  obs.forEach((o, i) => box(s, o, 0.6 + i * 1.5, 1.95, 1.35, 0.7, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 13 }));
  T(s, "관측 히스토리 h개", { x: 0.6, y: 2.7, w: 5.8, h: 0.3, fontSize: 11.5, color: C.muted });
  box(s, "minGPT (작은 GPT)", 0.6, 3.1, 5.85, 0.8, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 16 });
  arrow(s, 2.0, 3.95, 1.9, 4.4, BET); arrow(s, 5.0, 3.95, 5.1, 4.4, C.A);
  box(s, "군집 로짓 k개\n→ softmax", 0.6, 4.45, 2.8, 1.0, { fill: { color: BETt }, color: BET, bold: true, fontSize: 14 });
  box(s, "오프셋 k × (행동 차원)\n모든 군집에 대해 한꺼번에", 3.6, 4.45, 2.85, 1.0, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 12.5 });
  T(s, "손실", { x: 6.9, y: 1.9, w: 3, h: 0.4, fontSize: 17, bold: true });
  code(s, "L = Focal(p, j*) + λ·|| r̂[j*] − (a − c[j*]) ||²", 6.9, 2.4, 5.8, 0.7, 13.5);
  bullets(s, [
    "Focal loss: 이미 잘 맞히는 쉬운 샘플의 가중치를 낮춰 드문 군집도 배우게 하는 분류 손실",
    "오프셋은 정답 군집 j*의 것만 학습 (masked)",
    "학습 땐 정답 군집 오프셋, 실행 땐 뽑힌 군집 오프셋 사용",
    "히스토리를 보는 이유: 지금 어느 모드(순서)를 수행 중인지 알아야 함",
  ], 6.9, 3.3, 5.8, 3.5, 13.5);
  s.addNotes(`[1:17–1:20] (시간 없으면 한 줄: "분류 손실 + 정답 군집의 오프셋 회귀")
Focal loss = 쉬운 문제(자주 나오는 군집)의 가중치를 낮춰, 드문 군집도 잘 배우게 하는 분류 손실.
모델은 모든 군집에 대해 오프셋을 한꺼번에 출력하고, 학습 때는 정답 군집 것만, 실행 때는 뽑힌 군집 것만 씁니다.`);
}

// 22. BeT 실험/의의
{
  const s = content("BeT의 실험, 의의, 한계", "해법 1 · BeT");
  paperTag(s, "Shafiullah et al. · NeurIPS 2022");
  const env = [["CARLA 자율주행", "건물을 왼쪽/오른쪽으로 돌아가는 두 경로"], ["블록 밀기", "어느 블록을 어느 목표 칸으로 먼저"], ["Franka Kitchen", "7개 하위작업을 어떤 순서로"]];
  env.forEach((e, i) => card(s, 0.6 + i * 4.1, 1.85, 3.8, 1.55, e[0], e[1], BET, { fill: BETt, hs: 16, bs: 13 }));
  T(s, "결과: 기존 BC(MLP)는 모드가 섞여 실패, BeT는 시범에 있던 여러 모드(경로 · 순서)를 골고루 재현 (시뮬레이션)", { x: 0.6, y: 3.6, w: 12.1, h: 0.7, fontSize: 14.5, bold: true });
  card(s, 0.6, 4.45, 3.8, 2.35, "의의", "\"행동 이산화\" 계보의 시조 → 6주차 RT-1 · RT-2 · OpenVLA의 256-bin (더 단순한 균등 구간)", C.good, { hs: 16, bs: 13 });
  card(s, 4.7, 4.45, 3.8, 2.35, "한계", "k에 민감 · 정밀 조작에서 이산화 손실 · 원 논문은 목표/언어 조건 없음", C.bad, { hs: 16, bs: 13 });
  card(s, 8.8, 4.45, 3.9, 2.35, "후속", "C-BeT (목표 조건, 실제 로봇)\nVQ-BeT (2024, 잔차 벡터 양자화로 k-means 대체)", C.muted, { hs: 16, bs: 13 });
  s.addNotes(`[1:20–1:25] 실험 세 환경은 각각 다른 종류의 다봉성: 경로(CARLA), 대상·목표(블록), 순서(부엌).
의의를 강조: "6주차 RT-1/RT-2/OpenVLA가 행동을 256칸으로 나누는 것은 이 정신의 직계 후손입니다."
체크 질문: "BeT에서 softmax는 무엇을 담당하나?" → 봉우리 고르기.`);
}

// 23. DP 한 문장
{
  const s = content("Diffusion Policy: 관측을 조건으로 행동 시퀀스를 생성한다", "해법 2 · Diffusion Policy");
  paperTag(s, "Chi et al. · RSS 2023 / IJRR 2024");
  T(s, "2주차", { x: 0.6, y: 1.9, w: 2, h: 0.4, fontSize: 14, bold: true, color: C.muted });
  box(s, "노이즈 이미지 x_T", 0.6, 2.35, 3.2, 0.85, { fill: { color: C.card }, bold: true, fontSize: 14 });
  arrow(s, 3.85, 2.78, 5.55, 2.78); T(s, "denoise × K", { x: 3.9, y: 2.35, w: 1.6, h: 0.3, fontSize: 11, color: C.muted, align: "center" });
  box(s, "이미지 x_0", 5.6, 2.35, 2.4, 0.85, { fill: { color: C.card }, bold: true, fontSize: 14 });
  T(s, "조건 c = 텍스트", { x: 8.3, y: 2.35, w: 4.4, h: 0.85, fontSize: 15, valign: "middle" });
  T(s, "오늘", { x: 0.6, y: 3.45, w: 2, h: 0.4, fontSize: 14, bold: true, color: DP });
  box(s, "노이즈 행동열 A_K", 0.6, 3.9, 3.2, 0.85, { fill: { color: DPt }, color: DP, bold: true, fontSize: 14 });
  arrow(s, 3.85, 4.33, 5.55, 4.33, DP); T(s, "denoise × K", { x: 3.9, y: 3.9, w: 1.6, h: 0.3, fontSize: 11, color: DP, align: "center" });
  box(s, "행동 시퀀스 A_0", 5.6, 3.9, 2.4, 0.85, { fill: { color: DP }, color: C.white, bold: true, fontSize: 14 });
  T(s, "조건 c = 최근 관측\n(이미지 2장 + 로봇 상태)", { x: 8.3, y: 3.9, w: 4.4, h: 0.85, fontSize: 15, valign: "middle" });
  code(s, "L = || ε − ε_θ( A_k, k, 관측 ) ||²        ← 2주차 손실과 글자 하나 다르지 않다", 0.6, 5.2, 12.1, 0.75, 15);
  T(s, "조건부 생성모델 = 로봇 정책 (2주차 칠판의 등식이 논문이 된 것)", { x: 0.6, y: 6.15, w: 12.1, h: 0.45, fontSize: 16, bold: true, color: DP });
  s.addNotes(`[1:25–1:28] "지난주 예고를 회수하겠습니다. 관측을 조건으로, 미래 행동 시퀀스를 diffusion으로 생성한다. 끝입니다."
칠판 가운데에 2주차 그림을 다시 그리고 글자만 바꿉니다: x → 행동 시퀀스, c → 관측.
"손실 함수가 2주차와 글자 하나 다르지 않습니다." 3주간의 투자를 회수하는 순간.`);
}

// 24. receding horizon
{
  const s = content("작동 흐름: 계획은 길게, 실행은 앞부분만 (receding horizon)", "해법 2 · Diffusion Policy");
  paperTag(s, "Chi et al. · RSS 2023 / IJRR 2024");
  const x0 = 2.4, cw = 0.58;
  T(s, "관측 T_o = 2", { x: 0.6, y: 2.0, w: 1.7, h: 0.5, fontSize: 13, bold: true, color: C.V, valign: "middle" });
  for (let i = 0; i < 2; i++) s.addShape(pres.shapes.RECTANGLE, { x: x0 + i * cw, y: 2.0, w: cw - 0.06, h: 0.5, fill: { color: C.V }, line: { color: C.white } });
  T(s, "생성 T_p = 16", { x: 0.6, y: 2.8, w: 1.7, h: 0.5, fontSize: 13, bold: true, color: DP, valign: "middle" });
  for (let i = 0; i < 16; i++) s.addShape(pres.shapes.RECTANGLE, { x: x0 + i * cw, y: 2.8, w: cw - 0.06, h: 0.5, fill: { color: i >= 1 && i < 9 ? DP : DPt }, line: { color: C.white } });
  T(s, "실행 T_a = 8", { x: 0.6, y: 3.6, w: 1.7, h: 0.5, fontSize: 13, bold: true, color: C.good, valign: "middle" });
  for (let i = 1; i < 9; i++) s.addShape(pres.shapes.RECTANGLE, { x: x0 + i * cw, y: 3.6, w: cw - 0.06, h: 0.5, fill: { color: C.good }, line: { color: C.white } });
  arrow(s, x0 + 9 * cw, 3.85, x0 + 10.6 * cw, 3.85, C.good, 2.5);
  T(s, "새 관측 → 다시 생성", { x: x0 + 10.7 * cw, y: 3.6, w: 3.0, h: 0.5, fontSize: 13, bold: true, color: C.good, valign: "middle" });
  card(s, 0.6, 4.5, 5.9, 2.2, "왜 16개 만들고 8개만 쓰나", [
    "길게 계획 → 일관성 (한 모드를 유지)",
    "짧게 실행 → 반응성 (새 관측을 빨리 반영)",
  ], DP, { fill: DPt, bs: 14.5 });
  card(s, 6.8, 4.5, 5.9, 2.2, "비유", "내비게이션이 전체 경로를 그려주지만, 몇 초마다 다시 계산한다", C.muted, { bs: 15 });
  s.addNotes(`[1:28–1:33] 칠판에: 관측 2스텝 → 행동 16스텝 생성 → 앞 8스텝만 실행 → 새 관측 → 다시 생성.
"계획은 길게(일관성), 실행은 짧게(반응성)." 칠판 오른쪽 chunking 박스의 효과 1·2와 연결.`);
}

// 25. 왜 잘 되나
{
  const s = content("Diffusion Policy는 왜 잘 되나", "해법 2 · Diffusion Policy");
  paperTag(s, "Chi et al. · RSS 2023 / IJRR 2024");
  const why = [
    ["다봉성", "분포 자체를 배운다 → 데모 그림 (3)"],
    ["시간적 일관성", "16스텝을 한 덩어리로 생성 → chunking 내장"],
    ["정밀도", "이산화가 없다 → 연속 행동 그대로"],
    ["학습 안정성", "에너지 대신 score를 배운다 (다음 슬라이드)"],
  ];
  why.forEach((w, i) => {
    const x = 0.6 + (i % 2) * 6.2, y = 1.9 + Math.floor(i / 2) * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.9, h: 2.05, rectRadius: 0.08, fill: { color: DPt }, line: { color: DPt } });
    numCircle(s, i + 1, x + 0.3, y + 0.35, 0.6, DP, 18);
    T(s, w[0], { x: x + 1.1, y: y + 0.35, w: 4.6, h: 0.6, fontSize: 20, bold: true, valign: "middle" });
    T(s, w[1], { x: x + 1.1, y: y + 1.05, w: 4.6, h: 0.8, fontSize: 14.5 });
  });
  s.addNotes(`[1:33–1:36] 네 가지를 빠르게. 1~3은 앞에서 배운 것의 정리, 4는 3주차 연결(다음 슬라이드).`);
}

// 26. IBC vs DP
{
  const s = content("3주차 연결: 에너지 대신 score를 배운다", "해법 2 · Diffusion Policy");
  paperTag(s, "Chi et al. · RSS 2023 / IJRR 2024");
  card(s, 0.6, 1.9, 5.9, 3.4, "이전: IBC (암묵적 정책, 에너지 모델)", [
    "p(a|o) ∝ exp(−E(a, o))",
    "정규화 상수 Z(o) = ∫exp(−E) da 계산이 어렵다",
    "음성 샘플로 Z를 근사 → 학습 불안정",
  ], C.bad, { bs: 14.5 });
  card(s, 6.8, 1.9, 5.9, 3.4, "Diffusion Policy", [
    "∇_a log p(a|o) = −∇_a E(a, o)",
    "미분하면 Z가 사라진다 (상수이므로)",
    "노이즈 예측 = score 학습 → 음성 샘플 불필요, 안정적",
  ], C.good, { fill: "E6F4EC", bs: 14.5 });
  arrow(s, 6.52, 3.6, 6.78, 3.6, C.A, 3);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.6, w: 12.1, h: 1.1, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "3주차: \"diffusion 학습 = score 학습\"  →  여기서 실용적 이득(안정성)으로 나타난다", { x: 0.85, y: 5.6, w: 11.6, h: 1.1, fontSize: 16, bold: true, color: C.white, valign: "middle" });
  s.addNotes(`[1:36–1:38] (보너스, 시간 없으면 한 줄)
"이전 방법 IBC는 에너지 함수를 배웠는데, 확률로 바꾸려면 정규화 상수를 계산해야 해서 음성 샘플이 필요했고 학습이 불안정했습니다. DP는 에너지의 기울기, 즉 score를 배웁니다. 미분하면 정규화 상수가 사라지죠. 3주차에 배운 'diffusion 학습 = score 학습'이 여기서 실익이 됩니다."`);
}

// 27. DP 결과
{
  const s = content("Diffusion Policy의 결과", "해법 2 · Diffusion Policy");
  paperTag(s, "Chi et al. · RSS 2023 / IJRR 2024");
  stat(s, 0.6, 1.8, 3.9, "+46.9%", "기존 최고 방법 대비 평균 향상", DP);
  stat(s, 4.7, 1.8, 3.9, "15개", "과제 (4개 벤치마크)", DP);
  stat(s, 8.8, 1.8, 3.9, "3개", "비교 대상: LSTM-GMM · IBC · BET", C.muted);
  card(s, 0.6, 3.9, 5.9, 2.85, "시뮬레이션 벤치마크", [
    "Robomimic: Lift · Can · Square · Transport · ToolHang",
    "Push-T: T자 블록을 목표에 밀어 넣기",
    "Multimodal Block Pushing · Franka Kitchen",
  ], DP, { bs: 13.5 });
  card(s, 6.8, 3.9, 5.9, 2.85, "실제 로봇", [
    "Push-T · 머그컵 뒤집기 · 소스 붓기 · 펴 바르기",
    "사람이 방해해도 재계획으로 복구",
    "확장판(IJRR): 양팔 과제 추가",
  ], DP, { bs: 13.5 });
  T(s, "▶ 영상: diffusion-policy.cs.columbia.edu", { x: 0.6, y: 6.85, w: 8, h: 0.35, fontSize: 12, color: C.muted });
  s.addNotes(`[1:38–1:42] ★ 영상: Push-T에서 사람이 블록을 치워도 다시 계획해 복구하는 장면.
"4개 벤치마크 15개 과제에서 기존 최고 대비 평균 46.9% 향상. 비교 대상에 BET가 있다는 것에 주목 — 오늘 앞의 논문을 이겼습니다."
"receding horizon의 힘: 방해를 받아도 몇 스텝 뒤 새로 계획합니다."`);
}

// 28. 구조 두 가지
{
  const s = content("구조 두 가지와 실무 팁", "해법 2 · Diffusion Policy");
  paperTag(s, "Chi et al. · RSS 2023 / IJRR 2024");
  tbl(s, [
    ["", "CNN 기반 (1D 시간축 U-Net)", "Transformer 기반"],
    ["조건 주입", "FiLM: 관측으로 채널 scale/bias 조절", "cross-attention: 행동 토큰 → 관측"],
    ["추천", "기본값, 하이퍼파라미터에 덜 민감", "행동이 급변하는 과제, 튜닝 필요"],
  ], { x: 0.6, y: 1.9, w: 12.1, colW: [2.1, 5.0, 5.0], fontSize: 14, rowH: 0.6 });
  card(s, 0.6, 4.0, 5.9, 2.75, "비전 인코더", [
    "ResNet-18을 사전학습 없이 정책과 함께 학습",
    "전역 풀링 → spatial softmax (\"어디\" 정보 보존)",
    "BatchNorm → GroupNorm",
  ], C.V, { bs: 13.5 });
  card(s, 6.8, 4.0, 5.9, 2.75, "위치 제어 > 속도 제어", [
    "\"이 각도로 가라\"(절대) vs \"이만큼 더\"(델타)",
    "델타는 오차가 적분되어 누적 오차에 취약",
    "위치 제어가 행동 시퀀스 예측과 궁합이 좋음",
  ], DP, { bs: 13.5 });
  s.addNotes(`[1:42–1:46] (시간 없으면 스킵 가능)
CNN 기반이 기본값, 행동이 급변하는 과제는 Transformer 기반.
실무 팁 두 개만: (1) spatial softmax — 특징맵에서 '어디'를 좌표로 뽑는 풀링, 로봇은 '어디'가 생명. (2) 위치 제어가 더 좋음 — 누적 오차(적 1호)와 연결.`);
}

// 29. 약점
{
  const s = content("약점: 2주차에 심어 둔 걱정이 돌아왔다", "해법 2 · Diffusion Policy");
  paperTag(s, "Chi et al. · RSS 2023 / IJRR 2024");
  stat(s, 0.6, 1.8, 3.0, "K = 100", "학습 시 denoising 스텝", C.muted);
  stat(s, 3.9, 1.8, 3.0, "10스텝", "실행 시 DDIM으로 축소", DP);
  stat(s, 7.2, 1.8, 3.2, "~0.1초", "chunk 하나 생성 (RTX 3080)", C.bad);
  T(s, "\"로봇은 0.02초마다 행동을 내야 하는데?\" → chunking으로 상쇄하지만 근본적으로 반복적", { x: 0.6, y: 3.85, w: 12.1, h: 0.5, fontSize: 16, bold: true });
  card(s, 0.6, 4.55, 5.9, 2.2, "이후의 해답", [
    "3주차 flow matching → π0 (약 10스텝)",
    "12주차 효율화 · 비동기 실행 (RTC)",
  ], C.good, { bs: 14 });
  card(s, 6.8, 4.55, 5.9, 2.2, "유산", [
    "소규모 모방학습의 사실상 기본 베이스라인",
    "Octo의 diffusion head · GR00T의 DiT head · UMI",
  ], DP, { bs: 14 });
  s.addNotes(`[1:46–1:50] "2주차에 '생성은 느리다'는 걱정을 심어 뒀죠. 그게 이 논문의 한계로 등장합니다."
학습 100스텝 → 실행 DDIM 10스텝, 약 0.1초. chunking으로 어느 정도 상쇄.
"지난주의 걱정이 실제 논문의 한계로 등장하는 것" — 수업 서사 연결.
체크: "DP와 2주차 이미지 diffusion의 차이는?" → x가 행동 시퀀스, c가 관측. 수학은 같다.`);
}

// 30. ALOHA
{
  const s = content("ACT: 하드웨어가 반이다 — ALOHA", "해법 3 · ACT");
  paperTag(s, "Zhao, Kumar, Levine, Finn · RSS 2023");
  box(s, "리더 팔 × 2\n(사람이 잡고 움직임)", 0.6, 2.0, 2.8, 1.1, { fill: { color: C.Lt }, color: C.L, bold: true, fontSize: 14 });
  arrow(s, 3.45, 2.55, 4.35, 2.55, C.faint, 3);
  T(s, "관절 1:1", { x: 3.45, y: 2.1, w: 0.9, h: 0.3, fontSize: 11, color: C.muted, align: "center" });
  box(s, "팔로워 팔 × 2\n(ViperX 6자유도)", 4.4, 2.0, 2.8, 1.1, { fill: { color: ACT }, color: C.white, bold: true, fontSize: 14 });
  stat(s, 0.6, 3.35, 2.1, "4대", "카메라 (위·앞·양 손목)", ACT);
  stat(s, 2.9, 3.35, 2.1, "50Hz", "제어 주파수", ACT);
  stat(s, 5.1, 3.35, 2.3, "<$20k", "총 비용 (기성품+3D프린팅)", ACT);
  T(s, "실제 과제 6개", { x: 7.8, y: 1.9, w: 4.9, h: 0.4, fontSize: 17, bold: true });
  const tasks = ["지퍼백 열기", "건전지 끼우기", "투명 소스컵 뚜껑 열기", "벨크로(케이블 타이) 끼우기", "테이프 뜯어 붙이기", "신발 신기기"];
  tasks.forEach((t, i) => box(s, t, 7.8 + (i % 2) * 2.5, 2.45 + Math.floor(i / 2) * 0.85, 2.35, 0.7, { fill: { color: ACTt }, fontSize: 13, bold: true }));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.55, w: 12.1, h: 1.15, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "저가 로봇은 부정확하다 → 학습으로 보상할 수 있을까? 정밀 과제일수록 누적 오차가 치명적 → chunking이 주인공이 된 이유", { x: 0.85, y: 5.55, w: 11.6, h: 1.15, fontSize: 14.5, bold: true, valign: "middle" });
  s.addNotes(`[1:50–1:53] ★ 영상: tonyzhaozh.github.io/aloha — 건전지 끼우기, 컵 뚜껑 열기.
"저자들은 알고리즘 전에 장비부터 만들었습니다. 사람이 작은 리더 팔 두 개를 잡고 움직이면 큰 팔 두 개가 그대로 따라합니다."
하드웨어(ALOHA) + 알고리즘(ACT)이 세트인 논문. 후속 Mobile ALOHA, ALOHA 2로 이어짐.`);
}

// 31. ACT 개요
{
  const s = content("ACT = Action Chunking with Transformers", "해법 3 · ACT");
  paperTag(s, "Zhao, Kumar, Levine, Finn · RSS 2023");
  box(s, "이미지 4장", 0.6, 2.1, 2.2, 0.8, { fill: { color: C.Vt }, color: C.V, bold: true, fontSize: 14 });
  box(s, "관절 14개", 0.6, 3.1, 2.2, 0.8, { fill: { color: C.card }, color: C.muted, bold: true, fontSize: 14 });
  arrow(s, 2.85, 2.5, 3.65, 2.9); arrow(s, 2.85, 3.5, 3.65, 3.1);
  box(s, "Transformer\n(한 번의 forward)", 3.7, 2.2, 3.0, 1.6, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 16 });
  arrow(s, 6.75, 3.0, 7.45, 3.0, ACT, 2.5);
  for (let i = 0; i < 10; i++) s.addShape(pres.shapes.RECTANGLE, { x: 7.5 + i * 0.52, y: 2.7, w: 0.46, h: 0.6, fill: { color: ACT }, line: { color: C.white } });
  T(s, "미래 행동 k = 100개 (50Hz에서 2초)", { x: 7.5, y: 3.4, w: 5.2, h: 0.35, fontSize: 13, bold: true, color: ACT });
  card(s, 0.6, 4.4, 3.8, 2.35, "회귀", "diffusion처럼 반복하지 않는다 → 빠름. 손실은 L1 (정밀 과제에서 L2보다 좋음)", ACT, { fill: ACTt, bs: 13.5 });
  card(s, 4.7, 4.4, 3.8, 2.35, "장치 1 · CVAE", "사람 시범의 들쭉날쭉함(스타일)을 잠재변수 z로 흡수", ACT, { fill: ACTt, bs: 13.5 });
  card(s, 8.8, 4.4, 3.9, 2.35, "장치 2 · 시간 앙상블", "Temporal ensembling: 겹치는 예측을 가중평균해 부드럽게 실행", ACT, { fill: ACTt, bs: 13.5 });
  s.addNotes(`[1:53–1:55] "이름 그대로 chunking이 주인공입니다. 100개를 한 번에 회귀로 뽑습니다."
여기에 장치 두 개가 붙습니다: CVAE와 temporal ensembling. 하나씩 봅니다.`);
}

// 32. VAE 보충
{
  const s = content("5분 보충: VAE란?", "해법 3 · ACT");
  box(s, "데이터 x", 0.6, 2.4, 2.0, 1.0, { fill: { color: C.card }, bold: true, fontSize: 15 });
  arrow(s, 2.65, 2.9, 3.35, 2.9);
  box(s, "인코더", 3.4, 2.4, 1.8, 1.0, { fill: { color: ACT }, color: C.white, bold: true, fontSize: 15 });
  arrow(s, 5.25, 2.9, 5.95, 2.9);
  box(s, "z\n(짧은 요약)", 6.0, 2.3, 1.6, 1.2, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 15 });
  arrow(s, 7.65, 2.9, 8.35, 2.9);
  box(s, "디코더", 8.4, 2.4, 1.8, 1.0, { fill: { color: ACT }, color: C.white, bold: true, fontSize: 15 });
  arrow(s, 10.25, 2.9, 10.95, 2.9);
  box(s, "복원 x̂", 11.0, 2.4, 1.7, 1.0, { fill: { color: C.card }, bold: true, fontSize: 15 });
  bullets(s, [
    "데이터를 짧은 요약 벡터 z로 압축했다가, z에서 다시 복원하도록 학습하는 오토인코더의 확률 버전",
    "z가 표준정규분포 N(0, I)에서 너무 멀어지지 않게 KL 벌점을 준다",
    "CVAE(조건부 VAE): 인코더·디코더에 조건(여기선 관측)을 함께 넣는다",
  ], 0.6, 4.0, 12.1, 1.8, 15);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.95, w: 12.1, h: 0.8, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "ACT에서는 z = \"시범의 스타일 요약\"이라고만 이해하면 충분", { x: 0.85, y: 5.95, w: 11.6, h: 0.8, fontSize: 17, bold: true, color: C.white, valign: "middle" });
  s.addNotes(`[1:55–1:57] VAE를 모르는 학생을 위한 보충. "데이터를 압축 요약하는 벡터 z를 두고, z로부터 데이터를 복원하게 학습하는 오토인코더의 확률 버전. 여기서는 z가 '시범 스타일 요약'이라고만 이해하면 충분합니다."`);
}

// 33. CVAE train/infer
{
  const s = content("ACT의 CVAE: 훈련 때와 실행 때가 다르다", "해법 3 · ACT");
  paperTag(s, "Zhao, Kumar, Levine, Finn · RSS 2023");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.85, w: 12.1, h: 2.3, rectRadius: 0.08, fill: { color: C.card }, line: { color: C.card } });
  T(s, "훈련", { x: 0.85, y: 1.95, w: 2, h: 0.4, fontSize: 15, bold: true, color: ACT });
  box(s, "실제 시범 행동 100개\n+ 관절", 0.85, 2.45, 2.6, 1.3, { fill: { color: C.white }, bold: true, fontSize: 13 });
  arrow(s, 3.5, 3.1, 3.95, 3.1);
  box(s, "인코더", 4.0, 2.65, 1.5, 0.9, { fill: { color: ACT }, color: C.white, bold: true, fontSize: 14 });
  arrow(s, 5.55, 3.1, 6.0, 3.1);
  box(s, "z (32차원)\n\"빠르게/조심스럽게\"", 6.05, 2.55, 2.3, 1.1, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 12.5 });
  arrow(s, 8.4, 3.1, 8.85, 3.1);
  box(s, "디코더 (+이미지·관절)\n→ 행동 100개 복원", 8.9, 2.55, 3.55, 1.1, { fill: { color: ACT }, color: C.white, bold: true, fontSize: 13 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.35, w: 12.1, h: 1.55, rectRadius: 0.08, fill: { color: ACTt }, line: { color: ACTt } });
  T(s, "실행", { x: 0.85, y: 4.45, w: 2, h: 0.4, fontSize: 15, bold: true, color: ACT });
  T(s, "인코더 없음", { x: 0.85, y: 4.95, w: 2.6, h: 0.6, fontSize: 14, color: C.muted, valign: "middle" });
  box(s, "z = 0 (평균 스타일)", 6.05, 4.8, 2.3, 0.8, { fill: { color: C.At }, color: C.A, bold: true, fontSize: 13 });
  arrow(s, 8.4, 5.2, 8.85, 5.2);
  box(s, "디코더 → 행동 100개", 8.9, 4.8, 3.55, 0.8, { fill: { color: ACT }, color: C.white, bold: true, fontSize: 13 });
  code(s, "L = L1(복원) + β · KL( q(z | 행동, 관절) || N(0, I) ),   β = 10", 0.6, 6.1, 12.1, 0.65, 14);
  s.addNotes(`[1:57–2:00] "같은 장면에서 시범이 제각각인 것(다봉성 + 손떨림)을 z가 흡수합니다. z가 스타일을 떠맡으니 디코더는 평균 내기를 할 필요가 없습니다."
실행 때는 인코더를 버리고 z=0 → 가장 표준적인 스타일로 움직임.
(질문 대비) 인코더에는 속도를 위해 이미지를 넣지 않고 관절과 행동 시퀀스만 넣습니다.`);
}

// 34. temporal ensembling
{
  const s = content("Temporal Ensembling: 겹치는 예측을 평균해 부드럽게", "해법 3 · ACT");
  paperTag(s, "Zhao, Kumar, Levine, Finn · RSS 2023");
  const cw = 0.55, x0 = 1.9;
  for (let r = 0; r < 4; r++) {
    T(s, r === 3 ? "t에서 예측" : `t−${3 - r}에서 예측`, { x: 0.6, y: 1.95 + r * 0.55, w: 1.3, h: 0.45, fontSize: 11, color: C.muted, valign: "middle" });
    for (let i = 0; i < 8; i++) {
      const col = r + i;
      s.addShape(pres.shapes.RECTANGLE, { x: x0 + col * cw, y: 1.95 + r * 0.55, w: cw - 0.05, h: 0.45, fill: { color: col === 3 ? ACT : ACTt }, line: { color: C.white } });
    }
  }
  s.addShape(pres.shapes.RECTANGLE, { x: x0 + 3 * cw - 0.05, y: 1.85, w: cw + 0.05, h: 2.35, fill: { color: C.white, transparency: 100 }, line: { color: C.A, width: 2.5 } });
  T(s, "현재 시점 t", { x: x0 + 3 * cw - 0.6, y: 4.25, w: 1.7, h: 0.3, fontSize: 12, bold: true, color: C.A, align: "center" });
  code(s, "a_t = Σ w_i · Â_t(i) / Σ w_i\nw_i = exp(−m · i),  i=0이 가장 오래된 예측", 8.3, 1.95, 4.4, 1.2, 13);
  T(s, "m이 작을수록 새 관측을 빨리 반영 · 추가 학습 비용 없음", { x: 8.3, y: 3.3, w: 4.4, h: 0.8, fontSize: 13, color: C.muted });
  card(s, 0.6, 4.8, 5.9, 1.95, "Diffusion Policy", "8개 실행 후 새로 → 끊어서 실행", DP, { fill: DPt, bs: 14.5 });
  card(s, 6.8, 4.8, 5.9, 1.95, "ACT", "매 스텝 새 chunk 예측 + 겹친 예측 평균 → 부드럽게 실행", ACT, { fill: ACTt, bs: 14.5 });
  s.addNotes(`[2:00–2:04] "매 스텝 새 chunk를 예측하면, 현재 시점에 대해 과거 여러 chunk가 낸 예측이 겹칩니다. 그걸 가중평균해서 실행합니다."
주황 테두리 열이 현재 시점. "계획이 급변하지 않게 하는 완충 장치."
DP와 대비: DP는 끊어서, ACT는 부드럽게.`);
}

// 35. ACT 결과
{
  const s = content("ACT의 결과와 ablation", "해법 3 · ACT");
  paperTag(s, "Zhao, Kumar, Levine, Finn · RSS 2023");
  stat(s, 0.6, 1.8, 3.0, "50개", "과제당 시범 (약 10분 분량)", ACT);
  stat(s, 3.9, 1.8, 3.0, "80–90%", "정밀 양팔 과제 성공률 (대부분)", ACT);
  stat(s, 7.2, 1.8, 2.8, "~0%", "비교 방법 대부분 (BeT · RT-1 · VINN 등)", C.bad);
  T(s, "가장 어려운 벨크로 끼우기는 20%대", { x: 10.2, y: 2.3, w: 2.5, h: 0.9, fontSize: 12.5, color: C.muted });
  const ab = [
    ["chunk 크기", "k=1에서 매우 낮고, k를 키울수록 크게 향상. 다른 방법에 붙여도 좋아짐 → 일반 원리"],
    ["CVAE", "스크립트 데이터(일관됨)에선 차이 없음, 사람 데이터에선 빼면 급락 → 들쭉날쭉함을 흡수"],
    ["제어 주파수", "사람도 5Hz 조종은 50Hz보다 훨씬 어려움 → 고주파가 정밀 조작에 중요"],
  ];
  ab.forEach((a, i) => {
    const y = 4.0 + i * 0.95;
    box(s, a[0], 0.6, y, 2.4, 0.8, { fill: { color: ACT }, color: C.white, bold: true, fontSize: 14 });
    T(s, a[1], { x: 3.2, y, w: 9.5, h: 0.8, fontSize: 14, valign: "middle" });
  });
  s.addNotes(`[2:04–2:07] "충격 포인트: 데이터 50개 + 회귀로도, chunking과 스타일 처리만 잘하면 정밀 조작이 된다."
Ablation 세 개를 짚기 — 특히 CVAE: 스크립트 데이터에서는 차이가 없고 사람 데이터에서만 차이가 난다 → CVAE가 정말 사람의 들쭉날쭉함을 흡수한다는 증거.
(과제별 정확한 수치는 02_논문해설.md 참고, 질문 받으면 원문 Table 확인 약속)
체크: "ACT가 실행할 때 z에 무엇을 넣나?" → 0.`);
}

// 36. ACT 유산
{
  const s = content("ACT의 한계와 유산", "해법 3 · ACT");
  paperTag(s, "Zhao, Kumar, Levine, Finn · RSS 2023");
  card(s, 0.6, 1.9, 5.9, 2.5, "한계", [
    "과제별 학습, 언어 조건 없음",
    "실행 시 z = 0 → 사실상 한 가지 스타일",
    "좌/우처럼 완전히 다른 모드엔 여전히 위험",
  ], C.bad, { bs: 14 });
  card(s, 6.8, 1.9, 5.9, 2.5, "후속", [
    "Mobile ALOHA (2024): 바퀴 달린 전신 조작",
    "ALOHA 2 (2024): Gemini Robotics의 플랫폼",
    "ALOHA Unleashed (2024): 대규모 데이터 + diffusion",
  ], ACT, { fill: ACTt, bs: 14 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.7, w: 12.1, h: 2.05, rectRadius: 0.08, fill: { color: C.ink }, line: { color: C.ink } });
  T(s, "가장 큰 유산", { x: 0.85, y: 4.85, w: 5, h: 0.4, fontSize: 14, bold: true, color: ACT });
  T(s, "action chunking → π0의 50스텝 @ 50Hz, 거의 모든 VLA\nLeRobot의 기본 정책 → 저가 SO-100/101 팔 입문의 표준", { x: 0.85, y: 5.3, w: 11.6, h: 1.3, fontSize: 17, bold: true, color: C.white });
  s.addNotes(`[2:07–2:10] 한계: 실행 시 z=0이라 진짜 다봉 샘플링은 아님. 그래서 일부 후속은 diffusion head로 교체.
유산: chunking이 모든 VLA로. LeRobot에서 ACT가 기본 정책이라 지금도 입문 표준.`);
}

// 37. 비교표
{
  const s = content("세 해법 비교", "정리");
  tbl(s, [
    ["", "BeT", "Diffusion Policy", "ACT"],
    ["다봉성 해법", "이산화(분류) + 오프셋", "분포 자체를 생성", "CVAE로 스타일 흡수"],
    ["출력", "행동 토큰(+잔차)", "행동 시퀀스 생성 (16)", "행동 chunk 회귀 (100)"],
    ["실행 방식", "매 스텝", "8스텝 실행 후 재계획", "매 스텝 + temporal ensembling"],
    ["속도", "빠름", "느림 (반복 denoising)", "빠름 (1회 forward)"],
    ["손실", "focal + 오프셋 MSE", "ε 예측 MSE", "L1 + β·KL"],
    ["유산", "RT류의 행동 이산화", "Octo · GR00T · π0의 생성 head", "모든 VLA의 chunking"],
  ], { x: 0.6, y: 1.85, w: 12.1, colW: [2.0, 3.2, 3.5, 3.4], fontSize: 14, rowH: 0.66 });
  s.addNotes(`[2:10–2:14] 칠판 가운데를 정리하며 3열 표. 특히 마지막 줄 '유산'을 강조: 6주차 이후 논문에서 이 세 가지가 부품으로 다시 나옵니다.`);
}

// 38. 공유한 것 + 6주차
{
  const s = content("세 논문이 공유한 것, 그리고 6주차로", "정리");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.85, w: 12.1, h: 1.3, rectRadius: 0.08, fill: { color: "E6F4EC" }, line: { color: "E6F4EC" } });
  T(s, "다봉성을 푸는 방법은 달랐지만, 셋 다 여러 스텝을 한 번에 예측했다 (action chunking)", { x: 0.85, y: 1.85, w: 11.6, h: 1.3, fontSize: 19, bold: true, color: C.good, valign: "middle" });
  box(s, "작은 모델\n한 과제\n수십~수백 시범", 0.6, 3.6, 3.4, 2.0, { fill: { color: C.card }, bold: true, fontSize: 16 });
  arrow(s, 4.1, 4.6, 5.3, 4.6, C.A, 3);
  T(s, "+ 1주차 VLM", { x: 4.0, y: 4.0, w: 1.5, h: 0.4, fontSize: 12, bold: true, color: C.L, align: "center" });
  box(s, "VLA\n한 모델 · 수백 과제\n언어 명령", 5.4, 3.6, 3.4, 2.0, { fill: { color: C.ink }, color: C.white, bold: true, fontSize: 16 });
  T(s, "6주차: RT-1 → RT-2 → Octo → OpenVLA", { x: 9.1, y: 3.6, w: 3.6, h: 2.0, fontSize: 16, bold: true, valign: "middle" });
  T(s, "논문 읽기 습관: 행동 표현(이산/생성/회귀)? chunk 길이? 제어 주파수? 초 단위로 환산!", { x: 0.6, y: 6.0, w: 12.1, h: 0.6, fontSize: 15, bold: true, color: C.muted });
  s.addNotes(`[2:14–2:17] 오른쪽 chunking 박스를 가리키며: "오늘의 가장 중요한 유산입니다."
"세 논문은 '작은 모델 + 한 과제 + 수십~수백 시범'의 세계였습니다. 여기에 1주차의 VLM을 얹어 '한 모델이 수백 과제를, 언어 명령으로' 하게 만들면? 그게 VLA이고 6주차에서 봅니다."
"칠판을 사진 찍어 가세요. 6주차 이후 논문을 읽을 때 체크리스트가 됩니다."`);
}

// 39. O/X
{
  const s = content("마무리 O / X", "정리");
  const q = [
    ["BC는 지도학습이므로 훈련 · 테스트 분포가 항상 같다.", "X", "정책이 자기 테스트 분포를 만든다"],
    ["MSE 회귀는 다봉성 데이터에서 두 모드 중 하나를 고른다.", "X", "평균을 낸다"],
    ["BeT는 행동을 군집으로 나누고 오프셋으로 정밀도를 보완한다.", "O", ""],
    ["Diffusion Policy는 행동 16개를 생성해 16개를 모두 실행한다.", "X", "앞 8개만 실행"],
    ["ACT의 chunk 100은 50Hz에서 2초 분량이다.", "O", ""],
  ];
  q.forEach((x, i) => {
    const y = 1.85 + i * 0.98;
    numCircle(s, i + 1, 0.6, y + 0.12, 0.55, C.ink3, 15);
    T(s, x[0], { x: 1.35, y, w: 8.2, h: 0.8, fontSize: 16, valign: "middle" });
    T(s, x[1], { x: 9.7, y, w: 0.6, h: 0.8, fontSize: 24, bold: true, color: x[1] === "O" ? C.good : C.bad, valign: "middle", align: "center" });
    T(s, x[2], { x: 10.4, y, w: 2.3, h: 0.8, fontSize: 12.5, color: C.muted, valign: "middle" });
  });
  s.addNotes(`[2:17–2:19] 1분 O/X. 답(O/X 열)은 발표자 보기에서 숨기고 싶다면 애니메이션 없이 먼저 문제만 읽고 답을 말로 확인하세요.
공지: (선택) 리딩 리포트 — 세 논문 중 1편, 반 페이지 (strengths / weaknesses / question). 다음 주 수업 전날까지.`);
}

// 40. 다음 주
{
  const s = dark();
  T(s, "다음 주 · 5주차", { x: 0.8, y: 1.3, w: 8, h: 0.5, fontSize: 18, color: C.A, bold: true });
  T(s, "강화학습 기초: 시행착오에서 배우기", { x: 0.8, y: 1.9, w: 12, h: 1.0, fontSize: 38, bold: true, color: C.white, valign: "middle" });
  T(s, "\"모방학습의 천장은 시범자의 실력입니다. 사람보다 잘할 수 없고, 사람이 안 보여준 상황엔 약합니다.\n그 천장을 뚫는 도구가 강화학습입니다. 14주차 세미나 논문들이 전부 다음 주 언어로 쓰여 있습니다.\"", { x: 0.8, y: 3.2, w: 11.8, h: 1.6, fontSize: 18, color: "D7DCE3", lineSpacingMultiple: 1.3 });
  T(s, "과제(선택): 리딩 리포트 반 페이지 — BeT / Diffusion Policy / ACT 중 1편", { x: 0.8, y: 5.4, w: 11.8, h: 0.5, fontSize: 15, color: "B8C0CC" });
  vlaDots(s, 0.8, 6.3);
  s.addNotes(`[2:19–2:20] 다음 주 연결 멘트를 그대로 읽고 마칩니다.`);
}

pres.writeFile({ fileName: path.join(__dirname, "Week04_Imitation_Learning.pptx") }).then((f) => console.log("wrote", f, "slides:", pageNo));
