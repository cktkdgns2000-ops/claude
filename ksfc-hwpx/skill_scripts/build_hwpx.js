#!/usr/bin/env node
/**
 * build_hwpx.js — 보고서 마크업(.md) → .hwpx
 *
 * 사용법:  node build_hwpx.js input.md output.hwpx
 *
 * 마크업 해석과 작성 규칙은 build_docx.js와 같다. 문서 구조를 중간 표현(IR)으로 기록한 뒤
 * hwpx_writer.py가 원본 HWP 실측값(160% 줄간격, 앞 공백 층위, 빈 줄 간격 등)으로 hwpx를 쓰고,
 * 한글 줄나눔 모델로 줄 맞춤(어절 단위 기본, 긴 어절만 글자 단위+자간 보정)을 적용한다.
 *
 * 마크업 문법은 references/markup.md 참조.
 * 서체·여백·색상은 아래 CONFIG 상수만 바꾸면 전체에 반영된다.
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak,
  VerticalAlign, TableLayoutType, LineRuleType, VerticalMergeType, ImageRun, Header,
} = require("./ir_shim");

// ───────────────────────────── CONFIG ─────────────────────────────
const MM = 56.6929; // 1mm in DXA
const PT = 2;       // 1pt in half-points (docx size unit)
const CONFIG = {
  page: { top: 20, bottom: 15, left: 25, right: 25 },  // mm
  fonts: {
    title:    { name: "HY헤드라인M", size: 18 },
    meta:     { name: "나눔고딕",   size: 13 },
    header:   { name: "HY헤드라인M", size: 16 },   // "1. 개 요" 형 섹션 제목
    subhdr:   { name: "HY울릉도M",  size: 15 },   // 박스 소제목
    body:     { name: "나눔명조",   size: 15 },   // □ ㅇ -
    keyword:  { name: "HY울릉도M",  size: 15 },   // (키워드)
    note:     { name: "나눔고딕",   size: 13 },   // * ※
    box:      { name: "나눔고딕",   size: 13 },   // 참고 박스 내부
    boxTitle: { name: "HY울릉도M",  size: 13 },
    caption:  { name: "HY울릉도M",  size: 13 },   // < 표 제목 >
    table:    { name: "나눔고딕",   size: 12 },   // @tablefont 로 변경 가능
    conclusion:{ name: "HY울릉도M", size: 15 },  // ➡ 결론 박스
  },
  // 줄간격: HWP 160%는 글자 크기 기준(15pt→약 22.5pt). Word의 배수는 서체 고유 행높이 기준이라
  //          더 벌어지므로 '고정(exact)' 값으로 지정한다. 코퍼스 실측 최빈값: 본문 24pt, 각주 19pt
  lineExact: { body: 24, note: 19, box: 18, table: 16, tableNote: 13.5 },   // box: 코퍼스 박스 안 넘어간 줄 간격 중앙값 18.5pt(사분위 17~19)  // pt  [ledger A3·A4: 24pt 31건/22pt 15건, 각주 중앙값 19]
  spacing: { group: 15, inner: 10, note: 8, header: 18, headerAfter: 4 },  // pt (header: 본문→헤더 간격, 원본 실측 약 18pt)
  // 범위 규칙(코퍼스 실측, ’26.9): 기본값은 중앙값 근처, 쪽 경계가 애매하면 범위 안에서 조절(@box, @cell)
  boxRange: { line: [17, 19], gap: [5, 12], gapDefault: 10 },      // 박스 안: 줄 간격, 항목 앞 여백(항목 사이 중앙값 29pt = 줄 18 + 여백 약 11)
  cell: { line: 1.40, lineRange: [1.30, 1.50], gap: 7, gapRange: [4, 8], gapTop: 11, gapTopRange: [8, 14] },   // 하위 항목 앞(ㅇ·-) 실측 약 7.6pt, 상위 항목 앞(■·➊) 11~16pt   // 표 셀 안 목록: 줄 간격 = 글자×배수(중앙값 16pt/11pt), 항목 앞 여백(항목 사이 중앙값 22pt)
  fit: {
    shrinkParens: 2,      // 본문 괄호 안 글자 크기 -2pt (0이면 끔)
    autoTight: false,     // hwpx: 자동 자간은 hwpx_writer.py가 한글 줄나눔 모델로 판단(Word 폭 추정은 쓰지 않음)
    maxTightPt: 1.6,      // 자동 자간 축소 한도(pt) — 15pt 기준 약 -10%(HWP 자간 조정 관행 범위)
    tightThreshold: 0.5,  // 마지막 줄이 한 줄의 50% 이하만 차지할 때만 자간 축소 시도(한도 내에서)
    maxScaleOver: 0.22,   // □·ㅇ 한 줄 항목이 22%까지 넘치면 장평(글자 폭)으로 한 줄에 맞춤
    safety: 0.02,         // 폭 추정 안전 여유(2%): 경계선 문단은 살짝 좁혀서 넘김 방지
  },
  colors: {
    titleFill: "DCEBFA", titleLine: "4467E9",
    headerText: "1F5FBF", romanFill: "0265CB",
    summaryFill: "FBE5D6", summaryBorder: "9C9C9C",
    keywordHighlight: "DCF6DD", keywordHighlightSub: "FCF6CC",   // □ 연두, ㅇ·➊·① 연노랑 [ledger C2·C16: 2510·2506·2601 픽셀 확인]
    refHighlight: "FCF6CC",   // (☞ 첨부1) 참조 형광
    grayFill: "F0F0F0", grayBorder: "808080",
    frameFill: "E1EED2",
    conclusionBorder: "7F7F7F", blue: "0033CC",
    tintFill: "E4FFFD", tintText: "146990",
    attachFill: "3333A0",          // 첨부·참고 라벨(진파랑, 흰 글씨)
    subhdrFill: "E0EAF3",          // 소제목 박스 연파랑
    tableHeader: "FEF7CD", tableHeaderBlue: "E0EAF3", tableBorder: "333333", tableOuter: "000000",  // [ledger D4/D5: 노랑 24건·연파랑 23건]
    rowGreen: "D3F9EA", rowGray: "E7E6E6", rowYellow: "FEF7CD", firstColFill: "FEF7CD",
    red: "C00000",
  },
};
// 들여쓰기(mm): [left, hanging]
const INDENT = {
  "□": [0, 8.5], "ㅇ": [3.5, 6.5], "-": [8, 5], "·": [8, 5],
  "*": [7.5, 4], "※": [7.5, 4],
  "➊": [6, 6], "➋": [6, 6], "➌": [6, 6], "➍": [6, 6], "➎": [6, 6], "➏": [6, 6], "➐": [6, 6], "➑": [6, 6],
  "❶": [6, 6], "❷": [6, 6], "❸": [6, 6], "❹": [6, 6],
  "①": [6, 6], "②": [6, 6], "③": [6, 6], "④": [6, 6], "⑤": [6, 6], "⑥": [6, 6],
  "➡": [4, 7], "⇨": [4, 7], "⇒": [4, 7], "▶": [4, 6], "☞": [4, 6],
  "■": [2, 5], "‣": [4, 5], "▪": [4, 5], "◆": [1, 5],
 "◈": [0, 6], "◇": [0, 6], "∙": [4, 4], "•": [4, 4], "❺": [6, 6], "❻": [6, 6], "❼": [6, 6], "❽": [6, 6], "❾": [6, 6], "❿": [6, 6], "⑦": [6, 6], "⑧": [6, 6], "⑨": [6, 6], "⑩": [6, 6], "➒": [6, 6], "➓": [6, 6],
};
// 층위 → 기호 시작 위치(mm). 층위는 기호가 아니라 문장의 논리 구조로 정한다(작성자 확인):
//   마크업 줄 앞 공백(2칸 = 한 단계)으로 직접 표시하거나, 표시가 없으면 문맥(기호가 처음 나온 순서)으로 추정
const LEVEL_LEFT = [0, 3.5, 8, 12, 16];
const LEVEL_MARKERS = /^([□ㅇ\-·∙■▶‣▪◆○●◎▷►]|[①-⑳]|[➊-➓]|[❶-❿])$/;
const ARROW_LV = /^[➡⇨⇒]$/;
/** 기호 종류: ➊➋➌…, ①②③…, ❶❷❸…는 각각 한 종류(같은 층위) */
const markerFamily = (m) => /^[➊-➓]$/.test(m) ? "➊" : (/^[①-⑳]$/.test(m) ? "①" : (/^[❶-❿]$/.test(m) ? "❶" : m));
// 양쪽 정렬(점검용으로 FP_LEFT=1이면 왼쪽 정렬: 줄의 자연 폭을 재어 '간신히 들어간 줄'을 찾기 위함)
const JUSTIFY = process.env.FP_LEFT ? AlignmentType.LEFT : AlignmentType.BOTH;
const BOX_LEAD = /^([➡⇨⇒◈◇▶►▷☞■□▪‣•∙◆○●※*]|[➊-➓]|[①-⑳]|[❶-❿])$/;   // 결론·프레이밍·요약 박스 줄의 앞머리 기호(둘째 줄 내어쓰기 대상)
const BULLET_NOTES = new Set(["‣", "▶", "▪"]);   // 글꼴은 각주처럼 작지만 위치는 층위를 따르는 나열 기호
const BOX_TOP = /^([□■▶‣▪◆·∙]|[①②③④⑤⑥⑦⑧⑨]|[➊➋➌➍➎➏➐➑]|[❶❷❸❹❺❻])$/;
const NOTE_MARKERS = new Set(["*", "**", "※", "‣", "▶", "▪", "1)", "2)", "3)", "4)", "5)", "6)", "7)", "8)", "9)"]);
for (const n of "123456789") INDENT[n + ")"] = [7.5, 5];
INDENT["**"] = [7.5, 5.5];
for (const k of ["◈", "◇", "○", "●", "◎", "▷", "►", "∙", "❺", "❻", "❼", "❽", "❾", "❿", "⑦", "⑧", "⑨", "⑩", "➒", "➓"]) if (!INDENT[k]) INDENT[k] = [0, 6];   // 박스 안 ◈ 소제목 줄 등(누락 점검으로 추가)   // 둘째 각주 표시(줄 첫머리 "** ")
const ARROW_MARKERS = new Set(["➡", "⇨", "⇒"]);
let WIDTH_OVERRIDE = null;   // ::: cols 안에서는 열 폭
const CONTENT_W = () => WIDTH_OVERRIDE || (210 - CONFIG.page.left - CONFIG.page.right) * MM;

// ───────────────────────────── helpers ─────────────────────────────
const font = (f) => ({ font: { name: f.name, eastAsia: f.name, hint: "eastAsia" }, size: f.size * PT });
const HEAVY_FONTS = new Set(["HY울릉도M"]);            // 자체 굵기 있는 서체: bold 지정을 무시
const emph = (f) => (HEAVY_FONTS.has(f.name) ? {} : { bold: true });
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const border = (color, size = 4, style = BorderStyle.SINGLE) => ({ style, size, color });
const allBorders = (b) => ({ top: b, left: b, bottom: b, right: b }); // 스키마 순서 top→left→bottom→right
const shade = (fill) => ({ type: ShadingType.CLEAR, color: "auto", fill });
const lineOf = (pt = CONFIG.lineExact.body) => ({ line: Math.round(pt * 20), lineRule: LineRuleType.EXACT });
const lineAuto = (mult = 1.15) => ({ line: Math.round(240 * mult), lineRule: LineRuleType.AUTO });
const mmSp = (pt) => pt * 20; // pt → twips for spacing

/** 인라인 마크업 → TextRun[]
 *  **굵게**  __밑줄__  {{파란 굵게}}  ^^위첨자^^  ~~작게(병기)~~
 */
function inline(text, base, opts = {}) {
  if (typeof text === "string") text = textVS(text);   // ➡ 등 이모지 가능 기호는 글자 모양으로
  const runs = [];
  const { shrinkParens = 0, noRefHl = false, noLead = false, parenAll = false, ...runOpts } = opts;
  // (☞ 첨부1) 류 참조는 연노랑 형광 [ledger C6′: 2409·2604·2605]
  if (!noRefHl) text = text.replace(/(☞\s*(?:첨부|참고|별첨)\s*[\d~,\-]*)/g, "⟦$1⟧");
  // 문장 맨 앞의 괄호 라벨((참고1), (정원) 등)은 부연이 아니라 머리말 → 축소하지 않음(작성자 확인)
  if (shrinkParens && !noLead) text = text.replace(/^(\s*(?:\*\*)?\s*(?:(?:[∙•·▪■‣※]|-(?=\s)|\*(?=\s))\s*)?(?:\*\*)?)(\([^()]{1,20}\))/, "$1⟪$2⟫");   // ※·* 뒤 첫 괄호 라벨 포함
  // 위첨자 둘째 각주 표시 ^^**^^ 가 굵게(**) 구문과 충돌하지 않도록 임시 문자로 치환
  text = text.replace(/\^\^\*\*\^\^/g, "^^\u2051^^");
  const re = /(==.+?==|⟪.+?⟫|⟦.+?⟧|\[s\].+?\[\/s\]|\*\*.+?\*\*|__.+?__|\{\{.+?\}\}|!!.+?!!|\^\^.+?\^\^|~~.+?~~)/g;
  let last = 0, m;
  const heavy = HEAVY_FONTS.has(base.name);   // 자체 굵은 서체에는 **굵게**·{{ }}의 bold를 적용하지 않음
  const pushRaw = (t, extra = {}) => {
    if (!t) return;
    const o = { text: t, ...font(base), ...runOpts, ...extra };
    if (heavy) delete o.bold;
    runs.push(new TextRun(o));
  };
  // 괄호 안 글자는 -2pt (원본 관행: (5/20), (’25.7월), (은행 등) 등)
  // 괄호 깊이를 토큰 사이에서도 유지 → 괄호 안에 위첨자 등이 끼어 있어도 괄호 전체를 -2pt(작성자 확인)
  let parenDepth = 0;
  const small = (extra) => ({ ...extra, size: Math.max(9, base.size - shrinkParens) * PT });
  const push = (t, extra = {}) => {
    if (!t) return;
    if (!shrinkParens) return pushRaw(t, extra);
    let buf = "";
    const flush = () => { if (buf) pushRaw(buf, (parenDepth > 0 || parenAll) ? small(extra) : extra); buf = ""; };
    for (const ch of t) {
      if (ch === "(") { flush(); parenDepth++; buf = "("; }
      else if (ch === ")") { buf += ")"; flush(); parenDepth = Math.max(0, parenDepth - 1); }
      else buf += ch;
    }
    flush();
  };
  while ((m = re.exec(text)) !== null) {
    push(text.slice(last, m.index));
    const tok = m[0];
    if (tok.startsWith("==")) runs.push(...inline(tok.slice(2, -2), base, { ...opts, parenAll: parenAll || parenDepth > 0, noLead: true, shading: shade(CONFIG.colors.refHighlight) }));   // ==연노랑 형광==
    else if (tok.startsWith("⟪")) runs.push(...inline(tok.slice(1, -1), base, { ...opts, noLead: true, shrinkParens: 0 }));   // 머리 괄호 라벨: 원 크기, 안의 위첨자 등은 그대로 처리
    else if (tok.startsWith("[s]")) runs.push(...inline(tok.slice(3, -4), base, { ...opts, parenAll: parenAll || parenDepth > 0, noLead: true, strike: true }));
    else if (tok.startsWith("⟦")) runs.push(...inline(tok.slice(1, -1), base, { ...opts, parenAll: parenAll || parenDepth > 0, noLead: true, noRefHl: true, shading: shade(CONFIG.colors.refHighlight) }));
    else if (tok.startsWith("**")) runs.push(...inline(tok.slice(2, -2), base, { ...opts, parenAll: parenAll || parenDepth > 0, noLead: true, bold: true }));
    else if (tok.startsWith("__")) runs.push(...inline(tok.slice(2, -2), base, { ...opts, parenAll: parenAll || parenDepth > 0, noLead: true, underline: {} }));
    else if (tok.startsWith("{{")) runs.push(...inline(tok.slice(2, -2), base, { ...opts, parenAll: parenAll || parenDepth > 0, noLead: true, bold: true, color: CONFIG.colors.blue }));
    else if (tok.startsWith("!!")) runs.push(...inline(tok.slice(2, -2), base, { ...opts, parenAll: parenAll || parenDepth > 0, noLead: true, color: CONFIG.colors.red }));
    else if (tok.startsWith("^^")) pushRaw(tok.slice(2, -2).replace(/\u2051/g, "**"), (parenDepth > 0 || parenAll) && shrinkParens ? small({ superScript: true }) : { superScript: true });   // 괄호 안 위첨자도 괄호 크기에 맞춤
    else if (tok.startsWith("~~")) pushRaw(tok.slice(2, -2), { size: Math.round(base.size * 0.8) * PT });
    last = m.index + tok.length;
  }
  push(text.slice(last));
  return runs;
}

/** 줄 폭 추정(pt). 나눔고딕·나눔명조 실측 advance: 한글 0.94em, 공백 0.28, 숫자 0.58, 대문자 0.73, 소문자 0.53, 문장부호 0.33 */
function estWidthPt(text, sizePt, shrinkParens = 0, fontName = "") {
  // HY계열(울릉도M·헤드라인M)은 한글 1em, 나눔명조 0.95, 나눔고딕 0.94
  const hangulEm = /^HY/.test(fontName) ? 1.0 : (/명조/.test(fontName) ? 0.95 : 0.94);
  let w = 0, depth = 0, bold = false;
  // 위첨자(^^…^^)는 실제로 약 60% 크기로 찍히므로 폭도 그만큼만 센다(전폭으로 세면 넘침을 과대 추정해 자동 자간 축소가 안 걸림)
  const clean = text.replace(/__|\{\{|\}\}|~~|⟦|⟧|⟪|⟫|==|!!|\[s\]|\[\/s\]|\u200B|\uFE0E/g, "").replace(/\^\^(.+?)\^\^/g, "\u0001$1\u0002");
  let sup = false;
  for (let i = 0; i < clean.length; i++) {
    const ch = clean[i];
    if (ch === "\u0001") { sup = true; continue; }
    if (ch === "\u0002") { sup = false; continue; }
    if (ch === "*" && clean[i + 1] === "*") { bold = !bold; i++; continue; }
    const sz = (depth > 0 && shrinkParens ? sizePt - shrinkParens : sizePt) * (bold ? 1.03 : 1) * (sup ? 0.6 : 1);
    if (ch === "(") depth++;
    if (/[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF\u4E00-\u9FFF\u3000-\u303F\uFF00-\uFFEF\u2460-\u24FF\u2776-\u2793\u25A0-\u25FF\u2190-\u21FF\u2600-\u27BF]/.test(ch)) w += sz * hangulEm;
    else if (ch === " ") w += sz * 0.28;
    else if (/[0-9]/.test(ch)) w += sz * 0.58;
    else if (/[A-Z]/.test(ch)) w += sz * 0.73;
    else if (/[a-z]/.test(ch)) w += sz * 0.53;
    else w += sz * 0.33;
    if (ch === ")") depth--;
  }
  return w;
}
/** 줄 끝 빈 공간이 크게 남을 때(다음 어절이 길어 통째로 넘어가는 경우) 그 어절을 자연스러운 경계에서 나눌 수 있게 한다.
 *  허용 경계: 'ㆍ'·'·'·'/' 뒤, 복합어 구성 단어 사이(스테이블/코인, 증권/결제용 등). 글자 한두 개만 떼어 내는 분리는 하지 않음.
 *  구현: 해당 위치에 폭 없는 줄바꿈 기회(U+200B)를 넣고, 실제 줄바꿈은 Word가 판단 (작성자 확인: 예외적으로 어절 분리 허용) */
const SPLIT_WORDS = ["스테이블", "코인", "토큰", "증권", "결제", "디지털", "자산", "가상자산", "투자자", "예탁금", "예탁", "자본시장", "인프라",
  "기관", "컨소시엄", "플랫폼", "거래소", "장외", "조각투자", "분산원장", "블록체인", "금융위", "금융", "시장", "제도", "법안", "연구용역", "보고서",
  "별도예치", "예치", "보관", "운용", "발행", "수탁", "담보", "대출", "전자", "관리", "시스템", "정책", "입법", "규제", "투자"];
/** 박스 안 항목 앞 여백: 기본 10pt, @spacing으로 쪽을 맞출 때 함께 줄되 범위(5~12pt) 안에서 */
const boxGap = () => Math.max(CONFIG.boxRange.gap[0], Math.min(CONFIG.boxRange.gap[1], CONFIG.boxRange.gapOverride ?? Math.min(CONFIG.boxRange.gapDefault, CONFIG.spacing.inner)));
function softSplit(text) {
  // 긴 복합어(한글 5자 이상)의 자연스러운 경계에 줄바꿈 기회(U+200B)를 둔다. Word는 그 어절이 줄 끝에 걸릴 때만 그 자리에서 나누므로
  // 줄 끝 빈 공간이 크게 남는 경우에만 분리가 일어나고, 한 줄에 다 들어가면 아무 변화가 없다(폭 추정 오차와 무관하게 안전)
  return text.split(/(\s+)/).map(tk0 => {
    // 어절 안의 괄호 앞(태광그룹/(흥국생명 등))도 줄바꿈 기회: 앞 글자가 한글·영숫자·닫는 부호이고 괄호 안이 2자 이상일 때(작성자 확인)
    const tk = tk0.replace(/([가-힣A-Za-z0-9%」’)](?:\*\*|__|\}\})?)(?=\([^()\s]{2,})/g, "$1\u200B")   // 굵게·밑줄 닫힘 뒤의 괄호도 포함
                  .replace(/([가-힣A-Za-z0-9)])([ㆍ·])(?=[가-힣A-Za-z]{2,})/g, "$1$2\u200B");        // 가운뎃점 뒤(… 고객확인(KYC)ㆍ/자금세탁방지)
    if (tk !== tk0 && tk.replace(/[^가-힣]/g, "").length < 5) return tk;
    const m = tk.match(/^((?:\*\*|__|\{\{)*)([가-힣ㆍ·\/]+)(.*)$/s);
    if (!m || m[2].replace(/[ㆍ·\/]/g, "").length < 5) return tk;
    const core = m[2]; const cands = new Set();
    for (let q = 2; q <= core.length - 2; q++) if (/[ㆍ·\/]/.test(core[q - 1])) cands.add(q);
    for (const dw of SPLIT_WORDS) { let at = core.indexOf(dw); while (at >= 0) { if (at >= 2 && core.length - at >= 2) cands.add(at); at = core.indexOf(dw, at + 1); } }
    // 사전 단어 안쪽에 걸리는 경계 제외(가상자산·스테이블 등은 통째로)
    for (const q of [...cands]) for (const dw of SPLIT_WORDS) { let at = core.indexOf(dw); while (at >= 0) { if (q > at && q < at + dw.length) cands.delete(q); at = core.indexOf(dw, at + 1); } }
    if (!cands.size) return tk;
    let out = ""; for (let q = 0; q < core.length; q++) { if (cands.has(q)) out += "\u200B"; out += core[q]; }
    return m[1] + out + m[3];
  }).join("");
}
const EMOJI_CAPABLE = /([\u27A1\u25B6\u25C0\u25AA\u25AB\u25FC\u25FB\u2B05-\u2B07\u2194-\u2199\u21A9\u21AA\u2714\u2716\u2611\u2934\u2935\u203C\u2049])(?!\uFE0E)/g;
const textVS = (t) => t.replace(EMOJI_CAPABLE, "$1\uFE0E");   // Word: ➡ 등은 기본이 컬러 이모지(폭 넓음) → 글자 모양 강제
/** 박스·표 셀 등 본문 밖 문단에도 같은 줄 넘침 보정을 적용: [줄바꿈 기회를 넣은 글, 자간 옵션] */
function lineParts(text) {
  // 줄 끝 {tight:x}(최소 자간 축소)와 문장 끝 ' ※ 첨언'(오른쪽 끝 작은 글씨)을 떼어 냄 — 본문·박스·표 셀·타임라인 공통
  let t = text, forcedTw = null, tail = null;
  const tm = t.match(/\s*\{tight(?::([\d.]+))?\}\s*$/);
  if (tm) { forcedTw = -Math.round((tm[1] ? parseFloat(tm[1]) : 0.8) * 20); t = t.slice(0, tm.index); }
  const lm = t.match(/\s*\{loose(?::([\d.]+))?\}\s*$/);
  if (lm) { forcedTw = Math.round((lm[1] ? parseFloat(lm[1]) : 0.3) * 20); t = t.slice(0, lm.index); }
  const m2 = t.match(/^(.*\S)\s+(※\s+.+)$/s);
  if (m2 && !/^\s*※/.test(t) && m2[1].replace(/[*_{}^]/g, "").length > 8) { t = m2[1]; tail = m2[2]; }
  return { t, forcedTw, tail };
}
function fitLine(text, font, availPt, { kwText = "", shrink = CONFIG.fit.shrinkParens, forcedTw = null } = {}) {
  const t = softSplit(text);
  const cs = autoFit(t, font.size, availPt, shrink, { fontName: font.name, kwText, allowScale: false }) || {};
  let tw = cs.characterSpacing != null ? cs.characterSpacing : null;
  if (forcedTw != null) tw = forcedTw > 0 ? forcedTw : (tw != null ? Math.min(tw, forcedTw) : forcedTw);   // {tight}는 최소값, {loose}는 그대로
  return [t, tw != null ? { characterSpacing: tw } : {}];
}
function tailFitsInline(main, tail, baseFont, availPt) {
  // 본문장과 첨언을 이어 써도 한 줄에 들어가면 오른쪽 끝으로 떼지 않는다(첨언 올려 쓰기는 문장이 줄을 넘길 때의 기법)
  if (!tail || !availPt) return false;
  const w = estWidthPt(main, baseFont.size, CONFIG.fit.shrinkParens, baseFont.name) + estWidthPt(" " + tail, baseFont.size - 2, 0, CONFIG.fonts.note.name);
  return w * 1.04 <= availPt;
}
function tailRuns(tail, baseFont, inlineTail = false) {
  if (!tail) return [];
  const tf = { ...CONFIG.fonts.note, size: baseFont.size - 2 };
  return [new TextRun({ text: inlineTail ? " " : "\t", ...font(tf) }), ...inline(tail, tf, { shrinkParens: 0 })];
}
/** 넘침 보정: 자간(characterSpacing, twips) 또는 장평(scale, %)을 돌려준다.
 *  - 마지막 줄이 조금 넘칠 때: 글자당 필요 축소가 maxTightPt 이내면 자간 축소
 *  - 한 줄에서 최대 22%까지 넘칠 때: HWP 장평처럼 글자 폭 축소(원본에서 □ 한 줄 항목에 흔히 사용)
 */
function autoFit(text, sizePt, availPt, shrinkParens, { fontName = "", kwText = "", allowScale = true } = {}) {
  if (!CONFIG.fit.autoTight) return {};
  const w = estWidthPt(kwText, sizePt, 0, CONFIG.fonts.keyword.name) + estWidthPt(text, sizePt, shrinkParens, fontName);
  const cap = availPt * (1 - CONFIG.fit.safety);
  const lines = w / cap;
  if (lines <= 1) return {};
  const nchar = text.replace(/\s/g, "").length || 1;
  // 한글은 어절 단위로 줄이 바뀌므로 글자 수 비례보다 줄이 더 빨리 넘어간다 → 어절 단위로 줄을 채워 마지막 줄 길이를 구함
  const words = (kwText + text).split(/\s+/).filter(Boolean);
  let nLines = 1, cur = 0; const sp = sizePt * 0.28;
  for (const wd of words) {
    const ww = estWidthPt(wd, sizePt, shrinkParens, fontName);
    if (cur > 0 && cur + sp + ww > cap) { nLines++; cur = ww; } else cur += (cur > 0 ? sp : 0) + ww;
  }
  const lastFill = cur / cap;                    // 마지막 줄이 차지하는 비율
  let needA = null, needB = null;
  // (A) 글자 수 비례 추정: 전체 폭이 줄 수의 정수배를 조금 넘을 때
  const over = lines - Math.floor(lines);
  if (over > 0 && over <= CONFIG.fit.tightThreshold) { const n = (over * availPt) / nchar + 0.05; if (n <= CONFIG.fit.maxTightPt) needA = n; }
  // (B) 어절 단위 추정: 마지막 줄의 어절들을 앞 줄로 끌어올리는 데 필요한 글자당 축소량
  if (nLines >= 2 && lastFill <= CONFIG.fit.tightThreshold) {
    const prevChars = Math.max(1, nchar - (words.length ? words[words.length - 1].length : 0));
    const n = ((cur + sp) / prevChars) * 1.3 + 0.1; if (n <= CONFIG.fit.maxTightPt) needB = n;   // 폭 추정이 실제보다 작게 나오는 경향(미리보기 실측 보정 ×1.3)
  }
  // (C) 첫 줄 끝 빈 공간: 다음 조각(띄어쓰기·줄바꿈 기회 U+200B까지)이 조금만 모자라 통째로 넘어가면, 첫 줄을 조금 좁혀 끌어올림
  //     예: "…15억원(’25.7월) 및 / 태광그룹(흥국생명 등)" → "… 및 태광그룹 / (흥국생명 등)"
  let needC = null;
  {
    const segs = (kwText + text).split(/(\s+|\u200B)/).filter(x => x !== "");
    let cur = 0, n1 = 0;
    for (const sg of segs) {
      if (/^\s+$/.test(sg)) { cur += sizePt * 0.28; continue; }
      if (sg === "\u200B") continue;
      const ww = estWidthPt(sg, sizePt, shrinkParens, fontName);
      if (cur + ww <= cap) { cur += ww; n1 += sg.replace(/[\s*_{}^]/g, "").length; continue; }
      const short = cur + ww - cap;                       // 모자라는 폭
      const blankRatio = (cap - cur) / cap;
      if (blankRatio > 0.06 && n1 > 0) { const n = short / n1 * 1.15 + 0.1; if (n <= CONFIG.fit.maxTightPt) needC = n; }
      break;
    }
  }
  const need = Math.max(needA || 0, needB || 0, needC || 0);
  if (need > 0) return { characterSpacing: -Math.round(need * 20) };
  if (allowScale && lines <= 1 + CONFIG.fit.maxScaleOver) {   // 한 줄로 밀어 넣기: 장평(□·ㅇ 항목만)
    const scale = Math.floor(100 / lines) - 1;
    return { scale, characterSpacing: -6 };
  }
  return {};
}

/** 표 셀·타임라인 안의 한 줄: 기호(■ - ∙ ㅇ * ※ ➊ ① 1) 등)로 시작하면 둘째 줄부터 기호 다음 글자 시작점에 맞춤(내어쓰기) */
function cellHang(line, sizePt) {
  const m = line.match(/^((?:[■ㅇ·•∙▪‣▶◆➊➋➌➍➎①②③④⑤⑥❶❷❸❹\-\*※]|\*\*|\d\))\s+)\S/);
  if (!m) return undefined;
  const tw = Math.round(estWidthPt(m[1], sizePt) * 20);
  return { left: tw, hanging: tw };
}
function cellIndent(line, sizePt, lv) {
  const h = cellHang(line, sizePt);
  if (!h) return undefined;
  const base = lv > 1 ? Math.round(LEVEL_LEFT[Math.min(lv, LEVEL_LEFT.length) - 1] * (sizePt / 15) * MM) : 0;
  return { left: base + h.left, hanging: h.hanging };
}
/** 본문 한 줄(마커 포함) → Paragraph */
function bodyParagraph(line, ctx = {}) {
  const inBox = !!ctx.inBox;
  const m = line.match(/^(\S+)\s+(.*)$/s);
  let marker = m ? m[1] : null, rest = m ? m[2] : line, glued = false;
  if (!marker || !INDENT[marker]) {
    const g = line.match(/^([➊-➓①-⑳❶-❿])(\S.*)$/s);   // 흐름 나열(➊자료수집 → ➋…)처럼 기호에 글자가 붙은 줄
    if (g && INDENT[g[1]]) { marker = g[1]; rest = g[2]; glued = true; } else { marker = null; rest = line; }
  }

  const isNote = !!(marker && NOTE_MARKERS.has(marker));
  const isArrow = !!(marker && ARROW_MARKERS.has(marker));
  const boxF = ctx.boxFs ? { ...CONFIG.fonts.box, size: ctx.boxFs } : CONFIG.fonts.box;
  let baseFont = inBox ? boxF : (isNote ? CONFIG.fonts.note : CONFIG.fonts.body);
  // 박스 안에서 내용 전체가 굵은 ※ 줄은 각주가 아니라 소제목(예: ※ **(참고1) …**) → 박스 글자 크기 유지
  const boxSubTitle = inBox && marker === "※" && /^\*\*[^*]+\*\*\s*$/.test(rest.trim());
  if (isNote && (ctx.afterTable || inBox) && !boxSubTitle) baseFont = { ...(inBox ? boxF : CONFIG.fonts.note), size: (inBox ? boxF.size : CONFIG.fonts.note.size) - 2 };   // 표 각주·박스 안 각주는 2pt 작게
  const bold = false;   // 본문 화살표 줄은 본문 서체, 굵기는 마크업(**)으로 [ledger C9·C10 정정: 2604·2605·2607]

  // 줄 끝 지시어: {small} 또는 {size:-2} → 그 줄만 글자 크기 조정(한 줄만 삐져나올 때 각주를 한 줄로 만드는 용도)
  const zm = rest.match(/\s*\{(small|size:(-?\d+))\}\s*$/);
  if (zm) { baseFont = { ...baseFont, size: Math.max(8, baseFont.size + (zm[2] ? parseInt(zm[2], 10) : -2)) }; rest = rest.slice(0, zm.index); }
  // 줄 끝 지시어: {right:텍스트} 오른쪽 끝 정렬 꼬리말((단위 : 백만원), (’25.5월말 기준))
  let rightTail = null, autoTail = false;
  const preTight = rest.match(/\s*\{(?:tight|loose)(?::[\d.]+)?\}\s*$/);
  const tightSuffix = preTight ? preTight[0] : ""; if (preTight) rest = rest.slice(0, preTight.index);
  const rm = rest.match(/\s*\{right:([^}]*)\}\s*$/);
  if (rm) { rightTail = rm[1]; rest = rest.slice(0, rm.index); }
  // 문장 끝에 붙인 첨언 " ※ …" → 2pt 작게, 줄 오른쪽 끝에 붙임(작성자 확인: 넘친 줄에 아래 ※ 첨언을 올려 쓸 때)
  if (!rightTail && marker !== "※") {
    const tm2 = rest.match(/^(.*\S)\s+(※\s+.+)$/s);
    if (tm2 && tm2[1].length > 8) { rest = tm2[1]; rightTail = tm2[2]; autoTail = true; }
  }
  rest = rest + tightSuffix;
  // 줄 끝 지시어: {tight} / {tight:1.2} 자간 강제 축소, {narrow:88} 장평(글자 폭 %) 강제
  let forced = null;
  const tm = rest.match(/\s*\{tight(?::([\d.]+))?\}\s*$/);
  if (tm) { forced = { characterSpacing: -Math.round((tm[1] ? parseFloat(tm[1]) : 0.8) * 20) }; rest = rest.slice(0, tm.index); }
  const lm0 = rest.match(/\s*\{loose(?::([\d.]+))?\}\s*$/);
  if (lm0) { forced = { characterSpacing: Math.round((lm0[1] ? parseFloat(lm0[1]) : 0.3) * 20) }; rest = rest.slice(0, lm0.index); }
  const nm = rest.match(/\s*\{narrow(?::(\d+))?\}\s*$/);
  if (nm) { forced = { ...(forced || {}), scale: nm[1] ? parseInt(nm[1], 10) : 90 }; rest = rest.slice(0, nm.index); }

  let [left, hang] = marker ? INDENT[marker] : [0, 0];
  if (ctx.level && marker && (!NOTE_MARKERS.has(marker) || BULLET_NOTES.has(marker))) left = LEVEL_LEFT[Math.min(ctx.level, LEVEL_LEFT.length) - 1];   // 층위 기반 위치
  // 박스 안에서 박스의 첫 층위 기호(□ ■ ▶ ‣ ① ➊ ◆ 등)는 본문용 들여쓰기 없이 박스 안쪽 끝에서 시작(작성자 확인). ㅇ·-는 □ 기준 상대 들여쓰기 유지
  if (inBox && marker === "※" && /^\*\*[^*]+\*\*\s*$/.test(rest.trim())) { left = 0; hang = 4.5; }   // 박스 안 ※ 소제목은 왼쪽 끝
  // 내어쓰기 폭 = 기호 + 한 칸 폭(실측 추정). 둘째 줄부터 첫 단어 시작점에 맞춤(작성자 확인: 기호 바로 다음 단어의 시작점)
  if (marker && !(isNote && (ctx.afterTable || ctx.afterBox))) {
    const hf = inBox ? CONFIG.fonts.box : (isNote ? CONFIG.fonts.note : CONFIG.fonts.body);
    const need = (estWidthPt(marker + (glued ? "" : " "), hf.size, 0, hf.name) + (glued ? 0 : hf.size * 0.15)) * 0.3528;   // pt → mm (+0.15em 여유: Word 대체 글꼴 폭 대비)
    const markerOnly = estWidthPt(marker, hf.size, 0, hf.name) * 0.3528 + (glued ? 0 : 0.8);
    hang = Math.max(need, markerOnly);
  }   // 문맥 층위(헤더 직하 ①, 그 아래 ❶)
  if (isNote && (ctx.afterTable || ctx.afterBox)) { left = ctx.subNote ? 5.5 : 0; hang = ctx.afterBox ? 4.5 : (ctx.subNote ? 3.5 : 5.5); }   // 표 각주·박스 아래·표 제목 ※는 왼쪽 끝 기준, 1) 아래의 * 는 1) 글자 시작에
  else if (isNote && ctx.noteLeft != null) { left = ctx.noteLeft; hang = marker === "**" ? 5.5 : (marker === "※" ? 4.5 : 4); }  // 각주 시작 = 바로 위 문장의 글자 시작 위치(박스 안 포함)
  if (boxSubTitle) { left = 0; hang = 4.5; }   // 박스 안 ※ 소제목(표·박스 제목)은 각주 규칙보다 우선해 왼쪽 끝
  const availPt = (CONTENT_W() - left * MM - hang * MM - (inBox ? 300 : 0)) / 20;
  const shrink = CONFIG.fit.shrinkParens;   // 박스 안도 동일(같은 층위 본문 대비 -2pt)

  // (키워드) 소제목
  let kwRun = null, kwText = "";
  const kw = rest.match(/^(\((?:[^()]|\([^()]*\)){1,40}\))\s*(.*)$/s);   // 중첩 괄호 허용: (주주 구성(예상))
  if (kw && marker && !isNote) {
    const kwFont = { ...CONFIG.fonts.keyword, size: baseFont.size };
    const lvl = ctx.level || (marker === "□" ? 1 : 2);
    const hlColor = (inBox && !ctx.boxHl) ? null : (lvl === 1 ? CONFIG.colors.keywordHighlight : CONFIG.colors.keywordHighlightSub);   // 최상위 연두, 하위 연노랑
    const highlight = !!hlColor;
    // HY울릉도M은 자체가 굵은 서체 → 볼드 처리 안 함. 위첨자 등 인라인 마크업은 살림
    kwRun = { text: kw[1], kwFont, highlight, hlColor };
    kwText = kw[1] + " ";
    rest = kw[2];
  }
  // 줄 끝 빈 공간이 크게 남는 경우 긴 어절을 자연스러운 경계에서 나눌 수 있게(본문·박스)
  if (!forced) {
    rest = softSplit(rest);   // 본문·각주·표 각주·박스 공통
  }
  // 직접 지정한 {tight}는 '최소한 이만큼'으로 보고, 자동 판정이 더 필요하다고 보면 더 줄인다(한도 안에서)
  const autoCs = autoFit(rest, baseFont.size, availPt, shrink, { fontName: baseFont.name, kwText, allowScale: marker === "□" || marker === "ㅇ" });
  let cs = forced || autoCs;
  if (forced && forced.characterSpacing != null && forced.characterSpacing < 0 && autoCs && autoCs.characterSpacing != null && autoCs.characterSpacing < forced.characterSpacing)
    cs = { ...forced, characterSpacing: autoCs.characterSpacing };

  const runs = [];
  const tableNote = isNote && ctx.afterTable;   // 박스 안의 표 각주도 동일
  if (marker && tableNote) {
    // 표 각주: 마커를 표 왼쪽 끝에, *는 한 칸 들여 **와 글자 시작 위치를 맞춤(원본 관행), 마커 뒤 탭으로 본문 정렬
    const mk = marker === "*" ? "\u00A0*" : marker;
    runs.push(new TextRun({ text: mk + "\t", ...font(baseFont), bold }));
  } else if (marker) runs.push(new TextRun({ text: textVS(marker) + (glued ? "" : "\t"), ...font(baseFont), bold }));   // 기호 뒤 탭 → 첫 단어 시작점 = 내어쓰기 위치(둘째 줄과 일치)
  if (kwRun) {
    runs.push(...inline(kwRun.text, kwRun.kwFont, { ...cs, ...(kwRun.highlight ? { shading: shade(kwRun.hlColor) } : {}) }));
    runs.push(new TextRun({ text: " ", ...font(baseFont), ...cs }));
  }
  runs.push(...inline(rest, baseFont, { bold, shrinkParens: shrink, ...cs }));

  const tailInline = autoTail && tailFitsInline((kwText || "") + rest, rightTail, baseFont, availPt);
  if (rightTail) {
    const tailFont = { ...CONFIG.fonts.note, size: baseFont.size - 2 };
    runs.push(new TextRun({ text: tailInline ? " " : "\t", ...font(tailFont) }));
    runs.push(...inline(rightTail, tailFont, { shrinkParens: 0 }));
  }
  let before = marker === "□" ? CONFIG.spacing.group : (isNote ? CONFIG.spacing.note : CONFIG.spacing.inner);
  if (ctx.afterTable || ctx.tightGap) before = 2;
  const linePt = (isNote && ctx.afterTable) ? CONFIG.lineExact.tableNote   // 표 아래 각주는 촘촘(2505·2506·2509 실측 12.8~13.4pt)
    : (inBox ? CONFIG.lineExact.box : (isNote ? CONFIG.lineExact.note : CONFIG.lineExact.body));
  return new Paragraph({
    ...pb(), children: runs, keepLines: true,
    _cls: isNote ? (ctx.afterTable ? "tnote" : "note") : (inBox ? "box" : "body"), _marker: marker || null, _level: ctx.level || 0, _glued: glued,
    alignment: JUSTIFY,   // 양쪽 정렬: HWP 원본과 같이 줄 끝을 오른쪽 여백에 맞춤(Word 기준)
    ...(rightTail && !tailInline ? { tabStops: [{ type: "right", position: Math.round(CONTENT_W() - (inBox ? 300 : 0)) }] } : {}),
    indent: { left: Math.round(left * MM + hang * MM), hanging: Math.round(hang * MM) },
    // 박스 안 문단 간격(원본 실측): □ 항목 사이 약 12pt, ①·■·ㅇ·➊ 등 항목 사이 약 10pt(항목 간 줄 간격 26.9~27.5pt − 줄 17pt),
    // ‣·▶ 한 줄 나열은 촘촘하게(2409: 18.6pt 간격), 각주·첨언은 4pt
    // (@spacing으로 쪽을 맞출 때 박스 안 간격도 함께 줄어들도록 본문 간격 값을 상한으로 사용)
    spacing: { before: mmSp(inBox ? (ctx.firstInBox ? (ctx.boxTitled ? 4 : 0) : marker === "□" ? boxGap() + 2 : (BULLET_NOTES.has(marker) ? (ctx.tightBox ? 1.5 : 4) : (isNote ? 4 : boxGap()))) : before), after: 0, ...lineOf(linePt) },
  });
}

/** "◈ (키워드) 본문" 류 한 줄을 런으로: 키워드는 원 크기, 나머지는 괄호 안 -2pt */
function runsWithKeyword(text, f, opts = {}) {
  const kw = text.match(/^(\((?:[^()]|\([^()]*\)){1,40}\))\s*(.*)$/s);
  if (!kw) return inline(text, f, { ...opts, shrinkParens: CONFIG.fit.shrinkParens });
  return [...inline(kw[1], f, opts), new TextRun({ text: " ", ...font(f) }), ...inline(kw[2], f, { ...opts, shrinkParens: CONFIG.fit.shrinkParens })];
}
/** 1셀 박스 표 */
const BOX_PAD_TW = 100;   // 박스 위·아래 안쪽 여백(twips, 약 5pt)
function boxTable(paragraphs, { fill, borderSpec, widthDxa, marginTw = 100, indentMm = 0 } = {}) {
  const w = widthDxa || (CONTENT_W() - Math.round(indentMm * MM));
  return new Table({
    width: { size: w, type: WidthType.DXA }, columnWidths: [w],
    ...(indentMm ? { indent: { size: Math.round(indentMm * MM), type: WidthType.DXA } } : {}),
    layout: TableLayoutType.FIXED,
    borders: allBorders(borderSpec || noBorder),
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: fill ? shade(fill) : undefined,
      margins: { top: BOX_PAD_TW, bottom: BOX_PAD_TW, left: 140, right: 140 },   // 위·아래 여백 같게(약 5pt)
      children: paragraphs,
    })] })],
  });
}
/** 페이지 나누기(---)는 빈 문단 대신 다음 문단의 pageBreakBefore로 처리 → 페이지 끝에 나누기가 걸려 빈 페이지가 생기는 일 방지 */
let PENDING_BREAK = false;
let LEVEL_WARN = [];
let LEVEL_TRACE = [];   // 구조 점검 경고
let CONFIDENTIAL = null;   // @confidential 문구(페이지 머리말)
const pb = () => { if (!PENDING_BREAK) return {}; PENDING_BREAK = false; return { pageBreakBefore: true }; };
/** 정확히 pt 높이만 차지하는 빈 문단(표·박스 앞뒤 간격용). Word는 auto 줄간격의 빈 문단을 서체 행높이만큼 키우므로 exact로 고정 */
const spacer = (pt = 6) => { const p = new Paragraph({ ...pb(), children: [new TextRun({ text: "", size: 2 })], spacing: { before: 0, after: 0, ...lineOf(pt) } }); p._spacer = true; return p; };

// ───────────────────────────── images ─────────────────────────────
/** JPEG/PNG 픽셀 크기 읽기 */
function imageSize(buf) {
  if (buf[0] === 0x89 && buf[1] === 0x50) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20), type: "png" };
  let i = 2;
  while (i < buf.length) {
    if (buf[i] !== 0xFF) { i++; continue; }
    const m = buf[i + 1];
    if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7), type: "jpg" };
    i += 2 + buf.readUInt16BE(i + 2);
  }
  return { w: 800, h: 600, type: "jpg" };
}
/** ![캡션](경로){w=120} → ImageRun. w는 mm(생략 시 본문 폭). 상대 경로는 입력 파일 기준 */
let INPUT_DIR = ".";
function imageRun(spec, maxMm) {
  const m = spec.match(/^!\[([^\]]*)\]\(([^)]+)\)(?:\{w=([\d.]+)\})?/);
  if (!m) return null;
  const file = path.isAbsolute(m[2]) ? m[2] : path.join(INPUT_DIR, m[2]);
  if (!fs.existsSync(file)) return new TextRun({ text: `[그림 없음: ${m[2]}]`, ...font(CONFIG.fonts.note), color: CONFIG.colors.red });
  const buf = fs.readFileSync(file); const sz = imageSize(buf);
  const wMm = Math.min(m[3] ? parseFloat(m[3]) : maxMm, maxMm);
  const wPx = wMm / 25.4 * 96, hPx = wPx * sz.h / sz.w;
  return new ImageRun({ data: buf, type: sz.type, transformation: { width: Math.round(wPx), height: Math.round(hPx) } });
}
function imageBlock(spec, indentMm = 0) {
  const run = imageRun(spec, CONTENT_W() / MM - indentMm);
  return [new Paragraph({ ...pb(), alignment: AlignmentType.CENTER, ...(indentMm ? { indent: { left: Math.round(indentMm * MM) } } : {}),
    spacing: { before: mmSp(4), after: mmSp(4), ...lineAuto(1) }, children: run ? [run] : [] })];
}

// ───────────────────────────── block builders ─────────────────────────────
function titleBand(text) {
  // 제목 두 줄: `# 첫 줄<br>둘째 줄` (원본처럼 의미 단위로 줄을 나눔)
  const parts = text.split("<br>");
  // 제목도 줄 넘침 보정(한 줄에 조금 넘치는 제목은 자간을 줄여 한 줄로)
  const kids = parts.flatMap((t, k) => { const lp = lineParts(t.trim()); const [tt, csX] = fitLine(lp.t, CONFIG.fonts.title, CONTENT_W() / 20 - 12, { forcedTw: lp.forcedTw, shrink: 0 });
    return [...(k ? [new TextRun({ break: 1 })] : []), ...inline(tt, CONFIG.fonts.title, { bold: true, ...csX })]; });
  return new Paragraph({
    ...pb(), children: kids,
    alignment: AlignmentType.CENTER,
    shading: shade(CONFIG.colors.titleFill),
    border: {
      top: border(CONFIG.colors.titleLine, 12), left: noBorder,
      bottom: border(CONFIG.colors.titleLine, 12), right: noBorder,
    },
    spacing: { before: 0, after: 0, ...lineOf(30) },
  });
}
function metaLine(text) {
  return new Paragraph({
    ...pb(), children: inline(text, CONFIG.fonts.meta), alignment: AlignmentType.RIGHT,
    spacing: { before: mmSp(1), after: mmSp(10), ...lineOf(17) },
  });
}
function confidentialLine(text) {
  return new Paragraph({
    ...pb(), children: inline(text, { name: "나눔고딕", size: 10 }, { color: CONFIG.colors.red }),
    alignment: AlignmentType.RIGHT, spacing: { before: 0, after: mmSp(4), ...lineOf(14) },
  });
}
function summaryBox(lines, variant = "peach") {
  const paras = lines.map((l, i) => {
    const sm = (() => { const q = l.match(/^(\S+)\s+(.*)$/s); return q && BOX_LEAD.test(q[1]) ? q : null; })();
    const hangMm = estWidthPt((sm ? sm[1] : "◈") + " ", CONFIG.fonts.body.size, 0, CONFIG.fonts.body.name) * 0.3528;
    const lp = lineParts(sm ? sm[2] : l);
    const [txt2, csX] = fitLine(lp.t, CONFIG.fonts.body, CONTENT_W() / 20 - 12 - (sm ? hangMm / 0.3528 : 0), { forcedTw: lp.forcedTw });
    return new Paragraph({
      alignment: JUSTIFY,
    ...(lp.tail && !tailFitsInline(lp.t, lp.tail, CONFIG.fonts.body, CONTENT_W() / 20 - 18) ? { tabStops: [{ type: "right", position: Math.round(CONTENT_W() - 240) }] } : {}),
    children: [...(sm ? [new TextRun({ text: textVS(sm[1]) + "\t", ...font(CONFIG.fonts.body) })] : []), ...inline(txt2, CONFIG.fonts.body, { shrinkParens: CONFIG.fit.shrinkParens, ...csX }), ...tailRuns(lp.tail, CONFIG.fonts.body, tailFitsInline(lp.t, lp.tail, CONFIG.fonts.body, CONTENT_W() / 20 - 18))],
    indent: sm ? { left: Math.round(hangMm * MM), hanging: Math.round(hangMm * MM) } : undefined,
    spacing: { before: i ? mmSp(4) : 0, after: 0, ...lineOf(CONFIG.lineExact.body) },
  }); });
  // 기본 연주황(10건). {white}: 흰 바탕 + 진파랑 테두리(2604 대외 문서)
  return [boxTable(paras, variant === "white"
    ? { borderSpec: border("475E90", 8), marginTw: 120 }
    : variant === "outline" ? { borderSpec: border("E8A060", 12), marginTw: 120 }   // 흰 바탕 + 주황 테두리
    : { fill: CONFIG.colors.summaryFill, borderSpec: border(CONFIG.colors.summaryBorder, 4), marginTw: 120 }), spacer(8)];
}
function plainHeader(text) {
  return new Paragraph({
    ...pb(), keepNext: true,
    children: inline(text, CONFIG.fonts.header, { bold: true, color: CONFIG.colors.headerText }),
    spacing: { before: mmSp(CONFIG.spacing.header), after: mmSp(CONFIG.spacing.headerAfter), ...lineOf(24) },
  });
}
/** ## Ⅰ. 검토 배경 → 연파랑 띠 + 위아래 파란 선, 왼쪽 정렬(2605·2608 설명자료형) */
function bandHeader(text) {
  const hf = { ...CONFIG.fonts.title, size: 20 };
  return [new Paragraph({
    ...pb(), keepNext: true,
    indent: { left: 0, right: Math.round(CONTENT_W() * 0.12) },
    shading: shade(CONFIG.colors.titleFill),
    border: { top: border(CONFIG.colors.titleLine, 12), left: noBorder, bottom: border(CONFIG.colors.titleLine, 12), right: noBorder },
    spacing: { before: mmSp(CONFIG.spacing.header), after: mmSp(CONFIG.spacing.headerAfter + 6), ...lineOf(32) },
    children: [new TextRun({ text: "\u00A0", ...font(hf) }), ...inline(text, hf, { bold: true })],
  })];
}
/** Ⅰ 보고 배경 / 1 검토 배경 → 파란 음영 번호 + 제목 + 밑줄(폭 62%). 표를 쓰지 않아 위아래 간격이 정확히 제어됨 */
function boxedNumberHeader(num, text) {
  // 번호 박스는 고정 폭 칸(원본 실측 폭 8~10mm, 높이 8~9mm)으로 그린다 — 예전처럼 글자 앞뒤 공백에 음영을 주면
  // Word에서 서체의 공백 폭만큼 박스가 넓어짐(HY 서체는 공백이 넓음)
  const hf = CONFIG.fonts.header;
  const need = estWidthPt(text, hf.size, 0, hf.name) * 20 + Math.round(22 * MM);
  const lineW = Math.min(CONTENT_W(), Math.max(Math.round(CONTENT_W() * 0.62), need));
  const c1 = Math.round(9 * MM), gap = Math.round(2.5 * MM), c3 = lineW - c1 - gap;
  const para = (runs, align) => new Paragraph({ alignment: align, keepNext: true, spacing: { before: 0, after: 0, ...lineOf(22) }, children: runs });
  const bottom = { top: noBorder, left: noBorder, right: noBorder, bottom: border("A6A6A6", 6) };
  const cell = (width, children, opts) => new TableCell({ width: { size: width, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    borders: opts.borders, shading: opts.fill ? shade(opts.fill) : undefined, margins: { top: 0, bottom: 0, left: opts.pad ?? 0, right: 0 }, children });
  return [spacer(CONFIG.spacing.header - 2), new Table({
    width: { size: lineW, type: WidthType.DXA }, columnWidths: [c1, gap, c3], layout: TableLayoutType.FIXED, borders: allBorders(noBorder),
    rows: [new TableRow({ cantSplit: true, height: { value: Math.round(24 * 20), rule: "exact" }, children: [
      cell(c1, [para([new TextRun({ text: num, ...font(hf), bold: true, color: "FFFFFF" })], AlignmentType.CENTER)], { fill: CONFIG.colors.romanFill, borders: allBorders(noBorder) }),
      cell(gap, [para([], AlignmentType.LEFT)], { borders: bottom }),
      cell(c3, [para(inline(text, hf, { bold: true }), AlignmentType.LEFT)], { borders: bottom }),
    ] })],
  })];   // 칸 안 문단의 '다음과 함께' 설정으로 헤더가 다음 내용과 떨어지지 않음
}
/** ### 1. 소제목 → 테두리 박스 소제목 ; ### (1안) … → 연청록 소제목 (문단 테두리·음영, 오른쪽 여백으로 폭 제어) */
function subHeader(text) {
  const tint = /^\((\d+안|[가-힣]{1,3})\)/.test(text);
  const f = CONFIG.fonts.subhdr;
  const estW = Math.min(CONTENT_W(), Math.round(estWidthPt(text, f.size) * 20 + 10 * MM));
  const b = tint ? { bottom: border("7FB8B8", 6) } : allBorders({ ...border("404040", 6), space: 2 });
  return [new Paragraph({
    ...pb(), keepNext: true,
    indent: { left: 0, right: Math.max(0, CONTENT_W() - estW) },
    border: b, shading: shade(tint ? CONFIG.colors.tintFill : CONFIG.colors.subhdrFill),
    spacing: { before: mmSp(Math.max(8, CONFIG.spacing.group - 1)), after: mmSp(4), ...lineOf(26) },   // 소제목 앞 간격은 @spacing을 따름
    children: [new TextRun({ text: "\u00A0", ...font(f) }), ...inline(text, f, { ...emph(f), ...(tint ? { color: CONFIG.colors.tintText } : {}) })],
  })];
}
/** 그라데이션 가로 막대(표지): 좁은 셀 여러 개에 진파랑→연파랑 채움 */
function gradientBar(heightPt = 7) {
  const n = 24, total = CONTENT_W(), w = Math.floor(total / n);
  const c0 = [0x33, 0x3C, 0xE6], c1 = [0xB4, 0xC4, 0xF2];
  const hex = (k) => c0.map((v, j) => Math.round(v + (c1[j] - v) * k / (n - 1)).toString(16).padStart(2, "0")).join("").toUpperCase();
  return new Table({
    width: { size: w * n, type: WidthType.DXA }, columnWidths: Array(n).fill(w), layout: TableLayoutType.FIXED,
    borders: { ...allBorders(noBorder), insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [new TableRow({ height: { value: Math.round(heightPt * 20), rule: "exact" }, children: Array.from({ length: n }, (_, k) => new TableCell({
      width: { size: w, type: WidthType.DXA }, shading: shade(hex(k)), borders: allBorders(noBorder),
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [new Paragraph({ spacing: { before: 0, after: 0, ...lineOf(1) }, children: [] })] })) })],
  });
}
/** ::: cover — 제목 줄들, @date, @org. 표지 한 페이지 후 자동 나누기 (2601·2605 설명자료형) */
function coverPage(lines) {
  const title = [], meta = {};
  for (const l of lines.map(x => x.trim()).filter(Boolean)) {
    const m = l.match(/^@(date|org)\s+(.*)$/); if (m) meta[m[1]] = m[2]; else title.push(l);
  }
  // 제목 줄마다 폭에 맞춰 크기 자동 축소(최대 34pt), 괄호는 -2pt
  const maxW = CONTENT_W() / 20 * 0.96;
  const sizeFor = (t) => { let z = 34; while (z > 20 && estWidthPt(t, z, 4, "HY") > maxW) z -= 1; return z; };
  const center = (runs, before, after, line) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: mmSp(before), after: mmSp(after), ...lineOf(line) }, children: runs });
  const out = [spacer(160), gradientBar(), ...title.map((t, i) => { const z = Math.min(...title.map(sizeFor)); return center(inline(t, { ...CONFIG.fonts.title, size: z }, { bold: true, shrinkParens: 4 }), i ? 0 : 8, i === title.length - 1 ? 8 : 0, z * 1.45); }), gradientBar()];
  if (meta.date) out.push(center(inline(meta.date, { ...CONFIG.fonts.title, size: 22 }, { bold: true }), 120, 0, 30));
  if (meta.org) out.push(meta.org.startsWith("![") ? new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: mmSp(150), after: 0 }, children: [imageRun(meta.org, 60)].filter(Boolean) })
                                             : center(inline(meta.org, { ...CONFIG.fonts.title, size: 24 }, { bold: true }), 170, 0, 32));
  PENDING_BREAK = true;
  return out;
}
/** ::: toc — "항목 | 쪽" 줄. 앞 공백 2칸 이상이면 하위 항목, [첨부]·(첨부)로 시작하면 첨부 항목. 점선 리더 */
function tocPage(lines) {
  const top = spacer(24);            // 앞 페이지(표지)의 나누기를 여기서 소비
  const inner = Math.round(CONTENT_W() - 24 * MM);
  const paras = [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: mmSp(26), ...lineOf(30) },
    children: [new TextRun({ text: "목   차", ...font({ ...CONFIG.fonts.title, size: 20 }), bold: true })] })];
  for (const raw of lines) {
    if (!raw.trim()) continue;
    const sub = /^\s{2,}/.test(raw), [label, pg = ""] = raw.trim().split("|").map(x => x.trim());
    const att = /^[\[(]\s*(첨부|참고|별첨)/.test(label);
    const f = att ? { name: "나눔고딕", size: 13 } : (sub ? { ...CONFIG.fonts.title, size: 15 } : { ...CONFIG.fonts.title, size: 18 });
    paras.push(new Paragraph({
      indent: { left: Math.round((sub ? 8 : (att ? 6 : 2)) * MM) },
      tabStops: [{ type: "right", position: inner, leader: "dot" }],
      spacing: { before: mmSp(att ? 14 : (sub ? 20 : 48)), after: 0, ...lineOf(f.size * 1.5) },
      children: [...inline(label, f, { bold: !att }), new TextRun({ text: "\t" + pg, ...font(f), bold: !att })],
    }));
  }
  const w = CONTENT_W() - Math.round(10 * MM);
  const tbl = new Table({
    width: { size: w, type: WidthType.DXA }, columnWidths: [w], layout: TableLayoutType.FIXED,
    indent: { size: Math.round(5 * MM), type: WidthType.DXA },
    borders: allBorders(border("7F7F7F", 18)),
    rows: [new TableRow({ height: { value: Math.round(232 * MM), rule: "atLeast" }, children: [new TableCell({
      width: { size: w, type: WidthType.DXA }, borders: allBorders(border("7F7F7F", 18)),
      margins: { top: 300, bottom: 300, left: 360, right: 360 }, children: paras })] })],
  });
  PENDING_BREAK = true;              // 목차 다음은 새 페이지
  return [top, tbl];
}
/** ::: cols 40 — 좌우 배치(표+그림 등). 왼쪽 폭 %, 두 블록은 ||| 로 구분. 테두리 없음 */
function colsBlock(pct, lines, pctx = {}) {
  const k = lines.findIndex(l => l.trim() === "|||");
  const L = k < 0 ? lines : lines.slice(0, k), R = k < 0 ? [] : lines.slice(k + 1);
  const total = pctx.inBox ? Math.round(CONTENT_W() - 300) : CONTENT_W(), gap = Math.round(3 * MM);
  const wl = Math.round((total - gap) * pct / 100), wr = total - gap - wl;
  const saved = WIDTH_OVERRIDE, savedPB = PENDING_BREAK;
  PENDING_BREAK = false;
  WIDTH_OVERRIDE = wl; const lb = parseBlocks(L, { ...pctx, inCol: true });
  WIDTH_OVERRIDE = wr; const rb = parseBlocks(R, { ...pctx, inCol: true });
  WIDTH_OVERRIDE = saved; PENDING_BREAK = savedPB;
  const cell = (w, ch) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: allBorders(noBorder),
    margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP,
    children: ch.length ? ch : [new Paragraph({ children: [] })] });
  return [spacer(4), new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: [wl, gap, wr], layout: TableLayoutType.FIXED,
    borders: { ...allBorders(noBorder), insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [new TableRow({ cantSplit: true, children: [cell(wl, lb), cell(gap, []), cell(wr, rb)] })],
  }), spacer(4)];
}
/** ## 첨부 1 제목 / 참고 2 제목 / 별첨 제목 → [진파랑 라벨 박스] [테두리 제목 박스](원본 관행) */
function attachHeader(label, text) {
  const hf = CONFIG.fonts.header;
  const w = CONTENT_W(), c1 = Math.round(20 * MM), gap = Math.round(3 * MM), c3 = w - c1 - gap;
  const cell = (width, children, opts) => new TableCell({
    width: { size: width, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    borders: opts.borders, shading: opts.fill ? shade(opts.fill) : undefined,
    margins: { top: 30, bottom: 30, left: opts.pad || 60, right: 60 }, children,
  });
  const para = (runs, align) => new Paragraph({ alignment: align, keepNext: true, spacing: { before: 0, after: 0, ...lineOf(22) }, children: runs });
  return [spacer(6), new Table({
    width: { size: w, type: WidthType.DXA }, columnWidths: [c1, gap, c3], layout: TableLayoutType.FIXED,
    borders: allBorders(noBorder),
    rows: [new TableRow({ height: { value: Math.round(26 * 20), rule: "exact" }, children: [
      cell(c1, [para([new TextRun({ text: label, ...font({ ...hf, size: 15 }), bold: true, color: "FFFFFF" })], AlignmentType.CENTER)],
        { fill: CONFIG.colors.attachFill, borders: allBorders(noBorder) }),
      cell(gap, [para([], AlignmentType.CENTER)], { borders: allBorders(noBorder) }),
      cell(c3, [para(inline(text, { ...hf, size: 15 }, { bold: true }), AlignmentType.LEFT)],
        { borders: allBorders(border(CONFIG.colors.tableOuter, 6)), pad: 160 }),
    ] })],
  }), spacer(10)];
}
function caption(text, first = false) {
  return new Paragraph({
    ...pb(), keepNext: true,
    children: inline(text, CONFIG.fonts.caption, emph(CONFIG.fonts.caption)), alignment: AlignmentType.CENTER,
    spacing: { before: first ? 0 : mmSp(8), after: mmSp(2), ...lineOf(18) },
  });
}
function conclusionBox(lines, variant = "white") {
  const fixed = lines.map((l, i) => {
    const m = l.match(/^(\S+)\s+(.*)$/s);
    const marker = m && BOX_LEAD.test(m[1]) ? m[1] : null, rest = marker ? m[2] : l;
    const f = CONFIG.fonts.conclusion;
    const hangPt = marker ? estWidthPt(marker + " ", f.size, 0, f.name) + f.size * 0.15 : 0;
    const lp = lineParts(rest);
    const [rest2, csX] = fitLine(lp.t, f, CONTENT_W() / 20 - 13 - hangPt, { forcedTw: lp.forcedTw });   // 박스 안쪽 폭(셀 여백 130twips×2) − 기호 폭
    return new Paragraph({
      alignment: JUSTIFY,
      ...(lp.tail && !tailFitsInline(lp.t, lp.tail, f, CONTENT_W() / 20 - 13 - hangPt) ? { tabStops: [{ type: "right", position: Math.round(CONTENT_W() - 260) }] } : {}),
      children: [...(marker ? [new TextRun({ text: textVS(marker) + "\t", ...font(f), ...emph(f) })] : []), ...runsWithKeyword(rest2, f, { ...emph(f), ...csX }), ...tailRuns(lp.tail, f, tailFitsInline(lp.t, lp.tail, f, CONTENT_W() / 20 - 13 - hangPt))],
      indent: marker ? { left: Math.round(hangPt * 0.3528 * MM), hanging: Math.round(hangPt * 0.3528 * MM) } : undefined,   // 둘째 줄 = 기호 다음 글자 시작점
      spacing: { before: i ? mmSp(3) : 0, after: 0, ...lineOf(CONFIG.lineExact.body) },
    });
  });
  // 기본 흰 바탕+점선. {green} 연두+점선(2506), {yellow} 연노랑+실선(2601·2605·2505)
  const fill = variant === "green" ? "EBF7E1" : (variant === "yellow" ? "FDF3DC" : undefined);
  const bspec = variant === "yellow" ? border("404040", 8) : border(CONFIG.colors.conclusionBorder, 8, BorderStyle.DOTTED);
  return [spacer(10), boxTable(fixed, { fill, borderSpec: bspec, marginTw: 130 }), spacer(8)];
}
function frameBox(lines, variant = "green") {
  const f = CONFIG.fonts.conclusion;
  const fill = variant === "peach" ? CONFIG.colors.summaryFill : (variant === "yellow" ? "FFF8DC" : (variant === "gray" ? CONFIG.colors.grayFill : (variant === "white" ? undefined : CONFIG.colors.frameFill)));
  const bspec = variant === "green" ? noBorder : border(CONFIG.colors.grayBorder, 6, BorderStyle.DOTTED);
  const paras = lines.map((l, i) => {
    const m = (() => { const q = l.match(/^(\S+)\s+(.*)$/s); return q && BOX_LEAD.test(q[1]) ? q : null; })();
    const lp = lineParts(m ? m[2] : l);
    const [body2, csX] = fitLine(lp.t, f, CONTENT_W() / 20 - 9 - (m ? estWidthPt("◈ ", f.size, 0, f.name) : 0), { forcedTw: lp.forcedTw });
    return new Paragraph({
      alignment: JUSTIFY,
      ...(lp.tail && !tailFitsInline(lp.t, lp.tail, f, CONTENT_W() / 20 - 15) ? { tabStops: [{ type: "right", position: Math.round(CONTENT_W() - 180) }] } : {}),
      children: [...(m ? [new TextRun({ text: textVS(m[1]) + "\t", ...font(f), ...emph(f) })] : []), ...runsWithKeyword(body2, f, { ...emph(f), ...csX }), ...tailRuns(lp.tail, f, tailFitsInline(lp.t, lp.tail, f, CONTENT_W() / 20 - 15))],
      indent: m ? { left: Math.round(estWidthPt(m[1] + " ", f.size, 0, f.name) * 0.3528 * MM), hanging: Math.round(estWidthPt(m[1] + " ", f.size, 0, f.name) * 0.3528 * MM) } : undefined,   // 기호 없는 줄은 들여쓰기 없음
      spacing: { before: i ? mmSp(3) : 0, after: 0, ...lineOf(CONFIG.lineExact.body) },
    });
  });
  return [spacer(4), boxTable(paras, { fill, borderSpec: bspec, marginTw: 90 }), spacer(4)];
}
function grayBox(title, lines, tableFont, solid = false, white = false, tightB = false, yellowB = false, hlB = false, nested = false, boxFsArg = null) {
  const paras = [];
  if (title) paras.push(new Paragraph({
    children: (() => { const lpT = lineParts(title); const [tt, csX] = fitLine(lpT.t.startsWith("※") ? lpT.t : `※ (참고) ${lpT.t}`, CONFIG.fonts.boxTitle, CONTENT_W() / 20 - 14, { forcedTw: lpT.forcedTw }); return inline(tt, CONFIG.fonts.boxTitle, { ...emph(CONFIG.fonts.boxTitle), shrinkParens: CONFIG.fit.shrinkParens, ...csX }); })(),
    spacing: { before: 0, after: mmSp(3), ...lineOf(CONFIG.lineExact.box) },
  }));
  // 박스 안에 표가 올 수도 있음
  const blocks = parseBlocks(lines, { inBox: true, tableFont, tightBox: tightB, boxHl: hlB, boxTitled: !!title, boxFs: boxFsArg });   // 박스 안 제목 다음 첫 문단은 4pt
  paras.push(...blocks);
  // [ledger D2] 회색 박스 테두리는 점선이 다수(14건 vs 7건) → 기본 점선, {solid}로 실선
  return [spacer(6), boxTable(paras, { fill: yellowB ? "FFF8DC" : (white ? undefined : CONFIG.colors.grayFill), borderSpec: border(white ? "404040" : CONFIG.colors.grayBorder, 6, solid ? BorderStyle.SINGLE : BorderStyle.DOTTED), marginTw: 110, ...(nested ? { widthDxa: Math.round(CONTENT_W() - 300) } : {}) }), spacer(4)];
}

/** 마크다운 파이프 표 → Table
 *  - 첫 행(또는 @header n 행) 헤더 연노랑. @kv: 홀수열 라벨. @firstcol: 첫 열 라벨. @dotted: 안쪽 가로선 점선
 *  - 셀 "<" : 왼쪽 셀과 가로 병합, "^" : 위 셀과 세로 병합
 *  - 행 첫 셀 접두 {g}(연두 강조행) {gray}(합계행) {y}(노랑) {b}(굵게) {blue}(파란 굵게)
 *  - 정렬: 순수 숫자 → 오른콽, 14자 이하 → 가운데, 그 외·항목형 → 왼쪽. 괄호 안 -2pt
 *  - 테두리: 위·아래 굵은 검정, 좌·우 없음, 안쪽 가는 선(원본 관행)
 */
function mdTable(rows, fontSize, { inBox = false, kv = false, headerRows = 1, firstCol = false, dotted = false, blueHdr = false, compact = false, rowH = 0, indentMm = 0, equal = false, center = false, rowPad = false, dense = false, colAlignSpec = null } = {}) {
  const tight = compact || kv;   // 밀집 표: @compact, 라벨-값 표(원본 회사개요표는 촘촘). @dotted는 선 모양만 바꿈
  const roomy = !!rowPad;   // @pad: 여유 있는 셀 여백(명단·내역 표, 원본 한 줄 행 약 23pt)
  const raw = rows.map(r => r.replace(/^\||\|$/g, "").split("|").map(c => c.trim()));
  const ncol = Math.max(...raw.map(r => r.length));
  raw.forEach(r => { while (r.length < ncol) r.push(""); });
  // 행 옵션
  const rowOpt = raw.map(r => {
    const o = { fill: null, bold: false, blue: false };
    let c0 = r[0];
    let m; while ((m = c0.match(/^\{(g|gray|y|b|blue)\}/))) {
      if (m[1] === "g") o.fill = CONFIG.colors.rowGreen; else if (m[1] === "gray") { o.fill = CONFIG.colors.rowGray; o.bold = true; }
      else if (m[1] === "y") o.fill = CONFIG.colors.rowYellow; else if (m[1] === "b") o.bold = true; else if (m[1] === "blue") { o.blue = true; o.bold = true; }
      c0 = c0.slice(m[0].length).trim();
    }
    r[0] = c0; return o;
  });
  const plain = (t) => t.replace(/^(\{(bg|fg):[^}]*\}|\{(g|gray|y|b|blue)\})+/g, "").replace(/!\[[^\]]*\]\([^)]*\)(\{w=[\d.]+\})?/g, "").replace(/\*\*|__|\{\{|\}\}|\^\^|~~|!!/g, "");   // 폭 계산용: 셀 접두·그림·인라인 기호 제거
  // 열 너비: 병합 표시(<,^) 제외한 최대 글자폭 비례
  // 폭 계산에서 병합 표시 셀과 '가로 병합의 시작 셀'(다음 셀이 <)은 제외 — 긴 병합 제목이 첫 열 폭을 부풀리지 않게
  const cellTexts = (i) => raw.map(r => (r[i] === "<" || r[i] === "^" || r[i + 1] === "<") ? "" : plain(r[i]));
  const lens = Array.from({ length: ncol }, (_, i) => Math.max(3, ...cellTexts(i).map(x => Math.max(0, ...x.split("<br>").map(y => estWidthPt(y, 10) / 10)))));
  // 가로 병합 셀(다음 셀이 <)의 필요 폭을 병합된 열들에 나눠 반영 — 병합 헤더가 줄바꿈되지 않게
  raw.forEach(r => r.forEach((t, i) => {
    if (t === "<" || r[i + 1] !== "<") return;
    let span = 1; while (i + span < ncol && r[i + span] === "<") span++;
    const need = Math.max(...plain(t).split("<br>").map(y => estWidthPt(y, 10) / 10)) + 0.8;
    const have = lens.slice(i, i + span).reduce((a, b) => a + b, 0);
    if (need > have) for (let k = i; k < i + span; k++) lens[k] += (need - have) / span;
  }));
  // 최소 폭: 열에서 가장 긴 '끊을 수 없는 토큰'(공백 기준, 12자 한도) — 숫자·짧은 라벨 열이 지나치게 좁아지는 것 방지
  const mins = Array.from({ length: ncol }, (_, i) => Math.max(3, ...cellTexts(i).flatMap(x => x.split(/<br>|\s+/)).map(tok => Math.min(12, estWidthPt(tok, 10) / 10))));
  const total = Math.round(CONTENT_W() - (inBox ? 300 : 0) - indentMm * MM);
  const fs = (fontSize || CONFIG.fonts.table.size);
  const minW = mins.map(m => Math.round(m * fs * 20 + 160));       // 글자폭(em→pt→twips) + 셀 여백
  // 셀 안 그림 {w=mm}은 그 폭을 최소 폭으로 보장
  raw.forEach(r => r.forEach((t, i) => { const m = t.match(/!\[[^\]]*\]\([^)]*\)\{w=([\d.]+)\}/); if (m) minW[i] = Math.max(minW[i], Math.round((parseFloat(m[1]) + 4) * MM)); }));
  let rest = total - minW.reduce((a, b) => a + b, 0);
  let widths;
  if (rest <= 0) { const sum = minW.reduce((a, b) => a + b, 0); widths = minW.map(w => Math.round(total * w / sum)); }
  else {
    // 남는 폭의 15%는 모든 열에 균등, 85%는 내용 길이 초과분 비례 → 긴 열이 전부 흡수하지 않게(원본 표의 숫자 열 여유)
    const extra = lens.map((l, i) => Math.max(0, Math.min(l, 40) - mins[i]));
    const esum = extra.reduce((a, b) => a + b, 0) || 1;
    widths = minW.map((w, i) => w + Math.round(rest * 0.15 / ncol) + Math.round(rest * 0.85 * extra[i] / esum));
  }
  if (equal) { const w0 = Math.floor(total / ncol); widths = Array(ncol).fill(w0); }
  widths[widths.length - 1] += total - widths.reduce((a, b) => a + b, 0);
  const tf = { ...CONFIG.fonts.table, size: fontSize || CONFIG.fonts.table.size };
  const thin = border(CONFIG.colors.tableBorder, 4), thick = border(CONFIG.colors.tableOuter, 12);
  const innerH = dotted ? border(CONFIG.colors.tableBorder, 4, BorderStyle.DOTTED) : thin;
  const nrow = raw.length;
  const isNumeric = (t) => /^[\d.,%△▲▽▼+\-\s]+$/.test(t) && /\d/.test(t);

  // 셀 하나의 자연 정렬: 목록형·긴 문장 → 왼쪽, 금액·주식수 → 오른쪽, 그 외 → 가운데
  const stripPrefix = (t) => t.replace(/^(\{(bg|fg):[^}]*\})+/, "").trim();
  const naturalAlign = (t) => {
    const ss = stripPrefix(t).split("<br>").map(x => x.trim());
    if (ss.some(x => /^[■ㅇ·•∙▪‣①②③④⑤➊➋➌]/.test(x) || /^-\s+\S/.test(x)) || ss.some(x => plain(x).length > 26)) return AlignmentType.LEFT;
    if (ss.every(x => x === "" || (isNumeric(plain(x)) && (plain(x).length >= 5 || /[,.]/.test(plain(x)))))) return AlignmentType.RIGHT;
    return AlignmentType.CENTER;
  };
  // 열 단위 정렬: 본문 셀 중 하나라도 왼쪽이면 열 전체 왼쪽, 전부 오른쪽이면 오른쪽, 아니면 가운데
  const colAlign = Array.from({ length: ncol }, (_, ci) => {
    const al = raw.slice(headerRows).map(r => r[ci]).filter((t, k) => t && t !== "<" && t !== "^" && raw[headerRows + k][ci + 1] !== "<" && !t.startsWith("![")).map(naturalAlign);
    if (colAlignSpec && colAlignSpec[ci]) return ({ l: AlignmentType.LEFT, r: AlignmentType.RIGHT, c: AlignmentType.CENTER })[colAlignSpec[ci]] || AlignmentType.CENTER;   // @align
    if (center) return AlignmentType.CENTER;   // @center: 긴 문장이어도 가운데(목록형 셀은 아래에서 왼쪽 유지)
    if (al.includes(AlignmentType.LEFT)) return AlignmentType.LEFT;
    if (al.length && al.every(a => a === AlignmentType.RIGHT)) return AlignmentType.RIGHT;
    return AlignmentType.CENTER;
  });
  // 같은 층위의 형제 열(비교표의 대상 열: 美DTCCㆍ일본 ProgmatㆍProject Helvetia 등)은 정렬을 하나로(작성자 확인)
  //   형제 열 = 첫 열(구분)을 뺀 2개 이상의 열로, 내용 길이가 서로 비슷한 경우(평균 길이 차 2.5배 이내, 숫자 열 제외)
  //   정렬: 형제 열 중 하나라도 목록형(■·ㅇ·- 등으로 시작하는 셀)이면 모두 왼쪽, 아니면 모두 가운데(긴 문장은 가운데로 줄바꿈)
  if (!colAlignSpec && !center && ncol >= 3) {
    const body = raw.slice(headerRows);
    const cellsOf = ci => body.map(r => r[ci]).filter(t => t && t !== "<" && t !== "^" && !t.startsWith("!["));
    const plainLen = t => t.replace(/<br>/g, " ").replace(/\{[^}]*\}|\*\*|\^\^|__|~~/g, "").length;
    const sib = []; for (let ci = 1; ci < ncol; ci++) { const cs = cellsOf(ci); if (cs.length && colAlign[ci] !== AlignmentType.RIGHT) sib.push(ci); }
    const avg = ci => { const cs = cellsOf(ci); return cs.reduce((a, t) => a + plainLen(t), 0) / Math.max(1, cs.length); };
    if (sib.length >= 2) {
      const avgs = sib.map(avg), mx = Math.max(...avgs), mn = Math.max(1, Math.min(...avgs));
      if (mx / mn <= 2.5) {
        const listy = sib.some(ci => cellsOf(ci).some(t => t.split("<br>").some(x => /^[■ㅇ·•∙▪‣▶➊-➓①-⑳❶-❿]/.test(x.trim()) || /^-\s+\S/.test(x.trim()))));
        for (const ci of sib) colAlign[ci] = listy ? AlignmentType.LEFT : AlignmentType.CENTER;
      }
    }
  }
  const mkCell = (text, ri, ci, span, vmerge) => {
    const isHeader = ri < headerRows, isLabel = (kv && ci % 2 === 0) || (firstCol && ci === 0);
    const ro = rowOpt[ri];
    // 셀 접두 {bg:HEX} 채움, {fg:HEX|white} 글자색 (컬러 다이어그램형 표)
    let cellBg = null, cellFg = null, cm;
    while ((cm = text.match(/^\{(bg|fg):([0-9A-Fa-f]{6}|white|black)\}/))) {
      const v = cm[2] === "white" ? "FFFFFF" : (cm[2] === "black" ? "000000" : cm[2].toUpperCase());
      if (cm[1] === "bg") cellBg = v; else cellFg = v;
      text = text.slice(cm[0].length).trim();
    }
    const subs = text.split("<br>").map(x => x.trim());
    const isList = subs.some(x => /^[■ㅇ·•∙▪‣①②③④⑤➊➋➌]/.test(x) || /^-\s+\S/.test(x));
    let align = AlignmentType.CENTER;
    if (!isHeader && !isLabel) align = span > 1 ? naturalAlign(text) : colAlign[ci];   // 같은 열(같은 층위)의 셀은 정렬 통일(작성자 확인)
    if (center && !isHeader && !isLabel && isList) align = AlignmentType.LEFT;
    const w = widths.slice(ci, ci + span).reduce((a, b) => a + b, 0);
    const bd = { top: ri === 0 ? thick : (ri === headerRows && headerRows > 0 ? thin : innerH), left: ci === 0 ? noBorder : thin,
                 bottom: ri === nrow - 1 ? thick : (ri < headerRows ? thin : innerH), right: ci + span >= ncol ? noBorder : thin };
    const hdrFill = blueHdr ? CONFIG.colors.tableHeaderBlue : CONFIG.colors.tableHeader;
    const fill = cellBg || (isHeader || isLabel ? hdrFill : (ro.fill || undefined));
    const extra = { bold: isHeader || isLabel || ro.bold || !!cellBg, ...(cellFg ? { color: cellFg } : (ro.blue ? { color: CONFIG.colors.blue } : {})), shrinkParens: 2 };
    return new TableCell({
      width: { size: w, type: WidthType.DXA }, borders: bd, columnSpan: span > 1 ? span : undefined,
      verticalMerge: vmerge, shading: fill ? shade(fill) : undefined,
      verticalAlign: VerticalAlign.CENTER, margins: { top: tight ? 10 : (roomy ? 90 : 45), bottom: tight ? 10 : (roomy ? 90 : 45), left: ncol >= 8 ? 20 : 50, right: ncol >= 8 ? 20 : 50 },  // 점선(밀집) 표 13.5pt 행, 일반 표는 여유 있게
      children: (() => { const st = []; return subs.map(x => { const mm = x.match(/^([■□▪‣▶◆◈ㅇ·•∙➊-➓①-⑳❶-❿]|-(?=\s))\s/); if (!mm || !LEVEL_MARKERS.test(mm[1])) return 0;
        const fam = markerFamily(mm[1]); const k = st.indexOf(fam); if (k >= 0) { st.length = k + 1; return k + 1; } st.push(fam); return st.length; }); })()
        .map((lv, xi) => [subs[xi], lv]).map(([x, cellLv], xi) => x.startsWith("![")
        ? new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0, ...lineAuto(1) }, children: [imageRun(x, (w - 120) / MM)].filter(Boolean) })
        : new Paragraph({
        alignment: align,
        indent: align === AlignmentType.LEFT ? (cellIndent(x, tf.size, (dotted || dense) ? 1 : cellLv) || (() => { for (let q = xi - 1; q >= 0; q--) { if (subs[q].startsWith("![")) return undefined; const ci = cellIndent(subs[q], tf.size, 1); if (ci) return { left: ci.left, hanging: 0 }; } return undefined; })()) : undefined,   // 기호 없이 이어지는 줄(<br>로 나눈 줄)은 앞 기호 줄의 글자 시작점   // 약력·조문처럼 촘촘한 표는 원본도 층위 들여쓰기 없음   // 셀 안 층위(■ → ㅇ) 들여쓰기 + 기호 다음 글자 시작점 내어쓰기
        // 셀 안 줄 간격(코퍼스 중앙값 1.4배, 밀집 표 1.25배), 목록 항목 사이 여백(■·ㅇ·- 등으로 시작하는 둘째 줄부터)
        spacing: { before: (xi > 0 && isList && !dotted && !dense) ? (/^[■▪‣▶◆◈➊-➓①-⑳❶-❿]/.test(x) ? mmSp(CONFIG.cell.gapTop) : (/^[ㅇ·•∙]|^-\s/.test(x) ? mmSp(CONFIG.cell.gap) : 0)) : 0, after: 0, ...lineOf(Math.round(tf.size * (tight ? 1.15 : CONFIG.cell.line) * 2) / 2) },   // 밀집 표(주주 구성 등)는 원본도 촘촘(약 1.15~1.2배)
        // 셀 안 긴 줄도 본문과 같은 줄 넘침 보정(줄바꿈 기회 + 자동 자간), 왼쪽 정렬 셀만
        ...((lineParts(x).tail && !tailFitsInline(lineParts(x).t, lineParts(x).tail, tf, (w - 200) / 20)) ? { tabStops: [{ type: "right", position: Math.max(200, w - 220) }] } : {}),
        children: (() => { const lp = lineParts(x); const hg = align === AlignmentType.LEFT ? cellHang(lp.t, tf.size) : null;
          const avail = (w - 200) / 20 - (hg ? hg.left / 20 : 0);
          const [x2, csX] = fitLine(lp.t, tf, avail, { forcedTw: lp.forcedTw });
          return [...inline(x2, tf, { ...extra, ...csX }), ...tailRuns(lp.tail, tf, tailFitsInline(lp.t, lp.tail, tf, (w - 200) / 20))]; })(),   // 표 셀도 본문과 같은 줄바꿈 규칙(정렬 무관)
      })),
    });
  };
  const trows = raw.map((r, ri) => {
    const cells = [];
    for (let ci = 0; ci < ncol; ci++) {
      const t = r[ci];
      if (t === "<") continue;                       // 왼쪽 셀에 병합됨
      let span = 1; while (ci + span < ncol && r[ci + span] === "<") span++;
      let vmerge;
      if (t === "^") vmerge = VerticalMergeType.CONTINUE;
      else if (ri + 1 < nrow && raw[ri + 1][ci] === "^") vmerge = VerticalMergeType.RESTART;
      cells.push(mkCell(t === "^" ? "" : t, ri, ci, span, vmerge));
    }
    return new TableRow({ tableHeader: !kv && ri < headerRows, cantSplit: true, children: cells, ...(rowH ? { height: { value: Math.round(rowH * 20), rule: "atLeast" } } : {}) });
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
    ...(indentMm ? { indent: { size: Math.round(indentMm * MM), type: WidthType.DXA } } : {}),
    borders: { top: thick, left: noBorder, bottom: thick, right: noBorder, insideHorizontal: innerH, insideVertical: thin },
    rows: trows,
  });
}

/** ::: timeline  →  "헤더 :: 본문<br>본문" 줄들을 화살표로 연결한 박스열 */
function timeline(lines, fontSize) {
  const items = lines.filter(l => l.trim()).map(l => { const [h, b = ""] = l.split("::"); return { h: h.trim(), b: b.trim() }; });
  const n = items.length, total = CONTENT_W();
  const arrowW = Math.round(8 * MM), itemW = Math.round((total - arrowW * (n - 1)) / n);
  const widths = []; items.forEach((_, i) => { widths.push(itemW); if (i < n - 1) widths.push(arrowW); });
  widths[widths.length - 1] += total - widths.reduce((a, b) => a + b, 0);
  const tf = { ...CONFIG.fonts.table, size: fontSize || CONFIG.fonts.table.size };
  const bd = border(CONFIG.colors.tableBorder, 4);
  const cell = (w, children, opts = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: allBorders(opts.border || noBorder),
    shading: opts.fill ? shade(opts.fill) : undefined, verticalAlign: VerticalAlign.CENTER,
    margins: { top: 50, bottom: 50, left: 70, right: 70 }, children,
  });
  const row = (key, isHead) => new TableRow({ children: items.flatMap((it, i) => {
    const subs = it[key].split("<br>").map(s => s.trim()).filter(Boolean);
    const paras = subs.map(s => new Paragraph({
      alignment: isHead || subs.length === 1 ? AlignmentType.CENTER : AlignmentType.LEFT,
      indent: (isHead || subs.length === 1) ? undefined : cellHang(s, tf.size),
      spacing: { before: 0, after: 0, line: 276 },
      ...(lineParts(s).tail && !tailFitsInline(lineParts(s).t, lineParts(s).tail, tf, (itemW - 140) / 20) ? { tabStops: [{ type: "right", position: Math.max(200, itemW - 150) }] } : {}),
      children: (() => { const lp = lineParts(s); const [s2, csX] = fitLine(lp.t, tf, (itemW - 140) / 20, { forcedTw: lp.forcedTw });
        return [...inline(s2, tf, { bold: isHead, shrinkParens: isHead ? 0 : CONFIG.fit.shrinkParens, ...csX }), ...tailRuns(lp.tail, tf, tailFitsInline(lp.t, lp.tail, tf, (itemW - 140) / 20))]; })(),
    }));
    const c = cell(itemW, paras.length ? paras : [new Paragraph({ children: [] })], { border: bd, fill: isHead ? CONFIG.colors.tableHeader : undefined });
    if (i === n - 1) return [c];
    const arrow = cell(arrowW, [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0, line: 276 },
      children: isHead ? [] : [new TextRun({ text: "⇨", ...font({ ...tf, size: 14 }) })] })]);
    return [c, arrow];
  }) });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
    borders: allBorders(noBorder), rows: [row("h", true), row("b", false)] });
}

// ───────────────────────────── parser ─────────────────────────────
const variantOf = (arr, def) => { let v = def; for (let k = 0; k < arr.length; k++) { const m = arr[k].match(/\s*\{(green|peach|white|yellow|gray|outline)\}\s*$/); if (m) { v = m[1]; arr[k] = arr[k].slice(0, m.index); } } return v; };
function parseBlocks(lines, ctx = {}) {
  const out = [];
  let i = 0, tableFont = ctx.tableFont, kvNext = false, headerRows = 1, firstColNext = false, dottedNext = false, blueNext = false, compactNext = false, rowHNext = 0, equalNext = false, centerNext = false, padNext = false, denseNext = false, alignNext = null, prevWasTable = false, prevWasNote = false, prevWasBox = false, noteLeft = null, lastTableNote = null, underBox = false, circleTop = false, indentNext = 0, lvStack = [], prevLevel = 0;
  // 줄 앞 공백으로 층위를 직접 표시한 블록: 공백 표시가 일관될 때만(하위 기호 ㅇ·- 등이 모두 들여써진 경우) 직접 표시로 본다.
  // 일부 줄만 공백이 있으면(원문 붙여넣기 등) 섞인 것으로 보고 문맥 추정을 쓴다 — 섞인 채로 직접 표시로 보면 공백 없는 ㅇ가 1층위로 올라감
  const lvLines = lines.filter(l => /^\s*(□|ㅇ|-|·|∙|■|▶|‣|▪|◆|[①-⑨]|[➊-➑]|[❶-❻])\s/.test(l));
  const indented = lvLines.filter(l => /^ {2,}/.test(l));
  const subUnindented = lvLines.filter(l => /^(ㅇ|-|·|∙)\s/.test(l));
  const explicitLv = indented.length > 0 && subUnindented.length === 0;
  while (i < lines.length) {
    const raw = lines[i], line = raw.trimEnd();
    const t = line.trim();
    if (!t) { i++; continue; }
    if (prevWasTable && !t.startsWith("|") && !/^(\*|※|\d\)|[■‣▶▪])/.test(t) && !t.startsWith("@")) { out.push(spacer(4)); prevWasTable = false; }
    if (prevWasBox && !t.startsWith(":::") && !NOTE_MARKERS.has(t.split(/\s+/)[0]) && !t.startsWith("@")) prevWasBox = false;

    if (t.startsWith("@tablefont")) { tableFont = parseInt(t.split(/\s+/)[1], 10); i++; continue; }
    if (t.startsWith("@spacing")) { const a = t.split(/\s+/).slice(1).map(Number); if (a[0]) CONFIG.spacing.group = a[0]; if (a[1]) CONFIG.spacing.inner = a[1]; if (a[2]) CONFIG.spacing.note = a[2]; i++; continue; }
    if (t.startsWith("@box ")) { const a = t.split(/\s+/).slice(1).map(Number); const R = CONFIG.boxRange;
      if (a[0]) CONFIG.lineExact.box = Math.max(R.line[0], Math.min(R.line[1], a[0])); if (a[1] != null && !isNaN(a[1])) R.gapOverride = Math.max(R.gap[0], Math.min(R.gap[1], a[1])); i++; continue; }
    if (t.startsWith("@cell ")) { const a = t.split(/\s+/).slice(1).map(Number); const C = CONFIG.cell;
      if (a[0]) C.line = Math.max(C.lineRange[0], Math.min(C.lineRange[1], a[0])); if (a[1] != null && !isNaN(a[1])) C.gap = Math.max(C.gapRange[0], Math.min(C.gapRange[1], a[1]));
      if (a[2] != null && !isNaN(a[2])) C.gapTop = Math.max(C.gapTopRange[0], Math.min(C.gapTopRange[1], a[2])); i++; continue; }   // @cell 줄배수 하위간격 상위간격 (범위 안에서만)
    if (t.startsWith("@line")) { const a = t.split(/\s+/).slice(1).map(Number); if (a[0]) CONFIG.lineExact.body = a[0]; if (a[1]) CONFIG.lineExact.note = a[1]; i++; continue; }
    if (t === "@kv") { kvNext = true; i++; continue; }
    if (t.startsWith("@header")) { headerRows = parseInt(t.split(/\s+/)[1] || "1", 10); i++; continue; }
    if (t === "@firstcol") { firstColNext = true; i++; continue; }
    if (t === "@dotted") { dottedNext = true; i++; continue; }
    if (t === "@bluehdr") { blueNext = true; i++; continue; }
    if (t === "@compact") { compactNext = true; i++; continue; }
    if (t === "@equal") { equalNext = true; i++; continue; }
    if (t === "@center") { centerNext = true; i++; continue; }
    if (t === "@pad") { padNext = true; i++; continue; }
    if (t === "@dense") { denseNext = true; i++; continue; }
    if (t.startsWith("@align ")) { alignNext = t.split(/\s+/).slice(1); i++; continue; }   // 열별 정렬 c/l/r
    if (t === "@indent") { indentNext = noteLeft || 0; i++; continue; }   // 다음 표·그림·제목 박스를 위 문장 글자 시작점에 맞춤
    if (t.startsWith("@rowh")) { rowHNext = parseFloat(t.split(/\s+/)[1]) || 0; i++; continue; }

    // ::: 블록
    if (t.startsWith(":::")) {
      let head = t.slice(3).trim();
      const tightB = /\{tight\}/.test(head); head = head.replace(/\s*\{tight\}/, "");
      // 박스 안 키워드 형광: 코퍼스 다수(키워드가 있는 박스 8/10~11)라 기본 켬, {nohl}로 끔. {hl}은 하위 호환
      const hlB = !/\{nohl\}/.test(head); head = head.replace(/\s*\{(no)?hl\}/, "");
      const fsM = head.match(/\s*\{fs:(\d+(?:\.\d+)?)\}/); const boxFs = fsM ? parseFloat(fsM[1]) : null; if (fsM) head = head.replace(fsM[0], "");   // {fs:11} 박스 글자 크기
      const yellowB = /\{yellow\}/.test(head); head = head.replace(/\s*\{yellow\}/, "");
      const white = /\{white\}\s*$/.test(head); head = head.replace(/\s*\{white\}\s*$/, "");
      const solid = white || /\{solid\}\s*$/.test(head); head = head.replace(/\s*\{solid\}\s*$/, "");
      const body = []; i++;
      let depth = 0;   // 박스 안 박스(중첩) 지원: 안쪽 ::: 제목 ~ ::: 쌍은 통째로 본문에 포함
      while (i < lines.length && !(lines[i].trim() === ":::" && depth === 0)) {
        const lt = lines[i].trim();
        if (lt.startsWith(":::") && lt.length > 3) depth++; else if (lt === ":::") depth--;
        body.push(lines[i]); i++; }
      i++; // closing
      prevWasBox = true;
      if (head.startsWith("cover")) { out.push(...coverPage(body)); prevWasBox = false; continue; }
      if (head.startsWith("toc")) { out.push(...tocPage(body)); prevWasBox = false; continue; }
      if (head.startsWith("cols")) { out.push(...colsBlock(parseFloat(head.split(/\s+/)[1]) || 50, body, ctx).slice(0, -1)); prevWasBox = false; prevWasTable = true; continue; }
      if (head.startsWith("timeline")) { out.push(spacer(4), timeline(body, tableFont), spacer(4)); }
      else if (head.startsWith("※") || head.startsWith("*")) {   // 제목은 박스 밖 위에, 박스는 점선 테두리(원본: "※ MOU 체결 개요" + 박스, "* (참고) 그간의 경과" + 박스)
        // 바로 위 문장의 글자 시작 위치에 제목과 박스를 맞춤(각주와 같은 규칙). 위 문장이 없으면 여백 1mm
        const ind = indentNext || 0;   // 기본은 왼쪽 끝까지 꽉 채움(작성자 확인). @indent를 앞에 두면 위 문장 글자 시작점에 맞춤
        indentNext = 0;
        out.push(new Paragraph({ ...pb(), children: (() => { const lpH = lineParts(head); const [hh, csX] = fitLine(lpH.t, CONFIG.fonts.boxTitle, CONTENT_W() / 20, { forcedTw: lpH.forcedTw }); return inline(hh, CONFIG.fonts.boxTitle, { ...emph(CONFIG.fonts.boxTitle), shrinkParens: CONFIG.fit.shrinkParens, ...csX }); })(), indent: { left: Math.round(ind * MM) },
          spacing: { before: mmSp(6), after: mmSp(2), ...lineOf(CONFIG.lineExact.note) } }));
        const paras = parseBlocks(body, { inBox: true, tableFont, tightBox: true, boxFs, boxHl: hlB });
        out.push(boxTable(paras, { fill: white ? undefined : (ctx.inBox ? "FFFFFF" : CONFIG.colors.grayFill), borderSpec: border(white || ctx.inBox ? "404040" : CONFIG.colors.grayBorder, 6, white ? BorderStyle.SINGLE : BorderStyle.DOTTED), marginTw: 45, indentMm: ind,
          ...(ctx.inBox ? { widthDxa: Math.round(CONTENT_W() - 300 - ind * MM) } : {}) }), spacer(4));   // 박스 안 박스: 바깥 박스 안쪽 폭, 흰 바탕+점선
      }
      else if (head.startsWith("참고") || head.startsWith("box")) {
        const title = head.startsWith("box") ? "" : head.replace(/^참고\s*/, "");
        out.push(...grayBox(head.startsWith("box") ? "" : (title ? title : "※ (참고)"), body, tableFont, solid, white, tightB, yellowB, hlB, !!ctx.inBox, boxFs));
      } else out.push(...grayBox(head, body, tableFont, solid, white, tightB, yellowB, hlB, !!ctx.inBox, boxFs));
      continue;
    }
    // 표
    if (t.startsWith("|")) {
      const rows = []; while (i < lines.length && lines[i].trim().startsWith("|")) { const r = lines[i].trim(); if (!/^\|[\s:-|]+\|$/.test(r)) rows.push(r); i++; }
      out.push(spacer(3), mdTable(rows, tableFont, { inBox: ctx.inBox, kv: kvNext, headerRows: kvNext ? 0 : headerRows, firstCol: firstColNext, dotted: dottedNext, blueHdr: blueNext, compact: compactNext, rowH: rowHNext, indentMm: indentNext, equal: equalNext, center: centerNext, rowPad: padNext, dense: denseNext, colAlignSpec: alignNext }));
      indentNext = 0; equalNext = false; centerNext = false; padNext = false; denseNext = false; alignNext = null;
      kvNext = false; firstColNext = false; dottedNext = false; blueNext = false; compactNext = false; rowHNext = 0; headerRows = 1; prevWasTable = true; continue;
    }
    // 결론 박스 / 프레이밍 박스 (연속 줄 묶음)
    if (t.startsWith(">>>")) { const b = []; while (i < lines.length && lines[i].trim().startsWith(">>>")) b.push(lines[i].trim().slice(3).trim()), i++; out.push(...frameBox(b, variantOf(b, "green"))); continue; }
    if (t.startsWith(">>")) { const b = []; while (i < lines.length && lines[i].trim().startsWith(">>") && !lines[i].trim().startsWith(">>>")) b.push(lines[i].trim().slice(2).trim()), i++; out.push(...conclusionBox(b, variantOf(b, "white"))); continue; }
    if (t.startsWith(">") && !ctx.inBox) { const b = []; while (i < lines.length && /^>(?!>)/.test(lines[i].trim())) b.push(lines[i].trim().slice(1).trim()), i++; out.push(...summaryBox(b, variantOf(b, "peach"))); continue; }

    if (t === "---") {   // 나누기 직전의 간격용 빈 문단은 제거(페이지 끝에 남아 빈 페이지를 만들 수 있음)
      while (out.length && out[out.length - 1]._spacer) out.pop();
      PENDING_BREAK = true; i++; continue;
    }
    if (t.startsWith("# ") && !ctx.inBox) { out.push(titleBand(t.slice(2).trim())); i++; continue; }
    if (t.startsWith("@doctitle ")) {   // 첨부 안에 실은 문서(양해각서 등)의 제목
      out.push(new Paragraph({ ...pb(), keepNext: true, alignment: AlignmentType.CENTER, spacing: { before: mmSp(6), after: mmSp(10), ...lineOf(24) },
        children: inline(t.slice(10).trim(), { ...CONFIG.fonts.title, size: 16 }, { bold: true }) }));
      i++; continue;
    }
    if (t.startsWith("@subtitle ")) {   // 제목 아래 가운데 부제(※ 주제 : …)
      out.push(new Paragraph({ ...pb(), alignment: AlignmentType.CENTER, spacing: { before: mmSp(2), after: 0, ...lineOf(18) },
        children: inline(t.slice(10).trim(), { ...CONFIG.fonts.subhdr, size: 13 }, { color: CONFIG.colors.headerText }) }));
      i++; continue;
    }
    if (t.startsWith("@confidential")) { CONFIDENTIAL = t.replace("@confidential", "").trim(); i++; continue; }   // 모든 페이지 머리말로
    if (t.startsWith("@ ") || t.startsWith("@meta")) { out.push(metaLine(t.replace(/^@(meta)?\s*/, ""))); i++; continue; }
    if (t.startsWith("### ")) { out.push(...subHeader(t.slice(4).trim())); noteLeft = null; lvStack.length = 0; prevLevel = 0; i++; continue; }
    if (t.startsWith("![")) { out.push(...imageBlock(t, indentNext)); indentNext = 0; i++; continue; }
    if (t.startsWith("## ")) {
      noteLeft = null; lvStack.length = 0; prevLevel = 0;   // 헤더에서 층위 초기화, 헤더 다음의 박스·표는 헤더 시작점(왼쪽 끝)
      const h = t.slice(3).trim();
      let m;
      if ((m = h.match(/^(첨\s?부\s?\d*|참\s?고\s?\d*|별\s?첨\s?\d*)\s+(.*)$/))) out.push(...attachHeader(m[1], m[2]));
      else if ((m = h.match(/^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|[IVX]{1,4})\.\s+(.*)$/))) out.push(...bandHeader(h));
      else if ((m = h.match(/^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|[IVX]{1,4}|\d{1,2})\s+(.*)$/))) out.push(...boxedNumberHeader(m[1], m[2]));
      else out.push(plainHeader(h));
      i++; continue;
    }
    if (/^[<\[\【].*[>\]\】]$/.test(t)) { out.push(caption(t, !!ctx.inBox && out.length === 0)); i++; continue; }
    let mk = t.split(/\s+/)[0]; if (!INDENT[mk] && /^[➊-➓①-⑳❶-❿]\S/.test(t)) mk = t[0];   // 기호에 글자가 붙은 줄
    if (!INDENT[mk] && mk.length <= 2 && /^[^\p{L}\p{N}(「“‘'"\[<{*^_~!]/u.test(mk)) LEVEL_WARN.push(`들여쓰기 규칙이 없는 기호 "${mk}": "${t.slice(0, 30)}"`);   // 규칙 누락 점검
    const isNoteLine = NOTE_MARKERS.has(mk);
    // 표 제목 역할의 ※·* 줄(바로 다음이 표): 표 왼쪽 끝에 맞춤
    let titlesTable = false;
    if (isNoteLine) { let k = i + 1; while (k < lines.length && (!lines[k].trim() || lines[k].trim().startsWith("@"))) k++; titlesTable = k < lines.length && lines[k].trim().startsWith("|"); }
    // 층위 판단: □ 아래의 ①은 하위(연노랑), □ 없이 헤더 바로 아래 ①은 최상위(□ 위치·연두), 그 아래 ❶·➊은 한 칸 들여씀(연노랑)
    let lv = {};
    if ((!isNoteLine || BULLET_NOTES.has(mk)) && (LEVEL_MARKERS.test(mk) || ARROW_LV.test(mk))) {
      const spaces = (raw.match(/^ */) || [""])[0].length;
      let level;
      if (explicitLv) level = 1 + Math.floor(spaces / 2);
      else if (ARROW_LV.test(mk)) level = lvStack.length ? 2 : 1;          // 화살표 줄은 층위를 만들지 않고 2층위 위치
      else {
        const fam = markerFamily(mk);
        const k = lvStack.indexOf(fam);
        if (k >= 0) { lvStack.length = k + 1; level = k + 1; }             // 이미 나온 종류로 돌아오면 그 층위로 복귀(➊ 다음 ➋도 같은 층위)
        else { lvStack.push(fam); level = lvStack.length; }                 // 처음 나온 종류는 한 단계 깊게
      }
      if (level > prevLevel + 1 && prevLevel > 0) LEVEL_WARN.push(`층위가 두 단계 이상 깊어짐: "${t.slice(0, 30)}"`);
      if (explicitLv && level > 1 && prevLevel === 0) LEVEL_WARN.push(`위 층위 없이 ${level}층위로 시작: "${t.slice(0, 30)}"`);
      prevLevel = level; lv = { level };
      if (process.env.LEVEL_TRACE) LEVEL_TRACE.push(`${level}\t${ctx.inBox ? "box" : "body"}\t${t.slice(0, 40)}`);
    }
    if ((!isNoteLine || BULLET_NOTES.has(mk)) && INDENT[mk]) { const ind = lv.indentOverride || INDENT[mk]; const bf = ctx.inBox ? CONFIG.fonts.box : CONFIG.fonts.body; noteLeft = (lv.level ? LEVEL_LEFT[Math.min(lv.level, LEVEL_LEFT.length) - 1] : ind[0]) + estWidthPt(mk + " ", bf.size, 0, bf.name) * 0.3528; }   // 본문 문단의 글자 시작 위치(mm) 기억
    if (!isNoteLine && !INDENT[mk]) noteLeft = null;
    if (prevWasTable) {
      if (isNoteLine) {   // 표 바로 아래 각주(연속된 각주 포함): 간격 2pt, 표 왼쪽 끝 정렬, 11pt
        const subNote = /^\*{1,2}$/.test(mk) && /^\d\)$/.test(lastTableNote || "");
        out.push(bodyParagraph(t, { ...ctx, afterTable: true, subNote }));
        if (!/^\*{1,2}$/.test(mk)) lastTableNote = mk;
        prevWasNote = true; i++; continue;   // 표 각주 모드 유지
      }
      out.push(spacer(4)); out.push(bodyParagraph(t, { ...ctx, ...lv }));
      prevWasTable = false; lastTableNote = null;
    } else if (prevWasBox && isNoteLine) {
      out.push(bodyParagraph(t, { ...ctx, afterBox: true }));   // 박스 바로 아래 ※·* 는 왼쪽 끝 정렬(크기는 그대로)
    } else {
      out.push(bodyParagraph(t, { ...ctx, ...lv, tightGap: isNoteLine && prevWasNote, noteLeft, afterBox: titlesTable || undefined, firstInBox: !!ctx.inBox && out.length === 0 }));  // 각주가 연달아 오면 간격 2pt
    }
    prevWasNote = isNoteLine; i++;
  }
  return out;
}

// ───────────────────────────── main ─────────────────────────────
const DEFAULTS = JSON.parse(JSON.stringify({ spacing: CONFIG.spacing, lineExact: CONFIG.lineExact, cell: CONFIG.cell, boxRange: CONFIG.boxRange }));
function build(inputPath, outputPath) {
  PENDING_BREAK = false; LEVEL_WARN = []; LEVEL_TRACE = []; CONFIDENTIAL = null; WIDTH_OVERRIDE = null; INPUT_DIR = path.dirname(path.resolve(inputPath));
  Object.assign(CONFIG.spacing, DEFAULTS.spacing); Object.assign(CONFIG.lineExact, DEFAULTS.lineExact);
  CONFIG.cell = JSON.parse(JSON.stringify(DEFAULTS.cell)); CONFIG.boxRange = JSON.parse(JSON.stringify(DEFAULTS.boxRange));
  const src = fs.readFileSync(inputPath, "utf8").replace(/\r\n/g, "\n");
  const lines = src.split("\n");
  const children = parseBlocks(lines);
  const doc = new Document({
    styles: { default: { document: { run: font(CONFIG.fonts.body) } } },
    sections: [{
      ...(CONFIDENTIAL ? { headers: { default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 0, ...lineOf(12) },
        children: inline(CONFIDENTIAL, { name: "나눔고딕", size: 10 }, { color: CONFIG.colors.red }) })] }) } } : {}),
      properties: { page: {
        size: { width: Math.round(210 * MM), height: Math.round(297 * MM) },
        margin: { top: Math.round(CONFIG.page.top * MM), bottom: Math.round(CONFIG.page.bottom * MM),
                  left: Math.round(CONFIG.page.left * MM), right: Math.round(CONFIG.page.right * MM) },
      } },
      children,
    }],
  });
  return Packer.toBuffer(doc).then(async buf => {
    const irPath = outputPath.replace(/\.hwpx$/i, "") + ".ir.json";
    const ir = JSON.parse(buf.toString());
    ir.config = { page: CONFIG.page, fonts: CONFIG.fonts, colors: CONFIG.colors, inputDir: INPUT_DIR };
    fs.writeFileSync(irPath, JSON.stringify(ir));
    const { spawnSync } = require("child_process");
    const r = spawnSync("python3", [path.join(__dirname, "hwpx_writer.py"), irPath, outputPath, ...(process.env.HWPX_KEEP_IR ? [] : ["--rm-ir"])], { stdio: "inherit" });
    if (r.status !== 0) throw new Error("hwpx_writer.py 실패");
    if (LEVEL_WARN.length) console.error("[층위 점검]\n  " + LEVEL_WARN.join("\n  "));
    if (process.env.LEVEL_TRACE) fs.writeFileSync(process.env.LEVEL_TRACE, LEVEL_TRACE.join("\n") + "\n");
  });
}

if (require.main === module) {
  const [inp, outp] = process.argv.slice(2);
  if (!inp || !outp) { console.error("usage: node build_hwpx.js input.md output.hwpx"); process.exit(1); }
  build(inp, outp).catch(e => { console.error(e); process.exit(1); });
}
module.exports = { build, CONFIG };
