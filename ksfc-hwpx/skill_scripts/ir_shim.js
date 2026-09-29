/**
 * ir_shim.js — build_hwpx.js가 쓰는 "docx" 라이브러리 대역.
 * 마크업 해석·작성 규칙(build_docx.js에서 가져온 부분)은 그대로 두고, 만들어진 문서 구조를
 * 중간 표현(IR, JSON)으로만 기록한다. 실제 hwpx 쓰기는 hwpx_writer.py가 맡는다.
 */
const AlignmentType = { LEFT: "left", CENTER: "center", RIGHT: "right", BOTH: "both", JUSTIFIED: "both", DISTRIBUTE: "distribute" };
const BorderStyle = { NONE: "none", SINGLE: "single", DOTTED: "dotted", DASHED: "dashed", DOUBLE: "double", THICK: "single" };
const WidthType = { DXA: "dxa", PERCENTAGE: "pct", AUTO: "auto" };
const ShadingType = { CLEAR: "clear", SOLID: "solid" };
const VerticalAlign = { TOP: "top", CENTER: "center", BOTTOM: "bottom" };
const TableLayoutType = { FIXED: "fixed", AUTOFIT: "autofit" };
const LineRuleType = { EXACT: "exact", AUTO: "auto", AT_LEAST: "atLeast" };
const VerticalMergeType = { RESTART: "restart", CONTINUE: "continue" };

class Node {
  constructor(type, opts) { this._t = type; Object.assign(this, opts || {}); }
  toIR() {
    const o = { type: this._t };
    for (const [k, v] of Object.entries(this)) {
      if (k === "_t") continue;
      o[k === "type" ? "kind" : k] = toIR(v);   // ImageRun의 type(jpg/png)이 요소 종류와 겹치지 않게
    }
    return o;
  }
}
function toIR(v) {
  if (v instanceof Node) return v.toIR();
  if (Buffer.isBuffer(v)) return { $b64: v.toString("base64") };
  if (Array.isArray(v)) return v.map(toIR);
  if (v && typeof v === "object") { const o = {}; for (const [k, x] of Object.entries(v)) o[k] = toIR(x); return o; }
  return v;
}
class Paragraph extends Node { constructor(o) { super("p", o); } }
class TextRun extends Node { constructor(o) { super("r", typeof o === "string" ? { text: o } : o); } }
class ImageRun extends Node { constructor(o) { super("img", o); } }
class Table extends Node { constructor(o) { super("tbl", o); } }
class TableRow extends Node { constructor(o) { super("tr", o); } }
class TableCell extends Node { constructor(o) { super("tc", o); } }
class Header extends Node { constructor(o) { super("header", o); } }
class PageBreak extends Node { constructor() { super("pagebreak", {}); } }
class Document extends Node { constructor(o) { super("doc", o); } }
const Packer = { toBuffer: async (doc) => Buffer.from(JSON.stringify(doc.toIR())) };

module.exports = {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, BorderStyle, WidthType,
  ShadingType, PageBreak, VerticalAlign, TableLayoutType, LineRuleType, VerticalMergeType, ImageRun, Header,
};
