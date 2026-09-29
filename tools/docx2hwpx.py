#!/usr/bin/env python3
"""docx(HWP 서식 재현본) → hwpx 변환.

기계적 치환이 아니라 docx가 흉내 내던 HWP 설정을 HWP 고유 방식으로 되돌린다.
  - 줄간격  : Word '고정(exact)' → 글자 크기 대비 %(글자에 따라). 간격용 빈 줄만 고정값
  - 자간    : Word twip(글자당 고정 pt) → HWP %(글자 폭 대비), 장평 w:w → ratio
  - 띄어쓰기: HWP 기본은 0.5em이라 Word(글꼴 고유 폭)보다 넓다 → '글꼴에 어울리는 빈칸'
  - 줄나눔  : Word wordWrap(어절) → HWP 한글 어절 단위(BREAK_WORD)
  - 기호+탭+내어쓰기 → HWP 내어쓰기(왼쪽 여백=첫 줄 시작) + 내어쓰기용 자동 탭
  - 여백    : HWP는 위쪽 여백 아래에 머리말 영역을 더 두므로 머리말·꼬리말 0
  - 표      : 글자처럼 취급, 셀 단위 쪽 나눔, 제목 행 반복, 셀별 테두리·음영·안 여백
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP_DECL = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)
HUC = "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"
LANGS = ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"]
LANG_ATTRS = ["hangul", "latin", "hanja", "japanese", "other", "symbol", "user"]

TW = 5  # 1 twip = 5 HWPUNIT (1/1440 in → 1/7200 in)

# 글꼴별 한글 한 글자 폭(em). HWP 자간은 '각 글자 폭'에 대한 %이므로 환산에 쓴다.
# 나눔 계열은 글꼴 파일 실측, HY 계열은 1em(전각).
HANGUL_ADV = {"나눔명조": 0.950, "나눔고딕": 0.940, "HY헤드라인M": 1.0, "HY울릉도M": 1.0}
# 글꼴에 없는 기호(➊➋➌, ➡ 등)를 Word는 기호 글꼴로 대체해 그린다 → HWP에도 명시
SYMBOL_FALLBACK = "Segoe UI Symbol"
FONT_FAMILY = {"나눔명조": "FCAT_MYUNGJO", "HY헤드라인M": "FCAT_GOTHIC", "HY울릉도M": "FCAT_GOTHIC",
               "나눔고딕": "FCAT_GOTHIC", SYMBOL_FALLBACK: "FCAT_GOTHIC"}

BORDER_WIDTHS = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]
BORDER_TYPE = {"single": "SOLID", "dotted": "DOT", "dashed": "DASH", "double": "DOUBLE_SLIM",
               "dotDash": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT", "thick": "SOLID"}


def q(tag):
    return "{%s}%s" % (W, tag)


def wattr(el, name, default=None):
    if el is None:
        return default
    return el.get(q(name), default)


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


# ───────────────────────── 글꼴 글리프 보유 여부 ─────────────────────────
def _font_cmaps():
    cmaps = {}
    try:
        from fontTools.ttLib import TTFont
        for face, fn in (("나눔명조", "NanumMyeongjo"), ("나눔고딕", "NanumGothic")):
            p = Path(f"/usr/share/fonts/truetype/nanum/{fn}.ttf")
            if p.exists():
                cmaps[face] = set(TTFont(str(p)).getBestCmap())
    except Exception:
        pass
    return cmaps


CMAPS = _font_cmaps()


def has_glyph(face, ch):
    if ch in " \t":
        return True
    if face in CMAPS:
        return ord(ch) in CMAPS[face]
    try:  # HY 글꼴: KS X 1001(완성형) 범위는 갖춤
        ch.encode("euc-kr")
        return True
    except UnicodeEncodeError:
        return False


# ───────────────────────── 헤더(스타일 표) 레지스트리 ─────────────────────────
class Registry:
    def __init__(self, skel_header):
        self.root = etree.fromstring(skel_header)
        self.fonts = []            # 추가 글꼴 이름
        self.char = {}             # key → id
        self.char_xml = []
        self.para = {}
        self.para_xml = []
        self.bf = {}
        self.bf_xml = []
        self.tab = {}
        self.tab_xml = []
        rl = self.root.find("{*}refList")
        self.base_font = len(rl.find("{*}fontfaces")[0])       # 스켈레톤 글꼴 수(2)
        self.base_bf = len(rl.find("{*}borderFills")) + 1     # 다음 id(3)
        self.base_char = len(rl.find("{*}charProperties"))
        self.base_para = len(rl.find("{*}paraProperties"))
        self.base_tab = len(rl.find("{*}tabProperties"))

    def font_id(self, face):
        if face not in self.fonts:
            self.fonts.append(face)
        return self.base_font + self.fonts.index(face)

    # borderFill: sides = dict(left/right/top/bottom → (type, width_mm, color)), fill=hex|None
    def border_fill(self, sides, fill=None):
        key = (tuple(sorted(sides.items())), fill)
        if key in self.bf:
            return self.bf[key]
        bid = self.base_bf + len(self.bf_xml)
        parts = [f'<hh:borderFill id="{bid}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">',
                 '<hh:slash type="NONE" Crooked="0" isCounter="0"/>',
                 '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>']
        for side in ("left", "right", "top", "bottom"):
            t, wmm, col = sides.get(side, ("NONE", 0.1, "#000000"))
            parts.append(f'<hh:{side}Border type="{t}" width="{fmt_mm(wmm)} mm" color="{col}"/>')
        parts.append('<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>')
        if fill:
            parts.append(f'<hc:fillBrush><hc:winBrush faceColor="#{fill}" hatchColor="#999999" alpha="0"/></hc:fillBrush>')
        parts.append("</hh:borderFill>")
        self.bf_xml.append("".join(parts))
        self.bf[key] = bid
        return bid

    def char_pr(self, face, height, bold=False, color="000000", shade=None, spacing=0, ratio=100,
                sym_face=None):
        key = (face, height, bold, color, shade, spacing, ratio, sym_face)
        if key in self.char:
            return self.char[key]
        cid = self.base_char + len(self.char_xml)
        fid = self.font_id(face)
        refs = {a: fid for a in LANG_ATTRS}
        if sym_face:  # 글꼴에 없는 기호만 담은 조각: 모든 언어 칸을 기호 글꼴로
            sid = self.font_id(sym_face)
            refs = {a: sid for a in LANG_ATTRS}
        attrs = lambda d: " ".join(f'{a}="{d[a]}"' for a in LANG_ATTRS)
        same = lambda v: {a: v for a in LANG_ATTRS}
        shade_attr = f"#{shade}" if shade else "none"
        x = (f'<hh:charPr id="{cid}" height="{height}" textColor="#{color}" shadeColor="{shade_attr}" '
             f'useFontSpace="1" useKerning="0" symMark="NONE" borderFillIDRef="2">'
             f'<hh:fontRef {attrs(refs)}/>'
             f'<hh:ratio {attrs(same(ratio))}/>'
             f'<hh:spacing {attrs(same(spacing))}/>'
             f'<hh:relSz {attrs(same(100))}/>'
             f'<hh:offset {attrs(same(0))}/>'
             + ("<hh:bold/>" if bold else "") +
             '<hh:underline type="NONE" shape="SOLID" color="#000000"/>'
             '<hh:strikeout shape="NONE" color="#000000"/>'
             '<hh:outline type="NONE"/>'
             '<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
             '</hh:charPr>')
        self.char_xml.append(x)
        self.char[key] = cid
        return cid

    def tab_pr(self, hanging):
        """내어쓰기용 자동 탭(autoTabLeft) + 같은 위치의 명시 탭(문단 왼쪽 여백 기준)."""
        key = hanging
        if key in self.tab:
            return self.tab[key]
        tid = self.base_tab + len(self.tab_xml)
        x = (f'<hh:tabPr id="{tid}" autoTabLeft="1" autoTabRight="0">'
             f'<hp:switch><hp:case hp:required-namespace="{HUC}">'
             f'<hh:tabItem pos="{hanging}" type="LEFT" leader="NONE" unit="HWPUNIT"/></hp:case>'
             f'<hp:default><hh:tabItem pos="{hanging * 2}" type="LEFT" leader="NONE"/></hp:default>'
             f'</hp:switch></hh:tabPr>')
        self.tab_xml.append(x)
        self.tab[key] = tid
        return tid

    def para_pr(self, align="JUSTIFY", left=0, intent=0, right=0, prev=0, nxt=0,
                ls_type="PERCENT", ls_val=160, keep_next=False, keep_lines=False,
                border=2, tab=0):
        key = (align, left, intent, right, prev, nxt, ls_type, ls_val, keep_next, keep_lines, border, tab)
        if key in self.para:
            return self.para[key]
        pid = self.base_para + len(self.para_xml)

        def margin(dbl):
            m = 2 if dbl else 1
            return ("<hh:margin>"
                    f'<hc:intent value="{intent * m}" unit="HWPUNIT"/>'
                    f'<hc:left value="{left * m}" unit="HWPUNIT"/>'
                    f'<hc:right value="{right * m}" unit="HWPUNIT"/>'
                    f'<hc:prev value="{prev * m}" unit="HWPUNIT"/>'
                    f'<hc:next value="{nxt * m}" unit="HWPUNIT"/>'
                    "</hh:margin>"
                    f'<hh:lineSpacing type="{ls_type}" value="{ls_val if ls_type == "PERCENT" else ls_val * m}" unit="HWPUNIT"/>')
        x = (f'<hh:paraPr id="{pid}" tabPrIDRef="{tab}" condense="0" fontLineHeight="0" snapToGrid="0" '
             f'suppressLineNumbers="0" checked="0" textDir="LTR">'
             f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
             '<hh:heading type="NONE" idRef="0" level="0"/>'
             f'<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" '
             f'keepWithNext="{int(keep_next)}" keepLines="{int(keep_lines)}" pageBreakBefore="0" lineWrap="BREAK"/>'
             '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
             f'<hp:switch><hp:case hp:required-namespace="{HUC}">{margin(False)}</hp:case>'
             f'<hp:default>{margin(True)}</hp:default></hp:switch>'
             f'<hh:border borderFillIDRef="{border}" offsetLeft="0" offsetRight="0" offsetTop="0" '
             f'offsetBottom="0" connect="0" ignoreMargin="0"/>'
             '</hh:paraPr>')
        self.para_xml.append(x)
        self.para[key] = pid
        return pid

    def build(self):
        rl = self.root.find("{*}refList")
        HH = "http://www.hancom.co.kr/hwpml/2011/head"
        wrap = lambda s: etree.fromstring(f"<x {NSMAP_DECL}>{s}</x>")
        # 글꼴
        for ff in rl.find("{*}fontfaces"):
            for i, face in enumerate(self.fonts):
                fam = FONT_FAMILY.get(face, "FCAT_GOTHIC")
                el = wrap(f'<hh:font id="{self.base_font + i}" face="{esc(face)}" type="TTF" isEmbedded="0">'
                          f'<hh:typeInfo familyType="{fam}" weight="6" proportion="4" contrast="0" '
                          f'strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font>')[0]
                ff.append(el)
            ff.set("fontCnt", str(len(ff)))
        for name, xs in (("borderFills", self.bf_xml), ("charProperties", self.char_xml),
                         ("tabProperties", self.tab_xml), ("paraProperties", self.para_xml)):
            box = rl.find("{%s}%s" % (HH, name))
            for x in xs:
                box.append(wrap(x)[0])
            box.set("itemCnt", str(len(box)))
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fmt_mm(v):
    s = ("%.2f" % v).rstrip("0").rstrip(".")
    return s if "." in s else s + ".0"


def border_mm(sz_eighths):
    mm = sz_eighths / 8 * 0.352778
    return min(BORDER_WIDTHS, key=lambda w: (abs(w - mm), w))


def wborder(el):
    """w:top 등 → (type, width_mm, #color) 또는 None(지정 없음)."""
    if el is None:
        return None
    val = wattr(el, "val", "none")
    if val in ("none", "nil"):
        return ("NONE", 0.1, "#000000")
    col = wattr(el, "color", "000000")
    col = "000000" if col in ("auto", None) else col
    return (BORDER_TYPE.get(val, "SOLID"), border_mm(int(wattr(el, "sz", "4"))), "#" + col.upper())


# ───────────────────────── 변환기 ─────────────────────────
class Skeleton:
    """한컴 뼈대 문서(header.xml 기본 스타일·version.xml 등). 기본값: python-hwpx 내장 Skeleton.hwpx"""

    def __init__(self, path=None):
        if path is None:
            import hwpx
            path = Path(hwpx.__file__).parent / "data" / "Skeleton.hwpx"
        path = Path(path)
        if path.is_dir():
            self.read = lambda name: (path / name).read_bytes()
        else:
            z = zipfile.ZipFile(path)
            self.read = z.read


class Converter:
    def __init__(self, docx_path, skel):
        z = zipfile.ZipFile(docx_path)
        self.doc = etree.fromstring(z.read("word/document.xml"))
        styles = etree.fromstring(z.read("word/styles.xml"))
        rpr = styles.find(f".//{q('rPrDefault')}/{q('rPr')}")
        f = rpr.find(q("rFonts"))
        self.def_font = wattr(f, "eastAsia") or wattr(f, "ascii") or "나눔명조"
        self.def_sz = int(wattr(rpr.find(q("sz")), "val", "20"))
        self.reg = Registry(skel.read("Contents/header.xml"))
        self.pid = 0
        self.oid = 1000000
        self.stats = {"percent": 0, "fixed": 0, "zwsp": 0, "fe0e": 0, "tables": 0}
        self.plain = []

    def next_pid(self):
        self.pid += 1
        return self.pid

    # ── 글자 모양 ──
    def run_props(self, r):
        rpr = r.find(q("rPr"))
        face, sz, bold, color, shade, sp, w = self.def_font, self.def_sz, False, "000000", None, 0, 100
        if rpr is not None:
            f = rpr.find(q("rFonts"))
            if f is not None and (wattr(f, "eastAsia") or wattr(f, "ascii")):
                face = wattr(f, "eastAsia") or wattr(f, "ascii")
            if rpr.find(q("sz")) is not None:
                sz = int(wattr(rpr.find(q("sz")), "val"))
            b = rpr.find(q("b"))
            bold = b is not None and wattr(b, "val", "1") not in ("0", "false")
            c = rpr.find(q("color"))
            if c is not None and wattr(c, "val") not in (None, "auto"):
                color = wattr(c, "val").upper()
            s = rpr.find(q("shd"))
            if s is not None and wattr(s, "fill") not in (None, "auto"):
                shade = wattr(s, "fill").upper()
            if rpr.find(q("spacing")) is not None:
                sp = int(wattr(rpr.find(q("spacing")), "val"))
            if rpr.find(q("w")) is not None:
                w = int(wattr(rpr.find(q("w")), "val"))
        return face, sz, bold, color, shade, sp, w

    def char_id(self, props, sym=False):
        face, sz, bold, color, shade, sp, w = props
        size_pt = sz / 2
        # Word 자간: 글자마다 sp/20 pt 고정 → HWP 자간: 글자 폭(장평 반영) 대비 %
        pct = 0
        if sp:
            glyph = size_pt * HANGUL_ADV.get(face, 1.0) * w / 100
            pct = max(-50, min(50, round((sp / 20) / glyph * 100)))
        return self.reg.char_pr(face, sz * 50, bold, color, shade, pct, w,
                                sym_face=SYMBOL_FALLBACK if sym else None)

    # ── 문단 ──
    def runs_of(self, p):
        """문단 안의 (props, text|'\t') 조각 목록. 표시용 제어 문자는 정리."""
        out = []
        for r in p.iter(q("r")):
            props = self.run_props(r)
            for ch in r:
                if ch.tag == q("t"):
                    t = ch.text or ""
                    self.stats["zwsp"] += t.count("​")
                    self.stats["fe0e"] += t.count("︎")
                    # U+200B(Word용 줄바꿈 기회)·U+FE0E(Word용 글자 모양 선택자)는 HWP에서 불필요
                    t = t.replace("​", "").replace("︎", "")
                    if t:
                        out.append((props, t))
                elif ch.tag == q("tab"):
                    out.append((props, "\t"))
                elif ch.tag == q("br"):
                    out.append((props, "\n"))
        return out

    def para_xml(self, p, in_cell=False, page_break=False, first_extra=""):
        ppr = p.find(q("pPr"))
        sp = ppr.find(q("spacing")) if ppr is not None else None
        ind = ppr.find(q("ind")) if ppr is not None else None
        jc = wattr(ppr.find(q("jc")) if ppr is not None else None, "val", "left")
        align = {"both": "JUSTIFY", "center": "CENTER", "right": "RIGHT", "left": "LEFT",
                 "distribute": "DISTRIBUTE", "start": "LEFT", "end": "RIGHT"}.get(jc, "LEFT")
        before = int(wattr(sp, "before", "0") or 0)
        after = int(wattr(sp, "after", "0") or 0)
        line = wattr(sp, "line")
        rule = wattr(sp, "lineRule", "auto")
        left = int(wattr(ind, "left", "0") or 0)
        hanging = int(wattr(ind, "hanging", "0") or 0)
        first = int(wattr(ind, "firstLine", "0") or 0)
        keep_next = ppr is not None and ppr.find(q("keepNext")) is not None
        keep_lines = ppr is not None and ppr.find(q("keepLines")) is not None
        pbb = ppr is not None and ppr.find(q("pageBreakBefore")) is not None

        pieces = self.runs_of(p)
        text_pieces = [(pr, t) for pr, t in pieces if t.strip()]
        has_text = bool(text_pieces)

        # 줄간격
        if has_text:
            max_pt = max(pr[1] for pr, _ in text_pieces) / 2
        else:
            # 빈 문단: 글자 크기는 (있으면) 빈 run의 크기, 없으면 문단 기호(기본) 크기
            rs = [self.run_props(r) for r in p.iter(q("r"))]
            max_pt = (rs[0][1] if rs else self.def_sz) / 2
        if line is None:
            ls_type, ls_val = "PERCENT", 115          # Word '1줄' = 글꼴 행높이(나눔 1.15em)
        elif rule == "exact" and has_text:
            ls_type, ls_val = "PERCENT", round(int(line) / 20 / max_pt * 100)   # 예: 24pt/15pt = 160%
        elif rule == "exact":
            ls_type, ls_val = "FIXED", int(line) * TW  # 간격용 빈 줄: 높이를 그대로
        elif rule == "atLeast":
            ls_type, ls_val = "AT_LEAST", int(line) * TW
        else:  # auto: 240 = 1줄(나눔 계열 1.15em)
            ls_type, ls_val = "PERCENT", round(int(line) / 240 * 115)
        self.stats["percent" if ls_type == "PERCENT" else "fixed"] += 1

        # 들여쓰기: Word(왼쪽=둘째 줄 이후, 내어쓰기=첫 줄만 왼쪽으로) → HWP(왼쪽=첫 줄 시작, 내어쓰기=나머지 줄)
        if hanging:
            h_left, h_intent = (left - hanging) * TW, -hanging * TW
        else:
            h_left, h_intent = left * TW, first * TW
        tab = self.reg.tab_pr(hanging * TW) if (hanging and any(t == "\t" for _, t in pieces)) else 0

        # 문단 테두리·배경(제목 띠)
        border = 2
        pbdr = ppr.find(q("pBdr")) if ppr is not None else None
        pshd = ppr.find(q("shd")) if ppr is not None else None
        if pbdr is not None or (pshd is not None and wattr(pshd, "fill") not in (None, "auto")):
            sides = {}
            for s in ("left", "right", "top", "bottom"):
                b = wborder(pbdr.find(q(s))) if pbdr is not None else None
                sides[s] = b or ("NONE", 0.1, "#000000")
            fill = wattr(pshd, "fill").upper() if pshd is not None and wattr(pshd, "fill") not in (None, "auto") else None
            border = self.reg.border_fill(sides, fill)

        ppid = self.reg.para_pr(align, h_left, h_intent, 0, before * TW, after * TW, ls_type, ls_val,
                                keep_next, keep_lines, border, tab)

        # 글자 조각 → run (같은 모양은 합침, 글꼴에 없는 기호는 기호 글꼴 조각으로 분리)
        segs = []
        for pr, t in pieces:
            face = pr[0]
            buf, sym = "", None
            for chh in t:
                s = not has_glyph(face, chh) and chh not in "\t\n"
                if sym is None or s == sym:
                    buf += chh
                    sym = s
                else:
                    segs.append((self.char_id(pr, sym), buf))
                    buf, sym = chh, s
            if buf:
                segs.append((self.char_id(pr, sym), buf))
        merged = []
        for cid, t in segs:
            if merged and merged[-1][0] == cid:
                merged[-1][1] += t
            else:
                merged.append([cid, t])
        if not merged:
            rs = [self.run_props(r) for r in p.iter(q("r"))]
            base = rs[0] if rs else (self.def_font, self.def_sz, False, "000000", None, 0, 100)
            if not has_text and ls_type == "FIXED":
                # 빈 줄 글자 크기는 줄 높이를 넘지 않게(고정값보다 큰 글자는 HWP에서 넘침 표시)
                base = (base[0], min(base[1], max(2, int(line) // 10)), False, "000000", None, 0, 100)
            merged = [[self.char_id(base), ""]]

        runs = []
        for i, (cid, t) in enumerate(merged):
            inner = first_extra if i == 0 else ""
            runs.append(f'<hp:run charPrIDRef="{cid}">{inner}{self.t_xml(t)}</hp:run>')
        self.plain.append("".join(t for _, t in merged))
        pb = "1" if (pbb or page_break) else "0"
        return (f'<hp:p id="{self.next_pid()}" paraPrIDRef="{ppid}" styleIDRef="0" pageBreak="{pb}" '
                f'columnBreak="0" merged="0">{"".join(runs)}</hp:p>')

    @staticmethod
    def t_xml(t):
        if t == "":
            return "<hp:t/>"
        parts = re.split(r"(\t|\n)", t)
        out = []
        for s in parts:
            if s == "\t":
                out.append('<hp:tab width="0" leader="0" type="1"/>')
            elif s == "\n":
                out.append("<hp:lineBreak/>")
            elif s:
                out.append(esc(s))
        return "<hp:t>" + "".join(out) + "</hp:t>"

    # ── 표 ──
    def table_xml(self, tbl):
        self.stats["tables"] += 1
        tpr = tbl.find(q("tblPr"))
        grid = [int(wattr(g, "w")) for g in tbl.find(q("tblGrid")).findall(q("gridCol"))]
        tb = tpr.find(q("tblBorders"))
        tb_side = {s: wborder(tb.find(q(s))) if tb is not None else None
                   for s in ("top", "left", "bottom", "right", "insideH", "insideV")}
        tcm = tpr.find(q("tblCellMar"))

        def mar(el, side, default):
            if el is None or el.find(q(side)) is None:
                return default
            return int(wattr(el.find(q(side)), "w"))
        def_mar = {s: mar(tcm, s, d) for s, d in (("left", 108), ("right", 108), ("top", 0), ("bottom", 0))}

        rows = tbl.findall(q("tr"))
        # 격자 배치(가로 병합 gridSpan, 세로 병합 vMerge)
        cells = []   # dict(row, col, span, tc)
        owner = {}   # (r, c) → cell
        for ri, tr in enumerate(rows):
            ci = 0
            for tc in tr.findall(q("tc")):
                tcpr = tc.find(q("tcPr"))
                span = int(wattr(tcpr.find(q("gridSpan")) if tcpr is not None else None, "val", "1"))
                vm = tcpr.find(q("vMerge")) if tcpr is not None else None
                if vm is not None and wattr(vm, "val", "continue") == "continue" and (ri - 1, ci) in owner:
                    top = owner[(ri - 1, ci)]
                    top["rspan"] += 1
                    for k in range(span):
                        owner[(ri, ci + k)] = top
                else:
                    cell = dict(row=ri, col=ci, span=span, rspan=1, tc=tc)
                    cells.append(cell)
                    for k in range(span):
                        owner[(ri, ci + k)] = cell
                ci += span
        nrows, ncols = len(rows), len(grid)
        header_rows = set()
        min_h = {}
        for ri, tr in enumerate(rows):
            trpr = tr.find(q("trPr"))
            if trpr is not None and trpr.find(q("tblHeader")) is not None:
                header_rows.add(ri)
            if trpr is not None and trpr.find(q("trHeight")) is not None:
                min_h[ri] = int(wattr(trpr.find(q("trHeight")), "val")) * TW

        row_h = [0] * nrows
        cell_xml = {}
        for cell in cells:
            tc, ri, ci = cell["tc"], cell["row"], cell["col"]
            tcpr = tc.find(q("tcPr"))
            tcb = tcpr.find(q("tcBorders")) if tcpr is not None else None
            last_r, last_c = ri + cell["rspan"] - 1, ci + cell["span"] - 1
            sides = {}
            for s, edge, inner in (("top", ri == 0, "insideH"), ("bottom", last_r == nrows - 1, "insideH"),
                                   ("left", ci == 0, "insideV"), ("right", last_c == ncols - 1, "insideV")):
                b = wborder(tcb.find(q(s))) if tcb is not None and tcb.find(q(s)) is not None else None
                if b is None:
                    b = tb_side[s] if edge else tb_side[inner]
                sides[s] = b or ("NONE", 0.1, "#000000")
            shd = tcpr.find(q("shd")) if tcpr is not None else None
            fill = wattr(shd, "fill").upper() if shd is not None and wattr(shd, "fill") not in (None, "auto") else None
            bf = self.reg.border_fill(sides, fill)
            tm = tcpr.find(q("tcMar")) if tcpr is not None else None
            m = {s: mar(tm, s, def_mar[s]) * TW for s in ("left", "right", "top", "bottom")}
            va = wattr(tcpr.find(q("vAlign")) if tcpr is not None else None, "val", "top")
            va = {"center": "CENTER", "bottom": "BOTTOM"}.get(va, "TOP")
            width = sum(grid[ci:ci + cell["span"]]) * TW
            paras = [self.para_xml(p, in_cell=True) for p in tc.findall(q("p"))]
            # 셀 높이는 최소값으로만 쓰인다(내용이 크면 한글이 늘림) → 여백 + 한 줄로 낮게 잡음
            est = m["top"] + m["bottom"] + 1000
            if cell["rspan"] == 1:
                row_h[ri] = max(row_h[ri], min_h.get(ri, 0), est)
            cell["x"] = dict(bf=bf, m=m, va=va, width=width, paras=paras,
                             header=int(ri in header_rows))
        for ri in range(nrows):
            row_h[ri] = row_h[ri] or max(min_h.get(ri, 0), 1000)

        trs = []
        for ri in range(nrows):
            tcs = []
            for cell in [c for c in cells if c["row"] == ri]:
                x = cell["x"]
                h = sum(row_h[cell["row"]:cell["row"] + cell["rspan"]])
                m = x["m"]
                tcs.append(
                    f'<hp:tc name="" header="{x["header"]}" hasMargin="1" protect="0" editable="0" dirty="0" '
                    f'borderFillIDRef="{x["bf"]}">'
                    f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="{x["va"]}" '
                    f'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
                    f'{"".join(x["paras"])}</hp:subList>'
                    f'<hp:cellAddr colAddr="{cell["col"]}" rowAddr="{ri}"/>'
                    f'<hp:cellSpan colSpan="{cell["span"]}" rowSpan="{cell["rspan"]}"/>'
                    f'<hp:cellSz width="{x["width"]}" height="{h}"/>'
                    f'<hp:cellMargin left="{m["left"]}" right="{m["right"]}" top="{m["top"]}" bottom="{m["bottom"]}"/>'
                    "</hp:tc>")
            trs.append("<hp:tr>" + "".join(tcs) + "</hp:tr>")

        none_bf = self.reg.border_fill({s: ("NONE", 0.1, "#000000") for s in ("left", "right", "top", "bottom")})
        self.oid += 1
        tw, th = sum(grid) * TW, sum(row_h)
        dm = {s: def_mar[s] * TW for s in def_mar}
        return (f'<hp:tbl id="{self.oid}" zOrder="{self.stats["tables"]}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
                f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" '
                f'repeatHeader="{int(bool(header_rows))}" rowCnt="{nrows}" colCnt="{ncols}" cellSpacing="0" '
                f'borderFillIDRef="{none_bf}" noAdjust="0">'
                f'<hp:sz width="{tw}" widthRelTo="ABSOLUTE" height="{th}" heightRelTo="ABSOLUTE" protect="0"/>'
                '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" '
                'vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
                '<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
                f'<hp:inMargin left="{dm["left"]}" right="{dm["right"]}" top="{dm["top"]}" bottom="{dm["bottom"]}"/>'
                + "".join(trs) + "</hp:tbl>")

    def table_host(self, tbl, page_break=False):
        """표를 담는 문단: 1pt 글자·줄간격 100% → 줄 높이 = 표 높이(여분 없음)."""
        ppid = self.reg.para_pr("LEFT", 0, 0, 0, 0, 0, "PERCENT", 100, False, False, 2, 0)
        cid = self.char_id((self.def_font, 2, False, "000000", None, 0, 100))
        t = self.table_xml(tbl)
        self.plain.append("")
        return (f'<hp:p id="{self.next_pid()}" paraPrIDRef="{ppid}" styleIDRef="0" '
                f'pageBreak="{int(page_break)}" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="{cid}">{t}<hp:t/></hp:run></hp:p>')

    # ── 구역 ──
    def sec_pr(self):
        sp = self.doc.find(f".//{q('sectPr')}")
        pg = sp.find(q("pgSz"))
        mg = sp.find(q("pgMar"))
        w, h = int(wattr(pg, "w")) * TW, int(wattr(pg, "h")) * TW
        if abs(w - 59528) < 20 and abs(h - 84186) < 20:
            w, h = 59528, 84186   # A4(HWP 표준값)
        g = lambda k: int(wattr(mg, k, "0")) * TW
        return (f'<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" '
                f'tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" '
                f'textVerticalWidthHead="0" masterPageCnt="0">'
                '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
                '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
                '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" '
                'fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
                '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
                f'<hp:pagePr landscape="WIDELY" width="{w}" height="{h}" gutterType="LEFT_ONLY">'
                # 머리말·꼬리말이 없는 문서: HWP는 위쪽 여백 + 머리말 아래에서 본문이 시작하므로 0으로 둬야
                # Word와 같은 위치(위 20mm·아래 15mm)에 본문이 놓인다
                f'<hp:margin header="0" footer="0" gutter="{g("gutter")}" left="{g("left")}" right="{g("right")}" '
                f'top="{g("top")}" bottom="{g("bottom")}"/></hp:pagePr>'
                '<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
                '<hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
                '<hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr>'
                '<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
                '<hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
                '<hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>'
                + "".join(
                    f'<hp:pageBorderFill type="{t}" borderFillIDRef="1" textBorder="PAPER" headerInside="0" '
                    'footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
                    for t in ("BOTH", "EVEN", "ODD"))
                + "</hp:secPr>"
                '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>')

    def convert(self):
        body = self.doc.find(q("body"))
        out = []
        first = True
        for el in body:
            if el.tag == q("p"):
                out.append(self.para_xml(el, first_extra=self.sec_pr() if first else ""))
                first = False
            elif el.tag == q("tbl"):
                if first:  # 문서가 표로 시작하면 구역 정보를 담을 빈 문단 먼저
                    out.append(self.para_xml(etree.Element(q("p")), first_extra=self.sec_pr()))
                    first = False
                out.append(self.table_host(el))
        section = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {NSMAP_DECL}>'
                   + "".join(out) + "</hs:sec>")
        return section.encode("utf-8"), self.reg.build()


def build(docx, out, skel_path=None, title="", preview_png=None):
    skel = Skeleton(skel_path)
    conv = Converter(docx, skel)
    section, header = conv.convert()
    hpf = skel.read("Contents/content.hpf").decode("utf-8")
    hpf = hpf.replace("<opf:title/>", f"<opf:title>{esc(title)}</opf:title>")
    hpf = re.sub(r'(<opf:meta name="(?:creator|lastsaveby)" content="text">)[^<]*', r"\1", hpf)
    hpf = re.sub(r'(<opf:meta name="(?:CreatedDate|ModifiedDate|date)" content="text">)[^<]*', r"\1", hpf)
    text = "\r\n".join(t for t in conv.plain if t.strip())
    files = [
        ("mimetype", skel.read("mimetype")),
        ("version.xml", skel.read("version.xml")),
        ("Contents/header.xml", header),
        ("Contents/section0.xml", section),
        ("Preview/PrvText.txt", text[:1024].encode("utf-8")),
        ("settings.xml", skel.read("settings.xml")),
        ("META-INF/container.rdf", skel.read("META-INF/container.rdf")),
        ("Contents/content.hpf", hpf.encode("utf-8")),
        ("META-INF/container.xml", skel.read("META-INF/container.xml")),
        ("META-INF/manifest.xml", skel.read("META-INF/manifest.xml")),
        ("Preview/PrvImage.png", Path(preview_png).read_bytes() if preview_png else skel.read("Preview/PrvImage.png")),
    ]
    with zipfile.ZipFile(out, "w") as z:
        for name, data in files:
            comp = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
            z.writestr(zipfile.ZipInfo(name, date_time=(2026, 9, 29, 0, 0, 0)), data, compress_type=comp)
    return conv


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="docx(HWP 서식 재현본) → hwpx")
    ap.add_argument("docx")
    ap.add_argument("hwpx")
    ap.add_argument("--title", default="")
    ap.add_argument("--preview", help="미리보기 PNG(PrvImage)")
    ap.add_argument("--skeleton", help="뼈대 hwpx(기본: python-hwpx 내장)")
    a = ap.parse_args()
    c = build(a.docx, a.hwpx, a.skeleton, a.title, a.preview)
    print("stats", c.stats, "charPr", len(c.reg.char_xml), "paraPr", len(c.reg.para_xml),
          "borderFill", len(c.reg.bf_xml), "fonts", c.reg.fonts)
