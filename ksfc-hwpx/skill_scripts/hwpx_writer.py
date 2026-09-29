#!/usr/bin/env python3
"""IR(build_hwpx.js가 기록한 문서 구조) → .hwpx

원본 HWP 51건 실측값(references/hwp-native.md)을 기준으로 한글 고유 방식으로 쓴다.
  - 줄간격: 글자에 따라 %(본문 15pt × 160% = 24pt). 간격은 작은 글자 빈 줄(문단 위·아래 간격 0)
  - 층위: 앞 공백 + 기호 + 공백, 내어쓰기 = 본문 첫 글자 위치(탭 없음)
  - 띄어쓰기 0.5em(글꼴에 어울리는 빈칸 끔), 양쪽 정렬, 최소 공백 0%
  - 줄 나눔: 어절 단위 기본. 띄어쓰기가 0.5em 넘게 벌어지는 줄이 생기면 자간(-10% 한도)으로 다음 어절을
    끌어올리고, 그래도 안 되면 그 문단만 글자 단위 + 자간 보정(쪼개지는 자리를 자연스럽게)
  - 표: 글자처럼 취급, 셀 단위 쪽 나눔, 제목 행 반복. 선 0.12mm(안쪽)·0.5mm(굵은 선)
  - 키워드 형광: 괄호를 뺀 글자에만 음영(원본 85%)
사용: python3 hwpx_writer.py doc.ir.json out.hwpx [--rm-ir]
"""
import base64
import io
import json
import os
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwp_metrics as HM  # noqa: E402

NS = ('xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
      'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
      'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
      'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
      'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" '
      'xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
      'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" '
      'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"')
HUC = "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"
LANGS = ["hangul", "latin", "hanja", "japanese", "other", "symbol", "user"]
TW = 5                       # twip → HWPUNIT
PAGE_W, PAGE_H = 59528, 84188
MM = 7200 / 25.4             # mm → HWPUNIT
# 원본 쪽 여백(51건 최빈값): 위 10 + 머리말 10, 아래 10 + 꼬리말 10, 좌우 25mm
PAGE = dict(top=round(10 * MM), header=round(10 * MM), bottom=round(10 * MM), footer=round(10 * MM),
            left=round(25 * MM), right=round(25 * MM))
BODY_W = PAGE_W - PAGE["left"] - PAGE["right"]
# docx용 색(쪽 이미지 픽셀 실측) → 원본 설정값
COLOR_MAP = {"DCF6DD": "D8FFD8", "FCF6CC": "FFF7CC", "FEF7CD": "FFF7CC"}
# 쪽 맞춤 단계(스킬 순서: 자간 → 간격 → 줄간격). 구간(쪽 나누기 사이)의 마지막 쪽이 거의 비면 한 단계씩 올림
FIT_LEVELS = [dict(sp=1.0, ls=None), dict(sp=0.9, ls=None), dict(sp=0.8, ls=None), dict(sp=0.9, ls=150), dict(sp=0.8, ls=150), dict(sp=0.7, ls=150)]
PULL_FILL = 0.5              # 마지막 줄이 한 줄의 50% 이하면 자간(-10% 한도)으로 끌어올림(스킬 기존 규칙)
LOOSE_EM = 0.5               # 양쪽 정렬로 띄어쓰기가 이만큼(em) 넘게 벌어지면 보정
MAX_TIGHT = -10              # 자간 보정 한도(%)
SKELETON = None


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def color(c, default="000000"):
    if not c or c in ("auto",):
        return default
    c = c.upper().lstrip("#")
    if c == "WHITE":
        return "FFFFFF"
    return COLOR_MAP.get(c, c)


def border_mm(sz):
    """docx 선 굵기(1/8pt) → 원본에서 쓰는 한글 선 굵기."""
    sz = sz or 4
    if sz <= 8:
        return "0.12 mm"
    if sz <= 14:
        return "0.5 mm"
    if sz <= 20:
        return "0.7 mm"
    return "1.0 mm"


def border_of(b):
    if not b or b.get("style") in (None, "none", "nil") or not b.get("size", 1):
        return ("NONE", "0.1 mm", "#000000")
    st = {"single": "SOLID", "dotted": "DASH", "dashed": "DASH", "double": "DOUBLE_SLIM"}.get(b["style"], "SOLID")
    return (st, border_mm(b.get("size")), "#" + color(b.get("color")))


# ───────────────────────── 헤더(모양 목록) ─────────────────────────
class Styles:
    def __init__(self, skel_header):
        from lxml import etree
        self.etree = etree
        self.root = etree.fromstring(skel_header)
        rl = self.root.find("{*}refList")
        self.base_font = len(rl.find("{*}fontfaces")[0])
        self.base_bf = len(rl.find("{*}borderFills")) + 1
        self.base_char = len(rl.find("{*}charProperties"))
        self.base_para = len(rl.find("{*}paraProperties"))
        self.base_tab = len(rl.find("{*}tabProperties"))
        self.fonts, self.chars, self.paras, self.bfs, self.tabs = [], {}, {}, {}, {}
        self.char_x, self.para_x, self.bf_x, self.tab_x = [], [], [], []

    def font_id(self, face):
        if face not in self.fonts:
            self.fonts.append(face)
        return self.base_font + self.fonts.index(face)

    def border_fill(self, sides, fill=None):
        key = (tuple(sorted(sides.items())), fill)
        if key not in self.bfs:
            bid = self.base_bf + len(self.bf_x)
            x = [f'<hh:borderFill id="{bid}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
                 '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>']
            for s in ("left", "right", "top", "bottom"):
                t, w, c = sides.get(s, ("NONE", "0.1 mm", "#000000"))
                x.append(f'<hh:{s}Border type="{t}" width="{w}" color="{c}"/>')
            x.append('<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>')
            if fill:
                x.append(f'<hc:fillBrush><hc:winBrush faceColor="#{fill}" hatchColor="#999999" alpha="0"/></hc:fillBrush>')
            x.append("</hh:borderFill>")
            self.bf_x.append("".join(x))
            self.bfs[key] = bid
        return self.bfs[key]

    def none_bf(self):
        return self.border_fill({s: ("NONE", "0.1 mm", "#000000") for s in ("left", "right", "top", "bottom")})

    def char(self, cs):
        key = cs.key()
        if key not in self.chars:
            cid = self.base_char + len(self.char_x)
            fid = self.font_id(cs.face)
            a = lambda v: " ".join(f'{l}="{v}"' for l in LANGS)
            x = (f'<hh:charPr id="{cid}" height="{round(cs.pt * 100)}" textColor="#{cs.color}" '
                 f'shadeColor="{("#" + cs.shade) if cs.shade else "none"}" useFontSpace="0" useKerning="0" '
                 f'symMark="NONE" borderFillIDRef="2"><hh:fontRef {a(fid)}/><hh:ratio {a(cs.ratio)}/>'
                 f'<hh:spacing {a(cs.spacing)}/><hh:relSz {a(100)}/><hh:offset {a(0)}/>'
                 + ("<hh:bold/>" if cs.bold else "")
                 + f'<hh:underline type="{"BOTTOM" if cs.underline else "NONE"}" shape="SOLID" color="#{cs.color}"/>'
                 + f'<hh:strikeout shape="{"CONTINUOUS" if cs.strike else "NONE"}" color="#{cs.color}"/>'
                 '<hh:outline type="NONE"/><hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
                 + ("<hh:supscript/>" if cs.sup else "") + "</hh:charPr>")
            self.char_x.append(x)
            self.chars[key] = cid
        return self.chars[key]

    def tab(self, items):
        """items: [(pos_hu, 'LEFT'|'RIGHT', leader)]"""
        key = tuple(items)
        if key not in self.tabs:
            tid = self.base_tab + len(self.tab_x)
            body = "".join(
                f'<hp:switch><hp:case hp:required-namespace="{HUC}"><hh:tabItem pos="{p}" type="{t}" leader="{l}" unit="HWPUNIT"/></hp:case>'
                f'<hp:default><hh:tabItem pos="{p * 2}" type="{t}" leader="{l}"/></hp:default></hp:switch>' for p, t, l in items)
            self.tab_x.append(f'<hh:tabPr id="{tid}" autoTabLeft="0" autoTabRight="0">{body}</hh:tabPr>')
            self.tabs[key] = tid
        return self.tabs[key]

    def para(self, ps):
        key = ps.key()
        if key not in self.paras:
            pid = self.base_para + len(self.para_x)

            def margin(m):
                return ("<hh:margin>"
                        f'<hc:intent value="{ps.intent * m}" unit="HWPUNIT"/><hc:left value="{ps.left * m}" unit="HWPUNIT"/>'
                        f'<hc:right value="{ps.right * m}" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/>'
                        f'<hc:next value="0" unit="HWPUNIT"/></hh:margin>'
                        f'<hh:lineSpacing type="{ps.ls_type}" value="{ps.ls_val if ps.ls_type == "PERCENT" else ps.ls_val * m}" unit="HWPUNIT"/>')
            x = (f'<hh:paraPr id="{pid}" tabPrIDRef="{ps.tab}" condense="0" fontLineHeight="0" snapToGrid="0" '
                 'suppressLineNumbers="0" checked="0" textDir="LTR">'
                 f'<hh:align horizontal="{ps.align}" vertical="BASELINE"/><hh:heading type="NONE" idRef="0" level="0"/>'
                 f'<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="{"KEEP_WORD" if ps.charwrap else "BREAK_WORD"}" '
                 f'widowOrphan="0" keepWithNext="{int(ps.keep_next)}" keepLines="{int(ps.keep_lines)}" pageBreakBefore="0" lineWrap="BREAK"/>'
                 '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
                 f'<hp:switch><hp:case hp:required-namespace="{HUC}">{margin(1)}</hp:case><hp:default>{margin(2)}</hp:default></hp:switch>'
                 f'<hh:border borderFillIDRef="{ps.border}" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
                 "</hh:paraPr>")
            self.para_x.append(x)
            self.paras[key] = pid
        return self.paras[key]

    def build(self):
        et = self.etree
        rl = self.root.find("{*}refList")
        wrap = lambda s: et.fromstring(f"<x {NS}>{s}</x>")[0]
        for ff in rl.find("{*}fontfaces"):
            for i, face in enumerate(self.fonts):
                fam = "FCAT_MYUNGJO" if "명조" in face else "FCAT_GOTHIC"
                ff.append(wrap(f'<hh:font id="{self.base_font + i}" face="{esc(face)}" type="TTF" isEmbedded="0">'
                               f'<hh:typeInfo familyType="{fam}" weight="6" proportion="4" contrast="0" strokeVariation="1" '
                               'armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font>'))
            ff.set("fontCnt", str(len(ff)))
        for name, xs in (("borderFills", self.bf_x), ("charProperties", self.char_x),
                         ("tabProperties", self.tab_x), ("paraProperties", self.para_x)):
            box = rl.find("{http://www.hancom.co.kr/hwpml/2011/head}" + name)
            for x in xs:
                box.append(wrap(x))
            box.set("itemCnt", str(len(box)))
        return et.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


class CS:
    """글자 모양"""
    __slots__ = ("face", "pt", "bold", "color", "shade", "spacing", "ratio", "underline", "strike", "sup")

    def __init__(self, face="나눔명조", pt=15.0, bold=False, color="000000", shade=None, spacing=0, ratio=100,
                 underline=False, strike=False, sup=False):
        self.face, self.pt, self.bold, self.color, self.shade = face, pt, bold, color, shade
        self.spacing, self.ratio, self.underline, self.strike, self.sup = spacing, ratio, underline, strike, sup

    def key(self):
        return tuple(getattr(self, k) for k in self.__slots__)

    def copy(self, **kw):
        c = CS(*[getattr(self, k) for k in self.__slots__])
        for k, v in kw.items():
            setattr(c, k, v)
        return c

    def adv(self, ch, extra=0):
        sp = max(-50, min(50, self.spacing + extra))
        a = HM.advance(ch, self.face, self.bold, self.pt, self.ratio, sp)
        return a * 0.6 if self.sup else a


class PS:
    """문단 모양"""
    __slots__ = ("align", "left", "intent", "right", "ls_type", "ls_val", "keep_next", "keep_lines", "border", "tab", "charwrap")

    def __init__(self, **kw):
        self.align, self.left, self.intent, self.right = "JUSTIFY", 0, 0, 0
        self.ls_type, self.ls_val, self.keep_next, self.keep_lines = "PERCENT", 160, False, False
        self.border, self.tab, self.charwrap = 2, 0, False
        for k, v in kw.items():
            setattr(self, k, v)

    def key(self):
        for k in ("left", "intent", "right", "ls_val"):
            setattr(self, k, int(round(getattr(self, k))))
        return tuple(getattr(self, k) for k in self.__slots__)


class Para:
    def __init__(self, ps, runs, width, page_break=False, hints=()):
        self.ps, self.runs, self.width, self.page_break = ps, runs, width, page_break   # runs: [[CS, text]] 또는 ("obj", xml)
        self.prefix = None        # (글자 모양, 앞 공백+기호+공백, 앞 공백 수) — 빠른 내어쓰기 계산용
        self.hints = set(hints)   # 자연스러운 분리 위치(글자 번호 앞) — build_hwpx.js의 U+200B 자리
        self.delta = 0            # 문단 자간 보정(%)

    def text(self):
        return "".join(t for cs, t in self.runs if cs != "obj")

    def chars(self, delta=None):
        d = self.delta if delta is None else delta
        out = []
        for cs, t in self.runs:
            if cs == "obj":
                continue
            for ch in t:
                out.append((ch, cs.adv(ch, d)))
        return out

    def widths(self):
        w = self.width - self.ps.left - self.ps.right
        return w - max(self.ps.intent, 0), w - max(-self.ps.intent, 0)


# ───────────────────────── 줄 맞춤(작성자 결정 규칙) ─────────────────────────
FIT_MARGIN = 0.01            # 줄 폭 여유(1%): 51건 한글 저장본 1,227개 줄 나눔 중 이 여유로 부족했던 곳 1건(98.5%) — 폭 추정 오차로 마지막 어절이 다시 넘어가지 않게


def _loosest(p, delta, mode, margin=0.0):
    fw, rw = p.widths()
    fw, rw = fw * (1 - margin), rw * (1 - margin)
    ch = p.chars(delta)
    starts, used, gaps = HM.layout(ch, fw, rw, mode)
    em = max(cs.pt for cs, t in p.runs if cs != "obj" and t.strip()) * 100
    worst = 0.0
    for i in range(len(used) - 1):        # 마지막 줄은 양쪽 정렬 대상이 아님
        w = (fw if i == 0 else rw) / (1 - margin)
        worst = max(worst, (w - used[i]) / max(1, gaps[i]) / em)
    return worst, starts, used


def _awkward_splits(p, starts):
    """어절 중간 줄 바뀜 중 자연스러운 경계(복합어 사이·가운뎃점 뒤 등, build_hwpx.js가 표시한 자리)가 아닌 것의 수."""
    t = p.text()
    bad = 0
    for s in starts[1:]:
        if 0 < s < len(t) and t[s - 1] != " " and t[s] != " ":
            if s not in p.hints:
                bad += 1
    return bad


def _lo(p):
    """자간 합계 -10% 한도(직접 지정한 {tight}가 더 크면 그 값)에서 이 문단이 더 줄일 수 있는 최소 보정값"""
    sps = [cs.spacing for cs, t in p.runs if cs != "obj" and t.strip()]
    return min(0, MAX_TIGHT - min(0, min(sps))) if sps else 0


def _char_stable(p, d):
    """글자 단위 줄 나눔이 줄 폭 ±1%(FIT_MARGIN)에서도 같은 자연스러운 자리에서만 끊기는지.
    한글이 계산보다 한 글자 더 넣으면(9차 저장본 `ㆍ|자금세탁` → `자|금세탁`) 어절 중간 끊김이 되므로 양쪽 모두 확인."""
    s0 = _loosest(p, d, "char")[1]
    return (_awkward_splits(p, s0) == 0 and s0 == _loosest(p, d, "char", FIT_MARGIN)[1]
            and s0 == _loosest(p, d, "char", -FIT_MARGIN)[1])


def _finalize(p):
    """내어쓰기까지 정한 최종 상태 점검: 글자 단위가 어절 중간(자연스러운 경계 아님)에서 끊기면 어절 단위로 되돌리고,
    줄 폭 여유가 부족하면 자간을 한 단계씩 더 줄임(-10% 한도). 자간을 줄이면 내어쓰기도 작아지므로 수렴."""
    lo = _lo(p)
    for _ in range(12):
        mode = "char" if p.ps.charwrap else "word"
        if p.ps.charwrap and not _char_stable(p, p.delta):
            p.ps.charwrap = False
            apply_quick_indent(p)
            continue
        if _loosest(p, p.delta, mode, FIT_MARGIN)[1] == _loosest(p, p.delta, mode)[1] or p.delta - 1 < lo:
            break
        p.delta -= 1
        apply_quick_indent(p)


def fit(p, stats):
    """어절 단위가 원칙(작성자 결정). 띄어쓰기가 0.5em 넘게 벌어지는 줄이 생기면 자간(-10% 한도)으로 다음 어절을 끌어올리고,
    그래도 안 되면 자연스러운 경계에서만 글자 단위로 나눔(없으면 어절 단위 유지). 끌어올릴 때는 줄 폭 1% 여유(FIT_MARGIN)를 남김."""
    if not p.text().strip() or any(cs == "obj" for cs, _ in p.runs) or "\t" in p.text():
        return
    base_sp = min(cs.spacing for cs, t in p.runs if cs != "obj" and t.strip())
    lo = min(0, MAX_TIGHT - min(0, base_sp))   # 자간 합계가 -10%를 넘지 않게(직접 지정한 {tight}는 존중)
    p.delta, p.ps.charwrap = 0, False
    if p.ps.align != "JUSTIFY":                  # 가운데·왼쪽 정렬(표 칸 등)은 넘침 여유만 확보
        if "\n" not in p.text():
            _settle(p, lo)
        return
    worst, starts, used = _loosest(p, 0, "word")
    if len(starts) < 2:
        _settle(p, lo)
        return
    if worst > LOOSE_EM:
        best = None
        for d in range(-1, lo - 1, -1):
            w2, s2, _ = _loosest(p, d, "word", FIT_MARGIN)
            if w2 <= LOOSE_EM:
                best = d
                break
        if best is not None:
            p.delta = best
            stats["tight"] += 1
        else:
            # 자연스러운 경계에서만 글자 단위 허용
            for d in sorted(range(lo, 1), key=abs):
                w2, s2, _ = _loosest(p, d, "char", FIT_MARGIN)
                # 여유를 둔 폭과 실제 폭에서 같은 자리로 끊길 때만(폭 오차로 끊김 자리가 어절 중간으로 밀리지 않게)
                if (_awkward_splits(p, s2) == 0 and len(s2) > 1 and _char_stable(p, d)
                        and s2 != _loosest(p, d, "word", FIT_MARGIN)[1]):
                    p.ps.charwrap, p.delta = True, d
                    stats["charwrap"] += 1
                    return
            # 없으면 어절 단위 유지, 벌어짐이 가장 작은 자간
            cands = [(_loosest(p, d, "word", FIT_MARGIN)[0], -d, d) for d in range(0, lo - 1, -1)]
            p.delta = min(cands)[2]
            if p.delta:
                stats["tight"] += 1
    # 마지막 줄에 짧은 조각만 남으면 끌어올리기(원본 관행: 자간을 줄여 한 줄 줄임), 1% 여유 확보
    worst, starts, used = _loosest(p, p.delta, "word")
    fw, rw = p.widths()
    if len(starts) >= 2 and used[-1] <= PULL_FILL * rw:
        for d in range(p.delta - 1, lo - 1, -1):
            w2, s2, _ = _loosest(p, d, "word", FIT_MARGIN)
            if len(s2) < len(starts) and w2 <= LOOSE_EM:
                p.delta = d
                stats["pullup"] += 1
                break
    _settle(p, lo)


def _settle(p, lo):
    """줄이 여유 없이 딱 맞으면(줄 폭 99% 넘게 채움) 한글에서는 끝 어절이 넘어가기도 함(4차 한글 저장본 416문단 중 1건,
    99.9%) → 여유(FIT_MARGIN)를 두어도 줄 나눔이 같아질 때까지 자간을 한 단계씩 줄임(-10% 한도). 자간을 따로 조정하지 않은 문단도 포함."""
    while p.delta - 1 >= lo:
        _, s_m, _ = _loosest(p, p.delta, "word", FIT_MARGIN)
        _, s_0, _ = _loosest(p, p.delta, "word")
        if s_m == s_0:
            break
        p.delta -= 1


def fit_single_lines(p, stats):
    """제목·소제목: 한 줄로 쓴 줄(강제 줄바꿈 사이)이 조금 넘치면 자간(-10% 한도)으로 한 줄에 맞춤.
    그래도 넘치면 장평을 95%·90%까지 줄임(원본 제목에서 쓰는 방식)."""
    width = p.width - p.ps.left - p.ps.right
    segs, cur = [], []
    for cs, t in p.runs:
        if cs == "obj":
            continue
        for ch in t:
            if ch == "\n":
                segs.append(cur); cur = []
            else:
                cur.append((cs, ch))
        # 조각 끝
    segs.append(cur)

    def over(delta, ratio):
        # HY 글꼴 폭은 근사값이라 2% 여유
        return max(sum(cs.copy(ratio=min(cs.ratio, ratio)).adv(ch, delta) for cs, ch in sg) for sg in segs if sg) > width * 0.98

    if not any(segs) or not over(0, 100):
        return
    for ratio in (100, 95, 90):
        for d in range(0, MAX_TIGHT - 1, -1):
            if not over(d, ratio):
                p.delta = d
                if ratio < 100:
                    for r in p.runs:
                        if r[0] != "obj":
                            r[0] = r[0].copy(ratio=min(r[0].ratio, ratio))
                stats["tight"] += 1
                return


# ───────────────────────── IR → 문단 ─────────────────────────
WIDE_MARK_EM = 1.04   # 한글이 전각 기호(□ ㅇ ➊ ① ※ ◈ ■ ▶ 등)를 그리는 폭 — 원본 내어쓰기 역산값


def prefix_width(cs, prefix):
    """앞 공백 + 기호 + 공백의 폭(= 내어쓰기, 한글에서 Shift+Tab으로 잡는 위치)."""
    return round(sum(cs.adv(c) for c in prefix))


def apply_quick_indent(p):
    """빠른 내어쓰기(Shift+Tab)와 같은 위치: 기호 뒤 커서의 실제 x.
    첫 줄이 양쪽 정렬로 늘어나면 기호 뒤 공백도 그만큼 늘어남(앞 공백은 안 늘어남) — 한글 보정 912문단 검증."""
    if not getattr(p, "prefix", None):
        return
    cs0, prefix, nlead = p.prefix
    base = sum(cs0.adv(c, p.delta) for c in prefix)
    extra = 0.0
    if p.ps.align == "JUSTIFY":
        fw, rw = p.widths()
        starts, used, gaps = HM.layout(p.chars(), fw, rw, "char" if p.ps.charwrap else "word")
        if len(starts) > 1:
            t = p.text()
            line = t[:starts[1]].rstrip(" ")
            ngap = line.count(" ") - nlead
            n_after = prefix.count(" ") - nlead
            if ngap > 0:
                extra = (fw - used[0]) / ngap * n_after
    p.ps.intent = -round(base + extra)

MARKER_RE = re.compile(r"^([□ㅇ\-·∙•■▪‣▶◆◈◇○●◎▷►※➡⇨⇒☞]|\*\*|\*|\d\)|[➊-➓①-⑳❶-❿])$")


class Writer:
    def __init__(self, ir, skel, levels=None):
        self.ir, self.skel = ir, skel
        self.levels = levels or {}   # 구간 번호 → FIT_LEVELS 단계
        self.block = 0
        self.st = Styles(skel("Contents/header.xml"))
        self.stats = {"tight": 0, "charwrap": 0, "pullup": 0, "paras": 0, "tables": 0, "images": 0}
        self.bins = []            # (id, ext, bytes)
        self.pid = 0
        self.oid = 1000

    # 글자 모양
    def cs_of(self, r):
        f = (r.get("font") or {}).get("name") or "나눔명조"
        size = (r.get("size") or 30) / 2
        cs = CS(face=f, pt=size, bold=bool(r.get("bold")) and not f.startswith("HY"), color=color(r.get("color")),
                shade=color(r["shading"]["fill"]) if r.get("shading") and r["shading"].get("fill") not in (None, "auto") else None,
                ratio=int(r.get("scale") or 100), underline=bool(r.get("underline") is not None and r.get("underline") is not False),
                strike=bool(r.get("strike")), sup=bool(r.get("superScript")))
        if r.get("characterSpacing"):
            glyph = size * HM.em(f, cs.bold, "가") * cs.ratio / 100
            cs.spacing = max(-50, min(50, round(r["characterSpacing"] / 20 / glyph * 100)))
        return cs

    def runs_of(self, children, width):
        """IR 문단 자식 → [[CS, text] | ("obj", xml)], 분리 힌트 위치"""
        runs, hints, pos = [], [], 0
        for r in children:
            if r["type"] == "r":
                cs = self.cs_of(r)
                t = r.get("text") or ""
                if r.get("break"):
                    runs.append((cs, "\n"))
                    pos += 1
                t = t.replace("︎", "").replace(" ", " ")
                clean = ""
                for ch in t:
                    if ch == "​":
                        hints.append(pos)
                        continue
                    clean += ch
                    pos += 1
                if clean:
                    runs.append([cs, clean])
            elif r["type"] == "img":
                x = self.pic_xml(r, width)
                if x:
                    runs.append(("obj", x))
        return runs, hints

    def split_paren_shade(self, runs):
        """형광 키워드 '(키워드)'는 괄호를 뺀 글자에만 음영(원본 85%)."""
        out = []
        for cs, t in runs:
            if cs != "obj" and cs.shade and len(t) > 2 and t.startswith("(") and t.endswith(")"):
                plain = cs.copy(shade=None)
                out += [[plain, "("], [cs, t[1:-1]], [plain, ")"]]
            else:
                out.append([cs, t])
        return out

    def lv(self):
        return FIT_LEVELS[self.levels.get(self.block, 0)]

    def spacer(self, gap_pt, width, keep_next=False, page_break=False, face="나눔명조"):
        """간격용 빈 줄: 원본처럼 작은 글자 + 160% (10pt 빈 줄 = 16pt 간격)."""
        gap_pt *= self.lv()["sp"]
        if gap_pt < 3:
            ps = PS(align="LEFT", ls_type="FIXED", ls_val=max(100, round(gap_pt * 100)), keep_next=keep_next)
            cs = CS(face=face, pt=max(1.0, gap_pt))
        else:
            ps = PS(align="LEFT", ls_type="PERCENT", ls_val=160, keep_next=keep_next)
            cs = CS(face=face, pt=round(gap_pt / 1.6 * 2) / 2)
        return Para(ps, [[cs, ""]], width, page_break)

    def convert_p(self, p, width):
        """IR 문단 → [Para] (앞뒤 간격 빈 줄 포함)"""
        runs, hints = self.runs_of(p.get("children") or [], width)
        runs = self.split_paren_shade(runs)
        text_runs = [(cs, t) for cs, t in runs if cs != "obj" and t.strip()]
        sp = p.get("spacing") or {}
        before, after = (sp.get("before") or 0) / 20, (sp.get("after") or 0) / 20
        line, rule = sp.get("line"), sp.get("lineRule")
        ind = p.get("indent") or {}
        left, hanging, right = (ind.get("left") or 0), (ind.get("hanging") or 0), (ind.get("right") or 0)
        align = {"both": "JUSTIFY", "center": "CENTER", "right": "RIGHT", "left": "LEFT"}.get(p.get("alignment"), "LEFT")
        page_break = bool(p.get("pageBreakBefore"))
        out = []
        keep_next = bool(p.get("keepNext"))
        face0 = runs[0][0].face if runs and runs[0][0] != "obj" else "나눔명조"

        if p.get("_spacer") or (not text_runs and not any(cs == "obj" for cs, _ in runs)):
            gap = (int(line) / 20 if line else 6) + before + after
            return [self.spacer(gap, width, keep_next, page_break, face0)]

        if before >= 1:
            out.append(self.spacer(before, width, True, page_break, face0))
            page_break = False

        # 줄간격: 글자 크기 대비 %
        max_pt = max(cs.pt for cs, _ in text_runs) if text_runs else 15
        if line and rule == "exact":
            ls_type, ls_val = "PERCENT", round(int(line) / 20 / max_pt * 100)
            if self.lv()["ls"] and ls_val >= 155 and align == "JUSTIFY":   # 쪽 맞춤: 본문 160% → 150%(원본 38%가 150%)
                ls_val = self.lv()["ls"]
        elif line and rule == "atLeast":
            ls_type, ls_val = "AT_LEAST", int(line) * TW
        elif line:
            ls_type, ls_val = "PERCENT", round(int(line) / 240 * 115)
        else:
            ls_type, ls_val = "PERCENT", 115 if not text_runs else 160

        ps = PS(align=align, right=right * TW, ls_type=ls_type, ls_val=ls_val, keep_next=keep_next,
                keep_lines=bool(p.get("keepLines")))

        # 기호 뒤 탭 + 내어쓰기 → 앞 공백 + 기호 + 공백 + 내어쓰기(원본 방식)
        quick = None
        start_tw = left - hanging
        first = runs[0] if runs else None
        if first and first[0] != "obj" and "\t" in first[1] and hanging > 0 and first[1].index("\t") <= 3:
            mk = first[1][:first[1].index("\t")].strip()
            cs0 = first[0]
            nsp = max(0, round(start_tw / 20 / (0.5 * cs0.pt)))
            tab_at = first[1].index("\t")
            first[1] = " " * nsp + mk + " " + first[1][tab_at + 1:]
            # 탭까지(tab_at + 1글자)가 앞 공백 + 기호 + 공백(nsp + len(mk) + 1글자)으로 바뀐 만큼 힌트 위치 이동
            hints = [h + nsp + len(mk) - tab_at for h in hints]
            prefix = " " * nsp + mk + " "
            ps.left = 0
            ps.intent = -prefix_width(cs0, prefix)
            quick = (cs0, prefix, nsp)
        elif hanging > 0:
            # 탭 없는 내어쓰기(표 셀 목록 등): 첫 글자 기호 + 공백까지를 내어쓰기로
            txt = "".join(t for cs, t in runs if cs != "obj")
            m = re.match(r"^(\s*\S{1,2}\s+)", txt)
            cs0 = runs[0][0] if runs and runs[0][0] != "obj" else CS()
            nsp = max(0, round(start_tw / 20 / (0.5 * cs0.pt)))
            if nsp and runs and runs[0][0] != "obj":
                runs[0][1] = " " * nsp + runs[0][1]
                hints = [h + nsp for h in hints]
            ps.left = 0
            prefix = " " * nsp + (m.group(1) if m else "")
            ps.intent = -prefix_width(cs0, prefix) if m else -hanging * TW
            if m:
                quick = (cs0, prefix, len(prefix) - len(prefix.lstrip(" ")))
        else:
            ps.left = max(0, left) * TW
            ps.intent = (ind.get("firstLine") or 0) * TW

        # 오른쪽 탭(꼬리말·목차 쪽번호)
        tabs = p.get("tabStops") or []
        if tabs:
            items = []
            for t in tabs:
                pos = max(0, int(t["position"]) * TW - ps.left)
                items.append((pos, "RIGHT" if t.get("type") == "right" else "LEFT", "DOT" if t.get("leader") == "dot" else "NONE"))
            ps.tab = self.st.tab(items)

        # 문단 테두리·배경(제목 띠, 소제목)
        pb, sh = p.get("border"), p.get("shading")
        if pb or (sh and sh.get("fill") not in (None, "auto")):
            sides = {s: border_of((pb or {}).get(s)) for s in ("left", "right", "top", "bottom")}
            fill = color(sh["fill"]) if sh and sh.get("fill") not in (None, "auto") else None
            ps.border = self.st.border_fill(sides, fill)

        para = Para(ps, runs, width, page_break, hints)
        para.prefix = quick
        out.append(para)
        if after >= 1:
            out.append(self.spacer(after, width, keep_next, False, face0))
        return out

    # ── 제목 띠·소제목: 원본처럼 1칸 표 + 세로 가운데 ──
    def band_para(self, p, width):
        kind = p.get("_band")
        runs, _ = self.runs_of(p.get("children") or [], width)
        runs = [[cs, t.lstrip(" ") if i == 0 else t] for i, (cs, t) in enumerate(runs) if cs != "obj"]
        pb, sh = p.get("border") or {}, p.get("shading")
        sides = {s: border_of(pb.get(s)) for s in ("left", "right", "top", "bottom")}
        fill = color(sh["fill"]) if sh and sh.get("fill") not in (None, "auto") else None
        bf = self.st.border_fill(sides, fill)
        align = {"center": "CENTER", "right": "RIGHT"}.get(p.get("alignment"), "LEFT")
        pad_lr, pad_tb = 510, 141
        line_w = sum(cs.adv(ch) for cs, t in runs for ch in t if ch != "\n")
        if kind == "sub":
            tw = min(width, round(line_w + 2 * pad_lr + 4 * MM))
        elif kind == "band":
            tw = round(width * 0.88)
        else:
            tw = width
        min_h = {"title": 3079, "band": 3300, "sub": 2600}[kind]
        ps = PS(align=align, ls_type="PERCENT", ls_val=100, keep_next=True)
        inner = Para(ps, runs, tw - 2 * pad_lr)
        fit_single_lines(inner, self.stats)
        cell = self.para_xml(inner)
        self.oid += 1
        tbl = (f'<hp:tbl id="{self.oid}" zOrder="{self.oid - 1000}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" '
               f'lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" rowCnt="1" colCnt="1" cellSpacing="0" '
               f'borderFillIDRef="{self.st.none_bf()}" noAdjust="0"><hp:sz width="{tw}" widthRelTo="ABSOLUTE" height="{min_h}" heightRelTo="ABSOLUTE" protect="0"/>'
               '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
               'horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
               f'<hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:inMargin left="{pad_lr}" right="{pad_lr}" top="{pad_tb}" bottom="{pad_tb}"/>'
               f'<hp:tr><hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="{bf}">'
               '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" '
               f'textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">{cell}</hp:subList><hp:cellAddr colAddr="0" rowAddr="0"/>'
               f'<hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="{tw}" height="{min_h}"/>'
               f'<hp:cellMargin left="{pad_lr}" right="{pad_lr}" top="{pad_tb}" bottom="{pad_tb}"/></hp:tc></hp:tr></hp:tbl>')
        host = Para(PS(align="LEFT", ls_type="PERCENT", ls_val=100, keep_next=True), [[CS(pt=1.0), ""], ("obj", tbl)], width,
                    bool(p.get("pageBreakBefore")))
        sp = p.get("spacing") or {}
        out = []
        if (sp.get("before") or 0) >= 20:
            out.append(self.spacer(sp["before"] / 20, width, True, host.page_break))
            host.page_break = False
        out.append(host)
        if (sp.get("after") or 0) >= 20:
            out.append(self.spacer(sp["after"] / 20, width, True))
        return out

    # ── 표 ──
    def table_xml(self, t, width):
        self.stats["tables"] += 1
        cols = [int(c) * TW for c in (t.get("columnWidths") or [])]
        rows = t.get("rows") or []
        tb = t.get("borders") or {}
        grid, cells = {}, []
        for ri, tr in enumerate(rows):
            ci = 0
            for tc in tr.get("children") or []:
                while (ri, ci) in grid:
                    ci += 1
                span = int(tc.get("columnSpan") or 1)
                vm = tc.get("verticalMerge")
                if vm == "continue" and (ri - 1, ci) in grid:
                    top = grid[(ri - 1, ci)]
                    top["rspan"] += 1
                    for k in range(span):
                        grid[(ri, ci + k)] = top
                else:
                    cell = dict(row=ri, col=ci, span=span, rspan=1, tc=tc)
                    cells.append(cell)
                    for k in range(span):
                        grid[(ri, ci + k)] = cell
                ci += span
        nrows, ncols = len(rows), max(1, len(cols))
        if not cols:
            cols = [width]
        header_rows = {ri for ri, tr in enumerate(rows) if tr.get("tableHeader")}
        min_h = {}
        for ri, tr in enumerate(rows):
            h = tr.get("height")
            if h and h.get("value"):
                min_h[ri] = int(h["value"]) * TW
        row_h = [0] * nrows
        body = []
        for cell in cells:
            tc, ri, ci = cell["tc"], cell["row"], cell["col"]
            lr, lc = ri + cell["rspan"] - 1, ci + cell["span"] - 1
            cb = tc.get("borders") or {}
            sides = {}
            for s, edge, inner in (("top", ri == 0, "insideHorizontal"), ("bottom", lr == nrows - 1, "insideHorizontal"),
                                   ("left", ci == 0, "insideVertical"), ("right", lc == ncols - 1, "insideVertical")):
                b = cb.get(s)
                if b is None:
                    b = tb.get(s) if edge else tb.get(inner)
                sides[s] = border_of(b)
            sh = tc.get("shading")
            fill = color(sh["fill"]) if sh and sh.get("fill") not in (None, "auto") else None
            bf = self.st.border_fill(sides, fill)
            mg = tc.get("margins") or {}
            m = {s: int(mg.get(s, 108 if s in ("left", "right") else 0)) * TW for s in ("left", "right", "top", "bottom")}
            if tc.get("_md"):
                m["left"] = m["right"] = 510
                m["top"] = m["bottom"] = max(141, m["top"])
            va = {"center": "CENTER", "bottom": "BOTTOM"}.get(tc.get("verticalAlign"), "TOP")
            w = sum(cols[ci:ci + cell["span"]])
            inner_w = w - m["left"] - m["right"]
            paras = self.blocks(tc.get("children") or [], inner_w)
            xml = "".join(self.para_xml(x) for x in paras) or self.para_xml(self.spacer(6, inner_w))
            est = m["top"] + m["bottom"] + 1000
            if cell["rspan"] == 1:
                row_h[ri] = max(row_h[ri], min_h.get(ri, 0), est)
            cell["x"] = (bf, m, va, w, xml, int(ri in header_rows))
        for ri in range(nrows):
            row_h[ri] = row_h[ri] or max(min_h.get(ri, 0), 1000)
        trs = []
        for ri in range(nrows):
            tcs = []
            for cell in [c for c in cells if c["row"] == ri]:
                bf, m, va, w, xml, hdr = cell["x"]
                h = sum(row_h[ri:ri + cell["rspan"]])
                tcs.append(f'<hp:tc name="" header="{hdr}" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="{bf}">'
                           f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="{va}" linkListIDRef="0" '
                           f'linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">{xml}</hp:subList>'
                           f'<hp:cellAddr colAddr="{cell["col"]}" rowAddr="{ri}"/><hp:cellSpan colSpan="{cell["span"]}" rowSpan="{cell["rspan"]}"/>'
                           f'<hp:cellSz width="{w}" height="{h}"/>'
                           f'<hp:cellMargin left="{m["left"]}" right="{m["right"]}" top="{m["top"]}" bottom="{m["bottom"]}"/></hp:tc>')
            trs.append("<hp:tr>" + "".join(tcs) + "</hp:tr>")
        self.oid += 1
        return (f'<hp:tbl id="{self.oid}" zOrder="{self.oid - 1000}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" '
                f'lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="{int(bool(header_rows))}" rowCnt="{nrows}" '
                f'colCnt="{ncols}" cellSpacing="0" borderFillIDRef="{self.st.none_bf()}" noAdjust="0">'
                f'<hp:sz width="{sum(cols)}" widthRelTo="ABSOLUTE" height="{sum(row_h)}" heightRelTo="ABSOLUTE" protect="0"/>'
                '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
                'horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
                '<hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:inMargin left="510" right="510" top="141" bottom="141"/>'
                + "".join(trs) + "</hp:tbl>")

    def table_para(self, t, width):
        ind = (t.get("indent") or {}).get("size") or 0
        ps = PS(align="LEFT", left=int(ind) * TW, ls_type="PERCENT", ls_val=100)
        cs = CS(pt=1.0)
        return Para(ps, [[cs, ""], ("obj", self.table_xml(t, width - int(ind) * TW))], width)

    # ── 그림 ──
    def pic_xml(self, r, width):
        data = base64.b64decode(r["data"]["$b64"]) if isinstance(r.get("data"), dict) else None
        if not data:
            return None
        self.stats["images"] += 1
        ext = "png" if r.get("kind") == "png" else "jpg"
        bid = f"image{len(self.bins) + 1}"
        self.bins.append((bid, ext, data))
        tr = r.get("transformation") or {}
        w, h = int(tr.get("width", 400)) * 75, int(tr.get("height", 300)) * 75     # px(96dpi) → HWPUNIT
        if w > width:
            h, w = round(h * width / w), width
        cap = round((PAGE_H - sum(PAGE[k] for k in ("top", "bottom", "header", "footer"))) * 0.85)
        if h > cap:
            w, h = round(w * cap / h), cap
        self.oid += 1
        return (f'<hp:pic id="{self.oid}" zOrder="{self.oid - 1000}" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" '
                f'lock="0" dropcapstyle="None" href="" groupLevel="0" instid="{self.oid}" reverse="0">'
                f'<hp:offset x="0" y="0"/><hp:orgSz width="{w}" height="{h}"/><hp:curSz width="{w}" height="{h}"/>'
                f'<hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="{w // 2}" centerY="{h // 2}" rotateimage="1"/>'
                '<hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
                '<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>'
                f'<hc:img binaryItemIDRef="{bid}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
                f'<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="{w}" y="0"/><hc:pt2 x="{w}" y="{h}"/><hc:pt3 x="0" y="{h}"/></hp:imgRect>'
                f'<hp:imgClip left="0" right="{w}" top="0" bottom="{h}"/><hp:inMargin left="0" right="0" top="0" bottom="0"/>'
                f'<hp:imgDim dimwidth="{w}" dimheight="{h}"/><hp:effects/>'
                f'<hp:sz width="{w}" widthRelTo="ABSOLUTE" height="{h}" heightRelTo="ABSOLUTE" protect="0"/>'
                '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
                'horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
                '<hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:shapeComment>그림입니다.</hp:shapeComment></hp:pic>')

    # ── 블록 ──
    def blocks(self, children, width, clean=True):
        out = []
        children = self.gap_after_header(children, width)
        for c in children:
            if c.get("_gap"):
                out.append(self.spacer(c["_gap"], width, True))
                continue
            if c["type"] == "p" and c.get("_band"):
                out += self.band_para(c, width)
            elif c["type"] == "p":
                out += self.convert_p(c, width)
            elif c["type"] == "tbl":
                out.append(self.table_para(c, width))
        # 쪽 나누기 바로 뒤의 간격용 빈 줄은 없앰: 원본은 새 쪽이 제목·표로 바로 시작
        if clean:
            cleaned, carry = [], False
            for p in out:
                is_space = not p.text().strip() and not any(cs == "obj" for cs, _ in p.runs)
                if (p.page_break or carry) and is_space:
                    carry = True
                    continue
                if carry:
                    p.page_break, carry = True, False
                cleaned.append(p)
            out = cleaned
        for p in out:
            fit(p, self.stats)
            if getattr(p, "prefix", None):
                # 내어쓰기는 첫 줄 양쪽 정렬 늘어남에 따라 달라지고, 내어쓰기가 바뀌면 줄 맞춤도 달라짐 → 둘 다 바뀌지 않을 때까지
                for _ in range(4):
                    before = (p.delta, p.ps.intent, p.ps.charwrap)
                    apply_quick_indent(p)
                    fit(p, self.stats)
                    if (p.delta, p.ps.intent, p.ps.charwrap) == before:
                        break
                apply_quick_indent(p)
                # 두 상태를 오가다 끝나면 자간과 내어쓰기가 어긋날 수 있음 → 최종 내어쓰기에서 여유 확인, 부족하면 자간 한 단계씩
                # (자간을 줄이면 첫 줄 늘어남이 줄어 내어쓰기도 작아지므로 수렴)
                _finalize(p)
        self.stats["paras"] += len(out)
        return out

    @staticmethod
    def gap_after_header(children, width):
        """번호 머리글(Ⅴ 향후 일정) 바로 뒤에 표·박스가 오면 한 줄 띄움.
        원본 51건: 머리글 뒤 42곳 중 41곳이 빈 줄(표 앞이면 표 글자 크기의 빈 줄, 160%)."""
        out = list(children)
        for i, c in enumerate(children):
            if c.get("type") != "tbl" or not c.get("_hdr"):
                continue
            j = i + 1
            while j < len(children) and children[j].get("_spacer"):
                j += 1
            if j >= len(children) or children[j].get("type") != "tbl":
                continue
            sizes = []
            def scan(o):
                if isinstance(o, dict):
                    if o.get("type") == "r" and o.get("size") and (o.get("text") or "").strip():
                        sizes.append(o["size"] / 2)
                    for v in o.values():
                        scan(v)
                elif isinstance(o, list):
                    for v in o:
                        scan(v)
            rows = children[j].get("rows") or []
            scan(rows[1:] or rows)
            fs = min(15, max(10, sizes[0] if sizes else 13))
            for k in range(i + 1, j):
                out[k] = None
            out[i] = [c, {"type": "p", "_gap": fs * 1.6}]
        flat = []
        for c in out:
            if c is None:
                continue
            flat += c if isinstance(c, list) else [c]
        return flat

    def para_xml(self, p, extra=""):
        self.pid += 1
        pid = self.st.para(p.ps)
        runs = []
        first = True
        for cs, t in p.runs:
            if cs == "obj":
                runs.append(f'<hp:run charPrIDRef="{self.st.char(CS(pt=1.0))}">{extra if first else ""}{t}<hp:t/></hp:run>')
            else:
                c2 = cs.copy(spacing=max(-50, min(50, cs.spacing + p.delta)))
                runs.append(f'<hp:run charPrIDRef="{self.st.char(c2)}">{extra if first else ""}{self.t_xml(t)}</hp:run>')
            first = False
        return (f'<hp:p id="{self.pid}" paraPrIDRef="{pid}" styleIDRef="0" pageBreak="{int(p.page_break)}" columnBreak="0" merged="0">'
                + "".join(runs) + "</hp:p>")

    @staticmethod
    def t_xml(t):
        if not t:
            return "<hp:t/>"
        out = []
        for s in re.split(r"(\t|\n)", t):
            if s == "\t":
                out.append('<hp:tab width="0" leader="0" type="1"/>')
            elif s == "\n":
                out.append("<hp:lineBreak/>")
            elif s:
                out.append(esc(s))
        return "<hp:t>" + "".join(out) + "</hp:t>"

    def sec_pr(self):
        g = PAGE
        return (f'<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" '
                'outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">'
                '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
                '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" '
                'hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
                f'<hp:pagePr landscape="WIDELY" width="{PAGE_W}" height="{PAGE_H}" gutterType="LEFT_ONLY">'
                f'<hp:margin header="{g["header"]}" footer="{g["footer"]}" gutter="0" left="{g["left"]}" right="{g["right"]}" '
                f'top="{g["top"]}" bottom="{g["bottom"]}"/></hp:pagePr>'
                '<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
                '<hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
                '<hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr>'
                '<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
                '<hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
                '<hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>'
                + "".join(f'<hp:pageBorderFill type="{t}" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">'
                          '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>' for t in ("BOTH", "EVEN", "ODD"))
                + '</hp:secPr><hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>')

    def header_ctrl(self):
        hdr = ((self.ir["sections"][0].get("headers") or {}).get("default"))
        if not hdr:
            return ""
        paras = self.blocks(hdr.get("children") or [], BODY_W)
        xml = "".join(self.para_xml(p) for p in paras)
        return (f'<hp:ctrl><hp:header id="1" applyPageType="BOTH"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
                f'vertAlign="BOTTOM" linkListIDRef="0" linkListNextIDRef="0" textWidth="{BODY_W}" textHeight="{PAGE["header"]}" '
                f'hasTextRef="0" hasNumRef="0">{xml}</hp:subList></hp:header></hp:ctrl>')

    def build(self):
        sec = self.ir["sections"][0]
        paras = []
        for c in self.gap_after_header(sec.get("children") or [], BODY_W):
            if c.get("pageBreakBefore") and paras:
                self.block += 1
            paras += self.blocks([c], BODY_W, clean=False)
        # 쪽 나누기 뒤 빈 줄 정리(구간 경계를 넘어서도)
        cleaned, carry = [], False
        for p in paras:
            is_space = not p.text().strip() and not any(cs == "obj" for cs, _ in p.runs)
            if (p.page_break or carry) and is_space and cleaned:
                carry = True
                continue
            if carry:
                p.page_break, carry = True, False
            cleaned.append(p)
        paras = cleaned
        head = self.header_ctrl()
        body = []
        for i, p in enumerate(paras):
            body.append(self.para_xml(p, (self.sec_pr() + head) if i == 0 else ""))
        section = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {NS}>' + "".join(body) + "</hs:sec>").encode()
        text = "\r\n".join(p.text() for p in paras if p.text().strip())
        return section, self.st.build(), text


def skeleton():
    import hwpx
    z = zipfile.ZipFile(Path(hwpx.__file__).parent / "data" / "Skeleton.hwpx")
    return z.read


def _render(ir, skel, levels):
    w = Writer(ir, skel, levels)
    section, header, text = w.build()
    return w, section, header, text


def write(ir_path, out_path, title=None):
    ir = json.load(open(ir_path, encoding="utf-8"))
    skel = skeleton()
    levels = {}
    w, section, header, text = _render(ir, skel, levels)
    # 자동 쪽 맞춤: 구간의 마지막 쪽이 25% 미만으로 조금 넘치면 그 구간만 단계적으로 압축
    import hwpx_layout as HL
    for _ in range(len(FIT_LEVELS) * 3):
        tmp = io.BytesIO()
        _pack(tmp, skel, w, section, header, text, title)
        fills = HL.Doc(tmp).page_fill()
        changed = False
        for b, pages in enumerate(fills):
            if len(pages) > 1 and pages[-1] < 0.25 and levels.get(b, 0) < len(FIT_LEVELS) - 1:
                levels[b] = levels.get(b, 0) + 1
                changed = True
        if not changed:
            break
        w, section, header, text = _render(ir, skel, levels)
    # 압축해도 쪽이 줄지 않은 구간은 원래대로
    tmp = io.BytesIO(); _pack(tmp, skel, w, section, header, text, title)
    final = HL.Doc(tmp).page_fill()
    base_w, bs, bh, bt = _render(ir, skel, {})
    tmp2 = io.BytesIO(); _pack(tmp2, skel, base_w, bs, bh, bt, title)
    base = HL.Doc(tmp2).page_fill()
    keep = {b: lv for b, lv in levels.items() if b < len(final) and b < len(base) and len(final[b]) < len(base[b])}
    if keep != levels:
        levels = keep
        w, section, header, text = _render(ir, skel, levels)
    w.stats["fit_levels"] = levels
    _pack(out_path, skel, w, section, header, text, title)
    return w.stats


def _pack(out_path, skel, w, section, header, text, title=None):
    hpf = skel("Contents/content.hpf").decode()
    if title is None:
        title = next((t for t in text.split("\r\n") if t.strip()), "")
    hpf = hpf.replace("<opf:title/>", f"<opf:title>{esc(title)}</opf:title>")
    hpf = re.sub(r'(<opf:meta name="(?:creator|lastsaveby|CreatedDate|ModifiedDate|date)" content="text">)[^<]*', r"\1", hpf)
    items = "".join(f'<opf:item id="{bid}" href="BinData/{bid}.{ext}" media-type="image/{"png" if ext == "png" else "jpg"}" isEmbeded="1"/>'
                    for bid, ext, _ in w.bins)
    hpf = hpf.replace('<opf:item id="settings"', items + '<opf:item id="settings"')
    files = [("mimetype", skel("mimetype")), ("version.xml", skel("version.xml")), ("Contents/header.xml", header),
             ("Contents/section0.xml", section)]
    files += [(f"BinData/{bid}.{ext}", data) for bid, ext, data in w.bins]
    files += [("Preview/PrvText.txt", text[:1024].encode()), ("settings.xml", skel("settings.xml")),
              ("META-INF/container.rdf", skel("META-INF/container.rdf")), ("Contents/content.hpf", hpf.encode()),
              ("META-INF/container.xml", skel("META-INF/container.xml")), ("META-INF/manifest.xml", skel("META-INF/manifest.xml")),
              ("Preview/PrvImage.png", skel("Preview/PrvImage.png"))]
    with zipfile.ZipFile(out_path, "w") as z:
        for name, data in files:
            comp = zipfile.ZIP_STORED if name in ("mimetype",) or name.startswith("BinData/") else zipfile.ZIP_DEFLATED
            z.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0)), data, compress_type=comp)


if __name__ == "__main__":
    ir_path, out = sys.argv[1], sys.argv[2]
    stats = write(ir_path, out)
    if "--rm-ir" in sys.argv:
        os.remove(ir_path)
    lv = stats.pop("fit_levels")
    fitmsg = (" | 쪽 맞춤: " + ", ".join(f"구간{b + 1} 간격×{FIT_LEVELS[l]['sp']}" + (f"·줄간격 {FIT_LEVELS[l]['ls']}%" if FIT_LEVELS[l]['ls'] else "") for b, l in sorted(lv.items()))) if lv else ""
    print("written:", out, "| 줄 맞춤: 자간 %(tight)d · 끌어올림 %(pullup)d · 글자 단위 %(charwrap)d | 표 %(tables)d · 그림 %(images)d" % stats + fitmsg)
