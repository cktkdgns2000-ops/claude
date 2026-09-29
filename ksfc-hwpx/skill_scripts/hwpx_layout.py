#!/usr/bin/env python3
"""hwpx 쪽 배치 추정(한글 규칙 모델). 생성본의 쪽 경계 점검과, 원본(한글 저장본)과의 대조에 쓴다.

- 줄: hwp_metrics.layout(어절/글자 단위), 줄 높이 = 줄에서 가장 큰 글자 × 줄간격%(고정값이면 그 값)
- 표·그림(글자처럼 취급): 한 줄로 보고 통째로 다음 쪽으로 넘어감
- 쪽 나누기(pageBreak), 다음 문단과 함께(keepWithNext), 문단 보호(keepLines) 반영
- 원본처럼 줄 배치 기록(linesegarray)이 있는 파일은 한글이 실제로 나눈 쪽 경계도 읽을 수 있음(actual_pages)
사용: python3 hwpx_layout.py file.hwpx          → 쪽별 첫 줄 추정
      python3 hwpx_layout.py file.hwpx --actual → 한글 실제 쪽 경계(줄 배치 기록)와 비교
"""
import re
import sys
import zipfile
from lxml import etree

import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwp_metrics as HM  # noqa: E402

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"


class Doc:
    def __init__(self, path, use_stored=False):
        self.use_stored = use_stored   # True: 한글이 저장한 표 높이 사용(원본 검증용)
        z = zipfile.ZipFile(path)
        h = etree.fromstring(z.read("Contents/header.xml"))
        fonts = {}
        for ff in h.iter(f"{HH}fontface"):
            if ff.get("lang") == "HANGUL":
                fonts = {f.get("id"): f.get("face") for f in ff}
        self.chars = {}
        for c in h.iter(f"{HH}charPr"):
            g = lambda t: c.find(f"{HH}{t}").get("hangul")
            self.chars[c.get("id")] = dict(face=fonts.get(g("fontRef")), pt=int(c.get("height")) / 100,
                                           ratio=int(g("ratio")), sp=int(g("spacing")), bold=c.find(f"{HH}bold") is not None,
                                           sup=c.find(f"{HH}supscript") is not None or c.find(f"{HH}subscript") is not None,
                                           fontspace=c.get("useFontSpace") == "1")
        self.paras = {}
        for p in h.iter(f"{HH}paraPr"):
            case = next((x for x in p.iter(f"{HP}case") if x.find(f"{HH}margin") is not None), p)
            m = case.find(f"{HH}margin")
            if m is None:
                continue
            mv = {e.tag.split("}")[1]: int(e.get("value")) for e in m}
            ls = case.find(f"{HH}lineSpacing")
            bs = p.find(f"{HH}breakSetting")
            self.paras[p.get("id")] = dict(left=mv["left"], right=mv["right"], intent=mv["intent"], prev=mv["prev"], next=mv["next"],
                                           ls=(ls.get("type"), int(ls.get("value"))), mode="char" if bs.get("breakNonLatinWord") == "KEEP_WORD" else "word",
                                           keep_next=bs.get("keepWithNext") == "1", keep_lines=bs.get("keepLines") == "1")
        self.sections = []
        for n in sorted(x for x in z.namelist() if re.match(r"Contents/section\d+\.xml", x)):
            sec = etree.fromstring(z.read(n))
            pp = sec.find(f".//{HP}pagePr")
            mg = pp.find(f"{HP}margin")
            W = int(pp.get("width")) - int(mg.get("left")) - int(mg.get("right"))
            H = int(pp.get("height")) - int(mg.get("top")) - int(mg.get("bottom")) - int(mg.get("header")) - int(mg.get("footer"))
            self.sections.append((sec, W, H))

    # ── 문단 → 줄 목록 [(높이, 글자높이, 첫 글자 위치)] ──
    def para_lines(self, p, width):
        pr = self.paras[p.get("paraPrIDRef")]
        seq = []   # (ch, adv, height, is_obj, obj_h)
        for run in p.findall(f"{HP}run"):
            c = self.chars[run.get("charPrIDRef")]
            for el in run:
                tag = el.tag.split("}")[1]
                if tag == "t":
                    parts = [(el.text or "")]
                    for sub in el:
                        st = sub.tag.split("}")[1]
                        parts.append("\t" if st == "tab" else ("\n" if st == "lineBreak" else " "))
                        parts.append(sub.tail or "")
                    for ch in "".join(parts):
                        if ch == "\n":
                            seq.append(("\n", 0, c["pt"] * 100, False, 0))
                            continue
                        if ch == " " and c["fontspace"]:
                            a = HM.em(c["face"], c["bold"], " ") * 0 + 0.3 * c["pt"] * 100
                        else:
                            a = HM.advance(ch if ch != "\t" else " ", c["face"], c["bold"], c["pt"], c["ratio"], c["sp"])
                        seq.append((ch, a * (0.6 if c["sup"] else 1), c["pt"] * 100, False, 0))
                elif tag in ("tbl", "pic"):
                    w, h = self.obj_size(el, width)
                    seq.append(("￼", w, h, True, h))
        if not seq:
            c = self.chars[p.findall(f"{HP}run")[0].get("charPrIDRef")] if p.findall(f"{HP}run") else dict(pt=10)
            seq = []
            empty_h = c["pt"] * 100
        runs0 = p.findall(f"{HP}run")
        base_h = max([self.chars[r.get("charPrIDRef")]["pt"] * 100 for r in runs0] or [1000])
        W = width - pr["left"] - pr["right"]
        fw, rw = W - max(pr["intent"], 0), W - max(-pr["intent"], 0)
        lines = []
        # 강제 줄바꿈으로 나눈 조각별 배치
        chunks, cur = [], []
        for x in seq:
            if x[0] == "\n":
                chunks.append(cur); cur = []
            else:
                cur.append(x)
        chunks.append(cur)
        pos = 0
        for k, chunk in enumerate(chunks):
            if not chunk:
                h = seq[0][2] if seq else empty_h
                lines.append((self.pitch(pr, h), h, pos))
                pos += 1
                continue
            starts, used, gaps = HM.layout([(x[0], x[1]) for x in chunk], fw if k == 0 else rw, rw, pr["mode"])
            for i, s in enumerate(starts):
                e = starts[i + 1] if i + 1 < len(starts) else len(chunk)
                seg = chunk[s:e]
                objs = [x[4] for x in seg if x[3]]
                texth = max([x[2] for x in seg if not x[3]] or [0])
                if objs:
                    # 글자처럼 취급한 표·그림 줄: 개체 높이 + 문단 글자 크기 × (줄간격% − 100%) (원본 441건 확인)
                    h = max(objs)
                    fh = texth or base_h
                    extra = fh * (pr["ls"][1] - 100) / 100 if pr["ls"][0] == "PERCENT" else 0
                    lines.append((max(h, texth) + max(0, extra), max(h, texth), pos + s))
                else:
                    lines.append((self.pitch(pr, texth), texth, pos + s))
            pos += len(chunk) + 1
        return lines, pr

    @staticmethod
    def pitch(pr, h):
        t, v = pr["ls"]
        if t == "PERCENT":
            return h * v / 100
        if t == "FIXED":
            return v
        if t == "AT_LEAST":
            return max(v, h)
        return h + v   # BETWEEN_LINES(여백만 지정)

    def obj_size(self, el, width):
        sz = el.find(f"{HP}sz")
        w = int(sz.get("width"))
        if el.tag.endswith("pic"):
            om = el.find(f"{HP}outMargin")
            extra = (int(om.get("top")) + int(om.get("bottom"))) if om is not None else 0
            return w, int(sz.get("height")) + extra
        if self.use_stored:
            h = int(sz.get("height"))
            om = el.find(f"{HP}outMargin")
            ex = int(om.get("top")) + int(om.get("bottom")) if om is not None else 0
            if 0 < h < 84188 and 0 <= ex < 20000:
                return w, h + ex
        return w, self.table_height(el)

    def table_height(self, tbl):
        om = tbl.find(f"{HP}outMargin")
        extra = int(om.get("top")) + int(om.get("bottom")) if om is not None else 0
        extra = extra if 0 <= extra < 20000 else 0        # 비정상 저장값 무시
        rows = tbl.findall(f"{HP}tr")
        hts = [0] * len(rows)
        spans = []
        for ri, tr in enumerate(rows):
            for tc in tr.findall(f"{HP}tc"):
                cm = tc.find(f"{HP}cellMargin")
                if tc.get("hasMargin") == "1":
                    m = {a: int(cm.get(a)) for a in ("left", "right", "top", "bottom")}
                else:
                    im = tbl.find(f"{HP}inMargin")
                    m = {a: int(im.get(a)) for a in ("left", "right", "top", "bottom")}
                cw = int(tc.find(f"{HP}cellSz").get("width")) - m["left"] - m["right"]
                inner = 0
                for p in tc.find(f"{HP}subList").findall(f"{HP}p"):
                    ls, pr = self.para_lines(p, cw)
                    inner += pr["prev"] + pr["next"] + sum(l[0] for l in ls[:-1]) + (ls[-1][0] if ls else 0)
                stored = int(tc.find(f"{HP}cellSz").get("height"))
                stored = stored if 0 < stored < 84188 else 0    # 비정상 저장값(음수가 부호 없이 저장된 경우 등) 무시
                h = max(stored, inner + m["top"] + m["bottom"])
                rs = int(tc.find(f"{HP}cellSpan").get("rowSpan"))
                if rs == 1:
                    hts[ri] = max(hts[ri], h)
                else:
                    spans.append((ri, rs, h))
        for ri, rs, h in spans:
            cur = sum(hts[ri:ri + rs])
            if h > cur:
                hts[ri + rs - 1] += h - cur
        return sum(hts) + extra

    # ── 쪽 나누기 ──
    def paginate(self):
        pages = []   # 쪽별 첫 줄 텍스트
        for sec, W, H in self.sections:
            items = []
            for p in sec.findall(f"{HP}p"):
                lines, pr = self.para_lines(p, W)
                text = "".join("".join(t.itertext()) for t in p.iter(f"{HP}t"))
                items.append(dict(p=p, lines=lines, pr=pr, text=text, pb=p.get("pageBreak") == "1"))
            y, page_first = 0.0, None
            for i, it in enumerate(items):
                if it["pb"] and (y > 0 or pages):
                    pages.append(None); y = 0.0
                y += it["pr"]["prev"]
                ls = it["lines"]
                need_whole = it["pr"]["keep_lines"] or any(l[1] > 0 and it["text"] == "" for l in ls)
                # 다음 문단과 함께: 이 문단 전체 + 다음 문단 첫 줄이 안 들어가면 넘김
                if it["pr"]["keep_next"] and i + 1 < len(items) and items[i + 1]["lines"] and y > 0:
                    tot = sum(l[0] for l in ls) + items[i + 1]["lines"][0][1]
                    if y + tot > H and tot <= H:
                        pages.append(None); y = 0.0
                if need_whole and y > 0 and y + sum(l[0] for l in ls[:-1]) + ls[-1][1] > H:
                    pages.append(None); y = 0.0
                for (pitch, th, pos) in ls:
                    if y > 0 and y + th > H:
                        pages.append(None); y = 0.0
                    if not pages:
                        pages.append(None)
                    if pages[-1] is None and it["text"].strip():
                        pages[-1] = it["text"][pos:pos + 20]
                    y += pitch
                y += it["pr"]["next"]
        return [p or "(빈 쪽)" for p in pages]

    def actual_pages(self):
        """한글이 저장한 줄 배치 기록으로 실제 쪽 경계(쪽별 첫 줄)."""
        pages = []
        for sec, W, H in self.sections:
            last = None
            for p in sec.findall(f"{HP}p"):
                text = "".join("".join(t.itertext()) for t in p.iter(f"{HP}t"))
                for s in p.findall(f"{HP}linesegarray/{HP}lineseg"):
                    v = int(s.get("vertpos"))
                    if last is None or v < last:
                        pages.append(None)
                    if pages[-1] is None and text.strip():
                        tp = int(s.get("textpos"))
                        pages[-1] = text[tp:tp + 20]
                    last = v
        return [p or "(빈 쪽)" for p in pages]


def norm(t):
    return re.sub(r"\s", "", t or "")[:10]


if __name__ == "__main__":
    d = Doc(sys.argv[1])
    est = d.paginate()
    if "--actual" in sys.argv:
        act = d.actual_pages()
        print(f"추정 {len(est)}쪽 / 실제 {len(act)}쪽")
        for i in range(max(len(est), len(act))):
            a = act[i] if i < len(act) else "-"
            e = est[i] if i < len(est) else "-"
            print(f"  {i + 1:2d} {'✓' if norm(a) == norm(e) else '✗'} 실제: {a[:20]:22s} 추정: {e[:20]}")
    else:
        for i, e in enumerate(est, 1):
            print(f"{i:2d}쪽 첫 줄: {e}")
