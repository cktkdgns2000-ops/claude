"""원본 hwpx(한글 저장본 변환)에서 문단 종류별 실제 서식값을 집계한다."""
import json, re, zipfile, sys
from collections import Counter, defaultdict
from lxml import etree

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"


def load_header(z):
    h = etree.fromstring(z.read("Contents/header.xml"))
    fonts = {}
    for ff in h.iter(f"{HH}fontface"):
        if ff.get("lang") == "HANGUL":
            fonts = {f.get("id"): f.get("face") for f in ff}
    chars = {}
    for c in h.iter(f"{HH}charPr"):
        g = lambda t, a="hangul": c.find(f"{HH}{t}").get(a)
        chars[c.get("id")] = dict(
            font=fonts.get(g("fontRef")), pt=int(c.get("height")) / 100, ratio=int(g("ratio")),
            spacing=int(g("spacing")), bold=c.find(f"{HH}bold") is not None,
            color=c.get("textColor"), shade=c.get("shadeColor"), fontspace=c.get("useFontSpace"),
            underline=c.find(f"{HH}underline").get("type"),
            sup=c.find(f"{HH}supscript") is not None)
    paras = {}
    for p in h.iter(f"{HH}paraPr"):
        case = next((x for x in p.iter(f"{HP}case") if x.find(f"{HH}margin") is not None), None)
        box = case if case is not None else p
        m = box.find(f"{HH}margin")
        if m is None:
            continue
        mv = {e.tag.split("}")[1]: int(e.get("value")) for e in m}
        ls = box.find(f"{HH}lineSpacing")
        bs = p.find(f"{HH}breakSetting")
        paras[p.get("id")] = dict(
            align=p.find(f"{HH}align").get("horizontal"), left=mv["left"], intent=mv["intent"],
            right=mv["right"], prev=mv["prev"], next=mv["next"],
            ls=(ls.get("type"), int(ls.get("value"))), condense=p.get("condense"),
            nonlatin=bs.get("breakNonLatinWord"), latin=bs.get("breakLatinWord"),
            widow=bs.get("widowOrphan"), keepnext=bs.get("keepWithNext"),
            tab=p.get("tabPrIDRef"), border=p.find(f"{HH}border").get("borderFillIDRef"))
    tabs = {}
    for t in h.iter(f"{HH}tabPr"):
        items = [(i.get("pos"), i.get("type")) for i in t.iter(f"{HH}tabItem") if i.get("unit") or True]
        tabs[t.get("id")] = dict(auto=t.get("autoTabLeft"), items=items[:2])
    bfs = {}
    for b in h.iter(f"{HH}borderFill"):
        d = {}
        for s in ("left", "right", "top", "bottom"):
            e = b.find(f"{HH}{s}Border")
            d[s] = (e.get("type"), e.get("width"), e.get("color")) if e is not None else None
        wb = b.find(f".//{HC}winBrush")
        d["fill"] = wb.get("faceColor") if wb is not None else None
        bfs[b.get("id")] = d
    return chars, paras, tabs, bfs


MARKERS = [("□", r"^□"), ("ㅇ", r"^[ㅇ○]"), ("-", r"^-\s"), ("·", r"^[·∙•]"), ("**", r"^\*\*"),
           ("*", r"^\*"), ("※", r"^※"), ("➊", r"^[➊-➓❶-❿]"), ("①", r"^[①-⑳]"), ("◈", r"^◈"),
           ("➡", r"^[➡⇨⇒▶►]"), ("■", r"^■"), ("Ⅰ", r"^[ⅠⅡⅢⅣⅤⅥI]\s*\S"), ("1.", r"^\d\.\s*\S"),
           ("<표제>", r"^<.*>$"), ("( )", r"^\(.*\)$")]


def classify(text):
    t = text.strip()
    for name, rx in MARKERS:
        if re.match(rx, t):
            return name
    return "기타"


def para_info(p, chars, paras):
    runs = []
    for r in p.findall(f"{HP}run"):
        txt = "".join("".join(t.itertext()) for t in r.findall(f"{HP}t"))
        runs.append((r.get("charPrIDRef"), txt))
    text = "".join(t for _, t in runs)
    # 주 글자 모양: 기호 뒤 본문에서 글자 수가 가장 많은 run
    cnt = Counter()
    for cid, t in runs:
        cnt[cid] += len(re.sub(r"\s", "", t))
    main = cnt.most_common(1)[0][0] if cnt and cnt.most_common(1)[0][1] else (runs[0][0] if runs else None)
    marker_cid = next((cid for cid, t in runs if t.strip()), None)
    segs = p.findall(f"{HP}linesegarray/{HP}lineseg")
    pitch = [int(s.get("vertsize")) + int(s.get("spacing")) for s in segs]
    tabs = len(p.findall(f".//{HP}tab"))
    return dict(text=text, cls=classify(text), para=paras.get(p.get("paraPrIDRef")), char=chars.get(main),
                mchar=chars.get(marker_cid), lines=len(segs), pitch=pitch, tabs=tabs,
                pagebreak=p.get("pageBreak"), segs=[(int(s.get("textpos")), int(s.get("vertpos"))) for s in segs])


def analyze(mapping):
    rows = []
    tables = []
    pages = []
    for key, m in mapping.items():
        if not m:
            continue
        z = zipfile.ZipFile(f"conv/{m['conv']}.hwpx")
        chars, paras, tabs, bfs = load_header(z)
        for sn in sorted(n for n in z.namelist() if re.match(r"Contents/section\d+\.xml", n)):
            sec = etree.fromstring(z.read(sn))
            pp = sec.find(f".//{HP}pagePr")
            if pp is not None:
                mg = pp.find(f"{HP}margin")
                pages.append((key, {a: int(mg.get(a)) for a in ("top", "bottom", "left", "right", "header", "footer")}))
            for p in sec.findall(f"{HP}p"):
                tbl = p.find(f".//{HP}tbl")
                if tbl is not None:
                    pos = tbl.find(f"{HP}pos")
                    cells = []
                    for tc in tbl.iter(f"{HP}tc"):
                        cm = tc.find(f"{HP}cellMargin")
                        cps = [para_info(x, chars, paras) for x in tc.find(f"{HP}subList").findall(f"{HP}p")]
                        cells.append(dict(bf=bfs.get(tc.get("borderFillIDRef")), hasMargin=tc.get("hasMargin"),
                                          margin={a: int(cm.get(a)) for a in ("left", "right", "top", "bottom")},
                                          header=tc.get("header"), paras=cps,
                                          row=int(tc.find(f"{HP}cellAddr").get("rowAddr"))))
                    im = tbl.find(f"{HP}inMargin"); om = tbl.find(f"{HP}outMargin")
                    tables.append(dict(doc=key, treat=pos.get("treatAsChar"), width=int(tbl.find(f"{HP}sz").get("width")),
                                       rows=int(tbl.get("rowCnt")), cols=int(tbl.get("colCnt")),
                                       inm={a: int(im.get(a)) for a in ("left", "right", "top", "bottom")},
                                       outm={a: int(om.get(a)) for a in ("left", "right", "top", "bottom")},
                                       repeat=tbl.get("repeatHeader"), pagebreak=tbl.get("pageBreak"),
                                       host=para_info(p, chars, paras), cells=cells, tblbf=bfs.get(tbl.get("borderFillIDRef"))))
                    continue
                info = para_info(p, chars, paras)
                info["doc"] = key
                rows.append(info)
    return rows, tables, pages


if __name__ == "__main__":
    mapping = json.load(open("conv/mapping.json"))
    rows, tables, pages = analyze(mapping)
    import pickle
    pickle.dump((rows, tables, pages), open("conv/analysis.pkl", "wb"))
    print("body paragraphs", len(rows), "tables", len(tables), "sections", len(pages))
