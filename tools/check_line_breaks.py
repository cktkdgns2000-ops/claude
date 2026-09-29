"""hwpx 문단별 줄 수를 Hancom 줄나눔 규칙 + 실제 글꼴 폭으로 추정해 원본 docx 렌더(PDF) 줄 수와 비교.
사용: python3 check_line_breaks.py out.hwpx word_like.pdf [--pages] [-v]"""
import re, sys, zipfile, subprocess
from lxml import etree
import pymupdf
from fontTools.ttLib import TTFont

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"
HUC = "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"

FILES = {("나눔명조", False): "NanumMyeongjo", ("나눔명조", True): "NanumMyeongjoBold",
         ("나눔고딕", False): "NanumGothic", ("나눔고딕", True): "NanumGothicBold"}
_cache = {}


def font(face, bold):
    key = (face, bold)
    if key not in _cache:
        if key in FILES:
            path = f"/usr/share/fonts/truetype/nanum/{FILES[key]}.ttf"
        else:
            path = subprocess.run(["fc-match", "-f", "%{file}", face], capture_output=True, text=True).stdout
        f = TTFont(path)
        _cache[key] = (f.getBestCmap(), f["hmtx"].metrics, f["head"].unitsPerEm)
    return _cache[key]


def adv_em(face, bold, ch):
    cmap, hm, upm = font(face, bold)
    g = cmap.get(ord(ch))
    if g is None:
        return 1.0
    return hm[g][0] / upm


NO_START = set(")]}’”,.:;!?ㆍ·、。」』〉》")
NO_END = set("([{‘“「『〈《")


def load(hwpx):
    z = zipfile.ZipFile(hwpx)
    head = etree.fromstring(z.read("Contents/header.xml"))
    sec = etree.fromstring(z.read("Contents/section0.xml"))
    fonts = {int(f.get("id")): f.get("face") for f in head.find(f".//{HH}fontface")}
    chars = {}
    for c in head.iter(f"{HH}charPr"):
        chars[c.get("id")] = dict(face=fonts[int(c.find(f"{HH}fontRef").get("hangul"))],
                                  pt=int(c.get("height")) / 100, bold=c.find(f"{HH}bold") is not None,
                                  ratio=int(c.find(f"{HH}ratio").get("hangul")),
                                  sp=int(c.find(f"{HH}spacing").get("hangul")))
    paras = {}
    for p in head.iter(f"{HH}paraPr"):
        case = next((c for c in p.iter(f"{HP}case") if c.find(f"{HH}margin") is not None), None)
        if case is None: continue
        m = {e.tag.split("}")[1]: int(e.get("value")) for e in case.find(f"{HH}margin")}
        paras[p.get("id")] = m
    return sec, chars, paras


def para_lines(p, width, chars, paras):
    m = paras[p.get("paraPrIDRef")]
    left, intent = m["left"], m["intent"]
    first_w = width - left - max(intent, 0)
    rest_w = width - left - max(-intent, 0)
    items = []   # (ch, advance_hu, spacing_factor_part)
    for run in p.findall(f"{HP}run"):
        c = chars[run.get("charPrIDRef")]
        for t in run.findall(f"{HP}t"):
            seq = []
            if t.text:
                seq += list(t.text)
            for ch in t:
                if ch.tag == f"{HP}tab":
                    seq.append("\t")
                if ch.tail:
                    seq += list(ch.tail)
            for ch in seq:
                if ch == "\t":
                    items.append(("\t", 0, 0))
                    continue
                base = adv_em(c["face"], c["bold"], ch) * c["pt"] * 100 * c["ratio"] / 100
                items.append((ch, base * (1 + c["sp"] / 100), base))
    if not items:
        return 0
    # 탭: 첫 줄에서 내어쓰기 위치로 이동(자동 탭)
    if any(ch == "\t" for ch, _, _ in items):
        k = [i for i, x in enumerate(items) if x[0] == "\t"][0]
        pre = sum(a for _, a, _ in items[:k])
        items[k] = ("\t", max(0, -intent - pre), 0)
    text = [ch for ch, _, _ in items]
    lines, start, n = 1, 0, len(items)
    widths = [first_w, rest_w]
    while True:
        wlim = widths[min(lines - 1, 1)]
        used, end = 0.0, start
        while end < n:
            ch, a, raw = items[end]
            if ch == " ":
                used += a; end += 1; continue
            if used + raw > wlim and end > start:
                break
            used += a; end += 1
        if end >= n:
            return lines
        opts = [i for i in range(start + 1, end + 1) if text[i - 1] == " " and text[i] != " "
                and text[i] not in NO_START and text[i - 1] not in NO_END]
        # 라틴 단어 안에서는 끊지 않음, 한글 어절도 통째(BREAK_WORD)
        start = max(opts) if opts else end
        while start < n and text[start] == " ":
            start += 1
        if start >= n:
            return lines
        lines += 1


def hwpx_counts(hwpx):
    sec, chars, paras = load(hwpx)
    out = []
    BODY = 59528 - 2 * 7085

    def walk(container, width):
        for p in container.findall(f"{HP}p"):
            tbls = p.findall(f".//{HP}run/{HP}tbl")
            if tbls:
                for tbl in tbls:
                    for tc in tbl.iter(f"{HP}tc"):
                        cm = tc.find(f"{HP}cellMargin")
                        w = int(tc.find(f"{HP}cellSz").get("width")) - int(cm.get("left")) - int(cm.get("right"))
                        walk(tc.find(f"{HP}subList"), w)
                continue
            txt = "".join("".join(t.itertext()) for t in p.iter(f"{HP}t"))
            if txt.strip():
                out.append((re.sub(r"\s", "", txt), para_lines(p, width, chars, paras)))
    walk(sec, BODY)
    return out


def pdf_counts(pdf, targets):
    lines = []
    for page in pymupdf.open(pdf):
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = re.sub(r"[\s​︎]", "", "".join(s["text"] for s in l["spans"]))
                if t:
                    lines.append((page.number + 1, round(l["bbox"][1]), t))
    res, used, i = [], set(), 0
    def try_at(j, norm):
        acc, cnt = "", 0
        while j < len(lines) and len(acc) < len(norm) and j not in used:
            acc += lines[j][2]; cnt += 1; j += 1
        return cnt if acc == norm else None
    for norm, _ in targets:
        found = None
        for j in list(range(i, len(lines))) + list(range(0, i)):
            if j in used: continue
            c = try_at(j, norm)
            if c:
                found = (j, c); break
        if found:
            j, c = found
            for k in range(j, j + c): used.add(k)
            res.append((c, lines[j][0])); i = j + c
        else:
            res.append((None, None))
    return res


if __name__ == "__main__":
    h = hwpx_counts(sys.argv[1])
    p = pdf_counts(sys.argv[2], h)
    diff = 0
    for (t, n), (m, pg) in zip(h, p):
        flag = "" if n == m else "  <-- 차이"
        if n != m:
            diff += 1
        if n != m or "-v" in sys.argv:
            print(f"p{pg} hwpx={n} word={m}{flag} | {t[:40]}")
    print("paragraphs", len(h), "differences", diff, "unmatched", sum(1 for m, _ in p if m is None))


def page_heights(hwpx):
    sec, chars, paras = load(hwpx)
    head = etree.fromstring(zipfile.ZipFile(hwpx).read("Contents/header.xml"))
    ls = {}
    for p in head.iter(f"{HH}paraPr"):
        case = next((c for c in p.iter(f"{HP}case") if c.find(f"{HH}margin") is not None), None)
        if case is None: continue
        l = case.find(f"{HH}lineSpacing"); ls[p.get("id")] = (l.get("type"), int(l.get("value")))
    BODY = 59528 - 2 * 7085

    def maxpt(p):
        hs = [chars[r.get("charPrIDRef")]["pt"] for r in p.findall(f"{HP}run")
              if "".join("".join(t.itertext()) for t in r.findall(f"{HP}t")).strip()]
        if not hs:
            hs = [chars[r.get("charPrIDRef")]["pt"] for r in p.findall(f"{HP}run")]
        return max(hs) * 100

    def ph(p, width):
        tbls = p.findall(f".//{HP}run/{HP}tbl")
        if tbls:
            return sum(th(t) for t in tbls)
        m = paras[p.get("paraPrIDRef")]
        t, v = ls[p.get("paraPrIDRef")]
        n = max(1, para_lines(p, width, chars, paras))
        pitch = maxpt(p) * v / 100 if t == "PERCENT" else v
        return m["prev"] + m["next"] + n * pitch

    def th(tbl):
        rows = tbl.findall(f"{HP}tr")
        hts = [0] * len(rows)
        spans = []
        for ri, tr in enumerate(rows):
            for tc in tr.findall(f"{HP}tc"):
                cm = tc.find(f"{HP}cellMargin")
                w = int(tc.find(f"{HP}cellSz").get("width")) - int(cm.get("left")) - int(cm.get("right"))
                h = int(cm.get("top")) + int(cm.get("bottom")) + sum(ph(p, w) for p in tc.find(f"{HP}subList").findall(f"{HP}p"))
                h = max(h, int(tc.find(f"{HP}cellSz").get("height")))
                rs = int(tc.find(f"{HP}cellSpan").get("rowSpan"))
                if rs == 1: hts[ri] = max(hts[ri], h)
                else: spans.append((ri, rs, h))
        for ri, rs, h in spans:
            cur = sum(hts[ri:ri + rs])
            if h > cur: hts[ri + rs - 1] += h - cur
        return sum(hts)

    pages, cur = [], 0
    for p in sec.findall(f"{HP}p"):
        if p.get("pageBreak") == "1":
            pages.append(cur); cur = 0
        cur += ph(p, BODY)
    pages.append(cur)
    body_h = 84186 - 5670 - 4250
    for i, h in enumerate(pages, 1):
        print(f"쪽{i}: 내용 {h/100:.1f}pt / 본문 {body_h/100:.1f}pt  여유 {(body_h-h)/100:.1f}pt")


if "--pages" in sys.argv:
    page_heights(sys.argv[1])
