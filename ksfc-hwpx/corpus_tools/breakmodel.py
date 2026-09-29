"""원본 본문 문단을 한글 규칙으로 다시 배치: (A) 글자 단위 모델이 원본 줄 바뀜을 맞히는지 검증,
(B) 같은 문단을 어절 단위로 배치했을 때 줄 끝 여백(=양쪽 정렬로 벌어지는 양) 분포."""
import json, re, zipfile, subprocess, sys
from collections import Counter
from lxml import etree
from fontTools.ttLib import TTFont

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
FILES = {("나눔명조", False): "/usr/share/fonts/truetype/nanum/NanumMyeongjo.ttf",
         ("나눔명조", True): "/usr/share/fonts/truetype/nanum/NanumMyeongjoBold.ttf",
         ("나눔고딕", False): "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
         ("나눔고딕", True): "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"}
for face in ("HY울릉도M", "HY헤드라인M"):
    path = subprocess.run(["fc-match", "-f", "%{file}", face], capture_output=True, text=True).stdout
    FILES[(face, False)] = FILES[(face, True)] = path
_f = {}


def adv(face, bold, ch):
    key = (face, bold)
    if key not in _f:
        t = TTFont(FILES[key]); _f[key] = (t.getBestCmap(), t["hmtx"].metrics, t["head"].unitsPerEm)
    cmap, hm, upm = _f[key]
    g = cmap.get(ord(ch))
    return hm[g][0] / upm if g else 1.0


NO_START = set(")]}’”,.:;!?ㆍ·、。」』〉》%")
NO_END = set("([{‘“「『〈《")


def layout(chars, first_w, rest_w, mode):
    """chars: [(ch, advance)] → 줄 시작 위치 목록과 각 줄의 사용 폭."""
    n = len(chars); starts = [0]; used_list = []; s = 0
    while s < n:
        w = first_w if len(starts) == 1 else rest_w
        used = 0.0; e = s; last_nonspace = s
        while e < n:
            ch, a = chars[e]
            if ch != " " and used + a > w + 0.5 and e > s:
                break
            used += a; e += 1
        if e >= n:
            used_list.append(sum(a for c, a in chars[s:n] if True)); break
        if mode == "word":
            opts = [i for i in range(s + 1, e + 1) if chars[i - 1][0] == " " and chars[i][0] != " "]
            b = max(opts) if opts else e
        else:
            b = e
            while b > s + 1 and (chars[b][0] in NO_START or chars[b - 1][0] in NO_END):
                b -= 1
        line = chars[s:b]
        while line and line[-1][0] == " ":
            line = line[:-1]
        used_list.append(sum(a for _, a in line))
        s = b
        while s < n and chars[s][0] == " ":
            s += 1
        if s < n:
            starts.append(s)
    return starts, used_list


def doc_paragraphs(path):
    z = zipfile.ZipFile(path)
    h = etree.fromstring(z.read("Contents/header.xml"))
    fonts = {}
    for ff in h.iter(f"{HH}fontface"):
        if ff.get("lang") == "HANGUL":
            fonts = {f.get("id"): f.get("face") for f in ff}
    chars = {c.get("id"): (fonts.get(c.find(f"{HH}fontRef").get("hangul")), int(c.get("height")) / 100,
                           int(c.find(f"{HH}ratio").get("hangul")), int(c.find(f"{HH}spacing").get("hangul")),
                           c.find(f"{HH}bold") is not None, c.get("useFontSpace"))
             for c in h.iter(f"{HH}charPr")}
    paras = {}
    for p in h.iter(f"{HH}paraPr"):
        case = next((x for x in p.iter(f"{HP}case") if x.find(f"{HH}margin") is not None), p)
        m = case.find(f"{HH}margin")
        if m is None: continue
        mv = {e.tag.split("}")[1]: int(e.get("value")) for e in m}
        paras[p.get("id")] = (mv["left"], mv["intent"], mv["right"], p.find(f"{HH}align").get("horizontal"))
    out = []
    for sn in sorted(n for n in z.namelist() if re.match(r"Contents/section\d+\.xml", n)):
        sec = etree.fromstring(z.read(sn))
        pp = sec.find(f".//{HP}pagePr"); mg = pp.find(f"{HP}margin")
        body = int(pp.get("width")) - int(mg.get("left")) - int(mg.get("right"))
        for p in sec.findall(f"{HP}p"):
            segs = p.findall(f"{HP}linesegarray/{HP}lineseg")
            runs = p.findall(f"{HP}run")
            if len(segs) < 1 or any(len(r) != len(r.findall(f"{HP}t")) for r in runs):
                continue   # 표·제어 문자가 섞인 문단 제외(textpos 정렬 보장)
            if any(len(t) for r in runs for t in r.findall(f"{HP}t")):
                continue   # hp:t 안에 탭 등 요소가 있는 문단 제외
            seq = []; ok = True
            for r in runs:
                face, pt, ratio, sp, bold, fs = chars[r.get("charPrIDRef")]
                if (face, bold) not in FILES or fs != "0":
                    ok = False; break
                for t in r.findall(f"{HP}t"):
                    for ch in (t.text or ""):
                        base = 0.5 if ch == " " else adv(face, bold, ch)
                        seq.append((ch, base * pt * 100 * ratio / 100 * (1 + sp / 100)))
            if not ok or not seq: continue
            left, intent, right, align = paras[p.get("paraPrIDRef")]
            if align != "JUSTIFY": continue
            width = body - left - right
            fw = width - max(intent, 0); rw = width - max(-intent, 0)
            out.append(dict(seq=seq, fw=fw, rw=rw, orig=[int(s.get("textpos")) for s in segs],
                            text="".join(c for c, _ in seq)))
    return out


if __name__ == "__main__":
    mapping = json.load(open("conv/mapping.json"))
    hit = Counter(); loose = []; mid_orig = 0; examples = []
    for key, m in mapping.items():
        for p in doc_paragraphs(f"conv/{m['conv']}.hwpx"):
            if len(p["orig"]) < 2:
                continue
            st, _ = layout(p["seq"], p["fw"], p["rw"], "char")
            hit["문단"] += 1
            hit["글자모델 일치" if st == p["orig"] else "불일치"] += 1
            if st != p["orig"]: continue            # 모델이 원본을 재현한 문단만 (B) 분석
            ws, used = layout(p["seq"], p["fw"], p["rw"], "word")
            t = p["text"]
            for i, s in enumerate(p["orig"][1:]):
                mid = t[s - 1] != " " and re.match("[가-힣]", t[s - 1]) and re.match("[가-힣]", t[s])
                mid_orig += bool(mid)
            for i, u in enumerate(used[:-1]):       # 마지막 줄은 양쪽 정렬 대상 아님
                w = p["fw"] if i == 0 else p["rw"]
                line = t[ws[i]: ws[i + 1]].rstrip()
                gaps = max(1, line.count(" "))
                pt = max(a for _, a in p["seq"]) / 100
                per_gap_em = (w - u) / gaps / (pt * 100)   # 띄어쓰기 하나가 더 벌어지는 양(em)
                loose.append((per_gap_em, (w - u) / w, key, line[-12:] + " / " + t[ws[i + 1]: ws[i + 1] + 10]))
    print(hit, "원본의 어절 중간 끊김:", mid_orig)
    import statistics
    per = sorted(x[0] for x in loose)
    print("어절 배치 시 줄 수", len(per))
    for th in (0.25, 0.5, 0.75, 1.0, 1.5):
        print(f"  띄어쓰기가 {th}em 넘게 벌어지는 줄: {sum(1 for x in per if x > th)} ({sum(1 for x in per if x > th)*100/len(per):.1f}%)")
    print("  중앙값 %.2fem, 90백분위 %.2fem" % (statistics.median(per), per[int(len(per) * .9)]))
    for x in sorted(loose, reverse=True)[:8]:
        print(f"   +{x[0]:.2f}em/칸 (줄 {x[1]*100:.0f}% 빔) {x[2]}: {x[3]}")
