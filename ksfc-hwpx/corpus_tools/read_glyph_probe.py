"""한글 저장본 glyph_probe.hwpx → 글자별 실제 폭(em). make_glyph_probe.py의 json과 함께 사용.
사용: python3 read_glyph_probe.py <한글 저장본.hwpx> <glyph_probe.json> [결과.json]
한 줄 개수 c(마지막 줄 제외) → 폭 ∈ (W/(c+1), W/c]. 크기 두 개의 구간을 겹쳐 가운데 값을 em으로."""
import json
import os
import sys
import zipfile
from collections import defaultdict

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "skill_scripts"))
import hwp_metrics as HM  # noqa: E402

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"


def main(saved, meta, out=None):
    m = json.load(open(meta, encoding="utf-8"))
    W, rows = m["width"], m["rows"]
    sec = etree.fromstring(zipfile.ZipFile(saved).read("Contents/section0.xml"))
    paras = sec.findall(f"{HP}p")
    iv = defaultdict(lambda: [0.0, 9.9])
    bad = 0
    for r in rows:
        segs = [int(s.get("textpos")) for s in paras[r["para"]].findall(f"{HP}linesegarray/{HP}lineseg")]
        counts = [b - a for a, b in zip(segs, segs[1:])]
        if not counts:
            bad += 1
            continue
        lo_c, hi_c = min(counts), max(counts)
        # 줄마다 c개가 들어가고 c+1개는 넘침 → c*w ≤ W < (c+1)*w
        lo, hi = W / (hi_c + 1) / (r["pt"] * 100), W / lo_c / (r["pt"] * 100)
        k = (r["face"], r["bold"], r["ch"])
        iv[k][0], iv[k][1] = max(iv[k][0], lo), min(iv[k][1], hi)
    res, diff = {}, []
    for (face, bold, ch), (lo, hi) in sorted(iv.items()):
        em = (lo + hi) / 2
        res[f"{face}|{int(bold)}|{ch}"] = dict(em=round(em, 4), lo=round(lo, 4), hi=round(hi, 4))
        cur = HM.em(face, bold, ch)
        if not (lo - 0.005 <= cur <= hi + 0.005):
            diff.append((face, bold, ch, cur, lo, hi))
    print(f"측정 {len(res)}개 (줄 기록 없음 {bad}) · 현재 모델과 다른 글자 {len(diff)}개")
    for face, bold, ch, cur, lo, hi in diff:
        print(f"  {face}{' 굵게' if bold else ''} {ch!r}: 모델 {cur:.3f} → 한글 {lo:.3f}~{hi:.3f}")
    if out:
        json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=0)


if __name__ == "__main__":
    main(*sys.argv[1:4])
