"""한글 일괄 저장본(hwp_batch_save.py 결과) ↔ 생성본 대조: 문단별 줄 위치, 쪽수·쪽 첫 줄.
사용: python3 compare_batch.py <한글저장 폴더> <생성본 폴더>   (파일 이름 같은 것끼리 대조)"""
import glob
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "skill_scripts"))
from hwpx_layout import Doc, HP, norm  # noqa: E402


def para_width(p, W):
    par = p.getparent()
    if par.tag != f"{HP}subList":
        return W
    tc = par.getparent()
    cs, cm = tc.find(f"{HP}cellSz"), tc.find(f"{HP}cellMargin")
    if cs is None or cm is None:
        return None
    return int(cs.get("width")) - int(cm.get("left")) - int(cm.get("right"))


def line_diffs(path):
    d = Doc(path, use_stored=True)
    tot, bad = 0, []
    for sec, W, H in d.sections:
        for p in sec.iter(f"{HP}p"):
            segs = p.findall(f"{HP}linesegarray/{HP}lineseg")
            txt = "".join(t.text or "" for t in p.iter(f"{HP}t"))
            if not segs or not txt.strip() or p.find(f".//{HP}tbl") is not None or p.find(f".//{HP}pic") is not None:
                continue
            w = para_width(p, W)
            if w is None:
                continue
            lines, _ = d.para_lines(p, w)
            st, ms = [int(s.get("textpos")) for s in segs], [l[2] for l in lines]
            tot += 1
            if st != ms:
                bad.append((txt[:30], [txt[a:a + 6] for a in st[1:]], [txt[a:a + 6] for a in ms[1:]]))
    return tot, bad


def main(saved_dir, gen_dir):
    T = dict(docs=0, paras=0, bad=0, pages=0, first=0, first_tot=0)
    for f in sorted(glob.glob(os.path.join(saved_dir, "*.hwpx"))):
        g = os.path.join(gen_dir, os.path.basename(f))
        if not os.path.exists(g):
            continue
        tot, bad = line_diffs(f)
        act, est = Doc(f, use_stored=True).actual_pages(), Doc(g).paginate()
        same = sum(1 for a, b in zip(act, est) if norm(a) == norm(b))
        T["docs"] += 1; T["paras"] += tot; T["bad"] += len(bad)
        T["pages"] += len(act) == len(est); T["first"] += same; T["first_tot"] += len(act)
        flag = "✓" if not bad and len(act) == len(est) and same == len(act) else "✗"
        print(f"{flag} {os.path.basename(f)[:45]:45s} 줄 {tot - len(bad)}/{tot} · 쪽 한글 {len(act)}/계산 {len(est)} · 첫 줄 {same}/{len(act)}")
        for b in bad:
            print("     ", b[0], "| 한글", b[1], "| 계산", b[2])
    print(f"\n문서 {T['docs']} · 문단 줄 위치 {T['paras'] - T['bad']}/{T['paras']} · 쪽수 일치 {T['pages']} · 쪽 첫 줄 {T['first']}/{T['first_tot']}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
