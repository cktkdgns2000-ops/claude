#!/usr/bin/env python3
"""생성한 hwpx 점검: 서식 검증 · 쪽 배치(한글 줄 나눔 모델) · 쪽별 채움 · 짧은 마지막 줄 · 잔여 마크업.
사용: python3 check_hwpx.py <파일.hwpx> [원본 hwpx]
원본(한글이 저장한 hwpx)을 주면 쪽별 첫 줄을 원본과 대조한다(재현 작업용).
쪽 배치는 한글 저장본 51건 5,348문단으로 검증한 모델(hwpx_layout.py) 계산이다. 한글 화면이 최종 기준."""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from hwpx_layout import Doc, HP, norm  # noqa: E402
from regress_hwpx import leftover_markup, validate  # noqa: E402


def short_last_lines(doc, limit=0.12):
    """마지막 줄에 한두 글자만 남은 문단(줄 폭 대비 limit 이하) — 자간 한도(-10%) 때문에 못 올린 경우"""
    out = []
    for sec, W, H in doc.sections:
        for p in sec.findall(f"{HP}p"):
            if p.find(f".//{HP}tbl") is not None or p.find(f".//{HP}pic") is not None:
                continue
            lines, pr = doc.para_lines(p, W)
            text = "".join("".join(t.itertext()) for t in p.iter(f"{HP}t"))
            if len(lines) < 2 or "\n" in text:
                continue
            tail = text[lines[-1][2]:].strip()
            width = W - pr["left"] - pr["right"] - max(-pr["intent"], 0)
            if 0 < len(tail) <= 4 and len(tail) * 1500 / max(1, width) <= limit:
                out.append((text[:24], tail))
    return out


def main(path, orig=None):
    ok = validate(path)
    print("서식 검증: " + ("통과" if ok else ("건너뜀(python-hwpx 없음 — pip install python-hwpx)" if ok is None else "실패")))
    d = Doc(path)
    pages, fills = d.paginate(), d.page_fill()
    print(f"쪽수(계산): {len(pages)}")
    flat = [f for blk in fills for f in blk]
    for i, first in enumerate(pages):
        f = flat[i] if i < len(flat) else 0
        print(f"  {i + 1:>2}쪽 채움 {f * 100:5.1f}% | {first[:30]}")
    for b, blk in enumerate(fills):
        if len(blk) > 1 and blk[-1] < 0.25:
            print(f"  ! 구간 {b + 1}(쪽 나누기 사이)의 마지막 쪽이 {blk[-1] * 100:.0f}%만 참 — 조금 넘친 것. 축약·간격 조정 검토")
    tails = short_last_lines(d)
    if tails:
        print(f"짧은 마지막 줄 {len(tails)}곳(자간 -10% 한도로 못 올림 — 문장 축약 검토):")
        for head, tail in tails[:10]:
            print(f"  {head}… → '{tail}'")
    lo = leftover_markup(path)
    if lo:
        print("잔여 마크업:", lo)
    if orig:
        act = Doc(orig, use_stored=True).actual_pages()
        same = sum(1 for a, b in zip(act, pages) if norm(a) == norm(b))
        print(f"원본 대조: 원본 {len(act)}쪽, 쪽 첫 줄 일치 {same}/{len(act)}")
        for i, (a, b) in enumerate(zip(act, pages)):
            if norm(a) != norm(b):
                print(f"  {i + 1}쪽 원본 '{a[:20]}' / 생성 '{b[:20]}'")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(*sys.argv[1:3])
