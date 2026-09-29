"""글자 폭 측정 문서(glyph_probe.hwpx) 생성 — 한글에서 저장하면 줄 나눔 위치로 글자별 실제 폭을 잰다.

문단 하나 = 한 글자를 N번 반복(글자 단위 줄 나눔, 왼쪽 정렬, 자간 0, 장평 100). 한 줄에 들어간 개수 c로
폭 ∈ (본문 폭/(c+1), 본문 폭/c]. 두 크기(6·10pt)로 재어 구간을 좁힌다.
글꼴 파일이 없거나(HY·휴먼명조·맑은 고딕) 글꼴에 없는 기호(∙ 등)를 한글이 대체 글꼴로 그리는 폭을 직접 얻기 위함.
사용: python3 make_glyph_probe.py <출력.hwpx>   → 같은 이름 .json(문단 번호 ↔ 글꼴·크기·글자)도 만든다.
분석: corpus_tools/read_glyph_probe.py <한글 저장본> <json>
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import hwpx_writer as W  # noqa: E402

FACES = [("나눔명조", False), ("나눔명조", True), ("나눔고딕", False), ("나눔고딕", True),
         ("HY울릉도M", False), ("HY헤드라인M", False), ("휴먼명조", False), ("맑은 고딕", False)]
_ASCII = "".join(chr(c) for c in range(0x21, 0x7F))
_EXTRA = ("가ㅇㆍ∙＝·※➊➋❶①②③◈◇□■▶▷►▸▹◆○●◎☞⇨➡⇒→➔‣‧・･｢｣「」『』‘’“”▪•…～±×÷°℃"
          "➀➁➂" "美日他旣現全中內韓社前有月高錢舊對相計先令新可完")
CHARS = "".join(dict.fromkeys(_ASCII + _EXTRA))   # 중복 제거, 순서 유지
SIZES = (6, 10)
N = 1200   # 반복 상한


def main(out):
    w = W.Writer({"sections": [{"children": []}]}, W.skeleton())
    rows, body = [], []
    for face, bold in FACES:
        for pt in SIZES:
            for ch in CHARS:
                ps = W.PS(align="LEFT", ls_type="PERCENT", ls_val=100)
                ps.charwrap = True
                cs = W.CS(face=face, pt=float(pt), bold=bold)
                # 추정 폭의 절반까지 좁아도 두 줄은 넘게(한 줄 개수가 두 번 이상 나오도록)
                n = min(N, max(40, int(2.6 * W.BODY_W / max(1.0, 0.5 * cs.adv(ch)))))
                p = W.Para(ps, [[cs, ch * n]], W.BODY_W)
                rows.append(dict(para=len(body), face=face, bold=bold, pt=pt, ch=ch, n=n))
                body.append(w.para_xml(p, (w.sec_pr() + w.header_ctrl()) if not body else ""))
    section = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {W.NS}>' + "".join(body) + "</hs:sec>").encode()
    W._pack(out, w.skel, w, section, w.st.build(), "glyph probe", "글자 폭 측정")
    json.dump(dict(width=W.BODY_W, rows=rows), open(os.path.splitext(out)[0] + ".json", "w", encoding="utf-8"), ensure_ascii=False)
    print(out, len(rows), "문단")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "glyph_probe.hwpx")
