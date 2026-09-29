#!/usr/bin/env python3
"""빠른 내어쓰기(Shift+Tab, ParagraphShapeIndentAtCaret) 보정용 hwpx 생성.
글꼴·크기·자간·앞 공백·기호 조합별 문단을 만들고, 각 문단의 '기호 뒤 커서 위치'를 calib.json에 기록한다.
사용자 PC에서 hwp_indent_at_caret.py로 한글의 빠른 내어쓰기를 실행·저장한 뒤, 저장본의 내어쓰기 값을 읽어 표로 만든다."""
import json, os, sys, io
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwpx_writer as HW

COMBOS = [("나눔명조", 15), ("나눔명조", 13), ("나눔고딕", 13), ("나눔고딕", 12), ("나눔고딕", 11), ("나눔고딕", 10), ("HY울릉도M", 15), ("HY울릉도M", 13)]
MARKERS = ["□", "ㅇ", "-", "*", "**", "※", "➊", "①", "❶", "◈", "➡", "⇨", "■", "‣", "▶", "∙", "·", "◇", "1)"]
BODY = "(키워드) 가나다라마바사아자차카타파하 가나다라마바사아자차카타파하 가나다라마바사아자차카타파하 가나다라마바사 두 줄 확인용 문장"


def main(out_dir):
    w = HW.Writer({"sections": [{"children": []}]}, HW.skeleton())
    # 첫 문단은 구역 설정이 들어가므로 안내 문단으로 두고, 보정 문단은 둘째 문단(번호 1)부터
    paras = [HW.Para(HW.PS(align="LEFT"), [[HW.CS(face="나눔고딕", pt=11), "빠른 내어쓰기 보정 파일 — hwp_indent_at_caret.py로 처리 후 저장"]], HW.BODY_W)]
    rows = []
    for face, pt in COMBOS:
        for sp in (0, -5):
            for mk in MARKERS:
                for nsp in (0, 1, 3):
                    prefix = " " * nsp + mk + " "
                    cs = HW.CS(face=face, pt=pt, spacing=sp)
                    ps = HW.PS(align="JUSTIFY", ls_type="PERCENT", ls_val=160)
                    paras.append(HW.Para(ps, [[cs, prefix + BODY]], HW.BODY_W))
                    rows.append(dict(para=len(paras) - 1, face=face, pt=pt, spacing=sp, marker=mk, nsp=nsp, caret=len(prefix)))
    xml = [w.para_xml(p, w.sec_pr() if i == 0 else "") for i, p in enumerate(paras)]
    section = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {HW.NS}>' + "".join(xml) + "</hs:sec>").encode()
    header = w.st.build()
    os.makedirs(out_dir, exist_ok=True)
    HW._pack(os.path.join(out_dir, "indent_calib.hwpx"), HW.skeleton(), w, section, header, "보정", "빠른 내어쓰기 보정")
    json.dump(rows, open(os.path.join(out_dir, "indent_calib.json"), "w", encoding="utf-8"), ensure_ascii=False)
    print(len(rows), "문단")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
