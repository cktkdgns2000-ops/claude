# -*- coding: utf-8 -*-
"""한글 '빠른 내어쓰기'(Shift+Tab = ParagraphShapeIndentAtCaret) 일괄 실행 — Windows + 한글 설치 PC용

사용법(명령 프롬프트):
    pip install pywin32
    python hwp_indent_at_caret.py indent_calib.hwpx indent_calib.json
→ 같은 폴더에 indent_calib_한글저장.hwpx 가 생긴다. 이 파일을 올려주면 된다.

처리 내용: json에 적힌 문단마다 '기호 뒤' 커서 위치로 이동해 빠른 내어쓰기를 실행하고 hwpx로 저장.
처음 실행 시 한글이 파일 접근 허용 여부를 물으면 '모두 허용'을 누른다.
"""
import json
import os
import sys

import win32com.client as win32


def main(src, rows_path):
    src = os.path.abspath(src)
    dst = os.path.splitext(src)[0] + "_한글저장.hwpx"
    rows = json.load(open(rows_path, encoding="utf-8"))
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")   # 보안 모듈이 설치돼 있으면 확인 창 생략
    except Exception:
        pass
    hwp.XHwpWindows.Item(0).Visible = True
    if not hwp.Open(src, "HWPX", "forceopen:true"):
        sys.exit("파일을 열지 못했습니다: " + src)
    done = 0
    for r in rows:
        # SetPos(목록, 문단 번호, 글자 위치): 본문 목록 0, 기호 뒤 위치에 커서
        if hwp.SetPos(0, r["para"], r["caret"]):
            hwp.HAction.Run("ParagraphShapeIndentAtCaret")
            done += 1
    hwp.SaveAs(dst, "HWPX", "")
    print(f"빠른 내어쓰기 {done}/{len(rows)}개 문단 처리 → {dst}")
    hwp.Quit()


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "indent_calib.hwpx")
    rows = sys.argv[2] if len(sys.argv) > 2 else os.path.join(here, "indent_calib.json")
    main(src, rows)
