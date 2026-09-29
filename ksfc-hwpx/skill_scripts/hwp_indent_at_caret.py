# -*- coding: utf-8 -*-
"""한글 '빠른 내어쓰기'(Shift+Tab = ParagraphShapeIndentAtCaret) 일괄 실행 — Windows + 한글 설치 PC용

사용법:
    1) (처음 한 번) 명령 프롬프트에서:  pip install pywin32
    2) 이 파일을 더블클릭하거나, 명령 프롬프트에서:  python hwp_indent_at_caret.py
       → 아래 FOLDER에서 indent_calib.hwpx / indent_calib.json 을 찾아 처리하고
         같은 폴더에 indent_calib_한글저장.hwpx 를 만든다. 이 파일을 올려주면 된다.
    다른 위치의 파일을 쓰려면:  python hwp_indent_at_caret.py <hwpx 경로> <json 경로>

처음 실행 시 한글이 파일 접근 허용 여부를 물으면 '모두 허용'을 누른다.
"""
import glob
import json
import os
import sys

FOLDER = r"C:\Users\ksfc\Desktop\다운로드\빠른내어쓰기_보정"


def find(name):
    """FOLDER(하위 폴더 포함) → 이 스크립트가 있는 폴더 순으로 파일을 찾는다."""
    here = os.path.dirname(os.path.abspath(__file__))
    for base in (FOLDER, here):
        direct = os.path.join(base, name)
        if os.path.exists(direct):
            return direct
        hits = glob.glob(os.path.join(base, "**", name), recursive=True)
        if hits:
            return hits[0]
    return None


def main(src, rows_path):
    import win32com.client as win32

    src = os.path.abspath(src)
    dst = os.path.join(os.path.dirname(src), "indent_calib_한글저장.hwpx")
    rows = json.load(open(rows_path, encoding="utf-8"))
    print("보정 파일:", src)
    print("문단 목록:", rows_path, f"({len(rows)}개)")

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")   # 보안 모듈이 있으면 확인 창 생략
    except Exception:
        pass
    hwp.XHwpWindows.Item(0).Visible = True
    if not hwp.Open(src, "HWPX", "forceopen:true"):
        raise SystemExit("파일을 열지 못했습니다: " + src)

    done, failed = 0, []
    for r in rows:
        # SetPos(목록, 문단 번호, 글자 위치): 본문 목록 0, 기호 뒤(= Shift+Tab을 누를 위치)에 커서
        if hwp.SetPos(0, r["para"], r["caret"]):
            hwp.HAction.Run("ParagraphShapeIndentAtCaret")
            done += 1
        else:
            failed.append(r["para"])
    hwp.SaveAs(dst, "HWPX", "")
    hwp.Quit()
    print(f"\n빠른 내어쓰기 {done}/{len(rows)}개 문단 처리")
    if failed:
        print("커서 이동 실패 문단:", failed[:20], "…" if len(failed) > 20 else "")
    print("저장 완료 →", dst)


if __name__ == "__main__":
    try:
        src = sys.argv[1] if len(sys.argv) > 1 else find("indent_calib.hwpx")
        rows = sys.argv[2] if len(sys.argv) > 2 else find("indent_calib.json")
        if not src or not rows:
            raise SystemExit(f"보정 파일을 찾지 못했습니다. 폴더를 확인하세요:\n  {FOLDER}\n"
                             "  (indent_calib.hwpx, indent_calib.json 이 있어야 합니다)")
        main(src, rows)
    except SystemExit as e:
        print(e)
    except Exception as e:
        print("오류:", type(e).__name__, e)
    input("\n엔터를 누르면 창이 닫힙니다...")
