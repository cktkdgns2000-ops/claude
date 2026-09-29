# -*- coding: utf-8 -*-
"""생성본 hwpx 일괄 '한글로 열어 저장' — Windows + 한글 설치 PC용

한글이 저장해야 파일 안에 실제 줄 배치 기록이 생긴다. 이 기록으로 생성기 계산과 한글 결과를 문단마다 대조한다.

사용법:
    1) (처음 한 번) 명령 프롬프트에서:  pip install pywin32
    2) 받은 zip을 아무 폴더에나 풀고, 그 안의 이 파일을 더블클릭(또는 python hwp_batch_save.py)
       → 이 파일이 있는 폴더의 *.hwpx를 하나씩 열어 같은 폴더의 '한글저장' 폴더에 같은 이름으로 저장한다.
         (이전에 만든 '한글저장' 폴더가 있어도 같은 이름 파일은 덮어쓴다)
    3) '한글저장' 폴더를 zip으로 묶어 올려주면 된다.
    다른 폴더를 처리하려면:  python hwp_batch_save.py <폴더 경로>

처음 실행 시 한글이 파일 접근 허용 여부를 물으면 '모두 허용'을 누른다.
"""
import glob
import os
import sys
import time
import zipfile

# 기본 폴더 = 이 파일이 있는 폴더(더블클릭해도, 어느 폴더에 풀어도 동작)
FOLDER = os.path.dirname(os.path.abspath(__file__))


def has_layout(path):
    """저장본에 한글 줄 배치 기록(lineseg)이 들어갔는지"""
    try:
        with zipfile.ZipFile(path) as z:
            return any(b"<hp:lineseg " in z.read(n) for n in z.namelist() if n.startswith("Contents/section"))
    except Exception:
        return False


def layout_all(hwp):
    """문서 끝까지 이동하고 쪽 수를 읽어 한글이 전체 쪽 배치를 계산하게 함(열자마자 저장하면 배치 기록이 빠짐)"""
    hwp.HAction.Run("MoveDocEnd")
    try:
        hwp.KeyIndicator()          # 현재 쪽 번호 계산 → 끝까지 배치
    except Exception:
        pass
    _ = hwp.PageCount
    hwp.HAction.Run("MoveDocBegin")


def main(folder):
    import win32com.client as win32

    folder = os.path.abspath(folder)
    out_dir = os.path.join(folder, "한글저장")
    os.makedirs(out_dir, exist_ok=True)
    files = sorted(glob.glob(os.path.join(folder, "*.hwpx")))   # 이 폴더에 바로 있는 파일만(하위 폴더·'한글저장'은 제외)
    if not files:
        raise SystemExit("hwpx 파일이 없습니다: " + folder)
    print(f"{len(files)}개 파일 처리 → {out_dir}")

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")   # 보안 모듈이 있으면 확인 창 생략
    except Exception:
        pass
    hwp.XHwpWindows.Item(0).Visible = True
    failed = []
    for i, src in enumerate(files, 1):
        dst = os.path.join(out_dir, os.path.basename(src))
        try:
            if not hwp.Open(src, "HWPX", "forceopen:true"):
                raise RuntimeError("열기 실패")
            for attempt in range(4):
                layout_all(hwp)
                time.sleep(0.5 + attempt)
                hwp.SaveAs(dst, "HWPX", "")
                if has_layout(dst):
                    break
            else:
                raise RuntimeError("줄 배치 기록 없이 저장됨")
            hwp.Clear(1)   # 저장하지 않고 닫기(이미 저장함)
            print(f"  [{i}/{len(files)}] {os.path.basename(src)}")
        except Exception as e:
            failed.append((os.path.basename(src), str(e)))
            print(f"  [{i}/{len(files)}] 실패: {os.path.basename(src)} ({e})")
    hwp.Quit()
    print(f"\n완료: {len(files) - len(failed)}개 저장, 실패 {len(failed)}개")
    for name, err in failed:
        print("  실패:", name, err)
    input("\nEnter를 누르면 창이 닫힙니다...")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else FOLDER)
