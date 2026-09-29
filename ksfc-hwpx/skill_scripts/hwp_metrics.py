"""한글(Hancom) 배치 규칙에 따른 글자 폭·줄 나눔 모델.

- 글자 폭 = 글꼴 전진폭 × 크기 × 장평 × (1 + 자간%)   (자간은 각 글자 폭에 대한 %)
- 띄어쓰기 = 0.5em (글꼴에 어울리는 빈칸 끔, 원본 51건 전부)
- 한글 줄 나눔: BREAK_WORD = 어절 단위, KEEP_WORD = 글자 단위 (한글에서 이름과 반대로 동작)
- 줄 끝 띄어쓰기는 여백 밖으로 걸침, 금칙(닫는 부호로 시작·여는 부호로 끝나는 줄 금지)
원본 여러 줄 문단 178개 중 147개에서 한글이 실제로 나눈 줄 위치를 모두 재현(corpus_tools/breakmodel.py).
"""
import os
import subprocess
from functools import lru_cache

NANUM = "/usr/share/fonts/truetype/nanum"
FONT_FILES = {
    ("나눔명조", False): f"{NANUM}/NanumMyeongjo.ttf", ("나눔명조", True): f"{NANUM}/NanumMyeongjoBold.ttf",
    ("나눔고딕", False): f"{NANUM}/NanumGothic.ttf", ("나눔고딕", True): f"{NANUM}/NanumGothicBold.ttf",
}
NO_START = set(")]}’”,.:;!?ㆍ·、。」』〉》%")
NO_END = set("([{‘“「『〈《")


@lru_cache(None)
def _font(face, bold):
    from fontTools.ttLib import TTFont
    path = FONT_FILES.get((face, bold)) or FONT_FILES.get((face, False))
    if not path:
        try:
            path = subprocess.run(["fc-match", "-f", "%{file}", face], capture_output=True, text=True).stdout
        except Exception:
            path = ""
    if not path or not os.path.exists(path):
        return None
    t = TTFont(path)
    return t.getBestCmap(), t["hmtx"].metrics, t["head"].unitsPerEm


def _wide(ch):
    o = ord(ch)
    return (0xAC00 <= o <= 0xD7A3 or 0x1100 <= o <= 0x11FF or 0x3130 <= o <= 0x318F or 0x3000 <= o <= 0x303F
            or 0xFF00 <= o <= 0xFFEF or 0x4E00 <= o <= 0x9FFF or 0x2190 <= o <= 0x21FF or 0x2460 <= o <= 0x24FF
            or 0x25A0 <= o <= 0x27BF or o in (0x203B,))


@lru_cache(None)
def em(face, bold, ch):
    """글자 전진폭(em). 띄어쓰기는 0.5em."""
    if ch == " ":
        return 0.5
    f = _font(face, bold)
    if f:
        cmap, hm, upm = f
        g = cmap.get(ord(ch))
        if g is not None:
            return hm[g][0] / upm
    return 1.0 if _wide(ch) else 0.55


def advance(ch, face, bold, pt, ratio=100, spacing=0):
    """HWPUNIT(1/7200in). pt 글자 → 100 HU/pt."""
    return em(face, bold, ch) * pt * 100 * ratio / 100 * (1 + spacing / 100)


def layout(chars, first_w, rest_w, mode="word"):
    """chars: [(ch, adv_hu)] → (줄 시작 위치 목록, 줄별 사용 폭(끝 띄어쓰기 제외), 줄별 띄어쓰기 수)."""
    n = len(chars)
    starts, used, gaps = [0], [], []
    s = 0
    while s < n:
        w = first_w if len(starts) == 1 else rest_w
        acc, e = 0.0, s
        while e < n:
            ch, a = chars[e]
            if ch != " " and acc + a > w + 0.5 and e > s:
                break
            acc += a
            e += 1
        if e >= n:
            line = chars[s:n]
        else:
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
        used.append(sum(a for _, a in line))
        gaps.append(sum(1 for c, _ in line[1:] if c == " "))
        if e >= n:
            break
        s = b
        while s < n and chars[s][0] == " ":
            s += 1
        if s < n:
            starts.append(s)
    return starts, used, gaps
