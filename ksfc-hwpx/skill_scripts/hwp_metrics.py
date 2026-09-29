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
# 빠른 내어쓰기(ParagraphShapeIndentAtCaret) 위치 계산용 기호 폭(em) — 보정 912문단에서 역산(줄 배치는 측정 폭 glyph_em.json 기준)
# 한글의 내어쓰기 위치는 기호 실제 폭보다 1.5~3% 작게 잡힘(나눔명조 □ 폭 0.951, 내어쓰기 0.935). 911/912문단 0.5pt 이내
MARKER_EM = {
    ('HY울릉도M', '*'): 0.575,
    ('HY울릉도M', '-'): 0.855,
    ('HY울릉도M', '·'): 0.99,
    ('HY울릉도M', '‣'): 0.456,   # 새 PC 폭 0.4671 - 글꼴 차이 0.011 (PC마다 대체 글꼴이 달라 새 PC 기준)
    ('HY울릉도M', '※'): 0.99,
    ('HY울릉도M', '⇨'): 0.96,
    ('HY울릉도M', '∙'): 0.488,   # 새 PC 폭 0.4995 - 글꼴 차이 0.011 (PC마다 대체 글꼴이 달라 새 PC 기준)
    ('HY울릉도M', '①'): 0.99,
    ('HY울릉도M', '■'): 0.99,
    ('HY울릉도M', '□'): 0.99,
    ('HY울릉도M', '▶'): 0.99,
    ('HY울릉도M', '◇'): 0.99,
    ('HY울릉도M', '◈'): 0.99,
    ('HY울릉도M', '❶'): 0.96,
    ('HY울릉도M', '➊'): 0.96,
    ('HY울릉도M', '➡'): 0.96,
    ('HY울릉도M', 'ㅇ'): 0.99,
    ('나눔고딕', '*'): 0.575,
    ('나눔고딕', '-'): 0.335,
    ('나눔고딕', '·'): 0.26,
    ('나눔고딕', '‣'): 0.439,   # 새 PC 폭 0.4685 - 글꼴 차이 0.030 (PC마다 대체 글꼴이 달라 새 PC 기준)
    ('나눔고딕', '※'): 0.91,
    ('나눔고딕', '⇨'): 0.91,
    ('나눔고딕', '∙'): 0.47,   # 새 PC 폭 0.4995 - 글꼴 차이 0.030 (PC마다 대체 글꼴이 달라 새 PC 기준)
    ('나눔고딕', '①'): 0.91,
    ('나눔고딕', '■'): 0.91,
    ('나눔고딕', '□'): 0.91,
    ('나눔고딕', '▶'): 0.91,
    ('나눔고딕', '◇'): 0.91,
    ('나눔고딕', '◈'): 0.91,
    ('나눔고딕', '❶'): 0.91,
    ('나눔고딕', '➊'): 0.91,
    ('나눔고딕', '➡'): 0.94,
    ('나눔고딕', 'ㅇ'): 0.91,
    ('나눔명조', '*'): 0.62,
    ('나눔명조', '-'): 0.62,
    ('나눔명조', '·'): 0.27,
    ('나눔명조', '‣'): 0.452,   # 새 PC 폭 0.4671 - 글꼴 차이 0.015 (PC마다 대체 글꼴이 달라 새 PC 기준)
    ('나눔명조', '※'): 0.935,
    ('나눔명조', '⇨'): 0.935,
    ('나눔명조', '∙'): 0.484,   # 새 PC 폭 0.4995 - 글꼴 차이 0.015 (PC마다 대체 글꼴이 달라 새 PC 기준)
    ('나눔명조', '①'): 0.935,
    ('나눔명조', '■'): 0.935,
    ('나눔명조', '□'): 0.935,
    ('나눔명조', '▶'): 0.935,
    ('나눔명조', '◇'): 0.935,
    ('나눔명조', '◈'): 0.935,
    ('나눔명조', '❶'): 0.935,
    ('나눔명조', '➊'): 0.935,
    ('나눔명조', '➡'): 0.95,
    ('나눔명조', 'ㅇ'): 0.935,
}

HY_NONHANGUL = 1.1

NO_START = set(")]}’”,.:;!?、。」』〉》%")   # 가운뎃점(ㆍ·)은 한글이 줄 첫머리에 둠(원본 10곳, 생성본 저장본 2곳)
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


def _load_measured():
    """한글에서 잰 글자 폭(glyph_em.json): 글꼴 파일이 없는 글꼴(HY·휴먼명조·맑은 고딕)과 나눔 글꼴에 없는 글자.
    make_glyph_probe.py → 한글 저장 → corpus_tools/read_glyph_probe.py. 키 '글꼴|굵게(0/1)|글자'"""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "glyph_em.json")
    try:
        import json
        return json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}


MEASURED = _load_measured()


@lru_cache(None)
def marker_em(face, bold, ch):
    """내어쓰기(빠른 내어쓰기 위치) 계산용 기호 폭: 보정 912문단에서 역산한 값(MARKER_EM), 없으면 em()"""
    return MARKER_EM.get((face, ch), em(face, bold, ch))


@lru_cache(None)
def em(face, bold, ch):
    """글자 전진폭(em). 띄어쓰기는 0.5em. 한글 측정값 → 글꼴 파일 → 기본값 순."""
    if ch == " ":
        return 0.5
    for k in (f"{face}|{int(bool(bold))}|{ch}", f"{face}|0|{ch}"):
        if k in MEASURED:
            return MEASURED[k]
    f = _font(face, bold)
    if f:
        cmap, hm, upm = f
        g = cmap.get(ord(ch))
        if g is not None:
            w = hm[g][0] / upm
            # HY 글꼴(미리보기 파일 폭)의 한글 음절 외 글자는 한글에서 약 10% 넓게 그려짐:
            # 한글 저장본 4·5·6차 12건에서 HY 영문·숫자·기호가 든 줄 90개의 줄 나눔이 1.1배일 때 전부 일치(1.0배 2건, 1.15배 1건 어긋남)
            if face.startswith("HY") and not ("가" <= ch <= "힣" or "\u3130" <= ch <= "\u318f"):   # 한글 음절·자모 제외
                w *= HY_NONHANGUL
            return w
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
