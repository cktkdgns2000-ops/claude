"""한글에서 저장한 시험본의 줄 위치(lineseg textpos)와 생성기 줄 나눔 모델(hwp_metrics.layout) 대조.
사용: python3 compare_hancom_saved.py <한글 저장본.hwpx> ...   (표·그림 없는 문단만 비교)"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "skill_scripts"))
from hwpx_layout import Doc, HP
tot=bad=0
for f in sys.argv[1:]:
    d=Doc(f,use_stored=True)
    for sec,W,H in d.sections:
        def visit(p,width):
            global tot,bad
            segs=p.findall(f"{HP}linesegarray/{HP}lineseg")
            if not segs: return
            txt=''.join(t.text or '' for t in p.iter(f"{HP}t"))
            if p.find(f".//{HP}tbl") is not None or p.find(f".//{HP}pic") is not None or not txt.strip(): return
            lines,pr=d.para_lines(p,width)
            st=[int(s.get('textpos')) for s in segs]
            ms=[l[2] for l in lines]
            tot+=1
            if st!=ms:
                bad+=1
                full=''.join(''.join((t.text or '') + ''.join(' '+(x.tail or '') for x in t) ) for t in p.iter(f"{HP}t"))
                print(f.split('/')[-1][:14], 'hancom',[full[a:a+6] for a in st[1:]], 'model',[full[a:a+6] for a in ms[1:]], full[:30])
        for p in sec.iter(f"{HP}p"):
            # determine width: top-level vs cell
            par=p.getparent()
            if par.tag==f"{HP}subList":
                tc=par.getparent(); cs=tc.find(f"{HP}cellSz"); cm=tc.find(f"{HP}cellMargin")
                w=int(cs.get('width'))-int(cm.get('left'))-int(cm.get('right'))
            else: w=W
            visit(p,w)
print('total',tot,'mismatch',bad)
