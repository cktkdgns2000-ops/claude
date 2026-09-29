"""HY 글꼴 한글 음절 외 글자의 폭 배율 추정: 한글 저장본의 줄 나눔(맞은 줄은 폭 이하, 넘긴 줄은 다음 어절까지 폭 초과)을
가장 잘 만족하는 배율을 찾음. 사용: 저장본 폴더 U를 고쳐서 실행."""
import os, sys, glob
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "skill_scripts"))
from hwpx_layout import Doc, HP
import hwp_metrics as HM
U='/root/.claude/uploads/9e47dde0-71e9-58ae-85bb-5e9c8ce17a56/'
files=[f for f in glob.glob(U+'*.hwpx') if '__4_' in f or '__5_' in f or '__6_' in f]
FACES=('HY울릉도M','HY헤드라인M')
recs=[]  # (face_share_hy, other_w, hy_w, W, wrapped(bool: line+next > W) )
for f in files:
    try: d=Doc(f,use_stored=True)
    except Exception: continue
    for sec,W,H in d.sections:
        for p in sec.iter(f"{HP}p"):
            if p.find(f".//{HP}tbl") is not None or p.find(f".//{HP}pic") is not None: continue
            segs=[int(s.get('textpos')) for s in p.findall(f"{HP}linesegarray/{HP}lineseg")]
            if len(segs)<2: continue
            par=p.getparent(); w=W
            if par.tag==f"{HP}subList":
                tc=par.getparent(); cs=tc.find(f"{HP}cellSz"); cm=tc.find(f"{HP}cellMargin")
                if cs is None or cm is None: continue
                w=int(cs.get('width'))-int(cm.get('left'))-int(cm.get('right'))
            pr=d.paras.get(p.get('paraPrIDRef'))
            if not pr or pr['ls'][0]!='PERCENT': pass
            seq=[]
            ok=True
            for r in p.findall(f"{HP}run"):
                c=d.chars[r.get('charPrIDRef')]
                for el in r:
                    tag=el.tag.split('}')[1]
                    if tag=='t':
                        if len(el): ok=False
                        for ch in el.text or '':
                            a=HM.advance(ch,c['face'],c['bold'],c['pt'],c['ratio'],c['sp'])*(0.6 if c['sup'] else 1)
                            seq.append((ch,a,(c['face'] in FACES) and not ('가'<=ch<='힣') and ch!=' ' ))
                    elif tag not in ('linesegarray',): ok=False
            if not ok or len(seq)<segs[-1]: continue
            W0=w-pr['left']-pr['right']; fw,rw=W0-max(pr['intent'],0),W0-max(-pr['intent'],0)
            b=segs+[len(seq)]
            for i in range(len(segs)-1):
                line=seq[b[i]:b[i+1]]
                while line and line[-1][0]==' ': line=line[:-1]
                # next unit: word or char
                j=b[i+1]
                if pr['mode']=='word':
                    k=j
                    while k<len(seq) and seq[k][0]!=' ': k+=1
                    nxt=seq[b[i]:k]
                else:
                    nxt=seq[b[i]:j+1]
                WW=fw if i==0 else rw
                hy=sum(x[1] for x in line if x[2]); ot=sum(x[1] for x in line if not x[2])
                hy2=sum(x[1] for x in nxt if x[2]); ot2=sum(x[1] for x in nxt if not x[2])
                recs.append((hy,ot,hy2,ot2,WW,f,''.join(x[0] for x in line)[-12:],''.join(x[0] for x in nxt)[-8:]))

n=sum(1 for r in recs if r[0]>0 or r[2]>0)
print('records with HY non-hangul',n)
for k in [0.9,1.0,1.05,1.1,1.15,1.2,1.3]:
    v1=sum(1 for hy,ot,hy2,ot2,W,f,*_ in recs if hy>0 and hy*k+ot>W)
    v2=sum(1 for hy,ot,hy2,ot2,W,f,*_ in recs if hy2>0 and hy2*k+ot2<=W)
    print(k,v1,v2)
v2=sum(1 for hy,ot,hy2,ot2,W,f,*_ in recs if hy2==0 and ot2<=W); v1=sum(1 for hy,ot,hy2,ot2,W,f,*_ in recs if hy==0 and ot>W)
print('no-HY lines: fit viol',v1,'wrap viol',v2,'of',sum(1 for r in recs if r[2]==0))

print('--- fills where Hancom wrapped (model line+next <= W), and fits >99%')
for hy,ot,hy2,ot2,W,f,a,b in recs:
    if hy2+ot2<=W: print(' WRAP  next-fill %.2f%%  hyshare %.2f %s | %s'%((hy2+ot2)/W*100, hy2/(hy2+ot2), a, b))
for hy,ot,hy2,ot2,W,f,a,b in recs:
    if (hy+ot)/W>0.985: print(' FIT   fill %.2f%%  hyshare %.2f %s'%((hy+ot)/W*100, hy/(hy+ot), a))
