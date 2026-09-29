"""빠른 내어쓰기 보정 기호 폭 재산출: 줄 배치는 측정 폭(glyph_em)으로 두고, 기호 폭 하나로 내어쓰기를 가장 잘 맞추는 값.
사용: python3 refit_marker_em.py <calib_result.json> <보정 한글 저장본.hwpx> <read_glyph_probe 결과.json>"""
import sys,json,collections
import os; sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'skill_scripts'))
import hwp_metrics as HM
from hwpx_layout import Doc, HP
rows=json.load(open(sys.argv[1]))
d=Doc(sys.argv[2],use_stored=True); sec,Wd,H=d.sections[0]
paras=sec.findall(f"{HP}p")
g=json.load(open(sys.argv[3]))
groups=collections.defaultdict(list)
for r in rows:
    if not r.get('textok') or r.get('intent') is None: continue
    p=paras[r['para']]
    seq=[]
    for run in p.findall(f"{HP}run"):
        c=d.chars[run.get('charPrIDRef')]
        for x in run.iter(f"{HP}t"):
            for ch in x.text or '': seq.append((ch,c))
    groups[(r['face'],r['marker'])].append((r,seq))
def predict(r,seq,m):
    mk=r['marker']; nsp=r['nsp']
    chars=[]
    for i,(ch,c) in enumerate(seq):
        if nsp<=i<nsp+len(mk):
            a=m*c['pt']*100*(1+c['sp']/100)/len(mk)   # 기호 전체 폭 m(em)을 글자 수로 나눔
        else:
            a=HM.advance(ch,c['face'],c['bold'],c['pt'],c['ratio'],c['sp'])
        chars.append((ch,a))
    c0=seq[0][1]
    prefix=sum(a for ch,a in chars[:nsp+len(mk)+1])
    starts,used,gaps=HM.layout(chars,Wd,Wd+r['intent'] if False else Wd,'word')
    # 첫 줄 폭은 전체(내어쓰기 음수), 둘째 줄부터 W - |intent| 이지만 첫 줄 배치에는 영향 없음
    extra=0
    if len(starts)>1:
        t=''.join(ch for ch,_ in chars); line=t[:starts[1]].rstrip(' ')
        ngap=line.count(' ')-nsp
        if ngap>0: extra=(Wd-used[0])/ngap*1
    return -(prefix+extra)
res={}
tot_ok=tot=0
for (face,mk),items in sorted(groups.items()):
    best=None
    for i in range(0,241):
        m=0.2+i*0.005
        m_tot=m*len(mk) if len(mk)>1 else m
        errs=[abs(predict(r,seq,m_tot)-r['intent'])/100 for r,seq in items]
        sc=(sum(e>0.5 for e in errs),sum(errs))
        if best is None or sc<best[0]: best=(sc,m,max(errs))
    (nbad,s),m,mx=best
    pk=g.get(f"{face}|0|{mk}",{}).get('em')
    res[(face,mk)]=m
    tot+=len(items); tot_ok+=len(items)-nbad
    print(f"{face} {mk!r}: 내어쓰기 맞춤 {m:.3f} (0.5pt 초과 {nbad}/{len(items)}, 최대 {mx:.2f}pt) | 측정문서 {pk} | 기존 {HM.MARKER_EM.get((face,mk))}")
print('total ok',tot_ok,'/',tot)
