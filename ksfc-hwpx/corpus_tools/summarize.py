import pickle
from collections import Counter, defaultdict
rows, tables, pages = pickle.load(open("conv/analysis.pkl","rb"))
def mm(hu): return round(hu/7200*25.4,1)
def top(c, n=3, docs=None):
    tot=sum(c.values()); return ", ".join(f"{k} {v*100//tot}%" for k,v in c.most_common(n))
print("## 쪽 여백(mm) 위+머리말 / 아래+꼬리말 / 좌 / 우")
c=Counter((mm(p['top']+p['header']), mm(p['bottom']+p['footer']), mm(p['left']), mm(p['right'])) for _,p in pages)
print(top(c,5))
c=Counter((mm(p['top']),mm(p['header'])) for _,p in pages); print(" top/header split:", top(c,4))
by=defaultdict(list)
for r in rows:
    if r['text'].strip(): by[r['cls']].append(r)
print("\n## 문단 종류별 (문단 수 / 문서 수)")
for cls in ["□","ㅇ","-","·","*","**","※","➊","①","◈","➡","■","1.","Ⅰ","<표제>","기타"]:
    rs=[r for r in by.get(cls,[]) if r['para'] and r['char']]
    if not rs: continue
    nd=len({r['doc'] for r in rs})
    f=Counter(f"{r['char']['font']} {r['char']['pt']:g}pt" for r in rs)
    ls=Counter(f"{r['para']['ls'][0]} {r['para']['ls'][1]}" for r in rs)
    ind=Counter(f"L{mm(r['para']['left'])}/I{mm(r['para']['intent'])}" for r in rs)
    prev=Counter(f"{r['para']['prev']/100:g}pt" for r in rs)
    al=Counter(r['para']['align'] for r in rs)
    pit=Counter(f"{p/100:g}" for r in rs for p in r['pitch'][:-1] if r['lines']>1)
    sp=Counter(r['char']['spacing'] for r in rs); ra=Counter(r['char']['ratio'] for r in rs)
    tabs=Counter(r['tabs']>0 for r in rs)
    fs=Counter(r['char']['fontspace'] for r in rs); nl=Counter(r['para']['nonlatin'] for r in rs); cd=Counter(r['para']['condense'] for r in rs)
    mf=Counter(f"{r['mchar']['font']} {r['mchar']['pt']:g}" for r in rs if r['mchar'])
    print(f"\n[{cls}] {len(rs)}문단/{nd}건")
    print("  글꼴:",top(f)); print("  기호 글꼴:",top(mf)); print("  줄간격:",top(ls)); print("  실측 줄피치(pt):",top(pit))
    print("  들여쓰기(mm):",top(ind,4)); print("  문단 위:",top(prev,4)); print("  정렬:",top(al,2))
    print("  자간:",top(sp,4),"| 장평:",top(ra,3),"| 탭 사용:",top(tabs,2))
    print("  빈칸(useFontSpace):",top(fs,2),"| 한글 줄나눔:",top(nl,2),"| 최소공백:",top(cd,2))
