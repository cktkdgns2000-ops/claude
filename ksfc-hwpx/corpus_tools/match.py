import re, json, zipfile, glob, os, difflib
K="/root/.claude/skills/synced/35c86b7f-299f-4e16-8fcb-fef00f689948_ebddc635-488e-483f-b011-242d3d7ca8ce/ksfc-report"
F=re.findall(r'\("([^"]+)","([^"]+)",(\d+)\)',open(K+"/scripts/regress.py",encoding="utf8").read())
idx=json.load(open("conv/index.json"))
def norm(s): return re.sub(r'[^0-9A-Za-z가-힣]','',s)
def hwpx_text(k):
    z=zipfile.ZipFile(f"conv/{k}.hwpx")
    x="".join(z.read(n).decode() for n in sorted(z.namelist()) if re.match(r'Contents/section\d+\.xml',n))
    return re.sub(r'<[^>]+>','',"".join(re.findall(r'<hp:t>(.*?)</hp:t>',x)))
def md_text(p):
    s=open(p,encoding="utf8").read()
    s=re.sub(r'!\[[^\]]*\]\([^)]*\)','',s); s=re.sub(r'\{[^{}]*\}','',s); s=re.sub(r'^@.*$','',s,flags=re.M)
    return s
def grams(s,n=5): return {s[i:i+n] for i in range(len(s)-n+1)}
docs={k:grams(norm(hwpx_text(k))) for k,r in idx.items() if r.get("ok")}
res=[]
for key,std,_ in F:
    g=grams(norm(md_text(f"{K}/assets/examples/{key}.md")))
    sc=[]
    for k,dg in docs.items():
        cov=len(g&dg)/max(1,len(g))          # 재현 원고가 원본에 담긴 비율
        prec=len(g&dg)/max(1,len(dg))        # 원본이 재현 원고에 담긴 비율
        name=difflib.SequenceMatcher(None,norm(std),norm(os.path.basename(idx[k]["src"]))).ratio()
        sc.append((round(cov,3),round(prec,3),round(name,2),k))
    sc.sort(reverse=True)
    res.append((key,std,sc[:3]))
json.dump(res,open("conv/match.json","w"),ensure_ascii=False,indent=1)
for key,std,top in res:
    b=top[0]; flag=""
    if b[0]<0.6: flag="  ?? 낮음"
    elif top[1][0]>b[0]-0.03: flag="  ?? 후보 근접"
    print(f"{key:36s} cov={b[0]:.2f} prec={b[1]:.2f} name={b[2]:.2f} {os.path.basename(idx[b[3]]['src'])[:60]}{flag}")
    if flag: 
        for t in top[1:]: print(f"{'':36s}   2nd cov={t[0]:.2f} prec={t[1]:.2f} name={t[2]:.2f} {os.path.basename(idx[t[3]]['src'])[:60]}")
