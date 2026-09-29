import glob, os, re, json, zipfile, io, warnings, traceback
warnings.filterwarnings("ignore")
from hwpx.hwp5.package import convert, to_hwpx_bytes
out = {}
files = sorted(glob.glob('orig/**/*.hwp', recursive=True) + glob.glob('orig/**/*.hwpx', recursive=True))
for i, f in enumerate(files):
    key = f"d{i:03d}"
    rec = {"src": f}
    try:
        if f.endswith('.hwp'):
            c = convert(open(f, 'rb').read())
            b = to_hwpx_bytes(c.files)
            rec["dropped"] = {k: v for k, v in (c.report.dropped + c.report.unconverted).items()}
        else:
            b = open(f, 'rb').read()
        open(f"conv/{key}.hwpx", "wb").write(b)
        z = zipfile.ZipFile(io.BytesIO(b))
        secs = sorted(n for n in z.namelist() if re.match(r'Contents/section\d+\.xml', n))
        xml = "".join(z.read(n).decode('utf-8') for n in secs)
        txt = re.sub(r'<[^>]+>', '', "".join(re.findall(r'<hp:t>(.*?)</hp:t>|<hp:t/>', xml)))
        rec.update(ok=True, chars=len(txt), linesegs=xml.count('<hp:lineseg '), paras=xml.count('<hp:p '),
                   text=txt[:3000])
    except Exception as e:
        rec.update(ok=False, err=f"{type(e).__name__}: {e}"[:200])
    out[key] = rec
json.dump(out, open('conv/index.json', 'w'), ensure_ascii=False, indent=1)
ok = [r for r in out.values() if r.get("ok")]
print("total", len(out), "ok", len(ok), "fail", len(out) - len(ok))
for k, r in out.items():
    if not r.get("ok"): print("FAIL", r["src"], r["err"])
print("with linesegs", sum(1 for r in ok if r["linesegs"]), "chars sum", sum(r["chars"] for r in ok))
print("dropped features:", {k for r in ok for k in r.get("dropped", {})})
