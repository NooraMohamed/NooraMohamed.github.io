"""Build majors/data.js from the UNL 'Majors by College' workbook.
Usage: python3 -I build_data.py path/to/workbook.xlsx > data.js
Each major is summed across colleges/degrees per Fall term, with undergrad/grad split."""
import sys, re, json, collections, openpyxl

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
def year(title):
    return int(re.search(r"(20\d\d)", title).group(1))

def clean(n):
    n = re.sub(r"\s+", " ", n.strip().upper())
    n = re.sub(r"\s*\((RAIKES|RAIKES PROGAM ONLY)\)$", "", n)   # Raikes-program duplicates -> base major
    n = re.sub(r"\s*\(RAIKES PROGAM ONLY\)$", "", n)
    n = re.sub(r"\(OMAHA\)", "(OMAHA)", n)
    return n

def level(college):
    c = college.strip()
    if c.startswith("Grad Studies"): return "g"
    if c == "Law": return "l"
    return "u"

data = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(int)))
colleges = collections.defaultdict(set)
years = []
for ws in wb:
    y = year(ws.title); years.append(y)
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if not r[0] or not r[1] or not isinstance(r[4], (int, float)): continue
        n = clean(r[1])
        data[n][level(r[0])][y] += int(r[4])
        colleges[n].add(r[0].strip())
years.sort()
out = []
for n, lv in data.items():
    tot = [sum(lv[l].get(y, 0) for l in lv) for y in years]
    out.append({"n": n, "c": sorted(colleges[n]),
                "t": tot,
                **{l: [lv[l].get(y, 0) for y in years] for l in lv}})
out.sort(key=lambda d: d["n"])
print("window.MAJORS_YEARS=%s;\nwindow.MAJORS_DATA=%s;" % (json.dumps(years), json.dumps(out, separators=(",", ":"))))
