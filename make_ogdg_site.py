"""OGDG web catalogue generator.

Reads the master workbook and writes one self-contained HTML page with
Datasets and Use cases browsing (search, filters, detail pages, hash links).

Usage:  python3 make_ogdg_site.py <workbook.xlsx> [output.html] [report.pdf]

Needs: pip install openpyxl pycountry   (pycountry optional: country names)
Sheets read: Datasets, Use_Cases, Themes, "Carinal 4".
"""
import sys, re, json
import warnings
import openpyxl
warnings.filterwarnings("ignore", module="openpyxl")
try:
    import pycountry
except ImportError:  # names fall back to ISO3 codes
    pycountry = None

XLSX = sys.argv[1] if len(sys.argv) > 1 else "Dataset_Level_updates_16.xlsx"
# Optional 3rd argument: the report PDF. It is embedded in the HTML, so the site stays one file.
REPORT_PDF = sys.argv[3] if len(sys.argv) > 3 else ""
REPORT_NAME = "OGDG_report.pdf"   # file name the visitor's browser saves
OUT = sys.argv[2] if len(sys.argv) > 2 else "OGDG_site.html"

URL = re.compile(r"https?://[^\s;,)\]]+[^\s;,.)\]]")
DSID = re.compile(r"\b([A-Z]{3}-\d{3})\b")
TIER = re.compile(r"^\s*\[(\w+)\]\s*")
CPTAG = re.compile(r"\[(CP\d(?:\s*&\s*CP\d)*)\]")
SPECIAL = {"AFRICA": "Africa", "GLOBAL": "Global", "COD": "DR Congo",
           "TZA": "Tanzania", "SSD": "South Sudan"}


def txt(v):
    return "" if v is None else str(v).strip()


def link_of(cell):
    """Hyperlink behind a cell, else the first URL written in it."""
    if cell.hyperlink is not None and cell.hyperlink.target:
        return cell.hyperlink.target.strip()
    t = txt(cell.value)
    m = URL.search(t)
    if m:
        return m.group(0)
    # plain domain written without https:// (e.g. "globalsolaratlas.info/...")
    m = re.match(r"^(?:www\.)?[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+(?:/\S*)?$", t)
    return "https://" + t if m else ""


def split_top(s, sep=";"):
    """Split on `sep` only outside brackets, so "(a; b)" stays in one item."""
    out, cur, depth = [], "", 0
    for ch in s or "":
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth = max(0, depth - 1)
        if ch == sep and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return [x.strip() for x in out if x.strip()]


def tier(s):
    m = TIER.match(s or "")
    return (m.group(1).capitalize() if m else ""), TIER.sub("", s or "", count=1).strip()


def country(code):
    code = code.strip().upper()
    if code in SPECIAL:
        return SPECIAL[code]
    if pycountry:
        c = pycountry.countries.get(alpha_3=code)
        if c:
            return getattr(c, "common_name", None) or c.name
    return code


def area(cov_spatial):
    return "Africa" if cov_spatial == "Africa" else "Global"


FORMAT_GROUPS = {  # Excel value -> filter group; anything not listed goes to "Other"
    "GeoTIFF": "GeoTIFF", "NetCDF": "NetCDF", "HDF5": "HDF", "HDF-EOS": "HDF", "HDF": "HDF",
    "GRIB": "GRIB", "Shapefile": "Shapefile", "GeoPackage": "GeoPackage", "GeoJSON": "GeoJSON",
    "GeoParquet": "GeoParquet", "Geodatabase": "Geodatabase", "KML": "KML/KMZ", "KMZ": "KML/KMZ",
    "Esri Grid": "Esri raster (Grid/ASCII)", "ASCII": "Esri raster (Grid/ASCII)",
    "CSV": "CSV/Excel", "Excel": "CSV/Excel", "JSON": "JSON",
    "GeoTIFF (incl. COG)": "GeoTIFF", "File Geodatabase": "Geodatabase", "Esri ASCII grid": "Esri raster (Grid/ASCII)", "Esri ASCII": "Esri raster (Grid/ASCII)"}


def format_facets(s):
    """Formats cell -> filter groups (GeoTIFF (COG) counts as GeoTIFF; rare formats -> Other)."""
    out = []
    for f in re.split(r",|;|/(?![^()]*\))", s or ""):
        f = re.sub(r"^(Raster|Vector):\s*", "", f.strip())
        f = re.sub(r"\s*\(COG\)$", "", f)
        f = re.sub(r"^NetCDF-?4$", "NetCDF", f)
        if not f:
            continue
        g = FORMAT_GROUPS.get(f, "Other")
        if g not in out:
            out.append(g)
    return out


UPDATE_ORDER = ["Near-real-time", "Daily", "Sub-monthly", "Monthly", "Quarterly",
                "Annual", "Irregular", "Static"]


def update_facets(s):
    """Update frequency cell -> filter values: drop bracketed detail, keep controlled terms
    (e.g. 'Daily (L3), Near-real-time (3-hourly, L4)' -> Daily, Near-real-time)."""
    base = re.sub(r"\([^)]*\)", "", s or "")
    out = []
    for part in base.split(","):
        part = part.strip().lower()
        for term in UPDATE_ORDER:
            if part == term.lower() and term not in out:
                out.append(term)
    return out


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    # Themes sheet: name, code, overview, general usage
    themes = []
    for r in wb["Themes"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            themes.append(dict(name=txt(r[0]), code=txt(r[1]), overview=txt(r[2]), usage=txt(r[3])))

    # Cardinal Points sheet
    cps = []
    for r in wb["Carinal 4"].iter_rows(min_row=2, values_only=True):
        if r[0] and str(r[0]).startswith("CP"):
            cps.append(dict(code=txt(r[0]), name=txt(r[1]), covers=txt(r[2]), url=txt(r[3])))

    # Datasets
    ws = wb["Datasets"]
    H = [txt(c.value) for c in ws[1]]
    col = {h: i for i, h in enumerate(H)}
    datasets = []
    for row in ws.iter_rows(min_row=2):
        v = [c.value for c in row]
        if not v[0]:
            continue
        g = lambda h: txt(v[col[h]]) if h in col else ""
        access = []
        for k in range(1, 6):
            c = row[col[f"Access {k}"]]
            if txt(c.value):
                access.append(dict(label=txt(c.value), url=link_of(c)))
        datasets.append(dict(
            id=g("ID"), theme=g("Theme(s)"), name=g("Dataset"),
            datasetUrl=link_of(row[col["Dataset (URL)"]]),
            description=g("Description"), provider=g("Provider") or g("Main Provider"), africa=g("Africa priority"),
            area=area(g("Coverage (spatial)")),
            covSpatial=g("Coverage (spatial)"), covTemporal=g("Coverage (temporal)"),
            update=g("Update frequency"), upd=update_facets(g("Update frequency")), resolution=g("Spatial resolution"), formats=g("Formats"), fmt=format_facets(g("Formats")),
            access=access, licence=g("Licence"), licenceName=g("Licence Name"),
            notes=g("Notes"), priority=g("Priority"), stars=g("Priority").count("★"),
            useCases=[]))
    dmap = {d["id"]: d for d in datasets}

    # Use cases
    ws = wb["Use_Cases"]
    H = [txt(c.value) for c in ws[1]]
    col = {}
    for i, h in enumerate(H):
        col.setdefault(h, i)  # first occurrence (sheet has two 'Date' columns)
    usecases = []
    for row in ws.iter_rows(min_row=2):
        v = [c.value for c in row]
        if not v[0]:
            continue
        g = lambda h: txt(v[col[h]]) if h in col else ""
        rel_t, rel = tier(g("Relevance to AfDB"))
        cpm = CPTAG.search(rel)
        cp = re.findall(r"CP\d", cpm.group(1)) if cpm else []
        rel = CPTAG.sub("", rel, count=1).strip()
        rep_t, rep = tier(g("Replicability (data&code availability)"))
        dif_t, dif = tier(g("Methodological difficulty"))
        sectors = [s.strip() for s in g("Sector").split("/") if s.strip()]
        codes = [c for c in re.split(r"[,;\s]+", g("Countries covered (ISO3)")) if c]
        cat = g("Geospatial Data  — catalogued (Dataset IDs)")
        linked = [i for i in dict.fromkeys(DSID.findall(cat)) if i in dmap]
        dc = []
        for part in [p.strip() for p in re.split(r"[;\n]+", g("Data & code")) if p.strip()]:
            m = URL.search(part)
            dc.append(dict(label=part[:m.start()].rstrip(" :") if m else part, url=m.group(0) if m else ""))
        uc = dict(
            id=g("Use-case ID"), code=g("Use-case ID").replace("UC-", ""), title=g("Title"),
            sectors=sectors, year=g("Date")[:4], type=g("Type"),
            types=[t.strip() for t in re.split(r"\s+/\s+", g("Type")) if t.strip()],
            summary=g("Summary"), problem=g("Policy problem"),
            countries=[country(c) for c in codes], countryCodes=codes,
            datasets=linked, nonCatalogued=g("Geospatial Data — non-catalogued"),
            auxiliary=g("Auxiliary data"), method=g("Method (short)") or g("Method (short, for report)") or g("Method"), outputs=g("Main outputs"),
            findings=split_top(g("Key findings")),
            dataCode=dc, reference=g("Reference and link"),
            referenceUrl=link_of(row[col["Reference and link"]]),
            relevance=rel, relevanceTier=rel_t, cp=cp,
            replicability=rep, replicabilityTier=rep_t,
            difficulty=dif, difficultyTier=dif_t)
        usecases.append(uc)
        for did in linked:
            dmap[did]["useCases"].append(uc["id"])
    import datetime
    meta = dict(built=datetime.date.today().isoformat(), reportName=REPORT_NAME, hasReport=bool(REPORT_PDF))
    return dict(themes=themes, cps=cps, datasets=datasets, usecases=usecases, meta=meta)


TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Open Geospatial Data Glossary</title>
<link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --ink:#16241d; --ink-soft:#4a5a51; --paper:#f2f4ee; --card:#ffffff; --line:#d8ddd2;
  --green:#1b5e3a; --green-deep:#123f27; --green-tint:#e3ede6; --gold:#b8722e; --sand:#f5e6d5;
  --hero:#123f27; --radius:10px;
}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font-family:'IBM Plex Sans',system-ui,sans-serif;font-size:15px;line-height:1.55}
h1,h2,h3{font-family:'Source Serif 4',Georgia,serif;font-weight:600;margin:0;color:var(--ink)}
a{color:var(--green);text-decoration:none} a:hover{text-decoration:underline}
header.topbar{position:sticky;top:0;z-index:20;background:var(--card);border-bottom:1px solid var(--line)}
.topbar-row{max-width:1180px;margin:0 auto;padding:12px 20px;display:flex;align-items:center;gap:24px;flex-wrap:wrap}
.brand{color:var(--green-deep);font-family:'Source Serif 4',serif;cursor:pointer;line-height:1.15}
.brand .mark{display:block;font-size:18px;font-weight:700;letter-spacing:.02em}
.brand .full{display:block;font-size:10.5px;color:var(--ink-soft);font-family:'IBM Plex Sans'}
nav.tabs{display:flex;gap:4px}
nav.tabs button{background:transparent;border:0;color:var(--ink-soft);font:500 14px 'IBM Plex Sans';padding:8px 14px;border-radius:999px;cursor:pointer}
nav.tabs button:hover{background:var(--green-tint);color:var(--ink)}
nav.tabs button.active{background:var(--green);color:#fff}
main{max-width:1180px;margin:0 auto;padding:28px 20px 60px}
.pagehead{margin-bottom:16px} .pagehead p{color:var(--ink-soft);margin:6px 0 0;max-width:none}
.topicband{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);margin-bottom:18px;padding:12px 16px}
.textlist{display:flex;flex-wrap:wrap;align-items:center;font-size:13.5px;line-height:1.9}
.textitem{color:var(--ink-soft);cursor:pointer;padding:2px 4px;white-space:nowrap}
.textitem:hover{color:var(--ink)} .textitem.active{color:var(--green);font-weight:600}
.sep{color:var(--ink-soft);opacity:.5;margin:0 6px}
.layout{display:grid;grid-template-columns:230px 1fr;gap:26px;align-items:start}
@media(max-width:780px){.layout{grid-template-columns:1fr}}
.filterpanel{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:16px;position:sticky;top:76px;max-height:calc(100vh - 96px);overflow:auto}
.fg{margin-bottom:16px} .fg:last-child{margin-bottom:0}
.fg h4{font:600 12px 'IBM Plex Sans';color:var(--ink-soft);margin:0 0 8px;display:flex;justify-content:space-between;cursor:pointer}
.fg.collapsed .gb{display:none} .fg h4 .chev{transition:transform .15s} .fg.collapsed h4 .chev{transform:rotate(-90deg)}
.checklist{display:flex;flex-direction:column;gap:7px}
.checklist label{display:flex;align-items:center;gap:8px;font-size:13.5px;cursor:pointer}
.checklist input{accent-color:var(--green);width:15px;height:15px;margin:0}
.cnt{color:var(--ink-soft);font-size:12px}
.checklist label.zero{opacity:.45}
.groupsearch{width:100%;padding:6px 8px;margin-bottom:8px;border:1px solid var(--line);border-radius:6px;background:var(--paper);font:12.5px inherit}
.clearbtn{margin-top:10px;font-size:12.5px;color:var(--ink-soft);background:none;border:0;cursor:pointer;text-decoration:underline;padding:0}
.toolbar{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:12px}
.searchbox{position:relative;flex:1;min-width:240px}
.searchbox input{width:100%;padding:10px 14px 10px 36px;border-radius:8px;border:1px solid var(--line);background:var(--card);font:14px 'IBM Plex Sans';color:var(--ink)}
.searchbox svg{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--ink-soft)}
.resultcount{font-size:13px;color:var(--ink-soft)}
.sortctrl{font-size:12.5px;color:var(--ink-soft)} .sortctrl button{background:none;border:0;color:var(--ink-soft);cursor:pointer;font-size:12.5px;padding:2px 6px}
.sortctrl button.active{color:var(--green);font-weight:600;text-decoration:underline}
.sortctrl select{font:inherit;font-size:12.5px;color:var(--ink);background:var(--card,#fff);border:1px solid var(--line,#ddd);border-radius:6px;padding:3px 6px;margin-left:4px;cursor:pointer}
.cardgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:15px 17px;cursor:pointer;display:flex;flex-direction:column;gap:6px}
.card:hover{border-color:var(--green);box-shadow:0 4px 14px rgba(27,94,58,.08)}
.kicker{font-size:11.5px;color:#a8916f;font-weight:500;display:flex;justify-content:space-between;gap:8px}
.card h3{font-size:16px}
.card p.desc{color:var(--ink-soft);font-size:13.3px;margin:0}
.badgerow{display:flex;flex-wrap:wrap;gap:6px}
.badge{font-size:11.5px;padding:3px 9px;border-radius:999px;font-weight:500;white-space:nowrap}
.b-africa{background:#f7e3c4;color:#8a4f0f}
.b-theme{background:#f2f0e9;color:#7a6638}
.b-area{background:transparent;color:#3a5a48;border:1px solid #b9cbbf} .b-global{background:transparent;color:#5f6862;border:1px solid #cfd5cd}
.b-lic{background:#e6ebef;color:#4e6273}
.b-country{background:transparent;color:#5f6862;border:1px solid #cfd5cd}
.tb{display:inline-block;padding:1px 7px;border-radius:4px;font-weight:600;font-size:11px;color:#fff;margin-left:3px}
.t-High,.t-Easy{background:#486d58} .t-Medium,.t-Moderate{background:#8fa79a} .t-Low,.t-Difficult{background:#b39463}
.stars{color:#9f8753;letter-spacing:1px;font-family:'Segoe UI Symbol',sans-serif}
.cptag{font-size:11px;font-weight:700;color:var(--gold);border:1px solid var(--gold);border-radius:4px;padding:0 5px}
.detailcard{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:26px 28px}
.backlink{display:inline-flex;align-items:center;gap:6px;font-size:13px;margin-bottom:14px;color:var(--ink-soft)}
.detailcard h2{font-size:23px;margin:6px 0 4px}
.detailcard .sub{color:var(--ink-soft);font-size:12px;letter-spacing:.06em;text-transform:uppercase;margin-top:8px}
.detailcard p.lead{margin:18px 0 0;font-size:15px;font-style:italic;color:var(--ink)}
.titlerow{display:flex;justify-content:space-between;align-items:baseline;gap:24px}
.dslink{font-size:13px;font-weight:600;color:var(--green);white-space:nowrap}
.note{color:#8c8c8c;font-size:13px;margin-top:8px}
.attrgrid{display:grid;grid-template-columns:190px minmax(0,1fr);row-gap:12px;column-gap:18px;margin-top:18px;border-top:1px solid var(--line);padding-top:18px}
.attrgrid dt{color:var(--ink-soft);font-size:13px} .attrgrid dd{margin:0;font-size:14px;min-width:0;overflow-wrap:anywhere}
@media(max-width:640px){.attrgrid{grid-template-columns:1fr}.attrgrid dt{margin-top:6px}}
.sect{margin-top:20px;border-top:1px solid var(--line);padding-top:16px}
.sect h4{font:600 13px 'IBM Plex Sans';color:var(--ink-soft);margin:0 0 10px}
.linklist{display:flex;flex-direction:column;gap:7px}
.linkrow{display:flex;gap:10px;align-items:baseline;cursor:pointer;color:var(--green);font-size:14px}
.linkrow[hidden]{display:none}
.showall{margin-top:8px;background:none;border:0;padding:0;color:var(--green);font:500 13px 'IBM Plex Sans';cursor:pointer;text-decoration:underline}
.linkrow:hover{text-decoration:underline} .linkrow .code{font-weight:600;font-size:12.5px;min-width:64px}
.chiplink{display:inline-block;border:1px solid var(--line);border-radius:999px;padding:4px 12px;font-size:13px;margin:0 6px 6px 0;background:var(--paper)}
ul.findings{margin:4px 0 0;padding-left:18px} ul.findings li{margin-bottom:4px}
.emptystate{padding:40px 0;text-align:center;color:var(--ink-soft)}
.hero{background:var(--hero);color:#fff;padding:64px 20px 48px;width:100vw;margin-left:calc(50% - 50vw);margin-top:-28px;position:relative;overflow:hidden}
.hero::before{content:"";position:absolute;inset:0;opacity:.35;pointer-events:none;background-image:repeating-radial-gradient(circle at 82% 15%,transparent 0,transparent 22px,rgba(255,255,255,.05) 23px,rgba(255,255,255,.05) 24px),repeating-radial-gradient(circle at 8% 90%,transparent 0,transparent 30px,rgba(255,255,255,.04) 31px,rgba(255,255,255,.04) 32px)}
.hero-inner{position:relative;max-width:700px;margin:0 auto;text-align:center}
.hero h1{font-size:31px;color:#fff;line-height:1.3} .hero h1 .accent{color:#e0a467}
.hero p.tag{color:rgba(255,255,255,.75);font-size:14.5px;margin:12px 0 26px}
.herosearch{position:relative;max-width:560px;margin:0 auto}
.herosearch input{width:100%;padding:14px 20px 14px 46px;border-radius:8px;border:0;font:15px 'IBM Plex Sans'}
.herosearch svg{position:absolute;left:17px;top:50%;transform:translateY(-50%);color:var(--ink-soft)}
.herostats{display:flex;justify-content:center;gap:28px;flex-wrap:wrap;margin-top:26px}
.herostat{font-size:13px;color:rgba(255,255,255,.78)} .herostat b{color:#fff;font-size:16px}
.homesection{margin:38px 0 0} .homesection h2{font-size:19px;margin-bottom:4px}
.sectionsub{color:var(--ink-soft);font-size:13.5px;margin-bottom:14px}
.carousel-wrap{position:relative}
.carousel{display:flex;gap:14px;overflow-x:auto;scroll-snap-type:x mandatory;padding-bottom:4px;scrollbar-width:none}
.carousel::-webkit-scrollbar{display:none}
.carousel-btn{position:absolute;top:50%;transform:translateY(-50%);width:34px;height:34px;border-radius:999px;background:var(--card);border:1px solid var(--line);cursor:pointer;display:flex;align-items:center;justify-content:center;box-shadow:0 2px 6px rgba(0,0,0,.08);z-index:2;color:var(--ink)}
.carousel-btn:hover{border-color:var(--green);color:var(--green)}
.carousel-btn.prev{left:-16px}.carousel-btn.next{right:-16px}
@media(max-width:700px){.carousel-btn{display:none}}
.browse2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:780px){.browse2{grid-template-columns:1fr}}
.browsebox{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:18px 20px}
.browsebox h2{font-size:18px;margin-bottom:12px}
.infobtn{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;border-radius:50%;border:1px solid var(--ink-soft);color:var(--ink-soft);font:600 9px 'IBM Plex Sans';line-height:1;cursor:pointer;vertical-align:4px;margin-left:6px;background:none;padding:0}
.infobtn:hover,.infobtn.on{border-color:var(--green);color:var(--green)}
.infowrap{position:relative;display:inline-block}
.infobtn.small{width:10px;height:10px;font-size:7px;vertical-align:1px;margin-left:4px;border-width:1px}
.pop.fpop{position:fixed;width:300px;z-index:60}
.pop p{margin:0 0 6px}.pop p:last-child{margin:0}
.pop b{font-weight:600}
dt .infobtn.small{vertical-align:2px}
.pop{position:absolute;left:-10px;top:30px;width:330px;background:var(--card);border:1px solid var(--line);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.12);padding:12px 14px;font:400 13px/1.5 'IBM Plex Sans';color:var(--ink);z-index:30}
.pop::before{content:"";position:absolute;top:-6px;left:16px;width:10px;height:10px;background:var(--card);border-left:1px solid var(--line);border-top:1px solid var(--line);transform:rotate(45deg)}
.fdiv{border:0;border-top:1px solid var(--line);margin:4px 0 16px}
.kicker{align-items:flex-start} .kicker>span:first-child{display:flex;flex-wrap:wrap;align-items:center;gap:4px 8px}
.kicker .badge{font-weight:500;font-size:11px;padding:2px 8px}
.cprow{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:-6px 0 18px}
@media(max-width:780px){.cprow{grid-template-columns:1fr 1fr}}
.cpbtn{background:#fdfbf8;border:1px solid #ece3d6;border-radius:8px;padding:8px 12px;cursor:pointer;font-size:12.5px;color:var(--ink-soft);line-height:1.35;text-align:left}
.cpbtn b{display:block;color:#c49a70;font-size:12px;font-weight:600}
.cpbtn:hover{border-color:#d9bf9e} .cpbtn.active{background:#f8efe3;border-color:#d9a874;color:var(--ink)} .cpbtn.active b{color:var(--gold)}
.carousel .card{min-width:280px;max-width:280px;scroll-snap-align:start}
.cpgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px}
.cpcard{background:var(--card);border:1px solid var(--line);border-top:3px solid var(--gold);border-radius:var(--radius);padding:15px 17px;cursor:pointer}
.cpcard:hover{border-color:var(--green)} .cpcard .n{font:600 26px 'Source Serif 4';color:var(--green)}
.cpcard .code{font-size:12px;font-weight:700;color:var(--gold)} .cpcard .nm{font-size:13.5px;margin-top:2px}
.chipcloud{display:flex;flex-wrap:wrap;gap:8px}
.chipcloud .badge{cursor:pointer;font-size:12.5px;padding:5px 12px}
.langtoggle{display:flex;gap:2px;margin-left:auto}
.langtoggle button{background:transparent;border:1px solid var(--line);color:var(--ink-soft);font:600 11.5px 'IBM Plex Sans';padding:4px 9px;cursor:pointer}
.langtoggle button:first-child{border-radius:6px 0 0 6px}.langtoggle button:last-child{border-radius:0 6px 6px 0}
.langtoggle button.active{background:var(--green);border-color:var(--green);color:#fff}
.cardbottom{margin-top:auto;padding-top:4px}
.cardfoot{padding-top:2px;font-size:12px;color:var(--ink-soft)}
.langnote{font-size:12.5px;color:var(--ink-soft);margin-top:6px;font-style:italic}
footer{text-align:center;color:var(--ink-soft);font-size:12.5px;padding:24px 20px 40px}
footer a{color:var(--ink-soft);text-decoration:underline}
.topright{display:flex;align-items:center;gap:12px;margin-left:auto}
.topright .langtoggle{margin-left:0}
.topsearch{position:relative}
.topsearch[hidden]{display:none}
.topsearch input{width:210px;padding:7px 12px 7px 32px;border-radius:999px;border:1px solid var(--line);background:var(--bg);font:13px 'IBM Plex Sans';color:var(--ink)}
.topsearch input:focus{outline:2px solid var(--green);outline-offset:0;background:var(--card)}
.topsearch svg{position:absolute;left:11px;top:50%;transform:translateY(-50%);width:14px;height:14px;color:var(--ink-soft)}
@media(max-width:760px){.topsearch input{width:140px}}
.searchband{background:#e7ede4;border:1px solid var(--line);border-radius:var(--radius);padding:30px 20px 26px;margin:6px 0 26px;text-align:center}
.searchband h1{font:600 26px 'Source Serif 4',Georgia,serif;margin:0 0 16px;color:var(--ink)}
.pagesearch{position:relative;max-width:560px;margin:0 auto;text-align:left}
.pagesearch input{width:100%;padding:11px 16px 11px 42px;border-radius:8px;border:1px solid var(--line);background:var(--card);font:15px 'IBM Plex Sans';color:var(--ink)}
.pagesearch svg{position:absolute;left:15px;top:50%;transform:translateY(-50%);color:var(--ink-soft)}
.reportbtn{font:500 13px 'IBM Plex Sans';color:var(--green);border:1px solid var(--line);border-radius:999px;padding:5px 12px;white-space:nowrap}
.reportbtn:hover{border-color:var(--green);text-decoration:none}
details.fold{margin-top:12px}
details.fold>summary{cursor:pointer;list-style:none;display:grid;grid-template-columns:190px minmax(0,1fr);column-gap:18px;align-items:baseline}
details.fold>summary .fl{font-size:13px;color:var(--ink-soft);display:flex;align-items:center;gap:5px}
details.fold>summary .fh{font-size:13px;color:#9aa39c}
details.fold[open]>summary .fh{visibility:hidden}
@media(max-width:640px){details.fold>summary{grid-template-columns:1fr}details.fold>summary .fh{display:none}}
details.fold>summary::-webkit-details-marker{display:none}
details.fold>summary .chev{transform:rotate(-90deg);transition:transform .15s}
details.fold[open]>summary .chev{transform:none}
details.fold>summary:hover .fl{color:var(--green)}
details.fold .attrgrid{border-top:0;padding-top:0;margin-top:12px}
details.fold .attrgrid dt:empty{min-height:0}
.attrgrid.flat{border-top:0;padding-top:0;margin-top:14px}
.viewtoggle{display:flex;gap:0}
.viewtoggle button{background:var(--card);border:1px solid var(--line);color:var(--ink-soft);font:500 12px 'IBM Plex Sans';padding:3px 10px;cursor:pointer}
.viewtoggle button:first-child{border-radius:6px 0 0 6px}.viewtoggle button:last-child{border-radius:0 6px 6px 0;border-left:0}
.viewtoggle button.active{background:var(--green);border-color:var(--green);color:#fff}
.tablewrap{overflow-x:auto;border-top:1px solid var(--line)}
table.list{width:100%;border-collapse:collapse;font-size:13.5px}
table.list th{text-align:left;font:600 12px 'IBM Plex Sans';color:var(--ink-soft);padding:10px 12px;border-bottom:1px solid var(--line);white-space:nowrap;cursor:pointer;user-select:none}
table.list th:hover{color:var(--ink)} table.list th .ar{font-size:10px;margin-left:3px}
table.list td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top}
table.list tr:last-child td{border-bottom:0}
table.list tbody tr{cursor:pointer} table.list tbody tr:hover{background:rgba(27,94,58,.05)}
table.list td.nm{font-weight:600;color:var(--green-deep);min-width:200px}
table.list td.mut{color:var(--ink-soft);white-space:nowrap}
.cphead{font:600 12px 'IBM Plex Sans';color:var(--ink-soft);margin:0 0 8px}
</style>
</head>
<body>
<header class="topbar"><div class="topbar-row">
  <div class="brand" onclick="go('')"><span class="mark">OGDG</span><span class="full" id="brandfull"></span></div>
  <nav class="tabs" id="tabs"></nav>
  <div class="topright"><div class="topsearch" id="topsearch"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg><input id="topq" type="search"></div><a class="reportbtn" id="reportbtn"></a>
  <div class="langtoggle" id="langtoggle"><button data-lang="en">EN</button><button data-lang="fr">FR</button></div></div>
</div></header>
<main id="app"></main>
<footer id="foot"></footer>
<script type="application/pdf;base64" id="reportpdf">__REPORT__</script>
<script>
const DATA = __DATA__;
const SECTOR_COLORS = {"Agriculture":"#6b8a3a","Biodiversity":"#3f7d6a","Climate":"#c0662a","Disaster and hazards":"#a35c4c",
 "Economic Development":"#86693a","Energy":"#a8862f","Food security":"#9a6a32","Forestry":"#3f6a46","Health":"#9c5068",
 "Infrastructure":"#5d6c7c","Fisheries and ocean":"#2f7889","Population":"#775c8b","Urban and settlement":"#66679a","Water resources":"#3a6ca3"};
const AREAS = ["Africa","Global"];
const LICENCES = ["Free / open","Attribution","Non-commercial"];
const UPDATES = ["Near-real-time","Daily","Sub-monthly","Monthly","Quarterly","Annual","Irregular","Static"];
const TIERS = {relevanceTier:["High","Medium","Low"], replicabilityTier:["High","Medium","Low"], difficultyTier:["Easy","Moderate","Difficult"]};
const FORMATS = (()=>{const c={};DATA.datasets.forEach(d=>d.fmt.forEach(x=>c[x]=(c[x]||0)+1));return Object.keys(c).sort((a,b)=>(a==='Other')-(b==='Other')||c[b]-c[a]||a.localeCompare(b));})();
const dmap = Object.fromEntries(DATA.datasets.map(d=>[d.id,d]));
const umap = Object.fromEntries(DATA.usecases.map(u=>[u.id,u]));
const esc = s => String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const linkify = s => esc(s).replace(/https?:\/\/[^\s|;,)]+[^\s|;,.)]/g, m=>`<a href="${m}" target="_blank" rel="noopener">${m}</a>`);
const cut = (s,n) => { s=s||''; return s.length>n ? s.slice(0,n).replace(/\s+\S*$/,'')+'…' : s; };
const SVG_SEARCH = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>';
const CHEV = '<svg class="chev" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>';

/* ---------- interface text (EN / FR) ----------
   Only the interface is translated. Catalogue content (descriptions, summaries, methods...)
   stays in English. Category labels (themes, sectors, CPs, tiers) are translated for display;
   filter values stay in English. */
const UI = {
 en:{datasets:"Datasets",usecases:"Use cases",brandFull:"Open Geospatial Data Glossary · AfDB DI4A",
  heroA:"Open geospatial data",heroB:"for African development policy",heroTag:"Find a dataset, or see how it has been used in documented policy work.",
  searchAll:"Search datasets, use cases and countries…",searchShort:"Search…",nDatasets:"datasets",nThemes:"themes",nUsecases:"use cases",nSectors:"sectors",
  byTheme:"Browse datasets by theme",bySector:"Browse use cases by sector",
  highPrio:"High-priority datasets",highPrioSub:"The team's indicative pick of datasets to explore first.",prioTip:"How priority is set",infoTip:"What the grades mean",
  info_relevance:"<p><b>High:</b> direct support for a Cardinal Point through an actionable output (baseline, risk assessment, targeting, siting, early warning or monitoring) in Africa, or in a global study whose results cover African countries.</p><p><b>Medium:</b> indirect support (validation or accuracy assessment, identifying the most useful data, methodological reference or general evidence), or an application outside Africa that is transferable to African countries.</p><p><b>Low:</b> no clear link to a Cardinal Point, or not transferable to Africa.</p><p>Cardinal Point tag: the Cardinal Point(s) the case links to, primary first and at most two. Low cases carry no tag.</p>",
  info_replicability:"<p><b>High:</b> research code is released (e.g. GitHub, Zenodo or a software package) and the key inputs are open or can be obtained again.</p><p><b>Medium:</b> inputs are open but the code is not released, or a key dataset is restricted or available only on request.</p><p><b>Low:</b> a key input is proprietary or collected in the field and cannot readily be obtained again.</p>",
  info_difficulty:"<p><b>Easy:</b> overlay, zonal statistics, thresholds, simple indices or digitising, with simple checks against a reference (e.g. correlation, RMSE), at the level of QGIS or a spreadsheet.</p><p><b>Moderate:</b> standard statistical or remote-sensing workflows: image processing, regression, drought or water-balance indices, travel-time and multi-criteria analysis, classical machine learning (e.g. random forest) and formal accuracy assessment.</p><p><b>Difficult:</b> models built or calibrated by the study itself: deep learning, Bayesian models, process simulation, data assimilation, metaheuristic optimisation or specialised sensor processing (e.g. SAR time series, lidar).</p><p>Judged on the method only, independent of data access.</p>",
  prioInfo:"Priority reflects how often a dataset appears in documented use cases and across sectors, its relevance to Africa, cloud or API access, and licence openness. It suggests where to start; it is not a quality rating.",
  dsLead:"Individual geospatial datasets — products from the same source are listed as separate entries, grouped by theme.",
  coverage:"Coverage area",priority:"Priority",format:"Format",licence:"Licence",clear:"Clear filters",search:"Search…",noMatch:"No matches.",
  dsSearch:"Search datasets (name, provider, description)…",of:"of",sort:"Sort",az:"A–Z",
  noDs:"No datasets match these filters.",dsNotFound:"Dataset not found.",allDs:"← All datasets",note:"Note",
  provider:"Provider",dsPage:"Dataset page",spatial:"Spatial coverage",temporal:"Temporal coverage",update:"Update frequency",
  resolution:"Spatial resolution",formats:"Formats",access:"Data access",notDoc:"Not documented.",noneYet:"None documented yet.",ucCount:"Use cases",
  ucLead:"Documented applications of the catalogued datasets to policy questions, browsable by sector.",
  relevance:"AfDB relevance",replicability:"Replicability",difficulty:"Methodological difficulty",type:"Type",year:"Year",countries:"Countries covered",
  ucSearch:"Search use cases (title, country, method)…",noUc:"No use cases match these filters.",ucNotFound:"Use case not found.",allUc:"← All use cases",
  problem:"Policy problem",method:"Method",outputs:"Main outputs",findings:"Key findings",catData:"Catalogued data",otherGeo:"Other geospatial data",
  aux:"Auxiliary data",dataCode:"Data & code",relevanceRow:"Relevance to AfDB",reference:"Reference",
  newest:"Newest",byId:"ID",dsPageBtn:"Dataset page",africaPrio:"Africa priority",afrFocus:"Africa",afrHigh:"High African use",oldest:"Oldest",
  results:"Results for",openIn:"open in",none:"None.",footer:"Generated from the OGDG master workbook · Data Innovation for Africa (DI4A), African Development Bank",
  contentNote:"",details:"Details",otherData:"Other data",cpTitle:"AfDB’s Four Cardinal Points",
  highRel:"High-relevance use cases",highRelSub:"Rated High for AfDB relevance and replicability; simplest methods first.",
  viewCards:"Cards",viewList:"List",name:"Name",theme:"Theme",titleCol:"Title",sector:"Sector",
  report:"Report",reportFull:"Download the report (PDF)",lastUpdated:"Last updated",showAll:"Show all"},
 fr:{datasets:"Jeux de données",usecases:"Cas d'usage",brandFull:"Glossaire des données géospatiales ouvertes · BAD DI4A",
  heroA:"Données géospatiales ouvertes",heroB:"pour les politiques de développement en Afrique",heroTag:"Trouvez un jeu de données, ou voyez comment il a été utilisé dans des travaux de politique publique documentés.",
  searchAll:"Rechercher jeux de données, cas d'usage et pays…",searchShort:"Rechercher…",nDatasets:"jeux de données",nThemes:"thèmes",nUsecases:"cas d'usage",nSectors:"secteurs",
  byTheme:"Parcourir les jeux de données par thème",bySector:"Parcourir les cas d'usage par secteur",
  highPrio:"Jeux de données prioritaires",highPrioSub:"Sélection indicative de l'équipe : les jeux de données à explorer en premier.",prioTip:"Comment la priorité est établie",infoTip:"Signification des niveaux",
  info_relevance:"<p><b>Élevée :</b> soutien direct à un point cardinal grâce à un résultat directement utilisable (situation de référence, évaluation des risques, ciblage, choix de sites, alerte précoce ou suivi) en Afrique, ou dans une étude mondiale dont les résultats couvrent des pays africains.</p><p><b>Moyenne :</b> soutien indirect (validation ou évaluation de la précision, identification des données les plus utiles, référence méthodologique ou éléments de preuve généraux), ou application hors d'Afrique transposable aux pays africains.</p><p><b>Faible :</b> aucun lien clair avec un point cardinal, ou non transposable à l'Afrique.</p><p>Étiquette de point cardinal : le ou les points cardinaux auxquels le cas se rattache, le principal en premier et deux au plus. Les cas de pertinence faible n'ont pas d'étiquette.</p>",
  info_replicability:"<p><b>Élevée :</b> le code de recherche est publié (p. ex. GitHub, Zenodo ou un logiciel) et les données clés sont ouvertes ou peuvent être obtenues à nouveau.</p><p><b>Moyenne :</b> les données sont ouvertes mais le code n'est pas publié, ou un jeu de données clé est restreint ou disponible uniquement sur demande.</p><p><b>Faible :</b> une donnée clé est propriétaire ou collectée sur le terrain et ne peut pas être facilement obtenue à nouveau.</p>",
  info_difficulty:"<p><b>Facile :</b> superposition, statistiques zonales, seuils, indices simples ou numérisation, avec des contrôles simples par rapport à une référence (p. ex. corrélation, RMSE), au niveau de QGIS ou d'un tableur.</p><p><b>Modérée :</b> traitements statistiques ou de télédétection courants : traitement d'images, régression, indices de sécheresse ou de bilan hydrique, analyse des temps de trajet et multicritère, apprentissage automatique classique (p. ex. forêt aléatoire) et évaluation formelle de la précision.</p><p><b>Difficile :</b> modèles construits ou calibrés par l'étude elle-même : apprentissage profond, modèles bayésiens, simulation de processus, assimilation de données, optimisation métaheuristique ou traitement spécialisé de capteurs (p. ex. séries temporelles SAR, lidar).</p><p>Évalué sur la méthode uniquement, indépendamment de l'accès aux données.</p>",
  prioInfo:"La priorité tient compte de la fréquence d'utilisation d'un jeu de données dans les cas d'usage documentés et de la diversité des secteurs concernés, de sa pertinence pour l'Afrique, de l'accès via le cloud ou une API, et de l'ouverture de sa licence. Elle indique par où commencer ; ce n'est pas une note de qualité.",
  dsLead:"Des jeux de données géospatiaux distincts — les produits d'une même source sont présentés comme des fiches séparées, classées par thème.",
  coverage:"Couverture",priority:"Priorité",format:"Format",licence:"Licence",clear:"Effacer les filtres",search:"Rechercher…",noMatch:"Aucun résultat.",
  dsSearch:"Rechercher (nom, fournisseur, description)…",of:"sur",sort:"Trier",az:"A–Z",
  noDs:"Aucun jeu de données ne correspond à ces filtres.",dsNotFound:"Jeu de données introuvable.",allDs:"← Tous les jeux de données",note:"Remarque",
  provider:"Fournisseur",dsPage:"Page du jeu de données",spatial:"Couverture spatiale",temporal:"Couverture temporelle",update:"Fréquence de mise à jour",
  resolution:"Résolution spatiale",formats:"Formats",access:"Accès aux données",notDoc:"Non documenté.",noneYet:"Aucun cas documenté pour l'instant.",ucCount:"Cas d'usage",
  ucLead:"Applications documentées des jeux de données du catalogue à des questions de politique publique, consultables par secteur.",
  relevance:"Pertinence pour la BAD",replicability:"Reproductibilité",difficulty:"Difficulté méthodologique",type:"Type",year:"Année",countries:"Pays couverts",
  ucSearch:"Rechercher (titre, pays, méthode)…",noUc:"Aucun cas d'usage ne correspond à ces filtres.",ucNotFound:"Cas d'usage introuvable.",allUc:"← Tous les cas d'usage",
  problem:"Problème de politique publique",method:"Méthode",outputs:"Principaux résultats",findings:"Constats clés",catData:"Données du catalogue",otherGeo:"Autres données géospatiales",
  aux:"Données auxiliaires",dataCode:"Données et code",relevanceRow:"Pertinence pour la BAD",reference:"Référence",
  newest:"Plus récents",byId:"ID",dsPageBtn:"Page du jeu de données",africaPrio:"Priorité Afrique",afrFocus:"Afrique",afrHigh:"Forte utilisation en Afrique",oldest:"Plus anciens",
  results:"Résultats pour",openIn:"ouvrir dans",none:"Aucun.",footer:"Généré à partir du classeur principal OGDG · Data Innovation for Africa (DI4A), Banque africaine de développement",
  contentNote:"Le contenu des fiches est en anglais.",details:"Détails",otherData:"Autres données",cpTitle:"Les quatre points cardinaux de la BAD",
  highRel:"Cas d'usage à forte pertinence",highRelSub:"Pertinence pour la BAD et reproductibilité élevées ; méthodes les plus simples en premier.",
  viewCards:"Cartes",viewList:"Liste",name:"Nom",theme:"Thème",titleCol:"Titre",sector:"Secteur",
  report:"Rapport",reportFull:"Télécharger le rapport (PDF, en anglais)",lastUpdated:"Dernière mise à jour",showAll:"Tout afficher"}
};
const LABEL_FR = {
 // themes
 "Elevation & Terrain":"Altitude et relief","Satellite Imagery":"Imagerie satellitaire","Land Cover & Use":"Occupation et utilisation des sols",
 "Climate & Weather":"Climat et météo","Population & Settlement":"Population et établissements humains","Hydrology & Water":"Hydrologie et eau",
 "Administrative Boundaries":"Limites administratives","Soil & Agriculture":"Sols et agriculture","Infrastructure & Built Environment":"Infrastructures et cadre bâti",
 "Biodiversity & Ecosystems":"Biodiversité et écosystèmes","Fire & Thermal":"Feux et anomalies thermiques","Atmosphere & Air Quality":"Atmosphère et qualité de l'air",
 "Nighttime Lights":"Lumières nocturnes","Ocean & Coastal":"Océans et zones côtières","Energy & Renewables":"Énergie et renouvelables","Connectivity & Digital":"Connectivité et numérique",
 // sectors
 "Agriculture":"Agriculture","Biodiversity":"Biodiversité","Climate":"Climat","Disaster and hazards":"Catastrophes et aléas","Economic Development":"Développement économique",
 "Energy":"Énergie","Food security":"Sécurité alimentaire","Forestry":"Foresterie","Health":"Santé","Infrastructure":"Infrastructures",
 "Fisheries and ocean":"Pêche et océans","Population":"Population","Urban and settlement":"Urbanisation et établissements humains","Water resources":"Ressources en eau",
 // Cardinal Points (working translation, not the Bank's official French wording)
 "CP1":"Améliorer l'accès aux capitaux : mobiliser les ressources financières de l'Afrique",
 "CP2":"Réformer et consolider les systèmes financiers, les institutions et les talents du continent",
 "CP3":"Tirer parti de la transformation démographique pour le développement économique",
 "CP4":"Bâtir des infrastructures résilientes au climat et accroître la valeur ajoutée des ressources naturelles",
 // tiers, coverage, licence, types
 "High":"Élevée","Medium":"Moyenne","Low":"Faible","Easy":"Facile","Moderate":"Modérée","Difficult":"Difficile",
 "Africa":"Afrique","Global":"Mondiale","Africa-first":"Conçu pour l'Afrique","High African use":"Forte utilisation en Afrique","Other":"Autre",
 "Near-real-time":"Quasi temps réel","Daily":"Quotidienne","Sub-monthly":"Infra-mensuelle","Monthly":"Mensuelle","Quarterly":"Trimestrielle","Annual":"Annuelle","Irregular":"Irrégulière","Static":"Statique","Free / open":"Libre / ouverte","Attribution":"Attribution","Non-commercial":"Non commerciale",
 "Dataset":"Jeu de données","Early warning":"Alerte précoce","Operational programme":"Programme opérationnel",
 "Research (application)":"Recherche (application)","Research (validation/accuracy)":"Recherche (validation/précision)"
};

const AFR = {'Africa-first':'afrFocus','Africa-focused':'afrFocus','Africa':'afrFocus','High African use':'afrHigh'};
let state = {ds:{f:{},q:'',sort:'az',open:{}}, uc:{f:{},q:'',sort:'new',open:{}}, gq:{}, info:false, lang:'en', view:{ds:'cards',uc:'cards'}, tsort:{ds:null,uc:null}};
try { const l = localStorage.getItem('ogdg-lang'); if (l==='fr'||l==='en') state.lang = l; } catch(e) {}
const T = k => (UI[state.lang][k] ?? UI.en[k]);
const L = v => state.lang==='fr' ? (LABEL_FR[v] || v) : v;                 // category label
const cpName = c => state.lang==='fr' ? (LABEL_FR[c.code] || c.name) : c.name;
const typeLabel = u => u.types.map(L).join(' / ');
const FRESH = {ds:()=>({f:{},q:'',sort:'az',open:{}}), uc:()=>({f:{},q:'',sort:'new',open:{}})};
let keepFilters=false;  // set by home-page theme/sector chips, which open a list with one filter applied
function resetList(sec){ state[sec]=FRESH[sec](); Object.keys(state.gq).forEach(k=>{ if(k.startsWith(sec)) delete state.gq[k]; }); }
function go(h){
  const target='#/'+h;
  if(location.hash===target){ if(!keepFilters && (h==='datasets'||h==='usecases')) resetList(h==='datasets'?'ds':'uc'); keepFilters=false; router(); }
  else location.hash=target;
}

/* ---------- small helpers ---------- */
function sectorBadge(s){
  const c = SECTOR_COLORS[s] || '#6b736d';
  return `<span class="badge" style="background:${c}1f;color:${c}">${esc(L(s))}</span>`;
}
function cpBadges(u){ return u.cp.map(c=>`<span class="cptag">${c}</span>`).join(' '); }
function areaBadge(a){ return `<span class="badge ${a==='Global'?'b-global':'b-area'}">${esc(L(a))}</span>`; }
function stars(n){ return n ? `<span class="stars" title="${esc(T('priority'))}">${'★'.repeat(n)}</span>` : ''; }
function countryLabel(u){ const c=u.countries; return c.length>2 ? c.slice(0,2).join(', ')+` +${c.length-2}` : c.join(', '); }
function primary(u){ return u.sectors[0] || ''; }
function toggle(sec,key,val){ const a=state[sec].f[key]=state[sec].f[key]||[]; const i=a.indexOf(val); i<0?a.push(val):a.splice(i,1); }
function check(sec,key,val,label,count){
  const on=(state[sec].f[key]||[]).includes(val);
  return `<label class="${count===0&&!on?'zero':''}"><input type="checkbox" data-sec="${sec}" data-key="${key}" data-val="${esc(val)}" ${on?'checked':''}> ${label}${count!==undefined?` <span class="cnt">(${count})</span>`:''}</label>`;
}
function ibtn(k){ return `<span class="infowrap"><button class="infobtn small fib ${state.finfo===k?'on':''}" data-info="${k}" title="${esc(T(k==='prio'?'prioTip':'infoTip'))}">i</button></span>`; }
function group(sec,key,title,body){
  const open = state[sec].open[key] === true;
  return `<div class="fg ${open?'':'collapsed'}"><h4 data-gt="${sec}:${key}"><span>${title}</span>${CHEV}</h4><div class="gb">${body}</div></div>`;
}
function checklist(sec,key,items,searchable){
  const q=(state.gq[sec+key]||'').toLowerCase();
  const list = q ? items.filter(it=>it.label.toLowerCase().includes(q)) : items;
  return (searchable?`<input class="groupsearch" id="gq-${sec+key}" data-gq="${sec+key}" placeholder="${T('search')}" value="${esc(state.gq[sec+key]||'')}">`:'') +
    `<div class="checklist" style="${searchable?'max-height:210px;overflow:auto':''}">${list.map(it=>check(sec,key,it.value,esc(it.label),it.count)).join('')||`<span class="cnt">${T('noMatch')}</span>`}</div>`;
}
function band(sec,key,items){
  const sel=state[sec].f[key]||[];
  return `<div class="topicband"><div class="textlist">${items.map((t,i)=>(i?'<span class="sep">|</span>':'')+
    `<span class="textitem ${sel.includes(t)?'active':''}" data-band="${sec}:${key}" data-val="${esc(t)}">${esc(L(t))}</span>`).join('')}</div></div>`;
}
function searchbar(sec,ph){ return `<div class="searchbox">${SVG_SEARCH}<input id="q-${sec}" placeholder="${ph}" value="${esc(state[sec].q)}"></div>`; }
let reportUrl=null;
function openReport(){
  if(!reportUrl){ const b64=document.getElementById('reportpdf').textContent.trim(); const bin=atob(b64); const u8=new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++) u8[i]=bin.charCodeAt(i); reportUrl=URL.createObjectURL(new Blob([u8],{type:'application/pdf'})); }
  const a=document.createElement('a'); a.href=reportUrl; a.download=DATA.meta.reportName; document.body.appendChild(a); a.click(); a.remove();
}
function viewToggle(sec){ const v=state.view[sec];
  return `<span class="viewtoggle"><button data-view="${sec}:cards" class="${v==='cards'?'active':''}">${T('viewCards')}</button><button data-view="${sec}:list" class="${v==='list'?'active':''}">${T('viewList')}</button></span>`; }
function th(sec,key,label){ const t=state.tsort[sec]; const ar=t&&t.key===key?`<span class="ar">${t.dir>0?'▲':'▼'}</span>`:'';
  return `<th data-tsort="${sec}:${key}">${label}${ar}</th>`; }
function tsorted(sec,list,getters){ const t=state.tsort[sec]; if(!t) return list; const g=getters[t.key];
  return [...list].sort((a,b)=>{ const x=g(a),y=g(b); return t.dir*(typeof x==='number'?x-y:String(x).localeCompare(String(y))); }); }
function dsTable(list){
  const G={name:d=>d.name,theme:d=>L(d.theme),provider:d=>d.provider,area:d=>L(d.area),licence:d=>L(d.licence),stars:d=>d.stars};
  const rows=tsorted('ds',list,G);
  return `<div class="tablewrap"><table class="list"><thead><tr>${th('ds','name',T('name'))}${th('ds','theme',T('theme'))}${th('ds','provider',T('provider'))}${th('ds','area',T('coverage'))}${th('ds','licence',T('licence'))}${th('ds','stars',T('priority'))}</tr></thead>
  <tbody>${rows.map(d=>`<tr data-go="datasets/${d.id}"><td class="nm">${esc(d.name)}</td><td>${esc(L(d.theme))}</td><td>${esc(d.provider)}</td><td class="mut">${esc(L(d.area))}</td><td class="mut">${esc(L(d.licence))}</td><td class="mut">${stars(d.stars)}</td></tr>`).join('')}</tbody></table></div>`;
}
function ucTable(list){
  const G={title:u=>u.title,sector:u=>L(primary(u)),year:u=>u.year,countries:u=>u.countries.join(', ')};
  const rows=tsorted('uc',list,G);
  return `<div class="tablewrap"><table class="list"><thead><tr>${th('uc','title',T('titleCol'))}${th('uc','sector',T('sector'))}${th('uc','year',T('year'))}${th('uc','countries',T('countries'))}</tr></thead>
  <tbody>${rows.map(u=>`<tr data-go="usecases/${u.id}"><td class="nm">${esc(u.title)}</td><td>${esc(u.sectors.map(L).join(', '))}</td><td class="mut">${esc(u.year)}</td><td title="${esc(u.countries.join(', '))}">${esc(countryLabel(u))}</td></tr>`).join('')}</tbody></table></div>`;
}
function fold(label,hint,inner,extra){ return `<details class="fold"${extra||''}><summary><span class="fl">${CHEV}${label}</span><span class="fh">${hint}</span></summary>${inner}</details>`; }
const DIFF_ORDER={Easy:0,Moderate:1,Difficult:2};
function carousel(id,cards){ return `<div class="carousel-wrap">
      <button class="carousel-btn prev" data-car="${id}:-1" aria-label="Previous">‹</button>
      <div class="carousel" id="${id}">${cards}</div>
      <button class="carousel-btn next" data-car="${id}:1" aria-label="Next">›</button></div>`; }
const contentNote = () => T('contentNote') ? `<div class="langnote">${T('contentNote')}</div>` : '';

/* ---------- home ---------- */
function renderHome(){
  const top = DATA.datasets.filter(d=>d.stars===5);
  const hiRel = DATA.usecases.filter(u=>u.relevanceTier==='High'&&u.replicabilityTier==='High')
    .sort((a,b)=>(DIFF_ORDER[a.difficultyTier]??9)-(DIFF_ORDER[b.difficultyTier]??9) || (b.year||'').localeCompare(a.year||'') || a.id.localeCompare(b.id));
  const sectorCount = s => DATA.usecases.filter(u=>primary(u)===s).length;
  const sectors = Object.keys(SECTOR_COLORS).filter(s=>sectorCount(s));
  return `
  <div class="hero"><div class="hero-inner">
    <h1>${T('heroA')}<br><span class="accent">${T('heroB')}</span></h1>
    <p class="tag">${T('heroTag')}</p>
    <div class="herosearch">${SVG_SEARCH}<input id="homeq" placeholder="${T('searchAll')}"></div>
    <div class="herostats">
      <div class="herostat"><b>${DATA.datasets.length}</b> ${T('nDatasets')}</div>
      <div class="herostat"><b>${DATA.themes.length}</b> ${T('nThemes')}</div>
      <div class="herostat"><b>${DATA.usecases.length}</b> ${T('nUsecases')}</div>
      <div class="herostat"><b>${sectors.length}</b> ${T('nSectors')}</div>
    </div></div></div>
  <div class="homesection"><h2>${T('byTheme')}</h2><div class="sectionsub">&nbsp;</div>
    <div class="chipcloud">${DATA.themes.map(t=>`<span class="badge b-theme" data-theme="${esc(t.name)}">${esc(L(t.name))} · ${DATA.datasets.filter(d=>d.theme===t.name).length}</span>`).join('')}</div></div>
  <div class="homesection"><h2>${T('bySector')}</h2><div class="sectionsub">&nbsp;</div>
    <div class="chipcloud">${sectors.map(s=>`<span data-sector="${esc(s)}">${sectorBadge(s)}</span>`).join('')}</div></div>
  <div class="homesection"><h2>${T('highPrio')}<span class="infowrap"><button class="infobtn ${state.info?'on':''}" id="infobtn" title="${esc(T('prioTip'))}">i</button>${state.info?`<div class="pop" id="pop">${T('prioInfo')}</div>`:''}</span></h2>
    <div class="sectionsub">${T('highPrioSub')}</div>
    ${carousel('carDs',top.map(dsCard).join(''))}</div>
  <div class="homesection"><h2>${T('highRel')}</h2>
    <div class="sectionsub">${T('highRelSub')}</div>
    ${carousel('carUc',hiRel.map(ucCard).join(''))}</div>`;
}

/* ---------- datasets ---------- */
function dsCard(d){
  return `<div class="card" data-go="datasets/${d.id}">
    <div class="kicker"><span>${d.id}</span>${stars(d.stars)}</div>
    <h3>${esc(d.name)}</h3>
    <p class="desc">${esc(cut(d.description,150))}</p>
    <div class="badgerow cardbottom"><span class="badge b-theme">${esc(L(d.theme))}</span>${areaBadge(d.area)}<span class="badge b-lic">${esc(L(d.licence))}</span></div>
    <div class="cardfoot">${T('ucCount')}: ${d.useCases.length}</div>
  </div>`;
}
function dsFiltered(skip){
  const f={...state.ds.f}, q=state.ds.q.toLowerCase();
  if(skip) delete f[skip];
  return DATA.datasets.filter(d=>
    (!f.theme?.length || f.theme.includes(d.theme)) && (!f.area?.length || f.area.includes(d.area)) &&
    (!f.format?.length || d.fmt.some(x=>f.format.includes(x))) &&
    (!f.licence?.length || f.licence.includes(d.licence)) && (!f.stars?.length || f.stars.includes(String(d.stars))) &&
    (!f.update?.length || d.upd.some(x=>f.update.includes(x))) &&
    (!q || (d.id+' '+d.name+' '+d.provider+' '+d.description+' '+d.theme+' '+L(d.theme)).toLowerCase().includes(q)));
}
function renderDatasets(){
  const f=state.ds.f, list=dsFiltered();
  const sorted=[...list].sort((a,b)=> state.ds.sort==='prio' ? b.stars-a.stars || a.name.localeCompare(b.name)
    : state.ds.sort==='uc' ? b.useCases.length-a.useCases.length || a.name.localeCompare(b.name) : a.name.localeCompare(b.name));
  const any = Object.values(f).some(a=>a.length) || state.ds.q;
  return `<div class="pagehead"><h1>${T('datasets')}</h1><p>${T('dsLead')}</p>${contentNote()}</div>
  ${band('ds','theme',DATA.themes.map(t=>t.name))}
  <div class="layout"><div class="filterpanel">
    ${group('ds','stars',`${T('priority')}${ibtn('prio')}`,checklist('ds','stars',[5,4,3,2,1].map(n=>({value:String(n),label:'★'.repeat(n),count:dsFiltered('stars').filter(d=>d.stars===n).length}))))}
    <hr class="fdiv">
    ${group('ds','area',T('coverage'),checklist('ds','area',AREAS.map(a=>({value:a,label:L(a),count:dsFiltered('area').filter(d=>d.area===a).length}))))}
    ${group('ds','format',T('format'),checklist('ds','format',FORMATS.map(x=>({value:x,label:x,count:dsFiltered('format').filter(d=>d.fmt.includes(x)).length})),true))}
    ${group('ds','licence',T('licence'),checklist('ds','licence',LICENCES.map(l=>({value:l,label:L(l),count:dsFiltered('licence').filter(d=>d.licence===l).length}))))}
    ${group('ds','update',T('update'),checklist('ds','update',UPDATES.filter(x=>DATA.datasets.some(d=>d.upd.includes(x))).map(x=>({value:x,label:L(x),count:dsFiltered('update').filter(d=>d.upd.includes(x)).length}))))}
    ${any?`<button class="clearbtn" data-clear="ds">${T('clear')}</button>`:''}
  </div><div>
    <div class="toolbar">${searchbar('ds',T('dsSearch'))}</div>
    <div class="toolbar"><span class="resultcount">${list.length} ${T('of')} ${DATA.datasets.length} ${T('nDatasets')}</span>
      <span class="sortctrl" style="margin-left:auto">${T('sort')}: <button data-sort="ds:az" class="${state.ds.sort==='az'&&!(state.view.ds==='list'&&state.tsort.ds)?'active':''}">${T('az')}</button><button data-sort="ds:prio" class="${state.ds.sort==='prio'&&!(state.view.ds==='list'&&state.tsort.ds)?'active':''}">${T('priority')}</button><button data-sort="ds:uc" class="${state.ds.sort==='uc'&&!(state.view.ds==='list'&&state.tsort.ds)?'active':''}">${T('ucCount')}</button></span>${viewToggle('ds')}</div>
    ${!sorted.length?`<div class="emptystate">${T('noDs')}</div>`: state.view.ds==='list' ? dsTable(sorted) : `<div class="cardgrid">${sorted.map(dsCard).join('')}</div>`}
  </div></div>`;
}
function renderDataset(id){
  const d=dmap[id]; if(!d) return `<div class="emptystate">${T('dsNotFound')}</div>`;
  const ucs=d.useCases.map(i=>umap[i]).filter(Boolean);
  const row=(k,v)=>v?`<dt>${T(k)}</dt><dd>${v}</dd>`:'';
  return `<a href="#/datasets" class="backlink">${T('allDs')}</a>
  <div class="detailcard">
    <div class="kicker"><span><span class="badge b-theme">${esc(L(d.theme))}</span></span>${stars(d.stars)}</div>
    <div class="titlerow"><h2>${esc(d.name)}</h2>${d.datasetUrl?`<a class="dslink" href="${esc(d.datasetUrl)}" target="_blank" rel="noopener">${T('dsPageBtn')} ↗</a>`:''}</div>
    <p style="margin:12px 0 0">${esc(d.description)}</p>
    ${d.notes?`<div class="note">${T('note')}: ${esc(d.notes)}</div>`:''}
    <dl class="attrgrid">
      ${row('provider', esc(d.provider))}
      ${row('spatial', esc(L(d.covSpatial)))}${row('temporal', esc(d.covTemporal))}
      ${row('update', esc(d.update))}${row('resolution', esc(d.resolution))}
      ${row('formats', esc(d.formats))}${row('licence', esc(L(d.licence) + (d.licenceName?` (${d.licenceName})`:'')))}
    </dl>
    <div class="sect"><h4>${T('access')}</h4>${d.access.map(a=>a.url?`<a class="chiplink" href="${esc(a.url)}" target="_blank" rel="noopener">${esc(a.label)} ↗</a>`:`<span class="chiplink">${esc(a.label)}</span>`).join('')||`<span class="cnt">${T('notDoc')}</span>`}</div>
    <div class="sect"><h4>${T('ucCount')} (${ucs.length})</h4>${ucs.length?`<div class="linklist">${ucs.map((u,i)=>`<div class="linkrow${i>=5?' more':''}" data-go="usecases/${u.id}"${i>=5?' hidden':''}><span class="code">${u.code}</span><span>${esc(u.title)}</span></div>`).join('')}</div>${ucs.length>5?`<button class="showall" id="showall">${T('showAll')}</button>`:''}`:`<span class="cnt">${T('noneYet')}</span>`}</div>
  </div>`;
}

/* ---------- use cases ---------- */
function ucCard(u){
  return `<div class="card" data-go="usecases/${u.id}">
    <div class="kicker"><span>${u.code} · ${u.year} · ${esc(typeLabel(u))}</span></div>
    <h3>${esc(u.title)}</h3>
    <p class="desc">${esc(cut(u.summary,150))}</p>
    <div class="badgerow cardbottom">${u.sectors.map(sectorBadge).join('')}<span class="badge b-country" title="${esc(u.countries.join(', '))}">${esc(countryLabel(u))}</span></div>
  </div>`;
}
function ucFiltered(skip){
  const f={...state.uc.f}, q=state.uc.q.toLowerCase();
  if(skip) delete f[skip];
  return DATA.usecases.filter(u=>
    (!f.sector?.length || f.sector.includes(primary(u))) &&
    (!f.relevanceTier?.length || f.relevanceTier.includes(u.relevanceTier)) &&
    (!f.replicabilityTier?.length || f.replicabilityTier.includes(u.replicabilityTier)) &&
    (!f.difficultyTier?.length || f.difficultyTier.includes(u.difficultyTier)) &&
    (!f.cp?.length || u.cp.some(c=>f.cp.includes(c))) &&
    (!f.type?.length || u.types.some(t=>f.type.includes(t))) &&
    (!f.year?.length || f.year.includes(u.year)) &&
    (!f.country?.length || u.countries.some(c=>f.country.includes(c))) &&
    (!q || (u.id+' '+u.title+' '+u.summary+' '+u.countries.join(' ')+' '+u.method+' '+u.sectors.join(' ')+' '+u.sectors.map(L).join(' ')).toLowerCase().includes(q)));
}
function renderUseCases(){
  const f=state.uc.f, list=ucFiltered();
  const dir = state.uc.sort==='old' ? -1 : 1;
  const sorted=[...list].sort((a,b)=> dir*(b.year||'').localeCompare(a.year||'') || a.id.localeCompare(b.id));
  const cnt=(fn)=>DATA.usecases.filter(fn).length;
  const cntK=(key,fn)=>ucFiltered(key).filter(fn).length;  // facet count: all other active filters applied
  const sectors=Object.keys(SECTOR_COLORS).filter(s=>cnt(u=>primary(u)===s));
  const types=[...new Set(DATA.usecases.flatMap(u=>u.types))].sort();
  const years=[...new Set(DATA.usecases.map(u=>u.year))].sort().reverse();
  const countries=[...new Set(DATA.usecases.flatMap(u=>u.countries))].sort();
  const any = Object.values(f).some(a=>a.length) || state.uc.q;
  const tierGroup=(key,title)=>group('uc',key,title,checklist('uc',key,TIERS[key].map(t=>({value:t,label:L(t),count:cntK(key,u=>u[key]===t)}))));
  const cpsel=f.cp||[];
  return `<div class="pagehead"><h1>${T('usecases')}</h1><p>${T('ucLead')}</p>${contentNote()}</div>
  ${band('uc','sector',sectors)}
  <div class="cphead">${T('cpTitle')}</div>
  <div class="cprow">${DATA.cps.map(c=>`<div class="cpbtn ${cpsel.includes(c.code)?'active':''}" data-band="uc:cp" data-val="${c.code}" title="${esc(cpName(c))}"><b>${c.code}</b>${esc(cpName(c))}</div>`).join('')}</div>
  <div class="layout"><div class="filterpanel">
    ${tierGroup('relevanceTier',T('relevance')+ibtn('relevance'))}${tierGroup('replicabilityTier',T('replicability')+ibtn('replicability'))}${tierGroup('difficultyTier',T('difficulty')+ibtn('difficulty'))}
    <hr class="fdiv">
    ${group('uc','type',T('type'),checklist('uc','type',types.map(t=>({value:t,label:L(t),count:cntK('type',u=>u.types.includes(t))}))))}
    ${group('uc','year',T('year'),checklist('uc','year',years.map(y=>({value:y,label:y,count:cntK('year',u=>u.year===y)}))))}
    ${group('uc','country',T('countries'),checklist('uc','country',countries.map(c=>({value:c,label:c,count:cntK('country',u=>u.countries.includes(c))})),true))}
    ${any?`<button class="clearbtn" data-clear="uc">${T('clear')}</button>`:''}
  </div><div>
    <div class="toolbar">${searchbar('uc',T('ucSearch'))}</div>
    <div class="toolbar"><span class="resultcount">${list.length} ${T('of')} ${DATA.usecases.length} ${T('nUsecases')}</span>
      <span class="sortctrl" style="margin-left:auto">${T('sort')}: <select id="ucsort"><option value="new"${state.uc.sort!=='old'?' selected':''}>${T('newest')}</option><option value="old"${state.uc.sort==='old'?' selected':''}>${T('oldest')}</option></select></span>${viewToggle('uc')}</div>
    ${!sorted.length?`<div class="emptystate">${T('noUc')}</div>`: state.view.uc==='list' ? ucTable(sorted) : `<div class="cardgrid">${sorted.map(ucCard).join('')}</div>`}
  </div></div>`;
}
function renderUseCase(id){
  const u=umap[id]; if(!u) return `<div class="emptystate">${T('ucNotFound')}</div>`;
  const ds=u.datasets.map(i=>dmap[i]).filter(Boolean);
  const row=(k,v)=>v?`<dt>${T(k)}</dt><dd>${v}</dd>`:'';
  const irow=(k,ik,v)=>v?`<dt>${T(k)}${ibtn(ik)}</dt><dd>${v}</dd>`:'';
  const tierTxt=(t,txt)=> (t?`<span class="tb t-${t}" style="margin:0 6px 0 0">${esc(L(t))}</span>`:'') + esc(txt);
  return `<a href="#/usecases" class="backlink">${T('allUc')}</a>
  <div class="detailcard">
    <div class="kicker"><span class="badgerow">${u.sectors.map(sectorBadge).join('')}</span></div>
    <h2>${esc(u.title)}</h2>
    <div class="sub">${esc(u.year)} · ${esc(typeLabel(u))}${u.countries.length?` · ${esc(u.countries.join(', '))}`:''}</div>
    <p class="lead">${esc(u.summary)}</p>
    ${(u.problem||u.method||u.outputs||u.findings.length)?fold(T('details'),[u.problem&&T('problem'),u.method&&T('method'),u.outputs&&T('outputs'),u.findings.length&&T('findings')].filter(Boolean).join(' · '),`<dl class="attrgrid">
      ${row('problem',esc(u.problem))}
      ${row('method',esc(u.method))}${row('outputs',esc(u.outputs))}
      ${row('findings', u.findings.length?`<ul class="findings">${u.findings.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>`:'')}
    </dl>`):''}
    <dl class="attrgrid">
      ${row('catData', ds.map(d=>`<div class="linkrow" data-go="datasets/${d.id}"><span class="code">${d.id}</span><span>${esc(d.name)}</span></div>`).join(''))}
    </dl>
    ${(u.nonCatalogued||u.auxiliary)?fold(T('otherData'),[u.nonCatalogued&&T('otherGeo'),u.auxiliary&&T('aux')].filter(Boolean).join(' · '),`<dl class="attrgrid">
      ${row('otherGeo',linkify(u.nonCatalogued))}${row('aux',esc(u.auxiliary))}
    </dl>`):''}
    ${u.dataCode.length?`<dl class="attrgrid flat">
      ${row('dataCode', u.dataCode.map(x=>x.url?`<a class="chiplink" href="${esc(x.url)}" target="_blank" rel="noopener">${esc(x.label||'Link')} ↗</a>`:`<span>${esc(x.label)}</span>`).join(''))}
    </dl>`:''}

    <dl class="attrgrid">
      ${irow('relevanceRow','relevance', tierTxt(u.relevanceTier,'') + (u.cp.length?cpBadges(u).replace(/class="cptag"/g,'class="cptag" style="margin-right:6px"'):'') + esc(u.relevance))}
      ${irow('replicability','replicability', tierTxt(u.replicabilityTier,u.replicability))}
      ${irow('difficulty','difficulty', tierTxt(u.difficultyTier,u.difficulty))}
    </dl>
    ${u.reference?`<div style="border-top:1px solid var(--line);margin-top:18px;padding-top:2px">${fold(T('reference'),esc(cut(u.reference,90)),`<dl class="attrgrid">
      <dt></dt><dd>${/https?:\/\//.test(u.reference) || !u.referenceUrl ? u.reference.split(/;\s+/).map(linkify).join('<br>') : `<a href="${esc(u.referenceUrl)}" target="_blank" rel="noopener">${esc(u.reference)}</a>`}</dd>
    </dl>`)}</div>`:''}
  </div>`;
}

/* ---------- search across both ---------- */
function renderSearch(q){
  // A home-page search starts from a clean slate: earlier list filters are cleared.
  resetList('ds'); resetList('uc');
  state.ds.q=q; state.uc.q=q;
  const d=dsFiltered(), u=ucFiltered();
  return `<div class="searchband"><h1>${T('results')} “${esc(q)}”</h1><div class="pagesearch">${SVG_SEARCH}<input id="srchq" type="search" value="${esc(q)}" placeholder="${T('searchAll')}"></div></div>
   <div class="sect" style="border:0;padding:0"><h4>${T('datasets')} (${d.length}) · <a href="#/datasets" class="openin" data-open="datasets">${T('openIn')} ${T('datasets')} →</a></h4>${d.length?dsTable(d.slice(0,12)):`<div class="emptystate">${T('none')}</div>`}</div>
   <div class="sect"><h4>${T('usecases')} (${u.length}) · <a href="#/usecases" class="openin" data-open="usecases">${T('openIn')} ${T('usecases')} →</a></h4>${u.length?ucTable(u.slice(0,12)):`<div class="emptystate">${T('none')}</div>`}</div>`;
}

/* ---------- router + events ---------- */
let lastHash=null;
function router(){
  const [sec,id] = location.hash.replace(/^#\/?/,'').split('/');
  document.documentElement.lang = state.lang;
  document.getElementById('brandfull').textContent = T('brandFull');
  const hasRep=DATA.meta.hasReport;
  document.getElementById('foot').innerHTML = `${esc(T('footer'))}<br>${T('lastUpdated')}: ${esc(DATA.meta.built)}${hasRep?` · <a href="#" class="replink">${esc(T('reportFull'))}</a>`:''}`;
  const ts=document.getElementById('topsearch'); ts.hidden=(sec===''||sec==='search'); const tq=document.getElementById('topq'); tq.placeholder=T('searchShort'); tq.setAttribute('aria-label',T('searchAll'));
  const rb=document.getElementById('reportbtn'); rb.style.display=hasRep?'':'none'; rb.href='#'; rb.classList.add('replink'); rb.textContent=T('report')+' ↓'; rb.title=T('reportFull');
  document.querySelectorAll('.replink').forEach(a=>a.onclick=e=>{e.preventDefault();openReport();});
  document.getElementById('tabs').innerHTML = [['datasets',T('datasets')],['usecases',T('usecases')]]
    .map(([k,l])=>`<button class="${sec===k?'active':''}" onclick="go('${k}')">${l}</button>`).join('');
  document.querySelectorAll('#langtoggle button').forEach(b=>b.classList.toggle('active', b.dataset.lang===state.lang));
  const app=document.getElementById('app');
  const focus=document.activeElement&&document.activeElement.id, pos=focus&&document.activeElement.selectionStart;
  const navigated = location.hash!==lastHash, y=window.scrollY;
  // Opening a list page from anywhere else (tab, 'All …' link, browser back) starts with no filters.
  if(navigated && !id && (sec==='datasets'||sec==='usecases')){ if(!keepFilters) resetList(sec==='datasets'?'ds':'uc'); }
  keepFilters=false; lastHash=location.hash;
  const fp=document.querySelector('.filterpanel'), fpY=fp?fp.scrollTop:0;
  if(sec==='datasets') app.innerHTML = id ? renderDataset(decodeURIComponent(id)) : renderDatasets();
  else if(sec==='usecases') app.innerHTML = id ? renderUseCase(decodeURIComponent(id)) : renderUseCases();
  else if(sec==='search') app.innerHTML = renderSearch(decodeURIComponent(id||''));
  else app.innerHTML = renderHome();
  bind(); showFpop();
  if(focus){ const el=document.getElementById(focus); if(el){ el.focus({preventScroll:true}); try{el.setSelectionRange(pos,pos);}catch(e){} } }
  if(navigated) window.scrollTo(0,0);
  else { window.scrollTo(0,y); const fp2=document.querySelector('.filterpanel'); if(fp2) fp2.scrollTop=fpY; }
}
function keepScroll(fn){ const y=window.scrollY; fn(); router(); window.scrollTo(0,y); }
function showFpop(){
  document.querySelectorAll('#fpop').forEach(x=>x.remove());
  document.querySelectorAll('.fib').forEach(x=>x.classList.toggle('on',x.dataset.info===state.finfo));
  const b=state.finfo && [...document.querySelectorAll('.fib')].find(x=>x.dataset.info===state.finfo && x.offsetParent);
  if(!b) return;
  const r=b.getBoundingClientRect(), d=document.createElement('div');
  d.id='fpop'; d.className='pop fpop'; d.innerHTML=state.finfo==='prio'?esc(T('prioInfo')):T('info_'+state.finfo);
  d.style.left=Math.max(8,Math.min(r.left-10, innerWidth-296))+'px'; d.style.top=(r.bottom+8)+'px';
  document.body.appendChild(d);
}
function bind(){
  document.querySelectorAll('[data-go]').forEach(el=>el.onclick=()=>go(el.dataset.go));
  document.querySelectorAll('input[data-key]').forEach(el=>el.onchange=()=>{toggle(el.dataset.sec,el.dataset.key,el.dataset.val);router();});
  document.querySelectorAll('[data-band]').forEach(el=>el.onclick=()=>{const [s,k]=el.dataset.band.split(':');toggle(s,k,el.dataset.val);router();});
  document.querySelectorAll('[data-gt]').forEach(el=>el.onclick=()=>{const [s,k]=el.dataset.gt.split(':');const o=state[s].open;o[k]=!(el.parentElement.classList.contains('collapsed')?false:true);router();});
  document.querySelectorAll('[data-sort]').forEach(el=>el.onclick=()=>{const [s,v]=el.dataset.sort.split(':');state[s].sort=v;state.tsort[s]=null;router();});
  const us=document.getElementById('ucsort'); if(us) us.onchange=()=>{state.uc.sort=us.value;state.tsort.uc=null;router();};
  document.querySelectorAll('[data-clear]').forEach(el=>el.onclick=()=>{const s=el.dataset.clear;state[s].f={};state[s].q='';router();});
  document.querySelectorAll('.groupsearch').forEach(el=>el.oninput=()=>{state.gq[el.dataset.gq]=el.value;router();});
  ['ds','uc'].forEach(s=>{const el=document.getElementById('q-'+s); if(el) el.oninput=()=>{state[s].q=el.value;router();};});
  ['homeq','srchq','topq'].forEach(id=>{ const el=document.getElementById(id); if(el) el.onkeydown=e=>{ if(e.key==='Enter'&&el.value.trim()){ const v=el.value.trim(); if(id==='topq') el.value=''; go('search/'+encodeURIComponent(v)); } }; });
  document.querySelectorAll('.fib').forEach(el=>el.onclick=e=>{e.stopPropagation(); state.finfo=state.finfo===el.dataset.info?null:el.dataset.info; showFpop();});
  const ib=document.getElementById('infobtn'); if(ib) ib.onclick=e=>{e.stopPropagation(); keepScroll(()=>state.info=!state.info);};
  document.querySelectorAll('[data-car]').forEach(el=>el.onclick=()=>{const [id,d]=el.dataset.car.split(':');document.getElementById(id).scrollBy({left:300*Number(d),behavior:'smooth'});});
  const sa=document.getElementById('showall'); if(sa) sa.onclick=()=>{document.querySelectorAll('.linkrow.more').forEach(x=>x.hidden=false); sa.remove();};
  document.querySelectorAll('[data-open]').forEach(el=>el.onclick=e=>{e.preventDefault(); keepFilters=true; location.hash='#/'+el.dataset.open;});
  document.querySelectorAll('[data-view]').forEach(el=>el.onclick=()=>{const [s,v]=el.dataset.view.split(':');state.view[s]=v;router();});
  document.querySelectorAll('[data-tsort]').forEach(el=>el.onclick=()=>{const [s,k]=el.dataset.tsort.split(':');const t=state.tsort[s];state.tsort[s]=(t&&t.key===k)?{key:k,dir:-t.dir}:{key:k,dir:1};router();});
  document.querySelectorAll('[data-theme]').forEach(el=>el.onclick=()=>{resetList('ds');state.ds.f={theme:[el.dataset.theme]};keepFilters=true;go('datasets');});
  document.querySelectorAll('[data-sector]').forEach(el=>el.onclick=()=>{resetList('uc');state.uc.f={sector:[el.dataset.sector]};keepFilters=true;go('usecases');});
}
document.querySelectorAll('#langtoggle button').forEach(b=>b.onclick=()=>keepScroll(()=>{
  state.lang=b.dataset.lang; try{localStorage.setItem('ogdg-lang',state.lang);}catch(e){}
}));
window.addEventListener('hashchange',router);
document.addEventListener('click',e=>{ if(state.info && !e.target.closest('#pop') && !e.target.closest('#infobtn')) keepScroll(()=>state.info=false);
  if(state.finfo && !e.target.closest('#fpop') && !e.target.closest('.fib')) { state.finfo=null; showFpop(); } });
window.addEventListener('scroll',()=>{ if(state.finfo){ state.finfo=null; showFpop(); } },true);
router();
</script>
</body>
</html>
'''


def main():
    data = load(XLSX)
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    import base64, os
    rep = ""
    if REPORT_PDF:
        with open(REPORT_PDF, "rb") as f:
            rep = base64.b64encode(f.read()).decode("ascii")
    html = TEMPLATE.replace("__DATA__", payload).replace("__REPORT__", rep)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"saved {OUT}: {len(data['datasets'])} datasets, {len(data['usecases'])} use cases, "
          f"{len(data['themes'])} themes")


if __name__ == "__main__":
    main()
