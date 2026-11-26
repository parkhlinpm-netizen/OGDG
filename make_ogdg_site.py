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
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@500&family=Source+Serif+4:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
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

/* ===== survey-map redesign ===== */
:root{--night:var(--green-deep);--contour:#486d58}
header.topbar{background:var(--night);border-bottom:0}
.brand{display:flex;align-items:center;gap:11px;color:#fff}
.brand .logo{width:34px;height:34px;color:#e0a467;flex:none}
.brand .mark{color:#fff;font:700 19px 'Source Serif 4',Georgia,serif;letter-spacing:.04em}
.brand .full{color:rgba(255,255,255,.62)}
nav.tabs button{color:rgba(255,255,255,.75);border-radius:0;border-bottom:2px solid transparent;padding:8px 4px;margin:0 10px}
nav.tabs button:hover{background:transparent;color:#fff}
nav.tabs button.active{background:transparent;color:#fff;border-bottom-color:#e0a467}
.langtoggle button{border-color:rgba(255,255,255,.3);color:rgba(255,255,255,.75)}
.langtoggle button.active{background:#e0a467;border-color:#e0a467;color:var(--night)}
.reportbtn{color:#fff;border-color:rgba(255,255,255,.35)} .reportbtn:hover{border-color:#e0a467}
.topsearch input{background:rgba(255,255,255,.1);border-color:rgba(255,255,255,.25);color:#fff}
.topsearch input::placeholder{color:rgba(255,255,255,.55)} .topsearch svg{color:rgba(255,255,255,.6)}
.topsearch input:focus{background:rgba(255,255,255,.16);outline-color:#e0a467}

.sheet{position:relative;width:100vw;margin-left:calc(50% - 50vw);margin-top:-28px;background:#e7ede9;border-bottom:1px solid var(--line);overflow:hidden}
.sheet .topo{position:absolute;inset:0;width:100%;height:100%;color:var(--contour);opacity:.2;pointer-events:none}
.sheet-inner{position:relative;max-width:1180px;margin:0 auto;padding:64px 20px 56px;display:grid;grid-template-columns:minmax(0,1.55fr) minmax(260px,1fr);gap:56px;align-items:end}
@media(max-width:860px){.sheet-inner{grid-template-columns:1fr;gap:32px}}
.kick{font:500 11.5px 'IBM Plex Mono',monospace;letter-spacing:.14em;text-transform:uppercase;color:var(--gold);margin-bottom:18px}
.sheet h1{font:600 50px/1.1 'Source Serif 4',Georgia,serif;color:var(--night);letter-spacing:-.015em}
@media(max-width:640px){.sheet h1{font-size:36px}}
.sheet .lede{font-size:17px;line-height:1.55;color:var(--ink);margin:18px 0 28px;max-width:44em}
.sheet .lede b{font-weight:600;color:var(--green-deep)}
.sheet .tag2{color:var(--ink-soft);font-size:14.5px;margin:0 0 26px}
.sheet .herosearch{margin:0;max-width:540px}
.sheet .herosearch input{border:1px solid #c9ccbd;border-radius:2px;box-shadow:0 1px 0 rgba(0,0,0,.04);background:#fff}
.sheet .herosearch input:focus{outline:2px solid var(--green);outline-offset:0}
.stamp{font:500 11.5px 'IBM Plex Mono',monospace;color:var(--ink-soft);margin-top:14px;letter-spacing:.04em}
.legend{background:rgba(255,255,255,.82);border:1px solid var(--night);padding:20px 22px 16px;position:relative}
.legend::before,.legend::after{content:"";position:absolute;width:10px;height:10px;border:1px solid var(--night)}
.legend::before{top:-6px;left:-6px;border-right:0;border-bottom:0}
.legend::after{bottom:-6px;right:-6px;border-left:0;border-top:0}
.legend .lh{font:500 11px 'IBM Plex Mono',monospace;letter-spacing:.14em;text-transform:uppercase;color:var(--ink-soft);border-bottom:1px solid var(--line);padding-bottom:8px;margin-bottom:6px}
.legend .row{display:flex;align-items:baseline;justify-content:space-between;padding:9px 0;border-bottom:1px dashed #d5d8cb;cursor:pointer}
.legend .row:last-child{border-bottom:0}
.legend .row span{font-size:14px;color:var(--ink-soft)}
.legend .row b{font:600 28px/1 'Source Serif 4',Georgia,serif;color:var(--night);font-variant-numeric:lining-nums}
.legend .row:hover span{color:var(--green)}

.homesection h2{font:600 22px/1.2 'Source Serif 4',Georgia,serif;color:var(--night)}
.homesection{margin-top:52px}
.sechead{display:flex;align-items:baseline;justify-content:space-between;gap:16px;border-bottom:1px solid var(--night);padding-bottom:8px;margin-bottom:0}
.sechead .seemore{font:500 12px 'IBM Plex Mono',monospace;letter-spacing:.08em;text-transform:uppercase;color:var(--green);cursor:pointer}
.themegrid{display:grid;grid-template-columns:repeat(4,1fr);border-left:1px solid var(--line)}
@media(max-width:900px){.themegrid{grid-template-columns:repeat(2,1fr)}}
.ttile{border-right:1px solid var(--line);border-bottom:1px solid var(--line);padding:16px 18px 15px;cursor:pointer;background:transparent;transition:background .15s;display:flex;flex-direction:column;gap:6px;min-height:118px}
.ttile:hover{background:var(--night)} .ttile:hover *{color:#fff!important}
.ttile .tc{display:flex;justify-content:space-between;font:500 11.5px 'IBM Plex Mono',monospace;letter-spacing:.08em;color:var(--gold)}
.ttile .tn{font:600 17px/1.25 'Source Serif 4',Georgia,serif;color:var(--night)}
.ttile .tk{margin-top:auto;font-size:12.5px;color:var(--ink-soft)}
.ttile .bar{height:3px;background:var(--line);position:relative;margin-top:2px}
.ttile .bar i{position:absolute;left:0;top:0;bottom:0;background:var(--green)}
.ttile:hover .bar{background:rgba(255,255,255,.2)} .ttile:hover .bar i{background:#e0a467}
.secgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));border-left:1px solid var(--line)}
.secgrid .sc{border-right:1px solid var(--line);border-bottom:1px solid var(--line);padding:11px 16px;cursor:pointer;display:flex;justify-content:space-between;align-items:baseline;font-size:14px}
.secgrid .sc b{font:600 18px 'Source Serif 4',Georgia,serif;color:var(--night);font-variant-numeric:lining-nums}
.secgrid .sc:hover{background:var(--green-tint)}
.secgrid .dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:9px;vertical-align:1px}
.homesection .sectionsub{margin:10px 0 14px}

.sheet .lede br{display:none}@media(min-width:700px){.sheet .lede br{display:inline}}
.sheet .lede mark{white-space:nowrap;background:none;color:#9a6b3c;font-weight:600;padding:0}
/* ---- original stable layout; heading wording restored; label in report green ---- */
.sheet h1 .accent{color:#9a6b3c}
.sheet h1 .hd{display:none}
.sheet .kick{color:var(--sage)}
@media(min-width:861px){.sheet h1{font-size:44px;line-height:1.15}}
@media(max-width:860px){.sheet h1 .hm{display:none}.sheet h1 .hd{display:inline}}
/* ---- left column as set; dashboard back to full size and stretched to the same height as the text column ---- */
.sheet .sheet-inner{grid-template-columns:minmax(0,1.55fr) minmax(250px,1fr);gap:56px;align-items:stretch}
@media(max-width:860px){.sheet .sheet-inner{grid-template-columns:1fr;gap:32px}}
.sheet .lede{font-size:16px}
.sheet .legend{padding:18px 22px 14px;display:flex;flex-direction:column}
.sheet .legend .lh{padding-bottom:8px;margin-bottom:0}
.sheet .legend .row{padding:0;flex:1;align-items:center}
.sheet .legend .row span{font-size:14px}
.sheet .legend .row b{font-size:28px}
.sheet .stamp{text-align:left}
@media(max-width:860px){.sheet .legend .row{padding:9px 0;flex:none}}
@media(min-width:641px) and (max-width:860px){.sheet h1{font-size:36px}}
@media(max-width:640px){.sheet h1{font-size:28px}}

/* ---- palette: original site colours (deep green + gold) ---- */
:root{--accent:#e0a467;--contour:#cfe0d4}
.searchband{background:transparent;border:0;border-radius:0;padding:12px 20px 22px;margin:0 0 18px;border-bottom:1px solid var(--line)}
header.topbar{background:var(--green-deep);border-bottom:1px solid rgba(255,255,255,.08)}
.sheet{background:var(--green-deep);border-bottom:0}.sheet .topo{opacity:.16}
.sheet .kick{color:var(--accent)}
.sheet h1{color:#fff}.sheet h1 .accent{color:var(--accent)}
.sheet .lede{color:rgba(255,255,255,.88)}.sheet .stamp{color:rgba(255,255,255,.6)}
.sheet .herosearch input{border:0;border-radius:8px}
.legend{background:rgba(255,255,255,.06);border-color:rgba(255,255,255,.4)}
.legend::before,.legend::after{border-color:rgba(255,255,255,.4)}
.legend .lh{color:rgba(255,255,255,.65);border-bottom-color:rgba(255,255,255,.2)}
.legend .row{border-bottom-color:rgba(255,255,255,.16)}.legend .row span{color:rgba(255,255,255,.78)}.legend .row b{color:#fff}
.legend .row:hover span{color:var(--accent)}
.sechead{border-bottom-color:var(--green-deep)}
.homesection h2,.ttile .tn,.secgrid .sc b{color:var(--green-deep)}
.ttile .bar i{background:var(--green)}.ttile:hover{background:var(--green-deep)}.ttile:hover .bar i{background:var(--accent)}

/* ---- home: hero back to about 400 px, tiles keep "N datasets" + bar, brown code/number row removed ---- */
@media(min-width:861px){.sheet h1{font-size:42px}}
.ttile .tc{display:none}
.ttile{padding:17px 18px 15px;min-height:104px;gap:6px}
.ttile .tn{font-size:16.5px;line-height:1.25}
.ttile .tk{margin-top:auto;font-size:12.5px}
.ttile .bar{margin-top:6px}
.homesection{margin-top:46px}
</style>
</head>
<body>
<header class="topbar"><div class="topbar-row">
  <div class="brand" onclick="go('')"><svg class="logo" viewBox="0 0 40 40" aria-hidden="true"><g fill="none" stroke="currentColor" stroke-width="1.6"><path d="M20 4c8 0 15 6 15 15s-6 17-15 17S5 29 5 20 12 4 20 4z"/><path d="M20 10c5 0 9 4 9 9s-4 11-9 11-9-6-9-11 4-9 9-9z"/><path d="M20 16c2.5 0 4 1.8 4 4s-1.8 4.5-4 4.5-4-2.3-4-4.5 1.5-4 4-4z"/></g><circle cx="20" cy="20.3" r="1.6" fill="currentColor"/></svg><span><span class="mark">OGDG</span><span class="full" id="brandfull"></span></span></div>
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
const TOPO = '__TOPO_SVG__';
const SVG_SEARCH = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>';
const CHEV = '<svg class="chev" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>';

/* ---------- interface text (EN / FR) ----------
   Only the interface is translated. Catalogue content (descriptions, summaries, methods...)
   stays in English. Category labels (themes, sectors, CPs, tiers) are translated for display;
   filter values stay in English. */
const UI = {
 en:{datasets:"Datasets",usecases:"Use cases",brandFull:"Open Geospatial Data Glossary · AfDB DI4A",
  heroA:"Open geospatial data",kick:"Data Innovation for Africa · AfDB Statistics Department",title:"Open Geospatial Data Glossary",lede:"Open geospatial datasets and documented use cases, <br>curated for <mark>African development work</mark>.",legendH:"At a glance",howTo:"Search by keyword, or browse by theme and sector below.",dsOne:"dataset",upd:"Last updated",seeDs:"All datasets",seeUc:"All use cases",ofTotal:"of datasets",heroB:"for African development",heroTag:"Find a dataset, or see how it has been applied in documented use cases.",
  searchAll:"Search datasets, use cases and countries…",searchShort:"Search…",nDatasets:"datasets",nThemes:"themes",nUsecases:"use cases",nSectors:"sectors",
  byTheme:"Browse datasets by theme",bySector:"Browse use cases by sector",
  highPrio:"High-priority datasets",highPrioSub:"The team's indicative pick of datasets to explore first.",prioTip:"How priority is set",infoTip:"What the grades mean",
  info_relevance:"<p><b>High:</b> direct support for a Cardinal Point through an actionable output (baseline, risk assessment, targeting, siting, early warning or monitoring) in Africa, or in a global study whose results cover African countries.</p><p><b>Medium:</b> indirect support (validation or accuracy assessment, identifying the most useful data, methodological reference or general evidence), or an application outside Africa that is transferable to African countries.</p><p><b>Low:</b> no clear link to a Cardinal Point, or not transferable to Africa.</p>",
  info_replicability:"<p><b>High:</b> research code is released (e.g. GitHub, Zenodo or a software package) and the key inputs are open or can be obtained again.</p><p><b>Medium:</b> inputs are open but the code is not released, or a key dataset is restricted or available only on request.</p><p><b>Low:</b> a key input is proprietary or collected in the field and cannot readily be obtained again.</p>",
  info_difficulty:"<p><b>Easy:</b> overlay, zonal statistics, thresholds, simple indices or digitising, with simple checks against a reference (e.g. correlation, RMSE), at the level of QGIS or a spreadsheet.</p><p><b>Moderate:</b> standard statistical or remote-sensing workflows: image processing, regression, drought or water-balance indices, travel-time and multi-criteria analysis, classical machine learning (e.g. random forest) and formal accuracy assessment.</p><p><b>Difficult:</b> models built or calibrated by the study itself: deep learning, Bayesian models, process simulation, data assimilation, metaheuristic optimisation or specialised sensor processing (e.g. SAR time series, lidar).</p>",
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
  heroA:"Données géospatiales ouvertes",kick:"Data Innovation for Africa · Département des statistiques de la BAD",title:"Glossaire des données géospatiales ouvertes",lede:"Jeux de données géospatiales ouverts et cas d'usage documentés, <br>sélectionnés pour le <mark>développement africain</mark>.",legendH:"En un coup d'œil",howTo:"Recherchez par mot-clé, ou parcourez par thème et secteur ci-dessous.",dsOne:"jeu de données",upd:"Dernière mise à jour",seeDs:"Tous les jeux de données",seeUc:"Tous les cas d'usage",ofTotal:"des jeux de données",heroB:"pour le développement africain",heroTag:"Trouvez un jeu de données, ou voyez comment il a été appliqué dans des cas d'usage documentés.",
  searchAll:"Rechercher jeux de données, cas d'usage et pays…",searchShort:"Rechercher…",nDatasets:"jeux de données",nThemes:"thèmes",nUsecases:"cas d'usage",nSectors:"secteurs",
  byTheme:"Parcourir les jeux de données par thème",bySector:"Parcourir les cas d'usage par secteur",
  highPrio:"Jeux de données prioritaires",highPrioSub:"Sélection indicative de l'équipe : les jeux de données à explorer en premier.",prioTip:"Comment la priorité est établie",infoTip:"Signification des niveaux",
  info_relevance:"<p><b>Élevée :</b> soutien direct à un point cardinal grâce à un résultat directement utilisable (situation de référence, évaluation des risques, ciblage, choix de sites, alerte précoce ou suivi) en Afrique, ou dans une étude mondiale dont les résultats couvrent des pays africains.</p><p><b>Moyenne :</b> soutien indirect (validation ou évaluation de la précision, identification des données les plus utiles, référence méthodologique ou éléments de preuve généraux), ou application hors d'Afrique transposable aux pays africains.</p><p><b>Faible :</b> aucun lien clair avec un point cardinal, ou non transposable à l'Afrique.</p>",
  info_replicability:"<p><b>Élevée :</b> le code de recherche est publié (p. ex. GitHub, Zenodo ou un logiciel) et les données clés sont ouvertes ou peuvent être obtenues à nouveau.</p><p><b>Moyenne :</b> les données sont ouvertes mais le code n'est pas publié, ou un jeu de données clé est restreint ou disponible uniquement sur demande.</p><p><b>Faible :</b> une donnée clé est propriétaire ou collectée sur le terrain et ne peut pas être facilement obtenue à nouveau.</p>",
  info_difficulty:"<p><b>Facile :</b> superposition, statistiques zonales, seuils, indices simples ou numérisation, avec des contrôles simples par rapport à une référence (p. ex. corrélation, RMSE), au niveau de QGIS ou d'un tableur.</p><p><b>Modérée :</b> traitements statistiques ou de télédétection courants : traitement d'images, régression, indices de sécheresse ou de bilan hydrique, analyse des temps de trajet et multicritère, apprentissage automatique classique (p. ex. forêt aléatoire) et évaluation formelle de la précision.</p><p><b>Difficile :</b> modèles construits ou calibrés par l'étude elle-même : apprentissage profond, modèles bayésiens, simulation de processus, assimilation de données, optimisation métaheuristique ou traitement spécialisé de capteurs (p. ex. séries temporelles SAR, lidar).</p>",
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
  const nDs = DATA.datasets.length, maxT = Math.max(...DATA.themes.map(t=>DATA.datasets.filter(d=>d.theme===t.name).length));
  const lede = T('lede');
  const heroM = T('heroB').match(/^(\S+)\s+([\s\S]+)$/), heroFirst = (heroM && state.lang==='en')?heroM[1]:'', heroRest = (heroM && state.lang==='en')?heroM[2]:T('heroB');
  return `
  <section class="sheet">${TOPO}<div class="sheet-inner">
    <div>
      <div class="kick">${T('kick')}</div>
      <h1>${T('heroA')}<span class="hm"> ${heroFirst}</span><br><span class="accent"><span class="hd">${heroFirst} </span>${heroRest}</span></h1>
      <p class="lede">${T('heroTag')}</p>
      <div class="herosearch">${SVG_SEARCH}<input id="homeq" placeholder="${T('searchAll')}"></div>
      <div class="stamp">${T('upd')} · ${esc(DATA.meta.built)}</div>
    </div>
    <aside class="legend">
      <div class="lh">${T('legendH')}</div>
      <div class="row" onclick="go('datasets')"><span>${T('nDatasets')}</span><b>${nDs}</b></div>
      <div class="row" onclick="go('datasets')"><span>${T('nThemes')}</span><b>${DATA.themes.length}</b></div>
      <div class="row" onclick="go('usecases')"><span>${T('nUsecases')}</span><b>${DATA.usecases.length}</b></div>
      <div class="row" onclick="go('usecases')"><span>${T('nSectors')}</span><b>${sectors.length}</b></div>
    </aside>
  </div></section>
  <div class="homesection"><div class="sechead"><h2>${T('byTheme')}</h2><span class="seemore" onclick="go('datasets')">${T('seeDs')} →</span></div>
    <div class="themegrid">${DATA.themes.map(t=>{const n=DATA.datasets.filter(d=>d.theme===t.name).length;
      return `<div class="ttile" data-theme="${esc(t.name)}"><div class="tc"><span>${esc(t.code)}</span><span>${String(n).padStart(2,'0')}</span></div>
        <div class="tn">${esc(L(t.name))}</div><div class="tk">${n} ${n===1?T('dsOne'):T('nDatasets')}</div><div class="bar"><i style="width:${Math.round(n/maxT*100)}%"></i></div></div>`;}).join('')}</div></div>
  <div class="homesection"><div class="sechead"><h2>${T('bySector')}</h2><span class="seemore" onclick="go('usecases')">${T('seeUc')} →</span></div>
    <div class="secgrid">${sectors.map(sc=>`<div class="sc" data-sector="${esc(sc)}"><span><i class="dot" style="background:${SECTOR_COLORS[sc]||'#888'}"></i>${esc(L(sc))}</span><b>${sectorCount(sc)}</b></div>`).join('')}</div></div>
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


CONTOUR = '<svg class="topo" viewBox="0 0 1400 640" preserveAspectRatio="xMidYMid slice" aria-hidden="true"><g fill="none" stroke="currentColor"><polygon points="883.1,150.0 884.7,152.1 885.6,154.5 885.6,156.8 884.6,159.0 882.7,160.8 880.3,162.1 877.7,163.1 875.1,163.7 872.6,164.2 870.4,164.7 868.3,165.2 866.3,165.8 864.2,166.3 862.1,166.6 860.0,166.7 857.9,166.5 855.8,166.1 853.9,165.5 851.9,164.9 849.8,164.4 847.6,163.9 845.2,163.4 842.7,162.8 840.3,161.8 838.2,160.3 836.7,158.5 836.0,156.4 836.2,154.2 837.1,152.0 838.5,150.0 840.0,148.3 841.3,146.7 842.3,145.3 842.9,143.7 843.1,142.0 843.3,140.1 843.8,138.0 844.6,136.0 846.0,134.1 847.8,132.7 850.1,131.7 852.5,131.2 855.1,131.0 857.6,131.0 860.0,131.2 862.4,131.4 864.7,131.8 867.0,132.4 869.1,133.2 871.0,134.4 872.5,135.9 873.6,137.7 874.3,139.5 874.8,141.2 875.4,142.7 876.2,144.1 877.5,145.3 879.2,146.7 881.2,148.2 883.1,150.0" stroke-width="0.7"/><polygon points="908.4,150.0 911.2,154.4 912.4,159.1 911.7,163.8 909.0,167.9 904.8,171.2 899.7,173.6 894.2,175.3 889.0,176.4 884.2,177.3 880.0,178.3 876.0,179.5 872.2,180.8 868.3,182.0 864.2,182.9 860.0,183.2 855.8,183.0 851.6,182.3 847.6,181.4 843.5,180.5 839.2,179.6 834.5,178.7 829.6,177.7 824.5,176.2 819.8,174.0 815.9,170.9 813.3,167.1 812.3,162.7 813.0,158.2 815.1,153.9 817.9,150.0 820.7,146.6 823.1,143.6 824.6,140.6 825.4,137.4 825.6,133.7 825.9,129.7 826.8,125.5 828.6,121.4 831.6,117.9 835.5,115.2 840.2,113.5 845.2,112.7 850.3,112.5 855.2,112.7 860.0,113.1 864.7,113.6 869.2,114.4 873.7,115.5 877.8,117.2 881.5,119.5 884.4,122.4 886.7,125.7 888.3,129.1 889.7,132.3 891.3,135.2 893.5,137.8 896.6,140.3 900.4,143.0 904.5,146.2 908.4,150.0" stroke-width="0.7"/><polygon points="935.3,150.0 938.8,156.8 939.6,163.9 937.4,170.6 932.5,176.5 925.5,181.0 917.4,184.2 909.2,186.3 901.6,187.8 894.8,189.3 888.9,191.0 883.4,193.1 878.0,195.4 872.3,197.5 866.3,199.1 860.0,199.9 853.6,199.9 847.2,199.2 840.9,198.1 834.5,196.9 827.8,195.7 820.6,194.5 813.1,192.7 805.5,190.2 798.7,186.5 793.3,181.6 790.1,175.5 789.3,168.8 790.9,162.0 794.2,155.7 798.4,150.0 802.3,145.0 805.3,140.5 807.0,135.9 807.6,130.9 807.6,125.2 807.9,119.0 809.3,112.6 812.3,106.6 817.0,101.5 823.3,97.8 830.5,95.6 838.1,94.6 845.6,94.6 853.0,95.1 860.0,95.7 866.9,96.5 873.6,97.6 880.1,99.2 886.3,101.6 891.7,104.9 896.2,109.1 899.8,113.8 902.6,118.5 905.3,123.0 908.3,127.1 912.4,130.9 917.6,134.7 923.6,138.9 929.9,144.0 935.3,150.0" stroke-width="0.7"/><polygon points="963.1,150.0 966.6,159.2 966.3,168.5 962.1,177.2 954.3,184.4 944.3,189.9 933.3,193.6 922.5,196.2 912.8,198.1 904.5,200.2 897.3,202.9 890.5,206.2 883.7,209.9 876.4,213.3 868.4,215.9 860.0,217.4 851.3,217.7 842.6,217.1 833.9,215.9 825.0,214.5 815.7,213.0 805.8,211.1 795.8,208.5 786.0,204.7 777.4,199.2 771.0,192.2 767.5,183.8 767.3,174.7 770.1,165.7 774.7,157.3 780.1,150.0 784.8,143.5 788.0,137.4 789.4,131.2 789.4,124.2 789.0,116.4 789.3,107.9 791.3,99.3 795.7,91.4 802.5,85.1 811.1,80.6 820.9,78.0 831.1,77.2 841.2,77.3 850.8,78.0 860.0,78.8 869.0,79.8 877.9,81.0 886.5,83.0 894.7,86.1 902.0,90.4 908.1,95.7 913.2,101.6 917.5,107.6 921.7,113.3 926.6,118.5 932.8,123.4 940.3,128.6 948.6,134.6 956.7,141.7 963.1,150.0" stroke-width="0.7"/><polygon points="991.1,150.0 993.9,161.5 991.8,173.0 984.9,183.3 974.1,191.7 960.9,197.8 947.2,201.9 934.3,204.8 923.1,207.4 913.7,210.6 905.5,214.6 897.8,219.6 889.7,225.0 880.7,229.9 870.7,233.8 860.0,236.1 848.9,236.9 837.6,236.3 826.3,235.0 814.8,233.3 802.8,231.2 790.4,228.5 777.9,224.7 766.2,219.3 756.3,211.8 749.3,202.4 746.2,191.6 746.9,180.1 750.9,169.0 756.8,158.9 763.1,150.0 768.1,142.1 770.9,134.5 771.6,126.4 770.7,117.4 769.7,107.2 770.0,96.4 772.9,85.7 778.9,76.1 787.8,68.6 799.1,63.5 811.6,60.9 824.4,60.1 836.8,60.5 848.6,61.4 860.0,62.3 871.1,63.2 882.1,64.6 893.0,66.8 903.2,70.5 912.4,75.6 920.3,81.9 927.0,89.0 933.0,96.1 939.1,102.9 946.1,109.3 954.5,115.5 964.3,122.2 974.7,130.0 984.2,139.3 991.1,150.0" stroke-width="1.3"/><polygon points="1018.4,150.0 1019.9,163.8 1015.4,177.1 1005.5,188.8 991.5,198.0 975.5,204.7 959.4,209.2 944.9,212.7 932.7,216.2 922.7,220.8 914.0,226.7 905.5,233.8 896.2,241.3 885.4,248.0 873.2,253.2 860.0,256.4 846.2,257.5 832.3,256.9 818.3,255.3 804.1,253.0 789.5,250.1 774.7,246.3 760.1,241.0 746.7,233.6 735.9,223.9 729.0,212.0 726.5,198.7 728.4,185.1 733.7,172.0 740.6,160.3 747.3,150.0 752.0,140.7 753.9,131.5 753.3,121.6 751.3,110.3 749.6,97.7 750.0,84.5 753.9,71.7 761.8,60.5 773.2,52.0 787.2,46.6 802.4,44.0 817.8,43.4 832.5,44.0 846.5,45.0 860.0,45.8 873.3,46.6 886.5,47.9 899.5,50.4 911.9,54.5 923.1,60.4 932.9,67.7 941.4,75.9 949.2,84.1 957.3,92.0 966.5,99.6 977.3,107.2 989.3,115.5 1001.4,125.4 1011.7,136.9 1018.4,150.0" stroke-width="0.7"/><polygon points="1044.5,150.0 1044.0,165.9 1036.9,180.8 1023.8,193.6 1006.7,203.6 988.1,210.7 970.3,215.7 954.9,220.1 942.4,225.0 932.2,231.5 923.3,239.9 914.0,249.5 903.3,259.3 890.6,268.0 876.0,274.5 860.0,278.4 843.4,279.6 826.6,278.7 809.9,276.5 793.1,273.3 776.0,269.3 759.0,264.0 742.6,256.9 728.2,247.3 717.0,235.2 710.4,220.8 708.9,205.2 712.0,189.4 718.4,174.7 725.9,161.6 732.4,150.0 736.1,139.3 736.5,128.5 734.2,116.5 730.7,102.8 728.3,87.7 729.0,72.0 734.3,57.2 744.3,44.6 758.4,35.4 775.3,29.7 793.4,27.3 811.3,27.0 828.3,27.8 844.4,28.6 860.0,29.2 875.4,29.7 890.9,30.9 906.1,33.5 920.7,38.2 934.0,44.9 945.8,53.2 956.2,62.4 966.0,71.8 976.2,80.8 987.6,89.6 1000.6,98.7 1014.6,108.8 1027.9,120.7 1038.6,134.6 1044.5,150.0" stroke-width="0.7"/><polygon points="1069.0,150.0 1066.2,167.8 1056.0,184.2 1039.9,197.9 1020.1,208.4 999.5,216.1 980.7,221.9 965.0,227.5 952.7,234.4 942.8,243.4 933.6,254.6 923.6,267.1 911.3,279.4 896.3,290.0 878.9,297.6 860.0,301.8 840.4,302.8 820.8,301.2 801.4,297.9 782.1,293.5 762.8,288.0 744.0,281.0 726.3,271.8 711.1,259.9 699.9,245.4 694.0,228.6 693.5,210.8 697.7,193.2 704.8,177.1 712.2,162.7 717.8,150.0 719.8,137.9 718.1,125.3 713.6,111.0 708.6,94.7 705.6,76.9 706.9,58.8 713.8,42.1 726.4,28.3 743.6,18.6 763.5,13.0 784.4,10.8 804.8,10.8 824.1,11.6 842.3,12.3 860.0,12.5 877.6,12.7 895.3,13.7 912.9,16.5 929.7,21.6 945.1,29.1 958.9,38.4 971.2,48.7 983.0,59.2 995.3,69.4 1009.0,79.5 1024.1,90.1 1039.7,102.1 1053.9,116.2 1064.4,132.4 1069.0,150.0" stroke-width="0.7"/><polygon points="1091.7,150.0 1086.4,169.5 1073.2,187.2 1054.2,201.8 1032.2,212.9 1010.4,221.2 991.2,228.2 975.9,235.5 964.2,244.9 954.7,256.9 945.4,271.3 934.3,286.8 920.1,301.6 902.5,313.8 882.1,322.2 860.0,326.3 837.4,326.5 815.0,323.8 793.0,319.1 771.5,313.0 750.4,305.7 730.0,296.7 711.4,285.4 695.9,271.2 685.0,254.3 679.8,235.3 680.2,215.6 685.2,196.6 692.3,179.2 698.9,163.9 702.7,150.0 702.3,136.4 697.8,121.7 690.9,105.0 684.4,85.9 681.1,65.3 683.5,44.9 692.6,26.4 708.2,11.7 728.6,1.7 751.8,-3.6 775.6,-5.4 798.5,-5.1 820.0,-4.3 840.3,-3.9 860.0,-4.2 879.8,-4.4 899.8,-3.6 919.7,-0.7 938.8,4.9 956.3,13.3 972.0,23.6 986.3,35.0 1000.1,46.6 1014.6,57.9 1030.4,69.3 1047.4,81.6 1064.3,95.6 1079.0,111.8 1088.9,130.3 1091.7,150.0" stroke-width="0.7"/><polygon points="1112.9,150.0 1104.9,171.1 1088.9,189.9 1067.5,205.3 1043.9,217.1 1021.5,226.4 1002.6,235.0 988.2,244.6 977.5,257.0 968.5,272.4 958.8,290.3 946.2,308.8 929.6,325.8 909.0,339.2 885.3,347.7 860.0,351.1 834.4,350.0 809.3,345.6 785.1,339.0 761.7,331.1 739.0,321.9 717.5,310.8 698.2,297.3 682.7,280.9 672.2,261.9 667.6,241.1 668.7,219.9 673.7,199.6 680.1,181.4 685.0,165.1 686.2,150.0 682.6,134.7 675.0,117.8 665.7,98.2 657.9,76.2 654.8,52.9 658.8,30.1 670.7,10.3 689.8,-5.0 713.8,-15.0 740.4,-19.9 767.1,-21.1 792.4,-20.5 816.0,-19.8 838.2,-19.8 860.0,-20.6 882.0,-21.4 904.3,-20.9 926.5,-17.8 947.7,-11.6 967.3,-2.4 985.0,8.9 1001.2,21.4 1017.1,34.0 1033.7,46.5 1051.5,59.3 1070.3,73.2 1088.3,89.2 1103.1,107.6 1112.0,128.3 1112.9,150.0" stroke-width="1.3"/><polygon points="1132.8,150.0 1122.2,172.6 1103.7,192.5 1080.4,208.7 1055.8,221.5 1033.5,232.1 1015.5,242.6 1002.4,255.1 992.7,270.8 984.0,290.0 973.6,311.3 959.1,332.6 939.8,351.3 915.8,365.3 888.6,373.3 860.0,375.3 831.5,372.3 804.1,365.8 777.9,357.2 752.9,347.3 729.0,336.0 706.7,323.1 687.0,307.6 671.4,289.3 661.2,268.5 656.9,246.1 658.0,223.8 662.3,202.7 667.1,183.6 669.4,166.4 667.2,150.0 660.0,132.8 649.1,113.2 637.5,90.7 628.9,65.6 626.7,39.6 633.1,14.8 648.5,-6.2 671.5,-21.7 699.5,-31.1 729.5,-35.3 759.1,-35.9 786.7,-35.0 812.2,-34.5 836.3,-35.1 860.0,-36.7 884.1,-38.1 908.7,-37.8 933.1,-34.6 956.6,-27.9 978.2,-17.9 997.8,-5.6 1016.1,7.9 1034.0,21.5 1052.7,35.2 1072.6,49.4 1092.9,65.0 1111.8,82.9 1126.4,103.6 1134.1,126.4 1132.8,150.0" stroke-width="0.7"/><polygon points="1151.9,150.0 1138.9,174.0 1118.2,195.0 1093.4,212.2 1068.5,226.1 1046.8,238.5 1030.2,251.4 1018.6,267.1 1009.9,286.5 1001.2,309.4 989.6,334.0 972.8,357.7 950.2,377.5 922.6,391.4 891.8,398.1 860.0,398.1 828.9,392.8 799.4,383.9 771.6,373.1 745.3,361.2 720.5,348.2 697.4,333.5 677.3,316.4 661.5,296.5 651.3,274.3 646.8,250.9 647.2,227.7 650.0,206.0 652.3,186.2 651.2,168.0 645.0,150.0 633.8,130.5 619.8,108.1 606.4,82.4 597.7,54.3 597.3,25.6 606.9,-0.8 626.5,-22.4 653.9,-37.7 686.0,-46.3 719.5,-49.5 751.7,-49.4 781.4,-48.4 808.6,-48.2 834.4,-49.6 860.0,-52.1 886.2,-54.2 913.0,-54.4 939.7,-51.1 965.3,-43.9 988.9,-33.1 1010.6,-20.0 1030.9,-5.6 1050.9,9.0 1071.8,23.8 1093.6,39.4 1115.4,56.8 1134.9,76.8 1149.2,99.6 1155.4,124.5 1151.9,150.0" stroke-width="0.7"/><polygon points="1170.6,150.0 1155.2,175.4 1132.7,197.5 1107.1,215.8 1082.4,231.2 1061.9,245.6 1047.0,261.4 1036.8,280.6 1028.9,303.8 1019.8,330.4 1006.4,357.9 986.7,383.4 960.5,403.6 929.1,416.5 894.8,421.3 860.0,418.8 826.6,410.8 795.3,399.4 766.3,386.5 739.0,372.8 713.3,358.4 689.5,342.4 668.8,324.1 652.5,303.2 641.7,280.1 636.3,255.9 635.2,232.1 635.6,209.8 634.5,189.3 629.4,169.9 618.9,150.0 603.8,127.9 587.0,102.4 572.6,73.4 564.9,42.3 567.2,11.4 581.0,-16.2 605.4,-38.0 637.6,-52.5 674.0,-60.0 710.7,-62.0 745.3,-61.2 776.7,-60.1 805.4,-60.6 832.7,-63.1 860.0,-66.8 888.2,-70.0 917.2,-70.7 946.2,-67.5 973.9,-59.8 999.7,-48.4 1023.4,-34.4 1045.8,-19.2 1068.1,-3.6 1091.1,12.3 1114.9,29.3 1138.1,48.5 1158.0,70.6 1171.7,95.7 1176.3,122.7 1170.6,150.0" stroke-width="0.7"/><polygon points="1203.6,470.0 1203.6,472.0 1203.5,474.1 1202.9,476.1 1201.9,478.0 1200.3,479.6 1198.3,480.9 1196.0,481.8 1193.6,482.4 1191.3,482.8 1189.3,483.3 1187.6,483.9 1185.9,484.9 1184.2,486.2 1182.3,487.6 1180.0,489.0 1177.4,489.9 1174.8,490.2 1172.1,489.8 1169.8,488.7 1167.9,487.1 1166.6,485.2 1165.5,483.2 1164.7,481.3 1164.0,479.5 1163.1,478.0 1162.1,476.5 1161.0,475.1 1160.0,473.5 1159.3,471.8 1158.9,470.0 1158.9,468.2 1159.2,466.4 1159.6,464.6 1160.1,462.8 1160.7,460.9 1161.3,458.8 1162.0,456.7 1163.1,454.6 1164.7,452.7 1166.8,451.3 1169.4,450.4 1172.2,450.3 1175.0,450.9 1177.7,452.0 1180.0,453.4 1181.9,454.8 1183.6,456.0 1185.3,456.7 1187.0,457.0 1189.1,457.1 1191.4,457.2 1193.9,457.4 1196.3,458.0 1198.5,459.0 1200.3,460.4 1201.6,462.1 1202.5,464.0 1203.0,466.0 1203.4,468.0 1203.6,470.0" stroke-width="0.7"/><polygon points="1227.4,470.0 1227.3,474.1 1226.6,478.1 1225.3,482.1 1222.9,485.7 1219.6,488.8 1215.6,491.2 1211.0,492.9 1206.5,494.1 1202.2,495.1 1198.5,496.3 1195.2,497.9 1192.0,500.2 1188.6,503.1 1184.6,506.0 1180.0,508.6 1174.8,510.3 1169.5,510.7 1164.3,509.5 1159.9,507.1 1156.3,503.6 1153.7,499.7 1151.7,495.8 1150.1,492.1 1148.3,488.9 1146.3,486.0 1144.0,483.1 1141.6,480.2 1139.4,477.1 1137.7,473.6 1136.9,470.0 1136.8,466.3 1137.3,462.6 1138.3,458.9 1139.3,455.2 1140.5,451.3 1141.8,447.3 1143.6,443.1 1146.1,439.1 1149.6,435.7 1154.1,433.1 1159.3,431.9 1165.0,432.1 1170.5,433.5 1175.6,435.9 1180.0,438.5 1183.7,441.0 1187.1,442.7 1190.4,443.6 1194.2,443.8 1198.6,443.5 1203.6,443.4 1208.8,443.8 1213.8,445.0 1218.3,447.2 1221.8,450.2 1224.3,453.8 1225.8,457.8 1226.7,461.9 1227.2,465.9 1227.4,470.0" stroke-width="0.7"/><polygon points="1250.9,470.0 1250.3,476.1 1249.0,482.0 1246.6,487.7 1242.9,493.0 1237.8,497.4 1231.7,500.8 1225.2,503.3 1218.7,505.2 1212.8,507.0 1207.6,509.2 1202.9,512.2 1198.3,516.2 1193.2,520.7 1187.1,525.3 1180.0,529.1 1172.1,531.2 1164.1,531.3 1156.6,529.1 1150.2,525.0 1145.1,519.6 1141.3,513.7 1138.4,507.9 1135.7,502.7 1132.8,498.1 1129.3,494.0 1125.4,489.9 1121.4,485.6 1117.9,480.8 1115.3,475.6 1114.0,470.0 1113.9,464.3 1114.7,458.6 1116.2,453.0 1118.0,447.4 1119.9,441.6 1122.2,435.6 1125.3,429.6 1129.4,424.0 1135.0,419.2 1142.0,416.0 1150.0,414.7 1158.4,415.5 1166.5,417.9 1173.8,421.3 1180.0,424.9 1185.4,427.8 1190.5,429.6 1195.8,430.1 1201.9,429.6 1209.0,428.8 1216.9,428.3 1225.1,428.9 1232.8,431.0 1239.3,434.6 1244.3,439.5 1247.7,445.3 1249.7,451.4 1250.6,457.7 1251.0,463.9 1250.9,470.0" stroke-width="0.7"/><polygon points="1273.4,470.0 1272.2,477.9 1270.1,485.7 1266.6,493.1 1261.5,499.7 1254.7,505.4 1246.8,509.8 1238.5,513.2 1230.5,516.0 1223.3,518.8 1216.9,522.4 1211.0,527.2 1205.0,533.0 1198.0,539.5 1189.7,545.7 1180.0,550.4 1169.4,552.7 1158.7,552.0 1148.9,548.4 1140.7,542.4 1134.2,535.0 1129.4,527.1 1125.4,519.7 1121.6,513.1 1117.2,507.4 1112.0,502.2 1106.3,496.9 1100.5,491.2 1095.6,484.7 1092.1,477.6 1090.4,470.0 1090.4,462.3 1091.8,454.6 1093.9,447.1 1096.5,439.5 1099.4,431.8 1102.8,424.0 1107.3,416.3 1113.4,409.3 1121.2,403.7 1130.8,400.1 1141.4,399.0 1152.4,400.3 1162.8,403.6 1172.0,407.8 1180.0,411.8 1187.1,414.7 1194.0,416.0 1201.6,415.5 1210.4,414.0 1220.6,412.4 1231.7,411.7 1242.8,412.8 1253.1,416.1 1261.5,421.4 1267.6,428.5 1271.5,436.6 1273.4,445.1 1274.1,453.6 1274.0,461.9 1273.4,470.0" stroke-width="0.7"/><polygon points="1294.4,470.0 1292.6,479.7 1289.7,489.1 1285.1,498.0 1278.7,506.0 1270.5,512.8 1261.1,518.3 1251.4,522.7 1242.2,526.7 1234.0,531.0 1226.7,536.3 1219.7,543.1 1212.1,551.1 1203.2,559.6 1192.5,567.2 1180.0,572.5 1166.6,574.6 1153.3,572.9 1141.4,567.5 1131.4,559.4 1123.8,549.9 1117.8,540.2 1112.8,531.2 1107.6,523.5 1101.5,516.8 1094.3,510.6 1086.5,504.1 1078.9,496.9 1072.6,488.7 1068.3,479.6 1066.4,470.0 1066.8,460.2 1068.8,450.6 1071.8,441.2 1075.3,431.8 1079.2,422.3 1083.9,412.8 1090.0,403.6 1098.0,395.4 1108.2,389.0 1120.3,385.2 1133.5,384.4 1146.8,386.3 1159.3,390.2 1170.4,394.7 1180.0,398.7 1188.8,401.0 1197.9,401.1 1208.0,399.3 1219.9,396.6 1233.3,394.2 1247.7,393.6 1261.9,395.5 1274.4,400.3 1284.3,407.9 1291.1,417.4 1294.9,428.1 1296.4,439.0 1296.4,449.7 1295.7,460.0 1294.4,470.0" stroke-width="1.3"/><polygon points="1313.6,470.0 1311.2,481.3 1307.6,492.2 1302.2,502.6 1294.8,511.9 1285.5,519.9 1275.0,526.6 1264.4,532.3 1254.4,537.7 1245.4,543.8 1237.3,551.3 1229.1,560.4 1219.9,570.7 1208.8,581.1 1195.4,589.8 1180.0,595.5 1163.7,596.9 1148.0,593.6 1134.0,586.1 1122.6,575.8 1113.7,564.1 1106.8,552.7 1100.5,542.4 1093.7,533.7 1085.6,526.2 1076.3,519.1 1066.3,511.5 1056.9,502.8 1049.3,492.8 1044.4,481.7 1042.5,470.0 1043.3,458.2 1046.1,446.7 1050.0,435.4 1054.5,424.2 1059.6,413.0 1065.6,401.9 1073.4,391.3 1083.4,382.0 1095.9,375.1 1110.4,371.2 1126.0,370.6 1141.5,372.9 1155.9,377.1 1168.7,381.6 1180.0,384.9 1190.8,385.9 1202.2,384.5 1215.2,381.2 1230.4,377.2 1247.3,374.4 1265.0,374.1 1281.8,377.3 1296.2,384.2 1307.0,394.3 1313.8,406.6 1317.1,419.9 1317.8,433.3 1317.0,446.1 1315.5,458.3 1313.6,470.0" stroke-width="0.7"/><polygon points="1330.8,470.0 1328.1,482.8 1324.1,495.1 1318.3,506.9 1310.3,517.6 1300.3,527.0 1289.2,535.0 1277.9,542.3 1267.4,549.6 1257.9,557.9 1248.9,567.8 1239.4,579.5 1228.3,592.0 1214.7,604.0 1198.4,613.6 1180.0,619.0 1160.9,619.1 1142.7,613.8 1126.9,603.9 1114.2,591.2 1104.3,577.5 1096.3,564.5 1088.6,553.2 1080.0,543.8 1069.8,535.7 1058.0,527.7 1045.8,519.0 1034.6,508.8 1025.8,496.9 1020.4,483.8 1018.7,470.0 1020.1,456.2 1023.8,442.8 1028.7,429.7 1034.3,416.8 1040.5,403.9 1047.9,391.3 1057.3,379.4 1069.3,369.2 1084.1,361.7 1100.9,357.7 1118.8,357.3 1136.3,359.7 1152.5,363.8 1166.9,367.6 1180.0,369.8 1192.9,369.2 1207.0,366.0 1223.2,361.0 1241.9,356.1 1262.3,353.1 1283.0,353.7 1302.1,358.8 1317.8,368.3 1328.9,381.3 1335.3,396.5 1337.6,412.5 1337.1,428.1 1335.3,442.9 1333.1,456.8 1330.8,470.0" stroke-width="0.7"/><polygon points="1346.3,470.0 1343.7,484.1 1339.9,497.9 1334.1,511.1 1325.9,523.3 1315.7,534.2 1304.3,544.0 1292.7,553.2 1281.8,562.7 1271.7,573.5 1261.7,586.1 1250.7,600.2 1237.4,614.9 1221.0,628.1 1201.5,637.9 1180.0,642.5 1158.1,640.8 1137.7,633.0 1120.3,620.5 1106.5,605.4 1095.7,589.8 1086.5,575.6 1077.3,563.6 1066.6,553.7 1053.9,545.1 1039.6,536.5 1025.0,526.6 1012.0,514.8 1002.1,501.0 996.5,485.8 995.1,470.0 997.4,454.3 1002.0,439.0 1007.9,424.1 1014.4,409.5 1021.7,395.1 1030.5,380.9 1041.6,367.8 1055.6,356.7 1072.5,348.7 1091.6,344.5 1111.6,344.0 1131.0,346.3 1148.8,349.8 1165.0,352.6 1180.0,353.2 1195.3,350.7 1212.3,345.5 1231.9,339.1 1254.1,333.5 1277.9,330.9 1301.4,333.0 1322.2,340.5 1338.5,353.0 1349.3,369.1 1354.6,387.3 1355.6,405.9 1354.0,423.7 1351.3,440.1 1348.7,455.5 1346.3,470.0" stroke-width="0.7"/><polygon points="1360.7,470.0 1358.9,485.4 1355.8,500.6 1350.5,515.4 1342.6,529.4 1332.4,542.1 1320.9,553.9 1309.2,565.4 1297.9,577.4 1287.1,590.8 1275.8,606.1 1262.9,622.6 1247.0,639.0 1227.4,653.0 1204.7,662.4 1180.0,665.4 1155.5,661.3 1133.1,650.8 1114.4,635.5 1099.6,618.1 1087.9,600.9 1077.5,585.7 1066.5,573.4 1053.4,563.4 1038.0,554.6 1020.9,545.3 1003.9,534.3 989.1,520.9 978.3,505.2 972.6,487.9 971.7,470.0 974.8,452.3 980.4,435.2 987.2,418.6 994.7,402.4 1003.1,386.3 1013.2,370.6 1025.9,356.2 1041.9,344.2 1061.0,335.7 1082.2,331.1 1104.2,330.4 1125.4,332.3 1145.0,334.9 1162.9,336.3 1180.0,334.9 1197.9,330.4 1218.0,323.4 1241.1,315.8 1266.9,310.0 1293.8,308.4 1319.4,312.6 1341.4,323.0 1357.7,338.8 1367.5,358.3 1371.4,379.4 1371.0,400.3 1368.2,419.8 1365.1,437.7 1362.5,454.3 1360.7,470.0" stroke-width="0.7"/><polygon points="580.8,520.0 579.6,521.7 578.7,523.3 578.2,524.8 577.9,526.5 577.7,528.4 577.3,530.3 576.6,532.3 575.4,534.0 573.7,535.5 571.7,536.6 569.5,537.4 567.2,538.1 564.8,538.5 562.4,538.9 560.0,539.2 557.5,539.3 555.1,539.1 552.7,538.4 550.6,537.3 548.9,535.7 547.7,533.9 547.0,531.9 546.4,530.0 545.9,528.4 545.1,527.0 543.9,525.9 542.2,524.7 540.1,523.5 538.0,521.9 536.1,520.0 534.8,517.8 534.4,515.5 534.9,513.3 536.2,511.3 538.1,509.6 540.3,508.3 542.6,507.2 544.9,506.3 547.1,505.4 549.2,504.6 551.3,503.9 553.4,503.4 555.6,503.2 557.9,503.3 560.0,503.7 562.0,504.2 563.9,504.8 565.8,505.3 567.8,505.6 570.1,505.7 572.6,505.8 575.3,506.1 578.0,506.7 580.4,507.8 582.3,509.4 583.4,511.4 583.7,513.7 583.1,516.0 582.1,518.1 580.8,520.0" stroke-width="0.7"/><polygon points="600.7,520.0 598.6,523.3 597.1,526.5 596.5,529.7 596.2,533.2 596.0,537.0 595.3,541.0 593.7,544.9 591.1,548.3 587.6,551.2 583.5,553.3 578.9,554.9 574.2,555.9 569.5,556.8 564.8,557.4 560.0,557.9 555.1,557.9 550.3,557.4 545.7,556.0 541.7,553.7 538.5,550.6 536.1,547.0 534.4,543.3 533.1,539.8 531.6,536.9 529.5,534.4 526.5,532.2 522.5,530.0 518.1,527.3 513.7,524.0 510.2,520.0 508.1,515.5 507.8,510.9 509.5,506.5 512.6,502.7 516.9,499.6 521.6,497.1 526.3,495.1 530.8,493.4 535.0,491.8 539.0,490.2 543.1,488.8 547.2,487.8 551.5,487.2 555.8,487.3 560.0,487.8 564.0,488.6 567.9,489.4 571.9,490.1 576.1,490.4 580.8,490.5 586.0,490.7 591.5,491.3 596.9,492.8 601.6,495.2 605.0,498.7 606.8,502.9 606.8,507.5 605.4,512.1 603.2,516.3 600.7,520.0" stroke-width="0.7"/><polygon points="619.7,520.0 616.9,524.9 615.4,529.6 614.9,534.6 614.9,540.1 614.8,545.9 613.7,552.0 611.2,557.8 607.1,562.9 601.6,566.9 595.2,570.0 588.3,572.0 581.2,573.5 574.2,574.6 567.1,575.5 560.0,576.1 552.8,576.1 545.7,575.1 539.0,572.9 533.2,569.4 528.4,564.9 524.8,559.7 522.0,554.6 519.6,549.9 516.7,545.8 512.7,542.4 507.4,539.2 500.9,535.8 493.9,531.5 487.5,526.2 482.8,520.0 480.5,513.1 481.1,506.2 484.5,499.9 490.0,494.5 496.9,490.1 504.2,486.8 511.3,484.0 517.8,481.6 523.8,479.2 529.6,476.8 535.3,474.6 541.3,472.8 547.5,471.7 553.8,471.5 560.0,471.9 566.1,472.7 572.1,473.4 578.3,473.9 585.0,474.0 592.4,474.0 600.4,474.4 608.8,475.6 616.6,478.2 623.2,482.3 627.7,488.0 629.7,494.6 629.1,501.6 626.7,508.4 623.2,514.6 619.7,520.0" stroke-width="0.7"/><polygon points="637.6,520.0 634.6,526.4 633.3,532.8 633.4,539.6 634.0,547.0 634.0,555.0 632.5,563.2 628.9,570.9 623.1,577.5 615.5,582.7 606.8,586.4 597.4,589.0 588.0,590.7 578.7,592.1 569.4,593.2 560.0,593.9 550.5,593.8 541.2,592.4 532.5,589.4 524.8,584.8 518.5,578.9 513.5,572.5 509.4,566.1 505.4,560.3 500.6,555.4 494.4,551.1 486.5,546.8 477.3,542.0 468.0,536.0 460.0,528.6 454.6,520.0 452.8,510.8 455.0,501.7 460.7,493.5 469.0,486.8 478.7,481.5 488.5,477.4 497.7,474.0 506.0,470.8 513.5,467.5 520.6,464.1 527.8,460.8 535.4,458.0 543.5,456.2 551.7,455.3 560.0,455.4 568.2,455.8 576.5,456.3 585.2,456.5 594.5,456.4 604.8,456.3 615.8,457.0 626.8,459.1 636.9,463.2 645.0,469.4 650.0,477.4 651.7,486.5 650.2,496.0 646.6,504.9 641.9,512.9 637.6,520.0" stroke-width="0.7"/><polygon points="654.5,520.0 651.7,527.9 651.2,535.9 652.2,544.6 653.5,554.1 653.7,564.4 651.8,574.7 646.9,584.2 639.3,592.2 629.4,598.4 618.2,602.7 606.5,605.7 594.8,607.7 583.2,609.4 571.6,610.7 560.0,611.6 548.3,611.3 536.8,609.5 526.0,605.8 516.5,600.1 508.5,593.1 502.0,585.4 496.4,578.0 490.5,571.3 483.5,565.6 474.7,560.4 464.0,555.1 452.3,548.7 440.9,540.8 431.8,531.0 426.5,520.0 425.9,508.4 430.3,497.4 438.8,487.7 450.1,479.9 462.5,473.8 474.6,469.1 485.5,465.0 495.1,460.9 503.6,456.4 511.8,451.6 520.3,446.9 529.5,442.9 539.3,440.0 549.5,438.3 560.0,437.7 570.6,437.7 581.3,437.7 592.7,437.6 604.9,437.3 618.1,437.4 631.9,438.8 645.4,442.2 657.3,448.1 666.3,456.7 671.4,467.3 672.3,479.0 669.7,490.8 664.7,501.7 659.1,511.5 654.5,520.0" stroke-width="1.3"/><polygon points="670.3,520.0 668.4,529.3 669.1,539.0 671.4,549.7 673.7,561.5 674.2,574.0 671.6,586.5 665.3,597.8 655.6,607.1 643.4,614.1 629.7,619.0 615.5,622.3 601.5,624.6 587.6,626.6 573.9,628.2 560.0,629.2 546.0,628.9 532.3,626.7 519.5,622.2 508.1,615.5 498.5,607.4 490.3,598.7 482.8,590.3 474.8,582.9 465.3,576.4 453.7,570.3 440.3,563.7 426.2,555.6 413.3,545.6 403.7,533.5 399.2,520.0 400.5,506.3 407.6,493.4 419.2,482.5 433.4,473.8 448.3,467.1 462.2,461.7 474.3,456.7 484.7,451.4 493.9,445.4 502.9,438.8 512.4,432.4 523.1,426.8 534.8,422.6 547.2,419.9 560.0,418.6 573.1,418.0 586.5,417.6 600.7,417.2 615.9,417.0 632.0,417.7 648.5,420.1 664.1,425.2 677.3,433.4 686.7,444.5 691.3,457.8 691.1,472.1 687.1,486.1 681.0,498.9 674.8,510.1 670.3,520.0" stroke-width="0.7"/></g><g stroke="currentColor" stroke-width="1"><line x1="0" y1="0" x2="0" y2="18"/><line x1="60" y1="0" x2="60" y2="10"/><line x1="120" y1="0" x2="120" y2="10"/><line x1="180" y1="0" x2="180" y2="10"/><line x1="240" y1="0" x2="240" y2="10"/><line x1="300" y1="0" x2="300" y2="18"/><line x1="360" y1="0" x2="360" y2="10"/><line x1="420" y1="0" x2="420" y2="10"/><line x1="480" y1="0" x2="480" y2="10"/><line x1="540" y1="0" x2="540" y2="10"/><line x1="600" y1="0" x2="600" y2="18"/><line x1="660" y1="0" x2="660" y2="10"/><line x1="720" y1="0" x2="720" y2="10"/><line x1="780" y1="0" x2="780" y2="10"/><line x1="840" y1="0" x2="840" y2="10"/><line x1="900" y1="0" x2="900" y2="18"/><line x1="960" y1="0" x2="960" y2="10"/><line x1="1020" y1="0" x2="1020" y2="10"/><line x1="1080" y1="0" x2="1080" y2="10"/><line x1="1140" y1="0" x2="1140" y2="10"/><line x1="1200" y1="0" x2="1200" y2="18"/><line x1="1260" y1="0" x2="1260" y2="10"/><line x1="1320" y1="0" x2="1320" y2="10"/><line x1="1380" y1="0" x2="1380" y2="10"/><line x1="0" y1="0" x2="18" y2="0"/><line x1="0" y1="60" x2="10" y2="60"/><line x1="0" y1="120" x2="10" y2="120"/><line x1="0" y1="180" x2="10" y2="180"/><line x1="0" y1="240" x2="10" y2="240"/><line x1="0" y1="300" x2="18" y2="300"/><line x1="0" y1="360" x2="10" y2="360"/><line x1="0" y1="420" x2="10" y2="420"/><line x1="0" y1="480" x2="10" y2="480"/><line x1="0" y1="540" x2="10" y2="540"/><line x1="0" y1="600" x2="18" y2="600"/><line x1="0" y1="660" x2="10" y2="660"/></g></svg>'


def main():
    data = load(XLSX)
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    import base64, os
    rep = ""
    if REPORT_PDF:
        with open(REPORT_PDF, "rb") as f:
            rep = base64.b64encode(f.read()).decode("ascii")
    html = TEMPLATE.replace("__TOPO_SVG__", CONTOUR).replace("__DATA__", payload).replace("__REPORT__", rep)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"saved {OUT}: {len(data['datasets'])} datasets, {len(data['usecases'])} use cases, "
          f"{len(data['themes'])} themes")


if __name__ == "__main__":
    main()
