"""
Turns the results of analyze.py into a review report with charts (HTML) and adds "Conclusions" as the first sheet of the Excel file.

Usage: python report.py <data folder>
  - runs all calculations of analyze.py first (the Excel file is regenerated);
  - reads conclusions.json in the same folder (conclusions and actions, written by the AI following step 6 of SKILL.md);
  - writes review_report_<hotel>.html: conclusions and actions first, charts and data notes below.

Charts are PNG files (drawn by charts.py); they are in the "charts" folder, embedded in the HTML and put in the "Charts" sheet of the Excel file, so nothing needs the internet.
"""
import base64
import html
import json
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import analyze as A  # noqa: E402
from charts import make_charts  # noqa: E402



try:  # the Trip.com English page shows Chinese reviews in traditional characters; convert back to simplified (only if opencc-python-reimplemented is installed)
    from opencc import OpenCC
    _t2s = OpenCC("t2s").convert
except Exception:  # noqa: BLE001
    _t2s = None

LANG_ZH = {"zh": "Chinese", "en": "English", "ko": "Korean", "ja": "Japanese", "zh-CN": "Simplified Chinese", "zh-HK": "Traditional Chinese (Hong Kong)",
           "zh-TW": "Traditional Chinese (Taiwan)", "th": "Thai", "fr": "French", "de": "German", "es": "Spanish", "ru": "Russian", "vi": "Vietnamese"}


def simple(text):
    t = str(text if text is not None else "")
    return _t2s(t).replace("衞", "卫") if _t2s else t


def esc(text):
    return html.escape(str(text if text is not None else ""))


def num(v, digits=2):
    try:
        return round(float(v), digits)
    except (TypeError, ValueError):
        return None


def build_data(res):
    """Take the numbers the charts need from the results."""
    t, reviews, pool, pf, today = res["tables"], res["reviews"], res["pool"], res["platforms"], res["today"]
    overall = t["Summary"]
    ov = dict(zip(overall["Item"], overall["Combined result"]))
    data = {}

    # platform ratings (10-point) and review counts
    rows = []
    for p in pf.rating_platforms():
        rows.append({"label": pf.disp(p), "value": pf.rating10(p), "count": int(pf.published(p) or 0)})
    data["platforms"] = rows


    # What guests complain about: how many of all reviews (including complaints inside praise reviews) complain about each theme. Low-score reviews are often only a few, too small a sample alone
    # All shares and counts use only normally collected reviews; topped-up low-score reviews (lowscore_*.json) are evidence quotes only, in no ratio
    data["neg_total"] = int(len(reviews[reviews["Low score"] == True]))  # noqa: E712
    data["all_total"] = int(len(reviews))
    data["complaint_reviews"] = int((reviews["Complaint themes"].fillna("") != "").sum())
    data["complaints"] = sorted(
        [{"label": th, "value": int(reviews["Complaint themes"].str.contains(re.escape(th)).sum())} for th in A.THEMES],
        key=lambda x: -x["value"])
    data["complaints"] = [x for x in data["complaints"] if x["value"] > 0]

    # theme impact scores
    imp = t["Theme impact scores"].dropna(subset=["Impact score"])
    data["impact"] = [{"label": r["Theme"], "value": num(r["Impact score"]), "count": int(r["Reviews complaining"])}
                      for _, r in imp.iterrows()]

    # sub-scores and guests: merged numbers from the Summary
    data["subscores"] = [{"label": k.replace("Sub-score: ", "").replace(" (10-point)", ""), "value": num(v)}
                         for k, v in ov.items() if str(k).startswith("Sub-score: ") and num(v) is not None]
    data["travellers"] = [{"label": k.replace("Guest: ", ""), "value": num(float(v) * 100, 1)}
                          for k, v in ov.items() if str(k).startswith("Guest: ") and num(v) is not None]

    # platform topics (Ctrip hot tags, Agoda topics)
    topics = t["Platform topics"]
    data["topics"] = []
    if len(topics) and "topic" in topics.columns:
        for _, r in topics.sort_values("mentions", ascending=False).head(12).iterrows():
            data["topics"].append({"label": f"{r['topic']} ({pf.disp(r['Platform'])})", "value": int(r["mentions"])})

    # KPI cards
    neg_n = re.search(r"(\d+) low-score reviews", str(ov.get("Low-score timing", "")))
    neg_recent = re.search(r"(\d+) in the past year", str(ov.get("Low-score timing", "")))
    cr, tot = data["complaint_reviews"], data["all_total"]
    # KPI cards: unified units (points, reviews); shares always "X% of the N reviews analysed", rounded, never mixing denominators
    nn = data["neg_total"]
    data["kpis"] = [
        {"label": "Overall rating", "value": f"{float(ov['Overall rating (10-point)']):.1f}" if num(ov.get("Overall rating (10-point)")) else "—",
         "note": "10-point scale, platform ratings weighted by review counts; platforms calculate differently, so use it as a reference, and judge change on the same platform"},
        {"label": "Reviews published on platforms", "value": f"{int(ov['Total published reviews'])}" if num(ov.get("Total published reviews")) else "—",
         "note": f"{tot} analysed" + "".join(
             f"; {pf.disp(p)} not fully collected" for p, f in res["summary"].get("platforms", {}).items() if f.get("scrape_note"))},
        {"label": "Reviews with complaints", "value": f"{cr}", "note": f"{cr / tot:.0%} of the {tot} reviews analysed" if tot else ""},
        {"label": "Low-score reviews (below 6)", "value": f"{nn}",
         "note": (f"{nn / tot:.0%} of the {tot} reviews analysed" if tot else "") + (f"; {int(neg_recent.group(1))} in the past year" if neg_recent else "")},
    ]
    data["overall"] = overall
    data["compare"] = t["Platform comparison"]
    data["today"] = str(today.date())
    return data


def table_html(df, columns):
    cols = [c for c in columns if c in df.columns]
    head = "".join(f"<th>{esc(c)}</th>" for c in cols)
    body = ""
    for _, r in df.iterrows():
        cells = []
        for c in cols:
            v = r[c]
            if isinstance(v, float) and ("share" in c.lower()) and not pd.isna(v):
                v = f"{v:.1%}"
            elif isinstance(v, pd.Timestamp):
                v = v.date()
            elif isinstance(v, float) and pd.isna(v):
                v = ""
            cells.append(f"<td>{esc(v)}</td>")
        body += "<tr>" + "".join(cells) + "</tr>"
    return f'<div class="tablewrap"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


PALETTE = ["#2a78d6", "#eb6834", "#1f9e74", "#e3a92b", "#8a63d2", "#5f6b7a"]


def donut(parts, center, sub, size=150):
    """Donut chart (SVG, readable in light and dark): parts = [(name, value, colour)], only for compositions that add up to 100% (one review counts in one class).
    Multi-select shares (one review can praise several things) must not be a pie, use horizontal bars."""
    total = sum(v for _, v, _ in parts) or 1
    r, w, cx = 54, 20, size / 2
    import math
    arcs, start = [], -math.pi / 2
    for _, v, col in parts:
        if v <= 0:
            continue
        ang = 2 * math.pi * v / total
        end = start + ang
        if ang >= 2 * math.pi - 1e-6:
            arcs.append(f'<circle cx="{cx}" cy="{cx}" r="{r}" fill="none" stroke="{col}" stroke-width="{w}"/>')
        else:
            x1, y1 = cx + r * math.cos(start), cx + r * math.sin(start)
            x2, y2 = cx + r * math.cos(end), cx + r * math.sin(end)
            arcs.append(f'<path d="M{x1:.2f},{y1:.2f} A{r},{r} 0 {1 if ang > math.pi else 0} 1 {x2:.2f},{y2:.2f}" '
                        f'fill="none" stroke="{col}" stroke-width="{w}"/>')
        start = end
    svg = (f'<svg viewBox="0 0 {size} {size}" width="{size}" height="{size}" role="img">{"".join(arcs)}'
           f'<text x="{cx}" y="{cx - 2}" text-anchor="middle" class="dc">{esc(center)}</text>'
           f'<text x="{cx}" y="{cx + 18}" text-anchor="middle" class="ds">{esc(sub)}</text></svg>')
    legend = "".join(f'<li><i style="background:{col}"></i><span>{esc(lbl)}</span><b>{v}</b>'
                     f'<em>{v / total:.0%}</em></li>' for lbl, v, col in parts)
    return f'<div class="donut">{svg}<ul>{legend}</ul></div>'


def render(res, conclusions, out, images):
    d = build_data(res)
    pf_disp = res["platforms"].disp
    hotel = esc(res["hotel"])
    c = conclusions or {}

    actions = "".join(
        f'<li><span class="badge{" later" if a.get("when") not in ("Now", "First") else ""}">{esc(a.get("when", ""))}</span>'
        f'<div><div class="aline"><strong>{esc(a.get("action", ""))}</strong>'
        + (f'<span class="basis">{esc(a["basis"])}</span>' if a.get("basis") else "") + "</div>"
        # what to do is written out (bullet points); only the "why" is collapsed
        + ("<ul class='adetail'>" + "".join(f"<li>{esc(x)}</li>" for x in (a["detail"] if isinstance(a["detail"], list) else [a["detail"]]))
           + "</ul>" if a.get("detail") else "")
        + (f'<details><summary>Why</summary><p class="muted">{esc(a["why"])}</p></details>' if a.get("why") else "")
        + "</div></li>"
        for a in c.get("actions", []))
    img = {Path(p).stem: (p, name) for p, name in images}

    def fig(key):
        p, name = img[key]
        return (f'<figure class="chart"><img src="data:image/png;base64,{base64.b64encode(Path(p).read_bytes()).decode()}" '
                f'alt="{esc(name)}"></figure>')

    used = set()
    parts = []
    for k, f in enumerate(c.get("findings", []), 1):
        keys = [x for x in f.get("charts", []) if x in img and x not in used]
        used.update(keys)
        ev = ("<details><summary>Evidence (quotes and numbers)</summary><ul class='evidence'>"
              + "".join(f"<li>{esc(e)}</li>" for e in f.get("evidence", [])) + "</ul></details>") if f.get("evidence") else ""
        stat = "".join(f'<div class="stat"><b>{esc(x.get("value", ""))}</b><span>{esc(x.get("label", ""))}</span></div>'
                       for x in f.get("stats", []))
        parts.append(f'<article class="finding"><div class="fhead"><span class="num">{k:02d}</span><div><h3>{esc(f.get("title", ""))}</h3>'
                     f'<p>{esc(f.get("detail", ""))}</p>{ev}</div></div>'
                     + (f'<div class="stats">{stat}</div>' if stat else "") + "".join(fig(x) for x in keys) + "</article>")
    findings = "".join(parts)
    if not c:
        findings = '<p class="muted">There is no conclusions.json yet: write the conclusions following step 6 of SKILL.md, then run report.py again.</p>'

    # At a glance: what guests praise / complain about (HTML bars)
    def bars(rows, cls):
        if not rows:
            return '<p class="muted">Not enough data</p>'
        top = max(r[1] for r in rows) or 1
        return "".join(f'<div class="bar{" few" if few else ""}"><span class="bl" title="{esc(lbl)}">{esc(re.split("[（(]", lbl)[0])}</span><span class="bt"><i class="{cls}" '
                       f'style="width:{max(4, v / top * 100):.0f}%"></i></span><span class="bv">{esc(txt)}</span></div>'
                       for lbl, v, txt, few in rows)

    pt = res["tables"].get("What guests praise")
    praise_rows = []
    if pt is not None and len(pt):
        for _, r in pt[pt["Reviews praising it"] > 0].iterrows():
            praise_rows.append((r["Aspect"], float(r["Share"]), f"{r['Share']:.0%}", r["Reviews praising it"] < A.MIN_THEME_MENTIONS))
    n_good = pt.attrs.get("n", 0) if pt is not None else 0
    comp_rows = [(x["label"], x["value"], f'{x["value"]}', x["value"] < A.MIN_THEME_MENTIONS) for x in d["complaints"]]
    pool = res["reviews"]  # only normally collected reviews
    n_neg = d["neg_total"]
    n_small = int(((pool["Complaint themes"].fillna("") != "") & (pool["Low score"] != True)).sum())  # noqa: E712
    n_ok = d["all_total"] - n_neg - n_small
    plats = sorted([x for x in d["platforms"] if x["count"]], key=lambda x: -x["count"])
    donuts = ('<div class="donuts"><section><h3>Where the reviews come from</h3><p class="cap">Reviews published on each platform</p>'
              + donut([(x["label"], x["count"], PALETTE[i % len(PALETTE)]) for i, x in enumerate(plats)],
                      f'{sum(x["count"] for x in plats)}', "reviews")
              + f'</section><section><h3>How many reviews complain</h3><p class="cap">The {d["all_total"]} reviews analysed</p>'
              + donut([("No complaint", n_ok, "#9fc3ee"), ("Complaint inside a good review", n_small, "#f3a17c"), ("Low score (below 6)", n_neg, "#eb6834")],
                      f'{n_ok / d["all_total"]:.0%}' if d["all_total"] else "—", "no complaint")
              + '</section></div>')
    glance = donuts + (f'<div class="glance"><section><h3>What guests praise</h3><p class="cap">Share of the {n_good} reviews analysed that praise it</p>'
              f'{bars(praise_rows, "pos")}</section><section><h3>What guests complain about</h3>'
              f'<p class="cap">Number of the {d["all_total"]} reviews analysed (including complaints inside good reviews) that complain about it</p>{bars(comp_rows, "neg")}'
              f'</section></div><p class="glance-note">Light colour = fewer than {A.MIN_THEME_MENTIONS} reviews, an isolated case: verify first (except anything about guest safety, privacy or being forced to leave: check at once however few).'
              f'One review can complain about several things, so the right-hand counts add up to more than the number of reviews with complaints.'
              f'Themes are detected by keywords, so a few are wrong.</p>')

    kpis = "".join(f'<div class="kpi"><span>{esc(k["label"])}</span><b>{esc(k["value"])}</b>'
                   f'<small>{esc(k["note"])}</small></div>' for k in d["kpis"])
    headline = c.get("headline", "")
    first = re.split(r"(?<=[.。])\s+", headline, maxsplit=1)[0]
    rest = headline[len(first):].strip()
    # The cover holds one short conclusion + a few "watch these" number blocks; without watch, the whole headline is used
    watch = "".join(f'<div class="watch"><span class="wl">{esc(w.get("label", ""))}</span><b>{esc(w.get("value", ""))}</b>'
                    + (f'<span class="ws">reviews complaining about it average {esc(w["score"])}</span>' if w.get("score") else "")
                    + (f'<small>{esc(w["note"])}</small>' if w.get("note") else "") + "</div>"
                    for w in c.get("watch", []))
    avg_all = res["reviews"]["score10"].mean()

    ov = d["overall"]
    how = dict(zip(ov["Item"], ov["Platforms used and how"]))
    scope = esc(dict(zip(ov["Item"], ov["Combined result"])).get("Platform scope", ""))

    hidden = [(p, f.get("hidden_reviews"), f.get("total_reviews")) for p, f in res["summary"].get("platforms", {}).items()
              if f.get("hidden_reviews")]
    # how many reviews collected per platform, and the date range
    rv = res["reviews"]
    scope_rows = []
    for p in sorted(rv["platform"].dropna().unique()):
        part = rv[rv["platform"] == p]
        pub = res["platforms"].published(p)
        scope_rows.append({"Platform": pf_disp(p), "Published": int(pub) if pub else "", "Collected": len(part),
                           "Date range": f"{part['date'].min().date()} to {part['date'].max().date()}" if part["date"].notna().any() else "",
                           "Languages": ", ".join(LANG_ZH.get(x, x) for x in sorted(part["review_language"].astype(str).unique()))})
    scope_html = table_html(pd.DataFrame(scope_rows), ["Platform", "Published", "Collected", "Date range", "Languages"])
    hidden_note = "".join(f"<li>{esc(pf_disp(p))}"
                          + (f"published {int(t)} reviews, {int(h)} fewer ({h / t:.0%}) can be seen on the page" if t else f"published count and visible count differ by {int(h)}")
                          + "; they may be hidden, folded, or score-only, and are not in the analysis.</li>" for p, h, t in hidden)
    fliggy_note = ("<li>Fliggy reviews have no overall score; a review's score is the average of its 4 sub-score stars.</li>"
                   if (res["reviews"]["platform"] == "Fliggy").any() else "")
    # platforms not fully collected (e.g. rate-limited while paging) have a scrape_note in summary.json
    hidden_note += "".join(f"<li>{esc(pf_disp(p))}not fully collected: {esc(f['scrape_note'])}.</li>"
                           for p, f in res["summary"].get("platforms", {}).items() if f.get("scrape_note"))

    page = f"""<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{hotel} review report</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>
/* Layout: dark cover with conclusion and key numbers → at a glance (praise/complaints side by side, all reviews) → what to do first (with evidence counts) → detailed findings (each with a chart) → appendix (data scope) */
:root {{
  --bg: #f4f5f7; --paper: #ffffff; --ink: #161a20; --ink-2: #48505c; --muted: #6f7782;
  --rule: #e1e4e9; --pos: #2a78d6; --neg: #eb6834; --track: #edf0f3; --accent: #1d5fb0;
  --band: #12304f; --band-ink: #ffffff; --band-2: #b9cce2; --band-card: rgba(255,255,255,.08);
  --sans: "Noto Sans SC", "PingFang SC", "Microsoft YaHei", system-ui, sans-serif;
  --serif: "Noto Serif SC", "Songti SC", "SimSun", serif;
}}
@media (prefers-color-scheme: dark) {{ :root:not([data-theme="light"]) {{
  --bg: #111317; --paper: #1a1d22; --ink: #eef0f3; --ink-2: #c4c9d1; --muted: #98a0ab;
  --rule: #2b3037; --pos: #3987e5; --neg: #d95926; --track: #262a31; --accent: #6aa5ec;
  --band: #0c1f35; --band-ink: #ffffff; --band-2: #a9bdd6; --band-card: rgba(255,255,255,.06); color-scheme: dark; }} }}
:root[data-theme="dark"] {{
  --bg: #111317; --paper: #1a1d22; --ink: #eef0f3; --ink-2: #c4c9d1; --muted: #98a0ab;
  --rule: #2b3037; --pos: #3987e5; --neg: #d95926; --track: #262a31; --accent: #6aa5ec;
  --band: #0c1f35; --band-ink: #ffffff; --band-2: #a9bdd6; --band-card: rgba(255,255,255,.06); color-scheme: dark; }}
* {{ box-sizing: border-box; }}
body {{ background: var(--bg); color: var(--ink); font: 15px/1.75 var(--sans); margin: 0; }}
.in {{ max-width: 980px; margin: 0 auto; padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px)); }}
p {{ margin: 0; }}
.muted {{ color: var(--muted); }}
.cover {{ background: var(--band); color: var(--band-ink); padding-block: 40px 32px; }}
.eyebrow {{ color: var(--band-2); font-size: 14px; }}
.hotel {{ font-size: 15px; font-weight: 500; color: var(--band-2); margin-top: 4px; }}
.cover h1 {{ font: 700 clamp(26px, 4.4vw, 40px)/1.3 var(--serif); margin: 12px 0 10px; text-wrap: balance; }}
.cover .sub {{ font-size: 17px; color: var(--band-2); max-width: 46em; }}
.wtitle {{ margin-top: 28px; font-size: 18px; font-weight: 700; display: flex; flex-wrap: wrap; align-items: baseline; gap: 4px 10px; }}
.wtitle span {{ font-size: 12.5px; font-weight: 400; color: var(--band-2); }}
.watchrow {{ display: grid; grid-template-columns: repeat(var(--n, 3), minmax(0, 1fr)); gap: 12px; margin-top: 10px; }}
.watch {{ display: flex; flex-direction: column; gap: 2px; background: var(--band-card); border-radius: 10px; padding: 14px 16px; border-top: 3px solid #f0a37e; }}
.watch .wl {{ font-size: 13px; color: var(--band-2); }}
.watch b {{ font-size: 30px; line-height: 1.25; font-variant-numeric: tabular-nums; }}
.watch .ws {{ font-size: 12px; }}
.watch small {{ font-size: 12px; color: var(--band-2); }}
.aline {{ display: flex; flex-wrap: wrap; align-items: baseline; gap: 4px 10px; }}
.actions details {{ margin-top: 4px; }}
.adetail {{ margin: 6px 0 0; padding-left: 18px; font-size: 14px; color: var(--ink-2); }}
.adetail li {{ margin: 2px 0; }}
.actions details p {{ margin-top: 4px; font-size: 13.5px; color: var(--ink-2); }}
.kpis {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-top: 28px; }}
.kpi {{ background: var(--band-card); border-radius: 10px; padding: 14px 16px; display: flex; flex-direction: column; gap: 2px; }}
.kpi span {{ font-size: 13px; color: var(--band-2); }}
.kpi b {{ font-size: 30px; line-height: 1.2; font-variant-numeric: tabular-nums; }}
.kpi small {{ font-size: 12px; color: var(--band-2); }}
main {{ padding-block: 8px 64px; }}
h2 {{ font: 700 22px/1.4 var(--serif); margin: 48px 0 16px; display: flex; align-items: baseline; gap: 10px; }}
h2 small {{ font: 400 13px var(--sans); color: var(--muted); }}
h3 {{ font-size: 17px; margin: 0; }}
.donuts {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }}
.donuts section {{ background: var(--paper); border: 1px solid var(--rule); border-radius: 12px; padding: 18px 20px; }}
.donut {{ display: flex; align-items: center; gap: 18px; flex-wrap: wrap; }}
.donut svg {{ flex: none; }}
.donut .dc {{ font: 700 26px var(--sans); fill: var(--ink); }}
.donut .ds {{ font: 400 12px var(--sans); fill: var(--muted); }}
.donut ul {{ list-style: none; margin: 0; padding: 0; display: grid; gap: 6px; font-size: 13.5px; flex: 1; min-width: 170px; }}
.donut li {{ display: grid; grid-template-columns: 12px 1fr auto 3em; align-items: center; gap: 8px; }}
.donut li i {{ width: 12px; height: 12px; border-radius: 3px; }}
.donut li b {{ font-weight: 500; font-variant-numeric: tabular-nums; }}
.donut li em {{ font-style: normal; color: var(--muted); text-align: right; font-variant-numeric: tabular-nums; }}
.stats {{ display: flex; flex-wrap: wrap; gap: 12px; margin-left: 52px; }}
.stat {{ background: var(--paper); border: 1px solid var(--rule); border-radius: 12px; padding: 12px 18px; display: flex; flex-direction: column; min-width: 140px; }}
.stat b {{ font-size: 26px; line-height: 1.25; color: var(--ink); font-variant-numeric: tabular-nums; }}
.stat span {{ font-size: 13px; color: var(--muted); }}
.glance {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
.glance section {{ background: var(--paper); border: 1px solid var(--rule); border-radius: 12px; padding: 18px 20px; }}
.cap {{ font-size: 12.5px; color: var(--muted); margin: 2px 0 12px; }}
.bar {{ display: grid; grid-template-columns: 9.5em 1fr 3.5em; align-items: center; gap: 10px; padding: 5px 0; font-size: 14px; }}
.bl {{ color: var(--ink); line-height: 1.3; overflow-wrap: anywhere; }}
.bt {{ background: var(--track); border-radius: 4px; height: 12px; }}
.bt i {{ display: block; height: 12px; border-radius: 4px; }}
.bt i.pos {{ background: var(--pos); }} .bt i.neg {{ background: var(--neg); }}
.bv {{ text-align: right; color: var(--ink-2); font-variant-numeric: tabular-nums; }}
.actions {{ list-style: none; padding: 0; margin: 0; background: var(--paper); border: 1px solid var(--rule); border-radius: 12px; }}
.actions > li {{ display: grid; grid-template-columns: 72px 1fr; gap: 14px; padding: 16px 20px; }}
.actions > li + li {{ border-top: 1px solid var(--rule); }}
.badge {{ justify-self: start; align-self: start; font-size: 12px; font-weight: 700; color: #fff; background: var(--accent); border-radius: 999px; padding: 2px 10px; margin-top: 2px; }}
.badge.later {{ background: transparent; color: var(--accent); border: 1px solid var(--accent); }}
.actions strong {{ font-size: 15.5px; }}
.basis {{ display: inline-block; font-size: 12.5px; color: var(--ink-2); background: var(--track); border-radius: 4px; padding: 1px 8px; }}
.bar.few .bt i {{ opacity: .38; }}
.bar.few .bl, .bar.few .bv {{ color: var(--muted); }}
.glance-note {{ font-size: 12.5px; color: var(--muted); margin-top: 8px; }}
.actions p {{ color: var(--muted); font-size: 13.5px; margin-top: 4px; }}
.findings {{ display: grid; gap: 28px; }}
.finding {{ display: grid; gap: 12px; }}
.fhead {{ display: grid; grid-template-columns: 44px 1fr; gap: 8px; }}
.num {{ font: 700 20px/1.6 var(--serif); color: var(--accent); }}
.finding h3 {{ font: 700 18px/1.6 var(--serif); }}
.finding p {{ color: var(--ink-2); max-width: 60em; margin-top: 2px; }}
.chart {{ margin: 0; background: #ffffff; border: 1px solid var(--rule); border-radius: 12px; padding: 10px; }}
.chart img {{ display: block; width: 100%; height: auto; }}
.charts {{ display: grid; gap: 16px; }}
details {{ margin-top: 6px; font-size: 13px; color: var(--muted); }}
summary {{ cursor: pointer; }}
.evidence {{ margin: 6px 0 0; padding-left: 18px; color: var(--ink-2); font-size: 13.5px; }}
.appendix {{ margin-top: 56px; padding-top: 8px; border-top: 2px solid var(--rule); }}
.tablewrap {{ overflow-x: auto; background: var(--paper); border: 1px solid var(--rule); border-radius: 12px; padding: 6px 14px; }}
table {{ border-collapse: collapse; width: 100%; font-size: 13.5px; font-variant-numeric: tabular-nums; }}
th, td {{ text-align: left; padding: 8px 8px; border-bottom: 1px solid var(--rule); vertical-align: top; }}
tr:last-child td {{ border-bottom: 0; }}
th {{ color: var(--muted); font-weight: 500; }}
.negs {{ padding-left: 22px; display: grid; gap: 10px; margin: 0; }}
.negs li {{ background: var(--paper); border: 1px solid var(--rule); border-radius: 10px; padding: 12px 14px; }}
.neg-meta {{ font-size: 12.5px; color: var(--muted); margin-bottom: 4px; }}
.negs p {{ color: var(--ink-2); font-size: 14px; }}
.notes {{ padding-left: 18px; }}
.notes li {{ color: var(--ink-2); font-size: 13.5px; }}
@media (max-width: 760px) {{
  .kpis {{ grid-template-columns: 1fr 1fr; }}
  .glance, .donuts {{ grid-template-columns: 1fr; }}
  .watchrow {{ grid-template-columns: 1fr 1fr; }}
  .stats {{ margin-left: 0; }}
  .actions > li {{ grid-template-columns: 1fr; gap: 6px; }}
  .bar {{ grid-template-columns: 7.5em 1fr 3.2em; }}
  .fhead {{ grid-template-columns: 34px 1fr; }}
}}
</style>
<div class="cover"><div class="in">
  <div class="eyebrow">{hotel} · OTA review summary · data as of {esc(d['today'])}</div>
  <h1>{esc(first)}</h1>
  {'' if watch else f'<p class="sub">{esc(rest)}</p>'}
  <div class="kpis">{kpis}</div>
  {(f'<div class="wtitle">Watch these {len(c.get("watch", []))}<span>For comparison: the {len(res["reviews"])} reviews analysed average {avg_all:.1f} (calculated from the collected reviews, a different method from the platform rating above); one review can count in several items</span></div>'
    f'<div class="watchrow" style="--n:{len(c.get("watch", []))}">{watch}</div>') if watch else ''}
</div></div>

<main class="in">
<h2>At a glance</h2>
{glance}

<h2>What to do first</h2>
<ol class="actions">{actions or '<li><span class="badge">—</span><div><p>No actions in conclusions.json.</p></div></li>'}</ol>

<h2>Detailed findings<small>each conclusion has its chart below; open "Evidence" for quotes and numbers</small></h2>
<div class="findings">{findings}</div>

<div class="appendix">
<h2>Appendix · data scope and notes</h2>
{scope_html}
<ul class="notes">
  <li>{scope}</li>
  <li>Reviews come from the public pages of each platform; reviews not visible on the page are not analysed.</li>
  {hidden_note}
  <li>Dates are when the guest wrote the review, not the stay date, usually a few days to weeks later.</li>
  <li>Low score: the rule of this report is a guest score below 6 on a 10-point scale; the denominator of the low-score share is the reviews analysed. The platforms' own negative counts (e.g. Ctrip "Negative (7)") are listed separately in the Excel file and may differ.</li>{fliggy_note}
  <li>Praise and complaints are detected by keywords and only the clause where the guest praises or complains is counted, so a few are misjudged or missed.</li>
</ul>
</div>
</main>
"""
    out.write_text(page, encoding="utf-8")
    return d


def add_conclusions_sheet(excel, conclusions, images=()):
    """Put the conclusions on the first Excel sheet and the charts on the second."""
    from openpyxl.drawing.image import Image as XLImage
    wb = load_workbook(excel)
    if images:
        if "Charts" in wb.sheetnames:
            del wb["Charts"]
        ws = wb.create_sheet("Charts", 0)
        row = 1
        for p, name in images:
            img = XLImage(str(p))
            img.width, img.height = img.width * 0.42, img.height * 0.42
            ws.add_image(img, f"A{row}")
            row += int(img.height / 20) + 2
    if not conclusions:
        wb.save(excel)
        return
    if "Conclusions" in wb.sheetnames:
        del wb["Conclusions"]
    ws = wb.create_sheet("Conclusions", 0)
    ws.append(["Headline", conclusions.get("headline", "")])
    ws.append([])
    ws.append(["What to do first", "Action", "Why (data evidence)"])
    for a in conclusions.get("actions", []):
        ws.append([a.get("when", ""), a.get("action", ""), a.get("why", "")])
    ws.append([])
    ws.append(["Conclusion", "Detail", "Evidence"])
    for f in conclusions.get("findings", []):
        ws.append([f.get("title", ""), f.get("detail", ""), "\n".join(f.get("evidence", []))])
    from openpyxl.styles import Alignment, Font
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.value in ("Headline", "What to do first", "Conclusion"):
                cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 70
    wb.save(excel)


def main(folder, today=None):
    folder = Path(folder)
    res = A.main(folder, today=today)
    path = folder / "conclusions.json"
    conclusions = json.loads(path.read_text(encoding="utf-8")) if path.exists() else None
    out = folder / f"review_report_{A.BAD_FILENAME.sub('_', str(res['hotel']))}.html"
    images = make_charts(res, folder)
    render(res, conclusions, out, images)
    add_conclusions_sheet(res["excel"], conclusions, images)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
