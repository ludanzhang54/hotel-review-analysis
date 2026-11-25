"""
把 analyze.py 的结果做成一份带图表的点评报告（HTML），并把"结论与建议"加到 Excel 第一页。

用法：python report.py <数据文件夹>
  - 先运行 analyze.py 的全部计算（会重新生成 Excel）；
  - 读同一文件夹里的 conclusions.json（结论与建议，由 AI 按 SKILL.md 第 6 步写）；
  - 生成 点评报告_<酒店名>.html：结论和建议在最前面，下面是图表和数据说明。

图表是 PNG（charts.py 画的），同时存在"图表"文件夹里、嵌在 HTML 里、放进 Excel 的"图表"页，打开不用联网。
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



try:  # Trip.com 英文页把中文点评显示成繁体，报告里转回简体（装了 opencc-python-reimplemented 才转）
    from opencc import OpenCC
    _t2s = OpenCC("t2s").convert
except Exception:  # noqa: BLE001
    _t2s = None

LANG_ZH = {"zh": "中文", "en": "英文", "ko": "韩文", "ja": "日文", "zh-CN": "简体中文", "zh-HK": "繁体中文（香港）",
           "zh-TW": "繁体中文（台湾）", "th": "泰文", "fr": "法文", "de": "德文", "es": "西班牙文", "ru": "俄文", "vi": "越南文"}


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
    """从计算结果里取出各张图要用的数字。"""
    t, reviews, pool, pf, today = res["tables"], res["reviews"], res["pool"], res["platforms"], res["today"]
    overall = t["综合汇总"]
    ov = dict(zip(overall["项目"], overall["合并结果"]))
    data = {}

    # 各平台评分（10 分制）和点评数
    rows = []
    for p in pf.rating_platforms():
        rows.append({"label": pf.disp(p), "value": pf.rating10(p), "count": int(pf.published(p) or 0)})
    data["platforms"] = rows


    # 客人在抱怨什么：全部点评里（含好评里顺带的抱怨）有多少条在抱怨这个主题。差评往往只有几条，只看差评样本太小
    # 所有比例和条数只用正常抓到的点评；补抓的低分点评（lowscore_*.json）只当原文证据，不进比例
    data["neg_total"] = int(len(reviews[reviews["差评"] == True]))  # noqa: E712
    data["all_total"] = int(len(reviews))
    data["complaint_reviews"] = int((reviews["命中主题"].fillna("") != "").sum())
    data["complaints"] = sorted(
        [{"label": th, "value": int(reviews["命中主题"].str.contains(re.escape(th)).sum())} for th in A.THEMES],
        key=lambda x: -x["value"])
    data["complaints"] = [x for x in data["complaints"] if x["value"] > 0]

    # 主题影响分
    imp = t["主题影响分"].dropna(subset=["影响分"])
    data["impact"] = [{"label": r["主题"], "value": num(r["影响分"]), "count": int(r["抱怨的点评条数"])}
                      for _, r in imp.iterrows()]

    # 分项、住客：取综合汇总里合并后的数
    data["subscores"] = [{"label": k.replace("分项：", "").replace("（10分制）", ""), "value": num(v)}
                         for k, v in ov.items() if str(k).startswith("分项：") and num(v) is not None]
    data["travellers"] = [{"label": k.replace("住客：", ""), "value": num(float(v) * 100, 1)}
                          for k, v in ov.items() if str(k).startswith("住客：") and num(v) is not None]

    # 平台话题（携程热门标签、Agoda 话题）
    topics = t["平台话题"]
    data["topics"] = []
    if len(topics) and "话题" in topics.columns:
        for _, r in topics.sort_values("提及次数", ascending=False).head(12).iterrows():
            data["topics"].append({"label": f"{r['话题']}（{pf.disp(r['平台'])}）", "value": int(r["提及次数"])})

    # 指标卡
    neg_n = re.search(r"共 (\d+) 条差评", str(ov.get("差评时间分布", "")))
    neg_recent = re.search(r"近1年 (\d+) 条", str(ov.get("差评时间分布", "")))
    cr, tot = data["complaint_reviews"], data["all_total"]
    # 指标卡：单位统一（分、条），比例统一用"本次分析的 N 条里占 X%"、取整，不混用不同分母
    nn = data["neg_total"]
    data["kpis"] = [
        {"label": "综合评分", "value": f"{float(ov['综合评分（10分制）']):.1f} 分" if num(ov.get("综合评分（10分制）")) else "—",
         "note": "10 分制，各平台公布的评分按点评数加权；各平台算法不同，只作参考，看变化要看同一平台自己的分"},
        {"label": "各平台公布点评", "value": f"{int(ov['平台公布点评总数'])} 条" if num(ov.get("平台公布点评总数")) else "—",
         "note": f"本次分析 {tot} 条" + "".join(
             f"；{pf.disp(p)}没抓全" for p, f in res["summary"].get("platforms", {}).items() if f.get("scrape_note"))},
        {"label": "提到不满的点评", "value": f"{cr} 条", "note": f"本次分析的 {tot} 条里占 {cr / tot:.0%}" if tot else ""},
        {"label": "差评（低于 6 分）", "value": f"{nn} 条",
         "note": (f"本次分析的 {tot} 条里占 {nn / tot:.0%}" if tot else "") + (f"；近一年 {int(neg_recent.group(1))} 条" if neg_recent else "")},
    ]
    data["overall"] = overall
    data["compare"] = t["各平台对比"]
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
            if isinstance(v, float) and ("占比" in c) and not pd.isna(v):
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
    """环形图（SVG，深浅色都能看）：parts = [(名称, 数值, 颜色)]，只用于加起来是 100% 的构成（一条点评只能算一类）。
    多选的比例（一条点评可以夸好几样）不能画成饼图，要用横条。"""
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
    legend = "".join(f'<li><i style="background:{col}"></i><span>{esc(lbl)}</span><b>{v} 条</b>'
                     f'<em>{v / total:.0%}</em></li>' for lbl, v, col in parts)
    return f'<div class="donut">{svg}<ul>{legend}</ul></div>'


def render(res, conclusions, out, images):
    d = build_data(res)
    pf_disp = res["platforms"].disp
    hotel = esc(res["hotel"])
    c = conclusions or {}

    actions = "".join(
        f'<li><span class="badge{" later" if a.get("when") not in ("立即", "先做") else ""}">{esc(a.get("when", ""))}</span>'
        f'<div><div class="aline"><strong>{esc(a.get("action", ""))}</strong>'
        + (f'<span class="basis">{esc(a["basis"])}</span>' if a.get("basis") else "") + "</div>"
        # 具体做什么直接写出来（分点），只有"为什么"收起来
        + ("<ul class='adetail'>" + "".join(f"<li>{esc(x)}</li>" for x in (a["detail"] if isinstance(a["detail"], list) else [a["detail"]]))
           + "</ul>" if a.get("detail") else "")
        + (f'<details><summary>为什么</summary><p class="muted">{esc(a["why"])}</p></details>' if a.get("why") else "")
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
        ev = ("<details><summary>依据（原文和数字）</summary><ul class='evidence'>"
              + "".join(f"<li>{esc(e)}</li>" for e in f.get("evidence", [])) + "</ul></details>") if f.get("evidence") else ""
        stat = "".join(f'<div class="stat"><b>{esc(x.get("value", ""))}</b><span>{esc(x.get("label", ""))}</span></div>'
                       for x in f.get("stats", []))
        parts.append(f'<article class="finding"><div class="fhead"><span class="num">{k:02d}</span><div><h3>{esc(f.get("title", ""))}</h3>'
                     f'<p>{esc(f.get("detail", ""))}</p>{ev}</div></div>'
                     + (f'<div class="stats">{stat}</div>' if stat else "") + "".join(fig(x) for x in keys) + "</article>")
    findings = "".join(parts)
    if not c:
        findings = '<p class="muted">还没有 conclusions.json：按 SKILL.md 第 6 步写好结论后重新运行 report.py。</p>'

    # 一眼看懂：客人夸什么 / 抱怨什么（HTML 横条）
    def bars(rows, cls):
        if not rows:
            return '<p class="muted">数据不足</p>'
        top = max(r[1] for r in rows) or 1
        return "".join(f'<div class="bar{" few" if few else ""}"><span class="bl" title="{esc(lbl)}">{esc(re.split("[（(]", lbl)[0])}</span><span class="bt"><i class="{cls}" '
                       f'style="width:{max(4, v / top * 100):.0f}%"></i></span><span class="bv">{esc(txt)}</span></div>'
                       for lbl, v, txt, few in rows)

    pt = res["tables"].get("好评在夸什么")
    praise_rows = []
    if pt is not None and len(pt):
        for _, r in pt[pt["夸它的点评条数"] > 0].iterrows():
            praise_rows.append((r["方面"], float(r["占比"]), f"{r['占比']:.0%}", r["夸它的点评条数"] < A.MIN_THEME_MENTIONS))
    n_good = pt.attrs.get("n", 0) if pt is not None else 0
    comp_rows = [(x["label"], x["value"], f'{x["value"]} 条', x["value"] < A.MIN_THEME_MENTIONS) for x in d["complaints"]]
    pool = res["reviews"]  # 只用正常抓到的点评
    n_neg = d["neg_total"]
    n_small = int(((pool["命中主题"].fillna("") != "") & (pool["差评"] != True)).sum())  # noqa: E712
    n_ok = d["all_total"] - n_neg - n_small
    plats = sorted([x for x in d["platforms"] if x["count"]], key=lambda x: -x["count"])
    donuts = ('<div class="donuts"><section><h3>点评来自哪些平台</h3><p class="cap">各平台公布的点评数</p>'
              + donut([(x["label"], x["count"], PALETTE[i % len(PALETTE)]) for i, x in enumerate(plats)],
                      f'{sum(x["count"] for x in plats)}', "条点评")
              + f'</section><section><h3>点评里有多少不满</h3><p class="cap">本次分析的 {d["all_total"]} 条点评</p>'
              + donut([("没有抱怨", n_ok, "#9fc3ee"), ("好评里顺带抱怨", n_small, "#f3a17c"), ("差评（低于 6 分）", n_neg, "#eb6834")],
                      f'{n_ok / d["all_total"]:.0%}' if d["all_total"] else "—", "没有抱怨")
              + '</section></div>')
    glance = donuts + (f'<div class="glance"><section><h3>客人夸什么</h3><p class="cap">本次分析的 {n_good} 条点评里，夸这一点的比例</p>'
              f'{bars(praise_rows, "pos")}</section><section><h3>客人抱怨什么</h3>'
              f'<p class="cap">本次分析的 {d["all_total"]} 条点评里（含好评里顺带的抱怨），抱怨这一点的条数</p>{bars(comp_rows, "neg")}'
              f'</section></div><p class="glance-note">浅色 = 不到 {A.MIN_THEME_MENTIONS} 条，属于个别情况，先核实（涉及客人安全、隐私、被迫离店的除外，再少也要马上查）。'
              f'一条点评可以同时抱怨好几样，所以右栏条数加起来比提到不满的点评多。'
              f'主题按关键词自动识别，会有少量误判。</p>')

    kpis = "".join(f'<div class="kpi"><span>{esc(k["label"])}</span><b>{esc(k["value"])}</b>'
                   f'<small>{esc(k["note"])}</small></div>' for k in d["kpis"])
    headline = c.get("headline", "")
    first, _, rest = headline.partition("。")
    # 封面只放一句短结论 + "要盯的"几个数字块；没写 watch 时退回整句结论
    watch = "".join(f'<div class="watch"><span class="wl">{esc(w.get("label", ""))}</span><b>{esc(w.get("value", ""))}</b>'
                    + (f'<span class="ws">抱怨它的点评平均 {esc(w["score"])} 分</span>' if w.get("score") else "")
                    + (f'<small>{esc(w["note"])}</small>' if w.get("note") else "") + "</div>"
                    for w in c.get("watch", []))
    avg_all = res["reviews"]["score10"].mean()

    ov = d["overall"]
    how = dict(zip(ov["项目"], ov["用了哪些平台、怎么算"]))
    scope = esc(dict(zip(ov["项目"], ov["合并结果"])).get("平台范围", ""))

    hidden = [(p, f.get("hidden_reviews"), f.get("total_reviews")) for p, f in res["summary"].get("platforms", {}).items()
              if f.get("hidden_reviews")]
    # 每个平台抓了多少、时间范围
    rv = res["reviews"]
    scope_rows = []
    for p in sorted(rv["platform"].dropna().unique()):
        part = rv[rv["platform"] == p]
        pub = res["platforms"].published(p)
        scope_rows.append({"平台": pf_disp(p), "平台公布条数": int(pub) if pub else "", "抓到条数": len(part),
                           "时间范围": f"{part['date'].min().date()} 至 {part['date'].max().date()}" if part["date"].notna().any() else "",
                           "语言": "、".join(LANG_ZH.get(x, x) for x in sorted(part["review_language"].astype(str).unique()))})
    scope_html = table_html(pd.DataFrame(scope_rows), ["平台", "平台公布条数", "抓到条数", "时间范围", "语言"])
    hidden_note = "".join(f"<li>{esc(pf_disp(p))}"
                          + (f"公布 {int(t)} 条点评，页面上能看到的少 {int(h)} 条（{h / t:.0%}）" if t else f"公布数和能看到的条数差 {int(h)} 条")
                          + "：可能是被隐藏、折叠，或只打分没写字，这部分不在分析内。</li>" for p, h, t in hidden)
    fliggy_note = ("<li>飞猪每条点评没有总分，单条分数按 4 个分项星级平均。</li>"
                   if (res["reviews"]["platform"] == "飞猪").any() else "")
    # 没抓全的平台（例如翻页被限流），summary.json 里写 scrape_note
    hidden_note += "".join(f"<li>{esc(pf_disp(p))}没抓全：{esc(f['scrape_note'])}。</li>"
                           for p, f in res["summary"].get("platforms", {}).items() if f.get("scrape_note"))

    page = f"""<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{hotel} 点评报告</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>
/* 版式：深色封面写结论和关键数字 → 一眼看懂（夸/抱怨并排，全部点评）→ 先做什么（带依据条数）→ 详细发现（每条配图）→ 附录（数据范围） */
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
  <div class="eyebrow">{hotel} · OTA 点评汇总 · 数据截至 {esc(d['today'])}</div>
  <h1>{esc(first)}</h1>
  {'' if watch else f'<p class="sub">{esc(rest)}</p>'}
  <div class="kpis">{kpis}</div>
  {(f'<div class="wtitle">最该盯的 {len(c.get("watch", []))} 件事<span>对比：本次分析的 {len(res["reviews"])} 条点评平均 {avg_all:.1f} 分（自己按抓到的点评算，和上面平台公布的综合评分算法不同）；一条点评可以同时算进好几件事</span></div>'
    f'<div class="watchrow" style="--n:{len(c.get("watch", []))}">{watch}</div>') if watch else ''}
</div></div>

<main class="in">
<h2>一眼看懂</h2>
{glance}

<h2>先做什么</h2>
<ol class="actions">{actions or '<li><span class="badge">—</span><div><p>conclusions.json 里没有写建议。</p></div></li>'}</ol>

<h2>详细发现<small>每条结论下面是对应的图；点"依据"看原文和数字</small></h2>
<div class="findings">{findings}</div>

<div class="appendix">
<h2>附录 · 数据范围和说明</h2>
{scope_html}
<ul class="notes">
  <li>{scope}</li>
  <li>点评来自各平台公开页面；页面上看不到的点评不在分析内。</li>
  {hidden_note}
  <li>日期是客人写点评的日期，不是入住日期，一般晚几天到几周。</li>
  <li>差评：本报告的规则，客人打分换算成 10 分制后低于 6 分；差评占比的分母是本次分析的点评。平台自己标的差评条数（如携程"差评 (7)"）另列在 Excel，两者不一定相同。</li>{fliggy_note}
  <li>夸什么、抱怨什么都是按关键词自动识别的，只算客人在夸或在抱怨的那句话，会有少量误判和漏判。</li>
</ul>
</div>
</main>
"""
    out.write_text(page, encoding="utf-8")
    return d


def add_conclusions_sheet(excel, conclusions, images=()):
    """把结论与建议放到 Excel 第一页，图表放到第二页。"""
    from openpyxl.drawing.image import Image as XLImage
    wb = load_workbook(excel)
    if images:
        if "图表" in wb.sheetnames:
            del wb["图表"]
        ws = wb.create_sheet("图表", 0)
        row = 1
        for p, name in images:
            img = XLImage(str(p))
            img.width, img.height = img.width * 0.42, img.height * 0.42
            ws.add_image(img, f"A{row}")
            row += int(img.height / 20) + 2
    if not conclusions:
        wb.save(excel)
        return
    if "结论与建议" in wb.sheetnames:
        del wb["结论与建议"]
    ws = wb.create_sheet("结论与建议", 0)
    ws.append(["一句话结论", conclusions.get("headline", "")])
    ws.append([])
    ws.append(["先做什么", "做法", "为什么（数据依据）"])
    for a in conclusions.get("actions", []):
        ws.append([a.get("when", ""), a.get("action", ""), a.get("why", "")])
    ws.append([])
    ws.append(["结论", "说明", "依据"])
    for f in conclusions.get("findings", []):
        ws.append([f.get("title", ""), f.get("detail", ""), "\n".join(f.get("evidence", []))])
    from openpyxl.styles import Alignment, Font
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.value in ("一句话结论", "先做什么", "结论"):
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
    out = folder / f"点评报告_{A.BAD_FILENAME.sub('_', str(res['hotel']))}.html"
    images = make_charts(res, folder)
    render(res, conclusions, out, images)
    add_conclusions_sheet(res["excel"], conclusions, images)
    print(f"已保存：{out}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
