"""
Charts (PNG) of the review report. Style: a one-line title above says what the chart shows, grey small text below gives the data source;
numbers are written on the colour blocks or bars, so the axes need not be read.

Usage: called by report.py as make_charts(res, folder); returns [(file name, title), ...].
"""
import re
import sys

import pandas as pd
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.colors import LinearSegmentedColormap  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import analyze as A  # noqa: E402

# CJK font (for Chinese review text, platform tags and room types): Microsoft YaHei on Windows, PingFang on Mac, Noto Sans CJK on Linux; whichever is available
_installed = {f.name for f in font_manager.fontManager.ttflist}
FONT = next((f for f in ["Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", "Noto Sans CJK SC", "Noto Sans CJK TC",
                         "Source Han Sans SC", "SimHei", "Arial Unicode MS"] if f in _installed), "sans-serif")
if FONT == "sans-serif":
    # Without a CJK font, Chinese text in the charts turns into boxes; better to stop than to hand over unreadable charts
    sys.exit("No CJK font found, so Chinese text in the charts would show as boxes. Install one first (e.g. Noto Sans CJK SC) and run again.")
plt.rcParams.update({"font.family": FONT, "axes.unicode_minus": False, "font.size": 11})

INK, INK2, MUTED, GRID = "#1f2328", "#4a505a", "#8a8f98", "#e6e8eb"
BLUE, ORANGE, GREEN, YELLOW, PINK = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"
RED = "#d64545"
BLUES = LinearSegmentedColormap.from_list("blues", ["#dce8f8", "#5b9be6", "#1d5fb0", "#123f7a"])
ORANGES = LinearSegmentedColormap.from_list("oranges", ["#fbe3d8", "#f5b08f", "#eb6834", "#c9581f"])
EMPTY = "#f0f1f3"
DARK = "#123f7a"
# Platform tag labels are shown as the platform wrote them


def _frame(fig, title, source):
    fig.text(0.012, 0.975, title, fontsize=15, color=INK2, va="top")
    if source:
        fig.text(0.012, 0.018, source, fontsize=9, color=MUTED, va="bottom", wrap=True)


def _ink_on(rgb):
    r, g, b = rgb[:3]
    return "white" if (0.299 * r + 0.587 * g + 0.114 * b) < 0.6 else INK


def heat_table(path, title, source, rows, cols, values, vmin, vmax, fmt, cmap=BLUES, bold_rows=(), col_notes=None,
               color_values=None):
    """Colour-block table: rows x cols; a value of None is shown as a dash."""
    nr, nc = len(rows), len(cols)
    H = 1.9 + 0.48 * nr
    fig = plt.figure(figsize=(max(8, 2.2 + 1.7 * nc), H), dpi=200)
    # Top and bottom margins are fixed in inches (title about 1.15 in, footnote about 0.75 in), so with many rows the table is not separated from the title by a big gap
    ax = fig.add_axes([0.02, 0.75 / H, 0.96, (H - 0.75 - 1.15) / H])
    ax.set_xlim(-2.6, nc)
    ax.set_ylim(nr, -0.2)
    ax.axis("off")
    for j, c in enumerate(cols):
        ax.text(j + 0.5, -0.38, str(c).replace(" (", "\n(") if len(str(c)) > 12 else c, ha="center", va="bottom", fontsize=11 if nc > 5 else 11.5, color=INK)
        if col_notes and col_notes[j]:
            ax.text(j + 0.5, -0.12, col_notes[j], ha="center", va="bottom", fontsize=9, color=MUTED)
    for i, r in enumerate(rows):
        bold = r in bold_rows
        ax.text(-0.1, i + 0.5, r, ha="right", va="center", fontsize=11.5, color=INK, fontweight="bold" if bold else "normal")
        for j in range(nc):
            v = values[i][j]
            if v is None:
                color, text, tc = EMPTY, "—", MUTED
            else:
                if color_values is not None:
                    cv = color_values[i][j]
                    hi = max(x for r in color_values for x in r if x is not None)
                    t = min(max(cv / hi, 0.15), 1) if hi else 0.5
                else:
                    t = min(max((v - vmin) / (vmax - vmin), 0), 1)
                color = cmap(t)
                text, tc = fmt(v), _ink_on(cmap(t))
            ax.add_patch(plt.Rectangle((j + 0.04, i + 0.06), 0.92, 0.88, color=color, lw=0))
            ax.text(j + 0.5, i + 0.5, text, ha="center", va="center", fontsize=11.5, color=tc,
                    fontweight="bold" if bold else "normal")
    _frame(fig, title, source)
    fig.savefig(path, facecolor="white")
    plt.close(fig)


def hbars(path, title, source, labels, values, fmt, color=BLUE, xmax=None, notes=None):
    n = len(labels)
    fig = plt.figure(figsize=(10, 1.5 + 0.42 * n), dpi=200)
    width = max(sum(0.016 if ord(ch) > 255 else 0.0085 for ch in str(x)) for x in labels)
    left = min(0.5, max(0.2, 0.04 + width))
    ax = fig.add_axes([left, 0.14, 0.93 - left, 0.70])
    y = range(n)
    ax.barh(list(y), values, color=color, height=0.62)
    ax.set_yticks(list(y), labels, fontsize=11.5, color=INK)
    ax.invert_yaxis()
    lim = xmax or (max(values) * 1.25 if values and max(values) > 0 else 1)
    ax.set_xlim(min(0, min(values) * 1.3 if values else 0), lim)
    for i, v in enumerate(values):
        ax.text(v + (lim * 0.01 if v >= 0 else -lim * 0.01), i, fmt(v, i), va="center",
                ha="left" if v >= 0 else "right", fontsize=11, color=INK2)
    for s in ["top", "right", "bottom"]:
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(GRID)
    ax.tick_params(axis="x", colors=MUTED, length=0)
    ax.tick_params(axis="y", length=0)
    ax.set_xticks([])
    _frame(fig, title, source)
    fig.savefig(path, facecolor="white")
    plt.close(fig)


MIN_NEG_FOR_MONTHS = 20  # with fewer than 20 low-score reviews the "low-score months" chart is not drawn


def make_charts(res, folder):
    t, reviews, pool, pf = res["tables"], res["reviews"], res["pool"], res["platforms"]
    ov = t["Summary"]
    item = dict(zip(ov["Item"], ov["Combined result"]))
    folder = Path(folder) / "charts"
    folder.mkdir(exist_ok=True)
    out = []
    date = str(res["today"].date())

    # 1. Rating overview: overall rating + sub-scores, per platform and merged
    plats = pf.rating_platforms()
    subs = t["Sub-scores"]
    groups = [g for g, _ in A.SUBSCORE_GROUPS]
    rows, values = ["Overall rating"], [[pf.rating10(p) for p in plats] + [A.wavg([(pf.rating10(p), pf.weight(p)[0]) for p in plats])]]
    for g in groups:
        part = subs[subs["Merged sub-score"] == g] if len(subs) else subs
        if part is None or part.empty:
            continue
        line = []
        for p in plats:
            m = part[part["Platform"].map(pf.same) == pf.same(p)]
            line.append(float(m["On 10-point scale"].iloc[0]) if len(m) else None)
        merged = item.get(next((k for k in item if str(k).startswith(f"Sub-score: {g}")), None))
        rows.append(g)
        values.append(line + [float(merged) if merged is not None else None])
    env = subs[(subs["Merged sub-score"].isna()) & subs["Sub-score"].astype(str).str.lower().isin(["环境", "location"])
               & subs["Platform"].isin(["Ctrip", "Trip.com"])] if len(subs) else subs
    if len(env):
        v = float(env["On 10-point scale"].iloc[0])
        rows.append("Environment (Ctrip)")
        values.append([v if pf.same(p) in ("Ctrip", "Trip.com") else None for p in plats] + [v])
    cols = [pf.disp(p) for p in plats] + ["Combined"]
    notes = [f"{int(pf.published(p) or 0)} reviews" for p in plats] + [f"{int(pf.all_published())} reviews"]
    vals = [v for r in values for v in r if v is not None]
    p1 = folder / "1_rating_overview.png"
    heat_table(p1, "Platform ratings and sub-scores (10-point scale, darker = higher)",
               f"Source: ratings published on each platform's hotel page, read on {date}. Combined = weighted by each platform's published review count."
               "\nSub-score names differ between platforms and are put into the same group (e.g. Cleanliness and Hygiene are both cleanliness); '—' = that platform has no such item. 'Environment' is a Ctrip sub-score and is not the same as 'Location'.",
               rows, cols, values, min(vals + [8]) - 0.3, 10, lambda v: f"{v:.1f}", bold_rows=("Overall rating",), col_notes=notes)
    out.append((p1, "Rating overview"))

    # 2. What guests complain about: all reviews (including complaints inside praise reviews), not only low-score ones, because low-score reviews are often only a few
    # Shares and counts use only normally collected reviews; topped-up low-score reviews are not in the chart (evidence quotes only)
    n_all = len(reviews)
    counts = []
    for th in A.THEMES:
        hit = reviews["Complaint themes"].str.contains(re.escape(th))
        if hit.sum():
            counts.append((th, int(hit.sum()), int((hit & (reviews["Low score"] == True)).sum())))  # noqa: E712
    counts.sort(key=lambda x: -x[1])
    if counts:
        p2 = folder / "2_complaints.png"
        hbars(p2, f"Among the {n_all} analysed reviews, how many complain about each problem",
              f"Source: all collected reviews (including complaints inside praise reviews); the clause where the guest complains is found by keywords automatically; one review can mention several problems."
              f"\nFewer than {A.MIN_THEME_MENTIONS} reviews are isolated cases and should not be read as a general problem.",
              [c[0] for c in counts], [c[1] for c in counts],
              lambda v, i: f"{v} reviews" + (f" ({counts[i][2]} of them low-score)" if counts[i][2] else ""), color=ORANGE,
              xmax=counts[0][1] * 1.7)
        out.append((p2, "What guests complain about"))

    # 3-5 time charts use only platforms and languages whose reviews were fully collected; on others, few old reviews only means they were not collected
    trend, trend_used = A.complete_sample(reviews, res["summary"])
    trend_note = "Platforms used: " + (", ".join(dict.fromkeys(pf.disp(x) for x in trend_used)) if trend_used else "none")

    # 3. When the complained-about problems appear (by half-year)
    allr = trend[trend["date"].notna()].copy()
    if len(allr):
        allr["period"] = allr["date"].dt.year.astype(str) + allr["date"].dt.month.map(lambda m: " H1" if m <= 6 else " H2")
        periods = sorted(allr["period"].unique())
        th_rows, th_vals = [], []
        for th in A.THEMES:
            line = [int(allr.loc[allr["period"] == p, "Complaint themes"].str.contains(re.escape(th)).sum()) for p in periods]
            if sum(line) >= 2:
                th_rows.append(th)
                th_vals.append([v if v else None for v in line])
        if th_rows:
            cnt = allr.groupby("period").size()
            p3 = folder / "3_problems_over_time.png"
            heat_table(p3, "When problems appear (cells = reviews)",
                       f"Source: {trend_note} (including complaints inside praise reviews), by review date in half-years; only problems complained about 2+ times are listed."
                       "\nA problem that appears across several periods deserves a check whether the same cause was never fixed.",
                       th_rows, [p for p in periods], th_vals, 0, max(max(v or 0 for v in r) for r in th_vals),
                       lambda v: f"{int(v)}", cmap=ORANGES, col_notes=[f"{cnt[p]} reviews" for p in periods],
                       color_values=[[(v / cnt[p] if v else None) for v, p in zip(r, periods)] for r in th_vals])
            out.append((p3, "Problems over time"))

    # 4. Reviews and average score per quarter (two small charts on one time axis, no double axis)
    sc = trend[trend["score10"].notna() & trend["date"].notna()].copy()
    if len(sc) >= 5:
        sc["q"] = sc["date"].dt.to_period("Q")
        g = sc.groupby("q")["score10"].agg(["size", "mean"])
        cn = {1: "Q1", 2: "Q2", 3: "Q3", 4: "Q4"}
        last_q = res["today"].to_period("Q")
        labels = [f"{q.year}\n{cn[q.quarter]}" + ("\n(incomplete)" if q == last_q else "") for q in g.index]
        fig, (a1, a2) = plt.subplots(2, 1, figsize=(11, 6.2), dpi=200, sharex=True,
                                     gridspec_kw={"height_ratios": [1, 1], "hspace": 0.35})
        fig.subplots_adjust(left=0.06, right=0.98, top=0.86, bottom=0.16)
        x = range(len(g))
        a1.bar(list(x), g["size"], color=BLUE, width=0.6)
        for i, v in enumerate(g["size"]):
            a1.text(i, v, f"{v}", ha="center", va="bottom", fontsize=10, color=INK2)
        a1.set_title("Number of reviews", loc="left", fontsize=11, color=INK)
        a2.plot(list(x), g["mean"], color=BLUE, lw=2, marker="o", ms=5)
        big = g[g["size"] >= A.MIN_PERIOD_REVIEWS]  # quarters with too few reviews do not count as "lowest"; one or two reviews can pull them down
        low = big["mean"].idxmin() if len(big) else None
        for i, (q, v) in enumerate(g["mean"].items()):
            a2.text(i, v + 0.12, f"{v:.1f}", ha="center", va="bottom", fontsize=10,
                    color=DARK if q == low else INK2, fontweight="bold" if q == low else "normal")
        a2.set_ylim(max(0, g["mean"].min() - 1), 10.5)
        a2.set_title(f"Average score (10-point; bold = lowest among quarters with {A.MIN_PERIOD_REVIEWS}+ reviews; the axis does not start at 0)", loc="left", fontsize=11, color=INK)
        for yv in range(int(max(0, g["mean"].min() - 1)) + 1, 11):
            a2.axhline(yv, color=GRID, lw=0.8, zorder=0)
            a2.text(-0.6, yv, f"{yv}", va="center", ha="right", fontsize=9, color=MUTED)
        a2.set_xticks(list(x), labels, fontsize=9.5, color=INK2)
        for a in (a1, a2):
            for s in ["top", "right", "left"]:
                a.spines[s].set_visible(False)
            a.spines["bottom"].set_color(GRID)
            a.set_yticks([])
        _frame(fig, "Reviews and average score per quarter", f"Source: {trend_note} (reviews with scores), by review date in quarters; in quarters with few reviews, one or two reviews can move the average.")
        p4 = folder / "4_quarterly_trend.png"
        fig.savefig(p4, facecolor="white")
        plt.close(fig)
        out.append((p4, "Quarterly trend"))

    # 5. In which months low-score reviews occur
    negd = trend[(trend["Low score"] == True) & trend["date"].notna()]  # noqa: E712
    if len(negd) >= MIN_NEG_FOR_MONTHS:  # with too few low-score reviews a month view is meaningless, so the chart is not drawn
        m = negd["date"].dt.month.value_counts()
        fig = plt.figure(figsize=(11, 3.6), dpi=200)
        ax = fig.add_axes([0.04, 0.2, 0.94, 0.58])
        vals = [int(m.get(i, 0)) for i in range(1, 13)]
        ax.bar(range(12), vals, color=[ORANGE if v else GRID for v in vals], width=0.6)
        for i, v in enumerate(vals):
            if v:
                ax.text(i, v, f"{v}", ha="center", va="bottom", fontsize=11, color=INK2)
        ax.set_xticks(range(12), ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], color=INK2)
        for s in ["top", "right", "left"]:
            ax.spines[s].set_visible(False)
        ax.spines["bottom"].set_color(GRID)
        ax.set_yticks([])
        _frame(fig, f"In which months the {len(negd)} low-score reviews occurred (all years added up)", f"Source: {trend_note}, normally collected low-score reviews, summed by the month of the review date; to see whether low-score reviews cluster in peak season.")
        p5 = folder / "5_low_score_months.png"
        fig.savefig(p5, facecolor="white")
        plt.close(fig)
        out.append((p5, "Low-score months"))

    # 6. Guest mix: one bar per platform, plus merged
    order = ["Business", "Family", "Couple", "Friends/group", "Solo", "Other"]
    if reviews["Guest group"].fillna("").isin(order).any():  # decided by the traveller types on reviews, not by the platform filter table
        colors = {"Business": BLUE, "Family": YELLOW, "Couple": ORANGE, "Friends/group": PINK, "Solo": GREEN, "Other": "#9aa1ac"}
        # Computed from traveller types marked on collected reviews (one review one group, same denominator as other guest-group numbers in the report);
        # counts on the platform filters (on Booking.com one review can count for several types) stay only in Excel for reference
        typed = reviews[reviews["Guest group"].fillna("").isin(order)]
        s = typed["Guest group"].value_counts()
        srcs6 = ", ".join(dict.fromkeys(pf.disp(p) for p in typed["platform"]))
        lines = [("Combined", s / s.sum(), int(s.sum()))]
        overlap, est = [], []
        fig = plt.figure(figsize=(11, 2.6), dpi=200)
        ax = fig.add_axes([0.2, 0.26, 0.77, 0.34])
        present = [o for o in order if o in s.index]
        for i, (name, share, n) in enumerate(lines):
            left = 0
            for grp in present:
                w = float(share.get(grp, 0))
                ax.barh(i, w, left=left, color=colors[grp], height=0.62, edgecolor="white", lw=1.5)
                if w >= 0.06:
                    ax.text(left + w / 2, i, f"{w:.0%}", ha="center", va="center", fontsize=10.5,
                            color="white" if grp in ("Business", "Couple") else INK, fontweight="bold" if name == "Combined" else "normal")
                left += w
            ax.text(-0.01, i, f"{name} ({n} reviews)", ha="right", va="center", fontsize=11, color=INK,
                    fontweight="bold" if name == "Combined" else "normal")
        ax.set_xlim(0, 1)
        ax.invert_yaxis()
        ax.axis("off")
        for k, grp in enumerate(present):
            fig.patches.append(plt.Rectangle((0.2 + k * 0.12, 0.70), 0.015, 0.05, color=colors[grp], transform=fig.transFigure))
            fig.text(0.22 + k * 0.12, 0.725, grp, fontsize=11, color=INK, va="center")
        _frame(fig, "Guest mix (all platforms merged)",
               f"Source: traveller types marked on collected reviews ({srcs6}), one review one type; Meituan and Fliggy reviews have no traveller type.")
        p6 = folder / "6_guest_mix.png"
        fig.savefig(p6, facecolor="white")
        plt.close(fig)
        out.append((p6, "Guest mix"))

    # 7. Average score of reviews complaining about each problem, next to the average of all reviews (same metric as the cover's "watch these" cards)
    sc7 = reviews[reviews["score10"].notna()]
    rows7 = []
    for th in A.THEMES:
        hit = sc7["Complaint themes"].str.contains(re.escape(th))
        if hit.sum() >= A.MIN_THEME_MENTIONS:
            rows7.append((th, float(sc7.loc[hit, "score10"].mean()), int(hit.sum())))
    rows7.sort(key=lambda x: x[1])
    if rows7:
        p7 = folder / "7_themes_lowering_score.png"
        labels7 = ["All reviews (for comparison)"] + [r[0] for r in rows7]
        vals7 = [float(sc7["score10"].mean())] + [r[1] for r in rows7]
        ns7 = [len(sc7)] + [r[2] for r in rows7]
        hbars(p7, "Average score of reviews complaining about this problem (10-point)",
              f"Source: all collected reviews (including complaints inside praise reviews), review counts in brackets; only problems complained about {A.MIN_THEME_MENTIONS}+ times are listed. A low score goes with the problem but is not necessarily caused by it.",
              labels7, vals7, lambda v, i: f"avg {v:.1f} ({ns7[i]} reviews)", color=[BLUE] + [ORANGE] * len(rows7), xmax=13)
        out.append((p7, "Themes lowering the score"))

    # 8. Platform tags
    tp = t["Platform topics"]
    if len(tp) and "topic" in tp.columns:
        # Tags of the same name on different platforms are merged into one (e.g. Ctrip and Meituan both have "well equipped"), not listed per channel
        tp = tp.assign(_l=tp["topic"].map(zh_label))
        srcs = ", ".join(dict.fromkeys(pf.disp(p) for p in tp["Platform"]))
        tp = tp.groupby("_l", sort=False)["mentions"].sum().sort_values(ascending=False).head(12)
        p8 = folder / "8_platform_tags.png"
        hbars(p8, "What guests mention most as good (tags counted by the platforms, merged)",
              f"Source: review tags and mention counts counted by {srcs}; tags of the same name are merged.",
              list(tp.index), [float(v) for v in tp.values], lambda v, i: f"{int(v)}", color=BLUE)
        out.append((p8, "Platform tags"))
    # 9. What guests praise
    pt = t.get("What guests praise")
    if pt is not None and len(pt) and pt["Reviews praising it"].sum():
        n_good = pt.attrs.get("n") or len(reviews[reviews["Low score"] != True])  # noqa: E712
        part = pt[pt["Reviews praising it"] > 0]
        p9 = folder / "9_praise.png"
        hbars(p9, f"Among the {n_good} analysed reviews, what guests praise",
              "Source: all collected reviews; low-score reviews do not count as praise; a praising wording near the aspect word is needed, detected automatically (misses some); one review can praise several aspects."
              + (f"\n{pt.attrs['unreadable']} Korean/Japanese reviews could not be read and are in the denominator, so shares are slightly low." if pt.attrs.get("unreadable") else ""),
              list(part["Aspect"]), [int(v) for v in part["Reviews praising it"]],
              lambda v, i: f"{v} ({part['Share'].iloc[i]:.0%})", color=BLUE)
        out.append((p9, "What guests praise"))

    # 10. What each guest group mentions (guest group x facilities/location/trip purpose): shows what main guest groups care about and what the listing should say
    sm = t.get("What each guest group mentions")
    if sm is not None and len(sm) and len(sm.attrs.get("groups", [])) >= 2:
        groups = sm.attrs["groups"]
        sell = sm[(sm["Category"] == "Selling point") & (sm["Mentions (all reviews)"] >= A.MIN_THEME_MENTIONS)]
        # sorted by the share among the two largest guest groups, most cared-about first
        top2 = [f"{g} ({n} reviews) share" for g, n in groups[:2]]
        sell = sell.assign(_k=sell[top2].max(axis=1)).sort_values("_k", ascending=False)
        base = sm[sm["Category"] == "Basics"]
        part = pd.concat([sell, base])
        rows = [r["Mentioned"] + (" (basics)" if r["Category"] == "Basics" else "") for _, r in part.iterrows()]
        vals = [[float(r[f"{g} ({n} reviews) share"]) * 100 for g, n in groups] for _, r in part.iterrows()]
        p10 = folder / "10_guest_group_mentions.png"
        heat_table(p10, "What each guest group mentions most in reviews (share of that group's reviews)",
                   "Source: reviews with a traveller type, counted per traveller type; the denominator is all reviews of that group; mentions inside low-score reviews are excluded; one review can mention several things."
                   f"\nOnly things mentioned {A.MIN_THEME_MENTIONS}+ times in all reviews are listed; the last two rows are basics that every hotel should do.",
                   rows, [f"{g} ({n} reviews)" for g, n in groups], vals,
                   0, max(max(v for v in r) for r in vals[:len(sell)]) if len(sell) else 1, lambda v: f"{v:.0f}%",
                   color_values=[[v if i < len(sell) else 0 for v in r] for i, r in enumerate(vals)])
        out.append((p10, "Guest group mentions"))
    # 11. What foreign-language reviews mention (by review language; language is not nationality): foreign-language reviews are usually few, so cells show counts and colour follows the share of that language's reviews
    lm = t.get("What foreign-language reviews mention")
    if lm is not None and len(lm):
        langs = [x for x in lm.attrs.get("langs", []) if x[0] != "Chinese"]
        if langs:
            cnt_cols = [f"{l} ({n} reviews) count" for l, n in langs] + [c for c in lm.columns if c.startswith("Chinese") and c.endswith("count")]
            share_cols = [c.replace("count", "share") for c in cnt_cols]
            fore = lm[[c for c in cnt_cols if not c.startswith("Chinese")]].sum(axis=1)
            part = lm[fore > 0].assign(_f=fore[fore > 0]).sort_values(["Category", "_f"], ascending=[True, False])
            if len(part):
                rows = [r["Mentioned"] + (" (basics)" if r["Category"] == "Basics" else "") for _, r in part.iterrows()]
                vals = [[int(r[c]) for c in cnt_cols] for _, r in part.iterrows()]
                # Colours: foreign-language columns by count (5+ is darkest, so a language with 1 review does not look dark because of '100%'); the Chinese column by share, relative to the largest value in that column
                zh_max = max([float(r[c]) for _, r in part.iterrows() for c in share_cols if c.startswith("Chinese")] or [1]) or 1
                shares = [[(float(r[s_]) / zh_max if s_.startswith("Chinese") else min(int(r[c]), 5) / 5)
                           for c, s_ in zip(cnt_cols, share_cols)] for _, r in part.iterrows()]
                p11 = folder / "11_foreign_language_mentions.png"
                heat_table(p11, "What foreign-language reviews mention (by review language; cells = counts)",
                           "Source: collected reviews, by review language; mentions inside low-score reviews are excluded; Korean and Japanese were translated into Chinese by hand before counting; foreign-language columns are coloured by count, the Chinese column by share."
                           "\nForeign-language reviews are few: read the counts and the originals, not only the shares; the Chinese column is for comparison.",
                           rows, [c.replace("count", "") for c in cnt_cols], vals, 0, 1, lambda v: f"{v}",
                           color_values=shares)
                out.append((p11, "Foreign-language mentions"))
    return out
