"""
点评报告的图表（PNG）。风格：上方一句话标题说明图在看什么，下方灰色小字写数据来源；
数字直接写在色块或柱子上，不用看坐标轴也能读。

用法：由 report.py 调用 make_charts(res, folder)，返回 [(文件名, 标题), ...]。
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

# 中文字体：Windows 微软雅黑、Mac 苹方、Linux 思源黑体，哪个有用哪个
_installed = {f.name for f in font_manager.fontManager.ttflist}
FONT = next((f for f in ["Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", "Noto Sans CJK SC", "Noto Sans CJK TC",
                         "Source Han Sans SC", "SimHei", "Arial Unicode MS"] if f in _installed), "sans-serif")
if FONT == "sans-serif":
    # 没有中文字体时图里的中文会变成方框；宁可停下，也不交出看不了的图
    sys.exit("没有找到中文字体，图里的中文会显示成方框。请先安装中文字体（例如 Noto Sans CJK SC / 思源黑体）再运行。")
plt.rcParams.update({"font.family": FONT, "axes.unicode_minus": False, "font.size": 11})

INK, INK2, MUTED, GRID = "#1f2328", "#4a505a", "#8a8f98", "#e6e8eb"
BLUE, ORANGE, GREEN, YELLOW, PINK = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"
RED = "#d64545"
BLUES = LinearSegmentedColormap.from_list("blues", ["#dce8f8", "#5b9be6", "#1d5fb0", "#123f7a"])
ORANGES = LinearSegmentedColormap.from_list("oranges", ["#fbe3d8", "#f5b08f", "#eb6834", "#c9581f"])
EMPTY = "#f0f1f3"
DARK = "#123f7a"
# 平台英文标签的中文（按"英文（中文）"显示）
EN_ZH = {"breakfast": "早餐", "in room amenities": "房内设施", "location": "位置", "cleanliness": "卫生",
         "service": "服务", "facilities": "设施", "value for money": "性价比", "room comfort": "房间舒适度",
         "staff": "员工", "comfort": "舒适度", "food": "餐饮", "pool": "泳池", "noise": "噪音", "wifi": "无线网络"}


def zh_label(text):
    t = str(text)
    return f"{t}（{EN_ZH[t.lower()]}）" if t.lower() in EN_ZH else t


def _frame(fig, title, source):
    fig.text(0.012, 0.975, title, fontsize=15, color=INK2, va="top")
    if source:
        fig.text(0.012, 0.018, source, fontsize=9, color=MUTED, va="bottom", wrap=True)


def _ink_on(rgb):
    r, g, b = rgb[:3]
    return "white" if (0.299 * r + 0.587 * g + 0.114 * b) < 0.6 else INK


def heat_table(path, title, source, rows, cols, values, vmin, vmax, fmt, cmap=BLUES, bold_rows=(), col_notes=None,
               color_values=None):
    """色块表：rows × cols，values[i][j] 为 None 时显示"—"。"""
    nr, nc = len(rows), len(cols)
    H = 1.9 + 0.48 * nr
    fig = plt.figure(figsize=(max(8, 2.2 + 1.7 * nc), H), dpi=200)
    # 上下留白按英寸固定（标题约 1.15 英寸、脚注约 0.75 英寸），行多时表格不会和标题隔一大片空白
    ax = fig.add_axes([0.02, 0.75 / H, 0.96, (H - 0.75 - 1.15) / H])
    ax.set_xlim(-2.6, nc)
    ax.set_ylim(nr, -0.2)
    ax.axis("off")
    for j, c in enumerate(cols):
        ax.text(j + 0.5, -0.38, c, ha="center", va="bottom", fontsize=11 if nc > 5 else 11.5, color=INK)
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


MIN_NEG_FOR_MONTHS = 20  # 差评少于 20 条不画"差评月份"


def make_charts(res, folder):
    t, reviews, pool, pf = res["tables"], res["reviews"], res["pool"], res["platforms"]
    ov = t["综合汇总"]
    item = dict(zip(ov["项目"], ov["合并结果"]))
    folder = Path(folder) / "图表"
    folder.mkdir(exist_ok=True)
    out = []
    date = str(res["today"].date())

    # 1. 评分总览：总评分 + 分项，各平台和合并
    plats = pf.rating_platforms()
    subs = t["分项评分"]
    groups = [g for g, _ in A.SUBSCORE_GROUPS]
    rows, values = ["总评分"], [[pf.rating10(p) for p in plats] + [A.wavg([(pf.rating10(p), pf.weight(p)[0]) for p in plats])]]
    for g in groups:
        part = subs[subs["统一分项"] == g] if len(subs) else subs
        if part is None or part.empty:
            continue
        line = []
        for p in plats:
            m = part[part["平台"].map(pf.same) == pf.same(p)]
            line.append(float(m["换算成10分制"].iloc[0]) if len(m) else None)
        merged = item.get(next((k for k in item if str(k).startswith(f"分项：{g}")), None))
        rows.append(g)
        values.append(line + [float(merged) if merged is not None else None])
    env = subs[(subs["统一分项"].isna()) & subs["分项"].astype(str).str.lower().isin(["环境", "location"])
               & subs["平台"].isin(["携程", "Trip.com"])] if len(subs) else subs
    if len(env):
        v = float(env["换算成10分制"].iloc[0])
        rows.append("环境（携程）")
        values.append([v if pf.same(p) in ("携程", "Trip.com") else None for p in plats] + [v])
    cols = [pf.disp(p) for p in plats] + ["合并"]
    notes = [f"{int(pf.published(p) or 0)} 条点评" for p in plats] + [f"{int(pf.all_published())} 条点评"]
    vals = [v for r in values for v in r if v is not None]
    p1 = folder / "1_评分总览.png"
    heat_table(p1, "各平台评分和分项（10 分制，颜色越深分越高）",
               f"来源：各平台酒店页面公布的评分，{date} 读取。合并 = 按各平台公布的点评数加权。"
               "\n各平台分项名称不同，已归成同一类（如英文 Cleanliness 和'卫生'都算卫生）；'—' = 该平台没有这一项。'环境'是携程的分项，和'位置'不是一回事。",
               rows, cols, values, min(vals + [8]) - 0.3, 10, lambda v: f"{v:.1f}", bold_rows=("总评分",), col_notes=notes)
    out.append((p1, "评分总览"))

    # 2. 客人在抱怨什么：用全部点评（含好评里顺带的抱怨），不只看差评——差评常常只有几条，样本太小
    # 比例和条数只用正常抓到的点评；补抓的低分点评不进图（只当原文证据）
    n_all = len(reviews)
    counts = []
    for th in A.THEMES:
        hit = reviews["命中主题"].str.contains(re.escape(th))
        if hit.sum():
            counts.append((th, int(hit.sum()), int((hit & (reviews["差评"] == True)).sum())))  # noqa: E712
    counts.sort(key=lambda x: -x[1])
    if counts:
        p2 = folder / "2_客人抱怨什么.png"
        hbars(p2, f"本次分析的 {n_all} 条点评里，每个问题被抱怨了几条",
              f"来源：抓到的全部点评（含好评里顺带的抱怨），按关键词自动识别客人在抱怨的那句话；一条点评可以提到多个问题。"
              f"\n少于 {A.MIN_THEME_MENTIONS} 条的属于个别情况，不能当成普遍问题。",
              [c[0] for c in counts], [c[1] for c in counts],
              lambda v, i: f"{v} 条" + (f"（其中差评 {counts[i][2]} 条）" if counts[i][2] else ""), color=ORANGE,
              xmax=counts[0][1] * 1.7)
        out.append((p2, "客人抱怨什么"))

    # 3–5 看时间的图只用点评已经抓全的平台和语言；没抓全的平台旧点评少只是没抓到
    trend, trend_used = A.complete_sample(reviews, res["summary"])
    trend_note = "用到的平台：" + ("、".join(dict.fromkeys(pf.disp(x) for x in trend_used)) if trend_used else "无")

    # 3. 抱怨的问题在哪些时间出现（按半年）
    allr = trend[trend["date"].notna()].copy()
    if len(allr):
        allr["期"] = allr["date"].dt.year.astype(str) + allr["date"].dt.month.map(lambda m: " 上半年" if m <= 6 else " 下半年")
        periods = sorted(allr["期"].unique())
        th_rows, th_vals = [], []
        for th in A.THEMES:
            line = [int(allr.loc[allr["期"] == p, "命中主题"].str.contains(re.escape(th)).sum()) for p in periods]
            if sum(line) >= 2:
                th_rows.append(th)
                th_vals.append([v if v else None for v in line])
        if th_rows:
            cnt = allr.groupby("期").size()
            p3 = folder / "3_问题出现的时间.png"
            heat_table(p3, "客人抱怨的问题出现在哪些时间（格子里是条数，颜色按占当期点评的比例）",
                       f"来源：{trend_note}（含好评里顺带的抱怨），按点评日期分半年；只列被抱怨 2 次以上的问题。"
                       "\n同一类问题跨好几个时期都出现，值得查是不是同一个原因没解决。",
                       th_rows, [p.replace("年 ", "年") for p in periods], th_vals, 0, max(max(v or 0 for v in r) for r in th_vals),
                       lambda v: f"{int(v)}", cmap=ORANGES, col_notes=[f"{cnt[p]} 条点评" for p in periods],
                       color_values=[[(v / cnt[p] if v else None) for v, p in zip(r, periods)] for r in th_vals])
            out.append((p3, "问题出现的时间"))

    # 4. 每季度点评量和平均分（两张小图，同一条时间轴，不用双坐标）
    sc = trend[trend["score10"].notna() & trend["date"].notna()].copy()
    if len(sc) >= 5:
        sc["q"] = sc["date"].dt.to_period("Q")
        g = sc.groupby("q")["score10"].agg(["size", "mean"])
        cn = {1: "一", 2: "二", 3: "三", 4: "四"}
        last_q = res["today"].to_period("Q")
        labels = [f"{q.year}年\n{cn[q.quarter]}季度" + ("\n（未完）" if q == last_q else "") for q in g.index]
        fig, (a1, a2) = plt.subplots(2, 1, figsize=(11, 6.2), dpi=200, sharex=True,
                                     gridspec_kw={"height_ratios": [1, 1], "hspace": 0.35})
        fig.subplots_adjust(left=0.06, right=0.98, top=0.86, bottom=0.16)
        x = range(len(g))
        a1.bar(list(x), g["size"], color=BLUE, width=0.6)
        for i, v in enumerate(g["size"]):
            a1.text(i, v, f"{v}", ha="center", va="bottom", fontsize=10, color=INK2)
        a1.set_title("点评条数", loc="left", fontsize=11, color=INK)
        a2.plot(list(x), g["mean"], color=BLUE, lw=2, marker="o", ms=5)
        big = g[g["size"] >= A.MIN_PERIOD_REVIEWS]  # 条数太少的季度不算"最低"，一两条就能拉低
        low = big["mean"].idxmin() if len(big) else None
        for i, (q, v) in enumerate(g["mean"].items()):
            a2.text(i, v + 0.12, f"{v:.1f}", ha="center", va="bottom", fontsize=10,
                    color=DARK if q == low else INK2, fontweight="bold" if q == low else "normal")
        a2.set_ylim(max(0, g["mean"].min() - 1), 10.5)
        a2.set_title(f"平均分（10 分制，加粗 = {A.MIN_PERIOD_REVIEWS} 条以上的季度里最低的；纵轴不从 0 开始）", loc="left", fontsize=11, color=INK)
        for yv in range(int(max(0, g["mean"].min() - 1)) + 1, 11):
            a2.axhline(yv, color=GRID, lw=0.8, zorder=0)
            a2.text(-0.6, yv, f"{yv}", va="center", ha="right", fontsize=9, color=MUTED)
        a2.set_xticks(list(x), labels, fontsize=9.5, color=INK2)
        for a in (a1, a2):
            for s in ["top", "right", "left"]:
                a.spines[s].set_visible(False)
            a.spines["bottom"].set_color(GRID)
            a.set_yticks([])
        _frame(fig, "每季度的点评条数和平均分", f"来源：{trend_note}（有分数的点评），按点评日期分季度；条数少的季度，一两条点评就能拉动平均分。")
        p4 = folder / "4_季度趋势.png"
        fig.savefig(p4, facecolor="white")
        plt.close(fig)
        out.append((p4, "季度趋势"))

    # 5. 差评发生在哪几个月
    negd = trend[(trend["差评"] == True) & trend["date"].notna()]  # noqa: E712
    if len(negd) >= MIN_NEG_FOR_MONTHS:  # 差评太少时按月份看没有意义，不出这张图
        m = negd["date"].dt.month.value_counts()
        fig = plt.figure(figsize=(11, 3.6), dpi=200)
        ax = fig.add_axes([0.04, 0.2, 0.94, 0.58])
        vals = [int(m.get(i, 0)) for i in range(1, 13)]
        ax.bar(range(12), vals, color=[ORANGE if v else GRID for v in vals], width=0.6)
        for i, v in enumerate(vals):
            if v:
                ax.text(i, v, f"{v}", ha="center", va="bottom", fontsize=11, color=INK2)
        ax.set_xticks(range(12), [f"{i}月" for i in range(1, 13)], color=INK2)
        for s in ["top", "right", "left"]:
            ax.spines[s].set_visible(False)
        ax.spines["bottom"].set_color(GRID)
        ax.set_yticks([])
        _frame(fig, f"{len(negd)} 条差评发生在哪几个月（各年份合计）", f"来源：{trend_note}，正常抓到的差评，按点评日期的月份合计；看差评是否集中在旺季。")
        p5 = folder / "5_差评月份.png"
        fig.savefig(p5, facecolor="white")
        plt.close(fig)
        out.append((p5, "差评月份"))

    # 6. 住客构成：每个平台一条，再加合并
    order = ["商务", "家庭", "情侣", "朋友/团体", "独自", "其他"]
    if reviews["客群"].fillna("").isin(order).any():  # 按点评上的住客类型判断要不要画，不看平台筛选项表
        colors = {"商务": BLUE, "家庭": YELLOW, "情侣": ORANGE, "朋友/团体": PINK, "独自": GREEN, "其他": "#9aa1ac"}
        # 用抓到的点评上标的住客类型算（一条点评只算一类，和报告里其他客群数字同一个分母）；
        # 平台筛选项上的条数（Booking.com 一条点评能算几类）只放在 Excel 里备查
        typed = reviews[reviews["客群"].fillna("").isin(order)]
        s = typed["客群"].value_counts()
        srcs6 = "、".join(dict.fromkeys(pf.disp(p) for p in typed["platform"]))
        lines = [("合并", s / s.sum(), int(s.sum()))]
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
                            color="white" if grp in ("商务", "情侣") else INK, fontweight="bold" if name == "合并" else "normal")
                left += w
            ax.text(-0.01, i, f"{name}（{n} 条）", ha="right", va="center", fontsize=11, color=INK,
                    fontweight="bold" if name == "合并" else "normal")
        ax.set_xlim(0, 1)
        ax.invert_yaxis()
        ax.axis("off")
        for k, grp in enumerate(present):
            fig.patches.append(plt.Rectangle((0.2 + k * 0.12, 0.70), 0.015, 0.05, color=colors[grp], transform=fig.transFigure))
            fig.text(0.22 + k * 0.12, 0.725, grp, fontsize=11, color=INK, va="center")
        _frame(fig, "住客构成（各平台合并）",
               f"来源：抓到的点评上标的住客类型（{srcs6}），一条点评算一类；美团、飞猪的点评没有住客类型。")
        p6 = folder / "6_住客构成.png"
        fig.savefig(p6, facecolor="white")
        plt.close(fig)
        out.append((p6, "住客构成"))

    # 7. 抱怨各问题的点评平均几分，和全部点评的平均分放在一起比（和报告封面"最该盯的事"用同一个指标）
    sc7 = reviews[reviews["score10"].notna()]
    rows7 = []
    for th in A.THEMES:
        hit = sc7["命中主题"].str.contains(re.escape(th))
        if hit.sum() >= A.MIN_THEME_MENTIONS:
            rows7.append((th, float(sc7.loc[hit, "score10"].mean()), int(hit.sum())))
    rows7.sort(key=lambda x: x[1])
    if rows7:
        p7 = folder / "7_拉低分数的主题.png"
        labels7 = ["全部点评（对比）"] + [r[0] for r in rows7]
        vals7 = [float(sc7["score10"].mean())] + [r[1] for r in rows7]
        ns7 = [len(sc7)] + [r[2] for r in rows7]
        hbars(p7, "抱怨这个问题的点评，平均几分（10 分制）",
              f"来源：抓到的全部点评（含好评里顺带的抱怨），括号里是点评条数；只列被抱怨 {A.MIN_THEME_MENTIONS} 条以上的问题。分数低和这个问题有关，不一定是它造成的。",
              labels7, vals7, lambda v, i: f"平均 {v:.1f} 分（{ns7[i]} 条）", color=[BLUE] + [ORANGE] * len(rows7), xmax=13)
        out.append((p7, "拉低分数的主题"))

    # 8. 平台标签
    tp = t["平台话题"]
    if len(tp) and "话题" in tp.columns:
        # 各平台同名标签合并成一个（例如携程、美团都有"设施齐全"），不分渠道列
        tp = tp.assign(_l=tp["话题"].map(zh_label))
        srcs = "、".join(dict.fromkeys(pf.disp(p) for p in tp["平台"]))
        tp = tp.groupby("_l", sort=False)["提及次数"].sum().sort_values(ascending=False).head(12)
        p8 = folder / "8_平台标签.png"
        hbars(p8, "客人最常提到的好处（平台自己统计的标签，各平台合并）",
              f"来源：{srcs}自己统计的点评标签和提及次数，同名标签已合并。",
              list(tp.index), [float(v) for v in tp.values], lambda v, i: f"{int(v)} 次", color=BLUE)
        out.append((p8, "平台标签"))
    # 9. 好评在夸什么
    pt = t.get("好评在夸什么")
    if pt is not None and len(pt) and pt["夸它的点评条数"].sum():
        n_good = pt.attrs.get("n") or len(reviews[reviews["差评"] != True])  # noqa: E712
        part = pt[pt["夸它的点评条数"] > 0]
        p9 = folder / "9_好评在夸什么.png"
        hbars(p9, f"本次分析的 {n_good} 条点评里，客人在夸什么",
              "来源：抓到的全部点评，差评不算夸；方面词附近有夸奖的说法才算，自动识别，会漏掉一些；一条点评可以夸好几个方面。"
              + (f"\n{pt.attrs['unreadable']} 条韩文、日文点评识别不了，算在分母里，比例略偏低。" if pt.attrs.get("unreadable") else ""),
              list(part["方面"]), [int(v) for v in part["夸它的点评条数"]],
              lambda v, i: f"{v} 条（{part['占比'].iloc[i]:.0%}）", color=BLUE)
        out.append((p9, "好评在夸什么"))

    # 10. 各类客人提到什么（客群 × 设施/位置/出行目的）：看主要客群在意什么，酒店介绍该写什么
    sm = t.get("各类客人提到什么")
    if sm is not None and len(sm) and len(sm.attrs.get("groups", [])) >= 2:
        groups = sm.attrs["groups"]
        sell = sm[(sm["类别"] == "卖点") & (sm["全部点评提到条数"] >= A.MIN_THEME_MENTIONS)]
        # 按最大的两类客人提到的比例排序，最在意的排前面
        top2 = [f"{g}（{n} 条）占比" for g, n in groups[:2]]
        sell = sell.assign(_k=sell[top2].max(axis=1)).sort_values("_k", ascending=False)
        base = sm[sm["类别"] == "基本功"]
        part = pd.concat([sell, base])
        rows = [r["提到的东西"] + ("（基本功）" if r["类别"] == "基本功" else "") for _, r in part.iterrows()]
        vals = [[float(r[f"{g}（{n} 条）占比"]) * 100 for g, n in groups] for _, r in part.iterrows()]
        p10 = folder / "10_不同客人夸什么.png"
        heat_table(p10, "各类客人在点评里最常提到什么（每类客人里提到的比例）",
                   "来源：有住客类型的点评，按住客类型分开算，分母是这类客人的全部点评，差评里提到的不算；一条点评可以提到好几样。"
                   f"\n只列全部点评里提到 {A.MIN_THEME_MENTIONS} 次以上的；最后两行是每家酒店都该做到的基本功。",
                   rows, [f"{g}（{n} 条）" for g, n in groups], vals,
                   0, max(max(v for v in r) for r in vals[:len(sell)]) if len(sell) else 1, lambda v: f"{v:.0f}%",
                   color_values=[[v if i < len(sell) else 0 for v in r] for i, r in enumerate(vals)])
        out.append((p10, "不同客人夸什么"))
    # 11. 外文点评提到什么（按点评语言，语言不等于国籍）：外文点评通常很少，格子里写条数，颜色按占这种语言点评的比例
    lm = t.get("外文点评提到什么")
    if lm is not None and len(lm):
        langs = [x for x in lm.attrs.get("langs", []) if x[0] != "中文"]
        if langs:
            cnt_cols = [f"{l}（{n} 条）条数" for l, n in langs] + [c for c in lm.columns if c.startswith("中文") and c.endswith("条数")]
            share_cols = [c.replace("条数", "占比") for c in cnt_cols]
            fore = lm[[c for c in cnt_cols if not c.startswith("中文")]].sum(axis=1)
            part = lm[fore > 0].assign(_f=fore[fore > 0]).sort_values(["类别", "_f"], ascending=[True, False])
            if len(part):
                rows = [r["提到的东西"] + ("（基本功）" if r["类别"] == "基本功" else "") for _, r in part.iterrows()]
                vals = [[int(r[c]) for c in cnt_cols] for _, r in part.iterrows()]
                # 颜色：外文列按条数（5 条以上最深，1 条的语言不会因为"100%"显得很深）；中文列按占比，和本列最大值比
                zh_max = max([float(r[c]) for _, r in part.iterrows() for c in share_cols if c.startswith("中文")] or [1]) or 1
                shares = [[(float(r[s_]) / zh_max if s_.startswith("中文") else min(int(r[c]), 5) / 5)
                           for c, s_ in zip(cnt_cols, share_cols)] for _, r in part.iterrows()]
                p11 = folder / "11_外文点评提到什么.png"
                heat_table(p11, "外文点评里提到什么（按点评语言，格子里是条数）",
                           "来源：抓到的点评，按点评语言分，差评里提到的不算；韩文、日文是人工译成中文后统计的；外文列颜色按条数，中文列按比例。"
                           "\n外文点评很少，看条数和原文，不要只看比例；中文一列用来对比。",
                           rows, [c.replace("条数", "") for c in cnt_cols], vals, 0, 1, lambda v: f"{v} 条",
                           color_values=shares)
                out.append((p11, "外文点评提到什么"))
    return out
