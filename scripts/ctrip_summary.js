// 携程国内站酒店详情页汇总读取
// 用法：在 https://hotels.ctrip.com/hotels/detail/?hotelId=编号 页面加载完后运行整段代码。
// 返回的 JSON 可以直接放进 summary.json 的 "platforms" 里（平台名"携程"），数字都是数值。
// hidden_reviews（被折叠的条数）这里读不到，要另外从 Trip.com 点评列表末尾读，见 SKILL.md。
(() => {
  const t = document.body.innerText;
  const n = (s) => (s == null ? null : parseFloat(String(s).replace(/,/g, "")));
  const sec = t.slice(t.indexOf("真实住客点评"));
  const sub = (name) => { const m = sec.match(new RegExp(name + "\\n(\\d\\.\\d)")); return m ? n(m[1]) : null; };
  const tagBlock = sec.slice(sec.indexOf("热门提及"), sec.indexOf("住客印象") > 0 ? sec.indexOf("住客印象") : undefined);
  // 标签区前面混着点评筛选项（所有点评、值得推荐、有图、差评等），不是话题，去掉
  const FILTERS = /^(所有点评|全部点评|全部|值得推荐|好评|中评|差评|有图|视频|有视频|有图\/视频|偏好相似|最新点评|追评|商务出差|家庭亲子|情侣出游|朋友出游|独自旅行|代人预订|其他)$/;
  const found = [...tagBlock.matchAll(/([^\n()（）]+?)\s*[(（]([\d,]+)[)）]/g)]
    .map((m) => ({ 话题: m[1].trim(), 提及次数: n(m[2]) }));
  const topics = found.filter((x) => !FILTERS.test(x.话题));
  const excluded = found.filter((x) => FILTERS.test(x.话题)).map((x) => x.话题);
  const out = {
    携程: {
      rating: n((t.match(/\n(\d\.\d)\n(?:超棒|很好|不错|好|一般)?\n?显示所有/) || [])[1]),
      scale: 5,
      total_reviews: n((t.match(/显示所有([\d,]+)条点评/) || [])[1]),
      subscores: Object.fromEntries(
        ["卫生", "设施", "环境", "服务"].map((k) => [k, sub(k)]).filter(([, v]) => v != null)),
      // 页面"差评 (N)"筛选项上的条数，用来算差评占比
      negative_reviews: (found.find((x) => x.话题 === "差评") || {}).提及次数 ?? null,
      topics,
      excluded,
      note: document.title.replace(/预订价格.*$/, "") + " " + location.href,
    },
  };
  return JSON.stringify(out);
})();
