// Read the summary of a Ctrip (China site) hotel detail page
// Usage: run the whole code after https://hotels.ctrip.com/hotels/detail/?hotelId=<id> has loaded.
// The returned JSON can go straight into "platforms" of summary.json (platform name "Ctrip"); numbers are real numbers.
// hidden_reviews (folded reviews) cannot be read here; read it from the end of the Trip.com review list instead, see SKILL.md.
(() => {
  const t = document.body.innerText;
  const n = (s) => (s == null ? null : parseFloat(String(s).replace(/,/g, "")));
  const sec = t.slice(t.indexOf("真实住客点评"));
  const sub = (name) => { const m = sec.match(new RegExp(name + "\\n(\\d\\.\\d)")); return m ? n(m[1]) : null; };
  const tagBlock = sec.slice(sec.indexOf("热门提及"), sec.indexOf("住客印象") > 0 ? sec.indexOf("住客印象") : undefined);
  // the tag area starts with review filters (all reviews, recommended, with photos, negative...), which are not topics: remove them
  const FILTERS = /^(所有点评|全部点评|全部|值得推荐|好评|中评|差评|有图|视频|有视频|有图\/视频|偏好相似|最新点评|追评|商务出差|家庭亲子|情侣出游|朋友出游|独自旅行|代人预订|其他)$/;
  const found = [...tagBlock.matchAll(/([^\n()（）]+?)\s*[(（]([\d,]+)[)）]/g)]
    .map((m) => ({ topic: m[1].trim(), mentions: n(m[2]) }));
  const topics = found.filter((x) => !FILTERS.test(x.topic));
  const excluded = found.filter((x) => FILTERS.test(x.topic)).map((x) => x.topic);
  const out = {
    Ctrip: {
      rating: n((t.match(/\n(\d\.\d)\n(?:超棒|很好|不错|好|一般)?\n?显示所有/) || [])[1]),
      scale: 5,
      total_reviews: n((t.match(/显示所有([\d,]+)条点评/) || [])[1]),
      subscores: Object.fromEntries(
        ["卫生", "设施", "环境", "服务"].map((k) => [k, sub(k)]).filter(([, v]) => v != null)),
      // the count on the page's "negative (N)" filter, used to calculate the low-score share
      negative_reviews: (found.find((x) => x.topic === "差评") || {}).mentions ?? null,
      topics,
      excluded,
      note: document.title.replace(/预订价格.*$/, "") + " " + location.href,
    },
  };
  return JSON.stringify(out);
})();
