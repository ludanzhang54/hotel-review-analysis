// 美团酒店点评抓取（手机版点评页）
// 前提：用户本人已在同一个浏览器里登录美团（手机号 + 验证码）。不要替用户登录。
// 页面：https://i.meituan.com/poi/美团酒店编号/feedbacks （编号从美团酒店搜索结果进入详情页后，网址里的 poiId=数字）
//   浏览器要用手机尺寸打开（电脑版酒店页会被拒绝访问）。
// 用法（每一页都这样做，慢慢翻）：
//   1. 打开第 1 页，运行整段代码，再运行 MeituanReviews.savePage()：读出这一页，存到浏览器里，返回已存条数。
//   2. 运行 MeituanReviews.next()：点"下一页"。等页面打开后至少再等 10 秒，重复第 1 步（重新运行整段代码 + savePage）。
//   3. 最后一页 next() 返回 "已经是最后一页"。运行 MeituanReviews.dump(0, 50)、dump(50, 100)…… 分段取出，存成 reviews_meituan_zh.json。
// 注意：不要用脚本连续请求翻页地址，美团会限流（返回 403 拒绝访问）。被限流就停下，过一段时间再接着翻，已存的不会丢。
// 读不到昵称以外的身份信息；昵称不保存。
(() => {
  // 按美团酒店编号分开存，连着分析两家酒店也不会混在一起
  const poi = (location.pathname.match(/poi\/(\d+)/) || [])[1] || "unknown";
  const KEY = "rs_meituan_" + poi;
  const load = () => { try { return JSON.parse(localStorage.getItem(KEY) || "[]"); } catch (e) { return []; } };
  const parse = () => [...document.querySelectorAll(".feedbackCard")].map((c) => {
    const stars = [...c.querySelectorAll(".stars img")].map((i) => i.className);
    const score = stars.reduce((s, cls) => s + (/star_full/.test(cls) ? 1 : /half/.test(cls) ? 0.5 : 0), 0);
    const cm = c.querySelector(".comment");
    return {
      review_language: "zh", score, score_scale: 5,
      date: (c.querySelector(".time") || {}).textContent || null,
      traveler_type: "", room_type: "", reviewer_country: "", title: "",
      // 长点评被折叠在 .feedbackmore 里，textContent 会连同折叠部分一起读出
      text: cm ? cm.textContent.replace(/\s+/g, " ").trim() : "",
    };
  });
  window.MeituanReviews = {
    reset() { localStorage.removeItem(KEY); return "已清空"; },
    summary() {
      const t = document.body.innerText;
      const total = (t.match(/评价\s*(\d+)/) || [])[1];
      const tags = [...t.matchAll(/^([^\n\d]{2,8}) (\d+)$/gm)].map((m) => ({ 话题: m[1].trim(), 提及次数: +m[2] }))
        .filter((x) => !/^(评价|晒图)$/.test(x.话题));
      return { total_reviews: total ? +total : null, topics: tags, note: location.href };
    },
    savePage() {
      if (/403/.test(document.title)) return "被限流（403），先停下，过一段时间再从这一页继续";
      const have = load();
      const seen = new Set(have.map((r) => r.date + r.text.slice(0, 40)));
      const rows = parse().filter((r) => !seen.has(r.date + r.text.slice(0, 40)));
      localStorage.setItem(KEY, JSON.stringify(have.concat(rows)));
      return `这一页新增 ${rows.length} 条，共 ${have.length + rows.length} 条`;
    },
    next() {
      const a = [...document.querySelectorAll("a")].find((x) => x.textContent.trim() === "下一页");
      if (!a || /disabled/.test(a.className)) return "已经是最后一页";
      a.click();
      return "已点下一页，页面打开后再等 10 秒";
    },
    count() { return load().length; },
    dump(from, to) { return JSON.stringify(load().slice(from, to)); },
  };
  return "MeituanReviews 已就绪";
})();
