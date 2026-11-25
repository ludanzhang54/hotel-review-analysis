// 飞猪酒店点评抓取
// 前提：用户本人已在同一个浏览器里登录某宝账号（扫码或账号），不登录看不到点评。不要替用户登录。
// 页面：https://hotel.fliggy.com/hotel_detail2.htm?shid=飞猪酒店编号#hotel-review
//   编号：在 https://hotel.fliggy.com/hotel_list3.htm?cityName=城市&keywords=酒店名 的结果里，酒店链接上的 shid=数字（不用登录也能搜）。
// 用法：运行整段代码，再运行 r = await FliggyReviews.collect()，一页一页点"下一页"（每页间隔 3–4 秒），
//   最后用 JSON.stringify(r.rows.slice(0, 50)) 等分段取出，存成 reviews_fliggy_zh.json。
//   FliggyReviews.summary() 读总分、点评数、好评/差评条数和分项。
// 飞猪的每条点评没有总分，只有 4 个分项星级（清洁程度/地理位置/服务体验/性价比，老点评是卫生/设施/服务/性价比），
// 总分按分项平均；0 星的分项当作没打分，不算进平均。
(() => {
  const parse = () => [...document.querySelectorAll("li.tb-r-comment")].map((li) => {
    const subs = {};
    li.querySelectorAll(".starscore li").forEach((x) => {
      const k = x.querySelector("span").textContent.replace(/[:：]/, "").trim();
      const em = x.querySelector("em");
      subs[k] = em ? parseFloat(em.style.width) / 20 : null;
    });
    const v = Object.values(subs).filter((s) => s);
    const info = li.querySelector(".tb-r-body .tb-r-info");
    const d = info && info.textContent.match(/\d{4}-\d\d-\d\d/);
    return {
      review_language: "zh", score: v.length ? Math.round(v.reduce((a, b) => a + b, 0) / v.length * 100) / 100 : null,
      score_scale: 5, date: d ? d[0] : null, subscores: subs,
      traveler_type: "", room_type: "", reviewer_country: "",
      title: ((li.querySelector(".comment-name") || {}).textContent || "").replace(/[“”\n]/g, " ").trim(),
      text: ((li.querySelector(".tb-r-cnt") || {}).innerText || "").trim(),
    };
  });
  const sleep = (ms) => new Promise((s) => setTimeout(s, ms));
  const sig = () => [...document.querySelectorAll("li.tb-r-comment .tb-r-cnt")].map((x) => x.textContent.slice(0, 20)).join("|");
  window.FliggyReviews = {
    summary() {
      const t = document.body.innerText;
      const sec = t.slice(t.indexOf("住客评价\n"));
      const n = (re) => { const m = sec.match(re); return m ? parseFloat(m[1]) : null; };
      const subs = {};
      for (const k of ["地理位置", "清洁程度", "服务体验", "性价比"]) {
        const v = n(new RegExp(k + "\\s*\\n?\\s*(\\d\\.\\d)"));
        if (v) subs[k] = v; // 页面显示 0.0 的分项是没公布，不填
      }
      return { rating: n(/住客评价\s*\n?\s*(\d\.\d)/), scale: 5, total_reviews: n(/(\d+)条点评/),
               positive_reviews: n(/好评 \((\d+)\)/), negative_reviews: n(/差评 \((\d+)\)/), subscores: subs, note: location.href };
    },
    async collect({ maxPages = 40, pageDelayMs = 3500 } = {}) {
      const rows = [];
      for (let p = 1; p <= maxPages; p++) {
        rows.push(...parse());
        const next = document.querySelector("a.next-page");
        if (!next) break;
        const before = sig();
        next.click();
        let ok = false;
        for (let i = 0; i < 30; i++) { await sleep(500); if (sig() !== before) { ok = true; break; } }
        if (!ok) break;
        await sleep(pageDelayMs + Math.random() * 1000);
      }
      return { platform: "飞猪", hotel: document.title.split("预订")[0], url: location.href, language: { code: "zh" }, rows };
    },
  };
  return "FliggyReviews 已就绪";
})();
