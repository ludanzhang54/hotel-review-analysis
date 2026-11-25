// Fliggy hotel review collector
// Prerequisite: the user has signed in to their own account in the same browser (QR code or account); reviews are not visible without signing in. Never sign in for the user.
// Page: https://hotel.fliggy.com/hotel_detail2.htm?shid=<Fliggy hotel id>#hotel-review
//   Id: in the results of https://hotel.fliggy.com/hotel_list3.htm?cityName=<city>&keywords=<hotel name>, the shid=number in the hotel link (searching needs no sign-in).
// Usage: run the whole code, then r = await FliggyReviews.collect(), click "next page" page by page (3-4 seconds between pages),
//   finally take the rows out in parts, e.g. JSON.stringify(r.rows.slice(0, 50)), and save as reviews_fliggy_zh.json.
//   FliggyReviews.summary() reads the overall rating, review count, positive/negative counts and sub-scores.
// A Fliggy review has no overall score, only 4 sub-score stars (cleanliness/location/service/value; old reviews use hygiene/facilities/service/value),
// so the score is the average of the sub-scores; a 0-star sub-score counts as not rated and is left out of the average.
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
        if (v) subs[k] = v; // a sub-score shown as 0.0 is not published, leave it out
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
      return { platform: "Fliggy", hotel: document.title.split("预订")[0], url: location.href, language: { code: "zh" }, rows };
    },
  };
  return "FliggyReviews ready";
})();
