// Meituan hotel review collector (mobile review page)
// Prerequisite: the user has signed in to Meituan in the same browser (phone number + code). Never sign in for the user.
// Page: https://i.meituan.com/poi/<Meituan hotel id>/feedbacks (the id is poiId=number in the address after opening a hotel from the Meituan hotel search results)
//   Open the browser at a phone size (the desktop hotel page is refused).
// Usage (do this for every page, slowly):
//   1. Open page 1, run the whole code, then MeituanReviews.savePage(): reads this page, stores it in the browser, returns the stored count.
//   2. Run MeituanReviews.next(): clicks "next page". After the page opens wait at least 10 more seconds, then repeat step 1 (run the whole code again + savePage).
//   3. On the last page next() returns "already the last page". Run MeituanReviews.dump(0, 50), dump(50, 100)... to take the rows out in parts and save as reviews_meituan_zh.json.
// Note: do not request page addresses in a row by script; Meituan rate-limits (403 refused). If limited, stop, wait a while and go on; what is stored is not lost.
// No identity information is read except the nickname, and the nickname is not saved.
(() => {
  // stored per Meituan hotel id, so analysing two hotels in a row does not mix them
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
      // long reviews are folded in .feedbackmore; textContent reads the folded part too
      text: cm ? cm.textContent.replace(/\s+/g, " ").trim() : "",
    };
  });
  window.MeituanReviews = {
    reset() { localStorage.removeItem(KEY); return "cleared"; },
    summary() {
      const t = document.body.innerText;
      const total = (t.match(/评价\s*(\d+)/) || [])[1];
      const tags = [...t.matchAll(/^([^\n\d]{2,8}) (\d+)$/gm)].map((m) => ({ topic: m[1].trim(), mentions: +m[2] }))
        .filter((x) => !/^(评价|晒图)$/.test(x.topic));
      return { total_reviews: total ? +total : null, topics: tags, note: location.href };
    },
    savePage() {
      if (/403/.test(document.title)) return "rate-limited (403): stop, wait a while, then continue from this page";
      const have = load();
      const seen = new Set(have.map((r) => r.date + r.text.slice(0, 40)));
      const rows = parse().filter((r) => !seen.has(r.date + r.text.slice(0, 40)));
      localStorage.setItem(KEY, JSON.stringify(have.concat(rows)));
      return `${rows.length} new on this page, ${have.length + rows.length} in total`;
    },
    next() {
      const a = [...document.querySelectorAll("a")].find((x) => x.textContent.trim() === "下一页");
      if (!a || /disabled/.test(a.className)) return "already the last page";
      a.click();
      return "clicked next page; wait 10 more seconds after it opens";
    },
    count() { return load().length; },
    dump(from, to) { return JSON.stringify(load().slice(from, to)); },
  };
  return "MeituanReviews ready";
})();
