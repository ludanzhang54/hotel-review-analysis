/*
 * Booking.com review collector.
 *
 * Load it into a Booking.com hotel page (https://www.booking.com/hotel/hk/<hotel>.en-gb.html),
 * by running it in the browser (AI browser tool or developer console), then call:
 *
 *   await ReviewScraper.listLanguages()
 *       -> [{code: "zh", label: "Chinese", count: 124}, ...]
 *   await ReviewScraper.collect({language: "zh", maxReviews: 40, sort: "newest", pageDelayMs: 3000})
 *       -> {platform, hotel, url, language, rows: [...]}
 *
 * One call = one language. Booking.com shows each review in its original language (with a
 * "Show translation" button), so the text kept here is the original. Reviews with no text
 * (score only) are skipped.
 */
(() => {
  const PLATFORM = "Booking.com";
  const SORTS = { newest: "NEWEST_FIRST", relevant: "MOST_RELEVANT", lowest: "SCORE_ASC", highest: "SCORE_DESC" };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const jitter = (ms) => ms + Math.round(Math.random() * 1000);
  const T = (e) => (e?.textContent || "").replace(/\s+/g, " ").trim();
  const q = (card, testid) => T(card.querySelector(`[data-testid="${testid}"]`));

  async function waitFor(check, timeoutMs = 15000) {
    for (let t = 0; t < timeoutMs; t += 300) {
      const value = check();
      if (value) return value;
      await sleep(300);
    }
    return null;
  }

  async function openReviewPanel() {
    const panelOpen = () => document.querySelector('select[name="languages"]');
    if (panelOpen()) return;
    const buttons = () => [...document.querySelectorAll("button")].filter((b) => T(b) === "Read all reviews");
    if (!(await waitFor(() => buttons().length))) {
      throw new Error("No 'Read all reviews' button: is this a Booking.com hotel page, in English (.en-gb.html)?");
    }
    // The button ignores clicks until the page has fully started, so wait and retry.
    for (let attempt = 0; attempt < 4 && !panelOpen(); attempt++) {
      await sleep(2500);
      buttons().forEach((b) => b.click());
      await waitFor(panelOpen, 6000);
    }
    if (!panelOpen()) throw new Error("The review panel did not open");
  }

  const select = (name) => document.querySelector(`select[name="${name}"]`);
  const cardsSignature = () => [...document.querySelectorAll('[data-testid="review"]')].slice(0, 3).map(T).join("|").slice(0, 300);

  async function setSelect(name, value) {
    const el = select(name);
    if (!el || el.value === value) return;
    const before = cardsSignature();
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value").set.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
    await waitFor(() => cardsSignature() !== before, 10000);
    await sleep(1500);
  }

  async function listLanguages() {
    await openReviewPanel();
    return [...select("languages").options]
      .filter((o) => o.value && o.value !== "0")
      .map((o) => {
        const m = o.text.match(/^(.*?)\s*\(([\d,]+)\)$/);
        return { code: o.value, label: m ? m[1] : o.text, count: m ? Number(m[2].replace(/,/g, "")) : null };
      });
  }

  function parseCard(card, language) {
    const avatar = card.querySelector('[data-testid="review-avatar"]');
    const avatarLeaves = avatar ? [...avatar.querySelectorAll("*")].filter((e) => !e.children.length).map(T).filter(Boolean) : [];
    const liked = q(card, "review-positive-text");
    const disliked = q(card, "review-negative-text");
    return {
      platform: PLATFORM,
      review_language: language.code,
      review_language_label: language.label,
      score: (q(card, "review-score").match(/[\d.]+/) || [""])[0],
      score_scale: 10,
      date: q(card, "review-date").replace("Reviewed: ", ""),
      traveler_type: q(card, "review-traveler-type"),
      room_type: q(card, "review-room-name"),
      reviewer_country: avatarLeaves.length > 1 ? avatarLeaves.at(-1) : "",
      stay: q(card, "review-stay-date"),
      title: q(card, "review-title"),
      text_liked: liked,
      text_disliked: disliked,
      text: [liked, disliked].filter(Boolean).join(" | "),
      text_is_original: true,
    };
  }

  async function collect({ language = "en", maxReviews = 40, sort = "newest", pageDelayMs = 3000 } = {}) {
    await openReviewPanel();
    const languages = await listLanguages();
    const lang = languages.find((l) => l.code === language) || { code: language, label: language, count: null };
    if (!languages.some((l) => l.code === language)) throw new Error(`Language '${language}' is not in this hotel's language filter`);
    await setSelect("languages", language);
    await setSelect("reviewListSorters", SORTS[sort] || SORTS.newest);

    const rows = [];
    const seen = new Set();
    const grab = () => {
      for (const card of document.querySelectorAll('[data-testid="review"]')) {
        const row = parseCard(card, lang);
        const key = [row.date, row.title, row.text.slice(0, 80)].join("|");
        if (seen.has(key) || !row.text) continue;
        seen.add(key);
        rows.push(row);
      }
    };

    grab();
    for (let page = 2; rows.length < maxReviews; page++) {
      const next = [...document.querySelectorAll("button[aria-label]")].find((b) => b.getAttribute("aria-label") === `Page ${page}`);
      if (!next) break;
      await sleep(jitter(pageDelayMs));
      const before = cardsSignature();
      next.click();
      if (!(await waitFor(() => cardsSignature() !== before, 12000))) break;
      await sleep(800);
      grab();
    }
    return { platform: PLATFORM, hotel: hotelName(), url: location.href.split(/[?#]/)[0], language: lang, rows: rows.slice(0, maxReviews) };
  }

  function hotelName() {
    return document.title.split(",")[0].trim();  // "<Hotel name>, <City> (updated prices ...)"
  }

  window.ReviewScraper = { platform: PLATFORM, listLanguages, collect };
})();
