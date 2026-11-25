/*
 * Trip.com review collector.
 *
 * Load it into a Trip.com hotel page in English
 * (https://www.trip.com/hotels/detail/?hotelId=<id>&locale=en-XX), either through
 * the AI browser tool or the developer console, then call:
 *
 *   await ReviewScraper.listLanguages()
 *       -> [{code: "zh", label: "Chinese", count: null}, ...]
 *   await ReviewScraper.collect({language: "zh", maxReviews: 40, sort: "newest", pageDelayMs: 3000})
 *       -> {platform, hotel, url, language, rows: [...]}
 *
 * One call = one language, on a freshly loaded page (reload the page before each
 * language). Trip.com shows an AI translation into English by default; the collector clicks
 * "Original Text" on every card before reading it, so the text kept is the original.
 * Trip.com's class names are scrambled, so cards are found by their text ("Posted <date>").
 */
(() => {
  const PLATFORM = "Trip.com";
  const SORTS = { newest: "Most Recent", relevant: "Most relevant", lowest: "Rating (low to high)", highest: "Rating (high to low)" };
  const LANGUAGE_CODES = {
    English: "en", Chinese: "zh", Korean: "ko", Japanese: "ja", French: "fr", German: "de", Spanish: "es", Italian: "it",
    Russian: "ru", Thai: "th", Indonesian: "id", Malay: "ms", Vietnamese: "vi", Dutch: "nl", Turkish: "tr",
    Portuguese: "pt", Arabic: "ar", Finnish: "fi", Polish: "pl", Danish: "da", Swedish: "sv", Norwegian: "no",
    Greek: "el", Hebrew: "he", Czech: "cs", Hungarian: "hu", Ukrainian: "uk", Filipino: "fil", Hindi: "hi",
  };
  const TRAVELER_TYPES = ["Business traveler", "Family", "Solo traveler", "Booked for others", "Couple", "Traveling with friends", "Other"];

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const jitter = (ms) => ms + Math.round(Math.random() * 1000);
  const T = (e) => (e?.textContent || "").replace(/\s+/g, " ").trim();
  const leaves = (root = document) => [...root.querySelectorAll("*")].filter((e) => !e.children.length);
  const leaf = (text, root = document) => leaves(root).filter((e) => T(e) === text);
  const click = (el) => {
    for (const type of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
      el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
    }
  };

  async function waitFor(check, timeoutMs = 15000) {
    for (let t = 0; t < timeoutMs; t += 300) {
      const value = check();
      if (value) return value;
      await sleep(300);
    }
    return null;
  }

  function cards() {
    const found = new Set();
    for (const posted of leaves()) {
      if (!/^Posted [A-Z][a-z]+ \d{1,2}, \d{4}$/.test(T(posted))) continue;
      if (posted.closest('[class*="sceneReviewSwiper"]')) continue;  // the highlights carousel, not the list
      let card = posted;
      while (card && !/Stayed in/.test(T(card))) card = card.parentElement;
      if (card && T(card).length < 10000) found.add(card);
    }
    return [...found];
  }
  const firstCard = () => T(cards()[0]).slice(0, 120);

  async function openReviews() {
    if (leaf("All Languages").length) return;
    const allReviews = () => [...document.querySelectorAll("*")].filter((e) => /^All [\d,]+ reviews$/.test(T(e))).at(-1);
    if (!(await waitFor(allReviews))) throw new Error("No 'All N reviews' button: is this a Trip.com hotel page in English (locale=en-XX)?");
    // The button ignores clicks until the page has fully started, so wait and retry.
    for (let attempt = 0; attempt < 4 && !leaf("All Languages").length; attempt++) {
      await sleep(2500);
      const button = allReviews();
      button.click();
      click(button);
      await waitFor(() => leaf("All Languages").length, 6000);
    }
    if (!leaf("All Languages").length) throw new Error("The review list did not open");
    await sleep(1500);
  }

  function optionsBox() {
    // The open language dropdown: the smallest element holding several known language names.
    const english = leaf("English").at(-1);
    let box = english;
    while (box && leaves(box).filter((e) => LANGUAGE_CODES[T(e)]).length < 3) box = box.parentElement;
    return box;
  }

  async function listLanguages() {
    await openReviews();
    const trigger = leaf("All Languages")[0];
    click(trigger);
    await waitFor(() => optionsBox(), 5000);
    const box = optionsBox();
    const labels = box ? [...new Set(leaves(box).map(T).filter((t) => t && t !== "All Languages"))] : [];
    click(trigger);  // close the dropdown again
    await sleep(800);
    return labels.map((label) => ({ code: LANGUAGE_CODES[label] || label.toLowerCase(), label, count: null }));
  }

  async function choose(triggerText, optionText) {
    const trigger = leaf(triggerText)[0];
    if (!trigger) return false;
    const before = firstCard();
    click(trigger);
    await sleep(1200);
    const option = leaf(optionText).at(-1);
    if (!option) return false;
    click(option);
    await waitFor(() => firstCard() !== before, 10000);
    await sleep(1500);
    return true;
  }

  async function showOriginals() {
    for (const button of leaf("Original Text")) {
      click(button);
      await sleep(300);
    }
    await sleep(800);
  }

  function parseCard(card, language) {
    const L = leaves(card).map(T).filter(Boolean);
    const typeAt = L.findIndex((x) => TRAVELER_TYPES.includes(x));
    const postedAt = L.findIndex((x) => /^Posted /.test(x));
    const responseAt = L.findIndex((x) => /^Response from Property/.test(x));
    const body = L.slice(postedAt + 1, responseAt > 0 ? responseAt : undefined)
      .filter((x) => x.length > 1 && !/^(Show (More|Less)|Translate|Original Text|Helpful|\d+\/\d+)$/.test(x));
    return {
      platform: PLATFORM,
      review_language: language.code,
      review_language_label: language.label,
      score: (T(card).match(/(\d+(?:\.\d)?)\/10/) || [])[1] || "",
      score_scale: 10,
      date: (L[postedAt] || "").replace("Posted ", ""),
      traveler_type: typeAt >= 0 ? L[typeAt] : "",
      room_type: typeAt > 0 ? L[typeAt - 1] : "",
      reviewer_country: typeAt > 1 ? L[typeAt - 2] : "",
      stay: L.find((x) => /^Stayed in/.test(x)) || "",
      title: "",
      text: body.sort((a, b) => b.length - a.length)[0] || "",
      text_is_original: !leaf("Original Text", card).length,
    };
  }

  async function collect({ language = "en", maxReviews = 40, sort = "newest", pageDelayMs = 3000, maxPages = 20 } = {}) {
    await openReviews();
    const languages = await listLanguages();
    const lang = languages.find((l) => l.code === language || l.label === language);
    if (!lang) throw new Error(`Language '${language}' is not in this hotel's language filter`);
    if (!(await choose("All Languages", lang.label))) throw new Error(`Could not select language ${lang.label}`);
    const sortTrigger = leaves().find((e) => /^Sort by:/.test(T(e)));
    if (sortTrigger && sort !== "relevant") {
      click(sortTrigger);
      await sleep(1000);
      const option = leaf(SORTS[sort] || SORTS.newest).at(-1);
      if (option) {
        const before = firstCard();
        click(option);
        await waitFor(() => firstCard() !== before, 10000);
        await sleep(1500);
      }
    }

    const rows = [];
    const seen = new Set();
    const grab = async () => {
      await showOriginals();
      for (const card of cards()) {
        const row = parseCard(card, lang);
        const key = [row.date, row.text.slice(0, 80)].join("|");
        if (!row.text || seen.has(key)) continue;
        seen.add(key);
        rows.push(row);
      }
    };
    // The page bar shows "1 2 3 4 5 ... 349": click page n if shown, otherwise the "next" arrow
    // (the last, text-less item of the bar).
    const pager = (n) => {
      const numbers = [...document.querySelectorAll("li")]
        .filter((e) => !e.children.length && /^\d+$/.test(T(e)))
        .filter((e) => [...e.parentElement.querySelectorAll("li")].map(T).includes("1"));
      const shown = numbers.find((e) => T(e) === String(n));
      if (shown) return shown;
      const items = numbers.length ? [...numbers[0].parentElement.querySelectorAll(":scope > li")] : [];
      const last = items.at(-1);
      return last && !T(last) ? last : null;
    };

    await grab();
    for (let page = 2; page <= maxPages && rows.length < maxReviews; page++) {
      const next = pager(page);
      if (!next) break;
      await sleep(jitter(pageDelayMs));
      const before = firstCard();
      click(next);
      if (!(await waitFor(() => firstCard() !== before, 10000))) break;
      await sleep(1000);
      await grab();
    }
    return { platform: PLATFORM, hotel: hotelName(), url: location.href.split(/&checkIn|#/)[0], language: lang, rows: rows.slice(0, maxReviews) };
  }

  function hotelName() {
    return T(document.querySelector("h1")) || document.title.split(/[,(|-]/)[0].trim();
  }

  window.ReviewScraper = { platform: PLATFORM, listLanguages, collect };
})();
