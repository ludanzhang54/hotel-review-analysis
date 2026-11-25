/*
 * Agoda review collector.
 *
 * Load it into an Agoda hotel page (https://www.agoda.com/<hotel>/hotel/hong-kong-hk.html),
 * by running it in the browser (AI browser tool or developer console), then call:
 *
 *   await ReviewScraper.listLanguages()
 *       -> [{code: "zh-CN", label: "Simplified Chinese (as the page shows it)", count: null}, ...]
 *   await ReviewScraper.collect({language: "zh-CN", maxReviews: 40, sort: "newest", pageDelayMs: 3000})
 *       -> {platform, hotel, url, language, rows: [...]}
 *
 * One call = one language of Agoda's language filter. With a language selected, Agoda shows the
 * original text; if a card still offers "Show original" (the text shown is a translation), the
 * collector clicks it first. Agoda often shows only a few reviews per page after filtering, and
 * some pages are empty, so it keeps paging (up to maxPages) until it has maxReviews.
 */
(() => {
  const PLATFORM = "Agoda";
  const SORTS = { newest: "1", highest: "2", lowest: "3", relevant: "7" };
  // Agoda's language filter shows names; map them to language codes.
  const LANGUAGE_CODES = {
    "English": "en", "繁體中文 (香港)": "zh-HK", "简体中文": "zh-CN", "繁體中文 (台灣)": "zh-TW", "한국어": "ko",
    "日本語": "ja", "ภาษาไทย": "th", "Bahasa Indonesia": "id", "Deutsch": "de", "Русский": "ru", "Français": "fr",
    "Tiếng Việt": "vi", "Español": "es", "Português (PT)": "pt", "Svenska": "sv", "العربية": "ar", "Italiano": "it",
    "Suomi": "fi", "עברית": "he", "Nederlands": "nl", "Bahasa Malaysia": "ms", "Türkçe": "tr", "Norsk": "no",
    "Português (BR)": "pt-BR", "Español / México": "es-MX", "Ελληνικά": "el", "Język polski": "pl", "Dansk": "da",
  };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const jitter = (ms) => ms + Math.round(Math.random() * 1000);
  const T = (e) => (e?.textContent || "").replace(/\s+/g, " ").trim();
  const E = (name) => document.querySelector(`[data-element-name="${name}"]`);
  // The filters also exist as non-<select> elements in some layouts; always use the <select>.
  const S = (name) => document.querySelector(`select[data-element-name="${name}"]`);

  async function waitFor(check, timeoutMs = 15000) {
    for (let t = 0; t < timeoutMs; t += 300) {
      const value = check();
      if (value) return value;
      await sleep(300);
    }
    return null;
  }

  function codeFor(label) {
    if (LANGUAGE_CODES[label]) return LANGUAGE_CODES[label];
    const english = label.match(/^English \/ (.+)$/);  // "English / Singapore" etc.
    return english ? `en (${english[1]})` : label;
  }

  async function openReviews() {
    // The review section loads lazily when it scrolls into view (scroll the page
    // with the mouse wheel first, which Agoda needs in some page versions).
    for (let i = 0; i < 40 && !(E("reviews-language-filter") && E("review-paginator-next")); i++) {
      window.scrollBy(0, 1500);
      await sleep(500);
    }
    if (!E("reviews-language-filter")) throw new Error("No review section found: is this an Agoda hotel page?");
    E("review-section")?.scrollIntoView();
    await sleep(1500);
  }

  // Agoda serves two page versions: the language filter is either a <select> or a button that
  // opens a list of [role=option] items. These helpers handle both.
  async function openLanguageMenu() {
    const button = E("reviews-language-filter");
    button.scrollIntoView({ block: "center" });
    await sleep(400);
    button.click();
    return await waitFor(() => {
      const options = [...document.querySelectorAll('[role="option"]')];
      return options.length ? options : null;
    }, 5000) || [];
  }

  async function languageLabels() {
    const select = S("reviews-language-filter");
    if (select) return [...select.options].map((o) => o.text.trim());
    const labels = (await openLanguageMenu()).map(T);
    E("reviews-language-filter").click();  // close the menu again
    await sleep(500);
    return labels;
  }

  async function chooseLanguage(label) {
    const select = S("reviews-language-filter");
    if (select) return setSelect(select, label);
    const option = (await openLanguageMenu()).find((o) => T(o) === label);
    if (!option) throw new Error(`Language '${label}' not found in the menu`);
    option.click();
    await sleep(4000);
  }

  const pageNumber = () => (T(E("review-paginator-step")?.parentElement).match(/review page (\d+)/) || [])[1];

  async function setSelect(el, value) {
    if (!el || el.value === value) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value").set.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
    await sleep(4000);
  }

  async function listLanguages() {
    await openReviews();
    return (await languageLabels())
      .filter((label) => label && !/^(Select your country|All languages)$/.test(label))
      .map((label) => ({ code: codeFor(label), label, count: null }));
  }

  async function showOriginals() {
    for (const control of document.querySelectorAll('[data-element-name="review-comment-translation"]')) {
      if (/Show original/i.test(T(control))) {
        (control.querySelector("button") || control).click();
        await sleep(400);
      }
    }
  }

  function parseCard(card, language) {
    const reviewer = [...card.querySelectorAll(".Review-comment-reviewer")].map(T);
    const right = T(card.querySelector(".Review-comment-right"));
    const translation = T(card.querySelector('[data-element-name="review-comment-translation"]'));
    return {
      platform: PLATFORM,
      review_language: language.code,
      review_language_label: language.label,
      score: (T(card.querySelector(".Review-comment-left")).match(/^[\d.]+/) || [""])[0],
      score_scale: 10,
      date: (right.match(/Reviewed ([A-Za-z]+ \d{1,2}, \d{4})/) || [])[1] || "",
      traveler_type: reviewer[1] || "",
      room_type: reviewer[2] || "",
      reviewer_country: (reviewer[0] || "").split(" from ").slice(1).join(" from "),
      stay: reviewer[3] || "",
      title: (right.match(/^“([^”]*)”/) || [])[1] || "",
      text: T(card.querySelector('[data-selenium="comment"]')),
      text_is_original: !/Show original/i.test(translation),
      _dedup: reviewer[0] || "",  // used only to de-duplicate; removed before output (no reviewer nicknames)
    };
  }

  async function collect({ language = "en", maxReviews = 40, sort = "newest", pageDelayMs = 3000, maxPages = 40 } = {}) {
    await openReviews();
    const languages = await listLanguages();
    const lang = languages.find((l) => l.code === language || l.label === language);
    if (!lang) throw new Error(`Language '${language}' is not in this hotel's language filter`);
    await chooseLanguage(lang.label);
    await setSelect(S("reviews-sort-dropdown"), SORTS[sort] || SORTS.newest);

    const rows = [];
    const seen = new Set();
    const grab = async () => {
      await showOriginals();
      for (const left of document.querySelectorAll(".Review-comment-left")) {
        const row = parseCard(left.parentElement, lang);
        const key = [row._dedup, row.date, row.text.slice(0, 80)].join("|");
        if (!row.text || seen.has(key)) continue;
        seen.add(key);
        rows.push(row);
      }
    };

    await grab();
    for (let page = 1; page < maxPages && rows.length < maxReviews; page++) {
      const next = E("review-paginator-next");
      if (!next || next.disabled) break;
      await sleep(jitter(pageDelayMs));
      const before = pageNumber();
      next.click();
      if (!(await waitFor(() => pageNumber() !== before, 10000))) break;
      await sleep(1500);
      await grab();
    }
    return { platform: PLATFORM, hotel: hotelName(), url: location.href.split(/[?#]/)[0], language: lang, rows: rows.slice(0, maxReviews) };
  }

  function hotelName() {
    return T(document.querySelector('[data-selenium="hotel-header-name"], h1')) || document.title.split(",")[0].trim();
  }

  window.ReviewScraper = { platform: PLATFORM, listLanguages, collect };
})();
