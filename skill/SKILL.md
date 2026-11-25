---
name: hotel-review-analysis
description: Collect a hotel's own guest reviews from Ctrip/Trip.com, Meituan, Fliggy, Booking.com, Agoda and Expedia and merge them into one easy-to-read review report (web page, can be saved as PDF): conclusions and prioritised actions first, then rating, what guests praise and complain about, main guest groups, foreign-language reviews, low-score reviews one by one, recurring problems and recent change, each finding with a chart. Use when asked to analyse reviews, check a hotel's reputation or collect reviews. Needs an AI tool that can drive a browser and run Python.
---

# Hotel Review Analysis (single hotel, many platforms)

The value is **one combined summary across platforms**, so the owner does not open each OTA or convert scores by hand. Platform-published ratings, sub-scores and guest types are merged (weighted by each platform's published review count); scraped reviews supply the low-score reviews and the 12 complaint themes. Analyse only the user's own hotel. Never analyse or recommend competitors.

**Report language:** the report, Excel and charts are written in English; reviews stay in their original language. Keyword lists in `scripts/analyze.py` cover Chinese and English. To produce another language, translate the labels in `analyze.py`, `charts.py` and `report.py`.

## Step 1: Ask (once)
1. Hotel name and city. 2. Each platform's hotel link if available; otherwise find it (step 2).

## Platforms

| Platform | What is read | Login | Status |
|---|---|---|---|
| Ctrip (full text read from Trip.com) | review text, score, date, traveller type, room type, country, language | No. Trip.com and Ctrip share the same reviews | Author-tested |
| Ctrip mainland site (summary only) | overall score, sub-scores, hot tags and counts | No | Author-tested |
| Booking.com | review text (liked/disliked parts), score, traveller type; sub-scores | No | Author-tested |
| Agoda | review text, score, traveller type; sub-scores; platform topics | No | Author-tested |
| Expedia | overall and sub-scores only | No | The review window cannot be opened automatically, scores only |
| Meituan | review text, stars, date; platform tags; overall score from search results | Yes: the user logs in (phone + code) | Author-tested; paging too fast gets rate-limited |
| Fliggy | review text, 4 sub-score stars, date; totals | Yes: the user logs in with their Alibaba account (QR code) | Author-tested |
| A short-video app | none | Web version has no hotel page | Not included in this version |

If a hotel cannot be found on a platform (common for independent hotels on Booking/Agoda/Expedia), skip it and report "not found this time, to be verified" (ask the hotel whether it is listed). Do not write "not listed". For platforms not covered, write "not included this time".

## Step 2: Find each platform's hotel page
- **Ctrip ID**: take `hotelId=<number>` from a Ctrip link, or search the web for "<hotel name> Ctrip hotels.ctrip.com". Ctrip's own search page needs login; do not use it. Trip.com uses the same ID.
- **Booking.com, Agoda, Expedia**: search by the hotel's English name and open the hotel page.
- **Meituan**: ask the user to log in first. Switch the browser to phone size, open `https://i.meituan.com/awp/h5/hotel/search/search.html`, pick the city, search the name, note the rating (5-point scale), open the hotel; `poiId=<number>` in the URL is the ID. The desktop hotel page refuses access. If the phone login's "this is my account" button does not respond to a mouse click, use `https://passport.meituan.com/account/unitivelogin`.
- **Fliggy**: open `https://hotel.fliggy.com/hotel_list3.htm?cityName=<city>&keywords=<hotel>` (no login), note rating and review count; `shid=<number>` is the ID. Reading reviews needs the user's Alibaba-account login.
- On every hotel page, check name and address. If they do not match, stop and ask.

## Step 3: Collect reviews
Scripts are in `scripts/`: `tripcom.js`, `booking.js`, `agoda.js`, `meituan.js`, `fliggy.js` (usage is at the top of each file).
- **Meituan**: open `https://i.meituan.com/poi/<id>/feedbacks`, **page one page at a time**, waiting at least 10 seconds between pages; never script rapid paging. On a rate limit (403), stop and resume later. If only part is collected, add `"scrape_note": "published N, collected only M, reason"` to that platform in summary.json and do not set all_reviews_scraped.
- **Fliggy**: open `https://hotel.fliggy.com/hotel_detail2.htm?shid=<id>#hotel-review` and run `collect` from `fliggy.js`; it pages slowly by itself. Fliggy reviews have no total score; use the average of the sub-score stars.

### General procedure
1. **The browser window must be visible** (not minimised or covered), or sites return no reviews.
2. Open the hotel page (table below) and run the whole `.js` file in the page.
3. Run `await ReviewScraper.listLanguages()`. Always use the returned `code` as `language`.
4. Collect each language separately: **refresh the page and re-run the `.js` before each language**, then run `collect`.
5. **Save to files**: results can be hundreds of KB, so read them out in slices (`JSON.stringify(r.rows.slice(0, 50))`, …) and save `{"platform", "hotel", "url", "language", "rows"}` as `reviews_<platform>_<language>.json`. Check the saved row count equals `r.rows.length`.

| Platform | Page | Languages | collect parameters |
|---|---|---|---|
| Trip.com | `https://www.trip.com/hotels/detail/?hotelId=<id>&locale=en-XX` | Chinese and English, plus others if present | `{language, maxReviews: 300, maxPages: 40, sort: "newest", pageDelayMs: 3000}` |
| Booking.com | hotel page with the language suffix changed to `.en-gb` (`xxx.zh-cn.html` → `xxx.en-gb.html`; `xxx.html` → `xxx.en-gb.html`) | top 3 languages by count in listLanguages | `{language, maxReviews: 300, sort: "newest", pageDelayMs: 3000}` |
| Agoda | hotel page in the English interface (`/en-gb/`) | Chinese (simplified, HK and TW traditional separately) and English | `{language, maxReviews: 300, maxPages: 60, sort: "newest", pageDelayMs: 3000}` |

**Agoda prerequisites** (both needed): the page must be in the English interface; before running `agoda.js`, scroll down with the **real mouse wheel** 20+ times until the review section appears (script scrolling does not load it).

**Cover a full year.** After each language, check the last review's date. If it is less than a year old and the list is not exhausted, raise `maxReviews` to 600 and double `maxPages`, and collect again. If the list is exhausted, mark the platform in summary.json: `"all_reviews_scraped": ["zh"]` (language codes) or `true` when all languages are done (for Ctrip put it on `Trip.com`). The time comparison checks each language separately and skips languages not covering a year, so "not fully collected" is never read as "fewer reviews".

**Low-score top-up** (skip if all reviews of the platform were collected): collect once more with `sort: "lowest"`, `maxReviews: 50`, in the hotel's main language, saved as `lowscore_<platform>.json`. These are evidence quotes only (shown in the low-score detail with a "source" column) and enter no ratio, count or chart.

Notes: on a login page, a human-check or a blocking pop-up, stop and ask the user to handle it, then continue. Never try to bypass. Do not speed up `pageDelayMs` or re-collect the same platform repeatedly. If collection fails, take a screenshot first; a site redesign means the matching `.js` needs fixing.

## Step 4: Read platform summaries, fill summary.json
Create `summary.json` in the data folder. **Platform names must match the review files exactly**: `Ctrip`, `Trip.com`, `Booking.com`, `Agoda`, `Expedia`, `Meituan`, `Fliggy`. Numbers without quotes. Leave out what the page does not show; never estimate. See `examples/sample_project/summary.json` for the shape.

- **Ctrip**: on the Ctrip mainland hotel page (`https://hotels.ctrip.com/hotels/detail/?hotelId=<id>`) run `scripts/ctrip_summary.js` and put the result under `platforms` (hot tags and the "negative (N)" count `negative_reviews`, used only in Excel and as a weight for merging themes; the report's low-score counts always use the "below 6" rule). Add `hidden_reviews` = Ctrip published count minus the sum of review counts visible on Trip.com, and say in a note that it is computed. In the report write "the published count and the visible count differ by N (may be hidden, folded, or rating-only)"; do not claim they are all hidden.
- **Trip.com**: the English page shows a 10-point score (more precise than Ctrip's 5-point), the review count and 4 sub-scores (Cleanliness, Amenities, Location, Service). When both exist, analyze.py uses Trip.com's score and sub-scores and Ctrip only for tags and negative counts, so nothing is counted twice.
- **Fill ratings and counts completely**: the combined rating and all merges weight by published review counts. Missing `total_reviews` on a scraped platform falls back to the scraped count (flagged); on a score-only platform, it is left out of the merge.
- **Booking.com, Agoda**: traveller-type counts are on the filter chips (e.g. "Families (20)").
- **Agoda's own rating and count only**: "Rating via Agoda" and "From N reviews" / "Agoda Reviews (N)". The header total also includes Booking.com and "Additional reviews" (from Ctrip/Trip.com), which duplicate other platforms and must not be used.
- **Expedia**: search by English name; if the hotel is not in the results (Expedia swaps the query for an address and lists nearby hotels), write "not found this time, to be verified".
- **Scope**: platforms not found go in `not_listed`; platforms this version cannot collect go in `not_included`.
- **Sub-scores**: fill every one the page publishes, with the platform's original name and scale. analyze.py groups them (hygiene, service, facilities, comfort, location, value); those that fit no group (e.g. Ctrip "environment") appear only in the detail.
- **Platform topics**: Ctrip hot tags with counts (the script removes filter items such as "recommended", "with photos", "negative"; check the `excluded` list when changing hotels); Agoda's topic analysis with keys `topic`/`mentions`/`positive_share`/`negative_share`, ratios as decimals (0.9).

## Step 5: Analyse
`pip install -r requirements.txt` (pandas 2.1+, openpyxl, matplotlib; opencc-python-reimplemented if traditional-Chinese reviews should be converted). A CJK font is required (for Chinese review text in charts) (Microsoft YaHei on Windows, Noto Sans CJK on Linux); `charts.py` stops with a message instead of drawing empty boxes. PDF needs Node.js and playwright (`npm i playwright`).

Run `python scripts/analyze.py <data-folder>`. It writes `review_analysis_<hotel>.xlsx` (numbers only; conclusions and charts come in step 6). After changing scripts or keywords run `python tests/test_core.py` and continue only when it prints the all-passed line.

The rules are at the top of `analyze.py`; do not change them ad hoc:
- **Low-score review**: guest score converted to 10-point below 6 (5-point below 3). No good/medium split, because platforms define "good" differently.
- **Merging**: platform numbers weighted by published review count; Ctrip and Trip.com are one set of reviews, counted once. Score-only platforms with a count are included in the overall rating and sub-scores.
- **12 complaint themes** (the 12th is "booking info does not match reality": breakfast included or not, extra charges, photos differ). Share of low-score reviews mentioning each theme, counting only dissatisfied wording: Booking's "disliked" part only; elsewhere reviews are split into clauses and a theme word counts only if a negative wording appears within 6 characters (English: negative words in the same clause). This is automatic and has some false hits. Platforms with fewer than 5 low-score reviews give no theme conclusion and do not take part in merging.
- **Theme impact score**: average score of reviews complaining about a theme minus the average of all others, computed inside each platform, then merged by number of complaints. Correlation, not causation. Fewer than 5 complaints: not counted; gaps below 0.3: not listed. Only normally collected reviews.
- **Sample warnings**: always state counts ("4 of 7 low-score reviews"); sub-scores from under 10% of published reviews are flagged; a period under 30 reviews warns that single reviews weigh a lot.
- **Guest mix**: from traveller types on collected reviews, one review one type, grouped into business, couple, solo, family, friends/group, other. Say "among reviews with a type, X% are family guests"; never present it as the actual guest mix (that needs the PMS).
- **Top-up low-score reviews** never enter any ratio, count, cover number or chart.
- **Charts over time** (timeline, quarterly trend, low-score months) use only platforms and languages marked fully collected.
- **Last 90 days vs 91 days to 1 year**: from collected reviews, only languages covering a full year, each period needing 10+ reviews per platform. Non-participating platforms and reasons are listed in the table.
- **Low-score share by language** is counted on the reviews collected per language, not weighted by each language's real review count.
- **Low-score share** is always "share of analysed reviews below 6", counted among collected reviews with text.

## Step 6: Write conclusions, build the report
Read the data and the low-score originals, form conclusions yourself, write `conclusions.json`, then run `python scripts/report.py <data-folder>` (needs matplotlib). It produces `review_report_<hotel>.html` (fixed order: one-line conclusion and metric cards → at a glance → what to do first → detailed findings with charts → appendix; images embedded, opens offline), the `charts/` PNG folder, and the "conclusions" and "charts" sheets in Excel. The report holds no full low-score originals; at most 1–3 sentences per finding with date and score.

### What to look for (check each; write it only if present)
1. **Overall level**: rating, low-score share, how many low-score reviews in the past year; which platform scores low or has few reviews. The combined rating is a reference only; to judge improvement, watch one platform's own score over time.
2. **What guests praise**: separate basics (service, cleanliness) from selling points (kitchen, laundry, suites, sea, gym, kids' services). Basics show good execution but are not selling points.
3. **Main guest groups and what they care about** (required; this is the basis for the hotel's listing copy): the two largest groups and their shares; what each group mentions more than others (with counts and shares); read their originals for concrete wording; say what to write for each group. Season is a clue only (review date is not stay date); ask to verify with PMS by month. Selling points to put in copy must first be verified as still offered and in which room types.
4. **What guests complain about and why**: use all reviews, not only low-score ones. Read originals to find causes. Theme names are only categories. Every theme that enters a recommendation must have its originals read one by one; wrong hits go into `corrections.json` and the report is rebuilt.
5. **Low-score reviews one by one** (required when fewer than 30): write "cause → what the hotel did → result" and look for the same failure chain (e.g. broken room → no room to move to → guest left). Do not write as fact what the review does not show.
6. **When low-score reviews cluster**: only with 20+ low-score reviews; otherwise no month pattern.
7. **Is a problem recurring**: same type across several half-years means not fixed. Compute yearly share, and say "every year" is not "getting worse". Repairs: add "record by room number".
8. **Recent change**: quarterly trend and the 90-day comparison; with few reviews say "one or two reviews, not yet a trend".
9. **Channels**: differences in review counts; where foreign-language reviews come from.
10. **Room types**: a room type with clearly lower scores and the same complaint repeated.
11. **Booking info vs reality**: cheap and fast to fix on the platform page.
Two groups' problems may differ: compare their average scores, complaint shares and themes; check whether the most common room type has the lowest score and whether that is the room itself. Add nearby place names (mall, beach, zone) to `feature_keywords` in summary.json, or they are not counted.
Foreign-language reviews are a separate finding: language is not nationality; read originals (translate Korean/Japanese into `translations.json`); results are clues, not a basis for market investment. If many foreign residents live nearby (foreign-owned factories, expat communities) but online foreign-language reviews are few, say so: they may book through corporate agreements or other channels.

### conclusions.json
Keep it short: one-line headline (max 16 words), 2–3 "watch" cards with the same metrics (reviews mentioning, their average score), actions with `when` (Now / First / Later / Verify first), a short action, `basis` (number of reviews), bullet `detail`, one-line `why`; findings with a title, a short detail, optional `stats` and `evidence`, and `charts` (file names without .png). See `examples/sample_project/conclusions.json`. Donut charts are only for shares that add to 100%.

**Order of findings** (skip what does not apply): overall reputation and platforms → what guests praise → main guest groups → foreign-language reviews → complaints, which hurts the score most, whether it recurs → recent change → individual incidents (few reviews, verify first) → channels.

**Chart names** for `charts`: 1_rating_overview, 2_complaints, 3_problems_over_time, 4_quarterly_trend, 5_low_score_months, 6_guest_mix, 7_themes_lowering_score, 8_platform_tags, 9_praise, 10_guest_group_mentions, 11_foreign_language_mentions. Give each chart to one finding only. "At a glance" already lists praise and complaints, so 9_ and 2_ are rarely needed again.

### Writing rules
- Conclusion first, numbers after, short sentences. Every number must be checkable in Excel or the originals. Quote originals with date and score; never reviewer nicknames.
- One decimal at most; percentages as integers; one denominator throughout (analysed reviews, written "308 reviews, 11%"; no mixing "about half" with "49%"); one name per problem, matching the charts; one metric wording for how serious a problem is ("reviews complaining about it average X (all reviews average Y)"); units always written (9.3 points, 431 reviews); no "small sample" label inside charts (state the count in text instead).
- Ctrip and Trip.com are one set of reviews: write only "Ctrip" in the report.
- **Recommendation thresholds**: "first/later" needs at least 5 complaints about the same problem not concentrated in one or two stays (or a large-sample praise). Safety, privacy, room-status errors (e.g. guest moved into an occupied room) and guests forced to leave are "now" even with one review. Under 5 → "verify first" with the internal record to check (PMS room status, maintenance tickets, complaint log, booking volume). Under 20 low-score reviews: never decide priorities from their distribution alone.
- **Price**: do not write unverified rules such as "all platforms require price parity". Write "compare direct price and platform price for the same date, room type and cancellation terms under the hotel's current contracts and price policies".
- Damp and smell are not always cleaning: also check linen drying and storage, ventilation, A/C and drains.
- Do not over-infer: review date is not stay date; many praising reviews naming staff may mean staff ask for reviews; one extreme review can drag a quarter; few reviews on a platform → check that channel's bookings first; small differences between groups are "indicative only"; praise shows what guests liked, not what brings bookings; when something is both praised and complained about, write both sides; a theme impact score is correlation.
- Only this hotel; no competitor comparison.

### Independent review before delivery (required)
Give the report to a reviewer who did not write it (a new subagent or a new window), with only the folder path, the report and the checklist. It marks each item "pass" or "problem + where", and only finds problems. Fix, rebuild, have it re-check. If no reviewer is possible, self-check and tell the user "self-checked, no independent review".
Checklist: tests pass; every number traceable and denominators stated; wording consistent (theme names, integer percentages, "low score" = below 6, "Ctrip" only, "foreign-language reviews" not nationality); recommendation thresholds met; themes in recommendations hand-checked; no over-inference; each finding has a chart or number and no walls of text; no nicknames, no competitors, no top-up reviews in any ratio; report opens at desktop and phone widths without horizontal scroll; every PDF page checked.

### Deliver
1. Deliver **one complete report** (do not send Excel or the skill files separately) (HTML; publish it as a shareable page if the tool can). Excel stays in the data folder unless asked.
2. PDF on request: `node scripts/to_pdf.js <absolute path to report.html> <output.pdf>`; check every page.
3. Reply with the one-line conclusion and what to do first; do not paste the report.

## Never
Bypass login, CAPTCHA or bot checks; collect at high frequency; write reviewer nicknames; analyse or recommend competitors.
