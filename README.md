# Hotel Review Analysis

**An open-source skill that turns a hotel's guest reviews from many booking platforms into one easy-to-read report.** It collects your own hotel's reviews (Ctrip/Trip.com, Meituan, Fliggy, Booking.com, Agoda; Expedia scores only), merges every platform into one summary, and writes a web page (saveable as PDF) with the conclusion and what to do first at the top, then charts: overall rating, what guests praise and complain about, main guest groups, foreign-language reviews, low-score reviews one by one, recurring problems and recent change.

![report preview](examples/sample_project/report_preview.png)

*The sample above uses invented data for a fictional hotel. The report, Excel and charts are in English; review text stays in its original language.*

## What it does

- **One combined rating** weighted by each platform's published review count (Ctrip and Trip.com are the same reviews, counted once)
- **Complaint themes**: 12 fixed themes with Chinese and English keywords (noise, repairs, hygiene, staff, breakfast, "booking info does not match reality" and more), counted only where guests are unhappy
- **What guests praise**, with basics (service, cleanliness) separated from real selling points
- **Main guest groups** and what each group mentions most, as the basis for listing copy
- **Low-score reviews one by one** (cause, what the hotel did, result), and whether a problem keeps coming back
- **Recent change**: last 90 days against 91 days to one year ago, with sample-size warnings
- **Actions ranked** as now / first / later / verify first, each with the number of reviews behind it
- Single hotel only. It does not analyse or recommend competitors

Design choices: low score = below 6 out of 10 (below 3 out of 5), no "good/medium" split; top-up low-score reviews are only evidence quotes and never enter a ratio or chart; no conclusion without a number you can check in Excel; no reviewer nicknames.

## Quick start (no account needed)

```bash
pip install -r requirements.txt            # pandas, openpyxl, matplotlib (a CJK font is needed so Chinese review text and room types show in charts)
python tests/test_core.py                  # should end with the all-passed line
python scripts/analyze.py examples/sample_project
python scripts/report.py  examples/sample_project
# open examples/sample_project/review_report_*.html in your browser
```

PDF (optional): `npm i playwright`, then `node scripts/to_pdf.js <absolute path to the .html> <output.pdf>`.

With your own hotel: let an AI assistant that can drive a browser and run Python follow `skill/SKILL.md`. It finds your hotel on each platform, collects the reviews with the scripts in `scripts/`, writes `summary.json` and `conclusions.json`, and builds the report. Keep your data folder outside this repository. Data formats are in `docs/DATA_FORMAT.md`.

## Platforms

| Platform | What is read | Status |
|---|---|---|
| Ctrip (full text read from Trip.com, summary from the mainland site) | reviews, scores, sub-scores, tags | Author-tested |
| Booking.com | reviews, sub-scores | Author-tested |
| Agoda | reviews, sub-scores, platform topics | Author-tested |
| Meituan | reviews, stars, tags (you log in yourself) | Author-tested, rate-limited if paged fast |
| Fliggy | reviews, sub-score stars (you log in yourself) | Author-tested |
| Expedia | scores only | Reviews cannot be opened automatically |
| A short-video app | none | Not included (no hotel page on the web version) |

If your hotel is not on a platform, the report says "not found this time, to be verified" (ask the hotel whether it is listed).

## Language

Review keywords cover Chinese and English. The report, Excel and charts are written in English; the Chinese keyword lists in `scripts/analyze.py` are needed to read Chinese reviews. To produce another output language, translate the labels in `scripts/analyze.py`, `scripts/charts.py` and `scripts/report.py`. Pull requests are welcome.

## Risks, in plain words

Collecting reviews automatically may be restricted by a platform's terms, and platforms may limit or block accounts that load pages too fast. The scripts stop at a login page or human check and never try to bypass it; you log in yourself and you decide what to collect and how often. You are responsible for following the terms of any site you use. The authors give no warranty (see `LICENSE`).

## Contributing

Pull requests are welcome: fixes when a platform changes its page, other platforms, translations, report improvements. Read `CONTRIBUTING.md`. Questions and ideas: open an issue.

## About the author

**Oriental Horizon Hotel Services** is a hotel management and consulting company. The founder is a hotel revenue
leader with 13 years in international hotel groups (Hilton, IHG, Accor, Hyatt) across Southeast Asia and China, a former
Director of Revenue at a five-star hotel in Singapore, and the head of a company that turned a loss-making
four-star hotel in a tier-3 city into a profitable one.

**Who we help**: hotel owners and hotel professionals (revenue, sales, marketing and general managers) who face a problem at work, especially at international brand and group hotels, where our own experience runs deepest.

**What we can help with**

- **Hotel business consulting**: how to get your hotel's business going and growing. We find where the business is strong and weak and what to do about it
- **OTA channel consulting**: how your hotel sells on the booking platforms, and how to improve it
- **AI for hotels**: we help your team put AI into daily operations, through training or by building the tools for you
- **Hands-on coaching**: we work alongside you and your team over time until the results show, not just hand over a report
- **Training**: courses on hotel business and AI, open to owners, hotel staff and students

**Free consultation: tell us your situation and we will tell you honestly whether we can help.**

- LinkedIn: https://www.linkedin.com/in/ludan-zhang54/
- WeChat: Oriental_Horizon
- WhatsApp: +65 86796635
- Email: ludanzhang54@gmail.com
- YouTube (Chinese-language videos and podcast): https://www.youtube.com/@TheHotelBusiness
- Podcast, English: *The Hotel Business* - https://rss.com/podcasts/the-hotel-business/
- Podcast, Chinese: *张小姐看酒店生意* ("Hotel Business with Ms. Zhang") - https://www.youtube.com/playlist?list=PLK66NvWFLv9U

## Licence

MIT. Free to use, change and share, including commercially. See `LICENSE`.
