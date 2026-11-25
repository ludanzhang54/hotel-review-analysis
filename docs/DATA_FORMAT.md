# Data format

All files live in one data folder per hotel (keep it outside this repository).

## `reviews_<platform>_<language>.json` (one file per platform and language)

```json
{"platform": "Trip.com", "hotel": "Hotel name", "url": "", "language": {"code": "zh"},
 "rows": [{"platform": "Trip.com", "date": "2026-08-12", "score": 9.0, "score_scale": 10, "text": "review text", "title": "",
           "traveler_type": "Couple", "room_type": "Standard King", "reviewer_country": "",
           "text_disliked": null}]}
```

- `platform` must be exactly one of `Ctrip`, `Trip.com`, `Booking.com`, `Agoda`, `Expedia`, `Meituan`, `Fliggy` (the platform names are used as keys).
- `score` on the platform's own scale; `score_scale` is that scale (5 or 10).
- `text_disliked`: Booking.com's separate "dislike" paragraph, if the platform has one.

## `lowscore_<platform>.json` (optional)
Lowest-score reviews fetched separately. Used only as evidence quotes, never in any ratio, count or chart.

## `summary.json`
Platform-level numbers read from the platform pages (rating, scale, published review count, sub-scores, traveller counts, platform topics).
See `examples/sample_project/summary.json` and `skill/SKILL.md`, step 4.

## `conclusions.json`, `corrections.json`, `translations.json`
Written by the AI assistant or the reviewer: findings and actions for the report, manual corrections of complaint themes, and Chinese translations of Korean or Japanese reviews. Formats are in `skill/SKILL.md` and in the docstrings of `scripts/analyze.py`.
