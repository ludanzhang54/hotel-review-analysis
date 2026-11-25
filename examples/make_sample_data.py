"""Generate the fictional sample project (no real hotel, no real guests).

Usage: python examples/make_sample_data.py examples/sample_project
Every review below is invented for demonstration only.
"""
import json
import random
import sys
from datetime import date, timedelta
from pathlib import Path

random.seed(7)
out = Path(sys.argv[1] if len(sys.argv) > 1 else "examples/sample_project")
out.mkdir(parents=True, exist_ok=True)
END = date(2026, 11, 1)

ZH_GOOD = ["早餐很丰富，种类多", "房间干净整洁，服务很好", "位置很方便，走路就到海边", "前台热情，办理入住很快",
           "房间宽敞，床很舒服", "有洗衣烘干机，带孩子住很方便", "有厨房，住几天很方便", "健身房设备不错，出差很合适",
           "整体不错，下次还会再来", "性价比很高，推荐"]
ZH_BAD = ["空调有点吵，晚上睡不好", "热水忽冷忽热，洗澡不舒服", "房间有点旧，需要翻新", "隔音一般，走廊很吵",
          "早餐没有写清楚不含，到店才说另外收费", "前台态度一般，问题没人理", "卫生间下水慢，有异味"]
EN_GOOD = ["Great breakfast and friendly staff.", "Clean room, comfortable bed, very close to the beach.",
           "Kitchen and laundry made our family stay easy.", "Quick check-in and a helpful front desk.",
           "Good gym, perfect for a business trip."]
EN_BAD = ["The air conditioning was noisy and the shower water kept changing temperature.",
          "Room felt dated and a bit worn.", "Breakfast was not included although the listing suggested it, extra charge at check-in.",
          "Thin walls, very noisy corridor at night."]
TYPES_ZH = ["商务出行", "家庭出游", "情侣出游", "独自旅行", "朋友出游"]
TYPES_EN = ["Business traveller", "Family with young children", "Couple", "Solo traveller", "Group of friends"]
ROOMS = ["Standard King", "Family Suite", "Deluxe Twin"]


def make(n, platform, lang, good, bad, types, scale, p_bad=0.14):
    rows = []
    for i in range(n):
        d = END - timedelta(days=int(random.random() ** 1.3 * 380))
        is_bad = random.random() < p_bad
        t = random.choice(bad if is_bad else good)
        if not is_bad and random.random() < 0.3:
            t += ("，" if lang == "zh" else " ") + random.choice(good)
        s10 = random.choice([3.5, 4.5, 5.0, 5.5] if is_bad else [8.0, 9.0, 9.5, 10.0, 10.0])
        rows.append({"platform": platform, "date": d.isoformat(), "text": t, "title": "",
                     "score": s10 * scale / 10, "score_scale": scale, "traveler_type": random.choice(types),
                     "room_type": random.choice(ROOMS), "reviewer_country": "", "text_disliked": None})
    return rows


def dump(name, platform, lang, rows):
    (out / name).write_text(json.dumps({"platform": platform, "hotel": "Harbour View Hotel (sample)", "url": "",
                                        "language": {"code": lang}, "rows": rows}, ensure_ascii=False, indent=0),
                            encoding="utf-8")


dump("reviews_tripcom_zh.json", "Trip.com", "zh", make(90, "Trip.com", "zh", ZH_GOOD, ZH_BAD, TYPES_ZH, 10))
dump("reviews_tripcom_en.json", "Trip.com", "en", make(30, "Trip.com", "en", EN_GOOD, EN_BAD, TYPES_EN, 10))
dump("reviews_booking_en.json", "Booking.com", "en", make(40, "Booking.com", "en", EN_GOOD, EN_BAD, TYPES_EN, 10))

(out / "summary.json").write_text(json.dumps({
    "hotel": "Harbour View Hotel (sample)", "not_listed": ["Expedia"], "not_included": ["某音"],
    "platforms": {
        "Trip.com": {"rating": 9.1, "scale": 10, "total_reviews": 120, "all_reviews_scraped": True,
                     "subscores": {"Cleanliness": 9.3, "Amenities": 9.0, "Location": 9.4, "Service": 9.1}},
        "Booking.com": {"rating": 8.9, "scale": 10, "total_reviews": 40, "all_reviews_scraped": True,
                        "subscores": {"Staff": 9.1, "Cleanliness": 9.0, "Value for money": 8.7}}}},
    ensure_ascii=False, indent=1), encoding="utf-8")
print("sample data written to", out)
