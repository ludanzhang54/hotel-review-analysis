"""
Hotel review analysis (one hotel, many platforms)
=====================================

Purpose: merge one hotel's reviews from all platforms into a single summary (the "Summary" sheet); per-platform detail sits in the later sheets.

Method:
- A low-score review is judged by the guest's own score: below 6 on a 10-point scale (Chinese reviews need no translation).
- Platforms are merged weighted by their published review counts; Ctrip and Trip.com are the same set of reviews and are counted once.
- 12 complaint themes, with Chinese and English keywords (traditional characters are converted to simplified first).
- Only one hotel is analysed.
- Complaint theme = share of low-score reviews mentioning the theme, counting only dissatisfied wording (Booking: only the "disliked" part, empty parts are not counted); titles are ignored;
  reviews topped up with the "lowest first" sort are evidence quotes only (visible in the low-score detail) and enter no ratio or count.
- Last 90 days vs 91 days to 1 year ago: from collected reviews, checked per language (per scrape file); only languages that reach back a full year take part,
  otherwise the "earlier" period looks small just because it was not fully collected; a platform needs 10+ reviews in both periods.
- Sub-scores: similar sub-scores from different platforms (Cleanliness etc.) are converted to a 10-point scale and merged with weights.
- Theme impact score = average score of reviews mentioning a theme minus the average of the rest, computed inside each platform first, then weighted by mentions;
  only normally collected reviews. This is correlation, not causation.
- Platform-published ratings, review counts, sub-scores, traveller types and platform topics are read from summary.json.

Input (in one data folder):
  reviews_*.json   result of collect() in the scraping step: {platform, hotel, url, language, rows:[...]} or a plain list of rows
  lowscore_*.json  reviews topped up with the "lowest first" sort (optional)
  summary.json     summary numbers read from the platform pages (format in SKILL.md, optional)

Output: review_analysis_<hotel name>.xlsx

Run: python analyze.py <data folder>     needs pandas 2.1+ and openpyxl (see requirements.txt)
"""

import json
import re
import sys
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------

# Low score: below 6 on a 10-point scale (below 3 on a 5-point scale). In testing this matched the "negative" count on the Ctrip page; Agoda also gives no rating words below 6.
NEGATIVE_BELOW = 6.0
MIN_NEGATIVE_REVIEWS = 5   # with fewer than 5 low-score reviews, theme shares are too unstable to conclude anything
RECENT_DAYS = 90           # "last 90 days"
COMPARE_DAYS = 365         # compared with "91 days to 1 year ago"
MIN_THEME_MENTIONS = 5     # a theme mentioned fewer than 5 times gets no impact score
MIN_PERIOD_REVIEWS = 10    # in the time comparison, a platform with fewer than 10 reviews in a period does not take part
SMALL_SHARE = 0.10         # sub-scores coming from under 10% of published reviews are flagged "small sample"
FEW_REVIEWS = 30           # in the time comparison, a period under 30 reviews warns that single reviews weigh a lot
MIN_IMPACT = 0.3           # "themes that lower the score most" lists only gaps of 0.3 points or more

# 12 complaint themes: English keywords + Chinese keywords (traditional characters are converted to simplified before matching)
THEMES = {
    "Dated facilities/needs renovation": {
        "en": ["old", "dated", "outdated", "tired", "worn", "renovat", "refurb", "run down", "run-down", "shabby"],
        "zh": ["老旧", "陈旧", "破旧", "残旧", "老化", "翻新", "年代感", "有年头", "古老", "又旧又破", "设施旧", "装修旧",
               "比较旧", "很旧", "有点旧", "太旧", "过时", "设备旧", "感觉旧", "旧了"],
    },
    "Hygiene/smell": {
        "en": ["dirty", "unclean", "stain", "mould", "mold", "smell", "musty", "mildew", "dust", "cockroach", "bug", "hair"],
        "zh": ["脏", "不干净", "不清洁", "不卫生", "卫生差", "卫生不好", "卫生一般", "清洁度", "污渍", "污垢", "发霉", "霉味",
               "霉斑", "霉菌", "异味", "怪味", "臭味", "很臭", "发臭", "恶臭", "有味道", "味道大", "味道重", "好大味", "烟味",
               "潮味", "腥味", "潮湿", "灰尘", "头发", "毛发", "蟑螂", "虫", "血迹"],
    },
    "Small room": {
        "en": ["small", "tiny", "cramped", "compact", "no space"],
        "zh": ["房间小", "很小", "太小", "有点小", "比较小", "偏小", "不大", "面积小", "空间小", "狭小", "局促", "房间拥挤", "好细"],
    },
    "Staff/service": {
        "en": ["rude", "unhelpful", "unfriendly", "impolite", "unprofessional", "attitude", "ignored", "poor service",
               "bad service", "slow service", "terrible service", "horrible service", "service was poor", "service was bad",
               "service was slow", "staff were rude", "staff was rude", "staff were unhelpful", "staff was unhelpful"],
        "zh": ["服务差", "服务不好", "服务很差", "服务一般", "服务态度", "态度差", "态度不好", "态度很差", "态度恶劣",
               "态度一般", "不专业", "不耐烦", "没人理", "不理", "冷漠", "敷衍", "前台态度", "员工态度", "工作人员态度",
               "服务员态度", "前台不", "前台很慢", "前台太慢"],
        # zh_ctx: words that are neutral by themselves; they count as a complaint theme only inside a complaining clause (not used for the impact score)
        "zh_ctx": ["服务", "前台", "员工", "工作人员", "服务员", "态度", "管理", "保安"],
    },
    "Check-in/waiting": {
        "en": ["check-in", "check in", "checkin", "queue", "waited", "waiting", "long line"],
        "zh": ["办理入住", "入住慢", "排队", "等了", "等候", "等很久", "等半天"],
    },
    "Noise": {
        "en": ["noise", "noisy", "loud", "thin wall", "hum", "humming"],
        "zh": ["噪音", "噪声", "很吵", "太吵", "有点吵", "好吵", "吵闹", "吵杂", "吵", "嘈", "隔音", "声音大", "很响",
               "太响", "一样响", "装修声"],
    },
    "Repairs (A/C, hot water, bathroom)": {
        "en": ["air con", "aircon", "air-con", "air conditioning", "broken", "leak", "not working", "toilet", "shower", "water", "flush"],
        "zh": ["空调", "冷气", "热水", "水温", "水压", "忽冷忽热", "漏水", "滴水", "水龙头", "堵", "坏", "不能用", "用不了",
               "故障", "马桶", "淋浴", "花洒", "下水", "排风", "忽冷忽烫", "忽大忽小", "不恒温", "排水慢", "下水慢",
               "反味", "不制冷", "热醒"],
        "zh_ctx": ["投屏", "电视", "排水", "地漏", "冰箱", "洗衣机", "温控", "水压"],
    },
    "Breakfast/F&B": {
        "en": ["breakfast", "buffet", "restaurant", "food", "coffee"],
        "zh": ["早餐", "早饭", "自助餐", "餐厅", "咖啡"],
    },
    "Facilities (pool, gym, lift)": {
        "en": ["pool", "gym", "lift", "elevator", "facilit", "lobby"],
        "zh": ["泳池", "游泳池", "健身", "电梯", "𨋢"],
    },
    "Booking info vs reality": {
        "en": ["not as described", "different from the photo", "different from photos", "extra charge", "hidden fee",
               "not included", "misleading"],
        "zh": ["与描述不符", "和描述不符", "跟描述不一样", "与图片不符", "和图片不一样", "照片不一样", "图片不一样", "实物不符",
               "货不对板", "只含一人", "只含一个人", "早餐不含", "不含早", "另外收费", "额外收费", "加收", "另收", "没有提前告知",
               "没提前说", "没有提醒", "没提醒"],
    },
    "Price/deposit/value": {
        "en": ["price", "expensive", "overpriced", "value", "worth", "deposit"],
        "zh": ["价格", "太贵", "很贵", "偏贵", "有点贵", "比较贵", "价格贵", "不值", "性价比", "押金", "按金", "价钱"],
    },
    "Location/transport": {
        "en": ["far from", "no shuttle", "long walk", "hard to find", "no parking", "expensive parking", "parking was expensive",
               "inconvenient location", "location was inconvenient"],
        "zh": ["位置偏", "太远", "很远", "比较远", "偏僻", "交通不便", "交通不方便", "出行不方便", "离地铁远", "地铁远",
               "打车难", "不好打车", "打不到车", "停车难", "不好停车", "没有停车", "停车不便", "停车费", "不好找", "难找"],
        "zh_ctx": ["位置", "交通", "地铁", "公交", "打车", "停车", "出行"],
    },
}

# Traditional -> simplified for the characters used in the keywords (only these; no extra software needed)
TRAD = str.maketrans({
    "舊": "旧", "陳": "陈", "殘": "残", "髒": "脏", "乾": "干", "淨": "净", "潔": "洁", "衛": "卫", "衞": "卫",
    "漬": "渍", "汙": "污", "黴": "霉", "異": "异", "煙": "烟", "塵": "尘", "頭": "头", "髮": "发", "蟲": "虫",
    "跡": "迹", "濕": "湿", "溼": "湿", "間": "间", "侷": "局", "擁": "拥", "擠": "挤", "狹": "狭", "細": "细",
    "較": "较", "點": "点", "覺": "觉", "設": "设", "備": "备", "過": "过", "時": "时", "裝": "装", "務": "务",
    "態": "态", "臺": "台", "檯": "台", "煩": "烦", "沒": "没", "禮": "礼", "賓": "宾", "員": "员", "辦": "办",
    "隊": "队", "聲": "声", "響": "响", "調": "调", "氣": "气", "熱": "热", "溫": "温", "壓": "压", "壞": "坏",
    "馬": "马", "灑": "洒", "風": "风", "龍": "龙", "飯": "饭", "廳": "厅", "電": "电", "價": "价", "貴": "贵",
    "無": "无", "問": "问", "題": "题", "譜": "谱", "爛": "烂", "亂": "乱", "訴": "诉", "絕": "绝", "噁": "恶",
    "遺": "遗", "鬧": "闹", "會": "会", "兒": "儿", "燙": "烫", "恆": "恒", "視": "视", "製": "制", "錯": "错", "說": "说", "説": "说", "話": "话", "錢": "钱", "難": "难", "費": "费", "專": "专", "離": "离", "遠": "远", "鐵": "铁", "車": "车", "惡": "恶", "積": "积", "廁": "厕", "續": "续",
})

# Traveller types are merged into 5 groups; the platform's original wording stays in the detail sheets
TRAVELLER_GROUPS = [
    ("Business", ["business", "商务", "商務", "出差"]),
    ("Couple", ["couple", "情侣", "情侶"]),
    ("Solo", ["solo", "独自", "獨自"]),
    ("Family", ["famil", "家庭", "亲子", "親子"]),
    ("Friends/group", ["friend", "group", "好友", "朋友", "团体", "團體"]),
]

# English keywords match whole words (old does not hit told, hair does not hit chair) and allow common endings (stains, smelly, leaking);
# the ones below are stems that may be followed by any letters (renovated, refurbished, facilities)
EN_STEMS = {"renovat", "refurb", "facilit"}
EN_PATTERNS = {
    theme: re.compile("|".join(
        rf"\b{re.escape(w)}" + ("" if w in EN_STEMS else r"(?:s|es|ed|ing|y|er)?\b") for w in words["en"]))
    for theme, words in THEMES.items()
}

# Sub-scores of different platforms are put into the same group before merging (matched in order; a name containing these words goes to that group); ones that fit no group (e.g. Ctrip "environment") are not merged
SUBSCORE_GROUPS = [
    ("Cleanliness", ["卫生", "cleanliness", "清洁"]),
    ("Service", ["服务", "staff", "service"]),
    ("Facilities", ["设施", "facilities", "amenities", "property condition"]),
    ("Comfort", ["comfort", "舒适"]),
    ("Location", ["location", "位置"]),
    ("Value", ["value", "性价比"]),
]

# English keywords that are neutral by themselves: they count as a complaint theme only when a dissatisfied wording is in the same clause
EN_NEUTRAL = {"breakfast", "buffet", "restaurant", "food", "coffee", "pool", "gym", "lift", "lifts", "elevator", "elevators",
              "lobby", "price", "prices", "value", "worth", "deposit", "check-in", "check in", "checkin", "toilet", "toilets",
              "shower", "showers", "water", "air con", "aircon", "air-con", "air conditioning", "attitude", "queue",
              "waiting", "waited", "compact", "flush", "shuttle", "facilit", "renovat", "refurb"}

BAD_FILENAME = re.compile(r'[\\/:*?"<>|]')


def traveller_group(label):
    text = str(label or "").lower()
    if not text.strip() or text == "nan":
        return None
    for group, words in TRAVELLER_GROUPS:
        if any(w in text for w in words):
            return group
    return "Other"


def is_negative(score10):
    return None if pd.isna(score10) else bool(score10 < NEGATIVE_BELOW)


# A complaint theme counts only the clause that is "complaining": a low-score review often also has praise ("rich breakfast", "friendly service"),
# so the review is cut into short clauses by punctuation, and a theme word counts only if a dissatisfied wording is in its clause.
# Keywords that are dissatisfied by themselves (dirty, very noisy, too expensive, mouldy) count directly.
NEG_CUES_ZH = ["不", "没", "无", "差", "慢", "坏", "脏", "吵", "嘈", "闹", "贵", "问题", "失望", "离谱", "糟", "难",
               "霉", "臭", "味", "烂", "旧", "破", "漏", "堵", "乱", "久", "投诉", "拒绝", "恶心", "腥", "灰", "虫",
               "烦", "敷衍", "冷漠", "一般", "遗憾", "可惜", "缺", "少", "押金", "按金", "窄", "挤", "响",
               # Praise reviews often voice a complaint in a suggesting tone: "breakfast could be richer", "temperature control has room to improve"
               "进步空间", "有待", "就好了", "更好了", "再丰富", "细致些", "希望", "建议增加", "建议改", "建议加",
               "有些前客人",
               # Booking info vs reality: "only told at check-in that breakfast is for one person", "different from the photos"
               "才说", "才告知", "才知道", "不符", "不一样", "只含", "加收", "另收", "另外收", "额外收"]
# First remove phrases that contain the characters above but are actually praise
POSITIVE_ZH = ["不错", "不愧", "无可挑剔", "没得说", "没话说", "没的说", "无微不至", "毫不犹豫", "不虚此行", "没问题",
               "没有问题", "不用担心", "不少", "少不了", "必不可少", "一尘不染", "不远", "没有异味", "无异味", "不吵",
               "不贵", "不脏", "不旧", "不慢", "不嘈", "无可", "不辞", "不厌", "没一会", "没多久", "无论", "不得不",
               "差不多", "诸多问题", "各种问题", "有问必答", "解决问题", "问题都能", "不管", "响应", "不用", "比一般",
               "美味", "口味", "趣味", "没有之一", "堵车", "不过", "来不及", "不止", "不差"]
NEG_CUES_EN = re.compile(r"\b(not|no|never|poor|bad|dirty|broken|rude|slow|noisy|small|old|smell\w*|expensive|"
                         r"but|however|wasn'?t|didn'?t|couldn'?t|isn'?t|doesn'?t|problem\w*|issue\w*|disappoint\w*|"
                         r"terrible|horrible|awful|worst|lack\w*|missing|unclean|stain\w*|leak\w*|cold|overpriced)\b")
# "没有…毛发/异味" (no hair / no smell) means "no problem"; remove before judging complaints
ABSENCE_ZH = re.compile(r"(没有|无|没)[^，。！？,.!?]{0,6}?(毛发|头发|异味|怪味|味道|灰尘|虫|蟑螂|噪音|霉|吵)")
# "虽然没有××那样大" is a comparison, not a complaint
COMPARE_ZH = re.compile(r"虽然?没有?[^，。！？,.!?]{0,20}?(那样|那么|这么)|不用担心[^，。！？,.!?]*|没有那种[^，。！？,.!?]*")
# Phrases inside a theme that are not complaints: "a small inconvenience" is not a small room, "the lobby is not big" is about the lobby, "child's clothes got dirty" is not a hygiene problem
THEME_NOT = {
    "Small room": re.compile(r"小困扰|小问题|小毛病|水压[^，。！？,.!?]{0,4}小|大堂[^，。！？,.!?]{0,3}(不大|小)"),
    "Hygiene/smell": re.compile(r"衣服[^，。！？,.!?]{0,4}脏"),
    # Expensive shops nearby are not the hotel's price problem: "the convenience store downstairs is expensive"
    "Price/deposit/value": re.compile(r"(便利店|超市|商场|商店|外卖|餐馆|饭店|市场|景区|门票|出租车|打车|停车场)[^，。！？,.!?]{0,8}(贵|不便宜)"),
}
# The repairs theme counts only "broken / not working"; "no bidet" means the facility is not provided, not a fault
LACK_ZH = re.compile(r"(没有|没)[^，。！？,.!?]{0,3}(卫洗丽|智能马桶|智能盖板|加热便座|加湿器|浴缸)")
CLAUSE = re.compile(r"[。！？；，、：:,.!?;\n|]+|\s+but\s+|\s+however\s+")


CJK = re.compile(r"[\u4e00-\u9fff]")
CUE_WINDOW = 6  # a dissatisfied wording must be within 6 characters of the theme word


def near_cue(clause, word):
    """The theme word appears in this clause and a dissatisfied wording is within 6 characters of it."""
    start = clause.find(word)
    while start >= 0:
        around = clause[max(0, start - CUE_WINDOW): start + len(word) + CUE_WINDOW]
        if any(n in around for n in NEG_CUES_ZH):
            return True
        start = clause.find(word, start + 1)
    return False


def is_complaint_zh(word):
    return any(c in word for c in NEG_CUES_ZH + ["小", "偏", "远"])


try:  # with opencc, convert traditional to simplified as a whole (Chinese reviews on the Trip.com English page are traditional); without it only the character table above is used
    from opencc import OpenCC
    _T2S = OpenCC("t2s").convert
except Exception:  # noqa: BLE001
    _T2S = None


def normalize(text):
    t = str(text or "").lower()
    if _T2S:
        t = _T2S(t)
    return t.translate(TRAD).replace("坏境", "环境")


# What guests praise: a praising wording within 10 characters of the aspect word, and no dissatisfied wording in the same clause.
# Words that are praise by themselves (clean, quiet, good value) and apartment-specific facilities (washer, kitchen) count whenever they appear.
# Concrete things guests mention in reviews (facilities, location, trip purpose): used to see what each guest group cares about and what the listing should say.
# Counts only "mentions" (mentions inside low-score reviews are excluded; the denominator is all reviews) and does not judge praise or complaint; split into "Basics" (every hotel should do) and "Selling points" (not every hotel has).
# In summary.json you can add nearby place names with "feature_keywords": {"Nearby mall": ["Wanda", "shopping centre"]}; they merge into an item of the same name or add a new one.
FEATURES = {
    # Selling points
    "Kitchen, cookware": ("Selling point", r"厨房|厨具|灶|锅碗|餐具|做饭|煮饭|kitchen|cook|stove"),
    "Washer/dryer": ("Selling point", r"洗衣|烘干|洗烘|washing|laundry|dryer|washer"),
    "Fridge, microwave": ("Selling point", r"冰箱|微波炉|fridge|refrigerator|microwave"),
    "Large room, suite": ("Selling point", r"房间.{0,3}大|宽敞|套房|两室|两居|一室一厅|客厅|spacious|suite|living room"),
    "Breakfast": ("Selling point", r"早餐|早饭|breakfast"),
    "Gym, sauna": ("Selling point", r"健身|桑拿|gym|fitness|sauna"),
    "Lounge, billiards, games": ("Selling point", r"台球|桌上足球|游戏|休闲廊|乒乓|billiard|pool table|game room"),
    "Free coffee": ("Selling point", r"咖啡|coffee"),
    "Kids, family": ("Selling point", r"孩子|小孩|小朋友|儿童|宝宝|闺女|儿子|女儿|挖沙|\bkids?\b|child|toddler|baby"),
    "Welcome fruit": ("Selling point", r"水果|fruit"),
    "Sea, beach": ("Selling point", r"海边|沙滩|海滩|栈桥|看海|海景|beach|seaside|sea view"),
    "Mall, shopping": ("Selling point", r"商场|超市|购物|\bmall\b|supermarket|shopping"),
    "Transport, airport, station": ("Selling point", r"机场|高铁|火车站|地铁|公交|打车|出租车|滴滴|airport|station|metro|\bcab\b|taxi|didi"),
    "Foreign-language staff": ("Selling point", r"会说韩语|会韩语|韩语|会说英语|会英语|会日语|english.speaking|speaks? english"),
    "Parking": ("Selling point", r"停车|parking"),
    "Quiet, soundproofing": ("Selling point", r"安静|隔音|\bquiet"),
    "Bed, pillows": ("Selling point", r"床垫|床品|枕头|床很舒服|床也很舒服|mattress|pillow|comfortable bed"),
    "Business trip, company": ("Selling point", r"出差|公司|工厂|开会|business trip|for work|work trip"),
    "Value for money": ("Selling point", r"性价比|实惠|划算|value for money|good value"),
    # Basics
    "Service, front desk": ("Basics", r"服务|前台|员工|工作人员|管家|staff|service|reception"),
    "Clean, hygiene": ("Basics", r"干净|卫生|整洁|一尘不染|clean"),
}


def feature_patterns(summary):
    pats = {k: v for k, v in FEATURES.items()}
    for name, words in (summary.get("feature_keywords") or {}).items():
        extra = "|".join(re.escape(w) for w in words)
        if name in pats:
            kind, pat = pats[name]
            pats[name] = (kind, pat + "|" + extra)
        else:
            pats[name] = ("Selling point", extra)
    return pats


LANG_NAMES = {"zh": "Chinese", "ko": "Korean", "ja": "Japanese", "en": "English"}


def language_mentions(reviews, summary):
    """What foreign-language reviews mention (by review language; language is not nationality): how many reviews in each foreign language mention each thing, compared with Chinese reviews.
    There are usually few foreign-language reviews, so counts are given, not only shares. Korean and Japanese reviews must first be translated into Chinese in translations.json to be counted."""
    # The denominator is all reviews in this language (same as other shares in the report); mentions inside low-score reviews are excluded (mostly complaints)
    ok = reviews.copy()
    ok["_t"] = ok["analysis text"].fillna("").astype(str).map(normalize).str.lower() + " " + ok["title"].fillna("").astype(str).str.lower()
    ok.loc[ok["Low score"] == True, "_t"] = ""  # noqa: E712
    langs = [l for l, c in ok["review_language"].value_counts().items() if l != "unknown"]
    rows = []
    for name, (kind, pat) in feature_patterns(summary).items():
        hit = ok["_t"].str.contains(pat, regex=True)
        row = {"Category": kind, "Mentioned": name}
        for l in langs:
            sub = ok["review_language"] == l
            row[f"{LANG_NAMES.get(l, l)} ({int(sub.sum())} reviews) count"] = int(hit[sub].sum())
            row[f"{LANG_NAMES.get(l, l)} ({int(sub.sum())} reviews) share"] = hit[sub].mean()
        rows.append(row)
    table = pd.DataFrame(rows)
    table.attrs["langs"] = [(LANG_NAMES.get(l, l), int((ok["review_language"] == l).sum())) for l in langs]
    return table


def room_type_table(reviews):
    """Room types: number of reviews, average score, share with complaints, most complained-about problem. Only room types with 10+ reviews are listed."""
    rows = []
    d = reviews[reviews["room_type"].fillna("").astype(str).str.strip() != ""]
    for rt, part in d.groupby("room_type"):
        if len(part) < MIN_GROUP_REVIEWS:
            continue
        comp = part["Complaint themes"].fillna("") != ""
        top = part.loc[comp, "Complaint themes"].str.split("; ").explode().value_counts()
        rows.append({"Room type": rt, "Reviews": len(part), "Average score (10-point)": round(part["score10"].mean(), 1),
                     "Share with complaints": comp.mean(), "Most complained about": ", ".join(f"{k} {v} reviews" for k, v in top.head(3).items()),
                     "Most common guests in this room type": ", ".join(f"{g} {n} reviews" for g, n in part["Guest group"].replace("", "not stated").fillna("not stated").value_counts().head(2).items())})
    return pd.DataFrame(rows).sort_values("Reviews", ascending=False) if rows else pd.DataFrame(columns=["Room type"])


def language_summary(reviews):
    """Review languages: number of reviews, average score, which platforms, traveller types."""
    rows = []
    for l, d in reviews.groupby("review_language"):
        rows.append({"Language": LANG_NAMES.get(l, l), "Reviews": len(d), "Average score (10-point)": round(d["score10"].mean(), 1),
                     "Platform": ", ".join(f"{p} {n} reviews" for p, n in d["platform"].value_counts().items()),
                     "Traveller type": ", ".join(f"{g} {n} reviews" for g, n in d["Guest group"].replace("", "not stated").fillna("not stated").value_counts().items())})
    return pd.DataFrame(rows).sort_values("Reviews", ascending=False)


def segment_mentions(reviews, summary):
    """What each guest group mentions in reviews: how many reviews of each group mention each thing. The denominator is all reviews of that group (same as in "What guests praise"),
    mentions inside low-score reviews are excluded (mostly complaints). Only groups with 10+ reviews are listed."""
    ok = reviews.copy()
    ok["_t"] = ok["analysis text"].fillna("").astype(str).map(normalize).str.lower() + " " + ok["title"].fillna("").astype(str).str.lower()
    ok.loc[ok["Low score"] == True, "_t"] = ""  # noqa: E712
    typed = ok[ok["Guest group"].fillna("") != ""]
    groups = [g for g, c in typed["Guest group"].value_counts().items() if c >= MIN_GROUP_REVIEWS and g != "Other"]
    rows = []
    for name, (kind, pat) in feature_patterns(summary).items():
        hit = ok["_t"].str.contains(pat, regex=True)
        row = {"Category": kind, "Mentioned": name, "Mentions (all reviews)": int(hit.sum()), "Share (all reviews)": hit.mean() if len(ok) else None}
        th = typed["_t"].str.contains(pat, regex=True)
        for g in groups:
            sub = typed["Guest group"] == g
            row[f"{g} ({int(sub.sum())} reviews) count"] = int(th[sub].sum())
            row[f"{g} ({int(sub.sum())} reviews) share"] = th[sub].mean()
        rows.append(row)
    table = pd.DataFrame(rows).sort_values(["Category", "Mentions (all reviews)"], ascending=[True, False])
    table.attrs["groups"] = [(g, int((typed["Guest group"] == g).sum())) for g in groups]
    table.attrs["n"] = len(ok)
    return table


PRAISE = {
    "Staff/service": {"zh": ["服务", "前台", "员工", "工作人员", "服务员", "管家", "保安", "经理", "接待"],
               "en": ["staff", "service", "reception", "team"]},
    "Cleanliness": {"zh": ["干净", "整洁", "一尘不染", "卫生"], "en": ["clean"]},
    "Breakfast/F&B": {"zh": ["早餐", "餐厅", "早饭"], "en": ["breakfast"]},
    "Location/transport": {"zh": ["位置", "交通", "海边", "沙滩", "地段", "周边", "商场", "出行", "地理"], "en": ["location", "beach", "mall"]},
    "Room (new, comfortable, well equipped)": {"zh": ["房间", "房型", "床", "空间", "装修", "设施"], "en": ["room", "bed", "spacious"]},
    "In-room facilities (kitchen, washer)": {"zh": ["洗衣机", "厨房", "冰箱", "微波炉", "烘干", "做饭", "厨具", "电磁炉"],
                         "en": ["kitchen", "laundry", "washing", "washer", "dryer"]},
    "Gym, lounge and free coffee": {"zh": ["健身房", "桌球", "高尔夫", "休闲", "咖啡", "泳池"], "en": ["gym", "lounge", "billiard", "coffee", "pool"]},
    "Quiet": {"zh": ["安静", "隔音"], "en": ["quiet"]},
    "Value for money": {"zh": ["性价比", "价格", "实惠", "划算", "物超所值"], "en": ["value", "price"]},
}
PRAISE_SELF = {"干净", "整洁", "一尘不染", "安静", "实惠", "划算", "物超所值", "洗衣机", "厨房", "做饭", "烘干",
               "健身房", "桌球", "高尔夫", "clean", "quiet", "kitchen", "laundry", "gym"}
POS_ZH = ["好", "棒", "赞", "不错", "满意", "热情", "热心", "周到", "贴心", "耐心", "专业", "干净", "整洁", "丰富", "方便",
          "便利", "舒适", "舒服", "新", "大", "宽敞", "安静", "齐全", "推荐", "喜欢", "温馨", "惊喜", "完美", "优秀", "高",
          "近", "nice", "实惠", "划算", "值", "亲切", "友好", "礼貌", "到位", "给力", "好吃", "美味", "丰盛", "多样",
          "足够", "还可以", "还行", "过得去", "媲美", "可口", "精致", "没话说", "没得说", "无可挑剔", "一流", "超出预期", "靠谱",
          "高效", "温柔", "细致", "用心"]
# The "大" in the praise word list is falsely triggered by "大堂/大厅" (lobby); mask these words before judging. When praising rooms, "clean" counts as hygiene, not as room
NOT_PRAISE = ["大堂", "大厅", "大型", "大家", "大概", "大约", "大叔", "大姐", "大哥", "大门", "大海", "大床"]
THEME_SKIP_POS = {"Room (new, comfortable, well equipped)": {"干净", "整洁"}}
# Korean/Japanese reviews cannot be read for praise yet; they are not counted in the praise denominator
UNREADABLE = re.compile(r"[\uac00-\ud7af\u3040-\u30ff]")
PRAISE_WINDOW = 10  # a praising wording must be within 10 characters of the aspect word
POS_EN = re.compile(r"\b(good|great|nice|excellent|friendly|helpful|clean|comfortable|convenient|quiet|spacious|delicious|"
                    r"recommend\w*|love\w*|perfect|amazing|kind|attentive|well|wonderful|best)\b")


PRAISE_CLAUSE = re.compile(r"[。！？；，、,.!?;\n|]+|\s+but\s+|\s+however\s+")


def clauses_of(text, splitter=None):
    """Cut into short clauses by punctuation; Chinese is also cut by spaces. Praise is not cut by colons (\"service: friendly\" must be read together)."""
    out = []
    for c in (splitter or CLAUSE).split(ABSENCE_ZH.sub("", text)):
        out += c.split() if CJK.search(c) else [c]
    return [c for c in out if c.strip()]


def praise_hits(text):
    """Which aspects a review praises (can be several)."""
    hits = []
    text = normalize(text)
    if UNREADABLE.search(text):
        return []
    clauses = clauses_of(text, PRAISE_CLAUSE)
    for theme, words in PRAISE.items():
        skip = THEME_SKIP_POS.get(theme, set())
        pos = [p for p in POS_ZH if p not in skip]
        for c in clauses:
            if CJK.search(c):  # English staff names inside a Chinese sentence would take up the window; remove them first
                c = re.sub(r"[a-z][a-z&/ ]*", "", c)
            for w in NOT_PRAISE:
                c = c.replace(w, "＿" * len(w))
            masked = COMPARE_ZH.sub("", c)
            if theme.startswith("Repairs"):
                masked = LACK_ZH.sub("", masked)
            if theme in THEME_NOT:
                masked = THEME_NOT[theme].sub("", masked)
            for w in POSITIVE_ZH:
                masked = masked.replace(w, "＿" * len(w))
            if any(n in masked for n in NEG_CUES_ZH) or NEG_CUES_EN.search(c):
                continue
            ok = False
            for kw in words["zh"]:
                i = c.find(kw)
                while i >= 0 and not ok:
                    around = c[max(0, i - PRAISE_WINDOW): i + len(kw) + PRAISE_WINDOW]
                    ok = kw in PRAISE_SELF or any(p in around for p in pos)
                    i = c.find(kw, i + 1)
                if ok:
                    break
            if not ok:
                ok = any(re.search(r"\b" + kw, c) and (kw in PRAISE_SELF or POS_EN.search(c)) for kw in words["en"])
            if ok:
                hits.append(theme)
                break
    return hits


def theme_hits(text, complaints_only=False):
    """Which themes a review mentions (can be several; each theme once).
    With complaints_only=True only complaining clauses count (used for the complaint themes)."""
    lower = str(text or "").lower()
    simple = normalize(text)
    if not complaints_only:
        return [theme for theme, words in THEMES.items()
                if EN_PATTERNS[theme].search(lower) or any(w in simple for w in words["zh"])]
    hits = []
    simple = simple.replace("坏境", "环境")  # common typo; "坏" would be read as a repair problem
    clauses = []
    for c in CLAUSE.split(ABSENCE_ZH.sub("", simple)):
        # Chinese reviews often use spaces to separate clauses; English must not be cut by spaces
        clauses += c.split() if CJK.search(c) else [c]
    clauses = [c for c in clauses if c.strip()]
    for theme, words in THEMES.items():
        for c in clauses:
            masked = COMPARE_ZH.sub("", c)
            if theme.startswith("Repairs"):
                masked = LACK_ZH.sub("", masked)
            if theme in THEME_NOT:
                masked = THEME_NOT[theme].sub("", masked)
            for w in POSITIVE_ZH:
                masked = masked.replace(w, "＿" * len(w))
            en = EN_PATTERNS[theme].search(c)
            if any(is_complaint_zh(w) and w in masked for w in words["zh"]) \
                    or any(near_cue(masked, w) for w in words["zh"] + words.get("zh_ctx", [])) \
                    or (en and NEG_CUES_EN.search(c)) \
                    or (en and not any(en.group(0).startswith(n) for n in EN_NEUTRAL)):
                hits.append(theme)
                break
    return hits


def to_number(value):
    if value is None or isinstance(value, (int, float)):
        return value
    match = re.search(r"-?[\d,]+(?:\.\d+)?", str(value))
    return float(match.group().replace(",", "")) if match else None


# ---------------------------------------------------------------------------
# Reading data
# ---------------------------------------------------------------------------

def read_rows(path):
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, dict):
        rows = data.get("rows", [])
        language = data.get("language")
        code = language.get("code") if isinstance(language, dict) else language
        for r in rows:
            r.setdefault("platform", data.get("platform"))
            r.setdefault("hotel", data.get("hotel"))
            r.setdefault("review_language", code)
    else:
        rows = data
    for r in rows:
        r["source file"] = path.stem  # one file = one language of one platform; the time comparison checks coverage per file
    return rows


def column(df, name, default=None):
    return df[name] if name in df.columns else pd.Series([default] * len(df), index=df.index)


def apply_translations(df, folder):
    """translations.json in the data folder: for languages the keywords cannot read (Korean, Japanese, ...), a person translates the original into Chinese:
    [{"platform": "Trip.com", "date": "2024-03-07", "text_starts": "호텔도 좋고", "zh": "酒店好，员工也亲切……"}]
    The translation is used to detect complaints, praise and mentions; the report and Excel keep both the original and the translation."""
    df["Translation"] = ""
    path = Path(folder) / "translations.json"
    if path.exists() and not df.empty:
        for c in json.loads(path.read_text(encoding="utf-8")):
            hit = (df["platform"] == c["platform"]) & (df["date"].dt.strftime("%Y-%m-%d") == c["date"]) \
                & df["text"].str.startswith(c["text_starts"])
            if not hit.any():
                print(f"translations.json: this entry matches no review: {c['platform']} {c['date']} {c['text_starts']}")
            df.loc[hit, "Translation"] = c["zh"]
    df["analysis text"] = [z if z else t for z, t in zip(df["Translation"], df["text"])]
    return df


def apply_corrections(df, folder):
    """Manual corrections after checking (corrections.json in the data folder); complaint themes the automatic detection got wrong are changed here:
    [{"platform": "Fliggy", "date": "2025-05-06", "text_starts": "差评给餐厅男经理", "themes": ["Staff/service"], "why": "complains about restaurant staff"}]
    themes = the correct complaint themes of this review (empty list = not a complaint)."""
    path = Path(folder) / "corrections.json"
    if not path.exists() or df.empty:
        return df
    for c in json.loads(path.read_text(encoding="utf-8")):
        hit = (df["platform"] == c["platform"]) & (df["date"].dt.strftime("%Y-%m-%d") == c["date"]) \
            & df["text"].map(normalize).str.startswith(normalize(c["text_starts"]))
        if not hit.any():
            print(f"corrections.json: this entry matches no review: {c}")
        df.loc[hit, "Complaint themes"] = "; ".join(c.get("themes", []))
    return df


# Older data files may use Chinese platform names; they are renamed on loading
LEGACY_PLATFORMS = {"携程": "Ctrip", "美团": "Meituan", "飞猪": "Fliggy"}


def load_reviews(folder, pattern):
    rows = []
    for path in sorted(folder.glob(pattern)):
        rows += read_rows(path)
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    for name in ["platform", "hotel", "title", "text", "text_disliked", "date", "traveler_type", "room_type",
                 "reviewer_country", "review_language", "source file"]:
        if name not in df.columns:
            df[name] = None
    df["platform"] = df["platform"].replace(LEGACY_PLATFORMS)
    df["text"] = df["text"].fillna("").astype(str).str.split().str.join(" ")
    # zh-CN, zh-HK, zh_TW, en-US etc. are all merged into zh, en; traditional Chinese is also Chinese
    df["review_language"] = (df["review_language"].fillna("unknown").astype(str).str.strip().str.lower()
                             .str.split(r"[-_]").str[0].replace({"": "unknown", "cn": "zh"}))
    scale = pd.to_numeric(column(df, "score_scale", 10), errors="coerce").fillna(10)
    df["score10"] = pd.to_numeric(column(df, "score"), errors="coerce") * 10 / scale
    df["date"] = pd.to_datetime(df["date"], errors="coerce", format="mixed")
    df = df[df["score10"].notna() | (df["text"] != "")]
    df["Low score"] = df["score10"].map(is_negative)
    df["Guest group"] = df["traveler_type"].map(traveller_group)
    # Complaint themes look only at the "dissatisfied" part: if Booking has a separate "disliked" part, only that part is used (praise in the liked part is not a complaint theme)
    # If the field exists but is empty (the guest wrote only the liked part), the review has no complaint theme, so praise is not counted in
    has_split = df["text_disliked"].notna()
    disliked = df["text_disliked"].fillna("").astype(str).str.strip()
    # Reviews the keywords cannot read (Korean, Japanese, ...) are read through the Chinese translation in translations.json (the original is kept)
    df = apply_translations(df, folder)
    df["Complaint themes"] = [";".join(theme_hits(d if split else t, complaints_only=True))
                   for split, d, t in zip(has_split, disliked, df["analysis text"])]
    df["Complaint themes"] = df["Complaint themes"].str.replace(";", "; ")
    df = apply_corrections(df, folder)
    # What guests praise: only reviews that are not low-score
    df["Praise aspects"] = ["; ".join(praise_hits(t)) if neg is not True else "" for t, neg in zip(df["analysis text"], df["Low score"])]
    # De-duplicate: only the same platform, text, date, score, country and room type count as duplicates, so identical short reviews by different guests are not deleted
    keys = ["platform", "text", "date", "score10", "reviewer_country", "room_type"]
    df = df.drop_duplicates(subset=keys)
    return df.reset_index(drop=True)


def load_summary(folder):
    path = folder / "summary.json"
    if not path.exists():
        return {}
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        return {}
    try:
        summary = json.loads(text) or {}
    except json.JSONDecodeError as e:
        sys.exit(f"summary.json has the wrong format ({e}); fix it following step 4 of SKILL.md and run again.")
    if isinstance(summary, dict) and isinstance(summary.get("platforms"), dict):
        summary["platforms"] = {LEGACY_PLATFORMS.get(k, k): v for k, v in summary["platforms"].items()}
        for f in summary["platforms"].values():
            for t in (f.get("topics") or []) if isinstance(f, dict) else []:
                for old, new in (("话题", "topic"), ("提及次数", "mentions"), ("正面占比", "positive_share"), ("负面占比", "negative_share")):
                    if old in t:
                        t[new] = t.pop(old)
    if not isinstance(summary, dict) or not isinstance(summary.get("platforms") or {}, dict):
        sys.exit("summary.json has the wrong format: the outermost level must be {\"hotel\": ..., \"platforms\": {...}}; fix it following step 4 of SKILL.md.")
    summary["platforms"] = {p: f for p, f in (summary.get("platforms") or {}).items() if isinstance(f, dict)}
    for f in summary["platforms"].values():
        for key in ["rating", "scale", "total_reviews", "hidden_reviews", "negative_reviews"]:
            if key in f:
                v = to_number(f[key])
                f[key] = None if v is None or pd.isna(v) else v
        f["subscores"] = {k: to_number(v) for k, v in (f.get("subscores") or {}).items()}
        f["subscores"] = {k: v for k, v in f["subscores"].items() if v is not None and not pd.isna(v)}
        f["traveller_counts"] = {k: to_number(v) for k, v in (f.get("traveller_counts") or {}).items()}
        for t in f.get("topics") or []:
            if "mentions" in t:
                t["mentions"] = to_number(t["mentions"])
            for key in ["positive_share", "negative_share"]:
                if key in t:
                    v = to_number(t[key])
                    t[key] = v / 100 if v is not None and (str(t[key]).strip().endswith("%") or v > 1) else v
        f.pop("excluded", None)
    return summary


# ---------------------------------------------------------------------------
# Merge helpers
# ---------------------------------------------------------------------------

def wavg(values):
    """Weighted average: values = [(value, weight)], skipping empty values and zero weights."""
    pairs = [(v, w) for v, w in values if v is not None and not pd.isna(v) and w and not pd.isna(w)]
    total = sum(w for _, w in pairs)
    return sum(v * w for v, w in pairs) / total if total else None


def platform_negative(p, reviews, pf):
    """(low-score share, note). If the platform page publishes a negative count (summary negative_reviews, e.g. Ctrip's "negative (7)") it is used,
    otherwise the share among the collected reviews."""
    neg, pub = pf.field(p, "negative_reviews"), pf.published(p)
    if neg is not None and pub:
        return neg / pub, f"{pf.disp(p)} page publishes {int(neg)}/{int(pub)} negative reviews"
    df = reviews[reviews["platform"] == p]
    share = negative_share(df)
    if share is None:
        return None, f"{pf.disp(p)}: collected reviews have no scores"
    n = int(df["Low score"].dropna().astype(bool).sum())
    return share, f"{pf.disp(p)}: {n} low-score reviews among the {int(df['Low score'].notna().sum())} collected"


def negative_share(df):
    flagged = df["Low score"].dropna()
    return flagged.astype(bool).mean() if len(flagged) else None


class Platforms:
    """Rating and weight of each platform. Weight = the platform's published review count.
    - Ctrip and Trip.com are the same set of reviews: reviews are collected from Trip.com, rating, review count and sub-scores use Ctrip's, counted once.
    - A platform that was scraped but has no review count in summary gets the scraped count as weight (stated in "how it is computed").
    - A score-only platform with no review count does not take part in merging (stated)."""

    def __init__(self, summary, reviews):
        self.info = summary.get("platforms", {})
        self.scraped = sorted(reviews["platform"].dropna().unique())
        self.counts = reviews["platform"].value_counts().to_dict()

    def same(self, p):
        """Treat Ctrip and Trip.com as one platform: both use the name Ctrip (Trip.com's rating is used when Ctrip has none)."""
        if p in ("Trip.com", "Ctrip") and "Ctrip" in self.info and "Trip.com" in self.info:
            # The Trip.com page is on a 10-point scale with one more digit of precision; use it when both have a rating
            return "Trip.com" if self.info["Trip.com"].get("rating") is not None else "Ctrip"
        return "Ctrip" if p == "Trip.com" and "Ctrip" in self.info else p

    def field(self, p, key):
        """Look at the representative platform first, then the other one (Ctrip/Trip.com)."""
        for name in [self.same(p), p, "Ctrip" if p == "Trip.com" else "Trip.com" if p == "Ctrip" else p]:
            v = (self.info.get(name) or {}).get(key)
            if v is not None:
                return v
        return None

    def published(self, p):
        return self.field(p, "total_reviews")

    def disp(self, p):
        # Ctrip and Trip.com are the same set of reviews; the report calls it "Ctrip" only
        return "Ctrip" if p in ("Trip.com", "Ctrip") and "Ctrip" in self.info else p

    def weight(self, p):
        """(weight, note); the weight is None when the platform cannot take part in merging."""
        pub = self.published(p)
        if pub:
            return pub, ""
        scraped = self.counts.get(p) or (self.counts.get("Trip.com") if p == "Ctrip" else None)
        if scraped:
            return scraped, f"{p}: published review count not filled in, using the {scraped} collected"
        return None, f"{p}: published review count not filled in, not included"

    def rating10(self, p):
        for name in [self.same(p), p]:
            f = self.info.get(name) or {}
            if f.get("rating") is not None:
                return round(f["rating"] * 10 / (f.get("scale") or 10), 2)
        return None

    def rating_platforms(self):
        """Platforms that take part in the overall rating (Ctrip and Trip.com: only one is kept)."""
        names = [p for p in self.info if self.rating10(p) is not None]
        return list(dict.fromkeys(self.same(p) for p in names))

    def all_published(self):
        """Sum of all platforms with a review count (Ctrip and Trip.com counted once)."""
        return sum(self.published(p) or 0 for p in dict.fromkeys(self.same(p) for p in self.info))

    def coverage(self, used):
        """Share of the total published reviews that these platforms make up; says so when a platform has no review count."""
        total = self.all_published()
        names = list(dict.fromkeys(self.same(p) for p in used))
        part = sum(self.published(p) or 0 for p in names)
        unknown = [p for p in names if not self.published(p)]
        text = f", {part / total:.0%} of the published reviews" if total and part else ""
        return text + (f" (also {', '.join(unknown)}: no review count, not in the share)" if unknown else "")


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------

def theme_shares(df):
    neg = df[df["Low score"] == True]  # noqa: E712
    row = {"Low-score reviews": len(neg)}
    for theme in THEMES:
        row[theme] = neg["Complaint themes"].str.contains(re.escape(theme)).mean() if len(neg) else None
    return row


def top_themes(row, n=3):
    if row["Low-score reviews"] < MIN_NEGATIVE_REVIEWS:
        return []
    shares = pd.Series({t: row[t] for t in THEMES}).astype(float)
    return list(shares[shares > 0].nlargest(n).items())


def theme_reason(row):
    if row.get("_no_score"):
        return "collected reviews have no scores, so low-score reviews cannot be told apart"
    if row["Low-score reviews"] < MIN_NEGATIVE_REVIEWS:
        return f"fewer than {MIN_NEGATIVE_REVIEWS} low-score reviews, no conclusion"
    return "the low-score reviews hit no theme; see the low-score detail"


def theme_table(pool, reviews, pf):
    """One row per platform plus "All platforms". Theme shares of all platforms = each platform's share weighted by "platform reviews x low-score share" (estimated low-score count);
    only normally collected reviews (topped-up low-score reviews are evidence quotes only and enter no ratio); platforms with fewer than 5 low-score reviews do not take part in merging."""
    rows, used, left = [], [], []
    for p in pf.scraped:
        row = {"Scope": p, **theme_shares(reviews[reviews["platform"] == p])}
        w, _ = pf.weight(p)
        share = platform_negative(p, reviews, pf)[0]
        row["_weight"] = (w or 0) * (share or 0)
        rows.append(row)
        if row["Low-score reviews"] >= MIN_NEGATIVE_REVIEWS and row["_weight"]:
            used.append(p)
        elif share is None:
            left.append(f"{p} (collected reviews have no scores)")
        elif row["Low-score reviews"] >= MIN_NEGATIVE_REVIEWS:
            left.append(f"{p} (no low-score reviews among the normally collected, weight cannot be estimated)")
        else:
            left.append(f"{p} (fewer than {MIN_NEGATIVE_REVIEWS} low-score reviews)")
    all_row = {"Scope": "All platforms", "Low-score reviews": sum(r["Low-score reviews"] for r in rows if r["Scope"] in used),
               "_no_score": all(negative_share(reviews[reviews["platform"] == p]) is None for p in pf.scraped)}
    for theme in THEMES:
        all_row[theme] = wavg([(r[theme], r["_weight"]) for r in rows if r["Scope"] in used])
    no_score = all_row.pop("_no_score")
    table = pd.DataFrame(rows + [all_row]).drop(columns="_weight")
    table["_no_score"] = [negative_share(reviews[reviews["platform"] == r["Scope"]]) is None for r in rows] + [no_score]
    table["Top complaint themes"] = ["; ".join(f"{t} {s:.0%}" for t, s in top_themes(r)) or theme_reason(r)
                          for r in table.to_dict("records")]
    table = table.drop(columns="_no_score")
    table["Merge note"] = [""] * len(rows) + [
        ("Merged: " + ", ".join(used) if used else "no platform has 5 low-score reviews") +
        ("; not merged: " + ", ".join(left) if left else "")]
    return table


def platform_ratings(summary):
    rows = []
    for platform, f in summary.get("platforms", {}).items():
        if f.get("rating") is None and f.get("total_reviews") is None:
            continue  # a platform with only traveller types filled in (e.g. Trip.com, rating same as Ctrip) is not listed
        scale = f.get("scale") or 10
        rating = f.get("rating")
        rows.append({
            "Platform": platform,
            "Platform rating": rating,
            "Scale": scale,
            "On 10-point scale": round(rating * 10 / scale, 2) if rating is not None else None,
            "Published reviews": f.get("total_reviews"),
            "Published minus visible (estimate)": f.get("hidden_reviews"),
            "Note": ("Ctrip and Trip.com are the same set of reviews and are counted once when merging; " if platform in ("Ctrip", "Trip.com")
                   and {"Ctrip", "Trip.com"} <= set(summary.get("platforms", {})) else "") + str(f.get("note", "")),
        })
    return pd.DataFrame(rows, columns=["Platform", "Platform rating", "Scale", "On 10-point scale", "Published reviews", "Published minus visible (estimate)", "Note"])


def subscore_group(name, platform=None):
    text = str(name).lower()
    if platform == "Trip.com" and text == "location":
        return None  # Location on the Trip.com English page is Ctrip's "environment" (surroundings), not a location rating; it does not take part in merging
    for group, words in SUBSCORE_GROUPS:
        if any(w in text for w in words):
            return group
    return None


def subscores(summary, pf):
    rows = []
    for platform, f in summary.get("platforms", {}).items():
        others = summary.get("platforms", {})
        if platform == "Ctrip" and (others.get("Trip.com") or {}).get("subscores"):
            continue  # same set of reviews as Trip.com; if Trip.com has sub-scores (10-point, finer) use Trip.com's
        if platform == "Trip.com" and not f.get("subscores"):
            continue
        scale = f.get("scale") or 10
        w, note = pf.weight(platform)
        for name, value in (f.get("subscores") or {}).items():
            rows.append({"Platform": platform, "Sub-score": name, "Score": value, "Scale": scale,
                         "On 10-point scale": round(value * 10 / scale, 2), "Merged sub-score": subscore_group(name, platform),
                         "Weight (platform reviews)": w, "Note": note})
    return pd.DataFrame(rows, columns=["Platform", "Sub-score", "Score", "Scale", "On 10-point scale", "Merged sub-score", "Weight (platform reviews)", "Note"])


def traveller_mix(summary, reviews):
    """Traveller-type counts published by the platform (numbers on the traveller-type filter in the review area). When both Ctrip and Trip.com have them, only Trip.com's are used.
    If a platform does not publish counts (e.g. Trip.com's filter has no numbers) but all reviews were collected (all_reviews_scraped is true),
    traveller types are counted from the collected reviews."""
    platforms = summary.get("platforms", {})
    rows = []
    has_counts = {p for p, f in platforms.items() if f.get("traveller_counts")}
    if "Trip.com" in has_counts or "Ctrip" in has_counts:
        has_counts |= {"Trip.com", "Ctrip"}
    for p in sorted(reviews["platform"].dropna().unique()):
        flag = (platforms.get(p) or {}).get("all_reviews_scraped")
        if p == "Trip.com" and flag is None:
            flag = (platforms.get("Ctrip") or {}).get("all_reviews_scraped")
        if p in has_counts or flag is not True:
            continue
        counts = reviews.loc[reviews["platform"] == p, "traveler_type"].fillna("").astype(str).str.strip()
        counts = counts[counts != ""].value_counts()
        total = counts.sum()
        for label, n in counts.items():
            rows.append({"Platform": f"{p} (counted from collected reviews)", "Traveller type (original)": label, "Guest group": traveller_group(label),
                         "Count": int(n), "Share": n / total if total else None})
    for platform, f in platforms.items():
        if platform == "Ctrip" and (platforms.get("Trip.com") or {}).get("traveller_counts"):
            continue
        counts = {k: v for k, v in (f.get("traveller_counts") or {}).items() if v}
        total = sum(counts.values())
        for label, n in counts.items():
            rows.append({"Platform": platform, "Traveller type (original)": label, "Guest group": traveller_group(label),
                         "Count": n, "Share": n / total if total else None})
    return pd.DataFrame(rows, columns=["Platform", "Traveller type (original)", "Guest group", "Count", "Share"])


def platform_topics(summary):
    rows = [{"Platform": platform, **t} for platform, f in summary.get("platforms", {}).items() for t in f.get("topics") or []]
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["Platform"])


def norm_lang(code):
    """Language codes are lower-cased and reduced to the main language: zh-HK, zh_CN, ZH -> zh; cn -> zh."""
    c = re.split(r"[-_]", str(code).strip().lower())[0]
    return {"": "unknown", "cn": "zh"}.get(c, c)


def norm_langs(codes):
    return [norm_lang(x) for x in codes]


def complete_sample(reviews, summary):
    """Sample for time trends: only platforms and languages whose reviews were fully collected (all_reviews_scraped true or a language list in summary.json).
    On a platform that was not fully collected, few old reviews only means they were not collected, and including them would make "fewer problems before" and "worse recently" look real. Returns (sample, note on platforms used)."""
    flags = {p: f.get("all_reviews_scraped") for p, f in summary.get("platforms", {}).items()}
    keep = pd.Series(False, index=reviews.index)
    used = []
    for p in reviews["platform"].dropna().unique():
        flag = flags.get(p)
        if p == "Trip.com" and flag is None:
            flag = flags.get("Ctrip")
        m = reviews["platform"] == p
        if flag is True:
            keep |= m
            used.append(p)
        elif isinstance(flag, list) and flag:
            keep |= m & reviews["review_language"].isin(norm_langs(flag))
            used.append(p)
    return reviews[keep], used


def periods(reviews, today, complete):
    """Last 90 days vs 91 days to 1 year ago, using collected reviews (with scores).
    Checked per language: only languages whose collected reviews reach back a full year take part; otherwise the "earlier" period of that language is small only because it was not fully collected;
    if a language really has under a year of reviews (the list was paged to the end), mark all_reviews_scraped (true or a list of language codes) in summary.json.
    A platform takes part in the comparison only if its languages have 10+ reviews in both periods."""
    cut_recent = today - pd.Timedelta(days=RECENT_DAYS)
    cut_year = today - pd.Timedelta(days=COMPARE_DAYS)
    out = {}
    for p in sorted(reviews["platform"].dropna().unique()):
        df = reviews[(reviews["platform"] == p) & reviews["score10"].notna()]
        flag = complete.get(p)
        if flag is None and p == "Trip.com":
            flag = complete.get("Ctrip")
        langs = []
        for source, part in df.groupby("source file"):
            lang = str(part["review_language"].iloc[0])
            done = flag is True or (isinstance(flag, list) and lang in norm_langs(flag))
            if done or (part["date"].notna().any() and part["date"].min() <= cut_year):
                langs.append(lang if lang != "unknown" else source)
        covered = [src for src, part in df.groupby("source file")
                   if (str(part["review_language"].iloc[0]) if str(part["review_language"].iloc[0]) != "unknown" else src) in langs]
        use = df[df["source file"].isin(covered)]
        recent = use[use["date"] >= cut_recent]
        earlier = use[(use["date"] < cut_recent) & (use["date"] >= cut_year)]
        if df.empty:
            reason = "collected reviews have no scores"
        elif df["date"].isna().all():
            reason = "collected reviews have no dates"
        elif not langs:
            reason = "collected reviews do not reach back one year"
        elif len(recent) < MIN_PERIOD_REVIEWS or len(earlier) < MIN_PERIOD_REVIEWS:
            reason = f"{len(recent)} reviews in the last 90 days, {len(earlier)} earlier; one period has fewer than {MIN_PERIOD_REVIEWS}"
        else:
            reason = ""
        out[p] = {"recent": recent, "earlier": earlier, "Language": langs, "excluded reason": reason}
    return out


def platform_table(reviews, pf, themes, per):
    """One row per platform (scraped and score-only ones), to see which platform raises or lowers the overall result."""
    t = themes.set_index("Scope")
    names = list(pf.scraped) + [p for p in pf.rating_platforms() if p not in pf.scraped
                                and not (p == "Ctrip" and "Trip.com" in pf.scraped)]
    rows = []
    for p in names:
        df = reviews[reviews["platform"] == p]
        w, note = pf.weight(p)
        joined = p in per and not per[p]["excluded reason"]
        rows.append({
            "Platform": "Ctrip (reviews collected from Trip.com)" if p == "Trip.com" and "Ctrip" in pf.info else p,
            "Platform rating (10-point)": pf.rating10(p),
            "Published reviews": pf.published(p),
            "Collected reviews": len(df) if len(df) else "score only, no reviews collected",
            "Latest collected review": df["date"].max() if len(df) else None,
            "Earliest collected review": df["date"].min() if len(df) else None,
            "Average score of collected reviews (10-point)": round(df["score10"].mean(), 2) if df["score10"].notna().any() else None,
            "Low-score share": platform_negative(p, reviews, pf)[0] if len(df) else None,
            "How low-score share is computed": platform_negative(p, reviews, pf)[1] if len(df) else "",
            f"Average score, last {RECENT_DAYS} days": round(per[p]["recent"]["score10"].mean(), 2) if joined else None,
            "Average score, 91 days to 1 year ago": round(per[p]["earlier"]["score10"].mean(), 2) if joined else None,
            "In time comparison": ("Yes (" + ", ".join(per[p]["Language"]) + ")") if joined else
                        ("No: " + per[p]["excluded reason"] if p in per else "No: no reviews collected"),
            "Low-score reviews": t.loc[p, "Low-score reviews"] if p in t.index else None,
            "Top complaint themes": t.loc[p, "Top complaint themes"] if p in t.index else "",
            "Note": note,
        })
    return pd.DataFrame(rows)


def overall_table(reviews, pool, pf, ratings, subs, travellers, themes, impact, per, today, summary):
    """One summary page: the combined result of all platforms. Each row says which platforms were used and how many reviews."""
    rows = []
    add = lambda item, value, how: rows.append({"Item": item, "Combined result": value, "Platforms used and how": how})
    notes = lambda ps: "; ".join(n for n in (pf.weight(p)[1] for p in ps) if n)
    info = summary.get("platforms", {})

    # 0. Platform scope
    included = list(dict.fromkeys([pf.disp(p) for p in pf.scraped] + [pf.disp(p) for p in pf.rating_platforms()]))
    scope = "Included: " + (", ".join(included) or "none")
    if summary.get("not_listed"):
        scope += "; not found this time, to be verified: " + ", ".join(summary["not_listed"])
    if summary.get("not_included"):
        scope += "; not included this time: " + ", ".join(summary["not_included"])
    add("Platform scope", scope, "Not found this time = the hotel cannot be found on that platform, ask the hotel whether it is listed; not included this time = a platform this version of the skill cannot collect")

    # 1. Rating and review counts
    rated = pf.rating_platforms()
    used = [p for p in rated if pf.weight(p)[0]]
    overall = wavg([(pf.rating10(p), pf.weight(p)[0]) for p in used])
    if not summary:
        how = "summary.json not filled in, no platform-published rating"
    elif not rated:
        how = "summary.json has no rating for any platform"
    else:
        how = ("Platform-published ratings converted to a 10-point scale and weighted by published review counts: " + ", ".join(
            f"{pf.disp(p)} {pf.rating10(p):g} ({int(pf.weight(p)[0])} reviews)" for p in used)
               + (f" ({notes(rated)})" if notes(rated) else ""))
    add("Overall rating (10-point)", round(overall, 2) if overall is not None else None, how)
    add("Total published reviews", int(pf.all_published()) if pf.all_published() else None,
        (", ".join(pf.disp(p) for p in rated if pf.published(p)) + " published review counts added up (Ctrip and Trip.com counted once)") if rated else "summary.json not filled in")
    add("Reviews analysed", len(reviews), "Collected reviews with text, after de-duplication (topped-up low-score reviews excluded): " + ", ".join(
        f"{pf.disp(p)} {pf.counts[p]} reviews" for p in pf.scraped))

    # 2. Low-score reviews
    negs = {p: platform_negative(p, reviews, pf) for p in pf.scraped}
    usable = [p for p in pf.scraped if negs[p][0] is not None]
    neg = wavg([(negs[p][0], pf.weight(p)[0]) for p in usable])
    add("Low-score share of analysed reviews", (reviews["Low score"] == True).sum() / len(reviews) if len(reviews) else None,  # noqa: E712
        f"Share of the {len(reviews)} normally collected reviews that score below 6 (the report uses this; topped-up low-score reviews are not in the ratio)")
    add("Low-score share (weighted by platform-published negative counts)", neg,
        ("Weighted by platform review counts: " + "; ".join(negs[p][1] for p in usable)) if usable else "no collected review has a score")
    all_themes = themes.set_index("Scope").loc["All platforms"]
    merged = [p for p in pf.scraped if p in str(all_themes["Merge note"]).split("; ")[0]]
    neg_pool = reviews[(reviews["Low score"] == True) & reviews["platform"].isin(merged)]  # noqa: E712
    year_ago = today - pd.Timedelta(days=COMPARE_DAYS)
    # Low-score count and timing: low-score reviews of all platforms count (the theme merge uses only platforms with 5+ low-score reviews, which is a different matter)
    all_negs = reviews[reviews["Low score"] == True]  # noqa: E712  ratios and counts use only normally collected reviews; topped-up low-score reviews are evidence quotes only
    recent_all = int((all_negs["date"] >= year_ago).sum())
    add("Low-score timing", f"{len(all_negs)} low-score reviews: {recent_all} in the past year, {len(all_negs) - recent_all} earlier",
        "Low-score reviews collected from all platforms (below 6, normally collected, topped-up low-score reviews excluded): " + ", ".join(
            f"{pf.disp(p)} {n} reviews" for p, n in all_negs["platform"].value_counts().items()) + "; few in the past year, so these problems may no longer exist")
    tops = top_themes(all_themes)
    if not tops:
        value = ("cannot merge" if all_themes["Low-score reviews"] == 0 and "not merged" in all_themes["Merge note"]
                 else all_themes["Top complaint themes"])
        add("Top complaint themes", value, all_themes["Merge note"])
    for i, (theme, share) in enumerate(tops, 1):
        k = int(neg_pool["Complaint themes"].str.contains(re.escape(theme)).sum())
        add(f"Top complaint theme {i}", f"{theme} ({share:.0%}, mentioned in {k} of {len(neg_pool)} low-score reviews)",
            "Share of low-score reviews complaining about this theme, platforms weighted by estimated low-score count. " + all_themes["Merge note"])
    worst = impact.dropna(subset=["Impact score"])
    worst = worst[worst["Impact score"] <= -MIN_IMPACT].head(3)
    impact_platforms = [c[:-len("_complaints")] for c in impact.columns if c.endswith("_complaints")]
    impact_how = ("Average score of reviews complaining about this theme minus the average of the other reviews, computed inside each platform and then merged; correlation, not causation; "
                  f"themes complained about fewer than {MIN_THEME_MENTIONS} times are not counted") if impact_platforms else "collected reviews have no scores"
    if not impact_platforms:
        add("Themes that lower the score most", "cannot compute", impact_how)
    elif worst.empty:
        add("Themes that lower the score most", f"no theme is complained about {MIN_THEME_MENTIONS}+ times and lowers the score by {MIN_IMPACT:g}+ points", impact_how)
    for i, row in enumerate(worst.to_dict("records"), 1):
        add(f"Theme that lowers the score {i}", f"{row['Theme']} (the {row['Reviews complaining']} reviews complaining about it average {-row['Impact score']:.1f} points lower)", impact_how)

    # 2b. What guests praise
    pt = praise_table(reviews)
    n_good = pt.attrs.get("n", 0)
    for i, row in enumerate(pt[pt["Reviews praising it"] > 0].head(3).to_dict("records"), 1):
        add(f"Most praised {i}", f"{row['Aspect']} ({row['Share']:.0%}, {row['Reviews praising it']} of the {n_good} analysed reviews)",
            "Share of non-low-score Chinese and English reviews praising this aspect (a praising wording near the aspect word is needed; automatic, misses some)"
            + (f"; {pt.attrs['unreadable']} Korean/Japanese reviews were not read" if pt.attrs.get("unreadable") else ""))

    # 3. Sub-scores: similar sub-scores of platforms converted to a 10-point scale, weighted by platform review counts
    grouped = subs.dropna(subset=["Merged sub-score"]) if len(subs) else subs
    shown = False
    total_pub = pf.all_published()
    for group, _ in SUBSCORE_GROUPS:
        part = grouped[grouped["Merged sub-score"] == group] if len(grouped) else grouped
        part = part[part["Weight (platform reviews)"].notna()] if len(part) else part
        if part.empty:
            continue
        shown = True
        value = wavg([(row["On 10-point scale"], row["Weight (platform reviews)"]) for row in part.to_dict("records")])
        n = int(part["Weight (platform reviews)"].sum())
        small = total_pub and n / total_pub < SMALL_SHARE
        add(f"Sub-score: {group} (10-point)" + (" (small sample, for reference)" if small else ""), round(value, 2),
            "; ".join(f"{pf.disp(row['Platform'])} {row['Sub-score']} {row['On 10-point scale']:g} ({int(row['Weight (platform reviews)'])} reviews)" for row in part.to_dict("records"))
            + ", weighted by platform review counts" + (f", {n} reviews in total, {n / total_pub:.0%} of the published reviews" if total_pub else ""))
    if not shown:
        add("Sub-scores", "none", "not published by the platforms or not filled in summary.json")

    # 4. Guest mix
    # Guest mix uses the traveller types marked on collected reviews (one review one type), the same denominator as other guest-group numbers in the report;
    # counts on the platform filters (on Booking.com one review can count for several types) stay only in the "Traveller type" sheet for reference
    typed = reviews[reviews["Guest group"].fillna("").isin(["Business", "Family", "Couple", "Friends/group", "Solo"])]
    if len(typed):
        counts = typed["Guest group"].value_counts()
        total = int(counts.sum())
        by_pf = ", ".join(f"{pf.disp(p)} {n} reviews" for p, n in typed["platform"].value_counts().items())
        for group, n in counts.items():
            add(f"Guest: {group}", n / total, f"Traveller types marked on collected reviews, one review one type, {total} reviews ({by_pf}); counts on the platform filters are in the 'Traveller type' sheet")
    else:
        add("Guest mix", "none", "no collected review marks a traveller type")

    # 5. Last 90 days vs 91 days to 1 year ago: collected reviews, only platforms and languages that reach back a year
    joined = [p for p in per if not per[p]["excluded reason"]]
    left = [f"{pf.disp(p)} ({per[p]['excluded reason']})" for p in per if per[p]["excluded reason"]]
    joined_desc = "; ".join(pf.disp(p) + " (" + ", ".join(per[p]["Language"]) + ")" for p in joined) or "none"
    how = (f"As of {today.date()}, computed from collected reviews (a different method from the platform-published overall rating above; do not compare them); "
           f"taking part: {joined_desc}, weighted by platform review counts"
           + (f"; not taking part: {'; '.join(left)}" if left else ""))
    if not joined:
        add("Last 90 days vs 91 days to 1 year ago", "cannot compare", how)
    else:
        stats = {}
        for key in ["recent", "earlier"]:
            n = sum(len(per[p][key]) for p in joined)
            k = sum(int(per[p][key]["Low score"].dropna().astype(bool).sum()) for p in joined)
            stats[key] = (wavg([(per[p][key]["score10"].mean(), pf.weight(p)[0]) for p in joined]),
                          wavg([(negative_share(per[p][key]), pf.weight(p)[0]) for p in joined]), n, k)
        (ra, rn, rc, rk), (ea, en, ec, ek) = stats["recent"], stats["earlier"]
        if min(rc, ec) < FEW_REVIEWS:
            how += f"; one period has under {FEW_REVIEWS} reviews, so one or two extreme reviews can change the result a lot; read the low-score detail before concluding"
        add(f"Average score, last {RECENT_DAYS} days (collected reviews, 10-point)", f"{ra:.2f} ({rc} reviews)", how)
        add("Average score, 91 days to 1 year ago (collected reviews, 10-point)", f"{ea:.2f} ({ec} reviews)", how)
        change = f"{ra - ea:+.2f} points"
        if rc < FEW_REVIEWS:
            recent_all = pd.concat([per[p]["recent"] for p in joined])
            low = recent_all["score10"].min()
            rest = recent_all["score10"].drop(recent_all["score10"].idxmin())
            if len(rest):
                change += f" (lowest review in the last {RECENT_DAYS} days is {low:g} points; without it the change is {rest.mean() - ea:+.2f} points)"
        add(f"Change in average score (last {RECENT_DAYS} days - earlier)", change, how)
        add(f"Low-score reviews, last {RECENT_DAYS} days", f"{rn:.0%} ({rk} of {rc} reviews)", how)
        add("Low-score reviews, 91 days to 1 year ago", f"{en:.0%} ({ek} of {ec} reviews)", how)
    return pd.DataFrame(rows)


MIN_GROUP_REVIEWS = 10  # a guest group with fewer than 10 reviews is not listed separately for praise


def praise_table(reviews):
    """What guests praise: number of reviews praising each aspect; the denominator of the shares is all analysed reviews (same as other shares in the report);
    low-score reviews do not count as praise (polite words inside them do not count); Korean/Japanese reviews cannot be read but are in the denominator, so shares are slightly low. Then split by guest group."""
    unreadable = reviews["analysis text"].astype(str).str.contains(UNREADABLE) & (reviews["Low score"] != True)  # noqa: E712
    good = reviews
    n = len(good)
    rows = []
    groups = [g for g, c in good["Guest group"].value_counts().items() if c >= MIN_GROUP_REVIEWS and g != "Other"]
    for theme in PRAISE:
        hit = good["Praise aspects"].str.contains(re.escape(theme))
        row = {"Aspect": theme, "Reviews praising it": int(hit.sum()), "Share": hit.mean() if n else None}
        for g in groups:
            sub = good[good["Guest group"] == g]
            row[f"{g} ({len(sub)} reviews) share"] = sub["Praise aspects"].str.contains(re.escape(theme)).mean()
        rows.append(row)
    table = pd.DataFrame(rows).sort_values("Reviews praising it", ascending=False)
    table.attrs["n"] = n
    table.attrs["unreadable"] = int(unreadable.sum())
    return table


def theme_impact(reviews):
    """Theme impact score: average score of reviews complaining about a theme minus the average of reviews not complaining about it (10-point).
    Only complaining clauses count (same judgement as the complaint themes); praise does not count, otherwise praise like "very good service" would distort the result.
    The difference is computed inside each platform first and then averaged weighted by each platform's number of mentions, so "the theme appears mostly on a low-scoring platform" is not read as lowering the score.
    Only normally collected reviews, not topped-up low-score reviews (they would skew the score). Correlation, not causation."""
    scored = reviews[reviews["score10"].notna()]
    platforms = sorted(scored["platform"].dropna().unique())
    rows = []
    for theme in THEMES:
        row = {"Theme": theme}
        total_hits, weighted, weight = 0, 0.0, 0
        for p in platforms:
            df = scored[scored["platform"] == p]
            hit = df["Complaint themes"].str.contains(re.escape(theme))
            n = int(hit.sum())
            total_hits += n
            diff = None
            if n and (~hit).any():
                diff = df.loc[hit, "score10"].mean() - df.loc[~hit, "score10"].mean()
                weighted += diff * n
                weight += n
            row[f"{p}_complaints"] = n
            row[f"{p}_impact"] = round(diff, 2) if diff is not None and n >= MIN_THEME_MENTIONS else None
        row["Reviews mentioning"] = total_hits
        if total_hits < MIN_THEME_MENTIONS:
            row["Impact score"], row["Note"] = None, f"fewer than {MIN_THEME_MENTIONS} complaints, no impact score"
        elif not weight:
            row["Impact score"], row["Note"] = None, "all reviews complain about this theme, nothing to compare with"
        else:
            row["Impact score"], row["Note"] = round(weighted / weight, 2), ""
        rows.append(row)
    for row in rows:
        row["Reviews complaining"] = row.pop("Reviews mentioning")
    table = pd.DataFrame(rows, columns=["Theme", "Impact score", "Reviews complaining", "Note"]) if not platforms else pd.DataFrame(rows)
    front = ["Theme", "Impact score", "Reviews complaining", "Note"]
    table = table[front + [c for c in table.columns if c not in front]]
    return table.sort_values("Impact score", na_position="last")


SHEET_NOTES = {
    "Summary": "Combined result of all platforms: overall rating, total reviews, low-score share, top complaint themes, themes that lower the score most, sub-scores, guest mix, last 90 days vs earlier. Which platforms were used and how is in the third column",
    "Platform comparison": "One row per platform (score-only platforms too), to see which platform raises or lowers the overall result",
    "Platform ratings": "Overall rating, review count and published-minus-visible count (estimate) on each platform page, including score-only platforms and platforms with no reviews collected",
    "Sub-scores": "Sub-scores published by each platform; 'Merged sub-score' is the group used when merging (empty = not merged, e.g. Ctrip 'environment' and Trip.com Location, which is Ctrip's 'environment')",
    "Traveller type": "Counts on the traveller-type filter of each platform; where a platform does not publish counts and all reviews were collected (e.g. Trip.com), they are counted from collected reviews and the platform name says so",
    "Complaint themes by platform": "Each platform and all platforms: number of low-score reviews and the share mentioning each theme (Booking: disliked part only)",
    "What guests praise": "Among all analysed reviews (low-score reviews do not count as praise): number and share praising each aspect; the later columns split by guest group (listed with 10+ reviews). Not detecting praise does not mean the guest was unhappy",
    "What each guest group mentions": "The denominator is all reviews of that guest group (mentions inside low-score reviews are excluded): number and share mentioning each thing (facilities, location, trip purpose). 'Basics' are what every hotel should do, 'Selling points' are not available everywhere. Shows what the main guest groups care about and what the listing should say",
    "Room types": "Room types with 10+ reviews: count, average score, share with complaints, most complained-about problem, most common guests. Shows whether one room type is complained about again and again",
    "Review languages": "By review language: count, average score, platforms, traveller types. Review language is not nationality, only a clue",
    "Foreign-language mentions": "Review language is not nationality, only a clue. The denominator is all reviews in that language (mentions inside low-score reviews are excluded): number and share mentioning each thing (Korean and Japanese are counted from the Chinese translations in translations.json)",
    "Theme impact scores": f"Each theme: average score of reviews complaining about it minus the average of the other reviews (praise does not count), computed inside each platform and then weighted by complaints. The more negative, the lower the scores of reviews complaining about it (correlation, not causation). Fewer than {MIN_THEME_MENTIONS} complaints: not counted",
    "Platform topics": "Topic statistics of the platforms themselves (Ctrip hot tags have mention counts only; Agoda topic analysis has positive and negative shares)",
    "Low-score detail": "All low-score reviews and the themes they hit (the 'Source' column marks topped-up low-score reviews, which are evidence quotes only), newest first",
    "All reviews": "All collected reviews (topped-up low-score reviews excluded), newest first",
}


def write_excel(tables, out):
    readme = pd.DataFrame([
        ["Low-score review", f"The rule of this report: the guest's score converted to a 10-point scale is below {NEGATIVE_BELOW:g} (below 3 on a 5-point scale). The platform's own 'negative' count (e.g. Ctrip 'negative (7)') is listed separately and may differ"],
        ["Merging", "Platform numbers are merged weighted by published review counts, so platforms with more reviews weigh more; Ctrip and Trip.com are the same set of reviews, counted once. Which platforms were used is in the third column of 'Summary'"],
        ["Collected reviews", "Only reviews with text are collected (rating-only reviews are not); each language is counted by the number collected, not weighted by the real number of reviews per language"],
        ["Low-score share", "The report always uses 'share of analysed reviews below 6' (Summary: 'Low-score share of analysed reviews'). 'Platform comparison' and the extra 'weighted by platform-published negative counts' row in Summary use the platform's own negative count (e.g. Ctrip 'negative (7)') and are kept for reference and as weights for the theme merge"],
        ["Complaint themes", "Share of low-score reviews complaining about a theme: Booking uses only the disliked part; other platforms are cut into clauses by punctuation and spaces, and a theme word counts only if a dissatisfied wording (not, no, poor, broken, dirty, noisy, expensive ...) is within 6 characters. Automatic, with some false hits; one review can mention several themes. Reviews topped up with the 'lowest first' sort are evidence quotes only (marked in the 'Source' column of 'Low-score detail') and enter no ratio or count"],
        ["Too few low-score reviews", f"When a scope has fewer than {MIN_NEGATIVE_REVIEWS} low-score reviews, no theme conclusion is listed and it does not take part in the merged theme shares"],
        ["Published minus visible", "Estimate: reviews may be hidden, folded, or rating-only, so not all are necessarily hidden by the platform; the content is unknown and not part of this analysis"],
    ] + [[name, SHEET_NOTES.get(name, "")] for name in tables], columns=["Item", "Notes"])
    # Ctrip and Trip.com are the same set of reviews; the report says "Ctrip" only
    def _cn(v):
        return v.replace("Trip.com", "Ctrip") if isinstance(v, str) else v
    readme = readme.map(_cn)
    tables = {name: t.rename(columns=_cn).map(_cn) for name, t in tables.items()}
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        readme.to_excel(writer, sheet_name="Notes", index=False)
        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name[:31], index=False)
        for name, sheet in writer.sheets.items():
            for col in sheet.columns:
                width = max(len(str(c.value or "")) for c in col[:50])
                sheet.column_dimensions[col[0].column_letter].width = min(max(10, width * 1.6), 60)
                header = str(col[0].value or "")
                if "share" in header.lower() or header in THEMES:
                    for c in col[1:]:
                        c.number_format = "0%"
            if name == "Summary":  # share and guest rows are shown as percentages
                for row in sheet.iter_rows(min_row=2):
                    if "share" in str(row[0].value).lower() or str(row[0].value).startswith("Guest: "):
                        row[1].number_format = "0%"
            sheet.freeze_panes = "B2"


def main(folder, today=None):
    folder = Path(folder)
    reviews = load_reviews(folder, "reviews_*.json")
    if reviews.empty:
        sys.exit("No reviews_*.json found; finish the scraping step first.")
    low = load_reviews(folder, "lowscore_*.json")
    reviews["Source"] = "Collected normally"
    if len(low):
        low["Source"] = "Topped-up low-score review (evidence quote only, in no ratio)"
    pool = pd.concat([reviews, low], ignore_index=True) if len(low) else reviews
    pool = pool.drop_duplicates(subset=["platform", "text", "date", "score10", "reviewer_country", "room_type"])
    summary = load_summary(folder)
    today = pd.Timestamp(today or pd.Timestamp.today().normalize())

    names = reviews["hotel"].dropna()
    hotel = summary.get("hotel") or (names.iloc[0] if len(names) else "Hotel")
    pf = Platforms(summary, reviews)
    ratings = platform_ratings(summary)
    subs = subscores(summary, pf)
    travellers = traveller_mix(summary, reviews)
    themes = theme_table(pool, reviews, pf)
    impact = theme_impact(reviews)
    complete = {p: f.get("all_reviews_scraped") for p, f in summary.get("platforms", {}).items()}
    per = periods(reviews, today, complete)
    # reviewer_country is not output: when a Trip.com guest did not fill in a country, this field shows the nickname
    columns = ["platform", "date", "score10", "traveler_type", "Guest group", "room_type",
               "review_language", "title", "text", "Translation", "Complaint themes", "Praise aspects", "Source"]
    tables = {
        "Summary": overall_table(reviews, pool, pf, ratings, subs, travellers, themes, impact, per, today, summary),
        "Platform comparison": platform_table(reviews, pf, themes, per),
        "Platform ratings": ratings,
        "Sub-scores": subs,
        "Traveller type": travellers,
        "Complaint themes by platform": themes,
        "Theme impact scores": impact,
        "What guests praise": praise_table(reviews),
        "What each guest group mentions": segment_mentions(reviews, summary),
        "Review languages": language_summary(reviews),
        "Room types": room_type_table(reviews),
        "Foreign-language mentions": language_mentions(reviews, summary),
        "Platform topics": platform_topics(summary),
        "Low-score detail": pool[pool["Low score"] == True].reindex(columns=columns).sort_values("date", ascending=False),  # noqa: E712
        "All reviews": reviews.reindex(columns=columns).sort_values("date", ascending=False)
                      .rename(columns={"Complaint themes": "Complaint themes (automatic, for reference only)"}),
    }
    out = folder / f"review_analysis_{BAD_FILENAME.sub('_', str(hotel))}.xlsx"
    write_excel(tables, out)
    pd.set_option("display.width", 200)
    pd.set_option("display.max_colwidth", 80)
    print(tables["Summary"][["Item", "Combined result"]].to_string(index=False))
    print(f"\nSaved: {out}")
    return {"tables": tables, "reviews": reviews, "pool": pool, "summary": summary, "today": today,
            "hotel": hotel, "excel": out, "platforms": pf}


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
