"""
酒店点评整合分析（单家酒店，多平台）
=====================================

目的：把一家酒店在各平台的点评合成一份总结（"综合汇总"表），各平台的明细放在后面的表里。

方法：
- 差评按客人自己打的分数判断：10 分制低于 6 分（中文点评不用翻译）。
- 合并各平台时，按平台公布的点评数加权；携程和携程国际版（Trip.com）是同一套点评，只算一次。
- 12 个差评主题，中英文关键词都认（繁体先转简体）。
- 只分析一家酒店。
- 差评主题 = 差评里提到该主题的比例，只看客人不满意的部分（Booking 只看差评段，差评段为空的不计），不看标题；
  按"低分优先"补抓的点评只当原文证据（差评明细里能看到），不进任何比例和条数。
- 近 90 天和 91 天到 1 年前对比：用抓到的点评，按语言（抓取文件）检查，只用已覆盖到 1 年前的语言，
  否则"更早"那一段少是因为没抓全，不是真的少；参加的平台两段都要有 10 条以上。
- 分项评分：各平台同类的分项（卫生/Cleanliness 等）换算成 10 分制后加权合并。
- 主题影响分 = 提到某主题的点评平均分 − 没提到的平均分，先在每个平台内算，再按提到条数加权；
  只用正常抓取的点评。这是相关，不是因果。
- 平台公布的评分、点评数、分项评分、住客类型、平台话题，从 summary.json 读取。

输入（放在同一个数据文件夹里）：
  reviews_*.json   抓取步骤 collect() 的结果：{platform, hotel, url, language, rows:[...]} 或直接是 rows 列表
  lowscore_*.json  按"低分优先"补抓的点评（可选）
  summary.json     各平台页面上读到的汇总数字（格式见 SKILL.md，可选）

输出：点评分析_<酒店名>.xlsx

运行：python analyze.py <数据文件夹>     需要 pandas 2.1 以上、openpyxl（见 requirements.txt）
"""

import json
import re
import sys
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# 规则
# ---------------------------------------------------------------------------

# 差评：10 分制低于 6 分（5 分制低于 3 分），在实测中与携程页面上"差评"的条数一致；Agoda 6 分以下也不给评价词。
NEGATIVE_BELOW = 6.0
MIN_NEGATIVE_REVIEWS = 5   # 差评少于 5 条时，主题占比太不稳定，不下结论
RECENT_DAYS = 90           # "近 90 天"
COMPARE_DAYS = 365         # 和"91 天到 1 年前"对比
MIN_THEME_MENTIONS = 5     # 一个主题被提到少于 5 次时，不算影响分
MIN_PERIOD_REVIEWS = 10    # 时间对比里，一段少于 10 条的平台不参加
SMALL_SHARE = 0.10         # 分项只来自不到 10% 的公布点评时，标"样本少"
FEW_REVIEWS = 30           # 时间对比里一段不到 30 条时提醒"单条点评影响大"
MIN_IMPACT = 0.3           # "拉低分数最多的主题"只列低 0.3 分以上的

# 12 个差评主题：英文关键词 + 中文关键词（中文先把繁体字转成简体再匹配）
THEMES = {
    "设施老旧/需翻新": {
        "en": ["old", "dated", "outdated", "tired", "worn", "renovat", "refurb", "run down", "run-down", "shabby"],
        "zh": ["老旧", "陈旧", "破旧", "残旧", "老化", "翻新", "年代感", "有年头", "古老", "又旧又破", "设施旧", "装修旧",
               "比较旧", "很旧", "有点旧", "太旧", "过时", "设备旧", "感觉旧", "旧了"],
    },
    "卫生/异味": {
        "en": ["dirty", "unclean", "stain", "mould", "mold", "smell", "musty", "mildew", "dust", "cockroach", "bug", "hair"],
        "zh": ["脏", "不干净", "不清洁", "不卫生", "卫生差", "卫生不好", "卫生一般", "清洁度", "污渍", "污垢", "发霉", "霉味",
               "霉斑", "霉菌", "异味", "怪味", "臭味", "很臭", "发臭", "恶臭", "有味道", "味道大", "味道重", "好大味", "烟味",
               "潮味", "腥味", "潮湿", "灰尘", "头发", "毛发", "蟑螂", "虫", "血迹"],
    },
    "房间小": {
        "en": ["small", "tiny", "cramped", "compact", "no space"],
        "zh": ["房间小", "很小", "太小", "有点小", "比较小", "偏小", "不大", "面积小", "空间小", "狭小", "局促", "房间拥挤", "好细"],
    },
    "员工/服务": {
        "en": ["rude", "unhelpful", "unfriendly", "impolite", "unprofessional", "attitude", "ignored", "poor service",
               "bad service", "slow service", "terrible service", "horrible service", "service was poor", "service was bad",
               "service was slow", "staff were rude", "staff was rude", "staff were unhelpful", "staff was unhelpful"],
        "zh": ["服务差", "服务不好", "服务很差", "服务一般", "服务态度", "态度差", "态度不好", "态度很差", "态度恶劣",
               "态度一般", "不专业", "不耐烦", "没人理", "不理", "冷漠", "敷衍", "前台态度", "员工态度", "工作人员态度",
               "服务员态度", "前台不", "前台很慢", "前台太慢"],
        # zh_ctx：本身不带好坏的词，只在抱怨的短句里才算差评主题（不用于主题影响分）
        "zh_ctx": ["服务", "前台", "员工", "工作人员", "服务员", "态度", "管理", "保安"],
    },
    "入住/等候": {
        "en": ["check-in", "check in", "checkin", "queue", "waited", "waiting", "long line"],
        "zh": ["办理入住", "入住慢", "排队", "等了", "等候", "等很久", "等半天"],
    },
    "噪音": {
        "en": ["noise", "noisy", "loud", "thin wall", "hum", "humming"],
        "zh": ["噪音", "噪声", "很吵", "太吵", "有点吵", "好吵", "吵闹", "吵杂", "吵", "嘈", "隔音", "声音大", "很响",
               "太响", "一样响", "装修声"],
    },
    "维修（空调、热水、卫浴）": {
        "en": ["air con", "aircon", "air-con", "air conditioning", "broken", "leak", "not working", "toilet", "shower", "water", "flush"],
        "zh": ["空调", "冷气", "热水", "水温", "水压", "忽冷忽热", "漏水", "滴水", "水龙头", "堵", "坏", "不能用", "用不了",
               "故障", "马桶", "淋浴", "花洒", "下水", "排风", "忽冷忽烫", "忽大忽小", "不恒温", "排水慢", "下水慢",
               "反味", "不制冷", "热醒"],
        "zh_ctx": ["投屏", "电视", "排水", "地漏", "冰箱", "洗衣机", "温控", "水压"],
    },
    "早餐/餐饮": {
        "en": ["breakfast", "buffet", "restaurant", "food", "coffee"],
        "zh": ["早餐", "早饭", "自助餐", "餐厅", "咖啡"],
    },
    "设施（泳池、健身房、电梯）": {
        "en": ["pool", "gym", "lift", "elevator", "facilit", "lobby"],
        "zh": ["泳池", "游泳池", "健身", "电梯", "𨋢"],
    },
    "订房信息和实际不符": {
        "en": ["not as described", "different from the photo", "different from photos", "extra charge", "hidden fee",
               "not included", "misleading"],
        "zh": ["与描述不符", "和描述不符", "跟描述不一样", "与图片不符", "和图片不一样", "照片不一样", "图片不一样", "实物不符",
               "货不对板", "只含一人", "只含一个人", "早餐不含", "不含早", "另外收费", "额外收费", "加收", "另收", "没有提前告知",
               "没提前说", "没有提醒", "没提醒"],
    },
    "价格/押金/性价比": {
        "en": ["price", "expensive", "overpriced", "value", "worth", "deposit"],
        "zh": ["价格", "太贵", "很贵", "偏贵", "有点贵", "比较贵", "价格贵", "不值", "性价比", "押金", "按金", "价钱"],
    },
    "位置/交通": {
        "en": ["far from", "no shuttle", "long walk", "hard to find", "no parking", "expensive parking", "parking was expensive",
               "inconvenient location", "location was inconvenient"],
        "zh": ["位置偏", "太远", "很远", "比较远", "偏僻", "交通不便", "交通不方便", "出行不方便", "离地铁远", "地铁远",
               "打车难", "不好打车", "打不到车", "停车难", "不好停车", "没有停车", "停车不便", "停车费", "不好找", "难找"],
        "zh_ctx": ["位置", "交通", "地铁", "公交", "打车", "停车", "出行"],
    },
}

# 关键词里用到的字的繁体 → 简体（只换这些字，不需要额外装软件）
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

# 住客类型统一成 5 组，平台原文保留在明细里
TRAVELLER_GROUPS = [
    ("商务", ["business", "商务", "商務", "出差"]),
    ("情侣", ["couple", "情侣", "情侶"]),
    ("独自", ["solo", "独自", "獨自"]),
    ("家庭", ["famil", "家庭", "亲子", "親子"]),
    ("朋友/团体", ["friend", "group", "好友", "朋友", "团体", "團體"]),
]

# 英文关键词按整词匹配（old 不会命中 told，hair 不会命中 chair），允许常见词尾（stains、smelly、leaking）；
# 下面这些是词干，后面可以接任意字母（renovated、refurbished、facilities）
EN_STEMS = {"renovat", "refurb", "facilit"}
EN_PATTERNS = {
    theme: re.compile("|".join(
        rf"\b{re.escape(w)}" + ("" if w in EN_STEMS else r"(?:s|es|ed|ing|y|er)?\b") for w in words["en"]))
    for theme, words in THEMES.items()
}

# 各平台的分项归成同一类再合并（按顺序匹配，名称里含这些词就归到这一类）；归不进去的（如携程"环境"）不合并
SUBSCORE_GROUPS = [
    ("卫生", ["卫生", "cleanliness", "清洁"]),
    ("服务", ["服务", "staff", "service"]),
    ("设施", ["设施", "facilities", "amenities", "property condition"]),
    ("舒适度", ["comfort", "舒适"]),
    ("位置", ["location", "位置"]),
    ("性价比", ["value", "性价比"]),
]

# 本身不带好坏的英文关键词：要和不满的说法在同一句里才算差评主题
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
    return "其他"


def is_negative(score10):
    return None if pd.isna(score10) else bool(score10 < NEGATIVE_BELOW)


# 差评主题只算"在抱怨"的那句话：一条差评里也常有夸奖（"早餐很丰富""服务态度蛮好"），
# 所以把点评按标点切成短句，主题词所在的短句里有不满的说法才算。
# 关键词本身就是不满（脏、很吵、太贵、发霉）的，直接算。
NEG_CUES_ZH = ["不", "没", "无", "差", "慢", "坏", "脏", "吵", "嘈", "闹", "贵", "问题", "失望", "离谱", "糟", "难",
               "霉", "臭", "味", "烂", "旧", "破", "漏", "堵", "乱", "久", "投诉", "拒绝", "恶心", "腥", "灰", "虫",
               "烦", "敷衍", "冷漠", "一般", "遗憾", "可惜", "缺", "少", "押金", "按金", "窄", "挤", "响",
               # 好评里常用建议的口气提不满："早餐再丰富些就好了""温控还有进步空间""还是有些前客人的头发"
               "进步空间", "有待", "就好了", "更好了", "再丰富", "细致些", "希望", "建议增加", "建议改", "建议加",
               "有些前客人",
               # 订房信息和实际不符："到店才说只含一人早餐""和图片不一样"
               "才说", "才告知", "才知道", "不符", "不一样", "只含", "加收", "另收", "另外收", "额外收"]
# 先去掉含上面字但其实是夸奖的说法
POSITIVE_ZH = ["不错", "不愧", "无可挑剔", "没得说", "没话说", "没的说", "无微不至", "毫不犹豫", "不虚此行", "没问题",
               "没有问题", "不用担心", "不少", "少不了", "必不可少", "一尘不染", "不远", "没有异味", "无异味", "不吵",
               "不贵", "不脏", "不旧", "不慢", "不嘈", "无可", "不辞", "不厌", "没一会", "没多久", "无论", "不得不",
               "差不多", "诸多问题", "各种问题", "有问必答", "解决问题", "问题都能", "不管", "响应", "不用", "比一般",
               "美味", "口味", "趣味", "没有之一", "堵车", "不过", "来不及", "不止", "不差"]
NEG_CUES_EN = re.compile(r"\b(not|no|never|poor|bad|dirty|broken|rude|slow|noisy|small|old|smell\w*|expensive|"
                         r"but|however|wasn'?t|didn'?t|couldn'?t|isn'?t|doesn'?t|problem\w*|issue\w*|disappoint\w*|"
                         r"terrible|horrible|awful|worst|lack\w*|missing|unclean|stain\w*|leak\w*|cold|overpriced)\b")
# "没有…毛发/异味"这类是在说"没有问题"，抱怨判断前先去掉
ABSENCE_ZH = re.compile(r"(没有|无|没)[^，。！？,.!?]{0,6}?(毛发|头发|异味|怪味|味道|灰尘|虫|蟑螂|噪音|霉|吵)")
# "虽然没有××那样大"是在比较，不是抱怨
COMPARE_ZH = re.compile(r"虽然?没有?[^，。！？,.!?]{0,20}?(那样|那么|这么)|不用担心[^，。！？,.!?]*|没有那种[^，。！？,.!?]*")
# 某个主题里不算抱怨的说法："有点小困扰"不是房间小，"大堂不大"说的是大堂，"孩子衣服搞脏了"不是卫生问题
THEME_NOT = {
    "房间小": re.compile(r"小困扰|小问题|小毛病|水压[^，。！？,.!?]{0,4}小|大堂[^，。！？,.!?]{0,3}(不大|小)"),
    "卫生/异味": re.compile(r"衣服[^，。！？,.!?]{0,4}脏"),
    # 周边商户贵不是酒店的价格问题："楼下便利店东西很贵"
    "价格/押金/性价比": re.compile(r"(便利店|超市|商场|商店|外卖|餐馆|饭店|市场|景区|门票|出租车|打车|停车场)[^，。！？,.!?]{0,8}(贵|不便宜)"),
}
# 维修主题只算"坏了、不好用"；"没有卫洗丽"是没配这个设施，不是故障
LACK_ZH = re.compile(r"(没有|没)[^，。！？,.!?]{0,3}(卫洗丽|智能马桶|智能盖板|加热便座|加湿器|浴缸)")
CLAUSE = re.compile(r"[。！？；，、：:,.!?;\n|]+|\s+but\s+|\s+however\s+")


CJK = re.compile(r"[\u4e00-\u9fff]")
CUE_WINDOW = 6  # 不满的说法要在主题词前后 6 个字以内


def near_cue(clause, word):
    """主题词在这句里出现，且前后 6 个字以内有不满的说法。"""
    start = clause.find(word)
    while start >= 0:
        around = clause[max(0, start - CUE_WINDOW): start + len(word) + CUE_WINDOW]
        if any(n in around for n in NEG_CUES_ZH):
            return True
        start = clause.find(word, start + 1)
    return False


def is_complaint_zh(word):
    return any(c in word for c in NEG_CUES_ZH + ["小", "偏", "远"])


try:  # 有 opencc 就先把繁体整体转成简体（Trip.com 英文页的中文点评是繁体），没有就只靠上面的字表
    from opencc import OpenCC
    _T2S = OpenCC("t2s").convert
except Exception:  # noqa: BLE001
    _T2S = None


def normalize(text):
    t = str(text or "").lower()
    if _T2S:
        t = _T2S(t)
    return t.translate(TRAD).replace("坏境", "环境")


# 好评在夸什么：方面词前后 10 个字内有夸奖的说法，而且这一句里没有不满的说法才算。
# 本身就是夸奖的词（干净、安静、实惠）和公寓特有的设施（洗衣机、厨房），出现就算。
# 客人在点评里提到的具体东西（设施、位置、出行目的）。用来看每类客人最在意什么，酒店介绍该写什么。
# 只数"提到"（差评里提到的不算，分母是全部点评），不判断夸还是骂；分成"基本功"（每家酒店都该做到）和"卖点"（不是每家都有）。
# summary.json 里可以用 "feature_keywords": {"附近商场": ["万达", "购物中心"]} 补这家酒店周边的地名，会并进同名项或新增一项。
FEATURES = {
    # 卖点
    "厨房、厨具": ("卖点", r"厨房|厨具|灶|锅碗|餐具|做饭|煮饭|kitchen|cook|stove"),
    "洗衣机、烘干": ("卖点", r"洗衣|烘干|洗烘|washing|laundry|dryer|washer"),
    "冰箱、微波炉": ("卖点", r"冰箱|微波炉|fridge|refrigerator|microwave"),
    "房间大、套房": ("卖点", r"房间.{0,3}大|宽敞|套房|两室|两居|一室一厅|客厅|spacious|suite|living room"),
    "早餐": ("卖点", r"早餐|早饭|breakfast"),
    "健身房、桑拿": ("卖点", r"健身|桑拿|gym|fitness|sauna"),
    "休闲区、台球、游戏": ("卖点", r"台球|桌上足球|游戏|休闲廊|乒乓|billiard|pool table|game room"),
    "免费咖啡": ("卖点", r"咖啡|coffee"),
    "儿童、亲子": ("卖点", r"孩子|小孩|小朋友|儿童|宝宝|闺女|儿子|女儿|挖沙|\bkids?\b|child|toddler|baby"),
    "欢迎水果": ("卖点", r"水果|fruit"),
    "海边、沙滩": ("卖点", r"海边|沙滩|海滩|栈桥|看海|海景|beach|seaside|sea view"),
    "商场、购物": ("卖点", r"商场|超市|购物|\bmall\b|supermarket|shopping"),
    "交通、机场、车站": ("卖点", r"机场|高铁|火车站|地铁|公交|打车|出租车|滴滴|airport|station|metro|\bcab\b|taxi|didi"),
    "会外语的员工": ("卖点", r"会说韩语|会韩语|韩语|会说英语|会英语|会日语|english.speaking|speaks? english"),
    "停车": ("卖点", r"停车|parking"),
    "安静、隔音": ("卖点", r"安静|隔音|\bquiet"),
    "床、枕头": ("卖点", r"床垫|床品|枕头|床很舒服|床也很舒服|mattress|pillow|comfortable bed"),
    "出差、公司": ("卖点", r"出差|公司|工厂|开会|business trip|for work|work trip"),
    "性价比": ("卖点", r"性价比|实惠|划算|value for money|good value"),
    # 基本功
    "服务、前台": ("基本功", r"服务|前台|员工|工作人员|管家|staff|service|reception"),
    "干净、卫生": ("基本功", r"干净|卫生|整洁|一尘不染|clean"),
}


def feature_patterns(summary):
    pats = {k: v for k, v in FEATURES.items()}
    for name, words in (summary.get("feature_keywords") or {}).items():
        extra = "|".join(re.escape(w) for w in words)
        if name in pats:
            kind, pat = pats[name]
            pats[name] = (kind, pat + "|" + extra)
        else:
            pats[name] = ("卖点", extra)
    return pats


LANG_NAMES = {"zh": "中文", "ko": "韩文", "ja": "日文", "en": "英文"}


def language_mentions(reviews, summary):
    """外文点评（按点评语言，语言不等于国籍）里提到什么：每种外语点评里提到每样东西的条数，和中文点评对比。
    外语点评通常很少，所以给条数，不只给比例。韩文、日文要先在 translations.json 里译成中文才数得到。"""
    # 分母是这种语言的全部点评（和报告里其他比例一样）；差评里提到的不算（多半是在抱怨）
    ok = reviews.copy()
    ok["_t"] = ok["分析用文本"].fillna("").astype(str).map(normalize).str.lower() + " " + ok["title"].fillna("").astype(str).str.lower()
    ok.loc[ok["差评"] == True, "_t"] = ""  # noqa: E712
    langs = [l for l, c in ok["review_language"].value_counts().items() if l != "未知"]
    rows = []
    for name, (kind, pat) in feature_patterns(summary).items():
        hit = ok["_t"].str.contains(pat, regex=True)
        row = {"类别": kind, "提到的东西": name}
        for l in langs:
            sub = ok["review_language"] == l
            row[f"{LANG_NAMES.get(l, l)}（{int(sub.sum())} 条）条数"] = int(hit[sub].sum())
            row[f"{LANG_NAMES.get(l, l)}（{int(sub.sum())} 条）占比"] = hit[sub].mean()
        rows.append(row)
    table = pd.DataFrame(rows)
    table.attrs["langs"] = [(LANG_NAMES.get(l, l), int((ok["review_language"] == l).sum())) for l in langs]
    return table


def room_type_table(reviews):
    """各房型：点评条数、平均分、有抱怨的比例、被抱怨最多的问题。只列 10 条以上的房型。"""
    rows = []
    d = reviews[reviews["room_type"].fillna("").astype(str).str.strip() != ""]
    for rt, part in d.groupby("room_type"):
        if len(part) < MIN_GROUP_REVIEWS:
            continue
        comp = part["命中主题"].fillna("") != ""
        top = part.loc[comp, "命中主题"].str.split("；").explode().value_counts()
        rows.append({"房型": rt, "点评条数": len(part), "平均分（10分制）": round(part["score10"].mean(), 1),
                     "有抱怨的比例": comp.mean(), "被抱怨最多的问题": "、".join(f"{k} {v} 条" for k, v in top.head(3).items()),
                     "住这个房型最多的客人": "、".join(f"{g} {n} 条" for g, n in part["客群"].replace("", "没写").fillna("没写").value_counts().head(2).items())})
    return pd.DataFrame(rows).sort_values("点评条数", ascending=False) if rows else pd.DataFrame(columns=["房型"])


def language_summary(reviews):
    """各语言点评：条数、平均分、来自哪些平台、住客类型。"""
    rows = []
    for l, d in reviews.groupby("review_language"):
        rows.append({"语言": LANG_NAMES.get(l, l), "点评条数": len(d), "平均分（10分制）": round(d["score10"].mean(), 1),
                     "平台": "、".join(f"{p} {n} 条" for p, n in d["platform"].value_counts().items()),
                     "住客类型": "、".join(f"{g} {n} 条" for g, n in d["客群"].replace("", "没写").fillna("没写").value_counts().items())})
    return pd.DataFrame(rows).sort_values("点评条数", ascending=False)


def segment_mentions(reviews, summary):
    """各类客人在点评里提到什么：每类客人有多少条点评提到每样东西。分母是这类客人的全部点评（和"好评在夸什么"同一个分母），
    差评里提到的不算（多半是在抱怨）。只列 10 条以上的客群。"""
    ok = reviews.copy()
    ok["_t"] = ok["分析用文本"].fillna("").astype(str).map(normalize).str.lower() + " " + ok["title"].fillna("").astype(str).str.lower()
    ok.loc[ok["差评"] == True, "_t"] = ""  # noqa: E712
    typed = ok[ok["客群"].fillna("") != ""]
    groups = [g for g, c in typed["客群"].value_counts().items() if c >= MIN_GROUP_REVIEWS and g != "其他"]
    rows = []
    for name, (kind, pat) in feature_patterns(summary).items():
        hit = ok["_t"].str.contains(pat, regex=True)
        row = {"类别": kind, "提到的东西": name, "全部点评提到条数": int(hit.sum()), "全部点评占比": hit.mean() if len(ok) else None}
        th = typed["_t"].str.contains(pat, regex=True)
        for g in groups:
            sub = typed["客群"] == g
            row[f"{g}（{int(sub.sum())} 条）条数"] = int(th[sub].sum())
            row[f"{g}（{int(sub.sum())} 条）占比"] = th[sub].mean()
        rows.append(row)
    table = pd.DataFrame(rows).sort_values(["类别", "全部点评提到条数"], ascending=[True, False])
    table.attrs["groups"] = [(g, int((typed["客群"] == g).sum())) for g in groups]
    table.attrs["n"] = len(ok)
    return table


PRAISE = {
    "员工/服务": {"zh": ["服务", "前台", "员工", "工作人员", "服务员", "管家", "保安", "经理", "接待"],
               "en": ["staff", "service", "reception", "team"]},
    "卫生/干净": {"zh": ["干净", "整洁", "一尘不染", "卫生"], "en": ["clean"]},
    "早餐/餐饮": {"zh": ["早餐", "餐厅", "早饭"], "en": ["breakfast"]},
    "位置/交通": {"zh": ["位置", "交通", "海边", "沙滩", "地段", "周边", "商场", "出行", "地理"], "en": ["location", "beach", "mall"]},
    "房间（新、舒适、设施全）": {"zh": ["房间", "房型", "床", "空间", "装修", "设施"], "en": ["room", "bed", "spacious"]},
    "房内设施（厨房、洗衣机）": {"zh": ["洗衣机", "厨房", "冰箱", "微波炉", "烘干", "做饭", "厨具", "电磁炉"],
                         "en": ["kitchen", "laundry", "washing", "washer", "dryer"]},
    "健身房、休闲区和免费咖啡": {"zh": ["健身房", "桌球", "高尔夫", "休闲", "咖啡", "泳池"], "en": ["gym", "lounge", "billiard", "coffee", "pool"]},
    "安静": {"zh": ["安静", "隔音"], "en": ["quiet"]},
    "性价比": {"zh": ["性价比", "价格", "实惠", "划算", "物超所值"], "en": ["value", "price"]},
}
PRAISE_SELF = {"干净", "整洁", "一尘不染", "安静", "实惠", "划算", "物超所值", "洗衣机", "厨房", "做饭", "烘干",
               "健身房", "桌球", "高尔夫", "clean", "quiet", "kitchen", "laundry", "gym"}
POS_ZH = ["好", "棒", "赞", "不错", "满意", "热情", "热心", "周到", "贴心", "耐心", "专业", "干净", "整洁", "丰富", "方便",
          "便利", "舒适", "舒服", "新", "大", "宽敞", "安静", "齐全", "推荐", "喜欢", "温馨", "惊喜", "完美", "优秀", "高",
          "近", "nice", "实惠", "划算", "值", "亲切", "友好", "礼貌", "到位", "给力", "好吃", "美味", "丰盛", "多样",
          "足够", "还可以", "还行", "过得去", "媲美", "可口", "精致", "没话说", "没得说", "无可挑剔", "一流", "超出预期", "靠谱",
          "高效", "温柔", "细致", "用心"]
# 夸奖词表里的"大"会被"大堂""大厅"误触发，判断前先遮掉这些词；夸房间时"干净"算到卫生，不算房间
NOT_PRAISE = ["大堂", "大厅", "大型", "大家", "大概", "大约", "大叔", "大姐", "大哥", "大门", "大海", "大床"]
THEME_SKIP_POS = {"房间（新、舒适、设施全）": {"干净", "整洁"}}
# 韩文、日文点评目前认不出在夸什么，不算进好评分析的分母
UNREADABLE = re.compile(r"[\uac00-\ud7af\u3040-\u30ff]")
PRAISE_WINDOW = 10  # 夸奖的说法在方面词前后 10 个字内
POS_EN = re.compile(r"\b(good|great|nice|excellent|friendly|helpful|clean|comfortable|convenient|quiet|spacious|delicious|"
                    r"recommend\w*|love\w*|perfect|amazing|kind|attentive|well|wonderful|best)\b")


PRAISE_CLAUSE = re.compile(r"[。！？；，、,.!?;\n|]+|\s+but\s+|\s+however\s+")


def clauses_of(text, splitter=None):
    """按标点切成短句；中文再按空格切。夸奖判断不按冒号切（"服务：热情"要连在一起看）。"""
    out = []
    for c in (splitter or CLAUSE).split(ABSENCE_ZH.sub("", text)):
        out += c.split() if CJK.search(c) else [c]
    return [c for c in out if c.strip()]


def praise_hits(text):
    """一条点评在夸哪些方面（可以多个）。"""
    hits = []
    text = normalize(text)
    if UNREADABLE.search(text):
        return []
    clauses = clauses_of(text, PRAISE_CLAUSE)
    for theme, words in PRAISE.items():
        skip = THEME_SKIP_POS.get(theme, set())
        pos = [p for p in POS_ZH if p not in skip]
        for c in clauses:
            if CJK.search(c):  # 中文句子里夹的员工英文名会占掉窗口，先去掉
                c = re.sub(r"[a-z][a-z&/ ]*", "", c)
            for w in NOT_PRAISE:
                c = c.replace(w, "＿" * len(w))
            masked = COMPARE_ZH.sub("", c)
            if theme.startswith("维修"):
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
    """一条点评提到了哪些主题（可以多个，每个主题只记一次）。
    complaints_only=True 时只算在抱怨的短句（用于差评主题）。"""
    lower = str(text or "").lower()
    simple = normalize(text)
    if not complaints_only:
        return [theme for theme, words in THEMES.items()
                if EN_PATTERNS[theme].search(lower) or any(w in simple for w in words["zh"])]
    hits = []
    simple = simple.replace("坏境", "环境")  # 常见错别字，"坏"会被当成维修
    clauses = []
    for c in CLAUSE.split(ABSENCE_ZH.sub("", simple)):
        # 中文点评常用空格分句；英文不能按空格切
        clauses += c.split() if CJK.search(c) else [c]
    clauses = [c for c in clauses if c.strip()]
    for theme, words in THEMES.items():
        for c in clauses:
            masked = COMPARE_ZH.sub("", c)
            if theme.startswith("维修"):
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
# 读数据
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
        r["来源文件"] = path.stem  # 一个文件 = 一个平台的一种语言，时间对比按文件检查覆盖范围
    return rows


def column(df, name, default=None):
    return df[name] if name in df.columns else pd.Series([default] * len(df), index=df.index)


def apply_translations(df, folder):
    """数据文件夹里的 translations.json：关键词识别不了的语言（韩文、日文等），由人把原文译成中文：
    [{"platform": "Trip.com", "date": "2024-03-07", "text_starts": "호텔도 좋고", "zh": "酒店好，员工也亲切……"}]
    识别抱怨、夸奖、提到什么时用译文；报告和 Excel 里原文、译文都保留。"""
    df["译文"] = ""
    path = Path(folder) / "translations.json"
    if path.exists() and not df.empty:
        for c in json.loads(path.read_text(encoding="utf-8")):
            hit = (df["platform"] == c["platform"]) & (df["date"].dt.strftime("%Y-%m-%d") == c["date"]) \
                & df["text"].str.startswith(c["text_starts"])
            if not hit.any():
                print(f"translations.json 里这条没对上任何点评：{c['platform']} {c['date']} {c['text_starts']}")
            df.loc[hit, "译文"] = c["zh"]
    df["分析用文本"] = [z if z else t for z, t in zip(df["译文"], df["text"])]
    return df


def apply_corrections(df, folder):
    """人工核对后的更正（数据文件夹里的 corrections.json），自动识别错了的抱怨主题按这里改：
    [{"platform": "飞猪", "date": "2025-05-06", "text_starts": "差评给餐厅男经理", "themes": ["员工/服务"], "why": "骂的是餐厅员工"}]
    themes 是这条点评正确的抱怨主题（空列表 = 不是抱怨）。"""
    path = Path(folder) / "corrections.json"
    if not path.exists() or df.empty:
        return df
    for c in json.loads(path.read_text(encoding="utf-8")):
        hit = (df["platform"] == c["platform"]) & (df["date"].dt.strftime("%Y-%m-%d") == c["date"]) \
            & df["text"].map(normalize).str.startswith(normalize(c["text_starts"]))
        if not hit.any():
            print(f"corrections.json 里这条没对上任何点评：{c}")
        df.loc[hit, "命中主题"] = "；".join(c.get("themes", []))
    return df


def load_reviews(folder, pattern):
    rows = []
    for path in sorted(folder.glob(pattern)):
        rows += read_rows(path)
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    for name in ["platform", "hotel", "title", "text", "text_disliked", "date", "traveler_type", "room_type",
                 "reviewer_country", "review_language", "来源文件"]:
        if name not in df.columns:
            df[name] = None
    df["text"] = df["text"].fillna("").astype(str).str.split().str.join(" ")
    # zh-CN、zh-HK、zh_TW、en-US 这类都归成 zh、en；繁体中文也是中文
    df["review_language"] = (df["review_language"].fillna("未知").astype(str).str.strip().str.lower()
                             .str.split(r"[-_]").str[0].replace({"": "未知", "cn": "zh"}))
    scale = pd.to_numeric(column(df, "score_scale", 10), errors="coerce").fillna(10)
    df["score10"] = pd.to_numeric(column(df, "score"), errors="coerce") * 10 / scale
    df["date"] = pd.to_datetime(df["date"], errors="coerce", format="mixed")
    df = df[df["score10"].notna() | (df["text"] != "")]
    df["差评"] = df["score10"].map(is_negative)
    df["客群"] = df["traveler_type"].map(traveller_group)
    # 差评主题只看客人说"不满意"的部分：Booking 有单独的差评段，就只用差评段（好评段里的夸奖不算差评主题）
    # 有差评段字段但是空的（客人只写了好评段），这条不计差评主题，免得把夸奖算进去
    has_split = df["text_disliked"].notna()
    disliked = df["text_disliked"].fillna("").astype(str).str.strip()
    # 韩文、日文等关键词识别不了的点评，用 translations.json 里的中文译文来识别（原文保留）
    df = apply_translations(df, folder)
    df["命中主题"] = [";".join(theme_hits(d if split else t, complaints_only=True))
                   for split, d, t in zip(has_split, disliked, df["分析用文本"])]
    df["命中主题"] = df["命中主题"].str.replace(";", "；")
    df = apply_corrections(df, folder)
    # 好评在夸什么：只看不是差评的点评
    df["夸的方面"] = ["；".join(praise_hits(t)) if neg is not True else "" for t, neg in zip(df["分析用文本"], df["差评"])]
    # 去重：同一平台、同一正文、同一日期、同一分数、同一国家和房型才算重复，避免把不同客人的同样短评删掉
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
        sys.exit(f"summary.json 格式不对（{e}），请按 SKILL.md 第 4 步的格式改好再运行。")
    if not isinstance(summary, dict) or not isinstance(summary.get("platforms") or {}, dict):
        sys.exit("summary.json 格式不对：最外层要是 {\"hotel\": ..., \"platforms\": {...}}，请按 SKILL.md 第 4 步改好再运行。")
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
            if "提及次数" in t:
                t["提及次数"] = to_number(t["提及次数"])
            for key in ["正面占比", "负面占比"]:
                if key in t:
                    v = to_number(t[key])
                    t[key] = v / 100 if v is not None and (str(t[key]).strip().endswith("%") or v > 1) else v
        f.pop("excluded", None)
    return summary


# ---------------------------------------------------------------------------
# 合并用的工具
# ---------------------------------------------------------------------------

def wavg(values):
    """加权平均：values = [(数值, 权重)]，跳过空值和权重为 0 的。"""
    pairs = [(v, w) for v, w in values if v is not None and not pd.isna(v) and w and not pd.isna(w)]
    total = sum(w for _, w in pairs)
    return sum(v * w for v, w in pairs) / total if total else None


def platform_negative(p, reviews, pf):
    """(差评占比, 说明)。平台页面公布了差评条数（summary 的 negative_reviews，如携程的"差评 (7)"）就用公布的，
    否则用抓到的点评里的比例。"""
    neg, pub = pf.field(p, "negative_reviews"), pf.published(p)
    if neg is not None and pub:
        return neg / pub, f"{pf.disp(p)} 页面公布差评 {int(neg)}/{int(pub)} 条"
    df = reviews[reviews["platform"] == p]
    share = negative_share(df)
    if share is None:
        return None, f"{pf.disp(p)} 抓到的点评没有分数"
    n = int(df["差评"].dropna().astype(bool).sum())
    return share, f"{pf.disp(p)} 抓到的 {int(df['差评'].notna().sum())} 条里 {n} 条差评"


def negative_share(df):
    flagged = df["差评"].dropna()
    return flagged.astype(bool).mean() if len(flagged) else None


class Platforms:
    """各平台的评分和权重。权重 = 平台公布的点评数。
    - 携程和 Trip.com 是同一套点评：点评从 Trip.com 抓，评分、点评数、分项用携程的，只算一次。
    - 抓了点评但 summary 没填点评数的平台，权重用抓到的条数（会在"怎么算"里写明）。
    - 只读评分、又没填点评数的平台，不参加合并（会写明）。"""

    def __init__(self, summary, reviews):
        self.info = summary.get("platforms", {})
        self.scraped = sorted(reviews["platform"].dropna().unique())
        self.counts = reviews["platform"].value_counts().to_dict()

    def same(self, p):
        """携程和 Trip.com 当成同一个平台：都用携程的名字（携程没填评分时用 Trip.com 的）。"""
        if p in ("Trip.com", "携程") and "携程" in self.info and "Trip.com" in self.info:
            # Trip.com 页面是 10 分制、多一位精度，两边都有评分时用 Trip.com 的
            return "Trip.com" if self.info["Trip.com"].get("rating") is not None else "携程"
        return "携程" if p == "Trip.com" and "携程" in self.info else p

    def field(self, p, key):
        """先看代表平台，没有再看另一个（携程/Trip.com）。"""
        for name in [self.same(p), p, "携程" if p == "Trip.com" else "Trip.com" if p == "携程" else p]:
            v = (self.info.get(name) or {}).get(key)
            if v is not None:
                return v
        return None

    def published(self, p):
        return self.field(p, "total_reviews")

    def disp(self, p):
        # 携程和 Trip.com 是同一套点评，报告里只叫"携程"（国内用户不熟悉 Trip.com）
        return "携程" if p in ("Trip.com", "携程") and "携程" in self.info else p

    def weight(self, p):
        """(权重, 说明)；不能参加合并时权重为 None。"""
        pub = self.published(p)
        if pub:
            return pub, ""
        scraped = self.counts.get(p) or (self.counts.get("Trip.com") if p == "携程" else None)
        if scraped:
            return scraped, f"{p}未填公布点评数，按抓到的 {scraped} 条算"
        return None, f"{p}未填公布点评数，未计入"

    def rating10(self, p):
        for name in [self.same(p), p]:
            f = self.info.get(name) or {}
            if f.get("rating") is not None:
                return round(f["rating"] * 10 / (f.get("scale") or 10), 2)
        return None

    def rating_platforms(self):
        """参加综合评分的平台（携程和 Trip.com 只留一个）。"""
        names = [p for p in self.info if self.rating10(p) is not None]
        return list(dict.fromkeys(self.same(p) for p in names))

    def all_published(self):
        """所有填了点评数的平台相加（携程和 Trip.com 只算一次）。"""
        return sum(self.published(p) or 0 for p in dict.fromkeys(self.same(p) for p in self.info))

    def coverage(self, used):
        """这些平台占各平台公布点评总数的比例；有平台没填点评数时写明。"""
        total = self.all_published()
        names = list(dict.fromkeys(self.same(p) for p in used))
        part = sum(self.published(p) or 0 for p in names)
        unknown = [p for p in names if not self.published(p)]
        text = f"，占各平台公布点评数的 {part / total:.0%}" if total and part else ""
        return text + (f"（另有 {'、'.join(unknown)} 未填点评数，比例里不含）" if unknown else "")


# ---------------------------------------------------------------------------
# 表格
# ---------------------------------------------------------------------------

def theme_shares(df):
    neg = df[df["差评"] == True]  # noqa: E712
    row = {"差评条数": len(neg)}
    for theme in THEMES:
        row[theme] = neg["命中主题"].str.contains(re.escape(theme)).mean() if len(neg) else None
    return row


def top_themes(row, n=3):
    if row["差评条数"] < MIN_NEGATIVE_REVIEWS:
        return []
    shares = pd.Series({t: row[t] for t in THEMES}).astype(float)
    return list(shares[shares > 0].nlargest(n).items())


def theme_reason(row):
    if row.get("_没有分数"):
        return "抓到的点评没有分数，分不出差评"
    if row["差评条数"] < MIN_NEGATIVE_REVIEWS:
        return f"差评不足{MIN_NEGATIVE_REVIEWS}条，不下结论"
    return "差评没有命中任何主题，请看差评明细"


def theme_table(pool, reviews, pf):
    """每个平台一行，再加"全部平台"。全部平台的主题占比 = 各平台占比按"平台点评数 × 差评占比"（估计的差评条数）加权；
    只用正常抓取的点评（补抓的低分点评只当原文证据，不进任何比例）；差评不足 5 条的平台不参加合并。"""
    rows, used, left = [], [], []
    for p in pf.scraped:
        row = {"范围": p, **theme_shares(reviews[reviews["platform"] == p])}
        w, _ = pf.weight(p)
        share = platform_negative(p, reviews, pf)[0]
        row["_权重"] = (w or 0) * (share or 0)
        rows.append(row)
        if row["差评条数"] >= MIN_NEGATIVE_REVIEWS and row["_权重"]:
            used.append(p)
        elif share is None:
            left.append(f"{p}（抓到的点评没有分数）")
        elif row["差评条数"] >= MIN_NEGATIVE_REVIEWS:
            left.append(f"{p}（正常抓取的点评里没有差评，无法估计权重）")
        else:
            left.append(f"{p}（差评不足{MIN_NEGATIVE_REVIEWS}条）")
    all_row = {"范围": "全部平台", "差评条数": sum(r["差评条数"] for r in rows if r["范围"] in used),
               "_没有分数": all(negative_share(reviews[reviews["platform"] == p]) is None for p in pf.scraped)}
    for theme in THEMES:
        all_row[theme] = wavg([(r[theme], r["_权重"]) for r in rows if r["范围"] in used])
    no_score = all_row.pop("_没有分数")
    table = pd.DataFrame(rows + [all_row]).drop(columns="_权重")
    table["_没有分数"] = [negative_share(reviews[reviews["platform"] == r["范围"]]) is None for r in rows] + [no_score]
    table["差评最多的主题"] = ["；".join(f"{t} {s:.0%}" for t, s in top_themes(r)) or theme_reason(r)
                          for r in table.to_dict("records")]
    table = table.drop(columns="_没有分数")
    table["合并说明"] = [""] * len(rows) + [
        ("参加合并：" + "、".join(used) if used else "没有平台的差评达到5条") +
        ("；未参加：" + "、".join(left) if left else "")]
    return table


def platform_ratings(summary):
    rows = []
    for platform, f in summary.get("platforms", {}).items():
        if f.get("rating") is None and f.get("total_reviews") is None:
            continue  # 只填了住客类型的平台（如 Trip.com，评分同携程）不列
        scale = f.get("scale") or 10
        rating = f.get("rating")
        rows.append({
            "平台": platform,
            "平台评分": rating,
            "满分": scale,
            "换算成10分制": round(rating * 10 / scale, 2) if rating is not None else None,
            "平台公布点评数": f.get("total_reviews"),
            "公布数和能看到的条数之差（估算）": f.get("hidden_reviews"),
            "备注": ("携程和 Trip.com 是同一套点评，合并时只算一次；" if platform in ("携程", "Trip.com")
                   and {"携程", "Trip.com"} <= set(summary.get("platforms", {})) else "") + str(f.get("note", "")),
        })
    return pd.DataFrame(rows, columns=["平台", "平台评分", "满分", "换算成10分制", "平台公布点评数", "公布数和能看到的条数之差（估算）", "备注"])


def subscore_group(name, platform=None):
    text = str(name).lower()
    if platform == "Trip.com" and text == "location":
        return None  # Trip.com 英文页的 Location 就是携程的"环境"（周边环境），不是位置评分，不参加合并
    for group, words in SUBSCORE_GROUPS:
        if any(w in text for w in words):
            return group
    return None


def subscores(summary, pf):
    rows = []
    for platform, f in summary.get("platforms", {}).items():
        others = summary.get("platforms", {})
        if platform == "携程" and (others.get("Trip.com") or {}).get("subscores"):
            continue  # 和 Trip.com 是同一套点评，Trip.com 填了分项（10 分制、更细）就用 Trip.com 的
        if platform == "Trip.com" and not f.get("subscores"):
            continue
        scale = f.get("scale") or 10
        w, note = pf.weight(platform)
        for name, value in (f.get("subscores") or {}).items():
            rows.append({"平台": platform, "分项": name, "分数": value, "满分": scale,
                         "换算成10分制": round(value * 10 / scale, 2), "统一分项": subscore_group(name, platform),
                         "权重_平台点评数": w, "备注": note})
    return pd.DataFrame(rows, columns=["平台", "分项", "分数", "满分", "换算成10分制", "统一分项", "权重_平台点评数", "备注"])


def traveller_mix(summary, reviews):
    """平台公布的住客类型条数（点评区住客类型筛选项上的数字）。携程和 Trip.com 都填了时只用 Trip.com 的。
    平台不公布条数（如 Trip.com 的筛选项不带数字）、但点评已经全部抓下来的（all_reviews_scraped 为 true），
    就按抓到的点评里的住客类型统计。"""
    platforms = summary.get("platforms", {})
    rows = []
    has_counts = {p for p, f in platforms.items() if f.get("traveller_counts")}
    if "Trip.com" in has_counts or "携程" in has_counts:
        has_counts |= {"Trip.com", "携程"}
    for p in sorted(reviews["platform"].dropna().unique()):
        flag = (platforms.get(p) or {}).get("all_reviews_scraped")
        if p == "Trip.com" and flag is None:
            flag = (platforms.get("携程") or {}).get("all_reviews_scraped")
        if p in has_counts or flag is not True:
            continue
        counts = reviews.loc[reviews["platform"] == p, "traveler_type"].fillna("").astype(str).str.strip()
        counts = counts[counts != ""].value_counts()
        total = counts.sum()
        for label, n in counts.items():
            rows.append({"平台": f"{p}（按抓到的点评统计）", "住客类型原文": label, "客群": traveller_group(label),
                         "条数": int(n), "占比": n / total if total else None})
    for platform, f in platforms.items():
        if platform == "携程" and (platforms.get("Trip.com") or {}).get("traveller_counts"):
            continue
        counts = {k: v for k, v in (f.get("traveller_counts") or {}).items() if v}
        total = sum(counts.values())
        for label, n in counts.items():
            rows.append({"平台": platform, "住客类型原文": label, "客群": traveller_group(label),
                         "条数": n, "占比": n / total if total else None})
    return pd.DataFrame(rows, columns=["平台", "住客类型原文", "客群", "条数", "占比"])


def platform_topics(summary):
    rows = [{"平台": platform, **t} for platform, f in summary.get("platforms", {}).items() for t in f.get("topics") or []]
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["平台"])


def norm_lang(code):
    """语言代码统一成小写、只取主语言：zh-HK、zh_CN、ZH → zh；cn → zh。"""
    c = re.split(r"[-_]", str(code).strip().lower())[0]
    return {"": "未知", "cn": "zh"}.get(c, c)


def norm_langs(codes):
    return [norm_lang(x) for x in codes]


def complete_sample(reviews, summary):
    """看时间趋势用的样本：只用点评已经抓全的平台和语言（summary.json 里 all_reviews_scraped 为 true 或语言列表）。
    没抓全的平台，旧点评少只是没抓到，放进来会让"以前问题少""最近变差"看起来像真的。返回 (样本, 用到的平台说明)。"""
    flags = {p: f.get("all_reviews_scraped") for p, f in summary.get("platforms", {}).items()}
    keep = pd.Series(False, index=reviews.index)
    used = []
    for p in reviews["platform"].dropna().unique():
        flag = flags.get(p)
        if p == "Trip.com" and flag is None:
            flag = flags.get("携程")
        m = reviews["platform"] == p
        if flag is True:
            keep |= m
            used.append(p)
        elif isinstance(flag, list) and flag:
            keep |= m & reviews["review_language"].isin(norm_langs(flag))
            used.append(p)
    return reviews[keep], used


def periods(reviews, today, complete):
    """近 90 天，和 91 天到 1 年前，用抓到的点评（有分数的）。
    按语言检查：只有抓到的点评已经覆盖到 1 年前的语言才参加，否则这种语言"更早"那一段少是没抓全，不是真的少；
    这种语言的点评本来就不到 1 年（列表翻到底了），summary.json 里标 all_reviews_scraped（true 或语言代码列表）。
    参加的语言两段都要有 10 条以上，这个平台才参加对比。"""
    cut_recent = today - pd.Timedelta(days=RECENT_DAYS)
    cut_year = today - pd.Timedelta(days=COMPARE_DAYS)
    out = {}
    for p in sorted(reviews["platform"].dropna().unique()):
        df = reviews[(reviews["platform"] == p) & reviews["score10"].notna()]
        flag = complete.get(p)
        if flag is None and p == "Trip.com":
            flag = complete.get("携程")
        langs = []
        for source, part in df.groupby("来源文件"):
            lang = str(part["review_language"].iloc[0])
            done = flag is True or (isinstance(flag, list) and lang in norm_langs(flag))
            if done or (part["date"].notna().any() and part["date"].min() <= cut_year):
                langs.append(lang if lang != "未知" else source)
        covered = [src for src, part in df.groupby("来源文件")
                   if (str(part["review_language"].iloc[0]) if str(part["review_language"].iloc[0]) != "未知" else src) in langs]
        use = df[df["来源文件"].isin(covered)]
        recent = use[use["date"] >= cut_recent]
        earlier = use[(use["date"] < cut_recent) & (use["date"] >= cut_year)]
        if df.empty:
            reason = "抓到的点评没有分数"
        elif df["date"].isna().all():
            reason = "抓到的点评没有日期"
        elif not langs:
            reason = "抓到的点评没有覆盖到1年前"
        elif len(recent) < MIN_PERIOD_REVIEWS or len(earlier) < MIN_PERIOD_REVIEWS:
            reason = f"近90天{len(recent)}条、更早{len(earlier)}条，有一段少于{MIN_PERIOD_REVIEWS}条"
        else:
            reason = ""
        out[p] = {"近期": recent, "更早": earlier, "语言": langs, "不参加的原因": reason}
    return out


def platform_table(reviews, pf, themes, per):
    """每个平台一行（抓了点评的和只读评分的都列），看综合结果是哪个平台拉高或拉低的。"""
    t = themes.set_index("范围")
    names = list(pf.scraped) + [p for p in pf.rating_platforms() if p not in pf.scraped
                                and not (p == "携程" and "Trip.com" in pf.scraped)]
    rows = []
    for p in names:
        df = reviews[reviews["platform"] == p]
        w, note = pf.weight(p)
        joined = p in per and not per[p]["不参加的原因"]
        rows.append({
            "平台": "携程（点评从 Trip.com 抓）" if p == "Trip.com" and "携程" in pf.info else p,
            "平台评分_10分制": pf.rating10(p),
            "平台公布点评数": pf.published(p),
            "抓到条数": len(df) if len(df) else "只读评分，没抓点评",
            "抓到的最新点评": df["date"].max() if len(df) else None,
            "抓到的最早点评": df["date"].min() if len(df) else None,
            "抓到的点评平均分_10分制": round(df["score10"].mean(), 2) if df["score10"].notna().any() else None,
            "差评占比": platform_negative(p, reviews, pf)[0] if len(df) else None,
            "差评占比怎么算": platform_negative(p, reviews, pf)[1] if len(df) else "",
            f"近{RECENT_DAYS}天平均分": round(per[p]["近期"]["score10"].mean(), 2) if joined else None,
            "91天到1年前平均分": round(per[p]["更早"]["score10"].mean(), 2) if joined else None,
            "参加时间对比": ("是（" + "、".join(per[p]["语言"]) + "）") if joined else
                        ("否：" + per[p]["不参加的原因"] if p in per else "否：没抓点评"),
            "差评条数": t.loc[p, "差评条数"] if p in t.index else None,
            "差评最多的主题": t.loc[p, "差评最多的主题"] if p in t.index else "",
            "备注": note,
        })
    return pd.DataFrame(rows)


def overall_table(reviews, pool, pf, ratings, subs, travellers, themes, impact, per, today, summary):
    """一页汇总：所有平台合在一起的结果。每行写明用了哪些平台、多少条。"""
    rows = []
    add = lambda item, value, how: rows.append({"项目": item, "合并结果": value, "用了哪些平台、怎么算": how})
    notes = lambda ps: "；".join(n for n in (pf.weight(p)[1] for p in ps) if n)
    info = summary.get("platforms", {})

    # 0. 平台范围
    included = list(dict.fromkeys([pf.disp(p) for p in pf.scraped] + [pf.disp(p) for p in pf.rating_platforms()]))
    scope = "已包含：" + ("、".join(included) or "无")
    if summary.get("not_listed"):
        scope += "；本次没有查到，待核实：" + "、".join(summary["not_listed"])
    if summary.get("not_included"):
        scope += "；本次未包含：" + "、".join(summary["not_included"])
    add("平台范围", scope, "本次没有查到 = 在该平台搜不到这家酒店，待酒店核实是否上架；本次未包含 = 本版 Skill 还不能抓的平台")

    # 1. 评分和点评数
    rated = pf.rating_platforms()
    used = [p for p in rated if pf.weight(p)[0]]
    overall = wavg([(pf.rating10(p), pf.weight(p)[0]) for p in used])
    if not summary:
        how = "summary.json 未填，没有平台公布的评分"
    elif not rated:
        how = "summary.json 里没有填任何平台的评分"
    else:
        how = ("各平台公布的评分换算成10分制，按平台公布点评数加权：" + "、".join(
            f"{pf.disp(p)} {pf.rating10(p):g}（{int(pf.weight(p)[0])} 条）" for p in used)
               + (f"（{notes(rated)}）" if notes(rated) else ""))
    add("综合评分（10分制）", round(overall, 2) if overall is not None else None, how)
    add("平台公布点评总数", int(pf.all_published()) if pf.all_published() else None,
        ("、".join(pf.disp(p) for p in rated if pf.published(p)) + " 公布的点评数相加（携程和 Trip.com 只算一次）") if rated else "summary.json 未填")
    add("本次分析的点评条数", len(reviews), "抓到的写了文字的点评，去重后（不含补抓的低分点评）：" + "、".join(
        f"{pf.disp(p)} {pf.counts[p]} 条" for p in pf.scraped))

    # 2. 差评
    negs = {p: platform_negative(p, reviews, pf) for p in pf.scraped}
    usable = [p for p in pf.scraped if negs[p][0] is not None]
    neg = wavg([(negs[p][0], pf.weight(p)[0]) for p in usable])
    add("差评占本次分析的点评", (reviews["差评"] == True).sum() / len(reviews) if len(reviews) else None,  # noqa: E712
        f"正常抓到的 {len(reviews)} 条点评里低于 6 分的比例（报告里用这个；补抓的低分点评不算进比例）")
    add("差评占比（按平台公布的差评条数加权）", neg,
        ("按平台点评数加权：" + "；".join(negs[p][1] for p in usable)) if usable else "抓到的点评都没有分数")
    all_themes = themes.set_index("范围").loc["全部平台"]
    merged = [p for p in pf.scraped if p in str(all_themes["合并说明"]).split("；")[0]]
    neg_pool = reviews[(reviews["差评"] == True) & reviews["platform"].isin(merged)]  # noqa: E712
    year_ago = today - pd.Timedelta(days=COMPARE_DAYS)
    # 差评条数和时间：所有平台抓到的差评都算（差评主题的合并只用差评满 5 条的平台，那是另一回事）
    all_negs = reviews[reviews["差评"] == True]  # noqa: E712  比例和条数只用正常抓到的点评；补抓的低分点评只当原文证据
    recent_all = int((all_negs["date"] >= year_ago).sum())
    add("差评时间分布", f"共 {len(all_negs)} 条差评：近1年 {recent_all} 条，1年前 {len(all_negs) - recent_all} 条",
        "所有平台抓到的差评（低于6分，正常抓取的，不含补抓的低分点评）：" + "、".join(
            f"{pf.disp(p)} {n} 条" for p, n in all_negs["platform"].value_counts().items()) + "；近1年的条数少，说明这些问题不一定还存在")
    tops = top_themes(all_themes)
    if not tops:
        value = ("无法合并" if all_themes["差评条数"] == 0 and "未参加" in all_themes["合并说明"]
                 else all_themes["差评最多的主题"])
        add("差评最多的主题", value, all_themes["合并说明"])
    for i, (theme, share) in enumerate(tops, 1):
        k = int(neg_pool["命中主题"].str.contains(re.escape(theme)).sum())
        add(f"差评最多的主题 {i}", f"{theme}（{share:.0%}，{len(neg_pool)} 条差评里 {k} 条提到）",
            "差评里在抱怨这个主题的比例，各平台按估计的差评条数加权。" + all_themes["合并说明"])
    worst = impact.dropna(subset=["影响分"])
    worst = worst[worst["影响分"] <= -MIN_IMPACT].head(3)
    impact_platforms = [c[:-len("_抱怨条数")] for c in impact.columns if c.endswith("_抱怨条数")]
    impact_how = ("抱怨这个主题的点评平均分 − 其他点评平均分，在每个平台内算再合并，是相关不是因果；"
                  f"抱怨不到{MIN_THEME_MENTIONS}次的主题不算") if impact_platforms else "抓到的点评没有分数"
    if not impact_platforms:
        add("拉低分数最多的主题", "无法计算", impact_how)
    elif worst.empty:
        add("拉低分数最多的主题", f"没有主题被抱怨{MIN_THEME_MENTIONS}次以上、且让分数低{MIN_IMPACT:g}分以上", impact_how)
    for i, row in enumerate(worst.itertuples(), 1):
        add(f"拉低分数最多的主题 {i}", f"{row.主题}（抱怨它的 {row.抱怨的点评条数} 条点评平均低 {-row.影响分:.1f} 分）", impact_how)

    # 2b. 好评在夸什么
    pt = praise_table(reviews)
    n_good = pt.attrs.get("n", 0)
    for i, row in enumerate(pt[pt["夸它的点评条数"] > 0].head(3).itertuples(), 1):
        add(f"好评最常夸的 {i}", f"{row.方面}（{row.占比:.0%}，本次分析的 {n_good} 条点评里 {row.夸它的点评条数} 条）",
            "不是差评的中文、英文点评里，在夸这个方面的比例（方面词附近有夸奖的说法才算，自动识别，会漏掉一些）"
            + (f"；另有 {pt.attrs['unreadable']} 条韩文、日文点评没有识别" if pt.attrs.get("unreadable") else ""))

    # 3. 分项评分：各平台同类分项换算成 10 分制，按平台点评数加权
    grouped = subs.dropna(subset=["统一分项"]) if len(subs) else subs
    shown = False
    total_pub = pf.all_published()
    for group, _ in SUBSCORE_GROUPS:
        part = grouped[grouped["统一分项"] == group] if len(grouped) else grouped
        part = part[part["权重_平台点评数"].notna()] if len(part) else part
        if part.empty:
            continue
        shown = True
        value = wavg([(row.换算成10分制, row.权重_平台点评数) for row in part.itertuples()])
        n = int(part["权重_平台点评数"].sum())
        small = total_pub and n / total_pub < SMALL_SHARE
        add(f"分项：{group}（10分制）" + ("（样本少，仅供参考）" if small else ""), round(value, 2),
            "、".join(f"{pf.disp(row.平台)} {row.分项} {row.换算成10分制:g}（{int(row.权重_平台点评数)} 条）" for row in part.itertuples())
            + "，按平台点评数加权" + (f"，共 {n} 条，占公布点评数的 {n / total_pub:.0%}" if total_pub else ""))
    if not shown:
        add("分项评分", "无", "平台未公布或 summary.json 未填分项评分")

    # 4. 住客构成
    # 住客构成用抓到的点评上标的住客类型（一条点评算一类），和报告里其他客群数字同一个分母；
    # 平台筛选项上的条数（Booking.com 一条点评能算几类）只留在"住客类型"表里备查
    typed = reviews[reviews["客群"].fillna("").isin(["商务", "家庭", "情侣", "朋友/团体", "独自"])]
    if len(typed):
        counts = typed["客群"].value_counts()
        total = int(counts.sum())
        by_pf = "、".join(f"{pf.disp(p)} {n} 条" for p, n in typed["platform"].value_counts().items())
        for group, n in counts.items():
            add(f"住客：{group}", n / total, f"抓到的点评上标的住客类型，一条点评算一类，共 {total} 条（{by_pf}）；平台筛选项上的条数见'住客类型'表")
    else:
        add("住客构成", "无", "抓到的点评上都没有标住客类型")

    # 5. 近 90 天和 91 天到 1 年前：用抓到的点评，只用覆盖到 1 年前的平台和语言
    joined = [p for p in per if not per[p]["不参加的原因"]]
    left = [f"{pf.disp(p)}（{per[p]['不参加的原因']}）" for p in per if per[p]["不参加的原因"]]
    joined_desc = "、".join(pf.disp(p) + "（" + "、".join(per[p]["语言"]) + "）" for p in joined) or "无"
    how = (f"以 {today.date()} 为准，用抓到的点评算（和上面平台公布的综合评分算法不同，不要拿来和综合评分比）；"
           f"参加：{joined_desc}，按平台点评数加权"
           + (f"；没参加：{'、'.join(left)}" if left else ""))
    if not joined:
        add("近90天和91天到1年前对比", "无法对比", how)
    else:
        stats = {}
        for key in ["近期", "更早"]:
            n = sum(len(per[p][key]) for p in joined)
            k = sum(int(per[p][key]["差评"].dropna().astype(bool).sum()) for p in joined)
            stats[key] = (wavg([(per[p][key]["score10"].mean(), pf.weight(p)[0]) for p in joined]),
                          wavg([(negative_share(per[p][key]), pf.weight(p)[0]) for p in joined]), n, k)
        (ra, rn, rc, rk), (ea, en, ec, ek) = stats["近期"], stats["更早"]
        if min(rc, ec) < FEW_REVIEWS:
            how += f"；有一段不到 {FEW_REVIEWS} 条，一两条极端点评就能明显改变结果，看差评明细再下判断"
        add(f"近{RECENT_DAYS}天平均分（抓到的点评，10分制）", f"{ra:.2f}（{rc} 条）", how)
        add("91天到1年前平均分（抓到的点评，10分制）", f"{ea:.2f}（{ec} 条）", how)
        change = f"{ra - ea:+.2f} 分"
        if rc < FEW_REVIEWS:
            recent_all = pd.concat([per[p]["近期"] for p in joined])
            low = recent_all["score10"].min()
            rest = recent_all["score10"].drop(recent_all["score10"].idxmin())
            if len(rest):
                change += f"（近{RECENT_DAYS}天最低一条 {low:g} 分，去掉它后变化 {rest.mean() - ea:+.2f} 分）"
        add(f"平均分变化（近{RECENT_DAYS}天 − 更早）", change, how)
        add(f"近{RECENT_DAYS}天差评", f"{rn:.0%}（{rc} 条里 {rk} 条）", how)
        add("91天到1年前差评", f"{en:.0%}（{ec} 条里 {ek} 条）", how)
    return pd.DataFrame(rows)


MIN_GROUP_REVIEWS = 10  # 某类客人少于 10 条点评时，不单独列他们夸什么


def praise_table(reviews):
    """客人在夸什么：夸每个方面的条数，比例的分母是本次分析的全部点评（和报告里其他比例同一个分母）；
    差评不算夸（差评里的客气话不算），韩文、日文点评识别不了，也算在分母里，所以比例略偏低。再按客群分开看。"""
    unreadable = reviews["分析用文本"].astype(str).str.contains(UNREADABLE) & (reviews["差评"] != True)  # noqa: E712
    good = reviews
    n = len(good)
    rows = []
    groups = [g for g, c in good["客群"].value_counts().items() if c >= MIN_GROUP_REVIEWS and g != "其他"]
    for theme in PRAISE:
        hit = good["夸的方面"].str.contains(re.escape(theme))
        row = {"方面": theme, "夸它的点评条数": int(hit.sum()), "占比": hit.mean() if n else None}
        for g in groups:
            sub = good[good["客群"] == g]
            row[f"{g}（{len(sub)} 条）占比"] = sub["夸的方面"].str.contains(re.escape(theme)).mean()
        rows.append(row)
    table = pd.DataFrame(rows).sort_values("夸它的点评条数", ascending=False)
    table.attrs["n"] = n
    table.attrs["unreadable"] = int(unreadable.sum())
    return table


def theme_impact(reviews):
    """主题影响分：抱怨某主题的点评平均分，减去没抱怨它的点评平均分（10 分制）。
    只算在抱怨的短句（和差评主题同一套判断），夸奖不算，否则"服务很好"这类好评会让结果失真。
    先在每个平台内部算差值，再按各平台提到的条数加权平均，避免"这个主题多出现在低分平台"被算成拉低分数。
    只用正常抓取的点评，不用补抓的低分点评（会把分数拉偏）。这是相关，不是因果。"""
    scored = reviews[reviews["score10"].notna()]
    platforms = sorted(scored["platform"].dropna().unique())
    rows = []
    for theme in THEMES:
        row = {"主题": theme}
        total_hits, weighted, weight = 0, 0.0, 0
        for p in platforms:
            df = scored[scored["platform"] == p]
            hit = df["命中主题"].str.contains(re.escape(theme))
            n = int(hit.sum())
            total_hits += n
            diff = None
            if n and (~hit).any():
                diff = df.loc[hit, "score10"].mean() - df.loc[~hit, "score10"].mean()
                weighted += diff * n
                weight += n
            row[f"{p}_抱怨条数"] = n
            row[f"{p}_影响分"] = round(diff, 2) if diff is not None and n >= MIN_THEME_MENTIONS else None
        row["提到的点评条数"] = total_hits
        if total_hits < MIN_THEME_MENTIONS:
            row["影响分"], row["备注"] = None, f"抱怨少于{MIN_THEME_MENTIONS}次，不算影响分"
        elif not weight:
            row["影响分"], row["备注"] = None, "所有点评都抱怨这个主题，没有可对比的点评"
        else:
            row["影响分"], row["备注"] = round(weighted / weight, 2), ""
        rows.append(row)
    for row in rows:
        row["抱怨的点评条数"] = row.pop("提到的点评条数")
    table = pd.DataFrame(rows, columns=["主题", "影响分", "抱怨的点评条数", "备注"]) if not platforms else pd.DataFrame(rows)
    front = ["主题", "影响分", "抱怨的点评条数", "备注"]
    table = table[front + [c for c in table.columns if c not in front]]
    return table.sort_values("影响分", na_position="last")


SHEET_NOTES = {
    "综合汇总": "所有平台合在一起的结果：综合评分、点评总数、差评占比、差评最多的主题、拉低分数最多的主题、分项评分、住客构成、近90天和更早对比。每一项用了哪些平台、怎么算，写在第三列",
    "各平台对比": "每个平台一行（只读评分的平台也列），看综合结果是哪个平台拉高或拉低的",
    "各平台评分": "各平台页面上公布的总评分、点评数、公布数和能看到的条数之差（估算，包括只读评分、没抓点评的平台）",
    "分项评分": "各平台公布的分项评分，'统一分项'是合并时归到哪一类（空的不参加合并，如携程的'环境'、Trip.com 的 Location——它就是携程的'环境'）",
    "住客类型": "各平台点评区住客类型筛选项上的条数；平台不公布条数、点评又全部抓下来的（如 Trip.com），按抓到的点评统计，平台名后会注明",
    "差评主题": "每个平台和全部平台：差评条数、差评里提到每个主题的比例（Booking 只看差评段）",
    "好评在夸什么": "本次分析的全部点评里（差评不算夸），夸每个方面的条数和比例；后面几列是不同客群分开看（10 条以上才列）。没识别出夸奖不等于客人不满意",
    "各类客人提到什么": "分母是这类客人的全部点评（差评里提到的不算），每类客人提到每样东西（设施、位置、出行目的）的条数和比例；'基本功'是每家酒店都该做到的，'卖点'不是每家都有。用来看主要客群在意什么、酒店介绍该写什么",
    "各房型点评": "10 条以上点评的房型：条数、平均分、有抱怨的比例、被抱怨最多的问题、住这个房型最多的客人。看是不是某个房型反复被抱怨",
    "各语言点评": "按点评语言分：条数、平均分、来自哪些平台、住客类型。点评语言不等于国籍，只能当线索",
    "外文点评提到什么": "点评语言不等于国籍，只能当线索。分母是这种语言的全部点评（差评里提到的不算），每种语言的点评提到每样东西的条数和比例（韩文、日文用 translations.json 的中文译文来数）",
    "主题影响分": f"每个主题：抱怨它的点评平均分 − 其他点评平均分（夸奖不算），先在每个平台内算，再按抱怨条数加权。负数越大，抱怨它的点评分数越低（相关，不是因果）。抱怨少于{MIN_THEME_MENTIONS}次不算",
    "平台话题": "平台自己的话题统计（携程热门标签只有提及次数；Agoda 话题分析有正面、负面比例）",
    "差评明细": "全部差评和命中的主题（“来源”列标明哪些是补抓的低分点评，只当原文证据），按日期从新到旧",
    "全部点评": "抓到的全部点评（不含补抓的低分点评），按日期从新到旧",
}


def write_excel(tables, out):
    readme = pd.DataFrame([
        ["差评", f"本报告的规则：客人打分换算成10分制后低于{NEGATIVE_BELOW:g}分（5分制低于3分）。平台自己标的'差评'条数（如携程'差评 (7)'）另列，两者不一定相同"],
        ["合并方法", "各平台的数字按平台公布的点评数加权合并，点评多的平台占的比重大；携程和携程国际版（Trip.com）是同一套点评，只算一次。每一项用了哪些平台，写在'综合汇总'第三列"],
        ["抓到的点评", "只抓写了文字的点评（页面上只打分没写字的不抓）；各语言按抓到的条数算，没有按各语言的实际点评数加权"],
        ["差评占比", "报告里一律用'本次分析的点评里低于 6 分的比例'（综合汇总'差评占本次分析的点评'）。'各平台对比'和综合汇总里另列的'按平台公布的差评条数加权'是用平台自己标的差评条数（如携程'差评 (7)'）算的，只作备查和主题合并的权重"],
        ["差评主题", "差评里在抱怨该主题的比例：Booking 只看差评段；其他平台按标点和空格分句，主题词前后 6 个字内有不满的说法（不、没、差、坏、脏、吵、贵等）才算。自动识别，会有少量误判；一条差评可以提到多个主题。按'低分优先'补抓的点评只当原文证据（'差评明细'里'来源'列标明），不进任何比例和条数"],
        ["差评不足", f"某个范围的差评少于{MIN_NEGATIVE_REVIEWS}条时，不列主题结论，也不参加全部平台的主题合并"],
        ["公布数和能看到的条数之差", "估算值：可能是被隐藏、折叠，或只打分没写字，不一定全是平台隐藏的；内容未知，不在本分析内"],
    ] + [[name, SHEET_NOTES.get(name, "")] for name in tables], columns=["项目", "说明"])
    # 用户只认"携程"：Excel 里不出现 Trip.com 字样，抓取来源写成"携程国际版"
    def _cn(v):
        return v.replace("Trip.com", "携程国际版") if isinstance(v, str) else v
    readme = readme.map(_cn)
    tables = {name: t.rename(columns=_cn).map(_cn) for name, t in tables.items()}
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        readme.to_excel(writer, sheet_name="说明", index=False)
        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name[:31], index=False)
        for name, sheet in writer.sheets.items():
            for col in sheet.columns:
                width = max(len(str(c.value or "")) for c in col[:50])
                sheet.column_dimensions[col[0].column_letter].width = min(max(10, width * 1.6), 60)
                header = str(col[0].value or "")
                if "占比" in header or header in THEMES:
                    for c in col[1:]:
                        c.number_format = "0%"
            if name == "综合汇总":  # 占比、住客两类行显示成百分比
                for row in sheet.iter_rows(min_row=2):
                    if "占比" in str(row[0].value) or str(row[0].value).startswith("住客："):
                        row[1].number_format = "0%"
            sheet.freeze_panes = "B2"


def main(folder, today=None):
    folder = Path(folder)
    reviews = load_reviews(folder, "reviews_*.json")
    if reviews.empty:
        sys.exit("没有找到 reviews_*.json，先完成抓取。")
    low = load_reviews(folder, "lowscore_*.json")
    reviews["来源"] = "正常抓取"
    if len(low):
        low["来源"] = "补抓的低分点评（只当原文证据，不进任何比例）"
    pool = pd.concat([reviews, low], ignore_index=True) if len(low) else reviews
    pool = pool.drop_duplicates(subset=["platform", "text", "date", "score10", "reviewer_country", "room_type"])
    summary = load_summary(folder)
    today = pd.Timestamp(today or pd.Timestamp.today().normalize())

    names = reviews["hotel"].dropna()
    hotel = summary.get("hotel") or (names.iloc[0] if len(names) else "酒店")
    pf = Platforms(summary, reviews)
    ratings = platform_ratings(summary)
    subs = subscores(summary, pf)
    travellers = traveller_mix(summary, reviews)
    themes = theme_table(pool, reviews, pf)
    impact = theme_impact(reviews)
    complete = {p: f.get("all_reviews_scraped") for p, f in summary.get("platforms", {}).items()}
    per = periods(reviews, today, complete)
    # 不输出 reviewer_country：Trip.com 的客人没填国家时，这一栏显示的是昵称
    columns = ["platform", "date", "score10", "traveler_type", "客群", "room_type",
               "review_language", "title", "text", "译文", "命中主题", "夸的方面", "来源"]
    tables = {
        "综合汇总": overall_table(reviews, pool, pf, ratings, subs, travellers, themes, impact, per, today, summary),
        "各平台对比": platform_table(reviews, pf, themes, per),
        "各平台评分": ratings,
        "分项评分": subs,
        "住客类型": travellers,
        "差评主题": themes,
        "主题影响分": impact,
        "好评在夸什么": praise_table(reviews),
        "各类客人提到什么": segment_mentions(reviews, summary),
        "各语言点评": language_summary(reviews),
        "各房型点评": room_type_table(reviews),
        "外文点评提到什么": language_mentions(reviews, summary),
        "平台话题": platform_topics(summary),
        "差评明细": pool[pool["差评"] == True].reindex(columns=columns).sort_values("date", ascending=False),  # noqa: E712
        "全部点评": reviews.reindex(columns=columns).sort_values("date", ascending=False)
                      .rename(columns={"命中主题": "抱怨主题（自动识别，仅供参考）"}),
    }
    out = folder / f"点评分析_{BAD_FILENAME.sub('_', str(hotel))}.xlsx"
    write_excel(tables, out)
    pd.set_option("display.width", 200)
    pd.set_option("display.max_colwidth", 80)
    print(tables["综合汇总"][["项目", "合并结果"]].to_string(index=False))
    print(f"\n已保存：{out}")
    return {"tables": tables, "reviews": reviews, "pool": pool, "summary": summary, "today": today,
            "hotel": hotel, "excel": out, "platforms": pf}


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
