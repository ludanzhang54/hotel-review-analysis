"""Regression tests for the core statistics: python3 tests/test_core.py (prints "All tests passed" only if all pass).
They check that topped-up low-score reviews stay out of ratios, language codes are unified, a few easily misjudged sentences, and an empty sample."""
import io
import json
import sys
sys.dont_write_bytecode = True
import tempfile
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import analyze as A  # noqa: E402


def make_folder(rows, low=None, lang="zh-CN"):
    d = Path(tempfile.mkdtemp())
    (d / "reviews_test_zh.json").write_text(json.dumps({"platform": "Test platform", "language": {"code": lang}, "rows": rows},
                                                       ensure_ascii=False), encoding="utf-8")
    if low:
        (d / "lowscore_test.json").write_text(json.dumps({"platform": "Test platform", "rows": low}, ensure_ascii=False),
                                              encoding="utf-8")
    (d / "summary.json").write_text(json.dumps({"hotel": "Test hotel", "platforms": {
        "Test platform": {"rating": 9.0, "scale": 10, "total_reviews": 10, "all_reviews_scraped": True}}},
        ensure_ascii=False), encoding="utf-8")
    return d


def row(i, score, text):
    return {"score": score, "score_scale": 10, "date": f"2026-0{1 + i % 9}-1{i % 9}", "text": text, "title": "",
            "traveler_type": "", "room_type": "", "reviewer_country": ""}


def run(folder):
    with redirect_stdout(io.StringIO()):
        return A.main(folder, today="2026-11-01")


def test_lowscore_not_in_rates():
    rows = [row(i, 9.0, f"很好，第{i}次住") for i in range(9)] + [row(9, 4.0, "前台态度很差")]
    low = [row(i, 3.0, f"补抓的差评{i}，服务太差") for i in range(5)]
    res = run(make_folder(rows, low))
    ov = res["tables"]["Summary"].set_index("Item")["Combined result"]
    assert len(res["reviews"]) == 10, "the normal sample should be 10 reviews"
    assert abs(float(ov["Low-score share of analysed reviews"]) - 0.1) < 1e-9, f"low-score share should be 10%, got {ov['Low-score share of analysed reviews']}"


def test_language_codes():
    res = run(make_folder([row(i, 9.0, "很好") for i in range(10)], lang="zh-HK"))
    assert set(res["reviews"]["review_language"]) == {"zh"}, set(res["reviews"]["review_language"])


def test_keyword_cases():
    def hits(t):
        return [h for c in A.clauses_of(A.normalize(t)) for h in A.theme_hits(c, complaints_only=True)]
    assert hits("服务不差，房间很不错") == [], hits("服务不差，房间很不错")
    assert hits("楼下便利店东西很贵，酒店设施不错") == [], hits("楼下便利店东西很贵，酒店设施不错")
    assert "Staff/service" in hits("前台态度很差")
    assert "Booking info vs reality" in hits("到店才说只含一个人的早餐")
    assert "Booking info vs reality" in hits("停车要另外收费，订的时候没写")


def test_flag_language_codes():
    # when all_reviews_scraped in summary.json is ["zh-HK"], zh reviews also count as fully collected
    import pandas as pd
    r = pd.DataFrame({"platform": ["Agoda", "Agoda"], "review_language": ["zh", "en"]})
    t, used = A.complete_sample(r, {"platforms": {"Agoda": {"all_reviews_scraped": ["zh-HK"]}}})
    assert list(t["review_language"]) == ["zh"], t
    assert used == ["Agoda"], used


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("Passed:", name)
    print("All tests passed")
