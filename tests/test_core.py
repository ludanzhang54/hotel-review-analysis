"""核心统计的回归测试：python3 tests/test_core.py（全部通过才打印"全部通过"）。
测补抓的低分点评不进比例、语言代码统一、几句容易误判的话、空样本。"""
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
    (d / "reviews_test_zh.json").write_text(json.dumps({"platform": "测试平台", "language": {"code": lang}, "rows": rows},
                                                       ensure_ascii=False), encoding="utf-8")
    if low:
        (d / "lowscore_test.json").write_text(json.dumps({"platform": "测试平台", "rows": low}, ensure_ascii=False),
                                              encoding="utf-8")
    (d / "summary.json").write_text(json.dumps({"hotel": "测试酒店", "platforms": {
        "测试平台": {"rating": 9.0, "scale": 10, "total_reviews": 10, "all_reviews_scraped": True}}},
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
    ov = res["tables"]["综合汇总"].set_index("项目")["合并结果"]
    assert len(res["reviews"]) == 10, "正常样本应是 10 条"
    assert abs(float(ov["差评占本次分析的点评"]) - 0.1) < 1e-9, f"差评占比应是 10%，实际 {ov['差评占本次分析的点评']}"


def test_language_codes():
    res = run(make_folder([row(i, 9.0, "很好") for i in range(10)], lang="zh-HK"))
    assert set(res["reviews"]["review_language"]) == {"zh"}, set(res["reviews"]["review_language"])


def test_keyword_cases():
    def hits(t):
        return [h for c in A.clauses_of(A.normalize(t)) for h in A.theme_hits(c, complaints_only=True)]
    assert hits("服务不差，房间很不错") == [], hits("服务不差，房间很不错")
    assert hits("楼下便利店东西很贵，酒店设施不错") == [], hits("楼下便利店东西很贵，酒店设施不错")
    assert "员工/服务" in hits("前台态度很差")
    assert "订房信息和实际不符" in hits("到店才说只含一个人的早餐")
    assert "订房信息和实际不符" in hits("停车要另外收费，订的时候没写")


def test_flag_language_codes():
    # summary.json 里 all_reviews_scraped 写成 ["zh-HK"] 时，zh 的点评也算抓全
    import pandas as pd
    r = pd.DataFrame({"platform": ["Agoda", "Agoda"], "review_language": ["zh", "en"]})
    t, used = A.complete_sample(r, {"platforms": {"Agoda": {"all_reviews_scraped": ["zh-HK"]}}})
    assert list(t["review_language"]) == ["zh"], t
    assert used == ["Agoda"], used


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("通过：", name)
    print("全部通过")
