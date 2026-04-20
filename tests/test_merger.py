"""merger の単体テスト。ローカルでpytest実行可能。"""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "pyapp"))

from merger import merge_all  # noqa: E402


SAMPLE_DIR = ROOT / "sample_data"


def _read(name: str) -> bytes:
    return (SAMPLE_DIR / name).read_bytes()


@pytest.fixture
def sample_bytes():
    return {
        "rms": _read("RMSデータ.csv"),
        "set_item": _read("NE-セット商品コード.csv"),
        "set_component": _read("NE-セット構成物コード.csv"),
        "non_set": _read("NE-セット無しコード.csv"),
    }


def test_merge_returns_rows(sample_bytes):
    res = merge_all(
        rms_bytes=sample_bytes["rms"],
        set_item_bytes=sample_bytes["set_item"],
        set_component_bytes=sample_bytes["set_component"],
        non_set_bytes=sample_bytes["non_set"],
    )
    assert res.stats["output_rows"] > 0


def test_price_excluded_when_empty(sample_bytes):
    res = merge_all(
        rms_bytes=sample_bytes["rms"],
        set_item_bytes=sample_bytes["set_item"],
        set_component_bytes=sample_bytes["set_component"],
        non_set_bytes=sample_bytes["non_set"],
    )
    for r in res.rows:
        if not r.get("_blank"):
            assert r["通常価格"] is not None


def test_set_product_cost_computed(sample_bytes):
    res = merge_all(
        rms_bytes=sample_bytes["rms"],
        set_item_bytes=sample_bytes["set_item"],
        set_component_bytes=sample_bytes["set_component"],
        non_set_bytes=sample_bytes["non_set"],
    )
    # セット: 140×2 + 5×1 = 285
    target = next(
        (r for r in res.rows if r.get("セット商品コード") == "ep-anker-25-01-q10-no-no"),
        None,
    )
    assert target is not None
    assert target["原価"] == 285
