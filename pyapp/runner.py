"""
JavaScript側から呼び出すエントリポイント。

Pyodideの `pyodide.globals.get()` 経由で使うため、トップレベルに関数を公開する。
"""
from __future__ import annotations

import json

from merger import merge_all
from excel_builder import build_xlsx_bytes


def process(
    rms_bytes: bytes,
    set_item_bytes: bytes,
    set_component_bytes: bytes,
    non_set_bytes: bytes,
) -> dict:
    """4CSV → Excelバイト列 と統計情報を返す。

    戻り値:
      {
        "xlsx": bytes,
        "stats": dict,
        "warnings": list[str],
      }
    """
    result = merge_all(
        rms_bytes=rms_bytes,
        set_item_bytes=set_item_bytes,
        set_component_bytes=set_component_bytes,
        non_set_bytes=non_set_bytes,
    )
    xlsx = build_xlsx_bytes(result.rows)
    return {
        "xlsx": xlsx,
        "stats": result.stats,
        "warnings": result.warnings,
    }


def preview(
    rms_bytes: bytes,
    set_item_bytes: bytes,
    set_component_bytes: bytes,
    non_set_bytes: bytes,
) -> str:
    """結合結果の統計・警告だけ返す（Excel生成はしない）。

    戻り値: JSON文字列
    """
    result = merge_all(
        rms_bytes=rms_bytes,
        set_item_bytes=set_item_bytes,
        set_component_bytes=set_component_bytes,
        non_set_bytes=non_set_bytes,
    )
    sample = [r for r in result.rows if not r.get("_blank")][:20]
    return json.dumps(
        {
            "stats": result.stats,
            "warnings": result.warnings[:50],
            "sample_rows": sample,
        },
        ensure_ascii=False,
    )
