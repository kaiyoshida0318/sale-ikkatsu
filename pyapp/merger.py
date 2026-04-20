"""
4つのCSVを結合し、利益計算表の1行分に相当する辞書のリストを生成する。

Pyodide（ブラウザ内Python）で動作。入力はすべてbytesで受け取る。
"""
from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Any

import pandas as pd


ENCODINGS_TO_TRY = ["shift_jis", "cp932", "utf-8-sig", "utf-8"]


def _read_csv(raw: bytes) -> pd.DataFrame:
    """複数のエンコーディングを試してCSVを読み込む。"""
    last_err: Exception | None = None
    for enc in ENCODINGS_TO_TRY:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, dtype=str).fillna("")
        except (UnicodeDecodeError, pd.errors.ParserError) as e:
            last_err = e
            continue
    raise ValueError(
        f"CSVの読み込みに失敗しました。対応エンコーディング: {ENCODINGS_TO_TRY}"
    ) from last_err


@dataclass
class MergeResult:
    rows: list[dict[str, Any]]
    stats: dict[str, int]
    warnings: list[str]


def _to_number(s: str) -> float | None:
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def merge_all(
    rms_bytes: bytes,
    set_item_bytes: bytes,
    set_component_bytes: bytes,
    non_set_bytes: bytes,
) -> MergeResult:
    """4つのCSVを結合する。

    結合ルール:
      RMS「システム連携用SKU番号」 = NE-セット商品コード「セット商品コード」
      NE-セット商品コード「商品コード」 = NE-セット構成物コード「商品コード」
      NE-セット構成物コードに無い商品コードは、NE-セット無しコードから原価取得
    """
    warnings: list[str] = []

    rms = _read_csv(rms_bytes)
    set_item = _read_csv(set_item_bytes)
    set_component = _read_csv(set_component_bytes)
    non_set = _read_csv(non_set_bytes)

    rms_required = [
        "商品管理番号(商品URL)",
        "商品名",
        "システム連携用SKU番号",
        "通常購入販売価格",
    ]
    # 全角括弧のバリエーションも許容
    alt_mgr_keys = ["商品管理番号（商品URL）", "商品管理番号(商品URL)"]
    mgr_key = next((k for k in alt_mgr_keys if k in rms.columns), None)
    if mgr_key is None:
        raise ValueError(
            f"RMSデータに商品管理番号列がありません（許容キー: {alt_mgr_keys}）"
        )

    for col in ["商品名", "システム連携用SKU番号", "通常購入販売価格"]:
        if col not in rms.columns:
            raise ValueError(f"RMSデータに必須列がありません: {col}")

    for col in ["セット商品コード", "商品コード", "数量"]:
        if col not in set_item.columns:
            raise ValueError(f"NE-セット商品コードに必須列がありません: {col}")

    for col in ["商品コード", "原価"]:
        if col not in set_component.columns:
            raise ValueError(f"NE-セット構成物コードに必須列がありません: {col}")
        if col not in non_set.columns:
            raise ValueError(f"NE-セット無しコードに必須列がありません: {col}")

    # セット商品コード -> [{"商品コード": str, "数量": int}]
    # 改行区切りで複数構成物が入っているセルを展開
    set_components_map: dict[str, list[dict[str, Any]]] = {}
    for _, row in set_item.iterrows():
        set_code = str(row["セット商品コード"]).strip()
        item_code_raw = str(row["商品コード"])
        qty_raw = str(row["数量"])
        if not set_code:
            continue

        item_codes = [
            c.strip()
            for c in item_code_raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            if c.strip()
        ]
        qty_tokens = [
            q.strip()
            for q in qty_raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        ]

        components: list[dict[str, Any]] = []
        for idx, code in enumerate(item_codes):
            q_str = qty_tokens[idx] if idx < len(qty_tokens) else ""
            qty = _to_number(q_str)
            if qty is None or qty <= 0:
                qty = 1
            components.append({"商品コード": code, "数量": int(qty)})

        if components:
            set_components_map.setdefault(set_code, []).extend(components)

    component_cost_map: dict[str, float] = {}
    for _, row in set_component.iterrows():
        code = str(row["商品コード"]).strip()
        cost = _to_number(row["原価"])
        if code and cost is not None:
            component_cost_map[code] = cost

    non_set_cost_map: dict[str, float] = {}
    for _, row in non_set.iterrows():
        code = str(row["商品コード"]).strip()
        cost = _to_number(row["原価"])
        if code and cost is not None:
            non_set_cost_map[code] = cost

    # 商品管理番号ごとに商品名を前方向補完
    name_by_mgr: dict[str, str] = {}
    for _, rrow in rms.iterrows():
        mgr = str(rrow[mgr_key]).strip()
        name = str(rrow.get("商品名", "")).strip()
        if mgr and name and mgr not in name_by_mgr:
            name_by_mgr[mgr] = name

    rows: list[dict[str, Any]] = []
    stats = {
        "rms_total": len(rms),
        "excluded_no_price": 0,
        "cost_from_component": 0,
        "cost_from_non_set": 0,
        "cost_not_found": 0,
        "output_rows": 0,
    }

    last_mgr_code: str | None = None

    for _, rrow in rms.iterrows():
        mgr_code = str(rrow[mgr_key]).strip()
        sku = str(rrow["システム連携用SKU番号"]).strip()
        price = _to_number(rrow["通常購入販売価格"])

        if price is None:
            stats["excluded_no_price"] += 1
            continue

        if last_mgr_code is not None and mgr_code != last_mgr_code:
            rows.append({"_blank": True})

        cost: float | None = None
        components = set_components_map.get(sku)
        if components:
            total = 0.0
            all_found = True
            for c in components:
                item_cost = component_cost_map.get(c["商品コード"])
                if item_cost is None:
                    item_cost = non_set_cost_map.get(c["商品コード"])
                if item_cost is None:
                    all_found = False
                    warnings.append(
                        f"原価が見つからない商品コード: {c['商品コード']} "
                        f"(セット {sku} の構成物)"
                    )
                    break
                total += item_cost * c["数量"]
            if all_found:
                cost = total
                stats["cost_from_component"] += 1
        else:
            c = non_set_cost_map.get(sku)
            if c is not None:
                cost = c
                stats["cost_from_non_set"] += 1

        if cost is None:
            stats["cost_not_found"] += 1
            warnings.append(f"原価が特定できないSKU: {sku}")

        rows.append(
            {
                "_blank": False,
                "商品名": name_by_mgr.get(mgr_code, str(rrow.get("商品名", ""))),
                "セット商品コード": sku,
                "商品管理番号": mgr_code,
                "SS用タイトル": "",
                "通常価格": int(price),
                "原価": int(cost) if cost is not None else None,
                "送料": None,
                "バリエ1": str(rrow.get("バリエーション項目選択肢1", "")),
                "バリエ2": str(rrow.get("バリエーション項目選択肢2", "")),
                "バリエ3": str(rrow.get("バリエーション項目選択肢3", "")),
                "バリエ4": str(rrow.get("バリエーション項目選択肢4", "")),
            }
        )
        stats["output_rows"] += 1
        last_mgr_code = mgr_code

    return MergeResult(rows=rows, stats=stats, warnings=warnings)
