/**
 * SALE一括 - 4CSVを結合して利益計算表用の行データを生成する。
 *
 * 入力: 各CSVの生バイト列 (Uint8Array) または文字列
 * 出力: { rows, stats, warnings }
 */
(function (global) {
  'use strict';

  // ---- CSVパーサー（Shift_JIS/UTF-8自動判別） -----------------------------
  function decodeBytes(bytes) {
    if (typeof bytes === 'string') return bytes;
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);

    const encodings = ['shift_jis', 'utf-8', 'cp932'];
    for (const enc of encodings) {
      try {
        const decoder = new TextDecoder(enc, { fatal: true });
        const text = decoder.decode(u8);
        // 日本語が出てこなければ失敗とみなす精度は不要、fatal:trueでエラーが出なければOK
        return text;
      } catch (e) {
        continue;
      }
    }
    // 最後の手段: UTF-8（非fatal）
    return new TextDecoder('utf-8').decode(u8);
  }

  /**
   * シンプルなCSVパーサ。以下のルールに対応：
   * - カンマ区切り
   * - ダブルクォートでの囲み + 内部のエスケープ ("" → ")
   * - クォート内の改行（セル内改行）を維持
   * - 行末の \r\n / \r / \n 対応
   *
   * 戻り値: 2次元配列 [[col0, col1, ...], ...]
   */
  function parseCSV(text) {
    const rows = [];
    let row = [];
    let field = '';
    let i = 0;
    let inQuotes = false;
    const len = text.length;

    while (i < len) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < len && text[i + 1] === '"') {
            field += '"';
            i += 2;
            continue;
          } else {
            inQuotes = false;
            i++;
            continue;
          }
        } else {
          field += ch;
          i++;
          continue;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
          i++;
          continue;
        }
        if (ch === ',') {
          row.push(field);
          field = '';
          i++;
          continue;
        }
        if (ch === '\r') {
          // \r or \r\n
          row.push(field);
          field = '';
          rows.push(row);
          row = [];
          if (i + 1 < len && text[i + 1] === '\n') i += 2;
          else i++;
          continue;
        }
        if (ch === '\n') {
          row.push(field);
          field = '';
          rows.push(row);
          row = [];
          i++;
          continue;
        }
        field += ch;
        i++;
      }
    }
    // 最後の1フィールド
    if (field.length > 0 || row.length > 0) {
      row.push(field);
      rows.push(row);
    }
    // 末尾の空行を除去（全列が空文字のもの）
    while (rows.length > 0 && rows[rows.length - 1].every((c) => c === '')) {
      rows.pop();
    }
    return rows;
  }

  /**
   * 2次元配列の1行目をヘッダーとして、以降を { colName: value } のオブジェクト配列に変換。
   * BOM(\uFEFF)があれば先頭から除去。
   */
  function csvToObjects(rows2d) {
    if (rows2d.length === 0) return { columns: [], records: [] };
    const headers = rows2d[0].map((h, i) => {
      if (i === 0 && h.charCodeAt(0) === 0xfeff) {
        return h.slice(1);
      }
      return h;
    });
    const records = [];
    for (let r = 1; r < rows2d.length; r++) {
      const row = rows2d[r];
      const obj = {};
      for (let c = 0; c < headers.length; c++) {
        obj[headers[c]] = (row[c] !== undefined ? row[c] : '');
      }
      records.push(obj);
    }
    return { columns: headers, records };
  }

  function readCSV(bytes) {
    const text = decodeBytes(bytes);
    const rows2d = parseCSV(text);
    return csvToObjects(rows2d);
  }

  // ---- ユーティリティ ---------------------------------------------------
  function toNumber(s) {
    if (s === null || s === undefined) return null;
    const t = String(s).trim().replace(/,/g, '');
    if (t === '') return null;
    const v = Number(t);
    return Number.isFinite(v) ? v : null;
  }

  // ---- 結合ロジック -----------------------------------------------------
  function mergeAll({ rmsBytes, setItemBytes, setComponentBytes, nonSetBytes }) {
    const warnings = [];

    const rms = readCSV(rmsBytes);
    const setItem = readCSV(setItemBytes);
    const setComponent = readCSV(setComponentBytes);
    const nonSet = readCSV(nonSetBytes);

    // 商品管理番号の列名は半角/全角の両方を許容
    const mgrKey = ['商品管理番号(商品URL)', '商品管理番号(商品URL)', '商品管理番号(商品URL)']
      .find(k => rms.columns.includes(k))
      || (rms.columns.includes('商品管理番号(商品URL)') ? '商品管理番号(商品URL)' : null)
      || (rms.columns.includes('商品管理番号(商品URL)') ? '商品管理番号(商品URL)' : null);

    const mgrKeyFinal = rms.columns.find(c => c.includes('商品管理番号'));
    if (!mgrKeyFinal) {
      throw new Error('RMSデータに商品管理番号列がありません');
    }

    for (const col of ['商品名', 'システム連携用SKU番号', '通常購入販売価格']) {
      if (!rms.columns.includes(col)) throw new Error(`RMSデータに必須列がありません: ${col}`);
    }
    for (const col of ['セット商品コード', '商品コード', '数量']) {
      if (!setItem.columns.includes(col)) throw new Error(`NE-セット商品コードに必須列がありません: ${col}`);
    }
    for (const col of ['商品コード', '原価']) {
      if (!setComponent.columns.includes(col)) throw new Error(`NE-セット構成物コードに必須列がありません: ${col}`);
      if (!nonSet.columns.includes(col)) throw new Error(`NE-セット無しコードに必須列がありません: ${col}`);
    }

    // セット商品コード -> [{商品コード, 数量}]
    // NE-セット商品コードCSVは、1セットに構成物が複数ある場合、
    // 商品コード/数量のセル内に改行区切りで複数値が格納されている。
    const setComponentsMap = new Map();
    for (const row of setItem.records) {
      const setCode = String(row['セット商品コード'] || '').trim();
      const itemCodeRaw = String(row['商品コード'] || '');
      const qtyRaw = String(row['数量'] || '');
      if (!setCode) continue;

      const itemCodes = itemCodeRaw
        .replace(/\r\n/g, '\n').replace(/\r/g, '\n')
        .split('\n').map(s => s.trim()).filter(s => s.length > 0);
      const qtyTokens = qtyRaw
        .replace(/\r\n/g, '\n').replace(/\r/g, '\n')
        .split('\n').map(s => s.trim());

      const components = [];
      for (let idx = 0; idx < itemCodes.length; idx++) {
        let qty = toNumber(qtyTokens[idx] || '');
        if (!(qty !== null && qty > 0)) qty = 1;
        components.push({ code: itemCodes[idx], qty: Math.round(qty) });
      }
      if (components.length > 0) {
        const existing = setComponentsMap.get(setCode) || [];
        setComponentsMap.set(setCode, existing.concat(components));
      }
    }

    const componentCostMap = new Map();
    for (const row of setComponent.records) {
      const code = String(row['商品コード'] || '').trim();
      const cost = toNumber(row['原価']);
      if (code && cost !== null) componentCostMap.set(code, cost);
    }

    const nonSetCostMap = new Map();
    for (const row of nonSet.records) {
      const code = String(row['商品コード'] || '').trim();
      const cost = toNumber(row['原価']);
      if (code && cost !== null) nonSetCostMap.set(code, cost);
    }

    // 商品管理番号ごとに商品名を前方向補完
    const nameByMgr = new Map();
    for (const row of rms.records) {
      const mgr = String(row[mgrKeyFinal] || '').trim();
      const name = String(row['商品名'] || '').trim();
      if (mgr && name && !nameByMgr.has(mgr)) {
        nameByMgr.set(mgr, name);
      }
    }

    const rows = [];
    const stats = {
      rms_total: rms.records.length,
      excluded_no_price: 0,
      cost_from_component: 0,
      cost_from_non_set: 0,
      cost_not_found: 0,
      output_rows: 0,
    };

    let lastMgrCode = null;
    for (const rrow of rms.records) {
      const mgrCode = String(rrow[mgrKeyFinal] || '').trim();
      const sku = String(rrow['システム連携用SKU番号'] || '').trim();
      const price = toNumber(rrow['通常購入販売価格']);

      if (price === null) {
        stats.excluded_no_price++;
        continue;
      }

      if (lastMgrCode !== null && mgrCode !== lastMgrCode) {
        rows.push({ _blank: true });
      }

      // 原価算出
      let cost = null;
      const components = setComponentsMap.get(sku);
      if (components && components.length > 0) {
        let total = 0;
        let allFound = true;
        for (const c of components) {
          let itemCost = componentCostMap.get(c.code);
          if (itemCost === undefined) itemCost = nonSetCostMap.get(c.code);
          if (itemCost === undefined) {
            allFound = false;
            warnings.push(`原価が見つからない商品コード: ${c.code} (セット ${sku} の構成物)`);
            break;
          }
          total += itemCost * c.qty;
        }
        if (allFound) {
          cost = total;
          stats.cost_from_component++;
        }
      } else {
        const c = nonSetCostMap.get(sku);
        if (c !== undefined) {
          cost = c;
          stats.cost_from_non_set++;
        }
      }
      if (cost === null) {
        stats.cost_not_found++;
        warnings.push(`原価が特定できないSKU: ${sku}`);
      }

      rows.push({
        _blank: false,
        商品名: nameByMgr.get(mgrCode) || String(rrow['商品名'] || ''),
        セット商品コード: sku,
        商品管理番号: mgrCode,
        SS用タイトル: '',
        通常価格: Math.round(price),
        原価: (cost !== null ? Math.round(cost) : null),
        送料: null,
        バリエ1: String(rrow['バリエーション項目選択肢1'] || ''),
        バリエ2: String(rrow['バリエーション項目選択肢2'] || ''),
        バリエ3: String(rrow['バリエーション項目選択肢3'] || ''),
        バリエ4: String(rrow['バリエーション項目選択肢4'] || ''),
      });
      stats.output_rows++;
      lastMgrCode = mgrCode;
    }

    return { rows, stats, warnings };
  }

  // エクスポート
  const api = { mergeAll, readCSV, parseCSV, decodeBytes };
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.SaleIkkatsuMerger = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
