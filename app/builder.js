/**
 * SALE一括 - 結合済みの行データから利益計算表Excelを生成する。
 * Python版 (openpyxl) と同じ書式・数式・配色を再現。
 *
 * 依存: ExcelJS (window.ExcelJS または require('exceljs'))
 */
(function (global) {
  'use strict';

  const ExcelJS = (typeof require !== 'undefined') ? require('exceljs') : global.ExcelJS;

  // ---- 配色 --------------------------------------------------------------
  const COLOR_RED = 'EE4D3D';
  const COLOR_GREEN = '3DEE56';
  const COLOR_GRAY = '9E9DA7';
  const COLOR_BLUE = 'DDEBF7';
  const COLOR_ORANGE = 'FCE4D6';
  const COLOR_LIGHTGREEN = 'E2F0D9';
  const COLOR_BROWN = 'EDE1D1';

  // ---- 列幅 (納品版と一致) ----------------------------------------------
  const COLUMN_WIDTHS = {
    A: 55.0, B: 35.57, C: 15.14, D: 15.57, E: 10.29, F: 6.0, G: 6.43,
    H: 26.0, I: 26.0, J: 26.0, K: 26.0, L: 12.71, M: 17.57, N: 15.14, O: 11.14,
    P: 15.14, Q: 17.57, R: 13.0, S: 15.14, T: 17.57, U: 13.0,
    V: 15.14, W: 17.57, X: 13.0, Y: 15.14, Z: 17.57, AA: 13.0, AB: 35.57,
  };

  // ---- ヘッダー定義 -----------------------------------------------------
  const HEADERS = [
    ['A', '商品名', null],
    ['B', 'セット商品コード', COLOR_RED],
    ['C', '商品管理番号', COLOR_RED],
    ['D', 'SS用タイトル', COLOR_RED],
    ['E', '通常価格', COLOR_GREEN],
    ['F', '原価', COLOR_GREEN],
    ['G', '送料', COLOR_GREEN],
    ['H', 'バリエ内容1(送料用)', COLOR_GRAY],
    ['I', 'バリエ内容2(送料用)', COLOR_GRAY],
    ['J', 'バリエ内容3(送料用)', COLOR_GRAY],
    ['K', 'バリエ内容4(送料用)', COLOR_GRAY],
    ['L', '楽天手数料', COLOR_GRAY],
    ['M', '楽天ペイ利用料', COLOR_GRAY],
    ['N', 'ポイント原資', COLOR_GRAY],
    ['O', 'NE手数料', COLOR_GRAY],
    ['P', '10%OFF価格', COLOR_BLUE],
    ['Q', '10%OFF利益額', COLOR_BLUE],
    ['R', '10%OFF利益率', COLOR_BLUE],
    ['S', '20%OFF価格', COLOR_ORANGE],
    ['T', '20%OFF利益額', COLOR_ORANGE],
    ['U', '20%OFF利益率', COLOR_ORANGE],
    ['V', '30%OFF価格', COLOR_LIGHTGREEN],
    ['W', '30%OFF利益額', COLOR_LIGHTGREEN],
    ['X', '30%OFF利益率', COLOR_LIGHTGREEN],
    ['Y', '50%OFF価格', COLOR_BROWN],
    ['Z', '50%OFF利益額', COLOR_BROWN],
    ['AA', '50%OFF利益率', COLOR_BROWN],
    ['AB', 'セット商品コード', null],
  ];

  const DISCOUNT_FILLS = {
    P: COLOR_BLUE, Q: COLOR_BLUE, R: COLOR_BLUE,
    S: COLOR_ORANGE, T: COLOR_ORANGE, U: COLOR_ORANGE,
    V: COLOR_LIGHTGREEN, W: COLOR_LIGHTGREEN, X: COLOR_LIGHTGREEN,
    Y: COLOR_BROWN, Z: COLOR_BROWN, AA: COLOR_BROWN,
  };

  const GROUP_COLUMNS = ['H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];

  const FIXED_L = 6;
  const FIXED_M = 3.5;
  const FIXED_N = 1;
  const FIXED_O = 10;

  const FONT_NAME = '游ゴシック';
  const FONT_SIZE = 11;

  // ヘルパー: 塗りつぶし
  function solidFill(hex) {
    return {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF' + hex },
      bgColor: { argb: 'FF' + hex },
    };
  }

  const THIN_BORDER = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };

  const ALIGN_CENTER = { horizontal: 'center', vertical: 'middle' };
  const ALIGN_LEFT = { horizontal: 'left', vertical: 'middle' };

  /**
   * 行データからExcelバイナリ(ArrayBuffer)を生成する。
   *
   * @param {Array} rows - merger.mergeAll の rows
   * @returns {Promise<ArrayBuffer>}
   */
  async function buildXlsx(rows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('利益計算表');

    const baseFont = { name: FONT_NAME, size: FONT_SIZE };
    const boldFont = { name: FONT_NAME, size: FONT_SIZE, bold: true };

    // ヘッダー
    for (const [colLetter, label, color] of HEADERS) {
      const cell = ws.getCell(`${colLetter}1`);
      cell.value = label;
      cell.font = baseFont;
      cell.alignment = ALIGN_CENTER;
      cell.border = THIN_BORDER;
      if (color) cell.fill = solidFill(color);
    }

    // データ行
    let r = 2;
    for (const rec of rows) {
      if (rec._blank) {
        // 空白行: 割引列の塗りだけ維持
        for (const [col, color] of Object.entries(DISCOUNT_FILLS)) {
          const cell = ws.getCell(`${col}${r}`);
          cell.fill = solidFill(color);
          cell.font = baseFont;
          cell.border = THIN_BORDER;
        }
        r++;
        continue;
      }

      // A: 商品名
      const aCell = ws.getCell(`A${r}`);
      aCell.value = rec['商品名'];
      aCell.alignment = ALIGN_LEFT;

      // B: セット商品コード
      ws.getCell(`B${r}`).value = rec['セット商品コード'];
      ws.getCell(`B${r}`).alignment = ALIGN_CENTER;

      // C: 商品管理番号 (数値化できれば数値)
      const mgr = rec['商品管理番号'];
      const mgrNum = Number(mgr);
      ws.getCell(`C${r}`).value = Number.isFinite(mgrNum) && String(mgrNum) === String(mgr) ? mgrNum : mgr;
      ws.getCell(`C${r}`).alignment = ALIGN_CENTER;

      // D: SS用タイトル (空)
      ws.getCell(`D${r}`).value = '';
      ws.getCell(`D${r}`).alignment = ALIGN_LEFT;

      // E: 通常価格
      ws.getCell(`E${r}`).value = rec['通常価格'];
      ws.getCell(`E${r}`).numFmt = '0';
      ws.getCell(`E${r}`).alignment = ALIGN_CENTER;

      // F: 原価
      if (rec['原価'] !== null && rec['原価'] !== undefined) {
        ws.getCell(`F${r}`).value = rec['原価'];
      }
      ws.getCell(`F${r}`).numFmt = '0';
      ws.getCell(`F${r}`).alignment = ALIGN_CENTER;

      // G: 送料 (空欄、通貨書式)
      ws.getCell(`G${r}`).numFmt = '\\¥#,##0';
      ws.getCell(`G${r}`).alignment = ALIGN_CENTER;

      // H-K: バリエ
      ws.getCell(`H${r}`).value = rec['バリエ1'];
      ws.getCell(`I${r}`).value = rec['バリエ2'];
      ws.getCell(`J${r}`).value = rec['バリエ3'];
      ws.getCell(`K${r}`).value = rec['バリエ4'];
      for (const c of ['H', 'I', 'J', 'K']) {
        ws.getCell(`${c}${r}`).alignment = ALIGN_CENTER;
      }

      // L-O: 固定値
      ws.getCell(`L${r}`).value = FIXED_L;
      ws.getCell(`L${r}`).numFmt = '0';
      ws.getCell(`M${r}`).value = FIXED_M;
      ws.getCell(`M${r}`).numFmt = '0.0';
      ws.getCell(`N${r}`).value = FIXED_N;
      ws.getCell(`N${r}`).numFmt = '0';
      ws.getCell(`O${r}`).value = FIXED_O;
      ws.getCell(`O${r}`).numFmt = '0';
      for (const c of ['L', 'M', 'N', 'O']) {
        ws.getCell(`${c}${r}`).alignment = ALIGN_CENTER;
      }

      // P-AA 数式
      ws.getCell(`P${r}`).value = { formula: `ROUNDDOWN(E${r}*0.9,-1)` };
      ws.getCell(`Q${r}`).value = { formula: `ROUNDDOWN(P${r}-F${r}-G${r}-(P${r}*0.01*L${r})-(P${r}*0.01*M${r})-(P${r}*0.01*N${r})-O${r},0)` };
      ws.getCell(`R${r}`).value = { formula: `IFERROR(ROUNDDOWN((Q${r}/P${r})*100,0),"")` };

      ws.getCell(`S${r}`).value = { formula: `ROUNDDOWN(E${r}*0.8,-1)` };
      ws.getCell(`T${r}`).value = { formula: `ROUNDDOWN(S${r}-F${r}-G${r}-(S${r}*0.01*L${r})-(S${r}*0.01*M${r})-(S${r}*0.01*N${r})-O${r},0)` };
      ws.getCell(`U${r}`).value = { formula: `IFERROR(ROUNDDOWN((T${r}/S${r})*100,0),"")` };

      ws.getCell(`V${r}`).value = { formula: `ROUNDDOWN(E${r}*0.7,-1)` };
      ws.getCell(`W${r}`).value = { formula: `ROUNDDOWN(V${r}-F${r}-G${r}-(V${r}*0.01*L${r})-(V${r}*0.01*M${r})-(V${r}*0.01*N${r})-O${r},0)` };
      ws.getCell(`X${r}`).value = { formula: `IFERROR(ROUNDDOWN((W${r}/V${r})*100,0),"")` };

      ws.getCell(`Y${r}`).value = { formula: `ROUNDDOWN(E${r}*0.5,-1)` };
      ws.getCell(`Z${r}`).value = { formula: `ROUNDDOWN(Y${r}-F${r}-G${r}-(Y${r}*0.01*L${r})-(Y${r}*0.01*M${r})-(Y${r}*0.01*N${r})-O${r},0)` };
      ws.getCell(`AA${r}`).value = { formula: `IFERROR(ROUNDDOWN((Z${r}/Y${r})*100,0),"")` };

      // 割引列の書式
      for (const [col, color] of Object.entries(DISCOUNT_FILLS)) {
        const cell = ws.getCell(`${col}${r}`);
        cell.alignment = ALIGN_CENTER;
        cell.numFmt = '0';
        cell.fill = solidFill(color);
      }
      // 価格列(P/S/V/Y)は太字
      for (const col of ['P', 'S', 'V', 'Y']) {
        ws.getCell(`${col}${r}`).font = boldFont;
      }

      // AB: セット商品コード (B列コピー)
      ws.getCell(`AB${r}`).value = rec['セット商品コード'];
      ws.getCell(`AB${r}`).alignment = ALIGN_CENTER;

      // 全セルにフォント・罫線
      for (let colIdx = 1; colIdx <= 28; colIdx++) {
        const cell = ws.getRow(r).getCell(colIdx);
        if (!cell.font || (cell.font.name !== FONT_NAME && !cell.font.bold)) {
          cell.font = baseFont;
        }
        cell.border = THIN_BORDER;
      }

      r++;
    }

    // 列幅
    for (const [col, width] of Object.entries(COLUMN_WIDTHS)) {
      ws.getColumn(col).width = width;
    }

    // グループ化 (H-O)
    for (const col of GROUP_COLUMNS) {
      ws.getColumn(col).outlineLevel = 1;
      ws.getColumn(col).hidden = false;
    }

    // フリーズペイン
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    // バイナリ出力
    const buf = await wb.xlsx.writeBuffer();
    return buf;
  }

  const api = { buildXlsx };
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.SaleIkkatsuBuilder = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
