'use strict';

// 時給・通勤手当の変換で共通に使う処理

const SOURCE_SHEET = '元データ';
const OUTPUT_SHEET = 'インポート用';

class ConvertError extends Error {
  constructor(message, details = []) {
    super(message);
    this.details = details;
  }
}

const str = (v) => (v === null || v === undefined ? '' : String(v).trim());
// 見出しは空白(全角含む)を除き、全角かっこを半角に揃えて完全一致で探す
const normalizeHeader = (v) => str(v).replace(/\s/g, '').replace(/（/g, '(').replace(/）/g, ')');

// columns: { key: 見出し名 } → { key: 列index }
function findColumns(headerRow, columns) {
  const names = (headerRow || []).map(normalizeHeader);
  const col = {};
  const missing = [];
  for (const [key, name] of Object.entries(columns)) {
    const i = names.indexOf(name);
    if (i < 0) missing.push(name);
    col[key] = i;
  }
  if (missing.length) throw new ConvertError('元データに必要な列が見つかりません', missing);
  return col;
}

// 社員番号が入っている行だけを返す
function dataRows(sheetRows, col) {
  const rows = sheetRows.slice(1).filter((r) => r[col.emp]);
  if (!rows.length) throw new ConvertError('データ行がありません');
  return rows;
}

// 従業員ごとにまとめ、メインの行(1店舗のみの人はその行、複数店舗の人は親店舗の行)を決める。
// 複数店舗の人は親店舗1行+子店舗の行で構成されている前提。col には emp/parent/child が必要
function groupByEmployee(rows, col) {
  const byEmp = new Map();
  for (const r of rows) {
    const emp = str(r[col.emp]);
    if (!byEmp.has(emp)) byEmp.set(emp, []);
    byEmp.get(emp).push(r);
  }

  const result = new Map();
  const badStores = [];
  for (const [emp, empRows] of byEmp) {
    if (empRows.length === 1) {
      result.set(emp, { main: empRows[0], children: [], rows: empRows });
      continue;
    }
    const parents = empRows.filter((r) => r[col.parent]);
    const children = empRows.filter((r) => r[col.child]);
    if (parents.length !== 1 || parents.length + children.length !== empRows.length) {
      badStores.push(emp);
      continue;
    }
    result.set(emp, { main: parents[0], children, rows: empRows });
  }
  if (badStores.length) {
    throw new ConvertError('親店舗・子店舗の指定が想定と違う従業員がいます', badStores);
  }
  return result;
}

// XLSX: SheetJS
// sheetName を省略すると先頭のシートを読む
function readSheetRows(XLSX, data, sheetName) {
  const wb = XLSX.read(data, { type: 'array' });
  const name = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) throw new ConvertError(`「${name}」シートが見つかりません`);
  if (!ws['!ref']) throw new ConvertError('データ行がありません');
  // 見出しが1行目・A列始まりになるよう読み取り範囲を固定する
  const range = XLSX.utils.decode_range(ws['!ref']);
  range.s = { r: 0, c: 0 };
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, range });
}

const readSourceRows = (XLSX, data) => readSheetRows(XLSX, data, SOURCE_SHEET);

function buildWorkbook(XLSX, { formatVersion, header, colWidths, outRows }) {
  const ws = XLSX.utils.aoa_to_sheet([['__format_version', formatVersion], header, ...outRows]);
  ws['!cols'] = colWidths.map((wch) => ({ wch }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, OUTPUT_SHEET);
  return wb;
}

function outputFileName(prefix, d = new Date()) {
  const p = (n) => String(n).padStart(2, '0');
  return `${prefix}_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.xlsx`;
}
