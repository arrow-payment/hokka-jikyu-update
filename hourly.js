'use strict';

// ほっか直営店 時給データ → インポート用Excel 変換ロジック
// (旧 hokka-jikyu-import スキルの convert.py を移植。変換ルールはユーザー確認済み・インポート実績あり)
//
// 元データの列は見出し名で探す(列の追加で位置がずれても壊れないように)。
// 使うのは HOURLY_COLUMNS の列だけで、勤務曜日・残業・深夜・資格・宅配手当などは使わない。
//
// 変換ルール:
// - 従業員単位にまとめ、1店舗のみの人はその条件をメイン時給(E列空欄)として登録
// - 複数店舗の人は親店舗の条件をメイン時給、子店舗で条件が違う場合だけ
//   E列にその所属コードを入れた例外時給の行を追加する(親と同条件の子店舗は出さない)
// - 休日手当あり → 基本給の行(該当曜日以外) + 基本給+日別手当の行(該当曜日)の2グループ
// - 休日手当なし → 全曜日+祝に1の1行
// - 早朝手当時間あり → 00:00-05:00 / 05:00-08:00(+早朝手当) / 08:00-24:00 の3分割
// - 両方 → 2グループ×3分割=6行。土日祝の早朝は日別手当と早朝手当を両方加算
// - 1行目に __format_version / 1.1.1、2行目にヘッダー。曜日フラグは文字列の '1'
// - 社員名・カナは親店舗(1店舗のみの人はその行)の氏名・フリガナ

const DAYS = ['日', '月', '火', '水', '木', '金', '土', '祝'];
const PREM_MAP = { '土日祝': new Set(['土', '日', '祝']), '日祝': new Set(['日', '祝']) };
// 2026-07-15追加分にハイフン区切りの表記ゆれあり。同一時間帯として扱う(ユーザー確認済み)
const EARLY_BAND_ALIASES = new Set(['5:00~8:00', '5:00-8:00']);
const HOURLY_COLUMNS = {
  store: '所属コード',
  emp: '社員番号',
  name: '氏名',
  kana: '氏名(フリガナ)',
  parent: '親店舗',
  child: '子店舗',
  base: '基本給',
  early: '早朝手当',
  daily: '日別手当',
  holiday: '休日手当該当曜日',
  earlyBand: '早朝手当時間',
};

// 「土・日・祝」のような中黒区切りも同じ曜日指定として扱う
const normalizeHoliday = (v) => str(v).replace(/[・、,\s]/g, '');

function condOf(r, col) {
  return {
    base: r[col.base] ?? null,
    early: r[col.early] || 0,
    daily: r[col.daily] || 0,
    holiday: normalizeHoliday(r[col.holiday]),
    earlyBand: str(r[col.earlyBand]),
  };
}

const sameCond = (a, b) => JSON.stringify(a) === JSON.stringify(b);

function validateCond(emp, { base, early, daily, holiday, earlyBand }) {
  if (typeof base !== 'number' || base <= 0) throw new ConvertError(`${emp}: 基本給が不正です(${base})`);
  if (typeof early !== 'number') throw new ConvertError(`${emp}: 早朝手当が数値ではありません(${early})`);
  if (typeof daily !== 'number') throw new ConvertError(`${emp}: 日別手当が数値ではありません(${daily})`);
  if (holiday && !PREM_MAP[holiday]) throw new ConvertError(`${emp}: 想定外の休日手当該当曜日です(${holiday})`);
  if (earlyBand && !EARLY_BAND_ALIASES.has(earlyBand)) throw new ConvertError(`${emp}: 想定外の早朝手当時間です(${earlyBand})`);
}

function buildHourlyRows([emp, name, kana], storeId, { base, early, daily, holiday, earlyBand }) {
  const prem = PREM_MAP[holiday] || new Set();
  const groups = prem.size
    ? [[new Set(DAYS.filter((d) => !prem.has(d))), 0], [prem, daily]]
    : [[new Set(DAYS), 0]];
  const rows = [];
  for (const [days, add] of groups) {
    const flags = DAYS.map((d) => (days.has(d) ? '1' : null));
    if (earlyBand) {
      rows.push([emp, name, kana, base + add, storeId, '00:00', '05:00', ...flags]);
      rows.push([emp, name, kana, base + add + early, storeId, '05:00', '08:00', ...flags]);
      rows.push([emp, name, kana, base + add, storeId, '08:00', '24:00', ...flags]);
    } else {
      rows.push([emp, name, kana, base + add, storeId, '00:00', '24:00', ...flags]);
    }
  }
  return rows;
}

// sheetRows: 「元データ」シートの全行(1行目が見出し、A列始まりの配列の配列)
function convertHourly(sheetRows) {
  const col = findColumns(sheetRows[0], HOURLY_COLUMNS);
  const rows = dataRows(sheetRows, col);

  const byEmp = groupByEmployee(rows, col);

  const outRows = [];
  const stats = { '手当なし': 0, '休日手当のみ': 0, '早朝手当のみ': 0, '両方': 0, '例外時給あり': 0 };
  for (const [emp, { main, children }] of byEmp) {
    const person = [emp, str(main[col.name]) || null, str(main[col.kana]) || null];
    const mainCond = condOf(main, col);
    validateCond(emp, mainCond);
    const { holiday, earlyBand } = mainCond;
    if (PREM_MAP[holiday] && earlyBand) stats['両方']++;
    else if (PREM_MAP[holiday]) stats['休日手当のみ']++;
    else if (earlyBand) stats['早朝手当のみ']++;
    else stats['手当なし']++;
    // E列(店舗ID)空欄 = メイン時給
    outRows.push(...buildHourlyRows(person, null, mainCond));

    let hasException = false;
    for (const r of children) {
      const cond = condOf(r, col);
      if (sameCond(cond, mainCond)) continue;
      validateCond(emp, cond);
      const storeId = str(r[col.store]);
      if (!storeId) throw new ConvertError(`${emp}: 子店舗の所属コードが空欄です`);
      outRows.push(...buildHourlyRows(person, storeId, cond));
      hasException = true;
    }
    if (hasException) stats['例外時給あり']++;
  }

  // 出力後の自己検証
  const outEmps = new Set(outRows.map((r) => r[0]));
  if (outEmps.size !== byEmp.size || [...byEmp.keys()].some((e) => !outEmps.has(e))) {
    throw new ConvertError('社員コードが入力と一致しません');
  }
  if (outRows.some((r) => !r.slice(7, 15).some(Boolean))) {
    throw new ConvertError('曜日フラグが1つもない行があります');
  }

  return { outRows, empCount: byEmp.size, stats };
}

const HOURLY = {
  key: 'hourly',
  label: '時給',
  formatVersion: '1.1.1',
  header: ['社員コード', '社員名', '社員名(カナ)', '時給', '対象部署/店舗ID',
    '開始時間', '終了時間',
    '日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '祝日'],
  previewLabels: { 2: 'カナ', 4: '店舗ID', 5: '開始', 6: '終了',
    7: '日', 8: '月', 9: '火', 10: '水', 11: '木', 12: '金', 13: '土', 14: '祝' },
  colWidths: [14, 10, 12, 8, 16, 10, 10],
  filePrefix: '時給インポートデータ',
  // 店舗IDが入る行 = 例外時給なのでプレビューで色を変える
  highlighter: () => (row) => Boolean(row[4]),
  highlightLabel: '例外時給ありのみ',
  convert: convertHourly,
};
