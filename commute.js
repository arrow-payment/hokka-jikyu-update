'use strict';

// ほっか直営店 通勤手当データ → インポート用Excel 変換ロジック
//
// 変換ルール(ユーザー指定・アウトプットサンプル_通勤手当.xlsx準拠):
// - 元データ1行(=1店舗)につき1行。往復運賃が0・空欄の行は出さない
// - 通勤手当 = 往復運賃、日額通勤手当の月上限額 = 定期代
// - 親店舗・子店舗に関係なく、E列(対象部署/店舗ID)には所属コードを入れる
// - 日額通勤手当の月上限は常に「使用する」
// - 元データの「通勤手当」列は使わない
// - 1行目に __format_version / 1.1.0。値はサンプルに合わせて文字列で出す

const COMMUTE_COLUMNS = {
  store: '所属コード',
  emp: '社員番号',
  name: '氏名',
  kana: '氏名(フリガナ)',
  fare: '往復運賃',
  pass: '定期代',
};

function convertCommute(sheetRows) {
  const col = findColumns(sheetRows[0], COMMUTE_COLUMNS);
  const rows = dataRows(sheetRows, col);

  // 同じ人の行がまとまるよう、従業員ごと(元データでの初出順)に並べる
  const byEmp = new Map();
  const duplicates = [];
  for (const r of rows) {
    const emp = str(r[col.emp]);
    const fare = r[col.fare];
    const pass = r[col.pass];
    if (fare === null || fare === '' || fare === 0) continue;
    if (typeof fare !== 'number' || fare < 0) throw new ConvertError(`${emp}: 往復運賃が不正です(${fare})`);
    if (typeof pass !== 'number' || pass <= 0) throw new ConvertError(`${emp}: 往復運賃があるのに定期代が不正です(${pass})`);

    const storeId = str(r[col.store]);
    if (!storeId) throw new ConvertError(`${emp}: 所属コードが空欄です`);
    if (!byEmp.has(emp)) byEmp.set(emp, []);
    const empRows = byEmp.get(emp);
    if (empRows.some((x) => x[4] === storeId)) duplicates.push(`${emp}(${storeId})`);

    empRows.push([emp, str(r[col.name]) || null, str(r[col.kana]) || null,
      String(fare), storeId, '使用する', String(pass)]);
  }
  if (duplicates.length) {
    throw new ConvertError('同じ店舗の行が重複している従業員がいます', duplicates);
  }
  if (!byEmp.size) throw new ConvertError('往復運賃が入っている行がありません');

  const outRows = [...byEmp.values()].flat();
  const multi = [...byEmp.values()].filter((x) => x.length > 1).length;
  return { outRows, empCount: byEmp.size, stats: { '複数店舗': multi } };
}

const COMMUTE = {
  key: 'commute',
  label: '通勤手当',
  formatVersion: '1.1.0',
  header: ['社員コード', '社員名', '社員名(カナ)', '通勤手当', '対象部署/店舗ID',
    '日額通勤手当の月上限', '日額通勤手当の月上限額'],
  previewLabels: { 2: 'カナ', 4: '店舗ID', 5: '月上限', 6: '月上限額' },
  colWidths: [14, 10, 12, 10, 16, 20, 22],
  filePrefix: '通勤手当インポートデータ',
  // 複数店舗(複数行)の人の行はプレビューで色を変える
  highlighter: (outRows) => {
    const counts = new Map();
    for (const [emp] of outRows) counts.set(emp, (counts.get(emp) || 0) + 1);
    return (row) => counts.get(row[0]) > 1;
  },
  highlightLabel: '複数店舗のみ',
  convert: convertCommute,
};
