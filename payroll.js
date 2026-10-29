'use strict';

// ほっか直営店 給与設定データ変換ロジック
//
// ARROWからエクスポートした現在の給与設定(アウトプットサンプル_給与設定.xlsx と同じ形式)を
// ベースに、元データで次の列だけを上書きする。それ以外の列は現在の登録内容のまま出す。
//
// - 通勤手当(元データの往復運賃が1店舗でもある人のみ。ない人は現在の登録のまま):
//   通勤手当=日額 / 通勤方法=公共交通機関のみ / 通勤手当の金額=0 / 日額通勤手当の月上限=使用しない
//   (金額は通勤手当インポートで店舗ごとの例外通勤手当として設定するため0にする)
// - 有給休暇全日使用単価 = 基本給 × 1日の勤務時間(四捨五入)。単価を入れるため計算方法は「手動」にする
// - 個人別単価1 = 宅配手当、個人別単価2 = 残業2(資格手当による増額分)
// - 複数店舗の人は親店舗の行の値を使う(往復運賃のみ全店舗を見る)
// - ARROWにいて元データにいない人は出力しない。元データにいてARROWにいない人は
//   上書き先がないので出力できず、一覧で知らせる
// - 1行目の __format_version と2行目の見出しはエクスポートのものをそのまま使う
// - 上書きは PAYROLL_FIELDS の項目単位で選べる。選ばなかった項目は現在の登録のまま出す

const PAYROLL_SOURCE_COLUMNS = {
  emp: '社員番号',
  parent: '親店舗',
  child: '子店舗',
  base: '基本給',
  hours: '勤務時間(開始から終了までのうち、1日何時間か)',
  delivery: '宅配手当',
  qualOvertime: '残業2(資格手当による増額分)',
  fare: '往復運賃',
};
const PAYROLL_CURRENT_COLUMNS = {
  emp: '社員コード',
  commute: '通勤手当',
  commuteMethod: '通勤方法',
  commuteAmount: '通勤手当の金額',
  commuteCap: '日額通勤手当の月上限',
  paidLeaveMethod: '有給休暇全日使用単価の計算方法',
  paidLeave: '有給休暇全日使用単価',
  personal1: '個人別単価1',
  personal2: '個人別単価2',
};

// 画面で選べる更新項目。cols は PAYROLL_CURRENT_COLUMNS のキー
const PAYROLL_FIELDS = [
  { key: 'commute', label: '通勤手当', cols: ['commute', 'commuteMethod', 'commuteAmount', 'commuteCap'] },
  { key: 'paidLeave', label: '有給休暇全日使用単価', cols: ['paidLeaveMethod', 'paidLeave'] },
  { key: 'personal1', label: '個人別単価1(宅配手当)', cols: ['personal1'] },
  { key: 'personal2', label: '個人別単価2(残業2)', cols: ['personal2'] },
];

// 「04:30:00」「4:30」→ 分。Excelの時刻セル(1日=1の小数)にも対応
function parseMinutes(v) {
  if (typeof v === 'number') return Math.round(v < 1 ? v * 24 * 60 : v * 60);
  const m = str(v).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

const amount = (v) => (v === null || v === '' ? 0 : v);

// sourceRows: 元データシートの全行 / currentRows: ARROWのエクスポートの全行(1行目が __format_version)
// selected: 上書きする PAYROLL_FIELDS の key の集合(省略時は全項目)
function convertPayroll(sourceRows, currentRows, selected = new Set(PAYROLL_FIELDS.map((f) => f.key))) {
  const selectedCols = new Set(PAYROLL_FIELDS.filter((f) => selected.has(f.key)).flatMap((f) => f.cols));
  const src = findColumns(sourceRows[0], PAYROLL_SOURCE_COLUMNS);
  const byEmp = groupByEmployee(dataRows(sourceRows, src), src);

  if (str(currentRows[0]?.[0]) !== '__format_version') {
    throw new ConvertError('ARROWの給与設定エクスポートではないようです');
  }
  const header = currentRows[1] || [];
  const cur = findColumns(header, PAYROLL_CURRENT_COLUMNS);

  const outRows = [];
  // 出力行 → 現在の登録から値が変わる列index の集合
  const changed = new Map();
  const seen = new Set();
  const duplicates = [];
  const stats = { '通勤手当あり': 0, '宅配手当あり': 0, '資格手当の残業増額あり': 0 };
  for (const current of currentRows.slice(2)) {
    const emp = str(current[cur.emp]);
    if (!emp || !byEmp.has(emp)) continue;
    if (seen.has(emp)) duplicates.push(emp);
    seen.add(emp);

    const { main, rows } = byEmp.get(emp);
    const minutes = parseMinutes(main[src.hours]);
    const base = main[src.base];
    if (typeof base !== 'number' || base <= 0) throw new ConvertError(`${emp}: 基本給が不正です(${base})`);
    if (minutes === null) throw new ConvertError(`${emp}: 1日の勤務時間が読めません(${main[src.hours]})`);
    const delivery = amount(main[src.delivery]);
    const qualOvertime = amount(main[src.qualOvertime]);
    if (typeof delivery !== 'number') throw new ConvertError(`${emp}: 宅配手当が数値ではありません(${delivery})`);
    if (typeof qualOvertime !== 'number') throw new ConvertError(`${emp}: 資格手当の残業増額が数値ではありません(${qualOvertime})`);

    const updates = {
      paidLeaveMethod: '手動',
      paidLeave: String(Math.round((base * minutes) / 60)),
      personal1: String(delivery),
      personal2: String(qualOvertime),
    };
    if (rows.some((r) => typeof r[src.fare] === 'number' && r[src.fare] > 0)) {
      Object.assign(updates, {
        commute: '日額',
        commuteMethod: '公共交通機関のみ',
        commuteAmount: '0',
        commuteCap: '使用しない',
      });
      stats['通勤手当あり']++;
    }
    if (delivery) stats['宅配手当あり']++;
    if (qualOvertime) stats['資格手当の残業増額あり']++;

    const row = header.map((_, i) => current[i] ?? null);
    const changedCols = new Set();
    for (const [key, value] of Object.entries(updates)) {
      if (!selectedCols.has(key)) continue;
      const i = cur[key];
      if (str(row[i]) !== value) changedCols.add(i);
      row[i] = value;
    }
    if (changedCols.size) changed.set(row, changedCols);
    outRows.push(row);
  }
  if (duplicates.length) {
    throw new ConvertError('ARROWの給与設定に同じ社員コードが複数あります', duplicates);
  }
  if (!outRows.length) throw new ConvertError('元データとARROWの給与設定で一致する社員コードがありません');

  const missing = [...byEmp.keys()].filter((e) => !seen.has(e));
  return {
    outRows,
    empCount: outRows.length,
    stats,
    changed,
    header,
    formatVersion: str(currentRows[0][1]),
    colWidths: header.map((h) => Math.max(8, str(h).length * 2)),
    notice: missing.length
      ? { message: `ARROWに登録がないため出力できない従業員がいます(${missing.length}名)`, details: missing }
      : null,
  };
}

const PAYROLL = {
  key: 'payroll',
  label: '給与設定',
  current: { label: 'ARROWの給与設定を選択' },
  previewLabels: {},
  filePrefix: '給与設定インポートデータ',
  // 現在の登録から値が変わる行の色を変える
  highlighter: (outRows, result) => (row) => result.changed.has(row),
  highlightLabel: '変更ありのみ',
  fields: PAYROLL_FIELDS,
  convert: convertPayroll,
};
