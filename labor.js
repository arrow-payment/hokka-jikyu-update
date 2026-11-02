'use strict';

// ほっか直営店 労働条件(週所定労働日数・時間)データ変換ロジック
//
// ARROWからエクスポートした現在の労働条件(employee_employee_master_labor_general と同じ形式)を
// ベースに、元データで次の列だけを上書きする。それ以外の列は現在の登録内容のまま出す。
//
// - 週所定労働日数 = 勤務曜日の曜日数(「日・祝」のような祝は数えない)
// - 週所定労働時間 = 曜日数 × 1日の勤務時間。「6:00」の形式で出す
// - 複数店舗の人は親店舗の行の値を使う。子店舗の行にも同じ勤務曜日・時間が入っているため
//   合算しない(ユーザー確認済み)
// - 勤務曜日が読めない・空欄・0の人、1日の勤務時間が読めない・0・24時間超の人は
//   出力せず、一覧で知らせる
// - ARROWにいて元データにいない人は出力しない。元データにいてARROWにいない人は
//   上書き先がないので出力できず、一覧で知らせる
// - 1行目の __format_version と2行目の見出しはエクスポートのものをそのまま使う
// - 上書きは LABOR_FIELDS の項目単位で選べる。選ばなかった項目は現在の登録のまま出す

const LABOR_SOURCE_COLUMNS = {
  emp: '社員番号',
  parent: '親店舗',
  child: '子店舗',
  hours: '勤務時間(開始から終了までのうち、1日何時間か)',
  days: '勤務曜日[勤務曜日]',
};
const LABOR_CURRENT_COLUMNS = {
  emp: '社員コード',
  weekDays: '週所定労働日数',
  weekHours: '週所定労働時間',
};
const LABOR_FIELDS = [
  { key: 'weekDays', label: '週所定労働日数', cols: ['weekDays'] },
  { key: 'weekHours', label: '週所定労働時間', cols: ['weekHours'] },
];
const WEEKDAYS = ['月', '火', '水', '木', '金', '土', '日'];

// 「月, 水, 土, 日・祝」→ 曜日の集合。読めない表記があれば null
function parseWeekdays(v) {
  const days = new Set();
  for (const t of str(v).split(/[,、，・\s]+/).filter(Boolean)) {
    if (t === '祝') continue;
    if (!WEEKDAYS.includes(t)) return null;
    days.add(t);
  }
  return days;
}

const formatMinutes = (m) => `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')}`;

// sourceRows: 元データシートの全行 / currentRows: ARROWのエクスポートの全行(1行目が __format_version)
// selected: 上書きする LABOR_FIELDS の key の集合(省略時は全項目)
function convertLabor(sourceRows, currentRows, selected = new Set(LABOR_FIELDS.map((f) => f.key))) {
  const selectedCols = new Set(LABOR_FIELDS.filter((f) => selected.has(f.key)).flatMap((f) => f.cols));
  const src = findColumns(sourceRows[0], LABOR_SOURCE_COLUMNS);
  const byEmp = groupByEmployee(dataRows(sourceRows, src), src);

  if (str(currentRows[0]?.[0]) !== '__format_version') {
    throw new ConvertError('ARROWの労働条件エクスポートではないようです');
  }
  const header = currentRows[1] || [];
  const cur = findColumns(header, LABOR_CURRENT_COLUMNS);

  const outRows = [];
  // 出力行 → 現在の登録から値が変わる列index の集合
  const changed = new Map();
  const seen = new Set();
  const duplicates = [];
  const skipped = [];
  const stats = { '複数店舗': 0 };
  for (const current of currentRows.slice(2)) {
    const emp = str(current[cur.emp]);
    if (!emp || !byEmp.has(emp)) continue;
    if (seen.has(emp)) duplicates.push(emp);
    seen.add(emp);

    const { main, rows } = byEmp.get(emp);
    const days = parseWeekdays(main[src.days]);
    const minutes = parseMinutes(main[src.hours]);
    if (!days?.size || !minutes) {
      skipped.push(`${emp}(勤務曜日・勤務時間が空欄か0)`);
      continue;
    }
    if (minutes > 24 * 60) {
      skipped.push(`${emp}(1日の勤務時間が24時間超: ${str(main[src.hours])})`);
      continue;
    }
    const weekMinutes = days.size * minutes;
    if (rows.length > 1) stats['複数店舗']++;

    const updates = {
      weekDays: String(days.size),
      weekHours: formatMinutes(weekMinutes),
    };
    // 「06:00」と「6:00」のような表記の違いは変更とみなさない
    const same = {
      weekDays: (v) => Number(str(v)) === days.size,
      weekHours: (v) => parseMinutes(v) === weekMinutes,
    };

    const row = header.map((_, i) => current[i] ?? null);
    const changedCols = new Set();
    for (const [key, value] of Object.entries(updates)) {
      if (!selectedCols.has(key)) continue;
      const i = cur[key];
      if (!same[key](row[i])) changedCols.add(i);
      row[i] = value;
    }
    if (changedCols.size) changed.set(row, changedCols);
    outRows.push(row);
  }
  if (duplicates.length) {
    throw new ConvertError('ARROWの労働条件に同じ社員コードが複数あります', duplicates);
  }
  if (!outRows.length) throw new ConvertError('元データとARROWの労働条件で一致し、出力できる従業員がいません');

  const missing = [...byEmp.keys()].filter((e) => !seen.has(e));
  const details = [
    ...skipped,
    ...missing.map((e) => `${e}(ARROWに登録なし)`),
  ];
  return {
    outRows,
    empCount: outRows.length,
    stats,
    changed,
    header,
    formatVersion: str(currentRows[0][1]),
    colWidths: header.map((h) => Math.max(8, str(h).length * 2)),
    notice: details.length
      ? { message: `出力しなかった従業員がいます(${details.length}名)`, details }
      : null,
  };
}

const LABOR = {
  key: 'labor',
  label: '週所定労働',
  current: { label: 'ARROWの労働条件を選択' },
  previewLabels: {},
  filePrefix: '週所定労働インポートデータ',
  // 現在の登録から値が変わる行の色を変える
  highlighter: (outRows, result) => (row) => result.changed.has(row),
  highlightLabel: '変更ありのみ',
  downloadFlaggedOnly: true,
  fields: LABOR_FIELDS,
  convert: convertLabor,
};
