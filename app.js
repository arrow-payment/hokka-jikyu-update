'use strict';

// 元データ・ARROWのエクスポートはブラウザのメモリ上でだけ処理し、外部送信も保存(localStorage等)もしない
// 1つの元データから、タブごとに各インポート用データを作る

const CONVERTERS = [HOURLY, COMMUTE, PAYROLL, LABOR];
const $ = (id) => document.getElementById(id);

// 検索用に、元データの社員番号 → 「社員番号 氏名 フリガナ」を持つ。
// 空白を除き、全角英数・半角カナ・ひらがな・小文字の違いは無視して探せるようにする
let searchKeys = new Map();
const toSearchKey = (v) => str(v).normalize('NFKC').replace(/\s/g, '')
  .replace(/[ぁ-ゖ]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0x60)).toUpperCase();

function buildSearchKeys(sheetRows) {
  const names = (sheetRows[0] || []).map(normalizeHeader);
  const [emp, name, kana] = ['社員番号', '氏名', '氏名(フリガナ)'].map((h) => names.indexOf(h));
  const keys = new Map();
  if (emp < 0) return keys;
  for (const r of sheetRows.slice(1)) {
    const code = str(r[emp]);
    if (code && !keys.has(code)) keys.set(code, toSearchKey([code, r[name], r[kana]].join('')));
  }
  return keys;
}

function fillMessage(box, { message, details }) {
  box.querySelector('.message').textContent = message;
  box.querySelector('ul').replaceChildren(...(details || []).map((d) => {
    const li = document.createElement('li');
    li.textContent = d;
    return li;
  }));
  box.hidden = false;
}

const toConvertError = (err) => (err instanceof ConvertError ? err : new Error('変換できませんでした'));

function setupDropZone(zone, onFile) {
  const input = zone.querySelector('input');
  input.addEventListener('change', () => {
    onFile(input.files[0]);
    input.value = '';
  });
  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('over');
    onFile(e.dataTransfer.files[0]);
  });
}

function renderPreview(view, conv, result) {
  const { outRows } = result;
  const header = result.header || conv.header;
  const head = document.createElement('tr');
  for (const [i, name] of header.entries()) {
    const th = document.createElement('th');
    th.textContent = conv.previewLabels[i] || name;
    head.append(th);
  }
  view.querySelector('thead').replaceChildren(head);
  // 出力の社員名・カナ列も検索対象にする(元データで名前が引けない場合にも探せるように)
  const nameCols = header.map(normalizeHeader).flatMap((h, i) => (/^(社員名|氏名)/.test(h) ? [i] : []));

  // 従業員ごとに背景を交互にし、highlight 対象の行は色を変える。
  // チェックボックスの絞り込みは従業員単位(対象行が1つでもある人の全行を残す)
  const highlight = conv.highlighter(outRows, result);
  const flaggedEmps = new Set(outRows.filter(highlight).map((r) => r[0]));
  let band = false;
  const body = document.createDocumentFragment();
  outRows.forEach((row, n) => {
    const tr = document.createElement('tr');
    if (n > 0 && row[0] !== outRows[n - 1][0]) {
      band = !band;
      tr.classList.add('emp-start');
    }
    tr.dataset.emp = row[0];
    tr.dataset.search = (searchKeys.get(str(row[0])) ?? '') + toSearchKey([row[0], ...nameCols.map((i) => row[i])].join(''));
    if (flaggedEmps.has(row[0])) tr.dataset.flagged = '';
    if (band) tr.classList.add('band');
    if (highlight(row)) tr.classList.add('exception');
    const changedCols = result.changed?.get(row);
    row.forEach((v, i) => {
      const td = document.createElement('td');
      td.textContent = v ?? '';
      if (changedCols?.has(i)) td.classList.add('changed');
      tr.append(td);
    });
    body.append(tr);
  });
  view.querySelector('tbody').replaceChildren(body);
  return flaggedEmps;
}

// 更新項目のチェックボックスを作る。変更のたびに onChange で作り直す
function renderFields(view, conv, selected, onChange) {
  if (!conv.fields) return;
  view.querySelector('.field-list').replaceChildren(...conv.fields.map(({ key, label }) => {
    const box = document.createElement('label');
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = selected.has(key);
    input.addEventListener('change', () => {
      if (input.checked) selected.add(key);
      else selected.delete(key);
      onChange();
    });
    const span = document.createElement('span');
    span.textContent = label;
    box.append(input, span);
    return box;
  }));
  view.querySelector('.fields').hidden = false;
}

// convert を実行し、結果(またはエラー)の表示を panel の .view に作り直す。
// conv.fields がある場合は selected(選択中の更新項目の key の集合)を convert に渡す
function renderResult(panel, conv, convert, selected) {
  const prev = panel.querySelector('.view');
  const view = $('result-template').content.firstElementChild.cloneNode(true);
  prev.replaceWith(view);
  try {
    const result = convert(selected);
    const { outRows, empCount, stats, notice } = result;
    view.querySelector('.emp-count').textContent = `${empCount}名`;
    view.querySelector('.out-count').textContent = `${outRows.length}行`;
    view.querySelector('.breakdown').textContent = Object.entries(stats)
      .filter(([, v]) => v)
      .map(([k, v]) => `${k} ${v}名`)
      .join(' / ');
    if (notice) fillMessage(view.querySelector('.notice'), notice);
    const flaggedEmps = renderPreview(view, conv, result);

    renderFields(view, conv, selected, () => renderResult(panel, conv, convert, selected));

    // 項目の選択で作り直したときも絞り込み条件は引き継ぐ
    const filter = view.querySelector('.filter');
    const onlyFlagged = view.querySelector('.only-flagged input');
    filter.value = prev.querySelector('.filter')?.value ?? '';
    onlyFlagged.checked = prev.querySelector('.only-flagged input')?.checked ?? false;
    view.querySelector('.only-flagged span').textContent = conv.highlightLabel;
    // downloadFlaggedOnly のタブは、チェック中は対象の従業員だけをダウンロードする
    const download = view.querySelector('.download');
    const flaggedDownload = () => conv.downloadFlaggedOnly && onlyFlagged.checked;
    const applyFilter = () => {
      const q = toSearchKey(filter.value);
      for (const tr of view.querySelector('tbody').rows) {
        tr.hidden = (q !== '' && !tr.dataset.search.includes(q))
          || (onlyFlagged.checked && !('flagged' in tr.dataset));
      }
      download.textContent = flaggedDownload() ? `${conv.highlightLabel}をダウンロード(${flaggedEmps.size}名)` : 'ダウンロード';
      download.disabled = flaggedDownload() && !flaggedEmps.size;
    };
    applyFilter();
    filter.addEventListener('input', applyFilter);
    onlyFlagged.addEventListener('change', applyFilter);
    download.addEventListener('click', () => {
      const rows = flaggedDownload() ? outRows.filter((r) => flaggedEmps.has(r[0])) : outRows;
      XLSX.writeFile(buildWorkbook(XLSX, { ...conv, ...result, outRows: rows }), outputFileName(conv.filePrefix));
    });
  } catch (err) {
    console.error(err);
    view.querySelector('.result').hidden = true;
    fillMessage(view.querySelector('.error'), toConvertError(err));
  }
}

function buildPanel(conv, sheetRows) {
  const panel = $('panel-template').content.firstElementChild.cloneNode(true);
  panel.id = `panel-${conv.key}`;
  if (!conv.current) {
    renderResult(panel, conv, () => conv.convert(sheetRows));
    return panel;
  }

  // 現在の登録状況(ARROWのエクスポート)と突き合わせるタブは、追加でそのファイルを受け取る
  const selected = new Set(conv.fields?.map((f) => f.key));
  const zone = panel.querySelector('.current');
  zone.hidden = false;
  zone.querySelector('.drop-main').textContent = conv.current.label;
  setupDropZone(zone, async (file) => {
    if (!file) return;
    const data = await file.arrayBuffer();
    zone.querySelector('.drop-sub').textContent = file.name;
    // 読み取りエラーも renderResult で表示するよう、初回の convert で読む
    let currentRows;
    renderResult(panel, conv, (sel) => conv.convert(sheetRows, currentRows ??= readSheetRows(XLSX, data), sel), selected);
  });
  return panel;
}

function selectTab(key) {
  for (const tab of $('tabs').children) tab.setAttribute('aria-selected', tab.dataset.key === key);
  for (const panel of $('panels').children) panel.hidden = panel.id !== `panel-${key}`;
}

function reset() {
  $('outputs').hidden = true;
  $('error').hidden = true;
  $('tabs').replaceChildren();
  $('panels').replaceChildren();
}

async function handleFile(file) {
  reset();
  if (!file) return;
  try {
    const sheetRows = readSourceRows(XLSX, await file.arrayBuffer());
    searchKeys = buildSearchKeys(sheetRows);
    for (const conv of CONVERTERS) {
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.role = 'tab';
      tab.className = 'tab';
      tab.dataset.key = conv.key;
      tab.textContent = conv.label;
      tab.addEventListener('click', () => selectTab(conv.key));
      $('tabs').append(tab);
      $('panels').append(buildPanel(conv, sheetRows));
    }
    $('result-file').textContent = file.name;
    selectTab(CONVERTERS[0].key);
    $('outputs').hidden = false;
  } catch (err) {
    console.error(err);
    fillMessage($('error'), toConvertError(err));
  }
}

setupDropZone($('drop'), handleFile);
