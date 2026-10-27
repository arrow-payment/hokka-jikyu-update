'use strict';

// 元データ・ARROWのエクスポートはブラウザのメモリ上でだけ処理し、外部送信も保存(localStorage等)もしない
// 1つの元データから、タブごとに各インポート用データを作る

const CONVERTERS = [HOURLY, COMMUTE, PAYROLL];
const $ = (id) => document.getElementById(id);

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
}

// convert を実行し、結果(またはエラー)の表示を panel の .view に作り直す
function renderResult(panel, conv, convert) {
  const view = $('result-template').content.firstElementChild.cloneNode(true);
  panel.querySelector('.view').replaceWith(view);
  try {
    const result = convert();
    const { outRows, empCount, stats, notice } = result;
    view.querySelector('.emp-count').textContent = `${empCount}名`;
    view.querySelector('.out-count').textContent = `${outRows.length}行`;
    view.querySelector('.breakdown').textContent = Object.entries(stats)
      .filter(([, v]) => v)
      .map(([k, v]) => `${k} ${v}名`)
      .join(' / ');
    if (notice) fillMessage(view.querySelector('.notice'), notice);
    renderPreview(view, conv, result);

    const filter = view.querySelector('.filter');
    const onlyFlagged = view.querySelector('.only-flagged input');
    view.querySelector('.only-flagged span').textContent = conv.highlightLabel;
    const applyFilter = () => {
      const q = filter.value.trim().toUpperCase();
      for (const tr of view.querySelector('tbody').rows) {
        tr.hidden = (q !== '' && !tr.dataset.emp.toUpperCase().includes(q))
          || (onlyFlagged.checked && !('flagged' in tr.dataset));
      }
    };
    filter.addEventListener('input', applyFilter);
    onlyFlagged.addEventListener('change', applyFilter);
    view.querySelector('.download').addEventListener('click', () => {
      XLSX.writeFile(buildWorkbook(XLSX, { ...conv, ...result }), outputFileName(conv.filePrefix));
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
  const zone = panel.querySelector('.current');
  zone.hidden = false;
  zone.querySelector('.drop-main').textContent = conv.current.label;
  setupDropZone(zone, async (file) => {
    if (!file) return;
    const data = await file.arrayBuffer();
    zone.querySelector('.drop-sub').textContent = file.name;
    renderResult(panel, conv, () => conv.convert(sheetRows, readSheetRows(XLSX, data)));
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
