(function () {
  const els = {
    tabs: Array.from(document.querySelectorAll('.tab')),
    docKindLabel: document.getElementById('doc-kind-label'),
    date: document.getElementById('doc-date'),
    role: document.getElementById('job-role'),
    fio: document.getElementById('fio'),
    fioGen: document.getElementById('fio-gen'),
    product: document.getElementById('product'),
    article: document.getElementById('article'),
    quantity: document.getElementById('quantity'),
    eo: document.getElementById('eo'),
    author: document.getElementById('author'),
    authorRole: document.getElementById('author-role'),
    formExp: document.getElementById('form-exp'),
    expOp: document.getElementById('exp-op'),
    expIssue: document.getElementById('exp-issue'),
    expReason1: document.getElementById('exp-reason-1'),
    expReason2: document.getElementById('exp-reason-2'),
    expMeasures: document.getElementById('exp-measures'),
    output: document.getElementById('output'),
    btnGenerate: document.getElementById('btn-generate'),
    btnClear: document.getElementById('btn-clear'),
    btnCopy: document.getElementById('btn-copy'),
    btnPrint: document.getElementById('btn-print'),
    btnDownload: document.getElementById('btn-download'),
  };

  const DI = {
    receiving: {
      label: 'Кладовщик (участок приема)',
      duty: 'раздел 3 ДИ, п. 3.1.1 (приемка ТМЦ, работа в ТСД/учетных системах, контроль корректности операций), раздел 3.1.3 (отчетность)',
      resp: 'раздел 5 ДИ (ответственность за ненадлежащее исполнение обязанностей, последствия ошибок, материальный ущерб)',
    },
    placement: {
      label: 'Кладовщик (участок размещения)',
      duty: 'раздел 3 ДИ, п. 3.1.1 (размещение ТМЦ, работа в ТСД/учетных системах, корректное оформление операций), раздел 3.1.3 (отчетность)',
      resp: 'раздел 5 ДИ (персональная ответственность за последствия решений и ошибок в операциях)',
    },
    forklift: {
      label: 'Водитель погрузчика (участок размещения)',
      duty: 'раздел 3 ДИ, п. 3.1.1 и 3.1.2 (выполнение работ ПРТ и погрузо-разгрузочных операций по установленным правилам)',
      resp: 'раздел 5 ДИ (ответственность за нарушения требований, причиненный ущерб и последствия решений)',
    },
  };

  const DOC_KIND = {
    bidu: 'Служебная (BIDU)',
    surplus: 'Служебная (Излишки)',
    exp: 'Объяснительная',
  };

  let kind = 'bidu';
  let generatedText = '';

  function setToday() {
    const d = new Date();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    els.date.value = `${d.getFullYear()}-${m}-${day}`;
  }

  function fmtDate(v) {
    if (!v) return '___ . ___ . ______';
    const [y, m, d] = v.split('-');
    return `${d}.${m}.${y}`;
  }

  function nonEmpty(v, fallback) {
    return (v || '').trim() || fallback;
  }

  function ruPlural(n, one, few, many) {
    const num = Math.abs(Number(n));
    if (!Number.isFinite(num)) return many;
    const mod10 = num % 10;
    const mod100 = num % 100;
    if (mod10 === 1 && mod100 !== 11) return one;
    if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) return few;
    return many;
  }

  function qtyWithWord(raw, one, few, many) {
    const text = String(raw || '').trim().replace(',', '.');
    if (!text) return '________________';
    const n = Number(text);
    if (!Number.isFinite(n)) return text;
    return `${text} ${ruPlural(n, one, few, many)}`;
  }

  function fioNominative() {
    return nonEmpty(els.fio.value, '________________');
  }

  function fioGenitive() {
    const g = (els.fioGen.value || '').trim();
    return g || fioNominative();
  }

  function buildHeader(title, subtitle) {
    return [
      'СЛУЖЕБНАЯ ЗАПИСКА',
      subtitle,
      '',
      `Дата: ${fmtDate(els.date.value)}`,
      '',
    ].join('\n');
  }

  function buildBidu() {
    const r = DI[els.role.value];
    return [
      buildHeader('СЛУЖЕБНАЯ ЗАПИСКА', 'О выявленных нарушениях в процессе работы'),
      `Настоящим сообщаю, что ${fmtDate(els.date.value)} у сотрудника ${fioGenitive()} выявлено нарушение:`,
      'некорректное применение кода BIDU.',
      '',
      `Товар: ${nonEmpty(els.product.value, '________________')}`,
      `Артикул: ${nonEmpty(els.article.value, '________________')}`,
      `Количество: ${qtyWithWord(els.quantity.value, 'единица', 'единицы', 'единиц')}`,
      `ЕО: ${nonEmpty(els.eo.value, '________________')}`,
      '',
      'Обоснование (ДИ):',
      `1. Нарушены обязанности: ${r.duty}.`,
      `2. Подлежит оценке ответственность: ${r.resp}.`,
      '',
      'Прошу:',
      '1. Запросить письменную объяснительную у сотрудника.',
      '2. Провести служебную проверку обстоятельств.',
      '3. Принять решение о мерах воздействия в соответствии с локальными актами и ТК РФ.',
      '',
      `Составил: ${nonEmpty(els.author.value, '________________')}`,
      `Должность: ${nonEmpty(els.authorRole.value, '________________')}`,
      'Подпись: __________________',
    ].join('\n');
  }

  function buildSurplus() {
    const r = DI[els.role.value];
    return [
      buildHeader('СЛУЖЕБНАЯ ЗАПИСКА', 'О выявленных нарушениях в процессе работы'),
      `Настоящим сообщаю, что ${fmtDate(els.date.value)} у сотрудника ${fioGenitive()} выявлено нарушение формирования отправления:`,
      'обнаружен излишек ТМЦ.',
      '',
      `Товар: ${nonEmpty(els.product.value, '________________')}`,
      `Артикул: ${nonEmpty(els.article.value, '________________')}`,
      `Излишек в количестве: ${qtyWithWord(els.quantity.value, 'единица', 'единицы', 'единиц')}`,
      `ЕО: ${nonEmpty(els.eo.value, '________________')}`,
      '',
      'Обоснование (ДИ):',
      `1. Нарушены обязанности: ${r.duty}.`,
      `2. Подлежит оценке ответственность: ${r.resp}, при наличии ущерба — с учетом ст. 243 ТК РФ.`,
      '',
      'Прошу:',
      '1. Запросить письменную объяснительную у сотрудника.',
      '2. Провести служебную проверку причин возникновения излишка.',
      '3. Принять корректирующие меры для исключения повторения.',
      '',
      `Составил: ${nonEmpty(els.author.value, '________________')}`,
      `Должность: ${nonEmpty(els.authorRole.value, '________________')}`,
      'Подпись: __________________',
    ].join('\n');
  }

  function buildExp() {
    const r = DI[els.role.value];
    const measures = (els.expMeasures.value || '')
      .split('\n')
      .map(s => s.trim())
      .filter(Boolean);
    const measuresBlock = measures.length
      ? measures.map((m, i) => `${i + 1}. ${m}`).join('\n')
      : '1. Усилить самоконтроль при выполнении операций.\n2. Проводить двойную сверку по ТСД.';

    return [
      'ОБЪЯСНИТЕЛЬНАЯ ЗАПИСКА',
      '',
      `Я, ${fioNominative()}, должность «${r.label}», по факту нарушения от ${fmtDate(els.date.value)} сообщаю следующее:`,
      '',
      `В ходе операции «${nonEmpty(els.expOp.value, '________________')}» по товару «${nonEmpty(els.product.value, '________________')}» (артикул ${nonEmpty(els.article.value, '________________')}, ЕО ${nonEmpty(els.eo.value, '________________')}) мной была допущена ошибка:`,
      `${nonEmpty(els.expIssue.value, '________________')}`,
      '',
      'Причины:',
      `1. ${nonEmpty(els.expReason1.value, '________________')}`,
      `2. ${nonEmpty(els.expReason2.value, '________________')}`,
      '',
      'Признаю, что нарушение относится к требованиям должностной инструкции:',
      `1. ${r.duty}.`,
      `2. ${r.resp}.`,
      '',
      'Для недопущения повторения обязуюсь:',
      measuresBlock,
      '',
      `Дата: ${fmtDate(els.date.value)}`,
      `Подпись: __________________ / ${fioNominative()}`,
    ].join('\n');
  }

  function render() {
    generatedText = kind === 'bidu'
      ? buildBidu()
      : kind === 'surplus'
        ? buildSurplus()
        : buildExp();
    renderPaper(generatedText);
  }

  function esc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function renderPaper(text) {
    const lines = String(text || '').split('\n');
    const blocks = [];
    let inList = false;
    let i = 0;
    while (i < lines.length && !lines[i].trim()) i++;
    if (i < lines.length) blocks.push(`<div class="doc-center">${esc(lines[i])}</div>`);
    i++;
    if (i < lines.length && lines[i].trim()) blocks.push(`<div class="doc-sub">${esc(lines[i])}</div>`);
    i++;

    for (; i < lines.length; i++) {
      const ln = lines[i];
      const t = ln.trim();
      if (!t) continue;
      if (inList && !/^\d+\.\s+/.test(t)) {
        blocks.push('</ol>');
        inList = false;
      }
      if (t.startsWith('Дата:')) {
        blocks.push(`<div class="doc-date">${esc(t)}</div>`);
        continue;
      }
      if (t === 'Прошу:' || t === 'Причины:' || t.startsWith('Обоснование')) {
        blocks.push(`<p class="doc-p no-indent"><b>${esc(t)}</b></p>`);
        continue;
      }
      if (/^\d+\.\s+/.test(t)) {
        if (!inList) {
          blocks.push('<ol class="doc-list">');
          inList = true;
        }
        blocks.push(`<li>${esc(t.replace(/^\d+\.\s+/, ''))}</li>`);
        continue;
      }
      if (t.startsWith('Составил:') || t.startsWith('Должность:') || t.startsWith('Подпись:')) {
        blocks.push(`<p class="doc-p no-indent doc-sign">${esc(t)}</p>`);
        continue;
      }
      blocks.push(`<p class="doc-p">${esc(t)}</p>`);
    }
    if (inList) blocks.push('</ol>');
    els.output.innerHTML = blocks.join('\n');
  }

  function switchKind(next) {
    kind = next;
    els.tabs.forEach(t => t.classList.toggle('active', t.dataset.kind === next));
    els.formExp.classList.toggle('hidden', next !== 'exp');
    els.docKindLabel.textContent = DOC_KIND[next];
    render();
  }

  function clearFields() {
    [
      els.fio, els.product, els.article, els.quantity, els.eo,
      els.fioGen, els.author, els.authorRole, els.expOp, els.expIssue, els.expReason1, els.expReason2, els.expMeasures,
    ].forEach(el => { el.value = ''; });
    setToday();
    render();
  }

  function copyOutput() {
    const text = generatedText.trim();
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      els.btnCopy.textContent = 'Скопировано';
      setTimeout(() => { els.btnCopy.textContent = 'Скопировать'; }, 1200);
    });
  }

  function downloadOutput() {
    const htmlBody = els.output.innerHTML || '';
    const htmlDoc = `<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Документ</title>
  <style>
    @page { size: A4; margin: 20mm; }
    body { font-family: "Times New Roman", serif; font-size: 12pt; line-height: 1.45; color: #000; }
    .doc-center { text-align: center; font-weight: 700; }
    .doc-sub { text-align: center; margin-top: 4px; }
    .doc-date { text-align: right; margin-top: 10px; margin-bottom: 14px; }
    .doc-p { text-align: justify; text-indent: 1.25cm; margin: 0 0 8px 0; }
    .doc-p.no-indent { text-indent: 0; }
    .doc-list { margin: 0 0 10px 0; padding-left: 20px; }
    .doc-list li { margin-bottom: 4px; }
    .doc-sign { margin-top: 18px; }
  </style>
</head>
<body>${htmlBody}</body>
</html>`;
    const blob = new Blob([htmlDoc], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `документ_${kind}_${els.date.value || 'без_даты'}.doc`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function printOutput() {
    window.focus();
    window.print();
  }

  els.tabs.forEach(tab => tab.addEventListener('click', () => switchKind(tab.dataset.kind)));
  els.btnGenerate.addEventListener('click', render);
  els.btnClear.addEventListener('click', clearFields);
  els.btnCopy.addEventListener('click', copyOutput);
  els.btnPrint.addEventListener('click', printOutput);
  els.btnDownload.addEventListener('click', downloadOutput);
  [els.role, els.date].forEach(el => el.addEventListener('change', render));

  setToday();
  render();
})();
