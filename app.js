(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const input = $('file-input'), drop = $('dropzone'), status = $('status');
  let workbook, sourceFile, pdfUrl, busy = false;
  const extensions = ['xlsx','xls','xlsm','xlsb','csv','tsv','ods','fods','txt','xml','html','htm','numbers','slk','dif'];
  function message(text, type = '') { status.textContent = text; status.className = 'status ' + type; status.hidden = !text; }
  function setBusy(value) { busy = value; input.disabled = value; $('sheet-select').disabled = value; $('clear-button').disabled = value; drop.setAttribute('aria-disabled', String(value)); }
  function clearPdf() {
    if (pdfUrl) URL.revokeObjectURL(pdfUrl);
    pdfUrl = null; $('download').removeAttribute('href'); $('result').hidden = true; $('preview-body').replaceChildren();
  }
  function reset() {
    clearPdf(); workbook = null; sourceFile = null; input.value = ''; $('sheet-picker').hidden = true; $('sheet-select').replaceChildren(); $('file-info').hidden = true; message('');
  }
  function render(report) {
    const tbody = $('preview-body');
    const fragment = document.createDocumentFragment();
    for (const row of report.rows.slice(0, 100)) {
      const tr = document.createElement('tr');
      [row.code, row.description, row.quantityText, row.color, row.warehouse].forEach((value, i) => {
        const td = document.createElement('td'); td.textContent = value || '—'; if (i === 2) td.className = 'numeric'; tr.append(td);
      });
      fragment.append(tr);
    }
    tbody.replaceChildren(fragment);
    $('result-summary').textContent = `${StockPDF.numberFormat.format(report.rows.length)} registros · 5 colunas · Aba ${report.sheetName}`;
    $('preview-note').textContent = report.rows.length > 100 ? 'Prévia dos primeiros 100 registros. O PDF contém todos os ' + StockPDF.numberFormat.format(report.rows.length) + ' registros.' : 'Prévia dos dados incluídos no PDF. A ordem dos produtos foi preservada.';
    $('warning').textContent = report.warnings.join(' '); $('warning').hidden = !report.warnings.length;
    $('result').hidden = false;
  }
  async function generate(autoDownload = true) {
    clearPdf();
    setBusy(true); message('Organizando seu estoque e preparando o PDF…', 'busy');
    // Permite que o navegador desenhe o estado de processamento antes da conversão.
    await new Promise(resolve => setTimeout(resolve, 50));
    try {
      const report = StockPDF.extract(workbook, $('sheet-select').value);
      const doc = StockPDF.createPdf(report, sourceFile.name);
      pdfUrl = URL.createObjectURL(doc.output('blob'));
      const download = $('download');
      download.href = pdfUrl;
      const base = sourceFile.name.replace(/\.[^.]+$/, '').replace(/[^\p{L}\p{N}_-]+/gu, '-').slice(0, 70) || 'produtos';
      download.download = `estoque-${base}.pdf`;
      render(report);
      const cleaned = report.cleanup.blankRows + report.cleanup.repeatedHeaders + report.cleanup.totalRows;
      message(`PDF pronto! ${StockPDF.numberFormat.format(report.rows.length)} registros organizados.${cleaned ? ' Linhas vazias, cabeçalhos repetidos e totais de rodapé foram desconsiderados quando identificados.' : ''} Se o download não começar, use “Baixar PDF”.`);
      if (autoDownload) download.click();
    } catch (error) {
      clearPdf(); message(error.message || 'Não foi possível gerar o PDF. Verifique o arquivo e tente novamente.', 'error');
    } finally { setBusy(false); }
  }
  async function openFile(file) {
    if (!file || busy) return;
    reset();
    if (!window.XLSX || !window.jspdf || !window.StockPDF || !window.jspdf.jsPDF.API.autoTable) { message('Não foi possível carregar os componentes do site. Recarregue a página e tente novamente.', 'error'); return; }
    if (!extensions.includes(file.name.toLowerCase().split('.').pop())) { message('Esse formato não é uma planilha compatível. Envie Excel (XLSX, XLS, XLSM ou XLSB), CSV, TSV, ODS ou outro formato listado no seletor.', 'error'); return; }
    if (file.size > 50 * 1024 * 1024) { message('O arquivo excede 50 MB. Exporte o estoque em um arquivo menor e tente novamente.', 'error'); return; }
    setBusy(true); message('Lendo sua planilha…', 'busy');
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      if (!data.length) throw new Error('O arquivo está vazio. Selecione uma planilha com os produtos.');
      workbook = StockPDF.parseWorkbook(data, file.name);
      const sheets = StockPDF.inspectWorkbook(workbook);
      if (!sheets.length) throw new Error('Não encontrei uma aba de estoque com Código, Descrição, Qtde, Descor e DescDep nas colunas A a E. O arquivo pode estar vazio, protegido por senha ou ter uma estrutura diferente da exportação original.');
      sourceFile = file;
      $('file-name').textContent = file.name;
      $('file-details').textContent = `${new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1 }).format(file.size / 1024)} KB · ${sheets.length} aba(s) de estoque identificada(s)`;
      $('file-info').hidden = false;
      $('sheet-select').replaceChildren(...sheets.map(sheet => { const option = document.createElement('option'); option.value = sheet.name; option.textContent = sheet.name; return option; }));
      // Uma única aba identificada gera o PDF imediatamente; várias exigem escolha para não omitir estoques silenciosamente.
      $('sheet-picker').hidden = sheets.length < 2;
      if (sheets.length > 1) {
        message('Encontrei mais de uma aba de estoque. Selecione a aba desejada para gerar seu PDF.');
        const placeholder = document.createElement('option'); placeholder.value = ''; placeholder.textContent = 'Selecione uma aba'; placeholder.disabled = true; placeholder.selected = true; $('sheet-select').prepend(placeholder);
        setBusy(false); return;
      }
    } catch (error) {
      workbook = null; sourceFile = null;
      message(error.message?.startsWith('Não encontrei') || error.message?.startsWith('O arquivo') ? error.message : 'Não consegui ler essa planilha. Verifique se o arquivo abre no Excel ou LibreOffice e se não está protegido por senha.', 'error');
      setBusy(false); return;
    }
    setBusy(false); await generate();
  }
  input.addEventListener('change', () => openFile(input.files[0]));
  $('clear-button').addEventListener('click', () => { reset(); input.focus(); });
  $('sheet-select').addEventListener('change', () => generate());
  for (const event of ['dragenter', 'dragover']) drop.addEventListener(event, e => { e.preventDefault(); if (!busy) drop.classList.add('dragover'); });
  for (const event of ['dragleave', 'drop']) drop.addEventListener(event, e => { e.preventDefault(); drop.classList.remove('dragover'); });
  drop.addEventListener('drop', e => {
    if (e.dataTransfer.files.length > 1) { message('Envie uma planilha por vez para gerar o PDF correspondente.', 'error'); return; }
    openFile(e.dataTransfer.files[0]);
  });
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => e.preventDefault());
  window.addEventListener('beforeunload', () => { if (pdfUrl) URL.revokeObjectURL(pdfUrl); });
})();
