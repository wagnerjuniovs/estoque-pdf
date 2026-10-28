/* Leitura e geração independentes da interface. Nenhum arquivo é enviado à rede. */
(function (root) {
  'use strict';
  const LABELS = ['Código', 'Descrição', 'Quantidade', 'Cor', 'Depósito'];
  const ALIASES = [
    ['codigo', 'cod', 'codigoproduto', 'codigodoproduto'],
    ['descricao', 'descricaoproduto', 'descricaodoproduto'],
    ['qtde', 'qtd', 'quantidade'],
    ['descor', 'cor', 'descricaocor', 'descricaodacor'],
    ['descdep', 'deposito', 'descricaodeposito', 'descricaododeposito']
  ];
  const numberFormat = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 12 });
  const clean = value => String(value ?? '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim();
  const normalize = value => clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const isHeader = values => ALIASES.every((names, i) => names.includes(normalize(values[i])));
  function cellText(cell) {
    if (!cell || cell.t === 'z') return '';
    // O texto formatado preserva códigos numéricos com máscara 000000.
    return clean(cell.w ?? cell.v);
  }
  function quantity(cell) {
    if (!cell || clean(cell.v) === '') return { number: null, text: '—' };
    if (cell.t === 'e') throw new Error('quantidade contém erro do Excel');
    if (typeof cell.v === 'number') {
      if (!Number.isFinite(cell.v)) throw new Error('quantidade inválida');
      return { number: cell.v, text: numberFormat.format(cell.v) };
    }
    let value = clean(cell.v).replace(/[\s\u00a0]/g, '');
    if (!/^[+-]?\d[\d.,]*$/.test(value)) throw new Error('quantidade não numérica');
    if (value.includes(',')) {
      // Convenção brasileira em campos textuais: 1.234,50 ou 1234,50.
      if (!/^[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+),\d+$/.test(value)) throw new Error('quantidade com separadores inválidos');
      value = value.replace(/\./g, '').replace(',', '.');
    } else if (/^[+-]?\d{1,3}(?:\.\d{3})+$/.test(value)) {
      value = value.replace(/\./g, '');
    } else if (!/^[+-]?\d+(?:\.\d+)?$/.test(value)) throw new Error('quantidade com separadores inválidos');
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) throw new Error('quantidade inválida');
    return { number: parsed, text: numberFormat.format(parsed) };
  }
  function headerRow(sheet) {
    if (!sheet?.['!ref']) return null;
    const range = root.XLSX.utils.decode_range(sheet['!ref']);
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 199); r++) {
      const values = Array.from({ length: 5 }, (_, c) => cellText(sheet[root.XLSX.utils.encode_cell({ r, c })]));
      if (isHeader(values)) return r;
    }
    return null;
  }
  function inspectWorkbook(workbook) {
    return workbook.SheetNames.map(name => ({ name, header: headerRow(workbook.Sheets[name]) })).filter(sheet => sheet.header !== null);
  }
  function extract(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const header = headerRow(sheet);
    if (header === null) throw new Error('Não encontrei os cabeçalhos Código, Descrição, Qtde, Descor e DescDep nas colunas A a E. Envie a exportação de estoque com essas cinco colunas no início.');
    const range = root.XLSX.utils.decode_range(sheet['!ref']);
    const rows = [];
    const warnings = [];
    let blankRows = 0, repeatedHeaders = 0, totalRows = 0, missingQuantities = 0, incompleteRows = 0;
    for (let r = header + 1; r <= range.e.r; r++) {
      const cells = Array.from({ length: 5 }, (_, c) => sheet[root.XLSX.utils.encode_cell({ r, c })]);
      const values = cells.map(cellText);
      if (values.every(v => !v)) { blankRows++; continue; }
      if (isHeader(values)) { repeatedHeaders++; continue; }
      // Totais de exportação são descartados apenas quando claramente identificados.
      const totalLabels = ['total', 'totalgeral', 'subtotal', 'totais'];
      const totalLabel = values[0] || values[1];
      if (totalLabels.includes(normalize(totalLabel)) && (!values[0] || !values[1]) && !values[3] && !values[4]) { totalRows++; continue; }
      for (let c = 0; c < 5; c++) {
        if (cells[c]?.t === 'e') throw new Error(`A célula ${root.XLSX.utils.encode_cell({ r, c })} contém erro do Excel. Corrija esse valor na origem e envie novamente.`);
        if (cells[c]?.f && cells[c].v === undefined) throw new Error(`A célula ${root.XLSX.utils.encode_cell({ r, c })} tem fórmula sem resultado salvo. Abra e salve a planilha no Excel ou LibreOffice antes de enviar.`);
      }
      let qty;
      try { qty = quantity(cells[2]); } catch (error) { throw new Error(`Linha ${r + 1}: ${error.message} (coluna C). O PDF não foi gerado para evitar um estoque incorreto.`); }
      if (qty.number === null) missingQuantities++;
      if ([0, 1, 3, 4].some(i => !values[i])) incompleteRows++;
      rows.push({ code: values[0], description: values[1], quantity: qty.number, quantityText: qty.text, color: values[3], warehouse: values[4], sourceRow: r + 1 });
    }
    if (!rows.length) throw new Error('A aba tem cabeçalhos, mas não há produtos nas colunas A a E.');
    if (missingQuantities) warnings.push(`${missingQuantities} linha(s) sem quantidade: o PDF mostra um traço, sem transformar campos vazios em zero.`);
    if (incompleteRows) warnings.push(`${incompleteRows} linha(s) com código, descrição, cor ou depósito vazio: os registros foram mantidos e os campos vazios aparecem como traço.`);
    return { rows, warnings, cleanup: { blankRows, repeatedHeaders, totalRows }, sheetName };
  }
  function parseWorkbook(data, name) {
    const extension = name.toLowerCase().split('.').pop();
    // Ler CSV como texto evita converter 003627 em 3627 ou códigos em datas.
    if (['csv', 'tsv', 'txt'].includes(extension)) {
      let text = new TextDecoder('utf-8').decode(data);
      if (text.includes('\ufffd')) text = new TextDecoder('windows-1252').decode(data);
      if (data[0] === 255 && data[1] === 254) text = new TextDecoder('utf-16le').decode(data);
      if (data[0] === 254 && data[1] === 255) text = new TextDecoder('utf-16be').decode(data);
      text = text.replace(/^\ufeff/, '');
      return root.XLSX.read(text, { type: 'string', raw: true, cellText: true });
    }
    return root.XLSX.read(data, { type: 'array', cellText: true, cellNF: true });
  }
  function createPdf(report, sourceName, date = new Date()) {
    const { jsPDF } = root.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true, putOnlyUsedFonts: true });
    doc.setProperties({ title: 'Relatório de estoque', subject: 'Código, descrição, quantidade, cor e depósito', creator: 'Estoque em PDF', author: '' });
    const stamp = new Intl.DateTimeFormat('pt-BR', { dateStyle: 'short', timeStyle: 'short' }).format(date);
    const width = doc.internal.pageSize.getWidth();
    const height = doc.internal.pageSize.getHeight();
    const left = 12, right = 12;
    const source = clean(sourceName) + '  |  Aba: ' + clean(report.sheetName);
    doc.autoTable({
      head: [LABELS],
      body: report.rows.map(row => [row.code || '—', row.description || '—', row.quantityText, row.color || '—', row.warehouse || '—']),
      startY: 39,
      margin: { top: 39, bottom: 16, left, right },
      tableWidth: width - left - right,
      theme: 'plain',
      showHead: 'everyPage', rowPageBreak: 'avoid',
      styles: { font: 'helvetica', fontSize: 8, textColor: [37, 49, 46], cellPadding: { top: 3.2, bottom: 3.2, left: 3, right: 3 }, overflow: 'linebreak', valign: 'middle', lineWidth: { bottom: 0.15 }, lineColor: [229, 234, 228] },
      headStyles: { fillColor: [23, 47, 43], textColor: [255, 255, 255], fontSize: 8, fontStyle: 'bold', cellPadding: { top: 3.5, bottom: 3.5, left: 3, right: 3 }, lineWidth: 0 },
      alternateRowStyles: { fillColor: [246, 248, 243] },
      columnStyles: { 0: { cellWidth: 25 }, 1: { cellWidth: 113 }, 2: { cellWidth: 22, halign: 'right' }, 3: { cellWidth: 48 }, 4: { cellWidth: width - left - right - 208 } },
      didParseCell(data) { if (data.section === 'head' && data.column.index === 2) data.cell.styles.halign = 'right'; },
      willDrawPage() {
        doc.setFillColor(82, 120, 73); doc.rect(left, 11, 11, 1.3, 'F');
        doc.setFont('helvetica', 'bold'); doc.setFontSize(19); doc.setTextColor(23, 47, 43); doc.text('Relatório de estoque', left, 22);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(106, 119, 106);
        const sourceLines = doc.splitTextToSize(source, 176);
        // Metadados limitados a duas linhas; descrições dos produtos nunca são truncadas.
        doc.text(sourceLines.slice(0, 2), left, 28);
        doc.text('Gerado em ' + stamp, width - right, 17, { align: 'right' });
        doc.text(numberFormat.format(report.rows.length) + ' registros de estoque', width - right, 23, { align: 'right' });
      }
    });
    const pages = doc.getNumberOfPages();
    for (let i = 1; i <= pages; i++) {
      doc.setPage(i); doc.setDrawColor(216, 224, 213); doc.setLineWidth(.2); doc.line(left, height - 12, width - right, height - 12);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(112, 125, 108);
      doc.text('ESTOQUE  /  Código · Descrição · Quantidade · Cor · Depósito', left, height - 7);
      doc.text(`Página ${i} de ${pages}`, width - right, height - 7, { align: 'right' });
    }
    return doc;
  }
  root.StockPDF = { LABELS, clean, normalize, quantity, inspectWorkbook, extract, parseWorkbook, createPdf, numberFormat };
})(typeof window !== 'undefined' ? window : globalThis);
