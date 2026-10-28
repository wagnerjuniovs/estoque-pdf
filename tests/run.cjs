/* Testes reais de navegador e leitura do PDF. Execute: node tests/run.cjs */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { chromium } = require('playwright-core');
const XLSX = require('../vendor/xlsx.full.min.js');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
const root = path.resolve(__dirname, '..');
const artifacts = path.join(root, 'artifacts');
fs.mkdirSync(artifacts, { recursive: true });
const header = ['Codigo','descricao','qtde','Descor','DescDep','NAO_EXPORTAR'];
const products = [
  ['003627','RIBANA POLIÉSTER PADRÃO 21',210,'PRETO /MADRI','PRONTA ENTREGA','SEGREDO_FORA_DO_PDF'],
  ['009928','Descrição com acentuação e ç',4730,'AZUL','EXCEDENTES','SEGREDO_FORA_DO_PDF'],
  ['016140','CÓS DUPLO',174,'VERDE','ESTOQUE','SEGREDO_FORA_DO_PDF']
];
function book(rows, options = {}) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Estoque');
  if (options.second) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([header,...products]), 'Outro depósito');
  return wb;
}
function fixture(name, wb, bookType) {
  const filename = path.join(artifacts, name);
  fs.writeFileSync(filename, XLSX.write(wb, { type: 'buffer', bookType }));
  return filename;
}
async function pdfText(filename) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(filename)), useSystemFonts: true }).promise;
  let text = ''; const pageTexts = [];
  for (let i=1; i<=doc.numPages; i++) { const p = await doc.getPage(i); const t = await p.getTextContent(); pageTexts.push(t.items.map(i=>i.str).join(' ')); }
  text = pageTexts.join('\n'); await doc.destroy(); return { text, pages: pageTexts };
}
(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1080 }, acceptDownloads: true });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  const appUrl = pathToFileURL(path.join(root,'index.html')).href;
  let count = 0;
  function pass(name) { count++; console.log('OK', name); }
  async function upload(filename, outputName = 'result.pdf') {
    const downloading = page.waitForEvent('download');
    await page.locator('#file-input').setInputFiles(filename);
    const download = await downloading;
    const target = path.join(artifacts, outputName); await download.saveAs(target);
    await page.locator('#download').waitFor({state:'visible'});
    return pdfText(target);
  }
  try {
    await page.goto(appUrl);
    await page.screenshot({ path: path.join(artifacts, 'site-desktop.png'), fullPage: true });
    const original = path.resolve(root, '..', 'teste.xlsx');
    if (fs.existsSync(original)) {
      const result = await upload(original, 'estoque-teste.pdf');
      for (const code of ['003627','009928','016140']) assert(result.text.includes(code));
      for (const excluded of ['1234567926438','61179000','Custo','preco_com']) assert(!result.text.includes(excluded));
      assert(result.text.includes('4.730')); assert(result.text.includes('Depósito'));
      assert(!result.text.includes('teste.xlsx')); assert(!result.text.includes('Plan1')); assert(!result.text.includes('Aba:'));
      assert.equal(await page.locator('#preview-body tr').count(), 3);
      await page.screenshot({ path: path.join(artifacts,'site-com-estoque.png'), fullPage:true });
      pass('planilha original: 3 produtos, códigos preservados, apenas A–E no PDF');
    }
    const excludedProducts = [
      ['008621','DESENVOLVIMENTO VESTUARIO MODELO SIMPLES',27,'NAO APLICAVEL','EXCEDENTES'],
      ['008626','DESENVOLVIMENTO VESTUARIO MODELO JQD',5,'NAO APLICAVEL','EXCEDENTES'],
      ['008627','Produto em DeSeNvOlViMeNtO interno','erro','Cor','Depósito']
    ];
    const filtered = await upload(fixture('exclusao.xlsx',book([header,...excludedProducts,...products]),'xlsx'));
    for (const code of ['008621','008626','008627']) assert(!filtered.text.includes(code));
    for (const code of ['003627','009928','016140']) assert(filtered.text.includes(code));
    assert(!filtered.text.toLowerCase().includes('desenvolvimento'));
    assert(!filtered.text.includes('exclusao.xlsx')); assert(!filtered.text.includes('Aba:'));
    assert.equal(await page.locator('#preview-body tr').count(),3);
    assert((await page.locator('#status').textContent()).includes('3 item(ns)'));
    pass('exclusão integral por descrição, em qualquer posição e caixa; sem nome de arquivo ou aba');
    await page.locator('#file-input').setInputFiles(fixture('todos-excluidos.xlsx',book([header,...excludedProducts]),'xlsx'));
    await page.waitForFunction(()=>document.querySelector('#status').classList.contains('error'));
    assert.equal(await page.locator('#result').isVisible(),false);
    assert((await page.locator('#status').textContent()).includes('Todos os produtos'));
    pass('todos os itens excluídos: informa ausência de produtos e remove PDF anterior');
    for (const type of ['xlsx','xls','xlsb','ods','csv']) {
      const f = fixture('formato.'+type,book([header,...products]),type === 'xls' ? 'biff8' : type);
      const result = await upload(f);
      for (const code of ['003627','009928','016140']) assert(result.text.includes(code),type+' '+code);
      assert(!result.text.includes('SEGREDO_FORA_DO_PDF'));
      assert.equal(await page.locator('#preview-body tr').count(),3);
      assert(result.text.includes('Descrição com acentuação e ç'));
      pass('importação '+type+' e acentos no PDF');
    }
    const csvFile = path.join(artifacts, 'brasileiro.csv');
    fs.writeFileSync(csvFile, '\ufeffCodigo;descricao;qtde;Descor;DescDep\n000012;Produto um;"1.234,50";Azul;Depósito A\n000013;Produto dois;0;Verde;Depósito B\n000014;Produto três;-2;Preto;Depósito A');
    const csvResult = await upload(csvFile);
    assert(csvResult.text.includes('1.234,5')); assert(csvResult.text.includes('-2')); assert(csvResult.text.includes('000012'));
    pass('CSV brasileiro, quantidade decimal, zero, saldo negativo e zeros à esquerda');
    const cleanedFile = fixture('limpeza.xlsx',book([['Exportação do estoque'],[],header,products[0],[],header,products[1],['','Total geral',4940]]),'xlsx');
    await upload(cleanedFile);
    assert.equal(await page.locator('#preview-body tr').count(),2);
    pass('título inicial, linha vazia, cabeçalho repetido e total de rodapé');
    const formatted = book([header,[12,'Produto com máscara',5,'Azul','Depósito']]);
    formatted.Sheets.Estoque.A2.z = '000000';
    const formattedResult = await upload(fixture('mascara.xlsx',formatted,'xlsx'));
    assert(formattedResult.text.includes('000012')); pass('código numérico com máscara de zeros');
    const incomplete = await upload(fixture('vazios.xlsx',book([header,['000001','Produto sem dados',null,'','']]),'xlsx'));
    assert.equal(await page.locator('#preview-body tr').count(),1);
    assert.equal(await page.locator('#warning').isVisible(),true);
    assert(incomplete.text.includes('000001')); pass('campos vazios mantidos e sinalizados');
    const bad = fixture('invalido.xlsx',book([header,['000001','Produto','erro','Preto','Estoque']]),'xlsx');
    await page.locator('#file-input').setInputFiles(bad);
    await page.waitForFunction(()=>document.querySelector('#status').classList.contains('error'));
    assert.equal(await page.locator('#result').isVisible(),false);
    assert((await page.locator('#status').textContent()).includes('Linha 2'));
    pass('quantidade inválida bloqueia PDF e remove resultado anterior');
    const multi = fixture('varias-abas.xlsx',book([header,...products],{second:true}),'xlsx');
    await page.locator('#file-input').setInputFiles(multi);
    await page.locator('#sheet-picker').waitFor({state:'visible'});
    assert.equal(await page.locator('#result').isVisible(),false);
    const multiDownload = page.waitForEvent('download');
    await page.locator('#sheet-select').selectOption('Outro depósito');
    await (await multiDownload).saveAs(path.join(artifacts,'multiplas-abas.pdf'));
    await page.locator('#download').waitFor({state:'visible'});
    assert((await page.locator('#result-summary').textContent()).includes('Outro depósito'));
    pass('várias abas exigem seleção explícita');
    const longRows = Array.from({length:240},(_,i)=>[String(i+1).padStart(6,'0'),`ITEM ${i+1}: DESCRIÇÃO TÉCNICA COMPLETA `+'POLIÉSTER COM ACABAMENTO ESPECIAL '.repeat(i%4+1),i,'AZUL / VERDE','DEPÓSITO DE PRODUTOS ACABADOS']);
    const longPdf = await upload(fixture('paginacao.xlsx',book([header,...longRows]),'xlsx'),'estoque-multipaginas.pdf');
    assert(longPdf.pages.length > 3);
    for(let i=1;i<=240;i++) assert(longPdf.text.includes(String(i).padStart(6,'0')),'Produto ausente: '+i);
    for(let i=0;i<longPdf.pages.length;i++) {
      for(const h of ['Código','Descrição','Quantidade','Cor','Depósito']) assert(longPdf.pages[i].includes(h));
      assert(longPdf.pages[i].includes(`Página ${i+1} de ${longPdf.pages.length}`));
    }
    assert.equal(await page.locator('#preview-body tr').count(),100);
    pass(`paginação: 240 produtos completos em ${longPdf.pages.length} páginas, cabeçalhos e rodapés`);
    const injection = fixture('seguranca.xlsx',book([header,['000001','<img src=x onerror=alert(1)>',1,'Cor','Depósito']]),'xlsx');
    await upload(injection); assert.equal(await page.locator('#preview-body img').count(),0); pass('texto da planilha não executa HTML');
    await page.setViewportSize({ width:390,height:844 }); await page.goto(appUrl);
    await page.screenshot({path:path.join(artifacts,'site-celular.png'),fullPage:true});
    assert(await page.evaluate(()=>document.documentElement.scrollWidth <= innerWidth));
    await upload(fixture('mobile.xlsx',book([header,...products]),'xlsx'));
    assert(await page.locator('#download').isVisible());
    assert(await page.evaluate(()=>document.documentElement.scrollWidth <= innerWidth));
    pass('layout e download no celular, sem rolagem horizontal da página');
    assert.deepEqual(errors,[]);
    console.log(`\n${count} verificações passaram. Nenhum erro de JavaScript.`);
  } finally { await browser.close(); }
})().catch(error=>{console.error(error);process.exitCode=1;});
