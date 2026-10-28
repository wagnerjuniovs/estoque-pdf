# Estoque em PDF

Site estático que recebe uma planilha sem limpeza manual e gera um PDF com apenas cinco colunas:

| Coluna de origem | Campo | Título no PDF |
|---|---|---|
| A | Codigo | Código |
| B | descricao | Descrição |
| C | qtde | Quantidade |
| D | Descor | Cor |
| E | DescDep | Depósito |

## Usar agora no computador

Abra **index.html** no Edge, Chrome, Firefox ou Safari. Escolha ou arraste a planilha. O PDF será gerado e seu download será iniciado. O botão **Baixar PDF** permite repetir o download. Não precisa instalar dependências ou iniciar servidor.

## Publicar gratuitamente no GitHub Pages

1. Crie um repositório **público** chamado `estoque-pdf` na conta `wagnerjuniovs`.
2. Envie os arquivos deste projeto. O arquivo **index.html** deve ficar na raiz do repositório, junto com `app.js`, `core.js`, `styles.css`, `favicon.svg` e a pasta **vendor** completa. Não envie a planilha ou PDFs do seu estoque.
3. No repositório, abra **Settings → Pages**. Em **Build and deployment**, selecione **Deploy from a branch**. Escolha a branch **main**, a pasta **/ (root)** e clique em **Save**.
4. Aguarde a publicação indicada pelo GitHub. O endereço esperado, após a ativação, é **https://wagnerjuniovs.github.io/estoque-pdf/**.

Se usar outro nome de repositório, o último trecho do endereço muda. A hospedagem gratuita do Pages está disponível para repositórios públicos no GitHub Free: [documentação oficial](https://docs.github.com/en/pages/quickstart).

## Dados e formatos

- Leitura e PDF acontecem no navegador. Nenhuma planilha é enviada, salva no servidor ou incluída no repositório. Não há analytics, APIs, banco de dados ou login.
- O site funciona com bibliotecas incluídas em **vendor**, sem depender de CDN em tempo de uso.
- Formatos testados: **XLSX, XLS, XLSB, CSV e ODS**. Também aceita extensões XLSM, TSV, FODS, TXT, XML, HTML, NUMBERS, SLK e DIF quando sua estrutura é reconhecida pelo SheetJS. Não significa que qualquer arquivo arbitrário possa ser convertido. Arquivos corrompidos ou protegidos por senha não são suportados.
- Limite por arquivo: **50 MB**. O desempenho depende do tamanho do estoque e da memória do dispositivo.
- Identifica os cabeçalhos nas primeiras 200 linhas usadas de cada aba. As cinco colunas devem estar em **A a E**, na ordem descrita acima. Aceita variações como acentos, `Quantidade`, `Cor` e `Depósito`.
- Uma aba de estoque identificada gera o PDF automaticamente. Se houver mais de uma, pede a escolha da aba, para que nenhum estoque seja omitido sem aviso. Cada PDF corresponde à aba selecionada.
- Mantém a ordem dos registros. Não agrupa, não deduplica e não soma produtos de depósitos diferentes.
- Ignora linhas sem dados em A–E, cabeçalhos repetidos e linhas explicitamente identificadas como total/subtotal sem produto, cor ou depósito. As demais colunas são desconsideradas.
- Preserva códigos em texto e zeros definidos por máscara numérica. Zeros já perdidos no arquivo de origem não podem ser reconstruídos sem conhecer o padrão.
- Remove espaços excedentes e quebras de linha de preenchimento. Descrições dos produtos são quebradas em várias linhas no PDF, sem corte.
- Quantidades numéricas do Excel mantêm seu valor. Quantidades em texto seguem convenção brasileira: `1.234,50` → `1.234,5`, `1.234` → `1.234`. Decimal textual com ponto e sem agrupamento também é reconhecido, como `12.5`. CSV de outra convenção com valores ambíguos deve ser exportado como Excel numérico.
- Valores vazios aparecem como traço, sem serem convertidos em zero. Saldo zero e negativo são mantidos. Quantidade não numérica ou erro de célula bloqueia a geração e informa a linha.
- Fórmulas usam o resultado salvo no arquivo. O site não recalcula fórmulas; salve antes no Excel/LibreOffice quando necessário.
- PDF A4 horizontal, texto selecionável, faixas discretas entre linhas, cabeçalho em todas as páginas e rodapé com página atual/total. Os cinco campos são os únicos dados de produtos incluídos. A fonte de origem, aba e horário de geração aparecem como identificação do relatório.
- Prévia na tela limitada a 100 registros. **O PDF inclui todos os registros válidos da aba**.

## Arquivos

- `index.html`, `styles.css`: interface responsiva e acessível.
- `core.js`: identificação de abas, leitura, limpeza, validação e PDF.
- `app.js`: upload, seleção de aba, prévia e download.
- `vendor/`: bibliotecas e respectivas licenças.
- `.nojekyll`: publicação dos arquivos estáticos sem processamento Jekyll.
- `tests/`: testes de navegador e leitura do PDF. Não contém planilhas reais.

## Verificação técnica

As dependências de desenvolvimento são opcionais e não participam do site publicado. Com Node.js e Microsoft Edge instalados:

```sh
npm install
npm test
```

O teste usa Edge em modo headless, cria arquivos sintéticos em `artifacts/` e verifica importação em cinco formatos, preservação de códigos/acentos, paginação com 240 produtos, cabeçalhos e rodapés, campos vazios, erros, várias abas, conteúdo HTML e tela de celular. Se `../teste.xlsx` existir, verifica também a planilha original. Artefatos e arquivos reais são ignorados pelo Git.

## Bibliotecas

- SheetJS CE 0.20.3 — Apache-2.0. [Distribuição oficial](https://docs.sheetjs.com/docs/getting-started/installation/standalone/).
- jsPDF 4.2.1 — MIT. [Projeto](https://github.com/parallax/jsPDF).
- jsPDF-AutoTable 5.0.9 — MIT. [Projeto](https://github.com/simonbengtsson/jsPDF-AutoTable).

Licenças incluídas na pasta `vendor`. Não é necessário contratar nenhum serviço para usar o site.
