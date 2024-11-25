function main(workbook: ExcelScript.Workbook) {
  // Acessa a planilha 'OVERVIEW'
  let w_inputInicial = workbook.getWorksheet("2-OVERVIEW");

  // Determinar a última linha preenchida na coluna J
  let ultimaLinhaInicial = w_inputInicial.getCell(w_inputInicial.getRange().getRowCount() - 1, 9)
    .getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();

  // Acessar as planilhas 'INSTORE' e 'ONLINE'
  let w_instore = workbook.getWorksheet('3-INSTORE');
  let w_online = workbook.getWorksheet('3-ONLINE');

  let rangeLimpar = w_instore.getRange("A1:AB1");
  rangeLimpar.clear(ExcelScript.ClearApplyTo.contents);
  let rangeLimpar2 = w_online.getRange("A1:AB40");
  rangeLimpar2.clear(ExcelScript.ClearApplyTo.contents);

  // Definir variáveis para controle das linhas onde os dados serão colados
  let colInstore = 0;
  let colOnline = 0;

  // Loop para verificar cada linha da coluna H e copiar para a planilha correspondente
  for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial;) {
    // Obter o valor da coluna E (index 4) na linha atual
    let valorColunaE = w_inputInicial.getCell(linhaAtual, 4).getValue() as string;

    // Obter o valor da coluna ao lado (index 5, coluna F)
    let valorProduto = w_inputInicial.getCell(linhaAtual, 5).getValue();
    const valorDate = w_inputInicial.getCell(linhaAtual, 11).getValue();

    if (valorColunaE === "Instore") {
      // Colar na aba INSTORE de forma incremental
      colar_skus(w_instore, colInstore, valorProduto, valorDate);
      colInstore++;
      colar_skus(w_instore, colInstore, valorProduto, valorDate);
      colInstore++;

    } else {
      // Colar na aba ONLINE de forma incremental
      colar_skus(w_online,colOnline,valorProduto,valorDate);
      
      // SKU
      let cel_sku2 = w_online.getCell(2, colOnline);
      cel_sku2.setValue('SKU');
      cel_sku2.getFormat().getFill().setColor('e8e8e8');
      colOnline++;
    }
    linhaAtual++;
  }
  w_instore.getRange("A1:AB40").getFormat().autofitColumns();
  w_online.getRange("A1:AB40").getFormat().autofitColumns();


// Função ----------------------
  function colar_skus(sheet: ExcelScript.Worksheet, col_n: number, valorProduto: string, valorDate: string): string {
    // Produto
    sheet.getCell(0, col_n).setValue(valorProduto);
    sheet.getCell(0, col_n).getFormat().getFont().setBold(true);
    sheet.getCell(0, col_n).getFormat().getFill().setColor('e8e8e8');

    // Data
    sheet.getCell(1, col_n).setValue(valorDate);
    sheet.getCell(1, col_n).getFormat().getFill().setColor('e8e8e8');

  }
}
