function main(workbook: ExcelScript.Workbook) {
  // Acessa a planilha 'OVERVIEW'
  let w_inputInicial = workbook.getWorksheet("2-OVERVIEW");

  // Determinar a última linha preenchida na coluna J
  let ultimaLinhaInicial = w_inputInicial.getCell(w_inputInicial.getRange().getRowCount() - 1, 9)
    .getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();

  // Acessar as planilhas 'INSTORE' e 'ONLINE'
  let w_instore = workbook.getWorksheet('3-INSTORE');
  let w_online = workbook.getWorksheet('3-ONLINE');

  let rangeLimpar = w_instore.getRange("A1:AB4");
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
    console.log(valorColunaE)


    // Obter o valor da coluna ao lado (index 5, coluna F)
    let valorProduto = w_inputInicial.getCell(linhaAtual, 5).getValue();
    const valorDate = w_inputInicial.getCell(linhaAtual, 11).getValue();
    const campanha = w_inputInicial.getCell(linhaAtual, 10).getValue();
    let cel_sku_instore = w_instore.getCell(3, colInstore);
    
    if (valorColunaE === "Instore") {
      // Colar na aba INSTORE de forma incremental
      console.log('Entrou no Instore');
      
      // SKU
      let cel_sku_instore = w_instore.getCell(3, colInstore);
      cel_sku_instore.setValue('SKU');
      cel_sku_instore.getFormat().getFill().setColor('c2c2c2');
      
      // Loja (ao lado)
      let cel_loja_instore = w_instore.getCell(3, colInstore + 1);
      cel_loja_instore.setValue('LOJA');
      cel_loja_instore.getFormat().getFill().setColor('c2c2c2');
      
      // Inserir dados para SKU e Loja lado a lado
      colar_skus(w_instore, colInstore, valorProduto, campanha, valorDate);
      colar_skus(w_instore, colInstore + 1, "Nome da Loja", campanha, valorDate); // Ajuste para o valor da loja
      
      // Incrementa duas colunas
      colInstore += 2;
      }
    else {
      // Colar na aba ONLINE de forma incremental
      colar_skus(w_online,colOnline,valorProduto,campanha, valorDate);
      console.log('Entrou no online')
      console.log(campanha)
      
      // SKU
      let cel_sku2 = w_online.getCell(3, colOnline);
      cel_sku2.setValue('SKU');
      cel_sku2.getFormat().getFill().setColor('c2c2c2');
      colOnline++;
    }
    linhaAtual++;
  }
  w_instore.getRange("A1:AB40").getFormat().autofitColumns();
  w_online.getRange("A1:AB40").getFormat().autofitColumns();


// Função ----------------------
  function colar_skus(sheet: ExcelScript.Worksheet, col_n: number, valorProduto: string, campanha :string,valorDate: string): string {
    // Produto
    sheet.getCell(0, col_n).setValue(valorProduto);
    sheet.getCell(0, col_n).getFormat().getFont().setBold(true);
    sheet.getCell(0, col_n).getFormat().getFill().setColor('e8e8e8');

    // Campanha
    sheet.getCell(1, col_n).setValue(campanha);
    sheet.getCell(1, col_n).getFormat().getFont().setBold(true);
    sheet.getCell(1, col_n).getFormat().getFill().setColor('e8e8e8');

    // Data
    sheet.getCell(2, col_n).setValue(valorDate);
    sheet.getCell(2, col_n).getFormat().getFill().setColor('e8e8e8');


  }
}
