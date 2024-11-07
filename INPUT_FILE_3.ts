function main(workbook: ExcelScript.Workbook) {


  //! --------------INPUT FINAL 28.10.2024

  // Acessar as planilhas necessárias
  const inputFinal = workbook.getWorksheet("4-INPUT_FINAL");
  const inputInicial = workbook.getWorksheet("2-OVERVIEW");
  const w_instore = workbook.getWorksheet("3-INSTORE");
  const w_online = workbook.getWorksheet("3-ONLINE");

  // Limpar a área de destino na planilha 'INPUT_FINAL'
  inputFinal.getRange("A2:P40").clear(ExcelScript.ClearApplyTo.contents);

  // Obter a última linha da planilha 'OVERVIEW'
  const ultimaLinhaInicial = inputInicial.getCell(inputInicial.getRange().getRowCount() - 1, 9)
    .getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();


  // Variáveis para controle de colunas nas planilhas 'INSTORE' e 'ONLINE'
  let colAtual_Instore = 0;
  let colAtual_Online = 0;
  let linhaDestino = 1;

  const pm = inputInicial.getRange('B9').getValues()
  const bandeira = inputInicial.getRange('B10').getValues()


  //! Loop através das linhas da planilha 'OVERVIEW'
  for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial; linhaAtual++) {

    // Copiar os dados da linha atual (colunas D a K)
    const valoresLinha = inputInicial.getRangeByIndexes(linhaAtual, 3, 1, 12).getValues();

    // Colocar os valores copiados na planilha 'INPUT_FINAL'
    inputFinal.getRangeByIndexes(linhaDestino, 0, 1, valoresLinha[0].length).setValues(valoresLinha);

    // Obter o valor da coluna de mídia na linha atual
    const valorMidia = inputFinal.getCell(linhaDestino, 1).getValue();

    // Processar conforme o tipo de mídia (Instore ou Online)
    if (valorMidia === "Instore") {
      // Concatenar SKUs e lojas da planilha 'INSTORE'
      const concatenadoSkus = concatenarValores(w_instore, colAtual_Instore);
      colAtual_Instore++;
      const concatenadoLojas = concatenarValores(w_instore, colAtual_Instore);
      colAtual_Instore++;

      // Colocar os valores concatenados na planilha 'INPUT_FINAL'
      inputFinal.getRange(`O${linhaDestino + 1}`).setValue(concatenadoLojas);
      inputFinal.getRange(`P${linhaDestino + 1}`).setValue(concatenadoSkus);
    }

    else {
      // Concatenar SKUs da planilha 'ONLINE'
      const concatenadoSkus = concatenarValores(w_online, colAtual_Online);
      colAtual_Online++;
      // if (colAtual_Online > col_final_online) { break }

      // Colocar os SKUs concatenados na planilha 'INPUT_FINAL'
      inputFinal.getRange(`P${linhaDestino + 1}`).setValue(concatenadoSkus);
    }

    //! PM e Bandeira
    inputFinal.getRange(`M${linhaDestino + 1}`).setValue(pm);
    inputFinal.getRange(`N${linhaDestino + 1}`).setValue(bandeira);

    // Atualizar a linha de destino
    linhaDestino++;
  }
}

//! Função 
function concatenarValores(sheet: ExcelScript.Worksheet, colIndex: number): string {

  // verificar o valor da linha 3, se for null 
  let row_value3: string = sheet.getCell(3, colIndex).getValues()
  let row_value2: string = sheet.getCell(2, colIndex).getValue()
  let valores: string[] = [];


  if (row_value3 == "" && row_value2 == "") {
    throw new Error(`Prencher valores de SKUs ou Lojas, não deixar campanha em branco!`);
  }

  else if (row_value3 == "" && row_value2 != "") {
    // console.log(row_value2)
    return row_value2;
  }


  else {
    let row_final: number = sheet.getCell(2, colIndex)
      .getRangeEdge(ExcelScript.KeyboardDirection.down)
      .getRowIndex();

    let rangeLojas = sheet.getRangeByIndexes(2, colIndex, row_final - 1, 1);
    // Se não tiver dados, Raise Error
    // console.log('entrou no else')
    if (rangeLojas.getRowCount() > 100) {
      throw new Error(`Prencher valores de SKUs ou Lojas, não deixar campanha em branco!`);
    }
    valores = rangeLojas.getValues().map(valor => valor[0] as string);
    return valores.join(";");
  }

}
