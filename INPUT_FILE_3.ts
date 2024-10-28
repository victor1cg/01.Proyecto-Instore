function main(workbook: ExcelScript.Workbook) {


  //! --------------INPUT FINAL 28.10.2024

  // Acessar as planilhas necessárias
  const inputFinal = workbook.getWorksheet("INPUT_FINAL");
  const inputInicial = workbook.getWorksheet("OVERVIEW");
  const w_instore = workbook.getWorksheet("INSTORE");
  const w_online = workbook.getWorksheet("ONLINE");

  // Limpar a área de destino na planilha 'INPUT_FINAL'
  inputFinal.getRange("A2:N40").clear(ExcelScript.ClearApplyTo.contents);

  // Obter a última linha da planilha 'OVERVIEW'
  const ultimaLinhaInicial = inputInicial.getCell(inputInicial.getRange().getRowCount() - 1, 9)
    .getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();
  console.log(ultimaLinhaInicial)

  // Variáveis para controle de colunas nas planilhas 'INSTORE' e 'ONLINE'
  let colAtual_Instore = 0;
  let colAtual_Online = 0;
  let linhaDestino = 1;

  // ultima coluna de ONLINE e INSTORE
  // let col_final_instore: number = w_instore.getCell(0, 0)
  //   .getRangeEdge(ExcelScript.KeyboardDirection.right)
  //   .getColumnIndex();

  // let col_final_online: number = w_online.getCell(0, 0)
  //   .getRangeEdge(ExcelScript.KeyboardDirection.right)
  //   .getColumnIndex();

  //! Loop através das linhas da planilha 'OVERVIEW'
  for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial; linhaAtual++) {

    // Copiar os dados da linha atual (colunas D a K)
    const valoresLinha = inputInicial.getRangeByIndexes(linhaAtual, 3, 1, 12).getValues();

    // Colocar os valores copiados na planilha 'INPUT_FINAL'
    inputFinal.getRangeByIndexes(linhaDestino, 0, 1, valoresLinha[0].length).setValues(valoresLinha);

    // Obter o valor da coluna de mídia na linha atual
    const valorMidia = inputFinal.getCell(linhaDestino, 1).getValue();
    console.log('valorMidia '+ valorMidia)
    // Processar conforme o tipo de mídia (Instore ou Online)
    if (valorMidia === "Instore") {
      // Concatenar SKUs e lojas da planilha 'INSTORE'
      const concatenadoSkus = concatenarValores(w_instore, colAtual_Instore);
      colAtual_Instore++;
      const concatenadoLojas = concatenarValores(w_instore, colAtual_Instore);
      colAtual_Instore++;

      // Colocar os valores concatenados na planilha 'INPUT_FINAL'
      inputFinal.getRange(`M${linhaDestino + 1}`).setValue(concatenadoLojas);
      inputFinal.getRange(`N${linhaDestino + 1}`).setValue(concatenadoSkus);
    } else {
      // Concatenar SKUs da planilha 'ONLINE'
      const concatenadoSkus = concatenarValores(w_online, colAtual_Online);
      colAtual_Online++;
      // if (colAtual_Online > col_final_online) { break }

      // Colocar os SKUs concatenados na planilha 'INPUT_FINAL'
      inputFinal.getRange(`N${linhaDestino + 1}`).setValue(concatenadoSkus);
    }

    // Atualizar a linha de destino
    linhaDestino++;
  }
}

//! Função 
function concatenarValores(sheet: ExcelScript.Worksheet, colIndex: number): string {
  let row_final: number = sheet.getCell(2, colIndex)
    .getRangeEdge(ExcelScript.KeyboardDirection.down)
    .getRowIndex();
  console.log('colIndex ' + colIndex)
  console.log('row_final ' + row_final)

  let rangeLojas = sheet.getRangeByIndexes(2, colIndex, row_final - 1, 1);

  // Se não tiver dados, Raise Error
  if (rangeLojas.getRowCount() > 100) { throw new Error(`Prencher valores de SKUs ou Lojas, não deixar campanha em branco!`); }
  let valores = rangeLojas.getValues();
  return valores.map(valor => valor[0]).join(";");
}
