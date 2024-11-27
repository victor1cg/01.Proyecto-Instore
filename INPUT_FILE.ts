function main(workbook: ExcelScript.Workbook) {
  // Acessar as planilhas necessárias
  const inputFinal = workbook.getWorksheet("4-INPUT_FINAL");
  const inputInicial = workbook.getWorksheet("2-OVERVIEW");
  const w_instore = workbook.getWorksheet("3-INSTORE");
  const w_online = workbook.getWorksheet("3-ONLINE");
  const w_tarifario = workbook.getWorksheet("1-TARIFARIO");

  // Limpar a área de destino dinamicamente
  const rowCount = inputFinal.getUsedRange()?.getRowCount() || 0;
  inputFinal.getRange(`A2:S${rowCount}`).clear(ExcelScript.ClearApplyTo.contents);

  // Obter a última linha da planilha 'OVERVIEW'
  const ultimaLinhaInicial = inputInicial.getCell(inputInicial.getRange().getRowCount() - 1, 9)
    .getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();

  // Variáveis para controle
  let colAtual_Instore = 0;
  let colAtual_Online = 0;
  let linhaDestino = 1;

  const pm = inputInicial.getRange("B9").getValues()[0][0];
  const bandeira = inputInicial.getRange("B10").getValues()[0][0];
  const data_ini: Date = w_tarifario.getRange("F2").getValues()[0][0];
  const data_fim: Date = w_tarifario.getRange("F3").getValues()[0][0];

  // Loop principal
  for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial; linhaAtual++) {
    const valoresLinha = inputInicial.getRangeByIndexes(linhaAtual, 3, 1, 13).getValues();
    inputFinal.getRangeByIndexes(linhaDestino, 0, 1, valoresLinha[0].length).setValues(valoresLinha);

    const valorMidia = inputFinal.getCell(linhaDestino, 1).getValue();
    if (valorMidia === "Instore") {
      const concatenadoSkus = concatenarValores(w_instore, colAtual_Instore++);
      const concatenadoLojas = concatenarValores(w_instore, colAtual_Instore++);
      inputFinal.getRange(`P${linhaDestino + 1}`).setValue(concatenadoLojas);
      inputFinal.getRange(`Q${linhaDestino + 1}`).setValue(concatenadoSkus);
    } else {
      const concatenadoSkus = concatenarValores(w_online, colAtual_Online++);
      inputFinal.getRange(`Q${linhaDestino + 1}`).setValue(concatenadoSkus);
    }

    inputFinal.getRange(`N${linhaDestino + 1}`).setValue(pm);
    inputFinal.getRange(`O${linhaDestino + 1}`).setValue(bandeira);
    inputFinal.getRange(`R${linhaDestino + 1}`).setValue(data_ini);
    inputFinal.getRange(`S${linhaDestino + 1}`).setValue(data_fim);

    linhaDestino++;
  }
  inputFinal.getRange("A1:S40").getFormat().autofitColumns();
}

function concatenarValores(sheet: ExcelScript.Worksheet, colIndex: number): string {
  const row_value3 = sheet.getCell(5, colIndex).getValues()[0][0] as string;
  const row_value2 = sheet.getCell(4, colIndex).getValues()[0][0] as string;

  console.log(row_value3)
  console.log(row_value2)
  if (!row_value3 && !row_value2) {
    throw new Error(`Preencher valores de SKUs ou Lojas, não deixar campanha em branco!`);
  } else if (!row_value3 && row_value2) {
    return row_value2;
  } else {
    const row_final = sheet.getCell(4, colIndex)
      .getRangeEdge(ExcelScript.KeyboardDirection.down)
      .getRowIndex();
    console.log('row final '+row_final)
    
    const rangeLojas = sheet.getRangeByIndexes(4, colIndex, row_final - 3, 1);
    if (rangeLojas.getRowCount() > 100) {
      throw new Error(`Limite de SKUs/Lojas ultrapassado!`);
    }

    const valores = rangeLojas.getValues().map(valor => valor[0] as string);
    return valores.join(";");
  }
}
