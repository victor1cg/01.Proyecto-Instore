function main(workbook: ExcelScript.Workbook) {
    
    // Acessar a planilha 'INPUT_FINAL'
    let inputFinal = workbook.getWorksheet("INPUT_FINAL");
    let linhaDestino = 2;

    // LIMPA o intervalo de A2:U40 na aba 'input_final'
    let rangeLimpar = inputFinal.getRange("A2:U40");
    rangeLimpar.clear();
    
        //------------- INPUT INICIAL
    // Acessa a planilha 'INPUT_INICIAL'
    let inputInicial = workbook.getWorksheet("INPUT_INICIAL (2)");

    // Copia as células de D2:K2 da aba 'input_inicial'
    let rangeInicial = inputInicial.getRange("D2:K2").getValues();
    // let valoresLinha = rangeInicial.getValues();
    // pegar a ultima celula
    let ultimaLinhaInicial = inputInicial.getCell(inputInicial.getRange().getRowCount()-1, 5).getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex()

    //------------- LOJAS
    // Copia os dados de B9 da aba 'input_inicial'
    let rangeB = inputInicial.getRange("B9").getExtendedRange(ExcelScript.KeyboardDirection.down);
    let valoresB = rangeB.getValues();

    // Concatena os valores de B9:B25 em uma única string, separados por ";"
    let concatenados = valoresB.map(valor => valor[0]).join(";");
    console.log("ultimaLinhaInicial " + ultimaLinhaInicial)

    //------------- LOOP tabela
    // loop para copiar de D2:K abaixo até a última linha preenchida
    for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial; linhaAtual++) 
    {
        // Copia os dados da linha atual (colunas D a K)
        let rangeLinha = inputInicial.getRangeByIndexes(linhaAtual, 3, 1, 7); 
        // Linha -1 (para índice zero) e colunas 3 a 9 (D a K)
        let valoresLinha = rangeLinha.getValues();

        // Define o intervalo de destino na aba 'input_final' (na linha atual de destino)
        let destinoLinha = inputFinal.getRangeByIndexes(linhaDestino - 1, 0, 1, valoresLinha[0].length); // Colando em A2, A3, etc.

        // Coloca os valores copiados na aba 'input_final'
        destinoLinha.setValues(valoresLinha);

        // Atualiza a linha de destino para a próxima linha
        linhaDestino++;

        // console.log("linhaAtual " + linhaAtual)

        // console.log("linhaDestino " + linhaDestino)

    }

    //------------- LOJAS
    console.log("linhaDestino " + linhaDestino)
    let destinoConcatenado = inputFinal.getRange(`I2:I${linhaDestino}`); 
    // Preencher até a última linha copiada
    destinoConcatenado.setValue(concatenados);

    // ---------------VALORES FINANCEIROS
    const invest = inputInicial.getRange('B2').getValue();
    let destino_inv = inputFinal.getRange(`J2:J${linhaDestino}`);
    destino_inv.setValue(invest);

    const desconto_per = inputInicial.getRange('B3').getValue();
    let destino_desc = inputFinal.getRange(`K2:K${linhaDestino}`);
    destino_desc.setValue(desconto_per);

    const custos = inputInicial.getRange('B4').getValue();
    let destino_custos = inputFinal.getRange(`L2:L${linhaDestino}`);
    destino_custos.setValue(custos);

    const margem = inputInicial.getRange('B5').getValue();
    let destino_margem = inputFinal.getRange(`M2:M${linhaDestino}`);
    destino_margem.setValue(margem);

    const margem_per = inputInicial.getRange('B6').getValue();
    let destino_marg_perc = inputFinal.getRange(`N2:N${linhaDestino}`);
    destino_marg_perc.setValue(margem_per);

    
}

