function main(workbook: ExcelScript.Workbook) {

    //########## SCRIPT FINAL! ##########
    
	// Acessar a planilha 'INPUT_FINAL'
    let inputFinal = workbook.getWorksheet("INPUT_FINAL");
    let linhaDestino = 1;
	
    let rangeLimpar = inputFinal.getRange("A2:U40");
    rangeLimpar.clear(ExcelScript.ClearApplyTo.contents);

	// Acessar as planilhas 'INSTORE' e 'ONLINE'
    let w_instore = workbook.getWorksheet('INSTORE');
    let w_online = workbook.getWorksheet('ONLINE');

    // pegar a ultima coluna 
    let ultimaCol_Instore = w_instore.getCell(w_instore.getRange().getRowCount()-1, 0).getRangeEdge(ExcelScript.KeyboardDirection.right).getColumnIndex()
    let ultimaCol_Online = w_online.getCell(w_online.getRange().getRowCount()-1, 0).getRangeEdge(ExcelScript.KeyboardDirection.right).getColumnIndex()
    
    //------------- INPUT INICIAL
    let inputInicial = workbook.getWorksheet("OVERVIEW");

    // Copia as células de D2:L2 da aba 'input_inicial'
    let rangeInicial = inputInicial.getRange("D2:K2").getValues();
    let ultimaLinhaInicial = inputInicial.getCell(inputInicial.getRange().getRowCount()-1, 5).getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex()

    
    //!------------- LOOP tabela
    for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial; linhaAtual++) 
        {
 
            // Copia os dados da linha atual (colunas D a K)
            let rangeLinha = inputInicial.getRangeByIndexes(linhaAtual, 3, 1, 9); 
            // Linha -1 (para índice zero) e colunas 3 a 9 (D a K)
            let valoresLinha = rangeLinha.getValues();
    
            // Define o intervalo de destino na aba 'input_final' (na linha atual de destino)
            let destinoLinha = inputFinal.getRangeByIndexes(linhaDestino , 0, 1, valoresLinha[0].length);
    
            // Coloca os valores copiados na aba 'input_final'
            destinoLinha.setValues(valoresLinha);
    
            // Atualiza a linha de destino para a próxima linha
            linhaDestino++;
        
        }
    
    //!------------- LOOP INSTORE

    let colAtual_Instore = 0;
    let colAtual_Online = 0;
    let linhaDestino_Final = 1;

    for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial; linhaAtual++) 
    {

        let valorMidia = inputFinal.getCell(linhaDestino_Final,4).getValue()

        console.log(linhaAtual)
        console.log(ultimaLinhaInicial)
        console.log(valorMidia)

        if(valorMidia == 'Instore')
            {
            let rangeSkus = w_instore.getCell(2,colAtual_Instore).getExtendedRange(ExcelScript.KeyboardDirection.down);
            let valoresSkus = rangeSkus.getValues();
            let concatenadoSkus = valoresSkus.map(valor => valor[0]).join(";");
            colAtual_Instore++
            
            let rangeLojas = w_instore.getCell(2,colAtual_Instore).getExtendedRange(ExcelScript.KeyboardDirection.down);
            let valoresLojas = rangeLojas.getValues();
            let concatenadoLojas = valoresLojas.map(valor => valor[0]).join(";");
            colAtual_Instore++

            let destinoConcatenadoLojas = inputFinal.getRange(`J${linhaDestino_Final}`); 
            destinoConcatenadoLojas.setValue(concatenadoLojas)
    
            let destinoConcatenadoSkus = inputFinal.getRange(`K${linhaDestino_Final}`); 
            destinoConcatenadoSkus.setValue(concatenadoSkus)
            }
        else
            {
            let rangeSkus = w_online.getCell(1,colAtual_Online).getExtendedRange(ExcelScript.KeyboardDirection.down);
            let valoresSkus = rangeSkus.getValues();
            let concatenadoSkus = valoresSkus.map(valor => valor[0]).join(";");
            let destinoConcatenadoSkus = inputFinal.getRange(`K${linhaDestino_Final}`); 
            destinoConcatenadoSkus.setValue(concatenadoSkus)

            colAtual_Online++
            }   
                

            linhaDestino_Final++;
    }

    
}

