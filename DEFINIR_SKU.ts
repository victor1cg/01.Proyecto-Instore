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
    rangeLimpar.clear();
    let rangeLimpar2 = w_online.getRange("A1:AB40");
    rangeLimpar2.clear();

    // Definir variáveis para controle das linhas onde os dados serão colados
    let colInstore = 0;
    let colOnline = 0;

    // Loop para verificar cada linha da coluna H e copiar para a planilha correspondente
    for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial;) {
        // Obter o valor da coluna E (index 4) na linha atual
        let valorColunaE = w_inputInicial.getCell(linhaAtual, 4).getValue() as string;

        // Obter o valor da coluna ao lado (index 5, coluna F)
        let valorProduto = w_inputInicial.getCell(linhaAtual, 5).getValue();

        if (valorColunaE === "Instore") {
            // Colar na aba INSTORE de forma incremental
            w_instore.getCell(0, colInstore).setValue(valorProduto);
            w_instore.getCell(0, colInstore).getFormat().getFont().setBold(true);
            w_instore.getCell(0, colInstore).getFormat().getFill().setColor('e8e8e8');

            let cel_sku = w_instore.getCell(1, colInstore);
            cel_sku.setValue('SKU');
            cel_sku.getFormat().getFont().setBold(true);
            cel_sku.getFormat().getFill().setColor('e8e8e8');
            colInstore++;

            w_instore.getCell(0, colInstore).setValue(valorProduto);
            w_instore.getCell(0, colInstore).getFormat().getFont().setBold(true);
            w_instore.getCell(0, colInstore).getFormat().getFill().setColor('e8e8e8');

            let cel_loja = w_instore.getCell(1, colInstore);
            cel_loja.setValue('LOJA');
            cel_loja.getFormat().getFont().setBold(true);
            cel_loja.getFormat().getFill().setColor('e8e8e8');
            colInstore++;
        } else {
            // Colar na aba ONLINE de forma incremental
            w_online.getCell(0, colOnline).setValue(valorProduto);
            w_online.getCell(0, colOnline).getFormat().getFont().setBold(true);
            w_online.getCell(0, colOnline).getFormat().getFill().setColor('e8e8e8');

            let cel_sku2 = w_online.getCell(1, colOnline);
            cel_sku2.setValue('SKU');
            cel_sku2.getFormat().getFont().setBold(true);
            cel_sku2.getFormat().getFill().setColor('e8e8e8');
            colOnline++;
        }
        linhaAtual++;
    }
}
