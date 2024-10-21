function main(workbook: ExcelScript.Workbook) {
    // CRIAR OS PRODUTOS NAS ABAS ONLINE E INSTORE

    // Acessa a planilha 'INPUT_INICIAL'
    let w_inputInicial = workbook.getWorksheet("OVERVIEW");

    // Determinar a última linha preenchida na coluna H
    let ultimaLinhaInicial = w_inputInicial.getCell(w_inputInicial.getRange().getRowCount() - 1, 7)
        .getRangeEdge(ExcelScript.KeyboardDirection.up).getRowIndex();

    // Acessar as planilhas 'INSTORE' e 'ONLINE'
    let w_instore = workbook.getWorksheet('INSTORE');
    let w_online = workbook.getWorksheet('ONLINE');

    let rangeLimpar = w_instore.getRange("A1:U1");
    rangeLimpar.clear();
    let rangeLimpar2 = w_online.getRange("A1:U40");
    rangeLimpar2.clear();

    // Definir variáveis para controle das linhas onde os dados serão colados
    let colInstore = 0;
    let colOnline = 0;


    // Loop para verificar cada linha da coluna H e copiar para a planilha correspondente
    for (let linhaAtual = 1; linhaAtual <= ultimaLinhaInicial;) {
        // Obter o valor da coluna H (index 7) na linha atual
        let valorColunaH = w_inputInicial.getCell(linhaAtual, 7).getValue() as string;

        // Obter o valor da célula ao lado da coluna H (index 8, coluna I)
        let valorProduto = w_inputInicial.getCell(linhaAtual, 8).getValue();


        if (valorColunaH == "Instore") {
            // Colar na aba INSTORE de forma incremental
            w_instore.getCell(0, colInstore).setValue(valorProduto);

            let cel_sku = w_instore.getCell(1, colInstore);
            cel_sku.setValue('SKU');
            cel_sku.getFormat().getFont().setBold(true);
            colInstore++;

            w_instore.getCell(0, colInstore).setValue(valorProduto);

            let cel_loja = w_instore.getCell(1, colInstore);
            cel_loja.setValue('LOJA');
            cel_loja.getFormat().getFont().setBold(true);
            colInstore++;
        }
        else {
            // Colar na aba ONLINE de forma incremental
            w_online.getCell(0, colOnline).setValue(valorProduto);

            let cel_sku2 = w_online.getCell(1, colOnline);
            cel_sku2.setValue('SKU');
            cel_sku2.getFormat().getFont().setBold(true);
            colOnline++;
        }
        linhaAtual++;


    }
}
