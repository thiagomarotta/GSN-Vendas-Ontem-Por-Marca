function createOrResetBlingDetalhesVendasSheet() {
  const sheetName = "Detalhes Vendas";
  const headersConfig = [
    { name: "Venda ID", width: 90, align: "left", format: "0" },                   // Equivalente ao campo "ID"
    { name: "Numero", width: 80, align: "left", format: "0" },                     // Mesmo de "Numero"
    { name: "Numero Loja", width: 145, align: "left", format: "0" },               // Igual ao "Numero Loja"
    { name: "Data", width: 90, align: "center", format: "dd/MM/yyyy" },            // Igual
    { name: "Total", width: 120, align: "right", format: "R$ #,##0.00" },          // Igual
    { name: "Total Produtos", width: 120, align: "right", format: "R$ #,##0.00" }, // Igual
    { name: "Contato Nome", width: 300, align: "left" },                           // Igual
    // { name: "Situação ID", width: 100, align: "center", format: "0" },             // Igual
    // { name: "Loja ID", width: 80, align: "center", format: "0" },                  // Igual
    // { name: "Item ID", width: 120, align: "left" },
    { name: "Item Produto ID", width: 130, align: "left" },
    // { name: "Item Código", width: 120, align: "left" },
    { name: "Item Quantidade", width: 130, align: "right", format: "0" },
    { name: "Item Valor", width: 100, align: "right", format: "R$ #,##0.00" },
    
  ];
  const sheet = initializeSheet(sheetName, false, headersConfig, { autoFilter: true, frozenRows: 1 });
  return sheet;
}


function importSalesDetails() {
  const vendasSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendas");
  if (!vendasSheet) {
    Logger.log(`❌ Planilha "Vendas" não encontrada.`);
    return;
  }

  const detalhesSheet = createOrResetBlingDetalhesVendasSheet();

  const prefix = "gsn"; // Ou 'metabolik'
  const token = ensureValidBlingToken({ prefix, ...BLING_CONFIG[prefix] });
  const baseUrl = "https://api.bling.com.br/Api/v3/pedidos/vendas/";

  const lastRow = vendasSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log(`❌ Nenhum ID de venda encontrado para processar.`);
    return;
  }

  const idRange = vendasSheet.getRange(2, 1, lastRow - 1, 1); // IDs na Coluna A
  const ids = idRange.getValues().flat().filter(id => id);

  let row = 2; // Começar na linha 2 da nova planilha

  ids.forEach((idVenda, index) => {
    let attempts = 0;
    const maxRetries = 5;
    let success = false;

    while (attempts < maxRetries && !success) {
      attempts++;
      const url = `${baseUrl}${idVenda}`;
      Logger.log(`🔎 Tentativa ${attempts}/${maxRetries} para Venda ID ${idVenda} (${index + 1}/${ids.length})`);

      const options = {
        method: "get",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Accept": "application/json"
        },
        muteHttpExceptions: true
      };

      try {
        const response = UrlFetchApp.fetch(url, options);
        const status = response.getResponseCode();
        const content = response.getContentText();

        if (status === 429) {
          Logger.log(`⚠️ Rate limit para Venda ID ${idVenda}. Aguardando...`);
          Utilities.sleep(3000 * attempts);
          continue;
        }

        if (status !== 200) {
          Logger.log(`❌ Erro ${status} ao buscar venda ${idVenda}: ${content}`);
          Utilities.sleep(1000 * attempts);
          continue;
        }

        const json = JSON.parse(content);
        const venda = json.data;
        if (!venda || !venda.itens || venda.itens.length === 0) {
          Logger.log(`⚠️ Venda ID ${idVenda} não possui itens.`);
          break;
        }

        const vendaInfo = [
          venda.id || "",
          venda.numero || "",
          venda.numeroLoja || "",
          venda.data || "",
          venda.total || 0,
          venda.totalProdutos || 0,
          venda.contato?.nome || ""
        ];

        const outputData = venda.itens.map(item => ([
          ...vendaInfo,
          item.produto?.id || "",    // Item Produto ID (nova posição correta)
          item.quantidade || 0,      // Item Quantidade
          item.valor || 0            // Item Valor
        ]));

        detalhesSheet.getRange(row, 1, outputData.length, outputData[0].length).setValues(outputData);
        SpreadsheetApp.flush();

        Logger.log(`✅ Detalhes salvos para Venda ID ${idVenda} (linha ${row} em diante).`);
        row += outputData.length;
        success = true;

      } catch (e) {
        Logger.log(`❌ Erro ao processar venda ${idVenda} na tentativa ${attempts}: ${e.message}`);
        Utilities.sleep(1000 * attempts);
      }
    }

    if (!success) {
      Logger.log(`❌ Falha definitiva após ${maxRetries} tentativas para venda ${idVenda}.`);
    }
  });

  Logger.log(`✅ Importação de detalhes das vendas finalizada.`);
}
