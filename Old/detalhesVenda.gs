// function createOrResetBlingDetalhesVendasSheet() {
//   const sheetName = "Detalhes Vendas";
//   const headersConfig = [
//     { name: "Venda ID", width: 90, align: "left", format: "0" },
//     { name: "Numero", width: 80, align: "left", format: "0" },
//     { name: "Numero Loja", width: 145, align: "left", format: "0" },
//     { name: "Data", width: 90, align: "center", format: "dd/MM/yyyy" },
//     { name: "Total Produtos", width: 120, align: "right", format: "R$ #,##0.00" },
//     { name: "Total", width: 120, align: "right", format: "R$ #,##0.00" },
//     { name: "Contato Nome", width: 300, align: "left" },
//     { name: "Item Produto ID", width: 130, align: "left" },
//     { name: "Item Quantidade", width: 130, align: "right", format: "0" },
//     { name: "Item Valor", width: 100, align: "right", format: "R$ #,##0.00" },
//     { name: "Situação", width: 175, align: "center" },
//     { name: "Loja", width: 175, align: "center" }
//   ];
//   const sheet = initializeSheet(sheetName, false, headersConfig, { autoFilter: true, frozenRows: 1 });
//   return sheet;
// }

// function importSalesDetails() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const vendasSheet = ss.getSheetByName("Vendas");
//   if (!vendasSheet) {
//     Logger.log(`❌ Planilha "Vendas" não encontrada.`);
//     return;
//   }

//   const detalhesSheet = createOrResetBlingDetalhesVendasSheet();

//   const prefix = "gsn";
//   const token = ensureValidBlingToken({ prefix, ...BLING_CONFIG[prefix] });
//   const baseUrl = "https://api.bling.com.br/Api/v3/pedidos/vendas/";

//   const lastRow = vendasSheet.getLastRow();
//   if (lastRow < 2) {
//     Logger.log(`❌ Nenhum ID de venda encontrado para processar.`);
//     return;
//   }

//   const idRange = vendasSheet.getRange(2, 1, lastRow - 1, 1); // IDs na Coluna A
//   const ids = idRange.getValues().flat().filter(id => id);

//   let row = 2;

//   ids.forEach((idVenda, index) => {
//     let attempts = 0;
//     const maxRetries = 5;
//     let success = false;

//     while (attempts < maxRetries && !success) {
//       attempts++;
//       const url = `${baseUrl}${idVenda}`;
//       Logger.log(`🔎 Tentativa ${attempts}/${maxRetries} para Venda ID ${idVenda} (${index + 1}/${ids.length})`);

//       const options = {
//         method: "get",
//         headers: {
//           "Authorization": `Bearer ${token}`,
//           "Accept": "application/json"
//         },
//         muteHttpExceptions: true
//       };

//       try {
//         const response = UrlFetchApp.fetch(url, options);
//         const status = response.getResponseCode();
//         const content = response.getContentText();

//         if (status === 429) {
//           Logger.log(`⚠️ Rate limit para Venda ID ${idVenda}. Aguardando...`);
//           Utilities.sleep(3000 * attempts);
//           continue;
//         }

//         if (status !== 200) {
//           Logger.log(`❌ Erro ${status} ao buscar venda ${idVenda}: ${content}`);
//           Utilities.sleep(1000 * attempts);
//           continue;
//         }

//         const json = JSON.parse(content);
//         const venda = json.data;
//         if (!venda || !venda.itens || venda.itens.length === 0) {
//           Logger.log(`⚠️ Venda ID ${idVenda} não possui itens.`);
//           break;
//         }

//         const situacaoTexto = SITUACAO_ENUM[venda.situacao?.id] || venda.situacao?.id || "";
//         const lojaTexto = LOJA_ENUM[venda.loja?.id] || venda.loja?.id || "";

//         const vendaInfo = [
//           venda.id || "",
//           venda.numero || "",
//           venda.numeroLoja || "",
//           venda.data || "",
//           venda.totalProdutos || 0,
//           venda.total || 0,
//           capitalizeName(venda.contato?.nome) || ""
//         ];

//         const outputData = venda.itens.map(item => ([
//           ...vendaInfo,
//           item.produto?.id || "",
//           item.quantidade || 0,
//           item.valor || 0,
//           situacaoTexto,
//           lojaTexto
//         ]));

//         detalhesSheet.getRange(row, 1, outputData.length, outputData[0].length).setValues(outputData);
//         SpreadsheetApp.flush();

//         Logger.log(`✅ Detalhes salvos para Venda ID ${idVenda} (linha ${row} em diante).`);
//         row += outputData.length;
//         success = true;

//       } catch (e) {
//         Logger.log(`❌ Erro ao processar venda ${idVenda} na tentativa ${attempts}: ${e.message}`);
//         Utilities.sleep(1000 * attempts);
//       }
//     }

//     if (!success) {
//       Logger.log(`❌ Falha definitiva após ${maxRetries} tentativas para venda ${idVenda}.`);
//     }
//   });

//   Logger.log(`✅ Importação de detalhes das vendas finalizada.`);
// }
