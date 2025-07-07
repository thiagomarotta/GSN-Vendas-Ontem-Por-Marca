// function createOrResetBlingProdutosVendasOntemSheet() {
//   const sheetName = "Produtos Vendas";
//   const headersConfig = [
//     { name: "ID", width: 90, align: "left", format: "0" },
//     { name: "Código", width: 90, align: "left" },
//     { name: "EAN", width: 110, align: "left" },
//     { name: "Produto", width: 570, align: "left" },
//     { name: "Marca", width: 120, align: "left" },
//     { name: "Tipo", width: 55, align: "center" },
//     // { name: "Situação", width: 85, align: "center" },
//     { name: "Estoque", width: 100, align: "right", format: "0" },
//     { name: "PDV", width: 100, align: "right", format: "R$ #,##0.00" },
//     { name: "PDC", width: 100, align: "right", format: "R$ #,##0.00" }
//   ];
//   const sheet = initializeSheet(sheetName, false, headersConfig, { autoFilter: true, frozenRows: 1 });
//   return sheet;
// }

// function importSalesProducts() {
//   const detalhesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Detalhes Vendas");
//   if (!detalhesSheet) {
//     Logger.log(`❌ Planilha "Detalhes Vendas" não encontrada.`);
//     return;
//   }

//   const produtosSheet = createOrResetBlingProdutosVendasOntemSheet();

//   const prefix = "gsn"; // Ou 'metabolik'
//   const token = ensureValidBlingToken({ prefix, ...BLING_CONFIG[prefix] });
//   const baseUrl = "https://api.bling.com.br/Api/v3/produtos/";

//   const lastRow = detalhesSheet.getLastRow();
//   if (lastRow < 2) {
//     Logger.log(`❌ Nenhum produto encontrado na planilha de detalhes.`);
//     return;
//   }

//   // Agora buscando IDs dos produtos da coluna H = posição 8 = "Item Produto ID"
//   const idRange = detalhesSheet.getRange(2, 8, lastRow - 1, 1);
//   const produtoIds = [...new Set(idRange.getValues().flat().filter(id => id))];  // Remove duplicados e vazios

//   Logger.log(`🔎 Total de IDs únicos de produtos para buscar: ${produtoIds.length}`);

//   let row = 2;  // Começar a preencher na linha 2 da aba "Produtos Vendas"

//   produtoIds.forEach((idProduto, index) => {
//     let attempts = 0;
//     const maxRetries = 5;
//     let success = false;

//     while (attempts < maxRetries && !success) {
//       attempts++;
//       const url = `${baseUrl}${idProduto}`;
//       Logger.log(`🔎 Tentativa ${attempts}/${maxRetries} para Produto ID ${idProduto} (${index + 1}/${produtoIds.length})`);

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
//           Logger.log(`⚠️ Rate limit ao buscar Produto ID ${idProduto}. Aguardando...`);
//           Utilities.sleep(3000 * attempts);
//           continue;
//         }

//         if (status !== 200) {
//           Logger.log(`❌ Erro ${status} ao buscar produto ${idProduto}: ${content}`);
//           Utilities.sleep(1000 * attempts);
//           continue;
//         }

//         const json = JSON.parse(content);
//         const prod = json.data;
//         if (!prod) {
//           Logger.log(`⚠️ Produto ID ${idProduto} retornou vazio.`);
//           break;
//         }

//         const values = [[
//           prod.id || "",
//           prod.codigo || "",
//           prod.gtin || "",
//           prod.nome || "",
//           prod.marca || "",
//           prod.tipo || "",
//           // prod.situacao || "",
//           prod.estoque?.saldoVirtualTotal || 0,
//           prod.preco || 0,
//           prod.fornecedor?.precoCusto || 0
//         ]];

//         produtosSheet.getRange(row, 1, 1, values[0].length).setValues(values);
//         SpreadsheetApp.flush();

//         Logger.log(`✅ Produto ID ${idProduto} salvo (linha ${row}).`);
//         row++;
//         success = true;

//       } catch (e) {
//         Logger.log(`❌ Erro ao processar Produto ID ${idProduto} na tentativa ${attempts}: ${e.message}`);
//         Utilities.sleep(1000 * attempts);
//       }
//     }

//     if (!success) {
//       Logger.log(`❌ Falha definitiva após ${maxRetries} tentativas para Produto ID ${idProduto}.`);
//     }
//   });

//   Logger.log(`✅ Importação de produtos das vendas finalizada.`);
// }
