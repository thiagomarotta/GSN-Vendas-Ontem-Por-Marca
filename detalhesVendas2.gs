function createOrResetDetalhesVenda2Sheet() {
  const sheetName = "Vendas";
  const headersConfig = [
    { name: "Venda ID", width: 130, align: "left", format: "0" },
    { name: "Numero", width: 110, align: "left", format: "0" },
    { name: "Numero Loja", width: 155, align: "left" },
    { name: "Data", width: 95, align: "center", format: "dd/MM/yyyy" },
    { name: "Total", width: 115, align: "right", format: formatContabilidadeBR },
    { name: "Contato Nome", width: 250, align: "left" },
    { name: "Item Produto ID", width: 130, align: "left" },
    { name: "Produto", width: 300, align: "left" },
    { name: "Marca", width: 160, align: "left" },
    { name: "Item Quantidade", width: 130, align: "center", format: "0" },
    { name: "Item Valor", width: 110, align: "right", format: formatContabilidadeBR },
    { name: "Item Custo", width: 110, align: "right", format: formatContabilidadeBR },
    { name: "Valor Total", width: 120, align: "right", format: formatContabilidadeBR },
    { name: "Custo Total", width: 120, align: "right", format: formatContabilidadeBR },
    { name: "Markup", width: 120, align: "center", format: "0.00" }, // ← ADICIONADO AQUI
    { name: "Situação", width: 170, align: "center" },
    { name: "Loja", width: 150, align: "center" },
    { name: "Estoque Atual", width: 130, align: "right", format: "0" }
  ];

  const sheet = initializeSheet(sheetName, false, headersConfig, {
    autoFilter: true,
    frozenRows: 1
  });

  return sheet;
}



// function importDetalhesVenda2() {
//   const startTime = Date.now();
//   const vendasSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendas");
//   if (!vendasSheet) {
//     Logger.log(`❌ Planilha "Vendas" não encontrada.`);
//     return;
//   }

//   const detalhesSheet = createOrResetDetalhesVenda2Sheet();
//   const token = ensureValidBlingToken({ prefix: "gsn", ...BLING_CONFIG["gsn"] });
//   const baseVendaUrl = "https://api.bling.com.br/Api/v3/pedidos/vendas/";
//   const baseProdutoUrl = "https://api.bling.com.br/Api/v3/produtos/";

//   const ids = vendasSheet.getRange(2, 1, vendasSheet.getLastRow() - 1, 1).getValues().flat().filter(Boolean);
//   let row = 2;

//   for (let i = 0; i < ids.length; i++) {
//     const idVenda = ids[i];
//     let venda;
//     let attempts = 0;
//     const maxRetries = 5;

//     while (attempts < maxRetries) {
//       try {
//         const res = UrlFetchApp.fetch(`${baseVendaUrl}${idVenda}`, {
//           method: "get",
//           headers: {
//             "Authorization": `Bearer ${token}`,
//             "Accept": "application/json"
//           },
//           muteHttpExceptions: true
//         });

//         const status = res.getResponseCode();

//         if (status === 429) {
//           Logger.log(`⚠️ Rate limit atingido para venda ${idVenda}. Tentativa ${attempts + 1}/${maxRetries}`);
//           Utilities.sleep(3000 * (attempts + 1));
//           attempts++;
//           continue;
//         }

//         if (status !== 200) {
//           Logger.log(`❌ Erro ${status} ao buscar venda ${idVenda}: ${res.getContentText()}`);
//           break;
//         }

//         venda = JSON.parse(res.getContentText()).data;
//         break; // sucesso

//       } catch (e) {
//         Logger.log(`❌ Erro na tentativa ${attempts + 1} ao buscar venda ${idVenda}: ${e.message}`);
//         Utilities.sleep(3000 * (attempts + 1));
//         attempts++;
//       }
//     }

//     if (!venda) {
//       Logger.log(`❌ Falha definitiva ao buscar venda ${idVenda}.`);
//       continue;
//     }

//     const vendaInfo = [
//       venda.id || "",
//       venda.numero || "",
//       venda.numeroLoja || "",
//       venda.data || "",
//       venda.totalProdutos || 0,
//       venda.total || 0,
//       capitalizeName(venda.contato?.nome || "")
//     ];

//     const situacao = SITUACAO_ENUM[venda.situacao?.id] || "Desconhecida";
//     const loja = LOJA_ENUM[venda.loja?.id] || "Desconhecida";

//     const itens = venda.itens || [];
//     const outputData = [];

//     for (let item of itens) {
//       const prodId = item.produto?.id || "";
//       let produtoNome = "";
//       let marca = "";
//       let estoque = 0;
//       let custo = 0;

//       if (prodId) {
//         let prodAttempts = 0;
//         while (prodAttempts < maxRetries) {
//           try {
//             const prodRes = UrlFetchApp.fetch(`${baseProdutoUrl}${prodId}`, {
//               method: "get",
//               headers: {
//                 "Authorization": `Bearer ${token}`,
//                 "Accept": "application/json"
//               },
//               muteHttpExceptions: true
//             });

//             const prodStatus = prodRes.getResponseCode();

//             if (prodStatus === 429) {
//               Logger.log(`⚠️ Rate limit ao buscar produto ${prodId}. Tentativa ${prodAttempts + 1}/${maxRetries}`);
//               Utilities.sleep(3000 * (prodAttempts + 1));
//               prodAttempts++;
//               continue;
//             }

//             if (prodStatus !== 200) {
//               Logger.log(`❌ Erro ${prodStatus} ao buscar produto ${prodId}: ${prodRes.getContentText()}`);
//               break;
//             }

//             const prod = JSON.parse(prodRes.getContentText()).data;
//             produtoNome = prod.nome || "";
//             marca = prod.marca || "";
//             estoque = prod.estoque?.saldoVirtualTotal || 0;
//             custo = prod.fornecedor?.precoCusto || 0;
//             break; // sucesso

//           } catch (e) {
//             Logger.log(`❌ Erro ao buscar produto ${prodId} na tentativa ${prodAttempts + 1}: ${e.message}`);
//             Utilities.sleep(3000 * (prodAttempts + 1));
//             prodAttempts++;
//           }
//         }
//       }

//       const quantidade = item.quantidade || 0;
//       const valor = item.valor || 0;
//       const valorTotal = quantidade * valor;
//       const custoTotal = quantidade * custo;

//       outputData.push([
//         ...vendaInfo,
//         prodId,
//         produtoNome,
//         marca,
//         quantidade,
//         valor,
//         custo,
//         valorTotal,
//         custoTotal,
//         situacao,
//         loja,
//         estoque
//       ]);
//     }

//     if (outputData.length > 0) {
//       detalhesSheet.getRange(row, 1, outputData.length, outputData[0].length).setValues(outputData);
//       row += outputData.length;
//       SpreadsheetApp.flush();
//       Logger.log(`✅ Processada venda ${idVenda} com ${outputData.length} itens.`);
//     }
//   }

//   // Medição de tempo de execução
//   const elapsed = (Date.now() - startTime) / 1000; // segundos
//   const hours = Math.floor(elapsed / 3600);
//   const minutes = Math.floor((elapsed % 3600) / 60);
//   const seconds = Math.floor(elapsed % 60);
//   const timeStr = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
//   SpreadsheetApp.getUi().alert(`⏱️ Tempo total de execução: ${timeStr}`);


//   Logger.log("✅ Importação de Detalhes Venda #2 finalizada.");
// }

function importDetalhesVenda2() {
  const detalhesSheet = createOrResetDetalhesVenda2Sheet();
  const token = ensureValidBlingToken({ prefix: "gsn", ...BLING_CONFIG["gsn"] });
  const baseVendaUrl = "https://api.bling.com.br/Api/v3/pedidos/vendas/";
  const baseProdutoUrl = "https://api.bling.com.br/Api/v3/produtos/";
  const maxRetries = 5;
  const limite = 100;

  // Define data de ontem
  const timezone = Session.getScriptTimeZone();
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const dataStr = Utilities.formatDate(yesterday, timezone, "yyyy-MM-dd");

  let row = 2;
  let pagina = 1;
  const startTime = new Date();

  while (true) {
    const url = `${baseVendaUrl}?limit=${limite}&page=${pagina}&dataInicial=${dataStr}&dataFinal=${dataStr}`;
    let response, status, json;

    try {
      response = UrlFetchApp.fetch(url, {
        method: "get",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Accept": "application/json"
        },
        muteHttpExceptions: true
      });
      status = response.getResponseCode();

      if (status === 429) {
        Logger.log(`⚠️ Rate limit na página ${pagina}. Aguardando...`);
        Utilities.sleep(3000 * pagina);
        continue;
      }

      if (status !== 200) {
        Logger.log(`❌ Erro ${status} ao buscar página ${pagina}: ${response.getContentText()}`);
        break;
      }

      json = JSON.parse(response.getContentText());
    } catch (e) {
      Logger.log(`❌ Erro na requisição da página ${pagina}: ${e.message}`);
      break;
    }

    const vendas = json.data || [];
    if (vendas.length === 0) break;

    Logger.log(`📦 Página ${pagina} carregou ${vendas.length} vendas`);

    for (let venda of vendas) {
      const vendaId = venda.id;
      let vendaDetalhada;
      let attempts = 0;

      // Retry para venda detalhada
      while (attempts < maxRetries) {
        try {
          const res = UrlFetchApp.fetch(`${baseVendaUrl}${vendaId}`, {
            method: "get",
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/json"
            },
            muteHttpExceptions: true
          });

          const status = res.getResponseCode();

          if (status === 429) {
            Logger.log(`⚠️ Rate limit para venda ${vendaId}. Tentativa ${attempts + 1}/${maxRetries}`);
            Utilities.sleep(3000 * (attempts + 1));
            attempts++;
            continue;
          }

          if (status !== 200) {
            Logger.log(`❌ Erro ${status} ao buscar venda ${vendaId}: ${res.getContentText()}`);
            break;
          }

          vendaDetalhada = JSON.parse(res.getContentText()).data;
          break;
        } catch (e) {
          Logger.log(`❌ Erro na tentativa ${attempts + 1} ao buscar venda ${vendaId}: ${e.message}`);
          Utilities.sleep(3000 * (attempts + 1));
          attempts++;
        }
      }

      if (!vendaDetalhada) continue;

      const vendaInfo = [
        vendaDetalhada.id || "",
        vendaDetalhada.numero || "",
        vendaDetalhada.numeroLoja || "",
        vendaDetalhada.data || "",
        // vendaDetalhada.totalProdutos || 0,
        vendaDetalhada.total || 0,
        capitalizeName(vendaDetalhada.contato?.nome || "")
      ];

      const situacao = SITUACAO_ENUM[vendaDetalhada.situacao?.id] || "Desconhecida";
      const loja = LOJA_ENUM[vendaDetalhada.loja?.id] || "Desconhecida";
      const itens = vendaDetalhada.itens || [];

      const totalProdutosBruto = itens.reduce((sum, item) =>
        sum + (item.valor || 0) * (item.quantidade || 0), 0);

      const outputData = [];

      for (let item of itens) {
        const prodId = item.produto?.id || "";
        let produtoNome = "", marca = "", estoque = 0, custo = 0;

        if (prodId) {
          let prodAttempts = 0;
          while (prodAttempts < maxRetries) {
            try {
              const prodRes = UrlFetchApp.fetch(`${baseProdutoUrl}${prodId}`, {
                method: "get",
                headers: {
                  "Authorization": `Bearer ${token}`,
                  "Accept": "application/json"
                },
                muteHttpExceptions: true
              });

              const prodStatus = prodRes.getResponseCode();
              if (prodStatus === 429) {
                Logger.log(`⚠️ Rate limit ao buscar produto ${prodId}. Tentativa ${prodAttempts + 1}/${maxRetries}`);
                Utilities.sleep(3000 * (prodAttempts + 1));
                prodAttempts++;
                continue;
              }

              if (prodStatus !== 200) {
                Logger.log(`❌ Erro ${prodStatus} ao buscar produto ${prodId}: ${prodRes.getContentText()}`);
                break;
              }

              const prod = JSON.parse(prodRes.getContentText()).data;
              produtoNome = prod.nome || "";
              marca = prod.marca || "";
              estoque = prod.estoque?.saldoVirtualTotal || 0;
              custo = prod.fornecedor?.precoCusto || 0;
              break;
            } catch (e) {
              Logger.log(`❌ Erro ao buscar produto ${prodId}: ${e.message}`);
              Utilities.sleep(3000 * (prodAttempts + 1));
              prodAttempts++;
            }
          }
        }

        const quantidade = item.quantidade || 0;
        const valorUnitarioBruto = item.valor || 0;
        const valorTotalBruto = valorUnitarioBruto * quantidade;

        const proporcao = totalProdutosBruto ? (valorTotalBruto / totalProdutosBruto) : 0;
        const valorTotalAjustado = proporcao * vendaDetalhada.total;
        const valorUnitarioAjustado = quantidade ? (valorTotalAjustado / quantidade) : 0;
        const custoTotal = custo * quantidade;
        const markup = custoTotal > 0 ? (valorTotalAjustado / custoTotal) : "";


        outputData.push([
          ...vendaInfo,
          prodId,
          produtoNome,
          marca,
          quantidade,
          valorUnitarioAjustado,
          custo,
          valorTotalAjustado,
          custoTotal,
          markup,
          situacao,
          loja,
          estoque
        ]);
      }

      if (outputData.length > 0) {
        detalhesSheet.getRange(row, 1, outputData.length, outputData[0].length).setValues(outputData);
        row += outputData.length;
        SpreadsheetApp.flush();
        Logger.log(`✅ Processada venda ${vendaId} com ${outputData.length} itens.`);
      }
    }

    if (!json.pagination || pagina >= json.pagination.totalPages) {
      Logger.log("✅ Fim da paginação.");
      break;
    }

    pagina++;
  }

  const endTime = new Date();
  const durationMs = endTime - startTime;
  const hours = Math.floor(durationMs / (1000 * 60 * 60));
  const minutes = Math.floor((durationMs % (1000 * 60 * 60)) / (1000 * 60));
  const seconds = Math.floor((durationMs % (1000 * 60)) / 1000);

  const durationStr = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
  // SpreadsheetApp.getUi().alert(`✅ Importação de Detalhes Venda #2 finalizada.\n⏱ Tempo total: ${durationStr}`);
}
