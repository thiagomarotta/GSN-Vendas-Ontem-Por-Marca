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
    { name: "Markup", width: 120, align: "center", format: "0.00" },
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

function promptImportarVendasPorData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instruções");
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert("❌ Aba 'Instruções' não encontrada.");
    return;
  }

  const dataInicialStr = sheet.getRange("F20").getDisplayValue().trim();
  const dataFinalStr = sheet.getRange("F21").getDisplayValue().trim();

  const parseDate = (str) => {
    const regex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    const match = str.match(regex);
    if (!match) return null;
    const [_, d, m, y] = match;
    return new Date(`${y}-${m}-${d}T00:00:00`);
  };

  const dataInicio = parseDate(dataInicialStr);
  const dataFim = parseDate(dataFinalStr);

  if (!dataInicio || !dataFim) {
    ui.alert("❌ Datas inválidas. Verifique o formato (DD/MM/AAAA) nas células F20 e F21 da aba Instruções.");
    return;
  }

  if (dataFim < dataInicio) {
    ui.alert("❌ A data final não pode ser anterior à data inicial.");
    return;
  }

  ui.alert(`📅 Importando vendas de ${dataInicialStr} até ${dataFinalStr}...`);
  importDetalhesVenda2(dataInicio, dataFim);
}


function importDetalhesVenda2(dataInicial, dataFinal) {
  const sheet = createOrResetDetalhesVenda2Sheet();
  const empresas = [
    { prefix: "gsn", configKey: "gsn" },
    { prefix: "metabolik", configKey: "metabolik" }
  ];

  const timezone = Session.getScriptTimeZone();
  const hoje = new Date();
  const ontem = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 1);

  const inicio = dataInicial || ontem;
  const fim = dataFinal || ontem;

  const dataInicialStr = Utilities.formatDate(inicio, timezone, "yyyy-MM-dd");
  const dataFinalStr = Utilities.formatDate(fim, timezone, "yyyy-MM-dd");

  for (const empresa of empresas) {
    importarVendasPorEmpresa(empresa.prefix, BLING_CONFIG[empresa.configKey], sheet, dataInicialStr, dataFinalStr);
  }
}

function importarVendasPorEmpresa(prefix, config, detalhesSheet, dataInicialStr, dataFinalStr) {
  const token = ensureValidBlingToken({ prefix, ...config });
  const baseVendaUrl = "https://api.bling.com.br/Api/v3/pedidos/vendas/";
  const baseProdutoUrl = "https://api.bling.com.br/Api/v3/produtos/";
  const maxRetries = 5;
  const limite = 100;

  let row = detalhesSheet.getLastRow() + 1;
  let pagina = 1;

  while (true) {
    const url = `${baseVendaUrl}?limit=${limite}&page=${pagina}&dataInicial=${dataInicialStr}&dataFinal=${dataFinalStr}`;
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
        vendaDetalhada.total || 0,
        capitalizeName(vendaDetalhada.contato?.nome || "")
      ];

      const situacao = SITUACAO_ENUM[vendaDetalhada.situacao?.id] || vendaDetalhada.situacao?.id;
      const loja = LOJA_ENUM[vendaDetalhada.loja?.id] || vendaDetalhada.loja?.id;
      const itens = vendaDetalhada.itens || [];

      const totalProdutosBruto = itens.reduce((sum, item) => sum + (item.valor || 0) * (item.quantidade || 0), 0);
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
}
