function createOrResetReceitaPorMarcaSheet() {
  const sheetName = "Receita Marca";
  const headersConfig = [
    { name: "Marca", width: 250, align: "left" },
    { name: "Receita Total", width: 150, align: "right", format: "R$ #,##0.00" },
    { name: "Custo Total", width: 150, align: "right", format: "R$ #,##0.00" },
    { name: "Markup", width: 100, align: "center", format: "#,##0.00" }
  ];
  const sheet = initializeSheet(sheetName, false, headersConfig, { autoFilter: true, frozenRows: 1 });
  return sheet;
}

function importVendasOntemPorMarca() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const vendasSheet = ss.getSheetByName("Vendas");
  const detalhesSheet = ss.getSheetByName("Detalhes Vendas");
  const produtosSheet = ss.getSheetByName("Produtos Vendas");

  if (!vendasSheet || !detalhesSheet || !produtosSheet) {
    Logger.log("❌ Planilhas necessárias não encontradas.");
    return;
  }

  const outputSheet = createOrResetReceitaPorMarcaSheet();

  // --- Mapear Venda ID → Total Final (coluna F da Vendas)
  const vendasData = vendasSheet.getRange(2, 1, vendasSheet.getLastRow() - 1, 6).getValues();
  const vendaIdToTotalFinal = {};
  vendasData.forEach(row => {
    const id = row[0];
    const total = parseFloat((row[5] || "0").toString().replace(/[^\d,.-]/g, "").replace(",", "."));
    if (id) vendaIdToTotalFinal[id] = total;
  });

  // --- Mapear Produto ID → Marca e Custo Unitário (PDC na Coluna 9, index 8)
  const produtosData = produtosSheet.getRange(2, 1, produtosSheet.getLastRow() - 1, 9).getValues();
  const produtoIdToMarca = {};
  const produtoIdToCustoUnitario = {};
  produtosData.forEach(row => {
    const id = row[0];
    const marca = row[4] || "Sem Marca";
    let custo = row[8] || 0;

    if (typeof custo === "string") {
      custo = parseFloat(custo.replace(/[^\d,.-]/g, "").replace(",", "."));
    } else if (typeof custo !== "number") {
      custo = 0;
    }

    if (id) {
      produtoIdToMarca[id] = marca;
      produtoIdToCustoUnitario[id] = isNaN(custo) ? 0 : custo;
    }
  });

  // --- Mapear Venda ID → Itens (quantidade + valor)
  const detalhesData = detalhesSheet.getRange(2, 1, detalhesSheet.getLastRow() - 1, 10).getValues();
  const vendaIdToItens = {};

  detalhesData.forEach(row => {
    const vendaId = row[0];
    const produtoId = row[7];  // Coluna H = index 7 (Item Produto ID)
    const quantidade = parseFloat((row[8] || "0").toString().replace(/[^\d,.-]/g, "").replace(",", "."));
    const valorOriginal = parseFloat((row[9] || "0").toString().replace(/[^\d,.-]/g, "").replace(",", "."));

    if (vendaId && produtoId) {
      if (!vendaIdToItens[vendaId]) vendaIdToItens[vendaId] = [];
      vendaIdToItens[vendaId].push({
        produtoId: produtoId,
        valorOriginal: isNaN(valorOriginal) ? 0 : valorOriginal,
        quantidade: isNaN(quantidade) ? 0 : quantidade
      });
    }
  });

  // --- Calcular Receita Final + Custo por Marca
  const receitaPorMarca = {};
  const custoPorMarca = {};

  Object.keys(vendaIdToItens).forEach(vendaId => {
    const totalFinal = vendaIdToTotalFinal[vendaId] || 0;
    const itens = vendaIdToItens[vendaId];
    const totalOriginalItens = itens.reduce((sum, item) => sum + item.valorOriginal, 0);

    itens.forEach(item => {
      let valorFinal = 0;
      const custoUnitario = produtoIdToCustoUnitario[item.produtoId] || 0;
      const custoTotalItem = custoUnitario * item.quantidade;

      if (totalOriginalItens > 0) {
        const percentual = item.valorOriginal / totalOriginalItens;
        valorFinal = percentual * totalFinal;
      }

      const marca = produtoIdToMarca[item.produtoId] || "Sem Marca";

      if (!receitaPorMarca[marca]) receitaPorMarca[marca] = 0;
      if (!custoPorMarca[marca]) custoPorMarca[marca] = 0;

      receitaPorMarca[marca] += valorFinal;
      custoPorMarca[marca] += custoTotalItem;
    });
  });

  // --- Gerar saída com Markup calculado ---
  const outputData = Object.keys(receitaPorMarca)
    .map(marca => {
      const receita = receitaPorMarca[marca];
      const custo = custoPorMarca[marca];
      const markup = custo > 0 ? receita / custo : "";
      return [
        marca,
        receita,
        custo,
        typeof markup === "number" ? Number(markup.toFixed(2)) : ""
      ];
    })
    .sort((a, b) => b[1] - a[1]);

  if (outputData.length > 0) {
    outputSheet.getRange(2, 1, outputData.length, 4).setValues(outputData);
    SpreadsheetApp.flush();
    Logger.log(`✅ Receita, custo e markup por marca calculados com sucesso.`);
  } else {
    Logger.log(`⚠️ Nenhuma receita encontrada para calcular.`);
  }
}
