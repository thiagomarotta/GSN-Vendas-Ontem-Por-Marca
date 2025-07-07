function createOrResetResumoProdutosSheet() {
  const sheetName = "Produtos";


  const headersConfig = [
    { name: "ID", width: 110, align: "left" },
    { name: "Produto", width: 450, align: "left" },
    { name: "Marca", width: 190, align: "left" },
    { name: "Vendas", width: 125, align: "center", format: "0" },
    { name: "PDV Total", width: 120, align: "right", format: formatContabilidadeBR },
    { name: "PDC Total", width: 120, align: "right", format: formatContabilidadeBR },
    { name: "PDV Médio", width: 120, align: "right", format: formatContabilidadeBR },
    { name: "PDC Médio", width: 120, align: "right", format: formatContabilidadeBR },
    { name: "Markup", width: 125, align: "center", format: "0.00" },
    { name: "Estoque Atual", width: 150, align: "center", format: "0" }
  ];

  return initializeSheet(sheetName, false, headersConfig, {
    autoFilter: true,
    frozenRows: 1
  });
}


function importResumoProdutos() {
  const detalhesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendas");
  if (!detalhesSheet) {
    Logger.log(`❌ Planilha "Vendas" não encontrada.`);
    return;
  }

  const data = detalhesSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log(`⚠️ Nenhum dado para processar.`);
    return;
  }

  const resumoMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const produtoId = row[6]; // "Item Produto ID"
    const produtoNome = row[7]; // "Produto"
    const marca = row[8]; // "Marca"
    const quantidade = Number(row[9]) || 0;
    const valorTotal = Number(row[12]) || 0; // "Valor Total"
    const custoTotal = Number(row[13]) || 0; // "Custo Total"
    const estoqueAtual = Number(row[17]) || 0;

    if (!produtoId) continue;

    if (!resumoMap.has(produtoId)) {
      resumoMap.set(produtoId, {
        produtoId,
        produtoNome,
        marca,
        totalQuantidade: 0,
        totalValor: 0,
        totalCusto: 0,
        estoqueAtual
      });
    }

    const item = resumoMap.get(produtoId);
    item.totalQuantidade += quantidade;
    item.totalValor += valorTotal;
    item.totalCusto += custoTotal;

    if (estoqueAtual > item.estoqueAtual) {
      item.estoqueAtual = estoqueAtual;
    }
  }

  const resumoSheet = createOrResetResumoProdutosSheet();
  const output = [];

  for (const [id, dados] of resumoMap) {
    const markup = dados.totalCusto > 0
      ? dados.totalValor / dados.totalCusto
      : '';

    const pdvMedio = dados.totalQuantidade > 0
      ? dados.totalValor / dados.totalQuantidade
      : '';

    const pdcMedio = dados.totalQuantidade > 0
      ? dados.totalCusto / dados.totalQuantidade
      : '';

    output.push([
      dados.produtoId,
      dados.produtoNome,
      dados.marca,
      dados.totalQuantidade,
      dados.totalValor,
      dados.totalCusto,
      pdvMedio,
      pdcMedio,
      markup,
      dados.estoqueAtual
    ]);
  }

  if (output.length > 0) {
    resumoSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
    Logger.log(`✅ Resumo gerado com ${output.length} produtos.`);
  } else {
    Logger.log("⚠️ Nenhum dado encontrado para gerar resumo.");
  }
}
