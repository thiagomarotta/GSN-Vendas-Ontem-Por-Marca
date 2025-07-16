function createOrResetResumoMarcasSheet() {
  const sheetName = "Marcas";
  const headersConfig = [
    { name: "Marca", width: 200, align: "left" },
    { name: "Receita Total", width: 150, align: "right", format: formatContabilidadeBR },
    { name: "Custo Total", width: 150, align: "right", format: formatContabilidadeBR },
    { name: "Markup", width: 100, align: "center", format: "0.00" }
  ];

  return initializeSheet(sheetName, false, headersConfig, {
    autoFilter: true,
    frozenRows: 1
  });
}

function importResumoMarcas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const instrucoesSheet = ss.getSheetByName("Instruções");
  const detalhesSheet = ss.getSheetByName("Vendas");

  if (!instrucoesSheet || !detalhesSheet) {
    Logger.log("❌ Planilha 'Instruções' ou 'Vendas' não encontrada.");
    return;
  }

  const filtro = instrucoesSheet.getRange("F2").getDisplayValue().trim();
  const somentePagas = filtro === "Somente compras pagas";

  const data = detalhesSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("⚠️ Nenhum dado para processar.");
    return;
  }

  const resumoMap = new Map();
  let receitaTotalGeral = 0;
  let custoTotalGeral = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const situacao = (row[15] || "").toString().trim(); // Coluna "Situação"
    if (somentePagas && !statusPagos.includes(situacao)) {
      continue;
    }

    const marca = row[8] || "Sem Marca";
    const valorTotal = Number(row[12]) || 0;
    const custoTotal = Number(row[13]) || 0;

    if (!resumoMap.has(marca)) {
      resumoMap.set(marca, {
        marca,
        totalReceita: 0,
        totalCusto: 0
      });
    }

    const item = resumoMap.get(marca);
    item.totalReceita += valorTotal;
    item.totalCusto += custoTotal;

    receitaTotalGeral += valorTotal;
    custoTotalGeral += custoTotal;
  }

  const resumoSheet = createOrResetResumoMarcasSheet();
  const output = [];

  for (const [_, dados] of resumoMap) {
    const markup = dados.totalCusto > 0
      ? dados.totalReceita / dados.totalCusto
      : "";

    output.push([
      dados.marca,
      dados.totalReceita,
      dados.totalCusto,
      markup
    ]);
  }

  output.sort((a, b) => b[1] - a[1]); // ordena por receita decrescente

  // Adiciona linha de TOTAL
  const markupGeral = custoTotalGeral > 0 ? receitaTotalGeral / custoTotalGeral : "";
  const totalRow = ["TOTAL", receitaTotalGeral, custoTotalGeral, markupGeral];
  output.push(totalRow);

  if (output.length > 0) {
    const range = resumoSheet.getRange(2, 1, output.length, output[0].length);
    range.setValues(output);

    // Aplica negrito na última linha (TOTAL)
    const totalRange = resumoSheet.getRange(output.length + 1, 1, 1, output[0].length);
    totalRange.setFontWeight("bold");

    Logger.log(`✅ Resumo de marcas gerado com ${output.length - 1} marcas + TOTAL (Filtro: ${filtro || "Nenhum"}).`);
  } else {
    Logger.log("⚠️ Nenhum dado encontrado para gerar o resumo de marcas.");
  }
}
