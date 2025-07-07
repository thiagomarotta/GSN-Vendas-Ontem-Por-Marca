function createOrResetResumoMarcasSheet() {
  const sheetName = "Marcas";
  const headersConfig = [
    { name: "Marca", width: 200, align: "left" },
    { name: "Receita Total", width: 150, align: "right", format: "R$ #,##0.00" },
    { name: "Custo Total", width: 150, align: "right", format: "R$ #,##0.00" },
    { name: "Markup", width: 100, align: "center", format: "0.00" }
  ];

  return initializeSheet(sheetName, false, headersConfig, {
    autoFilter: true,
    frozenRows: 1
  });
}

function importResumoMarcas() {
  const detalhesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendas");
  if (!detalhesSheet) {
    Logger.log("❌ Planilha 'Detalhes Venda #2' não encontrada.");
    return;
  }

  const data = detalhesSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("⚠️ Nenhum dado para processar.");
    return;
  }

  const resumoMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
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

  if (output.length > 0) {
    resumoSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
    Logger.log(`✅ Resumo de marcas gerado com ${output.length} marcas.`);
  } else {
    Logger.log("⚠️ Nenhum dado encontrado para gerar o resumo de marcas.");
  }
}
