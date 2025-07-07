function createOrResetBlingVendasSheet() {
  const sheetName = "Vendas";
  const headersConfig = [
    { name: "ID", width: 90, align: "left", format: "0" },
    { name: "Numero", width: 80, align: "left", format: "0" },
    { name: "Numero Loja", width: 145, align: "left", format: "0" },
    { name: "Data", width: 90, align: "center", format: "dd/MM/yyyy" },
    { name: "Total Produtos", width: 120, align: "right", format: "R$ #,##0.00" },
    { name: "Total", width: 120, align: "right", format: "R$ #,##0.00" },
    { name: "Contato Nome", width: 300, align: "left" },
    { name: "Tipo Pessoa", width: 105, align: "center" },
    { name: "Numero Documento", width: 155, align: "center" },
    { name: "Situação", width: 175, align: "center" },
    { name: "Loja", width: 175, align: "center" }
  ];

  const sheet = initializeSheet(sheetName, false, headersConfig, { autoFilter: true, frozenRows: 1 });
  return sheet;
}

function applyBlingVendasFormatting(sheet, headersConfig) {
  headersConfig.forEach((config, i) => {
    const col = i + 1;
    const colLetter = getColumnLetter(col);

    if (config.width) sheet.setColumnWidth(col, config.width);
    if (config.align) sheet.getRange(`${colLetter}:${colLetter}`).setHorizontalAlignment(config.align);
    if (config.format) sheet.getRange(`${colLetter}:${colLetter}`).setNumberFormat(config.format);
  });
}

function importBlingSalesYesterday() {
  const sheet = createOrResetBlingVendasSheet();

  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const startDate = Utilities.formatDate(yesterday, "GMT-3", "yyyy-MM-dd");
  const endDate = startDate;

  const prefix = "gsn";
  const token = ensureValidBlingToken({ prefix, ...BLING_CONFIG[prefix] });

  const baseUrl = "https://api.bling.com.br/Api/v3/pedidos/vendas";
  let page = 1;
  const limit = 5;
  let allSales = [];
  let retryCount = 0;
  const maxRetries = 5;

  while (true) {
    const url = `${baseUrl}?dataInicial=${startDate}&dataFinal=${endDate}&limit=${limit}&page=${page}`;
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
      const json = JSON.parse(response.getContentText());

      if (status === 429) {
        Logger.log(`⚠️ Rate limit na página ${page}. Retentando após delay...`);
        Utilities.sleep(3000 * (retryCount + 1));
        retryCount++;
        if (retryCount > maxRetries) throw new Error("Excesso de retries por rate limit.");
        continue;
      }

      if (status !== 200) {
        Logger.log(`Erro ${status}: ${response.getContentText()}`);
        throw new Error(`Erro ${status}: ${response.getContentText()}`);
      }

      const vendas = json.data || [];
      const pagination = json.pagination || { currentPage: page, totalPages: page };

      allSales = allSales.concat(vendas);
      Logger.log(`✅ Página ${page}: ${vendas.length} vendas carregadas.`);

      if (pagination.currentPage >= pagination.totalPages) {
        Logger.log(`✅ Fim da paginação na página ${pagination.currentPage} de ${pagination.totalPages}.`);
        break;
      }

      page++;
      retryCount = 0;

    } catch (e) {
      Logger.log(`❌ Erro na página ${page}: ${e.message}`);
      throw e;
    }
  }

  if (allSales.length === 0) {
    sheet.getRange("A1").setValue("Nenhuma venda encontrada para ontem.");
    return;
  }

  const outputData = allSales.map(sale => {
    const situacaoId = sale.situacao?.id;
    const situacaoText = SITUACAO_ENUM[situacaoId] || `Outro (${situacaoId || ""})`;

    const lojaId = sale.loja?.id;
    const lojaText = LOJA_ENUM[lojaId] || `Outro (${lojaId || ""})`;

    return [
      sale.id || "",
      sale.numero || "",
      sale.numeroLoja || "",
      sale.data || "",
      sale.totalProdutos || 0,
      sale.total || 0,
      capitalizeName(sale.contato?.nome) || "",
      sale.contato?.tipoPessoa || "",
      formatCpfCnpj(sale.contato?.numeroDocumento) || "",
      situacaoText,
      lojaText
    ];
  });

  sheet.getRange(2, 1, outputData.length, 11).setValues(outputData);
  SpreadsheetApp.flush();

  Logger.log(`✅ Importação finalizada: ${allSales.length} vendas no total.`);
}
