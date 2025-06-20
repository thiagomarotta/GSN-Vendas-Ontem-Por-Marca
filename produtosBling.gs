function createOrResetBlingProdutosSheet() {
  const sheetName = "Produtos";
  const headersConfig = [
    { name: "ID", width: 80, align: "left", format: "0" },
    { name: "Código", width: 120, align: "left" },
    { name: "EAN", width: 150, align: "left" },
    { name: "Produto", width: 300, align: "left" },
    { name: "Marca", width: 150, align: "left" },
    { name: "Tipo", width: 80, align: "center" },
    { name: "Situação", width: 80, align: "center" },
    { name: "Estoque", width: 100, align: "right", format: "0" },
    { name: "PDV", width: 100, align: "right", format: "R$ #,##0.00" },
    { name: "PDC", width: 100, align: "right", format: "R$ #,##0.00" }
  ];
  const sheet = initializeSheet(sheetName, false, headersConfig, { autoFilter: true, frozenRows: 1 });
  return sheet;
}

function importBlingProducts() {
  const sheet = createOrResetBlingProdutosSheet();
  const prefix = "gsn"; // Ou 'metabolik' se quiser usar outra conta
  const token = ensureValidBlingToken({ prefix, ...BLING_CONFIG[prefix] });
  const baseUrl = "https://api.bling.com.br/Api/v3/produtos";
  const limite = 100;
  let pagina = 1;
  let row = 2;  // Começar na segunda linha da planilha
  let retryCount = 0;
  const maxRetries = 5;
  let totalProdutos = 0;

  while (true) {
    const url = `${baseUrl}?limite=${limite}&pagina=${pagina}`;
    Logger.log(`🔎 Fazendo request para página ${pagina}: ${url}`);

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
      const json = JSON.parse(content);

      if (status === 429) {
        Logger.log(`⚠️ Rate limit na página ${pagina}. Retentando após delay...`);
        Utilities.sleep(3000 * (retryCount + 1));
        retryCount++;
        if (retryCount > maxRetries) throw new Error("Excesso de retries por rate limit.");
        continue;
      }

      if (status !== 200) {
        Logger.log(`❌ Erro ${status}: ${content}`);
        throw new Error(`Erro ${status}: ${content}`);
      }

      const produtos = json.data || [];
      Logger.log(`📄 Página ${pagina} retornou ${produtos.length} produtos.`);

      if (produtos.length === 0) {
        Logger.log(`✅ Fim da paginação: Nenhum produto retornado na página ${pagina}.`);
        break;
      }

      const outputData = produtos.map(prod => ([
        prod.id || ""
        // prod.codigo || "",
        // prod.nome || "",
        // prod.tipo || "",
        // prod.preco || 0,
        // prod.estoque?.saldoVirtualTotal || 0,
        // prod.unidade || "",
        // prod.gtin || "",
        // prod.categoria?.descricao || ""
      ]));

      sheet.getRange(row, 1, outputData.length, 1).setValues(outputData);
      row += outputData.length;
      totalProdutos += produtos.length;
      SpreadsheetApp.flush();

      Logger.log(`📈 Total acumulado até agora: ${totalProdutos} produtos.`);

      // Se retornou menos que o limite, paramos
      if (produtos.length < limite) {
        Logger.log(`✅ Fim da paginação: Página ${pagina} tinha apenas ${produtos.length} produtos (limite é ${limite}).`);
        break;
      }

      pagina++;
      retryCount = 0;

    } catch (e) {
      Logger.log(`❌ Erro ao processar página ${pagina}: ${e.message}`);
      throw e;
    }
  }

  Logger.log(`✅ Importação de produtos finalizada com sucesso. Total final: ${totalProdutos} produtos.`);
}

function importAdditionalProductInformation() {
  const sheetName = "Produtos";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`❌ Planilha "${sheetName}" não encontrada.`);
    return;
  }

  const prefix = "gsn"; // Ou 'metabolik' se quiser outra conta
  const token = ensureValidBlingToken({ prefix, ...BLING_CONFIG[prefix] });
  const baseUrl = "https://api.bling.com.br/Api/v3/produtos/";

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log(`❌ Nenhum ID encontrado na planilha para processar.`);
    return;
  }

  const idRange = sheet.getRange(2, 1, lastRow - 1, 2); // Colunas A (ID) e B (Código)
  const idAndCodigoValues = idRange.getValues();

  const idsToProcess = idAndCodigoValues
    .map((row, idx) => ({ id: row[0], codigo: row[1], rowIndex: idx + 2 }))
    .filter(item => item.id && !item.codigo);  // Só os que têm ID mas não têm Código preenchido

  Logger.log(`🔎 Total de produtos faltando dados: ${idsToProcess.length}`);

  idsToProcess.forEach((item, index) => {
    let attempts = 0;
    const maxRetries = 5;
    let success = false;

    while (attempts < maxRetries && !success) {
      attempts++;
      const url = `${baseUrl}${item.id}`;
      Logger.log(`🔎 Tentativa ${attempts}/${maxRetries} para ID ${item.id} (linha ${item.rowIndex})`);

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
          Logger.log(`⚠️ Rate limit ao buscar o ID ${item.id}. Aguardando antes de retentar...`);
          Utilities.sleep(3000 * attempts);
          continue;
        }

        if (status !== 200) {
          Logger.log(`❌ Erro ${status} ao buscar ID ${item.id}: ${content}`);
          Utilities.sleep(1000 * attempts);
          continue;
        }

        const json = JSON.parse(content);
        const prod = json.data;

        if (!prod) {
          Logger.log(`⚠️ Produto ID ${item.id} retornou vazio.`);
          break;
        }

        const values = [
          prod.codigo || "",
          prod.gtin || "",
          prod.nome || "",
          prod.marca || "",
          prod.tipo || "",
          prod.situacao || "",
          prod.estoque?.saldoVirtualTotal || 0,
          prod.preco || 0,
          prod.fornecedor?.precoCusto || 0
        ];

        // Preenchendo colunas B até J (colunas 2 até 10)
        sheet.getRange(item.rowIndex, 2, 1, values.length).setValues([values]);
        SpreadsheetApp.flush();

        Logger.log(`✅ Dados preenchidos para ID ${item.id} (linha ${item.rowIndex}).`);
        success = true;

      } catch (e) {
        Logger.log(`❌ Erro ao processar ID ${item.id} na tentativa ${attempts}: ${e.message}`);
        Utilities.sleep(1000 * attempts);
      }
    }

    if (!success) {
      Logger.log(`❌ Falha definitiva após ${maxRetries} tentativas para o ID ${item.id} (linha ${item.rowIndex}).`);
    }
  });

  Logger.log(`✅ Importação de informações adicionais finalizada.`);
}
