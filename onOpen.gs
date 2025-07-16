// Subir código atual para GAP: clasp push
// Baixar código atual de GAP: clasp pull

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GSN")

    .addItem("- 🔑 Autenticação Bling", "openAuthBlingAllAccounts")

    .addItem("1. Obter vendas de ontem", "load")
    .addItem("2. Obter vendas por data específica", "load2")
    .addToUi();
}

function load() {
  importDetalhesVenda2();
  importResumoProdutos();
  importResumoMarcas();
}

function load2() {
  promptImportarVendasPorData();
  importResumoProdutos();
  importResumoMarcas();
}