// Subir código atual para GAP: clasp push
// Baixar código atual de GAP: clasp pull

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GSN")

    .addItem("- 🔑 Autenticação Bling", "openAuthBlingAllAccounts")

    .addItem("1. Obter vendas de ontem e seus detalhes", "importDetalhesVenda2")
    .addItem("2. Sumarizar vendas por produto", "importResumoProdutos")
    .addItem("3. Sumarizar vendas por marca", "importResumoMarcas")

    .addToUi();
}

function load() {
  importDetalhesVenda2();
  importResumoProdutos();
  importResumoMarcas();
}