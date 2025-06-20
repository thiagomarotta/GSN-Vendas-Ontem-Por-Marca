// Subir código atual para GAP: clasp push
// Baixar código atual de GAP: clasp pull

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GSN")
    .addItem("- 🔑 Autenticação Bling", "openAuthBlingAllAccounts")
    .addItem("1. Obter vendas de ontem", "importBlingSalesYesterday")
    .addItem("2. Obter detalhes de vendas do Bling", "importSalesDetails")
    .addItem("3. Obter produtos das vendas do Bling", "importSalesProducts")
    .addItem("4. Sumarize as vendas por marca", "importVendasOntemPorMarca")
    .addItem("- Load all data", "load")

    // .addItem("- Obter produtos do Bling", "importBlingProducts")
    // .addItem("- Obter dados complementares de produtos do Bling", "importAdditionalProductInformation")
    
    .addToUi();
}

function load() {
  importBlingSalesYesterday();
  importSalesDetails();
  importSalesProducts();
  importVendasOntemPorMarca();
}