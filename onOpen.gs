// Subir código atual para GAP: clasp push
// Baixar código atual de GAP: clasp pull

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GSN")
    .addItem("- 🔑 Autenticação Bling", "openAuthBlingAllAccounts")
    .addItem("- Obter vendas de ontem", "importBlingSalesYesterday")
    .addToUi();
}