function copiarAbaProdutos() {
  // ID das planilhas de origem e de destino
  const planilhaOrigemId = PropertiesService.getScriptProperties().getProperty('planilhaOrigemId')
  const planilhaDestinoId = PropertiesService.getScriptProperties().getProperty('planilhaDestinoId')

  // Aba de origem e de destino
  const abaOrigemNome = 'Produtos';
  const abaDestinoNome = 'Produtos';

  // Acessando as planilhas
  const planilhaOrigem = SpreadsheetApp.openById(planilhaOrigemId);
  const planilhaDestino = SpreadsheetApp.openById(planilhaDestinoId);

  const abaOrigem = planilhaOrigem.getSheetByName(abaOrigemNome);
  const abaDestino = planilhaDestino.getSheetByName(abaDestinoNome);

  // Verifica se as abas existem
  if (abaOrigem && abaDestino) {
    const dadosAbaOrigem = abaOrigem.getDataRange().getValues();

    // Apagando dados existentes na aba de destino (opcional)
    abaDestino.clear();

    // Copiando os dados para a aba de destino
    abaDestino.getRange(1, 1, dadosAbaOrigem.length, dadosAbaOrigem[0].length).setValues(dadosAbaOrigem);

    Logger.log('A aba "Produtos" foi copiada com sucesso para a planilha de destino.');
  } else {
    Logger.log('Não foi possível encontrar as abas especificadas.');
  }
}
