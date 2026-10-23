/**
 * MÓDULO: CONFIGURAÇÃO GLOBAL
 * Gerencia as requisições iniciais e renderiza a interface do Web App.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Fixação de Etiquetas — CBMAC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Permite acoplamento no AppSheet ou sistemas web
}

/**
 * Função utilitária para incluir arquivos HTML (Styles e Scripts) dentro da casca principal (Index).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}