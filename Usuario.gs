// Retorna o e-mail do usuário logado na conta Google
function obterUsuarioLogado() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return "usuario@dominio.com"; // Fallback para testes locais
  }
}