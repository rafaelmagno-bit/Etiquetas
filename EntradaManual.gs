/**
 * Registra manualmente uma nova entrada na aba ENTRADA.
 * @param {Object} dados - Objeto contendo os dados do formulário
 * @return {Object} Status do salvamento
 */
function salvarEntradaManual(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaEntrada = ss.getSheetByName("ENTRADA");
    
    if (!abaEntrada) {
      return { sucesso: false, mensagem: "Aba 'ENTRADA' não encontrada." };
    }

    if (!dados.itens || dados.itens.length === 0) {
      return { sucesso: false, mensagem: "Nenhum item enviado para gravação." };
    }

    let linhasParaInserir = [];

    // Loop por todos os itens enviados
    for (let i = 0; i < dados.itens.length; i++) {
      let item = dados.itens[i];

      // Formatação individual do Tombamento do item
      let tomboFormatado = "";
      const qtd = parseInt(item.quantidade) || 1;
      const tomboIni = String(item.tombamentoInicial || "").trim();

      if (qtd > 1 && tomboIni) {
        const numIni = parseInt(tomboIni, 10);
        if (!isNaN(numIni)) {
          const numFim = numIni + qtd - 1;
          const tamOriginal = tomboIni.length;
          const tomboFimStr = String(numFim).padStart(tamOriginal, '0');
          tomboFormatado = `${tomboIni} a ${tomboFimStr}`;
        } else {
          tomboFormatado = tomboIni;
        }
      } else {
        tomboFormatado = tomboIni;
      }

      // Prepara o vetor com 36 posições
      let novaLinha = new Array(36).fill("");

      // Dados compartilhados do registro
      novaLinha[0]  = dados.sequencial || "";        // Coluna A  (1)  - Sequencial
      novaLinha[1]  = dados.dataRegistro || "";      // Coluna B  (2)  - Data do Registro
      novaLinha[3]  = dados.tipo || "";              // Coluna D  (4)  - Tipo
      novaLinha[5]  = dados.processo || "";          // Coluna F  (6)  - Processo
      novaLinha[6]  = dados.ug || "";                // Coluna G  (7)  - UG
      novaLinha[7]  = dados.uorg || "";              // Coluna H  (8)  - UORG
      novalinha[15] = (dados.corresponsavel || "").toUpperCase().trim();

      // Dados específicos deste item
      novaLinha[24] = item.descricao || "";         // Coluna Y  (25) - Descrição Completa
      novaLinha[25] = item.marca || "";             // Coluna Z  (26) - Marca
      novaLinha[27] = qtd;                           // Coluna AB (28) - Quantidade
      novaLinha[30] = tomboFormatado;                // Coluna AE (31) - Tombamento
      novaLinha[32] = "Pendente";                    // Coluna AG (33) - Status (Fixo)
      novaLinha[35] = item.descricaoPlaqueta || ""; // Coluna AJ (36) - Descrição para a Plaqueta

      linhasParaInserir.push(novaLinha);
    }

    // Grava todas as linhas em lote (operação mais rápida e segura)
    if (linhasParaInserir.length > 0) {
      const proximaLinha = abaEntrada.getLastRow() + 1;
      abaEntrada.getRange(proximaLinha, 1, linhasParaInserir.length, 36).setValues(linhasParaInserir);
    }

    return { 
      sucesso: true, 
      mensagem: `${dados.itens.length} item(ns) registrado(s) com sucesso na aba ENTRADA!` 
    };

  } catch (error) {
    return { sucesso: false, mensagem: "Erro ao salvar entrada manual: " + error.toString() };
  }
}

/**
 * Busca dinamica de UGs na aba "UG"
 * Coluna A = Codigo | Coluna B = Descricao | Coluna C = Sigla
 */
function buscarUgsNoBackend(termo) {
  if (!termo || termo.toString().trim().length === 0) return [];

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UG");
  if (!sheet) return [];

  var dados = sheet.getDataRange().getValues();
  var termoLower = termo.toString().toLowerCase().trim();
  var resultados = [];

  // Ignora o cabeçalho (começa em i = 1)
  for (var i = 1; i < dados.length; i++) {
    var codigo = dados[i][0] ? dados[i][0].toString().trim() : "";
    var descricao = dados[i][1] ? dados[i][1].toString().trim() : "";
    var sigla = dados[i][2] ? dados[i][2].toString().trim() : "";

    // Verifica se o termo bate com Código, Descrição OU Sigla
    if (
      codigo.toLowerCase().includes(termoLower) ||
      descricao.toLowerCase().includes(termoLower) ||
      sigla.toLowerCase().includes(termoLower)
    ) {
      resultados.push({
        codigo: codigo,
        descricao: descricao,
        sigla: sigla
      });
    }

    // Limita a 15 resultados para manter a interface rápida
    if (resultados.length >= 15) break;
  }

  return resultados;
}

/**
 * Busca dinâmica de UORGs na aba "ORGANOGRAMA"
 * Coluna A = Código SIPAC | Coluna B = Descrição | Coluna C = Código UORG | Coluna D = UG vinculada
 */
function buscarUorgsNoBackend(termo, codigoUg) {
  if (!termo || termo.toString().trim().length === 0) return [];

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORGANOGRAMA");
  if (!sheet) return [];

  var dados = sheet.getDataRange().getValues();
  var termoLower = termo.toString().toLowerCase().trim();
  var ugFiltro = codigoUg ? codigoUg.toString().trim().toLowerCase() : "";
  var resultados = [];

  // Ignora o cabeçalho (começa em i = 1)
  for (var i = 1; i < dados.length; i++) {
    var sipac = dados[i][0] ? dados[i][0].toString().trim() : "";
    var descricao = dados[i][1] ? dados[i][1].toString().trim() : "";
    var uorg = dados[i][2] ? dados[i][2].toString().trim() : "";
    var ugVinculada = dados[i][3] ? dados[i][3].toString().trim() : "";

    // Se uma UG foi selecionada no formulário, só exibe as UORGs associadas a ela
    if (ugFiltro && ugVinculada.toLowerCase() !== ugFiltro) {
      continue;
    }

    // Busca por Código SIPAC, Descrição OU Código UORG
    if (
      sipac.toLowerCase().includes(termoLower) ||
      descricao.toLowerCase().includes(termoLower) ||
      uorg.toLowerCase().includes(termoLower)
    ) {
      resultados.push({
        sipac: sipac,
        descricao: descricao,
        uorg: uorg,
        ug: ugVinculada
      });
    }

    // Limite de 15 resultados para performance
    if (resultados.length >= 15) break;
  }

  return resultados;
}