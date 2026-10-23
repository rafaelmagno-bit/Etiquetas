/**
 * Busca dados para o Relatório Fotográfico cruzando ITENS_DEMANDA com ENTRADA.
 */
function obterDadosRelatorioFotografico(tipoFiltro, valorFiltro) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaItens = ss.getSheetByName("ITENS_DEMANDA");
    var abaEntrada = ss.getSheetByName("ENTRADA");

    if (!abaItens || !abaEntrada) {
      return { sucesso: false, message: "Abas necessárias (ITENS_DEMANDA / ENTRADA) não encontradas." };
    }

    var dadosItens = abaItens.getDataRange().getValues();
    var dadosEntrada = abaEntrada.getDataRange().getValues();

    // Mapeia ENTRADA (Coluna A/B: Sequencial, Coluna F: Processo)
    var mapaProcessoPorSeq = {};
    for (var i = 1; i < dadosEntrada.length; i++) {
      var seqEntrada = String(dadosEntrada[i][0] || dadosEntrada[i][1] || "").trim();
      var numProcesso = String(dadosEntrada[i][5] || "").trim(); // Coluna F
      if (seqEntrada) {
        mapaProcessoPorSeq[seqEntrada] = numProcesso;
      }
    }

    var termoBusca = String(valorFiltro).trim().toLowerCase();
    var resultados = [];

    // Percorre ITENS_DEMANDA (Col B: Sequencial, Col C: Tombo, Col D: Descrição, Col E: Foto Plaqueta, Col F: Foto Bem)
    for (var j = 1; j < dadosItens.length; j++) {
      var seqItem = String(dadosItens[j][1] || "").trim();      // Coluna B
      var tomboItem = String(dadosItens[j][2] || "").trim();    // Coluna C
      var descItem = String(dadosItens[j][3] || "").trim();     // Coluna D
      var caminhoPlaqueta = String(dadosItens[j][4] || "").trim(); // Coluna E
      var caminhoBem = String(dadosItens[j][5] || "").trim();      // Coluna F

      var numProcesso = mapaProcessoPorSeq[seqItem] || "Não Identificado";

      var atendeFiltro = false;

      if (tipoFiltro === "sequencial" && seqItem.toLowerCase() === termoBusca) {
        atendeFiltro = true;
      } else if (tipoFiltro === "tombo" && tomboItem.toLowerCase() === termoBusca) {
        atendeFiltro = true;
      } else if (tipoFiltro === "processo" && numProcesso.toLowerCase().includes(termoBusca)) {
        atendeFiltro = true;
      }

      if (atendeFiltro) {
        resultados.push({
          sequencial: seqItem,
          processo: numProcesso,
          tombo: tomboItem || "Sem Tombo",
          descricao: descItem || "Sem Descrição",
          fotoPlaquetaBase64: buscarImagemDriveEConverterBase64(caminhoPlaqueta),
          fotoBemBase64: buscarImagemDriveEConverterBase64(caminhoBem)
        });
      }
    }

    if (resultados.length === 0) {
      return { sucesso: false, message: "Nenhum patrimônio encontrado para os critérios informados." };
    }

    return { sucesso: true, dados: resultados };

  } catch (err) {
    return { sucesso: false, message: "Erro ao buscar dados do relatório: " + err.toString() };
  }
}

/**
 * Busca o arquivo no Google Drive seja por URL, ID ou pelo Nome do Arquivo (AppSheet)
 * e o converte para String Base64 para inclusão direta no PDF.
 */
function buscarImagemDriveEConverterBase64(caminhoOuUrl) {
  if (!caminhoOuUrl) return null;
  var str = String(caminhoOuUrl).trim();
  if (!str) return null;

  try {
    var file = null;

    // 1. Caso seja uma URL do Google Drive
    if (str.indexOf("http") !== -1) {
      var match = str.match(/[\w-]{25,}/);
      if (match) {
        try { file = DriveApp.getFileById(match[0]); } catch (e) {}
      }
    }

    // 2. Caso seja diretamente o ID do arquivo (sem barras ou pontos)
    if (!file && str.length >= 25 && str.indexOf("/") === -1 && str.indexOf(".") === -1) {
      try { file = DriveApp.getFileById(str); } catch (e) {}
    }

    // 3. Caso seja o caminho do AppSheet (ex: "15812641/ID-3cdc8274.FOTO DA PLAQUETA.144230.jpg")
    if (!file) {
      var nomeArquivo = str.split('/').pop().trim();
      var buscaArquivos = DriveApp.getFilesByName(nomeArquivo);
      if (buscaArquivos.hasNext()) {
        file = buscaArquivos.next();
      }
    }

    // Se encontrou o arquivo, gera a string Base64
    if (file) {
      var blob = file.getBlob();
      var bytes = blob.getBytes();
      var encoded = Utilities.base64Encode(bytes);
      return "data:" + blob.getContentType() + ";base64," + encoded;
    }

  } catch (e) {
    Logger.log("Erro ao carregar imagem do Drive: " + e.toString());
  }
  
  return null;
}