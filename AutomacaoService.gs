/**
 * MÓDULO: AUTOMAÇÃO E SINCRONIZAÇÃO
 * Sincroniza dados da planilha principal externa para a aba ENTRADA local.
 * Executado automaticamente via gatilho de tempo (Daily Trigger).
 */
function sincronizarPlanilhaPrincipal() {
  try {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    const abaEntradaLocal = ssLocal.getSheetByName("ENTRADA");
    if (!abaEntradaLocal) {
      Logger.log("Erro: Aba 'ENTRADA' local não encontrada.");
      return;
    }

    // 1. Abre a planilha principal externa usando o ID fornecido
    const idPlanilhaPrincipal = "1qrUFrqUnYvi51FoUfmAYdV2h6lOMh3fHM3-tjDz3Zyc";
    const ssPrincipal = SpreadsheetApp.openById(idPlanilhaPrincipal);
    const abaEntradaPrincipal = ssPrincipal.getSheetByName("ENTRADA");
    
    if (!abaEntradaPrincipal) {
      Logger.log("Erro: Aba de origem na planilha principal externa não localizada.");
      return;
    }

    // 2. Coleta os dados da planilha principal (Origem)
    const dadosPrincipal = abaEntradaPrincipal.getDataRange().getValues();
    if (dadosPrincipal.length <= 2) {
      Logger.log("Planilha principal sem registros novos para importar.");
      return;
    }

    const idxColunaAG = 32; // Coluna AG na Origem é a 33ª coluna (índice 32)
    const sequenciaisCancelados = new Set();

    // 🌟 PASSO A: Identifica todos os Sequenciais CANCELADOS na planilha externa
    for (let p = 2; p < dadosPrincipal.length; p++) {
      const linha = dadosPrincipal[p];
      const seq = String(linha[0] || "").trim();
      const statusBaixa = String(linha[idxColunaAG] || "").trim().toLowerCase();

      if (seq && statusBaixa === "cancelado") {
        sequenciaisCancelados.add(seq);
      }
    }

    // 🌟 PASSO B: Exclui da planilha LOCAL os registros que foram cancelados na Origem
    // Varre de baixo para cima para evitar desalinhamento de linhas durante a exclusão
    let dadosLocais = abaEntradaLocal.getDataRange().getValues();
    let contadorExcluidos = 0;

    for (let i = dadosLocais.length - 1; i >= 2; i--) {
      const seqLocal = String(dadosLocais[i][0] || "").trim();
      if (seqLocal && sequenciaisCancelados.has(seqLocal)) {
        abaEntradaLocal.deleteRow(i + 1); // +1 pois deleteRow usa índice de linha da planilha (inicia em 1)
        contadorExcluidos++;
      }
    }

    // 🌟 PASSO C: Mapeia os Sequenciais que PERMANECERAM na aba ENTRADA local
    dadosLocais = abaEntradaLocal.getDataRange().getValues(); // Recarrega os dados atualizados pós-exclusão
    const sequenciaisExistentes = {};
    for (let i = 2; i < dadosLocais.length; i++) {
      const seq = String(dadosLocais[i][0] || "").trim();
      if (seq) {
        sequenciaisExistentes[seq] = true;
      }
    }

    // Determina o tamanho horizontal fixo baseado no cabeçalho local
    const totalColunasLocal = abaEntradaLocal.getLastColumn(); 
    const colAG = 33; // Coluna 33 é a AG Local

    let contadorNovos = 0;

    // 🌟 PASSO D: Importa apenas os registros novos e que NÃO estejam cancelados
    for (let p = 2; p < dadosPrincipal.length; p++) {
      const linhaPrincipal = dadosPrincipal[p];
      const seqPrincipal = String(linhaPrincipal[0] || "").trim();
      const statusBaixa = String(linhaPrincipal[idxColunaAG] || "").trim().toLowerCase();

      // Ignora registros sem sequencial, cancelados na origem ou que já existem localmente
      if (!seqPrincipal || statusBaixa === "cancelado" || sequenciaisExistentes[seqPrincipal]) {
        continue;
      }

      // Cria uma nova linha para injetar respeitando o tamanho local
      let novaLinha = [];
      
      // Copia os dados originais respeitando o limite de colunas da planilha principal
      for (let c = 0; c < totalColunasLocal; c++) {
        novaLinha.push(linhaPrincipal[c] !== undefined ? linhaPrincipal[c] : "");
      }

      // Força a Coluna AG (Status do Fluxo Local) a ser "Pendente"
      novaLinha[colAG - 1] = "Pendente";

      // Garante que as colunas AH e AI entram vazias
      novaLinha[colAG] = "";     // AH (Agente Tombador)
      novaLinha[colAG + 1] = ""; // AI (Data Confirmação)

      // Adiciona a linha no final da aba local
      abaEntradaLocal.appendRow(novaLinha);
      sequenciaisExistentes[seqPrincipal] = true; // Marca no mapa para evitar duplicatas do mesmo loop
      contadorNovos++;
    }

    Logger.log(`Sincronização concluída com sucesso! novos adicionados: ${contadorNovos} | registros cancelados removidos: ${contadorExcluidos}.`);

  } catch (error) {
    Logger.log("Erro crítico na sincronização automatizada: " + error.toString());
  }
}


function atualizarItensDemanda() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEntrada = ss.getSheetByName("ENTRADA");
  var sheetItens = ss.getSheetByName("ITENS_DEMANDA");
  
  var dadosEntrada = sheetEntrada.getDataRange().getValues();
  var dadosItens = sheetItens.getDataRange().getValues();
  
  var tombosExistentes = {};
  for (var i = 1; i < dadosItens.length; i++) {
    var chaveUnica = dadosItens[i][1] + "_" + dadosItens[i][2]; // SEQUENCIAL_TOMBO
    tombosExistentes[chaveUnica] = true;
  }
  
  var novasLinhas = [];
  
  var colSequencial = 0;      // Coluna B ("SEQUENCIAL DOCUMENTO")
  var colDescricao = 35;     // Coluna Y ("DESCRIÇÃO")
  var colTombamentos = 30;    // Coluna AE ("TOMBAMENTO")
  var colStatusMae = 32;      // Coluna AG ("STATUS")
  
  Logger.log("Iniciando varredura com correção de intervalos...");
  
  for (var j = 1; j < dadosEntrada.length; j++) {
    var linha = dadosEntrada[j];
    var statusMae = String(linha[colStatusMae]).trim();
    
    if (statusMae === "Com Tombador") {
      var sequencial = linha[colSequencial];
      var descricao = linha[colDescricao];
      var stringTombos = String(linha[colTombamentos]).trim();
      
      if (!stringTombos || stringTombos === "" || stringTombos === "undefined") continue;
      
      // Padroniza o texto: transforma " a " ou " A " em um hífen "-" para facilitar a leitura do intervalo
      stringTombos = stringTombos.replace(/\s+a\s+/gi, "-");
      
      // Quebra o texto por vírgulas, ponto e vírgula, quebras de linha ou espaços restantes (sem quebrar números isolados)
      var partes = stringTombos.split(/[\n,;\s]+/);
      
      partes.forEach(function(parte) {
        parte = parte.trim();
        if (parte === "") return;
        
        // Verifica se a parte atual é um intervalo (ex: 16349952-16349953)
        if (parte.includes("-")) {
          var intervalo = parte.split("-");
          var inicio = parseInt(intervalo[0].trim(), 10);
          var fim = parseInt(intervalo[1].trim(), 10);
          
          // Se forem números válidos e o início for menor ou igual ao fim, gera a sequência
          if (!isNaN(inicio) && !isNaN(fim) && inicio <= fim) {
            for (var t = inicio; t <= fim; t++) {
              var tomboGerado = String(t);
              var chaveVerificacao = sequencial + "_" + tomboGerado;
              
              if (!tombosExistentes[chaveVerificacao]) {
                var idItem = "ID-" + Utilities.getUuid().substring(0, 8);
                novasLinhas.push([idItem, sequencial, tomboGerado, descricao, "", "", "Pendente"]);
              }
            }
          }
        } else {
          // Se for um número de tombo único isolado
          var tomboUnico = parte;
          var chaveVerificacao = sequencial + "_" + tomboUnico;
          
          if (!tombosExistentes[chaveVerificacao]) {
            var idItem = "ID-" + Utilities.getUuid().substring(0, 8);
            novasLinhas.push([idItem, sequencial, tomboUnico, descricao, "", "", "Pendente"]);
          }
        }
      });
    }
  }
  
  if (novasLinhas.length > 0) {
    sheetItens.getRange(sheetItens.getLastRow() + 1, 1, novasLinhas.length, 7).setValues(novasLinhas);
    Logger.log("SUCESSO: " + novasLinhas.length + " itens processados (incluindo intervalos)!");
  } else {
    Logger.log("Varredura concluída: Nenhum item novo.");
  }
}

/**
 * Verifica o status de vistoria dos itens detalhados na aba ITENS_DEMANDA.
 * Se TODOS os tombos de um sequencial que está "Com Tombador" na aba ENTRADA 
 * estiverem como "Etiqueta Fixada" ou "Etiquetas Entregues à Servidor",
 * o status da demanda na aba ENTRADA é alterado para "Concluído".
 */
function verificarEAtualizarStatusEntrada() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEntrada = ss.getSheetByName("ENTRADA");
  var sheetItens = ss.getSheetByName("ITENS_DEMANDA");
  
  var dadosEntrada = sheetEntrada.getDataRange().getValues();
  var dadosItens = sheetItens.getDataRange().getValues();
  
  // Mapeamento das posições das colunas na aba ENTRADA
  var colSequencialEntrada = 0; // Coluna A (Índice 0)
  var colStatusMae = 32;       // Coluna AG (Índice 32)
  
  // Mapeamento das posições das colunas na aba ITENS_DEMANDA
  var colSequencialItens = 1;  // Coluna B (Índice 1)
  var colStatusVistoria = 6;   // Coluna G (Índice 6)

  // 1. Agrupar os status de vistoria por Sequencial a partir da aba ITENS_DEMANDA
  // Estrutura: { "SEQUENCIAL": { total: X, concluidos: Y } }
  var statusPorSequencial = {};
  
  for (var i = 1; i < dadosItens.length; i++) {
    var seqItem = String(dadosItens[i][colSequencialItens]).trim();
    var statusVistoria = String(dadosItens[i][colStatusVistoria]).trim();
    
    if (seqItem === "" || seqItem === "undefined") continue;
    
    if (!statusPorSequencial[seqItem]) {
      statusPorSequencial[seqItem] = { total: 0, concluidos: 0 };
    }
    
    statusPorSequencial[seqItem].total++;
    
    // Verifica se atende a um dos dois critérios de conclusão
    if (statusVistoria === "Etiqueta Fixada" || statusVistoria === "Etiquetas Entregues à Servidor") {
      statusPorSequencial[seqItem].concluidos++;
    }
  }
  
  var alterouAlgo = false;

  // 2. Varrer a aba ENTRADA procurando por demandas "Com Tombador"
  for (var j = 1; j < dadosEntrada.length; j++) {
    var statusMae = String(dadosEntrada[j][colStatusMae]).trim();
    
    if (statusMae === "Com Tombador") {
      var sequencialMae = String(dadosEntrada[j][colSequencialEntrada]).trim();
      var infoItens = statusPorSequencial[sequencialMae];
      
      // Se existem registros gerados para este sequencial e TODOS estão com o status correto
      if (infoItens && infoItens.total > 0 && infoItens.total === infoItens.concluidos) {
        
        // Linha física na planilha (j + 1 por causa do índice 0 do array e cabeçalho)
        var numeroLinha = j + 1; 
        
        // Coluna AG é a 33ª coluna do Excel/Sheets
        sheetEntrada.getRange(numeroLinha, 33).setValue("Concluído");
        
        Logger.log("Demanda " + sequencialMae + " atualizada para Concluído. (" + infoItens.concluidos + " de " + infoItens.total + " itens validados).");
        alterouAlgo = true;
      }
    }
  }
  
  if (!alterouAlgo) {
    Logger.log("Nenhuma demanda elegível para conclusão no momento.");
  }
}

/**
 * Função auxiliar para normalizar o CADMAT:
 * Remove caracteres não numéricos (pontos, traços) e zeros à esquerda.
 */
function normalizarCadmat(valor) {
  if (valor === null || valor === undefined) return "";
  var str = String(valor).trim();
  if (str === "") return "";
  
  // Extrai apenas os números
  var apenasNumeros = str.replace(/\D/g, ""); 
  if (apenasNumeros === "") return str;
  
  // Remove zeros à esquerda
  var semZeros = apenasNumeros.replace(/^0+/, "");
  return semZeros === "" ? "0" : semZeros;
}

/**
 * Busca o CADMAT (Coluna X) na aba ENTRADA, localiza a correspondência na planilha externa
 * (Coluna A -> Coluna I) e grava a descrição diretamente na Coluna AJ.
 */
function atualizarDescricaoTombamentoPorCadmat() {
  try {
    var ID_PLANILHA_EXTERNA = "1yNiDK7LbRWf0U8wNQc8vUMZqF9rD5w8AuETswlYs_14";
    var GID_ABA_EXTERNA = 1690895173;

    // 1. Abrir planilha externa e carregar a aba
    var ssExterna = SpreadsheetApp.openById(ID_PLANILHA_EXTERNA);
    var abasExternas = ssExterna.getSheets();
    var abaExterna = abasExternas.find(s => s.getSheetId() === GID_ABA_EXTERNA) || abasExternas[0];
    
    var dadosExternos = abaExterna.getDataRange().getValues();
    
    // Mapa em memória: CADMAT Normalizado -> Descrição (Coluna I / índice 8)
    var mapaCadmat = {};
    for (var i = 1; i < dadosExternos.length; i++) {
      var cadmatRaw = dadosExternos[i][0]; // Coluna A (CADMAT)
      var cadmatNorm = normalizarCadmat(cadmatRaw);
      var descricao = String(dadosExternos[i][8] || "").trim(); // Coluna I
      
      if (cadmatNorm !== "") {
        mapaCadmat[cadmatNorm] = descricao;
      }
    }

    // 2. Abrir aba ENTRADA da planilha ativa
    var ssAtiva = SpreadsheetApp.getActiveSpreadsheet();
    var abaEntrada = ssAtiva.getSheetByName("ENTRADA");
    
    if (!abaEntrada) {
      return { sucesso: false, mensagem: "Aba 'ENTRADA' não foi encontrada." };
    }

    var dadosEntrada = abaEntrada.getDataRange().getValues();
    if (dadosEntrada.length <= 1) {
      return { sucesso: true, mensagem: "Não há dados para processar na aba ENTRADA." };
    }

    // DEFINIÇÃO EXATA DAS COLUNAS:
    // Coluna X  = índice 23 (base 0) -> CADMAT
    // Coluna AJ = índice 35 (base 0) -> Descrição para Tombamento (Coluna 36 na planilha)
    var idxCadmat = 23; 
    var idxDescricaoTombamento = 35; // Coluna AJ

    // Tenta validar o cabeçalho; se não achar palavra similar, mantém a Coluna AJ (35)
    var cabecalho = dadosEntrada[0].map(c => String(c).trim().toLowerCase());
    for (var col = 0; col < cabecalho.length; col++) {
      if (cabecalho[col].includes("descriç") && (cabecalho[col].includes("tombamento") || cabecalho[col].includes("tombo"))) {
        idxDescricaoTombamento = col;
        break;
      }
    }

    var totalEncontrados = 0;
    var totalAtualizados = 0;
    var colunaResultado = [];

    // 3. Cruzar informações linha a linha
    for (var r = 1; r < dadosEntrada.length; r++) {
      var cadmatLinhaRaw = dadosEntrada[r][idxCadmat];
      var cadmatLinhaNorm = normalizarCadmat(cadmatLinhaRaw);
      var descAtual = dadosEntrada[r][idxDescricaoTombamento];

      if (cadmatLinhaNorm !== "" && mapaCadmat[cadmatLinhaNorm] !== undefined) {
        var novaDesc = mapaCadmat[cadmatLinhaNorm];
        colunaResultado.push([novaDesc]);
        totalEncontrados++;
        
        if (descAtual !== novaDesc) {
          totalAtualizados++;
        }
      } else {
        // Mantém a descrição existente se não encontrar o CADMAT na busca
        colunaResultado.push([descAtual || ""]);
      }
    }

    // 4. Gravar os dados de uma vez na Coluna AJ (Coluna 36)
    if (colunaResultado.length > 0) {
      abaEntrada.getRange(2, idxDescricaoTombamento + 1, colunaResultado.length, 1)
        .setValues(colunaResultado);
    }

    return { 
      sucesso: true, 
      mensagem: `Sincronização concluída na Coluna AJ!\n\n` +
                `- Correspondências encontradas: ${totalEncontrados}\n` +
                `- Linhas recém-atualizadas: ${totalAtualizados}` 
    };

  } catch (erro) {
    return { sucesso: false, mensagem: "Erro durante a sincronização: " + erro.toString() };
  }
}