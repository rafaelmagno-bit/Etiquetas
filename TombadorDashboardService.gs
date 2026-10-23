/**
 * MÓDULO: PAINEL GESTOR - DETALHAMENTO POR TOMBADOR
 * Consolida os dados das abas ENTRADA e ITENS_DEMANDA com foco nos status "Com Tombador" e "Concluido".
 */
function obterDadosPainelGestorTombadores() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaEntrada = ss.getSheetByName("ENTRADA");
    const abaItens = ss.getSheetByName("ITENS_DEMANDA");
    
    if (!abaEntrada || !abaItens) {
      return { sucesso: false, message: "Abas necessárias não encontradas." };
    }

    const dadosEntrada = abaEntrada.getDataRange().getValues();
    const dadosItens = abaItens.getDataRange().getValues();

    const resumoVisitasGlobal = {
      fixada: 0,
      entregue: 0,
      naoLocalizado: 0,
      naoComporta: 0,
      pendente: 0,
      listaNaoLocalizados: []
    };

    const progressoItensPorSequencial = {};
    
    // 1. Mapear o progresso na aba ITENS_DEMANDA
    for (let j = 1; j < dadosItens.length; j++) {
      const seqItem = String(dadosItens[j][1]).trim();   // Coluna B (Sequencial)
      const tombo = String(dadosItens[j][2]).trim();     // Coluna C (Tombo)
      const descricao = String(dadosItens[j][3]).trim(); // Coluna D (Descrição)
      const statusVistoria = String(dadosItens[j][6] || "").trim(); // Coluna G (STATUS)
      const statusVistoriaLimpo = statusVistoria.toLowerCase();
      
      if (!seqItem) continue;
      
      if (!progressoItensPorSequencial[seqItem]) {
        progressoItensPorSequencial[seqItem] = { 
          total: 0, 
          realizados: 0, 
          descricoes: [] 
        };
      }
      
      progressoItensPorSequencial[seqItem].total++;
      
      if (descricao && !progressoItensPorSequencial[seqItem].descricoes.includes(descricao)) {
        progressoItensPorSequencial[seqItem].descricoes.push(descricao);
      }
      
      // 🌟 REGRA CORRIGIDA: Apenas vistorias efetivamente finalizadas somam na barra de progresso
      const eConcluido = statusVistoriaLimpo === "etiqueta fixada" ||
                         statusVistoriaLimpo.includes("entregue") ||
                         statusVistoriaLimpo.includes("servidor") ||
                         statusVistoriaLimpo.includes("nao localizado") ||
                         statusVistoriaLimpo.includes("não localizado") ||
                         statusVistoriaLimpo.includes("nao comporta") ||
                         statusVistoriaLimpo.includes("não comporta");

      if (eConcluido) {
        progressoItensPorSequencial[seqItem].realizados++;
      }

      // Incrementa os contadores macros dos cards inferiores
      if (statusVistoriaLimpo === "etiqueta fixada") {
        resumoVisitasGlobal.fixada++;
      } else if (statusVistoriaLimpo.includes("entregue") || statusVistoriaLimpo.includes("servidor")) {
        resumoVisitasGlobal.entregue++;
      } else if (statusVistoriaLimpo.includes("nao localizado") || statusVistoriaLimpo.includes("não localizado")) {
        resumoVisitasGlobal.naoLocalizado++;
        resumoVisitasGlobal.listaNaoLocalizados.push({
          sequencial: seqItem,
          tombo: tombo,
          descricao: descricao
        });
      } else if (statusVistoriaLimpo.includes("nao comporta") || statusVistoriaLimpo.includes("não comporta")) {
        resumoVisitasGlobal.naoComporta++;
      } else {
        resumoVisitasGlobal.pendente++; // "Aguardando Vistoria", "Pendente" ou vazios
      }
    }

    // 2. Processar a aba ENTRADA e agrupar por Tombador
    const painelTombadores = {};
    const resumoGeral = { totalComTombador: 0, totalConcluido: 0, totalBensAlocados: 0 };

    const colSequencial = 0;   // Coluna A
    const colQtdBens = 27;     // Coluna AB
    const colStatus = 32;      // Coluna AG
    const colTombador = 33;    // Coluna AH
    const colDataEnvio = 34;   // Coluna AI
    const colDescricao = 35;   // Coluna AJ

    for (let i = 2; i < dadosEntrada.length; i++) {
      const linha = dadosEntrada[i];
      if (!linha[colSequencial]) continue;

      const statusMacro = String(linha[colStatus]).trim();
      if (statusMacro !== "Com Tombador" && statusMacro !== "Concluido") continue;

      const tombadorNome = String(linha[colTombador] || "Não Atribuído").trim();
      const sequencial = String(linha[colSequencial]).trim();
      const dataEnvioRaw = linha[colDataEnvio];
      const qtdBens = parseInt(linha[colQtdBens]) || 0;

      if (statusMacro === "Com Tombador") resumoGeral.totalComTombador++;
      if (statusMacro === "Concluido") resumoGeral.totalConcluido++;
      resumoGeral.totalBensAlocados += qtdBens;

      if (!painelTombadores[tombadorNome]) {
        painelTombadores[tombadorNome] = {
          nome: tombadorNome,
          processosAtivos: 0,
          processosConcluidos: 0,
          totalBensCarga: 0,
          bensVistoriadosPonta: 0,
          detalhesCarga: []
        };
      }

      const t = painelTombadores[tombadorNome];
      if (statusMacro === "Com Tombador") t.processosAtivos++;
      if (statusMacro === "Concluido") t.processosConcluidos++;
      t.totalBensCarga += qtdBens;

      let diasEmPosse = 0;
      if (dataEnvioRaw instanceof Date) {
        const diffTime = Math.abs(new Date() - dataEnvioRaw);
        diasEmPosse = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      }

      const progressoItens = progressoItensPorSequencial[sequencial] || { total: qtdBens, realizados: 0, descricoes: [] };
      
      // 🌟 Atribui os itens vistoriados reais do sequencial ao total do Tombador
      t.bensVistoriadosPonta += progressoItens.realizados;

      const descEntrada = String(linha[colDescricao] || "").trim();
      const descItens = progressoItens.descricoes ? progressoItens.descricoes.join(", ") : "";
      const descricaoFinal = descEntrada || descItens || "-";

      t.detalhesCarga.push({
        sequencial: sequencial,
        status: statusMacro,
        processo: String(linha[5] || "-").trim(),
        uorg: String(linha[7] || "-").trim(),
        descricao: descricaoFinal,
        dataEnvio: dataEnvioRaw instanceof Date ? Utilities.formatDate(dataEnvioRaw, Session.getScriptTimeZone(), "dd/MM/yyyy") : "-",
        diasEmPosse: diasEmPosse,
        qtdBensProcesso: qtdBens,
        itensVistoriados: progressoItens.realizados,
        totalItensGerados: progressoItens.total
      });
    }

    return {
      sucesso: true,
      resumoGeral: resumoGeral,
      resumoVisitasGlobal: resumoVisitasGlobal,
      tombadores: Object.values(painelTombadores)
    };

  } catch (error) {
    return { sucesso: false, message: error.toString() };
  }
}