/**
 * MÓDULO: DASHBOARD SERVICE
 * Gerencia a inteligência analítica e agrupamento das UGs por siglas.
 * CORREÇÃO: Índice fixo para a coluna AG (índice 32) para evitar leitura errada da data (AI).
 */
function obterDadosDashboardEtiquetas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Carrega o mapa de UG
    const abaUg = ss.getSheetByName("UG");
    const mapaSiglasUg = {};
    if (abaUg) {
      const dadosUg = abaUg.getDataRange().getValues();
      for (let u = 1; u < dadosUg.length; u++) {
        const codUg = String(dadosUg[u][0]).trim().toUpperCase();
        const siglaUg = String(dadosUg[u][2] || "").trim().toUpperCase();
        if (codUg && siglaUg) mapaSiglasUg[codUg] = siglaUg;
      }
    }

    // 2. Carrega o mapa da UORG
    const abaOrganograma = ss.getSheetByName("ORGANOGRAMA");
    const mapaUorgDescricao = {};
    if (abaOrganograma) {
      const dadosOrg = abaOrganograma.getDataRange().getValues();
      for (let o = 1; o < dadosOrg.length; o++) {
        const codUorg = String(dadosOrg[o][2] || "").trim().toUpperCase();
        const descUorg = String(dadosOrg[o][1] || "").trim();
        if (codUorg && descUorg) mapaUorgDescricao[codUorg] = descUorg;
      }
    }

    const abaEntrada = ss.getSheetByName("ENTRADA");
    if (!abaEntrada) return { sucesso: false, message: "Aba ENTRADA não encontrada." };

    const dadosEntrada = abaEntrada.getDataRange().getValues();
    const porUg = {};
    const porCorresponsavel = {};
    const recentes = [];
    const statusTotais = {};

    const indiceStatus = 32; // Coluna AG

    // 🌟 INÍCIO EM i = 2 PARA IGNORAR O CABEÇALHO DA PLANILHA (LINHAS 1 E 2)
    for (let i = 2; i < dadosEntrada.length; i++) {
      const r = dadosEntrada[i];
      const seq = String(r[0] || "").trim();

      // Pula linhas vazias ou palavras de cabeçalho residuais
      if (!seq || seq.toUpperCase() === "SEQUENCIAL" || seq.toUpperCase() === "SEQ") continue;

      const ugCodigo = String(r[6] || "").trim().toUpperCase();
      const uorgCodigo = String(r[7] || "").trim().toUpperCase();
      const corresponsavelNome = String(r[15] || "SEM CORRESPONSÁVEL").trim().toUpperCase() || "SEM CORRESPONSÁVEL";

      // Filtra caso o texto da célula seja o próprio título da coluna
      if (corresponsavelNome === "CORRESPONSÁVEL" || corresponsavelNome === "CORRESPONSAVEL") continue;

      const processo = String(r[4] || "-").trim();
      const tipoEntrada = String(r[3] || "-").trim();
      const uorgExibicao = mapaUorgDescricao[uorgCodigo] || uorgCodigo;
      const infoComplementar = String(r[35] || "-").trim();
      const qtd = parseInt(r[27]) || 0;
      const status = String(r[indiceStatus] || "Pendente").trim();
      const ugExibicao = mapaSiglasUg[ugCodigo] || ugCodigo;

      // Acumular total por status
      if (status) {
        statusTotais[status] = (statusTotais[status] || 0) + qtd;
      }

      // Agrupamento por UG
      if (ugExibicao && ugExibicao !== "UG") {
        if (!porUg[ugExibicao]) {
          porUg[ugExibicao] = { ug: ugExibicao, pendentes: 0, comTombador: 0, detalhes: [] };
        }
        if (status.toLowerCase() === "pendente") {
          porUg[ugExibicao].pendentes += qtd;
        } else if (status.toLowerCase() === "com tombador") {
          porUg[ugExibicao].comTombador += qtd;
        }
        if (status.toLowerCase() === "pendente" || status.toLowerCase() === "com tombador") {
          porUg[ugExibicao].detalhes.push({
            sequencial: seq,
            processo: processo,
            uorg: uorgExibicao,
            tipo: tipoEntrada,
            info: infoComplementar,
            quantidade: qtd,
            status: status
          });
        }
      }

      // Agrupamento por CORRESPONSÁVEL
      if (!porCorresponsavel[corresponsavelNome]) {
        porCorresponsavel[corresponsavelNome] = { nome: corresponsavelNome, pendentes: 0, comTombador: 0, detalhes: [] };
      }
      if (status.toLowerCase() === "pendente") {
        porCorresponsavel[corresponsavelNome].pendentes += qtd;
      } else if (status.toLowerCase() === "com tombador") {
        porCorresponsavel[corresponsavelNome].comTombador += qtd;
      }
      if (status.toLowerCase() === "pendente" || status.toLowerCase() === "com tombador") {
        porCorresponsavel[corresponsavelNome].detalhes.push({
          sequencial: seq,
          processo: processo,
          ug: ugExibicao,
          uorg: uorgExibicao,
          tipo: tipoEntrada,
          info: infoComplementar,
          quantidade: qtd,
          status: status
        });
      }
    }

    return { 
      sucesso: true, 
      recentes: recentes, 
      statusTotais: statusTotais, 
      porUg: porUg, 
      porCorresponsavel: porCorresponsavel 
    };

  } catch (error) {
    return { sucesso: false, message: "Erro no servidor local: " + error.toString() };
  }
}

/**
 * Busca a aba ORGANOGRAMA e cria um dicionário/mapa { CÓDIGO: DESCRIÇÃO }
 */
function obterMapaUorgs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaOrganograma = ss.getSheetByName("ORGANOGRAMA");
  var mapa = {};
  
  if (!abaOrganograma) return mapa;

  var dados = abaOrganograma.getDataRange().getValues();
  
  // Começa do índice 1 para pular a linha do cabeçalho
  for (var i = 1; i < dados.length; i++) {
    var descricao = dados[i][1];             // Coluna B (índice 1)
    var codigo = String(dados[i][2]).trim(); // Coluna C (índice 2)
    
    if (codigo) {
      mapa[codigo] = descricao;
    }
  }
  
  return mapa;
}