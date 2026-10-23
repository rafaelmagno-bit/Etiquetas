/**
 * MÓDULO: DASHBOARD SERVICE
 * Gerencia a inteligência analítica e agrupamento das UGs por siglas.
 * CORREÇÃO: Índice fixo para a coluna AG (índice 32) para evitar leitura errada da data (AI).
 */
function obterDadosDashboardEtiquetas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Carrega o mapa de UG para traduzir Código numérico para SIGLA
    const abaUg = ss.getSheetByName("UG");
    const mapaSiglasUg = {};
    if (abaUg) {
      const dadosUg = abaUg.getDataRange().getValues();
      for (let u = 1; u < dadosUg.length; u++) {
        const codUg = String(dadosUg[u][0]).trim().toUpperCase();
        const siglaUg = String(dadosUg[u][2] || "").trim().toUpperCase(); 
        if (codUg && siglaUg) {
          mapaSiglasUg[codUg] = siglaUg;
        }
      }
    }

    // 2. 🌟 NOVO: Carrega o mapa da UORG (Coluna C = Código, Coluna B = Descrição)
    const abaOrganograma = ss.getSheetByName("ORGANOGRAMA");
    const mapaUorgDescricao = {};
    if (abaOrganograma) {
      const dadosOrg = abaOrganograma.getDataRange().getValues();
      for (let o = 1; o < dadosOrg.length; o++) {
        const descUorg = String(dadosOrg[o][1] || "").trim();             // Coluna B (Índice 1)
        const codUorg = String(dadosOrg[o][2] || "").trim().toUpperCase(); // Coluna C (Índice 2)
        if (codUorg && descUorg) {
          mapaUorgDescricao[codUorg] = descUorg;
        }
      }
    }

    const abaEntrada = ss.getSheetByName("ENTRADA");
    if (!abaEntrada) {
      return { sucesso: false, message: "Aba 'ENTRADA' não encontrada nesta planilha de etiquetas." };
    }

    const dados = abaEntrada.getDataRange().getValues();
    if (dados.length <= 2) {
      return { sucesso: true, recentes: [], statusTotais: {}, porUg: {} };
    }

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const tresDiasAtras = new Date(hoje);
    tresDiasAtras.setDate(tresDiasAtras.getDate() - 3);

    const recentes = [];
    const statusTotais = {};
    const porUg = {};
    
    // Índice fixado em 32, que corresponde exatamente à Coluna AG (Status)
    const indiceStatus = 32; 

    for (let i = 2; i < dados.length; i++) {
      const r = dados[i];
      if (!r[0]) continue; 

      const seq = String(r[0]).trim();
      const dataRegRaw = r[1];                       
      const tipoEntrada = String(r[3] || "-").trim(); 
      const processo = String(r[5] || "-").trim();
      const ugCodigo = String(r[6] || "SEM UG").trim().toUpperCase(); 
      
      // 🌟 TRADUÇÃO DA UORG: Busca a descrição pelo código ou mantém o código original se não achar
      const uorgCodigo = String(r[7] || "-").trim();
      const uorgExibicao = mapaUorgDescricao[uorgCodigo.toUpperCase()] || uorgCodigo;

      const infoComplementar = String(r[35] || "-").trim(); 
      const qtd = parseInt(r[27]) || 0;              
      
      // Captura o status real da coluna AG
      const status = String(r[indiceStatus] || "Pendente").trim();
      const ugExibicao = mapaSiglasUg[ugCodigo] || ugCodigo; 

      if (dataRegRaw instanceof Date) {
        let dataRegLimpa = new Date(dataRegRaw);
        dataRegLimpa.setHours(0, 0, 0, 0);
        if (dataRegLimpa >= tresDiasAtras) {
          recentes.push({
            sequencial: seq,
            dataRegistro: Utilities.formatDate(dataRegRaw, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            tipo: tipoEntrada,
            processo: processo,
            ug: ugExibicao, 
            uorg: uorgExibicao, // 👈 Agora envia a descrição
            info: infoComplementar,
            quantidade: qtd
          });
        }
      }

      // Acumular total por status
      if (status) {
        statusTotais[status] = (statusTotais[status] || 0) + qtd;
      }

      if (ugExibicao) {
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
            uorg: uorgExibicao, // 👈 Agora envia a descrição no detalhamento da modal
            tipo: tipoEntrada,
            info: infoComplementar,
            quantidade: qtd,
            status: status
          });
        }
      }
    }

    Object.keys(porUg).forEach(ugKey => {
      porUg[ugKey].detalhes.sort((a, b) => {
        let compUorg = a.uorg.localeCompare(b.uorg);
        if (compUorg !== 0) return compUorg;
        if (a.status.toLowerCase() === "pendente" && b.status.toLowerCase() !== "pendente") return -1;
        if (a.status.toLowerCase() !== "pendente" && b.status.toLowerCase() === "pendente") return 1;
        return 0;
      });
    });

    return { sucesso: true, recentes: recentes, statusTotais: statusTotais, porUg: porUg };
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