/**
 * MÓDULO: DASHBOARD SERVICE
 * Gerencia a inteligência analítica e agrupamento das UGs por siglas.
 * CORREÇÃO: Índice fixo para a coluna AG (índice 32) para evitar leitura errada da data (AI).
 */
function obterDadosDashboardEtiquetas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Carrega o mapa de UG (Código -> Sigla)
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

    // 2. Carrega o mapa do ORGANOGRAMA (Código -> Descrição UORG)
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

    // Filtro de data: últimos 3 dias
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const tresDiasAtras = new Date(hoje);
    tresDiasAtras.setDate(tresDiasAtras.getDate() - 3);

    const indiceStatus = 32; // Coluna AG

    // Ignora linhas de cabeçalho (linhas 1 e 2)
    for (let i = 2; i < dadosEntrada.length; i++) {
      const r = dadosEntrada[i];
      const seq = String(r[0] || "").trim();

      if (!seq || seq.toUpperCase() === "SEQUENCIAL" || seq.toUpperCase() === "SEQ") continue;

      const dataRegRaw = r[1]; // Coluna B
      const tipoEntrada = String(r[3] || "-").trim();
      
      // 🌟 CORREÇÃO: Coluna F (índice 5) é o Número do Processo
      const processo = String(r[5] || "-").trim(); 
      
      const ugCodigo = String(r[6] || "").trim().toUpperCase();
      const uorgCodigo = String(r[7] || "").trim().toUpperCase();
      const corresponsavelNome = String(r[15] || "SEM CORRESPONSÁVEL").trim().toUpperCase() || "SEM CORRESPONSÁVEL";
      const qtd = parseInt(r[27]) || 0;
      const status = String(r[indiceStatus] || "Pendente").trim();
      const infoComplementar = String(r[35] || "-").trim();

      if (corresponsavelNome === "CORRESPONSÁVEL" || corresponsavelNome === "CORRESPONSAVEL") continue;

      const ugExibicao = mapaSiglasUg[ugCodigo] || ugCodigo;
      const uorgExibicao = mapaUorgDescricao[uorgCodigo] || uorgCodigo;

      // 1. Processar lançamentos recentes (últimos 3 dias)
      if (dataRegRaw) {
        let dataRegObj = (dataRegRaw instanceof Date) ? dataRegRaw : new Date(dataRegRaw);
        if (!isNaN(dataRegObj.getTime())) {
          let dataRegLimpa = new Date(dataRegObj);
          dataRegLimpa.setHours(0, 0, 0, 0);

          if (dataRegLimpa >= tresDiasAtras) {
            recentes.push({
              sequencial: seq,
              dataRegistro: Utilities.formatDate(dataRegObj, Session.getScriptTimeZone(), "dd/MM/yyyy"),
              tipo: tipoEntrada,
              processo: processo,
              ug: ugExibicao,
              uorg: uorgExibicao,
              info: infoComplementar,
              quantidade: qtd
            });
          }
        }
      }

      // 2. Acumular totais de bens por status
      if (status) {
        statusTotais[status] = (statusTotais[status] || 0) + qtd;
      }

      // 3. Agrupamento por UG
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

      // 4. Agrupamento por CORRESPONSÁVEL
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