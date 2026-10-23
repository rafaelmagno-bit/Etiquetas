/**
 * MÓDULO: RELATÓRIO SERVICE
 * Concentra os cruzamentos de dados complexos com as bases de SERVIDORES, UG e ORGANOGRAMA.
 */
function getDadosRelatorioEntrada(sequencial) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Mapeamento da aba SERVIDORES
    const abaServidores = ss.getSheetByName("SERVIDORES");
    const mapaServidores = {};
    if (abaServidores) {
      const dadosServ = abaServidores.getDataRange().getValues();
      for (let s = 1; s < dadosServ.length; s++) {
        const chave = String(dadosServ[s][0]).trim();
        if (chave) {
          mapaServidores[chave] = {
            nome: String(dadosServ[s][1] || "").trim(),     
            email: String(dadosServ[s][7] || "-").trim(),    
            telefone: String(dadosServ[s][8] || "-").trim()  
          };
        }
      }
    }

    // Mapeamento da aba UG
    const abaUg = ss.getSheetByName("UG");
    const mapaUg = {};
    if (abaUg) {
      const dadosUg = abaUg.getDataRange().getValues();
      for (let u = 1; u < dadosUg.length; u++) {
        const codUg = String(dadosUg[u][0]).trim().toUpperCase();
        if (codUg) mapaUg[codUg] = String(dadosUg[u][1] || "").trim();
      }
    }

    // Mapeamento da aba ORGANOGRAMA
    const abaOrganograma = ss.getSheetByName("ORGANOGRAMA");
    const mapaUorg = {};
    if (abaOrganograma) {
      const dadosOrg = abaOrganograma.getDataRange().getValues();
      for (let o = 1; o < dadosOrg.length; o++) {
        const codUorg = String(dadosOrg[o][2]).trim().toUpperCase();
        if (codUorg) {
          mapaUorg[codUorg] = String(dadosOrg[o][1] || "").trim();
        }
      }
    }

    const abaEntrada = ss.getSheetByName("ENTRADA");
    if (!abaEntrada) return { sucesso: false, message: "Aba ENTRADA não encontrada." };

    const dados = abaEntrada.getDataRange().getValues();
    let dadosProcesso = null;
    const itens = [];

    for (let i = 2; i < dados.length; i++) {
      const r = dados[i];
      if (String(r[0]).trim() === String(sequencial).trim()) {
        
        if (!dadosProcesso) {
          let dataFormatada = "-";
          if (r[1] instanceof Date) {
            dataFormatada = Utilities.formatDate(r[1], Session.getScriptTimeZone(), "dd/MM/yyyy");
          }

          const ugCodigo = String(r[6] || "-").trim().toUpperCase();
          const uorgCodigo = String(r[7] || "-").trim().toUpperCase();
          const siapeCorresponsavel = String(r[17] || "").trim(); 
          const cpfCorresponsavel = String(r[16] || "").trim();
          
          let contatoEmail = "-";
          let contatoTelefone = "-";
          let servidorCadastrado = mapaServidores[siapeCorresponsavel] || mapaServidores[cpfCorresponsavel];
          if (servidorCadastrado) {
            contatoEmail = servidorCadastrado.email;
            contatoTelefone = servidorCadastrado.telefone;
          }

          dadosProcesso = {
            sequencial: String(r[0]).trim(),
            data: dataFormatada,
            tipo: String(r[3] || "-").trim(),
            processo: String(r[5] || "-").trim(),
            ug: ugCodigo,
            ugDescricao: mapaUg[ugCodigo] || "",
            uorg: uorgCodigo,
            uorgDescricao: mapaUorg[uorgCodigo] || "",
            corresponsel: String(r[15] || "").trim(), 
            cpfCorresponsavel: cpfCorresponsavel,
            siapeCorresponsavel: siapeCorresponsavel,
            emailCorresponsavel: contatoEmail, 
            telefoneCorresponsavel: contatoTelefone
          };
        }

        const valorTombamento = String(r[30] || "").trim();
        itens.push({
          ordem: r[22] || "-",       
          cadmat: r[23] || "-",      
          descricao: r[24] || "-",   
          qtd: parseInt(r[27]) || 0, 
          valorUnit: parseFloat(r[28]) || 0, 
          tombamento: valorTombamento !== "" ? valorTombamento : "Pendente"
        });
      }
    }

    if (dadosProcesso) {
      dadosProcesso.itens = itens;
      return { sucesso: true, dados: dadosProcesso };
    } else {
      return { sucesso: false };
    }
  } catch (error) {
    return { sucesso: false, error: error.toString() };
  }
}

/**
 * Gera o relatório individual de produtividade para comprovação no sistema POLARE / UFPE.
 * @param {string} dataInicioStr - Data inicial (YYYY-MM-DD)
 * @param {string} dataFimStr - Data final (YYYY-MM-DD)
 * @return {Object} Dados consolidados das atividades realizadas no período
 */
function obterRelatorioProdutividadePolare(dataInicioStr, dataFimStr) {
  try {
    var emailLogado = typeof obterUsuarioLogado === "function" ? obterUsuarioLogado() : Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Identificar o nome do Servidor na aba SERVIDORES
    var abaServidores = ss.getSheetByName("SERVIDORES");
    var dadosServidores = abaServidores ? abaServidores.getDataRange().getValues() : [];
    var nomeServidor = emailLogado;
    
    for (var i = 1; i < dadosServidores.length; i++) {
      var emailDaAba = String(dadosServidores[i][2] || "").trim().toLowerCase();
      if (emailDaAba === String(emailLogado).trim().toLowerCase()) {
        nomeServidor = String(dadosServidores[i][1] || "").trim();
        break;
      }
    }

    var dataInicio = new Date(dataInicioStr + "T00:00:00");
    var dataFim = new Date(dataFimStr + "T23:59:59");

    var historicoAcoes = [];
    var totalFixada = 0;
    var totalEntregue = 0;
    var totalNaoLocalizado = 0;
    var totalNaoComporta = 0;
    var totalDestinacoes = 0;
    var totalProcessosMovimentados = new Set();

    // 2. Processar Vistorias de Campo na aba ITENS_DEMANDA
    var abaItens = ss.getSheetByName("ITENS_DEMANDA");
    if (abaItens) {
      var dadosItens = abaItens.getDataRange().getValues();
      if (dadosItens.length > 1) {
        var cabecalhosItens = dadosItens[0].map(c => String(c).trim().toUpperCase());

        var idxSequencial = 1; // Coluna B
        var idxDescricao = 3;  // Coluna D
        var idxTombo = buscarIndiceColuna(cabecalhosItens, ["TOMBO", "PATRIMONIO", "BEM"], 4);
        var idxAgente = buscarIndiceColuna(cabecalhosItens, ["AGENTE", "EMAIL", "USUARIO", "RESPONSAVEL", "TOMBADOR"], 7);
        var idxStatus = buscarIndiceColuna(cabecalhosItens, ["STATUS", "SITUAÇÃO", "VISTORIA", "RESULTADO"], 6);
        var idxData = buscarIndiceColuna(cabecalhosItens, ["DATA", "ATUALIZAÇÃO", "HORA"], 10);

        for (var j = 1; j < dadosItens.length; j++) {
          var linha = dadosItens[j];
          var agenteReg = String(linha[idxAgente] || "").trim().toLowerCase();
          var emailAlvo = String(emailLogado).trim().toLowerCase();
          var nomeAlvo = String(nomeServidor).trim().toLowerCase();

          if (agenteReg.includes(emailAlvo) || (nomeAlvo && agenteReg.includes(nomeAlvo))) {
            var dataVistoria = parseDataSegura(linha[idxData]);

            if (!dataVistoria || (dataVistoria >= dataInicio && dataVistoria <= dataFim)) {
              var statusVistoria = String(linha[idxStatus] || "Vistoriado").trim();
              var sequencialValor = String(linha[idxSequencial] || "-").trim();
              
              if (sequencialValor !== "-") totalProcessosMovimentados.add(sequencialValor);

              var stLower = statusVistoria.toLowerCase();
              if (stLower.includes("fixada")) totalFixada++;
              else if (stLower.includes("entregue")) totalEntregue++;
              else if (stLower.includes("não localizado") || stLower.includes("nao localizado")) totalNaoLocalizado++;
              else if (stLower.includes("não comporta") || stLower.includes("nao comporta")) totalNaoComporta++;

              historicoAcoes.push({
                data: dataVistoria ? Utilities.formatDate(dataVistoria, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "Data N/I",
                tipoAcao: "Vistoria de Campo / Etiquetagem",
                processo:  sequencialValor,
                tombo: String(linha[idxTombo] || "-"),
                descricao: String(linha[idxDescricao] || "Item sem descrição"),
                statusResultante: statusVistoria
              });
            }
          }
        }
      }
    }

    // 3. Processar Destinações e Atribuições na aba ENTRADA
    var abaEntrada = ss.getSheetByName("ENTRADA");
    if (abaEntrada) {
      var dadosEntrada = abaEntrada.getDataRange().getValues();
      if (dadosEntrada.length > 1) {
        var idxDescricaoAJ = 35; // Coluna AJ
        var idxUserDestinacaoAK = 36; // Coluna AK (Quem fez a destinação)
        var idxDataDestinacaoAL = 37; // Coluna AL (Data/Hora da destinação)
        var idxQtdBensAB = 27; // Coluna AB (Quantidade real de bens)

        for (var k = 1; k < dadosEntrada.length; k++) {
          var lEntrada = dadosEntrada[k];
          var emailOperadorDestinacao = String(lEntrada[idxUserDestinacaoAK] || "").trim().toLowerCase();
          var tombadorAtribuido = String(lEntrada[8] || "").trim().toLowerCase();
          
          var emailAlvo = String(emailLogado).trim().toLowerCase();
          var nomeAlvo = String(nomeServidor).trim().toLowerCase();

          // Formatação correta do campo Tombo/Qtd para processo inteiro
          var qtdBens = parseInt(lEntrada[idxQtdBensAB]) || 0;
          var tomboTexto = qtdBens > 0 ? (qtdBens + " item(ns)") : "-";

          // AÇÃO A: O usuário logado foi quem realizou a DESTINAÇÃO (Coluna AK e AL)
          if (emailOperadorDestinacao.includes(emailAlvo)) {
            var dataDest = parseDataSegura(lEntrada[idxDataDestinacaoAL]);
            if (!dataDest || (dataDest >= dataInicio && dataDest <= dataFim)) {
              var seqProc = String(lEntrada[0] || lEntrada[2] || "-");
              if (seqProc !== "-") totalProcessosMovimentados.add(seqProc);
              totalDestinacoes++;

              historicoAcoes.push({
                data: dataDest ? Utilities.formatDate(dataDest, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "Data N/I",
                tipoAcao: "Destinação de Processo e Impressões", // Ajustado
                processo: seqProc,
                tombo: tomboTexto, // Correção: Exibe "-" ou a Qtd real
                descricao: "Processos destinados à tombamento com relatórios e etiquetas impressas",
                statusResultante: String(lEntrada[10] || "Destinado")
              });
            }
          }

          // AÇÃO B: O usuário logado foi o RECEPTOR/TOMBADOR do processo
          else if (tombadorAtribuido.includes(emailAlvo) || (nomeAlvo && tombadorAtribuido.includes(nomeAlvo))) {
            var dataEnvio = parseDataSegura(lEntrada[9]);
            if (!dataEnvio || (dataEnvio >= dataInicio && dataEnvio <= dataFim)) {
              var seqEntrada = String(lEntrada[0] || lEntrada[2] || "-");
              if (seqEntrada !== "-") totalProcessosMovimentados.add(seqEntrada);

              var descTombamento = String(lEntrada[idxDescricaoAJ] || "Atribuição de Carga de Trabalho para Vistoria").trim();

              historicoAcoes.push({
                data: dataEnvio ? Utilities.formatDate(dataEnvio, Session.getScriptTimeZone(), "dd/MM/yyyy") : "Data N/I",
                tipoAcao: "Recebimento / Posse de Processo",
                processo: seqEntrada,
                tombo: tomboTexto,
                descricao: descTombamento,
                statusResultante: String(lEntrada[10] || "Com Tombador")
              });
            }
          }
        }
      }
    }

    // Ordenar histórico do mais recente para o mais antigo
    historicoAcoes.sort((a, b) => new Date(b.data) - new Date(a.data));

    return {
      sucesso: true,
      servidor: nomeServidor,
      email: emailLogado,
      periodo: {
        inicio: Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        fim: Utilities.formatDate(dataFim, Session.getScriptTimeZone(), "dd/MM/yyyy")
      },
      resumo: {
        totalVistorias: totalFixada + totalEntregue + totalNaoLocalizado + totalNaoComporta,
        totalDestinacoes: totalDestinacoes,
        etiquetasFixadas: totalFixada,
        entreguesServidor: totalEntregue,
        naoLocalizados: totalNaoLocalizado,
        naoComporta: totalNaoComporta,
        processosAtendidos: totalProcessosMovimentados.size
      },
      detalhes: historicoAcoes
    };

  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao gerar relatório POLARE: " + e.toString() };
  }
}

/**
 * Função auxiliar para localizar colunas dinamicamente pelos nomes do cabeçalho
 */
function buscarIndiceColuna(cabecalhos, palavrasChave, indicePadrao) {
  for (var i = 0; i < cabecalhos.length; i++) {
    for (var k = 0; k < palavrasChave.length; k++) {
      if (cabecalhos[i].indexOf(palavrasChave[k]) !== -1) {
        return i;
      }
    }
  }
  return indicePadrao;
}

/**
 * Função auxiliar para tratar e converter datas em múltiplos formatos com segurança
 */
function parseDataSegura(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;
  var str = String(valor).trim();
  if (str === "") return null;
  
  if (str.indexOf("/") !== -1) {
    var partes = str.split(" ");
    var dPartes = partes[0].split("/");
    if (dPartes.length === 3) {
      var iso = dPartes[2] + "-" + dPartes[1] + "-" + dPartes[0];
      if (partes[1]) iso += "T" + partes[1];
      var d = new Date(iso);
      if (!isNaN(d.getTime())) return d;
    }
  }
  var d2 = new Date(str);
  return isNaN(d2.getTime()) ? null : d2;
}