/**
 * MÓDULO: TOMBADOR SERVICE
 * Retorna a lista de tombadores cadastrados na aba 'TOMBADORES' para preencher o select do HTML
 */
function obterListaTombadores() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("TOMBADORES");
    if (!aba) return [];
    
    const dados = aba.getDataRange().getValues();
    const lista = [];
    
    // Pula o cabeçalho (i=1) e pega as colunas A (Nome) e B (E-mail)
    for (let i = 1; i < dados.length; i++) {
      const nome = String(dados[i][0]).trim();
      const email = String(dados[i][1]).trim();
      if (nome) {
        lista.push({ nome: nome, email: email });
      }
    }
    return lista;
  } catch (error) {
    return [];
  }
}

/**
 * Registra o tombador (AH), a data (AI), altera o status (AG), envia e-mail de notificação
 * e grava a rastreabilidade do operador que destinou (AK) e data/hora (AL) para o POLARE.
 */
function destinarAotombador(sequencial, tombadorNome) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Captura o usuário logado que está executando a destinação
    const emailLogado = typeof obterUsuarioLogado === "function" ? obterUsuarioLogado() : Session.getActiveUser().getEmail();
    const dataHoraAtual = new Date();

    // 1. Identificar o e-mail do tombador selecionado na aba TOMBADORES
    const abaTombadores = ss.getSheetByName("TOMBADORES");
    let emailTombador = "";
    if (abaTombadores) {
      const dadosT = abaTombadores.getDataRange().getValues();
      for (let t = 1; t < dadosT.length; t++) {
        if (String(dadosT[t][0]).trim() === String(tombadorNome).trim()) {
          emailTombador = String(dadosT[t][1]).trim();
          break;
        }
      }
    }

    const abaEntrada = ss.getSheetByName("ENTRADA");
    if (!abaEntrada) return { sucesso: false, message: "Aba 'ENTRADA' não encontrada." };

    const dados = abaEntrada.getDataRange().getValues();
    
    // 🌟 DEFINIÇÃO FIXA DAS COLUNAS:
    const colAG = 33; // Coluna AG (Status do fluxo) - Fixo
    const colAH = 34; // Coluna AH (Agente Tombador) - Fixo
    const colAI = 35; // Coluna AI (Data de Confirmação) - Fixo
    const colAK = 37; // 🌟 Coluna AK (E-mail do Operador que Destinou) - Novo para POLARE
    const colAL = 38; // 🌟 Coluna AL (Data e Hora da Destinação) - Novo para POLARE
    
    let linhasAtualizadas = 0;
    let dadosResumo = null;

    // 2. Gravação de dados na Planilha
    for (let i = 2; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(sequencial).trim()) {
        
        // Coleta dados do primeiro item para compor o e-mail resumo
        if (!dadosResumo) {
          dadosResumo = {
            sequencial: sequencial,
            tipo: String(dados[i][3] || "-").trim(),
            processo: String(dados[i][5] || "-").trim(),
            qtdTotal: 0
          };
        }
        dadosResumo.qtdTotal += parseInt(dados[i][27]) || 0; // Acumula quantidade de bens

        // Atualizações exatas das colunas por número físico
        abaEntrada.getRange(i + 1, colAG).setValue("Com Tombador"); // Grava na AG
        abaEntrada.getRange(i + 1, colAH).setValue(tombadorNome);   // Grava na AH
        abaEntrada.getRange(i + 1, colAI).setValue(dataHoraAtual); // Grava na AI
        
        // 🌟 REGISTRO DE RASTREABILIDADE PARA O RELATÓRIO POLARE
        abaEntrada.getRange(i + 1, colAK).setValue(emailLogado);   // Grava e-mail na AK
        abaEntrada.getRange(i + 1, colAL).setValue(dataHoraAtual); // Grava data/hora na AL
        
        linhasAtualizadas++;
      }
    }

    if (linhasAtualizadas === 0) {
      return { sucesso: false, message: "Sequencial não localizado para alteração." };
    }

    // 3. Disparo Automatizado do E-mail (Se o e-mail do tombador for válido)
    let emailStatusInfo = "E-mail não enviado (endereço não localizado).";
    
    if (emailTombador && emailTombador.includes("@")) {
      const assunto = `[Patrimônio] Novo Tombamento Vinculado — Sequencial ${dadosResumo.sequencial}`;
      
      const corpoHtml = `
        <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; color: #333; line-height: 1.6;">
          <h3 style="color: #0d6efd; margin-bottom: 5px;">Módulo de Fixação de Etiquetas — CBMAC</h3>
          <p>Olá, <strong>${tombadorNome}</strong>,</p>
          <p>Informamos que um novo processo de tombamento foi associado ao seu perfil no sistema.</p>
          
          <div style="background-color: #f8f9fa; border-left: 4px solid #0d6efd; padding: 15px; margin: 20px 0; border-radius: 4px;">
            <h4 style="margin-top: 0; color: #495057;">Resumo do Processo</h4>
            <table style="font-size: 14px; width: 100%; border-collapse: collapse;">
              <tr><td style="padding: 4px 0; color: #6c757d; width: 120px;"><strong>Sequencial:</strong></td><td><strong>${dadosResumo.sequencial}</strong></td></tr>
              <tr><td style="padding: 4px 0; color: #6c757d;"><strong>Tipo de Entrada:</strong></td><td>${dadosResumo.tipo}</td></tr>
              <tr><td style="padding: 4px 0; color: #6c757d;"><strong>Nº Processo:</strong></td><td>${dadosResumo.processo}</td></tr>
              <tr><td style="padding: 4px 0; color: #6c757d;"><strong>Qtd. de Bens:</strong></td><td style="color: #0d6efd; font-weight: bold;">${dadosResumo.qtdTotal} itens</td></tr>
            </table>
          </div>
          
          <p style="background-color: #fff3cd; color: #664d03; padding: 12px; border-radius: 4px; border: 1px solid #ffe69c;">
            ⚠️ <strong>Importante:</strong> O documento físico juntamente com as respectivas etiquetas já encontram-se disponíveis para retirada na sala da <strong>Divisão de Bens Móveis</strong>.
          </p>
          <hr style="border: 0; border-top: 1px solid #eee; margin: 25px 0;">
          <small style="color: #999; display: block; text-align: center;">Sistema de Gestão Patrimonial — UFPE / CBMAC</small>
        </div>
      `;

      GmailApp.sendEmail(emailTombador, assunto, "", { htmlBody: corpoHtml });
      emailStatusInfo = `Notificação enviada para: ${emailTombador}`;
    }

    return { 
      sucesso: true, 
      message: `Status atualizado com sucesso (${linhasAtualizadas} itens). \n[OK] ${emailStatusInfo}` 
    };

  } catch (error) {
    return { sucesso: false, message: "Erro no fluxo do tombador: " + error.toString() };
  }
}