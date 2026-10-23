/**
 * Busca os bens do sequencial na aba ENTRADA e gera um PDF formatado
 * em tamanho de bobina/rolo para Rotuladora (Uma etiqueta de 50x25mm por página).
 * Resolve intervalos de tombamento do tipo "16180660 a 16180662" gerando etiquetas individuais.
 */
function gerarEtiquetasServidor(sequencialTarget) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaEntrada = ss.getSheetByName("ENTRADA");
  const dados = abaEntrada.getDataRange().getValues();

  // Mapeamento das Colunas: Y (35) e AE (30) | Sequencial: A (0)
  const colSequencial = 0;
  const colDescricao = 35; // Coluna Y
  const colTombo = 30;     // Coluna AE

  let itensEtiqueta = [];

  // Coleta todos os bens correspondentes ao Sequencial clicado
  for (let i = 2; i < dados.length; i++) {
    const seq = String(dados[i][colSequencial]).trim();
    if (seq === String(sequencialTarget).trim()) {
      
      const tomboCelula = String(dados[i][colTombo]).trim();
      const descricaoCel = String(dados[i][colDescricao]).trim().toUpperCase();
      
      // Verifica se a célula contém um intervalo separado por " a " ou " A "
      if (/\s+a\s+/i.test(tomboCelula)) {
        // Divide o texto removendo espaços extras
        const partes = tomboCelula.split(/\s+a\s+/i);
        const inicio = parseInt(partes[0].replace(/\D/g, ""), 10);
        const fim = parseInt(partes[1].replace(/\D/g, ""), 10);
        
        // Se a conversão para número funcionar e o início for menor ou igual ao fim
        if (!isNaN(inicio) && !isNaN(fim) && inicio <= fim) {
          // Loop para gerar cada número do intervalo de forma individual
          for (let t = inicio; t <= fim; t++) {
            itensEtiqueta.push({
              tombo: String(t),
              descricao: descricaoCel
            });
          }
        } else {
          // Caso a conversão falhe por segurança, joga o texto cru da célula
          itensEtiqueta.push({ tombo: tomboCelula, descricao: descricaoCel });
        }
      } else {
        // Se for apenas um tombo único normal, insere diretamente
        itensEtiqueta.push({ tombo: tomboCelula, descricao: descricaoCel });
      }
      
    }
  }

  if (itensEtiqueta.length === 0) return null;

  // Imagem do Brasão da UFPE em Base64 para renderização garantida no PDF
  const brasaoBase64 = "https://lh3.googleusercontent.com/d/1r3GLe3Q36gFfTiCeiTqo-Vc7CQkBZTyD";

  // HTML com tamanho de página customizado para a dimensão exata de UMA etiqueta por folha (Bobina/Rolo)
  let htmlConteudo = `
  <html>
  <head>
    <style>
      @page { 
        size: 50mm 25mm; /* TAMANHO REAL DA ETIQUETA */
        margin: 1.5mm 2mm 1.5mm 2mm; 
      }
      body { 
        font-family: 'Arial', sans-serif; 
        margin: 0; 
        padding: 0; 
        background-color: #fff; 
        color: #000;
        width: 46mm;  
        height: 22mm; 
      }
      .etiqueta-box {
        width: 100%;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-sizing: border-box;
        page-break-after: always; /* Uma etiqueta por página na bobina */
      }
      .topo-container {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        height: 5.5mm;
      }
      .brasao-ufpe {
        height: 5.5mm;
        width: auto;
        display: block;
      }
      .texto-siads-ufpe {
        text-align: right;
        line-height: 0.9;
      }
      .txt-siads {
        font-size: 10px;
        font-weight: 900;
      }
      .txt-ufpe-sub {
        font-size: 8px;
        font-weight: bold;
      }
      .barcode-container {
        text-align: center;
        height: 8.5mm;
        margin: 0.5mm 0;
      }
      .barcode-img {
        height: 8.5mm;
        width: 100%;
        object-fit: fill;
      }
      .dados-patrimonio {
        line-height: 1.0;
      }
      .num-tombo {
        font-size: 13px;
        font-weight: bold;
        margin: 0;
      }
      .desc-item {
        font-size: 9.5px; 
        font-weight: normal;
        margin: 0;
        text-transform: uppercase;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
    </style>
  </head>
  <body>
  `;

  // Monta as páginas do PDF percorrendo os itens individuais gerados a partir da expansão
  itensEtiqueta.forEach((item, index) => {
    const tomboTratado = encodeURIComponent(item.tombo);
    const urlBarcode = `https://bwipjs-api.metafloor.com/?bcid=code128&text=${tomboTratado}&includetext=false&scale=2`;

    const estiloEstojo = (index === itensEtiqueta.length - 1) ? 'style="page-break-after: avoid;"' : '';

    htmlConteudo += `
      <div class="etiqueta-box" ${estiloEstojo}>
        <div class="topo-container">
          <img src="${brasaoBase64}" class="brasao-ufpe" />
          <div class="texto-siads-ufpe">
            <div class="txt-siads">SIADS</div>
            <div class="txt-ufpe-sub">UFPE</div>
          </div>
        </div>

        <div class="barcode-container">
          <img src="${urlBarcode}" class="barcode-img" />
        </div>

        <div class="dados-patrimonio">
          <div class="num-tombo">${item.tombo || ''}</div>
          <div class="desc-item">${item.descricao || ''}</div>
        </div>
      </div>
    `;
  });

  htmlConteudo += `
  </body>
  </html>
  `;

  const arquivoTemp = DriveApp.createFile("temp_etiqueta_rotuladora.html", htmlConteudo, MimeType.HTML);
  const blobPdf = arquivoTemp.getAs(MimeType.PDF).setName("Etiquetas_Rotuladora_" + sequencialTarget + ".pdf");
  
  const arquivoPdf = DriveApp.createFile(blobPdf);
  arquivoTemp.setTrashed(true);

  arquivoPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return arquivoPdf.getUrl();
}