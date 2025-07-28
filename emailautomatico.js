function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📩 E-mails')
    .addItem('Enviar E-mails', 'enviarPlanilhasPorEmail')
    .addToUi();
}

function enviarPlanilhasPorEmail() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("BaseDREs");
  const dados = aba.getDataRange().getValues();
  const pastaId = ""; // ID da pasta com as planilhas
  const pasta = DriveApp.getFolderById(pastaId);

  for (let i = 1; i < dados.length; i++) {
    const nomeDRE = dados[i][0];
    const emailDestino = dados[i][1];

    if (!nomeDRE || !emailDestino) continue;

    const arquivos = pasta.getFilesByName(nomeDRE);
    if (!arquivos.hasNext()) continue;

    const arquivo = arquivos.next();
    if (arquivo.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const arquivoId = arquivo.getId();
    const nomeArquivo = arquivo.getName();

    const token = ScriptApp.getOAuthToken();
    const options = {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`
      },
      muteHttpExceptions: true
    };

    const urlXlsx = `https://docs.google.com/spreadsheets/d/${arquivoId}/export?format=xlsx`;
    const respostaXlsx = UrlFetchApp.fetch(urlXlsx, options);
    if (respostaXlsx.getResponseCode() !== 200) {
      Logger.log(`Erro ao exportar XLSX: ${nomeDRE}`);
      continue;
    }
    const blobXlsx = respostaXlsx.getBlob().setName(nomeDRE + ".xlsx");

    const urlPdf = `https://docs.google.com/spreadsheets/d/${arquivoId}/export?format=pdf&portrait=true&exportFormat=pdf`;
    const respostaPdf = UrlFetchApp.fetch(urlPdf, options);
    if (respostaPdf.getResponseCode() !== 200) {
      Logger.log(`Erro ao exportar PDF: ${nomeDRE}`);
      continue;
    }
    const blobPdf = respostaPdf.getBlob().setName(nomeDRE + ".pdf");

    const corpoEmail = `
    `;

    GmailApp.sendEmail(emailDestino, `Relatório - ${nomeDRE}`, "", {
      htmlBody: corpoEmail,
      attachments: [blobXlsx, blobPdf],
      name: "",
      cc: ''
    });

    const dataehora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    aba.getRange(i + 1, 5).setValue(dataehora); // Coluna E
  }
}

// script by @lucasonline0
