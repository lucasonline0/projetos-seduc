function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function enviarAvaliacao(formData) {
  const doc = DocumentApp.create(`Avaliação - ${formData.nome}`);
  const body = doc.getBody();

  body.appendParagraph('Relatório de Avaliação de Estágio Probatório')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Servidor Avaliado: ${formData.nome}`);
  body.appendParagraph(`Matrícula: ${formData.matricula}`);
  body.appendParagraph(`Cargo: ${formData.cargo}`);
  body.appendParagraph(`Assiduidade: ${formData.assiduidade}`);
  body.appendParagraph(`Pontualidade: ${formData.pontualidade}`);
  body.appendParagraph(`Competência Técnica: ${formData.competencia}`);
  body.appendParagraph(`Observações: ${formData.observacoes || 'Nenhuma'}`);
  body.appendParagraph(`Data/Hora da Avaliação: ${new Date().toLocaleString()}`);

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdf = file.getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdf);
  pdfFile.setName(`Avaliação - ${formData.nome}.pdf`);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  GmailApp.sendEmail(formData.email, `Relatório de Avaliação de ${formData.nome}`, 
    'Segue em anexo o relatório da avaliação de estágio probatório.', {
      attachments: [pdfFile]
    });

  return pdfFile.getUrl();
}
