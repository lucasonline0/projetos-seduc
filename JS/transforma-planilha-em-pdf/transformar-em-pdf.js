function converterPlanilhasParaPDF() {
  const pastaOrigemId = '';
  const pastaOrigem = DriveApp.getFolderById(pastaOrigemId);

  let pastaPDF;
  const pastas = pastaOrigem.getFoldersByName("pdf");
  pastaPDF = pastas.hasNext() ? pastas.next() : pastaOrigem.createFolder("pdf");

  const arquivos = pastaOrigem.getFilesByType(MimeType.GOOGLE_SHEETS);

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nomeArquivo = arquivo.getName();
    const planilha = SpreadsheetApp.openById(arquivo.getId());

    planilha.getSheets().forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const totalRows = sheet.getMaxRows();
      const totalCols = sheet.getMaxColumns();

      if (lastRow < totalRows) {
        sheet.deleteRows(lastRow + 1, totalRows - lastRow);
      }
      if (lastCol < totalCols) {
        sheet.deleteColumns(lastCol + 1, totalCols - lastCol);
      }
    });

    const url = 'https://docs.google.com/spreadsheets/d/' + arquivo.getId() + '/export?';

    const parametros = {
      exportFormat: 'pdf',
      format: 'pdf',
      size: 'A4',
      portrait: true,
      fitw: true,
      top_margin: 0.5,
      bottom_margin: 0.5,
      left_margin: 0.5,
      right_margin: 0.5,
      sheetnames: false,
      printtitle: false,
      pagenumbers: false,
      gridlines: false,
      fzr: false,
    };

    const token = ScriptApp.getOAuthToken();
    const options = {
      headers: {
        Authorization: 'Bearer ' + token,
      },
      muteHttpExceptions: true,
    };

    const queryString = Object.keys(parametros)
      .map(key => key + '=' + encodeURIComponent(parametros[key]))
      .join('&');

    const response = UrlFetchApp.fetch(url + queryString, options);
    const blob = response.getBlob().setName(nomeArquivo + ".pdf");

    pastaPDF.createFile(blob);
  }

  Logger.log("Conversão finalizada!");
}

//script by @lucasonline0