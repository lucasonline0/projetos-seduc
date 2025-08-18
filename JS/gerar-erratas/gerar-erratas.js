function gerarErratas() {
  const ui = SpreadsheetApp.getUi();
  const abaNome = "BASE_ER"; 
  const folha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(abaNome);
  if (!folha) {
    ui.alert("Aba BASE_ER não encontrada!");
    return;
  }

  const dados = folha.getRange(2, 1, folha.getLastRow() - 1, folha.getLastColumn()).getValues();
  const valores = dados.filter(linha => linha[6] && String(linha[6]).trim() !== "");
  if (valores.length === 0) {
    ui.alert("Nenhuma linha válida encontrada na aba BASE_ER.");
    return;
  }

  const pastaId = ""; //id da pasta
  let doc;
  if (pastaId) {
    const pasta = DriveApp.getFolderById(pastaId);
    doc = DocumentApp.create("ERRATAS");
    pasta.addFile(DriveApp.getFileById(doc.getId()));
  } else {
    doc = DocumentApp.create("ERRATAS");
  }

  const body = doc.getBody();
  body.appendParagraph("ERRATAS - Documentos de Férias")
      .setFontSize(12)
      .setBold(true)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph("");

  const COL_F = 5 - 1;  
  const COL_G = 6 - 1;  
  const COL_M = 12 - 1; 
  const COL_N = 13 - 1; 
  const COL_P = 16 - 1; 

  valores.forEach((linha, index) => {
    Logger.log(`Processando errata ${index + 1} de ${valores.length}: ${linha[COL_G]}`);
    
    const nome = linha[COL_G] || "";
    const matricula = linha[COL_F] || "";
    const periodoOriginal = linha[COL_M] || "";
    const periodoCorrigido = linha[COL_N] || "";
    const portaria = linha[COL_P] || "";

    body.appendParagraph(`Processando errata ${index + 1} de ${valores.length}...`)
        .setFontSize(8)
        .setForegroundColor("#888888");

    const splitPeriodo = s => s.includes("GOZO") ? s.split("GOZO:").map(t => t.trim()) : [s];
    const periodoOrigSplit = splitPeriodo(periodoOriginal);
    const periodoCorrSplit = splitPeriodo(periodoCorrigido);

    body.appendParagraph(`ERRATA na Portaria n° Col.: ${portaria}`)
        .setFontSize(10)
        .setBold(true);
    body.appendParagraph(`Dispõe sobre a concessão de férias regulamentares em relação ao servidor mencionado, Proc.2025/2770950`)
        .setFontSize(10);
    body.appendParagraph(`Nome: ${nome}(COLUNA G), Matrícula n° ${matricula}(COLUNA F)`)
        .setFontSize(10);
    body.appendParagraph(`Onde se lê:Período Aquisitivo:${periodoOrigSplit[0]}(COLUNA M)`);
    if (periodoOrigSplit[1]) body.appendParagraph(`Onde se lê:Período de Gozo:${periodoOrigSplit[1]}(COLUNA M)`);
    body.appendParagraph(`Leia-se:Período Aquisitivo:${periodoCorrSplit[0]}(COLUNA N)`);
    if (periodoCorrSplit[1]) body.appendParagraph(`Leia-se:Período de Gozo:${periodoCorrSplit[1]}(COLUNA N)`);
    body.appendParagraph(`Publicada no Diário Oficial n°. ${portaria}(COLUNA P)`);
    body.appendParagraph("");
  });

  doc.saveAndClose();
  ui.alert("Documento de ERRATAS gerado com sucesso!\n\n" + doc.getUrl());
  Logger.log("Documento de ERRATAS finalizado: " + doc.getUrl());
}
