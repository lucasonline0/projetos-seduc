function exportarDREsParaPlanilhas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBase = ss.getSheetByName("BASE_EFETIVOS");
  const dados = abaBase.getDataRange().getValues();

  Logger.log("Cabeçalho da BASE_EFETIVOS: " + dados[0]);

  const pastaId = "";
  const pastaDestino = DriveApp.getFolderById(pastaId);

  const listaDREs = [
   
  ];

  const colunasDesejadas = [26, 2, 4, 8, 5, 6, 13, 10, 17, 25];

  const cabecalho = [
    "DRE", "MUNICIPIO", "SETOR", "SERVIDOR", "MATRICULA",
    "VINCULO", "CARGO", "DT_EXERCICIO", "TERMINO ESTAGIO", "TIPO DE PROCESSO"
  ];

  listaDREs.forEach(dre => {
    const linhasFiltradas = dados.filter((linha, index) => {
      if (index === 0) return false;

      const colunaDRE = linha[26];   // Coluna AA
      const colunaT = linha[19];     // REALIZA EST PROB?
      const colunaX = linha[23];     // STATUS_GERAL
      const colunaAB = linha[27];    // ANO_HOM
      const tipoProcesso = linha[25]; // Coluna Z: TIPOS DE PROCESSOS
    

      const colunaDREClean = colunaDRE ? colunaDRE.toString().trim() : "";
      const colunaTClean = colunaT ? colunaT.toString().trim() : "";
      const colunaXClean = colunaX ? colunaX.toString().trim() : "";
      const colunaABClean = colunaAB ? colunaAB.toString().trim() : "";
      const tipoProcessoClean = tipoProcesso
        ? removerAcentos(tipoProcesso.toString().trim().toUpperCase())
        : "";

      const condicao = (
        colunaDREClean === dre &&
        colunaTClean === "Sim" &&
        colunaXClean === "Não Homologado" &&
        colunaABClean === "" &&
        tipoProcessoClean === "ETAPA UNICA"
      );

      if (condicao) {
        Logger.log(`? Encontrado para DRE: ${dre} | Linha: ${index + 1}`);
      }

      return condicao;
    });

    Logger.log(`?? DRE ${dre} - Total linhas filtradas: ${linhasFiltradas.length}`);

    if (linhasFiltradas.length > 0) {
      const novaPlanilha = SpreadsheetApp.create(dre);
      const novaAba = novaPlanilha.getActiveSheet();

      const dadosFiltrados = linhasFiltradas.map(linha =>
        colunasDesejadas.map(i => linha[i])
      );

      const dadosParaInserir = [cabecalho, ...dadosFiltrados];
      novaAba.getRange(1, 1, dadosParaInserir.length, dadosParaInserir[0].length)
        .setValues(dadosParaInserir);

      novaAba.setFrozenRows(1);
      novaAba.getRange(1, 1, 1, cabecalho.length)
        .setFontWeight("bold")
        .setHorizontalAlignment("center");

      novaAba.autoResizeColumns(1, cabecalho.length);
      novaAba.setColumnWidth(8, 150);
      novaAba.setColumnWidth(9, 180);
      novaAba.setColumnWidth(10, 200);

      const rangeTotal = novaAba.getRange(1, 1, dadosParaInserir.length, cabecalho.length);
      rangeTotal.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

      const arquivo = DriveApp.getFileById(novaPlanilha.getId());
      pastaDestino.addFile(arquivo);
      DriveApp.getRootFolder().removeFile(arquivo);
    }
  });
}

// Função para remover acentos
function removerAcentos(texto) {
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
