function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gerador DREs')
    .addItem('Gerar Todas as DREs', 'gerarTodasDREs')
    .addItem('Gerar DREs Selecionadas', 'gerarDREsSelecionadas')
    .addToUi();
}

function gerarTodasDREs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirmar Geração',
    'Deseja gerar planilhas para todas as DREs?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    exportarDREsParaPlanilhas([]);
  }
}

// Função para gerar DREs selecionadas
function gerarDREsSelecionadas() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'DREs Selecionadas',
    'Digite as DREs que deseja gerar, separadas por vírgula. Exemplo: DRE CASTANHAL, DRE BELEM 1, DRE BELEM 2',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const dresSelecionadas = response.getResponseText()
      .split(',')
      .map(dre => dre.trim())
      .filter(dre => dre !== '');
    
    if (dresSelecionadas.length > 0) {
      exportarDREsParaPlanilhas(dresSelecionadas);
    } else {
      ui.alert('Erro', 'Nenhuma DRE válida foi informada.');
    }
  }
}

// Lista completa de todas as DREs
const listaDREsCompleta = [
  "DRE XINGUARA",
  "DRE TUCURUI",
  "DRE ABAETETUBA",
  "DRE AFUA",
  "DRE ALTAMIRA",
  "DRE ANANINDEUA 1",
  "DRE ANANINDEUA 2",
  "DRE ANANINDEUA 3",
  "DRE ANANINDEUA 4",
  "DRE ANANINDEUA 5",
  "DRE BELEM 1",
  "DRE BELEM 10",
  "DRE BELEM 2",
  "DRE BELEM 3",
  "DRE BELEM 4",
  "DRE BELEM 5",
  "DRE BELEM 6",
  "DRE BELEM 7",
  "DRE BELEM 8",
  "DRE BELEM 9",
  "DRE BENEVIDES",
  "DRE BRAGANCA",
  "DRE BREVES",
  "DRE CACHOEIRA DO ARARI",
  "DRE CAMETA",
  "DRE CAPANEMA",
  "DRE CAPITAO POCO",
  "DRE CASTANHAL",
  "DRE CONCEICAO DO ARAGUAIA",
  "DRE CURRALINHO",
  "DRE ITAITUBA",
  "DRE MAE DO RIO",
  "DRE MARABA",
  "DRE MARACANA",
  "DRE MONTE ALEGRE",
  "DRE OBIDOS",
  "DRE PARAUAPEBAS",
  "DRE SANTA BARBARA",
  "DRE SANTA IZABEL DO PARA",
  "DRE SANTAREM"
];

function exportarDREsParaPlanilhas(dresSelecionadas = []) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBase = ss.getSheetByName("BASE_EFETIVOS");
  const dados = abaBase.getDataRange().getValues();

  Logger.log("Cabeçalho da BASE_EFETIVOS: " + dados[0]);

  const pastaId = "";
  const pastaDestino = DriveApp.getFolderById(pastaId);

  const listaDREs = dresSelecionadas.length > 0 ? dresSelecionadas : listaDREsCompleta;

  //CRIA PASTAS COM DATA E HORA
  const dataHora = new Date();
  const nomePasta = `DREs_${dataHora.getFullYear()}-${String(dataHora.getMonth() + 1).padStart(2, '0')}-${String(dataHora.getDate()).padStart(2, '0')}_${String(dataHora.getHours()).padStart(2, '0')}-${String(dataHora.getMinutes()).padStart(2, '0')}`;
  
  let novaPasta;
  if (pastaId) {
    novaPasta = pastaDestino.createFolder(nomePasta);
  } else {
    novaPasta = DriveApp.getRootFolder().createFolder(nomePasta);
  }

  const colunasDesejadas = [26, 2, 4, 8, 5, 6, 13, 10, 17, 25];

  const cabecalho = [
    "DRE", "MUNICIPIO", "SETOR", "SERVIDOR", "MATRICULA",
    "VINCULO", "CARGO", "DT_EXERCICIO", "TERMINO ESTAGIO", "TIPO DE PROCESSO"
  ];

  let totalGeradas = 0;

  listaDREs.forEach(dre => {
    const linhasFiltradas = dados.filter((linha, index) => {
      if (index === 0) return false;

      const colunaDRE = linha[26];   // Coluna AA
      const colunaT = linha[19];     // REALIZA EST PROB?
      const colunaX = linha[23];     // STATUS_GERAL
      const colunaAB = linha[27];    // ANO_HOM
      const tipoProcesso = linha[25]; // TIPOS DE PROCESSOS
    

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
      novaPasta.addFile(arquivo);
      DriveApp.getRootFolder().removeFile(arquivo);
      
      totalGeradas++;
    }
  });

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    `Foram geradas ${totalGeradas} planilhas na pasta "${nomePasta}".`,
    'Geração Concluída',
    ui.ButtonSet.OK
  );
}

function removerAcentos(texto) {
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
//script by @lucasonline0