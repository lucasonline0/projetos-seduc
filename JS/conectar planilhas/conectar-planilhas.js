const ID_PLANILHA_ORIGEM = '';
const NOME_ABA_ORIGEM = '';
const ID_PLANILHA_DESTINO = '';
const NOME_ABA_DESTINO = '';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sincronização')
    .addItem('Sincronizar Ingressados', 'sincronizarIngressados')
    .addItem('Verificar Duplicatas', 'verificarDuplicatas')
    .addToUi();
}

function sincronizarIngressados() {
  const planilhaOrigem = SpreadsheetApp.openById(ID_PLANILHA_ORIGEM).getSheetByName(NOME_ABA_ORIGEM);
  const planilhaDestino = SpreadsheetApp.openById(ID_PLANILHA_DESTINO).getSheetByName(NOME_ABA_DESTINO);

  const dadosOrigem = planilhaOrigem.getDataRange().getValues();
  const dadosDestino = planilhaDestino.getDataRange().getValues();

  const cabecalhoOrigem = dadosOrigem[0].map(c => String(c).trim().toUpperCase());
  const cabecalhoDestino = dadosDestino[0].map(c => String(c).trim().toUpperCase());

  const colunasComuns = cabecalhoOrigem
    .map((nome, idxOrigem) => {
      const idxDestino = cabecalhoDestino.findIndex(c => c.includes(nome));
      return idxDestino !== -1 ? { idxOrigem, idxDestino } : null;
    })
    .filter(Boolean);

  const idxMatriculaOrigem = cabecalhoOrigem.findIndex(c => c.includes('MATRICULA'));
  const idxMatriculaDestino = cabecalhoDestino.findIndex(c => c.includes('MATRICULA'));
  const idxStatusOrigem = cabecalhoOrigem.findIndex(c => c.includes('STATUS'));

  if (idxMatriculaOrigem === -1 || idxMatriculaDestino === -1 || idxStatusOrigem === -1) {
    Logger.log("Colunas 'MATRICULA' ou 'STATUS' não encontradas.");
    return;
  }

  const matriculasDestino = new Set(
    dadosDestino.slice(1).map(linha => String(linha[idxMatriculaDestino]).replace(/\s/g, '').toLowerCase())
  );

  let novasLinhas = 0;

  for (let i = 1; i < dadosOrigem.length; i++) {
    const status = String(dadosOrigem[i][idxStatusOrigem]).trim().toLowerCase();
    const matricula = String(dadosOrigem[i][idxMatriculaOrigem]).replace(/\s/g, '').toLowerCase();

    if (status.includes('ingressado') && !matriculasDestino.has(matricula)) {
      const novaLinha = [];
      for (let j = 0; j < cabecalhoDestino.length; j++) {
        const colunaComum = colunasComuns.find(c => c.idxDestino === j);
        novaLinha[j] = colunaComum ? dadosOrigem[i][colunaComum.idxOrigem] : '';
      }
      planilhaDestino.appendRow(novaLinha);
      novasLinhas++;
    }
  }

  Browser.msgBox('Sincronização completa!', `${novasLinhas} novos ingressados adicionados.`, Browser.Buttons.OK);
}

function verificarDuplicatas() {
  const planilhaDestino = SpreadsheetApp.openById(ID_PLANILHA_DESTINO).getSheetByName(NOME_ABA_DESTINO);
  const dados = planilhaDestino.getDataRange().getValues();
  const cabecalho = dados[0].map(c => String(c).trim().toUpperCase());

  const idxMatricula = cabecalho.findIndex(c => c.includes('MATRICULA'));
  const idxNome = cabecalho.findIndex(c => c.includes('NOME'));

  if (idxMatricula === -1 || idxNome === -1) {
    Logger.log("Colunas 'MATRICULA' ou 'NOME' não encontradas.");
    return;
  }

  const duplicados = [];
  const conjunto = {};

  for (let i = 1; i < dados.length; i++) {
    const matricula = String(dados[i][idxMatricula]).replace(/\s/g, '').toLowerCase();
    const nome = String(dados[i][idxNome]).trim().toLowerCase();
    const chave = matricula + '|' + nome;

    if (conjunto[chave]) {
      duplicados.push({linha: i+1, matricula: dados[i][idxMatricula], nome: dados[i][idxNome]});
    } else {
      conjunto[chave] = true;
    }
  }

  if (duplicados.length === 0) {
    Browser.msgBox('Duplicatas', 'Nenhuma duplicata encontrada.', Browser.Buttons.OK);
  } else {
    let mensagem = 'Duplicatas encontradas:\n';
    duplicados.forEach(d => {
      mensagem += `Linha ${d.linha}: ${d.nome} (Matrícula ${d.matricula})\n`;
    });
    Browser.msgBox('Duplicatas', mensagem, Browser.Buttons.OK);
  }
}
