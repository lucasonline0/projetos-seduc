function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("💬 Verificação")
      .addItem("Verificar e-mails", "verificarRespostas")
      .addToUi();
  }
  
  function verificarRespostas() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dados = sheet.getDataRange().getValues();
  
    //data e hora de corte: 07/08/2025 às 08:00
    const dataCorte = new Date("2025-08-07T08:00:00");
  
    for (let i = 1; i < dados.length; i++) {
      const email = dados[i][1]; // coluna B
      const statusResposta = (dados[i][2] || "").toString().trim().toUpperCase(); // coluna c
  
      if (!email) continue;
      if (statusResposta === "SIM") continue;
  
      const threads = GmailApp.search(`from:${email} after:${Math.floor(dataCorte.getTime() / 1000)}`);
      let respostaEncontrada = false;
      let trechoResposta = "";
  
      for (const thread of threads) {
        const mensagens = thread.getMessages();
  
        for (const msg of mensagens) {
          if (msg.getFrom().includes(email)) {
            const dataMsg = msg.getDate();
            if (dataMsg >= dataCorte) {
              respostaEncontrada = true;
  
              let corpo = msg.getPlainBody();
  
              corpo = corpo.split(/\n\s*>/)[0];
  
              corpo = corpo.split(/\nEm\s|^De:|^Para:|^Assunto:/mi)[0];
  
              // limpeza basica
              corpo = corpo.trim();
  
              trechoResposta = corpo.substring(0, 200);
              break;
            }
          }
        }
  
        if (respostaEncontrada) break;
      }
  
      if (respostaEncontrada) {
        sheet.getRange(i + 1, 3).setValue("SIM"); // coluna c
        sheet.getRange(i + 1, 4).setValue(trechoResposta); // coluna d
      } else if (statusResposta === "") {
        sheet.getRange(i + 1, 3).setValue("NAO");
      }
    }
  
    SpreadsheetApp.flush();
  }
  