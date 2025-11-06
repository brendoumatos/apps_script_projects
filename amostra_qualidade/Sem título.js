function teste() {

    
  //var corpo = mensagem + linkPecaPiloto + "\n\n";
  
   let mensagem = {
        //to:emailAnalista + "," + emailInspetor,
        //cc: emailTimeQualidade,
        to:"llbrendouwilerll@gmail.com" + "," + "brendouwiler_7@hotmail.com",
        subject: "Amostra de Qualidade: " + 123 + " - Op: " + 4567,
        htmlBody: "Olá time!<br><br>" +
        "Por meio do presente, informamos que a amostra <b>" + idParaLaudo + "</b>, referente à ordem: <b>" + ordem + "</b>, foi avaliada e consta no relatório em anexo e também, no aplicativo da " +
        "<a href='https://www.appsheet.com/start/84977eab-d116-4ba1-ba2e-92b77d076fa3?platform=desktop#appName=Pe%C3%A7aPiloto-1001230722-24-03-21&vss=H4sIAAAAAAAAA63PTQrCMBQE4KuUWfcE2Ym4EFEExY1xEZtXCLZJSVK1hJzGhQfpxUz9wY27unwz8DEv4KzosvGiOIHtw_daUAeGwLHtGuJgHFOjvTUVR86xEvUrnNTGeSsySdnaGtn29_5mOCLiIf9gnhxYGGGxP-7KoSRpr0pFdoAHJoFvJNUDkYKfAGKOuvXiWNHzrwTEmLLSFK0juUsjx45zcz27NkLLpZHJL0XlKD4Afo13DKYBAAA=&view=Amostra%20de%20Produ%C3%A7%C3%A3o'>" +
        "PeçaPiloto</a>.<br><br>" +
  
        "<b>Reiteramos sobre observações quanto à produção:</b><br><br>" +
        "<ul>" +
          "<li>Confecção fazer a separação das peças com defeito conforme anexo;</li>" +
          "<li>Fazer reposição do talhado que faltar;</li>" +
           "<li>Sobras de talhados separar, identificar e apontar a quebra;</li>" +
          "<li>Peça já montada identificar como segunda qualidade, não picotar peças prontas;</li>" +
          "<li>Áreas responsáveis, se possível, colocar a devolutiva de respostas via e-mail.</li>" +
        "</ul><br>" +
  
        "Atenciosamente,<br>" +
        "Time de Qualidade<br><br>",


        name: "brendou.matos@ciahering.com.br",
        //attachments: [planilha.getAs('application/pdf').setName("Amostra: " + idParaLaudo + " - " + numeroOrdem +".pdf")]
        };

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tab Referencia').getRange("A10").setValue(mensagem);


     

  MailApp.sendEmail(mensagem)
    
  
  }


