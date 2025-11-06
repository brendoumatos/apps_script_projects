function alertaFeedback() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName;
  var numeroAlerta = guia("Alerta de Qualidade").getRange("K13").getValue();
  var numLinhaAlerta = 0
  var lrAlertasGerados = guia("Alertas Gerados").getLastRow();
  
  // preparar para enviar e-mail

  for(numLinhaAlerta = 2; numLinhaAlerta <= lrAlertasGerados ;numLinhaAlerta++ ){

    var buscarAlerta = guia("Alertas Gerados").getRange("B"+numLinhaAlerta).getValue();
    var emailAutomativoEnviado = guia("Alertas Gerados").getRange("F"+numLinhaAlerta).getValue();
    var motivoAlerta = guia("Alertas Gerados").getRange("K"+numLinhaAlerta).getValue();
    var numeroOrdem = guia("Alertas Gerados").getRange("L"+numLinhaAlerta).getValue();

    if(buscarAlerta == numeroAlerta){

      if(emailAutomativoEnviado != "Alerta Enviado!" ){

        if(motivoAlerta == "Feedback"){

          guia("Email Automatico").getRange("K6").setValue(numeroAlerta); // inseri número do alerta de qualidade na tela de email automatico
          guia("Tab Referências").getRange("T2").setValue(numeroAlerta); // inseri número do alerta de qualidade na tela de email automatico
        
          //enviar email com anexo para analista, inspetor, supervisores

          //coletar email dos usuários e incluir na tabela de referência para efetuar procv
          var supervisorEmail = guia("Tab Referências").getRange("T5").getValue();
          var analistaEmail = guia("Tab Referências").getRange("T6").getValue();
          var inspetorEmail = guia("Tab Referências").getRange("T7").getValue();
          var analistaQualidadeEmail = guia("Tab Referências").getRange("E8").getValue();
          var analistaQualidadeCampo = guia("Tab Referências").getRange("E9").getValue();
          var emissorAlerta = guia("Alertas Gerados").getRange("H"+numLinhaAlerta).getValue();

          //preparar corpo do email
          var corpoEmail = guia("Tab Referências").getRange("V2").getValue();
          var replaceNumeroAlerta = corpoEmail.replace(/numeroAlerta/,numeroAlerta);
          var replaceNumeroOrdem = replaceNumeroAlerta.replace(/numeroOrdem/,numeroOrdem);
          corpoEmail = replaceNumeroOrdem.replace(/motivoAlerta/,motivoAlerta);
        
          //setar informação que o email foi validado e por quem
          guia("Alertas Gerados").getRange("C"+ numLinhaAlerta).setValue("Validado!");

          var horaValidacao = guia("Alertas Gerados").getRange("G"+numLinhaAlerta).getValue();

          guia("Alertas Gerados").getRange("D"+numLinhaAlerta).setValue(horaValidacao);
          guia("Alertas Gerados").getRange("E"+numLinhaAlerta).setValue(emissorAlerta);

          //preparar telas para enviar email
          guia("Email Automatico").showSheet(); // mostrar guia Email automatico

          guia("Alerta de qualidade").hideSheet(); // oculta guia alerta de qualidade

          var mensagem = {
            to: inspetorEmail +","+analistaEmail+","+emissorAlerta,
            cc: supervisorEmail + ","+analistaQualidadeEmail +","+analistaQualidadeCampo,
            subject: "Alerta de Qualidade: " + numeroAlerta + " - Op: " + numeroOrdem + " | " + motivoAlerta,
            body: corpoEmail,
            name: Session.getActiveUser(),
            attachments: [planilha.getAs('application/pdf').setName(numeroAlerta +" - "+numeroOrdem +".pdf")]
          }

          MailApp.sendEmail(mensagem);//enviar email

          // criar função para informar que o email foi enviado e salvar na ALERTAS GERADOS

          guia("EMAIL AUTOMATICO").getRange("K6").clearContent();
          guia("Tab Referências").getRange("T2").clearContent();

          guia("Alertas Gerados").getRange("F"+numLinhaAlerta).setValue("Alerta Enviado!");

          //
          guia("Alerta de qualidade").showSheet(); // mostrar guia alerta de qualidade
          guia("Email Automatico").hideSheet(); // ocultar guia Email automatico
        }

      }
      
    }
  }
}