function emailAutomaticoNovaEntrada() {
  
  criarNumeroAlerta(); // primerio criar ID e depois iniciar o envio do email

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName;

  var lrAlertasGerados = guia("Alertas Gerados").getLastRow();
  var ultimoAlertaEnviado = guia("Alertas Gerados").getRange("F1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

  for (ultimoAlertaEnviado + 1; ultimoAlertaEnviado <= lrAlertasGerados; ultimoAlertaEnviado++) {

    var numLinhaAlerta = ultimoAlertaEnviado;
    var emailAutomativoEnviado = guia("Alertas Gerados").getRange("F" + numLinhaAlerta).getValue(); 
    var numeroAlerta = guia("Alertas Gerados").getRange("B" + numLinhaAlerta).getValue();
    var motivoAlerta = guia("Alertas Gerados").getRange("J" + numLinhaAlerta).getValue();

    if (numeroAlerta != null) { //VERIFICAR SE HÁ VALOR PREENCHIDO NA LINHA SELECIONADA

      if (emailAutomativoEnviado != "Alerta Enviado!") { //VERIFICAR SE O ALERTA JÁ FOI ENVIADO
      
        var numeroOrdem = guia("Alertas Gerados").getRange("K" + numLinhaAlerta).getValue();
        var areaIdentificadora = guia("Alertas Gerados").getRange("I" + numLinhaAlerta).getValue();
        var areaGeradora = guia("Alertas Gerados").getRange("I" + numLinhaAlerta).getValue();
        var emissorAlerta = guia("Alertas Gerados").getRange("H" + numLinhaAlerta).getValue();
          
        guia("Email Automatico").getRange("K6").setValue(numeroAlerta); 
        guia("Tab Referências").getRange("E3").setValue(numeroAlerta); 
        guia("Tab Referências").getRange("E4").setValue(areaGeradora); 

        //PROCURAR LISTA DE EMAIL DA AREA GERADORA
        var linhaCentroDeCusto = 0;
        var lrCentroDeCusto = guia("Tab Referências").getRange("G:G").getLastRow(); 

        let idPlanilhaEmail = null; 

        for (linhaCentroDeCusto = 2; linhaCentroDeCusto <= lrCentroDeCusto; linhaCentroDeCusto++) { 

          var buscaAreaIdentificadora = guia("Tab Referências").getRange("G" + linhaCentroDeCusto).getValue(); 

          if (buscaAreaIdentificadora == areaIdentificadora) { 
            var identificadorAlerta = guia("Tab Referências").getRange("J" + linhaCentroDeCusto).getValue();
          }

          var buscaAreaGeradora = guia("Tab Referências").getRange("G" + linhaCentroDeCusto).getValue(); 

          if (buscaAreaGeradora == areaGeradora) { 
            idPlanilhaEmail = guia("Tab Referências").getRange("H" + linhaCentroDeCusto).getValue();
            var supervisorEmail = guia("Tab Referências").getRange("I" + linhaCentroDeCusto).getValue();
            guia("Lista_Email").getRange("A1").setValue("=" + idPlanilhaEmail);
          }

          if (identificadorAlerta != null && idPlanilhaEmail != null) {
            lrCentroDeCusto = 1;
          }
        }
        
        // coletar e-mails
        let analistaEmail = guia("Tab Referências").getRange("E7").getValue();
        let inspetorEmail = guia("Tab Referências").getRange("E8").getValue();
        let fornecedorEmail = guia("Tab Referências").getRange("E9").getValue();

        // preparar corpo do e-mail
        let corpoEmail = guia("Tab Referências").getRange("K3").getValue();
        corpoEmail = corpoEmail
          .replace(/numeroAlerta/, numeroAlerta)
          .replace(/numeroOrdem/, numeroOrdem)
          .replace(/motivoAlerta/, motivoAlerta);
        
        // preparar telas
        guia("Email Automatico").showSheet(); 
        guia("Alerta de qualidade").hideSheet(); 
        guia("Alertas Gerados").hideSheet(); 
        guia("Tab Referências").hideSheet(); 
        guia("Lista_Email").hideSheet(); 

        Utilities.sleep(600);

        // ✅ MONTAR LISTA DE E-MAILS COM FILTRO
        var listaEmails = [
          analistaEmail,
          inspetorEmail,
          supervisorEmail,
          fornecedorEmail
        ];

        var email_destinatario = listaEmails
          .filter(function(e) {
            return e &&                         
                   e.toString().trim() !== "" && 
                   e.toString().trim() !== "Inexistente" &&
                   e.toString().trim().toUpperCase() !== "#N/A" &&
                   !/^undefined$/i.test(e.toString().trim());
          })
          .join(",");

        if (!email_destinatario) {
          Logger.log("Nenhum destinatário válido encontrado para o alerta " + numeroAlerta);
          continue; 
        }

        if (areaGeradora == "74054 - Central Facção Goian") {
          MailApp.sendEmail({
            to: email_destinatario,
            cc: identificadorAlerta + "," + emissorAlerta,
            subject: "Alerta de Qualidade: " + numeroAlerta + " - Op: " + numeroOrdem + " | " + motivoAlerta,
            body: corpoEmail,
            name: Session.getActiveUser(),
            attachments: [planilha.getAs('application/pdf').setName(numeroAlerta + " - " + numeroOrdem + ".pdf")]
          });
        } else {
          MailApp.sendEmail({
            to: email_destinatario,
            cc: identificadorAlerta + "," + emissorAlerta,
            subject: "Alerta de Qualidade: " + numeroAlerta + " - Op: " + numeroOrdem + " | " + motivoAlerta,
            body: corpoEmail,
            name: Session.getActiveUser(),
            attachments: [planilha.getAs('application/pdf').setName(numeroAlerta + " - " + numeroOrdem + ".pdf")]
          });
        }

        // limpar dados para próximo alerta
        guia("EMAIL AUTOMATICO").getRange("K6").clearContent();
        guia("Tab Referências").getRange("E3:E4").clearContent();
        guia("LISTA_EMAIL").getRange("A1").clearContent();
        guia("Alertas Gerados").getRange("C" + numLinhaAlerta).setValue("Validado!");

        var horaValidacao = guia("Alertas Gerados").getRange("G" + numLinhaAlerta).getValue();
        guia("Alertas Gerados").getRange("D" + numLinhaAlerta).setValue(horaValidacao);
        guia("Alertas Gerados").getRange("E" + numLinhaAlerta).setValue(emissorAlerta);
        guia("Alertas Gerados").getRange("F" + numLinhaAlerta).setValue("Alerta Enviado!");

        guia("Alerta de qualidade").showSheet(); 
        guia("Email Automatico").hideSheet(); 

        identificadorAlerta = null;
        idPlanilhaEmail = null;
        numeroAlerta = null;
        emailAutomativoEnviado = null;
        motivoAlerta = null;
        
        Utilities.sleep(30000);
      }
    }
  }
}
