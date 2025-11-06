function novoEmailAutomaticoNovaEntrada() {
  
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

        if (motivoAlerta == "Inspeção na Origem" || motivoAlerta == "Rejeito") { //VERIFICAR SE O ALERTA SE ENQQUADRA NOS QUESITOS

          var numeroOrdem = guia("Alertas Gerados").getRange("K" + numLinhaAlerta).getValue();
          var areaIdentificadora = guia("Alertas Gerados").getRange("I" + numLinhaAlerta).getValue();
          var areaGeradora = guia("Alertas Gerados").getRange("I" + numLinhaAlerta).getValue();
          var emissorAlerta = guia("Alertas Gerados").getRange("H" + numLinhaAlerta).getValue();

          guia("Email Automatico").getRange("K6").setValue(numeroAlerta); // inseri número do alerta de qualidade na tela de email automatico
          guia("Tab Referências").getRange("E3").setValue(numeroAlerta); // inseri número do alerta de qualidade na tela de email automatico
          guia("Tab Referências").getRange("E4").setValue(areaGeradora); //inseri a área que gerou o problema

          //PROCURAR LISTA DE EMAIL DA AREA GERADORA

          var linhaCentroDeCusto = 0;
          var lrCentroDeCusto = guia("Tab Referências").getRange("G:G").getLastRow(); //PEGAR ÚLTIMA LINHA PREENCHIDA COM OS CENTRO DE CUSTOS

          let idPlanilhaEmail = null; // Redefinir a variável no início de cada iteração

          for (linhaCentroDeCusto = 2; linhaCentroDeCusto <= lrCentroDeCusto; linhaCentroDeCusto++) { // PROCURAR CENTRO DE CUSTO E CHECAR COM ÁREA GERADORA

            //PROCURAR AREA QUE IDENTIFICOU O PROBLEMA E COPIAR EMAIL
            var buscaAreaIdentificadora = guia("Tab Referências").getRange("G" + linhaCentroDeCusto).getValue();

            if (buscaAreaIdentificadora == areaIdentificadora) { //VERIFICAR SE DADOS ENCONTRADOS CONFEREM COM INFORMADOS

              var identificadorAlerta = guia("Tab Referências").getRange("J" + linhaCentroDeCusto).getValue();

            }

            //PROCURAR AREA QUE IDENTIFICOU O PROBLEMA E COPIAR EMAIL

            var buscaAreaGeradora = guia("Tab Referências").getRange("G" + linhaCentroDeCusto).getValue();

            if (buscaAreaGeradora == areaGeradora) { //VERIFICAR SE DADOS ENCONTRADOS CONFEREM COM INFORMADOS

              idPlanilhaEmail = guia("Tab Referências").getRange("H" + linhaCentroDeCusto).getValue();

              guia("Lista_Email").getRange("A1").setValue("=" + idPlanilhaEmail); // SETAR FÓRMULA PARA LISTA DE EMAIL DA ÁREA CORRESPONDENTE

            }
            // SE JÁ TIVER ENCONTRADO, PARAR BUSCA
            if (identificadorAlerta != null && idPlanilhaEmail != null) {
              lrCentroDeCusto = 1;
            }
          }

          //enviar email com anexo para analista, inspetor, supervisores

          //coletar email dos usuários e incluir na tabela de referência para efetuar procv
          let supervisorEmail = guia("Tab Referências").getRange("E6").getValue();
          let analistaEmail = guia("Tab Referências").getRange("E7").getValue();
          let inspetorEmail = guia("Tab Referências").getRange("E8").getValue();

          //preparar corpo do email
          let corpoEmail = guia("Tab Referências").getRange("K3").getValue();
          let replaceNumeroAlerta = corpoEmail.replace(/numeroAlerta/, numeroAlerta);
          let replaceNumeroOrdem = replaceNumeroAlerta.replace(/numeroOrdem/, numeroOrdem);
          corpoEmail = replaceNumeroOrdem.replace(/motivoAlerta/, motivoAlerta);

          //preparar telas para enviar email
          guia("Email Automatico").showSheet(); // mostrar guia Email automatico
          guia("Alerta de qualidade").hideSheet(); // oculta guia alerta de qualidade

          let mensagem = {
            to: inspetorEmail + "," + analistaEmail,
            cc: supervisorEmail + "," + identificadorAlerta,
            subject: "Alerta de Qualidade: " + numeroAlerta + " - Op: " + numeroOrdem + " | " + motivoAlerta,
            body: corpoEmail,
            name: Session.getActiveUser(),
            attachments: [planilha.getAs('application/pdf').setName(numeroAlerta + " - " + numeroOrdem + ".pdf")]
          };

          MailApp.sendEmail(mensagem);//enviar email

          // criar função para informar que o email foi enviado e salvar na ALERTAS GERADOS

          guia("EMAIL AUTOMATICO").getRange("K6").clearContent();
          guia("Tab Referências").getRange("E3:E4").clearContent();
          guia("LISTA_EMAIL").getRange("A1").clearContent();

          //setar informação que o email foi enviado e por quem
          guia("Alertas Gerados").getRange("C" + numLinhaAlerta).setValue("Validado!");

          var horaValidacao = guia("Alertas Gerados").getRange("G" + numLinhaAlerta).getValue();

          guia("Alertas Gerados").getRange("D" + numLinhaAlerta).setValue(horaValidacao);
          guia("Alertas Gerados").getRange("E" + numLinhaAlerta).setValue(emissorAlerta);
          guia("Alertas Gerados").getRange("F" + numLinhaAlerta).setValue("Alerta Enviado!");

          //
          guia("Alerta de qualidade").showSheet(); // mostrar guia alerta de qualidade
          guia("Email Automatico").hideSheet(); // ocultar guia Email automatico

          //LIMPAR VARIAVEIS PARA PRÓXIMO ALERTA

          identificadorAlerta = null;
          idPlanilhaEmail = null;
          numeroAlerta = null;
          emailAutomativoEnviado = null;
          motivoAlerta = null;
        }
      }
    }
  }
}
