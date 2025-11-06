//FUNÇÃO PARA VERIFICAR SE O ALERTA JÁ ESTÁ VALIDADO


function verificarAlerta(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var numeroAlerta = ss.getSheetByName("Alerta de Qualidade").getRange("K13").getValue();
  var numLinha = 0
  var buscarAlerta = 0

  for(numLinha = 2; buscarAlerta != numeroAlerta; numLinha++){

    var buscarAlerta = ss.getSheetByName("Alertas Gerados").getRange("b"+numLinha).getValue();

    if(buscarAlerta == numeroAlerta){

      var confirmacao = ss.getSheetByName("Alertas Gerados").getRange(numLinha, 3).getValue();

      if( confirmacao != "Validado!"){

        validarAlerta();

      }else{

        var dataValidacao = ss.getSheetByName("Alertas Gerados").getRange(numLinha, 3).getValue();
          
        var botaoAlerta = ui.alert('Atenção!',
          'Este alerta foi validado no dia: '+ dataValidacao +
          '. Deseja envia-lo via e-mail?',
          ui.ButtonSet.YES_NO);

          if (botaoAlerta == ui.Button.YES) {
          // User clicked "Yes".
            enviarEmail();
          } 
      }

    }
  }
}


//Função para validar alerta

function validarAlerta(){
  
  // pop up de coleta de dados
  SpreadsheetApp.getActiveSpreadsheet().toast('Coletando Dados para assinatura','Validação Alerta de Qualidade',3);

  var usuario = Session.getActiveUser().getEmail();
  var numeroAlerta = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Alerta de Qualidade").getRange("K13").getValue();
  var alertaOP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Alerta de Qualidade").getRange("D14").getValue();
  let agora = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Alerta de Qualidade").getRange("A4").getValue();
  


  //Insere nova linha
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validações Realizadas").insertRowAfter(1);

  // pop up assinatura
  ss.toast( 'Assinando documento em nome de:  '+Session.getActiveUser() ,'Validação Alerta de Qualidade',3);

  //Imputar dados
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validações Realizadas").getRange("A2").setValue(numeroAlerta);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validações Realizadas").getRange("B2").setValue(agora);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validações Realizadas").getRange("C2").setValue(usuario);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validações Realizadas").getRange ("D2").setValue(alertaOP);
  //Coletar formulário a validar
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validações Realizadas").getRange("E2").setValue("Validado!");
  
  //SpreadsheetApp.getUi().alert("Alerta de Qualidade nº" + numeroAlerta + ", Validado com Sucesso!");

  // pop up preparar para enviar e-mail
  ss.toast('Preparando e-mail','Validação Alerta de Qualidade',3);
  
  enviarEmail();
}

// FUNÇÃO PARA ENVIAR EMAIL

function enviarEmail(){
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var alertaValidado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab Referências").getRange("E7").getValue();
  var numeroAlerta = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PDF").getRange("K6").getValue();
  var numeroOrdem= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PDF").getRange("D7").getValue();
  var modeloAlerta = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PDF").getRange("M7").getValue();

//  Utilities.sleep(30)


  if(alertaValidado == 'Validado!'){
    
    //oculta alerta e deixa pdf visível

    planilha.getSheetByName("PDF").showSheet();
    planilha.getSheetByName("Alerta de Qualidade").hideSheet();

    //Constantes de dados dos usuário

    var supervisorEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TAB REFERÊNCIAS").getRange("E3").getValue();
    var analistaEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TAB REFERÊNCIAS").getRange("E4").getValue();
    var inspetorEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TAB REFERÊNCIAS").getRange("E5").getValue();
  
    //selecionar guia para enviar email
    

    var mensagem = {
      to: inspetorEmail, analistaEmail,
      cc: supervisorEmail,
      subject: "Alerta de Qualidade: " + numeroAlerta + " - Op: " + numeroOrdem + " | " + modeloAlerta,
      body: "Olá, você está recebendo o alerta de qualidade número: " + numeroAlerta+", referente à ordem de produção : " + numeroOrdem + "Este alerta de foi emitido por motivo de:  "+modeloAlerta,
      name: Session.getActiveUser(),
      attachments: [planilha.getAs('application/pdf').setName(numeroAlerta +" - "+numeroOrdem +".pdf")]
    }
    //pop up enviar e-mail
    ss.toast('Enviando e-mail','Validação Alerta de Qualidade',3);

    MailApp.sendEmail(mensagem);

    ui.alert("Alerta de qualidade nº " + numeroAlerta + ", ordem nº: " + numeroOrdem + " enviado com sucesso!" );

    //mostrar guias novamente

  }else{

    ui.alert('Erro! Alerta de qualidade nº: '+ numeroAlerta + ', ordem nº: ' + numeroOrdem + ', não está validado! Verifique!' );
  }

  planilha.getSheetByName("Alerta de Qualidade").showSheet();
  planilha.getSheetByName("PDF").hideSheet();

}
