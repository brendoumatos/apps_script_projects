function testegerar_laudo_amostra_registro_area(){
  

  Utilities.sleep(60);
  try{

  Logger.log("ETAPA 01 - Capturando Variáveis do Projeto");

  //VARIAVEIS GERAIS DO PROJETO
  var planilha = SpreadsheetApp.openById("1gG-ocY8HXiVD33sjziCnaiArNgVyK_HG95dzvcePPv0")
  var guia_amostra_registro = planilha.getSheetByName("Amostras_Registro")
  var guia_templante = planilha.getSheetByName("Template")
  var pastaFoto = DriveApp.getFoldersByName('Amostras_Registro_Images')
  var tabReferencias = planilha.getSheetByName("Tab Referencia")
  var listaEmail = planilha.getSheetByName("ListaEmail")

  //variaveis dde ultimos realizados
  var lr_laudo_gerado = guia_amostra_registro.getRange(1,3).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
  var lr_laudo_enviado = guia_amostra_registro.getRange(1,2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

    Logger.log("ETAPA 02 - Gerando Id Para Novos Laudos");
  //criar id para cada laudo gerado
  for(var linha_atual = lr_laudo_enviado; linha_atual <= lr_laudo_gerado; linha_atual++){
    
    //verificar necessidade de um novo id
    var lr_id_gerado = guia_amostra_registro.getRange(1,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
    var ultimo_id_amostra = guia_amostra_registro.getRange(lr_id_gerado, 1).getValue();
    var id_linha_atual = guia_amostra_registro.getRange(linha_atual, 1).getValue();

    if(id_linha_atual == ''){
      
      var id_extraido_amostra = ultimo_id_amostra.split(" / ");
      var id_amostra = Number(id_extraido_amostra[0]);
      var ano_atual = guia_amostra_registro.getRange(linha_atual, 3).getValue().getFullYear();
      
      var novo_id = (id_amostra + 1) + " / " + ano_atual;
     
     
      
      guia_amostra_registro.getRange(linha_atual, 1).setValue(novo_id);
      Logger.log("ETAPA 02.1 - Novo Id Gerado:" + novo_id);

      Logger.log("ETAPA 03 - Enviando dados do novo laudo para o template")
      //enviar novo id e area solicitante para template
      guia_templante.getRange("K6").setValue(novo_id);
      tabReferencias.getRange("E3").setValue(novo_id);

      var idParaLaudo = novo_id;

    }else{
      Logger.log("ETAPA 03 - Enviando dados do novo laudo para o template")
      //enviar novo id e area solicitante para template
      guia_templante.getRange("K6").setValue(id_linha_atual);
      tabReferencias.getRange("E3").setValue(id_linha_atual);
      
      var idParaLaudo = id_linha_atual;

      Logger.log("ETAPA 03.1 - ID a Utilizar: " + idParaLaudo);

    }

    Logger.log("ETAPA 04 - Verificando emissão do laudo");

    //verificar se ja foi gerado o laudo
    var status_laudo = guia_amostra_registro.getRange(linha_atual, 2).getValue();

    if(status_laudo == ''){ //se o status estiver vazio, gerar laudo

      Logger.log("ETAPA 04 - Enviando dados do novo laudo para o template")
      //enviar novo id e area solicitante para template
      guia_templante.getRange("K6").setValue(idParaLaudo);
      tabReferencias.getRange("E3").setValue(idParaLaudo);

      Utilities.sleep(5);

      Logger.log("ETAPA 05 - Obtendo dados das imagems")
      //verificar se há foto ativa e buscar id
        
        //capturar diretorio+nomearquivos
        var evidencia01 = guia_amostra_registro.getRange("J" + linha_atual).getValue();
        var evidencia02 = guia_amostra_registro.getRange("N" + linha_atual).getValue();
        var evidencia03 = guia_amostra_registro.getRange("R" + linha_atual).getValue();
        var evidencia04 = guia_amostra_registro.getRange("V" + linha_atual).getValue();
        
        if(evidencia01 != ''){ //if evidencia01 is not null extract values

          var nome_foto01 = evidencia01.toString().split("G.Q - São Luis/Quality/Amostras Qualidade/Amostras_Registro_Images/");//quebrar texto
          var nome_extraido_foto01 = nome_foto01[1]; //capturar apenas nome
          var arquivo_foto01 = pastaFoto.getFilesByName(nome_extraido_foto01).next(); //acessar arquivo
          var id_foto01 = arquivo_foto01.getId(); //obter id
          Logger.log("id_foto01 = " + id_foto01); //logar id

          guia_templante.getRange("C16").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto01 + '";1);"")'); //inseri a foto01
          guia_templante.getRange("B16").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto01 +'";1)');

          Logger.log("ETAPA 05.1 - foto01 setada");
        }

      if(evidencia02 != ""){ //if evidencia02 is not null extract values

          var nome_foto02 = evidencia02.toString().split("G.Q - São Luis/Quality/Amostras Qualidade/Amostras_Registro_Images/");//quebrar texto
          var nome_extraido_foto02 = nome_foto02[1]; //capturar apenas nome
          var arquivo_foto02 = pastaFoto.getFilesByName(nome_extraido_foto02).next(); //acessar arquivo
          var id_foto02 = arquivo_foto02.getId(); //obter id
          Logger.log("id_foto02 = " + id_foto02); //logar id

          guia_templante.getRange("I16").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto02 + '";1);"")'); //inseri a foto02
          guia_templante.getRange("H16").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto02 +'";2)');

          
          Logger.log("ETAPA 05.2 - foto02 setada");
        }

        if(evidencia03 != ""){ //if evidencia03 is not null extract values

          var nome_foto03 = evidencia03.toString().split("/");//quebrar texto
          var nome_extraido_foto03 = nome_foto03[1]; //capturar apenas nome
          var arquivo_foto03 = pastaFoto.getFilesByName(nome_extraido_foto03).next(); //acessar arquivo
          var id_foto03 = arquivo_foto03.getId(); //obter id
          Logger.log("id_foto03 = " + id_foto03); //logar id

          guia_templante.getRange("C17").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto03 + '";1);"")'); //inseri a foto03
          guia_templante.getRange("B17").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto03 +'";3)');

          Logger.log("ETAPA 05.3 - foto03 setada");
        }

        if(evidencia04 != ""){ //if evidencia04 is not null extract values

          var nome_foto04 = evidencia01.toString().split("/");//quebrar texto
          var nome_extraido_foto04 = nome_foto04[1]; //capturar apenas nome
          var arquivo_foto04 = pastaFoto.getFilesByName(nome_extraido_foto04).next(); //acessar arquivo
          var id_foto04 = arquivo_foto04.getId(); //obter id
          Logger.log("id_foto04 = " + id_foto04); //logar id
          
          guia_templante.getRange("I17").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto04 + '";1);"")'); //inseri a foto04
          guia_templante.getRange("H16").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto04 +'";3)')


         Logger.log("ETAPA 05.4 - foto04 setada");
        }

      Utilities.sleep(90); // aguardar 30 segundos para carregar as imagens

      //coletar dados dos remetenter
      Logger.log("ETAPA 06 - Buscando dados dos remetentes");

      var area_solicitante = guia_amostra_registro.getRange("F" + linha_atual).getValue();//captura area que enviou amostra
      var fornecedor = guia_amostra_registro.getRange("AD" + linha_atual).getValue();//captura o nome do fornecedor

      //coletar emails
      for( var linhaAreaRef = 2; linhaAreaRef <= tabReferencias.getRange("G2").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex(); linhaAreaRef++){//rodar linha a linha para encontrar a igual e retornar planilha analista/inspetor e time qualidade

        if(area_solicitante == tabReferencias.getRange("G"+linhaAreaRef).getValue()){ //capturar valor da coluna G

          Logger.log("Area Solicitante Encontrada: " + area_solicitante);
          var idPlanilhaFornecedor = tabReferencias.getRange("H" + linhaAreaRef).getValue(); //captura o id da planilha
          var emailFixo = tabReferencias.getRange("I" + linhaAreaRef).getValue(); // captura os email fixos do time
          var emailTimeQualidade = tabReferencias.getRange("J" + linhaAreaRef).getValue();

        } 
      }

      //coletar emails dos analistas e inspetores
      Logger.log("ETAPA 07 - Coletando emails dos Analistas e Inspetores");
        
      listaEmail.getRange("A1").setFormula("=" + idPlanilhaFornecedor); //seta planilha com nome dos inspetores
        
      Utilities.sleep(10);

      Logger.log("ETAPA 07.1 - Capturando o grupo do Fornecedor");
      
      var fornecedorString = fornecedor.toString().split("-");
      var gpFornecedor = Number(fornecedorString[0]);
      Logger.log("Grupo extraído do fornecedor: " + gpFornecedor);

      Logger.log("ETAPA 07.2 - Bucando email dos envolvidos")
      //rodar linha por linha buscando o grupo do fornecedor
      for( var linhaFornecedorRef = 2; linhaFornecedorRef < listaEmail.getRange("A1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex() ;linhaFornecedorRef++ ){
          
        var fornecedorRef = Number(listaEmail.getRange("A" + linhaFornecedorRef).getValue()); //captura o grupo atual

        //testa se encontrou o fornecedor
        if(fornecedorRef == gpFornecedor){

          Logger.log("Fornecedor encontrado: " + fornecedorRef);
          var emailAnalista = listaEmail.getRange("C" + linhaFornecedorRef).getValue();
          var emailInspetor = listaEmail.getRange("D" + linhaFornecedorRef).getValue();
            
        }

      }

      Utilities.sleep(60);
      Logger.log("ETAPA 08 - Preparando PDF");
      
      var ordem = guia_amostra_registro.getRange("E" + linha_atual).getValue();//captura numerio da ordem
      
      //ocultando guias;  
      planilha.getSheetByName("Amostras_Registro").hideSheet();
      planilha.getSheetByName("Amostras_Avaliadas").hideSheet();
      planilha.getSheetByName("Tab Referencia").hideSheet();
      planilha.getSheetByName("ListaEmail").hideSheet();
      planilha.getSheetByName("Lista Fornecedores").hideSheet();
      planilha.getSheetByName("Template").activate();

      //gerando pdf
      var numeroOrdem = guia_amostra_registro.getRange("E" + linha_atual).getValue();

      //setar modelo email para area especifica (confeccao ou nao?)
      var areaConfeccao = area_solicitante.indexOf("Confecção") > -1;
      Logger.log(areaConfeccao);
      
      let mensagem;
      
      if(areaConfeccao == true){
        
        mensagem = {
          to:emailAnalista + "," + emailInspetor + "," + emailFixo,
          cc: emailTimeQualidade,
          //to:"llbrendouwilerll@gmail.com" + "," + "brendouwiler_7@hotmail.com",
          subject: "Amostra de Qualidade: " + idParaLaudo + " - Op: " + ordem,
          htmlBody: "Olá time!<br><br>" +
            "Por meio deste, informamos que a amostra <b>" + idParaLaudo + "</b>, referente à ordem: <b>" + ordem + "</b>, foi avaliada e consta no relatório em anexo e, também no aplicativo da " +
            "<a href='https://www.appsheet.com/start/84977eab-d116-4ba1-ba2e-92b77d076fa3?platform=desktop#appName=Pe%C3%A7aPiloto-1001230722-24-03-21&vss=H4sIAAAAAAAAA63PTQrCMBQE4KuUWfcE2Ym4EFEExY1xEZtXCLZJSVK1hJzGhQfpxUz9wY27unwz8DEv4KzosvGiOIHtw_daUAeGwLHtGuJgHFOjvTUVR86xEvUrnNTGeSsySdnaGtn29_5mOCLiIf9gnhxYGGGxP-7KoSRpr0pFdoAHJoFvJNUDkYKfAGKOuvXiWNHzrwTEmLLSFK0juUsjx45zcz27NkLLpZHJL0XlKD4Afo13DKYBAAA=&view=Amostra%20de%20Produ%C3%A7%C3%A3o'>" +
            "Peça Piloto</a>.<br><br>" +
  
            "<b>Reiteramos sobre observações quanto à produção:</b><br>" +
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
          attachments: [planilha.getAs('application/pdf').setName("Amostra: " + idParaLaudo + " - " + numeroOrdem +".pdf")]
        };
      }else{

        mensagem = {
          //MENSAGEM PARA DEMAIS AREAS
          to:emailFixo,
          cc: emailTimeQualidade,
          //to:"llbrendouwilerll@gmail.com" + "," + "brendouwiler_7@hotmail.com",
          subject: "Amostra de Qualidade: " + idParaLaudo + " - Op: " + ordem,
          htmlBody: "Olá time!<br><br>" +
            "Por meio deste, informamos que a amostra <b>" + idParaLaudo + "</b>, referente à ordem: <b>" + ordem + "</b>, foi devidamento avaliada e encontra-se no relatório em anexo, bem como disponível no aplicativo da " +
            "<a href='https://www.appsheet.com/start/84977eab-d116-4ba1-ba2e-92b77d076fa3?platform=desktop#appName=Pe%C3%A7aPiloto-1001230722-24-03-21&vss=H4sIAAAAAAAAA63PTQrCMBQE4KuUWfcE2Ym4EFEExY1xEZtXCLZJSVK1hJzGhQfpxUz9wY27unwz8DEv4KzosvGiOIHtw_daUAeGwLHtGuJgHFOjvTUVR86xEvUrnNTGeSsySdnaGtn29_5mOCLiIf9gnhxYGGGxP-7KoSRpr0pFdoAHJoFvJNUDkYKfAGKOuvXiWNHzrwTEmLLSFK0juUsjx45zcz27NkLLpZHJL0XlKD4Afo13DKYBAAA=&view=Amostra%20de%20Produ%C3%A7%C3%A3o'>" +
            "Peça Piloto</a>.<br><br>" +
            "<br>"+
            "Enstamos encaminhando a amostra para seja realizada uma <b> análise criteriosa e profunda </b>, visando identificar com precisão oas possíveis origens e causas do defeito constatado <br>"+
            "Reforçamos a necessidade de uma <b> devolutiva estruturada </b>, acompanhada de uma <b> apresentação (.ppt)</b>, que nos permita respaldar tecnicamente as informações e, <b> subsidiar a interlocução com as área produtivas </b>.<br>"+
            "<br>"+
            "Atenciosamente,<br>" +
            "Time de Qualidade<br><br>",

          name: "brendou.matos@ciahering.com.br",
          attachments: [planilha.getAs('application/pdf').setName("Amostra: " + idParaLaudo + " - " + numeroOrdem +".pdf")]
        };

      }
      Logger.log("ETAPA 09 - Enviando Email");  
  
      MailApp.sendEmail(mensagem);//enviar email

      Utilities.sleep(60);

      guia_amostra_registro.getRange("B" + linha_atual).setValue("Enviado");//informando que o laudo foi enviado
      

    }

    Logger.log("ETAPA 10 - Limpar dados");

    guia_templante.getRange("K6").clearContent();
    guia_templante.getRange("B16:I17").clearContent();
    listaEmail.getRange("A1").clearContent();
    tabReferencias.getRange("E3").clearContent();
    tabReferencias.getRange("E4").clearContent();
    
  }

  }catch(e) {
    Logger.log(e.messagem);
    }
    
}