function gerar_alerta_qualidade() {

  Utilities.sleep(60);

  try{

    Logger.log("ETAPA 01 - Capturando Variáveis do Projeto");

    //VARIAVEIS GERAIS DO PROJETO
    
    var planilha_alerta = SpreadsheetApp.openById('1CIlbBxQmhUZ9pvHnQlNr9gEKG1JBrDAZCfRgVU6jpTE')
    var guia_alerta_gerado = planilha_alerta.getSheetByName("ALERTAS_GERADOS");
    var guia_template_alerta = planilha_alerta.getSheetByName("TEMPLATE_ALERTA");
    var pasta_foto_alerta = DriveApp.getFolderById("1EeU63fcHWy2h6U7vWv1c1EMMcTyTa8Pc");
    var tab_ref_areas = planilha_alerta.getSheetByName("Tab_Ref_Areas")
    var guia_lista_email = planilha_alerta.getSheetByName("Lista_Email")

    //VARIAVEIS DE CONTROLE DE EXECUÇÃO DO SCRIPT
    var ultimo_alerta_gerado = guia_alerta_gerado.getRange("C1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
    var ultimo_alerta_enviado = guia_alerta_gerado.getRange("B1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();


    //GERANDO NOVOS ID
    Logger.log("ETAPA 02 - Gerando novos IDs");
    for(var linha_atual = ultimo_alerta_enviado + 1; linha_atual <= ultimo_alerta_gerado +1; linha_atual++){

      try{

      
      
        var ultimo_id_gerado = guia_alerta_gerado.getRange("A1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getValue();
        var id_linha_atual = guia_alerta_gerado.getRange("A" + linha_atual).getValue();
        var status_alerta_atual = guia_alerta_gerado.getRange("B" + linha_atual).getValue();
        
        if( id_linha_atual == ''){

          var ultimo_id_extraido = ultimo_id_gerado.split(" / ");
          var id_extraido_numero = Number(ultimo_id_extraido[0])
          var ano_atual = guia_alerta_gerado.getRange("C" + linha_atual).getValue().getFullYear();

          var novo_id = (id_extraido_numero + 1) + " / " + ano_atual;

          Logger.log("ETAPA 02.1 - Gerado novo ID: " + novo_id);

          var id_para_laudo = novo_id;

          guia_alerta_gerado.getRange("A" + linha_atual).setValue(id_para_laudo);

        }else{ //SE FALSO REUTILIZA O MESMO LAUDO

          var id_para_laudo = id_linha_atual
        }


        //PREPARAR PARA GERAR LAUDO
        if( status_alerta_atual == ''){

          Logger.log("Etapa 03 - Coletando dados do alerta");

          var num_ordem = guia_alerta_gerado.getRange("G" + linha_atual).getValue();
          var emissor_alerta = guia_alerta_gerado.getRange("D" + linha_atual).getValue();
          var motivo_alerta = guia_alerta_gerado.getRange("F" + linha_atual).getValue();

          Utilities.sleep(10);

          Logger.log("Etapa 03.1 - Enviando dados para o template");
          
          guia_template_alerta.getRange("K6").setValue(id_para_laudo);

          Utilities.sleep(5)

          Logger.log("ETAPA 04 - Coletando imagens capturadas");

          //capturar diretorio+nomearquivos
          var evidencia01 = guia_alerta_gerado.getRange("M" + linha_atual).getValue();
          var evidencia02 = guia_alerta_gerado.getRange("Q" + linha_atual).getValue();
          var evidencia03 = guia_alerta_gerado.getRange("U" + linha_atual).getValue();
          var evidencia04 = guia_alerta_gerado.getRange("Y" + linha_atual).getValue();
          var evidencia05 = guia_alerta_gerado.getRange("AC" + linha_atual).getValue();
          var evidencia06 = guia_alerta_gerado.getRange("AG" + linha_atual).getValue();
          var evidencia07 = guia_alerta_gerado.getRange("AK" + linha_atual).getValue();
          var evidencia08 = guia_alerta_gerado.getRange("AO" + linha_atual).getValue();

          Utilities.sleep(60)

          Logger.log("ETAPA 05 - Inserindo Imagens ")

          if(evidencia01 != ''){ //if evidencia01 is not null extract values

          var nome_foto01 = evidencia01.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto01 = nome_foto01[1]; //capturar apenas nome
          var arquivo_foto01 = pasta_foto_alerta.getFilesByName(nome_extraido_foto01).next(); //acessar arquivo
          var id_foto01 = arquivo_foto01.getId(); //obter id
          Logger.log("id_foto01 = " + id_foto01); //logar id

          guia_template_alerta.getRange("C21").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto01 + '";1);"")'); //inseri a foto01
          guia_template_alerta.getRange("B21").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto01 +'";1)');

          Logger.log("ETAPA 05.1 - foto01 setada");
          }

          if(evidencia02 != ''){ //if evidencia02 is not null extract values

          var nome_foto02 = evidencia02.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto02 = nome_foto02[1]; //capturar apenas nome
          var arquivo_foto02 = pasta_foto_alerta.getFilesByName(nome_extraido_foto02).next(); //acessar arquivo
          var id_foto02 = arquivo_foto02.getId(); //obter id
          Logger.log("id_foto02 = " + id_foto02); //logar id

          guia_template_alerta.getRange("C22").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto02 + '";1);"")'); //inseri a foto02
          guia_template_alerta.getRange("B22").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto02 +'";2)');

          Logger.log("ETAPA 05.2 - foto02 setada");
          }

          if(evidencia03 != ''){ //if evidencia03 is not null extract values

          var nome_foto03 = evidencia03.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto03 = nome_foto03[1]; //capturar apenas nome
          var arquivo_foto03 = pasta_foto_alerta.getFilesByName(nome_extraido_foto03).next(); //acessar arquivo
          var id_foto03 = arquivo_foto03.getId(); //obter id
          Logger.log("id_foto03 = " + id_foto03); //logar id

          guia_template_alerta.getRange("C23").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto03 + '";1);"")'); //inseri a foto03
          guia_template_alerta.getRange("B23").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto03 +'";3)');

          Logger.log("ETAPA 05.3 - foto03 setada");
          }

          if(evidencia04 != ''){ //if evidencia04 is not null extract values

          var nome_foto04 = evidencia04.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto04 = nome_foto04[1]; //capturar apenas nome
          var arquivo_foto04 = pasta_foto_alerta.getFilesByName(nome_extraido_foto04).next(); //acessar arquivo
          var id_foto04 = arquivo_foto04.getId(); //obter id
          Logger.log("id_foto04 = " + id_foto04); //logar id

          guia_template_alerta.getRange("C24").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto04 + '";1);"")'); //inseri a foto04
          guia_template_alerta.getRange("B24").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto04 +'";4)');

          Logger.log("ETAPA 05.4 - foto04 setada");
          }

          
          if(evidencia05 != ''){ //if evidencia05 is not null extract values

          var nome_foto05 = evidencia05.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto05 = nome_foto05[1]; //capturar apenas nome
          var arquivo_foto05 = pasta_foto_alerta.getFilesByName(nome_extraido_foto05).next(); //acessar arquivo
          var id_foto05 = arquivo_foto05.getId(); //obter id
          Logger.log("id_foto05 = " + id_foto05); //logar id

          guia_template_alerta.getRange("I21").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto05 + '";1);"")'); //inseri a foto05
          guia_template_alerta.getRange("H21").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto05 +'";5)');

          Logger.log("ETAPA 05.5 - foto05 setada");
          }

          if(evidencia06 != ''){ //if evidencia06 is not null extract values

          var nome_foto06 = evidencia06.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto06 = nome_foto06[1]; //capturar apenas nome
          var arquivo_foto06 = pasta_foto_alerta.getFilesByName(nome_extraido_foto06).next(); //acessar arquivo
          var id_foto06 = arquivo_foto06.getId(); //obter id
          Logger.log("id_foto06 = " + id_foto06); //logar id

          guia_template_alerta.getRange("I22").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto06 + '";1);"")'); //inseri a foto06
          guia_template_alerta.getRange("H22").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto06 +'";6)');

          Logger.log("ETAPA 05.6 - foto06 setada");
          }

          if(evidencia07 != ''){ //if evidencia07 is not null extract values

          var nome_foto07 = evidencia07.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto07 = nome_foto07[1]; //capturar apenas nome
          var arquivo_foto07 = pasta_foto_alerta.getFilesByName(nome_extraido_foto07).next(); //acessar arquivo
          var id_foto07 = arquivo_foto07.getId(); //obter id
          Logger.log("id_foto07 = " + id_foto07); //logar id

          guia_template_alerta.getRange("I23").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto07 + '";1);"")'); //inseri a foto07
          guia_template_alerta.getRange("H23").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto07 +'";7)');

          Logger.log("ETAPA 05.7 - foto07 setada");
          }

          if(evidencia08 != ''){ //if evidencia08 is not null extract values

          var nome_foto08 = evidencia08.toString().split("Alerta_Qualidade_Images/");//quebrar texto
          var nome_extraido_foto08 = nome_foto08[1]; //capturar apenas nome
          var arquivo_foto08 = pasta_foto_alerta.getFilesByName(nome_extraido_foto08).next(); //acessar arquivo
          var id_foto08 = arquivo_foto08.getId(); //obter id
          Logger.log("id_foto08 = " + id_foto08); //logar id

          guia_template_alerta.getRange("I24").setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id_foto08 + '";1);"")'); //inseri a foto08
          guia_template_alerta.getRange("H24").setFormula('=HYPERLINK("https://drive.google.com/file/d/' + id_foto08 +'";8)');

          Logger.log("ETAPA 05.8 - foto08 setada");
          }

          
        Utilities.sleep(90); // aguardar 30 segundos para carregar as imagens

          //coletar dados dos remetentes
          Logger.log("ETAPA 06 - Buscando dados dos remetentes");

          var area_geradora = guia_alerta_gerado.getRange("E" + linha_atual).getValue();//captura area que enviou amostra
          var fornecedor = guia_alerta_gerado.getRange("AZ" + linha_atual).getValue();//captura o nome do fornecedor

          //coletar emails
          for( var linhaAreaRef = 2; linhaAreaRef <= tab_ref_areas.getRange("A2").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex(); linhaAreaRef++){//rodar linha a linha para encontrar a igual e retornar planilha analista/inspetor e time qualidade

            let area_linha_atual = tab_ref_areas.getRange("A" + linhaAreaRef).getValue();

            if(area_geradora == area_linha_atual){ //capturar valor da coluna A

              Logger.log("Area Geradora Encontrada: " + area_geradora);
              var id_planilha_fornecedor = tab_ref_areas.getRange("B" + linhaAreaRef).getValue(); //captura o id da planilha
              var email_fixo = tab_ref_areas.getRange("C" + linhaAreaRef).getValue(); // captura os email fixos do time
              var email_time_qualidade = tab_ref_areas.getRange("D" + linhaAreaRef).getValue();

              break
            }
            
          }

          //coletar emails dos analistas e inspetores
          Logger.log("ETAPA 07 - Coletando emails dos Analistas e Inspetores");
          
          guia_lista_email.getRange("A1").setFormula("=" + id_planilha_fornecedor); //seta planilha com nome dos inspetores
          
          Utilities.sleep(10);

          Logger.log("ETAPA 07.1 - Capturando o grupo do Fornecedor");
        
          var fornecedorString = fornecedor.toString().split("-");
          var grupo_fornecedor = Number(fornecedorString[0]);
          Logger.log("Grupo extraído do fornecedor: " + grupo_fornecedor);

          Logger.log("ETAPA 07.1.1 - Validando se é fornecedor interno")

          var conf_interna = area_geradora.toString().indexOf("Confecção Interna") > -1;

          if( conf_interna == true){
            
            let turno_fornecedor = guia_alerta_gerado.getRange("BB" + linha_atual).getValue(); // captura o turno do fornecedor

            grupo_fornecedor = grupo_fornecedor + "-" + turno_fornecedor;
            
          }

          Logger.log("ETAPA 07.2 - Bucando email dos envolvidos")
          //rodar linha por linha buscando o grupo do fornecedor

          let linha_fim_fornecedor = guia_lista_email.getRange("B1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

          for( var linha_ref_fornecedor = 2; linha_ref_fornecedor <= linha_fim_fornecedor ;linha_ref_fornecedor++ ){
            
            let fornecedorRef = guia_lista_email.getRange("A" + linha_ref_fornecedor).getValue(); //captura o grupo atual

            //testa se encontrou o fornecedor
            if(fornecedorRef == grupo_fornecedor || fornecedorRef == undefined){

              Logger.log("Fornecedor encontrado: " + fornecedorRef);

              if(conf_interna == true){

                var email_analista = '';
                var email_inspetor = guia_lista_email.getRange("D" + linha_ref_fornecedor).getValue();
              
              }else{

                var email_analista = guia_lista_email.getRange("C" + linha_ref_fornecedor).getValue();
                var email_inspetor = guia_lista_email.getRange("D" + linha_ref_fornecedor).getValue();
                var email_fornecedor = guia_lista_email.getRange("E" + linha_ref_fornecedor).getValue();

              }
              
              var listaEmails = [
                email_analista,
                email_inspetor,
                email_fornecedor
              ];

              var email_destinatario = listaEmails
                .filter(function(e) {
                  return e &&                         
                        e.toString().trim() !== "" && 
                        e.toString().trim() !== "Inexistente" &&
                        e.toString().trim() !== "Indefinido" &&
                        e.toString().trim() !== undefined &&
                        e.toString().trim().toUpperCase() !== "#N/A" &&
                        !/^undefined$/i.test(e.toString().trim());
                })
                .join(",");
              /*
              if (!email_destinatario) {
                Logger.log("Nenhum destinatário válido encontrado para o alerta " + numeroAlerta);
                continue; 
              }*/

              break
            }else{

              //não encontrou grupo do fornecedor, ignora email dos destinatarios

              email_destinatario = ''
            }

            if( email_destinatario == undefined){

                email_destinatario = ''

            }
          }

          Logger.log('Destinatários: ' + email_destinatario);

          Utilities.sleep(60)

          Logger.log("ETAPA 08 - Gerando PDF")

          
          planilha_alerta.getSheetByName("TEMPLATE_ALERTA").showSheet();
          planilha_alerta.getSheetByName("EMAIL AUTOMATICO").hideSheet();
          planilha_alerta.getSheetByName("ALERTAS_GERADOS").hideSheet();
          planilha_alerta.getSheetByName("ALERTA DE QUALIDADE").hideSheet();
          planilha_alerta.getSheetByName("Tab_Ref_Areas").hideSheet();
          planilha_alerta.getSheetByName("Lista_fornecedores").hideSheet();
          planilha_alerta.getSheetByName("Lista_Email").hideSheet();

          Utilities.sleep(60)

          //setar modelo email para area especifica (confeccao ou nao?)
          var areaConfeccao = area_geradora.indexOf("Confecção") > -1;
          Logger.log(areaConfeccao);
        
          var mensagem;
        
          if(areaConfeccao == true){
          
            mensagem = {
              to:email_fixo + "," + email_destinatario,
              cc: email_time_qualidade + "," + emissor_alerta,
              //to:"llbrendouwilerll@gmail.com" + "," + "brendouwiler_7@hotmail.com",
              subject: "Alerta de Qualidade: " + id_para_laudo + " - Op: " + num_ordem + " | " + motivo_alerta,
              htmlBody: "Olá time!<br><br>" +
                "Por meio deste, encaminhamos o alerta <b>" + id_para_laudo + "</b>, referente à ordem: <b>" + num_ordem + "</b> <br>" +
                "Este foi emitido pelo motivo de: <b>" + motivo_alerta + "</b>" +
                "<br>" +
                "Os detalhes deste alerta estão disponíveis no laudo em anexo <br>" +
                
                "Atenciosamente,<br>" +
                "Time de Qualidade<br><br>",

              name: "vinicius.dos@ciahering.com.br",
              attachments: [planilha_alerta.getAs('application/pdf').setName("Alerta: " + id_para_laudo + " - " + num_ordem +".pdf")]
            };
          }else{

            mensagem = {
              //MENSAGEM PARA DEMAIS AREAS
              to:emailFixo,
              cc: email_time_qualidade + "," + emissor_alerta,
              //to:"llbrendouwilerll@gmail.com" + "," + "brendouwiler_7@hotmail.com",
              subject: "Alerta de Qualidade: " + id_para_laudo + " - Op: " + num_ordem + " | " + motivo_alerta,
              htmlBody: "Olá time!<br><br>" +
                "Por meio deste, encaminhamos o alerta <b>" + id_para_laudo + "</b>, referente à ordem: <b>" + num_ordem + "</b> <br>" +
                "Este foi emitido pelo motivo de: <b>" + motivo_alerta + "</b>" +
                "<br><br>" +
                "Os detalhes deste alerta estão disponíveis no laudo em anexo <br>" +
                "<br><br>"+
                "Atenciosamente,<br>" +
                "Time de Qualidade<br><br>",

              name: "vinicius.dos@ciahering.com.br",
              
              attachments: [planilha_alerta.getAs('application/pdf').setName("Alerta: " + id_para_laudo + " - " + num_ordem +".pdf")]
            };
          }

          Logger.log("ETAPA 09 - Enviando Email")

          MailApp.sendEmail(mensagem);//enviar email

          Utilities.sleep(60);

          guia_alerta_gerado.getRange("B" + linha_atual).setValue("Alerta Enviado!");//informando que o laudo foi enviado
        
          Logger.log("ETAPA 10 - Limpando dados utlizandos");

          guia_template_alerta.getRange("K6").clearContent();
          guia_template_alerta.getRange("B21:I24").clearContent(); //limpa celulas de imagem

          planilha_alerta.getSheetByName("ALERTA DE QUALIDADE").showSheet();
          planilha_alerta.getSheetByName("TEMPLATE_ALERTA").hideSheet();
        }


        Logger.log(ultimo_id_gerado)

      }catch{

        let mensagem_erro = {

          //MENSAGEM PARA DEMAIS AREAS
              to:emissor_alerta,
              subject: "ERRO | Alerta de Qualidade: " + num_ordem,
              htmlBody: "Olá!<br><br>" +
                "Informamos que foi identificado um erro no Alerta de Qualidade referente à ordem: <b>" + num_ordem + "</b> <br>" +
                "Acesse o aplicativo Quality para verificar se as informações preenchidas estão <b>"+
                "<br><br>" +
                "Atenciosamente,<br>" +
                "Time de Qualidade<br><br>",

              name: "vinicius.dos@ciahering.com.br",
              


        };

        if(status_alerta_atual == null){
          
          MailApp.sendEmail(mensagem_erro);
        }

        Logger.log("Erro ao gerar o laudo na linha: " + linha_atual)

        

        continue

      }


    }

  }catch(e){
    Logger.log(e.messagem);
  }

  Logger.log("Finalizando envio dos alertas")

}