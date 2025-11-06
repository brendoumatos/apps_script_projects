function myFunction() {
  try {
    // Obtém o e-mail do usuário ativo
    Session.getActiveUser().getEmail();
    
    // Acessa a pasta no Drive
    var pastaFoto = DriveApp.getFolderById("15LH3jfOhBSypkZV6thp4n3i5R1pqHpO9");
    var template = SpreadsheetApp.openById("10tmxkKcBR0ifSyklnygMfhrj0LQj9iGDhIz1bPawkIg").getSheetByName("TEMPLATE");
    var registro_foto01 = template.getRange("C16").getValue();
    var nome_extraido = registro_foto01.toString().split("/");
    var nomefoto01 = String(nome_extraido[1]);

    // Obtém o arquivo de imagem
    var arquivoImagem = pastaFoto.getFilesByName(nomefoto01).next();
    var id_foto01 = arquivoImagem.getId();

    // Obtém o link de visualização direta do arquivo
    var linkImagem = "https://drive.google.com/uc?id=" + id_foto01;

    // Insere o link na célula I16
    template.getRange("I16").setFormula('=IMAGE("' + linkImagem + '", 1, 100, 100)');

    Logger.log("Imagem inserida com sucesso!");
  } catch (e) {
    Logger.log("Erro: " + e.message);

    pastaFoto.secy
  }
}





function inserirImagemNaPlanilha() {
  try {
    // ID da pasta no Google Drive
    const pastaId = '15LH3jfOhBSypkZV6thp4n3i5R1pqHpO9';
    
    // Nome da imagem na pasta

     var template = SpreadsheetApp.openById("10tmxkKcBR0ifSyklnygMfhrj0LQj9iGDhIz1bPawkIg").getSheetByName("TEMPLATE");

    const nomeImagem = template.getRange("C16").getValue();

    // Acessar a pasta no Drive
    const pasta = DriveApp.getFolderById(pastaId);
    const arquivos = pasta.getFilesByName(nomeImagem);

    if (!arquivos.hasNext()) {
      throw new Error('Imagem não encontrada na pasta!');
    }

    // Obter o arquivo de imagem
    const arquivoImagem = arquivos.next();
    const blob = arquivoImagem.getBlob();

    // Acessar a planilha e a aba desejada
    const planilha = SpreadsheetApp.openById('10tmxkKcBR0ifSyklnygMfhrj0LQj9iGDhIz1bPawkIg');
    const aba = planilha.getSheetByName('template');

    // Limpar a célula antes de inserir a nova imagem
    aba.getRange('I16').clearContent();

    // Inserir a imagem na célula
    aba.insertImage(blob, 9, 16); // Coluna I = 9, Linha 16

    Logger.log('Imagem inserida com sucesso!');
  } catch (e) {
    Logger.log('Erro ao inserir imagem: ' + e.message);
  }
}

