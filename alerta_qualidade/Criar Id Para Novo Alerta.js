function criarNumeroAlerta() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guia = planilha.getSheetByName; // Substitua pelo nome da guia onde as respostas do formulário são armazenadas.
  var lrAlertasGerados = guia("Alertas Gerados").getLastRow();

  //pega dados do ultimo ID gerados (Alerta anterior)
  var linhaUltimoIdGerado = guia("Alertas Gerados").getRange("B1").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

  for(linhaBuscaAlerta = linhaUltimoIdGerado; linhaBuscaAlerta < lrAlertasGerados; linhaBuscaAlerta++){

    var anoUltimoIdGerado = guia("Alertas Gerados").getRange(linhaBuscaAlerta, 7).getValue().getFullYear();

    //pega o ano para o novo alerta gerados (ultima linha)
    var anoUltimoAlertaGerado = guia("Alertas Gerados").getRange(linhaBuscaAlerta + 1,7).getValue().getFullYear();


    var ultimoAlertaGerado = guia("Alertas Gerados").getRange("B" + linhaBuscaAlerta).getValue(); // pega o ultimo alerta gerado

    var numeroIdExtraido = ultimoAlertaGerado.toString().split(" / ");

    var numeroIdUltimoAlerta = Number(numeroIdExtraido[0]); // extrai o número (ID) do último alerta

    var novoID = (numeroIdUltimoAlerta + 1) + " / " + anoUltimoAlertaGerado; // gera um novo ID

    // caso o ano seja novo, reiniciar a contagem

    if( anoUltimoAlertaGerado != anoUltimoIdGerado ){
    
      guia("Alertas Gerados").getRange(linhaBuscaAlerta + 1, 2).setValue(1 + " / " + anoUltimoAlertaGerado);

    }else{

      guia("Alertas Gerados").getRange(linhaBuscaAlerta + 1, 2).setValue(novoID);

    }

    
  }

  
}