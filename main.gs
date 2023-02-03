function so_request() {
  
  var spsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sh = spsheet.getActiveSheet();
  var row = sh.getLastRow(); 

  var name = sh.getRange(row,3).getValue();
  var pnumber = sh.getRange(row,12).getValue();
  var email = sh.getRange(row,11).getValue();
  var occupation = sh.getRange(row,2).getValue();

  

  var employer = sh.getRange(row,9).getValue();
  var place = sh.getRange(row, 20).getValue();

  switch (occupation) {
    case 'Aluno':

      var date = sh.getRange(row,5).getValue();
      var project = sh.getRange(row,8).getValue();
      var resources = sh.getRange(row,18).getValue();
      var material = sh.getRange(row,19).getValue();
      var place = sh.getRange(row, 20).getValue();
      var file = sh.getRange(row,17).getValue();

      var text = "Solicitante: " + name + "\nTelefone: " + pnumber + "\nEmail: " + email + "\nSolicitação: " + project + "\nIrá utilizar: " + resources +  "\nMaterial: " + material + "\nData de agendamento: " + date + "\nLink do arquivo: " + file + "\nFuncionário responsável: " + employer;

            
      text = encodeURIComponent(text);

      switch (place){
        case 'FabLab acadêmico':
          var chat_id = 01234567;
          message(text, chat_id);
          break;

        case 'FabLab profissional':
          var chat_id = 76543210;
          message(text, chat_id);
          break;
      }

      break;
            
    case 'Externo':
      var date = sh.getRange(row,6).getValue();
      var project = sh.getRange(row,8).getValue();
      var resources = sh.getRange(row,18).getValue();
      var material = sh.getRange(row,19).getValue();
      var place = sh.getRange(row, 20).getValue();
      var destination = sh.getRange(row,16).getValue();
      var file = sh.getRange(row,17).getValue();



      var text = "Solicitante: " + name + "\nTelefone: " + pnumber + "\nEmail: " + email + "\nSolicitação: " + project + "\nIrá utilizar: " + resources +  "\nMaterial: " + material + "\nData de agendamento: " + date + "\nDestino: " + destination +  "\nLink do arquivo: " + file +  "\nFuncionário responsável: " + employer;

            
      text = encodeURIComponent(text);

      switch (place){
        case 'FabLab acadêmico':
          var chat_id = 01234567;
          message(text, chat_id);
          break;

        case 'FabLab profissional':
          var chat_id = 76543210;
          message(text, chat_id);
          break;
      }

      break;
    
    case 'Funcionário':
    
      var date = sh.getRange(row,4).getValue();
      var project = sh.getRange(row,7).getValue();
      var material = sh.getRange(row,24).getValue();
      var quantity = sh.getRange(row,23).getValue();
      var place = sh.getRange(row,25).getValue();
      var destination = sh.getRange(row,15).getValue();
      var file = sh.getRange(row,22).getValue();

      var text = "Solicitante: " + name + "\nTelefone: " + pnumber + "\nEmail: " + email + "\nSolicitação: " + quantity + ", " + project + "\nMaterial: " + material + "\nData Limite: " + date +  "\nDestino: " + destination +  "\nLink do arquivo: " + file + "\nFuncionário responsável: " + employer;

            
      text = encodeURIComponent(text);

      switch (place){
        case 'FabLab acadêmico':
          var chat_id = 01234567;
          message(text, chat_id);
          break;

        case 'FabLab profissional':
          var chat_id = 76543210;
          message(text, chat_id);
          break;
      }

      
      break;
            
    
    }
}
    

function message(text,chat_id)
{
        
        var bot_id = "your-bot-id";
        var url = "https://api.telegram.org/bot" + bot_id + "/sendMessage?chat_id=-" + chat_id + "&text=" + text;
        Logger.log(url);

        UrlFetchApp.fetch(url);
}