function enviarEmailAprovador() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  var nomeSolicitante = sheet.getRange(lastRow, 4).getValue(); // Ajuste a coluna conforme necessário
  var emailSolicitante = sheet.getRange(lastRow, 7).getValue();
  var linkArquivo = sheet.getRange(lastRow, 3).getValue();
  var emailAprovador = sheet.getRange(lastRow, 6).getValue(); // Ajuste a coluna do e-mail do aprovador
  var comentario = sheet.getRange(lastRow, 5).getValue();
  var nomeProduto = sheet.getRange(lastRow, 2).getValue();
  
  var formAprovacaoUrl = 'https://docs.google.com/forms/d/e/1FAIpQLScAFKdKB5ub0oKCVAuWkRBTZPkOBRZ3YH6bR02NuAFHoy4Chg/viewform?usp=dialog'; 
  // Insira o link do Formulário de Aprovação
  
  // Criando link pré-preenchido para o Formulário de Aprovação
  var linkFormPreenchido = formAprovacaoUrl + "?usp=pp_url" +
    '&entry.469813180=' + encodeURIComponent(nomeSolicitante) + 
    '&entry.1536346065=' + encodeURIComponent(emailSolicitante) +
    '&entry.1462655805=' + encodeURIComponent(linkArquivo);

  // Enviar e-mail para o aprovador
  MailApp.sendEmail({
    to: emailAprovador,
    subject: "Aprovação de Arquivo - " + nomeSolicitante,
    body: "Olá,\n\nVocê tem um arquivo para aprovação enviado por " + nomeSolicitante + "\n\nO produto a qual esta relacionado o pedido é:" + nomeProduto + ".\n\nClique aqui para ver qual arquivo:\n" + linkArquivo +  "\n\nForam enviada essas considerações para o pedido:" + comentario + "\n\nAcesse o formulário para aprovar ou rejeitar:\n" + linkFormPreenchido + "\n\nAtenciosamente,\nSistema de Aprovação"
  });
