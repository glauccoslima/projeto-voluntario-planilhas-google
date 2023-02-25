function enviarEmailFinalMes() {
  Logger.log("Iniciando função enviarEmailFinalMes.");
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Planilha ativa: " + sheet.getName());
  
  var emails = sheet.getEditors().map(function(user) {
    return user.getEmail();
  });
  Logger.log("Lista de editores: " + emails.join(', '));
  
  var dataAtual = new Date();
  Logger.log("Data atual: " + dataAtual);
  
  var assunto = "Atenção! Último dia para enviar os pedidos ao site jw.org";
  var mensagem = "Prezados irmãos, \n\nVenho lembrá-los de que hoje é o último dia para enviar os pedidos ao site (https://hub.jw.org). \n\nSe os pedidos já foram enviados para o site (https://hub.jw.org), por favor, desconsiderem este e-mail. \n\nCaso ainda não tenham enviado os pedidos, por favor, informe ao irmão Ewerton Martins. \n\nEle estará disponível para ajudá-los e garantir que os pedidos sejam enviados corretamente. \n\nEste é um e-mail automático de lembrete, portanto, não é necessário responder. \n\nAtenciosamente, Seus irmãos Central Águas Claras.";
  
  // Solicita permissão para enviar e-mails em nome do usuário
  Logger.log("Solicitando permissão para enviar e-mails em nome do usuário.");
  MailApp.getRemainingDailyQuota();
  
  // Envia o e-mail para os usuários
  Logger.log("Enviando e-mail para os usuários.");
  MailApp.sendEmail(emails.join(','), assunto, mensagem);
    
  Logger.log("E-mail enviado para: " + emails.join(', '));
  
  Logger.log("Função enviarEmailFinalMes concluída.");
}

function criarDisparador() {
  var hoje = new Date();
  var ultimoDiaMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0).getDate();
  
  ScriptApp.newTrigger('enviarEmailFinalMes')
    .timeBased()
    .onMonthDay(ultimoDiaMes)
    .atHour(16)
    .create();
  
  Logger.log("Disparador criado com sucesso para o dia " + ultimoDiaMes + " às 16h.");
}
