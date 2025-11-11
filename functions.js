Office.onReady(() => {
  const item = Office.context.mailbox.item;

  // Añade handlers para detectar cambios en destinatarios
  if (item.to) item.to.addHandlerAsync(Office.EventType.RecipientsChanged, checkPublicDomains);
  if (item.cc) item.cc.addHandlerAsync(Office.EventType.RecipientsChanged, checkPublicDomains);
  if (item.bcc) item.bcc.addHandlerAsync(Office.EventType.RecipientsChanged, checkPublicDomains);
});

function checkPublicDomains(event) {
  const item = Office.context.mailbox.item;
  const publicDomains = ["gmail.com", "hotmail.com", "yahoo.com", "outlook.com"];

  // Recolecta todos los destinatarios actuales
  let recipients = [];
  if(item.to) recipients = recipients.concat(item.to);
  if(item.cc) recipients = recipients.concat(item.cc);
  if(item.bcc) recipients = recipients.concat(item.bcc);

  const flagged = recipients.filter(r =>
    r.emailAddress && publicDomains.some(d => r.emailAddress.endsWith("@" + d))
  );

  if (flagged.length > 0) {
    Office.context.mailbox.item.notificationMessages.addAsync("publicDomainAlert", {
      type: "warningMessage",
      message: "Se detectaron destinatarios con dominios públicos: " + flagged.map(r => r.emailAddress).join(", ")
    });
  } else {
    // Elimina aviso si ya no hay dominios públicos
    Office.context.mailbox.item.notificationMessages.removeAsync("publicDomainAlert");
  }

  if (event) event.completed();
}
