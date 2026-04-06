Office.onReady(() => {
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );
});

let mailData = null;

function handleMailData(arg) {

  mailData = JSON.parse(arg.message);

  document.getElementById("sender").innerText = mailData.sender;
  document.getElementById("subject").innerText = mailData.subject;
  document.getElementById("date").innerText = new Date(mailData.date).toLocaleString();
}

// 🔥 VERSION DEMO STABLE (SANS GRAPH)
document.getElementById("btnYes").onclick = () => {

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["Primo.DELLASIEGA.external2@test-banque-france.fr"],
    subject: "🚨 Signalement Hérisson",
    htmlBody: `
      Mail suspect signalé<br>
      Expéditeur: ${mailData.sender}<br>
      Sujet: ${mailData.subject}
    `
  });

  Office.context.ui.closeContainer();
};

document.getElementById("btnNo").onclick = () => {
  Office.context.ui.closeContainer();
};
