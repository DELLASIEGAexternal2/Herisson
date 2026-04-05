Office.onReady(async () => {
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );

  displayUser();
});

let mailData = null;

function handleMailData(arg) {
  mailData = JSON.parse(arg.message);
  document.getElementById("sender").innerText = mailData.sender;
  document.getElementById("subject").innerText = mailData.subject;
  document.getElementById("date").innerText = new Date(mailData.date).toLocaleString();
}

async function displayUser() {
 /* const token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
  const me = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` }
  }).then(r => r.json());

  document.getElementById("userInfo").innerText =`${me.displayName} – ${me.mail || me.userPrincipalName}`;*/
}

document.getElementById("btnYes").onclick = async () => {
  const token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });

  const itemId = Office.context.mailbox.convertToRestId(
    Office.context.mailbox.item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );

  const eml = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${itemId}/$value`,
    { headers: { Authorization: `Bearer ${token}` } }
  ).then(r => r.text());

  // ZIP du mail
  const zip = new JSZip();
  zip.file("mail.eml", eml);
  const zipContent = await zip.generateAsync({ type: "base64" });

  // Envoi Graph
  await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      message: {
        subject: "Signalement Hérisson",
        body: {
          contentType: "Text",
          content: `Mail signalé par ${mailData.senderName}`
        },
        toRecipients: [
          { emailAddress: { address: "security@banque-france.fr" } }
        ],
        attachments: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: "mail.zip",
            contentType: "application/zip",
            contentBytes: zipContent
          }
        ]
      }
    })
  });

  Office.context.ui.messageParent("OK");
  Office.context.ui.closeContainer();
};

document.getElementById("btnNo").onclick = () => {
  Office.context.ui.closeContainer();
};
