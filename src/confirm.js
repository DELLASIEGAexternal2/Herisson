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

document.getElementById("btnYes").onclick = async () => {
  try {
    console.log("START SIGNAL");

    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true
    });

    console.log("TOKEN OK");

    const itemId = Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    // 📩 Récupération mail
    const eml = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${itemId}/$value`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    ).then(r => r.text());

    console.log("MAIL OK");

    // 📦 ZIP simple (sans lib externe)
    const base64 = btoa(unescape(encodeURIComponent(eml)));

    // 📤 Envoi
    await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        message: {
          subject: "🚨 Signalement Hérisson",
          body: {
            contentType: "HTML",
            content: `
              <b>Mail signalé</b><br/>
              Expéditeur: ${mailData.senderName}<br/>
              Objet: ${mailData.subject}
            `
          },
          toRecipients: [
            {
              emailAddress: {
                address: "Primo.DELLASIEGA@icloud.com"
              }
            }
          ],
          attachments: [
            {
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: "mail.eml",
              contentType: "message/rfc822",
              contentBytes: base64
            }
          ]
        }
      })
    });

    console.log("MAIL SENT ✅");

    Office.context.ui.messageParent("OK");
    Office.context.ui.closeContainer();

  } catch (err) {
    console.error("ERROR:", err);
    alert("Erreur envoi");
  }
};

document.getElementById("btnNo").onclick = () => {
  Office.context.ui.closeContainer();
};
