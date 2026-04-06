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

// 🔥 BOUTON OUI (GRAPH FULL)
document.getElementById("btnYes").onclick = async () => {

  try {

    console.log("START GRAPH");

    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true
    });

    console.log("TOKEN OK");

    // 🔥 RÉCUPÉRATION MAIL VIA ID PASSÉ
    const eml = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    ).then(r => r.text());

    console.log("MAIL OK");

    const base64 = btoa(unescape(encodeURIComponent(eml)));

    // 🔥 ENVOI AU CERT
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
              Mail suspect<br>
              Expéditeur: ${mailData.sender}<br>
              Sujet: ${mailData.subject}
            `
          },
          toRecipients: [
            {
              emailAddress: {
                address: "Primo.DELLASIEGA.external2@test-banque-france.fr"
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

    alert("Signalement envoyé ✔");

    Office.context.ui.closeContainer();

  } catch (err) {

    console.error("ERROR:", err);

    alert("Erreur Graph ❌");
  }
};

document.getElementById("btnNo").onclick = () => {
  Office.context.ui.closeContainer();
};
