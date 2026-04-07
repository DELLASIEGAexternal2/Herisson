Office.onReady(() => {

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );

});

window.onload = () => {

  const btnYes = document.getElementById("btnYes");
  const btnNo = document.getElementById("btnNo");

  btnYes.onclick = sendMail;
  btnNo.onclick = () => Office.context.ui.closeContainer();
};

let mailData = null;

function handleMailData(arg) {

  mailData = JSON.parse(arg.message);

  document.getElementById("sender").innerText = mailData.sender || "-";
  document.getElementById("subject").innerText = mailData.subject || "-";
  document.getElementById("date").innerText =
    new Date(mailData.date).toLocaleString() || "-";
}


// 🔐 MSAL
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: "e92a8324-40d8-4ce5-876d-99df6b07acf9",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin
  }
});

async function getToken() {

  try {
    const r = await msalInstance.acquireTokenSilent({
      scopes: ["Mail.Read", "Mail.Send"]
    });
    return r.accessToken;

  } catch {

    const r = await msalInstance.loginPopup({
      scopes: ["Mail.Read", "Mail.Send"],
      prompt: "select_account"
    });

    return r.accessToken;
  }
}


// 🔥 ENVOI SIMPLE (SANS PIECE JOINTE POUR GARANTIR ENVOI)
async function sendMail() {

  try {

    const token = await getToken();

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/sendMail",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          message: {
            subject: "🚨 TEST SIGNAL HERISSON",
            body: {
              contentType: "Text",
              content: `
Expéditeur: ${mailData.sender}
Sujet: ${mailData.subject}
              `
            },
            toRecipients: [
              {
                emailAddress: {
                  address: "edellasiegankoume724@gmail.com"
                }
              }
            ]
          }
        })
      }
    );

    if (!response.ok) {
      const err = await response.text();
      alert("Erreur envoi: " + err);
      return;
    }

    alert("MAIL ENVOYÉ ✔");

    Office.context.ui.closeContainer();

  } catch (e) {

    alert("Erreur: " + e.message);
  }
}
