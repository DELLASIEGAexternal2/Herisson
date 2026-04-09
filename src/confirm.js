let mailData = null;

Office.onReady(() => {
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );

  document.getElementById("btnYes").onclick = sendMail;
  document.getElementById("btnNo").onclick = () =>
    Office.context.ui.closeContainer();
});

function handleMailData(arg) {
  mailData = JSON.parse(arg.message);
  document.getElementById("sender").innerText = mailData.sender || "—";
  document.getElementById("subject").innerText = mailData.subject || "—";
  document.getElementById("date").innerText =
    new Date(mailData.date).toLocaleString();
}

/* MSAL */
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

async function sendMail() {
  try {
    const token = await getToken();

    const mailResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    const eml = await mailResponse.text();
    const base64 = btoa(unescape(encodeURIComponent(eml)));

    const comment =
      document.getElementById("comment").value || "Aucun commentaire";

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
              <b>Signalement utilisateur</b><br><br>
              Expéditeur : ${mailData.sender}<br>
              Sujet : ${mailData.subject}<br>
              Date : ${mailData.date}<br><br>
              <b>Commentaire :</b><br>${comment}
            `
          },
          toRecipients: [
            {
              emailAddress: {
                address:
                  "PrimoSylvestreDELLASIEGA-NKOUME@dscoie091.onmicrosoft.com"
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

    alert("Signalement envoyé ✔");
    Office.context.ui.closeContainer();
  } catch (e) {
    alert("Erreur ❌ " + e.message);
  }
}