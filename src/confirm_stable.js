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

  try {

    mailData = JSON.parse(arg.message);

    document.getElementById("sender").innerText = mailData.sender || "-";
    document.getElementById("subject").innerText = mailData.subject || "-";
    document.getElementById("date").innerText =
      new Date(mailData.date).toLocaleString() || "-";

  } catch (e) {
    console.error("Erreur data:", e);
  }
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


// 🔥 ENVOI PRO AVEC .EML + COMMENTAIRE
async function sendMail() {

  try {

    if (!mailData) throw new Error("Aucune donnée mail");

    const token = await getToken();

    // 🔥 Récupération du mail réel
    const mailResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    if (!mailResponse.ok) {
      throw new Error("Erreur récupération mail");
    }

    const eml = await mailResponse.text();

    // 🔥 Encodage robuste
    const base64 = btoa(
      new Uint8Array([...eml].map(c => c.charCodeAt(0)))
        .reduce((data, byte) => data + String.fromCharCode(byte), '')
    );

    // 🔥 commentaire utilisateur
    const comment = document.getElementById("comment")?.value || "Aucun";

    // 🔥 ENVOI
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
            subject: "🚨 Signalement Hérisson",
            body: {
              contentType: "HTML",
              content: `
                <b>Signalement utilisateur</b><br><br>
                Expéditeur: ${mailData.sender}<br>
                Sujet: ${mailData.subject}<br>
                Date: ${mailData.date}<br><br>
                <b>Commentaire:</b><br>
                ${comment}
              `
            },
            toRecipients: [
              {
                emailAddress: {
                  address: "PrimoSylvestreDELLASIEGA-NKOUME@dscoie091.onmicrosoft.com"
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
      }
    );

    if (!response.ok) {
      const err = await response.text();
      throw new Error(err);
    }

    alert("Signalement envoyé ✔");

    Office.context.ui.closeContainer();

  } catch (e) {

    console.error(e);
    alert("Erreur ❌ " + e.message);
  }
}
