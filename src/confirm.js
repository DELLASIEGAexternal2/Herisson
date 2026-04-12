let mailData = null;

Office.onReady(() => {

  console.log("CONFIRM READY");

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );

  const btnYes = document.getElementById("btnYes");
  const btnNo = document.getElementById("btnNo");

  if (btnYes) btnYes.onclick = sendMail;

  if (btnNo) {
    btnNo.onclick = () => {
      console.log("CLICK NON OK");
      Office.context.ui.closeContainer();
    };
  }

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
      scopes: ["Mail.Read", "Mail.Send"]
    });
    return r.accessToken;
  }
}
function openHelp() {
  Office.context.ui.displayDialogAsync(
    "https://dellasiegaexternal2.github.io/Herisson/support.html",
    { height: 50, width: 40 }
  );
}
function showLoader() {
  const loader = document.getElementById("loader");
  if (loader) loader.style.display = "flex";
}

function hideLoader() {
  const loader = document.getElementById("loader");
  if (loader) loader.style.display = "none";
}

async function sendMail() {

  try {

    console.log("SEND START");

    const btn = document.getElementById("btnYes");

    btn.innerText = "Envoi...";
    btn.disabled = true;
    document.getElementById("btnNo").disabled = true;

    showLoader();

    const token = await getToken();
    console.log("TOKEN OK");

    const mailResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    console.log("MAIL FETCH OK");

    const eml = await mailResponse.text();
    const base64 = btoa(unescape(encodeURIComponent(eml)));

    const comment =
      document.getElementById("comment").value || "Aucun commentaire";

    const sendResponse = await fetch(
      "https://graph.microsoft.com/v1.0/me/sendMail",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          message: {
            subject: "Signalement Hérisson",
            body: {
              contentType: "HTML",
              content: `Expéditeur: ${mailData.sender}<br>Sujet: ${mailData.subject}`
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
      }
    );

    console.log("SEND OK", sendResponse);

    btn.innerText = "✔ Envoyé";

    setTimeout(() => {
      Office.context.ui.closeContainer();
    }, 1200);

  } catch (e) {

    console.error(e);

    hideLoader();

    alert("Erreur ❌ " + e.message);

    document.getElementById("btnYes").innerText = "Oui";
    document.getElementById("btnYes").disabled = false;
    document.getElementById("btnNo").disabled = false;
  }
}
