/* =====================================================
   CONFIRM.JS — VERSION FINALE STABLE
   ===================================================== */

let mailData = null;
let isSending = false;

/* =========================
   INIT
   ========================= */
Office.onReady(() => {
  console.log("CONFIRMATION PRÊTE");

  // Réception des données depuis commands.js
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );

  const btnYes = document.getElementById("btnYes");
  const btnNo = document.getElementById("btnNo");
  const helpBtn = document.getElementById("helpBtn");

  if (btnYes) btnYes.onclick = sendMail;

  if (btnNo) {
    btnNo.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      Office.context.ui.closeContainer();
    };
  }

  if (helpBtn) {
    helpBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      openHelp();
    };
  }
});

/* =========================
   DATA MAIL
   ========================= */
function handleMailData(arg) {
  mailData = JSON.parse(arg.message);

  console.log("MAIL DATA:", mailData);

  document.getElementById("sender").innerText =
    mailData.sender || "—";

  document.getElementById("subject").innerText =
    mailData.subject || "—";

  document.getElementById("date").innerText =
    mailData.date
      ? new Date(mailData.date).toLocaleString()
      : "—";
}

/* =========================
   MSAL
   ========================= */
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

/* =========================
   LOADER
   ========================= */
function showLoader() {
  const loader = document.getElementById("loader");
  if (loader) loader.style.display = "flex";
}

function hideLoader() {
  const loader = document.getElementById("loader");
  if (loader) loader.style.display = "none";
}

/* =========================
   HELP
   ========================= */
function openHelp() {
  Office.context.ui.displayDialogAsync(
    "https://dellasiegaexternal2.github.io/Herisson/support.html",
    { height: 50, width: 40 }
  );
}

/* =========================
   SEND MAIL
   ========================= */
async function sendMail(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }

  if (isSending) return;
  isSending = true;

  const btnYes = document.getElementById("btnYes");
  const btnNo = document.getElementById("btnNo");

  try {
    if (!mailData || !mailData.itemId) {
      throw new Error("Mail non chargé correctement");
    }

    btnYes.innerText = "Envoi...";
    btnYes.disabled = true;
    btnNo.disabled = true;

    showLoader();

    const token = await getToken();
    console.log("TOKEN OK");

    // Récupération du mail en EML
    const mailResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    if (!mailResponse.ok) {
      throw new Error("Erreur récupération mail");
    }

    const eml = await mailResponse.text();
    const base64 = btoa(unescape(encodeURIComponent(eml)));

    const comment =
      document.getElementById("comment").value || "Aucun commentaire";

    // Envoi du signalement
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
      }
    );

    if (!sendResponse.ok) {
      throw new Error("Erreur envoi Graph");
    }

    console.log("SEND OK");

    btnYes.innerText = "✔ Envoyé";

    // POPUP INFORMATION (APRÈS SUCCÈS)
    Office.context.ui.displayDialogAsync(
      "https://dellasiegaexternal2.github.io/Herisson/src/popup-info.html",
      {
        height: 35,
        width: 50,
        displayInIframe: true
      }
    );

    // Fermeture de la fenêtre de confirmation
    setTimeout(() => {
      Office.context.ui.closeContainer();
    }, 300);

  } catch (e) {
    console.error(e);

    hideLoader();

    alert("Erreur ❌ " + e.message);

    btnYes.innerText = "Oui";
    btnYes.disabled = false;
    btnNo.disabled = false;

    isSending = false;
  }
}
