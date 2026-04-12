let mailData = null;

Office.onReady(() => {
  console.log("CONFIRM READY");

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );

  const btnYes = document.getElementById("btnYes");
  const btnNo = document.getElementById("btnNo");
  const helpBtn = document.getElementById("helpBtn");

  if (btnYes) btnYes.onclick = sendMail;

  // ✅ FIX NON
  if (btnNo) {
    btnNo.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      console.log("CLICK NON");
      Office.context.ui.closeContainer();
    };
  }

  // ✅ FIX AIDE
  if (helpBtn) {
    helpBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      console.log("CLICK HELP");
      openHelp();
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
async function sendMail() {
  try {
    const btn = document.getElementById("btnYes");
    btn.innerText = "Envoi...";
    btn.disabled = true;
    document.getElementById("btnNo").disabled = true;

    showLoader();

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
    });

    btn.innerText = "✔ Envoyé";

    setTimeout(() => {
      Office.context.ui.closeContainer();
    }, 1200);

  } catch (e) {
    hideLoader();

    alert("Erreur ❌ " + e.message);

    document.getElementById("btnYes").innerText = "Oui";
    document.getElementById("btnYes").disabled = false;
    document.getElementById("btnNo").disabled = false;
  }
}
