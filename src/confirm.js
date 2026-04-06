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

// 🔐 MSAL CONFIG
const msalConfig = {
  auth: {
    clientId: "e92a8324-40d8-4ce5-876d-99df6b07acf9",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// 🔥 GET TOKEN GRAPH
async function getGraphToken() {

  const accounts = msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    await msalInstance.loginPopup({
      scopes: ["Mail.Read", "Mail.Send"]
    });
  }

  const response = await msalInstance.acquireTokenSilent({
    scopes: ["Mail.Read", "Mail.Send"],
    account: msalInstance.getAllAccounts()[0]
  });

  return response.accessToken;
}

// 🔥 BOUTON OUI (FINAL)
document.getElementById("btnYes").onclick = async () => {

  try {

    console.log("START");

    const token = await getGraphToken();

    console.log("TOKEN OK");

    // 🔥 GET MAIL
    const mailResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    if (!mailResponse.ok) {
      const err = await mailResponse.text();
      throw new Error(err);
    }

    const eml = await mailResponse.text();

    console.log("MAIL OK");

    // 🔥 ENCODAGE SAFE
    const base64 = btoa(
      new Uint8Array(
        [...eml].map(c => c.charCodeAt(0))
      ).reduce((data, byte) => data + String.fromCharCode(byte), '')
    );

    // 🔥 SEND MAIL
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
          },
          saveToSentItems: true
        })
      }
    );

    if (!sendResponse.ok) {
      const err = await sendResponse.text();
      console.error("GRAPH ERROR:", err);
      throw new Error(err);
    }

    console.log("MAIL SENT ✅");

    Office.context.ui.closeContainer();

  } catch (err) {

    console.error("ERROR:", err);
    alert("Erreur envoi ❌ " + err.message);
  }
};

document.getElementById("btnNo").onclick = () => {
  Office.context.ui.closeContainer();
};
