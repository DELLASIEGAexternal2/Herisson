Office.onReady(() => {

  console.log("Office READY");

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    handleMailData
  );

  document.addEventListener("DOMContentLoaded", () => {

    console.log("DOM READY");

    document.getElementById("btnYes").onclick = sendMail;
    document.getElementById("btnNo").onclick = () => {
      Office.context.ui.closeContainer();
    };

  });

});

let mailData = null;

function handleMailData(arg) {

  try {

    mailData = JSON.parse(arg.message);

    console.log("DATA:", mailData);

    document.getElementById("sender").innerText = mailData.sender || "-";
    document.getElementById("subject").innerText = mailData.subject || "-";
    document.getElementById("date").innerText =
      new Date(mailData.date).toLocaleString() || "-";

  } catch (e) {
    console.error("DATA ERROR:", e);
  }
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

// 🔥 ENVOI
async function sendMail() {

  try {

    console.log("START SEND");

    const token = await getGraphToken();

    console.log("TOKEN OK");

    const mailResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    if (!mailResponse.ok) {
      throw new Error(await mailResponse.text());
    }

    const eml = await mailResponse.text();

    const base64 = btoa(
      new Uint8Array([...eml].map(c => c.charCodeAt(0)))
        .reduce((data, byte) => data + String.fromCharCode(byte), '')
    );

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

    if (!sendResponse.ok) {
      throw new Error(await sendResponse.text());
    }

    console.log("MAIL SENT ✅");

    Office.context.ui.closeContainer();

  } catch (err) {

    console.error("ERROR:", err);
    alert("Erreur ❌ " + err.message);
  }
}
