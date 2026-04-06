// ===== BOUTON OUI =====
document.getElementById("btnYes").onclick = async () => {

  try {

    console.log("CLICK OUI");

    // 🔐 Token Azure automatique
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true
    });

    console.log("TOKEN OK");

    // 📩 Récupération du mail
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${mailData.itemId}/$value`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    if (!response.ok) {
      throw new Error("Erreur récupération mail");
    }

    const eml = await response.text();

    console.log("MAIL OK");

    // 📦 encodage pièce jointe
    const base64 = btoa(unescape(encodeURIComponent(eml)));

    // 📤 Envoi
    const send = await fetch(
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
                Signalement utilisateur<br>
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
      }
    );

    if (!send.ok) {
      throw new Error("Erreur envoi Graph");
    }

    console.log("MAIL SENT ✅");

    alert("Signalement envoyé ✔");

    Office.context.ui.closeContainer();

  } catch (err) {

    console.error(err);

    alert("Erreur Graph ❌");
  }
};



// ===== BOUTON NON =====
document.getElementById("btnNo").onclick = () => {
  console.log("ANNULATION");
  Office.context.ui.closeContainer();
};
