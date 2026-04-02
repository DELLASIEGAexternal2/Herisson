Office.onReady(() => {
    console.log("Dialog ready");
});

/**
 * Réception des données du parent
 */
Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg) => {
        const data = JSON.parse(arg.message);

        document.getElementById("sender").innerText  = data.sender;
        document.getElementById("subject").innerText = data.subject;
        document.getElementById("date").innerText    = data.date;
    }
);

/**
 * BOUTON OUI
 */
document.getElementById("btnYes").onclick = async () => {

    try {

        // 🔐 Token utilisateur
        const token = await OfficeRuntime.auth.getAccessToken();

        const itemId = Office.context.mailbox.item.itemId;

        // 🔥 1. récupérer le mail en MIME
        const mimeResponse = await fetch(
            `https://graph.microsoft.com/v1.0/me/messages/${itemId}/$value`,
            {
                headers: {
                    Authorization: "Bearer " + token
                }
            }
        );

        const mimeContent = await mimeResponse.text();

        // 🔥 2. encoder en base64
        const base64 = btoa(unescape(encodeURIComponent(mimeContent)));

        // 🔥 3. envoyer au CERT
        await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
            method: "POST",
            headers: {
                Authorization: "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                message: {
                    subject: "🚨 Signalement Hérisson",
                    body: {
                        contentType: "HTML",
                        content: "Mail suspect en pièce jointe"
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

        console.log("✔ Mail envoyé au CERT");

        Office.context.ui.messageParent("YES");
        Office.context.ui.closeContainer();

    } catch (err) {
        console.error("Erreur:", err);
    }
};

/**
 * BOUTON NON
 */
document.getElementById("btnNo").onclick = () => {
    Office.context.ui.messageParent("NO");
    Office.context.ui.closeContainer();
};
