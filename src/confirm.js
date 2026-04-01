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

        const token = await OfficeRuntime.auth.getAccessToken();

        console.log("TOKEN:", token);

        await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                message: {
                    subject: "Signalement Hérisson",
                    body: {
                        contentType: "HTML",
                        content: "<b>Mail suspect signalé</b>"
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address: "cert@entreprise.com"
                            }
                        }
                    ]
                }
            })
        });

        Office.context.ui.messageParent("YES");
        Office.context.ui.closeContainer();

    } catch (err) {
        console.error("Erreur Graph:", err);
    }
};

/**
 * BOUTON NON
 */
document.getElementById("btnNo").onclick = () => {
    Office.context.ui.messageParent("NO");
    Office.context.ui.closeContainer();
};
