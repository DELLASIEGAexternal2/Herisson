Office.onReady(() => {

    initButtons();

    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        (arg) => {

            const data = JSON.parse(arg.message);

            document.getElementById("sender").innerText  = data.sender;
            document.getElementById("subject").innerText = data.subject;
            document.getElementById("date").innerText    = data.date;
        }
    );
});

function initButtons() {

    document.getElementById("btnYes").onclick = async () => {

        try {

            const token = await OfficeRuntime.auth.getAccessToken({
                allowSignInPrompt: true
            });

            const itemId = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0
            );

            const mimeResponse = await fetch(
                `https://graph.microsoft.com/v1.0/me/messages/${itemId}/$value`,
                {
                    headers: {
                        Authorization: "Bearer " + token
                    }
                }
            );

            const mimeContent = await mimeResponse.text();

            const base64 = btoa(
                new TextEncoder().encode(mimeContent)
                    .reduce((data, byte) => data + String.fromCharCode(byte), '')
            );

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

            Office.context.ui.messageParent("YES");
            Office.context.ui.closeContainer();

        } catch (err) {

            console.error(err);

            Office.context.ui.messageParent("ERROR");
            Office.context.ui.closeContainer();
        }
    };

    document.getElementById("btnNo").onclick = () => {
        Office.context.ui.messageParent("NO");
        Office.context.ui.closeContainer();
    };
}
