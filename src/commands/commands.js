/* global Office */

Office.onReady(() => {});

/**
 * Ouvre la popup Hérisson avec les infos du mail
 */
function openConfirmDialog(event) {

    const url = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

    const item = Office.context.mailbox.item;

    const mailData = {
        sender: item.from?.displayName || item.from?.emailAddress || "(expéditeur inconnu)",
        subject: item.subject || "(sans sujet)",
        date: item.dateTimeCreated
            ? new Date(item.dateTimeCreated).toLocaleString()
            : "(date inconnue)"
    };

    Office.context.ui.displayDialogAsync(
        url,
        {
            height: 70,
            width: 60
        },
        function (asyncResult) {

            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                const dialog = asyncResult.value;

                // Envoi des données vers la popup
                setTimeout(() => {
                    dialog.messageChild(JSON.stringify({
                        type: "init",
                        data: mailData
                    }));
                }, 300);

                // Réponse utilisateur
                dialog.addEventHandler(
                    Office.EventType.DialogMessageReceived,
                    function (arg) {

                        try {
                            const msg = JSON.parse(arg.message);

                            if (msg.confirm === true) {

                                console.log("✔ Envoi confirmé");

                                // FUTUR : envoyer mail CERT ici

                            } else {
                                console.log("❌ Envoi annulé");
                            }

                        } catch (e) {
                            console.error("Erreur parsing message popup");
                        }

                        dialog.close();
                    }
                );
            }

            event.completed();
        }
    );
}

window.openConfirmDialog = openConfirmDialog;
