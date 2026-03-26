/* global Office */
Office.onReady(() => {});

/**
 * Ouvre la popup Hérisson en GRAND MODAL Outlook.
 * La popup utilisée = confirm.html (inchangée).
 */
function openConfirmDialog(event) {

    const url = "https://dellasiegaexternal2.github.io/Herisson/src/cfrm.html";

    // 1) Récupération du mail en cours
    const item = Office.context.mailbox.item;

    const mailData = {
        sender:
            item.from?.displayName ||
            item.from?.emailAddress ||
            "(expéditeur inconnu)",

        subject: item.subject || "(sans sujet)",

        date:
            item.dateTimeCreated
                ? new Date(item.dateTimeCreated).toLocaleString()
                : "(date inconnue)"
    };

    // 2) Ouvrir la popup modale (grande fenêtre)
    Office.context.ui.displayDialogAsync(
        url,
        {
            height: 70,       // Grande fenêtre
            width: 60,        // Large
            requireHTTPS: true,
            displayInIframe: true
        },
        (asyncResult) => {

            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                const dialog = asyncResult.value;

                // 3) Envoyer les données du mail vers confirm.html
                //    (ton HTML attend: sender, subject, date)
                setTimeout(() => {
                    dialog.messageChild(JSON.stringify({
                        type: "init",
                        data: mailData
                    }));
                }, 400); // petit délai Outlook obligatoire

                // 4) Écouter la réponse YES / NO depuis confirm.js
                dialog.addEventHandler(
                    Office.EventType.DialogMessageReceived,
                    (arg) => {

                        try {
                            const msg = JSON.parse(arg.message);

                            if (msg.confirm === true) {
                                console.log("Utilisateur a confirmé l’envoi Hérisson.");
                                // ICI ON ENVOIE AU CERT (Graph ou SMTP)
                                // → on fera cette partie étape suivante
                            } else {
                                console.log("Utilisateur a annulé l’envoi.");
                            }
                        } catch (e) {
                            console.error("Message popup invalide :", arg.message);
                        }

                        dialog.close();
                    }
                );
            }

            // Outlook exige que la commande soit libérée
            event.completed();
        }
    );
}

/* --- Exposer la fonction --- */
if (typeof window !== "undefined") {
    window.openConfirmDialog = openConfirmDialog;
}
