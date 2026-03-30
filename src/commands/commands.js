/* global Office */

/**
 * Initialisation Office
 */
Office.onReady(() => {
    console.log("Commands ready");
});

/**
 * Fonction appelée depuis le manifest
 * → ouvre la popup + envoie les infos du mail
 */
function openConfirmDialog(event) {

    try {

        const url = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

        const item = Office.context.mailbox.item;

        // Sécurisation (évite crash Outlook)
        if (!item) {
            console.error("Aucun mail sélectionné");
            event.completed();
            return;
        }

        const mailData = {
            sender: item.from?.displayName || item.from?.emailAddress || "(expéditeur inconnu)",
            subject: item.subject || "(sans sujet)",
            date: item.dateTimeCreated
                ? new Date(item.dateTimeCreated).toLocaleString()
                : "(date inconnue)"
        };

        // Ouverture popup
        Office.context.ui.displayDialogAsync(
            url,
            {
                height: 70,
                width: 60,
                displayInIframe: true
            },
            function (asyncResult) {

                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error("Erreur ouverture popup :", asyncResult.error);
                    event.completed();
                    return;
                }

                const dialog = asyncResult.value;

                // Envoi des données vers la popup
                setTimeout(() => {
                    dialog.messageChild(JSON.stringify(mailData));
                }, 300);

                // Réception réponse utilisateur
                dialog.addEventHandler(
                    Office.EventType.DialogMessageReceived,
                    function (arg) {

                        try {
                            const msg = JSON.parse(arg.message);

                            if (msg.confirm === true) {

                                console.log("✔ Confirmation utilisateur");

                                // 🔥 ICI TU AJOUTERAS :
                                // → envoi mail CERT
                                // → Graph API
                                // → pièces jointes

                            } else {
                                console.log("❌ Annulé utilisateur");
                            }

                        } catch (e) {
                            console.error("Erreur parsing réponse popup", e);
                        }

                        dialog.close();
                    }
                );
            }
        );

    } catch (error) {
        console.error("Erreur globale :", error);
    }

    // Obligatoire pour Outlook
    event.completed();
}

/**
 * Exposition globale (OBLIGATOIRE pour manifest)
 */
window.openConfirmDialog = openConfirmDialog;
