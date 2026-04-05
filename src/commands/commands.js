/* global Office */

Office.onReady(() => {
    console.log("Commands ready");
});

function openConfirmDialog(event) {

    const url = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

    const item = Office.context.mailbox.item;

    if (!item) {
        console.error("Aucun mail sélectionné");
        event.completed();
        return;
    }

    const mailData = {
        sender: item.from?.displayName || item.from?.emailAddress || "-",
        subject: item.subject || "-",
        date: item.dateTimeCreated
            ? new Date(item.dateTimeCreated).toLocaleString()
            : "-"
    };

    Office.context.ui.displayDialogAsync(
        url,
        { height: 70, width: 60, displayInIframe: true },
        function (asyncResult) {

            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Erreur ouverture popup");
                event.completed();
                return;
            }

            const dialog = asyncResult.value;

            setTimeout(() => {
                dialog.messageChild(JSON.stringify(mailData));
            }, 300);

            dialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                function (arg) {

                    console.log("Retour dialog:", arg.message);

                    dialog.close();
                }
            );
        }
    );

    event.completed();
}

window.openConfirmDialog = openConfirmDialog;
