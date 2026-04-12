/* global Office */

Office.onReady(() => {
  console.log("Herisson Pret");
});

function openConfirmDialog(event) {

  try {

    const item = Office.context.mailbox.item;

    if (!item || !item.subject) {
      console.error("Mail non disponible");
      event.completed();
      return;
    }

    const dialogUrl = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

    const mailData = {
      subject: item.subject,
      sender: item.from?.emailAddress,
      senderName: item.from?.displayName,
      date: item.dateTimeCreated,

      // 🔥 CRITIQUE
      itemId: Office.context.mailbox.convertToRestId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      )
    };

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      {
        height: 70,
        width: 60,
        displayInIframe: true
      },
      (asyncResult) => {

        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Dialog error:", asyncResult.error);
          return;
        }

        const dialog = asyncResult.value;

        // 🔥 sécurisation envoi data
        setTimeout(() => {
          try {
            dialog.messageChild(JSON.stringify(mailData));
            console.log("DATA SENT TO DIALOG");
          } catch (err) {
            console.error("MessageChild error:", err);
          }
        }, 800);
      }
    );

  } catch (e) {
    console.error("GLOBAL ERROR:", e);
  }

  event.completed();
}

// 🔥 CRITIQUE (OBLIGATOIRE)
Office.actions.associate("openConfirmDialog", openConfirmDialog);
