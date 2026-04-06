/* global Office */

Office.onReady(() => {
  console.log("Herisson Pret");
});

function openConfirmDialog(event) {
  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      console.error("Pas de mail");
      event.completed();
      return;
    }

    const dialogUrl = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

    const mailData = {
      subject: item.subject || "N/A",
      sender: item.from?.emailAddress || "N/A",
      senderName: item.from?.displayName || "N/A",
      date: item.dateTimeCreated || new Date()
    };

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 70, width: 60 },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(asyncResult.error);
          return;
        }

        const dialog = asyncResult.value;

        setTimeout(() => {
          dialog.messageChild(JSON.stringify(mailData));
        }, 500);
      }
    );

  } catch (e) {
    console.error("Dialog error:", e);
  }

  event.completed();
}

window.openConfirmDialog = openConfirmDialog;
