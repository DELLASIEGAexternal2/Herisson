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
      date: item.dateTimeCreated
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
          console.error(asyncResult.error);
          return;
        }

        const dialog = asyncResult.value;

        setTimeout(() => {
          dialog.messageChild(JSON.stringify(mailData));
        }, 800);
      }
    );

  } catch (e) {
    console.error(e);
  }

  event.completed();
}

// 🔥 CRITIQUE
Office.actions.associate("openConfirmDialog", openConfirmDialog);
