/* global Office */

Office.onReady(() => {
  console.log("Herisson ready");
});

function openConfirmDialog(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
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
    { height: 70, width: 60 },
    (asyncResult) => {
      const dialog = asyncResult.value;
      setTimeout(() => {
        dialog.messageChild(JSON.stringify(mailData));
      }, 300);
    }
  );

  event.completed();
}

window.openConfirmDialog = openConfirmDialog;
