/* global Office */
Office.onReady(() => {});

/* ***** AJOUT : FONCTION QUI OUVRE LA POPUP ***** */
function openConfirmDialog(event) {

  const url = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

  Office.context.ui.displayDialogAsync(
    url,
    {
    height: 60,
      width: 50,
      requireHTTPS: true,
      displayInIframe: true
    },
    (result) => {

      if (result.status === Office.AsyncResultStatus.Succeeded) {

        const dialog = result.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg) => {
            // ***** AJOUT : réception du message *****
            dialog.close();
          }
        );
      }

      event.completed();
    }
  );
}

/* ***** AJOUT : EXPOSER LA FONCTION ***** */
if (typeof window !== "undefined") {
  window.openConfirmDialog = openConfirmDialog;
}