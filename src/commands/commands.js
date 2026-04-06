/* global Office */

Office.onReady(() => {
  console.log("Commands loaded");
});

function openConfirmDialog(event) {
  Office.context.ui.displayDialogAsync(
    "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html",
    { height: 70, width: 60 },
    () => {}
  );

  event.completed();
}

window.openConfirmDialog = openConfirmDialog;
