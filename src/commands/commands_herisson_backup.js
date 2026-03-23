function openConfirmDialog(event) {
    const url = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

    Office.context.ui.displayDialogAsync(
        url,
        { height: 60, width: 50 },
        () => event.completed()
    );
}

Office.actions.associate("openConfirmDialog", openConfirmDialog);