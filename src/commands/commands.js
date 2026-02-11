function openConfirmDialog(event) {
    Office.context.ui.displayDialogAsync(
        window.location.origin + "/confirm.html",
        { height: 60, width: 50 },
        () => event.completed()
    );
}

Office.actions.associate("openConfirmDialog", openConfirmDialog);
