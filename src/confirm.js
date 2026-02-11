/*document.getElementById("btnYes").onclick = () => {
    alert("Confirmation UI uniquement (pas encore de traitement)");
};

document.getElementById("btnNo").onclick = () => {
    window.close();
};

document.querySelector(".help").onclick = () => {
    alert("Aide MailSuspect (UI uniquement)");
}; */

Office.onReady(() => {
    const item = Office.context.mailbox.item;

    document.getElementById("sender").innerText =
        item.from ? item.from.emailAddress : "—";

    document.getElementById("subject").innerText =
        item.subject || "—";

    document.getElementById("date").innerText =
        item.dateTimeCreated
            ? new Date(item.dateTimeCreated).toLocaleString()
            : "—";
});

document.getElementById("btnYes").onclick = () => {
    alert("Signalement confirmé (UI + Outlook OK)");
    Office.context.ui.closeContainer();
};

document.getElementById("btnNo").onclick = () => {
    Office.context.ui.closeContainer();
};

document.querySelector(".help").onclick = () => {
    window.open("https://intranet/aide-mailsuspect", "_blank");
};

