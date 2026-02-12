// Le dialog est isolé : il ne peut PAS accéder à Office.context.mailbox.item
// Il doit recevoir les données depuis commands.js via messageParent().

// ------------------------------
// 1) Initialisation Office.js
// ------------------------------
Office.onReady(() => {
    console.log("Dialog ready");
});

// ------------------------------
// 2) Réception des données envoyées par commands.js
// ------------------------------
Office.context.ui.addHandlerAsync(
    Office.EventType.DialogMessageReceived,
    (arg) => {
        try {
            const data = JSON.parse(arg.message);

            document.getElementById("sender").innerText  = data.sender || "—";
            document.getElementById("subject").innerText = data.subject || "—";
            document.getElementById("date").innerText    = data.date || "—";
        } catch (e) {
            console.error("Erreur parsing données dialog :", e);
        }
    }
);

// ------------------------------
// 3) Actions utilisateur (OUI / NON / AIDE)
// ------------------------------
document.addEventListener("DOMContentLoaded", () => {

    // Bouton OUI
    document.getElementById("btnYes").onclick = () => {
        Office.context.ui.messageParent("YES");
        Office.context.ui.closeContainer();
    };

    // Bouton NON
    document.getElementById("btnNo").onclick = () => {
        Office.context.ui.messageParent("NO");
        Office.context.ui.closeContainer();
    };

    // Bouton AIDE — ouvre support.html dans une nouvelle dialog
    document.querySelector(".help").onclick = () => {
        Office.context.ui.displayDialogAsync(
            "https://dellasiegaexternal2.github.io/Herisson/support.html",
            { height: 60, width: 50, displayInIframe: true }
        );
    };
});
