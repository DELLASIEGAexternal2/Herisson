/****************************************************
 *  MODE SIMULATEUR (UNIQUEMENT admin-simulator1)
 ****************************************************/
if (location.href.includes("admin-simulator1")) {
    console.warn("MODE SIMULATEUR ACTIVÉ – admin-simulator1 – Office.js désactivé.");

    document.addEventListener("DOMContentLoaded", () => {

        // Données figées pour la capture
        document.getElementById("sender").innerText  = "Microsoft au nom de...";
        document.getElementById("subject").innerText = "Vous avez des tâches en retard";
        document.getElementById("date").innerText    = "16/01/2026 08:40";

        // Boutons simulateur
        document.getElementById("btnYes").onclick = () => alert("YES (simulateur)");
        document.getElementById("btnNo").onclick  = () => alert("NO (simulateur)");

        document.querySelector(".help").onclick = () =>
            window.open("https://dellasiegaexternal2.github.io/Herisson/support.html");
    });

    // Stopper le mode Outlook
    throw new Error("Simulateur admin-simulator1 – Office désactivé.");
}


/****************************************************
 *  MODE NAVIGATEUR (GitHub Pages / Chrome)
 ****************************************************/
if (typeof Office === "undefined" || !Office.context) {
    console.warn("MODE NAVIGATEUR – Preview confirm.html");

    document.addEventListener("DOMContentLoaded", () => {
        document.getElementById("sender").innerText  = "(hors Outlook)";
        document.getElementById("subject").innerText = "(hors Outlook)";
        document.getElementById("date").innerText    = new Date().toLocaleString();

        document.getElementById("btnYes").onclick = () => alert("YES (preview)");
        document.getElementById("btnNo").onclick  = () => alert("NO (preview)");

        document.querySelector(".help").onclick = () =>
            window.open("https://dellasiegaexternal2.github.io/Herisson/support.html");
    });

    throw new Error("Confirm.js exécuté hors Outlook – Preview.");
}


/****************************************************
 *  MODE OUTLOOK (dialog réelle)
 ****************************************************/

Office.onReady(() => {
    console.log("Dialog ready");
});

// Réception des données envoyées par commands.js
Office.context.ui.addHandlerAsync(
    Office.EventType.DialogMessageReceived,
    (arg) => {
        if (!arg || !arg.message) {
            console.warn("DialogMessageReceived sans message");
            return;
        }

        try {
            const data = JSON.parse(arg.message);

            document.getElementById("sender").innerText  = data.sender || "—";
            document.getElementById("subject").innerText = data.subject || "—";
            document.getElementById("date").innerText    = data.date || "—";

        } catch (e) {
            console.error("JSON dialog invalide :", arg.message, e);
        }
    }
);


// Actions utilisateur Outlook
document.addEventListener("DOMContentLoaded", () => {

    document.getElementById("btnYes").onclick = () => {
        Office.context.ui.messageParent("YES");
        Office.context.ui.closeContainer();
    };

    document.getElementById("btnNo").onclick = () => {
        Office.context.ui.messageParent("NO");
        Office.context.ui.closeContainer();
    };

    document.querySelector(".help").onclick = () => {
        Office.context.ui.displayDialogAsync(
            "https://dellasiegaexternal2.github.io/Herisson/support.html",
            { height: 60, width: 50, displayInIframe: true }
        );
    };
});
