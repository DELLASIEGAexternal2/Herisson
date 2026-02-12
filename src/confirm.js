/****************************************************
 *  MODE SIMULATEUR / NAVIGATEUR (pas Outlook)
 *  Permet de figer la fenêtre dans le simulateur M365
 ****************************************************/
if (window.location.href.includes("admin-simulator")) {
    console.warn("MODE SIMULATEUR ACTIVÉ – Office.js désactivé.");

    document.addEventListener("DOMContentLoaded", () => {
        // Données factices
        document.getElementById("sender").innerText  = "exemple@domain.fr";
        document.getElementById("subject").innerText = "Message de démonstration";
        document.getElementById("date").innerText    = "12/02/2026 10:45";

        // Boutons (sans Office)
        document.getElementById("btnYes").onclick = () => alert("YES (simulateur)");
        document.getElementById("btnNo").onclick  = () => alert("NO (simulateur)");
        document.querySelector(".help").onclick = () =>
            window.open("https://dellasiegaexternal2.github.io/Herisson/support.html");
    });

    // STOP ici → on ne charge PAS Office.js
    throw new Error("Simulateur détecté – Office.onReady désactivé.");
}


/****************************************************
 *  MODE NAVIGATEUR HORS SIMULATEUR (GitHub, Chrome)
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

// Réception des données depuis commands.js
Office.context.ui.addHandlerAsync(
    Office.EventType.DialogMessageReceived,
    (arg) => {
        if (!arg || !arg.message) return;

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

// Actions utilisateur
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
