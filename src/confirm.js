/****************************************************
 *  MODE NAVIGATEUR (PAS DANS OUTLOOK)
 *  Empêche les erreurs "Office is undefined"
 *  et permet un aperçu local de confirm.html
 ****************************************************/
if (typeof Office === "undefined" || !Office.context) {
    console.warn("confirm.js chargé HORS Outlook – mode navigateur activé.");

    document.addEventListener("DOMContentLoaded", () => {
        // Remplir avec des valeurs fictives
        document.getElementById("sender").innerText  = "(hors Outlook)";
        document.getElementById("subject").innerText = "(hors Outlook)";
        document.getElementById("date").innerText    = new Date().toLocaleString();

        // Bouton OUI
        document.getElementById("btnYes").onclick = () => {
            alert("YES (mode navigateur)");
        };

        // Bouton NON
        document.getElementById("btnNo").onclick = () => {
            alert("NO (mode navigateur)");
        };

        // Aide → ouvre support.html même hors Outlook
        document.querySelector(".help").onclick = () => {
            window.open(
                "https://dellasiegaexternal2.github.io/Herisson/support.html",
                "_blank"
            );
        };
    });

    // STOP ici → ne pas continuer l'exécution Outlook
    throw new Error("Confirm.js exécuté hors Outlook – Office context absent.");
}


/****************************************************
 *    MODE OUTLOOK (dialog displayDialogAsync)
 ****************************************************/

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

        // Outlook peut envoyer un 1er event sans data → ignorer
        if (!arg || !arg.message) {
            console.warn("DialogMessageReceived sans message.");
            return;
        }

        let data;
        try {
            data = JSON.parse(arg.message);
        } catch (e) {
            console.error("Message JSON invalide :", arg.message, e);
            return;
        }

        // Mise à jour UI
        document.getElementById("sender").innerText  = data.sender || "—";
        document.getElementById("subject").innerText = data.subject || "—";
        document.getElementById("date").innerText    = data.date || "—";
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

    // Bouton AIDE : ouverture du support
    document.querySelector(".help").onclick = () => {
        Office.context.ui.displayDialogAsync(
            "https://dellasiegaexternal2.github.io/Herisson/support.html",
            { height: 60, width: 50, displayInIframe: true }
        );
    };
});
