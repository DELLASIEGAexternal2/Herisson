/* global Office */
Office.onReady(() => {});

/**
 * Ouvre la pop-up et transmet le contexte du message courant.
 * Appelée par la commande du ruban (ExecuteFunction).
 */
function openConfirmDialog(event) {
  const url = "https://dellasiegaexternal2.github.io/Herisson/src/confirm.html";

  // 1) Préparer le contexte du message
  const item = Office.context.mailbox.item;
  const payload = {
    sender:
      (item.from && (item.from.displayName || item.from.emailAddress)) ||
      (item.sender && (item.sender.displayName || item.sender.emailAddress)) ||
      "",
    subject: item.subject || "",
    date:
      (item.dateTimeCreated && new Date(item.dateTimeCreated).toLocaleString()) ||
      (item.dateTimeReceived && new Date(item.dateTimeReceived).toLocaleString()) ||
      ""
  };

  // 2) Ouvrir la dialog
  Office.context.ui.displayDialogAsync(
    url,
    { height: 60, width: 50, requireHTTPS: true, displayInIframe: true },
    (res) => {
      // Toujours compléter la commande du ruban
      try { event.completed(); } catch (_) {}

      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("displayDialogAsync error:", res.error);
        return;
      }

      const dialog = res.value;

      // 3) Envoyer le contexte à la dialog
      try {
        dialog.messageChild(JSON.stringify({ type: "init", data: payload }));
      } catch (e) {
        // Certains hôtes nécessitent un léger délai avant messageChild
        setTimeout(() => {
          try { dialog.messageChild(JSON.stringify({ type: "init", data: payload })); } catch {}
        }, 300);
      }

      // 4) Ecouter la réponse de la dialog (YES / NO)
      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        (arg) => {
          // arg.message est une string
          try {
            const msg = JSON.parse(arg.message);
            if (msg && msg.type === "result") {
              console.log("Dialog result:", msg.value); // "YES" ou "NO"
            }
          } catch {
            console.log("Dialog raw message:", arg.message);
          }
          dialog.close();
        }
      );

      // 5) Gestion fermeture / erreur de la dialog
      dialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        (evt) => {
          console.warn("DialogEventReceived:", evt);
        }
      );
    }
  );
}

// Associer la fonction si tu utilises Office.actions (facultatif avec ExecuteFunction)
if (Office.actions && Office.actions.associate) {
  Office.actions.associate("openConfirmDialog", openConfirmDialog);
}

// Exposer pour <FunctionName>
if (typeof window !== "undefined") {
  window.openConfirmDialog = openConfirmDialog;
}
