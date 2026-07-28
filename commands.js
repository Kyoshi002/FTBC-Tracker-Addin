// "Track Request" ribbon button — Engineer Request Tracker (FTBC Procurement)
//
// This runs inside Office's sandboxed webview: it can only talk to the mail
// item via Office.js. It cannot run local Python or reach the tracker
// pipeline directly. Its ONE job is to apply the "→ Track" category to the
// open email — the same tag-based intake path request_tracker.py's scan()
// already reads (proven working via COM regardless of which Outlook client
// applied the tag). Processing happens on Yoshi's next Sync Now press, the
// hotkey (classic Outlook only), or the 8AM Morning Brief refresh.

const CATEGORY = "→ Track";

Office.onReady(() => {
  // no-op: Office.actions.associate below is what actually wires the button
});

function trackRequest(event) {
  const item = Office.context.mailbox.item;

  item.categories.addAsync([CATEGORY], (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      notify(item, "Tagged '→ Track' — processed on your next Sync Now or 8AM refresh.", "informationalMessage");
    } else {
      // Fail loud, not silent: this button's only job is the tag, so if it
      // didn't land, the user must know now, not discover it later.
      notify(item, "Could not tag this email: " + (result.error ? result.error.message : "unknown error"), "errorMessage");
    }
    event.completed();
  });
}

function notify(item, message, type) {
  item.notificationMessages.addAsync("trackRequestResult", {
    type: type,
    message: message,
    icon: "icon-16",
    persistent: false,
  });
}

Office.actions.associate("trackRequest", trackRequest);
