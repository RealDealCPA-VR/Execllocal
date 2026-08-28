/*
 * Office command functions (FunctionFile). The manifest currently only uses
 * the ShowTaskpane action; this file stays as the extension point for future
 * UI-less commands. Keep it host-agnostic: no Outlook-only APIs here.
 */

/* global Office */

Office.onReady(() => {
  // Office.js is ready to be called from command functions.
});

/**
 * Placeholder command. Wire it to a manifest Control action when needed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  // Nothing to do yet; always signal completion so Office never hangs.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
