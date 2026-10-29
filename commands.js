/**
 * Office Add-in background command file
 */
Office.onReady();

/**
 * Example background action command (if using UI-less button)
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // Be sure to call event.completed() when done
  event.completed();
}

// Register function in global scope
Office.actions.associate("action", action);
