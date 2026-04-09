/* global Office */

Office.onReady(() => {
  // Commands are ready
});

// This function is called when the ribbon button is clicked (if using ExecuteFunction action type).
// Currently we use ShowTaskpane, so this is a placeholder for future function commands.
function analyzeEmailCommand(event) {
  event.completed();
}

// Register the function with Office
if (typeof Office !== "undefined") {
  Office.actions = Office.actions || {};
  Office.actions.associate("analyzeEmailCommand", analyzeEmailCommand);
}
