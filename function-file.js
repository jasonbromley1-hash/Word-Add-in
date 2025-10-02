/* global Office */

Office.onReady(() => {
    // This file is for ExecuteFunction commands
    // Currently not used but required by manifest
    console.log('Function file loaded');
});

// Example function command (for future use)
function insertQuickClause(event) {
    // Quick insert functionality could go here
    event.completed();
}

// Register functions
if (typeof Office !== 'undefined') {
    Office.actions.associate("insertQuickClause", insertQuickClause);
}
