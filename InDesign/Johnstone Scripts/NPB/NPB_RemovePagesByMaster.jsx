// Create a ScriptUI dialog
var dialog = new Window("dialog", "Which bulletin set would you like to remove?");

// Add buttons for 'Confidential Pages' and 'Non-Confidential Pages'
var confButton = dialog.add("button", undefined, "Confidential Pages");
var nonConfButton = dialog.add("button", undefined, "Non-Confidential Pages");

// Add a close button
var cancelButton = dialog.add("button", undefined, "Cancel", {name: "cancel"});

// Variable to store the user's choice
var choice = "";

// Add button actions
confButton.onClick = function() {
    choice = "confidential";
    dialog.close();
};

nonConfButton.onClick = function() {
    choice = "non-confidential";
    dialog.close();
};

// Show the dialog
dialog.show();

// Function to remove confidential pages
function removeConfPages() {
    var doc = app.activeDocument;
    for (var i = doc.pages.length - 1; i >= 0; i--) {
        var page = doc.pages[i];
        if (page.appliedMaster != null && 
            (page.appliedMaster.name == 'A-Conf' || page.appliedMaster.name == 'C-Conf-Blank')) {
            page.remove();
        }
    }
    alert("Pages with master 'A-Conf' or 'C-Conf-Blank' have been removed.");
}

// Function to remove non-confidential pages
function removeNonConfPages() {
    var doc = app.activeDocument;
    for (var i = doc.pages.length - 1; i >= 0; i--) {
        var page = doc.pages[i];
        if (page.appliedMaster != null && 
            (page.appliedMaster.name == 'B-NonConf' || page.appliedMaster.name == 'D-NonConf_Blank')) {
            page.remove();
        }
    }
    alert("Pages with master 'B-NonConf' or 'D-NonConf_Blank' have been removed.");
}

// Execute the appropriate function based on the user's choice
if (choice === "confidential") {
    removeConfPages();
} else if (choice === "non-confidential") {
    removeNonConfPages();
}
