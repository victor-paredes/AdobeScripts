/*
*
*   Replace Content By Key
*   v1.1 | May 2024
*   Written by Victor Paredes
*   https://victorpared.es
*
*   Description:
*   This Adobe Illustrator script will replace either text or image content based on key_ and object name pairings.
*   Example: "key_object_name" will replace content for objects named "object_name".
*
*   Special shoutout to Joonas Pääkkö (https://github.com/joonaspaakko) for his ScriptUI developer tool. 
*
*/







// Get the active document
var doc = app.activeDocument;

// Function to select all objects with the given name
function selectObjectsByName(name) {
    // Deselect everything
    doc.selection = null;
    
    // Iterate through all page items in the document
    var count = 0;
    for (var i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];

        // If the item's name matches the given name, add it to the selection
        if (item.name === name) {
            item.selected = true;
            count++;
        }
    }
    
    if (count > 0) {
        alert("Selected " + count + " object(s) with the name '" + name + "'.");
    } else {
        alert("No objects found with the name '" + name + "'.");
    }
}

// Check if an object is selected
if (doc.selection.length === 1) {
    var selectedItem = doc.selection[0];  // Get the selected object

    // Check if the selected item has a name
    if (selectedItem.name !== "") {
        var selectedItemName = selectedItem.name;
        selectObjectsByName(selectedItemName);
    } else {
        alert("The selected object does not have a name.");
    }
} else if (doc.selection.length === 0) {
    // No object is selected, prompt the user for a name using ScriptUI dialog
    var dialog = new Window('dialog', 'Select Objects by Name');
    
    // Add text and input field
    dialog.add('statictext', undefined, 'Enter the name of the objects to select:');
    var nameInput = dialog.add('edittext', undefined, '');
    nameInput.characters = 20;  // Set input width

    // Add OK and Cancel buttons
    var okButton = dialog.add('button', undefined, 'OK', {name: 'ok'});
    var cancelButton = dialog.add('button', undefined, 'Cancel', {name: 'cancel'});

    // Add button functionality
    okButton.onClick = function() {
        var enteredName = nameInput.text;
        dialog.close();  // Close the dialog
        selectObjectsByName(enteredName);  // Select objects by the entered name
    };

    cancelButton.onClick = function() {
        dialog.close();  // Close the dialog without doing anything
    };

    // Show the dialog
    dialog.show();
} else {
    alert("Please select exactly one object or none.");
}
