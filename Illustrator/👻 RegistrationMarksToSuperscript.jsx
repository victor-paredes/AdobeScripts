// Loop through all text objects in the document and make "®" superscript
function makeRegisteredTrademarkSuperscript() {
    if (app.documents.length === 0) {
        alert("No document open.");
        return;
    }

    var doc = app.activeDocument;
    var textFrames = doc.textFrames;
    
    // Loop through all text frames
    for (var i = 0; i < textFrames.length; i++) {
        var textFrame = textFrames[i];
        
        // Loop through each character in the text frame
        for (var j = 0; j < textFrame.characters.length; j++) {
            var character = textFrame.characters[j];
            
            // Check if the character is the "®" symbol
            if (character.contents === "®") {
                // Set the superscript attribute
                character.baselineShift = 7; // Adjust as needed
                character.horizontalScale = 60; // Shrinks the symbol
                character.verticalScale = 60; // Adjust as needed
            }
        }
    }

    alert("All ® symbols have been made superscript.");
}

// Call the function
makeRegisteredTrademarkSuperscript();
