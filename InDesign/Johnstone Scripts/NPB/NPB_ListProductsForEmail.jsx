// npb_list_product_for_email.jsx
// By Victor Paredes for Johnstone Supply
// August 2023

// Begin an undo group
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Create New Page and Copy Product Line Content");

function main() {
    // Create a new blank page at the top of the document
    var doc = app.activeDocument;
    var firstPage = doc.pages.add(LocationOptions.AT_BEGINNING);

    // Create a new text frame on the new page
    var textFrame = firstPage.textFrames.add();
    textFrame.geometricBounds = ["12.7mm", "12.7mm", "200mm", "185mm"];

    // Loop through all objects in the document to find those with the script label "product_line"
    var allPageItems = doc.allPageItems;
    for (var i = 0; i < allPageItems.length; i++) {
        var item = allPageItems[i];
        if (item.label === "product_line") {
            try {
                var content = item.contents;
                textFrame.contents += content + "\n";
            } catch (e) {
                $.writeln("Error processing item: " + e);
            }
        }
    }

    // Remove unwanted text
    var unwantedText = "ProductGroup_NewProductBulletinNo - ProductGroup_BulletinHeader";
    app.findGrepPreferences.findWhat = unwantedText;
    var foundItems = textFrame.findGrep();
    for (var j = 0; j < foundItems.length; j++) {
        foundItems[j].remove();
    }
    app.findGrepPreferences = NothingEnum.NOTHING;

    // Get total line count and adjust
    var lineCount = textFrame.lines.length;
    textFrame.contents += "\nTotal products: " + (lineCount - 1);

    // Store content and clear the frame
    var originalContent = textFrame.contents;
    textFrame.contents = "";

    // Generate today's date as YYYY_M_D
    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth() + 1; // zero-based
    var day = today.getDate();
    var formattedDate = year + "_" + month + "_" + day;

    // Add the intro text with dynamic date
    var introText = "Hi Team,\n\nThe New Product Bulletin has been uploaded to the following drive:\nS:\\Product Specialists\\NP Bulletins 2024\\" + formattedDate + "\n\nThe following products are included:\n\n";
    textFrame.contents = introText + originalContent;

    // Fit to content
    textFrame.fit(FitOptions.FRAME_TO_CONTENT);
}
