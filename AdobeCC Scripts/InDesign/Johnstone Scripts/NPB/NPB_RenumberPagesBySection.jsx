// Get the active document
var activeDoc = app.activeDocument;

// Begin undo group
app.doScript(function() {

    // Retain the default section (first section) and remove all other sections
    for (var sectionIndex = activeDoc.sections.length - 1; sectionIndex > 0; sectionIndex--) {
        activeDoc.sections[sectionIndex].remove();
    }

    // Loop through all the pages in the document
    for (var pageIndex = 0; pageIndex < activeDoc.pages.length; pageIndex++) {
        var currentPage = activeDoc.pages[pageIndex];

        // Check if the applied master page is "A-Conf" or "B-NonConf"
        if (
            currentPage.appliedMaster !== null &&
            (currentPage.appliedMaster.name === "A-Conf" || currentPage.appliedMaster.name === "B-NonConf")
        ) {
            // Add a new section starting at this page
            var newPageSection = activeDoc.sections.add(currentPage);

            // Set continue numbering to false to start a new section
            newPageSection.continueNumbering = false;

            // Restart the page numbering from 1
            newPageSection.pageNumberStart = 1;
        }
    }

}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Reset Page Numbering for A-Conf and B-NonConf");
