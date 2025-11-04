#target indesign

(function() {
    var doc = app.activeDocument;
    var pages = doc.pages;
    var totalPages = pages.length;

    // Confirm user wants to proceed
    if (!confirm("This will interleave pages based on master page assignments.\nAll content will be preserved.\n\nProceed?")) {
        return;
    }

    // Group pages by master
    var masterAGroup = [];
    var masterBGroup = [];

    for (var i = 0; i < totalPages; i++) {
        var page = pages[i];
        var appliedMaster = page.appliedMaster;

        if (!appliedMaster) continue;

        if (masterAGroup.length === 0) {
            masterAGroup.push(page);
        } else if (appliedMaster == masterAGroup[0].appliedMaster) {
            masterAGroup.push(page);
        } else {
            masterBGroup.push(page);
        }
    }

    // Sanity check
    if (masterAGroup.length !== masterBGroup.length) {
        alert("Error: The two groups do not have an equal number of pages.\nPlease ensure the document has an even number of pages split evenly between two master types.");
        return;
    }

    app.doScript(function() {
        var newPages = [];
        var pageCount = masterAGroup.length;

        // Insert new interleaved pages at the end
        for (var i = 0; i < pageCount; i++) {
            var pageA = masterAGroup[i];
            var pageB = masterBGroup[i];

            // Duplicate each to end of document
            var newPageA = pageA.duplicate(LocationOptions.AT_END);
            var newPageB = pageB.duplicate(LocationOptions.AT_END);

            newPages.push(newPageA);
            newPages.push(newPageB);
        }

        // Delete original pages
        // Important: Must delete from the end to avoid index shifting
        for (var i = totalPages - 1; i >= 0; i--) {
            pages[i].remove();
        }

    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Interleave Pages by Master");

})();
