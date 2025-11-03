
/*
===============================================================================
 Script: Trim Overset Text
 Description: Removes overset text from all text frames in the active InDesign
              document, including nested and story-linked frames. Displays a
              modal indicating how many frames were updated.
 
 Written: July 2025
 Author: Victor Paredes
 For: Johnstone Supply
===============================================================================
*/


function removeOversetTextProperly() {
    var doc = app.activeDocument;
    var updatedCount = 0;

    function trimOverset(tf) {
        if (!tf.overflows) return;

        try {
            var story = tf.parentStory;
            var fullText = story.texts[0];
            var lastVisibleInsertion = tf.insertionPoints[-1];
            var cutIndex = lastVisibleInsertion.index;

            if (cutIndex < fullText.characters.length) {
                var startChar = fullText.characters[cutIndex];
                var endChar = fullText.characters[-1];
                fullText.characters.itemByRange(startChar, endChar).remove();
                updatedCount++;
            }
        } catch (e) {
            $.writeln("Error processing frame: " + e);
        }
    }

    function processAllTextFrames(container) {
        for (var i = 0; i < container.textFrames.length; i++) {
            trimOverset(container.textFrames[i]);
        }

        if ("allPageItems" in container) {
            var items = container.allPageItems;
            for (var j = 0; j < items.length; j++) {
                var item = items[j];
                if (item instanceof TextFrame) {
                    processAllTextFrames(item);
                }
            }
        }
    }

    for (var p = 0; p < doc.pages.length; p++) {
        processAllTextFrames(doc.pages[p]);
    }

    for (var m = 0; m < doc.masterSpreads.length; m++) {
        processAllTextFrames(doc.masterSpreads[m]);
    }

    for (var s = 0; s < doc.stories.length; s++) {
        var story = doc.stories[s];
        for (var t = 0; t < story.textContainers.length; t++) {
            var tf = story.textContainers[t];
            if (tf instanceof TextFrame) {
                trimOverset(tf);
            }
        }
    }

    return updatedCount;
}

function showResultModal(count) {
    var w = new Window("dialog", "Overset Text Cleanup");
    w.alignChildren = "center";

    var message = count === 1
        ? "1 text frame was updated to remove overset text."
        : count + " text frames were updated to remove overset text.";
    w.add("statictext", undefined, message);

    var footer = w.add("statictext", undefined, "July 2025 | vp");
    footer.graphics.font = ScriptUI.newFont("Helvetica", "italic", 9);
    footer.alignment = "center";

    var btn = w.add("button", undefined, "OK", {name: "ok"});
    btn.alignment = "center";

    w.show();
}

if (app.documents.length === 0) {
    alert("No document is open.");
} else {
    var result = app.doScript(removeOversetTextProperly, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Trim Overset Text");
    showResultModal(result);
}
