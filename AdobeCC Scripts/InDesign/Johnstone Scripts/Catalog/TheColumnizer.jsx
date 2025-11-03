var hardcodedExcludedColumnsDefault = [0, 3]; // Exclude 1st and 4th columns (0-based index)
var hardcodedExcludedColumns = hardcodedExcludedColumnsDefault.slice();
var doc = app.activeDocument;

app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "THE COLUMNIZER!");

function main() {
    if (app.selection.length === 0) {
        alert("Please select a table cell.");
        return;
    }

    var sel = app.selection[0];
    var cell = null;

    if (sel instanceof Cell) {
        cell = sel;
    } else if (sel.parent instanceof Cell) {
        cell = sel.parent;
    } else if (sel.parent.parent instanceof Cell) {
        cell = sel.parent.parent;
    }

    if (!cell) {
        alert("Selection is not inside a table cell.");
        return;
    }

    var table = cell.parent;
    var colIndex = cell.parentColumn.index;
    var totalCols = table.columns.length;

    var dialog = new Window("dialog", "THE COLUMNIZER!");
    dialog.add("statictext", undefined, "Resize table columns based on cell contents.");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.preferredSize.width = 500;

    var contentGroup = dialog.add("group");
    contentGroup.orientation = "row";
    contentGroup.alignChildren = ["fill", "top"];

    var leftCol = contentGroup.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "left";

    var radioGroup = leftCol.add("panel", undefined, "Scope");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";

    var currentOnly = radioGroup.add("radiobutton", undefined, "Apply to current column only");
    var allColumns = radioGroup.add("radiobutton", undefined, "Apply to all columns in the table");
    currentOnly.value = true;

    var rulesPanel = leftCol.add("panel", undefined, "Rules");
    rulesPanel.orientation = "column";
    rulesPanel.alignChildren = "left";

    var includeAllRows = rulesPanel.add("radiobutton", undefined, "Include all rows");
    var excludeFirstRow = rulesPanel.add("radiobutton", undefined, "Exclude first row");
    var onlyFirstRow = rulesPanel.add("radiobutton", undefined, "Only include first row");
    includeAllRows.value = true;

    var paddingPanel = leftCol.add("panel", undefined, "Padding");
    paddingPanel.orientation = "column";
    paddingPanel.alignChildren = "left";

    var pad07 = paddingPanel.add("radiobutton", undefined, "0.07 pt (minimum recommended)");
    var pad10 = paddingPanel.add("radiobutton", undefined, "0.10 pt (improved legibility)");
    pad07.value = true;

    var rightCol = contentGroup.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = "fill";
    rightCol.margins = [20, 0, 0, 0];

    rightCol.add("statictext", undefined, "Exclude the following columns:");

    var colLabels = [];
    var visibleColIndices = [];

    function buildColumnList() {
        colLabels.length = 0;
        visibleColIndices.length = 0;
        for (var i = 0; i < totalCols; i++) {
            var isExcluded = false;
            for (var j = 0; j < hardcodedExcludedColumns.length; j++) {
                if (hardcodedExcludedColumns[j] === i) {
                    isExcluded = true;
                    break;
                }
            }
            if (isExcluded) continue;

            var rawText = "";
            try {
                rawText = table.rows[0].cells[i].texts[0].contents;
            } catch (e) {
                rawText = "";
            }

            var cleaned = rawText.replace(/[\u00A0\r\n\t\s]+/g, "");
            if (cleaned !== "") {
                var label = rawText.replace(/^\s+|\s+$/g, "");
                colLabels.push((i + 1) + ": " + label);
                visibleColIndices.push(i);
            }
        }
    }

    buildColumnList();
    var columnList = rightCol.add("listbox", undefined, colLabels, { multiselect: true });
    columnList.preferredSize.height = 200;
    columnList.enabled = false;

    var toggleHardcodedCheckbox = rightCol.add("checkbox", undefined, "Catalog: Hide empty columns.");
    toggleHardcodedCheckbox.value = true;

    toggleHardcodedCheckbox.onClick = function () {
        hardcodedExcludedColumns = toggleHardcodedCheckbox.value ? hardcodedExcludedColumnsDefault.slice() : [];
        buildColumnList();
        columnList.removeAll();
        for (var i = 0; i < colLabels.length; i++) {
            columnList.add("item", colLabels[i]);
        }
    };

    var clearButton = rightCol.add("button", undefined, "Clear");
    clearButton.enabled = false;

    clearButton.onClick = function () {
        for (var i = 0; i < columnList.items.length; i++) {
            columnList.items[i].selected = false;
        }
        if (columnList.items.length > 0) {
            columnList.items[0].selected = true;
            columnList.items[0].selected = false;
        }
        dialog.active = true;
    };

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, "Apply", { name: "ok" });
    buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });

    var footer = dialog.add("statictext", undefined, "Designed by Victor Paredes | 2025");
    footer.alignment = "center";
    footer.graphics.font = ScriptUI.newFont("Arial", "italic", 10);

    currentOnly.onClick = function () {
        columnList.enabled = false;
        clearButton.enabled = false;
        for (var i = 0; i < columnList.items.length; i++) {
            columnList.items[i].selected = false;
        }
    };

    allColumns.onClick = function () {
        columnList.enabled = true;
        clearButton.enabled = true;
    };

    if (dialog.show() !== 1) return;

    var bufferAmount = pad10.value ? 0.10 : 0.07;
    var excludedIndices = [];

    if (allColumns.value) {
        for (var i = 0; i < columnList.items.length; i++) {
            if (columnList.items[i].selected) {
                excludedIndices.push(visibleColIndices[i]);
            }
        }
    }

    var useOnlyFirst = onlyFirstRow.value;
    var excludeFirst = excludeFirstRow.value;

    if (allColumns.value) {
        for (var v = 0; v < visibleColIndices.length; v++) {
            var colIndexReal = visibleColIndices[v];
            if (contains(excludedIndices, colIndexReal)) continue;
            if (contains(hardcodedExcludedColumns, colIndexReal)) continue;
            if (isEntireColumnEmpty(table, colIndexReal)) continue;

            var maxWidth = 0;
            for (var r = 0; r < table.bodyRowCount; r++) {
                if (excludeFirst && r === 0) continue;
                if (useOnlyFirst && r > 0) break;

                try {
                    var cell = table.rows[r].cells[colIndexReal];
                    if (cell.columnSpan > 1) continue;
                    var width = measureTextWidth(cell.texts[0]);
                    if (!isNaN(width) && width > maxWidth) maxWidth = width;
                } catch (e) {
                    continue;
                }
            }

            var finalWidth = maxWidth + bufferAmount;
            if (!isNaN(finalWidth) && finalWidth > 0) {
                table.columns[colIndexReal].width = finalWidth;
            }
        }
    } else {
        if (!contains(hardcodedExcludedColumns, colIndex)) {
            var maxWidth = 0;
            for (var r = 0; r < table.bodyRowCount; r++) {
                if (excludeFirst && r === 0) continue;
                if (useOnlyFirst && r > 0) break;

                try {
                    var cell = table.rows[r].cells[colIndex];
                    if (cell.columnSpan > 1) continue;
                    var width = measureTextWidth(cell.texts[0]);
                    if (!isNaN(width) && width > maxWidth) maxWidth = width;
                } catch (e) {
                    continue;
                }
            }

            var finalWidth = maxWidth + bufferAmount;
            if (!isNaN(finalWidth) && finalWidth > 0) {
                table.columns[colIndex].width = finalWidth;
            }
        }
    }
}

function measureTextWidth(text) {
    var tempFrame = doc.textFrames.add({
        geometricBounds: ["0p0", "0p0", "100p0", "100p0"]
    });
    tempFrame.contents = text.contents;
    tempFrame.texts[0].properties = text.properties;

    var justification = text.justification;
    var width;

    if (justification === Justification.RIGHT_ALIGN || justification === Justification.CENTER_ALIGN) {
        var start = tempFrame.texts[0].insertionPoints[0].horizontalOffset;
        var end = tempFrame.texts[0].insertionPoints[-1].horizontalOffset;
        width = Math.abs(end - start);
    } else {
        width = tempFrame.texts[0].endHorizontalOffset;
    }

    tempFrame.remove();
    return width;
}

function contains(array, value) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] === value) return true;
    }
    return false;
}

function isEntireColumnEmpty(table, colIndex) {
    for (var r = 0; r < table.bodyRowCount; r++) {
        try {
            var cell = table.rows[r].cells[colIndex];
            if (cell.columnSpan > 1) continue;
            var text = cell.texts[0].contents;
            if (text.replace(/[\u00A0\r\n\t\s]+/g, "") !== "") return false;
        } catch (e) {}
    }
    return true;
}
