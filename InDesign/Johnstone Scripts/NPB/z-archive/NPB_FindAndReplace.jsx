if (app.documents.length === 0) {
    alert("Please open a document before running this script.");
} else {
    var doc = app.activeDocument;

    var options = [
        {
            id: "cleanSupTag",
            labelBase: "Replace <sup>®</sup> with ® (no formatting)",
            findWhat: "<sup>^r</sup>",
            changeTo: "^r",
            format: null
        },
        {
            id: "rSuperscript",
            labelBase: "Replace ® with ® (superscript)",
            findWhat: "^r",
            changeTo: "^r",
            format: Position.SUPERSCRIPT
        },
        {
            id: "tmSuperscript",
            labelBase: "Replace ™ with ™ (superscript)",
            findWhat: "^d",
            changeTo: "^d",
            format: Position.SUPERSCRIPT
        },
        {
            id: "disableHyphenation",
            labelBase: "Disable hyphenation in all text frames",
            findWhat: null,
            changeTo: null,
            format: null,
            isCustomHandler: true,
            handler: function () {
                var stories = doc.stories;
                for (var s = 0; s < stories.length; s++) {
                    var paras = stories[s].paragraphs;
                    for (var p = 0; p < paras.length; p++) {
                        paras[p].hyphenation = false;
                    }
                }
            }
        }
    ];

    // Count matches and build labels
    for (var i = 0; i < options.length; i++) {
        if (options[i].findWhat) {
            app.findTextPreferences = NothingEnum.nothing;
            app.findTextPreferences.findWhat = options[i].findWhat;
            var results = doc.findText();
            options[i].count = results.length;
            options[i].label = options[i].labelBase;
            options[i].countText = "(" + results.length + ")";
        } else {
            options[i].count = 0;
            options[i].label = options[i].labelBase;
            options[i].countText = "";
        }
    }
    app.findTextPreferences = NothingEnum.nothing;

    var dlg = new Window("dialog", "Choose and Order Find/Change Options");
    dlg.alignChildren = "fill";

    var checkboxGroup = dlg.add("group");
    checkboxGroup.orientation = "column";
    checkboxGroup.alignChildren = "left";

    var checkboxes = {};
    var labels = {};
    var listbox = null;

    function updateListNumbers() {
        for (var i = 0; i < listbox.items.length; i++) {
            var plainLabel = listbox.items[i].text.replace(/^\d+\)\s*/, "");
            listbox.items[i].text = (i + 1) + ") " + plainLabel;
        }
    }

    function addItem(labelBase) {
        for (var i = 0; i < listbox.items.length; i++) {
            if (listbox.items[i].text.indexOf(labelBase) > -1) return;
        }
        listbox.add("item", labelBase);
        updateListNumbers();
    }

    function removeItem(labelBase) {
        for (var i = listbox.items.length - 1; i >= 0; i--) {
            if (listbox.items[i].text.indexOf(labelBase) > -1) {
                listbox.remove(i);
            }
        }
        updateListNumbers();
    }

    for (var i = 0; i < options.length; i++) {
        var row = checkboxGroup.add("group");
        row.orientation = "row";

        var cb = row.add("checkbox", undefined, options[i].label);
        checkboxes[options[i].id] = cb;

        var countLabel = row.add("statictext", undefined, options[i].countText);
        countLabel.graphics.foregroundColor = countLabel.graphics.newPen(
            countLabel.graphics.PenType.SOLID_COLOR,
            [0.5, 0.5, 0.5],
            1
        );

        labels[options[i].id] = countLabel;

        cb.value = true;
    }

    var listGroup = dlg.add("group");
    listGroup.orientation = "column";
    listGroup.alignChildren = "fill";
    listGroup.add("statictext", undefined, "Execution Order:");

    listbox = listGroup.add("listbox", undefined, [], { multiselect: false });
    listbox.preferredSize = [400, 100];

    var buttonRow = listGroup.add("group");
    var upBtn = buttonRow.add("button", undefined, "↑");
    var downBtn = buttonRow.add("button", undefined, "↓");

    var footer = dlg.add("group");
    footer.alignment = "right";
    var cancelBtn = footer.add("button", undefined, "Cancel", { name: "cancel" });
    var okBtn = footer.add("button", undefined, "Update", { name: "ok" });

    // Populate listbox on load
    for (var i = 0; i < options.length; i++) {
        if (checkboxes[options[i].id].value) {
            addItem(options[i].label);
        }
    }

    for (var i = 0; i < options.length; i++) {
        (function (opt) {
            checkboxes[opt.id].onClick = function () {
                if (this.value) {
                    addItem(opt.label);
                } else {
                    removeItem(opt.label);
                }
            };
        })(options[i]);
    }

    upBtn.onClick = function () {
        var sel = listbox.selection;
        if (sel && sel.index > 0) {
            var i = sel.index;
            var temp = listbox.items[i - 1].text;
            listbox.items[i - 1].text = listbox.items[i].text;
            listbox.items[i].text = temp;
            listbox.selection = i - 1;
            updateListNumbers();
        }
    };

    downBtn.onClick = function () {
        var sel = listbox.selection;
        if (sel && sel.index < listbox.items.length - 1) {
            var i = sel.index;
            var temp = listbox.items[i + 1].text;
            listbox.items[i + 1].text = listbox.items[i].text;
            listbox.items[i].text = temp;
            listbox.selection = i + 1;
            updateListNumbers();
        }
    };

    if (dlg.show() === 1) {
        app.doScript(function () {
            for (var i = 0; i < listbox.items.length; i++) {
                var plainLabel = listbox.items[i].text.replace(/^\d+\)\s*/, "");
                for (var j = 0; j < options.length; j++) {
                    if (options[j].label === plainLabel) {
                        if (options[j].isCustomHandler) {
                            options[j].handler();
                        } else {
                            app.findTextPreferences = NothingEnum.nothing;
                            app.changeTextPreferences = NothingEnum.nothing;
                            app.findTextPreferences.findWhat = options[j].findWhat;
                            app.changeTextPreferences.changeTo = options[j].changeTo;
                            if (options[j].format !== null) {
                                app.changeTextPreferences.position = options[j].format;
                            }
                            doc.changeText();
                        }
                        break;
                    }
                }
            }

            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Run Ordered Find/Change");
    }
}
