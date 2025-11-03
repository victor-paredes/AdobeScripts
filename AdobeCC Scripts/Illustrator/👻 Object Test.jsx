/*
  Ghostbyte ObjectCreator.jsx
  Illustrator ExtendScript + ScriptUI
  - All measurement fields have unit dropdowns (auto / pixels / inches)
  - Behaves like Illustrator Character panel
  - Create Front View button outside tabs
*/

#target illustrator
#targetengine "ghostbyte_objectcreator"

// === Global Data Variables ===
var csvData = [];
var currentPage = 1;
var rowsPerPage = 20;
var featureColumns = []; // Dynamic list of feature column names
var baseColumns = ["Position", "Product", "SKU", "Price"]; // Fixed columns before Features

// === Default Preferences ===
var PREFS = {
    // Front View
    frontWidth: "360",
    frontWidthUnit: 0, // 0=pixels, 1=inches
    frontBgElements: 0, // 0=Include, 1=None
    frontPosLabels: 0, // 0=Include, 1=None
    frontPegLineSpacing: "40",
    frontPegLineSpacingUnit: 0, // 0=pixels, 1=inches
    frontShelfHeight: "20",
    frontShelfHeightUnit: 0, // 0=pixels, 1=inches
    frontProdRectSpacing: "10",
    frontProdRectSpacingUnit: 0, // 0=pixels, 1=inches
    frontRectHeight: "72",
    frontRectHeightUnit: 0, // 0=pixels, 1=inches
    frontRectWidth: "72",
    frontRectWidthUnit: 0, // 0=pixels, 1=inches
    frontPosLabelFontSize: "12",
    frontPosLabelFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    frontSkuFontSize: "10",
    frontSkuFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    frontProductFontSize: "12",
    frontProductFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    frontPriceFontSize: "14",
    frontPriceFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    frontTracking: "0",
    frontTrackingUnit: 0, // 0=auto, 1=pixels, 2=inches
    frontLeading: "",
    frontLeadingUnit: 0, // 0=auto, 1=pixels, 2=inches
    frontBorderColor: "#000000",
    
    // List View
    listWidth: "300",
    listWidthUnit: 0, // 0=pixels, 1=inches
    listIdentifiers: 0, // 0=Identifiers, 1=No Identifiers
    listDividerSpacing: "10",
    listDividerSpacingUnit: 0, // 0=pixels, 1=inches
    listObjectSpacing: "10",
    listObjectSpacingUnit: 1, // 0=auto, 1=pixels, 2=inches
    listTextSpacing: "5",
    listTextSpacingUnit: 1, // 0=auto, 1=pixels, 2=inches
    listPosNumFontSize: "12",
    listPosNumFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    listSkuFontSize: "12",
    listSkuFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    listProductFontSize: "12",
    listProductFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    listPriceFontSize: "12",
    listPriceFontSizeUnit: 1, // 0=auto, 1=pixels, 2=inches
    listTracking: "0",
    listTrackingUnit: 0, // 0=auto, 1=pixels, 2=inches
    listLeading: "",
    listLeadingUnit: 0, // 0=auto, 1=pixels, 2=inches
    listBorderColor: "#000000"
};

var win = new Window("dialog", "GhostByte PlanoMako", undefined, {resizeable:true});
win.orientation = "column";
win.alignChildren = ["fill","top"];

// === Tabs ===
var tabs = win.add("tabbedpanel");
tabs.alignChildren = ["fill", "fill"];
tabs.preferredSize = [800, 600];

// === Data Tab ===
var tabData = tabs.add("tab", undefined, "Data");
tabData.orientation = "column";
tabData.alignChildren = ["fill", "top"];

// === Features Tab ===
var tabFeatureColors = tabs.add("tab", undefined, "Features");
tabFeatureColors.orientation = "column";
tabFeatureColors.alignChildren = ["fill","top"];
tabFeatureColors.spacing = 10;

// === Front View Tab ===
var tabLayout = tabs.add("tab", undefined, "Front View");
tabLayout.orientation = "column";
tabLayout.alignChildren = ["fill","top"];
tabLayout.spacing = 10;

// === List View Tab ===
var tabList = tabs.add("tab", undefined, "List View");
tabList.orientation = "column";
tabList.alignChildren = ["fill","top"];
tabList.spacing = 10;

// Add Feature button at the top
var addFeatureBtn = tabFeatureColors.add("button", undefined, "Add Feature");
addFeatureBtn.alignment = "left";

// Default pastel colors for features (20 colors total)
var defaultFeatureColors = [
    "#F6DFDD", // Light pink
    "#D3E5F4", // Light blue
    "#D7EBE6", // Light mint
    "#FFE6CC", // Light peach
    "#E6D9F2", // Light lavender
    "#FFF4CC", // Light yellow
    "#FFD9E6", // Light rose
    "#D9F2E6", // Light seafoam
    "#F2E6D9", // Light tan
    "#E6F2FF", // Light sky blue
    "#FFE6F2", // Light pink-purple
    "#E6FFE6", // Light lime
    "#F2D9E6", // Light mauve
    "#D9E6F2", // Light periwinkle
    "#FFEBE6", // Light coral
    "#E6F2D9", // Light pistachio
    "#F2E6FF", // Light lilac
    "#E6FFFF", // Light cyan
    "#FFF2E6", // Light cream
    "#E6E6FF"  // Light blue-purple
];

// Container for dynamically created feature color controls
var featureColorContainer = tabFeatureColors.add("group");
featureColorContainer.orientation = "column";
featureColorContainer.alignChildren = ["fill","top"];
featureColorContainer.spacing = 5;

// Arrays to store dynamic color control references
var featureColorInputs = [];
var featureColorSwatches = [];
var featureColorPickerBtns = [];

// Function to rebuild feature color controls based on current feature columns
function rebuildFeatureColors(){
    // Clear existing controls
    while(featureColorContainer.children.length > 0){
        featureColorContainer.remove(featureColorContainer.children[0]);
    }
    
    // Clear arrays
    featureColorInputs = [];
    featureColorSwatches = [];
    featureColorPickerBtns = [];
    
    // Create color picker for each feature
    for(var i = 0; i < featureColumns.length; i++){
        var featureName = featureColumns[i];
        var defaultColor = defaultFeatureColors[i % defaultFeatureColors.length];
        
        var colorGroup = featureColorContainer.add("group");
        colorGroup.orientation = "row";
        
        var label = colorGroup.add("statictext", undefined, featureName + " Color (hex):");
        label.preferredSize.width = 180;
        
        var colorInput = colorGroup.add("edittext", undefined, defaultColor);
        colorInput.characters = 8;
        
        var colorSwatch = colorGroup.add("panel", undefined, undefined, {borderStyle: "black"});
        colorSwatch.size = [20, 20];
        
        var pickerBtn = colorGroup.add("button", undefined, "Pick...");
        
        // Store references
        featureColorInputs.push(colorInput);
        featureColorSwatches.push(colorSwatch);
        featureColorPickerBtns.push(pickerBtn);
        
        // Set up event handlers using closure
        (function(input, swatch, btn){
            btn.onClick = function(){
                var chosen = showColorPickerDialog(input.text);
                if(chosen) {
                    input.text = chosen;
                    updateColorSwatch(swatch, chosen);
                }
            };
            input.onChange = function(){
                updateColorSwatch(swatch, input.text);
            };
            // Initialize swatch
            updateColorSwatch(swatch, input.text);
        })(colorInput, colorSwatch, pickerBtn);
    }
    
    // Force layout update
    featureColorContainer.layout.layout(true);
    tabFeatureColors.layout.layout(true);
}

// Initial build with empty features
rebuildFeatureColors();

// Copy all controls from Front View tab for List View
// Width
var listWidthGroup = tabList.add("group");
listWidthGroup.orientation="row";
listWidthGroup.add("statictext", undefined, "Total Width:");
var listWidthInput = listWidthGroup.add("edittext", undefined, "300"); listWidthInput.characters=6;
var listWidthUnit = listWidthGroup.add("dropdownlist", undefined, ["pixels","inches"]); listWidthUnit.selection=0;

// Identifiers
var listIdGroup = tabList.add("group");
listIdGroup.orientation="row";
listIdGroup.add("statictext", undefined, "Identifiers:");
var listIdDrop = listIdGroup.add("dropdownlist", undefined, ["Identifiers","No Identifiers"]); listIdDrop.selection=0;

// Group Divider spacing (new control)
var listDividerGroup = tabList.add("group");
listDividerGroup.orientation="row";
listDividerGroup.add("statictext", undefined, "Group divider spacing:");
var listDividerInput = listDividerGroup.add("edittext", undefined, "10"); listDividerInput.characters=4;
var listDividerUnit = listDividerGroup.add("dropdownlist", undefined, ["pixels","inches"]); listDividerUnit.selection=0;

// Object Spacing
var listPadGroup = tabList.add("group");
listPadGroup.orientation="row";
listPadGroup.add("statictext", undefined, "Object spacing:");
var listPadInput = listPadGroup.add("edittext", undefined, "10"); listPadInput.characters=4;
var listPadUnit = listPadGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listPadUnit.selection=1;

// Text Spacing
var listTextGroup = tabList.add("group");
listTextGroup.orientation="row";
listTextGroup.add("statictext", undefined, "Text spacing:");
var listTextSpaceInput = listTextGroup.add("edittext", undefined, "5"); listTextSpaceInput.characters=4;
var listTextSpaceUnit = listTextGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listTextSpaceUnit.selection=1;

// Position Number Font Controls
var listPosNumFontGroup = tabList.add("group");
listPosNumFontGroup.orientation="row";
listPosNumFontGroup.add("statictext", undefined, "Position Number Font:");
var listPosNumFontDrop = listPosNumFontGroup.add("dropdownlist", undefined, []);
var listPosNumStyleDrop = listPosNumFontGroup.add("dropdownlist", undefined, []);
var listPosNumSizeInput = listPosNumFontGroup.add("edittext", undefined, "12"); listPosNumSizeInput.characters=4;
var listPosNumSizeUnit = listPosNumFontGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listPosNumSizeUnit.selection=1;

// SKU Font Controls
var listSkuFontGroup = tabList.add("group");
listSkuFontGroup.orientation="row";
listSkuFontGroup.add("statictext", undefined, "SKU Font:");
var listSkuFontDrop = listSkuFontGroup.add("dropdownlist", undefined, []);
var listSkuStyleDrop = listSkuFontGroup.add("dropdownlist", undefined, []);
var listSkuSizeInput = listSkuFontGroup.add("edittext", undefined, "12"); listSkuSizeInput.characters=4;
var listSkuSizeUnit = listSkuFontGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listSkuSizeUnit.selection=1;

// Product Name Font Controls
var listProductFontGroup = tabList.add("group");
listProductFontGroup.orientation="row";
listProductFontGroup.add("statictext", undefined, "Product Name Font:");
var listProductFontDrop = listProductFontGroup.add("dropdownlist", undefined, []);
var listProductStyleDrop = listProductFontGroup.add("dropdownlist", undefined, []);
var listProductSizeInput = listProductFontGroup.add("edittext", undefined, "12"); listProductSizeInput.characters=4;
var listProductSizeUnit = listProductFontGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listProductSizeUnit.selection=1;

// Price Font Controls
var listPriceFontGroup = tabList.add("group");
listPriceFontGroup.orientation="row";
listPriceFontGroup.add("statictext", undefined, "Price Font:");
var listPriceFontDrop = listPriceFontGroup.add("dropdownlist", undefined, []);
var listPriceStyleDrop = listPriceFontGroup.add("dropdownlist", undefined, []);
var listPriceSizeInput = listPriceFontGroup.add("edittext", undefined, "12"); listPriceSizeInput.characters=4;
var listPriceSizeUnit = listPriceFontGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listPriceSizeUnit.selection=1;

// Tracking
var listTrackGroup = tabList.add("group");
listTrackGroup.orientation="row";
listTrackGroup.add("statictext", undefined, "Tracking:");
var listTrackingInput = listTrackGroup.add("edittext", undefined, "0"); listTrackingInput.characters=4;
var listTrackingUnit = listTrackGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listTrackingUnit.selection=0;

// Leading
var listLeadGroup = tabList.add("group");
listLeadGroup.orientation="row";
listLeadGroup.add("statictext", undefined, "Leading:");
var listLeadingInput = listLeadGroup.add("edittext", undefined, ""); listLeadingInput.characters=4;
var listLeadingUnit = listLeadGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listLeadingUnit.selection=0;

// Border
var listBorderGroup = tabList.add("group");
listBorderGroup.orientation="row";
listBorderGroup.add("statictext", undefined, "Border width:");
var listBorderWidthInput = listBorderGroup.add("edittext", undefined, "0.5"); listBorderWidthInput.characters=4;
var listBorderWidthUnit = listBorderGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); listBorderWidthUnit.selection=1;

listBorderGroup.add("statictext", undefined, "Style:");
var listBorderStyleDrop = listBorderGroup.add("dropdownlist", undefined, ["Solid", "Dashed", "Dotted"]);
listBorderStyleDrop.selection = 0;

listBorderGroup.add("statictext", undefined, "Color (hex):");
var listBorderColorInput = listBorderGroup.add("edittext", undefined, "#000000");
listBorderColorInput.characters = 8;
var listBorderColorSwatch = listBorderGroup.add("panel", undefined, undefined, {borderStyle: "black"});
listBorderColorSwatch.size = [20, 20];
var listBorderColorPickerBtn = listBorderGroup.add("button", undefined, "Pick...");

// === Data Tab Controls ===
// (Import CSV and Add Row buttons moved to pagination row)

// === FRONT VIEW CONTROLS ===

// --- Width ---
var widthGroup = tabLayout.add("group");
widthGroup.orientation="row";
widthGroup.add("statictext", undefined, "Total Width:");
var widthInput = widthGroup.add("edittext", undefined, "360"); widthInput.characters=6;
var widthUnit = widthGroup.add("dropdownlist", undefined, ["pixels","inches"]); widthUnit.selection=0;

// --- Identifiers ---
var idGroup = tabLayout.add("group");
idGroup.orientation="row";
idGroup.add("statictext", undefined, "Identifiers:");
var idDrop = idGroup.add("dropdownlist", undefined, ["Identifiers","No Identifiers"]); idDrop.selection=0;

// --- Object Spacing ---
var padGroup = tabLayout.add("group");
padGroup.orientation="row";
padGroup.add("statictext", undefined, "Object spacing:");
var padInput = padGroup.add("edittext", undefined, "5"); padInput.characters=4;
var padUnit = padGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); padUnit.selection=1;

// --- Group Spacing ---
var groupSpacingGroup = tabLayout.add("group");
groupSpacingGroup.orientation="row";
groupSpacingGroup.add("statictext", undefined, "Group spacing:");
var groupSpacingInput = groupSpacingGroup.add("edittext", undefined, "18"); groupSpacingInput.characters=4;
var groupSpacingUnit = groupSpacingGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); groupSpacingUnit.selection=1;

// --- Text Spacing ---
var textGroup = tabLayout.add("group");
textGroup.orientation="row";
textGroup.add("statictext", undefined, "Text spacing:");
var textSpaceInput = textGroup.add("edittext", undefined, "5"); textSpaceInput.characters=4;
var textSpaceUnit = textGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); textSpaceUnit.selection=1;

// --- Background Elements ---
var bgElementsGroup = tabLayout.add("group");
bgElementsGroup.orientation="row";
bgElementsGroup.add("statictext", undefined, "Background Elements:");
var bgElementsDrop = bgElementsGroup.add("dropdownlist", undefined, ["Include","None"]); bgElementsDrop.selection=0;

// --- Font ---
var fontGroup = tabLayout.add("group");
fontGroup.orientation="row";
fontGroup.add("statictext", undefined, "Font:");
var fontDrop = fontGroup.add("dropdownlist", undefined, []);
var styleDrop = fontGroup.add("dropdownlist", undefined, []);

// --- Font Size ---
var sizeGroup = tabLayout.add("group");
sizeGroup.orientation="row";
sizeGroup.add("statictext", undefined, "Font size:");
var fontSizeInput = sizeGroup.add("edittext", undefined, "12"); fontSizeInput.characters=4;
var fontSizeUnit = sizeGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); fontSizeUnit.selection=1;

// --- Tracking ---
var trackGroup = tabLayout.add("group");
trackGroup.orientation="row";
trackGroup.add("statictext", undefined, "Tracking:");
var trackingInput = trackGroup.add("edittext", undefined, "0"); trackingInput.characters=4;
var trackingUnit = trackGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); trackingUnit.selection=0;

// --- Leading ---
var leadGroup = tabLayout.add("group");
leadGroup.orientation="row";
leadGroup.add("statictext", undefined, "Leading:");
var leadingInput = leadGroup.add("edittext", undefined, ""); leadingInput.characters=4;
var leadingUnit = leadGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); leadingUnit.selection=0;

// --- Border ---
var borderGroup = tabLayout.add("group");
borderGroup.orientation="row";
borderGroup.add("statictext", undefined, "Border width:");
var borderWidthInput = borderGroup.add("edittext", undefined, "0.5"); borderWidthInput.characters=4;
var borderWidthUnit = borderGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]); borderWidthUnit.selection=1;

borderGroup.add("statictext", undefined, "Style:");
var borderStyleDrop = borderGroup.add("dropdownlist", undefined, ["Solid", "Dashed", "Dotted"]);
borderStyleDrop.selection = 0;

borderGroup.add("statictext", undefined, "Color (hex):");
var borderColorInput = borderGroup.add("edittext", undefined, "#000000");
borderColorInput.characters = 8;
var borderColorSwatch = borderGroup.add("panel", undefined, undefined, {borderStyle: "black"});
borderColorSwatch.size = [20, 20];
var borderColorPickerBtn = borderGroup.add("button", undefined, "Pick...");

// Show a simple ScriptUI color picker (RGB sliders + hex preview)
function hexToRgb(hex){
    if(!hex) return null;
    hex = hex.replace(/[^0-9a-fA-F]/g,'');
    if(hex.length===3){
        hex = hex.split('').map(function(c){return c+c;}).join('');
    }
    if(hex.length!==6) return null;
    var r = parseInt(hex.substring(0,2),16);
    var g = parseInt(hex.substring(2,4),16);
    var b = parseInt(hex.substring(4,6),16);
    if(isNaN(r)||isNaN(g)||isNaN(b)) return null;
    return [r,g,b];
}
function hexToRgbArray(hex){
    var rgb = hexToRgb(hex);
    if(!rgb) return [1,1,1]; // Default to white if invalid
    return [rgb[0]/255, rgb[1]/255, rgb[2]/255];
}

function updateColorSwatch(swatch, hexValue) {
    var rgbArray = hexToRgbArray(hexValue);
    swatch.onDraw = function() {
        try {
            var g = this.graphics;
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, rgbArray);
            var pen = g.newPen(g.PenType.SOLID_COLOR, [0,0,0], 1);
            g.rectPath(0, 0, this.size[0], this.size[1]);
            g.fillPath(brush);
            g.strokePath(pen);
        } catch(e) {}
    };
}
function rgbToHex(r,g,b){
    function c(v){
        v = Math.max(0, Math.min(255, Math.round(v)));
        var s = v.toString(16);
        return (s.length===1)?"0"+s:s;
    }
    return "#"+c(r)+c(g)+c(b);
}

// Parse a color input which can be hex (#RRGGBB or #RGB) or "R,G,B"; returns [r,g,b] or null
function parseColorInput(text){
    if(text===undefined||text===null) return null;
    var t = (text && text.toString) ? text.toString() : String(text);
    // ExtendScript may not implement String.prototype.trim, use regex
    t = t.replace(/^\s+|\s+$/g, "");
    var byHex = hexToRgb(t);
    if(byHex) return byHex;
    var parts = t.split(",");
    if(parts.length===3){
        var pr = parseInt(parts[0],10);
        var pg = parseInt(parts[1],10);
        var pb = parseInt(parts[2],10);
        if(!isNaN(pr)&&!isNaN(pg)&&!isNaN(pb)) return [pr,pg,pb];
    }
    return null;
}

// Helper function to get active features for a row
function getActiveFeatures(rowData){
    var activeFeatures = [];
    for(var i=0; i<featureColumns.length; i++){
        var featureName = featureColumns[i];
        if(rowData[featureName] === "X"){
            activeFeatures.push(featureName);
        }
    }
    return activeFeatures;
}

// Helper function to get color for a feature from the Feature Colors tab
function getFeatureColor(featureName){
    // Find the index of this feature in featureColumns
    var featureIndex = -1;
    for(var i = 0; i < featureColumns.length; i++){
        if(featureColumns[i] === featureName){
            featureIndex = i;
            break;
        }
    }
    
    // If found and we have a color input for it, use that
    if(featureIndex >= 0 && featureIndex < featureColorInputs.length){
        var colorFromInput = parseColorInput(featureColorInputs[featureIndex].text);
        if(colorFromInput) return colorFromInput;
    }
    
    // Fallback to default color based on index
    if(featureIndex >= 0){
        var defaultHex = defaultFeatureColors[featureIndex % defaultFeatureColors.length];
        return parseColorInput(defaultHex) || [200, 200, 200];
    }
    
    // Ultimate fallback
    return [200, 200, 200];
}

function showColorPickerDialog(initial){
    var rgb = hexToRgb(initial) || [0,0,0];
    var dlg = new Window("dialog", "Pick Color");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill","top"];

    var preview = dlg.add("panel");
    preview.preferredSize = [200,40];

    function updatePreview(){
        var col = rgbToHex(rgb[0],rgb[1],rgb[2]);
        preview.graphics.backgroundColor = preview.graphics.newBrush(preview.graphics.BrushType.SOLID_COLOR, [rgb[0]/255, rgb[1]/255, rgb[2]/255, 1]);
        hexField.text = col;
    }

    var sliders = dlg.add("group"); sliders.orientation = "column";
    var rGroup = sliders.add("group"); rGroup.add("statictext", undefined, "R:");
    var rSlider = rGroup.add("slider", undefined, rgb[0], 0, 255); rSlider.preferredSize.width = 180;
    var rVal = rGroup.add("edittext", undefined, rgb[0].toString()); rVal.characters=4;

    var gGroup = sliders.add("group"); gGroup.add("statictext", undefined, "G:");
    var gSlider = gGroup.add("slider", undefined, rgb[1], 0, 255); gSlider.preferredSize.width = 180;
    var gVal = gGroup.add("edittext", undefined, rgb[1].toString()); gVal.characters=4;

    var bGroup = sliders.add("group"); bGroup.add("statictext", undefined, "B:");
    var bSlider = bGroup.add("slider", undefined, rgb[2], 0, 255); bSlider.preferredSize.width = 180;
    var bVal = bGroup.add("edittext", undefined, rgb[2].toString()); bVal.characters=4;

    var hexRow = dlg.add("group");
    hexRow.add("statictext", undefined, "Hex:");
    var hexField = hexRow.add("edittext", undefined, rgbToHex(rgb[0],rgb[1],rgb[2])); hexField.characters = 8;

    var okCancel = dlg.add("group"); okCancel.alignment = "right";
    var okBtn = okCancel.add("button", undefined, "OK");
    var cancelBtn = okCancel.add("button", undefined, "Cancel");

    function syncFromSliders(){
        rgb[0] = Math.round(rSlider.value);
        rgb[1] = Math.round(gSlider.value);
        rgb[2] = Math.round(bSlider.value);
        rVal.text = rgb[0]; gVal.text = rgb[1]; bVal.text = rgb[2];
        hexField.text = rgbToHex(rgb[0],rgb[1],rgb[2]);
        updatePreview();
    }
    function syncFromEdits(){
        var rv = parseInt(rVal.text,10); if(isNaN(rv)) rv=0; rv=Math.max(0,Math.min(255,rv));
        var gv = parseInt(gVal.text,10); if(isNaN(gv)) gv=0; gv=Math.max(0,Math.min(255,gv));
        var bv = parseInt(bVal.text,10); if(isNaN(bv)) bv=0; bv=Math.max(0,Math.min(255,bv));
        rSlider.value = rv; gSlider.value = gv; bSlider.value = bv;
        syncFromSliders();
    }

    rSlider.onChanging = syncFromSliders; gSlider.onChanging = syncFromSliders; bSlider.onChanging = syncFromSliders;
    rVal.onChange = syncFromEdits; gVal.onChange = syncFromEdits; bVal.onChange = syncFromEdits;

    hexField.onChange = function(){
        var h = hexField.text;
        var parsed = hexToRgb(h);
        if(parsed){ rgb = parsed; rSlider.value=rgb[0]; gSlider.value=rgb[1]; bSlider.value=rgb[2]; syncFromSliders(); }
    };

    updatePreview();

    if(dlg.show()==1){
        return hexField.text;
    }
    return null;
}

// Feature Colors Tab - handlers are now set up dynamically in rebuildFeatureColors()

// Add Feature button handler
addFeatureBtn.onClick = function(){
    // Prompt for feature name
    var featureName = prompt("Enter feature name:", "Feature " + (featureColumns.length + 1));
    if(!featureName) return; // User cancelled
    
    // Trim whitespace
    featureName = featureName.replace(/^\s+|\s+$/g, "");
    if(!featureName) {
        alert("Feature name cannot be empty.");
        return;
    }
    
    // Check if feature already exists
    for(var i = 0; i < featureColumns.length; i++){
        if(featureColumns[i] === featureName){
            alert("Feature '" + featureName + "' already exists.");
            return;
        }
    }
    
    // Save current data first (before adding new feature)
    saveRows();
    
    // Add the new feature to featureColumns
    featureColumns.push(featureName);
    
    // Add the feature column to all existing rows (initialize as empty)
    for(var r = 0; r < csvData.length; r++){
        csvData[r][featureName] = ""; // Initialize as empty
    }
    
    // Rebuild the feature colors tab to show new color picker
    rebuildFeatureColors();
    
    // Rebuild the data table to show new checkbox column
    rebuildTable();
};

borderColorPickerBtn.onClick = function(){
    var chosen = showColorPickerDialog(borderColorInput.text);
    if(chosen) {
        borderColorInput.text = chosen;
        updateColorSwatch(borderColorSwatch, chosen);
    }
};

listBorderColorPickerBtn.onClick = function(){
    var chosen = showColorPickerDialog(listBorderColorInput.text);
    if(chosen) {
        listBorderColorInput.text = chosen;
        updateColorSwatch(listBorderColorSwatch, chosen);
    }
};

// onChange handlers to update swatches when hex values are typed
borderColorInput.onChange = function(){ updateColorSwatch(borderColorSwatch, borderColorInput.text); };
listBorderColorInput.onChange = function(){ updateColorSwatch(listBorderColorSwatch, listBorderColorInput.text); };

// Initialize swatches with default colors
updateColorSwatch(borderColorSwatch, "#000000");
updateColorSwatch(listBorderColorSwatch, "#000000");

// --- Rectangle Background Color ---
var rectBgGroup = tabLayout.add("group");
rectBgGroup.orientation = "row";
rectBgGroup.add("statictext", undefined, "Rect Fill (hex):");
var rectFillInput = rectBgGroup.add("edittext", undefined, "#ffffff"); rectFillInput.characters = 8;
var rectFillPick = rectBgGroup.add("button", undefined, "Pick...");
rectFillPick.onClick = function(){ var c = showColorPickerDialog(rectFillInput.text); if(c) rectFillInput.text = c; };

// --- Circle Background Color ---
var circBgGroup = tabLayout.add("group");
circBgGroup.orientation = "row";
circBgGroup.add("statictext", undefined, "Circle Fill (hex):");
var circFillInput = circBgGroup.add("edittext", undefined, "#124B8E"); circFillInput.characters = 8;
var circFillPick = circBgGroup.add("button", undefined, "Pick...");
circFillPick.onClick = function(){ var c = showColorPickerDialog(circFillInput.text); if(c) circFillInput.text = c; };

// --- Circle Text Color ---
var circTextGroup = tabLayout.add("group");
circTextGroup.orientation = "row";
circTextGroup.add("statictext", undefined, "Circle Text (hex):");
var circTextInput = circTextGroup.add("edittext", undefined, "#ffffff"); circTextInput.characters = 8;
var circTextPick = circTextGroup.add("button", undefined, "Pick...");
circTextPick.onClick = function(){ var c = showColorPickerDialog(circTextInput.text); if(c) circTextInput.text = c; };

// --- Rounded Corners ---
var roundGroup = tabLayout.add("group");
roundGroup.orientation = "row";
roundGroup.add("statictext", undefined, "Rounded corners:");
var cornerRadiusInput = roundGroup.add("edittext", undefined, "0");
cornerRadiusInput.characters = 4;
var cornerRadiusUnit = roundGroup.add("dropdownlist", undefined, ["auto","pixels","inches"]);
cornerRadiusUnit.selection = 1;

// === Populate Fonts ===
function populateFontDropdown(dropdown, styleDropdown, preferredStyle) {
    try {
        var fonts = app.textFonts;
        // Build a unique list of font families
        var families = {};
        var familyOrder = [];
        for (var i=0; i<fonts.length; i++){
            var fam = fonts[i].family || fonts[i].name || "";
            if(!families[fam]){ families[fam]=true; familyOrder.push(fam); }
        }
        for(var fi=0; fi<familyOrder.length; fi++) dropdown.add("item", familyOrder[fi]);
        
        // Prefer ITCFranklinGothic LT Pro
        var preferredFamily = "ITCFranklinGothic LT Pro";
        var prefLower = preferredFamily.toLowerCase();
        var famIndex = 0;
        for(var k=0;k<familyOrder.length;k++){
            var famName = (familyOrder[k]||"").toString();
            var famLower = famName.toLowerCase();
            if(famLower===prefLower || famLower.indexOf(prefLower)!==-1){ famIndex=k; break; }
        }
        dropdown.selection = famIndex;
        
        // Update styles and select preferred style if specified
        updateFontStyles(dropdown, styleDropdown, preferredStyle);
        
        // Add change handler
        dropdown.onChange = function() {
            updateFontStyles(dropdown, styleDropdown, preferredStyle);
        };
    } catch(e) { alert("Unable to load fonts for " + dropdown.parent.children[0].text + ": "+e); }
}

function updateFontStyles(dropdown, styleDropdown, preferredStyle) {
    styleDropdown.removeAll();
    try {
        var base = dropdown.selection ? dropdown.selection.text : "";
        var seen = {};
        for (var i=0; i<app.textFonts.length; i++){
            var f = app.textFonts[i];
            if(f.family==base && !seen[f.style]){
                styleDropdown.add("item", f.style);
                seen[f.style]=true;
            }
        }
        // Select preferred style if specified and exists, otherwise first style
        if(styleDropdown.items.length>0){
            if (preferredStyle) {
                var prefLower = preferredStyle.toLowerCase();
                var prefIdx = -1;
                for(var k=0; k<styleDropdown.items.length; k++){
                    var s = (styleDropdown.items[k].text||"").toString().toLowerCase();
                    if(s===prefLower || s.indexOf(prefLower)!==-1){ prefIdx=k; break; }
                }
                styleDropdown.selection = (prefIdx>=0) ? styleDropdown.items[prefIdx] : styleDropdown.items[0];
            } else {
                styleDropdown.selection = styleDropdown.items[0];
            }
        }
    } catch(e) {}
}

try {
    // Populate all font dropdowns with their respective preferred styles
    populateFontDropdown(fontDrop, styleDrop);  // Front View
    populateFontDropdown(listPosNumFontDrop, listPosNumStyleDrop, "CnDm");  // Position Number
    populateFontDropdown(listSkuFontDrop, listSkuStyleDrop, "CnDm");  // SKU
    populateFontDropdown(listProductFontDrop, listProductStyleDrop, "CnBk");  // Product Name
    populateFontDropdown(listPriceFontDrop, listPriceStyleDrop, "CnBk");  // Price
} catch(e) { alert("Unable to load fonts: "+e); }

function updateFontStyles(){
    styleDrop.removeAll();
    try {
        var base = fontDrop.selection?fontDrop.selection.text:"";
        var seen = {};
        for (var i=0;i<app.textFonts.length;i++){
            var f = app.textFonts[i];
            if(f.family==base && !seen[f.style]){
                styleDrop.add("item", f.style);
                seen[f.style]=true;
            }
        }
        // Prefer style "CnDm" if it exists, otherwise pick the first
        if(styleDrop.items.length>0){
            var preferredStyle = "CnDm";
            var prefLower = preferredStyle.toLowerCase();
            var prefIdx = -1;
            for(var k=0;k<styleDrop.items.length;k++){
                var s = (styleDrop.items[k].text||"").toString().toLowerCase();
                if(s===prefLower || s.indexOf(prefLower)!==-1){ prefIdx=k; break; }
            }
            styleDrop.selection = (prefIdx>=0)? styleDrop.items[prefIdx] : styleDrop.items[0];
        }
    }catch(e){}
}
fontDrop.onChange = updateFontStyles;

// === Data Table ===
// (Data variables declared globally at top of file)

var container = tabData.add("group");
container.orientation = "column";
container.alignChildren = ["fill","top"];
container.minimumSize.height = 360;

// Initialize column widths
var widths = [30,40,180,90,60]; // #, Pos, Product, SKU, Price

// Pagination controls with Add Row and Import CSV buttons
var paginationGroup = container.add("group");
paginationGroup.orientation = "row";
paginationGroup.alignment = "fill";
var addBtn = paginationGroup.add("button", undefined, "Add Row");
var importBtn = paginationGroup.add("button", undefined, "Import CSV");
// Add flexible space to push pagination to center-right
var spacer = paginationGroup.add("group");
spacer.alignment = "fill";
var pageInfoLabel = paginationGroup.add("statictext", undefined, "Page 1 of 1");
pageInfoLabel.preferredSize.width = 100;
pageInfoLabel.justify = "center";
var prevPageBtn = paginationGroup.add("button", undefined, "Previous");
var nextPageBtn = paginationGroup.add("button", undefined, "Next");

var scrollPanel = container.add("panel");
scrollPanel.orientation = "column";
scrollPanel.alignChildren = ["fill","top"];
scrollPanel.minimumSize = [700, 260];
// Enable scrolling
scrollPanel.properties = {
    borderless: true,
    su1PanelCoordinates: true
};
scrollPanel.spacing = 2; // Add some spacing between rows

// === CSV Helpers ===
function parseCSV(text){
    var lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    
    // Parse feature color metadata from comment lines
    var featureColorMap = {};
    var dataLines = [];
    for(var i = 0; i < lines.length; i++){
        var line = lines[i] || "";
        line = line.replace(/^\s+|\s+$/g, "");
        
        // Check if it's a color metadata comment: # FeatureName=ColorValue
        if(line.indexOf("# ") === 0 && line.indexOf("=") > 0){
            var metaLine = line.substring(2); // Remove "# "
            var equalPos = metaLine.indexOf("=");
            if(equalPos > 0){
                var featureName = metaLine.substring(0, equalPos);
                featureName = featureName.replace(/^\s+|\s+$/g, "");
                var colorValue = metaLine.substring(equalPos + 1);
                colorValue = colorValue.replace(/^\s+|\s+$/g, "");
                featureColorMap[featureName] = colorValue;
            }
        } else if(line.indexOf("#") !== 0){
            // Not a comment line, keep it
            dataLines.push(line);
        }
        // Skip other comment lines
    }
    
    var headers = parseCSVLine(dataLines.shift());
    
    // Find the "Features" column index
    var featuresIndex = -1;
    for(var h=0; h<headers.length; h++){
        var header = headers[h] || "";
        header = header.replace(/^\s+|\s+$/g, "");
        if(header === "Features"){
            featuresIndex = h;
            break;
        }
    }
    
    // Determine which columns are features (all columns after "Features")
    featureColumns = [];
    if(featuresIndex >= 0){
        for(var f=featuresIndex+1; f<headers.length; f++){
            var featureHeader = headers[f] || "";
            featureHeader = featureHeader.replace(/^\s+|\s+$/g, "");
            if(featureHeader){
                featureColumns.push(featureHeader);
            }
        }
    }
    
    var data = [];
    
    for (var i = 0; i < dataLines.length; i++){
        var line = dataLines[i] || "";
        // trim via regex for ExtendScript compatibility
        line = line.replace(/^\s+|\s+$/g, "");
        if(!line) continue;
        
        var parts = parseCSVLine(line);
        var row = {};
        for (var j = 0; j < headers.length; j++){
            var h = headers[j] || "";
            h = h.replace(/^\s+|\s+$/g, "");
            var p = parts[j] || "";
            p = p.replace(/^\s+|\s+$/g, "");
            row[h] = p;
        }
        // Keep all non-empty rows
        var hasAnyData = false;
        for(var key in row){ if(row[key]){ hasAnyData = true; break; } }
        if(hasAnyData) data.push(row);
    }
    
    // Return both data and feature color map
    return {data: data, featureColors: featureColorMap};
}

// Helper function to parse a single CSV line handling quoted fields
function parseCSVLine(line){
    var result = [];
    var current = "";
    var inQuotes = false;
    
    for(var i = 0; i < line.length; i++){
        var c = line.charAt(i);
        var nextChar = (i + 1 < line.length) ? line.charAt(i + 1) : "";
        
        if(c === '"'){
            if(inQuotes && nextChar === '"'){
                // Escaped quote (two quotes in a row)
                current += '"';
                i++; // Skip the next quote
            } else {
                // Toggle quote mode
                inQuotes = !inQuotes;
            }
        } else if(c === ',' && !inQuotes){
            // End of field
            result.push(current);
            current = "";
        } else {
            current += c;
        }
    }
    // Add the last field
    result.push(current);
    return result;
}
function getUniquePositions(data){
    var seen={}, order=[];
    for (var i=0;i<data.length;i++){
        var n=parseInt(data[i]["Position"],10);
        if(!isNaN(n)&&!seen[n]){ seen[n]=true; order.push(n); }
    }
    order.sort(function(a,b){return a-b;});
    return order;
}
function highestPosition(){
    var max=0;
    for (var i=0;i<csvData.length;i++){
        var n=parseInt(csvData[i]["Position"],10);
        if(!isNaN(n)&&n>max) max=n;
    }
    return max;
}
function reorderData(){
    csvData.sort(function(a,b){
        var na=parseInt(a["Position"],10);
        var nb=parseInt(b["Position"],10);
        if(isNaN(na)) na=9999;
        if(isNaN(nb)) nb=9999;
        return na-nb;
    });
}

// === Table ===
function rebuildTable(){
    while(scrollPanel.children.length>0) scrollPanel.remove(scrollPanel.children[0]);
    
    // Calculate pagination
    var totalRows = csvData.length;
    var totalPages = Math.ceil(totalRows / rowsPerPage);
    if(totalPages === 0) totalPages = 1;
    
    // Ensure current page is valid
    if(currentPage < 1) currentPage = 1;
    if(currentPage > totalPages) currentPage = totalPages;
    
    // Update page info label
    pageInfoLabel.text = "Page " + currentPage + " of " + totalPages;
    
    // Enable/disable pagination buttons
    prevPageBtn.enabled = (currentPage > 1);
    nextPageBtn.enabled = (currentPage < totalPages);
    
    // Calculate which rows to display
    var startIndex = (currentPage - 1) * rowsPerPage;
    var endIndex = Math.min(startIndex + rowsPerPage, totalRows);
    
    var positions=getUniquePositions(csvData);
    for(var p=0;p<positions.length;p++){
        // Add spacing before position label (except first)
        if(p > 0) {
            scrollPanel.add("panel", undefined, undefined, {height: 10});
        }
        
        var labelGroup=scrollPanel.add("group");
        labelGroup.orientation="row";
        labelGroup.alignChildren = ["left", "center"];
        var posNum=positions[p];
        
        var labelText = labelGroup.add("statictext",undefined,"Position "+posNum+" label:");
        labelText.graphics.font = ScriptUI.newFont(labelText.graphics.font.name, ScriptUI.FontStyle.BOLD, labelText.graphics.font.size);
        
        var letter=String.fromCharCode(65 + p);
        var defaultText=(p===0)? letter+" (Pegboard)" : letter;
        var labelInput = labelGroup.add("edittext",undefined,defaultText);
        labelInput.characters=30;
        labelInput.graphics.font = ScriptUI.newFont(labelInput.graphics.font.name, ScriptUI.FontStyle.BOLD, labelInput.graphics.font.size);
        
        labelGroup.add("statictext",undefined," Type:");
        var posTypeDrop = labelGroup.add("dropdownlist",undefined,["Peg","Shelf"]);
        posTypeDrop.selection = (p===0) ? 0 : 1; // First position is Peg, others are Shelf
    }
    
    // Add horizontal rule after position labels
    var hrPanel = scrollPanel.add("panel");
    hrPanel.alignment = "fill";
    hrPanel.minimumSize = [0, 2];
    hrPanel.maximumSize = [10000, 2];
    
    // Add header row after position labels, before data rows
    var headerGroup = scrollPanel.add("group");
    headerGroup.orientation = "row";
    headerGroup.spacing = 0;
    
    // Build header labels dynamically
    var headerLabels = ["#","Pos","Product","SKU","Price"];
    var currentWidths = widths.slice(); // Copy base widths
    
    // Add feature column headers
    for(var i=0; i<featureColumns.length; i++){
        headerLabels.push(featureColumns[i]);
        currentWidths.push(70); // Default width for feature columns
    }
    
    // Add delete column header
    headerLabels.push("");
    currentWidths.push(25);
    
    // Create header cells
    for (var j=0; j<headerLabels.length; j++){
        var h = headerGroup.add("statictext", undefined, headerLabels[j]);
        h.preferredSize.width = currentWidths[j];
        h.justify = "left";
        h.graphics.font = ScriptUI.newFont(h.graphics.font.name, ScriptUI.FontStyle.BOLD, h.graphics.font.size);
    }
    
    // Only display rows for current page
    for(var r=startIndex; r<endIndex; r++){
        var g=scrollPanel.add("group");
        g.orientation="row";
        g.spacing=0;
        // Count label - show actual row number (not page-relative)
        var countLabel = g.add("statictext", undefined, (r+1).toString()); countLabel.preferredSize.width = currentWidths[0]; countLabel.justify = "left";
        var p=g.add("edittext",undefined,csvData[r]["Position"]||"");
        p.preferredSize.width=currentWidths[1];
        p.onChanging=function(){this.text=this.text.replace(/[^\d]/g,"");};
        p.onChange=function(){saveRows();reorderData();rebuildTable();};
        var prod=g.add("edittext",undefined,csvData[r]["Product"]||""); prod.preferredSize.width=currentWidths[2];
        var sku=g.add("edittext",undefined,csvData[r]["SKU"]||""); sku.preferredSize.width=currentWidths[3];
        var price=g.add("edittext",undefined,csvData[r]["Price"]||""); price.preferredSize.width=currentWidths[4];
        
        // Dynamically add checkboxes for each feature column
        for(var f=0; f<featureColumns.length; f++){
            var featureName = featureColumns[f];
            var checkbox = g.add("checkbox",undefined,"");
            checkbox.value = (csvData[r][featureName] === "X");
            checkbox.preferredSize.width = currentWidths[5 + f]; // widths index starts at 5 for features
        }
        
        (function(index, group){
            var delBtn = group.add("button",undefined,"X");
            delBtn.preferredSize.width = currentWidths[currentWidths.length - 1]; // Last width is for delete button
            delBtn.onClick=function(){
                csvData.splice(index,1);
                // Stay on same page if possible, otherwise go to last page
                var newTotalPages = Math.ceil(csvData.length / rowsPerPage);
                if(currentPage > newTotalPages && newTotalPages > 0) {
                    currentPage = newTotalPages;
                }
                rebuildTable();
            };
        })(r, g);
    }
    win.layout.layout(true); win.update();
    
    // Update total rows count
    totalRowsLabel.text = "Total Rows: " + totalRows;
}
function saveRows(){
    // With pagination, only update the rows that are currently visible in the UI
    var startIndex = (currentPage - 1) * rowsPerPage;
    var visibleRowIndex = 0;
    
    for(var i=0;i<scrollPanel.children.length;i++){
        var g=scrollPanel.children[i];
        if(g.type=="group"){
            // detect data rows - they should have base columns (5) + feature columns + delete button
            var children = g.children;
            var expectedChildren = 1 + 4 + featureColumns.length + 1; // count + base + features + delete
            if(children.length >= expectedChildren && children[1].type=="edittext"){
                // Calculate the actual index in csvData
                var actualIndex = startIndex + visibleRowIndex;
                
                // Get position type from the position label row for this position
                var posNum = children[1].text;
                var posType = "Shelf"; // default
                
                // Find the position label row for this position
                for(var j=0; j<scrollPanel.children.length; j++){
                    var labelGroup = scrollPanel.children[j];
                    if(labelGroup.type=="group" && labelGroup.children.length>=4){
                        var firstChild = labelGroup.children[0];
                        if(firstChild.type=="statictext" && firstChild.text.indexOf("Position "+posNum+" label:") === 0){
                            // Found the label row, get the dropdown value
                            for(var k=0; k<labelGroup.children.length; k++){
                                if(labelGroup.children[k].type=="dropdownlist"){
                                    posType = labelGroup.children[k].selection ? labelGroup.children[k].selection.text : "Shelf";
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
                
                // Build the row object with base columns
                var rowData = {
                    "Position":children[1].text,
                    "PositionType":posType,
                    "Product":children[2].text,
                    "SKU":children[3].text,
                    "Price":children[4].text
                };
                
                // Add dynamic feature columns
                for(var f=0; f<featureColumns.length; f++){
                    var featureName = featureColumns[f];
                    var checkboxIndex = 5 + f; // checkboxes start at index 5
                    rowData[featureName] = children[checkboxIndex].value ? "X" : "";
                }
                
                // Update the existing row in csvData
                if(actualIndex < csvData.length){
                    csvData[actualIndex] = rowData;
                }
                visibleRowIndex++;
            }
        }
    }
}

// === CSV Import ===
importBtn.onClick=function(){
    var f=File.openDialog("Select CSV");
    if(!f) return;
    f.open("r");
    var text=f.read();
    f.close();
    var parseResult=parseCSV(text);
    var imported = parseResult.data;
    var featureColors = parseResult.featureColors;
    
    if(imported.length===0){ alert("No valid rows found in CSV."); return; }
    
    // Clear existing data
    csvData = [];
    
    // Add each imported row using the Add Row logic
    for(var i=0; i<imported.length; i++){
        var row = imported[i];
        
        // Determine the position number
        var posNum = parseInt(row["Position"], 10);
        if(isNaN(posNum)) posNum = 1;
        
        // Determine default type based on whether this is the first row overall
        var defaultType = (csvData.length === 0) ? "Peg" : "Shelf";
        
        // Build row data with base columns
        var rowData = {
            "Position": posNum.toString(),
            "PositionType": row["PositionType"] || defaultType,
            "Product": row["Product"] || "",
            "SKU": row["SKU"] || "",
            "Price": row["Price"] || ""
        };
        
        // Add feature columns from imported data
        for(var f=0; f<featureColumns.length; f++){
            rowData[featureColumns[f]] = row[featureColumns[f]] || "";
        }
        
        csvData.push(rowData);
    }
    
    currentPage = 1; // Reset to first page after import
    rebuildFeatureColors(); // Rebuild feature color controls based on detected features
    
    // Apply imported feature colors
    for(var fc = 0; fc < featureColumns.length; fc++){
        var featureName = featureColumns[fc];
        if(featureColors[featureName] && fc < featureColorInputs.length){
            featureColorInputs[fc].text = featureColors[featureName];
            updateColorSwatch(featureColorSwatches[fc], featureColors[featureName]);
        }
    }
    
    rebuildTable();
};

// === Add Row ===
addBtn.onClick=function(){
    saveRows();
    var nextPos=highestPosition()||1;
    var defaultType = (csvData.length === 0) ? "Peg" : "Shelf"; // First row is Peg, others are Shelf
    
    // Build new row with base columns
    var newRow = {
        "Position":nextPos.toString(),
        "PositionType":defaultType,
        "Product":"",
        "SKU":"",
        "Price":""
    };
    
    // Add empty feature columns
    for(var f=0; f<featureColumns.length; f++){
        newRow[featureColumns[f]] = "";
    }
    
    csvData.push(newRow);
    
    // Go to the page containing the new row
    var newRowIndex = csvData.length - 1;
    currentPage = Math.ceil(csvData.length / rowsPerPage);
    
    rebuildTable();
    updateRowCount();
};

// === Pagination Buttons ===
prevPageBtn.onClick = function(){
    if(currentPage > 1){
        saveRows();
        currentPage--;
        rebuildTable();
    }
};

nextPageBtn.onClick = function(){
    var totalPages = Math.ceil(csvData.length / rowsPerPage);
    if(currentPage < totalPages){
        saveRows();
        currentPage++;
        rebuildTable();
    }
};

// === Escape Key ===
win.addEventListener("keydown", function(k){
    if(k.keyName=="Escape") win.close();
    if(k.keyName=="Enter") {
        k.preventDefault();
        try {
            if(!app.documents.length){alert("Open a document first.");return;}
            saveRows(); if(csvData.length===0){alert("No rows.");return;}
            
            var savedData = csvData.slice();
            
            // Get position titles from UI before closing window
            var savedPositionTitles = {};
            for(var i=0; i<scrollPanel.children.length; i++){
                var child = scrollPanel.children[i];
                if(child.type=="group" && child.children.length>=4){
                    var firstChild = child.children[0];
                    if(firstChild.type=="statictext" && firstChild.text.indexOf("Position ") === 0 && firstChild.text.indexOf(" label:") > 0){
                        var posMatch = firstChild.text.match(/Position (\d+) label:/);
                        if(posMatch && posMatch[1]){
                            var posNum = posMatch[1];
                            var titleInput = child.children[1];
                            if(titleInput.type=="edittext"){
                                savedPositionTitles[posNum] = titleInput.text;
                            }
                        }
                    }
                }
            }
            
            function toPoints(value, unit){
                if(unit=="inches") return value*72;
                if(unit=="pixels") return value;
                return value;
            }
            var frontViewWidth = toPoints(parseFloat(widthInput.text)||800, widthUnit.selection.text);
            var spacing = 50;
            
            win.close(); // Close window first
            
            // Then execute both view creations
            try {
                createFrontView(false);
            } catch(e) {
                alert("Create Front View error: "+e);
            }
            
            csvData = savedData;
            
            // Create List View with offset
            var originalListClick = createListBtn.onClick;
            createListBtn.onClick = function(){
                try {
                    if(!app.documents.length){alert("Open a document first.");return;}
                    saveRows(); if(csvData.length===0){alert("No rows.");return;}

                    function toPointsLocal(value, unit){
                        if(unit=="inches") return value*72;
                        if(unit=="pixels") return value;
                        return value;
                    }

                    var totalWidth = toPointsLocal(parseFloat(listWidthInput.text)||300, listWidthUnit.selection.text);
                    var pad = listPadUnit.selection.text=="auto"?10:toPointsLocal(parseFloat(listPadInput.text)||10, listPadUnit.selection.text);
                    var groupDivider = toPointsLocal(parseFloat(listDividerInput.text)||10, listDividerUnit.selection.text);
                    var textSpacing = listTextSpaceUnit.selection.text=="auto"?5:toPointsLocal(parseFloat(listTextSpaceInput.text)||5, listTextSpaceUnit.selection.text);
                    
                    var includeIDs=(listIdDrop.selection.text=="Identifiers");
                    
                    function getFontByFamilyAndStyle(familyName, preferredStyle) {
                        try {
                            for(var fi=0; fi<app.textFonts.length; fi++){
                                var f = app.textFonts[fi];
                                if(f.family==familyName && (!preferredStyle || f.style==preferredStyle)) { return f; }
                            }
                            for(var fi2=0; fi2<app.textFonts.length; fi2++){
                                if(app.textFonts[fi2].family==familyName){ return app.textFonts[fi2]; }
                            }
                        } catch(e) {}
                        return app.textFonts[0];
                    }
                    
                    var posNumFont = getFontByFamilyAndStyle(
                        listPosNumFontDrop.selection ? listPosNumFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                        listPosNumStyleDrop.selection ? listPosNumStyleDrop.selection.text : "CnDm"
                    );
                    var skuFont = getFontByFamilyAndStyle(
                        listSkuFontDrop.selection ? listSkuFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                        listSkuStyleDrop.selection ? listSkuStyleDrop.selection.text : "CnDm"
                    );
                    var productFont = getFontByFamilyAndStyle(
                        listProductFontDrop.selection ? listProductFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                        listProductStyleDrop.selection ? listProductStyleDrop.selection.text : "CnBk"
                    );
                    var priceFont = getFontByFamilyAndStyle(
                        listPriceFontDrop.selection ? listPriceFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                        listPriceStyleDrop.selection ? listPriceStyleDrop.selection.text : "CnBk"
                    );
                    
                    var strokeClr = new RGBColor();
                    strokeClr.red = 0;
                    strokeClr.green = 0;
                    strokeClr.blue = 0;

                    var doc=app.activeDocument, layer=doc.activeLayer;
                    var grouped={};
                    for(var i=0;i<csvData.length;i++){
                        var pos=csvData[i]["Position"];
                        if(!grouped[pos])grouped[pos]=[];
                        grouped[pos].push(csvData[i]);
                    }

                    // Use the saved position titles from before window closed
                    var positionTitles = savedPositionTitles;

                    var xOffset = frontViewWidth + spacing;
                    var yPos = 0;
                    var globalCount = 1;
                    var baseHeight = toPointsLocal(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                    var rectHeight = baseHeight * 2.5;
                    var allPositionGroups = []; // Track all position groups for final grouping
                    
                    for(var pos in grouped){
                        var rows = grouped[pos];
                        var posGroup = layer.groupItems.add();
                        posGroup.name = "Position " + pos;
                        
                        for(var j=0; j<rows.length; j++){
                            var row = rows[j];
                            
                            var rect = posGroup.pathItems.rectangle(yPos, xOffset, totalWidth, rectHeight);
                            rect.stroked = true;
                            rect.strokeWidth = 0.5;
                            rect.strokeColor = strokeClr;
                            rect.filled = true;
                            
                            // Get active features for this row dynamically
                            var activeFeatures = getActiveFeatures(row);
                            var featureCount = activeFeatures.length;
                            
                            if(featureCount === 0){
                                // No features - white background
                                var whiteFill = new RGBColor();
                                whiteFill.red = 255;
                                whiteFill.green = 255;
                                whiteFill.blue = 255;
                                rect.fillColor = whiteFill;
                            } else if(featureCount === 1){
                                // One feature - solid color
                                var featureColorArr = getFeatureColor(activeFeatures[0]);
                                var featureColor = new RGBColor();
                                featureColor.red = featureColorArr[0]; featureColor.green = featureColorArr[1]; featureColor.blue = featureColorArr[2];
                                rect.fillColor = featureColor;
                            } else if(featureCount === 2){
                                // Two features - make rect transparent, create two color wedges
                                rect.filled = false;
                                
                                var gradientColors = [];
                                for(var fc=0; fc<activeFeatures.length; fc++){
                                    gradientColors.push(getFeatureColor(activeFeatures[fc]));
                                }
                                
                                var rectBounds = rect.geometricBounds;
                                var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                                var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                                var rectLeft = rectBounds[0];
                                var rectTop = rectBounds[1];
                                
                                // Create first color triangle (top-left to bottom-right)
                                var wedge1 = posGroup.pathItems.add();
                                wedge1.setEntirePath([[rectLeft, rectTop], [rectLeft + rectW, rectTop], [rectLeft, rectTop - rectH]]);
                                wedge1.filled = true;
                                wedge1.stroked = false;
                                var color1 = new RGBColor();
                                color1.red = gradientColors[0][0]; color1.green = gradientColors[0][1]; color1.blue = gradientColors[0][2];
                                wedge1.fillColor = color1;
                                wedge1.move(rect, ElementPlacement.PLACEBEFORE);
                                
                                // Create second color triangle (bottom-right)
                                var wedge2 = posGroup.pathItems.add();
                                wedge2.setEntirePath([[rectLeft + rectW, rectTop], [rectLeft + rectW, rectTop - rectH], [rectLeft, rectTop - rectH]]);
                                wedge2.filled = true;
                                wedge2.stroked = false;
                                var color2 = new RGBColor();
                                color2.red = gradientColors[1][0]; color2.green = gradientColors[1][1]; color2.blue = gradientColors[1][2];
                                wedge2.fillColor = color2;
                                wedge2.move(rect, ElementPlacement.PLACEBEFORE);
                                
                                // Move the transparent rectangle to the front so border is visible
                                rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                            } else {
                                // Three or more features - make rect transparent, create horizontal bars
                                rect.filled = false;
                                
                                var gradientColors = [];
                                for(var fc=0; fc<activeFeatures.length; fc++){
                                    gradientColors.push(getFeatureColor(activeFeatures[fc]));
                                }
                                
                                var rectBounds = rect.geometricBounds;
                                var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                                var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                                var rectLeft = rectBounds[0];
                                var rectTop = rectBounds[1];
                                var barHeight = rectH / featureCount;
                                
                                // Create horizontal bars for each feature
                                for(var b=0; b<featureCount; b++){
                                    var bar = posGroup.pathItems.add();
                                    bar.setEntirePath([
                                        [rectLeft, rectTop - (barHeight * b)],
                                        [rectLeft + rectW, rectTop - (barHeight * b)],
                                        [rectLeft + rectW, rectTop - (barHeight * (b+1))],
                                        [rectLeft, rectTop - (barHeight * (b+1))]
                                    ]);
                                    bar.filled = true;
                                    bar.stroked = false;
                                    var barColor = new RGBColor();
                                    barColor.red = gradientColors[b][0]; 
                                    barColor.green = gradientColors[b][1]; 
                                    barColor.blue = gradientColors[b][2];
                                    bar.fillColor = barColor;
                                    bar.move(rect, ElementPlacement.PLACEBEFORE);
                                }
                                
                                // Move the transparent rectangle to the front so border is visible
                                rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                            }
                            
                            var verticalCenter = yPos - (rectHeight / 2);
                            var textGroup = posGroup.groupItems.add();
                            var xPos = xOffset + textSpacing;
                            
                            if(includeIDs){
                                var posNumSize = toPointsLocal(parseFloat(listPosNumSizeInput.text)||12, listPosNumSizeUnit.selection.text);
                                var circleDiameter = Math.max(rectHeight * 0.6, posNumSize * 1.5);
                                var circleTopY = yPos - (rectHeight - circleDiameter)/2;
                                
                                var circle = textGroup.pathItems.ellipse(circleTopY, xPos, circleDiameter, circleDiameter);
                                circle.filled = true;
                                circle.stroked = false;
                                var circleColor = new RGBColor();
                                circleColor.red = 18;
                                circleColor.green = 75;
                                circleColor.blue = 142;
                                circle.fillColor = circleColor;
                                
                                var posText = textGroup.textFrames.pointText([0, 0]);
                                posText.contents = globalCount++;
                                posText.textRange.characterAttributes.textFont = posNumFont;
                                posText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listPosNumSizeInput.text)||12, listPosNumSizeUnit.selection.text);
                                
                                var whiteColor = new RGBColor();
                                whiteColor.red = 255;
                                whiteColor.green = 255;
                                whiteColor.blue = 255;
                                posText.textRange.characterAttributes.fillColor = whiteColor;
                                
                                var bounds = posText.geometricBounds;
                                var textWidth = bounds[2] - bounds[0];
                                var textHeight = bounds[1] - bounds[3];
                                posText.position = [xPos + (circleDiameter - textWidth)/2, verticalCenter + textHeight/2];
                                
                                xPos += circleDiameter + textSpacing;
                            }
                            
                            var skuText = textGroup.textFrames.pointText([0, 0]);
                            skuText.contents = row["SKU"];
                            skuText.textRange.characterAttributes.textFont = skuFont;
                            skuText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listSkuSizeInput.text)||12, listSkuSizeUnit.selection.text);
                            var skuBounds = skuText.geometricBounds;
                            var skuHeight = skuBounds[1] - skuBounds[3];
                            skuText.position = [xPos, verticalCenter + skuHeight/2];
                            xPos += 43 + textSpacing;
                            
                            var productText = textGroup.textFrames.pointText([0, 0]);
                            productText.contents = row["Product"];
                            productText.textRange.characterAttributes.textFont = productFont;
                            productText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                            productText.textRange.characterAttributes.leading = 12;
                            var productBounds = productText.geometricBounds;
                            var productHeight = productBounds[1] - productBounds[3];
                            productText.position = [xPos, verticalCenter + productHeight/2];
                            
                            var priceText = textGroup.textFrames.pointText([0, 0]);
                            priceText.contents = row["Price"];
                            priceText.textRange.characterAttributes.textFont = priceFont;
                            priceText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listPriceSizeInput.text)||12, listPriceSizeUnit.selection.text);
                            var priceBounds = priceText.geometricBounds;
                            var priceWidth = priceBounds[2] - priceBounds[0];
                            var priceHeight = priceBounds[1] - priceBounds[3];
                            priceText.position = [xOffset + totalWidth - priceWidth - textSpacing, verticalCenter + priceHeight/2];
                            
                            yPos -= rectHeight;
                        }
                        
                        // Add position title rectangle on the left side of the position group
                        var positionTitle = positionTitles[pos] || "";
                        if(positionTitle) {
                            // Calculate the total height of this position group
                            var groupHeight = rows.length * rectHeight;
                            var groupTopY = yPos + groupHeight;
                            
                            var posTitleWidth = 20;
                            var posTitleHeight = groupHeight;
                            var posTitleX = xOffset - posTitleWidth; // Position to the left of the list view
                            var posTitleY = groupTopY;
                            
                            var posTitleRect = posGroup.pathItems.rectangle(posTitleY, posTitleX, posTitleWidth, posTitleHeight);
                            posTitleRect.stroked = true;
                            posTitleRect.strokeWidth = 0.5;
                            var posTitleStrokeClr = new RGBColor();
                            posTitleStrokeClr.red = 0;
                            posTitleStrokeClr.green = 0;
                            posTitleStrokeClr.blue = 0;
                            posTitleRect.strokeColor = posTitleStrokeClr;
                            posTitleRect.filled = true;
                            var posTitleFillClr = new RGBColor();
                            posTitleFillClr.red = 15;
                            posTitleFillClr.green = 76;
                            posTitleFillClr.blue = 143;
                            posTitleRect.fillColor = posTitleFillClr;
                            
                            // Add rotated position title text
                            var posTitleText = posGroup.textFrames.add();
                            posTitleText.contents = positionTitle;
                            
                            // Use product font settings
                            posTitleText.textRange.characterAttributes.textFont = productFont;
                            posTitleText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                            
                            // Set white text color
                            var whiteTxtClr = new RGBColor();
                            whiteTxtClr.red = 255;
                            whiteTxtClr.green = 255;
                            whiteTxtClr.blue = 255;
                            posTitleText.textRange.characterAttributes.fillColor = whiteTxtClr;
                            
                            // Rotate text 90 degrees clockwise
                            posTitleText.rotate(90);
                            
                            // Center the rotated text
                            var textBounds = posTitleText.geometricBounds;
                            var textWidth = textBounds[2] - textBounds[0];
                            var textHeight = textBounds[1] - textBounds[3];
                            
                            var rectCenterX = posTitleX + (posTitleWidth / 2);
                            var rectCenterY = posTitleY - (posTitleHeight / 2);
                            
                            posTitleText.position = [
                                rectCenterX - (textWidth / 2),
                                rectCenterY + (textHeight / 2)
                            ];
                        }
                        
                        // Add this position group to the array
                        allPositionGroups.push(posGroup);
                        
                        yPos -= groupDivider;
                    }
                    
                    // Group all position groups together into a single "List View" group
                    if(allPositionGroups.length > 0) {
                        var listViewMasterGroup = layer.groupItems.add();
                        listViewMasterGroup.name = "List View";
                        for(var g=0; g<allPositionGroups.length; g++) {
                            allPositionGroups[g].moveToBeginning(listViewMasterGroup);
                        }
                    }
                } catch(e) {
                    alert("Create List View error: "+e);
                }
            };
            
            createListBtn.onClick();
            createListBtn.onClick = originalListClick;
            
        } catch(e) {
            alert("Create Both error: "+e);
        }
    }
});

// === Tools Tab ===
var tabTools = tabs.add("tab", undefined, "Tools");
tabTools.orientation = "column";
tabTools.alignChildren = ["left","top"];
tabTools.spacing = 15;

// About section
var aboutGroup = tabTools.add("group");
aboutGroup.orientation = "column";
aboutGroup.alignChildren = ["left","top"];
aboutGroup.spacing = 5;

var aboutTitle = aboutGroup.add("statictext", undefined, "PlanoMako");
aboutTitle.graphics.font = ScriptUI.newFont(aboutTitle.graphics.font.name, ScriptUI.FontStyle.BOLD, 14);

var aboutText = aboutGroup.add("statictext", undefined, "PlanoMako is a tool for creating planogram layouts from CSV data.\nGenerate front-view and list-view layouts with dynamic features,\ncustom colors, and flexible positioning options.", {multiline: true});
aboutText.preferredSize.width = 700;

// Export CSV button with description
var exportCSVGroup = tabTools.add("group");
exportCSVGroup.orientation = "row";
exportCSVGroup.alignChildren = ["left","center"];
exportCSVGroup.spacing = 10;

var exportCSVBtn = exportCSVGroup.add("button", undefined, "Export CSV");
exportCSVBtn.preferredSize.width = 120;

var exportCSVDesc = exportCSVGroup.add("statictext", undefined, "Export a CSV that you can import later.\nThis will save your data rows, features, and feature colors.", {multiline: true});
exportCSVDesc.preferredSize.width = 550;

// Clone Script button with description
var cloneScriptGroup = tabTools.add("group");
cloneScriptGroup.orientation = "row";
cloneScriptGroup.alignChildren = ["left","center"];
cloneScriptGroup.spacing = 10;

var cloneScriptBtn = cloneScriptGroup.add("button", undefined, "Clone Script");
cloneScriptBtn.preferredSize.width = 120;

var cloneScriptDesc = cloneScriptGroup.add("statictext", undefined, "Save all of your preferences by exporting a new PlanoMako\nscript file with current preferences serving as default values.", {multiline: true});
cloneScriptDesc.preferredSize.width = 550;

// === Buttons (Bottom) ===
var bottom = win.add("group");
bottom.orientation = "row";
bottom.alignment = "fill";

var left = bottom.add("group");
left.orientation = "row";
var clearBtn = left.add("button", undefined, "Clear");
var cancelBtn = left.add("button", undefined, "Cancel");

var right = bottom.add("group");
right.orientation = "row";
right.alignment = "right";
right.add("statictext", undefined, "Create:");
var createFrontBtn = right.add("button", undefined, "Front View");
var createListBtn = right.add("button", undefined, "List View");
var createBothBtn = right.add("button", undefined, "Create Both");

// === Button Handlers ===

// Export CSV Handler
exportCSVBtn.onClick = function(){
    saveRows(); // Save current data first
    
    if(csvData.length === 0){
        alert("No data to export.");
        return;
    }
    
    var f = File.saveDialog("Save CSV file", "*.csv");
    if(!f) return;
    
    // Ensure .csv extension
    var filePath = f.fsName;
    if(filePath.toLowerCase().indexOf(".csv") !== filePath.length - 4){
        f = new File(filePath + ".csv");
    }
    
    f.open("w");
    
    // Write header comment with feature colors metadata
    f.writeln("# PlanoMako Export");
    f.writeln("# Feature Colors:");
    for(var i = 0; i < featureColumns.length; i++){
        var colorValue = "#CCCCCC"; // Default
        if(i < featureColorInputs.length && featureColorInputs[i]){
            colorValue = featureColorInputs[i].text;
        }
        f.writeln("# " + featureColumns[i] + "=" + colorValue);
    }
    f.writeln("#");
    
    // Build header row
    var headers = baseColumns.slice(); // ["Position", "Product", "SKU", "Price"]
    
    // Add "Features" divider column
    headers.push("Features");
    
    // Add all feature columns
    for(var j = 0; j < featureColumns.length; j++){
        headers.push(featureColumns[j]);
    }
    
    // Write header row
    f.writeln(headers.join(","));
    
    // Write data rows
    for(var r = 0; r < csvData.length; r++){
        var row = csvData[r];
        var rowValues = [];
        
        // Add base columns
        for(var b = 0; b < baseColumns.length; b++){
            var val = row[baseColumns[b]] || "";
            // Escape quotes and wrap in quotes if contains comma or quote
            if(val.indexOf(",") >= 0 || val.indexOf('"') >= 0){
                val = '"' + val.replace(/"/g, '""') + '"';
            }
            rowValues.push(val);
        }
        
        // Add empty "Features" divider column
        rowValues.push("");
        
        // Add feature columns
        for(var fc = 0; fc < featureColumns.length; fc++){
            var featureVal = row[featureColumns[fc]] || "";
            rowValues.push(featureVal);
        }
        
        f.writeln(rowValues.join(","));
    }
    
    f.close();
    alert("CSV exported successfully!");
};

// Clone Script Handler
cloneScriptBtn.onClick = function(){
    // Try to get the current script file path
    var currentScript = null;
    
    // Try multiple methods to find the script
    if($.fileName){
        currentScript = File($.fileName);
    }
    
    // If not found, ask user to locate it
    if(!currentScript || !currentScript.exists){
        alert("Please locate the current PlanoMako script file to clone.");
        currentScript = File.openDialog("Select the current PlanoMako script", "*.jsx");
        if(!currentScript) return;
    }
    
    // Ask user where to save the cloned script
    var newScript = File.saveDialog("Save cloned PlanoMako script", "*.jsx");
    if(!newScript) return;
    
    // Ensure .jsx extension
    var scriptPath = newScript.fsName;
    if(scriptPath.toLowerCase().indexOf(".jsx") !== scriptPath.length - 4){
        newScript = new File(scriptPath + ".jsx");
    }
    
    // Read the current script
    currentScript.open("r");
    var scriptContent = currentScript.read();
    currentScript.close();
    
    // Build preferences object with current values
    var currentPrefs = {
        // Front View
        frontWidth: widthInput.text || "360",
        frontWidthUnit: widthUnit.selection ? widthUnit.selection.index : 0,
        frontBgElements: bgElementsDrop.selection ? bgElementsDrop.selection.index : 0,
        frontPosLabels: posLabelsDrop.selection ? posLabelsDrop.selection.index : 0,
        frontPegLineSpacing: pegLineSpacingInput.text || "40",
        frontPegLineSpacingUnit: pegLineSpacingUnit.selection ? pegLineSpacingUnit.selection.index : 0,
        frontShelfHeight: shelfHeightInput.text || "20",
        frontShelfHeightUnit: shelfHeightUnit.selection ? shelfHeightUnit.selection.index : 0,
        frontProdRectSpacing: rectSpacingInput.text || "10",
        frontProdRectSpacingUnit: rectSpacingUnit.selection ? rectSpacingUnit.selection.index : 0,
        frontRectHeight: rectHeightInput.text || "72",
        frontRectHeightUnit: rectHeightUnit.selection ? rectHeightUnit.selection.index : 0,
        frontRectWidth: rectWidthInput.text || "72",
        frontRectWidthUnit: rectWidthUnit.selection ? rectWidthUnit.selection.index : 0,
        frontPosLabelFontSize: posLabelSizeInput.text || "12",
        frontPosLabelFontSizeUnit: posLabelSizeUnit.selection ? posLabelSizeUnit.selection.index : 1,
        frontSkuFontSize: skuSizeInput.text || "10",
        frontSkuFontSizeUnit: skuSizeUnit.selection ? skuSizeUnit.selection.index : 1,
        frontProductFontSize: productSizeInput.text || "12",
        frontProductFontSizeUnit: productSizeUnit.selection ? productSizeUnit.selection.index : 1,
        frontPriceFontSize: priceSizeInput.text || "14",
        frontPriceFontSizeUnit: priceSizeUnit.selection ? priceSizeUnit.selection.index : 1,
        frontTracking: trackingInput.text || "0",
        frontTrackingUnit: trackingUnit.selection ? trackingUnit.selection.index : 0,
        frontLeading: leadingInput.text || "",
        frontLeadingUnit: leadingUnit.selection ? leadingUnit.selection.index : 0,
        frontBorderColor: borderColorInput.text || "#000000",
        
        // List View
        listWidth: listWidthInput.text || "300",
        listWidthUnit: listWidthUnit.selection ? listWidthUnit.selection.index : 0,
        listIdentifiers: listIdDrop.selection ? listIdDrop.selection.index : 0,
        listDividerSpacing: listDividerInput.text || "10",
        listDividerSpacingUnit: listDividerUnit.selection ? listDividerUnit.selection.index : 0,
        listObjectSpacing: listPadInput.text || "10",
        listObjectSpacingUnit: listPadUnit.selection ? listPadUnit.selection.index : 1,
        listTextSpacing: listTextSpaceInput.text || "5",
        listTextSpacingUnit: listTextSpaceUnit.selection ? listTextSpaceUnit.selection.index : 1,
        listPosNumFontSize: listPosNumSizeInput.text || "12",
        listPosNumFontSizeUnit: listPosNumSizeUnit.selection ? listPosNumSizeUnit.selection.index : 1,
        listSkuFontSize: listSkuSizeInput.text || "12",
        listSkuFontSizeUnit: listSkuSizeUnit.selection ? listSkuSizeUnit.selection.index : 1,
        listProductFontSize: listProductSizeInput.text || "12",
        listProductFontSizeUnit: listProductSizeUnit.selection ? listProductSizeUnit.selection.index : 1,
        listPriceFontSize: listPriceSizeInput.text || "12",
        listPriceFontSizeUnit: listPriceSizeUnit.selection ? listPriceSizeUnit.selection.index : 1,
        listTracking: listTrackingInput.text || "0",
        listTrackingUnit: listTrackingUnit.selection ? listTrackingUnit.selection.index : 0,
        listLeading: listLeadingInput.text || "",
        listLeadingUnit: listLeadingUnit.selection ? listLeadingUnit.selection.index : 0,
        listBorderColor: listBorderColorInput.text || "#000000"
    };
    
    // Replace PREFS values in the script
    for(var key in currentPrefs){
        var pattern = new RegExp("("+key+":\\s*)([\"']?)([^,\\n]+)\\2", "g");
        var newValue = currentPrefs[key];
        // If value is a string (non-numeric), wrap in quotes
        if(typeof newValue === "string" && isNaN(parseFloat(newValue))){
            newValue = '"' + newValue + '"';
        }
        scriptContent = scriptContent.replace(pattern, "$1" + newValue);
    }
    
    // Write the cloned script
    newScript.open("w");
    newScript.write(scriptContent);
    newScript.close();
    
    alert("Script cloned successfully with your current preferences as defaults!");
};

clearBtn.onClick = function(){csvData=[];currentPage=1;rebuildTable();};
cancelBtn.onClick = function(){win.close();};
createListBtn.onClick = function(){
    try {
        if(!app.documents.length){alert("Open a document first.");return;}
        saveRows(); if(csvData.length===0){alert("No rows.");return;}

        function toPoints(value, unit){
            if(unit=="inches") return value*72;
            if(unit=="pixels") return value;
            return value;
        }

        // Get List View specific settings
        var totalWidth = toPoints(parseFloat(listWidthInput.text)||300, listWidthUnit.selection.text);
        var pad = listPadUnit.selection.text=="auto"?10:toPoints(parseFloat(listPadInput.text)||10, listPadUnit.selection.text);
        var groupDivider = toPoints(parseFloat(listDividerInput.text)||10, listDividerUnit.selection.text);
        var textSpacing = listTextSpaceUnit.selection.text=="auto"?5:toPoints(parseFloat(listTextSpaceInput.text)||5, listTextSpaceUnit.selection.text);
        // Default font size will be handled individually for each text element
        
        var includeIDs=(listIdDrop.selection.text=="Identifiers");
        
        // Get font settings for each text element
        function getFontByFamilyAndStyle(familyName, preferredStyle) {
            try {
                for(var fi=0; fi<app.textFonts.length; fi++){
                    var f = app.textFonts[fi];
                    if(f.family==familyName && (!preferredStyle || f.style==preferredStyle)) { return f; }
                }
                // Fallback to any style of the family
                for(var fi2=0; fi2<app.textFonts.length; fi2++){
                    if(app.textFonts[fi2].family==familyName){ return app.textFonts[fi2]; }
                }
            } catch(e) {}
            return app.textFonts[0]; // Ultimate fallback
        }
        
        var posNumFont = getFontByFamilyAndStyle(
            listPosNumFontDrop.selection ? listPosNumFontDrop.selection.text : "ITCFranklinGothic LT Pro",
            listPosNumStyleDrop.selection ? listPosNumStyleDrop.selection.text : "CnDm"
        );
        var skuFont = getFontByFamilyAndStyle(
            listSkuFontDrop.selection ? listSkuFontDrop.selection.text : "ITCFranklinGothic LT Pro",
            listSkuStyleDrop.selection ? listSkuStyleDrop.selection.text : "CnDm"
        );
        var productFont = getFontByFamilyAndStyle(
            listProductFontDrop.selection ? listProductFontDrop.selection.text : "ITCFranklinGothic LT Pro",
            listProductStyleDrop.selection ? listProductStyleDrop.selection.text : "CnBk"
        );
        var priceFont = getFontByFamilyAndStyle(
            listPriceFontDrop.selection ? listPriceFontDrop.selection.text : "ITCFranklinGothic LT Pro",
            listPriceStyleDrop.selection ? listPriceStyleDrop.selection.text : "CnBk"
        );
        
        // Set up colors
        var strokeClr = new RGBColor();
        strokeClr.red = 0;
        strokeClr.green = 0;
        strokeClr.blue = 0;

        var doc=app.activeDocument, layer=doc.activeLayer;
        var grouped={};
        for(var i=0;i<csvData.length;i++){
            var pos=csvData[i]["Position"];
            if(!grouped[pos])grouped[pos]=[];
            grouped[pos].push(csvData[i]);
        }

        // Get position titles from UI before closing window
        var positionTitles = {};
        for(var i=0; i<scrollPanel.children.length; i++){
            var child = scrollPanel.children[i];
            if(child.type=="group" && child.children.length>=4){
                var firstChild = child.children[0];
                if(firstChild.type=="statictext" && firstChild.text.indexOf("Position ") === 0 && firstChild.text.indexOf(" label:") > 0){
                    var posMatch = firstChild.text.match(/Position (\d+) label:/);
                    if(posMatch && posMatch[1]){
                        var posNum = posMatch[1];
                        var titleInput = child.children[1];
                        if(titleInput.type=="edittext"){
                            positionTitles[posNum] = titleInput.text;
                        }
                    }
                }
            }
        }

        win.close();
        
        var yPos = 0;
        var globalCount = 1; // Track position count across all groups
                // Use product name font size as base for rectangle height
                var baseHeight = toPoints(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                var rectHeight = baseHeight * 2.5;
        var allPositionGroups = []; // Track all position groups for final grouping
        for(var pos in grouped){
            var rows = grouped[pos];
            var posGroup = layer.groupItems.add();
            posGroup.name = "Position " + pos;
            
            for(var j=0; j<rows.length; j++){
                var row = rows[j];
                
                // Create rectangle first
                var rect = posGroup.pathItems.rectangle(yPos, 0, totalWidth, rectHeight);
                rect.stroked = true;
                rect.strokeWidth = 0.5;
                rect.strokeColor = strokeClr;
                rect.filled = true;
                
                // Get active features for this row dynamically
                var activeFeatures = getActiveFeatures(row);
                var featureCount = activeFeatures.length;
                
                if(featureCount === 0){
                    // No features - white background
                    var whiteFill = new RGBColor();
                    whiteFill.red = 255;
                    whiteFill.green = 255;
                    whiteFill.blue = 255;
                    rect.fillColor = whiteFill;
                } else if(featureCount === 1){
                    // One feature - solid color
                    var featureColorArr = getFeatureColor(activeFeatures[0]);
                    var featureColor = new RGBColor();
                    featureColor.red = featureColorArr[0]; featureColor.green = featureColorArr[1]; featureColor.blue = featureColorArr[2];
                    rect.fillColor = featureColor;
                } else if(featureCount === 2){
                    // Two features - make rect transparent, create two color wedges
                    rect.filled = false;
                    
                    var gradientColors = [];
                    for(var fc=0; fc<activeFeatures.length; fc++){
                        gradientColors.push(getFeatureColor(activeFeatures[fc]));
                    }
                    
                    var rectBounds = rect.geometricBounds;
                    var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                    var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                    var rectLeft = rectBounds[0];
                    var rectTop = rectBounds[1];
                    
                    // Create first color triangle (top-left to bottom-right)
                    var wedge1 = posGroup.pathItems.add();
                    wedge1.setEntirePath([[rectLeft, rectTop], [rectLeft + rectW, rectTop], [rectLeft, rectTop - rectH]]);
                    wedge1.filled = true;
                    wedge1.stroked = false;
                    var color1 = new RGBColor();
                    color1.red = gradientColors[0][0]; color1.green = gradientColors[0][1]; color1.blue = gradientColors[0][2];
                    wedge1.fillColor = color1;
                    wedge1.move(rect, ElementPlacement.PLACEBEFORE);
                    
                    // Create second color triangle (bottom-right)
                    var wedge2 = posGroup.pathItems.add();
                    wedge2.setEntirePath([[rectLeft + rectW, rectTop], [rectLeft + rectW, rectTop - rectH], [rectLeft, rectTop - rectH]]);
                    wedge2.filled = true;
                    wedge2.stroked = false;
                    var color2 = new RGBColor();
                    color2.red = gradientColors[1][0]; color2.green = gradientColors[1][1]; color2.blue = gradientColors[1][2];
                    wedge2.fillColor = color2;
                    wedge2.move(rect, ElementPlacement.PLACEBEFORE);
                    
                    // Move the transparent rectangle to the front so border is visible
                    rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                } else {
                    // Three or more features - make rect transparent, create horizontal bars
                    rect.filled = false;
                    
                    var gradientColors = [];
                    for(var fc=0; fc<activeFeatures.length; fc++){
                        gradientColors.push(getFeatureColor(activeFeatures[fc]));
                    }
                    
                    var rectBounds = rect.geometricBounds;
                    var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                    var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                    var rectLeft = rectBounds[0];
                    var rectTop = rectBounds[1];
                    var barHeight = rectH / featureCount;
                    
                    // Create horizontal bars for each feature
                    for(var b=0; b<featureCount; b++){
                        var bar = posGroup.pathItems.add();
                        bar.setEntirePath([
                            [rectLeft, rectTop - (barHeight * b)],
                            [rectLeft + rectW, rectTop - (barHeight * b)],
                            [rectLeft + rectW, rectTop - (barHeight * (b+1))],
                            [rectLeft, rectTop - (barHeight * (b+1))]
                        ]);
                        bar.filled = true;
                        bar.stroked = false;
                        var barColor = new RGBColor();
                        barColor.red = gradientColors[b][0]; 
                        barColor.green = gradientColors[b][1]; 
                        barColor.blue = gradientColors[b][2];
                        bar.fillColor = barColor;
                        bar.move(rect, ElementPlacement.PLACEBEFORE);
                    }
                    
                    // Move the transparent rectangle to the front so border is visible
                    rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                }
                
                // Calculate vertical center of current rectangle
                var verticalCenter = yPos - (rectHeight / 2);
                
                // Create container for text elements
                var textGroup = posGroup.groupItems.add();
                var xPos = textSpacing;
                
                if(includeIDs){
                    var posNumSize = toPoints(parseFloat(listPosNumSizeInput.text)||12, listPosNumSizeUnit.selection.text);
                    var circleDiameter = Math.max(rectHeight * 0.6, posNumSize * 1.5); // Ensure circle is large enough for the number
                    var circleTopY = yPos - (rectHeight - circleDiameter)/2;
                    
                    // Create and style circle
                    var circle = textGroup.pathItems.ellipse(
                        circleTopY,
                        xPos,
                        circleDiameter,
                        circleDiameter
                    );
                    circle.filled = true;
                    circle.stroked = false;
                    var circleColor = new RGBColor();
                    circleColor.red = 18;
                    circleColor.green = 75;
                    circleColor.blue = 142;
                    circle.fillColor = circleColor;
                    
                    // Create and style number
                    var posText = textGroup.textFrames.pointText([0, 0]);
                    posText.contents = globalCount++;  // Use and increment global counter
                    posText.textRange.characterAttributes.textFont = posNumFont;
                    posText.textRange.characterAttributes.size = toPoints(parseFloat(listPosNumSizeInput.text)||12, listPosNumSizeUnit.selection.text);
                    
                    // Set white color
                    var whiteColor = new RGBColor();
                    whiteColor.red = 255;
                    whiteColor.green = 255;
                    whiteColor.blue = 255;
                    posText.textRange.characterAttributes.fillColor = whiteColor;
                    
                    // Center number in circle
                    var bounds = posText.geometricBounds;
                    var textWidth = bounds[2] - bounds[0];
                    var textHeight = bounds[1] - bounds[3];
                    posText.position = [
                        xPos + (circleDiameter - textWidth)/2,
                        verticalCenter + textHeight/2
                    ];
                    
                    xPos += circleDiameter + textSpacing;
                }
                
                // SKU
                var skuText = textGroup.textFrames.pointText([0, 0]);
                skuText.contents = row["SKU"];
                skuText.textRange.characterAttributes.textFont = skuFont;
                skuText.textRange.characterAttributes.size = toPoints(parseFloat(listSkuSizeInput.text)||12, listSkuSizeUnit.selection.text);
                var skuBounds = skuText.geometricBounds;
                var skuHeight = skuBounds[1] - skuBounds[3];
                skuText.position = [xPos, verticalCenter + skuHeight/2];
                xPos += 43 + textSpacing;  // Added 3 pixels to the spacing
                
                // Product name
                var productText = textGroup.textFrames.pointText([0, 0]);
                productText.contents = row["Product"];
                productText.textRange.characterAttributes.textFont = productFont;
                productText.textRange.characterAttributes.size = toPoints(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                productText.textRange.characterAttributes.leading = 12;
                var productBounds = productText.geometricBounds;
                var productHeight = productBounds[1] - productBounds[3];
                productText.position = [xPos, verticalCenter + productHeight/2];
                
                // Price (right-aligned)
                var priceText = textGroup.textFrames.pointText([0, 0]);
                priceText.contents = row["Price"];
                priceText.textRange.characterAttributes.textFont = priceFont;
                priceText.textRange.characterAttributes.size = toPoints(parseFloat(listPriceSizeInput.text)||12, listPriceSizeUnit.selection.text);
                var priceBounds = priceText.geometricBounds;
                var priceWidth = priceBounds[2] - priceBounds[0];
                var priceHeight = priceBounds[1] - priceBounds[3];
                priceText.position = [
                    totalWidth - priceWidth - textSpacing,
                    verticalCenter + priceHeight/2
                ];
                
                // Move to next row
                yPos -= rectHeight;
            }
            
            // Add position title rectangle on the left side of the position group
            var positionTitle = positionTitles[pos] || "";
            if(positionTitle) {
                // Calculate the total height of this position group
                var groupHeight = rows.length * rectHeight;
                var groupTopY = yPos + groupHeight; // Top of the first rectangle in this group
                
                var posTitleWidth = 20;
                var posTitleHeight = groupHeight;
                var posTitleX = -posTitleWidth; // Position to the left, touching at x=0
                var posTitleY = groupTopY;
                
                var posTitleRect = posGroup.pathItems.rectangle(posTitleY, posTitleX, posTitleWidth, posTitleHeight);
                posTitleRect.stroked = true;
                posTitleRect.strokeWidth = 0.5;
                var posTitleStrokeClr = new RGBColor();
                posTitleStrokeClr.red = 0;
                posTitleStrokeClr.green = 0;
                posTitleStrokeClr.blue = 0;
                posTitleRect.strokeColor = posTitleStrokeClr;
                posTitleRect.filled = true;
                var posTitleFillClr = new RGBColor();
                posTitleFillClr.red = 15;
                posTitleFillClr.green = 76;
                posTitleFillClr.blue = 143;
                posTitleRect.fillColor = posTitleFillClr;
                
                // Add rotated position title text
                var posTitleText = posGroup.textFrames.add();
                posTitleText.contents = positionTitle;
                
                // Use product font settings for the position title
                posTitleText.textRange.characterAttributes.textFont = productFont;
                posTitleText.textRange.characterAttributes.size = toPoints(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                
                // Set white text color
                var whiteTxtClr = new RGBColor();
                whiteTxtClr.red = 255;
                whiteTxtClr.green = 255;
                whiteTxtClr.blue = 255;
                posTitleText.textRange.characterAttributes.fillColor = whiteTxtClr;
                
                // Rotate text 90 degrees clockwise
                posTitleText.rotate(90);
                
                // Center the rotated text in the position title rectangle
                var textBounds = posTitleText.geometricBounds;
                var textWidth = textBounds[2] - textBounds[0];
                var textHeight = textBounds[1] - textBounds[3];
                
                // Calculate center position
                var rectCenterX = posTitleX + (posTitleWidth / 2);
                var rectCenterY = posTitleY - (posTitleHeight / 2);
                
                // Position text so its center aligns with rectangle center
                posTitleText.position = [
                    rectCenterX - (textWidth / 2),
                    rectCenterY + (textHeight / 2)
                ];
            }
            
            // Add this position group to the array
            allPositionGroups.push(posGroup);
            
            // Add group divider space
            yPos -= groupDivider;
        }
        
        // Group all position groups together into a single "List View" group
        if(allPositionGroups.length > 0) {
            var listViewMasterGroup = layer.groupItems.add();
            listViewMasterGroup.name = "List View";
            for(var g=0; g<allPositionGroups.length; g++) {
                allPositionGroups[g].moveToBeginning(listViewMasterGroup);
            }
        }
    } catch(e) {
        alert("Create List View error: "+e);
    }
};

// === Create Front View ===
function createFrontView(closeWindow){
    try{
        if(!app.documents.length){alert("Open a document first.");return;}
        saveRows(); if(csvData.length===0){alert("No rows.");return;}

    function toPoints(value, unit){
        if(unit=="inches") return value*72;
        if(unit=="pixels") return value;
        return value;
    }

    var totalWidth = toPoints(parseFloat(widthInput.text)||800, widthUnit.selection.text);
    var pad = padUnit.selection.text=="auto"?10:toPoints(parseFloat(padInput.text)||10, padUnit.selection.text);
    var groupSpacing = groupSpacingUnit.selection.text=="auto"?10:toPoints(parseFloat(groupSpacingInput.text)||10, groupSpacingUnit.selection.text);
    var textSpacing = textSpaceUnit.selection.text=="auto"?5:toPoints(parseFloat(textSpaceInput.text)||5, textSpaceUnit.selection.text);
    var fontSize = fontSizeUnit.selection.text=="auto"?12:toPoints(parseFloat(fontSizeInput.text)||12, fontSizeUnit.selection.text);
    var tracking = trackingUnit.selection.text=="auto"?0:parseFloat(trackingInput.text)||0;
    var leading = (leadingUnit.selection.text=="auto"||leadingInput.text=="")?null:toPoints(parseFloat(leadingInput.text), leadingUnit.selection.text);
    var borderWidth = borderWidthUnit.selection.text=="auto"?1:toPoints(parseFloat(borderWidthInput.text)||1, borderWidthUnit.selection.text);
    var cornerRadius = cornerRadiusUnit.selection.text=="auto"?0:toPoints(parseFloat(cornerRadiusInput.text)||0, cornerRadiusUnit.selection.text);

    var includeIDs=(idDrop.selection.text=="Identifiers");
    var fontName=fontDrop.selection?fontDrop.selection.text:"Arial";
    var borderStyle=borderStyleDrop.selection.text;
    // Accept hex (#RRGGBB or #RGB) or legacy "R,G,B"
    var bcText = (borderColorInput && borderColorInput.text)? borderColorInput.text : "#000000";
    var rgbArr = null;
    // Try hex first
    rgbArr = hexToRgb(bcText);
    if(!rgbArr){
        // Try comma-separated
        var parts = bcText.split(",");
        if(parts.length===3){
            var pr = parseInt(parts[0],10);
            var pg = parseInt(parts[1],10);
            var pb = parseInt(parts[2],10);
            if(!isNaN(pr)&&!isNaN(pg)&&!isNaN(pb)) rgbArr=[pr,pg,pb];
        }
    }
    if(!rgbArr) rgbArr=[0,0,0];
    var rC=[rgbArr[0]/255, rgbArr[1]/255, rgbArr[2]/255];

    function clamp255(v){ v = parseInt(v,10); if(isNaN(v)) v=0; return Math.max(0, Math.min(255, Math.round(v))); }
    var strokeClr = new RGBColor();
    strokeClr.red = clamp255(rgbArr[0]);
    strokeClr.green = clamp255(rgbArr[1]);
    strokeClr.blue = clamp255(rgbArr[2]);

    // Parse rect fill color (default white)
    var rectFillArr = parseColorInput(rectFillInput.text) || [255,255,255];
    var rectFillClr = new RGBColor(); rectFillClr.red = clamp255(rectFillArr[0]); rectFillClr.green = clamp255(rectFillArr[1]); rectFillClr.blue = clamp255(rectFillArr[2]);

    // Parse circle fill color (default #124B8E)
    var circFillArr = parseColorInput(circFillInput.text) || hexToRgb('#124B8E');
    var circFillClr = new RGBColor(); circFillClr.red = clamp255(circFillArr[0]); circFillClr.green = clamp255(circFillArr[1]); circFillClr.blue = clamp255(circFillArr[2]);

    // Parse circle text color (default white)
    var circTextArr = parseColorInput(circTextInput.text) || [255,255,255];
    var circTextClr = new RGBColor(); circTextClr.red = clamp255(circTextArr[0]); circTextClr.green = clamp255(circTextArr[1]); circTextClr.blue = clamp255(circTextArr[2]);

        // debug alert removed

    var doc=app.activeDocument, layer=doc.activeLayer;
    var grouped={};
    for(var i=0;i<csvData.length;i++){
        var pos=csvData[i]["Position"];
        if(!grouped[pos])grouped[pos]=[];
        grouped[pos].push(csvData[i]);
    }

    // Get position titles from UI before closing window
    var positionTitles = {};
    for(var i=0; i<scrollPanel.children.length; i++){
        var child = scrollPanel.children[i];
        if(child.type=="group" && child.children.length>=4){
            var firstChild = child.children[0];
            if(firstChild.type=="statictext" && firstChild.text.indexOf("Position ") === 0 && firstChild.text.indexOf(" label:") > 0){
                // Extract position number from "Position X label:"
                var posMatch = firstChild.text.match(/Position (\d+) label:/);
                if(posMatch && posMatch[1]){
                    var posNum = posMatch[1];
                    var titleInput = child.children[1];
                    if(titleInput.type=="edittext"){
                        positionTitles[posNum] = titleInput.text;
                    }
                }
            }
        }
    }
    
    if(closeWindow !== false) win.close();
        var yOff=0,idCount=1,rectCount=0;
        var allPositionGroups = []; // Track all position groups for final grouping
        for(var pos in grouped){
            var rows=grouped[pos];
            var group=layer.groupItems.add(); group.name="Position "+pos;
            var rectW=(totalWidth - pad*(rows.length-1))/rows.length;
            var rectH=60;
            
            // Get the position type for this position (Peg or Shelf)
            var positionType = "Shelf"; // default
            if(rows.length > 0 && rows[0]["PositionType"]) {
                positionType = rows[0]["PositionType"];
            }
            
            // Get position title text
            var positionTitle = positionTitles[pos] || "";
            
            // Check if background elements should be included
            var includeBgElements = (bgElementsDrop.selection.text === "Include");
            
            if(includeBgElements) {
                // Create background rectangle for the entire row group (20px padding left/right, 8px top/bottom)
                var bgPaddingHorizontal = 20;
                var bgPaddingVertical = 8;
                var bgWidth = totalWidth + (bgPaddingHorizontal * 2);
                var bgHeight = rectH + (bgPaddingVertical * 2);
                var bgX = -bgPaddingHorizontal;
                var bgY = -yOff + bgPaddingVertical;
                
                var bgRect = group.pathItems.rectangle(bgY, bgX, bgWidth, bgHeight);
                bgRect.stroked = true;
                bgRect.strokeWidth = 0.5;
                var bgStrokeClr = new RGBColor();
                bgStrokeClr.red = 0;
                bgStrokeClr.green = 0;
                bgStrokeClr.blue = 0;
                bgRect.strokeColor = bgStrokeClr;
                bgRect.filled = true;
                var bgFillClr = new RGBColor();
                bgFillClr.red = 255;
                bgFillClr.green = 255;
                bgFillClr.blue = 255;
                bgRect.fillColor = bgFillClr;
                
                // Only create grey rectangles if position type is "Peg"
                if(positionType === "Peg") {
                    var greyColor = new RGBColor();
                    greyColor.red = 189;
                    greyColor.green = 190;
                    greyColor.blue = 192;
                    
                    var greyGroup = group.groupItems.add();
                    var numGreyRects = 4;
                    var greyRectHeight = 4; // 4px tall rectangles
                    var totalGreyHeight = bgHeight * 0.80; // 80% of background height
                    
                    // Calculate spacing between grey rectangles
                    var greySpacing = (totalGreyHeight - (numGreyRects * greyRectHeight)) / (numGreyRects - 1);
                    
                    // Starting Y position (centered vertically in background)
                    var greyStartY = bgY - (bgHeight / 2) + (totalGreyHeight / 2);
                    
                    for(var g = 0; g < numGreyRects; g++){
                        var greyY = greyStartY - (g * (greyRectHeight + greySpacing));
                        var greyRect = greyGroup.pathItems.rectangle(greyY, bgX, bgWidth, greyRectHeight);
                        greyRect.stroked = false;
                        greyRect.filled = true;
                        var gClr = new RGBColor();
                        gClr.red = 189;
                        gClr.green = 190;
                        gClr.blue = 192;
                        greyRect.fillColor = gClr;
                    }
                }
                
                // Create shelf rectangle if position type is "Shelf" (below background, touching)
                if(positionType === "Shelf") {
                    var shelfHeight = 12; // 12px tall shelf rectangle
                    var shelfY = bgY - bgHeight; // Position it directly below the background rectangle
                    
                    var shelfRect = group.pathItems.rectangle(shelfY, bgX, bgWidth, shelfHeight);
                    shelfRect.stroked = true;
                    shelfRect.strokeWidth = 0.5;
                    var shelfStrokeClr = new RGBColor();
                    shelfStrokeClr.red = 0;
                    shelfStrokeClr.green = 0;
                    shelfStrokeClr.blue = 0;
                    shelfRect.strokeColor = shelfStrokeClr;
                    shelfRect.filled = true;
                    var shelfFillClr = new RGBColor();
                    shelfFillClr.red = 216;
                    shelfFillClr.green = 221;
                    shelfFillClr.blue = 227;
                    shelfRect.fillColor = shelfFillClr;
                }
                
                // Create position title rectangle on the left side (touching background)
                var posTitleWidth = 20;
                var posTitleHeight = bgHeight;
                var posTitleX = bgX - posTitleWidth; // Position to the left of background
                var posTitleY = bgY;
                
                var posTitleRect = group.pathItems.rectangle(posTitleY, posTitleX, posTitleWidth, posTitleHeight);
                posTitleRect.stroked = true;
                posTitleRect.strokeWidth = 0.5;
                var posTitleStrokeClr = new RGBColor();
                posTitleStrokeClr.red = 0;
                posTitleStrokeClr.green = 0;
                posTitleStrokeClr.blue = 0;
                posTitleRect.strokeColor = posTitleStrokeClr;
                posTitleRect.filled = true;
                var posTitleFillClr = new RGBColor();
                posTitleFillClr.red = 15;
                posTitleFillClr.green = 76;
                posTitleFillClr.blue = 143;
                posTitleRect.fillColor = posTitleFillClr;
                
                // Add rotated position title text
                if(positionTitle) {
                    var posTitleText = group.textFrames.add();
                    posTitleText.contents = positionTitle;
                    
                    // Get font from Front View settings with CnBk style
                    var titleFont = null;
                    try{
                        var familyName = fontDrop.selection?fontDrop.selection.text:"ITCFranklinGothic LT Pro";
                        var preferredStyle = "CnBk"; // Use CnBk style for position titles
                        for(var fi=0; fi<app.textFonts.length; fi++){
                            var f = app.textFonts[fi];
                            if(f.family==familyName && f.style==preferredStyle) { titleFont = f; break; }
                        }
                        if(!titleFont){
                            for(var fi2=0; fi2<app.textFonts.length; fi2++){
                                if(app.textFonts[fi2].family==familyName){ titleFont = app.textFonts[fi2]; break; }
                            }
                        }
                        if(!titleFont) titleFont = app.textFonts[0];
                    }catch(e){ titleFont = app.textFonts[0]; }
                    
                    posTitleText.textRange.characterAttributes.textFont = titleFont;
                    posTitleText.textRange.characterAttributes.size = fontSize;
                    posTitleText.textRange.characterAttributes.tracking = tracking;
                    if(leading) posTitleText.textRange.characterAttributes.leading = leading;
                    
                    // Set white text color
                    var whiteTxtClr = new RGBColor();
                    whiteTxtClr.red = 255;
                    whiteTxtClr.green = 255;
                    whiteTxtClr.blue = 255;
                    posTitleText.textRange.characterAttributes.fillColor = whiteTxtClr;
                    
                    // Rotate text 90 degrees (clockwise, so tops of letters are on the right)
                    posTitleText.rotate(90);
                    
                    // Center the rotated text in the position title rectangle
                    var textBounds = posTitleText.geometricBounds;
                    var textWidth = textBounds[2] - textBounds[0];
                    var textHeight = textBounds[1] - textBounds[3];
                    
                    // Calculate center position
                    var rectCenterX = posTitleX + (posTitleWidth / 2);
                    var rectCenterY = posTitleY - (posTitleHeight / 2);
                    
                    // Position text so its center aligns with rectangle center
                    posTitleText.position = [
                        rectCenterX - (textWidth / 2),
                        rectCenterY + (textHeight / 2)
                    ];
                }
                
                // Move background rect to the back of the group (behind grey rectangles if they exist)
                bgRect.zOrder(ZOrderMethod.SENDTOBACK);
            }
            
            for(var j=0;j<rows.length;j++){
                var x=j*(rectW+pad); var y=-yOff;
                var rect=group.pathItems.roundedRectangle(y,x,rectW,rectH,cornerRadius,cornerRadius);
                rect.stroked=true; rect.strokeWidth=borderWidth;
                // Assign a fresh RGBColor per rectangle (avoid shared object issues)
                var sc = new RGBColor();
                sc.red = strokeClr.red; sc.green = strokeClr.green; sc.blue = strokeClr.blue;
                rect.strokeColor = sc;
                
                // Get active features for this row dynamically
                var rowData = rows[j];
                var activeFeatures = getActiveFeatures(rowData);
                var featureCount = activeFeatures.length;
                
                if(featureCount === 0){
                    // No features selected - use white/default rect fill
                    rect.filled = true;
                    rect.fillColor = (function(){ var c = new RGBColor(); c.red = rectFillClr.red; c.green = rectFillClr.green; c.blue = rectFillClr.blue; return c; })();
                } else if(featureCount === 1){
                    // One feature - solid color
                    rect.filled = true;
                    var featureColorArr = getFeatureColor(activeFeatures[0]);
                    var featureClr = new RGBColor();
                    featureClr.red = featureColorArr[0]; featureClr.green = featureColorArr[1]; featureClr.blue = featureColorArr[2];
                    rect.fillColor = featureClr;
                } else if(featureCount === 2){
                    // Two features - make rect transparent, create two color wedges
                    rect.filled = false;
                    
                    var gradientColors = [];
                    for(var fc=0; fc<activeFeatures.length; fc++){
                        gradientColors.push(getFeatureColor(activeFeatures[fc]));
                    }
                    
                    var rectBounds = rect.geometricBounds;
                    var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                    var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                    var rectLeft = rectBounds[0];
                    var rectTop = rectBounds[1];
                    
                    // Create first color triangle (top-left to bottom-right)
                    var wedge1 = doc.pathItems.add();
                    wedge1.setEntirePath([[rectLeft, rectTop], [rectLeft + rectW, rectTop], [rectLeft, rectTop - rectH]]);
                    wedge1.filled = true;
                    wedge1.stroked = false;
                    var color1 = new RGBColor();
                    color1.red = gradientColors[0][0]; color1.green = gradientColors[0][1]; color1.blue = gradientColors[0][2];
                    wedge1.fillColor = color1;
                    wedge1.move(rect, ElementPlacement.PLACEBEFORE);
                    
                    // Create second color triangle (bottom-right)
                    var wedge2 = doc.pathItems.add();
                    wedge2.setEntirePath([[rectLeft + rectW, rectTop], [rectLeft + rectW, rectTop - rectH], [rectLeft, rectTop - rectH]]);
                    wedge2.filled = true;
                    wedge2.stroked = false;
                    var color2 = new RGBColor();
                    color2.red = gradientColors[1][0]; color2.green = gradientColors[1][1]; color2.blue = gradientColors[1][2];
                    wedge2.fillColor = color2;
                    wedge2.move(rect, ElementPlacement.PLACEBEFORE);
                    
                    // Move the transparent rectangle to the front so border is visible
                    rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                } else {
                    // Three or more features - make rect transparent, create horizontal bars
                    rect.filled = false;
                    
                    var gradientColors = [];
                    for(var fc=0; fc<activeFeatures.length; fc++){
                        gradientColors.push(getFeatureColor(activeFeatures[fc]));
                    }
                    
                    var rectBounds = rect.geometricBounds;
                    var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                    var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                    var rectLeft = rectBounds[0];
                    var rectTop = rectBounds[1];
                    var barHeight = rectH / featureCount;
                    
                    // Create horizontal bars for each feature
                    for(var b=0; b<featureCount; b++){
                        var bar = doc.pathItems.add();
                        bar.setEntirePath([
                            [rectLeft, rectTop - (barHeight * b)],
                            [rectLeft + rectW, rectTop - (barHeight * b)],
                            [rectLeft + rectW, rectTop - (barHeight * (b+1))],
                            [rectLeft, rectTop - (barHeight * (b+1))]
                        ]);
                        bar.filled = true;
                        bar.stroked = false;
                        var barColor = new RGBColor();
                        barColor.red = gradientColors[b][0]; 
                        barColor.green = gradientColors[b][1]; 
                        barColor.blue = gradientColors[b][2];
                        bar.fillColor = barColor;
                        bar.move(rect, ElementPlacement.PLACEBEFORE);
                    }
                    
                    // Move the transparent rectangle to the front so border is visible
                    rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                }
                
                if(borderStyle=="Dashed") rect.strokeDashes=[6,3];
                else if(borderStyle=="Dotted") rect.strokeDashes=[1,2];
                else rect.strokeDashes=[];

                // Pick a font object by family and preferred style if possible
                var tFont = null;
                try{
                    var preferredStyle = (styleDrop.selection && styleDrop.selection.text)? styleDrop.selection.text : null;
                    var familyName = fontDrop.selection?fontDrop.selection.text:fontName;
                    for(var fi=0; fi<app.textFonts.length; fi++){
                        var f = app.textFonts[fi];
                        if(f.family==familyName && (!preferredStyle || f.style==preferredStyle)) { tFont = f; break; }
                    }
                    if(!tFont){
                        // fallback: find any font with the family
                        for(var fi2=0; fi2<app.textFonts.length; fi2++){
                            if(app.textFonts[fi2].family==familyName){ tFont = app.textFonts[fi2]; break; }
                        }
                    }
                    if(!tFont) tFont = app.textFonts[0];
                }catch(e){ tFont = app.textFonts[0]; }

                // Create circle and SKU text first (without a group)
                var tempCirc = null;
                var tempIdLabel = null;
                if(includeIDs){
                    var circleSize=fontSize*2;
                    tempCirc=group.pathItems.ellipse(y-rectH/2+(circleSize),x+rectW/2-(circleSize/2),circleSize,circleSize);
                    tempCirc.stroked=false; tempCirc.filled=true;
                    tempCirc.fillColor = (function(){ var c = new RGBColor(); c.red = circFillClr.red; c.green = circFillClr.green; c.blue = circFillClr.blue; return c; })();
                    tempIdLabel=group.textFrames.add();
                    tempIdLabel.contents=idCount.toString();
                    tempIdLabel.textRange.characterAttributes.textFont=tFont;
                    tempIdLabel.textRange.characterAttributes.size=fontSize;
                    tempIdLabel.textRange.characterAttributes.tracking=tracking;
                    try{ var tClr = new RGBColor(); tClr.red = circTextClr.red; tClr.green = circTextClr.green; tClr.blue = circTextClr.blue; tempIdLabel.textRange.characterAttributes.fillColor = tClr; }catch(e){}
                    if(leading) tempIdLabel.textRange.characterAttributes.leading=leading;
                    tempIdLabel.left=tempCirc.left+(tempCirc.width/2)-(tempIdLabel.width/2);
                    tempIdLabel.top=tempCirc.top-(tempCirc.height/2)+(tempIdLabel.height/2);
                    idCount++;
                }

                var tempSkuText=group.textFrames.add();
                tempSkuText.contents=rows[j]["SKU"];
                tempSkuText.textRange.characterAttributes.textFont=tFont;
                tempSkuText.textRange.characterAttributes.size=fontSize;
                tempSkuText.textRange.characterAttributes.tracking=tracking;
                if(leading) tempSkuText.textRange.characterAttributes.leading=leading;
                tempSkuText.left=x+rectW/2 - tempSkuText.width/2;
                tempSkuText.top=y - rectH/2 - (includeIDs ? textSpacing : 0);

                // Now group the content as-is
                var innerGroup=group.groupItems.add();
                if(tempCirc) tempCirc.moveToBeginning(innerGroup);
                if(tempIdLabel) tempIdLabel.moveToBeginning(innerGroup);
                tempSkuText.moveToBeginning(innerGroup);

                // Center the group vertically in the rectangle
                var groupBounds = innerGroup.geometricBounds;
                var groupHeight = groupBounds[1] - groupBounds[3];
                var groupTop = groupBounds[1];
                
                // Rectangle bounds
                var rectTop = y;
                var rectBottom = y - rectH;
                var rectCenterY = rectTop - (rectH / 2);
                
                // Calculate where the group's center currently is
                var groupCenterY = groupTop - (groupHeight / 2);
                
                // Adjust to align centers
                var adjustment = rectCenterY - groupCenterY;
                innerGroup.translate(0, adjustment);
                
                rectCount++;
            }
            
            // Use different spacing depending on whether background elements are included
            if(includeBgElements) {
                // When background elements are included, space them so they touch (total background height)
                var bgPaddingVertical = 8;
                var bgHeight = rectH + (bgPaddingVertical * 2);
                
                // If position type is "Shelf", add extra height for the shelf rectangle
                if(positionType === "Shelf") {
                    yOff += bgHeight + 12; // Background + 12px shelf rectangle
                } else {
                    // Peg position: add background height + 12.5px gap
                    yOff += bgHeight + 12.5; // Background rectangle + gap
                }
            } else {
                // Normal spacing with groupSpacing setting
                yOff += rectH + groupSpacing;
            }
            
            // Add this position group to the array
            allPositionGroups.push(group);
        }
        
        // Group all position groups together into a single "Front View" group
        if(allPositionGroups.length > 0) {
            var frontViewMasterGroup = layer.groupItems.add();
            frontViewMasterGroup.name = "Front View";
            for(var g=0; g<allPositionGroups.length; g++) {
                allPositionGroups[g].moveToBeginning(frontViewMasterGroup);
            }
        }
        
        app.redraw();
        // alert(rectCount+" front-view boxes created on layer: "+layer.name);
    } catch(e){
        alert("Create Front View error: "+e);
    }
}

createFrontBtn.onClick = function(){
    createFrontView(true);
};

// === Create Both ===
createBothBtn.onClick = function(){
    try {
        if(!app.documents.length){alert("Open a document first.");return;}
        saveRows(); if(csvData.length===0){alert("No rows.");return;}
        
        // Store a copy of the data
        var savedData = csvData.slice();
        
        // Get position titles from UI BEFORE any operations
        var savedPositionTitles = {};
        for(var i=0; i<scrollPanel.children.length; i++){
            var child = scrollPanel.children[i];
            if(child.type=="group" && child.children.length>=4){
                var firstChild = child.children[0];
                if(firstChild.type=="statictext" && firstChild.text.indexOf("Position ") === 0 && firstChild.text.indexOf(" label:") > 0){
                    var posMatch = firstChild.text.match(/Position (\d+) label:/);
                    if(posMatch && posMatch[1]){
                        var posNum = posMatch[1];
                        var titleInput = child.children[1];
                        if(titleInput.type=="edittext"){
                            savedPositionTitles[posNum] = titleInput.text;
                        }
                    }
                }
            }
        }
        
        // Calculate Front View width for offset
        function toPoints(value, unit){
            if(unit=="inches") return value*72;
            if(unit=="pixels") return value;
            return value;
        }
        var frontViewWidth = toPoints(parseFloat(widthInput.text)||800, widthUnit.selection.text);
        var spacing = 50; // Add 50 pixels spacing between views
        
        // Create Front View (inline logic, no window close)
        try {
            createFrontView(false); // Don't close window
        } catch(e) {
            alert("Create Front View error: "+e);
        }
        
        // Restore data and create List View
        csvData = savedData;
        
        // Modify List View creation to offset horizontally
        var originalListClick = createListBtn.onClick;
        createListBtn.onClick = function(){
            try {
                if(!app.documents.length){alert("Open a document first.");return;}
                saveRows(); if(csvData.length===0){alert("No rows.");return;}

                function toPointsLocal(value, unit){
                    if(unit=="inches") return value*72;
                    if(unit=="pixels") return value;
                    return value;
                }

                // Get List View specific settings
                var totalWidth = toPointsLocal(parseFloat(listWidthInput.text)||300, listWidthUnit.selection.text);
                var pad = listPadUnit.selection.text=="auto"?10:toPointsLocal(parseFloat(listPadInput.text)||10, listPadUnit.selection.text);
                var groupDivider = toPointsLocal(parseFloat(listDividerInput.text)||10, listDividerUnit.selection.text);
                var textSpacing = listTextSpaceUnit.selection.text=="auto"?5:toPointsLocal(parseFloat(listTextSpaceInput.text)||5, listTextSpaceUnit.selection.text);
                
                var includeIDs=(listIdDrop.selection.text=="Identifiers");
                
                // Get font settings for each text element
                function getFontByFamilyAndStyle(familyName, preferredStyle) {
                    try {
                        for(var fi=0; fi<app.textFonts.length; fi++){
                            var f = app.textFonts[fi];
                            if(f.family==familyName && (!preferredStyle || f.style==preferredStyle)) { return f; }
                        }
                        for(var fi2=0; fi2<app.textFonts.length; fi2++){
                            if(app.textFonts[fi2].family==familyName){ return app.textFonts[fi2]; }
                        }
                    } catch(e) {}
                    return app.textFonts[0];
                }
                
                var posNumFont = getFontByFamilyAndStyle(
                    listPosNumFontDrop.selection ? listPosNumFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                    listPosNumStyleDrop.selection ? listPosNumStyleDrop.selection.text : "CnDm"
                );
                var skuFont = getFontByFamilyAndStyle(
                    listSkuFontDrop.selection ? listSkuFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                    listSkuStyleDrop.selection ? listSkuStyleDrop.selection.text : "CnDm"
                );
                var productFont = getFontByFamilyAndStyle(
                    listProductFontDrop.selection ? listProductFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                    listProductStyleDrop.selection ? listProductStyleDrop.selection.text : "CnBk"
                );
                var priceFont = getFontByFamilyAndStyle(
                    listPriceFontDrop.selection ? listPriceFontDrop.selection.text : "ITCFranklinGothic LT Pro",
                    listPriceStyleDrop.selection ? listPriceStyleDrop.selection.text : "CnBk"
                );
                
                var strokeClr = new RGBColor();
                strokeClr.red = 0;
                strokeClr.green = 0;
                strokeClr.blue = 0;

                var doc=app.activeDocument, layer=doc.activeLayer;
                var grouped={};
                for(var i=0;i<csvData.length;i++){
                    var pos=csvData[i]["Position"];
                    if(!grouped[pos])grouped[pos]=[];
                    grouped[pos].push(csvData[i]);
                }

                // Use saved position titles (captured before any window operations)
                var positionTitles = savedPositionTitles;

                win.close();
                
                var xOffset = frontViewWidth + spacing; // Offset List View to the right
                var yPos = 0;
                var globalCount = 1;
                var baseHeight = toPointsLocal(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                var rectHeight = baseHeight * 2.5;
                var allPositionGroups = []; // Track all position groups for final grouping
                
                for(var pos in grouped){
                    var rows = grouped[pos];
                    var posGroup = layer.groupItems.add();
                    posGroup.name = "Position " + pos;
                    
                    for(var j=0; j<rows.length; j++){
                        var row = rows[j];
                        
                        var rect = posGroup.pathItems.rectangle(yPos, xOffset, totalWidth, rectHeight);
                        rect.stroked = true;
                        rect.strokeWidth = 0.5;
                        rect.strokeColor = strokeClr;
                        rect.filled = true;
                        
                        // Get active features for this row dynamically
                        var activeFeatures = getActiveFeatures(row);
                        var featureCount = activeFeatures.length;
                        
                        if(featureCount === 0){
                            // No features - white background
                            var whiteFill = new RGBColor();
                            whiteFill.red = 255;
                            whiteFill.green = 255;
                            whiteFill.blue = 255;
                            rect.fillColor = whiteFill;
                        } else if(featureCount === 1){
                            // One feature - solid color
                            var featureColorArr = getFeatureColor(activeFeatures[0]);
                            var featureColor = new RGBColor();
                            featureColor.red = featureColorArr[0]; featureColor.green = featureColorArr[1]; featureColor.blue = featureColorArr[2];
                            rect.fillColor = featureColor;
                        } else if(featureCount === 2){
                            // Two features - make rect transparent, create two color wedges
                            rect.filled = false;
                            
                            var gradientColors = [];
                            for(var fc=0; fc<activeFeatures.length; fc++){
                                gradientColors.push(getFeatureColor(activeFeatures[fc]));
                            }
                            
                            var rectBounds = rect.geometricBounds;
                            var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                            var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                            var rectLeft = rectBounds[0];
                            var rectTop = rectBounds[1];
                            
                            // Create first color triangle (top-left to bottom-right)
                            var wedge1 = posGroup.pathItems.add();
                            wedge1.setEntirePath([[rectLeft, rectTop], [rectLeft + rectW, rectTop], [rectLeft, rectTop - rectH]]);
                            wedge1.filled = true;
                            wedge1.stroked = false;
                            var color1 = new RGBColor();
                            color1.red = gradientColors[0][0]; color1.green = gradientColors[0][1]; color1.blue = gradientColors[0][2];
                            wedge1.fillColor = color1;
                            wedge1.move(rect, ElementPlacement.PLACEBEFORE);
                            
                            // Create second color triangle (bottom-right)
                            var wedge2 = posGroup.pathItems.add();
                            wedge2.setEntirePath([[rectLeft + rectW, rectTop], [rectLeft + rectW, rectTop - rectH], [rectLeft, rectTop - rectH]]);
                            wedge2.filled = true;
                            wedge2.stroked = false;
                            var color2 = new RGBColor();
                            color2.red = gradientColors[1][0]; color2.green = gradientColors[1][1]; color2.blue = gradientColors[1][2];
                            wedge2.fillColor = color2;
                            wedge2.move(rect, ElementPlacement.PLACEBEFORE);
                            
                            // Move the transparent rectangle to the front so border is visible
                            rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                        } else {
                            // Three or more features - create horizontal bars dynamically
                            rect.filled = false;
                            
                            var gradientColors = [];
                            for(var fc=0; fc<activeFeatures.length; fc++){
                                gradientColors.push(getFeatureColor(activeFeatures[fc]));
                            }
                            
                            var rectBounds = rect.geometricBounds;
                            var rectW = Math.abs(rectBounds[2] - rectBounds[0]);
                            var rectH = Math.abs(rectBounds[3] - rectBounds[1]);
                            var rectLeft = rectBounds[0];
                            var rectTop = rectBounds[1];
                            var barHeight = rectH / featureCount;
                            
                            for(var b=0; b<featureCount; b++){
                                var bar = posGroup.pathItems.add();
                                var barTop = rectTop - (b * barHeight);
                                var barBottom = barTop - barHeight;
                                bar.setEntirePath([
                                    [rectLeft, barTop],
                                    [rectLeft + rectW, barTop],
                                    [rectLeft + rectW, barBottom],
                                    [rectLeft, barBottom]
                                ]);
                                bar.filled = true;
                                bar.stroked = false;
                                var barColor = new RGBColor();
                                barColor.red = gradientColors[b][0];
                                barColor.green = gradientColors[b][1];
                                barColor.blue = gradientColors[b][2];
                                bar.fillColor = barColor;
                                bar.move(rect, ElementPlacement.PLACEBEFORE);
                            }
                            
                            // Move the transparent rectangle to the front so border is visible
                            rect.zOrder(ZOrderMethod.BRINGTOFRONT);
                        }
                        
                        var verticalCenter = yPos - (rectHeight / 2);
                        var textGroup = posGroup.groupItems.add();
                        var xPos = xOffset + textSpacing;
                        
                        if(includeIDs){
                            var posNumSize = toPointsLocal(parseFloat(listPosNumSizeInput.text)||12, listPosNumSizeUnit.selection.text);
                            var circleDiameter = Math.max(rectHeight * 0.6, posNumSize * 1.5);
                            var circleTopY = yPos - (rectHeight - circleDiameter)/2;
                            
                            var circle = textGroup.pathItems.ellipse(circleTopY, xPos, circleDiameter, circleDiameter);
                            circle.filled = true;
                            circle.stroked = false;
                            var circleColor = new RGBColor();
                            circleColor.red = 18;
                            circleColor.green = 75;
                            circleColor.blue = 142;
                            circle.fillColor = circleColor;
                            
                            var posText = textGroup.textFrames.pointText([0, 0]);
                            posText.contents = globalCount++;
                            posText.textRange.characterAttributes.textFont = posNumFont;
                            posText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listPosNumSizeInput.text)||12, listPosNumSizeUnit.selection.text);
                            
                            var whiteColor = new RGBColor();
                            whiteColor.red = 255;
                            whiteColor.green = 255;
                            whiteColor.blue = 255;
                            posText.textRange.characterAttributes.fillColor = whiteColor;
                            
                            var bounds = posText.geometricBounds;
                            var textWidth = bounds[2] - bounds[0];
                            var textHeight = bounds[1] - bounds[3];
                            posText.position = [xPos + (circleDiameter - textWidth)/2, verticalCenter + textHeight/2];
                            
                            xPos += circleDiameter + textSpacing;
                        }
                        
                        var skuText = textGroup.textFrames.pointText([0, 0]);
                        skuText.contents = row["SKU"];
                        skuText.textRange.characterAttributes.textFont = skuFont;
                        skuText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listSkuSizeInput.text)||12, listSkuSizeUnit.selection.text);
                        var skuBounds = skuText.geometricBounds;
                        var skuHeight = skuBounds[1] - skuBounds[3];
                        skuText.position = [xPos, verticalCenter + skuHeight/2];
                        xPos += 43 + textSpacing;
                        
                        var productText = textGroup.textFrames.pointText([0, 0]);
                        productText.contents = row["Product"];
                        productText.textRange.characterAttributes.textFont = productFont;
                        productText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                        productText.textRange.characterAttributes.leading = 12;
                        var productBounds = productText.geometricBounds;
                        var productHeight = productBounds[1] - productBounds[3];
                        productText.position = [xPos, verticalCenter + productHeight/2];
                        
                        var priceText = textGroup.textFrames.pointText([0, 0]);
                        priceText.contents = row["Price"];
                        priceText.textRange.characterAttributes.textFont = priceFont;
                        priceText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listPriceSizeInput.text)||12, listPriceSizeUnit.selection.text);
                        var priceBounds = priceText.geometricBounds;
                        var priceWidth = priceBounds[2] - priceBounds[0];
                        var priceHeight = priceBounds[1] - priceBounds[3];
                        priceText.position = [xOffset + totalWidth - priceWidth - textSpacing, verticalCenter + priceHeight/2];
                        
                        yPos -= rectHeight;
                    }
                    
                    // Add position title rectangle on the left side of the position group
                    var positionTitle = positionTitles[pos] || "";
                    if(positionTitle) {
                        // Calculate the total height of this position group
                        var groupHeight = rows.length * rectHeight;
                        var groupTopY = yPos + groupHeight;
                        
                        var posTitleWidth = 20;
                        var posTitleHeight = groupHeight;
                        var posTitleX = xOffset - posTitleWidth; // Position to the left of the list view
                        var posTitleY = groupTopY;
                        
                        var posTitleRect = posGroup.pathItems.rectangle(posTitleY, posTitleX, posTitleWidth, posTitleHeight);
                        posTitleRect.stroked = true;
                        posTitleRect.strokeWidth = 0.5;
                        var posTitleStrokeClr = new RGBColor();
                        posTitleStrokeClr.red = 0;
                        posTitleStrokeClr.green = 0;
                        posTitleStrokeClr.blue = 0;
                        posTitleRect.strokeColor = posTitleStrokeClr;
                        posTitleRect.filled = true;
                        var posTitleFillClr = new RGBColor();
                        posTitleFillClr.red = 15;
                        posTitleFillClr.green = 76;
                        posTitleFillClr.blue = 143;
                        posTitleRect.fillColor = posTitleFillClr;
                        
                        // Add rotated position title text
                        var posTitleText = posGroup.textFrames.add();
                        posTitleText.contents = positionTitle;
                        
                        // Use product font settings
                        posTitleText.textRange.characterAttributes.textFont = productFont;
                        posTitleText.textRange.characterAttributes.size = toPointsLocal(parseFloat(listProductSizeInput.text)||12, listProductSizeUnit.selection.text);
                        
                        // Set white text color
                        var whiteTxtClr = new RGBColor();
                        whiteTxtClr.red = 255;
                        whiteTxtClr.green = 255;
                        whiteTxtClr.blue = 255;
                        posTitleText.textRange.characterAttributes.fillColor = whiteTxtClr;
                        
                        // Rotate text 90 degrees clockwise
                        posTitleText.rotate(90);
                        
                        // Center the rotated text
                        var textBounds = posTitleText.geometricBounds;
                        var textWidth = textBounds[2] - textBounds[0];
                        var textHeight = textBounds[1] - textBounds[3];
                        
                        var rectCenterX = posTitleX + (posTitleWidth / 2);
                        var rectCenterY = posTitleY - (posTitleHeight / 2);
                        
                        posTitleText.position = [
                            rectCenterX - (textWidth / 2),
                            rectCenterY + (textHeight / 2)
                        ];
                    }
                    
                    // Add this position group to the array
                    allPositionGroups.push(posGroup);
                    
                    yPos -= groupDivider;
                }
                
                // Group all position groups together into a single "List View" group
                if(allPositionGroups.length > 0) {
                    var listViewMasterGroup = layer.groupItems.add();
                    listViewMasterGroup.name = "List View";
                    for(var g=0; g<allPositionGroups.length; g++) {
                        allPositionGroups[g].moveToBeginning(listViewMasterGroup);
                    }
                }
            } catch(e) {
                alert("Create List View error: "+e);
            }
        };
        
        createListBtn.onClick();
        
        // Restore original handler
        createListBtn.onClick = originalListClick;
        
        // Align the bottoms of Front View and List View groups
        try {
            var doc = app.activeDocument;
            var layer = doc.activeLayer;
            var frontViewGroup = null;
            var listViewGroup = null;
            
            // Find the Front View and List View master groups
            for(var i=0; i<layer.groupItems.length; i++){
                if(layer.groupItems[i].name === "Front View"){
                    frontViewGroup = layer.groupItems[i];
                }
                if(layer.groupItems[i].name === "List View"){
                    listViewGroup = layer.groupItems[i];
                }
            }
            
            // Align bottoms if both groups exist
            if(frontViewGroup && listViewGroup){
                var frontBounds = frontViewGroup.geometricBounds;
                var listBounds = listViewGroup.geometricBounds;
                
                // geometricBounds: [left, top, right, bottom]
                var frontBottom = frontBounds[3];
                var listBottom = listBounds[3];
                
                // Calculate vertical adjustment needed to align bottoms
                var verticalAdjustment = listBottom - frontBottom;
                
                // Move Front View group vertically to align bottoms
                frontViewGroup.translate(0, verticalAdjustment);
            }
        } catch(e) {
            // Alignment failed but don't show error - views are still created
        }
        
    } catch(e) {
        alert("Create Both error: "+e);
    }
};

win.show();
