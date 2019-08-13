const OUTPUT_RESOLUTION = 200;
const OUTPUT_NAME = "Output";
const DOC_LOCATION = activeDocument.path;

// Save original units
var originalUnits = app.preferences.rulerUnits;
app.preferences.rulerUnits = Units.INCHES;

// Creates a temporary working copy of the document
activeDocument.duplicate();

// Adds an adjustment layer
activeDocument.activeLayer = activeDocument.layers[0];
addAdjustmentLayer();

// Moves every layer to a folder and merges it
var mergeGroup = activeDocument.layerSets.add();
while (activeDocument.layers.length > 1) {
    var moveLayer = activeDocument.layers[activeDocument.layers.length - 1];
    if (moveLayer.layerType == "ArtLayer") {
        moveLayer.move(mergeGroup, ElementPlacement.INSIDE);
    }
    else {
        moveLayerSet(moveLayer, mergeGroup);
    }
}
activeDocument.layerSets[0].merge();

// Adds margins
transformAndResize();

// Saves to a pdf file and closes the temp file
saveToPDF();
activeDocument.close(SaveOptions.DONOTSAVECHANGES);

// Reset the units
app.preferences.rulerUnits = originalUnits;

// Send completion alert
alert("Campbook Page Processing: Success!");

// ----------------HELPER FUNCTIONS-------------------

// Adds margins by shrinking the image
function transformAndResize() {
    var height = 11;
    var newSize = (100 / height) * 10;
    activeDocument.activeLayer.resize(newSize, newSize, AnchorPosition.MIDDLECENTER);
    activeDocument.resizeImage(new UnitValue(8.5, "in"), undefined, OUTPUT_RESOLUTION, ResampleMethod.AUTOMATIC);
}

// Saves to a pdf file
function saveToPDF() {
    var pdf = new File(DOC_LOCATION + "/" + OUTPUT_NAME + ".pdf");
    activeDocument.saveAs(pdf, new PDFSaveOptions(), true, Extension.LOWERCASE);
}

// Runs the default photoshop action to add an adjustment layer
function addAdjustmentLayer() {
    var mainAction = new ActionDescriptor();

    var adjRef = new ActionReference();
    adjRef.putClass(charIDToTypeID("AdjL"));
    mainAction.putReference(charIDToTypeID("null"), adjRef);
    
    mainAction.putObject(
        charIDToTypeID("Usng"), 
        charIDToTypeID("AdjL"), 
        createTypeAction());
    
    executeAction(
        charIDToTypeID("Mk  "), 
        mainAction, 
        DialogModes.NO);
}

function createTypeAction() {
    var typeAct = new ActionDescriptor();
    typeAct.putObject(
        charIDToTypeID("Type"),
        charIDToTypeID("BanW"),
        createColorActionDesc());
    return typeAct;
}

function createColorActionDesc() {
    var color1 = new ActionDescriptor();
    color1.putEnumerated(
        stringIDToTypeID("presetKind"),
        stringIDToTypeID("presetKindType"),
        stringIDToTypeID("presetKindDefault"));
    
    color1.putInteger(charIDToTypeID("Rd  "), 40);
    color1.putInteger(charIDToTypeID("Yllw"), 60);
    color1.putInteger(charIDToTypeID("Grn "), 40);
    color1.putInteger(charIDToTypeID("Cyn "), 60);
    color1.putInteger(charIDToTypeID("Bl  "), 20);
    color1.putInteger(charIDToTypeID("Mgnt"), 80);
    
    color1.putBoolean(stringIDToTypeID("useTint"), false);
    color1.putObject(
        stringIDToTypeID("tintColor"), 
        charIDToTypeID("RGBC"), 
        createColorActionDesc2());

    return color1;
}

function createColorActionDesc2() {
    var color2 = new ActionDescriptor();

    color2.putDouble(charIDToTypeID("Rd  "), 225.000458);
    color2.putDouble(charIDToTypeID("Grn "), 211.000671);
    color2.putDouble(charIDToTypeID("Bl  "), 179.001160);
    
    return color2;
}

// Moves a layerSet (bc photoshop is dumb and won't let me do it with its move function)
function moveLayerSet(fromLayer, toLayer){
    var actDesc = new ActionDescriptor();
    actDesc.putReference(charIDToTypeID("null"), getSourceRef(fromLayer));  
    actDesc.putReference(charIDToTypeID("T   "), getDestinationRef(toLayer));  
    actDesc.putBoolean(charIDToTypeID("Adjs"), false);  
    actDesc.putInteger(charIDToTypeID("Vrsn"), 5);  
    executeAction(charIDToTypeID("move"), actDesc, DialogModes.NO);  
}  

function getSourceRef(fromLayer) {
    sourceRef = new ActionReference();  
    sourceRef.putName(charIDToTypeID("Lyr "), fromLayer.name);  
    return sourceRef;
}

function getDestinationRef(toLayer) {
    var indexRef = new ActionReference();
    indexRef.putName(charIDToTypeID("Lyr "), toLayer.name);  
    var layerIndex = executeActionGet(indexRef).getInteger(stringIDToTypeID('itemIndex'));  
    var destinationRef = new ActionReference();  
    destinationRef.putIndex(charIDToTypeID("Lyr "), layerIndex-1);  
    return destinationRef;
}