#target photoshop
// photoshop_frontify.jsx — v1.1

(function () {
    var doc = app.activeDocument;
    var savedUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    var ptToPx = doc.resolution / 72;
    var processed = 0;

    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        if (layer.typename === "LayerSet" && isArtboard(layer)) {
            processArtboard(layer);
            processed++;
        }
    }

    app.preferences.rulerUnits = savedUnits;

    // Save As with "_frontify" appended before the .psd extension
    var origPath = doc.fullName;
    var origName = doc.name;
    var baseName = origName.replace(/\.psd$/i, "");
    var newFile = new File(origPath.parent + "/" + baseName + "_frontify.psd");
    var saveOpts = new PhotoshopSaveOptions();
    saveOpts.embedColorProfile = true;
    saveOpts.alphaChannels = true;
    saveOpts.layers = true;
    doc.saveAs(newFile, saveOpts, false);

    // Close the original (now saved as _frontify)
    doc.close(SaveOptions.DONOTSAVECHANGES);

    // Confirm
    alert("Successfully Frontified!");

    // ---------------------------------------------------------------

    function isArtboard(layer) {
        try {
            var ref = new ActionReference();
            ref.putIdentifier(charIDToTypeID("Lyr "), layer.id);
            return executeActionGet(ref).hasKey(stringIDToTypeID("artboard"));
        } catch (e) { return false; }
    }

    function processArtboard(artboard) {
        // Step 1: split mixed-format text layers
        var textLayers = [];
        gatherTextLayers(artboard.layers, textLayers);
        for (var i = 0; i < textLayers.length; i++) {
            splitTextLayer(textLayers[i]);
        }

        // Step 2: rasterize bg group and create bild layer
        var bgLayer = findLayerByName(artboard.layers, "bg");
        if (!bgLayer) {
            alert("Artboard '" + artboard.name + "': no layer named 'bg'. Skipping.");
            return;
        }

        doc.activeLayer = bgLayer;
        rasterizeActiveLayer();
        bgLayer = doc.activeLayer;

        doc.selection.selectAll();
        executeAction(charIDToTypeID("copy"), undefined, DialogModes.NO);
        executeAction(charIDToTypeID("past"), undefined, DialogModes.NO);

        var bildLayer = doc.activeLayer;
        bildLayer.name = "bild";
        doc.selection.deselect();

        bildLayer.move(artboard, ElementPlacement.PLACEATEND);
        bgLayer.remove();
    }

    // --- text splitting ---

    function gatherTextLayers(layers, out) {
        for (var i = 0; i < layers.length; i++) {
            var l = layers[i];
            if (l.typename === "ArtLayer" && l.kind === LayerKind.TEXT) {
                out.push(l);
            } else if (l.typename === "LayerSet") {
                gatherTextLayers(l.layers, out);
            }
        }
    }

    function splitTextLayer(layer) {
        var data = readTextData(layer);
        if (!data || data.styleRanges.length <= 1) return;

        var lines = data.text.split('\r');
        var bounds = layer.bounds;
        var originX = bounds[0].value;
        var originY = bounds[1].value;

        var fs = data.styleRanges[0];
        var lineHeightPx = (fs.leading > 0 ? fs.leading : fs.fontSize * 1.2) * ptToPx;

        var charOffset = 0;
        var lineIndex  = 0;

        for (var i = 0; i < lines.length; i++) {
            var lineText = lines[i];
            if (lineText.length > 0) {
                var style     = firstStyleInRange(data.styleRanges, charOffset, charOffset + lineText.length);
                var baselineY = originY + lineIndex * lineHeightPx + style.fontSize * ptToPx;
                makeTextLayer(lineText, style, originX, baselineY, layer);
                lineIndex++;
            }
            charOffset += lineText.length + 1;
        }

        layer.remove();
    }

    function readTextData(layer) {
        try {
            var ref = new ActionReference();
            ref.putIdentifier(charIDToTypeID("Lyr "), layer.id);
            var txtDesc   = executeActionGet(ref).getObjectValue(charIDToTypeID("Txt "));
            var rawText   = txtDesc.getString(charIDToTypeID("Txt "));
            var rangeList = txtDesc.getList(stringIDToTypeID("textStyleRange"));

            var ranges = [];
            for (var i = 0; i < rangeList.count; i++) {
                var rd = rangeList.getObjectValue(i);
                var sd = rd.getObjectValue(stringIDToTypeID("textStyle"));

                var fontSize = sd.hasKey(charIDToTypeID("Sz  "))
                    ? sd.getUnitDoubleValue(charIDToTypeID("Sz  ")) : 12;

                var leading = 0;
                if (sd.hasKey(stringIDToTypeID("leading"))) {
                    try { leading = sd.getUnitDoubleValue(stringIDToTypeID("leading")); } catch (e) {}
                }
                if (leading <= 0) leading = fontSize * 1.2;

                var color = new SolidColor();
                color.rgb.red = 0; color.rgb.green = 0; color.rgb.blue = 0;
                if (sd.hasKey(charIDToTypeID("Clr "))) {
                    try {
                        var cd = sd.getObjectValue(charIDToTypeID("Clr "));
                        color.rgb.red   = readChannel(cd, charIDToTypeID("Rd  "));
                        color.rgb.green = readChannel(cd, charIDToTypeID("Grn "));
                        color.rgb.blue  = readChannel(cd, charIDToTypeID("Bl  "));
                    } catch (e) {}
                }

                ranges.push({
                    from:     rd.getInteger(charIDToTypeID("From")),
                    to:       rd.getInteger(charIDToTypeID("To  ")),
                    font:     sd.hasKey(stringIDToTypeID("fontName")) ? sd.getString(stringIDToTypeID("fontName")) : "",
                    fontSize: fontSize,
                    leading:  leading,
                    color:    color
                });
            }
            return { text: rawText, styleRanges: ranges };
        } catch (e) { return null; }
    }

    function readChannel(cd, key) {
        try { return cd.getUnitDoubleValue(key); } catch (e) {}
        try { return cd.getDoubleValue(key);     } catch (e) {}
        return 0;
    }

    function firstStyleInRange(ranges, from, to) {
        for (var i = 0; i < ranges.length; i++) {
            if (ranges[i].to > from && ranges[i].from <= to) return ranges[i];
        }
        return ranges[0];
    }

    function makeTextLayer(text, style, x, baselineY, refLayer) {
        var newLayer = doc.artLayers.add();
        newLayer.kind = LayerKind.TEXT;
        var ti = newLayer.textItem;
        ti.contents = text;
        ti.position = [x, baselineY];
        try { ti.font  = style.font; }                          catch (e) {}
        try { ti.size  = new UnitValue(style.fontSize, "pt"); } catch (e) {}
        try { ti.color = style.color; }                         catch (e) {}
        newLayer.move(refLayer, ElementPlacement.PLACEBEFORE);
    }

    // --- shared helpers ---

    function findLayerByName(layers, name) {
        for (var i = 0; i < layers.length; i++) {
            if (layers[i].name === name) return layers[i];
        }
        return null;
    }

    function rasterizeActiveLayer() {
        executeAction(charIDToTypeID("Mrg2"), undefined, DialogModes.NO);
    }

})();
