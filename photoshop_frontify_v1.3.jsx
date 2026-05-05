#target photoshop
// photoshop_frontify.jsx — v1.3

(function () {
    var doc = app.activeDocument;
    var savedUnits     = app.preferences.rulerUnits;
    var savedTypeUnits = app.preferences.typeUnits;
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits  = TypeUnits.PIXELS;   // keep sizes in document px throughout
    var ptToPx = doc.resolution / 72;

    // Collect artboards first — prevents doc.layers shifting during the loop
    var artboards = [];
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].typename === "LayerSet" && isArtboard(doc.layers[i])) {
            artboards.push(doc.layers[i]);
        }
    }
    for (var i = 0; i < artboards.length; i++) {
        processArtboard(artboards[i]);
    }

    app.preferences.rulerUnits = savedUnits;
    app.preferences.typeUnits  = savedTypeUnits;

    // Strip any existing _frontify suffix before adding a new one
    var baseName = doc.name.replace(/_frontify/gi, "").replace(/\.psd$/i, "");
    var newFile  = new File(doc.fullName.parent + "/" + baseName + "_frontify.psd");
    var saveOpts = new PhotoshopSaveOptions();
    saveOpts.embedColorProfile = true;
    saveOpts.alphaChannels     = true;
    saveOpts.layers            = true;
    doc.saveAs(newFile, saveOpts, false);

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
        var textLayers = [];
        gatherTextLayers(artboard.layers, textLayers);
        for (var i = 0; i < textLayers.length; i++) {
            splitTextLayer(textLayers[i]);
        }

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

    // ── text splitting ───────────────────────────────────────────────

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
        var ti    = layer.textItem;
        var lines = ti.contents.split('\r');
        if (lines.length <= 1) return;

        var bounds  = layer.bounds;
        var boundsT = bounds[1].value;
        var boundsB = bounds[3].value;

        // Count non-empty lines; bail if nothing to split
        var nonEmpty = 0;
        for (var k = 0; k < lines.length; k++) { if (lines[k].length > 0) nonEmpty++; }
        if (nonEmpty <= 1) return;   // nothing meaningful to split

        // lineHeightPx: per-paragraph stride (original total height ÷ paragraph count).
        // Used as the minimum step between duplicates to preserve leading.
        var boundsH      = boundsB - boundsT;
        var lineHeightPx = boundsH / nonEmpty;

        // Skip layers where text is too small to be worth splitting
        if (lineHeightPx / 1.2 < 16) return;

        var styleRanges = getStyleRanges(layer);

        // Determine the font for each non-empty paragraph.
        // If every paragraph shares the same font there is nothing to split —
        // Frontify can handle the layer as-is with uniform styling.
        var defaultFont = "";
        try { defaultFont = ti.font; } catch (e) {}
        var paraFonts = [], tempOff = 0;
        for (var p = 0; p < lines.length; p++) {
            if (lines[p].length > 0) {
                var psr = firstStyleForLine(styleRanges, tempOff, tempOff + lines[p].length);
                paraFonts.push((psr && psr.font) ? psr.font : defaultFont);
            }
            tempOff += lines[p].length + 1;
        }
        var needsSplit = false;
        for (var pi = 1; pi < paraFonts.length; pi++) {
            if (paraFonts[pi] !== paraFonts[0]) { needsSplit = true; break; }
        }
        if (!needsSplit) return;

        // Probe: measure the per-line leading STRIDE, not just the glyph height.
        // stride = height("|\r|") - height("|") = baseline-to-baseline distance,
        // which is exactly what each \r spacer advances and what translate must match.
        var singleLineH = lineHeightPx; // safe fallback
        try {
            var probe1 = layer.duplicate(layer, ElementPlacement.PLACEBEFORE);
            probe1.textItem.contents = "|";
            var h1 = probe1.bounds[3].value - probe1.bounds[1].value;
            probe1.remove();

            var probe2 = layer.duplicate(layer, ElementPlacement.PLACEBEFORE);
            probe2.textItem.contents = "|\r|";
            var h2 = probe2.bounds[3].value - probe2.bounds[1].value;
            probe2.remove();

            var stride = h2 - h1;
            if (stride > 0) singleLineH = stride;
        } catch (e) {}

        var charOffset      = 0;
        var totalLinesSoFar = 0; // running visual-line count of all previous paragraphs

        for (var i = 0; i < lines.length; i++) {
            var lineText = lines[i];
            var lineLen  = lineText.length;

            if (lineLen > 0) {
                var sr = firstStyleForLine(styleRanges, charOffset, charOffset + lineLen);

                var dup = layer.duplicate(layer, ElementPlacement.PLACEBEFORE);
                var dti = dup.textItem;

                try { dti.contents = lineText; } catch (e) {}
                if (sr && sr.font)  { try { dti.font  = sr.font;  } catch (e) {} }
                if (sr && sr.color) { try { dti.color = sr.color; } catch (e) {} }
                try { dti.hyphenation = false; } catch (e) {}

                // Shift the duplicate down by the total visual-line height of all
                // preceding paragraphs. Each line = singleLineH px (measured by probe).
                // This places the content at exactly the same Y as in the original box.
                if (totalLinesSoFar > 0) {
                    try { dup.translate(0, totalLinesSoFar * singleLineH); } catch (e) {}
                }

                // Count how many visual lines THIS paragraph wraps to and accumulate.
                var dupH = 0;
                try {
                    dupH = dup.bounds[3].value - dup.bounds[1].value;
                    var linesInPara = Math.round(dupH / singleLineH);
                    if (linesInPara < 1) linesInPara = 1;
                    totalLinesSoFar += linesInPara;
                } catch (e) { totalLinesSoFar++; }

                // Shrink-to-fit: set autoSizePolicy to auto-height so the text box
                // resizes to hug its content (equivalent to the Properties panel
                // "Resize to Fit Content" button).
                selectLayer(dup);
                try {
                    var stfDesc = new ActionDescriptor();
                    var stfRef  = new ActionReference();
                    stfRef.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
                    stfDesc.putReference(charIDToTypeID("null"), stfRef);
                    var stfTxt = new ActionDescriptor();
                    stfTxt.putEnumerated(
                        stringIDToTypeID("autoSizePolicy"),
                        stringIDToTypeID("autoSizePolicy"),
                        stringIDToTypeID("autoSizeHeight")
                    );
                    stfDesc.putObject(charIDToTypeID("T   "), charIDToTypeID("TxLr"), stfTxt);
                    executeAction(charIDToTypeID("setd"), stfDesc, DialogModes.NO);
                } catch (e) {}
            }
            charOffset += lineLen + 1;
        }
        layer.remove();
    }

    // Three-method approach — whichever one finds style ranges first wins.
    function getStyleRanges(layer) {
        try {
            selectLayer(layer);
            var ref = new ActionReference();
            ref.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
            var txtDesc = executeActionGet(ref).getObjectValue(charIDToTypeID("Txt "));

            var ranges   = [];
            var KEY_tsr  = stringIDToTypeID("textStyleRange");
            var KEY_ts   = stringIDToTypeID("textStyle");

            // ── Method 1: getList by stringID ──────────────────────
            try {
                var list = txtDesc.getList(KEY_tsr);
                for (var i = 0; i < list.count; i++) {
                    var r = extractRange(list.getObjectValue(i));
                    if (r) ranges.push(r);
                }
            } catch (e) {}

            // ── Method 2: putIndex on a layer reference ────────────
            // Bypasses the known count=0 bug in some PS versions
            if (ranges.length === 0) {
                for (var idx = 0; idx < 200; idx++) {
                    try {
                        var idxRef = new ActionReference();
                        idxRef.putIndex(KEY_tsr, idx);
                        idxRef.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
                        var r = extractRange(executeActionGet(idxRef));
                        if (r) ranges.push(r);
                    } catch (e) { break; }
                }
            }

            // ── Method 3: enumerate ALL list-type keys in txtDesc ──
            // Finds the real key regardless of its ID — necessary when
            // stringIDToTypeID("textStyleRange") maps to the wrong ID
            // in this version of Photoshop.
            if (ranges.length === 0) {
                try {
                    for (var ki = 0; ki < txtDesc.count; ki++) {
                        var keyID = txtDesc.getKey(ki);
                        try {
                            if (txtDesc.getType(keyID) !== DescValueType.LISTTYPE) continue;
                            var cList = txtDesc.getList(keyID);
                            if (cList.count === 0) continue;
                            // Verify this list contains textStyle items
                            var firstObj = cList.getObjectValue(0);
                            if (!firstObj.hasKey(KEY_ts)) continue;
                            // Found the right list
                            for (var j = 0; j < cList.count; j++) {
                                var r = extractRange(cList.getObjectValue(j));
                                if (r) ranges.push(r);
                            }
                            break;
                        } catch (e2) {}
                    }
                } catch (e) {}
            }

            return ranges;
        } catch (e) { return []; }
    }

    function extractRange(rd) {
        try {
            var sd = rd.getObjectValue(stringIDToTypeID("textStyle"));

            // fontPostScriptName includes the weight suffix (e.g. "VWHead-Bold").
            // fontName is often just the base family ("VWHead") — use as fallback only.
            var font = "";
            try { font = sd.getString(stringIDToTypeID("fontPostScriptName")); } catch (e) {}
            if (!font) { try { font = sd.getString(stringIDToTypeID("fontName")); } catch (e) {} }

            // Action manager stores size as typographic points (1pt = 1/72 in).
            // Multiply by ptToPx to convert to document pixels, matching what
            // ti.size returns when typeUnits = PIXELS.
            // Store size in points (as the action manager holds it).
            // When assigned to ti.size, PS converts pt→px using document resolution —
            // tagging as "px" would make PS treat the raw number as literal pixels.
            var size = null;
            try { size = new UnitValue(sd.getUnitDoubleValue(charIDToTypeID("Sz  ")), "pt"); } catch (e) {}

            var color = null;
            try {
                var cd = sd.getObjectValue(charIDToTypeID("Clr "));
                color = new SolidColor();
                try { color.rgb.red   = cd.getUnitDoubleValue(charIDToTypeID("Rd  ")); } catch (e) { color.rgb.red   = cd.getDoubleValue(charIDToTypeID("Rd  ")); }
                try { color.rgb.green = cd.getUnitDoubleValue(charIDToTypeID("Grn ")); } catch (e) { color.rgb.green = cd.getDoubleValue(charIDToTypeID("Grn ")); }
                try { color.rgb.blue  = cd.getUnitDoubleValue(charIDToTypeID("Bl  ")); } catch (e) { color.rgb.blue  = cd.getDoubleValue(charIDToTypeID("Bl  ")); }
            } catch (e) {}

            var from = 0, to = 9999;
            try { from = rd.getInteger(charIDToTypeID("From")); } catch (e) {}
            // "To  " charID fails in some PS versions; try the stringID form as well
            try { to = rd.getInteger(charIDToTypeID("To  ")); } catch (e) {}
            if (to === 9999) { try { to = rd.getInteger(stringIDToTypeID("to")); } catch (e) {} }

            // Leading: stored in pt; null means auto-leading (leave PS to handle it).
            var leading = null;
            try {
                var autoLead = false;
                try { autoLead = sd.getBoolean(stringIDToTypeID("autoLeading")); } catch (e) {}
                if (!autoLead) {
                    leading = new UnitValue(sd.getUnitDoubleValue(charIDToTypeID("Ldng")), "pt");
                }
            } catch (e) {}

            return { from: from, to: to, font: font, size: size, color: color, leading: leading };
        } catch (e) { return null; }
    }

    function firstStyleForLine(ranges, from, to) {
        for (var i = 0; i < ranges.length; i++) {
            if (ranges[i].to > from && ranges[i].from <= to) return ranges[i];
        }
        return null;
    }

    function selectLayer(layer) {
        var desc = new ActionDescriptor();
        var ref  = new ActionReference();
        ref.putIdentifier(charIDToTypeID("Lyr "), layer.id);
        desc.putReference(charIDToTypeID("null"), ref);
        desc.putBoolean(charIDToTypeID("MkVs"), false);
        executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
    }

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
