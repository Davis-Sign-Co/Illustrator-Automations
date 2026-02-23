/**
 * ThruCut & KissCut Overcut Generator v7
 * 
 * 1. Decomposes "Thru-cut" filled rectangles into cut lines with overcuts
 * 2. Moves "Kiss-cut" filled shapes to a dedicated KISS layer
 * 3. Generates registration marks (REG_SF spot color) around the cut area
 *    - If Thru-cut exists: reg marks relative to thru-cut lines
 *    - If only Kiss-cut: reg marks relative to kiss-cut shapes
 * 4. Places an alignment star at the top-right
 * 5. Adds file name labels top and bottom
 * 6. Resizes artboard to fit all artwork
 * 
 * Layer order (top to bottom): REG → CUT → KISS → Layer 1 (labels)
 *   (CUT and KISS layers only created if respective shapes exist)
 * 
 * Overcut:     0.1" (7.2 pt)
 * Reg marks:   0.25" dia circles, 0.5" from cuts, ≤12" apart
 * Alignment:   0.2" 5-point star, 1" left of top-right reg mark
 * Labels:      File name in Arial (0.4" height), right-aligned, on Layer 1
 */

// ─── Configuration ───────────────────────────────────────────────────────────
var OVERCUT         = 7.2;     // 0.1"  in points
var TOLERANCE       = 0.1;     // coordinate matching tolerance (pt)
var STROKE_WEIGHT   = 0.5;     // cut line stroke weight (pt)

var REG_DIAMETER    = 18;      // 0.25" in points
var REG_RADIUS      = 9;       // half diameter
var REG_OFFSET      = 36;      // 0.5"  in points (gap from cut lines)
var REG_MAX_SPACING = 864;     // 12"   in points (max gap between marks)

var STAR_SIZE       = 14.4;    // 0.2"  outer diameter in points
var STAR_RADIUS     = 7.2;     // half star size
var STAR_INSET      = 72;      // 1"    left of top-right reg mark

var LABEL_OFFSET    = 36;      // 0.5"  from reg marks (pt)
var LABEL_FONT_SIZE = 28.8;    // 0.4" text height in points

var ARTBOARD_PAD    = 7.2;     // 0.1" padding around artwork for artboard


// ─── Main ────────────────────────────────────────────────────────────────────
function main() {
    if (!app.documents.length) {
        alert("No document open.");
        return;
    }

    var doc = app.activeDocument;

    // ─── Get file name ───────────────────────────────────────────────────
    var fileName = "Untitled";
    try {
        fileName = doc.name;
    } catch (e) {}

    // ─── Locate spot colors (neither is strictly required) ───────────────
    var thruCutSpot = findSpot(doc, "Thru-cut");
    var kissCutSpot = findSpot(doc, "Kiss-cut");

    if (!thruCutSpot && !kissCutSpot) {
        alert("Neither \"Thru-cut\" nor \"Kiss-cut\" spot color found.\n" +
              "Please add at least one to your swatches and try again.\n\n" +
              "Spot colors found: " + listSpots(doc));
        return;
    }

    var thruCutColor = thruCutSpot ? makeSpotColor(thruCutSpot) : null;
    var kissCutColor = kissCutSpot ? makeSpotColor(kissCutSpot) : null;

    // REG_SF spot (auto-create if missing)
    var regSpot = findSpot(doc, "REG_SF");
    if (!regSpot) {
        regSpot = doc.spots.add();
        regSpot.name = "REG_SF";
        var regRGB = new RGBColor();
        regRGB.red   = 0;
        regRGB.green = 0;
        regRGB.blue  = 0;
        regSpot.color = regRGB;
        regSpot.colorType = ColorModel.SPOT;
    }
    var regColor = makeSpotColor(regSpot);

    var starColor = new RGBColor();
    starColor.red   = 255;
    starColor.green = 0;
    starColor.blue  = 0;

    var labelColor = new RGBColor();
    labelColor.red   = 150;
    labelColor.green = 150;
    labelColor.blue  = 150;


    // ─── Collect qualifying shapes ───────────────────────────────────────
    var thruRects = [];
    var kissShapes = [];

    for (var i = doc.pathItems.length - 1; i >= 0; i--) {
        var item = doc.pathItems[i];
        if (thruCutSpot && isSpotRectangle(item, thruCutSpot)) {
            thruRects.push(item);
        } else if (kissCutSpot && isSpotShape(item, kissCutSpot)) {
            kissShapes.push(item);
        }
    }

    var hasThru = thruRects.length > 0;
    var hasKiss = kissShapes.length > 0;

    if (!hasThru && !hasKiss) {
        alert("No qualifying shapes found.\n\n" +
              "Requirements:\n" +
              "- Thru-cut: Closed 4-corner axis-aligned path, \"Thru-cut\" fill\n" +
              "- Kiss-cut: Closed path with \"Kiss-cut\" fill\n\n" +
              "Total pathItems: " + doc.pathItems.length + "\n" +
              "Spots: " + listSpots(doc));
        return;
    }


    // ─── Reference the original layer ────────────────────────────────────
    var origLayer;
    if (hasThru) {
        origLayer = thruRects[0].layer;
    } else {
        origLayer = kissShapes[0].layer;
    }


    // ─── Bounding box for registration marks ─────────────────────────────
    var bbLeft   =  Infinity, bbRight = -Infinity;
    var bbBottom =  Infinity, bbTop   = -Infinity;


    // ═════════════════════════════════════════════════════════════════════
    // KISS-CUT: Compute bounds BEFORE moving (anchor points, not geometricBounds)
    // ═════════════════════════════════════════════════════════════════════
    if (hasKiss && !hasThru) {
        for (var k2 = 0; k2 < kissShapes.length; k2++) {
            var pts = kissShapes[k2].pathPoints;
            for (var p = 0; p < pts.length; p++) {
                var ax = pts[p].anchor[0];
                var ay = pts[p].anchor[1];
                if (ax < bbLeft)   bbLeft   = ax;
                if (ax > bbRight)  bbRight  = ax;
                if (ay < bbBottom) bbBottom = ay;
                if (ay > bbTop)    bbTop    = ay;
            }
        }
    }


    // ═════════════════════════════════════════════════════════════════════
    // THRU-CUT PROCESSING
    // ═════════════════════════════════════════════════════════════════════
    var hLines = [];
    var vLines = [];
    var cutLayer = null;

    if (hasThru) {
        // Decompose into line segments
        for (var r = 0; r < thruRects.length; r++) {
            var corners = getRectCorners(thruRects[r]);
            var cL = corners.left,  cR = corners.right;
            var cT = corners.top,   cB = corners.bottom;

            hLines.push({ a: cL, b: cR, c: cT });
            hLines.push({ a: cL, b: cR, c: cB });
            vLines.push({ a: cB, b: cT, c: cL });
            vLines.push({ a: cB, b: cT, c: cR });
        }

        normalizeLines(hLines);
        normalizeLines(vLines);

        hLines = deduplicateLines(hLines);
        vLines = deduplicateLines(vLines);

        hLines = mergeCollinear(hLines);
        vLines = mergeCollinear(vLines);

        // Apply overcuts
        for (var h = 0; h < hLines.length; h++) {
            hLines[h].a -= OVERCUT;
            hLines[h].b += OVERCUT;
        }
        for (var v = 0; v < vLines.length; v++) {
            vLines[v].a -= OVERCUT;
            vLines[v].b += OVERCUT;
        }

        // Compute bounding box from overcut lines
        for (var h2 = 0; h2 < hLines.length; h2++) {
            if (hLines[h2].a < bbLeft)   bbLeft   = hLines[h2].a;
            if (hLines[h2].b > bbRight)  bbRight  = hLines[h2].b;
            if (hLines[h2].c < bbBottom) bbBottom = hLines[h2].c;
            if (hLines[h2].c > bbTop)    bbTop    = hLines[h2].c;
        }
        for (var v2 = 0; v2 < vLines.length; v2++) {
            if (vLines[v2].c < bbLeft)   bbLeft   = vLines[v2].c;
            if (vLines[v2].c > bbRight)  bbRight  = vLines[v2].c;
            if (vLines[v2].a < bbBottom) bbBottom = vLines[v2].a;
            if (vLines[v2].b > bbTop)    bbTop    = vLines[v2].b;
        }

        // Remove original thru-cut rectangles
        for (var d = 0; d < thruRects.length; d++) {
            thruRects[d].remove();
        }

        // Create CUT layer
        cutLayer = doc.layers.add();
        cutLayer.name = "CUT";
        cutLayer.move(origLayer, ElementPlacement.PLACEBEFORE);

        // Draw cut lines
        for (var h3 = 0; h3 < hLines.length; h3++) {
            var seg = hLines[h3];
            drawLine(cutLayer, seg.a, seg.c, seg.b, seg.c, thruCutColor);
        }
        for (var v3 = 0; v3 < vLines.length; v3++) {
            var seg2 = vLines[v3];
            drawLine(cutLayer, seg2.c, seg2.a, seg2.c, seg2.b, thruCutColor);
        }
    }


    // ═════════════════════════════════════════════════════════════════════
    // KISS-CUT: Move shapes to KISS layer (AFTER bounds are computed)
    // ═════════════════════════════════════════════════════════════════════
    var kissLayer = null;

    if (hasKiss) {
        kissLayer = doc.layers.add();
        kissLayer.name = "KISS";
        if (cutLayer) {
            kissLayer.move(cutLayer, ElementPlacement.PLACEAFTER);
        } else {
            kissLayer.move(origLayer, ElementPlacement.PLACEBEFORE);
        }

        // Move kiss-cut shapes to KISS layer
        for (var k = 0; k < kissShapes.length; k++) {
            kissShapes[k].move(kissLayer, ElementPlacement.PLACEATEND);
        }
    }


    // ═════════════════════════════════════════════════════════════════════
    // DEBUG: Verify bounding box
    // ═════════════════════════════════════════════════════════════════════
    var bbW = Math.round((bbRight - bbLeft) / 72 * 1000) / 1000;
    var bbH = Math.round((bbTop - bbBottom) / 72 * 1000) / 1000;


    // ═════════════════════════════════════════════════════════════════════
    // REGISTRATION MARKS, STAR, LABELS, ARTBOARD
    // ═════════════════════════════════════════════════════════════════════

    // ─── REG layer (always on top) ───────────────────────────────────────
    var regLayer = doc.layers.add();
    regLayer.name = "REG";
    if (cutLayer) {
        regLayer.move(cutLayer, ElementPlacement.PLACEBEFORE);
    } else if (kissLayer) {
        regLayer.move(kissLayer, ElementPlacement.PLACEBEFORE);
    } else {
        regLayer.move(origLayer, ElementPlacement.PLACEBEFORE);
    }


    // ─── Registration mark positions ─────────────────────────────────────
    var regLeft   = bbLeft   - REG_OFFSET;
    var regRight  = bbRight  + REG_OFFSET;
    var regTop    = bbTop    + REG_OFFSET;
    var regBottom = bbBottom - REG_OFFSET;

    var markPositions = {};

    // Top edge
    var topPts = subdivideEdge(regLeft, regRight);
    for (var t = 0; t < topPts.length; t++) {
        addMark(markPositions, topPts[t], regTop);
    }
    // Bottom edge
    var botPts = subdivideEdge(regLeft, regRight);
    for (var b = 0; b < botPts.length; b++) {
        addMark(markPositions, botPts[b], regBottom);
    }
    // Left edge
    var leftPts = subdivideEdge(regBottom, regTop);
    for (var le = 0; le < leftPts.length; le++) {
        addMark(markPositions, regLeft, leftPts[le]);
    }
    // Right edge
    var rightPts = subdivideEdge(regBottom, regTop);
    for (var ri = 0; ri < rightPts.length; ri++) {
        addMark(markPositions, regRight, rightPts[ri]);
    }

    // Draw registration marks
    var markCount = 0;
    for (var key in markPositions) {
        if (!markPositions.hasOwnProperty(key)) continue;
        var pos = markPositions[key];
        drawCircle(regLayer, pos.x, pos.y, REG_RADIUS, regColor);
        markCount++;
    }


    // ─── Alignment star ──────────────────────────────────────────────────
    var starX = regRight - STAR_INSET;
    var starY = regTop;
    drawStar(regLayer, starX, starY, STAR_RADIUS, starColor);


    // ─── File name labels on original layer ──────────────────────────────
    var labelTopY    = regTop    + LABEL_OFFSET;
    var labelBottomY = regBottom - LABEL_OFFSET;

    var topLabel = drawLabelRight(origLayer, fileName, regRight, labelTopY, labelColor, LABEL_FONT_SIZE);
    var bottomLabel = drawLabelRight(origLayer, fileName, regRight, labelBottomY, labelColor, LABEL_FONT_SIZE);

    // Rename original layer
    try {
        origLayer.name = "Layer 1";
    } catch (e) {}


    // ─── Resize artboard to fit all artwork ──────────────────────────────
    var artLeft   = regLeft   - REG_RADIUS;
    var artRight  = regRight  + REG_RADIUS;
    var artTop    = regTop    + REG_RADIUS;
    var artBottom = regBottom - REG_RADIUS;

    // Expand for alignment star
    if (starX - STAR_RADIUS < artLeft)   artLeft   = starX - STAR_RADIUS;
    if (starX + STAR_RADIUS > artRight)  artRight  = starX + STAR_RADIUS;
    if (starY + STAR_RADIUS > artTop)    artTop    = starY + STAR_RADIUS;
    if (starY - STAR_RADIUS < artBottom) artBottom = starY - STAR_RADIUS;

    // Expand for labels
    try {
        var tlb = topLabel.geometricBounds;
        if (tlb[1] < artLeft)   artLeft   = tlb[1];
        if (tlb[3] > artRight)  artRight  = tlb[3];
        if (tlb[0] > artTop)    artTop    = tlb[0];
        if (tlb[2] < artBottom) artBottom = tlb[2];
    } catch (e) {}

    try {
        var blb = bottomLabel.geometricBounds;
        if (blb[1] < artLeft)   artLeft   = blb[1];
        if (blb[3] > artRight)  artRight  = blb[3];
        if (blb[0] > artTop)    artTop    = blb[0];
        if (blb[2] < artBottom) artBottom = blb[2];
    } catch (e) {}

    // Add padding
    artLeft   -= ARTBOARD_PAD;
    artRight  += ARTBOARD_PAD;
    artTop    += ARTBOARD_PAD;
    artBottom -= ARTBOARD_PAD;

    // artboardRect: [left, top, right, bottom]
    doc.artboards[0].artboardRect = [artLeft, artTop, artRight, artBottom];


    // ─── Summary ─────────────────────────────────────────────────────────
    var artW = Math.round((artRight - artLeft) / 72 * 100) / 100;
    var artH = Math.round((artTop - artBottom) / 72 * 100) / 100;

    var summary = "Done!\n\n";
    if (hasThru) {
        summary += thruRects.length + " Thru-cut rectangle(s) → " +
                   (hLines.length + vLines.length) + " cut line(s) on CUT layer.\n";
    }
    if (hasKiss) {
        summary += kissShapes.length + " Kiss-cut shape(s) moved to KISS layer.\n";
    }
    summary += markCount + " registration mark(s) on REG layer.\n" +
               "1 alignment star placed.\n" +
               "File name labels added (top & bottom).\n\n" +
               "Reg bounds: " + bbW + "\" x " + bbH + "\"\n" +
               "Artboard: " + artW + "\" x " + artH + "\"\n\n" +
               "Reg marks relative to: " + (hasThru ? "Thru-cut" : "Kiss-cut") + " bounds\n" +
               "Overcut: 0.1\" per end\n" +
               "Reg offset: 0.5\" from cuts\n" +
               "Max reg spacing: 12\"";

    alert(summary);
}


// ═══════════════════════════════════════════════════════════════════════════════
// Shape Detection
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Check if a pathItem is an axis-aligned rectangle filled with a given spot.
 * Accepts 4 or 5-point closed paths.
 */
function isSpotRectangle(item, spot) {
    if (!item.closed) return false;

    var numPts = item.pathPoints.length;
    if (numPts < 4 || numPts > 5) return false;

    if (numPts === 5) {
        var first = item.pathPoints[0].anchor;
        var last  = item.pathPoints[4].anchor;
        if (Math.abs(first[0] - last[0]) > TOLERANCE ||
            Math.abs(first[1] - last[1]) > TOLERANCE) {
            return false;
        }
    }

    if (!item.filled) return false;

    var fc = item.fillColor;
    if (fc.typename !== "SpotColor") return false;
    if (fc.spot.name !== spot.name) return false;

    var pts = item.pathPoints;
    for (var p = 0; p < 4; p++) {
        var x1 = pts[p].anchor[0];
        var y1 = pts[p].anchor[1];
        var x2 = pts[(p + 1) % 4].anchor[0];
        var y2 = pts[(p + 1) % 4].anchor[1];

        var dx = Math.abs(x1 - x2);
        var dy = Math.abs(y1 - y2);

        if (dx > TOLERANCE && dy > TOLERANCE) return false;
    }

    return true;
}

/**
 * Check if a pathItem is ANY closed shape filled with a given spot color.
 * Does NOT require axis-alignment or specific point count.
 */
function isSpotShape(item, spot) {
    if (!item.closed) return false;
    if (!item.filled) return false;

    var fc = item.fillColor;
    if (fc.typename !== "SpotColor") return false;
    if (fc.spot.name !== spot.name) return false;

    return true;
}

function getRectCorners(item) {
    var pts = item.pathPoints;
    var minX =  Infinity, maxX = -Infinity;
    var minY =  Infinity, maxY = -Infinity;

    var count = Math.min(pts.length, 4);
    for (var i = 0; i < count; i++) {
        var x = pts[i].anchor[0];
        var y = pts[i].anchor[1];
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
    }

    return { left: minX, right: maxX, top: maxY, bottom: minY };
}


// ═══════════════════════════════════════════════════════════════════════════════
// Line Processing
// ═══════════════════════════════════════════════════════════════════════════════

function normalizeLines(lines) {
    for (var i = 0; i < lines.length; i++) {
        if (lines[i].a > lines[i].b) {
            var tmp = lines[i].a;
            lines[i].a = lines[i].b;
            lines[i].b = tmp;
        }
    }
}

function deduplicateLines(lines) {
    var removed = {};

    for (var i = 0; i < lines.length; i++) {
        if (removed[i]) continue;
        for (var j = i + 1; j < lines.length; j++) {
            if (removed[j]) continue;
            if (Math.abs(lines[i].c - lines[j].c) < TOLERANCE &&
                Math.abs(lines[i].a - lines[j].a) < TOLERANCE &&
                Math.abs(lines[i].b - lines[j].b) < TOLERANCE) {
                removed[j] = true;
            }
        }
    }

    var unique = [];
    for (var k = 0; k < lines.length; k++) {
        if (!removed[k]) unique.push(lines[k]);
    }
    return unique;
}

function mergeCollinear(lines) {
    var groups = [];
    var groupKeys = [];

    for (var i = 0; i < lines.length; i++) {
        var placed = false;
        for (var g = 0; g < groupKeys.length; g++) {
            if (Math.abs(lines[i].c - groupKeys[g]) < TOLERANCE) {
                groups[g].push(lines[i]);
                placed = true;
                break;
            }
        }
        if (!placed) {
            groupKeys.push(lines[i].c);
            groups.push([lines[i]]);
        }
    }

    var merged = [];
    for (var g2 = 0; g2 < groups.length; g2++) {
        var segs = groups[g2];
        segs.sort(function(x, y) { return x.a - y.a; });

        var current = { a: segs[0].a, b: segs[0].b, c: segs[0].c };
        for (var s = 1; s < segs.length; s++) {
            if (segs[s].a <= current.b + TOLERANCE) {
                current.b = Math.max(current.b, segs[s].b);
            } else {
                merged.push(current);
                current = { a: segs[s].a, b: segs[s].b, c: segs[s].c };
            }
        }
        merged.push(current);
    }

    return merged;
}


// ═══════════════════════════════════════════════════════════════════════════════
// Registration Marks
// ═══════════════════════════════════════════════════════════════════════════════

function subdivideEdge(start, end) {
    var length = Math.abs(end - start);
    var numSpans = Math.ceil(length / REG_MAX_SPACING);
    if (numSpans < 1) numSpans = 1;

    var step = (end - start) / numSpans;
    var points = [];
    for (var i = 0; i <= numSpans; i++) {
        points.push(start + i * step);
    }
    return points;
}

function addMark(dict, x, y) {
    var key = Math.round(x * 100) + "," + Math.round(y * 100);
    if (!dict[key]) {
        dict[key] = { x: x, y: y };
    }
}


// ═══════════════════════════════════════════════════════════════════════════════
// Drawing Functions
// ═══════════════════════════════════════════════════════════════════════════════

function drawLine(layer, x1, y1, x2, y2, spotColor) {
    var line = layer.pathItems.add();
    line.setEntirePath([[x1, y1], [x2, y2]]);
    line.closed = false;
    line.filled = false;
    line.stroked = true;
    line.strokeColor = spotColor;
    line.strokeWidth = STROKE_WEIGHT;
    line.strokeCap = StrokeCap.BUTTENDCAP;
    line.strokeJoin = StrokeJoin.MITERENDJOIN;
    return line;
}

function drawCircle(layer, cx, cy, radius, fillColor) {
    var dia = radius * 2;
    var circle = layer.pathItems.ellipse(
        cy + radius,
        cx - radius,
        dia,
        dia,
        false,
        true
    );
    circle.filled = true;
    circle.fillColor = fillColor;
    circle.stroked = false;
    return circle;
}

function drawStar(layer, cx, cy, outerRadius, fillColor) {
    var innerRadius = outerRadius * 0.381966;
    var numPoints = 5;
    var totalPoints = numPoints * 2;
    var angleStep = Math.PI / numPoints;
    var startAngle = -Math.PI / 2;

    var pathPoints = [];
    for (var i = 0; i < totalPoints; i++) {
        var angle = startAngle + (i * angleStep);
        var r2 = (i % 2 === 0) ? outerRadius : innerRadius;
        var px = cx + r2 * Math.cos(angle);
        var py = cy + r2 * Math.sin(angle);
        pathPoints.push([px, py]);
    }

    var star = layer.pathItems.add();
    star.setEntirePath(pathPoints);
    star.closed = true;
    star.filled = true;
    star.fillColor = fillColor;
    star.stroked = false;
    return star;
}

function drawLabelRight(layer, text, rightX, cy, textColor, fontSize) {
    var tf = layer.textFrames.add();
    tf.contents = text;

    var style = tf.textRange.characterAttributes;
    style.size = fontSize;
    style.fillColor = textColor;

    try {
        style.textFont = app.textFonts.getByName("ArialMT");
    } catch (e) {
        try {
            style.textFont = app.textFonts.getByName("Arial-Regular");
        } catch (e2) {}
    }

    tf.position = [0, cy];
    var tfWidth = tf.width;
    var tfHeight = tf.height;
    tf.position = [rightX - tfWidth, cy + tfHeight / 2];

    return tf;
}


// ═══════════════════════════════════════════════════════════════════════════════
// Utilities
// ═══════════════════════════════════════════════════════════════════════════════

function findSpot(doc, name) {
    for (var i = 0; i < doc.spots.length; i++) {
        if (doc.spots[i].name === name) return doc.spots[i];
    }
    return null;
}

function makeSpotColor(spot) {
    var sc = new SpotColor();
    sc.spot = spot;
    sc.tint = 100;
    return sc;
}

function listSpots(doc) {
    var names = [];
    for (var i = 0; i < doc.spots.length; i++) {
        names.push(doc.spots[i].name);
    }
    return names.join(", ");
}


// ─── Execute ─────────────────────────────────────────────────────────────────
main();
