/**
 * ThruCut Overcut Generator v2
 * 
 * Decomposes rectangles (filled with spot color "Thru-cut")
 * into individual cut lines with configurable overcuts.
 * Handles tiled/abutting rectangles by deduplicating shared edges
 * and merging collinear segments before applying overcuts.
 * 
 * Overcut distance: 0.1" (7.2 pt)
 * Output: Stroke lines using the "Thru-cut" spot color
 */

// ─── Configuration ───────────────────────────────────────────────────────────
var OVERCUT = 7.2;        // 0.1 inch in points (1 inch = 72 pt)
var TOLERANCE = 0.1;      // point tolerance for coordinate matching
var STROKE_WEIGHT = 0.5;  // output line stroke weight in points

// ─── Main ────────────────────────────────────────────────────────────────────
function main() {
    if (!app.documents.length) {
        alert("No document open.");
        return;
    }

    var doc = app.activeDocument;

    // --- Locate the "Thru-cut" spot color ---
    var thruCutSpot = null;
    for (var s = 0; s < doc.spots.length; s++) {
        if (doc.spots[s].name === "Thru-cut") {
            thruCutSpot = doc.spots[s];
            break;
        }
    }
    if (!thruCutSpot) {
        alert("Spot color \"Thru-cut\" not found in this document.\nPlease add it to your swatches and try again.");
        return;
    }

    var thruCutColor = new SpotColor();
    thruCutColor.spot = thruCutSpot;
    thruCutColor.tint = 100;

    // --- Collect qualifying rectangles ---
    var rects = [];
    for (var i = doc.pathItems.length - 1; i >= 0; i--) {
        var item = doc.pathItems[i];
        if (isThruCutRectangle(item, thruCutSpot)) {
            rects.push(item);
        }
    }

    if (rects.length === 0) {
        alert("No qualifying rectangles found.\n\n" +
              "Requirements:\n" +
              "- Closed path, 4 corners (axis-aligned)\n" +
              "- Fill: \"Thru-cut\" spot color\n\n" +
              "Total pathItems in doc: " + doc.pathItems.length + "\n" +
              "Spot colors: " + listSpots(doc));
        return;
    }

    // --- Decompose rectangles into line segments ---
    // hLines: horizontal segments {a: minX, b: maxX, c: Y}
    // vLines: vertical segments   {a: minY, b: maxY, c: X}
    var hLines = [];
    var vLines = [];

    for (var r = 0; r < rects.length; r++) {
        var corners = getRectCorners(rects[r]);
        var left   = corners.left;
        var right  = corners.right;
        var top    = corners.top;
        var bottom = corners.bottom;

        // Top edge (horizontal)
        hLines.push({ a: left, b: right, c: top });
        // Bottom edge (horizontal)
        hLines.push({ a: left, b: right, c: bottom });
        // Left edge (vertical)
        vLines.push({ a: bottom, b: top, c: left });
        // Right edge (vertical)
        vLines.push({ a: bottom, b: top, c: right });
    }

    // --- Normalize: ensure a <= b ---
    normalizeLines(hLines);
    normalizeLines(vLines);

    // --- Deduplicate exact overlapping segments ---
    hLines = deduplicateLines(hLines);
    vLines = deduplicateLines(vLines);

    // --- Merge collinear abutting/overlapping segments ---
    hLines = mergeCollinear(hLines);
    vLines = mergeCollinear(vLines);

    // --- Apply overcuts ---
    for (var h = 0; h < hLines.length; h++) {
        hLines[h].a -= OVERCUT;
        hLines[h].b += OVERCUT;
    }
    for (var v = 0; v < vLines.length; v++) {
        vLines[v].a -= OVERCUT;
        vLines[v].b += OVERCUT;
    }

    // --- Draw output lines ---
    var parentLayer = rects[0].layer;

    for (var h2 = 0; h2 < hLines.length; h2++) {
        var seg = hLines[h2];
        drawLine(parentLayer, seg.a, seg.c, seg.b, seg.c, thruCutColor);
    }
    for (var v2 = 0; v2 < vLines.length; v2++) {
        var seg2 = vLines[v2];
        drawLine(parentLayer, seg2.c, seg2.a, seg2.c, seg2.b, thruCutColor);
    }

    // --- Remove original rectangles ---
    for (var d = 0; d < rects.length; d++) {
        rects[d].remove();
    }

    alert("Done!\n\n" +
          rects.length + " rectangle(s) converted.\n" +
          (hLines.length + vLines.length) + " cut line(s) generated.\n" +
          "Overcut: 0.1\" per end.");
}


// ═══════════════════════════════════════════════════════════════════════════════
// Helper Functions
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Check if a pathItem is a qualifying axis-aligned rectangle
 * filled with the "Thru-cut" spot color.
 *
 * Accepts 4-point or 5-point closed paths (some EPS exporters
 * add a redundant lineto back to the start before closepath).
 * Stroke state is ignored — only fill color matters.
 */
function isThruCutRectangle(item, thruCutSpot) {
    // Must be a closed path
    if (!item.closed) return false;

    // Accept 4 or 5 points (handle redundant duplicate endpoint)
    var numPts = item.pathPoints.length;
    if (numPts < 4 || numPts > 5) return false;

    // If 5 points, verify the 5th is a duplicate of the 1st
    if (numPts === 5) {
        var first = item.pathPoints[0].anchor;
        var last  = item.pathPoints[4].anchor;
        if (Math.abs(first[0] - last[0]) > TOLERANCE ||
            Math.abs(first[1] - last[1]) > TOLERANCE) {
            return false; // 5th point is not a duplicate — not a simple rectangle
        }
    }

    // Must be filled
    if (!item.filled) return false;

    // Fill color must be the Thru-cut spot
    var fc = item.fillColor;
    if (fc.typename !== "SpotColor") return false;
    if (fc.spot.name !== thruCutSpot.name) return false;

    // Verify axis-aligned: each of the first 4 edges must be
    // either purely horizontal or purely vertical
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
 * Extract the bounding rectangle coordinates from a path's anchor points.
 * More reliable than geometricBounds for stroked paths.
 */
function getRectCorners(item) {
    var pts = item.pathPoints;
    var minX =  Infinity, maxX = -Infinity;
    var minY =  Infinity, maxY = -Infinity;

    // Only use first 4 points (ignore duplicate 5th)
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

/**
 * Ensure a <= b for all line segments.
 */
function normalizeLines(lines) {
    for (var i = 0; i < lines.length; i++) {
        if (lines[i].a > lines[i].b) {
            var tmp = lines[i].a;
            lines[i].a = lines[i].b;
            lines[i].b = tmp;
        }
    }
}

/**
 * Remove exact duplicate line segments.
 * Two lines are duplicates if all three values (a, b, c) match within TOLERANCE.
 * Handles cases where more than 2 copies exist.
 */
function deduplicateLines(lines) {
    var removed = {};

    for (var i = 0; i < lines.length; i++) {
        if (removed[i]) continue;

        for (var j = i + 1; j < lines.length; j++) {
            if (removed[j]) continue;

            if (Math.abs(lines[i].c - lines[j].c) < TOLERANCE &&
                Math.abs(lines[i].a - lines[j].a) < TOLERANCE &&
                Math.abs(lines[i].b - lines[j].b) < TOLERANCE) {
                // Mark duplicate for removal — do NOT break,
                // continue scanning for additional copies
                removed[j] = true;
            }
        }
    }

    var unique = [];
    for (var k = 0; k < lines.length; k++) {
        if (!removed[k]) {
            unique.push(lines[k]);
        }
    }
    return unique;
}

/**
 * Merge collinear overlapping/abutting segments.
 * Groups lines by their fixed coordinate (c), then within each group
 * sorts by start (a) and merges any that overlap or touch.
 */
function mergeCollinear(lines) {
    // Group by c-value (within tolerance)
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

        // Sort by start position
        segs.sort(function(x, y) { return x.a - y.a; });

        var current = { a: segs[0].a, b: segs[0].b, c: segs[0].c };

        for (var s = 1; s < segs.length; s++) {
            if (segs[s].a <= current.b + TOLERANCE) {
                // Overlapping or abutting — extend
                current.b = Math.max(current.b, segs[s].b);
            } else {
                // Gap — push current and start new
                merged.push(current);
                current = { a: segs[s].a, b: segs[s].b, c: segs[s].c };
            }
        }
        merged.push(current);
    }

    return merged;
}

/**
 * Draw a single line path between two points with the Thru-cut spot color.
 */
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

/**
 * List all spot color names in the document (for debug alert).
 */
function listSpots(doc) {
    var names = [];
    for (var i = 0; i < doc.spots.length; i++) {
        names.push(doc.spots[i].name);
    }
    return names.join(", ");
}


// ─── Execute ─────────────────────────────────────────────────────────────────
main();
