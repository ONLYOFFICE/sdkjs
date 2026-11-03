/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// A horizontal ink stroke has no Y span, and a vertical stroke has no X span.
// 0/0 is NaN. The geometry builder converts NaN to 0, so the stroke is drawn
// at the top or left edge of its rect instead of on the points.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "ink.js"), "utf8");
const start = source.indexOf("function inkRelative");
const end = source.indexOf("function getMinRect");
assert.ok(start > 0 && end > start, "inkRelative was not found");

const sandbox = {};
vm.runInNewContext(source.slice(start, end), sandbox);
const inkRelative = sandbox.inkRelative;

assert.strictEqual(inkRelative(0, 0), 0.5);
assert.strictEqual(inkRelative(4, 0), 0.5);
assert.strictEqual(inkRelative(0, 20), 0);
assert.strictEqual(inkRelative(20, 20), 1);
assert.strictEqual(inkRelative(5, 20), 0.25);

// A hair off horizontal still uses the real proportion.
assert.strictEqual(inkRelative(1e-8, 1e-8), 1);
assert.strictEqual(inkRelative(0, 1e-8), 0);

// The annotation rect is padded by the stroke width, then the inner box is
// inset by the same width. A finite center on a zero inner span is the
// original point. NaN collapses to the outer edge.
function place(boundMin, boundMax, lineWidth, relative) {
    const innerMin = boundMin + lineWidth;
    const innerMax = boundMax - lineWidth;
    return (innerMax - innerMin) * relative + innerMin;
}

const lineWidth = 4;
const drawnY = 120;
const rectTop = drawnY - lineWidth;
const rectBottom = drawnY + lineWidth;
assert.strictEqual(place(rectTop, rectBottom, lineWidth, inkRelative(0, 0)), drawnY);

const drawnX = 40;
assert.strictEqual(place(drawnX - lineWidth, drawnX + lineWidth, lineWidth, inkRelative(0, 0)), drawnX);

// Near-horizontal: the small span is preserved, not snapped to the center.
const near = inkRelative(0.2, 0.4);
assert.strictEqual(place(100, 100.4, 0, near), 100.2);

assert.strictEqual((0 * (0 / 0)) >> 0, 0);
