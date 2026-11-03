/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// A pen or highlighter click stores one track point. The track used to be
// discarded, and a one-point ink path has no length to stroke. The dot is a
// filled disk whose diameter is the stroke width.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const trackSource = fs.readFileSync(path.join(__dirname, "../../../common/Drawings/TrackObjects/PolyLine.js"), "utf8");
const trackStart = trackSource.indexOf("function inkTrackCanCreate");
const trackEnd = trackSource.indexOf("PolyLine.prototype.canCreateShape");
assert.ok(trackStart > 0 && trackEnd > trackStart);

const trackSandbox = {};
vm.runInNewContext(trackSource.slice(trackStart, trackEnd), trackSandbox);
const inkTrackCanCreate = trackSandbox.inkTrackCanCreate;

assert.strictEqual(inkTrackCanCreate(1, true), true);
assert.strictEqual(inkTrackCanCreate(0, true), false);
assert.strictEqual(inkTrackCanCreate(2, true), true);
assert.strictEqual(inkTrackCanCreate(1, false), false);
assert.strictEqual(inkTrackCanCreate(2, false), true);

const inkSource = fs.readFileSync(path.join(__dirname, "ink.js"), "utf8");
const inkStart = inkSource.indexOf("function isSinglePointInk");
const inkEnd = inkSource.indexOf("function fillShapeByPoints");
assert.ok(inkStart > 0 && inkEnd > inkStart);

const inkSandbox = {};
vm.runInNewContext(inkSource.slice(inkStart, inkEnd), inkSandbox);

assert.strictEqual(inkSandbox.isSinglePointInk([[12, 34]]), true);
assert.strictEqual(inkSandbox.isSinglePointInk([[12, 34, 56, 78]]), false);
assert.strictEqual(inkSandbox.isSinglePointInk([[12, 34], [56, 78]]), false);
assert.strictEqual(inkSandbox.isSinglePointInk([]), false);

// Rect padded by the full stroke width on each side. The disk diameter is that width.
const dot = inkSandbox.singlePointInkDot([10, 20, 14, 24]);
assert.strictEqual(dot.cx, 12);
assert.strictEqual(dot.cy, 22);
assert.strictEqual(dot.rx, 1);
assert.strictEqual(dot.ry, 1);
assert.strictEqual(inkSandbox.singlePointInkDot([5, 5, 5, 5]), null);

const commands = [];
const geometry = {
    AddPathCommand: function () { commands.push(Array.prototype.slice.call(arguments)); },
    AddRect: function () { commands.push(["rect"].concat(Array.prototype.slice.call(arguments))); }
};
assert.strictEqual(inkSandbox.addFilledInkDot(geometry, [10, 20, 14, 24]), true);
assert.deepStrictEqual(commands[0], [0, undefined, "norm", false, 43200, 43200]);
assert.strictEqual(commands[2][0], 1);
assert.ok(commands.filter((command) => command[0] === 5).length === 4);
assert.strictEqual(commands[commands.length - 1][0], 6);

// A flat stroke still uses the axis division. Only a single point is pinned to the center.
assert.ok(inkSource.indexOf("bSinglePoint ? 0.5 : nIndX / (xMax - xMin)") > 0);
assert.ok(inkSource.indexOf("bSinglePoint ? 0.5 : nIndY / (yMax - yMin)") > 0);
