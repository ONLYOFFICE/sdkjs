/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// Synthetic stand-in for ONLYOFFICE/DocumentServer#3844.
// Three Ink appearances: Multiply at /CA 1, Normal at /CA 1, Multiply at /CA 0.4.
// The shipped wasm reports blend mode 0, so the editor reads /BM from the
// appearance dictionaries. No customer PDF is required.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "base.js"), "utf8");
const start = source.indexOf('window["AscPDF"].indexAppearanceBlend = function');
const end = source.indexOf('window["AscPDF"].startBlendMode = function');
assert.ok(start > 0 && end > start, "indexAppearanceBlend was not found in base.js");

const sandbox = { TextDecoder: TextDecoder };
sandbox.window = sandbox;
sandbox.AscPDF = {};
vm.runInNewContext(source.slice(start, end), sandbox);

const indexAppearanceBlend = sandbox.AscPDF.indexAppearanceBlend;
const fixture = fs.readFileSync(path.join(__dirname, "ink-multiply.fixture.pdf"));
const byName = indexAppearanceBlend(fixture);

assert.strictEqual(byName["ink-9"], 1);
assert.strictEqual(byName["ink-11"], 1);
assert.strictEqual(byName["ink-10"], undefined);

// MuPDF writes the hyphen in /NM as an octal escape. The annotation's real
// name is fitz-A0, so the map key has to be that string or the highlight
// stays on Normal.
const octalName = [
	"1 0 obj",
	"<< /Type /Annot /Subtype /Ink /NM (fitz\\055A0) /AP << /N 2 0 R >> >>",
	"endobj",
	"2 0 obj",
	"<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> >> >> /Length 8 >>",
	"stream",
	"/R0 gs\n",
	"endstream",
	"endobj"
].join("\n");
const octalByName = indexAppearanceBlend(Buffer.from(octalName, "latin1"));
assert.strictEqual(octalByName["fitz-A0"], 1);
assert.strictEqual(octalByName["fitz055A0"], undefined);

assert.strictEqual(sandbox.AscPDF.blendModeForNewInk(true, 0.5, false), 1);
assert.strictEqual(sandbox.AscPDF.blendModeForNewInk(true, 1, false), 0);
assert.strictEqual(sandbox.AscPDF.blendModeForNewInk(true, 0.5, true), 0);
assert.strictEqual(sandbox.AscPDF.blendModeForNewInk(false, 0.5, false), 0);
assert.strictEqual(Object.keys(byName).length, 2);
assert.deepStrictEqual(Object.keys(indexAppearanceBlend(new Uint8Array())), []);

const text = fixture.toString("latin1");
assert.ok(text.indexOf("/BM /Multiply") > 0);
assert.ok(text.indexOf("/BM /Normal") > 0);
assert.ok(text.indexOf("/CA 1") > 0);
assert.ok(text.indexOf("/CA 0.4") > 0);

// The graphics state the content stream selects wins over an earlier /BM.
const selected = indexAppearanceBlend(`
1 0 obj
<< /Type /Annot /Subtype /Ink /NM (use-r1) /AP << /N 2 0 R >> >>
endobj
2 0 obj
<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Normal >> /R1 << /BM /Multiply >> >> >> /Length 8 >>
stream
/R1 gs
endstream
endobj
3 0 obj
<< /Type /Annot /Subtype /Ink /NM (use-normal) /AP << /N 4 0 R >> >>
endobj
4 0 obj
<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> /R1 << /BM /Normal >> >> >> /Length 8 >>
stream
/R1 gs
endstream
endobj
`);
assert.strictEqual(selected["use-r1"], 1);
assert.strictEqual(selected["use-normal"], undefined);

// Two gs operators: the last one is in effect. An indirect ExtGState dictionary
// is followed so the selected name is not confused with a sibling resource.
const lastGs = indexAppearanceBlend(`
1 0 obj
<< /Type /Annot /Subtype /Ink /NM (last-gs) /AP << /N 2 0 R >> >>
endobj
2 0 obj
<< /Type /XObject /Subtype /Form /Resources << /ExtGState 5 0 R >> /Length 16 >>
stream
/R0 gs
/R1 gs
endstream
endobj
5 0 obj
<< /R0 6 0 R /R1 7 0 R >>
endobj
6 0 obj
<< /Type /ExtGState /BM /Normal >>
endobj
7 0 obj
<< /Type /ExtGState /BM /Multiply >>
endobj
`);
assert.strictEqual(lastGs["last-gs"], 1);

// A later generation replaces the appearance. The stream body is not a dictionary.
const updated = indexAppearanceBlend(`
1 0 obj
<< /Type /Annot /Subtype /Ink /NM (ink) /AP << /N 2 0 R >> >>
endobj
2 0 obj
<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Normal >> >> >> /Length 8 >>
stream
/R0 gs
endstream
endobj
2 1 obj
<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> >> >> /Length 8 >>
stream
/R0 gs
endstream
endobj
`);
assert.strictEqual(updated["ink"], 1);

// Compressed bytes that happen to contain the marker are not a blend mode.
const trap = indexAppearanceBlend(`
1 0 obj
<< /Type /Annot /Subtype /Ink /NM (trap) /AP << /N 2 0 R >> >>
endobj
2 0 obj
<< /Type /XObject /Subtype /Form /Filter /FlateDecode /Length 14 >>
stream
/BM /Multiply
endstream
endobj
`);
assert.strictEqual(trap["trap"], undefined);

// Highlights and form fields stay on their own draw paths.
const others = indexAppearanceBlend(`
1 0 obj
<< /Type /Annot /Subtype /Highlight /NM (hi) /BM /Multiply >>
endobj
2 0 obj
<< /Type /Annot /Subtype /Widget /NM (field) /BM /Multiply >>
endobj
3 0 obj
<< /Type /Annot /Subtype /Stamp /NM (stamp) /AP << /N 4 0 R >> >>
endobj
4 0 obj
<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> >> >> /Length 8 >>
stream
/R0 gs
endstream
endobj
`);
assert.strictEqual(others["hi"], undefined);
assert.strictEqual(others["field"], undefined);
assert.strictEqual(others["stamp"], 1);

// Uncompressed object stream. A direct object of the same number would win.
(function () {
	const apBody = "<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> >> >> >>";
	const annotBody = "<< /Type /Annot /Subtype /Ink /NM (in-stm) /AP << /N 4 0 R >> >>";
	const header = "4 0 8 " + (apBody.length + 1);
	const stream = header + "\n" + apBody + "\n" + annotBody;
	const packed = indexAppearanceBlend(
		"9 0 obj\n<< /Type /ObjStm /N 2 /First " + (header.length + 1) + " /Length " + stream.length + " >>\nstream\n" +
		stream + "\nendstream\nendobj\n"
	);
	assert.strictEqual(packed["in-stm"], 1);
})();

// A direct object later in the file wins over an earlier object stream.
// An object stream later in the file replaces an earlier generation-0 object.
(function () {
	function pack(objNum, body) {
		const header = objNum + " 0";
		const stream = header + "\n" + body;
		const first = header.length + 1;
		return "9 0 obj\n<< /Type /ObjStm /N 1 /First " + first + " /Length " + stream.length + " >>\nstream\n" + stream + "\nendstream\nendobj\n";
	}
	const annot = "8 0 obj\n<< /Type /Annot /Subtype /Ink /NM (order) /AP << /N 3 0 R >> >>\nendobj\n";
	const normalDict = "<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Normal >> >> >> >>";
	const multiplyDict = "<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> >> >> >>";
	const normalAp = "3 0 obj\n<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Normal >> >> >> /Length 8 >>\nstream\n/R0 gs\nendstream\nendobj\n";
	const multiplyAp = "3 0 obj\n<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> >> >> /Length 8 >>\nstream\n/R0 gs\nendstream\nendobj\n";
	const directWins = indexAppearanceBlend(pack(3, normalDict) + annot + multiplyAp);
	assert.strictEqual(directWins["order"], 1);
	const streamWins = indexAppearanceBlend(annot + normalAp + pack(3, multiplyDict));
	assert.strictEqual(streamWins["order"], 1);
})();

// One scan of a large file still finds the annot at the end.
const pad = " ".repeat(1000000);
const large = indexAppearanceBlend(pad + `
1 0 obj
<< /Type /Annot /Subtype /Ink /NM (late) /AP << /N 2 0 R >> >>
endobj
2 0 obj
<< /Type /XObject /Subtype /Form /Resources << /ExtGState << /R0 << /BM /Multiply >> >> >> /Length 8 >>
stream
/R0 gs
endstream
endobj
`);
assert.strictEqual(large["late"], 1);

const onPage = sandbox.AscPDF.annotsCompositedBeforeOverlay;
function same(actual, expected) {
	assert.deepStrictEqual(Array.from(actual), expected);
}
same(onPage([
	{ markup: false, blend: 0 },
	{ markup: false, blend: 1 }
]), [0, 1]);
same(onPage([
	{ markup: false, blend: 1 },
	{ markup: false, blend: 0 }
]), [0]);
same(onPage([
	{ markup: true, blend: 1 },
	{ markup: false, blend: 1 },
	{ markup: false, blend: 0 }
]), [1]);
same(onPage([
	{ markup: false, blend: 0 },
	{ markup: false, blend: 0 }
]), []);

console.log("appearance-blend.test.js ok");
