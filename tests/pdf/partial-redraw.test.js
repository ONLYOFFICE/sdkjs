const assert = require("assert");

global.window = global;
global.AscCommon = {};
global.AscPDF = {};
global.AscFormat = {
	InitClass: function() {},
	CBaseNoIdObject: function() {}
};
global.AscDFH = { historyitem_type_Pdf_Page: 1 };

require("../../pdf/src/viewer.js");

function drawing(id, bounds) {
	return { id: id, bounds: bounds };
}

const drawings = [
	drawing("a", { l: 0, t: 0, r: 10, b: 10 }),
	drawing("b", { l: 9, t: 0, r: 20, b: 10 }),
	drawing("c", { l: 19, t: 0, r: 30, b: 10 }),
	drawing("d", { l: 100, t: 100, r: 110, b: 110 })
];

const result = AscPDF.collectDrawingsForPartialRedraw(drawings, { l: 1, t: 1, r: 2, b: 2 });

assert.deepStrictEqual(result.drawings.map(function(d) { return d.id; }), ["a"]);
assert.deepStrictEqual(result.bounds, { l: 1, t: 1, r: 2, b: 2 });

const diagonalChain = [
	drawing("root", { l: 9, t: 9, r: 20, b: 20 }),
	drawing("gap", { l: 0, t: 15, r: 5, b: 19 }),
	drawing("child", { l: 19, t: 19, r: 30, b: 30 })
];

const diagonalResult = AscPDF.collectDrawingsForPartialRedraw(diagonalChain, { l: 0, t: 0, r: 10, b: 10 });

assert.deepStrictEqual(diagonalResult.drawings.map(function(d) { return d.id; }), ["root"]);
assert.deepStrictEqual(diagonalResult.bounds, { l: 0, t: 0, r: 10, b: 10 });

console.log("partial redraw clipped drawings test passed");
