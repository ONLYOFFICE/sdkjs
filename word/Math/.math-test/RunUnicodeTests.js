/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */
import "./node.js";
import "../NamesOfLiterals.js";
import "../UnicodeParser.js";
import "./UnicodeTestsList/sqrt-tests.js";
import "./UnicodeTestsList/box-tests.js";
import "./UnicodeTestsList/brackets-test.js";
import "./UnicodeTestsList/fraction-test.js";
import "./UnicodeTestsList/literal-tests.js";
import "./UnicodeTestsList/aboveAndBelow-test.js";
import "./UnicodeTestsList/complex-stuff.js";
import "./UnicodeTestsList/hbrack-tests.js";
import "./UnicodeTestsList/script-tests.js";
import "./UnicodeTestsList/special_scripts-tests.js";
import { assert } from "chai";

const parser = window.AscMath.CUnicodeConverter;

const sqrt = window.AscMath.sqrt;
const box = window.AscMath.box;
const bracket = window.AscMath.bracket;
const fraction = window.AscMath.fraction;
const literal = window.AscMath.literal;
const aboveBelow = window.AscMath.aboveBelow;
const complex = window.AscMath.complex;
const hbrack = window.AscMath.hbrack;
const script = window.AscMath.script;
const special = window.AscMath.special;

// const CUnicodeConverter = eval(text)
// const assert = require("chai").assert;
//
// // const arr = [
// // 	"∑",
// // 	"⅀",
// // 	"⨊",
// // 	"∏",
// // 	"∐",
// // 	"⨋",
// // 	"∫",
// // 	"∬",
// // 	"∭",
// // 	"⨌",
// // 	"∮",
// // 	"∯",
// // 	"∰",
// // 	"∱",
// // 	"⨑",
// // 	"∲",
// // 	"∳",
// // 	"⨍",
// // 	"⨎",
// // 	"⨏",
// // 	"⨕",
// // 	"⨖",
// // 	"⨗",
// // 	"⨘",
// // 	"⨙",
// // 	"⨚",
// // 	"⨛",
// // 	"⨜",
// // 	"⨒",
// // 	"⨓",
// // 	"⨔",
// // 	"⋀",
// // 	"⋁",
// // 	"⋂",
// // 	"⋃",
// // 	"⨃",
// // 	"⨄",
// // 	"⨅",
// // 	"⨆",
// // 	"⨀",
// // 	"⨁",
// // 	"⨂",
// // 	"⨉",
// // 	"⫿",
// // 	"(",
// // 	"[",
// // 	"{",
// // 	"〈",
// // 	")",
// // 	"]",
// // 	"}",
// // 	"〉",
// // 	"├",
// // 	"┤",
// // 	"┬",
// // 	"┴",
// // 	"▁",
// // 	"¯",
// // 	"▭",
// // 	"□",
// // 	"&",
// // 	"▒",
// // 	"^",
// // 	"_",
// // 	"¦",
// // 	"√",
// // 	"∛",
// // 	"∜",
// // 	"⊘",
// // 	"/",
// // 	",",
// // 	".",
// // 	"⏜",
// // 	"⏝",
// // 	"⎴",
// // 	"⎵",
// // 	"⏞",
// // 	"⏟",
// // 	"⏠",
// // 	"⏡",
// // 	"&",
// // 	"■",
// // ];
// //
// // describe("Проверка литералов из операторов", function () {
// //   arr.forEach((literal) => li_test("\\" + literal));
// // });
//
describe("Проверка работоспособности простых литералов", function () {
	literal(test);
});
describe("Проверка работоспособности деления", function () {
	fraction(test);
});
describe("Проверка работоспособности радикалов", function () {
	sqrt(test);
});
describe("Проверка работоспособности скриптов", function () {
	script(test);
});
describe("Проверка работоспособности below/above", function () {
	aboveBelow(test);
});
describe("Проверка работоспособности hBrack", function () {
	hbrack(test);
});
describe("Проверка работоспособности скобок", function () {
	bracket(test);
});
describe("Проверка работоспособности комплексных выражений", function () {
	complex(test);
});
describe("Проверка box", function () {
	box(test);
});
describe("Проверка special_scripts", function () {
	special(test);
});

function test(program, expected, description = "Без описания") {
	it(description, function () {
		const ast = parser(program, undefined, true);
		assert.deepEqual(
			ast,
			expected
		);
	});
}
