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
import "../LaTeXParser.js";
import "./LaTeXList/fraction.js";
import "./LaTeXList/degree-tests.js";
import "./LaTeXList/brackets-test.js";
import "./LaTeXList/accents-tests.js";
import "./LaTeXList/numericFunctions-test.js";
import "./LaTeXList/sqrt-tests.js";
import "./LaTeXList/style-test.js";
import { assert } from "chai";

const parser = window.AscMath.ConvertLaTeXToTokensList;
const accent = window.AscMath.accents;
const fraction = window.AscMath.fraction;
const degree = window.AscMath.degree;
const brackets = window.AscMath.brackets;
const numericFunctions = window.AscMath.numericFunctions;
const sqrt = window.AscMath.sqrt;
const style = window.AscMath.style;

describe("Сhecking the health of fractions", function () {
	fraction(test);
});
describe("Сhecking the health of degrees and indexes", function () {
	degree(test);
});
describe("Сhecking the health of brackets", function () {
	brackets(test);
});
describe("Сhecking accents", function () {
 	accent(test);
});
describe("Сhecking standard numerical functions", function () {
	numericFunctions(test);
});
describe("Сhecking radical functions", function () {
	sqrt(test);
});
describe("Сhecking math fonts", function () {
	style(test);
});

function test(program, expected, description = "Без описания") {
	it(description, function () {
		const ast = parser(program, undefined, true);
		assert.deepEqual(ast, expected);
	});
}
