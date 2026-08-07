/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/*
 * Regression tests for AscCommon.compareNumbers.
 *
 * From this directory:  node comparenumbers.js   (exit code 1 if anything fails)
 * Or open comparenumbers.html and read the browser console.
 *
 * The boundaries below were read off Microsoft Excel 16.111.3 on macOS by building
 * base + k * ulp(base) and finding the largest k that still compares equal.
 */
(function () {
	"use strict";

	/* Decided up front: the node branch below defines window, so this cannot be
	   re-tested later on. */
	var isNode = (typeof window === "undefined");
	var root;
	if (!isNode) {
		/* The html page has already set up the namespaces and loaded NumFormat.js. */
		root = window;
	} else {
		/* Under node, give NumFormat.js the globals it writes into, then load it. */
		root = global;
		root.window = root;
		root.AscCommon = root.AscCommon || {};
		root.Asc = root.Asc || {};
		root.AscFonts = root.AscFonts || {};
		root.AscCommonExcel = root.AscCommonExcel || {};
		require("../NumFormat.js");
	}

	var cmp = root.AscCommon.compareNumbers;
	var failures = 0, total = 0;

	function log(s) {
		if (typeof console !== "undefined" && console.log) {
			console.log(s);
		}
	}

	function check(name, actual, expected) {
		total++;
		var ok = (actual === expected);
		if (!ok) {
			failures++;
		}
		log((ok ? "PASS  " : "FAIL  ") + name + "   got " + actual + ", expected " + expected);
	}

	function sign(n) {
		return n < 0 ? -1 : (n > 0 ? 1 : 0);
	}

	function checkCmp(name, a, b, expected) {
		check(name, sign(cmp(a, b)), expected);
	}

	/* Largest k for which base and base + direction * k * 2^ulpExp still compare equal. */
	function largestEqualK(base, ulpExp, direction) {
		var ulp = Math.pow(2, ulpExp);
		for (var k = 1; k < 4096; k++) {
			if (0 !== cmp(base, base + direction * k * ulp)) {
				return k - 1;
			}
		}
		return -1;
	}

	function checkBoundary(name, base, ulpExp, direction, expectedK) {
		check(name, largestEqualK(base, ulpExp, direction), expectedK);
	}

	function run() {
		/* The case this was written for: a column total against the same total reached
		   by multiplying. The two are one bit apart and straddle the 15-digit rounding
		   boundary, so rounding each operand separately does not make them equal. */
		checkCmp("total by SUM vs by product", 6.4583333333333348, 6.4583333333333339, 0);

		/* Ordinary comparisons must be unaffected. */
		checkCmp("plainly smaller", 1, 2, -1);
		checkCmp("plainly greater", 2, 1, 1);
		checkCmp("identical", 1.5, 1.5, 0);
		checkCmp("zero against zero", 0, 0, 0);
		checkCmp("zero against tiny", 0, 1e-300, -1);
		checkCmp("tiny against zero", 1e-300, 0, 1);
		checkCmp("negative against positive", -1e-20, 1e-20, -1);

		/* A difference in the last stored digit is a real difference. */
		checkCmp("15th digit differs", 1.00000000000001, 1.00000000000002, -1);
		checkCmp("15th digit differs, negated", -1.00000000000001, -1.00000000000002, 1);

		/* Boundaries, against the values measured in Excel. */
		checkBoundary("boundary at 1", 1, -52, 1, 22);
		checkBoundary("boundary at 1024", 1024, -42, 1, 21);
		checkBoundary("boundary at 0.03125", 0.03125, -57, 1, 7);
		checkBoundary("boundary at 3.125", 3.125, -51, 1, 11);
		checkBoundary("boundary at 1e200", 1e200, 612, 1, 29);

		/* The threshold follows the decimal exponent, so it steps by ten across a power
		   of ten even though the binary spacing is the same on both sides. */
		checkBoundary("boundary just below ten", 9.9, -49, 1, 2);
		checkBoundary("boundary just above ten", 10.1, -49, 1, 28);

		/* When the two operands sit on opposite sides of a power of ten, the threshold
		   comes from the smaller one. Sign plays no part in any of this. */
		checkBoundary("crossing one downwards", 1, -53, -1, 4);
		checkBoundary("crossing one downwards, negated", -1, -52, 1, 2);
		checkBoundary("negative, no crossing", -1, -52, -1, 22);

		log("");
		log(failures ? (failures + " of " + total + " failed") : ("all " + total + " passed"));
		return failures;
	}

	if (isNode) {
		process.exitCode = run() ? 1 : 0;
	} else {
		root.runCompareNumbersTests = run;
	}
})();
