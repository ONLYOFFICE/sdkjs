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

/**
 * Font slot selection in CParagraphTextShaper.GetFontSlot.
 *
 * The point of interest is that DrawingML text has no <w:cs/> or <w:rtl/> to
 * decide with, so the slot is taken from the character there and only there.
 * These tests pin both halves of that: the DrawingML behaviour and the
 * WordprocessingML behaviour that must not change.
 *
 * No fonts are loaded and nothing is measured - GetFontSlot is a pure decision
 * over a code point and the run properties, so the shaper is driven directly.
 *
 * Open fontslot.html in a browser, or run this file with node:
 *   node word/Editor/Paragraph/test/fontslot.js
 */
(function(root)
{
	"use strict";

	if (typeof require === "function" && typeof module === "object")
	{
		// node: load the same sources the html page loads
		root.window = root;
		let path = require("path");
		let sdk  = path.resolve(__dirname, "../../../..");
		require(path.join(sdk, "common/libfont/common.js"));
		// loader.js needs a document to build its canvas, so out of a browser
		// only the one function the lookup table needs is taken from it. This
		// is the same body, see AscFonts.allocate in common/libfont/loader.js.
		root.AscFonts.allocate = function(size) { return new Uint8Array(size); };
		require(path.join(sdk, "common/commonDefines.js"));
		require(path.join(sdk, "word/Editor/Paragraph/Run/FontClassification.js"));
		require(path.join(sdk, "common/libfont/textshaper.js"));
		require(path.join(sdk, "common/libfont/stringshaper.js"));
		require(path.join(sdk, "word/Editor/Paragraph/TextShaper.js"));
	}

	let AscWord  = root.AscWord;
	let AscFonts = root.AscFonts;
	let SCRIPT   = AscFonts.HB_SCRIPT;

	const SLOT_NAME = {};
	SLOT_NAME[AscWord.fontslot_None]     = "None";
	SLOT_NAME[AscWord.fontslot_ASCII]    = "ASCII";
	SLOT_NAME[AscWord.fontslot_EastAsia] = "EastAsia";
	SLOT_NAME[AscWord.fontslot_HAnsi]    = "HAnsi";
	SLOT_NAME[AscWord.fontslot_CS]       = "CS";

	/**
	 * Minimal stand-in for the compiled text properties. GetFontSlotByTextPr
	 * reads exactly these four values.
	 */
	function textPr(isCS, isRTL)
	{
		return {
			CS     : !!isCS,
			RTL    : !!isRTL,
			RFonts : {Hint : AscWord.fonthint_Default},
			Lang   : {EastAsia : 0x0409}
		};
	}

	/**
	 * @param nUnicode   code point under test
	 * @param isDrawing  true for <a:txBody> content, false for a document body
	 * @param oTextPr    run properties
	 * @param nScript    script of the segment being shaped (-1 when it has not
	 *                   started yet). Only reached by marks and joiners.
	 */
	function slotOf(nUnicode, isDrawing, oTextPr, nScript)
	{
		let shaper = AscWord.ParagraphTextShaper;
		shaper.Paragraph = {bFromDocument : !isDrawing};
		shaper.TextPr    = oTextPr || textPr(false, false);
		shaper.Script    = undefined !== nScript ? nScript : -1;
		return shaper.GetFontSlot(nUnicode);
	}

	let failed = 0;
	let total  = 0;

	function check(sName, nActual, nExpected)
	{
		++total;
		let isOk = (nActual === nExpected);
		if (!isOk)
			++failed;

		let sText = (isOk ? "PASS" : "FAIL") + "  " + sName
			+ "  ->  " + SLOT_NAME[nActual]
			+ (isOk ? "" : "  (expected " + SLOT_NAME[nExpected] + ")");

		if (root.document)
		{
			let div = root.document.createElement("div");
			div.textContent = sText;
			div.style.color = isOk ? "#207520" : "#c00000";
			root.document.body.appendChild(div);
		}
		console.log(sText);
	}

	/**
	 * Asserts the slot is the one the unpatched lookup would give. Used for
	 * everything that must keep its current behaviour, so the expected value
	 * cannot drift out of date.
	 */
	function checkSame(sName, nUnicode, isDrawing, oTextPr, nScript)
	{
		oTextPr = oTextPr || textPr(false, false);
		check(sName, slotOf(nUnicode, isDrawing, oTextPr, nScript),
			AscWord.GetFontSlotByTextPr(nUnicode, oTextPr));
	}

	const DRAWING  = true;
	const DOCUMENT = false;

	// ---- DrawingML: the scripts PowerPoint draws with <a:cs> ----------------
	check("DrawingML Hebrew alef U+05D0", slotOf(0x05D0, DRAWING), AscWord.fontslot_CS);
	check("DrawingML Arabic alef U+0627", slotOf(0x0627, DRAWING), AscWord.fontslot_CS);
	check("DrawingML Syriac alaph U+0710", slotOf(0x0710, DRAWING), AscWord.fontslot_CS);
	check("DrawingML Syriac Supplement U+0860", slotOf(0x0860, DRAWING), AscWord.fontslot_CS);
	check("DrawingML Thaana U+0780", slotOf(0x0780, DRAWING), AscWord.fontslot_CS);
	check("DrawingML Thaana U+07A6", slotOf(0x07A6, DRAWING), AscWord.fontslot_CS);

	// ---- DrawingML: everything else keeps the slot it had -------------------
	check("DrawingML Latin A U+0041", slotOf(0x0041, DRAWING), AscWord.fontslot_ASCII);
	check("DrawingML Hiragana U+3042", slotOf(0x3042, DRAWING), AscWord.fontslot_EastAsia);
	// Greek and Cyrillic land in fontslot_HAnsi, which the drawer resolves to
	// _rfonts.HAnsi - filled from <a:latin> just as Ascii is. Asserted against
	// the unpatched lookup rather than a hard coded slot.
	checkSame("DrawingML Greek U+03B1 unchanged", 0x03B1, DRAWING);
	checkSame("DrawingML Cyrillic U+0410 unchanged", 0x0410, DRAWING);

	// ---- WordprocessingML must not change ----------------------------------
	// Word draws a flagless Hebrew run with the ascii/hAnsi font, not the cs one.
	check("Document Hebrew, no flags", slotOf(0x05D0, DOCUMENT), AscWord.fontslot_ASCII);
	check("Document Arabic, no flags", slotOf(0x0627, DOCUMENT), AscWord.fontslot_ASCII);
	// Thaana is one of the four scripts that move in DrawingML, so this is the
	// case that shows the split is by content origin and not by script alone.
	checkSame("Document Thaana unchanged", 0x0780, DOCUMENT);
	checkSame("Document Japanese unchanged", 0x3042, DOCUMENT);
	check("Document Hebrew, <w:cs/>", slotOf(0x05D0, DOCUMENT, textPr(true, false)), AscWord.fontslot_CS);
	check("Document Hebrew, <w:rtl/>", slotOf(0x05D0, DOCUMENT, textPr(false, true)), AscWord.fontslot_CS);

	// ---- An explicit flag still wins in DrawingML too -----------------------
	// The flags are never set there today, but if a path ever sets them it takes
	// precedence over the character.
	check("DrawingML Latin, CS set", slotOf(0x0041, DRAWING, textPr(true, false)), AscWord.fontslot_CS);
	check("DrawingML Latin, RTL set", slotOf(0x0041, DRAWING, textPr(false, true)), AscWord.fontslot_CS);

	// ---- Marks and joiners follow the base letter --------------------------
	// A font change flushes the shaping buffer, so a mark that lands in another
	// slot than its base is shaped apart from it and the word comes out broken.
	//
	// U+060C..U+074A is folded into Arabic by GetTextScript, so these need no
	// help from the segment - Script is left unset to show that.
	check("DrawingML fatha U+064E", slotOf(0x064E, DRAWING, null, -1), AscWord.fontslot_CS);
	check("DrawingML shadda U+0651", slotOf(0x0651, DRAWING, null, -1), AscWord.fontslot_CS);
	check("DrawingML sukun U+0652", slotOf(0x0652, DRAWING, null, -1), AscWord.fontslot_CS);
	check("DrawingML tatweel U+0640", slotOf(0x0640, DRAWING, null, -1), AscWord.fontslot_CS);
	check("DrawingML Arabic comma U+060C", slotOf(0x060C, DRAWING, null, -1), AscWord.fontslot_CS);

	// Outside that range a mark is Inherited and carries no script of its own,
	// so it continues the segment. ZWJ inside an Arabic word is the real case.
	check("DrawingML ZWJ in Arabic segment",
		slotOf(0x200D, DRAWING, null, SCRIPT.HB_SCRIPT_ARABIC), AscWord.fontslot_CS);
	check("DrawingML ZWNJ in Arabic segment",
		slotOf(0x200C, DRAWING, null, SCRIPT.HB_SCRIPT_ARABIC), AscWord.fontslot_CS);
	check("DrawingML ZWJ in Hebrew segment",
		slotOf(0x200D, DRAWING, null, SCRIPT.HB_SCRIPT_HEBREW), AscWord.fontslot_CS);
	checkSame("DrawingML ZWJ in Latin segment unchanged",
		0x200D, DRAWING, null, SCRIPT.HB_SCRIPT_LATIN);

	// Common is deliberately not carried forward: a space after Hebrew keeps the
	// slot it has without this change, and it already ends a shaping segment.
	checkSame("DrawingML space after Hebrew segment unchanged",
		0x0020, DRAWING, null, SCRIPT.HB_SCRIPT_HEBREW);
	checkSame("DrawingML digit after Hebrew segment unchanged",
		0x0031, DRAWING, null, SCRIPT.HB_SCRIPT_HEBREW);

	let sResult = failed
		? (failed + " of " + total + " checks FAILED")
		: ("all " + total + " checks passed");

	if (root.document)
	{
		let h = root.document.createElement("h3");
		h.textContent = sResult;
		h.style.color = failed ? "#c00000" : "#207520";
		root.document.body.appendChild(h);
	}
	console.log(sResult);

	if (failed && typeof process === "object" && process.exit)
		process.exit(1);

})(typeof window !== "undefined" ? window : globalThis);
