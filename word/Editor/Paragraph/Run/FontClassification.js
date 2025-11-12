/*
 * (c) Copyright Ascensio System SIA 2010-2024
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

"use strict";

const fontslot_None     = 0x00;
const fontslot_ASCII    = 0x01;
const fontslot_EastAsia = 0x02;
const fontslot_CS       = 0x04;
const fontslot_HAnsi    = 0x08;
const fontslot_Unknown  = 0x10;

(function(window)
{

	const fonthint_Default  = 0x00;
	const fonthint_CS       = 0x01;
	const fonthint_EastAsia = 0x02;

	const TABLE_CHUNK_LEN = 0x10000;
	const TABLE_CHUNKS    = 3;
	let   LOOKUP_TABLE    = null;
	const HINT_EA_OFFSET  = TABLE_CHUNK_LEN;
	const HINT_ZH_OFFSET  = TABLE_CHUNK_LEN * 2;

	(function()
	{
		LOOKUP_TABLE = AscFonts.allocate(TABLE_CHUNK_LEN * TABLE_CHUNKS);

		function FillRange(nStart, nEnd, arrFontSlots)
		{
			for (let u = nStart; u <= nEnd; ++u)
			{
				LOOKUP_TABLE[u]                  = arrFontSlots[0];
				LOOKUP_TABLE[u + HINT_EA_OFFSET] = arrFontSlots[1];
				LOOKUP_TABLE[u + HINT_ZH_OFFSET] = arrFontSlots[2];
			}
		}

		function AddExceptions(arrExceptions, nOffset, nFontSlot)
		{
			for (let nIndex = 0, nCount = arrExceptions.length; nIndex < nCount; ++nIndex)
			{
				LOOKUP_TABLE[arrExceptions[nIndex] + nOffset] = nFontSlot;
			}
		}

		// Basic Latin
		FillRange(0x0000, 0x007F, [fontslot_ASCII, fontslot_ASCII, fontslot_ASCII]);

		// Latin-1 Supplement
		FillRange(0x00A0, 0x00FF, [fontslot_HAnsi, fontslot_HAnsi, fontslot_HAnsi]);
		AddExceptions([
			0xA1, 0xA4, 0xA7, 0xA8, 0xAA, 0xAD, 0xAF, 0xB0, 0xB1, 0xB2, 0xB3, 0xB4,
			0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBC, 0xBD, 0xBE, 0xBF, 0xD7, 0xF7
		], HINT_EA_OFFSET, fontslot_EastAsia);
		AddExceptions([
			0xA1, 0xA4, 0xA7, 0xA8, 0xAA, 0xAD, 0xAF, 0xB0, 0xB1, 0xB2, 0xB3, 0xB4,
			0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBC, 0xBD, 0xBE, 0xBF, 0xD7, 0xF7,
			0xE0, 0xE1, 0xE8, 0xE9, 0xEA, 0xEC, 0xED, 0xF2, 0xF3, 0xF9, 0xFA, 0xFC
		], HINT_ZH_OFFSET, fontslot_EastAsia);

		// Latin Extended-A
		// Latin Extended-B
		// IPA Extensions
		FillRange(0x0100, 0x02AF, [fontslot_HAnsi, fontslot_HAnsi, fontslot_EastAsia]);

		// Spacing Modifier Letters
		// Combining Diacritical Marks
		// Greek and Coptic
		// Cyrillic
		FillRange(0x02B0, 0x04FF, [fontslot_HAnsi, fontslot_EastAsia, fontslot_EastAsia]);

		// Hebrew
		// Arabic
		// Syriac
		// Arabic Supplement
		// Thaana
		FillRange(0x0590, 0x07BF, [fontslot_ASCII, fontslot_ASCII, fontslot_ASCII]);

		// Hangul Jamo
		FillRange(0x1100, 0x11FF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Latin Extended Additional
		FillRange(0x1E00, 0x1EFF,  [fontslot_HAnsi, fontslot_HAnsi, fontslot_EastAsia]);

		// Greek Extended
		FillRange(0x1F00, 0x1FFF, [fontslot_ASCII, fontslot_HAnsi, fontslot_HAnsi]);

		// General Punctuation
		// Superscripts and Subscripts
		// Currency Symbols
		// Combining Diacritical Marks for Symbols
		// Letter-like Symbols
		// Number Forms
		// Arrows
		// Mathematical Operators
		// Miscellaneous Technical
		// Control Pictures
		// Optical Character Recognition
		// Enclosed Alphanumerics
		// Box Drawing
		// Block Elements
		// Geometric Shapes
		// Miscellaneous Symbols
		// Dingbats
		FillRange(0x2000, 0x27BF, [fontslot_HAnsi, fontslot_EastAsia, fontslot_EastAsia]);

		// CJK Radicals Supplement
		// Kangxi Radicals
		// Ideographic Description Characters
		// CJK Symbols and Punctuation
		// Hiragana
		// Katakana
		// Bopomofo
		// Hangul Compatibility Jamo
		// Kanbun
		FillRange(0x2E80, 0x319F, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Enclosed CJK Letters and Months
		// CJK Compatibility
		// CJK Unified Ideographs Extension A
		FillRange(0x3200, 0x4DBF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// CJK Unified Ideographs
		FillRange(0x4E00, 0x9FAF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Yi Syllables
		// Yi Radicals
		FillRange(0xA000, 0xA4CF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Hangul Syllables
		FillRange(0xAC00, 0xD7AF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// High Surrogates
		// High Private Use Surrogates
		// Low Surrogates
		FillRange(0xD800, 0xDFFF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Private Use Area
		FillRange(0xE000, 0xF8FF, [fontslot_HAnsi, fontslot_EastAsia, fontslot_EastAsia]);

		// CJK Compatibility Ideographs
		FillRange(0xF900, 0xFAFF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Alphabetic Presentation Forms
		FillRange(0xFB00, 0xFB1C, [fontslot_HAnsi, fontslot_EastAsia, fontslot_EastAsia]);
		FillRange(0xFB1D, 0xFB4F, [fontslot_ASCII, fontslot_ASCII, fontslot_ASCII]);

		// Arabic Presentation Forms-A
		FillRange(0xFB50, 0xFDFF, [fontslot_ASCII, fontslot_ASCII, fontslot_ASCII]);

		// CJK Compatibility Forms
		// Small Form Variants
		FillRange(0xFE30, 0xFE6F, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Arabic Presentation Forms-B
		FillRange(0xFE70, 0xFEFE, [fontslot_ASCII, fontslot_ASCII, fontslot_ASCII]);

		// Halfwidth and Fullwidth Forms
		FillRange(0xFF00, 0xFFEF, [fontslot_EastAsia, fontslot_EastAsia, fontslot_EastAsia]);

		// Все, что выше, сделано согласно документу DR 09-0040
		// Все, что ниже, получено опытным путем

		// Devanagari
		// Bengali
		// Gurmukhi
		// Gujarati
		// Oriya
		// Tamil
		// Telegu
		// Kannada
		// Malayalam
		// Sinhala
		// Thai
		// Lao
		// Tibetan
		// Myanmar
		// Georgian
		FillRange(0x0900, 0x10FF, [fontslot_HAnsi, fontslot_HAnsi, fontslot_HAnsi]);

		// Ethiopic
		// Cherokee
		FillRange(0x1200, 0x13FF, [fontslot_HAnsi, fontslot_HAnsi, fontslot_HAnsi]);

	})();

	function GetFontSlot(nUnicode, nHint, nLangId, isCS, isRTL)
	{
		let nSlot;
		if (nUnicode > 0xFFFF)
		{
			if ((nUnicode >= 0x20000 && nUnicode <= 0x2A6DF) ||
				(nUnicode >= 0x2F800 && nUnicode <= 0x2FA1F))
			{
				nSlot = fontslot_EastAsia;
			}
			else if (nUnicode >= 0x1D400 && nUnicode <= 0x1D7FF)
			{
				nSlot = fontslot_ASCII;
			}
			else
			{
				nSlot = fontslot_HAnsi;
			}
		}
		else if (fonthint_EastAsia !== nHint)
		{
			nSlot = LOOKUP_TABLE[nUnicode];
		}
		else
		{
			if (lcid_zh === nLangId)
				nSlot = LOOKUP_TABLE[HINT_ZH_OFFSET + nUnicode];
			else
				nSlot = LOOKUP_TABLE[HINT_EA_OFFSET + nUnicode];

			if (fontslot_EastAsia === nSlot)
				return nSlot;
		}

		if (isCS || isRTL)
			return fontslot_CS;

		return nSlot ? nSlot : fontslot_ASCII;
	}
	function GetFontSlotByTextPr(nUnicode, oTextPr)
	{
		return GetFontSlot(nUnicode, oTextPr.RFonts.Hint, oTextPr.Lang.EastAsia, oTextPr.CS, oTextPr.RTL);
	}
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].GetFontSlot         = GetFontSlot;
	window['AscWord'].GetFontSlotByTextPr = GetFontSlotByTextPr;

	window['AscWord'].fontslot_None     = fontslot_None;
	window['AscWord'].fontslot_ASCII    = fontslot_ASCII;
	window['AscWord'].fontslot_EastAsia = fontslot_EastAsia;
	window['AscWord'].fontslot_CS       = fontslot_CS;
	window['AscWord'].fontslot_HAnsi    = fontslot_HAnsi;
	window['AscWord'].fontslot_Unknown  = fontslot_Unknown;

	window['AscWord'].fonthint_Default  = fonthint_Default;
	window['AscWord'].fonthint_CS       = fonthint_CS;
	window['AscWord'].fonthint_EastAsia = fonthint_EastAsia;

})(window);

