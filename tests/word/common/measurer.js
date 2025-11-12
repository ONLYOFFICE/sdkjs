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

(function(window)
{
	window.AscFonts = window.AscFonts || {};
	onLoadFontsModule = function () {};

	const CharWidth   = 0.5;
	const FontSize    = 10;
	const FontHeight  = 20;
	const FontAscent  = 15;
	const FontDescent = 5;

	// TODO: Сейчас fontSize=10, потому что по умолчанию такой в редакторе, нужно, чтобы здесь он выставлялся

	const GraphemeNormal       = 20;
	const GraphemeLigature_ffi = 25;
	const GraphemeLigature_ff  = 30;
	const GraphemeLigature_fi  = 35;

	const GraphemeCombining_xyz = 40;
	const GraphemeCombining_xy  = 45;

	function GetCharWidth(codePoint)
	{
		if (0x00BC === codePoint) // ¼
			return 0.25 * CharWidth;
		else if (0x00BD === codePoint) // ½
			return 0.5 * CharWidth;
		else if (0x00BE === codePoint) // ¾
			return 0.75 * CharWidth;
		
		return CharWidth;
	}

	let HB_String                            = [];
	AscFonts.HB_StartString                  = function()
	{
		HB_String.length = 0;
	};
	AscFonts.HB_EndString                    = function()
	{

	};
	AscFonts.HB_AppendToString               = function(u)
	{
		HB_String.push(u);
	};
	AscFonts.GetGraphemeWidth = function()
	{
		return CharWidth;
	};
	AscFonts.CTextShaper.prototype.FlushWord = function()
	{
		AscFonts.HB_EndString();

		for (let nIndex = 0, nCount = HB_String.length; nIndex < nCount; ++nIndex)
		{
			if (nCount - nIndex >= 3
				&& AscTest.Letter.f === HB_String[nIndex]
				&& AscTest.Letter.f === HB_String[nIndex + 1]
				&& AscTest.Letter.i === HB_String[nIndex + 2])
			{
				this.FlushGrapheme(GraphemeLigature_ffi, CharWidth * 3, 3, true);
				nIndex += 2;
			}
			else if (nCount - nIndex >= 2
				&& AscTest.Letter.f === HB_String[nIndex]
				&& AscTest.Letter.f === HB_String[nIndex + 1])
			{
				this.FlushGrapheme(GraphemeLigature_ff, CharWidth * 2, 2, true);
				nIndex += 1;
			}
			else if (nCount - nIndex >= 2
				&& AscTest.Letter.f === HB_String[nIndex]
				&& AscTest.Letter.i === HB_String[nIndex + 1])
			{
				this.FlushGrapheme(GraphemeLigature_fi, CharWidth * 2, 2, true);
				nIndex += 1;
			}
			else if (nCount - nIndex >= 3
				&& AscTest.Letter.x === HB_String[nIndex]
				&& AscTest.Letter.y === HB_String[nIndex + 1]
				&& AscTest.Letter.z === HB_String[nIndex + 2])
			{
				this.FlushGrapheme(GraphemeCombining_xyz, CharWidth * 3, 3, false);
				nIndex += 2;
			}
			else if (nCount - nIndex >= 2
				&& AscTest.Letter.x === HB_String[nIndex]
				&& AscTest.Letter.y === HB_String[nIndex + 1])
			{
				this.FlushGrapheme(GraphemeCombining_xy, CharWidth * 2, 2, false);
				nIndex += 1;
			}
			else
			{
				this.Direction = this.GetDirection(HB_String[nIndex]);
				this.FlushGrapheme(GraphemeNormal, GetCharWidth(HB_String[nIndex]), 1, false);
			}
		}

		AscFonts.HB_StartString();
	};
	const g_oTextMeasurer = window.g_oTextMeasurer || AscCommon.g_oTextMeasurer;
	g_oTextMeasurer.SetFontInternal = function()
	{
	};
	g_oTextMeasurer.SetTextPr       = function()
	{
	};
	g_oTextMeasurer.SetFontSlot     = function()
	{
	};
	g_oTextMeasurer.SetFont         = function()
	{
	};
	g_oTextMeasurer.GetHeight       = function()
	{
		return FontHeight;
	};
	g_oTextMeasurer.GetAscender     = function()
	{
		return FontAscent;
	};
	g_oTextMeasurer.GetDescender    = function()
	{
		return FontDescent;
	};
	g_oTextMeasurer.MeasureCode     = function(codePoint)
	{
		return {Width : GetCharWidth(codePoint) * FontSize};
	};
	g_oTextMeasurer.Measure2Code     = function()
	{
		return {Width : CharWidth * FontSize};
	};
	g_oTextMeasurer.Measure         = function(char)
	{
		return g_oTextMeasurer.MeasureCode(char ? char.codePointAt(0) : 0x0020);
	};
	AscCommon.g_fontApplication = AscCommon.g_fontApplication || window.g_fontApplication || {};
	AscCommon.g_fontApplication.LoadFont = function () {

	};
	//--------------------------------------------------------export----------------------------------------------------
	AscTest.CharWidth   = CharWidth;
	AscTest.FontSize    = FontSize;
	AscTest.FontHeight  = FontHeight;
	AscTest.FontAscent  = FontAscent;
	AscTest.FontDescent = FontDescent;

})(window);
