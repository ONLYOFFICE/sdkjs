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

(function ()
{
	const FLAG = {
		RTL      : 0x00000001,
		STRONG   : 0x00000010,
		WEAK     : 0x00000020,
		NEUTRAL  : 0x00000040,
		SENTINEL : 0x00000080,
		ARABIC   : 0x00000002,
		
		LETTER         : 0x00000100,
		NUMBER         : 0x00000200,
		NUMBER_SEP_TER : 0x00000400,
		SPACE          : 0x00000800,
		EXPLICIT       : 0x00001000,
		ISOLATE        : 0x00002000,
		SEPARATOR      : 0x00004000,
		OVERRIDE       : 0x00008000,
		
		ES : 0x00010000, // European Separator
		ET : 0x00020000, // European Terminator
		CS : 0x00040000, // Common Separator
		
		NSM : 0x00080000, // Non Spacing Mark
		BN  : 0x00100000, // Boundary Neutral
		
		BS  : 0x00200000, // Block Separator
		SS  : 0x00400000, // Segment Separator
		WS  : 0x00800000, // WhiteSpace
		
		FS  : 0x01000000  // FIRST
	};
	
	
	const TYPE = {
		L   : FLAG.STRONG | FLAG.LETTER,
		R   : FLAG.STRONG | FLAG.LETTER | FLAG.RTL,
		AL  : FLAG.STRONG | FLAG.LETTER | FLAG.ARABIC | FLAG.RTL,
		EN  : FLAG.WEAK | FLAG.NUMBER,
		ES  : FLAG.WEAK | FLAG.NUMBER_SEP_TER | FLAG.ES,
		ET  : FLAG.WEAK | FLAG.NUMBER_SEP_TER | FLAG.ET,
		AN  : FLAG.WEAK | FLAG.NUMBER | FLAG.ARABIC,
		CS  : FLAG.WEAK | FLAG.NUMBER_SEP_TER | FLAG.CS,
		NSM : FLAG.WEAK | FLAG.NSM,
		BN  : FLAG.WEAK | FLAG.SPACE | FLAG.BN,
		B   : FLAG.NEUTRAL | FLAG.SPACE | FLAG.SEPARATOR | FLAG.BS,
		S   : FLAG.NEUTRAL | FLAG.SPACE | FLAG.SEPARATOR | FLAG.SS,
		WS  : FLAG.NEUTRAL | FLAG.SPACE | FLAG.WS,
		ON  : FLAG.NEUTRAL,
		LRE : FLAG.STRONG | FLAG.EXPLICIT,
		LRO : FLAG.STRONG | FLAG.EXPLICIT | FLAG.OVERRIDE,
		RLE : FLAG.STRONG | FLAG.EXPLICIT | FLAG.RTL,
		RLO : FLAG.STRONG | FLAG.EXPLICIT | FLAG.OVERRIDE | FLAG.RTL,
		PDF : FLAG.WEAK | FLAG.EXPLICIT,
		LRI : FLAG.NEUTRAL | FLAG.ISOLATE,
		RLI : FLAG.NEUTRAL | FLAG.ISOLATE | FLAG.RTL,
		FSI : FLAG.NEUTRAL | FLAG.ISOLATE | FLAG.FS,
		PDI : FLAG.NEUTRAL | FLAG.WEAK | FLAG.ISOLATE,
		PM  : FLAG.STRONG // Paragraph end mark
	};
	
	const DIRECTION = {
		L : TYPE.L,
		R : TYPE.R
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscBidi.FLAG      = FLAG;
	AscBidi.TYPE      = TYPE;
	AscBidi.DIRECTION = DIRECTION;
})();
