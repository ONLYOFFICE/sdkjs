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
	const TYPES = {
		LTR : 'L', // Left-To-Right letter
		RTL : 'R', // Right-To-Left letter
		AL  : 'A', // Arabic Letter
		EN  : '1', // European Numeral
		AN  : '9', // Arabic Numeral
		ES  : 'w', // European number Separator
		ET  : 'w', // European number Terminator
		CS  : 'w', // Common Separator
		NSM : '`', // Non Spacing Mark
		BN  : 'b', // Boundary Neutral
		BS  : 'B', // Block Separator
		SS  : 'S', // Segment Separator
		WS  : '_', // WhiteSpace
		ON  : 'n', // Other Neutral
		LRE : '+', // Left-to-Right Embedding
		RLE : '+', // Right-to-Left Embedding
		LRO : '+', // Left-to-Right Override
		RLO : '+', // Right-to-Left Override
		PDF : '-', // Pop Directional Flag
		LRI : '+', // Left-to-Right Isolate
		RLI : '+', // Right-to-Left Isolate
		FSI : '+', // First-Strong Isolate
		PDI : '-'  // Pop Directional Isolate
	};
	
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
		
		FS  : 0x01000000, // FIRST
	};
	
	
	const TYPE = {
		
		LTR : FLAG.STRONG | FLAG.LETTER,
		RTL : FLAG.STRONG | FLAG.LETTER | FLAG.RTL,
		
		// Arabic letter
		AL : FLAG.STRONG | FLAG.LETTER | FLAG.ARABIC | FLAG.RTL,
		
		// Embedding
		LRE : FLAG.STRONG | FLAG.EXPLICIT,
		RLE : FLAG.STRONG | FLAG.EXPLICIT | FLAG.RTL,
		
		// Override
		LRO : FLAG.STRONG | FLAG.EXPLICIT | FLAG.OVERRIDE,
		RLO : FLAG.STRONG | FLAG.EXPLICIT | FLAG.OVERRIDE | FLAG.RTL,
		
		// Pop Directional Flag
		PDF : FLAG.WEAK | FLAG.EXPLICIT,
		
		EN : FLAG.WEAK | FLAG.NUMBER,
		AN : FLAG.WEAK | FLAG.NUMBER | FLAG.ARABIC,
		
		ES : FLAG.WEAK | FLAG.NUMBER_SEP_TER | FLAG.ES,
		ET : FLAG.WEAK | FLAG.NUMBER_SEP_TER | FLAG.ET,
		
		// Common Separator
		CS : FLAG.WEAK | FLAG.NUMBER_SEP_TER | FLAG.CS,
		
		// Non Spacing Mark
		NSM : FLAG.WEAK | FLAG.NSM,
		
		// Boundary Neutral
		BN : FLAG.WEAK | FLAG.SPACE | FLAG.BN,
		
		// Block Separator
		BS : FLAG.NEUTRAL | FLAG.SPACE | FLAG.SEPARATOR | FLAG.BS,
		
		// Segment Separator
		SS : FLAG.NEUTRAL | FLAG.SPACE | FLAG.SEPARATOR | FLAG.SS,
		
		// White Space
		WS : FLAG.NEUTRAL | FLAG.SPACE | FLAG.WS,
		
		// Other Neutral
		ON : FLAG.NEUTRAL,
		
		WLTR     : FLAG.WEAK,
		WRTL     : FLAG.WEAK | FLAG.RTL,
		SENTINEL : FLAG.SENTINEL,
		
		// Isolates
		LRI : FLAG.NEUTRAL | FLAG.ISOLATE,
		RLI : FLAG.NEUTRAL | FLAG.ISOLATE | FLAG.RTL,
		FSI : FLAG.NEUTRAL | FLAG.ISOLATE | FLAG.FS,
		PDI : FLGA.NEUTRAL | FLAG.WEAK | FLAG.ISOLATE // NEUTRAL ?
	};

	function generateTypes()
	{
	
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	AscBidi.TYPES = TYPES;
	AscBidi.TYPE = TYPE;
})();
