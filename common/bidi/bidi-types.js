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
		PM  : FLAG.STRONG // Paragraph end mark or strong with main direction of the text (tab/break)
	};
	
	const DIRECTION = {
		L : TYPE.L,
		R : TYPE.R
	};
	
	const DIRECTION_FLAG = {
		None  : 0x00,
		LTR   : 0x01,
		RTL   : 0x02,
		Other : 0x04
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscBidi.FLAG           = FLAG;
	AscBidi.TYPE           = TYPE;
	AscBidi.DIRECTION      = DIRECTION;
	AscBidi.DIRECTION_FLAG = DIRECTION_FLAG;
})();
