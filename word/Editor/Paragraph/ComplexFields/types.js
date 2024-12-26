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
	window['AscWord'] = window['AscWord'] || {};
	
	window['AscWord'].fieldtype_UNKNOWN      = 0x0000;
	window['AscWord'].fieldtype_MERGEFIELD   = 0x0001;
	window['AscWord'].fieldtype_PAGENUM      = 0x0002;
	window['AscWord'].fieldtype_PAGECOUNT    = 0x0003;
	window['AscWord'].fieldtype_FORMTEXT     = 0x0004;
	window['AscWord'].fieldtype_TOC          = 0x0005;
	window['AscWord'].fieldtype_PAGEREF      = 0x0006;
	window['AscWord'].fieldtype_PAGE         = window['AscWord'].fieldtype_PAGENUM;
	window['AscWord'].fieldtype_NUMPAGES     = window['AscWord'].fieldtype_PAGECOUNT;
	window['AscWord'].fieldtype_ASK          = 0x0007;
	window['AscWord'].fieldtype_REF          = 0x0008;
	window['AscWord'].fieldtype_HYPERLINK    = 0x0009;
	window['AscWord'].fieldtype_TIME         = 0x000A;
	window['AscWord'].fieldtype_DATE         = 0x000B;
	window['AscWord'].fieldtype_FORMULA      = 0x0010;
	window['AscWord'].fieldtype_SEQ          = 0x0011;
	window['AscWord'].fieldtype_STYLEREF     = 0x0012;
	window['AscWord'].fieldtype_NOTEREF      = 0x0013;
	window['AscWord'].fieldtype_ADDIN        = 0x0014;
	window['AscWord'].fieldtype_FORMCHECKBOX = 0x0015;
	window['AscWord'].fieldtype_FORMDROPDOWN = 0x0016;
	
})(window);
