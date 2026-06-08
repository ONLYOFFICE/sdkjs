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
