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

(function (window)
{
	/**
	 * @constructor
	 */
	function AutoHyphenationSettings()
	{
		this.autoHyphenation  = false;
		this.hyphenateCaps    = true;
		this.hyphenationZone  = 360; // twips
		this.hyphenationLimit = 0;
	}
	AutoHyphenationSettings.prototype.isAutoHyphenation = function()
	{
		return this.autoHyphenation;
	};
	AutoHyphenationSettings.prototype.setAutoHyphenation = function(isAuto)
	{
		this.autoHyphenation = isAuto;
	};
	AutoHyphenationSettings.prototype.isHyphenateCaps = function()
	{
		return this.hyphenateCaps;
	};
	AutoHyphenationSettings.prototype.setHyphenateCaps = function(isHyphenate)
	{
		this.hyphenateCaps = isHyphenate;
	};
	AutoHyphenationSettings.prototype.getHyphenationZone = function()
	{
		return this.hyphenationZone;
	};
	AutoHyphenationSettings.prototype.setHyphenationZone = function(zone)
	{
		this.hyphenationZone = zone;
	};
	AutoHyphenationSettings.prototype.getHyphenationLimit = function()
	{
		return this.hyphenationLimit;
	};
	AutoHyphenationSettings.prototype.setHyphenationLimit = function(limit)
	{
		this.hyphenationLimit = limit;
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].AutoHyphenationSettings = window['AscCommon']['AutoHyphenationSettings'] = AutoHyphenationSettings;
	
	const p = AutoHyphenationSettings.prototype;
	p["isAutoHyphenation"]   = p.isAutoHyphenation;
	p["setAutoHyphenation"]  = p.setAutoHyphenation;
	p["isHyphenateCaps"]     = p.isHyphenateCaps;
	p["setHyphenateCaps"]    = p.setHyphenateCaps;
	p["getHyphenationZone"]  = p.getHyphenationZone;
	p["setHyphenationZone"]  = p.setHyphenationZone;
	p["getHyphenationLimit"] = p.getHyphenationLimit;
	p["setHyphenationLimit"] = p.setHyphenationLimit;
	
})(window);
