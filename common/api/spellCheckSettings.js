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
	const FLAGS_MASK       = 0xFFFF;
	const IGNORE_UPPERCASE = 0x0001;
	const IGNORE_NUMBERS   = 0x0002;

	const NON_IGNORE_UPPERCASE = FLAGS_MASK ^ IGNORE_UPPERCASE;
	const NON_IGNORE_NUMBERS   = FLAGS_MASK ^ IGNORE_NUMBERS;

	/**
	 * @constructor
	 */
	function CSpellCheckSettings()
	{
		this.Flags = IGNORE_UPPERCASE | IGNORE_NUMBERS;
	}
	CSpellCheckSettings.prototype.Copy = function()
	{
		let oSettings = new CSpellCheckSettings();
		oSettings.Flags = this.Flags;
		return oSettings;
	};
	CSpellCheckSettings.prototype.IsEqual = function(oSettings)
	{
		return (this.Flags === oSettings.Flags);
	};
	CSpellCheckSettings.prototype.IsIgnoreWordsInUppercase = function()
	{
		return !!(this.Flags & IGNORE_UPPERCASE);
	};
	CSpellCheckSettings.prototype.SetIgnoreWordsInUppercase = function(isIgnore)
	{
		if (isIgnore)
			this.Flags |= IGNORE_UPPERCASE;
		else
			this.Flags &= NON_IGNORE_UPPERCASE;
	};
	CSpellCheckSettings.prototype.IsIgnoreWordsWithNumbers = function()
	{
		return !!(this.Flags & IGNORE_NUMBERS);
	};
	CSpellCheckSettings.prototype.SetIgnoreWordsWithNumbers = function(isIgnore)
	{
		if (isIgnore)
			this.Flags |= IGNORE_NUMBERS;
		else
			this.Flags &= NON_IGNORE_NUMBERS;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CSpellCheckSettings = window['AscCommon']['CSpellCheckSettings'] = CSpellCheckSettings;
	CSpellCheckSettings.prototype['get_IgnoreWordsInUppercase'] = CSpellCheckSettings.prototype.get_IgnoreWordsInUppercase = CSpellCheckSettings.prototype.IsIgnoreWordsInUppercase;
	CSpellCheckSettings.prototype['put_IgnoreWordsInUppercase'] = CSpellCheckSettings.prototype.put_IgnoreWordsInUppercase = CSpellCheckSettings.prototype.SetIgnoreWordsInUppercase;
	CSpellCheckSettings.prototype['get_IgnoreWordsWithNumbers'] = CSpellCheckSettings.prototype.get_IgnoreWordsWithNumbers = CSpellCheckSettings.prototype.IsIgnoreWordsWithNumbers;
	CSpellCheckSettings.prototype['put_IgnoreWordsWithNumbers'] = CSpellCheckSettings.prototype.put_IgnoreWordsWithNumbers = CSpellCheckSettings.prototype.SetIgnoreWordsWithNumbers;

})(window);
