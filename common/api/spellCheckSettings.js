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
