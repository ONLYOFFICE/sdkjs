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
	function CAutoCorrectSettings()
	{
		this.SmartQuotes            = true;
		this.HyphensWithDash        = true;
		this.AutomaticBulletedLists = true;
		this.AutomaticNumberedLists = true;
		this.FrenchPunctuation      = true;
		this.DoubleSpaceWithPeriod  = false;
		this.FirstLetterOfSentences = true;
		this.FirstLetterOfCells     = true;
		this.Hyperlinks             = true;
		this.FirstLetterExceptions  = new AscCommon.CFirstLetterExceptions();
	}
	//getters
	CAutoCorrectSettings.prototype.IsSmartQuotes  = function()
	{
		return this.SmartQuotes;
	};
	CAutoCorrectSettings.prototype.IsHyphensWithDash = function()
	{
		return this.HyphensWithDash;
	};
	CAutoCorrectSettings.prototype.IsAutomaticBulletedLists = function()
	{
		return this.AutomaticBulletedLists;
	};
	CAutoCorrectSettings.prototype.IsAutomaticNumberedLists = function()
	{
		return this.AutomaticNumberedLists;
	};
	CAutoCorrectSettings.prototype.IsFrenchPunctuation = function()
	{
		return this.FrenchPunctuation;
	};
	CAutoCorrectSettings.prototype.IsDoubleSpaceWithPeriod = function()
	{
		return this.DoubleSpaceWithPeriod;
	};
	CAutoCorrectSettings.prototype.IsFirstLetterOfSentences = function()
	{
		return this.FirstLetterOfSentences;
	};
	CAutoCorrectSettings.prototype.IsFirstLetterOfCells = function()
	{
		return this.FirstLetterOfCells;
	};
	CAutoCorrectSettings.prototype.IsHyperlinks = function()
	{
		return this.Hyperlinks;
	};
	//setters
	CAutoCorrectSettings.prototype.SetSmartQuotes  = function(bVal)
	{
		this.SmartQuotes = bVal;
	};
	CAutoCorrectSettings.prototype.SetHyphensWithDash = function(bVal)
	{
		this.HyphensWithDash = bVal;
	};
	CAutoCorrectSettings.prototype.SetAutomaticBulletedLists = function(bVal)
	{
		this.AutomaticBulletedLists = bVal;
	};
	CAutoCorrectSettings.prototype.SetAutomaticNumberedLists = function(bVal)
	{
		this.AutomaticNumberedLists = bVal;
	};
	CAutoCorrectSettings.prototype.SetFrenchPunctuation = function(bVal)
	{
		this.FrenchPunctuation = bVal;
	};
	CAutoCorrectSettings.prototype.SetDoubleSpaceWithPeriod = function(bVal)
	{
		this.DoubleSpaceWithPeriod = bVal;
	};
	CAutoCorrectSettings.prototype.SetFirstLetterOfSentences = function(bVal)
	{
		this.FirstLetterOfSentences = bVal;
	};
	CAutoCorrectSettings.prototype.SetFirstLetterOfCells = function(bVal)
	{
		this.FirstLetterOfCells = bVal;
	};
	CAutoCorrectSettings.prototype.SetHyperlinks = function(bVal)
	{
		this.Hyperlinks = bVal;
	};
	CAutoCorrectSettings.prototype.GetFirstLetterExceptionManager = function()
	{
		return this.FirstLetterExceptions;
	};
	CAutoCorrectSettings.prototype.CheckFirstLetterException = function(word, lang)
	{
		return this.FirstLetterExceptions.Check(word, lang);
	};
	CAutoCorrectSettings.prototype.GetFirstLetterExceptionsMaxLen = function()
	{
		return this.FirstLetterExceptions.GetMaxLen();
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CAutoCorrectSettings = CAutoCorrectSettings;

	CAutoCorrectSettings.prototype["get_FirstLetterExceptionManager"] = CAutoCorrectSettings.prototype.get_FirstLetterExceptionManager = CAutoCorrectSettings.prototype.GetFirstLetterExceptionManager;

})(window);
