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
