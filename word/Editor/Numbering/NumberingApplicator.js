/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
	const NumberingType = {
		Remove : 0,
		Bullet : 1,
		Number : 2,
		Hybrid : 3
	};

	/**
	 * Класс для применения нумерации к документу
	 * @param {AscWord.CDocument} document
	 * @constructor
	 */
	function CNumberingApplicator(document)
	{
		this.Document = document;

		this.NumPr      = null;
		this.Paragraphs = [];
		this.NumInfo    = null;
	}

	/**
	 * Применяем нумерацию по заданному объекту
	 * @param numInfo {object}
	 */
	CNumberingApplicator.prototype.Apply = function(numInfo)
	{
		if (!this.Document)
			return false;

		this.NumInfo    = numInfo;
		this.NumPr      = this.GetCurrentNumPr();
		this.Paragraphs = this.GetParagraphs();

		if (this.Paragraphs.length)
			return false;

		let result = false;
		if (this.IsRemoveNumbering())
			result = this.RemoveNumbering();
		else if (this.IsBulleted())
			result = this.ApplyBulleted();
		else if (this.IsNumbered())
			result = this.ApplyNumbered();
		else if (this.IsSingleLevel())
			result = this.ApplySingleLevel();
		else if (this.IsMultilevel())
			result = this.ApplyMultilevel();

		return result;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CNumberingApplicator.prototype.GetCurrentNumPr = function()
	{
		let document = this.Document;

		let numPr = document.GetSelectedNum();
		if (!numPr && !document.IsTextSelectionUse())
			numPr = document.GetSelectedNum(true);

		return numPr;
	};
	CNumberingApplicator.prototype.GetParagraphs = function()
	{
		let paragraphs = [];
		if (this.NumPr)
			paragraphs = this.Document.GetAllParagraphsByNumbering(oNumPr);
		else
			paragraphs = this.Document.GetSelectedParagraphs();

		return paragraphs;
	};
	CNumberingApplicator.prototype.IsRemoveNumbering = function()
	{
		return (NumberingType.Remove === this.NumInfo.Type);
	};
	CNumberingApplicator.prototype.IsBulleted = function()
	{
		return (NumberingType.Bullet === this.NumInfo.Type && (!this.NumInfo.Lvl || 0 === this.NumInfo.Lvl.length));
	};
	CNumberingApplicator.prototype.IsNumbered = function()
	{
		return (NumberingType.Number === this.NumInfo.Type && (!this.NumInfo.Lvl || 0 === this.NumInfo.Lvl.length));
	};
	CNumberingApplicator.prototype.IsSingleLevel = function()
	{
		return (NumberingType.Number === this.NumInfo.Type);
	};
	CNumberingApplicator.prototype.IsMultilevel = function()
	{
		return false;
	};
	CNumberingApplicator.prototype.RemoveNumbering = function()
	{

	};
	CNumberingApplicator.prototype.ApplyBulleted = function()
	{

	};
	CNumberingApplicator.prototype.ApplyNumbered = function()
	{

	};
	CNumberingApplicator.prototype.ApplySingleLevel = function()
	{

	};
	CNumberingApplicator.prototype.ApplyMultilevel = function()
	{

	};
	//---------------------------------------------------------export---------------------------------------------------
	window["AscWord"].CNumberingApplicator = CNumberingApplicator;

})(window);
