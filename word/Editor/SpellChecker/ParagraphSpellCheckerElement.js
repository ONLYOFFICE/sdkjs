/*
 * (c) Copyright Ascensio System SIA 2010-2022
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

(function(window, undefined)
{
	/**
	 * Отдельный элемент проверки орфографии внутри параграфа
	 * @constructor
	 */
	function CParagraphSpellCheckerElement(StartPos, EndPos, Word, Lang, Prefix, Ending)
	{
		this.StartPos = StartPos;
		this.EndPos   = EndPos;
		this.Word     = Word;
		this.Lang     = Lang;
		this.Checked  = null; // null - неизвестно, true - правильное слово, false - неправильное слово
		this.CurPos   = false;
		this.Variants = null;

		this.StartRun = null;
		this.EndRun   = null;

		// В некоторых языках слова идут вместе со знаками пунктуации до или после, например,
		// -abwicklung и bwz. (в немецком языке)
		this.Prefix = Prefix; // Символ приставки, если есть, например, "-"
		this.Ending = Ending; // Символ окончания, если есть, например, "."
	}
	CParagraphSpellCheckerElement.prototype.GetStartPos = function()
	{
		return this.StartPos;
	};
	CParagraphSpellCheckerElement.prototype.GetEndPos = function()
	{
		return this.EndPos;
	};
	CParagraphSpellCheckerElement.prototype.GetPrefix = function()
	{
		return this.Prefix;
	};
	CParagraphSpellCheckerElement.prototype.GetEnding = function()
	{
		return this.Ending;
	};
	CParagraphSpellCheckerElement.prototype.IsCorrect = function()
	{
		return (this.Checked === true);
	};
	CParagraphSpellCheckerElement.prototype.IsWrong = function()
	{
		return (this.Checked === false);
	};
	CParagraphSpellCheckerElement.prototype.IsUndefined = function()
	{
		return (this.Checked === null);
	};
	CParagraphSpellCheckerElement.prototype.IsCurrent = function()
	{
		return this.CurPos;
	};
	CParagraphSpellCheckerElement.prototype.GetStartRun = function()
	{
		return this.StartRun;
	};
	CParagraphSpellCheckerElement.prototype.GetEndRun = function()
	{
		return this.EndRun;
	};
	CParagraphSpellCheckerElement.prototype.GetWord = function()
	{
		return this.Word;
	};
	CParagraphSpellCheckerElement.prototype.GetLang = function()
	{
		return this.Lang;
	};
	CParagraphSpellCheckerElement.prototype.GetVariants = function()
	{
		return this.Variants;
	};
	CParagraphSpellCheckerElement.prototype.SetVariants = function(arrVariants)
	{
		this.Variants = arrVariants ? arrVariants : null;
	};
	CParagraphSpellCheckerElement.prototype.SetCorrect = function()
	{
		this.Checked  = true;
		this.Variants = null;
	};
	CParagraphSpellCheckerElement.prototype.SetUndefined = function()
	{
		this.Checked  = null;
		this.Variants = null;
	};
	CParagraphSpellCheckerElement.prototype.SetWrong = function(arrVariants)
	{
		this.Checked  = false;
		this.Variants = arrVariants ? arrVariants : null;
	};
	CParagraphSpellCheckerElement.prototype.SetCurrent = function()
	{
		this.Checked = true;
		this.CurPos  = true;
	};
	CParagraphSpellCheckerElement.prototype.ResetCurrent = function()
	{
		this.CurPos = false;
	};
	CParagraphSpellCheckerElement.prototype.ClearSpellingMarks = function()
	{
		if (this.StartRun !== this.EndRun)
		{
			if (this.StartRun)
				this.StartRun.ClearSpellingMarks();

			if (this.EndRun)
				this.EndRun.ClearSpellingMarks();
		}
		else
		{
			if (this.EndRun)
				this.EndRun.ClearSpellingMarks();
		}
	};
	CParagraphSpellCheckerElement.prototype.CheckPositionInside = function(oPos)
	{
		return (oPos && this.EndPos.Compare(oPos) >= 0 && this.StartPos.Compare(oPos) <= 0);
	};
	CParagraphSpellCheckerElement.prototype.CheckIntersection = function(oStartPos, oEndPos)
	{
		return (oStartPos && oEndPos && this.StartPos.Compare(oEndPos) <= 0 && this.EndPos.Compare(oStartPos) >= 0);
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CParagraphSpellCheckerElement = CParagraphSpellCheckerElement;

})(window);
