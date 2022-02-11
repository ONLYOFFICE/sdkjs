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
	const CHECKED_LIMIT = 2000;

	/**
	 * Класс для проверки орфографии внутри параграфа
	 * @param oSpellChecker
	 * @param isForceFullCheck
	 * @constructor
	 */
	function CParagraphSpellCheckerCollector(oSpellChecker, isForceFullCheck)
	{
		this.ContentPos   = new CParagraphContentPos();
		this.SpellChecker = oSpellChecker;

		this.CurLcid  = -1;
		this.bWord    = false;
		this.sWord    = "";
		this.StartPos = null; // CParagraphContentPos
		this.EndPos   = null; // CParagraphContentPos
		this.Prefix   = null;

		// Защита от проверки орфографии в большом параграфе
		// TODO: Возможно стоить заменить проверку с количества пройденных элементов на время выполнения
		this.CheckedCounter = 0;
		this.CheckedLimit   = CHECKED_LIMIT;
		this.FindStart      = false;
		this.ForceFullCheck = !!isForceFullCheck;
	}
	/**
	 * Обновляем текущую позицию на заданной глубине
	 * @param nPos {number}
	 * @param nDepth {number}
	 */
	CParagraphSpellCheckerCollector.prototype.UpdatePos = function(nPos, nDepth)
	{
		this.ContentPos.Update(nPos, nDepth);
	};
	/**
	 * Получаем текущую позицию на заданном уровне
	 * @param nDepth
	 * @return {number}
	 */
	CParagraphSpellCheckerCollector.prototype.GetPos = function(nDepth)
	{
		return this.ContentPos.Get(nDepth);
	};
	/**
	 * Проверяем превышен ли лимит возможнных проверок в параграфе за один проход таймера
	 * @return {boolean}
	 */
	CParagraphSpellCheckerCollector.prototype.IsExceedLimit = function()
	{
		return (!this.ForceFullCheck && this.CheckedCounter >= this.CheckedLimit);
	};
	/**
	 * Увеличиваем счетчик проверенных элементов
	 */
	CParagraphSpellCheckerCollector.prototype.IncreaseCheckedCounter = function()
	{
		this.CheckedCounter++;
	};
	/**
	 * Перестартовываем счетчик
	 */
	CParagraphSpellCheckerCollector.prototype.ResetCheckedCounter = function()
	{
		this.CheckedCounter = 0;
	};
	/**
	 * Если проверка была приостановлена и сейчас мы ищем начальную позицию
	 * @return {boolean}
	 */
	CParagraphSpellCheckerCollector.prototype.IsFindStart = function()
	{
		return this.FindStart;
	};
	/**
	 * Выставляем ищем ли мы место, где закончили проверку прошлый раз
	 * @param isFind {boolean}
	 */
	CParagraphSpellCheckerCollector.prototype.SetFindStart = function(isFind)
	{
		this.FindStart = isFind;
	};
	/**
	 * Получиаем возможную приставку до слова (обычно это знак "-")
	 * @returns {number}
	 */
	CParagraphSpellCheckerCollector.prototype.GetPrefix = function()
	{
		if (this.Prefix && this.Prefix.IsHyphen())
			return this.Prefix.GetCharCode();

		return 0;
	};
	CParagraphSpellCheckerCollector.prototype.CheckPrefix = function(oItem)
	{
		this.Prefix = oItem;
	};
	/**
	 * Данная команда останавливает сборку элемента для проверки орфографии
	 */
	CParagraphSpellCheckerCollector.prototype.FlushWord = function()
	{
		if (this.bWord)
		{
			this.SpellChecker.Add(this.StartPos, this.EndPos, this.sWord, this.CurLcid, this.GetPrefix(), 0);

			this.bWord = false;
			this.sWord = "";
		}
	};
	/**
	 * @param {CRunElementBase} oElement
	 * @param {CTextPr} oTextPr
	 */
	CParagraphSpellCheckerCollector.prototype.HandleRunElement = function(oElement, oTextPr)
	{
		if (oElement.IsText() && !oElement.IsPunctuation() && !oElement.IsNBSP() && !oElement.Is_SpecialSymbol())
		{
			if (!this.bWord)
			{
				this.StartPos = this.ContentPos.Copy();
				this.EndPos   = this.ContentPos.Copy();

				let nDepth = this.ContentPos.GetDepth();
				this.EndPos.Update(this.ContentPos.Get(nDepth) + 1, nDepth);

				this.bWord = true;
				this.sWord = oElement.GetCharForSpellCheck(oTextPr.Caps);
			}
			else
			{
				this.sWord += oElement.GetCharForSpellCheck(oTextPr.Caps);

				this.EndPos = this.ContentPos.Copy();

				let nDepth = this.ContentPos.GetDepth();
				this.EndPos.Update(this.ContentPos.Get(nDepth) + 1, nDepth);
			}
		}
		else
		{
			if (this.bWord)
			{
				this.bWord = false;
				this.SpellChecker.Add(this.StartPos, this.EndPos, this.sWord, this.CurLcid, this.GetPrefix(), oElement.IsDot() ? oElement.GetCharCode() : 0);
				this.CheckPrefix(null);
			}
			else
			{
				this.CheckPrefix(oElement);
			}
		}

		this.IncreaseCheckedCounter();
	};
	CParagraphSpellCheckerCollector.prototype.HandleLang = function(nLang)
	{
		if (this.CurLcid === nLang)
			return;

		this.FlushWord();

		this.CurLcid = nLang;
	};

	/**
	 * Метка начала/конца элемента для проверки
	 * @constructor
	 */
	function CParagraphSpellingMark(SpellCheckerElement, Start, Depth)
	{
		this.Element = SpellCheckerElement;
		this.Start   = Start;
		this.Depth   = Depth;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CParagraphSpellCheckerCollector = CParagraphSpellCheckerCollector;
	window['AscCommonWord'].CParagraphSpellingMark          = CParagraphSpellingMark;

})(window);
