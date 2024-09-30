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
	const NON_LETTER_SYMBOLS = [];
	NON_LETTER_SYMBOLS[0x00A0] = 1;
	NON_LETTER_SYMBOLS[0x00AE] = 1;

	const CHECKED_LIMIT = 2000;
	
	// Если значения совпадают - значит апостроф развернут в правильную сторону, если нет, то в значении лежит апостроф в нужном направлении
	const APOSTROPHES = {
		0x0027 : 0x0027,
		0x02BC : 0x02BC,
		0x02BD : 0x02BC,
		0x2018 : 0x2019,
		0x2019 : 0x2019
	};
	
	function isCorrectApostrophe(codePoint)
	{
		return APOSTROPHES[codePoint] === codePoint;
	}
	

	/**
	 * Класс для проверки орфографии внутри параграфа
	 * @param oSpellChecker
	 * @param isForceFullCheck
	 * @constructor
	 */
	function CParagraphSpellCheckerCollector(oSpellChecker, isForceFullCheck)
	{
		this.ContentPos   = new AscWord.CParagraphContentPos();
		this.SpellChecker = oSpellChecker;

		this.CurLcid  = -1;
		this.bWord    = false;
		this.sWord    = "";
		
		this.startRun      = null;
		this.startInRunPos = 0;
		this.endRun        = null;
		this.endInRunPos   = 0;
		
		this.Prefix   = null;
		
		this.apostrophe     = null;
		this.lastApostrophe = null;

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
			this.SpellChecker.Add(this.startRun, this.startInRunPos, this.endRun, this.endInRunPos, this.sWord, this.CurLcid, this.GetPrefix(), 0, this.apostrophe);

			this.bWord = false;
			this.sWord = "";
			
			this.apostrophe     = null;
			this.lastApostrophe = null;
		}
	};
	/**
	 * @param {AscWord.CRunElementBase} oElement
	 * @param {CTextPr} oTextPr
	 * @param {AscWord.Run} run
	 * @param {number} inRunPos
	 */
	CParagraphSpellCheckerCollector.prototype.HandleRunElement = function(oElement, oTextPr, run, inRunPos)
	{
		if (this.IsWordLetter(oElement))
		{
			if (!this.bWord)
			{
				this.startRun      = run;
				this.startInRunPos = inRunPos;
				this.endRun        = run;
				this.endInRunPos   = inRunPos + 1;
				
				this.bWord = true;
				this.sWord = oElement.GetCharForSpellCheck(oTextPr.Caps);
			}
			else
			{
				if (this.lastApostrophe)
				{
					this.sWord += isCorrectApostrophe(this.lastApostrophe) ? String.fromCharCode(0x0027) : String.fromCharCode(0x0020);
					this.apostrophe     = APOSTROPHES[this.lastApostrophe];
					this.lastApostrophe = null;
				}
				
				this.sWord += oElement.GetCharForSpellCheck(oTextPr.Caps);
				
				this.endRun      = run;
				this.endInRunPos = inRunPos + 1;
			}
		}
		else if (this.bWord && this.IsApostrophe(oElement))
		{
			if (this.lastApostrophe)
				this.FlushWord();
			else
				this.lastApostrophe = oElement.GetCodePoint();
		}
		else
		{
			if (this.bWord)
			{
				this.SpellChecker.Add(this.startRun, this.startInRunPos, this.endRun, this.endInRunPos, this.sWord, this.CurLcid, this.GetPrefix(), oElement.IsDot() ? oElement.GetCharCode() : 0, this.apostrophe);
				this.bWord = false;
				this.sWord = "";
				this.apostrophe     = null;
				this.lastApostrophe = null;
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
	CParagraphSpellCheckerCollector.prototype.IsPunctuation = function(oElement)
	{
		if (!oElement.IsPunctuation())
			return false;
		
		// Исключения, полученнные опытным путем
		let nUnicode = oElement.GetCodePoint();
		return (!(0x2019 === nUnicode && lcid_frFR === this.CurLcid)
			&& !(0x2018 === nUnicode && (lcid_uzLatnUZ === this.CurLcid || lcid_uzCyrlUZ === this.CurLcid)));
	};
	CParagraphSpellCheckerCollector.prototype.IsWordLetter = function(oElement)
	{
		return (oElement.IsText() && !this.IsPunctuation(oElement) && !NON_LETTER_SYMBOLS[oElement.GetCodePoint()]);
	};
	CParagraphSpellCheckerCollector.prototype.IsApostrophe = function(oElement)
	{
		if (!oElement.IsText())
			return false;
		
		return !!(APOSTROPHES[oElement.GetCodePoint()]);
	};

	/**
	 * Метка начала элемента для проверки
	 * @constructor
	 */
	function SpellMarkStart(spellCheckElement)
	{
		this.Element = spellCheckElement;
	}
	SpellMarkStart.prototype.getElement = function()
	{
		return this.Element;
	};
	SpellMarkStart.prototype.isStart = function()
	{
		return true;
	};
	SpellMarkStart.prototype.onAdd = function(pos)
	{
		if (this.Element.startInRunPos >= pos)
			++this.Element.startInRunPos;
	};
	SpellMarkStart.prototype.onRemove = function(pos, count)
	{
		if (this.Element.startInRunPos > pos + count)
			this.Element.startInRunPos -= count;
		else if (this.Element.startInRunPos > pos)
			this.Element.startInRunPos = Math.max(0, pos);
	};
	SpellMarkStart.prototype.movePos = function(shift)
	{
		this.Element.startInRunPos += shift;
	};
	SpellMarkStart.prototype.getPos = function()
	{
		return this.Element.startInRunPos;
	};
	SpellMarkStart.prototype.isMisspelled = function()
	{
		return false === this.Element.Checked;
	};
	/**
	 * Метка конца элемента для проверки
	 * @constructor
	 */
	function SpellMarkEnd(spellCheckElement)
	{
		this.Element = spellCheckElement;
	}
	SpellMarkEnd.prototype.getElement = function()
	{
		return this.Element;
	};
	SpellMarkEnd.prototype.isStart = function()
	{
		return false;
	};
	SpellMarkEnd.prototype.onAdd = function(pos)
	{
		if (this.Element.endInRunPos >= pos)
			++this.Element.endInRunPos;
	};
	SpellMarkEnd.prototype.onRemove = function(pos, count)
	{
		if (this.Element.endInRunPos > pos + count)
			this.Element.endInRunPos -= count;
		else if (this.Element.endInRunPos > pos)
			this.Element.endInRunPos = Math.max(0, pos);
	};
	SpellMarkEnd.prototype.movePos = function(shift)
	{
		this.Element.endInRunPos += shift;
	};
	SpellMarkEnd.prototype.getPos = function()
	{
		return this.Element.endInRunPos;
	};
	SpellMarkEnd.prototype.isMisspelled = function()
	{
		return false === this.Element.Checked;
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CParagraphSpellCheckerCollector = CParagraphSpellCheckerCollector;
	
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].SpellMarkStart = SpellMarkStart;
	window['AscWord'].SpellMarkEnd   = SpellMarkEnd;

})(window);
