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

(function(window, undefined)
{
	const IGNORE_UPPERCASE = 0x0001;
	const IGNORE_NUMBERS   = 0x0002;

	/**
	 * Класс для хранения элементов проверки орфографии
	 * @constructor
	 */
	function CParagraphSpellChecker(oParagraph)
	{
		this.Elements  = [];
		this.RecalcId  = -1;
		this.Paragraph = oParagraph;
		this.Words     = {};
		this.Collector = null;
		this.Flags     = IGNORE_UPPERCASE | IGNORE_NUMBERS;
	}

	CParagraphSpellChecker.prototype.Clear = function()
	{
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			this.Elements[nIndex].ClearSpellingMarks();
		}

		this.Elements = [];
		this.Words    = {};
	};
	/**
	 * @param oSettings {AscCommon.CSpellCheckSettings}
	 */
	CParagraphSpellChecker.prototype.UpdateSettings = function(oSettings)
	{
		this.Flags = 0;

		if (oSettings.IsIgnoreWordsInUppercase())
			this.Flags |= IGNORE_UPPERCASE;

		if (oSettings.IsIgnoreWordsWithNumbers())
			this.Flags |= IGNORE_NUMBERS;
	};
	CParagraphSpellChecker.prototype.SetRecalcId = function(nRecalcId)
	{
		this.RecalcId = nRecalcId;
	};
	CParagraphSpellChecker.prototype.Check = function(nRecalcId, isCheckCurrentWord)
	{
		let arrWords = [];
		let arrLangs = [];
		this.private_GetWordsListForRequest(arrWords, arrLangs, isCheckCurrentWord);

		let isFirst = (this.RecalcId === -1);
		if (undefined !== nRecalcId)
			this.SetRecalcId(nRecalcId);

		if (0 < arrWords.length
			&& true === this.GetDocumentSpellChecker().AddWaitingParagraph(this.Paragraph, this.RecalcId, arrWords, arrLangs))
		{
			editor.SpellCheckApi.spellCheck({
				"type"        : "spell",
				"ParagraphId" : this.Paragraph.GetId(),
				"RecalcId"    : this.RecalcId,
				"ElementId"   : 0,
				"usrWords"    : arrWords,
				"usrLang"     : arrLangs
			});
		}
		else
		{
			this.private_ClearMarksForCorrectWords();
		}

		return (arrWords.length || isFirst);
	};
	CParagraphSpellChecker.prototype.private_GetWordsListForRequest = function(arrWords, arrLangs, isCheckCurrentWord)
	{
		let oCurPos = this.GetCurrentPositionInParagraph();

		let isSkipCurrentWord = true !== isCheckCurrentWord && oCurPos ? this.SkipCurrentWord() : false;
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			let oElement = this.Elements[nIndex];
			let sWord    = oElement.GetWord();
			let nLang    = oElement.GetLang();

			oElement.ResetCurrent();
			if (!this.HaveDictionary(nLang) || !this.IsNeedCheckWord(sWord))
			{
				oElement.SetCorrect();
			}
			else if (oElement.IsUndefined() && isSkipCurrentWord && oElement.CheckPositionInside(oCurPos))
			{
				oElement.SetCurrent();
				this.AddCurrentParagraph();
			}

			if (oElement.IsUndefined())
			{
				arrWords.push(sWord);
				arrLangs.push(nLang);

				let nPrefix = this.Elements[nIndex].GetPrefix();
				let nEnding = this.Elements[nIndex].GetEnding();

				if (nPrefix)
				{
					arrWords.push(String.fromCharCode(nPrefix) + sWord);
					arrLangs.push(nLang);
				}

				if (nEnding)
				{
					arrWords.push(sWord + String.fromCharCode(nEnding));
					arrLangs.push(nLang);
				}

				if (nPrefix && nEnding)
				{
					arrWords.push(String.fromCharCode(nPrefix) + sWord + String.fromCharCode(nEnding));
					arrLangs.push(nLang);
				}
			}
		}
	};
	CParagraphSpellChecker.prototype.Add = function(startRun, startInRunPos, endRun, endInRunPos, Word, Lang, Prefix, Ending, apostrophe)
	{
		if (Word.length > 0)
		{
			if ('\'' === Word.charAt(Word.length - 1))
				Word = Word.substr(0, Word.length - 1);
			if ('\'' === Word.charAt(0))
				Word = Word.substr(1);
		}

		if (!this.HaveDictionary(Lang) || !this.IsNeedCheckWord(Word))
			return;
		
		let oElement = new AscWord.CParagraphSpellCheckerElement(startRun, startInRunPos, endRun, endInRunPos, Word, Lang, Prefix, Ending, apostrophe);
		startRun.AddSpellCheckerElement(new AscWord.SpellMarkStart(oElement));
		endRun.AddSpellCheckerElement(new AscWord.SpellMarkEnd(oElement));
		this.Elements.push(oElement);
	};
	CParagraphSpellChecker.prototype.SpellCheckResponse = function(nRecalcId, usrCorrect)
	{
		let oDocumentSpellChecker = this.GetDocumentSpellChecker();
		oDocumentSpellChecker.RemoveWaitingParagraph(this.Paragraph);

		if (nRecalcId !== this.RecalcId)
			return;

		for (let nIndex = 0, nCorrectIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			let oElement = this.Elements[nIndex];
			if (oElement.IsUndefined())
			{
				let isCorrect = false;
				if (oDocumentSpellChecker.IsIgnored(oElement.GetWord()))
				{
					isCorrect = true;
				}
				else if (oElement.GetPrefix() && oElement.GetEnding())
				{
					isCorrect = usrCorrect[nCorrectIndex] || usrCorrect[nCorrectIndex + 1] || usrCorrect[nCorrectIndex + 2] || usrCorrect[nCorrectIndex + 3];
					nCorrectIndex += 3;
				}
				else if (oElement.GetPrefix() || oElement.GetEnding())
				{
					isCorrect = usrCorrect[nCorrectIndex] || usrCorrect[nCorrectIndex + 1];
					nCorrectIndex++;
				}
				else
				{
					isCorrect = usrCorrect[nCorrectIndex];
				}

				if (isCorrect)
					oElement.SetCorrect();
				else
					oElement.SetWrong();

				nCorrectIndex++;
			}
		}

		this.private_UpdateWordsList();
		this.private_UpdateParagraphState();
	};
	CParagraphSpellChecker.prototype.private_UpdateParagraphState = function()
	{
		let oDocumentSpellChecker = this.GetDocumentSpellChecker();

		if (this.GetErrorsCount() > 0)
			oDocumentSpellChecker.AddParagraphWithErrors(this.Paragraph);
		else
			oDocumentSpellChecker.RemoveParagraphWithErrors(this.Paragraph);
	};
	CParagraphSpellChecker.prototype.SuggestResponse = function(nRecalcId, sElementId, usrVariants)
	{
		let oElement = this.Elements[sElementId];
		if (nRecalcId === this.RecalcId && oElement)
		{
			oElement.SetVariants(usrVariants);

			this.private_CheckEASTEGGS(sElementId);

			let sWord = oElement.GetWord();
			let nLang = oElement.GetLang();
			if (undefined !== this.Words[sWord] && undefined !== this.Words[sWord][nLang])
				this.Words[sWord][nLang] = oElement.GetVariants();
		}
	};
	CParagraphSpellChecker.prototype.GetElementByRange = function(oStartPos, oEndPos)
	{
		let oFoundElement = null;
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			let oElement = this.Elements[nIndex];
			if (oElement.CheckIntersection(oStartPos, oEndPos) && oElement.IsWrong())
			{
				if (oFoundElement)
				{
					oFoundElement = null;
					break;
				}
				else
				{
					oFoundElement = oElement;
				}
			}
		}

		return oFoundElement;
	};
	CParagraphSpellChecker.prototype.GetIndexByElement = function(oElement)
	{
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			if (oElement === this.Elements[nIndex])
				return nIndex;
		}

		return -1;
	};
	CParagraphSpellChecker.prototype.CheckVariants = function(oElement)
	{
		if (this.GetDocumentSpellChecker().IsWaitingParagraph(this.Paragraph))
			return;

		if (!oElement.IsWrong())
			return;

		let arrVariants = oElement.GetVariants();
		if (arrVariants)
			return;

		let nLang = oElement.GetLang();
		if (!this.HaveDictionary(nLang))
			return;

		let nElementIndex = this.GetIndexByElement(oElement);
		if (-1 === nElementIndex)
			return;

		editor.SpellCheckApi.spellCheck({
			"type"        : "suggest",
			"ParagraphId" : this.Paragraph.GetId(),
			"RecalcId"    : this.RecalcId,
			"ElementId"   : nElementIndex,
			"usrWords"    : [oElement.GetWord()],
			"usrLang"     : [nLang]
		});
	};
	CParagraphSpellChecker.prototype.Ignore = function(sWord)
	{
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oElement = this.Elements[nIndex];
			if (oElement.IsWrong() && oElement.GetWord() === sWord)
				oElement.SetCorrect();
		}

		if (this.Words[sWord])
			delete this.Words[sWord];

		this.private_UpdateParagraphState();
	};
	CParagraphSpellChecker.prototype.ResetElementsWithCurrentState = function()
	{
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			let oElement = this.Elements[nIndex];
			if (oElement.IsCurrent())
				oElement.SetUndefined();
		}
	};
	/**
	 * Получаем количество элементов проверки орфографии
	 * @returns {Number}
	 */
	CParagraphSpellChecker.prototype.GetElementsCount = function()
	{
		return this.Elements.length;
	};
	/**
	 * Получаем элемент проверки орфографии по номеру
	 * @param nIndex
	 * @returns {AscWord.CParagraphSpellCheckerElement}
	 */
	CParagraphSpellChecker.prototype.GetElement = function(nIndex)
	{
		return (nIndex < 0 || nIndex >= this.Elements.length ? null : this.Elements[nIndex]);
	};
	/**
	 * Приостанавливаем проверку орфографии, если параграф слишком большой
	 * @param {AscWord.CParagraphSpellCheckerCollector} oCollector
	 */
	CParagraphSpellChecker.prototype.Pause = function(oCollector)
	{
		this.Collector = oCollector;
	};
	/**
	 * Проверяем приостановлена ли проверка в данном параграфе
	 * @return {boolean}
	 */
	CParagraphSpellChecker.prototype.IsPaused = function()
	{
		return !!(this.Collector);
	};
	/**
	 * Очищаем остановленное состояние
	 */
	CParagraphSpellChecker.prototype.ClearCollector = function()
	{
		this.Collector = null;
	};
	/**
	 * Нужно ли проверять слово
	 * @param {string} sWord
	 * @returns {boolean}
	 */
	CParagraphSpellChecker.prototype.IsNeedCheckWord = function(sWord)
	{
		if (1 >= sWord.length || ((this.Flags & IGNORE_UPPERCASE) && AscCommon.IsAbbreviation(sWord)))
			return false;

		if (this.Flags & IGNORE_NUMBERS)
		{
			for (var oIterator = sWord.getUnicodeIterator(); oIterator.check(); oIterator.next())
			{
				let nCharCode = oIterator.value();
				if (AscCommon.IsDigit(nCharCode))
					return false;
			}
		}

		return true;
	};
	CParagraphSpellChecker.prototype.HaveDictionary = function(nLang)
	{
		return editor.SpellCheckApi.checkDictionary(nLang);
	};
	CParagraphSpellChecker.prototype.SkipCurrentWord = function()
	{
		return (editor.asc_IsSpellCheckCurrentWord() !== true);
	};
	CParagraphSpellChecker.prototype.GetCurrentPositionInParagraph = function()
	{
		let oCurPos = null;
		if (this.Paragraph.IsThisElementCurrent())
			oCurPos = this.Paragraph.Get_ParaContentPos(false, false);

		return oCurPos;
	};
	CParagraphSpellChecker.prototype.AddCurrentParagraph = function()
	{
		editor.WordControl.m_oLogicDocument.Spelling.AddCurrentParagraph(this.Paragraph);
	};
	CParagraphSpellChecker.prototype.GetErrorsCount = function()
	{
		let nErrorsCount = 0;
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			if (this.Elements[nIndex].IsWrong())
				nErrorsCount++;
		}

		return nErrorsCount;
	};
	CParagraphSpellChecker.prototype.GetCollector = function(isForceFullCheck)
	{
		let oCollector;
		if (this.IsPaused())
		{
			oCollector = this.Collector;
			oCollector.SetFindStart(true);
			oCollector.ResetCheckedCounter();
		}
		else
		{
			oCollector = new AscWord.CParagraphSpellCheckerCollector(this, isForceFullCheck);
			this.Elements = [];
		}

		return oCollector;
	};
	CParagraphSpellChecker.prototype.OnEndCollectingElements = function()
	{
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			let oElement = this.Elements[nIndex];

			let sWord = oElement.GetWord();
			let nLang = oElement.GetLang();

			if (undefined !== this.Words[sWord])
			{
				if (undefined === this.Words[sWord][nLang] || this.Words[sWord].Prefix !== oElement.GetPrefix() || this.Words[sWord].Ending !== oElement.GetEnding())
					oElement.SetUndefined();
				else if (true === this.Words[sWord][nLang])
					oElement.SetCorrect();
				else
					oElement.SetWrong(this.Words[sWord][nLang]);
			}
		}

		this.private_ClearWordsList();
		this.private_UpdateWordsList();
	};
	CParagraphSpellChecker.prototype.private_ClearWordsList = function()
	{
		this.Words = {};
	};
	CParagraphSpellChecker.prototype.private_UpdateWordsList = function()
	{
		for (let nIndex = this.Elements.length - 1; nIndex >= 0; --nIndex)
		{
			let oElement = this.Elements[nIndex];
			let sWord    = oElement.GetWord();
			let nLang    = oElement.GetLang();

			if (oElement.IsCorrect() && !oElement.IsCurrent())
			{
				if (!this.Words[sWord])
				{
					this.Words[sWord] = {
						Prefix : oElement.GetPrefix(),
						Ending : oElement.GetEnding()
					};
				}

				if (undefined === this.Words[sWord][nLang])
					this.Words[sWord][nLang] = true;
			}
			else if (oElement.IsWrong())
			{
				if (!this.Words[sWord])
				{
					this.Words[sWord] = {
						Prefix : oElement.GetPrefix(),
						Ending : oElement.GetEnding()
					};
				}

				if (undefined === this.Words[sWord][nLang])
					this.Words[sWord][nLang] = oElement.GetVariants();
			}
		}

		this.private_ClearMarksForCorrectWords();
	};
	CParagraphSpellChecker.prototype.private_ClearMarksForCorrectWords = function()
	{
		for (let nCount = this.Elements.length, nIndex = nCount - 1; nIndex >= 0; --nIndex)
		{
			let oElement = this.Elements[nIndex];
			if (oElement.IsCorrect() && !oElement.IsCurrent())
			{
				let startRun = oElement.GetStartRun();
				let endRun   = oElement.GetEndRun();

				if (startRun !== endRun)
				{
					startRun.RemoveSpellCheckerElement(oElement);
					endRun.RemoveSpellCheckerElement(oElement);
				}
				else
				{
					endRun.RemoveSpellCheckerElement(oElement);
				}

				this.Elements.splice(nIndex, 1);
			}
		}
	}
	CParagraphSpellChecker.prototype.private_CheckEASTEGGS = function(sId)
	{
		for (let nIndex = 0, nCount = EASTEGGS.length; nIndex < nCount; ++nIndex)
		{
			if (EASTEGGS[nIndex] === this.Elements[sId].Word)
			{
				this.Elements[sId].Variants = EASTEGGS_VARIANTS[nIndex];
				return;
			}
		}
	};
	CParagraphSpellChecker.prototype.GetDocumentSpellChecker = function()
	{
		return editor.WordControl.m_oLogicDocument.Spelling;
	};

	const EASTEGGS          = [];
	const EASTEGGS_VARIANTS = [];


	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CParagraphSpellChecker = CParagraphSpellChecker;

})(window);
