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
	const UNICODE_TO_SPACE = [];
	UNICODE_TO_SPACE[0x00A0] = 1; // Non-breaking-space
	UNICODE_TO_SPACE[0x2007] = 1; // FIGURE SPACE
	UNICODE_TO_SPACE[0x202F] = 1; // NARROW NO-BREAK SPACE
	UNICODE_TO_SPACE[0x2060] = 1; // WORD JOINER
	UNICODE_TO_SPACE[0x2000] = 1; // EN QUAD
	UNICODE_TO_SPACE[0x2001] = 1; // EM QUAD
	UNICODE_TO_SPACE[0x2004] = 1; // THREE-PER-EM SPACE
	UNICODE_TO_SPACE[0x2006] = 1; // SIX-PER-EM SPACE
	UNICODE_TO_SPACE[0x2008] = 1; // PUNCTUATION SPACE
	UNICODE_TO_SPACE[0x2009] = 1; // THIN SPACE
	UNICODE_TO_SPACE[0x200A] = 1; // HAIR SPACE
	UNICODE_TO_SPACE[0x200B] = 1; // ZERO-WIDTH SPACE
	UNICODE_TO_SPACE[0x202F] = 1; // NARROW NO-BREAK SPACE
	UNICODE_TO_SPACE[0x205F] = 1; // MEDIUM MATHEMATICAL SPACE
	UNICODE_TO_SPACE[0x2060] = 1; // WORD JOINER

	function IsToSpace(nUnicode)
	{
		return (AscCommon.IsSpace(nUnicode) || !!(UNICODE_TO_SPACE[nUnicode]));
	}

	/**
	 * Класс для изменения регистра выделенного текста
	 * @param {Asc.c_oAscChangeTextCaseType} nType
	 * @constructor
	 */
	function CChangeTextCaseEngine(nType)
	{
		this.ChangeType       = nType;
		this.StartSentence    = true;
		this.WordBuffer       = [];
		this.currentSentence  = "";
		this.word             = "";
		this.lineWords        = 0;
		this.SentenceSettings = [];
		this.flag             = 0;
		this.GlobalSettings   = true;
		this.CurrentParagraph = 0;
		this.isAllinTable     = true;
	}
	CChangeTextCaseEngine.prototype.ProcessParagraphs = function(arrParagraphs)
	{
		if (Asc.c_oAscChangeTextCaseType.SentenceCase === this.ChangeType || Asc.c_oAscChangeTextCaseType.CapitalizeWords === this.ChangeType)
			this.CollectWordsAndSentences(arrParagraphs);

		for (let nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
		{
			this.ResetParagraph();

			let oParagraph = arrParagraphs[nIndex];
			let oThis = this;

			if (Asc.c_oAscChangeTextCaseType.SentenceCase === this.ChangeType || Asc.c_oAscChangeTextCaseType.CapitalizeWords === this.ChangeType)
			{
				oParagraph.CheckRunContent(function(oRun)
				{
					oThis.ProcessTextRunInSentenceCase(oRun);
				});
			}
			else
			{
				oParagraph.CheckRunContent(function(oRun)
				{
					oThis.ProcessTextRunInStrictCase(oRun);
				});
			}

			this.FlushWord();
		}
	};
	CChangeTextCaseEngine.prototype.ResetParagraph = function()
	{
		this.StartSentence = true;
		this.WordBuffer    = [];
		this.currentSentence = "";
		this.word = "";
		//this.SentenceSettings = [];
		//this.flag = 0;
		this.lineWords = 0;
	};
	CChangeTextCaseEngine.prototype.FlushWord = function()
	{
		var sCurrentWord  = "";
		var bNeddToChange = true;
		var bIsAllUpper   = true;
		var bIsAllLower   = true;
		var bIsProperName = true;

		for (var nIndex = 0, nCount = this.WordBuffer.length; nIndex < nCount; ++nIndex)
		{
			var nCharCode   = this.WordBuffer[nIndex].Run.GetElement(this.WordBuffer[nIndex].Pos).GetCodePoint();

			var nLowerCode = String.fromCharCode(nCharCode).toLowerCase().charCodeAt(0);
			var nUpperCode = String.fromCharCode(nCharCode).toUpperCase().charCodeAt(0);

			if (nIndex === 0 && nCharCode === nLowerCode)
				bIsProperName = false;
			if (nCharCode === nLowerCode)
				bIsAllUpper = false;
			if (nCharCode === nUpperCode)
			{
				bIsAllLower = false;
				if (nIndex !== 0)
					bIsProperName = false;
			}
			sCurrentWord += String.fromCharCode(nCharCode);
		}

		if (sCurrentWord)
		{
			if (bIsProperName)
			{
				bNeddToChange = false;
			}
			else if (bIsAllUpper)
			{
				bNeddToChange = false;
			}
			else if (bIsAllLower)
			{
				bNeddToChange = false;
			}
			else
			{
				bNeddToChange = true;
			}
		}
		var bFlagForCheck = false;
		var nCaseType = this.ChangeType;
		if (this.SentenceSettings[0] && this.SentenceSettings[0].wordCount)
		{
			for (var nIndex = 0, nCount = this.WordBuffer.length; nIndex < nCount; ++nIndex)
			{
				if (!this.WordBuffer[nIndex].Change)
					continue;
				bFlagForCheck = true;

				var oRun      = this.WordBuffer[nIndex].Run;
				var nInRunPos = this.WordBuffer[nIndex].Pos;
				
				var nCharCode  = oRun.GetElement(nInRunPos).GetCodePoint();
				
				var nLowerCode = String.fromCharCode(nCharCode).toLowerCase().charCodeAt(0);
				var nUpperCode = String.fromCharCode(nCharCode).toUpperCase().charCodeAt(0);

				if (nLowerCode !== nCharCode || nUpperCode !== nCharCode)
				{
					if (nLowerCode === nCharCode
						&& ((Asc.c_oAscChangeTextCaseType.SentenceCase === nCaseType && (this.StartSentence && 0 === nIndex))
						|| Asc.c_oAscChangeTextCaseType.ToggleCase === nCaseType
						|| Asc.c_oAscChangeTextCaseType.UpperCase === nCaseType
						|| (Asc.c_oAscChangeTextCaseType.CapitalizeWords === nCaseType && 0 === nIndex)))
					{
						oRun.AddToContent(nInRunPos, new AscWord.CRunText(nUpperCode), false);
						oRun.RemoveFromContent(nInRunPos + 1, 1, false);
					}
					else if (nUpperCode === nCharCode
						&& (Asc.c_oAscChangeTextCaseType.ToggleCase === nCaseType
							|| Asc.c_oAscChangeTextCaseType.LowerCase === nCaseType
							|| (Asc.c_oAscChangeTextCaseType.CapitalizeWords === nCaseType && 0 !== nIndex && bNeddToChange)
							|| (Asc.c_oAscChangeTextCaseType.CapitalizeWords === nCaseType && this.SentenceSettings[0].allFirst === true && this.SentenceSettings[0].sentenceMistakes === true && !this.SentenceSettings[0].allUpperWithoutFirst && 0 !== nIndex)
							|| (Asc.c_oAscChangeTextCaseType.SentenceCase === nCaseType && this.GlobalSettings === true && this.isAllinTable === false && !(this.StartSentence && 0 === nIndex))
							|| (Asc.c_oAscChangeTextCaseType.SentenceCase === nCaseType && this.isAllinTable === true && this.SentenceSettings[0].allFirst === true && this.SentenceSettings[0].sentenceMistakes === true && !this.SentenceSettings[0].allUpperWithoutFirst && !(this.StartSentence && 0 === nIndex))
							|| (Asc.c_oAscChangeTextCaseType.SentenceCase === nCaseType && bNeddToChange && !(this.StartSentence && 0 === nIndex))
						))
					{
						oRun.AddToContent(nInRunPos, new AscWord.CRunText(nLowerCode), false);
						oRun.RemoveFromContent(nInRunPos + 1, 1, false);
					}
				}
			}
		}
		if (this.WordBuffer.length > 0 && bFlagForCheck && (Asc.c_oAscChangeTextCaseType.CapitalizeWords === nCaseType || Asc.c_oAscChangeTextCaseType.SentenceCase === nCaseType))
		{
			this.lineWords++;
			if (this.SentenceSettings[0] && this.SentenceSettings[0].wordCount)
			{
				if (this.lineWords === this.SentenceSettings[0].wordCount)
				{
					this.SentenceSettings.splice(0, 1);
					this.lineWords = 0;
				}
			}
		}
		if (this.WordBuffer.length > 0)
			this.StartSentence = false;

		this.WordBuffer = [];
	};
	CChangeTextCaseEngine.prototype.AddLetter = function(oRun, nInRunPos, isChange)
	{
		this.WordBuffer.push({
			Run    : oRun,
			Pos    : nInRunPos,
			Change : isChange
		});
	};
	CChangeTextCaseEngine.prototype.SetStartSentence = function(isStart)
	{
		this.StartSentence = isStart;
	};
	CChangeTextCaseEngine.prototype.CheckEachWord = function(sElement)
	{
		var el1 = sElement.slice(1);
		if (sElement[0] === sElement[0].toUpperCase() && el1 === el1.toLowerCase())
		{
			return true;
		}
		if (sElement === sElement.toUpperCase())
		{
			return true;
		}
		if (sElement === sElement.toLowerCase())
		{
			return true;
		}
		return false;
	};
	CChangeTextCaseEngine.prototype.CheckWords = function()
	{
		var sett = {
			allFirst: true,
			sentenceMistakes: true,
			allUpperWithoutFirst: true,
			wordCount: 0
		};
		var wordsInSentece = this.currentSentence.split(/[\-\ \|]/);

		for (var k = 0; k < wordsInSentece.length; k++)
		{
			if (wordsInSentece[k] === "")
			{
				wordsInSentece.splice(k, 1);
				k--;
			}
		}
		sett.wordCount = wordsInSentece.length;
		if (wordsInSentece.length !== 0)
		{
			for (var j = 0; j < wordsInSentece.length; j++)
			{
				if (wordsInSentece[j][0] !== wordsInSentece[j][0].toUpperCase())
				{
					sett.allFirst = false;
				}
				if (!this.CheckEachWord(wordsInSentece[j]))
				{
					sett.sentenceMistakes = false;
				}
				if (this.CurrentParagraph === 0)
				{
					if (sett.allFirst === false || sett.sentenceMistakes === false)
					{
						this.GlobalSettings = false;
					}
				}
				if (wordsInSentece.length > 1 && j >= 1)
				{
					if (wordsInSentece[j] !== wordsInSentece[j].toUpperCase())
					{
						sett.allUpperWithoutFirst = false;
					}
				}
			}
			var elem1 = wordsInSentece[0].slice(1);
			var bElem1IsEmpty;
			if (elem1 === "")
			{
				bElem1IsEmpty = false;
			}
			else if (elem1 != "" && elem1 === elem1.toLowerCase())
			{
				bElem1IsEmpty = true;
			}
			else
			{
				bElem1IsEmpty = false;
			}
			if (!(wordsInSentece[0][0] === wordsInSentece[0][0].toUpperCase() && bElem1IsEmpty))
			{
				sett.allUpperWithoutFirst = false;
			}
			this.SentenceSettings[this.flag] = sett;
			this.flag++;
		}
		this.currentSentence = "";
	};
	CChangeTextCaseEngine.prototype.CheckItemOnCollect = function(oItem, isInSelection)
	{
		if (oItem.IsText())
		{
			if (oItem.IsDot())
			{
				this.currentSentence += this.word;
				this.currentSentence += " ";
				this.word = "";
				this.CheckWords(this);
			}
			else
			{
				if (!oItem.IsPunctuation())
				{
					if (isInSelection)
					{
						let nCharCode  = oItem.GetCharCode();
						let nLowerCode = String.fromCharCode(nCharCode).toLowerCase().charCodeAt(0);
						let nUpperCode = String.fromCharCode(nCharCode).toUpperCase().charCodeAt(0);

						if (IsToSpace(nCharCode))
							this.word += " ";
						if (nLowerCode !== nCharCode || nUpperCode !== nCharCode || oItem.IsNumber())
							this.word += String.fromCharCode(nCharCode);
					}
				}
				else
				{
					this.currentSentence += this.word;
					this.currentSentence += " ";
					this.word = "";
					this.CheckWords(this);
				}
			}
		}
		else
		{
			this.currentSentence += this.word;
			this.currentSentence += " ";
			this.word = "";

			if (!oItem.IsTab() && !oItem.IsSpace())
				this.CheckWords(this);

			if (oItem.IsParaEnd() && this.SentenceSettings.length === 0)
				this.GlobalSettings = false;
		}
	};
	CChangeTextCaseEngine.prototype.CollectWordsAndSentences = function(arrParagraphs)
	{
		for (let nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
		{
			let oParagraph = arrParagraphs[nIndex];
			if (!oParagraph.IsTableCellContent())
				this.isAllinTable = false;

			let oThis = this;
			oParagraph.CheckRunContent(function(oRun, nStartPos, nEndPos)
			{
				if (undefined === nStartPos)
					nStartPos = 0;
				if (undefined === nEndPos)
					nEndPos   = -1;

				if (oRun.IsSelectionUse())
				{
					nStartPos = oRun.GetSelectionStartPos();
					nEndPos   = oRun.GetSelectionEndPos();
				}

				for (let nPos = 0, nCount = oRun.GetElementsCount(); nPos < nCount; ++nPos)
				{
					oThis.CheckItemOnCollect(oRun.GetElement(nPos), nPos >= nStartPos && nPos < nEndPos);
				}
			});
			
			this.CurrentParagraph++;
		}
	};
	CChangeTextCaseEngine.prototype.ProcessTextRunInSentenceCase = function(oRun)
	{
		let nStartPos = 0;
		let nEndPos   = -1;

		if (oRun.IsSelectionUse())
		{
			nStartPos = oRun.GetSelectionStartPos();
			nEndPos   = oRun.GetSelectionEndPos();
		}

		for (let nPos = 0, nCount = oRun.GetElementsCount(); nPos < nCount; ++nPos)
		{
			this.HandleItemInSentenceCase(oRun, nPos, nPos >= nStartPos && nPos < nEndPos);
		}
	};
	CChangeTextCaseEngine.prototype.HandleItemInSentenceCase = function(oRun, nPos, isInSelection)
	{
		let oItem = oRun.GetElement(nPos);
		if (oItem.IsText())
		{
			if (oItem.IsDot())
			{
				this.FlushWord();
				this.SetStartSentence(true);
			}
			else
			{
				if (!oItem.IsPunctuation())
				{
					if (!IsToSpace(oItem.GetCharCode()))
						this.AddLetter(oRun, nPos, isInSelection);
					else
						this.FlushWord();
				}
				else
				{
					this.FlushWord();

					let nCharCode = oItem.GetCharCode();
					if (33 === nCharCode || 63 === nCharCode || 46 === nCharCode)
						this.SetStartSentence(true);
					else
						this.SetStartSentence(false);
				}
			}
		}
		else
		{
			this.FlushWord();

			if (!oItem.IsTab() && !oItem.IsSpace())
				this.SetStartSentence(false);
		}
	};
	CChangeTextCaseEngine.prototype.ProcessTextRunInStrictCase = function(oRun)
	{
		let nStartPos = 0;
		let nEndPos   = -1;

		if (oRun.IsSelectionUse())
		{
			nStartPos = oRun.GetSelectionStartPos();
			nEndPos   = oRun.GetSelectionEndPos();
		}

		for (let nPos = nStartPos; nPos < nEndPos; ++nPos)
		{
			let oItem = oRun.GetElement(nPos);
			if (!oItem.IsText())
				continue;

			let nCharCode  = oItem.GetCharCode();
			let nLowerCode = String.fromCharCode(nCharCode).toLowerCase().charCodeAt(0);
			let nUpperCode = String.fromCharCode(nCharCode).toUpperCase().charCodeAt(0);

			if (nLowerCode !== nCharCode || nUpperCode !== nCharCode)
			{
				if (nLowerCode === nCharCode && (Asc.c_oAscChangeTextCaseType.ToggleCase === this.ChangeType || Asc.c_oAscChangeTextCaseType.UpperCase === this.ChangeType))
				{
					oRun.AddToContent(nPos, new AscWord.CRunText(nUpperCode), false);
					oRun.RemoveFromContent(nPos + 1, 1, false);
				}
				else if (nUpperCode === nCharCode && (Asc.c_oAscChangeTextCaseType.ToggleCase === this.ChangeType || Asc.c_oAscChangeTextCaseType.LowerCase === this.ChangeType))
				{
					oRun.AddToContent(nPos, new AscWord.CRunText(nLowerCode), false);
					oRun.RemoveFromContent(nPos + 1, 1, false);
				}
			}
		}
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CChangeTextCaseEngine = CChangeTextCaseEngine;

})(window);
