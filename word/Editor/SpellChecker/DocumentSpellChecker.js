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
	const DOCUMENT_SPELLING_MAX_PARAGRAPHS = 50;   // максимальное количество параграфов, которые мы проверяем в таймере
	const DOCUMENT_SPELLING_MAX_ERRORS     = 2000; // максимальное количество ошибок, которые отрисовываются

	const DOCUMENT_SPELLING_EXCEPTIONAL_WORDS = {
		"Teamlab" : true, "teamlab" : true, "OnlyOffice" : true, "ONLYOFFICE" : true, "API" : true
	};


	/**
	 * Проверка орфографии в документе. Здесь будут хранится параграфы, которые надо проверить. Параграфы, в
	 * которых уже есть неверно набранные слова, а также набор слов, которые игнорируются при проверке
	 * @constructor
	 */
	function CDocumentSpellChecker()
	{
		this.Use          = true;
		this.TurnOn       = 1;
		this.ErrorsExceed = false;
		this.Paragraphs  = {}; // Параграфы, в которых есть ошибки в орфографии (объект с ключом - Id параграфа)
		this.Words       = {}; // Слова, которые пользователь решил пропустить(нажал "пропустить все") при проверке орфографии
		this.CheckPara   = {}; // Параграфы, в которых нужно запустить проверку орфографии
		this.CurPara     = {}; // Параграфы, в которых мы не проверили некотырые слова, из-за того что в них стоял курсор
		this.Settings    = new AscCommon.CSpellCheckSettings();

		this.WaitingParagraphs = {}; // Параграфы, которые ждут ответа от сервера

		// Заполним начальный список исключений
		this.Words = DOCUMENT_SPELLING_EXCEPTIONAL_WORDS;
	}

	CDocumentSpellChecker.prototype.TurnOff = function()
	{
		this.TurnOn -= 1;
	};
	CDocumentSpellChecker.prototype.TurnOn = function()
	{
		this.TurnOn += 1;
	};
	CDocumentSpellChecker.prototype.SetSettings = function(oSettings)
	{
		if (!this.Settings.IsEqual(oSettings))
		{
			this.Settings = oSettings.Copy();
			return true;
		}

		return false;
	};
	CDocumentSpellChecker.prototype.GetSettings = function()
	{
		return this.Settings;
	};
	CDocumentSpellChecker.prototype.IsOn = function()
	{
		return (1 === this.TurnOn);
	};
	CDocumentSpellChecker.prototype.IsErrorsExceed = function()
	{
		return this.ErrorsExceed;
	};
	CDocumentSpellChecker.prototype.AddParagraphWithErrors = function(oParagraph)
	{
		this.Paragraphs[oParagraph.GetId()] = oParagraph;
	};
	CDocumentSpellChecker.prototype.RemoveParagraphWithErrors = function(oParagraph)
	{
		delete this.Paragraphs[oParagraph.GetId()];
	};
	CDocumentSpellChecker.prototype.Reset = function()
	{
		this.Paragraphs = {};
		this.CheckPara  = {};
		this.CurPara    = {};

		this.WaitingParagraphs = {};

		this.ErrorsExceed = false;
	};
	CDocumentSpellChecker.prototype.IsIgnored = function(sWord)
	{
		return (undefined !== this.Words[sWord]);
	};
	CDocumentSpellChecker.prototype.AddToIgnore = function(sWord)
	{
		this.Words[sWord] = true;

		for (let sId in this.Paragraphs)
		{
			this.Paragraphs[sId].SpellChecker.Ignore(sWord);
		}
	};
	CDocumentSpellChecker.prototype.AddParagraphToCheck = function(oParagraph)
	{
		this.CheckPara[oParagraph.GetId()] = oParagraph;
	};
	CDocumentSpellChecker.prototype.GetErrorsCount = function()
	{
		let nCount = 0;
		for (let sId in this.Paragraphs)
		{
			let oParagraph = this.Paragraphs[sId];
			nCount += oParagraph.GetSpellChecker().GetErrorsCount();
		}
		return nCount;
	};
	CDocumentSpellChecker.prototype.ContinueSpellCheck = function()
	{
		if (!this.IsOn() || this.IsErrorsExceed())
			return;

		// Пока не обработались предыдущие параграфы, новые не стартуем
		if (!this.private_HaveWaitingParagraphs())
		{
			// Эта функция запускается в таймере, поэтому здесь сразу все параграфы не проверяем, чтобы не было
			// притормоза большого, а запускаем по несколько штук.

			let nCounter = 0;
			for (let sId in this.CheckPara)
			{
				let oParagraph = this.CheckPara[sId];

				if (oParagraph.IsUseInDocument() && !oParagraph.ContinueSpellCheck(false))
					break;

				delete this.CheckPara[sId];
				nCounter++;

				if (nCounter > DOCUMENT_SPELLING_MAX_PARAGRAPHS)
					break;
			}

			this.private_CheckErrorsExceed();
		}


		for (let sId in this.CurPara)
		{
			let oParagraph = this.CurPara[sId];
			delete this.CurPara[sId];
			oParagraph.SpellChecker.ResetElementsWithCurrentState();
			oParagraph.SpellChecker.Check();
		}
	};
	CDocumentSpellChecker.prototype.private_HaveWaitingParagraphs = function()
	{
		for (let sId in this.WaitingParagraphs)
		{
			return true;
		}

		return false;
	};
	CDocumentSpellChecker.prototype.private_CheckErrorsExceed = function()
	{
		// TODO: Послать сообщение
		if (this.GetErrorsCount() > DOCUMENT_SPELLING_MAX_ERRORS)
			this.ErrorsExceed = true;
	};
	CDocumentSpellChecker.prototype.AddCurrentParagraph = function(oParagraph)
	{
		this.CurPara[oParagraph.GetId()] = oParagraph;
	};
	CDocumentSpellChecker.prototype.CheckCurrentParagraph = function()
	{
		for (let sId in this.CheckPara)
		{
			let oParagraph = this.CheckPara[sId];
			oParagraph.ContinueSpellCheck(true);
			delete this.CheckPara[sId];
		}

		for (let sId in this.CurPara)
		{
			let oParagraph = this.CurPara[sId];
			delete this.CurPara[sId];
			oParagraph.SpellChecker.ResetElementsWithCurrentState();
			oParagraph.SpellChecker.Check(undefined, true);
		}
	};
	CDocumentSpellChecker.prototype.AddWaitingParagraph = function(oParagraph, RecalcId, Words, Langs)
	{
		let sParaId = oParagraph.GetId();
		let oInfo   = this.WaitingParagraphs[sParaId];
		if (!oInfo || !oInfo.IsEqual(Words, Langs, RecalcId))
		{
			this.WaitingParagraphs[sParaId] = new CWaitingInfo(Words, Langs, RecalcId);
			return true;
		}

		return false;
	};
	CDocumentSpellChecker.prototype.IsWaitingParagraph = function(oParagraph)
	{
		return (undefined !== this.WaitingParagraphs[oParagraph.GetId()]);
	};
	CDocumentSpellChecker.prototype.RemoveWaitingParagraph = function(oParagraph)
	{
		let sParaId = oParagraph.GetId();
		if (this.WaitingParagraphs[sParaId])
			delete this.WaitingParagraphs[sParaId];
	};
	CDocumentSpellChecker.prototype.RemoveWaitingParagraphById = function(sParaId)
	{
		if (this.WaitingParagraphs[sParaId])
			delete this.WaitingParagraphs[sParaId]
	};

	/**
	 * Класс, содержащий информацию
	 * @constructor
	 */
	function CWaitingInfo(arrWords, arrLangs, nRecalcId)
	{
		this.Words    = arrWords;
		this.Langs    = arrLangs;
		this.RecalcId = nRecalcId;
	}
	CWaitingInfo.prototype.IsEqual = function(arrWords, arrLangs, nRecalcId)
	{
		let nLen = this.Words.length;
		if (this.RecalcId !== nRecalcId
			|| undefined === arrWords
			|| undefined === arrLangs
			|| nLen !== arrWords.length
			|| nLen !== arrLangs.length)
			return false;

		for (let nIndex = 0; nIndex < nLen; nIndex++)
		{
			if (this.Words[nIndex] !== arrWords[nIndex] || this.Langs[nIndex] !== arrLangs[nIndex])
				return false;
		}

		return true;
	};

	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CDocumentSpellChecker = CDocumentSpellChecker;

})(window);

