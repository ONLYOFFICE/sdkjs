/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	const g_nMaxElements = 1000; // Сколько максимально просматриваем элементов влево

	const AUTOCORRECT_FLAGS_NONE                     = 0x00000000;
	const AUTOCORRECT_FLAGS_ALL                      = 0xFFFFFFFF;
	const AUTOCORRECT_FLAGS_FRENCH_PUNCTUATION       = 0x00000001;
	const AUTOCORRECT_FLAGS_SMART_QUOTES             = 0x00000002;
	const AUTOCORRECT_FLAGS_HYPHEN_WITH_DASH         = 0x00000004;
	const AUTOCORRECT_FLAGS_HYPERLINK                = 0x00000008;
	const AUTOCORRECT_FLAGS_FIRST_LETTER_SENTENCE    = 0x00000010;
	const AUTOCORRECT_FLAGS_NUMBERING                = 0x00000020;
	const AUTOCORRECT_FLAGS_DOUBLE_SPACE_WITH_PERIOD = 0x00000040;

	/**
	 * Класс для выполнения автозамены
	 * @param oRun {AscCommonWord.ParaRun} ран в котором стартовала автозамена
	 * @param nPos {number} позиция, на которой был добавлен последний элемент, с которого стартовала автозамена
	 * @constructor
	 */
	function CRunAutoCorrect(oRun, nPos)
	{
		this.Document   = null;
		this.Paragraph  = null;
		this.Run        = oRun;
		this.ContentPos = null;
		this.Pos        = nPos;
		this.Lang       = 1033;
		this.RunItem    = null;

		this.HistoryActions = 1;
		this.Result         = AUTOCORRECT_FLAGS_NONE;
		this.Flags          = AUTOCORRECT_FLAGS_NONE;

		this.RunElementsBefore = null;
		this.Text              = "";
		this.AsYouType         = false;

		this.Init();
	}
	CRunAutoCorrect.prototype.Init = function()
	{
		let oRun = this.Run;

		if (!oRun)
			return;

		this.Lang = oRun.Get_CompiledPr(false).Lang ? oRun.Get_CompiledPr(false).Lang.Val : 1033;

		this.Paragraph = oRun.GetParagraph();

		if (!this.Paragraph)
			return;

		let oDocument = this.Paragraph.GetLogicDocument();
		if (!oDocument || (!oDocument.IsDocumentEditor() && !oDocument.IsPresentationEditor()))
			return;

		this.Document = oDocument;

		let oContentPos = this.Paragraph.GetPosByElement(oRun);
		if (!oContentPos)
			return;

		oContentPos.Update(this.Pos, oContentPos.GetDepth() + 1);
		this.ContentPos = oContentPos;

		this.RunItem = oRun.GetElement(this.Pos);
	};
	CRunAutoCorrect.prototype.IsValid = function()
	{
		return (!!this.RunItem);
	};
	/**
	 * @param {number} nFlags - флаги, какие автозамены мы пробуем делать
	 * @param {number} [nHistoryActions=1] Автозамене предществовало заданное количество точек в истории
	 * @returns {number}
	 */
	CRunAutoCorrect.prototype.DoAutoCorrect = function(nFlags, nHistoryActions)
	{
		this.Flags          = this.private_CheckFlags(nFlags);
		this.Result         = AUTOCORRECT_FLAGS_NONE;
		this.HistoryActions = (undefined === nHistoryActions || null === nHistoryActions) ? 1 : nHistoryActions;

		if (!this.IsValid() || !this.Flags)
			return this.Result;

		// Чтобы позиция ContentPos была актуальна, отключаем корректировку содержимого параграфа на время выполнения
		// автозамены. Все изменения должны происходить ТОЛЬКО внтури ранов
		this.Paragraph.TurnOffCorrectContent();

		this.private_CheckAsYouType();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_DOUBLE_SPACE_WITH_PERIOD, this.private_ProcessDoubleSpaceWithPeriod))
			return this.private_Return();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_FRENCH_PUNCTUATION, this.private_ProcessFrenchPunctuation))
			return this.private_Return();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_SMART_QUOTES, this.private_ProcessSmartQuotesAutoCorrect))
			return this.private_Return();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_HYPHEN_WITH_DASH, this.private_ProcessHyphenWithDashAutoCorrect))
			return this.private_Return();

		if (!this.private_CollectTextBefore())
			return this.private_Return();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_HYPHEN_WITH_DASH, this.private_ProcessSpaceHyphenWithDashAutoCorrect))
			return this.private_Return();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_HYPERLINK, this.private_ProcessHyperlinkAutoCorrect))
			return this.private_Return();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_FIRST_LETTER_SENTENCE, this.private_ProcessCapitalizeFirstLetterOfSentencesAutoCorrect))
			return this.private_Return();

		if (!this.private_ProcessAutoCorrect(AUTOCORRECT_FLAGS_NUMBERING, this.private_ProcessNumbering))
			return this.private_Return();

		return this.private_Return();
	};
	CRunAutoCorrect.prototype.private_CheckFlags = function(nFlags)
	{
		// Если автозамены будем включать в формах при каких-либо условиях, тогда нужно проверять, что
		// выполнение автозамены не выходит за пределы формы
		if (this.Run.GetParentForm())
			return AUTOCORRECT_FLAGS_NONE;

		return nFlags;
	};
	CRunAutoCorrect.prototype.private_Return = function()
	{
		this.Paragraph.TurnOnCorrectContent();
		return this.Result;
	};
	CRunAutoCorrect.prototype.private_IsDocumentLocked = function()
	{
		return (0 === this.HistoryActions
			&& this.Result === AUTOCORRECT_FLAGS_NONE
			&& this.Document.IsSelectionLocked(AscCommon.changestype_None, {
			Type      : AscCommon.changestype_2_ElementsArray_and_Type,
			Elements  : [this.Paragraph],
			CheckType : AscCommon.changestype_Paragraph_Properties
		}, true, false));
	};
	CRunAutoCorrect.prototype.private_UpdatePos = function()
	{
		let oRun = this.Run;
		if (oRun.GetElement(this.Pos) === this.RunItem)
		{
			oRun.State.ContentPos = this.Pos + 1;
			return;
		}

		for (let nPos = 0, nCount = oRun.GetElementsCount(); nPos < nCount; ++nPos)
		{
			if (oRun.GetElement(nPos) === this.RunItem)
			{
				oRun.State.ContentPos = nPos + 1;

				this.Pos = nPos;
				this.ContentPos.Update(nPos, this.ContentPos.GetDepth());
				return;
			}
		}

		this.RunItem = null;
	};
	CRunAutoCorrect.prototype.private_CollectTextBefore = function()
	{
		let oRunElementsBefore = new CParagraphRunElements(this.ContentPos, g_nMaxElements, [para_Text], false);
		oRunElementsBefore.SetBreakOnBadType(true);
		oRunElementsBefore.SetBreakOnDifferentClass(true);
		oRunElementsBefore.SetSaveContentPositions(true);

		this.Paragraph.GetPrevRunElements(oRunElementsBefore);

		var arrElements = oRunElementsBefore.GetElements();
		if (arrElements.length <= 0)
			return false;

		var sText = "";
		for (let nIndex = 0, nCount = arrElements.length; nIndex < nCount; ++nIndex)
		{
			if (para_Text !== arrElements[nCount - 1 - nIndex].Type)
				return false;

			sText += String.fromCodePoint(arrElements[nCount - 1 - nIndex].Value);
		}

		this.RunElementsBefore = oRunElementsBefore;
		this.Text              = sText;

		return true;
	};
	CRunAutoCorrect.prototype.private_CheckAsYouType = function()
	{
		let oRunElementsBefore = new CParagraphRunElements(this.ContentPos, 1, [para_Text], false);
		oRunElementsBefore.SetBreakOnBadType(true);
		oRunElementsBefore.SetBreakOnDifferentClass(true);
		this.Paragraph.GetPrevRunElements(oRunElementsBefore);

		this.AsYouType = false;

		let arrElements = oRunElementsBefore.GetElements();

		var oHistory = this.Document.GetHistory();
		if (arrElements.length > 0 && oHistory.CheckAsYouTypeAutoCorrect)
			this.AsYouType = oHistory.CheckAsYouTypeAutoCorrect(arrElements[0], this.HistoryActions);
	};
	CRunAutoCorrect.prototype.private_ProcessAutoCorrect = function(nType, pFunc)
	{
		if (this.Flags & nType && pFunc.call(this))
			this.Result |= nType;

		this.private_UpdatePos();

		return this.IsValid();
	};
	/**
	 * Производим автозамену для замены двух пробелов знаком точки
	 * @returns {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessDoubleSpaceWithPeriod = function()
	{
		let oDocument   = this.Document;
		let oRun        = this.Run;
		let oContentPos = this.ContentPos;
		let oParagraph  = this.Paragraph;

		if (!oDocument.IsAutoCorrectDoubleSpaceWithPeriod())
			return false;

		if (!this.RunItem.IsSpace())
			return false;

		var oRunElementsBefore = new CParagraphRunElements(oContentPos, 2, null, false);
		oRunElementsBefore.SetSaveContentPositions(true);
		oParagraph.GetPrevRunElements(oRunElementsBefore);
		var arrElements = oRunElementsBefore.GetElements();
		var oHistory    = oDocument.GetHistory();
		if (2 !== arrElements.length
			|| !arrElements[0].IsSpace()
			|| !this.private_CheckPrevSymbolForDoubleSpaceWithDot(arrElements[1])
			|| !oHistory.CheckAsYouTypeAutoCorrect(arrElements[0], 1, 500))
			return false;

		if (this.private_IsDocumentLocked())
			return false;

		oDocument.StartAction(AscDFH.historydescription_Document_AutoCorrectHyphensWithDash);

		var oDot = new AscWord.CRunText(46);
		oRun.AddToContent(this.Pos, oDot);
		oParagraph.RemoveRunElement(oRunElementsBefore.GetContentPositions()[0]);

		oDocument.Recalculate();
		oDocument.FinalizeAction();

		return true;
	};
	CRunAutoCorrect.prototype.private_CheckPrevSymbolForDoubleSpaceWithDot = function(oItem)
	{
		return (oItem.IsText()
			&& (!oItem.IsPunctuation()
				|| 0x23 === oItem.Value
				|| 0x24 === oItem.Value
				|| 0x25 === oItem.Value
				|| 0x40 === oItem.Value));
	};
	/**
	 * Производим автозамену для французской пунктуации
	 * @returns {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessFrenchPunctuation = function()
	{
		let oDocument   = this.Document;
		let oRun        = this.Run;
		let oContentPos = this.ContentPos;
		let oParagraph  = this.Paragraph;
		let nLang       = this.Lang;
		let oRunItem    = this.RunItem;

		if (!oDocument.IsAutoCorrectFrenchPunctuation())
			return false;

		if (!(para_Text === oRunItem.Type && (1036 === nLang && (0x003A === oRunItem.Value || 0x003B === oRunItem.Value || 0x003F === oRunItem.Value || 0x0021 === oRunItem.Value))))
			return false;

		var oRunElementsBefore = new CParagraphRunElements(oContentPos, 3, null, false);
		oRunElementsBefore.SetSaveContentPositions(true);
		oParagraph.GetPrevRunElements(oRunElementsBefore);
		var arrElements = oRunElementsBefore.GetElements();

		if ((arrElements.length > 0 &&
			para_Text === arrElements[0].Type
			&& (33 === arrElements[0].Value
				|| 58 === arrElements[0].Value
				|| 59 === arrElements[0].Value
				|| 63 === arrElements[0].Value))
			|| (arrElements.length >= 3
				&& para_Space === arrElements[0].Type
				&& para_Space === arrElements[1].Type
				&& para_Space === arrElements[2].Type))
			return false;

		if (this.private_IsDocumentLocked())
			return false;

		oDocument.StartAction(AscDFH.historydescription_Document_AutoCorrectCommon);

		oRun.AddToContent(this.Pos, new AscWord.CRunText(0x00A0));
		oRun.State.ContentPos = this.Pos + 2;

		if (arrElements.length >= 1 && (para_Space === arrElements[0].Type || (para_Text === arrElements[0].Type && arrElements[0].IsNBSP())))
		{
			if (oParagraph.RemoveRunElement(oRunElementsBefore.GetContentPositions()[0])
				&& (arrElements.length >= 2 && (para_Space === arrElements[1].Type || (para_Text === arrElements[1].Type && arrElements[1].IsNBSP()))))
			{
				oParagraph.RemoveRunElement(oRunElementsBefore.GetContentPositions()[1]);
			}
		}

		oDocument.FinalizeAction();

		return true;
	};
	/**
	 * Производим автозамену для умных кавычек
	 * @returns {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessSmartQuotesAutoCorrect = function()
	{
		let oDocument   = this.Document;
		let oContentPos = this.ContentPos;
		let oParagraph  = this.Paragraph;
		let nLang       = this.Lang;
		let oRunItem    = this.RunItem;

		if (!oDocument.IsAutoCorrectSmartQuotes())
			return false;

		if (!(para_Text === oRunItem.Type && (34 === oRunItem.Value || 39 === oRunItem.Value)))
			return false;

		var isOpenQuote   = true;
		var isDoubleQuote = 34 === oRunItem.Value;

		var oRunElementsBefore = new CParagraphRunElements(oContentPos, 1, null, false);
		oParagraph.GetPrevRunElements(oRunElementsBefore);
		var arrElements = oRunElementsBefore.GetElements();
		if (arrElements.length > 0)
			isOpenQuote = this.private_IsOpenQuoteAfter(arrElements[0]);

		if (!isDoubleQuote && (1050 === nLang || 1060 === nLang))
			return false;

		if (this.private_IsDocumentLocked())
			return false;

		oDocument.StartAction(AscDFH.historydescription_Document_AutoCorrectSmartQuotes);

		this.private_ReplaceSmartQuotes(nLang, isDoubleQuote, isOpenQuote);
		this.Run.State.ContentPos = this.Pos + 1;

		oDocument.FinalizeAction();

		this.RunItem = this.Run.GetElement(this.Pos);

		return true;
	};
	CRunAutoCorrect.prototype.private_IsOpenQuoteAfter = function(oPrevElement)
	{
		// nbsp - – − ( [ {
		return (!oPrevElement.IsText()
			|| oPrevElement.IsNBSP()
			|| 0x002D === oPrevElement.Value
			|| 0x2013 === oPrevElement.Value
			|| 0x2212 === oPrevElement.Value
			|| 0x0028 === oPrevElement.Value
			|| 0x005B === oPrevElement.Value
			|| 0x007B === oPrevElement.Value);
	};
	CRunAutoCorrect.prototype.private_ReplaceSmartQuotes = function(nLang, isDoubleQuote, isOpenQuote)
	{
		this.Run.RemoveFromContent(this.Pos, 1);
		if (isDoubleQuote)
		{
			switch (nLang)
			{
				case 1029:
				case 1031:
				case 1039:
				case 1050:
				case 1051:
				case 1061:
				{
					// „text“
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x201E : 0x201C));
					break;
				}
				case 1038:
				case 1045:
				case 1048:
				case 1062:
				{
					// „text”
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x201E : 0x201D));
					break;
				}
				case 1030:
				case 1035:
				case 1053:
				{
					// ”text”
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(0x201D));
					break;
				}
				case 1049:
				{
					// «text»
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x00AB : 0x00BB));
					break;
				}
				case 1060:
				{
					// »text«
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x00BB : 0x00AB));
					break;
				}
				case 1036:
				{
					// « text »
					if (isOpenQuote)
					{
						this.Run.AddToContent(this.Pos, new AscWord.CRunText(0x00AB));
						this.Run.AddToContent(this.Pos + 1, new AscWord.CRunText(0x00A0));
					}
					else
					{
						this.Run.AddToContent(this.Pos, new AscWord.CRunText(0x00A0));
						this.Run.AddToContent(this.Pos + 1, new AscWord.CRunText(0x00BB));
					}

					this.Pos++;

					break;
				}

				default:
				{
					// “text”
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x201C : 0x201D));
					break;
				}
			}
		}
		else
		{
			switch (nLang)
			{
				case 1029:
				case 1031:
				case 1039:
				case 1051:
				{
					// ‚text‘
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x201A : 0x2018));
					break;
				}
				case 1048:
				{
					// ‚text’
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x201A : 0x2019));
					break;
				}
				case 1030:
				case 1035:
				case 1038:
				case 1053:
				case 1061:
				{
					// ’text’
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(0x2019));
					break;
				}
				default:
				{
					// ‘text’
					this.Run.AddToContent(this.Pos, new AscWord.CRunText(isOpenQuote ? 0x2018 : 0x2019));
					break;
				}
			}
		}
	};
	/**
	 * Производим автозамену замены двух дефисов длинным тире
	 * @returns {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessHyphenWithDashAutoCorrect = function()
	{
		let oDocument   = this.Document;
		let oContentPos = this.ContentPos;
		let oParagraph  = this.Paragraph;
		let oRunItem    = this.RunItem;
		let oRun        = this.Run;

		if (!oDocument.IsAutoCorrectHyphensWithDash())
			return false;

		if (!(para_Text === oRunItem.Type && 45 === oRunItem.Value))
			return false;

		var oRunElementsBefore = new CParagraphRunElements(oContentPos, 1, null, false);
		oRunElementsBefore.SetSaveContentPositions(true);
		oParagraph.GetPrevRunElements(oRunElementsBefore);
		var arrElements = oRunElementsBefore.GetElements();
		if (arrElements.length > 0 && para_Text === arrElements[0].Type && 45 === arrElements[0].Value)
		{
			if (this.private_IsDocumentLocked())
				return false;

			oDocument.StartAction(AscDFH.historydescription_Document_AutoCorrectHyphensWithDash);

			var oDash = new AscWord.CRunText(8212);
			oRun.AddToContent(this.Pos + 1, oDash);
			var oStartPos = oRunElementsBefore.GetContentPositions()[0];
			var oEndPos   = oContentPos;
			oContentPos.Update(this.Pos + 1, oContentPos.GetDepth());

			oParagraph.RemoveSelection();
			oParagraph.SetSelectionUse(true);
			oParagraph.SetSelectionContentPos(oStartPos, oEndPos, false);
			oParagraph.Remove(1);
			oParagraph.RemoveSelection();

			oDocument.Recalculate();
			oDocument.FinalizeAction();

			this.RunItem = oDash;

			return true;
		}

		return false;
	};
	/**
	 * Производим автозамену hyphen на dash в случаях <space-hyphen> или <space-hyphen-space>
	 * @returns {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessSpaceHyphenWithDashAutoCorrect = function()
	{
		let oDocument          = this.Document;
		let oParagraph         = this.Paragraph;
		let oRunElementsBefore = this.RunElementsBefore;

		if (!this.AsYouType)
			return false;

		if (!oDocument.IsAutoCorrectHyphensWithDash())
			return false;

		var arrElements    = oRunElementsBefore.GetElements();
		var nElementsCount = arrElements.length;
		if (nElementsCount <= 0)
			return false;

		for (var nIndex = 0; nIndex < nElementsCount; ++nIndex)
		{
			var oItem = arrElements[nIndex];
			if (!oItem.IsLetter() && (nIndex !== nElementsCount - 1 || !oItem.IsHyphen()))
				return false;
		}

		var oChangePos = null;
		if (para_Text === arrElements[nElementsCount - 1].Type && 45 === arrElements[nElementsCount - 1].Value)
		{
			if (arrElements.length > 1 && para_Text === arrElements[nElementsCount - 2].Type && 45 !== arrElements[nElementsCount - 2].Value)
			{
				var oTempRunElementsBefore = new CParagraphRunElements(oRunElementsBefore.GetContentPositions()[nElementsCount - 1], 1, null, false);
				oTempRunElementsBefore.SetSaveContentPositions(true);
				oParagraph.GetPrevRunElements(oTempRunElementsBefore);
				arrElements = oTempRunElementsBefore.GetElements();
				if (arrElements.length > 0 && para_Space === arrElements[0].Type)
					oChangePos = oRunElementsBefore.GetContentPositions()[nElementsCount - 1];
			}
		}
		else
		{
			var oTempRunElementsBefore = new CParagraphRunElements(oRunElementsBefore.GetContentPositions()[nElementsCount - 1], 3, null, false);
			oTempRunElementsBefore.SetSaveContentPositions(true);
			oParagraph.GetPrevRunElements(oTempRunElementsBefore);
			arrElements = oTempRunElementsBefore.GetElements();
			if (3 === arrElements.length
				&& para_Space === arrElements[0].Type
				&& para_Text === arrElements[1].Type && 45 === arrElements[1].Value
				&& para_Space === arrElements[2].Type)
				oChangePos = oTempRunElementsBefore.GetContentPositions()[1];
		}

		if (oChangePos)
		{
			var nInRunPos = oChangePos.Get(oChangePos.GetDepth());
			oChangePos.DecreaseDepth(1);
			var oRun = oParagraph.GetClassByPos(oChangePos);

			if (oRun instanceof ParaRun && oRun.Content.length > nInRunPos)
			{
				if (this.private_IsDocumentLocked())
					return false;

				oDocument.StartAction(AscDFH.historydescription_Document_AutoCorrectHyphensWithDash);
				oRun.RemoveFromContent(nInRunPos, 1);
				oRun.AddToContent(nInRunPos, new AscWord.CRunText(8211));
				oDocument.Recalculate();
				oDocument.FinalizeAction();
				return true;
			}
		}

		return false;
	};
	/**
	 * Производим автозаменку для гиперссылок
	 * @returns {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessHyperlinkAutoCorrect = function()
	{
		let oDocument          = this.Document;
		let sText              = this.Text;
		let oParagraph         = this.Paragraph;
		let oContentPos        = this.ContentPos;
		let oRunElementsBefore = this.RunElementsBefore;
		let oRun               = this.Run;

		if (!oDocument.IsAutoCorrectHyperlinks())
			return false;

		var isPresentation = oDocument.IsPresentationEditor();

		if (oRun.IsInHyperlink())
			return false;

		if (AscCommon.rx_allowedProtocols.test(sText) || /^(www.)|@/i.test(sText))
		{
			// Удаляем концевые пробелы и переводы строки перед проверкой гиперссылок
			sText = sText.replace(/\s+$/, '');

			var nTypeHyper = AscCommon.getUrlType(sText);
			if (AscCommon.c_oAscUrlType.Invalid !== nTypeHyper)
			{
				if (!isPresentation && this.private_IsDocumentLocked())
					return false;

				oDocument.StartAction(AscDFH.historydescription_Document_AutomaticListAsType);
				var oTopElement;

				if (isPresentation)
				{
					var oParentContent = oParagraph.Parent;
					var oTable         = oParentContent.IsInTable(true);
					if (oTable)
					{
						oTopElement = oTable;
					}
					else
					{
						oTopElement = oParentContent;
					}
				}
				else
				{
					oTopElement = oDocument;
				}

				var arrContentPosition = oRunElementsBefore.GetContentPositions();
				var oStartPos          = arrContentPosition.length > 0 ? arrContentPosition[arrContentPosition.length - 1] : oRunElementsBefore.CurContentPos;
				var oEndPos            = oContentPos;
				oContentPos.Update(this.Pos, oContentPos.GetDepth());


				var oDocPos = [{Class : oRun, Position : this.Pos + 1}];
				oRun.GetDocumentPositionFromObject(oDocPos);
				oDocument.TrackDocumentPositions([oDocPos]);


				oParagraph.RemoveSelection();
				oParagraph.SetSelectionUse(true);
				oParagraph.SetSelectionContentPos(oStartPos, oEndPos, false);
				oParagraph.AddHyperlink(new Asc.CHyperlinkProperty({Value : AscCommon.prepareUrl(sText, nTypeHyper)}));
				oParagraph.RemoveSelection();

				oDocument.RefreshDocumentPositions([oDocPos]);
				oTopElement.SetContentPosition(oDocPos, 0, 0);
				oDocument.Recalculate();
				oDocument.FinalizeAction();

				// TODO: Надо обновить позицию Run + Position
				this.RunItem = null;

				return true;
			}
		}

		return false;
	};
	/**
	 * Производим автозамену для первого символа в предложении
	 * @return {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessCapitalizeFirstLetterOfSentencesAutoCorrect = function()
	{
		let oDocument          = this.Document;
		let sText              = this.Text;
		let oParagraph         = this.Paragraph;
		let oRunElementsBefore = this.RunElementsBefore;
		let nLang              = this.Lang;

		if (!this.AsYouType)
			return false;

		if (!oDocument.IsAutoCorrectFirstLetterOfSentences())
			return false;

		if (oRunElementsBefore.IsEnd()
			&& oParagraph.IsTableCellContent()
			&& !oDocument.IsAutoCorrectFirstLetterOfCells())
		{
			return false;
		}

		if ("www" === sText || "http" === sText || "https" === sText)
			return false;

		var arrElements = oRunElementsBefore.GetElements();
		if (arrElements.length <= 0)
			return false;

		for (var nIndex = 0, nCount = arrElements.length; nIndex < nCount; ++nIndex)
		{
			if (para_Text !== arrElements[nIndex].Type)
				return false;

			var sTemp = String.fromCharCode(arrElements[nIndex].Value);
			if (sTemp.toUpperCase() === sTemp && !arrElements[nIndex].IsNBSP())
				return false;
		}

		var arrContentPos = oRunElementsBefore.GetContentPositions();
		if (arrContentPos.length <= 0)
			return false;

		// Запоминаем позицию для автозамены
		var oAutoCorrectContentPos = arrContentPos[arrContentPos.length - 1];

		var oNextRunElementsBefore = new CParagraphRunElements(oAutoCorrectContentPos, g_nMaxElements, [para_Space, para_Tab], false);
		oNextRunElementsBefore.SetBreakOnBadType(true);
		oNextRunElementsBefore.SetBreakOnDifferentClass(true);
		oNextRunElementsBefore.SetSaveContentPositions(true);
		oParagraph.GetPrevRunElements(oNextRunElementsBefore);

		if (!oNextRunElementsBefore.IsEnd())
		{
			var oNextContentPos;

			arrContentPos = oNextRunElementsBefore.GetContentPositions();
			if (arrContentPos.length <= 0)
				oNextContentPos = oAutoCorrectContentPos;
			else
				oNextContentPos = arrContentPos[arrContentPos.length - 1];

			var oRunElements = new CParagraphRunElements(oNextContentPos, 1, null, true);
			oRunElements.SetSaveContentPositions(true);
			oParagraph.GetPrevRunElements(oRunElements);

			// TODO: Надо проверить окончание предложения со скобками, возможно надо проверять два последних символа
			if (oRunElements.Elements.length > 0
				&& !oRunElements.Elements[0].IsDot()
				&& !oRunElements.Elements[0].IsExclamationMark()
				&& !oRunElements.Elements[0].IsQuestionMark())
				return false;

			// Проверяем исключения
			if (1 === oRunElements.Elements.length)
			{
				let autoCorrectSettings = oDocument.GetAutoCorrectSettings();

				var nExceptionMaxLen = autoCorrectSettings.GetFirstLetterExceptionsMaxLen() + 1;
				var oDotContentPos   = oRunElements.GetContentPositions()[0];
				oRunElements         = new CParagraphRunElements(oDotContentPos, nExceptionMaxLen, null, false);
				oParagraph.GetPrevRunElements(oRunElements);

				arrElements         = oRunElements.GetElements();
				var sCheckException = "";
				for (var nIndex = 0, nCount = arrElements.length; nIndex < nCount; ++nIndex)
				{
					var oElement = arrElements[nIndex];

					if (!oElement.IsLetter())
						break;

					sCheckException = String.fromCharCode(oElement.Value) + sCheckException;
				}

				if (autoCorrectSettings.CheckFirstLetterException(sCheckException, nLang))
					return false;
			}
		}

		// Если мы дошли до этого момента, значит можно производить автозамену
		var nDepth = oAutoCorrectContentPos.GetDepth();
		if (nDepth <= 0)
			return false;

		var nInRunPos = oAutoCorrectContentPos.Get(nDepth);
		oAutoCorrectContentPos.DecreaseDepth(1);
		var oRun = oParagraph.GetElementByPos(oAutoCorrectContentPos);
		if (!oRun || !(oRun instanceof ParaRun))
			return false;

		var oItem = oRun.GetElement(nInRunPos);
		if (!oItem || oItem.Type !== para_Text)
			return false;

		if (this.private_IsDocumentLocked())
			return false;

		oDocument.StartAction(AscDFH.historydescription_Document_AutoCorrectFirstLetterOfSentence);

		var oNewItem = new AscWord.CRunText(String.fromCharCode(oItem.Value).toUpperCase().charCodeAt(0));
		oRun.RemoveFromContent(nInRunPos, 1, true);
		oRun.AddToContent(nInRunPos, oNewItem, true);

		oDocument.Recalculate();
		oDocument.FinalizeAction();

		return true;
	};
	/**
	 * Производим автозамену для первого символа в предложении
	 * @return {boolean}
	 */
	CRunAutoCorrect.prototype.private_ProcessNumbering = function()
	{
		let oParagraph         = this.Paragraph;
		let oRunElementsBefore = this.RunElementsBefore;

		if ((oParagraph.bFromDocument && oParagraph.GetNumPr())
			|| (!oParagraph.bFromDocument && !oParagraph.PresentationPr.Bullet.IsNone())
			|| !oRunElementsBefore.IsEnd())
			return false;

		if (oParagraph.bFromDocument)
			return this.private_ProcessNumberingInDocument();
		else
			return this.private_ProcessNumberingInPresentation();
	};
	CRunAutoCorrect.prototype.private_ProcessNumberingInDocument = function()
	{
		let oDocument   = this.Document;
		let oParagraph  = this.Paragraph;
		let oContentPos = this.ContentPos;

		if (this.private_IsDocumentLocked())
			return false;

		oDocument.StartAction(AscDFH.historydescription_Document_AutomaticListAsType);

		let oNumPr = this.private_GetSuitableNumPr();
		if (oNumPr)
		{
			var oStartPos = oParagraph.GetStartPos();
			var oEndPos   = oContentPos;
			oContentPos.Update(this.Pos + 1, oContentPos.GetDepth());

			oParagraph.RemoveSelection();
			oParagraph.SetSelectionUse(true);
			oParagraph.SetSelectionContentPos(oStartPos, oEndPos, false);
			oParagraph.Remove(1);
			oParagraph.RemoveSelection();
			oParagraph.MoveCursorToStartPos(false);

			oParagraph.ApplyNumPr(oNumPr.NumId, oNumPr.Lvl);

			oDocument.Recalculate();
		}

		oDocument.FinalizeAction();
		return (!!oNumPr);
	};
	CRunAutoCorrect.prototype.private_ProcessNumberingInPresentation = function()
	{
		let oDocument   = this.Document;
		let oParagraph  = this.Paragraph;
		let oContentPos = this.ContentPos;

		var oBullet = null;
		if (oDocument.IsAutomaticBulletedLists())
		{
			oBullet = this.private_GetSuitablePresentationBulletForAutoCorrect(this.Text);
		}
		if (!oBullet && oDocument.IsAutomaticNumberedLists())
		{
			oBullet = this.private_GetSuitablePresentationNumberingForAutoCorrect(this.Text);
		}

		if (oBullet)
		{
			oDocument.StartAction(AscDFH.historydescription_Document_AutomaticListAsType);
			var oStartPos = oParagraph.GetStartPos();
			var oEndPos   = oContentPos;
			oContentPos.Update(this.Pos + 1, oContentPos.GetDepth());

			oParagraph.RemoveSelection();
			oParagraph.SetSelectionUse(true);
			oParagraph.SetSelectionContentPos(oStartPos, oEndPos, false);
			oParagraph.Remove(1);
			oParagraph.RemoveSelection();
			oParagraph.MoveCursorToStartPos(false);
			oParagraph.Add_PresentationNumbering(oBullet);
			oDocument.FinalizeAction();

			return true;
		}

		return false;
	};
	CRunAutoCorrect.prototype.private_GetSuitableNumPr = function()
	{
		let oParagraph = this.Paragraph;
		let oDocument  = this.Document;

		let oNumPr = null;

		var oPrevNumPr = null;
		var oPrevParagraph = oParagraph.Get_DocumentPrev();
		if (oPrevParagraph && oPrevParagraph.IsParagraph())
			oPrevNumPr = oPrevParagraph.GetNumPr();

		if (oDocument.IsAutomaticBulletedLists())
		{
			var oNumLvl = this.private_GetSuitableBulletedLvlForAutoCorrect(this.Text);
			if (oNumLvl)
			{
				if (oPrevNumPr)
				{
					var oPrevNumLvl = oDocument.GetNumbering().GetNum(oPrevNumPr.NumId).GetLvl(oPrevNumPr.Lvl);
					if (oPrevNumLvl.IsSimilar(oNumLvl))
					{
						oNumPr = new CNumPr(oPrevNumPr.NumId, oPrevNumPr.Lvl);
					}
				}

				if (!oNumPr)
				{
					var oNum = oDocument.GetNumbering().CreateNum();
					oNum.CreateDefault(c_oAscMultiLevelNumbering.Bullet);
					oNum.SetLvl(oNumLvl, 0);
					oNumPr = new CNumPr(oNum.GetId(), 0);
				}
			}
		}

		if (oDocument.IsAutomaticNumberedLists())
		{
			var arrResult = this.private_GetSuitableNumberedLvlForAutoCorrect(this.Text);

			if (arrResult && arrResult.length > 0 && arrResult.length <= 9)
			{
				if (oPrevNumPr)
				{
					var isAdd      = false;
					var nResultLvL = oPrevNumPr.Lvl;

					var oResult = arrResult[arrResult.length - 1];
					if (oResult && -1 !== oResult.Value && oResult.Lvl)
					{
						var oNumInfo = oPrevParagraph.Parent.CalculateNumberingValues(oPrevParagraph, oPrevNumPr);
						var oPrevNum = oDocument.GetNumbering().GetNum(oPrevNumPr.NumId);
						var nPrevLvl = oPrevNumPr.Lvl;

						for (var nLvl = nPrevLvl; nLvl >= 0; --nLvl)
						{
							var oPrevNumLvl = oPrevNum.GetLvl(nLvl);
							if (oPrevNumLvl.IsSimilar(oResult.Lvl))
							{
								if (oResult.Value > oNumInfo[nLvl] && oResult.Value <= oNumInfo[nLvl] + 2 && arrResult.length - 1 >= nLvl)
								{
									var isCheckPrevLvls = true;
									for (var nLvl2 = 0; nLvl2 < nLvl; ++nLvl2)
									{
										if (arrResult[nLvl2].Value !== oNumInfo[nLvl2])
										{
											isCheckPrevLvls = false;
											break;
										}
									}

									if (isCheckPrevLvls)
									{
										isAdd      = true;
										nResultLvL = nLvl;
										break;
									}
								}
							}
						}

						if (!isAdd)
						{
							oResult.Lvl.ResetNumberedText(oPrevNumPr.Lvl);

							var oPrevNumLvl = oDocument.GetNumbering().GetNum(oPrevNumPr.NumId).GetLvl(oPrevNumPr.Lvl);
							if (oPrevNumLvl.IsSimilar(oResult.Lvl))
							{
								var oNumInfo = oPrevParagraph.Parent.CalculateNumberingValues(oPrevParagraph, oPrevNumPr);
								if (oResult.Value > oNumInfo[oPrevNumPr.Lvl] && oResult.Value <= oNumInfo[oPrevNumPr.Lvl] + 2)
									isAdd = true;
							}
						}

					}

					if (isAdd)
						oNumPr = new CNumPr(oPrevNumPr.NumId, nResultLvL);
				}
				else
				{
					var isCreateNew = true;
					for (var nIndex = 0, nCount = arrResult.length; nIndex < nCount; ++nIndex)
					{
						var oResult = arrResult[nIndex];
						if (!oResult || 1 !== oResult.Value || !oResult.Lvl)
						{
							isCreateNew = false;
							break;
						}
					}

					if (isCreateNew)
					{
						var oNum = oDocument.GetNumbering().CreateNum();
						oNum.CreateDefault(c_oAscMultiLevelNumbering.Numbered);
						for (var iLvl = 0, nCount = arrResult.length; iLvl < nCount; ++iLvl)
						{
							let oldLvl = oNum.GetLvl(iLvl);
							let newLvl = arrResult[iLvl].Lvl;
							newLvl.SetParaPr(oldLvl.GetParaPr());
							oNum.SetLvl(newLvl, iLvl);
						}

						oNumPr = new CNumPr(oNum.GetId(), arrResult.length - 1);
					}
				}
			}
		}

		return oNumPr;
	};
	/**
	 * Подбираем подходящий маркированный список
	 * @param sText {string}
	 * @returns {?CNumberingLvl}
	 */
	CRunAutoCorrect.prototype.private_GetSuitableBulletedLvlForAutoCorrect = function(sText)
	{
		var oNumberingLvl = new CNumberingLvl();
		oNumberingLvl.InitDefault(0, c_oAscMultiLevelNumbering.Bullet);

		if ('*' === sText)
		{
			var oTextPr = new CTextPr();
			oTextPr.RFonts.SetAll("Symbol");
			oNumberingLvl.SetByType(c_oAscNumberingLevel.Bullet, 0, String.fromCharCode(0x00B7), oTextPr);
			return oNumberingLvl;
		}
		else if ('-' === sText)
		{
			var oTextPr = new CTextPr();
			oTextPr.RFonts.SetAll("Arial");
			oNumberingLvl.SetByType(c_oAscNumberingLevel.Bullet, 0, String.fromCharCode(0x2013), oTextPr);
			return oNumberingLvl;
		}
		else if ('>' === sText)
		{
			var oTextPr = new CTextPr();
			oTextPr.RFonts.SetAll("Wingdings");
			oNumberingLvl.SetByType(c_oAscNumberingLevel.Bullet, 0, String.fromCharCode(0x00D8), oTextPr);
			return oNumberingLvl;
		}

		return null;
	};
	/**
	 * Подбираем подходящий нумерованный список
	 * @param sText {string}
	 * @returns {null | {Lvl : CNumberingLvl, Value : number}}
	 */
	CRunAutoCorrect.prototype.private_GetSuitableNumberedLvlForAutoCorrect = function(sText)
	{
		if (sText.length < 2)
			return null;

		var sLastChar = sText.charAt(sText.length - 1);
		if ('.' !== sLastChar && ')' !== sLastChar)
			return null;

		var nFirstCharCode = sText.charCodeAt(0);

		var sValue = sText.slice(0, sText.length - 1);

		function private_ParseNextInt(sText, nPos)
		{
			if (nPos >= sText.length)
				return null;

			var nNextParaPos = sText.indexOf(")", nPos);
			var nNextDotPos  = sText.indexOf(".", nPos);

			var nEndPos;
			if (-1 === nNextDotPos && -1 === nNextParaPos)
				return null;
			else if (-1 === nNextDotPos)
				nEndPos = nNextParaPos;
			else if (-1 === nNextParaPos)
				nEndPos = nNextDotPos;
			else
				nEndPos = Math.min(nNextDotPos, nNextParaPos);

			var sValue = sText.slice(nPos, nEndPos);
			var nValue = parseInt(sValue);

			if (isNaN(nValue))
				return null;

			return {Value : nValue, Char : sText.charAt(nEndPos), Pos : nEndPos + 1};
		}

		// Проверяем, либо у нас все числовое, либо у нас все буквенное (все заглавные, либо все не заглавные)
		if (48 <= nFirstCharCode && nFirstCharCode <= 57)
		{

			var arrResult = [], nPos = 0;

			var oNum = private_ParseNextInt(sText, nPos);
			var oPrevLvl = null;
			var nCurLvl  = 0;
			while (oNum)
			{
				nPos = oNum.Pos;

				var oNumberingLvl = new CNumberingLvl();
				oNumberingLvl.InitDefault(0, c_oAscMultiLevelNumbering.Numbered);
				if ('.' === oNum.Char)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.DecimalDot_Left, nCurLvl);
				else if (')' === oNum.Char)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.DecimalBracket_Left, nCurLvl);

				if (oPrevLvl)
				{
					var arrPrevLvlText = oPrevLvl.GetLvlText();
					var arrLvlText     = [];
					for (var nIndex = 0, nCount = arrPrevLvlText.length; nIndex < nCount; ++nIndex)
					{
						arrLvlText.push(arrPrevLvlText[nIndex].Copy());
					}
					oNumberingLvl.SetLvlText(arrLvlText.concat(oNumberingLvl.GetLvlText()));
				}

				arrResult.push({Lvl : oNumberingLvl, Value : oNum.Value});

				oNum = private_ParseNextInt(sText, nPos);
				oPrevLvl = oNumberingLvl;
				nCurLvl++;
			}

			if (arrResult.length > 9)
				return null;

			return arrResult;
		}
		else if (65 <= nFirstCharCode && nFirstCharCode <= 90)
		{
			var nRoman  = AscCommon.RomanToInt(sValue);
			var nLetter = AscCommon.LatinNumberingToInt(sValue);

			var arrResult = [];
			if (!isNaN(nRoman))
			{
				var oNumberingLvl = new CNumberingLvl();
				oNumberingLvl.InitDefault(0, c_oAscMultiLevelNumbering.Numbered);

				if ('.' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.UpperRomanDot_Right, 0);
				else if (')' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.UpperRomanBracket_Left, 0);

				arrResult.push({Lvl : oNumberingLvl, Value : nRoman});
			}

			if (!isNaN(nLetter))
			{
				var oNumberingLvl = new CNumberingLvl();
				oNumberingLvl.InitDefault(0, c_oAscMultiLevelNumbering.Numbered);

				if ('.' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.UpperLetterDot_Left, 0);
				else if (')' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.UpperLetterBracket_Left, 0);

				arrResult.push({Lvl : oNumberingLvl, Value : nLetter});
			}

			if (arrResult.length > 0)
				return arrResult;

			return null;
		}
		else if (97 <= nFirstCharCode && nFirstCharCode <= 122)
		{
			var nRoman  = AscCommon.RomanToInt(sValue);
			var nLetter = AscCommon.LatinNumberingToInt(sValue);

			var arrResult = [];

			if (!isNaN(nRoman))
			{
				var oNumberingLvl = new CNumberingLvl();
				oNumberingLvl.InitDefault(0, c_oAscMultiLevelNumbering.Numbered);

				if ('.' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.LowerRomanDot_Right, 0);
				else if (')' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.LowerRomanBracket_Left, 0);

				arrResult.push({Lvl : oNumberingLvl, Value : nRoman});
			}
			if (!isNaN(nLetter))
			{
				var oNumberingLvl = new CNumberingLvl();
				oNumberingLvl.InitDefault(0, c_oAscMultiLevelNumbering.Numbered);

				if ('.' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.LowerLetterDot_Left, 0);
				else if (')' === sLastChar)
					oNumberingLvl.SetByType(c_oAscNumberingLevel.LowerLetterBracket_Left, 0);

				arrResult.push({Lvl : oNumberingLvl, Value : nLetter});
			}

			if (arrResult.length > 0)
				return arrResult;

			return null;
		}

		return null;
	};
	CRunAutoCorrect.prototype.private_GetSuitablePresentationBulletForAutoCorrect = function(sText)
	{
		if ('*' === sText)
		{
			return AscFormat.fGetPresentationBulletByNumInfo({Type : 0, SubType : 0});
		}
		else if ('-' === sText)
		{
			return AscFormat.fGetPresentationBulletByNumInfo({Type : 0, SubType : 8});
		}
		return null;
	};
	CRunAutoCorrect.prototype.private_GetSuitablePresentationNumberingForAutoCorrect = function(sText)
	{
		var sLastChar = sText.charAt(sText.length - 1);
		if ('.' !== sLastChar && ')' !== sLastChar)
			return null;

		var nNumType = null;
		if (sText === "(a)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth;
		}
		else if (sText === "a)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_AlphaLcParenR;
		}
		else if (sText === "a.")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod;
		}
		else if (sText === "(A)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_AlphaUcParenBoth;
		}
		else if (sText === "A)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_AlphaUcParenR;
		}
		else if (sText === "A.")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod;
		}
		else if (sText === "(1)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_ArabicParenBoth;
		}
		else if (sText === "1)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_ArabicParenR;
		}
		else if (sText === "1.")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_ArabicPeriod;
		}
		else if (sText === "(i)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth;
		}
		else if (sText === "i)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_RomanLcParenR;
		}
		else if (sText === "i.")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_RomanLcPeriod;
		}
		else if (sText === "(I)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_RomanUcParenBoth;
		}
		else if (sText === "I)")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_RomanUcParenR;
		}
		else if (sText === "I.")
		{
			nNumType = AscFormat.numbering_presentationnumfrmt_RomanUcPeriod;
		}
		if (nNumType !== null)
		{
			var oBullet                    = new AscFormat.CBullet();
			oBullet.bulletType             = new AscFormat.CBulletType();
			oBullet.bulletType.type        = AscFormat.BULLET_TYPE_BULLET_AUTONUM;
			oBullet.bulletType.AutoNumType = nNumType;
			return oBullet;
		}
		return null;
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunAutoCorrect = CRunAutoCorrect;

	window['AscWord'].AUTOCORRECT_FLAGS_NONE                     = AUTOCORRECT_FLAGS_NONE;
	window['AscWord'].AUTOCORRECT_FLAGS_ALL                      = AUTOCORRECT_FLAGS_ALL;
	window['AscWord'].AUTOCORRECT_FLAGS_FRENCH_PUNCTUATION       = AUTOCORRECT_FLAGS_FRENCH_PUNCTUATION;
	window['AscWord'].AUTOCORRECT_FLAGS_SMART_QUOTES             = AUTOCORRECT_FLAGS_SMART_QUOTES;
	window['AscWord'].AUTOCORRECT_FLAGS_HYPHEN_WITH_DASH         = AUTOCORRECT_FLAGS_HYPHEN_WITH_DASH;
	window['AscWord'].AUTOCORRECT_FLAGS_HYPERLINK                = AUTOCORRECT_FLAGS_HYPERLINK;
	window['AscWord'].AUTOCORRECT_FLAGS_FIRST_LETTER_SENTENCE    = AUTOCORRECT_FLAGS_FIRST_LETTER_SENTENCE;
	window['AscWord'].AUTOCORRECT_FLAGS_NUMBERING                = AUTOCORRECT_FLAGS_NUMBERING;
	window['AscWord'].AUTOCORRECT_FLAGS_DOUBLE_SPACE_WITH_PERIOD = AUTOCORRECT_FLAGS_DOUBLE_SPACE_WITH_PERIOD;

})(window);

