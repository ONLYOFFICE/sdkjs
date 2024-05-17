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
	/**
	 * Механизм поиска. Хранит параграфы с найденной строкой
	 * @constructor
	 */
	function CDocumentSearch(oLogicDocument)
	{
		this.LogicDocument = oLogicDocument;
		this.Text          = "";
		this.MatchCase     = false;
		this.Word          = false;
		this.Pattern       = new AscCommonWord.CSearchPatternEngine();

		this.Prefix        = [];

		this.Id            = 0;
		this.Count         = 0;
		this.Elements      = {};
		this.ReplacedId    = [];
		this.CurId         = -1;
		this.Direction     = true; // направление true - вперед, false - назад
		this.ClearOnRecalc = true; // Флаг, говорящий о том, запустился ли пересчет из-за Replace
		this.Selection     = false;
		this.Footnotes     = [];
		this.Endnotes      = [];
		this.InsertPattern = new AscCommonWord.CSearchPatternEngine();

		this.TextAroundId     = -1;
		this.TextAroundTimer  = null;
		this.TextAroundUpdate = true;
		this.ReplaceEvent     = true;
		this.TextAroundEmpty  = true; // Флаг, что все очищено, чтобы не очищать повторно
	}

	CDocumentSearch.prototype.Reset = function()
	{
		this.Text      = "";
		this.MatchCase = false;
		this.Word	   = false;
	};
	/**
	 * @param {AscCommon.CSearchSettings} oProps
	 */
	CDocumentSearch.prototype.Compare = function(oProps)
	{
		return (oProps && this.Text === oProps.GetText()
			&& this.MatchCase === oProps.IsMatchCase()
			&& this.Word === oProps.IsWholeWords());
	};
	CDocumentSearch.prototype.Clear = function()
	{
		this.Reset();

		// Очищаем предыдущие элементы поиска
		for (var Id in this.Elements)
		{
			this.Elements[Id].ClearSearchResults();
		}

		this.Id         = 0;
		this.Count      = 0;
		this.Elements   = {};
		this.ReplacedId = [];
		this.CurId      = -1;
		this.Direction  = true;

		this.TextAroundUpdate = true;
		this.StopTextAround();
		this.SendClearAllTextAround();
	};
	CDocumentSearch.prototype.Add = function(Paragraph)
	{
		this.Count++;
		this.Elements[this.Id++] = Paragraph;
		return (this.Id - 1);
	};
	CDocumentSearch.prototype.GetElementsMap = function()
	{
		let map = {};
		for (let searchId in this.Elements)
		{
			let paraId = this.Elements[searchId].GetId();
			if (!map[paraId])
				map[paraId] = this.Elements[searchId];
		}
		return map;
	};
	CDocumentSearch.prototype.Select = function(nId, bUpdateStates)
	{
		var Paragraph = this.Elements[nId];
		if (Paragraph)
		{
			var SearchElement = Paragraph.SearchResults[nId];
			if (SearchElement)
			{
				Paragraph.Selection.Use   = true;
				Paragraph.Selection.Start = false;

				Paragraph.Set_SelectionContentPos(SearchElement.StartPos, SearchElement.EndPos);
				Paragraph.Set_ParaContentPos(SearchElement.StartPos, false, -1, -1);

				Paragraph.Document_SetThisElementCurrent(false !== bUpdateStates);
			}

			this.SetCurrent(nId);
		}
	};
	CDocumentSearch.prototype.SetCurrent = function(nId)
	{
		this.CurId = undefined !== nId ? nId : -1;
		let nIndex = -1 !== this.CurId ? this.GetElementIndexById(this.CurId) : -1;

		let oApi = this.LogicDocument.GetApi()
		oApi.sync_setSearchCurrent(nIndex, this.Count);
	};
	CDocumentSearch.prototype.ResetCurrent = function()
	{
		this.SetCurrent(-1);
	};
	CDocumentSearch.prototype.GetCount = function()
	{
		return this.Count;
	};
	CDocumentSearch.prototype.GetCurrent = function()
	{
		return this.CurId;
	};
	CDocumentSearch.prototype.Replace = function(sReplaceString, Id, bRestorePos)
	{
		this.InsertPattern.Set(sReplaceString);

		var oPara = this.Elements[Id];
		if (oPara)
		{
			if (oPara.IsInPlaceholder())
				return false;
			
			var oLogicDocument   = oPara.LogicDocument;
			var isTrackRevisions = oLogicDocument ? oLogicDocument.IsTrackRevisions() : false;

			var SearchElement = oPara.SearchResults[Id];
			if (SearchElement)
			{
				if (oPara.GetPlaceHolderObject(SearchElement.StartPos)
					|| oPara.GetPlaceHolderObject(SearchElement.EndPos))
					return false;
				
				var ContentPos, StartPos, EndPos, bSelection;
				if (true === bRestorePos)
				{
					// Сохраняем позицию состояние параграфа, чтобы курсор остался в том же месте и после замены.
					bSelection = oPara.IsSelectionUse();
					ContentPos = oPara.Get_ParaContentPos(false, false);
					StartPos   = oPara.Get_ParaContentPos(true, true);
					EndPos     = oPara.Get_ParaContentPos(true, false);

					oPara.Check_NearestPos({ContentPos : ContentPos});
					oPara.Check_NearestPos({ContentPos : StartPos});
					oPara.Check_NearestPos({ContentPos : EndPos});
				}

				if (isTrackRevisions)
				{
					// Встанем в конечную позицию поиска и добавим новый текст
					var oEndContentPos = SearchElement.EndPos;
					var oEndRun        = SearchElement.ClassesE[SearchElement.ClassesE.length - 1];

					var nRunPos = oEndContentPos.Get(SearchElement.ClassesE.length - 1);

					if (reviewtype_Add === oEndRun.GetReviewType() && oEndRun.GetReviewInfo().IsCurrentUser())
					{
						this.private_AddReplacedStringToRun(oEndRun, nRunPos);
					}
					else
					{
						var oRunParent      = oEndRun.GetParent();
						var nRunPosInParent = oEndRun.GetPosInParent(oRunParent);
						var oReplaceRun     = oEndRun.Split2(nRunPos, oRunParent, nRunPosInParent);

						if (!oReplaceRun.IsEmpty())
							oReplaceRun.Split2(0, oRunParent, nRunPosInParent + 1);

						this.private_AddReplacedStringToRun(oReplaceRun, 0);
						oReplaceRun.SetReviewType(reviewtype_Add);
					}
				}
				else
				{
					// Сначала в начальную позицию поиска добавляем новый текст
					var StartContentPos = SearchElement.StartPos;
					var StartRun        = SearchElement.ClassesS[SearchElement.ClassesS.length - 1];

					var RunPos = StartContentPos.Get(SearchElement.ClassesS.length - 1);
					this.private_AddReplacedStringToRun(StartRun, RunPos);
				}

				// Выделяем старый объект поиска и удаляем его
				oPara.Selection.Use = true;
				oPara.Set_SelectionContentPos(SearchElement.StartPos, SearchElement.EndPos);
				oPara.Remove();

				// Перемещаем курсор в конец поиска
				oPara.RemoveSelection();
				oPara.Set_ParaContentPos(SearchElement.StartPos, true, -1, -1);

				// Удаляем запись о данном элементе
				this.Count--;

				oPara.RemoveSearchResult(Id);
				delete this.Elements[Id];
				this.private_AddReplacedId(Id);

				if (true === bRestorePos)
				{
					oPara.Set_SelectionContentPos(StartPos, EndPos);
					oPara.Set_ParaContentPos(ContentPos, true, -1, -1);
					oPara.Selection.Use = bSelection;
					oPara.Clear_NearestPosArray();
				}

				if (this.ReplaceEvent && !this.TextAroundUpdate)
					this.LogicDocument.GetApi().sync_removeTextAroundSearch(Id);

				return true;
			}
		}

		return false;
	};
	CDocumentSearch.prototype.private_AddReplacedId = function(nId)
	{
		for (let nPos = 0, nCount = this.ReplacedId.length; nPos < nCount; ++nPos)
		{
			if (this.ReplacedId[nPos] > nId)
				return this.ReplacedId.splice(nPos, 0, nId);
		}

		this.ReplacedId.push(nId);
	};
	CDocumentSearch.prototype.private_AddReplacedStringToRun = function(oRun, nInRunPos)
	{
		var isMathRun = oRun.IsMathRun();
		for (var nIndex = 0, nAdd = 0, nCount = this.InsertPattern.GetLength(); nIndex < nCount; ++nIndex)
		{
			var oItem = this.InsertPattern.Get(nIndex).ToRunElement(isMathRun);
			if (oItem)
			{
				oRun.AddToContent(nInRunPos + nAdd, oItem, false);
				nAdd++;
			}
		}
	};
	CDocumentSearch.prototype.GetElementIndexById = function(nId)
	{
		for (let nPos = 0, nCount = this.ReplacedId.length; nPos < nCount; ++nPos)
		{
			if (this.ReplacedId[nPos] > nId)
				return (nId - nPos);
			else if (this.ReplacedId[nPos] === nId)
				return -1;
		}

		return (nId - this.ReplacedId.length);
	};
	CDocumentSearch.prototype.ReplaceAll = function(NewStr, bUpdateStates)
	{
		this.RemoveEvent = false;

		for (var Id = this.Id; Id >= 0; --Id)
		{
			if (this.Elements[Id])
				this.Replace(NewStr, Id, true);
		}

		this.RemoveEvent = true;

		this.Clear();
	};
	/**
	 * @param {AscCommon.CSearchSettings} oProps
	 */
	CDocumentSearch.prototype.Set = function(oProps)
	{
		if (!oProps)
			return;

		this.Text      = oProps.GetText();
		this.MatchCase = oProps.IsMatchCase();
		this.Word      = oProps.IsWholeWords();

		var _sText = this.Text;
		if (!this.MatchCase)
			_sText = this.Text.toLowerCase();

		this.Pattern.Set(_sText);
		this.private_CalculatePrefix();
	};
	CDocumentSearch.prototype.IsWholeWords = function()
	{
		return this.Word;
	};
	CDocumentSearch.prototype.IsMatchCase = function()
	{
		return this.MatchCase;
	};
	CDocumentSearch.prototype.SetFootnotes = function(arrFootnotes)
	{
		this.Footnotes = arrFootnotes;
	};
	CDocumentSearch.prototype.SetEndnotes = function(arrEndnotes)
	{
		this.Endnotes = arrEndnotes;
	};
	CDocumentSearch.prototype.GetFootnotes = function()
	{
		return this.Footnotes;
	};
	CDocumentSearch.prototype.GetEndnotes = function()
	{
		return this.Endnotes;
	};
	CDocumentSearch.prototype.GetDirection = function()
	{
		return this.Direction;
	};
	CDocumentSearch.prototype.SetDirection = function(bDirection)
	{
		this.Direction = bDirection;
	};
	CDocumentSearch.prototype.private_CalculatePrefix = function()
	{
		var nLen = this.Pattern.GetLength();

		this.Prefix    = new Int32Array(nLen);
		this.Prefix[0] = 0;

		for (var nPos = 1, nK = 0; nPos < nLen; ++nPos)
		{
			nK = this.Prefix[nPos - 1]
			while (nK > 0 && !(this.Pattern.Get(nPos).IsMatch(this.Pattern.Get(nK))))
				nK = this.Prefix[nK - 1];

			if (this.Pattern.Get(nPos).IsMatch(this.Pattern.Get(nK)))
				nK++;

			this.Prefix[nPos] = nK;
		}
	};
	CDocumentSearch.prototype.GetPrefix = function(nIndex)
	{
		return this.Prefix[nIndex];
	};
	CDocumentSearch.prototype.StartTextAround = function()
	{
		if (!this.TextAroundUpdate)
			return this.SendAllTextAround();

		this.TextAroundUpdate = false;
		this.StopTextAround();

		this.TextAroundId = 0;

		this.LogicDocument.GetApi().sync_startTextAroundSearch();

		let oThis = this;
		this.TextAroundTimer = setTimeout(function()
		{
			oThis.ContinueGetTextAround()
		}, 20);

		this.TextArround = [];
	};
	CDocumentSearch.prototype.ContinueGetTextAround = function()
	{
		let arrResult = [];

		let nStartTime = performance.now();
		while (performance.now() - nStartTime < 20)
		{
			if (this.TextAroundId >= this.Id)
				break;

			let sId = this.TextAroundId++;

			if (!this.Elements[sId])
				continue;

			let textAround = this.Elements[sId].GetTextAroundSearchResult(sId);
			this.TextArround[sId] = textAround;
			arrResult.push([sId, textAround]);
		}

		if (arrResult.length)
			this.TextAroundEmpty = false;

		this.LogicDocument.GetApi().sync_getTextAroundSearchPack(arrResult);

		let oThis = this;
		if (this.TextAroundId >= 0 && this.TextAroundId < this.Id)
		{
			this.TextAroundTimer = setTimeout(function()
			{
				oThis.ContinueGetTextAround();
			}, 20);
		}
		else
		{
			this.TextAroundId    = -1;
			this.TextAroundTimer = null;
			this.LogicDocument.GetApi().sync_endTextAroundSearch();
		}
	};
	CDocumentSearch.prototype.StopTextAround = function()
	{
		if (this.TextAroundTimer)
		{
			clearTimeout(this.TextAroundTimer);
			this.LogicDocument.GetApi().sync_endTextAroundSearch();
		}

		this.TextAroundTimer = null;
		this.TextAroundId    = -1;
	};
	CDocumentSearch.prototype.SendAllTextAround = function()
	{
		if (this.TextAroundTimer)
			return;

		let arrResult = [];
		for (let nId = 0; nId < this.Id; ++nId)
		{
			if (!this.Elements[nId] || undefined === this.TextArround[nId])
				continue;

			arrResult.push([nId, this.TextArround[nId]]);
		}

		let oApi = this.LogicDocument.GetApi();
		oApi.sync_startTextAroundSearch();
		oApi.sync_getTextAroundSearchPack(arrResult);
		oApi.sync_endTextAroundSearch();
	};
	CDocumentSearch.prototype.SendClearAllTextAround = function()
	{
		if (this.TextAroundEmpty)
			return;

		let oApi = this.LogicDocument.GetApi();
		if (!oApi)
			return;

		oApi.sync_startTextAroundSearch();
		oApi.sync_endTextAroundSearch();

		this.TextAroundEmpty = true;
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CDocumentSearch = CDocumentSearch;

})(window);

