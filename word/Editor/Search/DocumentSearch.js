/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window)
{
	/**
	 * Механизм поиска. Хранит параграфы с найденной строкой
	 * @constructor
	 */
	function DocumentSearch(oLogicDocument)
	{
		this.LogicDocument = oLogicDocument;
		this.Text          = "";
		this.MatchCase     = false;
		this.Word          = false;
		this.Pattern       = new AscCommonWord.CSearchPatternEngine();
		this.RegExp        = null;

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

	DocumentSearch.prototype.Reset = function()
	{
		this.Text      = "";
		this.MatchCase = false;
		this.Word	   = false;
		this.RegExp    = null;
	};
	/**
	 * @param {AscCommon.CSearchSettings} oProps
	 */
	DocumentSearch.prototype.Compare = function(oProps)
	{
		return (oProps && this.Text === oProps.GetText()
			&& this.MatchCase === oProps.IsMatchCase()
			&& this.Word === oProps.IsWholeWords());
	};
	DocumentSearch.prototype.Clear = function()
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
	DocumentSearch.prototype.Add = function(Paragraph)
	{
		this.Count++;
		this.Elements[this.Id++] = Paragraph;
		return (this.Id - 1);
	};
	DocumentSearch.prototype.GetElementsMap = function()
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
	DocumentSearch.prototype.Select = function(nId, bUpdateStates)
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
	DocumentSearch.prototype.SetCurrent = function(nId)
	{
		this.CurId = undefined !== nId ? nId : -1;
		let nIndex = -1 !== this.CurId ? this.GetElementIndexById(this.CurId) : -1;

		let oApi = this.LogicDocument.GetApi()
		oApi.sync_setSearchCurrent(nIndex, this.Count);
	};
	DocumentSearch.prototype.ResetCurrent = function()
	{
		this.SetCurrent(-1);
	};
	DocumentSearch.prototype.GetCount = function()
	{
		return this.Count;
	};
	DocumentSearch.prototype.GetCurrent = function()
	{
		return this.CurId;
	};
	DocumentSearch.prototype.Replace = function(sReplaceString, Id, bRestorePos)
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
	DocumentSearch.prototype.private_AddReplacedId = function(nId)
	{
		for (let nPos = 0, nCount = this.ReplacedId.length; nPos < nCount; ++nPos)
		{
			if (this.ReplacedId[nPos] > nId)
				return this.ReplacedId.splice(nPos, 0, nId);
		}

		this.ReplacedId.push(nId);
	};
	DocumentSearch.prototype.private_AddReplacedStringToRun = function(oRun, nInRunPos)
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
	DocumentSearch.prototype.GetElementIndexById = function(nId)
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
	DocumentSearch.prototype.ReplaceAll = function(NewStr, bUpdateStates)
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
	DocumentSearch.prototype.Set = function(oProps)
	{
		if (!oProps)
			return;
		
		this.RegExp    = null;
		this.Text      = oProps.GetText();
		this.MatchCase = oProps.IsMatchCase();
		this.Word      = oProps.IsWholeWords();
		
		if (this.Text instanceof RegExp)
		{
			this.SetRegExp(this.Text);
		}
		else
		{
			var _sText = this.Text;
			if (!this.MatchCase)
				_sText = this.Text.toLowerCase();
			
			this.Pattern.Set(_sText);
			this.private_CalculatePrefix();
		}
	};
	DocumentSearch.prototype.SetRegExp = function(regExp)
	{
		this.RegExp = null;
		this.Pattern.Set("");
		this.private_CalculatePrefix();
	};
	DocumentSearch.prototype.IsEmpty = function()
	{
		if (this.RegExp)
			return false;

		return !this.Pattern.GetLength();
	};
	DocumentSearch.prototype.IsWholeWords = function()
	{
		return this.Word;
	};
	DocumentSearch.prototype.IsMatchCase = function()
	{
		return this.MatchCase;
	};
	DocumentSearch.prototype.SetFootnotes = function(arrFootnotes)
	{
		this.Footnotes = arrFootnotes;
	};
	DocumentSearch.prototype.SetEndnotes = function(arrEndnotes)
	{
		this.Endnotes = arrEndnotes;
	};
	DocumentSearch.prototype.GetFootnotes = function()
	{
		return this.Footnotes;
	};
	DocumentSearch.prototype.GetEndnotes = function()
	{
		return this.Endnotes;
	};
	DocumentSearch.prototype.GetDirection = function()
	{
		return this.Direction;
	};
	DocumentSearch.prototype.SetDirection = function(bDirection)
	{
		this.Direction = bDirection;
	};
	DocumentSearch.prototype.private_CalculatePrefix = function()
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
	DocumentSearch.prototype.GetPrefix = function(nIndex)
	{
		return this.Prefix[nIndex];
	};
	DocumentSearch.prototype.StartTextAround = function()
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
	DocumentSearch.prototype.ContinueGetTextAround = function()
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
	DocumentSearch.prototype.StopTextAround = function()
	{
		if (this.TextAroundTimer)
		{
			clearTimeout(this.TextAroundTimer);
			this.LogicDocument.GetApi().sync_endTextAroundSearch();
		}

		this.TextAroundTimer = null;
		this.TextAroundId    = -1;
	};
	DocumentSearch.prototype.SendAllTextAround = function()
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
	DocumentSearch.prototype.SendClearAllTextAround = function()
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
	window['AscWord'] = window['AscWord'] || {};
	AscWord.DocumentSearch  = DocumentSearch;

	// for backward compatibility
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CDocumentSearch = DocumentSearch;

})(window);

