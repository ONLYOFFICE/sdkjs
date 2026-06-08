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
	 * Класс для поиска внутри параграфа
	 * @constructor
	 */
	function CParagraphSearch(Paragraph, SearchEngine, Type)
	{
		this.Paragraph    = Paragraph;
		this.SearchEngine = SearchEngine;
		this.Type         = Type;

		this.ContentPos   = new AscWord.CParagraphContentPos();

		this.StartPos     = null; // Запоминаем здесь стартовую позицию поиска
		this.SearchIndex  = 0;    // Номер символа, с которым мы проверяем совпадение
		this.StartPosBuf  = [];
	}

	CParagraphSearch.prototype.Reset = function()
	{
		this.StartPos    = null;
		this.SearchIndex = 0;
		this.StartPosBuf = [];
	};
	CParagraphSearch.prototype.Check = function(nIndex, oItem)
	{
		return this.SearchEngine.Pattern.Check(nIndex, oItem, this.SearchEngine);
	};
	CParagraphSearch.prototype.GetPrefix = function(nIndex)
	{
		return this.SearchEngine.GetPrefix(nIndex);
	};
	CParagraphSearch.prototype.CheckSearchEnd = function()
	{
		return (this.SearchIndex === this.SearchEngine.Pattern.GetLength());
	};

	/**
	 * Найденные элементы в параграфе. Записаны в массиве Paragraph.SearchResults
	 * @constructor
	 */
	function CParagraphSearchElement(StartPos, EndPos, Type, Id)
	{
		this.StartPos  = StartPos;
		this.EndPos    = EndPos;
		this.Type      = Type;
		this.ResultStr = "";
		this.Id        = Id;

		this.ClassesS = [];
		this.ClassesE = [];
	}
	CParagraphSearchElement.prototype.RegisterClass = function(isStart, oClass)
	{
		if (isStart)
			this.ClassesS.push(oClass);
		else
			this.ClassesE.push(oClass);
	};


	/**
	 * Метки начала и конца найденных элементов внутри параграфа
	 * @constructor
	 */
	function CParagraphSearchMark(SearchResult, Start, Depth)
	{
		this.SearchResult = SearchResult;
		this.Start        = Start;
		this.Depth        = Depth;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CParagraphSearch        = CParagraphSearch;
	window['AscCommonWord'].CParagraphSearchElement = CParagraphSearchElement;
	window['AscCommonWord'].CParagraphSearchMark    = CParagraphSearchMark;

})(window);
