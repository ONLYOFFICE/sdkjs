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
	/**
	 * Механизм поиска. Хранит параграфы (или PageMatch объект) с найденной строкой
	 * @constructor
	 */
	function CPdfSearch(oPdfDoc) {
		AscCommonWord.CDocumentSearch.call(this, oPdfDoc);

		this.Show			= false;	// флаг что надо хайлайтить PdfPageMatch результаты на странице
		this.PagesMatches	= [];		// результаты поиска по страницам
		this.PagesLines		= [];
	}

	CPdfSearch.prototype.constructor = CPdfSearch;
    CPdfSearch.prototype = Object.create(AscCommonWord.CDocumentSearch.prototype);

	CPdfSearch.prototype.GetPdfPageMatches = function(nPage) {
		if (!this.PagesMatches[nPage]) {
			return [];
		}

		return this.PagesMatches[nPage].filter(function(match) {
			return match.pdfPageMatch;
		});
	};
	CPdfSearch.prototype.Clear = function() {
		this.Reset();

		// Очищаем предыдущие элементы поиска
		for (var Id in this.Elements) {
			if (true != this.Elements[Id].pdfPageMatch) {
				this.Elements[Id].ClearSearchResults();
			}
		}

		this.Id         	= 0;
		this.Count      	= 0;
		this.Elements   	= {};
		this.ReplacedId 	= [];
		this.PagesMatches	= [];
		this.CurId      	= -1;
		this.Direction  	= true;

		this.TextAroundUpdate = true;
		this.StopTextAround();
		this.SendClearAllTextAround();
	};

	CPdfSearch.prototype.Search = function() {
		let oViewer		= Asc.editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let oPagesInfo  = oViewer.pagesInfo;
        
        for (let i = 0; i < oPagesInfo.pages.length; i++) {
            let oPageInfo = oPagesInfo.pages[i];

            // сначала ищем на самой странице, если не было конвертации в шейпы
            let nStartIdx = this.Id;
            let oPdfPageResult;
            if (true != oFile.pages[i].isRecognized) {
                oPdfPageResult = oFile.searchPage(i);

                for (let nMatch = 0; nMatch < oPdfPageResult.matches.length; nMatch++) {
                    this.Add(oPdfPageResult.matches[nMatch]);
                }
            }
            
            // затем ищем в drawings
            for (let j = 0; j < oPageInfo.drawings.length; ++j) {
                oPageInfo.drawings[j].Search && oPageInfo.drawings[j].Search(this, search_Common);
            }

			// to do (нужен поиск в формах, аннотациях?)

            this.PagesMatches[i] = [];
            // в Elements все результаты поиска, в PagesMatches - результаты по страницам
            for (let j = nStartIdx; j < this.Id; j++) {
                this.PagesMatches[i].push(this.Elements[j]);
            }

			if (oPdfPageResult) {
				this.PagesLines[i] = oPdfPageResult.pageLines;
			}
        }
	};

	CPdfSearch.prototype.Select = function(nId) {
		let oViewer		= Asc.editor.getDocumentRenderer();
		let oElm		= this.Elements[nId];

		if (!oElm) {
			return;
		}

		if (oElm instanceof AscWord.Paragraph) {
			oViewer.CurrentSearchNavi = null;

			let SearchElement = oElm.SearchResults[nId];
			if (SearchElement) {
				oElm.Selection.Use   = true;
				oElm.Selection.Start = false;

				oElm.Set_SelectionContentPos(SearchElement.StartPos, SearchElement.EndPos);
				oElm.Set_ParaContentPos(SearchElement.StartPos, false, -1, -1);

				oElm.Document_SetThisElementCurrent();
			}

			oViewer.navigateToPage(oElm.Get_AbsolutePage());			
		}
		else if (oElm.pdfPageMatch == true) {
			oViewer.CurrentSearchNavi = oElm;
			oViewer.ToSearchResult();
		}

		this.SetCurrent(nId);
	};

	CPdfSearch.prototype.StartTextAround = function() {
		if (!this.TextAroundUpdate)
			return this.SendAllTextAround();

		this.TextAroundUpdate = false;
		this.StopTextAround();

		this.TextAroundId = 0;

		this.LogicDocument.GetApi().sync_startTextAroundSearch();

		let oThis = this;
		this.TextAroundTimer = setTimeout(function() {
			oThis.ContinueGetTextAround()
		}, 20);

		this.TextArround = [];
	};
	CPdfSearch.prototype.ContinueGetTextAround = function(){
		let arrResult = [];

		let nStartTime = performance.now();
		while (performance.now() - nStartTime < 20) {
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
		if (this.TextAroundId >= 0 && this.TextAroundId < this.Id) {
			this.TextAroundTimer = setTimeout(function() {
				oThis.ContinueGetTextAround();
			}, 20);
		}
		else {
			this.TextAroundId    = -1;
			this.TextAroundTimer = null;
			
			window['AscPDF'].PdfPageMatch.lastPartInfo = null;
			this.LogicDocument.GetApi().sync_endTextAroundSearch();
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscPDF'].CPdfSearch = CPdfSearch;
})(window);

