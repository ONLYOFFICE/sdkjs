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
	 * Search mechanism. Stores paragraphs (or PageMatch object) with the found string
	 * @constructor
	 */
	function CPdfSearch(oPdfDoc) {
		AscCommonWord.CDocumentSearch.call(this, oPdfDoc);

		this.Show			= false;	// flag indicating to highlight PdfPageMatch results on the page
		this.PagesMatches	= [];		// search results by pages
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

		// Clear previous search elements
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

            // first search on the page itself, if there was no conversion to shapes
            let nStartIdx = this.Id;
            let oPdfPageResult;
            if (true != oFile.pages[i].isRecognized) {
                oPdfPageResult = oFile.searchPage(i);

                for (let nMatch = 0; nMatch < oPdfPageResult.matches.length; nMatch++) {
                    this.Add(oPdfPageResult.matches[nMatch]);
                }
            }
            
            // search in drawings in edit mode
			if (Asc.editor.canEdit()) {
				for (let j = 0; j < oPageInfo.drawings.length; ++j) {
					oPageInfo.drawings[j].Search && oPageInfo.drawings[j].Search(this, search_Common);
				}
			}

			// to do (is search needed in forms, annotations?)

            this.PagesMatches[i] = [];
            // Elements contains all search results, PagesMatches contains results by pages
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
		let oFile		= oViewer.file;
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

			oViewer.navigateToPage(oElm.GetAbsolutePage());
		}
		else if (oElm.pdfPageMatch == true) {
			oFile.removeSelection();

			oFile.Selection.Page1 = oElm[0].PageNum;
			oFile.Selection.Line1 = oElm[0].LineNum;
			oFile.Selection.Glyph1 = oElm[0].Char1;
			
			oFile.Selection.Page2 = oElm[oElm.length - 1].PageNum;
			oFile.Selection.Line2 = oElm[oElm.length - 1].LineNum;
			oFile.Selection.Glyph2 = oElm[oElm.length - 1].Char2;

			oFile.Selection.IsSelection = true;

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

