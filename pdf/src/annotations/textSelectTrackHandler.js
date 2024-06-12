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

(function(window) {
	/**
	 * Класс контролирует события работы трека формулы. Вызывать у этого класса события обновления можно
	 * сколько угодно раз, а этот класс уже отрисовщику и в интерфейс посылает события, только когда реально что-то
	 * изменилось
	 *
	 * @constructor
	 */
	function CTextSelectTrackHandler(drawingDocument, eventHandler) {
		this.DrawingDocument = drawingDocument;
		this.EventHandler    = eventHandler;
	}

	CTextSelectTrackHandler.prototype.Update = function() {
		this.OnChangePosition();
	};
	CTextSelectTrackHandler.prototype.OnChangePosition = function() {
		let bounds = this.GetBounds();
		if (!bounds) {
			this.OnHide();
			return;
		}

		this.OnShow(bounds);
	};

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CTextSelectTrackHandler.prototype.GetBounds = function() {
		let oViewer		= Asc.editor.getDocumentRenderer();
		let oFile		= oViewer.file;
		let aSelQuads	= oFile.getSelectionQuads();

		if (aSelQuads.length == 0) {
			return null;
		}

		let nPage		= aSelQuads[0].page;
		let aFirstQuads	= aSelQuads[0].quads[0];

        let oDoc    = oViewer.getPDFDoc();
        let oTr     = oDoc.pagesTransform[nPage].invert;

		let oPoint1 = oTr.TransformPoint(aFirstQuads[0], aFirstQuads[1]);
		let oPoint2 = oTr.TransformPoint(aFirstQuads[6], aFirstQuads[7]);

		return [Math.min(oPoint1.x, oPoint2.x), Math.min(oPoint1.y, oPoint2.y), Math.max(oPoint1.x, oPoint2.x), Math.max(oPoint1.y, oPoint2.y)];
	};
	CTextSelectTrackHandler.prototype.OnHide = function() {
		this.EventHandler.sendEvent("asc_onHideTextSelectTrack");
	};
	CTextSelectTrackHandler.prototype.OnShow = function(bounds) {
		this.EventHandler.sendEvent("asc_onShowTextSelectTrack", bounds);
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};
	window['AscPDF'].CTextSelectTrackHandler = CTextSelectTrackHandler;
})(window);
