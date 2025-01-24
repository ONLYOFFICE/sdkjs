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

	CTextSelectTrackHandler.prototype.Update = function(bCheckMouseUpPos) {
		this.OnChangePosition(bCheckMouseUpPos);
	};
	CTextSelectTrackHandler.prototype.OnChangePosition = function(bCheckMouseUpPos) {
		let oFile = Asc.editor.getDocumentRenderer().file;
		if (oFile.Selection.IsSelection || false === Asc.editor.NeedShowTextSelectPanel()) {
			this.OnHide();
			return;
		}
		
		let bounds = this.GetBounds();
		if (!bounds) {
			this.OnHide();
			return;
		}

		this.OnShow(bounds, bCheckMouseUpPos);
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

		let nFirstPage		= aSelQuads[0].page;
		let aFirstQuads		= aSelQuads[0].quads[0];
		let aMaxFirstRect	= [aFirstQuads[0], aFirstQuads[1], aFirstQuads[6], aFirstQuads[7]]; // x1, y1, x2, y2

		aSelQuads[0].quads.forEach(function(quads) {
			if (aMaxFirstRect[0] > quads[0]) {
				aMaxFirstRect[0] = quads[0];
			}
			if (aMaxFirstRect[1] > quads[1]) {
				aMaxFirstRect[1] = quads[1];
			}
			if (aMaxFirstRect[2] < quads[6]) {
				aMaxFirstRect[2] = quads[6];
			}
			if (aMaxFirstRect[3] < quads[7]) {
				aMaxFirstRect[3] = quads[7];
			}
		});

		let nLastPage		= aSelQuads[aSelQuads.length - 1].page;
		let aLastQuads		= aSelQuads[aSelQuads.length - 1].quads[0];
		let aMaxLastRect	= [aLastQuads[0], aLastQuads[1], aLastQuads[6], aLastQuads[7]]; // x1, y1, x2, y2

		aSelQuads[aSelQuads.length - 1].quads.forEach(function(quads) {
			if (aMaxLastRect[0] > quads[0]) {
				aMaxLastRect[0] = quads[0];
			}
			if (aMaxLastRect[1] > quads[1]) {
				aMaxLastRect[1] = quads[1];
			}
			if (aMaxLastRect[2] < quads[6]) {
				aMaxLastRect[2] = quads[6];
			}
			if (aMaxLastRect[3] < quads[7]) {
				aMaxLastRect[3] = quads[7];
			}
		});

		let oDoc    	= oViewer.getPDFDoc();
		let oFirtsTr	= oDoc.pagesTransform[nFirstPage].invert;
		let oLastTr     = oDoc.pagesTransform[nLastPage].invert;

		let oFirstPoint1 = oFirtsTr.TransformPoint(aMaxFirstRect[0], aMaxFirstRect[1]);
		let oFirstPoint2 = oFirtsTr.TransformPoint(aMaxFirstRect[2], aMaxFirstRect[3]);
		
		let oLastPoint1 = oLastTr.TransformPoint(aMaxLastRect[0], aMaxLastRect[1]);
		let oLastPoint2 = oLastTr.TransformPoint(aMaxLastRect[2], aMaxLastRect[3]);

		let x1 = Math.min(oFirstPoint1.x, oFirstPoint2.x, oLastPoint1.x, oLastPoint2.x);
		let x2 = Math.max(oFirstPoint1.x, oFirstPoint2.x, oLastPoint1.x, oLastPoint2.x);
		let y1 = Math.min(oFirstPoint1.y, oFirstPoint2.y, oLastPoint1.y, oLastPoint2.y);
		let y2 = Math.max(oFirstPoint1.y, oFirstPoint2.y, oLastPoint1.y, oLastPoint2.y);

		return [x1, y1, x2, y2];
	};
	CTextSelectTrackHandler.prototype.OnHide = function() {
		this.EventHandler.sendEvent("asc_onHideTextSelectTrack");
	};
	CTextSelectTrackHandler.prototype.OnShow = function(bounds, bCheckMouseUpPos) {
		let isMouseUpOnTop;

		if (bCheckMouseUpPos) {
			let oViewer = Asc.editor.getDocumentRenderer();
			let nRectH = bounds[3] - bounds[1];
			
			if (AscCommon.global_mouseEvent.Y - oViewer.y < bounds[1] + nRectH / 2) {
				isMouseUpOnTop = true;
			}
			else {
				isMouseUpOnTop = false;
			}
		}

		this.EventHandler.sendEvent("asc_onShowTextSelectTrack", bounds, isMouseUpOnTop);
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};
	window['AscPDF'].CTextSelectTrackHandler = CTextSelectTrackHandler;
})(window);
