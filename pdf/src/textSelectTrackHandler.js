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

(function(window) {
	/**
	 * This class controls text select track events. You can call update events on this class
	 * as many times as needed, and this class will only send events to the renderer and interface
	 * when something actually changes
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
		let oDoc = Asc.editor.getPDFDoc();
		let isCanEditShape = Asc.editor.canEdit();

		let oFile = Asc.editor.getDocumentRenderer().file;
		if (false == oFile.isSelectionUse() || false === Asc.editor.NeedShowTextSelectPanel() || (oDoc.activeDrawing && isCanEditShape)) {
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
