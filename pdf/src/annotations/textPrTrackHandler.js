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
	function CAnnotTextPrTrackHandler(drawingDocument, eventHandler) {
		this.DrawingDocument = drawingDocument;
		this.EventHandler    = eventHandler;

		this.Annot			= null;
		this.PageNum		= -1;
		this.ForceUpdate	= true;
	}

	CAnnotTextPrTrackHandler.prototype.SetTrackObject = function(oAnnot, pageNum, isActive) {
		if (oAnnot !== this.Annot || (oAnnot && (this.PageNum !== pageNum || this.ForceUpdate))) {
			this.Annot			= oAnnot;
			this.ForceUpdate	= false;
			this.PageNum		= pageNum;

			let bounds = null;
			if (this.Annot)
				bounds = this.GetBounds();

			if (bounds) {
				this.OnShow(bounds);
			}
			else {
				this.Annot    = null;
				this.PageNum = -1;

				this.OnHide();
			}
		}
	};
	CAnnotTextPrTrackHandler.prototype.Update = function() {
		this.ForceUpdate = true;
	};
	CAnnotTextPrTrackHandler.prototype.OnChangePosition = function() {
		let bounds = this.GetBounds();
		if (!bounds)
			return;

		this.OnShow(bounds);
	};

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CAnnotTextPrTrackHandler.prototype.GetBounds = function() {
		let oAnnot = this.Annot;

		if (!oAnnot)
			return null;
		
		let oDoc = oAnnot.GetDocument();
		if (!oDoc) {
			return null;
		}

        let aOrigRect   = oAnnot.GetOrigRect();
        let aRD         = oAnnot.GetRectangleDiff() || [0, 0, 0, 0]; // отступ координат фигуры с текстом от ректа аннотации
        let nPage       = oAnnot.GetPage();
		let oTr			= oDoc.pagesTransform[nPage].invert;
		
        let aFreeTextRect = []; // прямоугольник

        // левый верхний
        aFreeTextRect.push({
            x: (aOrigRect[0] + aRD[0]),
            y: (aOrigRect[1] + aRD[1])
        });
        // правый верхний
        aFreeTextRect.push({
            x: (aOrigRect[2] - aRD[2]),
            y: (aOrigRect[1] + aRD[1])
        });
        // правый нижний
        aFreeTextRect.push({
            x: (aOrigRect[2] - aRD[2]),
            y: (aOrigRect[3] - aRD[3])
        });
        // левый нижний
        aFreeTextRect.push({
            x: (aOrigRect[0] + aRD[0]),
            y: (aOrigRect[3] - aRD[3])
        });

		
		let oPoint1 = oTr.TransformPoint(aFreeTextRect[0].x, aFreeTextRect[0].y); // левый верхний
		let oPoint2 = oTr.TransformPoint(aFreeTextRect[1].x, aFreeTextRect[1].y); // правый верхний
		let oPoint3 = oTr.TransformPoint(aFreeTextRect[2].x, aFreeTextRect[2].y); // правый нижний
		let oPoint4 = oTr.TransformPoint(aFreeTextRect[3].x, aFreeTextRect[3].y); // левый нижний

		let xMin = Math.min(oPoint1.x, oPoint2.x, oPoint3.x, oPoint4.x);
		let yMin = Math.min(oPoint1.y, oPoint2.y, oPoint3.y, oPoint4.y);
		let xMax = Math.max(oPoint1.x, oPoint2.x, oPoint3.x, oPoint4.x);
		let yMax = Math.max(oPoint1.y, oPoint2.y, oPoint3.y, oPoint4.y);

		return [xMin, yMin, xMax, yMax];
	};
	CAnnotTextPrTrackHandler.prototype.OnHide = function() {
		this.EventHandler.sendEvent("asc_onHideAnnotTextPrTrack");
	};
	CAnnotTextPrTrackHandler.prototype.OnShow = function(bounds) {
		this.EventHandler.sendEvent("asc_onShowAnnotTextPrTrack", bounds);
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscPDF'] = window['AscPDF'] || {};
	window['AscPDF'].CAnnotTextPrTrackHandler = CAnnotTextPrTrackHandler;
})(window);
