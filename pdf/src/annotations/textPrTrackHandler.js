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

		let nPage		= oAnnot.GetPage();
		let aBoundsMM	= [];

		if (oAnnot.IsFreeText()) {
			let oTxBoxBounds = oAnnot.GetTextBoxShape().getRectBounds();
			aBoundsMM = [oTxBoxBounds.l, oTxBoxBounds.t, oTxBoxBounds.r, oTxBoxBounds.b];
		}

		let pos0 = this.DrawingDocument.ConvertCoordsToCursorWR(aBoundsMM[0], aBoundsMM[1], nPage);
		let pos1 = this.DrawingDocument.ConvertCoordsToCursorWR(aBoundsMM[2], aBoundsMM[3], nPage);

		return [pos0.X, pos0.Y, pos1.X, pos1.Y];
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
