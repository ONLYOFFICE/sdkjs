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

(function(window, undefined){

	var AscCommon = window['AscCommon'];

	function CPrintPreviewBase(api, parentElementId) {
		this.api = api;
		this.parentElementId = parentElementId;
		this.page = null;
		this.pageImage = null;
		this.canvas = document.createElement("canvas");
		this.init();
	}
	CPrintPreviewBase.prototype.init = function () {
		let parentElem = document.getElementById(this.parentElementId);
		parentElem.appendChild(this.canvas);
		this.resize();
	};
	CPrintPreviewBase.prototype.resize = function (parentElemSrc) {
		let parentElem = parentElemSrc ? parentElemSrc : document.getElementById(this.parentElementId);

		this.canvas.style.width  = parentElem.offsetWidth + "px";
		this.canvas.style.height = parentElem.offsetHeight + "px";

		AscCommon.calculateCanvasSize(this.canvas);
	};
	CPrintPreviewBase.prototype.checkGraphics = function(width, height, w_mm, h_mm) {
		let aspectMM = w_mm / h_mm;
		let aspect = width / height;

		let w, h;

		if (aspectMM > aspect)
		{
			w = width;
			h = (width * h_mm / w_mm) >> 0;
		}
		else
		{
			w = (height * w_mm / h_mm) >> 0;
			h = height;
		}

		if (this.pageImage === null)
			this.pageImage = document.createElement("canvas");

		this.pageImage.width = w;
		this.pageImage.height = h;

		let pageCtx = this.pageImage.getContext("2d");
		pageCtx.fillStyle = "#FFFFFF";
		pageCtx.fillRect(0, 0, w, h);

		let g = new AscCommon.CGraphics();
		g.init(this.pageImage.getContext("2d"), w, h, w_mm, h_mm);
		g.m_oFontManager = AscCommon.g_fontManager;
		g.transform(1, 0, 0, 1, 0, 0);

		g.IsNoDrawingEmptyPlaceholderText = true;
		g.IsNoDrawingEmptyPlaceholder = true;
		g.isPrintMode = true;
		g.isSupportEditFeatures = function() { return false; };

		return g;
	};
	CPrintPreviewBase.prototype.update = function(paperSize) {
		let width_canvas = this.canvas.width;
		let height_canvas = this.canvas.height;

		this.canvas.width = width_canvas;

		if (null === this.page)
			return;

		let offset = AscCommon.AscBrowser.convertToRetinaValue(25);
		let min_size = 3 * offset;
		if (width_canvas < min_size || height_canvas < min_size)
			return;

		let width = this.canvas.width - (offset << 1);
		let height = this.canvas.height - (offset << 1);

		let ctx = this.canvas.getContext("2d");

		let strokeRect = this.drawPrintPreview(width, height, paperSize);

		if (this.pageImage)
		{
			let x = (width_canvas - this.pageImage.width) >> 1;
			let y = (height_canvas - this.pageImage.height) >> 1;

			ctx.drawImage(this.pageImage, x, y);

			if (undefined === paperSize)
			{
				strokeRect = {
					x : x,
					y : y,
					w : this.pageImage.width,
					h : this.pageImage.height
				};
			}

			if (null != strokeRect)
			{
				ctx.strokeStyle = AscCommon.GlobalSkin.PageOutline;
				let lineW = AscCommon.AscBrowser.retinaPixelRatio >> 0;

				ctx.lineWidth = lineW;
				ctx.strokeRect(strokeRect.x + lineW / 2, strokeRect.y + lineW / 2, strokeRect.w - lineW, strokeRect.h - lineW);
				ctx.beginPath();
			}
		}
	};
	CPrintPreviewBase.prototype.close = function() {
		if (this.canvas)
		{
			let parentElem = document.getElementById(this.parentElementId);
			parentElem.removeChild(this.canvas);
			this.canvas = null;
		}
	};
	CPrintPreviewBase.prototype.drawPrintPreview = function(width, height, paperSize) {};

	function CDocumentPrintPreview(api, parentElementId) {
		CPrintPreviewBase.call(this, api, parentElementId);
	}
	AscFormat.InitClassWithoutType(CDocumentPrintPreview, CPrintPreviewBase);
	CDocumentPrintPreview.prototype.drawPrintPreview = function (width, height, paperSize) {
		if (this.api.WordControl.m_oDrawingDocument.IsFreezePage(this.page))
			return;

		let page = this.api.WordControl.m_oDrawingDocument.m_arrPages[this.page];
		let w_mm = page.width_mm;
		let h_mm = page.height_mm;

		let g = this.checkGraphics(width, height, w_mm, h_mm);

		let oldViewMode = this.api.isViewMode;
		let oldShowMarks = this.api.ShowParaMarks;

		this.api.isViewMode = true;
		this.api.ShowParaMarks = false;

		this.api.WordControl.m_oLogicDocument.SetupBeforeNativePrint({
			"drawPlaceHolders" : false,
			"drawFormHighlight" : false,
			"isPrint" : true
		}, g);
		this.api.WordControl.m_oLogicDocument.DrawPage(this.page, g);
		this.api.WordControl.m_oLogicDocument.RestoreAfterNativePrint();

		this.api.isViewMode = oldViewMode;
		this.api.ShowParaMarks = oldShowMarks;
	};

	function CPdfPrintPreview(api, parentElementId) {
		CPrintPreviewBase.call(this, api, parentElementId);
	}
	AscFormat.InitClassWithoutType(CPdfPrintPreview, CPrintPreviewBase);
	CPdfPrintPreview.prototype.drawPrintPreview = function (width, height, paperSize) {
		let viewer = this.api.WordControl.m_oDrawingDocument.m_oDocumentRenderer;
		let file = viewer.file;

		if (!file)
			return;

		let page = file.pages[this.page];

		let w_mm = page.W * 25.4 / page.Dpi;
		let h_mm = page.H * 25.4 / page.Dpi;

		let aspectMM = w_mm / h_mm;
		let aspect = width / height;
		let w, h;

		if (aspectMM > aspect)
		{
			w = width;
			h = (width * h_mm / w_mm) >> 0;
		}
		else
		{
			w = (height * w_mm / h_mm) >> 0;
			h = height;
		}

		this.pageImage = viewer.GetPrintPage(this.page, w, h, this.printContentType);
	};
	
	function CPresentationPrintPreview(api, parentElementId) {
		CPrintPreviewBase.call(this, api, parentElementId);
	}
	AscFormat.InitClassWithoutType(CPresentationPrintPreview, CPrintPreviewBase);
	CPresentationPrintPreview.prototype.initGraphicsFlags = function() {

	};
	CPresentationPrintPreview.prototype.restoreGraphicsFlags = function() {

	};
	CPresentationPrintPreview.prototype.getPrintPreviewGraphics = function() {
		let g = new AscCommon.CGraphics();
		g.init(this.canvas.getContext("2d"), this.canvas.width, this.canvas.height, this.canvas.width * AscCommon.g_dKoef_pix_to_mm, this.canvas.height * AscCommon.g_dKoef_pix_to_mm);
		g.m_oFontManager = AscCommon.g_fontManager;
		g.transform(1, 0, 0, 1, 0, 0);
		return g;
	};
	CPresentationPrintPreview.prototype.getGraphics = function() {
		let g = this.getPrintPreviewGraphics();
		g.IsNoDrawingEmptyPlaceholderText = true;
		g.IsNoDrawingEmptyPlaceholder = true;
		g.isPrintMode = true;
		g.isSupportEditFeatures = function() { return false; };
		return g;
	};
	CPresentationPrintPreview.prototype.update = function(paperSize) {
		if (null === this.page)
			return;

		let width_canvas = parseInt(this.canvas.style.width, 10);
		let height_canvas = parseInt(this.canvas.style.height, 10);

		let offset = 20;
		if (width_canvas < offset || height_canvas < offset)
			return;

		const graphics = this.getGraphics();


		let strokeRect = this.drawPrintPreview(width_canvas, height_canvas, offset, paperSize, graphics);
		if (this.pageImage)
		{
			let x = (width_canvas - this.pageImage.width) >> 1;
			let y = (height_canvas - this.pageImage.height) >> 1;


			if (undefined === paperSize)
			{
				strokeRect = {
					x : x,
					y : y,
					w : this.pageImage.width,
					h : this.pageImage.height
				};
			}

			if (null != strokeRect)
			{
				const color = AscCommon.GlobalSkin.PageOutline;
				graphics.p_color((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF, 0xFF);
				graphics.p_width(0);
				graphics._s();
				graphics._m(strokeRect.x + 1 / 2, strokeRect.y + 1 / 2);
				graphics._l(strokeRect.x + strokeRect.w - 1, strokeRect.y + 1 / 2);
				graphics._l(strokeRect.x + strokeRect.w - 1, strokeRect.y + strokeRect.h - 1);
				graphics._l(strokeRect.x + 1 / 2, strokeRect.y + strokeRect.h - 1);
				graphics._z();
				graphics.ds();
			}
		}
	};
	CPresentationPrintPreview.prototype.drawPrintPreview = function (widthCanvas, heightCanvas, offset, paperSize, graphics) {
		let w_mm = this.api.WordControl.m_oLogicDocument.GetWidthMM();
		let h_mm = this.api.WordControl.m_oLogicDocument.GetHeightMM();


		let strokeRect;
		let coef = 1;
		let width = widthCanvas - offset / 2;
		let height = heightCanvas - offset / 2;
		let x = 0;
		let y = 0;
		let w = width;
		let h = height;

		if (undefined !== paperSize)
		{
			let paperW = paperSize[0];
			let paperH = paperSize[1];
			width *= AscCommon.g_dKoef_pix_to_mm;
			height *= AscCommon.g_dKoef_pix_to_mm;
			coef = Math.min((width / w_mm), (height / h_mm));
			w = w_mm * coef;
			h = h_mm * coef;
			x = (widthCanvas - w) / 2;
			y = (heightCanvas - h) / 2;



		}
		strokeRect = {
			x : x,
			y : y,
			w : w,
			h : h
		};

		const m = new AscCommon.CMatrix();
		m.Scale(coef, coef);
		m.Translate(0, 0);
		graphics.SetBaseTransform(m);
		graphics.b_color1(255, 255, 255, 255);
		graphics.rect(0, 0, w_mm, h_mm);
		this.drawPage(this.page, graphics);
		graphics.ResetBaseTransform();
		return strokeRect;
	};
	CPresentationPrintPreview.prototype.drawPage = function (index, graphics, options) {
		// const m = new AscCommon.CMatrix();
		// m.Scale(0.2,0.2);
		// graphics.SetBaseTransform(m);
		this.api.WordControl.m_oLogicDocument.DrawPage(index, graphics);
	};

	AscCommon.CDocumentPrintPreview = CDocumentPrintPreview;
	AscCommon.CPdfPrintPreview = CPdfPrintPreview;
	AscCommon.CPresentationPrintPreview = CPresentationPrintPreview;
})(window);
