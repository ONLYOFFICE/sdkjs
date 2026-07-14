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

	const PRESENTATION_PRINT_PREVIEW_OFFSET = 10;
	function CPresentationPrintPreview(api, parentElementId) {
		CPrintPreviewBase.call(this, api, parentElementId);
		this.initCalculatedValues();
		this.printer = null;
	}
	AscFormat.InitClassWithoutType(CPresentationPrintPreview, CPrintPreviewBase);
	CPresentationPrintPreview.prototype.getPresentation = function () {
		return this.api.WordControl.m_oLogicDocument;
	};
	CPresentationPrintPreview.prototype.initCalculatedValues = function() {
		this.calculatedValues = {
			canvasWidth: null,
			canvasHeight: null
		};
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
		g.IsPrintPreview = true;
		g.isSupportEditFeatures = function() { return false; };
		return g;
	};
	CPresentationPrintPreview.prototype.draw = function (graphics, adaptSizes) {
		graphics.b_color1(255, 255, 255, 255);
		graphics.rect(adaptSizes.x, adaptSizes.y, adaptSizes.width, adaptSizes.height);
		graphics.df();

		const m = new AscCommon.CMatrix();
		m.Scale(adaptSizes.scale, adaptSizes.scale);
		m.Translate(adaptSizes.x, adaptSizes.y);

		graphics.SaveGrState();
		graphics.SetBaseTransform(m);
		graphics.reset();
		this.printer.drawPage(graphics, this.page);
		graphics.SetBaseTransform(null);
		graphics.reset();
		graphics.RestoreGrState();
	}
	CPresentationPrintPreview.prototype.initPrinter = function(advancedOptions) {
		this.printer = new AscCommonSlide.PrintManager(this.getPresentation(), advancedOptions);
	}
	CPresentationPrintPreview.prototype.update = function() {
		this.initCalculatedValues();
		if (null === this.page)
			return;

		const width_canvas = this.getCanvasWidthMM();
		const height_canvas = this.getCanvasHeightMM();


		let offset = 10;
		if (width_canvas < offset || height_canvas < offset)
			return;

		const graphics = this.getGraphics();
		graphics.b_color1(255, 255, 255, 255);
		graphics.rect(0, 0, width_canvas, height_canvas);
		graphics.df();
		const adaptPaperSizes = this.getAdaptPaperSizes();
		if (adaptPaperSizes)
		{
			this.draw(graphics, adaptPaperSizes);
				const color = parseInt(AscCommon.GlobalSkin.PageOutline.slice(1), 16);
				graphics.p_color((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF, 0xFF);
				graphics.p_width(0);
				graphics._s();
				graphics._m(adaptPaperSizes.x, adaptPaperSizes.y);
				graphics._l(adaptPaperSizes.x + adaptPaperSizes.width, adaptPaperSizes.y);
				graphics._l(adaptPaperSizes.x + adaptPaperSizes.width, adaptPaperSizes.y + adaptPaperSizes.height);
				graphics._l(adaptPaperSizes.x, adaptPaperSizes.y + adaptPaperSizes.height);
				graphics._z();
				graphics.ds();
		}
	};
	CPresentationPrintPreview.prototype.getPageSizes = function () {
		return this.printer.getPageSizes();
	};
	CPresentationPrintPreview.prototype.getAdaptPaperSizes = function() {
		const widthCanvas = this.getCanvasWidthMM();
		const heightCanvas = this.getCanvasHeightMM();
		const paperSize = this.getPageSizes();
		let width = widthCanvas - PRESENTATION_PRINT_PREVIEW_OFFSET;
		let height = heightCanvas - PRESENTATION_PRINT_PREVIEW_OFFSET;
		let x = 0;
		let y = 0;
		let paperScale = 1;
		if (paperSize)
		{
			let paperW = paperSize.width;
			let paperH = paperSize.height;
			paperScale = Math.min(width / paperW, height / paperH);
			width = paperW * paperScale;
			height = paperH * paperScale;
			x = (widthCanvas - width) / 2;
			y = (heightCanvas - height) / 2;
		}
		return {
			x : x,
			y : y,
			width : width,
			height : height,
			scale: paperScale
		};
	}
	CPresentationPrintPreview.prototype.getCanvasHeightMM = function() {
		if (this.calculatedValues.canvasHeight === null) {
			this.calculatedValues.canvasHeight = this.canvas.height * AscCommon.g_dKoef_pix_to_mm;
		}
		return this.calculatedValues.canvasHeight;
	};
	CPresentationPrintPreview.prototype.getCanvasWidthMM = function() {
		if (this.calculatedValues.canvasWidth === null) {
			this.calculatedValues.canvasWidth = this.canvas.width * AscCommon.g_dKoef_pix_to_mm;
		}
		return this.calculatedValues.canvasWidth;
	};
	CPresentationPrintPreview.prototype.updateInterfaceState = function () {
		const presentation = this.getPresentation();
		if (presentation) {
			presentation.UpdateNotePageOrientationState();
		}
	};

	AscCommon.CDocumentPrintPreview = CDocumentPrintPreview;
	AscCommon.CPdfPrintPreview = CPdfPrintPreview;
	AscCommon.CPresentationPrintPreview = CPresentationPrintPreview;
})(window);
