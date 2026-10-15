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

(function () {
	function SelectedTextImageBuilder(presentation, sourceContent) {
		this.presentation = presentation;
		this.sourceContent = sourceContent;
		this.contentIndents = 30;
	}

	SelectedTextImageBuilder.prototype.trimContent = function (docContent) {
		new AscCommonSlide.ContentTrimmer(docContent).trim();
	};
	SelectedTextImageBuilder.prototype.recalculateTextArt = function (docContent, contentWidth, contentHeight) {
		if (docContent.Parent && docContent.Parent.parent && docContent.Parent.parent instanceof AscFormat.CShape) {
			return docContent.Parent.parent.checkTextWarp(docContent, docContent.Parent.parent.getBodyPr(), contentWidth, contentHeight, true, false);
		}
		return null;
	};
	SelectedTextImageBuilder.prototype.getCanvas = function (imageWidth, imageHeight) {
		const canvas = document.createElement('canvas');
		canvas.width = ((imageWidth * AscCommon.g_dKoef_mm_to_pix) + 2 * this.contentIndents + 0.5) >> 0;
		canvas.height = ((imageHeight * AscCommon.g_dKoef_mm_to_pix) + 2 * this.contentIndents + 0.5) >> 0;
		return canvas;
	};
	SelectedTextImageBuilder.prototype.getGraphics = function (canvas, imageWidth, imageHeight) {
		const context = canvas.getContext('2d');
		const graphics = new AscCommon.CGraphics();
		graphics.init(context, canvas.width, canvas.height, imageWidth + 2.0 * this.contentIndents / AscCommon.g_dKoef_mm_to_pix, imageHeight + 2.0 * this.contentIndents / AscCommon.g_dKoef_mm_to_pix);
		graphics.m_oFontManager = AscCommon.g_fontManager;
		graphics.m_oCoordTransform.tx = this.contentIndents;
		graphics.m_oCoordTransform.ty = this.contentIndents;
		graphics.transform(1, 0, 0, 1, 0, 0);
		return graphics;
	};
	SelectedTextImageBuilder.prototype.draw = function (graphics, docContent, textWarpObject) {
		if (textWarpObject && textWarpObject.oTxWarpStructNoTransform) {
			textWarpObject.oTxWarpStructNoTransform.draw(graphics, docContent.Parent.parent.Get_Theme(), docContent.Parent.parent.Get_ColorMap());
		} else {
			const oldShowParaMarks = this.presentation.Api.ShowParaMarks;
			this.presentation.Api.ShowParaMarks = false;
			docContent.Draw(0, graphics);
			this.presentation.Api.ShowParaMarks = oldShowParaMarks;
		}
	};
	SelectedTextImageBuilder.prototype.getContentWidth = function (docContent) {
		let maxWidth = 0;
		const content = docContent.Content;
		for (let i = 0; i < content.length; ++i) {
			var oParagraph = content[i];
			for (let j = 0; j < oParagraph.Lines.length; ++j) {
				if (oParagraph.Lines[j].Ranges[0].W > maxWidth) {
					maxWidth = oParagraph.Lines[j].Ranges[0].W;
				}
			}
		}
		return maxWidth;
	};
	SelectedTextImageBuilder.prototype.build = function (selectedContent) {
		if (selectedContent.Elements.length === 0) {
			return null;
		}
		const docContent = this.sourceContent;
		const docContentForDraw = new AscFormat.CDrawingDocContent(docContent.Parent, docContent.DrawingDocument, 0, 0, 20000, 20000);
		selectedContent.ReplaceElements(docContentForDraw);
		this.trimContent(docContentForDraw);
		if (docContentForDraw.Content.length === 0) {
			return null;
		}

		docContentForDraw.Reset(0, 0, 20000, 20000);
		docContentForDraw.Recalculate_Page(0, true);
		const contentWidth = this.getContentWidth(docContentForDraw) + 1;
		docContentForDraw.Reset(0, 0, contentWidth, 20000);
		docContentForDraw.Recalculate_Page(0, true);
		const contentHeight = docContentForDraw.GetSummaryHeight();
		const textWartObject = this.recalculateTextArt(docContentForDraw, contentWidth, contentHeight);
		const canvas = this.getCanvas(contentWidth, contentHeight);
		let imageUrl;
		if (!window["NATIVE_EDITOR_ENJINE"]) {
			var oGraphics = this.getGraphics(canvas, contentWidth, contentHeight);
			this.draw(oGraphics, docContentForDraw, textWartObject);
			imageUrl = canvas.toDataURL("image/png");
		} else {
			imageUrl = "";
		}
		const image = AscFormat.DrawingObjectsController.prototype.createImage(imageUrl, 0, 0, canvas.width * AscCommon.g_dKoef_pix_to_mm, canvas.height * AscCommon.g_dKoef_pix_to_mm);
		return new DrawingCopyObject(image, 0, 0, contentWidth, contentHeight, imageUrl);
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].SelectedTextImageBuilder = SelectedTextImageBuilder;
})();
