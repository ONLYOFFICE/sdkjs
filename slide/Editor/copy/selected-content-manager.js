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
	function SelectedContentManager(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent) {
		this.presentation = presentation;
		this.sourceFormattingContent = sourceFormattingContent;
		this.endFormattingContent = endFormattingContent;
		this.imagesSelectedContent = imagesSelectedContent;
	}
	SelectedContentManager.prototype.fillSelectedContent = function () {};

	function OutlineSelectedContentManager(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent) {
		SelectedContentManager.call(this, presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent);
	}
	AscFormat.InitClassWithoutType(OutlineSelectedContentManager, SelectedContentManager);
	OutlineSelectedContentManager.prototype.getOutlineView = function () {
		return this.presentation.GetOutlineView();
	};
	OutlineSelectedContentManager.prototype.getSelectedShapeCount = function (contents) {
		let shapeCount = 0;
		for (let i = 0; i < contents.length; i += 1) {
			if (contents[i].content) {
				shapeCount += 1;
			}
		}
		return shapeCount;
	};
	OutlineSelectedContentManager.prototype.extractTextFromContent = function (docContent) {
		const presentation = this.presentation;
		const selectedContentBuilder = new AscCommonSlide.DocSelectedContentBuilder(docContent);

		const endFormatter = new AscCommonSlide.EndContentFormatter(presentation, selectedContentBuilder);
		this.endFormattingContent.DocContent = endFormatter.getFormattedContent();

		const sourceFormatter = new AscCommonSlide.SourceContentFormatter(presentation, selectedContentBuilder);
		this.sourceFormattingContent.DocContent = sourceFormatter.getFormattedContent();

		const selectedContentToDraw = sourceFormatter.getFormattedContent();
		const imageBuilder = new SelectedTextImageBuilder(presentation, docContent);
		const drawingCopy = imageBuilder.build(selectedContentToDraw);
		if (drawingCopy) {
			this.imagesSelectedContent.Drawings.push(drawingCopy);
		}
	};
	OutlineSelectedContentManager.prototype.fillDocContent = function () {
		const oThis = this;
		this.getOutlineView().forEachSelectedContent(function (selectProps) {
			if (!selectProps.content) {
				return;
			}
			oThis.extractTextFromContent(selectProps.content);
			return true;
		});
	};
	OutlineSelectedContentManager.prototype.fillMultipleParagraphs = function () {
		const presentation = this.presentation;
		const selectedSlides = presentation.GetSelectedSlideObjects();
		if (selectedSlides.length === 0) {
			return;
		}
		const collector = new AscCommonSlide.NormalSlidesCollector(presentation, this.sourceFormattingContent, this.endFormattingContent, this.imagesSelectedContent);
		collector.collect(selectedSlides);

		const selectedContentBuilder = new AscCommonSlide.OutlineSelectedContentBuilder(this.getOutlineView(), selectedSlides);
		const sourceFormatter = new AscCommonSlide.SourceContentFormatter(presentation, selectedContentBuilder);
		const endFormatter = new AscCommonSlide.EndContentFormatter(presentation, selectedContentBuilder);
		this.endFormattingContent.DocContent = endFormatter.getFormattedContent();
		this.sourceFormattingContent.DocContent = sourceFormatter.getFormattedContent();
	};
	OutlineSelectedContentManager.prototype.fillSelectedContent = function () {
		const outlineView = this.getOutlineView();
		if (!outlineView) {
			return;
		}
		const docContent = outlineView.getDocContent();
		if (!docContent) {
			return;
		}
		const contents = outlineView.getSelectionContentsInfo();
		if (this.getSelectedShapeCount(contents) > 1) {
			this.fillMultipleParagraphs();
		} else {
			this.fillDocContent();
		}
	};

	function ThumbnailSelectedContentManager(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent) {
		SelectedContentManager.call(this, presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent);
	}
	AscFormat.InitClassWithoutType(ThumbnailSelectedContentManager, SelectedContentManager);
	ThumbnailSelectedContentManager.prototype.fillSelectedContent = function () {
		const presentation = this.presentation;
		const sourceFormattingContent = this.sourceFormattingContent;
		const endFormattingContent = this.endFormattingContent;
		const imagesSelectedContent = this.imagesSelectedContent;
		const isSlidePageMode = presentation.IsSlidePageMode();
		const isMasterSlideMode = presentation.IsMasterSlideMode();
		if (isSlidePageMode || isMasterSlideMode) {
			let collector;
			if (isMasterSlideMode) {
				collector = new AscCommonSlide.MasterSlideCollector(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent);
			} else {
				collector = new AscCommonSlide.NormalSlidesCollector(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent);
			}
			const selectedSlides = presentation.GetSelectedSlideObjects();
			collector.collect(selectedSlides);
		}
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].OutlineSelectedContentManager = OutlineSelectedContentManager;
	window["AscCommonSlide"].ThumbnailSelectedContentManager = ThumbnailSelectedContentManager;
})();
