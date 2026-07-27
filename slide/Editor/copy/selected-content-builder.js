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
	function SelectedContentBuilder() {

	}
	SelectedContentBuilder.prototype.getSelectedContent = function () {

	};
	SelectedContentBuilder.prototype.getTheme = function () {

	};
	SelectedContentBuilder.prototype.getColorMap = function () {

	};
	function DocSelectedContentBuilder(docContent) {
		SelectedContentBuilder.call(this);
		this.docContent = docContent;
	}
	AscFormat.InitClassWithoutType(DocSelectedContentBuilder, SelectedContentBuilder);
	DocSelectedContentBuilder.prototype.getSelectedContent = function () {
		return this.docContent.GetSelectedContent();
	};
	DocSelectedContentBuilder.prototype.getTheme = function () {
		return this.docContent.Get_Theme();
	};
	DocSelectedContentBuilder.prototype.getColorMap = function () {
		return this.docContent.Get_ColorMap();
	};

	function OutlineSelectedContentBuilder(outlineView, selectedSlides) {
		SelectedContentBuilder.call(this);
		this.outlineView = outlineView;
		this.selectedSlides = selectedSlides;
	}
	AscFormat.InitClassWithoutType(OutlineSelectedContentBuilder, SelectedContentBuilder);
	OutlineSelectedContentBuilder.prototype.getOutlineView = function () {
		return this.outlineView;
	};
	OutlineSelectedContentBuilder.prototype.getSelectedContent = function () {
		const outlineView = this.getOutlineView();
		const selectedContent = new AscCommonWord.CSelectedContent();
		for (let i = 0; i < this.selectedSlides.length; i += 1) {
			const slide = this.selectedSlides[i];
			const paragraphs = outlineView.getOutlineSlideParagraphs(slide);
			for (let j = 0; j < paragraphs.length; j += 1) {
				const paragraph = outlineView.getSourceParagraph(paragraphs[j]);
				if (!paragraph) {
					continue;
				}

				const selectedElement = new AscCommonWord.CSelectedElement();
				selectedElement.SelectedAll = true;
				selectedElement.Element = paragraph.Copy(paragraph.Parent, paragraph.DrawingDocument, {});
				selectedContent.Elements.push(selectedElement);
			}
		}
		return selectedContent;
	};
	OutlineSelectedContentBuilder.prototype.getTheme = function () {
		if (this.selectedSlides.length) {
			const selectedSlide = this.selectedSlides[0];
			return selectedSlide.getTheme();
		}
		return null;
	};
	OutlineSelectedContentBuilder.prototype.getColorMap = function () {
		if (this.selectedSlides.length) {
			const selectedSlide = this.selectedSlides[0];
			return selectedSlide.getColorMap();
		}
		return null;
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].DocSelectedContentBuilder = DocSelectedContentBuilder;
	window["AscCommonSlide"].OutlineSelectedContentBuilder = OutlineSelectedContentBuilder;
})();
