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
	function ContentFormatter(presentation, selectedContentBuilder) {
		this.presentation = presentation;
		this.selectedContentBuilder = selectedContentBuilder;
	}
	ContentFormatter.prototype.getSelectedContent = function () {
		return this.selectedContentBuilder.getSelectedContent();
	};
	ContentFormatter.prototype.getFormattedContent = function () {

	};
	function EndContentFormatter(presentation, selectedContentBuilder) {
		ContentFormatter.call(this, presentation, selectedContentBuilder);
	}
	AscFormat.InitClassWithoutType(EndContentFormatter, ContentFormatter);
	EndContentFormatter.prototype.getFormattedContent = function () {
		const endFormattingContent = this.getSelectedContent();
		if (!endFormattingContent) {
			return;
		}

		for (let i = 0; i < endFormattingContent.Elements.length; ++i) {
			const element = endFormattingContent.Elements[i].Element;
			if (element.GetType() === AscCommonWord.type_Paragraph) {
				if (element.Pr && element.Pr.DefaultRunPr && AscFormat.isRealNumber(element.Pr.DefaultRunPr.FontSize)) {
					const textPr = element.Pr.Copy();
					textPr.DefaultRunPr.FontSize = undefined;
					element.Set_Pr(textPr);
				}
				this.presentation.internalResetElementsFontSize(element.Content);
			}
		}
		return endFormattingContent;
	};

	function SourceContentFormatter(presentation, selectedContentBuilder) {
		ContentFormatter.call(this, presentation, selectedContentBuilder);
		this.presentation = presentation;
	}
	AscFormat.InitClassWithoutType(SourceContentFormatter, ContentFormatter);
	SourceContentFormatter.prototype.getFormattedContent = function () {
		const sourceFormattingContent = this.getSelectedContent();
		if (!sourceFormattingContent) {
			return;
		}
		const sourceElements = [];
		for (let i = 0; i < sourceFormattingContent.Elements.length; ++i) {
			const paragraph = sourceFormattingContent.Elements[i].Element;
			paragraph.private_CompileParaPr();
			sourceElements.push(paragraph);
		}
		AscFormat.SaveContentSourceFormatting(sourceElements, sourceElements, this.getTheme(), this.getColorMap());
		return sourceFormattingContent;
	};
	SourceContentFormatter.prototype.getTheme = function () {
		return this.selectedContentBuilder.getTheme();
	};
	SourceContentFormatter.prototype.getColorMap = function () {
		return this.selectedContentBuilder.getColorMap();
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].SourceContentFormatter = SourceContentFormatter;
	window["AscCommonSlide"].EndContentFormatter = EndContentFormatter;
})();
