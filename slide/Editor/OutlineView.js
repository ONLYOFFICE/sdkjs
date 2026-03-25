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

(function(undefined) {
	
	function OutlineSlide() {
		this.title = null;
		this.subTitle = null;
		this.content = [];
	}

	OutlineSlide.prototype.setTitle = function (pr) {
		this.title = pr;
	};
	OutlineSlide.prototype.setSubTitle = function (pr) {
		this.subTitle = pr;
	};
	OutlineSlide.prototype.addContent = function (pr) {
		this.content.push(pr);
	};
	
	function OutlineView(presentation) {
		this.presentation = presentation;
		this.outlineShape = null;
	}
	OutlineView.prototype.getPresentation = function () {
		return this.presentation;
	};
	OutlineView.prototype.setOutlineShape = function (shape) {
	this.outlineShape = shape;
	};
	OutlineView.prototype.updateAll = function () {
		const presentation = this.getPresentation();
		const outlineSlides = [];
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			const outlineSlide = slide.getOutlineSlide();
			outlineSlides.push(outlineSlide);
		}
		const outlineShape = this.getOutlineShape(outlineSlides);
		this.setOutlineShape(outlineShape);
	};
	OutlineView.prototype.getOutlineShape = function (outlineSlides) {
		return AscFormat.ExecuteNoHistory(function () {
			const oSp = new AscFormat.CShape();
			oSp.setBDeleted(false);
			oSp.createTextBody();
			const oContent = oSp.txBody.content;
			for (let i = 0; i < outlineSlides.length; i += 1) {
				const slide = outlineSlides[i];
				if (slide.title) {
					const shapeContent = slide.title.txBody.content.Content;
					for (let i = 0 ; i < shapeContent.length; i += 1) {
						const paragraph = shapeContent[i];
						const copyParagraph = this.getCopyParagraph(oContent, paragraph);
						oContent.AddToContent(oContent.Content.length, copyParagraph);
					}
				}
			}
			oSp.recalculateContent();
			return oSp;
		}, this, []);
	};
	OutlineView.prototype.getCopyParagraph = function (parent, paragraph, props) {
		return AscFormat.ExecuteNoHistory(function () {
			const copyParagraph = paragraph.Copy2(parent);
			return copyParagraph;
		}, this, []);
	};
	OutlineView.prototype.draw = function (graphics) {
		if (this.outlineShape) {
			this.outlineShape.draw(graphics);
		}
	}


	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].OutlineSlide = OutlineSlide;
	window["AscCommonSlide"].OutlineView = OutlineView;
})();
