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
		this.slide = null;
	}
	OutlineSlide.prototype.setSlide = function (pr) {
		this.slide = pr;
	};
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
		this.outlineShapes = null;
	}
	OutlineView.prototype.getPresentation = function () {
		return this.presentation;
	};
	OutlineView.prototype.setOutlineShape = function (shape) {
	this.outlineShape = shape;
	};
	OutlineView.prototype.updateAll = function (width, height) {
		const presentation = this.getPresentation();
		const outlineSlides = [];
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			const outlineSlide = slide.getOutlineSlide();
			outlineSlides.push(outlineSlide);
		}
		const outlineShape = this.getOutlineShape(outlineSlides, width, height);
		this.setOutlineShape(outlineShape);
	};
	OutlineView.prototype.addContentToOutlineShape = function (outlineShape, slideShape, pr) {
		if (!slideShape) {
			return;
		}
		const outlineContent = outlineShape.txBody.content;
		const slideContent = slideShape.txBody.content;
		const paragraphs = slideContent.Content;
		for (let i = 0 ; i < paragraphs.length; i += 1) {
			const paragraph = paragraphs[i];
			const copyParagraph = this.getCopyParagraph(outlineContent, paragraph);
			outlineContent.AddToContent(outlineContent.Content.length, copyParagraph);
		}
	};
	OutlineView.prototype.fillOutlineShape = function (outlineShape, outlineSlides, width) {
		for (let i = 0; i < outlineSlides.length; i += 1) {
			const slide = outlineSlides[i];
			this.addContentToOutlineShape(outlineShape, slide.title, width);
			this.addContentToOutlineShape(outlineShape, slide.subTitle, width);
			for (let j = 0; j < slide.content.length; j += 1) {
				this.addContentToOutlineShape(outlineShape, slide.content[j], width);
			}
		}
	}
	OutlineView.prototype.applyParagraphProps = function (outlineParagraph, slideParagraph) {
			const compiledPr = slideParagraph.getCompiledPr();
			const copyParaPr = new CParaPr();
			if (compiledPr.ParaPr.Bullet) {
				copyParaPr.Bullet = compiledPr.ParaPr.Bullet.createDuplicate();
				copyParaPr.Lvl = compiledPr.ParaPr.Lvl;
				copyParaPr.Ind = compiledPr.ParaPr.Ind.Copy();
				copyParaPr.Ind.FirstLine *= 0.5;
				copyParaPr.Ind.Left *= 0.5;
				copyParaPr.Ind.Right *= 0.5;
			}
			outlineParagraph.SetPr(copyParaPr);
			outlineParagraph.CheckRunContent(function (run) {
				const textPr = new CTextPr();
				textPr.SetFontSize(15);
				run.SetPr(textPr);
			});

	};
	OutlineView.prototype.getOutlineShape = function (outlineSlides, width, height) {
		return AscFormat.ExecuteNoHistory(function () {
			const outlineShape = new AscFormat.CShape();
			outlineShape.setBDeleted(false);
			outlineShape.createTextBody();
			const outlineContent = outlineShape.txBody.content;
			outlineContent.ClearContent(false);
			this.fillOutlineShape(outlineShape, outlineSlides, width);
			outlineShape.extX = width;
			outlineShape.extY = 2000;
			outlineShape.recalculateContent();
			outlineShape.contentWidth = width;
			outlineShape.outlineView = this;
			return outlineShape;
		}, this, []);
	};
	OutlineView.prototype.getCopyParagraph = function (parent, paragraph, props) {
		return AscFormat.ExecuteNoHistory(function () {
			const copyParagraph = paragraph.Copy(parent, null, props);
			this.applyParagraphProps(copyParagraph, paragraph);
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
