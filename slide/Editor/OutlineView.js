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

(function (undefined) {

	function getRGBFromHex(color) {
		if (!color) {
			return {R: 0, G: 0, B: 0};
		}
		const prepareColor = parseInt(color.slice(1), 16);
		return {R: (prepareColor >> 16) & 0xff, G: (prepareColor >> 8) & 0xff, B: prepareColor & 0xff};
	}
	function rebuildPos(pos, sourceContent, sourceParagraph) {
		const newPos = pos.slice();
		newPos[0] = {Class: sourceContent, Position: sourceParagraph.GetIndex()};
		newPos[1] = {Class: sourceParagraph, Position: pos[1].Position};
		for (let i = 2; i < newPos.length; i += 1) {
			const previousPos = newPos[i - 1];
			const newClass = previousPos.Class.GetElement(previousPos.Position);
			if (newClass.Get_Type() === para_Math) {
				newPos[i].Class = newClass.Root;
			} else {
				newPos[i].Class = newClass;
			}
		}
		return newPos;
	}

	function getOutlineTextPr(isTitle, isMathRun) {
		const textPr = new CTextPr();
		textPr.SetFontSize(10);
		if (isTitle) {
			textPr.SetBold(true);
		}
		if (isMathRun) {
			textPr.SetFontFamily("Cambria Math");
			textPr.RFonts.SetAll("Cambria Math");
		} else {
			textPr.SetFontFamily("Arial");
			textPr.RFonts.SetAll("Arial");
		}
		return textPr;
	}


	const DECORATORDRAWER_OFFSET_LEFT = 2;
	const DECORATORDRAWER_SLIDE_WIDTH = 6;
	function DecoratorDrawer(outlineView, transform, selectedSlideIndex, focusSlideIndex) {
		this.cache = {
			slideNumbersWidth: null
		};
		this.outlineView = outlineView;

		this.selectedSlideIndex = selectedSlideIndex;
		this.focusSlideIndex = focusSlideIndex;

		this.backgroundColor = AscCommon.GlobalSkin.BackgroundColorThumbnails;
		this.normalSlideColor = AscCommon.GlobalSkin.ThumbnailsPageOutline || "#C0C0C0";
		this.activeSlideColor = AscCommon.GlobalSkin.ThumbnailsPageOutlineActive;
		this.hoverSlideColor = AscCommon.GlobalSkin.ThumbnailsPageOutlineHover;

		this.transform = transform;

		this.init();
	}
	DecoratorDrawer.prototype.init = function () {
		return this.getDecorCacheManager().resetColors();
	};
	DecoratorDrawer.prototype.getParagraphY = function (paragraph) {
		return this.outlineView.getParagraphY(paragraph);
	};
	DecoratorDrawer.prototype.getDecorCacheManager = function () {
		return this.outlineView.decorCacheManager;
	}
	DecoratorDrawer.prototype.getPresentation = function () {
		return this.outlineView.getPresentation();
	};
	DecoratorDrawer.prototype.getSlideNumbersWidth = function () {
		if (this.cache.slideNumbersWidth === null) {
			this.cache.slideNumbersWidth = this.outlineView.getSlideNumbersWidth();
		}
		return this.cache.slideNumbersWidth;
	};
	DecoratorDrawer.prototype.getSlideIconSizes = function (paragraph) {

	};
	DecoratorDrawer.prototype.drawDecorShape = function (graphics, shape, x, y) {
		var t = new AscCommon.CMatrix();
		t.tx = this.getX(x);
		t.ty = this.getY(y);
		shape.draw(graphics, t, t);
	};
	DecoratorDrawer.prototype.drawTitleDecorations = function (graphics, info) {
		const slideNumberWidth = this.getSlideNumbersWidth();
		const slideX = DECORATORDRAWER_OFFSET_LEFT + slideNumberWidth;
		const decorCacheManager = this.getDecorCacheManager();
		const measureInfo = decorCacheManager.getMeasureInfo();
		const height = measureInfo.height;
		const charWidth = measureInfo.charWidth;

		const paragraph = info.outlineParagraph;
		const slide = info.slide;
		const slideTopY = this.getParagraphY(paragraph);

		const currentSlideIndex = slide.num;
		const currentSlideLabel = String(currentSlideIndex + 1);
		const slideNumberShape = decorCacheManager.getSlideNumberShape(slideNumberWidth, height, currentSlideLabel);
		this.drawDecorShape(graphics, slideNumberShape, slideX - (currentSlideLabel.length * charWidth + 1), slideTopY);

		let penHex;
		if (currentSlideIndex === this.selectedSlideIndex) {
			penHex = this.activeSlideColor;
		} else if (currentSlideIndex === this.focusSlideIndex) {
			penHex = this.hoverSlideColor;
		} else {
			penHex = this.normalSlideColor
		}
		const barShape = decorCacheManager.getTitleShape(this.backgroundColor, penHex, DECORATORDRAWER_SLIDE_WIDTH, height);
		this.drawDecorShape(graphics, barShape, slideX, slideTopY);
	};
	DecoratorDrawer.prototype.drawContentDecorations = function (graphics, info) {
		const slideNumberWidth = this.getSlideNumbersWidth();
		const slideX = DECORATORDRAWER_OFFSET_LEFT + slideNumberWidth;

		const decorCacheManager = this.getDecorCacheManager();
		const measureInfo = decorCacheManager.getMeasureInfo();
		const height = measureInfo.height;
		const charWidth = measureInfo.charWidth;

		const outlineParagraph = info.getOutlineParagraph(this.outlineView);
		const contentY = this.getParagraphY(outlineParagraph);
		const label = String(info.index + 1);
		const numberWShape = charWidth * label.length + 1;
		const badgeShape = decorCacheManager.getContentShape(this.backgroundColor, this.normalSlideColor, numberWShape, height, label);
		this.drawDecorShape(graphics, badgeShape, slideX + DECORATORDRAWER_SLIDE_WIDTH - numberWShape, contentY);
	};
	DecoratorDrawer.prototype.getY = function (y) {
		return this.transform.TransformPointY(0, y);
	};
	DecoratorDrawer.prototype.getX = function (x) {
		return this.transform.TransformPointX(x, 0);
	};
	function OutlineSlide() {
		this.title = null;
		this.content = [];
		this.slide = null;
	}

	OutlineSlide.prototype.setSlide = function (pr) {
		this.slide = pr;
	};
	OutlineSlide.prototype.setTitle = function (pr) {
		this.title = pr;
	};
	OutlineSlide.prototype.addContent = function (pr) {
		this.content.push(pr);
	};
	OutlineSlide.prototype.forEachShape = function (callback) {
		let res = null;
		if (this.title) {
			res = callback(this.title, 0);
			if (res) {
				return res;
			}
		}
		for (let i = 0; i < this.content.length; i += 1) {
			res = callback(this.content[i], i + 1);
			if (res) {
				return res;
			}
		}
		return res;
	};
	OutlineSlide.prototype.getOutlineShapes = function () {
		return [this.title].concat(this.content);
	};
	function SavedPosition(startPos, endPos) {
		this.startPos = startPos || null;
		this.endPos = endPos || null;
	}

	SavedPosition.prototype.apply = function (outlineView) {
		const docContent = outlineView.getDocContent();
		if (!docContent) {
			return;
		}
		if (this.startPos) {
			docContent.RemoveSelection();
			const startPos = this.startPos.getPosition(outlineView);
			let endPos;
			if (this.endPos) {
				endPos = this.endPos.getPosition(outlineView);
			}
			if (startPos && endPos) {
				docContent.SetContentSelection(startPos, endPos, 0, 0, 0);
			} else if (startPos) {
				docContent.SetContentPosition(startPos, 0, 0);
			} else if (endPos) {
				docContent.SetContentPosition(endPos, 0, 0);
			}

			outlineView.updateSelectionState();
		}
	};
	function ContentPosition(pos) {
		this.pos = pos;
	}
	ContentPosition.prototype.isDocStatePosition = function () {
		return true;
	};
	ContentPosition.prototype.getPosition = function (outlineView) {
		return outlineView.rebuildSavedPositionPos(this.pos);
	};
	function SlidePosition(slide, isStart) {
		this.slide = slide;
		this.isStart = isStart;
	}
	SlidePosition.prototype.getPosition = function (outlineView) {
		return outlineView.rebuildSavedSlidePos(this.slide, this.isStart);
	};
	SlidePosition.prototype.isDocStatePosition = function () {
		return false;
	};

	function UpdateAddedSlidesManager(outlineView, newSlides) {
		this.slides = newSlides;
		this.outlineView = outlineView;
	}
	UpdateAddedSlidesManager.prototype.update = function () {
		this.sortSlides();
		const outlineView = this.outlineView;
		for (let i = 0; i < this.slides.length; i += 1) {
			const slide = this.slides[i];
			const insertIndex = this.getInsertIndex(slide);
			outlineView.addMockTitle(insertIndex, new OutlineSlideInfo(slide));
		}
	};
	UpdateAddedSlidesManager.prototype.getInsertIndex = function (slide) {
		const outlineView = this.outlineView;
		const presentation = outlineView.getPresentation();
		const slides = presentation.Slides;
		const slideNum = slide.num;
		const nextSlide = slides[slideNum + 1];
		const nextInfo = outlineView.outlineInfo.getSlideInfoBySlide(nextSlide);
		const docContent = outlineView.getDocContent();
		if (!nextInfo) {
			return docContent.Content.length;
		}
		const outlineParagraph = nextInfo.outlineParagraph;
		for (let i = outlineParagraph.Index - 1; i >= 0; i -= 1) {
			const paragraph = docContent.Content[i];
			const sourceParagraph = outlineView.getSourceParagraph(paragraph);
			if (sourceParagraph) {
				const sourceShape = sourceParagraph.GetParentShape(true);
				if (sourceShape.parent !== slide) {
					return i + 1;
				} else if (i === 0) {
					return i;
				}
			} else {
				break;
			}
		}
		return outlineParagraph.Index;
	};
	UpdateAddedSlidesManager.prototype.sortSlides = function () {
		this.slides.sort(function (a, b) {
			return b.num - a.num;
		});
	};

	function UpdateRemovedSlidesManager(outlineView, newSlides) {
		this.slides = newSlides;
		this.outlineView = outlineView;
	}
	UpdateRemovedSlidesManager.prototype.removeMockParagraph = function (slide) {
		const outlineView = this.outlineView;
		const slideInfo = outlineView.outlineInfo.getSlideInfoBySlide(slide);
		if (slideInfo) {
			const outlineParagraph = slideInfo.outlineParagraph;
			outlineView.removeParagraph(outlineParagraph, outlineParagraph.Index);
		}
	};
	UpdateRemovedSlidesManager.prototype.update = function () {
		this.sortSlides();
		for (let i = 0; i < this.slides.length; i += 1) {
			const slide = this.slides[i];
			this.removeMockParagraph(slide);
		}
	};
	UpdateRemovedSlidesManager.prototype.sortSlides = function () {
		this.slides.sort(function (a, b) {
			return b.num - a.num;
		});
	};
	function UpdateNewParagraphsManager(outlineView, newParagraphs) {
		this.outlineView = outlineView;
		this.paragraphs = newParagraphs;
		this.slideNumberToOutlineSlide = {};

		this.paragraphIdToSlideIndex = {};
		this.paragraphIdToShape = {};
		this.paragraphIdToShapeIndex = {};
	}
	UpdateNewParagraphsManager.prototype.getOutlineSlide = function (num) {
		if (!this.slideNumberToOutlineSlide[num]) {
			const presentation = this.outlineView.getPresentation();
			this.slideNumberToOutlineSlide[num] = presentation.Slides[num].getOutlineSlide();
		}
		return this.slideNumberToOutlineSlide[num];
	};
	UpdateNewParagraphsManager.prototype.getSlideIndex = function (paragraph) {
		const paragraphId = paragraph.Get_Id();
		if (this.paragraphIdToSlideIndex[paragraphId] === undefined) {
			const shape = paragraph.GetParentShape();
			const slide = shape.parent;
			this.paragraphIdToSlideIndex[paragraphId] = slide.num;
		}
		return this.paragraphIdToSlideIndex[paragraphId];
	};
	UpdateNewParagraphsManager.prototype.getShapeIndex = function (paragraph) {
		const paragraphId = paragraph.Get_Id();
		if (this.paragraphIdToShapeIndex[paragraphId] === undefined) {
			const oThis = this;
			const slideIndex = this.getSlideIndex(paragraph);
			const outlineSlide = this.getOutlineSlide(slideIndex);
			const paragraphShape = this.getParagraphShape(paragraph);
			outlineSlide.forEachShape(function (shape, idx) {
				if (paragraphShape === shape) {
					oThis.paragraphIdToShapeIndex[paragraphId] = idx;
					return true;
				}
			});
		}
		return this.paragraphIdToShapeIndex[paragraphId];
	};
	UpdateNewParagraphsManager.prototype.getContentIndex = function (paragraph) {
		return paragraph.Index;
	};
	UpdateNewParagraphsManager.prototype.getParagraphShape = function (paragraph) {
		const paragraphId = paragraph.Get_Id();
		if (!this.paragraphIdToShape[paragraphId]) {
			this.paragraphIdToShape[paragraphId] = paragraph.GetParentShape();
		}
		return this.paragraphIdToShape[paragraphId];
	};
	UpdateNewParagraphsManager.prototype.compareParagraphs = function (aParagraph, bParagraph) {
		const aSlideIndex = this.getSlideIndex(aParagraph);
		const bSlideIndex = this.getSlideIndex(bParagraph);
		if (aSlideIndex !== bSlideIndex) {
			return aSlideIndex - bSlideIndex;
		}
		const aShapeIndex = this.getShapeIndex(aParagraph);
		const bShapeIndex = this.getShapeIndex(bParagraph);
		if (aShapeIndex !== bShapeIndex) {
			return aShapeIndex - bShapeIndex;
		}
		const aContentIndex = this.getContentIndex(aParagraph);
		const bContentIndex = this.getContentIndex(bParagraph);
		return aContentIndex - bContentIndex;
	};
	UpdateNewParagraphsManager.prototype.update = function () {
		const outlineView = this.outlineView;
		const outlineDocContent = this.outlineView.getDocContent();
		const outlineParagraphs = outlineDocContent.Content;
		for (let i = 0; i < this.paragraphs.length; i++) {
			const sourceParagraph = this.paragraphs[i];
			const insertIndex = this.getInsertIndex(sourceParagraph, outlineParagraphs);
			const shapeIndex = this.getShapeIndex(sourceParagraph);
			let isReplaceMockParagraph = false;
			if (shapeIndex === 0 && sourceParagraph.Index === 0) {
				isReplaceMockParagraph = outlineView.addUpdatedParagraphWithCheckMockParagraph(sourceParagraph, insertIndex);
			}
			if (!isReplaceMockParagraph) {
				outlineView.addUpdatedParagraph(sourceParagraph, insertIndex);
				const shape = this.getParagraphShape(sourceParagraph);
				const slide = shape.parent;
				const slideInfo = outlineView.outlineInfo.getSlideInfoBySlide(slide);
				if (slideInfo) {
					slideInfo.resetContentShapeInfoMapByShape(shape);
				}
			}
		}
	};
	UpdateNewParagraphsManager.prototype.getInsertIndex = function (sourceParagraph, outlineParagraphs) {
		const outlineView = this.outlineView;
		for (let i = outlineParagraphs.length - 1; i >= 0; i -= 1) {
			const outlineParagraph = outlineParagraphs[i];
			const sourceOutlineParagraph = outlineView.getSourceParagraph(outlineParagraph);
			if (!sourceOutlineParagraph) {
				const slideInfo = outlineView.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
				const slideIndex = slideInfo.slide.num;
				const sourceSlideIndex = this.getSlideIndex(sourceParagraph);
				if (slideIndex <= sourceSlideIndex) {
					return i;
				}
			} else if (this.compareParagraphs(sourceOutlineParagraph, sourceParagraph) < 0) {
				return i + 1;
			}
		}
		return 0;


			//todo binary search
		// if (outlineParagraphs.length === 0) {
		// 	return 0;
		// }
		// const outlineView = this.outlineView;
		// let leftPointer = 0;
		// let rightPointer = outlineParagraphs.length - 1;
		// while (leftPointer !== rightPointer) {
		// 	const centerPointer = leftPointer + Math.floor((rightPointer - leftPointer) / 2);
		// 	const outlineParagraph = outlineParagraphs[centerPointer];
		// 	const sourceOutlineParagraph = outlineView.outlineToSourceMap[outlineParagraph.Get_Id()];
		// 	const compareResult = this.compareParagraphs(sourceOutlineParagraph, sourceParagraph);
		// 	if (compareResult < 0) {
		// 		leftPointer = centerPointer;
		// 	} else {
		// 		rightPointer = centerPointer;
		// 	}
		// }
		//
		// const centerOutlineParagraph = outlineParagraphs[leftPointer];
		// const sourceOutlineParagraph = outlineView.outlineToSourceMap[centerOutlineParagraph.Get_Id()];
		// const compareResult = this.compareParagraphs(sourceOutlineParagraph, sourceParagraph);
		// if (compareResult < 0) {
		// 	return leftPointer + 1
		// }
		// return leftPointer;
	};

	function ForEachSelectManager(outlineView) {
		this.outlineView = outlineView;
	}
	ForEachSelectManager.prototype.getSelectionUseContentsInfo = function () {
		const content = this.getDocContent();
		const startPos = content.GetContentPosition(true, true);
		const endPos = content.GetContentPosition(true, false);

		const contents = [];
		const direction = content.GetSelectDirection();
		let startPosIndex = startPos[0].Position;
		let endPosIndex = endPos[0].Position;
		if (direction === AscWord.Direction.BACKWARD) {
			const temp = startPosIndex;
			startPosIndex = endPosIndex;
			endPosIndex = temp;
		}
		for (let i = startPosIndex; i <= endPosIndex; i += 1) {
			const contentInfo = contents[contents.length - 1];
			const outlineParagraph = content.Content[i];
			const sourceParagraph = this.getSourceParagraph(outlineParagraph);
			if (sourceParagraph) {
				const sourceContent = sourceParagraph.GetParent();
				if (contentInfo && contentInfo.content === sourceContent) {
					if (direction === AscWord.Direction.FORWARD) {
						contentInfo.endPos = endPos;
						contentInfo.endParagraph = sourceParagraph;
					} else {
						contentInfo.startPos = startPos;
						contentInfo.startParagraph = sourceParagraph;
					}
				} else {
					contents.push({
						startPos: startPos,
						startParagraph: sourceParagraph,
						endPos: endPos,
						endParagraph: sourceParagraph,
						content: sourceContent,
						direction: direction
					});
				}
			} else {
				contents.push({
					content: null,
					outlineParagraph: outlineParagraph,
					direction: direction
				});
			}
		}
		return contents;
	};
	ForEachSelectManager.prototype.getSourceParagraph = function (outlineParagraph) {
		return this.outlineView.getSourceParagraph(outlineParagraph);
	};
	ForEachSelectManager.prototype.getDocContent = function () {
		return this.outlineView.getDocContent();
	};
	ForEachSelectManager.prototype.processNullableContent = function (outlineInfo, callback) {
		return callback(outlineInfo);
	};
	ForEachSelectManager.prototype.processSingleSelectedContent = function (contentInfo, callback) {
		if (contentInfo.content) {
			return this.processSingleSelectedNotNullableContent(contentInfo, callback);
		}
		return this.processNullableContent({outlineParagraph: contentInfo.outlineParagraph, index: 0, count: 1}, callback);

	};
	ForEachSelectManager.prototype.processSingleSelectedNotNullableContent = function (contentInfo, callback) {
		if (contentInfo.content !== null) {
			const content = contentInfo.content;
			const startPos = rebuildPos(contentInfo.startPos, content, contentInfo.startParagraph);
			const endPos = rebuildPos(contentInfo.endPos, content, contentInfo.endParagraph);
			content.SetContentPosition(startPos, 0, 0);
			content.SetContentSelection(startPos, endPos, 0, 0, 0, 0);
			const res = callback({content: content, index: 0, count: 1});
			content.RemoveSelection();
			return res;
		}
	};
	ForEachSelectManager.prototype.processStartContent = function (contentInfo, count, callback) {
		const content = contentInfo.content;
		if (!content) {
			return this.processNullableContent({outlineParagraph: contentInfo.outlineParagraph, index: 0, count: count, direction: contentInfo.direction}, callback);
		}

		content.MoveCursorToEndPos(false, true);
		const targetPos = content.GetContentPosition(true, true);
		if (contentInfo.direction === AscWord.Direction.FORWARD) {
			const startPos = rebuildPos(contentInfo.startPos, content, contentInfo.startParagraph);
			content.SetContentSelection(startPos, targetPos, 0, 0, 0);
		} else {
			const startPos = rebuildPos(contentInfo.endPos, content, contentInfo.endParagraph);
			content.SetContentSelection(targetPos, startPos, 0, 0, 0);
		}
		const res = callback({content: content, index: 0, count: count, direction: contentInfo.direction});
		content.RemoveSelection();
		return res;
	};
	ForEachSelectManager.prototype.processCenterContents = function (contents, callback) {
		for (let i = 1; i < contents.length - 1; i += 1) {
			const contentInfo = contents[i];
			const content = contentInfo.content;
			if (!content) {
				const nullableRes = this.processNullableContent({outlineParagraph: contentInfo.outlineParagraph, index: i, count: contents.length, direction: contentInfo.direction}, callback);
				if (nullableRes) {
					return nullableRes;
				}
			} else {
				content.SelectAll(contentInfo.direction);
				const res = callback({content: content, index: i, count: contents.length, direction: contentInfo.direction});
				content.RemoveSelection();
				if (res) {
					return true;
				}
			}
		}
	};
	ForEachSelectManager.prototype.processEndContent = function (contentInfo, count, callback) {
		const content = contentInfo.content;
		if (!content) {
			return this.processNullableContent({outlineParagraph: contentInfo.outlineParagraph, index: count - 1, count: count, direction: contentInfo.direction}, callback);
		}
		content.MoveCursorToStartPos();
		const targetPos = content.GetContentPosition(true, true);
		if (contentInfo.direction === AscWord.Direction.FORWARD) {
			const endPos = rebuildPos(contentInfo.endPos, content, contentInfo.endParagraph);
			content.SetContentSelection(targetPos, endPos, 0, 0, 0);
		} else {
			const endPos = rebuildPos(contentInfo.startPos, content, contentInfo.startParagraph);
			content.SetContentSelection(endPos, targetPos, 0, 0, 0);
		}
		const res = callback({content: content, index: count - 1, count: count, direction: contentInfo.direction});
		content.RemoveSelection();
		return res;
	};
	ForEachSelectManager.prototype.processMultipleSelectedContent = function (contents, callback) {
		const contentLength = contents.length;
		let res;
		res = this.processStartContent(contents[0], contentLength, callback);
		if (res) {
			return res;
		}
		res = this.processCenterContents(contents, callback);
		if (res) {
			return res;
		}
		return this.processEndContent(contents[contents.length - 1], contentLength, callback);
	};
	ForEachSelectManager.prototype.processSelectedContent = function (callback) {
		const contents = this.getSelectionUseContentsInfo();
		if (contents.length === 1) {
			const contentInfo = contents[0];
			return this.processSingleSelectedContent(contentInfo, callback);
		}
		if (contents.length > 1) {
			return this.processMultipleSelectedContent(contents, callback);
		}
	};
	ForEachSelectManager.prototype.processCursorContent = function (callback) {
		const content = this.getDocContent();
		const contentPos = content.GetContentPosition(false, false);
		const paragraph = content.Content[content.CurPos.ContentPos];
		const sourceParagraph = this.getSourceParagraph(paragraph);
		if (sourceParagraph) {
			return this.processRealCursorContent(sourceParagraph, contentPos, callback);
		}
		return this.processNullableContent({outlineParagraph: paragraph, index: 0, count: 1}, callback);
	};
	ForEachSelectManager.prototype.processRealCursorContent = function (sourceParagraph, contentPos, callback) {
		const sourceContent = sourceParagraph.GetParent();
		const startPos = rebuildPos(contentPos, sourceContent, sourceParagraph);
		sourceContent.SetContentPosition(startPos, 0, 0);
		const res = callback({content: sourceContent, index: 0, count: 1});
		sourceContent.RemoveSelection();
		return res;
	};
	ForEachSelectManager.prototype.forEachSelectedContent = function (callback) {
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		if (content.IsSelectionUse()) {
			return this.processSelectedContent(callback);
		}
		return this.processCursorContent(callback);
	};

	function RemoveOutlineParagraphsManager(outlineView) {
		this.outlineView = outlineView;
		this.checkSlidesForRemove = {};
		this.saveManager = new SavePosManager(outlineView);
		this.concatShapesManager = new ConcatShapesManager(outlineView);
	}
	RemoveOutlineParagraphsManager.prototype.getDocContent = function () {
		return this.outlineView.getDocContent();
	};
	RemoveOutlineParagraphsManager.prototype.getSourceParagraph = function (outlineParagraph) {
		return this.outlineView.getSourceParagraph(outlineParagraph);
	};
	RemoveOutlineParagraphsManager.prototype.getOutlineParagraph = function (sourceParagraph) {
		return this.outlineView.sourceToOutlineMap[sourceParagraph.Get_Id()];
	};
	RemoveOutlineParagraphsManager.prototype.getContentPosParagraph = function (docContent) {
		return docContent.Content[docContent.CurPos.ContentPos];
	};
	RemoveOutlineParagraphsManager.prototype.processNotSelectedContent = function () {
		if (this.isNeedProcessCursorAtBegin()) {
			this.processCursorAtBegin();
		} else {
			this.processCommonPosContent();
		}
	};
	RemoveOutlineParagraphsManager.prototype.processCommonPosContent = function () {
		const oThis = this;
		const outlineView = this.outlineView;
		outlineView.forEachSelectedContent(function (selectProps) {
			const content = selectProps.content;
			if (content) {
				content.Remove(-1, true, false, false);
				oThis.saveManager.saveContentPosSourceParagraph(content);
			}
		});
	};
	RemoveOutlineParagraphsManager.prototype.processCursorAtBegin = function () {
		const docContent = this.getDocContent();
		const contentPos = docContent.GetContentPosition(false, true);
		const outlineParagraph = contentPos[1].Class;
		if (outlineParagraph.Index === 0) {
			return;
		}
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			this.processRealCursorAtBegin(outlineParagraph);
		} else {
			this.processNullableContent(outlineParagraph, true);
		}
	};
	RemoveOutlineParagraphsManager.prototype.processRealCursorAtBegin = function (outlineParagraph) {
		const docContent = this.getDocContent();
		const currentSourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (currentSourceParagraph && (currentSourceParagraph.Index !== 0 || currentSourceParagraph.IsPresentationBulleted())) {
			this.processCommonPosContent();
		} else {
			const previousOutlineParagraph = docContent.Content[outlineParagraph.Index - 1];
			this.checkRemoveEdgeParagraphs(previousOutlineParagraph, outlineParagraph);
		}
	};
	RemoveOutlineParagraphsManager.prototype.checkRemoveEdgeParagraphs = function (firstParagraph, lastParagraph, isSkipConcatContents) {
		const lastSourceParagraph = this.getSourceParagraph(lastParagraph);
		if (!lastSourceParagraph) {
			return;
		}
		const lastContent = lastSourceParagraph.GetParent();
		const firstSourceParagraph = this.getSourceParagraph(firstParagraph);
		if (firstSourceParagraph) {
			const previousParagraphContent = firstSourceParagraph.GetParent();
			this.concatEdgeContents(previousParagraphContent, lastContent, isSkipConcatContents);
		} else {
			const titleDocContent = this.outlineView.createTitleContent(firstParagraph);
			this.concatEdgeContents(titleDocContent, lastContent);
		}
	};
	RemoveOutlineParagraphsManager.prototype.isNeedProcessCursorAtBegin = function () {
		const docContent = this.getDocContent();
		const paragraph = this.getContentPosParagraph(docContent);
		return paragraph.IsCursorAtBegin();
	};
	RemoveOutlineParagraphsManager.prototype.processSourceSelectedContent = function (content, idx, count, isOnAddText, edgeContents) {
		if (content.IsSelectedAll() || content.IsEmpty()) {
			const shape = content.Is_DrawingShape(true);
			const slide = shape.parent;
			const slideNum = slide.num;
			this.checkSlidesForRemove[slideNum] = slide;
			content.Remove(1, true, false, false);
		} else {
			content.Remove(1, true, false, count === 1 && isOnAddText);
			if (idx === 0) {
				edgeContents.firstContent = content;
			} else if (idx === count - 1) {
				edgeContents.lastContent = content;
			}
		}
	};
	RemoveOutlineParagraphsManager.prototype.saveFirstOutlineParagraph = function (edgeContents, selectProps) {
		if (selectProps.index !== 0) {
			return;
		}

		if (selectProps.outlineParagraph) {
			const outlineParagraph = selectProps.outlineParagraph;
			edgeContents.firstOutlineParagraph = outlineParagraph;
		} else {
			const content = selectProps.content;
			const sourceParagraph = content.Content[0];
			const outlineParagraph = sourceParagraph && this.getOutlineParagraph(sourceParagraph);
			if (outlineParagraph) {
				edgeContents.firstOutlineParagraph = outlineParagraph;
			}
		}
	}
	RemoveOutlineParagraphsManager.prototype.processSelectedContent = function (isOnAddText) {
		const outlineView = this.outlineView;
		const edgeContents = {firstContent: null, lastContent: null, firstOutlineParagraph: null};
		const oThis = this;
		let isSingleContent = false;
		outlineView.forEachSelectedContent(function (selectProps) {
			oThis.saveFirstOutlineParagraph(edgeContents, selectProps);
			if (selectProps.outlineParagraph) {
				const outlineParagraph = selectProps.outlineParagraph;
				oThis.processNullableContent(outlineParagraph);
			} else {
				const content = selectProps.content;
				const idx = selectProps.index;
				const count = selectProps.count;
				oThis.processSourceSelectedContent(content, idx, count, isOnAddText, edgeContents);
				if (count === 1) {
					isSingleContent = true;
					oThis.saveManager.saveContentPosSourceParagraph(content);
				}
			}
		});

		if (isSingleContent) {
			return;
		}

		if (edgeContents.firstContent && edgeContents.lastContent) {
			this.concatEdgeContents(edgeContents.firstContent, edgeContents.lastContent);
		} else if (edgeContents.firstContent) {
			const content = edgeContents.firstContent.Content;
			this.saveManager.saveEndPosSourceParagraph(content[content.length - 1]);
		} else if (edgeContents.lastContent && edgeContents.firstOutlineParagraph) {
			this.processLastContent(edgeContents.lastContent, edgeContents.firstOutlineParagraph);
		}
	};
	RemoveOutlineParagraphsManager.prototype.processLastContent = function (lastContent, firstOutlineParagraph) {
		const content = lastContent.Content;
		const shape = lastContent.Is_DrawingShape(true);
		if (shape.isGeneralTitle()) {
			this.saveManager.saveStartPosSourceParagraph(content[0]);
		} else if (firstOutlineParagraph.Index !== 0) {
			const outlineContent = this.outlineView.getDocContent();
			const previousOutlineParagraph = outlineContent.Content[firstOutlineParagraph.Index - 1];
			const lastContentSourceParagraph = content[0];
			const lastOutlineParagraph = this.getOutlineParagraph(lastContentSourceParagraph);
			this.checkRemoveEdgeParagraphs(previousOutlineParagraph, lastOutlineParagraph, true);
		}
	};
	RemoveOutlineParagraphsManager.prototype.remove = function (isOnAddText) {
		const docContent = this.getDocContent();
		if (isOnAddText) {
			if (!docContent.IsSelectionUse()) {
				return;
			}
		}
		const outlineView = this.outlineView;
		if (docContent.IsSelectionUse()) {
			this.processSelectedContent(isOnAddText);
		} else {
			this.processNotSelectedContent();
		}
		outlineView.removeSlides(this.checkSlidesForRemove);
	};
	function SavePosManager(outlineView) {
		this.outlineView = outlineView;
	}
	SavePosManager.prototype.getSourceParagraph = function (outlineParagraph) {
		return this.outlineView.getSourceParagraph(outlineParagraph);
	};
	SavePosManager.prototype.saveEndPosParagraph = function (outlineParagraph) {
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			this.saveEndPosSourceParagraph(sourceParagraph);
		}
	};
	SavePosManager.prototype.saveContentPosSourceParagraph = function (sourceContent) {
		const content = sourceContent;
		const contentPos = content.GetContentPosition(false, true);
		this.outlineView.setSavedPosition(new SavedPosition(new ContentPosition(contentPos)));
	};
	SavePosManager.prototype.saveEndPosSourceParagraph = function (sourceParagraph) {
		const contentPos = [{Class: sourceParagraph.GetParent(), Position: sourceParagraph.GetIndex()}];
		sourceParagraph.GetEndContentPosition(contentPos);
		this.outlineView.setSavedPosition(new SavedPosition(new ContentPosition(contentPos)));
	};
	SavePosManager.prototype.saveStartPosSourceParagraph = function (sourceParagraph) {
		const contentPos = [{Class: sourceParagraph.GetParent(), Position: sourceParagraph.GetIndex()}];
		sourceParagraph.GetStartContentPosition(contentPos);
		this.outlineView.setSavedPosition(new SavedPosition(new ContentPosition(contentPos)));
	};
	RemoveOutlineParagraphsManager.prototype.processNullableContent = function (outlineParagraph, isEdgeParagraph) {
		const outlineView = this.outlineView;
		const info = outlineView.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		const slide = info && info.slide;
		if (!slide) {
			return;
		}
		this.checkSlidesForRemove[slide.num] = slide;
		outlineView.checkParagraph(outlineParagraph);

		if (isEdgeParagraph) {
			this.processEdgeNullableContent(outlineParagraph);
		}
	};
	RemoveOutlineParagraphsManager.prototype.processEdgeNullableContent = function (outlineParagraph) {
		const docContent = this.getDocContent();
		const nextOutlineParagraph = docContent.Content[outlineParagraph.Index + 1];
		if (!nextOutlineParagraph) {
			return;
		}

		const previousOutlineParagraph = docContent.Content[outlineParagraph.Index - 1];
		const previousSourceParagraph = this.getSourceParagraph(previousOutlineParagraph);

		const nextSourceParagraph = this.getSourceParagraph(nextOutlineParagraph);
		const nextInfo = this.outlineView.outlineInfo.getSlideInfoByParagraph(nextOutlineParagraph);
		if (!nextInfo) {
			if (previousSourceParagraph) {
				const firstContent = previousSourceParagraph.GetParent();
				const lastContent = nextSourceParagraph.GetParent();
				this.concatEdgeContents(firstContent, lastContent, true);
			} else {
				const titleDocContent = this.outlineView.createTitleContent(previousOutlineParagraph);
				const nextDocContent = nextSourceParagraph.GetParent();
				this.concatEdgeContents(titleDocContent, nextDocContent);
			}
		}
		this.saveManager.saveEndPosParagraph(previousOutlineParagraph);
	};
	RemoveOutlineParagraphsManager.prototype.checkRemoveSlides = function (firstIndex, lastIndex) {
		const presentation = this.outlineView.getPresentation();
		for (let i = firstIndex; i <= lastIndex; i += 1) {
			this.checkSlidesForRemove[i] = presentation.Slides[i];
		}
	}

	function ConcatShapesManager(outlineView) {
		this.outlineView = outlineView;
	}
	ConcatShapesManager.prototype.concatContents = function (firstContent, lastContent, isSkipConcatParagraph) {
		const firstShape = firstContent.Is_DrawingShape(true);
		const lastShape = lastContent.Is_DrawingShape(true);
		const firstSlide = firstShape.parent;
		const lastSlide = lastShape.parent;

		if (!isSkipConcatParagraph) {
			const firstContentParagraph = firstContent.Content[firstContent.Content.length - 1];

			firstContentParagraph.Concat(lastContent.Content[0]);
			lastContent.Remove_FromContent(0, 1);
		}
		let isNeedUseNextShapeForInsert = false;
		if (firstShape.isGeneralTitle()) {
			isNeedUseNextShapeForInsert = true;
		}


		if (firstSlide !== lastSlide) {
			this.processEdgeSlides(firstShape, lastShape, isNeedUseNextShapeForInsert, !isSkipConcatParagraph);
		}
	};
	ConcatShapesManager.prototype.processEdgeSlides = function (firstShape, lastShape, isNeedUseNextShapeForInsert, isConcatenatedFirstParagraph) {
		const firstSlide = firstShape.parent;
		const lastSlide = lastShape.parent;
		const firstOutlineSlide = firstSlide.getOutlineSlide();
		const lastOutlineSlide = lastSlide.getOutlineSlide();
		const firstOutlineShapes = firstOutlineSlide.getOutlineShapes();
		const lastOutlineShapes = lastOutlineSlide.getOutlineShapes();

		let firstShapeIndex = firstOutlineShapes.indexOf(firstShape);
		if (isNeedUseNextShapeForInsert) {
			firstShapeIndex += 1;
			if (firstShapeIndex === firstOutlineShapes.length) {
				const content = this.outlineView.createSubtitleContent(firstSlide);
				firstOutlineShapes.push(content.Is_DrawingShape(true));
			}
		}
		let lastShapeIndex = lastOutlineShapes.indexOf(lastShape);

		for (let i = firstShapeIndex; i < firstOutlineShapes.length; i += 1) {
			while (lastShapeIndex <= i) {
				if (lastShapeIndex === lastOutlineShapes.length) {
					return;
				}
				const isSkipFirstParagraph = isConcatenatedFirstParagraph && lastShape === lastOutlineShapes[lastShapeIndex];
				this.concatShapeContent(firstOutlineShapes[i], lastOutlineShapes[lastShapeIndex], isSkipFirstParagraph);
				lastShapeIndex += 1;
			}
		}
		const edgeFirstShape = firstOutlineShapes[firstOutlineShapes.length - 1];
		for (let i = lastShapeIndex; i < lastOutlineShapes.length; i += 1) {
			const isSkipFirstParagraph = isConcatenatedFirstParagraph && lastShape === lastOutlineShapes[i];
			this.concatShapeContent(edgeFirstShape, lastOutlineShapes[i], isSkipFirstParagraph);
		}
	};
	ConcatShapesManager.prototype.concatShapeContent = function (firstShape, lastShape, skipFirstParagraph) {
		const lastContent = lastShape.getDocContent();
		if (lastContent.IsEmpty()) {
			return;
		}
		const firstContent = firstShape.getDocContent();
		let i = skipFirstParagraph ? 1 : 0;
		if (lastContent.Content[i] && firstContent.IsEmpty()) {
			const firstContentParagraph = firstContent.Content[firstContent.Content.length - 1];
			firstContentParagraph.Concat(lastContent.Content[i]);
			i += 1;
		}
		for (; i < lastContent.Content.length; i += 1) {
			firstContent.Add_ToContent(firstContent.Content.length, lastContent.Content[i].Copy(firstContent));
		}

		lastContent.Clear_Content();
	}

	RemoveOutlineParagraphsManager.prototype.concatEdgeContents = function (firstContent, lastContent, isSkipConcatParagraph) {

		const firstShape = firstContent.Is_DrawingShape(true);
		const lastShape = lastContent.Is_DrawingShape(true);
		const firstSlide = firstShape.parent;
		const lastSlide = lastShape.parent;
		const firstContentParagraph = firstContent.Content[firstContent.Content.length - 1];
		this.saveManager.saveEndPosSourceParagraph(firstContentParagraph);
		this.concatShapesManager.concatContents(firstContent, lastContent, isSkipConcatParagraph);
		this.checkRemoveSlides(firstSlide.num + 1, lastSlide.num);
	};

	function ParagraphAddManager(outlineView) {
		this.outlineView = outlineView;
	}
	ParagraphAddManager.prototype.forEachSelectedContent = function (callback) {
		return this.outlineView.forEachSelectedContent(callback);
	};
	ParagraphAddManager.prototype.setSavedRealContentState = function (content, idx, contentCount, direction) {
		return this.outlineView.setSavedRealContentState(content, idx, contentCount, direction);
	};
	ParagraphAddManager.prototype.processMockParagraphAdd = function (outlineParagraph, idx, contentCount, direction, paraItem) {
		const docContent = this.outlineView.createTitleContent(outlineParagraph);
		if (docContent) {
			this.processSourceContentAdd(docContent, idx, contentCount, direction, paraItem);
		}
	};
	ParagraphAddManager.prototype.processSourceContentAdd = function (content, idx, contentCount, direction, paraItem) {
		content.AddToParagraph(paraItem);
		this.setSavedRealContentState(content, idx, contentCount, direction);
	};
	ParagraphAddManager.prototype.processParagraphAdd = function (paraItem) {
		const oThis = this;
		this.forEachSelectedContent(function (selectProps) {
			const idx = selectProps.index;
			const contentCount = selectProps.count;
			const direction = selectProps.direction;
			if (selectProps.outlineParagraph) {
				const outlineParagraph = selectProps.outlineParagraph;
				oThis.processMockParagraphAdd(outlineParagraph, idx, contentCount, direction, paraItem);
			} else {
				const content = selectProps.content;
				oThis.processSourceContentAdd(content, idx, contentCount, direction, paraItem);
			}
		});
	};
	ParagraphAddManager.prototype.processRemoveContentTypes = function () {
		this.outlineView.remove(true);
	};
	ParagraphAddManager.prototype.paragraphAdd = function (paraItem) {
		switch (paraItem.Type) {
			case para_Math:
			case para_NewLine:
			case para_Text:
			case para_Space:
			case para_Tab:
			case para_PageNum:
			case para_Field:
			case para_FootnoteReference:
			case para_FootnoteRef:
			case para_Separator:
			case para_ContinuationSeparator:
			case para_InstrText:
			case para_EndnoteReference:
			case para_EndnoteRef: {
				this.processRemoveContentTypes();
				break;
			}
			default: {
				break;
			}
		}
		this.processParagraphAdd(paraItem);
	};

	function OutlineContentInsertManager(outlineView) {
		this.outlineView = outlineView;
	}
	OutlineContentInsertManager.prototype.insertIntoSourceContent = function (content, selectedContent) {
		const paragraph = content.Content[content.CurPos.ContentPos];
		if (!paragraph || !paragraph.IsParagraph()) {
			return false;
		}
		const nearPos = {Paragraph: paragraph, ContentPos: paragraph.Get_ParaContentPos(false, false)};
		paragraph.Check_NearestPos(nearPos);
		selectedContent.Insert(nearPos);
		return true;
	};
	OutlineContentInsertManager.prototype.insert = function (selectedContent) {
		const outlineView = this.outlineView;
		const oThis = this;
		const presentation = outlineView.getPresentation();
		outlineView.remove(true);
		let inserted = false;
		outlineView.forEachSelectedContent(function (selectProps) {
			let content = selectProps.content;
			if (!content && selectProps.outlineParagraph) {
				content = outlineView.createTitleContent(selectProps.outlineParagraph);
			}
			if (content) {
				const shape = content.Is_DrawingShape(true);
				presentation.PrepareInsertedDocContent(selectedContent, content, shape);
				selectedContent.EndCollect(outlineView.getPresentation());
				inserted = oThis.insertIntoSourceContent(content, selectedContent);
				outlineView.setSavedRealContentState(content, selectProps.index, selectProps.count, selectProps.direction);
			}
			return inserted;
		});
		return inserted;
	};

	function OutlineInfo() {
		this.outlineParagraphToInfoMap = {};
		this.slideToInfoMap = {};
	}
	OutlineInfo.prototype.deleteSlideInfoByParagraph = function (outlineParagraph) {
		const id = outlineParagraph.GetId();
		const info = this.outlineParagraphToInfoMap[id];
		if (info) {
			delete this.outlineParagraphToInfoMap[id];
			delete this.slideToInfoMap[info.slide.GetId()];
		}
	}
	OutlineInfo.prototype.addSlideInfo = function (slideInfo) {
		const outlineParagraph = slideInfo.outlineParagraph;
		const slide = slideInfo.slide;
		this.outlineParagraphToInfoMap[outlineParagraph.GetId()] = slideInfo;
		this.slideToInfoMap[slide.GetId()] = slideInfo;
	};
	OutlineInfo.prototype.getSlideInfoByParagraph = function (outlineParagraph) {
		return this.outlineParagraphToInfoMap[outlineParagraph.GetId()];
	};
	OutlineInfo.prototype.getSlideInfoBySlide = function (slide) {
		return this.slideToInfoMap[slide.GetId()];
	};
	OutlineInfo.prototype.getOutlineParagraphToInfoMap = function () {
		return this.outlineParagraphToInfoMap;
	};
	function OutlineContentShapeInfo(shape, index) {
		this.shape = shape;
		this.index = index;
	}
	OutlineContentShapeInfo.prototype.getOutlineParagraph = function (outlineView) {
		const shape = this.shape;
		const docContent = shape.getDocContent();
		const firstSourceParagraph = docContent.Content[0];
		return outlineView.sourceToOutlineMap[firstSourceParagraph.GetId()];
	};
	function OutlineSlideInfo(slide, outlineParagraph) {
		this.slide = slide;
		this.outlineParagraph = outlineParagraph || null;
		this.contentShapeInfoMap = null;
	}
	OutlineSlideInfo.prototype.setOutlineParagraph = function (outlineParagraph) {
		this.outlineParagraph = outlineParagraph;
	}
	OutlineSlideInfo.prototype.resetContentShapeInfoMap = function () {
		this.contentShapeInfoMap = null;
	};
	OutlineSlideInfo.prototype.resetContentShapeInfoMapByShape = function (shape) {
		if (this.contentShapeInfoMap && !this.contentShapeInfoMap[shape.GetId()]) {
			this.resetContentShapeInfoMap();
		}
	};
	OutlineSlideInfo.prototype.getContentShapeInfoMap = function () {
		if (this.contentShapeInfoMap === null) {
			this.contentShapeInfoMap = {};
			const outlineSlide = this.slide.getOutlineSlide();
			if (outlineSlide.content.length > 1) {
				for (let i = 0; i < outlineSlide.content.length; i += 1) {
					const shape = outlineSlide.content[i];
					if (shape.isOutlinePlaceholderWithContent()) {
						this.contentShapeInfoMap[shape.GetId()] = new OutlineContentShapeInfo(shape, i);
					}
				}
			}
		}
		return this.contentShapeInfoMap;
	};

	function DecorCacheManager() {
		this.titleShape = null;
		this.contentShape = null;
		this.slideNumberShape = null;
		this.emptySlideShape = null;
		this.measureInfo = null;
		this.resetColors();
	}
	DecorCacheManager.prototype.resetColors = function () {
		this.brushes = {};
		this.pens = {};
	};
	DecorCacheManager.prototype.getBrush = function (hex) {
		if (!hex) {
			return null;
		}
		if (!this.brushes[hex]) {
			const rgb = getRGBFromHex(hex);
			this.brushes[hex] = AscFormat.CreateSolidFillRGBA(rgb.R, rgb.G, rgb.B, 255);
		}
		return this.brushes[hex];
	};
	DecorCacheManager.prototype.getPen = function (hex, penW) {
		if (!hex) {
			return null;
		}
		if (!this.pens[hex]) {
			const brush = this.getBrush(hex);
			var pen = new AscFormat.CLn();
			pen.Fill = brush;
			this.pens[hex] = pen;
		}
		this.pens[hex].w = penW;
		return this.pens[hex];
	};
	DecorCacheManager.prototype.getTitleShape = function (brushHex, penHex, width, height) {
		if (this.titleShape === null) {
			this.titleShape = this.createDecorShape(width, height, "roundRect");
		}
		this.recalculateShape(this.titleShape, brushHex, penHex, 20000, width, height)
		return this.titleShape;
	};
	DecorCacheManager.prototype.getContentShape = function (brushHex, penHex, width, height, strContent) {
		if (this.contentShape === null) {
			this.contentShape = this.createDecorShape(width, height, "rect");
		}
		this.recalculateShape(this.contentShape, brushHex, penHex, 0, width, height);
		this.recalculateStrContent(this.contentShape, strContent);
		return this.contentShape;
	};
	DecorCacheManager.prototype.getSlideNumberShape = function (width, height, strContent) {
		if (this.slideNumberShape === null) {
			this.slideNumberShape = this.createDecorShape(width, height, "rect", AscCommon.align_Left);
		}
		this.recalculateShape(this.slideNumberShape, null, null, 0, width, height);
		this.recalculateStrContent(this.slideNumberShape, strContent);
		return this.slideNumberShape;
	};
	DecorCacheManager.prototype.getEmptySlideShape = function (width, height) {
		if (this.emptySlideShape === null) {
			this.emptySlideShape = this.createDecorShape(width, height, "rect", AscCommon.align_Left);
			this.recalculateStrContent(this.emptySlideShape, AscCommon.translateManager.getValue("Click to add first slide"), 16);
		}
		this.recalculateShape(this.emptySlideShape, null, null, 0, width, height, true);
		return this.emptySlideShape;
	};
	DecorCacheManager.prototype.checkShapeSizes = function (shape, width, height, isRecalculateContent) {
		if (!(AscFormat.fApproxEqual(shape.extX, width) && AscFormat.fApproxEqual(shape.extY, height))) {
			this.recalculateShapeSizes(shape, width, height);
			if (isRecalculateContent) {
				shape.recalculateContent();
			}
		}
	};
	DecorCacheManager.prototype.recalculateShapeSizes = function (shape, width, height) {
		shape.extX = width;
		shape.extY = height;

		if (shape.calcGeometry) {
			shape.calcGeometry.Recalculate(width, height);
		}
	};
	DecorCacheManager.prototype.recalculateShape = function (shape, brushHex, penHex, penW, width, height, isRecalculateContent) {
		this.checkShapeSizes(shape, width, height,isRecalculateContent);
		shape.brush = this.getBrush(brushHex);
		shape.pen = this.getPen(penHex, penW);
	};
	DecorCacheManager.prototype.recalculateStrContent = function (shape, label, fontSize) {
		AscFormat.ExecuteNoHistory(function () {
			const docContent = shape.getDocContent();
			const para = docContent.Content[0];
			var run = new ParaRun(para, false);
			var textPr = getOutlineTextPr();
			if (fontSize !== undefined) {
				textPr.SetFontSize(fontSize);
			}
			const rgb = getRGBFromHex(AscCommon.GlobalSkin.ThumbnailsPageNumberText);
			textPr.SetColor(rgb.R, rgb.G, rgb.B);
			run.Set_Pr(textPr);
			run.AddText(label);
			para.Content[0] = run;
			para.OnContentChange();
			shape.recalculateContent();
		}, this, []);
	};
	DecorCacheManager.prototype.createDecorShape = function (width, height, geometryType, paraAlign) {
		paraAlign = typeof paraAlign === "number" ? paraAlign : AscCommon.align_Center;
		return AscFormat.ExecuteNoHistory(function () {
			var shape = new AscFormat.CShape();
			shape.setBDeleted(false);
			shape.extX = width;
			shape.extY = height;

			geometryType = geometryType || "rect";
			shape.calcGeometry = AscFormat.CreateGeometry(geometryType);
			shape.calcGeometry.Recalculate(width, height);

			var txBody = shape.createTextBody();
			txBody.bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
			txBody.bodyPr.setInsets(0, 0, 0, 0);

			var para = txBody.content.Content[0];
			var paraPr = new CParaPr();
			paraPr.Jc = paraAlign;
			para.SetPr(paraPr);
			return shape;
		}, this, []);
	};
	DecorCacheManager.prototype.getMeasureInfo = function () {
		if (this.measureInfo === null) {
			g_oTextMeasurer.SetTextPr(getOutlineTextPr(), null);
			g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);
			const height = g_oTextMeasurer.GetHeight();
			let maxWidth = 0
			for (let i = 0; i <= 9; i += 1) {
				const measureInfo = g_oTextMeasurer.MeasureCode(48 + i);
				if (maxWidth < measureInfo.Width) {
					maxWidth = measureInfo.Width;
				}
			}
			this.measureInfo = {height: height, charWidth: maxWidth};
		}
		return this.measureInfo;
	};

	function OutlineView(api) {
		this.api = api;
		this.isInit = false;
		this.outlineShape = null;
		this.decorCacheManager = new DecorCacheManager();
		this.reset();
	}
	OutlineView.prototype.getSlideNumbersWidth = function () {
			const measureInfo = this.decorCacheManager.getMeasureInfo();
			const presentation = this.getPresentation();
			const slidesCount = presentation.GetSlidesCount();
			const firstSlideNumber = presentation.getFirstSlideNumber();
			const totalSlidesLength = String(slidesCount + firstSlideNumber).length;
			return measureInfo.charWidth * totalSlidesLength;
	};
	OutlineView.prototype.createTitleContent = function (outlineParagraph) {
		const slideInfo = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		if (!slideInfo) {
			return;
		}
		const slide = slideInfo.slide;
		const createdSp = slide.createTitle();
		createdSp.clearContent();
		const docContent = createdSp.getDocContent();
		docContent.Update_ContentIndexing();
		const sourceParagraph = docContent.Content[0];
		this.removeOutlineParagraph(outlineParagraph);
		this.addOutlineParagraph(sourceParagraph, outlineParagraph, slideInfo);
		const newOutlineParagraph = this.updateFromSourceParagraph(sourceParagraph);
		newOutlineParagraph.Index = outlineParagraph.Index;
		newOutlineParagraph.MoveCursorToStartPos();
		docContent.MoveCursorToStartPos();
		return docContent;
	};

	OutlineView.prototype.createSubtitleContent = function (slide) {
		const createdSp = slide.createContent();
		createdSp.clearContent();
		const docContent = createdSp.getDocContent();
		docContent.Update_ContentIndexing();
		docContent.MoveCursorToStartPos();
		return docContent;
	};

	OutlineView.prototype.reset = function () {
		this.outlineToSourceMap = {};
		this.sourceToOutlineMap = {};
		this.outlineInfo = new OutlineInfo();
		this.resetMapToCheckParagraphs();
		this.resetMapToCheckSlides();
		this.resetPosition();
	};
	OutlineView.prototype.getPresentation = function () {
		return this.api.private_GetLogicDocument();
	};
	OutlineView.prototype.setOutlineShape = function (shape) {
		this.outlineShape = shape;
	};
	OutlineView.prototype.updateAll = function (width, height) {
		this.isInit = true;
		const presentation = this.getPresentation();
		const outlineSlides = [];
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			const outlineSlide = slide.getOutlineSlide();
			outlineSlides.push(outlineSlide);
		}
		this.createOutlineShape(outlineSlides, width, height);
	};
	OutlineView.prototype.addOutlineParagraph = function (sourceParagraph, outlineParagraph, pr) {
		const outlineId = outlineParagraph.Get_Id();
		this.outlineToSourceMap[outlineId] = sourceParagraph;
		if (sourceParagraph) {
			this.sourceToOutlineMap[sourceParagraph.Get_Id()] = outlineParagraph;
		}
		if (pr) {
			pr.setOutlineParagraph(outlineParagraph);
			this.outlineInfo.addSlideInfo(pr);
		}
	};
	OutlineView.prototype.removeOutlineParagraph = function (outlineParagraph) {
		const outlineId = outlineParagraph.Get_Id();
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		delete this.outlineToSourceMap[outlineId];
		this.outlineInfo.deleteSlideInfoByParagraph(outlineParagraph);
		if (sourceParagraph) {
			delete this.sourceToOutlineMap[sourceParagraph.Get_Id()];
		}
	};
	OutlineView.prototype.addCopyParagraph = function (paragraph, pos, isFirstParagraph, pr) {
		return AscFormat.ExecuteNoHistory(function () {
			const outlineContent = this.getDocContent();
			const copyParagraph = this.getCopyParagraph(outlineContent, paragraph, !!(pr && pr.slide));
			outlineContent.AddToContent(pos, copyParagraph);
			this.addOutlineParagraph(paragraph, copyParagraph, isFirstParagraph ? pr : null);
			return copyParagraph;
		}, this, []);
	}
	OutlineView.prototype.removeParagraph = function (outlineParagraph, pos) {
		AscFormat.ExecuteNoHistory(function () {
			const outlineContent = this.getDocContent();
			outlineContent.Remove_FromContent(pos, 1, false);
			this.removeOutlineParagraph(outlineParagraph);
		}, this, []);
	}
	OutlineView.prototype.addContentToOutlineShape = function (outlineShape, slideShape, pr) {
		if (!slideShape) {
			return;
		}
		const outlineContent = outlineShape.txBody.content;
		const slideContent = slideShape.txBody.content;
		const paragraphs = slideContent.Content;
		for (let i = 0; i < paragraphs.length; i += 1) {
			const paragraph = paragraphs[i];
			this.addCopyParagraph(paragraph, outlineContent.Content.length, i === 0, pr);
		}
	};
	OutlineView.prototype.addMockTitleToOutlineShape = function (outlineShape, insertIndex, pr) {
		AscFormat.ExecuteNoHistory(function () {
			const outlineContent = outlineShape.txBody.content;
			const paragraph = this.getMockTitleParagraph(outlineContent);
			outlineContent.AddToContent(insertIndex, paragraph);
			this.addOutlineParagraph(null, paragraph, pr);
		}, this, []);

	};
	OutlineView.prototype.addMockTitle = function (insertIndex, pr) {
		this.addMockTitleToOutlineShape(this.outlineShape, insertIndex, pr);
	};
	OutlineView.prototype.getMockTitleParagraph = function (content) {
		return AscFormat.ExecuteNoHistory(function () {
			const paragraph = new AscWord.Paragraph(content, true);
			this.applyParagraphProps(paragraph, null, true);
			return paragraph;
		}, this, []);

	};
	OutlineView.prototype.fillOutlineShape = function (outlineShape, outlineSlides) {
		const paragraphs = outlineShape.txBody.content.Content;
		for (let i = 0; i < outlineSlides.length; i += 1) {
			const slide = outlineSlides[i];

			const titlePr = new OutlineSlideInfo(slide.slide);
			if (slide.title !== null) {
				this.addContentToOutlineShape(outlineShape, slide.title, titlePr);
			} else {
				this.addMockTitleToOutlineShape(outlineShape, paragraphs.length, titlePr);
			}

			let shapeCount = 0;
			for (let j = 0; j < slide.content.length; j += 1) {
				const contentShape = slide.content[j];
				if (!contentShape.txBody.content.IsEmpty()) {
					this.addContentToOutlineShape(outlineShape, slide.content[j], null);
					shapeCount += 1;
				}
			}
		}
	};
	OutlineView.prototype.getParaPr = function (compiledParaPr) {
		const copyParaPr = new CParaPr();
		if (compiledParaPr.ParaPr.Bullet) {
			copyParaPr.Bullet = compiledParaPr.ParaPr.Bullet.createDuplicate();
			copyParaPr.Bullet.bulletColor = new AscFormat.CBulletColor();
			copyParaPr.Bullet.bulletColor.type = AscFormat.BULLET_TYPE_COLOR_CLR;
			copyParaPr.Bullet.bulletColor.UniColor = AscFormat.CreateUniColorFromRGB(0, 0, 0);
			copyParaPr.Lvl = compiledParaPr.ParaPr.Lvl;
		}
		copyParaPr.Ind = compiledParaPr.ParaPr.Ind.Copy();
		copyParaPr.Ind.FirstLine *= 0.5;
		copyParaPr.Ind.Left *= 0.5;
		copyParaPr.Ind.Right *= 0.5;
		copyParaPr.Spacing.Before = 1;
		copyParaPr.Spacing.After = 1;
		copyParaPr.Spacing.Line = 1;
		return copyParaPr;
	};
	OutlineView.prototype.applyParagraphProps = function (outlineParagraph, slideParagraph, isTitle) {
		const compiledPr = slideParagraph ? slideParagraph.getCompiledPr() : {ParaPr: g_oDocumentDefaultParaPr};
		const copyParaPr = this.getParaPr(compiledPr);
		const oThis = this;
		outlineParagraph.SetPr(copyParaPr);
		outlineParagraph.CheckRunContent(function (run) {
			const textPr = getOutlineTextPr(isTitle, run.IsMathRun());
			run.SetPr(textPr);
		});
		const parTextPr = getOutlineTextPr();
		outlineParagraph.TextPr.Value = parTextPr;
		outlineParagraph.TextPr.CalcValue = parTextPr;
	};
	OutlineView.prototype.createOutlineShape = function (outlineSlides, width, height) {
		this.reset();
		return AscFormat.ExecuteNoHistory(function () {
			const outlineShape = new AscFormat.CShape();
			this.setOutlineShape(outlineShape);
			outlineShape.setBDeleted(false);
			outlineShape.createTextBody();
			outlineShape.txBody.bodyPr.setInsets(0, 0, 0, 0);
			outlineShape.themeOverride = AscFormat.GenerateDefaultTheme();
			const outlineContent = outlineShape.txBody.content;
			outlineContent.ClearContent(false);
			this.fillOutlineShape(outlineShape, outlineSlides);
			this.setSizes(width, height);
			return outlineShape;
		}, this, []);
	};
	OutlineView.prototype.setSizes = function (width, height) {
		if (this.outlineShape) {
			this.outlineShape.extX = width;
			this.outlineShape.extY = typeof height === "number" ? height : 2000;
			this.outlineShape.recalculateContent();
			this.outlineShape.contentWidth = width;
		}
	}
	OutlineView.prototype.getCopyParagraph = function (parent, paragraph, isTitle) {
		return AscFormat.ExecuteNoHistory(function () {
			const copyParagraph = paragraph.Copy2(parent, null, null);
			this.applyParagraphProps(copyParagraph, paragraph, isTitle);
			return copyParagraph;
		}, this, []);
	};
	OutlineView.prototype.draw = function (graphics) {
		if (this.outlineShape) {
			if (this.isEmptyPresentation()) {
				this.drawEmpty(graphics);
			} else {
				this.outlineShape.draw(graphics);
			}
		}
	};
	OutlineView.prototype.drawEmpty = function (graphics) {
		const shape = this.decorCacheManager.getEmptySlideShape(this.outlineShape.extX + this.outlineShape.x, this.outlineShape.extY);
		shape.draw(graphics);
	};

	OutlineView.prototype.getOutlineParagraphs = function () {
		const content = this.getDocContent();
		if (content) {
			return content.Content;
		}
		return [];
	};
	OutlineView.prototype.isHaveParagraphs = function () {
		return this.getOutlineParagraphs();
	};
	OutlineView.prototype.getParagraphY = function (paragraph) {
		for (let i = 0; i < paragraph.Pages.length; i += 1) {
			const page = paragraph.Pages[i];
			if (page.EndLine >= 0) {
				return this.getTransformY(page.Y);
			}
		}

		return this.getTransformY(paragraph.Y);
	};
	OutlineView.prototype.getTransformY = function (y) {
		if (this.outlineShape) {
			return this.outlineShape.transformText.TransformPointY(0, y);
		}
		return null;
	};
	OutlineView.prototype.getTransformX = function (x) {
		if (this.outlineShape) {
			return this.outlineShape.transformText.TransformPointX(x, 0);
		}
		return null;
	};
	OutlineView.prototype.getInvertTransformY = function (y) {
		if (this.outlineShape) {
			return this.outlineShape.invertTransformText.TransformPointY(0, y);
		}
		return null;
	};
	OutlineView.prototype.getInvertTransformX = function (x) {
		if (this.outlineShape) {
			return this.outlineShape.invertTransformText.TransformPointX(x, 0);
		}
		return null;
	};
	OutlineView.prototype.drawDecorations = function (graphics, selectedSlideIndex, focusSlideIndex) {
		if (!this.outlineShape || !this.isHaveParagraphs()) return;

		const decoratorDrawer = new DecoratorDrawer(this, new AscCommon.CMatrix(), selectedSlideIndex, focusSlideIndex);

		const paragraphMap = this.outlineInfo.getOutlineParagraphToInfoMap();
		for (let outlineId in paragraphMap) {
			const info = paragraphMap[outlineId];
			decoratorDrawer.drawTitleDecorations(graphics, info);
			const contentShapeInfoMap = info.getContentShapeInfoMap(this);
			for (let contentShapeInfoId in contentShapeInfoMap) {
				const contentShapeInfo = contentShapeInfoMap[contentShapeInfoId];
				decoratorDrawer.drawContentDecorations(graphics, contentShapeInfo);
			}
		}
	};
	OutlineView.prototype.getDocContent = function () {
		return this.outlineShape && this.outlineShape.txBody.content;
	};

	function DecorationSelectManager(outlineView, e, x, y) {
		this.outlineView = outlineView;
		this.event = e;
		this.x = x;
		this.y = y;
	}
	DecorationSelectManager.prototype.handle = function () {
		const outlineView = this.outlineView;
		const docContent = outlineView.getDocContent();
		if (docContent) {
			const tx = outlineView.getInvertTransformX(this.x);
			const ty = outlineView.getInvertTransformY(this.y);

			const nearestPos = docContent.Get_NearestPos(0, tx, ty);
			const paragraph = nearestPos.Paragraph;
			const isClickToSlideIcon = this.isClickToSlideIcon(paragraph, tx, ty);
			if (isClickToSlideIcon) {

			} else {

			}
		}

	};

	DecorationSelectManager.prototype.isExpandSelection = function () {
		return this.event.CtrlKey;
	};
	OutlineView.prototype.handleDecorationSelect = function () {

	};
	OutlineView.prototype.selectionSetStart = function (e, x, y) {
		if (this.outlineShape) {
			if (this.outlineShape.hit(x, y)) {
				this.outlineShape.selectionSetStart(e, x, y, 0);
			} else {
				this.handleDecorationSelect(e, x, y);
			}
		}
	};
	OutlineView.prototype.updateOutlineShapeTransform = function (offsetX, offsetY) {
		if (!this.outlineShape) {
			return;
		}
		this.outlineShape.x = offsetX;
		this.outlineShape.y = offsetY;
		this.outlineShape.transform.tx = offsetX || 0;
		this.outlineShape.transform.ty = offsetY || 0;
		this.outlineShape.transformText.tx = offsetX || 0;
		this.outlineShape.transformText.ty = offsetY || 0;
		this.outlineShape.invertTransform = AscCommon.global_MatrixTransformer.Invert(this.outlineShape.transform);
		this.outlineShape.invertTransformText = AscCommon.global_MatrixTransformer.Invert(this.outlineShape.transformText);
		this.outlineShape.transformText2 = this.outlineShape.transformText;
		this.outlineShape.invertTransformText2 = this.outlineShape.invertTransformText;
	};
	OutlineView.prototype.getSelectedSlide = function () {
		const docContent = this.getDocContent();
		if (!docContent) {
			return 0;
		}
		const outlineParagraphs = this.getOutlineParagraphs();
		const currentParagraph = docContent.GetCurrentParagraph();
		if (currentParagraph) {
			for (let i = currentParagraph.Index; i >= 0; i -= 1) {
				const paragraph = outlineParagraphs[i];
				const info = this.outlineInfo.getSlideInfoByParagraph(paragraph);
				if (info && info.slide !== undefined) {
					return info.slide.num;
				}
			}
		}
		return null;
	};
	OutlineView.prototype.getTransformText = function () {
		if (this.outlineShape) {
			return this.outlineShape.transformText;
		}
		return new AscCommon.CMatrix();
	};
	OutlineView.prototype.getOutlineHeight = function () {
		if (this.outlineShape) {
			const top = this.outlineShape.transformText.TransformPointY(0, 0);
			const bottom = this.outlineShape.transformText.TransformPointY(0, this.outlineShape.contentHeight);
			return bottom - top;
		}
		return 0;
	}
	OutlineView.prototype.getContentPos = function (callback) {
		let res;
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		if (content.IsSelectionUse()) {
			const startPos = content.GetContentPosition(true, true);
			const endPos = content.GetContentPosition(true, false);
			res = callback(true, startPos, endPos);
		} else {
			const contentPos = content.GetContentPosition(false, false);
			res = callback(false, contentPos);
		}
		return res;
	};
	OutlineView.prototype.forEachSelectedContent = function (callback) {
		const manager = new ForEachSelectManager(this);
		return manager.forEachSelectedContent(callback);
	}

	OutlineView.prototype.getSelectedText = function (oPr) {
		const docContent = this.getDocContent();
		if (!docContent) {
			return "";
		}
		return docContent.GetSelectedText(false, oPr);
	};

	OutlineView.prototype.getSelectionContentsInfo = function () {
		const manager = new ForEachSelectManager(this);
		const content = this.getDocContent();
		if (!content) {
			return [];
		}
		if (content.IsSelectionUse()) {
			return manager.getSelectionUseContentsInfo();
		}
		const contentPos = content.GetContentPosition(false, false);
		const outlineParagraph = content.Content[content.CurPos.ContentPos];
		const sourceParagraph = manager.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			return [{
				startPos: contentPos,
				startParagraph: sourceParagraph,
				endPos: contentPos,
				endParagraph: sourceParagraph,
				content: sourceParagraph.GetParent(),
				direction: AscWord.Direction.FORWARD
			}];
		}
		return [{
			content: null,
			outlineParagraph: outlineParagraph,
			direction: AscWord.Direction.FORWARD
		}];
	};

	OutlineView.prototype.remove = function (bOnAddText) {
		const manager = new RemoveOutlineParagraphsManager(this);
		manager.remove(bOnAddText);
		this.update();
		this.applySavedPositionToOutline();
	};
	OutlineView.prototype.removeSlides = function (slideForRemoveMap) {
		const slides = [];
		const slideNumbers = Object.keys(slideForRemoveMap);
		for (let i = 0 ; i < slideNumbers.length; i += 1) {
			const num = slideNumbers[i];
			const slide = slideForRemoveMap[num];
			if (!slide.isHaveOutlineShapesWithContent()) {
				slides.push(slideForRemoveMap[num]);
			}
		}
		if (!slides.length) {
			return;
		}
		const presentation = this.getPresentation();
		let minSlideIndex = presentation.GetSlidesCount() - 1;
		for (let i = 0; i < slides.length; i++) {
			const index = presentation.GetSlideIndex(slides[i]);
			if (minSlideIndex > index) {
				minSlideIndex = index;
			}
		}

		const unpreserveInfo = presentation.getUnpreserveLayoutsAndMasters(slides);
		const checkArray = [].concat(unpreserveInfo.layouts, unpreserveInfo.masters, slides);
		for (let i = 0; i < checkArray.length; i += 1) {
			const slide = checkArray[i];
			presentation.removeSlideByObject(slide);
		}
		presentation.updateSlideIndexes();
		presentation.DrawingDocument.UpdateThumbnailsAttack();
		this.update();
		presentation.DrawingDocument.m_oWordControl.GoToPage(Math.min(presentation.GetSlidesCount() - 1, minSlideIndex), undefined, undefined, true);
		presentation.Api.sync_HideComment();
	};
	OutlineView.prototype.rebuildSavedPositionPos = function (pos) {
		const docContent = this.getDocContent();
		if (docContent) {
			const sourceParagraph = pos[1].Class;
			const outlineParagraph = this.sourceToOutlineMap[sourceParagraph.GetId()];
			if (outlineParagraph) {
				return rebuildPos(pos, docContent, outlineParagraph);
			}
		}
		return null;
	};
	OutlineView.prototype.rebuildSavedSlidePos = function (slide, isStart) {
		if (!slide) {
			return null;
		}

		const info = this.outlineInfo.getSlideInfoBySlide(slide);
		const docContent = this.getDocContent();
		if (info && docContent) {
			const outlineParagraph = info.outlineParagraph;
			const contentPos = [{Class: docContent, Position: outlineParagraph.Index}];
			if (isStart) {
				outlineParagraph.GetStartContentPosition(contentPos);
			} else {
				outlineParagraph.GetEndContentPosition(contentPos, true);
			}
			return contentPos;
		}
		return null;
	};
	OutlineView.prototype.applySavedPositionToOutline = function () {
		const savedPosition = this.getSavedPosition();
		savedPosition.apply(this);
		this.resetPosition();
	}
	OutlineView.prototype.setSavedPosition = function (position) {
		this.savedPosition = position;
	};
	OutlineView.prototype.getSavedPosition = function () {
		return this.savedPosition;
	};
	OutlineView.prototype.resetPosition = function () {
		this.setSavedPosition(new SavedPosition());
	};
	OutlineView.prototype.resetMapToCheckParagraphs = function () {
		this.mapToCheckParagraphs = {};
	};
	OutlineView.prototype.resetMapToCheckSlides = function () {
		this.mapToCheckSlides = {};
	};
	OutlineView.prototype.paragraphAdd = function (paraItem) {
		const manager = new ParagraphAddManager(this);
		manager.paragraphAdd(paraItem);
		this.update();
		this.applySavedPositionToOutline();
	}
	OutlineView.prototype.insertContent = function (selectedContent) {
		const manager = new OutlineContentInsertManager(this);
		const inserted = manager.insert(selectedContent);
		this.update();
		this.applySavedPositionToOutline();
		return inserted;
	}

	OutlineView.prototype.getParagraphParaPr = function () {
		let paraPr;
		this.forEachSelectedContent(function (selectProps) {
			if (selectProps.content) {
				const content = selectProps.content;
				const contentParaPr = content.GetCalculatedParaPr();
				if (paraPr) {
					paraPr.Compare(contentParaPr);
				} else {
					paraPr = contentParaPr;
				}
			}

		});
		return paraPr;
	};

	OutlineView.prototype.getParagraphTextPr = function () {
		let textPr;
		this.forEachSelectedContent(function (selectProps) {
			if (selectProps.content) {
				const content = selectProps.content;
				const contentTextPr = content.GetCalculatedTextPr();
				if (textPr) {
					textPr.Compare(contentTextPr);
				} else {
					textPr = contentTextPr;
				}
			}
		});
		return textPr;
	};
	OutlineView.prototype.unlinkOutlineParagraph = function (outlineParagraph) {
		const index = outlineParagraph.Index;
		this.removeParagraph(outlineParagraph, index);
	};
	OutlineView.prototype.getPropertiesFromSourceShape = function (sourceParagraph) {
		if (sourceParagraph.Index === 0) {
			const parentShape = sourceParagraph.GetParentShape();
			if (parentShape.isGeneralTitle()) {
				return new OutlineSlideInfo(parentShape.parent);
			}
		}
		return null;
	};
	OutlineView.prototype.updateFromSourceParagraph = function (sourceParagraph) {
		const outlineParagraph = this.sourceToOutlineMap[sourceParagraph.Get_Id()];
		const index = outlineParagraph.Index;
		const info = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		this.removeParagraph(outlineParagraph, index);
		return this.addUpdatedParagraph(sourceParagraph, index, info);
	};
	OutlineView.prototype.addUpdatedParagraph = function (sourceParagraph, index, pr) {
		pr = pr || this.getPropertiesFromSourceShape(sourceParagraph);
		return this.addCopyParagraph(sourceParagraph, index, sourceParagraph.Index === 0, pr);
	};

	OutlineView.prototype.addUpdatedParagraphWithCheckMockParagraph = function (sourceParagraph, index) {
		const docContent = this.getDocContent();
		const outlineParagraph = docContent.Content[index];
		if (outlineParagraph) {
			const pr = this.getPropertiesFromSourceShape(sourceParagraph);
			const info = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
			if (info && pr && pr.slide === info.slide) {
				this.removeParagraph(outlineParagraph, index);
				this.addUpdatedParagraph(sourceParagraph, index, pr);
				return true;
			}
		}
		return false;
	};

	function ParagraphUpdateData() {
		this.existingParagraphs = [];
		this.newParagraphs = [];
	}
	ParagraphUpdateData.prototype.isNeedRecalculate = function () {
		return !!this.existingParagraphs.length || !!this.newParagraphs.length;
	};

	function SlideUpdateData() {
		this.removedSlides = [];
		this.addedSlides = [];
	}
	SlideUpdateData.prototype.isNeedRecalculate = function () {
		return !!this.removedSlides.length || !!this.addedSlides.length;
	};
	OutlineView.prototype.getParagraphUpdateData = function () {
		const updateData = new ParagraphUpdateData();
		for (let id in this.mapToCheckParagraphs) {
			const paragraph = this.mapToCheckParagraphs[id];
			const shape = paragraph.GetParentShape();
			if (!shape || shape.getObjectType() !== AscDFH.historyitem_type_Shape) {
				continue;
			}
			if (shape === this.outlineShape) {
				const sourceParagraph = this.getSourceParagraph(paragraph);
				if (!sourceParagraph) {
					updateData.existingParagraphs.push(paragraph)
				}
			} else if (shape.isOutlinePlaceholder()) {
				const content = shape.getDocContent();
				if (shape.parent && shape.parent.getObjectType() === AscDFH.historyitem_type_Slide) {
					const outlineParagraph = this.sourceToOutlineMap[paragraph.Get_Id()];
					if (outlineParagraph) {
						updateData.existingParagraphs.push(outlineParagraph);
					} else if (paragraph.IsUseInDocument()) {
						const isEmptyContent = content.IsEmpty();
						if (isEmptyContent && shape.isGeneralTitle() || !isEmptyContent) {
							updateData.newParagraphs.push(paragraph);
						}
					}
				}
			}
		}
		this.resetMapToCheckParagraphs();
		return updateData;
	};
	OutlineView.prototype.getSlideUpdateData = function () {
		const updateData = new SlideUpdateData();
		if (AscCommon.isEmptyObject(this.mapToCheckSlides)) {
			return updateData;
		}

		const presentation = this.getPresentation();
		const usedSlidesMap = presentation.GetUsedSlidesMap();
		for (let id in this.mapToCheckSlides) {
			const slide = this.mapToCheckSlides[id];
			if (!slide.isHaveTitleShape()) {
				const slideInfo = this.outlineInfo.getSlideInfoBySlide(slide);
				if (usedSlidesMap[slide.GetId()]) {
					if (!slideInfo) {
						updateData.addedSlides.push(slide);
					}
				} else {
					if (slideInfo) {
						updateData.removedSlides.push(slide);
					}
				}
			}
		}
		this.resetMapToCheckSlides();
		return updateData;
	};
	OutlineView.prototype.updateParagraphData = function () {
		const updateData = this.getParagraphUpdateData();
		if (updateData.isNeedRecalculate()) {
			const presentation = this.getPresentation();
			presentation.updateSlideIndexes();
			this.updateExistingParagraphs(updateData.existingParagraphs);
			this.updateNewParagraphs(updateData.newParagraphs);
			return true;
		}
		return false;
	};
	OutlineView.prototype.updateSlideData = function () {
		const updateData = this.getSlideUpdateData();
		if (updateData.isNeedRecalculate()) {
			const presentation = this.getPresentation();
			presentation.updateSlideIndexes();
			this.updateAddedSlides(updateData.addedSlides);
			this.updateRemovedSlides(updateData.removedSlides);
			return true;
		}
		return false;
	};
	OutlineView.prototype.update = function () {
		if (!this.isInit) {
			return;
		}
		let isRecalculate = false;
		isRecalculate |= this.updateParagraphData();
		isRecalculate |= this.updateSlideData();
		if (isRecalculate) {
			this.outlineShape && this.outlineShape.recalculateContent();
		}
	};
	function UseInDocumentManager() {
		this.useInDocumentShapes = {};
		this.useInDocumentParagraphs = {};
		this.useInDocumentSlides = null;
	}
	UseInDocumentManager.prototype.getPresentation = function () {
		return Asc.editor.WordControl.m_oLogicDocument;
	};
	UseInDocumentManager.prototype.isForceUseInDocumentParagraph = function (paragraph) {
		if (paragraph.Index === -1) {
			return false;
		}
		const shape = paragraph.GetParentShape();
		const isUseInDocumentShape = this.isUseInDocumentShape(shape);
		if (!isUseInDocumentShape) {
			return false;
		}
		const content = shape.getDocContent();
		if (content.Content[paragraph.GetIndex()] !== paragraph) {
			return false;
		}
		return true;
	};
	UseInDocumentManager.prototype.isUseInDocumentParagraph = function (paragraph) {
		if (!paragraph) {
			return false;
		}
		if (this.useInDocumentParagraphs[paragraph.GetId()] === undefined) {
			this.useInDocumentParagraphs[paragraph.GetId()] = this.isForceUseInDocumentParagraph(paragraph);
		}
		return this.useInDocumentParagraphs[paragraph.GetId()]
	};
	UseInDocumentManager.prototype.isUseInDocumentShape = function (shape) {
		if (!shape) {
			return false;
		}
		if (this.useInDocumentShapes[shape.GetId()] === undefined) {
			const slide = shape.parent;
			const isUseInDocumentSlide = this.isUseInDocumentSlide(slide);
			if (isUseInDocumentSlide) {
				for (let i = 0; i < slide.cSld.spTree.length; i += 1) {
					const sp = slide.cSld.spTree[i];
					this.useInDocumentShapes[sp.GetId()] = true;
				}
			}
		}
		return this.useInDocumentShapes[shape.GetId()];
	};
	UseInDocumentManager.prototype.isUseInDocumentSlide = function (slide) {
		if (!slide) {
			return false;
		}
		if (this.useInDocumentSlides === null) {
			this.useInDocumentSlides = {};
			const presentation = this.getPresentation();
			for (let i = 0; i < presentation.Slides.length; i += 1) {
				const slide = presentation.Slides[i];
				this.useInDocumentSlides[slide.GetId()] = true;
			}
		}
		return this.useInDocumentSlides[slide.GetId()];
	};
	function UpdateExistingParagraphManager(outlineView, useInDocumentManager) {
		this.outlineView = outlineView;
		this.useInDocumentManager = useInDocumentManager;
	}
	UpdateExistingParagraphManager.prototype.isNeedAddMockParagraph = function (sourceParagraph) {
		if (!sourceParagraph || sourceParagraph.Index !== 0) {
			return false;
		}
		const shape = sourceParagraph.GetParentShape();
		if (!shape.isGeneralTitle() || this.useInDocumentManager.isUseInDocumentShape(shape)) {
			return false;
		}

		const slide = shape.parent;
		if (!this.useInDocumentManager.isUseInDocumentSlide(slide)) {
			return false;
		}

		return true;
	};
	UpdateExistingParagraphManager.prototype.update = function (existingOutlineParagraphs) {
		existingOutlineParagraphs.sort(function (aParagraph, bParagraph) {
			return bParagraph.Index - aParagraph.Index;
		});
		const outlineView = this.outlineView;
		for (let i = 0; i < existingOutlineParagraphs.length; i += 1) {
			const outlineParagraph = existingOutlineParagraphs[i];
			const sourceParagraph = outlineView.getSourceParagraph(outlineParagraph);
			let content;
			let shape;
			let isEmptyContent = false;
			if (sourceParagraph) {
				content = sourceParagraph.GetParent();
				isEmptyContent = content.IsEmpty();
				shape = content.Is_DrawingShape(true);
			}

			if (this.useInDocumentManager.isUseInDocumentParagraph(sourceParagraph) && (!isEmptyContent || isEmptyContent && shape.isGeneralTitle())) {
				outlineView.updateFromSourceParagraph(sourceParagraph);
			} else {
				this.unlinkOutlineParagraph(outlineParagraph, sourceParagraph);
			}
		}
	}
	UpdateExistingParagraphManager.prototype.unlinkOutlineParagraph = function (outlineParagraph, sourceParagraph) {
		const outlineView = this.outlineView;
		const info = outlineView.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		outlineView.unlinkOutlineParagraph(outlineParagraph);
		if (this.isNeedAddMockParagraph(sourceParagraph)) {
			outlineView.addMockTitle(outlineParagraph.Index, info);
		} else if (sourceParagraph) {
			this.setRecalculateContentShapeIndexes(sourceParagraph);
		}
	}
	UpdateExistingParagraphManager.prototype.setRecalculateContentShapeIndexes = function (sourceParagraph) {
		const outlineView = this.outlineView;
		const shape = sourceParagraph.GetParentShape();
		const content = sourceParagraph.GetParent();
		const slide = shape.parent;
		if ((!this.useInDocumentManager.isUseInDocumentShape(shape) || content.IsEmpty()) && this.useInDocumentManager.isUseInDocumentSlide(slide)) {
			const info = outlineView.outlineInfo.getSlideInfoBySlide(slide);
			if (info) {
				info.resetContentShapeInfoMap();
			}
		}
	}
	OutlineView.prototype.updateExistingParagraphs = function (existingOutlineParagraphs) {
		if (!existingOutlineParagraphs.length) {
			return;
		}
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		content.Update_ContentIndexing();
		const manager = new UpdateExistingParagraphManager(this, new UseInDocumentManager());
		manager.update(existingOutlineParagraphs);
	}

	OutlineView.prototype.updateNewParagraphs = function (newParagraphs) {
		if (!newParagraphs.length) {
			return;
		}
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		content.Update_ContentIndexing();
		const updateManager = new UpdateNewParagraphsManager(this, newParagraphs);
		updateManager.update();
	};

	OutlineView.prototype.updateAddedSlides = function (addedSlides) {
		if (!addedSlides.length) {
			return;
		}
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		content.Update_ContentIndexing();
		const updateManager = new UpdateAddedSlidesManager(this, addedSlides);
		updateManager.update();
	};
	OutlineView.prototype.updateRemovedSlides = function (removedSlides) {
		if (!removedSlides.length) {
			return;
		}
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		content.Update_ContentIndexing();
		const updateManager = new UpdateRemovedSlidesManager(this, removedSlides);
		updateManager.update();
	};
	OutlineView.prototype.checkParagraph = function (paragraph) {
		if (AscCommon.History.IsOn()) {
			this.mapToCheckParagraphs[paragraph.GetId()] = paragraph;
		}
	};
	OutlineView.prototype.checkSourceSlide = function (slide) {
		if (!AscCommon.History.IsOn()) {
			return;
		}
		this.mapToCheckSlides[slide.GetId()] = slide;
		for (let i = 0; i < slide.cSld.spTree.length; i += 1) {
			const shape = slide.cSld.spTree[i];
			this.checkSourceShape(shape);
		}
	};
	OutlineView.prototype.checkSourceShape = function (shape) {
		if (!AscCommon.History.IsOn()) {
			return;
		}
		if (shape.isOutlinePlaceholder()) {
			const content = shape.getDocContent();
			for (let i = 0; i < content.Content.length; i += 1) {
				const paragraph = content.Content[i];
				this.checkParagraph(paragraph);
			}
		}
	};
	OutlineView.prototype.getApi = function () {
		return this.api;
	};
	OutlineView.prototype.selectAll = function () {
		const content = this.getDocContent();
		if (content) {
			content.SelectAll();
			const drawingDocument = this.outlineShape.getDrawingDocument();
			this.outlineShape.updateSelectionState(drawingDocument);
		}
	};
	OutlineView.prototype.executeShortcut = function (shortcutAction) {
		let res = {keyResult: keydownresult_PreventAll};
		switch (shortcutAction) {
			case Asc.c_oAscPresentationShortcutType.EditSelectAll: {
				this.selectAll();
				break;
			}
			case Asc.c_oAscPresentationShortcutType.EditUndo:
			case Asc.c_oAscPresentationShortcutType.EditRedo: {
				res.keyResult = keydownresult_PreventNothing;
				break;
			}
			default: {
				res = null;
			}
		}
		return res;
	}
	OutlineView.prototype.updateInterfaceState = function () {

	};
	OutlineView.prototype.updateSelectionState = function () {
		if (this.outlineShape) {
			if (this.isEmptyPresentation()) {
				this.outlineShape.resetSelectionState();
			} else {
				this.outlineShape.updateSelectionState();
			}
		}
	};
	OutlineView.prototype.moveCursorToStartPos = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToStartPos(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorToEndPos = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToEndPos(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorLeft = function (AddToSelect, Word) {
		const content = this.getDocContent();

		if (this.isRtl())
			content.MoveCursorRight(AddToSelect, Word);
		else
			content.MoveCursorLeft(AddToSelect, Word);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorRight = function (AddToSelect, Word) {
		const content = this.getDocContent();
		if (this.isRtl())
			content.MoveCursorLeft(AddToSelect, Word);
		else
			content.MoveCursorRight(AddToSelect, Word);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorUp = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorUp(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorDown = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorDown(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorToEndOfLine = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToEndOfLine(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};

	OutlineView.prototype.moveCursorToStartOfLine = function (AddToSelect) {
		const content = this.getDocContent();
		content.MoveCursorToStartOfLine(AddToSelect);
		this.updateSelectionState();
		this.updateInterfaceState();
	};
	OutlineView.prototype.isRtl = function () {
		let oContent = this.getDocContent();
		if (oContent) {
			let curPara = oContent.GetCurrentParagraph();
			return !!(curPara && curPara.isRtlDirection());
		}
		return false;
	}
	OutlineView.prototype.onKeyDown = function (e) {
		const api = this.getApi();
		const shortcutAction = api.getShortcut(e);
		const shortcutRes = this.executeShortcut(shortcutAction);
		if (shortcutRes) {
			return shortcutRes.keyResult;
		}
		return this.executeHotKey(e);
	};
	OutlineView.prototype.executeHotKey = function (e) {
		const isMacOs = AscCommon.AscBrowser.isMacOs;
		const presentation = this.getPresentation();
		switch (e.KeyCode) {
			case Asc.c_oAscKeyCodes.ArrowLeft: {
				if (isMacOs && e.CtrlKey) {
					this.moveCursorToStartOfLine(e.ShiftKey);
				} else {
					const bIsWord = isMacOs ? e.AltKey : e.CtrlKey;
					this.moveCursorLeft(e.ShiftKey, bIsWord);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowRight: {
				if (isMacOs && e.CtrlKey) {
					this.moveCursorToEndOfLine(e.ShiftKey);
				} else {
					const bIsWord = isMacOs ? e.AltKey : e.CtrlKey;
					this.moveCursorRight(e.ShiftKey, bIsWord);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowUp: {
				this.moveCursorUp(e.ShiftKey);
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowDown: {
				this.moveCursorDown(e.ShiftKey);
				break;
			}
			case Asc.c_oAscKeyCodes.Home: {
				if (e.CtrlKey) {
					this.moveCursorToStartPos(e.ShiftKey);
				} else {
					this.moveCursorToStartOfLine(e.ShiftKey);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.End: {
				if (e.CtrlKey) {
					this.moveCursorToEndPos(e.ShiftKey);
				} else {
					this.moveCursorToEndOfLine(e.ShiftKey);
				}
				break;
			}
			case Asc.c_oAscKeyCodes.Backspace: {
				presentation.StartAction();
				this.remove();
				presentation.FinalizeAction();
				this.updateInterfaceState()
				this.updateSelectionState();
				break;
			}
			case Asc.c_oAscKeyCodes.Enter: {
				presentation.StartAction();
				this.handleEnter(e);
				presentation.FinalizeAction();
				this.updateInterfaceState()
				this.updateSelectionState();
				break;
			}
		}
	};
	OutlineView.prototype.getDrawingDocument = function () {
		const presentation = this.getPresentation();
		return presentation.GetDrawingDocument();
	};
	OutlineView.prototype.isFocusOnOutline = function () {
		const presentation = this.getPresentation();
		return presentation.IsFocusOnOutline();
	};
	OutlineView.prototype.getSelectedSlidesRange = function () {
		const selectedArray = this.getSelectedSlideArray();
		if (selectedArray.length) {
			return {Min: selectedArray[0], Max: selectedArray[selectedArray.length - 1]};
		}
		return {Min: -1, Max: -1};
	};
	OutlineView.prototype.getSlideIndex = function (outlineParagraph) {
		const info = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			const parentShape = sourceParagraph.GetParentShape();
			if (parentShape && parentShape.parent) {
				return parentShape.parent.num;
			}
		} else if (info) {
			return info.slide.num;
		}
		return -1;
	};
	OutlineView.prototype.getSelectedSlideArrayFromSelection = function (startPos, endPos) {
		const startParagraph = startPos[1].Class;
		const endParagraph = endPos[1].Class;
		let startSlideIndex = this.getSlideIndex(startParagraph);
		let endSlideIndex = this.getSlideIndex(endParagraph);
		if (startSlideIndex !== -1 && endSlideIndex !== -1) {
			if (startSlideIndex > endSlideIndex) {
				const temp = startSlideIndex;
				startSlideIndex = endSlideIndex;
				endSlideIndex = temp;
			}
			const selectedSlides = [];
			for (let i = startSlideIndex; i <= endSlideIndex; i += 1) {
				selectedSlides.push(i);
			}
			return selectedSlides;
		} else if (startSlideIndex !== -1) {
			return [startSlideIndex];
		} else if (endSlideIndex !== -1) {
			return [endSlideIndex];
		}
		return [];
	};
	OutlineView.prototype.getSelectedSlideArrayFromCurPos = function (curPos) {
		const startParagraph = curPos[1].Class;
		const startSlideIndex = this.getSlideIndex(startParagraph);
		if (startSlideIndex !== -1) {
			return  [startSlideIndex];
		}
		return [];
	};

	OutlineView.prototype.getSelectedSlideArray = function () {
		if (this.isEmptyPresentation()) {
			return [];
		}
		if (this.isFocusOnOutline()) {
			const oThis = this;
			const selectedArray = this.getContentPos(function (isSelectionUse, startPos, endPos) {
				if (isSelectionUse) {
					return oThis.getSelectedSlideArrayFromSelection(startPos, endPos);
				} else {
					return oThis.getSelectedSlideArrayFromCurPos(startPos);
				}
			});
			return selectedArray || [];
		}
		const drawingDocument = this.getDrawingDocument();
		if (drawingDocument.SlideCurrent >= 0) {
			return [drawingDocument.SlideCurrent];
		}

		return [];
	};
	OutlineView.prototype.isSelectedPage = function (pageNum) {
		const selectedPages = this.getSelectedSlideArray();
		return selectedPages.indexOf(pageNum) !== -1;
	};
	OutlineView.prototype.getTitleSlideParagraph = function (pageNum) {
		const docContent = this.getDocContent();
		if (docContent) {
			for (let i = 0; i < docContent.Content.length; i += 1) {
				const paragraph = docContent.Content[i];
				const info = this.outlineInfo.getSlideInfoByParagraph(paragraph);
				if (info && info.slide.num === pageNum) {
					return paragraph;
				}
			}
		}
	};
	OutlineView.prototype.getPageY = function (pageNum) {
		const docContent = this.getDocContent();
		if (docContent) {
			const paragraph = this.getTitleSlideParagraph(pageNum);
			if (paragraph) {
				return this.getParagraphY(paragraph);
			}

		}
		return 0;
	};
	OutlineView.prototype.selectPage = function (pageNum) {
		const docContent = this.getDocContent();
		if (docContent) {
			docContent.RemoveSelection();
			const paragraph = this.getTitleSlideParagraph(pageNum);
			if (paragraph) {
				paragraph.MoveCursorToStartPos();
				paragraph.SetThisElementCurrent();
			}
			this.updateSelectionState();
		}
	};
	OutlineView.prototype.getThumbnailHeight = function () {
		return this.decorCacheManager.getMeasureInfo().height;
	};
	OutlineView.prototype.removeSelection = function () {
		const docContent = this.getDocContent();
		if (docContent) {
			docContent.RemoveSelection();
			this.outlineShape.resetSelectionState();
		}
	};

	OutlineView.prototype.getNearestPage = function (x, y) {
		if (this.isEmptyPresentation()) {
			return -1;
		}
		const docContent = this.getDocContent();
		if (docContent) {
			const tx = this.getInvertTransformX(x);
			const ty = this.getInvertTransformY(y);
			const nearestPos = docContent.Get_NearestPos(0, tx, ty);
			return this.getSlideIndex(nearestPos.Paragraph);
		}
		return -1;
	};
	OutlineView.prototype.hitInTextRect = function (x, y) {
		if (this.outlineShape && !this.isEmptyPresentation()) {
			return AscFormat.HitToRect(x, y, this.outlineShape.invertTransformText, 0, 0, this.outlineShape.contentWidth, 20000);
		}
		return false;
	};
	OutlineView.prototype.checkOutlineWidth = function (width) {
		if (!this.outlineShape) {
			return;
		}
		if (!AscFormat.fApproxEqual(width, this.outlineShape.extX)) {
			this.setSizes(width);
		}
	};
	OutlineView.prototype.isTitleOutlineParagraph = function (paragraph) {
		return !!this.outlineInfo.getSlideInfoByParagraph(paragraph);
	};
	OutlineView.prototype.getSourceParagraph = function (outlineParagraph) {
		return this.outlineToSourceMap[outlineParagraph.Get_Id()];
	};
	OutlineView.prototype.getOutlineSlideParagraphs = function (slide) {
		const info = this.outlineInfo.getSlideInfoBySlide(slide);
		if (!info) {
			return [];
		}
		const content = this.getDocContent();
		if (!content) {
			return [];
		}
		const outlineParagraph = info.outlineParagraph;
		const startIndex = outlineParagraph.Index + 1;
		const result = [outlineParagraph];
		for (let i = startIndex; i < content.Content.length; i += 1) {
			const currentOutlineParagraph = content.Content[i];
			if (this.isTitleOutlineParagraph(currentOutlineParagraph)) {
				break;
			}
			result.push(currentOutlineParagraph);
		}
		return result;
	};
	OutlineView.prototype.getThumbnailPagePosition = function (pageIndex) {
		const presentation = this.getPresentation();
		const paragraphs = this.getOutlineSlideParagraphs(presentation.Slides[pageIndex]);
		if (paragraphs.length) {
			const firstParagraph = paragraphs[0];
			const lastParagraph = paragraphs[paragraphs.length - 1];
			const lastLine = lastParagraph.Lines[lastParagraph.Lines.length - 1];
			if (!lastLine) {
				return null;
			}
			const Y = this.getTransformY(firstParagraph.Y);
			const H = Y - (this.getTransformY(lastParagraph.Y) + lastLine.Bottom);
			return {
				X: this.getTransformX(0),
				Y: Y,
				W: this.outlineShape.contentWidth,
				H: H
			};
		}
		return null;
	};
	OutlineView.prototype.isEmptyPresentation = function () {
		const presentation = this.getPresentation();
		return !presentation.Slides.length;
	};
	OutlineView.prototype.drawSelectionOnPage = function (drawingDocument) {
		const content = this.getDocContent();
		if (content) {
			drawingDocument.UpdateTargetTransform(this.outlineShape.transformText);
			content.DrawSelectionOnPage(0);
		}
	};
	OutlineView.prototype.recalculateCurPos = function (bUpdateX, bUpdateY) {
		const content = this.getDocContent();
		if (content) {
			content.RecalculateCurPos(bUpdateX, bUpdateY);
		}
	};
	OutlineView.prototype.setSavedRealContentState = function (content, idx, count, direction) {
		direction = typeof direction === 'number' ? direction : AscWord.Direction.FORWARD;
		if (count === 1) {
			if (content.IsSelectionUse()) {
				const startPos = content.GetContentPosition(true, true);
				const endPos = content.GetContentPosition(true, false);
				this.setSavedPosition(new SavedPosition(new ContentPosition(startPos), new ContentPosition(endPos)));
			} else {
				this.setSavedPosition(new SavedPosition(new ContentPosition(content.GetContentPosition(false, false))));
			}
		} else if (idx === 0) {
			const savedPosition = this.getSavedPosition();
			if (content.IsSelectionUse()) {
				if (direction === AscWord.Direction.BACKWARD) {
					savedPosition.endPos = new ContentPosition(content.GetContentPosition(true, false));
				} else {
					savedPosition.startPos = new ContentPosition(content.GetContentPosition(true, true));
				}
			} else if (!savedPosition.startPos) {
				savedPosition.startPos = new ContentPosition(content.GetContentPosition(false, false));
			}
		} else if (idx === count - 1) {
			const savedPosition = this.getSavedPosition();
			if (content.IsSelectionUse()) {
				if (direction === AscWord.Direction.BACKWARD) {
					savedPosition.startPos = new ContentPosition(content.GetContentPosition(true, true));
				} else {
					savedPosition.endPos = new ContentPosition(content.GetContentPosition(true, false));
				}
			} else if (!savedPosition.startPos) {
				savedPosition.startPos = new ContentPosition(content.GetContentPosition(false, false));
			}
		}
	};
	OutlineView.prototype.setSavedMockedParagraphState = function (outlineParagraph, idx, count, direction) {
		direction = typeof direction === 'number' ? direction : AscWord.Direction.FORWARD;

		const info = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		if (!info) {
			return;
		}
		if (count === 1) {
			this.setSavedPosition(new SavedPosition(new SlidePosition(info.slide)));
		} else if (idx === 0) {
			const savedPosition = this.getSavedPosition();
			if (outlineParagraph.IsSelectionUse()) {
				if (direction === AscWord.Direction.BACKWARD) {
					const endPos = outlineParagraph.Get_ParaContentPos(true, false);
					savedPosition.endPos = new SlidePosition(info.slide, outlineParagraph.IsCursorAtBegin(endPos));
				} else {
					const startPos = outlineParagraph.Get_ParaContentPos(true, true);
					savedPosition.startPos = new SlidePosition(info.slide, outlineParagraph.IsCursorAtBegin(startPos));
				}
			} else if (!savedPosition.startPos) {
				savedPosition.startPos = new SlidePosition(info.slide);
			}
		} else if (idx === count - 1) {
			const savedPosition = this.getSavedPosition();
			if (outlineParagraph.IsSelectionUse()) {
				if (direction === AscWord.Direction.BACKWARD) {
					const startPos = outlineParagraph.Get_ParaContentPos(true, true);
					savedPosition.startPos = new SlidePosition(info.slide, outlineParagraph.IsCursorAtBegin(startPos));
				} else {
					const endPos = outlineParagraph.Get_ParaContentPos(true, false);
					savedPosition.endPos = new SlidePosition(info.slide, outlineParagraph.IsCursorAtBegin(endPos));
				}
			} else if (!savedPosition.startPos) {
				savedPosition.startPos = new SlidePosition(info.slide);
			}
		}
	};
	OutlineView.prototype.saveDocumentStateBeforeLoadChanges = function (docState) {
		const presentation = this.getPresentation();
		const docContent = this.getDocContent();
		if (!docContent || !presentation) {
			return;
		}
		const oThis = this;
		this.forEachSelectedContent(function (selectProps) {
			const idx = selectProps.index;
			const contentCount = selectProps.count;
			const direction = selectProps.direction;
			if (selectProps.outlineParagraph) {
				oThis.setSavedMockedParagraphState(selectProps.outlineParagraph, idx, contentCount, direction);
			} else {
				oThis.setSavedRealContentState(selectProps.content, idx, contentCount, direction);
			}
		});
		const savedPosition = this.getSavedPosition();
		docState.SavedOutlinePosition = savedPosition;
		if (savedPosition.startPos && savedPosition.endPos) {
			if (savedPosition.endPos.isDocStatePosition()) {
				docState.EndPos = savedPosition.endPos.pos;
			}
			if (savedPosition.startPos.isDocStatePosition()) {
				docState.StartPos = savedPosition.startPos.pos;
			}
		} else if (savedPosition.startPos && savedPosition.startPos.isDocStatePosition()) {
			docState.Pos = savedPosition.startPos.pos;
		} else if (savedPosition.endPos && savedPosition.endPos.isDocStatePosition()) {
			docState.Pos = savedPosition.endPos.pos;
		}
		docState.DrawingSelection = docContent.Selection.Use;
		docState.Slide = presentation.GetCurrentSlide();
		this.resetPosition();
	};

	OutlineView.prototype.applyDocumentStateAfterLoadChanges = function (docState) {
		this.update();
		const savedPosition = docState.SavedOutlinePosition;
		if (savedPosition) {
			this.setSavedPosition(savedPosition);
			this.applySavedPositionToOutline();
		}
	};
	OutlineView.prototype.handleTitleEnter = function (outlineParagraph) {
		const slideInfo = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		if (!slideInfo) {
			return;
		}
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		const outlineSlide = slideInfo.slide.getOutlineSlide();
		const shapesWithContent = outlineSlide.content.filter(function (element) {
			return element.isOutlinePlaceholderWithContent();
		});

		const presentation = this.getPresentation();
		const newSlide = presentation.addNextSlideAction();
		if (!sourceParagraph && shapesWithContent.length === 0) {
			return newSlide;
		}

		if (sourceParagraph) {
			const newParagraph = sourceParagraph.Split();
			let titleSp = newSlide.getMatchingShape(AscFormat.phType_title, null, false, {});
			if (!titleSp) {
				titleSp = newSlide.createTitle();
			}
			const titleDocContent = titleSp.getDocContent();
			titleDocContent.Clear_Content();
			const newSlideParagraph = titleDocContent.Content[0];
			new SavePosManager(this).saveStartPosSourceParagraph(newSlideParagraph);
			newSlideParagraph.Concat(newParagraph);
		}

		if (shapesWithContent.length !== 0) {
			let contentShape = newSlide.cSld.getContentShape();
			if (!contentShape) {
				contentShape = newSlide.createContent();
			}
			const contentDocContent = contentShape.getDocContent();
			contentDocContent.Clear_Content();

			const firstShapeContent = shapesWithContent[0].getDocContent();
			const concatShapesManager = new ConcatShapesManager(this);
			concatShapesManager.concatContents(contentDocContent, firstShapeContent);
		}
		return newSlide;
	};
	OutlineView.prototype.handleEnter = function (e) {
		this.remove(true);
		const presentation = this.getPresentation();
		const oThis = this;
		let isMathHandle = false;
		let newSlide;
		this.forEachSelectedContent(function (selectProps) {
			if (selectProps.count !== 1) {
				return true;
			}
			const content = selectProps.content;
			if (selectProps.content) {
				let oShape = content.Is_DrawingShape(true);
				if (!oShape) {
					return;
				}
				var oSelectedInfo = new CSelectedElementsInfo();
				content.GetSelectedElementsInfo(oSelectedInfo);
				var oMath = oSelectedInfo.GetMath();
				if (null !== oMath && oMath.Is_InInnerContent()) {
					isMathHandle = true;
					if (oMath.Handle_AddNewLine()) {
						oShape.checkExtentsByDocContent();
						this.Recalculate();
					}
				} else if (!e.ShiftKey) {
					if (oShape.isGeneralTitle()) {
						const sourceParagraph = content.Content[0];
						const outlineParagraph = oThis.sourceToOutlineMap[sourceParagraph.GetId()];
						newSlide = oThis.handleTitleEnter(outlineParagraph);
					} else {

						content.AddNewParagraph();
						oThis.setSavedRealContentState(content, selectProps.idx, selectProps.count, selectProps.direction);
					}
				}
			} else {
				newSlide = oThis.handleTitleEnter(selectProps.outlineParagraph);
			}
		});
		if (e.ShiftKey && !isMathHandle) {
			presentation.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line), false, true);
		}

		this.update();
		this.applySavedPositionToOutline();
		if (newSlide) {
			presentation.DrawingDocument.m_oWordControl.GoToPage(newSlide.num);
		}
	};
	OutlineView.prototype.setSelectionState = function (slideSelectionState) {
		const state = slideSelectionState[0];
		if (state && state.textSelection) {
			this.update();
			const thumbnailsManager = Asc.editor.getThumbnailsManager();
			if (!thumbnailsManager) {
				return;
			}
			thumbnailsManager.SetFocusElement(AscCommon.FOCUS_OBJECT_THUMBNAILS);
			const docContent = this.getDocContent();
			if (docContent) {
				docContent.SetSelectionState(state.textSelection);
			}
		}
	};
	OutlineView.prototype.getSelectionState = function () {
		const selectionState = {};
		const docContent = this.getDocContent();
		if (docContent) {
			selectionState.focus = true;
			selectionState.textObject = this.outlineShape;
			selectionState.selectStartPage = selectionState.textObject.selectStartPage;
			selectionState.textSelection = docContent.GetSelectionState();
		}
		return [selectionState];
	};


	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].OutlineSlide = OutlineSlide;
	window["AscCommonSlide"].OutlineView = OutlineView;
	window["AscCommonSlide"].DecoratorDrawer = DecoratorDrawer;
})();
