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
		textPr.InitDefault();
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

	const CHECK_SLIDE_INFO_ADD = 0x10;
	const CHECK_SLIDE_INFO_REMOVE = 0x01;
	const CHECK_SLIDE_INFO_REPLACE = 0x11;
	const CHECK_SLIDE_INFO_NOTHING = 0x00;
	function CheckSlideInfo(slide) {
		this.slide = slide;
		this.flags = CHECK_SLIDE_INFO_NOTHING;
	}
	CheckSlideInfo.prototype.isReplace = function () {
		return this.flags === CHECK_SLIDE_INFO_REPLACE;
	};

	function SlideUpdateDataCollector(outlineView, checkSlideMap, usedSlideMap) {
		this.outlineView = outlineView;
		this.checkSlideMap = checkSlideMap;
		this.usedSlideMap = usedSlideMap;
	}
	SlideUpdateDataCollector.prototype.collect = function () {
		const updateData = new SlideUpdateData();
		const outlineView = this.outlineView;
		for (let id in this.checkSlideMap) {
			const checkInfo = this.checkSlideMap[id];
			const slide = checkInfo.slide;
			if (!slide.isHaveTitleShape()) {
				const slideInfo = outlineView.outlineInfo.getSlideInfoBySlide(slide);
				if (this.usedSlideMap[slide.GetId()]) {
					if (checkInfo.isReplace()) {
						updateData.removedSlides.push(slide);
						updateData.addedSlides.push(slide);
					} else if (!slideInfo) {
						updateData.addedSlides.push(slide);
					}
				} else {
					if (slideInfo) {
						updateData.removedSlides.push(slide);
					}
				}
			}
		}
		return updateData;
	};
	SlideUpdateDataCollector.prototype.addSlidesToCheck = function (slideMap) {
		this.checkSlideMap = Object.assign({}, this.checkSlideMap, slideMap);
	};
	function ParagraphUpdateDataCollector(outlineView, paragraphMap, additionalCorrectedContents) {
		this.outlineView = outlineView;
		this.paragraphMap = paragraphMap;
		this.additionalCheckSlideMap = {};
		this.additionalCorrectedContents = additionalCorrectedContents || {};
	}
	ParagraphUpdateDataCollector.prototype.checkExistingParagraphInOutline = function (paragraph, updateData) {
		const id = paragraph.Get_Id();
		const outlineView = this.outlineView;
		if (outlineView.sourceToOutlineMap[id]) {
			updateData.addExistingParagraph(outlineView.sourceToOutlineMap[id]);
			return;
		}
		const shape = paragraph.GetParentShape();
		if (shape && (shape === outlineView.outlineShape)) {
			updateData.addExistingParagraph(paragraph);
			const info = outlineView.outlineInfo.getSlideInfoByParagraph(paragraph);
			const slide = info && info.slide;
			if (slide) {
				this.additionalCheckSlideMap[slide.GetId()] = new CheckSlideInfo(slide);
			}
		}
	}
	ParagraphUpdateDataCollector.prototype.isUseInDocument = function (paragraph) {
		return paragraph.IsUseInDocument();
	};
	ParagraphUpdateDataCollector.prototype.checkExistingParagraphInPresentation = function (paragraph, updateData) {
		const shape = paragraph.GetParentShape();
		const isCorrectShape = shape && AscFormat.isSlideLikeObject(shape.parent) && shape.parent.isSlide() && shape.isOutlinePlaceholder();
		if (!isCorrectShape) {
			return;
		}
		const content = shape.getDocContent();
		const isCorrectContent = content && (shape.isTitlePlaceholder() || !content.Is_Empty() || this.additionalCorrectedContents[content.GetId()]);
		if (!isCorrectContent) {
			return;
		}
		const slide = shape.parent;
		const slideId = slide.GetId();
		if (this.isUseInDocument(paragraph)) {
			updateData.addNewParagraph(paragraph);
		} else {
			this.additionalCheckSlideMap[slideId] = new CheckSlideInfo(slide);
		}
	};
	ParagraphUpdateDataCollector.prototype.collect = function () {
		const updateData = new ParagraphUpdateData();
		for (let id in this.paragraphMap) {
			const paragraph = this.paragraphMap[id];
			this.checkExistingParagraphInOutline(paragraph, updateData);
			this.checkExistingParagraphInPresentation(paragraph, updateData);
		}
		return updateData;
	};

	function ParagraphUpdateData() {
		this.existingParagraphs = [];
		this.newParagraphs = [];
	}
	ParagraphUpdateData.prototype.isNeedRecalculate = function () {
		return !!this.existingParagraphs.length || !!this.newParagraphs.length;
	};
	ParagraphUpdateData.prototype.addExistingParagraph = function (paragraph) {
		this.existingParagraphs.push(paragraph);
	};
	ParagraphUpdateData.prototype.addNewParagraph = function (paragraph) {
		this.newParagraphs.push(paragraph);
	};

	function SlideUpdateData() {
		this.removedSlides = [];
		this.addedSlides = [];
	}
	SlideUpdateData.prototype.isNeedRecalculate = function () {
		return !!this.removedSlides.length || !!this.addedSlides.length;
	};

	const DECORATORDRAWER_OFFSET_LEFT = 2;
	const DECORATORDRAWER_SLIDE_WIDTH = 6;
	function DecoratorDrawer(outlineView, transform, selectedSlideMap, focusSlideIndex) {
		this.cache = {
			slideNumbersWidth: null
		};
		this.outlineView = outlineView;

		this.selectedSlideMap = selectedSlideMap || {};
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
	DecoratorDrawer.prototype.drawDecorShape = function (graphics, shape, x, y) {
		var t = new AscCommon.CMatrix();
		t.tx = this.getX(x);
		t.ty = this.getY(y);
		shape.draw(graphics, t, t);
	};
	DecoratorDrawer.prototype.drawTitleDecorations = function (graphics, info) {
		const slideNumberWidth = this.getSlideNumbersWidth();
		const decorCacheManager = this.getDecorCacheManager();
		const measureInfo = decorCacheManager.getMeasureInfo();
		const charWidth = measureInfo.charWidth;

		const paragraph = info.outlineParagraph;
		const slide = info.slide;
		const slideIconSizes = this.getSlideIconSizes(paragraph);
		if (!slideIconSizes) {
			return;
		}
		const presentation = this.getPresentation();
		const firstSlideNumber = presentation.getFirstSlideNumber();
		const currentSlideIndex = slide.num;
		const currentSlideLabel = String(currentSlideIndex + firstSlideNumber);
		const slideNumberShape = decorCacheManager.getSlideNumberShape(slideNumberWidth, slideIconSizes.height, currentSlideLabel);
		this.drawDecorShape(graphics, slideNumberShape, slideIconSizes.x - (currentSlideLabel.length * charWidth + 1), slideIconSizes.y);

		let penHex;
		if (this.selectedSlideMap[currentSlideIndex]) {
			penHex = this.activeSlideColor;
		} else if (currentSlideIndex === this.focusSlideIndex) {
			penHex = this.hoverSlideColor;
		} else {
			penHex = this.normalSlideColor
		}
		const barShape = decorCacheManager.getTitleShape(this.backgroundColor, penHex, slideIconSizes.width, slideIconSizes.height);
		this.drawDecorShape(graphics, barShape, slideIconSizes.x, slideIconSizes.y);
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
	DecoratorDrawer.prototype.getSlideIconSizes = function (outlineParagraph) {
		return this.outlineView.getSlideIconSizes(outlineParagraph);
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
	SavedPosition.prototype.isEmpty = function () {
		return !this.startPos && !this.endPos;
	};
	SavedPosition.prototype.apply = function (outlineView) {
		const docContent = outlineView.getDocContent();
		if (!docContent) {
			return false;
		}
		let isApplied = false;
		if (this.startPos) {
			docContent.RemoveSelection();
			const startPos = this.startPos.getPosition(outlineView);
			let endPos;
			if (this.endPos) {
				endPos = this.endPos.getPosition(outlineView);
			}
			if (startPos && endPos) {
				docContent.SetContentSelection(startPos, endPos, 0, 0, 0);
				isApplied = true;
			} else if (startPos) {
				docContent.SetContentPosition(startPos, 0, 0);
				isApplied = true;
			} else if (endPos) {
				docContent.SetContentPosition(endPos, 0, 0);
				isApplied = true;
			}

			outlineView.updateSelectionState();
		}
		return isApplied;
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
		let returnIndex;
		let startIndex;
		if (nextInfo) {
			const outlineParagraph = nextInfo.outlineParagraph;
			startIndex = outlineParagraph.Index - 1;
			returnIndex = outlineParagraph.Index;
		} else if (docContent.Content.length === 0) {
			return 0;
		} else {
			startIndex = docContent.Content.length - 1;
			returnIndex = startIndex;
		}
		for (let i = startIndex; i >= 0; i -= 1) {
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
				return i + 1;
			}
		}
		return returnIndex;
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
	UpdateNewParagraphsManager.prototype.isTitleFirstParagraph = function (sourceParagraph) {
		return this.getShapeIndex(sourceParagraph) === 0 && sourceParagraph.Index === 0;
	};
	UpdateNewParagraphsManager.prototype.update = function () {
		const outlineView = this.outlineView;
		const outlineDocContent = this.outlineView.getDocContent();
		const outlineParagraphs = outlineDocContent.Content;
		for (let i = 0; i < this.paragraphs.length; i++) {
			const sourceParagraph = this.paragraphs[i];
			const insertIndex = this.getInsertIndex(sourceParagraph, outlineParagraphs);
			let isUpdated = false;
			if (this.isTitleFirstParagraph(sourceParagraph)) {
				isUpdated = outlineView.addUpdatedParagraphWithCheckTitle(sourceParagraph, insertIndex);
			}
			if (!isUpdated) {
				outlineView.addUpdatedParagraph(sourceParagraph, insertIndex);
				this.resetContentShapeByShape(this.getParagraphShape(sourceParagraph));
			}
		}
	};
	UpdateNewParagraphsManager.prototype.resetContentShapeByShape = function (shape) {
		const outlineView = this.outlineView;
		const slide = shape.parent;
		const slideInfo = outlineView.outlineInfo.getSlideInfoBySlide(slide);
		if (slideInfo) {
			slideInfo.resetContentShapeInfoMapByShape(shape);
		}
	}
	UpdateNewParagraphsManager.prototype.getInsertIndex = function (sourceParagraph, outlineParagraphs) {
		const outlineView = this.outlineView;
		for (let i = outlineParagraphs.length - 1; i >= 0; i -= 1) {
			const outlineParagraph = outlineParagraphs[i];
			const sourceOutlineParagraph = outlineView.getSourceParagraph(outlineParagraph);
			if (!sourceOutlineParagraph) {
				const slideInfo = outlineView.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
				const slideIndex = slideInfo.slide.num;
				const sourceSlideIndex = this.getSlideIndex(sourceParagraph);
				if (slideIndex === sourceSlideIndex) {
					return this.isTitleFirstParagraph(sourceParagraph) ? i : i + 1;
				} else if (slideIndex < sourceSlideIndex) {
					return i + 1;
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
		if (!content.Content.length) {
			return contents;
		}
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
		return this.processNullableContent({outlineParagraph: contentInfo.outlineParagraph, index: 0, count: 1, direction: contentInfo.direction}, callback);

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
		const oldPos = sourceContent.GetContentPosition(false, false);
		sourceContent.SetContentPosition(startPos, 0, 0);
		const res = callback({content: sourceContent, index: 0, count: 1});
		sourceContent.RemoveSelection();
		sourceContent.SetContentPosition(oldPos, 0, 0);
		return res;
	};
	ForEachSelectManager.prototype.forEachSelectedContent = function (callback) {
		const content = this.getDocContent();
		if (!content || !content.Content.length) {
			return;
		}
		const api = this.outlineView.getApi();
		api.isDisableSetCurrentElement = true;
		let res;
		if (content.IsSelectionUse()) {
			res = this.processSelectedContent(callback);
		} else {
			res = this.processCursorContent(callback);
		}
		api.isDisableSetCurrentElement = false;
		return res;
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
	RemoveOutlineParagraphsManager.prototype.processNotSelectedContent = function (direction) {
		if (direction === AscWord.Direction.FORWARD) {
			if (this.isNeedProcessCursorAtEnd()) {
				this.processCursorAtEnd();
				return;
			}
		} else if (this.isNeedProcessCursorAtBegin()) {
			this.processCursorAtBegin();
			return;
		}
		this.processCommonPosContent(direction);
	};
	RemoveOutlineParagraphsManager.prototype.processCommonPosContent = function (direction) {
		const oThis = this;
		const outlineView = this.outlineView;
		outlineView.forEachSelectedContent(function (selectProps) {
			const content = selectProps.content;
			if (content) {
				content.Remove(direction, true, false, false);
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
		this.processEdgeParagraph(outlineParagraph, AscWord.Direction.BACKWARD);
	};
	RemoveOutlineParagraphsManager.prototype.processCursorAtEnd = function () {
		const docContent = this.getDocContent();
		const outlineParagraph = this.getContentPosParagraph(docContent);
		const nextOutlineParagraph = docContent.Content[outlineParagraph.Index + 1];
		if (!nextOutlineParagraph) {
			return;
		}
		this.processEdgeParagraph(nextOutlineParagraph, AscWord.Direction.FORWARD);
	};
	RemoveOutlineParagraphsManager.prototype.processEdgeParagraph = function (outlineParagraph, direction) {
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			this.processRealCursorAtBegin(outlineParagraph, direction);
		} else {
			this.processNullableContent(outlineParagraph, true);
		}
	};
	RemoveOutlineParagraphsManager.prototype.isNeedConcatEdgeParagraphs = function (sourceParagraph, direction) {
		if (sourceParagraph.Index !== 0) {
			return false;
		}
		return direction === AscWord.Direction.FORWARD || !sourceParagraph.IsPresentationBulleted();
	};
	RemoveOutlineParagraphsManager.prototype.processRealCursorAtBegin = function (outlineParagraph, direction) {
		const docContent = this.getDocContent();
		const currentSourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (currentSourceParagraph && !this.isNeedConcatEdgeParagraphs(currentSourceParagraph, direction)) {
			this.processCommonPosContent(direction);
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
	RemoveOutlineParagraphsManager.prototype.isNeedProcessCursorAtEnd = function () {
		const docContent = this.getDocContent();
		const paragraph = this.getContentPosParagraph(docContent);
		return paragraph.IsCursorAtEnd();
	};
	RemoveOutlineParagraphsManager.prototype.processSourceSelectedContent = function (content, idx, count, isOnAddText, edgeContents) {
		const shape = content.Is_DrawingShape(true);
		const slide = shape.parent;
		const slideNum = slide.num;
		this.checkSlidesForRemove[slideNum] = slide;
		if (content.IsSelectedAll() || content.IsEmpty()) {
			content.Remove(1, true, false, false);
		} else {
			content.Remove(1, true, false, count === 1 && isOnAddText);
		}
		if (idx === 0) {
			edgeContents.firstContent = content;
		} else if (idx === count - 1) {
			edgeContents.lastContent = content;
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
	RemoveOutlineParagraphsManager.prototype.saveMockParagraphSlidePos = function (outlineParagraph) {
		const info = this.outlineView.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		if (info) {
			this.saveManager.saveSlideStartPos(info.slide);
		}
	};
	RemoveOutlineParagraphsManager.prototype.processSelectedContent = function (isOnAddText) {
		const savedContentsOnUpdate = {};
		const outlineView = this.outlineView;
		const edgeContents = {firstContent: null, lastContent: null, firstOutlineParagraph: null};
		const oThis = this;
		let isSingleContent = false;
		outlineView.forEachSelectedContent(function (selectProps) {
			oThis.saveFirstOutlineParagraph(edgeContents, selectProps);
			if (selectProps.outlineParagraph) {
				const outlineParagraph = selectProps.outlineParagraph;
				if (isOnAddText && selectProps.count === 1) {
					oThis.saveMockParagraphSlidePos(outlineParagraph);
				}
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

		if (isOnAddText && edgeContents.firstContent) {
			savedContentsOnUpdate[edgeContents.firstContent.GetId()] = edgeContents.firstContent;
		}
		if (isSingleContent) {
			return savedContentsOnUpdate;
		}

		if (edgeContents.firstContent && edgeContents.lastContent) {
			this.concatEdgeContents(edgeContents.firstContent, edgeContents.lastContent);
		} else if (edgeContents.firstContent) {
			const content = edgeContents.firstContent.Content;
			this.saveManager.saveEndPosSourceParagraph(content[content.length - 1]);
		} else if (edgeContents.lastContent && edgeContents.firstOutlineParagraph) {
			this.processLastContent(edgeContents.lastContent, edgeContents.firstOutlineParagraph);
		}
		return savedContentsOnUpdate;
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
	RemoveOutlineParagraphsManager.prototype.remove = function (isOnAddText, direction) {
		const docContent = this.getDocContent();
		if (!docContent || !docContent.Content.length) {
			return;
		}
		if (isOnAddText) {
			if (!docContent.IsSelectionUse()) {
				return;
			}
		}
		const outlineView = this.outlineView;
		let savedContentsOnUpdate;
		if (docContent.IsSelectionUse()) {
			savedContentsOnUpdate = this.processSelectedContent(isOnAddText);
		} else {
			const removeDirection = direction === AscWord.Direction.FORWARD ? AscWord.Direction.FORWARD : AscWord.Direction.BACKWARD;
			this.processNotSelectedContent(removeDirection);
		}
		outlineView.removeSlides(this.checkSlidesForRemove, isOnAddText, savedContentsOnUpdate);
		return savedContentsOnUpdate;
	};
	function getSourceParagraphStartPos(sourceParagraph) {
		const contentPos = [{Class: sourceParagraph.GetParent(), Position: sourceParagraph.GetIndex()}];
		sourceParagraph.GetStartContentPosition(contentPos);
		return contentPos;
	}
	function getSourceParagraphEndPos(sourceParagraph, isBehindEnd) {
		const contentPos = [{Class: sourceParagraph.GetParent(), Position: sourceParagraph.GetIndex()}];
		sourceParagraph.GetEndContentPosition(contentPos, isBehindEnd);
		return contentPos;
	}
	function getSourceContentEdgePos(content, isStart) {
		const paragraph = isStart ? content.Content[0] : content.Content[content.Content.length - 1];
		if (!paragraph) {
			return null;
		}

		return isStart ? getSourceParagraphStartPos(paragraph) : getSourceParagraphEndPos(paragraph, true);
	}
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
	SavePosManager.prototype.saveSlideStartPos = function (slide) {
		this.outlineView.setSavedPosition(new SavedPosition(new SlidePosition(slide, true)));
	};
	SavePosManager.prototype.saveContentPosSourceParagraph = function (sourceContent) {
		const content = sourceContent;
		const contentPos = content.GetContentPosition(false, true);
		this.outlineView.setSavedPosition(new SavedPosition(new ContentPosition(contentPos)));
	};
	SavePosManager.prototype.saveEndPosSourceParagraph = function (sourceParagraph) {
		const contentPos = getSourceParagraphEndPos(sourceParagraph);
		this.outlineView.setSavedPosition(new SavedPosition(new ContentPosition(contentPos)));
	};
	SavePosManager.prototype.saveStartPosSourceParagraph = function (sourceParagraph) {
		const contentPos = getSourceParagraphStartPos(sourceParagraph);
		this.outlineView.setSavedPosition(new SavedPosition(new ContentPosition(contentPos)));
	};

	function getSourceParagraphLevel(sourceParagraph) {
		return AscFormat.isRealNumber(sourceParagraph.Pr.Lvl) ? sourceParagraph.Pr.Lvl : 0;
	}
	function getNestedParagraphs(sourceParagraphs) {
		return sourceParagraphs.filter(function (sourceParagraph) {
			return getSourceParagraphLevel(sourceParagraph) !== 0;
		});
	}
	function getTopLevelParagraphIndex(sourceParagraphs) {
		for (let i = 0; i < sourceParagraphs.length; i += 1) {
			if (getSourceParagraphLevel(sourceParagraphs[i]) === 0) {
				return i;
			}
		}
		return -1;
	}
	function ChangeOutlineLevelManager(outlineView) {
		this.outlineView = outlineView;
	}
	ChangeOutlineLevelManager.prototype.getSelectedOutlineParagraphs = function () {
		const docContent = this.outlineView.getDocContent();
		if (!docContent || !docContent.Content.length) {
			return [];
		}
		const selectedParagraphs = [];
		docContent.GetCurrentParagraph(false, selectedParagraphs);
		return selectedParagraphs;
	};
	ChangeOutlineLevelManager.prototype.isOutlineParagraphInTitle = function (outlineParagraph) {
		return !!this.outlineView.getSlideInfoByTitleParagraph(outlineParagraph);
	};
	ChangeOutlineLevelManager.prototype.getSlideOutlineContents = function (slide, contents) {
		contents = contents || [];
		const outlineSlide = slide.getOutlineSlide();
		const shapes = outlineSlide.getOutlineShapes();
		for (let i = 0; i < shapes.length; i += 1) {
			const shape = shapes[i];
			const content = shape && shape.getDocContent();
			if (!content || (shape !== outlineSlide.title && content.Is_Empty())) {
				continue;
			}
			contents.push(content);
		}
		return contents;
	};
	ChangeOutlineLevelManager.prototype.getShapeDocContent = function (shape, slide) {
		if (!shape) {
			return this.outlineView.createSubtitleContent(slide);
		}
		if (!shape.txBody || !shape.txBody.content) {
			shape.createTextBody();
		}
		return shape.getDocContent();
	};
	ChangeOutlineLevelManager.prototype.getSlideJoinContent = function (slide) {
		const contentShapes = slide.getOutlineSlide().content;
		let lastContentShape = contentShapes[0];
		for (let i = contentShapes.length - 1; i >= 1; i -= 1) {
			const shape = contentShapes[i];
			const content = shape.getDocContent();
			if (content && !content.Is_Empty()) {
				lastContentShape = shape;
				break;
			}
		}
		return this.getShapeDocContent(lastContentShape, slide);
	};
	ChangeOutlineLevelManager.prototype.getSlideFirstContent = function (slide) {
		return this.getShapeDocContent(slide.getOutlineSlide().content[0], slide);
	};
	ChangeOutlineLevelManager.prototype.collectContentsParagraphs = function (contents, paragraphs) {
		paragraphs = paragraphs || [];
		for (let i = 0; i < contents.length; i += 1) {
			const content = contents[i];
			content.GetAllParagraphs({All: true, Shapes: false}, paragraphs);
		}
		return paragraphs;
	};
	ChangeOutlineLevelManager.prototype.addParagraphsToContent = function (targetContent, sourceParagraphs) {
		const addedParagraphs = [];
		if (!sourceParagraphs.length) {
			return addedParagraphs;
		}
		if (targetContent.Is_Empty()) {
			targetContent.ClearContent(false);
		}
		for (let i = 0; i < sourceParagraphs.length; i += 1) {
			const newParagraph = sourceParagraphs[i].Copy2(targetContent);
			targetContent.Add_ToContent(targetContent.Content.length, newParagraph);
			addedParagraphs.push(newParagraph);
		}
		return addedParagraphs;
	};
	ChangeOutlineLevelManager.prototype.addSlideAfterSlide = function (slide) {
		const presentation = this.outlineView.getPresentation();
		return presentation.addNextSlideByPosAction(undefined, slide.num);
	};
	ChangeOutlineLevelManager.prototype.hasTitleOutlineParagraph = function (outlineParagraphs) {
		for (let i = 0; i < outlineParagraphs.length; i += 1) {
			if (this.isOutlineParagraphInTitle(outlineParagraphs[i])) {
				return true;
			}
		}
		return false;
	};
	ChangeOutlineLevelManager.prototype.getTitleSlides = function (outlineParagraphs) {
		const slides = [];
		for (let i = 0; i < outlineParagraphs.length; i += 1) {
			const outlineParagraph = outlineParagraphs[i];
			const slideInfo = this.outlineView.getSlideInfoByTitleParagraph(outlineParagraph);
			const slide = slideInfo && slideInfo.slide;
			if (slide && slides.indexOf(slide) === -1) {
				slides.push(slide);
			}
		}
		return slides;
	};
	ChangeOutlineLevelManager.prototype.getContentSourceParagraphs = function (outlineParagraphs) {
		const sourceParagraphs = [];
		for (let i = 0; i < outlineParagraphs.length; i += 1) {
			const outlineParagraph = outlineParagraphs[i];
			if (this.isOutlineParagraphInTitle(outlineParagraph)) {
				continue;
			}
			const sourceParagraph = this.outlineView.getSourceParagraph(outlineParagraph);
			if (sourceParagraph) {
				sourceParagraphs.push(sourceParagraph);
			}
		}
		return sourceParagraphs;
	};
	ChangeOutlineLevelManager.prototype.isCanChangeParagraphsLevel = function (sourceParagraphs, isIncrease) {
		if (!sourceParagraphs.length) {
			return false;
		}
		for (let i = 0; i < sourceParagraphs.length; i += 1) {
			if (!sourceParagraphs[i].Can_IncreaseLevel(isIncrease)) {
				return false;
			}
		}
		return true;
	};
	ChangeOutlineLevelManager.prototype.changeParagraphsLevel = function (sourceParagraphs, isIncrease) {
		if (!this.isCanChangeParagraphsLevel(sourceParagraphs, isIncrease)) {
			return false;
		}
		const contentMap = {};
		for (let i = 0; i < sourceParagraphs.length; i += 1) {
			const sourceParagraph = sourceParagraphs[i];
			sourceParagraph.Increase_Level(isIncrease);
			const content = sourceParagraph.GetParent();
			if (content) {
				contentMap[content.GetId()] = content;
			}
		}
		for (let id in contentMap) {
			const content = contentMap[id];
			const shape = content.Is_DrawingShape(true);
			if (shape && shape.checkExtentsByDocContent) {
				shape.checkExtentsByDocContent();
			}
		}
		return true;
	};
	ChangeOutlineLevelManager.prototype.changeSelectionLevel = function (sourceParagraphs, isIncrease) {
		if (!sourceParagraphs.length) {
			return;
		}
		const state = this.outlineView.saveDocumentStateBeforeLoadChanges();
		this.changeParagraphsLevel(sourceParagraphs, isIncrease);
		this.outlineView.applyDocumentStateAfterLoadChanges(state);
	};
	ChangeOutlineLevelManager.prototype.setSavedSelection = function (firstSourceParagraph, lastSourceParagraph, isBackward) {
		if (!firstSourceParagraph || !lastSourceParagraph) {
			return;
		}
		let startPos = getSourceParagraphStartPos(firstSourceParagraph);
		let endPos = getSourceParagraphEndPos(lastSourceParagraph, true);
		if (isBackward) {
			const temp = startPos;
			startPos = endPos;
			endPos = temp;
		}
		this.outlineView.setSavedPosition(new SavedPosition(new ContentPosition(startPos), new ContentPosition(endPos)));
	};
	ChangeOutlineLevelManager.prototype.getSlideTailParagraphs = function (sourceParagraph) {
		const shape = sourceParagraph.GetParentShape();
		const slide = shape && shape.parent;
		if (!slide) {
			return [];
		}
		const contentShapes = slide.getOutlineSlide().content;

		const tailContents = [];
		const startContentIndex = contentShapes.indexOf(shape) + 1;
		for (let i = startContentIndex; i < contentShapes.length; i += 1) {
			const content = contentShapes[i].getDocContent();
			if (content && !content.Is_Empty()) {
				tailContents.push(content);
			}
		}
		const sourceContent = sourceParagraph.GetParent();
		const startIndex = sourceParagraph.GetIndex() + 1;
		const paragraphs = sourceContent.Content.slice(startIndex);
		return this.collectContentsParagraphs(tailContents, paragraphs);
	};
	ChangeOutlineLevelManager.prototype.removeSlideParagraphsFrom = function (sourceParagraph) {
		const shape = sourceParagraph.GetParentShape();
		const slide = shape && shape.parent;
		if (!slide) {
			return;
		}
		const contentShapes = slide.getOutlineSlide().content;
		const sourceContent = sourceParagraph.GetParent();
		const startIndex = sourceParagraph.GetIndex();
		if (startIndex === 0) {
			sourceContent.Clear_Content();
		} else {
			sourceContent.Remove_FromContent(startIndex, sourceContent.Content.length - startIndex);
		}

		const startContentIndex = contentShapes.indexOf(shape) + 1;
		for (let i = startContentIndex; i < contentShapes.length; i += 1) {
			const content = contentShapes[i].getDocContent();
			if (content && !content.Is_Empty()) {
				content.Clear_Content();
			}
		}
	};
	ChangeOutlineLevelManager.prototype.getSplitGroups = function (sourceParagraphs, tailParagraphs) {
		const groups = [];
		for (let i = 0; i < sourceParagraphs.length; i += 1) {
			const sourceParagraph = sourceParagraphs[i];
			if (getSourceParagraphLevel(sourceParagraph) === 0) {
				groups.push({titleParagraph: sourceParagraph, contentParagraphs: [], promoteCount: 0});
			} else {
				const group = groups[groups.length - 1];
				group.contentParagraphs.push(sourceParagraph);
				group.promoteCount += 1;
			}
		}
		const lastGroup = groups[groups.length - 1];
		lastGroup.contentParagraphs = lastGroup.contentParagraphs.concat(tailParagraphs);
		return groups;
	};
	ChangeOutlineLevelManager.prototype.mapParagraphCopies = function (paragraphCopyMap, sourceParagraphs, addedParagraphs) {
		for (let i = 0; i < sourceParagraphs.length; i += 1) {
			paragraphCopyMap[sourceParagraphs[i].Get_Id()] = addedParagraphs[i];
		}
	};
	ChangeOutlineLevelManager.prototype.fillSplitSlide = function (newSlide, group, paragraphCopyMap) {
		const titleShape = newSlide.getMatchingShape(AscFormat.phType_title, null, false, {}) || newSlide.createTitle();
		if (!titleShape) {
			return [];
		}
		const titleContent = this.getShapeDocContent(titleShape, newSlide);
		titleContent.Clear_Content();
		const addedTitles = this.addParagraphsToContent(titleContent, [group.titleParagraph]);
		this.mapParagraphCopies(paragraphCopyMap, [group.titleParagraph], addedTitles);
		if (!group.contentParagraphs.length) {
			return addedTitles;
		}
		const slideContent = this.getSlideFirstContent(newSlide);
		slideContent.Clear_Content();
		const addedContents = this.addParagraphsToContent(slideContent, group.contentParagraphs);
		this.mapParagraphCopies(paragraphCopyMap, group.contentParagraphs, addedContents);
		this.changeParagraphsLevel(addedContents.slice(0, group.promoteCount), false);
		return addedTitles.concat(addedContents);
	};
	ChangeOutlineLevelManager.prototype.isCanSplitSlideByParagraph = function (sourceParagraph) {
		const shape = sourceParagraph.GetParentShape();
		const slide = shape && shape.parent;
		return !!slide && slide.getOutlineSlide().content.indexOf(shape) !== -1;
	};
	ChangeOutlineLevelManager.prototype.captureSelectionEdgePos = function (isStart) {
		return this.outlineView.getDocContent().GetContentPosition(true, isStart);
	};
	ChangeOutlineLevelManager.prototype.captureCursorPos = function () {
		return this.outlineView.getDocContent().GetContentPosition(false, false);
	};
	ChangeOutlineLevelManager.prototype.remapCapturedPos = function (pos, paragraphCopyMap) {
		if (!pos || !pos[1]) {
			return null;
		}
		const sourceParagraph = this.outlineView.getSourceParagraph(pos[1].Class);
		if (!sourceParagraph) {
			return null;
		}
		const targetParagraph = paragraphCopyMap[sourceParagraph.Get_Id()] || sourceParagraph;
		return rebuildPos(pos, targetParagraph.GetParent(), targetParagraph);
	};
	ChangeOutlineLevelManager.prototype.splitSlideByParagraphs = function (sourceParagraphs, topLevelIndex) {
		const outlineView = this.outlineView;
		const presentation = outlineView.getPresentation();
		const leadingParagraphs = sourceParagraphs.slice(0, topLevelIndex);
		const splitParagraphs = sourceParagraphs.slice(topLevelIndex);
		const firstParagraph = splitParagraphs[0];
		const slide = firstParagraph.GetParentShape().parent;
		const tailParagraphs = this.getSlideTailParagraphs(splitParagraphs[splitParagraphs.length - 1]);
		const groups = this.getSplitGroups(splitParagraphs, tailParagraphs);

		let previousSlide = slide;
		let firstNewSlide = null;
		const paragraphCopyMap = {};
		for (let i = 0; i < groups.length; i += 1) {
			const newSlide = this.addSlideAfterSlide(previousSlide);
			if (!newSlide) {
				break;
			}
			previousSlide = newSlide;
			firstNewSlide = firstNewSlide || newSlide;
			this.fillSplitSlide(newSlide, groups[i], paragraphCopyMap);
		}
		if (!firstNewSlide) {
			return;
		}
		this.removeSlideParagraphsFrom(firstParagraph);
		this.changeParagraphsLevel(leadingParagraphs, false);

		this.setResultPosition(firstParagraph, paragraphCopyMap);

		outlineView.update();
		presentation.DrawingDocument.m_oWordControl.GoToPage(presentation.GetSlideIndex(firstNewSlide));
		outlineView.applySavedPositionToOutline();
	};
	ChangeOutlineLevelManager.prototype.getJoinTargetSlide = function (slides) {
		const presentation = this.outlineView.getPresentation();
		return presentation.Slides[presentation.GetSlideIndex(slides[0]) - 1] || null;
	};
	ChangeOutlineLevelManager.prototype.joinSlidesToPreviousSlide = function (slides, startParagraph) {
		const outlineView = this.outlineView;
		const presentation = outlineView.getPresentation();
		const previousSlide = this.getJoinTargetSlide(slides);
		if (!previousSlide) {
			return;
		}
		let contents = [];
		const slidesForRemove = {};
		//todo split removeSlides and getSlidesFromMap
		for (let i = 0; i < slides.length; i += 1) {
			const slide = slides[i];
			this.getSlideOutlineContents(slide, contents);
			slidesForRemove[slide.num] = slide;
		}
		const movedParagraphs = this.collectContentsParagraphs(contents);
		const targetContent = this.getSlideJoinContent(previousSlide);
		const addedParagraphs = this.addParagraphsToContent(targetContent, movedParagraphs);
		const paragraphCopyMap = {};
		this.mapParagraphCopies(paragraphCopyMap, movedParagraphs, addedParagraphs);
		for (let i = 0; i < contents.length; i += 1) {
			contents[i].Clear_Content();
		}

		this.setResultPosition(startParagraph || movedParagraphs[0], paragraphCopyMap);

		outlineView.removeSlides(slidesForRemove);
		outlineView.update();
		presentation.DrawingDocument.m_oWordControl.GoToPage(presentation.GetSlideIndex(previousSlide));
		outlineView.applySavedPositionToOutline();
	};
	ChangeOutlineLevelManager.prototype.setResultPosition = function (fallbackParagraph, paragraphCopyMap) {
		const outlineView = this.outlineView;
		const resultParagraph = fallbackParagraph ? (paragraphCopyMap[fallbackParagraph.Get_Id()] || fallbackParagraph) : null;

		function saveCollapsedResultPosition() {
			if (resultParagraph) {
				new SavePosManager(outlineView).saveStartPosSourceParagraph(resultParagraph);
			}
		}
		if (outlineView.isSelectionUse()) {
			const docContent = outlineView.getDocContent();
			const startSelectionPos = this.captureSelectionEdgePos(true);
			const endSelectionPos = this.captureSelectionEdgePos(false);
			const startPos = this.remapCapturedPos(startSelectionPos, paragraphCopyMap);
			const endPos = this.remapCapturedPos(endSelectionPos, paragraphCopyMap);
			if (startPos && endPos) {
				outlineView.setSavedPosition(new SavedPosition(new ContentPosition(startPos), new ContentPosition(endPos)));
			} else if (docContent.Selection.StartPos === docContent.Selection.EndPos) {
				saveCollapsedResultPosition();
			} else {
				const isBackward = docContent.GetSelectDirection() === AscWord.Direction.BACKWARD;
				this.setSavedSelection(resultParagraph, resultParagraph, isBackward);
			}
		} else {
			const cursorPos = this.captureCursorPos();
			const remappedCursorPos = this.remapCapturedPos(cursorPos, paragraphCopyMap);
			if (remappedCursorPos) {
				outlineView.setSavedPosition(new SavedPosition(new ContentPosition(remappedCursorPos)));
			} else {
				saveCollapsedResultPosition();
			}
		}
	};
	ChangeOutlineLevelManager.prototype.increaseParagraphIndent = function (outlineParagraphs) {
		if (!this.isCanIncreaseParagraphIndent(outlineParagraphs)) {
			return;
		}
		const sourceParagraphs = this.getContentSourceParagraphs(outlineParagraphs);
		const titleSlides = this.getTitleSlides(outlineParagraphs);
		if (!titleSlides.length) {
			this.changeSelectionLevel(sourceParagraphs, true);
			return;
		}
		this.changeParagraphsLevel(sourceParagraphs, true);
		this.joinSlidesToPreviousSlide(titleSlides, this.outlineView.getSourceParagraph(outlineParagraphs[0]));
	};
	ChangeOutlineLevelManager.prototype.decreaseParagraphIndent = function (outlineParagraphs) {
		if (!this.isCanDecreaseParagraphIndent(outlineParagraphs)) {
			return;
		}
		const sourceParagraphs = this.getContentSourceParagraphs(outlineParagraphs);
		const topLevelIndex = getTopLevelParagraphIndex(sourceParagraphs);
		if (topLevelIndex === -1) {
			this.changeSelectionLevel(sourceParagraphs, false);
			return;
		}
		this.splitSlideByParagraphs(sourceParagraphs, topLevelIndex);
	};
	ChangeOutlineLevelManager.prototype.isCanIncreaseParagraphIndent = function (outlineParagraphs) {
		const titleSlides = this.getTitleSlides(outlineParagraphs);
		if (titleSlides.length && !this.getJoinTargetSlide(titleSlides)) {
			return false;
		}
		const sourceParagraphs = this.getContentSourceParagraphs(outlineParagraphs);
		if (sourceParagraphs.length) {
			return this.isCanChangeParagraphsLevel(sourceParagraphs, true);
		}
		return !!titleSlides.length;
	};
	ChangeOutlineLevelManager.prototype.isCanDecreaseParagraphIndent = function (outlineParagraphs) {
		if (this.hasTitleOutlineParagraph(outlineParagraphs)) {
			return false;
		}
		const sourceParagraphs = this.getContentSourceParagraphs(outlineParagraphs);
		if (!sourceParagraphs.length) {
			return false;
		}
		const nestedParagraphs = getNestedParagraphs(sourceParagraphs);
		if (nestedParagraphs.length && !this.isCanChangeParagraphsLevel(nestedParagraphs, false)) {
			return false;
		}
		const topLevelIndex = getTopLevelParagraphIndex(sourceParagraphs);
		if (topLevelIndex === -1) {
			return true;
		}
		return this.isCanSplitSlideByParagraph(sourceParagraphs[topLevelIndex]);
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
		const isNeedUseNextShapeForInsert = firstShape.isGeneralTitle();

		if (firstSlide !== lastSlide) {
			this.processEdgeSlides(firstShape, lastShape, isNeedUseNextShapeForInsert);
		}
	};
	ConcatShapesManager.prototype.processEdgeSlides = function (firstShape, lastShape, isNeedUseNextShapeForInsert) {
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
				this.concatShapeContent(firstOutlineShapes[i], lastOutlineShapes[lastShapeIndex]);
				lastShapeIndex += 1;
			}
		}
		const edgeFirstShape = firstOutlineShapes[firstOutlineShapes.length - 1];
		for (let i = lastShapeIndex; i < lastOutlineShapes.length; i += 1) {
			this.concatShapeContent(edgeFirstShape, lastOutlineShapes[i]);
		}
	};
	ConcatShapesManager.prototype.concatShapeContent = function (firstShape, lastShape) {
		const lastContent = lastShape.getDocContent();
		if (!lastContent || lastContent.IsEmpty()) {
			return;
		}
		if (!firstShape.txBody || !firstShape.txBody.content) {
			firstShape.createTextBody();
		}
		const firstContent = firstShape.getDocContent();
		let i = 0;
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

	function ParagraphAddManager(outlineView, isRecalculate) {
		this.outlineView = outlineView;
		this.isRecalculate = isRecalculate;
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
		content.AddToParagraph(paraItem, this.isRecalculate);
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
	OutlineContentInsertManager.prototype.insertIntoSourceContent = function (content, selectedContent, isSelect) {
		const paragraph = content.Content[content.CurPos.ContentPos];
		if (!paragraph || !paragraph.IsParagraph()) {
			return false;
		}
		const nearPos = {Paragraph: paragraph, ContentPos: paragraph.Get_ParaContentPos(false, false)};
		paragraph.Check_NearestPos(nearPos);
		selectedContent.Insert(nearPos, isSelect);
		return true;
	};
	OutlineContentInsertManager.prototype.insert = function (selectedContent, isSelect) {
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
				inserted = oThis.insertIntoSourceContent(content, selectedContent, isSelect);
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
	OutlineInfo.prototype.setOutlineParagraph = function (slideInfo, outlineParagraph) {
		const previousParagraph = slideInfo.outlineParagraph;
		if (previousParagraph) {
			delete this.outlineParagraphToInfoMap[previousParagraph.GetId()];
		}
		slideInfo.setOutlineParagraph(outlineParagraph);
		this.outlineParagraphToInfoMap[outlineParagraph.GetId()] = slideInfo;
		this.slideToInfoMap[slideInfo.slide.GetId()] = slideInfo;
	};
	OutlineInfo.prototype.getSlideInfoByParagraph = function (outlineParagraph) {
		return this.outlineParagraphToInfoMap[outlineParagraph.GetId()];
	};
	OutlineInfo.prototype.getSlideInfoBySlide = function (slide) {
		return this.slideToInfoMap[slide && slide.GetId()] || null;
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
			textPr.SetUnifill(AscFormat.CreateUnifillFromRGB(rgb.R, rgb.G, rgb.B));
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
			paraPr.InitDefault();
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

	function getSelectionSetter(outlineView, target, isResetSelection) {
		if (isResetSelection) {
			return new SelectionSetter(outlineView, target);
		}
		return new ExpandSelectionSetter(outlineView, target);
	}
	function DecorationSelectManager(outlineView, e, x, y) {
		this.outlineView = outlineView;
		this.event = e;
		this.x = x;
		this.y = y;
	}
	DecorationSelectManager.prototype.getDocContent = function () {
		return this.outlineView.getDocContent();
	};
	DecorationSelectManager.prototype.handle = function () {
		const docContent = this.getDocContent();
		if (!docContent) {
			return;
		}
		const tx = this.outlineView.getInvertTransformX(this.x);
		const ty = this.outlineView.getInvertTransformY(this.y);

		const nearestPos = docContent.Get_NearestPos(0, tx, ty);
		if (!nearestPos || !nearestPos.Paragraph) {
			return;
		}
		const paragraph = nearestPos.Paragraph;
		const isSlide = this.isClickToSlideIcon(paragraph);
		const target = this.getTargetContentPositions(paragraph, isSlide);
		if (!target) {
			return;
		}
		const setter = getSelectionSetter(this.outlineView, target, !this.isExpandSelection());
		setter.select();
	};
	DecorationSelectManager.prototype.isClickToSlideIcon = function (outlineParagraph) {
		const sizes = this.outlineView.getSlideIconSizes(outlineParagraph);
		if (!sizes) {
			return false;
		}
		return this.x >= sizes.x && this.x <= sizes.x + sizes.width &&
			this.y >= sizes.y && this.y <= sizes.y + sizes.height;
	};
	DecorationSelectManager.prototype.getTargetContentPositions = function (outlineParagraph, isSlide) {
		const positionGetter = isSlide ? new SlidePositionGetter(this.outlineView, outlineParagraph) : new ParagraphPositionGetter(this.outlineView, outlineParagraph);
		return positionGetter.get();
	};
	DecorationSelectManager.prototype.isExpandSelection = function () {
		return this.event.CtrlKey || this.event.ShiftKey;
	};
	
	function OutlinePositionGetter(outlineView, paragraph) {
		this.outlineView = outlineView;
		this.paragraph = paragraph;
	}
	OutlinePositionGetter.prototype.getContentPosition = function (outlineParagraph, isStart) {
		const docContent = this.outlineView.getDocContent();
		const pos = [{Class: docContent, Position: outlineParagraph.Index}];
		if (isStart) {
			outlineParagraph.GetStartContentPosition(pos);
		} else {
			outlineParagraph.GetEndContentPosition(pos, true);
		}
		return pos;
	}
	OutlinePositionGetter.prototype.get = function () {};

	function SlidePositionGetter(outlineView, paragraph) {
		OutlinePositionGetter.call(this, outlineView, paragraph);
	}
	AscFormat.InitClassWithoutType(SlidePositionGetter, OutlinePositionGetter);
	SlidePositionGetter.prototype.get = function () {
		let firstParagraph = this.paragraph;
		let lastParagraph = this.paragraph;
		const info = this.outlineView.outlineInfo.getSlideInfoByParagraph(this.paragraph);
		const slide = info && info.slide;
		if (slide) {
			const paragraphs = this.outlineView.getOutlineSlideParagraphs(slide);
			if (paragraphs.length) {
				firstParagraph = paragraphs[0];
				lastParagraph = paragraphs[paragraphs.length - 1];
			}
		}
		return {
			startPos: this.getContentPosition(firstParagraph, true),
			endPos: this.getContentPosition(lastParagraph, false)
		};
	};

	function ParagraphPositionGetter(outlineView, paragraph) {
		OutlinePositionGetter.call(this, outlineView, paragraph);
	}
	AscFormat.InitClassWithoutType(ParagraphPositionGetter, OutlinePositionGetter);
	ParagraphPositionGetter.prototype.get = function () {
		return {
			startPos: this.getContentPosition(this.paragraph, true),
			endPos: this.getContentPosition(this.paragraph, false)
		};
	};

	function SelectionSetter(outlineView, target) {
		this.outlineView = outlineView;
		this.target = target;
	}
	
	SelectionSetter.prototype.setSelection = function (positions) {
		const docContent = this.outlineView.getDocContent();
		if (!docContent) {
			return;
		}
		docContent.SetContentPosition(positions.endPos, 0, 0);
		docContent.SetContentSelection(positions.startPos, positions.endPos, 0, 0, 0);
		this.outlineView.updateSelectionState();
	}
	SelectionSetter.prototype.select = function () {
		this.setSelection(this.target);
	}
	function ExpandSelectionSetter(outlineView, target) {
		SelectionSetter.call(this, outlineView, target);
	}
	AscFormat.InitClassWithoutType(ExpandSelectionSetter, SelectionSetter);
	ExpandSelectionSetter.prototype.select = function () {
		const cur = this.getCurrentContentPositions();
		const targetStartIndex = this.target.startPos[0].Position;
		const currentStartIndex = cur.startPos[0].Position;
		const targetEndIndex = this.target.endPos[0].Position;
		const currentEndIndex = cur.endPos[0].Position;
		let startPos;
		let endPos;
		let isNeedSetSelection = false;
		let isSwitchSelection = false;
		if (targetStartIndex <= currentStartIndex) {
			isNeedSetSelection = true;
			isSwitchSelection = true;
			startPos = this.target.startPos;
		} else {
			startPos = cur.startPos;
		}
		if (targetEndIndex >= currentEndIndex) {
			isNeedSetSelection = true;
			isSwitchSelection = false;
			endPos = this.target.endPos;
		} else {
			endPos = cur.endPos;
		}
		if (isNeedSetSelection) {
			const newSelectionPositions = isSwitchSelection ? {startPos: endPos, endPos: startPos} : {startPos: startPos, endPos: endPos};
			this.setSelection(newSelectionPositions);
		}
	};
	ExpandSelectionSetter.prototype.getCurrentContentPositions = function () {
		const docContent = this.outlineView.getDocContent();
		if (docContent.IsSelectionUse()) {
			const startPos = docContent.GetContentPosition(true, true);
			const endPos = docContent.GetContentPosition(true, false);
			if (startPos[0].Position <= endPos[0].Position) {
				return {startPos: startPos, endPos: endPos};
			}
			return {startPos: endPos, endPos: startPos};
		}
		const pos = docContent.GetContentPosition(false, false);
		return {startPos: pos, endPos: pos};
	};


	function OutlineSlidesSelectionManager(outlineView, slideIndexes, isResetSelection) {
		this.outlineView = outlineView;
		this.indexes = slideIndexes;
		this.isResetSelection = isResetSelection;
	}
	OutlineSlidesSelectionManager.prototype.getSelectRange = function () {
		if (!this.indexes.length) {
			return null;
		}
		let startSlideIndex = this.indexes[0];
		let endSlideIndex = this.indexes[0];

		for (let i = 1; i < this.indexes.length; i += 1) {
			const slideIndex = this.indexes[i];
			if (slideIndex < startSlideIndex) {
				startSlideIndex = slideIndex;
			}
			if (slideIndex > endSlideIndex) {
				endSlideIndex = slideIndex;
			}
		}
		return {start: startSlideIndex, end: endSlideIndex};
	};
	OutlineSlidesSelectionManager.prototype.getSlidePositions = function (slideIndex) {
		const presentation = this.outlineView.getPresentation();
		const slide = presentation.Slides[slideIndex];
		if (!slide) {
			return null;
		}
		return this.outlineView.getSlidePositions(slide);
	};
	OutlineSlidesSelectionManager.prototype.select = function () {
		const presentation = this.outlineView.getPresentation();
		if (!presentation) {
			return;
		}
		if (!presentation.IsOutlineMode()) {
			return;
		}
		this.outlineView.update();
		const selectRange = this.getSelectRange();
		if (!selectRange) {
			return;
		}

		const startSlidePositions = this.getSlidePositions(selectRange.start);
		const endSlidePositions = this.getSlidePositions(selectRange.end);
		if (!startSlidePositions || !endSlidePositions) {
			return;
		}

		const target = {startPos: startSlidePositions.startPos, endPos: endSlidePositions.endPos};
		const selectionSetter = getSelectionSetter(this.outlineView, target, this.isResetSelection);
		selectionSetter.select();
	};

	const OUTLINEVIEW_INDENT_SCALE = 0.3;
	function OutlineView(api) {
		this.api = api;
		this.isInit = false;
		this.outlineShape = null;
		this.isDecorationMouseDown = false;
		this.decorCacheManager = new DecorCacheManager();
		this.recalcInfo = {
			recalculateContent: true
		};
		this.reset();
	}
	OutlineView.prototype.getSlidePositions = function (slide) {
		const info = this.outlineInfo.getSlideInfoBySlide(slide);
		if (!info) {
			return null;
		}
		return (new SlidePositionGetter(this, info.outlineParagraph)).get();
	};
	OutlineView.prototype.getSlideNumbersWidth = function () {
			const measureInfo = this.decorCacheManager.getMeasureInfo();
			const presentation = this.getPresentation();
			const slidesCount = presentation.GetSlidesCount();
			const firstSlideNumber = presentation.getFirstSlideNumber();
			const totalSlidesLength = String(slidesCount - 1 + firstSlideNumber).length;
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
		this.compositeInputInfo = null;
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
			this.outlineInfo.setOutlineParagraph(pr, outlineParagraph);
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
			const copyParagraph = this.getCopyParagraph(outlineContent, paragraph, this.isTitleSourceParagraph(paragraph));
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
		const slideContent = slideShape.txBody && slideShape.txBody.content;
		if (!slideContent) {
			return;
		}
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
				if (contentShape.txBody && contentShape.txBody.content && !contentShape.txBody.content.IsEmpty()) {
					this.addContentToOutlineShape(outlineShape, slide.content[j], null);
					shapeCount += 1;
				}
			}
		}
	};
	OutlineView.prototype.getParaPr = function (compiledParaPr, isTitle) {
		const copyParaPr = new CParaPr();
		copyParaPr.InitDefault();
		if (compiledParaPr.ParaPr.Bullet) {
			copyParaPr.Bullet = compiledParaPr.ParaPr.Bullet.createDuplicate();
			copyParaPr.Bullet.bulletColor = new AscFormat.CBulletColor();
			copyParaPr.Bullet.bulletColor.type = AscFormat.BULLET_TYPE_COLOR_CLR;
			copyParaPr.Bullet.bulletColor.UniColor = AscFormat.CreateUniColorFromRGB(0, 0, 0);
			copyParaPr.Lvl = compiledParaPr.ParaPr.Lvl;
		}
		copyParaPr.Ind = compiledParaPr.ParaPr.Ind.Copy();
		copyParaPr.Ind.FirstLine *= OUTLINEVIEW_INDENT_SCALE;
		if (isTitle) {
			copyParaPr.Ind.Left = copyParaPr.Ind.FirstLine < 0 ? -copyParaPr.Ind.FirstLine : 0;
			copyParaPr.Ind.Right = 0;
		} else {
			copyParaPr.Ind.Left *= OUTLINEVIEW_INDENT_SCALE;
			copyParaPr.Ind.Right *= OUTLINEVIEW_INDENT_SCALE;
		}

		copyParaPr.Spacing.Before = 1;
		copyParaPr.Spacing.After = 1;
		copyParaPr.Spacing.Line = 1;
		return copyParaPr;
	};
	OutlineView.prototype.applyParagraphProps = function (outlineParagraph, slideParagraph, isTitle) {
		const compiledPr = slideParagraph ? slideParagraph.getCompiledPr() : {ParaPr: g_oDocumentDefaultParaPr};
		const copyParaPr = this.getParaPr(compiledPr, isTitle);
		outlineParagraph.SetPr(copyParaPr);
		outlineParagraph.CheckRunContent(function (run) {
			const textPr = getOutlineTextPr(isTitle, run.IsMathRun());
			const runPr = run.GetDirectTextPr();
			if (runPr) {
				textPr.SetItalic(runPr.Italic);
				textPr.SetUnderline(runPr.Underline);
				textPr.SetStrikeout(runPr.Strikeout);
				textPr.SetDStrikeout(runPr.DStrikeout);
				textPr.SetCaps(runPr.Caps);
				textPr.SetSmallCaps(runPr.SmallCaps);
				textPr.SetVertAlign(runPr.VertAlign);
				textPr.SetSpacing(AscFormat.isRealNumber(runPr.Spacing) ? runPr.Spacing * OUTLINEVIEW_INDENT_SCALE : runPr.Spacing);
			}
			textPr.SetUnifill(AscFormat.CreateUnifillFromRGB(0, 0, 0));
			run.SetPr(textPr);
		});
		const parTextPr = getOutlineTextPr();
		outlineParagraph.TextPr.Value = parTextPr;
		outlineParagraph.TextPr.CalcValue = parTextPr;
	};
	OutlineView.prototype.getStyles = function () {
		return AscFormat.ExecuteNoHistory(function () {
			const styles = new CStyles(false);
			const default_style = new CStyle("defaultStyle", null, null, null, true);
			styles.Add(default_style);
			return {styles: styles, lastId: default_style.Id, outlineView: this};
		}, this, []);
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
			outlineShape.stylesOverride = this.getStyles();
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
			if (this.outlineShape.extX && this.outlineShape.extY) {
				this.outlineShape.recalculateContent();
			}
			this.outlineShape.contentWidth = width;
		}
	}
	OutlineView.prototype.getCopyParagraph = function (parent, paragraph, isTitle) {
		return AscFormat.ExecuteNoHistory(function () {
			const copyParagraph = paragraph.Copy2(parent);
			this.applyParagraphProps(copyParagraph, paragraph, isTitle);
			return copyParagraph;
		}, this, []);
	};
	OutlineView.prototype.draw = function (graphics) {
		if (this.outlineShape) {
			this.getPresentation().ensureRecalculateOnOpen();
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
	OutlineView.prototype.drawDecorations = function (graphics, selectedSlideMap, focusSlideIndex) {
		if (!this.outlineShape || !this.isHaveParagraphs()) return;

		const decoratorDrawer = new DecoratorDrawer(this, new AscCommon.CMatrix(), selectedSlideMap, focusSlideIndex);

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

	OutlineView.prototype.handleDecorationSelect = function (e, x, y) {
		const selectManager = new DecorationSelectManager(this, e, x, y);
		selectManager.handle();
	};
	OutlineView.prototype.selectionSetStart = function (e, x, y) {
		if (this.outlineShape) {
			if (x < this.outlineShape.x) {
				this.isDecorationMouseDown = true;
				this.handleDecorationSelect(e, x, y);
			} else {
				this.isDecorationMouseDown = false;
				this.outlineShape.selectionSetStart(e, x, y, 0);
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
			return this.outlineShape.contentHeight;
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

	OutlineView.prototype.isSelectionUse = function () {
		const content = this.getDocContent();
		return !!content && content.IsSelectionUse();
	};
	OutlineView.prototype.getSelectedText = function (oPr) {
		const docContent = this.getDocContent();
		if (!docContent) {
			return "";
		}
		return docContent.GetSelectedText(false, oPr);
	};

	OutlineView.prototype.remove = function (isOnAddText, direction) {
		const docContent = this.getDocContent();
		const isSelectionUse = docContent && docContent.IsSelectionUse();
		const removeStartIndex = this.getRemoveStartIndex();
		const manager = new RemoveOutlineParagraphsManager(this);
		 const removeInfo = manager.remove(isOnAddText, direction);
		this.update(removeInfo);
		const isPositionSaved = !this.isEmptySavedPosition();
		const isPositionApplied = this.applySavedPositionToOutline();
		if (isPositionSaved && !isPositionApplied && !isOnAddText) {
			this.moveCursorToOutlineIndex(removeStartIndex);
		}
		if (isOnAddText && isSelectionUse) {
			this.checkCurrentTextObjectExtends();
		}
	};
	OutlineView.prototype.getRemoveStartIndex = function () {
		const docContent = this.getDocContent();
		if (!docContent) {
			return -1;
		}
		if (!docContent.IsSelectionUse()) {
			return docContent.CurPos.ContentPos;
		}
		return Math.min(docContent.Selection.StartPos, docContent.Selection.EndPos);
	};
	OutlineView.prototype.moveCursorToOutlineIndex = function (index) {
		const docContent = this.getDocContent();
		if (!docContent || index < 0 || !docContent.Content.length) {
			return;
		}
		docContent.Update_ContentIndexing();
		docContent.RemoveSelection();
		if (index < docContent.Content.length) {
			const paragraph = docContent.Content[index];
			paragraph.SetThisElementCurrent();
			paragraph.MoveCursorToStartPos();
		} else {
			const paragraph = docContent.Content[docContent.Content.length - 1];
			paragraph.SetThisElementCurrent();
			paragraph.MoveCursorToEndPos();
		}
		this.updateSelectionState();
	};
	OutlineView.prototype.removeSlides = function (slideForRemoveMap, isOnAddText, savedContentsOnUpdate) {
		const slides = [];
		let slideNumbers = Object.keys(slideForRemoveMap).map(function (num) {
			return parseInt(num, 10);
		});
		if (!slideNumbers.length) {
			return;
		}
		let savedMinSlideIndex = null;
		if (isOnAddText) {
			let minSlideNumber = slideNumbers[0];
			for (let i = 1; i < slideNumbers.length; i += 1) {
				if (minSlideNumber > slideNumbers[i]) {
					minSlideNumber = slideNumbers[i];
				}
			}
			savedMinSlideIndex = minSlideNumber;
			slideNumbers = slideNumbers.filter(function (num) {
				return num !== minSlideNumber;
			});
		}
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
		let minSlideIndex = savedMinSlideIndex === null ? presentation.GetSlidesCount() - 1 : savedMinSlideIndex;
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
		this.update(savedContentsOnUpdate);
		const lastSlideIndex = presentation.GetSlidesCount() - 1;
		const currentPage = Math.min(lastSlideIndex, minSlideIndex);
		presentation.Recalculate();
		presentation.DrawingDocument.m_oWordControl.GoToPage(currentPage, undefined, undefined, true);
		const docContent = this.getDocContent();
		if (docContent && currentPage !== -1 && this.isEmptySavedPosition()) {
			docContent.RemoveSelection();
			const slidePositions = this.getSlidePositions(presentation.Slides[currentPage]);
			if (lastSlideIndex >= minSlideIndex) {
				docContent.SetContentPosition(slidePositions.startPos, 0, 0);
			} else {
				docContent.SetContentPosition(slidePositions.endPos, 0, 0);
			}
		}

		presentation.Api.sync_HideComment();
	};
	OutlineView.prototype.isEmptySavedPosition = function () {
		const savedPosition = this.getSavedPosition();
		if (!savedPosition) {
			return true;
		}
		return savedPosition.isEmpty();
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
		const isApplied = savedPosition.apply(this);
		this.resetPosition();
		return isApplied;
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
	OutlineView.prototype.paragraphAdd = function (paraItem, isRecalculate) {
		const manager = new ParagraphAddManager(this, isRecalculate);
		manager.paragraphAdd(paraItem);
		this.update();
		this.applySavedPositionToOutline();
		this.checkCurrentTextObjectExtends();
	}
	OutlineView.prototype.insertContent = function (selectedContent, isSelect) {
		const manager = new OutlineContentInsertManager(this);
		const inserted = manager.insert(selectedContent, isSelect);
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
	OutlineView.prototype.isTitleSourceParagraph = function (sourceParagraph) {
		const parentShape = sourceParagraph && sourceParagraph.GetParentShape();
		return !!(parentShape && parentShape.isGeneralTitle());
	};
	OutlineView.prototype.getTitleOutlineParagraph = function (sourceParagraph, slide) {
		const parentShape = sourceParagraph.GetParentShape() || slide.getOutlineTitleShape();
		if (!parentShape || parentShape.parent !== slide) {
			return null;
		}
		const content = parentShape.getDocContent();
		const firstSourceParagraph = content && content.Content[0];
		return firstSourceParagraph ? (this.sourceToOutlineMap[firstSourceParagraph.GetId()] || null) : null;
	};
	OutlineView.prototype.checkMoveSlideInfoToTitleParagraph = function (outlineParagraph) {
		const info = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		if (!info) {
			return;
		}
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (!sourceParagraph) {
			return;
		}
		const titleOutlineParagraph = this.getTitleOutlineParagraph(sourceParagraph, info.slide);
		if (titleOutlineParagraph && titleOutlineParagraph !== outlineParagraph) {
			this.outlineInfo.setOutlineParagraph(info, titleOutlineParagraph);
		}
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

	OutlineView.prototype.addUpdatedParagraphWithCheckTitle = function (sourceParagraph, index) {
		const docContent = this.getDocContent();
		const outlineParagraph = docContent.Content[index];
		if (outlineParagraph) {
			const pr = this.getPropertiesFromSourceShape(sourceParagraph);
			const info = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
			if (info && pr && pr.slide === info.slide) {
				if (!this.getSourceParagraph(outlineParagraph)) {
					this.removeParagraph(outlineParagraph, index);
				}
				this.addUpdatedParagraph(sourceParagraph, index, info);
				return true;
			}
		}
		return false;
	};

	OutlineView.prototype.updateParagraphData = function (updateData) {
		if (updateData.isNeedRecalculate()) {
			const presentation = this.getPresentation();
			presentation.updateSlideIndexes();
			this.updateExistingParagraphs(updateData.existingParagraphs);
			this.updateNewParagraphs(updateData.newParagraphs);
			return true;
		}
		return false;
	};
	OutlineView.prototype.updateSlideData = function (updateData) {
		if (updateData.isNeedRecalculate()) {
			const presentation = this.getPresentation();
			presentation.updateSlideIndexes();
			this.updateRemovedSlides(updateData.removedSlides);
			this.updateAddedSlides(updateData.addedSlides);
			return true;
		}
		return false;
	};
	OutlineView.prototype.update = function (savedContentsOnUpdate) {
		if (!this.isInit) {
			return;
		}
		if (this.dragDropManager) {
			this.dragDropManager.onOutlineUpdate();
		}
		let isRecalculate = this.recalcInfo.recalculateContent;

		const presentation = this.getPresentation();
		const paragraphUpdateDataCollector = new ParagraphUpdateDataCollector(this, this.mapToCheckParagraphs, savedContentsOnUpdate);
		const paragraphUpdateData = paragraphUpdateDataCollector.collect();
		isRecalculate |= this.updateParagraphData(paragraphUpdateData);
		this.resetMapToCheckParagraphs();

		const slideUpdateDataCollector = new SlideUpdateDataCollector(this, this.mapToCheckSlides, presentation.GetUsedSlidesMap());
		slideUpdateDataCollector.addSlidesToCheck(paragraphUpdateDataCollector.additionalCheckSlideMap);
		const updateData = slideUpdateDataCollector.collect();
		isRecalculate |= this.updateSlideData(updateData);
		this.resetMapToCheckSlides();

		if (isRecalculate) {
			this.recalcInfo.recalculateContent = false;
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
	UpdateExistingParagraphManager.prototype.update = function (existingOutlineParagraphs) {
		existingOutlineParagraphs.sort(function (aParagraph, bParagraph) {
			return bParagraph.Index - aParagraph.Index;
		});
		for (let i = 0; i < existingOutlineParagraphs.length; i += 1) {
			const outlineParagraph = existingOutlineParagraphs[i];
			this.checkRecalculateContentShapeIndexes(outlineParagraph);
			this.checkMoveSlideInfoToTitleParagraph(outlineParagraph);
			this.unlinkOutlineParagraph(outlineParagraph);

		}
	}
	UpdateExistingParagraphManager.prototype.checkRecalculateContentShapeIndexes = function (outlineParagraph) {
		const outlineView = this.outlineView;
		const sourceParagraph = outlineView.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			this.setRecalculateContentShapeIndexes(sourceParagraph);
		}
	};
	UpdateExistingParagraphManager.prototype.checkMoveSlideInfoToTitleParagraph = function (outlineParagraph) {
		this.outlineView.checkMoveSlideInfoToTitleParagraph(outlineParagraph);
	};
	UpdateExistingParagraphManager.prototype.unlinkOutlineParagraph = function (outlineParagraph) {
		const outlineView = this.outlineView;
		outlineView.unlinkOutlineParagraph(outlineParagraph);
	}
	UpdateExistingParagraphManager.prototype.setRecalculateContentShapeIndexes = function (sourceParagraph) {
		const outlineView = this.outlineView;
		const shape = sourceParagraph.GetParentShape();
		const content = sourceParagraph.GetParent();
		const slide = shape && shape.parent;
		if ((!this.useInDocumentManager.isUseInDocumentShape(shape) || content.IsEmpty()) && this.useInDocumentManager.isUseInDocumentSlide(slide)) {
			const info = outlineView.outlineInfo.getSlideInfoBySlide(slide);
			if (info) {
				info.resetContentShapeInfoMap();
			}
		}
	};
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
	OutlineView.prototype.checkSourceSlide = function (slide, isAdd) {
		if (!AscCommon.History.IsOn()) {
			return;
		}

		if (!this.mapToCheckSlides[slide.GetId()]) {
			this.mapToCheckSlides[slide.GetId()] = new CheckSlideInfo(slide);
		}
		const info = this.mapToCheckSlides[slide.GetId()];
		if (isAdd) {
			info.flags = info.flags | CHECK_SLIDE_INFO_ADD;
		} else {
			info.flags = info.flags | CHECK_SLIDE_INFO_REMOVE;
		}
		for (let i = 0; i < slide.cSld.spTree.length; i += 1) {
			const shape = slide.cSld.spTree[i];
			this.checkSourceShape(shape, true);
		}
	};
	OutlineView.prototype.checkSourceShape = function (shape, isFromSlide) {
		if (!AscCommon.History.IsOn()) {
			return;
		}
		if (shape.isOutlinePlaceholder()) {
			if (!isFromSlide) {
				const parent = shape.parent || shape.oldParent;
				if (parent) {
					this.mapToCheckSlides[parent.GetId()] = new CheckSlideInfo(parent);
				}
			}
			const content = shape.getDocContent();
			if (!content) {
				return;
			}
			for (let i = 0; i < content.Content.length; i += 1) {
				const paragraph = content.Content[i];
				this.checkParagraph(paragraph);
			}
		}
	};
	OutlineView.prototype.getApi = function () {
		return this.api;
	};
	OutlineView.prototype.selectAll = function (direction) {
		const content = this.getDocContent();
		if (content) {
			content.SelectAll(direction);
			this.updateInterfaceState();
			this.updateSelectionState();
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
				break;
			}
		}
		return res;
	}
	OutlineView.prototype.updateInterfaceState = function () {

	};
	OutlineView.prototype.getThumbnailsManager = function () {
		const api = this.getApi();
		return api.WordControl.Thumbnails;
	};
	OutlineView.prototype.updateSelectionState = function () {
		if (this.outlineShape) {
			if (this.isEmptyPresentation() || !this.isFocusOnOutline()) {
				this.outlineShape.resetSelectionState();
			} else {
				this.outlineShape.updateSelectionState();
			}
			const thumbnails = this.getThumbnailsManager();
			if (thumbnails) {
				thumbnails.OnUpdateOverlay();
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
		let returnValue = keydownresult_PreventNothing;
		switch (e.KeyCode) {
			case Asc.c_oAscKeyCodes.ArrowLeft: {
				if (isMacOs && e.CtrlKey) {
					this.moveCursorToStartOfLine(e.ShiftKey);
				} else {
					const bIsWord = isMacOs ? e.AltKey : e.CtrlKey;
					this.moveCursorLeft(e.ShiftKey, bIsWord);
				}
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowRight: {
				if (isMacOs && e.CtrlKey) {
					this.moveCursorToEndOfLine(e.ShiftKey);
				} else {
					const bIsWord = isMacOs ? e.AltKey : e.CtrlKey;
					this.moveCursorRight(e.ShiftKey, bIsWord);
				}
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowUp: {
				this.moveCursorUp(e.ShiftKey);
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.ArrowDown: {
				this.moveCursorDown(e.ShiftKey);
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.Home: {
				if (e.CtrlKey) {
					this.moveCursorToStartPos(e.ShiftKey);
				} else {
					this.moveCursorToStartOfLine(e.ShiftKey);
				}
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.End: {
				if (e.CtrlKey) {
					this.moveCursorToEndPos(e.ShiftKey);
				} else {
					this.moveCursorToEndOfLine(e.ShiftKey);
				}
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.Backspace:
			case Asc.c_oAscKeyCodes.Delete: {
				const oThis = this;
				const direction = e.KeyCode === Asc.c_oAscKeyCodes.Delete ? AscWord.Direction.FORWARD : AscWord.Direction.BACKWARD;
				this.checkSelectedObjectsAndCallback(function () {
					oThis.remove(false, direction);
				}, AscDFH.historydescription_Spreadsheet_Remove, true, true, true, direction);
				this.updateInterfaceState()
				this.updateSelectionState();
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.Tab: {
				this.paragraphIncDecIndentWithPoint(!e.ShiftKey);
				this.updateInterfaceState();
				this.updateSelectionState();
				returnValue = keydownresult_PreventAll;
				break;
			}
			case Asc.c_oAscKeyCodes.Enter: {
				presentation.StartAction();
				this.handleEnter(e);
				presentation.FinalizeAction();
				this.updateInterfaceState()
				this.updateSelectionState();
				returnValue = keydownresult_PreventAll;
				break;
			}
			default: {
				returnValue = null;
				break;
			}
		}
		return returnValue;
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
	OutlineView.prototype.getSlideLastParagraph = function (slideIndex) {
		const presentation = this.getPresentation();
		const docContent = this.getDocContent();
		const nextSlide = presentation.Slides[slideIndex + 1];
		if (nextSlide) {
			const nextInfo = this.outlineInfo.getSlideInfoBySlide(nextSlide);
			if (!nextInfo) {
				return null;
			}
			const nextSlideOutlineParagraph = nextInfo.outlineParagraph;
			return  docContent.Content[nextSlideOutlineParagraph.GetIndex() - 1];
		}
		return docContent.Content[docContent.Content.length - 1];

	};
	OutlineView.prototype.isFullSelectedSlide = function (slideIndex) {
		const presentation = this.getPresentation();
		const slide = presentation.Slides[slideIndex];
		if (!slide) {
			return false;
		}
		const info = this.outlineInfo.getSlideInfoBySlide(slide);
		if (!info) {
			return false;
		}
		const startOutlineParagraph = info.outlineParagraph;
		if (!startOutlineParagraph || !startOutlineParagraph.IsSelectedAll()) {
			return false;
		}

		let endOutlineParagraph = this.getSlideLastParagraph(slideIndex);
		if (endOutlineParagraph) {
			return endOutlineParagraph.IsSelectedAll();
		}
		return false;
	};
	OutlineView.prototype.isCursorAtBeginSlide = function (slideIndex) {
		const presentation = this.getPresentation();
		const slide = presentation.Slides[slideIndex];
		if (!slide) {
			return false;
		}
		const info = this.outlineInfo.getSlideInfoBySlide(slide);
		if (!info) {
			return false;
		}
		const startOutlineParagraph = info.outlineParagraph;
		return startOutlineParagraph.IsSelectionFromStart() || startOutlineParagraph.IsCursorAtBegin();
	};
	OutlineView.prototype.isCursorAtEndSlide = function (slideIndex) {
		let endOutlineParagraph = this.getSlideLastParagraph(slideIndex);
		if (!endOutlineParagraph) {
			return false;
		}
		return endOutlineParagraph.IsSelectionToEnd() || endOutlineParagraph.IsCursorAtEnd();
	};
	OutlineView.prototype.getSlidesForLock = function (isRemoveBeforeAdd, removeDirection) {
		const selectedArray = this.getSelectedSlideArray();
		if (selectedArray.length) {
			const presentation = this.getPresentation();
			let min = selectedArray[0];
			let max = selectedArray[selectedArray.length - 1];
			const docContent = this.getDocContent();
			const isSelectionUse = docContent.IsSelectionUse();

			if (isSelectionUse) {
				if (min === max) {
					if (this.isFullSelectedSlide(min)) {
						return [presentation.Slides[min]];
					}
				} else {
					const result = [];
					for (let i = min + 1; i <= max; i += 1) {
						result.push(presentation.Slides[i])
					}
					return result;
				}
			} else if (!isRemoveBeforeAdd) {
				if (removeDirection === AscWord.Direction.FORWARD) {
					if (presentation.Slides[max + 1] && this.isCursorAtEndSlide(max)) {
						return [presentation.Slides[max + 1]];
					}
				} else if (presentation.Slides[min - 1] && this.isCursorAtBeginSlide(min)) {
					return [presentation.Slides[min]];
				}
			}
		}

		return [];
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
		const paragraph = this.getTitleSlideParagraph(pageNum);
		if (paragraph) {
			this.moveCursorToOutlineIndex(paragraph.GetIndex());
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
			} else if (this.isSelectionUse()) {
				const beginPos = new ContentPosition(getSourceContentEdgePos(content, true));
				const endPos = new ContentPosition(getSourceContentEdgePos(content, false));
				if (direction === AscWord.Direction.BACKWARD) {
					this.setSavedPosition(new SavedPosition(endPos, beginPos));
				} else {
					this.setSavedPosition(new SavedPosition(beginPos, endPos));
				}
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
			} else if (direction === AscWord.Direction.BACKWARD) {
				if (!savedPosition.endPos) {
					savedPosition.endPos = new ContentPosition(getSourceContentEdgePos(content, true));
				}
			} else if (!savedPosition.startPos) {
				savedPosition.startPos = new ContentPosition(getSourceContentEdgePos(content, true));
			}
		} else if (idx === count - 1) {
			const savedPosition = this.getSavedPosition();
			if (content.IsSelectionUse()) {
				if (direction === AscWord.Direction.BACKWARD) {
					savedPosition.startPos = new ContentPosition(content.GetContentPosition(true, true));
				} else {
					savedPosition.endPos = new ContentPosition(content.GetContentPosition(true, false));
				}
			} else if (direction === AscWord.Direction.BACKWARD) {
				if (!savedPosition.startPos) {
					savedPosition.startPos = new ContentPosition(getSourceContentEdgePos(content, false));
				}
			} else if (!savedPosition.endPos) {
				savedPosition.endPos = new ContentPosition(getSourceContentEdgePos(content, false));
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
			if (outlineParagraph.IsSelectionUse()) {
				const beginPos = new SlidePosition(info.slide, true);
				const endPos = new SlidePosition(info.slide, false);
				if (direction === AscWord.Direction.BACKWARD) {
					this.setSavedPosition(new SavedPosition(endPos, beginPos));
				} else {
					this.setSavedPosition(new SavedPosition(beginPos, endPos));
				}
			} else {
				this.setSavedPosition(new SavedPosition(new SlidePosition(info.slide)));
			}
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
			} else if (direction === AscWord.Direction.BACKWARD) {
				if (!savedPosition.endPos) {
					savedPosition.endPos = new SlidePosition(info.slide, true);
				}
			} else if (!savedPosition.startPos) {
				savedPosition.startPos = new SlidePosition(info.slide, true);
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
			} else if (direction === AscWord.Direction.BACKWARD) {
				if (!savedPosition.startPos) {
					savedPosition.startPos = new SlidePosition(info.slide);
				}
			} else if (!savedPosition.endPos) {
				savedPosition.endPos = new SlidePosition(info.slide);
			}
		}
	};
	OutlineView.prototype.saveDocumentStateBeforeLoadChanges = function (docState) {
		docState = docState || {};
		const presentation = this.getPresentation();
		const docContent = this.getDocContent();
		if (!docContent || !presentation) {
			return docState;
		}
		const oThis = this;
		const compositePos = this.getSourcePosForCompositeRun();
		if (compositePos) {
			this.setSavedPosition(new SavedPosition(new ContentPosition(compositePos)));
		} else {
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
		}
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
		return docState;
	};

	OutlineView.prototype.applyDocumentStateAfterLoadChanges = function (docState) {
		this.update();
		const savedPosition = docState.SavedOutlinePosition;
		if (savedPosition) {
			this.setSavedPosition(savedPosition);
			this.applySavedPositionToOutline();
		}
	};
	OutlineView.prototype.getSlideInfoByTitleParagraph = function (outlineParagraph) {
		const slideInfo = this.outlineInfo.getSlideInfoByParagraph(outlineParagraph);
		if (slideInfo) {
			return slideInfo;
		}
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		if (!this.isTitleSourceParagraph(sourceParagraph)) {
			return null;
		}
		return this.outlineInfo.getSlideInfoBySlide(sourceParagraph.GetParentShape().parent);
	};
	OutlineView.prototype.moveTailTitleParagraphs = function (sourceParagraph, titleDocContent) {
		const sourceContent = sourceParagraph.GetParent();
		const startIndex = sourceParagraph.GetIndex() + 1;
		const count = sourceContent.Content.length - startIndex;
		if (count <= 0) {
			return;
		}
		for (let i = 0; i < count; i += 1) {
			const tailParagraph = sourceContent.Content[startIndex + i];
			titleDocContent.Add_ToContent(titleDocContent.Content.length, tailParagraph.Copy(titleDocContent));
		}
		sourceContent.Remove_FromContent(startIndex, count);
	};
	OutlineView.prototype.handleTitleEnter = function (outlineParagraph) {
		const slideInfo = this.getSlideInfoByTitleParagraph(outlineParagraph);
		if (!slideInfo) {
			return;
		}
		const sourceParagraph = this.getSourceParagraph(outlineParagraph);
		const outlineSlide = slideInfo.slide.getOutlineSlide();
		const shapesWithContent = outlineSlide.content.filter(function (element) {
			return element.isContentPlaceholder();
		});

		const presentation = this.getPresentation();
		const newSlide = presentation.addNextSlideAction();
		if (!sourceParagraph && shapesWithContent.length === 0) {
			new SavePosManager(this).saveSlideStartPos(newSlide);
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
			this.moveTailTitleParagraphs(sourceParagraph, titleDocContent);
		}

		if (shapesWithContent.length !== 0) {
			let contentShape = newSlide.cSld.getContentShape();
			if (!contentShape) {
				contentShape = newSlide.createContent();
			}
			const contentDocContent = contentShape.getDocContent();
			contentDocContent.Clear_Content();

			const firstShape = shapesWithContent[0];
			if (!firstShape.txBody || !firstShape.txBody.content) {
				firstShape.createTextBody();
			}
			let firstShapeContent = firstShape.getDocContent();
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
		let isNeedRecalculate = false;
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
						isNeedRecalculate = true;
					}
					oThis.setSavedRealContentState(content, selectProps.index, selectProps.count, selectProps.direction);
				} else if (!e.ShiftKey) {
					if (oShape.isGeneralTitle()) {
						const sourceParagraph = content.GetCurrentParagraph() || content.Content[0];
						const outlineParagraph = oThis.sourceToOutlineMap[sourceParagraph.GetId()];
						newSlide = oThis.handleTitleEnter(outlineParagraph);
					} else {

						content.AddNewParagraph();
						oThis.setSavedRealContentState(content, selectProps.index, selectProps.count, selectProps.direction);
					}
				}
			} else {
				newSlide = oThis.handleTitleEnter(selectProps.outlineParagraph);
			}
		});
		if (isNeedRecalculate) {
			presentation.Recalculate();
		}
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
		if (docContent && !this.isEmptyPresentation()) {
			selectionState.focus = true;
			selectionState.textObject = this.outlineShape;
			selectionState.selectStartPage = selectionState.textObject.selectStartPage;
			selectionState.textSelection = docContent.GetSelectionState();
		}
		return [selectionState];
	};
	OutlineView.prototype.getSlideIconSizes = function (outlineParagraph) {
		if (!this.isTitleOutlineParagraph(outlineParagraph)) {
			return null;
		}
		const slideX = DECORATORDRAWER_OFFSET_LEFT + this.getSlideNumbersWidth();
		const top = this.getParagraphY(outlineParagraph);
		const height = this.decorCacheManager.getMeasureInfo().height;
		return {x: slideX, y: top, width: DECORATORDRAWER_SLIDE_WIDTH, height: height};
	}
	OutlineView.prototype.getSelectedObjects = function () {
		const selectedObjects = [];
		this.forEachSelectedContent(function (selectProps) {
			if (selectProps.content) {
				const shape = selectProps.content.Is_DrawingShape(true);
				if (shape) {
					selectedObjects.push(shape);
				}
			}
		});
		return selectedObjects;
	};
	OutlineView.prototype.checkExistingContentSelectProps = function (selectProps) {
		if (!selectProps.content && selectProps.outlineParagraph) {
			return this.createTitleContent(selectProps.outlineParagraph);
		}
		return selectProps.content;
	};
	OutlineView.prototype.checkRemoveOnAddFromContentAndCallback = function (callback) {
		this.remove(true);
		const oThis = this;
		this.forEachSelectedContent(function (selectProps) {
			const content = oThis.checkExistingContentSelectProps(selectProps);
			if (content) {
				callback(content);
				oThis.setSavedRealContentState(content, selectProps.index, selectProps.count, selectProps.direction);
			}
			return true;
		});
		this.update();
		this.applySavedPositionToOutline();
	};

	OutlineView.prototype.getCompositeInputShape = function () {
		return this.compositeInputInfo ? this.compositeInputInfo.shape : null;
	};
	OutlineView.prototype.beginCompositeInput = function () {
		this.endCompositeInput();

		const docContent = this.getDocContent();
		if (!docContent || !docContent.Content.length) {
			return null;
		}

		let info = null;
		this.checkRemoveOnAddFromContentAndCallback(function (content) {
			if (info) {
				return;
			}
			const paragraph = content.Content[content.CurPos.ContentPos];
			if (!paragraph || !paragraph.IsParagraph()) {
				return;
			}
			const run = paragraph.Get_ElementByPos(paragraph.Get_ParaContentPos(false, false));
			if (!(run instanceof ParaRun)) {
				return;
			}
			info = {
				run: run,
				inRunPos: run.State.ContentPos,
				content: content,
				shape: content.Is_DrawingShape(true),
				outlineRun: null
			};
		});

		if (!info || !info.shape) {
			return null;
		}

		info.run.State.ContentPos = info.inRunPos;
		this.compositeInputInfo = info;
		return info;
	};
	OutlineView.prototype.getSourcePosForCompositeRun = function () {
		const info = this.compositeInputInfo;
		if (!info || !info.run) {
			return null;
		}
		const pos = info.run.GetDocumentPositionForCurrentPosition();
		if (!pos || pos.length < 3 || pos[1].Class !== info.run.GetParagraph()) {
			return null;
		}
		return pos;
	};
	OutlineView.prototype.getOutlineRunForCompositeRun = function () {
		const sourcePos = this.getSourcePosForCompositeRun();
		if (!sourcePos) {
			return null;
		}
		const outlinePos = this.rebuildSavedPositionPos(sourcePos);
		if (!outlinePos || !outlinePos.length) {
			return null;
		}
		const outlineCompositeRun = outlinePos[outlinePos.length - 1].Class;
		return outlineCompositeRun instanceof AscCommonWord.ParaRun ? outlineCompositeRun : null;
	};
	OutlineView.prototype.clearOutlineCompositeMarker = function () {
		const info = this.compositeInputInfo;
		if (info && info.outlineRun) {
			info.outlineRun.Set_CompositeInput(null);
			info.outlineRun = null;
		}
	};
	OutlineView.prototype.applyCompositeInputToOutline = function (compositeInput) {
		const sourcePos = this.getSourcePosForCompositeRun();
		if (!sourcePos) {
			return;
		}
		this.setSavedPosition(new SavedPosition(new ContentPosition(sourcePos)));
		this.applySavedPositionToOutline();

		this.clearOutlineCompositeMarker();
		if (!compositeInput || compositeInput.Length <= 0) {
			return;
		}
		const outlineRun = this.getOutlineRunForCompositeRun();
		if (!outlineRun) {
			return;
		}
		const marker = new AscWord.RunCompositeInput_Old(outlineRun);
		marker.Pos = compositeInput.Pos;
		marker.Length = compositeInput.Length;
		outlineRun.Set_CompositeInput(marker);
		this.compositeInputInfo.outlineRun = outlineRun;
	};
	OutlineView.prototype.checkCompositeInputExtends = function () {
		const info = this.compositeInputInfo;
		if (info && info.shape && info.shape.checkExtentsByDocContent) {
			info.shape.checkExtentsByDocContent(true, true);
		}
	};
	OutlineView.prototype.endCompositeInput = function () {
		this.clearOutlineCompositeMarker();
		this.compositeInputInfo = null;
	};

	OutlineView.prototype.applyDocContentFunction = function (docContentFunction, args) {
		const state = this.saveDocumentStateBeforeLoadChanges();
		this.forEachSelectedContent(function (selectProps) {
			const content = selectProps.content;
			if (content) {
				docContentFunction.apply(content, args);
				const shape = content.Is_DrawingShape(true);
				if (shape && shape.checkExtentsByDocContent) {
					shape.checkExtentsByDocContent();
				}
			}
		});
		this.applyDocumentStateAfterLoadChanges(state);
	};
	OutlineView.prototype.applyTextFunction = function (docContentFunction, args, isRemoveBeforeAdd) {
		const tableFunction = function () {};
		const oThis = this;
		let state = {};
		if (!isRemoveBeforeAdd) {
			this.saveDocumentStateBeforeLoadChanges(state);
		}

		this.forEachSelectedContent(function (selectProps) {
			let content = selectProps.content;
			if (isRemoveBeforeAdd) {
				content = oThis.checkExistingContentSelectProps(selectProps);
			}
			if (content) {
				const shape = content.Is_DrawingShape(true);
				if (shape) {
					shape.applyTextFunction(docContentFunction, tableFunction, args);
				}
				if (isRemoveBeforeAdd) {
					oThis.setSavedRealContentState(content, selectProps.index, selectProps.count, selectProps.direction);
				}
			}
		});

		if (isRemoveBeforeAdd) {
			this.update();
			this.applySavedPositionToOutline();
		} else {
			this.applyDocumentStateAfterLoadChanges(state);
		}
	};
	OutlineView.prototype.checkCurrentTextObjectExtends = function () {
		const selectedObjects = this.getSelectedObjects();
		for (let i = 0; i < selectedObjects.length; i += 1) {
			const selectedObject = selectedObjects[i];
			if (selectedObject && selectedObject.checkExtentsByDocContent) {
				selectedObject.checkExtentsByDocContent(true, true);
			}
		}
	};
	OutlineView.prototype.paragraphClearFormatting = function (isClearParaPr, isClearTextPr) {
		this.applyDocContentFunction(AscFormat.CDrawingDocContent.prototype.ClearParagraphFormatting, [isClearParaPr, isClearTextPr]);
	};
	OutlineView.prototype.paragraphIncDecFontSize = function (isIncrease) {
		this.applyDocContentFunction(CDocumentContent.prototype.IncreaseDecreaseFontSize, [isIncrease]);

	};
	OutlineView.prototype.paragraphIncDecIndent = function (isIncrease) {
		const docContent = this.getDocContent();
		if (!docContent) {
			return;
		}
		const manager = new ChangeOutlineLevelManager(this);
		const outlineParagraphs = manager.getSelectedOutlineParagraphs();
		if (!outlineParagraphs.length) {
			return;
		}
		if (isIncrease) {
			manager.increaseParagraphIndent(outlineParagraphs);
		} else {
			manager.decreaseParagraphIndent(outlineParagraphs);
		}
	};
	OutlineView.prototype.setDefaultTabSize = function (tabSize) {
		this.applyDocContentFunction(CDocumentContent.prototype.SetParagraphDefaultTabSize, [tabSize]);
	};
	OutlineView.prototype.setParagraphSpacing = function (spacing) {
		this.applyDocContentFunction(CDocumentContent.prototype.SetParagraphSpacing, [spacing]);
	};
	OutlineView.prototype.setParagraphBidi = function (isRtl) {
		this.applyDocContentFunction(CDocumentContent.prototype.SetParagraphBidi, [isRtl]);
	};
	OutlineView.prototype.setParagraphTabs = function (tabs) {
		this.applyTextFunction(CDocumentContent.prototype.SetParagraphTabs, [tabs]);
	};
	OutlineView.prototype.setParagraphNumbering = function (numInfo) {
		this.applyDocContentFunction(CDocumentContent.prototype.Set_ParagraphPresentationNumbering, [numInfo]);
	};

	OutlineView.prototype.setParagraphAlign = function (align) {
		function callback() {
			const info = new CSelectedElementsInfo();
			this.GetSelectedElementsInfo(info);
			const math = info.GetMath();
			if (null !== math && true !== math.Is_Inline()) {
				math.Set_Align(align);
				return;
			}
			this.SetParagraphAlign(align);
		}
		this.applyDocContentFunction(callback, []);
	};

	OutlineView.prototype.changeTextCase = function (caseType) {
		const oThis = this;
		const callback = function () {
			let paragraph;
			if (!this.IsSelectionUse()) {
				paragraph = this.GetCurrentParagraph();
				if (paragraph) {
					paragraph.SelectCurrentWord();
				}
			}
			if (this.IsSelectionUse() && !this.IsSelectionEmpty()) {
				const paragraphs = [];
				this.GetCurrentParagraph(false, paragraphs, {});

				const changeEngine = new AscCommonWord.CChangeTextCaseEngine(caseType);
				changeEngine.ProcessParagraphs(paragraphs);
			}
		}
		this.checkSelectedObjectsAndCallback(function () {
			oThis.applyDocContentFunction(callback, []);
		}, AscDFH.historydescription_Document_ChangeTextCase);

	};

	OutlineView.prototype.addTextWithPr = function (text, settings) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.remove(true);
			oThis.applyTextFunction(AscFormat.CDrawingDocContent.prototype.addTextWithPr, [text, settings], true);
		}, AscDFH.historydescription_Document_AddTextWithProperties, true, true);

	};
	OutlineView.prototype.setParagraphIndent = function (indent) {
		if(AscCommon.isRealObject(indent) && AscFormat.isRealNumber(indent.Left) && indent.Left < 0)
		{
			indent.Left = 0;
		}
		this.applyDocContentFunction(CDocumentContent.prototype.SetParagraphIndent, [indent]);
	};
	OutlineView.prototype.paragraphClearFormattingWithPoint = function (isClearParaPr, isClearTextPr) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.paragraphClearFormatting(isClearParaPr, isClearTextPr);
		}, AscDFH.historydescription_Presentation_ParagraphClearFormatting);
	};
	OutlineView.prototype.paragraphIncDecFontSizeWithPoint = function (isIncrease) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.paragraphIncDecFontSize(isIncrease);
		}, AscDFH.historydescription_Presentation_ParagraphIncDecFontSize);
	};
	OutlineView.prototype.paragraphIncDecIndentWithPoint = function (isIncrease) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.paragraphIncDecIndent(isIncrease);
		}, AscDFH.historydescription_Presentation_ParagraphIncDecIndent, true, true);
	};
	OutlineView.prototype.setParagraphSpacingWithPoint = function (spacing) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.setParagraphSpacing(spacing);
		}, AscDFH.historydescription_Presentation_SetParagraphSpacing);
	};
	OutlineView.prototype.setParagraphBidiWithPoint = function (isRtl) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.setParagraphBidi(isRtl);
		}, AscDFH.historydescription_Document_SetParagraphBidi);
	};
	OutlineView.prototype.setParagraphTabsWithPoint = function (tabs) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.setParagraphTabs(tabs);
		}, AscDFH.historydescription_Presentation_SetParagraphTabs);
	};
	OutlineView.prototype.setParagraphNumberingWithPoint = function (numInfo) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.setParagraphNumbering(numInfo);
		}, AscDFH.historydescription_Presentation_SetParagraphNumbering);
	};
	OutlineView.prototype.setParagraphAlignWithPoint = function (align) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.setParagraphAlign(align);
		}, AscDFH.historydescription_Presentation_SetParagraphAlign);
	};
	OutlineView.prototype.setParagraphIndentWithPoint = function (indent) {
		const oThis = this;
		this.checkSelectedObjectsAndCallback(function () {
			oThis.setParagraphIndent(indent);
		}, AscDFH.historydescription_Presentation_SetParagraphIndent);
	};



	OutlineView.prototype.checkSelectedObjectsAndCallback = function (callback, historyPointType, isNeedCheckDrawings, isNeedCheckSlides, isRemoveAction, removeDirection, additionalLockObjects) {
		const presentation = this.getPresentation();
		const pointType = AscFormat.isRealNumber(historyPointType) ? historyPointType : AscDFH.historydescription_CommonControllerCheckSelected;
		isNeedCheckDrawings = typeof isNeedCheckDrawings === "boolean" ? isNeedCheckDrawings : true;
		isNeedCheckSlides = typeof isNeedCheckSlides === "boolean" ? isNeedCheckSlides : false;
		let checkType = null;
		if (additionalLockObjects || (isNeedCheckDrawings && isNeedCheckSlides)) {
			checkType = AscCommon.changestype_OutlineSlideDrawingProps;
		} else if (isNeedCheckDrawings) {
			checkType = AscCommon.changestype_Drawing_Props;
		}

		const additionalData = {isRemoveAction: isRemoveAction, removeDirection: removeDirection, additionalLockObjects: additionalLockObjects};
		if(checkType === null || presentation.Document_Is_SelectionLocked(checkType, additionalData, undefined, null) === false) {
			presentation.StartAction(pointType);
			callback.call(this);
			presentation.FinalizeAction();
			return true;
		}
		return false;
	}
	OutlineView.prototype.getDragDropManager = function () {
		if (!this.dragDropManager) {
			this.dragDropManager = new AscCommonSlide.COutlineDragDropManager(this);
		}
		return this.dragDropManager;
	};
	OutlineView.prototype.isDragDropActive = function () {
		return !!this.dragDropManager && this.dragDropManager.isActive();
	};
	OutlineView.prototype.paraApplyCallback = function (props) {
		AscFormat.paraApplyCallback(props, this);
	};
	OutlineView.prototype.startRecalculate = function () {
		this.getPresentation().Recalculate();
	};
	OutlineView.prototype.isCanIncreaseParagraphLevel = function (isIncrease) {
		const manager = new ChangeOutlineLevelManager(this);
		const outlineParagraphs = manager.getSelectedOutlineParagraphs();
		if (!outlineParagraphs.length) {
			return false;
		}
		if (isIncrease) {
			return manager.isCanIncreaseParagraphIndent(outlineParagraphs);
		}
		return manager.isCanDecreaseParagraphIndent(outlineParagraphs);
	};
	OutlineView.prototype.refreshAllContentsFields = function () {
		const content = this.getDocContent();
		if (!content) {
			return;
		}
		const isOldUpdate = AscCommon.History.RecalculateData.Update;
		AscCommon.History.RecalculateData.Update = false;
		AscFormat.RefreshContentAllFields(content);
		AscCommon.History.RecalculateData.Update = isOldUpdate;
	};
	OutlineView.prototype.getSlideNumPresentationFieldInfo = function (outlineParagraph) {
		const presentation = this.getPresentation();
		const firstSlideNum = presentation.getFirstSlideNumber();
		const slideIndex = this.getSlideIndex(outlineParagraph);
		return {slideIndex: slideIndex, slideNum: firstSlideNum + slideIndex};
	};
	OutlineView.prototype.recalcFields = function () {
		this.refreshAllContentsFields();
		this.recalcInfo.recalculateContent = true;
	};


	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].OutlineSlide = OutlineSlide;
	window["AscCommonSlide"].OutlineView = OutlineView;
	window["AscCommonSlide"].DecoratorDrawer = DecoratorDrawer;
	window["AscCommonSlide"].OutlineSlidesSelectionManager = OutlineSlidesSelectionManager;
	window["AscCommonSlide"].SavedPosition = SavedPosition;
	window["AscCommonSlide"].ContentPosition = ContentPosition;
	window["AscCommonSlide"].SlidePosition = SlidePosition;
})();
