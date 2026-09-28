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

	const OutlineDragState = {
		none: 0,
		pending: 1,
		tracking: 2
	};
	const OutlineDragSelectionType = {
		text: 0,
		slides: 1,
		mixed: 2
	};
	const OutlineDropTargetKind = {
		none: 0,
		text: 1,
		boundary: 2
	};

	function getTextLength(paragraph) {
		return paragraph.GetCharOffsetByContentPos(paragraph.Get_EndPos(false));
	}

	function getParaContentPosByOffset(paragraph, offset) {
		if (offset <= 0) {
			return paragraph.Get_StartPos();
		}
		if (offset >= getTextLength(paragraph)) {
			return paragraph.Get_EndPos(false);
		}
		return paragraph.GetContentPosByCharOffset(offset);
	}

	function getDocPosByOffset(paragraph, offset, isBehindEnd) {
		const pos = [{Class: paragraph.GetParent(), Position: paragraph.GetIndex()}];
		if (isBehindEnd && offset >= getTextLength(paragraph)) {
			paragraph.GetEndContentPosition(pos, true);
			return pos;
		}
		return paragraph.UpdatePosArrayByContentPos(pos, getParaContentPosByOffset(paragraph, offset).Copy());
	}

	function getOffsetByDocPos(pos) {
		const paragraph = pos[1].Class;
		const paraContentPos = new AscWord.CParagraphContentPos();
		for (let i = 1; i < pos.length; i += 1) {
			paraContentPos.Update(pos[i].Position, i - 1);
		}
		return paragraph.GetCharOffsetByContentPos(paraContentPos);
	}

	function comparePoints(aIndex, aOffset, bIndex, bOffset) {
		if (aIndex !== bIndex) {
			return aIndex < bIndex ? -1 : 1;
		}
		if (aOffset !== bOffset) {
			return aOffset < bOffset ? -1 : 1;
		}
		return 0;
	}

	/**
	 * Selection of the outline normalized for drag'n'drop. All indexes are indexes of outline paragraphs,
	 * offsets are character offsets inside them. The start is always before the end in document order.
	 */
	function OutlineDragSelection(type, startIndex, startOffset, endIndex, endOffset, slides, isEndInTitle) {
		this.type = type;
		this.startIndex = startIndex;
		this.startOffset = startOffset;
		this.endIndex = endIndex;
		this.endOffset = endOffset;
		this.slides = slides;
		this.isEndInTitle = isEndInTitle;
	}
	OutlineDragSelection.prototype.isSame = function (other) {
		return !!other && this.type === other.type &&
			this.startIndex === other.startIndex && this.startOffset === other.startOffset &&
			this.endIndex === other.endIndex && this.endOffset === other.endOffset;
	};

	function OutlineDropTarget(kind) {
		this.kind = kind;
		this.selectionType = null;
		this.outlineIndex = -1;
		this.offset = 0;
		this.boundaryIndex = -1;
		this.isInsideSelection = false;
		this.x = 0;
		this.y = 0;
		this.height = 0;
	}
	OutlineDropTarget.prototype.isNone = function () {
		return this.kind === OutlineDropTargetKind.none;
	};

	/**
	 * Reads the outline model: which outline paragraphs are titles, which slide they belong to and so on.
	 */
	function OutlineModelReader(outlineView) {
		this.outlineView = outlineView;
	}
	OutlineModelReader.prototype.getDocContent = function () {
		return this.outlineView.getDocContent();
	};
	OutlineModelReader.prototype.getParagraphs = function () {
		const content = this.getDocContent();
		return content ? content.Content : [];
	};
	OutlineModelReader.prototype.getPresentation = function () {
		return this.outlineView.getPresentation();
	};
	OutlineModelReader.prototype.isTitle = function (outlineParagraph) {
		return !!this.outlineView.getSlideInfoByTitleParagraph(outlineParagraph);
	};
	OutlineModelReader.prototype.getSlide = function (outlineParagraph) {
		const slideIndex = this.outlineView.getSlideIndex(outlineParagraph);
		return slideIndex === -1 ? null : this.getPresentation().Slides[slideIndex] || null;
	};
	OutlineModelReader.prototype.getSlideParagraphs = function (slide) {
		return this.outlineView.getOutlineSlideParagraphs(slide);
	};
	OutlineModelReader.prototype.getTitleParagraphs = function (slide) {
		const paragraphs = this.getSlideParagraphs(slide);
		const result = [];
		for (let i = 0; i < paragraphs.length; i += 1) {
			if (!this.isTitle(paragraphs[i])) {
				break;
			}
			result.push(paragraphs[i]);
		}
		return result;
	};
	OutlineModelReader.prototype.getContentParagraphs = function (slide) {
		const paragraphs = this.getSlideParagraphs(slide);
		return paragraphs.slice(this.getTitleParagraphs(slide).length);
	};
	OutlineModelReader.prototype.getTitleRange = function (slide) {
		const titleParagraphs = this.getTitleParagraphs(slide);
		if (!titleParagraphs.length) {
			return null;
		}
		return {first: titleParagraphs[0].Index, last: titleParagraphs[titleParagraphs.length - 1].Index};
	};
	OutlineModelReader.prototype.getSlideLastIndex = function (slide) {
		const paragraphs = this.getSlideParagraphs(slide);
		return paragraphs.length ? paragraphs[paragraphs.length - 1].Index : -1;
	};
	OutlineModelReader.prototype.getSourceParagraph = function (outlineParagraph) {
		return this.outlineView.getSourceParagraph(outlineParagraph) || null;
	};
	OutlineModelReader.prototype.getParagraphBounds = function (outlineParagraph) {
		const pagesCount = outlineParagraph.Pages.length;
		if (!pagesCount) {
			return {top: outlineParagraph.Y, bottom: outlineParagraph.Y};
		}
		return {
			top: outlineParagraph.Get_PageBounds(0).Top,
			bottom: outlineParagraph.Get_PageBounds(pagesCount - 1).Bottom
		};
	};

	/**
	 * Classifies and normalizes the outline selection (FR-000 – FR-004a, FR-009a, FR-014b).
	 */
	function OutlineDragSelectionBuilder(outlineView) {
		this.reader = new OutlineModelReader(outlineView);
	}
	OutlineDragSelectionBuilder.prototype.getBounds = function () {
		const content = this.reader.getDocContent();
		if (!content || !content.IsSelectionUse() || content.IsSelectionEmpty()) {
			return null;
		}
		const firstPos = content.GetContentPosition(true, true);
		const secondPos = content.GetContentPosition(true, false);
		const first = {index: firstPos[0].Position, offset: getOffsetByDocPos(firstPos)};
		const second = {index: secondPos[0].Position, offset: getOffsetByDocPos(secondPos)};
		if (comparePoints(first.index, first.offset, second.index, second.offset) > 0) {
			return {start: second, end: first};
		}
		return {start: first, end: second};
	};
	OutlineDragSelectionBuilder.prototype.getLength = function (index) {
		return getTextLength(this.reader.getParagraphs()[index]);
	};
	OutlineDragSelectionBuilder.prototype.getSlidesBetween = function (startIndex, endIndex) {
		const paragraphs = this.reader.getParagraphs();
		const slides = [];
		for (let i = startIndex; i <= endIndex; i += 1) {
			const paragraph = paragraphs[i];
			if (this.reader.isTitle(paragraph)) {
				const slide = this.reader.getSlide(paragraph);
				if (slide && slides.indexOf(slide) === -1) {
					slides.push(slide);
				}
			}
		}
		return slides;
	};
	OutlineDragSelectionBuilder.prototype.hasTitleBetween = function (startIndex, endIndex) {
		const paragraphs = this.reader.getParagraphs();
		for (let i = startIndex; i <= endIndex; i += 1) {
			if (this.reader.isTitle(paragraphs[i])) {
				return true;
			}
		}
		return false;
	};
	OutlineDragSelectionBuilder.prototype.build = function () {
		const bounds = this.getBounds();
		if (!bounds) {
			return null;
		}
		const paragraphs = this.reader.getParagraphs();
		const start = bounds.start;
		let end = bounds.end;
		const startParagraph = paragraphs[start.index];
		const blockEnd = (end.offset === 0 && end.index > start.index) ?
			{index: end.index - 1, offset: this.getLength(end.index - 1)} : end;
		if (this.reader.isTitle(startParagraph)) {
			return this.buildFromTitle(start, end, blockEnd);
		}
		if (this.hasTitleBetween(start.index + 1, blockEnd.index)) {
			return this.buildMixed(start, blockEnd);
		}
		return new OutlineDragSelection(OutlineDragSelectionType.text, start.index, start.offset, end.index, end.offset, [], false);
	};
	OutlineDragSelectionBuilder.prototype.buildFromTitle = function (start, end, blockEnd) {
		const paragraphs = this.reader.getParagraphs();
		const startSlide = this.reader.getSlide(paragraphs[start.index]);
		const titleRange = this.reader.getTitleRange(startSlide);
		if (end.index <= titleRange.last) {
			const isWholeTitle = start.index === titleRange.first && start.offset === 0 &&
				end.index === titleRange.last && end.offset >= this.getLength(titleRange.last);
			if (!isWholeTitle) {
				return new OutlineDragSelection(OutlineDragSelectionType.text, start.index, start.offset, end.index, end.offset, [], false);
			}
			return this.buildSlides(titleRange.first, titleRange.last, this.getLength(titleRange.last), [startSlide], true);
		}
		end = blockEnd;
		if (end.index <= titleRange.last) {
			return this.buildSlides(titleRange.first, titleRange.last, this.getLength(titleRange.last), [startSlide], true);
		}
		const endParagraph = paragraphs[end.index];
		const endSlide = this.reader.getSlide(endParagraph);
		let isEndInTitle = false;
		if (this.reader.isTitle(endParagraph)) {
			const endTitleRange = this.reader.getTitleRange(endSlide);
			end = {index: endTitleRange.last, offset: this.getLength(endTitleRange.last)};
			isEndInTitle = true;
		}
		const slides = this.getSlidesBetween(titleRange.first, end.index);
		return this.buildSlides(titleRange.first, end.index, end.offset, slides, isEndInTitle);
	};
	OutlineDragSelectionBuilder.prototype.buildSlides = function (startIndex, endIndex, endOffset, slides, isEndInTitle) {
		const firstSlide = slides[0];
		const lastSlide = slides[slides.length - 1];
		if (firstSlide.num === 0) {
			endIndex = this.reader.getSlideLastIndex(lastSlide);
			endOffset = this.getLength(endIndex);
			isEndInTitle = this.reader.isTitle(this.reader.getParagraphs()[endIndex]);
		}
		return new OutlineDragSelection(OutlineDragSelectionType.slides, startIndex, 0, endIndex, endOffset, slides, isEndInTitle);
	};
	OutlineDragSelectionBuilder.prototype.buildMixed = function (start, end) {
		const paragraphs = this.reader.getParagraphs();
		const endParagraph = paragraphs[end.index];
		let isEndInTitle = false;
		if (this.reader.isTitle(endParagraph)) {
			const endTitleRange = this.reader.getTitleRange(this.reader.getSlide(endParagraph));
			end = {index: endTitleRange.last, offset: this.getLength(endTitleRange.last)};
			isEndInTitle = true;
		}
		const slides = this.getSlidesBetween(start.index + 1, end.index);
		return new OutlineDragSelection(OutlineDragSelectionType.mixed, start.index, start.offset, end.index, end.offset, slides, isEndInTitle);
	};

	/**
	 * Computes drop targets for a drag selection (FR-005, FR-012, FR-016).
	 */
	function OutlineDropTargetBuilder(outlineView, selection) {
		this.outlineView = outlineView;
		this.reader = new OutlineModelReader(outlineView);
		this.selection = selection;
	}
	OutlineDropTargetBuilder.prototype.isBlockSelection = function () {
		return this.selection.type === OutlineDragSelectionType.slides;
	};
	OutlineDropTargetBuilder.prototype.buildByPoint = function (x, y) {
		const content = this.reader.getDocContent();
		if (!content || !content.Content.length) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		const tx = this.outlineView.getInvertTransformX(x);
		const ty = this.outlineView.getInvertTransformY(y);
		const nearPos = content.Get_NearestPos(0, tx, ty);
		if (!nearPos || !nearPos.Paragraph) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		const paragraph = nearPos.Paragraph;
		if (this.isBlockSelection()) {
			const bounds = this.reader.getParagraphBounds(paragraph);
			const middle = (bounds.top + bounds.bottom) / 2;
			const boundaryIndex = ty < middle ? paragraph.Index : paragraph.Index + 1;
			return this.buildBoundary(boundaryIndex, ty);
		}
		return this.buildText(paragraph.Index, paragraph.GetCharOffsetByContentPos(nearPos.ContentPos), nearPos);
	};
	OutlineDropTargetBuilder.prototype.buildByOutlinePos = function (outlineIndex, offset) {
		const paragraph = this.reader.getParagraphs()[outlineIndex];
		if (!paragraph) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		if (this.isBlockSelection()) {
			return this.buildBoundary(offset > 0 ? outlineIndex + 1 : outlineIndex);
		}
		offset = Math.max(0, Math.min(offset, getTextLength(paragraph)));
		const nearPos = paragraph.Get_NearestPos(0, 0, 0);
		const target = this.buildText(outlineIndex, offset, null);
		if (nearPos) {
			target.height = nearPos.Height;
		}
		return target;
	};
	OutlineDropTargetBuilder.prototype.buildByBoundary = function (outlineIndex, isAfter) {
		if (!this.isBlockSelection()) {
			const paragraph = this.reader.getParagraphs()[outlineIndex];
			if (!paragraph) {
				return new OutlineDropTarget(OutlineDropTargetKind.none);
			}
			return this.buildByOutlinePos(outlineIndex, isAfter ? getTextLength(paragraph) : 0);
		}
		return this.buildBoundary(isAfter ? outlineIndex + 1 : outlineIndex);
	};
	OutlineDropTargetBuilder.prototype.buildText = function (outlineIndex, offset, nearPos) {
		const target = new OutlineDropTarget(OutlineDropTargetKind.text);
		target.selectionType = this.selection.type;
		target.outlineIndex = outlineIndex;
		target.offset = offset;
		if (nearPos) {
			target.x = nearPos.X;
			target.y = nearPos.Y;
			target.height = nearPos.Height;
		}
		target.isInsideSelection = this.isTextTargetInside(outlineIndex, offset);
		return target;
	};
	OutlineDropTargetBuilder.prototype.isTextTargetInside = function (outlineIndex, offset) {
		const selection = this.selection;
		const isAfterStart = comparePoints(outlineIndex, offset, selection.startIndex, selection.startOffset) >= 0;
		const isBeforeEnd = comparePoints(outlineIndex, offset, selection.endIndex, selection.endOffset) <= 0;
		if (isAfterStart && isBeforeEnd) {
			return true;
		}
		if (selection.type === OutlineDragSelectionType.mixed) {
			const paragraphs = this.reader.getParagraphs();
			const slide = this.reader.getSlide(paragraphs[outlineIndex]);
			const startSlide = this.reader.getSlide(paragraphs[selection.startIndex]);
			return slide === startSlide || selection.slides.indexOf(slide) !== -1;
		}
		return false;
	};
	OutlineDropTargetBuilder.prototype.snapBoundary = function (boundaryIndex, ty) {
		const paragraphs = this.reader.getParagraphs();
		if (boundaryIndex <= 0 || boundaryIndex >= paragraphs.length) {
			return Math.max(0, Math.min(boundaryIndex, paragraphs.length));
		}
		const previous = paragraphs[boundaryIndex - 1];
		const next = paragraphs[boundaryIndex];
		if (!this.reader.isTitle(previous) || !this.reader.isTitle(next)) {
			return boundaryIndex;
		}
		const slide = this.reader.getSlide(previous);
		if (slide !== this.reader.getSlide(next)) {
			return boundaryIndex;
		}
		const titleRange = this.reader.getTitleRange(slide);
		const before = titleRange.first;
		const after = titleRange.last + 1;
		if (typeof ty === "number") {
			const top = this.reader.getParagraphBounds(paragraphs[titleRange.first]).top;
			const bottom = this.reader.getParagraphBounds(paragraphs[titleRange.last]).bottom;
			return (ty - top) <= (bottom - ty) ? before : after;
		}
		return (boundaryIndex - before) <= (after - boundaryIndex) ? before : after;
	};
	OutlineDropTargetBuilder.prototype.buildBoundary = function (boundaryIndex, ty) {
		const paragraphs = this.reader.getParagraphs();
		boundaryIndex = this.snapBoundary(boundaryIndex, ty);
		const target = new OutlineDropTarget(OutlineDropTargetKind.boundary);
		target.selectionType = this.selection.type;
		target.boundaryIndex = boundaryIndex;
		target.x = 0;
		if (boundaryIndex < paragraphs.length) {
			target.y = this.reader.getParagraphBounds(paragraphs[boundaryIndex]).top;
		} else if (paragraphs.length) {
			target.y = this.reader.getParagraphBounds(paragraphs[paragraphs.length - 1]).bottom;
		}
		target.isInsideSelection = this.isBoundaryInside(boundaryIndex);
		return target;
	};
	OutlineDropTargetBuilder.prototype.isBoundaryInside = function (boundaryIndex) {
		const selection = this.selection;
		const endLength = getTextLength(this.reader.getParagraphs()[selection.endIndex]);
		const lastBoundary = selection.endOffset >= endLength ? selection.endIndex + 1 : selection.endIndex;
		return boundaryIndex >= selection.startIndex && boundaryIndex <= lastBoundary;
	};

	function TrackedPosition(paragraph, offset) {
		this.paragraph = paragraph;
		this.offset = offset;
	}

	/**
	 * Low level operations on slide paragraphs used by the drop plans: split, copy/move, join.
	 * Every operation goes to the slides, the outline is refreshed afterwards.
	 */
	function OutlineParagraphOperations(outlineView) {
		this.outlineView = outlineView;
		this.aliases = {};
		this.tracked = [];
		this.touchedContents = {};
	}
	OutlineParagraphOperations.prototype.touch = function (content) {
		if (content) {
			this.touchedContents[content.GetId()] = content;
		}
	};
	OutlineParagraphOperations.prototype.track = function (paragraph, offset) {
		const position = new TrackedPosition(paragraph, offset);
		this.tracked.push(position);
		return position;
	};
	/**
	 * Splits a paragraph at several character offsets. Returns the segments from left to right,
	 * the first segment is the original paragraph.
	 */
	OutlineParagraphOperations.prototype.split = function (paragraph, offsets) {
		const sortedOffsets = offsets.slice().sort(function (a, b) {
			return a - b;
		}).filter(function (offset, index, array) {
			return index === 0 || array[index - 1] !== offset;
		});
		const content = paragraph.GetParent();
		const rightSegments = [];
		for (let i = sortedOffsets.length - 1; i >= 0; i -= 1) {
			const rightParagraph = new AscWord.Paragraph(content, true);
			paragraph.Split(rightParagraph, getParaContentPosByOffset(paragraph, sortedOffsets[i]));
			content.Add_ToContent(paragraph.GetIndex() + 1, rightParagraph);
			rightSegments.unshift(rightParagraph);
		}
		this.touch(content);
		return {offsets: sortedOffsets, segments: [paragraph].concat(rightSegments)};
	};
	OutlineParagraphOperations.prototype.copyParagraphs = function (paragraphs) {
		const copies = [];
		for (let i = 0; i < paragraphs.length; i += 1) {
			copies.push(paragraphs[i].Copy2(paragraphs[i].GetParent()));
		}
		return copies;
	};
	OutlineParagraphOperations.prototype.removeParagraphs = function (paragraphs) {
		for (let i = paragraphs.length - 1; i >= 0; i -= 1) {
			const paragraph = paragraphs[i];
			const content = paragraph.GetParent();
			const index = paragraph.GetIndex();
			if (content && index !== -1 && content.Content[index] === paragraph) {
				content.Remove_FromContent(index, 1);
				this.touch(content);
			}
		}
	};
	OutlineParagraphOperations.prototype.insertParagraphs = function (content, index, paragraphs) {
		for (let i = 0; i < paragraphs.length; i += 1) {
			content.Add_ToContent(index + i, paragraphs[i]);
		}
		this.touch(content);
	};
	OutlineParagraphOperations.prototype.appendParagraphs = function (content, paragraphs) {
		if (!paragraphs.length) {
			return;
		}
		if (content.Is_Empty()) {
			content.ClearContent(false);
		}
		this.insertParagraphs(content, content.Content.length, paragraphs);
	};
	OutlineParagraphOperations.prototype.resolve = function (paragraph) {
		let result = paragraph;
		while (result && this.aliases[result.GetId()]) {
			result = this.aliases[result.GetId()];
		}
		return result;
	};
	/**
	 * Appends the text of the right paragraph to the left one and removes the right paragraph.
	 * The joined paragraph keeps the properties of the left one (FR-018).
	 */
	OutlineParagraphOperations.prototype.join = function (leftParagraph, rightParagraph) {
		const left = this.resolve(leftParagraph);
		const right = this.resolve(rightParagraph);
		if (!left || !right || left === right) {
			return left;
		}
		const leftLength = getTextLength(left);
		left.Concat(right);
		const rightContent = right.GetParent();
		const rightIndex = right.GetIndex();
		if (rightContent && rightIndex !== -1 && rightContent.Content[rightIndex] === right) {
			rightContent.Remove_FromContent(rightIndex, 1);
			this.touch(rightContent);
		}
		this.touch(left.GetParent());
		this.aliases[right.GetId()] = left;
		for (let i = 0; i < this.tracked.length; i += 1) {
			const position = this.tracked[i];
			if (position.paragraph === right) {
				position.paragraph = left;
				position.offset += leftLength;
			}
		}
		return left;
	};
	OutlineParagraphOperations.prototype.setLevel = function (paragraphs, level) {
		for (let i = 0; i < paragraphs.length; i += 1) {
			paragraphs[i].Set_PresentationLevel(level);
		}
	};
	OutlineParagraphOperations.prototype.isTitleContent = function (content) {
		const shape = content && content.Is_DrawingShape(true);
		return !!(shape && shape.isGeneralTitle && shape.isGeneralTitle());
	};
	/**
	 * The content that receives paragraphs added "to the end of a slide" (FR-009c).
	 */
	OutlineParagraphOperations.prototype.getReceiverContent = function (slide) {
		const contentShapes = slide.getOutlineSlide().content;
		for (let i = contentShapes.length - 1; i >= 0; i -= 1) {
			const content = contentShapes[i].getDocContent();
			if (content && !content.Is_Empty()) {
				return content;
			}
		}
		const firstShape = contentShapes[0];
		if (firstShape) {
			if (!firstShape.txBody || !firstShape.txBody.content) {
				firstShape.createTextBody();
			}
			return firstShape.getDocContent();
		}
		return this.outlineView.createSubtitleContent(slide);
	};
	OutlineParagraphOperations.prototype.checkShapesExtents = function () {
		for (let id in this.touchedContents) {
			const content = this.touchedContents[id];
			const shape = content.Is_DrawingShape(true);
			if (shape && shape.checkExtentsByDocContent) {
				shape.checkExtentsByDocContent();
			}
		}
	};

	/**
	 * Maps paragraphs of a slide to the same paragraphs of its duplicate (copy mode).
	 */
	function mapParagraphToSlideCopy(paragraph, slide, slideCopy) {
		const shape = paragraph.GetParentShape();
		const shapeIndex = slide.cSld.spTree.indexOf(shape);
		const shapeCopy = slideCopy.cSld.spTree[shapeIndex];
		const content = shapeCopy && shapeCopy.getDocContent();
		return content ? content.Content[paragraph.GetIndex()] || null : null;
	}

	/**
	 * Executes one drop: split → move → join (research R2). Built once per drop, from the current outline state.
	 */
	function OutlineDragDropExecutor(outlineView, selection, target, isCopy) {
		this.outlineView = outlineView;
		this.reader = new OutlineModelReader(outlineView);
		this.selection = selection;
		this.target = target;
		this.isCopy = !!isCopy;
		this.operations = new OutlineParagraphOperations(outlineView);
		this.resultStart = null;
		this.resultEnd = null;
		this.resultSlide = null;
	}
	OutlineDragDropExecutor.prototype.getPresentation = function () {
		return this.outlineView.getPresentation();
	};
	OutlineDragDropExecutor.prototype.getOutlineParagraph = function (index) {
		return this.reader.getParagraphs()[index];
	};
	OutlineDragDropExecutor.prototype.getSourceParagraphs = function (outlineParagraphs) {
		const result = [];
		for (let i = 0; i < outlineParagraphs.length; i += 1) {
			const sourceParagraph = this.reader.getSourceParagraph(outlineParagraphs[i]);
			if (sourceParagraph) {
				result.push(sourceParagraph);
			}
		}
		return result;
	};
	OutlineDragDropExecutor.prototype.getOrCreateSourceParagraph = function (outlineIndex) {
		const outlineParagraph = this.getOutlineParagraph(outlineIndex);
		const sourceParagraph = this.reader.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			return sourceParagraph;
		}
		const titleContent = this.outlineView.createTitleContent(outlineParagraph);
		return titleContent ? titleContent.Content[0] : null;
	};
	OutlineDragDropExecutor.prototype.getLockObjects = function () {
		const slides = [];
		const shapes = [];
		function addSlideShapes(slide) {
			if (!slide) {
				return;
			}
			const outlineShapes = slide.getOutlineSlide().getOutlineShapes();
			for (let i = 0; i < outlineShapes.length; i += 1) {
				if (outlineShapes[i] && shapes.indexOf(outlineShapes[i]) === -1) {
					shapes.push(outlineShapes[i]);
				}
			}
		}
		const presentation = this.getPresentation();
		const selection = this.selection;
		if (!this.isCopy) {
			for (let i = 0; i < selection.slides.length; i += 1) {
				slides.push(selection.slides[i]);
			}
			if (selection.type !== OutlineDragSelectionType.text) {
				const firstSlide = selection.slides[0];
				addSlideShapes(presentation.Slides[presentation.GetSlideIndex(firstSlide) - 1]);
				addSlideShapes(this.reader.getSlide(this.getOutlineParagraph(selection.startIndex)));
			}
		}
		if (this.target.kind === OutlineDropTargetKind.text) {
			addSlideShapes(this.reader.getSlide(this.getOutlineParagraph(this.target.outlineIndex)));
		} else {
			const boundary = this.getBoundarySlide();
			addSlideShapes(boundary.slide);
		}
		return {slides: slides, shapes: shapes, isOnlyAdditional: this.isCopy};
	};
	OutlineDragDropExecutor.prototype.execute = function () {
		switch (this.selection.type) {
			case OutlineDragSelectionType.text: {
				return this.executeText();
			}
			case OutlineDragSelectionType.slides: {
				return this.executeSlides();
			}
			case OutlineDragSelectionType.mixed: {
				return this.executeMixed();
			}
		}
		return false;
	};
	OutlineDragDropExecutor.prototype.finish = function () {
		const outlineView = this.outlineView;
		const presentation = this.getPresentation();
		this.operations.checkShapesExtents();
		presentation.updateSlideIndexes();
		const startPosition = this.resultStart;
		const endPosition = this.resultEnd;
		if (startPosition && endPosition) {
			outlineView.setSavedPosition(new AscCommonSlide.SavedPosition(startPosition, endPosition));
		}
		outlineView.update();
		if (this.resultSlide) {
			const slideIndex = presentation.GetSlideIndex(this.resultSlide);
			if (slideIndex !== -1) {
				presentation.DrawingDocument.m_oWordControl.GoToPage(slideIndex);
			}
		}
		outlineView.applySavedPositionToOutline();
	};
	OutlineDragDropExecutor.prototype.getTrackedContentPosition = function (trackedPosition, isEnd) {
		const paragraph = this.operations.resolve(trackedPosition.paragraph);
		return new AscCommonSlide.ContentPosition(getDocPosByOffset(paragraph, trackedPosition.offset, isEnd));
	};
	OutlineDragDropExecutor.prototype.getSplitSegments = function (splitResults, paragraph) {
		return splitResults[paragraph.GetId()];
	};
	OutlineDragDropExecutor.prototype.splitAtPoints = function (points) {
		const groups = {};
		const order = [];
		for (let i = 0; i < points.length; i += 1) {
			const point = points[i];
			const id = point.paragraph.GetId();
			if (!groups[id]) {
				groups[id] = {paragraph: point.paragraph, offsets: []};
				order.push(id);
			}
			groups[id].offsets.push(point.offset);
		}
		const results = {};
		for (let i = 0; i < order.length; i += 1) {
			const group = groups[order[i]];
			results[order[i]] = this.operations.split(group.paragraph, group.offsets);
		}
		return results;
	};
	function getSegmentStartingAt(splitResult, offset) {
		return splitResult.segments[splitResult.offsets.indexOf(offset) + 1];
	}
	function getSegmentEndingAt(splitResult, offset) {
		return splitResult.segments[splitResult.offsets.indexOf(offset)];
	}
	/**
	 * Source paragraphs of a text range of the outline, in document order (the whole paragraphs between the ends).
	 */
	OutlineDragDropExecutor.prototype.getRangeSourceParagraphs = function (startIndex, endIndex) {
		const paragraphs = this.reader.getParagraphs();
		return this.getSourceParagraphs(paragraphs.slice(startIndex, endIndex + 1));
	};

	OutlineDragDropExecutor.prototype.executeText = function () {
		const selection = this.selection;
		const operations = this.operations;
		const sourceParagraphs = this.getRangeSourceParagraphs(selection.startIndex, selection.endIndex);
		if (!sourceParagraphs.length) {
			return false;
		}
		const targetParagraph = this.getOrCreateSourceParagraph(this.target.outlineIndex);
		if (!targetParagraph) {
			return false;
		}
		const startParagraph = sourceParagraphs[0];
		const endParagraph = sourceParagraphs[sourceParagraphs.length - 1];
		const isTargetTitle = operations.isTitleContent(targetParagraph.GetParent());
		const isSourceTitle = operations.isTitleContent(startParagraph.GetParent());
		if (this.isCopy) {
			return this.executeTextCopy(sourceParagraphs, targetParagraph, isTargetTitle !== isSourceTitle);
		}

		const points = [
			{paragraph: startParagraph, offset: selection.startOffset},
			{paragraph: endParagraph, offset: selection.endOffset},
			{paragraph: targetParagraph, offset: this.target.offset}
		];
		const splitResults = this.splitAtPoints(points);
		const startSplit = this.getSplitSegments(splitResults, startParagraph);
		const endSplit = this.getSplitSegments(splitResults, endParagraph);
		const targetSplit = this.getSplitSegments(splitResults, targetParagraph);

		const sourceLeft = getSegmentEndingAt(startSplit, selection.startOffset);
		const sourceRight = getSegmentStartingAt(endSplit, selection.endOffset);
		const targetLeft = getSegmentEndingAt(targetSplit, this.target.offset);
		const targetRight = getSegmentStartingAt(targetSplit, this.target.offset);

		const pieces = [];
		const firstPiece = getSegmentStartingAt(startSplit, selection.startOffset);
		const lastPiece = getSegmentEndingAt(endSplit, selection.endOffset);
		pieces.push(firstPiece);
		for (let i = 1; i < sourceParagraphs.length - 1; i += 1) {
			pieces.push(sourceParagraphs[i]);
		}
		if (lastPiece !== firstPiece) {
			pieces.push(lastPiece);
		}

		const copies = operations.copyParagraphs(pieces);
		if (!this.isCopy) {
			operations.removeParagraphs(pieces);
		}
		const targetContent = targetLeft.GetParent();
		operations.insertParagraphs(targetContent, targetLeft.GetIndex() + 1, copies);
		if (isTargetTitle !== isSourceTitle) {
			operations.setLevel(copies, 0);
		}

		this.resultStart = operations.track(copies[0], 0);
		const lastCopy = copies[copies.length - 1];
		this.resultEnd = operations.track(lastCopy, getTextLength(lastCopy));

		operations.join(targetLeft, copies[0]);
		operations.join(lastCopy, targetRight);
		if (!this.isCopy) {
			this.joinSourceEdges(sourceLeft, sourceRight);
		}
		this.resultStart = this.getTrackedContentPosition(this.resultStart, false);
		this.resultEnd = this.getTrackedContentPosition(this.resultEnd, false);
		this.resultSlide = targetContent.Is_DrawingShape(true).parent;
		this.finish();
		return true;
	};
	OutlineDragDropExecutor.prototype.executeTextCopy = function (sourceParagraphs, targetParagraph, isLevelReset) {
		const selection = this.selection;
		const operations = this.operations;
		const copiesSource = operations.copyParagraphs(sourceParagraphs);
		const targetSplit = operations.split(targetParagraph, [this.target.offset]);
		const targetLeft = targetSplit.segments[0];
		const targetRight = targetSplit.segments[1];
		const content = targetLeft.GetParent();
		operations.insertParagraphs(content, targetLeft.GetIndex() + 1, copiesSource);
		const pieces = this.trimCopies(copiesSource, selection.startOffset, selection.endOffset);
		if (isLevelReset) {
			operations.setLevel(pieces, 0);
		}
		const lastPiece = pieces[pieces.length - 1];
		this.resultStart = operations.track(pieces[0], 0);
		this.resultEnd = operations.track(lastPiece, getTextLength(lastPiece));
		operations.join(targetLeft, pieces[0]);
		operations.join(lastPiece, targetRight);
		this.resultStart = this.getTrackedContentPosition(this.resultStart, false);
		this.resultEnd = this.getTrackedContentPosition(this.resultEnd, false);
		this.resultSlide = content.Is_DrawingShape(true).parent;
		this.finish();
		return true;
	};
	OutlineDragDropExecutor.prototype.trimCopies = function (copies, startOffset, endOffset) {
		const operations = this.operations;
		const firstCopy = copies[0];
		const lastCopy = copies[copies.length - 1];
		const extraSegments = [];
		let firstPiece;
		let lastPiece;
		if (firstCopy === lastCopy) {
			const split = operations.split(firstCopy, typeof endOffset === "number" ? [startOffset, endOffset] : [startOffset]);
			firstPiece = getSegmentStartingAt(split, startOffset);
			lastPiece = firstPiece;
			extraSegments.push(getSegmentEndingAt(split, startOffset));
			if (typeof endOffset === "number") {
				lastPiece = getSegmentEndingAt(split, endOffset);
				extraSegments.push(getSegmentStartingAt(split, endOffset));
			}
		} else {
			lastPiece = lastCopy;
			if (typeof endOffset === "number") {
				const endSplit = operations.split(lastCopy, [endOffset]);
				lastPiece = endSplit.segments[0];
				extraSegments.push(endSplit.segments[1]);
			}
			const startSplit = operations.split(firstCopy, [startOffset]);
			firstPiece = startSplit.segments[1];
			extraSegments.push(startSplit.segments[0]);
		}
		operations.removeParagraphs(extraSegments);
		const pieces = [firstPiece].concat(copies.slice(1, copies.length - 1));
		if (lastPiece !== firstPiece) {
			pieces.push(lastPiece);
		}
		return pieces;
	};
	OutlineDragDropExecutor.prototype.joinSourceEdges = function (sourceLeft, sourceRight) {
		const operations = this.operations;
		const left = operations.resolve(sourceLeft);
		const right = operations.resolve(sourceRight);
		if (left === right) {
			return;
		}
		const isLeftEmpty = getTextLength(left) === 0;
		const isRightEmpty = getTextLength(right) === 0;
		if (isLeftEmpty && !isRightEmpty) {
			operations.removeParagraphs([left]);
		} else if (isRightEmpty && !isLeftEmpty) {
			operations.removeParagraphs([right]);
		} else {
			operations.join(left, right);
		}
	};

	/**
	 * Resolves a boundary target to the slide after whose paragraphs the boundary lies.
	 * slide === null means "before the first slide".
	 */
	OutlineDragDropExecutor.prototype.getBoundarySlide = function () {
		const paragraphs = this.reader.getParagraphs();
		const boundaryIndex = this.target.boundaryIndex;
		if (boundaryIndex <= 0) {
			return {slide: null, contentIndex: 0};
		}
		const previous = paragraphs[boundaryIndex - 1];
		const slide = this.reader.getSlide(previous);
		const slideParagraphs = this.reader.getSlideParagraphs(slide);
		const titleCount = this.reader.getTitleParagraphs(slide).length;
		const count = previous.Index - slideParagraphs[0].Index + 1;
		return {slide: slide, contentIndex: Math.max(0, count - titleCount)};
	};
	OutlineDragDropExecutor.prototype.getSlideContentSources = function (slide) {
		return this.getSourceParagraphs(this.reader.getContentParagraphs(slide));
	};
	OutlineDragDropExecutor.prototype.getSlideTitleSources = function (slide) {
		return this.getSourceParagraphs(this.reader.getTitleParagraphs(slide));
	};
	OutlineDragDropExecutor.prototype.moveSlides = function (slides, afterSlide) {
		const presentation = this.getPresentation();
		const indexes = [];
		for (let i = 0; i < slides.length; i += 1) {
			indexes.push(presentation.GetSlideIndex(slides[i]));
		}
		indexes.sort(function (a, b) {
			return a - b;
		});
		let position = afterSlide ? presentation.GetSlideIndex(afterSlide) + 1 : 0;
		for (let i = 0; i < indexes.length; i += 1) {
			if (indexes[i] < position) {
				position -= 1;
			}
		}
		if (position !== indexes[0]) {
			presentation.moveSlides(indexes, position);
		}
		presentation.updateSlideIndexes();
		return slides;
	};
	OutlineDragDropExecutor.prototype.copySlides = function (slides, afterSlide) {
		const presentation = this.getPresentation();
		let position = afterSlide ? presentation.GetSlideIndex(afterSlide) + 1 : 0;
		const copies = [];
		for (let i = 0; i < slides.length; i += 1) {
			const idMap = {};
			const slideCopy = slides[i].createDuplicate(idMap, false);
			AscFormat.fResetConnectorsIds(slideCopy.cSld.spTree, idMap);
			presentation.insertSlideObjectToPos(position + i, slideCopy);
			copies.push(slideCopy);
		}
		presentation.updateSlideIndexes();
		return copies;
	};
	/**
	 * Splits the last selected paragraph of a block at the selection end and returns
	 * {lastSelected, leftover}: the last selected source paragraph and the not selected paragraphs of the slide.
	 */
	OutlineDragDropExecutor.prototype.splitBlockEnd = function (lastSlide) {
		const selection = this.selection;
		const operations = this.operations;
		const contentSources = this.getSlideContentSources(lastSlide);
		if (selection.isEndInTitle) {
			const titleSources = this.getSlideTitleSources(lastSlide);
			return {lastSelected: titleSources[titleSources.length - 1] || null, leftover: contentSources, selectedWholeCount: 0};
		}
		const endOutlineParagraph = this.getOutlineParagraph(selection.endIndex);
		const endSource = this.reader.getSourceParagraph(endOutlineParagraph);
		const endIndexInContent = contentSources.indexOf(endSource);
		if (endIndexInContent === -1) {
			return {lastSelected: endSource, leftover: [], selectedWholeCount: contentSources.length};
		}
		const leftover = contentSources.slice(endIndexInContent + 1);
		let selectedWholeCount = endIndexInContent + 1;
		if (selection.endOffset < getTextLength(endSource)) {
			const splitResult = operations.split(endSource, [selection.endOffset]);
			leftover.unshift(splitResult.segments[1]);
			selectedWholeCount = endIndexInContent;
		}
		return {lastSelected: endSource, leftover: leftover, selectedWholeCount: selectedWholeCount};
	};
	OutlineDragDropExecutor.prototype.moveParagraphsToSlideEnd = function (paragraphs, slide) {
		if (!paragraphs.length) {
			return [];
		}
		const operations = this.operations;
		const copies = operations.copyParagraphs(paragraphs);
		operations.removeParagraphs(paragraphs);
		operations.appendParagraphs(operations.getReceiverContent(slide), copies);
		return copies;
	};

	OutlineDragDropExecutor.prototype.executeSlides = function () {
		const presentation = this.getPresentation();
		const selection = this.selection;
		const slides = selection.slides;
		const firstSlide = slides[0];
		const lastSlide = slides[slides.length - 1];
		const receiverSlide = presentation.Slides[presentation.GetSlideIndex(firstSlide) - 1] || null;
		const boundary = this.getBoundarySlide();
		const targetSlide = boundary.slide;

		let targetTail = [];
		if (targetSlide && (this.isCopy || targetSlide !== lastSlide)) {
			targetTail = this.getSlideContentSources(targetSlide).slice(boundary.contentIndex);
		}

		if (this.isCopy) {
			return this.executeSlidesCopy(targetSlide, targetTail);
		}

		const blockEnd = this.splitBlockEnd(lastSlide);
		let leftoverToReceiver = blockEnd.leftover;
		let leftoverToEnd = [];
		let afterSlide = targetSlide;
		if (targetSlide === lastSlide) {
			const splitIndex = Math.max(0, boundary.contentIndex - blockEnd.selectedWholeCount);
			leftoverToReceiver = blockEnd.leftover.slice(0, splitIndex);
			leftoverToEnd = blockEnd.leftover.slice(splitIndex);
			afterSlide = receiverSlide;
		}

		if (receiverSlide) {
			this.moveParagraphsToSlideEnd(leftoverToReceiver, receiverSlide);
		}
		this.moveSlides(slides, afterSlide);
		this.moveParagraphsToSlideEnd(leftoverToEnd.concat(targetTail), lastSlide);

		this.setBlockResult(firstSlide, lastSlide, blockEnd.lastSelected);
		this.finish();
		return true;
	};
	OutlineDragDropExecutor.prototype.executeSlidesCopy = function (targetSlide, targetTail) {
		const selection = this.selection;
		const slides = selection.slides;
		const lastSlide = slides[slides.length - 1];
		const copies = this.copySlides(slides, targetSlide);
		const lastCopy = copies[copies.length - 1];
		const lastSelected = this.cutCopyLeftover(lastSlide, lastCopy);
		this.moveParagraphsToSlideEnd(targetTail, lastCopy);
		this.setBlockResult(copies[0], lastCopy, lastSelected);
		this.finish();
		return true;
	};
	OutlineDragDropExecutor.prototype.cutCopyLeftover = function (lastSlide, lastCopy) {
		const selection = this.selection;
		const operations = this.operations;
		const contentSources = this.getSlideContentSources(lastSlide);
		if (selection.isEndInTitle) {
			const copiesToRemove = [];
			for (let i = 0; i < contentSources.length; i += 1) {
				const paragraph = mapParagraphToSlideCopy(contentSources[i], lastSlide, lastCopy);
				if (paragraph) {
					copiesToRemove.push(paragraph);
				}
			}
			operations.removeParagraphs(copiesToRemove);
			const titleSources = this.getSlideTitleSources(lastSlide);
			const lastTitle = titleSources[titleSources.length - 1];
			return lastTitle ? mapParagraphToSlideCopy(lastTitle, lastSlide, lastCopy) : null;
		}
		const endSource = this.reader.getSourceParagraph(this.getOutlineParagraph(selection.endIndex));
		const endIndexInContent = contentSources.indexOf(endSource);
		const endCopy = mapParagraphToSlideCopy(endSource, lastSlide, lastCopy);
		if (endIndexInContent === -1 || !endCopy) {
			return endCopy;
		}
		const leftoverCopies = [];
		for (let i = endIndexInContent + 1; i < contentSources.length; i += 1) {
			const paragraph = mapParagraphToSlideCopy(contentSources[i], lastSlide, lastCopy);
			if (paragraph) {
				leftoverCopies.push(paragraph);
			}
		}
		if (selection.endOffset < getTextLength(endCopy)) {
			leftoverCopies.unshift(operations.split(endCopy, [selection.endOffset]).segments[1]);
		}
		operations.removeParagraphs(leftoverCopies);
		return endCopy;
	};
	OutlineDragDropExecutor.prototype.setBlockResult = function (firstSlide, lastSlide, lastSelected) {
		const titleSources = this.getTitleSourcesOfSlide(firstSlide);
		if (titleSources.length) {
			this.resultStart = new AscCommonSlide.ContentPosition(getDocPosByOffset(titleSources[0], 0, false));
		} else {
			this.resultStart = new AscCommonSlide.SlidePosition(firstSlide, true);
		}
		if (lastSelected && lastSelected.GetParent()) {
			this.resultEnd = new AscCommonSlide.ContentPosition(getDocPosByOffset(lastSelected, getTextLength(lastSelected), false));
		} else {
			this.resultEnd = new AscCommonSlide.SlidePosition(lastSlide, false);
		}
		this.resultSlide = firstSlide;
	};
	OutlineDragDropExecutor.prototype.getTitleSourcesOfSlide = function (slide) {
		const outlineSlide = slide.getOutlineSlide();
		const titleContent = outlineSlide.title && outlineSlide.title.getDocContent();
		return titleContent ? titleContent.Content.slice() : [];
	};

	OutlineDragDropExecutor.prototype.executeMixed = function () {
		const selection = this.selection;
		const operations = this.operations;
		const slides = selection.slides;
		const lastSlide = slides[slides.length - 1];
		const startOutlineParagraph = this.getOutlineParagraph(selection.startIndex);
		const startSlide = this.reader.getSlide(startOutlineParagraph);
		const startParagraph = this.reader.getSourceParagraph(startOutlineParagraph);
		const targetParagraph = this.getOrCreateSourceParagraph(this.target.outlineIndex);
		if (!startParagraph || !targetParagraph) {
			return false;
		}
		const targetSlide = targetParagraph.GetParentShape().parent;
		const startContentSources = this.getSlideContentSources(startSlide);
		const leadTail = startContentSources.slice(startContentSources.indexOf(startParagraph) + 1);
		const isTargetTitle = operations.isTitleContent(targetParagraph.GetParent());
		const targetContentSources = this.getSlideContentSources(targetSlide);
		const targetTail = isTargetTitle ? targetContentSources : targetContentSources.slice(targetContentSources.indexOf(targetParagraph) + 1);

		let sourceLeft = null;
		let targetLeft;
		let targetRight;
		let leadCopies;
		let lastSelected;
		let leftover = [];
		let lastSlideTarget = lastSlide;
		if (this.isCopy) {
			const leadSourceCopies = operations.copyParagraphs([startParagraph].concat(leadTail));
			const slideCopies = this.copySlides(slides, targetSlide);
			lastSlideTarget = slideCopies[slideCopies.length - 1];
			lastSelected = this.cutCopyLeftover(lastSlide, lastSlideTarget);
			const targetSplit = operations.split(targetParagraph, [this.target.offset]);
			targetLeft = targetSplit.segments[0];
			targetRight = targetSplit.segments[1];
			operations.insertParagraphs(targetLeft.GetParent(), targetLeft.GetIndex() + 1, leadSourceCopies);
			leadCopies = this.trimCopies(leadSourceCopies, selection.startOffset);
		} else {
			const splitResults = this.splitAtPoints([
				{paragraph: startParagraph, offset: selection.startOffset},
				{paragraph: targetParagraph, offset: this.target.offset}
			]);
			const startSplit = this.getSplitSegments(splitResults, startParagraph);
			const targetSplit = this.getSplitSegments(splitResults, targetParagraph);
			sourceLeft = getSegmentEndingAt(startSplit, selection.startOffset);
			targetLeft = getSegmentEndingAt(targetSplit, this.target.offset);
			targetRight = getSegmentStartingAt(targetSplit, this.target.offset);
			const lead = [getSegmentStartingAt(startSplit, selection.startOffset)].concat(leadTail);

			const blockEnd = this.splitBlockEnd(lastSlide);
			lastSelected = blockEnd.lastSelected;
			leftover = blockEnd.leftover;

			leadCopies = operations.copyParagraphs(lead);
			operations.removeParagraphs(lead);
			operations.insertParagraphs(targetLeft.GetParent(), targetLeft.GetIndex() + 1, leadCopies);
		}
		if (isTargetTitle) {
			operations.setLevel(leadCopies, 0);
		}

		let leftoverCopies = [];
		if (leftover.length && sourceLeft) {
			leftoverCopies = operations.copyParagraphs(leftover);
			operations.removeParagraphs(leftover);
			const leftContent = sourceLeft.GetParent();
			operations.insertParagraphs(leftContent, sourceLeft.GetIndex() + 1, leftoverCopies);
		}

		if (!this.isCopy) {
			this.moveSlides(slides, targetSlide);
		}

		const tailCopies = this.moveParagraphsToSlideEnd([targetRight].concat(targetTail), lastSlideTarget);
		const targetRightCopy = tailCopies[0];

		this.resultStart = operations.track(leadCopies[0], 0);
		if (lastSelected && lastSelected.GetParent()) {
			this.resultEnd = operations.track(lastSelected, getTextLength(lastSelected));
		}

		operations.join(targetLeft, leadCopies[0]);
		if (lastSelected && targetRightCopy) {
			operations.join(lastSelected, targetRightCopy);
		}
		if (leftoverCopies.length) {
			operations.join(sourceLeft, leftoverCopies[0]);
		} else if (sourceLeft && getTextLength(operations.resolve(sourceLeft)) === 0) {
			operations.removeParagraphs([operations.resolve(sourceLeft)]);
		}

		this.resultStart = this.getTrackedContentPosition(this.resultStart, false);
		this.resultEnd = this.resultEnd ? this.getTrackedContentPosition(this.resultEnd, false) : new AscCommonSlide.SlidePosition(lastSlideTarget, false);
		this.resultSlide = targetSlide;
		this.finish();
		return true;
	};

	/**
	 * Drag'n'drop in the outline view: the gesture state machine (StartTrackText / EndTrackText analogue)
	 * and the entry points for the model-level API.
	 */
	function COutlineDragDropManager(outlineView) {
		this.outlineView = outlineView;
		this.state = OutlineDragState.none;
		this.startX = 0;
		this.startY = 0;
		this.target = null;
	}
	COutlineDragDropManager.prototype.getPresentation = function () {
		return this.outlineView.getPresentation();
	};
	COutlineDragDropManager.prototype.isActive = function () {
		return this.state !== OutlineDragState.none;
	};
	COutlineDragDropManager.prototype.isTrackText = function () {
		return this.state === OutlineDragState.tracking;
	};
	COutlineDragDropManager.prototype.canEdit = function () {
		const presentation = this.getPresentation();
		return !!presentation && presentation.CanEdit() && !AscCommon.CollaborativeEditing.Get_GlobalLockSelection();
	};
	COutlineDragDropManager.prototype.getDragSelection = function () {
		if (this.outlineView.isEmptyPresentation()) {
			return null;
		}
		this.getPresentation().updateSlideIndexes();
		return new OutlineDragSelectionBuilder(this.outlineView).build();
	};
	/**
	 * Expands the visible selection: whole titles, the whole first slide (FR-002, FR-009a, FR-014b).
	 * Called only when a drag starts.
	 */
	COutlineDragDropManager.prototype.normalizeSelection = function () {
		const selection = this.getDragSelection();
		const content = this.outlineView.getDocContent();
		if (!selection || !content) {
			return false;
		}
		const bounds = new OutlineDragSelectionBuilder(this.outlineView).getBounds();
		if (bounds && bounds.start.index === selection.startIndex && bounds.start.offset === selection.startOffset &&
			bounds.end.index === selection.endIndex && bounds.end.offset === selection.endOffset) {
			return false;
		}
		const startParagraph = content.Content[selection.startIndex];
		const endParagraph = content.Content[selection.endIndex];
		const startPos = getDocPosByOffset(startParagraph, selection.startOffset, false);
		const endPos = getDocPosByOffset(endParagraph, selection.endOffset, false);
		content.RemoveSelection();
		content.SetContentPosition(endPos, 0, 0);
		content.SetContentSelection(startPos, endPos, 0, 0, 0);
		this.outlineView.updateSelectionState();
		return true;
	};
	COutlineDragDropManager.prototype.getDropTarget = function (x, y) {
		const selection = this.getDragSelection();
		if (!selection) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		return new OutlineDropTargetBuilder(this.outlineView, selection).buildByPoint(x, y);
	};
	COutlineDragDropManager.prototype.getDropTargetByOutlinePos = function (outlineIndex, offset) {
		const selection = this.getDragSelection();
		if (!selection) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		return new OutlineDropTargetBuilder(this.outlineView, selection).buildByOutlinePos(outlineIndex, offset);
	};
	COutlineDragDropManager.prototype.getDropTargetByBoundary = function (outlineIndex, isAfter) {
		const selection = this.getDragSelection();
		if (!selection) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		return new OutlineDropTargetBuilder(this.outlineView, selection).buildByBoundary(outlineIndex, isAfter);
	};
	COutlineDragDropManager.prototype.canDrop = function (target) {
		if (!target || target.isNone() || target.isInsideSelection) {
			return false;
		}
		const selection = this.getDragSelection();
		if (!selection || selection.type !== target.selectionType) {
			return false;
		}
		return this.canEdit();
	};
	/**
	 * Executes a drop as one history point. Locks are checked here only, at drop (FR-019a).
	 * @returns {boolean} true if the document was changed
	 */
	COutlineDragDropManager.prototype.drop = function (target, isCopy) {
		if (!this.canDrop(target)) {
			return false;
		}
		const selection = this.getDragSelection();
		const executor = new OutlineDragDropExecutor(this.outlineView, selection, target, isCopy);
		const lockObjects = executor.getLockObjects();
		let isDone = false;
		this.outlineView.checkSelectedObjectsAndCallback(function () {
			isDone = executor.execute();
		}, AscDFH.historydescription_Document_DragText, true, true, !isCopy, undefined, lockObjects);
		return isDone;
	};

	COutlineDragDropManager.prototype.checkMouseDown = function (e, x, y) {
		this.reset();
		const outlineView = this.outlineView;
		const content = outlineView.getDocContent();
		const outlineShape = outlineView.outlineShape;
		if (!content || !outlineShape || !content.IsSelectionUse() || content.IsSelectionEmpty()) {
			return false;
		}
		if (e.Button !== 0 || e.ShiftKey || e.ClickCount > 1 || x < outlineShape.x || !this.canEdit()) {
			return false;
		}
		const tx = outlineView.getInvertTransformX(x);
		const ty = outlineView.getInvertTransformY(y);
		if (!content.CheckPosInSelection(tx, ty, 0, undefined)) {
			return false;
		}
		this.state = OutlineDragState.pending;
		this.startX = x;
		this.startY = y;
		return true;
	};
	COutlineDragDropManager.prototype.startTrackText = function () {
		if (this.state !== OutlineDragState.pending) {
			return;
		}
		this.normalizeSelection();
		this.state = OutlineDragState.tracking;
		this.target = null;
	};
	COutlineDragDropManager.prototype.checkMouseMove = function (e, x, y) {
		if (this.state !== OutlineDragState.tracking) {
			return false;
		}
		this.target = this.getDropTarget(x, y);
		return true;
	};
	COutlineDragDropManager.prototype.checkMouseUp = function (e, x, y) {
		if (this.state === OutlineDragState.pending) {
			this.reset();
			this.placeCursor(x, y);
			return true;
		}
		if (this.state === OutlineDragState.tracking) {
			this.target = this.getDropTarget(x, y);
			this.endTrackText(false);
			return true;
		}
		return false;
	};
	COutlineDragDropManager.prototype.endTrackText = function (isOnlyMoveTarget) {
		const target = this.target;
		const isCopy = AscCommon.global_mouseEvent.CtrlKey;
		this.reset();
		if (isOnlyMoveTarget || !target) {
			return false;
		}
		return this.drop(target, isCopy);
	};
	COutlineDragDropManager.prototype.cancelTrackText = function () {
		const isActive = this.isActive();
		this.reset();
		return isActive;
	};
	COutlineDragDropManager.prototype.reset = function () {
		this.state = OutlineDragState.none;
		this.target = null;
	};
	COutlineDragDropManager.prototype.placeCursor = function (x, y) {
		const outlineView = this.outlineView;
		const outlineShape = outlineView.outlineShape;
		if (!outlineShape) {
			return;
		}
		const mouseEvent = AscCommon.global_mouseEvent;
		const type = mouseEvent.Type;
		mouseEvent.Type = AscCommon.g_mouse_event_type_down;
		outlineView.selectionSetStart(mouseEvent, x, y);
		mouseEvent.Type = AscCommon.g_mouse_event_type_up;
		outlineShape.selectionSetEnd(mouseEvent, x, y, 0);
		mouseEvent.Type = type;
		outlineView.updateSelectionState();
	};
	COutlineDragDropManager.prototype.drawTarget = function (overlay) {
		const target = this.target;
		if (!this.isTrackText() || !target || target.isNone() || target.isInsideSelection || !overlay) {
			return;
		}
		const outlineShape = this.outlineView.outlineShape;
		if (!outlineShape) {
			return;
		}
		const transform = this.outlineView.getTransformText();
		const pixelRatio = AscCommon.AscBrowser.retinaPixelRatio;
		const koef = AscCommon.g_dKoef_mm_to_pix * pixelRatio;
		const step = Math.max(1, Math.round(pixelRatio));
		const ctx = overlay.m_oContext;
		const oldAlpha = ctx.globalAlpha;
		ctx.globalAlpha = 1;
		ctx.lineWidth = step;
		const indent = 0.5 * step;
		overlay.SetBaseTransform();
		if (target.kind === OutlineDropTargetKind.text) {
			const x = (transform.TransformPointX(target.x, target.y) * koef) >> 0;
			const y = (transform.TransformPointY(target.x, target.y) * koef) >> 0;
			const height = (target.height * koef) >> 0;
			overlay.CheckRect(x, y, 2 * step, height);
			this.drawDashes(ctx, "#000000", 0, height, step, function (i) {
				ctx.moveTo(x, y + i + indent);
				ctx.lineTo(x + 2 * step, y + i + indent);
			});
			this.drawDashes(ctx, "#FFFFFF", step, height, step, function (i) {
				ctx.moveTo(x, y + i + indent);
				ctx.lineTo(x + 2 * step, y + i + indent);
			});
		} else {
			const x1 = (transform.TransformPointX(0, target.y) * koef) >> 0;
			const x2 = (transform.TransformPointX(outlineShape.contentWidth || outlineShape.extX, target.y) * koef) >> 0;
			const y = (transform.TransformPointY(0, target.y) * koef) >> 0;
			const width = x2 - x1;
			overlay.CheckRect(x1, y - step, width, 2 * step);
			this.drawDashes(ctx, "#000000", 0, width, step, function (i) {
				ctx.moveTo(x1 + i, y + indent);
				ctx.lineTo(x1 + i, y - 2 * step + indent);
			});
			this.drawDashes(ctx, "#FFFFFF", step, width, step, function (i) {
				ctx.moveTo(x1 + i, y + indent);
				ctx.lineTo(x1 + i, y - 2 * step + indent);
			});
		}
		ctx.globalAlpha = oldAlpha;
	};
	COutlineDragDropManager.prototype.drawDashes = function (ctx, color, start, length, step, drawDash) {
		ctx.beginPath();
		ctx.strokeStyle = color;
		for (let i = start; i < length; i += 2 * step) {
			drawDash(i);
		}
		ctx.stroke();
		ctx.beginPath();
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].COutlineDragDropManager = COutlineDragDropManager;
	window["AscCommonSlide"].OutlineDragSelection = OutlineDragSelection;
	window["AscCommonSlide"].OutlineDropTarget = OutlineDropTarget;
	window["AscCommonSlide"].OutlineDragDropExecutor = OutlineDragDropExecutor;
	window["AscCommonSlide"].OutlineDragState = OutlineDragState;
	window["AscCommonSlide"].OutlineDragSelectionType = OutlineDragSelectionType;
	window["AscCommonSlide"].OutlineDropTargetKind = OutlineDropTargetKind;
})();
