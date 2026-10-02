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
		tracking: 2,
		dropping: 3
	};
	const OutlineDragSelectionType = {
		text: 0,
		slides: 1,
		mixed: 2,
		paragraphs: 3
	};
	const OutlineDropTargetKind = {
		none: 0,
		text: 1,
		boundary: 2
	};

	function getDocPos(paragraph, contentPos) {
		const pos = [{Class: paragraph.GetParent(), Position: paragraph.GetIndex()}];
		return paragraph.UpdatePosArrayByContentPos(pos, contentPos.Copy());
	}

	function comparePositions(a, b) {
		if (a.index !== b.index) {
			return a.index < b.index ? -1 : 1;
		}
		return a.pos.Compare(b.pos);
	}

	function OutlineDragSelection(type, startIndex, startPos, endIndex, endPos, slides, isEndInTitle) {
		this.type = type;
		this.startIndex = startIndex;
		this.startPos = startPos;
		this.endIndex = endIndex;
		this.endPos = endPos;
		this.slides = slides;
		this.isEndInTitle = isEndInTitle;
	}
	OutlineDragSelection.prototype.getStart = function () {
		return {index: this.startIndex, pos: this.startPos};
	};
	OutlineDragSelection.prototype.getEnd = function () {
		return {index: this.endIndex, pos: this.endPos};
	};

	function OutlineDropTarget(kind) {
		this.kind = kind;
		this.selectionType = null;
		this.outlineIndex = -1;
		this.contentPos = null;
		this.boundaryIndex = -1;
		this.isInsideSelection = false;
		this.x = 0;
		this.y = 0;
		this.height = 0;
	}
	OutlineDropTarget.prototype.isNone = function () {
		return this.kind === OutlineDropTargetKind.none;
	};

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

	function OutlineDragSelectionBuilder(outlineView, isBlock) {
		this.reader = new OutlineModelReader(outlineView);
		this.isBlock = !!isBlock;
	}
	OutlineDragSelectionBuilder.prototype.getBounds = function () {
		const content = this.reader.getDocContent();
		if (!content || !content.IsSelectionUse() || content.IsSelectionEmpty()) {
			return null;
		}
		const firstIndex = content.Selection.StartPos;
		const secondIndex = content.Selection.EndPos;
		const first = {index: firstIndex, pos: content.Content[firstIndex].Get_ParaContentPos(true, true)};
		const second = {index: secondIndex, pos: content.Content[secondIndex].Get_ParaContentPos(true, false)};
		if (comparePositions(first, second) > 0) {
			return {start: second, end: first};
		}
		return {start: first, end: second};
	};
	OutlineDragSelectionBuilder.prototype.getParagraph = function (index) {
		return this.reader.getParagraphs()[index];
	};
	OutlineDragSelectionBuilder.prototype.getEndPoint = function (index) {
		return {index: index, pos: this.getParagraph(index).Get_EndPos(true)};
	};
	OutlineDragSelectionBuilder.prototype.isAtBegin = function (point) {
		return this.getParagraph(point.index).IsCursorAtBegin(point.pos);
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
		const start = bounds.start;
		const end = bounds.end;
		const startParagraph = this.getParagraph(start.index);
		const blockEnd = (end.index > start.index && this.isAtBegin(end)) ? this.getEndPoint(end.index - 1) : end;
		if (this.reader.isTitle(startParagraph)) {
			return this.buildFromTitle(start, blockEnd);
		}
		if (this.hasTitleBetween(start.index + 1, blockEnd.index)) {
			return this.buildMixed(start, blockEnd);
		}
		if (this.isBlock) {
			return this.buildParagraphs(start, blockEnd);
		}
		return this.buildText(start, blockEnd);
	};
	OutlineDragSelectionBuilder.prototype.buildText = function (start, end) {
		return new OutlineDragSelection(OutlineDragSelectionType.text, start.index, start.pos, end.index, end.pos, [], false);
	};
	OutlineDragSelectionBuilder.prototype.buildFromTitle = function (start, end) {
		const startSlide = this.reader.getSlide(this.getParagraph(start.index));
		const titleRange = this.reader.getTitleRange(startSlide);
		if (end.index <= titleRange.last) {
			const isWholeTitle = start.index === titleRange.first && end.index === titleRange.last &&
				this.getParagraph(start.index).IsSelectionFromStart() && this.getParagraph(end.index).IsSelectionToEnd();
			if (!isWholeTitle && !this.isBlock) {
				return this.buildText(start, end);
			}
			return this.buildSlides(titleRange.first, this.getEndPoint(titleRange.last), [startSlide], true);
		}
		const endParagraph = this.getParagraph(end.index);
		const isEndInTitle = this.reader.isTitle(endParagraph);
		const endIndex = isEndInTitle ? this.reader.getTitleRange(this.reader.getSlide(endParagraph)).last : end.index;
		end = this.getEndPoint(endIndex);
		const slides = this.getSlidesBetween(titleRange.first, end.index);
		return this.buildSlides(titleRange.first, end, slides, isEndInTitle);
	};
	OutlineDragSelectionBuilder.prototype.buildSlides = function (startIndex, end, slides, isEndInTitle) {
		const firstSlide = slides[0];
		const lastSlide = slides[slides.length - 1];
		if (firstSlide.num === 0) {
			end = this.getEndPoint(this.reader.getSlideLastIndex(lastSlide));
			isEndInTitle = this.reader.isTitle(this.getParagraph(end.index));
		}
		const startPos = this.getParagraph(startIndex).Get_StartPos();
		return new OutlineDragSelection(OutlineDragSelectionType.slides, startIndex, startPos, end.index, end.pos, slides, isEndInTitle);
	};

	OutlineDragSelectionBuilder.prototype.buildMixed = function (start, end) {
		const endParagraph = this.getParagraph(end.index);
		const isEndInTitle = this.reader.isTitle(endParagraph);
		const endIndex = isEndInTitle ? this.reader.getTitleRange(this.reader.getSlide(endParagraph)).last : end.index;
		end = this.getEndPoint(endIndex);
		const startPos = this.getParagraph(start.index).Get_StartPos();
		const slides = this.getSlidesBetween(start.index + 1, end.index);
		return new OutlineDragSelection(OutlineDragSelectionType.mixed, start.index, startPos, end.index, end.pos, slides, isEndInTitle);
	};
	OutlineDragSelectionBuilder.prototype.buildParagraphs = function (start, end) {
		end = this.getEndPoint(end.index);
		const startPos = this.getParagraph(start.index).Get_StartPos();
		return new OutlineDragSelection(OutlineDragSelectionType.paragraphs, start.index, startPos, end.index, end.pos, [], false);
	};

	function OutlineDropTargetBuilder(outlineView, selection) {
		this.outlineView = outlineView;
		this.reader = new OutlineModelReader(outlineView);
		this.selection = selection;
	}

	OutlineDropTargetBuilder.prototype.isBlockSelection = function () {
		const type = this.selection.type;
		return type === OutlineDragSelectionType.slides || type === OutlineDragSelectionType.mixed || type === OutlineDragSelectionType.paragraphs;
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
		return this.buildText(paragraph.Index, nearPos.ContentPos.Copy(), nearPos);
	};
	OutlineDropTargetBuilder.prototype.buildByOutlinePos = function (outlineIndex, contentPos) {
		const paragraph = this.reader.getParagraphs()[outlineIndex];
		if (!paragraph || !contentPos) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		if (this.isBlockSelection()) {
			return this.buildBoundary(paragraph.IsCursorAtBegin(contentPos) ? outlineIndex : outlineIndex + 1);
		}
		const nearPos = paragraph.Get_NearestPos(0, 0, 0);
		const target = this.buildText(outlineIndex, contentPos.Copy(), null);
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
			return this.buildByOutlinePos(outlineIndex, isAfter ? paragraph.Get_EndPos(false) : paragraph.Get_StartPos());
		}
		return this.buildBoundary(isAfter ? outlineIndex + 1 : outlineIndex);
	};
	OutlineDropTargetBuilder.prototype.buildText = function (outlineIndex, contentPos, nearPos) {
		const target = new OutlineDropTarget(OutlineDropTargetKind.text);
		target.selectionType = this.selection.type;
		target.outlineIndex = outlineIndex;
		target.contentPos = contentPos;
		if (nearPos) {
			target.x = nearPos.X;
			target.y = nearPos.Y;
			target.height = nearPos.Height;
		}
		target.isInsideSelection = this.isTextTargetInside({index: outlineIndex, pos: contentPos});
		return target;
	};
	OutlineDropTargetBuilder.prototype.isTextTargetInside = function (point) {
		const selection = this.selection;
		if (comparePositions(point, selection.getStart()) < 0) {
			return false;
		}
		if (comparePositions(point, selection.getEnd()) <= 0) {
			return true;
		}
		const paragraphs = this.reader.getParagraphs();
		return point.index === selection.endIndex + 1 &&
			paragraphs[selection.endIndex].Get_EndPos(true).Compare(selection.endPos) === 0 &&
			paragraphs[point.index].IsCursorAtBegin(point.pos);
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
		if (boundaryIndex <= 0 && (this.selection.type === OutlineDragSelectionType.mixed || this.selection.type === OutlineDragSelectionType.paragraphs)) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
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
		const endParagraph = this.reader.getParagraphs()[selection.endIndex];
		const lastBoundary = endParagraph.IsCursorAtEnd(selection.endPos) ? selection.endIndex + 1 : selection.endIndex;
		return boundaryIndex >= selection.startIndex && boundaryIndex <= lastBoundary;
	};

	function OutlineParagraphOperations(outlineView) {
		this.outlineView = outlineView;
		this.tracked = [];
		this.touchedContents = {};
	}
	OutlineParagraphOperations.prototype.touch = function (content) {
		if (content) {
			this.touchedContents[content.GetId()] = content;
		}
	};
	OutlineParagraphOperations.prototype.track = function (paragraph, contentPos) {
		const nearPos = {Paragraph: paragraph, ContentPos: contentPos.Copy()};
		paragraph.Check_NearestPos(nearPos);
		this.tracked.push(nearPos);
		return nearPos;
	};
	OutlineParagraphOperations.prototype.clearTracked = function () {
		for (let i = 0; i < this.tracked.length; i += 1) {
			this.tracked[i].Paragraph.Clear_NearestPosArray();
		}
		this.tracked = [];
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

	function SlideCopyContentMap(slide, slideCopy) {
		this.slide = slide;
		this.slideCopy = slideCopy;
		this.contents = null;
	}
	SlideCopyContentMap.prototype.fillContents = function () {
		this.contents = {};
		const shapes = this.slide.cSld.spTree;
		const shapeCopies = this.slideCopy.cSld.spTree;
		for (let i = 0; i < shapes.length; i += 1) {
			const content = shapes[i].getDocContent();
			if (content) {
				const shapeCopy = shapeCopies[i];
				this.contents[content.Get_Id()] = (shapeCopy && shapeCopy.getDocContent()) || null;
			}
		}
	};
	SlideCopyContentMap.prototype.getContentCopy = function (paragraph) {
		if (!this.contents) {
			this.fillContents();
		}
		const parent = paragraph.GetParent();
		return this.contents[parent.Get_Id()] || null;
	};
	SlideCopyContentMap.prototype.getParagraph = function (paragraph) {
		const content = this.getContentCopy(paragraph);
		return content && content.Content[paragraph.GetIndex()] || null;
	};

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
			if (selection.type === OutlineDragSelectionType.paragraphs) {
				addSlideShapes(this.reader.getSlide(this.getOutlineParagraph(selection.startIndex)));
			} else if (selection.type !== OutlineDragSelectionType.text) {
				const firstSlide = selection.slides[0];
				addSlideShapes(presentation.Slides[presentation.GetSlideIndex(firstSlide) - 1]);
				addSlideShapes(this.reader.getSlide(this.getOutlineParagraph(selection.startIndex)));
				addSlideShapes(selection.slides[selection.slides.length - 1]);
			}
		}
		if (this.target.kind === OutlineDropTargetKind.text) {
			addSlideShapes(this.reader.getSlide(this.getOutlineParagraph(this.target.outlineIndex)));
		} else {
			addSlideShapes(this.getBoundaryAnchor().slide);
		}
		return {slides: slides, shapes: shapes, isOnlyAdditional: this.isCopy};
	};
	OutlineDragDropExecutor.prototype.execute = function () {
		try {
			switch (this.selection.type) {
				case OutlineDragSelectionType.text: {
					return this.executeText();
				}
				case OutlineDragSelectionType.slides:
				case OutlineDragSelectionType.mixed: {
					return this.executeBlock();
				}
				case OutlineDragSelectionType.paragraphs: {
					return this.executeParagraphs();
				}
			}
			return false;
		} finally {
			this.operations.clearTracked();
		}
	};

	OutlineDragDropExecutor.prototype.showResult = function () {
		const presentation = this.getPresentation();
		const slideIndex = this.resultSlide ? presentation.GetSlideIndex(this.resultSlide) : -1;
		if (slideIndex !== -1) {
			const wordControl = presentation.DrawingDocument.m_oWordControl;
			const thumbnails = wordControl.Thumbnails;
			if (thumbnails) {
				const isSelectPageEnabled = thumbnails.SelectPageEnabled;
				thumbnails.SelectPageEnabled = false;
				wordControl.GoToPage(slideIndex);
				thumbnails.SelectPageEnabled = isSelectPageEnabled;
			}
		}
		if (this.resultStart && this.resultEnd) {
			this.outlineView.setSavedPosition(new AscCommonSlide.SavedPosition(this.resultStart, this.resultEnd));
			this.outlineView.applySavedPositionToOutline();
		}
	};
	OutlineDragDropExecutor.prototype.getTargetPoint = function () {
		const outlineParagraph = this.getOutlineParagraph(this.target.outlineIndex);
		const sourceParagraph = this.reader.getSourceParagraph(outlineParagraph);
		if (sourceParagraph) {
			return {nearPos: this.operations.track(sourceParagraph, this.target.contentPos), placeholder: null};
		}
		return {nearPos: this.operations.track(outlineParagraph, outlineParagraph.Get_StartPos()), placeholder: outlineParagraph};
	};
	OutlineDragDropExecutor.prototype.createTargetTitle = function (placeholder) {
		const titleContent = this.outlineView.createTitleContent(placeholder);
		const titleParagraph = titleContent ? titleContent.Content[0] : null;
		return titleParagraph ? this.operations.track(titleParagraph, titleParagraph.Get_StartPos()) : null;
	};
	OutlineDragDropExecutor.prototype.rollback = function () {
		AscCommon.History.UndoLastPoint();
		this.outlineView.update();
	};
	OutlineDragDropExecutor.prototype.collectSelectedContent = function () {
		const outlineView = this.outlineView;
		const presentation = this.getPresentation();
		return AscFormat.ExecuteNoHistory(function () {
			const selectedContent = new AscCommonWord.CSelectedContent();
			outlineView.forEachSelectedContent(function (selectProps) {
				if (selectProps.content) {
					selectProps.content.GetSelectedContent(selectedContent);
				}
			});
			selectedContent.EndCollect(presentation);
			return selectedContent;
		}, this, []);
	};
	OutlineDragDropExecutor.prototype.copySelectedContent = function (collected) {
		const result = new AscCommonWord.CSelectedContent();
		for (let i = 0; i < collected.Elements.length; i += 1) {
			const element = collected.Elements[i];
			const paragraph = element.Element;
			const selectedElement = new AscCommonWord.CSelectedElement();
			selectedElement.SelectedAll = element.SelectedAll;
			selectedElement.Element = paragraph.Copy(paragraph.Parent, paragraph.DrawingDocument, {});
			result.Add(selectedElement);
		}
		result.EndCollect(this.getPresentation());
		return result;
	};
	OutlineDragDropExecutor.prototype.executeText = function () {
		const outlineView = this.outlineView;
		const collected = this.collectSelectedContent();
		if (!collected.Elements.length) {
			return false;
		}
		const targetPoint = this.getTargetPoint();
		if (!collected.CanInsert(targetPoint.nearPos)) {
			return false;
		}
		const selectedContent = this.copySelectedContent(collected);
		const nearPos = targetPoint.placeholder ? this.createTargetTitle(targetPoint.placeholder) : targetPoint.nearPos;
		if (!nearPos) {
			this.rollback();
			return false;
		}
		if (!this.isCopy) {
			outlineView.remove(false, AscWord.Direction.BACKWARD);
		}
		const outlineParagraph = nearPos.Paragraph.IsUseInDocument() && outlineView.sourceToOutlineMap[nearPos.Paragraph.Get_Id()];
		if (!outlineParagraph) {
			this.rollback();
			return false;
		}
		this.resultSlide = this.reader.getSlide(outlineParagraph);
		const outlineContent = outlineView.getDocContent();
		outlineContent.RemoveSelection();
		outlineContent.SetContentPosition(getDocPos(outlineParagraph, nearPos.ContentPos), 0, 0);
		outlineView.insertContent(selectedContent, true);
		this.setResultFromOutlineSelection();
		return true;
	};
	OutlineDragDropExecutor.prototype.setResultFromOutlineSelection = function () {
		const bounds = new OutlineDragSelectionBuilder(this.outlineView).getBounds();
		if (!bounds) {
			return;
		}
		const startParagraph = this.reader.getSourceParagraph(this.getOutlineParagraph(bounds.start.index));
		const endParagraph = this.reader.getSourceParagraph(this.getOutlineParagraph(bounds.end.index));
		if (!startParagraph || !endParagraph) {
			return;
		}
		this.resultStart = new AscCommonSlide.ContentPosition(getDocPos(startParagraph, bounds.start.pos));
		this.resultEnd = new AscCommonSlide.ContentPosition(getDocPos(endParagraph, bounds.end.pos));
	};

	OutlineDragDropExecutor.prototype.getSlideContentSources = function (slide) {
		return this.getSourceParagraphs(this.reader.getContentParagraphs(slide));
	};
	OutlineDragDropExecutor.prototype.getSlideIndexes = function (slides) {
		const presentation = this.getPresentation();
		const indexes = [];
		for (let i = 0; i < slides.length; i += 1) {
			indexes.push(presentation.GetSlideIndex(slides[i]));
		}
		indexes.sort(AscCommon.fSortAscending);
		return indexes;
	};
	OutlineDragDropExecutor.prototype.getSlideInsertPos = function (afterSlide) {
		return afterSlide ? this.getPresentation().GetSlideIndex(afterSlide) + 1 : 0;
	};
	OutlineDragDropExecutor.prototype.moveSlides = function (slides, afterSlide) {
		const presentation = this.getPresentation();
		const indexes = this.getSlideIndexes(slides);
		const position = this.getSlideInsertPos(afterSlide);
		const isInPlace = position >= indexes[0] && position <= indexes[indexes.length - 1] + 1;
		if (!isInPlace) {
			presentation.getViewManager().shiftSlides(position, indexes, false);
		}
		presentation.updateSlideIndexes();
	};
	OutlineDragDropExecutor.prototype.copySlides = function (slides, afterSlide) {
		const presentation = this.getPresentation();
		const newIndexes = presentation.getViewManager().shiftSlides(this.getSlideInsertPos(afterSlide), this.getSlideIndexes(slides), true);
		presentation.updateSlideIndexes();
		const copies = [];
		for (let i = 0; i < newIndexes.length; i += 1) {
			copies.push(presentation.Slides[newIndexes[i]]);
		}
		return copies;
	};
	OutlineDragDropExecutor.prototype.getBlockEnd = function (lastSlide) {
		const selection = this.selection;
		const contentSources = this.getSlideContentSources(lastSlide);
		if (selection.isEndInTitle) {
			const titleSources = this.getTitleSourcesOfSlide(lastSlide);
			return {lastSelected: titleSources[titleSources.length - 1] || null, leftover: contentSources};
		}
		const endOutlineParagraph = this.getOutlineParagraph(selection.endIndex);
		const endSource = this.reader.getSourceParagraph(endOutlineParagraph);
		const endIndexInContent = contentSources.indexOf(endSource);
		if (endIndexInContent === -1) {
			return {lastSelected: endSource, leftover: []};
		}
		const leftover = contentSources.slice(endIndexInContent + 1);
		return {lastSelected: endSource, leftover: leftover};
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

	OutlineDragDropExecutor.prototype.cutCopyLeftover = function (lastSlide, lastCopy) {
		const selection = this.selection;
		const operations = this.operations;
		const contentSources = this.getSlideContentSources(lastSlide);
		const copyMap = new SlideCopyContentMap(lastSlide, lastCopy);
		if (selection.isEndInTitle) {
			const copiesToRemove = [];
			for (let i = 0; i < contentSources.length; i += 1) {
				const paragraph = copyMap.getParagraph(contentSources[i]);
				if (paragraph) {
					copiesToRemove.push(paragraph);
				}
			}
			operations.removeParagraphs(copiesToRemove);
			const titleSources = this.getTitleSourcesOfSlide(lastSlide);
			const lastTitle = titleSources[titleSources.length - 1];
			return lastTitle ? copyMap.getParagraph(lastTitle) : null;
		}
		const endSource = this.reader.getSourceParagraph(this.getOutlineParagraph(selection.endIndex));
		const endIndexInContent = contentSources.indexOf(endSource);
		const endCopy = copyMap.getParagraph(endSource);
		if (endIndexInContent === -1 || !endCopy) {
			return endCopy;
		}
		const leftoverCopies = [];
		for (let i = endIndexInContent + 1; i < contentSources.length; i += 1) {
			const paragraph = copyMap.getParagraph(contentSources[i]);
			if (paragraph) {
				leftoverCopies.push(paragraph);
			}
		}
		operations.removeParagraphs(leftoverCopies);
		return endCopy;
	};
	OutlineDragDropExecutor.prototype.getTitleSourcesOfSlide = function (slide) {
		const outlineSlide = slide.getOutlineSlide();
		const titleContent = outlineSlide.title && outlineSlide.title.getDocContent();
		return titleContent ? titleContent.Content.slice() : [];
	};

	OutlineDragDropExecutor.prototype.getLiveContentParagraphs = function (slide) {
		const result = [];
		const contentShapes = slide.getOutlineSlide().content;
		for (let i = 0; i < contentShapes.length; i += 1) {
			const content = contentShapes[i].getDocContent();
			if (content && !content.Is_Empty()) {
				for (let j = 0; j < content.Content.length; j += 1) {
					result.push(content.Content[j]);
				}
			}
		}
		return result;
	};
	OutlineDragDropExecutor.prototype.getBoundaryAnchor = function () {
		if (this.target.boundaryIndex <= 0) {
			return {slide: null, paragraph: null};
		}
		const previous = this.getOutlineParagraph(this.target.boundaryIndex - 1);
		const slide = this.reader.getSlide(previous);
		if (this.reader.isTitle(previous)) {
			return {slide: slide, paragraph: null};
		}
		return {slide: slide, paragraph: this.reader.getSourceParagraph(previous)};
	};
	OutlineDragDropExecutor.prototype.insertAtAnchor = function (slide, paragraph, paragraphs) {
		const operations = this.operations;
		if (paragraph) {
			operations.insertParagraphs(paragraph.GetParent(), paragraph.GetIndex() + 1, paragraphs);
			return;
		}
		const contentParagraphs = this.getLiveContentParagraphs(slide);
		if (contentParagraphs.length) {
			const first = contentParagraphs[0];
			operations.insertParagraphs(first.GetParent(), first.GetIndex(), paragraphs);
		} else {
			operations.appendParagraphs(operations.getReceiverContent(slide), paragraphs);
		}
	};
	OutlineDragDropExecutor.prototype.getParagraphsAfterInsertion = function (anchor, leadCopies) {
		if (!anchor.slide) {
			return [];
		}
		const contentParagraphs = this.getLiveContentParagraphs(anchor.slide);
		const lastBefore = leadCopies.length ? leadCopies[leadCopies.length - 1] : anchor.paragraph;
		if (!lastBefore) {
			return contentParagraphs;
		}
		return contentParagraphs.slice(contentParagraphs.indexOf(lastBefore) + 1);
	};
	OutlineDragDropExecutor.prototype.getLeadParagraphs = function () {
		const selection = this.selection;
		if (selection.type !== OutlineDragSelectionType.mixed) {
			return [];
		}
		const startOutlineParagraph = this.getOutlineParagraph(selection.startIndex);
		const startParagraph = this.reader.getSourceParagraph(startOutlineParagraph);
		if (!startParagraph) {
			return null;
		}
		const startContentSources = this.getSlideContentSources(this.reader.getSlide(startOutlineParagraph));
		return startContentSources.slice(startContentSources.indexOf(startParagraph));
	};
	OutlineDragDropExecutor.prototype.setBlockResult = function (firstSlide, lastSlide, leadCopies, lastSelected) {
		if (leadCopies.length) {
			const first = leadCopies[0];
			this.resultStart = new AscCommonSlide.ContentPosition(getDocPos(first, first.Get_StartPos()));
		} else {
			const titleSources = this.getTitleSourcesOfSlide(firstSlide);
			this.resultStart = titleSources.length ?
				new AscCommonSlide.ContentPosition(getDocPos(titleSources[0], titleSources[0].Get_StartPos())) :
				new AscCommonSlide.SlidePosition(firstSlide, true);
		}
		if (lastSelected && lastSelected.GetParent()) {
			this.resultEnd = new AscCommonSlide.ContentPosition(getDocPos(lastSelected, lastSelected.Get_EndPos(true)));
		} else {
			this.resultEnd = new AscCommonSlide.SlidePosition(lastSlide, false);
		}
	};
	OutlineDragDropExecutor.prototype.executeBlock = function () {
		const presentation = this.getPresentation();
		const operations = this.operations;
		const slides = this.selection.slides;
		const firstSlide = slides[0];
		const lastSlide = slides[slides.length - 1];
		const lead = this.getLeadParagraphs();
		if (!lead) {
			return false;
		}
		const anchor = this.getBoundaryAnchor();
		if (this.isCopy) {
			return this.executeBlockCopy(lead, anchor);
		}
		const receiverSlide = presentation.Slides[presentation.GetSlideIndex(firstSlide) - 1] || null;
		const blockEnd = this.getBlockEnd(lastSlide);
		const leadCopies = operations.copyParagraphs(lead);
		operations.removeParagraphs(lead);
		const anchorIndexInLeftover = anchor.slide === lastSlide ? blockEnd.leftover.indexOf(anchor.paragraph) : -1;
		let leadSlide;
		if (anchorIndexInLeftover !== -1) {
			const leftoverBefore = this.moveParagraphsToSlideEnd(blockEnd.leftover.slice(0, anchorIndexInLeftover + 1), receiverSlide);
			if (leadCopies.length) {
				this.insertAtAnchor(receiverSlide, leftoverBefore[leftoverBefore.length - 1], leadCopies);
			}
			this.moveSlides(slides, receiverSlide);
			leadSlide = receiverSlide;
		} else {
			if (receiverSlide) {
				this.moveParagraphsToSlideEnd(blockEnd.leftover, receiverSlide);
			}
			if (leadCopies.length) {
				this.insertAtAnchor(anchor.slide, anchor.paragraph, leadCopies);
			}
			const tail = this.getParagraphsAfterInsertion(anchor, leadCopies);
			this.moveSlides(slides, anchor.slide);
			this.moveParagraphsToSlideEnd(tail, lastSlide);
			leadSlide = anchor.slide;
		}
		this.setBlockResult(firstSlide, lastSlide, leadCopies, blockEnd.lastSelected);
		this.resultSlide = leadCopies.length ? leadSlide : firstSlide;
		this.operations.checkShapesExtents();
		return true;
	};
	OutlineDragDropExecutor.prototype.executeBlockCopy = function (lead, anchor) {
		const slides = this.selection.slides;
		const lastSlide = slides[slides.length - 1];
		const slideCopies = this.copySlides(slides, anchor.slide);
		const lastCopy = slideCopies[slideCopies.length - 1];
		const lastSelected = this.cutCopyLeftover(lastSlide, lastCopy);
		const leadCopies = this.operations.copyParagraphs(lead);
		if (leadCopies.length) {
			this.insertAtAnchor(anchor.slide, anchor.paragraph, leadCopies);
		}
		const tail = this.getParagraphsAfterInsertion(anchor, leadCopies);
		this.moveParagraphsToSlideEnd(tail, lastCopy);
		this.setBlockResult(slideCopies[0], lastCopy, leadCopies, lastSelected);
		this.resultSlide = leadCopies.length ? anchor.slide : slideCopies[0];
		this.operations.checkShapesExtents();
		return true;
	};
	OutlineDragDropExecutor.prototype.executeParagraphs = function () {
		const operations = this.operations;
		const selection = this.selection;
		const outlineParagraphs = this.reader.getParagraphs().slice(selection.startIndex, selection.endIndex + 1);
		const sources = this.getSourceParagraphs(outlineParagraphs);
		const anchor = this.getBoundaryAnchor();
		if (!sources.length || !anchor.slide) {
			return false;
		}
		const copies = operations.copyParagraphs(sources);
		if (!this.isCopy) {
			operations.removeParagraphs(sources);
		}
		this.insertAtAnchor(anchor.slide, anchor.paragraph, copies);
		const first = copies[0];
		const last = copies[copies.length - 1];
		this.resultStart = new AscCommonSlide.ContentPosition(getDocPos(first, first.Get_StartPos()));
		this.resultEnd = new AscCommonSlide.ContentPosition(getDocPos(last, last.Get_EndPos(true)));
		this.resultSlide = anchor.slide;
		operations.checkShapesExtents();
		return true;
	};

	function COutlineDragDropManager(outlineView) {
		this.outlineView = outlineView;
		this.state = OutlineDragState.none;
		this.target = null;
		this.selection = null;
		this.isBlockDrag = false;
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
		return new OutlineDragSelectionBuilder(this.outlineView, this.isBlockDrag).build();
	};

	COutlineDragDropManager.prototype.normalizeSelection = function () {
		const selection = this.getDragSelection();
		const content = this.outlineView.getDocContent();
		if (!selection || !content) {
			return false;
		}
		const bounds = new OutlineDragSelectionBuilder(this.outlineView).getBounds();
		if (bounds && comparePositions(bounds.start, selection.getStart()) === 0 &&
			comparePositions(bounds.end, selection.getEnd()) === 0) {
			return false;
		}
		const startPos = getDocPos(content.Content[selection.startIndex], selection.startPos);
		const endPos = getDocPos(content.Content[selection.endIndex], selection.endPos);
		content.RemoveSelection();
		content.SetContentPosition(endPos, 0, 0);
		content.SetContentSelection(startPos, endPos, 0, 0, 0);
		this.outlineView.updateSelectionState();
		return true;
	};
	COutlineDragDropManager.prototype.getTrackedSelection = function () {
		if (!this.selection) {
			this.selection = this.getDragSelection();
		}
		return this.selection;
	};
	COutlineDragDropManager.prototype.onOutlineUpdate = function () {
		this.selection = null;
	};
	COutlineDragDropManager.prototype.getDropTarget = function (x, y) {
		const selection = this.isTrackText() ? this.getTrackedSelection() : this.getDragSelection();
		if (!selection) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		return new OutlineDropTargetBuilder(this.outlineView, selection).buildByPoint(x, y);
	};
	COutlineDragDropManager.prototype.getDropTargetByOutlinePos = function (outlineIndex, contentPos) {
		const selection = this.getDragSelection();
		if (!selection) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		return new OutlineDropTargetBuilder(this.outlineView, selection).buildByOutlinePos(outlineIndex, contentPos);
	};
	COutlineDragDropManager.prototype.getDropTargetByBoundary = function (outlineIndex, isAfter) {
		const selection = this.getDragSelection();
		if (!selection) {
			return new OutlineDropTarget(OutlineDropTargetKind.none);
		}
		return new OutlineDropTargetBuilder(this.outlineView, selection).buildByBoundary(outlineIndex, isAfter);
	};
	COutlineDragDropManager.prototype.canDrop = function (target, selection) {
		if (!target || target.isNone() || target.isInsideSelection) {
			return false;
		}
		if (!selection) {
			selection = this.getDragSelection();
		}
		if (!selection || selection.type !== target.selectionType) {
			return false;
		}
		return this.canEdit();
	};

	COutlineDragDropManager.prototype.drop = function (target, isCopy, selection) {
		if (!selection) {
			selection = this.getDragSelection();
		}
		if (!this.canDrop(target, selection)) {
			return false;
		}
		const executor = new OutlineDragDropExecutor(this.outlineView, selection, target, isCopy);
		const lockObjects = executor.getLockObjects();
		let isDone = false;
		this.outlineView.checkSelectedObjectsAndCallback(function () {
			isDone = executor.execute();
		}, AscDFH.historydescription_Document_DragText, true, true, !isCopy, undefined, lockObjects);
		if (isDone) {
			executor.showResult();
		}
		return isDone;
	};

	COutlineDragDropManager.prototype.canStartPending = function (e) {
		const content = this.outlineView.getDocContent();
		if (!content || !content.IsSelectionUse() || content.IsSelectionEmpty()) {
			return false;
		}
		if (e.Button !== 0 || e.ShiftKey || e.ClickCount > 1) {
			return false;
		}
		return this.canEdit();
	};
	COutlineDragDropManager.prototype.checkMouseDown = function (e, x, y) {
		this.reset();
		if (!this.canStartPending(e)) {
			return false;
		}
		const outlineView = this.outlineView;
		const content = outlineView.getDocContent();
		const outlineShape = outlineView.outlineShape;
		if (!outlineShape || x < outlineShape.x) {
			return false;
		}
		const tx = outlineView.getInvertTransformX(x);
		const ty = outlineView.getInvertTransformY(y);
		if (!content.CheckPosInSelection(tx, ty, 0, undefined)) {
			return false;
		}
		this.state = OutlineDragState.pending;
		return true;
	};
	COutlineDragDropManager.prototype.checkDecorationMouseDown = function (e, x, y) {
		this.reset();
		if (!this.canStartPending(e)) {
			return false;
		}
		this.state = OutlineDragState.pending;
		this.isBlockDrag = true;
		return true;
	};
	COutlineDragDropManager.prototype.startTrackText = function () {
		if (this.state !== OutlineDragState.pending) {
			return;
		}
		this.normalizeSelection();
		this.state = OutlineDragState.tracking;
		this.target = null;
		this.selection = this.getDragSelection();
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
			const isBlockDrag = this.isBlockDrag;
			this.reset();
			if (!isBlockDrag) {
				this.placeCursor(x, y);
			}
			return true;
		}
		if (this.state === OutlineDragState.tracking) {
			this.target = this.getDropTarget(x, y);
			this.endTrackText();
			return true;
		}
		return false;
	};
	COutlineDragDropManager.prototype.endTrackText = function () {
		const target = this.target;
		const selection = this.selection;
		const isCopy = AscCommon.global_mouseEvent.CtrlKey;
		if (!target) {
			this.reset();
			return false;
		}
		this.state = OutlineDragState.dropping;
		const res = this.drop(target, isCopy, selection);
		this.reset();
		return res;
	};
	COutlineDragDropManager.prototype.cancelTrackText = function () {
		const isActive = this.isActive();
		this.reset();
		return isActive;
	};
	COutlineDragDropManager.prototype.reset = function () {
		this.state = OutlineDragState.none;
		this.target = null;
		this.selection = null;
		this.isBlockDrag = false;
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
