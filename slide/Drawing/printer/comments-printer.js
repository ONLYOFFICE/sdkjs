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

	const COMMENTSPRINTER_HORIZONTAL_FIELD = AscCommonSlide.CHANDOUTMASTER_HORIZONTAL_FIELD;
	const COMMENTSPRINTER_VERTICAL_FIELD = AscCommonSlide.CHANDOUTMASTER_VERTICAL_FIELD;
	const COMMENTSPRINTER_INITIALS_WIDTH = 25;
	const COMMENTSPRINTER_BLOCK_GAP = 5;
	const COMMENTSPRINTER_HEADER_LINE_OFFSET = 2;
	const COMMENTSPRINTER_REPLY_INDENT = 10;
	const COMMENTSPRINTER_MIN_FIRST_FRAGMENT_HEIGHT = 10;

	const COMMENTSPRINTER_FONT_HEADER = 18;
	const COMMENTSPRINTER_FONT_DATE = 10;
	const COMMENTSPRINTER_FONT_TEXT = 11;
	const COMMENTSPRINTER_FONT_LABEL = 10;

	const COMMENTSPRINTER_DATE_COLOR = {R: 110, G: 110, B: 110};

	const executeWithContentLimits = AscCommonSlide.executeWithContentLimits;
	function getSlideText() {
		return AscCommon.translateManager.getValue("Slide");
	}
	function getContinuedText() {
		return AscCommon.translateManager.getValue("continued");
	}

	function getTextPr(opts) {
		opts = opts || {};
		const textPr = new CTextPr();
		textPr.SetFontSize(opts.size || COMMENTSPRINTER_FONT_TEXT);
		textPr.SetBold(!!opts.bold);
		textPr.SetFontFamily("Arial");
		textPr.RFonts.SetAll("Arial");
		if (opts.color) {
			textPr.Color = new AscCommonWord.CDocumentColor(opts.color.R, opts.color.G, opts.color.B);
		}
		return textPr;
	}

	function getParaPr(opts) {
		opts = opts || {};
		const paraPr = new CParaPr();
		paraPr.Spacing.Before = (opts.spacingBefore != null) ? opts.spacingBefore : 0;
		paraPr.Spacing.After = (opts.spacingAfter != null) ? opts.spacingAfter : 0;
		return paraPr;
	}

	function formatCommentDate(timeValue) {
		const ms = parseInt(timeValue, 10);
		if (isNaN(ms) || ms <= 0) {
			return "";
		}
		try {
			return new Date(ms).toISOString();
		} catch (e) {
			return "";
		}
	}

	function splitTextIntoLines(text) {
		if (typeof text !== "string" || text.length === 0) {
			return [""];
		}
		const lines = text.split(new RegExp("\\r\\n|\\r|\\n"));
		while (lines.length > 1 && lines[lines.length - 1] === "") {
			lines.pop();
		}
		return lines;
	}

	function addParagraphToContent(content, runs, paraPr) {
		AscFormat.ExecuteNoHistory(function () {
			const paragraph = new AscWord.Paragraph(content, true);
			content.AddToContent(content.Content.length, paragraph);
			if (paraPr) {
				paragraph.Set_Pr(getParaPr(paraPr));
			}
			for (let i = 0; i < runs.length; i += 1) {
				const runInfo = runs[i];
				const run = new AscWord.Run(paragraph, false);
				run.AddText(runInfo.text || "");
				paragraph.AddToContent(paragraph.Content.length - 1, run);
				if (runInfo.textPr) {
					run.Set_Pr(getTextPr(runInfo.textPr));
				}
			}
		}, this, []);
	}

	function createEmptyBlockShape(width) {
		return AscFormat.ExecuteNoHistory(function () {
			const shape = new AscFormat.CShape();
			shape.setBDeleted(false);
			shape.extX = width;
			shape.extY = 100;
			shape.createTextBody();
			shape.txBody.bodyPr.setInsets(0, 0, 0, 0);
			shape.txBody.content.ClearContent(false);
			return shape;
		}, this, []);

	}

	function getRecalcPageBottom(content, pageIndex) {
		const bounds = content.GetContentBounds(pageIndex);
		if (!bounds) {
			return 0;
		}
		return Math.max(0, bounds.Bottom - bounds.Top);
	}

	function drawInlineShape(graphics, x, y, width, runs) {
		AscFormat.ExecuteNoHistory(function () {
			const shape = createEmptyBlockShape(width);
			const content = shape.getDocContent();
			addParagraphToContent(content, runs, {spacingAfter: 1});
			executeWithContentLimits(function () {
				content.Recalculate_Page(0, true);
			}, content, 0, 0, width, 1000);

			graphics.SaveGrState();
			const transform = new AscCommon.CMatrix();
			transform.tx = x;
			transform.ty = y;
			graphics.transform3(transform);
			content.Set_StartPage(0);
			content.Draw(0, graphics);
			graphics.RestoreGrState();
		}, this, []);
	}

	function CommentsBlock(slideIdx, indent, hasInitialsLabel, label) {
		this.slideIdx = slideIdx;
		this.indent = indent || 0;
		this.hasInitialsLabel = !!hasInitialsLabel;
		this.label = label || null;
		this.shape = null;
	}

	CommentsBlock.prototype.createShape = function (width) {
		this.shape = createEmptyBlockShape(width);
	};

	CommentsBlock.prototype.getContent = function () {
		return this.shape.getDocContent();
	};

	CommentsBlock.prototype.isHeader = function () {
		return false;
	};

	CommentsBlock.prototype.recalcPage = function (width, pageIdx, pageHeight) {
		const content = this.getContent();
		let res;
		executeWithContentLimits(function () {
			res = content.Recalculate_Page(pageIdx, true);
		}, content, 0, 0, width, pageHeight);
		return {
			result: res,
			height: getRecalcPageBottom(content, pageIdx)
		};
	};

	CommentsBlock.prototype.drawFragment = function (graphics, fragmentIndex, x, y) {
		graphics.SaveGrState();
		const transform = new AscCommon.CMatrix();
		transform.tx = x;
		transform.ty = y;
		graphics.transform3(transform);
		const content = this.getContent();
		content.Set_StartPage(0);
		content.Draw(fragmentIndex, graphics);
		graphics.RestoreGrState();
	};

	CommentsBlock.prototype.drawInitialsLabel = function (graphics, x, y) {
		if (!this.hasInitialsLabel) return;
		const labelLines = [{
			text  : this.label || "",
			textPr: {size: COMMENTSPRINTER_FONT_LABEL, bold: true, color: COMMENTSPRINTER_DATE_COLOR}
		}];
		drawInlineShape(graphics, x, y, COMMENTSPRINTER_INITIALS_WIDTH, labelLines);
	};

	function HeaderBlock(slideIdx, width) {
		CommentsBlock.call(this, slideIdx, 0, false, null);
		this.createShape(width);
		this.fillContent();
	}

	AscFormat.InitClassWithoutType(HeaderBlock, CommentsBlock);

	HeaderBlock.prototype.fillContent = function () {
		AscFormat.ExecuteNoHistory(function () {
			const content = this.getContent();
			addParagraphToContent(content, [{
				text  : this.getTitle(),
				textPr: {size: COMMENTSPRINTER_FONT_HEADER, bold: true, color: COMMENTSPRINTER_DATE_COLOR}
			}], {spacingAfter: 1});
		}, this, []);
	};
	HeaderBlock.prototype.getTitle = function () {
		return getSlideText() + " " + (this.slideIdx + 1);
	};
	HeaderBlock.prototype.isHeader = function () {
		return true;
	};

	function ContinuedHeaderBlock(slideIdx, width) {
		HeaderBlock.call(this, slideIdx, width);
	}
	AscFormat.InitClassWithoutType(ContinuedHeaderBlock, HeaderBlock);
	ContinuedHeaderBlock.prototype.getTitle = function () {
		return getSlideText() + " " + (this.slideIdx + 1) + " (" + getContinuedText() + ")";
	};

	function CommentBlock(opts, width) {
		CommentsBlock.call(this, opts.slideIdx, opts.indent || 0, true, opts.label || null);
		this.text = opts.text;
		this.authorName = opts.authorName;
		this.time = opts.time;
		this.createShape(width);
		this.fillContent();
	}

	AscFormat.InitClassWithoutType(CommentBlock, CommentsBlock);

	CommentBlock.prototype.fillContent = function () {
		AscFormat.ExecuteNoHistory(function () {
			const content = this.getContent();

			const textPr = {size: COMMENTSPRINTER_FONT_TEXT};
			const lines = splitTextIntoLines(this.text);
			for (let i = 0; i < lines.length; i += 1) {
				addParagraphToContent(content, [{text: lines[i], textPr: textPr}], {spacingAfter: 1});
			}

			const infoPr = {size: COMMENTSPRINTER_FONT_DATE, color: COMMENTSPRINTER_DATE_COLOR};
			const dateStr = formatCommentDate(this.time);
			const authorRuns = [{text: this.authorName || "", textPr: infoPr}];
			if (dateStr) {
				authorRuns.push({text: "; " + dateStr, textPr: infoPr});
			}
			addParagraphToContent(content, authorRuns, {spacingBefore: 1, spacingAfter: 1});
		}, this, []);
	};

	function CommentsFragment(block, fragmentIndex, y, height) {
		this.block = block;
		this.fragmentIndex = fragmentIndex;
		this.y = y;
		this.height = height;
	}

	CommentsFragment.prototype.isFirstFragment = function () {
		return this.fragmentIndex === 0;
	};

	function CommentsPage() {
		this.fragments = [];
	}

	CommentsPage.prototype.addFragment = function (block, fragmentIndex, y, height) {
		this.fragments.push(new CommentsFragment(block, fragmentIndex, y, height));
	};

	CommentsPage.prototype.isEmpty = function () {
		return this.fragments.length === 0;
	};

	CommentsPage.prototype.getFragments = function () {
		return this.fragments;
	};

	function CommentsSegment() {
		this.blocks = [];
		this.pages = [];
		this.headerContentHeight = 0;
		this.headerReservedHeight = 0;
	}

	CommentsSegment.prototype.addBlock = function (block) {
		this.blocks.push(block);
	};

	CommentsSegment.prototype.hasBlocks = function () {
		return this.blocks.length > 0;
	};

	CommentsSegment.prototype.getBlocks = function () {
		return this.blocks;
	};

	CommentsSegment.prototype.setPages = function (pages) {
		this.pages = pages;
	};

	CommentsSegment.prototype.getPage = function (idx) {
		return this.pages[idx];
	};

	CommentsSegment.prototype.getPagesCount = function () {
		return this.pages.length;
	};

	CommentsSegment.prototype.setHeaderMetrics = function (contentHeight, reservedHeight) {
		this.headerContentHeight = contentHeight;
		this.headerReservedHeight = reservedHeight;
	};

	function SegmentLayouter(segment, usableSize) {
		this.segment = segment;
		this.usable = usableSize;
		this.pages = [];
		this.currentPage = null;
		this.yCursor = 0;
		this.openSlideIdx = null;
	}

	SegmentLayouter.getBlockShapeWidth = function (block, usableWidth) {
		if (block.isHeader()) {
			return usableWidth;
		}
		return usableWidth - block.indent - (block.hasInitialsLabel ? COMMENTSPRINTER_INITIALS_WIDTH : 0);
	};

	SegmentLayouter.prototype.recalculate = function () {
		this.computeHeaderMetrics();
		const blocks = this.segment.getBlocks();
		for (let i = 0; i < blocks.length; i += 1) {
			this.recalculateBlock(blocks[i]);
		}
		this.flushPage();
		this.segment.setPages(this.pages);
	};

	SegmentLayouter.prototype.computeHeaderMetrics = function () {
		AscCommon.g_oTextMeasurer.SetFontInternal("Arial", COMMENTSPRINTER_FONT_HEADER, 1);
		const headerLineHeight = AscCommon.g_oTextMeasurer.GetHeight();
		this.segment.setHeaderMetrics(headerLineHeight, headerLineHeight + COMMENTSPRINTER_HEADER_LINE_OFFSET);
	};

	SegmentLayouter.prototype.flushPage = function () {
		if (this.currentPage && !this.currentPage.isEmpty()) {
			this.pages.push(this.currentPage);
		}
		this.currentPage = null;
		this.yCursor = 0;
	};

	SegmentLayouter.prototype.ensurePage = function (block) {
		if (this.currentPage) {
			return;
		}
		this.currentPage = new CommentsPage();
		this.yCursor = 0;
		if (!block.isHeader() && this.openSlideIdx !== null) {
			this.placeContinuedHeader(this.openSlideIdx);
		}
	};

	SegmentLayouter.prototype.placeContinuedHeader = function (slideIdx) {
		const width = this.usable.width;
		const block = new ContinuedHeaderBlock(slideIdx, width);
		block.recalcPage(width, 0, this.usable.height);
		this.placeFragment(block, 0, this.segment.headerReservedHeight);
	};

	SegmentLayouter.prototype.placeFragment = function (block, fragmentIndex, height) {
		const gap = (this.yCursor > 0) ? COMMENTSPRINTER_BLOCK_GAP : 0;
		const placeY = this.yCursor + gap;
		this.currentPage.addFragment(block, fragmentIndex, placeY, height);
		this.yCursor = placeY + height;
	};

	SegmentLayouter.prototype.getRemainingHeight = function () {
		const gap = (this.yCursor > 0) ? COMMENTSPRINTER_BLOCK_GAP : 0;
		return this.usable.height - this.yCursor - gap;
	};

	SegmentLayouter.prototype.recalculateBlock = function (block) {
		if (block.isHeader()) {
			this.openSlideIdx = block.slideIdx;
		}
		const blockWidth = SegmentLayouter.getBlockShapeWidth(block, this.usable.width);

		this.ensurePage(block);
		const probe = block.recalcPage(blockWidth, 0, this.usable.height);

		if (block.isHeader()) {
			this.recalculateHeaderBlock(block, blockWidth);
			return;
		}

		const remainingHeight = this.getRemainingHeight();
		if (probe.result === recalcresult2_End && probe.height <= remainingHeight) {
			this.placeFragment(block, 0, probe.height);
			return;
		}

		this.recalculateMultiFragmentBlock(block, blockWidth);
	};

	SegmentLayouter.prototype.recalculateHeaderBlock = function (block, blockWidth) {
		const placedHeight = this.segment.headerReservedHeight;
		if (placedHeight > this.getRemainingHeight() && this.yCursor > 0) {
			this.flushPage();
			this.ensurePage(block);
			block.recalcPage(blockWidth, 0, this.usable.height);
		}
		this.placeFragment(block, 0, placedHeight);
	};

	SegmentLayouter.prototype.recalculateMultiFragmentBlock = function (block, blockWidth) {
		let remainingHeight = this.getRemainingHeight();
		if (remainingHeight < COMMENTSPRINTER_MIN_FIRST_FRAGMENT_HEIGHT) {
			this.flushPage();
			this.ensurePage(block);
			remainingHeight = this.getRemainingHeight();
		}

		let pageIdx = 0;
		while (true) {
			const probe = block.recalcPage(blockWidth, pageIdx, remainingHeight);

			if (probe.height <= 0) {
				if (this.yCursor === 0) {
					break;
				}
				this.flushPage();
				this.ensurePage(block);
				remainingHeight = this.getRemainingHeight();
				pageIdx = 0;
				continue;
			}

			this.placeFragment(block, pageIdx, probe.height);
			if (probe.result === recalcresult2_End) {
				break;
			}
			pageIdx += 1;
			this.flushPage();
			this.ensurePage(block);
			remainingHeight = this.getRemainingHeight();
		}
	};

	function CommentsPrinter(presentation, printOptions, pageSizesProvider) {
		this.presentation = presentation;
		this.printOptions = printOptions;
		this.pageSizesProvider = pageSizesProvider || null;
		this.commentsCache = {};
	}

	CommentsPrinter.prototype.getPageSizes = function () {
		if (typeof this.pageSizesProvider === "function") {
			return this.pageSizesProvider();
		}
		const pageOptions = this.printOptions.pageOptions;
		return {width: pageOptions.width, height: pageOptions.height};
	};

	CommentsPrinter.prototype.getContentSize = function () {
		const pageSizes = this.getPageSizes();
		return {
			width : pageSizes.width - 2 * COMMENTSPRINTER_HORIZONTAL_FIELD,
			height: pageSizes.height - 2 * COMMENTSPRINTER_VERTICAL_FIELD
		};
	};

	CommentsPrinter.prototype.hasAnyCommentsForSlides = function (slides) {
		for (let i = 0; i < slides.length; i += 1) {
			const slide = slides[i];
			if (slide && slide.slideComments && slide.slideComments.comments && slide.slideComments.comments.length > 0) {
				return true;
			}
		}
		return false;
	};

	CommentsPrinter.prototype.getSegment = function (slides) {
		if (!this.hasAnyCommentsForSlides(slides)) {
			return null;
		}
		const segment = new CommentsSegment();
		this.buildSegmentBlocks(segment, slides);
		if (!segment.hasBlocks()) {
			return null;
		}
		this.recalculateSegment(segment);
		return segment;
	};

	CommentsPrinter.prototype.buildSegmentBlocks = function (segment, slides) {
		const contentSize = this.getContentSize();
		for (let s = 0; s < slides.length; s += 1) {
			const slide = slides[s];
			if (!slide || !slide.slideComments) continue;
			const comments = slide.slideComments.comments;
			if (!comments || comments.length === 0) continue;
			this.appendSlideBlocks(segment, slide.num, comments, contentSize);
		}
	};

	CommentsPrinter.prototype.appendSlideBlocks = function (segment, slideIdx, comments, contentSize) {
		segment.addBlock(new HeaderBlock(slideIdx, contentSize.width));
		for (let i = 0; i < comments.length; i += 1) {
			this.appendCommentBlocks(segment, slideIdx, comments[i], i + 1, contentSize);
		}
	};

	CommentsPrinter.prototype.getInitials = function (userName) {
		const presentation = this.presentation;
		const author = presentation.CommentAuthors[userName];
		return author ? author.Initials : AscCommon.translateManager.getValue("Unknown User");
	};
	CommentsPrinter.prototype.getColorIndex = function (userName) {
		const presentation = this.presentation;
		const author = presentation.CommentAuthors[userName];
		return author ? author.GetColorIndex() : 0;
	};
	CommentsPrinter.prototype.appendCommentBlocks = function (segment, slideIdx, comment, commentNumber, contentSize) {
		const data = comment.Data;
		if (!data) {
			return;
		}
		const label = this.getInitials(data.m_sUserName) + commentNumber;
		const colorIndex = this.getColorIndex(data.m_sUserName);
		segment.addBlock(new CommentBlock({
			slideIdx  : slideIdx,
			indent    : 0,
			label     : label,
			text      : data.m_sText,
			authorName: data.m_sUserName,
			time      : data.m_sTime
		}, contentSize.width - COMMENTSPRINTER_INITIALS_WIDTH));
		this.commentsCache[comment.GetId()] = {
			label: label,
			colorIndex: colorIndex
		};

		const replies = data.m_aReplies || [];
		const replyWidth = contentSize.width - COMMENTSPRINTER_INITIALS_WIDTH - COMMENTSPRINTER_REPLY_INDENT;
		for (let i = 0; i < replies.length; i += 1) {
			const reply = replies[i];
			segment.addBlock(new CommentBlock({
				slideIdx  : slideIdx,
				indent    : COMMENTSPRINTER_REPLY_INDENT,
				label     : this.getInitials(reply.m_sUserName) + commentNumber + " - " + (i + 1),
				text      : reply.m_sText,
				authorName: reply.m_sUserName,
				time      : reply.m_sTime
			}, replyWidth));
		}
	};

	CommentsPrinter.prototype.recalculateSegment = function (segment) {
		new SegmentLayouter(segment, this.getContentSize()).recalculate();
	};

	CommentsPrinter.prototype.getSegmentPagesCount = function (segment) {
		return segment ? segment.getPagesCount() : 0;
	};

	CommentsPrinter.prototype.drawSegmentPage = function (graphics, segment, localPageIndex) {
		const pageSizes = this.getPageSizes();
		const baseTransform = graphics.GetBaseTransform();
		const composedTransform = baseTransform ? baseTransform.CreateDublicate() : new AscCommon.CMatrix();

		graphics.SaveGrState();
		graphics.SetIntegerGrid(false);
		graphics.AddClipRect(0, 0, pageSizes.width, pageSizes.height);
		graphics.SetBaseTransform(composedTransform);
		graphics.reset();
		graphics.SetIntegerGrid(false);
		this.drawSegmentPageContent(graphics, segment, localPageIndex);
		graphics.SetBaseTransform(baseTransform);
		graphics.reset();
		graphics.RestoreGrState();
	};

	CommentsPrinter.prototype.drawSegmentPageContent = function (graphics, segment, localPageIndex) {
		const page = segment.getPage(localPageIndex);
		if (!page) {
			return;
		}
		const contentSize = this.getContentSize();
		const pageSizes = this.getPageSizes();
		const xOriginBase = COMMENTSPRINTER_HORIZONTAL_FIELD;
		const yOrigin = COMMENTSPRINTER_VERTICAL_FIELD;

		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, pageSizes.width, pageSizes.height);

		const fragments = page.getFragments();
		for (let i = 0; i < fragments.length; i += 1) {
			const frag = fragments[i];
			const block = frag.block;
			const blockX = xOriginBase + block.indent + (block.hasInitialsLabel ? COMMENTSPRINTER_INITIALS_WIDTH : 0);
			const blockY = yOrigin + frag.y;

			if (block.hasInitialsLabel && frag.isFirstFragment()) {
				block.drawInitialsLabel(graphics, xOriginBase + block.indent, blockY);
			}

			block.drawFragment(graphics, frag.fragmentIndex, blockX, blockY);

			if (block.isHeader()) {
				const lineY = blockY + segment.headerContentHeight + COMMENTSPRINTER_HEADER_LINE_OFFSET;
				this.drawHeaderLine(graphics, xOriginBase, lineY, contentSize.width);
			}
		}

		graphics.RestoreGrState();
	};

	CommentsPrinter.prototype.drawHeaderLine = function (graphics, x, y, width) {
		graphics.p_color(COMMENTSPRINTER_DATE_COLOR.R, COMMENTSPRINTER_DATE_COLOR.G, COMMENTSPRINTER_DATE_COLOR.B, 255);
		graphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, y, x, x + width, 0);
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].CommentsPrinter = CommentsPrinter;
})();
