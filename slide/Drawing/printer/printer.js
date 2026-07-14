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

	const executeWithContentLimits = AscCommonSlide.executeWithContentLimits;
	const CHANDOUTMASTER_HORIZONTAL_FIELD = AscCommonSlide.CHANDOUTMASTER_HORIZONTAL_FIELD;
	const CHANDOUTMASTER_VERTICAL_FIELD = AscCommonSlide.CHANDOUTMASTER_VERTICAL_FIELD;
	const SPECIFICPRINTER_FITPAGE_MARGIN = 1 * AscCommon.g_dKoef_pix_to_mm;
	const OUTLINEPRINTER_DECORATIONS_OFFSET_LEFT = 15;

	function CPrinterPage(printer) {
		this.printer = printer;
	}
	CPrinterPage.prototype.draw = function (graphics) {
		const printer = this.printer;
		const pageSizes = printer.getPageSizes();
		const contentSizes = printer.getContentSizes();
		const scaleCoefficient = printer.isScaleToFitPaper() ? printer.getScaleToFitPaperCoefficient() : 1;
		const m = new AscCommon.CMatrix();
		m.Scale(scaleCoefficient, scaleCoefficient);
		m.Translate((pageSizes.width - contentSizes.width * scaleCoefficient) / 2, (pageSizes.height - contentSizes.height * scaleCoefficient) / 2);
		const baseTransform = graphics.GetBaseTransform();
		const composedTransform = baseTransform ? baseTransform.CreateDublicate() : new AscCommon.CMatrix();
		composedTransform.Multiply(m, AscCommon.MATRIX_ORDER_PREPEND);

		graphics.SaveGrState();
		graphics.SetIntegerGrid(false);
		graphics.AddClipRect(0, 0, pageSizes.width, pageSizes.height);
		graphics.SetBaseTransform(composedTransform);
		graphics.reset();
		graphics.SetIntegerGrid(false);
		this.drawContent(graphics);
		graphics.SetBaseTransform(baseTransform);
		graphics.reset();
		graphics.RestoreGrState();

		if (printer.isDrawFrame()) {
			graphics.SaveGrState();
			graphics.SetBaseTransform(composedTransform);
			graphics.reset();
			graphics.SetIntegerGrid(false);
			printer.drawFrame(graphics);
			graphics.SetBaseTransform(baseTransform);
			graphics.reset();
			graphics.RestoreGrState();
		}
	};
	CPrinterPage.prototype.drawContent = function (graphics) {};

	function CSlidePage(printer, slide) {
		CPrinterPage.call(this, printer);
		this.slide = slide;
	}
	AscFormat.InitClassWithoutType(CSlidePage, CPrinterPage);
	CSlidePage.prototype.drawContent = function (graphics) {
		const printer = this.printer;
		const slide = this.slide;
		const contentSizes = printer.getContentSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, contentSizes.width, contentSizes.height);
		graphics.m_bNoPrintSlideBackground = printer.isSkipSlideBackground();
		slide.draw(graphics);
		graphics.m_bNoPrintSlideBackground = false;
		graphics.RestoreGrState();
	};

	function CHandoutPage(printer, pageIdx, slides) {
		CPrinterPage.call(this, printer);
		this.pageIdx = pageIdx;
		this.slides = slides;
	}
	AscFormat.InitClassWithoutType(CHandoutPage, CPrinterPage);
	CHandoutPage.prototype.drawContent = function (graphics) {
		const printer = this.printer;
		const presentation = printer.getPresentation();
		const handoutMaster = presentation.handoutMasters[0];
		const contentSizes = printer.getContentSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, contentSizes.width, contentSizes.height);
		const oldHandoutSettings = handoutMaster.handoutSettings;
		handoutMaster.handoutSettings = printer.getHandoutSettings();
		graphics.m_bNoPrintSlideBackground = printer.isSkipSlideBackground();
		handoutMaster.draw(graphics, this.pageIdx, this.slides, printer.isDrawSlideFrame(), printer.isDrawSlideNumber());
		graphics.m_bNoPrintSlideBackground = false;
		handoutMaster.handoutSettings = oldHandoutSettings;
		graphics.RestoreGrState();
	};

	function CNotePage(printer, note, pageIndex) {
		CPrinterPage.call(this, printer);
		this.note = note;
		this.pageIndex = pageIndex;
	}
	AscFormat.InitClassWithoutType(CNotePage, CPrinterPage);
	CNotePage.prototype.drawContent = function (graphics) {
		const printer = this.printer;
		const note = this.note;
		const contentSizes = printer.getContentSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, contentSizes.width, contentSizes.height);
		if (this.pageIndex === 0) {
			graphics.m_bNoPrintSlideBackground = printer.isSkipSlideBackground();
			note.draw(graphics);
			graphics.m_bNoPrintSlideBackground = false;
		} else {
			graphics.SaveGrState();
			note.drawForNextPrintPage(graphics);
			const transform = new AscCommon.CMatrix();
			transform.tx = CHANDOUTMASTER_HORIZONTAL_FIELD;
			transform.ty = CHANDOUTMASTER_VERTICAL_FIELD;
			graphics.transform3(transform);
			const notesShape = note.getBodyShape();
			const docContent = notesShape.getDocContent();
			docContent.Set_StartPage(0);
			docContent.Draw(this.pageIndex, graphics);
			graphics.RestoreGrState();
		}
		graphics.RestoreGrState();
	};

	function COutlinePage(printer, pageIndex) {
		CPrinterPage.call(this, printer);
		this.pageIndex = pageIndex;
	}
	AscFormat.InitClassWithoutType(COutlinePage, CPrinterPage);
	COutlinePage.prototype.draw = function (graphics) {
		CPrinterPage.prototype.draw.call(this, graphics);
		this.drawOutline(graphics);
	}
	COutlinePage.prototype.drawOutline = function (graphics) {
		const printer = this.printer;
		const pageSize = printer.getPageSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, pageSize.width, pageSize.height);
		printer.drawDecorationsByPage(graphics, this.pageIndex);
		const outlinePosition = printer.getOutlinePosition();
		const t = new AscCommon.CMatrix();
		t.tx = outlinePosition.x;
		t.ty = outlinePosition.y;
		graphics.transform3(t);
		const shape = printer.getOutlineShape();
		const content = shape.getDocContent();
		content.Set_StartPage(0);
		content.Draw(this.pageIndex, graphics);
		graphics.RestoreGrState();
	};
	COutlinePage.prototype.drawContent = function (graphics) {
		const printer = this.printer;
		const presentation = printer.getPresentation();
		const handoutMaster = presentation.handoutMasters[0];
		const contentSizes = printer.getContentSizes();
		graphics.SaveGrState();
		graphics.AddClipRect(0, 0, contentSizes.width, contentSizes.height);
		handoutMaster.drawMainPage(graphics, this.pageIndex);
		graphics.RestoreGrState();
	};

	function CCommentPage(commentsPrinter, commentsSegment, localIndex) {
		this.commentsPrinter = commentsPrinter;
		this.commentsSegment = commentsSegment;
		this.localIndex = localIndex;
	}
	CCommentPage.prototype.draw = function (graphics) {
		this.commentsPrinter.drawSegmentPage(graphics, this.commentsSegment, this.localIndex);
	};

	function CPrinterSegment(slides) {
		this.slides = slides || [];
		this.slidePages = [];
	}
	CPrinterSegment.prototype.addSlidePage = function (page) {
		this.slidePages.push(page);
	};
	CPrinterSegment.prototype.addSlidePages = function (pages) {
		for (let i = 0; i < pages.length; i += 1) {
			this.slidePages.push(pages[i]);
		}
	};
	CPrinterSegment.prototype.collectPages = function (pages, commentsPrinter) {
		for (let i = 0; i < this.slidePages.length; i += 1) {
			pages.push(this.slidePages[i]);
		}
		if (!commentsPrinter) {
			return;
		}
		const commentsSegment = commentsPrinter.getSegment(this.slides);
		if (!commentsSegment) {
			return;
		}
		const count = commentsPrinter.getSegmentPagesCount(commentsSegment);
		for (let i = 0; i < count; i += 1) {
			pages.push(new CCommentPage(commentsPrinter, commentsSegment, i));
		}
	};

	function PrintManager(presentation, printOptions) {
		this.printOptions = printOptions;
		this.presentation = presentation;
		this.specificPrinter = null;
		this.init();
	}
	PrintManager.prototype.init = function () {
		switch (this.getPrintType()) {
			case Asc.c_oAscSlidesOnPagePrintType.FullPageSlides: {
				this.specificPrinter = new SlidePrinter(this.presentation, this.printOptions);
				break;
			}
			case Asc.c_oAscSlidesOnPagePrintType.Handouts: {
				this.specificPrinter = new HandoutPrinter(this.presentation, this.printOptions);
				break;
			}
			case Asc.c_oAscSlidesOnPagePrintType.Outline: {
				this.specificPrinter = new OutlinePrinter(this.presentation, this.printOptions);
				break;
			}
			case Asc.c_oAscSlidesOnPagePrintType.SlideWithNotes: {
				this.specificPrinter = new NotesPrinter(this.presentation, this.printOptions);
				break;
			}
		}
	};
	PrintManager.prototype.getPrintType = function () {
		return this.printOptions.slidesOnPagePrintOptions.printType;
	};
	PrintManager.prototype.drawPage = function (graphics, index) {
		this.specificPrinter.drawPage(graphics, index);
	};
	PrintManager.prototype.getPagesCount = function () {
		return this.specificPrinter.getPagesCount();
	};
	PrintManager.prototype.getPageSizes = function () {
		return this.specificPrinter.getPageSizes();
	};

	function SpecificPrinter(presentation, printOptions) {
		this.presentation = presentation;
		this.printOptions = printOptions;
		this.init();
	}
	SpecificPrinter.prototype.init = function () {
		this.resetCache();
	};
	SpecificPrinter.prototype.resetCache = function () {
		this.cache = Object.assign({
			printSlides: null,
			commentsPrinter: null,
			pages: null
		}, this.getSpecificCache());
	};
	SpecificPrinter.prototype.getSpecificCache = function () {
		return {};
	};

	SpecificPrinter.prototype.getPresentation = function () {
		return this.presentation;
	}
	SpecificPrinter.prototype.getForcePrintIndexes = function () {
		const range = this.printOptions.rangeOptions;
		const presentation = this.getPresentation();
		const isDrawHiddenSlides = this.printOptions.slidePrintOptions.asc_getIsPrintHiddenSlides();
		const slides = presentation.Slides;
		let result;
		switch (range.rangeType) {
			case Asc.c_oAscPresentationRangeType.AllSlides: {
				result = [];
				for (let i = 0; i < slides.length; i += 1) {
					result.push(i);
				}
				break;
			}
			case Asc.c_oAscPresentationRangeType.SelectedSlides: {
				if (presentation.IsSlidePageMode() || presentation.IsOutlineMode()) {
					const thumbnails = presentation.DrawingDocument.m_oWordControl.Thumbnails;
					result = thumbnails.GetSelectedArray();
				} else {
					result = [];
					for (let i = 0; i < slides.length; i += 1) {
						result.push(i);
					}
				}
				break;
			}
			case Asc.c_oAscPresentationRangeType.ActiveSlide: {
				if (presentation.CurPage !== -1) {
					result = [presentation.CurPage];
				} else {
					return [];
				}
				break;
			}
			case Asc.c_oAscPresentationRangeType.CustomRange: {
				const customRange = range.customRange;
				const slides = presentation.Slides;
				const ranges = customRange.trim().split(",");
				result = [];
				for (let i = 0; i < ranges.length; i += 1) {
					const range = ranges[i].trim().split("-");
					if (range.length === 2) {
						const start = parseInt(range[0], 10) - 1;
						const end = parseInt(range[1], 10) - 1;
						if (Number.isNaN(start) || Number.isNaN(end)) {
							return [];
						}
						if (start < 0 || start >= slides.length || end < 0 || end >= slides.length) {
							return [];
						}
						if (start <= end) {
							for (let slideIndex = start; slideIndex <= end; slideIndex += 1) {
								result.push(slideIndex);
							}
						} else {
							for (let slideIndex = end; slideIndex >= start; slideIndex -= 1) {
								result.push(slideIndex);
							}
						}
					} else if (range.length === 1) {
						const slideIndex = parseInt(range[0], 10) - 1;
						if (Number.isNaN(slideIndex) || slideIndex < 0 || slideIndex >= slides.length) {
							return [];
						}
						result.push(slideIndex);
					} else {
						return [];
					}
				}
				break;
			}
		}
		if (isDrawHiddenSlides) {
			return result;
		}
		return result.filter(function (slideIndex) {
			const slide = slides[slideIndex];
			return slide.isVisible();
		});
	};
	SpecificPrinter.prototype.getPrintSlides = function () {
		if (this.cache.printSlides === null) {
			const presentation = this.getPresentation();
			const printIndexes = this.getForcePrintIndexes();
			const slides = [];
			for (let i = 0; i < printIndexes.length; i += 1) {
				slides.push(presentation.Slides[printIndexes[i]]);
			}
			this.cache.printSlides = slides;
		}

		return this.cache.printSlides;
	};
	SpecificPrinter.prototype.getContentSizes = function () {
		return null;
	};
	SpecificPrinter.prototype.getPageSizes = function () {
		const pageOptions = this.printOptions.pageOptions;
		return { width: pageOptions.width, height: pageOptions.height };
	};
	SpecificPrinter.prototype.isScaleToFitPaper = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsScaleToFitPaper();
	};
	SpecificPrinter.prototype.isSkipSlideBackground = function () {
		return !this.printOptions.slidePrintOptions.asc_getIsPrintSlideBackground();
	};
	SpecificPrinter.prototype.isPrintComments = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsPrintComments();
	};
	SpecificPrinter.prototype.getCommentsPrinter = function () {
		if (!this.cache.commentsPrinter) {
			const oThis = this;
			this.cache.commentsPrinter = new AscCommonSlide.CommentsPrinter(this.presentation, this.printOptions, function () {
				return oThis.getPageSizes();
			});
		}
		return this.cache.commentsPrinter;
	};
	SpecificPrinter.prototype.getSegments = function () {
		return null;
	};
	SpecificPrinter.prototype.getPages = function () {
		if (this.cache.pages) {
			return this.cache.pages;
		}
		const segments = this.getSegments();
		if (!segments) {
			return null;
		}
		const commentsPrinter = this.isPrintComments() ? this.getCommentsPrinter() : null;
		const pages = [];
		for (let i = 0; i < segments.length; i += 1) {
			segments[i].collectPages(pages, commentsPrinter);
		}
		this.cache.pages = pages;
		return pages;
	};
	SpecificPrinter.prototype.getPagesCount = function () {
		const pages = this.getPages();
		return pages ? pages.length : 0;
	};
	SpecificPrinter.prototype.drawPage = function (graphics, index) {
		const pages = this.getPages();
		if (!pages) {
			return;
		}
		const page = pages[index];
		if (!page) {
			return;
		}
		page.draw(graphics);
	};
	SpecificPrinter.prototype.isDrawFrame = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsFrameSlides();
	};
	SpecificPrinter.prototype.drawFrame = function (graphics) {
		const contentSizes = this.getContentSizes();
		graphics.p_color(0, 0, 0, 255);
		this.strokeRect(graphics, 0, 0, contentSizes.width, contentSizes.height);
	};
	SpecificPrinter.prototype.strokeRect = function (graphics, x, y, width, height) {
		if (graphics.AddSmartRect) {
			graphics.AddSmartRect(x, y, width, height, 0);
		} else {
			graphics.p_width(0);
			graphics._s();
			graphics._m(x, y);
			graphics._l(x + width, y);
			graphics._l(x + width, y + height);
			graphics._l(x, y + height);
			graphics._z();
			graphics.ds();
		}
	};
	SpecificPrinter.prototype.getScaleToFitPaperCoefficient = function () {
		const pageSizes = this.getPageSizes();
		const contentSizes = this.getContentSizes();
		const widthScale = (pageSizes.width - 2 * SPECIFICPRINTER_FITPAGE_MARGIN) / contentSizes.width;
		const heightScale = (pageSizes.height - 2 * SPECIFICPRINTER_FITPAGE_MARGIN) / contentSizes.height;
		return Math.min(widthScale, heightScale);
	}

	function NotesPrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(NotesPrinter, SpecificPrinter);
	NotesPrinter.prototype.getSpecificCache = function () {
		return {
			notePages: null
		};
	};
	NotesPrinter.prototype.getSegments = function () {
		const slides = this.getPrintSlides();
		const pagesByNote = this.getGroupPagesByNote(this.getNotesPages());
		const segments = [];
		for (let i = 0; i < slides.length; i += 1) {
			const slide = slides[i];
			const note = slide.notes;
			const noteId = note.GetId();
			if (!pagesByNote[noteId]) {
				continue;
			}
			const segment = new CPrinterSegment([slide]);
			segment.addSlidePages(pagesByNote[noteId]);
			segments.push(segment);
		}
		return segments;
	};
	NotesPrinter.prototype.getGroupPagesByNote = function (pages) {
		const map = {};
		for (let i = 0; i < pages.length; i += 1) {
			const page = pages[i];
			const noteId = page.note.GetId();
			if (!map[noteId]) {
				map[noteId] = [];
			}
			map[noteId].push(page);
		}
		return map;
	};
	NotesPrinter.prototype.getNotesPages = function () {
		if (this.cache.notePages === null) {
			const pages = [];
			const printer = this;
			const presentation = this.getPresentation();
			const notesPageHeight = presentation.GetNotesHeightMM() - CHANDOUTMASTER_VERTICAL_FIELD * 2;
			const notesPageWidth = presentation.GetNotesWidthMM() - CHANDOUTMASTER_HORIZONTAL_FIELD * 2;
			const printSlides = this.getPrintSlides();
			for (let i = 0; i < printSlides.length; i += 1) {
				const slide = printSlides[i];
				const note = slide.notes;
				const shape = note.getBodyShape();
				note.recalculate();
				if (shape) {
					note.cSld.forEachSp(function (shape) {
						shape.setRecalculateInfo();
						shape.recalculate()
					});
					shape.setRecalculateInfo();
					shape.recalculate();
					var oDocContent = shape.getDocContent();
					if(oDocContent) {
						oDocContent.CalculateAllFields();
						oDocContent.Reset(0, 0, shape.extX, 20000);
						var CurPage = 0;
						var RecalcResult = recalcresult2_NextPage;
						while (recalcresult2_End !== RecalcResult) {
							const callback = function () {
								pages.push(new CNotePage(printer, note, CurPage));
								RecalcResult = oDocContent.Recalculate_Page(CurPage++, true);
							};
							if (CurPage === 0) {
								executeWithContentLimits(callback, oDocContent, 0, 0, shape.extX, shape.extY);
							} else {
								executeWithContentLimits(callback, oDocContent, 0, 0, notesPageWidth, notesPageHeight);
							}
						}
						shape.contentWidth = shape.extX;
						shape.contentHeight = 20000;
					}
				}
			}
			this.cache.notePages = pages;
		}
		return this.cache.notePages;
	};
	NotesPrinter.prototype.getContentSizes = function () {
		const presentation = this.getPresentation();
		return { width: presentation.GetNotesWidthMM(), height: presentation.GetNotesHeightMM() };
	};

	function OutlinePrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(OutlinePrinter, SpecificPrinter);

	OutlinePrinter.prototype.getSpecificCache = function () {
		return {
			outlineView: null,
			decorationsByPage: null,
			outlinePosition: null
		};
	};
	OutlinePrinter.prototype.getPages = function () {
		if (this.cache.pages) {
			return this.cache.pages;
		}
		const outlineShape = this.getOutlineShape();
		const docContent = outlineShape.getDocContent();
		const pages = [];
		for (let i = 0; i < docContent.Pages.length; i += 1) {
			pages.push(new COutlinePage(this, i));
		}
		this.cache.pages = pages;
		return pages;
	};
	OutlinePrinter.prototype.getOutlinePosition = function () {
		if (this.cache.outlinePosition === null) {
			const contentSizes = this.getContentSizes();
			const pageSizes = this.getPageSizes();
			const scaleCoefficient = this.isScaleToFitPaper() ? this.getScaleToFitPaperCoefficient() : 1;
			const horizontalFieldSize = CHANDOUTMASTER_HORIZONTAL_FIELD;
			const verticalFieldSize = CHANDOUTMASTER_VERTICAL_FIELD;
			const scaledContentWidth = contentSizes.width * scaleCoefficient;
			const scaledContentHeight = contentSizes.height * scaleCoefficient;
			const outlineWidth = scaledContentWidth - horizontalFieldSize * 2;
			const textWidth = outlineWidth - OUTLINEPRINTER_DECORATIONS_OFFSET_LEFT;
			const height = scaledContentHeight - verticalFieldSize * 2;

			const x = horizontalFieldSize + OUTLINEPRINTER_DECORATIONS_OFFSET_LEFT + (pageSizes.width - scaledContentWidth) / 2;
			const y = verticalFieldSize + (pageSizes.height - scaledContentHeight) / 2;
			this.cache.outlinePosition = {
				decorX: horizontalFieldSize,
				x: x,
				y: y,
				width: textWidth,
				height: height
			}
		}
		return this.cache.outlinePosition;
	};
	OutlinePrinter.prototype.getOutlineView = function () {
		if (this.cache.outlineView === null) {
			const presentation = this.getPresentation();
			const outlineSlides = [];
			const printSlides = this.getPrintSlides();
			for (let i = 0; i < printSlides.length; i++) {
				const slide = printSlides[i];
				outlineSlides.push(slide.getOutlineSlide());
			}
			const outlinePosition = this.getOutlinePosition();
			const outlineView = new AscCommonSlide.OutlineView(presentation.Api);
			outlineView.createOutlineShape(outlineSlides, outlinePosition.width, outlinePosition.height);
			this.cache.outlineView = outlineView;
			this.recalculateOutlinePages();
		}
		return this.cache.outlineView;
	};
	OutlinePrinter.prototype.getOutlineShape = function () {
		const outlineView = this.getOutlineView();
		return outlineView.outlineShape;
	};
	OutlinePrinter.prototype.recalculateOutlinePages = function () {
		const outlineShape = this.getOutlineShape();
		const content = outlineShape.getDocContent();
		let recalResult = recalcresult2_NextPage;
		let curPage = 0;
		while (recalcresult2_End !== recalResult)
			executeWithContentLimits(function () {
				recalResult = content.Recalculate_Page(curPage++, true);
			}, content, 0, 0, outlineShape.extX, outlineShape.extY);
	};
	OutlinePrinter.prototype.getParagraphPage = function (paragraph) {
		const startPage = paragraph.PageNum;
		for (let i = 0; i < paragraph.Pages.length; i += 1) {
			const page = paragraph.Pages[i];
			if (page.EndLine >= 0) {
				return startPage + i;
			}
		}
		return startPage;
	};
	OutlinePrinter.prototype.getDecorationsByPage = function () {
		if (this.cache.decorationsByPage === null) {
			this.cache.decorationsByPage = {
				titleCache: {},
				contentCache: {}
			};
			const titleCache = this.cache.decorationsByPage.titleCache;
			const contentCache = this.cache.decorationsByPage.contentCache;
			const outlineView = this.getOutlineView();
			const paragraphMap = outlineView.outlineInfo.getOutlineParagraphToInfoMap();
			for (let outlineId in paragraphMap) {
				const info = paragraphMap[outlineId];
				const titleOutlineParagraph = info.outlineParagraph;
				const titlePageNum = this.getParagraphPage(titleOutlineParagraph);
				if (!titleCache[titlePageNum]) {
					titleCache[titlePageNum] = [];
				}
				titleCache[titlePageNum].push(info);
				const contentShapeInfoMap = info.getContentShapeInfoMap(outlineView);
				for (let contentShapeInfoId in contentShapeInfoMap) {
					const contentShapeInfo = contentShapeInfoMap[contentShapeInfoId];
					const contentOutlineParagraph = contentShapeInfo.getOutlineParagraph(outlineView);
					const contentPageNum = this.getParagraphPage(contentOutlineParagraph);
					if (!contentCache[contentPageNum]) {
						contentCache[contentPageNum] = [];
					}
					contentCache[contentPageNum].push(contentShapeInfo);
				}
			}
		}
		return this.cache.decorationsByPage;
	};
	OutlinePrinter.prototype.drawDecorationsByPage = function (graphics, index) {
		const t = new AscCommon.CMatrix();
		const position = this.getOutlinePosition();
		t.tx = position.decorX;
		t.ty = position.y;
		const decorationsInfo = this.getDecorationsByPage();
		const outlineView = this.getOutlineView();

		const decoratorDrawer = new AscCommonSlide.DecoratorDrawer(outlineView, t);
		const titlePageInfos = decorationsInfo.titleCache[index];
		const contentPageInfos = decorationsInfo.contentCache[index];
		if (titlePageInfos) {
			for (let i = 0; i < titlePageInfos.length; i += 1) {
				const info = titlePageInfos[i];
				decoratorDrawer.drawTitleDecorations(graphics, info);
			}
		}
		if (contentPageInfos) {
			for (let i = 0; i < contentPageInfos.length; i += 1) {
				const info = contentPageInfos[i];
				decoratorDrawer.drawContentDecorations(graphics, info);
			}
		}
	};
	OutlinePrinter.prototype.getContentSizes = function () {
		const presentation = this.getPresentation();
		return { width: presentation.GetNotesWidthMM(), height: presentation.GetNotesHeightMM() };
	};
	OutlinePrinter.prototype.isDrawFrame = function () {
		return false;
	};

	function HandoutPrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(HandoutPrinter, SpecificPrinter);
	HandoutPrinter.prototype.getSegments = function () {
		const allSlides = this.getPrintSlides();
		const slidesOnPageCount = this.getSlidesOnPageCount();
		const segments = [];
		for (let i = 0; i < allSlides.length; i += slidesOnPageCount) {
			const slides = allSlides.slice(i, i + slidesOnPageCount);
			const segment = new CPrinterSegment(slides);
			segment.addSlidePage(new CHandoutPage(this, segments.length, slides));
			segments.push(segment);
		}
		return segments;
	};
	HandoutPrinter.prototype.getSlidesOnPageCount = function () {
		return this.printOptions.slidesOnPagePrintOptions.slidesCount;
	};
	HandoutPrinter.prototype.getHandoutSettings = function () {
		const printOptions = this.printOptions;
		return AscCommonSlide.HandoutSettings.fromPrinterSettings(printOptions.slidesOnPagePrintOptions);
	};
	HandoutPrinter.prototype.getContentSizes = function () {
		const presentation = this.getPresentation();
		return { width: presentation.GetNotesWidthMM(), height: presentation.GetNotesHeightMM() };
	};
	HandoutPrinter.prototype.isDrawFrame = function () {
		return false;
	};
	HandoutPrinter.prototype.isDrawSlideFrame = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsFrameSlides();
	};
	HandoutPrinter.prototype.isDrawSlideNumber = function () {
		return this.printOptions.slidesOnPageOptions.asc_getIsPrintSlideNumbersOnHandouts();
	};

	function SlidePrinter(presentation, printOptions) {
		SpecificPrinter.call(this, presentation, printOptions);
	}
	AscFormat.InitClassWithoutType(SlidePrinter, SpecificPrinter);
	SlidePrinter.prototype.getSegments = function () {
		const slides = this.getPrintSlides();
		const segments = [];
		for (let i = 0; i < slides.length; i += 1) {
			const slide = slides[i];
			const segment = new CPrinterSegment([slide]);
			segment.addSlidePage(new CSlidePage(this, slide));
			segments.push(segment);
		}
		return segments;
	};
	SlidePrinter.prototype.getContentSizes = function () {
		const presentation = this.getPresentation();
		return { width: presentation.GetWidthMM(), height: presentation.GetHeightMM() };
	};
	SlidePrinter.prototype.getPageSizes = function () {
		const pageOptions = this.printOptions.pageOptions;
		const contentSizes = this.getContentSizes();
		const isSlideLandscape = contentSizes.width > contentSizes.height;
		const isPageLandscape = pageOptions.width > pageOptions.height;
		if (isSlideLandscape !== isPageLandscape) {
			return { width: pageOptions.height, height: pageOptions.width };
		}
		return { width: pageOptions.width, height: pageOptions.height };
	};

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].PrintManager = PrintManager;
})();
