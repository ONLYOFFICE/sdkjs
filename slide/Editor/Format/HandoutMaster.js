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

"use strict";

(function () {
	const History = AscCommon.History;

	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetTheme] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetHF] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetClrMap] = AscDFH.CChangesDrawingsObject;

	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetTheme] = function (oClass, value) {
		oClass.Theme = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetHF] = function (oClass, value) {
		oClass.hf = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetClrMap] = function (oClass, value) {
		oClass.clrMap = value;
	};

	const HANDOUT_ALIGNMENT_HORIZONTAL = 0;
	const HANDOUT_ALIGNMENT_VERTICAL = 1;

	const OUTLINE_LAYOUT_COUNT = -1;
	function HandoutSettings() {
		this.slideCount = 6;
		this.align = HANDOUT_ALIGNMENT_HORIZONTAL;
	}
	HandoutSettings.fromPrinterSettings = function (slidesOnPagePrintOptions) {
		const handoutSettings = new HandoutSettings();
		handoutSettings.setSlideCount(slidesOnPagePrintOptions.slidesCount);
		handoutSettings.setAlign(slidesOnPagePrintOptions.arrangementType === Asc.c_oAscSlidesOnPageArrangementType.Horizontal ? HANDOUT_ALIGNMENT_HORIZONTAL : HANDOUT_ALIGNMENT_VERTICAL);
		return handoutSettings;
	};
	HandoutSettings.prototype.setSlideCount = function (count) {
		this.slideCount = count;
	};
	HandoutSettings.prototype.setAlign = function (align) {
		this.align = align;
	};
	HandoutSettings.prototype.isOutlineLayout = function () {
		return this.slideCount === OUTLINE_LAYOUT_COUNT;
	};

	const CHANDOUTMASTER_GAP = 10;
	const CHANDOUTMASTER_HORIZONTAL_FIELD = 20;
	const CHANDOUTMASTER_VERTICAL_FIELD = 20;
	function CHandoutMaster() {
		AscCommonSlide.SlideBase.call(this);
		this.clrMap = new AscFormat.ClrMap();
		this.hf = null;

		this.handoutSettings = new HandoutSettings();

		this.Theme = null;
		this.kind = AscFormat.TYPE_KIND.HANDOUT_MASTER;
		this.setSlideSize(this.presentation.GetNotesWidthMM(), this.presentation.GetNotesHeightMM());
		this.recalcInfo = {
			recalculateBackground: true,
			recalculateSpTree    : true
		};
	}

	AscFormat.InitClass(CHandoutMaster, AscCommonSlide.SlideBase, AscDFH.historyitem_type_HandoutMaster);
	CHandoutMaster.prototype.fromPPTY = function (pReader) {
		var oStream = pReader.stream;
		var nStart = oStream.cur;
		oStream.Skip2(1);
		var nEnd = nStart + oStream.GetULong() + 4;
		this.readChildren(nEnd, pReader);
		oStream.Seek2(nEnd);
	};
	CHandoutMaster.prototype.readChild = function (nType, pReader) {
		var s = pReader.stream;
			switch (nType) {
				case 0: {
					pReader.ReadCSld(this.cSld);
					break;
				}
				case 1: {
					var clrMap = new AscFormat.ClrMap();
					pReader.ReadClrMap(clrMap);
					this.setClrMap(clrMap);
					break;
				}
				case 2: {
					this.setHF(pReader.ReadHF());
					break;
				}
				default: {
					s.SkipRecord();
					break;
				}
			}
	};
	CHandoutMaster.prototype.privateWriteAttributes = undefined;
	CHandoutMaster.prototype.writeChildren = function (pWriter) {
		pWriter.StartRecord(AscCommon.c_oMainTables.HandoutMasters);

		pWriter.WriteRecord1(0, this.cSld, pWriter.WriteCSld);
		pWriter.WriteRecord1(1, this.clrMap, pWriter.WriteClrMap);
		pWriter.WriteRecord2(2, this.hf, pWriter.WriteHF);

		pWriter.EndRecord();
	};
	CHandoutMaster.prototype.recalcSpTree = function () {
		this.recalcInfo.recalculateSpTree = true;
	};
	CHandoutMaster.prototype.recalcBackground = function () {
		this.recalcInfo.recalculateBackground = true;
	};
	CHandoutMaster.prototype.getObjectType = function () {
		return AscDFH.historyitem_type_HandoutMaster;
	};

	CHandoutMaster.prototype.setTheme = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_HandoutMasterSetTheme, this.Theme, pr));
		this.Theme = pr;
	};
	CHandoutMaster.prototype.setClrMap = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_HandoutMasterSetClrMap, this.clrMap, pr));
		this.clrMap = pr;
	};

	CHandoutMaster.prototype.setHF = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_HandoutMasterSetHF, this.hf, pr));
		this.hf = pr;
	};

	CHandoutMaster.prototype.removeFromSpTreeById = function (id) {
		for (var i = this.cSld.spTree.length - 1; i > -1; --i) {
			if (this.cSld.spTree[i].Get_Id() === id) {
				this.removeFromSpTreeByPos(i);
			}
		}
	};

	CHandoutMaster.prototype.getAllFonts = function (fonts) {
		var i;
		if (this.Theme) {
			this.Theme.Document_Get_AllFontNames(fonts);
		}

		for (i = 0; i < this.cSld.spTree.length; ++i) {
			if (typeof this.cSld.spTree[i].getAllFonts === "function")
				this.cSld.spTree[i].getAllFonts(fonts);
		}
	};

	CHandoutMaster.prototype.createDuplicate = function (IdMap) {
		var oIdMap = IdMap || {};
		var oPr = new AscFormat.CCopyObjectProperties();
		oPr.idMap = oIdMap;
		var i;
		var copy = new CHandoutMaster();
		if (this.clrMap) {
			copy.setClrMap(this.clrMap.createDuplicate());
		}
		if (typeof this.cSld.name === "string" && this.cSld.name.length > 0) {
			copy.setCSldName(this.cSld.name);
		}
		if (this.cSld.Bg) {
			copy.changeBackground(this.cSld.Bg.createFullCopy());
		}
		for (i = 0; i < this.cSld.spTree.length; ++i) {
			var _copy = this.cSld.spTree[i].copy(oPr);
			oIdMap[this.cSld.spTree[i].Id] = _copy.Id;
			copy.addToSpTreeToPos(copy.cSld.spTree.length, _copy);
			copy.cSld.spTree[copy.cSld.spTree.length - 1].setParent2(copy);
		}
		if (this.hf) {
			copy.setHF(this.hf.createDuplicate());
		}
		return copy;
	};
	CHandoutMaster.prototype.getThemeIndex = function () {
		let aMasters = Asc.editor.WordControl.m_oLogicDocument.notesMasters;
		for (let nIdx = 0; nIdx < aMasters.length; ++nIdx) {
			if (aMasters[nIdx] === this) {
				return -nIdx - 1;
			}
		}
		return 0;
	};
	CHandoutMaster.prototype.drawNoPlaceholders = function (graphics, slide) {

	};
	CHandoutMaster.prototype.draw = function (graphics, slideNum, handoutSlides, isDrawFrame, isDrawSlideNumber) {
		this.drawMainPage(graphics, slideNum);
		this.drawPlaceholders(graphics, handoutSlides, isDrawFrame, isDrawSlideNumber);
	};
	CHandoutMaster.prototype.drawMainPage = function (graphics, slideNum) {
		if (AscFormat.isRealNumber(slideNum)) {
			if (slideNum !== this.lastRecalcSlideIndex) {
				this.recalcSpTree();
				this.lastRecalcSlideIndex = slideNum;
				this.cSld.refreshAllContentsFields(true);
			}
		} else {
			if (-1 !== this.lastRecalcSlideIndex) {
				this.recalcSpTree();
				this.lastRecalcSlideIndex = -1;
				this.cSld.refreshAllContentsFields(true);

			}
		}
		this.recalculate();

		DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
		this.cSld.forEachSp(function (oSp) {
			if (!AscCommon.IsHiddenObj(oSp)) {
				oSp.draw(graphics);
			}
		});
	};
	CHandoutMaster.prototype.getGap = function (paperSize, slideSize, scale, repeatCount) {
		return repeatCount === 1 ? 0 : (paperSize - slideSize * repeatCount * scale) / (repeatCount - 1);
	}
	CHandoutMaster.prototype.getAlign = function () {
		if (this.handoutSettings.slideCount === 3) {
			const pageWidth = this.presentation.GetNotesWidthMM();
			const pageHeight = this.presentation.GetNotesHeightMM();
			if (pageWidth >= pageHeight) {
				return HANDOUT_ALIGNMENT_VERTICAL;
			} else {
				return HANDOUT_ALIGNMENT_HORIZONTAL;
			}
		}
		return this.handoutSettings.align;
	};
	CHandoutMaster.prototype.getSlidesCount = function () {
		return this.handoutSettings.slideCount;
	};
	CHandoutMaster.prototype.getOutlineHandouts = function () {
		const pageWidth = this.presentation.GetNotesWidthMM();
		const pageHeight = this.presentation.GetNotesHeightMM();
		const x = CHANDOUTMASTER_HORIZONTAL_FIELD;
		const y = CHANDOUTMASTER_VERTICAL_FIELD;
		const width = pageWidth - 2 * CHANDOUTMASTER_HORIZONTAL_FIELD;
		const height = pageHeight - 2 * CHANDOUTMASTER_VERTICAL_FIELD;
		return [new SlidePlaceholder(x, y, width, height)];
	};
	CHandoutMaster.prototype.getSlideHandouts = function (slides) {
		let align = this.getAlign();
		const handouts = [];
		let slidesCount = this.getSlidesCount();
		let isDrawWithSlideTextPlaceholder = false;
		if (slidesCount === 3) {
			isDrawWithSlideTextPlaceholder = true;
			slidesCount = 6;
		}
		const scaledFieldSize = CHANDOUTMASTER_HORIZONTAL_FIELD;
		let countSlidesOnRow = this.getSlidesCountOnRow(slidesCount);
		let rowsCount = Math.floor(slidesCount / countSlidesOnRow);
		if (this.presentation.GetNotesWidthMM() < this.presentation.GetNotesHeightMM()) {
			const temp = countSlidesOnRow;
			countSlidesOnRow = rowsCount;
			rowsCount = temp;
		}
		const w_mm = this.presentation.GetWidthMM();
		const h_mm = this.presentation.GetHeightMM();
		const paperWidth = this.presentation.GetNotesWidthMM() - scaledFieldSize * 2;
		const paperHeight = this.presentation.GetNotesHeightMM() - CHANDOUTMASTER_VERTICAL_FIELD * 2;
		const slidesWidth = w_mm * countSlidesOnRow;
		const slidesHeight = h_mm * rowsCount;
		const resultWidthWithMaxGap = (countSlidesOnRow - 1) * CHANDOUTMASTER_GAP;
		const resultHeightWithMaxGap = (rowsCount - 1) * CHANDOUTMASTER_GAP;
		const adaptSlideWidthScale = (paperWidth - resultWidthWithMaxGap) / slidesWidth;
		const adaptSlideHeightScale = (paperHeight - resultHeightWithMaxGap) / slidesHeight;

		let slideScale = Math.min(adaptSlideWidthScale, adaptSlideHeightScale);
		let horizontalGap = this.getGap(paperWidth, w_mm, slideScale, countSlidesOnRow);
		let verticalGap = this.getGap(paperHeight, h_mm, slideScale, rowsCount);

		const scaledPresentationWidth = w_mm * slideScale;
		const scaledPresentationHeight = h_mm * slideScale;

		const resultWidth = countSlidesOnRow * w_mm * slideScale + (countSlidesOnRow - 1) * horizontalGap;
		const resultHeight = rowsCount * h_mm * slideScale + (rowsCount - 1) * verticalGap;
		const startX = scaledFieldSize + (paperWidth - resultWidth) / 2;
		const startY = scaledFieldSize + (paperHeight - resultHeight) / 2;

		let slideIndex = 0;
		if (align === HANDOUT_ALIGNMENT_HORIZONTAL) {
			for (let i = 0; i < rowsCount; i += 1) {
				const slideY = startY + i * verticalGap + h_mm * i * slideScale;
				for (let j = 0; j < countSlidesOnRow; j += 1) {
					const slideX = startX + j * w_mm * slideScale + j * horizontalGap;
					if (isDrawWithSlideTextPlaceholder && handouts.length % 2 === 1) {
						handouts.push(new TextPlaceholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight));
					} else {
						let slide = null;
						if (slides) {
							slide = slides[slideIndex];
							if (!slide) {
								return handouts;
							}
							slideIndex += 1;
						}
						handouts.push(new SlidePlaceholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight, slide));
					}
				}
			}
		} else {
			for (let i = 0; i < countSlidesOnRow; i += 1) {
				const slideX = startX + i * horizontalGap + w_mm * i * slideScale;
				for (let j = 0; j < rowsCount; j += 1) {
					const slideY = startY + j * h_mm * slideScale + j * verticalGap;
					if (isDrawWithSlideTextPlaceholder && handouts.length % 2 === 1) {
						handouts.push(new TextPlaceholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight));
					} else {
						let slide = null;
						if (slides) {
							slide = slides[slideIndex];
							if (!slide) {
								return handouts;
							}
							slideIndex += 1;
						}
						handouts.push(new SlidePlaceholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight, slide));
					}
				}
			}
		}
		return handouts;
	};
	CHandoutMaster.prototype.getHandouts = function (slides) {
		if (this.isOutlineLayout()) {
			return this.getOutlineHandouts();
		} else {
			return this.getSlideHandouts(slides);
		}
	};
	CHandoutMaster.prototype.isOutlineLayout = function () {
		return this.handoutSettings.isOutlineLayout();
	};
	CHandoutMaster.prototype.drawPlaceholders = function (g, handoutSlides, isDrawFrame, isDrawSlideNumber) {
		const handouts = this.getHandouts(handoutSlides);
		for (let i = 0; i < handouts.length; i += 1) {
			const handout = handouts[i];
			handout.draw(g, isDrawFrame, isDrawSlideNumber);
		}
	};
	CHandoutMaster.prototype.getSlidesCountOnRow = function (slidesCount) {
		if (slidesCount % 3 === 0) {
			return 3;
		} else if (slidesCount % 2 === 0) {
			return 2;
		}
		return 1;
	}
	CHandoutMaster.prototype.getDrawingsForController = function () {
		return this.cSld.spTree;
	};
	CHandoutMaster.prototype.setSlideSize = function (nW, nH) {
		this.Width = nW;
		this.Height = nH;
	};
	CHandoutMaster.prototype.getTheme = function () {
		return this.Theme || null;
	};
	CHandoutMaster.prototype.getColorMap = function () {
		return this.Get_ColorMap();
	};
	CHandoutMaster.prototype.recalculateBackground = function () {
		let backgroundFill = null;
		let rgba = {R: 0, G: 0, B: 0, A: 255};

		const layout = null;
		const master = this;
		const theme = this.Theme;
		if (this.cSld.Bg != null) {
			const ret = this.cSld.recalculateBackground(theme, this, layout, master, rgba);
			backgroundFill = ret.backFill;
			rgba = ret.RGBA || rgba;
		}
		if (backgroundFill != null)
			backgroundFill.calculate(theme, this, layout, master, rgba);

		this.backgroundFill = backgroundFill;
	};
	CHandoutMaster.prototype.checkSlideColorScheme = function () {
		this.recalcSpTree();
		this.recalcBackground();
		for (let i = 0; i < this.cSld.spTree.length; ++i) {
			const shape = this.cSld.spTree[i];
			if (!shape.isPlaceholder()) {
				shape.handleUpdateFill();
				shape.handleUpdateLn();
			}
		}
	};
	CHandoutMaster.prototype.createFontMap = function (oFontsMap, oCheckedMap, isNoPh) {
		if (oCheckedMap[this.Get_Id()]) {
			return;
		}
		const spTree = this.cSld.spTree;
		for (let idx = 0; idx < spTree.length; ++idx) {
			const oSp = spTree[idx];
			if (isNoPh) {
				if (oSp.isPlaceholder()) {
					continue;
				}
			}
			oSp.createFontMap(oFontsMap);
		}
		oCheckedMap[this.Get_Id()] = this;
	};
	CHandoutMaster.prototype.recalculate = function () {
		if (this.recalcInfo.recalculateSpTree) {
			this.recalcInfo.recalculateSpTree = false;
			for (let i = 0; i < this.cSld.spTree.length; ++i) {
				this.cSld.spTree[i].recalculate();
			}
		}
		if (this.recalcInfo.recalculateBackground) {
			this.recalculateBackground();
			this.recalcInfo.recalculateBackground = false;
		}
	};
	CHandoutMaster.prototype.isAnimated = function () {
		return false;
	};
	CHandoutMaster.prototype.isLockedObject = function () {
		return false;
	};
	//todo
	CHandoutMaster.prototype.isSlide = function () {
		return false;
	};
	CHandoutMaster.prototype.isLayout = function () {
		return false;
	};
	CHandoutMaster.prototype.isMaster = function () {
		return false;
	};
	CHandoutMaster.prototype.getParentObjects = function () {
		return {
			presentation: Asc.editor.private_GetLogicDocument(),
			master      : this,
			layout      : null,
			slide       : null
		};
	};
	CHandoutMaster.prototype.getWorksheet = function () {
		return null;
	};
	CHandoutMaster.prototype.Get_ColorMap = function () {
		if (this.clrMap) {
			return this.clrMap;
		}
		return AscFormat.GetDefaultColorMap();
	};
	CHandoutMaster.prototype.IsUseInDocument = function () {
		let oPresentation = Asc.editor.private_GetLogicDocument();
		if (!oPresentation) return false;
		for (let nIdx = 0; nIdx < oPresentation.handoutMasters.length; ++nIdx) {
			if (oPresentation.handoutMasters[nIdx] === this) {
				return true;
			}
		}
		return false;
	};
	CHandoutMaster.prototype.IsUseInSlides = function () {
		return false;
	};
	CHandoutMaster.prototype.isPreserve = function () {
		return false;
	};
	CHandoutMaster.prototype.shapeRemove = function (pos, count) {
		if (pos > -1 && pos < this.cSld.spTree.length) {
			this.cSld.removeFromSpTree(pos, count);
		}
	};
	CHandoutMaster.prototype.removeFromSpTreeByPos = function (pos) {
		this.shapeRemove(pos, 1);
	};
	CHandoutMaster.prototype.shapeAdd = function (pos, item) {
		this.checkDrawingUniNvPr(item);
		var _pos = (AscFormat.isRealNumber(pos) && pos > -1 && pos <= this.cSld.spTree.length) ? pos : this.cSld.spTree.length;
		this.cSld.checkAndAssignName(item);
		this.cSld.addToSpTree(_pos, item);
		item.setParent2(this);
	};
	CHandoutMaster.prototype.addToSpTreeToPos = function (pos, item) {
		this.shapeAdd(pos, item);
	};
	CHandoutMaster.prototype.handleCSldRecalc = function (data) {
		if (data) {
			switch (data.Type) {
				case AscDFH.historyitem_CSld_SetBg:
					this.recalcInfo.recalculateBackground = true;
					break;
			}
			this.addToRecalculate();
		}
	};
	CHandoutMaster.prototype.setSlideCount = function (count) {
		this.handoutSettings.setSlideCount(count);
	};
	CHandoutMaster.prototype.getSlideCount = function (count) {
		return this.handoutSettings.slideCount;
	};

	CHandoutMaster.prototype.collectRedrawSlides = function (redrawSlides, force) {
		redrawSlides.addRedrawSlide(this);
	};
	CHandoutMaster.prototype.isHandoutMaster = function () {
		return true;
	};
	CHandoutMaster.prototype.Write_ToBinary2 = AscCommonSlide.SlideBase.prototype.writeMasterBinary;
	CHandoutMaster.prototype.Read_FromBinary2 = AscCommonSlide.SlideBase.prototype.readMasterBinary;
	CHandoutMaster.prototype.isMasterLikeObject = function () {
		return true;
	};

	function PlaceholderBase(x, y, width, height) {
		this.x = x;
		this.y = y;
		this.width = width;
		this.height = height;
	}

	PlaceholderBase.prototype.draw = function (g, isDrawFrame, isDrawSlideNumber) {

	};

	function SlidePlaceholder(x, y, width, height, slide) {
		PlaceholderBase.call(this, x, y, width, height);
		this.slide = slide || null;
	}

	AscFormat.InitClassWithoutType(SlidePlaceholder, PlaceholderBase);
	SlidePlaceholder.prototype.draw = function (graphics, isDrawFrame, isDrawSlideNumber) {
		if (this.slide) {
			this.drawWithSlide(graphics, isDrawFrame, isDrawSlideNumber);
		} else {
			this.drawPlaceholder(graphics);
		}
	};
	SlidePlaceholder.prototype.strokeRect = function (graphics, x, y, width, height) {
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
	SlidePlaceholder.prototype.drawPlaceholder = function (graphics) {
		if (graphics.isBoundsChecker()) {
			return;
		}
		graphics.SaveGrState();
		const transform = new AscCommon.CMatrix();
		transform.tx = this.x;
		transform.ty = this.y;
		graphics.transform3(transform);
		graphics.p_color(127, 127, 127, 255);
		graphics.p_dash(AscCommon.DashPatternPresets[10].slice());
		this.strokeRect(graphics, 0, 0, this.width, this.height);
		graphics.p_dash(null);
		graphics.RestoreGrState();
	};
	SlidePlaceholder.prototype.drawWithSlide = function (graphics, isDrawFrame, isDrawSlideNumber) {
		if (graphics.isBoundsChecker()) {
			return;
		}
		graphics.SaveGrState();
		const baseTransform = graphics.GetBaseTransform();
		const slide = this.slide;
		const presentation = slide.presentation;
		const slideWidth = presentation.GetWidthMM();
		const slideHeight = presentation.GetHeightMM();
		const scaleCoefficient = this.width / slideWidth;

		const localTransform = new AscCommon.CMatrix();
		localTransform.Scale(scaleCoefficient, scaleCoefficient);
		localTransform.Translate(this.x, this.y);

		const composedTransform = baseTransform ? baseTransform.CreateDublicate() : new AscCommon.CMatrix();
		composedTransform.Multiply(localTransform, AscCommon.MATRIX_ORDER_PREPEND);

		graphics.SetBaseTransform(composedTransform);
		graphics.reset();
		graphics.AddClipRect(0, 0, slideWidth, slideHeight);
		slide.draw(graphics);
		graphics.SetBaseTransform(baseTransform);
		graphics.reset();
		graphics.RestoreGrState();
		if (isDrawFrame) {
			this.drawFrame(graphics);
		}
		if (isDrawSlideNumber) {
			this.drawSlideNumber(graphics);
		}
	};
	SlidePlaceholder.prototype.drawFrame = function (graphics) {
		graphics.p_color(0, 0, 0, 255);
		this.strokeRect(graphics, this.x, this.y, this.width, this.height);
	};
	SlidePlaceholder.prototype.drawSlideNumber = function (graphics) {
		AscFormat.ExecuteNoHistory(function () {
			const shape = this.getNumberShape();
			const content = shape.txBody.content;
			shape.txBody.replaceContentFitText("" + (this.slide.num + 1));
			content.ApplyToAll = true;
			content.AddToParagraph(new AscCommonWord.ParaTextPr({
				FontSize  : 13,
				FontFamily: {Name: "Arial", Index: -1},
				Color     : AscWord.BLACK_COLOR
			}), false);
			content.ApplyToAll = false;
			shape.recalculateContent();
			const transform = shape.transformText;
			transform.tx = this.x;
			transform.ty = this.y + this.height + 1;
			shape.draw(graphics);
		}, this, []);
	};

	SlidePlaceholder.prototype.getNumberShape = function () {
		return AscFormat.ExecuteNoHistory(function () {
			const shape = new AscFormat.CShape();
			shape.setBDeleted(false);
			shape.extX = this.width;
			shape.extY = this.height;
			shape.createTextBody();
			const bodyPr = new AscFormat.CBodyPr();
			bodyPr.setInsets(0, 0, 0, 0);
			shape.txBody.setBodyPr(bodyPr);
			return shape;
		}, this, []);

	};
	SlidePlaceholder.prototype.drawText = function (graphics, text) {

	};

	const TEXTPLACEHOLDER_STEP_COUNT = 7;

	function TextPlaceholder(x, y, width, height) {
		PlaceholderBase.call(this, x, y, width, height);
	}

	AscFormat.InitClassWithoutType(TextPlaceholder, PlaceholderBase);
	TextPlaceholder.prototype.draw = function (graphics, isDrawFrame) {
		if (graphics.isBoundsChecker()) {
			return;
		}
		const step = this.height / TEXTPLACEHOLDER_STEP_COUNT;
		graphics.SaveGrState();
		const transform = new AscCommon.CMatrix();
		transform.tx = this.x;
		transform.ty = this.y;
		graphics.transform3(transform);
		graphics.p_color(0, 0, 0, 255);
		for (let i = 0; i < TEXTPLACEHOLDER_STEP_COUNT; i++) {
			graphics.drawHorLine(AscCommon.c_oAscLineDrawingRule.Center, step * (i + 1), 0, this.width, 0);
		}
		graphics.RestoreGrState();
	};

	function createHandoutMaster() {
		var oHandoutMaster = new CHandoutMaster();
		var oBG = new AscFormat.CBg();
		var oBgRef = new AscFormat.StyleRef();
		oBgRef.idx = 1001;
		var oUniColor = new AscFormat.CUniColor();
		oUniColor.color = new AscFormat.CSchemeColor();
		oUniColor.color.id = 6;
		oBgRef.Color = oUniColor;
		oBG.bgRef = oBgRef;
		oHandoutMaster.changeBackground(oBG);

		AscCommonSlide.addHeaderShape(oHandoutMaster);
		AscCommonSlide.addDateShape(oHandoutMaster);
		AscCommonSlide.addFooterShape(oHandoutMaster);
		AscCommonSlide.addNumberShape(oHandoutMaster);
		//clrMap
		oHandoutMaster.clrMap.setClr(0, 0);
		oHandoutMaster.clrMap.setClr(1, 1);
		oHandoutMaster.clrMap.setClr(2, 2);
		oHandoutMaster.clrMap.setClr(3, 3);
		oHandoutMaster.clrMap.setClr(4, 4);
		oHandoutMaster.clrMap.setClr(5, 5);
		oHandoutMaster.clrMap.setClr(10, 10);
		oHandoutMaster.clrMap.setClr(11, 11);
		oHandoutMaster.clrMap.setClr(6, 12);
		oHandoutMaster.clrMap.setClr(7, 13);
		oHandoutMaster.clrMap.setClr(15, 8);
		oHandoutMaster.clrMap.setClr(16, 9);
		return oHandoutMaster;
	}

	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonSlide'].CHandoutMaster = CHandoutMaster;
	window['AscCommonSlide'].createHandoutMaster = createHandoutMaster;
	window['AscCommonSlide'].HandoutSettings = HandoutSettings;
	window['AscCommonSlide'].HANDOUT_ALIGNMENT_HORIZONTAL = HANDOUT_ALIGNMENT_HORIZONTAL;
	window['AscCommonSlide'].HANDOUT_ALIGNMENT_VERTICAL = HANDOUT_ALIGNMENT_VERTICAL;
	window['AscCommonSlide'].CHANDOUTMASTER_HORIZONTAL_FIELD = CHANDOUTMASTER_HORIZONTAL_FIELD;
	window['AscCommonSlide'].CHANDOUTMASTER_VERTICAL_FIELD = CHANDOUTMASTER_VERTICAL_FIELD;
})();
