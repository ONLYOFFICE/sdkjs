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

"use strict";

(function() {
	const History = AscCommon.History;

	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetTheme] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetHF] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterAddToSpTree] = AscDFH.CChangesDrawingsContentPresentation;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterRemoveFromTree] = AscDFH.CChangesDrawingsContentPresentation;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetBg] = AscDFH.CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetName] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_HandoutMasterSetClrMap] = AscDFH.CChangesDrawingsObject;


	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetTheme] = function(oClass, value) {
		oClass.Theme = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetHF] = function(oClass, value) {
		oClass.hf = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetName] = function(oClass, value) {
		oClass.cSld.name = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetClrMap] = function(oClass, value) {
		oClass.clrMap = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_HandoutMasterSetBg] = function(oClass, value, FromLoad) {
		oClass.cSld.Bg = value;
		if (FromLoad) {
			var Fill;
			if (oClass.cSld.Bg && oClass.cSld.Bg.bgPr && oClass.cSld.Bg.bgPr.Fill) {
				Fill = oClass.cSld.Bg.bgPr.Fill;
			}
			if (typeof AscCommon.CollaborativeEditing !== "undefined") {
				if (Fill && Fill.fill && Fill.fill.type === Asc.c_oAscFill.FILL_TYPE_BLIP && typeof Fill.fill.RasterImageId === "string" && Fill.fill.RasterImageId.length > 0) {
					AscCommon.CollaborativeEditing.Add_NewImage(Fill.fill.RasterImageId);
				}
			}
		}
	};

	AscDFH.drawingsConstructorsMap[AscDFH.historyitem_HandoutMasterSetBg] = AscFormat.CBg;

	AscDFH.drawingContentChanges[AscDFH.historyitem_HandoutMasterAddToSpTree] = function(oClass) {
		return oClass.cSld.spTree;
	};
	AscDFH.drawingContentChanges[AscDFH.historyitem_HandoutMasterRemoveFromTree] = function(oClass) {
		return oClass.cSld.spTree;
	};

	function CHandoutMaster() {
		AscCommonSlide.SlideBase.call(this);
		this.clrMap = new AscFormat.ClrMap();
		this.cSld = new AscFormat.CSld(this);
		this.hf = null;

		this.slideCount = 6;

		this.Theme = null;
		this.kind = AscFormat.TYPE_KIND.HANDOUT_MASTER;
		this.m_oContentChanges = new AscCommon.CContentChanges();
		this.setSlideSize(this.presentation.GetNotesWidthMM(), this.presentation.GetNotesHeightMM());
		this.recalcInfo = {
			recalculateBackground: true,
			recalculateSpTree:     true
		};
	}

	AscFormat.InitClass(CHandoutMaster, AscCommonSlide.SlideBase, AscDFH.historyitem_type_HandoutMaster);


	CHandoutMaster.prototype.getObjectType = function() {
		return AscDFH.historyitem_type_HandoutMaster;
	};

	CHandoutMaster.prototype.setTheme = function(pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_HandoutMasterSetTheme, this.Theme, pr));
		this.Theme = pr;
	};
	CHandoutMaster.prototype.setClrMap = function(pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_HandoutMasterSetClrMap, this.clrMap, pr));
		this.clrMap = pr;
	};

	CHandoutMaster.prototype.setHF = function(pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_HandoutMasterSetHF, this.hf, pr));
		this.hf = pr;
	};

	CHandoutMaster.prototype.removeFromSpTreeById = function(id) {
		for (var i = this.cSld.spTree.length - 1; i > -1; --i) {
			if (this.cSld.spTree[i].Get_Id() === id) {
				this.removeFromSpTreeByPos(i);
			}
		}
	};

	CHandoutMaster.prototype.changeBackground = function(bg) {
		History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_HandoutMasterSetBg, this.cSld.Bg, bg));
		this.cSld.Bg = bg;
	};

	CHandoutMaster.prototype.setCSldName = function(pr) {
		History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_HandoutMasterSetName, this.cSld.name, pr));
		this.cSld.name = pr;
	};

	CHandoutMaster.prototype.getAllFonts = function(fonts) {
		var i;
		if (this.Theme) {
			this.Theme.Document_Get_AllFontNames(fonts);
		}

		for (i = 0; i < this.cSld.spTree.length; ++i) {
			if (typeof this.cSld.spTree[i].getAllFonts === "function")
				this.cSld.spTree[i].getAllFonts(fonts);
		}
	};

	CHandoutMaster.prototype.createDuplicate = function(IdMap) {
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
	CHandoutMaster.prototype.getThemeIndex = function() {
		let aMasters = Asc.editor.WordControl.m_oLogicDocument.notesMasters;
		for (let nIdx = 0; nIdx < aMasters.length; ++nIdx) {
			if (aMasters[nIdx] === this) {
				return -nIdx - 1;
			}
		}
		return 0;
	};
	CHandoutMaster.prototype.drawNoPlaceholders = function(graphics, slide) {

	};
	CHandoutMaster.prototype.draw = function(graphics, slide) {
		if (slide) {
			if (AscFormat.isRealNumber(slide.num) && slide.num !== this.lastRecalcSlideIndex) {
				this.lastRecalcSlideIndex = slide.num;
				this.cSld.refreshAllContentsFields(true);
			}
		} else {
			if (-1 !== this.lastRecalcSlideIndex) {
				this.lastRecalcSlideIndex = -1;
				this.cSld.refreshAllContentsFields(true);

			}
		}
		this.recalculate();

		DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
		this.cSld.forEachSp(function(oSp) {
			if (!AscCommon.IsHiddenObj(oSp)) {
				oSp.draw(graphics);
			}
		});
		this.drawPlaceholders(graphics);
	};
	const CHANDOUTMASTER_GAP = 10;
	const CHANDOUTMASTER_FIELDSIZE = 20;
	CHandoutMaster.prototype.getGap = function(paperSize, slideSize, scale, repeatCount) {
		return repeatCount === 1 ? 0 : (paperSize - slideSize * repeatCount * scale) / (repeatCount - 1);
	}
	CHandoutMaster.prototype.getHandouts = function() {
		//todo align
		let align = 1;

		const handouts = [];
		let slidesCount = this.slideCount;
		let isDrawWithSlideTextPlaceholder = false;
		if (slidesCount === 3) {
			isDrawWithSlideTextPlaceholder = true;
			slidesCount = 6;
		}
		const scaledFieldSize = CHANDOUTMASTER_FIELDSIZE;
		let countSlidesOnRow = this.getSlidesCountOnRow(slidesCount);
		let rowsCount = Math.floor(slidesCount / countSlidesOnRow);
		if (this.presentation.GetNotesWidthMM() < this.presentation.GetNotesHeightMM()) {
			const temp = countSlidesOnRow;
			countSlidesOnRow = rowsCount;
			rowsCount = temp;
			align = 0;
		}
		const w_mm = this.presentation.GetWidthMM();
		const h_mm = this.presentation.GetHeightMM();
		const paperWidth = this.presentation.GetNotesWidthMM() - scaledFieldSize * 2;
		const paperHeight = this.presentation.GetNotesHeightMM() - scaledFieldSize * 2;
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

		if (align === 0) {
			for (let i = 0; i < rowsCount; i += 1) {
				const slideY = startY + i * verticalGap + h_mm * i * slideScale;
				for (let j = 0; j < countSlidesOnRow; j += 1) {
					const slideX = startX + j * w_mm * slideScale + j * horizontalGap;
					if (isDrawWithSlideTextPlaceholder && handouts.length % 2 === 1) {
						handouts.push(new TextPlaceholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight));
					} else {
						handouts.push(new SlidePlaceholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight));
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
						handouts.push(new SlidePlaceholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight));
					}
				}
			}
		}
		return handouts;
	};
	CHandoutMaster.prototype.drawPlaceholders = function(g) {
		const handouts = this.getHandouts();
		for (let i = 0; i < handouts.length; i += 1) {
			const handout = handouts[i];
			handout.draw(g);
		}
	};
	CHandoutMaster.prototype.getSlidesCountOnRow = function(slidesCount) {
		if (slidesCount % 3 === 0) {
			return 3;
		} else if (slidesCount % 2) {
			return 2;
		}
		return 1;
	}
	CHandoutMaster.prototype.getDrawingsForController = function() {
		return this.cSld.spTree;
	};
	CHandoutMaster.prototype.setSlideSize = function(nW, nH) {
		this.Width = nW;
		this.Height = nH;
	};
	CHandoutMaster.prototype.getTheme = function() {
		return this.Theme || null;
	};
	CHandoutMaster.prototype.getColorMap = function() {
		return this.Get_ColorMap();
	};
	CHandoutMaster.prototype.recalculateBackground = function() {
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
	CHandoutMaster.prototype.checkSlideColorScheme = function() {
		this.recalcInfo.recalculateSpTree = true;
		this.recalcInfo.recalculateBackground = true;
		for (let i = 0; i < this.cSld.spTree.length; ++i) {
			const shape = this.cSld.spTree[i];
			if (!shape.isPlaceholder()) {
				shape.handleUpdateFill();
				shape.handleUpdateLn();
			}
		}
	};
	CHandoutMaster.prototype.createFontMap = function(oFontsMap, oCheckedMap, isNoPh) {
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
	CHandoutMaster.prototype.recalculate = function() {
		if (this.recalcInfo.recalculateSpTree) {
			this.recalcInfo.recalculateSpTree = false;
			for (let i = 0; i < this.cSld.spTree.length; ++i) {
				this.cSld.spTree[i].recalculate();
			}
		}
		if (this.recalcInfo.recalculateBackground) {
			this.recalcInfo.recalculateBackground = false;
		}
	};
	CHandoutMaster.prototype.isAnimated = function() {
		return false;
	};
	CHandoutMaster.prototype.isLockedObject = function() {
		return false;
	};
	//todo
	CHandoutMaster.prototype.isSlide = function() {
		return false;
	};
	CHandoutMaster.prototype.isLayout = function() {
		return false;
	};
	CHandoutMaster.prototype.isMaster = function() {
		return false;
	};
	//todo
	CHandoutMaster.prototype.getParentObjects = function() {
		return {};
	};
	CHandoutMaster.prototype.getWorksheet = function() {
		return null;
	};
	CHandoutMaster.prototype.Get_ColorMap = function() {
		if (this.clrMap) {
			return this.clrMap;
		}
		return AscFormat.GetDefaultColorMap();
	};
	CHandoutMaster.prototype.IsUseInDocument = function() {
		let oPresentation = Asc.editor.private_GetLogicDocument();
		if (!oPresentation) return false;
		for (let nIdx = 0; nIdx < oPresentation.handoutMasters.length; ++nIdx) {
			if (oPresentation.handoutMasters[nIdx] === this) {
				return true;
			}
		}
		return false;
	};
	CHandoutMaster.prototype.IsUseInSlides = function() {
		return false;
	};
	CHandoutMaster.prototype.isPreserve = function() {
		return false;
	};
	CHandoutMaster.prototype.shapeRemove = function(pos, count) {
		if (pos > -1 && pos < this.cSld.spTree.length) {
			History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_HandoutMasterRemoveFromTree, pos, this.cSld.spTree.slice(pos, pos + count), false));
			this.cSld.spTree.splice(pos, count);
		}
	};
	CHandoutMaster.prototype.removeFromSpTreeByPos = function(pos) {
		this.shapeRemove(pos, 1);
	};
	CHandoutMaster.prototype.shapeAdd = function(pos, item) {
		this.checkDrawingUniNvPr(item);
		var _pos = (AscFormat.isRealNumber(pos) && pos > -1 && pos <= this.cSld.spTree.length) ? pos : this.cSld.spTree.length;
		History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_HandoutMasterAddToSpTree, _pos, [item], true, true));
		this.cSld.spTree.splice(_pos, 0, item);
		item.setParent2(this);
	};
	CHandoutMaster.prototype.addToSpTreeToPos = function(pos, item) {
		this.shapeAdd(pos, item);
	};
	CHandoutMaster.prototype.Refresh_RecalcData = function(data) {
		if (data) {
			switch (data.Type) {
				case AscDFH.historyitem_HandoutMasterSetBg: {
					this.recalcInfo.recalculateBackground = true;
					break;
				}
				// case AscDFH.historyitem_HandoutMasterAddToSpTree:
				//  this.recalcInfo.recalculateSpTree = true;
				// 	break;
			}
			this.addToRecalculate();
		}
	};
	CHandoutMaster.prototype.setSlideCount = function(count) {
		this.slideCount = count;
	}

	function PlaceholderBase(x, y, width, height) {
		this.x = x;
		this.y = y;
		this.width = width;
		this.height = height;
	}

	PlaceholderBase.prototype.draw = function(g) {

	};

	function SlidePlaceholder(x, y, width, height) {
		PlaceholderBase.call(this, x, y, width, height);
	};
	AscFormat.InitClassWithoutType(SlidePlaceholder, PlaceholderBase);
	SlidePlaceholder.prototype.draw = function(graphics) {
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
		graphics.AddSmartRect(0, 0, this.width, this.height, 0);
		graphics.p_dash(null);
		graphics.RestoreGrState();
	};

	const TEXTPLACEHOLDER_STEP_COUNT = 7;

	function TextPlaceholder(x, y, width, height) {
		PlaceholderBase.call(this, x, y, width, height);
	}

	AscFormat.InitClassWithoutType(TextPlaceholder, PlaceholderBase);
	TextPlaceholder.prototype.draw = function(graphics) {
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
})();
