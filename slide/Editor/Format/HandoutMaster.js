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

		this.slideCounts = 6;

		this.Theme = null;
		this.kind = AscFormat.TYPE_KIND.HANDOUT_MASTER;
		this.m_oContentChanges = new AscCommon.CContentChanges();
		this.setSlideSize(this.presentation.GetNotesWidthMM(), this.presentation.GetNotesHeightMM());
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

	CHandoutMaster.prototype.addToSpTreeToPos = function(pos, obj) {
		var _pos = Math.max(0, Math.min(pos, this.cSld.spTree.length));
		History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_HandoutMasterAddToSpTree, _pos, [obj], true));
		this.cSld.spTree.splice(_pos, 0, obj);
		obj.setParent2(this);
	};

	CHandoutMaster.prototype.removeFromSpTreeByPos = function(pos) {
		if (pos > -1 && pos < this.cSld.spTree.length) {
			History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_HandoutMasterRemoveFromTree, pos, this.cSld.spTree.splice(pos, 1), false));
		}
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
	CHandoutMaster.prototype.getTheme = function() {
		return this.Theme || null;
	};
	//todo
	CHandoutMaster.prototype.getColorMap = function() {
		return AscFormat.GetDefaultColorMap();
	};
	CHandoutMaster.prototype.drawNoPlaceholders = function(graphics, slide) {

	};
	CHandoutMaster.prototype.recalculate = function() {
		for (let i = 0; i < this.cSld.spTree.length; ++i) {
			this.cSld.spTree[i].recalculate();
		}
	}
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
	const gap = 10;
	const maxGap = 30;
	const fieldSize = 20;
	CHandoutMaster.prototype.getGap = function(paperSize, slideSize, scale, repeatCount) {
		return repeatCount === 1 ? 0: (paperSize - slideSize * repeatCount * scale) / (repeatCount - 1);
	}
	CHandoutMaster.prototype.getHandouts = function () {
		const handouts = [];
		const slidesCount = this.slideCounts;
		//todo align
		const align = 1;

		const scaledFieldSize = fieldSize;
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
		const paperHeight = this.presentation.GetNotesHeightMM() - scaledFieldSize * 2;
		const slidesWidth = w_mm * countSlidesOnRow;
		const slidesHeight = h_mm * rowsCount;
		const resultWidthWithMaxGap = slidesWidth + (countSlidesOnRow - 1) * maxGap;
		const resultHeightWithMaxGap = slidesHeight + (rowsCount - 1) * maxGap;
		let slideScale = Math.min(paperWidth / resultWidthWithMaxGap, paperHeight / resultHeightWithMaxGap);
		let horizontalGap = this.getGap(paperWidth, w_mm, slideScale, countSlidesOnRow);
		let verticalGap = this.getGap(paperHeight, h_mm, slideScale, rowsCount);
		if (horizontalGap < gap || verticalGap < gap) {
			const paperWidthWithoutMinGap = paperWidth - gap * (countSlidesOnRow - 1);
			const paperHeightWithoutMinGap = paperHeight - gap * (rowsCount - 1);
			slideScale = Math.min(paperWidthWithoutMinGap / slidesWidth, paperHeightWithoutMinGap / slidesHeight);
			horizontalGap = this.getGap(paperWidth, w_mm, slideScale, countSlidesOnRow);
			verticalGap = this.getGap(paperHeight, h_mm, slideScale, rowsCount);
		}

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
					handouts.push(new Placeholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight));
				}
			}
		} else {
			for (let i = 0; i < countSlidesOnRow; i += 1) {
				const slideX = startX + i * horizontalGap + w_mm * i * slideScale;
				for (let j = 0; j < rowsCount; j += 1) {
					const slideY = startY + j * h_mm * slideScale + j * verticalGap;
					handouts.push(new Placeholder(slideX, slideY, scaledPresentationWidth, scaledPresentationHeight));
				}
			}
		}
		return handouts;
	};
	CHandoutMaster.prototype.drawPlaceholders = function (g) {
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
	CHandoutMaster.prototype.getDrawingsForController = function () {
		return this.cSld.spTree;
	};
	CHandoutMaster.prototype.setSlideSize = function (nW, nH) {
		this.Width = nW;
		this.Height = nH;
	};
	CHandoutMaster.prototype.getParentObjects = function() {
		return {}
	};

	function PlaceholderBase(x, y, width, height) {
		this.x = x;
		this.y = y;
		this.width = width;
		this.height = height;
	}
	PlaceholderBase.prototype.draw = function (g) {

	};
	function Placeholder(x, y, width, height) {
		PlaceholderBase.call(this, x, y, width, height);
	};
	AscFormat.InitClassWithoutType(Placeholder, PlaceholderBase);
	Placeholder.prototype.draw = function (graphics) {
		graphics.SaveGrState();
		const transform = new AscCommon.CMatrix();
		transform.tx = this.x;
		transform.ty = this.y;
		graphics.transform3(transform);
		const color = parseInt(AscCommon.GlobalSkin.PageOutline.slice(1), 16);
		graphics.p_color((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF, 0xFF);
		graphics.p_width(0);
		graphics._s();
		graphics._m(0, 0);
		graphics._l(this.width, 0);
		graphics._l(this.width, this.height);
		graphics._l(0, this.height);
		graphics._z();
		graphics.ds();
		graphics.RestoreGrState();
	};

	function CreateHandoutMaster() {
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

		var oSp = new AscFormat.CShape();
		oSp.setBDeleted(false);
		var oNvSpPr = new AscFormat.UniNvPr();
		var oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(2);
		oCNvPr.setName("Header Placeholder 1");
		var oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_hdr);
		oPh.setSz(2);
		oNvSpPr.nvPr.setPh(oPh);
		oSp.setNvSpPr(oNvSpPr);
		oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
		oSp.setSpPr(new AscFormat.CSpPr());
		oSp.spPr.setParent(oSp);
		oSp.spPr.setXfrm(new AscFormat.CXfrm());
		oSp.spPr.xfrm.setParent(oSp.spPr);
		oSp.spPr.xfrm.setOffX(0);
		oSp.spPr.xfrm.setOffY(0);
		oSp.spPr.xfrm.setExtX(2971800 / 36000);
		oSp.spPr.xfrm.setExtY(458788 / 36000);
		oSp.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
		oSp.spPr.geometry.setParent(oSp.spPr);
		oSp.createTextBody();
		var oBodyPr = oSp.txBody.bodyPr.createDuplicate();
		oBodyPr.vert = AscFormat.nVertTThorz;
		oBodyPr.lIns = 91440 / 36000;
		oBodyPr.tIns = 45720 / 36000;
		oBodyPr.rIns = 91440 / 36000;
		oBodyPr.bIns = 45720 / 36000;
		oBodyPr.rtlCol = false;
		oBodyPr.anchor = 1;
		oSp.txBody.setBodyPr(oBodyPr);
		var oTxLstStyle = new AscFormat.TextListStyle();
		oTxLstStyle.levels[0] = new CParaPr();
		oTxLstStyle.levels[0].Jc = AscCommon.align_Left;
		oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
		oTxLstStyle.levels[0].DefaultRunPr.FontSize = 12;
		oSp.txBody.setLstStyle(oTxLstStyle);
		oSp.setParent(oHandoutMaster);
		oHandoutMaster.addToSpTreeToPos(0, oSp);

		oSp = new AscFormat.CShape();
		oSp.setBDeleted(false);
		oNvSpPr = new AscFormat.UniNvPr();
		oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(3);
		oCNvPr.setName("Date Placeholder 2");
		oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_dt);
		oPh.setIdx(3 + "");
		oNvSpPr.nvPr.setPh(oPh);
		oSp.setNvSpPr(oNvSpPr);
		oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
		oSp.setSpPr(new AscFormat.CSpPr());
		oSp.spPr.setParent(oSp);
		oSp.spPr.setXfrm(new AscFormat.CXfrm());
		oSp.spPr.xfrm.setParent(oSp.spPr);
		oSp.spPr.xfrm.setOffX(3884613 / 36000);
		oSp.spPr.xfrm.setOffY(0);
		oSp.spPr.xfrm.setExtX(2971800 / 36000);
		oSp.spPr.xfrm.setExtY(458788 / 36000);
		oSp.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
		oSp.spPr.geometry.setParent(oSp.spPr);
		oSp.createTextBody();
		oBodyPr = oSp.txBody.bodyPr.createDuplicate();
		oBodyPr.vert = AscFormat.nVertTThorz;
		oBodyPr.lIns = 91440 / 36000;
		oBodyPr.tIns = 45720 / 36000;
		oBodyPr.rIns = 91440 / 36000;
		oBodyPr.bIns = 45720 / 36000;
		oSp.txBody.setBodyPr(oBodyPr);
		oTxLstStyle = new AscFormat.TextListStyle();
		oTxLstStyle.levels[0] = new CParaPr();
		oTxLstStyle.levels[0].Jc = AscCommon.align_Right;
		oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
		oTxLstStyle.levels[0].DefaultRunPr.FontSize = 12;
		//endParaPr
		oSp.txBody.setLstStyle(oTxLstStyle);
		oSp.setParent(oHandoutMaster);
		oHandoutMaster.addToSpTreeToPos(1, oSp);

		oSp = new AscFormat.CShape();
		oSp.setBDeleted(false);
		oNvSpPr = new AscFormat.UniNvPr();
		oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(4);
		oCNvPr.setName("Footer Placeholder 3");
		oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_ftr);
		oPh.setIdx(2 + "");
		oPh.setSz(2);
		oNvSpPr.nvPr.setPh(oPh);
		oSp.setNvSpPr(oNvSpPr);
		oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
		oSp.setSpPr(new AscFormat.CSpPr());
		oSp.spPr.setParent(oSp);
		oSp.spPr.setXfrm(new AscFormat.CXfrm());
		oSp.spPr.xfrm.setParent(oSp.spPr);
		oSp.spPr.xfrm.setOffX(0);
		oSp.spPr.xfrm.setOffY(8685213 / 36000);
		oSp.spPr.xfrm.setExtX(2971800 / 36000);
		oSp.spPr.xfrm.setExtY(458787 / 36000);
		oSp.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
		oSp.spPr.geometry.setParent(oSp.spPr);
		oSp.createTextBody();
		oBodyPr = oSp.txBody.bodyPr.createDuplicate();
		oBodyPr.vert = AscFormat.nVertTThorz;
		oBodyPr.lIns = 91440 / 36000;
		oBodyPr.tIns = 45720 / 36000;
		oBodyPr.rIns = 91440 / 36000;
		oBodyPr.bIns = 45720 / 36000;
		oBodyPr.rtlCol = false;
		oBodyPr.anchor = 0;
		oSp.txBody.setBodyPr(oBodyPr);
		oTxLstStyle = new AscFormat.TextListStyle();
		oTxLstStyle.levels[0] = new CParaPr();
		oTxLstStyle.levels[0].Jc = AscCommon.align_Left;
		oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
		oTxLstStyle.levels[0].DefaultRunPr.FontSize = 12;
		//endParaPr
		oSp.txBody.setLstStyle(oTxLstStyle);
		oSp.setParent(oHandoutMaster);
		oHandoutMaster.addToSpTreeToPos(2, oSp);

		oSp = new AscFormat.CShape();
		oSp.setBDeleted(false);
		oNvSpPr = new AscFormat.UniNvPr();
		oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(5);
		oCNvPr.setName("Slide Number Placeholder 4");
		oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_sldNum);
		oPh.setIdx(3 + "");
		oPh.setSz(2);
		oNvSpPr.nvPr.setPh(oPh);
		oSp.setNvSpPr(oNvSpPr);
		oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
		oSp.setSpPr(new AscFormat.CSpPr());
		oSp.spPr.setParent(oSp);
		oSp.spPr.setXfrm(new AscFormat.CXfrm());
		oSp.spPr.xfrm.setParent(oSp.spPr);
		oSp.spPr.xfrm.setOffX(3884613 / 36000);
		oSp.spPr.xfrm.setOffY(8685213 / 36000);
		oSp.spPr.xfrm.setExtX(2971800 / 36000);
		oSp.spPr.xfrm.setExtY(458787 / 36000);
		oSp.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
		oSp.spPr.geometry.setParent(oSp.spPr);
		oSp.createTextBody();
		oBodyPr = oSp.txBody.bodyPr.createDuplicate();
		oBodyPr.vert = AscFormat.nVertTThorz;
		oBodyPr.lIns = 91440 / 36000;
		oBodyPr.tIns = 45720 / 36000;
		oBodyPr.rIns = 91440 / 36000;
		oBodyPr.bIns = 45720 / 36000;
		oBodyPr.rtlCol = false;
		oBodyPr.anchor = 0;
		oSp.txBody.setBodyPr(oBodyPr);
		oTxLstStyle = new AscFormat.TextListStyle();
		oTxLstStyle.levels[0] = new CParaPr();
		oTxLstStyle.levels[0].Jc = AscCommon.align_Right;
		oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
		oTxLstStyle.levels[0].DefaultRunPr.FontSize = 12;
		//endParaPr
		oSp.txBody.setLstStyle(oTxLstStyle);
		oSp.setParent(oHandoutMaster);
		oHandoutMaster.addToSpTreeToPos(3, oSp);

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
	window['AscCommonSlide'].CreateHandoutMaster = CreateHandoutMaster;
})();
