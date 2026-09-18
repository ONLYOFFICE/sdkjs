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
	AscDFH.changesFactory[AscDFH.historyitem_NotesSetClrMap] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_NotesSetShowMasterPhAnim] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_NotesSetShowMasterSp] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_NotesSetSlide] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_NotesSetNotesMaster] = AscDFH.CChangesDrawingsObject;


	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetClrMap] = function (oClass, value) {oClass.clrMap = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetShowMasterPhAnim] = function (oClass, value) {oClass.showMasterPhAnim = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetShowMasterSp] = function (oClass, value) {oClass.showMasterSp = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetSlide] = function (oClass, value) {oClass.slide = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetNotesMaster] = function (oClass, value) {oClass.Master = value;};

	//Temporary function
	function GetNotesWidth() {
		return editor.WordControl.m_oDrawingDocument.Notes_GetWidth();
	}

	function CNotes() {
		AscCommonSlide.SlideBase.call(this);
		this.showMasterPhAnim = null;
		this.showMasterSp = true;
		this.slide = null;
		this.clrMap = null;

		this.Master = null;
		this.recalcInfo =
			{
				recalculateBackground: true,
				recalculateSpTree    : true,
				recalculateSlide     : true
			};

		this.kind = AscFormat.TYPE_KIND.NOTES;

		this.setSlideSize(this.presentation.GetNotesWidthMM(), this.presentation.GetNotesHeightMM());

		this.Lock = new AscCommon.CLock();
		this.graphicObjects = new AscFormat.DrawingObjectsController(this);
	}

	AscFormat.InitClass(CNotes, AscCommonSlide.SlideBase, AscDFH.historyitem_type_Notes);
	CNotes.prototype.setClMapOverride = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesSetClrMap, this.clrMap, pr));
		this.clrMap = pr;
	};

	CNotes.prototype.setShowMasterPhAnim = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_NotesSetShowMasterPhAnim, this.showMasterPhAnim, pr));
		this.showMasterPhAnim = pr;
	};

	CNotes.prototype.setShowMasterSp = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_NotesSetShowMasterSp, this.showMasterSp, pr));
		this.showMasterSp = pr;
	};

	CNotes.prototype.removeFromSpTreeById = function (id) {
		for (var i = this.cSld.spTree.length - 1; i > -1; --i) {
			if (this.cSld.spTree[i].Get_Id() === id) {
				this.removeFromSpTreeByPos(i);
			}
		}
	};


	CNotes.prototype.setCSldName = function (pr) {
		this.cSld.setName(pr);
	};
	CNotes.prototype.Write_ToBinary2 = AscCommonSlide.SlideBase.prototype.writeSlideBinary;
	CNotes.prototype.Read_FromBinary2 = AscCommonSlide.SlideBase.prototype.readSlideBinary;
	CNotes.prototype.setSlide = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesSetSlide, this.slide, pr));
		this.slide = pr;
	};

	CNotes.prototype.setNotesMaster = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesSetNotesMaster, this.Master, pr));
		this.Master = pr;
	};

	CNotes.prototype.getWidth = function () {
		return GetNotesWidth();
	};

	CNotes.prototype.createBodyShape = function () {
		const oSp = new AscFormat.CShape();
		oSp.setBDeleted(false);

		const oNvSpPr = new AscFormat.UniNvPr();

		const oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(3);
		oCNvPr.setName('Notes Placeholder 2');

		const oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_body);
		oPh.setIdx(1 + "");

		oNvSpPr.nvPr.setPh(oPh);
		oSp.setNvSpPr(oNvSpPr);
		oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
		oSp.setSpPr(new AscFormat.CSpPr());
		oSp.spPr.setParent(oSp);
		oSp.createTextBody();

		const oBodyPr = new AscFormat.CBodyPr();
		oSp.txBody.setBodyPr(oBodyPr);

		const oTxLstStyle = new AscFormat.TextListStyle();
		oSp.txBody.setLstStyle(oTxLstStyle);

		oSp.setParent(this);
		this.addToSpTreeToPos(1, oSp);

		return oSp;
	};

	CNotes.prototype.recalcBackground = function () {
		this.recalcInfo.recalculateBackground = true;
	};
	CNotes.prototype.handleCSldRecalc = function (data) {
		if (data) {
			switch (data.Type) {
				case AscDFH.historyitem_CSld_SetBg:
					this.recalcBackground();
					break;
			}
			this.addToRecalculate();
		}
	};

	CNotes.prototype.Refresh_RecalcData = function (data) {
		if (data) {
			this.addToRecalculate();
		}
	};

	CNotes.prototype.Refresh_RecalcData2 = function () {

	};

	CNotes.prototype.createDuplicate = function (IdMap) {

		var oIdMap = IdMap || {};
		var oPr = new AscFormat.CCopyObjectProperties();
		oPr.idMap = oIdMap;
		var copy = new CNotes();
		if (this.clrMap) {
			copy.setClMapOverride(this.clrMap.createDuplicate());
		}

		this.cSld.fillObject(copy.cSld, oPr);
		if (AscFormat.isRealBool(this.showMasterPhAnim)) {
			copy.setShowMasterPhAnim(this.showMasterPhAnim);
		}
		if (AscFormat.isRealBool(this.showMasterSp)) {
			copy.setShowMasterSp(this.showMasterSp);
		}
		copy.setNotesMaster(this.Master);

		return copy;
	};

	//todo think about it
	CNotes.prototype.showDrawingObjects = function () {
		var oPresentation = editor.WordControl.m_oLogicDocument;
		if (oPresentation.IsNotesPageMode()) {
			AscCommonSlide.SlideBase.prototype.showDrawingObjects.call(this);
		} else if (this.slide) {
			if (oPresentation.CurPage === this.slide.num) {
				editor.WordControl.m_oDrawingDocument.Notes_OnRecalculate(this.slide.num, this.slide.NotesWidth, this.slide.getNotesHeight());
			}
		}
	};
	CNotes.prototype.isViewerMode = function () {
		return editor.WordControl.m_oLogicDocument.IsViewMode();
	};
	CNotes.prototype.convertPixToMM = function (pix) {
		return editor.WordControl.m_oDrawingDocument.GetMMPerDot(pix);
	};

	CNotes.prototype.drawViewPrMarks = function (oGraphics) {
		if (oGraphics.isSupportTextDraw && !oGraphics.isSupportTextDraw()) return;
		return AscCommonSlide.Slide.prototype.drawViewPrMarks.call(this, oGraphics);
	};
	CNotes.prototype.getDrawingsForController = function () {
		return this.cSld.spTree;
	};
	CNotes.prototype.getTheme = function () {
		return this.Master && this.Master.getTheme() || null;
	};
	CNotes.prototype.getColorMap = function () {
		if (this.Master) {
			if (this.Master.clrMap) {
				return this.Master.clrMap;
			}
		}
		return AscFormat.GetDefaultColorMap();
	};
	CNotes.prototype.recalculateBackground = function () {
		let RGBA = {R: 0, G: 0, B: 0, A: 255};

		const _master = this.Master;
		const _theme = _master.Theme;
		let ret;
		if (this && this.cSld.Bg) {
			ret = this.cSld.recalculateBackground(_theme, this, null, _master, RGBA);
		} else if (_master && _master.cSld.Bg) {
			ret = _master.cSld.recalculateBackground(_theme, this, null, _master, RGBA);
		} else {
			ret = this.getDefaultBackFill();
		}
		RGBA = ret.RGBA || RGBA;
		const backFill = ret.backFill;
		if (backFill != null)
			backFill.calculate(_theme, this, this, _master, RGBA);

		this.backgroundFill = backFill;
	};
	CNotes.prototype.getAllFonts = function (fonts) {
		var i;
		for (i = 0; i < this.cSld.spTree.length; ++i) {
			if (typeof this.cSld.spTree[i].getAllFonts === "function")
				this.cSld.spTree[i].getAllFonts(fonts);
		}
	};
	CNotes.prototype.needMasterSpDraw = function () {
		return this.showMasterSp;
	};
	CNotes.prototype.createFontMap = function (oFontsMap, oCheckedMap, isNoPh) {
		if (oCheckedMap[this.Get_Id()]) {
			return;
		}
		var aSpTree = this.cSld.spTree;
		var nSp, nSpCount = aSpTree.length;
		for (nSp = 0; nSp < nSpCount; ++nSp) {
			aSpTree[nSp].createFontMap(oFontsMap);
		}
		if (this.needMasterSpDraw()) {
			this.Master.createFontMap(oFontsMap, oCheckedMap, true);
		}
		oCheckedMap[this.Get_Id()] = this;
	};
	CNotes.prototype.recalculate = function (props) {
		if (!this.Master) {
			return;
		}
		props = props || {};
		if (this.recalcInfo.recalculateSlide) {
			this.recalcInfo.recalculateSlide = false;
			this.recalculateSlide();
		}
		if (this.slide) {
			this.slide.isSkipNotesRecalculate = true;
			this.slide.recalculate();
			this.slide.isSkipNotesRecalculate = false;
		}
		this.Master.recalculate();
		if (this.recalcInfo.recalculateBackground) {
			this.recalculateBackground();
			this.recalcInfo.recalculateBackground = false;
		}
		if (this.recalcInfo.recalculateSpTree) {
			const isSkipNoteShape = !!props.isSkipNoteShape;
			if (isSkipNoteShape) {
				for (let i = 0; i < this.cSld.spTree.length; ++i) {
					const shape = this.cSld.spTree[i];
					if (shape.getPlaceholderType() !== AscFormat.phType_body) {
						shape.recalculate();
					}
				}
			} else {
				for (let i = 0; i < this.cSld.spTree.length; ++i)
					this.cSld.spTree[i].recalculate();
			}

			this.recalcInfo.recalculateSpTree = false;
		}
		//todo cachedImage
		this.cachedImage = null;
	};
	CNotes.prototype.recalculateSlide = function () {
		const presentation = this.presentation;
		if (presentation) {
			for (let i = 0; i < presentation.Slides.length; i += 1) {
				const slide = presentation.Slides[i];
				if (slide.notes === this) {
					this.slide = slide;
					return;
				}
			}
		}
	};
	CNotes.prototype.checkSlideColorScheme = function () {
		this.recalcInfo.recalculateSpTree = true;
		this.recalcBackground();
		this.cSld.forEachSp(function (oSp) {
			oSp.handleUpdateFill();
			oSp.handleUpdateLn();
		});
	};

	CNotes.prototype.getParentObjects = function () {
		return {
			presentation: this.presentation,
			master      : this.Master,
			layout      : null,
			slide       : this
		};
	};
	CNotes.prototype.Get_ColorMap = function () {
		if (this.clrMap) {
			return this.clrMap;
		} else if (this.Master && this.Master.clrMap) {
			return this.Master.clrMap;
		}
		return AscFormat.GetDefaultColorMap();
	};
	CNotes.prototype.IsUseInDocument = function () {
		if (this.slide) {
			return this.slide.IsUseInDocument();
		}
		return false;
	};
	CNotes.prototype.draw = function (graphics, slide, replacedNoteShape) {
		let i;
		this.drawBgMaster(graphics);
		for (i = 0; i < this.cSld.spTree.length; ++i) {
			let oSp = this.cSld.spTree[i];
			if (AscCommon.IsHiddenObj(oSp)) {
				continue;
			}
			if (replacedNoteShape && oSp.getPlaceholderType() === AscFormat.phType_body) {
				replacedNoteShape.draw(graphics);
			} else {
				oSp.draw(graphics);
			}
		}

		if (slide) {
			this.drawViewPrMarks(graphics);
		}
	};
	CNotes.prototype.drawForNextPrintPage = function (graphics) {
		this.drawBgMaster(graphics);
		this.cSld.forEachSp(function (shape) {
			const placeholderType = shape.getPlaceholderType();
			switch (placeholderType) {
				case AscFormat.phType_dt:
				case AscFormat.phType_sldNum:
				case AscFormat.phType_ftr:
				case AscFormat.phType_hdr: {
					shape.draw(graphics);
				}
			}
		});
	};
	CNotes.prototype.drawBgMaster = function (graphics) {
		DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
		if (this.needMasterSpDraw()) {
			this.Master.drawNoPlaceholdersShapesOnly(graphics, this);
		}
	};
	CNotes.prototype.shapeRemove = function (pos, count) {
		if (pos > -1 && pos < this.cSld.spTree.length) {
			this.cSld.removeFromSpTree(pos, 1);
		}
	};
	CNotes.prototype.shapeAdd = function (pos, obj) {
		var _pos = Math.max(0, Math.min(pos, this.cSld.spTree.length));
		this.cSld.checkAndAssignName(obj);
		this.cSld.addToSpTree(_pos, obj);
		obj.setParent2(this);
	};
	CNotes.prototype.setSlideSize = function (w, h) {
		this.Width = w;
		this.Height = h;
	};
	CNotes.prototype.isNote = function () {
		return true;
	};
	CNotes.prototype.collectRedrawSlides = function (redrawSlides, force) {
		redrawSlides.addRedrawSlide(this);
	};
	CNotes.prototype.recalcBody = function () {
		const body = this.getBodyShape();
		if (body) {
			body.setRecalculateInfo();
		}
	};
	CNotes.prototype.checkOnDeletePlaceholderInSpTree = function (placeholdersMap) {
		this.cSld.forEachSp(function (shape) {
			if (shape.isPlaceholder() && shape.getPlaceholderType() === AscFormat.phType_body) {
				shape.recalculateTransform();
			}
			shape.checkOnDeletePlaceholder(placeholdersMap);
		});
	};

	function copyMasterPlaceholderToNotes(oN, oMasterSp, bIsSpecialPh) {
		const sp = oMasterSp.copy(undefined);
		sp.setParent(oN);
		if (sp.txBody) {
			sp.txBody.setBodyPr(new AscFormat.CBodyPr());
			sp.txBody.setLstStyle(new AscFormat.TextListStyle());
		}
		if (sp.spPr) {
			sp.spPr.setXfrm(null);
		}
		if (!bIsSpecialPh) {
			sp.clearContent();
		}
		oN.addToSpTreeToPos(oN.cSld.spTree.length, sp);
		return sp;
	}

	function CreateNotes(oNotesMaster) {
		const oNotes = new CNotes();
		const oMasterSldImg = oNotesMaster && oNotesMaster.getMatchingShape(AscFormat.phType_sldImg, null, false, {});
		if (oMasterSldImg) {
			copyMasterPlaceholderToNotes(oNotes, oMasterSldImg, false);
		} else {
			const oAddedSldImg = AscCommonSlide.addSlideImage(oNotes);
			oAddedSldImg.clearContent();
		}

		const oMasterBody = oNotesMaster && oNotesMaster.getMatchingShape(AscFormat.phType_body, null, true);
		if (oMasterBody) {
			copyMasterPlaceholderToNotes(oNotes, oMasterBody, false);
		} else {
			const oAddedBody = AscCommonSlide.addBodyShape(oNotes);
			oAddedBody.clearContent();
		}

		const oMasterSldNum = oNotesMaster && oNotesMaster.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
		if (oMasterSldNum) {
			copyMasterPlaceholderToNotes(oNotes, oMasterSldNum, true);
		} else {
			AscCommonSlide.addNumberShape(oNotes);
		}
		return oNotes;
	}

	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonSlide'].CNotes = CNotes;
	window['AscCommonSlide'].GetNotesWidth = GetNotesWidth;
	window['AscCommonSlide'].CreateNotes = CreateNotes;
})();
