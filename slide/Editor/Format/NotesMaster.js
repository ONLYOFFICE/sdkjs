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

	AscDFH.changesFactory[AscDFH.historyitem_NotesMasterSetNotesTheme] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_NotesMasterSetHF] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_NotesMasterSetNotesStyle] = AscDFH.CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_NotesMasterSetClrMap] = AscDFH.CChangesDrawingsObject;


	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesMasterSetNotesTheme] = function (oClass, value) {
		oClass.Theme = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesMasterSetHF] = function (oClass, value) {
		oClass.hf = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesMasterSetNotesStyle] = function (oClass, value) {
		oClass.txStyles = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesMasterSetClrMap] = function (oClass, value) {
		oClass.clrMap = value;
	};
	AscDFH.drawingsConstructorsMap[AscDFH.historyitem_NotesMasterSetNotesStyle] = AscFormat.TextListStyle;

	function CNotesMaster() {
		AscCommonSlide.SlideBase.call(this);
		this.clrMap = new AscFormat.ClrMap();
		this.hf = null;
		this.txStyles = null;

		this.Theme = null;
		this.kind = AscFormat.TYPE_KIND.NOTES_MASTER;
		this.setSlideSize(this.presentation.GetNotesWidthMM(), this.presentation.GetNotesHeightMM());
	}

	AscFormat.InitClass(CNotesMaster, AscCommonSlide.SlideBase, AscDFH.historyitem_type_NotesMaster);

	CNotesMaster.prototype.setTheme = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesMasterSetNotesTheme, this.Theme, pr));
		this.Theme = pr;
	};
	CNotesMaster.prototype.setClrMap = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesMasterSetClrMap, this.clrMap, pr));
		this.clrMap = pr;
	};

	CNotesMaster.prototype.setHF = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesMasterSetHF, this.hf, pr));
		this.hf = pr;
	};

	CNotesMaster.prototype.setNotesStyle = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_NotesMasterSetNotesStyle, this.txStyles, pr));
		this.txStyles = pr;
	};

	CNotesMaster.prototype.createDuplicate = function (IdMap) {
		var oIdMap = IdMap || {};
		var oPr = new AscFormat.CCopyObjectProperties();
		oPr.idMap = oIdMap;
		var i;
		var copy = new CNotesMaster();
		if (this.clrMap) {
			copy.setClrMap(this.clrMap.createDuplicate());
		}
		this.cSld.fillObject(copy.cSld, oPr);
		if (this.hf) {
			copy.setHF(this.hf.createDuplicate());
		}
		if (this.txStyles) {
			copy.setNotesStyle(this.txStyles.createDuplicate());
		}
		return copy;
	};


	CNotesMaster.prototype.Clear_ContentChanges = function () {
		this.m_oContentChanges.Clear();
	};

	CNotesMaster.prototype.Add_ContentChanges = function (Changes) {
		this.m_oContentChanges.Add(Changes);
	};

	CNotesMaster.prototype.Refresh_ContentChanges = function () {
		this.m_oContentChanges.Refresh();
	};

	CNotesMaster.prototype.handleCSldRecalc = function (data) {
		if (data) {
			switch (data.Type) {
				case AscDFH.historyitem_CSld_SetBg:
					this.recalcInfo.recalculateBackground = true;
					break;
				case AscDFH.historyitem_CSld_AddToSpTree:
				case AscDFH.historyitem_CSld_RemoveFromSpTree:
					this.recalcInfo.recalculateSpTree = true;
					break;
			}
			this.addToRecalculate();
		}
	};
	CNotesMaster.prototype.Refresh_RecalcData = function (data) {
		if (data) {
			switch (data.Type) {
				case AscDFH.historyitem_NotesMasterSetNotesTheme: {
					this.presentation.bNeedUpdateThemes = true;
					this.checkSlideTheme();
					for (let i = 0; i < this.presentation.notes.length; i += 1) {
						if (this.presentation.notes[i].Master === this) {
							this.presentation.notes[i].checkSlideTheme();
							this.presentation.notes[i].addToRecalculate();
						}
					}
					break;
				}
			}
			this.addToRecalculate();
		}
	};
	CNotesMaster.prototype.getThemeIndex = function () {
		let aMasters = Asc.editor.WordControl.m_oLogicDocument.notesMasters;
		for (let nIdx = 0; nIdx < aMasters.length; ++nIdx) {
			if (aMasters[nIdx] === this) {
				return -nIdx - 1;
			}
		}
		return 0;
	};
	CNotesMaster.prototype.getDrawingsForController = function () {
		return this.cSld.spTree;
	};
	CNotesMaster.prototype.getTheme = function () {
		return this.Theme || null;
	};
	CNotesMaster.prototype.getColorMap = function () {
		return AscFormat.GetDefaultColorMap();
	};
	CNotesMaster.prototype.recalculateBackground = function () {
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
	CNotesMaster.prototype.checkSlideColorScheme = function () {
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
	CNotesMaster.prototype.getAllFonts = function (fonts) {
		var i;
		if (this.Theme) {
			this.Theme.Document_Get_AllFontNames(fonts);
		}

		if (this.txStyles) {
			this.txStyles.Document_Get_AllFontNames(fonts);
		}

		for (i = 0; i < this.cSld.spTree.length; ++i) {
			if (typeof this.cSld.spTree[i].getAllFonts === "function")
				this.cSld.spTree[i].getAllFonts(fonts);
		}
	};
	CNotesMaster.prototype.createFontMap = function (oFontsMap, oCheckedMap, isNoPh) {
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
	CNotesMaster.prototype.recalculate = function () {
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
	//todo
	CNotesMaster.prototype.getParentObjects = function () {
		return {
			presentation: Asc.editor.private_GetLogicDocument(),
			master      : this,
			layout      : null,
			slide       : null
		};
	};
	CNotesMaster.prototype.Get_ColorMap = function () {
		return this.getColorMap();
	};
	CNotesMaster.prototype.IsUseInDocument = function () {
		const presentation = this.presentation;
		for (let i = 0; i < presentation.notesMasters.length; i += 1) {
			if (presentation.notesMasters[i] === this) {
				return true;
			}
		}
		return false;
	};
	CNotesMaster.prototype.draw = function (graphics, slide) {
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
		this.cSld.forEachSp(function (oSp) {
			if (!AscCommon.IsHiddenObj(oSp)) {
				oSp.draw(graphics);
			}
		});
	};
	CNotesMaster.prototype.drawNoPlaceholdersShapesOnly = function (graphics, slide) {
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
		this.cSld.forEachSp(function (oSp) {
			if (!AscCommon.IsHiddenObj(oSp) && !oSp.isPlaceholder()) {
				oSp.draw(graphics);
			}
		});
	};
	CNotesMaster.prototype.shapeRemove = function (pos, count) {
		if (pos > -1 && pos < this.cSld.spTree.length) {
			this.cSld.removeFromSpTree(pos, 1);
		}
	};
	CNotesMaster.prototype.shapeAdd = function (pos, obj) {
		var _pos = Math.max(0, Math.min(pos, this.cSld.spTree.length));
		this.cSld.checkAndAssignName(obj);
		this.cSld.addToSpTree(_pos, obj);
		obj.setParent2(this);
	};
	CNotesMaster.prototype.setSlideSize = function (w, h) {
		this.Width = w;
		this.Height = h;
	};
	CNotesMaster.prototype.isNoteMaster = function () {
		return true;
	};
	CNotesMaster.prototype.collectRedrawSlides = function (redrawSlides, force) {
		redrawSlides.addRedrawSlide(this);
		const presentation = redrawSlides.presentation;
		for (let i = 0; i < presentation.notes.length; i += 1) {
			const note = presentation.notes[i];
			if (note.Master === this) {
				redrawSlides.addRedrawSlide(note);
			}
		}
	};
	CNotesMaster.prototype.Write_ToBinary2 = AscCommonSlide.SlideBase.prototype.writeMasterBinary;
	CNotesMaster.prototype.Read_FromBinary2 = AscCommonSlide.SlideBase.prototype.readMasterBinary;
	CNotesMaster.prototype.isMasterLikeObject = function () {
		return true;
	};

	function CreateNotesMaster() {
		var oNM = new CNotesMaster();
		var oBG = new AscFormat.CBg();
		var oBgRef = new AscFormat.StyleRef();
		oBgRef.idx = 1001;
		var oUniColor = new AscFormat.CUniColor();
		oUniColor.color = new AscFormat.CSchemeColor();
		oUniColor.color.id = 6;//bg1
		oBgRef.Color = oUniColor;
		oBG.bgRef = oBgRef;
		oNM.changeBackground(oBG);
		AscCommonSlide.addHeaderShape(oNM);
		AscCommonSlide.addDateShape(oNM);
		AscCommonSlide.addBodyShape(oNM);
		AscCommonSlide.addFooterShape(oNM);
		AscCommonSlide.addSlideImage(oNM);
		AscCommonSlide.addNumberShape(oNM);

		//clrMap
		oNM.clrMap.setClr(0, 0);
		oNM.clrMap.setClr(1, 1);
		oNM.clrMap.setClr(2, 2);
		oNM.clrMap.setClr(3, 3);
		oNM.clrMap.setClr(4, 4);
		oNM.clrMap.setClr(5, 5);
		oNM.clrMap.setClr(10, 10);
		oNM.clrMap.setClr(11, 11);
		oNM.clrMap.setClr(6, 12);
		oNM.clrMap.setClr(7, 13);
		oNM.clrMap.setClr(15, 8);
		oNM.clrMap.setClr(16, 9);

		const oTxLstStyle = new AscFormat.TextListStyle();

		oTxLstStyle.levels[0] = new CParaPr();
		oTxLstStyle.levels[0].Ind.Left = 0;
		oTxLstStyle.levels[0].Jc = AscCommon.align_Left;
		oTxLstStyle.levels[0].DefaultTab = 914400 / 36000;

		oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
		oTxLstStyle.levels[0].DefaultRunPr.FontSize = 12;
		oTxLstStyle.levels[0].DefaultRunPr.Unifill = AscFormat.CreateUniFillSchemeColorWidthTint(15, 0);

		oTxLstStyle.levels[0].DefaultRunPr.RFonts.SetFontStyle(AscFormat.fntStyleInd_minor);

		oTxLstStyle.levels[1] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[1].Ind.Left = 457200 / 36000;

		oTxLstStyle.levels[2] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[2].Ind.Left = 914400 / 36000;

		oTxLstStyle.levels[3] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[3].Ind.Left = 1371600 / 36000;

		oTxLstStyle.levels[4] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[4].Ind.Left = 1828800 / 36000;

		oTxLstStyle.levels[5] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[5].Ind.Left = 2286000 / 36000;

		oTxLstStyle.levels[6] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[6].Ind.Left = 2743200 / 36000;

		oTxLstStyle.levels[7] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[7].Ind.Left = 3200400 / 36000;

		oTxLstStyle.levels[8] = oTxLstStyle.levels[0].Copy();
		oTxLstStyle.levels[8].Ind.Left = 3657600 / 36000;
		oNM.setNotesStyle(oTxLstStyle);
		return oNM;
	}

	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonSlide'].CNotesMaster = CNotesMaster;
	window['AscCommonSlide'].CreateNotesMaster = CreateNotesMaster;
})();
