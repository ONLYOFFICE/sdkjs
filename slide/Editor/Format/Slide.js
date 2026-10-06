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
	// Import
	var g_oTableId = AscCommon.g_oTableId;
	var History = AscCommon.History;


	function CChangesDrawingsContentComments(Class, Type, Pos, Items, isAdd) {
		AscDFH.CChangesDrawingsContent.call(this, Class, Type, Pos, Items, isAdd);
	}

	CChangesDrawingsContentComments.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
	CChangesDrawingsContentComments.prototype.constructor = CChangesDrawingsContentComments;
	CChangesDrawingsContentComments.prototype.addToInterface = function () {
		for (var i = 0; i < this.Items.length; ++i) {
			var oComment = this.Items[i];
			oComment.Data.bDocument = !(oComment.Parent && (oComment.Parent.slide instanceof Slide));
			editor.sync_AddComment(oComment.Get_Id(), oComment.Data);
		}
	};
	CChangesDrawingsContentComments.prototype.removeFromInterface = function () {
		for (var i = 0; i < this.Items.length; ++i) {
			editor.sync_RemoveComment(this.Items[i].Get_Id());
		}
	};
	CChangesDrawingsContentComments.prototype.Undo = function () {
		AscDFH.CChangesDrawingsContent.prototype.Undo.call(this);
		if (this.IsAdd()) {
			this.removeFromInterface();
		} else {
			this.addToInterface();
		}
	};
	CChangesDrawingsContentComments.prototype.Redo = function () {
		AscDFH.CChangesDrawingsContent.prototype.Redo.call(this);
		if (this.IsAdd()) {
			this.addToInterface();
		} else {
			this.removeFromInterface();
		}
	};

	CChangesDrawingsContentComments.prototype.Load = function () {
		AscDFH.CChangesDrawingsContent.prototype.Load.call(this);
		if (this.IsAdd()) {
			this.addToInterface();
		} else {
			this.removeFromInterface();
		}
	};

	AscDFH.CChangesDrawingsContentComments = CChangesDrawingsContentComments;


	AscDFH.changesFactory[AscDFH.historyitem_SlideSetLocks] = AscDFH.CChangesDrawingSlideLocks;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetComments] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetShow] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetShowPhAnim] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetShowMasterSp] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetLayout] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetNum] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetTransition] = AscDFH.CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetSize] = AscDFH.CChangesDrawingsObjectNoId;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetClrMapOverride] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_SlideCommentsAddComment] = AscDFH.CChangesDrawingsContentComments;
	AscDFH.changesFactory[AscDFH.historyitem_SlideCommentsRemoveComment] = AscDFH.CChangesDrawingsContentComments;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetNotes] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetTiming] = AscDFH.CChangesDrawingsObject;


	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetComments] = function (oClass, value) {oClass.slideComments = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetShow] = function (oClass, value) {oClass.show = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetShowPhAnim] = function (oClass, value) {oClass.showMasterPhAnim = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetShowMasterSp] = function (oClass, value) {oClass.showMasterSp = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetLayout] = function (oClass, value) {oClass.Layout = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetTiming] = function (oClass, value) {oClass.timing = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetNum] = function (oClass, value) {oClass.num = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetTransition] = function (oClass, value) {oClass.transition = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetSize] = function (oClass, value) {
		oClass.Width = value.a;
		oClass.Height = value.b;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetClrMapOverride] = function (oClass, value) {oClass.clrMap = value;};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetNotes] = function (oClass, value) {oClass.notes = value;};


	AscDFH.drawingContentChanges[AscDFH.historyitem_SlideCommentsAddComment] =
		AscDFH.drawingContentChanges[AscDFH.historyitem_SlideCommentsRemoveComment] = function (oClass) {return oClass.comments;};

	AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideSetSize] = AscFormat.CDrawingBaseCoordsWritable;
	AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideSetBg] = AscFormat.CBg;
	AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideSetTransition] = Asc.CAscSlideTransition;

	function Slide(presentation, slideLayout, slideNum) {
		AscCommonSlide.SlideBase.call(this);
		this.kind = AscFormat.TYPE_KIND.SLIDE;

		this.graphicObjects = new AscFormat.DrawingObjectsController(this);
		this.cSld.initCollaborativeMarks();
		this.clrMap = null; // override ClrMap

		this.show = true;
		this.showMasterPhAnim = false;
		this.showMasterSp = true;

		this.backgroundFill = null;

		this.notes = null;

		this.transition = new Asc.CAscSlideTransition();
		this.transition.setDefaultParams();

		this.timing = null;

		this.recalcInfo =
			{
				recalculateBackground: true,
				recalculateSpTree    : true
			};

		this.searchingArray = [];  // array of objects for selection
		this.selectionArray = [];  // array of objects for search


		this.writecomments = [];

		this.commentX = 0;
		this.commentY = 0;

		this.Lock = new AscCommon.CLock();


		this.notesShape = null;

		this.lastLayoutType = null;
		this.lastLayoutMatchingName = null;
		this.lastLayoutName = null;

		this.NotesWidth = -10.0;

		this.animationPlayer = null;

		if (presentation) {
			this.setSlideSize(presentation.GetWidthMM(), presentation.GetHeightMM());
			this.setSlideComments(new SlideComments(this));
		}

		if (slideLayout) {
			this.setLayout(slideLayout);
		}
		if (typeof slideNum === "number") {
			this.setSlideNum(slideNum);
		}

		this.isSkipNotesRecalculate = false;
	}

	AscFormat.InitClass(Slide, AscCommonSlide.SlideBase, AscDFH.historyitem_type_Slide);
	Slide.prototype.Clear_CollaborativeMarks = function () {
		this.cSld.Clear_CollaborativeMarks();
	};
	Slide.prototype.createDuplicate = function (IdMap, bCacheImage) {
		var oIdMap = IdMap || {};
		var oPr = new AscFormat.CCopyObjectProperties();
		oPr.idMap = oIdMap;
		oPr.cacheImage = bCacheImage !== false;
		var copy = new Slide(this.presentation, this.Layout, 0), i;
		this.cSld.fillObject(copy.cSld, oPr);

		if (this.clrMap) {
			copy.setClMapOverride(this.clrMap.createDuplicate());
		}
		if (AscFormat.isRealBool(this.show)) {
			copy.setShow(this.show);
		}
		if (AscFormat.isRealBool(this.showMasterPhAnim)) {
			copy.setShowPhAnim(this.showMasterPhAnim);
		}
		if (AscFormat.isRealBool(this.showMasterSp)) {
			copy.setShowMasterSp(this.showMasterSp);
		}

		copy.applyTransition(this.transition.createDuplicate());
		copy.setSlideSize(this.Width, this.Height);

		if (this.notes) {
			copy.setNotes(this.notes.createDuplicate(undefined, oPr.cacheImage));
		}
		if (this.slideComments) {
			if (!copy.slideComments) {
				copy.setSlideComments(new SlideComments(copy));
			}
			var aComments = this.slideComments.comments;
			for (i = 0; i < aComments.length; ++i) {
				copy.slideComments.addComment(aComments[i].createDuplicate(copy.slideComments, true));
			}
		}

		if (!this.recalcInfo.recalculateBackground && !this.recalcInfo.recalculateSpTree) {
			if (!oPr || false !== oPr.cacheImage) {
				copy.cachedImage = this.getBase64Img();
			}
		}
		if (this.timing) {
			copy.setTiming(this.timing.createDuplicate(oIdMap));
		}
		return copy;
	};

	Slide.prototype.refreshAllContentsFields = function () {
		this.cSld.refreshAllContentsFields();
		if (this.notesShape) {
			this.notesShape.handleAllContents(AscFormat.RefreshContentAllFields);
		}
	};
	Slide.prototype.changeNum = function (num) {
		this.num = num;
	};
	Slide.prototype.addComment = function (comment) {
		if (AscCommon.isRealObject(this.slideComments)) {
			this.slideComments.addComment(comment);
		}
	};
	Slide.prototype.changeComment = function (id, commentData) {
		if (AscCommon.isRealObject(this.slideComments)) {
			this.slideComments.changeComment(id, commentData);
		}
	};
	Slide.prototype.removeMyComments = function () {
		if (AscCommon.isRealObject(this.slideComments)) {
			this.slideComments.removeMyComments();
		}
	};
	Slide.prototype.removeAllComments = function () {
		if (AscCommon.isRealObject(this.slideComments)) {
			this.slideComments.removeAllComments();
		}
	};
	Slide.prototype.removeComment = function (id, bForce) {
		if (AscCommon.isRealObject(this.slideComments)) {
			this.slideComments.removeComment(id, bForce);
		}
	};
	Slide.prototype.getAllMyComments = function (aAllComments) {
		if (this.slideComments) {
			this.slideComments.getAllMyComments(aAllComments, this);
		}
	};
	Slide.prototype.getAllComments = function (aAllComments, isMine, isCurrent, aIds) {
		if (this.slideComments) {
			this.slideComments.getAllComments(aAllComments, isMine, isCurrent, aIds);
		}
	};
	Slide.prototype.Refresh_RecalcData = function (data) {
		if (data) {
			switch (data.Type) {
				case AscDFH.historyitem_SlideSetLayout: {
					this.checkSlideTheme();
					if (this.Layout) {
						this.lastLayoutType = this.Layout.type;
						this.lastLayoutMatchingName = this.Layout.matchingName;
						this.lastLayoutName = this.Layout.cSld.name;
					}
					break;
				}
				case AscDFH.historyitem_SlideSetComments: {
					this.recalcComments();
					break;
				}
			}
			this.addToRecalculate();
		}
	};
	Slide.prototype.handleCSldRecalc = function (data) {
		if (data && data.Type === AscDFH.historyitem_CSld_SetBg) {
			this.recalcInfo.recalculateBackground = true;
		}
		this.addToRecalculate();
	};
	Slide.prototype.setNotes = function (pr) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetNotes, this.notes, pr));
		this.notes = pr;
	};
	Slide.prototype.setSlideComments = function (comments) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetComments, this.slideComments, comments));
		this.slideComments = comments;
	};
	Slide.prototype.setShow = function (bShow) {
		History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_SlideSetShow, this.show, bShow));
		this.show = bShow;
	};
	Slide.prototype.setShowPhAnim = function (bShow) {
		History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_SlideSetShowPhAnim, this.showMasterPhAnim, bShow));
		this.showMasterPhAnim = bShow;
	};
	Slide.prototype.setShowMasterSp = function (bShow) {
		History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_SlideSetShowMasterSp, this.showMasterSp, bShow));
		this.showMasterSp = bShow;
	};
	Slide.prototype.setLayout = function (layout) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetLayout, this.Layout, layout));
		this.Layout = layout;
		if (layout) {
			this.lastLayoutType = layout.type;
			this.lastLayoutMatchingName = layout.matchingName;
			this.lastLayoutName = layout.cSld.name;
		}
	};

	Slide.prototype.changeLayout = function (layout) {
		this.setLayout(layout);
		this.bChangeLayout = true;
		for (var j = this.cSld.spTree.length - 1; j > -1; --j) {
			var shape = this.cSld.spTree[j];
			if (shape.isEmptyPlaceholder()) {
				this.removeFromSpTreeById(shape.Get_Id());
			} else {
				var oInfo = {};
				var hierarchy = shape.getHierarchy(undefined, oInfo);
				var bNoPlaceholder = true;
				var bNeedResetTransform = false;
				var oNotNullPH = null;
				for (var t = 0; t < hierarchy.length; ++t) {
					if (hierarchy[t]) {
						if (hierarchy[t].parent && (hierarchy[t].parent instanceof AscCommonSlide.SlideLayout)) {
							bNoPlaceholder = false;
						}
						if (hierarchy[t].spPr && hierarchy[t].spPr.xfrm && hierarchy[t].spPr.xfrm.isNotNull()) {
							bNeedResetTransform = true;
							oNotNullPH = hierarchy[t];
						}
					}
				}
				if (bNoPlaceholder) {
					if (this.cSld.spTree[j].isEmptyPlaceholder()) {
						this.removeFromSpTreeById(this.cSld.spTree[j].Get_Id());
					} else {
						var hierarchy2 = shape.getHierarchy(undefined, undefined);
						for (var t = 0; t < hierarchy2.length; ++t) {
							if (hierarchy2[t]) {
								if (hierarchy2[t].spPr && hierarchy2[t].spPr.xfrm && hierarchy2[t].spPr.xfrm.isNotNull()) {
									break;
								}
							}
						}
						if (t === hierarchy2.length) {
							AscFormat.CheckSpPrXfrm(shape);
						}
					}
				} else {
					if (bNeedResetTransform) {
						if (shape.spPr && shape.spPr.xfrm && shape.spPr.xfrm.isNotNull()) {
							if (shape.getObjectType() !== AscDFH.historyitem_type_GraphicFrame
								&& shape.getObjectType() !== AscDFH.historyitem_type_SmartArt
								&& shape.getObjectType() !== AscDFH.historyitem_type_ChartSpace) {
								shape.spPr.setXfrm(null);
							} else {
								if (oNotNullPH && oNotNullPH.spPr && oNotNullPH.spPr.xfrm) {
									if (!shape.spPr) {
										shape.setSpPr(oNotNullPH.spPr.createDuplicate());
										shape.spPr.setParent(shape);
									} else {
										shape.spPr.setXfrm(oNotNullPH.spPr.xfrm.createDuplicate());
									}
								}
							}
						}
					}
				}
			}
		}
		for (let i = 0; i < layout.cSld.spTree.length; ++i) {
			if (layout.cSld.spTree[i].isPlaceholder()) {
				const phType = layout.cSld.spTree[i].getPlaceholderType();
				const hf = layout.Master.hf;
				const isSpecialPh = phType === AscFormat.phType_dt || phType === AscFormat.phType_ftr || phType === AscFormat.phType_hdr || phType === AscFormat.phType_sldNum;
				if (!isSpecialPh || hf && ((phType === AscFormat.phType_dt && (hf.dt !== false)) ||
					(phType === AscFormat.phType_ftr && (hf.ftr !== false)) ||
					(phType === AscFormat.phType_hdr && (hf.hdr !== false)) ||
					(phType === AscFormat.phType_sldNum && (hf.sldNum !== false)))) {
					const matchingShape = this.getMatchingShape(layout.cSld.spTree[i].getPlaceholderType(), layout.cSld.spTree[i].getPlaceholderIndex(), layout.cSld.spTree[i].getIsSingleBody ? layout.cSld.spTree[i].getIsSingleBody() : false);
					if (matchingShape == null && layout.cSld.spTree[i]) {
						const sp = layout.cSld.spTree[i].copy(undefined);
						sp.setParent(this);
						if (sp.txBody) {
							sp.txBody.setBodyPr(new AscFormat.CBodyPr());
							if (sp.txBody.lstStyle) {
								sp.txBody.setLstStyle(null);
							}
						}
						if (sp.spPr) {
							if (sp.spPr.xfrm) {
								sp.spPr.setXfrm(null);
							}
							if (sp.spPr.Fill) {
								sp.spPr.setFill(null);
							}
						}
						!isSpecialPh && sp.clearContent && sp.clearContent();
						this.addToSpTreeToPos(this.cSld.spTree.length, sp)
					}
				}
			}
		}
	};
	Slide.prototype.setSlideNum = function (num) {
		History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_SlideSetNum, this.num, num));
		this.num = num;
	};
	Slide.prototype.applyTransition = function (transition) {
		var oldTransition = this.transition.createDuplicate();
		this.transition.applyProps(transition);
		History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideSetTransition, oldTransition, this.transition.createDuplicate()));
	};
	Slide.prototype.setTiming = function (oTiming) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetTiming, this.timing, oTiming));
		this.timing = oTiming;
		if (this.timing) {
			this.timing.setParent(this);
		}
	};
	Slide.prototype.setSlideSize = function (w, h) {
		History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideSetSize, new AscFormat.CDrawingBaseCoordsWritable(this.Width, this.Height), new AscFormat.CDrawingBaseCoordsWritable(w, h)));
		this.Width = w;
		this.Height = h;
	};

	Slide.prototype.shapeAdd = function (pos, item) {
		this.checkDrawingUniNvPr(item);
		var _pos = (AscFormat.isRealNumber(pos) && pos > -1 && pos <= this.cSld.spTree.length) ? pos : this.cSld.spTree.length;
		this.cSld.checkAndAssignName(item);
		this.cSld.addToSpTree(_pos, item);
		item.setParent2(this);
		const presentation = Asc.editor.WordControl.m_oLogicDocument;
		if (presentation) {
			presentation.checkShapeOutlineView(item);
		}
	};

	Slide.prototype.shapeRemove = function (pos, count) {
		if (pos > -1 && pos < this.cSld.spTree.length) {
			this.cSld.removeFromSpTree(pos, count);
		}
	};

	Slide.prototype.isEditingInFastMultipleUsers = function () {
		var oPresentation = editor.WordControl.m_oLogicDocument;
		if (oPresentation && oPresentation.IsEditingInFastMultipleUsers()) {
			return true;
		}
		return false;
	};
	Slide.prototype.checkNeedCopyTimingBeforeEdit = function () {
		// if(this.isEditingInFastMultipleUsers()) {
		//     if(this.timing) {
		//         this.setTiming(this.timing.createDuplicate());
		//     }
		// }
	};
	Slide.prototype.addAnimation = function (nPresetClass, nPresetId, nPresetSubtype, oColor, bReplace) {
		this.checkNeedCopyTimingBeforeEdit();
		if (!this.timing) {
			this.setTiming(new AscFormat.CTiming());
		}
		return this.timing.addAnimation(nPresetClass, nPresetId, nPresetSubtype, oColor, bReplace);
	};
	Slide.prototype.setAnimationProperties = function (oPr) {
		if (!this.timing) {
			return;
		}
		this.checkNeedCopyTimingBeforeEdit();
		this.timing.setAnimationProperties(oPr);
		this.showDrawingObjects();
	};

	Slide.prototype.isVisible = function () {
		return this.show !== false;
	};


	Slide.prototype.CheckLayout = function () {
		var bRet = true;
		if (!this.Layout || !this.Layout.CheckCorrect()) {
			var oMaster = this.presentation.slideMasters[0];
			if (!oMaster) {
				bRet = false;
			} else {
				var oLayout = oMaster.getMatchingLayout(this.lastLayoutType, this.lastLayoutMatchingName, this.lastLayoutName, undefined);
				if (oLayout) {
					this.setLayout(oLayout);
				} else {
					bRet = false;
				}
			}
		}
		return bRet;
	};

	Slide.prototype.correctContent = function () {


		for (var i = this.cSld.spTree.length - 1; i > -1; --i) {
			if (this.cSld.spTree[i].CheckCorrect && !this.cSld.spTree[i].CheckCorrect() || this.cSld.spTree[i].bDeleted) {
				if (this.cSld.spTree[i].setBDeleted) {
					this.cSld.spTree[i].setBDeleted(true);
				}
				this.removeFromSpTreeById(this.cSld.spTree[i].Get_Id());
			}
		}
		for (var i = this.cSld.spTree.length - 1; i > -1; --i) {
			for (var j = i - 1; j > -1; --j) {
				if (this.cSld.spTree[i] === this.cSld.spTree[j]) {
					this.removeFromSpTreeByPos(i);
					break;
				}
			}
		}
	};

	Slide.prototype.replaceSp = function (oPh, oObject) {
		var aSpTree = this.cSld.spTree;
		for (var i = 0; i < aSpTree.length; ++i) {
			if (aSpTree[i] === oPh) {
				break;
			}
		}
		this.removeFromSpTreeByPos(i);
		this.addToSpTreeToPos(i, oObject);
		var oNvProps = oObject.getNvProps && oObject.getNvProps();
		if (oNvProps) {
			if (oPh) {
				var oNvPropsPh = oPh.getNvProps && oPh.getNvProps();
				var oPhPr = oNvPropsPh.ph;
				if (oPhPr) {
					oNvProps.setPh(oPhPr.createDuplicate());
				}
			}
		}
	};

	Slide.prototype.setCSldName = function (name) {
		this.cSld.setName(name);
	};

	Slide.prototype.setClMapOverride = function (clrMap) {
		History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetClrMapOverride, this.clrMap, clrMap));
		this.clrMap = clrMap;
	};
	Slide.prototype.Write_ToBinary2 = AscCommonSlide.SlideBase.prototype.writeSlideBinary;
	Slide.prototype.Read_FromBinary2 = AscCommonSlide.SlideBase.prototype.readSlideBinary;
	Slide.prototype.getAllFonts = function (fonts) {
		this.cSld.forEachSp(function (oSp) {
			oSp.getAllFonts(fonts);
		});
	};

	Slide.prototype.getParentObjects = function () {
		var oRet = {master: null, layout: null, slide: this};
		if (this.Layout) {
			oRet.layout = this.Layout;
			if (this.Layout.Master) {
				oRet.master = this.Layout.Master;
			}
		}
		return oRet;
	};

	Slide.prototype.getAllRasterImagesOnSlide = function (aImages) {
		aImages = aImages || [];
		this.getAllRasterImages(aImages);
		const oLayout = this.Layout;
		if (oLayout) {
			oLayout.getAllRasterImages(aImages);
			const oMaster = oLayout.Master;
			if (oMaster) {
				oMaster.getAllRasterImages(aImages);
			}
		}
		return aImages;
	};
	Slide.prototype.checkPlaceholders = function (aPlaceholders) {
		let bResult = false;
		this.cSld.forEachSp(function (oSp) {
			let bSpResult = oSp.checkPlaceholders(aPlaceholders);
			if (bSpResult) {
				bResult = true;
			}
		});
		return bResult;
	};

	Slide.prototype.checkSlideColorScheme = function () {
		this.recalcInfo.recalculateSpTree = true;
		this.recalcInfo.recalculateBackground = true;
		this.cSld.forEachSp(function (oSp) {
			oSp.handleUpdateFill();
			oSp.handleUpdateLn();
		});
	};
	Slide.prototype.Get_ColorMap = function () {
		if (this.clrMap) {
			return this.clrMap;
		} else if (this.Layout && this.Layout.clrMap) {
			return this.Layout.clrMap;
		} else if (this.Layout.Master && this.Layout.Master.clrMap) {
			return this.Layout.Master.clrMap;
		}
		return AscFormat.GetDefaultColorMap();
	};

	Slide.prototype.recalculate = function () {
		if (!this.Layout || !AscFormat.isRealNumber(this.num)) {
			return;
		}
		this.Layout.recalculate();
		if (this.recalcInfo.recalculateBackground) {
			this.recalculateBackground();
			this.recalcInfo.recalculateBackground = false;
		}
		if (this.recalcInfo.recalculateSpTree) {
			this.recalculateSpTree();
			this.recalcInfo.recalculateSpTree = false;
		}
		if (!this.isSkipNotesRecalculate) {
			const viewManager = this.presentation && this.presentation.getViewManager();
			if (viewManager) {
				viewManager.recalculateNotesShapeOnSlide(this);
			}
		}
		this.cachedImage = null;
	};

	Slide.prototype.recalculateBackground = function () {
		let RGBA = {R: 0, G: 0, B: 0, A: 255};

		const _layout = this.Layout;
		const _master = _layout.Master;
		const _theme = _master.Theme;
		let ret;
		if (this.cSld.Bg) {
			ret = this.cSld.recalculateBackground(_theme, this, _layout, _master, RGBA);
		} else if (_layout && _layout.cSld.Bg) {
			ret = _layout.cSld.recalculateBackground(_theme, this, _layout, _master, RGBA);
		} else if (_master && _master.cSld.Bg) {
			ret = _master.cSld.recalculateBackground(_theme, this, _layout, _master, RGBA);
		} else {
			ret = this.getDefaultBackFill();
		}
		RGBA = ret.RGBA || RGBA;
		const backFill = ret.backFill;
		if (backFill != null)
			backFill.calculate(_theme, this, _layout, _master, RGBA);

		this.backgroundFill = backFill;
	};

	Slide.prototype.recalculateSpTree = function () {
		for (var i = 0; i < this.cSld.spTree.length; ++i)
			this.cSld.spTree[i].recalculate();
	};

	Slide.prototype.getNotesHeight = function () {
		if (!this.notesShape) {
			return 0;
		}

		var oDocContent = this.notesShape.getDocContent();
		if (oDocContent) {
			return oDocContent.GetSummaryHeight();
		}
		return 0;
	};

	Slide.prototype.recalculateNotesShape = function () {
		AscFormat.ExecuteNoHistory(function () {
			if (!this.notes) {
				this.notesShape = null;//ToDo: Create notes body shape
			} else {
				this.notesShape = this.notes.getBodyShape();
				if (this.notesShape && this.notesShape.getObjectType() !== AscDFH.historyitem_type_Shape) {
					this.notesShape = null;
				}
			}

			if (this.notesShape) {
				this.notes.slide = this;
				this.notes.graphicObjects.selectObject(this.notesShape, 0);
				this.notes.graphicObjects.selection.textSelection = this.notesShape;
				this.notesShape.recalculateNotesShape();
			}
		}, this, []);
	};

	Slide.prototype.needMasterSpDraw = function () {
		return this.showMasterSp ? this.Layout.showMasterSp : false;
	};

	Slide.prototype.needLayoutSpDraw = function () {
		return this.showMasterSp;
	};

	Slide.prototype.isEqualBgMasterAndLayout = function (oSlide) {
		if (!(this.backgroundFill === oSlide.backgroundFill ||
			this.backgroundFill && this.backgroundFill.isEqual(oSlide.backgroundFill))) {
			return false;
		}
		if (this.needMasterSpDraw() && !oSlide.needMasterSpDraw() || oSlide.needMasterSpDraw() && !this.needMasterSpDraw()) {
			return false;
		}
		if (this.needMasterSpDraw()) {
			if (this.Layout.Master !== oSlide.Layout.Master) {
				return false;
			}
		}
		if (this.needLayoutSpDraw()) {
			if (this.Layout !== oSlide.Layout) {
				return false;
			}
		}
		return true;
	};

	Slide.prototype.drawBgMasterAndLayout = function (graphics, bClipBySlide, bCheckBounds) {
		let _bounds;
		if (!graphics.m_bNoPrintSlideBackground) {
			DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
		}
		if (bClipBySlide) {
			graphics.SaveGrState();
			graphics.AddClipRect(0, 0, this.Width, this.Height);
		}
		if (this.needMasterSpDraw()) {
			if (bCheckBounds) {
				_bounds = this.Layout.Master.bounds;
				graphics.rect(_bounds.l, _bounds.t, _bounds.w, _bounds.h);
			} else {
				this.Layout.Master.drawNoPlaceholdersShapesOnly(graphics, this);
			}
		}
		if (this.needLayoutSpDraw()) {
			if (bCheckBounds) {
				_bounds = this.Layout.bounds;
				graphics.rect(_bounds.l, _bounds.t, _bounds.w, _bounds.h);
			} else {
				this.Layout.drawNoPlaceholdersShapesOnly(graphics, this);
			}
		}
	};
	Slide.prototype.draw = function (graphics) {
		if (this.presentation) {
			this.presentation.checkSlideRecalculateOnOpen(this);
		}
		let bCheckBounds = graphics.isBoundsChecker();
		let bSlideShow = this.graphicObjects.isSlideShow();
		let bClipBySlide = !this.graphicObjects.canEdit();
		if (bCheckBounds && (bSlideShow || bClipBySlide)) {
			graphics.rect(0, 0, this.Width, this.Height);
			return;
		}
		let _bounds, i;
		let backgroundLayer = graphics.transitionBackgroundLayer;
		if (backgroundLayer === false) {
			if (bClipBySlide) {
				graphics.SaveGrState();
				graphics.AddClipRect(0, 0, this.Width, this.Height);
			}
		} else {
			this.drawBgMasterAndLayout(graphics, bClipBySlide, bCheckBounds);
			if (backgroundLayer === true) {
				if (bClipBySlide) {
					graphics.RestoreGrState();
				}
				return;
			}
		}
		if (this.cSld.collaborativeMarks) {
			this.cSld.collaborativeMarks.Init_Drawing();
		}
		let oCollColor;
		let fDist = 3;
		let oIdentityMtx = new AscCommon.CMatrix();
		for (i = 0; i < this.cSld.spTree.length; ++i) {
			let oSp = this.cSld.spTree[i];
			if (AscCommon.IsHiddenObj(oSp)) {
				continue;
			}
			if (this.cSld.collaborativeMarks) {
				oCollColor = this.cSld.collaborativeMarks.Check(i);
				if (oCollColor) {
					var oBounds = oSp.bounds;
					graphics.transform3(oIdentityMtx);
					if (graphics.put_GlobalAlpha) {
						graphics.put_GlobalAlpha(true, 0.5);
					}
					let dX = oBounds.l - fDist;
					let dY = oBounds.t - fDist;
					let dW = oBounds.r - oBounds.l + 2 * fDist;
					let dH = oBounds.b - oBounds.t + 2 * fDist;
					graphics.drawCollaborativeChanges(dX, dY, dW, dH, oCollColor);
					if (graphics.put_GlobalAlpha) {
						graphics.put_GlobalAlpha(false, 1);
					}
				}
			}
			oSp.draw(graphics);
		}

		if (this.timing) {
			let aShapes = this.timing.getMoveEffectsShapes();
			if (aShapes) {
				for (i = 0; i < aShapes.length; ++i) {
					aShapes[i].draw(graphics);
				}
			}
			this.timing.drawEffectsLabels(graphics);
		}
		if (this.slideComments && (!graphics.PrintComments || graphics.PrintComments.isPrintComments)) {
			let aComments = this.slideComments.comments;
			for (i = 0; i < aComments.length; ++i) {
				var oComment = aComments[i];
				if (AscCommon.UserInfoParser.canViewComment(oComment.GetUserName()) !== false) {
					oComment.draw(graphics);
				}
			}
		}
		if (!bCheckBounds && !bSlideShow) {
			this.drawViewPrMarks(graphics);
		}
		if (bClipBySlide) {
			graphics.RestoreGrState();
		}
	};

	function CStrideData(oPresentation) {
		this.presentation = oPresentation;
		this.slideWidth = null;
		this.slideHeight = null;
		this.stride = null;
		this.originX = null;
		this.originY = null;

	}

	CStrideData.prototype.checkUpdate = function () {
		if (this.slideWidth !== this.presentation.GetWidthEMU() ||
			this.slideHeight !== this.presentation.GetHeightEMU() ||
			this.stride !== this.presentation.getViewPropertiesStride()) {
			this.update();
			return true;
		}
		return false;
	};
	CStrideData.prototype.update = function () {
		this.stride = this.presentation.getViewPropertiesStride();
		this.slideWidth = this.presentation.GetWidthEMU();
		this.slideHeight = this.presentation.GetHeightEMU();

		this.originX = this.getStartStridePos(this.stride, this.slideWidth);
		this.originY = this.getStartStridePos(this.stride, this.slideHeight);
	};
	CStrideData.prototype.getStartStridePos = function (stride, len) {
		let nCenterPos = len / 2 + 0.5 >> 0;
		let nPos = nCenterPos - ((nCenterPos / stride >> 0) + 1) * stride;
		return nPos;
	};
	CStrideData.prototype.getNearestLinearPoint = function (nDX, nOrigin) {
		let nDX_ = Math.abs(nDX);
		let nCX = nDX_ / this.stride >> 0;
		let dRX = nDX_ - nCX * this.stride;
		let nX;
		if ((dRX / this.stride) < 0.5) {
			nX = nCX * this.stride;
		} else {
			nX = (nCX + 1) * this.stride;
		}
		if (nDX < 0) {
			nX = -nX;
		}
		return nOrigin + nX;
	};
	CStrideData.prototype.getNearestPoint = function (x, y) {
		this.checkUpdate();
		let em = function (dValue) {return dValue / 36000;};
		let me = function (dValue) { return dValue * 36000 + 0.5 >> 0;};
		let nDX = me(x) - this.originX;
		let nDY = me(y) - this.originY;
		let nX = this.getNearestLinearPoint(nDX, this.originX);
		let nY = this.getNearestLinearPoint(nDY, this.originY);
		return {x: em(nX), y: em(nY)};
	};
	CStrideData.prototype.drawGrid = function (oGraphics) {
		let oContext = oGraphics.m_oContext;
		if (!oContext) {
			return;
		}
		this.checkUpdate();
		let dPixScale = oGraphics.m_oCoordTransform.sx;
		let mmp = function (dMM) {
			return (dMM * dPixScale + 0.5) >> 0;
		};
		let em = function (val) {return val / 36000;};
		let ep = function (nEmu) {
			return mmp(em(nEmu));
		};

		let nGridSpacingPix = ep(this.stride);

		let nMinLineStridePix = ep(720000);
		let nMinInsideLineStridePix = ep(100000);

		let nGridSpacing = this.stride;
		let bPixel = true;
		let nStrideInsideLine = nGridSpacing;
		let nStrideLine;

		let nStrideInsideLinePix = nGridSpacingPix;
		let nStrideLinePix;

		let nHorStart;
		let nVertStart;
		let nSlideWidth = this.slideWidth;
		let nSlideHeight = this.slideHeight;

		while (nStrideInsideLinePix < nMinInsideLineStridePix) {
			nStrideInsideLine *= 2;
			nStrideInsideLinePix = ep(nStrideInsideLine);
		}
		nStrideLine = nStrideInsideLine;
		nStrideLinePix = nStrideInsideLinePix;
		while (nStrideLinePix < nMinLineStridePix) {
			nStrideLine += nStrideInsideLine;
			nStrideLinePix = ep(nStrideLine);
		}
		bPixel = nStrideInsideLinePix < AscCommon.AscBrowser.convertToRetinaValue(17, true);

		oGraphics.SaveGrState();
		oGraphics.transform3(new AscCommon.CMatrix());
		oGraphics.SetIntegerGrid(true);
		oGraphics.b_color1(0, 0, 0, 255);
		oGraphics._s();

		let nX, nY;
		let oT = oGraphics.m_oFullTransform;
		let oImageCanvas = document.createElement('canvas');
		let oImageContext;
		if (bPixel) {
			oImageCanvas.width = 1;
			oImageCanvas.height = 1;
			oImageContext = oImageCanvas.getContext("2d");
			oImageContext.fillStyle = 'white';
			oImageContext.fillRect(0, 0, oImageCanvas.width, oImageCanvas.height);
			oImageContext.fill();
		} else {
			oImageCanvas.width = 3;
			oImageCanvas.height = 3;
			oImageContext = oImageCanvas.getContext("2d");
			oImageContext.fillStyle = 'white';
			oImageContext.fillRect(1, 0, 1, 1);
			oImageContext.fillRect(0, 1, 1, 1);
			oImageContext.fillRect(2, 1, 1, 1);
			oImageContext.fillRect(1, 2, 1, 1);
			oImageContext.fill();
		}

		function dp() {
			let nHP = em(nHorPos);
			let nVP = em(nVertPos);
			nX = oT.TransformPointX(nHP, nVP) + 0.5 >> 0;
			nY = oT.TransformPointY(nHP, nVP) + 0.5 >> 0;
			if (bPixel) {
				oContext.drawImage(oImageCanvas, nX, nY);
			} else {
				oContext.drawImage(oImageCanvas, nX - 1, nY - 1);
			}
		}

		nHorStart = this.getStartStridePos(nStrideInsideLine, nSlideWidth);
		nVertStart = this.getStartStridePos(nStrideLine, nSlideHeight);
		let nVertPos = nVertStart;
		let nHorPos;
		let nBottom = nSlideHeight + nStrideInsideLine;
		let nRight = nSlideWidth + nStrideInsideLine;
		while (nVertPos < nBottom) {
			nHorPos = nHorStart;
			while (nHorPos < nRight) {
				dp();
				nHorPos += nStrideInsideLine;
			}
			nVertPos += nStrideLine;
		}
		if (nStrideLine !== nStrideInsideLine) {
			nHorStart = this.getStartStridePos(nStrideLine, nSlideWidth);
			nVertStart = this.getStartStridePos(nStrideInsideLine, nSlideHeight);
			nHorPos = nHorStart;

			while (nHorPos < nRight) {
				nVertPos = nVertStart;
				while (nVertPos < nBottom) {
					dp();
					nVertPos += nStrideInsideLine;
				}
				nHorPos += nStrideLine;
			}
		}

		oGraphics.df();
		oGraphics.RestoreGrState();
	};
	Slide.prototype.drawNotes = function (g) {
		if (this.presentation) {
			this.presentation.checkSlideRecalculateOnOpen(this);
		}
		if (this.notesShape) {
			this.notesShape.draw(g);
			var oLock = this.notesShape.Lock;
			if (oLock && AscCommon.c_oAscLockTypes.kLockTypeNone != oLock.Get_Type()) {
				var bCoMarksDraw = true;
				if (typeof editor !== "undefined" && editor && AscFormat.isRealBool(editor.isCoMarksDraw)) {
					bCoMarksDraw = editor.isCoMarksDraw;
				}
				if (bCoMarksDraw) {
					g.transform3(this.notesShape.transformText);
					var Width = this.notesShape.txBody.content.XLimit - 2;
					Width = Math.max(Width, 1);
					var Height = this.notesShape.txBody.content.GetSummaryHeight();
					g.DrawLockObjectRect(oLock.Get_Type(), 0, 0, Width, Height);
				}
			}
		}
	};
	Slide.prototype.isAnimated = function () {
		let oTr = this.transition;
		if (oTr
			&& oTr.TransitionType !== undefined
			&& oTr.TransitionType !== null
			&& oTr.TransitionType !== c_oAscSlideTransitionTypes.None) {
			return true;
		}
		if (this.timing) {
			if (this.timing.hasEffects()) {
				return true;
			}
		}
		return false;
	};
	Slide.prototype.getTheme = function () {
		return this.Layout && this.Layout.getTheme() || null;
	};
	Slide.prototype.getNum = function () {
		return this.num;
	}

	Slide.prototype.drawNotesSelect = function () {

		if (this.notesShape) {
			var content = this.notesShape.getDocContent();
			if (content) {
				this.presentation.DrawingDocument.UpdateTargetTransform(this.notesShape.transformText);
				content.DrawSelectionOnPage(0);
			}
		}
	};

	Slide.prototype.removeAllCommentsToInterface = function () {
		if (this.slideComments) {
			var aComments = this.slideComments.comments;
			for (var i = aComments.length - 1; i > -1; --i) {
				editor.sync_RemoveComment(aComments[i].Get_Id());
			}
		}
	};

	Slide.prototype.paragraphAdd = function (paraItem, bRecalculate) {
		this.graphicObjects.paragraphAdd(paraItem, bRecalculate);
	};

	Slide.prototype.sendGraphicObjectProps = function () {
		editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
	};

	Slide.prototype.isViewerMode = function () {
		return editor.isViewMode;
	};

	Slide.prototype.onMouseDown = function (e, x, y) {
		this.graphicObjects.onMouseDown(e, x, y);
	};

	Slide.prototype.onMouseMove = function (e, x, y) {
		this.graphicObjects.onMouseMove(e, x, y);
	};

	Slide.prototype.onMouseUp = function (e, x, y) {
		this.graphicObjects.onMouseUp(e, x, y);
	};


	Slide.prototype.getColorMap = function () {
		if (this.clrMap) {
			return this.clrMap;
		} else if (this.Layout) {
			if (this.Layout.clrMap) {
				return this.Layout.clrMap;
			} else if (this.Layout.Master) {
				if (this.Layout.Master.clrMap) {
					return this.Layout.Master.clrMap;
				}
			}
		} else if (this.Master) {
			if (this.Master.clrMap) {
				return this.Master.clrMap;
			}
		}
		return AscFormat.GetDefaultColorMap();
	};

	Slide.prototype.showComment = function (Id, x, y) {
		//editor.sync_HideComment();
		editor.sync_ShowComment(Id, x, y);
	};


	Slide.prototype.getSlideIndex = function () {
		return this.num;
	};

	Slide.prototype.Clear_ContentChanges = function () {
		this.m_oContentChanges.Clear();
	};

	Slide.prototype.Add_ContentChanges = function (Changes) {
		this.m_oContentChanges.Add(Changes);
	};

	Slide.prototype.Refresh_ContentChanges = function () {
		this.m_oContentChanges.Refresh();
	};

	Slide.prototype.checkNoTransformPlaceholder = function () {
		var sp_tree = this.cSld.spTree;
		for (var i = 0; i < sp_tree.length; ++i) {
			var sp = sp_tree[i];
			if (sp.getObjectType() === AscDFH.historyitem_type_Shape || sp.getObjectType() === AscDFH.historyitem_type_ImageShape || sp.getObjectType() === AscDFH.historyitem_type_OleObject) {
				if (sp.isPlaceholder && sp.isPlaceholder()) {
					var hierarchy = sp.getHierarchy();
					for (var j = 0; j < hierarchy.length; ++j) {
						if (AscCommon.isRealObject(hierarchy[j]))
							break;
					}
					if (j === hierarchy.length) {
						AscFormat.CheckSpPrXfrm(sp, true);
					}
				} else if (sp.getObjectType() === AscDFH.historyitem_type_Shape) {
					sp.handleUpdateTheme();
					sp.checkExtentsByDocContent();
				}
			} else {
				if (sp.isGroupObject()) {
					sp.handleUpdateTheme();
					sp.checkExtentsByDocContent();
				}
			}
		}
	};

	Slide.prototype.getSnapArrays = function () {
		var snapX = [];
		var snapY = [];
		for (var i = 0; i < this.cSld.spTree.length; ++i) {
			if (this.cSld.spTree[i].getSnapArrays) {
				this.cSld.spTree[i].getSnapArrays(snapX, snapY);
			}
		}
		return {snapX: snapX, snapY: snapY};
	};


	Slide.prototype.Load_Comments = function (authors) {
		AscCommonSlide.fLoadComments(this, authors);
	};

	Slide.prototype.getDrawingsForController = function () {
		if (this.timing) {
			let aShapes = this.timing.getMoveEffectsShapes();
			if (aShapes && aShapes.length > 0) {
				return this.cSld.spTree.concat(aShapes);
			}
		}
		return this.cSld.spTree;
	};

	Slide.prototype.createFontMap = function (oFontsMap, oCheckedMap, isNoPh) {
		var aSpTree = this.cSld.spTree;
		var nSp, nSpCount = aSpTree.length;
		for (nSp = 0; nSp < nSpCount; ++nSp) {
			aSpTree[nSp].createFontMap(oFontsMap);
		}
		if (this.needMasterSpDraw()) {
			this.Layout.Master.createFontMap(oFontsMap, oCheckedMap, true);
		}
		if (this.needLayoutSpDraw()) {
			this.Layout.createFontMap(oFontsMap, oCheckedMap, true);
		}
		oCheckedMap[this.Get_Id()] = this;
		if (this.notesShape) {
			this.notesShape.createFontMap(oFontsMap);
		}
	};
	Slide.prototype.getAnimationPlayer = function () {
		if (!this.animationPlayer) {
			var oDemoManager = editor.WordControl.DemonstrationManager;
			this.animationPlayer = new AscFormat.CAnimationPlayer(this, oDemoManager);
		}
		return this.animationPlayer;
	};
	Slide.prototype.isAdvanceAfterTransition = function () {
		var oTransition = this.transition;
		if (!oTransition) {
			return false;
		}
		if (this.presentation) {
			var oShowPr = this.presentation.showPr;
			if (oShowPr && oShowPr.useTimings === false) {
				return false;
			}
		}
		return oTransition.SlideAdvanceAfter === true;
	};
	Slide.prototype.getAdvanceDuration = function () {
		var oTransition = this.transition;
		if (!oTransition) {
			return 0;
		}
		return oTransition.SlideAdvanceDuration;
	};
	Slide.prototype.isSlide = function () {
		return true;
	};

	Slide.prototype.IsUseInDocument = function () {
		let oPresentation = Asc.editor.private_GetLogicDocument();
		if (!oPresentation) return false;
		for (let nSld = 0; nSld < oPresentation.Slides.length; ++nSld) {
			if (oPresentation.Slides[nSld] === this) {
				return true;
			}
		}
		return false;
	};
	Slide.prototype.IsUseInSlides = function () {
		return this.IsUseInDocument();
	};
	Slide.prototype.collectRedrawSlides = function (redrawSlides, force) {
		redrawSlides.addRedrawSlide(this);
	};
	Slide.prototype.getOutlineTitleShape = function () {
		let titleShape = null;
		let minIdx = 0;
		for (let i = 0; i < this.cSld.spTree.length; i += 1) {
			const shape = this.cSld.spTree[i];
			if (shape.getObjectType() !== AscDFH.historyitem_type_Shape || !shape.isTitlePlaceholder()) {
				continue;
			}

			const idx = parseInt(shape.getPlaceholderIndex() || "0", 10);
			if (titleShape === null || idx < minIdx) {
				titleShape = shape;
				minIdx = idx;
			}
		}
		return titleShape;
	};
	Slide.prototype.getOutlineSlide = function () {
		const result = new AscCommonSlide.OutlineSlide();
		result.setSlide(this);
		const titleShape = this.getOutlineTitleShape();
		result.setTitle(titleShape);
		for (let i = 0; i < this.cSld.spTree.length; i += 1) {
			const shape = this.cSld.spTree[i];
			if (shape.getObjectType() !== AscDFH.historyitem_type_Shape) {
				continue;
			}
			if (shape.isContentPlaceholder() || shape.isTitlePlaceholder() && shape !== titleShape) {
				result.addContent(shape);
			}
		}
		result.content.sort(function (aShape, bShape) {
			const aIsTitle = aShape.isTitlePlaceholder();
			if (aIsTitle !== bShape.isTitlePlaceholder()) {
				return aIsTitle ? 1 : -1;
			}
			const aPhIndex = aShape.getPlaceholderIndex() || "0";
			const bPhIndex = bShape.getPlaceholderIndex() || "0";
			return parseInt(aPhIndex, 10) - parseInt(bPhIndex, 10);
		});
		return result;
	};

	Slide.prototype.isHaveOutlineShapes = function () {
		for (let i = 0; i < this.cSld.spTree.length; i += 1) {
			const shape = this.cSld.spTree[i];
			if (shape.isOutlinePlaceholder()) {
				return true;
			}
		}
		return false;
	};
	Slide.prototype.isHaveOutlineShapesWithContent = function () {
		for (let i = 0; i < this.cSld.spTree.length; i += 1) {
			const shape = this.cSld.spTree[i];
			if (shape.isOutlinePlaceholderWithContent()) {
				return true;
			}
		}
		return false;
	};
	Slide.prototype.isHaveTitleShape = function () {
		for (let i = 0; i < this.cSld.spTree.length; i += 1) {
			const shape = this.cSld.spTree[i];
			if (shape.isTitlePlaceholder()) {
				return true;
			}
		}
		return false;
	};

	Slide.prototype.createTitle = function () {
		const titleSp = this.getMatchingShape(AscFormat.phType_title, null, false, {});
		if (titleSp) {
			return null;
		}
		let matchingShape = null;
		const layout = this.Layout;
		const master = this.Layout.Master;
		if (layout) {
			matchingShape = layout.getMatchingShape(AscFormat.phType_title, null, false, {});
			if (!matchingShape && master) {
				matchingShape = master.getMatchingShape(AscFormat.phType_title, null, false, {});
			}
		}
		let copySp;
		if (matchingShape) {
			copySp = matchingShape.copy();
		} else {
			copySp = AscCommonSlide.CreatePlaceholder(AscFormat.phType_title);
		}
		this.shapeAdd(this.cSld.spTree.length, copySp);

		return copySp;
	};
	Slide.prototype.createContent = function () {
		let matchingShape;
		matchingShape = this.cSld.getContentShape();
		if (matchingShape) {
			return null;
		}

		const layout = this.Layout;
		const master = this.Layout.Master;

		if (layout) {
			matchingShape = layout.cSld.getContentShape();
			if (!matchingShape && master) {
				matchingShape = master.cSld.getContentShape();
			}
		}
		let copySp;
		if (matchingShape) {
			copySp = matchingShape.copy();
		} else {
			copySp = AscCommonSlide.CreatePlaceholder(AscFormat.phType_body);
		}
		this.shapeAdd(this.cSld.spTree.length, copySp);
		return copySp;
	};
	Slide.prototype.checkOutlineView = function (isAdd) {
		this.presentation && this.presentation.checkSlideOutlineView(this, isAdd);
	};
	Slide.prototype.recalcComments = function () {
		this.presentation && this.presentation.recalcComments();
	};

	function fLoadComments(oObject, authors) {
		var _comments_count = oObject.writecomments.length;
		var _comments_id = [];
		var _comments_data = [];
		var _comments_data_author_id = [];
		var _comments = [];

		var oComments = oObject.slideComments ? oObject.slideComments : oObject.comments;
		if (!oComments) {
			return;
		}
		for (var i = 0; i < _comments_count; i++) {
			var _wc = oObject.writecomments[i];

			if (0 == _wc.WriteParentAuthorId || 0 == _wc.WriteParentCommentId) {
				var commentData = new AscCommon.CSlideCommentData();

				commentData.m_sText = _wc.WriteText;
				commentData.m_sUserId = ("" + _wc.WriteAuthorId);
				commentData.m_sUserName = "";
				commentData.m_sTime = _wc.WriteTime;
				commentData.m_nTimeZoneBias = _wc.timeZoneBias;
				if (commentData.m_sTime && null != commentData.m_nTimeZoneBias) {
					commentData.m_sOOTime = (parseInt(commentData.m_sTime) + commentData.m_nTimeZoneBias * 60000) + "";
				}

				for (var k in authors) {
					if (_wc.WriteAuthorId == authors[k].Id) {
						commentData.m_sUserName = authors[k].Name;
						break;
					}
				}

				//if ("" != commentData.m_sUserName)
				{
					_comments_id.push(_wc.WriteCommentId);
					_comments_data.push(commentData);
					_comments_data_author_id.push(_wc.WriteAuthorId);

					_wc.ParceAdditionalData(commentData);

					var comment = new AscCommon.CSlideComment(oComments, new AscCommon.CSlideCommentData());
					comment.setPosition(_wc.x / 22.66, _wc.y / 22.66);
					_comments.push(comment);
				}
			} else {
				var commentData = new AscCommon.CSlideCommentData();

				commentData.m_sText = _wc.WriteText;
				commentData.m_sUserId = ("" + _wc.WriteAuthorId);
				commentData.m_sUserName = "";
				commentData.m_sTime = _wc.WriteTime;
				commentData.m_nTimeZoneBias = _wc.timeZoneBias;
				if (commentData.m_sTime && null != commentData.m_nTimeZoneBias) {
					commentData.m_sOOTime = (parseInt(commentData.m_sTime) + commentData.m_nTimeZoneBias * 60000) + "";
				}

				for (var k in authors) {
					if (_wc.WriteAuthorId == authors[k].Id) {
						commentData.m_sUserName = authors[k].Name;
						break;
					}
				}

				_wc.ParceAdditionalData(commentData);

				var _parent = null;
				for (var j = 0; j < _comments_data.length; j++) {
					if ((_wc.WriteParentAuthorId == _comments_data_author_id[j]) && (_wc.WriteParentCommentId == _comments_id[j])) {
						_parent = _comments_data[j];
						break;
					}
				}

				if (null != _parent) {
					_parent.m_aReplies.push(commentData);
				}
			}
		}

		for (var i = 0; i < _comments.length; i++) {
			_comments[i].Set_Data(_comments_data[i]);
			oObject.addComment(_comments[i]);
		}

		oObject.writecomments = [];
	}

	AscFormat.CTextBody.prototype.GetAbsoluteStartPage = function () {
		if (this.parent && this.parent.GetAbsoluteStartPage) {
			return this.parent.GetAbsoluteStartPage();
		}
		return 0;
	};
	AscFormat.CTextBody.prototype.GetAbsolutePage = function (CurPage) {
		return this.GetAbsoluteStartPage();
	};
	AscFormat.CTextBody.prototype.GetAbsoluteColumn = function (CurPage) {
		return 0;//TODO;
	};
	AscFormat.CTextBody.prototype.checkCurrentPlaceholder = function () {
		var presentation = editor.WordControl.m_oLogicDocument;
		var oCurController = presentation.GetCurrentController();
		if (oCurController) {
			return oCurController.getTargetDocContent() === this.content;
		}
		return false;
	};


	function SlideComments(slide) {
		this.comments = [];
		this.m_oContentChanges = new AscCommon.CContentChanges(); // list of changes (adding/removing elements)
		this.slide = slide;
		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		g_oTableId.Add(this, this.Id);
	}

	SlideComments.prototype =
		{
			getObjectType: function () {
				return AscDFH.historyitem_type_SlideComments;
			},

			Get_Id: function () {
				return this.Id;
			},

			Clear_ContentChanges: function () {
				this.m_oContentChanges.Clear();
			},

			Add_ContentChanges: function (Changes) {
				this.m_oContentChanges.Add(Changes);
			},

			Refresh_ContentChanges: function () {
				this.m_oContentChanges.Refresh();
			},
			recalcComments: function () {
				this.slide && this.slide.recalcComments();
			},
			addComment: function (comment) {
				History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsAddComment, this.comments.length, [comment], true));
				this.comments.splice(this.comments.length, 0, comment);
				comment.slideComments = this;
			},


			getSlideIndex: function () {
				if (this.slide) {
					return this.slide.num;
				}
				return null;
			},

			changeComment: function (id, commentData) {
				for (var i = 0; i < this.comments.length; ++i) {
					if (this.comments[i].Get_Id() === id) {
						this.comments[i].Set_Data(commentData);
						return;
					}
				}
			},

			removeComment: function (id, bForce) {
				for (var i = 0; i < this.comments.length; ++i) {
					var oComment = this.comments[i];
					if (oComment.Get_Id() === id && (bForce || oComment.canBeDeleted())) {
						History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsRemoveComment, i, this.comments.splice(i, 1), false));
						editor.sync_RemoveComment(id);
						return;
					}
				}
			},

			removeMyComments: function () {
				if (!editor.DocInfo) {
					return;
				}
				var sUserId = editor.DocInfo.get_UserId();
				for (var i = this.comments.length - 1; i > -1; --i) {
					var oComment = this.comments[i];
					var oCommentData = oComment.Data;
					if (oCommentData.m_sUserId === sUserId) {
						if (oComment.canBeDeleted()) {
							History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsRemoveComment, i, this.comments.splice(i, 1), false));
							editor.sync_RemoveComment(oComment.Get_Id());
						}
					} else {
						oComment.removeUserReplies(sUserId);
					}
				}
			},

			removeAllComments: function () {
				for (var i = this.comments.length - 1; i > -1; --i) {
					var oComment = this.comments[i];
					if (oComment.canBeDeleted()) {
						History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsRemoveComment, i, this.comments.splice(i, 1), false));
						editor.sync_RemoveComment(oComment.Get_Id());
					}
				}
			},

			removeSelectedComment: function () {
				var comment = this.getSelectedComment();
				if (comment) {
					this.removeComment(comment.Get_Id());
				}
			},

			getSelectedComment: function () {
				for (var i = 0; i < this.comments.length; ++i) {
					if (this.comments[i].selected) {
						return this.comments[i];
					}
				}
				return null;
			},

			getAllMyComments: function (aAllComments, oSlide) {
				if (!editor.DocInfo) {
					return;
				}
				var sUserId = editor.DocInfo.get_UserId();
				for (var i = 0; i < this.comments.length; ++i) {
					var oComment = this.comments[i];
					var oCommentData = oComment.Data;
					if (oCommentData.m_sUserId === sUserId) {
						aAllComments.push({comment: oComment, slide: oSlide});
					} else {
						for (var j = 0; j < oCommentData.m_aReplies.length; ++j) {
							if (oCommentData.m_aReplies[j].m_sUserId === sUserId) {
								aAllComments.push({comment: oComment, slide: oSlide});
								break;
							}
						}
					}
				}
			},
			getAllComments  : function (aAllComments, isMine, isCurrent, aIds) {
				var oSlide = null;
				if (this.slide && (this.slide instanceof Slide)) {
					oSlide = this.slide;
				}
				var nComment, oComment;
				if (Array.isArray(aIds)) {
					var oIdMap = {};
					for (var nId = 0; nId < aIds.length; ++nId) {
						oIdMap[aIds[nId]] = true;
					}
					for (nComment = 0; nComment < this.comments.length; ++nComment) {
						oComment = this.comments[nComment];
						if (oIdMap[oComment.Id]) {
							aAllComments.push({comment: oComment, slide: oSlide});
						}
					}
				} else {
					if (isCurrent) {
						if (oSlide) {
							var oPresentation = oSlide.presentation;
							if (oPresentation && oPresentation.Slides[oPresentation.CurPage] === oSlide) {
								for (nComment = 0; nComment < this.comments.length; ++nComment) {
									oComment = this.comments[nComment];
									if (oComment.selected && (!isMine || oComment.isMineComment())) {
										aAllComments.push({comment: oComment, slide: oSlide});
										return;
									}
								}
							}
						}
					} else {
						for (nComment = 0; nComment < this.comments.length; ++nComment) {
							oComment = this.comments[nComment];
							if ((!isMine || oComment.isMineComment())) {
								aAllComments.push({comment: oComment, slide: oSlide});
							}
						}
					}
				}
			},

			recalculate: function () {},

			Write_ToBinary2: function (w) {
				w.WriteLong(AscDFH.historyitem_type_SlideComments);
				w.WriteString2(this.Id);
				AscFormat.writeObject(w, this.slide);
			},

			Read_FromBinary2: function (r) {
				this.Id = r.GetString2();
				this.slide = AscFormat.readObject(r);
			},

			Refresh_RecalcData: function () {
				History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
				this.recalcComments();
			},

			collectRedrawSlides: function (redrawSlides, force) {
				if (this.slide) {
					redrawSlides.collectRedrawSlides(this.slide, force);
				}
			}
		};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonSlide'].Slide = Slide;
	window['AscCommonSlide'].SlideComments = SlideComments;
	window['AscCommonSlide'].fLoadComments = fLoadComments;
	window['AscCommonSlide'].CStrideData = CStrideData;

})();
