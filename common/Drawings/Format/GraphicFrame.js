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

(function (window, undefined) {

// Import
	var CShape = AscFormat.CShape;
	var HitInLine = AscFormat.HitInLine;

	var isRealObject = AscCommon.isRealObject;
	

	window['AscDFH'].drawingsChangesMap[AscDFH.historyitem_GraphicFrameSetSpPr] = function (oClass, value) {
		oClass.spPr = value;
	};
	window['AscDFH'].drawingsChangesMap[AscDFH.historyitem_GraphicFrameSetGraphicObject] = function (oClass, value) {
		oClass.graphicObject = value;
		if (value) {
			value.Parent = oClass;
			oClass.graphicObject.Index = 0;
		}
	};
	window['AscDFH'].drawingsChangesMap[AscDFH.historyitem_GraphicFrameSetSetNvSpPr] = function (oClass, value) {
		oClass.nvGraphicFramePr = value;
	};
	window['AscDFH'].drawingsChangesMap[AscDFH.historyitem_GraphicFrameSetSetParent] = function (oClass, value) {
		oClass.oldParent = oClass.parent;
		oClass.parent = value;
	};
	window['AscDFH'].drawingsChangesMap[AscDFH.historyitem_GraphicFrameSetSetGroup] = function (oClass, value) {
		oClass.group = value;
	};

	AscDFH.changesFactory[AscDFH.historyitem_GraphicFrameSetSpPr] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_GraphicFrameSetGraphicObject] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_GraphicFrameSetSetNvSpPr] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_GraphicFrameSetSetParent] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_GraphicFrameSetSetGroup] = AscDFH.CChangesDrawingsObject;

	function CGraphicFrame() {
		AscFormat.CGraphicObjectBase.call(this);
		this.graphicObject = null;
		this.nvGraphicFramePr = null;
		this.Pages = [];
		this.compiledStyles = [];
		this.recalcInfo =
			{
				recalculateTransform: true,
				recalculateSizes: true,
				recalculateNumbering: true,
				recalculateTable: true
			};
		this.RecalcInfo = {};
	}

	AscFormat.InitClass(CGraphicFrame, AscFormat.CGraphicObjectBase, AscDFH.historyitem_type_GraphicFrame);

	CGraphicFrame.prototype.addToRecalculate = CShape.prototype.addToRecalculate;

	CGraphicFrame.prototype.Get_Theme = CShape.prototype.Get_Theme;

	CGraphicFrame.prototype.Get_ColorMap = CShape.prototype.Get_ColorMap;

	CGraphicFrame.prototype.getSlideIndex = CShape.prototype.getSlideIndex;
	CGraphicFrame.prototype.IsUseInDocument = CShape.prototype.IsUseInDocument;
	CGraphicFrame.prototype.convertPixToMM = CShape.prototype.convertPixToMM;
	CGraphicFrame.prototype.hit = CShape.prototype.hit;
	CGraphicFrame.prototype.hitInPath = function () {
		return false;
	};
	CGraphicFrame.prototype.hitInInnerArea = function (x, y) {
		let oInvTransform = this.getInvertTransform();
		if (!oInvTransform) {
			return false;
		}
		let tx = oInvTransform.TransformPointX(x, y);
		let ty = oInvTransform.TransformPointY(x, y);
		return tx > 0 && tx < this.extX && ty > 0 && ty < this.extY;
	};

	CGraphicFrame.prototype.hitInTextRect = function (x, y) {
		return this.hitInInnerArea(x, y);
	};
	CGraphicFrame.prototype.hitInTextHyperlink = function(x, y) {
		if(!this.graphicObject || !(this.graphicObject instanceof AscWord.CTable)) {
			return null;
		}
		let oInvTransform = this.getInvertTransform();
		if (!oInvTransform) {
			return false;
		}
		let tx = oInvTransform.TransformPointX(x, y);
		let ty = oInvTransform.TransformPointY(x, y);

		let oCellPos = this.graphicObject.private_GetCellByXY(tx, ty, 0);
		if(!oCellPos) {
			return null;
		}
		let oRow = this.graphicObject.GetRow(oCellPos.Row);
		if(!oRow) {
			return null;
		}
		let oCell    = oRow.GetCell(oCellPos.Cell);
		let oContent = oCell.Content;
		if(!oContent) {
			return null;
		}
		let oHitParagraph = oContent.IsInText(tx, ty, 0);
		if (oHitParagraph) {
			if (oHitParagraph.IsInText(tx, ty, 0)) {
				let oHyperlink = oHitParagraph.CheckHyperlink(tx, ty, 0);
				if (oHyperlink) {
					return oHyperlink;
				}
			}
		}
		return null;
	};
	CGraphicFrame.prototype.GetDocumentPositionFromObject = function (arrPos) {
		if (!arrPos)
			arrPos = [];
		return arrPos;
	};

	CGraphicFrame.prototype.Is_DrawingShape = function (bRetShape) {
		if (bRetShape === true) {
			return null;
		}
		return false;
	};
	CGraphicFrame.prototype.GetSelectedText = function(bClearText, oPr) {
		if (this.graphicObject && this.graphicObject.GetSelectedText) {
			return this.graphicObject.GetSelectedText(bClearText, oPr);
		}
		return "";
	};
	CGraphicFrame.prototype.IsTableFirstRowOnNewPage = function () {
		return false;
	};
	CGraphicFrame.prototype.IsBlockLevelSdtContent = function () {
		return false;
	};
	CGraphicFrame.prototype.IsBlockLevelSdtFirstOnNewPage = function () {
		return false;
	};

	CGraphicFrame.prototype.handleUpdatePosition = function () {
		this.recalcInfo.recalculateTransform = true;
		this.addToRecalculate();
	};

	CGraphicFrame.prototype.handleUpdateTheme = function () {
		this.compiledStyles = [];
		if (this.graphicObject) {
			this.graphicObject.Recalc_CompiledPr2();
			this.graphicObject.RecalcInfo.Recalc_AllCells();
			this.recalcInfo.recalculateSizes = true;
			this.recalcInfo.recalculateTable = true;
			this.addToRecalculate();
		}
	};

	CGraphicFrame.prototype.handleUpdateFill = function () {
	};

	CGraphicFrame.prototype.handleUpdateLn = function () {
	};

	CGraphicFrame.prototype.handleUpdateExtents = function () {
		this.recalcInfo.recalculateTransform = true;
		this.addToRecalculate();
	};

	CGraphicFrame.prototype.recalcText = function () {
		this.compiledStyles = [];
		if (this.graphicObject) {
			this.graphicObject.Recalc_CompiledPr2();
			this.graphicObject.RecalcInfo.Reset(true);
		}
		this.recalcInfo.recalculateTable = true;
		this.recalcInfo.recalculateSizes = true;
	};

	CGraphicFrame.prototype.Get_TextBackGroundColor = function () {
		return undefined;
	};

	CGraphicFrame.prototype.GetPrevElementEndInfo = function () {
		return null;
	};

	CGraphicFrame.prototype.Get_PageFields = function () {
		return editor.WordControl.m_oLogicDocument.Get_PageFields();
	};

	CGraphicFrame.prototype.Get_ParentTextTransform = function () {
		return this.transformText;
	};

	CGraphicFrame.prototype.getDocContent = function () {
		if (this.graphicObject && this.graphicObject.CurCell && (false === this.graphicObject.Selection.Use || (true === this.graphicObject.Selection.Use && table_Selection_Text === this.graphicObject.Selection.Type))) {
			return this.graphicObject.CurCell.Content;
		}
		return null;
	};

	CGraphicFrame.prototype.setSpPr = function (spPr) {
		AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GraphicFrameSetSpPr, this.spPr, spPr));
		this.spPr = spPr;
		if (spPr) {
			spPr.setParent(this);
		}
	};

	CGraphicFrame.prototype.setGraphicObject = function (graphicObject) {
		AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GraphicFrameSetGraphicObject, this.graphicObject, graphicObject));
		this.graphicObject = graphicObject;
		if (this.graphicObject) {
			this.graphicObject.Index = 0;
			this.graphicObject.Parent = this;
		}
	};

	CGraphicFrame.prototype.setNvSpPr = function (pr) {
		AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GraphicFrameSetSetNvSpPr, this.nvGraphicFramePr, pr));
		this.nvGraphicFramePr = pr;
	};

	CGraphicFrame.prototype.setParent = function (parent) {
		AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GraphicFrameSetSetParent, this.parent, parent));
		this.parent = parent;
	};

	CGraphicFrame.prototype.setGroup = function (group) {
		AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GraphicFrameSetSetGroup, this.group, group));
		this.group = group;
	};

	CGraphicFrame.prototype.Search = function (SearchEngine, Type) {
		if (this.graphicObject) {
			this.graphicObject.Search(SearchEngine, Type);
		}
	};

	CGraphicFrame.prototype.GetSearchElementId = function (bNext, bCurrent) {
		if (this.graphicObject) {
			return this.graphicObject.GetSearchElementId(bNext, bCurrent);
		}

		return null;
	};

	CGraphicFrame.prototype.FindNextFillingForm = function (isNext, isCurrent) {
		if (this.graphicObject)
			return this.graphicObject.FindNextFillingForm(isNext, isCurrent);

		return null;
	};

	CGraphicFrame.prototype.copy = function (oPr) {
		var ret = new CGraphicFrame();
		if (this.graphicObject) {
			ret.setGraphicObject(this.graphicObject.Copy(ret));
			if (editor && editor.WordControl && editor.WordControl.m_oLogicDocument && isRealObject(editor.WordControl.m_oLogicDocument.globalTableStyles)) {
				ret.graphicObject.Reset(0, 0, this.graphicObject.XLimit, this.graphicObject.YLimit, ret.graphicObject.PageNum);
			}
		}
		if (this.nvGraphicFramePr) {
			ret.setNvSpPr(this.nvGraphicFramePr.createDuplicate());
		}
		if (this.spPr) {
			ret.setSpPr(this.spPr.createDuplicate());
			ret.spPr.setParent(ret);
		}
		ret.setBDeleted(false);
		if (this.macro !== null) {
			ret.setMacro(this.macro);
		}
		if (this.textLink !== null) {
			ret.setTextLink(this.textLink);
		}
		if (!this.recalcInfo.recalculateTable && !this.recalcInfo.recalculateSizes && !this.recalcInfo.recalculateTransform) {
			if (!oPr || false !== oPr.cacheImage) {
				ret.cachedImage = this.getBase64Img();
				ret.cachedPixH = this.cachedPixH;
				ret.cachedPixW = this.cachedPixW;
			}
		}
		return ret;
	};

	CGraphicFrame.prototype.getAllFonts = function (fonts) {
		if (this.graphicObject) {
			for (var i = 0; i < this.graphicObject.Content.length; ++i) {
				var row = this.graphicObject.Content[i];
				var cells = row.Content;
				for (var j = 0; j < cells.length; ++j) {
					cells[j].Content.Document_Get_AllFontNames(fonts);
				}
			}
			delete fonts["+mj-lt"];
			delete fonts["+mn-lt"];
			delete fonts["+mj-ea"];
			delete fonts["+mn-ea"];
			delete fonts["+mj-cs"];
			delete fonts["+mn-cs"];
		}
	};

	CGraphicFrame.prototype.MoveCursorToStartPos = function () {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.MoveCursorToStartPos();
			this.graphicObject.RecalculateCurPos();

		}
	};

	CGraphicFrame.prototype.MoveCursorToEndPos = function () {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.MoveCursorToEndPos();
			this.graphicObject.RecalculateCurPos();

		}
	};


	CGraphicFrame.prototype.pasteFormatting = function (oFormatData) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.PasteFormatting(oFormatData);

			this.recalcInfo.recalculateContent = true;
			this.recalcInfo.recalculateTransformText = true;
			editor.WordControl.m_oLogicDocument.recalcMap[this.Id] = this;
		}

	};

	CGraphicFrame.prototype.ClearParagraphFormatting = function (isClearParaPr, isClearTextPr) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.ClearParagraphFormatting(isClearParaPr, isClearTextPr);

			this.recalcInfo.recalculateContent = true;
			this.recalcInfo.recalculateTransformText = true;
			editor.WordControl.m_oLogicDocument.recalcMap[this.Id] = this;
		}

	};

	CGraphicFrame.prototype.Set_Props = function (props) {
		if (this.graphicObject) {
			var bApplyToAll = this.parent.graphicObjects.State.textObject !== this;
			// if(bApplyToAll)
			//     this.graphicObject.SetApplyToAll(true);
			this.graphicObject.Set_Props(props, bApplyToAll);
			//if(bApplyToAll)
			//    this.graphicObject.SetApplyToAll(false);
			this.OnContentRecalculate();
			editor.WordControl.m_oLogicDocument.recalcMap[this.Id] = this;
		}
	};

	CGraphicFrame.prototype.updateCursorType = function (x, y, e) {
		var oApi = Asc.editor || editor;
		var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
		if (isDrawHandles === false) {
			return;
		}
		var tx = this.invertTransform.TransformPointX(x, y);
		var ty = this.invertTransform.TransformPointY(x, y);
		this.graphicObject.UpdateCursorType(tx, ty, 0)
	};

	CGraphicFrame.prototype.getIsSingleBody = CShape.prototype.getIsSingleBody;

	CGraphicFrame.prototype.getHierarchy = CShape.prototype.getHierarchy;

	CGraphicFrame.prototype.getAllImages = function (images) {
	};
	CGraphicFrame.prototype.recalculateTable = function () {
		if (this.graphicObject) {
			this.graphicObject.Set_PositionH(Asc.c_oAscHAnchor.Page, false, 0, false);
			this.graphicObject.Set_PositionV(Asc.c_oAscVAnchor.Page, false, 0, false);
			this.graphicObject.Parent = this;
			this.graphicObject.Reset(0, 0, this.extX, 10000, 0);
			this.graphicObject.Recalculate_Page(0);
		}
	};

	CGraphicFrame.prototype.recalculate = function () {
		if (this.bDeleted || !this.parent)
			return;
		AscFormat.ExecuteNoHistory(function () {


			if (this.recalcInfo.recalculateTable) {
				this.extX = this.getXfrmExtX();
				this.recalculateTable();
				this.recalcInfo.recalculateTable = false;
			}
			let bSizes = this.recalcInfo.recalculateSizes;
			if(this.recalcInfo.recalculateSizes) {
				this.recalculateSizes();
				this.recalcInfo.recalculateSizes = false;
			}
			if (this.recalcInfo.recalculateTransform) {
				this.recalculateTransform();
				this.recalculateSnapArrays();
				this.recalcInfo.recalculateTransform = false;
				this.transformText = this.transform;
				this.invertTransformText = this.invertTransform;
				this.cachedImage = null;
				this.recalcInfo.recalculateSizes = true;
			}
			if (bSizes) {
				this.bounds.l = this.x;
				this.bounds.t = this.y;
				this.bounds.r = this.x + this.extX;
				this.bounds.b = this.y + this.extY;
				this.bounds.x = this.x;
				this.bounds.y = this.y;
				this.bounds.w = this.extX;
				this.bounds.h = this.extY;
			}
		}, this, []);

	};

	CGraphicFrame.prototype.recalculateSizes = function () {
		if (this.graphicObject) {
			this.graphicObject.XLimit -= this.graphicObject.X;
			this.graphicObject.X = 0;
			this.graphicObject.Y = 0;
			this.graphicObject.X_origin = 0;
			var _page_bounds = this.graphicObject.Get_PageBounds(0);
			this.extX = _page_bounds.Right - _page_bounds.Left;
			this.extY = _page_bounds.Bottom - _page_bounds.Top;
		}
	};

	CGraphicFrame.prototype.IsSelectedSingleElement = function () {
		return true;
	};

	CGraphicFrame.prototype.recalculateCurPos = function () {
		this.graphicObject.RecalculateCurPos();
	};

	CGraphicFrame.prototype.getTypeName = function () {
		if (this.isTable()) {
			return AscCommon.translateManager.getValue("Table");
		}
		return AscFormat.CGraphicObjectBase.prototype.getTypeName.call(this)
	};

	CGraphicFrame.prototype.CanAddHyperlink = function (bCheck) {
		if (this.graphicObject)
			return this.graphicObject.CanAddHyperlink(bCheck);
		return false;
	};

	CGraphicFrame.prototype.IsCursorInHyperlink = function (bCheck) {
		if (this.graphicObject)
			return this.graphicObject.IsCursorInHyperlink(bCheck);
		return false;
	};

	CGraphicFrame.prototype.getTransformMatrix = function () {
		return this.transform;
		if (this.recalcInfo.recalculateTransform) {
			this.recalculateTransform();
			this.recalcInfo.recalculateTransform = false;
		}
		return this.transform;
	};

	CGraphicFrame.prototype.OnContentReDraw = function () {
	};

	CGraphicFrame.prototype.changeSize = function (kw, kh) {
		if (this.spPr && this.spPr.xfrm && this.spPr.xfrm.isNotNull()) {
			var xfrm = this.spPr.xfrm;
			xfrm.setOffX(xfrm.offX * kw);
			xfrm.setOffY(xfrm.offY * kh);
		}
		this.recalcTransform && this.recalcTransform();
	};

	CGraphicFrame.prototype.recalcTransform = function () {
		this.recalcInfo.recalculateTransform = true;
	};

	CGraphicFrame.prototype.getTransform = function () {
		if (this.recalcInfo.recalculateTransform) {
			this.recalculateTransform();
			this.recalcInfo.recalculateTransform = false;
		}
		return {
			x: this.x,
			y: this.y,
			extX: this.extX,
			extY: this.extY,
			rot: this.rot,
			flipH: this.flipH,
			flipV: this.flipV
		};
	};

	CGraphicFrame.prototype.canRotate = function () {
		return false;
	};


	CGraphicFrame.prototype.canGroup = function () {
		return false;
	};

	CGraphicFrame.prototype.createRotateTrack = function () {
		return new AscFormat.RotateTrackShapeImage(this);
	};

	CGraphicFrame.prototype.createResizeTrack = function (cardDirection) {
		return new AscFormat.ResizeTrackShapeImage(this, cardDirection);
	};

	CGraphicFrame.prototype.createMoveTrack = function () {
		return new AscFormat.MoveShapeImageTrack(this);
	};

	CGraphicFrame.prototype.getSnapArrays = function (snapX, snapY) {
		var transform = this.getTransformMatrix();
		snapX.push(transform.tx);
		snapX.push(transform.tx + this.extX * 0.5);
		snapX.push(transform.tx + this.extX);
		snapY.push(transform.ty);
		snapY.push(transform.ty + this.extY * 0.5);
		snapY.push(transform.ty + this.extY);
	};



	CGraphicFrame.prototype.getInvertTransform = function () {
		if (this.recalcInfo.recalculateTransform)
			this.recalculateTransform();
		return this.invertTransform;
	};

	CGraphicFrame.prototype.Document_UpdateRulersState = function (margins) {
		if (this.graphicObject) {
			this.graphicObject.Document_UpdateRulersState(this.getParentNum());
		}
	};

	CGraphicFrame.prototype.Get_PageLimits = function (PageIndex) {
		return {X: 0, Y: 0, XLimit: Page_Width, YLimit: Page_Height};
	};

	CGraphicFrame.prototype.getParentObjects = CShape.prototype.getParentObjects;

	CGraphicFrame.prototype.IsHdrFtr = function (bool) {
		if (bool)
			return null;

		return false;
	};
	CGraphicFrame.prototype.resize = function (extX, extY) {
		var newExtX = AscFormat.isRealNumber(extX) ? extX : this.extX;
		var newExtY = AscFormat.isRealNumber(extY) ? extY : this.extY;
		if (!AscFormat.fApproxEqual(newExtX, this.extX) || !AscFormat.fApproxEqual(newExtY, this.extY)) {
			this.graphicObject.Resize(newExtX, newExtY);
			this.recalculateTable();
			this.recalculateSizes();
			return true;
		}
		return false;
	};
	CGraphicFrame.prototype.setFrameTransform = function (oPr) {
		var bResult = this.resize(oPr.FrameWidth, oPr.FrameHeight);
		var newX = AscFormat.isRealNumber(oPr.FrameX) ? oPr.FrameX : this.x;
		var newY = AscFormat.isRealNumber(oPr.FrameY) ? oPr.FrameY : this.y;
		this.setNoChangeAspect(oPr.FrameLockAspect ? true : undefined);
		if (!AscFormat.fApproxEqual(newX, this.x) || !AscFormat.fApproxEqual(newY, this.y)) {
			AscFormat.CheckSpPrXfrm(this, true);
			var xfrm = this.spPr.xfrm;
			xfrm.setOffX(newX);
			xfrm.setOffY(newY);
			bResult = true;
			this.recalculate();
		}
		return bResult;
	};

	CGraphicFrame.prototype.IsFootnote = function (bReturnFootnote) {
		if (bReturnFootnote)
			return null;

		return false;
	};

	CGraphicFrame.prototype.IsTableCellContent = function (isReturnCell) {
		if (true === isReturnCell)
			return null;

		return false;
	};

	CGraphicFrame.prototype.Check_AutoFit = function () {
		return false;
	};

	CGraphicFrame.prototype.IsInTable = function () {
		return null;
	};

	CGraphicFrame.prototype.selectionSetStart = function (e, x, y, slideIndex) {
		if (AscCommon.g_mouse_button_right === e.Button) {
			this.rightButtonFlag = true;
			return;
		}
		if (isRealObject(this.graphicObject)) {
			var tx, ty;
			tx = this.invertTransform.TransformPointX(x, y);
			ty = this.invertTransform.TransformPointY(x, y);
			if (AscCommon.g_mouse_event_type_down === e.Type) {
				if (this.graphicObject.IsTableBorder(tx, ty, 0)) {
					if (editor.WordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props) !== false) {
						return;
					} else {
					}
				}
			}

			if (!(/*content.IsTextSelectionUse() && */e.ShiftKey)) {
				if (editor.WordControl.m_oLogicDocument.CurPosition) {
					editor.WordControl.m_oLogicDocument.CurPosition.X = tx;
					editor.WordControl.m_oLogicDocument.CurPosition.Y = ty;
				}
				this.graphicObject.Selection_SetStart(tx, ty, this.getParentNum(), e);
			} else {
				if (!this.graphicObject.IsSelectionUse()) {
					this.graphicObject.StartSelectionFromCurPos();
				}
				this.graphicObject.Selection_SetEnd(tx, ty, this.getParentNum(), e);
			}
			this.graphicObject.RecalculateCurPos();

		}
	};

	CGraphicFrame.prototype.selectionSetEnd = function (e, x, y, slideIndex) {
		if (AscCommon.g_mouse_event_type_move === e.Type) {
			this.rightButtonFlag = false;
		}
		if (this.rightButtonFlag && AscCommon.g_mouse_event_type_up === e.Type) {
			this.rightButtonFlag = false;
			return;
		}
		if (isRealObject(this.graphicObject)) {
			var tx, ty;
			tx = this.invertTransform.TransformPointX(x, y);
			ty = this.invertTransform.TransformPointY(x, y);
			//var bBorder = this.graphicObject.Selection.Type2 === table_Selection_Border;
			this.graphicObject.Selection_SetEnd(tx, ty, 0, e);
			//if(g_mouse_event_type_up === e.Type && bBorder)
			//    editor.WordControl.m_oLogicDocument.Recalculate();  TODO: пересчет вызывается в CTable
		}
	};

	CGraphicFrame.prototype.updateSelectionState = function () {
		if (isRealObject(this.graphicObject)) {
			var drawingDocument = this.parent.presentation.DrawingDocument;
			var Doc = this.graphicObject;
			if (true === Doc.IsSelectionUse() && !Doc.IsSelectionEmpty()) {
				drawingDocument.UpdateTargetTransform(this.transform);
				drawingDocument.TargetEnd();
				drawingDocument.SelectEnabled(true);
				drawingDocument.SelectClear();
				Doc.DrawSelectionOnPage(0);
				drawingDocument.SelectShow();
			} else {
				drawingDocument.SelectEnabled(false);
				Doc.RecalculateCurPos();
				drawingDocument.UpdateTargetTransform(this.transform);
				drawingDocument.TargetShow();
			}
		} else {
			this.parent.presentation.DrawingDocument.UpdateTargetTransform(null);
			this.parent.presentation.DrawingDocument.TargetEnd();
			this.parent.presentation.DrawingDocument.SelectEnabled(false);
			this.parent.presentation.DrawingDocument.SelectClear();
			this.parent.presentation.DrawingDocument.SelectShow();
		}
	};

	CGraphicFrame.prototype.Get_AbsolutePage = function (CurPage) {
		return this.Get_StartPage_Absolute();
	};

	CGraphicFrame.prototype.Get_AbsoluteColumn = function (CurPage) {
		return 0;
	};

	CGraphicFrame.prototype.Is_TopDocument = function () {
		return false;
	};

	CGraphicFrame.prototype.GetTopElement = function () {
		return null;
	};

	CGraphicFrame.prototype.drawAdjustments = function () {
	};

	CGraphicFrame.prototype.recalculateTransform = CShape.prototype.recalculateTransform;

	CGraphicFrame.prototype.recalculateLocalTransform = CShape.prototype.recalculateLocalTransform;


	CGraphicFrame.prototype.Update_ContentIndexing = function () {
	};

	CGraphicFrame.prototype.GetTopDocumentContent = function (isOneLevel) {
		return null;
	};
	CGraphicFrame.prototype.GetElement = function (nIndex) {
		return this.graphicObject;
	};

	CGraphicFrame.prototype.draw = function (graphics) {
		if (graphics.isBoundsChecker() === true) {
			graphics.transform3(this.transform);
			graphics._s();
			graphics._m(0, 0);
			graphics._l(this.extX, 0);
			graphics._l(this.extX, this.extY);
			graphics._l(0, this.extY);
			graphics._e();
			return;
		}
		if (graphics.animationDrawer) {
			graphics.animationDrawer.drawObject(this, graphics);
			return;
		}
		if (this.graphicObject) {
			graphics.SaveGrState();
			graphics.transform3(this.transform);
			graphics.SetIntegerGrid(false);
			this.graphicObject.Draw(0, graphics);
			this.drawLocks(this.transform, graphics);
			graphics.RestoreGrState();
		}
	};

	CGraphicFrame.prototype.Select = function () {
	};

	CGraphicFrame.prototype.Set_CurrentElement = function () {
		if (this.parent && this.parent.graphicObjects) {
			this.parent.graphicObjects.resetSelection(true);
			if (this.group) {
				var main_group = this.group.getMainGroup();
				this.parent.graphicObjects.selectObject(main_group, 0);
				main_group.selectObject(this, 0);
				main_group.selection.textSelection = this;
			} else {
				this.parent.graphicObjects.selectObject(this, 0);
				this.parent.graphicObjects.selection.textSelection = this;
			}
			if (editor.WordControl.m_oLogicDocument.CurPage !== this.getParentNum()) {
				editor.WordControl.m_oLogicDocument.Set_CurPage(this.getParentNum());
				editor.WordControl.GoToPage(this.getParentNum());
			}
		}
	};

	CGraphicFrame.prototype.OnContentRecalculate = function () {
		this.recalcInfo.recalculateSizes = true;
		this.recalcInfo.recalculateTransform = true;
		editor.WordControl.m_oLogicDocument.Document_UpdateRulersState();
	};

	CGraphicFrame.prototype.getTextSelectionState = function () {
		return this.graphicObject.GetSelectionState();
	};

	CGraphicFrame.prototype.setTextSelectionState = function (Sate) {
		return this.graphicObject.SetSelectionState(Sate, Sate.length - 1);
	};

	CGraphicFrame.prototype.paragraphAdd = function (paraItem, bRecalculate) {
	};

	CGraphicFrame.prototype.applyTextFunction = function (docContentFunction, tableFunction, args) {
		if (tableFunction === CTable.prototype.AddToParagraph) {
			if ((args[0].Type === para_NewLine
					|| args[0].Type === para_Text
					|| args[0].Type === para_Space
					|| args[0].Type === para_Tab
					|| args[0].Type === para_PageNum)
				&& this.graphicObject.Selection.Use) {
				this.graphicObject.Remove(1, true, undefined, true);
			}
		} else if (tableFunction === CTable.prototype.AddNewParagraph) {
			this.graphicObject.Selection.Use && this.graphicObject.Remove(1, true, undefined, true);
		}
		tableFunction.apply(this.graphicObject, args);
	};

	CGraphicFrame.prototype.remove = function (Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord) {
		this.graphicObject.Remove(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord);
		this.recalcInfo.recalculateSizes = true;
		this.recalcInfo.recalculateTransform = true;
	};

	CGraphicFrame.prototype.addNewParagraph = function () {
		this.graphicObject.AddNewParagraph(false);
		this.recalcInfo.recalculateContent = true;
		this.recalcInfo.recalculateTransformText = true;
	};

	CGraphicFrame.prototype.setParagraphAlign = function (val) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.SetParagraphAlign(val);
			this.recalcInfo.recalculateContent = true;
			this.recalcInfo.recalculateTransform = true;
		}
	};

	CGraphicFrame.prototype.applyAllAlign = function (val) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.SetApplyToAll(true);
			this.graphicObject.SetParagraphAlign(val);
			this.graphicObject.SetApplyToAll(false);
		}
	};

	CGraphicFrame.prototype.setParagraphSpacing = function (val) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.SetParagraphSpacing(val);
		}
	};

	CGraphicFrame.prototype.applyAllSpacing = function (val) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.SetApplyToAll(true);
			this.graphicObject.SetParagraphSpacing(val);
			this.graphicObject.SetApplyToAll(false);
		}
	};

	CGraphicFrame.prototype.setParagraphNumbering = function (val) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.SetParagraphNumbering(val);
		}
	};

	CGraphicFrame.prototype.setParagraphIndent = function (val) {
		if (isRealObject(this.graphicObject)) {
			this.graphicObject.SetParagraphIndent(val);
		}
	};

	CGraphicFrame.prototype.setWordFlag = function (bPresentation, Document) {
		if (this.graphicObject) {
			this.graphicObject.bPresentation = bPresentation;
			for (var i = 0; i < this.graphicObject.Content.length; ++i) {
				var row = this.graphicObject.Content[i];
				for (var j = 0; j < row.Content.length; ++j) {
					var content = row.Content[j].Content;
					if (!bPresentation && Document) {
						content.Styles = Document.Styles;
					} else {
						content.Styles = null;
					}
					content.bPresentation = bPresentation;
					for (var k = 0; k < content.Content.length; ++k) {
						content.Content[k].bFromDocument = !bPresentation;
					}
				}
			}
		}
	};
	CGraphicFrame.prototype.getWordTable = function () {
		if (!this.graphicObject) {
			return null;
		}
		var bOldBDeleted = this.bDeleted;
		this.bDeleted = true;
		this.setWordFlag(false);
		var oTable = this.graphicObject.Copy();
		oTable.Set_TableStyle2(undefined);
		this.setWordFlag(true);
		this.bDeleted = bOldBDeleted;
		var aRows = oTable.Content;
		var aCells;
		var nRow, nCell;
		var oRow, oCell;
		var oShd;
		for (nRow = 0; nRow < aRows.length; ++nRow) {
			oRow = aRows[nRow];
			aCells = oRow.Content;
			for (nCell = 0; nCell < aCells.length; ++nCell) {
				oCell = aCells[nCell];
				if (oCell.Pr) {
					if (oCell.Pr.Shd) {
						oShd = oCell.Pr.Shd;
						oCell.Set_Shd(ConvertToWordTableShd(oShd));
					}

					if (oCell.Pr.TableCellBorders) {
						var oBorders = oCell.Pr.TableCellBorders;
						// 0 - Top, 1 - Right, 2- Bottom, 3- Left
						oCell.Set_Border(ConvertToWordTableBorder(oBorders.Top), 0);
						oCell.Set_Border(ConvertToWordTableBorder(oBorders.Right), 1);
						oCell.Set_Border(ConvertToWordTableBorder(oBorders.Bottom), 2);
						oCell.Set_Border(ConvertToWordTableBorder(oBorders.Left), 3);
					}
				}
			}
		}
		return oTable;

	};
	CGraphicFrame.prototype.Get_Styles = function (level) {
		if (AscFormat.isRealNumber(level)) {
			if (!this.compiledStyles[level]) {
				CShape.prototype.recalculateTextStyles.call(this, level);
			}
			return this.compiledStyles[level];
		} else {
			let oLogicDoc;
			if(Asc.editor.private_GetLogicDocument) {
				oLogicDoc = Asc.editor.private_GetLogicDocument();
			}
			if(oLogicDoc) {

				if(oLogicDoc.globalTableStyles) {
					return oLogicDoc.globalTableStyles;
				}
				if(oLogicDoc.Get_Styles) {
					return oLogicDoc.Get_Styles(level);
				}
			}
			return this.Get_Styles(0);
		}
	};

	CGraphicFrame.prototype.Get_PageContentStartPos = function (PageNum) {
		var presentation = editor.WordControl.m_oLogicDocument;
		return {
			X: 0,
			XLimit: presentation.GetWidthMM(),
			Y: 0,
			YLimit: presentation.GetHeightMM(),
			MaxTopBorder: 0
		};


	};

	CGraphicFrame.prototype.Get_PageContentStartPos2 = function () {
		return this.Get_PageContentStartPos();
	};

	CGraphicFrame.prototype.Refresh_RecalcData = function () {
		this.Refresh_RecalcData2();
	};

	CGraphicFrame.prototype.Refresh_RecalcData2 = function () {
		this.recalcInfo.recalculateTable = true;
		this.recalcInfo.recalculateSizes = true;
		this.addToRecalculate();
	};
	CGraphicFrame.prototype.checkTypeCorrect = function () {
		if (!this.graphicObject) {
			return false;
		}
		return true;
	};
	CGraphicFrame.prototype.IsThisElementCurrent = function () {
		if (this.parent && this.parent.graphicObjects) {
			if (this.group) {
				var main_group = this.group.getMainGroup();
				return main_group.selection.textSelection === this;
			} else {
				return this.parent.graphicObjects.selection.textSelection === this;
			}
		}
		return false;
	};


	CGraphicFrame.prototype.GetAllContentControls = function (arrContentControls) {
		if (this.graphicObject) {
			this.graphicObject.GetAllContentControls(arrContentControls);
		}
	};
	CGraphicFrame.prototype.IsElementStartOnNewPage = function () {
		return true;
	};

	CGraphicFrame.prototype.getCopyWithSourceFormatting = function () {
		var ret = this.copy(undefined);
		var oCopyTable = ret.graphicObject;
		var oSourceTable = this.graphicObject;
		var oTheme = this.Get_Theme();
		var oColorMap = this.Get_ColorMap();
		var RGBA;
		if (oCopyTable && oSourceTable) {
			var oPr2 = oCopyTable.Pr;
			if (oSourceTable.CompiledPr.Pr && oSourceTable.CompiledPr.Pr.TablePr) {
				oCopyTable.Set_Pr(oSourceTable.CompiledPr.Pr.TablePr.Copy());
				oPr2 = oCopyTable.Pr;
				if (oPr2.TableBorders) {
					if (oPr2.TableBorders) {
						if (oPr2.TableBorders.Bottom && oPr2.TableBorders.Bottom.Unifill) {
							oPr2.TableBorders.Bottom.Unifill.check(oTheme, oColorMap);
							RGBA = oPr2.TableBorders.Bottom.Unifill.getRGBAColor();
							oPr2.TableBorders.Bottom.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
						}
						if (oPr2.TableBorders.Left && oPr2.TableBorders.Left.Unifill) {
							oPr2.TableBorders.Left.Unifill.check(oTheme, oColorMap);
							RGBA = oPr2.TableBorders.Left.Unifill.getRGBAColor();
							oPr2.TableBorders.Left.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
						}
						if (oPr2.TableBorders.Right && oPr2.TableBorders.Right.Unifill) {
							oPr2.TableBorders.Right.Unifill.check(oTheme, oColorMap);
							RGBA = oPr2.TableBorders.Right.Unifill.getRGBAColor();
							oPr2.TableBorders.Right.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
						}
						if (oPr2.TableBorders.Top && oPr2.TableBorders.Top.Unifill) {
							oPr2.TableBorders.Top.Unifill.check(oTheme, oColorMap);
							RGBA = oPr2.TableBorders.Top.Unifill.getRGBAColor();
							oPr2.TableBorders.Top.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
						}
						if (oPr2.TableBorders.InsideH && oPr2.TableBorders.InsideH.Unifill) {
							oPr2.TableBorders.InsideH.Unifill.check(oTheme, oColorMap);
							RGBA = oPr2.TableBorders.InsideH.Unifill.getRGBAColor();
							oPr2.TableBorders.InsideH.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
						}
						if (oPr2.TableBorders.InsideV && oPr2.TableBorders.InsideV.Unifill) {
							oPr2.TableBorders.InsideV.Unifill.check(oTheme, oColorMap);
							RGBA = oPr2.TableBorders.InsideV.Unifill.getRGBAColor();
							oPr2.TableBorders.InsideV.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
						}
					}
				}
			}
			if (oCopyTable.Content.length === oSourceTable.Content.length) {
				for (var i = 0; i < oCopyTable.Content.length; ++i) {
					var aSourceCells = oSourceTable.Content[i].Content;
					var aCopyCells = oCopyTable.Content[i].Content;
					if (aSourceCells.length === aCopyCells.length) {
						for (var j = 0; j < aSourceCells.length; ++j) {
							var aSourceContent = aSourceCells[j].Content.Content;
							var aCopyContent = aCopyCells[j].Content.Content;
							AscFormat.SaveContentSourceFormatting(aSourceContent, aCopyContent, oTheme, oColorMap);
							if (aSourceCells[j].CompiledPr.Pr) {
								var oCellPr = new CTableCellPr();
								if (oPr2.TableBorders) {
									if (i === 0) {
										if (oPr2.TableBorders.Top && oPr2.TableBorders.Top.Color) {
											oCellPr.TableCellBorders.Top = oPr2.TableBorders.Top.Copy();
										}
									}
									if (i === oCopyTable.Content.length - 1) {
										if (oPr2.TableBorders.Bottom && oPr2.TableBorders.Bottom.Color) {
											oCellPr.TableCellBorders.Bottom = oPr2.TableBorders.Bottom.Copy();
										}
									}
									if (oPr2.TableBorders.InsideH && oPr2.TableBorders.InsideH.Color) {
										if (i !== 0) {
											oCellPr.TableCellBorders.Top = oPr2.TableBorders.InsideH.Copy();
										}
										if (i !== oCopyTable.Content.length - 1) {
											oCellPr.TableCellBorders.Bottom = oPr2.TableBorders.InsideH.Copy();
										}
									}
									if (j === 0) {
										if (oPr2.TableBorders.Left && oPr2.TableBorders.Left.Color) {
											oCellPr.TableCellBorders.Left = oPr2.TableBorders.Left.Copy();
										}
									}
									if (j === aSourceCells.length - 1) {
										if (oPr2.TableBorders.Right && oPr2.TableBorders.Right.Color) {
											oCellPr.TableCellBorders.Right = oPr2.TableBorders.Right.Copy();
										}
									}
									if (oPr2.TableBorders.InsideV && oPr2.TableBorders.InsideV.Color) {
										if (j !== 0) {
											oCellPr.TableCellBorders.Left = oPr2.TableBorders.InsideV.Copy();
										}
										if (j !== aSourceCells.length - 1) {
											oCellPr.TableCellBorders.Right = oPr2.TableBorders.InsideV.Copy();
										}
									}
								}
								oCellPr.Merge(aSourceCells[j].CompiledPr.Pr.Copy());
								aCopyCells[j].Set_Pr(oCellPr);
								var oPr = aCopyCells[j].Pr;
								if (oPr.Shd && oPr.Shd.Unifill) {
									if (oPr.Shd.Unifill) {
										oPr.Shd.Unifill.check(oTheme, oColorMap);
										RGBA = oPr.Shd.Unifill.getRGBAColor();
										oPr.Shd.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
									}
								}
								if (oPr.TableCellBorders) {
									if (oPr.TableCellBorders.Bottom && oPr.TableCellBorders.Bottom.Unifill) {
										oPr.TableCellBorders.Bottom.Unifill.check(oTheme, oColorMap);
										RGBA = oPr.TableCellBorders.Bottom.Unifill.getRGBAColor();
										oPr.TableCellBorders.Bottom.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
									}

									if (oPr.TableCellBorders.Left && oPr.TableCellBorders.Left.Unifill) {
										oPr.TableCellBorders.Left.Unifill.check(oTheme, oColorMap);
										RGBA = oPr.TableCellBorders.Left.Unifill.getRGBAColor();
										oPr.TableCellBorders.Left.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
									}
									if (oPr.TableCellBorders.Right && oPr.TableCellBorders.Right.Unifill) {
										oPr.TableCellBorders.Right.Unifill.check(oTheme, oColorMap);
										RGBA = oPr.TableCellBorders.Right.Unifill.getRGBAColor();
										oPr.TableCellBorders.Right.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
									}
									if (oPr.TableCellBorders.Top && oPr.TableCellBorders.Top.Unifill) {
										oPr.TableCellBorders.Top.Unifill.check(oTheme, oColorMap);
										RGBA = oPr.TableCellBorders.Top.Unifill.getRGBAColor();
										oPr.TableCellBorders.Top.Unifill = AscFormat.CreateSolidFillRGB(RGBA.R, RGBA.G, RGBA.B, 255);
									}
								}
							}
						}
					}

				}
			}
		}
		return ret;
	};
	CGraphicFrame.prototype.documentCreateFontMap = function (oMap) {
		if (this.graphicObject && this.graphicObject.Document_CreateFontMap) {
			this.graphicObject.Document_CreateFontMap(oMap);
		}
	};

	CGraphicFrame.prototype.getSpTreeDrawing = function () {
		if (this.isTable()) {
			return this;
		} else {
			let oGraphicObject = this.graphicObject;
			if (oGraphicObject) {
				if (oGraphicObject instanceof AscFormat.CT_GraphicalObject) {
					let oGraphicData = oGraphicObject.GraphicData;
					if (oGraphicData) {
						let oDrawing = oGraphicData.graphicObject;
						if (oDrawing) {
							return oDrawing;
						}
					}
				} else {
					return oGraphicObject;
				}
			}
			return null;
		}
	};

	CGraphicFrame.prototype.static_CreateGraphicFrameFromDrawing = function (oDrawing) {
		let Graphic = new AscFormat.CT_GraphicalObject();
		Graphic.Namespace = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
		Graphic.GraphicData = new AscFormat.CT_GraphicalObjectData();
		let nDrawingType = oDrawing.getObjectType();
		if (nDrawingType === AscDFH.historyitem_type_ChartSpace) {
			if (oDrawing.isChartEx()) {
				Graphic.GraphicData.Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex";
			} else {
				Graphic.GraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
			}
		} else if (nDrawingType === AscDFH.historyitem_type_SlicerView) {
			Graphic.GraphicData.Uri = "http://schemas.microsoft.com/office/drawing/2010/slicer";
		} else if (nDrawingType === AscDFH.historyitem_type_TimelineSlicerView) {
			Graphic.GraphicData.Uri = "http://schemas.microsoft.com/office/drawing/2012/timeslicer";
		} else if (nDrawingType === AscDFH.historyitem_type_SmartArt) {
			Graphic.GraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
		}
		Graphic.GraphicData.graphicObject = oDrawing;

		let newGraphicObject = AscFormat.ExecuteNoHistory(function () {
			return new AscFormat.CGraphicFrame();
		}, this, []);
		newGraphicObject.spPr = oDrawing.spPr;
		newGraphicObject.graphicObject = Graphic;
		return newGraphicObject;
	};
	CGraphicFrame.prototype.Get_ShapeStyleForPara = function () {
		return null;
	};

	CGraphicFrame.prototype.compareForMorph = function(oDrawingToCheck, oCurCandidate, oMapPaired) {
		if(!oDrawingToCheck) {
			return oCurCandidate;
		}
		const nOwnType = this.getObjectType();
		const nCheckType = oDrawingToCheck.getObjectType();
		if(nOwnType !== nCheckType) {
			return oCurCandidate;
		}
		const oTable = this.graphicObject;
		if(!oTable) {
			return oCurCandidate;
		}
		const sName = this.getOwnName();
		const nRows = oTable.GetRowsCount();
		const nCols = oTable.GetColsCount();
		const oCheckTable = oDrawingToCheck.graphicObject;
		if(sName && sName.startsWith(AscFormat.OBJECT_MORPH_MARKER)) {
			const sCheckName = oDrawingToCheck.getOwnName();
			if(sName !== sCheckName) {
				return oCurCandidate;
			}
		}
		else {
			if(oCheckTable.GetColsCount() !== nCols || oCheckTable.GetRowsCount() !== nRows) {
				return oCurCandidate;
			}
		}
		if(!oMapPaired || !oMapPaired[oDrawingToCheck.Id]) {
			if(!oCurCandidate) {
				if(oMapPaired && oMapPaired[oDrawingToCheck.Id]) {
					let oParedDrawing = oMapPaired[oDrawingToCheck.Id].drawing;
					if(oParedDrawing.getOwnName() === oDrawingToCheck.getOwnName()) {
						return oCurCandidate;
					}
				}
				return oDrawingToCheck;
			}
			const dDistCheck = this.getDistanceL1(oDrawingToCheck);
			const dDistCur = this.getDistanceL1(oCurCandidate);
			let dSizeMCandidate = Math.abs(oCurCandidate.extX - this.extX) + Math.abs(oCurCandidate.extY - this.extY);
			let dSizeMCheck = Math.abs(oDrawingToCheck.extX - this.extX) + Math.abs(oDrawingToCheck.extY - this.extY);
			if(dSizeMCandidate < dSizeMCheck) {
				return  oCurCandidate;
			}
			else {
				if(dDistCur < dDistCheck) {
					return  oCurCandidate;
				}
			}
			if(!oMapPaired || !oMapPaired[oDrawingToCheck.Id]) {
				return oDrawingToCheck;
			}
			else {
				let oParedDrawing = oMapPaired[oDrawingToCheck.Id].drawing;
				if(oParedDrawing.getOwnName() === oDrawingToCheck.getOwnName()) {
					return oCurCandidate;
				}
				else {
					return oDrawingToCheck;
				}
			}
		}
		return  oCurCandidate;
	};
	function ConvertToWordTableBorder(oBorder) {
		if (!oBorder) {
			return undefined;
		}
		var oFill = oBorder.Unifill;
		if (!oFill) {
			return oBorder;
		}
		var oNewBorder;
		var oRGBA;
		if (oFill.isSolidFillRGB()) {
			if (oFill.fill.color.color) {
				oRGBA = oFill.fill.color.color.RGBA;
			}
		}

		oNewBorder = oBorder.Copy();
		oNewBorder.Unifill = undefined;
		if (oRGBA) {
			oNewBorder.Color = new AscCommonWord.CDocumentColor(oRGBA.R, oRGBA.G, oRGBA.B, false);
		} else {
			oNewBorder.Color = new AscCommonWord.CDocumentColor(0, 0, 0, false);
		}
		return oNewBorder;
	}

	function ConvertToWordTableShd(oShd) {
		if (!oShd) {
			return undefined;
		}
		var oFill = oShd.Unifill;
		if (!oFill) {
			return oShd;
		}
		var oNewShd;
		var oCopyFill;
		if (oFill) {
			if (oFill.isSolidFillRGB()) {
				var oRGBA;
				if (oFill.fill.color.color) {
					oRGBA = oFill.fill.color.color.RGBA;
					if (oRGBA) {
						oNewShd = new AscCommonWord.CDocumentShd();

						oCopyFill = oFill.createDuplicate();
						oNewShd.Set_FromObject({
							Value: Asc.c_oAscShd.Clear,
							Color: {
								r: oRGBA.R,
								g: oRGBA.G,
								b: oRGBA.B,
								Auto: false
							},
							ThemeColor: oCopyFill,
							ThemeFill: oCopyFill
						});
						return oNewShd;
					}
				}
			}
			if (oFill.isSolidFillScheme()) {
				oNewShd = new AscCommonWord.CDocumentShd();
				oCopyFill = oFill.createDuplicate();
				oNewShd.Set_FromObject({
					Value: Asc.c_oAscShd.Clear,
					ThemeColor: oCopyFill,
					ThemeFill: oCopyFill
				});
				return oNewShd;
			}
		}
		return undefined;
	}


	function updateRowHeightAfterOpen(oRow, fRowHeight) {
		oRow.Set_Height(fRowHeight, Asc.linerule_AtLeast);
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscFormat'] = window['AscFormat'] || {};
	window['AscFormat'].CGraphicFrame = CGraphicFrame;
	window['AscFormat'].updateRowHeightAfterOpen = updateRowHeightAfterOpen;
})(window);
