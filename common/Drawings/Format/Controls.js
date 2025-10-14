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

(function () {

	const CFormControlPr_checked_unchecked = 0;
	const CFormControlPr_checked_checked = 1;
	const CFormControlPr_checked_mixed = 2;

	const CFormControlPr_selType_single = 0;
	const CFormControlPr_selType_multi = 1;
	const CFormControlPr_selType_extended = 2;

	const CFormControlPr_objectType_button = 0;
	const CFormControlPr_objectType_checkBox = 1;
	const CFormControlPr_objectType_drop = 2;
	const CFormControlPr_objectType_gBox = 3;
	const CFormControlPr_objectType_label = 4;
	const CFormControlPr_objectType_list = 5;
	const CFormControlPr_objectType_radio = 6;
	const CFormControlPr_objectType_scroll = 7;
	const CFormControlPr_objectType_spin = 8;
	const CFormControlPr_objectType_editBox = 9;
	const CFormControlPr_objectType_dialog = 10;
	const CFormControlPr_objectType_toggleButton = 11;
	const CFormControlPr_objectType_tabStrip = 12;
	const CFormControlPr_objectType_image = 13;

	const CFormControlPr_verticalAlignment_bottom = 0;
	const CFormControlPr_verticalAlignment_center = 1;
	const CFormControlPr_verticalAlignment_distributed = 2;
	const CFormControlPr_verticalAlignment_justify = 3;
	const CFormControlPr_verticalAlignment_top = 4;

	const CFormControlPr_horizontalAlignment_center = 0;
	const CFormControlPr_horizontalAlignment_continuous = 1;
	const CFormControlPr_horizontalAlignment_distributed = 2;
	const CFormControlPr_horizontalAlignment_fill = 3;
	const CFormControlPr_horizontalAlignment_general = 4;
	const CFormControlPr_horizontalAlignment_justify = 5;
	const CFormControlPr_horizontalAlignment_left = 6;
	const CFormControlPr_horizontalAlignment_right = 7;
	const CFormControlPr_horizontalAlignment_centerContinuous = 8;
	function CStringCacheManager() {
		this.paragraphsByString = {};
		this.paragraphsForDelete = null;
	}
	CStringCacheManager.prototype.getParagraphWithText = function(sText) {
		if (!this.paragraphsByString[sText]) {
			const oParagraph = AscFormat.ExecuteNoHistory(function() {
				const oShape = new AscFormat.CShape();
				oShape.createTextBody();
				const oParagraph = oShape.txBody.content.GetAllParagraphs()[0];
				oParagraph.MoveCursorToStartPos();
				oParagraph.Pr = new AscCommonWord.CParaPr();
				const oParaRun = new AscCommonWord.ParaRun(oParagraph);
				const oTextPr = getListBoxItemTextPr();
				oParaRun.Set_Pr(oTextPr);
				oParaRun.AddText(sText);
				oParagraph.AddToContent(0, oParaRun);
				oParagraph.SetParagraphAlign(AscCommon.align_Left);
				oParagraph.Reset(0, 0, 1000, 1000, 0, 0, 1);
				oParagraph.Recalculate_Page(0);
				oParagraph.LineNumbersInfo = null;
				return oParagraph;
			}, this, []);
			this.paragraphsByString[sText] = oParagraph;
		}
		if (this.paragraphsForDelete) {
			delete this.paragraphsForDelete[sText];
		}
		return this.paragraphsByString[sText];
	};
	CStringCacheManager.prototype.startCheckDeleteParagraphs = function() {
		this.paragraphsForDelete = Object.assign({}, this.paragraphsByString);
	};
	CStringCacheManager.prototype.endCheckDeleteParagraphs = function() {
		for (let sText in this.paragraphsForDelete) {
			delete this.paragraphsByString[sText];
		}
		this.paragraphsForDelete = null;
	};
	function drawParagraph(paragraph, graphics) {
		const oApi = Asc.editor;
		let bOldViewMode = false;
		if (oApi) {
			bOldViewMode = oApi.isViewMode;
			oApi.isViewMode = true;
		}
		paragraph.Draw(0, graphics);
		if (oApi) {
			oApi.isViewMode = bOldViewMode;
		}
	}
	function getVerticalAlignFromControlPr(nPr) {
		switch (nPr) {
			case CFormControlPr_verticalAlignment_bottom:
				return AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM;
			case CFormControlPr_verticalAlignment_center:
				return AscFormat.VERTICAL_ANCHOR_TYPE_CENTER;
			case CFormControlPr_verticalAlignment_distributed:
				return AscFormat.VERTICAL_ANCHOR_TYPE_DISTRIBUTED;
			case CFormControlPr_verticalAlignment_justify:
				return AscFormat.VERTICAL_ANCHOR_TYPE_JUSTIFIED;
			case CFormControlPr_verticalAlignment_top:
				return AscFormat.VERTICAL_ANCHOR_TYPE_TOP;
			default:
				return null;
		}
	}

	function getHorizontalAlignFromControl(nPr) {
		switch (nPr) {
			case CFormControlPr_horizontalAlignment_center:
				return AscCommon.align_Center;
			case CFormControlPr_horizontalAlignment_continuous:
				return null;
			case CFormControlPr_horizontalAlignment_distributed:
				return AscCommon.align_Distributed;
			case CFormControlPr_horizontalAlignment_fill:
				return null;
			case CFormControlPr_horizontalAlignment_general:
				return null;
			case CFormControlPr_horizontalAlignment_justify:
				return AscCommon.align_Justify;
			case CFormControlPr_horizontalAlignment_left:
				return AscCommon.align_Left;
			case CFormControlPr_horizontalAlignment_right:
				return AscCommon.align_Right;
			case CFormControlPr_horizontalAlignment_centerContinuous:
				return AscCommon.align_CenterContinuous;
			default:
				return null;
		}
	}

	function getVerticalControlAlignFromBodyPr(nPr) {
		switch (nPr) {
			case AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM:
				return CFormControlPr_verticalAlignment_bottom;
			case AscFormat.VERTICAL_ANCHOR_TYPE_CENTER:
				return CFormControlPr_verticalAlignment_center;
			case AscFormat.VERTICAL_ANCHOR_TYPE_DISTRIBUTED:
				return CFormControlPr_verticalAlignment_distributed;
			case AscFormat.VERTICAL_ANCHOR_TYPE_JUSTIFIED:
				return CFormControlPr_verticalAlignment_justify;
			case AscFormat.VERTICAL_ANCHOR_TYPE_TOP:
				return CFormControlPr_verticalAlignment_top;
			default:
				return null;
		}
	}

	function getHorizontalAlignFromContentControl(nPr) {
		switch (nPr) {
			case AscCommon.align_Center:
				return CFormControlPr_horizontalAlignment_center;
			case AscCommon.align_Distributed:
				return CFormControlPr_horizontalAlignment_distributed;
			case AscCommon.align_Justify:
				return CFormControlPr_horizontalAlignment_justify;
			case AscCommon.align_Left:
				return CFormControlPr_horizontalAlignment_left;
			case AscCommon.align_Right:
				return CFormControlPr_horizontalAlignment_right;
			case AscCommon.align_CenterContinuous:
				return CFormControlPr_horizontalAlignment_centerContinuous;
			default:
				return null;
		}
	}
	function startRoundControl(graphics, x, y, extX, extY, nRadiusPx, arrColor) {
		const nRadius = Math.min(nRadiusPx * AscCommon.g_dKoef_pix_to_mm * AscCommon.AscBrowser.retinaPixelRatio, extX / 2, extY / 2);
		function draw() {
			graphics._s();
			graphics._m(x, y + nRadius);
			graphics._c2(x, y, x + nRadius, y);
			graphics._l(x + extX - nRadius, y);
			graphics._c2(x + extX, y, x + extX, y + nRadius);
			graphics._l(x + extX, y + extY - nRadius);
			graphics._c2(x + extX, y + extY, x + extX - nRadius, y + extY);
			graphics._l(x + nRadius, y + extY);
			graphics._c2(x, y + extY, x, y + extY - nRadius);
			graphics._z();
		}
		graphics.SaveGrState();
		graphics.p_color.apply(graphics, arrColor);
		graphics.p_width(0);
		draw();
		graphics.ds();
		graphics.AddClipRect(x, y, extX, extY);
		graphics.StartClipPath();
		draw();
		graphics.EndClipPath();
	}
	function endRoundControl(graphics) {
		graphics.RestoreGrState();
	}
	function CStepManager() {
		this.timeoutId = null;
	}
	CStepManager.prototype.start = function (fCallback) {
		const oThis = this;
		if (oThis.timeoutId === null) {
			fCallback();
			oThis.timeoutId = setTimeout(function () {
				(function f() {
					fCallback();
					oThis.timeoutId = setTimeout(f, 100);
				})();
			}, 500);
		}
	};
	CStepManager.prototype.end = function () {
		clearTimeout(this.timeoutId);
		this.timeoutId = null;
	};
function getFlatPenColor() {
	return [202, 202, 202, 255];
}
	function getFlatCheckBoxPenColor() {
		return [51, 51, 51, 255];
	}
	function CButtonBase(oController) {
		this.x = null;
		this.y = null;
		this.extX = null;
		this.extY = null;
		this.transform = new AscCommon.CMatrix();
		this.invertTransform = new AscCommon.CMatrix();
		this.controller = oController;
		this.isHold = false;
		this.isHover = false;
	}
	CButtonBase.prototype.getFlatFillColor = function () {
		if (this.isHold) {
			return [225, 225, 225, 255];
		}
		if (this.isHover) {
			return [234, 234, 234, 255];
		}
		return [244, 244, 244, 255];
	};
	CButtonBase.prototype.getFlatPenColor = function () {
		return getFlatPenColor();
	};
	CButtonBase.prototype.draw = function (graphics) {
		graphics.SaveGrState();
		graphics.transform3(this.transform);
		const arrPenColor = this.getFlatPenColor();
		graphics.p_width(0);
		graphics.p_color.apply(graphics, arrPenColor);
		graphics.b_color1.apply(graphics, this.getFlatFillColor());
		graphics._s();
		graphics._m(0, 0);
		graphics._l(this.extX, 0);
		graphics._l(this.extX, this.extY);
		graphics._l(0, this.extY);
		graphics._z();
		graphics.ds();
		graphics.df();
		graphics.RestoreGrState();
		graphics._e();
	};
	CButtonBase.prototype.hit = function (nX, nY) {
		return AscFormat.HitToRect(nX, nY, this.invertTransform, 0, 0, this.extX, this.extY);
	};
	CButtonBase.prototype.setIsHold = function (pr) {
		if (this.isHold !== pr) {
			this.controller.checkNeedUpdate();
		}
		this.isHold = pr;
	};
	CButtonBase.prototype.setIsHover = function (pr) {
		if (this.isHover !== pr) {
			this.controller.checkNeedUpdate();
		}
		this.isHover = pr;
	};
	CButtonBase.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (!this.hit(nX, nY)) {
			return false;
		}

		this.setIsHold(true);
		if (this._onMouseDown(e, nX, nY, nPageIndex, oDrawingController)) {
			this.controller.checkNeedUpdate();
		}
		return true;
	};
	CButtonBase.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (this.isHold) {
			this.setIsHold(false);
			if (this._onMouseUp(e, nX, nY, nPageIndex, oDrawingController)) {
				this.controller.checkNeedUpdate();
			}
			return true;
		}
		return false;
	};
	CButtonBase.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		this.setIsHover(this.hit(nX, nY));
		if (this.isHover || this.isHold) {
			return this._onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		}
		return false;
	};
	CButtonBase.prototype._onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {

	};
	CButtonBase.prototype._onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {

	};
	CButtonBase.prototype._onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {

	};

	AscDFH.changesFactory[AscDFH.historyitem_Control_ControlPr] = AscDFH.CChangesDrawingsObject;
	AscDFH.changesFactory[AscDFH.historyitem_Control_FormControlPr] = AscDFH.CChangesDrawingsObject;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_Control_ControlPr] = function (oClass, pr) {
		oClass.controlPr = pr;
	}
	AscDFH.drawingsChangesMap[AscDFH.historyitem_Control_FormControlPr] = function (oClass, pr) {
		oClass.formControlPr = pr;
	}

	function CControl() {
		AscFormat.CShape.call(this);
		this.name = null;
		this.link = null;
		this.rId = null;
		this.setFormControlPr(new CFormControlPr());
		this.setControlPr(new CControlPr());
		this.controller = null;
	}

	AscFormat.InitClass(CControl, AscFormat.CShape, AscDFH.historyitem_type_Control);
	CControl.prototype.superclass = AscFormat.CGraphicObjectBase;
	CControl.prototype.fillObject = function (oCopy, oPr) {
		AscFormat.CShape.prototype.fillObject.call(this, oCopy, oPr);
		if (this.controlPr) {
			oCopy.setControlPr(this.controlPr.createDuplicate());
		}
		if (this.formControlPr) {
			oCopy.setFormControlPr(this.formControlPr.createDuplicate());
		}
	};
	CControl.prototype.isNeedResetState = function() {
		return this.controller.isNeedResetState();
	};
	CControl.prototype.initController = function () {
		switch (this.formControlPr.objectType) {
			case CFormControlPr_objectType_checkBox: {
				this.controller = new CCheckBoxController(this);
				break;
			}
			case CFormControlPr_objectType_button: {
				this.controller = new CButtonController(this);
				break;
			}
			case CFormControlPr_objectType_spin: {
				this.controller = new CSpinController(this);
				break;
			}
			case CFormControlPr_objectType_scroll: {
				this.controller = new CScrollController(this);
				break;
			}
			case CFormControlPr_objectType_list: {
				this.controller = new CListBoxController(this);
				break;
			}
			case CFormControlPr_objectType_drop: {
				this.controller = new CComboBoxController(this);
				break;
			}
			default: {
				return false;
			}
		}
		return true;
	}
	CControl.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		const oUR = graphics.updatedRect;
		if (oUR && this.bounds) {
			if (!oUR.isIntersectOther(this.bounds)) {
				return;
			}
		}
		this.controller.draw(graphics, transform, transformText, pageIndex, opt);
	};
	CControl.prototype.hitInInnerArea = function (x, y) {
		return this.controller.hitInInnerArea(x, y);
	}
	CControl.prototype.hitInPath = CControl.prototype.hitInInnerArea;
	CControl.prototype.hitInTextRect = function (x, y) {
		if (this.selected) {
			return AscFormat.CShape.prototype.hitInTextRect.call(this, x, y);
		}
		return false;
	};
	CControl.prototype.isControl = function () {
		return true;
	}
	CControl.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		const bRet = this.controller.onMouseDown(e, nX, nY, nPageIndex, oDrawingController);
		this.controller.update();
		return bRet;
	}
	CControl.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		const bRet = this.controller.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		this.controller.update();
		return bRet;
	}
	CControl.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		const bRet = this.controller.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
		this.controller.update();
		return bRet;
	}
	CControl.prototype.getCursorInfo = function (e, nX, nY) {
		return this.controller.getCursorInfo(e, nX, nY);
	}
	CControl.prototype.getTextRect = function () {
		return this.controller.getTextRect();
	};
	CControl.prototype.canRotate = function () {
		return false;
	};
	CControl.prototype.canEditText = function () {
		return this.controller.canEditText();
	};
	CControl.prototype.canEditGeometry = function () {
		return false;
	};
	CControl.prototype.setControlPr = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_Control_ControlPr, this.controlPr, pr));
		this.controlPr = pr;
		this.controlPr.setParent(this);
	};
	CControl.prototype.setFormControlPr = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_Control_FormControlPr, this.formControlPr, pr));
		this.formControlPr = pr;
		this.formControlPr.setParent(this);
	};
	CControl.prototype.clearVmlTxBody = function () {
		const oDocContent = this.getDocContent();
		for (let i = 0; i < oDocContent.Content.length; i++) {
			const oParagraph = oDocContent.Content[i];
			oParagraph.CheckRunContent(function (oRun) {
				let nCount = 0;
				for (let i = oRun.Content.length - 1; i >= 0; i -= 1) {
					const oElement = oRun.Content[i];
					switch (oElement.Type) {
						case para_Space: {
							nCount += 1;
							break;
						}
						case para_NewLine: {
							oRun.Content.splice(i, nCount);
							nCount = 0;
							break;
						}
						default: {
							nCount = 0;
							break;
						}
					}
				}
			});
		}
	}
	CControl.prototype.getControlPr = function () {
		return this.controlPr;
	};
	CControl.prototype.getFormControlPr = function () {
		return this.formControlPr;
	};
	CControl.prototype.getChecked = function () {
		return this.controller.getChecked();
	}
	CControl.prototype.copy = function (oPr) {
		var copy = new CControl();
		this.fillObject(copy, oPr);
		copy.initController();
		return copy;
	};
	CControl.prototype.applySpecialPasteProps = function (oPastedWb) {
		this.controller.applySpecialPasteProps(oPastedWb);
	};
	CControl.prototype.initTextProperties = function () {
		this.controller.initTextProperties();
	};
	CControl.prototype.setMacro = function (pr) {
		this.controlPr.setMacro(pr);
	};
	CControl.prototype.getTextHAlign = function () {
		const oTxBody = this.txBody;
		const oContent = oTxBody && oTxBody.content;
		if (oContent) {
			oContent.SetApplyToAll(true);
			const oTextPr = oContent.GetDirectParaPr();
			oContent.SetApplyToAll(false);
			return getHorizontalAlignFromContentControl(oTextPr.Jc);
		}
		return null;
	};
	CControl.prototype.getTextVAlign = function () {
		const oTxBody = this.txBody;
		const oBodyPr = oTxBody && oTxBody.bodyPr;
		if (oBodyPr) {
			return getVerticalControlAlignFromBodyPr(oBodyPr.anchor);
		}
		return null;
	};
	CControl.prototype.recalculateTransform = function () {
		AscFormat.CShape.prototype.recalculateTransform.call(this);
		this.controller.recalculateTransform();
	};
	CControl.prototype.recalculate = function () {
		AscFormat.CShape.prototype.recalculate.call(this);
		this.controller.recalculate();
	};
	CControl.prototype.handleChangeRanges = function (arrRanges) {
		return this.controller.handleChangeRanges(arrRanges);
	}
	CControl.prototype.onUpdate = function () {
		this.controller.onUpdate();
	};

	function CControlControllerBase(oControl) {
		this.control = oControl;
		this.isNeedUpdate = false;
	}
	CControlControllerBase.prototype.getFormControlPr = function () {
		return this.control.formControlPr;
	};
	CControlControllerBase.prototype.getControlPr = function () {
		return this.control.controlPr;
	};
	CControlControllerBase.prototype.getWorksheet = function () {
		return this.control.Get_Worksheet();
	};
	CControlControllerBase.prototype.getParsedFmlaLink = function () {
		const oFormControlPr = this.getFormControlPr();
		const oRef = this.getParsedRef(oFormControlPr.fmlaLink);
		if (oRef) {
			return new AscCommonExcel.Range(oRef.worksheet, oRef.bbox.r1, oRef.bbox.c1, oRef.bbox.r1, oRef.bbox.c1);
		}
		return null;
	};
	CControlControllerBase.prototype.getParsedFmlaRange = function () {
		const oFormControlPr = this.getFormControlPr();
		return this.getParsedRef(oFormControlPr.fmlaRange);
	};
	CControlControllerBase.prototype.getParsedRef = function (sStrRef) {
		if (sStrRef) {
			const oWs = this.getWorksheet();
			let aParsedRef = AscCommonExcel.getRangeByRef(sStrRef, oWs, true, true, true);
			const oRef = aParsedRef[0];
			if (oRef) {
				return oRef;
			}
		}
		return null;
	};
	CControlControllerBase.prototype.setRangeValue = function (oRef, oCellValue) {
		oRef._foreach(function (oCell) {
			if (oCell) {
				oCell.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, new AscCommonExcel.CCellValue(oCellValue)));

			}
		});
		const oWb = Asc.editor.wb;
		if (oWb) {
			const oActiveWs = oWb.getWorksheet(null, true);
			const nWorksheetIndex = oRef.worksheet.getIndex();
			const oWs = oWb.getWorksheet(nWorksheetIndex, true);
			if (oWs) {
				oWs._updateRange(oRef.bbox);
			}
			if (oActiveWs) {
				oActiveWs.draw();
			}
		}
	};
	CControlControllerBase.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {};
	CControlControllerBase.prototype.getCursorInfo = function (e, nX, nY) {};
	CControlControllerBase.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {};
	CControlControllerBase.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {return false;};
	CControlControllerBase.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {};
	CControlControllerBase.prototype.initTextProperties = function () {return null;};
	CControlControllerBase.prototype.applySpecialPasteProps = function (oPastedWb) {};
	CControlControllerBase.prototype.handleChangeRanges = function (aRanges) {
		const bHandleFmlaRange = this.handleFmlaRange(aRanges);
		const bHandleFmlaLink = this.handleFmlaLink(aRanges);
		return bHandleFmlaRange || bHandleFmlaLink;
	};
	CControlControllerBase.prototype.getTextRect = function () {
		return AscFormat.CShape.prototype.getTextRect.call(this.control);
	};
	CControlControllerBase.prototype.getChecked = function () {
		return null;
	};
	CControlControllerBase.prototype.recalculateTransform = function () {};
	CControlControllerBase.prototype.recalculate = function() {};
	CControlControllerBase.prototype.handleFmlaLink = function (aRanges) {return false;};
	CControlControllerBase.prototype.handleFmlaRange = function (aRanges) {return false;};
	CControlControllerBase.prototype.onUpdate = function () {
		return AscFormat.CShape.prototype.onUpdate.call(this.control);
	};
	CControlControllerBase.prototype.hitInInnerArea = function (nX, nY) {
		const oControl = this.control;
		return AscFormat.HitToRect(nX, nY, oControl.invertTransform, 0, 0, oControl.extX, oControl.extY);
	};
	CControlControllerBase.prototype.handleRef = function(aRanges, oRef, fCallback) {
		if (!oRef) {
			return false;
		}
		for (let i = 0; i < aRanges.length; i += 1) {
			if (oRef.isIntersect(aRanges[i])) {
				fCallback();
				return true;
			}
		}
		return false;
	};
	CControlControllerBase.prototype.isNeedResetState = function () {
		return true;
	};
	CControlControllerBase.prototype.checkNeedUpdate = function() {
		this.isNeedUpdate = true;
	};
	CControlControllerBase.prototype.update = function() {
		if (this.isNeedUpdate) {
			this.isNeedUpdate = false;
			this.control.onUpdate();
			this.forceUpdate();
		}
	};
	CControlControllerBase.prototype.forceUpdate = function() {
		this.control.onUpdate();
	};
	CControlControllerBase.prototype.canEditText = function () {
		return false;
	};

	const CHECKBOX_SIDE_SIZE = 3;
	const CHECKBOX_X_OFFSET = 1.5;
	const CHECKBOX_BODYPR_INSETS_L = 27432 / 36000;
	const CHECKBOX_BODYPR_INSETS_R = 0;
	const CHECKBOX_BODYPR_INSETS_T = 32004 / 36000;
	const CHECKBOX_BODYPR_INSETS_B = 32004 / 36000;
	const CHECKBOX_OFFSET_X = CHECKBOX_SIDE_SIZE + (CHECKBOX_X_OFFSET * 2 - CHECKBOX_BODYPR_INSETS_L);

	function CCheckBox(oController) {
		CButtonBase.call(this, oController);
	}
	AscFormat.InitClassWithoutType(CCheckBox, CButtonBase);
	CCheckBox.prototype.getFlatPenColor = function () {
		return getFlatCheckBoxPenColor();
	};
	CCheckBox.prototype.isChecked = function () {

	};
	CCheckBox.prototype.isMixed = function () {

	};
	CCheckBox.prototype.getFlatFillColor = function () {
		if (this.isHold) {
			return [225, 225, 225, 255];
		}
		if (this.isHover) {
			return [234, 234, 234, 255];
		}
		return [255, 255, 255, 255];
	};
	CCheckBox.prototype.draw = function (graphics) {
		graphics.SaveGrState();
		graphics.transform3(this.transform);
		startRoundControl(graphics, 0, 0, this.extX, this.extY, 2, getFlatCheckBoxPenColor());
		CButtonBase.prototype.draw.call(this, graphics);
		graphics.p_color.apply(graphics, this.getFlatPenColor());
		graphics.p_width(400);
		graphics._s();
		if (this.isChecked()) {
			graphics._m(2.5, 0.75);
			graphics._l(1, 2.25);
			graphics._l(0.5, 1.75);
			graphics.ds();

		} else if (this.isMixed()) {
			graphics._m(CHECKBOX_SIDE_SIZE * 0.2, CHECKBOX_SIDE_SIZE * 0.5);
			graphics._l(CHECKBOX_SIDE_SIZE * 0.8, CHECKBOX_SIDE_SIZE * 0.5);
			graphics.ds();
		}
		graphics._e();
		endRoundControl(graphics);
		graphics.RestoreGrState();
	};

	function CCheckBoxController(oControl) {
		CControlControllerBase.call(this, oControl);
		this.checkBox = new CCheckBox(this);
		this.initCheckBoxHandlers();
	}
	AscFormat.InitClassWithoutType(CCheckBoxController, CControlControllerBase);
	CCheckBoxController.prototype.initCheckBoxHandlers = function () {
		const oThis = this;
		this.checkBox.isChecked = function () {
			return oThis.isChecked();
		};
		this.checkBox.isMixed = function () {
			return oThis.isMixed();
		};
	};
	CCheckBoxController.prototype.initTextProperties = function () {
		const oControl = this.control;
		const oTxBody = oControl.txBody;
		if (oTxBody) {
			const oBodyPr = new AscFormat.CBodyPr();
			oBodyPr.setInsets(CHECKBOX_BODYPR_INSETS_L, CHECKBOX_BODYPR_INSETS_T, CHECKBOX_BODYPR_INSETS_R, CHECKBOX_BODYPR_INSETS_B);
			oBodyPr.setAnchor(getVerticalAlignFromControlPr(this.control.formControlPr.textVAlign));
			oBodyPr.vertOverflow = AscFormat.nVOTClip;
			oBodyPr.wrap = AscFormat.nTWTSquare;
			oBodyPr.upright = true;
			oTxBody.setBodyPr(oBodyPr);
		}
	};
	CCheckBoxController.prototype.recalculateTransform = function () {
		const oControl = this.control;
		const oCheckBoxTransform = oControl.transform.CreateDublicate();
		const nXOffset = CHECKBOX_X_OFFSET;
		const nYOffset = (oControl.extY - CHECKBOX_SIDE_SIZE) / 2;
		oCheckBoxTransform.tx += nXOffset;
		oCheckBoxTransform.ty += nYOffset;
		this.checkBox.x = oControl.x + nXOffset;
		this.checkBox.y = oControl.y + nYOffset;
		this.checkBox.extX = CHECKBOX_SIDE_SIZE;
		this.checkBox.extY = CHECKBOX_SIDE_SIZE;
		this.checkBox.transform = oCheckBoxTransform;
		this.checkBox.invertTransform = oCheckBoxTransform.CreateDublicate().Invert();
	};
	CCheckBoxController.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		const oControl = this.control;
		AscFormat.CShape.prototype.draw.call(oControl, graphics, transform, transformText, pageIndex, opt);
		this.checkBox.draw(graphics);
	};
	CCheckBoxController.prototype.isChecked = function () {
		const nCheckValue = this.getChecked();
		return nCheckValue === CFormControlPr_checked_checked;
	};
	CCheckBoxController.prototype.isMixed = function () {
		const nCheckValue = this.getChecked();
		return nCheckValue === CFormControlPr_checked_mixed;
	};
	CCheckBoxController.prototype.isEmpty = function () {
		return !(this.isChecked() || this.isMixed());
	};
	CCheckBoxController.prototype.isExternalCheckBox = function () {
		const oRef = this.getParsedFmlaLink();
		const oWbModel = Asc.editor && Asc.editor.wbModel;
		if (oRef && oWbModel) {
			return oRef.worksheet.workbook !== oWbModel;
		}
		return false;
	};
	CCheckBoxController.prototype.getCursorInfo = function (e, nX, nY) {
		const oControl = this.control;
		if (oControl.selected) {
			return null;
		}
		if (!oControl.hit(nX, nY)) {
			return null;
		}
		return {cursorType: "pointer", objectId: oControl.GetId()};
	};
	CCheckBoxController.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		const oControl = this.control;
		if (oControl.selected) {
			return false;
		}
		if (e.button !== 0) {
			return false;
		}
		if (e.CtrlKey) {
			return false;
		}
		if (!oControl.hit(nX, nY)) {
			return false;
		}
		this.checkBox.setIsHold(true);
		return true;
	}
	CCheckBoxController.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		const bIsHit = this.control.hit(nX, nY);
		this.checkBox.setIsHover(bIsHit);
		return bIsHit;
	}

	CCheckBoxController.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oController) {
		this.checkBox.setIsHold(false);
		if (this.isExternalCheckBox()) {
			this.checkNeedUpdate();
			return false;
		}
		const oThis = this;
		oController.checkObjectsAndCallback(function () {
			const oFormControlPr = oThis.getFormControlPr();
			if (!oThis.isEmpty()) {
				oFormControlPr.setChecked(CFormControlPr_checked_unchecked);
			} else {
				oFormControlPr.setChecked(CFormControlPr_checked_checked);
			}
			oThis.updateCellFromControl();
		}, [], false, AscDFH.historydescription_Spreadsheet_SwitchCheckbox, [this.control]);
		this.checkNeedUpdate();
		return true;
	}
	CCheckBoxController.prototype.getCheckedFromRange = function () {
		const oRef = this.getParsedFmlaLink();
		let nRetValue = null;
		if (oRef) {
			oRef._foreachNoEmpty(function (oCell) {
				if (oCell) {
					const bValue = oCell.getBoolValue();
					if (oCell.type === AscCommon.CellValueType.Bool || oCell.type === AscCommon.CellValueType.Number) {
						nRetValue = bValue ? CFormControlPr_checked_checked : CFormControlPr_checked_unchecked;
					} else if (oCell.type === AscCommon.CellValueType.Error) {
						nRetValue = CFormControlPr_checked_mixed;
					}
				}
			});
		}
		return nRetValue;
	};
	CCheckBoxController.prototype.getCellValueFromControl = function () {
		const oFormControlPr = this.getFormControlPr();
		const oCellValue = new AscCommonExcel.CCellValue();
		if (oFormControlPr.checked === CFormControlPr_checked_checked) {
			oCellValue.type = AscCommon.CellValueType.Bool;
			oCellValue.number = 1;
		} else if (oFormControlPr.checked === CFormControlPr_checked_mixed) {
			oCellValue.type = AscCommon.CellValueType.Error;
			oCellValue.text = AscCommonExcel.cError.prototype.getStringFromErrorType(AscCommonExcel.cErrorType.not_available);
		} else {
			oCellValue.type = AscCommon.CellValueType.Bool;
			oCellValue.number = 0;
		}
		return oCellValue;
	}
	CCheckBoxController.prototype.updateCellFromControl = function () {
		const oThis = this;
		const oRef = this.getParsedFmlaLink();
		if (oRef) {
			const oCellValue = oThis.getCellValueFromControl();
			this.setRangeValue(oRef, oCellValue);
		}
	};
	CCheckBoxController.prototype.getTextRect = function () {
		const oTextRect = AscFormat.CShape.prototype.getTextRect.call(this.control);
		oTextRect.l += CHECKBOX_OFFSET_X;
		oTextRect.r += CHECKBOX_OFFSET_X;
		return oTextRect;
	};
	CCheckBoxController.prototype.getChecked = function () {
		const nRangeValue = this.getCheckedFromRange();
		if (nRangeValue !== null) {
			return nRangeValue;
		}
		const oFormControlPr = this.getFormControlPr();
		return oFormControlPr.getChecked();
	};
	CCheckBoxController.prototype.applySpecialPasteProps = function (oPastedWb) {
		this.addExternalReferenceToEditor(oPastedWb);
	};
	CCheckBoxController.prototype.addExternalReferenceToEditor = function (oPastedWb) {
		const oApi = Asc.editor;
		const oWbModel = oApi && oApi.wbModel;
		if (!oWbModel) {
			return;
		}
		const oFormControlPr = this.getFormControlPr();
		const sRef = oFormControlPr.fmlaLink;
		if (!sRef) {
			return;
		}
		const oMockWb = new AscCommonExcel.Workbook(undefined, undefined, false);
		oMockWb.externalReferences = oPastedWb.externalReferences;
		oMockWb.dependencyFormulas = oPastedWb.dependencyFormulas;
		const oWorksheets = this.getWorksheetsFromControlValue(oMockWb);
		const oMainExternalReference = oWbModel.addExternalReferenceFromWorksheets(oWorksheets, oPastedWb, oMockWb);
		if (oMainExternalReference) {
			const sNewRef = AscFormat.updateRefToExternal(sRef, oMainExternalReference, oPastedWb.externalReferences, oMockWb);
			oFormControlPr.setFmlaLink(sNewRef);
		} else if (oPastedWb.externalReferences.length) {
			const sNewRef = AscFormat.updateRefToExternal(sRef, oMainExternalReference, oPastedWb.externalReferences, oPastedWb);
			oFormControlPr.setFmlaLink(sNewRef);
		}
	};
	CCheckBoxController.prototype.getWorksheetsFromControlValue = function (oParentWb) {

		const oFormControlPr = this.getFormControlPr();
		const sRef = oFormControlPr.fmlaLink;
		const oRes = {};
		if (sRef) {
			const arrF = AscFormat.getParsedCopyRefs(sRef, oParentWb);
			if (arrF.length) {
				const oFirstRef = arrF[0];
				const sSheetName = oFirstRef.sheet;
				const oWorksheet = new AscCommonExcel.Worksheet(oParentWb);
				oWorksheet.sName = sSheetName;
				const oRange = oWorksheet.getRange2(oFirstRef.range);
				if (oRange) {
					const oBBox = oRange.bbox;
					const oWorksheetInfo = {};
					oRes[sSheetName] = oWorksheetInfo;
					oRes[sSheetName].defNames = [];
					oRes[sSheetName].ws = oWorksheet;
					oRes[sSheetName].maxR = oBBox.r1;
					oRes[sSheetName].maxC = oBBox.c1;
					oRes[sSheetName].minC = oBBox.c1;
					oRes[sSheetName].minR = oBBox.r1;

					const oCellValue = this.getCellValueFromControl();
					oWorksheet._getCell(oBBox.c1, oBBox.r1, function (oCell) {
						oCell.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));
					});
					if (oFirstRef.defName) {
						oWorksheetInfo.defNames.push(oFirstRef.defName);
					}
				}
			}
		}
		return oRes;
	};
	CCheckBoxController.prototype.canEditText = function () {
		return AscFormat.CShape.prototype.canEditText.call(this.control);
	};

	const BUTTON_BODYPR_INSETS = 27432 / 36000;

	function CButtonController(oControl) {
		CControlControllerBase.call(this, oControl);
		this.button = new CButtonBase(this);
		this.initButton();
	};
	AscFormat.InitClassWithoutType(CButtonController, CControlControllerBase);
	CButtonController.prototype.initButton = function () {
		const oThis = this;
		this.button._onMouseMove = function () {
			return true;
		};
		this.button._onMouseDown = function () {
			return true;
		};
		this.button._onMouseUp = function () {
			oThis.runMacros();
			return true;
		}
	};
	CButtonController.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		const oControl = this.control;
		transform = transform || oControl.transform;
		transformText = transformText || oControl.transformText;
		graphics.SaveGrState();
		graphics.transform3(transform);
		startRoundControl(graphics, 0, 0, oControl.extX, oControl.extY, 4, getFlatPenColor());
		this.button.draw(graphics);
		oControl.drawTxBody(graphics, transform, transformText, pageIndex);
		endRoundControl(graphics);
		graphics.RestoreGrState();
	};
	CButtonController.prototype.recalculateTransform = function () {
		const oControl = this.control;
		this.button.x = oControl.x;
		this.button.y = oControl.y;
		this.button.extX = oControl.extX;
		this.button.extY = oControl.extY;
		this.button.transform = oControl.transform.CreateDublicate();
		this.button.invertTransform = oControl.invertTransform.CreateDublicate();
	};
	CButtonController.prototype.onMacroError = function (sMacro) {
		const oApi = Asc.editor;
		if (oApi) {
			oApi.sendEvent("asc_onError", Asc.c_oAscError.ID.MacroUnavailableWarning, c_oAscError.Level.NoCritical, sMacro);
		}
	};
	CButtonController.prototype.getCursorInfo = function (e, nX, nY) {
		const oControl = this.control;
		if (oControl.selected) {
			return null;
		}
		if (!oControl.hit(nX, nY)) {
			return null;
		}
		return {cursorType: "pointer", objectId: oControl.GetId()};
	};
	CButtonController.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		const oControl = this.control;
		if (oControl.selected) {
			return false;
		}
		if (e.button !== 0) {
			return false;
		}
		if (e.CtrlKey) {
			return false;
		}
		const oControlPr = this.getControlPr();
		if (oControlPr.getMacroId() === null) {
			return false;
		}
		return this.button.onMouseDown(e, nX, nY, nPageIndex, oDrawingController);
	};
	CButtonController.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oController) {
		return this.button.onMouseUp(e, nX, nY, nPageIndex, oController);
	};
	CButtonController.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oController) {
		return this.button.onMouseMove(e, nX, nY, nPageIndex, oController);
	};
	CButtonController.prototype.runMacros = function () {
		const oControlPr = this.getControlPr();
		const sMacro = oControlPr.getJSAMacroId();
		if (sMacro === null) {
			this.onMacroError(oControlPr.getMacroId());
			return;
		}
		const oApi = Asc.editor;
		if (oApi) {
			const sMacroName = oApi.macros && oApi.macros.getNameByGuid(sMacro);
			if (!sMacroName) {
				this.onMacroError(sMacro);
			}
			oApi.asc_runMacros(sMacro);
		}
	};
	CButtonController.prototype.initTextProperties = function () {
		const oControl = this.control;
		const oBodyPr = new AscFormat.CBodyPr();
		oBodyPr.setInsets(BUTTON_BODYPR_INSETS, BUTTON_BODYPR_INSETS, BUTTON_BODYPR_INSETS, BUTTON_BODYPR_INSETS);
		oBodyPr.setAnchor(getVerticalAlignFromControlPr(oControl.formControlPr.textVAlign));
		oBodyPr.vertOverflow = AscFormat.nVOTClip;
		oBodyPr.wrap = AscFormat.nTWTSquare;
		oBodyPr.upright = true;
		const oTextBody = oControl.txBody;
		oTextBody.setBodyPr(oBodyPr);
		const oContent = oTextBody.content;
		if (oContent) {
			oContent.SetApplyToAll(true);
			oContent.SetParagraphAlign(getHorizontalAlignFromControl(oControl.formControlPr.textHAlign));
			oContent.SetApplyToAll(false);
		}
		return oBodyPr;
	};
	CButtonController.prototype.canEditText = function () {
		return AscFormat.CShape.prototype.canEditText.call(this.control);
	};

	const SPINBUTTON_RECTANGLE_SIDE_SCALE = 0.25;
	const SPINBUTTON_DIRECTION_UP = 1;
	const SPINBUTTON_DIRECTION_DOWN = 2;
	const SPINBUTTON_DIRECTION_RIGHT = 3;
	const SPINBUTTON_DIRECTION_LEFT = 4;
	function CSpinButton(oController, nDirection) {
		CButtonBase.call(this, oController);
		this.direction = nDirection || SPINBUTTON_DIRECTION_UP;
	}
	AscFormat.InitClassWithoutType(CSpinButton, CButtonBase);
	CSpinButton.prototype.draw = function (graphics) {
		CButtonBase.prototype.draw.call(this, graphics);
		const arrRectColor = [51, 51, 51, 255];
		graphics.SaveGrState();
		graphics.transform3(this.transform);
		const nRectCX = this.extX / 2;
		const nRectCY = this.extY / 2;
		const nRectHeight = (SPINBUTTON_RECTANGLE_SIDE_SCALE) * Math.min(this.extX, this.extY);
		const nHalfRectHeight = nRectHeight / 2;
		graphics.p_width(nRectHeight * 200);
		graphics.p_color.apply(graphics, arrRectColor);
		graphics._s();

		switch (this.direction) {
			case SPINBUTTON_DIRECTION_UP: {
				graphics._m(nRectCX - nRectHeight, nRectCY + nHalfRectHeight);
				graphics._l(nRectCX, nRectCY - nHalfRectHeight);
				graphics._l(nRectCX + nRectHeight, nRectCY + nHalfRectHeight);
				break;
			}
			case SPINBUTTON_DIRECTION_DOWN: {
				graphics._m(nRectCX - nRectHeight, nRectCY - nHalfRectHeight);
				graphics._l(nRectCX, nRectCY + nHalfRectHeight);
				graphics._l(nRectCX + nRectHeight, nRectCY - nHalfRectHeight);
				break;
			}
			case SPINBUTTON_DIRECTION_LEFT: {
				graphics._m(nRectCX + nHalfRectHeight, nRectCY - nRectHeight);
				graphics._l(nRectCX - nHalfRectHeight, nRectCY);
				graphics._l(nRectCX + nHalfRectHeight, nRectCY + nRectHeight);
				break;
			}
			case SPINBUTTON_DIRECTION_RIGHT: {
				graphics._m(nRectCX - nHalfRectHeight, nRectCY - nRectHeight);
				graphics._l(nRectCX + nHalfRectHeight, nRectCY);
				graphics._l(nRectCX - nHalfRectHeight, nRectCY + nRectHeight);
				break;
			}
			default: {
				break;
			}
		}
		graphics.ds();
		graphics._z();
		// graphics.df();
		graphics.RestoreGrState();
		graphics._e();
	};

	function CSpinController(oControl) {
		CControlControllerBase.call(this, oControl);
		this.upButton = new CSpinButton(this, SPINBUTTON_DIRECTION_UP);
		this.downButton = new CSpinButton(this, SPINBUTTON_DIRECTION_DOWN);
		this.stepManager = new CStepManager();
		this.initButtonEventHandlers();
	};
	AscFormat.InitClassWithoutType(CSpinController, CControlControllerBase);
	CSpinController.prototype.startChangeValue = function (fCallback) {
		Asc.editor.startGroupActions();
		this.stepManager.start(fCallback);
	};
	CSpinController.prototype.endChangeValue = function (oDrawingController) {
		const oFormControlPr = this.getFormControlPr();
		const nCurrentValue = oFormControlPr.val;
		Asc.editor.cancelGroupActions();
		this.stepManager.end();
		if (nCurrentValue !== null) {
			const oThis = this;
			oDrawingController.checkObjectsAndCallback(function () {
				oFormControlPr.setVal(nCurrentValue);
				const oRef = oThis.getParsedFmlaLink();
				if (oRef) {
					const oCellValue = new AscCommonExcel.CCellValue();
					oCellValue.number = nCurrentValue;
					oThis.setRangeValue(oRef, oCellValue);
				}
			}, [], undefined, AscDFH.historydescription_Spreadsheet_IncrementControl, [], true);
		}
	};
	CSpinController.prototype.initButtonEventHandlers = function () {
		const oController = this;

		this.downButton._onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
			oController.startChangeValue(oController.decrement.bind(oController, oDrawingController));
		};
		this.upButton._onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
			oController.startChangeValue(oController.increment.bind(oController, oDrawingController));
		};
		this.downButton._onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
			oController.endChangeValue(oDrawingController);
			return true;
		};
		this.upButton._onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
			oController.endChangeValue(oDrawingController);
			return true;
		};
	};
	CSpinController.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		graphics.SaveGrState();
		transform = transform || this.control.transform;
		graphics.transform3(transform);
		startRoundControl(graphics, 0, 0, this.control.extX, this.control.extY, 2, getFlatPenColor());
		this.downButton.draw(graphics);
		this.upButton.draw(graphics);
		endRoundControl(graphics);
		graphics.RestoreGrState();
	};
	CSpinController.prototype.getCursorInfo = function (e, nX, nY) {
		const oControl = this.control;
		if (!oControl.hit(nX, nY)) {
			return null;
		}
		return {cursorType: "pointer", objectId: oControl.GetId()};
	};
	CSpinController.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (e.button !== 0) {
			return false;
		}
		if (e.CtrlKey) {
			return false;
		}
		if (this.upButton.onMouseDown(e, nX, nY, nPageIndex, oDrawingController)) {
			return true;
		}
		if (this.downButton.onMouseDown(e, nX, nY, nPageIndex, oDrawingController)) {
			return true;
		}
		return false;
	};
	CSpinController.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		this.upButton.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
		this.downButton.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
	};
	CSpinController.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		let bRet = 0;
		bRet |= this.upButton.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		bRet |= this.downButton.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		return !!bRet;
	};
	CSpinController.prototype.initTextProperties = function () {};
	CSpinController.prototype.applySpecialPasteProps = function (oPastedWb) {};
	CSpinController.prototype.hit = function (nX, nY) {
		return this.upButton.hit(nX, nY) || this.downButton.hit(nX, nY);
	};
	CSpinController.prototype.increment = function (oDrawingController) {
		this.doAction(oDrawingController, true);
	};
	CSpinController.prototype.decrement = function (oDrawingController) {
		this.doAction(oDrawingController, false);
	};
	CSpinController.prototype.doAction = function (oDrawingController, isIncrement) {
		const oRef = this.getParsedFmlaLink();
		const oFormControlPr = this.getFormControlPr();
		const nIncValue = AscFormat.isRealNumber(oFormControlPr.inc) ? oFormControlPr.inc : 1;
		const nMaxValue = oFormControlPr.max || 0;
		const nMinValue = oFormControlPr.min || 0;
		let nCurrentValue;
		if (oRef) {
			let nRefValue = 0;
			oRef.worksheet._getCellNoEmpty(oRef.bbox.r1, oRef.bbox.c1, function (oCell) {
				if (oCell && !oCell.isNullText()) {
					nRefValue = oCell.number;
				}
			});
			if (nRefValue === null) {
				nRefValue = oFormControlPr.val;
			}
			nCurrentValue = nRefValue || 0;
		} else {
			nCurrentValue = oFormControlPr.val || 0;
		}
		const nNewValue = Math.min(Math.max(isIncrement ? nCurrentValue + nIncValue : nCurrentValue - nIncValue, nMinValue), nMaxValue);
		oFormControlPr.setVal(nNewValue);
		if (oRef) {
			const oCellValue = new AscCommonExcel.CCellValue();
			oCellValue.number = nNewValue;
			this.setRangeValue(oRef, oCellValue);
		}
	};
	CSpinController.prototype.recalculateTransform = function () {
		const oControl = this.control;
		const oControlMatrix = oControl.transform;
		const nHalfHeight = oControl.extY / 2;
		this.upButton.x = oControl.x;
		this.upButton.y = oControl.y;
		this.upButton.extX = oControl.extX;
		this.upButton.extY = nHalfHeight;
		this.upButton.transform = oControlMatrix.CreateDublicate();
		this.upButton.invertTransform = global_MatrixTransformer.Invert(this.upButton.transform);

		const oDownMatrix = oControlMatrix.CreateDublicate();
		const nUpX = oDownMatrix.TransformPointX(0, 0);
		const nUpY = oDownMatrix.TransformPointY(0, 0);
		const nDownX = oDownMatrix.TransformPointX(0, nHalfHeight);
		const nDownY = oDownMatrix.TransformPointY(0, nHalfHeight);
		oDownMatrix.Translate(nDownX - nUpX, nDownY - nUpY);
		this.downButton.transform = oDownMatrix;
		this.downButton.invertTransform = global_MatrixTransformer.Invert(this.downButton.transform);
		this.downButton.x = oControl.x;
		this.downButton.y = oControl.y + nHalfHeight;
		this.downButton.extX = oControl.extX;
		this.downButton.extY = nHalfHeight;
	};


	const SCROLLBAR_THUMB_MIN_SIZE_RATIO = 0.2;

	function CTrackArea(oController) {
		CButtonBase.call(this, oController);
		this.trackAreaHoldDirection = 0;
		this.currentMousePostion = {
			x: 0,
			y: 0
		};
	}
	AscFormat.InitClassWithoutType(CTrackArea, CButtonBase);
	CTrackArea.prototype.getFlatFillColor = function () {
		return [244, 244, 244, 255];
	}
	CTrackArea.prototype.draw = function (graphics) {
		const bOldHold = this.isHold;
		this.isHold = false;
		CButtonBase.prototype.draw.call(this, graphics);
		this.isHold = bOldHold;
	}

	function CScrollController(oControl) {
		CControlControllerBase.call(this, oControl);
		this.scroll = new CScrollContainer(this);
		this.initButtonEventHandlers();
	}

	AscFormat.InitClassWithoutType(CScrollController, CControlControllerBase);
	CScrollController.prototype.startGroupPoints = function () {
		Asc.editor.startGroupActions();
	};
	CScrollController.prototype.endGroupPoints = function (oDrawingController) {
		const nCurrentValue = this.scroll.getCurrentValue();
		Asc.editor.cancelGroupActions();
		this.scroll.setValue(nCurrentValue, oDrawingController);
	};
	CScrollController.prototype.initButtonEventHandlers = function () {
		const oController = this;
		this.scroll.getMinValue = function () {
			const oFormControlPr = oController.getFormControlPr();
			return oFormControlPr.min || 0;
		};
		this.scroll.getMaxValue = function () {
			const oFormControlPr = oController.getFormControlPr();
			return oFormControlPr.max || 0;
		};
		this.scroll.getCurrentValue = function () {
			const oFormControlPr = oController.getFormControlPr();
			const oRef = oController.getParsedFmlaLink();
			let nCurrentValue;
			if (oRef) {
				let nRefValue = 0;
				oRef.worksheet._getCellNoEmpty(oRef.bbox.r1, oRef.bbox.c1, function (oCell) {
					if (oCell && !oCell.isNullText()) {
						nRefValue = oCell.number;
					}
				});
				if (nRefValue === null) {
					nRefValue = oFormControlPr.val;
				}
				nCurrentValue = nRefValue || 0;
			} else {
				nCurrentValue = oFormControlPr.val || 0;
			}
			return nCurrentValue;
		};
		this.scroll._setValue = function (newValue, oDrawingController) {
			const oThis = this;
			oDrawingController.checkObjectsAndCallback(function () {
				const nRoundedValue = Math.round(newValue);
				const oFormControlPr = oController.getFormControlPr();
				oFormControlPr.setVal(nRoundedValue);

				const oRef = oController.getParsedFmlaLink();
				if (oRef) {
					const oCellValue = new AscCommonExcel.CCellValue();
					oCellValue.number = nRoundedValue;
					oController.setRangeValue(oRef, oCellValue);
				}

				oThis.recalculateThumbTransform(oThis.getMinValue(), oThis.getMaxValue(), nRoundedValue);
			}, [], undefined, AscDFH.historydescription_Spreadsheet_IncrementControl, [], true);
		};
		this.scroll.getPageValue = function () {
			const oFormControlPr = oController.getFormControlPr();
			const nPage = oFormControlPr.getPage();
			return AscFormat.isRealNumber(nPage) ? nPage : 1;
		};
		this.scroll.getIncrementValue = function () {
			const oFormControlPr = oController.getFormControlPr();
			const nInc = oFormControlPr.getInc();
			return AscFormat.isRealNumber(nInc) ? nInc : 1;
		};
		this.scroll.onStartChangeValues = function () {
			oController.startGroupPoints();
		};
		this.scroll.onEndChangeValues = function (oDrawingController) {
			oController.endGroupPoints(oDrawingController);
		};
	};

	CScrollController.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		graphics.SaveGrState();
		const oControl = this.control;
		transform = transform || oControl.transform;
		graphics.transform3(transform);
		startRoundControl(graphics, 0, 0, oControl.extX, oControl.extY, 2, getFlatPenColor());
		this.scroll.draw(graphics);
		endRoundControl(graphics);
		graphics.RestoreGrState();
	};

	CScrollController.prototype.getCursorInfo = function (e, nX, nY) {
		return this.scroll.getCursorInfo(nX, nY);
	};
	CScrollController.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		return this.scroll.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
	};
	CScrollController.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		const oControl = this.control;
		if (e.button !== 0) {
			return false;
		}
		if (e.CtrlKey) {
			return false;
		}
		return this.scroll.onMouseDown(e, nX, nY, nPageIndex, oDrawingController);
	};

	CScrollController.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		return this.scroll.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
	};

	CScrollController.prototype.recalculateTransform = function () {
		const oControl = this.control;
		this.scroll.x = oControl.x;
		this.scroll.y = oControl.y;
		this.scroll.extX = oControl.extX;
		this.scroll.extY = oControl.extY;
		this.scroll.transform = oControl.transform.CreateDublicate();
		this.scroll.invertTransform = oControl.invertTransform.CreateDublicate();
		this.scroll.recalculateTransform();
	};

	function CThumbButton(oController) {
		CButtonBase.call(this, oController);
	}
	AscFormat.InitClassWithoutType(CThumbButton, CButtonBase);
	CThumbButton.prototype.getFlatFillColor = function () {
		if (this.isHold || this.isHover) {
			return [51, 51, 51, 255];
		}
		return [148, 148, 148, 255];
	};
	CThumbButton.prototype.draw = function (graphics) {
		graphics.SaveGrState();
		graphics.transform3(this.transform);
		startRoundControl(graphics, 0, 0, this.extX, this.extY, 4, getFlatPenColor());
		CButtonBase.prototype.draw.call(this, graphics);
		endRoundControl(graphics);
		graphics.RestoreGrState();
	};

	function CScrollContainer(oController) {
		this.controller = oController;
		this.upButton = new CSpinButton(oController);
		this.downButton = new CSpinButton(oController);
		this.thumb = new CThumbButton(oController);
		this.trackArea = new CTrackArea(oController);
		this.stepManager = new CStepManager();
		this.x = null;
		this.y = null;
		this.extX = null;
		this.extY = null;
		this.transform = new AscCommon.CMatrix();
		this.invertTransform = new AscCommon.CMatrix();
		this.isChangeValues = false;
		this.initButtonEventHandlers();
	}
	CScrollContainer.prototype.getMinValue = function () {

	};
	CScrollContainer.prototype.getMaxValue = function () {

	};
	CScrollContainer.prototype.getCurrentValue = function () {

	};
	CScrollContainer.prototype.getIncrementValue = function () {
		return 1;
	}
	CScrollContainer.prototype.getPageValue = function () {
		return 1;
	}
	CScrollContainer.prototype._setValue = function () {

	};
	CScrollContainer.prototype.onEndChangeValues = function () {

	};
	CScrollContainer.prototype.onStartChangeValues = function () {

	};
	CScrollContainer.prototype.isCanScroll = function () {
		return (this.getMaxValue() - this.getMinValue()) > 0;
	};
	CScrollContainer.prototype.setValue = function (nNewValue, oDrawingController) {
		this.startChangeValues();
		this._setValue(nNewValue, oDrawingController);
	};
	CScrollContainer.prototype.isVertical = function () {
		return this.extY > this.extX;
	};
	CScrollContainer.prototype.increment = function (oDrawingController) {
		this.doIncrement(oDrawingController, true);
	};
	CScrollContainer.prototype.decrement = function (oDrawingController) {
		this.doIncrement(oDrawingController, false);
	};
	CScrollContainer.prototype.doIncrement = function (oDrawingController, bIsIncrement) {
		const nIncValue = this.getIncrementValue();
		const nMinValue = this.getMinValue();
		const nMaxValue = this.getMaxValue();
		const nCurrentValue = this.getCurrentValue();
		const nNewValue = bIsIncrement ?
			Math.min(nMaxValue, nCurrentValue + nIncValue) :
			Math.max(nMinValue, nCurrentValue - nIncValue);
		this.setValue(nNewValue, oDrawingController);
	};
	CScrollContainer.prototype.endChangeValues = function (oDrawingController) {
		if (this.isChangeValues) {
			this.onEndChangeValues(oDrawingController);
			this.isChangeValues = false;
		}
	};
	CScrollContainer.prototype.startChangeValues = function () {
		if (!this.isChangeValues) {
			this.isChangeValues = true;
			this.onStartChangeValues();
		}
	};
	CScrollContainer.prototype.initButtonEventHandlers = function () {
		const oThis = this;

		this.upButton._onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
			oThis.stepManager.start(oThis.decrement.bind(oThis, oDrawingController));
		};

		this.upButton._onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
			oThis.endChangeValues(oDrawingController);
			oThis.stepManager.end();
			return true;
		};

		this.downButton._onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
			oThis.stepManager.start(oThis.increment.bind(oThis, oDrawingController));
		};
		this.downButton._onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
			oThis.endChangeValues(oDrawingController);
			oThis.stepManager.end();
			return true;
		};

		this.thumb._onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
			return true;
		};
		this.thumb._onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
			if (this.isHold) {

				const nTrackSize = oThis.getThumbTrackSpace();
				const bIsVertical = oThis.isVertical();
				let nRelPos;
				if (bIsVertical) {
					nRelPos = nY - (oThis.trackArea.y + oThis.getThumbTrackOffset());
				} else {
					nRelPos = nX - (oThis.trackArea.x + oThis.getThumbTrackOffset());
				}

				const nMin = oThis.getMinValue();
				const nMax = oThis.getMaxValue();
				const nRange = nMax - nMin;

				let nNewValue = (nRelPos / nTrackSize) * nRange;
				nNewValue = Math.max(nMin, Math.min(nMax, nNewValue));
				oThis.setValue(nNewValue, oDrawingController);
				return true;
			}
		};
		this.thumb._onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
			oThis.endChangeValues(oDrawingController);
			return true;
		};
		this.trackArea.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
			if (!this.hit(nX, nY)) {
				return false;
			}

			this.currentMousePostion.x = nX;
			this.currentMousePostion.y = nY;
			if (!oThis.initPerformTrackAreaStep()) {
				return false;
			}
			this.setIsHold(true);
			oThis.stepManager.start(oThis.performTrackAreaStep.bind(oThis, oDrawingController));
			return true;
		};
		this.trackArea._onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
			this.trackAreaHoldDirection = 0;
			oThis.stepManager.end();
			oThis.endChangeValues(oDrawingController);
			return true;
		};
		this.trackArea._onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
			if (this.isHold) {
				this.currentMousePostion.x = nX;
				this.currentMousePostion.y = nY;
			}
			return false;
		};
	};
	CScrollContainer.prototype.initPerformTrackAreaStep = function () {
		let nLocalPos, nThumbMin, nThumbMax;
		const bIsVertical = this.isVertical();
		const oThumb = this.thumb;
		const oTrackArea = this.trackArea;
		const nX = oTrackArea.currentMousePostion.x;
		const nY = oTrackArea.currentMousePostion.y;
		if (bIsVertical) {
			nLocalPos = oTrackArea.invertTransform.TransformPointY(nX, nY);
			nThumbMin = oThumb.y - oTrackArea.y;
			nThumbMax = nThumbMin + oThumb.extY;
		} else {
			nLocalPos = oTrackArea.invertTransform.TransformPointX(nX, nY);
			nThumbMin = oThumb.x - oTrackArea.x;
			nThumbMax = nThumbMin + oThumb.extX;
		}
		if (nLocalPos >= nThumbMin && nLocalPos <= nThumbMax) {
			return false;
		}

		const nTempTrackAreaHold = (nLocalPos < nThumbMin) ? -1 : 1;

		if (oTrackArea.trackAreaHoldDirection === 0 || nTempTrackAreaHold === oTrackArea.trackAreaHoldDirection) {
			oTrackArea.trackAreaHoldDirection = nTempTrackAreaHold;
			return true;
		}
		return false;
	};
	CScrollContainer.prototype.performTrackAreaStep = function (oDrawingController, bSkipUpdate) {
		if (this.initPerformTrackAreaStep()) {
			const nMinValue = this.getMinValue();
			const nMaxValue = this.getMaxValue();
			const nPageSize = this.getPageValue();

			const nValue = this.getCurrentValue();
			const nNewValue = this.trackArea.trackAreaHoldDirection === -1 ? nValue - nPageSize : nValue + nPageSize;
			this.setValue(Math.min(Math.max(nMinValue, nNewValue), nMaxValue), oDrawingController);

			if (!bSkipUpdate) {
				this.controller.checkNeedUpdate();
			}
			return true;
		} else {
			this.stepManager.end();
		}
		return false;
	};
	CScrollContainer.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (!this.isCanScroll()) {
			return !this.selected && !e.CtrlKey && e.button === 0 && this.hit(nX, nY);
		}
		return this.upButton.onMouseDown(e, nX, nY, nPageIndex, oDrawingController) || this.downButton.onMouseDown(e, nX, nY, nPageIndex, oDrawingController) ||
			this.thumb.onMouseDown(e, nX, nY, nPageIndex, oDrawingController) || this.trackArea.onMouseDown(e, nX, nY, nPageIndex, oDrawingController);
	};
	CScrollContainer.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (!this.isCanScroll()) {
			return;
		}
		this.trackArea.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
		this.upButton.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
		this.downButton.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
		this.thumb.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
	};
	CScrollContainer.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (!this.isCanScroll()) {
			return;
		}
		let bButtonRet = 0;
		bButtonRet |= this.upButton.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		bButtonRet |= this.downButton.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		bButtonRet |= this.thumb.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		bButtonRet |= this.trackArea.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		return bButtonRet;
	};

	CScrollContainer.prototype.recalculateTransform = function () {
		if (!this.isCanScroll()) {
			return;
		}

		const oControlMatrix = this.transform;

		const nScrollWidth = this.extX;
		const nScrollHeight = this.extY;

		const bIsVertical = this.isVertical();
		const nMinValue = this.getMinValue();
		const nMaxValue = this.getMaxValue();
		const nCurrentValue = this.getCurrentValue();
		if (bIsVertical) {
			const nButtonWidth = this.extX;
			const nButtonHeight = Math.min(this.extY / 2, nButtonWidth);
			const nTrackHeight = nScrollHeight - (nButtonHeight * 2);
			this.upButton.x = this.x;
			this.upButton.y = this.y;
			this.upButton.extX = nButtonWidth;
			this.upButton.extY = nButtonHeight;
			this.upButton.transform = oControlMatrix.CreateDublicate();
			this.upButton.invertTransform = global_MatrixTransformer.Invert(this.upButton.transform);
			this.upButton.direction = SPINBUTTON_DIRECTION_UP;

			const oDownButtonMatrix = oControlMatrix.CreateDublicate();
			global_MatrixTransformer.TranslateAppend(oDownButtonMatrix, 0, nScrollHeight - nButtonHeight);
			this.downButton.transform = oDownButtonMatrix;
			this.downButton.invertTransform = global_MatrixTransformer.Invert(this.downButton.transform);
			this.downButton.x = this.x;
			this.downButton.y = this.y + nScrollHeight - nButtonHeight;
			this.downButton.extX = nButtonWidth;
			this.downButton.extY = nButtonHeight;
			this.downButton.direction = SPINBUTTON_DIRECTION_DOWN;

			this.trackArea.x = this.x;
			this.trackArea.y = this.y + nButtonHeight;
			this.trackArea.extX = nScrollWidth;
			this.trackArea.extY = nTrackHeight;
			const oTrackMatrix = oControlMatrix.CreateDublicate();
			global_MatrixTransformer.TranslateAppend(oTrackMatrix, 0, nButtonHeight);
			this.trackArea.transform = oTrackMatrix;
			this.trackArea.invertTransform = global_MatrixTransformer.Invert(this.trackArea.transform);
		} else {
			const nButtonHeight = this.extY;
			const nButtonWidth = Math.min(nButtonHeight, this.extX / 2);

			const nTrackWidth = nScrollWidth - (nButtonWidth * 2);

			this.upButton.x = this.x;
			this.upButton.y = this.y;
			this.upButton.extX = nButtonWidth;
			this.upButton.extY = nButtonHeight;
			this.upButton.transform = oControlMatrix.CreateDublicate();
			this.upButton.invertTransform = global_MatrixTransformer.Invert(this.upButton.transform);
			this.upButton.direction = SPINBUTTON_DIRECTION_LEFT;

			const oDownButtonMatrix = oControlMatrix.CreateDublicate();
			global_MatrixTransformer.TranslateAppend(oDownButtonMatrix, nScrollWidth - nButtonWidth, 0);
			this.downButton.transform = oDownButtonMatrix;
			this.downButton.invertTransform = global_MatrixTransformer.Invert(this.downButton.transform);
			this.downButton.x = this.x + nScrollWidth - nButtonWidth;
			this.downButton.y = this.y;
			this.downButton.extX = nButtonWidth;
			this.downButton.extY = nButtonHeight;
			this.downButton.direction = SPINBUTTON_DIRECTION_RIGHT;

			this.trackArea.x = this.x + nButtonWidth;
			this.trackArea.y = this.y;
			this.trackArea.extX = nTrackWidth;
			this.trackArea.extY = nScrollHeight;
			const oTrackMatrix = oControlMatrix.CreateDublicate();
			global_MatrixTransformer.TranslateAppend(oTrackMatrix, nButtonWidth, 0);
			this.trackArea.transform = oTrackMatrix;
			this.trackArea.invertTransform = global_MatrixTransformer.Invert(this.trackArea.transform);
		}
		this.recalculateThumbTransform(nMinValue, nMaxValue, nCurrentValue);
	};
	const ThumbTrackSpaceOffset = 2;
	CScrollContainer.prototype.getThumbTrackSpace = function () {
		if (this.isVertical()) {
			return this.trackArea.extY - ThumbTrackSpaceOffset - this.thumb.extY;
		}
		return this.trackArea.extX - ThumbTrackSpaceOffset - this.thumb.extX;
	};
	CScrollContainer.prototype.getThumbTrackOffset = function () {
		return ThumbTrackSpaceOffset / 2;
	};
	CScrollContainer.prototype.recalculateThumbTransform = function (nMinValue, nMaxValue, nCurrentValue) {
		const bIsVertical = this.isVertical();
		const nValueRange = nMaxValue - nMinValue;
		const nThumbPositionRatio = nValueRange > 0 ? (nCurrentValue - nMinValue) / nValueRange : 0;
		const nButtonSide = this.upButton.extY;

		if (bIsVertical) {
			const nScrollWidth = this.extX;
			const nTrackHeight = this.trackArea.extY;
			const nThumbMinHeight = nScrollWidth * SCROLLBAR_THUMB_MIN_SIZE_RATIO;
			const nThumbHeight = Math.max(nThumbMinHeight, nTrackHeight / (nValueRange + 1));
			this.thumb.extY = nThumbHeight;

			const nThumbTrackSpace = this.getThumbTrackSpace();
			const nThumbY = nThumbPositionRatio * nThumbTrackSpace + this.getThumbTrackOffset();

			const nThumbWidth = nScrollWidth * 0.8;
			const nXOffset = (nScrollWidth - nThumbWidth) / 2;
			this.thumb.x = this.x + nXOffset;
			this.thumb.y = this.y + nButtonSide + nThumbY;
			this.thumb.extX = nThumbWidth;


			const oThumbMatrix = this.transform.CreateDublicate();
			global_MatrixTransformer.TranslateAppend(oThumbMatrix, nXOffset, nButtonSide + nThumbY);
			this.thumb.transform = oThumbMatrix;
			this.thumb.invertTransform = global_MatrixTransformer.Invert(this.thumb.transform);
		} else {
			const nScrollHeight = this.extY;
			const nTrackWidth = this.trackArea.extX;
			const nThumbMinWidth = nScrollHeight * SCROLLBAR_THUMB_MIN_SIZE_RATIO;
			const nThumbWidth = Math.max(nThumbMinWidth, nTrackWidth / (nValueRange + 1));
			this.thumb.extX = nThumbWidth;
			const nThumbTrackSpace = this.getThumbTrackSpace();
			const nThumbX = nThumbPositionRatio * nThumbTrackSpace + this.getThumbTrackOffset();

			const nThumbHeight = nScrollHeight * 0.8;
			const nYOffset = (nScrollHeight - nThumbHeight) / 2;
			this.thumb.x = this.x + nButtonSide + nThumbX;
			this.thumb.y = this.y + nYOffset;

			this.thumb.extY = nThumbHeight;

			const oThumbMatrix = this.transform.CreateDublicate();
			global_MatrixTransformer.TranslateAppend(oThumbMatrix, nButtonSide + nThumbX, nYOffset);
			this.thumb.transform = oThumbMatrix;
			this.thumb.invertTransform = global_MatrixTransformer.Invert(this.thumb.transform);
		}
	};

	CScrollContainer.prototype.draw = function (graphics) {
		if (this.isCanScroll()) {
			this.trackArea.draw(graphics);
			this.thumb.draw(graphics);
			this.downButton.draw(graphics);
			this.upButton.draw(graphics);
		} else {
			graphics.SaveGrState();
			graphics.transform3(this.transform);
			graphics.b_color1(244,244,244, 255);
			graphics._s();
			graphics._m(0, 0);
			graphics._l(this.extX, 0);
			graphics._l(this.extX, this.extY);
			graphics._l(0, this.extY);
			graphics._z(0, this.extY);
			graphics.df();
			graphics._e();
			graphics.RestoreGrState();
		}
	};
	CScrollContainer.prototype.hit = function (nX, nY) {
		return AscFormat.HitToRect(nX, nY, this.invertTransform, 0, 0, this.extX, this.extY);
	}
	CScrollContainer.prototype.getCursorInfo = function (nX, nY) {
		const oControl = this.controller.control;
		if (this.thumb.hit(nX, nY) || this.downButton.hit(nX, nY) || this.upButton.hit(nX, nY)) {
			return {cursorType: "pointer", objectId: oControl.GetId()};
		} else if (this.trackArea.hit(nX, nY)) {
			return {cursorType: "default", objectId: oControl.GetId()};
		}
		return null;
	};


	const LISTBOX_ITEM_HEIGHT = 2;
	const LISTBOX_SCROLL_WIDTH = 4;
	const LISTBOX_MAX_ITEM_HEIGHT = 4;

	function getListBoxItemTextPr() {
		const oTextPr = new AscCommonWord.CTextPr();
		oTextPr.Set_FromObject({
			FontFamily: {
				Name:  "Arial",
				Index: -1,
			},
			FontSize:   8,
			Color: {r: 0,	g: 0, b: 0,	a: 255},
			Bold: false
		});
		oTextPr.RFonts.SetAll("Arial");
		return oTextPr;
	}

	let FONT_HEIGHT_LISTBOX_ITEM = null;
	function getListBoxItemFontHeight() {
		if (FONT_HEIGHT_LISTBOX_ITEM === null) {
			const oTextPr = getListBoxItemTextPr();
			AscCommon.g_oTextMeasurer.SetTextPr(oTextPr);
			AscCommon.g_oTextMeasurer.SetFontSlot(fontslot_ASCII);
			FONT_HEIGHT_LISTBOX_ITEM = AscCommon.g_oTextMeasurer.GetHeight();
		}
		return FONT_HEIGHT_LISTBOX_ITEM;
	}

	const LISTBOX_TEXT_PADDING = 1;
	function CListBoxItem(oListBox, oParagraph) {
		this.listBox = oListBox;
		this.text = oParagraph || null;
		this.x = 0;
		this.y = 0;
		this.extX = 0;
		this.extY = LISTBOX_ITEM_HEIGHT;
		this.textY = 0;
		this.isSelected = false;
		this.isHovered = false;
	}

	CListBoxItem.prototype.setPosition = function (nX, nY, nWidth, nHeight) {
		this.x = nX;
		this.y = nY;
		this.extX = nWidth;
		this.extY = nHeight || LISTBOX_ITEM_HEIGHT;
	};

	CListBoxItem.prototype.setSelected = function (bSelected) {
		this.isSelected = bSelected;
	};

	CListBoxItem.prototype.setHovered = function (bHovered) {
		this.isHovered = bHovered;
	};
	CListBoxItem.prototype.getItemBackgroundColor = function () {
		if (this.isSelected) {
			return [225, 225, 225, 255];
		}
		if (this.isHovered) {
			return [234, 234, 234, 255];
		}
		return [255, 255, 255, 255];
	};
	CListBoxItem.prototype.getItemTextColor = function () {
		return [51, 51, 51, 255];
	};

	CListBoxItem.prototype.draw = function (graphics, oListBoxTransform) {
		graphics.SaveGrState();
		graphics.b_color1.apply(graphics, this.getItemBackgroundColor());
		graphics.SetIntegerGrid(true);
		graphics.TableRect(this.x, this.y, this.extX, this.extY);

		if (this.text) {
			const nTextX = this.x + LISTBOX_TEXT_PADDING;
			const oTransform = oListBoxTransform.CreateDublicate();
			oTransform.Translate(nTextX, this.textY);
			graphics.transform3(oTransform);
			drawParagraph(this.text, graphics);
		}
		graphics.RestoreGrState();
	};

	CListBoxItem.prototype.hit = function (nX, nY) {
		return nX >= this.x && nX <= this.x + this.extX &&
			   nY >= this.y && nY <= this.y + this.extY;
	};

	CListBoxItem.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController, nItemIndex) {
		if (this.hit(nX, nY)) {
			this.listBox.handleItemClick(nItemIndex, e.CtrlKey, oDrawingController);
			return true;
		}
		return false;
	};
	CListBoxItem.prototype.recalculateTextPosition = function () {
		this.textY = this.y + (this.extY - getListBoxItemFontHeight()) / 2;
	};



	function CListBox(oController) {
		this.controller = oController;
		this.selectedIndices = {};
		this.scrollPosition = 0;
		this.scrollContainer = new CScrollContainer(oController);
		this.listItems = [];
		this.visibleItemsCount = 0;
		this.x = 0;
		this.y = 0;
		this.extX = 0;
		this.extY = 0;
		this.transform = new AscCommon.CMatrix();
		this.invertTransform = new AscCommon.CMatrix();
		this.stringCacheManager = new CStringCacheManager();
		this.initScrollContainer();
	}
	CListBox.prototype.isMultiSelection = function () {
		return false;
	};

	CListBox.prototype.isExtendedSelection = function () {
		return false;
	};

	CListBox.prototype.isShowScroll = function () {
		return true;
	};

	CListBox.prototype.onChangeSelection = function (oDrawingController) {
	};

	CListBox.prototype.initScrollContainer = function () {
		const oThis = this;


		this.scrollContainer.getMinValue = function () {
			return 0;
		};
		
		this.scrollContainer.getMaxValue = function () {
			return Math.max(0, oThis.listItems.length - oThis.visibleItemsCount);
		};
		
		this.scrollContainer.getCurrentValue = function () {
			return oThis.scrollPosition;
		};
		
		this.scrollContainer.getIncrementValue = function () {
			return 1;
		};
		
		this.scrollContainer.getPageValue = function () {
			return Math.max(1, Math.floor(oThis.visibleItemsCount / 2));
		};
		
		this.scrollContainer._setValue = function (newValue, oDrawingController) {
			oThis.scrollPosition = Math.round(newValue);
			oThis.recalculateTransform();
			oThis.controller.checkNeedUpdate();
		};
		this.scrollContainer.isVertical = function () {
			return true;
		}
	};
	CListBox.prototype.updateListItems = function () {
		this.listItems = [];
		const oStringCacheManager = this.stringCacheManager;
		oStringCacheManager.startCheckDeleteParagraphs();
		const oFormControlPr = this.controller.getFormControlPr();
		if (oFormControlPr.itemLst.length) {
			for (let i = 0; i < oFormControlPr.itemLst.length; i += 1) {
				this.listItems.push(new CListBoxItem(this, oStringCacheManager.getParagraphWithText(oFormControlPr.itemLst[i])));
			}
		}
		else {
			const oRange = this.controller.getParsedFmlaRange();
			if (oRange) {
				const oThis = this;
				oRange._foreach(function (oCell) {
					const sItem = oCell && !oCell.isNullText() ? oCell.getValue() : "";
					oThis.listItems.push(new CListBoxItem(oThis, oStringCacheManager.getParagraphWithText(sItem)));
				});
			}
		}
		oStringCacheManager.endCheckDeleteParagraphs();
	};

	CListBox.prototype.recalculateItemPositions = function () {
		if (this.listItems.length === 0) {
			return;
		}
		const nItemCount = this.extY / LISTBOX_MAX_ITEM_HEIGHT;
		const nMaxItemCount = Math.ceil(nItemCount);
		let nCalcItemHeight = this.extY / nMaxItemCount;
		const nListWidth = this.extX;
		if (nCalcItemHeight < LISTBOX_ITEM_HEIGHT) {
			nCalcItemHeight = LISTBOX_ITEM_HEIGHT;
			this.visibleItemsCount = Math.ceil(this.extY / LISTBOX_ITEM_HEIGHT);
		} else {
			this.visibleItemsCount = nMaxItemCount;
		}
		for (let i = 0; i < this.listItems.length; i += 1) {
			this.listItems[i].setPosition(0, 0, nListWidth, nCalcItemHeight);
		}
		this.recalculateVisibleItemPositions();
	};

	CListBox.prototype.recalculateVisibleItemPositions = function () {
		let nIndex = 0;
		this.checkVisibleItems(function (oItem) {
			oItem.y = oItem.extY * nIndex;
			oItem.recalculateTextPosition();
			nIndex++;
		});
	};
	CListBox.prototype.resetSelectedIndices = function () {
		for (let sIndex in this.selectedIndices) {
			this.selectedIndices[sIndex].setSelected(false);
		}
		this.selectedIndices = {};
	};
	CListBox.prototype.addSelectedIndex = function (nIndex) {
		nIndex = Math.max(0, Math.min(nIndex, this.listItems.length - 1));
		if (this.listItems[nIndex]) {
			this.selectedIndices[nIndex] = this.listItems[nIndex];
			this.selectedIndices[nIndex].setSelected(true);
			return nIndex;
		}
		return null;
	};
	CListBox.prototype.removeSelectedIndex = function (nIndex) {
		if (this.selectedIndices[nIndex]) {
			this.selectedIndices[nIndex].setSelected(false);
			delete this.selectedIndices[nIndex];
			return nIndex;
		}
		return null;
	};
	CListBox.prototype.handleItemClick = function (nItemIndex, bCtrlKey, oDrawingController) {
		if (this.isMultiSelection() || this.isExtendedSelection() && bCtrlKey) {
			if (this.selectedIndices[nItemIndex]) {
				this.removeSelectedIndex(nItemIndex);
			} else {
				this.addSelectedIndex(nItemIndex);
			}
		} else {
			this.resetSelectedIndices();
			this.addSelectedIndex(nItemIndex);
		}
		this.onChangeSelection(oDrawingController);
		this.controller.checkNeedUpdate();
	};
	CListBox.prototype.recalculateTransform = function() {
		this.recalculateItemPositions();
		this.recalculateScrollTransform();
	};
	CListBox.prototype.draw = function (graphics, transform) {
		const oTransform = transform || this.transform;
		graphics.SaveGrState();
		graphics.transform3(oTransform);
		startRoundControl(graphics, 0, 0, this.extX, this.extY, 2, getFlatPenColor());
		graphics.b_color1(255, 255, 255, 255);
		graphics.TableRect(0, 0, this.extX, this.extY);
		this.checkVisibleItems(function (oItem) {
			oItem.draw(graphics, oTransform);
		});

		if (this.isShowScroll() && this.listItems.length > this.visibleItemsCount) {
			this.scrollContainer.draw(graphics);
		}
		endRoundControl(graphics);
		graphics.RestoreGrState();
	};

	CListBox.prototype.recalculateScrollTransform = function () {
		const nScrollWidth = this.isShowScroll() ? LISTBOX_SCROLL_WIDTH : 0;
		const nListWidth = this.extX - nScrollWidth;
		
		this.scrollContainer.x = this.x + nListWidth;
		this.scrollContainer.y = this.y;
		this.scrollContainer.extX = nScrollWidth;
		this.scrollContainer.extY = this.extY;
		this.scrollContainer.transform = this.transform.CreateDublicate();
		global_MatrixTransformer.TranslateAppend(this.scrollContainer.transform, nListWidth, 0);
		this.scrollContainer.invertTransform = global_MatrixTransformer.Invert(this.scrollContainer.transform);
		this.scrollContainer.recalculateTransform();
	};
	CListBox.prototype.checkVisibleItems = function (fCallback) {
		const nStartIndex = this.scrollPosition;
		const nEndIndex = Math.min(nStartIndex + this.visibleItemsCount + 1, this.listItems.length);
		for (let i = nStartIndex; i < nEndIndex; i++) {
			if (fCallback(this.listItems[i], i)) {
				return true;
			}
		}
		return false;
	}
	CListBox.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (!this.hit(nX, nY)) {
			return false;
		}
		if (e.CtrlKey && !this.isExtendedSelection()) {
			return false;
		}
		if (this.isShowScroll() && this.scrollContainer.onMouseDown(e, nX, nY, nPageIndex, oDrawingController)) {
			return true;
		}

		const nLocalX = this.invertTransform.TransformPointX(nX, nY);
		const nLocalY = this.invertTransform.TransformPointY(nX, nY);
		this.checkVisibleItems(function (oItem, nItemIndex) {
			return oItem.onMouseDown(e, nLocalX, nLocalY, nPageIndex, oDrawingController, nItemIndex);
		});
		return true;
	};

	CListBox.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (this.isShowScroll()) {
			return this.scrollContainer.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
		}
		return false;
	};
	CListBox.prototype.resetHoverEffect = function () {
		let bUpdate = false;
		this.checkVisibleItems(function (oItem) {
			if (oItem.isHovered) {
				bUpdate = true;
				oItem.setHovered(false);
			}
			return bUpdate;
		});
		if (bUpdate) {
			this.controller.checkNeedUpdate();
		}
	};
	CListBox.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (this.isShowScroll()) {
			const bIsScrollMove = this.scrollContainer.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
			const bIsHitScroll = this.scrollContainer.hit(nX, nY);
			if (bIsHitScroll || bIsScrollMove) {
				if (bIsHitScroll) {
					this.resetHoverEffect();
				}
				return bIsScrollMove;
			}
		}

			const nLocalX = this.invertTransform.TransformPointX(nX, nY);
			const nLocalY = this.invertTransform.TransformPointY(nX, nY);

			let bUpdate = 0;
			let bHit = false;
			this.checkVisibleItems(function (oItem) {
				if (bHit) {
					bUpdate |= oItem.isHovered;
					oItem.setHovered(false);
				} else {
					bHit = oItem.hit(nLocalX, nLocalY);
					bUpdate |= oItem.isHovered !== bHit;
					oItem.setHovered(bHit);
				}
			});
			
			if (bUpdate) {
				this.controller.checkNeedUpdate();
			}
		
		return false;
	};

	CListBox.prototype.hit = function (nX, nY) {
		return AscFormat.HitToRect(nX, nY, this.invertTransform, 0, 0, this.extX, this.extY);
	};

	CListBox.prototype.getSingleSelectedIndex = function () {
		if (!this.isMultiSelection() && !this.isExtendedSelection()) {
			for (let sIndex in this.selectedIndices) {
				return parseInt(sIndex, 10);
			}
		}
		return -1;
	};

	function CListBoxController(oControl) {
		CControlControllerBase.call(this, oControl);
		this.listBox = new CListBox(this);
		this.recalcInfo = {
			recalculateItems: true
		};
		this.setupListBoxBehavior();
	}

	AscFormat.InitClassWithoutType(CListBoxController, CControlControllerBase);
	CListBoxController.prototype.handleFmlaRange = function (aRanges) {
		return this.handleRef(aRanges, this.getParsedFmlaRange(), this.recalculateItems.bind(this));
	};
	CListBoxController.prototype.handleFmlaLink = function (aRanges) {
		if (this.listBox.isMultiSelection() || this.listBox.isExtendedSelection()) {
			return false;
		}
		return this.handleRef(aRanges, this.getParsedFmlaLink(), this.updateSelectedIndices.bind(this));
	};
	CListBoxController.prototype.recalculate = function() {
		if (this.recalcInfo.recalculateItems) {
			this.recalcInfo.recalculateItems = false;
			this.recalculateItems();
		}
	};
	CListBoxController.prototype.recalculateItems = function() {
		this.updateListItems();
		this.updateSelectedIndices();
		this.listBox.recalculateTransform();
	};
	CListBoxController.prototype.setupListBoxBehavior = function () {
		const oThis = this;
		const oFormControlPr = this.getFormControlPr();


		this.listBox.isMultiSelection = function () {
			return oFormControlPr.selType === CFormControlPr_selType_multi;
		};
		
		this.listBox.isExtendedSelection = function () {
			return oFormControlPr.selType === CFormControlPr_selType_extended;
		};
		
		this.listBox.isShowScroll = function () {
			return true;
		};
		
		this.listBox.onChangeSelection = function (oDrawingController) {
			oDrawingController.checkObjectsAndCallback(function () {
				oThis.updateSelectionInFormControl();
				oThis.updateLinkedCell();
			}, [], undefined, AscDFH.historydescription_Spreadsheet_SelectListBox, [], true);
		};
	};

	CListBoxController.prototype.updateListItems = function () {
		this.listBox.updateListItems();
	};
	CListBoxController.prototype.getFmlaLinkIndex = function () {
		const oParsedLink = this.getParsedFmlaLink();
		if (oParsedLink) {
			let nIndex = null;
			oParsedLink._foreach(function (oCell) {
				if (oCell) {
					const nNumber = oCell.getNumberValue();
					if (nNumber !== null) {
						nIndex = nNumber - 1;
					}
				}
			});
			return nIndex;
		}
		return null;
	};
	CListBoxController.prototype.updateSelectedIndices = function () {
		const oFormControlPr = this.getFormControlPr();
		this.listBox.resetSelectedIndices();
		if (oFormControlPr.selType === CFormControlPr_selType_multi) {
			if (oFormControlPr.multiSel) {
				const aIndices = oFormControlPr.multiSel.split(", ");
				for (let i = 0; i < aIndices.length; i++) {
					const nIndex = parseInt(aIndices[i], 10) - 1;
					if (!isNaN(nIndex) && nIndex >= 0) {
						this.listBox.addSelectedIndex(nIndex);
					}
				}
			}
		} else if (!this.listBox.isMultiSelection() && !this.listBox.isExtendedSelection()) {
			const nFmlaLinkIndex = this.getFmlaLinkIndex();
			if (nFmlaLinkIndex !== null) {
				if (nFmlaLinkIndex >= 0) {
					this.listBox.addSelectedIndex(nFmlaLinkIndex);
				}
			} else if (typeof oFormControlPr.sel === "number" && oFormControlPr.sel > 0) {
				const nSelectedIndex = oFormControlPr.sel - 1;
				if (nSelectedIndex >= 0) {
					this.listBox.addSelectedIndex(nSelectedIndex);
				}
			}
		}
	};

	CListBoxController.prototype.updateSelectionInFormControl = function () {
		const oFormControlPr = this.getFormControlPr();
		const oListBox = this.listBox;
		if (oFormControlPr.selType === CFormControlPr_selType_multi || oFormControlPr.selType === CFormControlPr_selType_extended) {
			const arrIndexes = [];
			for (let sIndex in oListBox.selectedIndices) {
				arrIndexes.push(parseInt(sIndex, 10) + 1);
			}
			oFormControlPr.setMultiSel(arrIndexes.join(", "));
		} else {
			const nIndex = oListBox.getSingleSelectedIndex() + 1;
			oFormControlPr.setSel(nIndex ? nIndex : null);
		}
	};

	CListBoxController.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		this.listBox.draw(graphics, transform);
	};

	CListBoxController.prototype.recalculateTransform = function () {
		const oControl = this.control;
		this.listBox.x = oControl.x;
		this.listBox.y = oControl.y;
		this.listBox.extX = oControl.extX;
		this.listBox.extY = oControl.extY;
		this.listBox.transform = oControl.transform.CreateDublicate();
		this.listBox.invertTransform = oControl.invertTransform.CreateDublicate();

		this.listBox.recalculateTransform();
	};

	CListBoxController.prototype.getCursorInfo = function (e, nX, nY) {
		const oControl = this.control;
		if (!oControl.hit(nX, nY)) {
			return null;
		}
		return {cursorType: "pointer", objectId: oControl.GetId()};
	};

	CListBoxController.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (e.button !== 0) {
			return false;
		}
		
		return this.listBox.onMouseDown(e, nX, nY, nPageIndex, oDrawingController);
	};

	CListBoxController.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		return this.listBox.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
	};

	CListBoxController.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		return this.listBox.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
	};

	CListBoxController.prototype.updateLinkedCell = function () {
		const oFormControlPr = this.getFormControlPr();
		if (oFormControlPr.selType === CFormControlPr_selType_multi || oFormControlPr.selType === CFormControlPr_selType_extended) {
			return;
		}
		const nIndex = this.listBox.getSingleSelectedIndex() + 1;
		if (!nIndex) {
			return;
		}

		const oRef = this.getParsedFmlaLink();
		if (oRef) {
			const oCellValue = new AscCommonExcel.CCellValue();
			oCellValue.type = AscCommon.CellValueType.Number;
			oCellValue.number = nIndex;
			this.setRangeValue(oRef, oCellValue);
		}
	};

	function CComboBoxController(oControl) {
		CControlControllerBase.call(this, oControl);
		this.listBox = new CListBox(this);
		this.dropButton = new CSpinButton(this, SPINBUTTON_DIRECTION_DOWN);
		this.isDropdownOpen = false;
		this.selectedText = null;
		this.recalcInfo = {
			recalculateItems: true
		};
		this.setupComboBoxBehavior();
	}

	AscFormat.InitClassWithoutType(CComboBoxController, CControlControllerBase);
	CComboBoxController.prototype.recalculate = function () {
		if (this.recalcInfo.recalculateItems) {
			this.recalcInfo.recalculateItems = false;
			this.recalculateItems();
		}
	};
	CComboBoxController.prototype.recalculateItems = function () {
		this.updateListItems();
		this.updateSelectedIndex();
		this.recalculateListBoxTransform();
	};
	CComboBoxController.prototype.recalculateTransform = function () {
		const oControl = this.control;
		this.dropButton.extY = oControl.extY;
		this.dropButton.extX = Math.min(oControl.extY, oControl.extX);
		const nXOffset = (oControl.extX - this.dropButton.extX);
		this.dropButton.x = oControl.x + nXOffset;
		this.dropButton.y = oControl.y;
		this.dropButton.transform = oControl.transform.CreateDublicate();
		this.dropButton.transform.Translate(nXOffset, 0);
		this.dropButton.invertTransform = this.dropButton.transform.CreateDublicate().Invert();
		this.recalculateListBoxTransform();
	};
	CComboBoxController.prototype.setupComboBoxBehavior = function () {
		const oThis = this;


		this.listBox.isMultiSelection = function () {
			return false;
		};
		
		this.listBox.isExtendedSelection = function () {
			return false;
		};
		
		this.listBox.isShowScroll = function () {
			const oFormControlPr = oThis.getFormControlPr();
			const nDropLines = oFormControlPr.dropLines || 8;
			return oThis.listBox.listItems.length > nDropLines;
		};
		
		this.listBox.onChangeSelection = function (oDrawingController) {
			const nIndex = this.getSingleSelectedIndex();
			if (nIndex === -1) {
				return;
			}
			oDrawingController.checkObjectsAndCallback(function () {
				oThis.selectedText = oThis.listBox.listItems[nIndex].text;
				oThis.updateLinkedCell();
				oThis.updateSelectionInFormControl();
				oThis.closeDropdown(oDrawingController);
			}, [], undefined, AscDFH.historydescription_Spreadsheet_SelectListBox, [], true);
		};
	};
	CComboBoxController.prototype.updateSelectionInFormControl = function () {
		const oFormControlPr = this.getFormControlPr();
		const nIndex = this.listBox.getSingleSelectedIndex() + 1;
		oFormControlPr.setSel(nIndex ? nIndex : null);
	};
	CComboBoxController.prototype.updateListItems = function () {
		this.listBox.updateListItems();
	};

	CComboBoxController.prototype.updateSelectedIndex = function () {
		this.listBox.resetSelectedIndices();
		const nFmlaLinkIndex = this.getFmlaLinkIndex();
		if (nFmlaLinkIndex !== null) {
			if (nFmlaLinkIndex >= 0) {
				const nIndex = this.listBox.addSelectedIndex(nFmlaLinkIndex);
				if (nIndex !== null) {
					this.selectedText = this.listBox.listItems[nIndex].text;
				}
			}
		} else {
			const oFormControlPr = this.getFormControlPr();
			const nSelectedIndex = AscFormat.isRealNumber(oFormControlPr.sel) ? oFormControlPr.sel - 1 : -1;
			if (nSelectedIndex >= 0 && this.listBox.listItems[nSelectedIndex]) {
				const nIndex = this.listBox.addSelectedIndex(nSelectedIndex);
				if (nIndex !== null) {
					this.selectedText = this.listBox.listItems[nIndex].text;
				}
			} else {
				this.selectedText = null;
			}
		}
	};

	CComboBoxController.prototype.openDropdown = function (oDrawingController) {
		if (!this.isDropdownOpen) {
			this.isDropdownOpen = true;
			this.dropButton.setIsHold(true);
			oDrawingController.addDropDown(this);
			this.checkNeedUpdate();
		}
	};

	CComboBoxController.prototype.closeDropdown = function (oDrawingController) {
		if (this.isDropdownOpen) {
			this.dropButton.setIsHold(false);
			this.isDropdownOpen = false;
			oDrawingController.resetDropDowns();
			this.checkNeedUpdate();
		}
	};
	CComboBoxController.prototype.drawDropDown = function(graphics) {
		if (this.isDropdownOpen) {
			this.listBox.draw(graphics);
		}
	};
	CComboBoxController.prototype.draw = function (graphics, transform, transformText, pageIndex, opt) {
		const oControl = this.control;
		const oTransform = transform || oControl.transform;
		
		graphics.SaveGrState();
		graphics.transform3(oTransform);
		startRoundControl(graphics, 0, 0, oControl.extX, oControl.extY, 2, getFlatPenColor());
		const nLabelWidth = oControl.extX - this.dropButton.extX;
		graphics.b_color1(255, 255, 255, 255);
		graphics.p_width(0);
		graphics._s();
		graphics._m(0, 0);
		graphics._l(nLabelWidth, 0);
		graphics._l(nLabelWidth, oControl.extY);
		graphics._l(0, oControl.extY);
		graphics._z();
		graphics.df();
		graphics._e();
		if (this.selectedText) {
			const nTextX = LISTBOX_TEXT_PADDING;
			const nTextY = (this.control.extY - getListBoxItemFontHeight()) / 2;
			const transform = oTransform.CreateDublicate();
			transform.Translate(nTextX, nTextY);
			graphics.transform3(transform);
			drawParagraph(this.selectedText, graphics);
		}

		this.dropButton.draw(graphics);
		endRoundControl(graphics);
		graphics.RestoreGrState();
	};

	CComboBoxController.prototype.recalculateListBoxTransform = function () {
		const oControl = this.control;
		const oFormControlPr = this.getFormControlPr();
		const nDropLines = oFormControlPr.dropLines || 8;

		const nItemHeight = LISTBOX_MAX_ITEM_HEIGHT;
		const nActualItems = this.listBox.listItems.length;
		const nVisibleItems = Math.min(nDropLines, nActualItems);
		const nDropdownHeight = nVisibleItems * nItemHeight;
		const nYOffset = 1 + oControl.extY;
		this.listBox.x = oControl.x;
		this.listBox.y = oControl.y + nYOffset;
		this.listBox.extX = oControl.extX;
		this.listBox.extY = nDropdownHeight;
		const oDropdownTransform = oControl.transform.CreateDublicate();
		global_MatrixTransformer.TranslateAppend(oDropdownTransform, 0, nYOffset);
		this.listBox.transform = oDropdownTransform;
		this.listBox.invertTransform = global_MatrixTransformer.Invert(oDropdownTransform);
		this.listBox.recalculateTransform();
	};

	CComboBoxController.prototype.getCursorInfo = function (e, nX, nY) {
		const oControl = this.control;
		if (!this.hitInInnerArea(nX, nY)) {
			return null;
		}
		if (oControl.selected) {
			return {cursorType: "default", objectId: oControl.GetId()};
		}
		return {cursorType: "pointer", objectId: oControl.GetId()};
	};

	CComboBoxController.prototype.onMouseDown = function (e, nX, nY, nPageIndex, oDrawingController) {
		const oControl = this.control;
		if (oControl.selected) {
			return true;
		}
		if (e.button !== 0) {
			this.closeDropdown(oDrawingController);
			return false;
		}

		if (this.isDropdownOpen) {
			if (!this.listBox.onMouseDown(e, nX, nY, nPageIndex, oDrawingController)) {
				this.closeDropdown(oDrawingController);
			}
			return oControl.hit(nX, nY);
		} else if (oControl.hit(nX, nY)) {
			this.openDropdown(oDrawingController);
			return true;
		}
		return false;
	};

	CComboBoxController.prototype.onMouseUp = function (e, nX, nY, nPageIndex, oDrawingController) {
		if (this.isDropdownOpen) {
			return this.listBox.onMouseUp(e, nX, nY, nPageIndex, oDrawingController);
		}
		return false;
	};

	CComboBoxController.prototype.onMouseMove = function (e, nX, nY, nPageIndex, oDrawingController) {
		let bRet = this.hitInComboBox(nX, nY);
		this.dropButton.setIsHover(bRet);
		if (this.isDropdownOpen) {
			bRet |= this.listBox.onMouseMove(e, nX, nY, nPageIndex, oDrawingController);
		}
		return !!bRet;
	};

	CComboBoxController.prototype.updateLinkedCell = function () {
		const nIndex = this.listBox.getSingleSelectedIndex() + 1;
		if (!nIndex) {
			return;
		}

		const oRef = this.getParsedFmlaLink();
		if (oRef) {
			const oCellValue = new AscCommonExcel.CCellValue();
			oCellValue.type = AscCommon.CellValueType.Number;
			oCellValue.number = nIndex;
			this.setRangeValue(oRef, oCellValue);
		}
	};
	CComboBoxController.prototype.onUpdate = function () {
		const oRect = new AscFormat.CGraphicBounds(this.control.x - 1, this.control.y - 1, this.control.x + this.control.extX + 1, this.listBox.y + this.listBox.extY + 1);
		return AscFormat.CShape.prototype.onUpdate.call(this.control, oRect);
	};
	CComboBoxController.prototype.hitInInnerArea = function (nX, nY) {
		if (this.isDropdownOpen) {
			const oControl = this.control;
			return AscFormat.HitToRect(nX, nY, oControl.invertTransform, 0, 0, oControl.extX, oControl.extY + this.listBox.extY + 1);
		} else {
			return this.hitInComboBox(nX, nY);
		}
	};
	CComboBoxController.prototype.hitInComboBox = function (nX, nY) {
		const oControl = this.control;
		return AscFormat.HitToRect(nX, nY, oControl.invertTransform, 0, 0, oControl.extX, oControl.extY);
	};
	CComboBoxController.prototype.isNeedResetState = function () {
		return !this.isDropdownOpen;
	};
	CComboBoxController.prototype.handleFmlaRange = function (aRanges) {
		return this.handleRef(aRanges, this.getParsedFmlaRange(), this.recalculateItems.bind(this));
	};
	CComboBoxController.prototype.handleFmlaLink = function (aRanges) {
		return this.handleRef(aRanges, this.getParsedFmlaLink(), this.updateSelectedIndex.bind(this));
	};
	CComboBoxController.prototype.getFmlaLinkIndex = function () {
		const oParsedLink = this.getParsedFmlaLink();
		if (oParsedLink) {
			let nIndex = null;
			oParsedLink._foreach(function (oCell) {
				if (oCell) {
					const nNumber = oCell.getNumberValue();
					if (nNumber !== null) {
						nIndex = nNumber - 1;
					}
				}
			});
			return nIndex;
		}
		return null;
	};
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AltText] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AutoFill] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AutoLine] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_AutoPict] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Dde] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_DefaultSize] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Disabled] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Cf] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_LinkedCell] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_ListFillRange] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_RId] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Locked] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Macro] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_Print] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_RecalcAlways] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_ControlPr_UiObject] = AscDFH.CChangesDrawingsBool;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AltText] = function (oClass, value) {
		this.altText = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AutoFill] = function (oClass, value) {
		this.autoFill = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AutoLine] = function (oClass, value) {
		this.autoLine = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_AutoPict] = function (oClass, value) {
		this.autoPict = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Dde] = function (oClass, value) {
		this.dde = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_DefaultSize] = function (oClass, value) {
		this.defaultSize = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Disabled] = function (oClass, value) {
		this.disabled = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Cf] = function (oClass, value) {
		this.cf = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_LinkedCell] = function (oClass, value) {
		this.linkedCell = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_ListFillRange] = function (oClass, value) {
		this.listFillRange = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_RId] = function (oClass, value) {
		this.rId = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Locked] = function (oClass, value) {
		this.locked = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Macro] = function (oClass, value) {
		this.macro = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_Print] = function (oClass, value) {
		this.print = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_RecalcAlways] = function (oClass, value) {
		this.recalcAlways = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_ControlPr_UiObject] = function (oClass, value) {
		this.uiObject = value;
	};

	function CControlPr() {
		AscFormat.CBaseFormatObject.call(this);
		this.altText = null;
		this.autoFill = null;
		this.autoLine = null;
		this.autoPict = null;
		this.dde = null;
		this.defaultSize = null;
		this.disabled = null;
		this.cf = null;
		this.linkedCell = null;
		this.listFillRange = null;
		this.rId = null;
		this.locked = null;
		this.macro = null;
		this.print = null;
		this.recalcAlways = null;
		this.uiObject = null;
	}

	AscFormat.InitClass(CControlPr, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_ControlPr);
	CControlPr.prototype.setAltText = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_AltText, this.altText, pr));
		this.altText = pr;
	}
	CControlPr.prototype.getAltText = function () {
		return this.altText;
	};
	CControlPr.prototype.setAutoFill = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_AutoFill, this.autoFill, pr));
		this.autoFill = pr;
	}
	CControlPr.prototype.getAutoFill = function () {
		return this.autoFill;
	};
	CControlPr.prototype.setAutoLine = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_AutoLine, this.autoLine, pr));
		this.autoLine = pr;
	}
	CControlPr.prototype.getAutoLine = function () {
		return this.autoLine;
	};
	CControlPr.prototype.setAutoPict = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_AutoPict, this.autoPict, pr));
		this.autoPict = pr;
	}
	CControlPr.prototype.getAutoPict = function () {
		return this.autoPict;
	};
	CControlPr.prototype.setDde = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Dde, this.dde, pr));
		this.dde = pr;
	}
	CControlPr.prototype.getDde = function () {
		return this.dde;
	};
	CControlPr.prototype.setDefaultSize = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_DefaultSize, this.defaultSize, pr));
		this.defaultSize = pr;
	}
	CControlPr.prototype.getDefaultSize = function () {
		return this.defaultSize;
	};
	CControlPr.prototype.setDisabled = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Disabled, this.disabled, pr));
		this.disabled = pr;
	}
	CControlPr.prototype.getDisabled = function () {
		return this.disabled;
	};
	CControlPr.prototype.setCf = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_Cf, this.cf, pr));
		this.cf = pr;
	}
	CControlPr.prototype.getCf = function () {
		return this.cf;
	};
	CControlPr.prototype.setLinkedCell = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_LinkedCell, this.linkedCell, pr));
		this.linkedCell = pr;
	}
	CControlPr.prototype.getLinkedCell = function () {
		return this.linkedCell;
	};
	CControlPr.prototype.setListFillRange = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_ListFillRange, this.listFillRange, pr));
		this.listFillRange = pr;
	}
	CControlPr.prototype.getListFillRange = function () {
		return this.listFillRange;
	};
	CControlPr.prototype.setRId = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_ControlPr_RId, this.rId, pr));
		this.rId = pr;
	}
	CControlPr.prototype.getRId = function () {
		return this.rId;
	};
	CControlPr.prototype.setLocked = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Locked, this.locked, pr));
		this.locked = pr;
	}
	CControlPr.prototype.getLocked = function () {
		return this.locked;
	};
	CControlPr.prototype.setMacro = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ControlPr_Macro, this.macro, pr));
		this.macro = pr;
	}
	CControlPr.prototype.getMacro = function () {
		return this.macro;
	};
	CControlPr.prototype.setPrint = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_Print, this.print, pr));
		this.print = pr;
	}
	CControlPr.prototype.getPrint = function () {
		return this.print;
	};
	CControlPr.prototype.setRecalcAlways = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_RecalcAlways, this.recalcAlways, pr));
		this.recalcAlways = pr;
	}
	CControlPr.prototype.getRecalcAlways = function () {
		return this.recalcAlways;
	};
	CControlPr.prototype.setUiObject = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ControlPr_UiObject, this.uiObject, pr));
		this.uiObject = pr;
	};
	CControlPr.prototype.getUiObject = function () {
		return this.uiObject;
	};
	CControlPr.prototype.fillObject = function (oCopy, oPr) {
		oCopy.setAltText(this.altText);
		oCopy.setAutoFill(this.autoFill);
		oCopy.setAutoLine(this.autoLine);
		oCopy.setAutoPict(this.autoPict);
		oCopy.setDde(this.dde);
		oCopy.setDefaultSize(this.defaultSize);
		oCopy.setDisabled(this.disabled);
		oCopy.setCf(this.cf);
		oCopy.setLinkedCell(this.linkedCell);
		oCopy.setListFillRange(this.listFillRange);
		oCopy.setRId(this.rId);
		oCopy.setLocked(this.locked);
		oCopy.setMacro(this.macro);
		oCopy.setPrint(this.print);
		oCopy.setRecalcAlways(this.recalcAlways);
		oCopy.setUiObject(this.uiObject);
	};
	CControlPr.prototype.getMacroId = function () {
		const sMacro = this.macro;
		return typeof sMacro === "string" ? sMacro : null;
	};
	CControlPr.prototype.getJSAMacroId = function () {
		const sMacro = this.getMacroId();
		if (sMacro !== null && sMacro.indexOf(AscFormat.MACRO_PREFIX) === 0) {
			return sMacro.slice(AscFormat.MACRO_PREFIX.length);
		}
		return null;
	};

	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_DropLines] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_ObjectType] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Checked] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_DropStyle] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Dx] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Inc] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Min] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Max] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Page] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Sel] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_SelType] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_TextHAlign] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_TextVAlign] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Val] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_WidthMin] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_EditVal] = AscDFH.CChangesDrawingsLong;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaGroup] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaLink] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaRange] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FmlaTxbx] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Colored] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_FirstButton] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_Horiz] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_JustLastX] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_LockText] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_MultiSel] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_NoThreeD] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_NoThreeD2] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_MultiLine] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_VerticalBar] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_PasswordEdit] = AscDFH.CChangesDrawingsBool;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_AddItemToLst] = AscDFH.CChangesDrawingsContentString;
	AscDFH.changesFactory[AscDFH.historyitem_FormControlPr_RemoveItemFromLst] = AscDFH.CChangesDrawingsContentString;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_DropLines] = function (oClass, value) {
		oClass.dropLines = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_ObjectType] = function (oClass, value) {
		oClass.objectType = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Checked] = function (oClass, value) {
		oClass.checked = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_DropStyle] = function (oClass, value) {
		oClass.dropStyle = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Dx] = function (oClass, value) {
		oClass.dx = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Inc] = function (oClass, value) {
		oClass.inc = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Min] = function (oClass, value) {
		oClass.min = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Max] = function (oClass, value) {
		oClass.max = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Page] = function (oClass, value) {
		oClass.page = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Sel] = function (oClass, value) {
		if(oClass.parent) {
			oClass.parent.controller.recalcInfo.recalculateItems = true;
			oClass.parent.addToRecalculate();
		}
		oClass.sel = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_SelType] = function (oClass, value) {
		oClass.selType = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_TextHAlign] = function (oClass, value) {
		oClass.textHAlign = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_TextVAlign] = function (oClass, value) {
		oClass.textVAlign = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Val] = function (oClass, value) {
			if(oClass.parent) {
				oClass.parent.recalcInfo.recalculateTransform = true;
				oClass.parent.addToRecalculate();
			}
		oClass.val = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_WidthMin] = function (oClass, value) {
		oClass.widthMin = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_EditVal] = function (oClass, value) {
		oClass.editVal = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaGroup] = function (oClass, value) {
		oClass.fmlaGroup = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaLink] = function (oClass, value) {
		oClass.fmlaLink = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaRange] = function (oClass, value) {
		oClass.fmlaRange = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FmlaTxbx] = function (oClass, value) {
		oClass.fmlaTxbx = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Colored] = function (oClass, value) {
		oClass.colored = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_FirstButton] = function (oClass, value) {
		oClass.firstButton = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_Horiz] = function (oClass, value) {
		oClass.horiz = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_JustLastX] = function (oClass, value) {
		oClass.justLastX = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_LockText] = function (oClass, value) {
		oClass.lockText = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_MultiSel] = function (oClass, value) {
		if(oClass.parent) {
			oClass.parent.recalcInfo.recalculateTransform = true;
			oClass.parent.addToRecalculate();
		}
		oClass.multiSel = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_NoThreeD] = function (oClass, value) {
		oClass.noThreeD = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_NoThreeD2] = function (oClass, value) {
		oClass.noThreeD2 = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_MultiLine] = function (oClass, value) {
		oClass.multiLine = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_VerticalBar] = function (oClass, value) {
		oClass.verticalBar = value;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_FormControlPr_PasswordEdit] = function (oClass, value) {
		oClass.passwordEdit = value;
	};
	AscDFH.drawingContentChanges[AscDFH.historyitem_FormControlPr_AddItemToLst] = function (oClass) {
		return oClass.itemLst;
	};
	AscDFH.drawingContentChanges[AscDFH.historyitem_FormControlPr_RemoveItemFromLst] = function (oClass) {
		return oClass.itemLst;
	};

	function CFormControlPr() {
		AscFormat.CBaseFormatObject.call(this);
		this.dropLines = null;
		this.objectType = null;
		this.checked = null;
		this.dropStyle = null;
		this.dx = null;
		this.inc = null;
		this.min = null;
		this.max = null;
		this.page = null;
		this.sel = null;
		this.selType = null;
		this.textHAlign = null;
		this.textVAlign = null;
		this.val = null;
		this.widthMin = null;
		this.editVal = null;
		this.fmlaGroup = null;
		this.fmlaLink = null;
		this.fmlaRange = null;
		this.fmlaTxbx = null;
		this.colored = null;
		this.firstButton = null;
		this.horiz = null;
		this.justLastX = null;
		this.lockText = null;
		this.multiSel = null;
		this.noThreeD = null;
		this.noThreeD2 = null;
		this.multiLine = null;
		this.verticalBar = null;
		this.passwordEdit = null;
		this.itemLst = [];

		this.isExternalFmlaLink = false;
	}

	AscFormat.InitClass(CFormControlPr, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_FormControlPr);
	CFormControlPr.prototype.setDropLines = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_DropLines, this.dropLines, pr));
		this.dropLines = pr;
	}
	CFormControlPr.prototype.getDropLines = function () {
		return this.dropLines;
	}
	CFormControlPr.prototype.setObjectType = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_ObjectType, this.objectType, pr));
		this.objectType = pr;
	}
	CFormControlPr.prototype.getObjectType = function () {
		return this.objectType;
	}
	CFormControlPr.prototype.setChecked = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Checked, this.checked, pr));
		this.checked = pr;
	}
	CFormControlPr.prototype.getChecked = function () {
		return this.checked;
	}
	CFormControlPr.prototype.setDropStyle = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_DropStyle, this.dropStyle, pr));
		this.dropStyle = pr;
	}
	CFormControlPr.prototype.getDropStyle = function () {
		return this.dropStyle;
	}
	CFormControlPr.prototype.setDx = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Dx, this.dx, pr));
		this.dx = pr;
	}
	CFormControlPr.prototype.getDx = function () {
		return this.dx;
	}
	CFormControlPr.prototype.setInc = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Inc, this.inc, pr));
		this.inc = pr;
	}
	CFormControlPr.prototype.getInc = function () {
		return this.inc;
	}
	CFormControlPr.prototype.setMin = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Min, this.min, pr));
		this.min = pr;
	}
	CFormControlPr.prototype.getMin = function () {
		return this.min;
	}
	CFormControlPr.prototype.setMax = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Max, this.max, pr));
		this.max = pr;
	}
	CFormControlPr.prototype.getMax = function () {
		return this.max;
	}
	CFormControlPr.prototype.setPage = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Page, this.page, pr));
		this.page = pr;
	}
	CFormControlPr.prototype.getPage = function () {
		return this.page;
	}
	CFormControlPr.prototype.setSel = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Sel, this.sel, pr));
		this.sel = pr;
	}
	CFormControlPr.prototype.getSel = function () {
		return this.sel;
	}
	CFormControlPr.prototype.setSelType = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_SelType, this.selType, pr));
		this.selType = pr;
	}
	CFormControlPr.prototype.getSelType = function () {
		return this.selType;
	}
	CFormControlPr.prototype.setTextHAlign = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_TextHAlign, this.textHAlign, pr));
		this.textHAlign = pr;
	}
	CFormControlPr.prototype.getTextHAlign = function () {
		return this.textHAlign;
	}
	CFormControlPr.prototype.setTextVAlign = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_TextVAlign, this.textVAlign, pr));
		this.textVAlign = pr;
	}
	CFormControlPr.prototype.getTextVAlign = function () {
		return this.textVAlign;
	}
	CFormControlPr.prototype.setVal = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_Val, this.val, pr));
		this.val = pr;
	}
	CFormControlPr.prototype.getVal = function () {
		return this.val;
	}
	CFormControlPr.prototype.setWidthMin = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_WidthMin, this.widthMin, pr));
		this.widthMin = pr;
	}
	CFormControlPr.prototype.getWidthMin = function () {
		return this.widthMin;
	}
	CFormControlPr.prototype.setEditVal = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_FormControlPr_EditVal, this.editVal, pr));
		this.editVal = pr;
	}
	CFormControlPr.prototype.getEditVal = function () {
		return this.editVal;
	}
	CFormControlPr.prototype.setFmlaGroup = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaGroup, this.fmlaGroup, pr));
		this.fmlaGroup = pr;
	}
	CFormControlPr.prototype.getFmlaGroup = function () {
		return this.fmlaGroup;
	}
	CFormControlPr.prototype.setFmlaLink = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaLink, this.fmlaLink, pr));
		this.fmlaLink = pr;
	}
	CFormControlPr.prototype.getFmlaLink = function () {
		return this.fmlaLink;
	}
	CFormControlPr.prototype.setFmlaRange = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaRange, this.fmlaRange, pr));
		this.fmlaRange = pr;
	}
	CFormControlPr.prototype.getFmlaRange = function () {
		return this.fmlaRange;
	}
	CFormControlPr.prototype.setFmlaTxbx = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_FmlaTxbx, this.fmlaTxbx, pr));
		this.fmlaTxbx = pr;
	}
	CFormControlPr.prototype.getFmlaTxbx = function () {
		return this.fmlaTxbx;
	}
	CFormControlPr.prototype.setColored = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_Colored, this.colored, pr));
		this.colored = pr;
	}
	CFormControlPr.prototype.getColored = function () {
		return this.colored;
	}
	CFormControlPr.prototype.setFirstButton = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_FirstButton, this.firstButton, pr));
		this.firstButton = pr;
	}
	CFormControlPr.prototype.getFirstButton = function () {
		return this.firstButton;
	}
	CFormControlPr.prototype.setHoriz = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_Horiz, this.horiz, pr));
		this.horiz = pr;
	}
	CFormControlPr.prototype.getHoriz = function () {
		return this.horiz;
	}
	CFormControlPr.prototype.setJustLastX = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_JustLastX, this.justLastX, pr));
		this.justLastX = pr;
	}
	CFormControlPr.prototype.getJustLastX = function () {
		return this.justLastX;
	}
	CFormControlPr.prototype.setLockText = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_LockText, this.lockText, pr));
		this.lockText = pr;
	}
	CFormControlPr.prototype.getLockText = function () {
		return this.lockText;
	}
	CFormControlPr.prototype.setMultiSel = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_FormControlPr_MultiSel, this.multiSel, pr));
		this.multiSel = pr;
	}
	CFormControlPr.prototype.getMultiSel = function () {
		return this.multiSel;
	}
	CFormControlPr.prototype.setNoThreeD = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_NoThreeD, this.noThreeD, pr));
		this.noThreeD = pr;
	}
	CFormControlPr.prototype.getNoThreeD = function () {
		return this.noThreeD;
	}
	CFormControlPr.prototype.setNoThreeD2 = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_NoThreeD2, this.noThreeD2, pr));
		this.noThreeD2 = pr;
	}
	CFormControlPr.prototype.getNoThreeD2 = function () {
		return this.noThreeD2;
	}
	CFormControlPr.prototype.setMultiLine = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_MultiLine, this.multiLine, pr));
		this.multiLine = pr;
	}
	CFormControlPr.prototype.getMultiLine = function () {
		return this.multiLine;
	}
	CFormControlPr.prototype.setVerticalBar = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_VerticalBar, this.verticalBar, pr));
		this.verticalBar = pr;
	}
	CFormControlPr.prototype.getVerticalBar = function () {
		return this.verticalBar;
	}
	CFormControlPr.prototype.setPasswordEdit = function (pr) {
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_FormControlPr_PasswordEdit, this.passwordEdit, pr));
		this.passwordEdit = pr;
	}
	CFormControlPr.prototype.getPasswordEdit = function () {
		return this.passwordEdit;
	}
	CFormControlPr.prototype.addItemToLst = function (nIdx, sPr) {
		var nInsertIdx = Math.min(this.itemLst.length, Math.max(0, nIdx));
		AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsContentString(this, AscDFH.historyitem_FormControlPr_AddItemToLst, nInsertIdx, [sPr], true));
		this.itemLst.splice(nInsertIdx, 0, sPr);
	};
	CFormControlPr.prototype.removeItemFromLst = function (nIdx) {
		if (nIdx > -1 && nIdx < this.itemLst.length) {
			AscCommon.History.CanAddChanges() && AscCommon.History.Add(new AscDFH.CChangesDrawingsContentString(this, AscDFH.historyitem_CCommonDataListRemove, nIdx, [this.itemLst[nIdx]], false));
			this.itemLst.splice(nIdx, 1);
		}
	};
	CFormControlPr.prototype.fillObject = function (oCopy, oPr) {
		oCopy.setDropLines(this.dropLines);
		oCopy.setObjectType(this.objectType);
		oCopy.setChecked(this.checked);
		oCopy.setDropStyle(this.dropStyle);
		oCopy.setDx(this.dx);
		oCopy.setInc(this.inc);
		oCopy.setMin(this.min);
		oCopy.setMax(this.max);
		oCopy.setPage(this.page);
		oCopy.setSel(this.sel);
		oCopy.setSelType(this.selType);
		oCopy.setTextHAlign(this.textHAlign);
		oCopy.setTextVAlign(this.textVAlign);
		oCopy.setVal(this.val);
		oCopy.setWidthMin(this.widthMin);
		oCopy.setEditVal(this.editVal);
		oCopy.setFmlaGroup(this.fmlaGroup);
		oCopy.setFmlaLink(this.fmlaLink);
		oCopy.setFmlaRange(this.fmlaRange);
		oCopy.setFmlaTxbx(this.fmlaTxbx);
		oCopy.setColored(this.colored);
		oCopy.setFirstButton(this.firstButton);
		oCopy.setHoriz(this.horiz);
		oCopy.setJustLastX(this.justLastX);
		oCopy.setLockText(this.lockText);
		oCopy.setMultiSel(this.multiSel);
		oCopy.setNoThreeD(this.noThreeD);
		oCopy.setNoThreeD2(this.noThreeD2);
		oCopy.setMultiLine(this.multiLine);
		oCopy.setVerticalBar(this.verticalBar);
		oCopy.setPasswordEdit(this.passwordEdit);
		for (let i = 0; i < this.itemLst.length; i += 1) {
			oCopy.addItemToLst(oCopy.itemLst.length, this.itemLst[i]);
		}
	};

	window["AscFormat"] = window["AscFormat"] || {};
	window["AscFormat"].CControl = CControl;
	window["AscFormat"].CFormControlPr_checked_unchecked = CFormControlPr_checked_unchecked;
	window["AscFormat"].CFormControlPr_checked_checked = CFormControlPr_checked_checked;
	window["AscFormat"].CFormControlPr_checked_mixed = CFormControlPr_checked_mixed;
	window["AscFormat"].getVerticalControlAlignFromBodyPr = getVerticalControlAlignFromBodyPr;
	window["AscFormat"].getHorizontalAlignFromControl = getHorizontalAlignFromControl;
})();
