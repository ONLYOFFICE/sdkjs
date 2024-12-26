/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	var InitClass = AscFormat.InitClass;
	var CAnimTexture = AscFormat.CAnimTexture;


	const STATE_FLAG_SELECTED = 1;
	const STATE_FLAG_HOVERED = 2;
	const STATE_FLAG_PRESSED = 4;
	const STATE_FLAG_DISABLED = 8;

	const CONTROL_TYPE_UNKNOWN = 0;
	const CONTROL_TYPE_LABEL = 1;
	const CONTROL_TYPE_IMAGE = 2;
	const CONTROL_TYPE_BUTTON = 3;
	const CONTROL_TYPE_HEADER = 4;
	const CONTROL_TYPE_SCROLL_VERT = 5;
	const CONTROL_TYPE_SCROLL_HOR = 6;
	const CONTROL_TYPE_TIMELINE_CONTAINER = 7;
	const CONTROL_TYPE_TIMELINE = 8;
	const CONTROL_TYPE_SEQ_LIST_CONTAINER = 9;
	const CONTROL_TYPE_SEQ_LIST = 10;
	const CONTROL_TYPE_ANIM_SEQ = 11;
	const CONTROL_TYPE_ANIM_GROUP_LIST = 12;
	const CONTROL_TYPE_ANIM_GROUP = 13;
	const CONTROL_TYPE_ANIM_ITEM = 14;
	const CONTROL_TYPE_EFFECT_BAR = 15;


	function CControl(oParentControl) {
		AscFormat.ExecuteNoHistory(function () {
			AscFormat.CShape.call(this);
			this.setRecalculateInfo();
			this.setBDeleted(false);
			this.setLayout(0, 0, 0, 0);
		}, this, []);

		this.parent = editor.WordControl.m_oLogicDocument.Slides[0];
		this.parentControl = oParentControl;
		this.state = 0;
		this.hidden = false;
		this.previous = null;
		this.next = null;
	}

	InitClass(CControl, AscFormat.CShape, CONTROL_TYPE_UNKNOWN);

	CControl.prototype.DEFALT_WRAP_OBJECT = {
		oTxWarpStruct: null,
		oTxWarpStructParamarks: null,
		oTxWarpStructNoTransform: null,
		oTxWarpStructParamarksNoTransform: null
	};
	CControl.prototype.setHidden = function (bVal) {
		if (this.hidden !== bVal) {
			this.hidden = bVal;
			this.onUpdate();
		}
	};
	CControl.prototype.show = function () {
		this.setHidden(false);
	};
	CControl.prototype.hide = function () {
		this.setHidden(true);
	};
	CControl.prototype.isHidden = function () {
		return this.hidden;
	};
	CControl.prototype.notAllowedWithoutId = function () {
		return false;
	};
	//define shape methods
	CControl.prototype.getBodyPr = function () {
		return this.bodyPr;
	};
	CControl.prototype.getScrollOffsetX = function (oChild) {
		return 0;
	};
	CControl.prototype.getScrollOffsetY = function (oChild) {
		return 0;
	};
	CControl.prototype.getParentScrollOffsetX = function (oChild) {
		if (this.parentControl) {
			return this.parentControl.getScrollOffsetX(oChild);
		}
		return 0;
	};
	CControl.prototype.getParentScrollOffsetY = function (oChild) {
		if (this.parentControl) {
			return this.parentControl.getScrollOffsetY(oChild);
		}
		return 0;
	};
	CControl.prototype.getFullTransformMatrix = function () {
		return this.transform;
	};
	CControl.prototype.getInvFullTransformMatrix = function () {
		return this.invertTransform;
	};
	CControl.prototype.multiplyParentTransforms = function (oLocalTransform) {
		var oMT = AscCommon.global_MatrixTransformer;
		var oTransform = oMT.CreateDublicateM(oLocalTransform);
		var oScrollMatrix = new AscCommon.CMatrix();
		oScrollMatrix.tx = this.getParentScrollOffsetX(this);
		oScrollMatrix.ty = this.getParentScrollOffsetY(this);
		oMT.MultiplyAppend(oTransform, oScrollMatrix);
		var oParentTransform = this.parentControl && this.parentControl.getFullTransformMatrix();
		oParentTransform && oMT.MultiplyAppend(oTransform, oParentTransform);
		return oTransform;
	};
	CControl.prototype.getFullTransform = function () {
		return this.transform;
	};
	CControl.prototype.getFullTextTransform = function () {
		return this.transformText;
	};
	CControl.prototype.recalculate = function () {
		AscFormat.CShape.prototype.recalculate.call(this);
	};
	CControl.prototype.recalculateBrush = function () {
		this.brush = null;
	};
	CControl.prototype.recalculatePen = function () {
		this.pen = null;
	};
	CControl.prototype.recalculateContent = function () {
	};
	CControl.prototype.recalculateGeometry = function () {
		//this.calcGeometry = AscFormat.CreateGeometry("rect");
		//this.calcGeometry.Recalculate(this.extX, this.extY);
	};
	CControl.prototype.recalculateTransform = function () {
		if (!this.transform) {
			this.transform = new AscCommon.CMatrix();
		}
		var tx = this.getLeft();
		var ty = this.getTop();
		this.x = tx;
		this.y = ty;
		this.rot = 0;
		this.extX = this.getWidth();
		this.extY = this.getHeight();
		this.flipH = false;
		this.flipV = false;
		ty += this.getParentScrollOffsetY(this);
		var oCurParent = this.parentControl;

		if (oCurParent) {
			tx += oCurParent.transform.tx;
			ty += oCurParent.transform.ty
		}
		this.transform.tx = tx;
		this.transform.ty = ty;
		if (!this.invertTransform) {
			this.invertTransform = new AscCommon.CMatrix();
		}
		this.invertTransform.tx = -tx;
		this.invertTransform.ty = -ty;
		this.localTransform = this.transform;
	};
	CControl.prototype.recalculateTransformText = function () {
		if (!this.transformText) {
			this.transformText = new AscCommon.CMatrix();
		}
		this.transformText.tx = this.transform.tx;
		this.transformText.ty = this.transform.ty;

		if (!this.invertTransformText) {
			this.invertTransformText = new AscCommon.CMatrix();
		}
		this.invertTransformText.tx = -this.transform.tx;
		this.invertTransformText.ty = -this.transform.ty;
		this.localTransformText = this.transformText;
	};
	CControl.prototype.recalculateBounds = function () {
		var dX = this.transform.tx;
		var dY = this.transform.ty;
		this.bounds.reset(dX, dY, dX + this.getWidth(), dY + this.getHeight())
	};
	CControl.prototype.recalculateSnapArrays = function () {
	};
	CControl.prototype.checkAutofit = function (bIgnoreWordShape) {
		return false;
	};
	CControl.prototype.checkTextWarp = function (oContent, oBodyPr, dWidth, dHeight, bNeedNoTransform, bNeedWarp) {
		return this.DEFALT_WRAP_OBJECT;
	};
	CControl.prototype.addToRecalculate = function () {
	};
	CControl.prototype.canHandleEvents = function () {
		return true;
	};
	CControl.prototype.getPenWidth = function (graphics) {
		var fScale = graphics.m_oCoordTransform.sx;
		var nPenW = AscCommon.AscBrowser.convertToRetinaValue(1, true) / fScale;
		return nPenW;
	};
	CControl.prototype.draw = function (graphics) {
		if (this.isHidden()) { return false; }
		if (!this.checkUpdateRect(graphics.updatedRect)) { return false; }

		this.recalculateTransform();
		this.recalculateTransformText();

		const sFillColor = this.getFillColor();
		const sOutlineColor = this.getOutlineColor();
		let oColor;
		if (sOutlineColor || sFillColor) {
			graphics.SaveGrState();
			graphics.transform3(this.transform);
			var x = 0;
			var y = 0;
			var extX = this.getWidth();
			var extY = this.getHeight();
			if (sFillColor) {
				oColor = AscCommon.RgbaHexToRGBA(sFillColor);
				graphics.b_color1(oColor.R, oColor.G, oColor.B, 0xFF);
				graphics.rect(x, y, extX, extY);
				graphics.df();
			}
			if (sOutlineColor) {
				oColor = AscCommon.RgbaHexToRGBA(sOutlineColor);
				graphics.SetIntegerGrid(true);

				var nPenW = this.getPenWidth(graphics);
				//graphics.p_width(100);//AscCommon.AscBrowser.convertToRetinaValue(1, true);
				graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
				graphics.drawHorLine(0, y, x, x + extX, nPenW);
				graphics.drawHorLine(0, y + extY, x, x + extX, nPenW);
				graphics.drawVerLine(2, x, y, y + extY, nPenW);
				graphics.drawVerLine(2, x + extX, y, y + extY, nPenW);
				graphics.ds();
			}
			graphics.RestoreGrState();
		}
		AscFormat.CShape.prototype.draw.call(this, graphics);
		return true;
	};
	CControl.prototype.hit = function (x, y) {
		if (this.parentControl && !this.parentControl.hit(x, y)) {
			return false;
		}
		var oInv = this.invertTransform;
		var tx = oInv.TransformPointX(x, y);
		var ty = oInv.TransformPointY(x, y);
		return tx >= 0 && tx <= this.extX && ty >= 0 && ty <= this.extY;
	};
	CControl.prototype.setStateFlag = function (nFlag, bValue) {
		var nOldState = this.state;
		if (bValue) {
			this.state |= nFlag;
		} else {
			this.state &= (~nFlag);
		}
		if (nOldState !== this.state) {
			this.onUpdate();
		}
	};
	CControl.prototype.getStateFlag = function (nFlag) {
		return (this.state & nFlag) !== 0;
	};
	CControl.prototype.isHovered = function () {
		return this.getStateFlag(STATE_FLAG_HOVERED);
	};
	CControl.prototype.isActive = function () {
		if (this.parentControl) {
			if (!this.eventListener && this.parentControl.isEventListener(this)) {
				return true;
			}
		}
		return false;
	};
	CControl.prototype.setHoverState = function () {
		this.setStateFlag(STATE_FLAG_HOVERED, true);
	};
	CControl.prototype.setNotHoverState = function () {
		this.setStateFlag(STATE_FLAG_HOVERED, false);
	};
	CControl.prototype.onMouseMove = function (e, x, y) {
		if (e.IsLocked) {
			return false;
		}
		if (!this.canHandleEvents()) {
			return false;
		}
		const bHover = !!this.hit(x, y);
		const bHoverChanged = bHover !== this.isHovered();
		if(bHoverChanged) {
			this.setStateFlag(STATE_FLAG_HOVERED, bHover);
		}
		return false;
	};
	CControl.prototype.onMouseDown = function (e, x, y) {
		if (!this.canHandleEvents()) {
			return false;
		}
		if (this.hit(x, y)) {
			if (this.parentControl) {
				this.parentControl.setEventListener(this);
			}
			return true;
		}
		return false;
	};
	CControl.prototype.onMouseUp = function (e, x, y) {
		if (this.parentControl) {
			if (this.parentControl.isEventListener(this)) {
				this.parentControl.setEventListener(null);
				return true;
			}
		}
		return false;
	};
	CControl.prototype.onMouseWheel = function (e, deltaY, X, Y) {
		return false;
	};
	CControl.prototype.onUpdate = function () {
		if (this.parentControl) {
			var oBounds = this.getBounds();
			this.parentControl.onChildUpdate(oBounds);
		}
	};
	CControl.prototype.onChildUpdate = function (oBounds) {
		if (this.parentControl) {
			this.parentControl.onChildUpdate(oBounds);
		}
	};
	CControl.prototype.checkUpdateRect = function (oUpdateRect) {
		var oBounds = this.getBounds();
		if (oUpdateRect && oBounds) {
			if (!oUpdateRect.isIntersectOther(oBounds)) {
				return false;
			}
		}
		return true;
	};
	CControl.prototype.recalculate = function () {
		AscFormat.CShape.prototype.recalculate.call(this);
	};
	/**
	 * Sets the location and dimensions of the control inside the parent container.
	 *
	 * @param {number} dX - Offset of the element along the X axis relative to the upper-left corner of the parent container.
	 * @param {number} dY - Offset of the element along the Y axis relative to the upper-left corner of the parent container.
	 * @param {number} dExtX - Width of the element.
	 * @param {number} dExtY - Height of the element.
	 *
	 * @note
	 * - Negative values for dX and dY are supported with behavior similar to "overflow: hidden" in CSS.
	 * - Negative values for dExtX and dExtY are not supported and may lead to unexpected behavior.
	 * - It is recommended to avoid using negative values for dExtX and dExtY to ensure proper rendering and hit detection.
	 */
	CControl.prototype.setLayout = function (dX, dY, dExtX, dExtY) {
		if (!this.spPr) {
			this.spPr = new AscFormat.CSpPr();
		}
		if (!this.spPr.xfrm) {
			this.spPr.xfrm = new AscFormat.CXfrm();
		}

		this.spPr.xfrm.offX = dX;
		this.spPr.xfrm.offY = dY;
		this.spPr.xfrm.extX = dExtX;
		this.spPr.xfrm.extY = dExtY;
		this.handleUpdateExtents();
	};
	CControl.prototype.getLeft = function () {
		return this.spPr.xfrm.offX;
	};
	CControl.prototype.getTop = function () {
		return this.spPr.xfrm.offY;
	};
	CControl.prototype.getRight = function () {
		return this.spPr.xfrm.offX + this.spPr.xfrm.extX;
	};
	CControl.prototype.getBottom = function () {
		return this.spPr.xfrm.offY + this.spPr.xfrm.extY;
	};
	CControl.prototype.getWidth = function () {
		return this.spPr.xfrm.extX;
	};
	CControl.prototype.getHeight = function () {
		return this.spPr.xfrm.extY;
	};
	CControl.prototype.getBounds = function () {
		this.recalculateBounds();
		this.recalculateTransform();
		this.recalculateTransformText();
		return this.bounds;
	};
	CControl.prototype.convertRelToAbs = function (oPos) {
		var oAbsPos = { x: oPos.x, y: oPos.y };
		var oParent = this;
		while (oParent) {
			oAbsPos.x += oParent.getLeft();
			oAbsPos.y += oParent.getTop();
			oParent = oParent.parentControl;
		}
		return oAbsPos;
	};
	CControl.prototype.convertAbsToRel = function (oPos) {
		var oRelPos = { x: oPos.x, y: oPos.y };
		var oParent = this;
		while (oParent) {
			oRelPos.x -= oParent.getLeft();
			oRelPos.y -= oParent.getTop();
			oParent = oParent.parentControl;
		}
		return oRelPos;
	};
	CControl.prototype.getNext = function () {
		return this.next;
	};
	CControl.prototype.getPrevious = function () {
		return this.previous;
	};
	CControl.prototype.setNext = function (v) {
		this.next = v;
	};
	CControl.prototype.setPrevious = function (v) {
		this.previous = v;
	};
	CControl.prototype.setParentControl = function (v) {
		this.parentControl = v;
	};
	CControl.prototype.getTiming = function () {
		var oSlide = this.getSlide();
		if (oSlide) {
			return oSlide.timing;
		}
		return null;
	};
	CControl.prototype.getSlide = function () {
		var oSlide = null;
		if (editor.WordControl && editor.WordControl.m_oLogicDocument) {
			oSlide = editor.WordControl.m_oLogicDocument.GetCurrentSlide();
			return oSlide;
		}
		return null;
	};
	CControl.prototype.getSlideNum = function () {
		var oSlide = this.getSlide();
		if (oSlide) {
			return oSlide.num;
		}
		return -1;
	};
	CControl.prototype.getFillColor = function () {
		console.error('Method "getFillColor" must be implemented in ' + this.constructor.name);
		return null;
	};
	CControl.prototype.getOutlineColor = function () {
		console.error('Method "getOutlineColor" must be implemented in ' + this.constructor.name);
		return null;
	};
	CControl.prototype.drawShdw = function () {

	};


	function CControlContainer(oParentControl) {
		CControl.call(this, oParentControl);
		this.children = [];
		this.recalcInfo.recalculateChildrenLayout = true;
		this.recalcInfo.recalculateChildren = true;

		this.eventListener = null;
	}

	InitClass(CControlContainer, CControl, CONTROL_TYPE_UNKNOWN);

	CControlContainer.prototype.isEventListener = function (oChild) {
		return this.eventListener === oChild;
	};
	CControlContainer.prototype.onScroll = function () {
	};
	CControlContainer.prototype.onStartScroll = function () {
	};
	CControlContainer.prototype.onEndScroll = function () {
	};
	CControlContainer.prototype.clear = function () {
		for (var nIdx = this.children.length - 1; nIdx > -1; --nIdx) {
			this.removeControl(this.children[nIdx]);
		}
	};
	CControlContainer.prototype.addControl = function (oChild) {
		var oLast = this.children[this.children.length - 1];
		this.children.push(oChild);
		if (oLast) {
			oLast.setNext(oChild);
			oChild.setPrevious(oLast);
			oChild.setParentControl(this);
		}
		return oChild;
	};
	CControlContainer.prototype.removeControl = function (oChild) {
		var nIdx = this.getChildIdx(oChild);
		this.removeByIdx(nIdx);
	};
	CControlContainer.prototype.removeByIdx = function (nIdx) {
		if (nIdx > -1 && nIdx < this.children.length) {
			var oChild = this.children[nIdx];
			oChild.setNext(null);
			oChild.setPrevious(null);
			oChild.setParentControl(null);
			var oPrev = this.children[nIdx - 1] || null;
			var oNext = this.children[nIdx + 1] || null;
			if (oPrev) {
				oPrev.setNext(oNext);
			}
			if (oNext) {
				oNext.setPrevious(oPrev);
			}
			this.children.splice(nIdx, 1);
		}
	};
	CControlContainer.prototype.getChildIdx = function (oChild) {
		for (var nChild = 0; nChild < this.children.length; ++nChild) {
			if (this.children[nChild] === oChild) {
				return nChild;
			}
		}
		return -1;
	};
	CControlContainer.prototype.getChildByType = function (nType) {
		for (var nChild = 0; nChild < this.children.length; ++nChild) {
			var oChild = this.children[nChild];
			if (oChild.getObjectType() === nType) {
				return oChild;
			}
		}
		return null;
	};
	CControlContainer.prototype.getChild = function (nIdx) {
		if (nIdx > -1 && nIdx < this.children.length) {
			return this.children[nIdx];
		}
	};
	CControlContainer.prototype.draw = function (graphics) {
		if (!CControl.prototype.draw.call(this, graphics)) {
			return false;
		}
		this.clipStart(graphics);
		for (var nChild = 0; nChild < this.children.length; ++nChild) {
			this.children[nChild].draw(graphics);
		}
		this.clipEnd(graphics);
		return true;
	};
	CControlContainer.prototype.clipStart = function (graphics) {
	};
	CControlContainer.prototype.clipEnd = function (graphics) {
	};
	CControlContainer.prototype.recalculateChildrenLayout = function () {
	};
	CControlContainer.prototype.recalculateChildren = function () {
	};
	CControlContainer.prototype.recalculate = function () {
		AscFormat.ExecuteNoHistory(function () {
			CControl.prototype.recalculate.call(this);
			if (this.recalcInfo.recalculateChildren) {
				this.recalculateChildren();
				this.recalcInfo.recalculateChildren = false;
			}
			if (this.recalcInfo.recalculateChildrenLayout) {
				this.recalculateChildrenLayout();
				this.recalcInfo.recalculateChildrenLayout = false;
			}
			for (var nChild = 0; nChild < this.children.length; ++nChild) {
				this.children[nChild].recalculate();
			}
		}, this, []);
	};
	CControlContainer.prototype.setLayout = function (dX, dY, dExtX, dExtY) {
		AscFormat.ExecuteNoHistory(function () {
			CControl.prototype.setLayout.call(this, dX, dY, dExtX, dExtY);
			this.recalcInfo.recalculateChildrenLayout = true;
		}, this, []);
	};
	CControlContainer.prototype.handleUpdateExtents = function () {
		this.recalcInfo.recalculateChildrenLayout = true;
		CControl.prototype.handleUpdateExtents.call(this);
	};
	CControlContainer.prototype.setEventListener = function (oChild) {
		if (oChild) {
			this.eventListener = oChild;
			if (this.parentControl) {
				this.parentControl.setEventListener(this);
			}
		} else {
			this.eventListener = null;
			if (this.parentControl) {
				this.parentControl.setEventListener(null);
			}
		}
	};
	CControlContainer.prototype.onMouseDown = function (e, x, y) {
		for (var nChild = this.children.length - 1; nChild >= 0; --nChild) {
			if (this.children[nChild].onMouseDown(e, x, y)) {
				return true;
			}
		}
		return CControl.prototype.onMouseDown.call(this, e, x, y);
	};
	CControlContainer.prototype.onMouseMove = function (e, x, y) {
		for (var nChild = this.children.length - 1; nChild >= 0; --nChild) {
			if (this.children[nChild].onMouseMove(e, x, y)) {
				return true;
			}
		}
		return CControl.prototype.onMouseMove.call(this, e, x, y);
	};
	CControlContainer.prototype.onMouseUp = function (e, x, y) {
		for (var nChild = this.children.length - 1; nChild >= 0; --nChild) {
			if (this.children[nChild].onMouseUp(e, x, y)) {
				return true;
			}
		}
		return CControl.prototype.onMouseUp.call(this, e, x, y);
	};
	CControlContainer.prototype.onMouseWheel = function (e, deltaY, X, Y) {
		for (var nChild = 0; nChild < this.children.length; ++nChild) {
			if (this.children[nChild].onMouseWheel(e, deltaY, X, Y)) {
				return true;
			}
		}
		return CControl.prototype.onMouseWheel.call(this, e, deltaY, X, Y);
	};
	CControlContainer.prototype.isScrolling = function () {
		for (var nChild = 0; nChild < this.children.length; ++nChild) {
			var oChild = this.children[nChild];
			if (oChild.isOnScroll && oChild.isOnScroll()) {
				return true;
			}
		}
		return false;
	};
	CControlContainer.prototype.canHandleEvents = function () {
		return false;
	};
	CControlContainer.prototype.onResize = function () {
		this.handleUpdateExtents();
		this.recalculate();
	};
	CControlContainer.prototype.getFillColor = function () {
		return null;
	};
	CControlContainer.prototype.getOutlineColor = function () {
		return null;
	};


	function CTopControl(oDrawer) {
		CControlContainer.call(this, null);
		this.drawer = oDrawer;
	}

	InitClass(CTopControl, CControlContainer, CONTROL_TYPE_UNKNOWN);

	CTopControl.prototype.onUpdateRect = function (oBounds) {
		if (this.drawer) {
			this.drawer.OnAnimPaneChanged(oBounds);
		}
	};
	CTopControl.prototype.onUpdate = function () {
		var oBounds = this.getBounds();
		this.onUpdateRect(oBounds);
	};
	CTopControl.prototype.onChildUpdate = function (oBounds) {
		this.onUpdateRect(oBounds);
	};
	CTopControl.prototype.onResize = function () {
		this.setLayout(0, 0, this.drawer.GetWidth(), this.drawer.GetHeight());
		CControlContainer.prototype.onResize.call(this);
		this.onUpdate();
	};


	function CLabel(oParentControl, sString, nFontSize, bBold, nParaAlign) {
		CControl.call(this, oParentControl);
		this.string = sString;
		this.fontSize = nFontSize;
		this.bold = bBold;
		this.paraAlign = nParaAlign;

		AscFormat.ExecuteNoHistory(function () {
			const oRGB = AscCommon.RgbaHexToRGBA(AscCommon.GlobalSkin.AnimPaneText);
			const oColor = new AscCommonWord.CDocumentColor(oRGB.R, oRGB.G, oRGB.B, false);

			var oTxLstStyle = new AscFormat.TextListStyle();
			oTxLstStyle.levels[0] = new CParaPr();
			oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
			oTxLstStyle.levels[0].DefaultRunPr.FontSize = this.fontSize;
			oTxLstStyle.levels[0].DefaultRunPr.Bold = this.bold;
			oTxLstStyle.levels[0].DefaultRunPr.Color = oColor;
			oTxLstStyle.levels[0].DefaultRunPr.RFonts.SetAll("Arial", -1);
			if (AscFormat.isRealNumber(this.paraAlign)) oTxLstStyle.levels[0].Jc = this.paraAlign;

			this.createTextBody();
			this.txBody.setLstStyle(oTxLstStyle);

			this.bodyPr = new AscFormat.CBodyPr();
			this.bodyPr.setDefault();
			this.bodyPr.anchor = 1; //vertical align ctr
			this.bodyPr.resetInsets();
			this.bodyPr.horzOverflow = AscFormat.nHOTClip;
			this.bodyPr.vertOverflow = AscFormat.nVOTClip;
		}, this, []);
	}

	InitClass(CLabel, CControl, CONTROL_TYPE_LABEL);

	CLabel.prototype.getString = function () {
		return AscCommon.translateManager.getValue(this.string);
	};
	CLabel.prototype.recalculateContent = function () {
		this.txBody.content.Recalc_AllParagraphs_CompiledPr();
		const oRGB = AscCommon.RgbaHexToRGBA(AscCommon.GlobalSkin.AnimPaneText);
		const oColor = new AscCommonWord.CDocumentColor(oRGB.R, oRGB.G, oRGB.B, false);
		this.txBody.lstStyle.levels[0].DefaultRunPr.Color = oColor;

		if (!this.txBody.bFit || !AscFormat.isRealNumber(this.txBody.fitWidth) || this.txBody.fitWidth > this.getWidth()) {
			this.txBody.recalculateOneString(this.getString());
		}

		// this.recalculateGeometry();
		this.recalculateTransform();
	};
	CLabel.prototype.canHandleEvents = function () {
		return false;
	};
	CLabel.prototype.getFillColor = function () {
		return null;
	};
	CLabel.prototype.getOutlineColor = function () {
		return null;
	};
	CLabel.prototype.recalculateTransformText = function () {
		var Y = this.getHeight() / 2 - this.txBody.content.GetSummaryHeight() / 2;
		if (!this.transformText) {
			this.transformText = new AscCommon.CMatrix();
		}
		this.transformText.tx = this.transform.tx;
		this.transformText.ty = this.transform.ty + Y;

		if (!this.invertTransformText) {
			this.invertTransformText = new AscCommon.CMatrix();
		}
		this.invertTransformText.tx = -this.transformText.tx;
		this.invertTransformText.ty = -this.transformText.ty;
		this.localTransformText = this.transformText;
	};
	CLabel.prototype.recalculateTransformText2 = function () {
		return null;
	};

	CLabel.prototype.getContentOneStringSizes = function () {
		this.recalculateContent();
		return this.txBody.getContentOneStringSizes();
	};

	CLabel.prototype.setLayout = function (dX, dY, dExtX, dExtY) {
		this.superclass.prototype.setLayout.call(this, dX, dY, dExtX, dExtY);
		if(this.txBody) {
			this.txBody.bFit = false;
		}
	};


	function CImageControl(oParentControl, sBase64Image, width, height) {
		CControl.call(this, oParentControl)
		this.src = sBase64Image;
		this.imgWidth = width || 0;
		this.imgHeight = height || 0;
	}

	InitClass(CImageControl, CControl, CONTROL_TYPE_IMAGE);

	CImageControl.prototype.canHandleEvents = function () {
		return false;
	};
	CImageControl.prototype.draw = function(graphics) {
		if (!this.src) { return false; }
		if (this.isHidden()) { return false; }
		if (!this.checkUpdateRect(graphics.updatedRect)) { return false; }

		const left = this.bounds.l + (this.getWidth() - this.imgWidth) / 2;
		const top = this.bounds.t + (this.getHeight() - this.imgHeight) / 2;

		graphics.SaveGrState();
		graphics.SetIntegerGrid(false);
		graphics.drawImage(this.src, left, top, this.imgWidth, this.imgHeight);
		graphics.RestoreGrState();
	};
	CImageControl.prototype.getFillColor = function () {
		return null;
	};
	CImageControl.prototype.getOutlineColor = function () {
		return null;
	};


	function CButton(oParentControl, fOnMouseDown, fOnMouseMove, fOnMouseUp) {
		CControlContainer.call(this, oParentControl);
		this.onMouseDownCallback = fOnMouseDown;
		this.onMouseMoveCallback = fOnMouseMove;
		this.onMouseUpCallback = fOnMouseUp;
	}

	InitClass(CButton, CControlContainer, CONTROL_TYPE_BUTTON);

	CButton.prototype.onMouseDown = function (e, x, y) {
		if (this.onMouseDownCallback && this.onMouseDownCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseDown.call(this, e, x, y);
	};
	CButton.prototype.onMouseMove = function (e, x, y) {
		if (this.onMouseMoveCallback && this.onMouseMoveCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseMove.call(this, e, x, y);
	};
	CButton.prototype.onMouseUp = function (e, x, y) {
		if (this.onMouseUpCallback && this.onMouseUpCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseUp.call(this, e, x, y);
	};
	CButton.prototype.canHandleEvents = function () {
		return true;
	};

	CButton.prototype.draw = function (graphics) {
		if (!CControl.prototype.draw.call(this, graphics)) { return false; }

		const contentOpacity = this.getContentOpacity();
		graphics.put_GlobalAlpha(true, contentOpacity);
		for (var nChild = 0; nChild < this.children.length; ++nChild) {
			this.children[nChild].draw(graphics);
		}
		graphics.put_GlobalAlpha(false);

		return true;
	};

	CButton.prototype.getContentOpacity = function () {
		if (this.isDisabled()) { return 0.6 }
		return 1;
	}
	CButton.prototype.getFillColor = function () {
		const oSkin = AscCommon.GlobalSkin;
		if (this.isDisabled()) { return null; }
		if (this.isHovered()) { return oSkin.AnimPaneButtonFillHovered; }
		if (this.sName === 'playButton') { return oSkin.type === 'light' ? oSkin.AnimPanePlayButtonFill : null; }
		return null;
	};
	CButton.prototype.getOutlineColor = function () {
		if (this.sName === 'playButton') { return AscCommon.GlobalSkin.AnimPanePlayButtonOutline; }
		return null;
	};
	CButton.prototype.isPressed = function () {
		return this.getStateFlag(STATE_FLAG_PRESSED);
	};
	CButton.prototype.disable = function () {
		return this.setStateFlag(STATE_FLAG_DISABLED, true)
	};
	CButton.prototype.enable = function () {
		return this.setStateFlag(STATE_FLAG_DISABLED, false)
	};
	CButton.prototype.isDisabled = function () {
		return this.getStateFlag(STATE_FLAG_DISABLED);
	};


	function CAnimPaneHeader(oDrawer) {
		CTopControl.call(this, oDrawer);
		this.label = this.addControl(new CLabel(this, 'Animation Pane', HEADER_LABEL_FONTSIZE, true, AscCommon.align_Left));

		this.playButton = this.addControl(new CButton(this, null, null, managePreview));
		this.playButton.sName = 'playButton';
		this.playButton.icon = this.playButton.addControl(new CImageControl(
			this.playButton, null,
			10 * AscCommon.g_dKoef_pix_to_mm, 10 * AscCommon.g_dKoef_pix_to_mm /* 10x10 svg icon (play and stop icons both) */
		));
		this.playButton.label = this.playButton.addControl(new CLabel(this.playButton, '', PLAY_BUTTON_LABEL_FONTSIZE));

		this.moveUpButton = this.addControl(new CButton(this, null, null, moveChosenUp));
		this.moveUpButton.icon = this.moveUpButton.addControl(new CImageControl(
			this.moveUpButton,
			AscCommon.GlobalSkin.type == 'light' ? arrowUpIcon_dark : arrowUpIcon_light,
			12 * AscCommon.g_dKoef_pix_to_mm, 7 * AscCommon.g_dKoef_pix_to_mm /* 12x7 svg icon*/
		));

		this.moveDownButton = this.addControl(new CButton(this, null, null, moveChosenDown));
		this.moveDownButton.icon = this.moveDownButton.addControl(new CImageControl(
			this.moveDownButton,
			AscCommon.GlobalSkin.type == 'light' ? arrowDownIcon_dark : arrowDownIcon_light,
			12 * AscCommon.g_dKoef_pix_to_mm, 7 * AscCommon.g_dKoef_pix_to_mm /* 12x7 svg icon*/
		));

		this.closeButton = this.addControl(new CButton(this, null, null, closePanel));
		this.closeButton.icon = this.closeButton.addControl(new CImageControl(
			this.closeButton,
			AscCommon.GlobalSkin.type == 'light' ? closeIcon_dark : closeIcon_light,
			10 * AscCommon.g_dKoef_pix_to_mm, 10 * AscCommon.g_dKoef_pix_to_mm /* 10x10 svg icon*/
		));

		// Event handlers for button of CAnimPaneHeader ---

		function managePreview(event, x, y) {
			if (!this.hit(x, y)) { return; }
			if (this.isDisabled()) { return; }
			if(Asc.editor.asc_IsStartedAnimationPreview()) {
				Asc.editor.asc_StopAnimationPreview();
			} else {
				const timing = this.parentControl.getTiming();
				const bIncludeFollowing = timing && (timing.getSelectedEffects().length === 1);
				Asc.editor.asc_StartAnimationPreview(this.parentControl.isStartAllPreview(), bIncludeFollowing);
			}
		}

		function moveChosenUp(event, x, y) {
			if (!this.hit(x, y)) { return }
			if (this.isDisabled()) { return }
			if (Asc.editor.asc_canMoveAnimationEarlier()) {
				if (Asc.editor.asc_IsStartedAnimationPreview()) {
					Asc.editor.asc_StopAnimationPreview()
				}
				Asc.editor.asc_moveAnimationEarlier()
			}
		}

		function moveChosenDown(event, x, y) {
			if (!this.hit(x, y)) { return }
			if (this.isDisabled()) { return }
			if (Asc.editor.asc_canMoveAnimationLater()) {
				if (Asc.editor.asc_IsStartedAnimationPreview()) {
					Asc.editor.asc_StopAnimationPreview()
				}
				Asc.editor.asc_moveAnimationLater()
			}
		}

		function closePanel(event, x, y) {
			if (!this.hit(x, y)) { return }
			if (this.isDisabled()) { return }
			Asc.editor.asc_ShowAnimPane(false);
			Asc.editor.sendEvent('asc_onCloseAnimPane');
		}

		// --- end of event handlers for buttons of CAnimPaneHeader
	}

	InitClass(CAnimPaneHeader, CTopControl, CONTROL_TYPE_HEADER);

	CAnimPaneHeader.prototype.isStartAllPreview = function () {
		const timing = this.getTiming();
		if (!timing) { return true; }
		const aSelectedEffects = timing.getSelectedEffects();
		return aSelectedEffects.length === 0;
	};
	CAnimPaneHeader.prototype.getPlayButtonText = function () {
		if (Asc.editor.asc_IsStartedAnimationPreview()) { return 'Stop'; }
		if (this.isStartAllPreview()) { return 'Play All'; } // No timing or zero selected effects
		if (this.getTiming().getSelectedEffects().length === 1) { return 'Play From'; } // One selected
		return 'Play Selected'; // Many selected
	};
	CAnimPaneHeader.prototype.getPlayButtonIcon = function() {
		let sPlayButtonIcon = "";
		if(Asc.editor.asc_IsStartedAnimationPreview()) {
			sPlayButtonIcon = AscCommon.GlobalSkin.type == 'light' ? stopIcon_dark : stopIcon_light;
		}
		else {
			sPlayButtonIcon = AscCommon.GlobalSkin.type == 'light' ? playIcon_dark : playIcon_light;
		}
		return sPlayButtonIcon;
	};
	CAnimPaneHeader.prototype.checkLayout = function() {
		if(this.playButton.label.string !== this.getPlayButtonText()) {
			this.recalculateChildrenLayout();
			this.onUpdate();
		}
	};
	CAnimPaneHeader.prototype.recalculateChildrenLayout = function () {
		let gap;
		this.moveUpButton.icon.src =AscCommon.GlobalSkin.type == 'light' ? arrowUpIcon_dark : arrowUpIcon_light;
		this.moveDownButton.icon.src = AscCommon.GlobalSkin.type == 'light' ? arrowDownIcon_dark : arrowDownIcon_light;
		this.closeButton.icon.src = AscCommon.GlobalSkin.type == 'light' ? closeIcon_dark : closeIcon_light;

		this.label.setLayout(COMMON_LEFT_MARGIN, 0, HEADER_LABEL_WIDTH, this.getHeight());



		let sPlayButtonText = this.getPlayButtonText();
		let sPlayButtonIcon = this.getPlayButtonIcon();
		this.playButton.icon.src = sPlayButtonIcon;
		gap = (PLAY_BUTTON_HEIGHT - PLAY_BUTTON_ICON_SIZE) / 2;
		this.playButton.icon.setLayout(PLAY_BUTTON_LEFT_PADDING, gap, PLAY_BUTTON_ICON_SIZE, PLAY_BUTTON_ICON_SIZE);

		let oButtonLabel = this.playButton.label;
		oButtonLabel.string = sPlayButtonText;
		oButtonLabel.setLayout(
			this.playButton.icon.getRight() + PLAY_BUTTON_LABEL_LEFT_MARGIN,
			0,
			PLAY_BUTTON_MAX_LABEL_WIDTH,
			PLAY_BUTTON_HEIGHT
		);
		// let dLabelWidth = Math.min(PLAY_BUTTON_MAX_LABEL_WIDTH, oButtonLabel.getContentOneStringSizes().w)
		let dLabelWidth = PLAY_BUTTON_MAX_LABEL_WIDTH;
		oButtonLabel.setLayout(
			this.playButton.icon.getRight() + PLAY_BUTTON_LABEL_LEFT_MARGIN,
			0,
			dLabelWidth,
			PLAY_BUTTON_HEIGHT
		);
		oButtonLabel.recalculate();

		this.playButton.setLayout(
			this.label.getRight() + PLAY_BUTTON_LEFT_MARGIN,
			(HEADER_HEIGHT - PLAY_BUTTON_HEIGHT) / 2,
			this.playButton.icon.getRight() + PLAY_BUTTON_LABEL_LEFT_MARGIN + oButtonLabel.getWidth() + PLAY_BUTTON_RIGHT_PADDING,
			PLAY_BUTTON_HEIGHT
		);


		this.moveUpButton.icon.setLayout(0, 0, MOVE_BUTTON_SIZE, MOVE_BUTTON_SIZE);

		gap = (HEADER_HEIGHT - MOVE_BUTTON_SIZE) / 2;
		this.moveUpButton.setLayout(
			this.playButton.getRight() + MOVE_UP_BUTTON_LEFT_MARGIN,
			gap,
			MOVE_BUTTON_SIZE,
			MOVE_BUTTON_SIZE
		);
		this.moveUpButton.recalculate();

		this.moveDownButton.icon.setLayout(0, 0, MOVE_BUTTON_SIZE, MOVE_BUTTON_SIZE);

		this.moveDownButton.setLayout(
			this.moveUpButton.getRight() + MOVE_DOWN_BUTTON_LEFT_MARGIN,
			gap,
			MOVE_BUTTON_SIZE,
			MOVE_BUTTON_SIZE
		);
		this.moveDownButton.recalculate();

		this.closeButton.icon.setLayout(0, 0, CLOSE_BUTTON_SIZE, CLOSE_BUTTON_SIZE);
		gap = (HEADER_HEIGHT - CLOSE_BUTTON_SIZE) / 2;
		this.closeButton.setLayout(
			this.getRight() - COMMON_RIGHT_MARGIN - CLOSE_BUTTON_SIZE,
			gap,
			CLOSE_BUTTON_SIZE,
			CLOSE_BUTTON_SIZE
		);
		this.closeButton.recalculate();
	};
	CAnimPaneHeader.prototype.onMouseMove = function (e, x, y) {
		const animPane = Asc.editor.WordControl.m_oAnimPaneApi;
		animPane.SetCursorType('default', new CMouseMoveData());

		for (var nChild = this.children.length - 1; nChild >= 0; --nChild) {
			if (this.children[nChild].onMouseMove(e, x, y)) {
				return true;
			}
		}
		return CControl.prototype.onMouseMove.call(this, e, x, y);
	};


	function CTimelineContainer(oDrawer) {
		CTopControl.call(this, oDrawer);
		this.drawer = oDrawer;

		this.zoomOutButton = this.addControl(
			new CButton(this, null, null, function (e, x, y) {
				if (this.hit(x, y)) editor.asc_ZoomOutTimeline();
			})
		);
		this.zoomOutButton.icon = this.zoomOutButton.addControl(
			new CImageControl(
				this.zoomOutButton,
				AscCommon.GlobalSkin.type == 'light' ? zoomOutIcon_dark : zoomOutIcon_light,
				10 * AscCommon.g_dKoef_pix_to_mm, 1 * AscCommon.g_dKoef_pix_to_mm /* 10x1 svg image */
			)
		);

		this.zoomLabel = this.addControl(new CLabel(this, 'Zoom', ZOOM_LABEL_FONTSIZE, false, AscCommon.align_Center));

		this.zoomInButton = this.addControl(
			new CButton(this, null, null, function (e, x, y) {
				if (this.hit(x, y)) editor.asc_ZoomInTimeline();
			})
		);
		this.zoomInButton.icon = this.zoomInButton.addControl(
			new CImageControl(
				this.zoomInButton,
				AscCommon.GlobalSkin.type == 'light' ? zoomInIcon_dark : zoomInIcon_light,
				11 * AscCommon.g_dKoef_pix_to_mm, 11 * AscCommon.g_dKoef_pix_to_mm /* 11x11 svg image */
			)
		);

		this.timeline = this.addControl(new CTimeline(this));

		this.onMouseDownCallback = function (event, x, y) {
			if(Asc.editor.asc_IsStartedAnimationPreview()) {
				Asc.editor.asc_StopAnimationPreview();
			}
		}
	}

	InitClass(CTimelineContainer, CTopControl, CONTROL_TYPE_TIMELINE_CONTAINER);

	CTimelineContainer.prototype.recalculateChildrenLayout = function () {
		this.zoomOutButton.icon.src = AscCommon.GlobalSkin.type == 'light' ? zoomOutIcon_dark : zoomOutIcon_light;
		this.zoomInButton.icon.src = AscCommon.GlobalSkin.type == 'light' ? zoomInIcon_dark : zoomInIcon_light;
		this.zoomInButton.setLayout(
			TIMELINE_SCROLL_ABSOLUTE_LEFT - TIMELINE_HEIGHT + (TIMELINE_HEIGHT - ZOOM_BUTTON_SIZE) / 2,
			(TIMELINE_HEIGHT - ZOOM_BUTTON_SIZE) / 2,
			ZOOM_BUTTON_SIZE,
			ZOOM_BUTTON_SIZE
		);
		this.zoomInButton.icon.setLayout(0, 0, ZOOM_BUTTON_SIZE, ZOOM_BUTTON_SIZE);

		this.zoomLabel.setLayout(
			this.zoomInButton.getLeft() - ZOOM_LABEL_WIDTH,
			0,
			ZOOM_LABEL_WIDTH,
			this.getHeight()
		);

		this.zoomOutButton.setLayout(
			this.zoomLabel.getLeft() - ZOOM_BUTTON_SIZE,
			(TIMELINE_HEIGHT - ZOOM_BUTTON_SIZE) / 2,
			ZOOM_BUTTON_SIZE,
			ZOOM_BUTTON_SIZE
		);
		this.zoomOutButton.icon.setLayout(0, 0, ZOOM_BUTTON_SIZE, ZOOM_BUTTON_SIZE);

		const timelineWidth = this.getWidth() -
			(COMMON_LEFT_MARGIN + COMMON_RIGHT_MARGIN) -
			(SCALE_BUTTON_LEFT_MARGIN + SCALE_BUTTON_WIDTH + TIMELINE_SCROLL_LEFT_MARGIN) - (ANIM_ITEM_HEIGHT - MENU_BUTTON_SIZE) / 2;
		this.timeline.setLayout(
			TIMELINE_SCROLL_ABSOLUTE_LEFT,
			(TIMELINE_HEIGHT - TIMELINE_SCROLL_HEIGHT) / 2,
			timelineWidth,
			TIMELINE_SCROLL_HEIGHT
		);
	};
	CTimelineContainer.prototype.draw = function (graphics) {
		if (!CTopControl.prototype.draw.call(this, graphics)) {
			return false;
		}
		// this.clipStart(graphics);
		for (var nChild = 0; nChild < this.children.length; ++nChild) {
			this.children[nChild].draw(graphics);
		}
		// graphics.RemoveClipRect();
		// this.clipEnd(graphics);
		return true;
	};
	CTimelineContainer.prototype.onMouseDown = function (e, x, y) {
		if (this.onMouseDownCallback && this.onMouseDownCallback.call(this, e, x, y)) {
			return true;
		}
		return CTopControl.prototype.onMouseDown.call(this, e, x, y);
	};


	function CTimeline(oParentControl, oContainer, oChild) {
		CControlContainer.call(this, oParentControl);

		this.container = oContainer;
		this.scrolledChild = oChild;

		this.isScrollerHovered;
		this.isStickedToPointer;
		
		this.startButton = this.addControl(new CButton(this, onFirstBtnMouseDown, null, onMouseUp));
		this.startButton.icon = this.startButton.addControl(new CImageControl(
			this.startButton,
			AscCommon.GlobalSkin.type == 'light' ? arrowLeftIcon_dark : arrowLeftIcon_light,
			5 * AscCommon.g_dKoef_pix_to_mm, 9 * AscCommon.g_dKoef_pix_to_mm /* 5x9 svg image */
		));

		this.endButton = this.addControl(new CButton(this, onSecondBtnMouseDown, null, onMouseUp));
		this.endButton.icon = this.endButton.addControl(new CImageControl(
			this.endButton,
			AscCommon.GlobalSkin.type == 'light' ? arrowRightIcon_dark : arrowRightIcon_light,
			5 * AscCommon.g_dKoef_pix_to_mm, 9 * AscCommon.g_dKoef_pix_to_mm /* 5x9 svg image */
		));
		
		function onFirstBtnMouseDown(e, x, y) {
			if (!this.hit(x, y)) { return }
			this.parentControl.setEventListener(this);
			let step = SCROLL_STEP * this.parentControl.getWidth()
			this.parentControl.startScroll(-step);
			return true;
		}

		function onSecondBtnMouseDown(e, x, y) {
			if (!this.hit(x, y)) { return }
			this.parentControl.setEventListener(this);
			let step = SCROLL_STEP * this.parentControl.getWidth()
			this.parentControl.startScroll(step);
			return true;
		}

		function onMouseUp(e, x, y) {
			if (this.parentControl.isEventListener(this)) {
				this.parentControl.setEventListener(null);
				this.parentControl.endScroll();
				return true;
			}
			return false;
		}
		
		this.timerId = null;
		this.timeoutId = null;

		// This fields supposed to be private
		// so it should not be changed directly.
		// Use set methods instead (setScrollOffset, setStartTime)
		this.scrollOffset = 0; // in millimeters
		this.startTime = 0; // in seconds
		this.timeScaleIndex = 2;

		// Tmp field for demoPreview
		this.tmpScrollOffset = null;

		// Labels cache
		this.labels = {};
		this.usedLabels = {};
		this.cachedParaPr = null

		this.onMouseDownCallback = function stickToPointer(event, x, y) {
			if (!this.hitInScroller(x, y)) { return; }
			this.isStickedToPointer = true;
			this.onUpdate()
		}

		this.onMouseUpCallback = function unstickFromPointer(event, x, y) {
			this.isStickedToPointer = false;
			if (this.isOnScroll()) { this.endScroll(); }
			this.onUpdate();
		}

		this.onMouseMoveCallback = function handlePointerMovement(event, x, y) {
			// Updating hover state of the scroller
			const tmpIsScrollerHovered = this.hitInScroller(x, y);
			if (this.isScrollerHovered !== tmpIsScrollerHovered) {
				this.isScrollerHovered = tmpIsScrollerHovered;
				this.onUpdate();
			}

			if (!this.isStickedToPointer) { return; }

			let oInv = this.getInvFullTransformMatrix();
			let tx = oInv.TransformPointX(x, y);

			let newScrollOffset = tx - this.getRulerStart() - TIMELINE_SCROLLER_WIDTH / 2;

			// Check if the boundaried are reached and start scrolling if so
			let leftBorder = this.getRulerStart();
			let rightBorder = this.getRulerEnd();
			if (tx <= leftBorder || tx >= rightBorder) {
				if (!this.isOnScroll()) {
					let scrollStep = this.getWidth() * SCROLL_STEP / 10;
					scrollStep = tx <= leftBorder ? -scrollStep : scrollStep;
					let scrollTimerDelay = 0;
					let scrollTimerInterval = 50;
					this.startScroll(scrollStep, scrollTimerDelay, scrollTimerInterval);
				}
			}
			else this.endScroll();

			// Updating scrollOffset
			this.setScrollOffset(newScrollOffset)
		}
	}

	InitClass(CTimeline, CControlContainer, CONTROL_TYPE_TIMELINE);

	CTimeline.prototype.limitScrollOffset = function (newScrollOffset /* in millimeters */) {
		return Math.max(0, Math.min(newScrollOffset, this.getMaxScrollOffset()));
	};
	CTimeline.prototype.getScrollOffset = function () {
		return this.tmpScrollOffset !== null ? this.tmpScrollOffset : this.scrollOffset;
	};
	CTimeline.prototype.setScrollOffset = function (newScrollOffset /* in millimeters */) {
		let oldScrollOffset = this.getScrollOffset()

		this.scrollOffset = this.limitScrollOffset(newScrollOffset)

		let difference = this.posToTime(this.getScrollOffset()) - this.posToTime(oldScrollOffset) // difference in seconds
		this.setStartTime(this.getStartTime() + difference)

		this.parentControl.onScroll();
		this.onUpdate();
	};
	CTimeline.prototype.getMaxScrollOffset = function () {
		return this.getWidth() - 2 * TIMELINE_SCROLL_BUTTON_SIZE - TIMELINE_SCROLLER_WIDTH;
	};

	CTimeline.prototype.getStartTime = function () {
		return this.startTime;
	};
	CTimeline.prototype.setStartTime = function (newStartTime /* in seconds */) {
		this.startTime = Math.max(0, newStartTime)

		this.parentControl.onScroll();
		this.onUpdate();

		// also updating seqList to redraw effect bars
		editor.WordControl.m_oAnimPaneApi.list.Control.seqList.onUpdateSeqList()
	};
	CTimeline.prototype.getCurrentTime = function() {
		return this.posToTime(this.getScrollOffset() + this.startButton.getWidth() + TIMELINE_SCROLLER_WIDTH / 2)
	}

	CTimeline.prototype.startScroll = function (step /* in millimeters */, scrollTimerDelay, scrollTimerInterval) {
		if (typeof scrollTimerDelay === 'undefined') { scrollTimerDelay = SCROLL_TIMER_DELAY }
		if (typeof scrollTimerInterval === 'undefined') { scrollTimerInterval = SCROLL_TIMER_INTERVAL }

		this.endScroll();
		var oScroll = this;
		oScroll.addScroll(step);

		this.timeoutId = setTimeout(function () {
			oScroll.timeoutId = null;
			oScroll.timerId = setInterval(function () {
				oScroll.addScroll(step);
			}, scrollTimerInterval);
		}, scrollTimerDelay);
	};
	CTimeline.prototype.addScroll = function (step /* in millimeters */) {
		let newStartTime = this.posToTime(this.getZeroShift() + step)

		const seqList = editor.WordControl.m_oAnimPaneApi.list.Control.seqList
		seqList.forEachAnimItem(function (animItem) {
			if (!animItem.hitResult) { return }
			animItem.handleTimelineScroll(step);
		})

		this.setStartTime(newStartTime)
	};
	CTimeline.prototype.endScroll = function () {
		if (this.timerId !== null) {
			clearInterval(this.timerId);
			this.timerId = null;
		}
		if (this.timeoutId !== null) {
			clearTimeout(this.timeoutId);
			this.timeoutId = null;
		}

		this.setStateFlag(STATE_FLAG_SELECTED, false);
	};
	CTimeline.prototype.isOnScroll = function () {
		return this.timerId !== null || this.timeoutId !== null;
		// return this.timerId !== null || this.timeoutId !== null || this.parentControl.isEventListener(this);
	};

	CTimeline.prototype.startDrawLabels = function () {
		this.usedLabels = {};
	};
	CTimeline.prototype.endDrawLabels = function () {
		for (var nTime in this.labels) {
			if (!this.usedLabels[nTime]) {
				var oLabel = this.labels[nTime];
				oLabel.parentControl = null;
				oLabel.bDeleted = true;
				delete this.labels[nTime];
			}
		}
	};
	CTimeline.prototype.getLabel = function (nTime, scale) {
		this.usedLabels[nTime] = true;
		if (this.labels[nTime] && AscFormat.fApproxEqual(this.labels[nTime].scale, scale, 0.01)) {
			return this.labels[nTime];
		}
		return this.cacheLabel(nTime, scale);
	};
	CTimeline.prototype.cacheLabel = function (nTime, scale) {
		var oLabel = new CLabel(this, this.getTimeString(nTime), TIMELINE_LABEL_FONTSIZE, false, AscCommon.align_Center);
		var oContent = oLabel.txBody.content;
		oLabel.setLayout(0, 0, TIMELINE_LABEL_WIDTH, this.getHeight());
		if (this.cachedParaPr) {
			oContent.Content[0].CompiledPr = this.cachedParaPr;
		} else {
			oContent.SetApplyToAll(true);
			oContent.SetParagraphAlign(AscCommon.align_Center);
			oContent.SetApplyToAll(false);
		}
		oLabel.recalculate();
		if (!this.cachedParaPr) {
			this.cachedParaPr = oContent.Content[0].CompiledPr;
		}
		var oBaseTexture = oLabel.getAnimTexture(scale);
		if (oBaseTexture) {
			this.labels[nTime] = new CAnimTexture(this, oBaseTexture.canvas, oBaseTexture.scale, oBaseTexture.x, oBaseTexture.y);
		}
		return this.labels[nTime];
	};
	CTimeline.prototype.getTimeString = function (nTime) {
		if (nTime < 60) {
			return "" + nTime;
		}

		const nSeconds = nTime % 60;
		const nMinutes = ((nTime / 60) >> 0) % 60;

		let sSeconds = padZero(nSeconds);
		let sMinutes = padZero(nMinutes);

		if (nTime < 3600) {
			return (sMinutes + ":") + sSeconds;
		}

		return (((nTime / 3600) >> 0) + ":") + sMinutes + ":" + sSeconds;

		function padZero(number) {
			return number < 10 ? "0" + number : "" + number;
		}
	};
	CTimeline.prototype.drawLabel = function (graphics, dPos, nTime) {
		var oLabelTexture = this.getLabel(nTime, graphics.m_oCoordTransform.sx);
		var oMatrix = new AscCommon.CMatrix();
		var dWidth = oLabelTexture.canvas.width / oLabelTexture.scale;
		var dHeight = oLabelTexture.canvas.height / oLabelTexture.scale;
		graphics.drawImage2(oLabelTexture.canvas,
			dPos - dWidth / 2, this.getHeight() / 2 - dHeight / 2,
			dWidth,
			dHeight);
		// var oContent = oLabel.txBody.content;
		// oContent.ShiftView(dPos - TIMELINE_LABEL_WIDTH / 2, this.getHeight() / 2 - oContent.GetSummaryHeight() / 2);
		// oContent.Draw(0, graphics);
		// oContent.ResetShiftView();
	};
	CTimeline.prototype.drawMark = function (graphics, dPos) {
		var dHeight = this.getHeight() / 3;
		var nPenW = this.getPenWidth(graphics);
		graphics.drawVerLine(1, dPos, dHeight, dHeight + dHeight, nPenW);
	};
	CTimeline.prototype.start = function (graphics, dPos) {
		var dHeight = this.getHeight() / 3;
		var nPenW = this.getPenWidth(graphics);
		graphics.drawVerLine(1, dPos, dHeight, dHeight + dHeight, nPenW);
	};
	CTimeline.prototype.handleUpdateExtents = function () {
		this.labels = {};
		this.usedLabels = {};
		this.cachedParaPr = null;
		if(this.startButton && this.endButton) {
			this.startButton.icon.src = AscCommon.GlobalSkin.type == 'light' ? arrowLeftIcon_dark : arrowLeftIcon_light;
			this.endButton.icon.src = AscCommon.GlobalSkin.type == 'light' ? arrowRightIcon_dark : arrowRightIcon_light;
		}
		CControlContainer.prototype.handleUpdateExtents.call(this);
	};
	CTimeline.prototype.draw = function (graphics) {
		if (this.isHidden()) { return false }
		if (!this.checkUpdateRect(graphics.updatedRect)) { return false }

		this.getStartTime() === 0 ? this.startButton.disable() : this.startButton.enable();

		graphics.SaveGrState();
		// var dPenW = this.getPenWidth(graphics);
		// graphics.SetIntegerGrid(true);
		// graphics.p_width(dPenW);
		// var sColor = this.children[0].getOutlineColor();
		// var oColor = AscCommon.RgbaHexToRGBA(sColor);
		// graphics.p_color(oColor.R, oColor.G, oColor.B, 255);
		// var dPaneLeft = this.children[0].getRight();
		// var dPaneWidth = this.getWidth() - (this.children[0].getWidth() + this.children[1].getWidth());
		// graphics.rect(dPaneLeft, 0, dPaneWidth, this.getHeight());
		// graphics.ds();
		// graphics.RestoreGrState();
		const sColor = AscCommon.GlobalSkin.AnimPaneTimelineRulerOutline;
		var oColor = AscCommon.RgbaHexToRGBA(sColor);
		var dPaneLeft = this.getRulerStart();
		var dPaneWidth = this.getRulerEnd() - dPaneLeft;
		var x = dPaneLeft;
		var y = 0;
		var extX = dPaneWidth;
		var extY = this.getHeight();
		graphics.transform3(this.transform);
		graphics.SetIntegerGrid(true);
		var nPenW = this.getPenWidth(graphics);
		graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
		graphics.drawHorLine(0, y, x, x + extX, nPenW);
		graphics.drawHorLine(0, y + extY, x, x + extX, nPenW);
		graphics.drawVerLine(2, x, y, y + extY, nPenW);
		graphics.drawVerLine(2, x + extX, y, y + extY, nPenW);
		graphics.ds();

		//draw marks
		//find first visible
		var fStartTime = this.posToTime(this.getRulerStart());
		var fTimeInterval = TIME_SCALES[this.timeScaleIndex];
		var nMarksCount = TIME_INTERVALS[this.timeScaleIndex] === LONG_TIME_INTERVAL ? 10 : 2;

		var dTimeOfSmallInterval = fTimeInterval / nMarksCount;
		var nStartIntervalIdx = this.startTime / dTimeOfSmallInterval >> 0;
		var nEndIntervalIdx = this.posToTime(this.getRulerEnd()) / dTimeOfSmallInterval + 0.5 >> 0;
		this.startDrawLabels();

		graphics.SaveGrState();
		var nInterval;
		graphics.AddClipRect(x, y, extX, extY);
		oColor = AscCommon.RgbaHexToRGBA(AscCommon.GlobalSkin.AnimPaneTimelineRulerTick);
		graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
		for (nInterval = nStartIntervalIdx; nInterval <= nEndIntervalIdx; ++nInterval) {
			var dTime = nInterval * dTimeOfSmallInterval;
			var dPos = this.timeToPos(dTime);
			if (nInterval % nMarksCount !== 0) {
				this.drawMark(graphics, dPos);
			} else {
				this.drawLabel(graphics, dPos, dTime);
			}
		}
		graphics.ds();
		// for(nInterval = nFirstInterval; nInterval <= nLastInterval; ++nInterval) {
		//     var dTime = nInterval*dSmallInterval;
		//     var dPos = this.timeToPos(dTime);
		//     if(nInterval % nMarksCount === 0) {
		//         this.drawLabel(graphics, dPos, dTime);
		//     }
		// }

		graphics.RestoreGrState();
		this.endDrawLabels();
		//

		this.drawScroller(graphics);

		graphics.RestoreGrState();

		if (!CControlContainer.prototype.draw.call(this, graphics)) {
			return false;
		}
	};
	CTimeline.prototype.drawScroller = function (graphics) {
		let x = this.getRulerStart() + this.getScrollOffset();
		let y = 0;
		let extX = TIMELINE_SCROLLER_WIDTH;
		let extY = this.getHeight();

		const oSkin = AscCommon.GlobalSkin;
		const oFillColor = AscCommon.RgbaHexToRGBA(oSkin.AnimPaneTimelineScrollerFill);

		let nOpacity;
		if (this.isStickedToPointer) {
			nOpacity = 0.5;//oSkin.AnimPaneTimelineScrollerOpacityActive;
		} else if (this.isScrollerHovered) {
			nOpacity = oSkin.AnimPaneTimelineScrollerOpacityHovered;
		} else {
			nOpacity = oSkin.AnimPaneTimelineScrollerOpacity;
		}

		graphics.b_color1(oFillColor.R, oFillColor.G, oFillColor.B, nOpacity * 255);
		graphics.rect(x, y, extX, extY);
		graphics.df();

		let nPenW = this.getPenWidth(graphics);


		let sColor = oSkin.AnimPaneTimelineScrollerOutline;
		let oColor = AscCommon.RgbaHexToRGBA(sColor);
		graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
		graphics.drawHorLine(0, y, x, x + extX, nPenW);
		graphics.drawHorLine(0, y + extY, x, x + extX, nPenW);
		graphics.drawVerLine(2, x, y, y + extY, nPenW);
		graphics.drawVerLine(2, x + extX, y, y + extY, nPenW);

		if (this.tmpScrollOffset !== null) {
			graphics.SaveGrState();
			graphics.RemoveClipRect();

			// const seqList = Asc.editor.WordControl.m_oAnimPaneApi.list.Control.seqList
			// graphics.drawVerLine(1, x + extX / 2, seqList.getTop(), y, nPenW);
			graphics.drawVerLine(1, x + extX / 2, y - 1000, y, nPenW);

			graphics.RestoreGrState();
		}

		return true;
	};
	CTimeline.prototype.hitInScroller = function(x, y) {
		// x, y - relatively to this.parentContainer
		// tx, ty - relatively to this

		let oInv = this.getInvFullTransformMatrix();
		let tx = oInv.TransformPointX(x, y);
		let ty = oInv.TransformPointY(x, y);

		let l = this.getRulerStart() + this.getScrollOffset();
		let t = 0;
		let r = l + TIMELINE_SCROLLER_WIDTH;
		let b = t + this.getHeight();

		return tx >= l && tx <= r && ty >= t && ty <= b;
	}

	CTimeline.prototype.onPreviewStart = function() {
		this.tmpScrollOffset = 0;
		this.setStartTime(0);

		const previewTimings = Asc.editor.WordControl.m_oLogicDocument.previewPlayer.timings;
		this.demoTiming = previewTimings[0]; // effects are smoothed to follow each other
		this.rawDemoTiming = previewTimings[1]; // timing with only effects for preview (without smoothing)

		const oPaneApi = Asc.editor.WordControl.m_oAnimPaneApi;
		oPaneApi.list.Control.recalculateByTiming(this.rawDemoTiming);

		oPaneApi.header.Control.recalculateChildrenLayout();
		oPaneApi.header.OnPaint();
		oPaneApi.timeline.OnPaint();
		oPaneApi.list.OnPaint();
		// this.onUpdate();
	}
	CTimeline.prototype.onPreviewStop = function() {
		this.demoTiming = null;
		this.rawDemoTiming = null;
		this.tmpScrollOffset = null;
		this.setStartTime(0);

		const oPaneApi = Asc.editor.WordControl.m_oAnimPaneApi;
		oPaneApi.list.Control.recalculateByTiming(this.getTiming());

		oPaneApi.header.Control.recalculateChildrenLayout();
		oPaneApi.header.OnPaint();
		oPaneApi.timeline.OnPaint();
		oPaneApi.list.OnPaint();
		// this.onUpdate();
	}
	CTimeline.prototype.onPreview = function(elapsedTicks) {
		if (this.tmpScrollOffset === null) { return; }
		if (!this.demoTiming) { return; }

		const currentlyPlayingDemoEffects = this.getCurrentlyPlayingDemoEffects(elapsedTicks);

		const currentlyPlayingDemoEffect = currentlyPlayingDemoEffects[0]; // first in group
		const correction = (currentlyPlayingDemoEffect)
			? currentlyPlayingDemoEffect.originalNode.getBaseTime() - currentlyPlayingDemoEffect.getBaseTime()
			: 0;
		this.tmpScrollOffset = this.getNewTmpScrollOffset(elapsedTicks, correction);

		const seqList = Asc.editor.WordControl.m_oAnimPaneApi.list.Control.seqList;
		seqList.setCurrentlyPlaying(currentlyPlayingDemoEffects);

		// this.parentControl.drawer == editor.WordControl.m_oAnimPaneApi.timeline
		Asc.editor.WordControl.m_oAnimPaneApi.timeline.OnPaint();
		Asc.editor.WordControl.m_oAnimPaneApi.list.OnPaint();
	}
	CTimeline.prototype.getCurrentlyPlayingDemoEffects = function (elapsedTicks) {
		const demoEffects = this.demoTiming.getRootSequences()[0].getAllEffects();
		const rawDemoEffects = this.rawDemoTiming.getRootSequences()[0].getAllEffects();
		rawDemoEffects.forEach(function (effect, index) {
			effect.originalDemoNode = demoEffects[index];
		});

		// Getting level 3 Time Node Containers
		// Each contains either 'after'-effect or 'click'-effect with mulpiple 'with'-effects
		const lvl3DemoTimingNodes = this.demoTiming.getRootSequences(0)[0].getChildrenTimeNodes()[0].getChildrenTimeNodes();

		// Getting first active level 3 Time Node Container
		// to get currently active demo effect
		let activeDemoEffect = null;
		for (let nodeIndex = lvl3DemoTimingNodes.length - 1; nodeIndex >= 0; --nodeIndex) {
			const node = lvl3DemoTimingNodes[nodeIndex];
			if (node.isActive()) {
				activeDemoEffect = node.getAllAnimEffects()[0];
				break;
			}
		}

		if(activeDemoEffect) {
			// Get index of active demo effect (in array of all raw demo effects)
			let activeDemoEffectIndex;
			for (let nEffect = 0; nEffect < rawDemoEffects.length; nEffect++) {
				if (rawDemoEffects[nEffect].originalNode === activeDemoEffect.originalNode) {
					activeDemoEffectIndex = nEffect;
					break;
				}
			}

			// Get group of active raw demo effects and their corresponding demo effects
			const activeRawDemoEffects = rawDemoEffects[activeDemoEffectIndex].getTimeNodeWithLvl(2).getAllAnimEffects();
			const activeDemoEffects = activeRawDemoEffects.map(function (rawEffect) {
				return rawEffect.originalDemoNode;
			});

			return activeDemoEffects;
		}
		return [];
	};
	CTimeline.prototype.getNewTmpScrollOffset = function (elapsedTicks, correction) {
		const leftLimit = 0;
		const rightLimit = this.getRulerEnd() - this.getZeroShift();

		let newTmpScrollOffset = ms_to_mm(elapsedTicks + correction) - ms_to_mm(this.getStartTime() * 1000);
		if (newTmpScrollOffset < leftLimit) {
			this.setStartTime(0);
			newTmpScrollOffset = 0;
		}
		if (newTmpScrollOffset > rightLimit) {
			const rulerDur = mm_to_ms(this.getRulerEnd() - this.getRulerStart()) / 1000; // seconds
			this.setStartTime(this.getStartTime() + rulerDur / 2);
			newTmpScrollOffset -= ms_to_mm(rulerDur / 2);
		}

		return newTmpScrollOffset;

		function ms_to_mm(nMilliseconds) {
			const index = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline.timeScaleIndex;
			return nMilliseconds * TIME_INTERVALS[index] / TIME_SCALES[index] / 1000;
		}
		function mm_to_ms(nMillimeters) {
			const index = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline.timeScaleIndex;
			return nMillimeters / TIME_INTERVALS[index] * TIME_SCALES[index] * 1000;
		}
	};

	CTimeline.prototype.getRulerStart = function () {
		return this.startButton.getRight();
	};
	CTimeline.prototype.getRulerEnd = function () {
		return this.getWidth() - this.endButton.getWidth();
	};
	CTimeline.prototype.getZeroShift = function () {
		// Returns the value (in millimeters) of the left margin of the start of the ruler
		return this.getRulerStart() + TIMELINE_SCROLLER_WIDTH / 2;
	};

	/*
	 * Functions to convert time to pos and vice versa
	 */
	CTimeline.prototype.getLinearCoeffs = function () {
		//linear relationship x = a*t + b
		var a = TIME_INTERVALS[this.timeScaleIndex] / TIME_SCALES[this.timeScaleIndex];
		var b = this.getZeroShift() - a * this.startTime;
		return { a: a, b: b };
	};
	CTimeline.prototype.timeToPos = function (fTime) {
		//linear relationship x = a*t + b
		var oCoefs = this.getLinearCoeffs();
		return oCoefs.a * fTime + oCoefs.b;
	};
	CTimeline.prototype.posToTime = function (fPos) {
		//linear relationship x = a*t + b 
		var oCoefs = this.getLinearCoeffs();
		return (fPos - oCoefs.b) / oCoefs.a;
	};
	CTimeline.prototype.changeTimelineScale = function (bZoomOut) {
		this.timeScaleIndex = bZoomOut ?
			Math.min(this.timeScaleIndex + 1, TIME_SCALES.length - 1) :
			Math.max(this.timeScaleIndex - 1, 0)

		this.onUpdate();

		// also updating seqList to redraw effect bars
		editor.WordControl.m_oAnimPaneApi.list.Control.seqList.onUpdateSeqList();
	}
	CTimeline.prototype.canHandleEvents = function () {
		return true;
	};
	CTimeline.prototype.recalculateChildrenLayout = function () {
		this.startButton.setLayout(0, 0, TIMELINE_SCROLL_BUTTON_SIZE, TIMELINE_SCROLL_BUTTON_SIZE);
		this.startButton.icon.setLayout(0, 0, TIMELINE_SCROLL_BUTTON_SIZE, TIMELINE_SCROLL_BUTTON_SIZE);

		this.endButton.setLayout(this.getWidth() - TIMELINE_SCROLL_BUTTON_SIZE, 0, TIMELINE_SCROLL_BUTTON_SIZE, TIMELINE_SCROLL_BUTTON_SIZE);
		this.endButton.icon.setLayout(0, 0, TIMELINE_SCROLL_BUTTON_SIZE, TIMELINE_SCROLL_BUTTON_SIZE);

		const currentScrollOffset = this.getScrollOffset()
		if (currentScrollOffset >= this.getMaxScrollOffset()) {
			this.setScrollOffset(currentScrollOffset)
		}
	};
	CTimeline.prototype.onMouseDown = function (e, x, y) {
		if (this.onMouseDownCallback && this.onMouseDownCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseDown.call(this, e, x, y);
	};
	CTimeline.prototype.onMouseMove = function (e, x, y) {
		if (this.onMouseMoveCallback && this.onMouseMoveCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseMove.call(this, e, x, y);
	};
	CTimeline.prototype.onMouseUp = function (e, x, y) {
		if (this.onMouseUpCallback && this.onMouseUpCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseUp.call(this, e, x, y);
	};


	function CSeqListContainer(oDrawer) {
		CTopControl.call(this, oDrawer);
		this.seqList = this.addControl(new CSeqList(this));

		this.onMouseDownCallback = function (event, x, y) {
			if(Asc.editor.asc_IsStartedAnimationPreview()) {
				Asc.editor.asc_StopAnimationPreview();
			}

			if (this.seqList.hit(x, y)) { return }

			this.seqList.forEachAnimItem(function (animItem) { animItem.effect.deselect() })
			Asc.editor.WordControl.m_oLogicDocument.RedrawCurSlide()
			Asc.editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState()
		}
	}

	InitClass(CSeqListContainer, CTopControl, CONTROL_TYPE_SEQ_LIST_CONTAINER);

	CSeqListContainer.prototype.recalculateChildrenLayout = function () {
		this.seqList.setLayout(
			COMMON_LEFT_MARGIN,
			0,
			this.getWidth() - COMMON_LEFT_MARGIN - COMMON_RIGHT_MARGIN,
			this.seqList.getHeight());
		this.seqList.recalculate();
		this.setLayout(0, 0, this.getWidth(), this.seqList.getHeight());
	};
	CSeqListContainer.prototype.recalculateByTiming = function (customTiming) {
		if (!customTiming) { return; }
		this.seqList.recalculateChildren(customTiming);
		this.seqList.recalculateChildrenLayout();
		this.seqList.parentControl.recalculateChildrenLayout();
		this.seqList.parentControl.onUpdate();
		this.seqList.parentControl.drawer.CheckScroll();
	};
	CSeqListContainer.prototype.onScroll = function () {
		this.onUpdate();
	};
	CSeqListContainer.prototype.onMouseWheel = function (e, deltaY, X, Y) {
		return false;
	};
	CSeqListContainer.prototype.onMouseDown = function (e, x, y) {
		if (this.onMouseDownCallback && this.onMouseDownCallback.call(this, e, x, y)) {
			return true;
		}
		return CTopControl.prototype.onMouseDown.call(this, e, x, y);
	};


	function CSeqList(oParentControl) {
		CControlContainer.call(this, oParentControl);
		// this.children - mainSeq, interactiveSeq

		// Tmp field for animItems moving up/down
		this.pressingPoint = null;
		this.nPressedSlot = null;
		this.nCurrentSlot = null;
		this.bTopPart = false;

		this.onMouseDownCallback = function (event, x, y) {
			const oThis = this;
			this.forEachAnimItem(function (animItem, index, groupIndex, seqIndex) {
				const hit = animItem.hit(x, y);
				const hitInEffectBar = animItem.hitInEffectBar(x, y);
				const hitInMenuButton = animItem.contextMenuButton.hit(x, y);
				if (hit && !hitInEffectBar && !hitInMenuButton) {
					oThis.nPressedSlot = index + seqIndex;
					oThis.pressingPoint = { x: x, y: y };
				}
			})
			this.onUpdate();
		}
		this.onMouseMoveCallback = function (event, x, y) {
			if (this.nPressedSlot === null) { return }
			const oThis = this;

			const minRadius = 2; // in millimeters
			const currentShift = Math.sqrt(Math.pow(oThis.pressingPoint.x - x, 2) + Math.pow(oThis.pressingPoint.y - y, 2));
			const bDistancePassed = currentShift >= minRadius;
			if (!bDistancePassed) { return }

			let nLastCheckedSeq = null;
			let hit = null;
			this.forEachAnimItem(function (animItem, index, groupIndex, seqIndex) {
				const seqLabel = oThis.children[seqIndex].label;
				if (seqIndex !== nLastCheckedSeq && seqLabel !== null && seqLabel.hit(x, y)) {
					hit = 'top';
					nLastCheckedSeq = seqIndex;
				} else {
					hit = animItem.hit(x, y);
				}

				if (hit) {
					oThis.nCurrentSlot = index + seqIndex;
					oThis.bTopPart = (hit === 'top');
				}
			})
			this.onUpdate();
		}
		this.onMouseUpCallback = function (event, x, y) {
			if (this.nCurrentSlot !== null && this.nPressedSlot !== null) {
				let moves = this.nCurrentSlot - this.nPressedSlot;
				if (moves > 0 && this.bTopPart) { moves -= 1; }
				if (moves < 0 && !this.bTopPart) { moves += 1; }

				if (moves > 0) {
					if (Asc.editor.asc_canMoveAnimationLater(moves)) {
						Asc.editor.asc_moveAnimationLater(moves)
					}
				}

				if (moves < 0) {
					if (Asc.editor.asc_canMoveAnimationEarlier(Math.abs(moves))) {
						Asc.editor.asc_moveAnimationEarlier(Math.abs(moves))
					}
				}
			}

			this.pressingPoint = null;
			this.nPressedSlot = null;
			this.nCurrentSlot = null;
			this.onUpdate();
		}
	}

	InitClass(CSeqList, CControlContainer, CONTROL_TYPE_SEQ_LIST);

	CSeqList.prototype.recalculateChildren = function (oCustomTiming) {
		this.clear();

		const oTiming = oCustomTiming || this.getTiming();
		if (!oTiming) { return }

		const aAllSeqs = oTiming.getRootSequences();
		let oLastSeqView = null; // Зачем нужна эта переменная?
		for (var nSeq = 0; nSeq < aAllSeqs.length; ++nSeq) {
			const oSeqView = new CAnimSequence(this, aAllSeqs[nSeq]);
			this.addControl(oSeqView);
			oLastSeqView = oSeqView;
		}
	};
	CSeqList.prototype.recalculateChildrenLayout = function () {
		let dLastBottom = 0;

		for (let nChild = 0; nChild < this.children.length; ++nChild) {
			const oSeq = this.children[nChild];
			oSeq.setLayout(0, dLastBottom, this.getWidth(), 0);
			oSeq.recalculate();
			dLastBottom = oSeq.getBottom();
		}
		this.setLayout(this.getLeft(), this.getTop(), this.getWidth(), dLastBottom);
	};
	CSeqList.prototype.draw = function (graphics) {
		if (!CControlContainer.prototype.draw.call(this, graphics)) { return false; }

		// Draw horizontal line when animItem is moving
		if (this.nCurrentSlot !== null) {
			const oThis = this;
			let currentAnimItem;
			this.forEachAnimItem(function(animItem, index, groupIndex, seqIndex) {
				if (index + seqIndex === oThis.nCurrentSlot) currentAnimItem = animItem;
			})

			const transformX = function (x, y) { return graphics.m_oFullTransform.TransformPointX(x, y) + 0.5 >> 0; }
			const transformY = function (x, y) { return graphics.m_oFullTransform.TransformPointY(x, y) + 0.5 >> 0; }

			const yCord = this.bTopPart ? currentAnimItem.bounds.t : currentAnimItem.bounds.b;
			const triangle_diagonal_half = 1; // mm

			const outerLeft = transformX(this.getLeft(), yCord);
			const innerLeft = transformX(this.getLeft() + triangle_diagonal_half, yCord);
			const innerRight = transformX(this.getRight() - triangle_diagonal_half, yCord);
			const outerRight = transformX(this.getRight(), yCord);
			const top = transformY(0, yCord - triangle_diagonal_half);
			const center = transformY(0, yCord);
			const bottom = transformY(0, yCord + triangle_diagonal_half);

			graphics.SaveGrState();
			graphics.RemoveClipRect();

			let ctx = graphics.m_oContext;
			ctx.beginPath();
			ctx.moveTo(innerLeft, center);
			ctx.lineTo(outerLeft, top);
			ctx.lineTo(outerLeft, bottom);
			ctx.lineTo(innerLeft, center);
			ctx.lineTo(innerRight, center);
			ctx.lineTo(outerRight, top);
			ctx.lineTo(outerRight, bottom);
			ctx.lineTo(innerRight, center);
			graphics.df();
			graphics.ds();

			graphics.RestoreGrState();
		}

		// Draw preview line
		const timeline = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline;
		if (timeline.tmpScrollOffset !== null) {
			graphics.SaveGrState();
			graphics.RemoveClipRect();

			let sColor = AscCommon.GlobalSkin.AnimPaneTimelineScrollerOutline;
			let oColor = AscCommon.RgbaHexToRGBA(sColor);
			graphics.p_color(oColor.R, oColor.G, oColor.B, 255);
			const xCord = timeline.getLeft() + timeline.getZeroShift() + timeline.tmpScrollOffset;
			const height = this.parentControl.drawer.GetHeight() + editor.WordControl.m_oAnimPaneApi.list.Scroll * g_dKoef_pix_to_mm;
			graphics.drawVerLine(1, xCord, this.getTop(), this.getTop() + height, this.getPenWidth(graphics));

			graphics.RestoreGrState();
		}

		return true;
	}
	CSeqList.prototype.onMouseDown = function (e, x, y) {
		if (this.onMouseDownCallback) this.onMouseDownCallback(e, x, y);
		return CControlContainer.prototype.onMouseDown.call(this, e, x, y);
	};
	CSeqList.prototype.onMouseMove = function (e, x, y) {
		if (this.onMouseMoveCallback) this.onMouseMoveCallback(e, x, y);
		return CControlContainer.prototype.onMouseMove.call(this, e, x, y);
	};
	CSeqList.prototype.onMouseUp = function (e, x, y) {
		if (this.onMouseUpCallback) this.onMouseUpCallback(e, x, y);
		return CControlContainer.prototype.onMouseUp.call(this, e, x, y);
	};

	CSeqList.prototype.onUpdateSeqList = function () {
		if (Asc.editor.WordControl.m_oAnimPaneApi.list.Control) {
			this.onUpdate()
		}
	}

	CSeqList.prototype.checkCachedTexture = function (graphics) {
		var dGraphicsScale = graphics.m_oCoordTransform.sx;
		if (this.cachedCanvas) {
			var dScale = this.cachedCanvas.scale;
			if (AscFormat.fApproxEqual(dScale, dGraphicsScale)) {
				return this.cachedCanvas;
			}
		}
		this.bDrawTexture = true;
		var oBaseTexture = this.getAnimTexture(dGraphicsScale);
		if (oBaseTexture) {
			this.cachedCanvas = new CAnimTexture(this, oBaseTexture.canvas, oBaseTexture.scale, oBaseTexture.x, oBaseTexture.y);
		}
		else {
			this.cachedCanvas = null;
		}
		this.bDrawTexture = false;
		return this.cachedCanvas;
	};
	CSeqList.prototype.clearCachedTexture = function () {
		if (this.cachedCanvas) {
			this.cachedCanvas = null;
		}
	};
	CSeqList.prototype.setCurrentlyPlaying = function (demoEffects) {
		if (!demoEffects) { return; }

		const originalEffects = demoEffects.map(
			function (demoEffect) {
				return demoEffect.originalNode;
			}
		);
		this.forEachAnimItem(
			function (animItem) {
				animItem.isCurrentlyPlaying = (originalEffects.indexOf(animItem.effect.originalNode) > -1);
			}
		)
	};
	CSeqList.prototype.forEachAnimItem = function (callback) {
		// У счетчиков сквозная нумерация
		let seqCounter = 0;
		let groupCounter = 0;
		let itemCounter = 0;
		this.children.forEach(function (seq) {
			seq.animGroups.forEach(function (group) {
				group.children.forEach(function (animItem) {
					callback(animItem, itemCounter, groupCounter, seqCounter);
					itemCounter++;
				})
				groupCounter++;
			})
			seqCounter++;
		})
	};


	// mainSeq or interactiveSeq
	function CAnimSequence(oParentControl, oSeq) {
		CControlContainer.call(this, oParentControl);
		this.seq = oSeq;
		this.label = null; //this.addControl(new CLabel(this, "seq"));
		this.animGroups = [];
	}

	InitClass(CAnimSequence, CControlContainer, CONTROL_TYPE_ANIM_SEQ);

	CAnimSequence.prototype.getSeq = function () {
		return this.seq;
	};
	CAnimSequence.prototype.recalculateChildren = function () {
		this.clear();

		let sLabel = this.seq.getLabel();
		if (typeof sLabel === "string" && sLabel.length > 0) {
			this.label = this.addControl(new CLabel(this, sLabel, 9, true, AscCommon.align_Left));
		}

		const aAllEffects = this.seq.getAllEffects();
		const animGroups = AscFormat.groupBy(aAllEffects, function (effect) { return effect.getIndexInSequence(); })

		for (let indexInSequence in animGroups) {
			const oAnimGroup = this.addControl(new CAnimGroup(this, animGroups[indexInSequence]));
			this.animGroups[this.animGroups.length] = oAnimGroup;
		}
	};
	CAnimSequence.prototype.recalculateChildrenLayout = function () {
		let dCurY = 0;
		if (this.label) {
			dCurY += SEQ_LABEL_MARGIN;
			this.label.setLayout(SEQ_LABEL_MARGIN, dCurY, this.getWidth(), SEQ_LABEL_HEIGHT);
			this.label.recalculate();
			dCurY += this.label.getHeight() + SEQ_LABEL_MARGIN;
		}
		for (let nGroup = 0; nGroup < this.animGroups.length; ++nGroup) {
			this.animGroups[nGroup].setLayout(0, dCurY, this.getWidth(), 0);
			this.animGroups[nGroup].recalculate();
			dCurY += this.animGroups[nGroup].getHeight();
		}
		this.setLayout(this.getLeft(), this.getTop(), this.getWidth(), dCurY);
	};


	function CAnimGroup(oParentControl, aAllGroupEffects) {
		CControlContainer.call(this, oParentControl);
		this.effects = aAllGroupEffects;
	}

	InitClass(CAnimGroup, CControlContainer, CONTROL_TYPE_ANIM_GROUP_LIST);

	CAnimGroup.prototype.getSeq = function () {
		return this.parentControl.getSeq();
	};
	CAnimGroup.prototype.recalculateChildren = function () {
		this.clear();

		for (let nCurEffect = this.effects.length - 1; nCurEffect >= 0; --nCurEffect) {
			const oItem = new CAnimItem(this, this.effects[nCurEffect]);
			this.addControl(oItem);
		}
	};
	CAnimGroup.prototype.recalculateChildrenLayout = function () {
		let dLastBottom = 0;

		for (let nChild = 0; nChild < this.children.length; ++nChild) {
			let oChild = this.children[nChild];
			oChild.setLayout(0, dLastBottom, this.getWidth(), ANIM_ITEM_HEIGHT);
			oChild.recalculate();
			dLastBottom = oChild.getBottom();
		}
		this.setLayout(this.getLeft(), this.getTop(), this.getWidth(), dLastBottom);
	};
	CAnimGroup.prototype.draw = function(graphics) {
		if (this.isHidden()) { return false; }
		if (!this.checkUpdateRect(graphics.updatedRect)) { return false; }
		if (!CControlContainer.prototype.draw.call(this, graphics)) { return false; }

		// Connection lines drawing
		let bShouldDraw = false;
		this.effects.some(function (effect) {
			if (effect.isSelected()) return bShouldDraw = true;
		})
		if (!bShouldDraw) { return }

		const oThis = this;
		const timeline = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline;
		const timelineShift = ms_to_mm(timeline.getStartTime() * 1000);

		let afterItems = []
		this.children.forEach(function (animItem) {
			if (animItem.effect.isAfterEffect()) afterItems[afterItems.length] = animItem;
		})
		if (afterItems.length === 0) { return }

		graphics.SaveGrState();
		graphics.AddClipRect(
			afterItems[0].getLeftBorder(),
			oThis.parentControl.getTop(),
			afterItems[0].getRightBorder() - afterItems[0].getLeftBorder(),
			oThis.parentControl.getBottom() - oThis.parentControl.getTop()
		);

		for (let i = 0; i < afterItems.length; i++) {
			const animItem = afterItems[i];
			
			if (animItem.effect === this.effects[this.effects.length - 1]) {
				// effects in group are arranged backwards
				continue;
			}

			const align = 0;
			const x = ms_to_mm(animItem.effect.getBaseTime()) + animItem.getLeftBorder() - timelineShift;
			let top = afterItems[i-1] ? oThis.getTop() + afterItems[i-1].getTop() : oThis.getTop();
			top += oThis.parentControl.getTop();
			let bottom = afterItems[i+1] ? oThis.getTop() + afterItems[i+1].getTop() : oThis.getBottom();
			bottom += oThis.parentControl.getTop();

			graphics.drawVerLine(align, x, top, bottom, oThis.getPenWidth(graphics));
		}

		graphics.RestoreGrState();

		function ms_to_mm(nMilliseconds) {
			const index = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline.timeScaleIndex;
			return nMilliseconds * TIME_INTERVALS[index] / TIME_SCALES[index] / 1000;
		}
	}


	function CAnimItem(oParentControl, oEffect) {
		CControlContainer.call(this, oParentControl);
		this.effect = oEffect;

		if (this.effect.isClickEffect() || !this.effect.getPreviousEffect()) {
			this.indexLabel = this.addControl(new CLabel(this, this.effect.getIndexInSequence() + "", INDEX_LABEL_FONTSIZE, false, AscCommon.align_Center))
		}

		// Event type image
		const eventImg = this.getEventImage();
		this.eventTypeImage = this.addControl(new CImageControl(this, eventImg.src, eventImg.width, eventImg.height));

		// Effect type image
		const effectImg = this.getEffectImage();
		this.effectTypeImage = this.addControl(new CImageControl(this, effectImg.src, effectImg.width, effectImg.height));

		this.effectLabel = this.addControl(new CLabel(this, this.getEffectLabelText(), EFFECT_LABEL_FONTSIZE, false, AscCommon.align_Left));
		this.contextMenuButton = this.addControl(new CButton(this, function (e, x, y) {
			if (this.hit(x, y) && (e.Button === AscCommon.g_mouse_button_left)) {
				this.pressedFlag = true;
			}
		}, null, showContextMenu));
		this.contextMenuButton.icon = this.contextMenuButton.addControl(new CImageControl(
			this.contextMenuButton,
			AscCommon.GlobalSkin.type == 'light' ? menuButtonIcon_dark : menuButtonIcon_light,
			10 * AscCommon.g_dKoef_pix_to_mm, 2 * AscCommon.g_dKoef_pix_to_mm /* 10x2 svg image */
		));

		this.contextMenuButton.sendContextMenuEvent = function (customX, customY) {
			const coords = editor.WordControl.m_oDrawingDocument.ConvertAnimPaneCoordsToCursor(
				AscFormat.isRealNumber(customX) ? customX : this.bounds.l,
				(AscFormat.isRealNumber(customY) ? customY : this.bounds.t) + HEADER_HEIGHT - editor.WordControl.m_oAnimPaneApi.list.Scroll * g_dKoef_pix_to_mm
			);

			const data = new AscCommonSlide.CContextMenuData()
			data.Type = Asc.c_oAscContextMenuTypes.AnimEffect;
			data.X_abs = coords.X;
			data.Y_abs = coords.Y;
			if (!AscFormat.isRealNumber(customX)) {
				data.ButtonWidth = (this.bounds.r - this.bounds.l) * g_dKoef_mm_to_pix;
			}
			if (!AscFormat.isRealNumber(customY)) {
				data.ButtonHeight = (this.bounds.b - this.bounds.t) * g_dKoef_mm_to_pix;
			}
			data.EffectStartType = this.parentControl.effect.getNodeType();

			editor.sync_ContextMenuCallback(data);
		}

		function showContextMenu(e, x, y) {
			if (this.hit(x, y) && !this.isHidden()) {
				if (this.pressedFlag) {
					this.sendContextMenuEvent();
					this.pressedFlag = false;
				}
			}
		}

		// Temp fields for effect bar movement
		this.tmpDelay = null;
		this.tmpDuration = null;
		this.tmpRepeatCount = null;

		// Callback functions for effect bar events
		this.onMouseDownCallback = function (event, x, y) {
			if (!this.hit(x, y)) { return }

			if (!this.contextMenuButton.hit(x, y)) {
				this.updateSelectState(event);
			}

			const hitRes = this.hitInEffectBar(x, y);
			if (hitRes) {
				this.hitResult = hitRes;
				this.tmpDelay = this.getDelay();
				this.tmpDuration = this.getDuration();

				if (this.effect.isUntilEffect() && hitRes.type === 'right') {
					this.tmpRepeatCount = this.getRepeatCount();
					this.initialTmpRepeatCount = this.tmpRepeatCount;
				}

				this.onUpdate();
			}
		}
		this.onMouseMoveCallback = function (event, x, y) {
			if (this.hit(x, y)) {
				this.updateCursorType(x, y);
			}

			if (!this.hitResult) { return }
			this.handleMovement(x, y);
			this.handleScrollCondition(x, y);

			this.onUpdate();
		}
		this.onMouseUpCallback = function (event, x, y) {
			if (this.hit(x, y)) {
				if (event.Button === AscCommon.g_mouse_button_right) {
					this.contextMenuButton.sendContextMenuEvent(x, y);
				}
			}

			if (!this.hitResult) { return }
			this.setNewEffectParams(this.tmpDelay, this.tmpDuration, this.tmpRepeatCount);
			this.hitResult = this.tmpDelay = this.tmpDuration = this.tmpRepeatCount = null;

			this.onUpdate()
		}
	}

	InitClass(CAnimItem, CControlContainer, CONTROL_TYPE_ANIM_ITEM);

	CAnimItem.prototype.recalculateChildrenLayout = function () {
		if (this.indexLabel) this.indexLabel.setLayout(0, 0, INDEX_LABEL_WIDTH, ANIM_ITEM_HEIGHT)

		this.eventTypeImage.setLayout(INDEX_LABEL_WIDTH, 0, EVENT_TYPE_ICON_SIZE, EVENT_TYPE_ICON_SIZE);
		this.effectTypeImage.src = this.getEffectImage().src;
		this.effectTypeImage.setLayout(this.eventTypeImage.getRight(), 0, EFFECT_TYPE_ICON_SIZE, EFFECT_TYPE_ICON_SIZE);

		const zeroPos = COMMON_LEFT_MARGIN + SCALE_BUTTON_LEFT_MARGIN + SCALE_BUTTON_WIDTH + TIMELINE_SCROLL_LEFT_MARGIN + TIMELINE_SCROLL_BUTTON_SIZE;
		const labelWidth = zeroPos - COMMON_LEFT_MARGIN - (INDEX_LABEL_WIDTH + EVENT_TYPE_ICON_SIZE + EFFECT_TYPE_ICON_SIZE);
		const gap = (ANIM_ITEM_HEIGHT - EFFECT_BAR_HEIGHT) / 2;
		this.effectLabel.setLayout(this.effectTypeImage.getRight(), gap, labelWidth, EFFECT_BAR_HEIGHT);

		const menuBtnGap = (ANIM_ITEM_HEIGHT - MENU_BUTTON_SIZE) / 2;
		const menuBtnLeft = this.getRight() - MENU_BUTTON_SIZE - menuBtnGap;
		this.contextMenuButton.setLayout(menuBtnLeft, menuBtnGap, MENU_BUTTON_SIZE, MENU_BUTTON_SIZE);

		this.contextMenuButton.icon.src = AscCommon.GlobalSkin.type == 'light' ? menuButtonIcon_dark : menuButtonIcon_light;
		this.contextMenuButton.icon.setLayout(0, 0, MENU_BUTTON_SIZE, MENU_BUTTON_SIZE);
	};
	CAnimItem.prototype.getEventImage = function () {
		let eventImg = {};
		if (this.effect.isClickEffect()) {
			eventImg.src = clickEffectIcon;
			eventImg.width = 11 * AscCommon.g_dKoef_pix_to_mm;
			eventImg.height = 16 * AscCommon.g_dKoef_pix_to_mm;
		}
		if (this.effect.isAfterEffect()) {
			eventImg.src = afterEffectIcon;
			eventImg.width = 16 * AscCommon.g_dKoef_pix_to_mm;
			eventImg.height = 16 * AscCommon.g_dKoef_pix_to_mm;
		}
		return eventImg;
	};
	CAnimItem.prototype.getEffectImage = function () {
		let effectImg = {};
		if (this.effect.cTn.presetClass === AscFormat.PRESET_CLASS_ENTR) {
			effectImg.src = entrEffectIcon;
			effectImg.width = 20 * AscCommon.g_dKoef_pix_to_mm;
			effectImg.height = 20 * AscCommon.g_dKoef_pix_to_mm;
		}
		if (this.effect.cTn.presetClass === AscFormat.PRESET_CLASS_EMPH) {
			effectImg.src = emphEffectIcon;
			effectImg.width = 20 * AscCommon.g_dKoef_pix_to_mm;
			effectImg.height = 20 * AscCommon.g_dKoef_pix_to_mm;
		}
		if (this.effect.cTn.presetClass === AscFormat.PRESET_CLASS_EXIT) {
			effectImg.src = exitEffectIcon;
			effectImg.width = 20 * AscCommon.g_dKoef_pix_to_mm;
			effectImg.height = 20 * AscCommon.g_dKoef_pix_to_mm;
		}
		if (this.effect.cTn.presetClass === AscFormat.PRESET_CLASS_PATH) {
			effectImg.src = AscCommon.GlobalSkin.type === 'light' ? pathEffectIcon_dark : pathEffectIcon_light;
			effectImg.width = 20 * AscCommon.g_dKoef_pix_to_mm;
			effectImg.height = 20 * AscCommon.g_dKoef_pix_to_mm;
		}
		return effectImg;
	};
	CAnimItem.prototype.updateSelectState = function (event) {
		const oThis = this;
		const seqList = Asc.editor.WordControl.m_oAnimPaneApi.list.Control.seqList;
		if (event.Button === AscCommon.g_mouse_button_right && !oThis.effect.isSelected()) {
			seqList.forEachAnimItem(function (animItem) {
				animItem.effect === oThis.effect ? animItem.effect.select() : animItem.effect.deselect();
			})
		}
		if (event.Button === AscCommon.g_mouse_button_left) {
			if (event.CtrlKey) {
				oThis.effect.toggleSelect();
			} else {
				seqList.forEachAnimItem(function (animItem) {
					animItem.effect === oThis.effect ? animItem.effect.select() : animItem.effect.deselect();
				})
			}
		}
		Asc.editor.WordControl.m_oLogicDocument.RedrawCurSlide();
		Asc.editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
	}
	CAnimItem.prototype.updateCursorType = function (x, y) {
		const cursorType = this.getNewCursorType(x, y);
		const mouseMoveData = this.getMouseMoveData(x, y);

		const animPane = Asc.editor.WordControl.m_oAnimPaneApi;
		animPane.SetCursorType(cursorType, mouseMoveData);
		animPane.sentMouseMoveData = mouseMoveData;
	}
	CAnimItem.prototype.getNewCursorType = function (x, y) {
		const isVerticalDrag = null !== editor.WordControl.m_oAnimPaneApi.list.Control.seqList.nPressedSlot;
		if (isVerticalDrag) {
			return 'grabbing';
		}

		let draggingAnimItem;
		editor.WordControl.m_oAnimPaneApi.list.Control.seqList.forEachAnimItem(function (animItem) {
			if (animItem.hitResult) {
				draggingAnimItem = animItem;
			}
		});

		const hitRes = draggingAnimItem
			? draggingAnimItem.hitResult
			: (this.hitResult || this.hitInEffectBar(x, y));

		const cursorTypes = {
			'left': 'col-resize',
			'right': 'col-resize',
			'partition': 'col-resize',
			'center': 'ew-resize'
		};
		const cursorType = hitRes
			? cursorTypes[hitRes.type]
			: this.contextMenuButton.hit(x, y) ? 'default' : 'ns-resize';
		return cursorType;
	};
	CAnimItem.prototype.getMouseMoveData = function (x, y) {
		const coords = editor.WordControl.m_oDrawingDocument.ConvertAnimPaneCoordsToCursor(
			x, y + HEADER_HEIGHT - editor.WordControl.m_oAnimPaneApi.list.Scroll * g_dKoef_pix_to_mm
		);

		const mouseMoveData = new CMouseMoveData();
		mouseMoveData.X_abs = coords.X;
		mouseMoveData.Y_abs = coords.Y;

		const isVerticalDrag = null !== editor.WordControl.m_oAnimPaneApi.list.Control.seqList.nPressedSlot;
		if (!this.contextMenuButton.hit(x, y) && !isVerticalDrag) {
			mouseMoveData.Type = Asc.c_oAscMouseMoveDataTypes.EffectInfo;
			const tooltipInfo = this.getInfoForTooltip(x, y);
			if (typeof tooltipInfo === 'string') {
				mouseMoveData.EffectText = tooltipInfo;
			} else {
				mouseMoveData.EffectDescription = tooltipInfo;
			}
		}

		return mouseMoveData;
	};
	CAnimItem.prototype.getInfoForTooltip = function (x, y) {
		// If there is a pressed animItem - we take the information from it,
		// otherwise - from this one (literally 'this' xd)
		let currentAnimItem = this;
		const seqList = editor.WordControl.m_oAnimPaneApi.list.Control.seqList
		seqList.forEachAnimItem(function (animItem) {
			if (animItem.hitResult) { currentAnimItem = animItem; }
		})

		const templateStrings = {
			startTime: AscCommon.translateManager.getValue('Start: ${0}s'),
			endTime: AscCommon.translateManager.getValue('End: ${0}s'),
			loopTime: AscCommon.translateManager.getValue('Loop: ${0}s'),
		};

		// When dragging (when animItem's bar is pressed)
		if (currentAnimItem.hitResult) {
			let time;
			switch (currentAnimItem.hitResult.type) {
				case 'center':
					time = currentAnimItem.getDelay() / 1000;
					return templateStrings.startTime.replace('${0}', time.toFixed(1));
				case 'left':
					time = currentAnimItem.getDelay() / 1000;
					return templateStrings.startTime.replace('${0}', time.toFixed(1));
				case 'right':
					time = currentAnimItem.getDelay() / 1000 + currentAnimItem.getDuration() / 1000;
					return templateStrings.endTime.replace('${0}', time.toFixed(1));
				case 'partition':
					time = (currentAnimItem.getDuration() / 1000);
					return templateStrings.loopTime.replace('${0}', time.toFixed(1));
			}
		}

		if (currentAnimItem.hitInEffectBar(x, y)) {
			const startTime = (currentAnimItem.getDelay() / 1000).toFixed(1);
			const endTime = ((currentAnimItem.getDelay() + currentAnimItem.getDuration()) / 1000).toFixed(1);
			const result = [
				templateStrings.startTime.replace('${0}', startTime),
				templateStrings.endTime.replace('${0}', endTime),
			];
			return result.join(', ');
		} else {
			// Belongs to [AscFormat.NODE_TYPE_AFTEREFFECT, AscFormat.NODE_TYPE_CLICKEFFECT, AscFormat.NODE_TYPE_WITHEFFECT]
			const eventType = currentAnimItem.effect.getNodeType();
			// Belongs to [AscFormat.PRESET_CLASS_ENTR, AscFormat.PRESET_CLASS_EMPH, AscFormat.PRESET_CLASS_EXIT, AscFormat.PRESET_CLASS_PATH]
			const presetClass = currentAnimItem.effect.cTn.presetClass;
			// Belongs to presets of animation - [AscFormat.ENTRANCE_APPEAR, ..., AscFormat.EMPHASIS_FILL_COLOR, ...]
			const presetId = currentAnimItem.effect.cTn.presetID;
			const shapeName = currentAnimItem.effect.getObjectName();
			return [eventType, presetClass, presetId, shapeName];
		}
	}
	CAnimItem.prototype.getEffectLabelText = function () {
		const objectName = this.effect.getObjectName();
		const objectText = this.effect.getObjectText();
		return objectText ? (objectName + ': ' + objectText) : objectName;
	};
	CAnimItem.prototype.handleScrollCondition = function (x, y) {
		const leftBorder = this.getLeftBorder();
		const rightBorder = this.getRightBorder();

		const timeline = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline;
		if (x <= leftBorder || x >= rightBorder) {
			if (!timeline.isOnScroll()) {
				let scrollStep = timeline.getWidth() * SCROLL_STEP / 10;
				scrollStep = x <= leftBorder ? -scrollStep : scrollStep;
				let scrollTimerDelay = 0;
				let scrollTimerInterval = 50;
				timeline.startScroll(scrollStep, scrollTimerDelay, scrollTimerInterval);
			}
		} else timeline.endScroll();
	}
	CAnimItem.prototype.handleMovement = function (x, y) {
		const timeline = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline;
		const timelineShift = this.ms_to_mm(timeline.getStartTime() * 1000);
		const repeats = this.getRepeatCount() / 1000;

		let pointOfLanding = x - this.getLeftBorder() + timelineShift;

		if (this.hitResult.type === 'right') {
			if (this.effect.isUntilEffect()) {
				const pointOfContact = this.ms_to_mm(this.effect.getFullDelay() + this.effect.asc_getDuration() * this.initialTmpRepeatCount / 1000);
				let diff = this.mm_to_ms(pointOfLanding - pointOfContact);

				const newTmpRepeatCount = this.initialTmpRepeatCount + diff / (this.effect.asc_getDuration() / 1000);
				this.tmpRepeatCount = Math.max(newTmpRepeatCount, MIN_ALLOWED_REPEAT_COUNT);
			} else {
				const pointOfContact = this.ms_to_mm(this.effect.getFullDelay() + this.effect.asc_getDuration() * repeats);
				let diff = this.mm_to_ms(pointOfLanding - pointOfContact);

				const newTmpDuration = this.effect.asc_getDuration() + diff / repeats;
				this.tmpDuration = Math.max(MIN_ALLOWED_DURATION, newTmpDuration);
			}
		}

		if (this.hitResult.type === 'left') {
			const pointOfContact = this.ms_to_mm(this.effect.getFullDelay());
			const basePoint = this.ms_to_mm(this.effect.getBaseTime());
			pointOfLanding = Math.max(pointOfLanding, basePoint);
			const diff = this.mm_to_ms(pointOfLanding - pointOfContact);

			const newTmpDuration = this.effect.asc_getDuration() - diff / repeats;
			const newTmpDelay = this.effect.getFullDelay() + diff;

			const maxNewTmpDuration = this.effect.getFullDelay() / repeats + this.effect.asc_getDuration();
			const maxNewTmpDelay = this.effect.getFullDelay() + (this.effect.asc_getDuration() - MIN_ALLOWED_DURATION) * repeats;

			if (this.effect.isUntilEffect()) {
				this.tmpDelay = Math.max(newTmpDelay, this.effect.getBaseTime());
			} else {
				this.tmpDuration = Math.min(Math.max(newTmpDuration, MIN_ALLOWED_DURATION), maxNewTmpDuration);
				this.tmpDelay = Math.min(Math.max(newTmpDelay, this.effect.getBaseTime()), maxNewTmpDelay);
			}
		}

		if (this.hitResult.type === 'center') {
			const pointOfContact = this.ms_to_mm(this.effect.getFullDelay()) + this.hitResult.offset;
			const diff = this.mm_to_ms(pointOfLanding - pointOfContact);

			const newTmpDelay = this.effect.getFullDelay() + diff;
			this.tmpDelay = Math.max(newTmpDelay, this.effect.getBaseTime());
		}

		if (this.hitResult.type === 'partition') {
			const pointOfContact = this.ms_to_mm(this.effect.getFullDelay() + this.effect.asc_getDuration() * this.hitResult.index);
			const diff = this.mm_to_ms(pointOfLanding - pointOfContact);

			const newTmpDuration = this.effect.asc_getDuration() + diff / this.hitResult.index;
			this.tmpDuration = Math.max(MIN_ALLOWED_DURATION, newTmpDuration);
		}
	}
	CAnimItem.prototype.handleTimelineScroll = function (step) {
		if (!this.hitResult) { return }

		const repeats = this.getRepeatCount() / 1000;
		const diff = this.mm_to_ms(step);
		let newTmpDelay;
		let newTmpDuration;
		let newTmpRepeatCount;
		switch (this.hitResult.type) {
			case 'center':
				newTmpDelay = this.tmpDelay + diff;
				this.tmpDelay = Math.max(newTmpDelay, this.effect.getBaseTime());
				break;

			case 'right':
				if (this.effect.isUntilEffect()) {
					newTmpRepeatCount = this.tmpRepeatCount + diff / (this.effect.asc_getDuration() / 1000);
					this.tmpRepeatCount = Math.max(newTmpRepeatCount, MIN_ALLOWED_REPEAT_COUNT);
				} else {
					newTmpDuration = this.tmpDuration + diff / repeats;
					this.tmpDuration = Math.max(MIN_ALLOWED_DURATION, newTmpDuration);
				}
				break;

			case 'left':
				newTmpDuration = this.tmpDuration - diff / repeats;
				newTmpDelay = this.tmpDelay + diff;

				const maxNewTmpDuration = (this.effect.getFullDelay() - this.effect.getBaseTime()) / repeats + this.effect.asc_getDuration();
				const maxNewTmpDelay = this.effect.getFullDelay() + (this.effect.asc_getDuration() - MIN_ALLOWED_DURATION) * repeats;

				if (this.effect.isUntilEffect()) {
					this.tmpDelay = Math.max(newTmpDelay, this.effect.getBaseTime());
				} else {
					this.tmpDuration = Math.min(Math.max(newTmpDuration, MIN_ALLOWED_DURATION), maxNewTmpDuration);
					this.tmpDelay = Math.min(Math.max(newTmpDelay, this.effect.getBaseTime()), maxNewTmpDelay);
				}
				break;

			case 'partition':
				newTmpDuration = this.tmpDuration + diff / this.hitResult.index;
				this.tmpDuration = Math.max(MIN_ALLOWED_DURATION, newTmpDuration);
				break;
		}

		this.onUpdate();
	}

	CAnimItem.prototype.ms_to_mm = function (nMilliseconds) {
		if (nMilliseconds === null || nMilliseconds === undefined) { return }

		const index = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline.timeScaleIndex;
		return nMilliseconds * TIME_INTERVALS[index] / TIME_SCALES[index] / 1000;
	};
	CAnimItem.prototype.mm_to_ms = function (nMillimeters) {
		if (nMillimeters === null || nMillimeters === undefined) { return }

		const index = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline.timeScaleIndex;
		return nMillimeters / TIME_INTERVALS[index] * TIME_SCALES[index] * 1000;
	};

	CAnimItem.prototype.getDelay = function () {
		return this.tmpDelay !== null ? this.tmpDelay : this.effect.getFullDelay();
	}
	CAnimItem.prototype.getDuration = function () {
		return this.tmpDuration !== null ? this.tmpDuration : this.effect.asc_getDuration()
	}
	CAnimItem.prototype.getRepeatCount = function () {
		if (this.tmpRepeatCount !== null) { return this.tmpRepeatCount; }
		else if (this.effect.asc_getRepeatCount() > 0) { return this.effect.asc_getRepeatCount(); }
		else {
			const bounds = this.getEffectBarBounds();
			const width = bounds.r - bounds.l;
			const totalWidth = this.getRightBorder() - bounds.l;
			return (totalWidth / width * 1000) >> 0; // approximate repeat counter
		}
	}

	CAnimItem.prototype.getLeftBorder = function () {
		const timeline = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline;
		return timeline.getLeft() + timeline.getZeroShift();
	}
	CAnimItem.prototype.getRightBorder = function () {
		const timeline = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline;
		return timeline.getRight() - TIMELINE_SCROLL_BUTTON_SIZE;
	}
	CAnimItem.prototype.getEffectBarBounds = function () {
		const timeline = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control.timeline;
		const timelineShift = timeline.getStartTime() * 1000;

		let l = this.ms_to_mm(this.getDelay()) + this.getLeftBorder() - this.ms_to_mm(timelineShift);

		let r = l + this.ms_to_mm(this.getDuration());

		let t = this.bounds.t + (ANIM_ITEM_HEIGHT - EFFECT_BAR_HEIGHT) / 2;

		let b = t + EFFECT_BAR_HEIGHT;

		if (this.effect.isInstantEffect()) {
			return { l: l, r: l + EFFECT_BAR_HEIGHT, t: t, b: b }
		}

		return { l: l, r: r, t: t, b: b }
	};

	CAnimItem.prototype.draw = function drawEffectBar(graphics) {
		const timelineContainer = Asc.editor.WordControl.m_oAnimPaneApi.timeline.Control
		if (!timelineContainer) { return }

		this.effect.isSelected() ? this.contextMenuButton.show() : this.contextMenuButton.hide();

		if (!CControlContainer.prototype.draw.call(this, graphics)) { return false }
		if (this.isHidden()) { return false }
		if (!this.checkUpdateRect(graphics.updatedRect)) { return false }

		graphics.SaveGrState();

		const clipL = this.getLeftBorder();
		const clipT = this.bounds.t;
		const clipW = this.getRightBorder() - clipL;
		const clipH = this.bounds.b - this.bounds.t;
		graphics.AddClipRect(clipL, clipT, clipW, clipH);

		const oSkin = AscCommon.GlobalSkin;
		let sFillColor, sOutlineColor;
		switch (this.effect.cTn.presetClass) {
			case AscFormat.PRESET_CLASS_ENTR:
				sFillColor = oSkin.AnimPaneEffectBarFillEntrance;
				sOutlineColor = oSkin.AnimPaneEffectBarOutlineEntrance;
				break;

			case AscFormat.PRESET_CLASS_EMPH:
				sFillColor = oSkin.AnimPaneEffectBarFillEmphasis;
				sOutlineColor = oSkin.AnimPaneEffectBarOutlineEmphasis;
				break;

			case AscFormat.PRESET_CLASS_EXIT:
				sFillColor = oSkin.AnimPaneEffectBarFillExit;
				sOutlineColor = oSkin.AnimPaneEffectBarOutlineExit;
				break;

			case AscFormat.PRESET_CLASS_PATH:
				sFillColor = oSkin.AnimPaneEffectBarFillPath;
				sOutlineColor = oSkin.AnimPaneEffectBarOutlinePath;
				break;

			default:
				sFillColor = '#A0A0A0';
				sOutlineColor = '#404040';
		}

		// hex to rgba
		const oFillColorRGBA = AscCommon.RgbaHexToRGBA(sFillColor);
		const oOutlineColorRGBA = AscCommon.RgbaHexToRGBA(sOutlineColor);

		// rgba to CShapeColor
		let oFillColor = new AscFormat.CShapeColor(oFillColorRGBA.R, oFillColorRGBA.G, oFillColorRGBA.B);
		let oOutlineColor = new AscFormat.CShapeColor(oOutlineColorRGBA.R, oOutlineColorRGBA.G, oOutlineColorRGBA.B);

		// change brightness of CShapeColor during demo preview
		if (Asc.editor.asc_IsStartedAnimationPreview()) {
			if (!this.isCurrentlyPlaying) {
				oFillColor = oFillColor.getColorData(0.4);
				oOutlineColor = oOutlineColor.getColorData(0.4);
			}
		}

		graphics.b_color1(oFillColor.r, oFillColor.g, oFillColor.b, 255);
		graphics.p_color(oOutlineColor.r, oOutlineColor.g, oOutlineColor.b, 255);

		const bounds = this.getEffectBarBounds();
		if (this.effect.isInstantEffect()) {
			// In case we need to draw a triangle

			let transform = graphics.m_oFullTransform;
			let left = (transform.TransformPointX(bounds.l, bounds.t) + 0.5) >> 0;
			let top = (transform.TransformPointY(bounds.l, bounds.t) + 0.5) >> 0;
			let right = (transform.TransformPointX(bounds.r, bounds.t) + 0.5) >> 0;
			let bottom = (transform.TransformPointY(bounds.l, bounds.b) + 0.5) >> 0;

			let ctx = graphics.m_oContext;
			ctx.beginPath();
			ctx.moveTo(left, top);
			ctx.lineTo(left + 5, top);
			ctx.lineTo(right, top + (bottom - top) / 2);
			ctx.lineTo(left + 5, bottom);
			ctx.lineTo(left, bottom);
			ctx.lineTo(left, top);
			graphics.df();
			graphics.ds();
		} else {
			let repeats;
			if (this.effect.isUntilEffect() && this.tmpRepeatCount === null) {
				// In case we need to draw an infinite bar with an arrow

				const barWidth = Math.max(this.getRightBorder() - bounds.l - EFFECT_BAR_HEIGHT, this.ms_to_mm(MIN_ALLOWED_DURATION));
				// repeats = barWidth / (bounds.r - bounds.l);
				repeats = this.getRepeatCount() / 1000;

				let transform = graphics.m_oFullTransform;
				let left = (transform.TransformPointX(bounds.l, bounds.t) + 0.5) >> 0;
				let top = (transform.TransformPointY(bounds.l, bounds.t) + 0.5) >> 0;
				let right = (transform.TransformPointX(bounds.l + barWidth, bounds.t) + 0.5) >> 0;
				let bottom = (transform.TransformPointY(bounds.l, bounds.b) + 0.5) >> 0;

				let ctx = graphics.m_oContext;
				ctx.beginPath();
				ctx.moveTo(left, top);
				ctx.lineTo(right, top);
				ctx.lineTo(right + 5, top);
				ctx.lineTo(right + EFFECT_BAR_HEIGHT * g_dKoef_mm_to_pix, top + (bottom - top) / 2);
				ctx.lineTo(right + 5, bottom);
				ctx.lineTo(right, bottom);
				ctx.lineTo(left, bottom);
				ctx.lineTo(left, top);

				graphics.df();
				graphics.ds();

				// renew clipRect to clip marks
				graphics.RemoveClipRect();
				graphics.AddClipRect(clipL, clipT, clipW - EFFECT_BAR_HEIGHT, clipH);
			} else {
				// In case we need to draw a bar

				repeats = this.getRepeatCount() / 1000;
				const barWidth = (bounds.r - bounds.l) * repeats;
				graphics.rect(bounds.l, bounds.t, barWidth, bounds.b - bounds.t);

				graphics.df();
				graphics.ds();
			}

			// draw marks
			if ((bounds.r - bounds.l) >= 2 * g_dKoef_pix_to_mm) {
				const gap = (bounds.b - bounds.t) / 5;
				for (let markIndex = 1; markIndex < repeats; markIndex++) {
					const xCord = bounds.l + markIndex * (bounds.r - bounds.l)
					graphics.drawVerLine(2, xCord, bounds.t + gap, bounds.b - gap, this.getPenWidth(graphics));
				}
			}
		}

		graphics.RestoreGrState();
	};
	CAnimItem.prototype.hitInEffectBar = function (x, y) {
		if (!this.hit(x, y)) { return null; }

		const bounds = this.getEffectBarBounds();
		const isOutOfBorders = x < this.getLeftBorder() || x > this.getRightBorder() || y < bounds.t || y > bounds.b
		if (isOutOfBorders) { return null; }

		const width = bounds.r - bounds.l;
		const repeats = this.getRepeatCount() / 1000;
		const delta = AscFormat.DIST_HIT_IN_LINE / 2

		let barRight = this.effect.isUntilEffect() ? this.getRightBorder() : bounds.l + width * repeats;
		barRight = Math.max(bounds.l + this.ms_to_mm(MIN_ALLOWED_DURATION), barRight);

		if (!this.effect.isInstantEffect()) {
			if (x >= bounds.l - delta && x <= bounds.l + delta) {
				return { type: 'left' };
			}

			if (this.effect.isUntilEffect()) {
				if (x >= barRight - EFFECT_BAR_HEIGHT - delta && x <= barRight + delta) {
					return { type: 'right' };
				}
			} else {
				if (x >= barRight - delta && x <= barRight + delta) {
					return { type: 'right' };
				}
			}

			const partitionIndex = (x - bounds.l) / width >> 0;
			// if effect isUntilEffect condition (partitionIndex < repeats) doesnt matter
			if (partitionIndex > 0 && (this.effect.isUntilEffect() || partitionIndex < repeats)) {
				const partitionPos = bounds.l + partitionIndex * width;
				if (x <= partitionPos + delta && x >= partitionPos - delta) {
					return { type: 'partition', index: partitionIndex };
				}
			}
		}

		if (x > bounds.l && x < barRight) {
			return { type: 'center', offset: x - bounds.l };
		}

		return null;
	};
	CAnimItem.prototype.hit = function (x, y) {
		const headerY = y + HEADER_HEIGHT - editor.WordControl.m_oAnimPaneApi.list.Scroll * g_dKoef_pix_to_mm
		if (editor.WordControl.m_oAnimPaneApi.header.Control.hit(x, headerY)) {
			return false;
		}

		if (this.parentControl && !this.parentControl.hit(x, y)) { return false; }

		const oInv = this.invertTransform;
		const tx = oInv.TransformPointX(x, y);
		const ty = oInv.TransformPointY(x, y);

		if (tx >= 0 && tx <= this.extX) {
			if (ty >= 0 && ty <= this.extY / 2) { return 'top' }
			if (ty > this.extY / 2 && ty <= this.extY) { return 'bottom' }
		}
		return false;
	};

	CAnimItem.prototype.setNewEffectParams = function (newDelay, newDuration, newRepeatCount) {
		const minAllowedDelta = 1 // in ms
		let dOwnNewDelay = newDelay - this.effect.getBaseTime();
		const delayDiff = Math.abs(dOwnNewDelay - this.effect.asc_getDelay());
		const durationDiff = Math.abs(newDuration - this.effect.asc_getDuration());
		const repeatCountDiff = Math.abs(newRepeatCount - this.effect.asc_getRepeatCount());

		const effectCopy = AscFormat.ExecuteNoHistory(function () {
			let oCopy = this.effect.createDuplicate();
			oCopy.merge(this.effect);
			return oCopy;
		}, this, []);

		if (dOwnNewDelay !== null && dOwnNewDelay !== undefined && delayDiff >= minAllowedDelta) {
			effectCopy.asc_putDelay(dOwnNewDelay);
		}
		if (newDuration !== null && newDuration !== undefined && durationDiff >= minAllowedDelta) {
			effectCopy.asc_putDuration(newDuration);
		}
		if (newRepeatCount !== null && newRepeatCount !== undefined && repeatCountDiff >= 1) {
			effectCopy.asc_putRepeatCount(newRepeatCount);
		}

		if (this.effect.isEqualProperties(effectCopy)) { return }
		Asc.editor.WordControl.m_oLogicDocument.SetAnimationProperties(effectCopy, false);
	};

	CAnimItem.prototype.onMouseDown = function (e, x, y) {
		if (this.onMouseDownCallback && this.onMouseDownCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseDown.call(this, e, x, y);
	};
	CAnimItem.prototype.onMouseMove = function (e, x, y) {
		if (this.onMouseMoveCallback && this.onMouseMoveCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseMove.call(this, e, x, y);
	};
	CAnimItem.prototype.onMouseUp = function (e, x, y) {
		if (this.onMouseUpCallback && this.onMouseUpCallback.call(this, e, x, y)) {
			return true;
		}
		return CControlContainer.prototype.onMouseUp.call(this, e, x, y);
	};
	CAnimItem.prototype.canHandleEvents = function () {
		return true;
	};
	CAnimItem.prototype.getFillColor = function() {
		const oSkin = AscCommon.GlobalSkin;
		if (this.effect.isSelected()) { return oSkin.AnimPaneItemFillSelected; }
		if (this.isHovered()) { return oSkin.AnimPaneItemFillHovered; }
		return null;
	};


	// HEADER
	const HEADER_HEIGHT = 40 * AscCommon.g_dKoef_pix_to_mm;
	const HEADER_LABEL_FONTSIZE = 10;
	const HEADER_LABEL_WIDTH = 101 * AscCommon.g_dKoef_pix_to_mm;

	const PLAY_BUTTON_HEIGHT = 24 * AscCommon.g_dKoef_pix_to_mm;
	const PLAY_BUTTON_LEFT_MARGIN = 30 * AscCommon.g_dKoef_pix_to_mm;
	const PLAY_BUTTON_LEFT_PADDING = 8 *  AscCommon.g_dKoef_pix_to_mm;
	const PLAY_BUTTON_RIGHT_PADDING = 12 *  AscCommon.g_dKoef_pix_to_mm;
	const PLAY_BUTTON_ICON_SIZE = 20 * AscCommon.g_dKoef_pix_to_mm;;
	const PLAY_BUTTON_LABEL_FONTSIZE = 9;
	const PLAY_BUTTON_MAX_LABEL_WIDTH = 100 * g_dKoef_pix_to_mm;
	const PLAY_BUTTON_LABEL_LEFT_MARGIN = 4 * AscCommon.g_dKoef_pix_to_mm;

	const MOVE_BUTTON_SIZE = 24 * AscCommon.g_dKoef_pix_to_mm;
	const MOVE_UP_BUTTON_LEFT_MARGIN = 14 * AscCommon.g_dKoef_pix_to_mm;
	const MOVE_DOWN_BUTTON_LEFT_MARGIN = 4 * AscCommon.g_dKoef_pix_to_mm;

	const CLOSE_BUTTON_SIZE = 24 * AscCommon.g_dKoef_pix_to_mm;


	// TIMELINE
	const TIMELINE_HEIGHT = 40 * AscCommon.g_dKoef_pix_to_mm;
	const TIMELINE_SCROLL_HEIGHT = 17 * AscCommon.g_dKoef_pix_to_mm;
	const TIMELINE_SCROLL_ABSOLUTE_LEFT = 143 * AscCommon.g_dKoef_pix_to_mm;
	const TIMELINE_SCROLL_LEFT_MARGIN = 10 * AscCommon.g_dKoef_pix_to_mm;
	const TIMELINE_SCROLL_RIGHT_MARGIN = 40 * AscCommon.g_dKoef_pix_to_mm;
	const TIMELINE_SCROLL_BUTTON_SIZE = 17 * AscCommon.g_dKoef_pix_to_mm;
	const TIMELINE_SCROLLER_WIDTH = 16 * AscCommon.g_dKoef_pix_to_mm;

	const ZOOM_BUTTON_SIZE = 20 * AscCommon.g_dKoef_pix_to_mm;
	const ZOOM_LABEL_FONTSIZE = 9;
	const ZOOM_LABEL_WIDTH = 40 * AscCommon.g_dKoef_pix_to_mm;

	const SCALE_BUTTON_WIDTH = 76 * AscCommon.g_dKoef_pix_to_mm;
	const SCALE_BUTTON_LEFT_MARGIN = 43 * AscCommon.g_dKoef_pix_to_mm;

	const TIMELINE_LABEL_WIDTH = 100;
	const TIMELINE_LABEL_FONTSIZE = 7.5;

	const SCROLL_TIMER_INTERVAL = 150;
	const SCROLL_TIMER_DELAY = 600;
	const SCROLL_STEP = 0.26

	const TIME_SCALES = [0.25, 1, 1, 2, 5, 10, 20, 60, 120, 300, 600, 600]; // in seconds

	const SMALL_TIME_INTERVAL = 15;
	const MIDDLE_1_TIME_INTERVAL = 20;
	const MIDDLE_2_TIME_INTERVAL = 25;
	const LONG_TIME_INTERVAL = 30;

	const TIME_INTERVALS = [
		SMALL_TIME_INTERVAL,
		LONG_TIME_INTERVAL, //1
		SMALL_TIME_INTERVAL, //1
		SMALL_TIME_INTERVAL, //2
		MIDDLE_1_TIME_INTERVAL, //5
		MIDDLE_1_TIME_INTERVAL,//10
		MIDDLE_1_TIME_INTERVAL,//20
		MIDDLE_2_TIME_INTERVAL,//60
		MIDDLE_2_TIME_INTERVAL,//120
		MIDDLE_2_TIME_INTERVAL,//300
		MIDDLE_2_TIME_INTERVAL,//600
		SMALL_TIME_INTERVAL//600
	]; // in mms


	// SEQUENCE LIST
	const SEQ_LABEL_MARGIN = 3;
	const SEQ_LABEL_HEIGHT = 15 * AscCommon.g_dKoef_pix_to_mm;

	const ANIM_ITEM_HEIGHT = 24 * AscCommon.g_dKoef_pix_to_mm;

	const INDEX_LABEL_FONTSIZE = 7.5;
	const INDEX_LABEL_WIDTH = 19 * AscCommon.g_dKoef_pix_to_mm;
	const EVENT_TYPE_ICON_SIZE = ANIM_ITEM_HEIGHT;
	const EFFECT_TYPE_ICON_SIZE = ANIM_ITEM_HEIGHT;
	const EFFECT_LABEL_FONTSIZE = 7.5;
	const EFFECT_BAR_HEIGHT = ANIM_ITEM_HEIGHT * 5 / 8;
	const MENU_BUTTON_SIZE = 15 * AscCommon.g_dKoef_pix_to_mm;

	const MIN_ALLOWED_DURATION = 10; // milliseconds
	const MIN_ALLOWED_REPEAT_COUNT = 10; // equals 0.01 of full effect duration


	// COMMON
	const COMMON_LEFT_MARGIN = 14 * AscCommon.g_dKoef_pix_to_mm;
	const COMMON_RIGHT_MARGIN = 20 * AscCommon.g_dKoef_pix_to_mm;
	// const ALIGNMENT_LINE = MATH.max(CLOSE_BUTTON_SIZE, MENU_BUTTON_SIZE, TIMELINE_SCROLL_BUTTON_SIZE) / 2;


	// ICONS
	const playIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMTAiIHZpZXdCb3g9IjAgMCAxMCAxMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0yIDEwTDEwIDVMMiAtNC43NjgzN2UtMDdWMTBaIiBmaWxsPSJibGFjayIvPgo8L3N2Zz4K';
	const playIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMTAiIHZpZXdCb3g9IjAgMCAxMCAxMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0yIDEwTDEwIDVMMiAtNC43NjgzN2UtMDdWMTBaIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K';
	const stopIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMTAiIHZpZXdCb3g9IjAgMCAxMCAxMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3Qgd2lkdGg9IjEwIiBoZWlnaHQ9IjEwIiBmaWxsPSJibGFjayIvPgo8L3N2Zz4K';
	const stopIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMTAiIHZpZXdCb3g9IjAgMCAxMCAxMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHJlY3Qgd2lkdGg9IjEwIiBoZWlnaHQ9IjEwIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K';
	
	const clickEffectIcon = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxMiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTIgMC41SDEwQzEwLjgyODQgMC41IDExLjUgMS4xNzE1NyAxMS41IDJWMTJDMTEuNSAxMy45MzMgOS45MzMgMTUuNSA4IDE1LjVINEMyLjA2NyAxNS41IDAuNSAxMy45MzMgMC41IDEyVjJDMC41IDEuMTcxNTcgMS4xNzE1NyAwLjUgMiAwLjVaIiBmaWxsPSJ3aGl0ZSIgc3Ryb2tlPSIjNDQ0NDQ0Ii8+CjxyZWN0IHg9IjUiIHk9IjIiIHdpZHRoPSIyIiBoZWlnaHQ9IjQiIGZpbGw9IiM0NDQ0NDQiLz4KPC9zdmc+Cg==';
	const afterEffectIcon = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTggMTUuNUMxMi4xNDIxIDE1LjUgMTUuNSAxMi4xNDIxIDE1LjUgOEMxNS41IDMuODU3ODYgMTIuMTQyMSAwLjUgOCAwLjVDMy44NTc4NiAwLjUgMC41IDMuODU3ODYgMC41IDhDMC41IDEyLjE0MjEgMy44NTc4NiAxNS41IDggMTUuNVoiIGZpbGw9IndoaXRlIiBzdHJva2U9IiM0NDQ0NDQiLz4KPHBhdGggZD0iTTExIDguNUg3IiBzdHJva2U9IiM0NDQ0NDQiLz4KPHBhdGggZD0iTTcuNSA0VjkiIHN0cm9rZT0iIzQ0NDQ0NCIvPgo8L3N2Zz4K';

	const entrEffectIcon = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMCAxLjAwMDAyQzkuNzExNjQgMC45OTgxNDIgOS40MjIyNSAxLjE2Mjk3IDkuMzE0NTQgMS40OTQ0OUw3LjM3MDc5IDcuNTIzNDZDNy4yOTE4MyA3Ljc2NjQ3IDcuMDkxMzIgNy45NDQ4IDYuODQ5NjkgOEgxLjcxN0MxLjAyMzc1IDggMC43MzU1MTYgOC44ODcxIDEuMjk2MzYgOS4yOTQ1OEw1LjgyNzIyIDEyLjQ4NzlDNS43ODMzMiAxMi42NjA5IDUuNzMxNzEgMTIuODQwMyA1LjY3MTI5IDEzLjAyNjJMNC4wMzcyOCAxOC4wNTUyQzMuODIzMDUgMTguNzE0NSA0LjU3NzY3IDE5LjI2MjggNS4xMzg1MiAxOC44NTUzTDEwIDE1LjQyODlMMTQuODYxNSAxOC44NTUzQzE1LjQyMjMgMTkuMjYyOCAxNi4xNzY5IDE4LjcxNDUgMTUuOTYyNyAxOC4wNTUyTDE0LjMyODcgMTMuMDI2MkMxNC4yNjgzIDEyLjg0MDMgMTQuMjE2NyAxMi42NjA5IDE0LjE3MjggMTIuNDg3OUwxOC43MDM2IDkuMjk0NThDMTkuMjY0NSA4Ljg4NzEgMTguOTc2MiA4IDE4LjI4MyA4SDEzLjE1MDNDMTIuOTA4NyA3Ljk0NDggMTIuNzA4MiA3Ljc2NjQ3IDEyLjYyOTIgNy41MjM0NkwxMC42ODU1IDEuNDk0NDlDMTAuNTc3NyAxLjE2Mjk3IDEwLjI4ODQgMC45OTgxNDIgMTAgMS4wMDAwMlpNMTAgMi42MDAzMUM5LjMwMDY2IDQuNzM5NDYgNy44MTc1MSA5LjAwMDAxIDcuODE3NTEgOS4wMDAwMUgyLjYwMDFMNy4wMDAxIDEyLjFDNi42MTQ0OSAxMy40NDk2IDYuMTcxNzMgMTQuNzQ0OCA1LjcyMjY5IDE2LjA1ODRDNS41NDg0NSAxNi41NjgxIDUuMzczMjYgMTcuMDgwNSA1LjIwMDEgMTcuNkwxMCAxNC4yMzIyTDE0Ljc5OTkgMTcuNkMxNC42MjY3IDE3LjA4MDYgMTQuNDUxNiAxNi41NjgxIDE0LjI3NzMgMTYuMDU4NUMxMy44MjgzIDE0Ljc0NDkgMTMuMzg1NSAxMy40NDk2IDEyLjk5OTkgMTIuMUwxNy4zOTk5IDkuMDAwMDFIMTIuMTgyNUMxMi4xODI1IDkuMDAwMDEgMTAuNjk5MyA0LjczOTQ2IDEwIDIuNjAwMzFaIiBmaWxsPSIjMEU4QTI2Ii8+CjxwYXRoIG9wYWNpdHk9IjAuNSIgZD0iTTcuODE3NTEgOS4wMDAyOUM3LjgxNzUxIDkuMDAwMjkgOS4zMDA2NiA0LjczOTc0IDEwIDIuNjAwNTlDMTAuNjk5MyA0LjczOTc0IDEyLjE4MjUgOS4wMDAyOSAxMi4xODI1IDkuMDAwMjlIMTcuMzk5OUwxMi45OTk5IDEyLjEwMDNDMTMuMzg1NSAxMy40NDk5IDEzLjgyODMgMTQuNzQ1MSAxNC4yNzczIDE2LjA1ODdDMTQuNDUxNiAxNi41NjgzIDE0LjYyNjcgMTcuMDgwOCAxNC43OTk5IDE3LjYwMDNMMTAgMTQuMjMyNUw1LjIwMDEgMTcuNjAwM0M1LjM3MzI2IDE3LjA4MDggNS41NDg0NCAxNi41NjgzIDUuNzIyNjggMTYuMDU4N0M2LjE3MTczIDE0Ljc0NTEgNi42MTQ0OSAxMy40NDk5IDcuMDAwMSAxMi4xMDAzTDIuNjAwMSA5LjAwMDI5SDcuODE3NTFaIiBmaWxsPSIjMEU4QTI2Ii8+CjxyZWN0IHg9IjMiIHk9IjIiIHdpZHRoPSI0IiBoZWlnaHQ9IjEiIGZpbGw9IiMwRThBMjYiLz4KPHJlY3QgeD0iMSIgeT0iNSIgd2lkdGg9IjQiIGhlaWdodD0iMSIgZmlsbD0iIzBFOEEyNiIvPgo8cmVjdCB4PSIxIiB5PSIxMiIgd2lkdGg9IjMiIGhlaWdodD0iMSIgZmlsbD0iIzBFOEEyNiIvPgo8cmVjdCB4PSIxIiB5PSIxNSIgd2lkdGg9IjIiIGhlaWdodD0iMSIgZmlsbD0iIzBFOEEyNiIvPgo8L3N2Zz4K';
	const emphEffectIcon = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMCAxLjAwMDAyQzkuNzExNjQgMC45OTgxNDIgOS40MjIyNSAxLjE2Mjk3IDkuMzE0NTQgMS40OTQ0OUw3LjM3MDc5IDcuNTIzNDZDNy4yOTE4MyA3Ljc2NjQ3IDcuMDkxMzIgNy45NDQ4IDYuODQ5NjkgOEgxLjcxN0MxLjAyMzc1IDggMC43MzU1MTYgOC44ODcxIDEuMjk2MzYgOS4yOTQ1OEw1LjgyNzIyIDEyLjQ4NzlDNS43ODMzMiAxMi42NjA5IDUuNzMxNzEgMTIuODQwMyA1LjY3MTI5IDEzLjAyNjJMNC4wMzcyOCAxOC4wNTUyQzMuODIzMDUgMTguNzE0NSA0LjU3NzY3IDE5LjI2MjggNS4xMzg1MiAxOC44NTUzTDEwIDE1LjQyODlMMTQuODYxNSAxOC44NTUzQzE1LjQyMjMgMTkuMjYyOCAxNi4xNzY5IDE4LjcxNDUgMTUuOTYyNyAxOC4wNTUyTDE0LjMyODcgMTMuMDI2MkMxNC4yNjgzIDEyLjg0MDMgMTQuMjE2NyAxMi42NjA5IDE0LjE3MjggMTIuNDg3OUwxOC43MDM2IDkuMjk0NThDMTkuMjY0NSA4Ljg4NzEgMTguOTc2MiA4IDE4LjI4MyA4SDEzLjE1MDNDMTIuOTA4NyA3Ljk0NDggMTIuNzA4MiA3Ljc2NjQ3IDEyLjYyOTIgNy41MjM0NkwxMC42ODU1IDEuNDk0NDlDMTAuNTc3NyAxLjE2Mjk3IDEwLjI4ODQgMC45OTgxNDIgMTAgMS4wMDAwMlpNMTAgMi42MDAzMUM5LjMwMDY2IDQuNzM5NDYgNy44MTc1MSA5LjAwMDAxIDcuODE3NTEgOS4wMDAwMUgyLjYwMDFMNy4wMDAxIDEyLjFDNi42MTQ0OSAxMy40NDk2IDYuMTcxNzMgMTQuNzQ0OCA1LjcyMjY5IDE2LjA1ODRDNS41NDg0NSAxNi41NjgxIDUuMzczMjYgMTcuMDgwNSA1LjIwMDEgMTcuNkwxMCAxNC4yMzIyTDE0Ljc5OTkgMTcuNkMxNC42MjY3IDE3LjA4MDYgMTQuNDUxNiAxNi41NjgxIDE0LjI3NzMgMTYuMDU4NUMxMy44MjgzIDE0Ljc0NDkgMTMuMzg1NSAxMy40NDk2IDEyLjk5OTkgMTIuMUwxNy4zOTk5IDkuMDAwMDFIMTIuMTgyNUMxMi4xODI1IDkuMDAwMDEgMTAuNjk5MyA0LjczOTQ2IDEwIDIuNjAwMzFaIiBmaWxsPSIjRkY4RTAwIi8+CjxwYXRoIG9wYWNpdHk9IjAuNSIgZD0iTTcuODE3NTEgOS4wMDAyOUM3LjgxNzUxIDkuMDAwMjkgOS4zMDA2NiA0LjczOTc0IDEwIDIuNjAwNTlDMTAuNjk5MyA0LjczOTc0IDEyLjE4MjUgOS4wMDAyOSAxMi4xODI1IDkuMDAwMjlIMTcuMzk5OUwxMi45OTk5IDEyLjEwMDNDMTMuMzg1NSAxMy40NDk5IDEzLjgyODMgMTQuNzQ1MSAxNC4yNzczIDE2LjA1ODdDMTQuNDUxNiAxNi41NjgzIDE0LjYyNjcgMTcuMDgwOCAxNC43OTk5IDE3LjYwMDNMMTAgMTQuMjMyNUw1LjIwMDEgMTcuNjAwM0M1LjM3MzI2IDE3LjA4MDggNS41NDg0NCAxNi41NjgzIDUuNzIyNjggMTYuMDU4N0M2LjE3MTczIDE0Ljc0NTEgNi42MTQ0OSAxMy40NDk5IDcuMDAwMSAxMi4xMDAzTDIuNjAwMSA5LjAwMDI5SDcuODE3NTFaIiBmaWxsPSIjRkY4RTAwIi8+CjxwYXRoIGQ9Ik01Ljg5NDU4IDYuMTkzMzZMMi4zOTQ1OCAxLjY5MzM2TDEuNjA1MjIgMi4zMDczTDUuMTA1MjIgNi44MDczTDUuODk0NTggNi4xOTMzNloiIGZpbGw9IiNGRjhFMDAiLz4KPHBhdGggZD0iTTEuODUzNTkgMTUuODUzOUw0LjM1MzU5IDEzLjM1MzlMMy42NDY0OCAxMi42NDY4TDEuMTQ2NDggMTUuMTQ2OEwxLjg1MzU5IDE1Ljg1MzlaIiBmaWxsPSIjRkY4RTAwIi8+CjxwYXRoIGQ9Ik0xNC4xNDY0IDYuMTkzMzZMMTcuNjQ2NCAxLjY5MzM2TDE4LjQzNTggMi4zMDczTDE0LjkzNTggNi44MDczTDE0LjE0NjQgNi4xOTMzNloiIGZpbGw9IiNGRjhFMDAiLz4KPHBhdGggZD0iTTE4LjE4NzQgMTUuODUzOUwxNS42ODc0IDEzLjM1MzlMMTYuMzk0NSAxMi42NDY4TDE4Ljg5NDUgMTUuMTQ2OEwxOC4xODc0IDE1Ljg1MzlaIiBmaWxsPSIjRkY4RTAwIi8+Cjwvc3ZnPgo=';
	const exitEffectIcon = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMCAxLjAwMDAyQzkuNzExNjQgMC45OTgxNDIgOS40MjIyNSAxLjE2Mjk3IDkuMzE0NTQgMS40OTQ0OUw3LjM3MDc5IDcuNTIzNDZDNy4yOTE4MyA3Ljc2NjQ3IDcuMDkxMzIgNy45NDQ4IDYuODQ5NjkgOEgxLjcxN0MxLjAyMzc1IDggMC43MzU1MTYgOC44ODcxIDEuMjk2MzYgOS4yOTQ1OEw1LjgyNzIyIDEyLjQ4NzlDNS43ODMzMiAxMi42NjA5IDUuNzMxNzEgMTIuODQwMyA1LjY3MTI5IDEzLjAyNjJMNC4wMzcyOCAxOC4wNTUyQzMuODIzMDUgMTguNzE0NSA0LjU3NzY3IDE5LjI2MjggNS4xMzg1MiAxOC44NTUzTDEwIDE1LjQyODlMMTQuODYxNSAxOC44NTUzQzE1LjQyMjMgMTkuMjYyOCAxNi4xNzY5IDE4LjcxNDUgMTUuOTYyNyAxOC4wNTUyTDE0LjMyODcgMTMuMDI2MkMxNC4yNjgzIDEyLjg0MDMgMTQuMjE2NyAxMi42NjA5IDE0LjE3MjggMTIuNDg3OUwxOC43MDM2IDkuMjk0NThDMTkuMjY0NSA4Ljg4NzEgMTguOTc2MiA4IDE4LjI4MyA4SDEzLjE1MDNDMTIuOTA4NyA3Ljk0NDggMTIuNzA4MiA3Ljc2NjQ3IDEyLjYyOTIgNy41MjM0NkwxMC42ODU1IDEuNDk0NDlDMTAuNTc3NyAxLjE2Mjk3IDEwLjI4ODQgMC45OTgxNDIgMTAgMS4wMDAwMlpNMTAgMi42MDAzMUM5LjMwMDY2IDQuNzM5NDYgNy44MTc1MSA5LjAwMDAxIDcuODE3NTEgOS4wMDAwMUgyLjYwMDFMNy4wMDAxIDEyLjFDNi42MTQ0OSAxMy40NDk2IDYuMTcxNzMgMTQuNzQ0OCA1LjcyMjY5IDE2LjA1ODRDNS41NDg0NSAxNi41NjgxIDUuMzczMjYgMTcuMDgwNSA1LjIwMDEgMTcuNkwxMCAxNC4yMzIyTDE0Ljc5OTkgMTcuNkMxNC42MjY3IDE3LjA4MDYgMTQuNDUxNiAxNi41NjgxIDE0LjI3NzMgMTYuMDU4NUMxMy44MjgzIDE0Ljc0NDkgMTMuMzg1NSAxMy40NDk2IDEyLjk5OTkgMTIuMUwxNy4zOTk5IDkuMDAwMDFIMTIuMTgyNUMxMi4xODI1IDkuMDAwMDEgMTAuNjk5MyA0LjczOTQ2IDEwIDIuNjAwMzFaIiBmaWxsPSIjRjIzRDNEIi8+CjxwYXRoIG9wYWNpdHk9IjAuNSIgZD0iTTcuODE3NTEgOS4wMDAyOUM3LjgxNzUxIDkuMDAwMjkgOS4zMDA2NiA0LjczOTc0IDEwIDIuNjAwNTlDMTAuNjk5MyA0LjczOTc0IDEyLjE4MjUgOS4wMDAyOSAxMi4xODI1IDkuMDAwMjlIMTcuMzk5OUwxMi45OTk5IDEyLjEwMDNDMTMuMzg1NSAxMy40NDk5IDEzLjgyODMgMTQuNzQ1MSAxNC4yNzczIDE2LjA1ODdDMTQuNDUxNiAxNi41NjgzIDE0LjYyNjcgMTcuMDgwOCAxNC43OTk5IDE3LjYwMDNMMTAgMTQuMjMyNUw1LjIwMDEgMTcuNjAwM0M1LjM3MzI2IDE3LjA4MDggNS41NDg0NCAxNi41NjgzIDUuNzIyNjggMTYuMDU4N0M2LjE3MTczIDE0Ljc0NTEgNi42MTQ0OSAxMy40NDk5IDcuMDAwMSAxMi4xMDAzTDIuNjAwMSA5LjAwMDI5SDcuODE3NTFaIiBmaWxsPSIjRjIzRDNEIi8+CjxwYXRoIGQ9Ik0xMyAySDE3VjNIMTNWMloiIGZpbGw9IiNGMjNEM0QiLz4KPHBhdGggZD0iTTE1IDZWNUgxOVY2SDE1WiIgZmlsbD0iI0YyM0QzRCIvPgo8cGF0aCBkPSJNMTkgMTJIMTZWMTNIMTlWMTJaIiBmaWxsPSIjRjIzRDNEIi8+CjxwYXRoIGQ9Ik0xNyAxNUgxOVYxNkgxN1YxNVoiIGZpbGw9IiNGMjNEM0QiLz4KPC9zdmc+Cg==';
	const pathEffectIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTkuOTk2NzUgMS41MDAwMUw5Ljk5Njc1IDEuNTAwMDVMMTAuMDAzMiAxLjUwMDAxQzEwLjA1MjUgMS40OTk2OSAxMC4wOTcgMS41MTM4IDEwLjEzMDQgMS41MzY0QzEwLjE2MDUgMS41NTY3NSAxMC4xOTA1IDEuNTg5NjEgMTAuMjA5OCAxLjY0ODU4QzEwLjIwOTggMS42NDg3MiAxMC4yMDk5IDEuNjQ4ODYgMTAuMjA5OSAxLjY0ODk5TDEyLjE1MzMgNy42NzY4OUwxMi4xNTM3IDcuNjc3OTdDMTIuMjg3OSA4LjA5MDk3IDEyLjYyODYgOC4zOTM2OSAxMy4wMzg5IDguNDg3NDRMMTMuMDkzOSA4LjVIMTMuMTUwM0gxOC4yODNDMTguNDkxIDguNSAxOC41NzggOC43NjUwNCAxOC40MTE5IDguODg4NTFMMTMuODg0NyAxMi4wNzkzTDEzLjYwMzUgMTIuMjc3NEwxMy42ODgxIDEyLjYxMDlDMTMuNzM0OCAxMi43OTQ2IDEzLjc4OTQgMTIuOTg0NSAxMy44NTMyIDEzLjE4MDhMMTUuNDg3MiAxOC4yMDk3QzE1LjU1MTcgMTguNDA4NCAxNS4zMjQ0IDE4LjU3MzYgMTUuMTU1NCAxOC40NTA4TDE1LjE1NTQgMTguNDUwOEwxNS4xNDk1IDE4LjQ0NjZMMTAuMjg4IDE1LjAyMDJMMTAgMTQuODE3Mkw5LjcxMTk1IDE1LjAyMDJMNC44NTA0NyAxOC40NDY2TDQuODUwNDQgMTguNDQ2Nkw0Ljg0NDYyIDE4LjQ1MDhDNC42NzU2NCAxOC41NzM2IDQuNDQ4MjYgMTguNDA4NCA0LjUxMjgxIDE4LjIwOTdMNi4xNDY4MiAxMy4xODA4QzYuMjEwNTggMTIuOTg0NSA2LjI2NTI0IDEyLjc5NDYgNi4zMTE4NiAxMi42MTA5TDYuMzk2NDcgMTIuMjc3NEw2LjExNTI2IDEyLjA3OTNMMS41ODgxNSA4Ljg4ODUzQzEuNDIyIDguNzY1MDYgMS41MDg5OSA4LjUgMS43MTcgOC41SDYuODQ5NjlINi45MDYwOEw2Ljk2MTA1IDguNDg3NDRDNy4zNzE0MyA4LjM5MzY5IDcuNzEyMTMgOC4wOTA5NyA3Ljg0NjMyIDcuNjc3OTdMNy44NDY2NyA3LjY3Njg5TDkuNzkwMDcgMS42NDg5OUM5Ljc5MDExIDEuNjQ4ODYgOS43OTAxNSAxLjY0ODcyIDkuNzkwMiAxLjY0ODU5QzkuODA5NDUgMS41ODk2MSA5LjgzOTU0IDEuNTU2NzYgOS44Njk2MyAxLjUzNjRDOS45MDMwNSAxLjUxMzggOS45NDc1NCAxLjQ5OTY5IDkuOTk2NzUgMS41MDAwMVoiIHN0cm9rZT0id2hpdGUiIHN0cm9rZS1vcGFjaXR5PSIwLjgiLz4KPHBhdGggZD0iTTEyIDNDMTIgNC4xMDQ1NyAxMS4xMDQ2IDUgMTAgNUM4Ljg5NTQzIDUgOCA0LjEwNDU3IDggM0M4IDEuODk1NDMgOC44OTU0MyAxIDEwIDFDMTEuMTA0NiAxIDEyIDEuODk1NDMgMTIgM1oiIGZpbGw9IiMzQUJENTkiLz4KPHBhdGggZD0iTTkgN0M5IDguMTA0NTcgOC4xMDQ1NyA5IDcgOUM1Ljg5NTQzIDkgNSA4LjEwNDU3IDUgN0M1IDUuODk1NDMgNS44OTU0MyA1IDcgNUM4LjEwNDU3IDUgOSA1Ljg5NTQzIDkgN1oiIGZpbGw9IiNGNTJDMkMiLz4KPC9zdmc+Cg==';
	const pathEffectIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAiIGhlaWdodD0iMjAiIHZpZXdCb3g9IjAgMCAyMCAyMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTkuOTk2NzUgMS41MDAwMUw5Ljk5Njc1IDEuNTAwMDVMMTAuMDAzMiAxLjUwMDAxQzEwLjA1MjUgMS40OTk2OSAxMC4wOTcgMS41MTM4IDEwLjEzMDQgMS41MzY0QzEwLjE2MDUgMS41NTY3NSAxMC4xOTA1IDEuNTg5NjEgMTAuMjA5OCAxLjY0ODU4QzEwLjIwOTggMS42NDg3MiAxMC4yMDk5IDEuNjQ4ODYgMTAuMjA5OSAxLjY0ODk5TDEyLjE1MzMgNy42NzY4OUwxMi4xNTM3IDcuNjc3OTdDMTIuMjg3OSA4LjA5MDk3IDEyLjYyODYgOC4zOTM2OSAxMy4wMzg5IDguNDg3NDRMMTMuMDkzOSA4LjVIMTMuMTUwM0gxOC4yODNDMTguNDkxIDguNSAxOC41NzggOC43NjUwNCAxOC40MTE5IDguODg4NTFMMTMuODg0NyAxMi4wNzkzTDEzLjYwMzUgMTIuMjc3NEwxMy42ODgxIDEyLjYxMDlDMTMuNzM0OCAxMi43OTQ2IDEzLjc4OTQgMTIuOTg0NSAxMy44NTMyIDEzLjE4MDhMMTUuNDg3MiAxOC4yMDk3QzE1LjU1MTcgMTguNDA4NCAxNS4zMjQ0IDE4LjU3MzYgMTUuMTU1NCAxOC40NTA4TDE1LjE1NTQgMTguNDUwOEwxNS4xNDk1IDE4LjQ0NjZMMTAuMjg4IDE1LjAyMDJMMTAgMTQuODE3Mkw5LjcxMTk1IDE1LjAyMDJMNC44NTA0NyAxOC40NDY2TDQuODUwNDQgMTguNDQ2Nkw0Ljg0NDYyIDE4LjQ1MDhDNC42NzU2NCAxOC41NzM2IDQuNDQ4MjYgMTguNDA4NCA0LjUxMjgxIDE4LjIwOTdMNi4xNDY4MiAxMy4xODA4QzYuMjEwNTggMTIuOTg0NSA2LjI2NTI0IDEyLjc5NDYgNi4zMTE4NiAxMi42MTA5TDYuMzk2NDcgMTIuMjc3NEw2LjExNTI2IDEyLjA3OTNMMS41ODgxNSA4Ljg4ODUzQzEuNDIyIDguNzY1MDYgMS41MDg5OSA4LjUgMS43MTcgOC41SDYuODQ5NjlINi45MDYwOEw2Ljk2MTA1IDguNDg3NDRDNy4zNzE0MyA4LjM5MzY5IDcuNzEyMTMgOC4wOTA5NyA3Ljg0NjMyIDcuNjc3OTdMNy44NDY2NyA3LjY3Njg5TDkuNzkwMDcgMS42NDg5OUM5Ljc5MDExIDEuNjQ4ODYgOS43OTAxNSAxLjY0ODcyIDkuNzkwMiAxLjY0ODU5QzkuODA5NDUgMS41ODk2MSA5LjgzOTU0IDEuNTU2NzYgOS44Njk2MyAxLjUzNjRDOS45MDMwNSAxLjUxMzggOS45NDc1NCAxLjQ5OTY5IDkuOTk2NzUgMS41MDAwMVoiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS1vcGFjaXR5PSIwLjgiLz4KPHBhdGggZD0iTTEyIDNDMTIgNC4xMDQ1NyAxMS4xMDQ2IDUgMTAgNUM4Ljg5NTQzIDUgOCA0LjEwNDU3IDggM0M4IDEuODk1NDMgOC44OTU0MyAxIDEwIDFDMTEuMTA0NiAxIDEyIDEuODk1NDMgMTIgM1oiIGZpbGw9IiMwRThBMjYiLz4KPHBhdGggZD0iTTkgN0M5IDguMTA0NTcgOC4xMDQ1NyA5IDcgOUM1Ljg5NTQzIDkgNSA4LjEwNDU3IDUgN0M1IDUuODk1NDMgNS44OTU0MyA1IDcgNUM4LjEwNDU3IDUgOSA1Ljg5NTQzIDkgN1oiIGZpbGw9IiNGMjNEM0QiLz4KPC9zdmc+Cg==';

	const arrowUpIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iNyIgdmlld0JveD0iMCAwIDEyIDciIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNNS45OTk5NiAwLjI5Mjg1NUw1LjY0NjQxIDAuNjQ2NDA4TDAuMTQ2NDA4IDYuMTQ2NDFMMC44NTM1MTYgNi44NTM1MUw1Ljk5OTk2IDEuNzA3MDdMMTEuMTQ2NCA2Ljg1MzUyTDExLjg1MzUgNi4xNDY0MUw2LjM1MzUyIDAuNjQ2NDA5TDUuOTk5OTYgMC4yOTI4NTVaIiBmaWxsPSJibGFjayIvPgo8L3N2Zz4K';
	const arrowUpIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iNyIgdmlld0JveD0iMCAwIDEyIDciIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNNS45OTk5NiAwLjI5Mjg1NUw1LjY0NjQxIDAuNjQ2NDA4TDAuMTQ2NDA4IDYuMTQ2NDFMMC44NTM1MTYgNi44NTM1MUw1Ljk5OTk2IDEuNzA3MDdMMTEuMTQ2NCA2Ljg1MzUyTDExLjg1MzUgNi4xNDY0MUw2LjM1MzUyIDAuNjQ2NDA5TDUuOTk5OTYgMC4yOTI4NTVaIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K';

	const arrowDownIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iNyIgdmlld0JveD0iMCAwIDEyIDciIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNNi4wMDAwNCA2LjcwNzE0TDYuMzUzNTkgNi4zNTM1OUwxMS44NTM2IDAuODUzNTkxTDExLjE0NjUgMC4xNDY0ODRMNi4wMDAwNCA1LjI5MjkzTDAuODUzNTkxIDAuMTQ2NDg0TDAuMTQ2NDg0IDAuODUzNTkxTDUuNjQ2NDggNi4zNTM1OUw2LjAwMDA0IDYuNzA3MTRaIiBmaWxsPSJibGFjayIvPgo8L3N2Zz4K';
	const arrowDownIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iNyIgdmlld0JveD0iMCAwIDEyIDciIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxwYXRoIGZpbGwtcnVsZT0iZXZlbm9kZCIgY2xpcC1ydWxlPSJldmVub2RkIiBkPSJNNi4wMDAwNCA2LjcwNzE0TDYuMzUzNTkgNi4zNTM1OUwxMS44NTM2IDAuODUzNTkxTDExLjE0NjUgMC4xNDY0ODRMNi4wMDAwNCA1LjI5MjkzTDAuODUzNTkxIDAuMTQ2NDg0TDAuMTQ2NDg0IDAuODUzNTkxTDUuNjQ2NDggNi4zNTM1OUw2LjAwMDA0IDYuNzA3MTRaIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K';

	const closeIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iMTIiIHZpZXdCb3g9IjAgMCAxMiAxMiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgMUwxMSAxMU0xMSAxTDEgMTEiIHN0cm9rZT0iYmxhY2siIHN0cm9rZS13aWR0aD0iMS41Ii8+Cjwvc3ZnPgo=';
	const closeIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iMTIiIHZpZXdCb3g9IjAgMCAxMiAxMiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEgMUwxMSAxMU0xMSAxTDEgMTEiIHN0cm9rZT0id2hpdGUiIHN0cm9rZS13aWR0aD0iMS41Ii8+Cjwvc3ZnPgo=';

	const menuButtonIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMiIgdmlld0JveD0iMCAwIDEwIDIiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxwYXRoIGQ9Ik04IDFDOCAwLjQ0NzcxNSA4LjQ0NzcyIDguMjg1MmUtMDggOSAxLjMxMTM0ZS0wN0M5LjU1MjI4IDEuNzk0MTdlLTA3IDEwIDAuNDQ3NzE2IDEwIDFDMTAgMS41NTIyOCA5LjU1MjI4IDIgOSAyQzguNDQ3NzIgMiA4IDEuNTUyMjggOCAxWiIgZmlsbD0iYmxhY2siLz4KPHBhdGggZD0iTTQgMUM0IDAuNDQ3NzE1IDQuNDQ3NzIgLTIuNjY4MzllLTA3IDUgLTIuMTg1NTdlLTA3QzUuNTUyMjkgLTEuNzAyNzVlLTA3IDYgMC40NDc3MTUgNiAxQzYgMS41NTIyOCA1LjU1MjI5IDIgNSAyQzQuNDQ3NzIgMiA0IDEuNTUyMjggNCAxWiIgZmlsbD0iYmxhY2siLz4KPHBhdGggZD0iTTEgMy44NTQyNmUtMDdDMC40NDc3MTYgMy4zNzE0NGUtMDcgLTEuOTU3MDNlLTA4IDAuNDQ3NzE2IC00LjM3MTE0ZS0wOCAxQy02Ljc4NTI2ZS0wOCAxLjU1MjI4IDAuNDQ3NzE2IDIgMSAyQzEuNTUyMjkgMiAyIDEuNTUyMjkgMiAxQzIgMC40NDc3MTYgMS41NTIyOSA0LjMzNzA5ZS0wNyAxIDMuODU0MjZlLTA3WiIgZmlsbD0iYmxhY2siLz4KPC9zdmc+Cg==';
	const menuButtonIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMiIgdmlld0JveD0iMCAwIDEwIDIiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxwYXRoIGQ9Ik04IDFDOCAwLjQ0NzcxNSA4LjQ0NzcyIDguMjg1MmUtMDggOSAxLjMxMTM0ZS0wN0M5LjU1MjI4IDEuNzk0MTdlLTA3IDEwIDAuNDQ3NzE2IDEwIDFDMTAgMS41NTIyOCA5LjU1MjI4IDIgOSAyQzguNDQ3NzIgMiA4IDEuNTUyMjggOCAxWiIgZmlsbD0id2hpdGUiLz4KPHBhdGggZD0iTTQgMUM0IDAuNDQ3NzE1IDQuNDQ3NzIgLTIuNjY4MzllLTA3IDUgLTIuMTg1NTdlLTA3QzUuNTUyMjkgLTEuNzAyNzVlLTA3IDYgMC40NDc3MTUgNiAxQzYgMS41NTIyOCA1LjU1MjI5IDIgNSAyQzQuNDQ3NzIgMiA0IDEuNTUyMjggNCAxWiIgZmlsbD0id2hpdGUiLz4KPHBhdGggZD0iTTEgMy44NTQyNmUtMDdDMC40NDc3MTYgMy4zNzE0NGUtMDcgLTEuOTU3MDNlLTA4IDAuNDQ3NzE2IC00LjM3MTE0ZS0wOCAxQy02Ljc4NTI2ZS0wOCAxLjU1MjI4IDAuNDQ3NzE2IDIgMSAyQzEuNTUyMjkgMiAyIDEuNTUyMjkgMiAxQzIgMC40NDc3MTYgMS41NTIyOSA0LjMzNzA5ZS0wNyAxIDMuODU0MjZlLTA3WiIgZmlsbD0id2hpdGUiLz4KPC9zdmc+Cg==';

	const arrowLeftIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNSIgaGVpZ2h0PSI5IiB2aWV3Qm94PSIwIDAgNSA5IiBmaWxsPSJub25lIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPgo8cGF0aCBkPSJNNSA5LjUzNjc0ZS0wN0w1IDlMMC41IDQuNUw1IDkuNTM2NzRlLTA3WiIgZmlsbD0iYmxhY2siLz4KPC9zdmc+Cg==';
	const arrowLeftIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNSIgaGVpZ2h0PSI5IiB2aWV3Qm94PSIwIDAgNSA5IiBmaWxsPSJub25lIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPgo8cGF0aCBkPSJNNSA5LjUzNjc0ZS0wN0w1IDlMMC41IDQuNUw1IDkuNTM2NzRlLTA3WiIgZmlsbD0id2hpdGUiLz4KPC9zdmc+Cg==';
	const arrowRightIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNSIgaGVpZ2h0PSI5IiB2aWV3Qm94PSIwIDAgNSA5IiBmaWxsPSJub25lIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPgo8cGF0aCBkPSJNMCA5TDAgMEw0LjUgNC41TDAgOVoiIGZpbGw9ImJsYWNrIi8+Cjwvc3ZnPgo=';
	const arrowRightIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNSIgaGVpZ2h0PSI5IiB2aWV3Qm94PSIwIDAgNSA5IiBmaWxsPSJub25lIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPgo8cGF0aCBkPSJNMCA5TDAgMEw0LjUgNC41TDAgOVoiIGZpbGw9IndoaXRlIi8+Cjwvc3ZnPgo=';

	const zoomInIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTEiIGhlaWdodD0iMTEiIHZpZXdCb3g9IjAgMCAxMSAxMSIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMSA2TDExIDVMNiA1TDYgLTIuMTg1NTdlLTA3TDUgLTIuNjIyNjhlLTA3TDUgNUwtMi4xODU1N2UtMDcgNUwtMi42MjI2OGUtMDcgNkw1IDZMNSAxMUw2IDExTDYgNkwxMSA2WiIgZmlsbD0iYmxhY2siLz4KPC9zdmc+Cg==';
	const zoomInIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTEiIGhlaWdodD0iMTEiIHZpZXdCb3g9IjAgMCAxMSAxMSIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0xMSA2TDExIDVMNiA1TDYgLTIuMTg1NTdlLTA3TDUgLTIuNjIyNjhlLTA3TDUgNUwtMi4xODU1N2UtMDcgNUwtMi42MjI2OGUtMDcgNkw1IDZMNSAxMUw2IDExTDYgNkwxMSA2WiIgZmlsbD0id2hpdGUiLz4KPC9zdmc+Cg==';
	const zoomOutIcon_dark = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMSIgdmlld0JveD0iMCAwIDEwIDEiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxyZWN0IHg9IjEwIiB3aWR0aD0iMSIgaGVpZ2h0PSIxMCIgdHJhbnNmb3JtPSJyb3RhdGUoOTAgMTAgMCkiIGZpbGw9ImJsYWNrIi8+Cjwvc3ZnPgo=';
	const zoomOutIcon_light = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMSIgdmlld0JveD0iMCAwIDEwIDEiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxyZWN0IHg9IjEwIiB3aWR0aD0iMSIgaGVpZ2h0PSIxMCIgdHJhbnNmb3JtPSJyb3RhdGUoOTAgMTAgMCkiIGZpbGw9IndoaXRlIi8+Cjwvc3ZnPgo=';

	const getIconsForLoad = function () {
		return [
			clickEffectIcon, afterEffectIcon,
			entrEffectIcon, emphEffectIcon, exitEffectIcon,
			pathEffectIcon_dark, pathEffectIcon_light,
			playIcon_dark, playIcon_light,
			stopIcon_dark, stopIcon_light,
			arrowUpIcon_dark, arrowUpIcon_light, 
			arrowDownIcon_dark, arrowDownIcon_light,
			closeIcon_dark, closeIcon_light,
			menuButtonIcon_dark, menuButtonIcon_light,
			arrowLeftIcon_dark, arrowLeftIcon_light,
			arrowRightIcon_dark, arrowRightIcon_light,
			zoomInIcon_dark, zoomInIcon_light,
			zoomOutIcon_dark, zoomOutIcon_light,
		];
	}


	// EXPORTS
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CAnimPaneHeader = CAnimPaneHeader;
	window['AscCommon'].CSeqListContainer = CSeqListContainer;
	window['AscCommon'].CTimelineContainer = CTimelineContainer;

	window['AscCommon'].getIconsForLoad = getIconsForLoad;
})(window);

