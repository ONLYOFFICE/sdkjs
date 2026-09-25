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
	function isShowParaMarksInTextArt(shape, bCurrentEditText) {
		if(!Asc.editor.ShowParaMarks) return false;
		if(bCurrentEditText) return true;
		if(shape.chekBodyPrTransform(shape.getBodyPr())) return false;
		return true;
	}

	function GraphicObjectDrawer(graphicObject, graphics, transform, transformText, pageIndex, options) {
		this.graphicObject = graphicObject;
		this.graphics = graphics;
		this.transform = transform;
		this.transformText = transformText;
		this.pageIndex = pageIndex;
		this.options = options;
		this.state = {
			clipRect: null
		}
	}
	GraphicObjectDrawer.prototype.getPageIndex = function () {
		return this.pageIndex;
	};
	GraphicObjectDrawer.prototype.checkSignature = function () {
		const graphicObject = this.graphicObject;
		if (graphicObject.signatureLine) {
			let signatureUrl = null;

			const api = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
			if (api) {
				signatureUrl = api.asc_getSignatureImage(graphicObject.signatureLine.id);
			}
			if (typeof signatureUrl === "string" && signatureUrl.length > 0) {
				graphicObject.brush = AscFormat.CreateBlipFillUniFillFromUrl(signatureUrl);
			}
		}
	};
	GraphicObjectDrawer.prototype.checkHorizontalRule = function () {
		const graphicObject = this.graphicObject;
		const horizontalRule = graphicObject.getHorizontalRule();
		if (horizontalRule) {
			if (horizontalRule.noshade) {
				if (!graphicObject.brush || !graphicObject.brush.fill) {
					graphicObject.brush = AscFormat.CreateSolidFillRGBA(160, 160, 160, 255);
				}
				if (!graphicObject.pen || !graphicObject.pen.Fill || !graphicObject.pen.Fill.fill) {
					graphicObject.pen = null;
				}
			} else {
				const graphics = this.getGraphics();
				graphics.transform3(this.getTransform(), false);
				graphics.SetIntegerGrid(true);
				let hrExtX = graphicObject.extX;
				let hrExtY = graphicObject.extY;
				graphics.p_color(128, 128, 128, 255);
				graphics.drawHorLine(0, 0, 0, hrExtX, 0);
				graphics.drawVerLine(0, 0, 0, hrExtY, 0);
				graphics.p_color(212, 208, 200, 255);
				graphics.drawHorLine(2, hrExtY, 0, hrExtX, 0);
				graphics.drawVerLine(2, hrExtX, 0, hrExtY, 0);
				graphicObject.brush = null;
				graphicObject.pen = null;
			}
		}
	};
	GraphicObjectDrawer.prototype.isTargetTextObject = function () {
		const graphicObject = this.graphicObject;
		const controller = graphicObject.getDrawingObjectsController && graphicObject.getDrawingObjectsController();
		return controller && AscFormat.getTargetTextObject(controller) === graphicObject;
	};
	GraphicObjectDrawer.prototype.isPdfAnnotations = function () {
		const graphicObject = this.graphicObject;
		if (Asc.editor.isPdfEditor()) {
			if (graphicObject.IsAnnot && graphicObject.IsAnnot()) {
				return true;
			}
		}
		return false;
	};
	GraphicObjectDrawer.prototype.clipPdfAnnotationsRect = function () {
		const graphicObject = this.graphicObject;
		if (graphicObject instanceof AscFormat.CShape) {
			if (!this.isTargetTextObject() && !this.isPdfAnnotations()) {
				const graphics = this.getGraphics();
				const transform = this.getTransform();
				const transformText = this.getTransformText();
				graphicObject.clipTextRect(graphics, transform, transformText, this.pageIndex);
			}
		}
	};
	GraphicObjectDrawer.prototype.drawTxBody = function () {
		const graphicObject = this.graphicObject;
		const txBody = graphicObject.txBody;
		if (txBody) {
			const graphics = this.getGraphics();
			graphics.SaveGrState();
			graphics.SetIntegerGrid(false);
			const contentTransform = this.getContentTransform();
			this.clipPdfAnnotationsRect();

			graphics.transform3(contentTransform, true);
			if (graphics.CheckUseFonts2 !== undefined)
				graphics.CheckUseFonts2(contentTransform);

			graphics.SetIntegerGrid(true);
			txBody.draw(graphics);
			if (graphics.UncheckUseFonts2 !== undefined)
				graphics.UncheckUseFonts2(contentTransform);
			graphics.RestoreGrState();
		}
	};
	GraphicObjectDrawer.prototype.drawTextBoxContent = function () {
		const graphicObject = this.graphicObject;
		const graphics = this.getGraphics();
		const pageIndex = this.getPageIndex();
		if (graphicObject.textBoxContent && graphics.isSupportTextDraw() && graphicObject.transformText) {
			const oldStartPage = graphicObject.textBoxContent.GetRelativeStartPage();
			graphicObject.textBoxContent.Set_StartPage(pageIndex);

			graphics.SaveGrState();
			graphics.SetIntegerGrid(false);
			graphicObject.clipTextRect(graphics, this.getTransform(), this.getTransformText(), pageIndex);
			let resultPageIndex = AscFormat.isRealNumber(graphics.shapePageIndex) ? graphics.shapePageIndex : oldStartPage;

			if (graphics.CheckUseFonts2 !== undefined)
				graphics.CheckUseFonts2(graphicObject.transformText, graphicObject.isForm && graphicObject.isForm() ? true : false);

			if (AscCommon.IsShapeToImageConverter) {
				graphicObject.textBoxContent.Set_StartPage(0);
				resultPageIndex = 0;
			}

			graphicObject.textBoxContent.Set_StartPage(resultPageIndex);
			graphicObject.textBoxContent.Draw(resultPageIndex, graphics);

			if (graphics.UncheckUseFonts2 !== undefined)
				graphics.UncheckUseFonts2();

			graphicObject.textBoxContent.Set_StartPage(oldStartPage);
			graphics.RestoreGrState();
		}
	};
	GraphicObjectDrawer.prototype.isShowParaMarksInTextArt = function () {
		return isShowParaMarksInTextArt(this.graphicObject, this.isTargetTextObject());
	};
	GraphicObjectDrawer.prototype.drawPlaceholderTextArt = function () {
		const graphics = this.getGraphics();
		const graphicObject = this.graphicObject;
		const theme = graphicObject.getParentObjects().theme;
		const colorMap = graphicObject.Get_ColorMap();
		if (graphics.IsNoDrawingEmptyPlaceholder !== true && graphics.IsNoDrawingEmptyPlaceholderText !== true && !graphics.isPdf()) {
			if (this.isShowParaMarksInTextArt()) {
				graphicObject.txWarpStructParamarks2.draw(graphics, graphicObject.transformTextWordArt2, theme, colorMap);
			} else {
				if (graphicObject.txWarpStruct2)
					graphicObject.txWarpStruct2.draw(graphics, graphicObject.transformTextWordArt2, theme, colorMap);
			}
		}
	};
	GraphicObjectDrawer.prototype.drawTextArtContentWithParaMarks = function () {
		const graphicObject = this.graphicObject;
		const graphics = this.getGraphics();
		const theme = graphicObject.getParentObjects().theme;
		const colorMap = graphicObject.Get_ColorMap();
		if (this.isTargetTextObject() && graphicObject.txWarpStructParamarksNoTransform) {
			graphicObject.txWarpStructParamarksNoTransform.draw(graphics, graphicObject.transformText, theme, colorMap);
		} else if (graphicObject.txWarpStructParamarks) {
			graphicObject.txWarpStructParamarks.draw(graphics, graphicObject.transformTextWordArt, theme, colorMap);
			if (graphicObject.checkNeedRecalcDocContentForTxWarp(graphicObject.bodyPr)) {
				if (graphicObject.txWarpStructParamarksNoTransform) {
					graphicObject.txWarpStructParamarksNoTransform.drawComments(graphics, undefined, graphicObject.transformTextWordArt);
				}
			}
		}
	};
	GraphicObjectDrawer.prototype.drawTextArtContentWithoutParaMarks = function () {
		const graphics = this.getGraphics();
		const graphicObject = this.graphicObject;
		const theme = graphicObject.getParentObjects().theme;
		const colorMap = graphicObject.Get_ColorMap();
		if (this.isTargetTextObject() && graphicObject.txWarpStructNoTransform) {
			graphicObject.txWarpStructNoTransform.draw(graphics, graphicObject.transformText, theme, colorMap);
		} else if (graphicObject.txWarpStruct) {
			graphicObject.txWarpStruct.draw(graphics, graphicObject.transformTextWordArt, theme, colorMap);
			if (graphicObject.checkNeedRecalcDocContentForTxWarp(graphicObject.bodyPr)) {
				if (graphicObject.txWarpStructNoTransform) {
					graphicObject.txWarpStructNoTransform.drawComments(graphics, undefined, graphicObject.transformTextWordArt);
				}
			}
		}
	};
	GraphicObjectDrawer.prototype.drawTextArtContent = function () {
		const graphicObject = this.graphicObject;
		const graphics = this.getGraphics();
		var content = graphicObject.getDocContent();
		var resultPageIndex = AscFormat.isRealNumber(graphics.shapePageIndex) ? graphics.shapePageIndex : (content ? content.GetRelativeStartPage() : 0);
		graphics.PageNum = resultPageIndex;
		var isNeedRestoreState = false;
		if (graphicObject.bWordShape && graphicObject.clipRect) {
			isNeedRestoreState = true;
			var clipRect = graphicObject.clipRect;
			if (!graphicObject.bodyPr.upright) {
				graphics.SaveGrState();
				graphics.SetIntegerGrid(false);
				graphics.transform3(graphicObject.transform);
				graphics.AddClipRect(clipRect.x, clipRect.y, clipRect.w, clipRect.h);
			} else {
				graphics.SaveGrState();
				graphics.SetIntegerGrid(false);
				graphics.transform3(graphicObject.transformText, true);
				graphics.AddClipRect(clipRect.x, clipRect.y, clipRect.w, clipRect.h);
			}
		}

		if (this.isShowParaMarksInTextArt()) {
			this.drawTextArtContentWithParaMarks();
		} else {
			this.drawTextArtContentWithoutParaMarks();
		}
		delete graphics.PageNum;
		if (isNeedRestoreState) {
			graphics.RestoreGrState();
		}
	};
	GraphicObjectDrawer.prototype.drawTextArt = function () {
		const graphicObject = this.graphicObject;
		if (!graphicObject.bWordShape && this.isNeedUsePlaceholderTransform()) {
			this.drawPlaceholderTextArt();
		} else {
			this.drawTextArtContent();
		}
	};
	GraphicObjectDrawer.prototype.isDrawContent = function () {
		const graphicObject = this.graphicObject;
		return !graphicObject.txWarpStruct && !graphicObject.txWarpStructParamarksNoTransform || (!graphicObject.txWarpStructParamarksNoTransform && this.isTargetTextObject() || (!graphicObject.txBody && !graphicObject.textBoxContent));
	};
	GraphicObjectDrawer.prototype.drawContent = function () {
		const graphicObject = this.graphicObject;
		if (!graphicObject.cropObject) {
			if (this.isDrawContent()) {
				this.drawTxBody();
				this.drawTextBoxContent();
			} else {
				this.drawTextArt();
			}
		}
	};
	GraphicObjectDrawer.prototype.drawShape = function () {
		const graphicObject = this.graphicObject;
		const graphics = this.getGraphics();
		const transform = this.getTransform();
		const oldBrush = graphicObject.brush;
		const oldPen = graphicObject.pen;
		const geometry = graphicObject.getGeometry();
		this.checkSignature();
		this.checkHorizontalRule();

		if ((geometry || (graphicObject.getObjectType && (graphicObject.getObjectType() === AscDFH.historyitem_type_DLbl || graphicObject.getObjectType() === AscDFH.historyitem_type_Title || graphicObject.getObjectType() === AscDFH.historyitem_type_Legend))) && (graphicObject.style || (graphicObject.brush && graphicObject.brush.fill) || (graphicObject.pen && graphicObject.pen.Fill && graphicObject.pen.Fill.fill))) {
			graphics.SetIntegerGrid(false);
			graphics.transform3(transform, false);

			var shape_drawer = new AscCommon.CShapeDrawer();
			shape_drawer.fromShape2(graphicObject, graphics, geometry);
			shape_drawer.draw(geometry);
		}

		graphicObject.brush = oldBrush;
		graphicObject.pen = oldPen;

		this.drawContent();
	};

	GraphicObjectDrawer.prototype.getTransform = function () {
		return this.transform || this.graphicObject.transform;
	};
	GraphicObjectDrawer.prototype.getTransformText = function () {
		return this.transformText || this.graphicObject.transformText;
	};
	GraphicObjectDrawer.prototype.getTransformText2 = function () {
		return this.options && this.options.transformText2 || this.graphicObject.transformText2;
	};
	GraphicObjectDrawer.prototype.getContentTransform = function () {
		if (this.isNeedUsePlaceholderTransform()) {
			return this.getTransformText2() || this.getTransformText();
		}
		return this.getTransformText();
	};
	GraphicObjectDrawer.prototype.isNeedUsePlaceholderTransform = function () {
		if (AscCommon.IsShapeToImageConverter) {
			return false;
		}
		const graphicObject = this.graphicObject;

		const isEmptyContent = !graphicObject.txBody.content || graphicObject.txBody.content.Is_Empty();
		if (!isEmptyContent) {
			return false;
		}
		const isHavePlaceholderContent = graphicObject.txBody.content2;
		if (!isHavePlaceholderContent) {
			return false;
		}

		const isCurrentPlaceholder = graphicObject.txBody.checkCurrentPlaceholder();
		if (isCurrentPlaceholder) {
			return false;
		}

		const isEmptyPlaceholder = graphicObject.isEmptyPlaceholder ? graphicObject.isEmptyPlaceholder() : false;
		if (!isEmptyPlaceholder) {
			return false;
		}
		return true;
	};
	GraphicObjectDrawer.prototype.getGeometry = function () {
		return this.graphicObject.getGeometry();
	};
	GraphicObjectDrawer.prototype.getGraphics = function () {
		return this.graphics;
	};
	GraphicObjectDrawer.prototype.isInDrawingArea = function () {
		const graphics = this.getGraphics();
		const graphicObjectBounds = this.graphicObject.bounds;
		const oUR = graphics.updatedRect;
		if (oUR && graphicObjectBounds) {
			if (!oUR.isIntersectOther(graphicObjectBounds)) {
				return false;
			}
		}
		return true;
	};
	GraphicObjectDrawer.prototype.checkBounds = function () {
		const graphics = this.getGraphics();
		const geometry = this.getGeometry();
		const transform = this.getTransform();
		const graphicObject = this.graphicObject;
		graphicObject.drawShdw && graphicObject.drawShdw(graphics);
		graphics.transform3(transform);
		if (!graphicObject.spPr || null == geometry || geometry.isEmpty() || !graphics.IsShapeNeedBounds(geometry.preset)) {
			graphics._s();
			graphics._m(0, 0);
			graphics._l(graphicObject.extX, 0);
			graphics._l(graphicObject.extX, graphicObject.extY);
			graphics._l(0, graphicObject.extY);
			graphics._e();
		} else {
			geometry.check_bounds(graphics, graphicObject);
			const shapeDrawer = new AscCommon.CShapeDrawer();
			shapeDrawer.fromShape2(graphicObject, graphics, geometry);
			if (!shapeDrawer.bIsNoStrokeAttack && shapeDrawer.isArrowPresent()) {
				shapeDrawer.drawArrows(false);
			}
		}

		if (graphics.AutoCheckLineWidth) {
			graphics.CheckLineWidth(graphicObject);
			graphics.CorrectBounds2();
		}

		if (graphicObject.txBody) {
			graphics.SetIntegerGrid(false);

			const contentTransform = this.getContentTransform();
			graphics.transform3(contentTransform);

			if (graphics.CheckUseFonts2 !== undefined)
				graphics.CheckUseFonts2(contentTransform);
			graphicObject.txBody.draw(graphics);
			if (graphics.UncheckUseFonts2 !== undefined)
				graphics.UncheckUseFonts2(contentTransform);
			graphics.SetIntegerGrid(true);
		}

		graphics.reset();
	};
	GraphicObjectDrawer.prototype.drawGraphicObject = function () {
		const graphicObject = this.graphicObject;
		if (graphicObject.drawGraphicObject) {
			graphicObject.drawGraphicObject(this);
		} else {
			this.drawShape();
		}
	};
	GraphicObjectDrawer.prototype.isNeedRecalculateGraphicObject = function () {
		return this.graphicObject.checkNeedRecalculate && this.graphicObject.checkNeedRecalculate();
	};
	GraphicObjectDrawer.prototype.isNeedDrawAnimations = function () {
		const graphics = this.getGraphics();
		return !!graphics.animationDrawer;
	};
	GraphicObjectDrawer.prototype.drawAnimations = function () {
		const graphics = this.getGraphics();
		graphics.animationDrawer.drawObject(this.graphicObject, graphics);
	};
	GraphicObjectDrawer.prototype.startDrawShape = function () {
		const graphics = this.getGraphics();
		if (graphics.StartDrawShape) {
			const graphicObject = this.graphicObject;
			graphics.StartDrawShape(undefined, !!(graphicObject.isForm && graphicObject.isForm()));
		}
	};
	GraphicObjectDrawer.prototype.endDrawShape = function () {
		const graphics = this.getGraphics();
		if (graphics.EndDrawShape) {
			graphics.EndDrawShape();
		}
	};
	GraphicObjectDrawer.prototype.getClipRect = function () {
		return this.graphicObject.getClipRect && this.graphicObject.getClipRect();
	};
	GraphicObjectDrawer.prototype.startClipRect = function () {
		const graphics = this.getGraphics();
		if (!graphics.isBoundsChecker()) {
			const clipRect = this.getClipRect();
			if (clipRect) {
				graphics.SaveGrState();
				graphics.AddClipRect(clipRect.x, clipRect.y, clipRect.w, clipRect.h);
				this.state.clipRect = clipRect;
			}
		}
	};
	GraphicObjectDrawer.prototype.endClipRect = function () {
		if (this.state.clipRect) {
			this.state.clipRect = null;
			const graphics = this.getGraphics();
			graphics.RestoreGrState();
		}
	};
	GraphicObjectDrawer.prototype.isNeedDrawPlaceholderRect = function () {
		const graphics = this.getGraphics();
		if (graphics.IsPrintPreview) {
			return false;
		}
		if (graphics.isSmartArtPreviewDrawer) {
			return false;
		}
		if (graphics.isPdf()) {
			return false;
		}
		if (AscCommon.IsShapeToImageConverter) {
			return false;
		}
		const graphicObject = this.graphicObject;
		let isMasterPlaceholder = false;
		const viewManager = Asc.editor && Asc.editor.presentationViewManager;
		if (viewManager && viewManager.isMasterPlaceholderShape(graphicObject)) {
			isMasterPlaceholder = true;
		}

		if (graphics.IsNoDrawingEmptyPlaceholder === true && !isMasterPlaceholder) {
			return false;
		}
		/*(!(this.pen && this.pen.Fill && this.pen.Fill.fill && !(this.pen.Fill.fill instanceof AscFormat.CNoFill)) && AscCommon.IS_GENERATE_SMARTART_AND_TEXT_ON_OPEN) || */
		return (!graphicObject.bWordShape && (graphicObject.isEmptyPlaceholder() && !graphicObject.isObjectInSmartArt() || isMasterPlaceholder) && !graphicObject.isSlideNoteShape() && !(graphicObject.pen && graphicObject.pen.isVisible()))
			|| (Asc.editor.isPdfEditor() && graphicObject.IsDrawing && graphicObject.IsDrawing() && graphicObject.ShouldDrawImaginaryBorder(graphics));
	};
	GraphicObjectDrawer.prototype.isNeedDrawDashPlaceholderRect = function () {
		const graphics = this.getGraphics();
		return typeof editor !== "undefined" && editor && graphics.m_oContext !== undefined && graphics.m_oContext !== null && !graphics.isTrack() && (Asc.editor.isPdfEditor() || !this.isTargetTextObject())
	};
	GraphicObjectDrawer.prototype.drawPlaceholderSmartDashRect = function () {
		const graphics = this.getGraphics();
		const transform = this.getTransform();
		const graphicObject = this.graphicObject;
		graphics.transform3(transform, false);
		const fullTransform = graphics.m_oFullTransform;
		graphics.SetIntegerGrid(true);

		const _x = fullTransform.TransformPointX(0, 0);
		const _y = fullTransform.TransformPointY(0, 0);
		const _r = fullTransform.TransformPointX(graphicObject.extX, graphicObject.extY);
		const _b = fullTransform.TransformPointY(graphicObject.extX, graphicObject.extY);

		const __x = Math.min(_x, _r);
		const __y = Math.min(_y, _b);
		const __r = Math.max(_x, _r);
		const __b = Math.max(_y, _b);
		graphics.m_oContext.lineWidth = 1;
		graphics.p_color(127, 127, 127, 255);

		graphics._s();
		editor.WordControl.m_oDrawingDocument.AutoShapesTrack.AddRectDashClever(graphics.m_oContext, __x >> 0, __y >> 0, __r >> 0, __b >> 0, 2, 2, true);
		graphics._s();
	};
	GraphicObjectDrawer.prototype.drawPlaceholderRotatedDashRect = function () {
		const graphics = this.getGraphics();
		const transform = this.getTransform();
		const graphicObject = this.graphicObject;
		graphics.transform3(transform, false);
		const fullTransform = graphics.m_oFullTransform;
		graphics.SetIntegerGrid(true);

		const _r = graphicObject.extX;
		const _b = graphicObject.extY;

		const x1 = fullTransform.TransformPointX(0, 0) >> 0;
		const y1 = fullTransform.TransformPointY(0, 0) >> 0;

		const x2 = fullTransform.TransformPointX(_r, 0) >> 0;
		const y2 = fullTransform.TransformPointY(_r, 0) >> 0;

		const x3 = fullTransform.TransformPointX(0, _b) >> 0;
		const y3 = fullTransform.TransformPointY(0, _b) >> 0;

		const x4 = fullTransform.TransformPointX(_r, _b) >> 0;
		const y4 = fullTransform.TransformPointY(_r, _b) >> 0;

		graphics.m_oContext.lineWidth = 1;
		graphics.p_color(127, 127, 127, 255);

		graphics._s();
		editor.WordControl.m_oDrawingDocument.AutoShapesTrack.AddRectDash(graphics.m_oContext, x1, y1, x2, y2, x3, y3, x4, y4, 3, 1, true);
		graphics._s();
	};
	GraphicObjectDrawer.prototype.drawPlaceholderSolidRect = function () {
		const graphics = this.getGraphics();
		const transform = this.getTransform();
		const graphicObject = this.graphicObject;
		graphics.SetIntegerGrid(false);
		graphics.p_width(70);
		graphics.transform3(transform, false);
		graphics.p_color(0, 0, 0, 255);
		graphics._s();
		graphics._m(0, 0);
		graphics._l(graphicObject.extX, 0);
		graphics._l(graphicObject.extX, graphicObject.extY);
		graphics._l(0, graphicObject.extY);
		graphics._z();
		graphics.ds();

		graphics.SetIntegerGrid(true);
	};
	GraphicObjectDrawer.prototype.drawPlaceholderDashRect = function () {
		const transform = this.getTransform();
		const graphics = this.getGraphics();
		var angle = transform.GetRotation();
		if (!graphics.isSlidePaceholder &&
			(AscFormat.fApproxEqual(angle, 0.0, 0.0) ||
				AscFormat.fApproxEqual(angle, 90.0, 0.0) ||
				AscFormat.fApproxEqual(angle, 180.0, 0.0) ||
				AscFormat.fApproxEqual(angle, 270.0, 0.0))) {
			this.drawPlaceholderSmartDashRect();
		} else {
			this.drawPlaceholderRotatedDashRect();
		}
	};
	GraphicObjectDrawer.prototype.drawPlaceholderRect = function () {
		if (this.isNeedDrawPlaceholderRect()) {
			if (this.isNeedDrawDashPlaceholderRect()) {
				this.drawPlaceholderDashRect();
			} else {
				this.drawPlaceholderSolidRect();
			}
		}
	};
	GraphicObjectDrawer.prototype.drawShadow = function () {
		const graphicObject = this.graphicObject;
		const graphics = this.getGraphics();
		graphicObject.drawShdw && graphicObject.drawShdw(graphics);
	};
	GraphicObjectDrawer.prototype.drawLocks = function () {
		const graphicObject = this.graphicObject;
		const graphics = this.getGraphics();
		const transform = this.getTransform();
		graphicObject.drawLocks && graphicObject.drawLocks(transform, graphics);
	};
	GraphicObjectDrawer.prototype.draw = function () {
		if (this.isNeedRecalculateGraphicObject()) {
			return;
		}

		if (!this.isInDrawingArea()) {
			return;
		}

		if (this.isNeedDrawAnimations()) {
			this.drawAnimations();
			return;
		}

		const graphics = this.getGraphics();
		if (graphics.isBoundsChecker()) {
			this.checkBounds();
			return;
		}

		this.startDrawShape();
		this.startClipRect();

		this.drawShadow();

		this.drawGraphicObject();

		this.endClipRect();

		this.drawPlaceholderRect();

		this.drawLocks();

		graphics.SetIntegerGrid(true);
		graphics.reset();

		this.endDrawShape();
	};

	function drawWithGraphicObjectDrawer(graphicObject, graphics, transform, transformText, pageIndex, opt) {
		const drawer = new GraphicObjectDrawer(graphicObject, graphics, transform, transformText, pageIndex, opt);
		drawer.draw();
	}
	function drawTxBodyWithGraphicObjectDrawer(graphicObject, graphics, transform, transformText, pageIndex, opt) {
		const drawer = new GraphicObjectDrawer(graphicObject, graphics, transform, transformText, pageIndex, opt);
		drawer.drawTxBody();
	}
	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].drawWithGraphicObjectDrawer = drawWithGraphicObjectDrawer;
	window["AscCommon"].drawTxBodyWithGraphicObjectDrawer = drawTxBodyWithGraphicObjectDrawer;
})();
