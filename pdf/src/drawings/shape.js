/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(){

    /**
	 * Class representing a pdf text shape.
	 * @constructor
    */
    function CPdfShape() {
        AscFormat.CShape.call(this);
        AscPDF.CPdfDrawingPrototype.call(this);
    }
    
    CPdfShape.prototype.constructor = CPdfShape;
    CPdfShape.prototype = Object.create(AscFormat.CShape.prototype);
    Object.assign(CPdfShape.prototype, AscPDF.CPdfDrawingPrototype.prototype);
    
    CPdfShape.prototype.IsShape = function() {
        return true;
    };
    CPdfShape.prototype.SetEditField = function(oForm) {
        this.editField = oForm;
    };
    CPdfShape.prototype.IsEditFieldShape = function() {
        return !!this.editField;
    };
    CPdfShape.prototype.GetEditField = function() {
        return this.editField;
    };
    CPdfShape.prototype.hitInTextRect = function(x, y) {
        if (this.IsEditFieldShape()) {
            return false;
        }
        
        return this.hitInTextRectWord(x, y);
    };
    CPdfShape.prototype.GetDocument = function() {
        if (this.IsEditFieldShape()) {
            return this.editField.GetDocument();
        }
        else {
            return AscPDF.CPdfDrawingPrototype.prototype.GetDocument.call(this);
        }
    };
    CPdfShape.prototype.GetPage = function() {
        if (this.IsEditFieldShape()) {
            return this.editField.GetPage();
        }
        else {
            return AscPDF.CPdfDrawingPrototype.prototype.GetPage.call(this);
        }
    };
    CPdfShape.prototype.ShouldDrawImaginaryBorder = function(graphicsWord) {
        let bDraw = !!(this.spPr && (!this.spPr.Fill || this.spPr.hasNoFill()) && !(this.pen && this.pen.Fill && this.pen.Fill.fill && !(this.pen.Fill.fill instanceof AscFormat.CNoFill)));
        bDraw = bDraw && this.IsFromScan();
        bDraw = bDraw && !Asc.editor.isRestrictionView();
        bDraw = bDraw && !graphicsWord.isThumbnails;
    
        return bDraw;
    };
    CPdfShape.prototype.CheckTextOnOpen = function() {
        let oContent = this.GetDocContent();
        if (oContent) {
            oContent.SetApplyToAll(true);
            AscFonts.FontPickerByCharacter.getFontsByString(oContent.GetSelectedText());
            oContent.SetApplyToAll(false);
        }
    };
    CPdfShape.prototype.canRotate = function () {
        if (this.cropObject) {
            return false;
        }
        
        if (this.signatureLine) {
            return false;
        }

        if (!this.canEdit()) {
			return false;
		}

        let oGroup = this.getMainGroup();
        if (oGroup && oGroup.IsAnnot()) {
            return false;
        }

        if (this.IsEditFieldShape()) {
            return false;
        }

		return this.getNoRot() === false;
    };
    CPdfShape.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculate();
        this.recalculateShdw();
        this.SetNeedRecalc(false);
    };
    CPdfShape.prototype.recalculateBounds = function(bLine) {
        let boundsChecker = new AscFormat.CSlideBoundsChecker();
        
        bLine && boundsChecker.CheckLineWidth(this);
        boundsChecker.DO_NOT_DRAW_ANIM_LABEL = true;
        this.draw(boundsChecker);
        boundsChecker.CorrectBounds();

        this.bounds.x = boundsChecker.Bounds.min_x;
        this.bounds.y = boundsChecker.Bounds.min_y;
        this.bounds.l = boundsChecker.Bounds.min_x;
        this.bounds.t = boundsChecker.Bounds.min_y;
        this.bounds.r = boundsChecker.Bounds.max_x;
        this.bounds.b = boundsChecker.Bounds.max_y;
        this.bounds.w = boundsChecker.Bounds.max_x - boundsChecker.Bounds.min_x;
        this.bounds.h = boundsChecker.Bounds.max_y - boundsChecker.Bounds.min_y;
    };
    CPdfShape.prototype.onMouseDown = function(x, y, e) {
        let oDoc                = this.GetDocument();
        let oViewer             = oDoc.Viewer;
        let oDrawingObjects     = oDoc.Viewer.DrawingObjects;
        this.selectStartPage    = this.GetPage();

        // координаты клика на странице в MM
        var pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
		let docContent = this.GetDocContent();
		if (docContent)
			docContent.RecalculateCurPos();
    };
    CPdfShape.prototype.GetDocContent = function() {
        return this.getDocContent();
    };
    CPdfShape.prototype.createTextBody = function () {
        AscFormat.CShape.prototype.createTextBody.call(this);
        this.SetNeedRecalc(true);
    };

    CPdfShape.prototype.getTrackGeometry = function () {
        let oGroup = this.getMainGroup();

        // заглушка для трека геометрии с клауд бордером для FreeText
        if (oGroup && oGroup.IsAnnot && oGroup.IsAnnot() && oGroup.GetTextBoxShape() == this) {
            return AscFormat.ExecuteNoHistory(
                function () {
                    var _ret = AscFormat.CreateGeometry("rect");
                    _ret.Recalculate(this.extX, this.extY);
                    return _ret;
                }, this, []
            );
        }

		const oOwnGeometry = this.getGeometry();
		if(oOwnGeometry) {
			return oOwnGeometry;
		}
		if(this.rectGeometry) {
			return this.rectGeometry;
		}
		return AscFormat.ExecuteNoHistory(
			function () {
				var _ret = AscFormat.CreateGeometry("rect");
				_ret.Recalculate(this.extX, this.extY);
				return _ret;
			}, this, []
		);
	};
    CPdfShape.prototype.onMouseUp = function(x, y, e) {
        let oViewer         = Asc.editor.getDocumentRenderer();
        
        this.selectStartPage    = this.GetPage();
        let oContent            = this.GetDocContent();

        if (global_mouseEvent.ClickCount == 2) {
            oContent.SelectAll();
            if (oContent.IsSelectionEmpty() == false)
                oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
            else
                oContent.RemoveSelection();
        }
                
        if (oContent.IsSelectionEmpty()) {
			oContent.RemoveSelection();
			oContent.RecalculateCurPos();
		}
    };
	CPdfShape.prototype.GetAllFonts = function(fontMap) {
		fontMap = fontMap || {};
		
		let docContent = this.GetDocContent();
		if (!docContent)
			return fontMap;
		
		for (let i = 0, count = docContent.GetElementsCount(); i < count; ++i) {
			let para = docContent.GetElement(i);
			if (!para || !para.IsParagraph())
				continue;
			
			para.Get_CompiledPr2(false).TextPr.Document_Get_AllFontNames(fontMap);
			
			if (para.Pr.Bullet)
				para.Pr.Bullet.Get_AllFontNames(fontMap);
			
			if (para.Pr.DefaultRunPr)
				para.Pr.DefaultRunPr.Document_Get_AllFontNames(fontMap);
			
			para.CheckRunContent(function(run) {
				run.Get_CompiledPr(false).Document_Get_AllFontNames(fontMap);
			});
		}
		
		delete fontMap["+mj-lt"];
		delete fontMap["+mn-lt"];
		delete fontMap["+mj-ea"];
		delete fontMap["+mn-ea"];
		delete fontMap["+mj-cs"];
		delete fontMap["+mn-cs"];
		
		return fontMap;
	};

    CPdfShape.prototype.hitToAdjustment = function (x, y) {
        if (!AscFormat.canSelectDrawing(this)) {
            return false;
        }
        var oApi = Asc.editor || editor;
        var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
        let isViewerObj = oApi.getPDFDoc().IsViewerObject(this);

        if (isViewerObj) {
            isDrawHandles = true;
        }

        if (isDrawHandles === false) {
            return {hit: false, adjPolarFlag: null, adjNum: null, warp: false};
        }
        var invert_transform;
        var t_x, t_y, ret;
        var _calcGeom = this.getGeometry();
        var _dist;
        if (global_mouseEvent && global_mouseEvent.AscHitToHandlesEpsilon) {
            _dist = global_mouseEvent.AscHitToHandlesEpsilon;
        } else {
            _dist = this.convertPixToMM(global_mouseEvent.KoefPixToMM * AscCommon.TRACK_CIRCLE_RADIUS);
        }
        if (_calcGeom) {
            invert_transform = this.getInvertTransform();
            if (!invert_transform) {
                return {hit: false, adjPolarFlag: null, adjNum: null, warp: false};
            }
            t_x = invert_transform.TransformPointX(x, y);
            t_y = invert_transform.TransformPointY(x, y);
            ret = _calcGeom.hitToAdj(t_x, t_y, _dist);
            if (ret.hit) {
                ret.warp = false;
                return ret;
            }
        }
        if (this.recalcInfo.warpGeometry && this.invertTransformTextWordArt) {
            invert_transform = this.invertTransformTextWordArt;
            t_x = invert_transform.TransformPointX(x, y);
            t_y = invert_transform.TransformPointY(x, y);
            ret = this.recalcInfo.warpGeometry.hitToAdj(t_x, t_y, _dist);
            ret.warp = true;
            return ret;
        }

        return {hit: false, adjPolarFlag: null, adjNum: null, warp: false};
    };
    CPdfShape.prototype.canMove = function () {
		var oApi = Asc.editor || editor;
		var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;

        let oEditField = this.GetEditField();
        if (oEditField && oEditField.IsLocked()) {
            return false;
        }
        
        if (oApi.getPDFDoc().IsViewerObject(this))
            return true;

		if (isDrawHandles === false) {
			return false;
		}
		if (!this.canEdit()) {
			return false;
		}
		return this.getNoMove() === false;
	};
    CPdfShape.prototype.canResize = function () {
        let oEditField = this.GetEditField();
        if (oEditField && oEditField.IsLocked()) {
            return false;
        }

		return AscFormat.CGraphicObjectBase.prototype.canResize.call(this);
	};
    CPdfShape.prototype.ConvertToAnnot = function() {
        let oDoc = Asc.editor.getPDFDoc();
        let oXfrm = this.getXfrm();

        let oAnnot;
        let aRect = [oXfrm.offX * g_dKoef_mm_to_pt, oXfrm.offY * g_dKoef_mm_to_pt, (oXfrm.offX + oXfrm.extX) * g_dKoef_mm_to_pt, (oXfrm.offY + oXfrm.extY) * g_dKoef_mm_to_pt];
        let sCreationDate = (new Date().getTime()).toString();
        let nLineW = this.pen.w / 36000.0 * g_dKoef_mm_to_pt;
        let oRGAStroke = this.pen.Fill.fill.color.RGBA;
        let aStrokeColor = [oRGAStroke.R / 255, oRGAStroke.G / 255, oRGAStroke.B / 255];
        let oRGBAFill = !this.brush.isNoFill() ? this.brush.fill.color.RGBA : null;
        let aFillColor = oRGBAFill ? [oRGBAFill.R / 255, oRGBAFill.G / 255, oRGBAFill.B / 255] : null;

        function rotateRect(rect, rad) {
            const [x1, y1, x2, y2] = rect;
            const cx = (x1 + x2) * 0.5, cy = (y1 + y2) * 0.5;
            const hw = (x2 - x1) * 0.5, hh = (y2 - y1) * 0.5;

            const c = Math.cos(rad), s = Math.sin(rad);
            const newHW = Math.abs(hw * c) + Math.abs(hh * s);
            const newHH = Math.abs(hw * s) + Math.abs(hh * c);

            return [cx - newHW, cy - newHH, cx + newHW, cy + newHH];
        }

        let oProps = {
            rect:           this.rot ? rotateRect(aRect, this.rot) : aRect,
            contents:       null,
            creationDate:   sCreationDate,
            modDate:        sCreationDate,
            hidden:         false
        }

        let oGeometry = this.getGeometry();
        switch (oGeometry.preset) {
            case "rect":
                oProps.type = AscPDF.ANNOTATIONS_TYPES.Square;
                break;
            case "ellipse":
                oProps.type = AscPDF.ANNOTATIONS_TYPES.Circle;
                break;
            case "line":
                oProps.type = AscPDF.ANNOTATIONS_TYPES.Line;
                break;
            default: {
                let aCommands = oGeometry.pathLst[0].ArrPathCommand;
                
                function isPolyLine(commands) {
                    let count0 = 0;

                    for (let i = 0; i < commands.length; i++) {
                        if (commands[i].id === AscFormat.moveTo) {
                            count0++;

                            if (count0 > 1) {
                                return false;
                            }
                        }
                        else if (commands[i].id !== AscFormat.lineTo) {
                            return false;
                        }
                    }

                    return count0 === 1;
                };

                function isPolygon(commands) {
                    let count0 = 0;

                    for (let i = 0; i < commands.length; i++) {
                        if (commands[i].id === AscFormat.moveTo) {
                            count0++;

                            if (count0 > 1) {
                                return false;
                            }
                        }
                        else if (commands[i].id !== AscFormat.lineTo && commands[i].id !== AscFormat.close) {
                            return false;
                        }
                    }

                    return count0 === 1;
                };

                let isCanConvert = oGeometry.pathLst.length == 1;
                if (!isCanConvert) {
                    return null;
                }
                else if (isPolyLine(aCommands)) {
                    oProps.type = AscPDF.ANNOTATIONS_TYPES.PolyLine;
                }
                else if (isPolygon(aCommands)) {
                    oProps.type = AscPDF.ANNOTATIONS_TYPES.Polygon;
                }
                break;
            }
        }

        oAnnot = AscPDF.CreateAnnotByProps(oProps, oDoc);
        oAnnot.SetWidth(nLineW);
        oAnnot.SetStrokeColor(aStrokeColor);
        aFillColor && oAnnot.SetFillColor(aFillColor);

        switch (oAnnot.GetType()) {
            case AscPDF.ANNOTATIONS_TYPES.Square:
            case AscPDF.ANNOTATIONS_TYPES.Circle: {
                let nLineW = oAnnot.GetWidth() * g_dKoef_pt_to_mm;
                oAnnot.recalcBounds();
                oAnnot.recalcGeometry();
                oAnnot.Recalculate(true);
                
                let oGrBounds = oAnnot.bounds;
                let oShapeBounds = oAnnot.getRectBounds();

                aRect[0] = Math.round(oGrBounds.l - nLineW) * g_dKoef_mm_to_pt;
                aRect[1] = Math.round(oGrBounds.t - nLineW) * g_dKoef_mm_to_pt;
                aRect[2] = Math.round(oGrBounds.r + nLineW) * g_dKoef_mm_to_pt;
                aRect[3] = Math.round(oGrBounds.b + nLineW) * g_dKoef_mm_to_pt;

                oAnnot.SetRect(aRect);
                oAnnot.SetRectangleDiff([
                    Math.round(oShapeBounds.l - oGrBounds.l + nLineW) * g_dKoef_mm_to_pt,
                    Math.round(oShapeBounds.t - oGrBounds.t + nLineW) * g_dKoef_mm_to_pt,
                    Math.round(oGrBounds.r - oShapeBounds.r + nLineW) * g_dKoef_mm_to_pt,
                    Math.round(oGrBounds.b - oShapeBounds.b + nLineW) * g_dKoef_mm_to_pt
                ]);
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Line: {
                oAnnot.SetLineEnd(AscPDF.LINE_END_TYPE.OpenArrow);
                let aCommands = oGeometry.pathLst[0].ArrPathCommand;
                let oPt1 = { X: aCommands[0].X, Y: aCommands[0].Y };
                let oPt2 = { X: aCommands[1].X, Y: aCommands[1].Y };

                if (this.flipH) {
                    oPt1.X = oXfrm.extX - oPt1.X;
                    oPt2.X = oXfrm.extX - oPt2.X;
                }
                if (this.flipV) {
                    oPt1.Y = oXfrm.extY - oPt1.Y;
                    oPt2.Y = oXfrm.extY - oPt2.Y;
                }

                let aLinePoints = [];
                let oTranform   = oAnnot.transform;
                // считаем новые точки linePoints (в оригинальных координатах - в пикселях, без скейлов)
                aLinePoints.push(oTranform.TransformPointX(oPt1.X, 0) * g_dKoef_mm_to_pt);
                aLinePoints.push(oTranform.TransformPointY(0, oPt1.Y) * g_dKoef_mm_to_pt);
                aLinePoints.push(oTranform.TransformPointX(oPt2.X, 0) * g_dKoef_mm_to_pt);
                aLinePoints.push(oTranform.TransformPointY(0, oPt2.Y) * g_dKoef_mm_to_pt);

                oAnnot.SetLinePoints(aLinePoints);
                oAnnot.SetRect(oAnnot.private_CalcBoundingRect());

                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.PolyLine: {
                let aCommands = this.getGeometry().pathLst[0].ArrPathCommand;
                let oTranform = oAnnot.transform;
                let aVertices = [];

                aCommands.forEach(function(command) {
                    let oPt = { X: command.X, Y: command.Y };
                    aVertices.push(oTranform.TransformPointX(oPt.X, 0) * g_dKoef_mm_to_pt);
                    aVertices.push(oTranform.TransformPointY(0, oPt.Y) * g_dKoef_mm_to_pt);
                });

                oAnnot.SetVertices(aVertices);
                
                aRect[0] -= nLineW;
                aRect[1] -= nLineW;
                aRect[2] += nLineW;
                aRect[3] += nLineW;

                oAnnot.SetRect(aRect);
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Polygon: {
                let aCommands = this.getGeometry().pathLst[0].ArrPathCommand;
                let oTranform = oAnnot.transform;
                let aVertices = [];

                aCommands.forEach(function(command) {
                    let oPt = { X: command.X, Y: command.Y };
                    if (command.id == AscFormat.close) {
                        oPt = { X: aCommands[0].X, Y: aCommands[0].Y };
                    }
                    
                    aVertices.push(oTranform.TransformPointX(oPt.X, 0) * g_dKoef_mm_to_pt);
                    aVertices.push(oTranform.TransformPointY(0, oPt.Y) * g_dKoef_mm_to_pt);
                });

                oAnnot.SetVertices(aVertices);
                
                aRect[0] -= nLineW;
                aRect[1] -= nLineW;
                aRect[2] += nLineW;
                aRect[3] += nLineW;

                oAnnot.SetRect(aRect);
                break;
            }
        }

        return oAnnot;
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfShape.prototype.updateSelectionState = function () {
        var drawing_document = this.getDrawingDocument();

        if (drawing_document) {
            var content = this.getDocContent();
            if (content) {
                var oMatrix = null;
                if (this.transformText) {
                    oMatrix = this.transformText.CreateDublicate();
                }
                drawing_document.UpdateTargetTransform(oMatrix);
                if (true === content.IsSelectionUse()) {
                    // Выделение нумерации
                    if (selectionflag_Numbering == content.Selection.Flag) {
                        drawing_document.TargetEnd();
                    }
                    // Обрабатываем движение границы у таблиц
                    else if (null != content.Selection.Data && true === content.Selection.Data.TableBorder && type_Table == content.Content[content.Selection.Data.Pos].GetType()) {
                        // Убираем курсор, если он был
                        drawing_document.TargetEnd();
                    } else {
                        if (false === content.IsSelectionEmpty()) {
                            drawing_document.Overlay && content.DrawSelectionOnPage(0);
                            drawing_document.TargetEnd();
                        } else {
                            if (true !== content.Selection.Start) {
                                content.RemoveSelection();
                            }
                            content.RecalculateCurPos();

                            drawing_document.TargetStart(true);
                        }
                    }
                } else {
                    content.RecalculateCurPos();

                    drawing_document.TargetStart(true);
                }
            } else {
                drawing_document.UpdateTargetTransform(new AscCommon.CMatrix());
                drawing_document.TargetEnd();
            }
        }
    };
    CPdfShape.prototype.Set_CurrentElement = function(bUpdate, pageIndex) {
        let oDoc = this.GetDocument();
        let oController = oDoc.GetController();

        this.SetControllerTextSelection(oController, this.GetPage());

        let oGroup = this.getMainGroup();
        if (!oGroup)
            oDoc.SetMouseDownObject(this);
    };
    CPdfShape.prototype.setRecalculateInfo = function() {
        this.recalcInfo =
        {
            recalculateContent:        true,
            recalculateBrush:          true,
            recalculatePen:            true,
            recalculateTransform:      true,
            recalculateTransformText:  true,
            recalculateBounds:         true,
            recalculateGeometry:       true,
            recalculateStyle:          true,
            recalculateFill:           true,
            recalculateLine:           true,
            recalculateTransparent:    true,
            recalculateTextStyles:     [true, true, true, true, true, true, true, true, true],
            recalculateContent2: true,
            oContentMetrics: null

        };
        this.compiledStyles = [];
        this.lockType = AscCommon.c_oAscLockTypes.kLockTypeNone;
    };
    CPdfShape.prototype.copy = function (oPr) {
        let copy = new CPdfShape();
        this.fillObject(copy, oPr);

        if ((!oPr || !oPr.bSkipRedactsIds) && this.GetRedactIds) {
            this.GetRedactIds().forEach(function(id) {
                copy.AddRedactId(id);
            });
        }
        
        return copy;
    };
    window["AscPDF"].CPdfShape = CPdfShape;
})();

