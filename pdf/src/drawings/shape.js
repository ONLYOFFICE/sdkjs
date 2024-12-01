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
    }
    
    CPdfShape.prototype.constructor = CPdfShape;
    CPdfShape.prototype = Object.create(AscFormat.CShape.prototype);
    Object.assign(CPdfShape.prototype, AscPDF.PdfDrawingPrototype.prototype);
    
    CPdfShape.prototype.IsShape = function() {
        return true;
    };
    CPdfShape.prototype.ShouldDrawImaginaryBorder = function(graphicsWord) {
        let bDraw = !!(this.spPr && this.spPr.hasNoFill() && !(this.pen && this.pen.Fill && this.pen.Fill.fill && !(this.pen.Fill.fill instanceof AscFormat.CNoFill)));
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

        if (this.group && this.group.IsAnnot()) {
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
    CPdfShape.prototype.recalculateBounds = function() {
        let boundsChecker = new AscFormat.CSlideBoundsChecker();
        
        // boundsChecker.CheckLineWidth(this);
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
        // заглушка для трека геометрии с клауд бордером для FreeText
        if (this.group && this.group.IsAnnot && this.group.IsAnnot() && this.group.GetTextBoxShape() == this) {
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
        if (!this.group)
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
        return copy;
    };
    window["AscPDF"].CPdfShape = CPdfShape;
})();

