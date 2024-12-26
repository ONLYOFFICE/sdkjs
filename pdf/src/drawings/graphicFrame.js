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
    function CPdfGraphicFrame()
    {
        AscFormat.CGraphicFrame.call(this);
    }
    
    CPdfGraphicFrame.prototype.constructor = CPdfGraphicFrame;
    CPdfGraphicFrame.prototype = Object.create(AscFormat.CGraphicFrame.prototype);
    Object.assign(CPdfGraphicFrame.prototype, AscPDF.PdfDrawingPrototype.prototype);

    CPdfGraphicFrame.prototype.IsGraphicFrame = function() {
        return true;
    };
    CPdfGraphicFrame.prototype.copy = function (oPr) {
        let ret = new CPdfGraphicFrame();
        let oDoc = Asc.editor.getPDFDoc();
        if (this.graphicObject) {
            ret.setGraphicObject(this.graphicObject.Copy(ret));
            if (oDoc && isRealObject(oDoc.globalTableStyles)) {
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

        ret._page = this._page; 
        return ret;
    };
    CPdfGraphicFrame.prototype.handleUpdateRot = function() {
        this.SetNeedRecalc(true);
    };
    CPdfGraphicFrame.prototype.Get_PageContentStartPos = function(nPage) {
        return this.GetDocument().Get_PageLimits(nPage);
    };
    CPdfGraphicFrame.prototype.Get_PageContentStartPos2 = function(nPage) {
        return this.Get_PageContentStartPos(nPage);
    };
    CPdfGraphicFrame.prototype.GetDocContent = function() {
        return this.getDocContent();
    };
    CPdfGraphicFrame.prototype.updateCursorType = function (x, y, e) {
		var tx = this.invertTransform.TransformPointX(x, y);
		var ty = this.invertTransform.TransformPointY(x, y);
		this.graphicObject.UpdateCursorType(tx, ty, 0)
	};
    /**
     * Removes char in current position by direction.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CPdfGraphicFrame.prototype.Remove = function(nDirection, isCtrlKey) {
        let oContent = this.GetDocContent();

        if (oContent) {
            oContent.Remove(nDirection, true, false, false, isCtrlKey);
        }
        else {
            this.graphicObject.Remove(nDirection, true, false, false, isCtrlKey);
        }
        
        this.SetNeedRecalc(true);

        if (AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
        }
    };
    CPdfGraphicFrame.prototype.CheckTextOnOpen = function() {
        let oTable = this.graphicObject;
        if (oTable) {
            oTable.SetApplyToAll(true);
            AscFonts.FontPickerByCharacter.getFontsByString(oTable.GetSelectedText(false));
            oTable.SetApplyToAll(false);
        }
    };
    CPdfGraphicFrame.prototype.GetAllFonts = function (fontMap) {
        fontMap = fontMap || {};

		if (this.graphicObject) {
			for (var i = 0; i < this.graphicObject.Content.length; ++i) {
				var row = this.graphicObject.Content[i];
				var cells = row.Content;
				for (var j = 0; j < cells.length; ++j) {
					cells[j].Content.Document_Get_AllFontNames(fontMap);
				}
			}
			delete fontMap["+mj-lt"];
			delete fontMap["+mn-lt"];
			delete fontMap["+mj-ea"];
			delete fontMap["+mn-ea"];
			delete fontMap["+mj-cs"];
			delete fontMap["+mn-cs"];
		}

        return fontMap;
	};
    CPdfGraphicFrame.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalcInfo.recalculateTransform = true;
        this.recalcInfo.recalculateSizes = true;
        
        this.recalculate();
        this.updateTransformMatrix();
        this.SetNeedRecalc(false);
    };
    CPdfGraphicFrame.prototype.MoveCursorToCell = function(bNext) {
        this.graphicObject.MoveCursorToCell(bNext);
    };
    CPdfGraphicFrame.prototype.checkExtentsByDocContent = function() {
        this.Recalculate();
        let oXfrm = this.getXfrm();
        if (Math.abs(oXfrm.extY - this.extY) > 0.001) {
            let nRot = this.GetRot();
            let oldExtX = oXfrm.extX;
            let oldExtY = oXfrm.extY;
            let oldOffX = oXfrm.offX;
            let oldOffY = oXfrm.offY;

            let deltaX = -(oldExtX - this.extX) / 2;
            let deltaY = -(oldExtY - this.extY) / 2;

            let _sin = Math.sin(nRot);
            let _cos = Math.cos(nRot);

            let newOffX = oldOffX + (deltaX*_cos - deltaY*_sin) - deltaX;
            let newOffY = oldOffY + (deltaX*_sin + deltaY*_cos) - deltaY;

            oXfrm.setOffX(newOffX);
            oXfrm.setOffY(newOffY);

            oXfrm.setExtX(this.extX);
            oXfrm.setExtY(this.extY);
        }
    };
    CPdfGraphicFrame.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
        if (bRecalc == false) {
            this._needRecalc = false;
        }
        else {
            // this.GetDocument().SetNeedUpdateTarget(true);
            this._needRecalc = true;
            this.recalcInfo.recalculateTable = true;
           
            if (bSkipAddToRedraw != true)
                this.AddToRedraw();
        }
    };
    CPdfGraphicFrame.prototype.onMouseDown = function(x, y, e) {
        let oViewer             = Asc.editor.getDocumentRenderer();
        let oDoc                = this.GetDocument();
        let oDrawingObjects     = oDoc.Viewer.DrawingObjects;
        this.selectStartPage    = this.GetPage();

        let pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
    };
    
    CPdfGraphicFrame.prototype.onMouseUp = function(x, y, e) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc    = this.GetDocument();
        let oDrDoc  = oDoc.GetDrawingDocument();

        this.selectStartPage    = this.GetPage();
        let oContent            = this.GetDocContent();

        this.graphicObject.Selection.Start = false;

        if (oDrDoc.m_sLockedCursorType.indexOf("resize") != -1) {
            this.SetNeedRecalc(true);
        }

        if (!oContent) {
            return;
        }

        if (global_mouseEvent.ClickCount == 2) {
            oContent.SelectAll();
            if (oContent.IsSelectionEmpty() == false)
                oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
            else
                oContent.RemoveSelection();
        }
                
        if (oContent.IsSelectionEmpty())
            oContent.RemoveSelection();
    };
    
    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfGraphicFrame.prototype.Get_Styles = function (level) {
		if (AscFormat.isRealNumber(level)) {
			if (!this.compiledStyles[level]) {
				AscFormat.CShape.prototype.recalculateTextStyles.call(this, level);
			}
			return this.compiledStyles[level];
		}
        else {
			return Asc.editor.getPDFDoc().globalTableStyles;
		}
	};
    CPdfGraphicFrame.prototype.canRotate = function() {
        return true;
    };
    CPdfGraphicFrame.prototype.getRotateAngle = function (x, y) {
        var transform = this.getTransformMatrix();
        var rotate_distance = this.convertPixToMM(AscCommon.TRACK_DISTANCE_ROTATE);
        var hc = this.extX * 0.5;
        var vc = this.extY * 0.5;
        var xc_t = transform.TransformPointX(hc, vc);
        var yc_t = transform.TransformPointY(hc, vc);
        var rot_x_t = transform.TransformPointX(hc, -rotate_distance);
        var rot_y_t = transform.TransformPointY(hc, -rotate_distance);

        var invert_transform = this.getInvertTransform();
        if (!invert_transform) {
            return 0.0;
        }
        var rel_x = invert_transform.TransformPointX(x, y);

        var v1_x, v1_y, v2_x, v2_y;
        v1_x = x - xc_t;
        v1_y = y - yc_t;

        v2_x = rot_x_t - xc_t;
        v2_y = rot_y_t - yc_t;

        var flip_h = this.getFullFlipH();
        var flip_v = this.getFullFlipV();
        var same_flip = flip_h && flip_v || !flip_h && !flip_v;
        var angle = rel_x > this.extX * 0.5 ? Math.atan2(Math.abs(v1_x * v2_y - v1_y * v2_x), v1_x * v2_x + v1_y * v2_y) : -Math.atan2(Math.abs(v1_x * v2_y - v1_y * v2_x), v1_x * v2_x + v1_y * v2_y);
        return same_flip ? angle : -angle;
    };
    CPdfGraphicFrame.prototype.selectionSetStart = function (e, x, y) {
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
				this.graphicObject.Selection_SetStart(tx, ty, this.GetPage(), e);
			} else {
				if (!this.graphicObject.IsSelectionUse()) {
					this.graphicObject.StartSelectionFromCurPos();
				}
				this.graphicObject.Selection_SetEnd(tx, ty, this.GetPage(), e);
			}
			this.graphicObject.RecalculateCurPos();
		}
	};
    CPdfGraphicFrame.prototype.draw = function (graphics) {
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
            graphics.SetIntegerGrid(true);
            if (this.graphicObject.IsTable()) {
                let oTable  = this.graphicObject;
                let nRows   = oTable.GetRowsCount();

                for (let nRow = 0; nRow < nRows; nRow++) {
                    let oRow = oTable.GetRow(nRow);
                    let nCells = oRow.GetCellsCount();

                    for (let nCell = 0; nCell < nCells; nCell++) {
                        let oContent = oRow.GetCell(nCell).GetContent();
                        oContent.Set_StartPage(0);
                    }
                }
            }

            this.graphicObject.Draw(0, graphics);
            
            this.drawLocks(this.transform, graphics);
            graphics.RestoreGrState();
        }
        graphics.SetIntegerGrid(true);
        graphics.reset();
    };
    CPdfGraphicFrame.prototype.select = function (drawingObjectsController, pageIndex) {
        if (!AscFormat.canSelectDrawing(this)) {
            return;
        }
        this.selected = true;
        this.selectStartPage = pageIndex;
        let content = this.getDocContent && this.getDocContent();
        if (content)
            content.Set_StartPage(0);
        let selected_objects;
        if (!AscCommon.isRealObject(this.group))
            selected_objects = drawingObjectsController ? drawingObjectsController.selectedObjects : [];
        else
            selected_objects = this.group.getMainGroup().selectedObjects;
        for (var i = 0; i < selected_objects.length; ++i) {
            if (selected_objects[i] === this)
                break;
        }
        if (i === selected_objects.length)
            selected_objects.push(this);

        if (drawingObjectsController) {
            drawingObjectsController.onChangeDrawingsSelection();
        }
    };
    CPdfGraphicFrame.prototype.updateSelectionState = function () {
        let oDoc    = this.GetDocument();
        let oDrDoc  = oDoc.GetDrawingDocument();

		if (isRealObject(this.graphicObject)) {
			let graphicObject = this.graphicObject;
			if (true === graphicObject.IsSelectionUse() && !graphicObject.IsSelectionEmpty()) {
				oDrDoc.UpdateTargetTransform(this.transform);
				oDrDoc.TargetEnd();
				oDrDoc.Overlay && graphicObject.DrawSelectionOnPage(0);
			} else {
				graphicObject.RecalculateCurPos();
				oDrDoc.UpdateTargetTransform(this.transform);
				oDrDoc.TargetStart(true);
			}
		} else {
			oDrDoc.UpdateTargetTransform(null);
			oDrDoc.TargetEnd();
		}
	};
    CPdfGraphicFrame.prototype.Set_CurrentElement = function () {
        let oDoc        = this.GetDocument();
        let oController = oDoc.GetController();

		oController.resetSelection(true);
        if (this.group) {
            var main_group = this.group.getMainGroup();
            oController.selectObject(main_group, 0);
            main_group.selectObject(this, 0);
            main_group.selection.textSelection = this;
        }
        else {
            oController.selectObject(this, 0);
            oController.selection.textSelection = this;
        }
	};
    CPdfGraphicFrame.prototype.Get_PageFields = function (nPage) {
        return this.Get_PageLimits(nPage);
    };

    window["AscPDF"].CPdfGraphicFrame = CPdfGraphicFrame;
})();

