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
                
        this._page          = undefined;
        this._apIdx         = undefined; // индекс объекта в файле
        this._rect          = [];       // scaled rect
        this._richContents  = [];

        this._isFromScan = false; // флаг, что был прочитан из скана текста 

        this._doc                   = undefined;
        this._needRecalc            = true;
        this._wasChanged            = false; // была ли изменена
        this._bDrawFromStream       = false; // нужно ли рисовать из стрима
        this._hasOriginView         = false; // имеет ли внешний вид из файла
    }
    
    CPdfGraphicFrame.prototype.constructor = CPdfGraphicFrame;
    CPdfGraphicFrame.prototype = Object.create(AscFormat.CGraphicFrame.prototype);

    CPdfGraphicFrame.prototype.SetFromScan = function(bFromScan) {
        this._isFromScan = bFromScan;

        // выставляем пунктирный бордер если нет заливки и  
        if (this.spPr.Fill.isNoFill() && this.spPr.ln.Fill.isNoFill()) {
            this.spPr.ln.setPrstDash(Asc.c_oDashType.sysDot);
            this.spPr.ln.setW(25.4 / 72.0 * 36000);
            this.spPr.ln.setFill(AscFormat.CreateSolidFillRGBA(0, 0, 0, 255));
        }
    };
    CPdfGraphicFrame.prototype.IsFromScan = function() {
        return this._isFromScan;
    };
    CPdfGraphicFrame.prototype.SetDocument = function(oDoc) {
        this._doc = oDoc;
    };
    CPdfGraphicFrame.prototype.SetPage = function(nPage) {
        this._page = nPage;
    };
    CPdfGraphicFrame.prototype.IsNeedDrawFromStream = function() {
       return false; 
    };
    CPdfGraphicFrame.prototype.IsAnnot = function() {
       return false;
    };
    CPdfGraphicFrame.prototype.IsForm = function() {
       return false;
    };
    CPdfGraphicFrame.prototype.IsTextShape = function() {
       return false;
    };
    CPdfGraphicFrame.prototype.IsImage = function() {
        return true;
    };
    CPdfGraphicFrame.prototype.IsChart = function() {
        return false;
    };
    CPdfGraphicFrame.prototype.IsDrawing = function() {
        return true;
    };

    CPdfGraphicFrame.prototype.ShowBorder = function(bShow) {
        let oLine = this.pen;

        if (bShow) {
            oLine.setFill(AscFormat.CreateSolidFillRGBA(0, 0, 0, 255));
        }
        else {
            oLine.setFill(AscFormat.CreateNoFillUniFill());
        }

        this.AddToRedraw();
    };
    CPdfGraphicFrame.prototype.SetApIdx = function(nIdx) {
        this.GetDocument().UpdateApIdx(nIdx);
        this._apIdx = nIdx;
    };
    CPdfGraphicFrame.prototype.GetApIdx = function() {
        return this._apIdx;
    };
    CPdfGraphicFrame.prototype.GetDocument = function() {
        if (this.group)
            return this.group.GetDocument();

        return this._doc;
    };
    CPdfGraphicFrame.prototype.GetPage = function() {
        if (this.group)
            return this.group.GetPage();
        
        return this._page;
    };
    CPdfGraphicFrame.prototype.SetInTextBox = function(bIn) {
        this.isInTextBox = bIn;
    };
    CPdfGraphicFrame.prototype.IsInTextBox = function() {
        return this.isInTextBox;
    };
    CPdfGraphicFrame.prototype.GetDocContent = function() {
        return this.getDocContent();
    };
    CPdfGraphicFrame.prototype.EnterText = function(aChars) {
        let oDoc        = this.GetDocument();
        let oContent    = this.GetDocContent();
        let oParagraph  = oContent.GetCurrentParagraph();

        oDoc.CreateNewHistoryPoint({objects: [this]});

        // удаляем текст в селекте
        if (oContent.IsSelectionUse())
            oContent.Remove(-1);

        for (let index = 0; index < aChars.length; ++index) {
            let oRun = AscPDF.codePointToRunElement(aChars[index]);
            if (oRun)
                oParagraph.AddToParagraph(oRun, true);
        }

        this.SetNeedRecalc(true);
        oContent.RecalculateCurPos();

        return true;
    };
    /**
     * Removes char in current position by direction.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CPdfGraphicFrame.prototype.Remove = function(nDirection, isCtrlKey) {
        let oDoc = this.GetDocument();
        oDoc.CreateNewHistoryPoint({objects: [this]});

        let oContent = this.GetDocContent();
        oContent.Remove(nDirection, true, false, false, isCtrlKey);
        oContent.RecalculateCurPos();
        this.SetNeedRecalc(true);

        if (AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
        }
        else {
            this.SetNeedRecalc(true);
        }
    };
    CPdfGraphicFrame.prototype.AddToRedraw = function() {
        let oViewer = Asc.editor.getDocumentRenderer();
        let _t      = this;
        
        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[_t.GetPage()])
                oViewer.pagesInfo.pages[_t.GetPage()].needRedrawTextShapes = true;
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    CPdfGraphicFrame.prototype.GetRect = function() {
        return this._rect;
    };
    
    CPdfGraphicFrame.prototype.SetRot = function(dAngle) {
        let oDoc = this.GetDocument();

        oDoc.History.Add(new CChangesPDFTxShapeRot(this, this.GetRot(), dAngle));

        this.changeRot(dAngle);
        this.SetNeedRecalc(true);
    };
    CPdfGraphicFrame.prototype.GetRot = function() {
        return this.rot;
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
    CPdfGraphicFrame.prototype.IsNeedRecalc = function() {
       return this._needRecalc;
    };
    CPdfGraphicFrame.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
       if (bRecalc == false) {
           this._needRecalc = false;
       }
       else {
           this._needRecalc = true;
           this.recalcInfo.recalculateTable = true;
           
           if (bSkipAddToRedraw != true)
               this.AddToRedraw();
       }
    };
    CPdfGraphicFrame.prototype.Draw = function(oGraphicsWord) {
        this.Recalculate();
        this.draw(oGraphicsWord);
    };
    CPdfGraphicFrame.prototype.SelectionSetStart = function(X, Y, e) {
        this.selectStartPage = this.GetPage();

        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        let oTransform  = this.getInvertTransform();
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        this.graphicObject.Selection_SetStart(xContent, yContent, 0, e);
        oContent.RecalculateCurPos();
    };
    
    CPdfGraphicFrame.prototype.SelectionSetEnd = function(X, Y, e) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        let oTransform  = this.getInvertTransform();
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        this.graphicObject.Selection_SetEnd(xContent, yContent, 0, e);
    };
    CPdfGraphicFrame.prototype.onMouseDown = function(x, y, e) {
        let oDoc                = this.GetDocument();
        let oDrawingObjects     = oDoc.Viewer.DrawingObjects;
        let oDrDoc              = oDoc.GetDrawingDocument();
        this.selectStartPage    = this.GetPage();

        let oPos    = oDrDoc.ConvertCoordsFromCursor2(x, y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        if (this.hitInBoundingRect(X, Y) || this.hitToHandles(X, Y) != -1 || this.hitInPath(X, Y)) {
            this.SetInTextBox(false);
        }
        else {
            this.SetInTextBox(true);
            // oDoc.SelectionSetStart(x, y, e);
        }

        oDrawingObjects.OnMouseDown(e, X, Y, this.selectStartPage);
    };
    CPdfGraphicFrame.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorLeft(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfGraphicFrame.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorRight(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfGraphicFrame.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorDown(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfGraphicFrame.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorUp(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    
    /**
     * Exit from this drawing.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CPdfGraphicFrame.prototype.Blur = function() {};

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
    
    /////////////////////////////
    /// saving
    ////////////////////////////

    CPdfGraphicFrame.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfGraphicFrame.prototype.Get_AbsolutePage = function() {
        return this.GetPage();
    };
    CPdfGraphicFrame.prototype.getLogicDocument = function() {
        return this.GetDocument();
    };
    CPdfGraphicFrame.prototype.IsThisElementCurrent = function() {
        return true;
    };
    CPdfGraphicFrame.prototype.getDrawingDocument = function() {
        return Asc.editor.getPDFDoc().GetDrawingDocument();
    };
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
    CPdfGraphicFrame.prototype.Get_PageFields = function (nPage) {
        return this.Get_PageLimits(nPage);
    };

    window["AscPDF"].CPdfGraphicFrame = CPdfGraphicFrame;
})();

