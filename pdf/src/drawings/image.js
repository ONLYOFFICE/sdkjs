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
    function CPdfImage()
    {
        AscFormat.CImageShape.call(this);
                
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
    
    CPdfImage.prototype.constructor = CPdfImage;
    CPdfImage.prototype = Object.create(AscFormat.CImageShape.prototype);

    CPdfImage.prototype.SetFromScan = function(bFromScan) {
        this._isFromScan = bFromScan;

        // выставляем пунктирный бордер если нет заливки и  
        if (this.spPr.Fill.isNoFill() && this.spPr.ln.Fill.isNoFill()) {
            this.spPr.ln.setPrstDash(Asc.c_oDashType.sysDot);
            this.spPr.ln.setW(25.4 / 72.0 * 36000);
            this.spPr.ln.setFill(AscFormat.CreateSolidFillRGBA(0, 0, 0, 255));
        }
    };
    CPdfImage.prototype.IsFromScan = function() {
        return this._isFromScan;
    };
    CPdfImage.prototype.SetDocument = function(oDoc) {
        this._doc = oDoc;
    };
    CPdfImage.prototype.SetPage = function(nPage) {
        this._page = nPage;
    };
    CPdfImage.prototype.IsNeedDrawFromStream = function() {
       return false; 
    };
    CPdfImage.prototype.IsAnnot = function() {
       return false;
    };
    CPdfImage.prototype.IsForm = function() {
       return false;
    };
    CPdfImage.prototype.IsTextShape = function() {
       return false;
    };
    CPdfImage.prototype.IsImage = function() {
        return true;
    };
    CPdfImage.prototype.IsChart = function() {
        return false;
    };
    CPdfImage.prototype.IsDrawing = function() {
        return true;
    };

    CPdfImage.prototype.ShowBorder = function(bShow) {
        let oLine = this.pen;

        if (bShow) {
            oLine.setFill(AscFormat.CreateSolidFillRGBA(0, 0, 0, 255));
        }
        else {
            oLine.setFill(AscFormat.CreateNoFillUniFill());
        }

        this.AddToRedraw();
    };
    CPdfImage.prototype.SetApIdx = function(nIdx) {
        this.GetDocument().UpdateApIdx(nIdx);
        this._apIdx = nIdx;
    };
    CPdfImage.prototype.GetApIdx = function() {
        return this._apIdx;
    };
    CPdfImage.prototype.GetDocument = function() {
        if (this.group)
            return this.group.GetDocument();

        return this._doc;
    };
    CPdfImage.prototype.GetPage = function() {
        if (this.group)
            return this.group.GetPage();
        
        return this._page;
    };
    CPdfImage.prototype.IsInTextBox = function() {
        return false;
    };
    CPdfImage.prototype.GetDocContent = function() {
        return null;
    };
    CPdfImage.prototype.AddToRedraw = function() {
        let oViewer = Asc.editor.getDocumentRenderer();
        let _t      = this;
        
        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[_t.GetPage()])
                oViewer.pagesInfo.pages[_t.GetPage()].needRedrawTextShapes = true;
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    CPdfImage.prototype.GetRect = function() {
        return this._rect;
    };
    
    CPdfImage.prototype.SetRect = function(aRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();

        oDoc.History.Add(new CChangesPDFTxShapeRect(this, this.GetRect(), aRect));

        this._rect = aRect;

        this._pagePos = {
            x: aRect[0],
            y: aRect[1],
            w: (aRect[2] - aRect[0]),
            h: (aRect[3] - aRect[1])
        };

        this.spPr.xfrm.extX = this._pagePos.w * g_dKoef_pix_to_mm;
        this.spPr.xfrm.extY = this._pagePos.h * g_dKoef_pix_to_mm;
        this.spPr.xfrm.offX = aRect[0] * g_dKoef_pix_to_mm;
        this.spPr.xfrm.offY = aRect[1] * g_dKoef_pix_to_mm;
        this.updateTransformMatrix();

        this.SetNeedRecalc(true);
    };
    CPdfImage.prototype.SetRot = function(dAngle) {
        let oDoc = this.GetDocument();

        oDoc.History.Add(new CChangesPDFTxShapeRot(this, this.GetRot(), dAngle));

        this.changeRot(dAngle);
        this.SetNeedRecalc(true);
    };
    CPdfImage.prototype.GetRot = function() {
        return this.rot;
    };
    CPdfImage.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalcGeometry();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CPdfImage.prototype.IsNeedRecalc = function() {
       return this._needRecalc;
    };
    CPdfImage.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
       if (bRecalc == false) {
           this._needRecalc = false;
       }
       else {
           this._needRecalc = true;
           if (bSkipAddToRedraw != true)
               this.AddToRedraw();
       }
    };
    CPdfImage.prototype.Draw = function(oGraphicsWord) {
        this.Recalculate();
        this.draw(oGraphicsWord);
    };
    CPdfImage.prototype.onMouseDown = function(x, y, e) {
        let oDoc                = this.GetDocument();
        let oDrawingObjects     = oDoc.Viewer.DrawingObjects;
        let oDrDoc              = oDoc.GetDrawingDocument();
        this.selectStartPage    = this.GetPage();

        let oPos    = oDrDoc.ConvertCoordsFromCursor2(x, y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        oDrawingObjects.OnMouseDown(e, X, Y, this.selectStartPage);
    };
    CPdfImage.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorLeft(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfImage.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorRight(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfImage.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorDown(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfImage.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
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
    CPdfImage.prototype.Blur = function() {};

    CPdfImage.prototype.onMouseUp = function(x, y, e) {};
    
    /////////////////////////////
    /// saving
    ////////////////////////////

    CPdfImage.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfImage.prototype.Get_AbsolutePage = function() {
        return this.GetPage();
    };
    CPdfImage.prototype.getLogicDocument = function() {
        return this.GetDocument();
    };
    CPdfImage.prototype.IsThisElementCurrent = function() {
        return true;
    };

    window["AscPDF"].CPdfImage = CPdfImage;
})();

