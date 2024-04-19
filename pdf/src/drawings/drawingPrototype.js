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
    function CPdfDrawingPrototype()
    {
        this._page          = undefined;
        this._apIdx         = undefined; // индекс объекта в файле

        this._isFromScan = false;   // флаг, что был прочитан из скана страницы 

        this._doc                   = undefined;
        this._needRecalc            = true;
    }
    
    CPdfDrawingPrototype.prototype.constructor = CPdfDrawingPrototype;
    
    CPdfDrawingPrototype.prototype.IsAnnot = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsForm = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsTextShape = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsImage = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsChart = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsDrawing = function() {
        return true;
    };
    CPdfDrawingPrototype.prototype.IsSmartArt = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsGraphicFrame = function() {
        return false;
    };

    CPdfDrawingPrototype.prototype.IsUseInDocument = function() {
        if (this.GetDocument().drawings.indexOf(this) == -1)
            return false;

        return true;
    };

    CPdfDrawingPrototype.prototype.SetFromScan = function(bFromScan) {
        this._isFromScan = bFromScan;
    };
    CPdfDrawingPrototype.prototype.IsFromScan = function() {
        return this._isFromScan;
    };
    CPdfDrawingPrototype.prototype.SetDocument = function(oDoc) {
        this._doc = oDoc;
    };
    CPdfDrawingPrototype.prototype.GetDocument = function() {
        if (this.group)
            return this.group.getLogicDocument();

        return this._doc;
    };
    CPdfDrawingPrototype.prototype.SetPage = function(nPage) {
        let nCurPage = this.GetPage();
        if (nPage == nCurPage)
            return;

        // initial set
        if (nCurPage == undefined) {
            this._page = nPage;
            return;
        }

        let oViewer = editor.getDocumentRenderer();
        let oDoc    = this.GetDocument();
        
        let nCurIdxOnPage = oViewer.pagesInfo.pages[nCurPage] && oViewer.pagesInfo.pages[nCurPage].drawings ? oViewer.pagesInfo.pages[nCurPage].drawings.indexOf(this) : -1;
        if (oViewer.pagesInfo.pages[nPage]) {
            if (oDoc.drawings.indexOf(this) != -1) {
                if (nCurIdxOnPage != -1) {
                    oViewer.pagesInfo.pages[nCurPage].drawings.splice(nCurIdxOnPage, 1);
                    oDoc.History.Add(new CChangesPDFDrawingPage(this, nCurPage, nPage));
                }
    
                if (this.IsUseInDocument() && oViewer.pagesInfo.pages[nPage].drawings.indexOf(this) == -1)
                    oViewer.pagesInfo.pages[nPage].drawings.push(this);

                // добавляем в перерисовку исходную страницу
                this.AddToRedraw();
            }

            this._page = nPage;
            this.selectStartPage = nPage;
            this.AddToRedraw();
        }
    };
    CPdfDrawingPrototype.prototype.GetPage = function() {
        if (this.group)
            return this.group.Get_AbsolutePage();
        
        return this._page;
    };
    CPdfDrawingPrototype.prototype.AddToRedraw = function() {
        let oViewer = Asc.editor.getDocumentRenderer();
        let nPage   = this.GetPage();
        
        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[nPage]) {
                oViewer.pagesInfo.pages[nPage].needRedrawTextShapes = true;
                oViewer.thumbnails && oViewer.thumbnails._repaintPage(nPage);
            }
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    
    CPdfDrawingPrototype.prototype.SetRot = function(dAngle) {
        let oDoc = this.GetDocument();

        oDoc.History.Add(new CChangesPDFDrawingRot(this, this.GetRot(), dAngle));

        this.changeRot(dAngle);
        this.SetNeedRecalc(true);
    };
    CPdfDrawingPrototype.prototype.GetRot = function() {
        return this.rot;
    };
    CPdfDrawingPrototype.prototype.Recalculate = function() {
    };
    CPdfDrawingPrototype.prototype.IsNeedRecalc = function() {
       return this._needRecalc;
    };
    CPdfDrawingPrototype.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
        if (bRecalc == false) {
            this._needRecalc = false;
        }
        else {
            let oDoc = Asc.editor.getPDFDoc();
            oDoc.ClearSearch();

            oDoc.SetNeedUpdateTarget(true);
            this._needRecalc = true;
            if (bSkipAddToRedraw != true)
                this.AddToRedraw();
        }
    };
    CPdfDrawingPrototype.prototype.Draw = function(oGraphicsWord) {
        this.Recalculate();
        this.draw(oGraphicsWord);
    };
    CPdfDrawingPrototype.prototype.onMouseDown = function(x, y, e) {};
    CPdfDrawingPrototype.prototype.onMouseUp = function(x, y, e) {};
    CPdfDrawingPrototype.prototype.GetDocContent = function() {
        return null;
    };
    CPdfDrawingPrototype.prototype.SetInTextBox = function(bIn) {
        this.isInTextBox = bIn;
    };
    CPdfDrawingPrototype.prototype.IsInTextBox = function() {
        return this.isInTextBox;
    };
    
    /////////////////////////////
    /// saving
    ////////////////////////////

    CPdfDrawingPrototype.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfDrawingPrototype.prototype.Get_AbsolutePage = function() {
        return this.GetPage();
    };
    CPdfDrawingPrototype.prototype.getLogicDocument = function() {
        return this.GetDocument();
    };
    CPdfDrawingPrototype.prototype.IsThisElementCurrent = function() {
        return this.selected;
    };
    CPdfDrawingPrototype.prototype.getDrawingDocument = function() {
        return Asc.editor.getPDFDoc().GetDrawingDocument();
    };

    /////////////////////////////
    /// saving
    ////////////////////////////

    CPdfDrawingPrototype.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    window["AscPDF"].PdfDrawingPrototype = CPdfDrawingPrototype;
})();

