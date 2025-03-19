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
    function CPdfDrawingPrototype() {
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
    CPdfDrawingPrototype.prototype.IsShape = function() {
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
    CPdfDrawingPrototype.prototype.IsPdfObject = function() {
        return true;
    };
    CPdfDrawingPrototype.prototype.OnContentChange = function() {
        return this.SetNeedRecalc(true);
    };
    CPdfDrawingPrototype.prototype.IsPdfDrawing = function() {
        return true;
    };
    CPdfDrawingPrototype.prototype.IsSmartArt = function() {
        return false;
    };
    CPdfDrawingPrototype.prototype.IsGraphicFrame = function() {
        return false;
    };
	CPdfDrawingPrototype.prototype.IsUseInDocument = function() {
		if (this.group && this.group.IsUseInDocument)
			return this.group.IsUseInDocument();
		
        let oPage = this.GetParentPage();
        if (oPage && oPage.drawings.includes(this)) {
            return true;
        }

		return false;
	};
    CPdfDrawingPrototype.prototype.OnBlur = function() {
        AscCommon.History.ForbidUnionPoint();
    };
    CPdfDrawingPrototype.prototype.recalculateContent = function() {
        let parentPrototype = Object.getPrototypeOf(Object.getPrototypeOf(this));
        let oRecalcData = null;
        // Вызов родительского метода
        if (parentPrototype && parentPrototype.recalculateContent && parentPrototype.recalculateContent != CPdfDrawingPrototype.prototype.recalculateContent) {
            oRecalcData = parentPrototype.recalculateContent.call(this);
        }

        return oRecalcData;
    };
    CPdfDrawingPrototype.prototype.SetParentPage = function(oParent) {
        this.parent = oParent;
    };
    CPdfDrawingPrototype.prototype.setParent2 = function() {};
    CPdfDrawingPrototype.prototype.GetParentPage = function() {
        return this.parent;
    };
    
    CPdfDrawingPrototype.prototype.GetSelectionQuads = function() {
        let oDoc        = this.GetDocument();
        let oViewer     = oDoc.Viewer;
        let oFile       = oViewer.file;
        let oDrDoc      = oDoc.GetDrawingDocument();
        let oContent    = this.GetDocContent();
        let aInfo       = [];
        let nPage       = this.GetPage();

        if (!oContent || !oContent.IsSelectionUse()) {
            return aInfo;
        }

        let nStart = oContent.Selection.StartPos;
        let nEnd   = oContent.Selection.EndPos;
        if (nStart > nEnd) [nStart, nEnd] = [nEnd, nStart];

        let oInfo = {
            page: nPage,
            quads: []
        }

        for (let i = nStart; i <= nEnd; i++) {
            let oPara = oContent.GetElement(i);

            let nStartInPara = oPara.Selection.StartPos;
            let nEndInPara   = oPara.Selection.EndPos;
            if (nStartInPara > nEndInPara) [nStartInPara, nEndInPara] = [nEndInPara, nStartInPara];

            let nStartLine = oPara.Pages[0].StartLine;
			let nEndLine   = oPara.Pages[0].EndLine;

            if (nStartInPara > oPara.Lines[nEndLine].Get_EndPos() || nEndInPara < oPara.Lines[nStartLine].Get_StartPos()) {
				return aInfo;
            }
			else {
				nStartInPara = Math.max(nStartInPara, oPara.Lines[nStartLine].Get_StartPos());
				nEndInPara   = Math.min(nEndInPara, ( nEndLine !== oPara.Lines.length - 1 ? oPara.Lines[nEndLine].Get_EndPos() : oPara.Content.length - 1 ));
			}

            let oDrawSelectionState = new AscWord.ParagraphDrawSelectionState(oPara);
			oDrawSelectionState.resetPage(0);

            for (let iLine = nStartLine; iLine <= nEndLine; ++iLine) {
                let oLine = oPara.Lines[iLine];
                oDrawSelectionState.resetLine(iLine);
                
                for (let iRange = 0, rangeCount = oLine.Ranges.length; iRange < rangeCount; ++iRange) {
                    let oRange = oLine.Ranges[iRange];
                    
                    let nRangeStart = oRange.StartPos;
                    let nRangeEnd   = oRange.EndPos;
                    
                    // Если пересечение пустое с селектом, тогда пропускаем данный отрезок
                    if (nStartInPara > nRangeEnd || nEndInPara < nRangeStart)
                        continue;
                    
                    oDrawSelectionState.beginRange(iRange);
                    
                    for (let pos = nRangeStart; pos <= nRangeEnd; ++pos) {
                        oPara.Content[pos].drawSelectionInRange(iLine, iRange, oDrawSelectionState);
                    }
                    
                    oDrawSelectionState.endRange();
                    
                    let aAnchored = oDrawSelectionState.getAnchoredObjects();
                    for (let index = 0; index < aAnchored.length; ++index) {
                        aAnchored[index].Draw_Selection();
                    }
                }
            }

            let aSelectionRanges = oDrawSelectionState.getSelectionRanges();
            for (let iSel = 0; iSel < aSelectionRanges.length; ++iSel) {
                let x = aSelectionRanges[iSel].x;
                let w = aSelectionRanges[iSel].w;
                let y = aSelectionRanges[iSel].y;
                let h = aSelectionRanges[iSel].h;
                
                if (oPara.CalculatedFrame) {
                    let Frame_X_min = oPara.CalculatedFrame.L2;
                    let Frame_Y_min = oPara.CalculatedFrame.T2;
                    let Frame_X_max = oPara.CalculatedFrame.L2 + oPara.CalculatedFrame.W2;
                    let Frame_Y_max = oPara.CalculatedFrame.T2 + oPara.CalculatedFrame.H2;
                    
                    x = Math.min(Math.max(Frame_X_min, x), Frame_X_max);
                    y = Math.min(Math.max(Frame_Y_min, y), Frame_Y_max);
                    w = Math.min(w, Frame_X_max - x);
                    h = Math.min(h, Frame_Y_max - y);
                }
                
                let isTextMatrixUse = ((null != oDrDoc.TextMatrix) && !global_MatrixTransformer.IsIdentity(oDrDoc.TextMatrix));
                if (isTextMatrixUse) {
                    let oPt1 = oDrDoc.TextMatrix.TransformPoint(x, y);            // левый верхний
                    let oPt2 = oDrDoc.TextMatrix.TransformPoint(x + w, y);        // правый верхний
                    let oPt3 = oDrDoc.TextMatrix.TransformPoint(x + w, y + h);    // правый нижний
                    let oPt4 = oDrDoc.TextMatrix.TransformPoint(x, y + h);        // левый нижний

                    let nKoeff = oViewer.getDrawingPageScale(nPage) * g_dKoef_pix_to_mm;

                    oInfo.quads.push([oPt1.x / nKoeff, oPt1.y / nKoeff, oPt2.x / nKoeff, oPt2.y / nKoeff, oPt4.x / nKoeff, oPt4.y / nKoeff, oPt3.x / nKoeff, oPt3.y / nKoeff]);
                }
            }
        }

        if (oInfo.quads.length != 0) {
            aInfo.push(oInfo);
        }

        return aInfo;
    };
    CPdfDrawingPrototype.prototype.GetOrigRect = function() {
        let oXfrm = this.getXfrm();

        return [oXfrm.offX * g_dKoef_mm_to_pt, oXfrm.offY * g_dKoef_mm_to_pt, (oXfrm.offX + this.extX) * g_dKoef_mm_to_pt, (oXfrm.offY + this.extY) * g_dKoef_mm_to_pt];
    };
    CPdfDrawingPrototype.prototype.SetFromScan = function(bFromScan) {
        this._isFromScan = bFromScan;
    };
    CPdfDrawingPrototype.prototype.IsFromScan = function() {
        return this._isFromScan;
    };
    CPdfDrawingPrototype.prototype.SetDocument = function(oDoc) {
        if (this._doc == oDoc) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFDocumentSetDocument(this, this._doc, oDoc));
        this._doc = oDoc;
    };
    CPdfDrawingPrototype.prototype.OnContentChange = function() {
        if (this.group) {
            this.group.SetNeedRecalc(true);
        }
        else {
            this.SetNeedRecalc(true);
        }
    };
    CPdfDrawingPrototype.prototype.OnTextPrChange = function() {
        if (this.group) {
            this.group.SetNeedRecalc(true);
        }
        else {
            this.SetNeedRecalc(true);
        }
    };
    CPdfDrawingPrototype.prototype.GetDocument = function() {
        if (this.group)
            return this.group.getLogicDocument();

        return this._doc;
    };
    CPdfDrawingPrototype.prototype.SetPage = function(nPage) {
        if (this.GetPage() == nPage) {
            return;
        }
        
        let oDoc        = this.GetDocument();
        let oNewPage    = oDoc.GetPageInfo(nPage);

        if (oNewPage) {
            oDoc.RemoveDrawing(this.GetId(), true);
            oDoc.AddDrawing(this, nPage);
            this.selectStartPage = nPage;
        }
    };
    CPdfDrawingPrototype.prototype.GetPage = function() {
        if (this.group)
            return this.group.GetPage();
        
        let oParentPage = this.GetParentPage();
        if (!oParentPage || !(oParentPage instanceof AscPDF.CPageInfo)) {
            return -1;
        }

        return oParentPage.GetIndex();
    };
    
    CPdfDrawingPrototype.prototype.AddToRedraw = function() {
        let oViewer = Asc.editor.getDocumentRenderer();
        let nPage   = this.GetPage();
        
        if (false == this.IsUseInDocument()) {
            return;
        }
        
        if (this.group && this.group.IsAnnot()) {
            this.group.AddToRedraw();
            return;
        }

        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[nPage]) {
                oViewer.pagesInfo.pages[nPage].needRedrawDrawings = true;
                oViewer.thumbnails && oViewer.thumbnails._repaintPage(nPage);
            }
        }

        oViewer.paint(setRedrawPageOnRepaint);
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
            if (this.group) {
                if (!this.group.IsPdfObject || !this.group.IsPdfObject()) {
                    return;
                }

                this.group.SetNeedRecalc(bRecalc, bSkipAddToRedraw);
                return;
            }

            let oDoc = Asc.editor.getPDFDoc();
            oDoc.ClearSearch();

            oDoc.SetNeedUpdateTarget(true);
            this._needRecalc = true;
            if (bSkipAddToRedraw != true)
                this.AddToRedraw();

            if (this.group) {
                this.group.SetNeedRecalc(true);
            }
        }
    };
    CPdfDrawingPrototype.prototype.addToRecalculate = function() {
        this.SetNeedRecalc(true);
    };
    CPdfDrawingPrototype.prototype.GetAllFonts = function(fontMap) {
        fontMap = fontMap || {};
        return fontMap;
    };
    CPdfDrawingPrototype.prototype.CheckTextOnOpen = function() {};
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
        let oDoc = editor.getPDFDoc();
        let oController = oDoc.GetController();

        if (oDoc.GetActiveObject() == this && this == oController.getTargetTextObject()) {
            return !!this.GetDocContent();
        }

        return false;
    };
	CPdfDrawingPrototype.prototype.Remove = function(direction, isWord) {
		let doc = this.GetDocument();
		let content = this.GetDocContent();
		
		if (!doc || !content)
			return;
		
		content.Remove(direction, true, false, false, isWord);
		this.SetNeedRecalc(true);
		content.RecalculateCurPos();
	};
	CPdfDrawingPrototype.prototype.EnterText = function(value) {
		let doc = this.GetDocument();
		let content = this.GetDocContent();
		if (!doc || !content)
			return false;
		
		let result = content.EnterText(value);
		content.RecalculateCurPos();
        
        this.checkExtentsByDocContent && this.checkExtentsByDocContent();
		return result;
	};
	CPdfDrawingPrototype.prototype.CorrectEnterText = function(oldValue, newValue) {
		let doc = this.GetDocument();
		let content = this.GetDocContent();
		if (!doc || !content)
			return false;
		
		let result = content.CorrectEnterText(oldValue, newValue, function(run, inRunPos, codePoint){return true;});
		content.RecalculateCurPos();
		return result;
	};
	CPdfDrawingPrototype.prototype.canBeginCompositeInput = function() {
		return true;
	};
	CPdfDrawingPrototype.prototype.beforeCompositeInput = function() {
		let docContent = this.GetDocContent();
		if (docContent.IsSelectionUse()) {
			docContent.Remove(1, true, false, true);
			docContent.RemoveSelection();
		}
	};
    
    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfDrawingPrototype.prototype.Get_AbsolutePage = function(nCurPage) {
        return this.GetPage() != undefined ? this.GetPage() : nCurPage;
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

