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
    function CPdfShape()
    {
        AscFormat.CShape.call(this);
                
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
    
    CPdfShape.prototype.constructor = CPdfShape;
    CPdfShape.prototype = Object.create(AscFormat.CShape.prototype);

    CPdfShape.prototype.SetFromScan = function(bFromScan) {
        this._isFromScan = bFromScan;

        // выставляем пунктирный бордер если нет заливки и  
        if (this.spPr.Fill.isNoFill() && this.spPr.ln.Fill.isNoFill()) {
            this.spPr.ln.setPrstDash(Asc.c_oDashType.sysDot);
            this.spPr.ln.setW(25.4 / 72.0 * 36000);
            this.spPr.ln.setFill(AscFormat.CreateSolidFillRGBA(0, 0, 0, 255));
        }
    };
    CPdfShape.prototype.IsFromScan = function() {
        return this._isFromScan;
    };
    CPdfShape.prototype.SetDocument = function(oDoc) {
        this._doc = oDoc;
    };
    CPdfShape.prototype.SetPage = function(nPage) {
        this._page = nPage;
    };
    CPdfShape.prototype.IsNeedDrawFromStream = function() {
       return false; 
    };
    CPdfShape.prototype.IsAnnot = function() {
       return false;
    };
    CPdfShape.prototype.IsForm = function() {
       return false;
    };
    CPdfShape.prototype.IsTextShape = function() {
        return true;
    };
    CPdfShape.prototype.IsImage = function() {
        return false;
    };
    CPdfShape.prototype.IsChart = function() {
        return false;
    };
    CPdfShape.prototype.IsDrawing = function() {
        return true;
    };
    CPdfShape.prototype.ShowBorder = function(bShow) {
        let oLine = this.pen;

        if (bShow) {
            oLine.setFill(AscFormat.CreateSolidFillRGBA(0, 0, 0, 255));
        }
        else {
            oLine.setFill(AscFormat.CreateNoFillUniFill());
        }

        this.AddToRedraw();
    };
    CPdfShape.prototype.SetApIdx = function(nIdx) {
        this.GetDocument().UpdateApIdx(nIdx);
        this._apIdx = nIdx;
    };
    CPdfShape.prototype.GetApIdx = function() {
        return this._apIdx;
    };
    CPdfShape.prototype.GetDocument = function() {
        if (this.group)
            return this.group.GetDocument();

        return this._doc;
    };
    CPdfShape.prototype.GetPage = function() {
        if (this.group)
            return this.group.GetPage();
        
        return this._page;
    };
    CPdfShape.prototype.AddToRedraw = function() {
        let oViewer = Asc.editor.getDocumentRenderer();
        let _t      = this;
        
        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[_t.GetPage()])
                oViewer.pagesInfo.pages[_t.GetPage()].needRedrawTextShapes = true;
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    CPdfShape.prototype.GetRect = function() {
        return this._rect;
    };
    
    CPdfShape.prototype.SetRect = function(aRect) {
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
    CPdfShape.prototype.SetRot = function(dAngle) {
        let oDoc = this.GetDocument();

        oDoc.History.Add(new CChangesPDFTxShapeRot(this, this.GetRot(), dAngle));

        this.changeRot(dAngle);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.GetRot = function() {
        return this.rot;
    };
    CPdfShape.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalcGeometry();
        this.recalculateContent();
        this.checkExtentsByDocContent(true, true);
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CPdfShape.prototype.IsNeedRecalc = function() {
       return this._needRecalc;
    };
    CPdfShape.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
       if (bRecalc == false) {
           this._needRecalc = false;
       }
       else {
           this._needRecalc = true;
           if (bSkipAddToRedraw != true)
               this.AddToRedraw();
       }
    };
    CPdfShape.prototype.Draw = function(oGraphicsWord) {
        this.Recalculate();
        this.draw(oGraphicsWord);
    };
    CPdfShape.prototype.onMouseDown = function(x, y, e) {
        let oDoc                = this.GetDocument();
        let oDrawingObjects     = oDoc.Viewer.DrawingObjects;
        let oDrDoc              = oDoc.GetDrawingDocument();
        this.selectStartPage    = this.GetPage();

        let oPos    = oDrDoc.ConvertCoordsFromCursor2(x, y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        if ((this.hitInInnerArea(X, Y) && !this.hitInTextRect(X, Y)) || this.hitToHandles(X, Y) != -1 || this.hitInPath(X, Y)) {
            this.SetInTextBox(false);
        }
        else {
            this.SetInTextBox(true);
            oDoc.SelectionSetStart(x, y, e);
        }

        oDrawingObjects.OnMouseDown(e, X, Y, this.selectStartPage);
    };
    CPdfShape.prototype.AddNewParagraph = function() {
        this.GetDocContent().AddNewParagraph();
        this.FitTextBox();
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.SelectionSetStart = function(X, Y, e) {
        this.selectStartPage = this.GetPage();

        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        let oTransform  = this.invertTransformText;
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        oContent.Selection_SetStart(xContent, yContent, 0, e);
        oContent.RecalculateCurPos();
    };
    
    CPdfShape.prototype.SelectionSetEnd = function(X, Y, e) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        let oTransform  = this.invertTransformText;
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        oContent.Selection_SetEnd(xContent, yContent, 0, e);
    };
    CPdfShape.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorLeft(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfShape.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorRight(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfShape.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorDown(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfShape.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent();
        if (!oContent)
            return;

        oContent.MoveCursorUp(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CPdfShape.prototype.GetDocContent = function() {
        return this.getDocContent();
    };
    CPdfShape.prototype.SetInTextBox = function(bIn) {
        this.isInTextBox = bIn;
    };
    CPdfShape.prototype.IsInTextBox = function() {
        return this.isInTextBox;
    };
    CPdfShape.prototype.EnterText = function(aChars) {
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

        this.FitTextBox();
        this.SetNeedRecalc(true);
        oContent.RecalculateCurPos();

        return true;
    };
    /**
     * Removes char in current position by direction.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CPdfShape.prototype.Remove = function(nDirection, isCtrlKey) {
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
    CPdfShape.prototype.SelectAllText = function() {
        this.GetDocContent().SelectAll();
    };
    /**
     * Exit from this annot.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CPdfShape.prototype.Blur = function() {
        let oDoc        = this.GetDocument();
        let oContent    = this.GetDocContent();
        let oPara       = oContent.GetElement(0);

        oPara.SetApplyToAll(true);
        let sText = oPara.GetSelectedText(true, {NewLine: true});
        oPara.SetApplyToAll(false);

        this.SetInTextBox(false);

        if (this.GetContents() != sText) {
            oDoc.CreateNewHistoryPoint();
            this.SetContents(sText);
            oDoc.TurnOffHistory();
        }
        
        oDoc.GetDrawingDocument().TargetEnd();
    };

    CPdfShape.prototype.FitTextBox = function() {
        return;
        let oDocContent = this.GetDocContent();
        this.recalculateContent();                

        let nContentH = oDocContent.GetSummaryHeight();

        if (nContentH > this.extY) {
            let aCurTextBoxRect = this.GetRect();
            
            // Находим новый textbox rect 
            let xMin = aCurTextBoxRect[0];
            let yMin = aCurTextBoxRect[1];
            let xMax = aCurTextBoxRect[2];
            let yMax = aCurTextBoxRect[3] + (nContentH - this.extY + 0.5) * g_dKoef_mm_to_pix;

            this.SetRect([xMin, yMin, xMax, yMax]);
        }
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
                
        if (oContent.IsSelectionEmpty())
            oContent.RemoveSelection();
    };
    
    /////////////////////////////
    /// saving
    ////////////////////////////

    CPdfShape.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    /////////////////////////////
    /// work with text properties
    ////////////////////////////

    CPdfShape.prototype.SetParaTextPr = function(oParaTextPr) {
        let oContent = this.GetDocContent();
        
        false == this.IsInTextBox() && oContent.SetApplyToAll(true);
        oContent.AddToParagraph(oParaTextPr.Copy());
        false == this.IsInTextBox() && oContent.SetApplyToAll(false);

        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.SetAlign = function(nType) {
        let oContent = this.GetDocContent();
        oContent.SetParagraphAlign(nType);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.GetAlign = function() {
        return this.GetDocContent().GetCalculatedParaPr().GetJc();
    };
    CPdfShape.prototype.SetBold = function(bBold) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Bold : bBold}));
    };
    CPdfShape.prototype.GetBold = function() {
        return !!this.GetCalculatedTextPr().GetBold();        
    };
    CPdfShape.prototype.SetItalic = function(bItalic) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Italic : bItalic}));
    };
    CPdfShape.prototype.GetItalic = function() {
        return !!this.GetCalculatedTextPr().GetItalic();
    };
    CPdfShape.prototype.SetUnderline = function(bUnderline) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Underline : bUnderline}));
    };
    CPdfShape.prototype.GetUnderline = function() {
        return !!this.GetCalculatedTextPr().GetUnderline();
    };
    CPdfShape.prototype.SetHighlight = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({HighlightColor : AscFormat.CreateUniColorRGB(r, g, b)}));
    };
    CPdfShape.prototype.SetStrikeout = function(bStrikeout) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({
			Strikeout  : bStrikeout,
			DStrikeout : false
        }));
    };
    CPdfShape.prototype.GetStrikeout = function() {
        return !!this.GetCalculatedTextPr().GetStrikeout();
    };
    CPdfShape.prototype.SetBaseline = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({VertAlign : nType}));
    };
    CPdfShape.prototype.GetBaseline = function() {
        return this.GetCalculatedTextPr().GetVertAlign();
    };
    CPdfShape.prototype.SetFontSize = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({FontSize : nType}));
    };
    CPdfShape.prototype.GetFontSize = function() {
        return this.GetCalculatedTextPr().GetFontSize();
    };
    CPdfShape.prototype.IncreaseDecreaseFontSize = function(bIncrease) {
        this.GetDocContent().IncreaseDecreaseFontSize(bIncrease);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.SetSpacing = function(nSpacing) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Spacing : nSpacing}));
    };
    CPdfShape.prototype.GetSpacing = function() {
        return this.GetCalculatedTextPr().GetSpacing();
    };
    CPdfShape.prototype.SetFontFamily = function(sFontFamily) {
        let oParaTextPr = new AscCommonWord.ParaTextPr();
		oParaTextPr.Value.RFonts.SetAll(sFontFamily, -1);
        this.SetParaTextPr(oParaTextPr);
    };
    CPdfShape.prototype.GetFontFamily = function() {
        return this.GetCalculatedTextPr().GetFontFamily();
    };
    CPdfShape.prototype.SetTextColor = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Unifill : AscFormat.CreateSolidFillRGB(r, g, b)}));
    };
    CPdfShape.prototype.ChangeTextCase = function(nCaseType) {
        let oContent    = this.GetDocContent();
        let oState      = oContent.GetSelectionState();

        let oChangeEngine = new AscCommonWord.CChangeTextCaseEngine(nCaseType);
		oChangeEngine.ProcessParagraphs(this.GetSelectedParagraphs());

        oContent.SetSelectionState(oState);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.GetSelectedParagraphs = function() {
        let oContent        = this.GetDocContent();
        let aSelectedParas  = [];

        oContent.GetCurrentParagraph(false, aSelectedParas);
        return aSelectedParas;
    };
    CPdfShape.prototype.SetVertAlign = function(nType) {
        this.setVerticalAlign(nType);
        this.SetNeedRecalc(true);
    };

    CPdfShape.prototype.GetAllFonts = function(aFonts) {
        let oContent    = this.GetDocContent();
        let fontMap     = {};
		aFonts          = aFonts || [];

        if (!oContent)
            return aFonts;

        let oPara;
        for (let nPara = 0, nCount = oContent.GetElementsCount(); nPara < nCount; nPara++) {
            oPara = oContent.GetElement(nPara);
            oPara.Get_CompiledPr().TextPr.Document_Get_AllFontNames(fontMap);

            let oRun;
            for (let nRun = 0, nRunCount = oPara.GetElementsCount(); nRun < nRunCount; nRun++) {
                oRun = oPara.GetElement(nRun);
                oRun.Get_CompiledTextPr().Document_Get_AllFontNames(fontMap);
            }
        }
		
        delete fontMap["+mj-lt"];
        delete fontMap["+mn-lt"];
        delete fontMap["+mj-ea"];
        delete fontMap["+mn-ea"];
        delete fontMap["+mj-cs"];
        delete fontMap["+mn-cs"];

        for (let key in fontMap) {
			if (aFonts.includes(key) == false)
                aFonts.push(key);
		}
		
		return aFonts;
    };

    CPdfShape.prototype.SetLineSpacing = function(oSpacing) {
        this.GetDocContent().SetParagraphSpacing(oSpacing);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.GetLineSpacing = function() {
        let oCalcedPr = this.GetCalculatedParaPr();
        return {
            After:  oCalcedPr.GetSpacingAfter(),
            Before: oCalcedPr.GetSpacingBefor()
        }
    };
    CPdfShape.prototype.SetColumnNumber = function(nColumns) {
        this.setColumnNumber(nColumns);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.IncreaseDecreaseIndent = function(bIncrease) {
        // Increase_ParagraphLevel для шейпов из презентаций
        this.GetDocContent().Increase_ParagraphLevel(bIncrease);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.SetNumbering = function(oBullet) {
        this.GetDocContent().Set_ParagraphPresentationNumbering(oBullet);
        this.SetNeedRecalc(true);
    };
    CPdfShape.prototype.ClearFormatting = function(bParaPr, bTextText) {
        this.GetDocContent().ClearParagraphFormatting(bParaPr, bTextText);
        this.SetNeedRecalc(true);
    };
    
    /**
     * Получаем рассчитанные настройки текста (полностью заполненные)
     * @returns {CTextPr}
     */
    CPdfShape.prototype.GetCalculatedTextPr = function() {
        return this.GetDocContent().GetCalculatedTextPr();
    };
    CPdfShape.prototype.GetCalculatedParaPr = function() {
        return this.GetDocContent().GetCalculatedParaPr();
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfShape.prototype.Get_AbsolutePage = function() {
        return this.GetPage();
    };
    CPdfShape.prototype.getLogicDocument = function() {
        return this.GetDocument();
    };
    CPdfShape.prototype.IsThisElementCurrent = function() {
        return true;
    };
    CPdfShape.prototype.getDrawingDocument = function() {
        return Asc.editor.getPDFDoc().GetDrawingDocument();
    };

    window["AscPDF"].CPdfShape = CPdfShape;
})();

