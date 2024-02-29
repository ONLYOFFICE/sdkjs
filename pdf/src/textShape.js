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
    function CTextShape(sName, nPage, aOrigRect, oDoc)
    {
        AscFormat.CShape.call(this);
                
        this._page          = nPage;
        this._apIdx         = undefined; // индекс объекта в файле
        this._rect          = [];       // scaled rect
        this._origRect      = [];
        this._richContents  = [];

        // internal
        this._doc                   = oDoc;
        this._needRecalc            = true;
        this._wasChanged            = false; // была ли изменена
        this._bDrawFromStream       = false; // нужно ли рисовать из стрима
        this._hasOriginView         = false; // имеет ли внешний вид из файла

        this.Internal_InitRect(aOrigRect);
        initTextShape(this);
    }
    
    CTextShape.prototype.constructor = CTextShape;
    CTextShape.prototype = Object.create(AscFormat.CShape.prototype);

    CTextShape.prototype.Internal_InitRect = function(aOrigRect) {
        let nPage = this.GetPage();
        let oViewer = Asc.editor.getDocumentRenderer();
        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        this._rect = [aOrigRect[0] * nScaleX, aOrigRect[1] * nScaleY, aOrigRect[2] * nScaleX, aOrigRect[3] * nScaleY];
        this._origRect = aOrigRect;
    };
    CTextShape.prototype.IsNeedDrawFromStream = function() {
       return false; 
    };
    CTextShape.prototype.IsAnnot = function() {
       return false;
    };
    CTextShape.prototype.IsForm = function() {
       return false;
    };
    CTextShape.prototype.IsTextShape = function() {
       return true;
    };
    CTextShape.prototype.ShowBorder = function(bShow) {
        let oLine = this.pen;

        if (bShow) {
            oLine.setFill(AscFormat.CreateSolidFillRGBA(0, 0, 0, 255));
        }
        else {
            oLine.setFill(AscFormat.CreateNoFillUniFill());
        }

        this.AddToRedraw();
    };
    CTextShape.prototype.SetApIdx = function(nIdx) {
        this.GetDocument().UpdateApIdx(nIdx);
        this._apIdx = nIdx;
    };
    CTextShape.prototype.GetApIdx = function() {
        return this._apIdx;
    };
    CTextShape.prototype.GetDocument = function() {
        return this._doc;
    };
    CTextShape.prototype.GetPage = function() {
        return this._page;
    };
    CTextShape.prototype.AddToRedraw = function() {
        let oViewer = Asc.editor.getDocumentRenderer();
        let _t      = this;
        
        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[_t.GetPage()])
                oViewer.pagesInfo.pages[_t.GetPage()].needRedrawTextShapes = true;
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    CTextShape.prototype.GetRect = function() {
        return this._rect;
    };
    CTextShape.prototype.GetOrigRect = function() {
        return this._origRect;
    };
    
    CTextShape.prototype.SetRect = function(aRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();
        let nPage       = this.GetPage();

        oDoc.History.Add(new CChangesPDFTxShapeRect(this, this.GetRect(), aRect));

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        this._rect = aRect;

        this._pagePos = {
            x: aRect[0],
            y: aRect[1],
            w: (aRect[2] - aRect[0]),
            h: (aRect[3] - aRect[1])
        };

        this._origRect[0] = this._rect[0] / nScaleX;
        this._origRect[1] = this._rect[1] / nScaleY;
        this._origRect[2] = this._rect[2] / nScaleX;
        this._origRect[3] = this._rect[3] / nScaleY;

        this.spPr.xfrm.extX = this._pagePos.w * g_dKoef_pix_to_mm;
        this.spPr.xfrm.extY = this._pagePos.h * g_dKoef_pix_to_mm;
        this.spPr.xfrm.offX = aRect[0] * g_dKoef_pix_to_mm;
        this.spPr.xfrm.offY = aRect[1] * g_dKoef_pix_to_mm;
        this.updateTransformMatrix();

        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.SetRot = function(dAngle) {
        let oDoc = this.GetDocument();

        oDoc.History.Add(new CChangesPDFTxShapeRot(this, this.GetRot(), dAngle));

        this.changeRot(dAngle);
        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.GetRot = function() {
        return this.rot;
    };
    CTextShape.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        let aRect = this.GetRect();
        this.spPr.xfrm.setOffX(aRect[0] * g_dKoef_pix_to_mm);
        this.spPr.xfrm.setOffY(aRect[1] * g_dKoef_pix_to_mm);
        
        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalcGeometry();
        this.recalculateContent();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CTextShape.prototype.IsNeedRecalc = function() {
       return this._needRecalc;
    };
    CTextShape.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
       if (bRecalc == false) {
           this._needRecalc = false;
       }
       else {
           this._needRecalc = true;
           if (bSkipAddToRedraw != true)
               this.AddToRedraw();
       }
    };
    CTextShape.prototype.Draw = function(oGraphicsWord) {
        this.Recalculate();
        this.draw(oGraphicsWord);
    };
    CTextShape.prototype.onMouseDown = function(x, y, e) {
        let oDoc                = this.GetDocument();
        let oDrDoc              = oDoc.GetDrawingDocument();
        this.selectStartPage    = this.GetPage();

        let oPos    = oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        if (this.hitToHandles(X, Y) != -1 || this.hitInPath(X, Y)) {
            this.SetInTextBox(false);
        }
        else {
            this.SetInTextBox(true);
            oDoc.SelectionSetStart(x, y, e);
            oDoc.SetLocalHistory();
        }
    };
    CTextShape.prototype.AddNewParagraph = function() {
        this.GetDocContent().AddNewParagraph();
        this.FitTextBox();
        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.SelectionSetStart = function(X, Y, e) {
        this.selectStartPage = this.GetPage();

        let oContent = this.GetDocContent();
        
        let oTransform  = this.invertTransformText;
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        oContent.Selection_SetStart(xContent, yContent, 0, e);
        oContent.RecalculateCurPos();
    };
    
    CTextShape.prototype.SelectionSetEnd = function(X, Y, e) {
        let oContent = this.GetDocContent();
        
        let oTransform  = this.invertTransformText;
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        oContent.Selection_SetEnd(xContent, yContent, 0, e);
    };
    CTextShape.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorLeft(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CTextShape.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorRight(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CTextShape.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorDown(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CTextShape.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorUp(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CTextShape.prototype.GetDocContent = function() {
        return this.getDocContent();
    };
    CTextShape.prototype.SetNeedUpdateRC = function(bUpdate) {
        this._needUpdateRC = bUpdate;
    };
    CTextShape.prototype.IsNeedUpdateRC = function() {
        return this._needUpdateRC;
    };
    CTextShape.prototype.SetInTextBox = function(bIn) {
        let oDoc = this.GetDocument();

        this.isInTextBox = bIn;

        if (bIn == false && this.IsNeedUpdateRC()) {
            oDoc.SetGlobalHistory();
            oDoc.CreateNewHistoryPoint();
            oDoc.History.Add(new CChangesPDFTxShapeRC(this, this.GetRichContents(), this.GetRichContents(true)));
            this.SetNeedUpdateRC(false);
        }
    };
    CTextShape.prototype.IsInTextBox = function() {
        return this.isInTextBox;
    };
    CTextShape.prototype.SetRichContents = function(aRCInfo) {
        let oDoc            = this.GetDocument();
        let oContent        = this.GetDocContent();
        oContent.ClearContent();
        
        let oLastUsedPara   = oContent.GetElement(0);
        oLastUsedPara.RemoveFromContent(0, oLastUsedPara.GetElementsCount());

        let oRCInfo;
        let oRun, oRFonts;

        this._richContents = aRCInfo;
        oDoc.History.Add(new CChangesPDFTxShapeRC(this, this.GetRichContents(), aRCInfo));

        for (let i = 0; i < aRCInfo.length; i++) {
            oRCInfo = aRCInfo[i];

            oRun = new ParaRun(oLastUsedPara, false);
            oRFonts = new CRFonts();
            // oRFonts.SetAll(oRCInfo["name"], -1);
            oRFonts.SetAll("Arial", -1);

            oRun.SetBold(Boolean(oRCInfo["bold"]));
            oRun.SetItalic(Boolean(oRCInfo["italic"]));
            oRun.SetStrikeout(Boolean(oRCInfo["strikethrough"]));
            oRun.SetUnderline(Boolean(oRCInfo["underlined"]));
            oRun.SetFontSize(oRCInfo["size"]);
            oRun.Set_RFonts2(oRFonts);

            editor.private_CreateApiRun(oRun).SetColor(0, 0, 0, false);

            let oIterator = oRCInfo["text"].replace('\r', '').getUnicodeIterator();
            while (oIterator.check()) {
                let runElement = AscPDF.codePointToRunElement(oIterator.value());
				oRun.Add(runElement);
                
                oIterator.next();
            }

            oLastUsedPara.AddToContentToEnd(oRun);

            if (oRCInfo["text"].indexOf('\r') != -1) {
                oLastUsedPara = new AscWord.Paragraph(oContent, true);
                oContent.Internal_Content_Add(oContent.GetElementsCount(), oLastUsedPara);
            }
        }

        this.SetNeedRecalc(true);
        this.SetNeedUpdateRC(false);
    };
    CTextShape.prototype.GetRichContents = function(bCalced) {
        if (!bCalced)
            return this._richContents;

        let oContent = this.GetDocContent();
        let aRCInfo = [];

        for (let i = 0, nCount = oContent.GetElementsCount(); i < nCount; i++) {
            let oPara = oContent.GetElement(i);

            for (let j = 0, nRunsCount = oPara.GetElementsCount(); j < nRunsCount; j++) {
                let oRun = oPara.GetElement(j);
                let sText = oRun.GetText();
                if (sText) {
                    aRCInfo.push({
                        "bold":             oRun.Get_Bold(),
                        "italic":           oRun.Get_Italic(),
                        "strikethrough":    oRun.Get_Strikeout(),
                        "underlined":       oRun.Get_Underline(),
                        "size":             oRun.Get_FontSize(),
                        "name":             oRun.Get_RFonts().Ascii.Name,
                        "text":             sText
                    });
                }
            }

            if (aRCInfo[aRCInfo.length - 1])
                aRCInfo[aRCInfo.length - 1]["text"] += '\r';
        }

        return aRCInfo;
    };
    CTextShape.prototype.EnterText = function(aChars) {
        let oDoc        = this.GetDocument();
        let oContent    = this.GetDocContent();
        let oParagraph  = oContent.GetCurrentParagraph();

        oDoc.SetLocalHistory();
        oDoc.CreateNewHistoryPoint(this);

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
        this.SetNeedUpdateRC(true);
        oContent.RecalculateCurPos();

        return true;
    };
    /**
     * Removes char in current position by direction.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CTextShape.prototype.Remove = function(nDirection, isCtrlKey) {
        let oDoc = this.GetDocument();
        oDoc.CreateNewHistoryPoint(this);

        let oContent = this.GetDocContent();
        oContent.Remove(nDirection, true, false, false, isCtrlKey);
        oContent.RecalculateCurPos();
        this.SetNeedRecalc(true);

        if (AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
        }
        else {
            this.SetNeedRecalc(true);
            this.SetNeedUpdateRC(true);
        }
    };
    CTextShape.prototype.SelectAllText = function() {
        this.GetDocContent().SelectAll();
    };
    /**
     * Exit from this annot.
     * @memberof CTextField
     * @typeofeditors ["PDF"]
     */
    CTextShape.prototype.Blur = function() {
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

    CTextShape.prototype.FitTextBox = function() {
        let oDocContent     = this.GetDocContent();
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

    CTextShape.prototype.onMouseUp = function(x, y, e) {
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

    CTextShape.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    /////////////////////////////
    /// work with text properties
    ////////////////////////////

    CTextShape.prototype.SetParaTextPr = function(oParaTextPr) {
        let oContent = this.GetDocContent();
        var StartPos = oContent.Selection.StartPos;
        var EndPos   = oContent.Selection.EndPos;

        if (EndPos < StartPos)
        {
            var Temp = StartPos;
            StartPos = EndPos;
            EndPos   = Temp;
        }

        for (var nIndex = StartPos; nIndex <= EndPos; nIndex++) {
            oContent.GetElement(nIndex).Add(oParaTextPr.Copy());
        }

        this.SetNeedRecalc(true);
        this.SetNeedUpdateRC(true);
    };
    CTextShape.prototype.SetAlign = function(nType) {
        let oContent = this.GetDocContent();
        oContent.SetParagraphAlign(nType);
        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.GetAlign = function() {
        return this.GetDocContent().GetCalculatedParaPr().GetJc();
    };
    CTextShape.prototype.SetBold = function(bBold) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Bold : bBold}));
    };
    CTextShape.prototype.GetBold = function() {
        return !!this.GetCalculatedTextPr().GetBold();        
    };
    CTextShape.prototype.SetItalic = function(bItalic) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Italic : bItalic}));
    };
    CTextShape.prototype.GetItalic = function() {
        return !!this.GetCalculatedTextPr().GetItalic();
    };
    CTextShape.prototype.SetUnderline = function(bUnderline) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Underline : bUnderline}));
    };
    CTextShape.prototype.GetUnderline = function() {
        return !!this.GetCalculatedTextPr().GetUnderline();
    };
    CTextShape.prototype.SetHighlight = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({HighlightColor : AscFormat.CreateUniColorRGB(r, g, b)}));
    };
    CTextShape.prototype.SetStrikeout = function(bStrikeout) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({
			Strikeout  : bStrikeout,
			DStrikeout : false
        }));
    };
    CTextShape.prototype.GetStrikeout = function() {
        return !!this.GetCalculatedTextPr().GetStrikeout();
    };
    CTextShape.prototype.SetBaseline = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({VertAlign : nType}));
    };
    CTextShape.prototype.GetBaseline = function() {
        return this.GetCalculatedTextPr().GetVertAlign();
    };
    CTextShape.prototype.SetFontSize = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({FontSize : nType}));
    };
    CTextShape.prototype.GetFontSize = function() {
        return this.GetCalculatedTextPr().GetFontSize();
    };
    CTextShape.prototype.IncreaseDecreaseFontSize = function(bIncrease) {
        this.GetDocContent().IncreaseDecreaseFontSize(bIncrease);
        this.SetNeedRecalc(true);
        this.SetNeedUpdateRC(true);
    };
    CTextShape.prototype.SetSpacing = function(nSpacing) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Spacing : nSpacing}));
    };
    CTextShape.prototype.GetSpacing = function() {
        return this.GetCalculatedTextPr().GetSpacing();
    };
    CTextShape.prototype.SetFontFamily = function(sFontFamily) {
        let oParaTextPr = new AscCommonWord.ParaTextPr();
		oParaTextPr.Value.RFonts.SetAll(sFontFamily, -1);
        this.SetParaTextPr(oParaTextPr);
    };
    CTextShape.prototype.GetFontFamily = function() {
        return this.GetCalculatedTextPr().GetFontFamily();
    };
    CTextShape.prototype.SetTextColor = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Unifill : AscFormat.CreateSolidFillRGB(r, g, b)}));
    };
    CTextShape.prototype.ChangeTextCase = function(nCaseType) {
        let oContent    = this.GetDocContent();
        let oState      = oContent.GetSelectionState();

        let oChangeEngine = new AscCommonWord.CChangeTextCaseEngine(nCaseType);
		oChangeEngine.ProcessParagraphs(this.GetSelectedParagraphs());

        oContent.SetSelectionState(oState);
        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.GetSelectedParagraphs = function() {
        let oContent        = this.GetDocContent();
        let aSelectedParas  = [];

        oContent.GetCurrentParagraph(false, aSelectedParas);
        return aSelectedParas;
    };
    CTextShape.prototype.SetVertAlign = function(nType) {
        this.setVerticalAlign(nType);
        this.SetNeedRecalc(true);
    };

    CTextShape.prototype.GetAllFonts = function(aFonts) {
        let oContent    = this.GetDocContent();
        let fontMap     = {};
		aFonts          = aFonts || [];

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

    CTextShape.prototype.SetLineSpacing = function(oSpacing) {
        this.GetDocContent().SetParagraphSpacing(oSpacing);
    };
    CTextShape.prototype.GetLineSpacing = function() {
        let oCalcedPr = this.GetCalculatedParaPr();
        return {
            After:  oCalcedPr.GetSpacingAfter(),
            Before: oCalcedPr.GetSpacingBefor()
        }
    };
    CTextShape.prototype.SetColumnNumber = function(nColumns) {
        this.setColumnNumber(nColumns);
        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.IncreaseDecreaseIndent = function(bIncrease) {
        // Increase_ParagraphLevel для шейпов из презентаций
        this.GetDocContent().Increase_ParagraphLevel(bIncrease);
        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.SetNumbering = function(oBullet) {
        this.GetDocContent().Set_ParagraphPresentationNumbering(oBullet);
        this.SetNeedRecalc(true);
    };
    CTextShape.prototype.ClearFormatting = function(bParaPr, bTextText) {
        this.GetDocContent().ClearParagraphFormatting(bParaPr, bTextText);
        this.SetNeedRecalc(true);
    };
    
    /**
     * Получаем рассчитанные настройки текста (полностью заполненные)
     * @returns {CTextPr}
     */
    CTextShape.prototype.GetCalculatedTextPr = function() {
        return this.GetDocContent().GetCalculatedTextPr();
    };
    CTextShape.prototype.GetCalculatedParaPr = function() {
        return this.GetDocContent().GetCalculatedParaPr();
    };

    function initTextShape(oParentTextShape) {
        let aShapeRectInMM = oParentTextShape.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });
        let xMax = aShapeRectInMM[2];
        let xMin = aShapeRectInMM[0];
        let yMin = aShapeRectInMM[1];
        let yMax = aShapeRectInMM[3];

        oParentTextShape.setSpPr(new AscFormat.CSpPr());
        oParentTextShape.spPr.setParent(oParentTextShape);
        oParentTextShape.spPr.setXfrm(new AscFormat.CXfrm());
        oParentTextShape.spPr.xfrm.setParent(oParentTextShape.spPr);
        oParentTextShape.spPr.setLn(new AscFormat.CLn());

        oParentTextShape.spPr.xfrm.setOffX(xMin);
        oParentTextShape.spPr.xfrm.setOffY(xMax);
        oParentTextShape.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oParentTextShape.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oParentTextShape.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
        // oParentTextShape.setStyle(AscFormat.CreateDefaultShapeStyle());
        oParentTextShape.setBDeleted(false);
        oParentTextShape.recalculate();
        oParentTextShape.brush = AscFormat.CreateNoFillUniFill();

        oParentTextShape.createTextBody();
        oParentTextShape.txBody.bodyPr.setInsets(0.5,0.5,0.5,0.5);
        oParentTextShape.txBody.bodyPr.horzOverflow = AscFormat.nHOTClip;
        oParentTextShape.txBody.bodyPr.vertOverflow = AscFormat.nVOTClip;

        let oLine       = oParentTextShape.pen;
        oLine.setPrstDash(Asc.c_oDashType.sysDot);
        oLine.setW(25.4 / 72.0 * 36000);
        if (oParentTextShape.GetDocument().IsTextEditMode() == false)
            oLine.setFill(AscFormat.CreateNoFillUniFill());

        oParentTextShape.SetRichContents(JSON.parse('[{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"Text\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"dwa"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"dawd\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"aw\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"da\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"d\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"dw"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"adawd\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"daw\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"dwa\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"d\\r"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"awd3212312323233"},{"bold":0,"italic":0,"strikethrough":0,"underlined":0,"size":10,"color":[0,0,0],"name":"","text":"\\r"}]'));
        // editor.private_CreateApiRun(oRun).SetColor(0, 0, 0, false);
        
        oParentTextShape.setTxBox(true);
    }

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    AscFormat.CTextBody.prototype.getDrawingDocument = function() {
        return this.parent && this.parent.getDrawingDocument && this.parent.getDrawingDocument();
    };

    window["AscPDF"].CTextShape = CTextShape;
})();

