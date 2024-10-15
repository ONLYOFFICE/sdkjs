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
    function CPdfSmartArt()
    {
        AscFormat.SmartArt.call(this);
                
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
    
    CPdfSmartArt.prototype.constructor = CPdfSmartArt;
    CPdfSmartArt.prototype = Object.create(AscFormat.SmartArt.prototype);
    Object.assign(CPdfSmartArt.prototype, AscPDF.PdfDrawingPrototype.prototype);

    CPdfSmartArt.prototype.IsSmartArt = function() {
        return true;
    };

    CPdfSmartArt.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalcGeometry();
        this.checkExtentsByDocContent(true, true);
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CPdfSmartArt.prototype.onMouseDown = function(x, y, e) {
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
    CPdfSmartArt.prototype.AddNewParagraph = function() {
        this.GetDocContent().AddNewParagraph();
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.GetDocContent = function() {
        return this.getTargetDocContent();
    };
    CPdfSmartArt.prototype.SelectAllText = function() {
        this.GetDocContent().SelectAll();
    };

    CPdfSmartArt.prototype.onMouseUp = function(x, y, e) {
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

    CPdfSmartArt.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    /////////////////////////////
    /// work with text properties
    ////////////////////////////

    CPdfSmartArt.prototype.SetParaTextPr = function(oParaTextPr) {
        let oContent = this.GetDocContent();
        
        false == this.IsInTextBox() && oContent.SetApplyToAll(true);
        oContent.AddToParagraph(oParaTextPr.Copy());
        false == this.IsInTextBox() && oContent.SetApplyToAll(false);

        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.SetAlign = function(nType) {
        let oContent = this.GetDocContent();
        oContent.SetParagraphAlign(nType);
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.GetAlign = function() {
        return this.GetDocContent().GetCalculatedParaPr().GetJc();
    };
    CPdfSmartArt.prototype.SetBold = function(bBold) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Bold : bBold}));
    };
    CPdfSmartArt.prototype.GetBold = function() {
        return !!this.GetCalculatedTextPr().GetBold();        
    };
    CPdfSmartArt.prototype.SetItalic = function(bItalic) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Italic : bItalic}));
    };
    CPdfSmartArt.prototype.GetItalic = function() {
        return !!this.GetCalculatedTextPr().GetItalic();
    };
    CPdfSmartArt.prototype.SetUnderline = function(bUnderline) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Underline : bUnderline}));
    };
    CPdfSmartArt.prototype.GetUnderline = function() {
        return !!this.GetCalculatedTextPr().GetUnderline();
    };
    CPdfSmartArt.prototype.SetHighlight = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({HighlightColor : AscFormat.CreateUniColorRGB(r, g, b)}));
    };
    CPdfSmartArt.prototype.SetStrikeout = function(bStrikeout) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({
			Strikeout  : bStrikeout,
			DStrikeout : false
        }));
    };
    CPdfSmartArt.prototype.GetStrikeout = function() {
        return !!this.GetCalculatedTextPr().GetStrikeout();
    };
    CPdfSmartArt.prototype.SetBaseline = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({VertAlign : nType}));
    };
    CPdfSmartArt.prototype.GetBaseline = function() {
        return this.GetCalculatedTextPr().GetVertAlign();
    };
    CPdfSmartArt.prototype.SetFontSize = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({FontSize : nType}));
    };
    CPdfSmartArt.prototype.GetFontSize = function() {
        return this.GetCalculatedTextPr().GetFontSize();
    };
    CPdfSmartArt.prototype.IncreaseDecreaseFontSize = function(bIncrease) {
        this.GetDocContent().IncreaseDecreaseFontSize(bIncrease);
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.SetSpacing = function(nSpacing) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Spacing : nSpacing}));
    };
    CPdfSmartArt.prototype.GetSpacing = function() {
        return this.GetCalculatedTextPr().GetSpacing();
    };
    CPdfSmartArt.prototype.SetFontFamily = function(sFontFamily) {
        let oParaTextPr = new AscCommonWord.ParaTextPr();
		oParaTextPr.Value.RFonts.SetAll(sFontFamily, -1);
        this.SetParaTextPr(oParaTextPr);
    };
    CPdfSmartArt.prototype.GetFontFamily = function() {
        return this.GetCalculatedTextPr().GetFontFamily();
    };
    CPdfSmartArt.prototype.SetTextColor = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Unifill : AscFormat.CreateSolidFillRGB(r, g, b)}));
    };
    CPdfSmartArt.prototype.ChangeTextCase = function(nCaseType) {
        let oContent    = this.GetDocContent();
        let oState      = oContent.GetSelectionState();

        let oChangeEngine = new AscCommonWord.CChangeTextCaseEngine(nCaseType);
		oChangeEngine.ProcessParagraphs(this.GetSelectedParagraphs());

        oContent.SetSelectionState(oState);
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.GetSelectedParagraphs = function() {
        let oContent        = this.GetDocContent();
        let aSelectedParas  = [];

        oContent.GetCurrentParagraph(false, aSelectedParas);
        return aSelectedParas;
    };
    CPdfSmartArt.prototype.SetVertAlign = function(nType) {
        this.setVerticalAlign(nType);
        this.SetNeedRecalc(true);
    };

    CPdfSmartArt.prototype.GetAllFonts = function(fontMap) {
        let aContents = [];
        this.getAllDocContents(aContents);

        fontMap = fontMap || {};

        for (let i = 0; i < aContents.length; i++) {
            let oContent = aContents[i];

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
        }
        
        
        delete fontMap["+mj-lt"];
        delete fontMap["+mn-lt"];
        delete fontMap["+mj-ea"];
        delete fontMap["+mn-ea"];
        delete fontMap["+mj-cs"];
        delete fontMap["+mn-cs"];
        
        return fontMap;
    };

    CPdfSmartArt.prototype.SetLineSpacing = function(oSpacing) {
        this.GetDocContent().SetParagraphSpacing(oSpacing);
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.GetLineSpacing = function() {
        let oCalcedPr = this.GetCalculatedParaPr();
        return {
            After:  oCalcedPr.GetSpacingAfter(),
            Before: oCalcedPr.GetSpacingBefor()
        }
    };
    CPdfSmartArt.prototype.SetColumnNumber = function(nColumns) {
        this.setColumnNumber(nColumns);
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.IncreaseDecreaseIndent = function(bIncrease) {
        // Increase_ParagraphLevel для шейпов из презентаций
        this.GetDocContent().Increase_ParagraphLevel(bIncrease);
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.SetNumbering = function(oBullet) {
        this.GetDocContent().Set_ParagraphPresentationNumbering(oBullet);
        this.SetNeedRecalc(true);
    };
    CPdfSmartArt.prototype.ClearFormatting = function(bParaPr, bTextText) {
        this.GetDocContent().ClearParagraphFormatting(bParaPr, bTextText);
        this.SetNeedRecalc(true);
    };
    
    /**
     * Получаем рассчитанные настройки текста (полностью заполненные)
     * @returns {CTextPr}
     */
    CPdfSmartArt.prototype.GetCalculatedTextPr = function() {
        return this.GetDocContent().GetCalculatedTextPr();
    };
    CPdfSmartArt.prototype.GetCalculatedParaPr = function() {
        return this.GetDocContent().GetCalculatedParaPr();
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfSmartArt.prototype.updateSelectionState = function () {
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

    window["AscPDF"].CPdfSmartArt = CPdfSmartArt;
})();

