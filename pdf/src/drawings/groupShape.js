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
    function CPdfGroupShape() {
        AscFormat.CGroupShape.call(this);
    }
    
    CPdfGroupShape.prototype.constructor = CPdfGroupShape;
    CPdfGroupShape.prototype = Object.create(AscFormat.CGroupShape.prototype);
    Object.assign(CPdfGroupShape.prototype, AscPDF.CPdfDrawingPrototype.prototype);

    CPdfGroupShape.prototype.IsGroup = function() {
        return true;
    };

    CPdfGroupShape.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalcGeometry();
        this.checkExtentsByDocContent(true, true);
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CPdfGroupShape.prototype.onMouseDown = function(x, y, e) {
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
    CPdfGroupShape.prototype.AddNewParagraph = function() {
        this.GetDocContent().AddNewParagraph();
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.GetDocContent = function() {
        return this.getTargetDocContent();
    };
    CPdfGroupShape.prototype.SelectAllText = function() {
        this.GetDocContent().SelectAll();
    };

    CPdfGroupShape.prototype.onMouseUp = function(x, y, e) {
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
    /// work with text properties
    ////////////////////////////

    CPdfGroupShape.prototype.SetParaTextPr = function(oParaTextPr) {
        let oContent = this.GetDocContent();
        
        false == this.IsInTextBox() && oContent.SetApplyToAll(true);
        oContent.AddToParagraph(oParaTextPr.Copy());
        false == this.IsInTextBox() && oContent.SetApplyToAll(false);

        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.SetAlign = function(nType) {
        let oContent = this.GetDocContent();
        oContent.SetParagraphAlign(nType);
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.GetAlign = function() {
        return this.GetDocContent().GetCalculatedParaPr().GetJc();
    };
    CPdfGroupShape.prototype.SetBold = function(bBold) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Bold : bBold}));
    };
    CPdfGroupShape.prototype.GetBold = function() {
        return !!this.GetCalculatedTextPr().GetBold();        
    };
    CPdfGroupShape.prototype.SetItalic = function(bItalic) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Italic : bItalic}));
    };
    CPdfGroupShape.prototype.GetItalic = function() {
        return !!this.GetCalculatedTextPr().GetItalic();
    };
    CPdfGroupShape.prototype.SetUnderline = function(bUnderline) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Underline : bUnderline}));
    };
    CPdfGroupShape.prototype.GetUnderline = function() {
        return !!this.GetCalculatedTextPr().GetUnderline();
    };
    CPdfGroupShape.prototype.SetHighlight = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({HighlightColor : AscFormat.CreateUniColorRGB(r, g, b)}));
    };
    CPdfGroupShape.prototype.SetStrikeout = function(bStrikeout) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({
			Strikeout  : bStrikeout,
			DStrikeout : false
        }));
    };
    CPdfGroupShape.prototype.GetStrikeout = function() {
        return !!this.GetCalculatedTextPr().GetStrikeout();
    };
    CPdfGroupShape.prototype.SetBaseline = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({VertAlign : nType}));
    };
    CPdfGroupShape.prototype.GetBaseline = function() {
        return this.GetCalculatedTextPr().GetVertAlign();
    };
    CPdfGroupShape.prototype.SetFontSize = function(nType) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({FontSize : nType}));
    };
    CPdfGroupShape.prototype.GetFontSize = function() {
        return this.GetCalculatedTextPr().GetFontSize();
    };
    CPdfGroupShape.prototype.IncreaseDecreaseFontSize = function(bIncrease) {
        this.GetDocContent().IncreaseDecreaseFontSize(bIncrease);
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.SetSpacing = function(nSpacing) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Spacing : nSpacing}));
    };
    CPdfGroupShape.prototype.GetSpacing = function() {
        return this.GetCalculatedTextPr().GetSpacing();
    };
    CPdfGroupShape.prototype.SetFontFamily = function(sFontFamily) {
        let oParaTextPr = new AscCommonWord.ParaTextPr();
		oParaTextPr.Value.RFonts.SetAll(sFontFamily, -1);
        this.SetParaTextPr(oParaTextPr);
    };
    CPdfGroupShape.prototype.GetFontFamily = function() {
        return this.GetCalculatedTextPr().GetFontFamily();
    };
    CPdfGroupShape.prototype.SetTextColor = function(r, g, b) {
        this.SetParaTextPr(new AscCommonWord.ParaTextPr({Unifill : AscFormat.CreateSolidFillRGB(r, g, b)}));
    };
    CPdfGroupShape.prototype.ChangeTextCase = function(nCaseType) {
        let oContent    = this.GetDocContent();
        let oState      = oContent.GetSelectionState();

        let oChangeEngine = new AscCommonWord.CChangeTextCaseEngine(nCaseType);
		oChangeEngine.ProcessParagraphs(this.GetSelectedParagraphs());

        oContent.SetSelectionState(oState);
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.GetSelectedParagraphs = function() {
        let oContent        = this.GetDocContent();
        let aSelectedParas  = [];

        oContent.GetCurrentParagraph(false, aSelectedParas);
        return aSelectedParas;
    };
    CPdfGroupShape.prototype.SetVertAlign = function(nType) {
        this.setVerticalAlign(nType);
        this.SetNeedRecalc(true);
    };

    CPdfGroupShape.prototype.GetAllFonts = function(fontMap) {
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

    CPdfGroupShape.prototype.SetLineSpacing = function(oSpacing) {
        this.GetDocContent().SetParagraphSpacing(oSpacing);
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.GetLineSpacing = function() {
        let oCalcedPr = this.GetCalculatedParaPr();
        return {
            After:  oCalcedPr.GetSpacingAfter(),
            Before: oCalcedPr.GetSpacingBefor()
        }
    };
    CPdfGroupShape.prototype.SetColumnNumber = function(nColumns) {
        this.setColumnNumber(nColumns);
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.IncreaseDecreaseIndent = function(bIncrease) {
        // Increase_ParagraphLevel для шейпов из презентаций
        this.GetDocContent().Increase_ParagraphLevel(bIncrease);
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.SetNumbering = function(oBullet) {
        this.GetDocContent().Set_ParagraphPresentationNumbering(oBullet);
        this.SetNeedRecalc(true);
    };
    CPdfGroupShape.prototype.ClearFormatting = function(bParaPr, bTextText) {
        this.GetDocContent().ClearParagraphFormatting(bParaPr, bTextText);
        this.SetNeedRecalc(true);
    };
    
    /**
     * Получаем рассчитанные настройки текста (полностью заполненные)
     * @returns {CTextPr}
     */
    CPdfGroupShape.prototype.GetCalculatedTextPr = function() {
        return this.GetDocContent().GetCalculatedTextPr();
    };
    CPdfGroupShape.prototype.GetCalculatedParaPr = function() {
        return this.GetDocContent().GetCalculatedParaPr();
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfGroupShape.prototype.updateSelectionState = function () {
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
    CPdfGroupShape.prototype.copy = function(oPr) {
        let copy = new CPdfGroupShape();
        this.copy2(copy, oPr);
        let drawing = copy.getDrawing();

        if (drawing) {
            for (let i = 0; i < drawing.spTree.length; i += 1) {
                let obj = drawing.spTree[i];
                if (obj.getObjectType() === AscDFH.historyitem_type_Shape) {
                    obj.copyTextInfoFromShapeToPoint();
                }
            }
        }

        if (oPr && oPr.contentCopyPr && oPr.contentCopyPr.Comparison) {
            copy.generateDrawingPart();
        }

        if ((!oPr || !oPr.bSkipRedactsIds) && this.GetRedactIds) {
            this.GetRedactIds().forEach(function(id) {
                copy.AddRedactId(id);
            });
        }

        return copy;
    };

    window["AscPDF"].CPdfGroupShape = CPdfGroupShape;
})();

