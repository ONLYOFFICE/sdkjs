/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
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
	 * Class representing a text field.
	 * @constructor
     * @extends {CBaseField}
	 */
    function CTextField(sName, nPage, aRect, oDoc)
    {
        AscPDF.CBaseField.call(this, sName, AscPDF. FIELD_TYPES.text, nPage, aRect, oDoc);
        
        this._alignment         = AscPDF.ALIGN_TYPE.left;
        this._charLimit         = 0;
        this._comb              = false;
        this._defaultStyle      = Object.assign({}, {}); // to do (must not be fileSelect flag)
        this._doNotScroll       = false;
        this._doNotSpellCheck   = false;
        this._multiline         = false;
        this._password          = false;
        this._richText          = false; // to do связанные свойства, методы
        this._richValue         = [];
        this._textFont          = AscPDF.DEFAULT_FIELD_FONT;
        this._fileSelect        = false;
        this._value             = undefined;
		this._displayValue      = "";
		this._useDisplayValue   = true;

        // internal
        oDoc.StartNoHistoryMode();
		this.content = new AscPDF.CTextBoxContent(this, oDoc);
        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
		this.contentFormat = new AscPDF.CTextBoxContent(this, oDoc, true);
        oDoc.EndNoHistoryMode();

        this._scrollInfo = null;
        this._markRect = {};
    }
    CTextField.prototype = Object.create(AscPDF.CBaseField.prototype);
	CTextField.prototype.constructor = CTextField;
    
    CTextField.prototype.SetComb = function(bComb) {
        if (this._comb == bComb && this.GetCharLimit() != 0)
            return;
            
        if (bComb == true) {
            this._comb = true;
            this._doNotScroll = true;
        }
        else {
            this._comb = false;
        }

        this.content.GetElement(0).Content.forEach(function(run) {
            run.RecalcInfo.Measure = true;
        });
        this.contentFormat.GetElement(0).Content.forEach(function(run) {
            run.RecalcInfo.Measure = true;
        });

        this.SetNeedRecalc(true);
        this.SetWasChanged(true);
    };
    CTextField.prototype.IsComb = function() {
        return this._comb;
    };
    CTextField.prototype.IsCanEditText = function() {
        let oDoc = this.GetDocument();
        if (oDoc.activeForm == this && this.IsInForm())
            return true;
        
        return false;
    };
    CTextField.prototype.SetCharLimit = function(nChars) {
        let oViewer = editor.getDocumentRenderer();

        if (typeof(nChars) == "number" && nChars <= 500 && nChars > 0 && this.IsFileSelect() === false) {
            nChars = Math.round(nChars);
            if (this._charLimit != nChars) {

                if (oViewer.IsOpenFormsInProgress != true) {
                    let sText = this.content.GetElement(0).GetText({ParaEndToSpace: false});
                    this.content.replaceAllText(sText.slice(0, Math.min(nChars, sText.length)));
                }
            }

            this._charLimit = nChars;
            this.SetNeedRecalc(true);
            this.SetWasChanged(true);
        }
    };
    CTextField.prototype.GetCharLimit = function() {
        return this._charLimit;
    };
    CTextField.prototype.SetDoNotScroll = function(bNot) {
        this._doNotScroll = bNot;
    };
    CTextField.prototype.IsDoNotScroll = function() {
        return this._doNotScroll;
    };
    CTextField.prototype.SetDoNotSpellCheck = function(bNot) {
        this._doNotSpellCheck = bNot;
    };
    CTextField.prototype.IsDoNotSpellCheck = function() {
        return this._doNotSpellCheck;
    };
    CTextField.prototype.SetFileSelect = function(bFileSelect) {
        if (bFileSelect === true && this.IsMultiline() != true && this._charLimit === 0
            && this.password != true && this.defaultValue == "") {
                this._fileSelect = true;
            }
        else if (bFileSelect === false) {
            this._fileSelect = false;
        }
    };
    CTextField.prototype.IsFileSelect = function() {
        return this._fileSelect;
    };
    CTextField.prototype.SetMultiline = function(bMultiline) {
        if (bMultiline == true && this.fileSelect != true) {
            this.content.SetUseXLimit(true);
            this.contentFormat.SetUseXLimit(true);
            this._multiline = true;
            this.SetWasChanged(true);
            this.SetNeedRecalc(true);
        }
        else if (bMultiline === false) {
            this.content.SetUseXLimit(false);
            this.contentFormat.SetUseXLimit(false);
            this._multiline = false;
            this.SetWasChanged(true);
            this.SetNeedRecalc(true);
        }
    };
    CTextField.prototype.IsMultiline = function() {
        return this._multiline;
    };
    CTextField.prototype.SetPassword = function(bPassword) {
        if (bPassword === true && this.fileSelect != true) {
            this._password = true;
        }
        else if (bPassword === false) {
            this._password = false;
        }
    };
    CTextField.prototype.IsPassword = function() {
        return this._password;
    };
    CTextField.prototype.SetRichText = function(bRichText) {
        this._richText = bRichText;
    };
    CTextField.prototype.IsRichText = function() {
        return this._richText;
    };
	CTextField.prototype.SetValue = function(sValue) {
		if (this.IsWidget()) {
            let oDoc        = this.GetDocument();
            let isOnOpen    = oDoc.Viewer.IsOpenFormsInProgress;

			if (isOnOpen != true)
				this.SetWasChanged(true);
			
			if (isOnOpen == true && !this.GetParent())
				this.SetApiValue(sValue);
			
			this.UpdateDisplayValue(sValue);
		}
		else {
			this.SetApiValue(sValue);
		}
	};
	CTextField.prototype.UpdateDisplayValue = function(displayValue) {
        let oDoc        = this.GetDocument();
        let isOnOpen    = oDoc.Viewer.IsOpenFormsInProgress;
        let _t          = this;

        oDoc.StartNoHistoryMode();

        AscFonts.FontPickerByCharacter.getFontsByString(displayValue);
        if (!oDoc.checkFieldFont(this, function() {
            _t.UpdateDisplayValue(displayValue);
        })) {
            return;
        }

        if (isOnOpen == false && this.GetType() == AscPDF.FIELD_TYPES.text && typeof(displayValue) == "string") {
            let aChars      = displayValue.codePointsArray();
            let nCharsCount = AscWord.GraphemesCounter.GetCount(aChars, this.content.GetCalculatedTextPr());
            let nCharLimit  = this.GetCharLimit();

            if (0 !== nCharLimit && nCharsCount > nCharLimit)
                aChars.length = nCharLimit;
            
            displayValue = String.fromCharCode.apply(null, aChars);
        }

        if (displayValue === this._displayValue && this._useDisplayValue == true)
			return;
		
		this._displayValue      = displayValue;
		this._useDisplayValue   = true;
		
        if (isOnOpen == true) {
            if (_t._displayValue !== displayValue)
				return;
			
			_t.content.replaceAllText(displayValue);
			_t.SetNeedRecalc(true);
        }
        else {
            if (_t._displayValue !== displayValue)
                return;
            
            _t.content.replaceAllText(displayValue);
            _t.SetNeedRecalc(true);
            _t.content.MoveCursorToStartPos();
        }
		
        oDoc.EndNoHistoryMode();
	};
    CTextField.prototype.GetCalcOrderIndex = function() {
        return this.GetDocument().GetCalculateInfo().ids.indexOf(this.GetApIdx());
    };
    CTextField.prototype.SetCalcOrderIndex = function(nIdx) {
        let oCalcInfo   = this.GetDocument().GetCalculateInfo();
        let oWidget     = null;
        if (this.IsWidget() || this.IsAllKidsWidgets()) {
            oWidget = this.GetKid(0) || this;
        }
        let oCalcTrigget = oWidget.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Calculate);
        if (oCalcTrigget == null || nIdx < 0)
            return false;

        let nCurIdx = oCalcInfo.ids.indexOf(this.GetApIdx());
        if (nCurIdx == nIdx)
            return true;

        oCalcInfo.ids.splice(nCurIdx, 1);
        if (nIdx > oCalcInfo.ids.length)
            oCalcInfo.ids.push(this.GetApIdx());
        else
            oCalcInfo.ids.splice(nIdx, 0, this.GetApIdx());

        return true;
    };

    // /**
	//  * Sets the value to childs fields.
	//  * @memberof CTextField
	//  * @typeofeditors ["PDF"]
	//  */
    // CTextField.prototype.SetValueToKids = function(sValue) {
    //     let oField, sName;
    //     let aDoneFields = [];
    //     for (let i = 0; i < this._kids.length; i++) {
    //         oField = this._kids[i];
    //         sName = oField.GetPartialName();

    //         if (oField.IsWidget()) {
    //             if (aDoneFields.includes(sName) == false) {
    //                 aDoneFields.push(oField.GetFullName());
    //                 oField.SetValue(sValue);
    //                 oField.Commit();
    //             }
    //         }
    //         else
    //             oField.SetValueToKids(sValue);
    //     }
    // };
	
	/**
	 * Gets the value of current form (can be not commited).
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 * @returns {string | Array} - can be array of rich value
	 */
	CTextField.prototype.GetValue = function() {
		return (this._useDisplayValue ? this._displayValue : this.content.getAllText());
	};
	CTextField.prototype.OnContentChange = function() {
		this._useDisplayValue = false;
        this.SetNeedRecalc(true);
	};
        
    CTextField.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        let oDoc = this.GetDocument();

        this.Recalculate();
        this.DrawBackground(oGraphicsPDF);
                
        let oContentToDraw = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format) && this.IsNeedDrawHighlight() ? this.contentFormat : this.content;
        this.curContent = oContentToDraw; // запоминаем текущий контент

        if (this.IsMultiline() == true) {
            oContentToDraw.ResetShiftView();
            oContentToDraw.ShiftView(this._curShiftView.x, this._curShiftView.y);
        }

        if (this._bAutoShiftContentView && oDoc.activeForm == this)
            this.CheckFormViewWindow();

        oGraphicsWord.AddClipRect(this.contentClipRect.X, this.contentClipRect.Y, this.contentClipRect.W, this.contentClipRect.H);
        oContentToDraw.Draw(0, oGraphicsWord);

        oGraphicsWord.RemoveLastClip();
        this.DrawBorders(oGraphicsPDF, oGraphicsWord);
        // redraw target cursor if field is selected
        if (oDoc.activeForm == this && oContentToDraw.IsSelectionUse() == false && this.IsCanEditText())
            oContentToDraw.RecalculateCurPos();
    };
    CTextField.prototype.DrawDateMarker = function(oCtx) {
        if (this.IsHidden())
            return;

        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let nScale      = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom * oViewer.getDrawingPageScale(nPage);
        let aOrigRect   = this.GetOrigRect();

        let xCenter = oViewer.width >> 1;
        if (oViewer.documentWidth > oViewer.width)
		{
			xCenter = (oViewer.documentWidth >> 1) - (oViewer.scrollX) >> 0;
		}
		let yPos    = oViewer.scrollY >> 0;
        let page    = oViewer.drawingPages[nPage];
        let w       = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let h       = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
        let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

        let isLandscape = oViewer.isLandscapePage(nPage);
        if (isLandscape) {
            indLeft = indLeft + (w - h) / 2;
        }
        
        let X       = aOrigRect[0] * nScale + indLeft;
        let Y       = aOrigRect[1] * nScale + indTop;
        let nWidth  = (aOrigRect[2] - aOrigRect[0]) * nScale;
        let nHeight = (aOrigRect[3] - aOrigRect[1]) * nScale;
        
        let oMargins = this.GetBordersWidth();
        
        let nMarkWidth  = 18;
        let nMarkX      = (X + nWidth) - oMargins.left * nScale - nMarkWidth;
        let nMarkHeight = nHeight - 2 * oMargins.top * nScale;
        let nMarkY      = Y + oMargins.top * nScale;

        // marker rect
        oCtx.setLineDash([]);
        oCtx.beginPath();
        oCtx.globalAlpha = 1;
        oCtx.fillStyle = "rgb(240, 240, 240)";
        oCtx.fillRect(nMarkX, nMarkY, nMarkWidth, nMarkHeight);

        // marker border right part
        oCtx.beginPath();
        oCtx.strokeStyle = "rgb(100, 100, 100)";
        oCtx.lineWidth = 1;
        oCtx.moveTo(nMarkX, nMarkY + nMarkHeight);
        oCtx.lineTo(nMarkX + nMarkWidth, nMarkY + nMarkHeight);
        oCtx.lineTo(nMarkX + nMarkWidth, nMarkY);
        oCtx.stroke();

        // marker border left part
        oCtx.beginPath();
        oCtx.strokeStyle = "rgb(255, 255, 255)";
        oCtx.moveTo(nMarkX, nMarkY + nMarkHeight);
        oCtx.lineTo(nMarkX, nMarkY);
        oCtx.lineTo(nMarkX + nMarkWidth, nMarkY);
        oCtx.stroke();

        // marker icon
        let nIconW = 5 * 1.5;
        let nIconH = 3 * 1.5;
        let nStartIconX = nMarkX + nMarkWidth/2 - (nIconW)/2;
        let nStartIconY = nMarkY + nMarkHeight/2 - (nIconH)/2;

        oCtx.beginPath();
        oCtx.fillStyle = "rgb(0, 0, 0)";
        
        oCtx.moveTo(nStartIconX, nStartIconY);
        oCtx.lineTo(nStartIconX + nIconW, nStartIconY);
        oCtx.lineTo(nStartIconX + nIconW/2, nStartIconY + nIconH);
        oCtx.lineTo(nStartIconX, nStartIconY);
        oCtx.fill();

        this._markRect = {
            x1: (nMarkX - indLeft) / nScale,
            y1: (nMarkY - indTop) / nScale,
            x2: ((nMarkX - indLeft) + (nMarkWidth)) / nScale,
            y2: ((nMarkY - indTop) + (nMarkHeight)) / nScale
        }
    };
    CTextField.prototype.IsDateFormat = function() {
        let oFormatTrigger      = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format);
        let oActionRunScript    = oFormatTrigger ? oFormatTrigger.GetActions()[0] : null;
        if (oActionRunScript && oActionRunScript.script.startsWith('AFDate_Format')) {
            return true;
        }

        return false;
    };
    CTextField.prototype.GetDateFormat = function() {
        let oFormatTrigger      = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format);
        let oActionRunScript    = oFormatTrigger ? oFormatTrigger.GetActions()[0] : null;
        if (oActionRunScript && oActionRunScript.script.startsWith('AFDate_Format')) {
            const regex = /(AFDate_Format|AFDate_FormatEx)\(["']([^"']+)["']\)/;
            const match = oActionRunScript.script.match(regex);

            if (match) {
                return match[2];
            }
            
            return "mmmm d, yyyy";
        }

        return "";
    }
    CTextField.prototype.ProcessAutoFitContent = function(oContent) {
        let oPara       = oContent.GetElement(0);
        let oRun        = oPara.GetElement(0);

        let oTextPr = oRun.Get_CompiledPr(true);
        let oBounds = this.getFormRelRect();

        g_oTextMeasurer.SetTextPr(oTextPr, null);
        g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);

        var nTextHeight = g_oTextMeasurer.GetHeight();
        var nMaxWidth   = oPara.RecalculateMinMaxContentWidth(false).Max;
        var nFontSize   = oTextPr.FontSize;

        if (nMaxWidth < 0.001 || nTextHeight < 0.001 || oBounds.W < 0.001 || oBounds.H < 0.001)
    	    return nTextHeight;

        var nNewFontSize = nFontSize;

        if (this.IsMultiline() == true) {
            const nFontStep = 0.05;
            oPara.Recalculate_Page(0);
            let oContentBounds = oContent.GetContentBounds(0);
            if (oContentBounds.Bottom - oContentBounds.Top > oBounds.H)
            {
                nNewFontSize = AscCommon.CorrectFontSize(nFontSize, true);
                while (nNewFontSize > 1)
                {
                    oRun.SetFontSize(nNewFontSize);
                    oPara.Recalculate_Page(0);

                    oContentBounds = oContent.GetContentBounds(0);
                    if (oContentBounds.Bottom - oContentBounds.Top < oBounds.H)
                        break;

                    nNewFontSize -= nFontStep;
                }
            }
            else
            {
                let nMaxFontSize = 12;

                while (nNewFontSize <= nMaxFontSize)
                {
                    oRun.SetFontSize(nNewFontSize);
                    
                    oPara.Recalculate_Page(0);

                    let oContentBounds = oContent.GetContentBounds(0);
                    if (oContentBounds.Bottom - oContentBounds.Top > oBounds.H)
                    {
                        nNewFontSize -= nFontStep;
                        oRun.SetFontSize(nNewFontSize);
                        oPara.Recalculate_Page(0);
                        break;
                    }

                    nNewFontSize += nFontStep;
                }

                nNewFontSize = Math.min(nNewFontSize, nMaxFontSize);
            }

        }
        else {
            nNewFontSize = (Math.min(nFontSize * oBounds.H / nTextHeight * 0.9, 100, nFontSize * oBounds.W / nMaxWidth) * 100 >> 0) / 100;
            oRun.SetFontSize(nNewFontSize);
            oPara.Recalculate_Page(0);
        }

        oTextPr.FontSize    = nNewFontSize;
        oTextPr.FontSizeCS  = nNewFontSize;

        this.AddToRedraw();
    };
    CTextField.prototype.GetTextHeight = function(oContent) {
        let oPara   = oContent.GetElement(0);
        let oRun    = oPara.GetElement(0);
        let oTextPr = oRun.Get_CompiledPr(true);

        g_oTextMeasurer.SetTextPr(oTextPr, null);
        g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);

        return g_oTextMeasurer.GetHeight();
    };
    CTextField.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        if (this.IsNeedCheckAlign()) {
            this.CheckAlignInternal();
        }

        this.RecalcMeasureContent();
        
        if (this.GetTextSize() == 0) {
            if (null == this.getFormRelRect()) {
                this.CalculateContentClipRect();
            }
            this.ProcessAutoFitContent(this.content);
            this.ProcessAutoFitContent(this.contentFormat);
        }

        if (false == this.RecalculateContentRect()) {
            this.contentFormat.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this.content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
        }

        this.SetNeedRecalc(false);
    };
    CTextField.prototype.RecalculateContentRect = function() {
        let aOrigRect = this.GetOrigRect();

        let X       = aOrigRect[0];
        let Y       = aOrigRect[1];
        let nWidth  = ((aOrigRect[2]) - (aOrigRect[0]));
        let nHeight = ((aOrigRect[3]) - (aOrigRect[1]));

        let oMargins = this.GetMarginsFromBorders();
        
        let contentX = (this.IsComb() ? (X + oMargins.left) : (X + 2 * oMargins.left)) * g_dKoef_pt_to_mm;
        let contentY = (Y + (this.IsMultiline() ? (2.5 * oMargins.top) : (2 * oMargins.top))) * g_dKoef_pt_to_mm;
        let contentXLimit = (this.IsComb() ? (X + nWidth - oMargins.left) : (X + nWidth - 2 * oMargins.left)) * g_dKoef_pt_to_mm;
        
        if ((this.borderStyle == "solid" || this.borderStyle == "dashed") && 
        this._comb == true && this._charLimit > 1) {
            contentX = (X) * g_dKoef_pt_to_mm;
            contentXLimit = (X + nWidth) * g_dKoef_pt_to_mm;
        }
        
        let contentYFormat = contentY;
        if (this.IsMultiline() == false) {
            let nContentH       = this.GetTextHeight(this.content);
            let nContentHFormat = this.GetTextHeight(this.contentFormat);
            
            contentY        = Y * g_dKoef_pt_to_mm + (nHeight * g_dKoef_pt_to_mm - nContentH) / 2;
            contentYFormat  = Y * g_dKoef_pt_to_mm + (nHeight * g_dKoef_pt_to_mm - nContentHFormat) / 2;
        }

        if (contentX != this.content.X || contentY != this.content.Y ||
        contentXLimit != this.content.XLimit || contentYFormat != this.contentFormat.YFormat) {
            this.content.X      = this.contentFormat.X = contentX;
            this.content.Y      = contentY;
            this.contentFormat.Y= contentYFormat;
            this.content.XLimit = this.contentFormat.XLimit = contentXLimit;
            this.content.YLimit = this.contentFormat.YLimit = 20000;
            
            this.CalculateContentClipRect();
            this.content.Recalculate_Page(0, true);
            this.contentFormat.Recalculate_Page(0, true);

            return true;
        }

        return false;
    };
    CTextField.prototype.CalculateContentClipRect = function() {
        if (!this.content)
            return;

        let aRect       = this.GetOrigRect();
        let X           = aRect[0];
        let Y           = aRect[1];
        let nWidth      = aRect[2] - aRect[0];
        let nHeight     = aRect[3] - aRect[1];
        let oMargins    = this.GetMarginsFromBorders();

        let contentX = (this.IsComb() ? (X + oMargins.left) : (X + 2 * oMargins.left)) * g_dKoef_pt_to_mm;
        let contentXLimit = (this.IsComb() ? (X + nWidth - oMargins.left) : (X + nWidth - 2 * oMargins.left)) * g_dKoef_pt_to_mm;

        this.contentClipRect = {
            X: contentX,
            Y: (Y + (this.IsMultiline() ? (2.5 * oMargins.top) : (oMargins.top))) * g_dKoef_pt_to_mm,
            W: contentXLimit - contentX,
            H: (nHeight - (this.IsMultiline() ? 2.5 * oMargins.top : oMargins.top) - oMargins.bottom) * g_dKoef_pt_to_mm,
            Page: this.GetPage()
        }
    };
    CTextField.prototype.onMouseDown = function(x, y, e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        let isInFocus   = oDoc.activeForm === this;
        let isInForm    = this.IsInForm();

        oDoc.activeForm = this;

        function callbackAfterFocus(x, y, e) {
            this.SetInForm(true);
            oDoc.SetLocalHistory();
            if (false == e.ShiftKey) {
                oDoc.SelectionSetStart(x, y, e);
            }
            else {
                this.content.StartSelectionFromCurPos();
                oDoc.SelectionSetEnd(x, y, e);
            }
            
            let pageObject = oViewer.getPageByCoords(x, y);

            if (this.IsDateFormat() && this.IsInForm() && pageObject.x >= this._markRect.x1 && pageObject.x <= this._markRect.x2 && pageObject.y >= this._markRect.y1 && pageObject.y <= this._markRect.y2) {
                editor.sendEvent("asc_onShowPDFFormsActions", this, x, y);
                this.content.MoveCursorToStartPos();
            }

            if (this.IsDoNotScroll() == false && this.IsMultiline())
                this.UpdateScroll(true);

            this.SetDrawHighlight(false);
            if (this.IsNeedDrawFromStream() == true) {
                this.SetDrawFromStream(false);
                this.AddToRedraw();
            }
            else if (this.curContent === this.contentFormat || false == isInForm) {
                this.AddToRedraw();
            }
        }

        let oOnFocus = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);
        // вызываем выставление курсора после onFocus. Если уже в фокусе, тогда сразу.
        if (false == isInFocus && oOnFocus && oOnFocus.Actions.length > 0)
            oActionsQueue.callbackAfterFocus = callbackAfterFocus.bind(this, x, y, e);
        else
            callbackAfterFocus.bind(this, x, y, e)();

        if (isInFocus) {
            this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseDown);
        }
        else {
            this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseDown, AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);
        }
    };
    CTextField.prototype.onMouseUp = function(x, y, e) {
        let oDoc    = this.GetDocument();
        let oDrDoc  = oDoc.GetDrawingDocument();

        if (global_mouseEvent.ClickCount == 2) {
            this.SelectAllText();
            oDrDoc.TargetEnd();
        }

        if (this.content.IsSelectionEmpty()) {
            this.content.RemoveSelection();
            if (false == oDoc.GetActionsQueue().IsInProgress()) {
                oDrDoc.TargetStart(true);
            }
        }

        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseUp);
        oDoc.Viewer.onUpdateOverlay();
    };
    CTextField.prototype.ScrollVertical = function(scrollY, maxYscroll) {
        this._bAutoShiftContentView = false;

        let nScrollCoeff                = scrollY / maxYscroll;
        this._curShiftView.y            = -nScrollCoeff * maxYscroll;
        this._scrollInfo.scrollCoeff    = nScrollCoeff;
        this.AddToRedraw();
    };
    CTextField.prototype.GetScrollInfo = function() {
        return this._scrollInfo;
    };
    CTextField.prototype.SetScrollInfo = function(oInfo) {
        this._scrollInfo = oInfo;
    };
    CTextField.prototype.UpdateScroll = function(bShow) {
        if (this.IsMultiline() == false)
            return;
        
        let oContentBounds  = this.content.GetContentBounds(0);
        let oContentRect    = this.getFormRelRect();
        let aOrigRect       = this.GetOrigRect();

        let nContentH   = oContentBounds.Bottom - oContentBounds.Top;
        let oScrollInfo = this.GetScrollInfo();
        if ((bShow == false || nContentH < oContentRect.H || this.IsDoNotScroll())) {
            if (oScrollInfo) {
                oScrollInfo.docElem.style.display = "none";
            }
            
            return;
        }

        let oDoc        = this.GetDocument();
        let nPage       = this.GetPage();
        let oTransform  = oDoc.pagesTransform[nPage].invert;
        let oViewer     = oDoc.Viewer;
        let isLandscape = oViewer.isLandscapePage(nPage);
        let nRotAngle   = oViewer.getPageRotate(nPage);

        let oGlobalCoords1  = oTransform.TransformPoint(aOrigRect[0], aOrigRect[1]);
        let oGlobalCoords2  = oTransform.TransformPoint(aOrigRect[2], aOrigRect[3]);

        let nLeftPos;
        let nTopPos;

        let bInvertScroll = false;
        switch (nRotAngle) {
            case 0:
                nLeftPos    = Math.round(oGlobalCoords2.x);
                nTopPos     = Math.round(oGlobalCoords1.y);
                break
            case 90:
                nLeftPos    = Math.round(oGlobalCoords2.x);
                nTopPos     = Math.round(oGlobalCoords2.y);
                bInvertScroll = true;
                break;
            case 180:
                nLeftPos    = Math.round(oGlobalCoords2.x) - 14;
                nTopPos     = Math.round(oGlobalCoords2.y);
                bInvertScroll = true;
                break;
            case 270:
                nLeftPos    = Math.round(oGlobalCoords1.x);
                nTopPos     = Math.round(oGlobalCoords2.y) - 14;
                break;
        }
        
        if (oContentBounds.Bottom - oContentBounds.Top > oContentRect.H) {
            let oScrollDocElm;
            if (oScrollInfo == null) {
                oViewer.scrollCount++;
                oScrollDocElm = document.createElement('div');
                document.getElementById('editor_sdk').appendChild(oScrollDocElm);
                oScrollDocElm.id = "formScroll_" + oViewer.scrollCount;
            }
            else {
                oScrollDocElm = oScrollInfo.docElem;
            }
            
            oScrollDocElm.style.top         = nTopPos + 'px';
            oScrollDocElm.style.left        = nLeftPos + 'px';
            oScrollDocElm.style.position    = "absolute";
            oScrollDocElm.style.display     = "block";
			oScrollDocElm.style.width       = isLandscape ? Math.round(Math.abs(oGlobalCoords2.x - oGlobalCoords1.x)) + "px" : "14px";
			oScrollDocElm.style.height      = isLandscape ? "14px" : Math.round(Math.abs(oGlobalCoords2.y - oGlobalCoords1.y)) + "px";
            oScrollDocElm.style.zIndex      = 0;

            let nMaxShift = oContentRect.H - nContentH;

            let oScrollSettings = Asc.editor.WordControl.CreateScrollSettings();
            oScrollSettings.isHorizontalScroll  = isLandscape;
		    oScrollSettings.isVerticalScroll    = !isLandscape;
            oScrollSettings.contentW            = isLandscape ? Math.abs(nMaxShift) : 0;
		    oScrollSettings.contentH            = isLandscape ? 0 : Math.abs(nMaxShift);
            oScrollSettings.screenH             = 0;
            oScrollSettings.screenW             = 0;
            oScrollSettings.scrollerMinHeight   = 5;
            oScrollSettings.scrollerMinWidth    = 5;
            
            let nScrollCoeff = this._curShiftView.y / nMaxShift;
            
            let oPara = this.content.GetElement(0);
            let oCurParaHeight = oPara.Lines[0].Bottom - oPara.Lines[0].Top;

            if (isLandscape) {
                oScrollSettings.hscrollStep = oCurParaHeight;
            }
            else {
                oScrollSettings.vscrollStep = oCurParaHeight;
            }
            
            let oScroll;
            if (oScrollInfo == null) {
                oScroll = new AscCommon.ScrollObject(oScrollDocElm.id, oScrollSettings);
            }
            else {
                oScroll = oScrollInfo.scroll;
            }

            let oThis = this;
            if (isLandscape) {
                oScroll.bind("scrollhorizontal", function (evt) {
                    if (false == bInvertScroll) {
                        oThis.ScrollVertical(evt.scrollD, evt.maxScrollX);
                    }
                    else {
                        oThis.ScrollVertical(evt.maxScrollX - evt.scrollD, evt.maxScrollX);
                    }
				});
                
                oScroll.scrollHCurrentX = false == bInvertScroll ? oScroll.maxScrollX * nScrollCoeff : oScroll.maxScrollX - (oScroll.maxScrollX * nScrollCoeff);
            }
            else {
                oScroll.bind("scrollvertical", function(evt) {
                    if (false == bInvertScroll) {
                        oThis.ScrollVertical(evt.scrollD, evt.maxScrollY);
                    }
                    else {
                        oThis.ScrollVertical(evt.maxScrollY - evt.scrollD, evt.maxScrollY);
                    }
                });

                oScroll.scrollVCurrentY = false == bInvertScroll ? oScroll.maxScrollY * nScrollCoeff : oScroll.maxScrollY - (oScroll.maxScrollY * nScrollCoeff);
            }
            
            if (oScrollInfo == null) {
                this.SetScrollInfo({
                    scroll:         oScroll,
                    docElem:        oScrollDocElm,
                    baseYPos:       parseInt(oScrollDocElm.style.top),
                    oldZoom:        oViewer.zoom,
                    scrollCoeff:    nScrollCoeff, // проскроленная часть
                    rot:            nRotAngle 
                });
            }

            oScroll.Repos(oScrollSettings, false, undefined, undefined, true);
        }
    };

    CTextField.prototype.SelectionSetStart = function(X, Y, e) {
        this.content.Selection_SetStart(X, Y, 0, e);
        this.content.RecalculateCurPos();
    };
    CTextField.prototype.SelectionSetEnd = function(X, Y, e) {
        this.content.Selection_SetEnd(X, Y, 0, e);
    };
    CTextField.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey) {
        this.content.MoveCursorLeft(isShiftKey, isCtrlKey);
        let oCurPos = this.content.RecalculateCurPos();
            
        let oFieldBounds = this.getFormRelRect();
        let nCursorH = g_oTextMeasurer.GetHeight();
        if ((oCurPos.X < oFieldBounds.X || oCurPos.Y - nCursorH * 0.75 < oFieldBounds.Y) && this.IsDoNotScroll() != true)
        {
            this.AddToRedraw();
            this.UpdateScroll(true);
        }
        
        this._bAutoShiftContentView = true && this.IsDoNotScroll() == false;
    };
    CTextField.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey) {
        this.content.MoveCursorRight(isShiftKey, isCtrlKey);
        let oCurPos = this.content.RecalculateCurPos();
            
        let oFieldBounds = this.getFormRelRect();
        if ((oCurPos.X > oFieldBounds.X + oFieldBounds.W || oCurPos.Y > oFieldBounds.Y + oFieldBounds.H) && this.IsDoNotScroll() != true)
        {
            this.AddToRedraw();
            this.UpdateScroll(true);
        }
        
        this._bAutoShiftContentView = true && this.IsDoNotScroll() == false;
    };
    CTextField.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        this.content.MoveCursorDown(isShiftKey, isCtrlKey);
        let oCurPos = this.content.RecalculateCurPos();
            
        let oFieldBounds = this.getFormRelRect();
        if (oCurPos.Y > oFieldBounds.Y + oFieldBounds.H && this.IsDoNotScroll() != true)
        {
            this.AddToRedraw();
            this.UpdateScroll(true);
        }
        
        this._bAutoShiftContentView = true && this.IsDoNotScroll() == false;
    };
    CTextField.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        this.content.MoveCursorUp(isShiftKey, isCtrlKey);
        let oCurPos = this.content.RecalculateCurPos();

        let oFieldBounds = this.getFormRelRect();
        let nCursorH = g_oTextMeasurer.GetHeight();
        if (oCurPos.Y - nCursorH * 0.75 < oFieldBounds.Y && this.IsDoNotScroll() != true)
        {
            this.AddToRedraw();
            this.UpdateScroll(true);
        }
        
        this._bAutoShiftContentView = true && this.IsDoNotScroll() == false;
    };
    CTextField.prototype.getRemainCharCount = function(removeCount) {
		if (0 === this._charLimit)
			return -1;
		
		let charCount = 0;
		this.content.CheckRunContent(function(oRun) {
			var nCurPos = oRun.Content.length;
			for (var nPos = 0; nPos < nCurPos; ++nPos) {
				if (para_Text === oRun.Content[nPos].Type || para_Space === oRun.Content[nPos].Type || para_Tab === oRun.Content[nPos].Type)
					++charCount;
			}
		});
		
		if (removeCount > charCount)
			removeCount = charCount;
		
		return Math.max(0, this._charLimit - (charCount - removeCount));
	};
	CTextField.prototype.EnterText = function(aChars) {
		let selectedCount = this.content.GetSelectedText(true, {NewLine: true}).length;
		let maxToAdd      = this.getRemainCharCount(selectedCount);
		
        let nCharsCount = AscWord.GraphemesCounter.GetCount(aChars, this.content.GetCalculatedTextPr());
		if (-1 !== maxToAdd && nCharsCount > maxToAdd)
			aChars.length = maxToAdd;
		
		if (!this.DoKeystrokeAction(aChars))
			return false;
		
		let doc = this.GetDocument();
		aChars = AscWord.CTextFormFormat.prototype.GetBuffer(doc.event["change"]);
		if (0 === nCharsCount)
			return false;
		
		if (!this.content.EnterText(aChars))
			return false;
		
		this.SetNeedRecalc(true);
		this.SetNeedCommit(true); // флаг что значение будет применено к остальным формам с таким именем
		this._bAutoShiftContentView = true && this.IsDoNotScroll() == false;
		
		if (this.IsDoNotScroll()) {
			let isOutOfForm = this.IsTextOutOfForm(this.content);
			if ((this.IsMultiline() && isOutOfForm.ver) || (isOutOfForm.hor && this.IsMultiline() == false)) {
				AscCommon.History.ForbidUnionPoint();
				AscCommon.History.Undo();
				AscCommon.History.Clear_Redo();
			}
			this.AddToRedraw();
		}
		
		return true;
	};
	CTextField.prototype.CorrectEnterText = function(oldValue, newValue) {
		let maxToAdd = this.getRemainCharCount(oldValue.length);
		
        let nCharsCount = AscWord.GraphemesCounter.GetCount(newValue, this.content.GetCalculatedTextPr());
		if (-1 !== maxToAdd && nCharsCount > maxToAdd)
			newValue.length = maxToAdd;
		
		if (!this.DoKeystrokeAction(newValue))
			return false;
		
		let doc = this.GetDocument();
		newValue = AscWord.CTextFormFormat.prototype.GetBuffer(doc.event["change"]);
		
		if (!this.content.CorrectEnterText(oldValue, newValue, function(run, inRunPos, codePoint){return true;}))
			return false;
		
		this.SetNeedRecalc(true);
		this.SetNeedCommit(true); // флаг что значение будет применено к остальным формам с таким именем
		this._bAutoShiftContentView = true && this.IsDoNotScroll() == false;
		
		if (this.IsDoNotScroll()) {
			let isOutOfForm = this.IsTextOutOfForm(this.content);
			if ((this.IsMultiline() && isOutOfForm.ver) || (isOutOfForm.hor && this.IsMultiline() == false)) {
				AscCommon.History.ForbidUnionPoint();
				AscCommon.History.Undo();
				AscCommon.History.Clear_Redo();
			}
			this.AddToRedraw();
		}
		return true;
	};
    CTextField.prototype.CheckAlignInternal = function() {
        this.SetNeedCheckAlign(false);

        // если выравнивание по центру или справа, то оно должно переключаться на left если ширина контента выходит за пределы формы
        // вызывается на момент коммита формы
        if ([AscPDF.ALIGN_TYPE.center, AscPDF.ALIGN_TYPE.right].includes(this.GetAlign())) {

            if (this.IsTextOutOfForm(this.content).hor) {
                if (this.content.GetAlign() != AscPDF.ALIGN_TYPE.left) {
                    this.content.SetAlign(AscPDF.ALIGN_TYPE.left);
                }
            }
            else if (this.content.GetAlign() != this.GetAlign()) {
                this.content.SetAlign(this.GetAlign());
            }

            if (this.IsTextOutOfForm(this.contentFormat).hor) {
                if (this.contentFormat.GetAlign() != AscPDF.ALIGN_TYPE.left) {
                    this.contentFormat.SetAlign(AscPDF.ALIGN_TYPE.left);
                }
            }
            else if (this.contentFormat.GetAlign() != this.GetAlign()) {
                this.contentFormat.SetAlign(this.GetAlign());
            }
        }
    };
    CTextField.prototype.SetNeedCheckAlign = function(bCheck) {
        this._needCheckAlign = bCheck;
    };
    CTextField.prototype.IsNeedCheckAlign = function() {
        return this._needCheckAlign;
    };
    CTextField.prototype.InsertChars = function(aChars) {
		this.content.EnterText(aChars);
	};
	CTextField.prototype.canBeginCompositeInput = function() {
		return true;
	};
    CTextField.prototype.SelectAllText = function() {
        this.content.SelectAll();
    };
	CTextField.prototype.beforeCompositeInput = function() {
		this.DoKeystrokeAction();
		this.removeBeforePaste();
	};
    /**
	 * Checks is text in form is out of form bounds.
     * Note: in vertical case one line always be valid even if form is very short.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
     * @returns {object} - {hor: {boolean}, ver: {boolean}}
	 */
    CTextField.prototype.IsTextOutOfForm = function(oContent) {
        if (null == this.getFormRelRect())
            this.Recalculate();
        else
            oContent.GetElement(0).Recalculate_Page(0);
        
        let oPageBounds = oContent.GetContentBounds(0);
        let oFormBounds = this.getFormRelRect();
        let nContentW   = oContent.GetElement(0).GetContentWidthInRange();
        
        let oResult = {
            hor: false,
            ver: false
        }

        if (nContentW > oFormBounds.W) {
            oResult.hor = true;
        }

        if (this.IsMultiline && this.IsMultiline() && oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
            oResult.ver = true;
        }

        return oResult;
    };

    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.Commit = function() {
        let oDoc        = this.GetDocument();
        let aFields     = this.GetDocument().GetAllWidgets(this.GetFullName());

        oDoc.StartNoHistoryMode();
        if (this.DoFormatAction() == false) {
            this.UndoNotAppliedChanges();
            if (this.IsChanged() == false)
                this.SetDrawFromStream(true);

            oDoc.EndNoHistoryMode();
            return;
        }
        oDoc.EndNoHistoryMode();
        
        if (this.GetApiValue() != this.GetValue()) {
            AscCommon.History.Add(new CChangesPDFFormValue(this, this.GetApiValue(), this.GetValue()));
            this.RevertContentView();
            this.SetApiValue(this.GetValue());
        }

        oDoc.StartNoHistoryMode();

        if (aFields.length == 1)
            this.SetNeedCommit(false);

        if (oDoc.event["rc"] == false) {
            this.needValidate = true;
            return;
        }

		let sValue = this.GetValue();
        this.UpdateDisplayValue(sValue);

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i].IsChanged() == false)
                aFields[i].SetWasChanged(true); // фиксируем, что форма была изменена

            if (aFields[i].HasShiftView()) {
                aFields[i].content.MoveCursorToStartPos();

                if (aFields[i] == this) {
                    aFields[i].AddToRedraw();
                    continue;
                }
            }
                
            if (aFields[i] == this)
                continue;

            aFields[i].UpdateDisplayValue(sValue);
            aFields[i].SetNeedRecalc(true);
        }

        let sFormatValue = this.contentFormat.getAllText();
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] == this)
                continue;

            aFields[i].contentFormat.replaceAllText(sFormatValue);
            aFields[i].SetNeedRecalc(true);
        }

        // когда выравнивание посередине или справа, то после того
        // как ширина контента будет больше чем размер формы, выравнивание становится слева, пока текста вновь не станет меньше чем размер формы
        aFields.forEach(function(field) {
            field.SetNeedCheckAlign(true);
        });

        this.SetNeedCommit(false);
        this.needValidate = true;

        oDoc.EndNoHistoryMode();
    };
	CTextField.prototype.SetAlign = function(nAlignType) {
        this._alignment = nAlignType;
		this.content.SetAlign(nAlignType);
		if (this.contentFormat)
			this.contentFormat.SetAlign(nAlignType);
		
        this.SetWasChanged(true);
		this.SetNeedRecalc(true);
	};
	CTextField.prototype.GetAlign = function() {
		return this._alignment;
	};

    CTextField.prototype.DoFormatAction = function() {
        let oDoc                = this.GetDocument();
        let oFormatTrigger      = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format);
        let oActionRunScript    = oFormatTrigger ? oFormatTrigger.GetActions()[0] : null;
        
        let isCanFormat = AscCommon.History.UndoRedoInProgress != true ? this.DoKeystrokeAction(null, false, true) : true;
        if (!isCanFormat) {
            let oWarningInfo = oDoc.GetWarningInfo();
            if (!oWarningInfo) {
                oWarningInfo = {
                    "format": "",
                    "target": this
                };
            }
            editor.sendEvent("asc_onFormatErrorPdfForm", oWarningInfo);
            return false;
        }

        if (oActionRunScript) {
            this.SetNeedRecalc(true);
            oActionRunScript.RunScript();
        }

        return true;
    };
    CTextField.prototype.DoKeystrokeAction = function(aChars, nRemoveType, isOnCommit, isCtrlKey) {
		if (!aChars)
			aChars = [];
		else if (typeof (aChars) === "string")
			aChars = aChars.codePointsArray();

        let oKeystrokeTrigger = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke);
        let oActionRunScript = oKeystrokeTrigger ? oKeystrokeTrigger.GetActions()[0] : null;
        let oDoc = this.GetDocument();
        
        let isCanEnter = true;

        function GetSelectionRange(p)
        {
            let selectedText = p.GetSelectedText(undefined, {NewLine: true});
            let state = p.SaveSelectionState();

            let selStart = p.Get_ParaContentPos(p.IsSelectionUse(), p.GetSelectDirection() > 0);

            p.RemoveSelection();
            p.StartSelectionFromCurPos();
            p.SetSelectionContentPos(p.GetStartPos(), selStart);
            let preText = p.GetSelectedText(undefined, {NewLine: true});

            p.LoadSelectionState(state);

            return {nSelStart : preText.length, nSelEnd : preText.length + selectedText.length};
        }

        let sValue      = this.GetValue(true);
        let oSelRange   = GetSelectionRange(this.content.GetElement(0));
        let nSelStart   = oSelRange.nSelStart;
        let nSelEnd     = oSelRange.nSelEnd;
        
        // если нет селекта, тогда обрабатывает ctrl
        if (isCtrlKey && nSelStart == nSelEnd && nRemoveType !== undefined) {
            // удаляем текст до пробела
            if (nRemoveType == -1) { // backspace
                let nSpacePos = sValue.lastIndexOf(' ', nSelEnd - 2);
                nSelStart = nSpacePos != - 1 ? nSpacePos + 1 : 0;
            }
            else if (nRemoveType == 1) { // delete
                let nSpacePos = sValue.indexOf(' ', nSelEnd + 2);
                nSelEnd = nSpacePos != - 1 ? nSpacePos : sValue.length;
            }
        }
        
        let nCharsCount = this.content.GetElement(0).GetText({ParaEndToSpace: false}).length;
        if (nRemoveType && nSelStart == nSelEnd) {
            if (nRemoveType == -1 && nSelStart != 0) {
                nSelStart--;
            }
            else if (nRemoveType == 1 && nSelEnd != nCharsCount) {
                nSelEnd++;
            }
        }

        this.GetDocument().SetEvent({
            "target":   this.GetFormApi(),
            "value":    sValue,
            "change":   aChars.map(function(char) {
                return String.fromCharCode(char);
            }).join(""),
            "willCommit": !!isOnCommit,
            "selStart": nSelStart,
            "selEnd": nSelEnd
        });

        if (oActionRunScript) {
            oActionRunScript.RunScript();
            isCanEnter = oDoc.event["rc"];
        }

        return isCanEnter;
    };
    CTextField.prototype.DoValidateAction = function(value) {
        let oDoc = this.GetDocument();

        oDoc.SetEvent({
            "taget": this.GetFormApi(),
            "rc": true,
            "value": value
        });

        let oValidateTrigger = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Validate);
        let oValidateScript = oValidateTrigger ? oValidateTrigger.GetActions()[0] : null;

        if (oValidateScript == null)
            return true;

        oDoc.isOnValidate = true;
        oValidateScript.RunScript();
        let isValid = oDoc.event["rc"];
        oDoc.isOnValidate = false;

        if (isValid == false) {
            let oWarningInfo = oDoc.GetWarningInfo();
            if (!oWarningInfo) {
                oWarningInfo = {
                    "target": this
                };
            }
            if (oWarningInfo["greater"] != null || oWarningInfo["less"] != null)
                editor.sendEvent("asc_onValidateErrorPdfForm", oWarningInfo);
            
            return isValid;
        }
        
        return isValid;
    };

    /**
	 * Removes char in current position by direction.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
	CTextField.prototype.Remove = function(nDirection, isCtrlKey) {
		if (this.IsCanEditText() == false)
			return false;
		
		if (!this.DoKeystrokeAction(null, nDirection, false, isCtrlKey))
			return false;
		
		let oDoc = this.GetDocument();
		this.UpdateSelectionByEvent();
		
		if (this.content.IsSelectionUse())
			this.content.Remove(nDirection, true, false, false, isCtrlKey);

        // скрипт keystroke мог поменять change значение, поэтому
        this.InsertChars(AscWord.CTextFormFormat.prototype.GetBuffer(oDoc.event["change"].toString()));

        if (false == AscCommon.History.Is_LastPointEmpty()) {
            this._bAutoShiftContentView = true && this.IsDoNotScroll() == false;
            this.SetNeedRecalc(true);
            this.SetNeedCommit(true);
        }
    };
    CTextField.prototype.RecalcMeasureContent = function() {
        function setRecalc(oRun) {
            oRun.RecalcInfo.Measure = true;
            oRun.RecalcInfo.Recalc = true;
        }
        if ([AscPDF.ALIGN_TYPE.center, AscPDF.ALIGN_TYPE.right].includes(this.content.GetAlign())) {
            this.content.CheckRunContent(setRecalc);
        }
        if ([AscPDF.ALIGN_TYPE.center, AscPDF.ALIGN_TYPE.right].includes(this.contentFormat.GetAlign())) {
            this.contentFormat.CheckRunContent(setRecalc);
        }
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.SyncField = function() {
        let aFields = this.GetDocument().GetAllWidgets(this.GetFullName());
        
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this._alignment         = aFields[i]._alignment;
                this._charLimit         = aFields[i]._charLimit;
                this._comb              = aFields[i]._comb;
                this._doNotScroll       = aFields[i]._doNotScroll;
                this._doNotSpellCheckl  = aFields[i]._doNotSpellCheckl;
                this._fileSelect        = aFields[i]._fileSelect;
                this._multiline         = aFields[i]._multiline;
                this._password          = aFields[i]._password;
                this._richText          = aFields[i]._richText;
                this._richValue         = aFields[i]._richValue.slice();
                this._textFont          = aFields[i]._textFont;
                this._borderStyle       = aFields[i]._borderStyle;

                this._triggers = aFields[i]._triggers ? aFields[i]._triggers.Copy(this) : null;

                if (this._multiline)
                    this.content.SetUseXLimit(true);

                let oPara = this.content.GetElement(0);
                let oParaToCopy = aFields[i].content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();

                // format content
                oPara = this.contentFormat.GetElement(0);
                oParaToCopy = aFields[i].contentFormat.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                break;
            }
        }
    };
    CTextField.prototype.CheckFormViewWindow = function()
    {
        // размеры всего контента
        let oPageBounds = this.content.GetContentBounds(0);
        let oFormBounds = this.getFormRelRect();

        let oParagraph = this.content.GetElement(0);
        
        let nDx = 0, nDy = 0;

        if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
        {
            if (oPageBounds.Left > oFormBounds.X)
                nDx = -oPageBounds.Left + oFormBounds.X;
            else if (oPageBounds.Right < oFormBounds.X + oFormBounds.W)
                nDx = oFormBounds.X + oFormBounds.W - oPageBounds.Right;
        }
        else
        {
            nDx = -this.content.ShiftViewX;
        }

        if (Math.abs(nDx) > 0.001 || Math.abs(nDy))
        {
            this.content.ShiftView(nDx, nDy);
        }

        var oCursorPos  = oParagraph.GetCalculatedCurPosXY();
        var oLineBounds = oParagraph.GetLineBounds(oCursorPos.Internal.Line);
        var oLastLineBounds = oParagraph.GetLineBounds(oParagraph.GetLinesCount() - 1);

	    nDx = 0;
	    nDy = 0;

        var nCursorT = Math.min(oCursorPos.Y, oLineBounds.Top);
        var nCursorB = Math.max(oCursorPos.Y + oCursorPos.Height, oLineBounds.Bottom);
        var nCursorH = Math.max(0, nCursorB - nCursorT);

        if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
        {
            if (oCursorPos.X < oFormBounds.X)
                nDx = oFormBounds.X - oCursorPos.X;
            else if (oCursorPos.X > oFormBounds.X + oFormBounds.W)
                nDx = oFormBounds.X + oFormBounds.W - oCursorPos.X;
        }

        if (this.IsMultiline && this.IsMultiline()) {
            // если высота контента больше чем высота формы
            if (oParagraph.IsSelectionUse()) {
                if (oParagraph.GetSelectDirection() == 1) {
                    if (nCursorT + nCursorH - nCursorH/4 > oFormBounds.Y + oFormBounds.H)
                        nDy = oFormBounds.Y + oFormBounds.H - (nCursorT + nCursorH);
                }
                else {
                    if (nCursorT + nCursorH/4 < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                }
            }
            else {
                if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
                    if (oLastLineBounds.Bottom - Math.floor(((oFormBounds.Y + oFormBounds.H) * 1000)) / 1000 < 0)
                        nDy = oFormBounds.Y + oFormBounds.H - oLastLineBounds.Bottom;
                    else if (nCursorT + nCursorH/4 < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                    else if (nCursorT + nCursorH - nCursorH/4 > oFormBounds.Y + oFormBounds.H)
                        nDy = oFormBounds.Y + oFormBounds.H - (nCursorT + nCursorH);
                }
                else
                    nDy = -this.content.ShiftViewY;
            }
        }

        if (Math.abs(nDx) > 0.001 || Math.round(nDy) != 0)
        {
            this.content.ShiftView(nDx, nDy);
            this._curShiftView = {
                x: this.content.ShiftViewX,
                y: this.content.ShiftViewY
            }
        }
    };
    /**
	 * Calculates doc position by selection position.
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.CalcDocPos = function(nStart, nEnd) {
        let oPara = this.content.GetElement(0);

        var isStartDone = false;
		var isEndDone	= false;
		var StartPos    = null;
		var EndPos      = null;
		var nCharsCount  = 0;
		var DocPos, DocPosInRun;

		function callback(oRun)
		{
			var nRangePos = 0;

			var nCurPos = oRun.Content.length;
			for (var nPos = 0; nPos < nCurPos; ++nPos)
			{
                if (isStartDone && isEndDone)
                    break;

				nRangePos++;
				if (nStart - nCharsCount === nRangePos - 1 && !isStartDone)
				{
					DocPosInRun =
					{
						Class : oRun,
						Position : nPos,
					};
		
					DocPos = oRun.GetDocumentPositionFromObject();
					DocPos.push(DocPosInRun);
		
					StartPos = DocPos;
					isStartDone = true;
				}
				
				if (nEnd - nCharsCount === nRangePos - 1 && !isEndDone)
				{
					DocPosInRun =
					{
						Class : oRun,
						Position : nPos,
					};
		
					DocPos = oRun.GetDocumentPositionFromObject();
					DocPos.push(DocPosInRun);
		
					EndPos = DocPos;
					isEndDone = true;
				}
			}

			if (nRangePos !== 0)
				nCharsCount += nRangePos;
		}

		oPara.CheckRunContent(callback);

        return { startPos: StartPos, endPos: EndPos }
    };
	CTextField.prototype.UpdateSelectionByEvent = function() {
		// убираем селект, выставляем из nSelStart/nSelEnd
		let doc = this.GetDocument();
		
		let selStart = doc.event["selStart"];
		let selEnd   = doc.event["selEnd"];
		
		let docPos = this.CalcDocPos(selStart, selEnd);
		let startPos = docPos.startPos;
		let endPos   = docPos.endPos;
		
		this.content.RemoveSelection();
		if (selStart === selEnd)
			this.content.SetContentPosition(startPos, 0, 0);
		else
			this.content.SetSelectionByContentPositions(startPos, endPos);
	};
	CTextField.prototype.WriteToBinary = function(memory) {
		memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        // длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);

        let sValue = this.GetApiValue(false);
        if (sValue != null && this.IsPassword() == false) {
            memory.fieldDataFlags |= (1 << 9);
            memory.WriteString(sValue);
        }

        let nCharLimit = this.GetCharLimit();
        if (nCharLimit != 0) {
            memory.fieldDataFlags |= (1 << 10);
            memory.WriteLong(nCharLimit);
        }

        memory.fieldDataFlags |= (1 << 13);
        this.WriteRenderToBinary(memory);

        //
        // rich value
        //

        if (this.IsMultiline()) {
            memory.widgetFlags |= (1 << 12);
        }
        if (this.IsPassword()) {
            memory.widgetFlags |= (1 << 13);
        }
        if (this.IsFileSelect()) {
            memory.widgetFlags |= (1 << 20);
        }
        if (this.IsDoNotSpellCheck()) {
            memory.widgetFlags |= (1 << 22);
        }
        if (this.IsDoNotScroll()) {
            memory.widgetFlags |= (1 << 23);
        }
        if (this.IsComb()) {
            memory.widgetFlags |= (1 << 24);
        }
        // if (this.IsRichText()) {
        //     memory.widgetFlags |= (1 << 25);
        // }

        let nEndPos = memory.GetCurPosition();

        // запись флагов
        memory.Seek(memory.posForWidgetFlags);
        memory.WriteLong(memory.widgetFlags);
        memory.Seek(memory.posForFieldDataFlags);
        memory.WriteLong(memory.fieldDataFlags);

        // запись длины комманд
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
	//------------------------------------------------------------------------------------------------------------------
	CTextField.prototype.getParagraph = function() {
		return this.content.GetElement(0);
	};
	CTextField.prototype.removeBeforePaste = function() {
		let pdfDoc = this.GetDocument();
		
		let selStart = pdfDoc.event["selStart"];
		let selEnd   = pdfDoc.event["selEnd"];
		
		this.content.RemoveSelection();
		
		let docPos = this.CalcDocPos(selStart, selEnd);
		
		if (selStart === selEnd) {
			this.content.SetContentPosition(docPos.startPos, 0, 0);
		}
		else {
			this.content.SetSelectionByContentPositions(docPos.startPos, docPos.endPos);
			this.content.Remove(-1, true, false, false, false);
		}
		
		this.SetNeedRecalc(true);
	};

    function codePointToRunElement(codePoint)
	{
		let element = null;
		if (9 === codePoint)
            element = new AscWord.CRunTab();
        else if (10 === codePoint || 13 === codePoint)
            element = new AscWord.CRunBreak(AscWord.break_Line);
        else if (AscCommon.IsSpace(codePoint))
            element = new AscWord.CRunSpace(codePoint);
        else
            element = new AscWord.CRunText(codePoint);

		return element;
	}
	
    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CTextField                             = CTextField;
	window["AscPDF"].CTextField.prototype["asc_GetValue"]   = CTextField.prototype.GetValue;
	window["AscPDF"].codePointToRunElement                  = codePointToRunElement;
})();

