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
	 * Class representing a combobox field.
	 * @constructor
     * @extends {CBaseListField}
	 */
    function CComboBoxField(sName, nPage, aRect, oDoc)
    {
        AscPDF.CBaseListField.call(this, sName, AscPDF.FIELD_TYPES.combobox, nPage, aRect, oDoc);

        this._calcOrderIndex    = 0;
        this._doNotSpellCheck   = false;
        this._editable          = false;

        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
		this.contentFormat = new AscPDF.CTextBoxContent(this, oDoc, true);

        this._markRect = {};
        this._useDisplayValue   = true;
    }
    CComboBoxField.prototype = Object.create(AscPDF.CBaseListField.prototype);
	CComboBoxField.prototype.constructor = CComboBoxField;

    CComboBoxField.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        let oDoc = this.GetDocument();
        
        this.Recalculate();
        this.DrawBackground(oGraphicsPDF);
        
        let oContentToDraw = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format) && this.IsNeedDrawHighlight() ? this.contentFormat : this.content;
        this.curContent = oContentToDraw; // запоминаем текущий контент

        if (oDoc.activeForm == this)
            this.CheckFormViewWindow();

        oGraphicsWord.AddClipRect(this.contentClipRect.X, this.contentClipRect.Y, this.contentClipRect.W, this.contentClipRect.H);
        oContentToDraw.Draw(0, oGraphicsWord);
        // redraw target cursor if field is selected
        if (oDoc.activeForm == this && oContentToDraw.IsSelectionUse() == false && this.IsCanEditText())
            oContentToDraw.RecalculateCurPos();
        
        oGraphicsWord.RemoveLastClip();
        this.DrawBorders(oGraphicsPDF);

        this.DrawLocks(oGraphicsPDF);
    };
    CComboBoxField.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        if (this.IsNeedCheckAlign()) {
            this.CheckAlignInternal();
        }
        
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
    CComboBoxField.prototype.RecalculateContentRect = function() {
        let aOrigRect = this.GetOrigRect();

        let X       = aOrigRect[0];
        let Y       = aOrigRect[1];
        let nWidth  = (aOrigRect[2] - aOrigRect[0]);
        let nHeight = (aOrigRect[3] - aOrigRect[1]);

        let oMargins = this.GetMarginsFromBorders();

        let contentX        = (X + 2 * oMargins.left) * g_dKoef_pt_to_mm;
        let contentY        = (Y + oMargins.top) * g_dKoef_pt_to_mm;
        let contentXLimit   = (X + nWidth - 2 * oMargins.left) * g_dKoef_pt_to_mm;

        let nContentH       = this.GetTextHeight(this.content);
        let nContentHFormat = this.GetTextHeight(this.contentFormat);

        contentY            = Y * g_dKoef_pt_to_mm + (nHeight * g_dKoef_pt_to_mm - nContentH) / 2;
        let contentYFormat  = Y * g_dKoef_pt_to_mm + (nHeight * g_dKoef_pt_to_mm - nContentHFormat) / 2;

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
    CComboBoxField.prototype.CalculateContentClipRect = function() {
        if (!this.content)
            return;

        let aRect       = this.GetOrigRect();
        let X           = aRect[0];
        let Y           = aRect[1];
        let nWidth      = aRect[2] - aRect[0];
        let nHeight     = aRect[3] - aRect[1];
        let oMargins    = this.GetMarginsFromBorders();

        let contentX        = (X + 2 * oMargins.left) * g_dKoef_pt_to_mm;
        let contentXLimit   = (X + nWidth - 2 * oMargins.left) * g_dKoef_pt_to_mm;

        this.contentClipRect = {
            X: contentX,
            Y: (Y + oMargins.top) * g_dKoef_pt_to_mm,
            W: contentXLimit - contentX,
            H: (nHeight - oMargins.top - oMargins.bottom) * g_dKoef_pt_to_mm,
            Page: this.GetPage()
        }
    };
    CComboBoxField.prototype.DrawMarker = function(oCtx) {
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
    CComboBoxField.prototype.onMouseDown = function(x, y, e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        let bHighlight  = this.IsNeedDrawHighlight();
        let isInFocus   = oDoc.activeForm === this;

        oDoc.activeForm = this;

        function callbackAfterFocus(x, y, e) {
            oDoc.SetLocalHistory();
            if (false == e.ShiftKey) {
                oDoc.SelectionSetStart(x, y, e);
            }
            else {
                this.content.StartSelectionFromCurPos();
                oDoc.SelectionSetEnd(x, y, e);
            }
            
            this.SetInForm(true);

            var pageObject = oViewer.getPageByCoords(x, y);

            if (pageObject.x >= this._markRect.x1 && pageObject.x <= this._markRect.x2 && pageObject.y >= this._markRect.y1 && pageObject.y <= this._markRect.y2 && this._options.length != 0) {
                editor.sendEvent("asc_onShowPDFFormsActions", this, x, y);
                this.content.MoveCursorToStartPos();
            }
            
            this.SetDrawHighlight(false);
            if (this.IsNeedDrawFromStream() == true) {
                this.SetDrawFromStream(false);
                this.AddToRedraw();
            }
            else if (this.curContent === this.contentFormat || bHighlight) {
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
    
    /**
	 * Selects the specified option.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.SelectOption = function(nIdx) {
        if (this.GetCurIdxs() == nIdx)
            return;

        let oPara = this.content.GetElement(0);
        let oRun = oPara.GetElement(0);

        oRun.ClearContent();
        if (Array.isArray(this._options[nIdx])) {
            oRun.AddText(this._options[nIdx][0]);
        }
        else {
            oRun.AddText(this._options[nIdx]);
        }

        this.SetNeedRecalc(true);
        this.SetNeedCommit(true);
        this.AddToRedraw();

        this.content.MoveCursorToStartPos();
        if (!Asc.editor.getDocumentRenderer().IsOpenFormsInProgress)
            this.SetNeedCheckAlign(true);
    };
    CComboBoxField.prototype.SetCurIdxs = function(aIdxs) {
        if (this.IsWidget()) {
            if (undefined !== aIdxs[0]) {
                this.SelectOption(aIdxs[0]);
            }

            if (editor.getDocumentRenderer().IsOpenFormsInProgress)
                this.SetApiCurIdxs(aIdxs);
        }
        else 
            this.SetApiCurIdxs(aIdxs);

        this.SetNeedCommit(false);
    };
    
    CComboBoxField.prototype.SetValue = function(sValue) {
        let aIdxs = [];
        if (this.IsWidget()) {
            let oDoc        = this.GetDocument();
            let isOnOpen    = oDoc.Viewer.IsOpenFormsInProgress;

            let sTextToAdd = "";
            for (let i = 0; i < this._options.length; i++) {
                if (Array.isArray(this._options[i]) && this._options[i][1] == sValue) {
                    sTextToAdd = this._options[i][0];
                    aIdxs.push(i);
                    break;
                }
            }
            if (sTextToAdd == "") {
                for (let i = 0; i < this._options.length; i++) {
                    if (this._options[i] == sValue) {
                        sTextToAdd = this._options[i];
                        aIdxs.push(i);
                        break;
                    }
                }
            }
            
            if (sTextToAdd == "")
                sTextToAdd = sValue;

            
            this.UpdateDisplayValue(sTextToAdd);
            this.SetNeedRecalc(true);
            this.SetWasChanged(true);

            if (isOnOpen) {
                this.SetApiValue(sValue);
                this.SetApiCurIdxs(aIdxs);
            }
        }
        else {
            this.SetApiValue(sValue);
            this.SetApiCurIdxs(aIdxs);
        }
    };
    
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.SyncField = function() {
        let aFields = this.GetDocument().GetAllWidgets(this.GetFullName());
        
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {

                this.SetCalcOrderIndex(aFields[i].GetCalcOrderIndex());
                this.SetDoNotSpellCheck(aFields[i].IsDoNotSpellCheck());
                this.SetEditable(aFields[i].IsEditable());

                let oPara = this.content.GetElement(0);
                let oParaToCopy = aFields[i].content.GetElement(0);

                oPara.ClearContent();
                for (var nPos = 0; nPos < oParaToCopy.Content.length - 1; nPos++) {
                    oPara.Internal_Content_Add(nPos, oParaToCopy.GetElement(nPos).Copy());
                }
                oPara.CheckParaEnd();
                
                this._options = aFields[i]._options.slice();
                break;
            }
        }

        oDoc.EndNoHistoryMode();
    };
	CComboBoxField.prototype.EnterText = function(aChars) {
		if (!this.DoKeystrokeAction(aChars))
			return false;
		
		let doc = this.GetDocument();
		aChars = AscWord.CTextFormFormat.prototype.GetBuffer(doc.event["change"]);
		if (!aChars.length)
			return false;
		
		this.UpdateSelectionByEvent();
		
		this.content.EnterText(aChars);
		
		this.SetNeedRecalc(true);
		this.SetNeedCommit(true); // флаг что значение будет применено к остальным формам с таким именем
		this._bAutoShiftContentView = true && this._doNotScroll == false;
		return true;
	};
	CComboBoxField.prototype.CorrectEnterText = function(oldValue, newValue) {
		if (!this.DoKeystrokeAction(newValue))
			return false;
		
		let doc = this.GetDocument();
		newValue = AscWord.CTextFormFormat.prototype.GetBuffer(doc.event["change"]);
		if (!newValue.length && !oldValue.length)
			return false;
		
		let result = this.content.CorrectEnterText(oldValue, newValue, function(run, inRunPos, codePoint){return true;});
		
		this.SetNeedRecalc(true);
		this.SetNeedCommit(true); // флаг что значение будет применено к остальным формам с таким именем
		this._bAutoShiftContentView = true && this._doNotScroll == false;
		return result;
	};
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.Commit = function() {
        let oDoc        = this.GetDocument();
        let aFields     = oDoc.GetAllWidgets(this.GetFullName());

        oDoc.StartNoHistoryMode();
        if (this.DoFormatAction() == false) {
            this.UndoNotAppliedChanges();
            if (this.IsChanged() == false)
                this.SetDrawFromStream(true);

            oDoc.EndNoHistoryMode();
            return;
        }
        oDoc.EndNoHistoryMode();
        
        let aCurIdxs = this.GetCurIdxs();
        let aApiIdxs = this.GetApiCurIdxs();

        let isChanged = false;
        for (let i = 0; i < aCurIdxs.length; i++) {
            if (aCurIdxs[i] === undefined || aApiIdxs[i] === undefined || aCurIdxs[i] !== aApiIdxs[i]) {
                isChanged = true;
                break;
            }
        }

        if (isChanged) {
            oDoc.History.Add(new CChangesPDFListFormCurIdxs(this, this.GetApiCurIdxs(), aCurIdxs));
            this.SetApiValue(this.GetValue());
            this.SetApiCurIdxs(this.GetCurIdxs());
        }
        
        oDoc.StartNoHistoryMode();

        if (aFields.length == 1)
            this.SetNeedCommit(false);

        let sDisplayValue = this.content.getAllText();
        this.UpdateDisplayValue(sDisplayValue);

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

            aFields[i].UpdateDisplayValue(sDisplayValue);
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
	CComboBoxField.prototype.InsertChars = function(aChars) {
		this.content.EnterText(aChars);
	};
	CComboBoxField.prototype.canBeginCompositeInput = function() {
		return this.IsEditable();
	};
	CComboBoxField.prototype.beforeCompositeInput = function() {
		this.DoKeystrokeAction();
		this.content.Remove(1, true, false, false);
	};
	/**
	 * Checks curValueIndex, corrects it and return.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CComboBoxField.prototype.UpdateIndexies = function() {
        let sValue = this.content.GetElement(0).GetText({ParaEndToSpace: false});
        let nIdx = -1;
        for (let i = 0; i < this._options.length; i++) {
            if (this._options[i][0] === sValue) {
                nIdx = i;
                break;
            }
        }
        for (let i = 0; i < this._options.length; i++) {
            if (this._options[i] === sValue) {
                nIdx = i;
                break;
            }
        }

        this.SetApiCurIdxs([nIdx]);
        return nIdx;
    };
    
    CComboBoxField.prototype.SetEditable = function(bValue) {
        this._editable = bValue;
    };
    CComboBoxField.prototype.IsEditable = function() {
        return this._editable;
    };
    CComboBoxField.prototype.IsCanEditText = function() {
        let oDoc = this.GetDocument();
        if (oDoc.activeForm == this && this.IsEditable() && this.IsInForm())
            return true;
        
        return false;
    };
    CComboBoxField.prototype.SetOptions = function(aOpt) {
        let aOptToPush = [];
        for (let i = 0; i < aOpt.length; i++) {
            if (aOpt[i] == null)
                continue;
            if (typeof(aOpt[i]) == "string" && aOpt[i] != "")
                aOptToPush.push(aOpt[i]);
            else if (Array.isArray(aOpt[i]) && aOpt[i][0] != undefined && aOpt[i][1] != undefined) {
                if (aOpt[i][0].toString && aOpt[i][1].toString) {
                    aOptToPush.push([aOpt[i][0].toString(), aOpt[i][1].toString()])
                }
            }
            else if (typeof(aOpt[i]) != "string" && aOpt[i].toString) {
                aOptToPush.push(aOpt[i].toString());
            }
        }

        this._options = aOptToPush;
    };
    
    CComboBoxField.prototype.GetValue = function(bDisplayValue) {
        let sValue = this._useDisplayValue ? this._displayValue : this.content.getAllText();
        if (bDisplayValue)
            return sValue;

        let aOptions = this.GetOptions();
        for (let i = 0; i < aOptions.length; i++) {
            if (Array.isArray(aOptions[i])) {
                if (aOptions[i][0] == sValue)
                    return aOptions[i][1];
            }
        }
        
        return sValue;
    };

    /**
	 * Gets current index.
	 * @memberof CComboBoxField
     * @param {boolean} [bApiValue=false] - if true -> returns api value (if false -> current value of this field)
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.GetCurIdxs = function(bApiValue) {
        if (bApiValue)
            return this._currentValueIndices;

        let sValue = this.content.GetElement(0).GetText({ParaEndToSpace: false});
        for (let i = 0; i < this._options.length; i++) {
            if (Array.isArray(this._options[i])) {
                if (this._options[i][0] == sValue)
                    return [i];
            }
        }
        for (let i = 0; i < this._options.length; i++) {
            if (this._options[i] === sValue) {
                return [i];
            }
        }
        
        return -1;
    };
	
    CComboBoxField.prototype.ProcessAutoFitContent = function(oContent) {
        let oPara   = oContent.GetElement(0);
        let oRun    = oPara.GetElement(0);
        let oTextPr = oRun.Get_CompiledPr(true);
        let oBounds = this.getFormRelRect();

        g_oTextMeasurer.SetTextPr(oTextPr, null);
        g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);

        var nTextHeight = g_oTextMeasurer.GetHeight();
        var nMaxWidth   = oPara.RecalculateMinMaxContentWidth(false).Max;
        var nFontSize   = oTextPr.FontSize;

        if (nMaxWidth < 0.001 || nTextHeight < 0.001 || oBounds.W < 0.001 || oBounds.H < 0.001)
    	    return nTextHeight;

        let nNewFontSize = (Math.min(nFontSize * oBounds.H / nTextHeight * 0.9, 100, nFontSize * oBounds.W / nMaxWidth) * 100 >> 0) / 100;
        oRun.SetFontSize(nNewFontSize);
        oPara.Recalculate_Page(0);

        oTextPr.FontSize    = nNewFontSize;
        oTextPr.FontSizeCS  = nNewFontSize;

        this.AddToRedraw();
    };
    CComboBoxField.prototype.GetTextHeight = function(oContent) {
        let oPara   = oContent.GetElement(0);
        let oRun    = oPara.GetElement(0);
        let oTextPr = oRun.Get_CompiledPr(true);

        g_oTextMeasurer.SetTextPr(oTextPr, null);
        g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);

        return g_oTextMeasurer.GetHeight();
    };
    
    CComboBoxField.prototype.WriteToBinary = function(memory) {
		memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        // длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);

        let value = this.GetApiValue(false);
        if (value != null && Array.isArray(value) == false) {
            memory.fieldDataFlags |= (1 << 9);
            memory.WriteString(value);
        }

        // элементы списка выбора
        let aOptions = this.GetOptions();
        if (aOptions) {
            memory.fieldDataFlags |= (1 << 10);
            memory.WriteLong(aOptions.length);
            for (let i = 0; i < aOptions.length; i++) {
                memory.WriteString(Array.isArray(aOptions[i]) ? aOptions[i][1] : "");
                memory.WriteString(Array.isArray(aOptions[i]) ? aOptions[i][0] : aOptions[i]);
            }
        }

        if (value != null && Array.isArray(value) == true) {
            // флаг что значение - это массив
            memory.fieldDataFlags |= (1 << 13);
            memory.WriteLong(value.length);
            for (let i = 0; i < value.length; i++) {
                memory.WriteString(value[i]);
            }
        }

        // массив I (выделенные значения списка)
        let curIdxs;
        if ([AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.listbox].includes(this.GetType())) {
            curIdxs = this.GetApiCurIdxs(false);
        }
        if (curIdxs) {
            memory.fieldDataFlags |= (1 << 14);
            memory.WriteLong(curIdxs.length);
            for (let i = 0; i < curIdxs.length; i++) {
                memory.WriteLong(curIdxs[i]);
            }
        }

        memory.fieldDataFlags |= (1 << 15);
        this.WriteRenderToBinary(memory);
        
        // top index
        
        if (this.IsEditable()) {
            memory.widgetFlags |= (1 << 18);
        }
        if (this.IsDoNotSpellCheck()) {
            memory.fieldDataFlags |= (1 << 22);
        }
        if (this.IsCommitOnSelChange()) {
            memory.widgetFlags |= (1 << 26);
        }
        
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
	CComboBoxField.prototype.IsDoNotScroll = function() {
		return true;
	};
	
	if (!window["AscPDF"])
		window["AscPDF"] = {};
	
	CComboBoxField.prototype.Remove                 = AscPDF.CTextField.prototype.Remove;
	CComboBoxField.prototype.MoveCursorLeft         = AscPDF.CTextField.prototype.MoveCursorLeft;
	CComboBoxField.prototype.MoveCursorRight        = AscPDF.CTextField.prototype.MoveCursorRight;
	CComboBoxField.prototype.SelectionSetStart      = AscPDF.CTextField.prototype.SelectionSetStart;
	CComboBoxField.prototype.SelectionSetEnd        = AscPDF.CTextField.prototype.SelectionSetEnd;
	CComboBoxField.prototype.CheckFormViewWindow    = AscPDF.CTextField.prototype.CheckFormViewWindow;
	CComboBoxField.prototype.SetAlign               = AscPDF.CTextField.prototype.SetAlign;
	CComboBoxField.prototype.GetAlign               = AscPDF.CTextField.prototype.GetAlign;
	CComboBoxField.prototype.CheckAlignInternal     = AscPDF.CTextField.prototype.CheckAlignInternal;
	CComboBoxField.prototype.IsTextOutOfForm        = AscPDF.CTextField.prototype.IsTextOutOfForm;
	CComboBoxField.prototype.SetDoNotSpellCheck     = AscPDF.CTextField.prototype.SetDoNotSpellCheck;
	CComboBoxField.prototype.IsDoNotSpellCheck      = AscPDF.CTextField.prototype.IsDoNotSpellCheck;
	CComboBoxField.prototype.DoValidateAction       = AscPDF.CTextField.prototype.DoValidateAction;
	CComboBoxField.prototype.DoKeystrokeAction      = AscPDF.CTextField.prototype.DoKeystrokeAction;
	CComboBoxField.prototype.DoFormatAction         = AscPDF.CTextField.prototype.DoFormatAction;
	CComboBoxField.prototype.CalcDocPos             = AscPDF.CTextField.prototype.CalcDocPos;
	CComboBoxField.prototype.GetCalcOrderIndex      = AscPDF.CTextField.prototype.GetCalcOrderIndex;
	CComboBoxField.prototype.SetCalcOrderIndex      = AscPDF.CTextField.prototype.SetCalcOrderIndex;
	CComboBoxField.prototype.UndoNotAppliedChanges  = AscPDF.CTextField.prototype.UndoNotAppliedChanges;
	CComboBoxField.prototype.SelectAllText          = AscPDF.CTextField.prototype.SelectAllText;
	CComboBoxField.prototype.UpdateDisplayValue     = AscPDF.CTextField.prototype.UpdateDisplayValue;
	CComboBoxField.prototype.onMouseUp              = AscPDF.CTextField.prototype.onMouseUp;
	CComboBoxField.prototype.OnContentChange        = AscPDF.CTextField.prototype.OnContentChange;
	CComboBoxField.prototype.UpdateSelectionByEvent = AscPDF.CTextField.prototype.UpdateSelectionByEvent;
	CComboBoxField.prototype.SetNeedCheckAlign      = AscPDF.CTextField.prototype.SetNeedCheckAlign;
	CComboBoxField.prototype.IsNeedCheckAlign       = AscPDF.CTextField.prototype.IsNeedCheckAlign;

	window["AscPDF"].CComboBoxField = CComboBoxField;
})();

