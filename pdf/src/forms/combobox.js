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
    function CComboBoxField(sName, aRect, oDoc)
    {
        AscPDF.CBaseListField.call(this, sName, AscPDF.FIELD_TYPES.combobox, aRect, oDoc);

        this._calcOrderIndex    = 0;
        this._doNotSpellCheck   = false;
        this._editable          = false;

        AscCommon.History.StartNoHistoryMode();
        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
        this.content = new AscPDF.CTextBoxContent(this, oDoc);
		this.contentFormat = new AscPDF.CTextBoxContent(this, oDoc, true);
        AscCommon.History.EndNoHistoryMode();

        this._markRect = {};
        this._useDisplayValue   = true;
    }
    CComboBoxField.prototype.constructor = CComboBoxField;
    AscFormat.InitClass(CComboBoxField, AscPDF.CBaseListField, AscDFH.historyitem_type_Pdf_Combobox_Field);

    CComboBoxField.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() && !this.IsEditMode())
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
        this.DrawEdit(oGraphicsWord);
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
        
        if (this.IsNeedRecalcTextTransform()) {
            this.RecalculateTextTransform();
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
        contentXLimit != this.content.XLimit || contentYFormat != this.contentFormat.Y) {
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

        let aRect = this.GetOrigRect();
        if (!aRect) {
            return;
        }

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
        return this.contentClipRect;
    };
    CComboBoxField.prototype.onMouseDown = function(x, y, e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        let bHighlight  = this.IsNeedDrawHighlight();
        let isInFocus   = oDoc.activeForm === this;

        oDoc.activeForm = this;

        if (oDoc.IsEditFieldsMode()) {
            let oController = oDoc.GetController();
            this.editShape.select(oController, this.GetPage());
            if (false == this.IsLocked()) {
                this.editShape.onMouseDown(x, y, e)
            }
            return;
        }

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

            let aOptions = this.GetOptions();
            let hasOptions = aOptions && aOptions.length != 0;

            if (pageObject.x >= this._markRect.x1 && pageObject.x <= this._markRect.x2 && pageObject.y >= this._markRect.y1 && pageObject.y <= this._markRect.y2 && hasOptions) {
                if (this.IsReadOnly()) {
                    return;
                }
                
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
        if (this.GetCurIdxs().includes(nIdx))
            return;
        
        let oDoc = this.GetDocument();
        let oPara = this.content.GetElement(0);
        let oRun = oPara.GetElement(0);

        // use history only in local form history
        if (oDoc.History == AscCommon.History) {
            AscCommon.History.StartNoHistoryMode();
        }

        oRun.ClearContent();
        let aOptions = this.GetOptions();
        if (aOptions[nIdx]) {
            if (Array.isArray(aOptions[nIdx])) {
                oRun.AddText(aOptions[nIdx][0]);
            }
            else {
                oRun.AddText(aOptions[nIdx]);
            }
        }

        if (oDoc.History == AscCommon.History) {
            AscCommon.History.EndNoHistoryMode();
        }

        this.SetNeedRecalc(true);
        this.SetNeedCommit(true);
        this.AddToRedraw();

        this.content.MoveCursorToStartPos();
        if (!Asc.editor.getDocumentRenderer().IsOpenFormsInProgress)
            this.SetNeedCheckAlign(true);

        this._bAutoShiftContentView = true;
        if (this.IsCommitOnSelChange() == true) {
            oDoc.EnterDownActiveField();
        }
    };
    CComboBoxField.prototype.SetCurIdxs = function(aIdxs) {
        if (this.IsWidget()) {
            let oDoc = this.GetDocument();
            oDoc.History.Add(new CChangesPDFListFormCurIdxs(this, this.GetCurIdxs(), aIdxs));

            if (undefined !== aIdxs[0]) {
                this.SelectOption(aIdxs[0]);
            }

            if (Asc.editor.getDocumentRenderer().IsOpenFormsInProgress)
                this.SetParentCurIdxs(aIdxs);
        }
        else 
            this.SetParentCurIdxs(aIdxs);

        this.SetNeedCommit(false);
    };
    
    CComboBoxField.prototype.SetValue = function(sValue) {
        let aIdxs = [];
        if (this.IsWidget()) {
            let oDoc        = this.GetDocument();
            let isOnOpen    = oDoc.Viewer.IsOpenFormsInProgress;

            let sTextToAdd = "";
            let aOptions = this.GetOptions();
            for (let i = 0; i < aOptions.length; i++) {
                if (Array.isArray(aOptions[i]) && aOptions[i][1] == sValue) {
                    sTextToAdd = aOptions[i][0];
                    aIdxs.push(i);
                    break;
                }
            }
            if (sTextToAdd == "") {
                for (let i = 0; i < aOptions.length; i++) {
                    if (aOptions[i] == sValue) {
                        sTextToAdd = aOptions[i];
                        aIdxs.push(i);
                        break;
                    }
                }
            }
            
            if (sTextToAdd == "")
                sTextToAdd = sValue;

            oDoc.History.Add(new CChangesPDFFormValue(this, this.GetParentValue(), sValue));

            this.UpdateDisplayValue(sTextToAdd);
            this.SetNeedRecalc(true);
            this.SetWasChanged(true);

            if (isOnOpen) {
                this.SetParentValue(sValue);
                this.SetParentCurIdxs(aIdxs);
            }
        }
        else {
            this.SetParentValue(sValue);
            this.SetParentCurIdxs(aIdxs);
        }
    };
    CComboBoxField.prototype.private_SetValue = function(sValue) {
        let aIdxs = [];
        if (this.IsWidget()) {
            let aOptions = this.GetOptions();
            let sTextToAdd = "";
            for (let i = 0; i < aOptions.length; i++) {
                if (Array.isArray(aOptions[i]) && aOptions[i][1] == sValue) {
                    sTextToAdd = aOptions[i][0];
                    aIdxs.push(i);
                    break;
                }
            }
            if (sTextToAdd == "") {
                for (let i = 0; i < aOptions.length; i++) {
                    if (aOptions[i] == sValue) {
                        sTextToAdd = aOptions[i];
                        aIdxs.push(i);
                        break;
                    }
                }
            }
            
            if (sTextToAdd == "")
                sTextToAdd = sValue;

            this.UpdateDisplayValue(sTextToAdd);
            this.SetNeedRecalc(true);
        }
        else {
            this.SetParentValue(sValue);
            this.SetParentCurIdxs(aIdxs);
        }
    };
    
    CComboBoxField.prototype.DrainLogicFrom = function(oFieldToInherit, bClearFrom) {
        AscPDF.CBaseField.prototype.DrainLogicFrom.call(this, oFieldToInherit, bClearFrom);

        this.SetOptions(oFieldToInherit.GetOptions());
        this.SetParentCurIdxs(oFieldToInherit.GetParentCurIdxs());
        this.SetEditable(oFieldToInherit.IsEditable());
        this.SetCommitOnSelChange(oFieldToInherit.IsCommitOnSelChange());

        if (bClearFrom !== false) {
            oFieldToInherit.SetOptions([]);
            oFieldToInherit.SetParentCurIdxs([]);
            oFieldToInherit.SetEditable(false);
            oFieldToInherit.SetCommitOnSelChange(false);
        }
    };
	CComboBoxField.prototype.EnterText = function(aChars) {
        let oKeystrokeEvent = this.DoKeystrokeAction(aChars);
		if (!oKeystrokeEvent["rc"])
			return false;
		
		let oKeystrokeTrigger = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke);
        if (oKeystrokeTrigger) {
            aChars = AscWord.CTextFormFormat.prototype.GetBuffer(oKeystrokeEvent["change"]);
        }
        
		if (!aChars.length)
			return false;
		
		this.UpdateSelectionByEvent();
		
		this.content.EnterText(aChars);
		
		this.SetNeedRecalc(true);
		this.SetNeedCommit(true); // флаг что значение будет применено к остальным формам с таким именем
		this._bAutoShiftContentView = true && this._doNotScroll == false;
		return true;
	};
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
	 */
    CComboBoxField.prototype.Commit = function() {
        let oDoc        = this.GetDocument();
        let aFields     = oDoc.GetAllWidgets(this.GetFullName());

        let aCurIdxs = this.GetCurIdxs();
        let aApiIdxs = this.GetParentCurIdxs();

        let sNewValue = this.GetValue();
        let isChanged = sNewValue !== this.GetParentValue();
        for (let i = 0; i < aCurIdxs.length; i++) {
            if (!aApiIdxs || aCurIdxs[i] === undefined || aApiIdxs[i] === undefined || aCurIdxs[i] !== aApiIdxs[i]) {
                isChanged = true;
                break;
            }
        }
        if (!isChanged) {
            return;
        }

        this.DoFormatAction();

        for (let i = 0; i < aFields.length; i++) {
            aFields[i].SetWasChanged(true);

            if (aFields[i].HasShiftView()) {
                aFields[i].content.MoveCursorToStartPos();

                if (aFields[i] == this) {
                    aFields[i].AddToRedraw();
                }
            }
            
            if (aFields[i] == this) {
                aFields[i]._useDisplayValue = true;
            }

            if (aCurIdxs && aCurIdxs.length != 0) {
                aFields[i].SetCurIdxs(aCurIdxs);
            }
            else {
                aFields[i].SetValue(sNewValue);
            }
            
            aFields[i].SetNeedRecalc(true);
        }

        let sFormatValue = this.GetFormatValue();
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] == this) {
                continue;
            }
            
            aFields[i].SetFormatValue(sFormatValue);
        }

        this.SetParentValue(this.GetValue());
        this.SetParentCurIdxs(aCurIdxs);

        // когда выравнивание посередине или справа, то после того
        // как ширина контента будет больше чем размер формы, выравнивание становится слева, пока текста вновь не станет меньше чем размер формы
        aFields.forEach(function(field) {
            field.SetNeedCheckAlign(true);
        });

        this.SetNeedCommit(false);
    };
	CComboBoxField.prototype.InsertChars = function(aChars) {
		this.content.EnterText(aChars);
	};
	CComboBoxField.prototype.canBeginCompositeInput = function() {
		return this.IsEditable();
	};
	/**
	 * Checks curValueIndex, corrects it and return.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CComboBoxField.prototype.UpdateIndexies = function() {
        let aOptions = this.GetOptions();
        let sValue = this.content.getAllText();
        let nIdx = -1;
        for (let i = 0; i < aOptions.length; i++) {
            if (aOptions[i][0] === sValue) {
                nIdx = i;
                break;
            }
        }
        for (let i = 0; i < aOptions.length; i++) {
            if (aOptions[i] === sValue) {
                nIdx = i;
                break;
            }
        }

        this.SetParentCurIdxs([nIdx]);
        return nIdx;
    };
    
    CComboBoxField.prototype.SetEditable = function(bValue) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets()) {
            oParent.SetEditable(bValue);
        }
        else {
            AscCommon.History.Add(new CChangesPDFComboboxFieldEditable(this, this._editable, bValue));
            this._editable = bValue;
        }

        this.SetWasChanged(true);
    };
    CComboBoxField.prototype.IsEditable = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.IsEditable();

        return this._editable;
    };
    CComboBoxField.prototype.IsCanEditText = function() {
        let oDoc = this.GetDocument();
        if (oDoc.activeForm == this && this.IsEditable() && this.IsInForm() && !this.IsReadOnly())
            return true;
        
        return false;
    };
    CComboBoxField.prototype.AddOption = function(option, nPos) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets())
            return oParent.AddOption(option, nPos);

        if (option == null) return;
        
        let formattedOption;
        
        if (Array.isArray(option) && option[0] !== undefined && option[1] !== undefined) {
            if (option[0].toString && option[1].toString) {
                formattedOption = [option[0].toString(), option[1].toString()];
            }
        } else if (option.toString) {
            formattedOption = option.toString();
        }
        
        if (formattedOption !== undefined) {
            if (nPos == undefined) {
                nPos = this._options.length;
            }

            this._options.splice(nPos, 0, formattedOption);
        }

        AscCommon.History.Add(new CChangesPDFListOption(this, nPos, [formattedOption], true));

        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
    };
    CComboBoxField.prototype.RemoveOption = function(nPos) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets())
            return oParent.RemoveOption(nPos);

        if (Number.isInteger(nPos) && nPos >= 0 && nPos < this._options.length) {
            if (this.GetCurIdxs().includes(nPos)) {
                AscCommon.History.StartNoHistoryMode();

                let oPara = this.content.GetElement(0);
                let oRun = oPara.GetElement(0);
                oRun.ClearContent();

                AscCommon.History.EndNoHistoryMode();
            }

            let option = this._options.splice(nPos, 1);

            AscCommon.History.Add(new CChangesPDFListOption(this, nPos, [option], false));

            this.SetWasChanged(true);
            this.SetNeedRecalc(true);

            return option;
        }
    };
    CComboBoxField.prototype.SetOptions = function(aOpt) {
        while (this.GetOptions().length > 0) {
            this.RemoveOption(0);
        }
        for (let i = 0; i < aOpt.length; i++) {
            this.AddOption(aOpt[i]);
        }
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

        let aOptions = this.GetOptions();
        let sValue = this.content.getAllText();
        for (let i = 0; i < aOptions.length; i++) {
            if (Array.isArray(aOptions[i])) {
                if (aOptions[i][0] == sValue)
                    return [i];
            }
        }
        for (let i = 0; i < aOptions.length; i++) {
            if (aOptions[i] === sValue) {
                return [i];
            }
        }
        
        return [];
    };
    CComboBoxField.prototype.SyncValue = function() {
        const aCurIdxs = this.GetParentCurIdxs();
        if (aCurIdxs.length != 0) {
            this.SetCurIdxs(aCurIdxs);
        }
        else {
            this.SetValue(this.GetParentValue());
        }

        this.SetFormatValue(this.GetFormatValue());
        
        this.SetNeedRecalc(true);
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

        let value = this.GetParentValue(false);
        if (value != null && Array.isArray(value) == false) {
            memory.fieldDataFlags |= (1 << 9);
            memory.WriteString(value);
        }

        // элементы списка выбора
        let aOptions = this.GetOptions(false);
        if (aOptions && aOptions.length !== 0) {
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
            curIdxs = this.GetParentCurIdxs(false);
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
        
        if (this.IsEditable(false)) {
            memory.widgetFlags |= (1 << 18);
        }
        if (this.IsDoNotSpellCheck(false)) {
            memory.widgetFlags |= (1 << 22);
        }
        if (this.IsCommitOnSelChange(false)) {
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

        this.CheckWidgetFlags(memory);
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
	CComboBoxField.prototype.GetFormatValue         = AscPDF.CTextField.prototype.GetFormatValue;
	CComboBoxField.prototype.SetFormatValue         = AscPDF.CTextField.prototype.SetFormatValue;
	CComboBoxField.prototype.IsDateFormat           = AscPDF.CTextField.prototype.IsDateFormat;
    CComboBoxField.prototype.IsNumberFormat         = AscPDF.CTextField.prototype.IsNumberFormat;
    CComboBoxField.prototype.IsSpecialKeystroke     = AscPDF.CTextField.prototype.IsSpecialKeystroke;
    CComboBoxField.prototype.GetDateFormat          = AscPDF.CTextField.prototype.GetDateFormat;
    CComboBoxField.prototype.GetFormatType          = AscPDF.CTextField.prototype.GetFormatType;
    CComboBoxField.prototype.GetFormatArgs          = AscPDF.CTextField.prototype.GetFormatArgs;
    CComboBoxField.prototype.GetValidateType        = AscPDF.CTextField.prototype.GetValidateType;
    CComboBoxField.prototype.GetValidateArgs        = AscPDF.CTextField.prototype.GetValidateArgs;
    CComboBoxField.prototype.SetPlaceholder         = AscPDF.CTextField.prototype.SetPlaceholder;
    CComboBoxField.prototype.SetRegularExp          = AscPDF.CTextField.prototype.SetRegularExp;
    CComboBoxField.prototype.SetArbitaryMask        = AscPDF.CTextField.prototype.SetArbitaryMask;
    CComboBoxField.prototype.ClearFormat            = AscPDF.CTextField.prototype.ClearFormat;
    CComboBoxField.prototype.SetDrawFromStream      = AscPDF.CTextField.prototype.SetDrawFromStream;
    CComboBoxField.prototype.DrawMarker             = AscPDF.CTextField.prototype.DrawMarker;
    CComboBoxField.prototype.beforeCompositeInput   = AscPDF.CTextField.prototype.beforeCompositeInput;
    CComboBoxField.prototype.IsCanCommit            = AscPDF.CTextField.prototype.IsCanCommit;

	window["AscPDF"].CComboBoxField = CComboBoxField;
})();

