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
	 * Class representing a text field.
	 * @constructor
     * @extends {CBaseField}
	 */
    function CTextField(sName, nPage, aRect, oDoc)
    {
        AscPDF.CBaseField.call(this, sName, AscPDF. FIELD_TYPES.text, nPage, aRect, oDoc);
        
        this._alignment         = AscPDF.ALIGN_TYPE.left;
        this._charLimit         = 0; // to do
        this._comb              = false;
        this._defaultStyle      = Object.assign({}, {}); // to do (must not be fileSelect flag)
        this._doNotScroll       = false;
        this._doNotSpellCheck   = false;
        this._multiline         = false;
        this._password          = false;
        this._richText          = false; // to do связанные свойства, методы
        this._richValue         = [];
        this._textFont          = "ArialMT";
        this._fileSelect        = false;
        this._value             = undefined;

        // internal
        TurnOffHistory();
		this.content = new AscPDF.CTextBoxContent(this, oDoc);

        // content for formatting value
        // Note: draw this content instead of main if form has a "format" action
		this.contentFormat = new AscPDF.CTextBoxContent(this, oDoc);

        this._scrollInfo = null;
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
    CTextField.prototype.IsEditable = function() {
        return this.IsNeedDrawHighlight() == false;
    };
    CTextField.prototype.SetCharLimit = function(nChars) {
        let oViewer = editor.getDocumentRenderer();

        if (typeof(nChars) == "number" && nChars <= 500 && nChars > 0 && this.GetFileSelect() === false) {
            nChars = Math.round(nChars);
            if (this._charLimit != nChars) {

                if (oViewer.IsOpenFormsInProgress != true) {
                    let aChars = [];
                    let sText = this.content.GetElement(0).GetText({ParaEndToSpace: false});
                    if (sText.length > nChars) {
                        for (let i = 0, nCount = Math.min(nChars, sText.length); i < nCount; i++) {
                            aChars.push(sText[i].charCodeAt(0));
                        }
    
                        this.EnterText(aChars);
                    }
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
    CTextField.prototype.GetDoNotScroll = function() {
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
    CTextField.prototype.GetFileSelect = function() {
        return this._fileSelect;
    };
    CTextField.prototype.SetMultiline = function(bMultiline) {
        if (bMultiline == true && this.fileSelect != true) {
            this.content.SetUseXLimit(true);
            this.contentFormat.SetUseXLimit(true);
            this._multiline = true;
        }
        else if (bMultiline === false) {
            this.content.SetUseXLimit(false);
            this.contentFormat.SetUseXLimit(false);
            this._multiline = false;
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
    CTextField.prototype.SetRichText = function(bRichText) {
        this._richText = bRichText;
    };
    CTextField.prototype.SetValue = function(sValue) {
        if (this.IsWidget()) {
            let oPara   = this.content.GetElement(0);
            let oRun    = oPara.GetElement(0);
            oPara.RemoveFromContent(1, oPara.GetElementsCount() - 1);
            oRun.ClearContent();

            let oTextFormat = new AscWord.CTextFormFormat();
            let arrBuffer   = oTextFormat.GetBuffer(sValue);
            
            if (sValue) {
                for (let index = 0; index < arrBuffer.length; ++index) {
                    let codePoint = arrBuffer[index];
                    if (9 === codePoint) // \t
                        oRun.AddToContent(index, new AscWord.CRunTab(), true);
                    else if (10 === codePoint || 13 === codePoint) // \n \r
                        oRun.AddToContent(index, new AscWord.CRunBreak(AscWord.break_Line), true);
                    else if (AscCommon.IsSpace(codePoint)) // space
                        oRun.AddToContent(index, new AscWord.CRunSpace(codePoint), true);
                    else
                        oRun.AddToContent(index, new AscWord.CRunText(codePoint), true);
                }
                
                this.content.MoveCursorToStartPos();
            }

            this.SetNeedRecalc(true);
            this.SetWasChanged(true);

            if (editor.getDocumentRenderer().IsOpenFormsInProgress && this.GetParent() == null)
                this.SetApiValue(sValue);
        }
        else
            this.SetApiValue(sValue);
    };
    
    CTextField.prototype.GetCalcOrderIndex = function() {
        return this.field.GetDocument().GetCalculateInfo().names.indexOf(this.field.GetFullName());
    };
    CTextField.prototype.SetCalcOrderIndex = function(nIdx) {
        let oCalcInfo = this.GetDocument().GetCalculateInfo();
        let oCalcTrigget = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Calculate);
        if (oCalcTrigget == null || nIdx < 0)
            return false;

        let nCurIdx = oCalcInfo.names.indexOf(this.GetFullName());
        if (nCurIdx == nIdx)
            return true;

        oCalcInfo.names.splice(nCurIdx, 1);
        if (nIdx > oCalcInfo.names.length)
            oCalcInfo.names.splice(nIdx, 0, this.GetFullName());
        else
            oCalcInfo.names.push(this.GetFullName());

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
        // to do обработать rich value
        let oPara = this.content.GetElement(0);
        oPara.SetApplyToAll(true);
        let sValue = oPara.GetSelectedText(true, {NewLine: true});
        oPara.SetApplyToAll(false);

        return sValue;
    };
        
    CTextField.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();

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

        oGraphicsWord.AddClipRect(this.contentRect.X, this.contentRect.Y, this.contentRect.W, this.contentRect.H);
        oContentToDraw.Draw(0, oGraphicsWord);

        oGraphicsWord.RemoveLastClip();
        this.DrawBorders(oGraphicsPDF);

        // redraw target cursor if field is selected
        if (oDoc.activeForm == this && oContentToDraw.IsSelectionUse() == false && this.IsEditable())
            oContentToDraw.RecalculateCurPos();
    };
    CTextField.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        let aRect = this.GetRect();

        let X       = aRect[0];
        let Y       = aRect[1];
        let nWidth  = ((aRect[2]) - (aRect[0]));
        let nHeight = ((aRect[3]) - (aRect[1]));

        // save pos in page.
        this._pagePos = {
            x: X,
            y: Y,
            w: nWidth,
            h: nHeight
        };

        let oMargins = this.GetMarginsFromBorders(false, false);
        
        let contentX = this.IsComb() ? (X + oMargins.left) * g_dKoef_pix_to_mm : (X + 2 * oMargins.left) * g_dKoef_pix_to_mm;
        let contentY = (Y + 2.5 * oMargins.top) * g_dKoef_pix_to_mm;
        let contentXLimit = this.IsComb() ? (X + nWidth - oMargins.left) * g_dKoef_pix_to_mm : (X + nWidth - 2 * oMargins.left) * g_dKoef_pix_to_mm;
        let contentYLimit = (Y + nHeight - oMargins.bottom) * g_dKoef_pix_to_mm;
        
        if ((this.borderStyle == "solid" || this.borderStyle == "dashed") && 
        this._comb == true && this._charLimit > 1) {
            contentX = (X) * g_dKoef_pix_to_mm;
            contentXLimit = (X + nWidth) * g_dKoef_pix_to_mm;
        }
        
        if (this.IsMultiline() == false) {
            // выставляем текст посередине
            let nContentH = this.content.GetElement(0).Get_EmptyHeight();
            contentY = (Y + nHeight / 2) * g_dKoef_pix_to_mm - nContentH / 2;
        }

        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;
        
        this.contentRect.X = contentX;
        this.contentRect.Y = contentY;
        this.contentRect.W = contentXLimit - contentX;
        this.contentRect.H = contentYLimit - contentY;

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this.content.X      = this.contentFormat.X = this._oldContentPos.X = contentX;
            this.content.Y      = this.contentFormat.Y = this._oldContentPos.Y = contentY;
            this.content.XLimit = this.contentFormat.XLimit = this._oldContentPos.XLimit = contentXLimit;
            this.content.YLimit = this.contentFormat.YLimit = this._oldContentPos.YLimit = 20000;
            this.content.Recalculate_Page(0, true);
            this.contentFormat.Recalculate_Page(0, true);
        }
        else if (this.IsNeedRecalc()) {
            this.contentFormat.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
            this.content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
        }

        this.SetNeedRecalc(false);
    };

    CTextField.prototype.onMouseDown = function(x, y, e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        function callbackAfterFocus(x, y, e) {
            
            let oPos    = AscPDF.GetPageCoordsByGlobalCoords(x, y, this.GetPage());
            let X       = oPos["X"];
            let Y       = oPos["Y"];

            editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
            editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;

            oViewer.Api.WordControl.m_oDrawingDocument.TargetStart();
            oViewer.Api.WordControl.m_oDrawingDocument.showTarget(true);

            this.content.Selection_SetStart(X, Y, 0, e);
            this.content.RecalculateCurPos();

            if (this._doNotScroll == false && this.IsMultiline())
                this.UpdateScroll(true);

            if (this.IsNeedDrawFromStream() == true) {
                this.SetDrawFromStream(false);
                this.AddToRedraw();
            }
            else if (this.curContent === this.contentFormat) {
                this.AddToRedraw();
            }
        }
        
        // вызываем выставление курсора после onFocus, если уже в фокусе, тогда сразу.
        if (oDoc.activeForm != this && this._triggers.OnFocus && this._triggers.OnFocus.Actions.length > 0)
            oActionsQueue.callBackAfterFocus = callbackAfterFocus.bind(this, x, y, e);
        else
            callbackAfterFocus.bind(this, x, y, e)();

        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseDown);
        if (oDoc.activeForm != this)
            this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);

        oDoc.activeForm = this;
    };
      
    CTextField.prototype.ScrollVertical = function(scrollY, maxYscroll) {
        let oViewer = editor.getDocumentRenderer();
        this._bAutoShiftContentView = false;

        let nScrollCoeff                = scrollY / maxYscroll;
        this._curShiftView.y            = -nScrollCoeff * maxYscroll;
        this._scrollInfo.scrollCoeff    = nScrollCoeff;
        this.AddToRedraw();
        oViewer._paint();
        oViewer.onUpdateOverlay();
    };
    CTextField.prototype.ScrollVerticalEnd = function() {
        let nHeightPerPara  = this.content.GetElement(1).Y - this.content.GetElement(0).Y;
        let nShiftCount     = this._curShiftView.y / nHeightPerPara; // количество смещений в длинах параграфов
        if (Math.abs(Math.round(nShiftCount) - nShiftCount) <= 0.001)
            return;

        let nMaxShiftY                  = this._scrollInfo.scroll.maxScrollY;
        this._curShiftView.y            = Math.round(nShiftCount) * nHeightPerPara;
        this._bAutoShiftContentView     = false;
        this._scrollInfo.scrollCoeff    = Math.abs(this._curShiftView.y / nMaxShiftY);
        
        this.AddToRedraw();
        editor.getDocumentRenderer()._paint();
    };
    CTextField.prototype.UpdateScroll = function(bShow) {
        if (this.IsMultiline() == false)
            return;
            
        let oContentBounds  = this.content.GetContentBounds(0);
        let oContentRect    = this.getFormRelRect();
        let oFormRect       = this.getFormRect();

        let nContentH       = oContentBounds.Bottom - oContentBounds.Top;
        let oScroll, oScrollDocElm, oScrollSettings;

        if (typeof(bShow) != "boolean" && this._scrollInfo)
            bShow = this._scrollInfo.scroll.canvas.style.display == "none" ? false : true;

        if (nContentH < oContentRect.H || this._doNotScroll) {
            
            if (this._scrollInfo)
                this._scrollInfo.scroll.canvas.style.display = "none";
            return;
        }

        let oViewer = editor.getDocumentRenderer();
        
        let oGlobalCoords1  = AscPDF.GetGlobalCoordsByPageCoords(oFormRect.X, oFormRect.Y, this.GetPage());
        let oGlobalCoords2  = AscPDF.GetGlobalCoordsByPageCoords(oFormRect.X + oFormRect.W, oFormRect.Y + oFormRect.H, this.GetPage());
        let oBorderWidth    = this.GetBordersWidth(true);
        
        if (this._scrollInfo == null && oContentBounds.Bottom - oContentBounds.Top > oContentRect.H) {
            oViewer.scrollCount++;
            oScrollDocElm = document.createElement('div');
            document.getElementById('editor_sdk').appendChild(oScrollDocElm);
            oScrollDocElm.id = "formScroll_" + oViewer.scrollCount;
            oScrollDocElm.style.top         = Math.round(oGlobalCoords1["Y"]) + 'px';
            oScrollDocElm.style.left        = Math.round(oGlobalCoords2["X"]) + 'px';
            oScrollDocElm.style.position    = "absolute";
            oScrollDocElm.style.display     = "block";
			oScrollDocElm.style.width       = "14px";
			oScrollDocElm.style.height      = Math.round(oGlobalCoords2["Y"]) - Math.round(oGlobalCoords1["Y"]) + "px";
            oScrollDocElm.style.zIndex      = 0;

            let nMaxShiftY = oContentRect.H - nContentH;

            oScrollSettings = editor.WordControl.CreateScrollSettings();
            oScrollSettings.isHorizontalScroll = false;
		    oScrollSettings.isVerticalScroll = true;
		    oScrollSettings.contentH = Math.abs(nMaxShiftY);
            oScrollSettings.screenH = 0;
            oScrollSettings.scrollerMinHeight = 5;
            
            let nScrollCoeff = this._curShiftView.y / nMaxShiftY;
            
            let oPara = this.content.GetElement(0);
            let oCurParaHeight  = oPara.Lines[0].Bottom - oPara.Lines[0].Top;

            oScrollSettings.vscrollStep = oCurParaHeight;
            oScroll = new AscCommon.ScrollObject(oScrollDocElm.id, oScrollSettings);

            let oThis = this;
            oScroll.bind("scrollvertical", function(evt) {
                oThis.ScrollVertical(evt.scrollD, evt.maxScrollY);
            });
            oScroll.bind("mouseup", function(evt) {
                if (oThis.GetType() == AscPDF.FIELD_TYPES.listbox)
                    oThis.ScrollVerticalEnd();
            });

            oScroll.scrollVCurrentY = oScroll.maxScrollY * nScrollCoeff;
            
            this._scrollInfo = {
                scroll:         oScroll,
                docElem:        oScrollDocElm,
                baseYPos:       parseInt(oScrollDocElm.style.top),
                oldZoom:        oViewer.zoom,
                scrollCoeff:    nScrollCoeff // проскроленная часть
            }

            oScroll.Repos(oScrollSettings, false);
        }
        else if (this._scrollInfo) {
            let nMaxShiftY = oContentRect.H - nContentH;
            let needUpdatePos = this._scrollInfo.oldZoom != oViewer.zoom || oGlobalCoords1["Y"] - oBorderWidth.top != this._scrollInfo.baseYPos;

            if (needUpdatePos) {
                oScrollSettings = editor.WordControl.CreateScrollSettings();
                oScrollSettings.isHorizontalScroll = false;
                oScrollSettings.isVerticalScroll = true;
                oScrollSettings.contentH = Math.abs(nMaxShiftY);
                oScrollSettings.screenH = 0;
                oScrollSettings.scrollerMinHeight = 5;
                this._scrollInfo.scroll.scrollVCurrentY = this._scrollInfo.scroll.maxScrollY * this._scrollInfo.scrollCoeff;
                this._scrollInfo.scroll.Repos(oScrollSettings, false);
                let nScrollCoeff = this.content.ShiftViewY / nMaxShiftY;
                this._scrollInfo.scrollCoeff = nScrollCoeff;

                this._scrollInfo.docElem.style.top      = Math.round(oGlobalCoords1["Y"]) + 'px';
                this._scrollInfo.docElem.style.left     = Math.round(oGlobalCoords2["X"]) + 'px';
                this._scrollInfo.docElem.style.height   = Math.round(oGlobalCoords2["Y"]) - Math.round(oGlobalCoords1["Y"]) + "px";
            
                this._scrollInfo.oldZoom = oViewer.zoom;
                this._scrollInfo.baseYPos = parseInt(this._scrollInfo.docElem.style.top);
            }

            if (bShow === true)
                this._scrollInfo.scroll.canvas.style.display = "";
            if (bShow === false)
                this._scrollInfo.scroll.canvas.style.display = "none";
        }
    };

    CTextField.prototype.SelectionSetStart = function(x, y, e) {
        let oPos    = AscPDF.GetPageCoordsByGlobalCoords(x, y, this.GetPage());
        let X       = oPos["X"];
        let Y       = oPos["Y"];
        this.content.Selection_SetStart(X, Y, 0, e);
    };
    CTextField.prototype.SelectionSetEnd = function(x, y, e) {
        let oPos    = AscPDF.GetPageCoordsByGlobalCoords(x, y, this.GetPage());
        let X       = oPos["X"];
        let Y       = oPos["Y"];
        this.content.Selection_SetEnd(X, Y, 0, e);
    };
    CTextField.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey)
    {
        this.content.MoveCursorLeft(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this.content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey)
    {
        this.content.MoveCursorRight(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this.content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        this.content.MoveCursorDown(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this.content.RecalculateCurPos();
    };
    CTextField.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        this.content.MoveCursorUp(isShiftKey, isCtrlKey);
        this._bAutoShiftContentView = true && this._doNotScroll == false;
        return this.content.RecalculateCurPos();
    };
    CTextField.prototype.EnterText = function(aChars)
    {
        let oDoc = this.GetDocument();
        let oPara = this.content.GetElement(0);
        this.CreateNewHistoryPoint(true);

        let nChars = 0;
        function countChars(oRun) {
            var nCurPos = oRun.Content.length;
			for (var nPos = 0; nPos < nCurPos; ++nPos)
			{
				if (para_Text === oRun.Content[nPos].Type || para_Space === oRun.Content[nPos].Type || para_Tab === oRun.Content[nPos].Type)
                    nChars++;
            }
        }
        // считаем символы в форме
        if (this._charLimit != 0)
            this.content.CheckRunContent(countChars);


        // считаем максимум символов для вставки
        let nSelChars = this.content.GetSelectedText(true, {NewLine: true}).length;
        let nMaxCharsToAdd = Math.min(this._charLimit != 0 ? this._charLimit - (nChars - nSelChars) : aChars.length, aChars.length);
        if (nMaxCharsToAdd > 0) aChars.length = nMaxCharsToAdd;
        else aChars.length = 0;

        if (this.DoKeystrokeAction(aChars) == false) {
            AscCommon.History.Remove_LastPoint();
            return false;
        }

        let nSelStart = oDoc.event["selStart"];
        let nSelEnd = oDoc.event["selEnd"];

        // убираем селект, выставляем из nSelStart/nSelEnd
        if (this.content.IsSelectionUse())
            this.content.RemoveSelection();

        let oDocPos     = this.CalcDocPos(nSelStart, nSelEnd);
        let startPos    = oDocPos.startPos;
        let endPos      = oDocPos.endPos;
        
        if (nSelStart == nSelEnd) {
            this.content.SetContentPosition(startPos, 0, 0);
            this.content.RecalculateCurPos();
        }
        else
            this.content.SetSelectionByContentPositions(startPos, endPos);

        if (nSelStart != nSelEnd)
            this.content.Remove(-1, true, false, false, false);

        this.SetNeedRecalc(true);
        aChars = AscWord.CTextFormFormat.prototype.GetBuffer(oDoc.event["change"]);
        if (aChars.length == 0) {
            return false;
        }

        this.InsertChars(aChars);
        this.SetNeedCommit(true); // флаг что значение будет применено к остальным формам с таким именем
        this._bAutoShiftContentView = true && this._doNotScroll == false;

        if (this._doNotScroll) {
            oPara.Recalculate_Page(0);
            let isOutOfForm = this.IsTextOutOfForm();
            if ((this.IsMultiline() && isOutOfForm.ver) || (isOutOfForm.hor && this.IsMultiline() == false))
                AscCommon.History.Undo();

            this.AddToRedraw();
        }

        return true;
    };
    CTextField.prototype.InsertChars = function(aChars) {
        let oPara = this.content.GetElement(0);

        for (let index = 0; index < aChars.length; ++index) {
            let codePoint = aChars[index];
            if (9 === codePoint) // \t
				oPara.AddToParagraph(new AscWord.CRunTab(), true);
			else if (10 === codePoint || 13 === codePoint) // \n \r
				oPara.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line), true);
			else if (AscCommon.IsSpace(codePoint)) // space
				oPara.AddToParagraph(new AscWord.CRunSpace(codePoint), true);
			else
				oPara.AddToParagraph(new AscWord.CRunText(codePoint), true);
        }
    };
    /**
	 * Checks is text in form is out of form bounds.
     * Note: in vertical case one line always be valid even if form is very short.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
     * @returns {object} - {hor: {boolean}, ver: {boolean}}
	 */
    CTextField.prototype.IsTextOutOfForm = function() {
        let oPageBounds = this.content.GetContentBounds(0);
        let oFormBounds = this.getFormRelRect();
        let oContentW   = this.content.GetElement(0).GetContentWidthInRange();
        
        let oResult = {
            hor: false,
            ver: false
        }

        if (oContentW > oFormBounds.W) {
            oResult.hor = true;
        }

        if (this.IsMultiline() && oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
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
        let aFields     = this._doc.GetFields(this.GetFullName());
        let oThisPara   = this.content.GetElement(0);
        
        if (this.DoFormatAction() == false) {
            this.UndoNotAppliedChanges();
            if (this.IsChanged() == false)
                this.SetDrawFromStream(true);

            return;
        }
        
        this.CorrectHistoryPoints();
        if (true != editor.getDocumentRenderer().isOnUndoRedo) {
            if (this.GetApiValue() != this.GetValue()) {
                if (this.GetDocument().IsNeedSkipHistory() == false) {
                    this.CreateNewHistoryPoint();
                    AscCommon.History.Add(new CChangesPDFFormValue(this, this.GetApiValue(), this.GetValue()));
                }
                
                this.SetApiValue(this.GetValue());
            }
        }

        TurnOffHistory();
        
        if (aFields.length == 1)
            this.SetNeedCommit(false);

        if (oDoc.event["rc"] == false) {
            this.needValidate = true;
            return;
        }

        // устанавливаем дефолтное значение формы
        if (this.GetValue() == "" && this.GetDefaultValue() != null) {
            this.SetValue(this.GetDefaultValue());
            this.SetNeedRecalc(true);
            this.AddToRedraw();
        }

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

            aFields[i].SetApiValue(this.GetApiValue());

            let oFieldPara = aFields[i].content.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oThisPara.Content.length - 1; nItem++) {
                oThisRun = oThisPara.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, oThisRun.Content[nRunPos].Copy());
                }
            }

            aFields[i].SetNeedRecalc(true);
        }

        let oParaFromFormat = this.contentFormat.GetElement(0);
        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] == this)
                continue;

            let oFieldPara = aFields[i].contentFormat.GetElement(0);
            let oThisRun, oFieldRun;
            for (let nItem = 0; nItem < oParaFromFormat.Content.length - 1; nItem++) {
                oThisRun = oParaFromFormat.Content[nItem];
                oFieldRun = oFieldPara.Content[nItem];
                oFieldRun.ClearContent();

                for (let nRunPos = 0; nRunPos < oThisRun.Content.length; nRunPos++) {
                    oFieldRun.AddToContent(nRunPos, oThisRun.Content[nRunPos].Copy());
                }
            }

            aFields[i].SetNeedRecalc(true);
        }

        this.SetNeedCommit(false);
        this.needValidate = true;
    };
	CTextField.prototype.SetAlign = function(nAlignType) {
		this.content.SetAlign(nAlignType);
		if (this.contentFormat)
			this.contentFormat.SetAlign(nAlignType);
		
		this.SetNeedRecalc(true);
	};
	CTextField.prototype.GetAlign = function() {
		return this.content.GetAlign();
	};

    CTextField.prototype.DoFormatAction = function() {
        let oViewer             = editor.getDocumentRenderer();
        let oDoc                = this.GetDocument();
        let oFormatTrigger      = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format);
        let oActionRunScript    = oFormatTrigger ? oFormatTrigger.GetActions()[0] : null;
        
        let isCanFormat = oViewer.isOnUndoRedo == false ? this.DoKeystrokeAction(null, false, true) : true;
        if (!isCanFormat) {
            editor.sendEvent("asc_onFormatErrorPdfForm", oDoc.GetWarningInfo());
            return false;
        }

        if (oActionRunScript) {
            this.SetNeedRecalc(true);
            oActionRunScript.RunScript();
        }

        return true;
    };
    CTextField.prototype.DoKeystrokeAction = function(aChars, nRemoveType, isOnCommit) {
        if (!aChars)
            aChars = [];

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

        let oSelRange   = GetSelectionRange(this.content.GetElement(0));
        let nSelStart   = oSelRange.nSelStart;
        let nSelEnd     = oSelRange.nSelEnd;

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
            "value":    this.GetValue(),
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
            if ((oWarningInfo["greater"] != null || oWarningInfo["less"] != null))
                editor.sendEvent("asc_onValidateErrorPdfForm", oWarningInfo);
            
            return isValid;
        }
        
        return isValid;
    };

    CTextField.prototype.UndoNotAppliedChanges = function() {
        this.UnionLastHistoryPoints();
        let nPoint = AscCommon.History.Index;
        AscCommon.History.Undo();
        
        // удаляем точки
        AscCommon.History.Points.length = nPoint;

        this.SetNeedRecalc(true);
        this.AddToRedraw();
        this.SetNeedCommit(false);
    };

    /**
	 * Unions the last history points of this form.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.UnionLastHistoryPoints = function() {
        let oTmpPoint;
        let oResultPoint = {
            State      : undefined,
            Items      : [],
            Time       : new Date().getTime(),
            Additional : {FormFilling: this, CanUnion: false},
            Description: undefined
        };

        let i = 0;
        for (i = AscCommon.History.Points.length - 1; i >= 0 ; i--) {
            oTmpPoint = AscCommon.History.Points[i];
            if (oTmpPoint.Additional.FormFilling === this && oTmpPoint.Additional.CanUnion != false) {
                oResultPoint.Items = oTmpPoint.Items.concat(oResultPoint.Items);
            }
            else {
                break;
            }
        }

        i++; // индекс точки, в которую поместим результирующее значение.
        
        // если больше 2х точек, то заменяем
        if (i < AscCommon.History.Points.length - 1) {
            AscCommon.History.Index = i;
            AscCommon.History.Points.splice(i, AscCommon.History.Points.length - i, oResultPoint);
        }
        else if (AscCommon.History.Points[i])
            AscCommon.History.Points[i].Additional.CanUnion = false; // запрещаем объединять последнюю добавленную точку
    };
    /**
	 * Используется при применении (Apply метод) формы.
     * Удаляем все будущие точки (данной формы) на момент применения значения формы, а так же удаляем все предыдущие не объединённые точки.
     * Точки которые нельзя удалить (объединить) помечаются canUnion == false.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.CorrectHistoryPoints = function() {
        let oViewer = editor.getDocumentRenderer();

        if (oViewer.isOnUndoRedo != true) {
            let nCurIdx = AscCommon.History.Index;
            let nFromPoint = nCurIdx;
            let nToPoint = nCurIdx;

            if (null == AscCommon.History.Points[nCurIdx] || AscCommon.History.Points[nCurIdx].Additional.CanUnion == false)
                return;

            if (AscCommon.History.Points[nCurIdx]) {
                for (let i = nCurIdx - 1; i >= 0; i--) {
                    if (!AscCommon.History.Points[i] || AscCommon.History.Points[i].Additional.CanUnion === false) {
                        break;
                    }
                    nFromPoint--;
                }
            }

            for (let i = nCurIdx + 1; i < AscCommon.History.Points.length; i++) {
                if (!AscCommon.History.Points[i] || AscCommon.History.Points[i].Additional.CanUnion === false) {
                    break;
                }

                nToPoint++;
            }

            let nNewIdx = nFromPoint - 1;
            AscCommon.History.Points.splice(nFromPoint, Math.max((nToPoint - nFromPoint) + 1));
            AscCommon.History.Index = nNewIdx
        }
        else {
            let nToPoint = AscCommon.History.Index + 1;;

            if (!AscCommon.History.Points[nToPoint] || AscCommon.History.Points[nToPoint].Additional.CanUnion === false) {
                return;
            }

            AscCommon.History.Points.length = nToPoint;
        }
    };
    /**
	 * Removes char in current position by direction.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.Remove = function(nDirection, bWord) {
        let oDoc = this.GetDocument();

        // для combobox
        if (this._editable == false)
            return false;

        this.CreateNewHistoryPoint(true);

        if (this.DoKeystrokeAction(null, nDirection, false) == false) {
            AscCommon.History.Remove_LastPoint();
            return false;
        }
        
        let nSelStart = oDoc.event["selStart"];
        let nSelEnd = oDoc.event["selEnd"];

        if (this.content.IsSelectionUse())
            this.content.RemoveSelection();

        let oDocPos     = this.CalcDocPos(nSelStart, nSelEnd);
        let startPos    = oDocPos.startPos;
        let endPos      = oDocPos.endPos;

        if (nSelStart == nSelEnd) {
            this.content.SetContentPosition(startPos, 0, 0);
            this.content.RecalculateCurPos();
        }
        else
            this.content.SetSelectionByContentPositions(startPos, endPos);

        if (nSelStart != nSelEnd)
            this.content.Remove(nDirection, true, false, false, bWord);
        
        // скрипт keystroke мог поменять change значение, поэтому
        this.InsertChars(AscWord.CTextFormFormat.prototype.GetBuffer(oDoc.event["change"].toString()));

        if (AscCommon.History.Is_LastPointEmpty())
            AscCommon.History.Remove_LastPoint();
        else {
            this.SetNeedRecalc(true);
            this.SetNeedCommit(true);
        }
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CTextField.prototype.SyncField = function() {
        let aFields = this._doc.GetFields(this.GetFullName());
        
        TurnOffHistory();

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
            this._originShiftView = {
                x: this.content.ShiftViewX,
                y: this.content.ShiftViewY
            }
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
                    if (nCursorT + nCursorH > oFormBounds.Y + oFormBounds.H)
                        nDy = oFormBounds.Y + oFormBounds.H - (nCursorT + nCursorH);
                }
                else {
                    if (nCursorT < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                }
            }
            else {
                if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
                    if (oLastLineBounds.Bottom - Math.floor(((oFormBounds.Y + oFormBounds.H) * 1000)) / 1000 < 0)
                        nDy = oFormBounds.Y + oFormBounds.H - oLastLineBounds.Bottom;
                    else if (nCursorT < oFormBounds.Y)
                        nDy = oFormBounds.Y - nCursorT;
                    else if (nCursorT + nCursorH > oFormBounds.Y + oFormBounds.H)
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
	
    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CTextField = CTextField;
})();

