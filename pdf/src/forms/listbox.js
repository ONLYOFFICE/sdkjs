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
    let LISTBOX_SELECTED_COLOR = {
        r: 153,
        g: 193,
        b: 218
    }

    /**
	 * Class representing a listbox field.
	 * @constructor
     * @extends {CBaseListField}
	 */
    function CListBoxField(sName, nPage, aRect, oDoc)
    {
        AscPDF.CBaseListField.call(this, sName, AscPDF.FIELD_TYPES.listbox, nPage, aRect, oDoc);

        this._multipleSelection = false;

        // internal
        this._scrollInfo = null;
        this._bAutoShiftContentView = true;

        this._internalMargins = {
            bottom: undefined
        }
    };
    CListBoxField.prototype = Object.create(AscPDF.CBaseListField.prototype);
	CListBoxField.prototype.constructor = CListBoxField;

    CListBoxField.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        // когда выравнивание посередине или справа, то после того
        // как ширина параграфа будет больше чем размер формы, выравнивание становится слева, пока текста вновь не станет меньше чем размер формы
        this.CheckAlignInternal();
        
        this.Recalculate();
        this.DrawBackground(oGraphicsPDF);
        
        if (this._bAutoShiftContentView || this._bShiftByTopIndex) {
            this.CheckFormViewWindow();
        }
        else {
            this.content.ResetShiftView();
            this.content.ShiftView(this._curShiftView.x, this._curShiftView.y);
        }

        oGraphicsWord.AddClipRect(this.contentClipRect.X, this.contentClipRect.Y, this.contentClipRect.W, this.contentClipRect.H);
        this.content.Draw(0, oGraphicsWord);
        oGraphicsWord.RemoveLastClip();

        this.DrawBorders(oGraphicsPDF);

        this.DrawLocks(oGraphicsPDF);
    };
    CListBoxField.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        if (!this.RecalculateContentRect()) {
            this.content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
        }

        this.SetNeedRecalc(false);
    };
    CListBoxField.prototype.RecalculateContentRect = function() {
        let aOrigRect = this.GetOrigRect();

        let X       = aOrigRect[0];
        let Y       = aOrigRect[1];
        let nWidth  = (aOrigRect[2] - aOrigRect[0]);

        let oMargins = this.GetMarginsFromBorders();

        let contentX        = (X + oMargins.left) * g_dKoef_pt_to_mm;
        let contentY        = (Y + oMargins.top) * g_dKoef_pt_to_mm;
        let contentXLimit   = (X + nWidth - oMargins.left) * g_dKoef_pt_to_mm;

        if (contentX != this.content.X || contentY != this.content.Y ||
        contentXLimit != this.content.XLimit) {
            this.content.X      = contentX;
            this.content.Y      = contentY;
            this.content.XLimit = contentXLimit;
            this.content.YLimit = 20000;

            this.content.Content.forEach(function(para) {
                para.Pr.Ind.FirstLine = oMargins.left * g_dKoef_pt_to_mm;
                para.RecalcCompiledPr(true);
            });

            this.CalculateContentClipRect();
            this.content.Recalculate_Page(0, true);

            return true;
        }

        return false;
    };
    CListBoxField.prototype.CalculateContentClipRect = function() {
        if (!this.content)
            return;

        let aRect       = this.GetOrigRect();
        let X           = aRect[0];
        let Y           = aRect[1];
        let nWidth      = aRect[2] - aRect[0];
        let nHeight     = aRect[3] - aRect[1];
        let oMargins    = this.GetMarginsFromBorders();

        let contentX        = (X + oMargins.left) * g_dKoef_pt_to_mm;
        let contentXLimit   = (X + nWidth - oMargins.left) * g_dKoef_pt_to_mm;

        this.contentClipRect = {
            X: contentX,
            Y: (Y + oMargins.top) * g_dKoef_pt_to_mm,
            W: contentXLimit - contentX,
            H: (nHeight - oMargins.top - oMargins.bottom) * g_dKoef_pt_to_mm,
            Page: this.GetPage()
        }
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.SyncField = function() {
        let aFields = this.GetDocument().GetAllWidgets(this.GetFullName());
        
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        for (let i = 0; i < aFields.length; i++) {
            if (aFields[i] != this) {
                this.SetMultipleSelection(aFields[i].IsMultipleSelection());
                this.content.Internal_Content_RemoveAll();
                for (let nItem = 0; nItem < aFields[i].content.Content.length; nItem++) {
                    this.content.Internal_Content_Add(nItem, aFields[i].content.Content[nItem].Copy());
                }
                
                this._options = aFields[i]._options.slice();
                this._currentValueIndices = aFields.multipleSelection ? aFields[i]._currentValueIndices.slice() : aFields[i]._currentValueIndices;

                let oPara;
                for (let i = 0; i < this.content.Content.length; i++) {
                    oPara = this.content.GetElement(i);
                    if (oPara.Pr.Shd && oPara.Pr.Shd.IsNil() == false)
                        oPara.RecalcCompiledPr(true);
                }
                break;
            }
        }

        oDoc.EndNoHistoryMode();
    };
    /**
	 * Applies value of this field to all fields with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.Commit = function() {
        let oDoc    = this.GetDocument();
        let aFields = oDoc.GetAllWidgets(this.GetFullName());
        
        let aCurIdxs = this.GetCurIdxs();
        let aApiIdxs = this.GetParentCurIdxs();

        this.ScrollVerticalEnd(true);
        let isChanged = false;
        for (let i = 0; i < aCurIdxs.length; i++) {
            if (aCurIdxs[i] === undefined || aApiIdxs[i] === undefined || aCurIdxs[i] !== aApiIdxs[i]) {
                isChanged = true;
                break;
            }
        }
        if (!isChanged) {
            return;
        }

        this.SetTopIndex(undefined);
        if (false == this.IsNeedDrawFromStream()) {
            this.Recalculate();
            this.CheckFormViewWindow();
        }

        for (let i = 0; i < aFields.length; i++) {
            aFields[i].SetWasChanged(true);
            aFields[i].SetNeedRecalc(true);
            aFields[i].SetCurIdxs(aCurIdxs);

            if (aFields[i] !== this) {
                aFields[i].SetTopIndex(undefined);
            }
        }
        
        this._bAutoShiftContentView = true;
        
        this.SetParentValue(this.GetValue());
        this.SetApiCurIdxs(aCurIdxs);
    };
    CListBoxField.prototype.UpdateTopIndex = function() {
        let oParaBounds     = this.content.GetElement(0).GetPageBounds(0);
        let nHeightPerPara  = oParaBounds.Bottom - oParaBounds.Top;
        let nTopIndex       = Math.round(-this._curShiftView.y / nHeightPerPara); // количество смещений в параграфах
        
        AscCommon.History.Add(new CChangesPDFListTopIndex(this, this.GetTopIndex(), nTopIndex));
        this._topIdx = nTopIndex;
    };
    CListBoxField.prototype.GetTopIndex = function() {
        return this._topIdx;
    };
    CListBoxField.prototype.SetTopIndex = function(nTopIndex) {
        // Обновляем _topIdx и добавляем изменение в историю
        AscCommon.History.Add(new CChangesPDFListTopIndex(this, this.GetTopIndex(), nTopIndex));
        this._topIdx = nTopIndex;

        if (false == Asc.editor.getDocumentRenderer().IsOpenFormsInProgress) {
            if (nTopIndex != undefined) {
                this._bAutoShiftContentView = false;
                this._bShiftByTopIndex = true;
            }
            else {
                this._bAutoShiftContentView = true;
                this._bShiftByTopIndex = false;
            }
        }
        
        this.AddToRedraw();
    };
    
    CListBoxField.prototype.SetMultipleSelection = function(bValue) {
        if (bValue == true) {
            this._multipleSelection = true;
        }
        else {
            this._multipleSelection = false;
        }
    };
    CListBoxField.prototype.IsMultipleSelection = function() {
        return this._multipleSelection;
    };

    CListBoxField.prototype.SelectOption = function(nIdx, isSingleSelect) {
        let curIdxs = this.GetCurIdxs();
        if (Array.isArray(curIdxs)) {
            if (curIdxs.includes(nIdx) && curIdxs.length == 1)
                return;
        }
        else if (curIdxs == nIdx)
            return;
            
        let oPara = this.content.GetElement(nIdx);
        let oApiPara;
        
        AscCommon.History.StartNoHistoryMode();

        this.content.Set_CurrentElement(nIdx);
        if (isSingleSelect) {
            this.content.Content.forEach(function(para) {
                oApiPara = editor.private_CreateApiParagraph(para);
                if (para.Pr.Shd && para.Pr.Shd.IsNil() == false) {
                    oApiPara.SetShd('nil');
                    para.RecalcCompiledPr(true);
                }
            });
        }

        if (oPara) {
            oApiPara = editor.private_CreateApiParagraph(oPara);
            oApiPara.SetShd('clear', LISTBOX_SELECTED_COLOR.r, LISTBOX_SELECTED_COLOR.g, LISTBOX_SELECTED_COLOR.b);
            oApiPara.Paragraph.RecalcCompiledPr(true);
        }

        AscCommon.History.EndNoHistoryMode();

        this.SetNeedRecalc(true);
        this.SetNeedCommit(true);
        this.AddToRedraw();
    };
    CListBoxField.prototype.UnselectOption = function(nIdx) {
        let oApiPara = editor.private_CreateApiParagraph(this.content.GetElement(nIdx));
        oApiPara.SetShd('nil');
        oApiPara.Paragraph.RecalcCompiledPr(true);
        this.SetNeedRecalc(true);
    };
    CListBoxField.prototype.SetOptions = function(aOpt) {
        this.content.Internal_Content_RemoveAll();
        for (let i = 0; i < aOpt.length; i++) {
            if (aOpt[i] == null)
                continue;
            let sCaption = "";
            if (typeof(aOpt[i]) == "string" && aOpt[i] != "") {
                sCaption = aOpt[i];
                this._options.push(aOpt[i]);
            }
            else if (Array.isArray(aOpt[i]) && aOpt[i][0] != undefined && aOpt[i][1] != undefined) {
                if (aOpt[i][0].toString && aOpt[i][1].toString) {
                    this._options.push([aOpt[i][0].toString(), aOpt[i][1].toString()]);
                    sCaption = aOpt[i][0].toString();
                }
            }
            else if (typeof(aOpt[i]) != "string" && aOpt[i].toString) {
                this._options.push(aOpt[i].toString());
                sCaption = aOpt[i].toString();
            }

            if (sCaption !== "") {
                AscFonts.FontPickerByCharacter.getFontsByString(sCaption);

                let oPara = new AscWord.Paragraph(this.content, false);
                let oRun = new AscWord.ParaRun(oPara, false);
                this.content.Internal_Content_Add(i, oPara);
                oPara.Add(oRun);
                oRun.AddText(sCaption);
            }
        }
    };
    CListBoxField.prototype.SetValue = function(value) {
        let aIndexes = [];
        if (this.IsWidget()) {
            if (Array.isArray(value)) {
                for (let i = 0; i < value.length; i++) {
                    let sVal = value[i];
                    let isFound = false;
                    for (let i = 0; i < this._options.length; i++) {
                        if (Array.isArray(this._options[i]) && this._options[i][1] == sVal) {
                            if (aIndexes.includes(i))
                                continue;
                            else {
                                isFound = true;
                                aIndexes.push(i);
                                break;
                            }
                        }
                    }
                    if (isFound == false) {
                        for (let i = 0; i < this._options.length; i++) {
                            if (this._options[i] == sVal) {
                                if (aIndexes.includes(i))
                                    continue;
                                else {
                                    aIndexes.push(i);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            else {
                for (let i = 0; i < this._options.length; i++) {
                    if (this._options[i][1] && this._options[i][1] == value) {
                        aIndexes.push(i);
                        break;
                    }
                }
                if (aIndexes.length == 0) {
                    for (let i = 0; i < this._options.length; i++) {
                        if (this._options[i] == value) {
                            aIndexes.push(i);
                            break;
                        }
                    }
                }
            }

            this.content.Content.forEach(function(para) {
                let oApiPara = editor.private_CreateApiParagraph(para);
                if (para.Pr.Shd && para.Pr.Shd.IsNil() == false) {
                    oApiPara.SetShd('nil');
                    para.RecalcCompiledPr(true);
                }
            });

            for (let i = 0; i < aIndexes.length; i++) {
                if (this.IsMultipleSelection()) {
                    this.SelectOption(aIndexes[i], false);
                }
                else
                    this.SelectOption(aIndexes[i], true);
            }

            if (editor.getDocumentRenderer().IsOpenFormsInProgress) {
                this.SetParentValue(value);
                this.SetApiCurIdxs(aIndexes);
            }
                
        }
        else {
            this.SetParentValue(value);
            this.SetApiCurIdxs(aIndexes);
        }
    };
    CListBoxField.prototype.private_SetValue = CListBoxField.prototype.SetValue;
    CListBoxField.prototype.InsertOption = function(sName, sExport, nIdx) {
        let optToInsert = sExport ? [sName, sExport] : sName;
        if (nIdx == -1 || nIdx > this._options.length) {
            nIdx = this._options.length;
        }

        let oIdxItem = this.content.GetElement(nIdx);

        this._options = this._options.splice(nIdx, 0, optToInsert);
        
        let oPara = new AscWord.Paragraph(this.content, false);
        let oRun = new AscWord.ParaRun(oPara, false);
        this.content.Internal_Content_Add(nIdx, oPara);
        oPara.Add(oRun);
        oRun.AddText(sName);
    };

    CListBoxField.prototype.onMouseDown = function(x, y, e) {
        let oDoc            = this.GetDocument();
        let oDrDoc          = oDoc.GetDrawingDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oDrDoc.TargetEnd();

        let isInFocus   = oDoc.activeForm === this;
        let isInForm    = this.IsInForm();
        
        oDoc.activeForm = this;

        function callbackAfterFocus(x, y, e) {
            this.SetInForm(true);
            this.SetDrawHighlight(false);

            if (this._options.length == 0)
                return;

            let oPos    = AscPDF.GetPageCoordsByGlobalCoords(x, y, this.GetPage());
            let X       = oPos["X"];
            let Y       = oPos["Y"];

            let nPos    = this.content.Internal_GetContentPosByXY(X, Y, 0);
            let oPara   = this.content.GetElement(nPos);
            let oShd    = oPara.Pr.Shd;

            this.UpdateScroll(true);
            if (this.IsNeedDrawFromStream() == true) {
                this.SetDrawFromStream(false);
                this.AddToRedraw();
            }
            else if (false == isInForm) {
                this.AddToRedraw();
            }

            if (this.IsMultipleSelection() == true) {
                if (e.CtrlKey == true) {
                    if (oShd && oShd.IsNil() == false) {
                        this.UnselectOption(nPos);
                    }
                    else {
                        this.SelectOption(nPos, false);
                    }
                }
                else {
                    this.SelectOption(nPos, true);
                }
            }
            else {
                this.SelectOption(nPos, true);
            }

            if (this.IsNeedCommit()) {
                this._bAutoShiftContentView = true;
                if (this.IsCommitOnSelChange() == true) {
                    oDoc.EnterDownActiveField();
                }
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
    CListBoxField.prototype.MoveSelectDown = function() {
        this._bAutoShiftContentView = true;
        this.content.MoveCursorDown();

        if (this.IsMultipleSelection() == true) {
            this.SelectOption(this.content.CurPos.ContentPos, true);
        }
        else {
            this.SelectOption(this.content.CurPos.ContentPos, true);
        }
        
        this.AddToRedraw();
        editor.getDocumentRenderer()._paint();
        this.UpdateScroll(true);
    };
    CListBoxField.prototype.MoveSelectUp = function() {
        this._bAutoShiftContentView = true;
        this.content.MoveCursorUp();

        if (this.IsMultipleSelection() == true) {
            this.SelectOption(this.content.CurPos.ContentPos, true);
        }
        else {
            this.SelectOption(this.content.CurPos.ContentPos, true);
        }

        this.AddToRedraw();
        editor.getDocumentRenderer()._paint();
        this.UpdateScroll(true);
    };
    CListBoxField.prototype.UpdateScroll = function(bShow) {
        let oContentBounds  = this.content.GetContentBounds(0);
        let oContentRect    = this.getFormRelRect();
        let aOrigRect       = this.GetOrigRect();

        let nContentH   = oContentBounds.Bottom - oContentBounds.Top;
        let oScrollInfo = this.GetScrollInfo();
        if (bShow == false || nContentH < oContentRect.H) {
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
            
            oScroll.bind("mouseup", function(evt) {
                if (oThis.GetType() == AscPDF.FIELD_TYPES.listbox)
                    oThis.ScrollVerticalEnd();
            });

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
    CListBoxField.prototype.ScrollVertical = function(scrollY, maxYscroll) {
        this._bAutoShiftContentView = false;

        let nScrollCoeff                = scrollY / maxYscroll;
        this._curShiftView.y            = -nScrollCoeff * maxYscroll;
        this._scrollInfo.scrollCoeff    = nScrollCoeff;
        this.AddToRedraw();
    };
    CListBoxField.prototype.ScrollVerticalEnd = function(isOnCommit) {
        if (!this._scrollInfo) {
            return;
        }

        let oParaBounds     = this.content.GetElement(0).GetPageBounds(0);
        let nHeightPerPara  = oParaBounds.Bottom - oParaBounds.Top;
        let nShiftCount     = this._curShiftView.y / nHeightPerPara; // количество смещений в длинах параграфов
        if (Math.abs(Math.round(nShiftCount) - nShiftCount) <= 0.001)
            return;

        let nMaxShiftY                  = this._scrollInfo.scroll.maxScrollY;
        this._curShiftView.y            = Math.round(nShiftCount) * nHeightPerPara;
        this._bAutoShiftContentView     = false;
        this._scrollInfo.scrollCoeff    = Math.abs(this._curShiftView.y / nMaxShiftY);
        
        if (isOnCommit) {
            this.content.ResetShiftView();
            this.content.ShiftView(this._curShiftView.x, this._curShiftView.y);
        }

        this.AddToRedraw();
    };
    CListBoxField.prototype.GetScrollInfo = function() {
        return this._scrollInfo;
    };
    CListBoxField.prototype.SetScrollInfo = function(oInfo) {
        this._scrollInfo = oInfo;
    };
    CListBoxField.prototype.CheckFormViewWindow = function()
    {
        if (this._bShiftByTopIndex) {
            let oParaBounds = this.content.GetElement(0).GetPageBounds(0);
            let nHeightPerPara = oParaBounds.Bottom - oParaBounds.Top;
            
            // Устанавливаем _curShiftView.y по заданному nTopIndex
            this._curShiftView.y = -this.GetTopIndex() * nHeightPerPara;

            this.content.ResetShiftView();
            this.content.ShiftView(this._curShiftView.x, this._curShiftView.y);
            this._oldShiftView = {
                x: this._curShiftView.x,
                y: this._curShiftView.y
            }

            this._bShiftByTopIndex = false;
            this._bAutoShiftContentView = false;
            return;
        }

        let curIdx = this.GetCurIdxs();
        
        let nFirstSelectedPara = 0;
        if (curIdx != null) {
            if (Array.isArray(curIdx)) {
                if (curIdx.length == 0)
                    return;

                nFirstSelectedPara = Number(curIdx[0]);
            }
            else
                nFirstSelectedPara = Number(curIdx);
        }
        
        let oParagraph  = this.content.GetElement(nFirstSelectedPara);

        // размеры всего контента
        let oPageBounds     = this.content.GetContentBounds(0);
        let oCurParaHeight  = oParagraph.Lines[0].Bottom - oParagraph.Lines[0].Top;

        let oFormBounds = this.getFormRelRect();
        let nDy = 0, nDx = 0;
        
        if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
            if (oParagraph.Y + oCurParaHeight - oCurParaHeight * 0.1> oFormBounds.Y + oFormBounds.H) {
                nDy = oFormBounds.Y + oFormBounds.H - (oParagraph.Y + oCurParaHeight) - (this._internalMargins.bottom != undefined ? this._internalMargins.bottom : 0);
                // nDy = oFormBounds.Y + oFormBounds.H - (oParagraph.Y + oCurParaHeight);
            }
                
            else if (oParagraph.Y < oFormBounds.Y)
                nDy = oFormBounds.Y - oParagraph.Y;
            else if (oParagraph.Y + oCurParaHeight < oFormBounds.Y)
                nDy = oCurParaHeight;
        }

        if (Math.abs(nDx) > 0.001 || Math.abs(nDy))
        {
            this.content.ShiftView(nDx, nDy);
            this._oldShiftView = {
                x: this.content.ShiftViewX,
                y: this.content.ShiftViewY
            }
            
            this._curShiftView.x = this._oldShiftView.x;
            this._curShiftView.y = this._oldShiftView.y;
        }
        else {
            this._oldShiftView.x = this._curShiftView.x;
            this._oldShiftView.y = this._curShiftView.y;
        }

        if (nDy == 0) {
            let nCurMarginBottom = this._internalMargins.bottom != undefined ? this._internalMargins.bottom : (oFormBounds.Y + oFormBounds.H) - (oParagraph.Y + oCurParaHeight);
            this._internalMargins.bottom = Math.min(nCurMarginBottom, (oFormBounds.Y + oFormBounds.H) - (oParagraph.Y + oCurParaHeight));
        }

        if (undefined == this.GetTopIndex()) {
            this.UpdateTopIndex();
        }
    };
    /**
	 * Gets real form value (can be not commited).
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.GetValue = function() {
        let oPara, oShd;
        if (this.IsMultipleSelection()) {
            let aValues = [];
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                oPara = this.content.GetElement(i);
                oShd = oPara.Pr.Shd;
                if (oShd && oShd.IsNil() == false)
                    aValues.push(Array.isArray(this._options[i]) ? this._options[i][1] : this._options[i]);
            }

            return aValues;
        }
        else {
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                oPara = this.content.GetElement(i);
                oShd = oPara.Pr.Shd;
                if (oShd && oShd.IsNil() == false)
                    return Array.isArray(this._options[i]) ? this._options[i][1] : this._options[i];
            }
        }
    };
    CListBoxField.prototype.GetCurIdxs = function(bApiValue) {
        if (bApiValue)
            return this._currentValueIndices;
            
        let oPara, oShd;
        let aIndexes = [];
        if (this.IsMultipleSelection()) {
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                oPara = this.content.GetElement(i);
                oShd = oPara.Pr.Shd;
                if (oShd && oShd.IsNil() == false)
                    aIndexes.push(i);
            }
        }
        else {
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                oPara = this.content.GetElement(i);
                oShd = oPara.Pr.Shd;
                if (oShd && oShd.IsNil() == false)
                    aIndexes.push(i);
            }
        }

        return aIndexes;
    };
    CListBoxField.prototype.SetCurIdxs = function(aIdxs) {
        if (this.IsWidget()) {
            let oDoc = this.GetDocument();
            oDoc.History.Add(new CChangesPDFListFormCurIdxs(this, this.GetParentCurIdxs(), aIdxs));

            oDoc.History.StartNoHistoryMode();
            // сначала снимаем выделение с текущих
            let aCurIdxs = this.GetCurIdxs();
            for (let i = 0; i < aCurIdxs.length; i++) {
                this.UnselectOption(aCurIdxs[i]);
            }

            if (aIdxs.length !== 0) {
                this.SelectOption(aIdxs[0], true);
                for (let i = 1; i < aIdxs.length; i++) {
                    this.SelectOption(aIdxs[i]);
                }
            }
            
            oDoc.History.EndNoHistoryMode();
            if (editor.getDocumentRenderer().IsOpenFormsInProgress)
                this.SetApiCurIdxs(aIdxs);
        }
        else {
            this.SetApiCurIdxs(aIdxs);
        }

        this.SetNeedCommit(false);
    };
    CListBoxField.prototype.SetAlign = function(nAlignType) {
        this._alignment = nAlignType;
		this.content.SetAlign(nAlignType);
        this.SetWasChanged(true);
		this.SetNeedRecalc(true);
	};
	CListBoxField.prototype.GetAlign = function() {
		return this._alignment;
	};
    /**
	 * Checks is paragraph text in form is out of form bounds.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
     * @returns {boolean}
	 */
    CListBoxField.prototype.IsParaOutOfForm = function(oPara) {
        if (null == this.getFormRelRect())
            this.Recalculate();
        else
            oPara.Recalculate_Page(0);
        
        let oFormBounds = this.getFormRelRect();
        let nContentW   = oPara.GetContentWidthInRange();
        
        if (nContentW > oFormBounds.W) {
            return true;
        }

        return false;
    };
    CListBoxField.prototype.CheckAlignInternal = function() {
        // когда выравнивание посередине или справа, то после того
        // как ширина параграфа будет больше чем размер формы, выравнивание становится слева, пока текста вновь не станет меньше чем размер формы

        if ([AscPDF.ALIGN_TYPE.center, AscPDF.ALIGN_TYPE.right].includes(this.GetAlign())) {
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                let oPara = this.content.GetElement(i);

                if (this.IsParaOutOfForm(oPara)) {
                    if (getPdfAlignType(oPara.GetParagraphAlign()) != AscPDF.ALIGN_TYPE.left) {
                        oPara.SetParagraphAlign(getInternalAlignType(AscPDF.ALIGN_TYPE.left));
                        this.SetNeedRecalc(true);
                    }
                }
                else if (getPdfAlignType(oPara.GetParagraphAlign()) != this.GetAlign()) {
                    oPara.SetParagraphAlign(getInternalAlignType(this.GetAlign()));
                    this.SetNeedRecalc(true);
                }
            }
        }
    };
    CListBoxField.prototype.WriteToBinary = function(memory) {
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
        
        //
        // top index
        //

        if (this.IsMultipleSelection()) {
            memory.widgetFlags |= (1 << 21);
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

    function getPdfAlignType(nPdfAlign) {
        switch (nPdfAlign) {
			case align_Left: return AscPDF.ALIGN_TYPE.left;
			case align_Center: return AscPDF.ALIGN_TYPE.center;
			case align_Right: return AscPDF.ALIGN_TYPE.right;
		}
    }

    function getInternalAlignType(nInternalAlign) {
        let _alignType = AscCommon.align_Left;
		switch (nInternalAlign) {
			case AscPDF.ALIGN_TYPE.left:
				_alignType = AscCommon.align_Left;
				break;
			case AscPDF.ALIGN_TYPE.center:
				_alignType = AscCommon.align_Center;
				break;
			case AscPDF.ALIGN_TYPE.right:
				_alignType = AscCommon.align_Right;
				break;
		}

        return _alignType;
    }
    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
    window["AscPDF"].CListBoxField      = CListBoxField;
})();

