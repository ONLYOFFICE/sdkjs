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
    };
    CListBoxField.prototype = Object.create(AscPDF.CBaseListField.prototype);
	CListBoxField.prototype.constructor = CListBoxField;

    CListBoxField.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        this.Recalculate();
        this.DrawBackground(oGraphicsPDF);
        
        if (this._bAutoShiftContentView)
            this.CheckFormViewWindow();
        else {
            this.content.ResetShiftView();
            this.content.ShiftView(this._curShiftView.x, this._curShiftView.y);
        }

        oGraphicsWord.AddClipRect(this.contentRect.X, this.contentRect.Y, this.contentRect.W, this.contentRect.H);
        this.content.Draw(0, oGraphicsWord);
        oGraphicsWord.RemoveLastClip();

        this.DrawBorders(oGraphicsPDF);
    };
    CListBoxField.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        let aRect = this.GetRect();

        let X = aRect[0];
        let Y = aRect[1];
        let nWidth = (aRect[2] - aRect[0]);
        let nHeight = (aRect[3] - aRect[1]);

        // save pos in page.
        this._pagePos = {
            x: X,
            y: Y,
            w: nWidth,
            h: nHeight
        };

        let oMargins = this.GetMarginsFromBorders(false, false);

        let contentX        = (X + oMargins.left) * g_dKoef_pix_to_mm;
        let contentY        = (Y + oMargins.top) * g_dKoef_pix_to_mm;
        let contentXLimit   = (X + nWidth - oMargins.left) * g_dKoef_pix_to_mm;
        let contentYLimit   = (Y + nHeight - oMargins.bottom) * g_dKoef_pix_to_mm;
        
        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;

        this.contentRect.X = contentX;
        this.contentRect.Y = contentY;
        this.contentRect.W = contentXLimit - contentX;
        this.contentRect.H = contentYLimit - contentY;

        this.content.Content.forEach(function(para) {
            para.Pr.Ind.FirstLine = oMargins.left * g_dKoef_pix_to_mm;
            para.RecalcCompiledPr(true);
        });

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this.content.X      = this._oldContentPos.X        = contentX;
            this.content.Y      = this._oldContentPos.Y        = contentY;
            this.content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this.content.YLimit = this._oldContentPos.YLimit   = 20000;
            this.content.Recalculate_Page(0, true);
        }
        else if (this.IsNeedRecalc()) {
            this.content.Content.forEach(function(element) {
                element.Recalculate_Page(0);
            });
        }

        this.SetNeedRecalc(false);
    };

    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.SyncField = function() {
        let aFields = this._doc.GetFields(this.GetFullName());
        
        TurnOffHistory();

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
    };
    /**
	 * Applies value of this field to all field with the same name.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
	 */
    CListBoxField.prototype.Commit = function() {
        let aFields = this._doc.GetFields(this.GetFullName());
        let oThis = this;
        
        this.CheckFormViewWindow();
        this.RevertContentViewToOriginal();

        let oThisBounds = this.getFormRelRect();

        if (true != editor.getDocumentRenderer().isOnUndoRedo && this.GetApiValue() != this.GetValue()) {
            if (this.GetDocument().IsNeedSkipHistory() == false) {
                this.CreateNewHistoryPoint();
                AscCommon.History.Add(new CChangesPDFFormValue(this, this.GetApiValue(), this.GetValue()));
            }
            this.UpdateIndexies();
            this.SetApiValue(this.GetValue());
        }
        
        TurnOffHistory();

        aFields.forEach(function(field) {
            field.SetWasChanged(true);

            field._bAutoShiftContentView = false;

            field.SetNeedRecalc(true);
            if (field.HasShiftView()) {
                if (field == oThis) {
                    field.AddToRedraw();
                    return;
                }
            }

            field.SetApiValue(oThis.GetApiValue());

            if (oThis == field)
                return;

            if (oThis.IsMultipleSelection()) {
                // снимаем выделение с тех, которые не присутсвуют в поле, от которого применяем ко всем
                for (let i = 0; i < field._currentValueIndices.length; i++) {
                    if (oThis._currentValueIndices.includes(field._currentValueIndices[i]) == false) {
                        field.UnselectOption(field._currentValueIndices[i]);
                    }
                }
                
                for (let i = 0; i < oThis._currentValueIndices.length; i++) {
                    // добавляем выделение тем, которые не присутсвуют в текущем поле, но присутсвуют в том, от которого применяем
                    if (field._currentValueIndices.includes(oThis._currentValueIndices[i]) == false) {
                        field.SelectOption(oThis._currentValueIndices[i], false);
                    }
                }
                field._currentValueIndices = oThis._currentValueIndices.slice();
            }
            else {
                field._currentValueIndices = oThis._currentValueIndices;
                field.SelectOption(field._currentValueIndices, true);
            }

            let oFieldBounds = field.getFormRelRect();
            if (Math.abs(oFieldBounds.H - oThisBounds.H) > 0.001) {
                field._bAutoShiftContentView = true;
            }
            else {
                field._curShiftView.x = oThis._curShiftView.x;
                field._curShiftView.y = oThis._curShiftView.y;
                field._originShiftView.x = oThis._originShiftView.x;
                field._originShiftView.y = oThis._originShiftView.y;
            }
        });
    };
    
    CListBoxField.prototype.SetMultipleSelection = function(bValue) {
        if (bValue == true) {
            this._multipleSelection = true;
            this._currentValueIndices = [this._currentValueIndices];
        }
        else {
            this._multipleSelection = false;
            this._currentValueIndices = this._currentValueIndices[0];
            this._currentValueIndices != -1 && this.SelectOption(this._currentValueIndices, true);
        }
    };
    CListBoxField.prototype.IsMultipleSelection = function() {
        return this._multipleSelection;
    };
    CListBoxField.prototype.SelectOption = function(nIdx, isSingleSelect) {
        if (this.GetCurIdxs() == nIdx)
            return;
            
        let oPara = this.content.GetElement(nIdx);
        let oApiPara;
        
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
                let oPara = new AscCommonWord.Paragraph(this.content.DrawingDocument, this.content, false);
                let oRun = new AscCommonWord.ParaRun(oPara, false);
                this.content.Internal_Content_Add(i, oPara);
                oPara.Add(oRun);
                oRun.AddText(sCaption);
            }
        }
    };
    CListBoxField.prototype.SetValue = function(value) {
        if (this.IsWidget()) {
            let aIndexes = [];
            if (Array.isArray(value)) {
                for (let i = 0; i < value.length; i++) {
                    let sVal = value[i];
                    let isFound = false;
                    for (let i = 0; i < this._options.length; i++) {
                        if (this._options[i][1] && this._options[i][1] == sVal) {
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

            this._currentValueIndices = this.IsMultipleSelection() == true ? aIndexes : aIndexes[0];
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

            if (editor.getDocumentRenderer().IsOpenFormsInProgress)
                this.SetApiValue(value);
        }
        else
            this.SetApiValue(value);
    };
    CListBoxField.prototype.InsertOption = function(sName, sExport, nIdx) {
        let optToInsert = sExport ? [sName, sExport] : sName;
        if (nIdx == -1 || nIdx > this._options.length) {
            nIdx = this._options.length;
        }

        let oIdxItem = this.content.GetElement(nIdx);

        this._options = this._options.splice(nIdx, 0, optToInsert);
        
        let oPara = new AscCommonWord.Paragraph(this.content.DrawingDocument, this.content, false);
        let oRun = new AscCommonWord.ParaRun(oPara, false);
        this.content.Internal_Content_Add(nIdx, oPara);
        oPara.Add(oRun);
        oRun.AddText(sName);
    };

    CListBoxField.prototype.onMouseDown = function(x, y, e) {
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        function callbackAfterFocus(x, y, e) {
            if (this._options.length == 0)
                return;

            let bHighlight = this.IsNeedDrawHighlight();
            this.SetDrawHighlight(false);

            let oPos    = AscPDF.GetPageCoordsByGlobalCoords(x, y, this.GetPage());
            let X       = oPos["X"];
            let Y       = oPos["Y"];

            editor.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
            editor.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;

            let nPos = this.content.Internal_GetContentPosByXY(X, Y, 0);
            let oPara = this.content.GetElement(nPos);
            let oShd = oPara.Pr.Shd;

            this.UpdateScroll(true);
            if (this.IsNeedDrawFromStream() == true) {
                this.SetDrawFromStream(false);
                this.AddToRedraw();
            }
            else if (bHighlight) {
                this.AddToRedraw();
            }

            if (this.IsMultipleSelection() == true) {
                if (e.ctrlKey == true) {
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
                this.UnionLastHistoryPoints(false);

                if (this.IsCommitOnSelChange() == true) {
                    this.Commit();
                }
            }

            oDoc.activeForm = this;
        }

        // вызываем выставление курсора после onFocus, если уже в фокусе, тогда сразу.
        if (oDoc.activeForm != this && this._triggers.OnFocus && this._triggers.OnFocus.Actions.length > 0)
            oActionsQueue.callBackAfterFocus = callbackAfterFocus.bind(this, x, y, e);
        else
            callbackAfterFocus.bind(this, x, y, e)();

        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseDown);
        if (oDoc.activeForm != this)
            this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);
    };
    CListBoxField.prototype.MoveSelectDown = function() {
        this._bAutoShiftContentView = true;
        this.content.MoveCursorDown();

        if (this.IsMultipleSelection() == true) {
            this.SelectOption(this.content.CurPos.ContentPos, true);
            this._currentValueIndices = [this.content.CurPos.ContentPos];
        }
        else {
            this.SelectOption(this.content.CurPos.ContentPos, true);
            this._currentValueIndices = this.content.CurPos.ContentPos;
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
            this._currentValueIndices = [this.content.CurPos.ContentPos];
        }
        else {
            this.SelectOption(this.content.CurPos.ContentPos, true);
            this._currentValueIndices = this.content.CurPos.ContentPos;
        }

        this.AddToRedraw();
        editor.getDocumentRenderer()._paint();
        this.UpdateScroll(true);
    };
    CListBoxField.prototype.UpdateScroll = function(bShow) {
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
    CListBoxField.prototype.ScrollVertical = function(scrollY, maxYscroll) {
        this._bAutoShiftContentView = false;

        let nScrollCoeff                = scrollY / maxYscroll;
        this._curShiftView.y            = -nScrollCoeff * maxYscroll;
        this._scrollInfo.scrollCoeff    = nScrollCoeff;
        this.AddToRedraw();
        editor.getDocumentRenderer()._paint();
    };
    CListBoxField.prototype.ScrollVerticalEnd = function() {
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
    /**
	 * Checks curValueIndices, corrects it and return.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CListBoxField.prototype.UpdateIndexies = function() {
        let nIdx;
        if (this.IsMultipleSelection())
            nIdx = []
        else
            nIdx = -1;

        let oCurPara;
        for (let i = 0; i < this.content.Content.length; i++) {
            oCurPara = this.content.GetElement(i);
            if (oCurPara.Pr.Shd && oCurPara.Pr.Shd.IsNil() == false) {
                if (this.IsMultipleSelection())
                    nIdx.push(i);
                else {
                    nIdx = i;
                    break;
                }
            }
        }

        this._currentValueIndices = nIdx;
        return nIdx;
    };
    /**
	 * Updates the selections according to the indices.
	 * @memberof CListBoxField
	 * @typeofeditors ["PDF"]
     * @returns {number}
	 */
    CListBoxField.prototype.UpdateSelection = function() {
        if (Array.isArray(this._currentValueIndices)) {
            for (let i = 0; i < this._currentValueIndices.length; i++) {
                let nIdx = this._currentValueIndices[i];
                this.SelectOption(nIdx, false);
            }
        }
        else
            this.SelectOption(this._currentValueIndices, true);
    };
    
    CListBoxField.prototype.CheckFormViewWindow = function()
    {
        let curIdx = this.GetCurIdxs();
        
        let nFirstSelectedPara = 0;
        if (curIdx != null) {
            nFirstSelectedPara = Array.isArray(curIdx) ? Number(curIdx[0]) : Number(curIdx)
        }
        
        let oParagraph  = this.content.GetElement(nFirstSelectedPara);

        // размеры всего контента
        let oPageBounds     = this.content.GetContentBounds(0);
        let oCurParaHeight  = oParagraph.Lines[0].Bottom - oParagraph.Lines[0].Top;

        let oFormBounds = this.getFormRelRect();
        let nDy = 0, nDx = 0;
        
        if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H) {
            if (oParagraph.Y + oCurParaHeight > oFormBounds.Y + oFormBounds.H)
                nDy = oFormBounds.Y + oFormBounds.H - (oParagraph.Y + oCurParaHeight);
            else if (oParagraph.Y < oFormBounds.Y)
                nDy = oFormBounds.Y - oParagraph.Y;
            else if (oParagraph.Y + oCurParaHeight < oFormBounds.Y)
                nDy = oCurParaHeight;
        }

        if (Math.abs(nDx) > 0.001 || Math.abs(nDy))
        {
            this.content.ShiftView(nDx, nDy);
            this._originShiftView = {
                x: this.content.ShiftViewX,
                y: this.content.ShiftViewY
            }
            
            this._curShiftView.x = this._originShiftView.x;
            this._curShiftView.y = this._originShiftView.y;
        }
        else {
            this._originShiftView.x = this._curShiftView.x;
            this._originShiftView.y = this._curShiftView.y;
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
            let aIndexes = [];
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                oPara = this.content.GetElement(i);
                oShd = oPara.Pr.Shd;
                if (oShd && oShd.IsNil() == false)
                    aIndexes.push(Array.isArray(this._options[i]) ? this._options[i][1] : this._options[i]);
            }

            return aIndexes;
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
        if (this.IsMultipleSelection()) {
            let aIndexes = [];
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                oPara = this.content.GetElement(i);
                oShd = oPara.Pr.Shd;
                if (oShd && oShd.IsNil() == false)
                    aIndexes.push(i);
            }

            return aIndexes;
        }
        else {
            for (let i = 0, nCount = this.content.GetElementsCount(); i < nCount; i++) {
                oPara = this.content.GetElement(i);
                oShd = oPara.Pr.Shd;
                if (oShd && oShd.IsNil() == false)
                    return i;
            }
        }
    };
	CListBoxField.prototype.UndoNotAppliedChanges = function() {
        this.SetValue(this.GetApiValue());
        this.SetNeedRecalc(true);
        this.AddToRedraw();
        this.SetNeedCommit(false);
    };

    CListBoxField.prototype.WriteToBinary = function(memory) {
        // TODO
		/*
		if (шрифт_у_CTextField_как-либо_изменялся)
		{
			// Также как функция SetFont в common/Drawings/Metafile.js
			
			if (шрифт_изменялся)
			{
				memory.WriteByte(AscCommon.CommandType.ctFontName);
				memory.WriteString(this.m_oFont.Name);
			}
			
			if (размер_шрифта_изменялся)
			{
				memory.WriteByte(AscCommon.CommandType.ctFontSize);
				memory.WriteDouble(this.m_oFont.FontSize);
			}
			
			if (стиль_шрифта_изменялся)
			{
				memory.WriteByte(AscCommon.CommandType.ctFontStyle);
				memory.WriteLong(style);
			}
		}
		*/

        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        // длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);

        let sValue = this.GetValue();
        if (sValue != null) {
            memory.fieldFlags2 |= (1 << 9);
            memory.WriteString(sValue);
        }

        let aOptions = this.GetOptions();
        if (aOptions && aOptions.length != 0) {
            memory.WriteLong(aOptions.length);
            for (let i = 0; i < aOptions.length; i++) {
                memory.WriteString(aOptions[i][1] != undefined ? aOptions[i][1] : "");
                memory.WriteString(aOptions[i][0] != undefined ? aOptions[i][0] : "");
            }
        }
        
        //
        // top index
        //

        if (this.IsMultipleSelection()) {
            memory.fieldFlags1 |= (1 << 21);
        }

        // if (this.IsDoNotSpellCheck()) {
        //     memory.fieldFlags2 |= (1 << 22);
        // }

        if (this.IsCommitOnSelChange()) {
            memory.fieldFlags1 |= (1 << 26);
        }

        let nEndPos = memory.GetCurPosition();

        // запись флагов
        memory.Seek(memory.posForFlags1);
        memory.WriteLong(memory.fieldFlags1);
        memory.Seek(memory.posForFlags2);
        memory.WriteLong(memory.fieldFlags2);

        // запись длины комманд
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CListBoxField = CListBoxField;
})();

