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
        
        this._formRect.X = X * g_dKoef_pix_to_mm;
        this._formRect.Y = Y * g_dKoef_pix_to_mm;
        this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        this._formRect.H = nHeight * g_dKoef_pix_to_mm;

        this.content.Content.forEach(function(para) {
            para.Pr.Ind.FirstLine   = oMargins.left * g_dKoef_pix_to_mm;
            para.RecalcCompiledPr(true);
        });

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
        contentXLimit != this._oldContentPos.XLimit) {
            this.content.X      = this._oldContentPos.X        = contentX;
            this.content.Y      = this._oldContentPos.Y        = contentY;
            this.content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this.content.YLimit = this._oldContentPos.YLimit   = 20000;
            this.CalculateContentRect();
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

        let aCurIdxs = this.GetCurIdxs();
        if (this.GetApiValue() != this.GetValue()) {
            if (this.GetDocument().IsNeedSkipHistory() == false && true != editor.getDocumentRenderer().isOnUndoRedo) {
                this.CreateNewHistoryPoint();
                AscCommon.History.Add(new CChangesPDFFormValue(this, this.GetApiValue(), this.GetValue()));
                AscCommon.History.Add(new CChangesPDFListFormCurIdxs(this, this.GetApiCurIdxs(), aCurIdxs));
            }
            else if (true == editor.getDocumentRenderer().isOnUndoRedo) {
                // из истории выставляет curIdxs для родительского поля. Это выставление не меняет выделение параграфов.
                // Поэтому вызываем SetCurIdxs
                this._bAutoShiftContentView = true;
                aCurIdxs = this.GetApiCurIdxs();
                this.SetCurIdxs(aCurIdxs);
            }

            this.SetApiValue(this.GetValue());
            this.SetApiCurIdxs(aCurIdxs);
        }
        
        TurnOffHistory();

        aFields.forEach(function(field) {
            field.SetWasChanged(true);
            field.SetNeedRecalc(true);
            if (field.HasShiftView()) {
                if (field == oThis) {
                    field.AddToRedraw();
                    return;
                }
            }

            if (oThis == field)
                return;

            field._bAutoShiftContentView = false;
            field.SetCurIdxs(aCurIdxs);

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
                this.SetApiValue(value);
                this.SetApiCurIdxs(aIndexes);
            }
                
        }
        else {
            this.SetApiValue(value);
            this.SetApiCurIdxs(aIndexes);
        }
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
        let oFormRect       = this.getFormRect();

        let nContentH       = oContentBounds.Bottom - oContentBounds.Top;
        let oScroll, oScrollDocElm, oScrollSettings;

        if (typeof(bShow) != "boolean" && this._scrollInfo)
            bShow = this._scrollInfo.docElem.style.display == "none" ? false : true;

        if (nContentH < oContentRect.H || this._doNotScroll) {
            
            if (this._scrollInfo)
                this._scrollInfo.docElem.style.display = "none";
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
                this._scrollInfo.docElem.style.display = "";
            if (bShow === false)
                this._scrollInfo.docElem.style.display = "none";
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
    
    CListBoxField.prototype.CheckFormViewWindow = function()
    {
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

        if (nDy == 0) {
            let nCurMarginBottom = this._internalMargins.bottom != undefined ? this._internalMargins.bottom : (oFormBounds.Y + oFormBounds.H) - (oParagraph.Y + oCurParaHeight);
            this._internalMargins.bottom = Math.min(nCurMarginBottom, (oFormBounds.Y + oFormBounds.H) - (oParagraph.Y + oCurParaHeight));
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
            // сначала снимаем выделение с текущих
            let aCurIdxs = this.GetCurIdxs();
            for (let i = 0; i < aCurIdxs.length; i++) {
                this.UnselectOption(aCurIdxs[i]);
            }

            this.SelectOption(aIdxs[0], true);
            for (let i = 1; i < aIdxs.length; i++) {
                this.SelectOption(aIdxs[i]);
            }
            if (editor.getDocumentRenderer().IsOpenFormsInProgress)
                this.SetApiCurIdxs(aIdxs);
        }
        else {
            this.SetApiCurIdxs(aIdxs);
        }

        this.SetNeedCommit(false);
    };
	CListBoxField.prototype.UndoNotAppliedChanges = function() {
        this.SetValue(this.GetApiValue());
        this.SetNeedRecalc(true);
        this.AddToRedraw();
        this.SetNeedCommit(false);
    };
    CListBoxField.prototype.SetAlign = function(nAlignType) {
        this._alignment = nAlignType;
		this.content.SetAlign(nAlignType);
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
        if (!this._pagePos)
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

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

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

