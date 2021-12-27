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

"use strict";

(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
    function(window, undefined) {

    // Import
    var isRealObject = AscCommon.isRealObject;
    var History = AscCommon.History;

//-----------------------------

    AscDFH.changesFactory[AscDFH.historyitem_TextBodySetParent] = AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_TextBodySetBodyPr] = AscDFH.CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_TextBodySetContent] = AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_TextBodySetLstStyle] = AscDFH.CChangesDrawingsObjectNoId;

    AscDFH.drawingsChangesMap[AscDFH.historyitem_TextBodySetParent] = function(oClass, value) {
        oClass.parent = value;
    };
    AscDFH.drawingsChangesMap[AscDFH.historyitem_TextBodySetBodyPr] = function(oClass, value) {
        if(CheckNeedRecalcAutoFit(oClass.bodyPr, value))
            if(oClass.parent) {
                oClass.parent.recalcInfo.recalculateContent = true;
                oClass.parent.recalcInfo.recalculateContent2 = true;
                oClass.parent.recalcInfo.recalculateTransformText = true;
            }
        if(oClass.content) {
            oClass.content.Recalc_AllParagraphs_CompiledPr();
        }
        oClass.bodyPr = value;
        oClass.recalcInfo.recalculateBodyPr = true;
    };
    AscDFH.drawingsChangesMap[AscDFH.historyitem_TextBodySetContent] = function(oClass, value) {
        oClass.content = value;
    };
    AscDFH.drawingsChangesMap[AscDFH.historyitem_TextBodySetLstStyle] = function(oClass, value) {
        oClass.lstStyle = value;
    };
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_TextBodySetBodyPr] = AscFormat.CBodyPr;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_TextBodySetLstStyle] = AscFormat.TextListStyle;


    var InitClass = AscFormat.InitClass;
    var CBaseFormatObject = AscFormat.CBaseFormatObject;

    function CTextBody() {
        CBaseFormatObject.call(this);
        this.bodyPr = null;
        this.lstStyle = null;
        this.content = null;
        this.parent = null;

        this.content2 = null;
        this.compiledBodyPr = null;
        this.parent = null;
        this.bFit = true;
        this.recalcInfo =
        {
            recalculateBodyPr: true,
            recalculateContent2: true
        };
    }
    InitClass(CTextBody, CBaseFormatObject, AscDFH.historyitem_type_TextBody);
    CTextBody.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.bodyPr)
            oCopy.setBodyPr(this.bodyPr.createDuplicate());
        if(this.lstStyle)
            oCopy.setLstStyle(this.lstStyle.createDuplicate());
        if(this.content)
            oCopy.setContent(this.content.Copy(oCopy));
    };
    CTextBody.prototype.createDuplicate2 = function() {
        var ret = new CTextBody();
        if(this.bodyPr)
            ret.setBodyPr(this.bodyPr.createDuplicate());
        if(this.lstStyle)
            ret.setLstStyle(this.lstStyle.createDuplicate());
        if(this.content) {
            var bTrackRevision = false;
            if(typeof editor !== "undefined" && editor && editor.WordControl && editor.WordControl.m_oLogicDocument.TrackRevisions === true) {
                bTrackRevision = editor.WordControl.m_oLogicDocument.GetLocalTrackRevisions();
                editor.WordControl.m_oLogicDocument.SetLocalTrackRevisions(false);
            }
            ret.setContent(this.content.Copy3(ret));
            if(false !== bTrackRevision) {
                editor.WordControl.m_oLogicDocument.SetLocalTrackRevisions(bTrackRevision);
            }
        }
        return ret;
    };
    CTextBody.prototype.createDuplicateForSmartArt = function (oPr) {
        var arrayOfTextBody = [];
        for (var i = 0; i < oPr.pointContentLength; i += 1) {
            arrayOfTextBody.push(new CTextBody());
        }
        var that = this;
        arrayOfTextBody.forEach(function (txBody) {
            if(that.bodyPr)
                txBody.setBodyPr(that.bodyPr.createDuplicateForSmartArt(oPr));
            if(that.lstStyle)
                txBody.setLstStyle(that.lstStyle.createDuplicate());
        })
        if(this.content) {
            this.content.createDuplicateForSmartArt(oPr, arrayOfTextBody);
        }
        return arrayOfTextBody;
    }

    CTextBody.prototype.Is_TopDocument = function() {
        return false;
    };
    CTextBody.prototype.Is_DrawingShape = function(bRetShape) {
        if(bRetShape === true) {
            return this.parent;
        }
        return true;
    };
    CTextBody.prototype.IsInTable = function(bReturnTopTable) {
        if(true === bReturnTopTable)
            return null;

        return false;
    };
    CTextBody.prototype.Get_Theme = function() {
        if(this.parent) {
            return this.parent.Get_Theme();
        }
        return null;
    };
    CTextBody.prototype.Get_ColorMap = function() {
        if(this.parent) {
            return this.parent.Get_ColorMap();
        }
        return null;
    };
    CTextBody.prototype.setBodyPr = function(pr) {
        History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_TextBodySetBodyPr, this.bodyPr, pr));
        this.bodyPr = pr;
        var oParent = this.parent;
        if(oParent) {
            if(oParent.recalcInfo) {
                oParent.recalcInfo.recalculateContent = true;
                oParent.recalcInfo.recalculateContent2 = true;
                oParent.recalcInfo.recalculateTransformText = true;
                if(this.content) {
                    this.content.Recalc_AllParagraphs_CompiledPr();
                }
                if(oParent.addToRecalculate) {
                    oParent.addToRecalculate();
                }
            }
            if(oParent.onChartInternalUpdate) {
                oParent.onChartInternalUpdate();
            }
        }
    };
    CTextBody.prototype.setContent = function(pr) {
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_TextBodySetContent, this.content, pr));
        this.content = pr;
    };
    CTextBody.prototype.setLstStyle = function(lstStyle) {
        History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_TextBodySetLstStyle, this.lstStyle, lstStyle));
        this.lstStyle = lstStyle;
    };
    CTextBody.prototype.recalculate = function() {
    };
    CTextBody.prototype.recalculateBodyPr = function() {
        AscFormat.ExecuteNoHistory(function() {
            if(!this.compiledBodyPr)
                this.compiledBodyPr = new AscFormat.CBodyPr();
            this.compiledBodyPr.setDefault();
            if(this.parent && this.parent.isPlaceholder && this.parent.isPlaceholder()) {
                var hierarchy = this.parent.getHierarchy();
                for(var i = hierarchy.length - 1; i > -1; --i) {
                    if(isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].txBody) && isRealObject(hierarchy[i].txBody.bodyPr))
                        this.compiledBodyPr.merge(hierarchy[i].txBody.bodyPr)
                }
            }
            if(isRealObject(this.bodyPr)) {
                this.compiledBodyPr.merge(this.bodyPr);
            }
        }, this, []);
    };
    CTextBody.prototype.Check_AutoFit = function() {
        return this.parent && this.parent.Check_AutoFit && this.parent.Check_AutoFit(true) || false;
    };
    CTextBody.prototype.Refresh_RecalcData = function(Data) {
        if(this.parent && this.parent.recalcInfo) {
            this.parent.recalcInfo.recalculateContent = true;
            this.parent.recalcInfo.recalculateContent2 = true;
            this.parent.recalcInfo.recalculateTransformText = true;

            if(this.parent.addToRecalculate) {
                this.parent.addToRecalculate();
            }
        }
        if(AscCommon.isRealObject(Data)) {
            if(Data.Type === AscDFH.historyitem_TextBodySetBodyPr) {
                this.recalcInfo.recalculateBodyPr = true;
            }
        }
    };
    CTextBody.prototype.isEmpty = function() {
        return this.content.Is_Empty();
    };
    CTextBody.prototype.OnContentReDraw = function() {
        if(this.parent && this.parent.OnContentReDraw) {
            this.parent.OnContentReDraw();
        }
    };
    CTextBody.prototype.Get_StartPage_Absolute = function() {
        return 0;//TODO;
    };
    CTextBody.prototype.Get_AbsolutePage = function(CurPage) {
        if(this.parent && this.parent.Get_AbsolutePage) {
            return this.parent.Get_AbsolutePage();
        }
        return 0;//TODO;
    };
    CTextBody.prototype.Get_AbsoluteColumn = function(CurPage) {
        return 0;//TODO;
    };
    CTextBody.prototype.Get_TextBackGroundColor = function() {
        if(this.parent && this.parent.Get_TextBackGroundColor) {
            return this.parent.Get_TextBackGroundColor();
        }
        return undefined;
    };
    CTextBody.prototype.IsHdrFtr = function(bReturnHdrFtr) {
        if(bReturnHdrFtr)
            return null;

        return false;
    };
    CTextBody.prototype.IsFootnote = function(bReturnFootnote) {
        if(bReturnFootnote)
            return null;

        return false;
    };
    CTextBody.prototype.Get_PageContentStartPos = function(pageNum) {
        return {X: 0, Y: 0, XLimit: this.contentWidth, YLimit: 20000};
    };
    CTextBody.prototype.Get_Numbering = function() {
        return new CNumbering();
    };
    CTextBody.prototype.Set_CurrentElement = function(bUpdate, pageIndex) {
        if(this.parent.Set_CurrentElement) {
            this.parent.Set_CurrentElement(bUpdate, pageIndex);
        }
    };
    CTextBody.prototype.checkDocContent = function() {
        this.parent && this.parent.checkDocContent && this.parent.checkDocContent();
    };
    CTextBody.prototype.getBodyPr = function() {
        if(this.recalcInfo.recalculateBodyPr) {
            this.recalculateBodyPr();
            this.recalcInfo.recalculateBodyPr = false;
        }
        return this.compiledBodyPr;
    };
    CTextBody.prototype.getSummaryHeight = function() {
        return this.content.GetSummaryHeight();
    };
    CTextBody.prototype.getSummaryHeight2 = function() {
        return this.content2 ? this.content2.GetSummaryHeight() : 0;
    };
    CTextBody.prototype.getSummaryHeight3 = function() {
        if(this.content && this.content.GetSummaryHeight_) {
            return this.content.GetSummaryHeight_();
        }
        return 0;
    };
    CTextBody.prototype.getCompiledBodyPr = function() {
        this.recalculateBodyPr();
        return this.compiledBodyPr;
    };
    CTextBody.prototype.Get_TableStyleForPara = function() {
        return null;
    };
    CTextBody.prototype.checkCurrentPlaceholder = function() {
        return false;
    };
    CTextBody.prototype.draw = function(graphics) {
        if((!this.content || this.content.Is_Empty()) && !AscCommon.IsShapeToImageConverter && this.parent.isEmptyPlaceholder() && !this.checkCurrentPlaceholder()) {
            if(graphics.IsNoDrawingEmptyPlaceholder !== true && graphics.IsNoDrawingEmptyPlaceholderText !== true && this.content2) {
                if(graphics.IsNoSupportTextDraw) {
                    var _w2 = this.content2.XLimit;
                    var _h2 = this.content2.GetSummaryHeight();
                    graphics.rect(this.content2.X, this.content2.Y, _w2, _h2);
                }

                this.content2.Set_StartPage(0);
                this.content2.Draw(0, graphics);
            }
        }
        else if(this.content) {
            if(graphics.IsNoSupportTextDraw) {
                var bEmpty = this.content.IsEmpty();
                var _w = bEmpty ? 0.1 : this.content.XLimit;
                var _h = this.content.GetSummaryHeight();
                graphics.rect(this.content.X, this.content.Y, _w, _h);
            }
            var old_start_page = this.content.StartPage;
            this.content.Set_StartPage(0);
            this.content.Draw(0, graphics);
            this.content.Set_StartPage(old_start_page);
        }
    };
    CTextBody.prototype.Get_Styles = function(level) {
        if(this.parent && this.parent.getStyles) {
            return this.parent.getStyles(level);
        }
        return AscFormat.ExecuteNoHistory(function() {
            var oStyles = new CStyles(false);
            var Style_Para_Def = new CStyle("Normal", null, null, styletype_Paragraph);
            Style_Para_Def.CreateNormal();
            oStyles.Default.Paragraph = oStyles.Add(Style_Para_Def);
            return {styles: oStyles, lastId: oStyles.Default.Paragraph};
        }, this, []);
    };
    CTextBody.prototype.IsCell = function(isReturnCell) {
        if(true === isReturnCell)
            return null;

        return false;
    };
    CTextBody.prototype.OnContentRecalculate = function() {
    };
    CTextBody.prototype.getMargins = function() {
        var _parent_transform = this.parent.transform;
        var _l;
        var _r;
        var _b;
        var _t;
        var _body_pr = this.getBodyPr();
        var sp = this.parent;
        var _rect = sp.getTextRect && sp.getTextRect();
        if(_rect) {
            _l = _rect.l + _body_pr.lIns;
            _t = _rect.t + _body_pr.tIns;
            _r = _rect.r - _body_pr.rIns;
            _b = _rect.b - _body_pr.bIns;
        }
        else {
            _l = _body_pr.lIns;
            _t = _body_pr.tIns;
            _r = sp.extX - _body_pr.rIns;
            _b = sp.extY - _body_pr.bIns;
        }

        var x_lt, y_lt, x_rb, y_rb;

        x_lt = _parent_transform.TransformPointX(_l, _t);
        y_lt = _parent_transform.TransformPointY(_l, _t);

        x_rb = _parent_transform.TransformPointX(_r, _b);
        y_rb = _parent_transform.TransformPointY(_r, _b);

        var hc = (_r - _l) / 2;
        var vc = (_b - _t) / 2;

        var xc = (x_lt + x_rb) / 2;
        var yc = (y_lt + y_rb) / 2;

        return {L: xc - hc, T: yc - vc, R: xc + hc, B: yc + vc, textMatrix: this.parent.transform};
    };
    CTextBody.prototype.Refresh_RecalcData2 = function(pageIndex) {
        this.parent && this.parent.Refresh_RecalcData2 && this.parent.Refresh_RecalcData2(pageIndex, this);
    };
    CTextBody.prototype.checkContentFit = function(sText) {
        var oContent = this.content;
        if(!oContent.Is_Empty()) {
            var oFirstPara = oContent.Content[0];
            oFirstPara.Content = [oFirstPara.Content[oFirstPara.Content.length - 1]];
        }
        AscFormat.AddToContentFromString(oContent, sText);
        AscFormat.CShape.prototype.recalculateContent.call(this.parent);
        var oFirstParagraph = oContent.Content[0];
        return oFirstParagraph.Lines.length === 1;
    };
    CTextBody.prototype.recalculateOneString = function(sText) {
        if(this.checkContentFit(sText)) {
            this.bFit = true;
            this.fitWidth = this.content.Content[0].Lines[0].Ranges[0].W;
            return;
        }
        this.fitWidth = null;
        this.bFit = false;
        var nLeftPos = 0, nRightPos = sText.length;
        var nMiddlePos;
        var sEnd = "...";
        var sFitText = sText + sEnd;

        while(nRightPos - nLeftPos > 1) {
            nMiddlePos = (nRightPos + nLeftPos) / 2 + 0.5 >> 0;
            sFitText = sText.slice(0, nMiddlePos - 1);
            sFitText += sEnd;
            if(!this.checkContentFit(sFitText)) {
                nRightPos = nMiddlePos;
            }
            else {
                nLeftPos = nMiddlePos;
            }
        }
        sFitText = sText.slice(0, nLeftPos - 1);
        sFitText += sEnd;
        if(!this.checkContentFit(sFitText)) {
            var bFound = false;
            for(var i = sEnd.length - 1; i > -1; i--) {
                sFitText = sEnd.slice(0, i);
                if(this.checkContentFit(sFitText)) {
                    bFound = true;
                    break;
                }
            }
            if(!bFound) {
                this.checkContentFit("");
            }
        }
    };
    CTextBody.prototype.getContentOneStringSizes = function() {
        return GetContentOneStringSizes(this.content);
    };
    CTextBody.prototype.recalculateByMaxWord = function() {
        var max_content = this.content.RecalculateMinMaxContentWidth().Max;
        this.content.SetApplyToAll(true);
        this.content.SetParagraphAlign(AscCommon.align_Center);
        this.content.SetApplyToAll(false);
        this.content.Reset(0, 0, max_content, 20000);
        this.content.Recalculate_Page(0, true);
        return {w: max_content, h: this.content.GetSummaryHeight()};
    };
    CTextBody.prototype.GetFirstElementInNextCell = function() {
        return null;
    };
    CTextBody.prototype.GetLastElementInPrevCell = function() {
        return null;
    };
    CTextBody.prototype.getRectWidth = function(maxWidth) {
        var body_pr = this.getBodyPr();
        var r_ins = body_pr.rIns;
        var l_ins = body_pr.lIns;
        var max_content_width = maxWidth - r_ins - l_ins;
        this.content.Reset(0, 0, max_content_width, 20000);
        this.content.Recalculate_Page(0, true);
        var max_width = 0;
        for(var i = 0; i < this.content.Content.length; ++i) {
            var par = this.content.Content[i];
            for(var j = 0; j < par.Lines.length; ++j) {
                if(par.Lines[j].Ranges[0].W > max_width) {
                    max_width = par.Lines[j].Ranges[0].W;
                }
            }
        }
        return max_width + 2 + r_ins + l_ins;
    };
    CTextBody.prototype.getMaxContentWidth = function(maxWidth, bLeft) {
        this.content.Reset(0, 0, maxWidth - 0.01, 20000);
        if(bLeft) {
            this.content.SetApplyToAll(true);
            this.content.SetParagraphAlign(AscCommon.align_Left);
            this.content.SetApplyToAll(false);
        }
        this.content.Recalculate_Page(0, true);
        var max_width = 0, arr_content = this.content.Content, paragraph_lines, i, j;
        for(i = 0; i < arr_content.length; ++i) {
            paragraph_lines = arr_content[i].Lines;
            for(j = 0; j < paragraph_lines.length; ++j) {
                if(paragraph_lines[j].Ranges[0].W > max_width)
                    max_width = paragraph_lines[j].Ranges[0].W;
            }
        }
        return max_width + 0.01;
    };
    CTextBody.prototype.GetPrevElementEndInfo = function(CurElement) {
        return null;
    };
    CTextBody.prototype.Is_UseInDocument = function(Id) {
        if(Id != undefined) {
            if(!this.content || this.content.Get_Id() !== Id) {
                return false;
            }
        }
        if(this.parent && this.parent.Is_UseInDocument) {
            return this.parent.Is_UseInDocument();
        }
        return false;
    };
    CTextBody.prototype.Get_ParentTextTransform = function() {
        if(this.parent && this.parent.transformText) {
            return this.parent.transformText.CreateDublicate();
        }
        return null;
    };
    CTextBody.prototype.Is_ThisElementCurrent = function() {
        if(this.parent && this.parent.Is_ThisElementCurrent) {
            return this.parent.Is_ThisElementCurrent();
        }
        return false;
    };
    CTextBody.prototype.getFirstParaParaPr = function() {
        if(!this.content) {
            return null;
        }
        var oParagraph = this.content.GetFirstParagraph();
        if(!oParagraph) {
            return null;
        }
        oParagraph.Set_DocumentIndex(0); //TODO: ?
        return oParagraph.Pr;
    };

    function GetContentOneStringSizes(oContent) {
        oContent.Reset(0, 0, 20000, 20000);
        oContent.Recalculate_Page(0, true);
        return {w: oContent.Content[0].Lines[0].Ranges[0].W + 0.1, h: oContent.GetSummaryHeight() + 0.1};
    }


    function CheckNeedRecalcAutoFit(oBP1, oBP2) {
        if(AscCommon.isFileBuild()) {
            return false;
        }
        var oTF1 = oBP1 && oBP1.textFit;
        var oTF2 = oBP2 && oBP2.textFit;
        var oTFType1 = oTF1 && oTF1.type || 0;
        var oTFType2 = oTF2 && oTF2.type || 0;
        if(oTFType1 === AscFormat.text_fit_NormAuto && oTFType2 === AscFormat.text_fit_NormAuto) {
            return oTF1.lnSpcReduction !== oTF2.lnSpcReduction || oTF1.fontScale !== oTF2.fontScale;
        }
        return oTFType1 === AscFormat.text_fit_NormAuto || oTFType2 === AscFormat.text_fit_NormAuto;
    }

    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].GetContentOneStringSizes = GetContentOneStringSizes;
    window['AscFormat'].CTextBody = CTextBody;
    window['AscFormat'].CheckNeedRecalcAutoFit = CheckNeedRecalcAutoFit;
})(window);
