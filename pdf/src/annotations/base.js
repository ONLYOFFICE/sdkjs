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
	 * Class representing a base annotation.
	 * @constructor
    */
    function CAnnotationBase(sName, nType, aOrigRect, oDoc)
    {
        // если аннотация не shape based
        if (this.Id == undefined) {
            this.Id = AscCommon.g_oIdCounter.Get_NewId();
            if ((AscCommon.g_oIdCounter.m_bLoad || AscCommon.History.CanAddChanges())) {
				AscCommon.g_oTableId.Add(this, this.Id);
			}
        }

        this._apIdx = undefined;
        this.type = nType;

        this._author                = undefined;
        this._borderEffectIntensity = undefined;
        this._borderEffectStyle     = undefined;
        this._contents              = undefined;
        this._creationDate          = undefined;
        this._delay                 = false; // пока не используется
        this._doc                   = undefined;
        this._inReplyTo             = undefined;
        this._intent                = undefined;
        this._lock                  = undefined; // pdf format lock
        this._lockContent           = undefined; // pdf format lock
        this._modDate               = undefined;
        this._name                  = undefined;
        this._opacity               = 1;
        this._rect                  = [];
        this._refType               = undefined;
        this._seqNum                = undefined;
        this._strokeColor           = undefined;
        this._style                 = undefined;
        this._subject               = undefined;
        this._toggleNoView          = undefined;
        this._richContents          = undefined;
        this._display               = undefined;
        this._noRotate              = undefined;
        this._noZoom                = undefined;
        this._fillColor             = undefined;
        this._dash                  = undefined;
        this._rectDiff              = undefined;
        this._popupIdx              = undefined;
        this._meta                  = {};
        this._replies               = []; // тут будут храниться ответы (text аннотации)

        // internal
        this._bDrawFromStream   = false; // нужно ли рисовать из стрима
        this._originView = {
            normal:     null,
            mouseDown:  null,
            rollover:   null
        }
        this._wasChanged = false;
        this.Lock = new AscCommon.CLock();
        this.repliesArrayChanges = new AscCommon.CContentChanges();

        this.uid = "";

        this.SetDocument(oDoc);
        this.SetName(sName);
        this.SetRect(aOrigRect);
    };
    CAnnotationBase.prototype = Object.create(AscFormat.CBaseNoIdObject.prototype);
	CAnnotationBase.prototype.constructor = CAnnotationBase;
    
    CAnnotationBase.prototype.copy = function() {
        return this.Copy();
    };
    CAnnotationBase.prototype.Copy = function(isForMove) {
        let oCopy;

        switch (this.GetType()) {
            case AscPDF.ANNOTATIONS_TYPES.Circle: {
                oCopy = new AscPDF.CAnnotationCircle(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.FreeText: {
                oCopy = new AscPDF.CAnnotationFreeText(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Ink: {
                oCopy = new AscPDF.CAnnotationInk(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Ink: {
                oCopy = new AscPDF.CAnnotationInk(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Line: {
                oCopy = new AscPDF.CAnnotationLine(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Polygon: {
                oCopy = new AscPDF.CAnnotationPolygon(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.PolyLine: {
                oCopy = new AscPDF.CAnnotationPolyLine(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Square: {
                oCopy = new AscPDF.CAnnotationSquare(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Stamp: {
                oCopy = new AscPDF.CAnnotationStamp(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
            case AscPDF.ANNOTATIONS_TYPES.Text: {
                oCopy = new AscPDF.CAnnotationText(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());
                break;
            }
        }

        if (isForMove) {
            oCopy.movingCopy = true;
            oCopy._originView = this._originView;
            oCopy._apIdx = this._apIdx;
            oCopy.SetOriginPage(this.GetOriginPage());
            oCopy.SetDrawFromStream(this.IsNeedDrawFromStream(), true);
        }

        // base props
        let sDate           = ((new Date).getTime()).toString();
        let aStrokeColor    = this.GetStrokeColor();
        let aFillColor      = this.GetFillColor();
        let aRD             = this.GetRectangleDiff();
        let aDash           = this.GetDash();

        oCopy.SetAuthor(AscCommon.UserInfoParser.getCurrentName());
        oCopy.SetUserId(this.GetUserId());
        oCopy.SetDisplay(this.GetDisplay());
        oCopy.SetModDate(sDate);
        oCopy.SetCreationDate(sDate);
        oCopy.SetContents(this.GetContents());
        oCopy.SetBorder(this.GetBorder());
        oCopy.SetBorderEffectStyle(this.GetBorderEffectStyle());
        oCopy.SetBorderEffectIntensity(this.GetBorderEffectIntensity());
        oCopy.SetOpacity(this.GetOpacity());
        oCopy.SetWidth(this.GetWidth());
        oCopy.SetMeta(JSON.parse(JSON.stringify(this.GetMeta())));

        aStrokeColor && oCopy.SetStrokeColor(aStrokeColor.slice());
        aFillColor && oCopy.SetFillColor(aFillColor.slice());
        aDash && oCopy.SetDash(aDash.slice());
        aRD && oCopy.SetRectangleDiff(aRD.slice());
        
        // copy replies
        let oAscCommData = this.GetAscCommentData();
        if (oAscCommData) {
            let oCommentData = new AscCommon.CCommentData();
            oCommentData.Read_FromAscCommentData(oAscCommData);
            if (false == this.IsFreeText()) {
                oCommentData.SetUserData(oCopy.GetId());
            }
            oCopy.EditCommentData(oCommentData)
        }

        return oCopy;
    };

    CAnnotationBase.prototype.AddReplyByCommentData = function(CommentData, nPos) {
        let oReply = new AscPDF.CAnnotationText(AscCommon.CreateGUID(), this.GetRect().slice(), this.GetDocument());

        oReply.SetCreationDate(CommentData.m_sOOTime);
        oReply.SetModDate(CommentData.m_sOOTime);
        oReply.SetAuthor(CommentData.m_sUserName);
        oReply.SetUserId(CommentData.m_sUserId);
        oReply.SetDisplay(window["AscPDF"].Api.Types.display["visible"]);
        oReply.SetReplyTo(this.GetReplyTo() || this);
        CommentData.SetUserData(oReply.GetId());
        oReply.SetContents(CommentData.m_sText);
        oReply._wasChanged = true;
        
        if (!nPos) {
            nPos = this._replies.length;
        }

        this._replies.splice(nPos, 0, oReply);
    };
    CAnnotationBase.prototype.AddReply = function(oReplyAnnot, nPos) {
        if (!nPos) {
            nPos = this._replies.length;
        }

        this._replies.splice(nPos, 0, oReplyAnnot);

        AscCommon.History.Add(new CChangesPDFAnnotReply(this, nPos, [oReplyAnnot], true));
    };
    CAnnotationBase.prototype.Add_ContentChanges = function(Changes) {
        this.repliesArrayChanges.Add(Changes);
    };
    CAnnotationBase.prototype.Refresh_ContentChanges = function() {
        this.repliesArrayChanges.Refresh();
    };
    CAnnotationBase.prototype.Clear_ContentChanges = function() {
        this.repliesArrayChanges.Clear();
    };
    CAnnotationBase.prototype.SetMeta = function(oMeta) {
        AscCommon.History.Add(new CChangesPDFAnnotMeta(this, this._meta, oMeta))

        this._meta = oMeta;
    };
    CAnnotationBase.prototype.GetMeta = function() {
        return this._meta;
    };
    CAnnotationBase.prototype.onMouseExit = function() {};
    CAnnotationBase.prototype.onMouseEnter = function() {};

    CAnnotationBase.prototype.IsEditFieldShape = function() {
        return false;
    };
    CAnnotationBase.prototype.SetUserId = function(sUID) {
        if (this.uid == sUID) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotUserId(this, this.uid, sUID));
        this.uid = sUID;
    };
    CAnnotationBase.prototype.GetUserId = function() {
        return this.uid;
    };
    CAnnotationBase.prototype.GetDocContent = function() {
        return null;
    };
    CAnnotationBase.prototype.SetReplyTo = function(oAnnot) {
        this._inReplyTo = oAnnot;
    };
    CAnnotationBase.prototype.GetReplyTo = function() {
        return this._inReplyTo;
    };
    CAnnotationBase.prototype.SetRefType = function(nType) {
        this._refType = nType;
    };
    CAnnotationBase.prototype.GetRefType = function() {
        return this._refType;
    };
    CAnnotationBase.prototype.SetRectangleDiff = function(aDiff) {
        AscCommon.History.Add(new CChangesPDFAnnotRD(this, this.GetRectangleDiff(), aDiff));

        this._rectDiff = aDiff;
        if (this.IsShapeBased()) {
            this.recalcGeometry();
            this.SetWasChanged(true);
            this.SetNeedRecalc(true);
        }
    };
    CAnnotationBase.prototype.GetRectangleDiff = function() {
        return this._rectDiff;
    };
    CAnnotationBase.prototype.SetNoRotate = function(bValue) {
        this._noRotate = bValue;
    };
    CAnnotationBase.prototype.SetToggleNoView = function(value) {
        this._toggleNoView = value;
    };
    CAnnotationBase.prototype.SetNoZoom = function(bValue) {
        this._noZoom = bValue;
    };
    CAnnotationBase.prototype.SetDash = function(aDash) {
        AscCommon.History.Add(new CChangesPDFAnnotBorderDash(this, this._dash, aDash));
        this._dash = aDash;
    };
    CAnnotationBase.prototype.GetDash = function() {
        return this._dash;
    };
    CAnnotationBase.prototype.SetFillColor = function(aColor) {
        AscCommon.History.Add(new CChangesPDFAnnotFill(this, this.GetFillColor(), aColor));

        this._fillColor = aColor;

        if (this.IsShapeBased()) {
            let oRGB    = this.GetRGBColor(aColor);
            let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
            this.setFill(oFill);
            this.SetNeedRecalc(true);
            this.SetNeedUpdateOpacity(true);
        }
        else {
            this.AddToRedraw();
        }
    };
    CAnnotationBase.prototype.GetFillColor = function() {
        return this._fillColor;
    };
    CAnnotationBase.prototype.SetWidth = function(nWidthPt) {
        AscCommon.History.Add(new CChangesPDFAnnotStrokeWidth(this, this.GetWidth(), nWidthPt));

        this._width = nWidthPt;

        this.SetWasChanged(true);
        this.private_UpdateLn();
    };
    CAnnotationBase.prototype.private_UpdateLn = function() {
        let nWidthPt = this.GetWidth();
        
        if (this.IsShapeBased()) {
            let oLine = this.spPr.ln;
            oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);

            if (nWidthPt == 0) {
                oLine.setFill(AscFormat.CreateNoFillUniFill());
            }
            else {
                AscCommon.History.StartNoHistoryMode();
                this.SetStrokeColor(this.GetStrokeColor());
                AscCommon.History.EndNoHistoryMode();
            }

            this.handleUpdateLn();
        }
    };
    CAnnotationBase.prototype.GetWidth = function() {
        return this._width;
    };
    CAnnotationBase.prototype.FillCommentsDataFrom = function(oAnnot) {
        let oCurAscCommData = oAnnot.GetAscCommentData();
        let oCurData = oCurAscCommData ? new AscCommon.CCommentData() : undefined;
        if (oCurData) {
            oCurData.Read_FromAscCommentData(oCurAscCommData);
            oCurData.SetUserData(this.GetId());
            oCurData.SetSolved(false);

            oCurData.m_aReplies.forEach(function(reply) {
                reply.m_sUserId = Asc.editor.documentUserId;
            });
        }
        
        this.SetUserId(Asc.editor.documentUserId);
        this.EditCommentData(oCurData);
    };
    CAnnotationBase.prototype.SetRichContents = function(aRCInfo) {
        AscCommon.History.Add(new CChangesPDFAnnotRC(this, this._richContents, aRCInfo));
        this._richContents = aRCInfo;

        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
    };
    CAnnotationBase.prototype.GetRichContents = function() {
        return this._richContents;
    };
    CAnnotationBase.prototype.SetIntent = function(nType) {
        AscCommon.History.Add(new CChangesPDFAnnotIntent(this, this._intent, nType));

        this._intent = nType;
    };
    CAnnotationBase.prototype.GetIntent = function() {
        return this._intent;
    };
    CAnnotationBase.prototype.SetLock = function(bValue) {
        this._lock = bValue;
    };
    CAnnotationBase.prototype.SetLockContent = function(bValue) {
        this._lockContent = bValue;
    };
    CAnnotationBase.prototype.SetBorder = function(nType) {
        if (this._border == nType) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotBorderType(this, this._border, nType));
        this._border = nType;
    };
    CAnnotationBase.prototype.GetBorder = function() {
        return this._border;
    };
    CAnnotationBase.prototype.SetBorderEffectIntensity = function(nValue) {
        if (nValue == this._borderEffectIntensity) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotBorderIntensity(this, this._borderEffectIntensity, nValue));
        this._borderEffectIntensity = nValue;
        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
    };
    CAnnotationBase.prototype.GetBorderEffectIntensity = function() {
        return this._borderEffectIntensity;
    };
    CAnnotationBase.prototype.SetBorderEffectStyle = function(nStyle) {
        if (nStyle == this._borderEffectStyle) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotBorderEffect(this, this._borderEffectStyle, nStyle));
        this._borderEffectStyle = nStyle;
        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
    };
    CAnnotationBase.prototype.GetBorderEffectStyle = function() {
        return this._borderEffectStyle;
    };

    CAnnotationBase.prototype.DrawSelected = function() {};
    CAnnotationBase.prototype.SetName = function(sName) {
        if (sName == this._name) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotName(this, this._name, sName));
        this._name = sName;
        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.GetName = function() {
        return this._name;
    };
    CAnnotationBase.prototype.SetOpacity = function (value) {
        if (this._opacity === value) return;

        AscCommon.History.Add(new CChangesPDFAnnotOpacity(this, this._opacity, value));
        this._opacity = value;
        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
        this.SetNeedUpdateOpacity(true);
    };
    CAnnotationBase.prototype.SetNeedUpdateOpacity = function(isNeed) {
        this._needUpdateOpacity = isNeed;
    };
    CAnnotationBase.prototype.IsNeedUpdateOpacity = function() {
        return this._needUpdateOpacity;
    };

    CAnnotationBase.prototype.UpdateOpacity = function() {
        const t = this.GetOpacity() * 100 * 2.55;

        if (!this.IsShapeBased()) {
            this.AddToRedraw();
            this.SetNeedUpdateOpacity(false);
            return;
        }

        if (this.IsFreeText()) {
            for (let i = 1; i < this.spTree.length; i++) {
                const oLine = this.spTree[i].spPr.ln;
                oLine.Fill.transparent = t;

                const oFill = this.spPr.Fill;
                oFill.transparent = t;

                this.spTree[i].handleUpdateLn();
                this.spTree[i].handleUpdateFill();
            }
        }
        else {
            const oLine = this.spPr.ln;
            oLine.Fill.transparent = t;

            const oFill = this.spPr.Fill;
            oFill.transparent = t;

            this.handleUpdateLn();
            this.handleUpdateFill();
        }

        this.SetNeedUpdateOpacity(false);
    }
    CAnnotationBase.prototype.GetOpacity = function() {
        return this._opacity;
    };
    CAnnotationBase.prototype.IsPdfObject = function() {
        return true;
    };
    CAnnotationBase.prototype.GetAllFonts = function(fontMap) {return fontMap};
    /**
	 * Invokes only on open forms.
	 * @memberof CAnnotationBase
	 * @typeofeditors ["PDF"]
	 */
    CAnnotationBase.prototype.SetOriginPage = function(nPage) {
        AscCommon.History.Add(new CChangesPDFAnnotOrigPage(this, this._origPage, nPage));

        this._origPage = nPage;
    };
    CAnnotationBase.prototype.GetOriginPage = function() {
        return this._origPage;
    };
    CAnnotationBase.prototype.SetWasChanged = function(isChanged, viewSync) {
        let oViewer   = Asc.editor.getDocumentRenderer();
        let canChange = !oViewer.IsOpenAnnotsInProgress && AscCommon.History.CanAddChanges();

        let prev    = this._wasChanged;
        let changed = prev !== isChanged && canChange;
        if (changed) {
            AscCommon.History.Add(new CChangesPDFAnnotChanged(this, prev, isChanged));
            this._wasChanged = isChanged;
        }

        if (false !== viewSync && canChange && (isChanged || changed)) {
            this.SetDrawFromStream(!isChanged);
        }
    };
    CAnnotationBase.prototype.IsChanged = function() {
        return this._wasChanged;  
    };
    CAnnotationBase.prototype.DrawFromStream = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;
            
        let originView = this.GetOriginView(oGraphicsPDF.GetDrawingPageW(), oGraphicsPDF.GetDrawingPageH());

        if (originView) {
            let aOrigRect = this.GetRect();
            
            let X = aOrigRect[0];
            let Y = aOrigRect[1];

            if (this.IsHighlight())
                AscPDF.startMultiplyMode(oGraphicsPDF.GetContext());
            
            oGraphicsPDF.SetGlobalAlpha(1);
            oGraphicsPDF.DrawImageXY(originView, X, Y, undefined, true);
            AscPDF.endMultiplyMode(oGraphicsPDF.GetContext());
        }

        // oGraphicsPDF.SetLineWidth(1);
        // let aOringRect  = this.GetRect();
        // let X       = aOringRect[0];
        // let Y       = aOringRect[1];
        // let nWidth  = aOringRect[2] - aOringRect[0];
        // let nHeight = aOringRect[3] - aOringRect[1];

        // Y += 1 / 2;
        // X += 1 / 2;
        // nWidth  -= 1;
        // nHeight -= 1;

        // oGraphicsPDF.SetStrokeStyle(0, 255, 255);
        // oGraphicsPDF.SetLineDash([]);
        // oGraphicsPDF.BeginPath();
        // oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
        // oGraphicsPDF.Stroke();
    };
    CAnnotationBase.prototype.SetSubject = function(sSubject) {
        if (this._subject == sSubject) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotSubject(this, this._subject, sSubject));
        this._subject = sSubject;

        this.SetWasChanged(true, false);
    };
    CAnnotationBase.prototype.GetSubject = function() {
        return this._subject;
    };
    /**
     * Returns a canvas with origin view (from appearance stream) of current annot.
	 * @memberof CAnnotationBase
	 * @typeofeditors ["PDF"]
     * @returns {canvas}
	 */
    CAnnotationBase.prototype.GetOriginView = function(nPageW, nPageH) {
        if (this.GetApIdx() == undefined)
            return null;

        nPageW = Math.round(nPageW);
        nPageH = Math.round(nPageH);

        let oViewer = editor.getDocumentRenderer();
        let oFile   = oViewer.file;
        
        let oApearanceInfo = this.GetOriginViewInfo(nPageW, nPageH);
        let oSavedView, oApInfoTmp;
        if (!oApearanceInfo)
            return null;
            
        oApInfoTmp = oApearanceInfo["N"];
        
        oSavedView = this._originView.normal;
        if (oSavedView) {
            if (oSavedView.width == oApearanceInfo["w"] && oSavedView.height == oApearanceInfo["h"]) {
                return oSavedView;
            }
        }
        
        let canvas  = document.createElement("canvas");
        let nWidth  = oApearanceInfo["w"];
        let nHeight = oApearanceInfo["h"];

        canvas.width    = nWidth;
        canvas.height   = nHeight;

        canvas.x    = oApearanceInfo["x"];
        canvas.y    = oApearanceInfo["y"];
        
        if (!oApInfoTmp)
            return null;

        let supportImageDataConstructor = (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge) ? false : true;

        let ctx             = canvas.getContext("2d");
        let mappedBuffer    = oFile.getUint8ClampedArray(oApInfoTmp["retValue"], 4 * nWidth * nHeight);
        let imageData       = null;

        if (supportImageDataConstructor)
        {
            imageData = new ImageData(mappedBuffer, nWidth, nHeight);
        }
        else
        {
            imageData = ctx.createImageData(nWidth, nHeight);
            imageData.data.set(mappedBuffer, 0);                    
        }
        if (ctx)
            ctx.putImageData(imageData, 0, 0);
        
        oViewer.file.free(oApInfoTmp["retValue"]);

        this._originView.normal = canvas;

        return canvas;
    };
    /**
     * Returns AP info of this annotation.
	 * @memberof CAnnotationBase
	 * @typeofeditors ["PDF"]
     * @returns {Object}
	 */
    CAnnotationBase.prototype.GetOriginViewInfo = function(nPageW, nPageH) {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let nPage       = this.GetOriginPage();
        let oOriginPage = oFile.pages.find(function(page) {
            return page.originIndex == nPage;
        });
        let nApIdx = this.GetApIdx();

        if (oOriginPage.annotsAPInfo == null || oOriginPage.annotsAPInfo.size.w != nPageW || oOriginPage.annotsAPInfo.size.h != nPageH) {
            oOriginPage.annotsAPInfo = {
                info: oFile.nativeFile["getAnnotationsAP"](nPage, nPageW, nPageH),
                size: {
                    w: nPageW,
                    h: nPageH
                }
            }
        }
        
        for (let i = 0; i < oOriginPage.annotsAPInfo.info.length; i++) {
            if (oOriginPage.annotsAPInfo.info[i]["i"] == nApIdx)
                return oOriginPage.annotsAPInfo.info[i];
        }

        return null;
    };
    
    CAnnotationBase.prototype.ClearCache = function() {
        this._originView.normal = null;
    };
    CAnnotationBase.prototype.SetPosition = function(x, y) {
        let aCurRect = this.GetRect();

        let nOldX = aCurRect[0];
        let nOldY = aCurRect[1];

        let nDeltaX = x - nOldX;
        let nDeltaY = y - nOldY;

        if (0 == nDeltaX && 0 == nDeltaY) {
            return;
        }

        if (this.IsInk()) {
            let aPath;
            for (let i = 0; i < this._gestures.length; i++) {
                aPath = this._gestures[i];
                for (let j = 0; j < aPath.length; j+=2) {
                    aPath[j] += nDeltaX;
                    aPath[j+1] += nDeltaY;
                }
            }
        }
        else if (this.IsLine()) {
            for (let i = 0; i < this._points.length; i+=2) {
                this._points[i] += nDeltaX;
                this._points[i+1] += nDeltaY;
            }
        }
        else if (this.IsPolygon() || this.IsPolyLine()) {
            for (let i = 0; i < this._vertices.length; i+=2) {
                this._vertices[i] += nDeltaX;
                this._vertices[i+1] += nDeltaY;
            }
        }
        else if (this.IsFreeText()) {
            let aCallout = this.GetCallout();
            if (aCallout) {
                for (let i = 0; i < aCallout.length; i+=2) {
                    aCallout[i] += nDeltaX;
                    aCallout[i+1] += nDeltaY;
                }
            }
        }

        AscCommon.History.Add(new CChangesPDFAnnotPos(this, [aCurRect[0], aCurRect[1]], [x, y]));

        let nWidth  = aCurRect[2] - aCurRect[0];
        let nHeight = aCurRect[3] - aCurRect[1];

        let aRD = this.GetRectangleDiff() || [0, 0, 0, 0];

        this._rect[0] = x;
        this._rect[1] = y;
        this._rect[2] = x + nWidth;
        this._rect[3] = y + nHeight;

        if (this.IsShapeBased()) {
            let oXfrm = this.getXfrm();
            AscCommon.History.StartNoHistoryMode();
            oXfrm.setOffX((this._rect[0] + aRD[0]) * g_dKoef_pt_to_mm);
            oXfrm.setOffY((this._rect[1] + aRD[1]) * g_dKoef_pt_to_mm);
            AscCommon.History.EndNoHistoryMode();
        }

        this.SetNeedRecalc(true);
        this.SetWasChanged(true, false);
    };
    CAnnotationBase.prototype.IsShapeBased = function() {
        return this instanceof AscPDF.CPdfShape || this instanceof AscFormat.CGroupShape;
    };
    CAnnotationBase.prototype.IsHighlight = function() {
        return false;
    };
    CAnnotationBase.prototype.IsRedact = function() {
        return false;
    };
    CAnnotationBase.prototype.IsTextMarkup = function() {
        return false;
    };
    CAnnotationBase.prototype.IsComment = function() {
        return false;
    };
    CAnnotationBase.prototype.IsInk = function() {
        return false;
    };
    CAnnotationBase.prototype.IsLine = function() {
        return false;
    };
    CAnnotationBase.prototype.IsCircle = function() {
        return false;
    };
    CAnnotationBase.prototype.IsSquare = function() {
        return false;
    };
    CAnnotationBase.prototype.IsPolygon = function() {
        return false;
    };
    CAnnotationBase.prototype.IsPolyLine = function() {
        return false;
    };
    CAnnotationBase.prototype.IsFreeText = function() {
        return false;
    };
    CAnnotationBase.prototype.IsStamp = function() {
        return false;
    };
    CAnnotationBase.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
        if (bRecalc == false) {
            this._needRecalc = false;
        }
        else {
            this._needRecalc = true;
            // note: movingCopy флаг означает, что объект был скопирован для отрисовки на overlay
            if (bSkipAddToRedraw != true && this.movingCopy != true)
                this.AddToRedraw();
        }
    };
    CAnnotationBase.prototype.IsNeedRecalc = function() {
        return this._needRecalc;
    };
    CAnnotationBase.prototype.GetRect = function() {
        return this._rect || (this.GetReplyTo() ? this.GetReplyTo().GetRect() : this._rect);
    };
    CAnnotationBase.prototype.IsNeedDrawFromStream = function() {
        return this._bDrawFromStream;
    };
    CAnnotationBase.prototype.SetDrawFromStream = function(bFromStream) {
        AscCommon.History.Add(new CChangesPDFAnnotChangedView(this, this._bDrawFromStream, bFromStream));
        this._bDrawFromStream = bFromStream;
    };
    CAnnotationBase.prototype.SetRect = function(aOrigRect) {
        AscCommon.History.Add(new CChangesPDFAnnotRect(this, this.GetRect(), aOrigRect));

        this._rect = aOrigRect;
        this.SetWasChanged(true);

        if (this.IsShapeBased()) {
            this.SetNeedRecalc(true);
            this.SetNeedRecalcSizes(true);
            this.SetWasChanged(true);
        }
    };
    CAnnotationBase.prototype.SetNeedRecalcSizes = function(bRecalc) {
        this._needRecalcSizes = bRecalc;
        this.recalcGeometry && this.recalcGeometry();
    };
    CAnnotationBase.prototype.IsNeedRecalcSizes = function() {
        return this._needRecalcSizes;
    };
    CAnnotationBase.prototype.IsUseInDocument = function() {
        let oPage = this.GetParentPage();
        if (oPage && oPage.annots.includes(this) && oPage.GetIndex() !== -1) {
            return true;
        }

        return false;
    };
    
    CAnnotationBase.prototype.GetId = function() {
        return this.Id;
    };
    CAnnotationBase.prototype.Get_Id = function() {
        return this.GetId();
    };
    CAnnotationBase.prototype.GetType = function() {
        return this.type;
    };
    CAnnotationBase.prototype.SetPopupIdx = function(nIdx) {
        this._popupIdx = nIdx;
    };
    CAnnotationBase.prototype.GetPopupIdx = function() {
        return this._popupIdx;
    };
    CAnnotationBase.prototype.SetParentPage = function(oParent) {
        this.parentPage = oParent;
    };
    CAnnotationBase.prototype.GetParentPage = function() {
        let oReplyTo = this.GetReplyTo();
        if (oReplyTo && !this.parentPage) {
            return oReplyTo.GetParentPage();
        }

        return this.parentPage;
    };
    CAnnotationBase.prototype.SetPage = function(nPage) {
        if (this.GetPage() == nPage) {
            return;
        }
        
        let oDoc        = this.GetDocument();
        let oNewPage    = oDoc.GetPageInfo(nPage);

        if (oNewPage) {
            oDoc.RemoveAnnot(this.GetId(), true);
            oDoc.AddAnnot(this, nPage);
            this.selectStartPage = nPage;
        }
    };
    CAnnotationBase.prototype.GetPage = function() {
        let oParentPage = this.GetParentPage();
        if (!oParentPage || !(oParentPage instanceof AscPDF.CPageInfo)) {
            return -1;
        }

        return oParentPage.GetIndex();
    };
    CAnnotationBase.prototype.SetDocument = function(oDoc) {
        if (this._doc == oDoc) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFObjectSetDocument(this, this._doc, oDoc));
        this._doc = oDoc;
    };
    CAnnotationBase.prototype.GetDocument = function() {
        return this._doc;
    };
    CAnnotationBase.prototype.IsHidden = function() {
        let nType = this.GetDisplay();
        if (nType == window["AscPDF"].Api.Types.display["hidden"] || nType == window["AscPDF"].Api.Types.display["noView"])
            return true;

        return false;
    };
    CAnnotationBase.prototype.SetDisplay = function(nType) {
        if (nType == this._display) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotDisplay(this, this._display, nType));

        this._display = nType;
        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.GetDisplay = function() {
        return this._display;
    };
    CAnnotationBase.prototype.onMouseUp = function(e) {
        if (e.button != 2) {
            this.GetDocument().ShowComment([this.GetId()]);
        }
    };
    CAnnotationBase.prototype.SetContents = function(contents) {
        if (this.GetContents() == contents)
            return;

        let oViewer         = Asc.editor.getDocumentRenderer();
        let oDoc            = Asc.editor.getPDFDoc();
        let sCurContents    = this.GetContents();

        let bSendAddCommentEvent = false;
        if (this._contents == null && contents != null)
            bSendAddCommentEvent = true;
        
        this._contents  = contents;
        
        if (AscCommon.History.UndoRedoInProgress == false) {
            AscCommon.History.Add(new CChangesPDFAnnotContents(this, sCurContents, contents));
        }
        
        this.SetWasChanged(true, false);
        if (bSendAddCommentEvent)
            oDoc.CheckComment(this);
        
        if (this._contents == null && this.IsUseInDocument())
            Asc.editor.sync_RemoveComment(this.GetId());
    };
    CAnnotationBase.prototype.IsUseContentAsComment = function() {
        return !(this.IsFreeText() || this.IsLine() && this.IsDoCaption());
    };
    CAnnotationBase.prototype.Recalculate = function(bForce) {
        if (!this.IsShapeBased()) {
            return;
        }

        if (true !== bForce && false == this.IsNeedRecalc()) {
            return;
        }

        if (this.IsNeedUpdateOpacity()) {
            this.UpdateOpacity();
        }
        
        if (this.IsNeedRecalcSizes()) {
            this.RecalcSizes();
        }
        if (this.recalcInfo.recalculateGeometry) {
            this.RefillGeometry();
        }

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CAnnotationBase.prototype.RecalcSizes = function() {
        if (!this.IsShapeBased()) {
            return;
        }

        let aRect = this.GetRect();
        let aRD = this.GetRectangleDiff() || [0, 0, 0, 0];

        let extX = ((aRect[2] - aRect[0]) - aRD[0] - aRD[2]) * g_dKoef_pt_to_mm;
        let extY = ((aRect[3] - aRect[1]) - aRD[1] - aRD[3]) * g_dKoef_pt_to_mm;

        AscCommon.History.StartNoHistoryMode();
        this.spPr.xfrm.setOffX((aRect[0] + aRD[0]) * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setOffY((aRect[1] + aRD[1]) * g_dKoef_pt_to_mm);

        this.spPr.xfrm.setExtX(extX);
        this.spPr.xfrm.setExtY(extY);
        AscCommon.History.EndNoHistoryMode();

        this.SetNeedRecalcSizes(false);
    };
    CAnnotationBase.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        this.Recalculate();
        this.draw(oGraphicsWord);

        // draw annot rect
        // if (oGraphicsPDF) {
        //     oGraphicsPDF.SetLineWidth(1);
        //     let aOringRect  = this.GetRect();
        //     let X       = aOringRect[0];
        //     let Y       = aOringRect[1];
        //     let nWidth  = aOringRect[2] - aOringRect[0];
        //     let nHeight = aOringRect[3] - aOringRect[1];

        //     Y += 1 / 2;
        //     X += 1 / 2;
        //     nWidth  -= 1;
        //     nHeight -= 1;

        //     oGraphicsPDF.SetStrokeStyle(0, 255, 255);
        //     oGraphicsPDF.SetLineDash([]);
        //     oGraphicsPDF.BeginPath();
        //     oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
        //     oGraphicsPDF.Stroke();

        //     // if (this.IsLine()) {
        //     //     let aPoints = this.GetLinePoints();

        //     //     oGraphicsPDF.BeginPath();
        //     //     oGraphicsPDF.SetStrokeStyle(255, 0, 0);
        //     //     oGraphicsPDF.MoveTo(aPoints[0], aPoints[1]);
        //     //     oGraphicsPDF.LineTo(aPoints[2], aPoints[3]);
        //     //     oGraphicsPDF.Stroke();
        //     // }
        // }
    };
    CAnnotationBase.prototype.changeFlipH = function () {
        return false;
    };
    CAnnotationBase.prototype.changeFlipV = function () {
        return false;
    };
    CAnnotationBase.prototype.GetReplies = function() {
        return this._replies;
    };
    CAnnotationBase.prototype.GetReply = function(nPos) {
        return this._replies[nPos];
    };
    CAnnotationBase.prototype.RemoveComment = function() {
        this.EditCommentData(undefined);
        this.SetContents(undefined);
    };
    CAnnotationBase.prototype.EditCommentData = function(oCommentData) {
        let oDoc = this.GetDocument();

        let oCurAscCommData = this.GetAscCommentData();
        let oCurData = oCurAscCommData ? new AscCommon.CCommentData() : undefined;
		oCurData && oCurData.Read_FromAscCommentData(oCurAscCommData);

        AscCommon.History.Add(new CChangesPDFAnnotCommentData(this, oCurData, oCommentData));

        if (oCommentData == null) {
            this._replies.length = 0;
            Asc.editor.sync_RemoveComment(this.GetId());
            return;
        }

        let oFirstCommToEdit;
        if (this.GetId() == oCommentData.m_sUserData)
            oFirstCommToEdit = this;
        else {
            oFirstCommToEdit = this._replies.find(function(oReply) {
                return oCommentData.m_sUserData == oReply.GetId(); 
            });
        }
        
        AscCommon.History.StartNoHistoryMode();
        if (null == oFirstCommToEdit) {
            this.AddReplyByCommentData(oCommentData);
            oFirstCommToEdit = this.GetReply(0);
            oDoc.CheckComment(this);
        }

        if (oFirstCommToEdit.GetContents() != oCommentData.m_sText) {
            oFirstCommToEdit.SetModDate(oCommentData.m_sOOTime);
            oFirstCommToEdit.SetAuthor(oCommentData.m_sUserName);
            oFirstCommToEdit.SetUserId(oCommentData.m_sUserId);
            oFirstCommToEdit.SetContents(oCommentData.m_sText);
        }

        let aReplyToDel = [];
        let oReply, oReplyCommentData;
        for (let i = 0; i < this._replies.length; i++) {
            oReply = this._replies[i];
            if (oFirstCommToEdit == oReply)
                continue;

            oReplyCommentData = oCommentData.m_aReplies.find(function(item) {
                return item.m_sUserData == oReply.GetId(); 
            });

            if (oReplyCommentData) {
                oReply.EditCommentData(oReplyCommentData);
            }
            else {
                aReplyToDel.push(oReply);
            }
        }

        for (let i = aReplyToDel.length - 1; i >= 0; i--) {
            this._replies.splice(this._replies.indexOf(aReplyToDel[i]), 1);
        }

        for (let i = 0; i < oCommentData.m_aReplies.length; i++) {
            oReplyCommentData = oCommentData.m_aReplies[i];
            if (!this._replies.find(function(reply) {
                return oReplyCommentData.m_sUserData == reply.GetId();
            })) {
               this.AddReplyByCommentData(oReplyCommentData, i);
            }
        }
        AscCommon.History.EndNoHistoryMode();

        if (this.IsComment()) {
            if (oCommentData.m_bSolved) {
                this.SetState(AscPDF.TEXT_ANNOT_STATE.Accepted);
            }
            else {
                this.SetState(AscPDF.TEXT_ANNOT_STATE.Unknown);
            }
        }

        for (let i = 0; i < this._replies.length; i++) {
            if (oCommentData.m_bSolved) {
                this._replies[i].SetState(AscPDF.TEXT_ANNOT_STATE.Accepted);
            }
            else {
                this._replies[i].SetState(AscPDF.TEXT_ANNOT_STATE.Unknown);
            }
        }

        Asc.editor.sync_ChangeCommentData(this.GetId(), oCommentData);
    };
    CAnnotationBase.prototype.GetAscCommentData = function() {
        let oAscCommData = new Asc.asc_CCommentDataWord(null);
        if (null == this.GetContents()) {
            return undefined;
        }

        oAscCommData.asc_putText(this.GetContents());
        let sModDate = this.GetModDate();
        if (sModDate)
            oAscCommData.asc_putOnlyOfficeTime(sModDate.toString());
        oAscCommData.asc_putUserId(this.GetUserId());
        oAscCommData.asc_putUserName(this.GetAuthor());
        oAscCommData.asc_putSolved(false);
        oAscCommData.asc_putQuoteText("");
        oAscCommData.asc_putUserData(this.GetId());

        this._replies.forEach(function(reply) {
            let oReplyAscCommData = reply.GetAscCommentData();
            if (oReplyAscCommData) {
                oAscCommData.m_aReplies.push(oReplyAscCommData);
            }
        });

        return oAscCommData;
    };
    CAnnotationBase.prototype.GetContents = function() {
        return this._contents;
    };
    CAnnotationBase.prototype.SetModDate = function(sDate) {
        if (sDate == this._modDate) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotModDate(this, this._modDate, sDate));
        this._modDate = sDate;
        this.SetWasChanged(true, false);
    };
    CAnnotationBase.prototype.GetModDate = function(bPDF) {
        if (this._modDate == undefined)
            return this._modDate;

        if (bPDF) {
            return formatTimestampToPDF(this._modDate); 
        }

        return this._modDate;
    };
    CAnnotationBase.prototype.SetCreationDate = function(sDate) {
        if (sDate == this._creationDate) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotCreationDate(this, this._creationDate, sDate));
        this._creationDate = sDate;
        this.SetWasChanged(true, false);
    };
    CAnnotationBase.prototype.GetCreationDate = function(bPDF) {
        if (this._creationDate == undefined)
            return this._creationDate;

        if (bPDF) {
            return formatTimestampToPDF(this._creationDate); 
        }

        return this._creationDate;
    };
    
    CAnnotationBase.prototype.SetAuthor = function(sAuthor) {
        if (sAuthor == this._author) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotAuthor(this, this._author, sAuthor));
        this._author = sAuthor;
        this.SetWasChanged(true, false);
    };
    CAnnotationBase.prototype.GetAuthor = function() {
        return this._author;
    };
    CAnnotationBase.prototype.IsAnnot = function() {
        return true;
    };
    CAnnotationBase.prototype.IsDrawing = function() {
        return false;
    };
    CAnnotationBase.prototype.IsForm = function() {
        return false;
    };
    CAnnotationBase.prototype.SetApIdx = function(nIdx) {
        if (undefined != this._apIdx) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotApIdx(this, this._apIdx, nIdx));
        
        this._apIdx = nIdx;
    };
    CAnnotationBase.prototype.GetApIdx = function() {
        if (undefined !== this._apIdx) {
            return this._apIdx;
        }
        else {
            let nPos = Object.keys(AscCommon.g_oTableId.m_aPairs).indexOf(this.GetId());
            if (-1 !== nPos) {
                return Asc.editor.getPDFDoc().GetCurMaxApIdx() + nPos;
            }
            
            return undefined;
        }
    };
    CAnnotationBase.prototype.AddToRedraw = function() {
        let oViewer = editor.getDocumentRenderer();
        let nPage   = this.GetPage();

        if (false == this.IsUseInDocument()) {
            return;
        }

        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[nPage]) {
                oViewer.pagesInfo.pages[nPage].needRedrawAnnots = true;
                oViewer.thumbnails && oViewer.thumbnails._repaintPage(nPage);
            }
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    /**
	 * Gets rgb color object from internal color array.
	 * @memberof CAnnotationBase
	 * @typeofeditors ["PDF"]
     * @returns {object}
	 */
    CAnnotationBase.prototype.GetRGBColor = function(aInternalColor) {
        let oColor = {
            r: 255,
            g: 255,
            b: 255
        };

        if (!aInternalColor || aInternalColor.length == 0) {
            return oColor;
        }
        
        if (aInternalColor.length == 1) {
            oColor = {
                r: Math.round(aInternalColor[0] * 255),
                g: Math.round(aInternalColor[0] * 255),
                b: Math.round(aInternalColor[0] * 255)
            }
        }
        else if (aInternalColor.length == 3) {
            oColor = {
                r: Math.round(aInternalColor[0] * 255),
                g: Math.round(aInternalColor[1] * 255),
                b: Math.round(aInternalColor[2] * 255)
            }
        }
        else if (aInternalColor.length == 4) {
            function cmykToRgb(c, m, y, k) {
                return {
                    r: Math.round(255 * (1 - c) * (1 - k)),
                    g: Math.round(255 * (1 - m) * (1 - k)),
                    b: Math.round(255 * (1 - y) * (1 - k))
                }
            }

            oColor = cmykToRgb(aInternalColor[0], aInternalColor[1], aInternalColor[2], aInternalColor[3]);
        }
        else if (aInternalColor.length > 4) {
            return this.GetRGBColor(aInternalColor.slice(0, 3));
        }

        return oColor;
    };
    CAnnotationBase.prototype.onMouseDown = function(x, y, e) {
        let oViewer         = Asc.editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;

        this.selectStartPage = this.GetPage();

        let pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
    };
    CAnnotationBase.prototype.createMoveTrack = function() {
        return new AscFormat.MoveAnnotationTrack(this);
    };

    CAnnotationBase.prototype.SetStrokeColor = function(aColor) {
        AscCommon.History.Add(new CChangesPDFAnnotStroke(this, this.GetStrokeColor(), aColor));

        this._strokeColor = aColor;
        this.SetWasChanged(true);

        if (!aColor) {
            aColor = [0, 0, 0];
        }
        
        if (this.IsShapeBased()) {
            let oRGB    = this.GetRGBColor(aColor);
            let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
            let oLine   = this.spPr.ln;
            oLine.setFill(oFill);
            this.SetNeedRecalc(true);
            this.SetNeedUpdateOpacity(true);
        }
        else {
            this.AddToRedraw();
        }
    };
    CAnnotationBase.prototype.GetStrokeColor = function() {
        return this._strokeColor;
    };

    CAnnotationBase.prototype.WriteToBinaryBase = function(memory) {
        // type
        memory.WriteByte(this.GetType());

        // apidx
        memory.WriteLong(this.GetApIdx());

        // annont flags
        let bHidden      = false;
        let bPrint       = false;
        let bNoView      = false;
        let ToggleNoView = false;
        let locked       = false;
        let lockedC      = false;
        let noZoom       = false;
        let noRotate     = false;

        let nDisplayType = this.GetDisplay();
        if (nDisplayType == 1) {
            bHidden = true;
        }
        else if (nDisplayType == 0 || nDisplayType == 3) {
            bPrint = true;
            if (nDisplayType == 3) {
                bNoView = true;
            }
        }
        let annotFlags = (bHidden << 1) |
        (bPrint << 2) |
        (noZoom << 3) |
        (noRotate << 4) |
        (bNoView << 5) |
        (locked << 7) |
        (ToggleNoView << 8) |
        (lockedC << 9);

        memory.WriteLong(annotFlags);

        // page
        let nPage = this.GetOriginPage();
        if (nPage == undefined)
            nPage = this.GetPage();

        memory.WriteLong(nPage);

        let sName           = this.GetName();
        let sContents       = this.GetContents();
        let BES             = this.GetBorderEffectStyle();
        let BEI             = this.GetBorderEffectIntensity();
        let aStrokeColor    = this.GetType() == AscPDF.ANNOTATIONS_TYPES.Text ? this.GetFillColor() : this.GetStrokeColor();
        let nBorder         = this.GetBorder();
        let nBorderW        = this.GetWidth();
        let sModDate        = this.GetModDate(true);

        // rect
        let aOrigRect = this.GetRect();
        if (this.IsStamp() && memory.docRenderer) {
            let nScale = this.GetOriginViewScale();

            // for not clipping by half border width
            memory.WriteDouble(aOrigRect[0] - (nBorderW / 2) * nScale); // x1
            memory.WriteDouble(aOrigRect[1] - (nBorderW / 2) * nScale); // y1
            memory.WriteDouble(aOrigRect[2] + (nBorderW / 2) * nScale); // x2
            memory.WriteDouble(aOrigRect[3] + (nBorderW / 2) * nScale); // y2
        }
        else {
            memory.WriteDouble(aOrigRect[0]); // x1
            memory.WriteDouble(aOrigRect[1]); // y1
            memory.WriteDouble(aOrigRect[2]); // x2
            memory.WriteDouble(aOrigRect[3]); // y2
        }
        

        // new flags
        let Flags = 0;
        let nPosForFlags = memory.GetCurPosition();
        memory.Skip(4);

        // name
        if (sName != null) {
            Flags |= (1 << 0);
            memory.WriteString(sName);
        }

        // contents
        if (sContents != null) {
            Flags |= (1 << 1);
            if (typeof(sContents) != "string")
                sContents = sContents.GetContents();

            memory.WriteString(sContents);
        }

        // border effect
        if (BES != null || BEI != null) {
            Flags |= (1 << 2);
            memory.WriteByte(BES);
            memory.WriteDouble(BEI);
        }

        if (aStrokeColor != null) {
            Flags |= (1 << 3);
            memory.WriteLong(aStrokeColor.length);
            for (let i = 0; i < aStrokeColor.length; i++)
                memory.WriteDouble(aStrokeColor[i]);
        }

        if (nBorder != null || nBorderW != null) {
            Flags |= (1 << 4);
            memory.WriteByte(nBorder);
            memory.WriteDouble(nBorderW);

            if (nBorder == 2) {
                let aDash = this.GetDash();
                memory.WriteLong(aDash.length);
                for (let i = 0; i < aDash.length; i++) {
                    memory.WriteDouble(aDash[i]);
                }
            }
        }

        if (sModDate != null) {
            Flags |= (1 << 5);
            memory.WriteString(sModDate);
        }

        // render
        let nEndPos = memory.GetCurPosition();
        this.WriteRenderToBinary(memory);
        if (nEndPos != memory.GetCurPosition())
            Flags |= (1 << 6);

        // uid
        let sUserId = this.GetUserId();
        if (sUserId) {
            Flags |= (1 << 7);
            memory.WriteString(sUserId);
        }

        // meta data
        let oMeta = this.GetMeta();
        if (oMeta != null) {
            Flags |= (1 << 9);
            if (memory.isForSplit || memory.isCopyPaste) {
                if (this.IsStamp()) {
                    if (this.GetRenderStructure()) {
                        oMeta["isOO"] = true;
                        oMeta["InRect"] = this.GetInRect();
                    }
                }
                if (this.GetOriginPage() == undefined) {
                    oMeta["isOO"] = true;

                    if (this.IsRedact()) {
                        let sRedactId = this.GetRedactId();
                        if (sRedactId) {
                            oMeta["redactId"] = sRedactId;
                        }
                    }
                }
            }
            
            memory.WriteString(JSON.stringify(oMeta));
        }

        nEndPos = memory.GetCurPosition();
        memory.Seek(nPosForFlags);
        memory.WriteLong(Flags);
        memory.Seek(nEndPos);
    };
    CAnnotationBase.prototype.ReadFromBinaryBase = function(memory) {
        // type
        let type = memory.GetUChar();
    
        // apidx
        let apIdx = memory.GetLong();
    
        // annot flags
        let annotFlags = memory.GetLong();
        let bHidden = (annotFlags & (1 << 1)) !== 0;
        let bPrint = (annotFlags & (1 << 2)) !== 0;
        let noZoom = (annotFlags & (1 << 3)) !== 0;
        let noRotate = (annotFlags & (1 << 4)) !== 0;
        let bNoView = (annotFlags & (1 << 5)) !== 0;
        let locked = (annotFlags & (1 << 7)) !== 0;
        let ToggleNoView = (annotFlags & (1 << 8)) !== 0;
        let lockedC = (annotFlags & (1 << 9)) !== 0;
    
        this.SetDisplay(bHidden ? 1 : (bPrint ? (bNoView ? 3 : 0) : -1));

        // page
        let nPage = memory.GetLong();
        this.SetPage(nPage);
    
        // rect
        let x1 = memory.GetDouble();
        let y1 = memory.GetDouble();
        let x2 = memory.GetDouble();
        let y2 = memory.GetDouble();
        this.SetOrigRect([x1, y1, x2, y2]);
    
        // new flags
        let Flags = memory.GetLong();
    
        // name
        if (Flags & (1 << 0)) {
            let sName = memory.GetString();
            this.SetName(sName);
        }
    
        // contents
        if (Flags & (1 << 1)) {
            let sContents = memory.GetString();
            this.SetContents(sContents);
        }
    
        // border effect
        if (Flags & (1 << 2)) {
            let BES = memory.GetUChar();
            let BEI = memory.GetDouble();
            this.SetBorderEffectStyle(BES);
            this.SetBorderEffectIntensity(BEI);
        }
    
        // stroke color
        if (Flags & (1 << 3)) {
            let colorLength = memory.GetLong();
            let aStrokeColor = [];
            for (let i = 0; i < colorLength; i++) {
                aStrokeColor.push(memory.GetDouble());
            }
            this.SetStrokeColor(aStrokeColor);
        }
    
        // border
        if (Flags & (1 << 4)) {
            let nBorder = memory.GetUChar();
            let nBorderW = memory.GetDouble();
            this.SetBorder(nBorder);
            this.SetWidth(nBorderW);
    
            if (nBorder === 2) {
                let dashLength = memory.GetLong();
                let aDash = [];
                for (let i = 0; i < dashLength; i++) {
                    aDash.push(memory.GetDouble());
                }
                this.SetDash(aDash);
            }
        }
    
        // modification date
        if (Flags & (1 << 5)) {
            let sModDate = memory.GetString();
            this.SetModDate(sModDate);
        }
    
        // user ID
        if (Flags & (1 << 7)) {
            let sUserId = memory.GetString();
            this.SetUserId(sUserId);
        }
    };
    CAnnotationBase.prototype.WriteToBinaryBase2 = function(memory) {
        let nType = this.GetType();
        if ((nType < 18 && nType != 1 && nType != 15) || nType == 25) {
            // запишем флаги в конце
            memory.annotFlags   = 0;
            memory.posForFlags  = memory.GetCurPosition();
            memory.Skip(4);

            let nPopupIdx       = this.GetPopupIdx();
            let sAuthor         = this.GetAuthor();
            let nOpacity        = this.GetOpacity();
            let aRC             = this.GetRichContents(true);
            let CrDate          = this.GetCreationDate(true);
            let oRefTo          = this.GetReplyTo();
            let nRefToReason    = this.GetRefType();
            let sSubject        = this.GetSubject();

            if (nPopupIdx != null) {
                memory.annotFlags |= (1 << 0);
                memory.WriteLong(nPopupIdx);
            }

            if (sAuthor != null) {
                memory.annotFlags |= (1 << 1);
                memory.WriteString(sAuthor);
            }

            if (nOpacity != null) {
                memory.annotFlags |= (1 << 2);
                memory.WriteDouble(nOpacity);
            }
                
            if (aRC != null) {
                memory.annotFlags |= (1 << 3);
                memory.WriteLong(aRC.length);

                for (let i = 0; i < aRC.length; i++) {
                    memory.WriteByte(aRC[i]["alignment"]);
                    let nFontStylePos = memory.GetCurPosition();
                    memory.Skip(4);

                    // font style
                    let nStyle = 0;
                    if (aRC[i]["bold"]) {
                        nStyle |= (1 << 0);
                    }
                    if (aRC[i]["italic"]) {
                        nStyle |= (1 << 1);
                    }
                    if (aRC[i]["strikethrough"]) {
                        nStyle |= (1 << 3);
                    }
                    if (aRC[i]["underlined"]) {
                        nStyle |= (1 << 4);
                    }
                    if (aRC[i]["vertical"]) {
                        nStyle |= (1 << 5);
                        memory.WriteDouble(aRC[i]["vertical"]);
                    }
                    if (aRC[i]["actual"]) {
                        nStyle |= (1 << 6);
                        memory.WriteString(aRC[i]["actual"]);
                    }
                    // запись флагов настроек шрифта
                    let nEndPos = memory.GetCurPosition();
                    memory.Seek(nFontStylePos);
                    memory.WriteLong(nStyle);
                    memory.Seek(nEndPos);

                    memory.WriteDouble(aRC[i]["size"]);
                    aRC[i]["color"].forEach(function(component) {
                        memory.WriteDouble(component);
                    });

                    memory.WriteString(aRC[i]["name"]);
                    memory.WriteString(aRC[i]["text"]);
                }
            }

            if (CrDate != null) {
                memory.annotFlags |= (1 << 4);
                memory.WriteString(CrDate);
            }

            if (oRefTo != null) {
                memory.annotFlags |= (1 << 5);
                memory.WriteLong(oRefTo.GetApIdx());
            }

            if (nRefToReason != null) {
                memory.annotFlags |= (1 << 6);
                memory.WriteByte(nRefToReason);
            }

            if (sSubject != null) {
                memory.annotFlags |= (1 << 7);
                memory.WriteString(sSubject);
            }
        }
    };
    CAnnotationBase.prototype.ReadFromBinaryBase2 = function(memory) {
        let nType = this.GetType();
        if ((nType < 18 && nType != 1 && nType != 15) || nType == 25) {
            // Чтение флагов
            memory.annotFlags = memory.GetLong();
    
            let nPopupIdx = null;
            if (memory.annotFlags & (1 << 0)) {
                nPopupIdx = memory.GetLong();
                this.SetPopupIdx(nPopupIdx);
            }
    
            let sAuthor = null;
            if (memory.annotFlags & (1 << 1)) {
                sAuthor = memory.GetString();
                this.SetAuthor(sAuthor);
            }
    
            let nOpacity = null;
            if (memory.annotFlags & (1 << 2)) {
                nOpacity = memory.GetDouble();
                this.SetOpacity(nOpacity);
            }
    
            let aRC = null;
            if (memory.annotFlags & (1 << 3)) {
                let rcLength = memory.GetLong();
                aRC = [];
    
                for (let i = 0; i < rcLength; i++) {
                    let rcItem = {};
    
                    rcItem["alignment"] = memory.GetUChar();
    
                    let nFontStyle = memory.GetLong();
                    rcItem["bold"] = (nFontStyle & (1 << 0)) !== 0;
                    rcItem["italic"] = (nFontStyle & (1 << 1)) !== 0;
                    rcItem["strikethrough"] = (nFontStyle & (1 << 3)) !== 0;
                    rcItem["underlined"] = (nFontStyle & (1 << 4)) !== 0;
                    rcItem["vertical"] = (nFontStyle & (1 << 5)) !== 0 ? memory.GetDouble() : null;
                    rcItem["actual"] = (nFontStyle & (1 << 6)) !== 0 ? memory.GetString() : null;
    
                    rcItem["size"] = memory.GetDouble();
    
                    rcItem["color"] = [];
                    for (let j = 0; j < 4; j++) { // Assuming color has 4 components
                        rcItem["color"].push(memory.GetDouble());
                    }
    
                    rcItem["name"] = memory.GetString();
                    rcItem["text"] = memory.GetString();
    
                    aRC.push(rcItem);
                }
                this.SetRichContents(aRC);
            }
    
            let CrDate = null;
            if (memory.annotFlags & (1 << 4)) {
                CrDate = memory.GetString();
                this.SetCreationDate(CrDate);
            }
    
            let oRefTo = null;
            if (memory.annotFlags & (1 << 5)) {
                let refApIdx = memory.GetLong();
                oRefTo = this.FindAnnotationByApIdx(refApIdx); // Assuming a method to find annotation by ApIdx exists
                this.SetReplyTo(oRefTo);
            }
    
            let nRefToReason = null;
            if (memory.annotFlags & (1 << 6)) {
                nRefToReason = memory.GetUChar();
                this.SetRefType(nRefToReason);
            }
    
            let sSubject = null;
            if (memory.annotFlags & (1 << 7)) {
                sSubject = memory.GetString();
                this.SetSubject(sSubject);
            }
        }
    };
    CAnnotationBase.prototype.WriteRenderToBinary = function(memory) {
        // пока только для основанных на фигурах
        if (false == this.IsShapeBased() || this.IsNeedDrawFromStream() || !memory.docRenderer || (memory.isForSplit || memory.isCopyPaste)) {
            return;
        }

        // тут будет длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.draw(memory.docRenderer); // для каждой страницы инициализируется свой renderer

        // запись длины комманд
        let nEndPos = memory.GetCurPosition();
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };

    function ConvertPt2Px(pt) {
        return (96 / 72) * pt;
    }
    function ConvertPx2Pt(px) {
        return px / (96 / 72);
    }

    function ParsePDFDate(sDate) {
        // Регулярное выражение для извлечения компонентов даты
        let regex = /D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})([Z\+\-]?)(\d{2})?'?(\d{2})?/;

        // Используем регулярное выражение для извлечения компонентов даты
        let match = sDate.match(regex);

        if (match) {
            // Извлекаем компоненты даты из совпадения
            let year = parseInt(match[1]);
            let month = parseInt(match[2]);
            let day = parseInt(match[3]);
            let hour = parseInt(match[4]);
            let minute = parseInt(match[5]);
            let second = parseInt(match[6]);
            let timeZoneSign = match[7];
            let timeZoneOffsetHours = parseInt(match[8]);
            let timeZoneOffsetMinutes = parseInt(match[9]);

            // Создаем объект Date с извлеченными компонентами даты
            let date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));

            // Учитываем смещение времени
            if (timeZoneSign === 'Z') {
                // Если указано "Z", это означает UTC
            } else if (timeZoneSign === '+') {
                date.setHours(date.getHours() - timeZoneOffsetHours);
                date.setMinutes(date.getMinutes() - timeZoneOffsetMinutes);
            } else if (timeZoneSign === '-') {
                date.setHours(date.getHours() + timeZoneOffsetHours);
                date.setMinutes(date.getMinutes() + timeZoneOffsetMinutes);
            }

            return date;
        }

        return null;
    }

    // переопределение методов cshape
    CAnnotationBase.prototype.canRotate = function() {
        return false;
    };
    CAnnotationBase.prototype.canEditText = function () {
        return false;
    };
    CAnnotationBase.prototype.canResize = function () {
        return true;
    };
    
    function formatTimestampToPDF(timestamp) {
        const date = new Date(parseInt(timestamp));
        
        const year      = date.getFullYear();
        const month     = (date.getMonth() + 1).toString().padStart(2, '0');
        const day       = date.getDate().toString().padStart(2, '0');
        const hours     = date.getHours().toString().padStart(2, '0');
        const minutes   = date.getMinutes().toString().padStart(2, '0');
        const seconds   = date.getSeconds().toString().padStart(2, '0');
      
        // Calculate timezone offset
        let timezoneOffsetMinutes = date.getTimezoneOffset();
        
        let timezoneOffsetSign;
        if (timezoneOffsetMinutes < 0)
            timezoneOffsetSign = '+';
        else if (timezoneOffsetMinutes > 0)
            timezoneOffsetSign = '-';
        else if (timezoneOffsetMinutes == 0)
            timezoneOffsetSign = '=';

        
        let timezoneOffsetHours = Math.abs(Math.floor(timezoneOffsetMinutes / 60)) >> 0;
        if (timezoneOffsetHours < 10)
            timezoneOffsetHours = '0' + timezoneOffsetHours.toString();
        else
            timezoneOffsetHours = timezoneOffsetHours.toString();
        
        let timezoneOffsetMinutesLeft = timezoneOffsetMinutes % 60;
        if (timezoneOffsetMinutesLeft < 10)
            timezoneOffsetMinutesLeft = '0' + timezoneOffsetMinutesLeft.toString();
        else
            timezoneOffsetMinutesLeft = timezoneOffsetMinutesLeft.toString();

        const formattedTimestamp = 'D:' + year + month + day + hours + minutes + seconds + timezoneOffsetSign + timezoneOffsetHours + "'" + timezoneOffsetMinutesLeft + "'";
        
        return formattedTimestamp;
    }

	window["AscPDF"].CAnnotationBase    = CAnnotationBase;
	window["AscPDF"].ConvertPt2Px       = ConvertPt2Px;
	window["AscPDF"].ConvertPx2Pt       = ConvertPx2Pt;
    window["AscPDF"].ParsePDFDate       = ParsePDFDate;

    window["AscPDF"].startMultiplyMode = function(ctx)
    {
        if (!AscCommon.AscBrowser.isIE)
            ctx.globalCompositeOperation = "multiply";
        else
        {
            ctx._multiplyGlobalAlpha = ctx.globalAlpha;
            ctx.globalAlpha = 0.3 * ctx._multiplyGlobalAlpha;
        }
    };
    window["AscPDF"].endMultiplyMode = function(ctx)
    {
        if (!AscCommon.AscBrowser.isIE)
            ctx.globalCompositeOperation = "source-over";
        else
        {
            ctx.globalAlpha = ctx._multiplyGlobalAlpha;
            delete ctx._multiplyGlobalAlpha;
        }
    };

})();

