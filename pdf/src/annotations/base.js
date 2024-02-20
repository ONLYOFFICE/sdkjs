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
	 * Class representing a base annotation.
	 * @constructor
    */
    function CAnnotationBase(sName, nType, nPage, aOrigRect, oDoc)
    {
        this.type = nType;

        this._author                = undefined;
        this._borderEffectIntensity = undefined;
        this._borderEffectStyle     = undefined;
        this._contents              = undefined;
        this._creationDate          = undefined;
        this._delay                 = false; // пока не используется
        this._doc                   = oDoc;
        this._inReplyTo             = undefined;
        this._intent                = undefined;
        this._lock                  = undefined;
        this._lockContent           = undefined;
        this._modDate               = undefined;
        this._name                  = sName;
        this._opacity               = 1;
        this._page                  = nPage;
        this._rect                  = undefined;
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

        this._replies               = []; // тут будут храниться ответы (text аннотации)

        // internal
        this._bDrawFromStream   = false; // нужно ли рисовать из стрима
        this._id                = AscCommon.g_oIdCounter.Get_NewId();
        this._originView = {
            normal:     null,
            mouseDown:  null,
            rollover:   null
        }
        this._wasChanged            = false;
        this.Internal_InitRect(aOrigRect);
    }
    CAnnotationBase.prototype.Internal_InitRect = function(aOrigRect) {
        let nPage = this.GetPage();
        let oViewer = editor.getDocumentRenderer();
        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        this._rect = [aOrigRect[0] * nScaleX, aOrigRect[1] * nScaleY, aOrigRect[2] * nScaleX, aOrigRect[3] * nScaleY];
        this._origRect = aOrigRect;
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
        this._rectDiff = aDiff;
    };
    CAnnotationBase.prototype.GetRectangleDiff = function() {
        return this._rectDiff;
    };
    CAnnotationBase.prototype.SetNoRotate = function(bValue) {
        this._noRotate = bValue;
    };
    CAnnotationBase.prototype.SetNoZoom = function(bValue) {
        this._noZoom = bValue;
    };
    CAnnotationBase.prototype.SetDash = function(aDash) {
        this._dash = aDash;
    };
    CAnnotationBase.prototype.GetDash = function() {
        return this._dash;
    };
    CAnnotationBase.prototype.SetFillColor = function(aColor) {
        this._fillColor = aColor;
    };
    CAnnotationBase.prototype.GetFillColor = function() {
        return this._fillColor;
    };
    CAnnotationBase.prototype.SetWidth = function(nWidth) {
        this._width = nWidth;
    };
    CAnnotationBase.prototype.GetWidth = function() {
        return this._width;
    };
    CAnnotationBase.prototype.SetRichContents = function(sText) {
        this._richContents = sText;
    };
    CAnnotationBase.prototype.GetRichContents = function() {
        return this._richContents;
    };
    CAnnotationBase.prototype.SetIntent = function(nType) {
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
        this._border = nType;
    };
    CAnnotationBase.prototype.GetBorder = function() {
        return this._border;
    };
    CAnnotationBase.prototype.SetBorderEffectIntensity = function(nValue) {
        this._borderEffectIntensity = nValue;
    };
    CAnnotationBase.prototype.GetBorderEffectIntensity = function() {
        return this._borderEffectIntensity;
    };
    CAnnotationBase.prototype.SetBorderEffectStyle = function(nStyle) {
        this._borderEffectStyle = nStyle;
    };
    CAnnotationBase.prototype.GetBorderEffectStyle = function() {
        return this._borderEffectStyle;
    };

    CAnnotationBase.prototype.DrawSelected = function(overlay) {
        let rPR         = AscCommon.AscBrowser.retinaPixelRatio;
        let style_blue  = "#939393";
        let indent      = 0.5 * Math.round(rPR);

        let nPage       = this.GetPage();
        let oViewer     = editor.getDocumentRenderer();
        let nScale      = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom * (96 / oViewer.file.pages[nPage].Dpi);
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

        
        let x1 = Math.round(indLeft + aOrigRect[0] * nScale);
        let y1 = Math.round(indTop + aOrigRect[1] * nScale);
        let x2 = Math.round(indLeft + aOrigRect[2] * nScale + 0.5);
        let y2 = Math.round(indTop + aOrigRect[3] * nScale + 0.5);

        overlay.m_oContext.lineWidth    = Math.round(rPR);
        overlay.m_oContext.globalAlpha  = 1;
        overlay.m_oContext.strokeStyle  = style_blue;
        overlay.m_oContext.beginPath();

        overlay.CheckPoint1(x1, y1);
        overlay.CheckPoint1(x2, y2);
        overlay.CheckPoint2(x1, y1);
        overlay.CheckPoint2(x2, y2);
        
        overlay.m_oContext.rect(x1 + indent, y1 + indent, x2 - x1, y2 - y1);
        overlay.m_oContext.stroke();
    };
    CAnnotationBase.prototype.GetName = function() {
        return this._name;
    };
    CAnnotationBase.prototype.SetOpacity = function(value) {
        this._opacity = value;
        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.GetOpacity = function() {
        return this._opacity;
    };
    CAnnotationBase.prototype.IsPdfObject = function() {
        return true;
    };
    /**
	 * Invokes only on open forms.
	 * @memberof CAnnotationBase
	 * @typeofeditors ["PDF"]
	 */
    CAnnotationBase.prototype.SetOriginPage = function(nPage) {
        this._origPage = nPage;
    };
    CAnnotationBase.prototype.GetOriginPage = function() {
        return this._origPage;
    };
    CAnnotationBase.prototype.SetWasChanged = function(isChanged) {
        let oViewer = editor.getDocumentRenderer();

        if (oViewer.IsOpenAnnotsInProgress == false) {
            this._wasChanged = isChanged;
        }
    };
    CAnnotationBase.prototype.IsChanged = function() {
        return this._wasChanged;  
    };
    CAnnotationBase.prototype.DrawFromStream = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;
            
        let originView      = this.GetOriginView();
        let nGrScale        = oGraphicsPDF.GetScale();

        if (originView) {
            let aOrigRect       = this.GetOrigRect();
            
            let X       = this.IsTextMarkup() ? originView.x / nGrScale : aOrigRect[0];
            let Y       = this.IsTextMarkup() ? originView.y / nGrScale : aOrigRect[1];
            let nWidth  = originView.width / nGrScale;
            let nHeight = originView.height / nGrScale;

            if (this.IsHighlight())
                AscPDF.startMultiplyMode(oGraphicsPDF.context);
            
            oGraphicsPDF.DrawImage(originView, 0, 0, nWidth, nHeight, X, Y, nWidth, nHeight);
            AscPDF.endMultiplyMode(oGraphicsPDF.context);
        }
    };
    CAnnotationBase.prototype.SetSubject = function(sSubject) {
        this._subject = sSubject;
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
    CAnnotationBase.prototype.GetOriginView = function() {
        if (this._apIdx == -1)
            return null;

        let oViewer = editor.getDocumentRenderer();
        let oFile   = oViewer.file;
        
        let oApearanceInfo  = this.GetOriginViewInfo();
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
        let mappedBuffer    = new Uint8ClampedArray(oFile.memory().buffer, oApInfoTmp["retValue"], 4 * nWidth * nHeight);
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
     * Returns AP info of this field.
	 * @memberof CAnnotationBase
	 * @typeofeditors ["PDF"]
     * @returns {Object}
	 */
    CAnnotationBase.prototype.GetOriginViewInfo = function() {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let nPage       = this.GetOriginPage();
        let oOriginPage = oFile.pages.find(function(page) {
            return page.originIndex == nPage;
        });

        let w = ((oOriginPage.W * 96 / oOriginPage.Dpi) >> 0) * AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom >> 0;
        let h = ((oOriginPage.H * 96 / oOriginPage.Dpi) >> 0) * AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom >> 0;

        if (oOriginPage.annotsAPInfo == null || oOriginPage.annotsAPInfo.size.w != w || oOriginPage.annotsAPInfo.size.h != h) {
            oOriginPage.annotsAPInfo = {
                info: oFile.nativeFile["getAnnotationsAP"](nPage, w, h),
                size: {
                    w: w,
                    h: h
                }
            }
        }
        
        for (let i = 0; i < oOriginPage.annotsAPInfo.info.length; i++) {
            if (oOriginPage.annotsAPInfo.info[i]["i"] == this._apIdx)
                return oOriginPage.annotsAPInfo.info[i];
        }

        return null;
    };
    
    CAnnotationBase.prototype.ClearCache = function() {
        this._originView.normal = null;
    };
    CAnnotationBase.prototype.SetPosition = function(x, y) {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = this.GetDocument();
        let nPage   = this.GetPage();

        let nOldX = this._rect[0];
        let nOldY = this._rect[1];

        let nDeltaX = x - nOldX;
        let nDeltaY = y - nOldY;

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        if (this.IsInk()) {
            let aPath;
            for (let i = 0; i < this._gestures.length; i++) {
                aPath = this._gestures[i];
                for (let j = 0; j < aPath.length; j++) {
                    aPath[j].x += nDeltaX * g_dKoef_pix_to_mm;
                    aPath[j].y += nDeltaY * g_dKoef_pix_to_mm;
                }
            }
        }
        else if (this.IsLine()) {
            for (let i = 0; i < this._points.length; i+=2) {
                this._points[i] += nDeltaX / nScaleX;
                this._points[i+1] += nDeltaY / nScaleY;
            }
        }
        else if (this.IsPolygon() || this.IsPolyLine()) {
            for (let i = 0; i < this._vertices.length; i+=2) {
                this._vertices[i] += nDeltaX / nScaleX;
                this._vertices[i+1] += nDeltaY / nScaleY;
            }
        }

        oDoc.History.Add(new CChangesPDFAnnotPos(this, [this._rect[0], this._rect[1]], [x, y]));

        let nWidth  = this._pagePos.w;
        let nHeight = this._pagePos.h;

        this._rect[0] = x;
        this._rect[1] = y;
        this._rect[2] = x + nWidth;
        this._rect[3] = y + nHeight;
        
        this._origRect[0] = this._rect[0] / nScaleX;
        this._origRect[1] = this._rect[1] / nScaleY;
        this._origRect[2] = this._rect[2] / nScaleX;
        this._origRect[3] = this._rect[3] / nScaleY;

        this._pagePos = {
            x: this._rect[0],
            y: this._rect[1],
            w: (this._rect[2] - this._rect[0]),
            h: (this._rect[3] - this._rect[1])
        };

        this.SetNeedRecalc(true);
        this.AddToRedraw();
        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.IsShapeBased = function() {
        return this instanceof AscFormat.CShape || this instanceof AscFormat.CGroupShape;
    };
    CAnnotationBase.prototype.IsHighlight = function() {
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
    CAnnotationBase.prototype.SetNeedRecalc = function(bRecalc) {
        this._needRecalc = bRecalc;
    };
    CAnnotationBase.prototype.IsNeedRecalc = function() {
        return this._needRecalc;
    };
    CAnnotationBase.prototype.GetOrigRect = function() {
        return this._origRect || this.GetReplyTo().GetOrigRect();
    };
    CAnnotationBase.prototype.IsNeedDrawFromStream = function() {
        return this._bDrawFromStream;
    };
    CAnnotationBase.prototype.SetDrawFromStream = function(bFromStream) {
        this._bDrawFromStream = bFromStream;
    };
    CAnnotationBase.prototype.SetRect = function(aRect) {
        let oViewer = editor.getDocumentRenderer();
        let nPage = this.GetPage();
        let oDoc = this.GetDocument();

        oDoc.History.Add(new CChangesPDFAnnotRect(this, this.GetRect(), aRect));

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        this._rect = aRect;
        this._rectDiff = [0, 0, 0, 0];

        this._pagePos = {
            x: aRect[0],
            y: aRect[1],
            w: (aRect[2] - aRect[0]),
            h: (aRect[3] - aRect[1])
        };

        this._origRect[0] = this._rect[0] / nScaleX;
        this._origRect[1] = this._rect[1] / nScaleY;
        this._origRect[2] = this._rect[2] / nScaleX;
        this._origRect[3] = this._rect[3] / nScaleY;

        this.SetWasChanged(true);
        if (oViewer.IsOpenAnnotsInProgress == false) {
            this.SetDrawFromStream(false);
        }
    };
    CAnnotationBase.prototype.IsUseInDocument = function() {
        if (this.GetDocument().annots.indexOf(this) == -1)
            return false;

        return true;
    };

    
    CAnnotationBase.prototype.GetRect = function() {
        return this._rect;
    };
    CAnnotationBase.prototype.GetId = function() {
        return this._id;
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
    CAnnotationBase.prototype.SetPage = function(nPage) {
        let nCurPage = this.GetPage();
        if (nPage == nCurPage)
            return;

        let oViewer = editor.getDocumentRenderer();
        let oDoc    = this.GetDocument();
        
        let nCurIdxOnPage = oViewer.pagesInfo.pages[nCurPage].annots ? oViewer.pagesInfo.pages[nCurPage].annots.indexOf(this) : -1;
        if (oViewer.pagesInfo.pages[nPage]) {
            if (oDoc.annots.indexOf(this) != -1) {
                if (oViewer.pagesInfo.pages[nPage].annots == null) {
                    oViewer.pagesInfo.pages[nPage].annots = [];
                }
    
                if (nCurIdxOnPage != -1)
                    oViewer.pagesInfo.pages[nCurPage].annots.splice(nCurIdxOnPage, 1);
    
                if (this.IsUseInDocument() && oViewer.pagesInfo.pages[nPage].annots.indexOf(this) == -1)
                    oViewer.pagesInfo.pages[nPage].annots.push(this);

                oDoc.History.Add(new CChangesPDFAnnotPage(this, nCurPage, nPage));

                // добавляем в перерисовку исходную страницу
                this.AddToRedraw();
            }

            this._page = nPage;
            this.selectStartPage = nPage;
            this.AddToRedraw();
        }
    };
    CAnnotationBase.prototype.GetPage = function() {
        return this._page;
    };
    CAnnotationBase.prototype.GetDocument = function() {
        return this._doc;
    };
    CAnnotationBase.prototype.IsHidden = function() {
        let nType = this.GetDisplay();
        if (nType == window["AscPDF"].Api.Objects.display["hidden"] || nType == window["AscPDF"].Api.Objects.display["noView"])
            return true;

        return false;
    };
    CAnnotationBase.prototype.SetDisplay = function(nType) {
        this._display = nType;
    };
    CAnnotationBase.prototype.GetDisplay = function() {
        return this._display;
    };
    CAnnotationBase.prototype.onMouseUp = function() {
        this.GetDocument().ShowComment([this.GetId()]);
    };
    CAnnotationBase.prototype._AddReplyOnOpen = function(oReplyInfo) {
        let oReply = new AscPDF.CAnnotationText(oReplyInfo["UniqueName"], this.GetPage(), [], this.GetDocument());

        oReply.SetContents(oReplyInfo["Contents"]);
        oReply.SetCreationDate(AscPDF.ParsePDFDate(oReplyInfo["CreationDate"]).getTime());
        oReply.SetModDate(AscPDF.ParsePDFDate(oReplyInfo["LastModified"]).getTime());
        oReply.SetAuthor(oReplyInfo["User"]);
        oReply.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);
        oReply.SetPopupIdx(oReplyInfo["Popup"]);
        oReply.SetSubject(oReplyInfo["Subj"]);

        oReply.SetReplyTo(this);
        oReply.SetApIdx(oReplyInfo["AP"]["i"]);
        
        this._replies.push(oReply);
    };
    CAnnotationBase.prototype.SetContents = function(contents) {
        if (this.GetContents() == contents)
            return;

        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oCurContents    = this.GetContents();

        let bSendAddCommentEvent = false;
        if (this._contents == null && contents != null)
            bSendAddCommentEvent = true;
        
        this._contents  = contents;
        
        if (oDoc.History.UndoRedoInProgress == false && oViewer.IsOpenAnnotsInProgress == false) {
            oDoc.History.Add(new CChangesPDFAnnotContents(this, oCurContents, contents));
        }
        
        this.SetWasChanged(true);
        if (bSendAddCommentEvent)
            oDoc.CheckComment(this);
        
        if (this._contents == null && this.IsUseInDocument())
            editor.sync_RemoveComment(this.GetId());
    };
    CAnnotationBase.prototype.IsUseContentAsComment = function() {
        return !(this.IsFreeText() || this.IsLine() && this.IsDoCaption());
    };
    CAnnotationBase.prototype.Recalculate = function() {
        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let aOrigRect   = this.GetOrigRect();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;
        
        this.handleUpdatePosition();
        this.recalculate();
        this.updatePosition(aOrigRect[0] * g_dKoef_pix_to_mm * nScaleX, aOrigRect[1] * g_dKoef_pix_to_mm * nScaleY)
    };
    CAnnotationBase.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        this.Recalculate();
        this.draw(oGraphicsWord);
    };
    CAnnotationBase.prototype.SetReplies = function(aReplies) {
        let oDoc = this.GetDocument();
        let oViewer = editor.getDocumentRenderer();

        if (oDoc.History.UndoRedoInProgress == false && oViewer.IsOpenAnnotsInProgress == false) {
            oDoc.History.Add(new CChangesPDFAnnotReplies(this, this._replies, aReplies));
        }
        this._replies = aReplies;
    };
    CAnnotationBase.prototype.GetReplies = function() {
        return this._replies;
    };
    CAnnotationBase.prototype.GetReply = function(nPos) {
        return this._replies[nPos];
    };
    CAnnotationBase.prototype.RemoveComment = function() {
        let oDoc = this.GetDocument();

        oDoc.CreateNewHistoryPoint();
        this.SetContents(null);
        this.SetReplies([]);
        oDoc.TurnOffHistory();
    };
    CAnnotationBase.prototype.EditCommentData = function(oCommentData) {
        let oFirstCommToEdit;
        if (this.GetApIdx() == oCommentData.m_sUserData)
            oFirstCommToEdit = this;
        else {
            oFirstCommToEdit = this._replies.find(function(oReply) {
                return oCommentData.m_sUserData == oReply.GetApIdx(); 
            });
        }
        
        if (oFirstCommToEdit.GetContents() != oCommentData.m_sText) {
            oFirstCommToEdit.SetContents(oCommentData.m_sText);
            oFirstCommToEdit.SetModDate(oCommentData.m_sOOTime);
        }

        let aReplyToDel = [];
        let oReply, oReplyCommentData;
        for (let i = 0; i < this._replies.length; i++) {
            oReply = this._replies[i];
            if (oFirstCommToEdit == oReply)
                continue;

            oReplyCommentData = oCommentData.m_aReplies.find(function(item) {
                return item.m_sUserData == oReply.GetApIdx(); 
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
                return oReplyCommentData.m_sUserData == reply.GetApIdx();
            })) {
                AscPDF.CAnnotationText.prototype.AddReply.call(this, oReplyCommentData);
            }
        }

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
    };
    CAnnotationBase.prototype.GetAscCommentData = function() {
        let oAscCommData = new Asc["asc_CCommentDataWord"](null);
        oAscCommData.asc_putText(this.GetContents());
        let sModDate = this.GetModDate();
        if (sModDate)
            oAscCommData.asc_putOnlyOfficeTime(sModDate.toString());
        oAscCommData.asc_putUserId(editor.documentUserId);
        oAscCommData.asc_putUserName(this.GetAuthor());
        oAscCommData.asc_putSolved(false);
        oAscCommData.asc_putQuoteText("");
        oAscCommData.m_sUserData = this.GetApIdx();

        this._replies.forEach(function(reply) {
            oAscCommData.m_aReplies.push(reply.GetAscCommentData());
        });

        return oAscCommData;
    };
    CAnnotationBase.prototype.GetContents = function() {
        return this._contents;
    };
    CAnnotationBase.prototype.SetModDate = function(sDate) {
        this._modDate = sDate;
        this.SetWasChanged(true);
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
        this._creationDate = sDate;
        this.SetWasChanged(true);
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
        this._author = sAuthor;
        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.GetAuthor = function() {
        return this._author;
    };
    CAnnotationBase.prototype.IsAnnot = function() {
        return true;
    };
    CAnnotationBase.prototype.SetApIdx = function(nIdx) {
        this.GetDocument().UpdateApIdx(nIdx);
        this._apIdx = nIdx;
    };
    CAnnotationBase.prototype.GetApIdx = function() {
        return this._apIdx;
    };
    CAnnotationBase.prototype.AddToRedraw = function() {
        let oViewer = editor.getDocumentRenderer();
        if (oViewer.pagesInfo.pages[this.GetPage()])
            oViewer.pagesInfo.pages[this.GetPage()].needRedrawAnnots = true;
    };
    /**
	 * Gets rgb color object from internal color array.
	 * @memberof CAnnotationBase
	 * @typeofeditors ["PDF"]
     * @returns {object}
	 */
    CAnnotationBase.prototype.GetRGBColor = function(aInternalColor) {
        let oColor = {};

        if (!aInternalColor)
            return {
                r: 255,
                g: 255,
                b: 255
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

        return oColor;
    };
    CAnnotationBase.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.TurnOffHistory();

        let oNewAnnot = new CAnnotationBase(AscCommon.CreateGUID(), this.type, this.GetPage(), this.GetOrigRect().slice(), oDoc);

        if (this._pagePos) {
            oNewAnnot._pagePos = {
                x: this._pagePos.x,
                y: this._pagePos.y,
                w: this._pagePos.w,
                h: this._pagePos.h
            }
        }

        if (this._origRect)
            oNewAnnot._origRect = this._origRect.slice();

        oNewAnnot._originView = this._originView;
        oNewAnnot._apIdx = this._apIdx;
        oNewAnnot.SetOriginPage(this.GetOriginPage());
        oNewAnnot.SetAuthor(this.GetAuthor());
        oNewAnnot.SetModDate(this.GetModDate());
        oNewAnnot.SetCreationDate(this.GetCreationDate());
        oNewAnnot.SetContents(this.GetContents());

        return oNewAnnot;
    };

    CAnnotationBase.prototype.onMouseDown = function(e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        let oDoc            = this.GetDocument();
        let oDrDoc          = oDoc.GetDrawingDocument();

        this.selectStartPage = this.GetPage();
        let oPos    = oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        let pageObject = oViewer.getPageByCoords3(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
    };
    CAnnotationBase.prototype.createMoveTrack = function() {
        return new AscFormat.MoveAnnotationTrack(this);
    };

    CAnnotationBase.prototype.SetStrokeColor = function(aColor) {
        let oViewer = editor.getDocumentRenderer();

        this._strokeColor = aColor;
        this.SetWasChanged(true);
        if (oViewer.IsOpenAnnotsInProgress == false) {
            this.SetDrawFromStream(false);
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

        memory.WriteLong(this.GetOriginPage());

        // rect
        let aOrigRect = this.GetOrigRect();
        memory.WriteDouble(aOrigRect[0]); // x1
        memory.WriteDouble(aOrigRect[1]); // y1
        memory.WriteDouble(aOrigRect[2]); // x2
        memory.WriteDouble(aOrigRect[3]); // y2

        // new flags
        let Flags = 0;
        let sName           = this.GetName();
        let sContents       = this.GetContents();
        let BES             = this.GetBorderEffectStyle();
        let BEI             = this.GetBorderEffectIntensity();
        let aStrokeColor    = this.IsFreeText() || this.IsComment() ? this.GetFillColor() : this.GetStrokeColor();
        let nBorder         = this.GetBorder();
        let nBorderW        = this.GetWidth();
        let sModDate        = this.GetModDate(true);

        if (sName != null)
            Flags |= (1 << 0);

        if (sContents != null)
            Flags |= (1 << 1);
        
        if (BES != null || BEI != null)
            Flags |= (1 << 2);
        if (aStrokeColor != null)
            Flags |= (1 << 3);
        if (nBorder != null || nBorderW != null)
            Flags |= (1 << 4);
        if (sModDate != null)
            Flags |= (1 << 5);
        
        memory.WriteLong(Flags);

        // name
        if (sName)
            memory.WriteString(sName);

        // contents
        if (sContents != null) {
            if (typeof(sContents) != "string")
                sContents = sContents.GetContents();

            memory.WriteString(sContents);
        }

        // border effect
        if (BES != null || BEI != null) {
            memory.WriteByte(BES);
            memory.WriteDouble(BEI);
        }

        if (aStrokeColor != null) {
            memory.WriteLong(aStrokeColor.length);
            for (let i = 0; i < aStrokeColor.length; i++)
                memory.WriteDouble(aStrokeColor[i]);
        }

        if (nBorder != null || nBorderW != null) {
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
            memory.WriteString(sModDate);
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
            let sRC             = this.GetRichContents();
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
                
            if (sRC != null) {
                memory.annotFlags |= (1 << 3);
                memory.WriteString(sRC);
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

