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

    let BORDER_EFFECT_STYLES = {
        None:   0,
        Cloud:  1
    }
	/**
	 * Class representing a base annotation.
	 * @constructor
    */
    function CAnnotationBase(sName, nType, nPage, aRect, oDoc)
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
        this._rect                  = aRect;
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

        // internal
        this._bDrawFromStream   = false; // нужно ли рисовать из стрима
        this._id                = AscCommon.g_oIdCounter.Get_NewId();
        this._originView = {
            normal:     null,
            mouseDown:  null,
            rollover:   null
        }
        this._wasChanged            = false;
    }
    CAnnotationBase.prototype.SetReqtangleDiff = function(aDiff) {
        this._rectDiff = aDiff;
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
    CAnnotationBase.prototype.SetFillColor = function(aColor) {
        this._fillColor = aColor;
    };
    CAnnotationBase.prototype.SetWidth = function(nWidth) {
        this._width = nWidth;
    };
    CAnnotationBase.prototype.SetRichContents = function(sText) {
        this._richContents = sText;
    };
    CAnnotationBase.prototype.SetIntent = function(nType) {
        this._intent = nType;
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
    CAnnotationBase.prototype.SetBorderEffectIntensity = function(nValue) {
        this._borderEffectIntensity = nValue;
    };
    CAnnotationBase.prototype.SetBorderEffectStyle = function(nStyle) {
        this._borderEffectIntensity = nStyle;
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
                oGraphicsPDF.context.globalCompositeOperation = "multiply";
            
            oGraphicsPDF.DrawImage(originView, 0, 0, nWidth, nHeight, X, Y, nWidth, nHeight);
            oGraphicsPDF.context.globalCompositeOperation = "source-over";
        }
    };
    // CAnnotationBase.prototype.SetCaptionPositioning
    CAnnotationBase.prototype.SetSubject = function(sSubject) {
        this._subject = sSubject;
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
            // sticky note всегда одинаковый при зуме
            if (this.IsComment() || oSavedView.width == oApearanceInfo["w"] && oSavedView.height == oApearanceInfo["h"]) {
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
        let nPage = this.GetPage();

        oDoc.History.Add(new CChangesPDFAnnotPos(this, [this._rect[0], this._rect[1]], [x, y]));

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

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

        this.AddToRedraw();
        this.SetWasChanged(true);
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
    CAnnotationBase.prototype.GetOrigRect = function() {
        return this._origRect;
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

        if (this.IsInDocument() == false || oDoc.History.UndoRedoInProgress) {
            oDoc.CreateNewHistoryPoint();
            oDoc.History.Add(new CChangesPDFAnnotRect(this, this.GetRect(), aRect));
            oDoc.TurnOffHistory();
        }

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        this._rect = aRect;

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
    CAnnotationBase.prototype.IsInDocument = function() {
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
    
                if (this.IsInDocument() && oViewer.pagesInfo.pages[nPage].annots.indexOf(this) == -1)
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
        let oPos = AscPDF.GetGlobalCoordsByPageCoords(this._pagePos.x + this._pagePos.w, this._pagePos.y + this._pagePos.h / 2, this.GetPage(), true);
        editor.sync_ShowComment([this.GetId()], oPos["X"], oPos["Y"])
    };
    CAnnotationBase.prototype._AddReplyOnOpen = function(oReplyInfo) {
        let oReply = new AscPDF.CAnnotationText(oReplyInfo["UniqueName"], this.GetPage(), [], this.GetDocument());

        oReply.SetContents(oReplyInfo["Contents"]);
        oReply.SetModDate(AscPDF.ParsePDFDate(oReplyInfo["LastModified"]).getTime());
        oReply.SetAuthor(oReplyInfo["User"]);
        oReply.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);

        if (this.IsComment())
            this._replies.push(oReply);
        else
            this._contents._replies.push(oReply);
    };
    CAnnotationBase.prototype._OnAfterSetContents = function() {
        let oAscCommData = this.IsComment() ? this.GetAscCommentData() : this._contents.GetAscCommentData();
        editor.sendEvent("asc_onAddComment", this.GetId(), oAscCommData);
    };
    CAnnotationBase.prototype.SetContents = function(contents) {
        if (this.GetContents() == contents)
            return;

        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oCurContents    = this.GetContents();
        let oNewContents;

        if (typeof(contents) == "string") {
            let oTextAnnot = new AscPDF.CAnnotationText(this.GetName(), this.GetPage(), [], this.GetDocument());
        
            oTextAnnot.SetContents(contents);
            oTextAnnot.SetModDate(this.GetModDate());
            oTextAnnot.SetAuthor(this.GetAuthor());
            oTextAnnot.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);

            this._contents  = oTextAnnot;
            oNewContents    = oTextAnnot;
        }
        else {
            oNewContents    = contents;
            this._contents  = contents;
            
            if (contents)
                this._OnAfterSetContents();
        }

        if (oDoc.History.UndoRedoInProgress == false && oViewer.IsOpenAnnotsInProgress == false) {
            oDoc.CreateNewHistoryPoint();
            oDoc.History.Add(new CChangesPDFAnnotContents(this, oCurContents, oNewContents));
            oDoc.TurnOffHistory();
        }
        
        if (oNewContents == null)
            editor.sync_RemoveComment(this.GetId());

        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.AddReply = function(CommentData) {
        this._contents.AddReply(CommentData);
    };
    CAnnotationBase.prototype.GetAscCommentData = function() {
        if (this._contents)
            return this._contents.GetAscCommentData();

        return null;
    };
    CAnnotationBase.prototype.EditCommentData = function(CommentData) {
        if (this._contents)
            this._contents.EditCommentData(CommentData);
    };
    CAnnotationBase.prototype.GetContents = function() {
        return this._contents;
    };
    CAnnotationBase.prototype.SetModDate = function(sDate) {
        this._modDate = sDate;
        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.GetModDate = function() {
        return this._modDate;
    };
    CAnnotationBase.prototype.SetCreationDate = function(sDate) {
        this._creationDate = sDate;
        this.SetWasChanged(true);
    };
    CAnnotationBase.prototype.GetCreationDate = function() {
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
        this._apIdx = nIdx;
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

        if (aInternalColor.length == 1) {
            oColor = {
                r: aInternalColor[0] * 255,
                g: aInternalColor[0] * 255,
                b: aInternalColor[0] * 255
            }
        }
        else if (aInternalColor.length == 3) {
            oColor = {
                r: aInternalColor[0] * 255,
                g: aInternalColor[1] * 255,
                b: aInternalColor[2] * 255
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

        let oNewAnnot = new CAnnotationBase(AscCommon.CreateGUID(), this.type, this.GetPage(), this.GetRect().slice(), oDoc);

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
    CAnnotationBase.prototype.Draw = function(oGraphics) {
        this.DrawFromStream(oGraphics);
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

    // аналоги методов Drawings
    CAnnotationBase.prototype.getObjectType = function() {
        return -1;
    };
    CAnnotationBase.prototype.hitInTextRect = function() {
        let oViewer = editor.getDocumentRenderer();

        if (oViewer.getPageAnnotByMouse() == this)
            return true;

        return false;
    };
    CAnnotationBase.prototype.getMediaFileName = function() {
        return false;
    };
    CAnnotationBase.prototype.canEdit = function() {
        return false;
    };
    // // заглушки для трекинга
    // CAnnotationBase.prototype.canChangeAdjustments = function() {
    //     return true;
    // };
    // CAnnotationBase.prototype.hitToAdjustment = function () {
    //     return {hit: false, adjPolarFlag: null, adjNum: null, warp: false};
    // };
    // CAnnotationBase.prototype.getObjectType = function() {
    //     return -1;
    // };

    function ConvertPt2Px(pt) {
        return (96 / 72) * pt;
    }
    function ConvertPx2Pt(px) {
        return px / (96 / 72);
    }

    function ParsePDFDate(sDate) {
        // Регулярное выражение для извлечения компонентов даты
        let regex = /D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})([Z\+\-])(\d{2})'(\d{2})/;

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

	window["AscPDF"].CAnnotationBase    = CAnnotationBase;
	window["AscPDF"].ConvertPt2Px       = ConvertPt2Px;
	window["AscPDF"].ConvertPx2Pt       = ConvertPx2Pt;
    window["AscPDF"].ParsePDFDate       = ParsePDFDate;
})();

