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
	 * Class representing a link annotation.
	 * @constructor
     * @extends {CAnnotationBase}
	 */
    function CAnnotationLink(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Link, aRect, oDoc);
        
        AscPDF.initShape(this);
        let oGeometry = AscFormat.CreateGeometry("rect");
        this.spPr.setGeometry(oGeometry);

        this._triggers      = new AscPDF.CPdfTriggers();
        this._quads         = [];
        this._rectDiff      = [0, 0, 0, 0];
        
        // states
        this._pressed = false;
        this._hovered = false;
    };
    
    CAnnotationLink.prototype.constructor = CAnnotationLink;
    AscFormat.InitClass(CAnnotationLink, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Link);
    Object.assign(CAnnotationLink.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationLink.prototype.IsLink = function() {
        return true;
    };
    CAnnotationLink.prototype.SetQuads = function(aFullQuads) {
        let oThis = this;

        for (let i = 0, nCount = this._quads.length; i < nCount; i++) {
            this.RemoveQuads(0);
        }

        aFullQuads.forEach(function(aQuads) {
            oThis.AddQuads(aQuads);
        });

        if (aFullQuads.length == 1) {
            this.SetNeedRecalcSizes(true);
        }
    };
    CAnnotationLink.prototype.GetQuads = function() {
        return this._quads;
    };
    CAnnotationLink.prototype.RecalcSizes = function() {
        if (!this.IsShapeBased()) {
            return;
        }

        let nLineW = this.GetWidth();
        let aQuads = this.GetQuads();
        
        AscCommon.History.StartNoHistoryMode();
        if (aQuads.length == 0 || aQuads.length > 1) {
            let aRect = this.GetRect();
            let aRD = [nLineW / 2, nLineW / 2, nLineW / 2, nLineW / 2];

            let extX = ((aRect[2] - aRect[0]) - aRD[0] - aRD[2]) * g_dKoef_pt_to_mm;
            let extY = ((aRect[3] - aRect[1]) - aRD[1] - aRD[3]) * g_dKoef_pt_to_mm;

            this.spPr.xfrm.setOffX((aRect[0] + aRD[0]) * g_dKoef_pt_to_mm);
            this.spPr.xfrm.setOffY((aRect[1] + aRD[1]) * g_dKoef_pt_to_mm);

            this.spPr.xfrm.setExtX(extX);
            this.spPr.xfrm.setExtY(extY);
        }
        else {
            let x1 = aQuads[0][0], y1 = aQuads[0][1]; // left up
            let x2 = aQuads[0][2], y2 = aQuads[0][3]; // right up
            let x3 = aQuads[0][4], y3 = aQuads[0][5]; // left down
            let x4 = aQuads[0][6], y4 = aQuads[0][7]; // right down

            let dx = x2 - x1;
            let dy = y2 - y1;

            let width  = Math.sqrt(dx * dx + dy * dy);
            let height = Math.sqrt((x3 - x1) * (x3 - x1) + (y3 - y1) * (y3 - y1));

            let angle = Math.atan2(dy, dx);

            if (angle < 0) angle += Math.PI;

            this.spPr.xfrm.setRot(angle);

            let cx = (x1 + x2 + x3 + x4) / 4;
            let cy = (y1 + y2 + y3 + y4) / 4;

            this.spPr.xfrm.setOffX((cx - width  / 2) * g_dKoef_pt_to_mm);
            this.spPr.xfrm.setOffY((cy - height / 2) * g_dKoef_pt_to_mm);

            this.spPr.xfrm.setExtX(width * g_dKoef_pt_to_mm);
            this.spPr.xfrm.setExtY(height * g_dKoef_pt_to_mm);
        }
        
        AscCommon.History.EndNoHistoryMode();
        this.SetNeedRecalcSizes(false);
    };
    CAnnotationLink.prototype.AddQuads = function(aQuads) {
        AscCommon.History.Add(new CChangesPDFAnnotQuads(this, this._quads.length, aQuads, true));
        this._quads.push(aQuads);
    };
    CAnnotationLink.prototype.RemoveQuads = function(nIdx) {
        AscCommon.History.Add(new CChangesPDFAnnotQuads(this, nIdx, this._quads[nIdx], false));
        this._quads.splice(nIdx, 1);
    };

    CAnnotationLink.prototype.DrawFromStream = function(oGraphicsPDF, oGraphicsWord) {
        let oViewer = editor.getDocumentRenderer();
        oGraphicsPDF.SetGlobalAlpha(1);
        
        let nImgType;
        if (this.IsPressed()) {
            nImgType = AscPDF.APPEARANCE_TYPES.mouseDown;
        }
        else if (this.IsHovered()) {
            nImgType = AscPDF.APPEARANCE_TYPES.rollover;
        }
        else
            nImgType = undefined;

        let originView = AscPDF.CBaseField.prototype.GetOriginView.call(this, nImgType, oGraphicsPDF.GetDrawingPageW(), oGraphicsPDF.GetDrawingPageH());
        if (!originView) {
            this.DrawLocks(oGraphicsPDF);
            this.DrawEdit(oGraphicsWord);
            return;
        }

        let oTr             = oGraphicsPDF.GetTransform();
        let highlightType   = this.GetHighlight();

        let aOrigRect = this.GetRect();

        let origX   = aOrigRect[0];
        let origY   = aOrigRect[1];
        let X       = originView.x;
        let Y       = originView.y;

        let nWidth  = originView.width;
        let nHeight = originView.height;

        let nLineWidth = this.GetWidth() + 1;

        // Create a new canvas element for the cropped area
        var croppedCanvas       = document.createElement('canvas');
        var oCroppedCtx         = croppedCanvas.getContext("2d");
        croppedCanvas.width     = nWidth;
        croppedCanvas.height    = nHeight;

        if (this.IsPressed() == false) {
            oGraphicsPDF.DrawImageXY(originView, origX, origY, undefined, true);
            return;
        }

        if (originView) {
            switch (highlightType) {
                case AscPDF.BUTTON_HIGHLIGHT_TYPES.none:
                case AscPDF.BUTTON_HIGHLIGHT_TYPES.push:
                    oGraphicsPDF.DrawImageXY(originView, origX, origY, undefined, true);
                    break;
                case AscPDF.BUTTON_HIGHLIGHT_TYPES.invert: {
                    let xCenter = oViewer.width >> 1;
                    if (oViewer.documentWidth > oViewer.width)
                    {
                        xCenter = (oViewer.documentWidth >> 1) - (oViewer.scrollX) >> 0;
                    }
                    let yPos    = oViewer.scrollY >> 0;
                    let page    = oViewer.drawingPages[this.GetPage()];
                    let w       = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
                    let h       = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
                    let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
                    let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

                    let x = X + indLeft;
                    let y = Y + indTop;
                    let nDWidth = 0;
                    let nDHeight = 0;

                    if (x < 0) {
                        nDWidth = nWidth - (nWidth + x);
                        X       += nDWidth;
                        nWidth  += x >> 0;
                        croppedCanvas.width = nWidth;
                        x = 0;
                    }
                    if (y < 0) {
                        nDHeight    = nHeight - (nHeight + y);
                        Y           += nDHeight;
                        nHeight     += y >> 0;
                        croppedCanvas.height = nHeight;
                        y = 0;
                    }

                    oCroppedCtx.drawImage(oViewer.canvas, x, y, nWidth, nHeight, 0, 0, nWidth, nHeight);
                    
                    if (page.ImageAnnots) {
                        oCroppedCtx.drawImage(page.ImageAnnots, X, Y, nWidth, nHeight, 0, 0, nWidth, nHeight);
                    }

                    oCroppedCtx.drawImage(originView, nDWidth, nDHeight, originView.width, originView.height, 0, 0, originView.width, originView.height);
                    oCroppedCtx.globalCompositeOperation='difference';
                    oCroppedCtx.fillStyle='white';
                    oCroppedCtx.fillRect(0, 0, croppedCanvas.width,croppedCanvas.height);
                    oGraphicsPDF.DrawImageXY(oCroppedCtx.canvas, origX, origY, undefined, true);
                    break;
                }
                case AscPDF.BUTTON_HIGHLIGHT_TYPES.outline: {
                    if (originView) {
                        oCroppedCtx.drawImage(originView, 0, 0);
                    }
                    else {
                        oCroppedCtx.drawImage(oViewer.canvasForms, X, Y, nWidth, nHeight, 0, 0, nWidth, nHeight);
                    }
    
                    oCroppedCtx.clearRect(nLineWidth * oTr.sy, nLineWidth * oTr.sy, croppedCanvas.width - 2 * nLineWidth * oTr.sy, croppedCanvas.height - 2 * nLineWidth * oTr.sy);
    
                    oCroppedCtx.globalCompositeOperation='difference';
                    oCroppedCtx.fillStyle='white';
                    oCroppedCtx.fillRect(0, 0, croppedCanvas.width,croppedCanvas.height);
                    oCroppedCtx.globalCompositeOperation='source-over';
                    oCroppedCtx.drawImage(originView, nLineWidth * oTr.sy, nLineWidth * oTr.sy, nWidth - 2 * nLineWidth * oTr.sy, nHeight - 2 * nLineWidth * oTr.sy, nLineWidth * oTr.sy, nLineWidth * oTr.sy, nWidth -  2 * nLineWidth * oTr.sy, nHeight - 2 * nLineWidth * oTr.sy);
    
                    oGraphicsPDF.DrawImageXY(oCroppedCtx.canvas, origX, origY, undefined, true);
                    break;
                }
            }
        }
    };

    CAnnotationLink.prototype.RefillGeometry = function() {
        let aQuads = this.GetQuads();
        if (aQuads.length == 0 || aQuads.length > 1 || this.GetBorder() !== AscPDF.BORDER_TYPES.underline) {
            return;
        }

        AscCommon.History.StartNoHistoryMode();

        this.spPr.geometry.gdLstInfo = [];
        this.spPr.geometry.pathLst = [];
        this.spPr.geometry.AddRect("0", "txT", "w", "txB");

        this.spPr.geometry.AddPathCommand(0, undefined, "norm");
        this.spPr.geometry.AddPathCommand(1, "0", "h");
        this.spPr.geometry.AddPathCommand(2, "w", "h");
        this.spPr.geometry.preset = undefined;

        AscCommon.History.EndNoHistoryMode();
        return this.spPr.geometry;
    };
    CAnnotationLink.prototype.SetPressed = function(bValue) {
        this._pressed = bValue;
        this.AddToRedraw();
    };
    CAnnotationLink.prototype.IsPressed = function() {
        return this._pressed;
    };
    CAnnotationLink.prototype.IsHovered = function() {
        return this._hovered;
    };
    CAnnotationLink.prototype.SetHovered = function(bValue) {
        this._hovered = bValue;
    };

    CAnnotationLink.prototype.onMouseDown = function(x, y, e) {
        if (Asc.editor.canEdit()) {
            AscPDF.CPdfShape.prototype.onMouseDown.call(this, x, y, e);
            return;
        }

        this.DrawPressed();
    };
    CAnnotationLink.prototype.onMouseEnter = function() {
        if (Asc.editor.canEdit()) {
            return;
        }

        this.SetHovered(true);
    };
    CAnnotationLink.prototype.onMouseExit = function() {
        if (Asc.editor.canEdit()) {
            return;
        }

        this.SetHovered(false);
    };
    CAnnotationLink.prototype.DrawPressed = function() {
        this.SetPressed(true);
        Asc.editor.getDocumentRenderer()._paint();
    };
    CAnnotationLink.prototype.DrawUnpressed = function() {
        this.SetPressed(false);
        Asc.editor.getDocumentRenderer()._paint();
    };
    CAnnotationLink.prototype.onMouseUp = function(x, y, e) {
        if (Asc.editor.canEdit()) {
            if (e.button != 2) {
                this.GetDocument().ShowComment([this.GetId()]);
            }
            return;
        }


        if (e.Button != 2) {
            this.GetDocument().ShowComment([this.GetId()]);

            this.DrawUnpressed();
            this.AddActionsToQueue(AscPDF.PDF_TRIGGERS_TYPES.MouseUp);
        }
    };
    CAnnotationLink.prototype.SetActions = function(nTriggerType, aActionsInfo) {
        let aActions = [];
        if (aActionsInfo) {
            for (let i = 0; i < aActionsInfo.length; i++) {
                let oAction;
                switch (aActionsInfo[i]["S"]) {
                    case AscPDF.ACTIONS_TYPES.JavaScript:
                        oAction = new AscPDF.CActionRunScript(aActionsInfo[i]["JS"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.ResetForm:
                        oAction = new AscPDF.CActionReset(aActionsInfo[i]["Fields"], Boolean(aActionsInfo[i]["Flags"]));
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.URI:
                        oAction = new AscPDF.CActionURI(aActionsInfo[i]["URI"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.HideShow:
                        oAction = new AscPDF.CActionHideShow(Boolean(aActionsInfo[i]["H"]), aActionsInfo[i]["T"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.GoTo:
                        let oRect = {
                            top:    aActionsInfo[i]["top"],
                            right:  aActionsInfo[i]["right"],
                            bottom: aActionsInfo[i]["bottom"],
                            left:   aActionsInfo[i]["left"]
                        }
    
                        let oDoc = Asc.editor.getPDFDoc();
                        let oPageInfo;
                        if (aActionsInfo[i]["pageId"]) {
                            oPageInfo = AscCommon.g_oTableId.GetById(aActionsInfo[i]["pageId"]);
                        }
                        else {
                            oPageInfo = oDoc.GetPageInfo(aActionsInfo[i]["page"]);
                        }

                        oAction = new AscPDF.CActionGoTo(oPageInfo.GetId(), aActionsInfo[i]["kind"], aActionsInfo[i]["zoom"], oRect);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.Named:
                        oAction = new AscPDF.CActionNamed(AscPDF.CActionNamed.GetInternalType(aActionsInfo[i]["N"]));
                        aActions.push(oAction);
                        break;
                }
            }
        }
        
        const oNewTrigger = aActions.length != 0 ? new AscPDF.CPdfTrigger(nTriggerType, aActions) : null;
        if (oNewTrigger) {
            oNewTrigger.SetParentField(this);
        }

        const aCurActionsInfo = this.GetActions(nTriggerType);
        AscCommon.History.Add(new CChangesPDFFormActions(this, aCurActionsInfo, aActionsInfo, nTriggerType));

        switch (nTriggerType) {
            case AscPDF.PDF_TRIGGERS_TYPES.MouseUp:
                this._triggers.MouseUp = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseDown:
                this._triggers.MouseDown = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseEnter:
                this._triggers.MouseEnter = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseExit:
                this._triggers.MouseExit = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.OnFocus:
                this._triggers.OnFocus = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.OnBlur:
                this._triggers.OnBlur = oNewTrigger;
                break;
        }

        return aActions;
    };
    CAnnotationLink.prototype.GetActions = function(nTriggerType) {
        // Get the trigger by type
        let oTrigger = this.GetTrigger(nTriggerType);
        if (!oTrigger || !oTrigger.Actions) {
            return [];
        }
        
        let aActionsInfo = [];
        // Iterate through all actions associated with the trigger
        for (let i = 0; i < oTrigger.Actions.length; i++) {
            let oAction = oTrigger.Actions[i];
            let actionInfo = {};
            
            // Determine the action type and populate the object with information
            switch (oAction.GetType()) {
                case AscPDF.ACTIONS_TYPES.JavaScript:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.JavaScript;
                    actionInfo["JS"] = oAction.GetScript();
                    break;
                case AscPDF.ACTIONS_TYPES.ResetForm:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.ResetForm;
                    actionInfo["Fields"] = oAction.GetNames();
                    actionInfo["Flags"] = Number(oAction.GetNeedAllExcept());
                    break;
                case AscPDF.ACTIONS_TYPES.URI:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.URI;
                    actionInfo["URI"] = oAction.GetURI();
                    break;
                case AscPDF.ACTIONS_TYPES.HideShow:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.HideShow;
                    actionInfo["H"] = oAction.GetHidden();
                    actionInfo["T"] = oAction.GetNames();
                    break;
                case AscPDF.ACTIONS_TYPES.GoTo:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.GoTo;
                    actionInfo["page"] = oAction.GetPageIdx();
                    actionInfo["pageId"] = oAction.GetPageId();
                    actionInfo["kind"] = oAction.GetKind();
                    actionInfo["zoom"] = oAction.GetZoom();
                    let oRect = oAction.GetRect();
                    actionInfo["top"] = oRect.top;
                    actionInfo["right"] = oRect.right;
                    actionInfo["bottom"] = oRect.bottom;
                    actionInfo["left"] = oRect.left;
                    break;
                case AscPDF.ACTIONS_TYPES.Named:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.Named;
                    actionInfo["N"] = oAction.GetNameStrType();
                    break;
                default:
                    // If the type is not recognized, add handling or skip
                    break;
            }
            
            aActionsInfo.push(actionInfo);
        }
        
        return aActionsInfo;
    };
    CAnnotationLink.prototype.GetTrigger = function(nType) {
        switch (nType) {
            case AscPDF.PDF_TRIGGERS_TYPES.MouseUp:
                return this._triggers.MouseUp;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseDown:
                return this._triggers.MouseDown;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseEnter:
                return this._triggers.MouseEnter;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseExit:
                return this._triggers.MouseExit;
            case AscPDF.PDF_TRIGGERS_TYPES.OnFocus:
                return this._triggers.OnFocus;
            case AscPDF.PDF_TRIGGERS_TYPES.OnBlur:
                return this._triggers.OnBlur;
        }

        return null;
    };
    CAnnotationLink.prototype.GetListActions = function() {
        let aActions = [];

        let oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseUp);
        if (oAction) {
            aActions.push(oAction);
        }
        
        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseDown);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseEnter);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseExit);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.OnFocus);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.OnBlur);
        if (oAction) {
            aActions.push(oAction);
        }

        return aActions;
    };
    CAnnotationLink.prototype.AddActionsToQueue = function() {
        let oThis           = this;
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        Object.values(arguments).forEach(function(type) {
            let oTrigger = oThis.GetTrigger(type);
        
            if (oTrigger && oTrigger.Actions.length > 0 && false == AscCommon.History.UndoRedoInProgress) {
                oActionsQueue.AddActions(oTrigger.Actions);
            }
        })
        
        if (oActionsQueue.actions.length !== 0) {
            oActionsQueue.Start();
        }
    };

    /**
     * Defines how a button reacts when a user clicks it. The four highlight modes supported are:
     * none — No visual indication that the button has been clicked.
     * invert — The region encompassing the button’s rectangle inverts momentarily.
     * push — The down face for the button (if any) is displayed momentarily.
     * outline — The border of the rectangleinverts momentarily.
     * @memberof CPushButtonField
     * @param {number} nType - AscPDF.BUTTON_HIGHLIGHT_TYPES
     * @typeofeditors ["PDF"]
     */
    CAnnotationLink.prototype.SetHighlight = function(nType) {
        AscCommon.History.Add(new CChangesPDFLinkAnnotHighlight(this, this._highlight, nType));

        this._highlight = nType;

        this.SetWasChanged(true);
    };
    CAnnotationLink.prototype.GetHighlight = function() {
        return this._highlight;
    };

    CAnnotationLink.prototype.SetPosition = function(x, y) {
        let aCurRect = this.GetRect();

        let nOldX = aCurRect[0];
        let nOldY = aCurRect[1];

        let nDeltaX = x - nOldX;
        let nDeltaY = y - nOldY;

        if (0 == nDeltaX && 0 == nDeltaY) {
            return;
        }

        let nWidth  = aCurRect[2] - aCurRect[0];
        let nHeight = aCurRect[3] - aCurRect[1];

        let aNewRect = [x, y, x + nWidth, y + nHeight];
        let aNewQuads = [];

        for (let i = 0; i < this._quads.length; i++) {
            let aQuadsRect = [];

            for (let j = 0; j < this._quads[i].length; j+=2) {
                aQuadsRect.push(this._quads[i][j] + nDeltaX);
                aQuadsRect.push(this._quads[i][j+1] + nDeltaY);
            }

            aNewQuads.push(aQuadsRect);
        }

        this.SetRect(aNewRect);
        this.SetQuads(aNewQuads);

        this.SetNeedRecalc(true);
        this.SetWasChanged(true, false);
    };

    CAnnotationLink.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        let nFlags = 0;
        let nPosForFlags = memory.GetCurPosition();
        memory.Skip(4);

        let oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseUp);
        if (oAction) {
            nFlags |= (1 << 0);
            oAction.WriteToBinary(memory);
        }
        
        //
        // PA action
        //

        // highlight
        let nHighlightType = this.GetHighlight();
        if (nHighlightType != null) {
            memory.fieldDataFlags |= (1 << 2);
            memory.WriteByte(nHighlightType);
        }

        // quads
        let aQuads = this.GetQuads();
        if (aQuads != null) {
            memory.fieldDataFlags |= (1 << 3);
            memory.WriteByte(nHighlightType);
            
            let nLen = 0;
            for (let i = 0; i < aQuads.length; i++) {
                nLen += aQuads[i].length;
            }
            memory.WriteLong(nLen);  
            for (let i = 0; i < aQuads.length; i++) {
                for (let j = 0; j < aQuads[i].length; j++) {
                    memory.WriteDouble(aQuads[i][j]);
                }
            }
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(nPosForFlags);
        memory.WriteLong(nFlags);
        
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    
    CAnnotationLink.prototype.hitInPath = function(x, y) {
        let invert_transform = this.getInvertTransform();
        if (!invert_transform) {
            return false;
        }
        let x_t = invert_transform.TransformPointX(x, y);
        let y_t = invert_transform.TransformPointY(x, y);
        let oGeometry = this.getGeometry();
        return oGeometry.hitInInnerArea(this.getCanvasContext(), x_t, y_t);
    };
    CAnnotationLink.prototype.canRotate = function() {
        return true;
    };

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
    
	window["AscPDF"].CAnnotationLink = CAnnotationLink;
})();

