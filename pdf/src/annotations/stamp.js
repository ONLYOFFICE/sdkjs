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

    let STAMP_TYPES = {
        // dinamic
        D_Approved:               'D_Approved',
        D_Revised:                'D_Revised',
        D_Reviewed:               'D_Reviewed',
        D_Received:               'D_Received',
        
        // standard
        SB_Approved:              'SB_Approved',
        SB_NotApproved:           'SB_NotApproved',
        SB_Revised:               'SB_Revised',
        SB_Confidential:          'SB_Confidential',
        SB_ForComment:            'SB_ForComment',
        SB_ForPublicRelease:      'SB_ForPublicRelease',
        SB_NotForPublicRelease:   'SB_NotForPublicRelease',
        SB_PreliminaryResults:    'SB_PreliminaryResults',
        SB_InformationOnly:       'SB_InformationOnly',
        SB_Draft:                 'SB_Draft',
        SB_Completed:             'SB_Completed',
        SB_Final:                 'SB_Final',
        SB_Void:                  'SB_Void',
        
        // sign
        SH_SignHere:              'SH_SignHere',
        SH_Witness:               'SH_Witness',
        SH_InitialHere:           'SH_InitialHere',
        
        // foxit
        Expired:                  'Expired',

        // onlyoffice
        Image:                    'Image'
    }

    /**
	 * Class representing a stamp annotation.
	 * @constructor
    */
    function CAnnotationStamp(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Stamp, aRect, oDoc);
        
        this._rotate = 0;
        this._stampType = undefined;

        this.Init();
    }
    
	CAnnotationStamp.prototype.constructor = CAnnotationStamp;
    AscFormat.InitClass(CAnnotationStamp, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Stamp);
    Object.assign(CAnnotationStamp.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationStamp.prototype.IsStamp = function() {
        return true;
    };
    CAnnotationStamp.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        let sType = this.GetIconType();
        if (sType == undefined) {
            return;
        }

        this.Recalculate();
        if (AscPDF.STAMP_TYPES.Image == sType) {
            this.draw(oGraphicsWord);
        }
        else {
            this._draw(oGraphicsWord);
        }
    };
    CAnnotationStamp.prototype._draw = function(oGraphicsWord) {
        let oStructure = this.GetRenderStructure();
        if (!oStructure) {
            return;
        }

        let nScale = this.GetOriginViewScale();
        let oTr = new AscCommon.CMatrix();
        
        // draw without rotate and scale for saving
        if (oGraphicsWord.isPdf()) {
            let hc = this.extX * 0.5;
            let vc = this.extY * 0.5;
            AscCommon.global_MatrixTransformer.TranslateAppend(oTr, -hc, -vc);
            AscCommon.global_MatrixTransformer.TranslateAppend(oTr, this.x + hc, this.y + vc);

            let aInRect = this.GetInRect();
            let nSourceX = aInRect[0] * g_dKoef_pt_to_mm;
            let nSourceY = aInRect[3] * g_dKoef_pt_to_mm;
            
            oTr.tx = nSourceX;
            oTr.ty = nSourceY;
        }
        else {
            oTr.Scale(nScale, nScale);
            let oOwnTr = this.getTransformMatrix();
            AscCommon.global_MatrixTransformer.MultiplyAppend(oTr, oOwnTr);
        }

        oStructure.draw(oGraphicsWord, oTr);
    }
    CAnnotationStamp.prototype.SetRenderStructure = function(oStructure) {
        this.renderStructure = oStructure;
    };
    CAnnotationStamp.prototype.GetRenderStructure = function() {
        if (this.renderStructure) {
            return this.renderStructure;
        }
        else {
            let oDoc = this.GetDocument();
            let oTextDrawer = oDoc.CreateStampRender(this.GetIconType(), this.GetAuthor(), this.GetCreationDate());
            this.SetRenderStructure(oTextDrawer && oTextDrawer.m_aStack[0]);
            return this.renderStructure;
        }
    };
    CAnnotationStamp.prototype.SetInRect = function(aInRect) {
        AscCommon.History.Add(new CChangesPDFAnnotStampInRect(this, this.inRect, aInRect));

        this.inRect = aInRect;
        
        function getDistance(x1, y1, x2, y2) {
            return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }

        if (!aInRect) {
            return;
        }
        
        let nShapeW = getDistance(aInRect[0], aInRect[1], aInRect[6], aInRect[7]);
        let nShapeH = getDistance(aInRect[0], aInRect[1], aInRect[2], aInRect[3]);

        if (nShapeH == 0 || nShapeW == 0) {
            return;
        }

        AscCommon.History.StartNoHistoryMode();
        this.spPr.xfrm.setExtX(nShapeW * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setExtY(nShapeH * g_dKoef_pt_to_mm);

        let aRect = this.GetRect();
        
        let nAnnotW = aRect[2] - aRect[0];
        let nAnnotH = aRect[3] - aRect[1];
        
        let nOffX = (aRect[0] - (nShapeW - nAnnotW) / 2);
        let nOffY = (aRect[1] - (nShapeH - nAnnotH) / 2);
        
        this.spPr.xfrm.setOffX(nOffX * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setOffY(nOffY * g_dKoef_pt_to_mm);

        AscCommon.History.EndNoHistoryMode();
    };
    CAnnotationStamp.prototype.GetInRect = function() {
        return this.inRect;
    };
    CAnnotationStamp.prototype.GetDrawing = function() {
        return this.content.GetAllDrawingObjects()[0];
    };
    CAnnotationStamp.prototype.SetWasChanged = function(isChanged) {
        let oViewer = Asc.editor.getDocumentRenderer();

        if (this._wasChanged !== isChanged && oViewer.IsOpenAnnotsInProgress == false) {
            this._wasChanged = isChanged;
        }
    };
    CAnnotationStamp.prototype.GetOriginViewScale = function() {
        let aInRect = this.GetInRect();
        if (!aInRect) {
            return 1;
        }

        function getDistance(x1, y1, x2, y2) {
            return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }

        let nSourceH = getDistance(aInRect[0], aInRect[1], aInRect[2], aInRect[3]) * g_dKoef_pt_to_mm;
        let nCurrentH = this.getXfrmExtY();

        return nCurrentH / nSourceH;
    };
    CAnnotationStamp.prototype.DrawFromStream = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;
            
        let nScale      = this.GetOriginViewScale();
        let originView  = this.GetOriginView(oGraphicsPDF.GetDrawingPageW() * nScale, oGraphicsPDF.GetDrawingPageH() * nScale);
        let nRot        = this.GetRot();

        if (originView) {
            let oXfrm = this.getXfrm();
            
            let X = oXfrm.offX / g_dKoef_pt_to_mm;
            let Y = oXfrm.offY / g_dKoef_pt_to_mm;

            if (this.IsHighlight())
                AscPDF.startMultiplyMode(oGraphicsPDF.GetContext());
            
            oGraphicsPDF.SetGlobalAlpha(1);
            oGraphicsPDF.DrawImageXY(originView, X, Y, nRot, true);
            AscPDF.endMultiplyMode(oGraphicsPDF.GetContext());
        }
    };
    CAnnotationStamp.prototype.ClearCache = function() {
        this._originView.normal = null;
        this.APInfo = null;
    };
    CAnnotationStamp.prototype.GetOriginViewInfo = function(nPageW, nPageH) {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let nPage       = this.GetOriginPage();

        if (this.APInfo == null || this.APInfo.size.w != nPageW || this.APInfo.size.h != nPageH) {
            this.APInfo = {
                info: oFile.nativeFile["getAnnotationsAP"](nPage, nPageW, nPageH, undefined, this.GetApIdx()),
                size: {
                    w: nPageW,
                    h: nPageH
                }
            }
        }
        
        for (let i = 0; i < this.APInfo.info.length; i++) {
            if (this.APInfo.info[i]["i"] == this._apIdx)
                return this.APInfo.info[i];
        }

        return null;
    };
    CAnnotationStamp.prototype.SetPosition = function(x, y) {
        let oDoc        = this.GetDocument();
        let aCurRect    = this.GetOrigRect();

        let nOldX = aCurRect[0];
        let nOldY = aCurRect[1];

        let nDeltaX = x - nOldX;
        let nDeltaY = y - nOldY;

        if (0 == nDeltaX && 0 == nDeltaY) {
            return;
        }

        oDoc.History.Add(new CChangesPDFAnnotPos(this, [aCurRect[0], aCurRect[1]], [x, y]));

        let nWidth  = aCurRect[2] - aCurRect[0];
        let nHeight = aCurRect[3] - aCurRect[1];

        this._origRect[0] = x;
        this._origRect[1] = y;
        this._origRect[2] = x + nWidth;
        this._origRect[3] = y + nHeight;
        
        let oXfrm = this.getXfrm();
        let nOffX = x - (this.getXfrmExtX() / g_dKoef_pt_to_mm - nWidth) / 2;
        let nOffY = y - (this.getXfrmExtY() / g_dKoef_pt_to_mm - nHeight) / 2;

        AscCommon.History.StartNoHistoryMode();
        oXfrm.setOffX(nOffX * g_dKoef_pt_to_mm);
        oXfrm.setOffY(nOffY * g_dKoef_pt_to_mm);
        AscCommon.History.EndNoHistoryMode();

        this.SetNeedRecalc(true);
        this.SetWasChanged(true);
    };
    CAnnotationStamp.prototype.SetRect = function(aRect, isOnRotate) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();
        let aCurRect    = this.GetRect();
        
        if (this.GetRotate() != 0 && this.getXfrm() && !isOnRotate) {
            oDoc.History.Add(new CChangesPDFAnnotStampRect(this, aCurRect, aRect, isOnRotate));
            this._origRect = aRect;

            this.SetNeedRecalcSizes(!isOnRotate);
            this.handleUpdateRot();
            return;
        }

        oDoc.History.Add(new CChangesPDFAnnotStampRect(this, aCurRect, aRect, isOnRotate));
        this._origRect = aRect;

        this.SetWasChanged(true);
        this.SetNeedRecalcSizes(!isOnRotate);
    };
    CAnnotationStamp.prototype.SetNeedRecalcSizes = function(bRecalc) {
        let oDoc = Asc.editor.getPDFDoc();
        if (oDoc.Viewer.IsOpenAnnotsInProgress) {
            return;
        }

        this._needRecalcSizes = bRecalc;
        this.recalcGeometry();
    };
    CAnnotationStamp.prototype.IsNeedRecalcSizes = function() {
        return this._needRecalcSizes;
    };
    CAnnotationStamp.prototype.canRotate = function() {
        return true;
    };
    CAnnotationStamp.prototype.recalculateBounds = function() {
        let sType = this.GetIconType();
        if (sType == undefined) {
            return;
        }

        let boundsChecker = new AscFormat.CSlideBoundsChecker();
        
        boundsChecker.DO_NOT_DRAW_ANIM_LABEL = true;
        if (AscPDF.STAMP_TYPES.Image == sType || !this.GetRenderStructure()) {
            this.draw(boundsChecker);
        } else {
            this._draw(boundsChecker);
        }
        boundsChecker.CorrectBounds();

        this.bounds.x = boundsChecker.Bounds.min_x;
        this.bounds.y = boundsChecker.Bounds.min_y;
        this.bounds.l = boundsChecker.Bounds.min_x;
        this.bounds.t = boundsChecker.Bounds.min_y;
        this.bounds.r = boundsChecker.Bounds.max_x;
        this.bounds.b = boundsChecker.Bounds.max_y;
        this.bounds.w = boundsChecker.Bounds.max_x - boundsChecker.Bounds.min_x;
        this.bounds.h = boundsChecker.Bounds.max_y - boundsChecker.Bounds.min_y;
    };
    CAnnotationStamp.prototype.Recalculate = function(bForce) {
        if (true !== bForce && false == this.IsNeedRecalc()) {
            return;
        }

        if (this.IsNeedRecalcSizes()) {
            let aRect = this.GetRect();

            let extX = ((aRect[2] - aRect[0])) * g_dKoef_pt_to_mm;
            let extY = ((aRect[3] - aRect[1])) * g_dKoef_pt_to_mm;

            this.spPr.xfrm.setOffX(aRect[0] * g_dKoef_pt_to_mm);
            this.spPr.xfrm.setOffY(aRect[1] * g_dKoef_pt_to_mm);

            this.spPr.xfrm.setExtX(extX);
            this.spPr.xfrm.setExtY(extY);

            this.SetNeedRecalcSizes(false);
        }

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CAnnotationStamp.prototype.RefillGeometry = function(oGeometry, aBounds) {
        oGeometry.Recalculate(aBounds[2] - aBounds[0], aBounds[3] - aBounds[1]);
        return oGeometry;
    };
    CAnnotationStamp.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oNewStamp = new CAnnotationStamp(AscCommon.CreateGUID(), this.GetOrigRect().slice(), oDoc);

        oNewStamp.SetInRect(this.inRect);
        oNewStamp.lazyCopy = true;

        this.fillObject(oNewStamp);

        let aStrokeColor    = this.GetStrokeColor();
        let aFillColor      = this.GetFillColor();

        oNewStamp._apIdx = this._apIdx;
        oNewStamp._originView = this._originView;
        oNewStamp.SetOriginPage(this.GetOriginPage());
        oNewStamp.SetAuthor(this.GetAuthor());
        oNewStamp.SetModDate(this.GetModDate());
        oNewStamp.SetCreationDate(this.GetCreationDate());
        aStrokeColor && oNewStamp.SetStrokeColor(aStrokeColor.slice());
        aFillColor && oNewStamp.SetFillColor(aFillColor.slice());
        oNewStamp.SetWidth(this.GetWidth());
        oNewStamp.SetOpacity(this.GetOpacity());
        oNewStamp.recalcGeometry()
        oNewStamp.SetNeedRecalcSizes(false);
        oNewStamp.Recalculate(true);
        oNewStamp.SetIconType(this.GetIconType());
        oNewStamp.SetRenderStructure(this.GetRenderStructure());

        oDoc.EndNoHistoryMode();
        return oNewStamp;
    };
    
    CAnnotationStamp.prototype.IsSelected = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        return oDrawingObjects.selectedObjects.includes(this);
    };
    
    CAnnotationStamp.prototype.SetIconType = function(sType) {
        if (typeof(sType) == "string") {
            let aSplitted = sType.split('#');
            sType = aSplitted[aSplitted.length - 1];
        }

        if (sType == this._stampType) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotStampType(this, this._stampType, sType));
        this._stampType = sType;
        this.SetWasChanged(true);
    };
    CAnnotationStamp.prototype.GetIconType = function() {
        return this._stampType;
    };
    CAnnotationStamp.prototype.SetRotate = function(nAngle) {
        let oViewer = Asc.editor.getDocumentRenderer();
        if (this._rotate == nAngle) {
            return;
        }
        
        AscCommon.History.Add(new CChangesPDFAnnotRotate(this, this._rotate, nAngle));

		let oXfrm = this.getXfrm();
        oXfrm.rot = -nAngle * (Math.PI / 180);
		
        this._rotate = nAngle;

        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
    };
    CAnnotationStamp.prototype.GetRotate = function() {
        return this._rotate;
    };
    CAnnotationStamp.prototype.handleUpdateRot = function(){
        AscFormat.CShape.prototype.handleUpdateRot.call(this);
        let oXfrm = this.getXfrm();

        this.SetRotate(-oXfrm.rot * (180 / Math.PI));
        this.recalcBounds();
        this.recalcGeometry();
        this.Recalculate(true);
        
        let aNewRect = [];
        let oGrBounds = this.bounds;
        aNewRect[0] = Math.round(oGrBounds.l) * g_dKoef_mm_to_pt;
        aNewRect[1] = Math.round(oGrBounds.t) * g_dKoef_mm_to_pt;
        aNewRect[2] = Math.round(oGrBounds.r) * g_dKoef_mm_to_pt;
        aNewRect[3] = Math.round(oGrBounds.b) * g_dKoef_mm_to_pt;

        this.SetRect(aNewRect, true);
    };
    CAnnotationStamp.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        memory.WriteString(this.GetIconType());
        memory.WriteDouble(this.GetRotate());
        
        let aInRect = this.GetInRect();
        let nBorderW = this.GetWidth();
        // original rect
        memory.WriteDouble(aInRect[0] - nBorderW / 2); // x1
        memory.WriteDouble(aInRect[3] - nBorderW / 2); // y1
        memory.WriteDouble(aInRect[4] + nBorderW / 2); // x2
        memory.WriteDouble(aInRect[1] + nBorderW / 2); // y2

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    CAnnotationStamp.prototype.SetDrawFromStream = function() {
        let oViewer = editor.getDocumentRenderer();
        if (oViewer.IsOpenAnnotsInProgress) {
            this._bDrawFromStream = true;
        }
    };
    CAnnotationStamp.prototype.GetAllFonts = function(fontMap) {
        if (this.IsNeedDrawFromStream() || this.fontsChecked) {
            return fontMap;
        }

        let oContent = this.getDocContent();
        if (!oContent) {
            return fontMap;
        }
        
        oContent.SetApplyToAll(true);
        AscFonts.FontPickerByCharacter.getFontsByString(oContent.GetSelectedText());
        oContent.SetApplyToAll(false);
        this.fontsChecked = true;
        this.renderStructure = null;
        this.AddToRedraw()

        return fontMap;
    };
    CAnnotationStamp.prototype.WriteRenderToBinary = function(memory) {
        // пока только для основанных на фигурах
        if (this.IsNeedDrawFromStream()) {
            return;
        }

        // тут будет длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.Draw(undefined, memory.docRenderer); // для каждой страницы инициализируется свой renderer

        // запись длины комманд
        let nEndPos = memory.GetCurPosition();
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    
    CAnnotationStamp.prototype.getNoChangeAspect = function() {
        return true;
    };
    CAnnotationStamp.prototype.Init = function() {
        let aOrigRect = this.GetRect();
        let aAnnotRectMM = aOrigRect ? aOrigRect.map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        }) : [];

        let nOffX = aAnnotRectMM[0];
        let nOffY = aAnnotRectMM[1];
        let nExtX = aAnnotRectMM[2] - aAnnotRectMM[0];
        let nExtY = aAnnotRectMM[3] - aAnnotRectMM[1];

        this.setSpPr(new AscFormat.CSpPr());
        this.spPr.setLn(new AscFormat.CLn());
        this.spPr.ln.setFill(AscFormat.CreateNoFillUniFill());
        this.spPr.setFill(AscFormat.CreateSolidFillRGBA(255, 255, 255, 255));
        this.spPr.setParent(this);
        this.spPr.setXfrm(new AscFormat.CXfrm());
        this.spPr.xfrm.setParent(this.spPr);
        
        this.spPr.xfrm.setOffX(nOffX);
        this.spPr.xfrm.setOffY(nOffY);
        this.spPr.xfrm.setExtX(nExtX);
        this.spPr.xfrm.setExtY(nExtY);
        
        this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

        this.setStyle(AscFormat.CreateDefaultShapeStyle());
        this.setBDeleted(false);
        this.recalculate();
    };
    
    window["AscPDF"].CAnnotationStamp = CAnnotationStamp;
    window["AscPDF"].STAMP_TYPES = window["AscPDF"]["STAMP_TYPES"] = STAMP_TYPES;
})();

