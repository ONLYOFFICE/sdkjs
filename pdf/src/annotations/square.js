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
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationSquare(sName, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Square, nPage, aRect, oDoc);
        AscFormat.CShape.call(this);
        AscPDF.initShape(this);
        this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._rectDiff      = undefined;

        // internal
        TurnOffHistory();
    }
    CAnnotationSquare.prototype.constructor = CAnnotationSquare;
    AscFormat.InitClass(CAnnotationSquare, AscFormat.CShape, AscDFH.historyitem_type_Shape);
    Object.assign(CAnnotationSquare.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationSquare.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.TurnOffHistory();

        let oSquare = new CAnnotationSquare(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oSquare.lazyCopy = true;

        oSquare._pagePos = {
            x: this._pagePos.x,
            y: this._pagePos.y,
            w: this._pagePos.w,
            h: this._pagePos.h
        }
        oSquare._origRect = this._origRect.slice();

        this.fillObject(oSquare);

        oSquare.pen = new AscFormat.CLn();
        oSquare._apIdx = this._apIdx;
        oSquare._originView = this._originView;
        oSquare.SetOriginPage(this.GetOriginPage());
        oSquare.SetAuthor(this.GetAuthor());
        oSquare.SetModDate(this.GetModDate());
        oSquare.SetCreationDate(this.GetCreationDate());
        oSquare.SetWidth(this.GetWidth());
        oSquare.SetStrokeColor(this.GetStrokeColor().slice());
        oSquare.SetFillColor(this.GetFillColor());
        oSquare.recalcInfo.recalculatePen = false;
        oSquare.recalcInfo.recalculateGeometry = true;
        this._rectDiff && oSquare.SetRectangleDiff(this._rectDiff.slice());
        oSquare.recalculate();

        return oSquare;
    };
    CAnnotationSquare.prototype.RefillGeometry = function(oGeometry, aShapeRectInMM) {
        if (this.GetBorderEffectStyle() !== AscPDF.BORDER_EFFECT_STYLES.Cloud)
            return;

        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let oDoc        = this.GetDocument();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom * g_dKoef_pix_to_mm;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom * g_dKoef_pix_to_mm;

        let aOrigRect   = this.GetOrigRect();
        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0];

        let aPoints;
        if (!oGeometry)
            oGeometry = this.spPr.geometry;
        if (!aShapeRectInMM) {
            aShapeRectInMM = [
                (aOrigRect[0] + aRD[0]) * nScaleX, (aOrigRect[1] + aRD[1]) * nScaleY,
                (aOrigRect[2] - aRD[2]) * nScaleX, (aOrigRect[3] - aRD[3]) * nScaleY
            ];

            aPoints = [
                {x: (aOrigRect[0] + aRD[0]) * nScaleX, y: (aOrigRect[1] + aRD[1]) * nScaleY},
                {x: (aOrigRect[2] - aRD[2]) * nScaleX, y: (aOrigRect[1] + aRD[1]) * nScaleY},
                {x: (aOrigRect[2] - aRD[2]) * nScaleX, y: (aOrigRect[3] - aRD[3]) * nScaleY},
                {x: (aOrigRect[0] + aRD[0]) * nScaleX, y: (aOrigRect[3] - aRD[3]) * nScaleY}
            ]
        }
        else {
            aPoints = [
                {x: aShapeRectInMM[0], y: aShapeRectInMM[1]},
                {x: aShapeRectInMM[2], y: aShapeRectInMM[1]},
                {x: aShapeRectInMM[2], y: aShapeRectInMM[3]},
                {x: aShapeRectInMM[0], y: aShapeRectInMM[3]}
            ]
        }

        oDoc.TurnOffHistory();
        AscPDF.generateCloudyGeometry(aPoints, aShapeRectInMM, oGeometry, this.GetBorderEffectIntensity());
        oGeometry.preset = undefined;
    };
    CAnnotationSquare.prototype.SetRectangleDiff = function(aDiff) {
        this._rectDiff = aDiff;

        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom * g_dKoef_pix_to_mm;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom * g_dKoef_pix_to_mm;

        let aOrigRect = this.GetOrigRect();

        this.spPr.xfrm.setOffX(aDiff[0] * nScaleX);
        this.spPr.xfrm.setOffY(aDiff[1] * nScaleY);
        let extX = ((aOrigRect[2] - aOrigRect[0]) - aDiff[0] - aDiff[2]) * nScaleX;
        let extY = ((aOrigRect[3] - aOrigRect[1]) - aDiff[1] - aDiff[3]) * nScaleY;

        this.spPr.xfrm.setExtX(extX);
        this.spPr.xfrm.setExtY(extY);
    };
    CAnnotationSquare.prototype.SetRect = function(aRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();
        let nPage       = this.GetPage();
        let aCurRect    = this.GetRect();

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

        this.SetRectangleDiff([0, 0, 0, 0]);
        oDoc.History.Add(new CChangesPDFAnnotRect(this, aCurRect, aRect));

        oDoc.TurnOffHistory();

        this.recalcGeometry();
        this.AddToRedraw();
        this.SetWasChanged(true);
    };
    CAnnotationSquare.prototype.SetRectangleDiff = function(aDiff) {
        let oDoc = this.GetDocument();
        oDoc.History.Add(new CChangesPDFAnnotRD(this, this.GetRectangleDiff(), aDiff));

        this._rectDiff  = aDiff;
        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom * g_dKoef_pix_to_mm;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom * g_dKoef_pix_to_mm;

        let aOrigRect = this.GetOrigRect();

        this.spPr.xfrm.setOffX(aDiff[0] * nScaleX);
        this.spPr.xfrm.setOffY(aDiff[1] * nScaleY);
        let extX = ((aOrigRect[2] - aOrigRect[0]) - aDiff[0] - aDiff[2]) * nScaleX;
        let extY = ((aOrigRect[3] - aOrigRect[1]) - aDiff[1] - aDiff[3]) * nScaleY;

        this.spPr.xfrm.setExtX(extX);
        this.spPr.xfrm.setExtY(extY);
    };
    CAnnotationSquare.prototype.IsSquare = function() {
        return true;
    };
    CAnnotationSquare.prototype.SetStrokeColor = function(aColor) {
        this._strokeColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
        let oLine   = this.pen;
        oLine.setFill(oFill);
    };
    CAnnotationSquare.prototype.SetFillColor = function(aColor) {
        this._fillColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
        this.setFill(oFill);
    };
    CAnnotationSquare.prototype.SetWidth = function(nWidthPt) {
        this._width = nWidthPt; 

        nWidthPt = nWidthPt > 0 ? nWidthPt : 0.5;
        let oLine = this.pen;
        oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);
    };
    CAnnotationSquare.prototype.Recalculate = function() {
        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let aOrigRect   = this.GetOrigRect();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom * g_dKoef_pix_to_mm;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom * g_dKoef_pix_to_mm;
        
        if (this.recalcInfo.recalculateGeometry)
            this.RefillGeometry();
        this.handleUpdatePosition();
        this.recalculate();
        this.updatePosition(aOrigRect[0] * nScaleX, aOrigRect[1] * nScaleY);
    };
    
    CAnnotationSquare.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // rectangle diff
        let aRD = this.GetRectangleDiff();
        if (aRD) {
            memory.annotFlags |= (1 << 15);
            for (let i = 0; i < 4; i++) {
                memory.WriteDouble(aRD[i]);
            }
        }
        
        // fill
        let aFill = this.GetFillColor();
        if (aFill != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteLong(aFill.length);
            for (let i = 0; i < aFill.length; i++)
                memory.WriteDouble(aFill[i]);
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    window["AscPDF"].CAnnotationSquare = CAnnotationSquare;
})();

