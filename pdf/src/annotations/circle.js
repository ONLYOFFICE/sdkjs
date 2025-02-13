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
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationCircle(sName, nPage, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Circle, nPage, aRect, oDoc);
        
        AscPDF.initShape(this);
        this.spPr.setGeometry(AscFormat.CreateGeometry("ellipse"));
        this.setStyle(AscFormat.CreateDefaultShapeStyle("ellipse"));

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._rectDiff      = [0, 0, 0, 0];
    }
	CAnnotationCircle.prototype.constructor = CAnnotationCircle;
    AscFormat.InitClass(CAnnotationCircle, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Circle);
    Object.assign(CAnnotationCircle.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationCircle.prototype.IsCircle = function() {
        return true;
    };
    CAnnotationCircle.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oCircle = new CAnnotationCircle(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);
        oCircle.lazyCopy = true;

        this.fillObject(oCircle);

        let aStrokeColor    = this.GetStrokeColor();
        let aFillColor      = this.GetFillColor();
        let aRD             = this.GetRectangleDiff();

        oCircle._apIdx = this._apIdx;
        oCircle._originView = this._originView;
        oCircle.SetOriginPage(this.GetOriginPage());
        oCircle.SetAuthor(this.GetAuthor());
        oCircle.SetModDate(this.GetModDate());
        oCircle.SetCreationDate(this.GetCreationDate());
        aStrokeColor && oCircle.SetStrokeColor(aStrokeColor.slice());
        aFillColor && oCircle.SetFillColor(aFillColor.slice());
        oCircle.SetWidth(this.GetWidth());
        oCircle.SetOpacity(this.GetOpacity());
        oCircle.recalcGeometry()
        aRD && oCircle.SetRectangleDiff(aRD.slice(), true);
        oCircle.SetDash(this.GetDash());
        oCircle.Recalculate(true);

        oDoc.EndNoHistoryMode();
        return oCircle;
    };
    CAnnotationCircle.prototype.RefillGeometry = function(oGeometry, aShapeRectInMM) {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();

        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0];
        let aOrigRect   = this.GetOrigRect();

        if (!oGeometry) {
            oGeometry = this.spPr.geometry;
        }
        if (!aShapeRectInMM) {
            aShapeRectInMM = [
                (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm, (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm,
                (aOrigRect[2] - aRD[2]) * g_dKoef_pt_to_mm, (aOrigRect[3] - aRD[3]) * g_dKoef_pt_to_mm
            ];
        }
        
        oDoc.StartNoHistoryMode();
        if (this.GetBorderEffectStyle() === AscPDF.BORDER_EFFECT_STYLES.Cloud) {
            generateCloudyGeometry(undefined, aShapeRectInMM, oGeometry, this.GetBorderEffectIntensity());
        }
        else {
            oGeometry.Recalculate(aShapeRectInMM[2] - aShapeRectInMM[0], aShapeRectInMM[3] - aShapeRectInMM[1]);
        }
        oDoc.EndNoHistoryMode();
    };
    CAnnotationCircle.prototype.SetRect = function(aOrigRect) {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();

        oDoc.History.Add(new CChangesPDFAnnotRect(this, this._origRect, aOrigRect));

        this._origRect = aOrigRect;
        
        this.SetWasChanged(true);
        this.SetNeedRecalcSizes(true);
    };
    CAnnotationCircle.prototype.SetRectangleDiff = function(aDiff, bOnResize) {
        let oDoc = this.GetDocument();
        oDoc.History.Add(new CChangesPDFAnnotRD(this, this.GetRectangleDiff(), aDiff));

        this._rectDiff = aDiff;

        this.SetWasChanged(true);
        this.SetNeedRecalcSizes(true);
        this.SetNeedRecalc(true);
    };
    CAnnotationCircle.prototype.SetNeedRecalcSizes = function(bRecalc) {
        this._needRecalcSizes = bRecalc;
        this.recalcGeometry();
    };
    CAnnotationCircle.prototype.IsNeedRecalcSizes = function() {
        return this._needRecalcSizes;
    };
    CAnnotationCircle.prototype.Recalculate = function(bForce) {
        if (true !== bForce && false == this.IsNeedRecalc()) {
            return;
        }

        if (this.IsNeedRecalcSizes()) {
            let aOrigRect = this.GetOrigRect();
            let aRD = this.GetRectangleDiff();

            let extX = ((aOrigRect[2] - aOrigRect[0]) - aRD[0] - aRD[2]) * g_dKoef_pt_to_mm;
            let extY = ((aOrigRect[3] - aOrigRect[1]) - aRD[1] - aRD[3]) * g_dKoef_pt_to_mm;

            this.spPr.xfrm.offX = (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm;
            this.spPr.xfrm.offY = (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm;

            this.spPr.xfrm.extX = extX;
            this.spPr.xfrm.extY = extY;

            this.SetNeedRecalcSizes(false);
        }
        if (this.recalcInfo.recalculateGeometry) {
            this.RefillGeometry();
        }

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CAnnotationCircle.prototype.WriteToBinary = function(memory) {
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
    
    function generateCloudyGeometry(arrPoints, aBounds, oGeometry, nIntensity) {
        let xMin = aBounds[0];
        let yMin = aBounds[1];
        let xMax = aBounds[2];
        let yMax = aBounds[3];

        let geometry = oGeometry ? oGeometry : new AscFormat.Geometry();
        if (oGeometry) {
            oGeometry.pathLst.length = 0;
        }

        let w = xMax - xMin, h = yMax-yMin;
        geometry.AddPathCommand(0, undefined, undefined, undefined, undefined, undefined);

        let dR  = nIntensity * 1.8;
        let dR2 = 1.5 *dR;
        
        let dXCE    = w / 2;
        let dYCE    = h / 2;
        let dA      = w / 2;
        let dB      = h / 2;
        let dAlpha  = 0;

        let aPoints = [];
        let p;
        
        function findPointsOnLine(pt1, pt2, N) {
            let points = [];
            let dx = pt2.x - pt1.x;
            let dy = pt2.y - pt1.y;
        
            for (let i = 0; i < N; i++) {
                let t = i / (N - 1);
                let x = pt1.x + t * dx;
                let y = pt1.y + t * dy;
                points.push({x: x, y: y});
            }
        
            return points;
        }

        if (arrPoints) {
            let oPt1;
            let oPt2;
            for (let i = 0; i < arrPoints.length; i++) {
                oPt1 = arrPoints[i];
                oPt2 = arrPoints[i + 1] || arrPoints[0];

                aPoints.splice(aPoints.length - 1, 1);
                if (oPt1.x == oPt2.x && oPt1.y == oPt2.y)
                    break;

                let nLineLenght = AscFormat.getLineLength(oPt1, oPt2);
                let nPointsCount = Math.ceil(nLineLenght / (dR2)) + 1;
                aPoints = aPoints.concat(findPointsOnLine(oPt1, oPt2, nPointsCount));
            }

            if (Math.abs(aPoints[0].x - aPoints[aPoints.length - 1].x) < 0.001 && Math.abs(aPoints[0].y - aPoints[aPoints.length - 1].y) < 0.001)
                aPoints.length = aPoints.length - 1;
        }
        else {
            while (dAlpha < 2* Math.PI) {
                p = AscFormat.getEllipsePoint(dXCE, dYCE, dA, dB, dAlpha);
                aPoints.push(p);
                dAlpha = AscFormat.ellipseCircleIntersection(dXCE, dYCE, dA, dB, dAlpha, dR2).alpha;
            }
        }

        let getLocationOnLine = function(p, p0, p1) {
            let x, y, x0, y0, x1, y1;
            x = p.x;
            y = p.y;
            x0 = p0.x;
            y0 = p0.y;
            x1 = p1.x;
            y1 = p1.y;
            let dDet = ((y0 - y1)*(x-x0) + (x1-x0)*(y-y0));
            if(dDet > 0) {
                return 1;
            }
            if(dDet < 0) {
                return -1;
            }
            return 0;
        }

        for(let nPt = 0; nPt < aPoints.length; ++nPt) {
            let oPrevPt, oCurPt, oNextPt;
            if(nPt > 0) {
                oPrevPt = aPoints[nPt - 1];
            }
            else {
                oPrevPt = aPoints[aPoints.length - 1];
            }
            oCurPt = aPoints[nPt];

            if(nPt < aPoints.length - 1) {
                oNextPt = aPoints[nPt + 1];
            }
            else {
                oNextPt = aPoints[0];
            }

            let dStAng, dEndAng, dSwAng;
            const dSwAdd = Math.PI / 8;
            let aInters1 = AscFormat.circlesIntersection(oPrevPt.x, oPrevPt.y, dR, oCurPt.x, oCurPt.y, dR);

            let oStartP;
            for(let nIdx = 0; nIdx < aInters1.length; ++nIdx) {
                let oTestP = aInters1[nIdx];
                if(getLocationOnLine(oTestP, oPrevPt, oCurPt) < 0) {
                    oStartP = oTestP;

                    break;
                }
            }
            if(!oStartP) {
                return;
            }

            let aInters2 = AscFormat.circlesIntersection(oCurPt.x, oCurPt.y, dR, oNextPt.x, oNextPt.y, dR);
            let oEndP;
            for(let nIdx = 0; nIdx < aInters2.length; ++nIdx) {
                let oTestP = aInters2[nIdx];
                if(getLocationOnLine(oTestP, oCurPt, oNextPt) < 0) {
                    oEndP = oTestP;

                    break;
                }
            }
            if(!oEndP) {
                return;
            }

            function c(dCoord) {
                return (dCoord * 36000 + 0.5 >> 0) + "";
            }

            function n(dAngRad) {
                let dResult = dAngRad;
                while (dResult >= 2*Math.PI) {
                    dResult -= 2*Math.PI;
                }
                while (dResult < 0) {
                    dResult += 2*Math.PI;
                }
                return dResult;
            }

            function a(dAngRad) {
                return (AscFormat.cToDeg * n(dAngRad) + 0.5 >> 0) + "";
            }

            //vector
            let dDX = 1;
            let dDY = 0;


            let dDXStart = (oStartP.x - oCurPt.x);
            let dDYStart = (oStartP.y - oCurPt.y);
            let dDXEnd = (oEndP.x - oCurPt.x);
            let dDYEnd = (oEndP.y - oCurPt.y);


            let dot = dDX*dDXStart + dDY*dDYStart;
            let det = dDX*dDYStart - dDY*dDXStart;
            dStAng = n(Math.atan2(det, dot));

            dot = dDX*dDXEnd + dDY*dDYEnd;
            det = dDX*dDYEnd - dDY*dDXEnd;
            dEndAng = n(Math.atan2(det, dot) + dSwAdd);



            while(dEndAng < dStAng) {
                dEndAng += Math.PI * 2;
            }

            dSwAng = dEndAng - dStAng;

            let dStartX = oStartP.x;
            let dStartY =  oStartP.y;


            if(nPt === 0) {
                if (arrPoints)
                    geometry.AddPathCommand(1, c(dStartX - xMin), c(dStartY - yMin));
                else
                    geometry.AddPathCommand(1, c(dStartX), c(dStartY));
            }

            geometry.AddPathCommand(3, c(dR), c(dR),
                a(dStAng), a(dSwAng));
            geometry.AddPathCommand(3, c(dR), c(dR), a(dStAng + dSwAng), "-"  + a(dSwAdd));
        }

        return geometry;
    }

    function calculateAngle(x1, y1, x2, y2) {
        let dy = y2 - y1;
        let dx = x2 - x1;
        let theta = Math.atan2(dy, dx); // диапазон [-PI, PI]
        theta *= 180 / Math.PI; // радианы в градусы
        // если нужен угол в диапазоне [0, 360)
        if (theta < 0) theta = 360 + theta;
        return theta;
    }

    window["AscPDF"].CAnnotationCircle      = CAnnotationCircle;
    window["AscPDF"].generateCloudyGeometry = generateCloudyGeometry;
})();

