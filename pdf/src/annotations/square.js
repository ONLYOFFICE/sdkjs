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
    function CAnnotationSquare(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Square, aRect, oDoc);
        
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
        this._rectDiff      = [0, 0, 0, 0];
    }
    CAnnotationSquare.prototype.constructor = CAnnotationSquare;
    AscFormat.InitClass(CAnnotationSquare, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Square);
    Object.assign(CAnnotationSquare.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationSquare.prototype.RefillGeometry = function(oGeometry, aShapeRectInMM) {
        if (this.GetBorderEffectStyle() !== AscPDF.BORDER_EFFECT_STYLES.Cloud)
            return;

        let aOrigRect   = this.GetRect();
        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0];

        let aPoints;
        if (!oGeometry)
            oGeometry = this.spPr.geometry;
        if (!aShapeRectInMM) {
            aShapeRectInMM = [
                (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm, (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm,
                (aOrigRect[2] - aRD[2]) * g_dKoef_pt_to_mm, (aOrigRect[3] - aRD[3]) * g_dKoef_pt_to_mm
            ];

            aPoints = [
                {x: (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm, y: (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm},
                {x: (aOrigRect[2] - aRD[2]) * g_dKoef_pt_to_mm, y: (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm},
                {x: (aOrigRect[2] - aRD[2]) * g_dKoef_pt_to_mm, y: (aOrigRect[3] - aRD[3]) * g_dKoef_pt_to_mm},
                {x: (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm, y: (aOrigRect[3] - aRD[3]) * g_dKoef_pt_to_mm}
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

        AscCommon.History.StartNoHistoryMode();
        AscPDF.generateCloudyGeometry(aPoints, aShapeRectInMM, oGeometry, this.GetBorderEffectIntensity());
        AscCommon.History.EndNoHistoryMode();

        oGeometry.preset = undefined;
    };
    CAnnotationSquare.prototype.IsSquare = function() {
        return true;
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

    window["AscPDF"].CAnnotationSquare = CAnnotationSquare;
})();

