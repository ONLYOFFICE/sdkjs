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

    let POLYLINE_INTENT_TYPE = {
        PolygonCloud:       0,
        PolyLineDimension:  1,
        PolygonDimension:   2
    }

    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationPolyLine(sName, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.PolyLine, nPage, aRect, oDoc);

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._lineStart     = undefined;
        this._lineEnd       = undefined;
        this._vertices      = undefined;

        // internal
        TurnOffHistory();
        this.content        = new AscPDF.CTextBoxContent(this, oDoc);
    }
    CAnnotationPolyLine.prototype = Object.create(AscPDF.CAnnotationBase.prototype);
	CAnnotationPolyLine.prototype.constructor = CAnnotationPolyLine;

    CAnnotationPolyLine.prototype.getObjectType = function() {
        return -1;
    };

    CAnnotationPolyLine.prototype.Draw = function(oGraphics) {
        if (this.IsHidden() == true)
            return;

        let oViewer = editor.getDocumentRenderer();
        let oGraphicsWord = oViewer.pagesInfo.pages[this.GetPage()].graphics.word;
        
        this.Recalculate();
        //this.DrawBackground();

        oGraphicsWord.AddClipRect(this.contentRect.X, this.contentRect.Y, this.contentRect.W, this.contentRect.H);

        this.content.Draw(0, oGraphicsWord);
        oGraphicsWord.RemoveClip();
    };
    CAnnotationPolyLine.prototype.Recalculate = function() {
        // if (this.IsNeedRecalc() == false)
        //     return;

        let oViewer = editor.getDocumentRenderer();
        let aRect   = this.GetRect();
        
        let X = aRect[0];
        let Y = aRect[1];
        let nWidth = (aRect[2] - aRect[0]);
        let nHeight = (aRect[3] - aRect[1]);

        let contentX;
        let contentY;
        let contentXLimit;
        let contentYLimit;
        
        contentX = (X) * g_dKoef_pix_to_mm;
        contentY = (Y) * g_dKoef_pix_to_mm;
        contentXLimit = (X + nWidth) * g_dKoef_pix_to_mm;
        contentYLimit = (Y + nHeight) * g_dKoef_pix_to_mm;

        // this._formRect.X = X * g_dKoef_pix_to_mm;
        // this._formRect.Y = Y * g_dKoef_pix_to_mm;
        // this._formRect.W = nWidth * g_dKoef_pix_to_mm;
        // this._formRect.H = nHeight * g_dKoef_pix_to_mm;
        
        if (!this.contentRect)
            this.contentRect = {};

        this.contentRect.X = contentX;
        this.contentRect.Y = contentY;
        this.contentRect.W = contentXLimit - contentX;
        this.contentRect.H = contentYLimit - contentY;

        if (!this._oldContentPos)
            this._oldContentPos = {};

        if (contentX != this._oldContentPos.X || contentY != this._oldContentPos.Y ||
            contentXLimit != this._oldContentPos.XLimit) {
            this.content.X      = this._oldContentPos.X        = contentX;
            this.content.Y      = this._oldContentPos.Y        = contentY;
            this.content.XLimit = this._oldContentPos.XLimit   = contentXLimit;
            this.content.YLimit = this._oldContentPos.YLimit   = 20000;
            this.content.Recalculate_Page(0, true);
        }
        // else if (this.IsNeedRecalc()) {
        //     this.content.Recalculate_Page(0, false);
        // }
    };
    
    CAnnotationPolyLine.prototype.SetVertices = function(aVertices) {
        this._vertices = aVertices;
    };
    CAnnotationPolyLine.prototype.GetVertices = function() {
        return this._vertices;
    };
    CAnnotationPolyLine.prototype.SetLineStart = function(nType) {
        this._lineStart = nType;
    };
    CAnnotationPolyLine.prototype.GetLineStart = function() {
        return this._lineStart;
    };
    CAnnotationPolyLine.prototype.SetLineEnd = function(nType) {
        this._lineEnd = nType;
    };
    CAnnotationPolyLine.prototype.GetLineEnd = function() {
        return this._lineEnd;
    };
    CAnnotationPolyLine.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // vertices
        let aVertices = this.GetVertices();
        if (aVertices) {
            memory.WriteLong(aVertices.length);
            for (let i = 0; i < aVertices.length; i++) {
                memory.WriteDouble(aVertices[i]);
            }
        }
        
        // line ending
        let nLS = this.GetLineStart();
        let nLE = this.GetLineEnd();
        if (nLE != null && nLS != null) {
            memory.annotFlags |= (1 << 15);
            memory.WriteByte(nLS);
            memory.WriteByte(nLE);
        }

        // fill
        let aFill = this.GetFillColor();
        if (aFill != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteLong(aFill.length);
            for (let i = 0; i < aFill.length; i++)
                memory.WriteDouble(aFill[i]);
        }

        // intent
        let nIntent = this.GetIntent();
        if (nIntent != null) {
            memory.annotFlags |= (1 << 20);
            memory.WriteDouble(nIntent);
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);

        this._replies.forEach(function(reply) {
            reply.WriteToBinary(memory); 
        });
    };
    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    window["AscPDF"].CAnnotationPolyLine = CAnnotationPolyLine;
})();

