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

    let LINE_CAP_STYLES = {
        Butt:       0,
        Round:      1,
        Projecting: 2
    }

    let LINE_INTENT_TYPE = {
        Dimension:  0,
        Arrow:      1
    }

    let LINE_END_TYPE = {
        Square:         0,
        Circle:         1,
        Diamond:        2,
        OpenArrow:      3,
        ClosedArrow:    4,
        None:           5,
        Butt:           6,
        ROpenArrow:     7,
        RClosedArrow:   8,
        Slash:          9
    }

    let CAPTION_POSITIONING = {
        Inline: 0,
        Top:    1
    }
    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationLine(sName, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Line, nPage, aRect, oDoc);

        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._points        = undefined;
        this._doCaption     = undefined;
        this._intent        = undefined;
        this._lineStart     = undefined;
        this._lineEnd       = undefined;
        this._leaderLength  = undefined; // LL
        this._leaderExtend  = undefined; // LLE
        this._leaderLineOffset  = undefined; // LLO
        this._captionPos        = CAPTION_POSITIONING.Inline; // CP
        this._captionOffset     = undefined;  // CO

        // internal
        TurnOffHistory();
        this.content        = new AscPDF.CTextBoxContent(this, oDoc);
    }
    CAnnotationLine.prototype = Object.create(AscPDF.CAnnotationBase.prototype);
	CAnnotationLine.prototype.constructor = CAnnotationLine;

    CAnnotationLine.prototype.SetCaptionOffset = function(array) {
        this._captionOffset = array;
    };
    CAnnotationLine.prototype.GetCaptionOffset = function() {
        return this._captionOffset;
    };

    CAnnotationLine.prototype.SetCaptionPos = function(nPosType) {
        this._captionPos = nPosType;
    };
    CAnnotationLine.prototype.GetCaptionPos = function() {
        return this._captionPos;
    };

    CAnnotationLine.prototype.SetLeaderLineOffset = function(nValue) {
        this._leaderLineOffset = nValue;
    };
    CAnnotationLine.prototype.GetLeaderLineOffset = function() {
        return this._leaderLineOffset;
    };
    CAnnotationLine.prototype.SetLeaderLength = function(nValue) {
        this._leaderLength = nValue;
    };
    CAnnotationLine.prototype.GetLeaderLength = function() {
        return this._leaderLength;
    };
    CAnnotationLine.prototype.SetLeaderExtend = function(nValue) {
        this._leaderExtend = nValue;
    };
    CAnnotationLine.prototype.GetLeaderExtend = function() {
        return this._leaderExtend;
    };
    CAnnotationLine.prototype.SetLinePoints = function(aPoints) {
        this._points = aPoints;
    };
    CAnnotationLine.prototype.GetLinePoints = function() {
        return this._points;
    };
    CAnnotationLine.prototype.SetDoCaption = function(nType) {
        this._doCaption = nType;
    };
    CAnnotationLine.prototype.GetDoCaption = function() {
        return this._doCaption;
    };
    CAnnotationLine.prototype.SetLineStart = function(nType) {
        this._lineStart = nType;
    };
    CAnnotationLine.prototype.SetLineEnd = function(nType) {
        this._lineEnd = nType;
    };
    CAnnotationLine.prototype.GetLineStart = function() {
        return this._lineStart;
    };
    CAnnotationLine.prototype.GetLineEnd = function() {
        return this._lineEnd;
    };

    CAnnotationLine.prototype.Draw = function(oGraphics) {
        if (this.IsHidden() == true)
            return;

        let oViewer = editor.getDocumentRenderer();
        let oGraphicsWord = oViewer.pagesInfo.pages[this.GetPage()].graphics.word;
        
        this.Recalculate();

        oGraphicsWord.AddClipRect(this.contentRect.X, this.contentRect.Y, this.contentRect.W, this.contentRect.H);

        this.content.Draw(0, oGraphicsWord);
        oGraphicsWord.RemoveClip();
    };
    CAnnotationLine.prototype.Recalculate = function() {
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
    };

    CAnnotationLine.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // line points
        let aLinePoints = this.GetLinePoints();
        for (let i = 0; i < aLinePoints.length; i++) {
            memory.WriteDouble(aLinePoints[i]);
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

        // leader leader
        let nLL = this.GetLeaderLength();
        if (nLL) {
            memory.annotFlags |= (1 << 17);
            memory.WriteDouble(nLL);
        }

        // leader extend
        let nLLE = this.GetLeaderExtend();
        if (nLLE) {
            memory.annotFlags |= (1 << 18);
            memory.WriteDouble(nLLE);
        }

        // do caption
        let bDoCaption = this.GetDoCaption();
        if (bDoCaption) {
            memory.annotFlags |= (1 << 19);
        }
        
        // intent
        let nIntent = this.GetIntent();
        if (nIntent != null) {
            memory.annotFlags |= (1 << 20);
            memory.WriteDouble(nIntent);
        }

        // leader Line Offset
        let nLLO = this.GetLeaderLineOffset();
        if (nLLO != null) {
            memory.annotFlags |= (1 << 21);
            memory.WriteDouble(nLLO);
        }

        // caption positioning
        let nCP = this.GetCaptionPos();
        if (nCP != null) {
            memory.annotFlags |= (1 << 22);
            memory.WriteByte(nCP);
        }

        // caption offset
        let aCO = this.GetCaptionOffset();
        if (aCO != null) {
            memory.annotFlags |= (1 << 23);
            memory.WriteDouble(aCO[0]);
            memory.WriteDouble(aCO[1]);
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

    window["AscPDF"].CAnnotationLine = CAnnotationLine;
})();

