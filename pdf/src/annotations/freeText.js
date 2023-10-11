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

    let FREE_TEXT_INTENT_TYPE = {
        FreeText:           0,
        FreeTextCallout:    1,
        FreeTextTypeWriter: 2
    }

    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationFreeText(sName, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.FreeText, nPage, aRect, oDoc);

        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._points        = undefined;
        this._intent        = undefined;
        this._lineEnd       = undefined;
        this._callout       = undefined;
        this._alignment     = undefined;
        this._defaultStyle  = undefined;

        // internal
        TurnOffHistory();
        this.content        = new AscPDF.CTextBoxContent(this, oDoc);
    }
    CAnnotationFreeText.prototype = Object.create(AscPDF.CAnnotationBase.prototype);
	CAnnotationFreeText.prototype.constructor = CAnnotationFreeText;

    CAnnotationFreeText.prototype.SetDefaultStyle = function(sStyle) {
        this._defaultStyle = sStyle;
    };
    CAnnotationFreeText.prototype.GetDefaultStyle = function() {
        return this._defaultStyle;
    };
    CAnnotationFreeText.prototype.SetAlign = function(nType) {
        this._alignment = nType;
    }
    CAnnotationFreeText.prototype.GetAlign = function() {
        return this._alignment;
    }
    CAnnotationFreeText.prototype.SetLineEnd = function(nType) {
        this._lineEnd = nType;
    };
    CAnnotationFreeText.prototype.GetLineEnd = function() {
        return this._lineEnd;
    };
    CAnnotationFreeText.prototype.SetCallout = function(aCallout) {
        this._callout = aCallout;
    };
    CAnnotationFreeText.prototype.GetCallout = function() {
        return this._callout;
    };
    CAnnotationFreeText.prototype.Draw = function(oGraphics) {
        if (this.IsHidden() == true)
            return;

        let oViewer = editor.getDocumentRenderer();
        let oGraphicsWord = oViewer.pagesInfo.pages[this.GetPage()].graphics.word;
        
        this.Recalculate();

        oGraphicsWord.AddClipRect(this.contentRect.X, this.contentRect.Y, this.contentRect.W, this.contentRect.H);

        this.content.Draw(0, oGraphicsWord);
        oGraphicsWord.RemoveClip();
    };
    CAnnotationFreeText.prototype.Recalculate = function() {
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
    CAnnotationFreeText.prototype.AddReply = function(oReply) {
        let oDoc = this.GetDocument();
        oDoc.CreateNewHistoryPoint();

        oReply.SetReplyTo(this);

        let aNewReplies = [].concat(this._replies);
        aNewReplies.push(oReply);
        this.SetReplies(aNewReplies);

        oDoc.TurnOffHistory();
    };
    CAnnotationFreeText.prototype.RemoveComment = function() {
        let oDoc = this.GetDocument();

        oDoc.CreateNewHistoryPoint();
        this.SetReplies([]);
        oDoc.TurnOffHistory();
    };
    CAnnotationFreeText.prototype.SetContents = function(contents) {
        if (this.GetContents() == contents)
            return;

        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oCurContents    = this.GetContents();

        this._contents  = contents;
        
        if (oDoc.History.UndoRedoInProgress == false && oViewer.IsOpenAnnotsInProgress == false) {
            oDoc.History.Add(new CChangesPDFAnnotContents(this, oCurContents, contents));
        }
        
        this.SetWasChanged(true);
    };
    CAnnotationFreeText.prototype.SetReplies = function(aReplies) {
        let oDoc = this.GetDocument();
        let oViewer = editor.getDocumentRenderer();

        if (oDoc.History.UndoRedoInProgress == false && oViewer.IsOpenAnnotsInProgress == false) {
            oDoc.History.Add(new CChangesPDFAnnotReplies(this, this._replies, aReplies));
        }
        this._replies = aReplies;
        if (aReplies.length != 0)
            this._OnAfterSetReply();
        else
            editor.sync_RemoveComment(this.GetId());

        this.SetWasChanged(true);   
    };
    CAnnotationFreeText.prototype.GetAscCommentData = function() {
        let oAscCommData = new Asc["asc_CCommentDataWord"](null);
        if (this._replies.length == 0)
            return oAscCommData;

        let oMainComm = this._replies[0];
        oAscCommData.asc_putText(oMainComm.GetContents());
        oAscCommData.asc_putOnlyOfficeTime(oMainComm.GetModDate().toString());
        oAscCommData.asc_putUserId(editor.documentUserId);
        oAscCommData.asc_putUserName(oMainComm.GetAuthor());
        oAscCommData.asc_putSolved(false);
        oAscCommData.asc_putQuoteText("");
        oAscCommData.m_sUserData = oMainComm.GetApIdx();

        this._replies.forEach(function(reply, index) {
            if (index == 0)
                return;
            
            oAscCommData.m_aReplies.push(reply.GetAscCommentData());
        });

        return oAscCommData;
    };
    CAnnotationFreeText.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // alignment
        let nAlign = this.GetAlign();
        if (nAlign != null)
            memory.WriteByte(nAlign);

        // rectangle diff
        let aRD = this.GetReqtangleDiff();
        if (aRD) {
            memory.annotFlags |= (1 << 15);
            for (let i = 0; i < 4; i++) {
                memory.WriteDouble(aRD[i]);
            }
        }

        // callout
        let aCallout = this.GetCallout();
        if (aCallout != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteLong(aCallout.length);
            for (let i = 0; i < aCallout.length; i++)
                memory.WriteDouble(aCallout[i]);
        }

        // default style
        let sDefaultStyle = this.GetDefaultStyle();
        if (sDefaultStyle != null) {
            memory.annotFlags |= (1 << 17);
            memory.WriteString(sDefaultStyle);
        }

        // line end
        let nLE = this.GetLineEnd();
        if (nLE != null) {
            memory.annotFlags |= (1 << 18);
            memory.WriteByte(nLE);
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

    window["AscPDF"].CAnnotationFreeText = CAnnotationFreeText;
})();

