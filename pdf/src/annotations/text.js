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

    let TEXT_ANNOT_STATE = {
        Marked:     0,
        Unmarked:   1,
        Accepted:   2,
        Rejected:   3,
        Cancelled:  4,
        Completed:  5,
        None:       6,
        Unknown:    7
    }

    let TEXT_ANNOT_STATE_MODEL = {
        Marked:     0,
        Review:     1,
        Unknown:    2
    }

    let NOTE_ICONS_TYPES = {
        Check1:         0,
        Check2:         1,
        Circle:         2,
        Comment:        3,
        Cross:          4,
        CrossH:         5,
        Help:           6,
        Insert:         7,
        Key:            8,
        NewParagraph:   9,
        Note:           10,
        Paragraph:      11,
        RightArrow:     12,
        RightPointer:   13,
        Star:           14,
        UpArrow:        15,
        UpLeftArrow:    16
    }

    /**
	 * Class representing a text annotation.
	 * @constructor
    */
    function CAnnotationText(sName, nPage, aOrigRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Text, nPage, aOrigRect, oDoc);

        this._noteIcon      = NOTE_ICONS_TYPES.Comment;
        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._fillColor     = [1, 0.82, 0];

        this._replies = [];
    }
    AscFormat.InitClass(CAnnotationText, AscPDF.CAnnotationBase, AscDFH.historyitem_type_Pdf_Annot_Text);
	CAnnotationText.prototype.constructor = CAnnotationText;
    
    CAnnotationText.prototype.SetState = function(nType) {
        this._state = nType;
    };
    CAnnotationText.prototype.GetState = function() {
        return this._state;
    };
    CAnnotationText.prototype.SetStateModel = function(nType) {
        this._stateModel = nType;
    };
    CAnnotationText.prototype.GetStateModel = function() {
        return this._stateModel;
    };
    CAnnotationText.prototype.ClearReplies = function() {
        this._replies = [];
    };
    CAnnotationText.prototype.AddReply = function(CommentData, nPos) {
        let oReply = new CAnnotationText(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), this.GetDocument());

        oReply.SetContents(CommentData.m_sText);
        oReply.SetCreationDate(CommentData.m_sOOTime);
        oReply.SetModDate(CommentData.m_sOOTime);
        oReply.SetAuthor(CommentData.m_sUserName);
        oReply.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);
        oReply.SetReplyTo(this.GetReplyTo() || this);
        CommentData.SetUserData(oReply.GetId());

        if (!nPos) {
            nPos = this._replies.length;
        }

        this._replies.splice(nPos, 0, oReply);
    };
    CAnnotationText.prototype.GetAscCommentData = function() {
        let oAscCommData = new Asc.asc_CCommentDataWord(null);
        if (null == this.GetContents()) {
            return undefined;
        }

        oAscCommData.asc_putText(this.GetContents());
        let sModDate = this.GetModDate();
        if (sModDate)
            oAscCommData.asc_putOnlyOfficeTime(sModDate.toString());
        oAscCommData.asc_putUserId(editor.documentUserId);
        oAscCommData.asc_putUserName(this.GetAuthor());
        
        let nState = this.GetState();
        let bSolved;
        if (nState == TEXT_ANNOT_STATE.Accepted || nState == TEXT_ANNOT_STATE.Completed)
            bSolved = true;
        oAscCommData.asc_putSolved(bSolved);
        oAscCommData.asc_putQuoteText("");
        oAscCommData.m_sUserData = this.GetId();

        this._replies.forEach(function(reply) {
            oAscCommData.m_aReplies.push(reply.GetAscCommentData());
        });

        return oAscCommData;
    };

    CAnnotationText.prototype.SetIconType = function(nType) {
        this._noteIcon = nType;
    };
    CAnnotationText.prototype.GetIconType = function() {
        return this._noteIcon;
    };
    CAnnotationText.prototype.GetIconDrawFunc = function() {
        let nType = this.GetIconType();
        switch (nType) {
            case NOTE_ICONS_TYPES.Check1:
            case NOTE_ICONS_TYPES.Check2:
                return drawIconCheck;
            case NOTE_ICONS_TYPES.Circle:
                return drawIconCircle;
            case NOTE_ICONS_TYPES.Comment:
                return drawIconComment;
            case NOTE_ICONS_TYPES.Cross:
                return NOTE_ICONS_IMAGES.Cross;
            case NOTE_ICONS_TYPES.CrossH:
                return NOTE_ICONS_IMAGES.CrossH;
            case NOTE_ICONS_TYPES.Help:
                return NOTE_ICONS_IMAGES.Help;
            case NOTE_ICONS_TYPES.Insert:
                return NOTE_ICONS_IMAGES.Insert;
            case NOTE_ICONS_TYPES.Key:
                return NOTE_ICONS_IMAGES.Key;
            case NOTE_ICONS_TYPES.NewParagraph:
                return NOTE_ICONS_IMAGES.NewParagraph;
            case NOTE_ICONS_TYPES.Note:
                return NOTE_ICONS_IMAGES.Note;
            case NOTE_ICONS_TYPES.Paragraph:
                return NOTE_ICONS_IMAGES.Paragraph;
            case NOTE_ICONS_TYPES.RightArrow:
                return NOTE_ICONS_IMAGES.RightArrow;
            case NOTE_ICONS_TYPES.RightPointer:
                return NOTE_ICONS_IMAGES.RightPointer;
            case NOTE_ICONS_TYPES.Star:
                return NOTE_ICONS_IMAGES.Star;
            case NOTE_ICONS_TYPES.UpArrow:
                return NOTE_ICONS_IMAGES.UpArrow;
            case NOTE_ICONS_TYPES.UpLeftArrow:
                return NOTE_ICONS_IMAGES.UpLeftArrow;
        }

        return null;
    };
    CAnnotationText.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oNewAnnot = new CAnnotationText(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oNewAnnot.lazyCopy = true;

        let aFillColor = this.GetFillColor();

        oNewAnnot._originView = this._originView;
        oNewAnnot._apIdx = this._apIdx;
        aFillColor && oNewAnnot.SetFillColor(aFillColor.slice());
        oNewAnnot.SetOriginPage(this.GetOriginPage());
        oNewAnnot.SetAuthor(this.GetAuthor());
        oNewAnnot.SetModDate(this.GetModDate());
        oNewAnnot.SetCreationDate(this.GetCreationDate());
        oNewAnnot.SetContents(this.GetContents());

        oDoc.EndNoHistoryMode();

        return oNewAnnot;
    };
    CAnnotationText.prototype.Draw = function(oGraphics) {
        if (this.IsHidden() == true)
            return;

        // note: oGraphic параметр для рисование track
        if (!this.graphicObjects)
            this.graphicObjects = new AscFormat.DrawingObjectsController(this);

        let oRGB = this.GetRGBColor(this.GetFillColor());

        let oDoc        = this.GetDocument();
        let nPage       = this.GetPage();
        let aOrigRect   = this.GetOrigRect();
        let nRotAngle   = oDoc.Viewer.getPageRotate(nPage);
        
        let nX          = aOrigRect[0];
        let nY          = aOrigRect[1];
        let nWidthPx    = (aOrigRect[2] - aOrigRect[0]);
        let nHeightPx   = (aOrigRect[3] - aOrigRect[1]);
        
        let oCtx = oGraphics.GetContext();
        oGraphics.EnableTransform();
        oCtx.iconFill = "rgb(" + oRGB.r + "," + oRGB.g + "," + oRGB.b + ")";

        let drawFunc = this.GetIconDrawFunc();
        drawFunc(oCtx, nX, nY, nWidthPx / 20, nHeightPx / 20);

        //////
        oGraphics.SetLineWidth(1);
        let aOringRect  = this.GetOrigRect();
        let X       = aOringRect[0];
        let Y       = aOringRect[1];
        let nWidth  = aOringRect[2] - aOringRect[0];
        let nHeight = aOringRect[3] - aOringRect[1];

        Y += 1 / 2;
        X += 1 / 2;
        nWidth  -= 1;
        nHeight -= 1;

        oGraphics.SetStrokeStyle(0, 255, 255);
        oGraphics.SetLineDash([]);
        oGraphics.BeginPath();
        oGraphics.Rect(X, Y, nWidth, nHeight);
        oGraphics.Stroke();
    };
    CAnnotationText.prototype.IsNeedDrawFromStream = function() {
        return false;
    };
    CAnnotationText.prototype.onMouseDown = function(x, y, e) {
        let oViewer         = Asc.editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;

        this.selectStartPage = this.GetPage();

        let pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
    };
    CAnnotationText.prototype.IsComment = function() {
        return true;
    };
    
    CAnnotationText.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // icon
        let nIconType = this.GetIconType();
        if (nIconType != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteByte(this.GetIconType());
        }
        
        // state model
        let nStateModel = this.GetStateModel();
        if (nStateModel != null) {
            memory.annotFlags |= (1 << 17);
            memory.WriteByte(nStateModel);
        }

        // state
        let nState = this.GetState();
        if (nState != null) {
            memory.annotFlags |= (1 << 18);
            memory.WriteByte(nState);
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    
    window["AscPDF"].CAnnotationText            = CAnnotationText;
    window["AscPDF"].TEXT_ANNOT_STATE           = TEXT_ANNOT_STATE;
    window["AscPDF"].TEXT_ANNOT_STATE_MODEL     = TEXT_ANNOT_STATE_MODEL;
	
    function drawIconCheck(ctx, x, y, xScale, yScale) {
        ctx.save();
        ctx.translate(x, y);
        ctx.scale(xScale, yScale);

        // Первый путь
        ctx.beginPath();
        ctx.moveTo(5.238, 8.8);
        ctx.lineTo(4, 11.8);
        ctx.lineTo(7.714, 16);
        ctx.bezierCurveTo(12.048, 9.4, 13.286, 8.2, 17, 4);
        ctx.bezierCurveTo(14.524, 4, 9.778, 8.8, 7.714, 11.8);
        ctx.closePath();
        ctx.fillStyle = ctx.iconFill;
        ctx.fill();

        // Второй путь
        ctx.beginPath();
        ctx.moveTo(12.536, 7.147);
        ctx.bezierCurveTo(10.809, 8.698, 9.134, 10.619, 8.126, 12.083);
        ctx.lineTo(7.75, 12.63);
        ctx.lineTo(5.382, 9.76);
        ctx.lineTo(4.582, 11.702);
        ctx.lineTo(7.656, 15.179);
        ctx.bezierCurveTo(11.244, 9.748, 12.669, 8.122, 15.439, 5.005);
        ctx.quadraticCurveTo(15.252, 5.105, 15.052, 5.225);
        ctx.bezierCurveTo(14.263, 5.7, 13.399, 6.372, 12.536, 7.147);

        ctx.moveTo(14.536, 4.369);
        ctx.bezierCurveTo(15.383, 3.859, 16.241, 3.5, 17, 3.5);
        ctx.lineTo(18.11, 3.5);
        ctx.lineTo(17.374, 4.331);
        ctx.bezierCurveTo(17, 4.755, 16.651, 5.147, 16.321, 5.518);
        ctx.bezierCurveTo(13.388, 8.818, 12.011, 10.368, 8.132, 16.274);
        ctx.lineTo(7.773, 16.821);
        ctx.lineTo(3.419, 11.898);
        ctx.lineTo(5.093, 7.839);
        ctx.lineTo(7.686, 10.979);
        ctx.bezierCurveTo(8.75, 9.539, 10.289, 7.821, 11.868, 6.403);
        ctx.bezierCurveTo(12.759, 5.603, 13.675, 4.887, 14.537, 4.369);
        ctx.fillStyle = '#000';
        ctx.fill();

        ctx.restore();
    }
    function drawIconCircle(ctx, x, y, xScale, yScale) {
        ctx.save();
        ctx.translate(x, y);
        ctx.scale(xScale, yScale);

        // Рисуем внешнюю окружность
        ctx.beginPath();
        ctx.arc(10, 10, 8, 0, 2 * Math.PI);
        ctx.fillStyle = "#000";
        ctx.fill();

        // Рисуем следующую окружность
        ctx.beginPath();
        ctx.arc(10, 10, 7, 0, 2 * Math.PI);
        ctx.fillStyle = ctx.iconFill;
        ctx.fill();

        // Рисуем среднюю окружность
        ctx.beginPath();
        ctx.arc(10, 10, 5, 0, 2 * Math.PI);
        ctx.fillStyle = "#000";
        ctx.fill();

        // Рисуем внутреннюю окружность
        ctx.beginPath();
        ctx.arc(10, 10, 4, 0, 2 * Math.PI);
        ctx.fillStyle = "white";
        ctx.fill();

        ctx.restore();
    }
    function drawIconComment(ctx, x, y, xScale, yScale) {
        ctx.save();
        ctx.translate(x, y);
        ctx.scale(xScale, yScale);

        // Основная фигура прямоугольника с вырезом
        ctx.fillStyle = ctx.iconFill;
        ctx.beginPath();
        ctx.moveTo(3, 3);
        ctx.lineTo(17, 3);
        ctx.quadraticCurveTo(18, 3, 18, 4);
        ctx.lineTo(18, 14);
        ctx.quadraticCurveTo(18, 15, 17, 15);
        ctx.lineTo(10, 15);
        ctx.lineTo(7.5, 17.5);
        ctx.lineTo(5, 15);
        ctx.lineTo(3, 15);
        ctx.quadraticCurveTo(2, 15, 2, 14);
        ctx.lineTo(2, 4);
        ctx.quadraticCurveTo(2, 3, 3, 3);
        ctx.closePath();
        ctx.fill();

        // Обводка темно-серым цветом с дополнительными деталями
        ctx.strokeStyle = '#333333';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(10, 15);
        ctx.lineTo(8.207, 16.793);
        ctx.lineTo(7.5, 17.5);
        ctx.lineTo(6.793, 16.793);
        ctx.lineTo(5, 15);
        ctx.lineTo(3, 15);
        ctx.quadraticCurveTo(2, 15, 2, 14);
        ctx.lineTo(2, 4);
        ctx.quadraticCurveTo(2, 3, 3, 3);
        ctx.lineTo(17, 3);
        ctx.quadraticCurveTo(18, 3, 18, 4);
        ctx.lineTo(18, 14);
        ctx.quadraticCurveTo(18, 15, 17, 15);
        ctx.lineTo(10, 15);
        ctx.lineTo(9.586, 15);
        ctx.closePath();
        ctx.stroke();

        // Линии для текста
        ctx.fillStyle = '#333333';
        ctx.beginPath();
        ctx.rect(5, 6, 10, 1); // Первая линия
        ctx.fill();

        ctx.beginPath();
        ctx.rect(5, 8, 10, 1); // Вторая линия
        ctx.fill();

        ctx.beginPath();
        ctx.rect(5, 10, 7, 1); // Третья линия
        ctx.fill();

        ctx.restore();
    }

    let SVG_ICON_COMMENT = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M2 4C2 3.44772 2.44772 3 3 3H17C17.5523 3 18 3.44772 18 4V14C18 14.5523 17.5523 15 17 15H10L7.5 17.5L5 15H3C2.44772 15 2 14.5523 2 14V4Z' fill='white'/>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M17 15H10L8.20711 16.7929L7.5 17.5L6.79289 16.7929L5 15H3C2.44772 15 2 14.5523 2 14V4C2 3.44772 2.44772 3 3 3H17C17.5523 3 18 3.44772 18 4V14C18 14.5523 17.5523 15 17 15ZM9.29289 14.2929L7.5 16.0858L5.70711 14.2929L5.41421 14H5H3V4H17V14H10H9.58579L9.29289 14.2929ZM15 6H5V7H15V6ZM5 8H15V9H5V8ZM12 10H5V11H12V10Z' fill='#333333'/>\
    </svg>";
    
    let SVG_ICON_CROSS = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='m5.404 4.697-.707.707L9.293 10l-4.596 4.596.707.707L10 10.707l4.596 4.596.707-.707L10.707 10l4.596-4.596-.707-.707L10 9.293z' fill='white'/>\
    <path d='m3.282 5.404 2.122-2.122L10 7.88l4.596-4.597 2.122 2.122L12.12 10l4.597 4.596-2.122 2.122L10 12.12l-4.596 4.597-2.122-2.122L7.88 10zm1.415 0L9.293 10l-4.596 4.596.707.707L10 10.707l4.596 4.596.707-.707L10.707 10l4.596-4.596-.707-.707L10 9.293 5.404 4.697z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_HELP = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M18 10.5a7.5 7.5 0 1 1-15 0 7.5 7.5 0 0 1 15 0' fill='white'/>\
    <path d='M11 14v-1h-1v1z' fill='#000'/>\
    <path d='M18 10.5a7.5 7.5 0 1 1-15 0 7.5 7.5 0 0 1 15 0m-1 0a6.5 6.5 0 1 0-13 0 6.5 6.5 0 0 0 13 0' fill='#000'/>\
    <path d='M10.207 7A1.207 1.207 0 0 0 9 8.207V8.5H8v-.293A2.207 2.207 0 0 1 10.207 6h.586A2.207 2.207 0 0 1 13 8.207V8.5a2.5 2.5 0 0 1-1 2l-.4.3A1.5 1.5 0 0 0 11 12h-1a2.5 2.5 0 0 1 1-2l.4-.3a1.5 1.5 0 0 0 .6-1.2v-.293A1.207 1.207 0 0 0 10.793 7z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_INSERT = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M10 3 2 16h16z' fill='white'/>\
    <path d='m10 5 6.187 10H3.813zm0-2L2 16h16z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_KEY = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M12 13a5 5 0 1 0-4.703-3.297L2.5 14.5l3 3L6 17v-1h1v-1h1v-1h1l1.297-1.297A5 5 0 0 0 12 13' fill='white'/>\
    <path d='m8.454 9.96-4.54 4.54L5 15.586V15h1v-1h1v-1h1.586l1.454-1.454.598.216a4 4 0 1 0-2.4-2.4zM9 14H8v1H7v1H6v1l-.5.5-3-3 4.797-4.797a5 5 0 1 1 3 3z' fill='#000'/>\
    <path d='M14 7a1 1 0 1 1-2 0 1 1 0 0 1 2 0' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_NEW_PARAGRAPH = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M4 7.464a1 1 0 0 1 .354-.763l5.5-4.654a1 1 0 0 1 1.292 0l5.5 4.654a1 1 0 0 1 .354.763V18a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z' fill='white'/>\
    <path d='m5 7.464 5.5-4.654L16 7.464V18H5zM4.354 6.7A1 1 0 0 0 4 7.464V18a1 1 0 0 0 1 1h11a1 1 0 0 0 1-1V7.464a1 1 0 0 0-.354-.763l-5.5-4.654a1 1 0 0 0-1.292 0z' fill='#000'/>\
    <path d='M9.5 8a2.5 2.5 0 1 0 0 5h.5v4h1V9h1v8h1V9h1V8z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_PARAGRAPH = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M14 4h-4v1c-1.5 0-3 .691-3 2.5 0 2 1.62 2.5 3 2.5v6h4z' fill='white'/>\
    <path d='M9.5 11q.255 0 .5-.035V16h4V5h1V4H9.5a3.5 3.5 0 1 0 0 7m3.5 4h-2V5h2zm-3-5c-1.38 0-3-.5-3-2.5C7 5.691 8.5 5 10 5z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_RIGHT_ARROW = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M8 14H2V7h6V3l10 7.5L8 18z' fill='white'/>\
    <path d='M9 13v3l7.333-5.5L9 5v3H3v5zM8 3l10 7.5L8 18v-4H2V7h6z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_RIGHT_POINTER = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M2.628 3.343 18.73 10.5 2.628 17.657 7.399 10.5z' fill='white'/>\
    <path d='M2.628 3.343 18.73 10.5 2.628 17.657 7.399 10.5zm2.744 2.314L8.601 10.5l-3.229 4.843L16.27 10.5z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_STAR = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M11.889 8.111 10 2 8.111 8.111H2l4.944 3.778L5.056 18 10 14.223 14.944 18l-1.888-6.111L18 8.11z' fill='white'/>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M11.889 8.111 10 2 8.111 8.111H2l4.944 3.778L5.056 18 10 14.223 14.944 18l-1.888-6.111L18 8.11zm3.155 1H11.15L10 5.387 8.85 9.111H4.955l3.15 2.406-1.171 3.79L10 12.963l3.065 2.342-1.17-3.789z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_NOTE = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M4 5a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v7.086a1 1 0 0 1-.293.707l-2.914 2.914a1 1 0 0 1-.707.293H5a1 1 0 0 1-1-1z' fill='white'/>\
    <path d='M15 5v7h-2a1 1 0 0 0-1 1v2H5V5zm-.414 8L13 14.586V13zM4 15a1 1 0 0 0 1 1h7.586a1 1 0 0 0 .707-.293l2.414-2.414a1 1 0 0 0 .293-.707V5a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_UP_ARROW = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='m3 12 7.5-10L18 12h-4v6H7v-6z' fill='white'/>\
    <path d='M13 11h3l-5.5-7.333L5 11h3v6h5zM3 12l7.5-10L18 12h-4v6H7v-6z' fill='#000'/>\
    </svg>";
    
    let SVG_ICON_UP_LEFT_ARROW = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M14.5 4.5h-10v10l3-3 5 5 4-4-5-5z' fill='white'/>\
    <path d='M4 4h11.707l-3.5 3.5 5 5-4.707 4.707-5-5-3.5 3.5zm1 1v8.293l2.5-2.5 5 5 3.293-3.293-5-5 2.5-2.5z' fill='#000'/>\
    </svg>";
	
})();

