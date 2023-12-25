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
    function CAnnotationText(sName, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Text, nPage, aRect, oDoc);

        this._noteIcon      = NOTE_ICONS_TYPES.Comment;
        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._strokeColor   = [1, 0.82, 0];

        // internal
        TurnOffHistory();
        this._replies = [];
    }
    CAnnotationText.prototype = Object.create(AscPDF.CAnnotationBase.prototype);
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
    CAnnotationText.prototype.AddReply = function(CommentData) {
        let oReply = new CAnnotationText(AscCommon.CreateGUID(), this.GetPage(), this.GetRect().slice(), this.GetDocument());

        oReply.SetContents(CommentData.m_sText);
        oReply.SetCreationDate(CommentData.m_sOOTime);
        oReply.SetModDate(CommentData.m_sOOTime);
        oReply.SetAuthor(CommentData.m_sUserName);
        oReply.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);
        oReply.SetReplyTo(this.GetReplyTo() || this);

        oReply.SetApIdx(this.GetDocument().GetMaxApIdx() + 2);
        CommentData.m_sUserData = oReply.GetApIdx();

        this._replies.push(oReply);
    };
    CAnnotationText.prototype.GetAscCommentData = function() {
        let oAscCommData = new Asc["asc_CCommentDataWord"](null);
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
        oAscCommData.m_sUserData = this.GetApIdx();

        this._replies.forEach(function(reply) {
            oAscCommData.m_aReplies.push(reply.GetAscCommentData());
        });

        return oAscCommData;
    };

    CAnnotationText.prototype.GetIconType = function() {
        return this._noteIcon;
    };
    CAnnotationText.prototype.GetIconImg = function() {
        let nType = this.GetIconType();
        switch (nType) {
            case NOTE_ICONS_TYPES.Check1:
                return NOTE_ICONS_IMAGES.Check1;
            case NOTE_ICONS_TYPES.Check2:
                return NOTE_ICONS_IMAGES.Check2;
            case NOTE_ICONS_TYPES.Circle:
                return NOTE_ICONS_IMAGES.Circle;
            case NOTE_ICONS_TYPES.Comment:
                return NOTE_ICONS_IMAGES.Comment;
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
        oDoc.TurnOffHistory();

        let oNewAnnot = new CAnnotationText(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

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
        oNewAnnot.SetFillColor(this.GetFillColor());
        oNewAnnot.SetOriginPage(this.GetOriginPage());
        oNewAnnot.SetAuthor(this.GetAuthor());
        oNewAnnot.SetModDate(this.GetModDate());
        oNewAnnot.SetCreationDate(this.GetCreationDate());
        oNewAnnot.SetContents(this.GetContents());

        return oNewAnnot;
    };
    CAnnotationText.prototype.Draw = function(oGraphics) {
        if (this.IsHidden() == true)
            return;

        // note: oGraphic параметр для рисование track
        if (!this.graphicObjects)
            this.graphicObjects = new AscFormat.DrawingObjectsController(this);

        let oRGB            = this.GetRGBColor(this.GetFillColor());
        let ICON_TO_DRAW    = this.GetIconImg();

        let aRect       = this.GetRect();
        let aOrigRect   = this.GetOrigRect();

        let nWidth  = (aRect[2] - aRect[0]) * AscCommon.AscBrowser.retinaPixelRatio;
        let nHeight = (aRect[3] - aRect[1]) * AscCommon.AscBrowser.retinaPixelRatio;
        
        let imgW = ICON_TO_DRAW.width;
        let imgH = ICON_TO_DRAW.height;

        let nScaleX = nWidth / imgW;
        let nScaleY = nHeight / imgH;
        let wScaled = Math.max(imgW * nScaleX + 0.5 >> 0, 40);
        let hScaled = Math.max(imgH * nScaleY + 0.5 >> 0, 40);

        let canvas = document.createElement('canvas');
        let context = canvas.getContext('2d');

        // Set the canvas dimensions to match the image
        canvas.width = wScaled;
        canvas.height = hScaled;

        // Draw the image onto the canvas
        context.drawImage(ICON_TO_DRAW, 0, 0, imgW, imgH, 0, 0, wScaled, hScaled);

        if (!AscCommon.AscBrowser.isIE || AscCommon.AscBrowser.isIeEdge) {
            if (oRGB.r != 255 || oRGB.g != 209 || oRGB.b != 0) {
                // Get the pixel data of the canvas
                let imageData = context.getImageData(0, 0, canvas.width, canvas.height);
                let data = imageData.data;
    
                // Loop through each pixel
                for (let i = 0; i < data.length; i += 4) {
                    const red = data[i];
                    const green = data[i + 1];
                    const blue = data[i + 2];
    
                    // Check if the pixel is black (R = 0, G = 0, B = 0)
                    if (red === 255 && green === 209 && blue === 0) {
                        // Change the pixel color to red (R = 255, G = 0, B = 0)
                        data[i] = oRGB.r; // Red
                        data[i + 1] = oRGB.g; // Green
                        data[i + 2] = oRGB.b; // Blue
                        // Note: The alpha channel (transparency) remains unchanged
                    }
                }
    
                // Put the modified pixel data back onto the canvas
                context.putImageData(imageData, 0, 0);
            }
        }

        // Draw the comment note
        oGraphics.SetIntegerGrid(true);
        oGraphics.DrawImageXY(canvas, aOrigRect[0], aOrigRect[1]);
        oGraphics.SetIntegerGrid(false);
    };
    CAnnotationText.prototype.IsNeedDrawFromStream = function() {
        return false;
    };
    CAnnotationText.prototype.onMouseDown = function(e) {
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
    CAnnotationText.prototype.IsComment = function() {
        return true;
    };
    CAnnotationText.prototype.getObjectType = function() {
        return -1;
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

        this._replies.forEach(function(reply) {
            reply.WriteToBinary(memory); 
        });
    };
    // CAnnotationText.prototype.ClearCache = function() {};
    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    window["AscPDF"].CAnnotationText            = CAnnotationText;
    window["AscPDF"].TEXT_ANNOT_STATE           = TEXT_ANNOT_STATE;
    window["AscPDF"].TEXT_ANNOT_STATE_MODEL     = TEXT_ANNOT_STATE_MODEL;
	
	function toBase64(str) {
		return window.btoa(unescape(encodeURIComponent(str)));
	}
	
	function getSvgImage(svg) {
		let image = new Image();
		if (!AscCommon.AscBrowser.isIE || AscCommon.AscBrowser.isIeEdge) {
			image.src = "data:image/svg+xml;utf8," + encodeURIComponent(svg);
		}
		else {
			image.src = "data:image/svg+xml;base64," + toBase64(svg);
			image.onload = function() {
				// Почему-то IE не определяет размеры сам
				this.width = 20;
				this.height = 20;
			};
		}
		
		return image;
	}

    let SVG_ICON_CHECK = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M5.2381 8.8L4 11.8L7.71429 16C12.0476 9.4 13.2857 8.2 17 4C14.5238 4 9.77778 8.8 7.71429 11.8L5.2381 8.8Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_CIRCLE = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M10 17C13.866 17 17 13.866 17 10C17 6.13401 13.866 3 10 3C6.13401 3 3 6.13401 3 10C3 13.866 6.13401 17 10 17ZM10 18C14.4183 18 18 14.4183 18 10C18 5.58172 14.4183 2 10 2C5.58172 2 2 5.58172 2 10C2 14.4183 5.58172 18 10 18Z' fill='black'/>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M10 14C12.2091 14 14 12.2091 14 10C14 7.79086 12.2091 6 10 6C7.79086 6 6 7.79086 6 10C6 12.2091 7.79086 14 10 14ZM10 15C12.7614 15 15 12.7614 15 10C15 7.23858 12.7614 5 10 5C7.23858 5 5 7.23858 5 10C5 12.7614 7.23858 15 10 15Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_COMMENT = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M2 4C2 3.44772 2.44772 3 3 3H17C17.5523 3 18 3.44772 18 4V14C18 14.5523 17.5523 15 17 15H10L7.5 17.5L5 15H3C2.44772 15 2 14.5523 2 14V4Z' fill='#FFD100'/>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M17 15H10L8.20711 16.7929L7.5 17.5L6.79289 16.7929L5 15H3C2.44772 15 2 14.5523 2 14V4C2 3.44772 2.44772 3 3 3H17C17.5523 3 18 3.44772 18 4V14C18 14.5523 17.5523 15 17 15ZM9.29289 14.2929L7.5 16.0858L5.70711 14.2929L5.41421 14H5H3V4H17V14H10H9.58579L9.29289 14.2929ZM15 6H5V7H15V6ZM5 8H15V9H5V8ZM12 10H5V11H12V10Z' fill='#333333'/>\
    </svg>";
    
    let SVG_ICON_CROSS = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M4.81055 5.70711L5.4718 6.36836L6.7943 7.69087L9.29585 10.1924L7.00452 12.4838L5.57703 13.9113L4.86328 14.625L5.57039 15.3321L6.28413 14.6184L7.71162 13.1909L10.003 10.8995L12.2943 13.1909L13.7218 14.6184L14.4355 15.3321L15.1427 14.625L14.4289 13.9113L13.0014 12.4838L10.7101 10.1924L13.2116 7.69087L14.5341 6.36836L15.1954 5.70711L14.4883 5L13.827 5.66125L12.5045 6.98377L10.003 9.48533L7.50141 6.98377L6.17891 5.66126L5.51766 5L4.81055 5.70711Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_HELP = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M17 10.5C17 14.0899 14.0899 17 10.5 17C6.91015 17 4 14.0899 4 10.5C4 6.91015 6.91015 4 10.5 4C14.0899 4 17 6.91015 17 10.5ZM18 10.5C18 14.6421 14.6421 18 10.5 18C6.35786 18 3 14.6421 3 10.5C3 6.35786 6.35786 3 10.5 3C14.6421 3 18 6.35786 18 10.5ZM11 14V13H10V14H11Z' fill='black'/>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M10.2071 7C9.88696 7 9.57993 7.12718 9.35355 7.35355C9.12718 7.57993 9 7.88696 9 8.20711V8.5H8V8.20711C8 7.62175 8.23253 7.06036 8.64645 6.64645C9.06036 6.23253 9.62175 6 10.2071 6H10.7929C11.3783 6 11.9396 6.23253 12.3536 6.64645C12.7675 7.06036 13 7.62175 13 8.20711V8.5C13 9.28689 12.6295 10.0279 12 10.5L11.6 10.8C11.2223 11.0833 11 11.5279 11 12H10C10 11.2131 10.3705 10.4721 11 10L11.4 9.7C11.7777 9.41672 12 8.97214 12 8.5V8.20711C12 7.88696 11.8728 7.57993 11.6464 7.35355C11.4201 7.12718 11.113 7 10.7929 7H10.2071Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_INSERT = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M10 5L16.1867 15H3.81333L10 5ZM10 3L2 16H18L10 3Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_KEY = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <mask id='path-1-inside-1_9139_69160' fill='white'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M12 13C14.7614 13 17 10.7614 17 8C17 5.23858 14.7614 3 12 3C9.23858 3 7 5.23858 7 8C7 8.59785 7.10493 9.1712 7.29737 9.70263L2.5 14.5L5.5 17.5L6 17L6 16H7V15H8L8 14H9L10.2974 12.7026C10.8288 12.8951 11.4021 13 12 13Z'/>\
    </mask>\
    <path d='M7.29737 9.70263L8.00448 10.4097L8.45415 9.96006L8.23762 9.36213L7.29737 9.70263ZM2.5 14.5L1.79289 13.7929L1.08579 14.5L1.79289 15.2071L2.5 14.5ZM5.5 17.5L4.79289 18.2071L5.5 18.9142L6.20711 18.2071L5.5 17.5ZM6 17L6.70711 17.7071L7 17.4142L7 17L6 17ZM6 16V15H5L5 16L6 16ZM7 16V17H8V16H7ZM7 15V14H6V15H7ZM8 15V16H9L9 15L8 15ZM8 14V13H7L7 14L8 14ZM9 14V15H9.41421L9.70711 14.7071L9 14ZM10.2974 12.7026L10.6379 11.7624L10.0399 11.5458L9.59027 11.9955L10.2974 12.7026ZM16 8C16 10.2091 14.2091 12 12 12V14C15.3137 14 18 11.3137 18 8H16ZM12 4C14.2091 4 16 5.79086 16 8H18C18 4.68629 15.3137 2 12 2V4ZM8 8C8 5.79086 9.79086 4 12 4V2C8.68629 2 6 4.68629 6 8H8ZM8.23762 9.36213C8.08414 8.9383 8 8.48008 8 8H6C6 8.71562 6.12572 9.4041 6.35713 10.0431L8.23762 9.36213ZM3.20711 15.2071L8.00448 10.4097L6.59027 8.99552L1.79289 13.7929L3.20711 15.2071ZM6.20711 16.7929L3.20711 13.7929L1.79289 15.2071L4.79289 18.2071L6.20711 16.7929ZM5.29289 16.2929L4.79289 16.7929L6.20711 18.2071L6.70711 17.7071L5.29289 16.2929ZM5 16L5 17L7 17L7 16L5 16ZM7 15H6V17H7V15ZM6 15V16H8V15H6ZM8 14H7V16H8V14ZM7 14L7 15L9 15L9 14L7 14ZM9 13H8V15H9V13ZM9.59027 11.9955L8.29289 13.2929L9.70711 14.7071L11.0045 13.4097L9.59027 11.9955ZM12 12C11.5199 12 11.0617 11.9159 10.6379 11.7624L9.95688 13.6429C10.5959 13.8743 11.2844 14 12 14V12Z' fill='black' mask='url(#path-1-inside-1_9139_69160)'/>\
    <circle cx='13' cy='7' r='1' fill='black'/>\
    </svg>";
    
    let SVG_ICON_NEW_PARAGRAPH = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <mask id='path-1-inside-1_9139_69160' fill='white'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M12 13C14.7614 13 17 10.7614 17 8C17 5.23858 14.7614 3 12 3C9.23858 3 7 5.23858 7 8C7 8.59785 7.10493 9.1712 7.29737 9.70263L2.5 14.5L5.5 17.5L6 17L6 16H7V15H8L8 14H9L10.2974 12.7026C10.8288 12.8951 11.4021 13 12 13Z'/>\
    </mask>\
    <path d='M7.29737 9.70263L8.00448 10.4097L8.45415 9.96006L8.23762 9.36213L7.29737 9.70263ZM2.5 14.5L1.79289 13.7929L1.08579 14.5L1.79289 15.2071L2.5 14.5ZM5.5 17.5L4.79289 18.2071L5.5 18.9142L6.20711 18.2071L5.5 17.5ZM6 17L6.70711 17.7071L7 17.4142L7 17L6 17ZM6 16V15H5L5 16L6 16ZM7 16V17H8V16H7ZM7 15V14H6V15H7ZM8 15V16H9L9 15L8 15ZM8 14V13H7L7 14L8 14ZM9 14V15H9.41421L9.70711 14.7071L9 14ZM10.2974 12.7026L10.6379 11.7624L10.0399 11.5458L9.59027 11.9955L10.2974 12.7026ZM16 8C16 10.2091 14.2091 12 12 12V14C15.3137 14 18 11.3137 18 8H16ZM12 4C14.2091 4 16 5.79086 16 8H18C18 4.68629 15.3137 2 12 2V4ZM8 8C8 5.79086 9.79086 4 12 4V2C8.68629 2 6 4.68629 6 8H8ZM8.23762 9.36213C8.08414 8.9383 8 8.48008 8 8H6C6 8.71562 6.12572 9.4041 6.35713 10.0431L8.23762 9.36213ZM3.20711 15.2071L8.00448 10.4097L6.59027 8.99552L1.79289 13.7929L3.20711 15.2071ZM6.20711 16.7929L3.20711 13.7929L1.79289 15.2071L4.79289 18.2071L6.20711 16.7929ZM5.29289 16.2929L4.79289 16.7929L6.20711 18.2071L6.70711 17.7071L5.29289 16.2929ZM5 16L5 17L7 17L7 16L5 16ZM7 15H6V17H7V15ZM6 15V16H8V15H6ZM8 14H7V16H8V14ZM7 14L7 15L9 15L9 14L7 14ZM9 13H8V15H9V13ZM9.59027 11.9955L8.29289 13.2929L9.70711 14.7071L11.0045 13.4097L9.59027 11.9955ZM12 12C11.5199 12 11.0617 11.9159 10.6379 11.7624L9.95688 13.6429C10.5959 13.8743 11.2844 14 12 14V12Z' fill='black' mask='url(#path-1-inside-1_9139_69160)'/>\
    <circle cx='13' cy='7' r='1' fill='black'/>\
    </svg>";
    
    let SVG_ICON_PARAGRAPH = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M8.5 11C8.66976 11 8.8367 10.9879 9 10.9646V16H10V5H12V16H13V5H14V4C12.165 4 10.2782 4 8.5 4C6.567 4 5 5.567 5 7.5C5 9.433 6.567 11 8.5 11Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_RIGHT_ARROW = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M9 13V16L16.3333 10.5L9 5V8H3V13H9ZM8 3L18 10.5L8 18V14H2V7H8V3Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_RIGHT_POINTER = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M2.62769 3.34302L18.7311 10.5001L2.62769 17.6572L7.39907 10.5001L2.62769 3.34302ZM5.3723 5.65716L8.60092 10.5001L5.3723 15.343L16.2689 10.5001L5.3723 5.65716Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_STAR = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M11.8885 8.11146L10 2L8.11146 8.11146H2L6.94427 11.8885L5.05573 18L10 14.2229L14.9443 18L13.0557 11.8885L18 8.11146H11.8885ZM15.0437 9.11146H11.1509L10 5.38705L8.8491 9.11146H4.9563L8.10564 11.5173L6.93486 15.3061L10 12.9645L13.0651 15.3061L11.8944 11.5173L15.0437 9.11146Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_NOTE = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M15 5H5L5 15H12V13C12 12.4477 12.4477 12 13 12H15V5ZM14.5858 13L13 14.5858V13H14.5858ZM4 15C4 15.5523 4.44772 16 5 16H12.5858C12.851 16 13.1054 15.8946 13.2929 15.7071L15.7071 13.2929C15.8946 13.1054 16 12.851 16 12.5858V5C16 4.44771 15.5523 4 15 4H5C4.44772 4 4 4.44772 4 5V15Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_UP_ARROW = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path fill-rule='evenodd' clip-rule='evenodd' d='M13 11L16 11L10.5 3.66667L5 11L8 11L8 17L13 17L13 11ZM3 12L10.5 2L18 12L14 12L14 18L7 18L7 12L3 12Z' fill='black'/>\
    </svg>";
    
    let SVG_ICON_UP_LEFT_ARROW = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M14.5 4.5H4.5V14.5L7.5 11.5L12.5 16.5L16.5 12.5L11.5 7.5L14.5 4.5Z' stroke='black'/>\
    </svg>";
	
	let NOTE_ICONS_IMAGES = {
		Check:          getSvgImage(SVG_ICON_CHECK),
		Circle:         getSvgImage(SVG_ICON_CIRCLE),
		Comment:        getSvgImage(SVG_ICON_COMMENT),
		Cross:          getSvgImage(SVG_ICON_CROSS),
		Help:           getSvgImage(SVG_ICON_HELP),
		Insert:         getSvgImage(SVG_ICON_INSERT),
		Key:            getSvgImage(SVG_ICON_KEY),
		NewParagraph:   getSvgImage(SVG_ICON_NEW_PARAGRAPH),
		Note:           getSvgImage(SVG_ICON_NOTE),
		Paragraph:      getSvgImage(SVG_ICON_PARAGRAPH),
		RightArrow:     getSvgImage(SVG_ICON_RIGHT_ARROW),
		RightPointer:   getSvgImage(SVG_ICON_RIGHT_POINTER),
		Star:           getSvgImage(SVG_ICON_STAR),
		UpArrow:        getSvgImage(SVG_ICON_UP_ARROW),
		UpLeftArrow:    getSvgImage(SVG_ICON_UP_LEFT_ARROW)
	}
    
})();

