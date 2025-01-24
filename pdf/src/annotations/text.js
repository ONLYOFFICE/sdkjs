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

    let HALF_SIZE = 11;
    
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
    
    CAnnotationText.prototype.select = AscFormat.CGraphicObjectBase.prototype.select;
    CAnnotationText.prototype.deselect = AscFormat.CGraphicObjectBase.prototype.deselect;
    CAnnotationText.prototype.canChangeAdjustments = function() {};
    CAnnotationText.prototype.hitToHandles = function() {};
    CAnnotationText.prototype.hitInBoundingRect = function() {};
    CAnnotationText.prototype.getNoChangeAspect = function() {};
    CAnnotationText.prototype.getMainGroup = function() {};
    CAnnotationText.prototype.getObjectName = function() {};
    CAnnotationText.prototype.isShape = function() {};
    CAnnotationText.prototype.isImage = function() {};
    CAnnotationText.prototype.canMove = function() {
        return true;
    };
    CAnnotationText.prototype.canResize = function() {
        return false;
    };
    CAnnotationText.prototype.canRotate = function() {
        return false;
    };
    
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

        oReply.SetCreationDate(CommentData.m_sOOTime);
        oReply.SetModDate(CommentData.m_sOOTime);
        oReply.SetAuthor(CommentData.m_sUserName);
        oReply.SetUserId(CommentData.m_sUserId);
        oReply.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);
        oReply.SetReplyTo(this.GetReplyTo() || this);
        CommentData.SetUserData(oReply.GetId());
        oReply.SetContents(CommentData.m_sText);

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
        oAscCommData.asc_putUserId(this.GetUserId());
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
                return drawIconCross;
            case NOTE_ICONS_TYPES.CrossH:
                return drawIconCrossHairs;
            case NOTE_ICONS_TYPES.Help:
                return drawIconHelp;
            case NOTE_ICONS_TYPES.Insert:
                return drawIconInsert;
            case NOTE_ICONS_TYPES.Key:
                return drawIconKey;
            case NOTE_ICONS_TYPES.NewParagraph:
                return drawIconNewParagraph;
            case NOTE_ICONS_TYPES.Note:
                return drawIconNote;
            case NOTE_ICONS_TYPES.Paragraph:
                return drawIconParagraph;
            case NOTE_ICONS_TYPES.RightArrow:
                return drawIconRightArrow;
            case NOTE_ICONS_TYPES.RightPointer:
                return drawIconRightPointer;
            case NOTE_ICONS_TYPES.Star:
                return drawIconStar;
            case NOTE_ICONS_TYPES.UpArrow:
                return drawIconUpArrow;
            case NOTE_ICONS_TYPES.UpLeftArrow:
                return drawIconUpLeftArrow;
        }

        return null;
    };
    CAnnotationText.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oNewAnnot = new CAnnotationText(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oNewAnnot.lazyCopy = true;
        oNewAnnot._originView = this._originView;
        oNewAnnot._apIdx = this._apIdx;

        let aFillColor = this.GetFillColor();
        aFillColor && oNewAnnot.SetFillColor(aFillColor.slice());
        oNewAnnot.SetOriginPage(this.GetOriginPage());
        oNewAnnot.SetAuthor(this.GetAuthor());
        oNewAnnot.SetModDate(this.GetModDate());
        oNewAnnot.SetCreationDate(this.GetCreationDate());
        oNewAnnot.SetContents(this.GetContents());
        oNewAnnot.SetIconType(this.GetIconType());

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
        let nWidth      = 21 / oDoc.Viewer.zoom;
        let nHeight     = 21 / oDoc.Viewer.zoom;
        
        let oCtx = oGraphics.GetContext();
        oCtx.save();
        oGraphics.EnableTransform();
        oCtx.iconFill = "rgb(" + oRGB.r + "," + oRGB.g + "," + oRGB.b + ")";

        let nScale = 1 / oDoc.Viewer.zoom;
        let drawFunc = this.GetIconDrawFunc();
        drawFunc(oCtx, nX , nY, nScale, nScale, -nRotAngle * Math.PI / 180);

        oCtx.restore();

        let aRegions = [[
            [nX, nY],
            [nX + nWidth, nY],
            [nX + nWidth, nY + nHeight],
            [nX, nY + nHeight]
        ]];

        oGraphics.DrawLockObjectRect(this.Lock.Get_Type(), aRegions);
        //// draw rect
        // oGraphics.SetLineWidth(1);
        // oGraphics.SetStrokeStyle(0, 255, 255);
        // oGraphics.SetLineDash([]);
        // oGraphics.BeginPath();
        // oGraphics.Rect(nX, nY, nWidth, nHeight);
        // oGraphics.Stroke();
    };
    CAnnotationText.prototype.drawLocks = function (transform, oGraphicsPDF) {
		if (AscCommon.IsShapeToImageConverter) {
			return;
		}
		var bNotes = !!(this.parent && this.parent.kind === AscFormat.TYPE_KIND.NOTES);
		if (!this.group && !bNotes) {
			var oLock;
			if (this.parent instanceof AscCommonWord.ParaDrawing) {
				oLock = this.parent.Lock;
			} else if (this.Lock) {
				oLock = this.Lock;
			}
			if (oLock && AscCommon.c_oAscLockTypes.kLockTypeNone !== oLock.Get_Type()) {
				var bCoMarksDraw = true;
				var oApi = editor || Asc['editor'];
				if (oApi) {

					switch (oApi.getEditorId()) {
						case AscCommon.c_oEditorId.Word: {
							bCoMarksDraw = (true === oApi.isCoMarksDraw || AscCommon.c_oAscLockTypes.kLockTypeMine !== oLock.Get_Type());
							break;
						}
						case AscCommon.c_oEditorId.Presentation: {
							bCoMarksDraw = (!AscCommon.CollaborativeEditing.Is_Fast() || AscCommon.c_oAscLockTypes.kLockTypeMine !== oLock.Get_Type());
							break;
						}
						case AscCommon.c_oEditorId.Spreadsheet: {
							bCoMarksDraw = (!oApi.collaborativeEditing.getFast() || AscCommon.c_oAscLockTypes.kLockTypeMine !== oLock.Get_Type());
							break;
						}
					}
				}
				if (bCoMarksDraw && oGraphicsPDF.DrawLockObjectRect) {
					oGraphicsPDF.transform3(transform);
					oGraphicsPDF.DrawLockObjectRect(oLock.Get_Type(), 0, 0, this.extX, this.extY);
					return true;
				}
			}
		}
		return false;
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
    CAnnotationText.prototype.DrawSelected = function(overlay) {
        overlay.m_oContext.lineWidth    = 3;
        overlay.m_oContext.globalAlpha  = 1;
        overlay.m_oContext.strokeStyle  = "rgb(33, 117, 200)";
        overlay.m_oContext.beginPath();

        let oViewer     = Asc.editor.getDocumentRenderer();
        let aOrigRect   = this.GetRect();
        let nX          = aOrigRect[0] + 0.5 >> 0;
        let nY          = aOrigRect[1] + 0.5 >> 0;
        let nWidth      = 21 / (oViewer.zoom);
        let nHeight     = 21 / (oViewer.zoom);

        let aRegions = [[
            [nX + nWidth, nY],
            [nX, nY],
            [nX, nY + nHeight],
            [nX + nWidth, nY + nHeight]
        ]];

        AscPDF.fillRegion({regions: aRegions}, overlay, this.GetPage());
        overlay.m_oContext.stroke();
    }
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
	
    function drawIconCheck(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(18.425,4.075);
        ctx.translate(17.7175,4.781713343584229);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.7848419132827078,-2.3567507403070853,1);
        ctx.rotate(0);
        ctx.translate(-17.7175,-4.781713343584229);
        ctx.lineTo(8.525000000000002,12.561);
        ctx.lineTo(4.99,9.025);
        ctx.translate(4.282500000000001,9.731713343584229);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.784841913282707,-2.356750740307086,1);
        ctx.rotate(0);
        ctx.translate(-4.282500000000001,-9.731713343584229);
        ctx.lineTo(1.4540000000000002,11.146);
        ctx.translate(2.1607133435842294,11.8535);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.3556382400776044,-3.927547067101982,1);
        ctx.rotate(0);
        ctx.translate(-2.1607133435842294,-11.8535);
        ctx.lineTo(5.696,16.802);
        ctx.lineTo(5.696,16.803);
        ctx.lineTo(7.818,18.925);
        ctx.translate(8.524999999999999,18.217786453749646);
        ctx.rotate(0);
        ctx.arc(0,0,1,2.3560434901900495,0.785549163399744,1);
        ctx.rotate(0);
        ctx.translate(-8.524999999999999,-18.217786453749646);
        ctx.lineTo(11.354,16.803);
        ctx.lineTo(20.546,7.611000000000001);
        ctx.translate(19.838786453749645,6.904000000000001);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7852471633951528,-0.7852471633951528,1);
        ctx.rotate(0);
        ctx.translate(-19.838786453749645,-6.904000000000001);
        ctx.lineTo(18.425,4.075000000000001);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(20.546,7.61);
        ctx.translate(19.838286957870636,6.9035);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7845404129767302,-0.7845404129767302,1);
        ctx.rotate(0);
        ctx.translate(-19.838286957870636,-6.9035);
        ctx.lineTo(18.425,4.075);
        ctx.translate(17.7175,4.781713343584229);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.7848419132827078,-2.3567507403070853,1);
        ctx.rotate(0);
        ctx.translate(-17.7175,-4.781713343584229);
        ctx.lineTo(8.525000000000002,12.561);
        ctx.lineTo(4.99,9.025);
        ctx.translate(4.282500000000001,9.731713343584229);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.784841913282707,-2.356750740307086,1);
        ctx.rotate(0);
        ctx.translate(-4.282500000000001,-9.731713343584229);
        ctx.lineTo(1.4540000000000002,11.146);
        ctx.translate(2.1607133435842294,11.8535);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.3556382400776044,-3.927547067101982,1);
        ctx.rotate(0);
        ctx.translate(-2.1607133435842294,-11.8535);
        ctx.lineTo(7.818,18.925);
        ctx.translate(8.524999999999999,18.217786453749646);
        ctx.rotate(0);
        ctx.arc(0,0,1,2.3560434901900495,0.785549163399744,1);
        ctx.rotate(0);
        ctx.translate(-8.524999999999999,-18.217786453749646);
        ctx.lineTo(20.546,7.61);
        ctx.closePath();
        ctx.moveTo(8.526,13.975000000000001);
        ctx.lineTo(4.281,9.732);
        ctx.lineTo(2.1609999999999996,11.854);
        ctx.lineTo(6.401999999999999,16.095);
        ctx.lineTo(6.401999999999999,16.096);
        ctx.lineTo(8.524,18.218);
        ctx.lineTo(10.645,16.096);
        ctx.lineTo(10.645999999999999,16.096);
        ctx.lineTo(19.838,6.904);
        ctx.lineTo(17.717000000000002,4.782);
        ctx.lineTo(8.524000000000003,13.975);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconCircle(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(11,21);
        ctx.bezierCurveTo(16.523,21,21,16.523,21,11);
        ctx.bezierCurveTo(21,5.477,16.523,1,11,1);
        ctx.bezierCurveTo(5.477,1,1,5.477,1,11);
        ctx.bezierCurveTo(1,16.523,5.477,21,11,21);
        ctx.closePath();
        ctx.moveTo(11,17);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,6,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,6,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(11,20);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,9,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,9,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.moveTo(17.21,11);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,6.21,0,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,6.21,3.141592653589793,6.283185307179586,0);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.closePath();
        ctx.moveTo(21,11);
        ctx.bezierCurveTo(21,16.523,16.523,21,11,21);
        ctx.bezierCurveTo(5.477,21,1,16.523,1,11);
        ctx.bezierCurveTo(1,5.477,5.477,1,11,1);
        ctx.bezierCurveTo(16.523,1,21,5.477,21,11);
        ctx.closePath();
        ctx.moveTo(16.21,11);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,5.21,0,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,5.21,3.141592653589793,6.283185307179586,0);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.restore();
    }
    function drawIconComment(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(2,3);
        ctx.translate(2,4);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-2,-4);
        ctx.lineTo(1,15);
        ctx.translate(2,15);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-2,-15);
        ctx.lineTo(4,16);
        ctx.lineTo(4,19.913);
        ctx.translate(4.499999998917781,19.913032897086396);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-3.1415268594180197,-5.419973419959298,1);
        ctx.rotate(0);
        ctx.translate(-4.499999998917781,-19.913032897086396);
        ctx.lineTo(9.833,16);
        ctx.lineTo(19,16);
        ctx.translate(19,15);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-19,-15);
        ctx.lineTo(20,4);
        ctx.translate(19,4);
        ctx.rotate(0);
        ctx.arc(0,0,1,0,-1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-19,-4);
        ctx.lineTo(2,3);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(5,15);
        ctx.lineTo(5,18.826);
        ctx.lineTo(9.463,15);
        ctx.lineTo(19,15);
        ctx.lineTo(19,4);
        ctx.lineTo(2,4);
        ctx.lineTo(2,15);
        ctx.lineTo(5,15);
        ctx.closePath();
        ctx.moveTo(4.825,20.293);
        ctx.lineTo(9.833,16);
        ctx.lineTo(19,16);
        ctx.translate(19,15);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-19,-15);
        ctx.lineTo(20,4);
        ctx.translate(19,4);
        ctx.rotate(0);
        ctx.arc(0,0,1,0,-1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-19,-4);
        ctx.lineTo(2,3);
        ctx.translate(2,4);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-2,-4);
        ctx.lineTo(1,15);
        ctx.translate(2,15);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-2,-15);
        ctx.lineTo(4,16);
        ctx.lineTo(4,19.913);
        ctx.translate(4.499999998917781,19.913032897086396);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-3.1415268594180197,-5.419973419959298,1);
        ctx.rotate(0);
        ctx.translate(-4.499999998917781,-19.913032897086396);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(5,7.5);
        ctx.translate(5.5,7.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-5.5,-7.5);
        ctx.lineTo(15.5,7);
        ctx.translate(15.5,7.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.5707963267948966,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-15.5,-7.5);
        ctx.lineTo(5.5,8);
        ctx.translate(5.5,7.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-5.5,-7.5);
        ctx.closePath();
        ctx.moveTo(5,11.5);
        ctx.translate(5.5,11.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-5.5,-11.5);
        ctx.lineTo(12.5,11);
        ctx.translate(12.5,11.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.5707963267948966,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-12.5,-11.5);
        ctx.lineTo(5.5,12);
        ctx.translate(5.5,11.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-5.5,-11.5);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconCross(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(19.485,16.657);
        ctx.translate(18.777786453749645,15.95);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7852471633951528,-0.7852471633951528,1);
        ctx.rotate(0);
        ctx.translate(-18.777786453749645,-15.95);
        ctx.lineTo(15.243,11);
        ctx.lineTo(19.485,6.758);
        ctx.translate(18.77828665641577,6.0504999999999995);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7859544135121894,-0.7859544135121896,1);
        ctx.rotate(0);
        ctx.translate(-18.77828665641577,-6.0504999999999995);
        ctx.lineTo(16.657,2.515);
        ctx.translate(15.9495,3.2217133435842293);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.7848419132827078,-2.3567507403070853,1);
        ctx.rotate(0);
        ctx.translate(-15.9495,-3.2217133435842293);
        ctx.lineTo(11,6.758);
        ctx.lineTo(6.757,2.515);
        ctx.translate(6.05,3.2222135462503534);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.7855491633997438,-2.356043490190049,1);
        ctx.rotate(0);
        ctx.translate(-6.05,-3.2222135462503534);
        ctx.lineTo(2.515,5.343);
        ctx.translate(3.2222135462503534,6.05);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.3563454901946406,-3.926839816984946,1);
        ctx.rotate(0);
        ctx.translate(-3.2222135462503534,-6.05);
        ctx.lineTo(6.757,11);
        ctx.lineTo(2.5149999999999997,15.243);
        ctx.translate(3.2222135462503525,15.95);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.3563454901946406,-3.926839816984946,1);
        ctx.rotate(0);
        ctx.translate(-3.2222135462503525,-15.95);
        ctx.lineTo(5.343,19.485);
        ctx.translate(6.05,18.777786453749645);
        ctx.rotate(0);
        ctx.arc(0,0,1,2.3560434901900495,0.785549163399744,1);
        ctx.rotate(0);
        ctx.translate(-6.05,-18.777786453749645);
        ctx.lineTo(11,15.243);
        ctx.lineTo(15.242,19.485);
        ctx.translate(15.9495,18.77828665641577);
        ctx.rotate(0);
        ctx.arc(0,0,1,2.3567507403070853,0.7848419132827078,1);
        ctx.rotate(0);
        ctx.translate(-15.9495,-18.77828665641577);
        ctx.lineTo(19.485,16.657);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(6.404,2.868);
        ctx.translate(6.0504999999999995,3.2216067731251763);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-0.7855491633997438,-2.356043490190049,1);
        ctx.rotate(0);
        ctx.translate(-6.0504999999999995,-3.2216067731251763);
        ctx.lineTo(2.868,5.697);
        ctx.translate(3.2216067731251763,6.0504999999999995);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-2.3563454901946406,-3.926839816984946,1);
        ctx.rotate(0);
        ctx.translate(-3.2216067731251763,-6.0504999999999995);
        ctx.lineTo(7.464,11);
        ctx.lineTo(2.8680000000000003,15.596);
        ctx.translate(3.2216067731251763,15.9495);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-2.3563454901946392,-3.9268398169849474,1);
        ctx.rotate(0);
        ctx.translate(-3.2216067731251763,-15.9495);
        ctx.lineTo(5.697000000000001,19.132);
        ctx.translate(6.050500000000001,18.778393226874826);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,2.3560434901900495,0.785549163399744,1);
        ctx.rotate(0);
        ctx.translate(-6.050500000000001,-18.778393226874826);
        ctx.lineTo(11,14.536);
        ctx.lineTo(15.596,19.131999999999998);
        ctx.translate(15.9495,18.778393226874822);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,2.356043490190051,0.7855491633997427,1);
        ctx.rotate(0);
        ctx.translate(-15.9495,-18.778393226874822);
        ctx.lineTo(19.132,16.302999999999997);
        ctx.translate(18.778393226874826,15.949499999999997);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,0.785247163395154,-0.7852471633951542,1);
        ctx.rotate(0);
        ctx.translate(-18.778393226874826,-15.949499999999997);
        ctx.lineTo(14.535,11);
        ctx.lineTo(19.132,6.404);
        ctx.translate(18.778393226874826,6.0504999999999995);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,0.7852471633951528,-0.7852471633951528,1);
        ctx.rotate(0);
        ctx.translate(-18.778393226874826,-6.0504999999999995);
        ctx.lineTo(16.303,2.868);
        ctx.translate(15.9495,3.221606773125176);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-0.7855491633997425,-2.356043490190051,1);
        ctx.rotate(0);
        ctx.translate(-15.9495,-3.221606773125176);
        ctx.lineTo(11,7.465);
        ctx.lineTo(6.404,2.868);
        ctx.closePath();
        ctx.moveTo(4.989,2.161);
        ctx.translate(6.05,3.2213202346461185);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,-2.3565149346587257,-0.7850777189310671,0);
        ctx.rotate(0);
        ctx.translate(-6.05,-3.2213202346461185);
        ctx.lineTo(11,6.051);
        ctx.lineTo(14.889,2.161);
        ctx.translate(15.949499999999999,3.2218203193755297);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,-2.356043490190049,-0.7855491633997442,0);
        ctx.rotate(0);
        ctx.translate(-15.949499999999999,-3.2218203193755297);
        ctx.lineTo(19.84,4.99);
        ctx.translate(18.77917968062447,6.0505);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,-0.7852471633951532,0.7852471633951532,0);
        ctx.rotate(0);
        ctx.translate(-18.77917968062447,-6.0505);
        ctx.lineTo(15.95,11.001000000000001);
        ctx.lineTo(19.84,14.889000000000001);
        ctx.translate(18.779679765353883,15.950000000000003);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,-0.78571860786383,0.78571860786383,0);
        ctx.rotate(0);
        ctx.translate(-18.779679765353883,-15.950000000000003);
        ctx.lineTo(17.011,19.839000000000002);
        ctx.translate(15.950499999999998,18.778179680624472);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,0.7855491633997433,2.35604349019005,0);
        ctx.rotate(0);
        ctx.translate(-15.950499999999998,-18.778179680624472);
        ctx.lineTo(11,15.95);
        ctx.lineTo(7.109999999999999,19.839);
        ctx.translate(6.049999999999999,18.77768006708627);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,0.7860203857802706,2.3555722678095226,0);
        ctx.rotate(0);
        ctx.translate(-6.049999999999999,-18.77768006708627);
        ctx.lineTo(2.16,17.011);
        ctx.translate(3.2203202346461186,15.95);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,2.355874045725964,3.9273112614536227,0);
        ctx.rotate(0);
        ctx.translate(-3.2203202346461186,-15.95);
        ctx.lineTo(6.051,11);
        ctx.lineTo(2.16,7.111);
        ctx.translate(3.2208203193755294,6.0504999999999995);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,2.3563454901946406,3.9268398169849466,0);
        ctx.rotate(0);
        ctx.translate(-3.2208203193755294,-6.0504999999999995);
        ctx.lineTo(4.989,2.16);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconCrossHairs(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(10.5,21);
        ctx.bezierCurveTo(16.299,21,21,16.299,21,10.5);
        ctx.bezierCurveTo(21,4.7010000000000005,16.299,0,10.5,0);
        ctx.bezierCurveTo(4.7010000000000005,0,0,4.701,0,10.5);
        ctx.bezierCurveTo(0,16.299,4.701,21,10.5,21);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(10.5,20);
        ctx.translate(10.5,10.5);
        ctx.rotate(0);
        ctx.arc(0,0,9.5,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-10.5,-10.5);
        ctx.translate(10.5,10.5);
        ctx.rotate(0);
        ctx.arc(0,0,9.5,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-10.5,-10.5);
        ctx.moveTo(21,10.5);
        ctx.bezierCurveTo(21,16.299,16.299,21,10.5,21);
        ctx.bezierCurveTo(4.7010000000000005,21,0,16.299,0,10.5);
        ctx.bezierCurveTo(0,4.7010000000000005,4.701,0,10.5,0);
        ctx.bezierCurveTo(16.299,0,21,4.701,21,10.5);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(11,6);
        ctx.translate(10.5,6);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,0,3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-10.5,-6);
        ctx.lineTo(10,10);
        ctx.lineTo(6,10);
        ctx.translate(6,10.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-6,-10.5);
        ctx.lineTo(10,11);
        ctx.lineTo(10,15);
        ctx.translate(10.5,15);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,3.141592653589793,6.283185307179586,1);
        ctx.rotate(0);
        ctx.translate(-10.5,-15);
        ctx.lineTo(11,11);
        ctx.lineTo(15,11);
        ctx.translate(15,10.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-15,-10.5);
        ctx.lineTo(11,10);
        ctx.lineTo(11,6);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconHelp(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x, y);
        ctx.rotate(rotationAngle);
        // ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(21,11);
        ctx.bezierCurveTo(21,16.523,16.523,21,11,21);
        ctx.bezierCurveTo(5.477,21,1,16.523,1,11);
        ctx.bezierCurveTo(1,5.477,5.477,1,11,1);
        ctx.bezierCurveTo(16.523,1,21,5.477,21,11);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(11,20);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,9,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.translate(11,11);
        ctx.rotate(0);
        ctx.arc(0,0,9,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-11,-11);
        ctx.moveTo(11,21);
        ctx.bezierCurveTo(16.523,21,21,16.523,21,11);
        ctx.bezierCurveTo(21,5.477,16.523,1,11,1);
        ctx.bezierCurveTo(5.477,1,1,5.477,1,11);
        ctx.bezierCurveTo(1,16.523,5.477,21,11,21);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(10.605,13.136);
        ctx.lineTo(10.605,13.075999999999999);
        ctx.bezierCurveTo(10.612,12.443999999999999,10.678,11.94,10.804,11.565999999999999);
        ctx.bezierCurveTo(10.93,11.190999999999999,11.109,10.886999999999999,11.341000000000001,10.655999999999999);
        ctx.bezierCurveTo(11.573,10.424,11.851,10.209999999999999,12.176000000000002,10.014);
        ctx.bezierCurveTo(12.371000000000002,9.894,12.547000000000002,9.754,12.703000000000001,9.591999999999999);
        ctx.translate(11.33200549848536,8.308544478052854);
        ctx.rotate(0);
        ctx.arc(0,0,1.878,0.7524319124568346,-0.0024198522300480363,1);
        ctx.rotate(0);
        ctx.translate(-11.33200549848536,-8.308544478052854);
        ctx.bezierCurveTo(13.21,7.978999999999998,13.134,7.696999999999998,12.981000000000002,7.458999999999999);
        ctx.translate(11.683486840435975,8.281043551618875);
        ctx.rotate(0);
        ctx.arc(0,0,1.536,-0.5647262666864303,-1.1074448898034213,1);
        ctx.rotate(0);
        ctx.translate(-11.683486840435975,-8.281043551618875);
        ctx.translate(11.543797192947725,8.567847350486774);
        ctx.rotate(0);
        ctx.arc(0,0,1.855,-1.1091838994037526,-1.5836253541441683,1);
        ctx.rotate(0);
        ctx.translate(-11.543797192947725,-8.567847350486774);
        ctx.bezierCurveTo(11.248000000000001,6.712999999999998,10.986,6.768999999999998,10.734000000000002,6.882999999999998);
        ctx.translate(11.360488251520538,8.267887168944366);
        ctx.rotate(0);
        ctx.arc(0,0,1.52,-1.995623499656061,-2.5450714665010894,1);
        ctx.rotate(0);
        ctx.translate(-11.360488251520538,-8.267887168944366);
        ctx.bezierCurveTo(9.933000000000002,7.655999999999998,9.836000000000002,7.972999999999998,9.809000000000001,8.363999999999997);
        ctx.lineTo(8.557,8.363999999999997);
        ctx.bezierCurveTo(8.583,7.799999999999997,8.729000000000001,7.317999999999997,8.994,6.916999999999997);
        ctx.translate(11.188788158959932,8.375125487497309);
        ctx.rotate(0);
        ctx.arc(0,0,2.635,-2.5551899298117338,-2.016019870935209,0);
        ctx.rotate(0);
        ctx.translate(-11.188788158959932,-8.375125487497309);
        ctx.bezierCurveTo(10.494,5.7849999999999975,10.983,5.678999999999998,11.52,5.678999999999998);
        ctx.bezierCurveTo(12.103,5.678999999999998,12.61,5.794999999999997,13.041,6.0269999999999975);
        ctx.bezierCurveTo(13.475,6.258999999999998,13.81,6.576999999999997,14.045,6.9819999999999975);
        ctx.bezierCurveTo(14.284,7.3859999999999975,14.403,7.846999999999998,14.403,8.363999999999997);
        ctx.bezierCurveTo(14.403,8.727999999999998,14.347000000000001,9.057999999999996,14.234,9.352999999999998);
        ctx.bezierCurveTo(14.124,9.647999999999998,13.966,9.910999999999998,13.757,10.142999999999997);
        ctx.translate(11.322229213793454,7.958662064003106);
        ctx.rotate(0);
        ctx.arc(0,0,3.271,0.7312344883785125,1.0282880066826148,0);
        ctx.rotate(0);
        ctx.translate(-11.322229213793454,-7.958662064003106);
        ctx.translate(14.603834133431905,13.405491149744073);
        ctx.rotate(0);
        ctx.arc(0,0,3.088,-2.112754051016296,-2.407588370032469,1);
        ctx.rotate(0);
        ctx.translate(-14.603834133431905,-13.405491149744073);
        ctx.translate(13.680004083858686,12.536261363664416);
        ctx.rotate(0);
        ctx.arc(0,0,1.82,-2.4221908650664257,-2.8694198138542006,1);
        ctx.rotate(0);
        ctx.translate(-13.680004083858686,-12.536261363664416);
        ctx.bezierCurveTo(11.847,12.322999999999997,11.805,12.665999999999997,11.798,13.076999999999996);
        ctx.lineTo(11.798,13.136999999999997);
        ctx.lineTo(10.605,13.136999999999997);
        ctx.closePath();
        ctx.moveTo(11.241,16.08);
        ctx.translate(11.230474667623017,15.219064336098013);
        ctx.rotate(0);
        ctx.arc(0,0,0.861,1.5585714783109954,2.3755273892032633,0);
        ctx.rotate(0);
        ctx.translate(-11.230474667623017,-15.219064336098013);
        ctx.translate(11.206935663901985,15.195525332376981);
        ctx.rotate(0);
        ctx.arc(0,0,0.861,2.3368615911814263,3.1538175020736943,0);
        ctx.rotate(0);
        ctx.translate(-11.206935663901985,-15.195525332376981);
        ctx.bezierCurveTo(10.346,14.938999999999998,10.434,14.729,10.61,14.552999999999999);
        ctx.translate(11.229522790210272,15.15092350046647);
        ctx.rotate(0);
        ctx.arc(0,0,0.861,-2.3739341014924373,-1.5574658393185912,0);
        ctx.rotate(0);
        ctx.translate(-11.229522790210272,-15.15092350046647);
        ctx.bezierCurveTo(11.487,14.29,11.697,14.377999999999998,11.873,14.552999999999999);
        ctx.translate(11.27509021988388,15.17253603191541);
        ctx.rotate(0);
        ctx.arc(0,0,0.861,-0.8031599211043225,0.01447665905744866,0);
        ctx.rotate(0);
        ctx.translate(-11.27509021988388,-15.17253603191541);
        ctx.translate(11.208169126214317,15.16728363327036);
        ctx.rotate(0);
        ctx.arc(0,0,0.928,0.019092072269415172,1.0261253363031761,0);
        ctx.rotate(0);
        ctx.translate(-11.208169126214317,-15.16728363327036);
        ctx.translate(11.256165015506811,15.22713497412754);
        ctx.rotate(0);
        ctx.arc(0,0,0.852,1.037909179046562,1.588596580754035,0);
        ctx.rotate(0);
        ctx.translate(-11.256165015506811,-15.22713497412754);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconInsert(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(9.65,1.7);
        ctx.translate(10.544460611816056,2.1471467476227972);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.678019781391401,-0.5042452305497682,0);
        ctx.rotate(0);
        ctx.translate(-10.544460611816056,-2.1471467476227972);
        ctx.lineTo(20.18,17.516000000000002);
        ctx.translate(19.304932003920698,18.000000002311737);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.5052200403546818,1.5707283307155433,0);
        ctx.rotate(0);
        ctx.translate(-19.304932003920698,-18.000000002311737);
        ctx.lineTo(2.618,19);
        ctx.translate(2.618533885657676,18.000000142516956);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.571330212477935,3.605001629134821,0);
        ctx.rotate(0);
        ctx.translate(-2.618533885657676,-18.000000142516956);
        ctx.lineTo(9.65,1.7);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(19.305,18);
        ctx.lineTo(10.545,2.1470000000000002);
        ctx.lineTo(2.617,18);
        ctx.lineTo(19.304000000000002,18);
        ctx.closePath();
        ctx.moveTo(11.419,1.664);
        ctx.translate(10.543460611816057,2.1471467476227972);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.5042452305497682,-2.678019781391401,1);
        ctx.rotate(0);
        ctx.translate(-10.543460611816057,-2.1471467476227972);
        ctx.lineTo(1.725,17.553);
        ctx.translate(2.619533369417748,18.000001175611576);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.6781825231473007,-4.710855610366059,1);
        ctx.rotate(0);
        ctx.translate(-2.619533369417748,-18.000001175611576);
        ctx.lineTo(19.305,19);
        ctx.translate(19.3049320039207,18.000000002311733);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.570728330715545,-0.5052200403546858,1);
        ctx.rotate(0);
        ctx.translate(-19.3049320039207,-18.000000002311733);
        ctx.lineTo(11.42,1.664);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconKey(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(0,0);
        ctx.lineTo(22,0);
        ctx.lineTo(22,22);
        ctx.lineTo(0,22);
        ctx.closePath();
        ctx.clip();
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(1.089,15.914);
        ctx.translate(2.502999954388261,17.328427138098668);
        ctx.rotate(0);
        ctx.arc(0,0,2,-2.3560434579425493,-3.1413790845388343,1);
        ctx.rotate(0);
        ctx.translate(-2.502999954388261,-17.328427138098668);
        ctx.lineTo(0.503,19.5);
        ctx.translate(1.5030000000000001,19.5);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-1.5030000000000001,-19.5);
        ctx.lineTo(4.503,20.5);
        ctx.translate(4.503,19.5);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-4.503,-19.5);
        ctx.lineTo(5.503,18.5);
        ctx.translate(6.503,18.5);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-6.503,-18.5);
        ctx.lineTo(7.503,17.5);
        ctx.translate(7.503,16.5);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-7.503,-16.5);
        ctx.lineTo(8.503,15.5);
        ctx.translate(9.503,15.5);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-9.503,-15.5);
        ctx.lineTo(9.675,14.5);
        ctx.translate(9.67457286190133,12.50000004561174);
        ctx.rotate(0);
        ctx.arc(0,0,2,1.5705827577439373,0.7852471311476521,1);
        ctx.rotate(0);
        ctx.translate(-9.67457286190133,-12.50000004561174);
        ctx.lineTo(11.903,13.1);
        ctx.translate(14.041236339311713,6.961763660688288);
        ctx.rotate(0);
        ctx.arc(0,0,6.5,1.9059978002596143,2.8063911801250754,1);
        ctx.rotate(0);
        ctx.translate(-14.041236339311713,-6.961763660688288);
        ctx.lineTo(1.0890000000000004,15.914);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(10.321,1.032);
        ctx.translate(14.0409196950033,6.96177212569982);
        ctx.rotate(0);
        ctx.arc(0,0,7,-2.1310688788215986,-4.422649250089384,0);
        ctx.rotate(0);
        ctx.translate(-14.0409196950033,-6.96177212569982);
        ctx.lineTo(11.443,14.267);
        ctx.translate(9.674466369765835,12.500000056952246);
        ctx.rotate(0);
        ctx.arc(0,0,2.5,0.784964371139951,1.570582874699609,0);
        ctx.rotate(0);
        ctx.translate(-9.674466369765835,-12.500000056952246);
        ctx.lineTo(9.503,15);
        ctx.translate(9.503,15.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-9.503,-15.5);
        ctx.lineTo(9.003,16.5);
        ctx.translate(7.503,16.5);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,0,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-7.503,-16.5);
        ctx.lineTo(6.503,18);
        ctx.translate(6.503,18.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-6.503,-18.5);
        ctx.lineTo(6.003,19.5);
        ctx.translate(4.503,19.5);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,0,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-4.503,-19.5);
        ctx.lineTo(1.5030000000000001,21);
        ctx.translate(1.5030000000000001,19.5);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-1.5030000000000001,-19.5);
        ctx.lineTo(0.0030000000000001137,17.328);
        ctx.translate(2.5029999565536376,17.327533918663608);
        ctx.rotate(0);
        ctx.arc(0,0,2.5,3.141406221054554,3.9268590015663256,0);
        ctx.rotate(0);
        ctx.translate(-2.5029999565536376,-17.327533918663608);
        ctx.lineTo(7.333,8.962999999999997);
        ctx.translate(14.040850650657916,6.961815438679338);
        ctx.rotate(0);
        ctx.arc(0,0,7,2.8516643635284926,4.152128071991463,0);
        ctx.rotate(0);
        ctx.translate(-14.040850650657916,-6.961815438679338);
        ctx.closePath();
        ctx.moveTo(14.715,1);
        ctx.translate(14.041349763408187,6.962063011973271);
        ctx.rotate(0);
        ctx.arc(0,0,6,-1.4582840552168919,-3.47664947422601,1);
        ctx.rotate(0);
        ctx.translate(-14.041349763408187,-6.962063011973271);
        ctx.translate(7.9027210323318595,9.09917239931948);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-0.3345506846354563,0.7861740477221655,0);
        ctx.rotate(0);
        ctx.translate(-7.9027210323318595,-9.09917239931948);
        ctx.lineTo(1.4420000000000002,16.267);
        ctx.translate(2.50199941874797,17.328320513442552);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,-2.355571720140669,-3.1407123111810455,1);
        ctx.rotate(0);
        ctx.translate(-2.50199941874797,-17.328320513442552);
        ctx.lineTo(1.0020000000000002,19.5);
        ctx.translate(1.5020000000000002,19.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-1.5020000000000002,-19.5);
        ctx.lineTo(4.502000000000001,20);
        ctx.translate(4.502000000000001,19.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-4.502000000000001,-19.5);
        ctx.lineTo(5.002000000000001,18.5);
        ctx.translate(6.502000000000001,18.5);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-6.502000000000001,-18.5);
        ctx.lineTo(7.502000000000001,17);
        ctx.translate(7.502000000000001,16.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-7.502000000000001,-16.5);
        ctx.lineTo(8.002,15.5);
        ctx.translate(9.502,15.5);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-9.502,-15.5);
        ctx.lineTo(9.675,14);
        ctx.translate(9.673679486557448,12.500000581252031);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,1.5699159843861503,0.7847753933457732,1);
        ctx.rotate(0);
        ctx.translate(-9.673679486557448,-12.500000581252031);
        ctx.lineTo(11.549000000000001,12.746);
        ctx.translate(11.902085734722661,13.10002042869778);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-2.3548726329111163,-1.2346743909056859,0);
        ctx.rotate(0);
        ctx.translate(-11.902085734722661,-13.10002042869778);
        ctx.translate(14.041012785170736,6.962024927341942);
        ctx.rotate(0);
        ctx.arc(0,0,6,1.906043010880377,4.824957772551787,1);
        ctx.rotate(0);
        ctx.translate(-14.041012785170736,-6.962024927341942);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(16,6);
        ctx.translate(16,5);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-16,-5);
        ctx.translate(16,5);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-16,-5);
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.restore();

        ctx.restore();
    }
    function drawIconNewParagraph(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(2,7.49);
        ctx.translate(2.9999997593287437,7.489306211520121);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.140898865054255,4.051270879791058,0);
        ctx.rotate(0);
        ctx.translate(-2.9999997593287437,-7.489306211520121);
        ctx.lineTo(10.386,0.4770000000000003);
        ctx.translate(11,1.2663060243023618);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.2319147323035042,-0.9096779212862891,0);
        ctx.rotate(0);
        ctx.translate(-11,-1.2663060243023618);
        ctx.lineTo(19.613999999999997,6.7);
        ctx.translate(19.000000240671252,7.4893062115201205);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.9096782262012644,0.0006937885355389195,0);
        ctx.rotate(0);
        ctx.translate(-19.000000240671252,-7.4893062115201205);
        ctx.lineTo(19.999999999999996,20);
        ctx.translate(18.999999999999996,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,0,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-18.999999999999996,-20);
        ctx.lineTo(3,21);
        ctx.translate(3,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-3,-20);
        ctx.lineTo(2,7.49);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(19,20);
        ctx.lineTo(19,7.49);
        ctx.lineTo(11,1.2670000000000003);
        ctx.lineTo(3,7.489000000000001);
        ctx.lineTo(3,20);
        ctx.lineTo(19,20);
        ctx.closePath();
        ctx.moveTo(2.386,6.7);
        ctx.translate(2.9999997593287437,7.489306211520121);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.231914427388528,-3.1422864421253314,1);
        ctx.rotate(0);
        ctx.translate(-2.9999997593287437,-7.489306211520121);
        ctx.lineTo(2,20);
        ctx.translate(3,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-3,-20);
        ctx.lineTo(19,21);
        ctx.translate(19,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-19,-20);
        ctx.lineTo(20,7.49);
        ctx.translate(19.00000024067126,7.4893062115201205);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.0006937885355384361,-0.9096782262012648,1);
        ctx.rotate(0);
        ctx.translate(-19.00000024067126,-7.4893062115201205);
        ctx.lineTo(11.614,0.4770000000000003);
        ctx.translate(11,1.2663060243023618);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.9096779212862891,-2.2319147323035042,1);
        ctx.rotate(0);
        ctx.translate(-11,-1.2663060243023618);
        ctx.lineTo(2.386000000000001,6.7);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(9.5,9);
        ctx.translate(9.5,11);
        ctx.rotate(0);
        ctx.arc(0,0,2,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-9.5,-11);
        ctx.lineTo(11,13);
        ctx.lineTo(11,9);
        ctx.lineTo(9.5,9);
        ctx.closePath();
        ctx.moveTo(9.5,8);
        ctx.translate(9.5,11);
        ctx.rotate(0);
        ctx.arc(0,0,3,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-9.5,-11);
        ctx.lineTo(11,14);
        ctx.lineTo(11,17.5);
        ctx.translate(11.5,17.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,3.141592653589793,6.283185307179586,1);
        ctx.rotate(0);
        ctx.translate(-11.5,-17.5);
        ctx.lineTo(12,9);
        ctx.lineTo(14,9);
        ctx.lineTo(14,17.5);
        ctx.translate(14.5,17.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,3.141592653589793,6.283185307179586,1);
        ctx.rotate(0);
        ctx.translate(-14.5,-17.5);
        ctx.lineTo(15,9);
        ctx.lineTo(15.5,9);
        ctx.translate(15.5,8.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-15.5,-8.5);
        ctx.lineTo(9.5,8);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconNote(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);

        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(2,3);
        ctx.translate(3,3);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-3,-3);
        ctx.lineTo(19,2);
        ctx.translate(19,3);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,0,0);
        ctx.rotate(0);
        ctx.translate(-19,-3);
        ctx.lineTo(20,13.586);
        ctx.translate(19.00000002280587,13.585786430950666);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.00021356905077231963,0.785549195647058,0);
        ctx.rotate(0);
        ctx.translate(-19.00000002280587,-13.585786430950666);
        ctx.lineTo(14.293000000000001,19.707);
        ctx.translate(13.585786430950666,19.00000002280587);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7852471311476532,1.570582757743939,0);
        ctx.rotate(0);
        ctx.translate(-13.585786430950666,-19.00000002280587);
        ctx.lineTo(3,20);
        ctx.translate(3,19);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-3,-19);
        ctx.lineTo(2,3);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(3,19);
        ctx.lineTo(3,3);
        ctx.lineTo(19,3);
        ctx.lineTo(19,13);
        ctx.lineTo(14.5,13);
        ctx.translate(14.5,14.5);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-14.5,-14.5);
        ctx.lineTo(13,19);
        ctx.lineTo(3,19);
        ctx.closePath();
        ctx.moveTo(14,18.586);
        ctx.lineTo(18.586,14);
        ctx.lineTo(14.5,14);
        ctx.translate(14.5,14.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-14.5,-14.5);
        ctx.lineTo(14,18.586);
        ctx.closePath();
        ctx.moveTo(2,3);
        ctx.translate(3,3);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-3,-3);
        ctx.lineTo(19,2);
        ctx.translate(19,3);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,0,0);
        ctx.rotate(0);
        ctx.translate(-19,-3);
        ctx.lineTo(20,13.586);
        ctx.translate(19.00000002280587,13.585786430950666);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.00021356905077231963,0.785549195647058,0);
        ctx.rotate(0);
        ctx.translate(-19.00000002280587,-13.585786430950666);
        ctx.lineTo(14.293000000000001,19.707);
        ctx.translate(13.585786430950666,19.00000002280587);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7852471311476532,1.570582757743939,0);
        ctx.rotate(0);
        ctx.translate(-13.585786430950666,-19.00000002280587);
        ctx.lineTo(3,20);
        ctx.translate(3,19);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-3,-19);
        ctx.lineTo(2,3);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconParagraph(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(8,12.5);
        ctx.lineTo(12.5,12.5);
        ctx.lineTo(12.5,20.5);
        ctx.lineTo(17.5,20.5);
        ctx.lineTo(17.5,2.5);
        ctx.lineTo(8,2.5);
        ctx.translate(8,7.5);
        ctx.rotate(0);
        ctx.arc(0,0,5,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-8,-7.5);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(8,2);
        ctx.translate(8,7.5);
        ctx.rotate(0);
        ctx.arc(0,0,5.5,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-8,-7.5);
        ctx.lineTo(12,13);
        ctx.lineTo(12,20.5);
        ctx.translate(12.5,20.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-12.5,-20.5);
        ctx.lineTo(17.5,21);
        ctx.translate(17.5,20.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-17.5,-20.5);
        ctx.lineTo(18,3);
        ctx.lineTo(19.5,3);
        ctx.translate(19.5,2.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.5707963267948966,4.71238898038469,1);
        ctx.rotate(0);
        ctx.translate(-19.5,-2.5);
        ctx.lineTo(8,2);
        ctx.closePath();
        ctx.moveTo(8,3);
        ctx.translate(8,7.5);
        ctx.rotate(0);
        ctx.arc(0,0,4.5,-1.5707963267948966,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-8,-7.5);
        ctx.lineTo(12,12);
        ctx.lineTo(12,3);
        ctx.lineTo(8,3);
        ctx.closePath();
        ctx.moveTo(13,3);
        ctx.lineTo(13,20);
        ctx.lineTo(17,20);
        ctx.lineTo(17,3);
        ctx.lineTo(13,3);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconRightArrow(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(8,16);
        ctx.translate(7,16);
        ctx.rotate(0);
        ctx.arc(0,0,1,0,-1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-7,-16);
        ctx.lineTo(2,15);
        ctx.translate(2,14);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-2,-14);
        ctx.lineTo(1,8);
        ctx.translate(2,8);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-2,-8);
        ctx.lineTo(7,7);
        ctx.translate(7,6);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-7,-6);
        ctx.lineTo(8,2.887);
        ctx.translate(8.999999627596988,2.8861369786310354);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.140729632113701,5.309191228420847,0);
        ctx.rotate(0);
        ctx.translate(-8.999999627596988,-2.8861369786310354);
        ctx.lineTo(20.84,9.713);
        ctx.translate(20.2786040403633,10.540547325839185);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.9747246262568916,0.9280500548604216,0);
        ctx.rotate(0);
        ctx.translate(-20.2786040403633,-10.540547325839185);
        ctx.lineTo(9.6,19.8);
        ctx.translate(9.000000000000002,19);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.9272952180016126,3.1415926535897936,0);
        ctx.rotate(0);
        ctx.translate(-9.000000000000002,-19);
        ctx.lineTo(8,16);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(2,14);
        ctx.lineTo(7,14);
        ctx.translate(7,16);
        ctx.rotate(0);
        ctx.arc(0,0,2,-1.5707963267948966,0,0);
        ctx.rotate(0);
        ctx.translate(-7,-16);
        ctx.lineTo(9,19);
        ctx.lineTo(20.28,10.54);
        ctx.lineTo(9,2.888);
        ctx.lineTo(9,6);
        ctx.translate(7,6);
        ctx.rotate(0);
        ctx.arc(0,0,2,0,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-7,-6);
        ctx.lineTo(2,8);
        ctx.lineTo(2,14);
        ctx.closePath();
        ctx.moveTo(7,15);
        ctx.translate(7,16);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,0,0);
        ctx.rotate(0);
        ctx.translate(-7,-16);
        ctx.lineTo(8,19);
        ctx.translate(9.000000000000002,19);
        ctx.rotate(0);
        ctx.arc(0,0,1,-3.141592653589793,-5.355890089177974,1);
        ctx.rotate(0);
        ctx.translate(-9.000000000000002,-19);
        ctx.lineTo(20.88,11.34);
        ctx.translate(20.27895466038126,10.540785072885523);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.9259879022029239,-0.9751482511590429,1);
        ctx.rotate(0);
        ctx.translate(-20.27895466038126,-10.540785072885523);
        ctx.lineTo(9.563,2.06);
        ctx.translate(8.999999852464112,2.8864567949231024);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.9727848639841764,-3.142135858693405,1);
        ctx.rotate(0);
        ctx.translate(-8.999999852464112,-2.8864567949231024);
        ctx.lineTo(8,6);
        ctx.translate(7,6);
        ctx.rotate(0);
        ctx.arc(0,0,1,0,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-7,-6);
        ctx.lineTo(2,7);
        ctx.translate(2,8);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-2,-8);
        ctx.lineTo(1,14);
        ctx.translate(2,14);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-2,-14);
        ctx.lineTo(7,15);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconRightPointer(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(1.991,3.008);
        ctx.bezierCurveTo(1.431,2.1630000000000003,2.341,1.118,3.2560000000000002,1.554);
        ctx.lineTo(20.109,9.597);
        ctx.translate(19.678310146393024,10.4995);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.1255393084200798,1.1255393084200798,0);
        ctx.rotate(0);
        ctx.translate(-19.678310146393024,-10.4995);
        ctx.lineTo(3.256,19.446);
        ctx.bezierCurveTo(2.3409999999999997,19.883000000000003,1.4319999999999997,18.837000000000003,1.9909999999999999,17.991);
        ctx.lineTo(6.585,11.050999999999998);
        ctx.translate(5.7508251082656585,10.499499999999998);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.5841613552973065,-0.5841613552973067,1);
        ctx.rotate(0);
        ctx.translate(-5.7508251082656585,-10.499499999999998);
        ctx.lineTo(1.99,3.008);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(1.574,3.284);
        ctx.bezierCurveTo(0.7340000000000001,2.016,2.098,0.44799999999999995,3.471,1.1029999999999998);
        ctx.lineTo(20.324,9.145999999999999);
        ctx.bezierCurveTo(21.463,9.69,21.463,11.309999999999999,20.324,11.854);
        ctx.lineTo(3.471,19.897);
        ctx.bezierCurveTo(2.098,20.552,0.7349999999999999,18.983999999999998,1.574,17.715999999999998);
        ctx.lineTo(6.168,10.775999999999996);
        ctx.translate(5.751077944934548,10.499999999999996);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,0.5847608688333814,-0.5847608688333814,1);
        ctx.rotate(0);
        ctx.translate(-5.751077944934548,-10.499999999999996);
        ctx.lineTo(1.5739999999999998,3.2839999999999963);
        ctx.closePath();
        ctx.moveTo(3.04,2.005);
        ctx.bezierCurveTo(2.583,1.787,2.128,2.31,2.408,2.7329999999999997);
        ctx.lineTo(7.002000000000001,9.673);
        ctx.translate(5.750902981379942,10.500499999999999);
        ctx.rotate(0);
        ctx.arc(0,0,1.5,-0.584361166726106,0.5843611667261063,0);
        ctx.rotate(0);
        ctx.translate(-5.750902981379942,-10.500499999999999);
        ctx.lineTo(2.4080000000000004,18.268);
        ctx.bezierCurveTo(2.128,18.69,2.583,19.213,3.0400000000000005,18.995);
        ctx.lineTo(19.894,10.951);
        ctx.translate(19.678131984768466,10.5);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.1243797886208347,-1.1243797886208347,1);
        ctx.rotate(0);
        ctx.translate(-19.678131984768466,-10.5);
        ctx.lineTo(3.04,2.005);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconStar(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(13.884,8.442);
        ctx.translate(13.951116865547602,7.94652515062915);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.7054364814756378,2.7550918035124354,0);
        ctx.rotate(0);
        ctx.translate(-13.951116865547602,-7.94652515062915);
        ctx.lineTo(11.463,3.143);
        ctx.translate(11,3.331761754600872);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-0.38712043758627585,-2.7544722160035175,1);
        ctx.rotate(0);
        ctx.translate(-11,-3.331761754600872);
        ctx.lineTo(8.511,8.135);
        ctx.translate(8.047883134452396,7.94652515062915);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,0.38650085007735785,1.4361561721141556,0);
        ctx.rotate(0);
        ctx.translate(-8.047883134452396,-7.94652515062915);
        ctx.lineTo(2.238999999999998,9.236);
        ctx.translate(2.3067508761404114,9.731388553341926);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.7067161947127014,-4.084420830362329,1);
        ctx.rotate(0);
        ctx.translate(-2.3067508761404114,-9.731388553341926);
        ctx.lineTo(6.494999999999998,13.378);
        ctx.translate(6.201964228953895,13.783129654416232);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-0.944594432449701,0.3360288912303172,0);
        ctx.rotate(0);
        ctx.translate(-6.201964228953895,-13.783129654416232);
        ctx.lineTo(4.667999999999999,19.668);
        ctx.translate(5.1403686039109004,19.831914313100675);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-2.8075883869416365,-5.263064248304401,1);
        ctx.rotate(0);
        ctx.translate(-5.1403686039109004,-19.831914313100675);
        ctx.lineTo(10.735999999999999,16.968);
        ctx.translate(10.9985,17.39355111326373);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-2.12351143989168,-1.0180812136981128,0);
        ctx.rotate(0);
        ctx.translate(-10.9985,-17.39355111326373);
        ctx.lineTo(16.596,20.258);
        ctx.translate(16.857631396089097,19.831914313100675);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,2.1214715947146026,-0.3340042666481535,1);
        ctx.rotate(0);
        ctx.translate(-16.857631396089097,-19.831914313100675);
        ctx.lineTo(15.323999999999998,13.948);
        ctx.translate(15.796035771046101,13.783129654416232);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,2.8055637623594762,4.086187086039494,0);
        ctx.rotate(0);
        ctx.translate(-15.796035771046101,-13.783129654416232);
        ctx.lineTo(19.985,10.136);
        ctx.translate(19.691249123859585,9.731388553341926);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,0.9428281767725357,-1.4348764588770901,1);
        ctx.rotate(0);
        ctx.translate(-19.691249123859585,-9.731388553341926);
        ctx.lineTo(13.883,8.441999999999998);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(14.305,7.49);
        ctx.lineTo(12.39,2.767);
        ctx.bezierCurveTo(11.884,1.5179999999999998,10.116,1.5179999999999998,9.610000000000001,2.767);
        ctx.lineTo(7.694,7.49);
        ctx.lineTo(2.106,8.245000000000001);
        ctx.bezierCurveTo(0.7639999999999998,8.426,0.32999999999999985,10.153,1.428,10.947000000000001);
        ctx.lineTo(5.608,13.97);
        ctx.lineTo(3.7249999999999996,19.337);
        ctx.bezierCurveTo(3.2569999999999997,20.671,4.725,21.852,5.927999999999999,21.11);
        ctx.lineTo(11,17.982);
        ctx.lineTo(16.072,21.11);
        ctx.bezierCurveTo(17.275,21.852,18.741999999999997,20.669999999999998,18.275,19.337);
        ctx.lineTo(16.392,13.969999999999999);
        ctx.lineTo(20.572,10.947);
        ctx.bezierCurveTo(21.669999999999998,10.152999999999999,21.236,8.427,19.894,8.245);
        ctx.lineTo(14.304999999999998,7.489999999999999);
        ctx.closePath();
        ctx.moveTo(13.488,8.135);
        ctx.translate(13.951116865547602,7.94652515062915);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,2.7550918035124354,1.7054364814756378,1);
        ctx.rotate(0);
        ctx.translate(-13.951116865547602,-7.94652515062915);
        ctx.lineTo(19.76,9.236);
        ctx.translate(19.692249123859586,9.731388553341926);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.4348764588770904,0.9428281767725355,0);
        ctx.rotate(0);
        ctx.translate(-19.692249123859586,-9.731388553341926);
        ctx.lineTo(15.504000000000001,13.378);
        ctx.translate(15.797035771046103,13.783129654416232);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-2.1969982211400922,-3.4776215448201104,1);
        ctx.rotate(0);
        ctx.translate(-15.797035771046103,-13.783129654416232);
        ctx.lineTo(17.331,19.668);
        ctx.translate(16.858631396089095,19.831914313100675);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-0.33400426664815364,2.1214715947146026,0);
        ctx.rotate(0);
        ctx.translate(-16.858631396089095,-19.831914313100675);
        ctx.lineTo(11.262,16.968);
        ctx.translate(10.999500000000001,17.39355111326373);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-1.018081213698113,-2.12351143989168,1);
        ctx.rotate(0);
        ctx.translate(-10.999500000000001,-17.39355111326373);
        ctx.lineTo(5.4030000000000005,20.258);
        ctx.translate(5.141368603910902,19.831914313100675);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.0201210588751854,3.47559692023795,0);
        ctx.rotate(0);
        ctx.translate(-5.141368603910902,-19.831914313100675);
        ctx.lineTo(6.675000000000001,13.948);
        ctx.translate(6.202964228953897,13.783129654416232);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,0.3360288912303172,-0.944594432449701,1);
        ctx.rotate(0);
        ctx.translate(-6.202964228953897,-13.783129654416232);
        ctx.lineTo(2.0140000000000002,10.136);
        ctx.translate(2.3077508761404135,9.731388553341926);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,2.198764476817258,4.576469112466885,0);
        ctx.rotate(0);
        ctx.translate(-2.3077508761404135,-9.731388553341926);
        ctx.lineTo(8.116,8.441999999999998);
        ctx.translate(8.048883134452398,7.94652515062915);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,1.4361561721141556,0.3865008500773579,1);
        ctx.rotate(0);
        ctx.translate(-8.048883134452398,-7.94652515062915);
        ctx.lineTo(10.536000000000001,3.142999999999998);
        ctx.translate(10.999500000000001,3.3305306641592236);
        ctx.rotate(0);
        ctx.arc(0,0,0.5,-2.7571297221386826,-0.38446293145111055,0);
        ctx.rotate(0);
        ctx.translate(-10.999500000000001,-3.3305306641592236);
        ctx.lineTo(13.488000000000001,8.134999999999998);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconUpArrow(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(16,13);
        ctx.translate(16,14);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,-3.141592653589793,1);
        ctx.rotate(0);
        ctx.translate(-16,-14);
        ctx.lineTo(15,20);
        ctx.translate(14,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,0,1.5707963267948966,0);
        ctx.rotate(0);
        ctx.translate(-14,-20);
        ctx.lineTo(8,21);
        ctx.translate(8,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,3.141592653589793,0);
        ctx.rotate(0);
        ctx.translate(-8,-20);
        ctx.lineTo(7,14);
        ctx.translate(6,14);
        ctx.rotate(0);
        ctx.arc(0,0,1,0,-1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-6,-14);
        ctx.lineTo(3.08,13);
        ctx.translate(3.080624593627877,12.000000195058618);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5714209204633842,3.816724436401862,0);
        ctx.rotate(0);
        ctx.translate(-3.080624593627877,-12.000000195058618);
        ctx.lineTo(10.219,1.4760000000000009);
        ctx.translate(11,2.100531024049247);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.4670617462982185,-0.674530907291575,0);
        ctx.rotate(0);
        ctx.translate(-11,-2.100531024049247);
        ctx.lineTo(19.701,11.376000000000001);
        ctx.translate(18.919576651974957,12.000001242916497);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.6738527541993078,1.572372979423071,0);
        ctx.rotate(0);
        ctx.translate(-18.919576651974957,-12.000001242916497);
        ctx.lineTo(16,13);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(14,20);
        ctx.lineTo(14,14);
        ctx.translate(16,14);
        ctx.rotate(0);
        ctx.arc(0,0,2,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-16,-14);
        ctx.lineTo(18.92,12);
        ctx.lineTo(11,2.1);
        ctx.lineTo(3.08,12);
        ctx.lineTo(6,12);
        ctx.translate(6,14);
        ctx.rotate(0);
        ctx.arc(0,0,2,-1.5707963267948966,0,0);
        ctx.rotate(0);
        ctx.translate(-6,-14);
        ctx.lineTo(8,20);
        ctx.lineTo(14,20);
        ctx.closePath();
        ctx.moveTo(15,14);
        ctx.translate(16,14);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,4.71238898038469,0);
        ctx.rotate(0);
        ctx.translate(-16,-14);
        ctx.lineTo(18.92,13);
        ctx.translate(18.919375406372126,12.000000195058618);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.57017173312641,-0.6751317828120689,1);
        ctx.rotate(0);
        ctx.translate(-18.919375406372126,-12.000000195058618);
        ctx.lineTo(11.780000000000003,1.4760000000000009);
        ctx.translate(10.999500000000003,2.1011557805859282);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.6753311077441445,-2.466261545845649,1);
        ctx.rotate(0);
        ctx.translate(-10.999500000000003,-2.1011557805859282);
        ctx.lineTo(2.299000000000003,11.376000000000001);
        ctx.translate(3.080424268677108,12.00000009000196);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.4677413747928547,-4.711964711694854,1);
        ctx.rotate(0);
        ctx.translate(-3.080424268677108,-12.00000009000196);
        ctx.lineTo(6,13);
        ctx.translate(6,14);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.5707963267948966,0,0);
        ctx.rotate(0);
        ctx.translate(-6,-14);
        ctx.lineTo(7,20);
        ctx.translate(8,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.141592653589793,1.5707963267948966,1);
        ctx.rotate(0);
        ctx.translate(-8,-20);
        ctx.lineTo(14,21);
        ctx.translate(14,20);
        ctx.rotate(0);
        ctx.arc(0,0,1,1.5707963267948966,0,1);
        ctx.rotate(0);
        ctx.translate(-14,-20);
        ctx.lineTo(15,14);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
    function drawIconUpLeftArrow(ctx, x, y, xScale, yScale, rotationAngle) {
        ctx.save();
        ctx.translate(x + HALF_SIZE * xScale, y + HALF_SIZE * yScale);
        ctx.rotate(rotationAngle);
        ctx.translate(-HALF_SIZE * xScale, -HALF_SIZE * yScale);
        
        ctx.scale(xScale, yScale);

        ctx.strokeStyle="rgba(0,0,0,0)";
        ctx.miterLimit=4;
        ctx.font="15px ''";
        ctx.fillStyle="rgba(0,0,0,0)";
        ctx.font="   15px ''";
        ctx.save();
        ctx.fillStyle=ctx.iconFill;
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(14.297,7.383);
        ctx.translate(15.078330716794145,8.00711722536231);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.4675914659956075,-3.926674074045537,1);
        ctx.rotate(0);
        ctx.translate(-15.078330716794145,-8.00711722536231);
        ctx.lineTo(18.688000000000002,13.031);
        ctx.translate(17.981286656415776,13.738500000000002);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.7859544135121902,0.7859544135121902,0);
        ctx.rotate(0);
        ctx.translate(-17.981286656415776,-13.738500000000002);
        ctx.lineTo(14.445000000000002,18.688000000000002);
        ctx.translate(13.738000000000003,17.980786453749648);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7855491633997438,2.356043490190049,0);
        ctx.rotate(0);
        ctx.translate(-13.738000000000003,-17.980786453749648);
        ctx.lineTo(8.714000000000002,14.371000000000002);
        ctx.translate(8.007117225362311,15.078330716794149);
        ctx.rotate(0);
        ctx.arc(0,0,1,-0.7857149063391532,-2.2447975143890813,1);
        ctx.rotate(0);
        ctx.translate(-8.007117225362311,-15.078330716794149);
        ctx.lineTo(5.382000000000003,15.896000000000003);
        ctx.translate(4.757390644027872,15.115062645001485);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.8961651185244527,3.010278237670715,0);
        ctx.rotate(0);
        ctx.translate(-4.757390644027872,-15.115062645001485);
        ctx.lineTo(2.173,3.297);
        ctx.translate(3.164288029149254,3.165288029149254);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.009496854561797,4.844484779412688,0);
        ctx.rotate(0);
        ctx.translate(-3.164288029149254,-3.165288029149254);
        ctx.lineTo(15.245999999999999,3.7670000000000003);
        ctx.translate(15.115062645001482,4.75839064402787);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.4394819108758188,0.6746312082704444,0);
        ctx.rotate(0);
        ctx.translate(-15.115062645001482,-4.75839064402787);
        ctx.lineTo(14.296,7.383000000000001);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        ctx.restore();
        ctx.save();
        ctx.fillStyle="#000";
        ctx.font="   15px ''";
        ctx.beginPath();
        ctx.moveTo(17.98,13.739);
        ctx.lineTo(13.665,9.42);
        ctx.translate(15.079105472815456,8.005678356330014);
        ctx.rotate(0);
        ctx.arc(0,0,2,2.356118062253705,3.8152380280952944,0);
        ctx.rotate(0);
        ctx.translate(-15.079105472815456,-8.005678356330014);
        ctx.lineTo(15.116,4.756);
        ctx.lineTo(3.164,3.165);
        ctx.lineTo(4.757,15.114999999999998);
        ctx.lineTo(6.757999999999999,13.516999999999998);
        ctx.translate(8.006790211860212,15.079217336596326);
        ctx.rotate(0);
        ctx.arc(0,0,2,-2.24515321531101,-0.7854008321798522,0);
        ctx.rotate(0);
        ctx.translate(-8.006790211860212,-15.079217336596326);
        ctx.lineTo(13.738,17.982);
        ctx.lineTo(17.981,13.738999999999999);
        ctx.closePath();
        ctx.moveTo(14.370000000000001,8.714);
        ctx.translate(15.077886116101272,8.007673413617287);
        ctx.rotate(0);
        ctx.arc(0,0,1,2.357297244455302,3.816305891162912,0);
        ctx.rotate(0);
        ctx.translate(-15.077886116101272,-8.007673413617287);
        ctx.lineTo(15.896,5.3820000000000014);
        ctx.translate(15.115062645001485,4.75739064402787);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.6746312082704438,-1.4394819108758186,1);
        ctx.rotate(0);
        ctx.translate(-15.115062645001485,-4.75739064402787);
        ctx.lineTo(3.297,2.173);
        ctx.translate(3.165288029149254,3.164288029149254);
        ctx.rotate(0);
        ctx.arc(0,0,1,-1.4387005277668998,-3.2736884526177907,1);
        ctx.rotate(0);
        ctx.translate(-3.165288029149254,-3.164288029149254);
        ctx.lineTo(3.7670000000000003,15.245999999999999);
        ctx.translate(4.75839064402787,15.115062645001482);
        ctx.rotate(0);
        ctx.arc(0,0,1,3.010278237670716,0.8961651185244528,1);
        ctx.rotate(0);
        ctx.translate(-4.75839064402787,-15.115062645001482);
        ctx.lineTo(7.383000000000001,14.296);
        ctx.translate(8.007117659339036,15.077330370138755);
        ctx.rotate(0);
        ctx.arc(0,0,1,-2.2447980698220285,-0.7843007542967229,0);
        ctx.rotate(0);
        ctx.translate(-8.007117659339036,-15.077330370138755);
        ctx.lineTo(13.032000000000002,18.688);
        ctx.translate(13.739,17.980786453749644);
        ctx.rotate(0);
        ctx.arc(0,0,1,2.3560434901900495,0.785549163399744,1);
        ctx.rotate(0);
        ctx.translate(-13.739,-17.980786453749644);
        ctx.lineTo(18.689,14.445999999999998);
        ctx.translate(17.98228665641577,13.738499999999998);
        ctx.rotate(0);
        ctx.arc(0,0,1,0.7859544135121888,-0.7859544135121888,1);
        ctx.rotate(0);
        ctx.translate(-17.98228665641577,-13.738499999999998);
        ctx.lineTo(14.372,8.713999999999999);
        ctx.closePath();
        ctx.fill("evenodd");
        ctx.stroke();
        ctx.restore();

        ctx.restore();
    }
})();

