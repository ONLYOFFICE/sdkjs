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

"use strict";

(function(window, undefined){

// Import
var CColor = AscCommon.CColor;
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var History = AscCommon.History;
var global_MatrixTransformer = AscCommon.global_MatrixTransformer;
    var g_dKoef_pix_to_mm = AscCommon.g_dKoef_pix_to_mm;
    var g_dKoef_mm_to_pix = AscCommon.g_dKoef_mm_to_pix;



    function CDrawingCollaborativeTarget(DrawingDocument)
    {
        AscCommon.CDrawingCollaborativeTargetBase.call(this);
        this.SheetId = null;
        this.DrawingDocument = DrawingDocument;
    }
    CDrawingCollaborativeTarget.prototype = Object.create(AscCommon.CDrawingCollaborativeTargetBase.prototype);

    CDrawingCollaborativeTarget.prototype.GetZoom = function()
    {
        return Asc.editor.wb.getZoom();
    };
    CDrawingCollaborativeTarget.prototype.ConvertCoords = function(x, y)
    {
        var oTrack = this.DrawingDocument.AutoShapesTrack;
        if(!oTrack)
        {
            return {X: 0, Y: 0};
        }
        var oGraphics = oTrack.Graphics;
        if(!oGraphics)
        {
            return {X: 0, Y: 0};
        }
        var oTransform = oGraphics.m_oCoordTransform;
        var _offX = 0;
        var _offY = 0;
        var dKoef = this.DrawingDocument.drawingObjects.convertMetric(1, 3, 0);
        if (oTransform)
        {
            _offX = oTransform.tx;
            _offY = oTransform.ty;
        }
        var _X = AscCommon.AscBrowser.convertToRetinaValue(_offX + dKoef * x, false);
        var _Y = AscCommon.AscBrowser.convertToRetinaValue(_offY + dKoef * y, false);
        return { X : _X, Y : _Y};
    };
    CDrawingCollaborativeTarget.prototype.GetMobileTouchManager = function()
    {
        return Asc.editor.wb.MobileTouchManager;
    };
    CDrawingCollaborativeTarget.prototype.GetParentElement = function()
    {
        return Asc.editor.HtmlElement;
    };
    CDrawingCollaborativeTarget.prototype.CheckPosition = function(_x, _y, _size, sheetId, _transform)
    {
        this.Transform = _transform;
        this.Size = _size;
        this.X = _x;
        this.Y = _y;
        this.SheetId = sheetId;
        this.Update();
    };
    CDrawingCollaborativeTarget.prototype.CheckStyleDisplay = function()
    {
    };
    CDrawingCollaborativeTarget.prototype.CheckNeedDraw = function()
    {
        var bShow = false;
        var oWorksheetView = Asc.editor.wb.getWorksheet();
        if(oWorksheetView)
        {
            var oModel = oWorksheetView.model;
            if(oModel)
            {
                if(oModel.Id === this.SheetId)
                {
                    bShow = true;
                }
            }
        }
        if(!bShow)
        {
            this.HtmlElement.style.display = "none";
        }
        else
        {
            this.HtmlElement.style.display = "block";
        }
        return bShow;
    };

function CDrawingDocument()
{
    this.IsLockObjectsEnable = false;

    AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.MarkerFormat, "14 8", "pointer");
    AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.SelectTableRow, "10 5", "default");
    AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.SelectTableColumn, "5 10", "default");
    AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.SelectTableContent, "10 10", "default");
    AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.TableEraser, "8 19", "pointer");

    this.m_oWordControl     = null;
    this.m_oLogicDocument   = null;
 	this.m_oDocumentRenderer = null;

    this.m_arrPages         = [];
    this.m_lPagesCount      = 0;

    this.m_lDrawingFirst    = -1;
    this.m_lDrawingEnd      = -1;
    this.m_lCurrentPage     = -1;

    this.m_lCountCalculatePages = 0;

    this.m_lTimerTargetId = -1;
    this.m_dTargetX = -1;
    this.m_dTargetY = -1;
    this.m_lTargetPage = -1;
    this.m_dTargetSize = 1;


    this.TargetHtmlElement = null;
    this.TargetHtmlElementLeft = 0;
    this.TargetHtmlElementTop = 0;

    this.CollaborativeTargets = [];
    this.CollaborativeTargetsUpdateTasks = [];

    this.m_bIsBreakRecalculate = false;

    this.m_bIsSelection = false;
    this.m_bIsSearching = false;

    this.MathTrack = new AscCommon.CMathTrack();

    this.CurrentSearchNavi = null;
    this.SearchTransform = null;

    this.m_lTimerUpdateTargetID = -1;
    this.m_tempX = 0;
    this.m_tempY = 0;
    this.m_tempPageIndex = 0;

    var oThis = this;
    this.m_sLockedCursorType = "";

    this.m_lCurrentRendererPage = -1;
    this.m_oDocRenderer = null;
    this.m_bOldShowMarks = false;

    this.UpdateTargetFromPaint = false;
    this.UpdateTargetCheck = false;
    this.NeedTarget = true;
    this.TextMatrix = null;
    this.TargetShowFlag = false;
    this.TargetShowNeedFlag = false;

    this.CanvasHit = document.createElement('canvas');
    this.CanvasHit.width = 10;
    this.CanvasHit.height = 10;
    this.CanvasHitContext = this.CanvasHit.getContext('2d');

    this.TargetCursorColor = {R: 0, G: 0, B: 0};

    this.GuiControlColorsMap = null;
    this.IsSendStandartColors = false;

    this.GuiCanvasFillTextureParentId = "";
    this.GuiCanvasFillTexture = null;
    this.GuiCanvasFillTextureCtx = null;
    this.LastDrawingUrl = "";

    this.GuiCanvasFillTextureParentIdTextArt = "";
    this.GuiCanvasFillTextureTextArt = null;
    this.GuiCanvasFillTextureCtxTextArt = null;
    this.LastDrawingUrlTextArt = "";

	this.SelectionMatrix = null;

    this.GuiCanvasTextProps = null;
    this.GuiLastTextProps = null;

    this.TableStylesLastLook = null;
    this.LastParagraphMargins = null;

    this.AutoShapesTrack = null;
    this.AutoShapesTrackLockPageNum = -1;

    this.Overlay = null;
    this.IsTextMatrixUse = false;

    this.placeholders = new AscCommon.DrawingPlaceholders(this);

    this.getDrawingObjects = function()
    {
        var oWs = Asc.editor.wb.getWorksheet();
        if(oWs) {
            return oWs.objectRender;
        }
        return null;
    };

    this.Start_CollaborationEditing = function()
    {
    };

    this.SetCursorType = function(sType, Data)
    {
    };

    this.LockCursorType    = function(sType)
    {
        this.m_sLockedCursorType                                    = sType;

        if ( Asc.editor &&  Asc.editor.wb) {
            Asc.editor.wb._onUpdateCursor(this.m_sLockedCursorType);
        }
    };
    this.LockCursorTypeCur = function()
    {
    };
    this.UnlockCursorType  = function()
    {
        this.m_sLockedCursorType = "";
        const oWBView = Asc.editor &&  Asc.editor.wb;
        if (oWBView) {
            const ws = oWBView.getWorksheet();
            if (ws) {
                const ct = ws.getCursorTypeFromXY(ws.objectRender.lastX, ws.objectRender.lastY);
                if (ct) {
                    oWBView._onUpdateCursor(ct.cursor);
                }
            }
        }
    };

    this.OnStartRecalculate = function(pageCount)
    {
    };

    this.OnRepaintPage = function(index)
    {
    };

    this.OnRecalculatePage = function(index, pageObject)
    {
    };

    this.OnEndRecalculate = function(isFull, isBreak)
    {
    };

	this.ChangePageAttack = function(pageIndex)
	{
	};

    this.StartRenderingPage = function(pageIndex)
    {
    };

    this.IsFreezePage = function(pageIndex)
    {
        return true;
    };

    this.RenderDocument = function(Renderer)
    {
    };

    this.ToRenderer = function()
    {
        return "";
    };

    this.ToRenderer2 = function()
    {
        var ret = "";
        return ret;
    };

    this.ToRendererPart = function()
    {
    };

    this.StopRenderingPage = function(pageIndex)
    {
    };

    this.ClearCachePages = function()
    {
    };

    this.CheckRasterImageOnScreen = function(src)
    {
    };

    this.FirePaint = function()
    {
    };

    this.ConvertCoordsFromCursor = function(x, y, bIsRul)
    {
        return { X : 0, Y : 0, Page: -1, DrawPage: -1 };
    };

    this.ConvertCoordsFromCursorPage = function(x, y, page, bIsRul)
    {
        return { X : 0, Y : 0, Page: -1, DrawPage: -1 };
    };

    this.ConvertCoordsToAnotherPage = function(x, y, pageCoord, pageNeed)
    {
            return { X : 0, Y : 0, Error: true };

    };

    this.ConvertCoordsFromCursor2 = function(x, y, zoomVal, isRuler)
    {
        return { X : 0, Y : 0, Page: -1, DrawPage: -1 };
    };

    this.ConvetToPageCoords = function(x,y,pageIndex)
    {
        return { X : 0, Y : 0, Page : pageIndex, Error: true };
    };

    this.IsCursorInTableCur = function(x, y, page)
    {
        return false;
    };

    this.ConvertCoordsToCursorWR = function(x, y, pageIndex, transform)
    {
	    let oRender = this.getDrawingObjects();
		if(!oRender)
		{
			return { X : 0, Y : 0, Error: true };
		}
		return oRender.convertCoordsToCursorWR(x, y);
    };

    this.ConvertCoordsToCursor = function(x, y, pageIndex, bIsRul)
    {
        return { X : 0, Y : 0, Error: true };
    };

    this.ConvertCoordsToCursor2 = function(x, y, pageIndex, bIsRul)
    {
        return { X : 0, Y : 0, Error: true };
    };

    this.ConvertCoordsToCursor3 = function(x, y, pageIndex)
    {
        return { X : 0, Y : 0, Error: true };
    };

    this.TargetStart = function()
    {
        if ( this.m_lTimerTargetId != -1 )
            clearInterval( this.m_lTimerTargetId );
        this.m_lTimerTargetId = setInterval( oThis.DrawTarget, 500 );
    };

    this.TargetEnd = function()
    {
        this.TargetShowFlag = false;
        this.TargetShowNeedFlag = false;

        if (this.m_lTimerTargetId != -1)
        {
            clearInterval( this.m_lTimerTargetId );
            this.m_lTimerTargetId = -1;
        }
        this.TargetHtmlElement.style.display = "none";
    };

    this.UpdateTargetNoAttack = function()
    {
    };

    this.GetTargetStyle = function()
    {
        return "rgb(" + this.TargetCursorColor.R + "," + this.TargetCursorColor.G + "," + this.TargetCursorColor.B + ")";
    };

    this.SetTargetColor = function(r, g, b)
    {
        this.TargetCursorColor.R = r;
        this.TargetCursorColor.G = g;
        this.TargetCursorColor.B = b;
    };

    this.CheckTargetDraw = function (x, y)
    {
        var drawingObjects = this.getDrawingObjects();
        if (!drawingObjects)
            return;

        var dKoef = drawingObjects.convertMetric(1, 3, 0);
        dKoef /= AscCommon.AscBrowser.retinaPixelRatio;

        var oldW = this.TargetHtmlElement.width_old;
        var oldH = this.TargetHtmlElement.height_old;

        var newW = 2;
        var newH = (this.m_dTargetSize * dKoef) >> 0;

        var _offX = 0;
        var _offY = 0;
        if (this.AutoShapesTrack && this.AutoShapesTrack.Graphics && this.AutoShapesTrack.Graphics.m_oCoordTransform)
        {
            _offX = this.AutoShapesTrack.Graphics.m_oCoordTransform.tx / AscCommon.AscBrowser.retinaPixelRatio;
            _offY = this.AutoShapesTrack.Graphics.m_oCoordTransform.ty / AscCommon.AscBrowser.retinaPixelRatio;
        }

        this.TargetHtmlElement.style.transformOrigin = "top left";

        if (oldW !== newW || oldH !== newH)
        {
            var pixNewW = ((newW * AscCommon.AscBrowser.retinaPixelRatio) >> 0) / AscCommon.AscBrowser.retinaPixelRatio;

            this.TargetHtmlElement.style.width = pixNewW + "px";
            this.TargetHtmlElement.style.height = newH + "px";
            this.TargetHtmlElement.width_old = newW;
            this.TargetHtmlElement.height_old = newH;
            this.TargetHtmlElement.oldColor = null;
        }

        var oldColor = this.TargetHtmlElement.oldColor;
        if (!oldColor ||
            oldColor.R !== this.TargetCursorColor.R ||
            oldColor.G !== this.TargetCursorColor.G ||
            oldColor.B !== this.TargetCursorColor.B)
        {
            this.TargetHtmlElement.style.backgroundColor = this.GetTargetStyle();
            this.TargetHtmlElement.oldColor = { R : this.TargetCursorColor.R, G : this.TargetCursorColor.G, B : this.TargetCursorColor.B };
        }

        if (null == this.TextMatrix || global_MatrixTransformer.IsIdentity2(this.TextMatrix))
        {
            if (null != this.TextMatrix)
            {
                x += this.TextMatrix.tx;
                y += this.TextMatrix.ty;
            }

            var pos1 = { X : _offX + x * dKoef, Y : _offY + y * dKoef };
            pos1.X -= (newW / 2);

            this.TargetHtmlElementLeft = pos1.X >> 0;
            this.TargetHtmlElementTop = pos1.Y >> 0;

            this.TargetHtmlElement.style["transform"] = "";
            this.TargetHtmlElement.style["msTransform"] = "";
            this.TargetHtmlElement.style["mozTransform"] = "";
            this.TargetHtmlElement.style["webkitTransform"] = "";

            if (!AscCommon.AscBrowser.isSafariMacOs || !AscCommon.AscBrowser.isWebkit)
            {
                this.TargetHtmlElement.style.left = this.TargetHtmlElementLeft + "px";
                this.TargetHtmlElement.style.top = this.TargetHtmlElementTop + "px";
            }
            else
            {
                this.TargetHtmlElement.style.left = "0px";
                this.TargetHtmlElement.style.top = "0px";
                this.TargetHtmlElement.style["webkitTransform"] = "matrix(1, 0, 0, 1, " + oThis.TargetHtmlElementLeft + "," + oThis.TargetHtmlElementTop + ")";
            }
        }
        else
        {
            var x1 = _offX + this.TextMatrix.TransformPointX(x, y) * dKoef;
            var y1 = _offY + this.TextMatrix.TransformPointY(x, y) * dKoef;

            var pos1 = { X : x1, Y : y1 };
            pos1.X -= (newW / 2);

            this.TargetHtmlElementLeft = pos1.X >> 0;
            this.TargetHtmlElementTop = pos1.Y >> 0;

            var transform = "matrix(" + this.TextMatrix.sx + ", " + this.TextMatrix.shy + ", " + this.TextMatrix.shx + ", " +
                this.TextMatrix.sy + ", " + pos1.X + ", " + pos1.Y + ")";

            this.TargetHtmlElement.style.left = "0px";
            this.TargetHtmlElement.style.top = "0px";

            this.TargetHtmlElement.style["transform"] = transform;
            this.TargetHtmlElement.style["msTransform"] = transform;
            this.TargetHtmlElement.style["mozTransform"] = transform;
            this.TargetHtmlElement.style["webkitTransform"] = transform;
        }

        if (AscCommon.g_inputContext)
            AscCommon.g_inputContext.move(this.TargetHtmlElementLeft, this.TargetHtmlElementTop);
    };

    this.UpdateTargetTransform = function(matrix)
    {
        this.TextMatrix = matrix;
    };

    this.UpdateTarget = function(x, y, pageIndex)
    {
        /*
        if (this.UpdateTargetFromPaint === false)
        {
            this.UpdateTargetCheck = true;

            if (this.NeedScrollToTargetFlag && this.m_dTargetX == x && this.m_dTargetY == y && this.m_lTargetPage == pageIndex)
                this.NeedScrollToTarget = false;
            else
                this.NeedScrollToTarget = true;

            return;
        }
        */

        if (-1 != this.m_lTimerUpdateTargetID)
        {
            clearTimeout(this.m_lTimerUpdateTargetID);
            this.m_lTimerUpdateTargetID = -1;
        }

        this.m_dTargetX = x;
        this.m_dTargetY = y;
        this.m_lTargetPage = pageIndex;

        this.CheckTargetDraw(x, y);
    };

    this.UpdateTargetTimer = function()
    {
    };

    this.SetTargetSize = function(size)
    {
        this.m_dTargetSize = size;
        //this.TargetHtmlElement.style.height = Number(this.m_dTargetSize * this.m_oWordControl.m_nZoomValue * AscCommon.g_dKoef_mm_to_pix / 100) + "px";
        //this.TargetHtmlElement.style.width = "2px";
    };

    this.DrawTarget = function()
    {
        if ( "block" != oThis.TargetHtmlElement.style.display && oThis.NeedTarget )
            oThis.TargetHtmlElement.style.display = "block";
        else
            oThis.TargetHtmlElement.style.display = "none";
    };

    this.TargetShow = function()
    {
        this.TargetShowNeedFlag = true;
    };

    this.CheckTargetShow = function()
    {
        if (this.TargetShowFlag && this.TargetShowNeedFlag)
        {
            this.TargetHtmlElement.style.display = "block";
            this.TargetShowNeedFlag = false;
            return;
        }

        if (!this.TargetShowNeedFlag)
            return;

        this.TargetShowNeedFlag = false;

        if ( -1 == this.m_lTimerTargetId )
            this.TargetStart();

        if (oThis.NeedTarget)
            this.TargetHtmlElement.style.display = "block";

        this.TargetShowFlag = true;
    };

    this.StartTrackImage = function(obj, x, y, w, h, type, pagenum)
    {
    };

    this.StartTrackTable = function(obj, transform)
    {
    };

    this.EndTrackTable = function(pointer, bIsAttack)
    {
    };

    this.CheckTrackTable = function()
    {
    };

    this.DrawTableTrack = function(overlay)
    {
    };

    this.DrawMathTrack = function (overlay)
    {
        if (!this.MathTrack.IsActive() || !this.TextMatrix)
            return;
        var drawingObjects = this.getDrawingObjects();
        if (!drawingObjects)
            return;

        overlay.Show();
        var nIndex, nCount;
        var oPath;
        var dKoefX, dKoefY;
        var PathLng = this.MathTrack.GetPolygonsCount();

        dKoefX = drawingObjects.convertMetric(1, 3, 0);
        dKoefY = dKoefX;

        var _offX = 0;
        var _offY = 0;
        if (this.AutoShapesTrack && this.AutoShapesTrack.Graphics && this.AutoShapesTrack.Graphics.m_oCoordTransform)
        {
            _offX = this.AutoShapesTrack.Graphics.m_oCoordTransform.tx;
            _offY = this.AutoShapesTrack.Graphics.m_oCoordTransform.ty;
        }

        // Draw methods apply retina scaling.
        if (true)
        {
            var rPR = AscCommon.AscBrowser.retinaPixelRatio;
			dKoefX /= rPR;
			dKoefY /= rPR;
			_offX /= rPR;
			_offY /= rPR;
        }

        for (nIndex = 0; nIndex < PathLng; nIndex++)
        {
            oPath = this.MathTrack.GetPolygon(nIndex);
            this.MathTrack.Draw(overlay, oPath, 0, 0, "#939393", dKoefX, dKoefY, _offX, _offY, this.TextMatrix);
            this.MathTrack.Draw(overlay, oPath, 1, 1, "#FFFFFF", dKoefX, dKoefY, _offX, _offY, this.TextMatrix);
        }
        for (nIndex = 0, nCount = this.MathTrack.GetSelectPathsCount(); nIndex < nCount; nIndex++)
        {
            oPath =  this.MathTrack.GetSelectPath(nIndex);
            this.MathTrack.DrawSelectPolygon(overlay, oPath, dKoefX, dKoefY, _offX, _offY, this.TextMatrix);
        }
    };

    this.SetCurrentPage = function(PageIndex)
    {
    };

    this.SelectEnabled = function(bIsEnabled)
    {
        var drawingObjects = this.getDrawingObjects();
        if(!drawingObjects) {
            return;
        }
        this.m_bIsSelection = bIsEnabled;
        if (false === this.m_bIsSelection)
        {
            this.SelectClear();
            drawingObjects.getOverlay().m_oContext.globalAlpha = 1.0;
        }
    };

	this.SelectClear = function()
    {

    };

    this.SearchClear = function()
    {
    };

    this.AddPageSearch = function(findText, rects, type)
    {

    };

    this.StartSearchTransform = function(transform)
    {
    };

    this.EndSearchTransform = function()
    {
    };

    this.StartSearch = function()
    {
    };

    this.EndSearch = function(bIsChange)
    {
    };

    this.private_StartDrawSelection = function(overlay)
    {
        this.Overlay = overlay;
        this.IsTextMatrixUse = ((null != this.TextMatrix) && !global_MatrixTransformer.IsIdentity(this.TextMatrix));

        this.Overlay.m_oContext.fillStyle = "rgba(51,102,204,255)";
        this.Overlay.m_oContext.beginPath();

        if (this.IsTextMatrixUse)
            this.Overlay.m_oContext.strokeStyle = "#9ADBFE";
    };

    this.private_EndDrawSelection = function()
    {
        var ctx = this.Overlay.m_oContext;

        ctx.globalAlpha = 0.2;
        ctx.fill();

        if (this.IsTextMatrixUse)
        {
            ctx.globalAlpha = 1.0;
			ctx.lineWidth = Math.round(AscCommon.AscBrowser.retinaPixelRatio);
            ctx.stroke();
        }

        ctx.beginPath();
        ctx.globalAlpha = 1.0;

        this.IsTextMatrixUse = false;
        this.Overlay = null;
    };

    this.AddPageSelection = function(pageIndex, x, y, w, h)
    {
        var drawingObjects = this.getDrawingObjects();
        if(!drawingObjects) {
            return;
        }
		if (null == this.SelectionMatrix)
			this.SelectionMatrix = this.TextMatrix;

        var dKoefX = drawingObjects.convertMetric(1, 3, 0);
        var dKoefY = dKoefX;
		var indent = 0.5 * Math.round(AscCommon.AscBrowser.retinaPixelRatio);

        var _offX = 0;
        var _offY = 0;
        if (this.AutoShapesTrack && this.AutoShapesTrack.Graphics && this.AutoShapesTrack.Graphics.m_oCoordTransform)
        {
            _offX = this.AutoShapesTrack.Graphics.m_oCoordTransform.tx;
            _offY = this.AutoShapesTrack.Graphics.m_oCoordTransform.ty;
        }

		if (null == this.TextMatrix || global_MatrixTransformer.IsIdentity(this.TextMatrix))
		{
			var _x = ((_offX + dKoefX * x + indent) >> 0) - indent;
			var _y = ((_offY + dKoefY * y + indent) >> 0) - indent;

			var _w = (dKoefX * w + 1) >> 0;
			var _h = (dKoefY * h + 1) >> 0;

			this.Overlay.CheckRect(_x, _y, _w, _h);
			this.Overlay.m_oContext.rect(_x,_y,_w,_h);
		}
		else
		{
			var _x1 = this.TextMatrix.TransformPointX(x, y);
			var _y1 = this.TextMatrix.TransformPointY(x, y);

			var _x2 = this.TextMatrix.TransformPointX(x + w, y);
			var _y2 = this.TextMatrix.TransformPointY(x + w, y);

			var _x3 = this.TextMatrix.TransformPointX(x + w, y + h);
			var _y3 = this.TextMatrix.TransformPointY(x + w, y + h);

			var _x4 = this.TextMatrix.TransformPointX(x, y + h);
			var _y4 = this.TextMatrix.TransformPointY(x, y + h);

			var x1 = _offX + dKoefX * _x1;
			var y1 = _offY + dKoefY * _y1;

			var x2 = _offX + dKoefX * _x2;
			var y2 = _offY + dKoefY * _y2;

			var x3 = _offX + dKoefX * _x3;
			var y3 = _offY + dKoefY * _y3;

			var x4 = _offX + dKoefX * _x4;
			var y4 = _offY + dKoefY * _y4;

			if (global_MatrixTransformer.IsIdentity2(this.TextMatrix))
			{
				x1 = (x1 >> 0) + indent;
				y1 = (y1 >> 0) + indent;

				x2 = (x2 >> 0) + indent;
				y2 = (y2 >> 0) + indent;

				x3 = (x3 >> 0) + indent;
				y3 = (y3 >> 0) + indent;

				x4 = (x4 >> 0) + indent;
				y4 = (y4 >> 0) + indent;
			}

			this.Overlay.CheckPoint(x1, y1);
			this.Overlay.CheckPoint(x2, y2);
			this.Overlay.CheckPoint(x3, y3);
			this.Overlay.CheckPoint(x4, y4);

			var ctx = this.Overlay.m_oContext;
			ctx.moveTo(x1, y1);
			ctx.lineTo(x2, y2);
			ctx.lineTo(x3, y3);
			ctx.lineTo(x4, y4);
			ctx.closePath();
		}
    };

    this.SelectShow = function()
    {
        var drawingObjects = this.getDrawingObjects();
        if(!drawingObjects) {
            return;
        }   
        drawingObjects.OnUpdateOverlay();
    };

    this.Set_RulerState_Table = function(markup, transform)
    {
    };

    this.Set_RulerState_Paragraph = function(margins)
    {
    };

    this.Set_RulerState_HdrFtr = function(bHeader, Y0, Y1)
    {
    };

    this.Update_MathTrack = function (IsActive, IsContentActive, oMath)
    {
        var PixelError = this.GetMMPerDot(1) * 3;
        this.MathTrack.Update(IsActive, IsContentActive, oMath, PixelError);
    };

    this.Update_ParaTab = function(Default_Tab, ParaTabs)
    {
    };

    this.UpdateTableRuler = function(isCols, index, position)
    {
    };

    this.GetDotsPerMM = function(value)
    {
        return value * Asc.editor.wb.getZoom() * AscCommon.g_dKoef_mm_to_pix;
    };

    this.GetMMPerDot = function(value)
    {
        return value / this.GetDotsPerMM( 1 );
    };
    this.GetVisibleMMHeight = function()
    {
        return 0;
    };

    // вот оооочень важная функция. она выкидывает из кэша неиспользуемые шрифты
    this.CheckFontCache = function()
    {
    };

    // при загрузке документа - нужно понять какие шрифты используются
    this.CheckFontNeeds = function()
    {
    };

    // фукнции для старта работы
    this.OpenDocument = function()
    {
    };

    // вот здесь весь трекинг
    this.DrawTrack = function(type, matrix, left, top, width, height, isLine, canRotate, isNoMove, isDrawHandles)
    {
        this.AutoShapesTrack.DrawTrack(type, matrix, left, top, width, height, isLine, canRotate, isNoMove, isDrawHandles);
    };

    this.DrawTrackSelectShapes = function(x, y, w, h)
    {
        this.AutoShapesTrack.DrawTrackSelectShapes(x, y, w, h);
    };

    this.DrawAdjustment = function(matrix, x, y, bTextWarp)
    {
        this.AutoShapesTrack.DrawAdjustment(matrix, x, y, bTextWarp);
    };

    this.LockTrackPageNum = function(nPageNum)
    {
        this.AutoShapesTrackLockPageNum = nPageNum;
    };
    this.UnlockTrackPageNum = function()
    {
        this.AutoShapesTrackLockPageNum = -1;
    };

    this.CheckGuiControlColors = function()
    {
    };

    this.SendControlColors = function()
    {
    };

    this.DrawImageTextureFillShape = function(url)
    {
        if (this.GuiCanvasFillTexture == null)
        {
            this.InitGuiCanvasShape(this.GuiCanvasFillTextureParentId);
        }

        if (this.GuiCanvasFillTexture == null || this.GuiCanvasFillTextureCtx == null || url == this.LastDrawingUrl)
            return;

        this.LastDrawingUrl = url;
        var _width  = this.GuiCanvasFillTexture.width;
        var _height = this.GuiCanvasFillTexture.height;

        this.GuiCanvasFillTextureCtx.clearRect(0, 0, _width, _height);

        if (null == this.LastDrawingUrl)
            return;

		var api = window["Asc"]["editor"];
        var _img = api.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(this.LastDrawingUrl)];
        if (_img != undefined && _img.Image != null && _img.Status != AscFonts.ImageLoadStatus.Loading)
        {
            var _x = 0;
            var _y = 0;
            var _w = Math.max(_img.Image.width, 1);
            var _h = Math.max(_img.Image.height, 1);

            var dAspect1 = _width / _height;
            var dAspect2 = _w / _h;

            _w = _width;
            _h = _height;
            if (dAspect1 >= dAspect2)
            {
                _w = dAspect2 * _height;
                _x = (_width - _w) / 2;
            }
            else
            {
                _h = _w / dAspect2;
                _y = (_height - _h) / 2;
            }

            this.GuiCanvasFillTextureCtx.drawImage(_img.Image, _x, _y, _w, _h);
        }
        else
        {
            this.GuiCanvasFillTextureCtx.lineWidth = 1;

            this.GuiCanvasFillTextureCtx.beginPath();
            this.GuiCanvasFillTextureCtx.moveTo(0, 0);
            this.GuiCanvasFillTextureCtx.lineTo(_width, _height);
            this.GuiCanvasFillTextureCtx.moveTo(_width, 0);
            this.GuiCanvasFillTextureCtx.lineTo(0, _height);
            this.GuiCanvasFillTextureCtx.strokeStyle = "#FF0000";
            this.GuiCanvasFillTextureCtx.stroke();

            this.GuiCanvasFillTextureCtx.beginPath();
            this.GuiCanvasFillTextureCtx.moveTo(0, 0);
            this.GuiCanvasFillTextureCtx.lineTo(_width, 0);
            this.GuiCanvasFillTextureCtx.lineTo(_width, _height);
            this.GuiCanvasFillTextureCtx.lineTo(0, _height);
            this.GuiCanvasFillTextureCtx.closePath();

            this.GuiCanvasFillTextureCtx.strokeStyle = "#000000";
            this.GuiCanvasFillTextureCtx.stroke();
            this.GuiCanvasFillTextureCtx.beginPath();
        }
    };

    this.DrawImageTextureFillTextArt = function(url)
    {
        if (this.GuiCanvasFillTextureTextArt == null)
        {
            this.InitGuiCanvasTextArt(this.GuiCanvasFillTextureParentIdTextArt);
        }

        if (this.GuiCanvasFillTextureTextArt == null || this.GuiCanvasFillTextureCtxTextArt == null || url == this.LastDrawingUrlTextArt)
            return;

        this.LastDrawingUrlTextArt = url;
        var _width  = this.GuiCanvasFillTextureTextArt.width;
        var _height = this.GuiCanvasFillTextureTextArt.height;

        this.GuiCanvasFillTextureCtxTextArt.clearRect(0, 0, _width, _height);

        if (null == this.LastDrawingUrlTextArt)
            return;

        var api = window["Asc"]["editor"];
        var _img = api.ImageLoader.map_image_index[AscCommon.getFullImageSrc2(this.LastDrawingUrlTextArt)];
        if (_img != undefined && _img.Image != null && _img.Status != AscFonts.ImageLoadStatus.Loading)
        {
            var _x = 0;
            var _y = 0;
            var _w = Math.max(_img.Image.width, 1);
            var _h = Math.max(_img.Image.height, 1);

            var dAspect1 = _width / _height;
            var dAspect2 = _w / _h;

            _w = _width;
            _h = _height;
            if (dAspect1 >= dAspect2)
            {
                _w = dAspect2 * _height;
                _x = (_width - _w) / 2;
            }
            else
            {
                _h = _w / dAspect2;
                _y = (_height - _h) / 2;
            }

            this.GuiCanvasFillTextureCtxTextArt.drawImage(_img.Image, _x, _y, _w, _h);
        }
        else
        {
            this.GuiCanvasFillTextureCtxTextArt.lineWidth = 1;

            this.GuiCanvasFillTextureCtxTextArt.beginPath();
            this.GuiCanvasFillTextureCtxTextArt.moveTo(0, 0);
            this.GuiCanvasFillTextureCtxTextArt.lineTo(_width, _height);
            this.GuiCanvasFillTextureCtxTextArt.moveTo(_width, 0);
            this.GuiCanvasFillTextureCtxTextArt.lineTo(0, _height);
            this.GuiCanvasFillTextureCtxTextArt.strokeStyle = "#FF0000";
            this.GuiCanvasFillTextureCtxTextArt.stroke();

            this.GuiCanvasFillTextureCtxTextArt.beginPath();
            this.GuiCanvasFillTextureCtxTextArt.moveTo(0, 0);
            this.GuiCanvasFillTextureCtxTextArt.lineTo(_width, 0);
            this.GuiCanvasFillTextureCtxTextArt.lineTo(_width, _height);
            this.GuiCanvasFillTextureCtxTextArt.lineTo(0, _height);
            this.GuiCanvasFillTextureCtxTextArt.closePath();

            this.GuiCanvasFillTextureCtxTextArt.strokeStyle = "#000000";
            this.GuiCanvasFillTextureCtxTextArt.stroke();
            this.GuiCanvasFillTextureCtxTextArt.beginPath();
        }
    };

    this.InitGuiCanvasShape = function(div_id)
    {
        if (this.GuiCanvasFillTextureParentId == div_id && null != this.GuiCanvasFillTexture)
            return;

        if (null != this.GuiCanvasFillTexture)
        {
            var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentId);
            if (_div_elem)
                _div_elem.removeChild(this.GuiCanvasFillTexture);

            this.GuiCanvasFillTexture = null;
            this.GuiCanvasFillTextureCtx = null;
        }

        this.GuiCanvasFillTextureParentId = div_id;
        var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentId);
        if (!_div_elem)
            return;

        var bIsAppend = true;
        if (_div_elem.childNodes && _div_elem.childNodes.length == 1)
        {
            this.GuiCanvasFillTexture = _div_elem.childNodes[0];
            bIsAppend = false;
        }
        else
        {
            this.GuiCanvasFillTexture = document.createElement('canvas');
        }

        this.GuiCanvasFillTexture.width = parseInt(_div_elem.style.width);
        this.GuiCanvasFillTexture.height = parseInt(_div_elem.style.height);

        this.LastDrawingUrl = "";
        this.GuiCanvasFillTextureCtx = this.GuiCanvasFillTexture.getContext('2d');

        if (bIsAppend)
        {
            _div_elem.appendChild(this.GuiCanvasFillTexture);
        }
    };

    this.InitGuiCanvasTextProps = function(div_id)
    {
		var _div_elem = document.getElementById(div_id);
		if (null != this.GuiCanvasTextProps)
		{
			var elem = _div_elem.getElementsByTagName('canvas');
			if (elem.length == 0)
			{
				_div_elem.appendChild(this.GuiCanvasTextProps);
			}
			else
			{
				var _width = parseInt(_div_elem.offsetWidth);
				var _height = parseInt(_div_elem.offsetHeight);
				if (0 == _width)
					_width = 300;
				if (0 == _height)
					_height = 80;

				this.GuiCanvasTextProps.style.width = _width + "px";
				this.GuiCanvasTextProps.style.height = _height + "px";
			}

			var old_width = this.GuiCanvasTextProps.width;
			var old_height = this.GuiCanvasTextProps.height;
			AscCommon.calculateCanvasSize(this.GuiCanvasTextProps);

			if (old_width !== this.GuiCanvasTextProps.width || old_height !== this.GuiCanvasTextProps.height)
				this.GuiLastTextProps = null;
		}
		else
		{
			this.GuiCanvasTextProps = document.createElement('canvas');
			this.GuiCanvasTextProps.style.position = "absolute";
			this.GuiCanvasTextProps.style.left = "0px";
			this.GuiCanvasTextProps.style.top = "0px";

			var _width = parseInt(_div_elem.offsetWidth);
			var _height = parseInt(_div_elem.offsetHeight);
			if (0 == _width)
				_width = 300;
			if (0 == _height)
				_height = 80;

			this.GuiCanvasTextProps.style.width = _width + "px";
			this.GuiCanvasTextProps.style.height = _height + "px";

			AscCommon.calculateCanvasSize(this.GuiCanvasTextProps);

			_div_elem.appendChild(this.GuiCanvasTextProps);
		}
    };

    this.InitGuiCanvasTextArt = function(div_id)
    {
        if (null != this.GuiCanvasFillTextureTextArt)
        {
            var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentIdTextArt);
            if (!_div_elem)
                _div_elem.removeChild(this.GuiCanvasFillTextureTextArt);

            this.GuiCanvasFillTextureTextArt = null;
            this.GuiCanvasFillTextureCtxTextArt = null;
        }

        this.GuiCanvasFillTextureParentIdTextArt = div_id;
        var _div_elem = document.getElementById(this.GuiCanvasFillTextureParentIdTextArt);
        if (!_div_elem)
            return;

        var bIsAppend = true;
        if (_div_elem.childNodes && _div_elem.childNodes.length == 1)
        {
            this.GuiCanvasFillTextureTextArt = _div_elem.childNodes[0];
            bIsAppend = false;
        }
        else
        {
            this.GuiCanvasFillTextureTextArt = document.createElement('canvas');
        }

        this.GuiCanvasFillTextureTextArt.width = parseInt(_div_elem.style.width);
        this.GuiCanvasFillTextureTextArt.height = parseInt(_div_elem.style.height);

        this.LastDrawingUrlTextArt = "";
        this.GuiCanvasFillTextureCtxTextArt = this.GuiCanvasFillTextureTextArt.getContext('2d');

        if (bIsAppend)
        {
            _div_elem.appendChild(this.GuiCanvasFillTextureTextArt);
        }
    };

    this.DrawGuiCanvasTextProps = function(props)
    {
        var bIsChange = false;
        if (null == this.GuiLastTextProps)
        {
            bIsChange = true;

            this.GuiLastTextProps = new Asc.asc_CParagraphProperty();

            this.GuiLastTextProps.Subscript     = props.Subscript;
            this.GuiLastTextProps.Superscript   = props.Superscript;
            this.GuiLastTextProps.SmallCaps     = props.SmallCaps;
            this.GuiLastTextProps.AllCaps       = props.AllCaps;
            this.GuiLastTextProps.Strikeout     = props.Strikeout;
            this.GuiLastTextProps.DStrikeout    = props.DStrikeout;

            this.GuiLastTextProps.TextSpacing   = props.TextSpacing;
            this.GuiLastTextProps.Position      = props.Position;
        }
        else
        {
            if (this.GuiLastTextProps.Subscript != props.Subscript)
            {
                this.GuiLastTextProps.Subscript = props.Subscript;
                bIsChange = true;
            }
            if (this.GuiLastTextProps.Superscript != props.Superscript)
            {
                this.GuiLastTextProps.Superscript = props.Superscript;
                bIsChange = true;
            }
            if (this.GuiLastTextProps.SmallCaps != props.SmallCaps)
            {
                this.GuiLastTextProps.SmallCaps = props.SmallCaps;
                bIsChange = true;
            }
            if (this.GuiLastTextProps.AllCaps != props.AllCaps)
            {
                this.GuiLastTextProps.AllCaps = props.AllCaps;
                bIsChange = true;
            }
            if (this.GuiLastTextProps.Strikeout != props.Strikeout)
            {
                this.GuiLastTextProps.Strikeout = props.Strikeout;
                bIsChange = true;
            }
            if (this.GuiLastTextProps.DStrikeout != props.DStrikeout)
            {
                this.GuiLastTextProps.DStrikeout = props.DStrikeout;
                bIsChange = true;
            }
            if (this.GuiLastTextProps.TextSpacing != props.TextSpacing)
            {
                this.GuiLastTextProps.TextSpacing = props.TextSpacing;
                bIsChange = true;
            }
            if (this.GuiLastTextProps.Position != props.Position)
            {
                this.GuiLastTextProps.Position = props.Position;
                bIsChange = true;
            }
        }

        if (undefined !== this.GuiLastTextProps.Position && isNaN(this.GuiLastTextProps.Position))
            this.GuiLastTextProps.Position = undefined;
        if (undefined !== this.GuiLastTextProps.TextSpacing && isNaN(this.GuiLastTextProps.TextSpacing))
            this.GuiLastTextProps.TextSpacing = undefined;

        if (!bIsChange)
            return;

        AscFormat.ExecuteNoHistory(function()
        {
            var shape = new AscFormat.CShape();
            shape.setTxBody(AscFormat.CreateTextBodyFromString("", this, shape));
            var par = shape.txBody.content.Content[0];
            par.Reset(0, 0, 1000, 1000, 0);
            par.MoveCursorToStartPos();
            var _paraPr = new CParaPr();
            par.Pr = _paraPr;
            var _textPr = new CTextPr();
            _textPr.FontFamily = { Name : "Arial", Index : -1 };
			_textPr.FontSize = (AscCommon.AscBrowser.convertToRetinaValue(11 << 1, true) >> 0) * 0.5;
            _textPr.RFonts.SetAll("Arial");

            _textPr.Strikeout  = this.GuiLastTextProps.Strikeout;

            if (true === this.GuiLastTextProps.Subscript)
                _textPr.VertAlign  = AscCommon.vertalign_SubScript;
            else if (true === this.GuiLastTextProps.Superscript)
                _textPr.VertAlign  = AscCommon.vertalign_SuperScript;
            else
                _textPr.VertAlign = AscCommon.vertalign_Baseline;

            _textPr.DStrikeout = this.GuiLastTextProps.DStrikeout;
            _textPr.Caps       = this.GuiLastTextProps.AllCaps;
            _textPr.SmallCaps  = this.GuiLastTextProps.SmallCaps;

            _textPr.Spacing    = this.GuiLastTextProps.TextSpacing;
            _textPr.Position   = this.GuiLastTextProps.Position;

            var parRun = new ParaRun(par);
            parRun.Set_Pr(_textPr);
            parRun.AddText("Hello World");
            par.Add_ToContent(0, parRun);

            par.Recalculate_Page(0);

            par.Recalculate_Page(0);

            var baseLineOffset = par.Lines[0].Y;
            var _bounds = par.Get_PageBounds(0);

            var ctx = this.GuiCanvasTextProps.getContext('2d');
            var _wPx = this.GuiCanvasTextProps.width;
            var _hPx = this.GuiCanvasTextProps.height;

            var _wMm = _wPx * g_dKoef_pix_to_mm;
            var _hMm = _hPx * g_dKoef_pix_to_mm;

            ctx.fillStyle = "#FFFFFF";
            ctx.fillRect(0, 0, _wPx, _hPx);

            var _pxBoundsW = par.Lines[0].Ranges[0].W * AscCommon.g_dKoef_mm_to_pix;//(_bounds.Right - _bounds.Left) * AscCommon.g_dKoef_mm_to_pix;
            var _pxBoundsH = (_bounds.Bottom - _bounds.Top) * AscCommon.g_dKoef_mm_to_pix;

            if (this.GuiLastTextProps.Position !== undefined && this.GuiLastTextProps.Position != null && this.GuiLastTextProps.Position != 0)
            {
                // TODO: нужна высота без учета Position
                // _pxBoundsH -= (this.GuiLastTextProps.Position * AscCommon.g_dKoef_mm_to_pix);
            }

            if (_pxBoundsH < _hPx && _pxBoundsW < _wPx)
            {
                // рисуем линию
                var _lineY = (((_hPx + _pxBoundsH) / 2) >> 0) + 0.5;
                var _lineW = (((_wPx - _pxBoundsW) / 4) >> 0);

                ctx.strokeStyle = "#000000";
                ctx.lineWidth = 1;

                ctx.beginPath();
                ctx.moveTo(0, _lineY);
                ctx.lineTo(_lineW, _lineY);

                ctx.moveTo(_wPx - _lineW, _lineY);
                ctx.lineTo(_wPx, _lineY);

                ctx.stroke();
                ctx.beginPath();
            }

            var _yOffset = (((_hPx + _pxBoundsH) / 2) - baseLineOffset * AscCommon.g_dKoef_mm_to_pix) >> 0;
            var _xOffset = ((_wPx - _pxBoundsW) / 2) >> 0;

            var graphics = new AscCommon.CGraphics();
            graphics.init(ctx, _wPx, _hPx, _wMm, _hMm);
            graphics.m_oFontManager = AscCommon.g_fontManager;

            graphics.m_oCoordTransform.tx = _xOffset;
            graphics.m_oCoordTransform.ty = _yOffset;

            graphics.transform(1,0,0,1,0,0);

            par.Draw(0, graphics);
        }, this, []);
    };

    this.CheckTableStyles = function(tableLook)
    {
    };

    this.GetTableStylesPreviews = function()
    {
        return [];
    };
    this.GetTableLook = function(isDefault)
    {
        let oTableLook;

        if (isDefault)
        {
            oTableLook = new AscCommon.CTableLook();
            oTableLook.SetDefault();
        }
        else
        {
            oTableLook = this.TableStylesLastLook;
        }

        return oTableLook;
    };

    this.IsMobileVersion = function()
    {
        return false;
    };

    this.OnSelectEnd = function()
    {
    };

    // collaborative targets
    this.Collaborative_UpdateTarget      = function(_id, _shortId, _x, _y, _size, sheetId, _transform, is_from_paint)
    {
        //if (is_from_paint !== true)
        //{
        //    this.CollaborativeTargetsUpdateTasks.push([_id, _shortId, _x, _y, _size, sheetId, _transform]);
        //    return;
        //}

        for (var i = 0; i < this.CollaborativeTargets.length; i++)
        {
            if (_id == this.CollaborativeTargets[i].Id)
            {
                this.CollaborativeTargets[i].CheckPosition(_x, _y, _size, sheetId, _transform);
                return;
            }
        }
        var _target     = new CDrawingCollaborativeTarget(this);
        _target.Id      = _id;
        _target.ShortId = _shortId;
        _target.SheetId = sheetId;
        _target.CheckPosition(_x, _y, _size, sheetId, _transform);
        this.CollaborativeTargets[this.CollaborativeTargets.length] = _target;
    };
    this.Collaborative_RemoveTarget      = function(_id)
    {
        var i = 0;
        for (i = 0; i < this.CollaborativeTargets.length; i++)
        {
            if (_id == this.CollaborativeTargets[i].Id)
            {
                this.CollaborativeTargets[i].Remove(this);
                this.CollaborativeTargets.splice(i, 1);
                i--;
            }
        }

        for (i = 0; i < this.CollaborativeTargetsUpdateTasks.length; i++)
        {
            var _tmp = this.CollaborativeTargetsUpdateTasks[i];
            if (_tmp[0] == _id)
            {
                this.CollaborativeTargetsUpdateTasks.splice(i, 1);
                i--;
            }
        }
    };
    this.Collaborative_TargetsUpdate     = function(bIsChangePosition)
    {
        if (bIsChangePosition)
        {
            for (var i = 0; i < this.CollaborativeTargets.length; i++)
            {
                this.CollaborativeTargets[i].Update();
            }
        }
    };
    this.Collaborative_GetTargetPosition = function(UserId)
    {
        for (var i = 0; i < this.CollaborativeTargets.length; i++)
        {
            if (UserId == this.CollaborativeTargets[i].Id)
                return {X : this.CollaborativeTargets[i].HtmlElementX, Y : this.CollaborativeTargets[i].HtmlElementY};
        }

        return null;
    };

	this.CloseFile = function ()
	{
		this.ClearCachePages();
		this.FirePaint();
		this.m_arrPages.splice(0, this.m_arrPages.length);
		this.m_lPagesCount = 0;

		this.m_lDrawingFirst = -1;
		this.m_lDrawingEnd = -1;
		this.m_lCurrentPage = -1;
	};
}

// заглушка
function CHtmlPage()
{
    this.drawingPage = { top: 0, left: 0, right: 0, bottom: 0 };
    this.width_mm = 0;
    this.height_mm = 0;
}

CHtmlPage.prototype.init = function(x, y, w_pix, h_pix, w_mm, h_mm) {
    this.drawingPage.top = y;
    this.drawingPage.left = x;
    this.drawingPage.right = w_pix;
    this.drawingPage.bottom = h_pix;
    this.width_mm = w_mm;
    this.height_mm = h_mm;
};

CHtmlPage.prototype.GetDrawingPageInfo = function() {

    return { drawingPage: this.drawingPage, width_mm: this.width_mm, height_mm: this.height_mm };
};

    //--------------------------------------------------------export----------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].CDrawingDocument = CDrawingDocument;
    window['AscCommon'].CHtmlPage = CHtmlPage;
})(window);
