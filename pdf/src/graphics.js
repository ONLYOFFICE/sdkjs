/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

function CPDFGraphics()
{
    this.context        = null;
    this.widthPx        = undefined;
    this.heightPx       = undefined;
    this.widthMM        = undefined;
    this.heightMM       = undefined;
    this.lineWidth      = 1;
    this.fillStyle      = null;
    this.strokeStyle    = null;
    this.globalAlpha    = 1;
    this.bIntegerGrid   = false;

    this.m_oTransform  = new AscCommon.CMatrixL();

    this.drawedRect = {};
}
CPDFGraphics.prototype.SetCurPage = function(nPage) {
    this.curPage = nPage;
};
CPDFGraphics.prototype.GetCurPage = function() {
    return this.curPage;
};
CPDFGraphics.prototype.GetScale = function() {
    let oViewer = editor.getDocumentRenderer();
    return AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom * (96 / oViewer.file.pages[this.GetCurPage()].Dpi);
};
CPDFGraphics.prototype.Transform = function(sx,shy,shx,sy,tx,ty)
{
    var _t = this.m_oTransform;
    _t.sx    = sx;
    _t.shx   = shx;
    _t.shy   = shy;
    _t.sy    = sy;
    _t.tx    = tx;
    _t.ty    = ty;

    this.context.setTransform(_t.sx,_t.shy,_t.shx,_t.sy,_t.tx,_t.ty);
};
CPDFGraphics.prototype.Clip = function() {
    this.context.clip();
};
CPDFGraphics.prototype.SetIntegerGrid = function() {
    this.bIntegerGrid = true;
};
CPDFGraphics.prototype.GetIntegerGrid = function() {
    return this.bIntegerGrid;
};

CPDFGraphics.prototype.SetStrokeStyle = function(r,g,b) {
    if (this.context)
        this.context.strokeStyle = "rgb(" + r + "," + g + "," + b + ")";

    this.strokeStyle = {
        r: r,
        g: g,
        b: b
    };
};
CPDFGraphics.prototype.SetFillStyle = function(r,g,b) {
    if (this.context)
        this.context.fillStyle = "rgb(" + r + "," + g + "," + b + ")";

    this.fillStyle = {
        r: r,
        g: g,
        b: b
    };
};
CPDFGraphics.prototype.SetLineWidth = function(width) {
    let nScale = this.GetScale();
    if (this.context)
        this.context.lineWidth = (nScale * width) + 0.5 >> 0;
        
    this.lineWidth = width;
};
CPDFGraphics.prototype.GetLineWidth = function(bScaled) {
    if (bScaled)
        return this.lineWidth * this.GetScale() + 0.5 >> 0;

    return this.lineWidth;
};
CPDFGraphics.prototype.Init = function(context, nWidthPx, nHeightPx) {
    this.context    = context;
    this.widthPx    = nWidthPx;
    this.heightPx   = nHeightPx;
    this.widthMM    = nWidthPx * g_dKoef_pix_to_mm;
    this.heightMM   = nHeightPx * g_dKoef_pix_to_mm;

    this.drawedRect = {
        xMin: nWidthPx,
        yMin: nHeightPx,
        xMax: 0,
        yMax: 0
    }
};
CPDFGraphics.prototype.CheckPoint = function(x, y) {
    return;
    if (this.drawedRect.xMin > x)
        this.drawedRect.xMin = x;
    if (this.drawedRect.xMax < x)
        this.drawedRect.xMax = x;

    if (this.drawedRect.yMin > y)
        this.drawedRect.yMin = y;
    if (this.drawedRect.yMax < y)
        this.drawedRect.yMax = y;
};
CPDFGraphics.prototype.GetDrawedRect = function(bScaled) {
    if (bScaled) {
        let nScale = this.GetScale();
        return {
            xMin: this.drawedRect.xMin * nScale >> 0,
            yMin: this.drawedRect.yMin * nScale >> 0,
            xMax: Math.round(this.drawedRect.xMax * nScale + 0.5),
            yMax: Math.round(this.drawedRect.yMax * nScale + 0.5)
        }
    }
        
    return this.drawedRect;
};

CPDFGraphics.prototype.Rect = function(x, y, w, h, isFill) {
    let nScale          = this.GetScale();
    let nLineW          = this.GetLineWidth();
    let bIntegerGrid    = this.GetIntegerGrid();

    this.CheckPoint(x - nLineW, y - nLineW);
    this.CheckPoint(x + w + nLineW, y + h + nLineW);

    let X1 = bIntegerGrid ? x * nScale >> 0 : x * nScale;
    let Y1 = bIntegerGrid ? y * nScale >> 0 : y * nScale;
    let X2 = bIntegerGrid ? (x + w) * nScale >> 0 : (x + w) * nScale;
    let Y2 = bIntegerGrid ? (y + h) * nScale >> 0 : (y + h) * nScale;

    let nLineOffsetX = (bIntegerGrid && !isFill) ? (0 === (nLineW * nScale % 2) ? 0 : 0.5) : 0;
    let nLineOffsetY = nLineOffsetX;

    this.context.moveTo(nLineOffsetX + X1, nLineOffsetY + Y1);
    this.context.lineTo(-nLineOffsetX + X2, nLineOffsetY + Y1);
    this.context.lineTo(-nLineOffsetX + X2, -nLineOffsetY + Y2);
    this.context.lineTo(nLineOffsetX + X1, -nLineOffsetY + Y2);
    this.context.closePath();
};
CPDFGraphics.prototype.BeginPath = function() {
    this.context.beginPath();
};
CPDFGraphics.prototype.ClosePath = function() {
    this.context.closePath();
};
CPDFGraphics.prototype.Stroke = function() {
    this.context.stroke();
};
CPDFGraphics.prototype.MoveTo = function(x, y) {
    let nScale = this.GetScale();
    
    this.CheckPoint(x, y);

    this.context.moveTo(x * nScale, y * nScale);
};
CPDFGraphics.prototype.LineTo = function(x, y) {
    let nScale = this.GetScale();

    this.CheckPoint(x, y);

    this.context.lineTo(x * nScale, y * nScale);
};
CPDFGraphics.prototype.FillRect = function(x, y, w, h) {
    let nScale = this.GetScale();

    this.CheckPoint(x, y);
    this.CheckPoint(x + w, y + h);

    this.context.beginPath();
    this.context.fillRect(x * nScale, y * nScale, w * nScale, h * nScale);
};
	CPDFGraphics.prototype.DrawImageXY = function(image, dx, dy) {
		
		let scale       = this.GetScale();
		let integerGrid = this.GetIntegerGrid();
		
		this.CheckPoint(dx, dy);
		this.CheckPoint(dx + image.width, dy + image.height);
		
		dx *= scale;
		if (integerGrid)
			dx = dx + 0.5 >> 0;
		
		dy *= scale;
		if (integerGrid)
			dy = dy + 0.5 >> 0;
		
		this.context.drawImage(image, dx, dy);
	};
CPDFGraphics.prototype.DrawImage = function(image, sx, sy, sWidth, sHeight, dx, dy, dWidth, dHeight) {
    let nScale          = this.GetScale();
    let bIntegerGrid    = this.GetIntegerGrid();

    this.CheckPoint(dx, dy);
    this.CheckPoint(dx + dWidth, dy + dHeight);

    if (sx != null) {
        sx *= nScale;
        if (bIntegerGrid)
            sx = sx + 0.5 >> 0;
    }
    if (sy != null) {
        sy *= nScale;
        if (bIntegerGrid)
            sy = sy + 0.5 >> 0;
    }
    if (sWidth != null) {
        sWidth *= nScale;
        if (bIntegerGrid)
            sWidth = sWidth + 0.5 >> 0;
    }
    if (sHeight != null) {
        sHeight *= nScale;
        if (bIntegerGrid)
            sHeight = sHeight + 0.5 >> 0;
    }
    if (dx != null) {
        dx *= nScale;
        if (bIntegerGrid)
            dx = dx + 0.5 >> 0;
    }
    if (dy != null) {
        dy *= nScale;
        if (bIntegerGrid)
            dy = dy + 0.5 >> 0;
    }
    if (dWidth != null) {
        dWidth *= nScale;
        if (bIntegerGrid)
            dWidth = dWidth + 0.5 >> 0;
    }
    if (dHeight != null) {
        dHeight *= nScale;
        if (bIntegerGrid)
            dHeight = dHeight + 0.5 >> 0;
    }
    this.context.drawImage(image, sx, sy, sWidth, sHeight, dx, dy, dWidth, dHeight);
};
CPDFGraphics.prototype.SetLineDash = function(dash) {
    if (this.context)
        this.context.setLineDash(dash);
};
CPDFGraphics.prototype.SetGlobalAlpha = function(value) {
    if (this.context)
        this.context.globalAlpha = value;
};

CPDFGraphics.prototype.Arc = function(x, y, radius, startAng, endAng, counterClockwise) {
    let nScale = this.GetScale();

    this.CheckPoint(x - radius, y - radius);
    this.CheckPoint(x - radius, y + radius);
    this.CheckPoint(x + radius, y - radius);
    this.CheckPoint(x + radius, y + radius);

    this.context.arc(x * nScale, y * nScale, radius * nScale, startAng, endAng, counterClockwise);
};
CPDFGraphics.prototype.ArcTo = function(x1, y1, x2, y2, r) {
    let nScale = this.GetScale();

    this.context.arcTo(x1 * nScale, y1 * nScale, x2 * nScale, y2 * nScale, r * nScale);
};
CPDFGraphics.prototype.Fill = function() {
    this.context.fill();
};
CPDFGraphics.prototype.ClearRect = function(x, y, w, h) {
    let nScale = this.GetScale();

    this.context.clearRect(x * nScale, y * nScale, w * nScale, h * nScale);
};
CPDFGraphics.prototype.HorLine = function(x1, x2, y)
{
    let nScale          = this.GetScale();
    let nLineW          = this.GetLineWidth(true);
    
    this.CheckPoint(x1, y);
    this.CheckPoint(x2, y);

    let X1  = x1 * nScale >> 0;
    let X2  = x2 * nScale >> 0;
    let Y   = y * nScale >> 0;
    
    let nLineOffsetY = (0 === (nLineW % 2) ? 0 : 0.5);

    this.context.moveTo(X1, nLineOffsetY + Y);
    this.context.lineTo(X2, nLineOffsetY + Y);
};
CPDFGraphics.prototype.VerLine = function(y1, y2, x)
{
    let nScale          = this.GetScale();
    let nLineW          = this.GetLineWidth(true);
    
    this.CheckPoint(x, y1);
    this.CheckPoint(x, y2);

    let Y1  = y1 * nScale >> 0;
    let Y2  = y2 * nScale >> 0;
    let X   = x * nScale >> 0;
    
    let nLineOffsetX = (0 === (nLineW % 2) ? 0 : 0.5);

    this.context.moveTo(nLineOffsetX + X, Y1);
    this.context.lineTo(nLineOffsetX + X, Y2);
};

    //------------------------------------------------------------export----------------------------------------------------
    window['AscPDF'] = window['AscPDF'] || {};
    window['AscPDF'].CPDFGraphics = CPDFGraphics;
})(window);
