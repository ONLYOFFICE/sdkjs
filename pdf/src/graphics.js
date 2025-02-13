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

    function CPDFGraphics()
    {
        this.m_oContext         = null;
        this.m_lWidthPix        = undefined;
        this.m_lHeightPix       = undefined;
        this.m_dPageWidthPix    = undefined;
        this.m_dPageHeightPix   = undefined;
        this.lineWidth          = 1;
        this.fillStyle          = null;
        this.strokeStyle        = null;
        this.globalAlpha        = 1;
        this.bIntegerGrid       = false;

        this.m_oCoordTransform = new AscCommon.CMatrix();
    }
    CPDFGraphics.prototype.SetCurPage = function(nPage) {
        this.curPage = nPage;
    };
    CPDFGraphics.prototype.GetCurPage = function() {
        return this.curPage;
    };
    CPDFGraphics.prototype.GetDrawingPageW = function() {
        return this.m_lWidthPix;
    };
    CPDFGraphics.prototype.GetDrawingPageH = function() {
        return this.m_lHeightPix;
    };
    CPDFGraphics.prototype.Clip = function() {
        this.m_oContext.clip();
    };
    CPDFGraphics.prototype.SetIntegerGrid = function(bInteger) {
        this.bIntegerGrid = bInteger;

        if (true == bInteger) {
            this.bIntegerGrid = true;
            this.m_oContext.setTransform(1,0,0,1,0,0);
        }
        else {
            this.bIntegerGrid = false;
        }
    };
    CPDFGraphics.prototype.GetIntegerGrid = function() {
        return this.bIntegerGrid;
    };

    CPDFGraphics.prototype.SetStrokeStyle = function(r,g,b,a) {
        if (a == undefined) {
            a = 255;
        }

        if (this.m_oContext)
            this.m_oContext.strokeStyle = "rgb(" + r + "," + g + "," + b + "," + a / 255 +")";

        this.strokeStyle = {
            r: r,
            g: g,
            b: b
        };
    };
    CPDFGraphics.prototype.SetFillStyle = function(r,g,b) {
        if (this.m_oContext)
            this.m_oContext.fillStyle = "rgb(" + r + "," + g + "," + b + ")";

        this.fillStyle = {
            r: r,
            g: g,
            b: b
        };
    };
    CPDFGraphics.prototype.SetLineWidth = function(width) {
        if (this.m_oContext)
            this.m_oContext.lineWidth = width + 0.5 >> 0;
            
        this.lineWidth = width;
    };
    CPDFGraphics.prototype.GetLineWidth = function(bTransform) {
        if (!bTransform)
            return this.lineWidth;

        let tr = this.GetTransform();
        return this.lineWidth * tr.sy + 0.5 >> 0;
    };
    CPDFGraphics.prototype.Init = function(context, nWidthPx, nHeightPx, nPageWidthPx, nPageHeightPx) {
        this.m_oContext     = context;
        this.m_lWidthPix    = nWidthPx;
        this.m_lHeightPix   = nHeightPx;

        this.m_dPageWidthPix    = nPageWidthPx;
        this.m_dPageHeightPix   = nPageHeightPx;

        this.m_oCoordTransform.sx   = this.m_lWidthPix / this.m_dPageWidthPix;
        this.m_oCoordTransform.sy   = this.m_lHeightPix / this.m_dPageHeightPix;
    };
    CPDFGraphics.prototype.GetTransform = function() {
        return this.m_oCoordTransform;
    };
    CPDFGraphics.prototype.GetContext = function() {
        return this.m_oContext;
    }
    CPDFGraphics.prototype.Rect = function(x, y, w, h) {
        let ctx = this.GetContext();
        ctx.beginPath();

        if (this.GetIntegerGrid())
        {
            let tr = this.GetTransform();
            if (0.0 === tr.shx && 0.0 === tr.shy)
            {
                let _x = (tr.TransformPointX(x, y) + 0.5) >> 0;
                let _y = (tr.TransformPointY(x, y) + 0.5) >> 0;
                let _r = (tr.TransformPointX(x + w, y) + 0.5) >> 0;
                let _b = (tr.TransformPointY(x, y + h) + 0.5) >> 0;

                ctx.rect(_x, _y, _r - _x, _b - _y);
            }
            else
            {
                let x1 = tr.TransformPointX(x, y);
                let y1 = tr.TransformPointY(x, y);
                let x2 = tr.TransformPointX(x + w, y);
                let y2 = tr.TransformPointY(x + w, y);
                let x3 = tr.TransformPointX(x + w, y + h);
                let y3 = tr.TransformPointY(x + w, y + h);
                let x4 = tr.TransformPointX(x, y + h);
                let y4 = tr.TransformPointY(x, y + h);

                ctx.moveTo(x1, y1);
                ctx.lineTo(x2, y2);
                ctx.lineTo(x3, y3);
                ctx.lineTo(x4, y4);
                ctx.closePath();
            }
        }
        else {
            ctx.save();
            this.EnableTransform();
            ctx.rect(x, y, w, h);
            ctx.restore();
        }
    };
    CPDFGraphics.prototype.EnableTransform = function() {
        let ctx = this.GetContext();
        let tr = this.GetTransform();
        
        ctx.setTransform(
            tr.sx,  // масштаб по оси X
            tr.shy, // наклон по оси Y
            tr.shx, // наклон по оси X
            tr.sy,  // масштаб по оси Y
            tr.tx,  // сдвиг по оси X
            tr.ty   // сдвиг по оси Y
        );
    }
    CPDFGraphics.prototype.BeginPath = function() {
        this.m_oContext.beginPath();
    };
    CPDFGraphics.prototype.ClosePath = function() {
        this.m_oContext.closePath();
    };
    CPDFGraphics.prototype.Stroke = function() {
        let ctx = this.GetContext();
        let tr = this.GetTransform();
        let nLineW = this.lineWidth;

        let _nLineW = nLineW * tr.sy + 0.5 >> 0;
        
        ctx.lineWidth = _nLineW;
        ctx.stroke();
        ctx.lineWidth = nLineW;
    };
    CPDFGraphics.prototype.MoveTo = function(x, y) {
        let ctx = this.GetContext();

        if (this.GetIntegerGrid()) {
            let tr = this.GetTransform();
            var _x = (tr.TransformPointX(x,y)) >> 0;
            var _y = (tr.TransformPointY(x,y)) >> 0;
            ctx.moveTo(_x + 0.5,_y + 0.5);
        }
        else {
            ctx.save();
            this.EnableTransform();
            ctx.moveTo(x,y);
            ctx.restore();
        }  
    };
    CPDFGraphics.prototype.LineTo = function(x, y) {
        let ctx = this.GetContext();

        if (this.GetIntegerGrid()) {
            let tr = this.GetTransform();
            var _x = (tr.TransformPointX(x,y)) >> 0;
            var _y = (tr.TransformPointY(x,y)) >> 0;
            ctx.lineTo(_x + 0.5,_y + 0.5);
        }
        else {
            ctx.save();
            this.EnableTransform();
            ctx.lineTo(x,y);
            ctx.restore();
        }
    };
    CPDFGraphics.prototype.FillRect = function(x, y, w, h) {
        var ctx = this.m_oContext;

        if (this.GetIntegerGrid()) {
            let tr = this.GetTransform();

            var _x = (tr.TransformPointX(x,y) >> 0) + 0.5;
            var _y = (tr.TransformPointY(x,y) >> 0) + 0.5;
            var _r = (tr.TransformPointX(x+w,y) >> 0) + 0.5;
            var _b = (tr.TransformPointY(x,y+h) >> 0) + 0.5;

            ctx.fillRect(_x - 0.5, _y - 0.5, _r - _x + 1, _b - _y + 1);
        }
        else {
            ctx.save();
            this.EnableTransform();
            ctx.fillRect(x, y, w, h);
            ctx.restore();
        }
    };
    CPDFGraphics.prototype.DrawImageXY = function(image, dx, dy, rot, minus1) {
        let tr = this.GetTransform();

        let _x = Math.round(tr.TransformPointX(dx, dy));
        let _y = Math.round(tr.TransformPointY(dx, dy));

        if (minus1)
        {
            _x--;
            _y--;
        }

        let context = this.m_oContext;

        context.save();
        context.translate(_x, _y);

        if (rot && rot !== 0) {
            let W = image.width;
            let H = image.height;
            let rotW = Math.round(W / 2);
            let rotH = Math.round(H / 2);

            context.translate(rotW, rotH);
            context.rotate(rot);
            context.translate(-rotW, -rotH);
        }

        context.drawImage(image, 0, 0);
        context.restore();
    };
    CPDFGraphics.prototype.DrawImage = function(image, sx, sy, sWidth, sHeight, dx, dy, dWidth, dHeight) {
        let ctx = this.GetContext();
        let tr  = this.GetTransform();
        
        ctx.save();

        ctx.setTransform(
            tr.sx,  // масштаб по оси X
            tr.shy, // наклон по оси Y
            tr.shx, // наклон по оси X
            tr.sy,  // масштаб по оси Y
            tr.tx,  // сдвиг по оси X
            tr.ty   // сдвиг по оси Y
        );

        ctx.drawImage(image, sx, sy, sWidth, sHeight, dx, dy, dWidth, dHeight);

        ctx.restore();
    };
    CPDFGraphics.prototype.SetLineDash = function(dash) {
        let tr = this.GetTransform();
        let ctx = this.GetContext();
        ctx.setLineDash(dash.map(function(measure) {
            return measure * tr.sy + 0.5 >> 0;
        }));
    };
    CPDFGraphics.prototype.SetGlobalAlpha = function(value) {
        if (this.m_oContext)
            this.m_oContext.globalAlpha = value;
    };

    CPDFGraphics.prototype.Arc = function(x, y, radius, startAng, endAng, counterClockwise) {
        let ctx = this.GetContext();

        ctx.save();
        this.EnableTransform();
        ctx.arc(x, y, radius, startAng, endAng, counterClockwise);
        ctx.restore();
    };
    CPDFGraphics.prototype.ArcTo = function(x1, y1, x2, y2, r) {
        let ctx = this.GetContext();

        ctx.save();
        this.EnableTransform();
        ctx.arcTo(x1, y1, x2, y2, r);
        ctx.restore();
    };
    CPDFGraphics.prototype.Fill = function() {
        let ctx = this.GetContext();

        ctx.save();
        this.EnableTransform();
        ctx.fill();
        ctx.restore();
    };
    CPDFGraphics.prototype.ClearRect = function(x, y, w, h) {
        let ctx = this.GetContext();

        ctx.save();
        this.EnableTransform();
        ctx.clearRect(x, y, w, h);
        ctx.restore();
    };
    CPDFGraphics.prototype.HorLine = function(x1, x2, y) {
        let ctx = this.GetContext();

        let nLineW = this.GetLineWidth(true);

        let tr = this.GetTransform();
        
        var _x1 = (tr.TransformPointX(x1,y)) >> 0;
        var _x2 = (tr.TransformPointX(x2,y)) >> 0;
        var _y  = (tr.TransformPointY(x1,y)) >> 0;

        let nLineOffsetY = (0 === (nLineW % 2) ? 0 : 0.5);

        ctx.moveTo(_x1, _y + nLineOffsetY);
        ctx.lineTo(_x2, _y + nLineOffsetY);
    };
    CPDFGraphics.prototype.VerLine = function(y1, y2, x) {
        let ctx = this.GetContext();

        let nLineW = this.GetLineWidth(true);

        let tr = this.GetTransform();
        
        var _y1 = (tr.TransformPointX(x,y1)) >> 0;
        var _y2 = (tr.TransformPointX(x,y2)) >> 0;
        var _x  = (tr.TransformPointY(y1,x)) >> 0;

        let nLineOffsetX = (0 === (nLineW % 2) ? 0 : 0.5);

        ctx.moveTo(nLineOffsetX + _x, _y1);
        ctx.lineTo(nLineOffsetX + _x, _y2);
    };
    CPDFGraphics.prototype.DrawLockObjectRect = function(lock_type, aRegions) {
        if (Asc.editor.isViewMode || this.IsThumbnail || this.IsDemonstrationMode || lock_type === AscCommon.c_oAscLockTypes.kLockTypeNone)
            return;
    
        if (Asc.editor.WordControl) {
            if (lock_type === AscCommon.c_oAscLockTypes.kLockTypeMine && true === AscCommon.CollaborativeEditing.Is_Fast())
                return;
            if (Asc.editor.WordControl.m_oDrawingDocument.IsLockObjectsEnable === false && lock_type === AscCommon.c_oAscLockTypes.kLockTypeMine)
                return;
        }
    
        this.SetGlobalAlpha(1);
        if (lock_type === AscCommon.c_oAscLockTypes.kLockTypeMine) {
            this.SetStrokeStyle(22, 156, 0, 255);
        } else {
            this.SetStrokeStyle(238, 53, 37, 255);
        }
    
        let ctx = this.m_oContext;
        let bIsInt = this.m_bIntegerGrid;
        if (!bIsInt)
            this.SetIntegerGrid(true);
    
        ctx.lineWidth = 1;
    
        let eps = 5; // Коэффициент расширения
    
        let tr = this.GetTransform();
    
        ctx.beginPath();
        for (let i = 0; i < aRegions.length; i++) {
            let aPoints = aRegions[i]; // Массив точек [[x1, y1], [x2, y2], ..., [xn, yn]]
            let expandedPoints = [];
    
            let numPoints = aPoints.length;
            for (let j = 0; j < numPoints; j++) {
                let current = aPoints[j];
                let next = aPoints[(j + 1) % numPoints];
    
                // Векторы текущего и следующего ребра
                let dx1 = next[0] - current[0];
                let dy1 = next[1] - current[1];
                let dx2 = aPoints[(j + 2) % numPoints][0] - next[0];
                let dy2 = aPoints[(j + 2) % numPoints][1] - next[1];
    
                // Нормали к ребрам
                let len1 = Math.sqrt(dx1 * dx1 + dy1 * dy1);
                let len2 = Math.sqrt(dx2 * dx2 + dy2 * dy2);
                let nx1 = dy1 / len1;
                let ny1 = -dx1 / len1;
                let nx2 = dy2 / len2;
                let ny2 = -dx2 / len2;
    
                // Усредненная нормаль
                let avgNx = (nx1 + nx2) / 2;
                let avgNy = (ny1 + ny2) / 2;
    
                // Нормализуем усредненную нормаль
                let avgLen = Math.sqrt(avgNx * avgNx + avgNy * avgNy);
                if (avgLen === 0) {
                    avgNx = 0;
                    avgNy = 0;
                } else {
                    avgNx /= avgLen;
                    avgNy /= avgLen;
                }
    
                // Смещение точки
                let x_expanded = next[0] + avgNx * eps;
                let y_expanded = next[1] + avgNy * eps;
    
                // Преобразуем координаты
                let tx = tr.TransformPointX(x_expanded, y_expanded);
                let ty = tr.TransformPointY(x_expanded, y_expanded);
    
                expandedPoints.push([tx, ty]);
            }
    
            // Рисуем пунктир вручную
            let w_dot = 2;  // Длина штриха
            let w_dist = 2; // Расстояние между штрихами
    
            for (let j = 0; j < expandedPoints.length; j++) {
                let start = expandedPoints[j];
                let end = expandedPoints[(j + 1) % expandedPoints.length];
    
                let dx = end[0] - start[0];
                let dy = end[1] - start[1];
                let length = Math.sqrt(dx * dx + dy * dy);
    
                let stepX = (dx / length) * w_dist;
                let stepY = (dy / length) * w_dist;
    
                let x = start[0];
                let y = start[1];
                let dist = 0;
    
                while (dist < length) {
                    ctx.moveTo(x, y);
                    x += (dx / length) * w_dot;
                    y += (dy / length) * w_dot;
                    dist += w_dot;
    
                    if (dist > length) {
                        x = end[0];
                        y = end[1];
                    }
    
                    ctx.lineTo(x, y);
    
                    x += stepX;
                    y += stepY;
                    dist += w_dist;
                }
            }
    
            ctx.stroke();
        }
    
        if (!bIsInt)
            this.SetIntegerGrid(false);
    };
    
    

    //------------------------------------------------------------export----------------------------------------------------
    window['AscPDF'] = window['AscPDF'] || {};
    window['AscPDF'].CPDFGraphics = CPDFGraphics;
})(window);
