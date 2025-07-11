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
var getFullImageSrc2 = AscCommon.getFullImageSrc2;

var CShapeColor = AscFormat.CShapeColor;

var c_oAscFill = Asc.c_oAscFill;

function DrawLineEnd(xEnd, yEnd, xPrev, yPrev, type, w, len, drawer, trans)
{
    switch (type)
    {
        case AscFormat.LineEndType.None:
            break;
        case AscFormat.LineEndType.Arrow:
        {
            if (Asc.editor.isPdfEditor() == true) {
                drawer.CheckDash();
            }

            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd + len * _ex;
            var tmpy = yEnd + len * _ey;

            var x1 = tmpx + _vx * w/2;
            var y1 = tmpy + _vy * w/2;

            var x3 = tmpx - _vx * w/2;
            var y3 = tmpy - _vy * w/2;

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer.ds();
            drawer._e();

            break;
        }
        case AscFormat.LineEndType.ReverseArrow:
        {
            if (Asc.editor.isPdfEditor() == true) {
                drawer.CheckDash();
            }

            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd - len * _ex;
            var tmpy = yEnd - len * _ey;

            var x1 = tmpx + _vx * w/2;
            var y1 = tmpy + _vy * w/2;

            var x3 = tmpx - _vx * w/2;
            var y3 = tmpy - _vy * w/2;

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer.ds();
            drawer._e();

            break;
        }
        case AscFormat.LineEndType.Diamond:
        {
            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd + len/2 * _ex;
            var tmpy = yEnd + len/2 * _ey;

            var x1 = xEnd + _vx * w/2;
            var y1 = yEnd + _vy * w/2;

            var x3 = xEnd - _vx * w/2;
            var y3 = yEnd - _vy * w/2;

            var tmpx2 = xEnd - len/2 * _ex;
            var tmpy2 = yEnd - len/2 * _ey;

            drawer._s();
            drawer._m(trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));
            drawer._l(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(tmpx2, tmpy2), trans.TransformPointY(tmpx2, tmpy2));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._z();
            if (Asc.editor.isPdfEditor()) {
                let oRGBColor;
                if (drawer.Shape.GetRGBColor) {
                    oRGBColor = drawer.Shape.GetRGBColor(drawer.Shape.GetFillColor());
                }
                else if (drawer.Shape.group) {
                    oRGBColor = drawer.Shape.group.GetRGBColor(drawer.Shape.group.GetFillColor());
                }

                if (oRGBColor) {
                    drawer.Graphics.m_oPen.Color.R = oRGBColor.r;
                    drawer.Graphics.m_oPen.Color.G = oRGBColor.g;
                    drawer.Graphics.m_oPen.Color.B = oRGBColor.b;
                }
            }
            drawer.drawStrokeFillStyle();
            drawer._e();

            drawer._s();
            drawer._m(trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));
            drawer._l(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(tmpx2, tmpy2), trans.TransformPointY(tmpx2, tmpy2));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._z();
            drawer.ds();
            drawer._e();

            break;
        }
        case AscFormat.LineEndType.Square:
        {
            var angle = Math.atan2(yEnd - yPrev, xEnd - xPrev);
            function rotatePoints(aPoints) {
                // Поворачиваем каждую вершину вокруг центра
                for (var i = 0; i < aPoints.length; i++) {
                    var x = aPoints[i].x - xEnd;
                    var y = aPoints[i].y - yEnd;

                    // Применяем матрицу поворота
                    var rotatedX = x * Math.cos(angle) - y * Math.sin(angle);
                    var rotatedY = x * Math.sin(angle) + y * Math.cos(angle);

                    // Возвращаем вершины на место
                    aPoints[i].x = rotatedX + xEnd;
                    aPoints[i].y = rotatedY + yEnd;
                }
            }

            var x1 = xEnd - w/2;
            var y1 = yEnd + w/2;

            var x2 = xEnd - w/2;
            var y2 = yEnd - w/2;

            var x3 = xEnd + w/2;
            var y3 = yEnd - w/2;

            var x4 = xEnd + w/2;
            var y4 = yEnd + w/2;

            let aSmall = [
                { x: x1, y: y1 },
                { x: x2, y: y2 },
                { x: x3, y: y3 },
                { x: x4, y: y4 }
            ]

            rotatePoints(aSmall);
            
            drawer._s();
            drawer._m(trans.TransformPointX(aSmall[0].x, aSmall[0].y), trans.TransformPointY(aSmall[0].x, aSmall[0].y));
            for (var i = 1; i < aSmall.length; i++) {
                drawer._l(trans.TransformPointX(aSmall[i].x, aSmall[i].y), trans.TransformPointY(aSmall[i].x, aSmall[i].y));
            }
            drawer._z();
            if (Asc.editor.isPdfEditor()) {
                let oRGBColor;
                if (drawer.Shape.GetRGBColor) {
                    oRGBColor = drawer.Shape.GetRGBColor(drawer.Shape.GetFillColor());
                }
                else if (drawer.Shape.group) {
                    oRGBColor = drawer.Shape.group.GetRGBColor(drawer.Shape.group.GetFillColor());
                }

                if (oRGBColor) {
                    drawer.Graphics.m_oPen.Color.R = oRGBColor.r;
                    drawer.Graphics.m_oPen.Color.G = oRGBColor.g;
                    drawer.Graphics.m_oPen.Color.B = oRGBColor.b;
                }
            }
            drawer.drawStrokeFillStyle();
            drawer._e();

            x1 = xEnd - w * 2/4;
            y1 = yEnd + w * 2/4;

            x2 = xEnd - w * 2/4;
            y2 = yEnd - w * 2/4;

            x3 = xEnd + w * 2/4;
            y3 = yEnd - w * 2/4;

            x4 = xEnd + w * 2/4;
            y4 = yEnd + w * 2/4;

            let aBig = [
                { x: x1, y: y1 },
                { x: x2, y: y2 },
                { x: x3, y: y3 },
                { x: x4, y: y4 }
            ]

            rotatePoints(aBig);
            
            drawer._s();
            drawer._m(trans.TransformPointX(aBig[0].x, aBig[0].y), trans.TransformPointY(aBig[0].x, aBig[0].y));
            for (var i = 1; i < aBig.length; i++) {
                drawer._l(trans.TransformPointX(aBig[i].x, aBig[i].y), trans.TransformPointY(aBig[i].x, aBig[i].y));
            }
            drawer._z();
            drawer.ds();
            drawer._e();
            break;
        }
        case AscFormat.LineEndType.Oval:
        {
            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd + len/2 * _ex;
            var tmpy = yEnd + len/2 * _ey;

            var tmpx2 = xEnd - len/2 * _ex;
            var tmpy2 = yEnd - len/2 * _ey;

            var cx1 = tmpx + _vx * 3*w/4;
            var cy1 = tmpy + _vy * 3*w/4;
            var cx2 = tmpx2 + _vx * 3*w/4;
            var cy2 = tmpy2 + _vy * 3*w/4;

            var cx3 = tmpx - _vx * 3*w/4;
            var cy3 = tmpy - _vy * 3*w/4;
            var cx4 = tmpx2 - _vx * 3*w/4;
            var cy4 = tmpy2 - _vy * 3*w/4;

            drawer._s();
            drawer._m(trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));
            drawer._c(trans.TransformPointX(cx1, cy1), trans.TransformPointY(cx1, cy1),
                trans.TransformPointX(cx2, cy2), trans.TransformPointY(cx2, cy2),
                trans.TransformPointX(tmpx2, tmpy2), trans.TransformPointY(tmpx2, tmpy2));

            drawer._c(trans.TransformPointX(cx4, cy4), trans.TransformPointY(cx4, cy4),
                trans.TransformPointX(cx3, cy3), trans.TransformPointY(cx3, cy3),
                trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));

            if (Asc.editor.isPdfEditor()) {
                let oRGBColor;
                if (drawer.Shape.GetRGBColor) {
                    oRGBColor = drawer.Shape.GetRGBColor(drawer.Shape.GetFillColor());
                }
                else if (drawer.Shape.group) {
                    oRGBColor = drawer.Shape.group.GetRGBColor(drawer.Shape.group.GetFillColor());
                }

                if (oRGBColor) {
                    drawer.Graphics.m_oPen.Color.R = oRGBColor.r;
                    drawer.Graphics.m_oPen.Color.G = oRGBColor.g;
                    drawer.Graphics.m_oPen.Color.B = oRGBColor.b;
                }
            }
            drawer.drawStrokeFillStyle();
            drawer._e();

            drawer._s();
            drawer._m(trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));
            drawer._c(trans.TransformPointX(cx1, cy1), trans.TransformPointY(cx1, cy1),
                trans.TransformPointX(cx2, cy2), trans.TransformPointY(cx2, cy2),
                trans.TransformPointX(tmpx2, tmpy2), trans.TransformPointY(tmpx2, tmpy2));

            drawer._c(trans.TransformPointX(cx4, cy4), trans.TransformPointY(cx4, cy4),
                trans.TransformPointX(cx3, cy3), trans.TransformPointY(cx3, cy3),
                trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));

            drawer.ds();
            drawer._e();
            break;
        }
        case AscFormat.LineEndType.Stealth:
        {
            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd + len * _ex;
            var tmpy = yEnd + len * _ey;

            var x1 = tmpx + _vx * w/2;
            var y1 = tmpy + _vy * w/2;

            var x3 = tmpx - _vx * w/2;
            var y3 = tmpy - _vy * w/2;

            var x4 = xEnd + (len - w/2) * _ex;
            var y4 = yEnd + (len - w/2) * _ey;

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._l(trans.TransformPointX(x4, y4), trans.TransformPointY(x4, y4));
            drawer._z();
            drawer.drawStrokeFillStyle();
            drawer._e();

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._l(trans.TransformPointX(x4, y4), trans.TransformPointY(x4, y4));
            drawer._z();
            drawer.ds();
            drawer._e();

            break;
        }
        case AscFormat.LineEndType.ReverseTriangle:
        {
            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd - len * _ex;
            var tmpy = yEnd - len * _ey;

            var x1 = tmpx + _vx * w/2;
            var y1 = tmpy + _vy * w/2;

            var x3 = tmpx - _vx * w/2;
            var y3 = tmpy - _vy * w/2;

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._z();
            if (Asc.editor.isPdfEditor()) {
                let oRGBColor;
                if (drawer.Shape.GetRGBColor) {
                    oRGBColor = drawer.Shape.GetRGBColor(drawer.Shape.GetFillColor());
                }
                else if (drawer.Shape.group) {
                    oRGBColor = drawer.Shape.group.GetRGBColor(drawer.Shape.group.GetFillColor());
                } 

                if (oRGBColor) {
                    drawer.Graphics.m_oPen.Color.R = oRGBColor.r;
                    drawer.Graphics.m_oPen.Color.G = oRGBColor.g;
                    drawer.Graphics.m_oPen.Color.B = oRGBColor.b;
                }
            }

            drawer.drawStrokeFillStyle();
            drawer._e();

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._z();
            drawer.ds();
            drawer._e();
            break;
        }
        case AscFormat.LineEndType.Triangle:
        {
            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd + len * _ex;
            var tmpy = yEnd + len * _ey;

            var x1 = tmpx + _vx * w/2;
            var y1 = tmpy + _vy * w/2;

            var x3 = tmpx - _vx * w/2;
            var y3 = tmpy - _vy * w/2;

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._z();
            if (Asc.editor.isPdfEditor()) {
                let oRGBColor;
                if (drawer.Shape.GetRGBColor) {
                    oRGBColor = drawer.Shape.GetRGBColor(drawer.Shape.GetFillColor());
                }
                else if (drawer.Shape.group) {
                    oRGBColor = drawer.Shape.group.GetRGBColor(drawer.Shape.group.GetFillColor());
                }

                if (oRGBColor) {
                    drawer.Graphics.m_oPen.Color.R = oRGBColor.r;
                    drawer.Graphics.m_oPen.Color.G = oRGBColor.g;
                    drawer.Graphics.m_oPen.Color.B = oRGBColor.b;
                }
            }

            drawer.drawStrokeFillStyle();
            drawer._e();

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
            drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._z();
            drawer.ds();
            drawer._e();
            break;
        }
        case AscFormat.LineEndType.Butt:
        {
            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd + len * _ex;
            var tmpy = yEnd + len * _ey;

            var angle = Math.atan2(yEnd - yPrev, xEnd - xPrev);
            // Вычисляем координаты конца перпендикулярной линии
            var perpendicularLength = w;
            var x1 = xEnd + perpendicularLength * Math.cos(angle - Math.PI / 2);
            var y1 = yEnd + perpendicularLength * Math.sin(angle - Math.PI / 2);
            var x2 = xEnd - perpendicularLength * Math.cos(angle - Math.PI / 2);
            var y2 = yEnd - perpendicularLength * Math.sin(angle - Math.PI / 2);

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));
            drawer.ds();
            break;
        }
        case AscFormat.LineEndType.Slash:
        {
            var _ex = xPrev - xEnd;
            var _ey = yPrev - yEnd;
            var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
            _ex /= _elen;
            _ey /= _elen;

            var _vx = _ey;
            var _vy = -_ex;

            var tmpx = xEnd + len * _ex;
            var tmpy = yEnd + len * _ey;

            var angle = Math.atan2(yEnd - yPrev, xEnd - xPrev) + (30 * Math.PI / 180);

            // Вычисляем координаты конца перпендикулярной линии
            var perpendicularLength = w;
            var x1 = xEnd + perpendicularLength * Math.cos(angle - Math.PI / 2);
            var y1 = yEnd + perpendicularLength * Math.sin(angle - Math.PI / 2);
            var x2 = xEnd - perpendicularLength * Math.cos(angle - Math.PI / 2);
            var y2 = yEnd - perpendicularLength * Math.sin(angle - Math.PI / 2);

            drawer._s();
            drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
            drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));
            drawer.ds();
            break;
        }
        // visio types are handled below
    }

    if (type === AscFormat.LineEndType.vsdxOpenASMEArrow ||
        type === AscFormat.LineEndType.vsdxFilledASMEArrow ||
        type === AscFormat.LineEndType.vsdxClosedASMEArrow) {
        len *= 1.5;
        let isFilled = type === AscFormat.LineEndType.vsdxFilledASMEArrow;
        let isOpen = type === AscFormat.LineEndType.vsdxOpenASMEArrow;
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, isFilled, isOpen, w, len);
    } else if (type === AscFormat.LineEndType.vsdxOpenSharpArrow ||
        type === AscFormat.LineEndType.vsdxFilledSharpArrow ||
        type === AscFormat.LineEndType.vsdxClosedSharpArrow) {
        let isFilled = type === AscFormat.LineEndType.vsdxFilledSharpArrow;
        let isOpen = type === AscFormat.LineEndType.vsdxOpenSharpArrow;
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, isFilled, isOpen, w, len);
    } else if (type === AscFormat.LineEndType.vsdxOpen90Arrow ||
        type === AscFormat.LineEndType.vsdxFilled90Arrow ||
        type === AscFormat.LineEndType.vsdxClosed90Arrow) {
        len /= 2;
        let isFilled = type === AscFormat.LineEndType.vsdxFilled90Arrow;
        let isOpen = type === AscFormat.LineEndType.vsdxOpen90Arrow;
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, isFilled, isOpen, w, len);
    } else if (type === AscFormat.LineEndType.vsdxIndentedFilledArrow ||
        type === AscFormat.LineEndType.vsdxIndentedClosedArrow ||
        type === AscFormat.LineEndType.vsdxOutdentedFilledArrow ||
        type === AscFormat.LineEndType.vsdxOutdentedClosedArrow) {
        let isArrowFilled = type === AscFormat.LineEndType.vsdxIndentedFilledArrow ||
            type === AscFormat.LineEndType.vsdxOutdentedFilledArrow;

        let isIndented = type === AscFormat.LineEndType.vsdxIndentedFilledArrow ||
            type === AscFormat.LineEndType.vsdxIndentedClosedArrow;

        var _ex = xPrev - xEnd;
        var _ey = yPrev - yEnd;
        var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
        _ex /= _elen; // cos a
        _ey /= _elen; // sin a

        // a now is 90 - a + invert cos
        var _vx = _ey;
        var _vy = -_ex;

        // (xEnd, yEnd) - right arrow point
        var tmpx = xEnd + len * _ex;
        var tmpy = yEnd + len * _ey;

        // (x1, y1) - top arrow point
        var x1 = tmpx + _vx * w/2;
        var y1 = tmpy + _vy * w/2;

        // (x3, y3) - bottom arrow point
        var x3 = tmpx - _vx * w/2;
        var y3 = tmpy - _vy * w/2;

        let controlPointShift = isIndented? 0.75 * len : 1.25 * len;

        var x4 = xEnd + controlPointShift * _ex;
        var y4 = yEnd + controlPointShift * _ey;

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
        drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
        let cpxLocal = trans.TransformPointX(x4, y4);
        let cpyLocal = trans.TransformPointY(x4, y4);
        let endXLocal = trans.TransformPointX(x1, y1);
        let endYLocal = trans.TransformPointY(x1, y1);
        drawer._c2(cpxLocal, cpyLocal, endXLocal, endYLocal);
        drawer._z();

        stokeOrFillPath(drawer, isArrowFilled);

        drawer._e();
    } else if (type === AscFormat.LineEndType.vsdxOpenFletch ||
        type === AscFormat.LineEndType.vsdxFilledFletch ||
        type === AscFormat.LineEndType.vsdxClosedFletch) {
        let isArrowFilled = false;
        let isArrowClosed = false;
        if (type === AscFormat.LineEndType.vsdxFilledFletch) {
            isArrowClosed = true;
            isArrowFilled = true;
        } else if (type === AscFormat.LineEndType.vsdxClosedFletch) {
            isArrowFilled = false;
            isArrowClosed = true;
        }
        var _ex = xPrev - xEnd;
        var _ey = yPrev - yEnd;
        var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
        _ex /= _elen; // cos a
        _ey /= _elen; // sin a

        // a now is 90 - a + invert cos
        var _vx = _ey;
        var _vy = -_ex;

        // (xEnd, yEnd) - right arrow point

        // left arrow point?
        var tmpx = xEnd + len * _ex;
        var tmpy = yEnd + len * _ey;

        // (x1, y1) - top arrow point
        var x1 = tmpx + _vx * w/2;
        var y1 = tmpy + _vy * w/2;

        // (x3, y3) - bottom arrow point
        var x3 = tmpx - _vx * w/2;
        var y3 = tmpy - _vy * w/2;

        let controlPointShift = 0.5 * len;

        var x4 = xEnd + controlPointShift * _ex;
        var y4 = yEnd + controlPointShift * _ey;

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        let cpxLocal = trans.TransformPointX(x4, y4);
        let cpyLocal = trans.TransformPointY(x4, y4);
        let endXLocal = trans.TransformPointX(xEnd, yEnd);
        let endYLocal = trans.TransformPointY(xEnd, yEnd);
        drawer._c2(cpxLocal, cpyLocal, endXLocal, endYLocal);

        // drawer._m(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
        drawer._c2(cpxLocal, cpyLocal, trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));

        if (isArrowClosed) {
            // drawer._m(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
            drawer._c2(cpxLocal, cpyLocal, trans.TransformPointX(x1, y1),  trans.TransformPointY(x1, y1));
        }

        // drawer._z();

        stokeOrFillPath(drawer, isArrowFilled);

        drawer._e();
    } else if (type === AscFormat.LineEndType.vsdxDimensionLine) {
        const drawLineAngle = Math.atan2(yEnd - yPrev, xEnd - xPrev) + (45 * Math.PI / 180);

        // Вычисляем координаты конца перпендикулярной линии
        const perpendicularLength = w * Math.sin(Math.PI / 4); // don't know why it's not just =w here
        const x1 = xEnd + perpendicularLength * Math.cos(drawLineAngle); // top right point for visio
        const y1 = yEnd + perpendicularLength * Math.sin(drawLineAngle); // top right point for visio
        const x2 = xEnd - perpendicularLength * Math.cos(drawLineAngle);
        const y2 = yEnd - perpendicularLength * Math.sin(drawLineAngle);

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));
        drawer.ds();
    } else if (type === AscFormat.LineEndType.vsdxFilledDot ||
        type === AscFormat.LineEndType.vsdxClosedDot) {
        let isFilled = type === AscFormat.LineEndType.vsdxFilledDot;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, 0, w, len, isFilled);
    } else if (type === AscFormat.LineEndType.vsdxFilledSquare ||
        type === AscFormat.LineEndType.vsdxClosedSquare) {
        let isArrowFilled = type === AscFormat.LineEndType.vsdxFilledSquare;
        let _ex = xPrev - xEnd;
        let _ey = yPrev - yEnd;
        let _elen = Math.sqrt(_ex*_ex + _ey*_ey);
        _ex /= _elen; // cos a
        _ey /= _elen; // sin a

        let isRotated = false;
        _ex = isRotated? _ex : 1; // cos a
        _ey = isRotated? _ey : 0; // sin a

        // a now is 90 - a + invert cos
        let _vx = _ey;
        let _vy = -_ex;

        // (xEnd, yEnd) - right arrow point

        let paramsScale = 0.4;

        // left arrow point?
        let tmpx = xEnd + len * paramsScale * _ex;
        let tmpy = yEnd + len * paramsScale * _ey;

        // (x1, y1) - left top arrow point
        let x1 = tmpx + _vx * w * paramsScale;
        let y1 = tmpy + _vy * w * paramsScale;

        // (x2, y2) - left bottom arrow point
        let x2 = tmpx - _vx * w * paramsScale;
        let y2 = tmpy - _vy * w * paramsScale;

        // right arrow point?
        let tmpxRight = xEnd - len * paramsScale * _ex;
        let tmpyRight = yEnd - len * paramsScale * _ey;

        // (x3, y3) - left top arrow point
        let x3 = tmpxRight + _vx * w * paramsScale;
        let y3 = tmpyRight + _vy * w * paramsScale;

        // (x4, y4) - left bottom arrow point
        let x4 = tmpxRight - _vx * w * paramsScale;
        let y4 = tmpyRight - _vy * w * paramsScale;

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
        drawer._l(trans.TransformPointX(x4, y4), trans.TransformPointY(x4, y4));
        drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));

        drawer._z();

        stokeOrFillPath(drawer, isArrowFilled);

        drawer._e();
    } else if (type === AscFormat.LineEndType.vsdxDiamond) {
        drawRhomb(drawer, xPrev, yPrev, xEnd, yEnd, 0);
    } else if (type === AscFormat.LineEndType.vsdxBackslash) {
        let _ex = xPrev - xEnd;
        let _ey = yPrev - yEnd;

        const arrowPartLen =  Math.sqrt(_ex * _ex + _ey * _ey);
        const arrowCos = _ex / arrowPartLen;
        const arrowSin = _ey / arrowPartLen;
        const xShift = len * arrowCos;
        const yShift = len * arrowSin;
        const drawLineAngle = Math.atan2(yEnd - yPrev, xEnd - xPrev) - (45 * Math.PI / 180);

        // Вычисляем координаты конца перпендикулярной линии
        const perpendicularLength = w * Math.sin(Math.PI / 4); // don't know why it's not just =w here
        const x1 = xEnd + perpendicularLength * Math.cos(drawLineAngle) + xShift; // top right point for visio
        const y1 = yEnd + perpendicularLength * Math.sin(drawLineAngle) + yShift; // top right point for visio
        const x2 = xEnd - perpendicularLength * Math.cos(drawLineAngle) + xShift;
        const y2 = yEnd - perpendicularLength * Math.sin(drawLineAngle) + yShift;

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));
        drawer.ds();

        drawer._e();
    } else if (type === AscFormat.LineEndType.vsdxOpenOneDash) {
        let shift = 0.75 * len;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
    } else if (type === AscFormat.LineEndType.vsdxOpenTwoDash) {
        let shift = 0.75 * len;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift * 1.5);
    } else if (type === AscFormat.LineEndType.vsdxOpenThreeDash) {
        let shift = 0.75 * len;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift * 1.5);
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift * 2);
    } else if (type === AscFormat.LineEndType.vsdxFork) {
        drawFork(drawer, xPrev, yPrev, xEnd, yEnd);
    } else if (type === AscFormat.LineEndType.vsdxDashFork) {
        drawFork(drawer, xPrev, yPrev, xEnd, yEnd);

        let shift = 1.5 * len * 0.75;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
    } else if (type === AscFormat.LineEndType.vsdxClosedFork) {
        let isFilled = false;
        drawFork(drawer, xPrev, yPrev, xEnd, yEnd);

        let shift = 0.75 * len * 1.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, isFilled);
    } else if (type === AscFormat.LineEndType.vsdxClosedPlus) {
        let shift = 0.75 * len * 1.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, false);

        shift = 0.75 * len * 0.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
    } else if (type === AscFormat.LineEndType.vsdxClosedOneDash) {
        let shift = 0.75 * len * 1.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);

        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, false);
    } else if (type === AscFormat.LineEndType.vsdxClosedTwoDash) {
        let shift = 0.75 * len * 2;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);

        shift = 0.75 * len * 1.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);

        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, false);
    } else if (type === AscFormat.LineEndType.vsdxClosedThreeDash) {
        let shift = 0.75 * len * 2.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 2;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 1.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, false);
    } else if (type === AscFormat.LineEndType.vsdxClosedDiamond) {
        let shift = 0.75 * len * 1.25;
        drawRhomb(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, false);
    } else if (type === AscFormat.LineEndType.vsdxFilledOneDash) {
        let shift = 0.75 * len * 1.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, true);
    } else if (type === AscFormat.LineEndType.vsdxFilledTwoDash) {
        let shift = 0.75 * len * 2;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 1.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, true);
    } else if (type === AscFormat.LineEndType.vsdxFilledThreeDash) {
        let shift = 0.75 * len * 2.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 2;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 1.5;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, true);
    } else if (type === AscFormat.LineEndType.vsdxFilledDiamond) {
        let shift = 0.75 * len * 1.25;
        drawRhomb(drawer, xPrev, yPrev, xEnd, yEnd, shift);
        shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, true);
    } else if (type === AscFormat.LineEndType.vsdxFilledDoubleArrow) {
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, true, false, w, len);
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, len, true, false, w, len);
    } else if (type === AscFormat.LineEndType.vsdxClosedDoubleArrow) {
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, false, false, w, len);
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, len, false, false, w, len);
    } else if (type === AscFormat.LineEndType.vsdxClosedNoDash) {
        let shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, false);
    } else if (type === AscFormat.LineEndType.vsdxFilledNoDash) {
        let shift = 0.75 * len * 0.5;
        drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, true);
    } else if (type === AscFormat.LineEndType.vsdxOpenDoubleArrow) {
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, false, true, w, len);
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, len, false, true, w, len);
    } else if (type === AscFormat.LineEndType.vsdxOpenArrowSingleDash) {
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, false, true, w, len);
        let shift = 0.75 * len * 2;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
    } else if (type === AscFormat.LineEndType.vsdxOpenDoubleArrowSingleDash) {
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, 0, false, true, w, len);
        drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, len, false, true, w, len);
        let shift = 0.75 * len * 2 + len;
        drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift);
    }

    /**
     * @param {CShapeDrawer} drawer
     * @param {boolean} isFilled - if not filled it is stroke bcs if both repeat all commands
     */
    function stokeOrFillPath(drawer, isFilled) {
        if (isFilled) {
            if (Asc.editor.isPdfEditor()) {
                let oRGBColor;
                if (drawer.Shape.GetRGBColor) {
                    oRGBColor = drawer.Shape.GetRGBColor(drawer.Shape.GetFillColor());
                }
                else if (drawer.Shape.group) {
                    oRGBColor = drawer.Shape.group.GetRGBColor(drawer.Shape.group.GetFillColor());
                }

                if (oRGBColor) {
                    drawer.Graphics.m_oPen.Color.R = oRGBColor.r;
                    drawer.Graphics.m_oPen.Color.G = oRGBColor.g;
                    drawer.Graphics.m_oPen.Color.B = oRGBColor.b;
                }
            }
            drawer.drawStrokeFillStyle();
        }
        let isStroke = !isFilled;
        if (isStroke) {
            drawer.ds();
        }
    }

    /**
     * Draws fork BEFORE line end
     * @param drawer
     * @param xPrev
     * @param yPrev
     * @param xEnd
     * @param yEnd
     */
    function drawFork(drawer, xPrev, yPrev, xEnd, yEnd) {
        var arrowAngle = Math.atan2(yEnd - yPrev, xEnd - xPrev);
        var perpendicularAngle = arrowAngle - (90 * Math.PI / 180);


        // Вычисляем координаты конца перпендикулярной линии
        var perpendicularLength = w / 2;
        let horLength = w * 0.75;
        var x1 = xEnd + perpendicularLength * Math.cos(perpendicularAngle); // top right point
        var y1 = yEnd + perpendicularLength * Math.sin(perpendicularAngle); // top right point
        var x2 = xEnd - perpendicularLength * Math.cos(perpendicularAngle); // bottom right point
        var y2 = yEnd - perpendicularLength * Math.sin(perpendicularAngle); // bottom right point
        let x3 = xEnd - horLength * Math.cos(arrowAngle); // left point
        let y3 = yEnd - horLength * Math.sin(arrowAngle); // left point

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
        drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));
        drawer.ds();

        drawer._e();
    }

    /**
     * Draws vertical line ON line end
     * @param drawer
     * @param xPrev
     * @param yPrev
     * @param xEnd
     * @param yEnd
     * @param shift
     */
    function drawVerticalLine(drawer, xPrev, yPrev, xEnd, yEnd, shift) {
        let arrowAngle = Math.atan2(yEnd - yPrev, xEnd - xPrev);
        var drawLineAngle = arrowAngle  - (90 * Math.PI / 180);

        // Вычисляем координаты конца перпендикулярной линии
        var perpendicularLength = w / 2;
        var x1 = xEnd + perpendicularLength * Math.cos(drawLineAngle) - shift * Math.cos(arrowAngle); // top right point for visio
        var y1 = yEnd + perpendicularLength * Math.sin(drawLineAngle) - shift * Math.sin(arrowAngle); // top right point for visio
        var x2 = xEnd - perpendicularLength * Math.cos(drawLineAngle) - shift * Math.cos(arrowAngle);
        var y2 = yEnd - perpendicularLength * Math.sin(drawLineAngle) - shift * Math.sin(arrowAngle);

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));
        drawer.ds();

        drawer._e();
    }

    /**
     * Draws center ON line end
     * @param drawer
     * @param xPrev
     * @param yPrev
     * @param xEnd
     * @param yEnd
     * @param shift - horizontal shift distance
     * @param w
     * @param len
     * @param isArrowFilled
     */
    function drawCircle(drawer, xPrev, yPrev, xEnd, yEnd, shift, w, len, isArrowFilled) {
        len *= 0.75;
        w *= 0.75;

        var _ex = xPrev - xEnd;
        var _ey = yPrev - yEnd;
        var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
        _ex /= _elen; // cos a
        _ey /= _elen; // sin a

        var _vx = _ey;
        var _vy = -_ex;

        var tmpx = xEnd + (len / 2 + shift) * _ex; // left circle point
        var tmpy = yEnd + (len / 2 + shift) * _ey; // left circle point

        var tmpx2 = xEnd - (len / 2 - shift) * _ex; // right circle point
        var tmpy2 = yEnd - (len / 2 - shift) * _ey; // right circle point

        // left top control point
        var cx1 = tmpx + _vx * 3*w/4;
        var cy1 = tmpy + _vy * 3*w/4;
        // right top control point
        var cx2 = tmpx2 + _vx * 3*w/4;
        var cy2 = tmpy2 + _vy * 3*w/4;

        // left bottom control point
        var cx3 = tmpx - _vx * 3*w/4;
        var cy3 = tmpy - _vy * 3*w/4;
        // right bottom control point
        var cx4 = tmpx2 - _vx * 3*w/4;
        var cy4 = tmpy2 - _vy * 3*w/4;

        drawer._s();
        drawer._m(trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));
        drawer._c(trans.TransformPointX(cx1, cy1), trans.TransformPointY(cx1, cy1),
            trans.TransformPointX(cx2, cy2), trans.TransformPointY(cx2, cy2),
            trans.TransformPointX(tmpx2, tmpy2), trans.TransformPointY(tmpx2, tmpy2));

        drawer._c(trans.TransformPointX(cx4, cy4), trans.TransformPointY(cx4, cy4),
            trans.TransformPointX(cx3, cy3), trans.TransformPointY(cx3, cy3),
            trans.TransformPointX(tmpx, tmpy), trans.TransformPointY(tmpx, tmpy));

        stokeOrFillPath(drawer, isArrowFilled);
        drawer._e();
    }

    /**
     * Draws BEFORE line end
     * @param drawer
     * @param xPrev
     * @param yPrev
     * @param xEnd
     * @param yEnd
     * @param shift
     */
    function drawRhomb(drawer, xPrev, yPrev, xEnd, yEnd, shift) {
        let isArrowFilled = false;
        let _ex = xPrev - xEnd;
        let _ey = yPrev - yEnd;
        const _elen = Math.sqrt(_ex * _ex + _ey * _ey);
        _ex /= _elen; // cos a
        _ey /= _elen; // sin a

        let isRotated = true;
        _ex = isRotated? _ex : 1; // cos a
        _ey = isRotated? _ey : 0; // sin a

        // a now is 90 - a + invert cos
        const _vx = _ey;
        const _vy = -_ex;

        // (xEnd, yEnd) - right arrow point
        xEnd = xEnd + shift * _ex;
        yEnd = yEnd + shift * _ey;

        let heightScale = 0.5;
        let widthScale = 1;

        // middle arrow point
        const tmpx = xEnd + len * widthScale * _ex;
        const tmpy = yEnd + len * widthScale * _ey;

        // (x1, y1) - middle top arrow point
        const x1 = tmpx + _vx * w * heightScale;
        const y1 = tmpy + _vy * w * heightScale;

        // (x2, y2) - middle bottom arrow point
        const x2 = tmpx - _vx * w * heightScale;
        const y2 = tmpy - _vy * w * heightScale;

        // most left point
        const x3 = tmpx + len * widthScale * _ex;
        const y3 = tmpy + len * widthScale * _ey;

        drawer._s();
        drawer._m(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
        drawer._l(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
        drawer._l(trans.TransformPointX(x2, y2), trans.TransformPointY(x2, y2));

        drawer._z();

        stokeOrFillPath(drawer, isArrowFilled);

        drawer._e();
    }

    /**
     * Draws arrow BEFORE line end
     * @param drawer
     * @param xPrev
     * @param yPrev
     * @param xEnd
     * @param yEnd
     * @param shift
     * @param isFilled
     * @param isOpen
     * @param w
     * @param len
     */
    function drawArrow(drawer, xPrev, yPrev, xEnd, yEnd, shift, isFilled, isOpen, w, len) {
        if (Asc.editor.isPdfEditor() == true) {
            drawer.CheckDash();
        }

        var _ex = xPrev - xEnd;
        var _ey = yPrev - yEnd;
        var _elen = Math.sqrt(_ex*_ex + _ey*_ey);
        _ex /= _elen;
        _ey /= _elen;

        var _vx = _ey;
        var _vy = -_ex;

        xEnd += shift * _ex;
        yEnd += shift * _ey;

        // (xEnd, yEnd) - right arrow point
        var tmpx = xEnd + len * _ex;
        var tmpy = yEnd + len * _ey;

        // (x1, y1) - top arrow point
        var x1 = tmpx + _vx * w/2;
        var y1 = tmpy + _vy * w/2;

        // (x3, y3) - bottom arrow point
        var x3 = tmpx - _vx * w/2;
        var y3 = tmpy - _vy * w/2;

        drawer._s();
        drawer._m(trans.TransformPointX(x1, y1), trans.TransformPointY(x1, y1));
        drawer._l(trans.TransformPointX(xEnd, yEnd), trans.TransformPointY(xEnd, yEnd));
        drawer._l(trans.TransformPointX(x3, y3), trans.TransformPointY(x3, y3));
        if (!isOpen) {
            drawer._z();
        }
        stokeOrFillPath(drawer, isFilled);
        drawer._e();
    }
}


function CShapeDrawer()
{
    this.Shape = null;
    this.Graphics = null;
    this.UniFill = null;
    this.Ln = null;
    this.Transform = null;

    this.bIsTexture = false;
    this.bIsNoFillAttack = false;
    this.bIsNoStrokeAttack = false;
    this.bDrawSmartAttack = false;
    this.FillUniColor = null;
    this.StrokeUniColor = null;
    this.StrokeWidth = 0;

    this.min_x = 0xFFFF;
    this.min_y = 0xFFFF;
    this.max_x = -0xFFFF;
    this.max_y = -0xFFFF;

    this.OldLineJoin = null;
    this.IsArrowsDrawing = false;
    this.IsCurrentPathCanArrows = true;

    this.bIsCheckBounds = false;

    this.IsRectShape = false;
}

CShapeDrawer.prototype =
{
    Clear : function()
    {
        //this.Shape = null;
        //this.Graphics = null;
        this.UniFill = null;
        this.Ln = null;
        this.Transform = null;

        this.bIsTexture = false;
        this.bIsNoFillAttack = false;
        this.bIsNoStrokeAttack = false;
        this.FillUniColor = null;
        this.StrokeUniColor = null;
        this.StrokeWidth = 0;

        this.min_x = 0xFFFF;
        this.min_y = 0xFFFF;
        this.max_x = -0xFFFF;
        this.max_y = -0xFFFF;

        this.OldLineJoin = null;
        this.IsArrowsDrawing = false;
        this.IsCurrentPathCanArrows = true;

        this.bIsCheckBounds = false;

        this.IsRectShape = false;
    },

    isPdf : function()
    {
        return this.Graphics.isPdf() || this.Graphics.isNativeDrawer();
    },

    CheckPoint : function(_x,_y)
    {
        // TODO: !!!
        var x = _x;
        var y = _y;

        if (false && this.Graphics.MaxEpsLine !== undefined)
        {
            x = this.Graphics.Graphics.m_oFullTransform.TransformPointX(_x,_y);
            y = this.Graphics.Graphics.m_oFullTransform.TransformPointY(_x,_y);
        }

        if (x < this.min_x)
            this.min_x = x;
        if (y < this.min_y)
            this.min_y = y;
        if (x > this.max_x)
            this.max_x = x;
        if (y > this.max_y)
            this.max_y = y;
    },

    CheckDash : function()
    {
        if (Asc.editor.isPdfEditor()) {
            let aDash;
            if (this.Shape.GetDash)
                aDash = this.Shape.GetDash();
            else if (this.Shape.group && this.Shape.group.GetDash)
                aDash = this.Shape.group.GetDash();
            else if (this.Shape.IsDrawing && this.Shape.IsDrawing()) {
                if (AscCommon.DashPatternPresets[this.Ln.prstDash]) {
                    aDash = AscCommon.DashPatternPresets[this.Ln.prstDash].slice();
                    for (var indexD = 0; indexD < aDash.length; indexD++)
                        aDash[indexD] *= 2 * this.StrokeWidth;
                }
            }

            if (aDash) {
                this.Graphics.p_dash(aDash.map(function(measure) {
                    return measure / 2;
                }));
            }
        }
        else if (this.Ln.prstDash != null && AscCommon.DashPatternPresets[this.Ln.prstDash])
        {
            var _arr = AscCommon.DashPatternPresets[this.Ln.prstDash].slice();
            for (var indexD = 0; indexD < _arr.length; indexD++)
                _arr[indexD] *= this.StrokeWidth;
            this.Graphics.p_dash(_arr);
        }
        else if (this.isPdf())
        {
            this.Graphics.p_dash(null);
        }
    },

    fromShape2 : function(shape, graphics, geom)
    {
        this.fromShape(shape, graphics);

        if (!geom && !graphics.bDrawRectWithLines)
        {
            this.IsRectShape = true;
        }
        else
        {
            if (geom.preset == "rect" && !graphics.bDrawRectWithLines)
                this.IsRectShape = true;
        }
    },

    fromShape : function(shape, graphics)
    {
        this.IsRectShape = false;
        this.Shape = shape;
        this.Graphics = graphics;
        this.UniFill = shape.brush;
        this.Ln = shape.pen;
        this.Transform = shape.TransformMatrix;

        this.min_x = 0xFFFF;
        this.min_y = 0xFFFF;
        this.max_x = -0xFFFF;
        this.max_y = -0xFFFF;

        var bIsCheckBounds = false;

        if (this.UniFill == null || this.UniFill.fill == null)
            this.bIsNoFillAttack = true;
        else
        {
            var _fill = this.UniFill.fill;
            switch (_fill.type)
            {
                case c_oAscFill.FILL_TYPE_BLIP:
                {
                    this.bIsTexture = true;
                    break;
                }
                case c_oAscFill.FILL_TYPE_SOLID:
                {
                    if(_fill.color)
                    {
                        this.FillUniColor = _fill.color.RGBA;
                    }
                    else
                    {
                        this.FillUniColor = new AscFormat.CUniColor().RGBA;
                    }
                    break;
                }
                case c_oAscFill.FILL_TYPE_GRAD:
                {
                    var _c = _fill.colors;
                    if (_c.length == 0)
                        this.FillUniColor = new AscFormat.CUniColor().RGBA;
                    else
                    {
                        if(_fill.colors[0].color)
                        {
                            this.FillUniColor = _fill.colors[0].color.RGBA;
                        }
                        else
                        {
                            this.FillUniColor = new AscFormat.CUniColor().RGBA;
                        }
                    }

                    bIsCheckBounds = true;

                    break;
                }
                case c_oAscFill.FILL_TYPE_PATT:
                {
                    bIsCheckBounds = true;
                    break;
                }
                case c_oAscFill.FILL_TYPE_NOFILL:
                {
                    this.bIsNoFillAttack = true;
                    break;
                }
                default:
                {
                    this.bIsNoFillAttack = true;
                    break;
                }
            }
        }

        if (this.Ln == null || this.Ln.Fill == null || this.Ln.Fill.fill == null)
        {
            this.bIsNoStrokeAttack = true;
            if (graphics.isTrack())
                graphics.Graphics.ArrayPoints = null;
            else
                graphics.ArrayPoints = null;
        }
        else
        {
            var _fill = this.Ln.Fill.fill;
            switch (_fill.type)
            {
                case c_oAscFill.FILL_TYPE_BLIP:
                {
                    this.StrokeUniColor = new AscFormat.CUniColor().RGBA;
                    break;
                }
                case c_oAscFill.FILL_TYPE_SOLID:
                {
                    if(_fill.color)
                    {
                        this.StrokeUniColor = _fill.color.RGBA;
                    }
                    else
                    {
                        this.StrokeUniColor = new AscFormat.CUniColor().RGBA;
                    }
                    break;
                }
                case c_oAscFill.FILL_TYPE_GRAD:
                {
                    var _c = _fill.colors;
                    if (_c == 0)
                        this.StrokeUniColor = new AscFormat.CUniColor().RGBA;
                    else
                    {
                        if(_fill.colors[0].color)
                        {
                            this.StrokeUniColor = _fill.colors[0].color.RGBA;
                        }
                        else
                        {
                            this.StrokeUniColor = new AscFormat.CUniColor().RGBA;
                        }
                    }

                    break;
                }
                case c_oAscFill.FILL_TYPE_PATT:
                {
                    if(_fill.fgClr)
                    {
                        this.StrokeUniColor = _fill.fgClr.RGBA;
                    }
                    else
                    {
                        this.StrokeUniColor = new AscFormat.CUniColor().RGBA;
                    }
                    break;
                }
                case c_oAscFill.FILL_TYPE_NOFILL:
                {
                    this.bIsNoStrokeAttack = true;
                    break;
                }
                default:
                {
                    this.bIsNoStrokeAttack = true;
                    break;
                }
            }

            this.StrokeWidth = (this.Ln.w == null) ? 12700 : parseInt(this.Ln.w);
            this.StrokeWidth /= 36000.0;

            this.p_width(1000 * this.StrokeWidth);
            
            this.CheckDash();

            if (graphics.isBoundsChecker() && !this.bIsNoStrokeAttack)
                graphics.LineWidth = this.StrokeWidth;

            var isUseArrayPoints = false;
            if ((this.Ln.headEnd != null && this.Ln.headEnd.type != null) || (this.Ln.tailEnd != null && this.Ln.tailEnd.type != null))
                isUseArrayPoints = true;

            if (graphics.isTrack() && graphics.Graphics)
                graphics.Graphics.ArrayPoints = isUseArrayPoints ? [] : null;
            else
                graphics.ArrayPoints = isUseArrayPoints ? [] : null;

            if (this.Graphics.m_oContext != null && this.Ln.Join != null && this.Ln.Join.type != null)
                this.OldLineJoin = this.Graphics.m_oContext.lineJoin;
        }

        if (this.bIsTexture || bIsCheckBounds)
        {
            // сначала нужно определить границы
            this.bIsCheckBounds = true;
            this.check_bounds();
            this.bIsCheckBounds = false;
        }
    },

    draw : function(geom)
    {
        if (this.bIsNoStrokeAttack && this.bIsNoFillAttack)
            return;

        var bIsPatt = false;
        if (this.UniFill != null && this.UniFill.fill != null &&
            ((this.UniFill.fill.type == c_oAscFill.FILL_TYPE_PATT) || (this.UniFill.fill.type == c_oAscFill.FILL_TYPE_GRAD)))
        {
            bIsPatt = true;
        }

        if (this.isPdf() && (this.bIsTexture || bIsPatt))
        {
            this.Graphics.put_TextureBoundsEnabled(true);
            this.Graphics.put_TextureBounds(this.min_x, this.min_y, this.max_x - this.min_x, this.max_y - this.min_y);
        }

        if(geom)
        {
            geom.draw(this);
        }
        else
        {
            this._s();
            this._m(0, 0);
            this._l(this.Shape.extX, 0);
            this._l(this.Shape.extX, this.Shape.extY);
            this._l(0, this.Shape.extY);
            this._z();
            this.drawFillStroke(true, "norm", true && !this.bIsNoStrokeAttack);
            this._e();
        }
        this.Graphics.ArrayPoints = null;

        if (this.isPdf() && (this.bIsTexture || bIsPatt))
        {
            this.Graphics.put_TextureBoundsEnabled(false);
        }

        if (this.Graphics.isBoundsChecker() && this.Graphics.AutoCheckLineWidth)
        {
            this.Graphics.CorrectBounds2();
        }

        this.Graphics.p_dash(null);
    },

    p_width : function(w)
    {
        this.Graphics.p_width(w);
    },

    _m : function(x, y)
    {
        if (this.bIsCheckBounds)
        {
            this.CheckPoint(x, y);
            return;
        }

        this.Graphics._m(x, y);
    },

    _l : function(x, y)
    {
        if (this.bIsCheckBounds)
        {
            this.CheckPoint(x, y);
            return;
        }

        this.Graphics._l(x, y);
    },

    _c : function(x1, y1, x2, y2, x3, y3)
    {
        if (this.bIsCheckBounds)
        {
            this.CheckPoint(x1, y1);
            this.CheckPoint(x2, y2);
            this.CheckPoint(x3, y3);
            return;
        }

        this.Graphics._c(x1, y1, x2, y2, x3, y3);
    },

    /**
     * @param x1 - cpx
     * @param y1 - cpy
     * @param x2 - end x
     * @param y2 - end y
     */
    _c2 : function(x1, y1, x2, y2)
    {
        if (this.bIsCheckBounds)
        {
            this.CheckPoint(x1, y1);
            this.CheckPoint(x2, y2);
            return;
        }

        this.Graphics._c2(x1, y1, x2, y2);
    },

    /**
     * @param {{x: Number, y: Number, z? :Number}} startPoint
     * @param {{x: Number, y: Number, z? :Number}[]} controlPoints
     * @param {{x: Number, y: Number, z? :Number}} endPoint
     */
    _cN : function(startPoint, controlPoints, endPoint)
    {
        if (this.bIsCheckBounds)
        {
            this.CheckPoint(startPoint.x, startPoint.y);
            controlPoints.forEach(function (controlPoint) {
                this.CheckPoint(controlPoint.x, controlPoint.y);
            })
            this.CheckPoint(endPoint.x, endPoint.y);
            return;
        }

        this.Graphics._cN(startPoint, controlPoints, endPoint);
    },

    _z : function()
    {
        this.IsCurrentPathCanArrows = false;
        if (this.bIsCheckBounds)
            return;

        this.Graphics._z();
    },

    _s : function()
    {
        this.IsCurrentPathCanArrows = true;
        this.Graphics._s();
        if (this.Graphics.ArrayPoints != null)
            this.Graphics.ArrayPoints = [];
    },
    _e : function()
    {
        this.IsCurrentPathCanArrows = true;
        this.Graphics._e();
        if (this.Graphics.ArrayPoints != null)
            this.Graphics.ArrayPoints = [];
    },

    drawTransitionTextures : function(oCanvas1, dAlpha1, oCanvas2, dAlpha2)
    {
        const dOldGlobalAlpha = this.Graphics.m_oContext.globalAlpha;
        const dX = this.min_x;
        const dY = this.min_y;
        const dW = this.max_x - this.min_x;
        const dH = this.max_y - this.min_y;
        this.Graphics.m_oContext.globalAlpha = dAlpha1;
        this.Graphics.drawImage(null, dX, dY, dW, dH, undefined, null, oCanvas1);
        this.Graphics.m_oContext.globalAlpha = dAlpha2;
        this.Graphics.drawImage(null, dX, dY, dW, dH, undefined, null, oCanvas2);
        this.Graphics.m_oContext.globalAlpha = dOldGlobalAlpha;
    },


    df : function(mode)
    {
        if (mode == "none" || this.bIsNoFillAttack)
            return;

        if (this.Graphics.isTrack())
            this.Graphics.m_oOverlay.ClearAll = true;

        if (this.Graphics.isBoundsChecker() === true)
            return;

        var editorInfo = this.getEditorInfo();

        var bIsIntegerGridTRUE = false;
        if (this.bIsTexture)
        {
            if (this.Graphics.m_bIntegerGrid === true)
            {
                this.Graphics.SetIntegerGrid(false);
                bIsIntegerGridTRUE = true;
            }

            if (this.isPdf())
            {
                if (null == this.UniFill.fill.tile || this.Graphics.m_oContext === undefined)
                {
                    this.Graphics.put_brushTexture(getFullImageSrc2(this.UniFill.fill.RasterImageId), 0);
                }
                else
                {
                    this.Graphics.put_brushTexture(getFullImageSrc2(this.UniFill.fill.RasterImageId), 1);
                }

                if (bIsIntegerGridTRUE)
                {
                    this.Graphics.SetIntegerGrid(true);
                }
                return;
            }

			var bIsUnusePattern = false;
			if (AscCommon.AscBrowser.isIE)
			{
				// ie падает иначе !!!
				if (this.UniFill.fill.RasterImageId)
				{
					if (this.UniFill.fill.RasterImageId.lastIndexOf(".svg") == this.UniFill.fill.RasterImageId.length - 4)
						bIsUnusePattern = true;
				}
			}

            if (bIsUnusePattern || null == this.UniFill.fill.tile || this.Graphics.m_oContext === undefined)
            {
                if (this.IsRectShape)
                {
                    if(this.UniFill.IsTransitionTextures)
                    {
                        this.drawTransitionTextures(this.UniFill.canvas1, this.UniFill.alpha1, this.UniFill.canvas2, this.UniFill.alpha2);
                    }
                    else if ((null == this.UniFill.transparent) || (this.UniFill.transparent == 255))
                    {
                        this.Graphics.drawImage(getFullImageSrc2(this.UniFill.fill.RasterImageId), this.min_x, this.min_y, (this.max_x - this.min_x), (this.max_y - this.min_y), undefined, this.UniFill.fill.srcRect, this.UniFill.fill.canvas);
                    }
                    else
                    {
                        var _old_global_alpha = this.Graphics.m_oContext.globalAlpha;
                        this.Graphics.m_oContext.globalAlpha = this.UniFill.transparent / 255;
                        this.Graphics.drawImage(getFullImageSrc2(this.UniFill.fill.RasterImageId), this.min_x, this.min_y, (this.max_x - this.min_x), (this.max_y - this.min_y), undefined, this.UniFill.fill.srcRect, this.UniFill.fill.canvas);
                        this.Graphics.m_oContext.globalAlpha = _old_global_alpha;
                    }
                }
                else
                {
                    this.Graphics.save();
                    this.Graphics.clip();

                    if(this.UniFill.IsTransitionTextures)
                    {
                        this.drawTransitionTextures(this.UniFill.canvas1, this.UniFill.alpha1, this.UniFill.canvas2, this.UniFill.alpha2);
                    }
                    else if (!this.Graphics.isSupportTextDraw() || this.Graphics.isTrack() || (null == this.UniFill.transparent) || (this.UniFill.transparent == 255))
                    {
                        this.Graphics.drawImage(getFullImageSrc2(this.UniFill.fill.RasterImageId), this.min_x, this.min_y, (this.max_x - this.min_x), (this.max_y - this.min_y), undefined, this.UniFill.fill.srcRect, this.UniFill.fill.canvas);
                    }
                    else
                    {
                        var _old_global_alpha = this.Graphics.m_oContext.globalAlpha;
                        this.Graphics.m_oContext.globalAlpha = this.UniFill.transparent / 255;
                        this.Graphics.drawImage(getFullImageSrc2(this.UniFill.fill.RasterImageId), this.min_x, this.min_y, (this.max_x - this.min_x), (this.max_y - this.min_y), undefined, this.UniFill.fill.srcRect, this.UniFill.fill.canvas);
                        this.Graphics.m_oContext.globalAlpha = _old_global_alpha;
                    }

                    this.Graphics.restore();
                }
            }
            else
            {
                var _img = editorInfo.editor.ImageLoader.map_image_index[getFullImageSrc2(this.UniFill.fill.RasterImageId)];
                var _img_native = this.UniFill.fill.canvas;
                if ((!_img_native) && (_img == undefined || _img.Image == null || _img.Status == AscFonts.ImageLoadStatus.Loading))
                {
                    this.Graphics.save();
                    this.Graphics.clip();

                    if (!this.Graphics.isSupportTextDraw() || this.Graphics.isTrack() || (null == this.UniFill.transparent) || (this.UniFill.transparent == 255))
                    {
                        this.Graphics.drawImage(getFullImageSrc2(this.UniFill.fill.RasterImageId), this.min_x, this.min_y, (this.max_x - this.min_x), (this.max_y - this.min_y));
                    }
                    else
                    {
                        var _old_global_alpha = this.Graphics.m_oContext.globalAlpha;
                        this.Graphics.m_oContext.globalAlpha = this.UniFill.transparent / 255;
                        this.Graphics.drawImage(getFullImageSrc2(this.UniFill.fill.RasterImageId), this.min_x, this.min_y, (this.max_x - this.min_x), (this.max_y - this.min_y));
                        this.Graphics.m_oContext.globalAlpha = _old_global_alpha;
                    }

                    this.Graphics.restore();
                }
                else
                {
                    var _is_ctx = false;
                    if (!this.Graphics.isSupportTextDraw() || undefined === this.Graphics.m_oContext || (null == this.UniFill.transparent) || (this.UniFill.transparent == 255))
                    {
                        _is_ctx = false;
                    }
                    else
                    {
                        _is_ctx = true;
                    }

                    var _gr = this.Graphics.isTrack() ? this.Graphics.Graphics : this.Graphics;
                    var _ctx = _gr.m_oContext;

                    var patt = !_img_native ? _ctx.createPattern(_img.Image, "repeat") : _ctx.createPattern(_img_native, "repeat");

                    _ctx.save();

                    var __graphics = (this.Graphics.MaxEpsLine === undefined) ? this.Graphics : this.Graphics.Graphics;
                    var bIsThumbnail = (__graphics.IsThumbnail === true) ? true : false;

                    var koefX = editorInfo.scale;
                    var koefY = editorInfo.scale;

                    if (bIsThumbnail)
                    {
                        koefX = __graphics.m_dDpiX / AscCommon.g_dDpiX;
                        koefY = __graphics.m_dDpiY / AscCommon.g_dDpiX;

                        koefX /= AscCommon.AscBrowser.retinaPixelRatio;
                        koefY /= AscCommon.AscBrowser.retinaPixelRatio;
                    }

                    // TODO: !!!
                    _ctx.translate(this.min_x, this.min_y);

                    _ctx.scale(koefX * __graphics.TextureFillTransformScaleX, koefY * __graphics.TextureFillTransformScaleY);

                    if (_is_ctx === true)
                    {
                        var _old_global_alpha = _ctx.globalAlpha;
                        _ctx.globalAlpha = this.UniFill.transparent / 255;
                        _ctx.fillStyle = patt;
                        _ctx.fill();
                        _ctx.globalAlpha = _old_global_alpha;
                    }
                    else
                    {
                        _ctx.fillStyle = patt;
                        _ctx.fill();
                    }

                    _ctx.restore();

                    _gr.m_bPenColorInit = false;
                    _gr.m_bBrushColorInit = false;
                }
            }

            if (bIsIntegerGridTRUE)
            {
                this.Graphics.SetIntegerGrid(true);
            }
            return;
        }

        if (this.UniFill != null && this.UniFill.fill != null)
        {
            var _fill = this.UniFill.fill;
            if (_fill.type == c_oAscFill.FILL_TYPE_PATT)
            {
                if (this.Graphics.m_bIntegerGrid === true)
                {
                    this.Graphics.SetIntegerGrid(false);
                    bIsIntegerGridTRUE = true;
                }

                var _is_ctx = false;
                if (!this.Graphics.isSupportTextDraw() || undefined === this.Graphics.m_oContext || (null == this.UniFill.transparent) || (this.UniFill.transparent == 255))
                {
                    _is_ctx = false;
                }
                else
                {
                    _is_ctx = true;
                }

                var _gr = this.Graphics.isTrack() ? this.Graphics.Graphics : this.Graphics;
                var _ctx = _gr.m_oContext;

                var _patt_name = AscCommon.global_hatch_names[_fill.ftype];
                if (undefined == _patt_name)
                    _patt_name = "cross";

                var _fc = _fill.fgClr && _fill.fgClr.RGBA || {R: 0, G: 0, B: 0, A: 255};
                var _bc = _fill.bgClr && _fill.bgClr.RGBA || {R: 255, G: 255, B: 255, A: 255};

                var __fa = (null === this.UniFill.transparent) ? _fc.A : 255;
                var __ba = (null === this.UniFill.transparent) ? _bc.A : 255;

                var _test_pattern = AscCommon.GetHatchBrush(_patt_name, _fc.R, _fc.G, _fc.B, __fa, _bc.R, _bc.G, _bc.B, __ba);
                var patt = _ctx.createPattern(_test_pattern.Canvas, "repeat");

                _ctx.save();

                var koefX = editorInfo.scale;
                var koefY = editorInfo.scale;
                if (this.Graphics.IsThumbnail)
                {
                    koefX = 1;
                    koefY = 1;
                }

                // TODO: !!!
                _ctx.translate(this.min_x, this.min_y);

                if (this.Graphics.MaxEpsLine === undefined)
                {
                    _ctx.scale(koefX * this.Graphics.TextureFillTransformScaleX, koefY * this.Graphics.TextureFillTransformScaleY);
                }
                else
                {
                    _ctx.scale(koefX * this.Graphics.Graphics.TextureFillTransformScaleX, koefY * this.Graphics.Graphics.TextureFillTransformScaleY);
                }

                if (_is_ctx === true)
                {
                    var _old_global_alpha = _ctx.globalAlpha;

                    if (null != this.UniFill.transparent)
                        _ctx.globalAlpha = this.UniFill.transparent / 255;

                    _ctx.fillStyle = patt;
                    _ctx.fill();
                    _ctx.globalAlpha = _old_global_alpha;
                }
                else
                {
                    _ctx.fillStyle = patt;
                    _ctx.fill();
                }

                _ctx.restore();

                _gr.m_bPenColorInit = false;
                _gr.m_bBrushColorInit = false;

                if (bIsIntegerGridTRUE)
                {
                    this.Graphics.SetIntegerGrid(true);
                }
                return;
            }
            else if (_fill.type == c_oAscFill.FILL_TYPE_GRAD)
            {
                if (this.Graphics.m_bIntegerGrid === true)
                {
                    this.Graphics.SetIntegerGrid(false);
                    bIsIntegerGridTRUE = true;
                }

                var _is_ctx = false;
                if (!this.Graphics.isSupportTextDraw() || undefined === this.Graphics.m_oContext || (null == this.UniFill.transparent) || (this.UniFill.transparent == 255))
                {
                    _is_ctx = false;
                }
                else
                {
                    _is_ctx = true;
                }

                var _gr = this.Graphics.isTrack() ? this.Graphics.Graphics : this.Graphics;
                var _ctx = _gr.m_oContext;

                var gradObj = null;
                if (_fill.lin)
                {
                    var _angle = _fill.lin.angle;
                    if (_fill.rotateWithShape === false)
                    {
                        var matrix_transform = this.Graphics.isTrack() ? this.Graphics.Graphics.m_oTransform : this.Graphics.m_oTransform;
                        if (matrix_transform)
                        {
                            //_angle -= (60000 * this.Graphics.m_oTransform.GetRotation());
                            _angle = AscCommon.GradientGetAngleNoRotate(_angle, matrix_transform);
                        }
                    }

                    var points = this.getGradientPoints(this.min_x, this.min_y, this.max_x, this.max_y, _angle, _fill.lin.scale);
                    gradObj = _ctx.createLinearGradient(points.x0, points.y0, points.x1, points.y1);
                }
                else if (_fill.path)
                {
                    var _cx = (this.min_x + this.max_x) / 2;
                    var _cy = (this.min_y + this.max_y) / 2;
                    var _r = Math.max(this.max_x - this.min_x, this.max_y - this.min_y) / 2;
                    gradObj = _ctx.createRadialGradient(_cx, _cy, 1, _cx, _cy, _r);
                }
                else
                {
                    //gradObj = _ctx.createLinearGradient(this.min_x, this.min_y, this.max_x, this.min_y);
                    var points = this.getGradientPoints(this.min_x, this.min_y, this.max_x, this.max_y, 0, false);
                    gradObj = _ctx.createLinearGradient(points.x0, points.y0, points.x1, points.y1);
                }

                const nTransparent = this.UniFill.transparent;
                let bUseGlobalAlpha = (null !== nTransparent && undefined !== nTransparent);
                const aColors = _fill.colors;
                const nClrCount = aColors.length;
                if(_fill.path && AscCommon.AscBrowser.isMozilla && bUseGlobalAlpha)
                {
                    bUseGlobalAlpha = false;
                    const dTransparent = nTransparent / 255.0;
                    for (let nClr = 0; nClr < nClrCount; nClr++)
                    {
                        let oClr = aColors[nClr];
                        gradObj.addColorStop(oClr.pos / 100000, oClr.color.getCSSWithTransparent(dTransparent));
                    }
                }
                else
                {
                    for (let nClr = 0; nClr < nClrCount; nClr++)
                    {
                        let oClr = aColors[nClr];
                        gradObj.addColorStop(oClr.pos / 100000, oClr.color.getCSSColor(nTransparent));
                    }
                }

                _ctx.fillStyle = gradObj;

                if (bUseGlobalAlpha)
                {
                    var _old_global_alpha = this.Graphics.m_oContext.globalAlpha;
                    _ctx.globalAlpha = this.UniFill.transparent / 255;
                    _ctx.fill();
                    _ctx.globalAlpha = _old_global_alpha;
                }
                else
                {
                    _ctx.fill();
                }

                _gr.m_bPenColorInit = false;
                _gr.m_bBrushColorInit = false;

                if (bIsIntegerGridTRUE)
                {
                    this.Graphics.SetIntegerGrid(true);
                }
                return;
            }
        }

        var rgba = this.FillUniColor;
        if (mode == "darken")
        {
            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
            var rgb = _color1.darken();
            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
        }
        else if (mode == "darkenLess")
        {
            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
            var rgb = _color1.darkenLess();
            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
        }
        else if (mode == "lighten")
        {
            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
            var rgb = _color1.lighten();
            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
        }
        else if (mode == "lightenLess")
        {
            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
            var rgb = _color1.lightenLess();
            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
        }
        if(rgba)
        {
            if (this.UniFill != null && this.UniFill.transparent != null)
                rgba.A = this.UniFill.transparent;
            this.Graphics.b_color1(rgba.R, rgba.G, rgba.B, rgba.A);
        }
        this.Graphics.df();
    },

    ds : function()
    {
        if (this.bIsNoStrokeAttack)
            return;

        if (this.Graphics.isTrack())
            this.Graphics.m_oOverlay.ClearAll = true;

        if (null != this.OldLineJoin && !this.IsArrowsDrawing)
        {
            switch (this.Ln.Join.type)
            {
                case AscFormat.LineJoinType.Round:
                {
                    this.Graphics.m_oContext.lineJoin = "round";
                    break;
                }
                case AscFormat.LineJoinType.Bevel:
                {
                    this.Graphics.m_oContext.lineJoin = "bevel";
                    break;
                }
                case AscFormat.LineJoinType.Empty:
                {
                    this.Graphics.m_oContext.lineJoin = "miter";
                    break;
                }
                case AscFormat.LineJoinType.Miter:
                {
                    this.Graphics.m_oContext.lineJoin = "miter";
                    break;
                }
            }
        }

		var arr = this.Graphics.isTrack() ? this.Graphics.Graphics.ArrayPoints : this.Graphics.ArrayPoints;
        var isArrowsPresent = (arr != null && arr.length > 1 && this.IsCurrentPathCanArrows === true) ? true : false;

        var rgba = this.StrokeUniColor;
        let nAlpha = 0xFF;
        if(!isArrowsPresent && !this.IsArrowsDrawing || Asc.editor.isPdfEditor() || this.Shape.isShadowSp)
        {
            if (this.Ln && this.Ln.Fill != null && this.Ln.Fill.transparent != null)
                nAlpha = this.Ln.Fill.transparent;
        }

        this.Graphics.p_color(rgba.R, rgba.G, rgba.B, nAlpha);

        if (this.IsRectShape && this.Graphics.AddSmartRect !== undefined)
        {
            if (undefined !== this.Shape.extX)
                this.Graphics.AddSmartRect(0, 0, this.Shape.extX, this.Shape.extY, this.StrokeWidth);
            else
                this.Graphics.ds();
        }
        else
        {
            this.Graphics.ds();
        }

        if (null != this.OldLineJoin && !this.IsArrowsDrawing)
        {
            this.Graphics.m_oContext.lineJoin = this.OldLineJoin;
        }

        if (isArrowsPresent)
        {
            this.IsArrowsDrawing = true;
            this.Graphics.p_dash(null);
            // значит стрелки есть. теперь:
            // определяем толщину линии "как есть"
            // трансформируем точки в окончательные.
            // и отправляем на отрисовку (с матрицей)

			var _graphicsCtx = this.Graphics.isTrack() ? this.Graphics.Graphics : this.Graphics;

			var trans = _graphicsCtx.m_oFullTransform;
            var trans1 = AscCommon.global_MatrixTransformer.Invert(trans);

            var x1 = trans.TransformPointX(0, 0);
            var y1 = trans.TransformPointY(0, 0);
            var x2 = trans.TransformPointX(1, 1);
            var y2 = trans.TransformPointY(1, 1);
            var dKoef = Math.sqrt(((x2-x1)*(x2-x1)+(y2-y1)*(y2-y1))/2);
            var _pen_w = this.Graphics.isTrack() ? (this.Graphics.Graphics.m_oContext.lineWidth * dKoef) : (this.Graphics.m_oContext.lineWidth * dKoef);
			var _max_w = undefined;
			if (_graphicsCtx.IsThumbnail === true)
			    _max_w = 2;

            var _max_delta_eps2 = 0.001;

            var arrKoef = this.isArrPix ? (1 / AscCommon.g_dKoef_mm_to_pix) : 1;

            if (this.Ln.headEnd != null)
            {
                var _x1 = trans.TransformPointX(arr[0].x, arr[0].y);
                var _y1 = trans.TransformPointY(arr[0].x, arr[0].y);
                var _x2 = trans.TransformPointX(arr[1].x, arr[1].y);
                var _y2 = trans.TransformPointY(arr[1].x, arr[1].y);

                var _max_delta_eps = Math.max(this.Ln.headEnd.GetLen(_pen_w), 5);

                var _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                var cur_point = 2;
                while (_max_delta < _max_delta_eps && cur_point < arr.length)
                {
                    _x2 = trans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
                    _y2 = trans.TransformPointY(arr[cur_point].x, arr[cur_point].y);
                    _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                    cur_point++;
                }

                if (_max_delta > _max_delta_eps2)
                {
					_graphicsCtx.ArrayPoints = null;
					DrawLineEnd(_x1, _y1, _x2, _y2, this.Ln.headEnd.type, arrKoef * this.Ln.headEnd.GetWidth(_pen_w, _max_w), arrKoef * this.Ln.headEnd.GetLen(_pen_w, _max_w), this, trans1);
					_graphicsCtx.ArrayPoints = arr;
                }
            }
            if (this.Ln.tailEnd != null)
            {
                var _1 = arr.length-1;
                var _2 = arr.length-2;
                var _x1 = trans.TransformPointX(arr[_1].x, arr[_1].y);
                var _y1 = trans.TransformPointY(arr[_1].x, arr[_1].y);
                var _x2 = trans.TransformPointX(arr[_2].x, arr[_2].y);
                var _y2 = trans.TransformPointY(arr[_2].x, arr[_2].y);

                var _max_delta_eps = Math.max(this.Ln.tailEnd.GetLen(_pen_w), 5);

                var _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                var cur_point = _2 - 1;
                while (_max_delta < _max_delta_eps && cur_point >= 0)
                {
                    _x2 = trans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
                    _y2 = trans.TransformPointY(arr[cur_point].x, arr[cur_point].y);
                    _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                    cur_point--;
                }

                if (_max_delta > _max_delta_eps2)
                {
					_graphicsCtx.ArrayPoints = null;
					DrawLineEnd(_x1, _y1, _x2, _y2, this.Ln.tailEnd.type, arrKoef * this.Ln.tailEnd.GetWidth(_pen_w, _max_w), arrKoef * this.Ln.tailEnd.GetLen(_pen_w, _max_w), this, trans1);
					_graphicsCtx.ArrayPoints = arr;
                }
            }
            this.IsArrowsDrawing = false;
            this.CheckDash();
        }
    },

    drawFillStroke : function(bIsFill, fill_mode, bIsStroke)
    {
        if (this.Graphics.isTrack())
            this.Graphics.m_oOverlay.ClearAll = true;

        if(this.Graphics.isBoundsChecker())
            return;
        if (!this.isPdf())
        {
            if (bIsFill)
                this.df(fill_mode);
            if (bIsStroke)
                this.ds();
        }
        else
        {
            if (this.bIsNoStrokeAttack)
                bIsStroke = false;

			var arr = this.Graphics.ArrayPoints;
			var isArrowsPresent = (arr != null && arr.length > 1 && this.IsCurrentPathCanArrows === true) ? true : false;

            if (bIsStroke)
            {
                if (null != this.OldLineJoin && !this.IsArrowsDrawing)
                {
                    this.Graphics.put_PenLineJoin(AscFormat.ConvertJoinAggType(this.Ln.Join.type));
                }

                var rgba = this.StrokeUniColor;
                let nAlpha = 0xFF;
                if(!isArrowsPresent && !this.IsArrowsDrawing)
                {
                    if (this.Ln && this.Ln.Fill != null && this.Ln.Fill.transparent != null)
                        nAlpha = this.Ln.Fill.transparent;
                }

                this.Graphics.p_color(rgba.R, rgba.G, rgba.B, nAlpha);

            }

            if (fill_mode == "none" || this.bIsNoFillAttack)
                bIsFill = false;

            var bIsPattern = false;

            if (bIsFill)
            {
                if (this.bIsTexture)
                {
                    if (null == this.UniFill.fill.tile)
                    {
                        if (null == this.UniFill.fill.srcRect)
                        {
                            if (this.UniFill.fill.RasterImageId && this.UniFill.fill.RasterImageId.indexOf(".svg") != 0)
                            {
                                this.Graphics.put_brushTexture(getFullImageSrc2(this.UniFill.fill.RasterImageId), 0);
                            }
                            else
                            {
                                if (this.UniFill.fill.canvas)
                                {
                                    this.Graphics.put_brushTexture(this.UniFill.fill.canvas.toDataURL("image/png"), 0);
                                }
                                else
                                {
                                    this.Graphics.put_brushTexture(getFullImageSrc2(this.UniFill.fill.RasterImageId), 0);
                                }
                            }
                        }
                        else
                        {
                            if (this.IsRectShape)
                            {
                                this.Graphics.drawImage(getFullImageSrc2(this.UniFill.fill.RasterImageId), this.min_x, this.min_y, (this.max_x - this.min_x), (this.max_y - this.min_y), undefined, this.UniFill.fill.srcRect);
                                bIsFill = false;
                            }
                            else
                            {
                                // TODO: support srcRect
                                this.Graphics.put_brushTexture(getFullImageSrc2(this.UniFill.fill.RasterImageId), 0);
                            }
                        }
                    }
                    else
                    {
                        if (this.UniFill.fill.canvas)
                        {
                            this.Graphics.put_brushTexture(this.UniFill.fill.canvas.toDataURL("image/png"), 1);
                        }
                        else
                        {
                            this.Graphics.put_brushTexture(getFullImageSrc2(this.UniFill.fill.RasterImageId), 1);
                        }
                    }
                    this.Graphics.put_BrushTextureAlpha(this.UniFill.transparent);
                }
                else
                {
                    var _fill = this.UniFill.fill;
                    if (_fill.type == c_oAscFill.FILL_TYPE_PATT)
                    {
                        var _patt_name = AscCommon.global_hatch_names[_fill.ftype];
                        if (undefined == _patt_name)
                            _patt_name = "cross";

                        var _fc = _fill.fgClr && _fill.fgClr.RGBA || {R: 0, G: 0, B: 0, A: 255};
                        var _bc = _fill.bgClr && _fill.bgClr.RGBA || {R: 255, G: 255, B: 255, A: 255};

                        var __fa = (null === this.UniFill.transparent) ? _fc.A : 255;
                        var __ba = (null === this.UniFill.transparent) ? _bc.A : 255;

                        var _pattern = AscCommon.GetHatchBrush(_patt_name, _fc.R, _fc.G, _fc.B, __fa, _bc.R, _bc.G, _bc.B, __ba);

                        var _url64 = "";
                        try
                        {
                            _url64 = _pattern.toDataURL();
                        }
                        catch (err)
                        {
                            _url64 = "";
                        }

                        this.Graphics.put_brushTexture(_url64, 1);

                        if (null != this.UniFill.transparent)
                            this.Graphics.put_BrushTextureAlpha(this.UniFill.transparent);
                        else
                            this.Graphics.put_BrushTextureAlpha(255);

                        bIsPattern = true;
                    }
                    else if (_fill.type == c_oAscFill.FILL_TYPE_GRAD)
                    {
                        var points = null;
                        if (_fill.lin)
                        {
                            var _angle = _fill.lin.angle;
                            if (_fill.rotateWithShape === false && this.Graphics.m_oTransform)
                            {
                                //_angle -= (60000 * this.Graphics.m_oTransform.GetRotation());
                                _angle = AscCommon.GradientGetAngleNoRotate(_angle, this.Graphics.m_oTransform);
                            }

                            points = this.getGradientPoints(this.min_x, this.min_y, this.max_x, this.max_y, _angle, _fill.lin.scale);
                        }
                        else if (_fill.path)
                        {
                            var _cx = (this.min_x + this.max_x) / 2;
                            var _cy = (this.min_y + this.max_y) / 2;
                            var _r = Math.max(this.max_x - this.min_x, this.max_y - this.min_y) / 2;

                            points = { x0 : _cx, y0 : _cy, x1 : _cx, y1 : _cy, r0 : 1, r1 : _r };
                        }
                        else
                        {
                            points = this.getGradientPoints(this.min_x, this.min_y, this.max_x, this.max_y, 0, false);
                        }

                        this.Graphics.put_BrushGradient(_fill, points, this.UniFill.transparent);
                    }
                    else
                    {
                        var rgba = this.FillUniColor;
                        if (fill_mode == "darken")
                        {
                            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
                            var rgb = _color1.darken();
                            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
                        }
                        else if (fill_mode == "darkenLess")
                        {
                            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
                            var rgb = _color1.darkenLess();
                            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
                        }
                        else if (fill_mode == "lighten")
                        {
                            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
                            var rgb = _color1.lighten();
                            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
                        }
                        else if (fill_mode == "lightenLess")
                        {
                            var _color1 = new CShapeColor(rgba.R, rgba.G, rgba.B);
                            var rgb = _color1.lightenLess();
                            rgba = { R: rgb.r, G: rgb.g, B: rgb.b, A: rgba.A };
                        }
                        if (rgba)
                        {
                            if (this.UniFill != null && this.UniFill.transparent != null)
                                rgba.A = this.UniFill.transparent;
                            this.Graphics.b_color1(rgba.R, rgba.G, rgba.B, rgba.A);
                        }
                    }
                }
            }

            if (bIsFill && bIsStroke)
            {
                if (this.bIsTexture || bIsPattern)
                {
                    this.Graphics.drawpath(256);
                    this.Graphics.drawpath(1);
                }
                else
                {
                    this.Graphics.drawpath(256 + 1);
                }
            }
            else if (bIsFill)
            {
                this.Graphics.drawpath(256);
            }
            else if (bIsStroke)
            {
                this.Graphics.drawpath(1);
            }
            else if (false)
            {
                // такого быть не должно по идее
                this.Graphics.b_color1(0, 0, 0, 0);
                this.Graphics.drawpath(256);
            }

            if (isArrowsPresent)
            {
                this.IsArrowsDrawing = true;
                this.Graphics.p_dash(null);
                // значит стрелки есть. теперь:
                // определяем толщину линии "как есть"
                // трансформируем точки в окончательные.
                // и отправляем на отрисовку (с матрицей)

                var trans = (!this.isPdf()) ? this.Graphics.m_oFullTransform : this.Graphics.GetTransform();
                var trans1 = AscCommon.global_MatrixTransformer.Invert(trans);

                var lineSize = (!this.isPdf()) ? this.Graphics.m_oContext.lineWidth : this.Graphics.GetLineWidth();

                var x1 = trans.TransformPointX(0, 0);
                var y1 = trans.TransformPointY(0, 0);
                var x2 = trans.TransformPointX(1, 1);
                var y2 = trans.TransformPointY(1, 1);
                var dKoef = Math.sqrt(((x2-x1)*(x2-x1)+(y2-y1)*(y2-y1))/2);
                var _pen_w = lineSize * dKoef;

                var _pen_w_max = 2.5 / AscCommon.g_dKoef_mm_to_pix;

                if (this.Ln.headEnd != null)
                {
                    var _x1 = trans.TransformPointX(arr[0].x, arr[0].y);
                    var _y1 = trans.TransformPointY(arr[0].x, arr[0].y);
                    var _x2 = trans.TransformPointX(arr[1].x, arr[1].y);
                    var _y2 = trans.TransformPointY(arr[1].x, arr[1].y);

                    var _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                    var cur_point = 2;
                    while (_max_delta < 0.001 && cur_point < arr.length)
                    {
                        _x2 = trans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
                        _y2 = trans.TransformPointY(arr[cur_point].x, arr[cur_point].y);
                        _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                        cur_point++;
                    }

                    if (_max_delta > 0.001)
                    {
                        this.Graphics.ArrayPoints = null;
                        DrawLineEnd(_x1, _y1, _x2, _y2, this.Ln.headEnd.type, this.Ln.headEnd.GetWidth(_pen_w, _pen_w_max), this.Ln.headEnd.GetLen(_pen_w, _pen_w_max), this, trans1);
                        this.Graphics.ArrayPoints = arr;
                    }
                }
                if (this.Ln.tailEnd != null)
                {
                    var _1 = arr.length-1;
                    var _2 = arr.length-2;
                    var _x1 = trans.TransformPointX(arr[_1].x, arr[_1].y);
                    var _y1 = trans.TransformPointY(arr[_1].x, arr[_1].y);
                    var _x2 = trans.TransformPointX(arr[_2].x, arr[_2].y);
                    var _y2 = trans.TransformPointY(arr[_2].x, arr[_2].y);

                    var _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                    var cur_point = _2 - 1;
                    while (_max_delta < 0.001 && cur_point >= 0)
                    {
                        _x2 = trans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
                        _y2 = trans.TransformPointY(arr[cur_point].x, arr[cur_point].y);
                        _max_delta = Math.max(Math.abs(_x1 - _x2), Math.abs(_y1 - _y2));
                        cur_point--;
                    }

                    if (_max_delta > 0.001)
                    {
                        this.Graphics.ArrayPoints = null;
                        DrawLineEnd(_x1, _y1, _x2, _y2, this.Ln.tailEnd.type, this.Ln.tailEnd.GetWidth(_pen_w, _pen_w_max), this.Ln.tailEnd.GetLen(_pen_w, _pen_w_max), this, trans1);
                        this.Graphics.ArrayPoints = arr;
                    }
                }
                this.IsArrowsDrawing = false;
                this.CheckDash();
            }
        }
    },

    drawStrokeFillStyle : function()
    {
        if (!this.isPdf())
        {
            var gr = this.Graphics.isTrack() ? this.Graphics.Graphics : this.Graphics;

            var tmp = gr.m_oBrush.Color1;
            var p_c = gr.m_oPen.Color;
            gr.b_color1(p_c.R, p_c.G, p_c.B, p_c.A);
            gr.df();
            gr.b_color1(tmp.R, tmp.G, tmp.B, tmp.A);
        }
        else
        {
            var tmp = this.Graphics.GetBrush().Color1;
            var p_c = this.Graphics.GetPen().Color;
            this.Graphics.b_color1(p_c.R, p_c.G, p_c.B, p_c.A);
            this.Graphics.df();
            this.Graphics.b_color1(tmp.R, tmp.G, tmp.B, tmp.A);
        }
    },

    check_bounds : function()
    {
        this.Shape.check_bounds(this);
    },

    // common funcs
    getNormalPoint : function(x0, y0, angle, x1, y1)
    {
        // точка - пересечение прямой, проходящей через точку (x0, y0) под углом angle и
        // перпендикуляра к первой прямой, проведенной из точки (x1, y1)
        var ex1 = Math.cos(angle);
        var ey1 = Math.sin(angle);

        var ex2 = -ey1;
        var ey2 = ex1;

        var a = ex1 / ey1;
        var b = ex2 / ey2;

        var x = ((a * b * y1 - a * b * y0) - (a * x1 - b * x0)) / (b - a);
        var y = (x - x0) / a + y0;

        return { X : x, Y : y };
    },

    getGradientPoints : function(min_x, min_y, max_x, max_y, _angle, scale)
    {
        var points = { x0 : 0, y0 : 0, x1 : 0, y1 : 0 };

        var angle = _angle / 60000;
        while (angle < 0)
            angle += 360;
        while (angle >= 360)
            angle -= 360;

        if (Math.abs(angle) < 1)
        {
            points.x0 = min_x;
            points.y0 = min_y;
            points.x1 = max_x;
            points.y1 = min_y;

            return points;
        }
        else if (Math.abs(angle - 90) < 1)
        {
            points.x0 = min_x;
            points.y0 = min_y;
            points.x1 = min_x;
            points.y1 = max_y;

            return points;
        }
        else if (Math.abs(angle - 180) < 1)
        {
            points.x0 = max_x;
            points.y0 = min_y;
            points.x1 = min_x;
            points.y1 = min_y;

            return points;
        }
        else if (Math.abs(angle - 270) < 1)
        {
            points.x0 = min_x;
            points.y0 = max_y;
            points.x1 = min_x;
            points.y1 = min_y;

            return points;
        }

        var grad_a = AscCommon.deg2rad(angle);
        if (!scale)
        {
            if (angle > 0 && angle < 90)
            {
                var p = this.getNormalPoint(min_x, min_y, grad_a, max_x, max_y);

                points.x0 = min_x;
                points.y0 = min_y;
                points.x1 = p.X;
                points.y1 = p.Y;

                return points;
            }
            if (angle > 90 && angle < 180)
            {
                var p = this.getNormalPoint(max_x, min_y, grad_a, min_x, max_y);

                points.x0 = max_x;
                points.y0 = min_y;
                points.x1 = p.X;
                points.y1 = p.Y;

                return points;
            }
            if (angle > 180 && angle < 270)
            {
                var p = this.getNormalPoint(max_x, max_y, grad_a, min_x, min_y);

                points.x0 = max_x;
                points.y0 = max_y;
                points.x1 = p.X;
                points.y1 = p.Y;

                return points;
            }
            if (angle > 270 && angle < 360)
            {
                var p = this.getNormalPoint(min_x, max_y, grad_a, max_x, min_y);

                points.x0 = min_x;
                points.y0 = max_y;
                points.x1 = p.X;
                points.y1 = p.Y;

                return points;
            }
            // никогда сюда не зайдем
            return points;
        }

        // scale == true
        var _grad_45 = (Math.PI / 2) - Math.atan2(max_y - min_y, max_x - min_x);
        var _grad_90_45 = (Math.PI / 2) - _grad_45;
        if (angle > 0 && angle < 90)
        {
            if (angle <= 45)
            {
                grad_a = (_grad_45 * angle / 45);
            }
            else
            {
                grad_a = _grad_45 + _grad_90_45 * (angle - 45) / 45;
            }

            var p = this.getNormalPoint(min_x, min_y, grad_a, max_x, max_y);

            points.x0 = min_x;
            points.y0 = min_y;
            points.x1 = p.X;
            points.y1 = p.Y;

            return points;
        }
        if (angle > 90 && angle < 180)
        {
            if (angle <= 135)
            {
                grad_a = Math.PI / 2 + _grad_90_45 * (angle - 90) / 45;
            }
            else
            {
				grad_a = Math.PI / 2 + _grad_90_45 + _grad_45 * (angle - 135) / 45;
            }

            var p = this.getNormalPoint(max_x, min_y, grad_a, min_x, max_y);

            points.x0 = max_x;
            points.y0 = min_y;
            points.x1 = p.X;
            points.y1 = p.Y;

            return points;
        }
        if (angle > 180 && angle < 270)
        {
            if (angle <= 225)
            {
                grad_a = Math.PI + _grad_45 * (angle - 180) / 45;
            }
            else
            {
				grad_a = Math.PI + _grad_45 + _grad_90_45 * (angle - 225) / 45;
            }

            var p = this.getNormalPoint(max_x, max_y, grad_a, min_x, min_y);

            points.x0 = max_x;
            points.y0 = max_y;
            points.x1 = p.X;
            points.y1 = p.Y;

            return points;
        }
        if (angle > 270 && angle < 360)
        {
            if (angle <= 315)
            {
                grad_a = 3 * Math.PI / 2 + _grad_90_45 * (angle - 270) / 45;
            }
            else
            {
				grad_a = 3 * Math.PI / 2 + _grad_90_45 + _grad_45 * (angle - 315) / 45;
            }

            var p = this.getNormalPoint(min_x, max_y, grad_a, max_x, min_y);

            points.x0 = min_x;
            points.y0 = max_y;
            points.x1 = p.X;
            points.y1 = p.Y;

            return points;
        }
        // никогда сюда не зайдем
        return points;
    },

    DrawPresentationComment : function(type, x, y, w, h)
    {
    },

    getEditorInfo: function()
    {
        var _ret = {};
        _ret.editor = Asc.editor || window.editor;
        switch (_ret.editor.editorId)
        {
            case AscCommon.c_oEditorId.Word:
            case AscCommon.c_oEditorId.Presentation:
            {
                _ret.scale = _ret.editor.WordControl.m_nZoomValue / 100;
                break;
            }
            case AscCommon.c_oEditorId.Spreadsheet:
            {
                _ret.scale = _ret.editor.asc_getZoom();
                break;
            }
            default:
                break;
        }
        return _ret;
    }
};

function ShapeToImageConverter(shape, pageIndex, sImageFormat)
{
    AscCommon.IsShapeToImageConverter = true;
    var _bounds_cheker = new AscFormat.CSlideBoundsChecker();

    var dKoef = AscCommon.g_dKoef_mm_to_pix;
    var w_mm = 210;
    var h_mm = 297;
    var w_px = (w_mm * dKoef) >> 0;
    var h_px = (h_mm * dKoef) >> 0;

    _bounds_cheker.init(w_px, h_px, w_mm, h_mm);
    _bounds_cheker.transform(1,0,0,1,0,0);

    _bounds_cheker.AutoCheckLineWidth = true;
	_bounds_cheker.CheckLineWidth(shape);
    shape.draw(_bounds_cheker, /*pageIndex*/0);
	_bounds_cheker.CorrectBounds2();

    var _need_pix_width     = _bounds_cheker.Bounds.max_x - _bounds_cheker.Bounds.min_x + 1;
    var _need_pix_height    = _bounds_cheker.Bounds.max_y - _bounds_cheker.Bounds.min_y + 1;

    if (_need_pix_width <= 0 || _need_pix_height <= 0)
        return null;

    /*
    if (shape.pen)
    {
        var _w_pen = (shape.pen.w == null) ? 12700 : parseInt(shape.pen.w);
        _w_pen /= 36000.0;
        _w_pen *= g_dKoef_mm_to_pix;

        _need_pix_width += (2 * _w_pen);
        _need_pix_height += (2 * _w_pen);

        _bounds_cheker.Bounds.min_x -= _w_pen;
        _bounds_cheker.Bounds.min_y -= _w_pen;
    }*/

	var _canvas = null;
	if (window["NATIVE_EDITOR_ENJINE"] === true && window["IS_NATIVE_EDITOR"] !== true)
	{
		_need_pix_width = _need_pix_width >> 0;
		_need_pix_height = _need_pix_height >> 0;
		_canvas = new CNativeGraphics();
		_canvas.width  = _need_pix_width;
		_canvas.height = _need_pix_height;
		_canvas.create(window["native"], _need_pix_width, _need_pix_height, _need_pix_width / dKoef, _need_pix_height / dKoef);
		_canvas.CoordTransformOffset(-_bounds_cheker.Bounds.min_x, -_bounds_cheker.Bounds.min_y);
		_canvas.transform(1, 0, 0, 1, 0, 0);
		shape.draw(_canvas, 0);
	}
	else
	{
		_canvas = document.createElement("canvas");
		_canvas.width = _need_pix_width >> 0;
		_canvas.height = _need_pix_height >> 0;
		var _ctx = _canvas.getContext("2d");
		var g = new AscCommon.CGraphics;
		g.init(_ctx, w_px, h_px, w_mm, h_mm);
		g.m_oFontManager = AscCommon.g_fontManager;
		g.m_oCoordTransform.tx = -_bounds_cheker.Bounds.min_x;
		g.m_oCoordTransform.ty = -_bounds_cheker.Bounds.min_y;
		g.transform(1, 0, 0, 1, 0, 0);
		shape.draw(g, 0);
	}

    if (AscCommon.g_fontManager) {
        AscCommon.g_fontManager.m_pFont = null;
    }
    if (AscCommon.g_fontManager2) {
        AscCommon.g_fontManager2.m_pFont = null;
    }
    AscCommon.IsShapeToImageConverter = false;

    var _ret = { ImageNative : _canvas, ImageUrl : "" };
    try
    {
        const sFormat = sImageFormat || "image/png";
        _ret.ImageUrl = _canvas.toDataURL(sFormat);
    }
    catch (err)
    {
        if (shape.brush != null && shape.brush.fill && shape.brush.fill.RasterImageId)
            _ret.ImageUrl = getFullImageSrc2(shape.brush.fill.RasterImageId);
        else
            _ret.ImageUrl = "";
    }
    if (_canvas.isNativeGraphics === true)
        _canvas.Destroy();
    return _ret;
}

//------------------------------------------------------------export----------------------------------------------------
window['AscCommon'] = window['AscCommon'] || {};
window['AscCommon'].CShapeDrawer = CShapeDrawer;
window['AscCommon'].ShapeToImageConverter = ShapeToImageConverter;
window['AscCommon'].IsShapeToImageConverter = false;
window['AscCommon'].DrawLineEnd = DrawLineEnd;
})(window);
