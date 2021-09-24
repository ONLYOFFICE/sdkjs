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

"use strict";

(function(window, undefined){

    function CPoint(x, y, bTemporary) {
        this.x = x;
        this.y = y;
        this.bTemporary = bTemporary === true;
    }
    CPoint.prototype.reset = function(x, y, bTemporary) {
        this.x = x;
        this.y = y;
        this.bTemporary = bTemporary === true;
    };
    CPoint.prototype.distance = function(x, y) {
        var dx = this.x - x;
        var dy = this.y - y;
        return Math.sqrt(dx*dx + dy*dy);
    };
    CPoint.prototype.distanceFromOther = function(oPoint) {
        return this.distance(oPoint.x, oPoint.y);
    };
    CPoint.prototype.isNear = function(x, y) {
        return this.distance(x, y) < 1;
    };
function PolyLine (drawingObjects, theme, master, layout, slide, pageIndex)
{

    AscFormat.ExecuteNoHistory(function(){

        this.drawingObjects = drawingObjects;
        this.arrPoint = [];
        this.Matrix = new AscCommon.CMatrixL();
        this.TransformMatrix = new AscCommon.CMatrixL();

        this.pageIndex = pageIndex;
        this.style  = AscFormat.CreateDefaultShapeStyle();
        var style = this.style;
        style.fillRef.Color.Calculate(theme, slide, layout, master, {R:0, G: 0, B:0, A:255});
        var RGBA = style.fillRef.Color.RGBA;
        var pen = theme.getLnStyle(style.lnRef.idx, style.lnRef.Color);
        style.lnRef.Color.Calculate(theme, slide, layout, master);
        RGBA = style.lnRef.Color.RGBA;

        if(pen.Fill)
        {
            pen.Fill.calculate(theme, slide, layout, master, RGBA);
        }

        this.pen = pen;

        this.polylineForDrawer = new PolylineForDrawer(this);
        this.continuousRanges = [];

    }, this, []);
}

    PolyLine.prototype.Draw = function(graphics)
    {
        graphics.SetIntegerGrid(false);
        graphics.transform3(this.Matrix);

        var shape_drawer = new AscCommon.CShapeDrawer();
        shape_drawer.fromShape(this, graphics);
        shape_drawer.draw(this);
    };
    PolyLine.prototype.draw = function(g)
    {
        if(AscFormat.isRealNumber(this.pageIndex) && g.SetCurrentPage)
        {
            g.SetCurrentPage(this.pageIndex);
        }
        this.polylineForDrawer.Draw(g);
    };

    PolyLine.prototype.getBounds = function()
    {
        var boundsChecker = new  AscFormat.CSlideBoundsChecker();
        this.draw(boundsChecker);
        boundsChecker.Bounds.posX = boundsChecker.Bounds.min_x;
        boundsChecker.Bounds.posY = boundsChecker.Bounds.min_y;
        boundsChecker.Bounds.extX = boundsChecker.Bounds.max_x - boundsChecker.Bounds.min_x;
        boundsChecker.Bounds.extY = boundsChecker.Bounds.max_y - boundsChecker.Bounds.min_y;
        return boundsChecker.Bounds;
    };
    PolyLine.prototype.getShape =  function(bWord, drawingDocument, drawingObjects)
    {
        var xMax = this.arrPoint[0].x, yMax = this.arrPoint[0].y, xMin = xMax, yMin = yMax;
        var i;

        var bClosed = false;
        var min_dist;
        if(drawingObjects)
        {
            min_dist = drawingObjects.convertPixToMM(3);
        }
        else
        {
            min_dist = editor.WordControl.m_oDrawingDocument.GetMMPerDot(3)
        }
        var oLastPoint = this.arrPoint[this.arrPoint.length-1];
        var nLastIndex = this.arrPoint.length-1;
        if(oLastPoint.bTemporary) {
            nLastIndex--;
        }
        if(nLastIndex > 1)
        {
            var dx = this.arrPoint[0].x - this.arrPoint[nLastIndex].x;
            var dy = this.arrPoint[0].y - this.arrPoint[nLastIndex].y;
            if(Math.sqrt(dx*dx +dy*dy) < min_dist)
            {
                bClosed = true;
            }
        }
        var _n = bClosed ? nLastIndex : nLastIndex + 1;
        for( i = 1; i<_n; ++i)
        {
            if(this.arrPoint[i].x > xMax)
            {
                xMax = this.arrPoint[i].x;
            }
            if(this.arrPoint[i].y > yMax)
            {
                yMax = this.arrPoint[i].y;
            }

            if(this.arrPoint[i].x < xMin)
            {
                xMin = this.arrPoint[i].x;
            }

            if(this.arrPoint[i].y < yMin)
            {
                yMin = this.arrPoint[i].y;
            }
        }




        var shape = new AscFormat.CShape();

        //  if(drawingObjects)
        //  {
        //      shape.setWorksheet(drawingObjects.getWorksheetModel());
        //      shape.addToDrawingObjects();
        //  }
        shape.setSpPr(new AscFormat.CSpPr());
        shape.spPr.setParent(shape);
        shape.spPr.setXfrm(new AscFormat.CXfrm());
        shape.spPr.xfrm.setParent(shape.spPr);
        if(!bWord)
        {
            shape.spPr.xfrm.setOffX(xMin);
            shape.spPr.xfrm.setOffY(yMin);
        }
        else
        {
            shape.setWordShape(true);
            shape.spPr.xfrm.setOffX(0);
            shape.spPr.xfrm.setOffY(0);
        }
        shape.spPr.xfrm.setExtX(xMax-xMin);
        shape.spPr.xfrm.setExtY(yMax - yMin);
        shape.setStyle(AscFormat.CreateDefaultShapeStyle());
        var geometry = new AscFormat.Geometry();


        var w = xMax - xMin, h = yMax-yMin;
        var kw, kh, pathW, pathH;
        if(w > 0)
        {
            pathW = 43200;
            kw = 43200/ w;
        }
        else
        {
            pathW = 0;
            kw = 0;
        }
        if(h > 0)
        {
            pathH = 43200;
            kh = 43200 / h;
        }
        else
        {
            pathH = 0;
            kh = 0;
        }
        geometry.AddPathCommand(0, undefined, bClosed ? "norm": "none", undefined, pathW, pathH);
        geometry.AddRect("l", "t", "r", "b");
        geometry.AddPathCommand(1, (((this.arrPoint[0].x - xMin) * kw) >> 0) + "", (((this.arrPoint[0].y - yMin) * kh) >> 0) + "");
        i = 1;
        var aRanges = this.continuousRanges;
        var aRange, nRange;
        var nEnd;
        var nPtsCount = this.arrPoint.length;
        var oPt1, oPt2, oPt3, nPt;
        for(nRange = 0; nRange < aRanges.length; ++nRange)
        {
            aRange = aRanges[nRange];
            nEnd = aRange[1];
            nPt = aRange[0] + 1;
            while(nPt <= nEnd)
            {
                if(nPt + 2 <= nEnd)
                {
                    //cubic bezier curve
                    oPt1 = this.arrPoint[nPt++];
                    oPt2 = this.arrPoint[nPt++];
                    oPt3 = this.arrPoint[nPt++];
                    geometry.AddPathCommand(5,
                        (((oPt1.x - xMin) * kw) >> 0) + "", (((oPt1.y - yMin) * kh) >> 0) + "",
                        (((oPt2.x - xMin) * kw) >> 0) + "", (((oPt2.y - yMin) * kh) >> 0) + "",
                        (((oPt3.x - xMin) * kw) >> 0) + "", (((oPt3.y - yMin) * kh) >> 0) + ""
                    );
                }
                else if(nPt + 1 <= nEnd)
                {
                    //quad bezier curve
                    oPt1 = this.arrPoint[nPt++];
                    oPt2 = this.arrPoint[nPt++];
                    geometry.AddPathCommand(4,
                        (((oPt1.x - xMin) * kw) >> 0) + "", (((oPt1.y - yMin) * kh) >> 0) + "",
                        (((oPt2.x - xMin) * kw) >> 0) + "", (((oPt2.y - yMin) * kh) >> 0) + ""
                    );
                }
                else
                {
                    //lineTo
                    oPt1 = this.arrPoint[nPt++];
                    geometry.AddPathCommand(2,
                        (((oPt1.x - xMin) * kw) >> 0) + "", (((oPt1.y - yMin) * kh) >> 0) + ""
                    );
                }
            }
        }
        if(bClosed)
        {
            geometry.AddPathCommand(6);
        }

        shape.spPr.setGeometry(geometry);
        shape.setBDeleted(false);
        shape.recalculate();
        shape.x = xMin;
        shape.y = yMin;
        return shape;
    };
    PolyLine.prototype.tryAddPoint = function(x, y)
    {
        var oLastPoint = this.arrPoint[this.arrPoint.length - 1];
        if(!oLastPoint) {
            this.addPoint(x, y);
            return;
        }
        if(oLastPoint.isNear(x, y)) {
            //oLastPoint.reset(x, y);
            return;
        }
        this.addPoint(x, y);
    };

    PolyLine.prototype.createContinuousRange = function()
    {
        var nIdx = this.arrPoint.length - 1;
        this.continuousRanges.push([nIdx, nIdx]);
    };
    PolyLine.prototype.getLastContinuousRange = function()
    {
        if(this.continuousRanges.length === 0) {
            this.createContinuousRange();
        }
        return this.continuousRanges[this.continuousRanges.length - 1];
    };
    PolyLine.prototype.addPoint = function(x, y, bTemporary)
    {
        this.arrPoint.push(new CPoint(x, y, bTemporary));
        var oLastRange = this.getLastContinuousRange();
        oLastRange[1] = this.arrPoint.length - 1;
    };
    PolyLine.prototype.replaceLastPoint = function(x, y, bTemporary)
    {
        var oLastPoint = this.arrPoint[this.arrPoint.length - 1];
        if(!oLastPoint) {
            this.addPoint(x, y, bTemporary);
            return;
        }
        oLastPoint.reset(x, y, bTemporary);
        var oLastRange = this.getLastContinuousRange();
        if(oLastRange[0] !== this.arrPoint.length - 1) {
            this.createContinuousRange();
        }
    };
    PolyLine.prototype.canCreateShape = function()
    {
        var nCount = this.arrPoint.length;
        if(nCount < 2) {
            return false;
        }
        var oLast = this.arrPoint[this.arrPoint.length - 1];
        if(oLast.bTemporary) {
            --nCount;
        }
        return nCount > 1;
    };
    PolyLine.prototype.getPointsCount = function()
    {
        return this.arrPoint.length;
    };

function PolylineForDrawer(polyline)
{
    this.polyline = polyline;
    this.pen = polyline.pen;
    this.brush = polyline.brush;
    this.TransformMatrix = polyline.TransformMatrix;
    this.Matrix = polyline.Matrix;

    this.Draw = function(graphics)
    {
        graphics.SetIntegerGrid(false);
        graphics.transform3(this.Matrix);

        var shape_drawer = new AscCommon.CShapeDrawer();
        shape_drawer.fromShape(this, graphics);
        shape_drawer.draw(this);
    };
    this.draw = function(g)
    {
        g._e();
        if(this.polyline.arrPoint.length < 2)
        {
            return;
        }
        g._m(this.polyline.arrPoint[0].x, this.polyline.arrPoint[0].y);
        for(var i = 1; i < this.polyline.arrPoint.length; ++i)
        {
            g._l(this.polyline.arrPoint[i].x, this.polyline.arrPoint[i].y);
        }
        g.ds();
    };
}

    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].PolyLine = PolyLine;
})(window);
