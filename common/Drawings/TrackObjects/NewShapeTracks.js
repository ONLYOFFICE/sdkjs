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

    var MIN_SHAPE_SIZE = 1.27;//размер меньше которого нельзя уменшить автофигуру или картинку по горизонтали или вертикали
    var MIN_SHAPE_SIZE_DIV2 = MIN_SHAPE_SIZE/2.0;

    var SHAPE_ASPECTS = {};
    SHAPE_ASPECTS["can"] = 3616635/4810125;
    SHAPE_ASPECTS["moon"] = 457200/914400;
    SHAPE_ASPECTS["leftBracket"] = 73152/914400;
    SHAPE_ASPECTS["rightBracket"] = 73152/914400;
    SHAPE_ASPECTS["leftBrace"] = 155448/914400;
    SHAPE_ASPECTS["rightBrace"] = 155448/914400;
    SHAPE_ASPECTS["triangle"] = 1060704/914400;
    SHAPE_ASPECTS["parallelogram"] = 1216152/914400;
    SHAPE_ASPECTS["trapezoid"] = 914400/1216152;
    SHAPE_ASPECTS["pentagon"] = 960120/914400;
    SHAPE_ASPECTS["hexagon"] = 1060704/914400;
    SHAPE_ASPECTS["bracePair"] = 1069848/914400;
    SHAPE_ASPECTS["rightArrow"] = 978408/484632;
    SHAPE_ASPECTS["leftArrow"] = 978408/484632;
    SHAPE_ASPECTS["upArrow"] = 484632/978408;
    SHAPE_ASPECTS["downArrow"] = 484632/978408;
    SHAPE_ASPECTS["leftRightArrow"] = 1216152/484632;
    SHAPE_ASPECTS["upDownArrow"] = 484632/1216152;
    SHAPE_ASPECTS["bentArrow"] = 813816/868680;
    SHAPE_ASPECTS["uturnArrow"] = 886968/877824;
    SHAPE_ASPECTS["bentUpArrow"] = 850392/731520;
    SHAPE_ASPECTS["curvedRightArrow"] = 731520/1216152;
    SHAPE_ASPECTS["curvedLeftArrow"] = 731520/1216152;
    SHAPE_ASPECTS["curvedUpArrow"] = 1216152/731520;
    SHAPE_ASPECTS["curvedDownArrow"] = 1216152/731520;
    SHAPE_ASPECTS["stripedRightArrow"] = 978408/484632;
    SHAPE_ASPECTS["notchedRightArrow"] = 978408/484632;
    SHAPE_ASPECTS["homePlate"] = 978408/484632;
    SHAPE_ASPECTS["leftRightArrowCallout"] = 1216152/576072;
    SHAPE_ASPECTS["flowChartProcess"] = 914400/612648;
    SHAPE_ASPECTS["flowChartAlternateProcess"] = 914400/612648;
    SHAPE_ASPECTS["flowChartDecision"] = 914400/612648;
    SHAPE_ASPECTS["flowChartInputOutput"] = 914400/612648;
    SHAPE_ASPECTS["flowChartPredefinedProcess"] = 914400/612648;
    SHAPE_ASPECTS["flowChartDocument"] = 914400/612648;
    SHAPE_ASPECTS["flowChartMultidocument"] = 1060704/758952;
    SHAPE_ASPECTS["flowChartTerminator"] = 914400/301752;
    SHAPE_ASPECTS["flowChartPreparation"] = 1060704/612648;
    SHAPE_ASPECTS["flowChartManualInput"] = 914400/457200;
    SHAPE_ASPECTS["flowChartManualOperation"] = 914400/612648;
    SHAPE_ASPECTS["flowChartPunchedCard"] = 914400/804672;
    SHAPE_ASPECTS["flowChartPunchedTape"] = 914400/804672;
    SHAPE_ASPECTS["flowChartPunchedTape"] = 457200/914400;
    SHAPE_ASPECTS["flowChartSort"] = 457200/914400;
    SHAPE_ASPECTS["flowChartOnlineStorage"] = 914400/612648;
    SHAPE_ASPECTS["flowChartMagneticDisk"] = 914400/612648;
    SHAPE_ASPECTS["flowChartMagneticDrum"] = 914400/685800;
    SHAPE_ASPECTS["flowChartDisplay"] = 914400/612648;
    SHAPE_ASPECTS["ribbon2"] = 1216152/612648;
    SHAPE_ASPECTS["ribbon"] = 1216152/612648;
    SHAPE_ASPECTS["ellipseRibbon2"] = 1216152/758952;
    SHAPE_ASPECTS["ellipseRibbon"] = 1216152/758952;
    SHAPE_ASPECTS["verticalScroll"] = 1033272/1143000;
    SHAPE_ASPECTS["horizontalScroll"] = 1143000/1033272;
    SHAPE_ASPECTS["wedgeRectCallout"] = 914400/612648;
    SHAPE_ASPECTS["wedgeRoundRectCallout"] = 914400/612648;
    SHAPE_ASPECTS["wedgeEllipseCallout"] = 914400/612648;
    SHAPE_ASPECTS["cloudCallout"] = 914400/612648;
    SHAPE_ASPECTS["borderCallout1"] = 914400/612648;
    SHAPE_ASPECTS["borderCallout2"] = 914400/612648;
    SHAPE_ASPECTS["borderCallout3"] = 914400/612648;
    SHAPE_ASPECTS["accentCallout1"] = 914400/612648;
    SHAPE_ASPECTS["accentCallout2"] = 914400/612648;
    SHAPE_ASPECTS["accentCallout3"] = 914400/612648;
    SHAPE_ASPECTS["callout1"] = 914400/612648;
    SHAPE_ASPECTS["callout2"] = 914400/612648;
    SHAPE_ASPECTS["callout3"] = 914400/612648;
    SHAPE_ASPECTS["accentBorderCallout1"] = 914400/612648;
    SHAPE_ASPECTS["accentBorderCallout2"] = 914400/612648;
    SHAPE_ASPECTS["accentBorderCallout3"] = 914400/612648;

function NewShapeTrack(presetGeom, startX, startY, theme, master, layout, slide, pageIndex, drawingsController, nPlaceholderType, bVertical, bSkipCheckConnector)
{
    this.presetGeom = presetGeom;
    this.startX = startX;
    this.startY = startY;

    this.x = null;
    this.y = null;
    this.extX = null;
    this.extY = null;
    this.rot = 0;
    this.arrowsCount = 0;
    this.transform = new AscCommon.CMatrix();
    this.pageIndex = pageIndex;
    this.theme = theme;
    this.drawingsController = drawingsController ? drawingsController : Asc.editor.getGraphicController();

    //for connectors
    this.bConnector = false;
    this.startConnectionInfo = null;
    this.oShapeDrawConnectors = null;
    this.lastSpPr = null;

    this.startShape = null;
    this.endShape = null;
    this.endConnectionInfo = null;
    this.placeholderType = nPlaceholderType;
    this.bVertical = bVertical;
    this.parentObject = slide || layout || master;

    AscFormat.ExecuteNoHistory(function(){

        if(!bSkipCheckConnector && this.drawingsController && !this.drawingsController.document){
            this.bConnector = AscFormat.isConnectorPreset(presetGeom);
            if(this.bConnector){

                var aSpTree = [];
                this.drawingsController.getAllSingularDrawings(this.drawingsController.getDrawingArray(), aSpTree);
                var oConnector = null;
                for(var i = aSpTree.length - 1; i > -1; --i){
                    oConnector = aSpTree[i].findConnector(startX, startY);
                    if(oConnector){
                        this.startConnectionInfo = oConnector;
                        this.startX = oConnector.x;
                        this.startY = oConnector.y;
                        this.startShape = aSpTree[i];
                        break;
                    }
                }
            }
        }


        var style;

        if(presetGeom.indexOf("WithArrow") > -1)
        {
            presetGeom = presetGeom.substr(0, presetGeom.length - 9);
            this.presetGeom = presetGeom;
            this.arrowsCount = 1;
        }
        if(presetGeom.indexOf("WithTwoArrows") > -1)
        {
            presetGeom = presetGeom.substr(0, presetGeom.length - 13);
            this.presetGeom = presetGeom;
            this.arrowsCount = 2;
        }

        var spDef = theme.spDef;
        let isTextRect = presetGeom && (presetGeom.indexOf("textRect") === 0);
        if(!isTextRect)
        {
            if(spDef && spDef.style)
            {
                style = spDef.style.createDuplicate();
            }
            else
            {
                style = AscFormat.CreateDefaultShapeStyle(this.presetGeom);
            }
        }
        else
        {
            style = AscFormat.CreateDefaultTextRectStyle();
        }
        var brush = theme ? theme.getFillStyle(style.fillRef.idx) : AscFormat.CreateNoFillUniFill();
        style.fillRef.Color.Calculate(theme, slide, layout, master, {R:0, G: 0, B:0, A:255});
        var RGBA = style.fillRef.Color.RGBA;
        if (style.fillRef.Color.color)
        {
            if (brush.fill && (brush.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID))
            {
                brush.fill.color = style.fillRef.Color.createDuplicate();
            }
        }
        var pen = theme.getLnStyle(style.lnRef.idx, style.lnRef.Color);
        style.lnRef.Color.Calculate(theme, slide, layout, master);
        RGBA = style.lnRef.Color.RGBA;

        if(isTextRect)
        {
            var ln, fill;
            ln = new AscFormat.CLn();
            ln.w = 6350;
            ln.Fill = new AscFormat.CUniFill();
            ln.Fill.fill = new AscFormat.CSolidFill();
            ln.Fill.fill.color = new AscFormat.CUniColor();
            ln.Fill.fill.color.color  = new AscFormat.CPrstColor();
            ln.Fill.fill.color.color.id = "black";

            fill = new AscFormat.CUniFill();
            fill.fill = new AscFormat.CSolidFill();
            fill.fill.color = new AscFormat.CUniColor();
            fill.fill.color.color  = new AscFormat.CSchemeColor();
            fill.fill.color.color.id = 12;
            pen.merge(ln);
            brush.merge(fill);
            if(slide)
            {
                brush = AscFormat.CreateNoFillUniFill();
                pen = AscFormat.CreateNoFillLine();
            }
        }

        if(this.arrowsCount > 0)
        {
            pen.setTailEnd(new AscFormat.EndArrow());
            pen.tailEnd.setType(AscFormat.LineEndType.Arrow);
            pen.tailEnd.setLen(AscFormat.LineEndSize.Mid);
            if(this.arrowsCount === 2)
            {
                pen.setHeadEnd(new AscFormat.EndArrow());
                pen.headEnd.setType(AscFormat.LineEndType.Arrow);
                pen.headEnd.setLen(AscFormat.LineEndSize.Mid);
            }
        }
        if(!isTextRect)
        {
            if(spDef && spDef.spPr )
            {
                if(spDef.spPr.Fill )
                {
                    brush.merge(spDef.spPr.Fill);
                }

                if(spDef.spPr.ln )
                {
                    pen.merge(spDef.spPr.ln);
                }
            }
        }

        var geometry = AscFormat.CreateGeometry(!isTextRect ? presetGeom : "rect");

        this.startGeom = geometry;
        if(pen.Fill)
        {
            pen.Fill.calculate(theme, slide, layout, master, RGBA);
        }
        brush.calculate(theme, slide, layout, master, RGBA);

        this.isLine = this.presetGeom === "line";

        this.overlayObject = new AscFormat.OverlayObject(geometry, 5, 5, brush, pen, this.transform);
        this.shape = null;

    }, this, []);

    /**
     *
     * @param e
     * @param x
     * @param y
     * @param {boolean?} isNoMinSize
     */
    this.track = function(e, x, y, isNoMinSize)
    {
        let minShapeSize = MIN_SHAPE_SIZE;
        let minShapeSizeDiv2 = MIN_SHAPE_SIZE_DIV2;

        // ignore minSize
        if (isNoMinSize === true) {
            minShapeSize = 0;
            minShapeSizeDiv2 = 0;
        }

        var bConnectorHandled = false;
        this.oShapeDrawConnectors = null;
        this.lastSpPr = null;

        this.endShape  = null;
        this.endConnectionInfo = null;

        if(this.bConnector){
            var aSpTree = [];
            this.drawingsController.getAllSingularDrawings(this.drawingsController.getDrawingArray(), aSpTree);
            var oConnector = null;
            var oEndConnectionInfo = null;
            for(var i = aSpTree.length - 1; i > -1; --i){
                oConnector = aSpTree[i].findConnector(x, y);
                if(oConnector){
                    oEndConnectionInfo = oConnector;
                    this.oShapeDrawConnectors = aSpTree[i];
                    this.endShape = aSpTree[i];
                    this.endConnectionInfo = oEndConnectionInfo;
                    break;
                }
            }
            if(this.startConnectionInfo || oEndConnectionInfo){

                var _startInfo = this.startConnectionInfo;
                var _endInfo = oEndConnectionInfo;
                if(!_startInfo){
                    _startInfo = AscFormat.fCalculateConnectionInfo(_endInfo, this.startX, this.startY);
                }
                else if(!_endInfo){
                    _endInfo = AscFormat.fCalculateConnectionInfo(_startInfo, x, y);
                }
                var oSpPr = AscFormat.fCalculateSpPr(_startInfo, _endInfo, this.presetGeom, this.overlayObject.pen.w);
                this.flipH = oSpPr.xfrm.flipH === true;
                this.flipV = oSpPr.xfrm.flipV === true;
                this.rot = AscFormat.isRealNumber(oSpPr.xfrm.rot) ? oSpPr.xfrm.rot : 0;
                this.extX = oSpPr.xfrm.extX;
                this.extY = oSpPr.xfrm.extY;
                this.x = oSpPr.xfrm.offX;
                this.y = oSpPr.xfrm.offY;
                this.overlayObject = new AscFormat.OverlayObject(oSpPr.geometry, 5, 5, this.overlayObject.brush, this.overlayObject.pen, this.transform);
                bConnectorHandled = true;
                this.lastSpPr = oSpPr;
            }
            if(!this.oShapeDrawConnectors){
                for(var i = aSpTree.length - 1; i > -1; --i){
                    var oCs = aSpTree[i].findConnectionShape(x, y);
                    if(oCs ){
                        oEndConnectionInfo = oConnector;
                        this.oShapeDrawConnectors = oCs;
                        break;
                    }
                }
            }
            if(false === bConnectorHandled){
                this.overlayObject = new AscFormat.OverlayObject(this.startGeom, 5, 5, this.overlayObject.brush, this.overlayObject.pen, this.transform);
            }
        }

        if(false === bConnectorHandled){
            var real_dist_x = x - this.startX;
            var abs_dist_x = Math.abs(real_dist_x);
            var real_dist_y = y - this.startY;
            var abs_dist_y = Math.abs(real_dist_y);
            this.flipH = false;
            this.flipV = false;
            this.rot = 0;

            if(this.isLine || this.bConnector)
            {
                if(x < this.startX)
                {
                    this.flipH = true;
                }
                if(y < this.startY)
                {
                    this.flipV = true;
                }
            }
            if(!(e.CtrlKey || e.ShiftKey) || (e.CtrlKey && !e.ShiftKey && this.isLine))
            {
                this.extX = abs_dist_x >= minShapeSize  ? abs_dist_x : (this.isLine ? 0 : minShapeSize);
                this.extY = abs_dist_y >= minShapeSize  ? abs_dist_y : (this.isLine ? 0 : minShapeSize);
                if(real_dist_x >= 0)
                {
                    this.x = this.startX;
                }
                else
                {
                    this.x = abs_dist_x >= minShapeSize  ? x : this.startX - this.extX;
                }

                if(real_dist_y >= 0)
                {
                    this.y = this.startY;
                }
                else
                {
                    this.y = abs_dist_y >= minShapeSize  ? y : this.startY - this.extY;
                }
            }
            else if(e.CtrlKey && !e.ShiftKey)
            {
                if(abs_dist_x >= minShapeSizeDiv2 )
                {
                    this.x = this.startX - abs_dist_x;
                    this.extX = 2*abs_dist_x;
                }
                else
                {
                    this.x = this.startX - minShapeSizeDiv2;
                    this.extX = minShapeSize;
                }

                if(abs_dist_y >= minShapeSizeDiv2 )
                {
                    this.y = this.startY - abs_dist_y;
                    this.extY = 2*abs_dist_y;
                }
                else
                {
                    this.y = this.startY - minShapeSizeDiv2;
                    this.extY = minShapeSize;
                }
            }
            else if(!e.CtrlKey && e.ShiftKey)
            {

                var new_width, new_height;
                var prop_coefficient = (typeof AscFormat.SHAPE_ASPECTS[this.presetGeom] === "number" ? AscFormat.SHAPE_ASPECTS[this.presetGeom] : 1);
                if(abs_dist_y === 0)
                {
                    new_width = abs_dist_x > minShapeSize ? abs_dist_x : minShapeSize;
                    new_height = abs_dist_x/prop_coefficient;
                }
                else
                {
                    var new_aspect = abs_dist_x/abs_dist_y;
                    if (new_aspect >= prop_coefficient)
                    {
                        new_width = abs_dist_x;
                        new_height = abs_dist_x/prop_coefficient;
                    }
                    else
                    {
                        new_height = abs_dist_y;
                        new_width = abs_dist_y*prop_coefficient;
                    }
                }

                if(new_width < minShapeSize || new_height < minShapeSize)
                {
                    var k_wh = new_width/new_height;
                    if(new_height < minShapeSize && new_width < minShapeSize)
                    {
                        if(new_height < new_width)
                        {
                            new_height = minShapeSize;
                            new_width = new_height*k_wh;
                        }
                        else
                        {
                            new_width = minShapeSize;
                            new_height = new_width/k_wh;
                        }
                    }
                    else if(new_height < minShapeSize)
                    {
                        new_height = minShapeSize;
                        new_width = new_height*k_wh;
                    }
                    else
                    {
                        new_width = minShapeSize;
                        new_height = new_width/k_wh;
                    }
                }
                this.extX = new_width;
                this.extY = new_height;
                if(real_dist_x >= 0)
                    this.x = this.startX;
                else
                    this.x = this.startX - this.extX;

                if(real_dist_y >= 0)
                    this.y = this.startY;
                else
                    this.y = this.startY - this.extY;

                if(this.isLine)
                {
                    var angle = Math.atan2(real_dist_y, real_dist_x);
                    if(angle >=0 && angle <= Math.PI/8
                        || angle <= 0 && angle >= -Math.PI/8
                        || angle >= 7*Math.PI/8 && angle <= Math.PI )
                    {
                        this.extY = 0;
                        this.y = this.startY;
                    }
                    else if(angle >= 3*Math.PI/8 && angle <= 5*Math.PI/8
                        || angle <= -3*Math.PI/8 && angle >= -5*Math.PI/8)
                    {
                        this.extX = 0;
                        this.x = this.startX;
                    }
                }
            }
            else
            {
                var new_width, new_height;
                var prop_coefficient = (typeof AscFormat.SHAPE_ASPECTS[this.presetGeom] === "number" ? AscFormat.SHAPE_ASPECTS[this.presetGeom] : 1);
                if(abs_dist_y === 0)
                {
                    new_width = abs_dist_x > minShapeSizeDiv2 ? abs_dist_x*2 : minShapeSize;
                    new_height = new_width/prop_coefficient;
                }
                else
                {
                    var new_aspect = abs_dist_x/abs_dist_y;
                    if (new_aspect >= prop_coefficient)
                    {
                        new_width = abs_dist_x*2;
                        new_height = new_width/prop_coefficient;
                    }
                    else
                    {
                        new_height = abs_dist_y*2;
                        new_width = new_height*prop_coefficient;
                    }
                }

                if(new_width < minShapeSize || new_height < minShapeSize)
                {
                    var k_wh = new_width/new_height;
                    if(new_height < minShapeSize && new_width < minShapeSize)
                    {
                        if(new_height < new_width)
                        {
                            new_height = minShapeSize;
                            new_width = new_height*k_wh;
                        }
                        else
                        {
                            new_width = minShapeSize;
                            new_height = new_width/k_wh;
                        }
                    }
                    else if(new_height < minShapeSize)
                    {
                        new_height = minShapeSize;
                        new_width = new_height*k_wh;
                    }
                    else
                    {
                        new_width = minShapeSize;
                        new_height = new_width/k_wh;
                    }
                }
                this.extX = new_width;
                this.extY = new_height;
                this.x = this.startX - this.extX*0.5;
                this.y = this.startY - this.extY*0.5;
            }
        }

        if (Asc.editor.isPdfEditor()) {
            let oDoc = Asc.editor.getPDFDoc();
            let nRotAngle = oDoc.Viewer.getPageRotate(this.pageIndex);
            this.rot = -nRotAngle * Math.PI / 180;

            if (nRotAngle === 90 || nRotAngle === 270) {

                this.x = this.x + (this.extX - this.extY) / 2;
                this.y = this.y - (this.extX - this.extY) / 2;

                let tempExtX = this.extX;
                this.extX = this.extY;
                this.extY = tempExtX;
            }
        }

        this.overlayObject.updateExtents(this.extX, this.extY);
        this.transform.Reset();
        var hc = this.extX * 0.5;
        var vc = this.extY * 0.5;
        AscCommon.global_MatrixTransformer.TranslateAppend(this.transform, -hc, -vc);
        if (this.flipH)
            AscCommon.global_MatrixTransformer.ScaleAppend(this.transform, -1, 1);
        if (this.flipV)
            AscCommon.global_MatrixTransformer.ScaleAppend(this.transform, 1, -1);

        AscCommon.global_MatrixTransformer.RotateRadAppend(this.transform, -this.rot);
        AscCommon.global_MatrixTransformer.TranslateAppend(this.transform, this.x + hc, this.y + vc);
    };

    this.draw = function(overlay)
    {
        if(AscFormat.isRealNumber(this.pageIndex) && overlay.SetCurrentPage)
        {
            overlay.SetCurrentPage(this.pageIndex);
        }
        if(this.oShapeDrawConnectors){
            this.oShapeDrawConnectors.drawConnectors(overlay);
        }
        this.overlayObject.draw(overlay);
    };


    this.isPlaceholderTrack = function() {
        return this.placeholderType !== undefined;
    };
    this.createDrawing = function() {
        let oDrawing;
        if(this.bConnector) {
            oDrawing = new AscFormat.CConnectionShape();
        }
        else if(this.isPlaceholderTrack()) {
            oDrawing = AscCommonSlide.CreatePlaceholder(this.placeholderType, this.bVertical);
        }
        else if(this.drawingsController) {
            oDrawing = this.drawingsController.createShape();
        }
        else {
            oDrawing = new AscFormat.CShape();
        }
        return oDrawing;
    };

    this.getShape = function(bFromWord, DrawingDocument, drawingObjects, isClickMouseEvent)
    {
        var _sp_pr, shape = this.createDrawing();
        if(this.bConnector){
            if(drawingObjects)
            {
                shape.setDrawingObjects(drawingObjects);
            }
            if(this.lastSpPr){
                _sp_pr = this.lastSpPr.createDuplicate();
            }
            else{
                _sp_pr = new AscFormat.CSpPr();
                _sp_pr.setXfrm(new AscFormat.CXfrm());
                var xfrm = _sp_pr.xfrm;
                xfrm.setParent(_sp_pr);
                var x, y;
                if(bFromWord)
                {
                    x = 0;
                    y = 0;
                }
                else
                {
                    x = this.x;
                    y = this.y;
                }
                xfrm.setOffX(x);
                xfrm.setOffY(y);
                xfrm.setExtX(this.extX);
                xfrm.setExtY(this.extY);
                xfrm.setFlipH(this.flipH);
                xfrm.setFlipV(this.flipV);
            }

            shape.setSpPr(_sp_pr);
            _sp_pr.setParent(shape);

            var nv_sp_pr = new AscFormat.UniNvPr();
            shape.setNvSpPr(nv_sp_pr);

            var nvUniSpPr = new AscFormat.CNvUniSpPr();
            if(this.startShape && this.startConnectionInfo)
            {
                nvUniSpPr.stCnxIdx = this.startConnectionInfo.idx;
                nvUniSpPr.stCnxId  = this.startShape.Id;
            }
            if(this.endShape && this.endConnectionInfo)
            {
                nvUniSpPr.endCnxIdx = this.endConnectionInfo.idx;
                nvUniSpPr.endCnxId  = this.endShape.Id;
            }
            shape.nvSpPr.setUniSpPr(nvUniSpPr);
        }
        else{
            if(!this.isPlaceholderTrack()) {
                if(drawingObjects)
                {
                    shape.setDrawingObjects(drawingObjects);
                }
                shape.setSpPr(new AscFormat.CSpPr());
                shape.spPr.setParent(shape);
                if(bFromWord)
                {
                    shape.setWordShape(true);
                }
            }
            var x, y;
            if(bFromWord)
            {
                x = 0;
                y = 0;
            }
            else
            {
                x = this.x;
                y = this.y;
            }

            shape.spPr.setXfrm(new AscFormat.CXfrm());
            var xfrm = shape.spPr.xfrm;
            xfrm.setParent(shape.spPr);
            xfrm.setOffX(x);
            xfrm.setOffY(y);
            xfrm.setExtX(this.extX);
            xfrm.setExtY(this.extY);
            xfrm.setFlipH(this.flipH);
            xfrm.setFlipV(this.flipV);
            xfrm.setRot(this.rot);
        }

        shape.setBDeleted(false);

        if(!this.isPlaceholderTrack())
        {
            if(this.presetGeom && this.presetGeom.indexOf("textRect") === 0)
            {
                let isPdf = Asc.editor.isPdfEditor();

                shape.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
                shape.setTxBox(true);
                var fill, ln;
                if((drawingObjects && drawingObjects.cSld) || isPdf)
                {
                    fill = new AscFormat.CUniFill();
                    fill.setFill(new AscFormat.CNoFill());
                    shape.spPr.setFill(fill);
                }
                else
                {
                    if(!bFromWord)
                    {
                        shape.setStyle(AscFormat.CreateDefaultTextRectStyle());
                    }

                    fill = new AscFormat.CUniFill();
                    fill.setFill(new AscFormat.CSolidFill());
                    fill.fill.setColor(new AscFormat.CUniColor());
                    fill.fill.color.setColor(new AscFormat.CSchemeColor());
                    fill.fill.color.color.setId(12);
                    shape.spPr.setFill(fill);

                    ln = new AscFormat.CLn();
                    ln.setW(6350);
                    ln.setFill(new AscFormat.CUniFill());
                    ln.Fill.setFill(new AscFormat.CSolidFill());
                    ln.Fill.fill.setColor(new AscFormat.CUniColor());
                    ln.Fill.fill.color.setColor(new AscFormat.CPrstColor());
                    ln.Fill.fill.color.color.setId("black");
                    shape.spPr.setLn(ln);    
                }
                var body_pr = new AscFormat.CBodyPr();
                body_pr.setDefault();
                if(this.presetGeom === "textRectVertical") {
                    body_pr.setVert(AscFormat.nVertTTvert);
                }
                if(bFromWord)
                {
                    shape.setTextBoxContent(new CDocumentContent(shape, DrawingDocument, 0, 0, 0, 0, false, false, false));
                    shape.setBodyPr(body_pr);
                }
                else
                {
                    shape.setTxBody(new AscFormat.CTextBody());
                    var content = new AscFormat.CDrawingDocContent(shape.txBody, DrawingDocument, 0, 0, 0, 0, false, false, true);
                    shape.txBody.setParent(shape);
                    shape.txBody.setContent(content);
                    var bNeedCheckExtents = false;
                    if(drawingObjects){
                        if((drawingObjects.cSld || Asc.editor.isPdfEditor()) && !this.isPlaceholderTrack()) {
                            body_pr.textFit = new AscFormat.CTextFit();
                            body_pr.textFit.type = AscFormat.text_fit_Auto;
                            if (isClickMouseEvent) {
                                body_pr.wrap = AscFormat.nTWTNone;
                            } else {
                                body_pr.wrap = AscFormat.nTWTSquare;
                            }
                            bNeedCheckExtents = true;
                        }
                        else{
                            body_pr.vertOverflow = AscFormat.nVOTClip;
                        }
                    }
                    shape.txBody.setBodyPr(body_pr);
                    if(bNeedCheckExtents){
                        var dOldX = shape.spPr.xfrm.offX;
                        var dOldY = shape.spPr.xfrm.offY;
                        shape.setParent(drawingObjects);
                        shape.recalculateContent();
                        shape.checkExtentsByDocContent(true, true);
                        shape.spPr.xfrm.setExtX(this.extX);
                        shape.spPr.xfrm.setOffX(dOldX);
                        shape.spPr.xfrm.setOffY(dOldY);
                    }
                }
            }
            else
            {
                if(!shape.spPr.geometry){
                    shape.spPr.setGeometry(AscFormat.CreateGeometry(this.presetGeom));
                }
                shape.setStyle(AscFormat.CreateDefaultShapeStyle(this.presetGeom));
                if(this.arrowsCount > 0)
                {
                    var ln = new AscFormat.CLn();
                    ln.setTailEnd(new AscFormat.EndArrow());
                    ln.tailEnd.setType(AscFormat.LineEndType.Arrow);
                    ln.tailEnd.setLen(AscFormat.LineEndSize.Mid);
                    if(this.arrowsCount === 2)
                    {
                        ln.setHeadEnd(new AscFormat.EndArrow());
                        ln.headEnd.setType(AscFormat.LineEndType.Arrow);
                        ln.headEnd.setLen(AscFormat.LineEndSize.Mid);
                    }
                    shape.spPr.setLn(ln);
                }
                var spDef = this.theme && this.theme.spDef;
                if(spDef)
                {
                    if(spDef.style)
                    {
                        shape.setStyle(spDef.style.createDuplicate());
                    }
                    if(spDef.spPr)
                    {
                        if(spDef.spPr.Fill)
                        {
                            shape.spPr.setFill(spDef.spPr.Fill.createDuplicate());
                        }
                        if(spDef.spPr.ln)
                        {
                            if(shape.spPr.ln)
                            {
                                shape.spPr.ln.merge(spDef.spPr.ln);
                            }
                            else
                            {
                                shape.spPr.setLn(spDef.spPr.ln.createDuplicate());
                            }
                        }
                    }
                }
            }
        }

        if(this.isPlaceholderTrack())
        {
            shape.checkDrawingUniNvPr();
            let oNvPr = shape.getNvProps();
            if(oNvPr) {
                let oPh = oNvPr.ph;
                let nMaxIdx = undefined;
                let aDrawings = this.parentObject.cSld.spTree;
                for(let nIdx = 0; nIdx < aDrawings.length; ++nIdx) {
                    let oSp = aDrawings[nIdx];
                    let nPhIdx = oSp.getPlaceholderIndex();
                    if(nMaxIdx === undefined || nMaxIdx === null) {
                        nMaxIdx = nPhIdx;
                    }
                    else {
                        if(nPhIdx !== null) {
                            nMaxIdx = Math.max(nMaxIdx, nPhIdx);
                        }
                    }
                }
                if(nMaxIdx !== undefined && nMaxIdx !== null) {
                    oPh.setIdx(nMaxIdx + 1);
                }
                else {
                    oPh.setIdx(1);
                }
            }
        }
        shape.x = this.x;
        shape.y = this.y;
        return shape;
    };


    this.getBounds = function()
    {
        var boundsChecker = new  AscFormat.CSlideBoundsChecker();
        this.draw(boundsChecker);
        var tr = this.transform;
        var arr_p_x = [];
        var arr_p_y = [];
        arr_p_x.push(tr.TransformPointX(0,0));
        arr_p_y.push(tr.TransformPointY(0,0));
        arr_p_x.push(tr.TransformPointX(this.extX,0));
        arr_p_y.push(tr.TransformPointY(this.extX,0));
        arr_p_x.push(tr.TransformPointX(this.extX,this.extY));
        arr_p_y.push(tr.TransformPointY(this.extX,this.extY));
        arr_p_x.push(tr.TransformPointX(0,this.extY));
        arr_p_y.push(tr.TransformPointY(0,this.extY));

        arr_p_x.push(boundsChecker.Bounds.min_x);
        arr_p_x.push(boundsChecker.Bounds.max_x);
        arr_p_y.push(boundsChecker.Bounds.min_y);
        arr_p_y.push(boundsChecker.Bounds.max_y);

        boundsChecker.Bounds.min_x = Math.min.apply(Math, arr_p_x);
        boundsChecker.Bounds.max_x = Math.max.apply(Math, arr_p_x);
        boundsChecker.Bounds.min_y = Math.min.apply(Math, arr_p_y);
        boundsChecker.Bounds.max_y = Math.max.apply(Math, arr_p_y);
        boundsChecker.Bounds.posX = this.x;
        boundsChecker.Bounds.posY = this.y;
        boundsChecker.Bounds.extX = this.extX;
        boundsChecker.Bounds.extY = this.extY;
        return boundsChecker.Bounds;
    }
}
    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].SHAPE_ASPECTS = SHAPE_ASPECTS;
    window['AscFormat'].NewShapeTrack = NewShapeTrack;
})(window);
