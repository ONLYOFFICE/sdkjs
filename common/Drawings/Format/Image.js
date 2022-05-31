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

(
/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {
// Import
var CShape = AscFormat.CShape;
var History = AscCommon.History;
var global_MatrixTransformer = AscCommon.global_MatrixTransformer;

var isRealObject = AscCommon.isRealObject;

    AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetSpPr] =  AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetBlipFill] =  AscDFH.CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetParent] =  AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetGroup] =  AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetStyle] =  AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetNvPicPr] =  AscDFH.CChangesDrawingsObject;

    AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetSpPr] =    function(oClass, value){oClass.spPr = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetBlipFill] =  function(oClass, value, FromLoad){
        oClass.blipFill = value;
        if(FromLoad){
            if(typeof AscCommon.CollaborativeEditing !== "undefined")
            {
                if(value && (typeof value.RasterImageId === "string") && value.RasterImageId.length > 0)
                {
                    AscCommon.CollaborativeEditing.Add_NewImage(value.RasterImageId);
                }
            }
        }
    };
    AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetParent] =  function(oClass, value){oClass.parent = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetGroup] =   function(oClass, value){oClass.group = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetStyle] =   function(oClass, value){oClass.style = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetNvPicPr] =  function(oClass, value){oClass.nvPicPr = value;};



    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ImageShapeSetBlipFill] = AscFormat.CBlipFill;



function CImageShape()
{
	AscFormat.CGraphicObjectBase.call(this);
    this.nvPicPr  = null;
    this.blipFill = null;
    this.style    = null;

    this.cropBrush = false;
    this.isCrop = false;
    this.parentCrop = null;

    this.shdwSp = null;
}

	CImageShape.prototype = Object.create(AscFormat.CGraphicObjectBase.prototype);
	CImageShape.prototype.constructor = CImageShape;

CImageShape.prototype.getObjectType = function()
{
    return AscDFH.historyitem_type_ImageShape;
};

CImageShape.prototype.setNvPicPr = function(pr)
{
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetNvPicPr, this.nvPicPr, pr));
    this.nvPicPr = pr;
};

CImageShape.prototype.setSpPr = function(pr)
{
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetSpPr, this.spPr, pr));
    this.spPr = pr;

    if(pr) {
        pr.setParent(this);
    }
};

CImageShape.prototype.setBlipFill = function(pr)
{
    History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ImageShapeSetBlipFill, this.blipFill, pr));
    this.blipFill = pr;
};

CImageShape.prototype.setParent = function(pr)
{
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetParent, this.parent, pr));
    this.parent = pr;
};

CImageShape.prototype.setGroup = function(pr)
{
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetGroup, this.group, pr));
    this.group = pr;
};

CImageShape.prototype.setStyle = function(pr)
{
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetStyle, this.style, pr));
    this.style = pr;
};

CImageShape.prototype.copy = function(oPr)
{
    var copy = new CImageShape();

    if(this.nvPicPr)
    {
        copy.setNvPicPr(this.nvPicPr.createDuplicate());
    }
    if(this.spPr)
    {
        copy.setSpPr(this.spPr.createDuplicate());
        copy.spPr.setParent(copy);
    }
    if(this.blipFill)
    {
        copy.setBlipFill(this.blipFill.createDuplicate());
    }
    if(this.style)
    {
        copy.setStyle(this.style.createDuplicate());
    }
    copy.setBDeleted(this.bDeleted);
    if(this.macro !== null) {
        copy.setMacro(this.macro);
    }
    if(this.textLink !== null) {
        copy.setTextLink(this.textLink);
    }
    if(this.clientData) {
        copy.setClientData(this.clientData.createDuplicate());
    }
    if(this.fLocksText !== null) {
        copy.setFLocksText(this.fLocksText);
    }
    copy.cachedImage = this.getBase64Img();
    copy.cachedPixH = this.cachedPixH;
    copy.cachedPixW = this.cachedPixW;
    copy.setLocks(this.locks);
    return copy;
};

CImageShape.prototype.getImageUrl = function()
{
    if(isRealObject(this.blipFill))
        return this.blipFill.RasterImageId;
    return null;
};

CImageShape.prototype.getSnapArrays = function(snapX, snapY)
{
    var transform = this.getTransformMatrix();
    snapX.push(transform.tx);
    snapX.push(transform.tx + this.extX*0.5);
    snapX.push(transform.tx + this.extX);
    snapY.push(transform.ty);
    snapY.push(transform.ty + this.extY*0.5);
    snapY.push(transform.ty + this.extY);
};

CImageShape.prototype.checkDrawingBaseCoords = CShape.prototype.checkDrawingBaseCoords;

CImageShape.prototype.setDrawingBaseCoords = CShape.prototype.setDrawingBaseCoords;

CImageShape.prototype.isPlaceholder  = function()
{
    return this.nvPicPr != null && this.nvPicPr.nvPr != undefined && this.nvPicPr.nvPr.ph != undefined;
};

CImageShape.prototype.isShape = function()
{
    return false;
};

CImageShape.prototype.isImage = function()
{
    return true;
};

CImageShape.prototype.isChart = function()
{
    return false;
};

CImageShape.prototype.isGroup = function()
{
    return false;
};


CImageShape.prototype.getWatermarkProps = function()
{
    var oProps = new Asc.CAscWatermarkProperties();
    oProps.put_Type(Asc.c_oAscWatermarkType.Image);
    oProps.put_ImageUrl2(this.blipFill.RasterImageId);
    oProps.put_Scale(-1);
    var oApi;
    if(window["Asc"] && window["Asc"]["editor"])
    {
        oApi = window["Asc"]["editor"];
    }
    else
    {
        oApi = editor;
    }
    if(oApi)
    {
        var oImgP = new Asc.asc_CImgProperty();
        oImgP.ImageUrl = this.blipFill.RasterImageId;
        var oSize = oImgP.asc_getOriginSize(oApi);

        if(oSize)
        {
            var dScale = (((this.extX /oSize.Width) * 100 + 0.5) >> 0) / 100 ;
            oProps.put_Scale(dScale);
            var dAspect = this.extX / this.extY;
            var dAspect2 = oSize.Width / oSize.Height;
            if(AscFormat.fApproxEqual(dAspect, dAspect2, 0.01))
            {
                var oParaDrawing = AscFormat.getParaDrawing(this);
                if(oParaDrawing) {
                    var oParentParagraph = oParaDrawing.Get_ParentParagraph();
                    if (oParentParagraph) {
                        var oSectPr = oParentParagraph.Get_SectPr();
                        if(oSectPr)
                        {
                            var Width = oSectPr.GetContentFrameWidth();
                            if(AscFormat.fApproxEqual(this.extX, Width, 1))
                            {
                                oProps.put_Scale(-1);
                            }
                        }
                    }
                }
            }
        }
    }
    return oProps;
};

CImageShape.prototype.getParentObjects = CShape.prototype.getParentObjects;

CImageShape.prototype.hitInPath = CShape.prototype.hitInPath;

CImageShape.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;

CImageShape.prototype.getRotateAngle = CShape.prototype.getRotateAngle;

CImageShape.prototype.changeSize = CShape.prototype.changeSize;

CImageShape.prototype.getRectBounds = function()
{
    var transform = this.getTransformMatrix();
    var w = this.extX;
    var h = this.extY;
    var rect_points = [{x:0, y:0}, {x: w, y: 0}, {x: w, y: h}, {x: 0, y: h}];
    var min_x, max_x, min_y, max_y;
    min_x = transform.TransformPointX(rect_points[0].x, rect_points[0].y);
    min_y = transform.TransformPointY(rect_points[0].x, rect_points[0].y);
    max_x = min_x;
    max_y = min_y;
    var cur_x, cur_y;
    for(var i = 1; i < 4; ++i)
    {
        cur_x = transform.TransformPointX(rect_points[i].x, rect_points[i].y);
        cur_y = transform.TransformPointY(rect_points[i].x, rect_points[i].y);
        if(cur_x < min_x)
            min_x = cur_x;
        if(cur_x > max_x)
            max_x = cur_x;

        if(cur_y < min_y)
            min_y = cur_y;
        if(cur_y > max_y)
            max_y = cur_y;
    }
    return {minX: min_x, maxX: max_x, minY: min_y, maxY: max_y};
};

CImageShape.prototype.canRotate = function()
{
    if(this.isCrop)
    {
        return false;
    }
    if(this.cropObject)
    {
        return false;
    }
    return AscFormat.CGraphicObjectBase.prototype.canRotate.call(this);
};


CImageShape.prototype.createRotateTrack = function()
{
    return new AscFormat.RotateTrackShapeImage(this);
};

CImageShape.prototype.createResizeTrack = function(cardDirection)
{
    return new AscFormat.ResizeTrackShapeImage(this, cardDirection);
};

CImageShape.prototype.createMoveTrack = function()
{
    return new AscFormat.MoveShapeImageTrack(this);
};

CImageShape.prototype.getInvertTransform = function()
{
    if(this.recalcInfo.recalculateTransform)
    {
        this.recalculateTransform();
        this.recalcInfo.recalculateTransform = true;
    }
    return this.invertTransform;
};

CImageShape.prototype.hitInTextRect = function(x, y)
{
    return false;
};

CImageShape.prototype.getBase64Img = CShape.prototype.getBase64Img;

CImageShape.prototype.convertToWord = function(document)
{
    this.setBDeleted(true);
    var oCopy = this.copy(undefined);
    oCopy.setBDeleted(false);
    return oCopy;
};

CImageShape.prototype.convertToPPTX = function(drawingDocument, worksheet)
{
    var ret = this.copy(undefined);
    ret.setWorksheet(worksheet);
    ret.setParent(null);
    ret.setBDeleted(false);
    return ret;
};

CImageShape.prototype.recalculateBrush = CShape.prototype.recalculateBrush;

CImageShape.prototype.recalculatePen = function()
{
    CShape.prototype.recalculatePen.call(this);
    if(this.pen)
    {
        // if(AscFormat.isRealNumber(this.pen.w)){
        //     this.pen.w *= 2;
        // }
       // if(!this.pen.Fill){
       //     this.pen = null;
       // }
    }
};

CImageShape.prototype.getCompiledLine = CShape.prototype.getCompiledLine;

CImageShape.prototype.getCompiledFill = CShape.prototype.getCompiledFill;

CImageShape.prototype.getCompiledTransparent = CShape.prototype.getCompiledTransparent;

CImageShape.prototype.getAllRasterImages = function(images)
{
    this.blipFill && typeof this.blipFill.RasterImageId === "string" && this.blipFill.RasterImageId.length > 0 && images.push(this.blipFill.RasterImageId);
};

CImageShape.prototype.getHierarchy = function()
{
    if(this.recalcInfo.recalculateShapeHierarchy)
    {
        this.compiledHierarchy.length = 0;
        var hierarchy = this.compiledHierarchy;
        if(this.isPlaceholder())
        {
            var ph_type = this.getPlaceholderType();
            var ph_index = this.getPlaceholderIndex();
            var b_is_single_body = this.getIsSingleBody();
            switch (this.parent.kind)
            {
                case AscFormat.TYPE_KIND.SLIDE:
                {
                    hierarchy.push(this.parent.Layout.getMatchingShape(ph_type, ph_index, b_is_single_body));
                    hierarchy.push(this.parent.Layout.Master.getMatchingShape(ph_type, ph_index, b_is_single_body));
                    break;
                }

                case AscFormat.TYPE_KIND.LAYOUT:
                {
                    hierarchy.push(this.parent.Master.getMatchingShape(ph_type, ph_index, b_is_single_body));
                    break;
                }
            }
        }
        this.recalcInfo.recalculateShapeHierarchy = true;
    }
    return this.compiledHierarchy;
};

CImageShape.prototype.recalculateTransform = function()
{
    this.cachedImage = null;
    if(!isRealObject(this.group))
    {
        if(this.spPr.xfrm.isNotNull())
        {
            var xfrm = this.spPr.xfrm;
            this.x = xfrm.offX;
            this.y = xfrm.offY;
            this.extX = xfrm.extX;
            this.extY = xfrm.extY;
            this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
            this.flipH = xfrm.flipH === true;
            this.flipV = xfrm.flipV === true;
        }
        else
        {
            if(this.isPlaceholder())
            {
                var hierarchy = this.getHierarchy();
                for(var i = 0; i < hierarchy.length; ++i)
                {
                    var hierarchy_sp = hierarchy[i];
                    if(isRealObject(hierarchy_sp) && hierarchy_sp.spPr.xfrm.isNotNull())
                    {
                        var xfrm = hierarchy_sp.spPr.xfrm;
                        this.x = xfrm.offX;
                        this.y = xfrm.offY;
                        this.extX = xfrm.extX;
                        this.extY = xfrm.extY;
                        this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
                        this.flipH = xfrm.flipH === true;
                        this.flipV = xfrm.flipV === true;
                        break;
                    }
                }
                if(i === hierarchy.length)
                {
                    this.x = 0;
                    this.y = 0;
                    this.extX = 5;
                    this.extY = 5;
                    this.rot = 0;
                    this.flipH = false;
                    this.flipV = false;
                }
            }
            else
            {
                this.x = 0;
                this.y = 0;
                this.extX = 5;
                this.extY = 5;
                this.rot = 0;
                this.flipH = false;
                this.flipV = false;
            }
        }
    }
    else
    {
        var xfrm;
        if(this.spPr.xfrm.isNotNull())
        {
            xfrm = this.spPr.xfrm;
        }
        else
        {
            if(this.isPlaceholder())
            {
                var hierarchy = this.getHierarchy();
                for(var i = 0; i < hierarchy.length; ++i)
                {
                    var hierarchy_sp = hierarchy[i];
                    if(isRealObject(hierarchy_sp) && hierarchy_sp.spPr.xfrm.isNotNull())
                    {
                        xfrm = hierarchy_sp.spPr.xfrm;
                        break;
                    }
                }
                if(i === hierarchy.length)
                {
                    xfrm = new AscFormat.CXfrm();
                    xfrm.offX = 0;
                    xfrm.offX = 0;
                    xfrm.extX = 5;
                    xfrm.extY = 5;
                }
            }
            else
            {
                xfrm = new AscFormat.CXfrm();
                xfrm.offX = 0;
                xfrm.offY = 0;
                xfrm.extX = 5;
                xfrm.extY = 5;
            }
        }
        var scale_scale_coefficients = this.group.getResultScaleCoefficients();
        this.x = scale_scale_coefficients.cx*(xfrm.offX - this.group.spPr.xfrm.chOffX);
        this.y = scale_scale_coefficients.cy*(xfrm.offY - this.group.spPr.xfrm.chOffY);
        this.extX = scale_scale_coefficients.cx*xfrm.extX;
        this.extY = scale_scale_coefficients.cy*xfrm.extY;
        this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
        this.flipH = xfrm.flipH === true;
        this.flipV = xfrm.flipV === true;
    }
    this.transform.Reset();
    var hc = this.extX*0.5;
    var vc = this.extY*0.5;
    global_MatrixTransformer.TranslateAppend(this.transform, -hc, -vc);
    if(this.flipH)
        global_MatrixTransformer.ScaleAppend(this.transform, -1, 1);
    if(this.flipV)
        global_MatrixTransformer.ScaleAppend(this.transform, 1, -1);
    global_MatrixTransformer.RotateRadAppend(this.transform, -this.rot);
    global_MatrixTransformer.TranslateAppend(this.transform, this.x + hc, this.y + vc);
    if(isRealObject(this.group))
    {
        global_MatrixTransformer.MultiplyAppend(this.transform, this.group.getTransformMatrix());
    }
    this.invertTransform = global_MatrixTransformer.Invert(this.transform);
   //if(this.drawingBase && !this.group)
   //{
   //    this.drawingBase.setGraphicObjectCoords();
   //}
};

CImageShape.prototype.Refresh_RecalcData = function(data)
{
    switch(data.Type)
    {
        case AscDFH.historyitem_ImageShapeSetBlipFill:
        {
            this.recalcBrush();
            this.recalcFill();
            this.addToRecalculate();
            break;
        }
        case AscDFH.historyitem_ShapeSetBDeleted:{
            if(!this.bDeleted){
                this.addToRecalculate();
            }
            break;
        }
    }
};

CImageShape.prototype.recalculateGeometry = function()
{
    this.calcGeometry = null;
    if(isRealObject(this.spPr.geometry)){
        this.calcGeometry = this.spPr.geometry;
    }
    else{
        var hierarchy = this.getHierarchy();
        for(var i = 0; i < hierarchy.length; ++i){
            if(hierarchy[i] && hierarchy[i].spPr && hierarchy[i].spPr.geometry){
                var _g = hierarchy[i].spPr.geometry;
                this.calcGeometry = AscFormat.ExecuteNoHistory(function(){
                    var _r = _g.createDuplicate();
                    _r.setParent(this);
                    return _r;
                }, this, []);
                break;
            }
        }
    }
    if(isRealObject(this.calcGeometry))
    {
        var transform = this.getTransform();
        this.calcGeometry.Recalculate(transform.extX, transform.extY);
    }
};

CImageShape.prototype.getTransformMatrix = function()
{
    if(this.recalcInfo.recalculateTransform)
    {
        this.recalculateTransform();
        this.recalcInfo.recalculateTransform = false;
    }
    return this.transform;
};

CImageShape.prototype.getTransform = function()
{
    if(this.recalcInfo.recalculateTransform)
    {
        this.recalculateTransform();
        this.recalcInfo.recalculateTransform = false;
    }
    return {x: this.x, y: this.y, extX: this.extX, extY: this.extY, rot: this.rot, flipH: this.flipH, flipV: this.flipV};
};

CImageShape.prototype.draw = function(graphics, transform)
{
    if(this.checkNeedRecalculate && this.checkNeedRecalculate()){
        return;
    }
    var oUR = graphics.updatedRect;
    if(oUR && this.bounds) {
        if(!oUR.isIntersectOther(this.bounds)) {
            return;
        }
    }
    if(graphics.animationDrawer) {
        graphics.animationDrawer.drawObject(this, graphics);
        return;
    }


    this.drawShdw &&  this.drawShdw(graphics);
    var oClipRect;
    if(!graphics.IsSlideBoundsCheckerType){
        oClipRect = this.getClipRect();
    }
    if(oClipRect){
        graphics.SaveGrState();
        graphics.AddClipRect(oClipRect.x, oClipRect.y, oClipRect.w, oClipRect.h);
    }
    var _transform = transform ? transform :this.transform;
    graphics.SetIntegerGrid(false);
    graphics.transform3(_transform, false);
    var shape_drawer = new AscCommon.CShapeDrawer();
    if(this.getObjectType() !== AscDFH.historyitem_type_OleObject && (this.pen || this.brush))
    {
        shape_drawer.fromShape2(this, graphics, this.calcGeometry);
        shape_drawer.draw(this.calcGeometry);
        shape_drawer.Clear();
    }
    var oldBrush = this.brush;
    var oldPen = this.pen;

    if(this.getObjectType() === AscDFH.historyitem_type_OleObject){
        var sImageId = this.blipFill && this.blipFill.RasterImageId;
        if(sImageId){
            var oApi = editor || window['Asc']['editor'];
            if(oApi){
                sImageId = AscCommon.getFullImageSrc2(sImageId);
                var _img = oApi.ImageLoader.map_image_index[sImageId];
                if ((_img && _img.Status === AscFonts.ImageLoadStatus.Loading) || (_img && _img.Image) || true === graphics.IsSlideBoundsCheckerType || true == graphics.RENDERER_PDF_FLAG){
					this.brush = CreateBrushFromBlipFill(this.blipFill);
                    this.pen = null;
                }
                else{
                    this.brush = AscFormat.CreateNoFillUniFill();
                }
            }
        }
        else{
            this.brush = CreateBrushFromBlipFill(this.blipFill);
            this.pen = null;
        }
    }
    else{
        this.brush = CreateBrushFromBlipFill(this.blipFill);
        //this.pen = null;
    }

    shape_drawer.fromShape2(this, graphics, this.calcGeometry);
    shape_drawer.draw(this.calcGeometry);
    if(this.cropBrush){
        this.brush = this.cropBrush;
        this.pen = null;
        shape_drawer.Clear();
        shape_drawer.fromShape2(this, graphics, this.calcGeometry);
        shape_drawer.draw(this.calcGeometry);
    }

    this.brush = oldBrush;
    this.pen = oldPen;

    this.drawLocks(_transform, graphics);
    if(oClipRect){
        graphics.RestoreGrState();
    }
    graphics.reset();
    graphics.SetIntegerGrid(true);
    this.drawAnimLabels && this.drawAnimLabels(graphics);
};


    CImageShape.prototype.handleUpdateLn = function()
    {
        this.recalcLine();
        this.recalcPen();
        this.addToRecalculate();
    };

    CImageShape.prototype.changePresetGeom = function (sPreset) {
        if(sPreset === "textRect"){
            return;
        }
        this.spPr.setGeometry( AscFormat.CreateGeometry(sPreset));
    };

    CImageShape.prototype.changeShadow = function (oShadow) {


        this.spPr && this.spPr.changeShadow(oShadow);
    };

    CImageShape.prototype.recalculateLocalTransform = CShape.prototype.recalculateLocalTransform;
CImageShape.prototype.hit = CShape.prototype.hit;

    CImageShape.prototype.changeLine = function (line)
    {
        if(this.recalcInfo.recalculatePen)
        {
            this.recalculatePen();
        }
        var stroke = AscFormat.CorrectUniStroke(line, this.pen);
        if(stroke.Fill)
        {
            stroke.Fill.convertToPPTXMods();
        }
        this.spPr.setLn(stroke);
    };


CImageShape.prototype.drawAdjustments = function(drawingDocument)
{
    if(!this.canChangeAdjustments()) {
        return;
    }
    if (this.calcGeometry) {
        this.calcGeometry.drawAdjustments(drawingDocument, this.transform, false);
    }
};

CImageShape.prototype.hitToAdjustment = CShape.prototype.hitToAdjustment;

CImageShape.prototype.getPlaceholderType = function()
{
    return this.isPlaceholder() ? this.nvPicPr.nvPr.ph.type : null;
};

CImageShape.prototype.getPlaceholderIndex = function()
{
    return this.isPlaceholder() ? this.nvPicPr.nvPr.ph.idx : null;
};

CImageShape.prototype.getPhType = function()
{
    return this.isPlaceholder() ? this.nvPicPr.nvPr.ph.type : null;
};

CImageShape.prototype.getPhIndex = function()
{
    return this.isPlaceholder() ? this.nvPicPr.nvPr.ph.idx : null;
};

CImageShape.prototype.getMediaFileName = function()
{
    if(this.nvPicPr && this.nvPicPr.nvPr && this.nvPicPr.nvPr.unimedia)
    {
        var oUniMedia = this.nvPicPr.nvPr.unimedia;
        if(oUniMedia.type === 7 || oUniMedia.type === 8)
        {
            if(typeof oUniMedia.media === "string" && oUniMedia.media.length > 0)
            {
                var sExt = AscCommon.GetFileExtension(oUniMedia.media);
                if(this.blipFill && typeof this.blipFill.RasterImageId === 'string')
                {
                    var sName = AscCommon.GetFileName(this.blipFill.RasterImageId);
                    var sMediaFile = sName + '.' + sExt;
                    return sMediaFile;
                }
            }
        }
    }
    return null;
};

CImageShape.prototype.setNvSpPr = function(pr)
{
    History.Add(new AscDFH.CChangesDrawingsObject(this,  AscDFH.historyitem_ImageShapeSetNvPicPr, this.nvPicPr, pr));
    this.nvPicPr = pr;
};

CImageShape.prototype.getAllImages = function(images)
{
    if(this.blipFill instanceof  AscFormat.CBlipFill && typeof this.blipFill.RasterImageId === "string")
    {
        images[AscCommon.getFullImageSrc2(this.blipFill.RasterImageId)] = true;
    }
};
CImageShape.prototype.checkTypeCorrect = function()
{
    if(!this.blipFill){
        return false;
    }
    if(!this.spPr){
        return false;
    }
    return true;
};

CImageShape.prototype.Load_LinkData = function(linkData)
{
};

    CImageShape.prototype.getTypeName = function() {
        return AscCommon.translateManager.getValue("Picture");
    };
	
	
	function CreateBrushFromBlipFill(oBlipFill) {
		if(!oBlipFill) {
			return null;
		}
		var oBrush = new AscFormat.CUniFill();
		oBrush.fill = oBlipFill;
		if(Array.isArray(oBlipFill.Effects)) {
			for(var nEffect = 0; nEffect < oBlipFill.Effects.length; ++nEffect) {
				var oEffect = oBlipFill.Effects[nEffect];
				if (oEffect && oEffect instanceof AscFormat.CAlphaModFix && AscFormat.isRealNumber(oEffect.amt)) {
					oBrush.setTransparent(255 * oEffect.amt / 100000);
					break;
				}
			}
		}
		return oBrush;
	}



    CImageShape.prototype.readChildXml = function (name, reader) {
        switch (name) {
            case "blipFill": {
                var uni_fill = new AscFormat.CUniFill();
                uni_fill.fromXml(reader, "blipFill");
                this.setBlipFill(uni_fill.fill);
                break;
            }
            case "spPr": {
                var spPr = new AscFormat.CSpPr();
                spPr.setParent(this);
                spPr.fromXml(reader);
                this.setSpPr(spPr);
                break;
            }
            case "nvPicPr": {
                let prop = new AscFormat.UniNvPr();
                prop.fromXml(reader);
                this.setNvPicPr(prop);
                this.setLocks(prop.getLocks());
                break;
            }
            case "style": {
                let prop = new AscFormat.CShapeStyle();
                prop.fromXml(reader);
                this.setStyle(prop);
                break;
            }
        }
    };
    CImageShape.prototype.writeChildren = function (writer) {
        //TODO:Implement in children
    };
    

	CImageShape.prototype.toXml = function(writer, name) {
        let namespace_ = "a";
        let bOle = false;
        let oContext = writer.context;
        
        if		(oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_XLSX)			namespace_ = "xdr";
        else if (oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_DOCX ||
                 oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_DOCX_GLOSSARY)	namespace_ = "pic";
        else if (oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_GRAPHICS)		namespace_ = "a";
        else if (oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_CHART_DRAWING)	namespace_ = "cdr";
        else if (oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_PPTX)		namespace_ = "p";

        if (oContext.docType !== AscFormat.XMLWRITER_DOC_TYPE_XLSX &&
            oContext.docType !== AscFormat.XMLWRITER_DOC_TYPE_DOCX &&
            oContext.docType !== AscFormat.XMLWRITER_DOC_TYPE_DOCX_GLOSSARY)
        {
            if(this.getObjectType() === AscDFH.historyitem_type_OleObject)
            {
                bOle = true;
                writer.WriteXmlString("<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"0\" name=\"\"/><p:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></p:cNvGraphicFramePr><p:nvPr><p:extLst><p:ext uri=\"{D42A27DB-BD31-4B8C-83A1-F6EECF244321}\"><p14:modId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"2157879785\"/></p:ext></p:extLst></p:nvPr></p:nvGraphicFramePr>");
                if(this.spPr.xfrm)
                {
                    let oldNamespace = this.spPr.xfrm.m_ns;
                    this.spPr.xfrm.m_ns = "p";
                    this.spPr.xfrm.toXml(writer);
                    this.spPr.xfrm.m_ns = oldNamespace;
                }
                writer.WriteXmlString("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\">");

                writer.WriteXmlNodeStart("p:oleObj");
                writer.WriteXmlAttributeString("name", "oleObj");
                // if(oleObject->m_oId.IsInit())
                // {
                //     writer.WriteAttribute2("r:id", oleObject->m_oId->get());
                // }

                let nCoeffPixToEmu = 635 * 20 * 3 / 4;
                if(this.m_nPixWidth !== null)
                {
                    writer.WriteXmlAttributeInt("imgW", (nCoeffPixToEmu * this.m_nPixWidth + 0.5) >> 0); //twips to emu
                }
                if(this.m_nPixHeight !== null)
                {
                    writer.WriteXmlAttributeInt("imgH", (nCoeffPixToEmu * this.m_nPixHeight + 0.5) >> 0); //twips to emu
                }
                writer.WriteXmlNullableAttributeString("progId", this.m_sApplicationId);
                writer.WriteXmlAttributesEnd();

                writer.WriteString("<p:embed/>");
            }
        }
        writer.WriteXmlNodeStart(namespace_ + ":pic");

        if (oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_DOCX ||
            oContext.docType === AscFormat.XMLWRITER_DOC_TYPE_DOCX_GLOSSARY)
        {
            writer.WriteXmlAttributeString("xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        }
        writer.WriteXmlAttributesEnd();

        if(this.nvPicPr) {
            if(this.nvPicPr.nvUniSpPr) {
                this.nvPicPr.nvUniSpPr.locks = this.locks;
            }
            this.nvPicPr.toXmlPic(writer);
            if(this.nvPicPr.nvUniSpPr) {
                this.nvPicPr.nvUniSpPr.locks = null;
            }
        }

        this.blipFill.toXml(writer, namespace_);

        writer.context.flag = 1;
        this.spPr.toXml(writer);
        writer.context.flag = 0;
        
        writer.WriteXmlNullable(this.style);

        writer.WriteXmlNodeEnd(namespace_ + ":pic");
        
        if (oContext.docType !== AscFormat.XMLWRITER_DOC_TYPE_XLSX &&
            oContext.docType !== AscFormat.XMLWRITER_DOC_TYPE_DOCX &&
            oContext.docType !== AscFormat.XMLWRITER_DOC_TYPE_DOCX_GLOSSARY)
        {
            if(bOle)
            {
                writer.WriteString("</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>");
            }
        }

		// var context = writer.context;
		// var objectId = context.objectId++;
		// writer.WriteXmlNodeStart(name);
		// writer.WriteXmlAttributesEnd();
        //
		// var ns = StaxParser.prototype.GetNSFromNodeName(name);
        //
		// writer.WriteXmlString('<'+ns+'nvPicPr>');
		// writer.WriteXmlString('<'+ns+'cNvPr id="' + objectId + '" name="Picture ' + objectId + '"/>');
		// writer.WriteXmlString('<'+ns+'cNvPicPr><a:picLocks noChangeAspect="1"/></'+ns+'cNvPicPr></'+ns+'nvPicPr>');
		// writer.WriteXmlNullable(this.blipFill, ns + "blipFill");
		// writer.WriteXmlNullable(this.spPr, ns + "spPr");
        //
		// writer.WriteXmlNodeEnd(name);
	};

    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].CImageShape = CImageShape;
})(window);
