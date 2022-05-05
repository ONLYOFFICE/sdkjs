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

/**
 * @param {Window} window
 * @param {undefined} undefined
 */
(function (window, undefined) {


    var CBaseObject = AscFormat.CBaseObject;
    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetLocks] = AscDFH.CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetDrawingBaseType] = AscDFH.CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetDrawingBaseEditAs] = AscDFH.CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetWorksheet] = AscDFH.CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_ShapeSetBDeleted] = AscDFH.CChangesDrawingsBool;

    AscDFH.changesFactory[AscDFH.historyitem_ShapeSetMacro]   = AscDFH.CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_ShapeSetTextLink]   = AscDFH.CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_ShapeSetModelId]   = AscDFH.CChangesDrawingsString;

    AscDFH.changesFactory[AscDFH.historyitem_ShapeSetFLocksText]   = AscDFH.CChangesDrawingsBool;



    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetDrawingBasePos] = AscDFH.CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetDrawingBaseExt] = AscDFH.CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_SetDrawingBaseCoors] = AscDFH.CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_ShapeSetClientData] = AscDFH.CChangesDrawingsObjectNoId;


    var drawingsChangesMap = window['AscDFH'].drawingsChangesMap;

    drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetLocks] = function(oClass, value){oClass.locks = value;};
    drawingsChangesMap[AscDFH.historyitem_ShapeSetBDeleted] = function(oClass, value){oClass.bDeleted = value;};
    drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetDrawingBaseType] = function(oClass, value){
        if(oClass.drawingBase){
            oClass.drawingBase.Type = value;
            oClass.handleUpdateExtents();
        }
    };
    drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetDrawingBaseEditAs] = function(oClass, value){
        if(oClass.drawingBase){
            oClass.drawingBase.editAs = value;
            oClass.handleUpdateExtents();
        }
    };

    drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetWorksheet] = function(oClass, value){
        if(typeof value === "string"){
            var oApi = window['Asc'] && window['Asc']['editor'];
            if(oApi && oApi.wbModel){
                oClass.worksheet = oApi.wbModel.getWorksheetById(value);
            }
            else{
                oClass.worksheet = null;
            }
        }
        else{
            oClass.worksheet = null;
        }
    };


    drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetDrawingBasePos] = function(oClass, value){
        if(value){
            if(oClass.drawingBase && oClass.drawingBase.Pos){
                oClass.drawingBase.Pos.X = value.a;
                oClass.drawingBase.Pos.Y = value.b;
                oClass.handleUpdatePosition();
            }
        }
    };

    drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetDrawingBaseExt] = function(oClass, value){
        if(value){
            if(oClass.drawingBase && oClass.drawingBase.ext){
                oClass.drawingBase.ext.cx = value.a;
                oClass.drawingBase.ext.cy = value.b;
                oClass.handleUpdateExtents();
            }
        }
    };
    drawingsChangesMap[AscDFH.historyitem_AutoShapes_SetDrawingBaseCoors] = function(oClass, value){
        if(value){
            if(oClass.drawingBase){
                oClass.drawingBase.from.col = value.fromCol;
                oClass.drawingBase.from.colOff = value.fromColOff;
                oClass.drawingBase.from.row = value.fromRow;
                oClass.drawingBase.from.rowOff = value.fromRowOff;
                oClass.drawingBase.to.col = value.toCol;
                oClass.drawingBase.to.colOff = value.toColOff;
                oClass.drawingBase.to.row = value.toRow;
                oClass.drawingBase.to.rowOff = value.toRowOff;
                oClass.drawingBase.Pos.X = value.posX;
                oClass.drawingBase.Pos.Y = value.posY;
                oClass.drawingBase.ext.cx = value.cx ;
                oClass.drawingBase.ext.cy = value.cy ;
                oClass.handleUpdateExtents();
            }
        }
    };
    drawingsChangesMap[AscDFH.historyitem_ShapeSetClientData] = function(oClass, value){
        oClass.clientData = value;
    };


    drawingsChangesMap[AscDFH.historyitem_ShapeSetMacro]            = function(oClass, value){
        oClass.macro = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ShapeSetTextLink]         = function(oClass, value){
        oClass.textLink = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ShapeSetModelId]         = function(oClass, value){
        oClass.modelId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ShapeSetFLocksText]         = function(oClass, value){
        oClass.fLocksText = value;
    };

    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_AutoShapes_SetDrawingBasePos] = CDrawingBaseCoordsWritable;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_AutoShapes_SetDrawingBaseExt] = CDrawingBaseCoordsWritable;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_AutoShapes_SetDrawingBaseCoors] = CDrawingBasePosWritable;
    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ShapeSetClientData] = CClientData;

    var LOCKS_MASKS =
    {
        noGrp: 1,
        noUngrp: 4,
        noSelect: 16,
        noRot: 64,
        noChangeAspect: 256,
        noMove: 1024,
        noResize: 4096,
        noEditPoints: 16384,
        noAdjustHandles: 65536,
        noChangeArrowheads: 262144,
        noChangeShapeType: 1048576,
        noDrilldown: 4194304,
        noTextEdit: 8388608,
        noCrop: 16777216,
        txBox: 33554432
    };

    function checkNormalRotate(rot)
    {
        var _rot = normalizeRotate(rot);
        return (_rot >= 0 && _rot < Math.PI * 0.25) || (_rot >= 3 * Math.PI * 0.25 && _rot < 5 * Math.PI * 0.25) || (_rot >= 7 * Math.PI * 0.25 && _rot < 2 * Math.PI);
    }

    function normalizeRotate(rot)
    {
        var new_rot = rot;
        if(AscFormat.isRealNumber(new_rot))
        {
            while(new_rot >= 2*Math.PI)
                new_rot -= 2*Math.PI;
            while(new_rot < 0)
                new_rot += 2*Math.PI;
            if(AscFormat.fApproxEqual(new_rot, 2*Math.PI, 0.001))
            {
                new_rot = 0.0;
            }
            return new_rot;
        }
        return new_rot;
    }


    function CDrawingBaseCoordsWritable(a, b){
        this.a = a;
        this.b = b;
    }

    CDrawingBaseCoordsWritable.prototype.Write_ToBinary = function(Writer){
        Writer.WriteDouble(this.a);
        Writer.WriteDouble(this.b);
    };

    CDrawingBaseCoordsWritable.prototype.Read_FromBinary = function(Reader){
        this.a = Reader.GetDouble();
        this.b = Reader.GetDouble();
    };

    window['AscFormat'].CDrawingBaseCoordsWritable = CDrawingBaseCoordsWritable;

    function CDrawingBasePosWritable(oObject){

            this.fromCol       = null;
            this.fromColOff    = null;
            this.fromRow       = null;
            this.fromRowOff    = null;
            this.toCol         = null;
            this.toColOff      = null;
            this.toRow         = null;
            this.toRowOff      = null;
            this.posX          = null;
            this.posY          = null;
            this.cx            = null;
            this.cy            = null;
            if(oObject){
                this.fromCol       = oObject.fromCol   ;
                this.fromColOff    = oObject.fromColOff;
                this.fromRow       = oObject.fromRow   ;
                this.fromRowOff    = oObject.fromRowOff;
                this.toCol         = oObject.toCol     ;
                this.toColOff      = oObject.toColOff  ;
                this.toRow         = oObject.toRow     ;
                this.toRowOff      = oObject.toRowOff  ;
                this.posX          = oObject.posX      ;
                this.posY          = oObject.posY      ;
                this.cx            = oObject.cx        ;
                this.cy            = oObject.cy        ;
            }
    }

    CDrawingBasePosWritable.prototype.Write_ToBinary = function(Writer){
        AscFormat.writeLong(Writer, this.fromCol);
        AscFormat.writeDouble(Writer, this.fromColOff);
        AscFormat.writeLong(Writer, this.fromRow);
        AscFormat.writeDouble(Writer, this.fromRowOff);
        AscFormat.writeLong(Writer, this.toCol);
        AscFormat.writeDouble(Writer, this.toColOff);
        AscFormat.writeLong(Writer, this.toRow);
        AscFormat.writeDouble(Writer, this.toRowOff);
        AscFormat.writeDouble(Writer, this.posX);
        AscFormat.writeDouble(Writer, this.posY);
        AscFormat.writeDouble(Writer, this.cx);
        AscFormat.writeDouble(Writer, this.cy);
    };
    CDrawingBasePosWritable.prototype.Read_FromBinary = function(Reader){
        this.fromCol      = AscFormat.readLong(Reader);
        this.fromColOff   = AscFormat.readDouble(Reader);
        this.fromRow      = AscFormat.readLong(Reader);
        this.fromRowOff   = AscFormat.readDouble(Reader);
        this.toCol        = AscFormat.readLong(Reader);
        this.toColOff     = AscFormat.readDouble(Reader);
        this.toRow        = AscFormat.readLong(Reader);
        this.toRowOff     = AscFormat.readDouble(Reader);
        this.posX         = AscFormat.readDouble(Reader);
        this.posY         = AscFormat.readDouble(Reader);
        this.cx           = AscFormat.readDouble(Reader);
        this.cy           = AscFormat.readDouble(Reader);
    };

    function CClientData(fLocksWithSheet, fPrintsWithSheet) {
        this.fLocksWithSheet = fLocksWithSheet !== undefined ? fLocksWithSheet : null;
        this.fPrintsWithSheet = fPrintsWithSheet !== undefined ? fPrintsWithSheet : null;
    }
    CClientData.prototype.Write_ToBinary = function(Writer) {
        AscFormat.writeBool(Writer, this.fLocksWithSheet);
        AscFormat.writeBool(Writer, this.fPrintsWithSheet);
    };
    CClientData.prototype.Read_FromBinary = function(Reader) {
        this.fLocksWithSheet  = AscFormat.readBool(Reader);
        this.fPrintsWithSheet = AscFormat.readBool(Reader);
    };
    CClientData.prototype.createDuplicate = function() {
        return new CClientData(this.fLocksWithSheet, this.fPrintsWithSheet);
    };

    /**
     * Class represent bounds graphical object
     * @param {number} l
     * @param {number} t
     * @param {number} r
     * @param {number} b
     * @constructor
     */
    function CGraphicBounds(l, t, r, b){
        this.l = l;
        this.t = t;
        this.r = r;
        this.b = b;
        this.checkWH();
    }

    CGraphicBounds.prototype.fromOther = function(oBounds){
        this.l = oBounds.l;
        this.t = oBounds.t;
        this.r = oBounds.r;
        this.b = oBounds.b;
        this.checkWH();
    };
    CGraphicBounds.prototype.copy = function(){
        return new CGraphicBounds(this.l, this.t, this.r, this.b);
    };
    CGraphicBounds.prototype.transform = function(oTransform){

        var xlt = oTransform.TransformPointX(this.l, this.t);
        var ylt = oTransform.TransformPointY(this.l, this.t);

        var xrt = oTransform.TransformPointX(this.r, this.t);
        var yrt = oTransform.TransformPointY(this.r, this.t);
        var xlb = oTransform.TransformPointX(this.l, this.b);
        var ylb = oTransform.TransformPointY(this.l, this.b);

        var xrb = oTransform.TransformPointX(this.r, this.b);
        var yrb = oTransform.TransformPointY(this.r, this.b);

        this.l = Math.min(xlb, xlt, xrb, xrt);
        this.t = Math.min(ylb, ylt, yrb, yrt);

        this.r = Math.max(xlb, xlt, xrb, xrt);
        this.b = Math.max(ylb, ylt, yrb, yrt);

        this.checkWH();
    };

    CGraphicBounds.prototype.checkByOther = function(oBounds){
        if(oBounds){
            if(oBounds.l < this.l){
                this.l = oBounds.l;
            }
            if(oBounds.t < this.t){
                this.t = oBounds.t;
            }
            if(oBounds.r > this.r){
                this.r = oBounds.r;
            }
            if(oBounds.b > this.b){
                this.b = oBounds.b;
            }
            this.checkWH();
        }
    };
    CGraphicBounds.prototype.checkWH = function(){

        this.x = this.l;
        this.y = this.t;
        this.w = this.r - this.l;
        this.h = this.b - this.t;
    };
    CGraphicBounds.prototype.reset = function(l, t, r, b){

        this.l = l;
        this.t = t;
        this.r = r;
        this.b = b;
        this.checkWH();
    };


    CGraphicBounds.prototype.isIntersect = function(l, t, r, b){

       if(l > this.r){
           return false;
       }
       if(r < this.l){
           return false;
       }
       if(t > this.b){
           return false;
       }
       if(b < this.t){
           return false;
       }
       return true;
    };

    CGraphicBounds.prototype.isIntersectOther = function(o){
        return this.isIntersect(o.l, o.t, o.r, o.b)
    };

    CGraphicBounds.prototype.intersection = function(o){
        var oRes = null;
        var l = Math.max(this.l, o.l);
        var t = Math.max(this.t, o.t);
        var r = Math.min(this.r, o.r);
        var b = Math.min(this.b, o.b);
        if(l <= r && t <= b) {
            return new CGraphicBounds(l, t, r, b);
        }
        return oRes;
    };
    CGraphicBounds.prototype.hit = function(x, y) {
        return x >= this.l && x <= this.r && y >= this.t && y <= this.b;  
    };

    function CCopyObjectProperties()
    {
        this.drawingDocument = null;
        this.idMap = null;
        this.bSaveSourceFormatting = null;
        this.contentCopyPr = null;
    }


    /**
     * Base class for all graphic objects
     * @constructor
     */
    function CGraphicObjectBase() {

        CBaseObject.call(this);
        /*Format fields*/
        this.spPr  = null;
        this.group = null;
        this.parent = null;
        this.bDeleted = true;
        this.locks = 0;
        this.macro = null;
        this.textLink = null;
        this.modelId = null;
        this.fLocksText = null;
        this.clientData = null;

        /*Calculated fields*/
        this.posX = null;
        this.posY = null;
        this.x    = 0;
        this.y    = 0;
        this.extX = 0;
        this.extY = 0;
        this.rot  = 0;
        this.flipH = false;
        this.flipV = false;
        this.bounds = new CGraphicBounds(0, 0, 0, 0);
        this.localTransform = new AscCommon.CMatrix();
        this.transform = new AscCommon.CMatrix();
        this.invertTransform = null;
        this.pen = null;
        this.brush = null;
        this.snapArrayX = [];
        this.snapArrayY = [];

        this.selected = false;

        this.cropObject = null;
        this.Lock = new AscCommon.CLock();
        this.setRecalculateInfo();
    }
    CGraphicObjectBase.prototype = Object.create(CBaseObject.prototype);
    CGraphicObjectBase.prototype.constructor = CGraphicObjectBase;

    CGraphicObjectBase.prototype.notAllowedWithoutId = function() {
        return true;
    };
    /**
     * Create a scheme color
     * @memberof CGraphicObjectBase
     * @returns {CGraphicBounds}
     */
    CGraphicObjectBase.prototype.checkBoundsRect = function(){
        var aCheckX = [], aCheckY = [];
        this.calculateSnapArrays(aCheckX, aCheckY, this.localTransform);
        return new CGraphicBounds(Math.min.apply(Math, aCheckX), Math.min.apply(Math, aCheckY), Math.max.apply(Math, aCheckX), Math.max.apply(Math, aCheckY));
    };
    /**
     * Set default recalculate info
     * @memberof CGraphicObjectBase
     */
    CGraphicObjectBase.prototype.setRecalculateInfo = function(){};
    /**
     * Get object bounds for defining group size
     * @memberof CGraphicObjectBase
     * @returns {CGraphicBounds}
     */
    CGraphicObjectBase.prototype.getBoundsInGroup = function () {
        var r = this.rot;
        if (!AscFormat.isRealNumber(r) || AscFormat.checkNormalRotate(r)) {
            return new CGraphicBounds(this.x, this.y, this.x + this.extX, this.y + this.extY);
        }
        else {
            var hc = this.extX * 0.5;
            var vc = this.extY * 0.5;
            var xc = this.x + hc;
            var yc = this.y + vc;
            return new CGraphicBounds(xc - vc, yc - hc, xc + vc, yc + hc);
        }
    };
    /**
     * Normalize a size object in group
     * @memberof CGraphicObjectBase
     */
    CGraphicObjectBase.prototype.normalize = function () {
        var new_off_x, new_off_y, new_ext_x, new_ext_y;
        var xfrm = this.spPr.xfrm;
        if (!AscCommon.isRealObject(this.group)) {
            new_off_x = xfrm.offX;
            new_off_y = xfrm.offY;
            new_ext_x = xfrm.extX;
            new_ext_y = xfrm.extY;
        }
        else {
            var scale_scale_coefficients = this.group.getResultScaleCoefficients();
            new_off_x = scale_scale_coefficients.cx * (xfrm.offX - this.group.spPr.xfrm.chOffX);
            new_off_y = scale_scale_coefficients.cy * (xfrm.offY - this.group.spPr.xfrm.chOffY);
            new_ext_x = scale_scale_coefficients.cx * xfrm.extX;
            new_ext_y = scale_scale_coefficients.cy * xfrm.extY;
            var txXfrm = this.txXfrm;
            if (txXfrm) {

                var new_tx_off_x = scale_scale_coefficients.cx * txXfrm.offX;
                var new_tx_off_y = scale_scale_coefficients.cy * txXfrm.offY;
                var new_tx_ext_x = scale_scale_coefficients.cx * txXfrm.extX;
                var new_tx_ext_y = scale_scale_coefficients.cy * txXfrm.extY;

                Math.abs(new_tx_off_x - txXfrm.offX) > AscFormat.MOVE_DELTA && txXfrm.setOffX(new_tx_off_x);
                Math.abs(new_tx_off_y - txXfrm.offY) > AscFormat.MOVE_DELTA && txXfrm.setOffY(new_tx_off_y);
                Math.abs(new_tx_ext_x - txXfrm.extX) > AscFormat.MOVE_DELTA && txXfrm.setExtX(new_tx_ext_x);
                Math.abs(new_tx_ext_y - txXfrm.extY) > AscFormat.MOVE_DELTA && txXfrm.setExtY(new_tx_ext_y);
            }
        }
        Math.abs(new_off_x - xfrm.offX) > AscFormat.MOVE_DELTA && xfrm.setOffX(new_off_x);
        Math.abs(new_off_y - xfrm.offY) > AscFormat.MOVE_DELTA && xfrm.setOffY(new_off_y);
        Math.abs(new_ext_x - xfrm.extX) > AscFormat.MOVE_DELTA && xfrm.setExtX(new_ext_x);
        Math.abs(new_ext_y - xfrm.extY) > AscFormat.MOVE_DELTA && xfrm.setExtY(new_ext_y);
    };

    CGraphicObjectBase.prototype.checkHiddenInAnimation = function() {
        if(AscCommonSlide.Slide && (this.parent instanceof AscCommonSlide.Slide)) {
            var oGrObjects = this.parent.graphicObjects;
            if(oGrObjects) {
                if(oGrObjects.isSlideShow()) {
                    var oAnimPlayer = oGrObjects.getAnimationPlayer();
                    if(oAnimPlayer) {
                        if(oAnimPlayer.isDrawingHidden(this.Get_Id())) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    };
    /**
     * Check point hit to bounds object
     * @memberof CGraphicObjectBase
     */
    CGraphicObjectBase.prototype.checkHitToBounds = function(x, y) {
        if(this.parent && (this.parent.Get_ParentTextTransform  && this.parent.Get_ParentTextTransform())) {
            return true;
        }
        if(this.checkHiddenInAnimation && this.checkHiddenInAnimation()) {
            return false;
        }

        var _x, _y;
        if(AscFormat.isRealNumber(this.posX) && AscFormat.isRealNumber(this.posY)) {
            _x = x - this.posX - this.bounds.x;
            _y = y - this.posY - this.bounds.y;
        }
        else {
            _x = x - this.bounds.x;
            _y = y - this.bounds.y;
        }
        var delta = 3 + (this.pen && AscFormat.isRealNumber(this.pen.w) ? this.pen.w/36000 : 0);
        if(_x >= -delta && _x <= this.bounds.w + delta && _y >= -delta && _y <= this.bounds.h + delta) {
            var oClipRect;
            if(this.getClipRect) {
                oClipRect = this.getClipRect();
            }
            if(oClipRect) {
                if(x < oClipRect.x || x > oClipRect.x + oClipRect.w
                    || y < oClipRect.y || y > oClipRect.y + oClipRect.h) {
                    return false;
                }
            }
            return true;
        }
        return false;
    };
    /**
     * Internal method for calculating snap arrays
     * @param {Array} snapArrayX
     * @param {Array} snapArrayY
     * @param {CMatrix} transform
     * @memberof CGraphicObjectBase
     */
    CGraphicObjectBase.prototype.calculateSnapArrays = function(snapArrayX, snapArrayY, transform)
    {
        var t = transform ? transform : this.transform;
        snapArrayX.push(t.TransformPointX(0, 0));
        snapArrayY.push(t.TransformPointY(0, 0));
        snapArrayX.push(t.TransformPointX(this.extX, 0));
        snapArrayY.push(t.TransformPointY(this.extX, 0));

        snapArrayX.push(t.TransformPointX(this.extX*0.5, this.extY*0.5));
        snapArrayY.push(t.TransformPointY(this.extX*0.5, this.extY*0.5));
        snapArrayX.push(t.TransformPointX(this.extX, this.extY));
        snapArrayY.push(t.TransformPointY(this.extX, this.extY));
        snapArrayX.push(t.TransformPointX(0, this.extY));
        snapArrayY.push(t.TransformPointY(0, this.extY));
    };
    /**
     * Public method for calculating snap arrays
     * @memberof CGraphicObjectBase
     */
    CGraphicObjectBase.prototype.recalculateSnapArrays = function()
    {
        this.snapArrayX.length = 0;
        this.snapArrayY.length = 0;
        this.calculateSnapArrays(this.snapArrayX, this.snapArrayY, null);
    };
    CGraphicObjectBase.prototype.setLocks = function(nLocks){
        AscCommon.History.Add( new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_AutoShapes_SetLocks, this.locks, nLocks));
        this.locks = nLocks;
    };
    CGraphicObjectBase.prototype.readMacro = function(oStream)
    {
        var nLength = oStream.GetULong();//length
        var nType = oStream.GetUChar();//attr type - 0
        this.setMacro(oStream.GetString2());
    };
    CGraphicObjectBase.prototype.writeMacro = function(oWriter)
    {
        if(typeof this.macro === "string" && this.macro.length > 0)
        {
            oWriter.StartRecord(0xA1);
            oWriter._WriteString1(0, this.macro);
            oWriter.EndRecord();
        }
    };
    CGraphicObjectBase.prototype.setMacro = function(sMacroName)
    {
        History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ShapeSetMacro, this.macro, sMacroName));
        this.macro = sMacroName;
    };
    CGraphicObjectBase.prototype.setModelId = function(sModelId)
    {
        History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ShapeSetModelId, this.modelId, sModelId));
        this.modelId = sModelId;
    };
    CGraphicObjectBase.prototype.setFLocksText = function(bLock)
    {
        History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ShapeSetFLocksText, this.fLocksText, bLock));
        this.fLocksText = bLock;
    };
    CGraphicObjectBase.prototype.setClientData = function(oClientData)
    {
        History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ShapeSetClientData, this.clientData, oClientData));
        this.clientData = oClientData;
    };
    CGraphicObjectBase.prototype.checkClientData = function()
    {
        if(!this.clientData)
        {
            this.setClientData(new CClientData());
        }
    };
    CGraphicObjectBase.prototype.getProtectionLockText = function()
    {
        return null;
    };
    CGraphicObjectBase.prototype.getProtectionLocked = function()
    {
        if(this.group) {
            return this.group.getProtectionLocked();
        }
        if(!this.clientData) {
            return true;
        }
        return this.clientData.fLocksWithSheet !== false;
    };
    CGraphicObjectBase.prototype.getProtectionPrint = function()
    {
        if(this.group) {
            return this.group.getProtectionPrint();
        }
        if(!this.clientData) {
            return false;
        }
        return this.clientData.fPrintsWithSheet !== false;
    };

    CGraphicObjectBase.prototype.setProtectionLockText = function(bVal)
    {
        if(this.getObjectType() === AscDFH.historyitem_type_Shape)
        {
            if(bVal === true || bVal === false)
            {
                var bValToSet = bVal ? null : false;
                if(this.fLocksText !== bValToSet)
                {
                    this.setFLocksText(bValToSet);
                }
            }
        }
    };
    CGraphicObjectBase.prototype.setProtectionLocked = function(bVal)
    {
        if(bVal === true || bVal === false)
        {
            if(this.group)
            {
                this.group.setProtectionLocked(bVal);
                return
            }
            this.checkClientData();
            var bValToSet = bVal ? undefined : false;
            if(this.clientData.fLocksWithSheet !== bValToSet)
            {
                var oData = this.clientData.createDuplicate();
                oData.fLocksWithSheet = bValToSet;
                this.setClientData(oData);
            }
        }
    };
    CGraphicObjectBase.prototype.setProtectionPrint = function(bVal)
    {
        if(bVal === true || bVal === false)
        {
            if(this.group)
            {
                this.group.setProtectionPrint(bVal);
                return
            }
            this.checkClientData();
            var bValToSet = bVal ? undefined : false;
            if(this.clientData.fPrintsWithSheet !== bValToSet)
            {
                var oData = this.clientData.createDuplicate();
                oData.fPrintsWithSheet = bValToSet;
                this.setClientData(oData);
            }
        }
    };

    CGraphicObjectBase.prototype.assignMacro = function(sGuid)
    {
        if(Array.isArray(this.spTree))
        {
            for(var nSp = 0; nSp < this.spTree.length; ++nSp)
            {
                this.spTree[nSp].assignMacro(sGuid);
            }
            return;
        }
        if(typeof sGuid === "string" && sGuid.length > 0)
        {
            this.setMacro(AscFormat.MACRO_PREFIX + sGuid);
        }
        else
        {
            this.setMacro(null);
        }
    };
    CGraphicObjectBase.prototype.setTextLink = function(sLink)
    {
        History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_ShapeSetTextLink, this.textLink, sLink));
        this.textLink = sLink;
    };
    CGraphicObjectBase.prototype.hasMacro = function()
    {
        var sMacro = this.getMacroOwnOrGroup();
        if(sMacro !== null) {
            return true;
        }
        return false;
    };
    CGraphicObjectBase.prototype.getMacroOwnOrGroup = function()
    {
        if(Array.isArray(this.spTree))
        {
            if(this.spTree.length > 0)
            {
                var oSp = this.spTree[0];
                var sMacro = oSp.getMacroOwnOrGroup();
                if(!sMacro)
                {
                    return null;
                }
                for(nSp = 1; nSp < this.spTree.length; ++nSp)
                {
                    if(sMacro !== this.spTree[nSp].getMacroOwnOrGroup())
                    {
                        return null;
                    }
                }
                return sMacro;
            }
            else
            {
                return null;
            }
        }
        if(typeof this.macro === "string" && this.macro.length > 0)
        {
            return this.macro;
        }
        return null;
    };
    CGraphicObjectBase.prototype.hasJSAMacro = function()
    {
        var sMacro = this.getMacroOwnOrGroup();
        if(typeof sMacro === "string" && sMacro.indexOf(AscFormat.MACRO_PREFIX) === 0) {
            return true;
        }
        return false;
    };
    CGraphicObjectBase.prototype.getJSAMacroId = function()
    {
        var sMacro = this.getMacroOwnOrGroup();
        if(typeof sMacro === "string" && sMacro.indexOf(AscFormat.MACRO_PREFIX) === 0) {
            return sMacro.slice(AscFormat.MACRO_PREFIX.length);
        }
        return null;
    };
    CGraphicObjectBase.prototype.getLockValue = function(nMask) {
        return  !!((this.locks & nMask) && (this.locks & (nMask << 1)));
    };
    CGraphicObjectBase.prototype.setLockValue = function(nMask, bValue) {
        if(!AscFormat.isRealBool(bValue)) {
            this.setLocks((~nMask) & this.locks);
        }
        else{
            this.setLocks(this.locks | nMask | (bValue ? nMask << 1 : 0));
        }
    };
    CGraphicObjectBase.prototype.getNoGrp = function(){
        return this.getLockValue(LOCKS_MASKS.noGrp);
    };
    CGraphicObjectBase.prototype.getNoUngrp = function(){
        return this.getLockValue(LOCKS_MASKS.noUngrp);
    };
    CGraphicObjectBase.prototype.getNoSelect = function(){
        return this.getLockValue(LOCKS_MASKS.noSelect);
    };
    CGraphicObjectBase.prototype.getNoRot = function(){
        return this.getLockValue(LOCKS_MASKS.noRot);
    };
    CGraphicObjectBase.prototype.getNoChangeAspect = function(){
        return this.getLockValue(LOCKS_MASKS.noChangeAspect);
    };
    CGraphicObjectBase.prototype.getNoMove = function(){
        return this.getLockValue(LOCKS_MASKS.noMove);
    };
    CGraphicObjectBase.prototype.getNoResize = function(){
        return this.getLockValue(LOCKS_MASKS.noResize);
    };
    CGraphicObjectBase.prototype.getNoEditPoints = function(){
        return this.getLockValue(LOCKS_MASKS.noEditPoints);
    };
    CGraphicObjectBase.prototype.getNoAdjustHandles = function(){
        return this.getLockValue(LOCKS_MASKS.noAdjustHandles);
    };
    CGraphicObjectBase.prototype.getNoChangeArrowheads = function(){
        return this.getLockValue(LOCKS_MASKS.noChangeArrowheads);
    };
    CGraphicObjectBase.prototype.getNoChangeShapeType = function(){
        return this.getLockValue(LOCKS_MASKS.noChangeShapeType);
    };
    CGraphicObjectBase.prototype.getNoDrilldown = function(){
        return this.getLockValue(LOCKS_MASKS.noDrilldown);
    };
    CGraphicObjectBase.prototype.getNoTextEdit = function(){
        return this.getLockValue(LOCKS_MASKS.noTextEdit);
    };
    CGraphicObjectBase.prototype.getNoCrop = function(){
        return this.getLockValue(LOCKS_MASKS.noCrop);
    };
    CGraphicObjectBase.prototype.getTxBox = function(){
        return this.getLockValue(LOCKS_MASKS.txBox);
    };
    CGraphicObjectBase.prototype.setTxBox = function(bValue){
        return this.setLockValue(LOCKS_MASKS.txBox, bValue);
    };
    CGraphicObjectBase.prototype.setNoChangeAspect = function(bValue){
        return this.setLockValue(LOCKS_MASKS.noChangeAspect, bValue);
    };
    CGraphicObjectBase.prototype.canEditGeometry = function(){
        return this.getObjectType() === AscDFH.historyitem_type_Shape &&
            !this.isPlaceholder() &&
            this.getNoEditPoints() !== true &&
            !!(this.spPr && this.spPr.geometry) && !(this.isObjectInSmartArt()); // todo: functionality not available in microsoft for smartart shapes, but the OOX format supports it, currently blocked due to resizing blocking
    };
    CGraphicObjectBase.prototype.canRotate = function() {
        if(!this.canEdit()) {
            return false;
        }
        return this.getNoRot() === false;
    };
    CGraphicObjectBase.prototype.canResize = function() {
        if(!this.canEdit()) {
            return false;
        }
        return this.getNoResize() === false;
    };
    CGraphicObjectBase.prototype.canMove = function() {
        var oApi = Asc.editor || editor;
        var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
        if(isDrawHandles === false)
        {
            return false;
        }
        if(!this.canEdit()) {
            return false;
        }
        return this.getNoMove() === false;
    };
    CGraphicObjectBase.prototype.canGroup = function() {
        if(!this.canEdit()) {
            return false;
        }
        return this.getNoGrp() === false;
    };
    CGraphicObjectBase.prototype.canUnGroup = function() {
        if(!this.canEdit()) {
            return false;
        }
        return (this.getObjectType() === AscDFH.historyitem_type_GroupShape || this.getObjectType() === AscDFH.historyitem_type_SmartArt && this.drawing)  && this.getNoUngrp() === false;
    };
    CGraphicObjectBase.prototype.canChangeAdjustments = function () {
        if(!this.canEdit()) {
            return false;
        }
        return !this.isObjectInSmartArt() && this.getNoAdjustHandles() === false;
    };
    CGraphicObjectBase.prototype.Reassign_ImageUrls = function(mapUrl){
        var blip_fill;
        if(this.blipFill){
            if(mapUrl[this.blipFill.RasterImageId]){
                if(this.setBlipFill){
                    blip_fill = this.blipFill.createDuplicate();
                    blip_fill.setRasterImageId(mapUrl[this.blipFill.RasterImageId]);
                    this.setBlipFill(blip_fill);
                }
            }
        }
        if(this.spPr && this.spPr.Fill && this.spPr.Fill.fill && this.spPr.Fill.fill.RasterImageId){
            if(mapUrl[this.spPr.Fill.fill.RasterImageId]){
                blip_fill = this.spPr.Fill.fill.createDuplicate();
                blip_fill.setRasterImageId(mapUrl[this.spPr.Fill.fill.RasterImageId]);
                var oUniFill = this.spPr.Fill.createDuplicate();
                oUniFill.setFill(blip_fill);
                this.spPr.setFill(oUniFill);
            }
        }
        if(Array.isArray(this.spTree)){
            for(var i = 0; i < this.spTree.length; ++i){
                if(this.spTree[i].Reassign_ImageUrls){
                    this.spTree[i].Reassign_ImageUrls(mapUrl);
                }
            }
        }
    };
    CGraphicObjectBase.prototype.getAllFonts = function(mapUrl){
    };
    CGraphicObjectBase.prototype.collectEquations3 = function(aEquations) {
        if(Array.isArray(this.spTree)) {
            for(var nSp = 0; nSp < this.spTree.length; ++nSp) {
                this.spTree[nSp].collectEquations3(aEquations);
            }
        }
        if(this.m_oMathObject) {
            aEquations.push(this);
        }
    };
    CGraphicObjectBase.prototype.getOuterShdw = function(){
        if(this.spPr && this.spPr.effectProps && this.spPr.effectProps.EffectLst && this.spPr.effectProps.EffectLst.outerShdw)
        {
            return this.spPr.effectProps.EffectLst.outerShdw;
        }
        return null;
    };
    CGraphicObjectBase.prototype.recalculateShdw = function(){

        this.shdwSp = null;
        var outerShdw = this.getOuterShdw && this.getOuterShdw();
        if(outerShdw)
        {
            AscFormat.ExecuteNoHistory(function(){
                var geometry = this.calcGeometry || this.spPr && this.spPr.geometry;

                var oParentObjects = this.getParentObjects();
                var track_object = new AscFormat.NewShapeTrack("rect", 0, 0, oParentObjects.theme, oParentObjects.master, oParentObjects.layout, oParentObjects.slide, 0);
                track_object.track({}, 0, 0);
                var shape = track_object.getShape(false, null, null);
                if(geometry)
                {
                    shape.spPr.setGeometry(geometry.createDuplicate());
                    shape.spPr.geometry.setParent(shape.spPr);
                }
                var oShadowFill;
                if(outerShdw.color)
                {
                    oShadowFill = AscFormat.CreateUniFillByUniColorCopy(outerShdw.color);
                }
                else
                {
                    oShadowFill = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(0, 0, 0));
                }

                if(this.getObjectType() === AscDFH.historyitem_type_Shape
                && (!this.brush || !this.brush.isVisible()))
                {
                    if(this.pen && this.pen.isVisible())
                    {
                        shape.spPr.Fill = AscFormat.CreateNoFillUniFill();
                        shape.spPr.ln = this.pen.createDuplicate();
                        shape.spPr.ln.Fill = oShadowFill;
                    }
                    else
                    {
                        shape.spPr.Fill = AscFormat.CreateNoFillUniFill();
                        shape.spPr.ln = AscFormat.CreateNoFillLine();
                    }
                }
                else
                {
                    shape.spPr.Fill = oShadowFill;
                    shape.spPr.ln = AscFormat.CreateNoFillLine();
                }
                var W = this.extX;
                var H = this.extY;
                var penW = 0;
                if(this.pen)
                {
                    penW = this.pen.w ? this.pen.w / 36000.0 : 12700.0 / 36000.0;
                    if(this.getObjectType() !== AscDFH.historyitem_type_ImageShape)
                    {
                        penW /= 2.0;
                    }
                }
                if(outerShdw.sx)
                {
                    W *= outerShdw.sx / 100000;
                }
                if(outerShdw.sy)
                {
                    H *= outerShdw.sy / 100000;
                }
                // W += penW;
                // H += penW;
                if(W < this.extX + penW)
                {
                    W = this.extX + penW + 1;
                }
                if(H < this.extY + penW)
                {
                    H = this.extY  + penW + 1;
                }
                shape.spPr.xfrm.setExtX(W);
                shape.spPr.xfrm.setExtY(H);
                shape.spPr.xfrm.setOffX(0);
                shape.spPr.xfrm.setOffY(0);
                if(!(this.parent && this.parent.Extent))
                {
                    shape.setParent(this.parent);
                }
                shape.recalculate();
                this.shdwSp = shape;
            }, this, []);
        }
    };
    CGraphicObjectBase.prototype.check_bounds = function(checker) {
        checker._s();
        checker._m(0, 0);
        checker._l(this.extX, 0);
        checker._l(this.extX, this.extY);
        checker._l(0, this.extY);
        checker._z();
        checker._e();
    };
    CGraphicObjectBase.prototype.handleObject = function(fCallback) {
        fCallback(this);
    };

    CGraphicObjectBase.prototype.isObjectInSmartArt = function () {
        if(this.group && this.group.isSmartArtObject()) {
            return true;
        }
        return false;

    };
    CGraphicObjectBase.prototype.isGroupObject = function () {
        var nType = this.getObjectType();
        return nType === AscDFH.historyitem_type_GroupShape || nType === AscDFH.historyitem_type_LockedCanvas || this.isSmartArtObject();
    };
    CGraphicObjectBase.prototype.isSmartArtObject = function () {
        var nType = this.getObjectType();
        return nType === AscDFH.historyitem_type_SmartArt ||
            nType === AscDFH.historyitem_type_SmartArtDrawing;
    };


    CGraphicObjectBase.prototype.drawShdw = function(graphics){
        var outerShdw = this.getOuterShdw && this.getOuterShdw();
        if(this.shdwSp && outerShdw && !graphics.IsSlideBoundsCheckerType)
        {
            var oTransform =  new AscCommon.CMatrix();
            var dist = outerShdw.dist ? outerShdw.dist /36000 : 0;
            var dir = outerShdw.dir ? outerShdw.dir : 0;
            oTransform.tx = dist*Math.cos(AscFormat.cToRad*dir) - (this.shdwSp.extX - this.extX) / 2.0;
            oTransform.ty = dist*Math.sin(AscFormat.cToRad*dir) - (this.shdwSp.extY - this.extY) / 2.0;
            global_MatrixTransformer.MultiplyAppend(oTransform, this.transform);
			this.shdwSp.bounds.x = this.bounds.x + this.shdwSp.bounds.l;
			this.shdwSp.bounds.y = this.bounds.y + this.shdwSp.bounds.t;
            this.shdwSp.transform = oTransform;
            this.shdwSp.draw(graphics);
        }
    };
    CGraphicObjectBase.prototype.getAllRasterImages = function(mapUrl){
    };
    CGraphicObjectBase.prototype.getAllSlicerViews = function(aSlicerView) {

    };
    CGraphicObjectBase.prototype.checkCorrect = function(){
        if(this.bDeleted === true){
            return false;
        }
        return this.checkTypeCorrect();
    };
    CGraphicObjectBase.prototype.Clear_ContentChanges = function()
    {
    };
    CGraphicObjectBase.prototype.Add_ContentChanges = function(Changes)
    {
    };
    CGraphicObjectBase.prototype.Refresh_ContentChanges = function()
    {
    };
    CGraphicObjectBase.prototype.getWatermarkProps = function()
    {
        var oProps = new Asc.CAscWatermarkProperties();
        oProps.put_Type(Asc.c_oAscWatermarkType.None);
        return oProps;
    };
    CGraphicObjectBase.prototype.CheckCorrect = function(){
        return this.checkCorrect();
    };
    CGraphicObjectBase.prototype.checkTypeCorrect = function(){
        return true;
    };
    CGraphicObjectBase.prototype.handleUpdateExtents = function(bExtX){
    };
    CGraphicObjectBase.prototype.handleUpdatePosition = function(){
    };
    CGraphicObjectBase.prototype.setDrawingObjects = function(drawingObjects){
        this.drawingObjects = drawingObjects;
        if(Array.isArray(this.spTree)) {
            for(var i = 0; i < this.spTree.length; ++i){
                this.spTree[i].setDrawingObjects(drawingObjects);
            }
        }
    };
    CGraphicObjectBase.prototype.setDrawingBase = function(drawingBase){
        this.drawingBase = drawingBase;
        //if(Array.isArray(this.spTree)) {
        //    for(var i = 0; i < this.spTree.length; ++i)
        //    {
        //        this.spTree[i].setDrawingBase(drawingBase);
        //    }
        //}
    };
    CGraphicObjectBase.prototype.setDrawingBaseType = function(nType){
        if(this.drawingBase){
            History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_AutoShapes_SetDrawingBaseType, this.drawingBase.Type, nType));
            this.drawingBase.Type = nType;
            this.handleUpdateExtents();
        }
    };
    CGraphicObjectBase.prototype.setDrawingBaseEditAs = function(nType){
        if(this.drawingBase){
            History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_AutoShapes_SetDrawingBaseEditAs, this.drawingBase.editAs, nType));
            this.drawingBase.editAs = nType;
            this.handleUpdateExtents();
        }
    };
    CGraphicObjectBase.prototype.setDrawingBasePos = function(fPosX, fPosY)
    {
        if(this.drawingBase && this.drawingBase.Pos)
        {
            History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_AutoShapes_SetDrawingBasePos, new CDrawingBaseCoordsWritable(this.drawingBase.Pos.X, this.drawingBase.Pos.Y), new CDrawingBaseCoordsWritable(fPosX, fPosY)));
            this.drawingBase.Pos.X = fPosX;
            this.drawingBase.Pos.Y = fPosY;
            this.handleUpdatePosition();
        }
    };
    CGraphicObjectBase.prototype.setDrawingBaseExt = function(fExtX, fExtY)
    {
        if(this.drawingBase && this.drawingBase.ext)
        {
            History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_AutoShapes_SetDrawingBaseExt, new CDrawingBaseCoordsWritable(this.drawingBase.ext.cx, this.drawingBase.ext.cy), new CDrawingBaseCoordsWritable(fExtX, fExtY)));
            this.drawingBase.ext.cx = fExtX;
            this.drawingBase.ext.cy = fExtY;
            this.handleUpdateExtents();
        }
    };
    CGraphicObjectBase.prototype.setTransformParams = function(x, y, extX, extY, rot, flipH, flipV)
    {
        if(!this.spPr)
        {
            this.setSpPr(new AscFormat.CSpPr());
            this.spPr.setParent(this);
        }
        if(!this.spPr.xfrm)
        {
            this.spPr.setXfrm(new AscFormat.CXfrm());
            this.spPr.xfrm.setParent(this.spPr);
        }
        this.spPr.xfrm.setOffX(x);
        this.spPr.xfrm.setOffY(y);
        this.spPr.xfrm.setExtX(extX);
        this.spPr.xfrm.setExtY(extY);
        if(this.isGroupObject()) {
            this.spPr.xfrm.setChOffX(0);
            this.spPr.xfrm.setChOffY(0);
            this.spPr.xfrm.setChExtX(extX);
            this.spPr.xfrm.setChExtY(extY);
        }
        this.spPr.xfrm.setRot(rot);
        this.spPr.xfrm.setFlipH(flipH);
        this.spPr.xfrm.setFlipV(flipV);
    };
    CGraphicObjectBase.prototype.getPlaceholderType = function()
    {
        return null;
    };
    CGraphicObjectBase.prototype.getPlaceholderIndex = function()
    {
        return null;
    };
    CGraphicObjectBase.prototype.getPhType = function()
    {
        return null;
    };
    CGraphicObjectBase.prototype.getPhIndex = function()
    {
        return null;
    };
    CGraphicObjectBase.prototype.getDrawingBaseType = function()
    {
        if(this.drawingBase)
        {
            if(this.drawingBase.Type === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell)
            {
                if(this.drawingBase.editAs !== null)
                {
                    return this.drawingBase.editAs;
                }
            }
            return this.drawingBase.Type;
        }
        return null;
    };

    CGraphicObjectBase.prototype.recalcBrush = function() {};
    CGraphicObjectBase.prototype.recalcPen = function() {};
    CGraphicObjectBase.prototype.recalcTransform = function() {};
    CGraphicObjectBase.prototype.recalcTransformText = function() {};
    CGraphicObjectBase.prototype.recalcBounds = function() {};
    CGraphicObjectBase.prototype.recalcGeometry = function() {};
    CGraphicObjectBase.prototype.recalcStyle = function() {};
    CGraphicObjectBase.prototype.recalcFill = function() {};
    CGraphicObjectBase.prototype.recalcLine = function() {};
    CGraphicObjectBase.prototype.recalcTransparent = function() {};
    CGraphicObjectBase.prototype.recalcTextStyles = function() {};
    CGraphicObjectBase.prototype.recalcTxBoxContent = function() {};
    CGraphicObjectBase.prototype.recalcWrapPolygon = function() {};
    CGraphicObjectBase.prototype.recalcContent = function() {};
    CGraphicObjectBase.prototype.recalcContent2 = function() {};

    CGraphicObjectBase.prototype.checkDrawingBaseCoords = function()
    {
        if(this.drawingBase && this.spPr && this.spPr.xfrm && !this.group) {
            var oldX = this.x, oldY = this.y, oldExtX = this.extX, oldExtY = this.extY;
            var oldRot = this.rot;
            this.x = this.spPr.xfrm.offX;
            this.y = this.spPr.xfrm.offY;
            this.extX = this.spPr.xfrm.extX;
            this.extY = this.spPr.xfrm.extY;
            this.rot = AscFormat.isRealNumber(this.spPr.xfrm.rot) ? AscFormat.normalizeRotate(this.spPr.xfrm.rot) : 0;

            var oldFromCol = this.drawingBase.from.col,
                oldFromColOff = this.drawingBase.from.colOff,
                oldFromRow = this.drawingBase.from.row,
                oldFromRowOff = this.drawingBase.from.rowOff,
                oldToCol = this.drawingBase.to.col,
                oldToColOff = this.drawingBase.to.colOff,
                oldToRow = this.drawingBase.to.row,
                oldToRowOff = this.drawingBase.to.rowOff,
                oldPosX = this.drawingBase.Pos.X,
                oldPosY = this.drawingBase.Pos.Y,
                oldCx = this.drawingBase.ext.cx,
                oldCy = this.drawingBase.ext.cy;


            this.drawingBase.setGraphicObjectCoords();
            this.x = oldX;
            this.y = oldY;
            this.extX = oldExtX;
            this.extY = oldExtY;
            this.rot = oldRot;
            var from = this.drawingBase.from, to = this.drawingBase.to;
            History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_AutoShapes_SetDrawingBaseCoors,
                new CDrawingBasePosWritable({
                    fromCol: oldFromCol,
                    fromColOff: oldFromColOff,
                    fromRow: oldFromRow,
                    fromRowOff: oldFromRowOff,
                    toCol: oldToCol,
                    toColOff: oldToColOff,
                    toRow: oldToRow,
                    toRowOff: oldToRowOff,
                    posX: oldPosX,
                    posY: oldPosY,
                    cx: oldCx,
                    cy: oldCy
                }),
                new CDrawingBasePosWritable({
                    fromCol: from.col,
                    fromColOff: from.colOff,
                    fromRow: from.row,
                    fromRowOff: from.rowOff,
                    toCol: to.col,
                    toColOff: to.colOff,
                    toRow: to.row,
                    toRowOff: to.rowOff,
                    posX: this.drawingBase.Pos.X,
                    posY: this.drawingBase.Pos.Y,
                    cx: this.drawingBase.ext.cx,
                    cy: this.drawingBase.ext.cy
                })));
            this.handleUpdateExtents();
        }
};
    CGraphicObjectBase.prototype.setDrawingBaseCoords = function(fromCol, fromColOff, fromRow, fromRowOff, toCol, toColOff, toRow, toRowOff, posX, posY, extX, extY)
    {
        if(this.drawingBase)
        {
            History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_AutoShapes_SetDrawingBaseCoors, new CDrawingBasePosWritable({
                    fromCol   : this.drawingBase.from.col,
                    fromColOff: this.drawingBase.from.colOff,
                    fromRow   : this.drawingBase.from.row,
                    fromRowOff: this.drawingBase.from.rowOff,
                    toCol     : this.drawingBase.to.col,
                    toColOff  : this.drawingBase.to.colOff,
                    toRow     : this.drawingBase.to.row,
                    toRowOff  : this.drawingBase.to.rowOff,
                    posX      : this.drawingBase.Pos.X,
                    posY      : this.drawingBase.Pos.Y,
                    cx        : this.drawingBase.ext.cx,
                    cy        : this.drawingBase.ext.cy
                }),
                new CDrawingBasePosWritable({
                    fromCol:    fromCol,
                    fromColOff: fromColOff,
                    fromRow   : fromRow,
                    fromRowOff: fromRowOff,
                    toCol:    toCol,
                    toColOff: toColOff,
                    toRow   : toRow,
                    toRowOff: toRowOff,
                    posX    : posX,
                    posY    : posY,
                    cx      : extX,
                    cy      : extY
                })));


                this.drawingBase.from.col    = fromCol;
                this.drawingBase.from.colOff = fromColOff;
                this.drawingBase.from.row    = fromRow;
                this.drawingBase.from.rowOff = fromRowOff;

                this.drawingBase.to.col    = toCol;
                this.drawingBase.to.colOff = toColOff;
                this.drawingBase.to.row    = toRow;
                this.drawingBase.to.rowOff = toRowOff;

                this.drawingBase.Pos.X  = posX;
                this.drawingBase.Pos.Y  = posY;
                this.drawingBase.ext.cx = extX;
                this.drawingBase.ext.cy = extY;

            this.handleUpdateExtents();
        }
    };
    CGraphicObjectBase.prototype.setWorksheet = function(worksheet)
    {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_AutoShapes_SetWorksheet, this.worksheet ? this.worksheet.getId() : null, worksheet ? worksheet.getId() : null));
        this.worksheet = worksheet;
        if(Array.isArray(this.spTree))
        {
            for(var i = 0; i < this.spTree.length; ++i)
            {
                this.spTree[i].setWorksheet(worksheet);
            }
        }
        if(Array.isArray(this.userShapes))
        {
            for(var nSp = 0; nSp < this.userShapes.length; ++nSp)
            {
                var oAnchor = this.userShapes[nSp];
                if(oAnchor)
                {
                    var oSp = oAnchor.object;
                    if(oSp && oSp.setWorksheet)
                    {
                        oSp.setWorksheet(worksheet);
                    }
                }
            }
        }
    };
    CGraphicObjectBase.prototype.getWorksheet = function()
    {
        return this.worksheet;
    };
    CGraphicObjectBase.prototype.getWorkbook = function()
    {
        var oWorksheet = this.getWorksheet();
        if(!oWorksheet)
        {
            return null;
        }
        if(oWorksheet.workbook)
        {
            return oWorksheet.workbook;
        }
        return null;
    };
    CGraphicObjectBase.prototype.getUniNvProps = function(){
        return this.nvSpPr || this.nvPicPr || this.nvGrpSpPr || this.nvGraphicFramePr || null;
    };
    CGraphicObjectBase.prototype.getCNvProps = function(){
        var oUniNvPr = this.getUniNvProps();
        if(oUniNvPr){
            return oUniNvPr.cNvPr;
        }
        return null;
    };
    CGraphicObjectBase.prototype.getFormatId = function(){
        var oCNvPr = this.getCNvProps();
        if(oCNvPr) {
            return oCNvPr.id;
        }
        return null;
    };
    CGraphicObjectBase.prototype.getNvProps = function(){
        var oUniNvPr = this.getUniNvProps();
        if(oUniNvPr){
            return oUniNvPr.nvPr;
        }
        return null;
    };
    CGraphicObjectBase.prototype.setTitle = function(sTitle){
        if(undefined === sTitle || null === sTitle){
            return;
        }
        var oNvPr = this.getCNvProps();
        if(oNvPr){
            oNvPr.setTitle(sTitle ? sTitle : null);
        }
    };
    CGraphicObjectBase.prototype.setDescription = function(sDescription){
        if(undefined === sDescription || null === sDescription){
            return;
        }
        var oNvPr = this.getCNvProps();
        if(oNvPr){
            oNvPr.setDescr(sDescription ? sDescription : null);
        }
    };
    CGraphicObjectBase.prototype.getTitle = function(){
        var oNvPr = this.getCNvProps();
        if(oNvPr){
            return oNvPr.title ? oNvPr.title : undefined;
        }
        return undefined;
    };
    CGraphicObjectBase.prototype.getDescription = function(){
        var oNvPr = this.getCNvProps();
        if(oNvPr){
            return oNvPr.descr ? oNvPr.descr : undefined;
        }
        return undefined;
    };
    CGraphicObjectBase.prototype.setBDeleted = function(pr)
    {
        History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_ShapeSetBDeleted, this.bDeleted, pr));
        this.bDeleted = pr;
    };
    CGraphicObjectBase.prototype.getEditorType = function()
    {
        return 1;
    };
    CGraphicObjectBase.prototype.isEmptyPlaceholder = function()
    {
        return false;
    };
    CGraphicObjectBase.prototype.RestartSpellCheck = function()
    {
    };
    CGraphicObjectBase.prototype.GetAllFields = function(isUseSelection, arrFields)
    {
        return arrFields ? arrFields : [];
    };
    CGraphicObjectBase.prototype.GetAllSeqFieldsByType = function(sType, aFields)
    {
    };
    CGraphicObjectBase.prototype.convertToConnectionParams = function(rot, flipH, flipV, oTransform, oBounds, oConnectorInfo){
        var _ret =  new AscFormat.ConnectionParams();
        var _rot = oConnectorInfo.ang*AscFormat.cToRad + rot;
        var _normalized_rot = AscFormat.normalizeRotate(_rot);
        _ret.dir = AscFormat.CARD_DIRECTION_E;
        if(_normalized_rot >= 0 && _normalized_rot < Math.PI * 0.25 || _normalized_rot >= 7 * Math.PI * 0.25 && _normalized_rot < 2 * Math.PI){
            _ret.dir = AscFormat.CARD_DIRECTION_E;
            if(flipH){
                _ret.dir = AscFormat.CARD_DIRECTION_W;
            }
        }
        else if(_normalized_rot >= Math.PI * 0.25 && _normalized_rot < 3 * Math.PI * 0.25){
            _ret.dir = AscFormat.CARD_DIRECTION_S;
            if(flipV){
                _ret.dir = AscFormat.CARD_DIRECTION_N;
            }
        }
        else if(_normalized_rot >= 3 * Math.PI * 0.25 && _normalized_rot < 5 * Math.PI * 0.25){
            _ret.dir = AscFormat.CARD_DIRECTION_W;
            if(flipH){
                _ret.dir = AscFormat.CARD_DIRECTION_E;
            }
        }
        else if(_normalized_rot >= 5 * Math.PI * 0.25 && _normalized_rot < 7 * Math.PI * 0.25){
            _ret.dir = AscFormat.CARD_DIRECTION_N;
            if(flipV){
                _ret.dir = AscFormat.CARD_DIRECTION_S;
            }
        }
        _ret.x = oTransform.TransformPointX(oConnectorInfo.x, oConnectorInfo.y);
        _ret.y = oTransform.TransformPointY(oConnectorInfo.x, oConnectorInfo.y);
        _ret.bounds.fromOther(oBounds);
        _ret.idx = oConnectorInfo.idx;
        return _ret;
    };
    CGraphicObjectBase.prototype.getGeom = function () {

        var _geom;
        if(this.rectGeometry){
            _geom = this.rectGeometry;
        }
        else if(this.calcGeometry){
            _geom = this.calcGeometry;
        }
        else if(this.spPr && this.spPr.geometry){
            _geom = this.spPr.geometry;
        }
        else{
            _geom = AscFormat.ExecuteNoHistory(
                function(){
                    var _ret = AscFormat.CreateGeometry("rect");
                    _ret.Recalculate(this.extX, this.extY);
                    return _ret;
                }, this, []
            );
        }
        return _geom;
    };
    CGraphicObjectBase.prototype.findGeomConnector = function(x, y){
        var _geom = this.getGeom();
        var oInvertTransform = this.invertTransform;
        var _x = oInvertTransform.TransformPointX(x, y);
        var _y = oInvertTransform.TransformPointY(x, y);
        return _geom.findConnector(_x, _y, this.convertPixToMM(AscCommon.global_mouseEvent.KoefPixToMM * AscCommon.TRACK_CIRCLE_RADIUS));

    };
    CGraphicObjectBase.prototype.findConnector = function(x, y){
        var oConnGeom = this.findGeomConnector(x, y);
        if(oConnGeom){
            var _rot = this.rot;
            var _flipH = this.flipH;
            var _flipV = this.flipV;
            if(this.group){
                _rot = AscFormat.normalizeRotate(this.group.getFullRotate() + _rot);
                if(this.group.getFullFlipH()){
                    _flipH = !_flipH;
                }
                if(this.group.getFullFlipV()){
                    _flipV = !_flipV;
                }
            }
            return this.convertToConnectionParams(_rot, _flipH, _flipV, this.transform, this.bounds, oConnGeom);
        }
        return null;
    };
    CGraphicObjectBase.prototype.findConnectionShape = function(x, y){
        if(this.hit(x, y)){
            return this;
        }
        return null;
    };
    CGraphicObjectBase.prototype.getAllDocContents = function(aDrawings){

    };
    CGraphicObjectBase.prototype.getFullRotate = function () {
        return !AscCommon.isRealObject(this.group) ? this.rot : this.rot + this.group.getFullRotate();
    };
    CGraphicObjectBase.prototype.getAspect = function (num) {
        var _tmp_x = this.extX !== 0 ? this.extX : 0.1;
        var _tmp_y = this.extY !== 0 ? this.extY : 0.1;
        return num === 0 || num === 4 ? _tmp_x / _tmp_y : _tmp_y / _tmp_x;
    };
    CGraphicObjectBase.prototype.getFullFlipH = function () {
        if (!AscCommon.isRealObject(this.group))
            return this.flipH;
        return this.group.getFullFlipH() ? !this.flipH : this.flipH;
    };
    CGraphicObjectBase.prototype.getFullFlipV = function () {
        if (!AscCommon.isRealObject(this.group))
            return this.flipV;
        return this.group.getFullFlipV() ? !this.flipV : this.flipV;
    };
    CGraphicObjectBase.prototype.getMainGroup = function () {
        if(!AscCommon.isRealObject(this.group)){
            if(this.isGroupObject()){
                return this;
            }
            return null;
        }
        return this.group.getMainGroup();
    };
    CGraphicObjectBase.prototype.drawConnectors = function(overlay)
    {
        var _geom = this.getGeom();
        _geom.drawConnectors(overlay, this.transform);
    };
    CGraphicObjectBase.prototype.getConnectionParams = function(cnxIdx, _group)
    {
        AscFormat.ExecuteNoHistory(
            function(){
                if(this.recalculateSizes){
                    this.recalculateSizes();
                }
                else if(this.recalculateTransform){
                    this.recalculateTransform();
                }
            }, this, []
        );
        if(cnxIdx !== null){
            var oConnectionObject = this.getGeom().cnxLst[cnxIdx];
            if(oConnectionObject){
                var g_conn_info =  {idx: cnxIdx, ang: oConnectionObject.ang, x: oConnectionObject.x, y: oConnectionObject.y};
                var _rot = AscFormat.normalizeRotate(this.getFullRotate());
                var _flipH =  this.getFullFlipH();
                var _flipV =  this.getFullFlipV();
                var _bounds = this.bounds;
                var _transform = this.transform;

                if(_group){
                    _rot = AscFormat.normalizeRotate((this.group ? this.group.getFullRotate() : 0) + _rot - _group.getFullRotate());
                    if(_group.getFullFlipH()){
                        _flipH = !_flipH;
                    }
                    if(_group.getFullFlipV()){
                        _flipV = !_flipV;
                    }
                    _bounds = _bounds.copy();
                    _bounds.transform(_group.invertTransform);
                    _transform = _transform.CreateDublicate();
                    AscCommon.global_MatrixTransformer.MultiplyAppend(_transform, _group.invertTransform);
                }
                return this.convertToConnectionParams(_rot, _flipH, _flipV, _transform, _bounds, g_conn_info);
            }
        }
        return null;
    };
    CGraphicObjectBase.prototype.getCardDirectionByNum = function (num) {
        var num_north = this.getNumByCardDirection(AscFormat.CARD_DIRECTION_N);
        var full_flip_h = this.getFullFlipH();
        var full_flip_v = this.getFullFlipV();
        var same_flip = !full_flip_h && !full_flip_v || full_flip_h && full_flip_v;
        if (same_flip)
            return ((num - num_north) + AscFormat.CARD_DIRECTION_N + 8) % 8;

        return (AscFormat.CARD_DIRECTION_N - (num - num_north) + 8) % 8;
    };
    CGraphicObjectBase.prototype.getTransformMatrix = function(){
        return this.transform;
    };
    CGraphicObjectBase.prototype.getNumByCardDirection = function (cardDirection) {
        var hc = this.extX * 0.5;
        var vc = this.extY * 0.5;
        var transform = this.getTransformMatrix();
        var y1, y3, y5, y7;
        y1 = transform.TransformPointY(hc, 0);
        y3 = transform.TransformPointY(this.extX, vc);
        y5 = transform.TransformPointY(hc, this.extY);
        y7 = transform.TransformPointY(0, vc);

        var north_number;
        var full_flip_h = this.getFullFlipH();
        var full_flip_v = this.getFullFlipV();
        switch (Math.min(y1, y3, y5, y7)) {
            case y1:
            {
                north_number = 1;
                break;
            }
            case y3:
            {
                north_number = 3;
                break;
            }
            case y5:
            {
                north_number = 5;
                break;
            }
            default:
            {
                north_number = 7;
                break;
            }
        }
        var same_flip = !full_flip_h && !full_flip_v || full_flip_h && full_flip_v;

        if (same_flip)
            return (north_number + cardDirection) % 8;
        return (north_number - cardDirection + 8) % 8;
    };
    CGraphicObjectBase.prototype.getInvertTransform = function(){
        return this.invertTransform;
    };
    CGraphicObjectBase.prototype.getResizeCoefficients = function (numHandle, x, y, aDrawings) {
        var cx, cy;
        cx = this.extX > 0 ? this.extX : 0.01;
        cy = this.extY > 0 ? this.extY : 0.01;

        var invert_transform = this.getInvertTransform();
        if(!invert_transform) {
            return { kd1: 1, kd2: 1 };
        }
        var t_x = invert_transform.TransformPointX(x, y);
        var t_y = invert_transform.TransformPointY(x, y);

        var bSnapH = false;
        var bSnapV = false;

        var oSnapHorObject, oSnapVertObject;
        var dSnapX = null;
        var dSnapY = null;
        var bOwnH = false;
        var bOwnV = false;
        if(numHandle === 0 || numHandle === 6 || numHandle === 7) {
            if(Math.abs(t_x) < AscFormat.SNAP_DISTANCE) {
                t_x = 0;
                bSnapH = true;
                bOwnH = true;
            }
        }
        if(numHandle === 2 || numHandle === 3 || numHandle === 4) {
            if(Math.abs(t_x - this.extX) < AscFormat.SNAP_DISTANCE) {
                t_x = this.extX;
                bSnapH = true;
                bOwnH = true;
            }
        }

        if(numHandle === 0 || numHandle === 1 || numHandle === 2) {
            if(Math.abs(t_y) < AscFormat.SNAP_DISTANCE) {
                t_y = 0;
                bSnapV = true;
                bOwnV = true;
            }
        }

        if(numHandle === 4 || numHandle === 5 || numHandle === 6) {
            if(Math.abs(t_y - this.extY) < AscFormat.SNAP_DISTANCE) {
                t_y = this.extY;
                bSnapV = true;
                bOwnV = true;
            }
        }


        if(!bSnapH) {
            if(Array.isArray(aDrawings)) {
                oSnapHorObject = AscFormat.GetMinSnapDistanceXObject(x, aDrawings, this);
                if(oSnapHorObject) {
                    if(Math.abs(oSnapHorObject.dist) < AscFormat.SNAP_DISTANCE) {
                        bSnapH = true;
                    }
                    else {
                        oSnapHorObject = null;
                    }
                }
            }
        }
        if(!bSnapV) {
            if(Array.isArray(aDrawings)) {
                oSnapVertObject = AscFormat.GetMinSnapDistanceYObject(y, aDrawings, this);
                if(oSnapVertObject && Math.abs(oSnapVertObject.dist) < AscFormat.SNAP_DISTANCE) {
                    bSnapV = true;
                }
                else {
                    oSnapVertObject = null;
                }
            }
        }
        if(oSnapHorObject || oSnapVertObject) {
            var newX = oSnapHorObject ? x + oSnapHorObject.dist : x;
            var newY = oSnapVertObject ? y + oSnapVertObject.dist : y;
            if(oSnapHorObject || !bSnapH) {
                t_x = invert_transform.TransformPointX(newX, newY);
            }
            if(oSnapVertObject || !bSnapV) {
                t_y = invert_transform.TransformPointY(newX, newY);
            }
        }
        if(bSnapH && !bOwnH) {
            if(numHandle !== 1 && numHandle !== 5 &&
                !((numHandle === 7 || numHandle === 3) && !this.transform.IsIdentity2())) {
                dSnapX = this.transform.TransformPointX(t_x, t_y);
            }
        }
        if(bSnapV && !bOwnV) {
            if(numHandle !== 7 && numHandle !== 3 &&
                !((numHandle === 1 || numHandle === 5) && !this.transform.IsIdentity2())) {
                dSnapY = this.transform.TransformPointY(t_x, t_y);
            }
        }

        switch (numHandle) {
            case 0:
                return { kd1: (cx - t_x) / cx, kd2: (cy - t_y) / cy, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY};
            case 1:
                return { kd1: (cy - t_y) / cy, kd2: 0, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
            case 2:
                return { kd1: (cy - t_y) / cy, kd2: t_x / cx, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
            case 3:
                return { kd1: t_x / cx, kd2: 0, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
            case 4:
                return { kd1: t_x / cx, kd2: t_y / cy, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
            case 5:
                return { kd1: t_y / cy, kd2: 0, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
            case 6:
                return { kd1: t_y / cy, kd2: (cx - t_x) / cx, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
            case 7:
                return { kd1: (cx - t_x) / cx, kd2: 0, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
        }
        return { kd1: 1, kd2: 1, snapH: bSnapH, snapV: bSnapV, snapX: dSnapX, snapY: dSnapY };
    };
    CGraphicObjectBase.prototype.GetAllContentControls = function(arrContentControls) {};
	CGraphicObjectBase.prototype.GetAllDrawingObjects = function(arrDrawingObjects) {};
	CGraphicObjectBase.prototype.GetAllOleObjects = function(sPluginId, arrObjects) {
    };
    CGraphicObjectBase.prototype.CheckContentControlEditingLock = function () {
        if(this.group){
            this.group.CheckContentControlEditingLock();
            return;
        }
        if(this.parent && this.parent.CheckContentControlEditingLock){
            this.parent.CheckContentControlEditingLock();
        }
    };
    CGraphicObjectBase.prototype.hit = function(x, y) {
        return false;
    };
    CGraphicObjectBase.prototype.hitToAdjustment = function()
    {
        return {hit: false};
    };
    CGraphicObjectBase.prototype.hitToHandles = function (x, y) {
        if(this.parent && this.parent.kind === AscFormat.TYPE_KIND.NOTES){
            return -1;
        }
        if(this.isProtected && this.isProtected()) {
            return -1;
        }
        return AscFormat.hitToHandles(x, y, this);
    };
    CGraphicObjectBase.prototype.onMouseMove = function (e, x, y) {
        return this.hit(x, y);
    };
    CGraphicObjectBase.prototype.drawLocks = function(transform, graphics){
        if(AscCommon.IsShapeToImageConverter) {
            return;
        }
        var bNotes = !!(this.parent && this.parent.kind === AscFormat.TYPE_KIND.NOTES);
        if(!this.group && !bNotes)
        {
            var oLock;
            if(this.parent instanceof AscCommonWord.ParaDrawing)
            {
                oLock = this.parent.Lock;
            }
            else if(this.Lock)
            {
                oLock = this.Lock;
            }
            if(oLock && AscCommon.locktype_None !== oLock.Get_Type())
            {
                var bCoMarksDraw = true;
                var oApi = editor || Asc['editor'];
                if(oApi){

                    switch(oApi.getEditorId()){
                        case AscCommon.c_oEditorId.Word:{
                            bCoMarksDraw = (true === oApi.isCoMarksDraw || AscCommon.locktype_Mine !== oLock.Get_Type());
                            break;
                        }
                        case AscCommon.c_oEditorId.Presentation:{
                            bCoMarksDraw = (!AscCommon.CollaborativeEditing.Is_Fast() || AscCommon.locktype_Mine !== oLock.Get_Type());
                            break;
                        }
                        case AscCommon.c_oEditorId.Spreadsheet:{
                            bCoMarksDraw = (!oApi.collaborativeEditing.getFast() || AscCommon.locktype_Mine !== oLock.Get_Type());
                            break;
                        }
                    }
                }
                if(bCoMarksDraw && graphics.DrawLockObjectRect){
                    graphics.transform3(transform);
                    graphics.DrawLockObjectRect(oLock.Get_Type(), 0, 0, this.extX, this.extY);
                    return true;
                }
            }
        }
        return false;
    };
    CGraphicObjectBase.prototype.getSignatureLineGuid = function(){
        return null;
    };
    CGraphicObjectBase.prototype.getMacrosName = function(){
        var sGuid = this.getJSAMacroId();
        if(sGuid) {
            var oApi = Asc.editor || editor;
            if(oApi) {
                return oApi.asc_getMacrosByGuid(sGuid);
            }
        }
        return this.macro;
    };
    CGraphicObjectBase.prototype.getCopyWithSourceFormatting = function(oIdMap){
        return this.copy(oIdMap);
    };
    CGraphicObjectBase.prototype.checkNeedRecalculate = function(){
        return false;
    };
    CGraphicObjectBase.prototype.handleAllContents = function(fCallback){
    };
    CGraphicObjectBase.prototype.canFill = function() {
        return false;
    };
    CGraphicObjectBase.prototype.getPaddings = function() {
        return null;
    };
    CGraphicObjectBase.prototype.getBodyPr = function() {
        var oThis = this;
        return AscFormat.ExecuteNoHistory(function(){
            var ret = new AscFormat.CBodyPr();
            ret.setDefault();
            return ret;
        }, oThis, []);
    };
    CGraphicObjectBase.prototype.getTextArtProperties = function() {
        return null;
    };
    CGraphicObjectBase.prototype.getColumnNumber = function() {
        return null;
    };
    CGraphicObjectBase.prototype.getColumnSpace = function() {
        return null;
    };
    CGraphicObjectBase.prototype.getTextFitType = function() {
        return null;
    };
    CGraphicObjectBase.prototype.getVertOverflowType = function() {
        return null;
    };
    CGraphicObjectBase.prototype.canChangeArrows = function () {
        if (!this.spPr || this.spPr.geometry == null) {
            return false;
        }
        var _path_list = this.spPr.geometry.pathLst;
        var _path_index;
        var _path_command_index;
        var _path_command_arr;
        for (_path_index = 0; _path_index < _path_list.length; ++_path_index) {
            _path_command_arr = _path_list[_path_index].ArrPathCommandInfo;
            for (_path_command_index = 0; _path_command_index < _path_command_arr.length; ++_path_command_index) {
                if (_path_command_arr[_path_command_index].id == 5) {
                    break;
                }
            }
            if (_path_command_index == _path_command_arr.length) {
                return true;
            }
        }
        return false;
    };
    CGraphicObjectBase.prototype.getStroke = function () {
        if(this.pen && this.pen.Fill)
        {
            return this.pen;
        }
        var ret = AscFormat.CreateNoFillLine();
        ret.w = 0;
        return ret;
    };
    CGraphicObjectBase.prototype.getPresetGeom = function () {
        if (this.spPr && this.spPr.geometry) {
            return this.spPr.geometry.preset;
        }
        else {
            if(this.calcGeometry)
            {
                return this.calcGeometry.preset;
            }
            return null;
        }
    };
    CGraphicObjectBase.prototype.getFill = function () {
        if(this.brush && this.brush.fill)
        {
            return this.brush;
        }
        return AscFormat.CreateNoFillUniFill();
    };
    CGraphicObjectBase.prototype.getClipRect = function(){
        if(this.parent && this.parent.GetClipRect){
            return this.parent.GetClipRect();
        }
        return null;
    };
    CGraphicObjectBase.prototype.getBlipFill = function(){
        if(this.getObjectType() === AscDFH.historyitem_type_ImageShape || this.getObjectType() === AscDFH.historyitem_type_Shape){
            if(this.blipFill){
                return this.blipFill;
            }
            if(this.brush && this.brush.fill && this.brush.fill.type === window['Asc'].c_oAscFill.FILL_TYPE_BLIP){
                return this.brush.fill;
            }
        }
        return null;
    };
    CGraphicObjectBase.prototype.checkSrcRect = function(){

        if(this.getObjectType() === AscDFH.historyitem_type_ImageShape){
            if(this.blipFill.tile || !this.blipFill.srcRect || this.blipFill.stretch){

                var blipFill = this.blipFill.createDuplicate();
                if(blipFill.tile){
                    blipFill.tile = null;
                }
                if(!blipFill.srcRect){
                    blipFill.srcRect = new AscFormat.CSrcRect();
                    blipFill.srcRect.l = 0;
                    blipFill.srcRect.t = 0;
                    blipFill.srcRect.r = 100;
                    blipFill.srcRect.b = 100;
                }
                if(blipFill.stretch){
                    blipFill.stretch = null;
                }
                this.setBlipFill(blipFill);
            }
        }
        else{
            if(this.brush.fill.tile || !this.brush.fill.srcRect || this.brush.fill.stretch){
                var brush = this.brush.createDuplicate();
                if(brush.fill.tile){
                    brush.fill.tile = null;
                }
                if(!brush.fill.srcRect){
                    brush.fill.srcRect = new AscFormat.CSrcRect();
                    brush.fill.srcRect.l = 0;
                    brush.fill.srcRect.t = 0;
                    brush.fill.srcRect.r = 100;
                    brush.fill.srcRect.b = 100;
                }
                if(brush.fill.stretch){
                    brush.fill.stretch = null;
                }
                this.brush = brush;
                this.spPr.setFill(brush);
            }
        }
    };
    CGraphicObjectBase.prototype.getCropObject = function(){
        if(!this.cropObject){
            this.createCropObject();
        }
        return this.cropObject;
    };
    CGraphicObjectBase.prototype.createCropObject = function(){
        return AscFormat.ExecuteNoHistory(function () {
            var oBlipFill = this.getBlipFill();
            if(!oBlipFill){
                return;
            }
            var srcRect = oBlipFill.srcRect;
            if(srcRect)
            {
                var sRasterImageId = oBlipFill.RasterImageId;
                var _l = srcRect.l ? srcRect.l : 0;
                var _t = srcRect.t ? srcRect.t : 0;
                var _r = srcRect.r ? srcRect.r : 100;
                var _b = srcRect.b ? srcRect.b : 100;
                var oShapeDrawer = new AscCommon.CShapeDrawer();
                oShapeDrawer.bIsCheckBounds = true;
                oShapeDrawer.Graphics = new AscFormat.CSlideBoundsChecker();
                this.check_bounds(oShapeDrawer);
                var boundsW = oShapeDrawer.max_x - oShapeDrawer.min_x;
                var boundsH = oShapeDrawer.max_y - oShapeDrawer.min_y;
                var wpct = (_r - _l)/100.0;
                var hpct = (_b - _t)/100.0;
                var extX = boundsW/wpct;
                var extY = boundsH/hpct;
                var DX = -extX*_l/100.0 + oShapeDrawer.min_x;
                var DY = -extY*_t/100.0 + oShapeDrawer.min_y;
                var XC = DX + extX/2.0;
                var YC = DY + extY/2.0;

                var oTransform = this.transform.CreateDublicate();
                // if(this.group)
                // {
                //     AscCommon.global_MatrixTransformer.MultiplyAppend(oTransform, this.group.invertTransform);
                // }

                var XC_ = oTransform.TransformPointX(XC, YC);
                var YC_ = oTransform.TransformPointY(XC, YC);

                var X = XC_ - extX/2.0;
                var Y = YC_ - extY/2.0;

                var oImage = AscFormat.DrawingObjectsController.prototype.createImage(sRasterImageId, X, Y, extX, extY);
                oImage.isCrop = true;
                oImage.parentCrop = this;
                oImage.worksheet = this.worksheet;
                oImage.drawingBase = this.drawingBase;
                oImage.spPr.xfrm.setRot(this.rot);
                oImage.spPr.xfrm.setFlipH(this.flipH);
                oImage.spPr.xfrm.setFlipV(this.flipV);
                // oImage.setGroup(this.group);


                oImage.setParent(this.parent);
                oImage.recalculate();
                oImage.setParent(null);
                oImage.recalculateTransform();
                oImage.recalculateGeometry();
                oImage.invertTransform = AscCommon.global_MatrixTransformer.Invert(oImage.transform);
                oImage.recalculateBounds();
                oImage.setParent(this.parent);
                oImage.selectStartPage = this.selectStartPage;
                oImage.cropBrush = AscFormat.CreateUnfilFromRGB(128, 128, 128);
                oImage.cropBrush.transparent = 100;
                oImage.pen = AscFormat.CreatePenBrushForChartTrack().pen;
                oImage.parent = this.parent;
                var oParentObjects = this.getParentObjects();
                oImage.cropBrush.calculate(oParentObjects.theme, oParentObjects.slide, oParentObjects.layout, oParentObjects.master, {R:0, G:0, B:0, A:255, needRecalc: true}, AscFormat.G_O_DEFAULT_COLOR_MAP);
                this.cropObject = oImage;
                return true;
            }
            return false;
        }, this, []);
    };
    CGraphicObjectBase.prototype.clearCropObject = function(){
        this.cropObject = null;
    };
    CGraphicObjectBase.prototype.drawCropTrack = function(graphics, srcRect, transform, cropObjectTransform){

    };
    CGraphicObjectBase.prototype.calculateSrcRect = function(){

        var oldTransform = this.transform.CreateDublicate();
        var oldExtX = this.extX;
        var oldExtY = this.extY;
        AscFormat.ExecuteNoHistory(function(){
            // this.cropObject.recalculateTransform();
            // this.recalculateTransform();
            var oldVal = this.recalcInfo.recalculateTransform;
            this.recalcInfo.recalculateTransform = false;
            this.recalculateGeometry();
            this.recalcInfo.recalculateTransform = oldVal;
        }, this, []);
        this.transform = oldTransform;
        this.extX = oldExtX;
        this.extY = oldExtY;
        this.setSrcRect(this.calculateSrcRect2());
        this.clearCropObject();
    };
    CGraphicObjectBase.prototype.setSrcRect = function(srcRect){

        if(this.getObjectType() === AscDFH.historyitem_type_ImageShape)
        {
            var blipFill = this.blipFill.createDuplicate();
            blipFill.srcRect = srcRect;
            this.setBlipFill(blipFill);
        }
        else
        {
            var brush = this.brush.createDuplicate();
            brush.fill.srcRect = srcRect;
            this.spPr.setFill(brush);
        }
    };
    CGraphicObjectBase.prototype.calculateSrcRect2 = function(){

        var oShapeDrawer = new AscCommon.CShapeDrawer();
        oShapeDrawer.bIsCheckBounds = true;
        oShapeDrawer.Graphics = new AscFormat.CSlideBoundsChecker();
        this.check_bounds(oShapeDrawer);
        return  CalculateSrcRect(this.transform, oShapeDrawer, this.cropObject.invertTransform, this.cropObject.extX, this.cropObject.extY);
    };
    CGraphicObjectBase.prototype.getMediaFileName = function(){
        return null;
    };
    CGraphicObjectBase.prototype.getLogicDocument = function()
    {
        var oApi = editor || Asc['editor'];
        if(oApi && oApi.WordControl)
        {
            return oApi.WordControl.m_oLogicDocument;
        }
        return null;
    };
    CGraphicObjectBase.prototype.updatePosition = function(x, y) {
        this.posX = x;
        this.posY = y;
        if(!this.group){
            this.x = this.localX + x;
            this.y = this.localY + y;
        }
        else{
            this.x = this.localX;
            this.y = this.localY;
        }
        if(this.updateTransformMatrix) {
            this.updateTransformMatrix();
        }
    };
    CGraphicObjectBase.prototype.copyComments = function(oLogicDocument)
    {
        if(!oLogicDocument)
        {
            return;
        }

        var aDocContents = [];
        this.getAllDocContents(aDocContents);
        for(var i = 0; i < aDocContents.length; ++i)
        {
        	aDocContents[i].CreateDuplicateComments();
        }
    };

    CGraphicObjectBase.prototype.writeRecord1 = function(pWriter, nType, oChild) {
        if(AscCommon.isRealObject(oChild)) {
            pWriter.WriteRecord1(nType, oChild, function(oChild) {
                oChild.toPPTY(pWriter);
            });
        }
        else {
            //TODO: throw an error
        }
    };
    CGraphicObjectBase.prototype.writeRecord2 = function(pWriter, nType, oChild) {
        if(AscCommon.isRealObject(oChild)) {
            this.writeRecord1(pWriter, nType, oChild);
        }
    };

    CGraphicObjectBase.prototype.ResetParametersWithResize = function(bNoResetRelSize)
    {
        var oParaDrawing = AscFormat.getParaDrawing(this);
        if(oParaDrawing && !(bNoResetRelSize === true))
        {
            if(oParaDrawing.SizeRelH)
            {
                oParaDrawing.SetSizeRelH(undefined);
            }
            if(oParaDrawing.SizeRelV)
            {
                oParaDrawing.SetSizeRelV(undefined);
            }
        }
        if(this instanceof AscFormat.CShape) {
            var oPropsToSet = null;
            if (this.bWordShape) {
                if (!this.textBoxContent)
                    return;
                if (this.bodyPr) {
                    oPropsToSet = this.bodyPr.createDuplicate();
                } else {
                    oPropsToSet = new AscFormat.CBodyPr();
                }
            } else {
                if (!this.txBody)
                    return;
                if (this.txBody.bodyPr) {
                    oPropsToSet = this.txBody.bodyPr.createDuplicate();
                } else {
                    oPropsToSet = new AscFormat.CBodyPr();
                }
            }
            var oBodyPr = this.getBodyPr();
            if (this.bWordShape) {
                if(oBodyPr.textFit && oBodyPr.textFit.type === AscFormat.text_fit_Auto)
                {
                    if(!oPropsToSet.textFit)
                    {
                        oPropsToSet.textFit = new AscFormat.CTextFit();
                    }
                    oPropsToSet.textFit.type = AscFormat.text_fit_No;
                }
            }
            if (oBodyPr.wrap === AscFormat.nTWTNone) {
                oPropsToSet.wrap = AscFormat.nTWTSquare;
            }
            if (this.bWordShape) {
                this.setBodyPr(oPropsToSet);
            } else {
                this.txBody.setBodyPr(oPropsToSet);
                if (this.checkExtentsByDocContent) {
                    this.checkExtentsByDocContent(true, true);
                }
            }
        }
    };
    CGraphicObjectBase.prototype.createPlaceholderControl = function(aControls)
    {
        if(!this.isEmptyPlaceholder())
        {
            return;
        }
        var phType = this.getPhType();
        var aButtons = [];
        var isLocalDesktop = window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsSupportMedia"] && window["AscDesktopEditor"]["IsSupportMedia"]();
        switch (phType)
        {
            case null:
            {
                aButtons.push(AscCommon.PlaceholderButtonType.Table);
                aButtons.push(AscCommon.PlaceholderButtonType.Chart);
                aButtons.push(AscCommon.PlaceholderButtonType.Image);
                aButtons.push(AscCommon.PlaceholderButtonType.ImageUrl);
                if(isLocalDesktop)
                {
                    aButtons.push(AscCommon.PlaceholderButtonType.Video);
                    aButtons.push(AscCommon.PlaceholderButtonType.Audio);
                }
                break;
            }
            case AscFormat.phType_body:
            {
                break;
            }
            case AscFormat.phType_chart:
            {
                aButtons.push(AscCommon.PlaceholderButtonType.Chart);
                break;
            }
            case AscFormat.phType_clipArt:
            {
                aButtons.push(AscCommon.PlaceholderButtonType.Image);
                aButtons.push(AscCommon.PlaceholderButtonType.ImageUrl);
                break;
            }
            case AscFormat.phType_ctrTitle:
            {
                break;
            }
            case AscFormat.phType_dgm:
            {
                break;
            }
            case AscFormat.phType_dt:
            {
                break;
            }
            case AscFormat.phType_ftr:
            {
                break;
            }
            case AscFormat.phType_hdr:
            {
                break;
            }
            case AscFormat.phType_media:
            {
                if(isLocalDesktop)
                {
                    aButtons.push(AscCommon.PlaceholderButtonType.Video);
                    aButtons.push(AscCommon.PlaceholderButtonType.Audio);
                }
                break;
            }
            case AscFormat.phType_obj:
            {
                aButtons.push(AscCommon.PlaceholderButtonType.Table);
                aButtons.push(AscCommon.PlaceholderButtonType.Chart);
                aButtons.push(AscCommon.PlaceholderButtonType.Image);
                aButtons.push(AscCommon.PlaceholderButtonType.ImageUrl);
                if(isLocalDesktop)
                {
                    aButtons.push(AscCommon.PlaceholderButtonType.Video);
                    aButtons.push(AscCommon.PlaceholderButtonType.Audio);
                }
                break;
            }
            case AscFormat.phType_pic:
            {

                aButtons.push(AscCommon.PlaceholderButtonType.Image);
                aButtons.push(AscCommon.PlaceholderButtonType.ImageUrl);
                break;
            }
            case AscFormat.phType_sldImg:
            {
                aButtons.push(AscCommon.PlaceholderButtonType.Image);
                aButtons.push(AscCommon.PlaceholderButtonType.ImageUrl);
                break;
            }
            case AscFormat.phType_sldNum:
            {
                break;
            }
            case AscFormat.phType_subTitle:
            {
                break;
            }
            case AscFormat.phType_tbl:
            {
                aButtons.push(AscCommon.PlaceholderButtonType.Table);
                break;
            }
            case AscFormat.phType_title:
            {
                break;
            }
        }
        var nSlideNum = 0;
        if(this.parent.getObjectType && this.parent.getObjectType() === AscDFH.historyitem_type_Slide)
        {
            nSlideNum = this.parent.num;
        }
        if(aButtons.length > 0)
        {
            aControls.push(AscCommon.CreateDrawingPlaceholder(this.Id, aButtons, nSlideNum, { x : 0, y : 0, w : this.extX, h : this.extY }, this.transform));
        }
    };
    CGraphicObjectBase.prototype.onSlicerUpdate = function(sName){
        return false;
    };

    CGraphicObjectBase.prototype.onSlicerLock = function(sName, bLock){
    };
    CGraphicObjectBase.prototype.onSlicerDelete = function(sName){
        return false;
    };
    CGraphicObjectBase.prototype.onSlicerChangeName = function(sName, sNewName){
        return false;
    };
    CGraphicObjectBase.prototype.onUpdate = function (oRect) {
        if(this.drawingBase) {
            this.drawingBase.onUpdate(oRect);
        }
        else {
            if(this.group) {
                this.group.onUpdate(oRect)
            }
        }
    };
    CGraphicObjectBase.prototype.getSlicerViewByName = function (name) {
        return null;
    };
    CGraphicObjectBase.prototype.setParent2 = function(parent) {
        this.setParent(parent);
        if(Array.isArray(this.spTree)) {
            for(var i = 0; i < this.spTree.length; ++i) {
                this.spTree[i].setParent2(parent);
            }
        }
    };
    CGraphicObjectBase.prototype.documentCreateFontMap = function(oMap) {
    };
    CGraphicObjectBase.prototype.createFontMap = function(oMap) {
        this.documentCreateFontMap(oMap);
    };
    CGraphicObjectBase.prototype.isComparable = function(oDrawing) {
        var oPr = this.getCNvProps();
        var oOtherPr = oDrawing.getCNvProps();
        if(!oPr && !oOtherPr) {
            return true;
        }
        if(!oPr)  {
            return false;
        }
        return oPr.hasSameNameAndId(oOtherPr);
    };
    CGraphicObjectBase.prototype.select = function (drawingObjectsController, pageIndex)
    {
        this.selected = true;
        this.selectStartPage = pageIndex;
        var content = this.getDocContent && this.getDocContent();
        if(content)
            content.Set_StartPage(pageIndex);
        var selected_objects;
        if (!AscCommon.isRealObject(this.group))
            selected_objects = drawingObjectsController ? drawingObjectsController.selectedObjects : [];
        else
            selected_objects = this.group.getMainGroup().selectedObjects;
        for (var i = 0; i < selected_objects.length; ++i) {
            if (selected_objects[i] === this)
                break;
        }
        if (i === selected_objects.length)
            selected_objects.push(this);

    
        if(drawingObjectsController) {
            drawingObjectsController.onChangeDrawingsSelection();
        }
    };
    CGraphicObjectBase.prototype.deselect = function (drawingObjectsController) {
        this.selected = false;
        var selected_objects;
        if (!AscCommon.isRealObject(this.group))
            selected_objects = drawingObjectsController ? drawingObjectsController.selectedObjects : [];
        else
            selected_objects = this.group.getMainGroup().selectedObjects;
        for (var i = 0; i < selected_objects.length; ++i) {
            if (selected_objects[i] === this) {
                selected_objects.splice(i, 1);
                break;
            }
        }
        if(this.graphicObject)
        {
            this.graphicObject.RemoveSelection();
        }
        
        if(drawingObjectsController) {
            drawingObjectsController.onChangeDrawingsSelection();
        }
        return this;
    };
    CGraphicObjectBase.prototype.Set_CurrentElement = function(bUpdate, pageIndex) {
        //TODO: refactor this
        if(AscFormat.CShape.prototype.Set_CurrentElement) {
            AscFormat.CShape.prototype.Set_CurrentElement.call(this, bUpdate, pageIndex);
        }
    };


    CGraphicObjectBase.prototype.SetControllerTextSelection = function(drawing_objects, nPageIndex) {
        if(drawing_objects) {
            var oContent = this.getDocContent && this.getDocContent();
            drawing_objects.resetSelection(true);
            if(this.group) {
                var main_group = this.group.getMainGroup();
                drawing_objects.selectObject(main_group, nPageIndex);
                main_group.selectObject(this, nPageIndex);
                if(oContent) {
                    main_group.selection.textSelection = this;
                }
                drawing_objects.selection.groupSelection = main_group;
            }
            else {
                drawing_objects.selectObject(this, nPageIndex);
                if(oContent) {
                    drawing_objects.selection.textSelection = this;
                }
            }
        }
    };

    CGraphicObjectBase.prototype.hitInBoundingRect = function (x, y) {
        if(this.parent && this.parent.kind === AscFormat.TYPE_KIND.NOTES){
            return false;
        }
        var invert_transform = this.getInvertTransform();
        if(!invert_transform)
        {
            return false;
        }
        var x_t = invert_transform.TransformPointX(x, y);
        var y_t = invert_transform.TransformPointY(x, y);

        var _hit_context = this.getCanvasContext();

        return !(AscFormat.CheckObjectLine(this)) && (AscFormat.HitInLine(_hit_context, x_t, y_t, 0, 0, this.extX, 0) ||
            AscFormat.HitInLine(_hit_context, x_t, y_t, this.extX, 0, this.extX, this.extY) ||
            AscFormat.HitInLine(_hit_context, x_t, y_t, this.extX, this.extY, 0, this.extY) ||
            AscFormat.HitInLine(_hit_context, x_t, y_t, 0, this.extY, 0, 0) ||
            (this.canRotate && this.canRotate() && AscFormat.HitInLine(_hit_context, x_t, y_t, this.extX * 0.5, 0, this.extX * 0.5, -this.convertPixToMM(AscCommon.TRACK_DISTANCE_ROTATE))));
    };
    CGraphicObjectBase.prototype.getCanvasContext = function() {
        return AscFormat.CShape.prototype.getCanvasContext.call(this);
    };
	CGraphicObjectBase.prototype.isForm = function() {
		return (this.parent && this.parent.IsForm && this.parent.IsForm());
	};
	CGraphicObjectBase.prototype.getFormHorPadding = function() {
		return 0;
	};
	CGraphicObjectBase.prototype.getInnerForm = function() {
		return null;
	};
	CGraphicObjectBase.prototype.isContainedInTopDocument = function() {
		var oParentContent = this.parent && this.parent.DocumentContent;
        if(!oParentContent) {
            return true;
        }
        return (oParentContent === oParentContent.GetLogicDocument());
	};
	CGraphicObjectBase.prototype.isContainedInMainDoc = function() {
		var oParentContent = this.parent && this.parent.DocumentContent;
        if(!oParentContent) {
            return true;
        }
        return (oParentContent.GetTopDocumentContent() === oParentContent.GetLogicDocument());
	};
    CGraphicObjectBase.prototype.getBoundsByDrawing = function() {
        var oCopy = this.bounds.copy();
        oCopy.l -= 3;
        oCopy.r += 3;
        oCopy.t -= 3;
        oCopy.b += 3;
        oCopy.checkWH();
        return oCopy;//TODO: do not count shape rect
    };
    CGraphicObjectBase.prototype.getPresentationSize = function() {
        var oPresentation = editor.WordControl.m_oLogicDocument;
        return {
            w: oPresentation.GetWidthMM(),
            h: oPresentation.GetHeightMM()
        }
    };
    CGraphicObjectBase.prototype.getAnimTexture = function(scale) {
        var oBounds = this.getBoundsByDrawing();
        var nWidth = oBounds.w * scale + 0.5 >> 0;
        var nHeight = oBounds.h * scale + 0.5 >> 0;
        if(nWidth === 0 || nHeight === 0) {
            return null;
        }
        var oCanvas = document.createElement('canvas');
        oCanvas.width = nWidth;
        oCanvas.height = nHeight;
        var oCtx = oCanvas.getContext('2d');
        var oGraphics = new AscCommon.CGraphics();
        oGraphics.init(oCtx, nWidth, nHeight, nWidth / scale, nHeight / scale);
        oGraphics.m_oFontManager = AscCommon.g_fontManager;
        var nX = oBounds.x * oGraphics.m_oCoordTransform.sx;
        var nY = oBounds.y * oGraphics.m_oCoordTransform.sy;
        oGraphics.m_oCoordTransform.tx = -nX;
        oGraphics.m_oCoordTransform.ty = -nY;
        AscCommon.IsShapeToImageConverter = true;
        this.draw(oGraphics);
        AscCommon.IsShapeToImageConverter = false;
        return new AscFormat.CBaseAnimTexture(oCanvas, scale, nX, nY)
    };

    CGraphicObjectBase.prototype.isOnProtectedSheet = function() {
        if(this.worksheet) {
            if(this.worksheet.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
                return true;
            }
        }
        return false;
    };
    CGraphicObjectBase.prototype.isProtectedText = function() {
        if(this.getProtectionLockText()) {
            if(this.isOnProtectedSheet()) {
                return true;
            }
        }
        return false;
    };
    CGraphicObjectBase.prototype.isProtected = function() {
        if(this.getProtectionLocked()) {
            if(this.isOnProtectedSheet()) {
                return true;
            }
        }
        return false;
    };
	CGraphicObjectBase.prototype.canEditText = function() {
        if(this.getObjectType() === AscDFH.historyitem_type_Shape) {
            if(!AscFormat.CheckLinePresetForParagraphAdd(this.getPresetGeom()) && !this.signatureLine) {
                if(this.isProtectedText()) {
                    return false;
                }
                if (this.isObjectInSmartArt()) {
                    return this.canEditTextInSmartArt();
                }
                return true;
            }
        }
        return false;
	};
    CGraphicObjectBase.prototype.canEdit = function() {
        if(this.isProtected()) {
            return false;
        }
        return true;
    };
    //for bug 52775. remove in the next version
    CGraphicObjectBase.prototype.applySmartArtTextStyle = function() {

    };
    CGraphicObjectBase.prototype.convertFromSmartArt = function() {
        return this;
    };

    CGraphicObjectBase.prototype.getDefaultRotSA = function () {
        if (this.isObjectInSmartArt()) {
            var point = this.getSmartArtShapePoint();
            if (point) {
                var custAng = point.getCustAng();
                if (custAng) {
                    var defaultRot = this.spPr && this.spPr.xfrm && this.spPr.xfrm.rot || 0;
                    if (custAng) {
                        if (custAng > defaultRot) {
                            defaultRot += Math.PI * 2 * Math.ceil(custAng / (Math.PI * 2));
                        }
                        defaultRot -= custAng;
                        return AscFormat.normalizeRotate(defaultRot);
                    }
                }
            }
            return AscFormat.normalizeRotate(this.rot);
        }
    };

    CGraphicObjectBase.prototype.changePositionInSmartArt = function (newX, newY) {}

    CGraphicObjectBase.prototype.changeRot = function(dAngle, bWord) {
        var oSmartArt;
        if (this.isObjectInSmartArt()) {
            oSmartArt = this.group.group;
            if (this.extX > oSmartArt.extX || this.extY > oSmartArt.extY || this.extX > oSmartArt.extY || this.extY > oSmartArt.extX) {
                return;
            }
        }

        if(this.spPr && this.spPr.xfrm) {
            var oXfrm = this.spPr.xfrm;
            var originalRot = oXfrm.rot || 0;
            var dRot = AscFormat.normalizeRotate(dAngle);
            oXfrm.setRot(dRot);
            if(this.isObjectInSmartArt()) {
                oSmartArt = this.group.group;
                var point = this.getSmartArtShapePoint();
                if (point) {
                    var prSet = point.getPrSet();
                    if (prSet) {
                        var defaultRot = originalRot;
                        if (prSet.custAng) {
                            var oldCustAng = prSet.custAng;
                            if (oldCustAng > defaultRot) {
                                defaultRot += Math.PI * 2 * Math.ceil(oldCustAng / (Math.PI * 2));
                            }
                            defaultRot -= oldCustAng;
                        }
                        if (originalRot !== dRot) {
                            var currentAngle = dRot;
                            if (defaultRot > currentAngle) {
                                currentAngle += Math.PI * 2 * Math.ceil(defaultRot / (Math.PI * 2));
                            }
                            var newCustAng = (currentAngle - defaultRot);
                            prSet.setCustAng(newCustAng);
                        }
                    }
                }
                this.recalculate();
                var oBounds = this.bounds;
                var diffX = null, diffY = null;
                var leftEdgeOfSmartArt = oSmartArt.x;
                var topEdgeOfSmartArt = oSmartArt.y;
                var rightEdgeOfSmartArt = oSmartArt.x + oSmartArt.extX;
                var bottomEdgeOfSmartArt = oSmartArt.y + oSmartArt.extY;
                if (bWord) {
                    oBounds = {
                        l: oBounds.l + leftEdgeOfSmartArt,
                        r: oBounds.r + leftEdgeOfSmartArt,
                        t: oBounds.t + topEdgeOfSmartArt,
                        b: oBounds.b + topEdgeOfSmartArt
                    };
                }
                if(oBounds.r > rightEdgeOfSmartArt) {
                    diffX = rightEdgeOfSmartArt - oBounds.r;
                }
                if(oBounds.l < leftEdgeOfSmartArt) {
                    diffX = leftEdgeOfSmartArt - oBounds.l;
                }
                if(oBounds.b > bottomEdgeOfSmartArt) {
                    diffY = bottomEdgeOfSmartArt - oBounds.b;
                }
                if(oBounds.t < topEdgeOfSmartArt) {
                    diffY = topEdgeOfSmartArt - oBounds.t;
                }

                if(diffX !== null) {
                    var newOffX = this.spPr.xfrm.offX + diffX;
                    this.spPr.xfrm.setOffX(newOffX);
                    this.txXfrm && this.txXfrm.setOffX(this.txXfrm.offX + diffX);
                }
                if(diffY !== null) {
                    var newOffY = this.spPr.xfrm.offY + diffY;
                    this.spPr.xfrm.setOffY(newOffY);
                    this.txXfrm && this.txXfrm.setOffY(this.txXfrm.offY + diffY);
                }

                var posX = this.spPr.xfrm.offX;
                var posY = this.spPr.xfrm.offY;
                this.changePositionInSmartArt(posX, posY);
            }
        }
    };
    CGraphicObjectBase.prototype.changeFlipH = function(bFlipH) {
        if(this.spPr && this.spPr.xfrm) {
            var oXfrm = this.spPr.xfrm;
            oXfrm.setFlipH(bFlipH);
            if(this.isObjectInSmartArt()) {
                var point = this.getSmartArtShapePoint();
                point && point.changeFlipH(bFlipH);
            }
        }
    };
    CGraphicObjectBase.prototype.changeFlipV = function(bFlipV) {
        if(this.spPr && this.spPr.xfrm) {
            var oXfrm = this.spPr.xfrm;
            oXfrm.setFlipV(bFlipV);
            if(this.isObjectInSmartArt()) {
                var point = this.getSmartArtShapePoint();
                point && point.changeFlipV(bFlipV);
            }
        }
    };
    CGraphicObjectBase.prototype.getTypeName = function() {
        return AscCommon.translateManager.getValue("Graphic Object");
    };
    CGraphicObjectBase.prototype.getObjectName = function() {
		var oCNvPr = this.getCNvProps();
		if(oCNvPr && typeof oCNvPr.name === "string" && oCNvPr.name.length > 0) {
			return oCNvPr.name;
		}
        return this.getTypeName() + " " + this.getFormatId();
    };
    
    CGraphicObjectBase.prototype.isPlaceholder = function() {
        return false;
    };
    CGraphicObjectBase.prototype.getPlaceholderName = function() {
        if(!this.isPlaceholder()) {
            return "";
        }
        var nPhType = this.getPlaceholderType();
        var sText = AscFormat.pHText[nPhType];
        if(!sText) {
            sText = AscFormat.pHText[AscFormat.phType_body];
        }
        var sTrText = AscCommon.translateManager.getValue(sText);
        return sTrText;
    };
    CGraphicObjectBase.prototype.getTimimng = function() {
        if(this.group) {
            return this.group.getTimimng();
        }
        return (this.parent && this.parent.timing) || null;
    };
    CGraphicObjectBase.prototype.drawAnimLabels = function(oGraphics) {
        if(oGraphics.IsThumbnail === true || oGraphics.IsDemonstrationMode === true 
            || AscCommon.IsShapeToImageConverter || oGraphics.DO_NOT_DRAW_ANIM_LABEL) {
            return;
        }
        if(this.group) {
            return;
        }
        var oApi = Asc.editor || editor;
        if(!oApi.isDrawAnimLabels || !oApi.isDrawAnimLabels()) {
            return;
        }
        var oTiming = this.getTimimng();
        if(!oTiming) {
            return;
        }
        var aEffects = oTiming.getObjectEffects(this.Get_Id());
        if(aEffects.length === 0) {
            return;
        }
        var oBounds = this.bounds;
        var dX = oBounds.x - this.convertPixToMM(9) - this.convertPixToMM(ANIM_LABEL_WIDTH_PIX);
        var dY = oBounds.y;
        var dW = this.convertPixToMM(ANIM_LABEL_WIDTH_PIX);
        var dH = this.convertPixToMM(ANIM_LABEL_HEIGHT_PIX);
        for(var nEffect = 0; nEffect < aEffects.length; ++nEffect) {
            aEffects[nEffect].drawEffectLabel(oGraphics, dX, dY, dW, dH);
            dY += (this.convertPixToMM(4) + dH);
        }
    };
    
    CGraphicObjectBase.prototype.isImage = function() {
        return this.getObjectType() === AscDFH.historyitem_type_ImageShape;
    };

    CGraphicObjectBase.prototype.Is_UseInDocument = function() {
        if(CShape.prototype.Is_UseInDocument) {
            return CShape.prototype.Is_UseInDocument.call(this);
        }
        return true;
    };

    

    var ANIM_LABEL_WIDTH_PIX = 22;
    var ANIM_LABEL_HEIGHT_PIX = 17;


    function CRelSizeAnchor() {
        CBaseObject.call(this);
        this.fromX = null;
        this.fromY = null;

        this.toX = null;
        this.toY = null;

        this.object = null;

        this.parent = null;
        this.drawingBase = null;
    }
    CRelSizeAnchor.prototype = Object.create(CBaseObject.prototype);
    CRelSizeAnchor.prototype.constructor = CRelSizeAnchor;
    CRelSizeAnchor.prototype.setDrawingBase = function(drawingBase){
        this.drawingBase = drawingBase;
    };
    CRelSizeAnchor.prototype.getObjectType = function () {
        return AscDFH.historyitem_type_RelSizeAnchor;
    };
    CRelSizeAnchor.prototype.setFromTo = function (fromX, fromY, toX, toY) {
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_RelSizeAnchorFromX, this.fromX, fromX));
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_RelSizeAnchorFromY, this.fromY, fromY));
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_RelSizeAnchorToX, this.toX, toX));
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_RelSizeAnchorToY, this.toY, toY));
        this.fromX = fromX;
        this.fromY = fromY;
        this.toX = toX;
        this.toY = toY;
    };
    CRelSizeAnchor.prototype.setObject = function (object) {
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_RelSizeAnchorObject, this.object, object));
        this.object = object;
        if(object){
            object.setParent(this);
        }
    };
    CRelSizeAnchor.prototype.setParent = function (object) {
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_RelSizeAnchorParent, this.parent, object));
        this.parent = object;
    };
    CRelSizeAnchor.prototype.copy = function(oPr){
        var copy = new CRelSizeAnchor();
        copy.setFromTo(this.fromX, this.fromY, this.toX, this.toY);
        if(this.object){
            copy.setObject(this.object.copy(oPr));
        }
        return copy;
    };
    CRelSizeAnchor.prototype.Refresh_RecalcData = function(drawingDocument){
        if(this.parent && this.parent.Refresh_RecalcData2)
        {
            this.parent.Refresh_RecalcData2();
        }
    };
    CRelSizeAnchor.prototype.Refresh_RecalcData2 = function(drawingDocument){
        if(this.parent && this.parent.Refresh_RecalcData2)
        {
            this.parent.Refresh_RecalcData2();
        }
    };

    AscDFH.drawingsChangesMap[AscDFH.historyitem_RelSizeAnchorFromX]  = function(oClass, value){oClass.fromX =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_RelSizeAnchorFromY]  = function(oClass, value){oClass.fromY =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_RelSizeAnchorToX]    = function(oClass, value){oClass.toX =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_RelSizeAnchorToY]    = function(oClass, value){oClass.toY =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_RelSizeAnchorObject] = function(oClass, value){oClass.object =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_RelSizeAnchorParent] = function(oClass, value){oClass.parent =  value;};

    AscDFH.changesFactory[AscDFH.historyitem_RelSizeAnchorFromX] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_RelSizeAnchorFromY] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_RelSizeAnchorToX] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_RelSizeAnchorToY] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_RelSizeAnchorObject] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_RelSizeAnchorParent] = window['AscDFH'].CChangesDrawingsObject;


    function CAbsSizeAnchor(){
        CBaseObject.call(this);
        this.fromX = null;
        this.fromY = null;
        this.toX = null;
        this.toY = null;
        this.object = null;

        this.parent = null;
        this.drawingBase = null;
    }


    CAbsSizeAnchor.prototype = Object.create(CBaseObject.prototype);
    CAbsSizeAnchor.prototype.constructor = CAbsSizeAnchor;
    CAbsSizeAnchor.prototype.setDrawingBase = function(drawingBase){
        this.drawingBase = drawingBase;
    };
    CAbsSizeAnchor.prototype.getObjectType = function () {
        return AscDFH.historyitem_type_AbsSizeAnchor;
    };
    CAbsSizeAnchor.prototype.setFromTo = function (fromX, fromY, extX, extY) {
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_AbsSizeAnchorFromX, this.fromX, fromX));
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_AbsSizeAnchorFromY, this.fromY, fromY));
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_AbsSizeAnchorExtX, this.toX, extX));
        History.Add(new AscDFH.CChangesDrawingsDouble(this, AscDFH.historyitem_AbsSizeAnchorExtY, this.toY, extY));
        this.fromX = fromX;
        this.fromY = fromY;
        this.toX = extX;
        this.toY = extY;
    };
    CAbsSizeAnchor.prototype.setObject = function (object) {
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_AbsSizeAnchorObject, this.object, object));
        this.object = object;
        if(object){
            object.setParent(this);
        }
    };

    CAbsSizeAnchor.prototype.setParent = function (object) {
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_AbsSizeAnchorParent, this.parent, object));
        this.parent = object;
    };

    CAbsSizeAnchor.prototype.copy = function(oPr){
        var copy = new CRelSizeAnchor();
        copy.setFromTo(this.fromX, this.fromY, this.toX, this.toY);
        if(this.object){
            copy.setObject(this.object.copy(oPr));
        }
        return copy;
    };

    CAbsSizeAnchor.prototype.Refresh_RecalcData = function(drawingDocument){
        if(this.parent && this.parent.Refresh_RecalcData2)
        {
            this.parent.Refresh_RecalcData2();
        }
    };
    CAbsSizeAnchor.prototype.Refresh_RecalcData2 = function(drawingDocument){
        if(this.parent && this.parent.Refresh_RecalcData2)
        {
            this.parent.Refresh_RecalcData2();
        }
    };


    function CalculateSrcRect(parentCropTransform, bounds, oInvertTransformCrop, cropExtX, cropExtY){
        var lt_x_abs = parentCropTransform.TransformPointX(bounds.min_x, bounds.min_y);
        var lt_y_abs = parentCropTransform.TransformPointY(bounds.min_x, bounds.min_y);
        var rb_x_abs = parentCropTransform.TransformPointX(bounds.max_x, bounds.max_y);
        var rb_y_abs = parentCropTransform.TransformPointY(bounds.max_x, bounds.max_y);

        var lt_x_rel = oInvertTransformCrop.TransformPointX(lt_x_abs, lt_y_abs);
        var lt_y_rel = oInvertTransformCrop.TransformPointY(lt_x_abs, lt_y_abs);
        var rb_x_rel = oInvertTransformCrop.TransformPointX(rb_x_abs, rb_y_abs);
        var rb_y_rel = oInvertTransformCrop.TransformPointY(rb_x_abs, rb_y_abs);
        var srcRect = new AscFormat.CSrcRect();
        var _l = (100*lt_x_rel / cropExtX);
        var _t = (100*lt_y_rel / cropExtY);
        var _r = (100*rb_x_rel / cropExtX);
        var _b = (100*rb_y_rel / cropExtY);
        srcRect.l = Math.min(_l, _r);
        srcRect.t = Math.min(_t, _b);
        srcRect.r = Math.max(_l, _r);
        srcRect.b = Math.max(_t, _b);

        return srcRect;
    }


    AscDFH.drawingsChangesMap[AscDFH.historyitem_AbsSizeAnchorFromX]  = function(oClass, value){oClass.fromX =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_AbsSizeAnchorFromY]  = function(oClass, value){oClass.fromY =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_AbsSizeAnchorExtX]    = function(oClass, value){oClass.toX =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_AbsSizeAnchorExtY]    = function(oClass, value){oClass.toY =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_AbsSizeAnchorObject] = function(oClass, value){oClass.object =  value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_AbsSizeAnchorParent] = function(oClass, value){oClass.parent =  value;};

    AscDFH.changesFactory[AscDFH.historyitem_AbsSizeAnchorFromX] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_AbsSizeAnchorFromY] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_AbsSizeAnchorExtX] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_AbsSizeAnchorExtY] = window['AscDFH'].CChangesDrawingsDouble;
    AscDFH.changesFactory[AscDFH.historyitem_AbsSizeAnchorObject] = window['AscDFH'].CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_AbsSizeAnchorParent] = window['AscDFH'].CChangesDrawingsObject;


    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].CGraphicObjectBase    = CGraphicObjectBase;
    window['AscFormat'].CGraphicBounds        = CGraphicBounds;
    window['AscFormat'].checkNormalRotate     = checkNormalRotate;
    window['AscFormat'].normalizeRotate       = normalizeRotate;
    window['AscFormat'].CRelSizeAnchor        = CRelSizeAnchor;
    window['AscFormat'].CAbsSizeAnchor        = CAbsSizeAnchor;
    window['AscFormat'].CalculateSrcRect      = CalculateSrcRect;
    window['AscFormat'].CCopyObjectProperties = CCopyObjectProperties;
    window['AscFormat'].CClientData = CClientData;
    window['AscFormat'].LOCKS_MASKS           = LOCKS_MASKS;
    window['AscFormat'].MACRO_PREFIX = "jsaProject_"
})(window);
