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
var CShape = AscFormat.CShape;

var History = AscCommon.History;



    CShape.prototype.getEditorType = function()
    {
        return 0;
    };

CShape.prototype.Get_Numbering =  function()
{
    return AscWord.DEFAULT_NUMBERING;
};

CShape.prototype.IsUseInDocument = function(){
    if(this.group)
    {
        var aSpTree = this.group.spTree;
        for(var i = 0; i < aSpTree.length; ++i)
        {
            if(aSpTree[i] === this)
            {
                return this.group.IsUseInDocument();
            }
        }
        return false;
    }
    if(this.drawingBase)
    {
        return this.drawingBase.isUseInDocument();
    }
    return false;
};

CShape.prototype.getTextArtPreviewManager = function()
{
    return Asc["editor"].textArtPreviewManager;
};

CShape.prototype.getDrawingObjectsController = function()
{
    var wsViews = Asc["editor"] && Asc["editor"].wb && Asc["editor"].wb.wsViews;
    if(wsViews)
    {
        for(var i = 0; i < wsViews.length; ++i)
        {
            if(wsViews[i] && wsViews[i].model === this.worksheet && wsViews[i].objectRender)
                return wsViews[i].objectRender.controller;
        }
    }
    return null;
};

CShape.prototype.hitInTextRect = function (x, y)
{
    if(!AscFormat.canSelectDrawing(this)) {
        return false;
    }
    var oController = this.getDrawingObjectsController && this.getDrawingObjectsController();
    if(oController && (AscFormat.getTargetTextObject(oController) === this || (oController.curState.startTargetTextObject === this)))
    {
        var content = this.getDocContent && this.getDocContent();
        if ( content && this.invertTransformText)
        {
            return AscFormat.HitToRect(x, y, this.invertTransformText, 0, 0, this.contentWidth, this.contentHeight);
        }
    }
    else
    {
        if (window.AscDisableTextSelection)
            return;


        return this.hitInTextRectWord(x, y);
    }

    return false;
};

function addToDrawings(worksheet, graphic, position, lockByDefault, anchor)
{

    var drawingObjects;
    var wsViews = Asc["editor"].wb && Asc["editor"].wb.wsViews;
    if(wsViews) {
        for(var i = 0; i < wsViews.length; ++i)
        {
            if(wsViews[i] && wsViews[i].model === worksheet)
            {
                drawingObjects = wsViews[i].objectRender;
                break;
            }
        }
    }
    if(!drawingObjects)
    {
        drawingObjects = new AscFormat.DrawingObjects();
    }
    var oldDrawingBase = graphic.drawingBase;

    var drawingObject = drawingObjects.createDrawingObject(anchor);
    drawingObject.graphicObject = graphic;
    graphic.setDrawingBase(drawingObject);
    if(!worksheet)
        return;
    var ret, aObjects = worksheet.Drawings;
    if (AscFormat.isRealNumber(position) && position > -1 && position <= aObjects.length) {
        aObjects.splice(position, 0, drawingObject);
        ret = position;
    }
    else {
        ret = aObjects.length;
        aObjects.push(drawingObject);
    }
    if(oldDrawingBase)
    {
        graphic.setDrawingBaseType(oldDrawingBase.Type);
        graphic.setDrawingBaseEditAs(oldDrawingBase.editAs);
        drawingObject.from.col = oldDrawingBase.from.col;
        drawingObject.from.colOff = oldDrawingBase.from.colOff;
        drawingObject.from.row = oldDrawingBase.from.row;
        drawingObject.from.rowOff = oldDrawingBase.from.rowOff;

        drawingObject.to.col = oldDrawingBase.to.col;
        drawingObject.to.colOff = oldDrawingBase.to.colOff;
        drawingObject.to.row = oldDrawingBase.to.row;
        drawingObject.to.rowOff = oldDrawingBase.to.rowOff;

        drawingObject.Pos.X = oldDrawingBase.Pos.X;
        drawingObject.Pos.Y = oldDrawingBase.Pos.Y;
        drawingObject.ext.cx = oldDrawingBase.ext.cx;
        drawingObject.ext.cy = oldDrawingBase.ext.cy;
    }
    if(graphic.recalcTransform)
    {
        graphic.recalcTransform();
        if(graphic.recalcBounds)
        {
            graphic.recalcBounds();
        }
        graphic.addToRecalculate();
    }

    return ret;
}

function CChangeContentDrawingWorksheet(Class, Pos, isAdd) {
    this.Type = AscDFH.historyitem_AutoShapes_AddToDrawingObjects;
    AscDFH.CChangesBaseContentChange.call(this, Class, Pos, [Class], isAdd);
}

CChangeContentDrawingWorksheet.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangeContentDrawingWorksheet.prototype.constructor = CChangeContentDrawingWorksheet;

    CChangeContentDrawingWorksheet.prototype.Undo = function() {

        if(this.IsAdd()) {
            AscFormat.deleteDrawingBase(this.Class.worksheet.Drawings, this.Class.Get_Id());
        }
        else {
            AscFormat.addToDrawings(this.Class.worksheet, this.Class, this.Pos);
        }
    };
    CChangeContentDrawingWorksheet.prototype.Redo = function() {
        if(this.IsAdd()) {
            AscFormat.addToDrawings(this.Class.worksheet, this.Class, this.Pos);
        }
        else {
            AscFormat.deleteDrawingBase(this.Class.worksheet.Drawings, this.Class.Get_Id());
        }
    };


    CChangeContentDrawingWorksheet.prototype.Load = function(Color){

        if(this.Class.worksheet && this.Class.worksheet.contentChanges) {
            if(this.IsAdd()) {
                var Pos  = this.Class.worksheet.contentChanges.Check(AscCommon.contentchanges_Add, true === this.UseArray && AscFormat.isRealNumber(this.PosArray[0]) ? this.PosArray[0] : this.Pos);
                if(Pos === false){
                    return;
                }
                AscFormat.addToDrawings(this.Class.worksheet, this.Class, Pos);
            }
            else {
                var Pos  = this.Class.worksheet.contentChanges.Check(AscCommon.contentchanges_Remove, true === this.UseArray && AscFormat.isRealNumber(this.PosArray[0]) ? this.PosArray[0] : this.Pos);
                if(Pos === false){
                    return;
                }
                AscFormat.deleteDrawingBase(this.Class.worksheet.Drawings, this.Class.Get_Id());
            }
        }
    };
    CChangeContentDrawingWorksheet.prototype.private_WriteItem = function(Writer, Item)
    {
        Writer.WriteString2(Item.Id);
    };
    CChangeContentDrawingWorksheet.prototype.private_ReadItem = function(Reader)
    {
        let oItem = AscCommon.g_oTableId.Get_ById(Reader.GetString2());
        return oItem;
    };



    function CChangeContentDrawingWorksheetAdd(Class, Pos) {
        CChangeContentDrawingWorksheet.call(this, Class, Pos, true);
        this.Type = AscDFH.historyitem_AutoShapes_AddToDrawingObjects;
    }

    CChangeContentDrawingWorksheetAdd.prototype = Object.create(CChangeContentDrawingWorksheet.prototype);
    CChangeContentDrawingWorksheetAdd.prototype.constructor = CChangeContentDrawingWorksheetAdd;
    CChangeContentDrawingWorksheetAdd.prototype.CreateReverseChange = function()
    {
        return this.private_CreateReverseChange(CChangeContentDrawingWorksheetRemove);
    };

    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_AddToDrawingObjects] = CChangeContentDrawingWorksheetAdd;

    function CChangeContentDrawingWorksheetRemove(Class, Pos) {
        CChangeContentDrawingWorksheet.call(this, Class, Pos, false);
        this.Type = AscDFH.historyitem_AutoShapes_RemoveFromDrawingObjects;
    }

    CChangeContentDrawingWorksheetRemove.prototype = Object.create(CChangeContentDrawingWorksheet.prototype);
    CChangeContentDrawingWorksheetRemove.prototype.constructor = CChangeContentDrawingWorksheetRemove;

    CChangeContentDrawingWorksheetRemove.prototype.CreateReverseChange = function()
    {
        return this.private_CreateReverseChange(CChangeContentDrawingWorksheetAdd);
    };

    AscDFH.changesFactory[AscDFH.historyitem_AutoShapes_RemoveFromDrawingObjects] = CChangeContentDrawingWorksheetRemove;



    CShape.prototype.Clear_ContentChanges = function()
    {
        if(this.worksheet && this.worksheet.contentChanges){
            this.worksheet.contentChanges.Clear();
        }
    };

    CShape.prototype.Add_ContentChanges = function(Changes)
    {
        if(this.worksheet && this.worksheet.contentChanges){
            this.worksheet.contentChanges.Add( Changes );
        }
    };

    CShape.prototype.Refresh_ContentChanges = function()
    {
        if(this.worksheet && this.worksheet.contentChanges){
            this.worksheet.contentChanges.Refresh();
        }
    };


    function editorAddToDrawingObjects(oGraphicObject, pos, type)
    {
        var position = addToDrawings(oGraphicObject.worksheet, oGraphicObject, pos, /*lockByDefault*/undefined, type);
        //var data = {Type: AscDFH.historyitem_AutoShapes_AddToDrawingObjects, Pos: position};
        History.Add(new CChangeContentDrawingWorksheetAdd(oGraphicObject, position));
        if(oGraphicObject.setDrawingBaseType)
        {
            if(oGraphicObject.drawingBase)
            {
                oGraphicObject.setDrawingBaseType && oGraphicObject.setDrawingBaseType(oGraphicObject.drawingBase.Type);
                oGraphicObject.setDrawingBaseEditAs && oGraphicObject.setDrawingBaseEditAs(oGraphicObject.drawingBase.editAs);
            }
            if(AscFormat.isRealNumber(type))
            {
                oGraphicObject.setDrawingBaseType(type);
                if(type === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell)
                {
                    oGraphicObject.setDrawingBaseEditAs(AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
                }
            }
        }
        //oGraphicObject.worksheet.addContentChanges(new AscCommon.CContentChangesElement(AscCommon.contentchanges_Add, position, 1, data));
        oGraphicObject.checkDrawingUniNvPr();
        if(oGraphicObject.signatureLine)
        {
            oGraphicObject.setSignature(oGraphicObject.signatureLine);
        }
        oGraphicObject.checkClientData();
        var oApi = Asc.editor;
        if(oApi && oGraphicObject.signatureLine)
        {
            oApi.sendEvent("asc_onAddSignature", oGraphicObject.signatureLine.id);
        }
    }

function editorDeleteDrawingBase(oGraphicObject, bCheckPlaceholder) {

    let oDrawingBase = oGraphicObject.drawingBase;
    if(oDrawingBase)
    {
        let oFrom = oDrawingBase.from;
        let oTo = oDrawingBase.to;
        let oPos = oDrawingBase.Pos;
        let oExt = oDrawingBase.ext;
        if(oFrom && oTo && oPos && oExt && oGraphicObject.setDrawingBaseType && oGraphicObject.setDrawingBaseCoords)
        {
            if(oDrawingBase.Type === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell)
            {
                oGraphicObject.setDrawingBaseEditAs(AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
            }
            oGraphicObject.setDrawingBaseType(oDrawingBase.Type);

            oGraphicObject.setDrawingBaseCoords(oFrom.col, oFrom.colOff, oFrom.row, oFrom.rowOff, oTo.col, oTo.colOff, oTo.row, oTo.rowOff, oPos.X, oPos.Y, oExt.cx, oExt.cy)
        }
    }
    if(oGraphicObject.signatureLine && oGraphicObject.setSignature)
    {
        oGraphicObject.setSignature(oGraphicObject.signatureLine);
    }
    let position = AscFormat.deleteDrawingBase(oGraphicObject.worksheet.Drawings, oGraphicObject.Get_Id());
    if(AscFormat.isRealNumber(position))
    {
        //var data = {Type: AscDFH.historyitem_AutoShapes_RemoveFromDrawingObjects, Pos: position};
        History.Add(new CChangeContentDrawingWorksheetRemove(oGraphicObject, position));
        //oGraphicObject.worksheet.addContentChanges(new AscCommon.CContentChangesElement(AscCommon.contentchanges_Remove, data.Pos, 1, data));
    }
    if(oGraphicObject.signatureLine && oGraphicObject.setSignature)
    {
        let oApi = Asc.editor;
        if(oApi)
        {
            oApi.sendEvent("asc_onAddSignature", oGraphicObject.signatureLine.id);
        }
        oGraphicObject.setSignature(oGraphicObject.signatureLine);
    }
    return position;
}


function getDrawingObjects_Sp(sp)
{
    var controller = sp.getDrawingObjectsController();
    return controller && controller.drawingObjects;
}

CShape.prototype.setRecalculateInfo = function()
{
    this.recalcInfo =
    {
        recalculateContent:        true,
        recalculateContent2:        true,
        recalculateBrush:          true,
        recalculatePen:            true,
        recalculateTransform:      true,
        recalculateTransformText:  true,
        recalculateBounds:         true,
        recalculateGeometry:       true,
        recalculateStyle:          true,
        recalculateFill:           true,
        recalculateLine:           true,
        recalculateTransparent:    true,
        recalculateTextStyles:     [true, true, true, true, true, true, true, true, true],
        oContentMetrics: null
    };
    this.compiledStyles = [];
    this.lockType = AscCommon.c_oAscLockTypes.kLockTypeNone;
};


    CShape.prototype.checkNeedRecalculate = function(){
        return this.recalcInfo.recalculateTransform === true || this.recalcInfo.recalculateContent === true;
    };

CShape.prototype.recalcContent = function()
{
    this.recalcInfo.recalculateContent = true;
    this.recalcInfo.recalculateContent2 = true;
};

CShape.prototype.getDrawingDocument = function()
{
    if(this.worksheet)
        return this.worksheet.getDrawingDocument();
    var drawingObjects =  getDrawingObjects_Sp(this);
    return drawingObjects && drawingObjects.drawingDocument;
};

CShape.prototype.recalcBrush = function()
{
    this.recalcInfo.recalculateBrush = true;
};

CShape.prototype.recalcPen = function()
{
    this.recalcInfo.recalculatePen = true;
};
CShape.prototype.recalcTransform = function()
{
    this.recalcInfo.recalculateTransform = true;
};
CShape.prototype.recalcTransformText = function()
{
    this.recalcInfo.recalculateTransformText = true;
};
CShape.prototype.recalcBounds = function()
{
    this.recalcInfo.recalculateBounds = true;
};
CShape.prototype.recalcGeometry = function()
{
    this.recalcInfo.recalculateGeometry = true;
};
CShape.prototype.recalcStyle = function()
{
    this.recalcInfo.recalculateStyle = true;
};
CShape.prototype.recalcFill = function()
{
    this.recalcInfo.recalculateFill = true;
};
CShape.prototype.recalcLine = function()
{
    this.recalcInfo.recalculateLine = true;
};
CShape.prototype.recalcTransparent = function()
{
    this.recalcInfo.recalculateTransparent = true;
};
CShape.prototype.recalcTextStyles = function()
{
    this.recalcInfo.recalculateTextStyles =  [true, true, true, true, true, true, true, true, true];
};
CShape.prototype.addToRecalculate = function()
{
    var controller = this.getDrawingObjectsController && this.getDrawingObjectsController();
    if(controller)
    {
        controller.objectsForRecalculate[this.Id] = this;
    }
};

CShape.prototype.handleUpdatePosition = function()
{
    this.recalcTransform();
    this.recalcBounds();
    this.recalcTransformText();
    this.addToRecalculate();
};
CShape.prototype.handleUpdateExtents = function()
{
    this.recalcContent();
    this.recalcGeometry();
    this.recalcBounds();
    this.recalcTransform();
    this.recalcTransformText();
    this.recalcContent();
    this.addToRecalculate();
};
CShape.prototype.handleUpdateRot = function()
{
    this.recalcTransform();
    if(this.txBody && this.txBody.bodyPr && this.txBody.bodyPr.upright)
    {
        this.recalcContent();
    }
    this.recalcTransformText();
    this.recalcBounds();
    this.addToRecalculate();
};
CShape.prototype.handleUpdateFlip = function()
{
    this.recalcTransform();
    this.recalcTransformText();
    this.recalcContent();
    this.addToRecalculate();
};
CShape.prototype.handleUpdateFill = function()
{
    this.recalcBrush();
    this.recalcFill();
    this.recalcTransparent();
    this.addToRecalculate();
};
CShape.prototype.handleUpdateLn = function()
{
    this.recalcPen();
    this.recalcLine();
    this.addToRecalculate();
};
CShape.prototype.handleUpdateGeometry = function()
{
    this.recalcGeometry();
    this.recalcBounds();
    this.recalcContent();
    this.recalcTransformText();
    this.addToRecalculate();
};
CShape.prototype.convertPixToMM = function(pix)
{
    var drawingObjects =  getDrawingObjects_Sp(this);
    var _ret = drawingObjects ? drawingObjects.convertMetric(AscCommon.AscBrowser.convertToRetinaValue(pix, true), 0, 3) : 0;
    return _ret;
};
CShape.prototype.getCanvasContext = function()
{
    var drawingObjects =  getDrawingObjects_Sp(this);
    return drawingObjects ? drawingObjects.getCanvasContext() : null;
};
CShape.prototype.getCompiledStyle = function()
{
    return this.style;
};
CShape.prototype.getHierarchy = function()
{
    return [];
};
CShape.prototype.getParentObjects = function ()
{
    return { slide: null, layout: null, master: null, theme: window["Asc"]["editor"].wbModel.theme};
};

CShape.prototype.recalcText = function()
{
    this.recalcInfo.recalculateContent = true;
    this.recalcInfo.recalculateContent2 = true;
    this.recalcInfo.recalculateTransformText = true;
};

/**
 * @memberof CShape
 */
CShape.prototype.recalculate = function ()
{
    if(this.bDeleted)
        return;
    AscFormat.ExecuteNoHistory(function(){
        var bRecalcShadow = this.recalcInfo.recalculateBrush ||
            this.recalcInfo.recalculatePen ||
            this.recalcInfo.recalculateTransform ||
            this.recalcInfo.recalculateGeometry ||
            this.recalcInfo.recalculateBounds;
        if (this.recalcInfo.recalculateBrush) {
            this.recalculateBrush();
            this.recalcInfo.recalculateBrush = false;
        }

        if (this.recalcInfo.recalculatePen) {
            this.recalculatePen();
            this.recalcInfo.recalculatePen = false;
        }
        if (this.recalcInfo.recalculateTransform) {
            this.recalculateTransform();
            this.recalculateSnapArrays();
            this.recalcInfo.recalculateTransform = false;
        }

        if (this.recalcInfo.recalculateGeometry) {
            this.recalculateGeometry();
            this.recalcInfo.recalculateGeometry = false;
        }

        if (this.recalcInfo.recalculateContent) {
            this.recalcInfo.oContentMetrics = this.recalculateContent();
            this.recalcInfo.recalculateContent = false;
        }
        if (this.recalcInfo.recalculateContent2) {
            this.recalculateContent2();
            this.recalcInfo.recalculateContent2 = false;
        }

        if (this.recalcInfo.recalculateTransformText) {
            this.recalculateTransformText();
            this.recalcInfo.recalculateTransformText = false;
        }

        if(this.recalcInfo.recalculateBounds)
        {
            this.recalculateBounds();
            this.recalcInfo.recalculateBounds = false;
        }
        if(bRecalcShadow)
        {
            this.recalculateShdw();
        }
        this.clearCropObject();

    }, this, []);
};

CShape.prototype.recalculateBounds = function()
{
    var boundsChecker = new  AscFormat.CSlideBoundsChecker();
    this.draw(boundsChecker);
    boundsChecker.CorrectBounds();

    this.bounds.x = boundsChecker.Bounds.min_x;
    this.bounds.y = boundsChecker.Bounds.min_y;
    this.bounds.l = boundsChecker.Bounds.min_x;
    this.bounds.t = boundsChecker.Bounds.min_y;
    this.bounds.r = boundsChecker.Bounds.max_x;
    this.bounds.b = boundsChecker.Bounds.max_y;
    this.bounds.w = boundsChecker.Bounds.max_x - boundsChecker.Bounds.min_x;
    this.bounds.h = boundsChecker.Bounds.max_y - boundsChecker.Bounds.min_y;
    if(this.drawingBase && !this.group)
    {
        this.drawingBase.checkBoundsFromTo();
    }
};
CShape.prototype.recalculateContent = function()
{
    var content = this.getDocContent();
    if(content){
        var body_pr = this.getBodyPr();

        var oRecalcObj = this.recalculateDocContent(content, body_pr);
        this.contentWidth = oRecalcObj.w;
        this.contentHeight = oRecalcObj.contentH;
        if(this.txBody)
        {
            this.txBody.contentWidth = this.contentWidth;
            this.txBody.contentHeight = this.contentHeight;
        }
        this.recalcInfo.oContentMetrics = oRecalcObj;
        if(this.recalcInfo.recalcTitle)
        {
            this.recalcInfo.bRecalculatedTitle = true;
            this.recalcInfo.recalcTitle = null;


            var oTextWarpContent = this.checkTextWarp(content, body_pr, oRecalcObj.textRectW + oRecalcObj.correctW, oRecalcObj.textRectH + oRecalcObj.correctH, true, false);
            this.txWarpStructParamarks = oTextWarpContent.oTxWarpStructParamarksNoTransform;
            this.txWarpStruct = oTextWarpContent.oTxWarpStructNoTransform;

            this.txWarpStructParamarksNoTransform = oTextWarpContent.oTxWarpStructParamarksNoTransform;
            this.txWarpStructNoTransform = oTextWarpContent.oTxWarpStructNoTransform;
        }
        else
        {
            var oTextWarpContent = this.checkTextWarp(content, body_pr, oRecalcObj.textRectW + oRecalcObj.correctW, oRecalcObj.textRectH + oRecalcObj.correctH, true, true);
            this.txWarpStructParamarks = oTextWarpContent.oTxWarpStructParamarks;
            this.txWarpStruct = oTextWarpContent.oTxWarpStruct;

            this.txWarpStructParamarksNoTransform = oTextWarpContent.oTxWarpStructParamarksNoTransform;
            this.txWarpStructNoTransform = oTextWarpContent.oTxWarpStructNoTransform;
        }
        return oRecalcObj;
    }
    else{
        this.txWarpStructParamarks = null;
        this.txWarpStruct = null;

        this.txWarpStructParamarksNoTransform = null;
        this.txWarpStructNoTransform = null;

        this.recalcInfo.warpGeometry = null;
    }
    return null;
};

CShape.prototype.Get_ColorMap = function()
{
    return AscFormat.GetDefaultColorMap();
};

CShape.prototype.getStyles = function(index)
{
    return this.Get_Styles(index);
};

CShape.prototype.Get_Worksheet = function()
{
    return this.worksheet;
};
CShape.prototype.checkDrawingPartWithHistory = function () {
	if (this.isObjectInSmartArt()) {
		return this.group.group.checkDrawingPartWithHistory(this);
	}
};



    CShape.prototype.Set_CurrentElement = function()
    {
        var drawing_objects = this.getDrawingObjectsController();
        if(drawing_objects)
        {
            this.SetControllerTextSelection(drawing_objects, 0);
        }
    };

AscFormat.CTextBody.prototype.Get_Worksheet = function()
{
    return this.parent && this.parent.Get_Worksheet && this.parent.Get_Worksheet();
};
AscFormat.CTextBody.prototype.getDrawingDocument = function()
{
    var oCellApi = window["Asc"] && window["Asc"]["editor"];
    if (oCellApi && oCellApi.wbModel)
    {
        return oCellApi.wbModel.DrawingDocument;
    }
    return null;
};
    AscFormat.CTextBody.prototype.checkCurrentPlaceholder = function()
    {
        var oCurController;
        var oApi = Asc.editor;
        if(oApi)
        {
            var ws = oApi.wb.getWorksheet();
            var oParaPr;
            if (ws && ws.objectRender && ws.objectRender.controller) {
                oCurController = ws.objectRender.controller;
            }
        }
        if(oCurController)
        {
            return oCurController.getTargetDocContent() === this.content;
        }
        return false;
    };

    //------------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].addToDrawings = addToDrawings;
    window['AscFormat'].editorDeleteDrawingBase = editorDeleteDrawingBase;
    window['AscFormat'].editorAddToDrawingObjects = editorAddToDrawingObjects;
})(window);
