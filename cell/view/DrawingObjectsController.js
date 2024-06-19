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
(
/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

    // Import
    var DrawingObjectsController = AscFormat.DrawingObjectsController;

    var History = AscCommon.History;

if(window.editor === "undefined" && window["Asc"]["editor"])
{
    window.editor = window["Asc"]["editor"];
	window.editor;
}

// ToDo убрать это отсюда!!!
AscCommon.CContentChangesElement.prototype.Refresh_BinaryData = function()
{
    if(this.m_aPositions.length > 0){
        this.m_pData.Pos = this.m_aPositions[0];
    }
    this.m_pData.UseArray = true;
    this.m_pData.PosArray = this.m_aPositions;
	if(editor && editor.isPresentationEditor)
	{
		var Binary_Writer = History.BinaryWriter;
		var Binary_Pos = Binary_Writer.GetCurPosition();

        this.m_pData.Data.UseArray = true;
        this.m_pData.Data.PosArray = this.m_aPositions;

        if ((Asc.editor || editor).binaryChanges) {
            Binary_Writer.WriteWithLen(this, function() {
                Binary_Writer.WriteString2(this.m_pData.Class.Get_Id());
                Binary_Writer.WriteLong(this.m_pData.Data.Type);
                this.m_pData.Data.WriteToBinary(Binary_Writer);
            });
        } else {
            Binary_Writer.WriteString2(this.m_pData.Class.Get_Id());
            Binary_Writer.WriteLong(this.m_pData.Data.Type);
            this.m_pData.Data.WriteToBinary(Binary_Writer);
        }
        var Binary_Len = Binary_Writer.GetCurPosition() - Binary_Pos;

		this.m_pData.Binary.Pos = Binary_Pos;
		this.m_pData.Binary.Len = Binary_Len;
	}
};

function CheckIdSatetShapeAdd(state)
{
    return !(state instanceof AscFormat.NullState);
}
DrawingObjectsController.prototype.getTheme = function()
{
    return window["Asc"]["editor"].wbModel.theme;
};

DrawingObjectsController.prototype.getDrawingArray = function()
{
    var ret = [];
    var drawing_bases = this.drawingObjects.getDrawingObjects();
    for(var i = 0; i < drawing_bases.length; ++i)
    {
        ret.push(drawing_bases[i].graphicObject);
    }
    return ret;
};

DrawingObjectsController.prototype.setTableProps = function(props)
{
    let by_type = this.getSelectedObjectsByTypes();
    if(by_type.tables.length === 1)
    {
        let sCaption = props.TableCaption;
        let sDescription = props.TableDescription;
        let sName = props.TableName;
        let dRowHeight = props.RowHeight;
        let oTable = by_type.tables[0];
        oTable.setTitle(sCaption);
        oTable.setDescription(sDescription);
        oTable.setName(sName);
        props.TableCaption = undefined;
        props.TableDescription = undefined;
        var bIgnoreHeight = false;
        if(AscFormat.isRealNumber(props.RowHeight))
        {
            if(AscFormat.fApproxEqual(props.RowHeight, 0.0))
            {
                props.RowHeight = 1.0;
            }
            bIgnoreHeight = false;
        }
        var target_text_object = AscFormat.getTargetTextObject(this);
        if(target_text_object === oTable)
        {
            oTable.graphicObject.Set_Props(props);
        }
        else
        {
            oTable.graphicObject.SelectAll();
            oTable.graphicObject.Set_Props(props);
            oTable.graphicObject.RemoveSelection();
        }
        props.TableCaption = sCaption;
        props.TableDescription = sDescription;
        props.RowHeight = dRowHeight;
        if(!oTable.setFrameTransform(props))
        {
            editor.WordControl.m_oLogicDocument.Check_GraphicFrameRowHeight(oTable, bIgnoreHeight);
        }
        
    }
};

DrawingObjectsController.prototype.RefreshAfterChangeColorScheme = function()
{
    var drawings = this.getDrawingArray();
    for(var i = 0; i < drawings.length; ++i)
    {
        if(drawings[i])
        {
            drawings[i].handleUpdateFill();
            drawings[i].handleUpdateLn();
            drawings[i].addToRecalculate();
        }
    }
};
DrawingObjectsController.prototype.updateOverlay = function()
{
    this.drawingObjects.OnUpdateOverlay();
};
DrawingObjectsController.prototype.updatePlaceholders = function ()
{
    const oWS = Asc.editor && Asc.editor.wbModel && Asc.editor.wbModel.getActiveWs();
    if (oWS)
    {
        const arrRet = [];
        const arrDrawingObjects = oWS.Drawings;
        for (let i = 0; i < arrDrawingObjects.length; i += 1)
        {
            const oGraphicObject = arrDrawingObjects[i] && arrDrawingObjects[i].graphicObject;
            if (oGraphicObject)
            {
                oGraphicObject.createPlaceholderControl(arrRet);
            }
        }
        const oDrawingDocument = this.getDrawingDocument();
        if (oDrawingDocument && oDrawingDocument.placeholders)
        {
            oDrawingDocument.placeholders.update(arrRet);
        }
    }
};
DrawingObjectsController.prototype.recalculate = function(bAll, Point, bCheckPoint)
{
    if(bCheckPoint !== false)
    {
        this.objectsForRecalculate = {};
        History.Get_RecalcData(Point);//Только для таблиц
    }
    this.recalculate2(bAll);
};

DrawingObjectsController.prototype.recalculate2 = function(bAll)
{
    if(bAll)
    {
        var drawings = this.getDrawingObjects();
        for(var i = 0; i < drawings.length; ++i)
        {
            if(drawings[i].recalcText)
            {
                drawings[i].recalcText();
            }
            drawings[i].recalculate();
        }
    }
    else
    {
        for(var key in this.objectsForRecalculate)
        {
            this.objectsForRecalculate[key].recalculate();
        }
    }
    this.objectsForRecalculate = {};
    this.updatePlaceholders();
};


DrawingObjectsController.prototype.updateRecalcObjects = function()
{};
DrawingObjectsController.prototype.getTheme = function()
{
    return window["Asc"]["editor"].wbModel.theme;
};

DrawingObjectsController.prototype.startRecalculate = function(bCheckPoint)
{
    this.recalculate(undefined, undefined, bCheckPoint);
    this.drawingObjects.showDrawingObjects();
    //this.updateSelectionState();
};

DrawingObjectsController.prototype.getDrawingObjects = function()
{
    //TODO: переделать эту функцию. Нужно где-то паралельно с массивом DrawingBas'ов хранить масси graphicObject'ов.
    var ret = [];
    var drawing_bases = this.drawingObjects.getDrawingObjects();
    for(var i = 0; i < drawing_bases.length; ++i)
    {
        ret.push(drawing_bases[i].graphicObject);
    }
    return ret;
};

DrawingObjectsController.prototype.checkSelectedObjectsAndFireCallback = function(callback, args)
{
    if(!this.canEdit()){
        return;
    }
    var oApi = Asc.editor;
    if(oApi && oApi.collaborativeEditing && oApi.collaborativeEditing.getGlobalLock()){
        return;
    }
    var selection_state = this.getSelectionState();
    var aId = [];
    for(var i = 0; i < this.selectedObjects.length; ++i)
    {
        aId.push(this.selectedObjects[i].Get_Id());
    }
    var _this = this;
    var callback2 = function(bLock, bSync)
    {
        if(bLock)
        {
            if(bSync !== true)
            {
                _this.setSelectionState(selection_state);
            }
            callback.apply(_this, args);
        }
    };
    oApi.checkObjectsLock(aId, callback2);
};
DrawingObjectsController.prototype.onMouseDown = function(e, x, y)
{
    e.ShiftKey = e.shiftKey;
    e.CtrlKey = e.metaKey || e.ctrlKey;
    e.Button = e.button;
    e.Type = AscCommon.g_mouse_event_type_down;
	e.IsLocked = e.isLocked;
    this.checkInkState();
    var ret = this.curState.onMouseDown(e, x, y, 0);
    if(e.ClickCount < 2)
    {
        if(this.drawingObjects && this.drawingObjects.getWorksheet){
            var ws = this.drawingObjects.getWorksheet();
            if(Asc.editor.wb.getWorksheet() !== ws){
                return ret;
            }
        }
        this.updateOverlay();
        this.updateSelectionState();
    }
    return ret;
};

DrawingObjectsController.prototype.OnMouseDown = DrawingObjectsController.prototype.onMouseDown;

DrawingObjectsController.prototype.onMouseMove = function(e, x, y)
{
    e.ShiftKey = e.shiftKey;
    e.CtrlKey = e.metaKey || e.ctrlKey;
    e.Button = e.button;
    e.Type = AscCommon.g_mouse_event_type_move;
	e.IsLocked = e.isLocked;
    this.curState.onMouseMove(e, x, y, 0);
};
DrawingObjectsController.prototype.OnMouseMove = DrawingObjectsController.prototype.onMouseMove;


DrawingObjectsController.prototype.onMouseUp = function(e, x, y)
{
    e.ShiftKey = e.shiftKey;
    e.CtrlKey = e.metaKey || e.ctrlKey;
    e.Button = e.button;
    e.Type = AscCommon.g_mouse_event_type_up;
    this.curState.onMouseUp(e, x, y, 0);
};
DrawingObjectsController.prototype.OnMouseUp = DrawingObjectsController.prototype.onMouseUp;

DrawingObjectsController.prototype.createGroup = function()
{
    var group = this.getGroup();
    if(group)
    {
        var group_array = this.getArrayForGrouping();
        for(var i = group_array.length - 1; i > -1; --i)
        {
            group_array[i].deleteDrawingBase();
        }
        this.resetSelection();
        this.drawingObjects.getWorksheetModel && group.setWorksheet(this.drawingObjects.getWorksheetModel());
        group.setDrawingObjects(this.drawingObjects);
        if(this.drawingObjects && this.drawingObjects.cSld)
        {
            group.setParent(this.drawingObjects);
        }
        group.addToDrawingObjects();
        group.checkDrawingBaseCoords();
        this.selectObject(group, 0);
        group.addToRecalculate();
        this.startRecalculate();
    }
		return group;
};
DrawingObjectsController.prototype.handleChartDoubleClick = function()
{
    var drawingObjects = this.drawingObjects;
    var oThis = this;
    this.checkSelectedObjectsAndFireCallback(function(){
        oThis.clearTrackObjects();
        oThis.clearPreTrackObjects();
        oThis.changeCurrentState(new AscFormat.NullState(this));
        drawingObjects.showChartSettings();
    }, []);
};


DrawingObjectsController.prototype.handleOleObjectDoubleClick = function(drawing, oleObject, e, x, y, pageIndex)
{
    let oThis = this;
    var fCallback = function(){
        if(oleObject.m_oMathObject) {
            Asc.editor.sendEvent("asc_onConvertEquationToMath", oleObject);
        } else if (oleObject.canEditTableOleObject()) {
            Asc.editor.asc_doubleClickOnTableOleObject(oleObject);
        } else {
            oleObject.runPlugin();
        }
        oThis.clearTrackObjects();
        oThis.clearPreTrackObjects();
        oThis.changeCurrentState(new AscFormat.NullState(oThis));
        oThis.onMouseUp(e, x, y);
    };
    if(!this.canEdit()){
        fCallback();
        return;
    }
    this.checkSelectedObjectsAndFireCallback(fCallback, []);
};

DrawingObjectsController.prototype.addChartDrawingObject = function(options)
{
    History.Create_NewPoint();
    var chart = this.getChartSpace(options, false);
    if(chart)
    {
        chart.setWorksheet(this.drawingObjects.getWorksheetModel());
        chart.setStyle(2);
        chart.setBDeleted(false);
        this.resetSelection();
        var w, h;
        if(AscCommon.isRealObject(options) && AscFormat.isRealNumber(options.width) && AscFormat.isRealNumber(options.height))
        {
            w = this.drawingObjects.convertMetric(options.width, 0, 3);
            h = this.drawingObjects.convertMetric(options.height, 0, 3);
        }
        else
        {
            w = this.drawingObjects.convertMetric(AscCommon.AscBrowser.convertToRetinaValue(AscCommon.c_oAscChartDefines.defaultChartWidth, true), 0, 3);
            h = this.drawingObjects.convertMetric(AscCommon.AscBrowser.convertToRetinaValue(AscCommon.c_oAscChartDefines.defaultChartHeight, true), 0, 3);
        }

        var chartLeft, chartTop;
        if(options && AscFormat.isRealNumber(options.left) && options.left >= 0 && AscFormat.isRealNumber(options.top) && options.top >= 0)
        {
            chartLeft = this.drawingObjects.convertMetric(options.left, 0, 3);
            chartTop = this.drawingObjects.convertMetric(options.top, 0, 3);
        }
        else
        {
            chartLeft =  -this.drawingObjects.convertMetric(this.drawingObjects.getScrollOffset().getX(), 0, 3) + (this.drawingObjects.convertMetric(this.drawingObjects.getContextWidth(), 0, 3) - w) / 2;
            if(chartLeft < 0)
            {
                chartLeft = 0;
            }
            chartTop =  -this.drawingObjects.convertMetric(this.drawingObjects.getScrollOffset().getY(), 0, 3) + (this.drawingObjects.convertMetric(this.drawingObjects.getContextHeight(), 0, 3) - h) / 2;
            if(chartTop < 0)
            {
                chartTop = 0;
            }
        }


        chart.setSpPr(new AscFormat.CSpPr());
        chart.spPr.setParent(chart);
        chart.spPr.setXfrm(new AscFormat.CXfrm());
        chart.spPr.xfrm.setParent(chart.spPr);
        chart.spPr.xfrm.setOffX(chartLeft);
        chart.spPr.xfrm.setOffY(chartTop);
        chart.spPr.xfrm.setExtX(w);
        chart.spPr.xfrm.setExtY(h);

        chart.setDrawingObjects(this.drawingObjects);
        chart.setWorksheet(this.drawingObjects.getWorksheetModel());
        chart.addToDrawingObjects();
        this.resetSelection();
        this.selectObject(chart, 0);
        if(options)
        {
            var old_range = options.getRange();
            options.putRange(null);
            options.putStyle(null);
            options.removeAllAxesProps();
            //options.putShowMarker(null);
            this.editChartCallback(options);
            options.putStyle(1);
           // options.bCreate = true;
            this.editChartCallback(options);
            options.putRange(old_range);
        }
        chart.addToRecalculate();
        chart.checkDrawingBaseCoords();
        this.startRecalculate();
        this.drawingObjects.sendGraphicObjectProps();
    }
	return chart;
};

DrawingObjectsController.prototype.isPointInDrawingObjects = function(x, y, e)
{
    this.handleEventMode = AscFormat.HANDLE_EVENT_MODE_CURSOR;
    this.checkInkState();
    var ret = this.curState.onMouseDown(e || {}, x, y, 0);
    this.handleEventMode = AscFormat.HANDLE_EVENT_MODE_HANDLE;
    return ret;
};

DrawingObjectsController.prototype.handleDoubleClickOnChart = function(chart)
{
    this.changeCurrentState(new AscFormat.NullState());
};

DrawingObjectsController.prototype.addImageFromParams = function(rasterImageId, x, y, extX, extY)
{
    var image = this.createImage(rasterImageId, x, y, extX, extY);
    image.setWorksheet(this.drawingObjects.getWorksheetModel());
    image.setDrawingObjects(this.drawingObjects);
    image.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorOneCell);
    image.checkDrawingBaseCoords();
    this.selectObject(image, 0);
    image.addToRecalculate();
};
DrawingObjectsController.prototype.addImage = function(sImageUrl, nPixW, nPixH, videoUrl, audioUrl)
{
    let options = {
        cell: this.drawingObjects.getWorksheetModel().selectionRange.activeCell,
        width: nPixW,
        height: nPixH
    }
    let _image = {
        src:  sImageUrl,
        Image: {
          src: sImageUrl
        }
    };
    this.drawingObjects.addImageObjectCallback(_image, options);
    this.startRecalculate();
    this.drawingObjects.getWorksheet().setSelectionShape(true);
};

DrawingObjectsController.prototype.addOleObjectFromParams = function(fPosX, fPosY, fWidth, fHeight, nWidthPix, nHeightPix, sLocalUrl, sData, sApplicationId, bSelect, arrImagesForAddToHistory){
    var oOleObject = this.createOleObject(sData, sApplicationId, sLocalUrl, fPosX, fPosY, fWidth, fHeight, nWidthPix, nHeightPix, arrImagesForAddToHistory);
    this.resetSelection();
    oOleObject.setWorksheet(this.drawingObjects.getWorksheetModel());
    oOleObject.setDrawingObjects(this.drawingObjects);
    oOleObject.addToDrawingObjects();
    oOleObject.checkDrawingBaseCoords();
    if(bSelect !== false) {
        this.selectObject(oOleObject, 0);
    }
    oOleObject.addToRecalculate();
    this.startRecalculate();
};

DrawingObjectsController.prototype.editOleObjectFromParams = function(oOleObject, sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory){
    oOleObject.editExternal(sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory);
    this.startRecalculate();
};


DrawingObjectsController.prototype.addTextArtFromParams = function(nStyle, dRectX, dRectY, dRectW, dRectH, wsmodel)
{
    History.Create_NewPoint();
    var oTextArt = this.createTextArt(nStyle, false, wsmodel);
    this.resetSelection();
    oTextArt.setWorksheet(this.drawingObjects.getWorksheetModel());
    oTextArt.setDrawingObjects(this.drawingObjects);
    oTextArt.addToDrawingObjects();
    oTextArt.checkExtentsByDocContent();
    var dNewPoX = dRectX + (dRectW - oTextArt.spPr.xfrm.extX) / 2;
    if(dNewPoX < 0)
        dNewPoX = 0;
    var dNewPoY = dRectY + (dRectH - oTextArt.spPr.xfrm.extY) / 2;
    if(dNewPoY < 0)
        dNewPoY = 0;
    oTextArt.spPr.xfrm.setOffX(dNewPoX);
    oTextArt.spPr.xfrm.setOffY(dNewPoY);

    oTextArt.checkDrawingBaseCoords();
    this.selectObject(oTextArt, 0);
    var oContent = oTextArt.getDocContent();
    this.selection.textSelection = oTextArt;
    oContent.SelectAll();
    oTextArt.addToRecalculate();
    this.startRecalculate();
};


DrawingObjectsController.prototype.getDrawingDocument = function()
{
    return this.drawingObjects.drawingDocument;
};
DrawingObjectsController.prototype.convertPixToMM = function(pix)
{
    var _ret = this.drawingObjects ? this.drawingObjects.convertMetric(pix, 0, 3) : 0;
    _ret *= AscCommon.AscBrowser.retinaPixelRatio;
    return _ret;
};

DrawingObjectsController.prototype.setParagraphNumbering = function(Bullet)
{
    this.applyDocContentFunction(CDocumentContent.prototype.Set_ParagraphPresentationNumbering, [Bullet], CTable.prototype.Set_ParagraphPresentationNumbering);
};

DrawingObjectsController.prototype.setParagraphIndent = function(Indent)
{
    if(AscCommon.isRealObject(Indent) && AscFormat.isRealNumber(Indent.Left) && Indent.Left < 0)
    {
        Indent.Left = 0;
    }
    this.applyDocContentFunction(CDocumentContent.prototype.SetParagraphIndent, [Indent], CTable.prototype.SetParagraphIndent);
};

DrawingObjectsController.prototype.paragraphIncDecIndent = function(bIncrease)
{
    this.applyDocContentFunction(CDocumentContent.prototype.Increase_ParagraphLevel, [bIncrease], CTable.prototype.Increase_ParagraphLevel);
};

DrawingObjectsController.prototype.canIncreaseParagraphLevel = function(bIncrease)
{
    let oDocContent = this.getTargetDocContent();
    if(oDocContent)
    {
        let oTextObject = AscFormat.getTargetTextObject(this);
        if(oTextObject && oTextObject.getObjectType() === AscDFH.historyitem_type_Shape)
        {
            if(oTextObject.isPlaceholder())
            {
                let nPhType = oTextObject.getPlaceholderType();
                if(nPhType === AscFormat.phType_title || nPhType === AscFormat.phType_ctrTitle)
                {
                    return false;
                }
            }
            return oDocContent.Can_IncreaseParagraphLevel(bIncrease);
        }
    }
    return false;
};


    DrawingObjectsController.prototype.checkMobileCursorPosition = function () {
        if(!this.drawingObjects){
            return;
        }
        var oWorksheet = this.drawingObjects.getWorksheet();
        if(!oWorksheet){
            return;
        }
        if(window["Asc"]["editor"].isMobileVersion){
            var oTargetDocContent = this.getTargetDocContent(false, false);
            if(oTargetDocContent){
                var oPos = oTargetDocContent.GetCursorPosXY();
                var oParentTextTransform = oTargetDocContent.Get_ParentTextTransform();
                var _x, _y;
                if(oParentTextTransform){
                    _x = oParentTextTransform.TransformPointX(oPos.X, oPos.Y);
                    _y = oParentTextTransform.TransformPointY(oPos.X, oPos.Y);
                }
                else{
                    _x = oPos.X;
                    _y = oPos.Y;
                }
                _x = this.drawingObjects.convertMetric(_x, 3, 0);
                _y = this.drawingObjects.convertMetric(_y, 3, 0);
                var oCell = oWorksheet.findCellByXY(_x, _y, true, false, false);
                if(oCell && oCell.col !== null && oCell.row !== null){
                    var oRange = new Asc.Range(oCell.col, oCell.row, oCell.col, oCell.row, false);
                    var oVisibleRange = oWorksheet.getVisibleRange();
                    if(!oRange.isIntersect(oVisibleRange)){
                        oWorksheet._scrollToRange(oRange);
                    }
                }
            }
        }
    };

DrawingObjectsController.prototype.onKeyPress = function(e)
{
    if (!this.canEdit())
        return false;
    if(e.CtrlKey || e.AltKey)
        return false;

    let Code;
    if (null != e.Which)
        Code = e.Which;
    else if (e.KeyCode)
        Code = e.KeyCode;
    else
        Code = 0;//special char

    let bRetValue = false;
    if ( Code >= 0x20 )
    {
        return this.enterText(Code);
    }

    return bRetValue;
};
    DrawingObjectsController.prototype.enterText = function (codePoints)
    {
        if (!this.canEdit())
            return false;
        if(this.checkSelectedObjectsProtectionText())
        {
            return true;
        }
        let fCallback = function()
        {
            let oItem;
            let Code;
            if(Array.isArray(codePoints))
            {
                for(let nIdx = 0; nIdx < codePoints.length; ++nIdx)
                {
                    Code = codePoints[nIdx];
                    if(AscCommon.IsSpace(Code))
                    {
                        oItem = new AscWord.CRunSpace(Code);
                    }
                    else
                    {
                        oItem = new AscWord.CRunText(Code);
                    }
                    this.paragraphAdd(oItem, true);
                }
            }
            else
            {
                Code = codePoints;
                if(AscCommon.IsSpace(Code))
                {
                    oItem = new AscWord.CRunSpace(Code);
                }
                else
                {
                    oItem = new AscWord.CRunText(Code);
                }
                this.paragraphAdd(oItem, true);
            }
            this.checkMobileCursorPosition();
            this.recalculateCurPos(true, true);
        };
        this.checkSelectedObjectsAndCallback(fCallback, [], false, AscDFH.historydescription_Spreadsheet_ParagraphAdd, undefined, window["Asc"]["editor"].collaborativeEditing.getFast());
        return AscCommon.isRealObject(this.getTargetDocContent());
    };

    DrawingObjectsController.prototype.checkSlicerCopies = function (aCopies)
    {
        var i;
        var aSlicers = [];
        var aSlicerViewNames = [];
        var oDrawing, oSlicer, sSlicerViewName;
        var oWSView;
        var oWB = Asc["editor"] && Asc["editor"].wbModel;
        var oControllerParent = this.drawingObjects;
        var oThis = this;
        if(oControllerParent && oControllerParent.getWorksheet)
        {
            oWSView = oControllerParent.getWorksheet();
        }
        if(!oWB || !oWSView)
        {
            return;
        }
        var aSlicerView = [];

        for(i = 0; i < aCopies.length; ++i)
        {
            aCopies[i].getAllSlicerViews(aSlicerView);
        }
        for(i = 0; i < aSlicerView.length; ++i)
        {
            oDrawing = aSlicerView[i];
            sSlicerViewName = oDrawing.getName();
            oSlicer = oWB.getSlicerByName(sSlicerViewName);
            if(oSlicer)
            {
                aSlicers.push(oSlicer);
                aSlicerViewNames.push(sSlicerViewName);
            }
        }
        if(aSlicers.length > 0)
        {
            History.StartTransaction();
            oWSView.tryPasteSlicers(aSlicers, function(bSuccess, aSlicerNames)
            {
                if(!bSuccess || aSlicerNames.length !== aSlicerViewNames.length)
                {
                    History.EndTransaction();
                    History.Undo();
                    return;
                }
                History.EndTransaction();
                var i, j;
                var oDrawing;
                var sOldSlicerName, sNewSlicerName;
                for(i = 0; i < aSlicerNames.length; ++i)
                {
                    sOldSlicerName = aSlicerViewNames[i];
                    sNewSlicerName = aSlicerNames[i];
                    for(j = 0; j < aSlicerView.length; ++j)
                    {
                        oDrawing = aSlicerView[j];
                        if(oDrawing.getName() === sOldSlicerName)
                        {
                            oDrawing.setName(sNewSlicerName);
                            oDrawing.onDataUpdate();
                            break;
                        }
                    }
                }
                oThis.startRecalculate();
                oThis.drawingObjects.sendGraphicObjectProps();
            });
        }
    };

    DrawingObjectsController.prototype.checkGraphicObjectPosition = function (x, y, w, h)
    {
        if(this.drawingObjects.checkGraphicObjectPosition)
        {
            return this.drawingObjects.checkGraphicObjectPosition(x, y, w, h);
        }
        return {x: 0, y: 0};
    };
//------------------------------------------------------------export---------------------------------------------------
window['AscCommonExcel'] = window['AscCommonExcel'] || {};
window['AscCommonExcel'].CheckIdSatetShapeAdd = CheckIdSatetShapeAdd;
})(window);
