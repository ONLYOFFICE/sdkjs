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
var HANDLE_EVENT_MODE_HANDLE = AscFormat.HANDLE_EVENT_MODE_HANDLE;
var HANDLE_EVENT_MODE_CURSOR = AscFormat.HANDLE_EVENT_MODE_CURSOR;

var isRealObject = AscCommon.isRealObject;
    var History = AscCommon.History;

var MOVE_DELTA = 1/100000;
var SNAP_DISTANCE = 1.27;



function StartAddNewShape(drawingObjects, preset, nPlaceholderType, bVertical)
{
    this.drawingObjects = drawingObjects;
    this.preset = preset;

    this.bStart = false;
    this.bMoved = false;//отошли ли мы от начальной точки

    this.startX = null;
    this.startY = null;

    this.oldConnector = null;
    this.placeholderType = nPlaceholderType;
    this.bVertical = bVertical;
}

StartAddNewShape.prototype =
{
    onMouseDown: function(e, x, y)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};

        let dStartX = x;
        let dStartY = y;
        let oNearestPos = this.drawingObjects.getSnapNearestPos(x, y);
        if(oNearestPos)
        {
            dStartX = oNearestPos.x;
            dStartY = oNearestPos.y;
        }
        this.startX = dStartX;
        this.startY = dStartY;
        this.drawingObjects.arrPreTrackObjects.length = 0;
        var layout = null, master = null, slide = null;
        if(this.drawingObjects.drawingObjects && this.drawingObjects.drawingObjects.cSld)
        {
            if(this.drawingObjects.drawingObjects.getParentObjects)
            {
                var oParentObjects = this.drawingObjects.drawingObjects.getParentObjects();
                if(isRealObject(oParentObjects))
                {
                    layout = oParentObjects.layout;
                    master = oParentObjects.master;
                    slide = oParentObjects.slide;
                }
            }
        }
        this.drawingObjects.arrPreTrackObjects.push(new AscFormat.NewShapeTrack(this.preset, dStartX, dStartY, this.drawingObjects.getTheme(), master, layout, slide, 0, this.drawingObjects, this.placeholderType, this.bVertical));
        this.bStart = true;
        this.drawingObjects.swapTrackObjects();
    },

    onMouseMove: function(e, x, y)
    {
        if(this.bStart && e.IsLocked)
        {
            if(!this.bMoved && (Math.abs(this.startX - x) > MOVE_DELTA || Math.abs(this.startY - y) > MOVE_DELTA ))
                this.bMoved = true;
            let oNearestPos = this.drawingObjects.getSnapNearestPos(x, y);
            let dX = x;
            let dY = y;
            if(oNearestPos)
            {
                dX = oNearestPos.x;
                dY = oNearestPos.y;
            }
            this.drawingObjects.arrTrackObjects[0].track(e, dX, dY);
            this.drawingObjects.updateOverlay();
        }
        else
        {
            if(AscFormat.isConnectorPreset(this.preset)){
                var oOldState = this.drawingObjects.curState;
                this.drawingObjects.connector = null;
                this.drawingObjects.changeCurrentState(new AscFormat.NullState(this.drawingObjects));
                var oResult;
                this.drawingObjects.handleEventMode = HANDLE_EVENT_MODE_CURSOR;
                oResult = this.drawingObjects.curState.onMouseDown(e, x, y, 0);
                this.drawingObjects.handleEventMode = HANDLE_EVENT_MODE_HANDLE;
                this.drawingObjects.changeCurrentState(oOldState);

                if(oResult) {
                    let oObject = AscCommon.g_oTableId.Get_ById(oResult.objectId);
					if(oObject.canConnectTo && oObject.canConnectTo()) {
						this.drawingObjects.connector = oObject;
					}
                }
                if(this.drawingObjects.connector !== this.oldConnector){
                    this.oldConnector = this.drawingObjects.connector;
                    this.drawingObjects.updateOverlay();
                }
                else{
                    this.oldConnector = this.drawingObjects.connector;
                }
            }
        }
    },

    onMouseUp: function(e, x, y)
    {
        var bRet = false;
        if(this.bStart && (this.drawingObjects.canEdit() || Asc.editor.isDrawSlideshowAnnotations()) && this.drawingObjects.arrTrackObjects.length > 0)
        {
            bRet = true;
            var oThis = this;
            var track =  oThis.drawingObjects.arrTrackObjects[0];
            if(!this.bMoved && this instanceof StartAddNewShape)
            {
                var ext_x, ext_y;
                var oExt = AscFormat.fGetDefaultShapeExtents(this.preset);
                ext_x = oExt.x;
                ext_y = oExt.y;
                this.onMouseMove({IsLocked: true}, this.startX + ext_x, this.startY + ext_y);
            }
            var oTrack = this.drawingObjects.arrTrackObjects[0];
            if(oTrack instanceof AscFormat.PolyLine)
            {

                if(!oTrack.canCreateShape())
                {
                    this.drawingObjects.clearTrackObjects();
                    this.drawingObjects.clearPreTrackObjects();
                    this.drawingObjects.updateOverlay();
                    if(Asc["editor"] && Asc["editor"].wb)
                    {
                        if(!e.fromWindow || this.bStart)
                        {
                            Asc["editor"].asc_endAddShape();
                        }
                    }
                    else if(editor && editor.sync_EndAddShape)
                    {
                        editor.sync_EndAddShape();
                    }
                    this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
                    return;
                }
            }
            if(this.bAnimCustomPath) {
                this.drawingObjects.clearTrackObjects();
                this.drawingObjects.clearPreTrackObjects();
                this.drawingObjects.updateOverlay();
                this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
                let oApi = this.drawingObjects.getEditorApi();
                let oPresentation = oApi.WordControl && oApi.WordControl.m_oLogicDocument;
                if(oPresentation) {
                    let oCurSlide = oPresentation.GetCurrentSlide();
                    if(oCurSlide) {
                        if(oPresentation.IsSelectionLocked(AscCommon.changestype_Timing) === false) {
                            oPresentation.StartAction(0);
                            let oTiming;
                            let aAddedEffects;
                            aAddedEffects = oCurSlide.addAnimation(AscFormat.PRESET_CLASS_PATH, AscFormat.MOTION_SQUARE, 0, null, this.bReplace);
                            oTiming = oCurSlide.timing;
                            if(!oTiming) {
                                oPresentation.FinalizeAction();
                                return;
                            }
                            for(let nEffect = 0; nEffect < aAddedEffects.length; ++nEffect) {
                                let oEffect = aAddedEffects[nEffect];
                                if(!oEffect) {
                                    continue;
                                }
                                let oPathShape = null;
                                oEffect.traverse(function(oChild) {
                                    if(oChild.getObjectType() === AscDFH.historyitem_type_AnimMotion) {
                                        oPathShape = oChild.createPathShape();
                                        return true;
                                    }
                                    return false;
                                });
                                if(!oPathShape) {
                                    continue;
                                }
                                let oPolylineShape = AscFormat.ExecuteNoHistory(function () {
                                    return oTrack.getShape(false, oPresentation.drawingDocument, oCurSlide.graphicObjects);
                                }, this, []);
                                if(!oPolylineShape) {
                                    continue;
                                }
                                let oSpPr = oPolylineShape.spPr;
                                let oXfrm = oSpPr.xfrm;
                                let oGeometry = null;
                                let oPath = null;
                                let dOffX = oXfrm.offX;
                                let dOffY = oXfrm.offY;
                                let oObjectBounds = oPathShape.objectBounds;
                                if(oSpPr.geometry) {
                                    oGeometry = oSpPr.geometry.createDuplicate();
                                    oGeometry.Recalculate(oXfrm.extX, oXfrm.extY);
                                    if(aAddedEffects.length > 1) {
                                        oPath = oGeometry.pathLst[0];
                                        if(oPath && oObjectBounds) {
                                            let oFirstCommand = oPath.ArrPathCommand[0];
                                            if(oFirstCommand && oFirstCommand.id === AscFormat.moveTo) {
                                                dOffX = oObjectBounds.x + oObjectBounds.w/2 - oFirstCommand.X;
                                                dOffY = oObjectBounds.y + oObjectBounds.h/2 - oFirstCommand.Y;
                                            }
                                        }
                                    }
                                }
                                oPathShape.updateAnimation(dOffX, dOffY, oXfrm.extX, oXfrm.extY, oXfrm.rot, oGeometry);
                                oEffect.cTn.setPresetID(AscFormat.MOTION_CUSTOM_PATH);
                                oEffect.cTn.setPresetSubtype(0);
                            }
                            oPresentation.FinalizeAction();
                            if(Asc["editor"] && Asc["editor"].wb)
                            {
                                if(!e.fromWindow || this.bStart)
                                {
                                    Asc["editor"].asc_endAddShape();
                                }
                            }
                            else if(editor && editor.sync_EndAddShape)
                            {
                                editor.sync_EndAddShape();
                            }
                            oPresentation.Document_UpdateInterfaceState();
                            if(this.bPreview && aAddedEffects.length > 0) {
                                let oTiming = oCurSlide.timing;
                                if(oTiming) {
                                    oCurSlide.graphicObjects.resetSelection();
                                    oTiming.resetSelection();
                                    for(let nEffect = 0; nEffect < aAddedEffects.length; ++nEffect) {
                                        aAddedEffects[nEffect].select();
                                    }
                                    oPresentation.StartAnimationPreview();
                                    oTiming.checkSelectedAnimMotionShapes();
                                }
                            }
                            else {
                                oPresentation.DrawingDocument.OnRecalculateSlide(oPresentation.CurPage);
                            }
                        }
                    }
                }
                return;
            }

            let callback = function(bLock, isClickMouseEvent)
            {
                if(bLock)
                {
                    History.Create_NewPoint(AscDFH.historydescription_CommonStatesAddNewShape);
                    var shape = track.getShape(false, oThis.drawingObjects.getDrawingDocument(), oThis.drawingObjects.drawingObjects, isClickMouseEvent);

                    if(!(oThis.drawingObjects.drawingObjects && oThis.drawingObjects.drawingObjects.cSld))
                    {
                        if(shape.spPr.xfrm.offX < 0)
                        {
                            shape.spPr.xfrm.setOffX(0);
                        }
                        if(shape.spPr.xfrm.offY < 0)
                        {
                            shape.spPr.xfrm.setOffY(0);
                        }
                    }
                    oThis.drawingObjects.drawingObjects.getWorksheetModel && shape.setWorksheet(oThis.drawingObjects.drawingObjects.getWorksheetModel());
                    if(oThis.drawingObjects.drawingObjects && oThis.drawingObjects.drawingObjects.cSld)
                    {
                        shape.setParent(oThis.drawingObjects.drawingObjects);
                        shape.setRecalculateInfo();
                    }
                    shape.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
                    shape.checkDrawingBaseCoords();
	                let oAPI = oThis.drawingObjects.getEditorApi();
					if(!oAPI.isDrawInkMode())
					{
						oThis.drawingObjects.checkChartTextSelection();
						oThis.drawingObjects.resetSelection();
						shape.select(oThis.drawingObjects, 0);
						if(oThis.preset === "textRect")
						{
							oThis.drawingObjects.selection.textSelection = shape;
							shape.recalculate();
							shape.selectionSetStart(e, x, y, 0);
							shape.selectionSetEnd(e, x, y, 0);
						}
					}
                    oThis.drawingObjects.startRecalculate();
					if(!oAPI.isDrawInkMode())
					{
						oThis.drawingObjects.drawingObjects.sendGraphicObjectProps();
					}
                    if(oThis.preset && oThis.preset.startsWith("actionButton"))
                    {
                        let sHyperText = "", sHyperValue, sHyperTooltip;
                        switch (oThis.preset) {
                            case "actionButtonBackPrevious": {
                                sHyperValue = "ppaction://hlinkshowjump?jump=previousslide";
                                sHyperTooltip = AscCommon.translateManager.getValue("Previous Slide");
                                break;
                            }
                            case "actionButtonBeginning": {
                                sHyperValue = "ppaction://hlinkshowjump?jump=firstslide";
                                sHyperTooltip = AscCommon.translateManager.getValue("First Slide");
                                break;
                            }
                            case "actionButtonEnd": {
                                sHyperValue = "ppaction://hlinkshowjump?jump=lastslide";
                                sHyperTooltip = AscCommon.translateManager.getValue("Last Slide");
                                break;
                            }
                            case "actionButtonForwardNext": {
                                sHyperValue = "ppaction://hlinkshowjump?jump=nextslide";
                                sHyperTooltip = AscCommon.translateManager.getValue("Next Slide");
                                break;
                            }
                            case "actionButtonHome": {
                                sHyperValue = "ppaction://hlinkshowjump?jump=firstslide";
                                sHyperTooltip = AscCommon.translateManager.getValue("First Slide");
                                break;
                            }
                            case "actionButtonReturn": {
                                sHyperValue = "ppaction://hlinkshowjump?jump=previousslide";
                                sHyperTooltip = AscCommon.translateManager.getValue("Previous Slide");
                                break;
                            }
                        }
                        if(sHyperValue) {
                            oAPI.sendEvent("asc_onDialogAddHyperlink", new Asc.CHyperlinkProperty({Text: sHyperText, Value: sHyperValue, ToolTip: sHyperTooltip}));
                        }
                    }

                }
	            oThis.drawingObjects.updateOverlay();
            };
            if(Asc.editor.isDrawSlideshowAnnotations())
            {
                AscFormat.ExecuteNoHistory(function () {
                    callback(true, e.ClickCount);
                }, this, []);
            }
            else if(Asc.editor.checkObjectsLock)
            {
                Asc.editor.checkObjectsLock([AscCommon.g_oIdCounter.Get_NewId()], callback);
            }
            else
            {
                callback(true, e.ClickCount);
            }
        }
        this.drawingObjects.clearTrackObjects();
        this.drawingObjects.clearPreTrackObjects();
        if(Asc["editor"] && Asc["editor"].wb)
        {
            if(!e.fromWindow || this.bStart)
            {
                Asc["editor"].asc_endAddShape();
            }
        }
        else if(editor && editor.sync_EndAddShape)
        {
            editor.sync_EndAddShape();
        }
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
        return bRet;
    }
};


function checkEmptyPlaceholderContent(content)
{
    if(!content){
        return content;
    }
    var oShape = content.Parent && content.Parent.parent;
    if (oShape) {
        if(content.Is_Empty()){
            if(oShape.isPlaceholder && oShape.isPlaceholder()) {
                return content;
            }
            if(content.isDocumentContentInSmartArtShape && content.isDocumentContentInSmartArtShape()) {
                return content;
            }
        }
        if(oShape.txWarpStruct){
            return content;
        }
        if(oShape.recalcInfo && oShape.recalcInfo.warpGeometry){
            return content;
        }
        var oBodyPr;
        if(oShape.getBodyPr){
            oBodyPr = oShape.getBodyPr();
            if(oBodyPr.vertOverflow !== AscFormat.nVOTOverflow){
                return content;
            }
        }
        var oParagraph = content.GetCurrentParagraph();
        if(oParagraph && oParagraph.IsEmptyWithBullet()) {
            return content;
        }
    }
    return null;
}


function NullState(drawingObjects)
{
    this.drawingObjects = drawingObjects;
    this.startTargetTextObject = null;
    this.lastMoveHandler = null;
}

NullState.prototype =
{
    checkRedrawOnMouseDown: function(oStartContent, oStartPara)
    {
        this.drawingObjects.checkRedrawOnChangeCursorPosition(oStartContent, oStartPara);
    },
    onMouseDown: function(e, x, y, pageIndex)
    {
        let start_target_doc_content, end_target_doc_content, selected_comment_index = -1;
        let oStartPara = null;
        let bHandleMode = this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_HANDLE;
        let sHitGuideId = this.drawingObjects.hitInGuide(x, y);
        let oAnimPlayer = this.drawingObjects.getAnimationPlayer && this.drawingObjects.getAnimationPlayer();
		let oAPI = this.drawingObjects.getEditorApi();
        if(bHandleMode)
        {
            start_target_doc_content = checkEmptyPlaceholderContent(this.drawingObjects.getTargetDocContent());
            if(start_target_doc_content)
            {
                oStartPara = start_target_doc_content.GetCurrentParagraph();
                if(!oStartPara.IsEmpty())
                {
                    oStartPara = null;
                }
            }
            this.startTargetTextObject = AscFormat.getTargetTextObject(this.drawingObjects);
        }
		else
		{
			if(oAPI.editorId === AscCommon.c_oEditorId.Presentation)
			{
				if(oAPI.isFormatPainterOn())
				{
					let oPainterData = oAPI.getFormatPainterData();
					let sType = "default";
					if(oPainterData)
					{
						if(oPainterData.isDrawingData())
						{
							sType = AscCommon.Cursors.ShapeCopy;
						}
						else
						{
							sType = AscCommon.Cursors.TextCopy;
						}
					}
					return {cursorType: sType, objectId: "1"};
				}
			}
        }
        var ret;
        ret = this.drawingObjects.handleSlideComments(e, x, y, pageIndex);
        if(ret)
        {
            if(ret.result)
            {
                return ret.result;
            }
            selected_comment_index = ret.selectedIndex;
        }
        
        var handleAnimLables = null;
        var oTiming = this.drawingObjects && this.drawingObjects.drawingObjects.timing;
        if(oTiming) 
        {
            handleAnimLables = oTiming.onMouseDown(e, x, y, bHandleMode);
        }
    
        var selection = this.drawingObjects.selection;
        if(!handleAnimLables) 
        {
            if(selection.groupSelection)
            {
    
                ret = AscFormat.handleSelectedObjects(this.drawingObjects, e, x, y, selection.groupSelection, pageIndex, false);
                if(ret)
                {
                    if(bHandleMode)
                    {
                        this.checkRedrawOnMouseDown(start_target_doc_content, oStartPara);
                        AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
                    }
                    return ret;
                }
                ret = AscFormat.handleFloatObjects(this.drawingObjects, selection.groupSelection.arrGraphicObjects, e, x, y, selection.groupSelection, pageIndex, false);
                if(ret)
                {
                    if(bHandleMode)
                    {
                        this.checkRedrawOnMouseDown(start_target_doc_content, oStartPara);
                        AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
                    }
                    return ret;
                }
            }
            else if(selection.chartSelection)
            {}
            ret = AscFormat.handleSelectedObjects(this.drawingObjects, e, x, y, null, pageIndex, false);
            if(ret)
            {
                if(bHandleMode)
                {
                    this.checkRedrawOnMouseDown(start_target_doc_content, oStartPara);
                    AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
                }
                return ret;
            }
    
            ret = AscFormat.handleFloatObjects(this.drawingObjects, this.drawingObjects.getDrawingArray(), e, x, y, null, pageIndex, false);
            if(ret)
            {
                if(bHandleMode)
                {
                    this.checkRedrawOnMouseDown(start_target_doc_content, oStartPara);
                    AscCommon.CollaborativeEditing.Update_ForeignCursorsPositions();
                }
                return ret;
            }    
        }

        if(bHandleMode)
        {
            let bRet =  this.drawingObjects.checkChartTextSelection(true);
            if(e.ClickCount < 2)
            {
                this.drawingObjects.resetSelection(undefined, undefined, undefined, !!handleAnimLables);
                if(handleAnimLables)
                {
                    if(oTiming)
                    {
                        oTiming.checkSelectedAnimMotionShapes();
                    }
                }
            }
            if(start_target_doc_content ||
                selected_comment_index > -1 ||
                bRet ||
                handleAnimLables)
            {
                this.drawingObjects.drawingObjects.showDrawingObjects();
            }
            if(this.drawingObjects.drawingObjects && this.drawingObjects.drawingObjects.cSld)
            {
                if(!this.drawingObjects.isSlideShow() && !handleAnimLables)
                {
                    this.drawingObjects.stX = x;
                    this.drawingObjects.stY = y;
                    this.drawingObjects.selectionRect = {x : x, y : y, w: 0, h: 0};
                    this.drawingObjects.changeCurrentState(new TrackSelectionRect(this.drawingObjects));
                }
            }
            if(oAnimPlayer) 
            {
                if(oAnimPlayer.onClick()) 
                {
                    return true;
                }
            }
            if(handleAnimLables) 
            {
                return handleAnimLables;
            }
        }
        else
        {
            if(this.lastMoveHandler)
            {
                if(!this.drawingObjects.isSlideShow()) 
                {
                    var oRet = {};
                    oRet.objectId = this.lastMoveHandler.Get_Id();
                    oRet.bMarker = false;
                    oRet.cursorType = "default";
                    oRet.tooltip = null;
                    return oRet;
                }
                if(handleAnimLables) 
                {
                    return handleAnimLables;
                }
            }
        }
	    if(!oAnimPlayer)
	    {
		    let oGuide = AscCommon.g_oTableId.Get_ById(sHitGuideId);
		    if(oGuide)
		    {
			    if(!bHandleMode)
			    {
				    let bHor = oGuide.isHorizontal();
				    return {cursorType: bHor ? "ns-resize" : "ew-resize", objectId: "1"};
			    }
			    else
			    {
					if(e.Button !== AscCommon.g_mouse_button_right)
					{
						this.drawingObjects.addPreTrackObject(new AscFormat.CGuideTrack(oGuide));
						this.drawingObjects.changeCurrentState(new TrackGuideState(this.drawingObjects, oGuide, x, y))
					}
				    return true;
			    }

		    }
	    }
        return null;
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        let aDrawings = this.drawingObjects.getDrawingArray();
        let _x = x, _y = y, oDrawing;
        this.lastMoveHandler = null;
        for(let nDrawing = aDrawings.length - 1; nDrawing > -1; --nDrawing) {
            oDrawing = aDrawings[nDrawing];
            if(oDrawing.onMouseMove(e, _x, _y)) {
                this.lastMoveHandler = oDrawing;
                _x = -1000;
            }
        }
        this.drawingObjects.checkShowMediaControlOnHover(this.lastMoveHandler);
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        var oTiming = this.drawingObjects && this.drawingObjects.drawingObjects.timing;
        if(oTiming) 
        {
            if(oTiming.onMouseDown(e, x, y, false)) 
            {
                editor.WordControl.m_oLogicDocument.noShowContextMenu = true;
            }
        }
    }
};



    function SlicerState(drawingObjects, oSlicer) {
        this.drawingObjects = drawingObjects;
        this.slicer = oSlicer;
    }
    SlicerState.prototype.onMouseDown = function (e, x, y, pageIndex) {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {cursorType: "default", objectId: this.slicer.Get_Id()};
        }
        return null;
    };
    SlicerState.prototype.onMouseMove = function (e, x, y, pageIndex) {
        if(!e.IsLocked) {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        this.slicer.onMouseMove(e, x, y, pageIndex);
    };
    SlicerState.prototype.onMouseUp = function (e, x, y, pageIndex) {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {cursorType: "default", objectId: this.slicer.Get_Id()};
        }
        var bRet = this.slicer.onMouseUp(e, x, y, pageIndex);
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
        return bRet;
    };
function TrackSelectionRect(drawingObjects)
{
    this.drawingObjects = drawingObjects;
}

TrackSelectionRect.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {cursorType: "default", objectId: "1"};
        }
        return null;
    },
    onMouseMove: function(e, x, y, pageIndex)
    {
        this.drawingObjects.selectionRect = {x : this.drawingObjects .stX, y : this.drawingObjects .stY, w : x - this.drawingObjects .stX, h : y - this.drawingObjects .stY};
        editor.WordControl.m_oDrawingDocument.m_oWordControl.OnUpdateOverlay(true);
    },
    onMouseUp: function(e, x, y, pageIndex)
    {
        var _glyph_index;
        var _glyphs_array;
        if(this.drawingObjects.drawingObjects && this.drawingObjects.drawingObjects.cSld)
        {
            _glyphs_array = this.drawingObjects.drawingObjects.cSld.spTree;
        }
        if(_glyphs_array)
        {
            var _glyph, _glyph_transform;
            var _xlt, _ylt, _xrt, _yrt, _xrb, _yrb, _xlb, _ylb;

            var _rect_l = Math.min(this.drawingObjects.selectionRect.x, this.drawingObjects.selectionRect.x + this.drawingObjects.selectionRect.w);
            var _rect_r = Math.max(this.drawingObjects.selectionRect.x, this.drawingObjects.selectionRect.x + this.drawingObjects.selectionRect.w);
            var _rect_t = Math.min(this.drawingObjects.selectionRect.y, this.drawingObjects.selectionRect.y + this.drawingObjects.selectionRect.h);
            var _rect_b = Math.max(this.drawingObjects.selectionRect.y, this.drawingObjects.selectionRect.y + this.drawingObjects.selectionRect.h);
            for(_glyph_index = 0; _glyph_index < _glyphs_array.length; ++_glyph_index)
            {
                _glyph = _glyphs_array[_glyph_index];
                _glyph_transform = _glyph.transform;

                _xlt = _glyph_transform.TransformPointX(0, 0);
                _ylt = _glyph_transform.TransformPointY(0, 0);

                _xrt = _glyph_transform.TransformPointX( _glyph.extX, 0);
                _yrt = _glyph_transform.TransformPointY( _glyph.extX, 0);

                _xrb = _glyph_transform.TransformPointX( _glyph.extX, _glyph.extY);
                _yrb = _glyph_transform.TransformPointY( _glyph.extX, _glyph.extY);

                _xlb = _glyph_transform.TransformPointX(0, _glyph.extY);
                _ylb = _glyph_transform.TransformPointY(0, _glyph.extY);

                if((_xlb >= _rect_l && _xlb <= _rect_r) && (_xrb >= _rect_l && _xrb <= _rect_r)
                    && (_xlt >= _rect_l && _xlt <= _rect_r) && (_xrt >= _rect_l && _xrt <= _rect_r) &&
                    (_ylb >= _rect_t && _ylb <= _rect_b) && (_yrb >= _rect_t && _yrb <= _rect_b)
                    && (_ylt >= _rect_t && _ylt <= _rect_b) && (_yrt >= _rect_t && _yrt <= _rect_b))
                {
                    this.drawingObjects.selectObject(_glyph, pageIndex);
                }
            }
        }
        this.drawingObjects.selectionRect = null;
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
        editor.WordControl.m_oDrawingDocument.m_oWordControl.OnUpdateOverlay(true);
        editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
    }

};




    function TrackGuideState(drawingObjects, oGuide, dStartX, dStartY) {
        this.drawingObjects = drawingObjects;
        this.guide = oGuide;
        this.tracked = false;
        this.startX = dStartX;
        this.startY = dStartY;
        let oTrack = this.drawingObjects.arrPreTrackObjects[0];
        if(oTrack) {
            let dPos = AscFormat.GdPosToMm(this.guide.pos);
	        dPos = (dPos * 10 + 0.5 >> 0) / 10;
            let oConvertedPos = editor.WordControl.m_oDrawingDocument.ConvertCoordsToCursorWR(dStartX, dStartY, 0);
            editor.sendEvent("asc_onTrackGuide", dPos, oConvertedPos.X, oConvertedPos.Y);
        }
    }
    TrackGuideState.prototype.onMouseDown = function (e, x, y, pageIndex) {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR) {
            let bHor = this.guide.isHorizontal();
            return {cursorType: bHor ? "ns-resize" : "ew-resize" , objectId: "1"};
        }
        return null;
    };
    TrackGuideState.prototype.onMouseMove = function (e, x, y, pageIndex) {
        if(!e.IsLocked) {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        let bHor = this.guide.isHorizontal();
        if(!this.tracked) {
            if(bHor && Math.abs(y - this.startY) > MOVE_DELTA ||
                !bHor && Math.abs(x - this.startX) > MOVE_DELTA) {
                this.tracked = true;
                this.drawingObjects.swapTrackObjects();
                this.onMouseMove(e, x, y, pageIndex);
                return;
            }
        }
        else {
            let oTrack = this.drawingObjects.arrTrackObjects[0];
            if(oTrack) {
	            let oNearestPos = this.drawingObjects.getSnapNearestPos(x, y);
				let dX = x;
				let dY = y;
	            if(oNearestPos) {
		            dX = oNearestPos.x;
		            dY = oNearestPos.y;
	            }
                oTrack.track(dX, dY);
                let oConvertedPos = editor.WordControl.m_oDrawingDocument.ConvertCoordsToCursorWR(bHor ? x : dX, !bHor ? y : dY, 0);
				let dGdPos = oTrack.getPos();
	            let dPos = AscFormat.GdPosToMm(dGdPos);
	            dPos = (dPos * 10 + 0.5 >> 0) / 10;
                editor.sendEvent("asc_onTrackGuide", dPos, oConvertedPos.X, oConvertedPos.Y)
                this.drawingObjects.updateOverlay();
            }
        }
    };
    TrackGuideState.prototype.onMouseUp = function (e, x, y, pageIndex) {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR) {
            let bHor = this.guide.isHorizontal();
            return {cursorType: bHor ? "ew-resize" : "ns-resize", objectId: "1"};
        }
        if(this.tracked) {
            let oPresentation = editor.WordControl.m_oLogicDocument;
            if(false === oPresentation.Document_Is_SelectionLocked(AscCommon.changestype_ViewPr, undefined, undefined, [])) {
                this.drawingObjects.trackEnd();
                oPresentation.Recalculate();
            }
        }
        this.drawingObjects.clearPreTrackObjects();
        this.drawingObjects.clearTrackObjects();
        this.drawingObjects.updateOverlay();
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
        editor.sendEvent("asc_onTrackGuide");
    };

function PreChangeAdjState(drawingObjects, majorObject)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
}

PreChangeAdjState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject && this.majorObject.Get_Id(), bMarker: true, cursorType: "crosshair"};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        this.drawingObjects.swapTrackObjects();
        this.drawingObjects.changeCurrentState(new ChangeAdjState(this.drawingObjects, this.majorObject));
        this.drawingObjects.OnMouseMove(e, x, y, pageIndex);
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        this.drawingObjects.clearPreTrackObjects();
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
    }
};

function ChangeAdjState(drawingObjects, majorObject)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
}

ChangeAdjState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject && this.majorObject.Get_Id(), bMarker: true, cursorType: "crosshair"};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        var t = AscFormat.CheckCoordsNeedPage(x, y, pageIndex, this.majorObject.selectStartPage, this.drawingObjects.getDrawingDocument());
        for(var i = 0; i < this.drawingObjects.arrTrackObjects.length; ++i){
            this.drawingObjects.arrTrackObjects[i].track(t.x, t.y);
        }
        this.drawingObjects.updateOverlay();
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.canEdit())
        {
            if(!this.drawingObjects.checkSelectedObjectsProtection())
            {
                var trackObjects = [].concat(this.drawingObjects.arrTrackObjects);
                var drawingObjects = this.drawingObjects;
                this.drawingObjects.checkSelectedObjectsAndCallback(function()
                {
                    var oOriginalObjects = [];
                    var oMapOriginalsIds = {};
                    for(var i = 0; i < trackObjects.length; ++i){
                        trackObjects[i].trackEnd();
                        if(trackObjects[i].originalObject && !trackObjects[i].processor3D){
                            oOriginalObjects.push(trackObjects[i].originalObject);
                            oMapOriginalsIds[trackObjects[i].originalObject.Get_Id()] = true;
                        }
                    }
                    var aAllConnectors = drawingObjects.getAllConnectorsByDrawings(oOriginalObjects, [],  undefined, true);
                    for(i = 0; i < aAllConnectors.length; ++i){
                        if(!oMapOriginalsIds[aAllConnectors[i].Get_Id()]){
                            aAllConnectors[i].calculateTransform();
                        }
                    }

                    drawingObjects.startRecalculate();
                },[], false, AscDFH.historydescription_CommonDrawings_ChangeAdj);
            }

        }
        this.drawingObjects.clearTrackObjects();
        this.drawingObjects.updateOverlay();
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
    }
};


function PreRotateState(drawingObjects, majorObject)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
}

PreRotateState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        this.drawingObjects.swapTrackObjects();
        this.drawingObjects.changeCurrentState(new RotateState(this.drawingObjects, this.majorObject));
        this.drawingObjects.OnMouseMove(e, x, y, pageIndex);
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        this.drawingObjects.clearPreTrackObjects();
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
    }
};

function RotateState(drawingObjects, majorObject)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
}

RotateState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject && this.majorObject.Get_Id(), bMarker: true, cursorType: "crosshair"};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        var coords = AscFormat.CheckCoordsNeedPage(x, y, pageIndex, this.majorObject.selectStartPage, this.drawingObjects.getDrawingDocument());
        this.drawingObjects.handleRotateTrack(e, coords.x, coords.y);
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.canEdit() && this.bSamePos !== true)
        {
            var tracks = [].concat(this.drawingObjects.arrTrackObjects);
            var group = this.group;
            var drawingObjects = this.drawingObjects;
            var oThis = this;
            var bIsMoveState = (this instanceof MoveState);
            var bIsChartFrame = Asc["editor"] && Asc["editor"].isChartEditor === true;
            var bIsTrackInChart = (tracks.length > 0 && (tracks[0] instanceof AscFormat.MoveChartObjectTrack));
            var bCopyOnMove = e.CtrlKey && bIsMoveState && !bIsChartFrame && !bIsTrackInChart;
            var bCopyOnMoveInGroup = (e.CtrlKey && oThis instanceof MoveInGroupState && !oThis.hasObjectInSmartArt);
            var i, j;
            var copy;
            if(bCopyOnMove)
            {
                this.drawingObjects.resetSelection();
                var oIdMap = {};
                var oCopyPr = new AscFormat.CCopyObjectProperties();
                oCopyPr.idMap = oIdMap;
                var aCopies = [];
                History.Create_NewPoint(AscDFH.historydescription_CommonDrawings_CopyCtrl);
                for(i = 0; i < tracks.length; ++i)
                {
	                tracks[i].checkDrawingPartWithHistory();
                    copy = tracks[i].originalObject.copy(oCopyPr);
                    oIdMap[tracks[i].originalObject.Id] = copy.Id;
                    this.drawingObjects.drawingObjects.getWorksheetModel && copy.setWorksheet(this.drawingObjects.drawingObjects.getWorksheetModel());
                    if(this.drawingObjects.drawingObjects && this.drawingObjects.drawingObjects.cSld)
                    {
                        copy.setParent2(this.drawingObjects.drawingObjects);
                        if(!copy.spPr || !copy.spPr.xfrm
                            || ((copy.getObjectType() === AscDFH.historyitem_type_GroupShape || copy.getObjectType() === AscDFH.historyitem_type_SmartArt) && !copy.spPr.xfrm.isNotNullForGroup() || copy.getObjectType() !== AscDFH.historyitem_type_GroupShape && !copy.spPr.xfrm.isNotNull()))
                        {
                            copy.recalculateTransform();
                        }
                    }
                    if(tracks[i].originalObject.drawingBase)
                    {
                        var drawingObject = tracks[i].originalObject.drawingBase;
                        var metrics = drawingObject.getGraphicObjectMetrics();
                        AscFormat.SetXfrmFromMetrics(copy, metrics);
                        copy.drawingBase = drawingObject;
                    }
                    copy.addToDrawingObjects();
                    aCopies.push(copy);

                    tracks[i].originalObject = copy;
                    tracks[i].trackEnd(false, true);
                    this.drawingObjects.selectObject(copy, 0);
                }
                if(!(this.drawingObjects.drawingObjects && this.drawingObjects.drawingObjects.cSld))
                {
                    if(History.StartTransaction && History.EndTransaction)
                    {
                        History.StartTransaction();
                    }
                    AscFormat.ExecuteNoHistory(function(){drawingObjects.checkSelectedObjectsAndCallback(function(){}, []);}, this, []);
                    if(History.StartTransaction && History.EndTransaction)
                    {
                        History.EndTransaction();
                    }
                    if(this.drawingObjects.checkSlicerCopies)
                    {
                        this.drawingObjects.checkSlicerCopies(aCopies);
                    }
                }
                else
                {
                    this.drawingObjects.startRecalculate();
                    this.drawingObjects.drawingObjects.sendGraphicObjectProps();
                }
                AscFormat.fResetConnectorsIds(aCopies, oIdMap);
            }
            else
            {
                if(bCopyOnMoveInGroup)
                {
                    if(!this.drawingObjects.checkSelectedObjectsProtection())
                    {
                        this.drawingObjects.checkSelectedObjectsAndCallback(function(){
                            var oIdMap = {};
                            var aCopies = [];
                            var oCopyPr = new AscFormat.CCopyObjectProperties();
                            oCopyPr.idMap = oIdMap;
                            group.resetSelection();
                            for(i = 0; i < tracks.length; ++i)
                            {
                                copy = tracks[i].originalObject.copy(oCopyPr);
                                aCopies.push(copy);
                                oThis.drawingObjects.drawingObjects.getWorksheetModel && copy.setWorksheet(oThis.drawingObjects.drawingObjects.getWorksheetModel());
                                if(oThis.drawingObjects.drawingObjects && oThis.drawingObjects.drawingObjects.cSld)
                                {
                                    copy.setParent2(oThis.drawingObjects.drawingObjects);
                                }
                                copy.setGroup(tracks[i].originalObject.group);
                                copy.group.addToSpTree(copy.group.length, copy);
                                tracks[i].originalObject = copy;
                                tracks[i].trackEnd(false);
                                group.selectObject(copy, 0);
                            }
                            AscFormat.fResetConnectorsIds(aCopies, oIdMap);
                            if(group)
                            {
                                group.updateCoordinatesAfterInternalResize();
                            }
                            if(!oThis.drawingObjects.drawingObjects || !oThis.drawingObjects.drawingObjects.cSld)
                            {
                                var min_x, min_y, drawing, arr_x2 = [], arr_y2 = [], oTransform;
                                for(i = 0; i < oThis.drawingObjects.selectedObjects.length; ++i)
                                {
                                    drawing = oThis.drawingObjects.selectedObjects[i];
                                    var rot = AscFormat.isRealNumber(drawing.spPr.xfrm.rot) ? drawing.spPr.xfrm.rot : 0;
                                    rot = AscFormat.normalizeRotate(rot);
                                    arr_x2.push(drawing.spPr.xfrm.offX);
                                    arr_y2.push(drawing.spPr.xfrm.offY);
                                    arr_x2.push(drawing.spPr.xfrm.offX + drawing.spPr.xfrm.extX);
                                    arr_y2.push(drawing.spPr.xfrm.offY + drawing.spPr.xfrm.extY);
                                    if (AscFormat.checkNormalRotate(rot))
                                    {
                                        min_x = drawing.spPr.xfrm.offX;
                                        min_y = drawing.spPr.xfrm.offY;
                                    }
                                    else
                                    {
                                        min_x = drawing.spPr.xfrm.offX + drawing.spPr.xfrm.extX/2 - drawing.spPr.xfrm.extY/2;
                                        min_y = drawing.spPr.xfrm.offY + drawing.spPr.xfrm.extY/2 - drawing.spPr.xfrm.extX/2;
                                        arr_x2.push(min_x);
                                        arr_y2.push(min_y);
                                        arr_x2.push(min_x + drawing.spPr.xfrm.extY);
                                        arr_y2.push(min_y + drawing.spPr.xfrm.extX);
                                    }
                                    if(min_x < 0)
                                    {
                                        drawing.spPr.xfrm.setOffX(drawing.spPr.xfrm.offX - min_x);
                                    }
                                    if(min_y < 0)
                                    {
                                        drawing.spPr.xfrm.setOffY(drawing.spPr.xfrm.offY - min_y);
                                    }
                                    drawing.checkDrawingBaseCoords();
                                    drawing.recalculateTransform();
                                    oTransform = drawing.transform;
                                    arr_x2.push(oTransform.TransformPointX(0, 0));
                                    arr_y2.push(oTransform.TransformPointY(0, 0));
                                    arr_x2.push(oTransform.TransformPointX(drawing.extX, 0));
                                    arr_y2.push(oTransform.TransformPointY(drawing.extX, 0));
                                    arr_x2.push(oTransform.TransformPointX(drawing.extX, drawing.extY));
                                    arr_y2.push(oTransform.TransformPointY(drawing.extX, drawing.extY));
                                    arr_x2.push(oTransform.TransformPointX(0, drawing.extY));
                                    arr_y2.push(oTransform.TransformPointY(0, drawing.extY));
                                }
                                oThis.drawingObjects.drawingObjects.checkGraphicObjectPosition(0, 0, Math.max.apply(Math, arr_x2), Math.max.apply(Math, arr_y2));
                            }
                            if(oThis.drawingObjects.checkSlicerCopies)
                            {
                                oThis.drawingObjects.checkSlicerCopies(aCopies);
                            }
                        }, [], false, AscDFH.historydescription_CommonDrawings_EndTrack);
                    }
                }
                else
                {
                    var oOriginalObjects = [];
                    var oMapOriginalsId = {};
                    var oMapAdditionalForCheck = {};
                    for(i = 0; i < tracks.length; ++i)
                    {
                        var oOrigObject = tracks[i].originalObject && tracks[i].chartSpace ? tracks[i].chartSpace : tracks[i].originalObject;
                        if(tracks[i].originalObject && !tracks[i].processor3D){
                            oOriginalObjects.push(oOrigObject);
                            oMapOriginalsId[oOrigObject.Get_Id()] = true;
                            var oGroup = oOrigObject.getMainGroup && oOrigObject.getMainGroup();
                            if(oGroup){
                                if(!oGroup.selected){
                                    oMapAdditionalForCheck[oGroup.Get_Id()] = oGroup;
                                }
                            }
                            else{
                                if(!oOrigObject.selected){
                                    oMapAdditionalForCheck[oOrigObject.Get_Id()] = oOrigObject;
                                }
                            }

                            if(Array.isArray(oOrigObject.arrGraphicObjects)){
                                for(j = 0; j < oOrigObject.arrGraphicObjects.length; ++j){
                                    oMapOriginalsId[oOrigObject.arrGraphicObjects[j].Get_Id()] = true;
                                }
                            }
                        }
                    }
                    var aAllConnectors = drawingObjects.getAllConnectorsByDrawings(oOriginalObjects, [],  undefined, true);
                    var bFlag = ((oThis instanceof MoveInGroupState) || (oThis instanceof MoveState));
                    var aConnectors = [];
                    for(i = 0; i < aAllConnectors.length; ++i){

                        var stSp = AscCommon.g_oTableId.Get_ById(aAllConnectors[i].getStCxnId());
                        var endSp = AscCommon.g_oTableId.Get_ById(aAllConnectors[i].getEndCxnId());
                        if((stSp && !oMapOriginalsId[stSp.Get_Id()]) || (endSp && !oMapOriginalsId[endSp.Get_Id()]) || !oMapOriginalsId[aAllConnectors[i].Get_Id()]){
                            var oGroup = aAllConnectors[i].getMainGroup && aAllConnectors[i].getMainGroup();
                            aConnectors.push(aAllConnectors[i]);
                            if(oGroup){
                                oMapAdditionalForCheck[oGroup.Id] = oGroup;
                            }
                            else{
                                oMapAdditionalForCheck[aAllConnectors[i].Get_Id()] = aAllConnectors[i];
                            }
                        }
                    }
                    var aAdditionalForCheck = [];
                    for(i in oMapAdditionalForCheck){
                        if(oMapAdditionalForCheck.hasOwnProperty(i)){
                            if(!oMapAdditionalForCheck[i].selected){
                                aAdditionalForCheck.push(oMapAdditionalForCheck[i]);
                            }
                        }
                    }
                    if(!this.drawingObjects.checkSelectedObjectsProtection())
                    {
                        this.drawingObjects.checkSelectedObjectsAndCallback(function () {

                                for(i = 0; i < tracks.length; ++i){
	                                tracks[i].checkDrawingPartWithHistory();
                                    tracks[i].trackEnd(false, bFlag);
                                }
                                if(tracks.length === 1 && tracks[0].chartSpace){
                                    return;
                                }
                                var oGroupMaps = {};
                                for(i = 0; i < aConnectors.length; ++i){
                                    aConnectors[i].calculateTransform(bFlag);
                                    var oGroup = aConnectors[i].getMainGroup && aConnectors[i].getMainGroup();
                                    if(oGroup){
                                        oGroupMaps[oGroup.Id] = oGroup;
                                    }
                                }
                                for(var key in oGroupMaps){
                                    if(oGroupMaps.hasOwnProperty(key)){
                                        oGroupMaps[key].updateCoordinatesAfterInternalResize();
                                    }
                                }
                                if(group)
                                {
                                    group.updateCoordinatesAfterInternalResize();
                                }
                                if(!oThis.drawingObjects.drawingObjects || !oThis.drawingObjects.drawingObjects.cSld)
                                {
                                    var min_x, min_y, drawing, arr_x2 = [], arr_y2 = [], oTransform;
                                    for(i = 0; i < oThis.drawingObjects.selectedObjects.length; ++i)
                                    {
                                        drawing = oThis.drawingObjects.selectedObjects[i];
                                        var rot = AscFormat.isRealNumber(drawing.spPr.xfrm.rot) ? drawing.spPr.xfrm.rot : 0;
                                        rot = AscFormat.normalizeRotate(rot);
                                        arr_x2.push(drawing.spPr.xfrm.offX);
                                        arr_y2.push(drawing.spPr.xfrm.offY);
                                        arr_x2.push(drawing.spPr.xfrm.offX + drawing.spPr.xfrm.extX);
                                        arr_y2.push(drawing.spPr.xfrm.offY + drawing.spPr.xfrm.extY);
                                        if (AscFormat.checkNormalRotate(rot))
                                        {
                                            min_x = drawing.spPr.xfrm.offX;
                                            min_y = drawing.spPr.xfrm.offY;
                                        }
                                        else
                                        {
                                            min_x = drawing.spPr.xfrm.offX + drawing.spPr.xfrm.extX/2 - drawing.spPr.xfrm.extY/2;
                                            min_y = drawing.spPr.xfrm.offY + drawing.spPr.xfrm.extY/2 - drawing.spPr.xfrm.extX/2;
                                            arr_x2.push(min_x);
                                            arr_y2.push(min_y);
                                            arr_x2.push(min_x + drawing.spPr.xfrm.extY);
                                            arr_y2.push(min_y + drawing.spPr.xfrm.extX);
                                        }
                                        if(min_x < 0)
                                        {
                                            drawing.spPr.xfrm.setOffX(drawing.spPr.xfrm.offX - min_x);
                                        }
                                        if(min_y < 0)
                                        {
                                            drawing.spPr.xfrm.setOffY(drawing.spPr.xfrm.offY - min_y);
                                        }
                                        drawing.checkDrawingBaseCoords && drawing.checkDrawingBaseCoords();
                                        drawing.recalculateTransform && drawing.recalculateTransform();
                                        oTransform = drawing.transform;
                                        arr_x2.push(oTransform.TransformPointX(0, 0));
                                        arr_y2.push(oTransform.TransformPointY(0, 0));
                                        arr_x2.push(oTransform.TransformPointX(drawing.extX, 0));
                                        arr_y2.push(oTransform.TransformPointY(drawing.extX, 0));
                                        arr_x2.push(oTransform.TransformPointX(drawing.extX, drawing.extY));
                                        arr_y2.push(oTransform.TransformPointY(drawing.extX, drawing.extY));
                                        arr_x2.push(oTransform.TransformPointX(0, drawing.extY));
                                        arr_y2.push(oTransform.TransformPointY(0, drawing.extY));
                                    }
                                    oThis.drawingObjects.drawingObjects.checkGraphicObjectPosition(0, 0, Math.max.apply(Math, arr_x2), Math.max.apply(Math, arr_y2));
                                }
                            }, [], false, AscDFH.historydescription_CommonDrawings_EndTrack, aAdditionalForCheck
                        );
                    }
                }
            }
        }
        this.drawingObjects.resetTracking();
    }
};


function PreResizeState(drawingObjects, majorObject, cardDirection)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
    this.cardDirection = cardDirection;
    this.handleNum = this.majorObject.getNumByCardDirection(cardDirection);
}

PreResizeState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        this.drawingObjects.swapTrackObjects();
        this.drawingObjects.changeCurrentState(new ResizeState(this.drawingObjects, this.majorObject, this.handleNum, this.cardDirection));
        this.drawingObjects.OnMouseMove(e, x, y, pageIndex);
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        this.drawingObjects.clearPreTrackObjects();
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
    }
};

function ResizeState(drawingObjects, majorObject, handleNum, cardDirection)
{
    this.drawingObjects = drawingObjects;
    this.majorObject  = majorObject;
    this.handleNum = handleNum;
    this.cardDirection = cardDirection;
}

ResizeState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        var start_arr = this.drawingObjects.getDrawingArray();
        var coords = AscFormat.CheckCoordsNeedPage(x, y, pageIndex, this.majorObject.selectStartPage, this.drawingObjects.getDrawingDocument());
        let dX = coords.x;
        let dY = coords.y;
        let oNearestPos = this.drawingObjects.getSnapNearestPos(coords.x, coords.y);
        if(oNearestPos)
        {
            dX = oNearestPos.x;
            dY = oNearestPos.y;
        }
        var resize_coef = this.majorObject.getResizeCoefficients(this.handleNum, dX, dY, start_arr, this.drawingObjects);
        this.drawingObjects.trackResizeObjects(resize_coef.kd1, resize_coef.kd2, e, dX, dY);
        if(this.drawingObjects.drawingObjects.cSld)
        {
            if(AscFormat.isRealNumber(resize_coef.snapX) && !resize_coef.horGuideSnap)
            {
                this.drawingObjects.getDrawingDocument().DrawVerAnchor(pageIndex, resize_coef.snapX);
            }
            if(AscFormat.isRealNumber(resize_coef.snapY) && !resize_coef.vertGuideSnap)
            {
                this.drawingObjects.getDrawingDocument().DrawHorAnchor(pageIndex, resize_coef.snapY);
            }
        }
        this.drawingObjects.updateOverlay();
    },

    onMouseUp: RotateState.prototype.onMouseUp
};


function PreMoveState(drawingObjects,  startX, startY, shift, ctrl, majorObject, majorObjectIsSelected, bInside, bGroupSelection)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
    this.startX = startX;
    this.startY = startY;
    this.shift = shift;
    this.ctrl = ctrl;
    this.majorObjectIsSelected = majorObjectIsSelected;
    this.bInside = bInside;
    this.bGroupSelection = bGroupSelection;
}

PreMoveState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "move", bMarker: true};
        }
        else{
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {

        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        if(Math.abs(this.startX - x) > MOVE_DELTA || Math.abs(this.startY - y) > MOVE_DELTA || pageIndex !== this.majorObject.selectStartPage)
        {
            if(this.drawingObjects.isSlideShow())
            {
                this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
                return;
            }
            this.drawingObjects.swapTrackObjects();
            this.drawingObjects.changeCurrentState(new MoveState(this.drawingObjects, this.majorObject, this.startX, this.startY));
            this.drawingObjects.OnMouseMove(e, x, y, pageIndex);
        }
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        return AscFormat.handleMouseUpPreMoveState(this.drawingObjects, e, x, y, pageIndex, true)
    }
};

function MoveState(drawingObjects, majorObject, startX, startY)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;

    this.startX = startX;
    this.startY = startY;

    var arr_x = [], arr_y = [];
    for(var i = 0; i < this.drawingObjects.arrTrackObjects.length; ++i)
    {
        var track = this.drawingObjects.arrTrackObjects[i];
        var transform = track.originalObject.transform;
        arr_x.push(transform.TransformPointX(0, 0));
        arr_y.push(transform.TransformPointY(0, 0));
        arr_x.push(transform.TransformPointX(track.originalObject.extX, 0));
        arr_y.push(transform.TransformPointY(track.originalObject.extX, 0));
        arr_x.push(transform.TransformPointX(track.originalObject.extX, track.originalObject.extY));
        arr_y.push(transform.TransformPointY(track.originalObject.extX, track.originalObject.extY));
        arr_x.push(transform.TransformPointX(0, track.originalObject.extY));
        arr_y.push(transform.TransformPointY(0, track.originalObject.extY));
    }
    this.rectX = Math.min.apply(Math, arr_x);
    this.rectY = Math.min.apply(Math, arr_y);
    this.rectW = Math.max.apply(Math, arr_x) - this.rectX;
    this.rectH = Math.max.apply(Math, arr_y) - this.rectY;
    this.bSamePos = true;
}

MoveState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "move", bMarker: true};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        let aTracks = this.drawingObjects.arrTrackObjects;
        let nTracksCount = aTracks.length;
        let nTrack;
        let bIsSlide = false;
        if(this.drawingObjects &&
            this.drawingObjects.drawingObjects &&
            this.drawingObjects.drawingObjects.cSld)
        {
            bIsSlide = true;
        }
        let dResultX, dResultY;
        if(!e.ShiftKey)
        {
            dResultX = x;
            dResultY = y;
        }
        else
        {
            let dAbsDistX = Math.abs(this.startX - x);
            let dAbsDistY = Math.abs(this.startY - y);
            if(dAbsDistX > dAbsDistY)
            {
                dResultX = x;
                dResultY = this.startY;
            }
            else
            {
                dResultX = this.startX;
                dResultY = y;
            }
        }
        let aDrawings = this.drawingObjects.getAllObjectsOnPage(0);
        let dMinDx = null, dMinDy = null;
        let dDx, dDy;
        let dCurDx = dResultX - this.startX;
        let dCurDy = dResultY - this.startY;
        let aSnapX = [], aSnapY = [];
        let oCurTrackOriginal;
        let aTrackSnapX, aTrackSnapY;
        let nSnapPos;
        let oSnapData;

        let aVertGuidesPos = this.drawingObjects.getVertGuidesPos();
        let aHorGuidesPos = this.drawingObjects.getHorGuidesPos();

        //-------------------------------------------------
        for(nTrack = 0; nTrack < nTracksCount; ++nTrack)
        {
            oCurTrackOriginal = aTracks[nTrack].originalObject;
            aTrackSnapX = oCurTrackOriginal.snapArrayX;
            if(!aTrackSnapX)
            {
                continue;
            }

            if(aDrawings.length > 0)
            {
                for(nSnapPos = 0; nSnapPos < aTrackSnapX.length; ++nSnapPos)
                {
                    oSnapData = AscFormat.GetMinSnapDistanceXObject(aTrackSnapX[nSnapPos] + dCurDx, aDrawings, undefined, aVertGuidesPos);
                    if(oSnapData)
                    {
                        dDx = oSnapData.dist;
                        if(dDx !== null)
                        {
                            if(dMinDx === null)
                            {
                                dMinDx = dDx;
                               !oSnapData.guide && aSnapX.push(oSnapData.pos);
                            }
                            else
                            {
                                if(AscFormat.fApproxEqual(dMinDx, dDx, 0.01))
                                {
                                    !oSnapData.guide && aSnapX.push(oSnapData.pos);
                                }
                                else if(Math.abs(dMinDx) > Math.abs(dDx))
                                {
                                    dMinDx = dDx;
                                    aSnapX.length = 0;
                                    !oSnapData.guide && aSnapX.push(oSnapData.pos);
                                }
                            }
                        }
                    }
                }
            }
        }
        if(AscFormat.fApproxEqual(dResultX, this.startX))
        {
            dMinDx = 0;
        }

        //-----------------------------
        for(nTrack = 0; nTrack < aTracks.length; ++nTrack)
        {
            oCurTrackOriginal = aTracks[nTrack].originalObject;
            aTrackSnapY = oCurTrackOriginal.snapArrayY;
            if(!aTrackSnapY)
            {
                continue;
            }
            if(aDrawings.length > 0)
            {
                for(nSnapPos = 0; nSnapPos < aTrackSnapY.length; ++nSnapPos)
                {
                    oSnapData = AscFormat.GetMinSnapDistanceYObject(aTrackSnapY[nSnapPos] + dCurDy, aDrawings, undefined, aHorGuidesPos);
                    if(oSnapData)
                    {
                        dDy = oSnapData.dist;
                        if(dDy !== null)
                        {
                            if(dMinDy === null)
                            {
                                dMinDy = dDy;
                                !oSnapData.guide && aSnapY.push(oSnapData.pos);
                            }
                            else
                            {
                                if(AscFormat.fApproxEqual(dMinDy, dDy, 0.01))
                                {
                                    !oSnapData.guide && aSnapY.push(oSnapData.pos);
                                }
                                else if(Math.abs(dMinDy) > Math.abs(dDy))
                                {
                                    dMinDy = dDy;
                                    aSnapY.length = 0;
                                    !oSnapData.guide && aSnapY.push(oSnapData.pos);
                                }
                            }
                        }
                    }
                }
            }
        }
        if(AscFormat.fApproxEqual(dResultY, this.startY))
        {
            dMinDy = 0;
        }

        let oMajorBounds = this.majorObject.getRectBounds();
        let oNearestSnapPos = this.drawingObjects.getSnapNearestPos(oMajorBounds.l + dCurDx, oMajorBounds.t + dCurDy);
        if(dMinDx === null || Math.abs(dMinDx) > SNAP_DISTANCE)
        {
            dMinDx = 0;
            if(oNearestSnapPos)
            {
                if(!AscFormat.fApproxEqual(dResultX, this.startX))
                {
                    let dDeltaX = oNearestSnapPos.x - oMajorBounds.x;
                    dMinDx = dDeltaX - dResultX + this.startX;
                }
            }
        }
        else
        {
            if(bIsSlide)
            {
                for(nSnapPos = 0; nSnapPos < aSnapX.length; ++nSnapPos)
                {
                    this.drawingObjects.getDrawingDocument().DrawVerAnchor(pageIndex, aSnapX[nSnapPos]);
                }

            }
        }

        if(dMinDy === null || Math.abs(dMinDy) > SNAP_DISTANCE)
        {
            dMinDy = 0;
            if(oNearestSnapPos)
            {
                if(!AscFormat.fApproxEqual(dResultY, this.startY))
                {
                    let dDeltaY = oNearestSnapPos.y - oMajorBounds.y;
                    dMinDy = dDeltaY - dResultY + this.startY;
                }
            }
        }
        else
        {
            if(bIsSlide)
            {
                for(nSnapPos = 0; nSnapPos < aSnapY.length; ++nSnapPos)
                {
                    this.drawingObjects.getDrawingDocument().DrawHorAnchor(pageIndex, aSnapY[nSnapPos]);
                }
            }
        }

        dDx = dResultX - this.startX + dMinDx;
        dDy = dResultY - this.startY + dMinDy;
        let oCheckPosition = this.drawingObjects.checkGraphicObjectPosition(this.rectX + dDx, this.rectY + dDy, this.rectW, this.rectH);
        for(nTrack = 0; nTrack < nTracksCount; ++nTrack)
        {
            aTracks[nTrack].track(dDx + oCheckPosition.x, dDy + oCheckPosition.y, pageIndex);
        }
        this.bSamePos = (AscFormat.fApproxEqual(dDx + oCheckPosition.x, 0) && AscFormat.fApproxEqual(dDy + oCheckPosition.y, 0));
        this.drawingObjects.updateOverlay();
    },

    onMouseUp: RotateState.prototype.onMouseUp
};

function PreMoveInGroupState(drawingObjects, group, startX, startY, ShiftKey, CtrlKey, majorObject,  majorObjectIsSelected)
{
    this.drawingObjects = drawingObjects;
    this.group = group;
    this.startX = startX;
    this.startY = startY;
    this.ShiftKey = ShiftKey;
    this.CtrlKey  = CtrlKey;
    this.majorObject = majorObject;
    this.majorObjectIsSelected = majorObjectIsSelected;
}

PreMoveInGroupState.prototype =
{
    onMouseDown: function(e, x,y,pageIndex)
    {},

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        if(Math.abs(this.startX - x) > MOVE_DELTA || Math.abs(this.startY - y) > MOVE_DELTA || pageIndex !== this.majorObject.selectStartPage)
        {
            if(this.drawingObjects.isSlideShow())
            {

                this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
                return;
            }
            this.drawingObjects.swapTrackObjects();
            this.drawingObjects.changeCurrentState(new MoveInGroupState(this.drawingObjects, this.majorObject, this.group, this.startX, this.startY));
            this.drawingObjects.OnMouseMove(e, x, y, pageIndex);
        }
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(e.CtrlKey && this.majorObjectIsSelected)
        {
            this.group.deselectObject(this.majorObject);
            if(this.group.selectedObjects.length === 0){
                this.drawingObjects.resetInternalSelection();
            }
            this.drawingObjects.drawingObjects && this.drawingObjects.drawingObjects.sendGraphicObjectProps && this.drawingObjects.drawingObjects.sendGraphicObjectProps();
            this.drawingObjects.updateOverlay();
        }
        this.drawingObjects.clearPreTrackObjects();
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
    }
};

function MoveInGroupState(drawingObjects, majorObject, group, startX, startY)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
    this.group = group;
    this.startX = startX;
    this.startY = startY;
    this.bSamePos = true;
	this.hasObjectInSmartArt = false;

    var arr_x = [], arr_y = [];
    for(var i = 0; i < this.drawingObjects.arrTrackObjects.length; ++i)
    {
        var track = this.drawingObjects.arrTrackObjects[i];
	    const oOriginalObject = track.originalObject;
	    var transform = oOriginalObject.transform;
        arr_x.push(transform.TransformPointX(0, 0));
        arr_y.push(transform.TransformPointY(0, 0));
        arr_x.push(transform.TransformPointX(oOriginalObject.extX, 0));
        arr_y.push(transform.TransformPointY(oOriginalObject.extX, 0));
        arr_x.push(transform.TransformPointX(oOriginalObject.extX, oOriginalObject.extY));
        arr_y.push(transform.TransformPointY(oOriginalObject.extX, oOriginalObject.extY));
        arr_x.push(transform.TransformPointX(0, oOriginalObject.extY));
        arr_y.push(transform.TransformPointY(0, oOriginalObject.extY));
				if (!this.hasObjectInSmartArt)
				{
					this.hasObjectInSmartArt = oOriginalObject.isObjectInSmartArt();
				}
    }
    this.rectX = Math.min.apply(Math, arr_x);
    this.rectY = Math.min.apply(Math, arr_y);
    this.rectW = Math.max.apply(Math, arr_x);
    this.rectH = Math.max.apply(Math, arr_y);
}

MoveInGroupState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "move", bMarker: true};
        }
    },

    onMouseMove: MoveState.prototype.onMouseMove,

    onMouseUp: MoveState.prototype.onMouseUp
};


function PreRotateInGroupState(drawingObjects, group, majorObject)
{
    this.drawingObjects = drawingObjects;
    this.group = group;
    this.majorObject = majorObject;
}

PreRotateInGroupState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        this.drawingObjects.swapTrackObjects();
        this.drawingObjects.changeCurrentState(new RotateInGroupState(this.drawingObjects, this.group, this.majorObject))
    },

    onMouseUp: PreMoveInGroupState.prototype.onMouseUp
};

function RotateInGroupState(drawingObjects, group, majorObject)
{
    this.drawingObjects = drawingObjects;
    this.group = group;
    this.majorObject = majorObject;
}

RotateInGroupState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },

    onMouseMove: RotateState.prototype.onMouseMove,

    onMouseUp: MoveInGroupState.prototype.onMouseUp
};

function PreResizeInGroupState(drawingObjects, group, majorObject, cardDirection)
{
    this.drawingObjects = drawingObjects;
    this.group = group;
    this.majorObject = majorObject;
    this.cardDirection = cardDirection;
}

PreResizeInGroupState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        this.drawingObjects.swapTrackObjects();
        this.drawingObjects.changeCurrentState(new ResizeInGroupState(this.drawingObjects, this.group, this.majorObject, this.majorObject.getNumByCardDirection(this.cardDirection), this.cardDirection));
        this.drawingObjects.OnMouseMove(e, x, y, pageIndex);
    },

    onMouseUp: PreMoveInGroupState.prototype.onMouseUp
};

function ResizeInGroupState(drawingObjects, group, majorObject, handleNum, cardDirection)
{
    this.drawingObjects = drawingObjects;
    this.group = group;
    this.majorObject = majorObject;
    this.handleNum = handleNum;
    this.cardDirection = cardDirection;
}

ResizeInGroupState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },
    onMouseMove: ResizeState.prototype.onMouseMove,
    onMouseUp: MoveInGroupState.prototype.onMouseUp
};

function PreChangeAdjInGroupState(drawingObjects, group)
{
    this.drawingObjects = drawingObjects;
    this.group = group;
}

PreChangeAdjInGroupState.prototype =
{
    onMouseDown: function(e, x, y,pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId:  this.group.Get_Id(), bMarker: true, cursorType: "crosshair"};
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        this.drawingObjects.swapTrackObjects();
        this.drawingObjects.changeCurrentState(new ChangeAdjInGroupState(this.drawingObjects, this.group));
        this.drawingObjects.OnMouseMove(e, x, y, pageIndex);
    },

    onMouseUp: PreMoveInGroupState.prototype.onMouseUp
};

function ChangeAdjInGroupState(drawingObjects, group)
{
    this.drawingObjects = drawingObjects;
    this.group = group;
    this.majorObject = drawingObjects.arrTrackObjects[0].originalShape;
}

ChangeAdjInGroupState.prototype =
{
    onMouseDown: function(e, x, y,pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            return {objectId: this.majorObject.Get_Id(), cursorType: "crosshair", bMarker: true};
        }
    },

    onMouseMove: ChangeAdjState.prototype.onMouseMove,

    onMouseUp: MoveInGroupState.prototype.onMouseUp
};

function TextAddState(drawingObjects, majorObject, startX, startY, button)
{
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
    this.startX = startX;
    this.startY = startY;
    this.button = button;
    this.bIsSelectionEmpty = this.isSelectionEmpty();
}

TextAddState.prototype =
{
    isSelectionEmpty: function()
    {
        if(this.majorObject.getObjectType() === AscDFH.historyitem_type_GraphicFrame)
        {
            if(this.majorObject.graphicObject)
            {
                return this.majorObject.graphicObject.IsSelectionEmpty();
            }
            return true;
        }
        var oContent = this.majorObject.getDocContent && this.majorObject.getDocContent();
        if(oContent)
        {
            if(oContent.IsSelectionEmpty()) {
                var oParagraph = oContent.GetCurrentParagraph();
                if(oParagraph && oParagraph.IsEmptyWithBullet()) {
                    return true;
                }
            }
            return false;
        }
        return true;
    },

    checkSelectionEmpty: function()
    {
        this.bIsSelectionEmpty = this.isSelectionEmpty();
    },

    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
        {
            var sId = this.majorObject.Id;
            if(this.majorObject.chart
                && this.majorObject.chart.getObjectType
                && this.majorObject.chart.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
                sId = this.majorObject.chart.Id;
            }
            return {objectId: sId, cursorType: "text"};
        }
    },
    onMouseMove: function(e, x, y, pageIndex)
    {
        if(!e.IsLocked)
        {
            if(this.button === AscCommon.g_mouse_button_right)
            {
                return this.endState(e, x, y, pageIndex);
            }
            //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
            return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
        }
        if(AscFormat.isRealNumber(this.startX) && AscFormat.isRealNumber(this.startY))
        {
            if(Math.abs(this.startX - x) < 0.001 && Math.abs(this.startY - y) < 0.001)
            {
                return;
            }
            this.startX = undefined;
            this.startY = undefined;
        }
        this.majorObject.selectionSetEnd(e, x, y, pageIndex);
        if(!(this.majorObject.getObjectType() === AscDFH.historyitem_type_GraphicFrame && this.majorObject.graphicObject.Selection.Type2 === table_Selection_Border))
            this.drawingObjects.updateSelectionState();
        if(this.bIsSelectionEmpty !== this.isSelectionEmpty())
        {
            this.drawingObjects.drawingObjects.showDrawingObjects();
        }
        this.checkSelectionEmpty();
    },
    onMouseUp: function(e, x, y, pageIndex)
    {
        var oldCtrl;
        if(this.drawingObjects.isSlideShow())
        {
            oldCtrl = e.CtrlKey;
            e.CtrlKey = true;
        }
        this.majorObject.selectionSetEnd(e, x, y, pageIndex);
        if(this.drawingObjects.isSlideShow())
        {
            e.CtrlKey = oldCtrl;
        }
        return this.endState(e, x, y, pageIndex);
    },

    endState: function(e, x, y, pageIndex)
    {
        this.drawingObjects.updateSelectionState();
        this.drawingObjects.drawingObjects.sendGraphicObjectProps();
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
        this.drawingObjects.handleEventMode = HANDLE_EVENT_MODE_CURSOR;
        this.drawingObjects.noNeedUpdateCursorType = true;
        var oApi = this.drawingObjects.getEditorApi();
        var cursor_type = this.drawingObjects.curState.onMouseDown(e, x, y, pageIndex);
        if(cursor_type && cursor_type.hyperlink)
        {
            this.drawingObjects.drawingObjects.showDrawingObjects();
            if(this.drawingObjects.isSlideShow())
            {
                cursor_type.hyperlink.Visited = true;
                oApi.sync_HyperlinkClickCallback(cursor_type.hyperlink.Value);
            }
        }
        this.drawingObjects.noNeedUpdateCursorType = false;
        this.drawingObjects.handleEventMode = HANDLE_EVENT_MODE_HANDLE;
        if(oApi)
        {
            if(oApi.editorId === AscCommon.c_oEditorId.Presentation)
            {
                let oPresentation = oApi.WordControl && oApi.WordControl.m_oLogicDocument;
                if(oApi.isFormatPainterOn())
                {
                    this.drawingObjects.paragraphFormatPaste2();
                    if (oApi.canTurnOffFormatPainter())
                    {
                        oApi.sync_PaintFormatCallback(c_oAscFormatPainterState.kOff);
                        if(oPresentation)
                        {
                            oPresentation.OnMouseMove(e, x, y, pageIndex)
                        }
                    }
                }
                else if(oApi.isMarkerFormat)
                {
                    if(oPresentation)
                    {
                        if(oPresentation.HighlightColor)
                        {
                            var oC = oPresentation.HighlightColor;
                            oPresentation.SetParagraphHighlight(true, oC.r, oC.g, oC.b);
                        }
                        else
                        {
                            oPresentation.SetParagraphHighlight(false);
                        }
                        oApi.sync_MarkerFormatCallback(true);
                    }
                }
            }
            else if(oApi.editorId === AscCommon.c_oEditorId.Spreadsheet)
            {
                this.drawingObjects.checkFormatPainterOnMouseEvent();
            }
        }
    }
};


function SplineBezierState(drawingObjects, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.polylineFlag = true;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}
SplineBezierState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        this.drawingObjects.startTrackPos = {x: x, y: y, pageIndex: pageIndex};
        this.drawingObjects.clearTrackObjects();
        this.drawingObjects.addPreTrackObject(new AscFormat.Spline(this.drawingObjects, this.drawingObjects.getTheme(), null, null, null, pageIndex));
        this.drawingObjects.arrPreTrackObjects[0].path.push(new AscFormat.SplineCommandMoveTo(x, y));
        this.drawingObjects.changeCurrentState(new SplineBezierState33(this.drawingObjects, x, y,pageIndex, this.bAnimCustomPath, this.bReplace, this.bPreview));
        if(!this.bAnimCustomPath) {
            this.drawingObjects.checkChartTextSelection();
            this.drawingObjects.resetSelection();
        }
        this.drawingObjects.updateOverlay();
    },

    onMouseMove: function(e, X, Y, pageIndex)
    {
    },

    onMouseUp: function(e, X, Y, pageIndex)
    {
        if(Asc["editor"] && Asc["editor"].wb)
        {
            Asc["editor"].asc_endAddShape();
        }
        else if(editor && editor.sync_EndAddShape)
        {
            editor.sync_EndAddShape();
        }
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
    }
};


function SplineBezierState33(drawingObjects, startX, startY, pageIndex, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.polylineFlag = true;
    this.pageIndex = pageIndex;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}

SplineBezierState33.prototype =
{

    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        var startPos = this.drawingObjects.startTrackPos;
        if(startPos.x === x && startPos.y === y && startPos.pageIndex === pageIndex)
            return;

        var tr_x, tr_y;
        if(pageIndex === startPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, startPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }
        this.drawingObjects.swapTrackObjects();
        this.drawingObjects.arrTrackObjects[0].path.push(new AscFormat.SplineCommandLineTo(tr_x, tr_y));
        this.drawingObjects.changeCurrentState(new SplineBezierState2(this.drawingObjects, this.pageIndex, this.bAnimCustomPath, this.bReplace, this.bPreview));
        this.drawingObjects.updateOverlay();
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
    }
};

function SplineBezierState2(drawingObjects,pageIndex, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.polylineFlag = true;
    this.pageIndex = pageIndex;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}

SplineBezierState2.prototype =
{

    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        if(e.ClickCount >= 2)
        {
            this.bStart = true;
            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        var startPos = this.drawingObjects.startTrackPos;
        var tr_x, tr_y;
        if(pageIndex === startPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, startPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }

        this.drawingObjects.arrTrackObjects[0].path[1].changeLastPoint(tr_x, tr_y);
        this.drawingObjects.updateOverlay();
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(e.fromWindow)
        {
            var nOldClickCount = e.ClickCount;
            e.ClickCount = 2;
            this.onMouseDown(e, x, y, pageIndex);
            e.ClickCount = nOldClickCount;
            return;
        }
        if( e.ClickCount < 2)
        {
            var tr_x, tr_y;
            if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
            {
                tr_x = x;
                tr_y = y;
            }
            else
            {
                var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
                tr_x = tr_point.x;
                tr_y = tr_point.y;
            }
            this.drawingObjects.changeCurrentState(new SplineBezierState3(this.drawingObjects,tr_x, tr_y, this.pageIndex, this.bAnimCustomPath, this.bReplace, this.bPreview));
        }
    }
};

function SplineBezierState3(drawingObjects, startX, startY, pageIndex, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.startX = startX;
    this.startY = startY;
    this.polylineFlag = true;
    this.pageIndex = pageIndex;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}

SplineBezierState3.prototype =
{

    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        if(e.ClickCount >= 2)
        {
            this.bStart = true;

            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(x === this.startX && y === this.startY && pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            return;
        }

        var tr_x, tr_y;
        if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }

        var x0, y0, x1, y1, x2, y2, x3, y3, x4, y4, x5, y5, x6, y6;
        var spline = this.drawingObjects.arrTrackObjects[0];
        x0 = spline.path[0].x;
        y0 = spline.path[0].y;
        x3 = spline.path[1].x;
        y3 = spline.path[1].y;
        x6 = tr_x;
        y6 = tr_y;

        var vx = (x6 - x0)/6;
        var vy = (y6 - y0)/6;

        x2 = x3 - vx;
        y2 = y3 - vy;

        x4 = x3 + vx;
        y4 = y3 + vy;

        x1 = (x0 + x2)*0.5;
        y1 = (y0 + y2)*0.5;

        x5 = (x4 + x6)*0.5;
        y5 = (y4 + y6)*0.5;


        spline.path.length = 1;
        spline.path.push(new AscFormat.SplineCommandBezier(x1, y1, x2, y2, x3, y3));


        spline.path.push(new AscFormat.SplineCommandBezier(x4, y4, x5, y5, x6, y6));
        this.drawingObjects.updateOverlay();
        this.drawingObjects.changeCurrentState(new SplineBezierState4(this.drawingObjects, this.pageIndex, this.bAnimCustomPath, this.bReplace, this.bPreview));
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(e.fromWindow)
        {
            var nOldClickCount = e.ClickCount;
            e.ClickCount = 2;
            this.onMouseDown(e, x, y, pageIndex);
            e.ClickCount = nOldClickCount;
            return;
        }
        if(e.ClickCount >= 2)
        {
            this.bStart = true;
            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
    }
};


function SplineBezierState4(drawingObjects, pageIndex, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.polylineFlag = true;
    this.pageIndex = pageIndex;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}


SplineBezierState4.prototype =
{

    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        if(e.ClickCount >= 2)
        {
            this.bStart = true;
            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        var spline = this.drawingObjects.arrTrackObjects[0];
        var lastCommand = spline.path[spline.path.length-1];
        var preLastCommand = spline.path[spline.path.length-2];
        var x0, y0, x1, y1, x2, y2, x3, y3, x4, y4, x5, y5, x6, y6;
        if(spline.path[spline.path.length-3].id == 0)
        {
            x0 = spline.path[spline.path.length-3].x;
            y0 = spline.path[spline.path.length-3].y;
        }
        else
        {
            x0 = spline.path[spline.path.length-3].x3;
            y0 = spline.path[spline.path.length-3].y3;
        }

        x3 = preLastCommand.x3;
        y3 = preLastCommand.y3;

        var tr_x, tr_y;
        if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }
        x6 = tr_x;
        y6 = tr_y;

        var vx = (x6 - x0)/6;
        var vy = (y6 - y0)/6;

        x2 = x3 - vx;
        y2 = y3 - vy;

        x4 = x3 + vx;
        y4 = y3 + vy;

        x5 = (x4 + x6)*0.5;
        y5 = (y4 + y6)*0.5;

        if(spline.path[spline.path.length-3].id == 0)
        {
            preLastCommand.x1 = (x0 + x2)*0.5;
            preLastCommand.y1 = (y0 + y2)*0.5;
        }

        preLastCommand.x2 = x2;
        preLastCommand.y2 = y2;
        preLastCommand.x3 = x3;
        preLastCommand.y3 = y3;

        lastCommand.x1 = x4;
        lastCommand.y1 = y4;
        lastCommand.x2 = x5;
        lastCommand.y2 = y5;
        lastCommand.x3 = x6;
        lastCommand.y3 = y6;

        this.drawingObjects.updateOverlay();
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(e.fromWindow)
        {
            var nOldClickCount = e.ClickCount;
            e.ClickCount = 2;
            this.onMouseDown(e, x, y, pageIndex);
            e.ClickCount = nOldClickCount;
            return;
        }
        if(e.ClickCount < 2 )
        {
            var tr_x, tr_y;
            if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
            {
                tr_x = x;
                tr_y = y;
            }
            else
            {
                var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
                tr_x = tr_point.X;
                tr_y = tr_point.Y;
            }
            this.drawingObjects.changeCurrentState(new SplineBezierState5(this.drawingObjects, tr_x, tr_y, this.pageIndex, this.bAnimCustomPath, this.bReplace, this.bPreview));
        }
    }
};

function SplineBezierState5(drawingObjects, startX, startY,pageIndex, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.startX = startX;
    this.startY = startY;
    this.polylineFlag = true;
    this.pageIndex = pageIndex;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;

}

SplineBezierState5.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        if(e.ClickCount >= 2)
        {
            this.bStart = true;
            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(x === this.startX && y === this.startY && pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            return;
        }
        var spline = this.drawingObjects.arrTrackObjects[0];
        var lastCommand = spline.path[spline.path.length-1];
        var x0, y0, x1, y1, x2, y2, x3, y3, x4, y4, x5, y5, x6, y6;

        if(spline.path[spline.path.length-2].id == 0)
        {
            x0 = spline.path[spline.path.length-2].x;
            y0 = spline.path[spline.path.length-2].y;
        }
        else
        {
            x0 = spline.path[spline.path.length-2].x3;
            y0 = spline.path[spline.path.length-2].y3;
        }

        x3 = lastCommand.x3;
        y3 = lastCommand.y3;


        var tr_x, tr_y;
        if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }
        x6 = tr_x;
        y6 = tr_y;

        var vx = (x6 - x0)/6;
        var vy = (y6 - y0)/6;


        x2 = x3 - vx;
        y2 = y3 - vy;

        x1 = (x2+x1)*0.5;
        y1 = (y2+y1)*0.5;

        x4 = x3 + vx;
        y4 = y3 + vy;

        x5 = (x4 + x6)*0.5;
        y5 = (y4 + y6)*0.5;

        if(spline.path[spline.path.length-2].id == 0)
        {
            lastCommand.x1 = x1;
            lastCommand.y1 = y1;
        }
        lastCommand.x2 = x2;
        lastCommand.y2 = y2;


        spline.path.push(new AscFormat.SplineCommandBezier(x4, y4, x5, y5, x6, y6));
        this.drawingObjects.updateOverlay();
        this.drawingObjects.changeCurrentState(new SplineBezierState4(this.drawingObjects, this.pageIndex, this.bAnimCustomPath, this.bReplace, this.bPreview));
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(e.ClickCount >= 2 || e.fromWindow)
        {
            this.bStart = true;
            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
    }
};

function PolyLineAddState(drawingObjects, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.polylineFlag = true;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}

PolyLineAddState.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        this.drawingObjects.startTrackPos = {x: x, y: y, pageIndex:pageIndex};
        this.drawingObjects.clearTrackObjects();
        this.drawingObjects.addTrackObject(new AscFormat.PolyLine(this.drawingObjects, this.drawingObjects.getTheme(), null, null, null, pageIndex));
        this.drawingObjects.arrTrackObjects[0].addPoint(x, y);
        if(!this.bAnimCustomPath) {
            this.drawingObjects.checkChartTextSelection();
            this.drawingObjects.resetSelection();
        }
        this.drawingObjects.updateOverlay();
        var _min_distance = this.drawingObjects.convertPixToMM(1);
        this.drawingObjects.changeCurrentState(new PolyLineAddState2(this.drawingObjects, _min_distance, this.bAnimCustomPath, this.bReplace, this.bPreview));
    },

    onMouseMove: function()
    {},

    onMouseUp: function()
    {

        if(Asc["editor"] && Asc["editor"].wb)
        {
            Asc["editor"].asc_endAddShape();
        }
        else if(editor && editor.sync_EndAddShape)
        {
            editor.sync_EndAddShape();
        }
        this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
    }
};


function PolyLineAddState2(drawingObjects, minDistance, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.polylineFlag = true;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;

}
PolyLineAddState2.prototype =
{

    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
	    if(!e.IsLocked)
	    {
		    //todo: implement inheritance from AscCommon.CDrawingControllerStateBase
		    return AscCommon.CDrawingControllerStateBase.prototype.emulateMouseUp.call(this, e, x, y, pageIndex);
	    }
        var tr_x, tr_y;
        if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }
        this.drawingObjects.arrTrackObjects[0].tryAddPoint(tr_x, tr_y);
        this.drawingObjects.updateOverlay();
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.arrTrackObjects[0].canCreateShape())
        {
            this.bStart = true;
            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
        else
        {
            this.drawingObjects.clearTrackObjects();
            this.drawingObjects.updateOverlay();
            this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));

            if(Asc["editor"] && Asc["editor"].wb)
            {
                Asc["editor"].asc_endAddShape();
            }
            else if(editor && editor.sync_EndAddShape)
            {
                editor.sync_EndAddShape();
            }
        }

    }
};



function AddPolyLine2State(drawingObjects, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.polylineFlag = true;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}
AddPolyLine2State.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        this.drawingObjects.startTrackPos = {x: x, y: y, pageIndex : pageIndex};
        this.drawingObjects.checkChartTextSelection();
        if(!this.bAnimCustomPath) {
            this.drawingObjects.resetSelection();
        }
        this.drawingObjects.updateOverlay();
        this.drawingObjects.clearTrackObjects();
        this.drawingObjects.addPreTrackObject(new AscFormat.PolyLine(this.drawingObjects, this.drawingObjects.getTheme(), null, null, null, pageIndex));
        this.drawingObjects.arrPreTrackObjects[0].addPoint(x, y);
        this.drawingObjects.changeCurrentState(new AddPolyLine2State2(this.drawingObjects, x, y, this.bAnimCustomPath, this.bReplace, this.bPreview));
    },

    onMouseMove: function(e, x, y, pageIndex)
    {},

    onMouseUp: function(e, x, y, pageIndex)
    {
    }
};

function AddPolyLine2State2(drawingObjects, x, y, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;
    this.X = x;
    this.Y = y;
    this.polylineFlag = true;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;


}
AddPolyLine2State2.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};
        if(e.ClickCount > 1)
        {
            if(Asc["editor"] && Asc["editor"].wb)
            {
                Asc["editor"].asc_endAddShape();
            }
            else if(editor && editor.sync_EndAddShape)
            {
                editor.sync_EndAddShape();
            }
            this.drawingObjects.clearPreTrackObjects();
            this.drawingObjects.clearTrackObjects();
            this.drawingObjects.changeCurrentState(new NullState(this.drawingObjects));
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(this.X !== x || this.Y !== y || this.drawingObjects.startTrackPos.pageIndex !== pageIndex)
        {
            var tr_x, tr_y;
            if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
            {
                tr_x = x;
                tr_y = y;
            }
            else
            {
                var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
                tr_x = tr_point.X;
                tr_y = tr_point.Y;
            }
            this.drawingObjects.swapTrackObjects();
            this.drawingObjects.arrTrackObjects[0].tryAddPoint(tr_x, tr_y);
            this.drawingObjects.changeCurrentState(new AddPolyLine2State3(this.drawingObjects, this.bAnimCustomPath, this.bReplace, this.bPreview));
        }
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
    }
};

function AddPolyLine2State3(drawingObjects, bAnimCustomPath, bReplace, bPreview)
{
    this.drawingObjects = drawingObjects;

    this.lastX = -1000;
    this.lastY = -1000;


    this.polylineFlag = true;
    this.bAnimCustomPath = bAnimCustomPath;
    this.bReplace = bReplace;
    this.bPreview = bPreview;
}
AddPolyLine2State3.prototype =
{
    onMouseDown: function(e, x, y, pageIndex)
    {
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: "1", bMarker: true, cursorType: "crosshair"};


        this.lastX = x;
        this.lastY = y;
        var tr_x, tr_y;
        if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }
        if(e.ClickCount > 1)
        {

            this.bStart = true;
            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
        else
        {
            var oTrack = this.drawingObjects.arrTrackObjects[0];
            oTrack.replaceLastPoint(tr_x, tr_y, false);
            oTrack.addPoint(tr_x, tr_y, true);
        }
    },

    onMouseMove: function(e, x, y, pageIndex)
    {
        if(AscFormat.fApproxEqual(x, this.lastX) && AscFormat.fApproxEqual(y, this.lastY)) {
            return;
        }
        var tr_x, tr_y;
        if(pageIndex === this.drawingObjects.startTrackPos.pageIndex)
        {
            tr_x = x;
            tr_y = y;
        }
        else
        {
            var tr_point = this.drawingObjects.getDrawingDocument().ConvertCoordsToAnotherPage(x, y, pageIndex, this.drawingObjects.startTrackPos.pageIndex);
            tr_x = tr_point.X;
            tr_y = tr_point.Y;
        }

        var oTrack = this.drawingObjects.arrTrackObjects[0];
        if(!e.IsLocked && oTrack.getPointsCount() > 1)
        {
            oTrack.replaceLastPoint(tr_x, tr_y, true);
        }
        else
        {
            oTrack.tryAddPoint(tr_x, tr_y);
        }
        this.drawingObjects.updateOverlay();
        this.lastX = x;
        this.lastY = y;
    },

    onMouseUp: function(e, x, y, pageIndex)
    {
        this.lastX = x;
        this.lastY = y;
        if(e.fromWindow)
        {
            var nOldClickCount = e.ClickCount;
            e.ClickCount = 2;
            this.onMouseDown(e, x, y, pageIndex);
            e.ClickCount = nOldClickCount;
            return;
        }
        if(e.ClickCount > 1)
        {

            this.bStart = true;

            this.pageIndex = this.drawingObjects.startTrackPos.pageIndex;
            StartAddNewShape.prototype.onMouseUp.call(this, e, x, y, pageIndex);
        }
    }
};

function TrackTextState(drawingObjects, majorObject, x, y) {
    this.drawingObjects = drawingObjects;
    this.majorObject = majorObject;
    this.startX = x;
    this.startY = y;
    this.bMove = false;
}
    TrackTextState.prototype.onMouseDown = function(e, x, y){
        if(this.drawingObjects.handleEventMode === HANDLE_EVENT_MODE_CURSOR)
            return {objectId: this.majorObject.Id, bMarker: true, cursorType: "default"};
        return null;
    };
    TrackTextState.prototype.onMouseMove = function(e, x, y){
        if(Math.abs(x - this.startX) > MOVE_DELTA || Math.abs(y - this.startY) > MOVE_DELTA)
        {
            this.bMove = true;
            this.drawingObjects.getDrawingDocument().StartTrackText();
        }
    };
    TrackTextState.prototype.onMouseUp   = function(e, x, y, pageIndex){
        if(!this.bMove)
        {
            this.majorObject.selectionSetStart(e, x, y, 0);
            this.majorObject.selectionSetEnd(e, x, y, 0);
            this.drawingObjects.updateSelectionState();
            this.drawingObjects.drawingObjects.sendGraphicObjectProps();
        }
        this.drawingObjects.changeCurrentState(new AscFormat.NullState(this.drawingObjects));

    };

    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].MOVE_DELTA = MOVE_DELTA;
    window['AscFormat'].SNAP_DISTANCE = SNAP_DISTANCE;
    window['AscFormat'].StartAddNewShape = StartAddNewShape;
    window['AscFormat'].NullState = NullState;
    window['AscFormat'].SlicerState = SlicerState;
    window['AscFormat'].PreChangeAdjState = PreChangeAdjState;
    window['AscFormat'].PreRotateState = PreRotateState;
    window['AscFormat'].RotateState = RotateState;
    window['AscFormat'].PreResizeState = PreResizeState;
    window['AscFormat'].PreMoveState = PreMoveState;
    window['AscFormat'].MoveState = MoveState;
    window['AscFormat'].PreMoveInGroupState = PreMoveInGroupState;
    window['AscFormat'].MoveInGroupState = MoveInGroupState;
    window['AscFormat'].PreRotateInGroupState = PreRotateInGroupState;
    window['AscFormat'].PreResizeInGroupState = PreResizeInGroupState;
    window['AscFormat'].PreChangeAdjInGroupState = PreChangeAdjInGroupState;
    window['AscFormat'].TextAddState = TextAddState;
    window['AscFormat'].SplineBezierState = SplineBezierState;
    window['AscFormat'].PolyLineAddState = PolyLineAddState;
    window['AscFormat'].AddPolyLine2State = AddPolyLine2State;
    window['AscFormat'].TrackTextState = TrackTextState;
    window['AscFormat'].checkEmptyPlaceholderContent = checkEmptyPlaceholderContent;
})(window);
