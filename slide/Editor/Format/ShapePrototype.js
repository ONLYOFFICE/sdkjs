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

CShape.prototype.IsUseInDocument = function()
{
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
    if(this.parent && this.parent.cSld){
        var aSpTree = this.parent.cSld.spTree;
        for(var i = 0; i < aSpTree.length; ++i){
            if(aSpTree[i] === this) {
                if(this.parent.IsUseInDocument) {
                    return this.parent.IsUseInDocument();
                }
                return true;
            }
        }
    }
    return false;
};

CShape.prototype.getDrawingObjectsController = function()
{
    if(this.parent && (AscFormat.isSlideLikeObject(this.parent) ||  this.parent.getObjectType() === AscDFH.historyitem_type_Notes))
    {
        return this.parent.graphicObjects;
    }
    return null;
};


function editorAddToDrawingObjects(oGraphicObject, pos, type)
{
    if(Asc.editor.isDrawSlideshowAnnotations())
    {
        let oAnnots = Asc.editor.getAnnotations();
        if(oAnnots)
        {
            oAnnots.addInk(oGraphicObject);
        }
        return;
    }
    if(oGraphicObject.parent && oGraphicObject.parent.cSld && oGraphicObject.parent.cSld.spTree)
    {
        if(oGraphicObject.signatureLine && oGraphicObject.setSignature)
        {
            oGraphicObject.setSignature(null);
        }
        oGraphicObject.parent.shapeAdd(pos, oGraphicObject);
    }
}
    function editorDeleteDrawingBase(oGraphicObject, bCheckPlaceholder)
    {
        let oSlide = oGraphicObject.parent;
        if(Asc.editor.isDrawSlideshowAnnotations())
        {
            let oAnnots = Asc.editor.getAnnotations();
            if(oAnnots)
            {
                oAnnots.eraseInk(oGraphicObject);
            }
            return;
        }
        if(AscFormat.isSlideLikeObject(oSlide))
        {
            let pos = oSlide.removeFromSpTreeById(oGraphicObject.Id);
            let phType = oGraphicObject.getPlaceholderType();
            if (oSlide.getObjectType() === AscDFH.historyitem_type_Slide)
            {
                if(bCheckPlaceholder && oGraphicObject.isPlaceholder() && !oGraphicObject.isEmptyPlaceholder()
                    && phType !== AscFormat.phType_hdr && phType !== AscFormat.phType_ftr
                    && phType !== AscFormat.phType_sldNum && phType !== AscFormat.phType_dt)
                {
                    let hierarchy = oGraphicObject.getHierarchy();
                    if(hierarchy[0])
                    {
                        var copy = hierarchy[0].copy(undefined);
                        copy.setParent(oSlide);
                        copy.addToDrawingObjects(pos);
                        var doc_content = copy.getDocContent && copy.getDocContent();
                        if(doc_content)
                        {
                            doc_content.SetApplyToAll(true);
                            doc_content.Remove(-1);
                            doc_content.SetApplyToAll(false);
                        }
                    }
                }
            }
            return pos;
        }
        return -1;
    }

CShape.prototype.setRecalculateInfo = function()
{
    this.recalcInfo =
    {
        recalculateContent:        true,
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
        recalculateContent2: true,
        oContentMetrics: null

    };
    this.compiledStyles = [];
    this.lockType = AscCommon.c_oAscLockTypes.kLockTypeNone;
};
CShape.prototype.recalcContent = function()
{
    this.recalcInfo.recalculateContent = true;
};
CShape.prototype.recalcContent2 = function()
{
    this.recalcInfo.recalculateContent2 = true;
};

CShape.prototype.getDrawingDocument = function()
{
    return editor.WordControl.m_oLogicDocument.DrawingDocument;
};

CShape.prototype.getTextArtPreviewManager = function()
{
    return editor.textArtPreviewManager;
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
    AscCommon.History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
};
CShape.prototype.getSlideIndex = function()
{
    return this.Get_StartPage_Absolute();
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
    this.recalcGeometry();
    this.recalcBounds();
    this.recalcTransform();
    this.recalcContent();
    this.recalcContent2();
    this.recalcTransformText();
    this.addToRecalculate();
};
CShape.prototype.handleUpdateTheme = function()
{
    this.setRecalculateInfo(); //TODO

    if(this.isPlaceholder()
        && !( this.spPr && this.spPr.xfrm &&
        (this.isGroupObject && this.isGroupObject() && this.spPr.xfrm.isNotNullForGroup() || this.getObjectType() !== AscDFH.historyitem_type_GroupShape && this.spPr.xfrm.isNotNull() ) ))
    {
        this.recalcTransform();
        this.recalcGeometry()
    }
    var content = this.getDocContent && this.getDocContent();
    if(content)
    {
        content.Recalc_AllParagraphs_CompiledPr();
    }
    this.recalcContent && this.recalcContent();
    this.recalcFill && this.recalcFill();
    this.recalcLine && this.recalcLine();
    this.recalcPen && this.recalcPen();
    this.recalcBrush && this.recalcBrush();
    this.recalcStyle && this.recalcStyle();
    this.recalcInfo.recalculateTextStyles && (this.recalcInfo.recalculateTextStyles = [true, true, true, true, true, true, true, true, true]);
    this.recalcBounds && this.recalcBounds();
    this.handleTitlesAfterChangeTheme && this.handleTitlesAfterChangeTheme();
    if(Array.isArray(this.spTree))
    {
        for(var i = 0; i < this.spTree.length; ++i)
        {
            this.spTree[i].handleUpdateTheme();
        }
    }

};
CShape.prototype.handleUpdateRot = function()
{
    this.recalcTransform();
    if(this.txBody && this.txBody.bodyPr && this.txBody.bodyPr.upright)
    {
        this.recalcContent();
        this.recalcContent2();
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
    this.recalcContent2();
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
    this.recalcContent2();
    this.recalcTransformText();
    this.addToRecalculate();
};
CShape.prototype.convertPixToMM = function(pix)
{
    return editor.WordControl.m_oLogicDocument.DrawingDocument.GetMMPerDot(pix);
};
CShape.prototype.getCanvasContext = function()
{
    return editor.WordControl.m_oLogicDocument.DrawingDocument.CanvasHitContext;
};
CShape.prototype.getCompiledStyle = function()
{
    if(this.style){
        return this.style;
    }
    var hierarchy = this.getHierarchy();
    for (var i = 0; i < hierarchy.length; ++i) {
        if (hierarchy[i] && hierarchy[i].style) {
            return hierarchy[i].style;
        }
    }
    return null;
};
CShape.prototype.getParentObjects = function ()
{
    let oParent;
    if(this.parent)
    {
        oParent = this.parent;
    }
    else
    {
        var oPresentation = editor.WordControl.m_oLogicDocument;
        if(oPresentation)
        {
            oParent = oPresentation.GetCurrentSlide();
        }
    }
    if(oParent)
    {
        switch(oParent.getObjectType())
        {
            case AscDFH.historyitem_type_Slide:
            {
                return {
                    presentation: editor.WordControl.m_oLogicDocument,
                    slide: oParent,
                    layout: oParent.Layout,
                    master: oParent.Layout ? oParent.Layout.Master : null,
                    theme: this.themeOverride ? this.themeOverride : (oParent.Layout && oParent.Layout.Master ? oParent.Layout.Master.Theme : null)
                };
            }
            case AscDFH.historyitem_type_SlideLayout:
            {
                return {
                    presentation: editor.WordControl.m_oLogicDocument,
                    slide: null,
                    layout: oParent,
                    master: oParent.Master,
                    theme: this.themeOverride ? this.themeOverride : (oParent.Master ? oParent.Master.Theme : null)
                };
            }
            case AscDFH.historyitem_type_SlideMaster:
            {
                return {
                    presentation: editor.WordControl.m_oLogicDocument,
                    slide: null,
                    layout: null,
                    master: oParent,
                    theme: this.themeOverride ? this.themeOverride : oParent.Theme
                };
            }
            case AscDFH.historyitem_type_Notes:
            {
                return {

                    presentation: editor.WordControl.m_oLogicDocument,
                    slide: null,
                    layout: null,
                    master: oParent.Master,
                    theme: this.themeOverride ? this.themeOverride : (oParent.Master ? oParent.Master.Theme : null),
                    notes: oParent
                }
            }
            case AscDFH.historyitem_type_NotesMaster:
            {
                return {
                    presentation: editor.WordControl.m_oLogicDocument,
                    slide: null,
                    layout: null,
                    master: oParent,
                    theme: this.themeOverride ? this.themeOverride : oParent.Theme,
                    notes: null
                }
            }
            case AscDFH.historyitem_type_RelSizeAnchor:
            case AscDFH.historyitem_type_AbsSizeAnchor:
            {
                if(oParent.parent)
                {
                    return oParent.parent.getParentObjects()
                }
                break;
            }
        }
    }
    return { slide: null, layout: null, master: null, theme: null};
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
    if(this.bDeleted || !this.parent)
        return;

    if(this.parent.getObjectType() === AscDFH.historyitem_type_Notes){
        return;
    }
    var check_slide_placeholder = !this.isPlaceholder() || (this.parent );
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

        if (this.recalcInfo.recalculateContent && check_slide_placeholder) {
            this.recalcInfo.oContentMetrics = this.recalculateContent();
            this.recalcInfo.recalculateContent = false;
        }
        if (this.recalcInfo.recalculateContent2 && check_slide_placeholder) {
            this.recalculateContent2();
            this.recalcInfo.recalculateContent2 = false;
        }

        if (this.recalcInfo.recalculateTransformText && check_slide_placeholder) {
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
    boundsChecker.DO_NOT_DRAW_ANIM_LABEL = true;
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
};
    CShape.prototype.getEditorType = function()
    {
        return 0;
    };

CShape.prototype.recalculateContent = function()
{
    var content = this.getDocContent();
    if(content)
    {
        var body_pr = this.getBodyPr();

        var oRecalcObject = this.recalculateDocContent(content, body_pr);

        this.contentWidth = oRecalcObject.w;
        this.contentHeight = oRecalcObject.contentH;
        if(this.recalcInfo.recalcTitle)
        {
            this.recalcInfo.bRecalculatedTitle = true;
            this.recalcInfo.recalcTitle = null;


            var oTextWarpContent = this.checkTextWarp(content, body_pr, oRecalcObject.textRectW + oRecalcObject.correctW, oRecalcObject.textRectH + oRecalcObject.correctH, true, false);
            this.txWarpStructParamarks = oTextWarpContent.oTxWarpStructParamarksNoTransform;
            this.txWarpStruct = oTextWarpContent.oTxWarpStructNoTransform;

            this.txWarpStructParamarksNoTransform = oTextWarpContent.oTxWarpStructParamarksNoTransform;
            this.txWarpStructNoTransform = oTextWarpContent.oTxWarpStructNoTransform;
        }
        else
        {
            var oTextWarpContent = this.checkTextWarp(content, body_pr, oRecalcObject.textRectW + oRecalcObject.correctW, oRecalcObject.textRectH + oRecalcObject.correctH, true, true);
            this.txWarpStructParamarks = oTextWarpContent.oTxWarpStructParamarks;
            this.txWarpStruct = oTextWarpContent.oTxWarpStruct;

            this.txWarpStructParamarksNoTransform = oTextWarpContent.oTxWarpStructParamarksNoTransform;
            this.txWarpStructNoTransform = oTextWarpContent.oTxWarpStructNoTransform;
        }
        return oRecalcObject;
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
    var parent_objects = this.getParentObjects();
    if(parent_objects.slide && parent_objects.slide.clrMap)
    {
        return parent_objects.slide.clrMap;
    }
    else if(parent_objects.layout && parent_objects.layout.clrMap)
    {
        return parent_objects.layout.clrMap;
    }
    else if(parent_objects.master && parent_objects.master.clrMap)
    {
        return parent_objects.master.clrMap;
    }
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
CShape.prototype.Get_Numbering =  function()
{
    return AscWord.DEFAULT_NUMBERING;
};
CShape.prototype.getIsSingleBody = function(x, y)
{
    if(!this.isPlaceholder())
        return false;
    var ph_type = this.getPlaceholderType();
    if(ph_type !== AscFormat.phType_body && ph_type !== null)
        return false;
    if(this.parent && this.parent.cSld && Array.isArray(this.parent.cSld.spTree))
    {
        var sp_tree = this.parent.cSld.spTree;
        for(var i = 0; i < sp_tree.length; ++i)
        {
            if(sp_tree[i] !== this && sp_tree[i].getPlaceholderType && sp_tree[i].getPlaceholderType() === AscFormat.phType_body)
                return false;
        }
    }
    return true;
};

CShape.prototype.Set_CurrentElement = function(bUpdate, pageIndex, bNoTextSelection){
    if(this.parent && this.parent.graphicObjects){
        var drawing_objects = this.parent.graphicObjects;
        
        if(bNoTextSelection !== true) 
        {
            this.SetControllerTextSelection(drawing_objects, 0);
        }
        else 
        {
            let oSelector;
            if (this.group)
            {
                let main_group = this.group.getMainGroup();
                oSelector = main_group;
            }
            else
            {
                oSelector = drawing_objects;
            }
            oSelector.resetSelection();
            oSelector.selectObject(this, 0);
        }
        var nSlideNum;
        if(this.parent instanceof AscCommonSlide.CNotes){
            editor.WordControl.m_oLogicDocument.FocusOnNotes = true;
            if(this.parent.slide){
                nSlideNum = this.parent.slide.num;
                this.parent.slide.graphicObjects.resetSelection();
            }
            else{
                nSlideNum = 0;
            }
        }
        else{
            nSlideNum = this.getParentNum();
            editor.WordControl.m_oLogicDocument.FocusOnNotes = false;
        }
        if(editor.WordControl.m_oLogicDocument.CurPage !== nSlideNum){
            editor.WordControl.m_oLogicDocument.Set_CurPage(nSlideNum);
            editor.WordControl.GoToPage(nSlideNum);
            if(this.parent instanceof AscCommonSlide.CNotes){
                editor.WordControl.m_oLogicDocument.FocusOnNotes = true;
            }
        }
		if (bUpdate && editor.WordControl.m_oLogicDocument) {
			editor.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
			editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		}
    }
};

CShape.prototype.OnContentReDraw = function(){
    if(AscCommonSlide){
        var oPresentation = editor.WordControl.m_oLogicDocument;
        if(this.parent instanceof AscCommonSlide.Slide) {
            oPresentation.DrawingDocument.OnRecalculateSlide(this.getParentNum());
        }
        else if(this.parent instanceof AscCommonSlide.CNotes) {
            var oCurSlide = oPresentation.Slides[oPresentation.CurPage];
            if(oCurSlide && oCurSlide.notes === this.parent){
                oPresentation.DrawingDocument.Notes_OnRecalculate(oPresentation.CurPage, oCurSlide.NotesWidth, oCurSlide.getNotesHeight());
            }
        }
    }
};

    CShape.prototype.IsThisElementCurrent = function()
    {
        if(this.parent && this.parent.graphicObjects)
        {
            if(this.group)
            {
                var main_group = this.group.getMainGroup();
                return main_group.selection.textSelection === this;
            }
            else
            {
                if(this.parent.getObjectType && this.parent.getObjectType() === AscDFH.historyitem_type_Notes)
                {
                    if(editor.WordControl.m_oLogicDocument.FocusOnNotes && this.parent.slide && this.parent.slide.num === editor.WordControl.m_oLogicDocument.CurPage)
                    {
                        return this.parent.graphicObjects.selection.textSelection === this;
                    }
                }
                else
                {
                    return this.parent.graphicObjects.selection.textSelection === this;
                }
            }
        }
        return false;
    };


    CShape.prototype.Get_StartPage_Absolute = function () {
        if(this.getParentObjects) {
            let oParents = this.getParentObjects();
            if(oParents && oParents.presentation) {
                if(oParents.slide) {
                    return oParents.slide.num;
                }
                if(oParents.notes && oParents.notes.slide) {
                    return oParents.notes.slide.num;
                }
                if(oParents.layout) {
                    return oParents.presentation.GetSlideIndex(oParents.layout);
                }
                if(oParents.master) {
                    return oParents.presentation.GetSlideIndex(oParents.master);
                }
            }
        }
        return 0;
    };
    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].editorDeleteDrawingBase = editorDeleteDrawingBase;
    window['AscFormat'].editorAddToDrawingObjects = editorAddToDrawingObjects;
})(window);
