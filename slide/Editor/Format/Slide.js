/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

// Import
var g_oTableId = AscCommon.g_oTableId;
var History = AscCommon.History;




function CChangesDrawingsContentComments(Class, Type, Pos, Items, isAdd){
	AscDFH.CChangesDrawingsContent.call(this, Class, Type, Pos, Items, isAdd);
}
CChangesDrawingsContentComments.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
CChangesDrawingsContentComments.prototype.constructor = CChangesDrawingsContentComments;
CChangesDrawingsContentComments.prototype.addToInterface = function(){
    for(var i = 0; i < this.Items.length; ++i){
        var oComment = this.Items[i];
        oComment.Data.bDocument = !(oComment.Parent && (oComment.Parent.slide instanceof Slide));
        editor.sync_AddComment(oComment.Get_Id(), oComment.Data);
    }
};
CChangesDrawingsContentComments.prototype.removeFromInterface = function(){
    for(var i = 0; i < this.Items.length; ++i){
        editor.sync_RemoveComment(this.Items[i].Get_Id());
    }
};
CChangesDrawingsContentComments.prototype.Undo = function(){
	AscDFH.CChangesDrawingsContent.prototype.Undo.call(this);
    if(this.IsAdd()){
        this.removeFromInterface();
    }
    else{
        this.addToInterface();
    }
};
CChangesDrawingsContentComments.prototype.Redo = function(){
	AscDFH.CChangesDrawingsContent.prototype.Redo.call(this);
    if(this.IsAdd()){
        this.addToInterface();
    }
    else{
        this.removeFromInterface();
    }
};

CChangesDrawingsContentComments.prototype.Load = function(){
	AscDFH.CChangesDrawingsContent.prototype.Load.call(this);
    if(this.IsAdd()){
        this.addToInterface();
    }
    else{
        this.removeFromInterface();
    }
};

AscDFH.CChangesDrawingsContentComments = CChangesDrawingsContentComments;


AscDFH.changesFactory[AscDFH.historyitem_SlideSetLocks             ] = AscDFH.CChangesDrawingSlideLocks;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetComments          ] = AscDFH.CChangesDrawingsObject    ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetShow              ] = AscDFH.CChangesDrawingsBool      ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetShowPhAnim        ] = AscDFH.CChangesDrawingsBool      ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetShowMasterSp      ] = AscDFH.CChangesDrawingsBool      ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetLayout            ] = AscDFH.CChangesDrawingsObject    ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetNum               ] = AscDFH.CChangesDrawingsLong      ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetTransition        ] = AscDFH.CChangesDrawingsObjectNoId;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetSize              ] = AscDFH.CChangesDrawingsObjectNoId;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetBg                ] = AscDFH.CChangesDrawingsObjectNoId;
AscDFH.changesFactory[AscDFH.historyitem_SlideAddToSpTree          ] = AscDFH.CChangesDrawingsContentPresentation   ;
AscDFH.changesFactory[AscDFH.historyitem_SlideRemoveFromSpTree     ] = AscDFH.CChangesDrawingsContentPresentation   ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetCSldName          ] = AscDFH.CChangesDrawingsString    ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetClrMapOverride    ] = AscDFH.CChangesDrawingsObject    ;
AscDFH.changesFactory[AscDFH.historyitem_PropLockerSetId           ] = AscDFH.CChangesDrawingsString    ;
AscDFH.changesFactory[AscDFH.historyitem_SlideCommentsAddComment   ] = AscDFH.CChangesDrawingsContentComments   ;
AscDFH.changesFactory[AscDFH.historyitem_SlideCommentsRemoveComment] = AscDFH.CChangesDrawingsContentComments   ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetNotes             ] = AscDFH.CChangesDrawingsObject   ;
AscDFH.changesFactory[AscDFH.historyitem_SlideSetTiming      ] = AscDFH.CChangesDrawingsObject   ;


AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetComments          ] = function(oClass, value){oClass.slideComments = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetShow              ] = function(oClass, value){oClass.show = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetShowPhAnim        ] = function(oClass, value){oClass.showMasterPhAnim = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetShowMasterSp      ] = function(oClass, value){oClass.showMasterSp = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetLayout            ] = function(oClass, value){oClass.Layout = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetTiming            ] = function(oClass, value){oClass.timing = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetNum               ] = function(oClass, value){oClass.num = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetTransition        ] = function(oClass, value){oClass.transition = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetSize              ] = function(oClass, value){oClass.Width = value.a; oClass.Height = value.b;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetBg                ] = function(oClass, value, FromLoad){
    oClass.cSld.Bg = value;
    if(FromLoad){
        var Fill;
        if(oClass.cSld.Bg && oClass.cSld.Bg.bgPr && oClass.cSld.Bg.bgPr.Fill)
        {
            Fill = oClass.cSld.Bg.bgPr.Fill;
        }
        if(typeof AscCommon.CollaborativeEditing !== "undefined")
        {
            if(Fill && Fill.fill && Fill.fill.type === Asc.c_oAscFill.FILL_TYPE_BLIP && typeof Fill.fill.RasterImageId === "string" && Fill.fill.RasterImageId.length > 0)
            {
                AscCommon.CollaborativeEditing.Add_NewImage(Fill.fill.RasterImageId);
            }
        }
    }
};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetCSldName          ] = function(oClass, value){oClass.cSld.name = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetClrMapOverride    ] = function(oClass, value){oClass.clrMap = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_PropLockerSetId           ] = function(oClass, value){oClass.objectId = value;};
AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetNotes             ] = function(oClass, value){oClass.notes = value;};



AscDFH.drawingContentChanges[AscDFH.historyitem_SlideAddToSpTree          ] =
AscDFH.drawingContentChanges[AscDFH.historyitem_SlideRemoveFromSpTree     ] = function(oClass){return oClass.cSld.spTree;};
AscDFH.drawingContentChanges[AscDFH.historyitem_SlideCommentsAddComment   ] =
AscDFH.drawingContentChanges[AscDFH.historyitem_SlideCommentsRemoveComment] = function(oClass){return oClass.comments;};

AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideSetSize              ] = AscFormat.CDrawingBaseCoordsWritable;
AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideSetBg                ] = AscFormat.CBg;
AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideSetTransition        ] = Asc.CAscSlideTransition;

function Slide(presentation, slideLayout, slideNum)
{
    AscFormat.CBaseFormatObject.call(this);
    this.kind = AscFormat.TYPE_KIND.SLIDE;

    this.presentation = editor && editor.WordControl && editor.WordControl.m_oLogicDocument;
    this.graphicObjects = new AscFormat.DrawingObjectsController(this);
    this.cSld = new AscFormat.CSld(this);
    this.collaborativeMarks = new CRunCollaborativeMarks();
    this.clrMap = null; // override ClrMap

    this.show = true;
    this.showMasterPhAnim = false;
    this.showMasterSp = true;

    this.backgroundFill = null;

    this.notes = null;

    this.transition = new Asc.CAscSlideTransition();
    this.transition.setDefaultParams();

    this.timing = null;

    this.recalcInfo =
    {
        recalculateBackground: true,
        recalculateSpTree: true
    };
    this.Width = 254;
    this.Height = 190.5;

    this.searchingArray = [];  // массив объектов для селекта
    this.selectionArray = [];  // массив объектов для поиска


    this.writecomments = [];

    this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)

    this.commentX = 0;
    this.commentY = 0;


    this.deleteLock     = null;
    this.backgroundLock = null;
    this.timingLock     = null;
    this.transitionLock = null;
    this.layoutLock     = null;
    this.showLock       = null;

    this.Lock = new AscCommon.CLock();


    this.notesShape = null;

    this.lastLayoutType = null;
    this.lastLayoutMatchingName = null;
    this.lastLayoutName = null;

    this.NotesWidth = -10.0;

    this.animationPlayer = null;

    if(presentation)
    {
        this.Width = presentation.GetWidthMM();
        this.Height = presentation.GetHeightMM();
        this.setSlideComments(new SlideComments(this));
        this.setLocks(new PropLocker(this.Id), new PropLocker(this.Id), new PropLocker(this.Id), new PropLocker(this.Id), new PropLocker(this.Id), new PropLocker(this.Id), new PropLocker(this.Id));
    }

    if(slideLayout)
    {
        this.setLayout(slideLayout);
    }
    if(typeof slideNum === "number")
    {
        this.setSlideNum(slideNum);
    }
}
AscFormat.InitClass(Slide, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_Slide);

    Slide.prototype.getDrawingDocument = function()
    {
        return editor.WordControl.m_oLogicDocument.DrawingDocument;
    };
    Slide.prototype.Reassign_ImageUrls = function(images_rename){
	    this.cSld.forEachSp(function(oSp) {
		    oSp.Reassign_ImageUrls(images_rename);
	    });
        if(this.cSld.Bg &&
            this.cSld.Bg.bgPr &&
            this.cSld.Bg.bgPr.Fill &&
            this.cSld.Bg.bgPr.Fill.fill instanceof AscFormat.CBlipFill &&
            typeof this.cSld.Bg.bgPr.Fill.fill.RasterImageId === "string" &&
            images_rename[this.cSld.Bg.bgPr.Fill.fill.RasterImageId])
        {
            let oBg = this.cSld.Bg.createFullCopy();
            oBg.bgPr.Fill.fill.RasterImageId = images_rename[oBg.bgPr.Fill.fill.RasterImageId];
            this.changeBackground(oBg);
        }
    };
    Slide.prototype.Clear_CollaborativeMarks = function()
    {
        this.collaborativeMarks.Clear();
    };
    Slide.prototype.createDuplicate = function(IdMap, bCacheImage)
    {
        var oIdMap = IdMap || {};
        var oPr = new AscFormat.CCopyObjectProperties();
        oPr.idMap = oIdMap;
        oPr.cacheImage = bCacheImage !== false;
        var copy = new Slide(this.presentation, this.Layout, 0), i;
        if(typeof this.cSld.name === "string" && this.cSld.name.length > 0)
        {
            copy.setCSldName(this.cSld.name);
        }
        if(this.cSld.Bg)
        {
            copy.changeBackground(this.cSld.Bg.createFullCopy());
        }

	    this.cSld.forEachSp(function(oSp) {
		    let oSpCopy = oSp.copy(oPr);
		    oIdMap[oSp.Id] = oSpCopy.Id;
		    copy.shapeAdd(copy.cSld.spTree.length, oSpCopy);
	    });

        if(this.clrMap)
        {
            copy.setClMapOverride(this.clrMap.createDuplicate());
        }
        if(AscFormat.isRealBool(this.show))
        {
            copy.setShow(this.show);
        }
        if(AscFormat.isRealBool(this.showMasterPhAnim))
        {
            copy.setShowPhAnim(this.showMasterPhAnim);
        }
        if(AscFormat.isRealBool(this.showMasterSp))
        {
            copy.setShowMasterSp(this.showMasterSp);
        }

        copy.applyTransition(this.transition.createDuplicate());
        copy.setSlideSize(this.Width, this.Height);

        if(this.notes){
            copy.setNotes(this.notes.createDuplicate());
        }
        if(this.slideComments){
            if(!copy.slideComments) {
                copy.setSlideComments(new SlideComments(copy));
            }
            var aComments = this.slideComments.comments;
            for(i = 0; i < aComments.length; ++i){
                copy.slideComments.addComment(aComments[i].createDuplicate(copy.slideComments, true));
            }
        }

        if(!this.recalcInfo.recalculateBackground && !this.recalcInfo.recalculateSpTree)
        {
            if(!oPr || false !== oPr.cacheImage) {
                copy.cachedImage = this.getBase64Img();
            }
        }
        if(this.timing) {
            copy.setTiming(this.timing.createDuplicate(oIdMap));
        }
        return copy;
    };
    Slide.prototype.handleAllContents = function(fCallback){
        this.cSld.handleAllContents(fCallback);
        if(this.notesShape){
            this.notesShape.handleAllContents(fCallback);
        }
    };
	Slide.prototype.refreshAllContentsFields = function() {
		this.cSld.refreshAllContentsFields();
		if(this.notesShape){
			this.notesShape.handleAllContents(AscFormat.RefreshContentAllFields);
		}
	};
    Slide.prototype.Search = function(Engine, Type ){
        var sp_tree = this.cSld.spTree;
        for(var i = 0; i < sp_tree.length; ++i){
            if (sp_tree[i].Search){
                sp_tree[i].Search(Engine, Type);
            }
        }
        if(this.notesShape){
            this.notesShape.Search(Engine, Type);
        }
    };
	Slide.prototype.GetSearchElementId = function(isNext, StartPos)
    {
        var sp_tree = this.cSld.spTree, i, Id;
        if(isNext)
        {
            for(i = StartPos; i < sp_tree.length; ++i)
            {
                if(sp_tree[i].GetSearchElementId)
                {
                    Id = sp_tree[i].GetSearchElementId(isNext, false);
                    if(Id !== null)
                    {
                        return Id;
                    }
                }
            }
        }
        else
        {
            for(i = StartPos; i > -1; --i)
            {
                if(sp_tree[i].GetSearchElementId)
                {
                    Id = sp_tree[i].GetSearchElementId(isNext, false);
                    if(Id !== null)
                    {
                        return Id;
                    }
                }
            }
        }
        return null;
    };
    Slide.prototype.getMatchingShape = function(type, idx, bSingleBody, info)
    {
        var _input_reduced_type;
        if(type == null)
        {
            _input_reduced_type = AscFormat.phType_body;
        }
        else
        {
            if(type == AscFormat.phType_ctrTitle)
            {
                _input_reduced_type = AscFormat.phType_title;
            }
            else
            {
                _input_reduced_type = type;
            }
        }

        var _input_reduced_index;
        if(idx == null)
        {
            _input_reduced_index = 0;
        }
        else
        {
            _input_reduced_index = idx;
        }

        var _sp_tree = this.cSld.spTree;
        var _shape_index;
        var _index, _type;
        var _final_index, _final_type;
        var _glyph;
        var body_count = 0;
        var last_body;
        var oNvObjPr;
        for(_shape_index = 0; _shape_index < _sp_tree.length; ++_shape_index)
        {
            _glyph = _sp_tree[_shape_index];
            if(_glyph.isPlaceholder())
            {
                oNvObjPr = _glyph.getUniNvProps();
                if(oNvObjPr)
                {
                    _index = oNvObjPr.nvPr.ph.idx;
                    _type = oNvObjPr.nvPr.ph.type;
                    if(_type == null)
                    {
                        _final_type = AscFormat.phType_body;
                    }
                    else
                    {
                        if(_type == AscFormat.phType_ctrTitle)
                        {
                            _final_type = AscFormat.phType_title;
                        }
                        else
                        {
                            _final_type = _type;
                        }
                    }

                    if(_index == null)
                    {
                        _final_index = 0;
                    }
                    else
                    {
                        _final_index = _index;
                    }

                    if(_input_reduced_type == _final_type && _input_reduced_index == _final_index)
                    {
                        if(info){
                            info.bBadMatch = !(_type === type && _index === idx);
                        }
                        return _glyph;
                    }
                    if(_input_reduced_type == AscFormat.phType_title && _input_reduced_type == _final_type)
                    {
                        if(info){
                            info.bBadMatch = !(_type === type && _index === idx);
                        }
                        return _glyph;
                    }
                    if(AscFormat.phType_body === _type)
                    {
                        ++body_count;
                        last_body = _glyph;
                    }
                }
            }
        }


        if(_input_reduced_type == AscFormat.phType_sldNum || _input_reduced_type == AscFormat.phType_dt || _input_reduced_type == AscFormat.phType_ftr || _input_reduced_type == AscFormat.phType_hdr)
        {
            for(_shape_index = 0; _shape_index < _sp_tree.length; ++_shape_index)
            {
                _glyph = _sp_tree[_shape_index];
                if(_glyph.isPlaceholder())
                {
                    oNvObjPr = _glyph.getUniNvProps();
                    if(oNvObjPr)
                    {
                        _type = oNvObjPr.nvPr.ph.type;
                        if(_input_reduced_type == _type)
                        {
                            if(info){
                                info.bBadMatch = !(_type === type && _index === idx);
                            }
                            return _glyph;
                        }
                    }
                }
            }
        }

        if(info){
            return null;
        }
        if(body_count === 1 && _input_reduced_type === AscFormat.phType_body && bSingleBody)
        {
            if(info){
                info.bBadMatch = !(_type === type && _index === idx);
            }
            return last_body;
        }

        for(_shape_index = 0; _shape_index < _sp_tree.length; ++_shape_index)
        {
            _glyph = _sp_tree[_shape_index];
            if(_glyph.isPlaceholder())
            {
                if(_glyph instanceof AscFormat.CShape)
                {
                    _index = _glyph.nvSpPr.nvPr.ph.idx;
                    _type = _glyph.nvSpPr.nvPr.ph.type;
                }
                if(_glyph instanceof AscFormat.CImageShape)
                {
                    _index = _glyph.nvPicPr.nvPr.ph.idx;
                    _type = _glyph.nvPicPr.nvPr.ph.type;
                }
                if(_glyph instanceof  AscFormat.CGroupShape)
                {
                    _index = _glyph.nvGrpSpPr.nvPr.ph.idx;
                    _type = _glyph.nvGrpSpPr.nvPr.ph.type;
                }

                if(_index == null)
                {
                    _final_index = 0;
                }
                else
                {
                    _final_index = _index;
                }

                if(_input_reduced_index == _final_index)
                {
                    if(info){
                        info.bBadMatch = true;
                    }
                    return _glyph;
                }
            }
        }


        if(body_count === 1 && bSingleBody)
        {
            if(info){
                info.bBadMatch = !(_type === type && _index === idx);
            }
            return last_body;
        }

        return null;
    };
    Slide.prototype.changeNum = function(num)
    {
        this.num = num;
    };
    Slide.prototype.recalcText = function()
    {
        this.recalcInfo.recalculateSpTree = true;
		this.cSld.forEachSp(function(oSp) {
			oSp.recalcText();
		});
    };
    Slide.prototype.addComment = function(comment)
    {
        if(AscCommon.isRealObject(this.slideComments))
        {
            this.slideComments.addComment(comment);
        }
    };
    Slide.prototype.changeComment = function(id, commentData)
    {
        if(AscCommon.isRealObject(this.slideComments))
        {
            this.slideComments.changeComment(id, commentData);
        }
    };
    Slide.prototype.removeMyComments = function()
    {
        if(AscCommon.isRealObject(this.slideComments))
        {
            this.slideComments.removeMyComments();
        }
    };
    Slide.prototype.removeAllComments = function()
    {
        if(AscCommon.isRealObject(this.slideComments))
        {
            this.slideComments.removeAllComments();
        }
    };
    Slide.prototype.removeComment = function(id, bForce)
    {
        if(AscCommon.isRealObject(this.slideComments))
        {
            this.slideComments.removeComment(id, bForce);
        }
    };
    Slide.prototype.addToRecalculate = function()
    {
        History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
    };
    Slide.prototype.getAllMyComments = function(aAllComments)
    {
        if(this.slideComments)
        {
            this.slideComments.getAllMyComments(aAllComments, this);
        }
    };
    Slide.prototype.getAllComments = function(aAllComments, isMine, isCurrent, aIds)
    {
        if(this.slideComments)
        {
            this.slideComments.getAllComments(aAllComments, isMine, isCurrent, aIds);
        }
    };
    Slide.prototype.Refresh_RecalcData = function(data)
    {
        if(data)
        {
            switch(data.Type)
            {
                case AscDFH.historyitem_SlideSetBg:
                {
                    this.recalcInfo.recalculateBackground = true;
                    break;
                }
                case AscDFH.historyitem_SlideSetLayout:
                {
                    this.checkSlideTheme();
                    if(this.Layout){
                        this.lastLayoutType = this.Layout.type;
                        this.lastLayoutMatchingName = this.Layout.matchingName;
                        this.lastLayoutName = this.Layout.cSld.name;
                    }
                    break;
                }
            }
            this.addToRecalculate();
        }
    };
    Slide.prototype.setNotes = function(pr){
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetNotes, this.notes, pr));
        this.notes = pr;
    };
    Slide.prototype.setSlideComments = function(comments)
    {
       History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetComments, this.slideComments, comments));
        this.slideComments = comments;
    };
    Slide.prototype.setShow = function(bShow)
    {
       History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_SlideSetShow, this.show, bShow));
        this.show = bShow;
    };
    Slide.prototype.setShowPhAnim = function(bShow)
    {
       History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_SlideSetShowPhAnim, this.showMasterPhAnim, bShow));
        this.showMasterPhAnim = bShow;
    };
    Slide.prototype.setShowMasterSp = function(bShow)
    {
       History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_SlideSetShowMasterSp, this.showMasterSp, bShow));
        this.showMasterSp = bShow;
    };
    Slide.prototype.setLayout = function(layout)
    {
       History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetLayout, this.Layout, layout));
        this.Layout = layout;
        if(layout){
            this.lastLayoutType = layout.type;
            this.lastLayoutMatchingName = layout.matchingName;
            this.lastLayoutName = layout.cSld.name;
        }
    };

	Slide.prototype.changeLayout = function(layout) {
		this.setLayout(layout);
        this.bChangeLayout = true;
        for (var j = this.cSld.spTree.length - 1; j > -1; --j) {
            var shape = this.cSld.spTree[j];
            if (shape.isEmptyPlaceholder()) {
                this.removeFromSpTreeById(shape.Get_Id());
            } else {
                var oInfo = {};
                var hierarchy = shape.getHierarchy(undefined, oInfo);
                var bNoPlaceholder = true;
                var bNeedResetTransform = false;
                var oNotNullPH = null;
                for (var t = 0; t < hierarchy.length; ++t) {
                    if (hierarchy[t]) {
                        if (hierarchy[t].parent && (hierarchy[t].parent instanceof AscCommonSlide.SlideLayout)) {
                            bNoPlaceholder = false;
                        }
                        if (hierarchy[t].spPr && hierarchy[t].spPr.xfrm && hierarchy[t].spPr.xfrm.isNotNull()) {
                            bNeedResetTransform = true;
                            oNotNullPH = hierarchy[t];
                        }
                    }
                }
                if (bNoPlaceholder) {
                    if (this.cSld.spTree[j].isEmptyPlaceholder()) {
                        this.removeFromSpTreeById(this.cSld.spTree[j].Get_Id());
                    } else {
                        var hierarchy2 = shape.getHierarchy(undefined, undefined);
                        for (var t = 0; t < hierarchy2.length; ++t) {
                            if (hierarchy2[t]) {
                                if (hierarchy2[t].spPr && hierarchy2[t].spPr.xfrm && hierarchy2[t].spPr.xfrm.isNotNull()) {
                                    break;
                                }
                            }
                        }
                        if (t === hierarchy2.length) {
                            AscFormat.CheckSpPrXfrm(shape);
                        }
                    }
                } else {
                    if (bNeedResetTransform) {
                        if (shape.spPr && shape.spPr.xfrm && shape.spPr.xfrm.isNotNull()) {
                            if (shape.getObjectType() !== AscDFH.historyitem_type_GraphicFrame) {
                                shape.spPr.setXfrm(null);
                            } else {
                                if (oNotNullPH) {
                                    if (!shape.spPr && oNotNullPH.spPr) {
                                        shape.setSpPr(oNotNullPH.spPr.createDuplicate());
                                        shape.spPr.setParent(shape);
                                    }
                                    if (!shape.spPr.xfrm && oNotNullPH.spPr && oNotNullPH.spPr.xfrm) {
                                        shape.spPr.setXfrm(oNotNullPH.spPr.xfrm.createDuplicate());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            // else
            // {
            //     if(shape.isPlaceholder() && (!shape.spPr || !shape.spPr.xfrm || !shape.spPr.xfrm.isNotNull()))
            //     {
            //         var hierarchy = shape.getHierarchy();
            //         for(var t = 0; t < hierarchy.length; ++t)
            //         {
            //             if(hierarchy[t] && hierarchy[t].spPr && hierarchy[t].spPr.xfrm && hierarchy[t].spPr.xfrm.isNotNull())
            //             {
            //                 break;
            //             }
            //         }
            //         if(t === hierarchy.length)
            //         {
            //             AscFormat.CheckSpPrXfrm(shape);
            //         }
            //     }
            // }
        }
        for (var j = 0; j < layout.cSld.spTree.length; ++j) {
            if (layout.cSld.spTree[j].isPlaceholder()) {
                var _ph_type = layout.cSld.spTree[j].getPlaceholderType();
                var hf = layout.Master.hf;
                var bIsSpecialPh = _ph_type === AscFormat.phType_dt || _ph_type === AscFormat.phType_ftr || _ph_type === AscFormat.phType_hdr || _ph_type === AscFormat.phType_sldNum;
                if (!bIsSpecialPh || hf && ((_ph_type === AscFormat.phType_dt && (hf.dt !== false)) ||
                    (_ph_type === AscFormat.phType_ftr && (hf.ftr !== false)) ||
                    (_ph_type === AscFormat.phType_hdr && (hf.hdr !== false)) ||
                    (_ph_type === AscFormat.phType_sldNum && (hf.sldNum !== false)))) {
                    var matching_shape = this.getMatchingShape(layout.cSld.spTree[j].getPlaceholderType(), layout.cSld.spTree[j].getPlaceholderIndex(), layout.cSld.spTree[j].getIsSingleBody ? layout.cSld.spTree[j].getIsSingleBody() : false);
                    if (matching_shape == null && layout.cSld.spTree[j]) {
                        var sp = layout.cSld.spTree[j].copy(undefined);
                        sp.setParent(this);
                        !bIsSpecialPh && sp.clearContent && sp.clearContent();
                        this.addToSpTreeToPos(this.cSld.spTree.length, sp)
                    }
                }
            }
        }
	};
    Slide.prototype.setSlideNum = function(num)
    {
        History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_SlideSetNum, this.num, num));
        this.num = num;
    };
    Slide.prototype.applyTransition = function(transition)
    {
        var oldTransition = this.transition.createDuplicate();
        this.transition.applyProps(transition);
       History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideSetTransition, oldTransition, this.transition.createDuplicate()));
    };
    Slide.prototype.setTiming = function(oTiming)
    {
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetTiming, this.timing, oTiming));
        this.timing = oTiming;
        if(this.timing)
        {
            this.timing.setParent(this);
        }
    };
    Slide.prototype.setSlideSize = function(w, h)
    {
       History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideSetSize, new AscFormat.CDrawingBaseCoordsWritable(this.Width, this.Height), new AscFormat.CDrawingBaseCoordsWritable(w, h)));
        this.Width = w;
        this.Height = h;
    };
    Slide.prototype.changeBackground = function(bg)
    {
       History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideSetBg, this.cSld.Bg , bg));
        this.cSld.Bg = bg;
    };
    Slide.prototype.setLocks = function(deleteLock, backgroundLock, timingLock, transitionLock, layoutLock, showLock)
    {
        this.deleteLock = deleteLock;
        this.backgroundLock = backgroundLock;
        this.timingLock = timingLock;
        this.transitionLock = transitionLock;
        this.layoutLock = layoutLock;
        this.showLock = showLock;
       History.Add(new AscDFH.CChangesDrawingSlideLocks(this, deleteLock, backgroundLock, timingLock, transitionLock, layoutLock, showLock));
    };
    Slide.prototype.shapeAdd = function(pos, item)
    {
        this.checkDrawingUniNvPr(item);
        var _pos = (AscFormat.isRealNumber(pos) && pos > -1 && pos <= this.cSld.spTree.length) ? pos : this.cSld.spTree.length;
       History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_SlideAddToSpTree, _pos, [item], true, true));
        this.cSld.spTree.splice(_pos, 0, item);
        item.setParent2(this);
        if(this.collaborativeMarks) {
            this.collaborativeMarks.Update_OnAdd(_pos);
        }
    };

    Slide.prototype.shapeRemove = function (pos, count) {
        if(pos > -1 && pos < this.cSld.spTree.length){
        History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_SlideRemoveFromSpTree, pos, this.cSld.spTree.slice(pos, pos + count), false));
        this.cSld.spTree.splice(pos, count);
        }
    };

    Slide.prototype.isEditingInFastMultipleUsers = function() {
        var oPresentation = editor.WordControl.m_oLogicDocument;
        if(oPresentation && oPresentation.IsEditingInFastMultipleUsers()) {
            return true;
        }
        return false;
    };
    Slide.prototype.checkNeedCopyTimingBeforeEdit = function() {
        // if(this.isEditingInFastMultipleUsers()) {
        //     if(this.timing) {
        //         this.setTiming(this.timing.createDuplicate());
        //     }
        // }
    };
    Slide.prototype.addAnimation = function(nPresetClass, nPresetId, nPresetSubtype, oColor, bReplace) {
        this.checkNeedCopyTimingBeforeEdit();
        if(!this.timing) {
            this.setTiming(new AscFormat.CTiming());
        }
        return this.timing.addAnimation(nPresetClass, nPresetId, nPresetSubtype, oColor, bReplace);
    };
    Slide.prototype.setAnimationProperties = function(oPr) {
        if(!this.timing) {
            return;
        }
        this.checkNeedCopyTimingBeforeEdit();
        this.timing.setAnimationProperties(oPr);
        this.showDrawingObjects();
    };

    Slide.prototype.isVisible = function(){
        return this.show !== false;
    };

    Slide.prototype.checkDrawingUniNvPr = function(drawing)
    {
        var nv_sp_pr;
        if(drawing)
        {
            drawing.checkDrawingUniNvPr();
        }
    };


    Slide.prototype.CheckLayout = function(){
        var bRet = true;
        if(!this.Layout || !this.Layout.CheckCorrect()){
            var oMaster =  this.presentation.slideMasters[0];
            if(!oMaster){
                bRet = false;
            }
            else{
                var oLayout = oMaster.getMatchingLayout(this.lastLayoutType, this.lastLayoutMatchingName, this.lastLayoutName, undefined);
                if(oLayout){
                    this.setLayout(oLayout);
                }
                else{
                    bRet = false;
                }
            }
        }
        return bRet;
    };

    Slide.prototype.correctContent = function(){


        for(var i = this.cSld.spTree.length - 1;  i > -1 ; --i){
            if(this.cSld.spTree[i].CheckCorrect && !this.cSld.spTree[i].CheckCorrect() || this.cSld.spTree[i].bDeleted){
                if(this.cSld.spTree[i].setBDeleted){
                    this.cSld.spTree[i].setBDeleted(true);
                }
                this.removeFromSpTreeById(this.cSld.spTree[i].Get_Id());
            }
        }
        for(var i = this.cSld.spTree.length - 1;  i > -1 ; --i){
            for(var j = i - 1; j > -1; --j){
                if(this.cSld.spTree[i] === this.cSld.spTree[j]){
                    this.removeFromSpTreeByPos(i);
                    break;
                }
            }
        }
    };


    Slide.prototype.removeFromSpTreeByPos = function(pos){
        if(pos > -1 && pos < this.cSld.spTree.length){
            var oSp = this.cSld.spTree[pos];
            History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_SlideRemoveFromSpTree, pos, [oSp], false));
            this.cSld.spTree.splice(pos, 1);
            if(this.timing && !AscCommon.IsChangingDrawingZIndex) {
                this.checkNeedCopyTimingBeforeEdit();
                this.timing.onRemoveObject(oSp.Get_Id());
            }
            if(this.collaborativeMarks) {
                this.collaborativeMarks.Update_OnRemove(pos, 1);
            }
        }
    };

    Slide.prototype.removeFromSpTreeById = function(id)
    {
        var sp_tree = this.cSld.spTree;
        for(var i = 0; i < sp_tree.length; ++i)
        {
            if(sp_tree[i].Get_Id() === id)
            {
                this.removeFromSpTreeByPos(i);
                return i;
            }
        }
        return null;
    };

    Slide.prototype.addToSpTreeToPos = function(pos, obj)
    {
        this.shapeAdd(pos, obj);
    };

    Slide.prototype.replaceSp = function(oPh, oObject)
    {
        var aSpTree = this.cSld.spTree;
        for(var i = 0; i < aSpTree.length; ++i)
        {
            if(aSpTree[i] === oPh)
            {
                break;
            }
        }
        this.removeFromSpTreeByPos(i);
        this.addToSpTreeToPos(i, oObject);
        var oNvProps = oObject.getNvProps && oObject.getNvProps();
        if(oNvProps)
        {
            if(oPh)
            {
                var oNvPropsPh = oPh.getNvProps && oPh.getNvProps();
                var oPhPr = oNvPropsPh.ph;
                if(oPhPr)
                {
                    oNvProps.setPh(oPhPr.createDuplicate());
                }
            }
        }
    };

    Slide.prototype.setCSldName = function(name)
    {
       History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_SlideSetCSldName, this.cSld.name, name));
        this.cSld.name = name;
    };

    Slide.prototype.setClMapOverride = function(clrMap)
    {
       History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideSetClrMapOverride, this.clrMap, clrMap));
        this.clrMap = clrMap;
    };

    Slide.prototype.getAllFonts = function(fonts)
    {
	    this.cSld.forEachSp(function(oSp) {
		    oSp.getAllFonts(fonts);
	    });
    };

    Slide.prototype.getParentObjects = function()
    {
        var oRet = {master: null, layout: null, slide: this};
        if(this.Layout)
        {
            oRet.layout = this.Layout;
            if(this.Layout.Master)
            {
                oRet.master = this.Layout.Master;
            }
        }
        return oRet;
    };

    Slide.prototype.copySelectedObjects = function(){
        var aSelectedObjects, i, fShift = 5.0;
        var oSelector = this.graphicObjects.selection.groupSelection ? this.graphicObjects.selection.groupSelection : this.graphicObjects;
        aSelectedObjects = [].concat(oSelector.selectedObjects);
        oSelector.resetSelection(undefined, false);
        var bGroup = this.graphicObjects.selection.groupSelection ? true : false;
        if(bGroup){
            oSelector.normalize();
        }
        for(i = 0; i < aSelectedObjects.length; ++i){
            var oCopy = aSelectedObjects[i].copy(undefined);
            oCopy.x = aSelectedObjects[i].x;
            oCopy.y = aSelectedObjects[i].y;
            oCopy.extX = aSelectedObjects[i].extX;
            oCopy.extY = aSelectedObjects[i].extY;
            AscFormat.CheckSpPrXfrm(oCopy, true);
            oCopy.spPr.xfrm.setOffX(oCopy.x + fShift);
            oCopy.spPr.xfrm.setOffY(oCopy.y + fShift);
            oCopy.setParent(this);
            if(!bGroup){
                this.addToSpTreeToPos(undefined, oCopy);
            }
            else{
                oCopy.setGroup(aSelectedObjects[i].group);
                aSelectedObjects[i].group.addToSpTree(undefined, oCopy);
            }
            oSelector.selectObject(oCopy, 0);
        }
        if(bGroup){
            oSelector.updateCoordinatesAfterInternalResize();
        }
    };

    Slide.prototype.getAllRasterImages = function(images) {
		let oBgPr = this.cSld.Bg && this.cSld.Bg.bgPr;
        if(oBgPr) {
	        oBgPr.checkBlipFillRasterImage(images);
        }
		this.cSld.forEachSp(function(oSp) {
			oSp.getAllRasterImages(images);
		});
    };

    Slide.prototype.changeSize = function(width, height)
    {
        var kw = width/this.Width, kh = height/this.Height;
        this.setSlideSize(width, height);
	    this.cSld.forEachSp(function(oSp) {
		    oSp.changeSize(kw, kh);
	    });
    };

    Slide.prototype.checkSlideSize = function()
    {
        this.recalcInfo.recalculateSpTree = true;
	    this.cSld.forEachSp(function(oSp) {
		    oSp.handleUpdateExtents();
	    });
    };

    Slide.prototype.checkSlideTheme = function()
    {

        this.bChangeLayout = undefined;
        this.recalcInfo.recalculateSpTree = true;
        this.recalcInfo.recalculateBackground = true;
		this.cSld.forEachSp(function(oSp) {
			oSp.handleUpdateTheme();
		});
    };

    Slide.prototype.checkSlideColorScheme = function()
    {
        this.recalcInfo.recalculateSpTree = true;
        this.recalcInfo.recalculateBackground = true;
	    this.cSld.forEachSp(function(oSp) {
		    oSp.handleUpdateFill();
		    oSp.handleUpdateLn();
	    });
    };
    Slide.prototype.Get_ColorMap = function()
    {
        if(this.clrMap)
        {
            return this.clrMap;
        }
        else if(this.Layout && this.Layout.clrMap)
        {
            return this.Layout.clrMap;
        }
        else if(this.Layout.Master && this.Layout.Master.clrMap)
        {
            return this.Layout.Master.clrMap;
        }
        return AscFormat.GetDefaultColorMap();
    };

    Slide.prototype.recalculate = function()
    {
        if(!this.Layout || !AscFormat.isRealNumber(this.num))
        {
            return;
        }
        this.Layout.recalculate();
        if(this.recalcInfo.recalculateBackground)
        {
            this.recalculateBackground();
            this.recalcInfo.recalculateBackground = false;
        }
        if(this.recalcInfo.recalculateSpTree)
        {
            this.recalculateSpTree();
            this.recalcInfo.recalculateSpTree = false;
        }
        this.recalculateNotesShape();
        this.cachedImage = null;
    };

    Slide.prototype.recalculateBackground = function()
    {
        var _back_fill = null;
        var RGBA = {R:0, G:0, B:0, A:255};

        var _layout = this.Layout;
        var _master = _layout.Master;
        var _theme = _master.Theme;
        if (this.cSld.Bg != null)
        {
            if (null != this.cSld.Bg.bgPr)
                _back_fill = this.cSld.Bg.bgPr.Fill;
            else if(this.cSld.Bg.bgRef != null)
            {
                this.cSld.Bg.bgRef.Color.Calculate(_theme, this, _layout, _master, RGBA);
                RGBA = this.cSld.Bg.bgRef.Color.RGBA;
                _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(this.cSld.Bg.bgRef.idx, this.cSld.Bg.bgRef.Color);
            }
        }
        else
        {
            if (_layout != null)
            {
                if (_layout.cSld.Bg != null)
                {
                    if (null != _layout.cSld.Bg.bgPr)
                        _back_fill = _layout.cSld.Bg.bgPr.Fill;
                    else if(_layout.cSld.Bg.bgRef != null)
                    {
                        _layout.cSld.Bg.bgRef.Color.Calculate(_theme, this, _layout, _master, RGBA);
                        RGBA = _layout.cSld.Bg.bgRef.Color.RGBA;
                        _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(_layout.cSld.Bg.bgRef.idx, _layout.cSld.Bg.bgRef.Color);
                    }
                }
                else if (_master != null)
                {
                    if (_master.cSld.Bg != null)
                    {
                        if (null != _master.cSld.Bg.bgPr)
                            _back_fill = _master.cSld.Bg.bgPr.Fill;
                        else if(_master.cSld.Bg.bgRef != null)
                        {
                            _master.cSld.Bg.bgRef.Color.Calculate(_theme, this, _layout, _master, RGBA);
                            RGBA = _master.cSld.Bg.bgRef.Color.RGBA;
                            _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(_master.cSld.Bg.bgRef.idx, _master.cSld.Bg.bgRef.Color);
                        }
                    }
                    else
                    {
                        _back_fill = new AscFormat.CUniFill();
                        _back_fill.fill = new AscFormat.CSolidFill();
                        _back_fill.fill.color =  new AscFormat.CUniColor();
                        _back_fill.fill.color.color = new AscFormat.CRGBColor();
                        _back_fill.fill.color.color.RGBA = {R:255, G:255, B:255, A:255};
                    }
                }
            }
        }

        if (_back_fill != null)
            _back_fill.calculate(_theme, this, _layout, _master, RGBA);

        this.backgroundFill = _back_fill;
    };

    Slide.prototype.recalculateSpTree = function()
    {
        for(var i = 0; i < this.cSld.spTree.length; ++i)
            this.cSld.spTree[i].recalculate();
    };

    Slide.prototype.getNotesHeight = function(){
        if(!this.notesShape){
            return 0;
        }

        var oDocContent = this.notesShape.getDocContent();
        if(oDocContent){
            return oDocContent.GetSummaryHeight();
        }
        return 0;
    };

    Slide.prototype.recalculateNotesShape = function(){
        AscFormat.ExecuteNoHistory(function(){
            if(!this.notes){
                this.notesShape = null;//ToDo: Create notes body shape
            }
            else{
                this.notesShape = this.notes.getBodyShape();
                if(this.notesShape && this.notesShape.getObjectType() !== AscDFH.historyitem_type_Shape){
                    this.notesShape = null;
                }
            }

            if(this.notesShape){
                this.notes.slide = this;
                this.notes.graphicObjects.selectObject(this.notesShape, 0);
                this.notes.graphicObjects.selection.textSelection = this.notesShape;
                var oDocContent = this.notesShape.getDocContent();
                if(oDocContent){
                    oDocContent.CalculateAllFields();
                    this.notesShape.transformText.tx = 3;
                    this.notesShape.transformText.ty = 3;
                    this.notesShape.invertTransformText = AscCommon.global_MatrixTransformer.Invert(this.notesShape.transformText);
                    var Width = AscCommonSlide.GetNotesWidth();
                    oDocContent.Reset(0, 0, Width, 2000);
                    var CurPage = 0;
                    var RecalcResult = recalcresult2_NextPage;
                    while ( recalcresult2_End !== RecalcResult  )
                        RecalcResult = oDocContent.Recalculate_Page( CurPage++, true );
                    this.notesShape.contentWidth = Width;
                    this.notesShape.contentHeight = 2000;
                }
                this.notesShape.transformText2 = this.notesShape.transformText;
                this.notesShape.invertTransformText2 = this.notesShape.invertTransformText;
                var oOldGeometry = this.notesShape.spPr.geometry;
                this.notesShape.spPr.geometry = null;
                this.notesShape.extX = Width;
                this.notesShape.extY = 2000;
                this.notesShape.recalculateContent2();
                this.notesShape.spPr.geometry = oOldGeometry;
                this.notesShape.pen = AscFormat.CreateNoFillLine();
            }
        }, this, []);
    };

    Slide.prototype.needMasterSpDraw = function() {
			return this.showMasterSp ? this.Layout.showMasterSp : false;
    };

    Slide.prototype.needLayoutSpDraw = function() {
			return this.showMasterSp;
    };

    Slide.prototype.isEqualBgMasterAndLayout = function(oSlide) {
        if(!(this.backgroundFill === oSlide.backgroundFill ||
            this.backgroundFill && this.backgroundFill.isEqual(oSlide.backgroundFill))) {
            return false;
        }
        if(this.needMasterSpDraw() && !oSlide.needMasterSpDraw() || oSlide.needMasterSpDraw() && !this.needMasterSpDraw()) {
            return false;
        }
        if(this.needMasterSpDraw()) {
            if(this.Layout.Master !== oSlide.Layout.Master) {
               return false;
            }
        }
        if(this.needLayoutSpDraw()) {
            if(this.Layout !== oSlide.Layout) {
                return false;
            }
        }
        return true;
    };

    Slide.prototype.drawBgMasterAndLayout = function(graphics, bClipBySlide, bCheckBounds) {
        let _bounds;
        DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
        if(bClipBySlide) {
            graphics.SaveGrState();
            graphics.AddClipRect(0, 0, this.Width, this.Height);
        }
        if(this.needMasterSpDraw()) {
            if (bCheckBounds) {
                _bounds =  this.Layout.Master.bounds;
                graphics.rect(_bounds.l, _bounds.t, _bounds.w, _bounds.h);
            }
            else {
                this.Layout.Master.draw(graphics, this);
            }
        }
        if(this.needLayoutSpDraw()) {
            if (bCheckBounds) {
                _bounds =  this.Layout.bounds;
                graphics.rect(_bounds.l, _bounds.t, _bounds.w, _bounds.h);
            }
            else {
                this.Layout.draw(graphics, this);
            }
        }
    };
    Slide.prototype.draw = function(graphics) {
        let bCheckBounds = graphics.isBoundsChecker();
        let bSlideShow = this.graphicObjects.isSlideShow();
        let bClipBySlide = !this.graphicObjects.canEdit();
        if (bCheckBounds && (bSlideShow || bClipBySlide)) {
            graphics.rect(0, 0, this.Width, this.Height);
            return;
        }
        let _bounds, i;
        this.drawBgMasterAndLayout(graphics, bClipBySlide, bCheckBounds);
        this.collaborativeMarks.Init_Drawing();
        let oCollColor;
        let fDist = 3;
        let oIdentityMtx = new AscCommon.CMatrix();
        for(i = 0; i < this.cSld.spTree.length; ++i) {
            let oSp = this.cSld.spTree[i];
            if(AscCommon.IsHiddenObj(oSp)) {
                continue;
            }
            if(this.collaborativeMarks) {
                oCollColor = this.collaborativeMarks.Check(i);
                if(oCollColor) {
                    var oBounds = oSp.bounds;
                    graphics.transform3(oIdentityMtx);
                    if(graphics.put_GlobalAlpha) {
                        graphics.put_GlobalAlpha(true, 0.5);
                    }
                    let dX = oBounds.l - fDist;
                    let dY = oBounds.t - fDist;
                    let dW = oBounds.r - oBounds.l + 2*fDist;
                    let dH = oBounds.b - oBounds.t + 2*fDist;
                    graphics.drawCollaborativeChanges(dX, dY, dW, dH, oCollColor);
                    if(graphics.put_GlobalAlpha) {
                        graphics.put_GlobalAlpha(false, 1);
                    }
                }
            }
            oSp.draw(graphics);
        }

        if(this.timing) {
            let aShapes = this.timing.getMoveEffectsShapes();
            if(aShapes) {
                for(i = 0; i < aShapes.length; ++i) {
                    aShapes[i].draw(graphics);
                }
            }
            this.timing.drawEffectsLabels(graphics);
        }
        if(this.slideComments) {
            let aComments = this.slideComments.comments;
            for(i = 0; i < aComments.length; ++i) {
                var oComment = aComments[i];
                if(AscCommon.UserInfoParser.canViewComment(oComment.GetUserName()) !== false) {
                    oComment.draw(graphics);
                }
            }
        }
        if(!bCheckBounds && !bSlideShow) {
            this.drawViewPrMarks(graphics);
        }
        if(bClipBySlide) {
            graphics.RestoreGrState();
        }
    };

    function CStrideData(oPresentation) {
        this.presentation = oPresentation;
        this.slideWidth = null;
        this.slideHeight = null;
        this.stride = null;
        this.originX = null;
        this.originY = null;

    }
    CStrideData.prototype.checkUpdate = function() {
        if(this.slideWidth !== this.presentation.GetWidthEMU() ||
        this.slideHeight !== this.presentation.GetHeightEMU() ||
        this.stride !== this.presentation.getViewPropertiesStride()) {
            this.update();
            return true;
        }
        return false;
    };
    CStrideData.prototype.update = function() {
        this.stride = this.presentation.getViewPropertiesStride();
        this.slideWidth = this.presentation.GetWidthEMU();
        this.slideHeight = this.presentation.GetHeightEMU();

        this.originX = this.getStartStridePos(this.stride, this.slideWidth);
        this.originY = this.getStartStridePos(this.stride, this.slideHeight);
    };
    CStrideData.prototype.getStartStridePos = function(stride, len) {
		let nCenterPos = len / 2 + 0.5 >> 0;
        let nPos = nCenterPos - ((nCenterPos / stride >> 0) + 1)*stride;
        return nPos;
    };
    CStrideData.prototype.getNearestLinearPoint = function(nDX, nOrigin) {
		let nDX_ = Math.abs(nDX);
        let nCX = nDX_ / this.stride >> 0;
        let dRX = nDX_ - nCX * this.stride;
        let nX;
        if((dRX / this.stride) < 0.5) {
            nX = nCX * this.stride;
        }
        else {
            nX = (nCX + 1) * this.stride;
        }
		if(nDX < 0) {
			nX = -nX;
		}
        return nOrigin + nX;
    };
    CStrideData.prototype.getNearestPoint = function(x, y) {
        this.checkUpdate();
        let em = function (dValue) {return dValue / 36000;};
        let me = function(dValue) { return dValue * 36000 + 0.5 >> 0;};
        let nDX = me(x) - this.originX;
        let nDY = me(y) - this.originY;
        let nX = this.getNearestLinearPoint(nDX, this.originX);
        let nY = this.getNearestLinearPoint(nDY, this.originY);
        return {x: em(nX), y: em(nY)};
    };
    CStrideData.prototype.drawGrid = function(oGraphics) {
        let oContext = oGraphics.m_oContext;
        if(!oContext) {
            return;
        }
        this.checkUpdate();
        let dPixScale = oGraphics.m_oCoordTransform.sx;
        let mmp = function(dMM) {
            return (dMM*dPixScale + 0.5) >> 0;
        };
        let em = function (val) {return val / 36000;};
        let ep = function(nEmu) {
            return mmp(em(nEmu));
        };

        let nGridSpacingPix = ep(this.stride);

        let nMinLineStridePix = ep(720000);
        let nMinInsideLineStridePix = ep(100000);

        let nGridSpacing = this.stride;
        let bPixel = true;
        let nStrideInsideLine = nGridSpacing;
        let nStrideLine;

        let nStrideInsideLinePix = nGridSpacingPix;
        let nStrideLinePix;

        let nHorStart;
        let nVertStart;
        let nSlideWidth = this.slideWidth;
        let nSlideHeight = this.slideHeight;

       while(nStrideInsideLinePix < nMinInsideLineStridePix) {
           nStrideInsideLine *= 2;
           nStrideInsideLinePix = ep(nStrideInsideLine);
       }
        nStrideLine = nStrideInsideLine;
        nStrideLinePix = nStrideInsideLinePix;
        while(nStrideLinePix < nMinLineStridePix) {
            nStrideLine += nStrideInsideLine;
            nStrideLinePix = ep(nStrideLine);
        }
        bPixel = nStrideInsideLinePix < AscCommon.AscBrowser.convertToRetinaValue(17, true);

        oGraphics.SaveGrState();
        oGraphics.transform3(new AscCommon.CMatrix());
        oGraphics.SetIntegerGrid(true);
        oGraphics.b_color1(0, 0, 0, 255);
	    oGraphics._s();

        let nX, nY;
        let oT = oGraphics.m_oFullTransform;
        let oImageCanvas = document.createElement('canvas');
        let oImageContext;
        if(bPixel) {
            oImageCanvas.width = 1;
            oImageCanvas.height = 1;
            oImageContext = oImageCanvas.getContext("2d");
            oImageContext.fillStyle = 'white';
            oImageContext.fillRect(0, 0, oImageCanvas.width, oImageCanvas.height);
            oImageContext.fill();
        }
        else {
            oImageCanvas.width = 3;
            oImageCanvas.height = 3;
            oImageContext = oImageCanvas.getContext("2d");
            oImageContext.fillStyle = 'white';
            oImageContext.fillRect(1, 0, 1, 1);
            oImageContext.fillRect(0, 1, 1, 1);
            oImageContext.fillRect(2, 1, 1, 1);
            oImageContext.fillRect(1, 2, 1, 1);
            oImageContext.fill();
        }

        function dp() {
            let nHP = em(nHorPos);
            let nVP = em(nVertPos);
            nX = oT.TransformPointX(nHP, nVP) + 0.5 >> 0;
            nY = oT.TransformPointY(nHP, nVP) + 0.5 >> 0;
            if(bPixel) {
                oContext.drawImage(oImageCanvas, nX, nY);
            }
            else {
                oContext.drawImage(oImageCanvas, nX - 1, nY - 1);
            }
        }

        nHorStart = this.getStartStridePos(nStrideInsideLine, nSlideWidth);
        nVertStart = this.getStartStridePos(nStrideLine, nSlideHeight);
        let nVertPos = nVertStart;
        let nHorPos;
		let nBottom = nSlideHeight + nStrideInsideLine;
		let nRight = nSlideWidth + nStrideInsideLine;
        while (nVertPos < nBottom) {
	        nHorPos = nHorStart;
	        while (nHorPos < nRight) {
		        dp();
		        nHorPos += nStrideInsideLine;
	        }
            nVertPos += nStrideLine;
        }
		if(nStrideLine !== nStrideInsideLine) {
			nHorStart = this.getStartStridePos(nStrideLine, nSlideWidth);
			nVertStart = this.getStartStridePos(nStrideInsideLine, nSlideHeight);
			nHorPos = nHorStart;

			while (nHorPos < nRight) {
				nVertPos = nVertStart;
				while (nVertPos < nBottom) {
					dp();
					nVertPos += nStrideInsideLine;
				}
				nHorPos += nStrideLine;
			}
		}

        oGraphics.df();
        oGraphics.RestoreGrState();
    };
    Slide.prototype.drawViewPrMarks = function(oGraphics) {
	    let oContext = oGraphics.m_oContext;
	    if( !oContext ||
			AscCommon.IsShapeToImageConverter ||
		    oGraphics.animationDrawer ||
		    oGraphics.IsThumbnail ||
		    oGraphics.IsDemonstrationMode ||
		    oGraphics.isBoundsChecker() ||
			oGraphics.IsNoDrawingEmptyPlaceholder) {
		    return;
	    }

	    let oApi = editor;
        if(!oApi) {
            return;
        }
        if(!oApi.WordControl) {
            return;
        }
        let oPresentation = oApi.WordControl.m_oLogicDocument;
        if(!oPresentation) {
            return;
        }
        if(oApi.asc_getShowGridlines()) {
            oPresentation.checkGridCache(oGraphics);
        }
        if(oApi.asc_getShowGuides()) {
            oPresentation.drawGuides(oGraphics);
        }
    };
    Slide.prototype.drawNotes = function (g) {
        if(this.notesShape) {
            this.notesShape.draw(g);
            var oLock = this.notesShape.Lock;
            if(oLock && AscCommon.c_oAscLockTypes.kLockTypeNone != oLock.Get_Type()) {
                var bCoMarksDraw = true;
                if(typeof editor !== "undefined" && editor && AscFormat.isRealBool(editor.isCoMarksDraw)) {
                    bCoMarksDraw = editor.isCoMarksDraw;
                }
                if(bCoMarksDraw) {
                    g.transform3(this.notesShape.transformText);
                    var Width = this.notesShape.txBody.content.XLimit - 2;
                    Width = Math.max(Width, 1);
                    var Height = this.notesShape.txBody.content.GetSummaryHeight();
                    g.DrawLockObjectRect(oLock.Get_Type(), 0, 0, Width, Height);
                }
            }
        }
    };
    Slide.prototype.drawAnimPane = function(oGraphics) {
        if(this.timing) {
            this.timing.drawAnimPane(oGraphics);
        }
    };
	Slide.prototype.isAnimated = function() {
		let oTr = this.transition;
		if(oTr
			&& oTr.TransitionType !== undefined
			&& oTr.TransitionType !== null
			&& oTr.TransitionType !== c_oAscSlideTransitionTypes.None) {
			return true;
		}
		if(this.timing) {
			if(this.timing.hasEffects()) {
				return true;
			}
		}
		return false;
	};
    Slide.prototype.onAnimPaneResize = function(oGraphics) {
        if(this.timing) {
            this.timing.onAnimPaneResize(oGraphics);
        }
    };

    Slide.prototype.onAnimPaneMouseDown = function(e, x, y) {
        if(this.timing) {
            this.timing.onAnimPaneMouseDown(e, x, y);
        }
    };
    Slide.prototype.onAnimPaneMouseMove = function(e, x, y) {
        if(this.timing) {
            this.timing.onAnimPaneMouseMove(e, x, y);
        }
    };
    Slide.prototype.onAnimPaneMouseUp = function(e, x, y) {
        if(this.timing) {
            this.timing.onAnimPaneMouseUp(e, x, y);
        }
    };
    Slide.prototype.onAnimPaneMouseWheel = function(e, deltaY, X, Y) {
        if(this.timing) {
            this.timing.onAnimPaneMouseWheel(e, deltaY, X, Y);
        }
    };
    Slide.prototype.getTheme = function(){
        return this.Layout && this.Layout.Master && this.Layout.Master.Theme || null;
    };

    Slide.prototype.drawSelect = function(_type)
    {
        if (_type === undefined)
        {
            this.graphicObjects.drawTextSelection(this.num);
            this.graphicObjects.drawSelect(0, this.presentation.DrawingDocument);
        }
        else if (_type == 1)
            this.graphicObjects.drawTextSelection(this.num);
        else if (_type == 2)
            this.graphicObjects.drawSelect(0, this.presentation.DrawingDocument);
    };

    Slide.prototype.drawNotesSelect = function(){

        if(this.notesShape){
            var content = this.notesShape.getDocContent();
            if(content)
            {
                this.presentation.DrawingDocument.UpdateTargetTransform(this.notesShape.transformText);
                content.DrawSelectionOnPage(0);
            }
        }
    };

    Slide.prototype.removeAllCommentsToInterface = function()
    {
        if(this.slideComments)
        {
            var aComments = this.slideComments.comments;
            for(var i = aComments.length - 1; i > -1; --i )
            {
                editor.sync_RemoveComment(aComments[i].Get_Id());
            }
        }
    };

    Slide.prototype.getDrawingObjects = function()
    {
        return this.cSld.spTree;
    };

    Slide.prototype.paragraphAdd = function(paraItem, bRecalculate)
    {
        this.graphicObjects.paragraphAdd(paraItem, bRecalculate);
    };

    Slide.prototype.OnUpdateOverlay = function()
    {
        this.presentation.DrawingDocument.m_oWordControl.OnUpdateOverlay();
    };

    Slide.prototype.sendGraphicObjectProps = function()
    {
        editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
    };

    Slide.prototype.isViewerMode = function()
    {
        return editor.isViewMode;
    };

    Slide.prototype.onMouseDown = function(e, x, y)
    {
        this.graphicObjects.onMouseDown(e, x, y);
    };

    Slide.prototype.onMouseMove = function(e, x, y)
    {
        this.graphicObjects.onMouseMove(e, x, y);
    };

    Slide.prototype.onMouseUp = function(e, x, y)
    {
        this.graphicObjects.onMouseUp(e, x, y);
    };


    Slide.prototype.getColorMap = function()
    {
        if(this.clrMap)
        {
            return this.clrMap;
        }
        else if(this.Layout)
        {
            if(this.Layout.clrMap)
            {
                return this.Layout.clrMap;
            }
            else if(this.Layout.Master)
            {
                if(this.Layout.Master.clrMap)
                {
                    return this.Layout.Master.clrMap;
                }
            }
        }
        else if(this.Master)
        {
            if(this.Master.clrMap)
            {
                return this.Master.clrMap;
            }
        }
        return AscFormat.GetDefaultColorMap();
    };


    Slide.prototype.showDrawingObjects = function()
    {
        editor.WordControl.m_oDrawingDocument.OnRecalculatePage(this.num, this);
    };

    Slide.prototype.showComment = function(Id, x, y)
    {
        //editor.sync_HideComment();
        editor.sync_ShowComment(Id, x, y );
    };


    Slide.prototype.getSlideIndex = function()
    {
        return this.num;
    };


    Slide.prototype.getWorksheet = function()
    {
        return null;
    };

    Slide.prototype.showChartSettings = function()
    {
        editor.asc_onOpenChartFrame();
        editor.sendEvent("asc_doubleClickOnChart", this.graphicObjects.getChartObject());
        this.graphicObjects.changeCurrentState(new AscFormat.NullState(this.graphicObjects));
    };


    Slide.prototype.Clear_ContentChanges  = function()
    {
        this.m_oContentChanges.Clear();
    };

    Slide.prototype.Add_ContentChanges  = function(Changes)
    {
        this.m_oContentChanges.Add( Changes );
    };

    Slide.prototype.Refresh_ContentChanges  = function()
    {
        this.m_oContentChanges.Refresh();
    };


    Slide.prototype.isLockedObject = function()
    {
      //  var sp_tree = this.cSld.spTree;
      //  for(var i = 0; i < sp_tree.length; ++i)
      //  {
      //      if(sp_tree[i].Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeMine && sp_tree[i].Lock.Type !== AscCommon.c_oAscLockTypes.kLockTypeNone)
      //          return true;
      //  }
        return false;
    };

    Slide.prototype.getPlaceholdersControls = function()
    {
        var ret = [];
        var aSpTree = this.cSld.spTree;
        for(var i = 0; i < aSpTree.length; ++i)
        {
            var oSp = aSpTree[i];
            oSp.createPlaceholderControl(ret);
        }
        return ret;
    };

    Slide.prototype.convertPixToMM = function(pix)
    {
        return editor.WordControl.m_oDrawingDocument.GetMMPerDot(pix);
    };

    Slide.prototype.getBase64Img = function()
    {
        if(window["NATIVE_EDITOR_ENJINE"])
        {
           return "";
        }
        if(typeof this.cachedImage === "string" && this.cachedImage.length > 0)
            return this.cachedImage;

        AscCommon.IsShapeToImageConverter = true;

        var dKoef = AscCommon.g_dKoef_mm_to_pix;
        var _need_pix_width     = ((this.Width*dKoef/3.0 + 0.5) >> 0);
        var _need_pix_height    = ((this.Height*dKoef/3.0 + 0.5) >> 0);

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

        var _canvas = document.createElement('canvas');
        _canvas.width = _need_pix_width;
        _canvas.height = _need_pix_height;

        var _ctx = _canvas.getContext('2d');

        var g = new AscCommon.CGraphics();
        g.init(_ctx, _need_pix_width, _need_pix_height, this.Width, this.Height);
        g.m_oFontManager = AscCommon.g_fontManager;

        g.m_oCoordTransform.tx = 0.0;
        g.m_oCoordTransform.ty = 0.0;
        g.transform(1,0,0,1,0,0);

        this.draw(g, /*pageIndex*/0);

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
            _ret.ImageUrl = _canvas.toDataURL("image/png");
        }
        catch (err)
        {
            _ret.ImageUrl = "";
        }
        return _ret.ImageUrl;
    };

    Slide.prototype.checkNoTransformPlaceholder = function()
    {
        var sp_tree = this.cSld.spTree;
        for(var i = 0; i < sp_tree.length; ++i)
        {
            var sp = sp_tree[i];
            if(sp.getObjectType() === AscDFH.historyitem_type_Shape || sp.getObjectType() === AscDFH.historyitem_type_ImageShape || sp.getObjectType() === AscDFH.historyitem_type_OleObject)
            {
                if(sp.isPlaceholder && sp.isPlaceholder())
                {
                    var hierarchy = sp.getHierarchy();
                    for(var j = 0; j < hierarchy.length; ++j)
                    {
                        if(AscCommon.isRealObject(hierarchy[j]))
                            break;
                    }
                    if(j === hierarchy.length)
                    {
                        AscFormat.CheckSpPrXfrm(sp, true);
                    }
                }
                else if(sp.getObjectType() === AscDFH.historyitem_type_Shape){
                    sp.handleUpdateTheme();
                    sp.checkExtentsByDocContent();
                }
            }
            else{
                if(sp.isGroupObject()){
                    sp.handleUpdateTheme();
                    sp.checkExtentsByDocContent();
                }
            }
        }
    };

    Slide.prototype.getSnapArrays = function()
    {
        var snapX = [];
        var snapY = [];
        for(var i = 0; i < this.cSld.spTree.length; ++i)
        {
            if(this.cSld.spTree[i].getSnapArrays)
            {
                this.cSld.spTree[i].getSnapArrays(snapX, snapY);
            }
        }
        return {snapX: snapX, snapY: snapY};
    };


    Slide.prototype.Load_Comments  = function(authors)
    {
        AscCommonSlide.fLoadComments(this, authors);
    };


	Slide.prototype.RestartSpellCheck = function()
    {
        for(var i = 0; i < this.cSld.spTree.length; ++i)
        {
            this.cSld.spTree[i].RestartSpellCheck();
        }
        if(this.notes)
        {
            var spTree = this.notes.cSld.spTree;
            for(i = 0; i < spTree.length; ++i)
            {
                spTree[i].RestartSpellCheck();
            }
        }
    };

    Slide.prototype.getDrawingsForController = function(){
        if(this.timing) {
            let aShapes = this.timing.getMoveEffectsShapes();
            if(aShapes && aShapes.length > 0) {
                return this.cSld.spTree.concat(aShapes);
            }
        }
        return this.cSld.spTree;
    };
    
    Slide.prototype.createFontMap = function (oFontsMap, oCheckedMap) {
        var aSpTree = this.cSld.spTree;
        var nSp, nSpCount = aSpTree.length;
        for(nSp = 0; nSp < nSpCount; ++nSp) {
            aSpTree[nSp].createFontMap(oFontsMap);
        }
        if(this.needMasterSpDraw()) {
            this.Layout.Master.createFontMap(oFontsMap, oCheckedMap, true);
        }
        if(this.needLayoutSpDraw()) {
            this.Layout.createFontMap(oFontsMap, oCheckedMap, true);
        }
        oCheckedMap[this.Get_Id()] = this;
        if(this.notesShape) {
            this.notesShape.createFontMap(oFontsMap);
        }
    };
    Slide.prototype.getAnimationPlayer = function() {
        if(!this.animationPlayer) {
            var oDemoManager = editor.WordControl.DemonstrationManager;
            this.animationPlayer = new AscFormat.CAnimationPlayer(this, oDemoManager);
        }
        return this.animationPlayer;
    };
    Slide.prototype.isAdvanceAfterTransition = function() {
        var oTransition = this.transition;
        if(!oTransition) {
            return false;
        }
        if(this.presentation) {
            var oShowPr = this.presentation.showPr;
            if(oShowPr && oShowPr.useTimings === false) {
                return false;
            }
        }
        return oTransition.SlideAdvanceAfter === true;
    };
    Slide.prototype.getAdvanceDuration = function() {
        var oTransition = this.transition;
        if(!oTransition) {
            return 0;
        }
        return oTransition.SlideAdvanceDuration;
    };
function fLoadComments(oObject, authors)
{
    var _comments_count = oObject.writecomments.length;
    var _comments_id = [];
    var _comments_data = [];
    var _comments_data_author_id = [];
    var _comments = [];

    var oComments = oObject.slideComments ? oObject.slideComments : oObject.comments;
    if(!oComments)
    {
        return;
    }
    for (var i = 0; i < _comments_count; i++)
    {
        var _wc = oObject.writecomments[i];

        if (0 == _wc.WriteParentAuthorId || 0 == _wc.WriteParentCommentId)
        {
            var commentData = new AscCommon.CCommentData();

            commentData.m_sText = _wc.WriteText;
            commentData.m_sUserId = ("" + _wc.WriteAuthorId);
            commentData.m_sUserName = "";
            commentData.m_sTime = _wc.WriteTime;
            commentData.m_nTimeZoneBias = _wc.timeZoneBias;
            if (commentData.m_sTime && null != commentData.m_nTimeZoneBias) {
                commentData.m_sOOTime = (parseInt(commentData.m_sTime) + commentData.m_nTimeZoneBias * 60000) + "";
            }

            for (var k in authors)
            {
                if (_wc.WriteAuthorId == authors[k].Id)
                {
                    commentData.m_sUserName = authors[k].Name;
                    break;
                }
            }

            //if ("" != commentData.m_sUserName)
            {
                _comments_id.push(_wc.WriteCommentId);
                _comments_data.push(commentData);
                _comments_data_author_id.push(_wc.WriteAuthorId);

                _wc.ParceAdditionalData(commentData);

                var comment = new AscCommon.CComment(oComments, new AscCommon.CCommentData());
                comment.setPosition(_wc.x / 22.66, _wc.y / 22.66);
                _comments.push(comment);
            }
        }
        else
        {
            var commentData = new AscCommon.CCommentData();

            commentData.m_sText = _wc.WriteText;
            commentData.m_sUserId = ("" + _wc.WriteAuthorId);
            commentData.m_sUserName = "";
            commentData.m_sTime = _wc.WriteTime;
            commentData.m_nTimeZoneBias = _wc.timeZoneBias;
            if (commentData.m_sTime && null != commentData.m_nTimeZoneBias) {
                commentData.m_sOOTime = (parseInt(commentData.m_sTime) + commentData.m_nTimeZoneBias * 60000) + "";
            }

            for (var k in authors)
            {
                if (_wc.WriteAuthorId == authors[k].Id)
                {
                    commentData.m_sUserName = authors[k].Name;
                    break;
                }
            }

            _wc.ParceAdditionalData(commentData);

            var _parent = null;
            for (var j = 0; j < _comments_data.length; j++)
            {
                if ((_wc.WriteParentAuthorId == _comments_data_author_id[j]) && (_wc.WriteParentCommentId == _comments_id[j]))
                {
                    _parent = _comments_data[j];
                    break;
                }
            }

            if (null != _parent)
            {
                _parent.m_aReplies.push(commentData);
            }
        }
    }

    for (var i = 0; i < _comments.length; i++)
    {
        _comments[i].Set_Data(_comments_data[i]);
        oObject.addComment(_comments[i]);
    }

    oObject.writecomments = [];
}

function PropLocker(objectId)
{
    this.objectId = null;
    this.Lock = new AscCommon.CLock();
    this.Id = AscCommon.g_oIdCounter.Get_NewId();
    g_oTableId.Add(this, this.Id);

    if(typeof  objectId === "string")
    {
        this.setObjectId(objectId);
    }

}

PropLocker.prototype = {

    getObjectType: function()
    {
        return AscDFH.historyitem_type_PropLocker;
    },
    setObjectId: function(id)
    {
       History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PropLockerSetId, this.objectId, id));
        this.objectId = id;
    },
    Get_Id: function()
    {
        return this.Id;
    },
    Write_ToBinary2: function(w)
    {
        w.WriteLong(AscDFH.historyitem_type_PropLocker);
        w.WriteString2(this.Id);
    },

    Read_FromBinary2: function(r)
    {
        this.Id = r.GetString2();
    },

    Refresh_RecalcData: function()
    {}

};

AscFormat.CTextBody.prototype.Get_StartPage_Absolute = function()
{
    if(this.parent)
    {
        if(this.parent.getParentObjects)
        {
            var parent_objects = this.parent.getParentObjects();
            if(parent_objects.slide)
            {
                return parent_objects.slide.num;
            }
            if(parent_objects.notes && parent_objects.notes.slide){
                return parent_objects.notes.slide.num;
            }
        }
    }
    return 0;
};
AscFormat.CTextBody.prototype.Get_AbsolutePage = function(CurPage)
{
    return this.Get_StartPage_Absolute();
};
AscFormat.CTextBody.prototype.Get_AbsoluteColumn = function(CurPage)
{
    return 0;//TODO;
};
AscFormat.CTextBody.prototype.checkCurrentPlaceholder = function()
{
    var presentation = editor.WordControl.m_oLogicDocument;
    var oCurController = presentation.GetCurrentController();
    if(oCurController)
    {
        return oCurController.getTargetDocContent() === this.content;
    }
    return false;
};



function SlideComments(slide)
{
    this.comments = [];
    this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)
    this.slide = slide;
    this.Id = AscCommon.g_oIdCounter.Get_NewId();
    g_oTableId.Add(this, this.Id);
}

SlideComments.prototype =
{
    getObjectType: function()
    {
        return AscDFH.historyitem_type_SlideComments;
    },

    Get_Id: function()
    {
        return this.Id;
    },

    Clear_ContentChanges : function()
    {
        this.m_oContentChanges.Clear();
    },

    Add_ContentChanges : function(Changes)
    {
        this.m_oContentChanges.Add( Changes );
    },

    Refresh_ContentChanges : function()
    {
        this.m_oContentChanges.Refresh();
    },

    addComment: function(comment)
    {
       History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsAddComment, this.comments.length, [comment], true));
        this.comments.splice(this.comments.length, 0, comment);
        comment.slideComments = this;
    },


    getSlideIndex: function()
    {
        if(this.slide)
        {
            return this.slide.num;
        }
        return null;
    },

    changeComment: function(id, commentData)
    {
        for(var i = 0; i < this.comments.length; ++i)
        {
            if(this.comments[i].Get_Id() === id)
            {
                this.comments[i].Set_Data(commentData);
                return;
            }
        }
    },

    removeComment: function(id, bForce)
    {
        for(var i = 0; i < this.comments.length; ++i)
        {
            var oComment = this.comments[i];
            if(oComment.Get_Id() === id && (bForce || oComment.canBeDeleted()))
            {
                History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsRemoveComment, i, this.comments.splice(i, 1), false));
                editor.sync_RemoveComment(id);
                return;
            }
        }
    },

    removeMyComments: function()
    {
        if(!editor.DocInfo)
        {
            return;
        }
        var sUserId = editor.DocInfo.get_UserId();
        for(var i = this.comments.length - 1; i > -1; --i)
        {
            var oComment = this.comments[i];
            var oCommentData = oComment.Data;
            if(oCommentData.m_sUserId === sUserId)
            {
                if(oComment.canBeDeleted())
                {
                    History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsRemoveComment, i, this.comments.splice(i, 1), false));
                    editor.sync_RemoveComment(oComment.Get_Id());
                }
            }
            else
            {
                oComment.removeUserReplies(sUserId);
            }
        }
    },

    removeAllComments: function()
    {
        for(var i = this.comments.length - 1; i > -1; --i)
        {
            var oComment = this.comments[i];
            if(oComment.canBeDeleted())
            {
                History.Add(new AscDFH.CChangesDrawingsContentComments(this, AscDFH.historyitem_SlideCommentsRemoveComment, i, this.comments.splice(i, 1), false));
                editor.sync_RemoveComment(oComment.Get_Id());
            }
        }
    },

    removeSelectedComment: function()
    {
        var comment = this.getSelectedComment();
        if(comment)
        {
            this.removeComment(comment.Get_Id());
        }
    },

    getSelectedComment: function()
    {
        for(var i = 0; i < this.comments.length; ++i)
        {
            if(this.comments[i].selected)
            {
                return this.comments[i];
            }
        }
        return null;
    },

    getAllMyComments: function(aAllComments, oSlide)
    {
        if(!editor.DocInfo)
        {
            return;
        }
        var sUserId = editor.DocInfo.get_UserId();
        for(var i = 0; i < this.comments.length; ++i)
        {
            var oComment = this.comments[i];
            var oCommentData = oComment.Data;
            if(oCommentData.m_sUserId === sUserId)
            {
                aAllComments.push({comment: oComment, slide: oSlide});
            }
            else
            {
                for(var j = 0; j < oCommentData.m_aReplies.length; ++j)
                {
                    if(oCommentData.m_aReplies[j].m_sUserId === sUserId)
                    {
                        aAllComments.push({comment: oComment, slide: oSlide});
                        break;
                    }
                }
            }
        }
    },
    getAllComments: function(aAllComments, isMine, isCurrent, aIds)
    {
        var oSlide = null;
        if(this.slide && (this.slide instanceof Slide))
        {
            oSlide = this.slide;
        }
        var nComment, oComment;
        if(Array.isArray(aIds))
        {
            var oIdMap = {};
            for(var nId = 0; nId < aIds.length; ++nId)
            {
                oIdMap[aIds[nId]] = true;
            }
            for(nComment = 0; nComment < this.comments.length; ++nComment)
            {
                oComment = this.comments[nComment];
                if(oIdMap[oComment.Id])
                {
                    aAllComments.push({comment: oComment, slide: oSlide});
                }
            }
        }
        else
        {
            if(isCurrent)
            {
                if(oSlide)
                {
                    var oPresentation = oSlide.presentation;
                    if(oPresentation && oPresentation.Slides[oPresentation.CurPage] === oSlide)
                    {
                        for(nComment = 0; nComment < this.comments.length; ++nComment)
                        {
                            oComment = this.comments[nComment];
                            if(oComment.selected && (!isMine || oComment.isMineComment()))
                            {
                                aAllComments.push({comment: oComment, slide: oSlide});
                                return;
                            }
                        }
                    }
                }
            }
            else
            {
                for(nComment = 0; nComment < this.comments.length; ++nComment)
                {
                    oComment = this.comments[nComment];
                    if((!isMine || oComment.isMineComment()))
                    {
                        aAllComments.push({comment: oComment, slide: oSlide});
                    }
                }
            }
        }
    },

    recalculate: function()
    {},

    Write_ToBinary2: function(w)
    {
        w.WriteLong(AscDFH.historyitem_type_SlideComments);
        w.WriteString2(this.Id);
        AscFormat.writeObject(w, this.slide);
    },

    Read_FromBinary2: function(r)
    {
        this.Id = r.GetString2();
        this.slide = AscFormat.readObject(r);
    },

    Refresh_RecalcData: function()
    {
        History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
    }
};

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonSlide'] = window['AscCommonSlide'] || {};
window['AscCommonSlide'].Slide = Slide;
window['AscCommonSlide'].PropLocker = PropLocker;
window['AscCommonSlide'].SlideComments = SlideComments;
window['AscCommonSlide'].fLoadComments = fLoadComments;
window['AscCommonSlide'].CStrideData = CStrideData;
