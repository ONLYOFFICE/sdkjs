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

// Import
var History = AscCommon.History;


 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetThemeIndex]     = AscDFH.CChangesDrawingsLong            ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetSize]           = AscDFH.CChangesDrawingsObjectNoId      ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetTheme]          = AscDFH.CChangesDrawingsObject          ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterAddToSpTree]       = AscDFH.CChangesDrawingsContent         ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetBg]             = AscDFH.CChangesDrawingsObjectNoId      ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetTxStyles]       = AscDFH.CChangesDrawingsObjectNoId      ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetCSldName]       = AscDFH.CChangesDrawingsString          ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetClrMapOverride] = AscDFH.CChangesDrawingsObject          ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetHF]             = AscDFH.CChangesDrawingsObject          ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterAddLayout]         = AscDFH.CChangesDrawingsContent         ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetTransition]     = AscDFH.CChangesDrawingsObjectNoId         ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterSetTiming]         = AscDFH.CChangesDrawingsObject         ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterRemoveLayout]      = AscDFH.CChangesDrawingsContent         ;
 AscDFH.changesFactory[AscDFH.historyitem_SlideMasterRemoveFromSpTree]  = AscDFH.CChangesDrawingsContent         ;

 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetThemeIndex]     = function(oClass, value){oClass.ThemeIndex = value;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetSize]           = function(oClass, value){oClass.Width = value.a; oClass.Height = value.b;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetTheme]          = function(oClass, value){oClass.Theme = value;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetBg]             = function(oClass, value, FromLoad){
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
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetTxStyles]       = function(oClass, value){oClass.txStyles = value;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetCSldName]       = function(oClass, value){oClass.cSld.name = value;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetClrMapOverride] = function(oClass, value){oClass.clrMap = value;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetHF]             = function(oClass, value){oClass.hf = value;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetTransition]     = function(oClass, value){oClass.transition = value;};
 AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideMasterSetTiming]         = function(oClass, value){oClass.timing = value;};


AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideMasterSetSize]      = AscFormat.CDrawingBaseCoordsWritable;
AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideMasterSetBg]        = AscFormat.CBg;
AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideMasterSetTxStyles]  = AscFormat.CTextStyles;
AscDFH.drawingsConstructorsMap[AscDFH.historyitem_SlideMasterSetTransition]  = Asc.CAscSlideTransition;


AscDFH.drawingContentChanges[AscDFH.historyitem_SlideMasterAddToSpTree]       = function(oClass){return oClass.cSld.spTree;};
AscDFH.drawingContentChanges[AscDFH.historyitem_SlideMasterAddLayout]         = function(oClass){return oClass.sldLayoutLst;};
AscDFH.drawingContentChanges[AscDFH.historyitem_SlideMasterRemoveLayout]      = function(oClass){return oClass.sldLayoutLst;};
AscDFH.drawingContentChanges[AscDFH.historyitem_SlideMasterRemoveFromSpTree]  = function(oClass){return oClass.cSld.spTree;};


function MasterSlide(presentation, theme)
{
    AscFormat.CBaseFormatObject.call(this);
    this.cSld = new AscFormat.CSld(this);
    this.clrMap = new AscFormat.ClrMap();

    this.hf = null;

    this.sldLayoutLst = [];

    this.txStyles = null;
    this.preserve = false;

    this.timing = null;
    this.transition = null;

    this.ImageBase64 = "";
    this.Width64 = 0;
    this.Height64 = 0;


    // pointers
    this.Theme = null;
    this.TableStyles = null;
    this.Vml = null;

    this.Width = 254;
    this.Height = 190.5;
    this.recalcInfo = {};
    this.DrawingDocument = editor.WordControl.m_oDrawingDocument;
    this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)

    this.bounds = new AscFormat.CGraphicBounds(0, 0, this.Width, this.Height);

//----------------------------------------------
    this.presentation = editor.WordControl.m_oLogicDocument;
    this.theme = theme;

    this.kind = AscFormat.TYPE_KIND.MASTER;
    this.recalcInfo =
    {
        recalculateBackground: true,
        recalculateSpTree: true,
        recalculateBounds: true,
        recalculateSlideLayouts: true
    };


    this.lastRecalcSlideIndex = 0;

    this.presentation = editor && editor.WordControl && editor.WordControl.m_oLogicDocument;
    this.graphicObjects = new AscFormat.DrawingObjectsController(this);

    this.deleteLock = new PropLocker(this.Id);
    this.backgroundLock = new PropLocker(this.Id);
    this.timingLock = new PropLocker(this.Id);
    this.transitionLock = new PropLocker(this.Id);
    this.layoutLock = new PropLocker(this.Id);
    this.showLock = new PropLocker(this.Id);
}
AscFormat.InitClass(MasterSlide, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_SlideMaster)

MasterSlide.prototype.addLayout = function (layout) {
    this.addToSldLayoutLstToPos(this.sldLayoutLst.length, layout);
};
MasterSlide.prototype.isAnimated = function () {
    return false;
};

MasterSlide.prototype.getDrawingDocument = function() {
    return editor.WordControl.m_oLogicDocument.DrawingDocument;
};
MasterSlide.prototype.getTheme = function(){
    return  this.Theme || null;
};
MasterSlide.prototype.getColorMap = function() {
    return AscFormat.GetDefaultColorMap();
};
MasterSlide.prototype.getMaster = function(){
    return this.getParentObjects().master;
};
MasterSlide.prototype.getDrawingsForController = function(){
    return this.cSld.spTree;
};

MasterSlide.prototype.sendGraphicObjectProps = function()
{
    editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
};
MasterSlide.prototype.showDrawingObjects = function()
{
    editor.WordControl.m_oDrawingDocument.OnRecalculateSlide(this.getNum());
};
MasterSlide.prototype.getNum = function () {
    let aSlides = this.presentation.GetAllSlides();
    for(let nIdx = 0; nIdx < aSlides.length; ++nIdx) {
        if(aSlides[nIdx] === this)
            return nIdx;
    }
    return -1;
};
MasterSlide.prototype.drawSelect = function(_type)
{
    if (_type === undefined)
    {
        this.graphicObjects.drawTextSelection(this.getNum());
        this.graphicObjects.drawSelect(0, this.presentation.DrawingDocument);
    }
    else if (_type == 1)
        this.graphicObjects.drawTextSelection(this.getNum());
    else if (_type == 2)
        this.graphicObjects.drawSelect(0, this.presentation.DrawingDocument);
};
MasterSlide.prototype.OnUpdateOverlay = function()
{
    this.presentation.DrawingDocument.m_oWordControl.OnUpdateOverlay();
};
MasterSlide.prototype.isLockedObject = function () {
    return false;
};
MasterSlide.prototype.Write_ToBinary2 = function (w) {
    w.WriteLong(AscDFH.historyitem_type_SlideMaster);
    w.WriteString2(this.Id);
    AscFormat.writeObject(w, this.theme);
};
MasterSlide.prototype.Read_FromBinary2 = function (r) {
    this.Id = r.GetString2();
    this.theme = AscFormat.readObject(r);
};


MasterSlide.prototype.recalculateBackground = function() {
    var _back_fill = null;
    var RGBA = {R: 0, G: 0, B: 0, A: 255};

    var _layout = null;
    var _master = this;
    var _theme = this.Theme;
    if (this.cSld.Bg != null) {
        if (null != this.cSld.Bg.bgPr)
            _back_fill = this.cSld.Bg.bgPr.Fill;
        else if (this.cSld.Bg.bgRef != null) {
            this.cSld.Bg.bgRef.Color.Calculate(_theme, this, _layout, _master, RGBA);
            RGBA = this.cSld.Bg.bgRef.Color.RGBA;
            _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(this.cSld.Bg.bgRef.idx, this.cSld.Bg.bgRef.Color);
        }
    }
    if (_back_fill != null)
        _back_fill.calculate(_theme, this, _layout, _master, RGBA);

    this.backgroundFill = _back_fill;
}

MasterSlide.prototype.drawNoPlaceholders = function(graphics, slide) {
    if(slide) {
        if(AscFormat.isRealNumber(slide.num) && slide.num !== this.lastRecalcSlideIndex) {
            this.lastRecalcSlideIndex = slide.num;
            this.cSld.refreshAllContentsFields(true);
        }
    }
    else {
        if(-1 !== this.lastRecalcSlideIndex){
            this.lastRecalcSlideIndex = -1;
            this.cSld.refreshAllContentsFields(true);

        }
    }
    this.recalculate();

    DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
    this.cSld.forEachSp(function(oSp) {
        if ( !AscCommon.IsHiddenObj(oSp) && !oSp.isPlaceholder()) {
            oSp.draw(graphics);
        }
    });
};
MasterSlide.prototype.drawNoPlaceholdersShapesOnly = function(graphics, slide) {
    if(slide) {
        if(AscFormat.isRealNumber(slide.num) && slide.num !== this.lastRecalcSlideIndex) {
            this.lastRecalcSlideIndex = slide.num;
            this.cSld.refreshAllContentsFields(true);
        }
    }
    else {
        if(-1 !== this.lastRecalcSlideIndex){
            this.lastRecalcSlideIndex = -1;
            this.cSld.refreshAllContentsFields(true);

        }
    }
    this.recalculate();
    this.cSld.forEachSp(function(oSp) {
        if ( !AscCommon.IsHiddenObj(oSp) && !oSp.isPlaceholder()) {
            oSp.draw(graphics);
        }
    });
};

MasterSlide.prototype.draw = function (graphics, slide) {
	if(slide) {
		if(AscFormat.isRealNumber(slide.num) && slide.num !== this.lastRecalcSlideIndex) {
			this.lastRecalcSlideIndex = slide.num;
			this.cSld.refreshAllContentsFields(true);
		}
	}
    else {
        if(-1 !== this.lastRecalcSlideIndex){
            this.lastRecalcSlideIndex = -1;
            this.cSld.refreshAllContentsFields(true);

        }
    }
    this.recalculate();

    DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
	this.cSld.forEachSp(function(oSp) {
		if ( !AscCommon.IsHiddenObj(oSp)) {
			oSp.draw(graphics);
		}
	});

    if(!slide) {
        this.drawViewPrMarks(graphics);
    }
};
MasterSlide.prototype.getMatchingLayout = function (type, matchingName, cSldName, themeFlag) {
    var layoutType = type;

    var _layoutName = null, _layout_index, _layout;

    if (type === AscFormat.nSldLtTTitle && !(themeFlag === true)) {
        layoutType = AscFormat.nSldLtTObj;
    }
    if (layoutType != null) {
        for (var i = 0; i < this.sldLayoutLst.length; ++i) {
            if (this.sldLayoutLst[i].type == layoutType) {
                return this.sldLayoutLst[i];
            }
        }
    }

    if (type === AscFormat.nSldLtTTitle && !(themeFlag === true)) {
        layoutType = AscFormat.nSldLtTTx;
        for (i = 0; i < this.sldLayoutLst.length; ++i) {
            if (this.sldLayoutLst[i].type == layoutType) {
                return this.sldLayoutLst[i];
            }
        }
    }


    if (matchingName != "" && matchingName != null) {
        _layoutName = matchingName;
    }
    else {
        if (cSldName != "" && cSldName != null) {
            _layoutName = cSldName;
        }
    }
    if (_layoutName != null) {
        var _layout_name;
        for (_layout_index = 0; _layout_index < this.sldLayoutLst.length; ++_layout_index) {
            _layout = this.sldLayoutLst[_layout_index];
            _layout_name = null;

            if (_layout.matchingName != null && _layout.matchingName != "") {
                _layout_name = _layout.matchingName;
            }
            else {
                if (_layout.cSld.name != null && _layout.cSld.name != "") {
                    _layout_name = _layout.cSld.name;
                }
            }
            if (_layout_name == _layoutName) {
                return _layout;
            }
        }
    }
    for (_layout_index = 0; _layout_index < this.sldLayoutLst.length; ++_layout_index) {
        _layout = this.sldLayoutLst[_layout_index];
        _layout_name = null;

        if (_layout.type != AscFormat.nSldLtTTitle) {
            return _layout;
        }

    }

    return this.sldLayoutLst[0];
};
MasterSlide.prototype.handleAllContents = Slide.prototype.handleAllContents;
MasterSlide.prototype.getAllRasterImagesForDraw = Slide.prototype.getAllRasterImagesForDraw;
MasterSlide.prototype.checkImageDraw = Slide.prototype.checkImageDraw;
MasterSlide.prototype.getMatchingShape = Slide.prototype.getMatchingShape;
MasterSlide.prototype.recalculate = function () {
    if (!this.Theme) return;
    var _shapes = this.cSld.spTree;
    var _shape_index, _slideLayout_index;
    var _shape_count = _shapes.length;
    var bRecalculateBounds = this.recalcInfo.recalculateBounds;
    var bRecalculateSlideLayouts = this.recalcInfo.recalculateSlideLayouts;
    var bRecalculateBackground = this.recalcInfo.recalculateBackground;
    var bRecalculateSpTree = this.recalcInfo.recalculateSpTree;
    if (bRecalculateBounds) {
        this.bounds.reset(this.Width + 100.0, this.Height + 100.0, -100.0, -100.0);
    }
    var bChecked = false;
    for (_shape_index = 0; _shape_index < _shape_count; ++_shape_index) {
        if (!_shapes[_shape_index].isPlaceholder() || true) {
            _shapes[_shape_index].recalculate();
            if (bRecalculateBounds) {
                this.bounds.checkByOther(_shapes[_shape_index].bounds);
            }
            bChecked = true;
        }
    }
    if (bRecalculateBounds) {
        if (bChecked) {
            this.bounds.checkWH();
        }
        else {
            this.bounds.reset(0.0, 0.0, 0.0, 0.0);
        }
        this.recalcInfo.recalculateBounds = false;
    }
    if(bRecalculateBackground) {
        this.recalculateBackground()
    }
    if (bRecalculateSlideLayouts || bRecalculateBackground || bRecalculateSpTree) {
        for (_slideLayout_index = 0; _slideLayout_index < this.sldLayoutLst.length; _slideLayout_index++) {
            let oLt = this.sldLayoutLst[_slideLayout_index];
            if (!oLt.cSld.Bg) {
                oLt.recalcInfo.recalculateBackground = true;
                oLt.recalcInfo.recalculateSpTree = true;
                oLt.ImageBase64 = "";
                oLt.recalculate();
            }
        }
        this.recalcInfo.recalculateSlideLayouts = false;
        this.recalcInfo.recalculateSpTree = false;
        this.recalcInfo.recalculateBackground = false;
    }
};
MasterSlide.prototype.checkSlideSize = Slide.prototype.checkSlideSize
MasterSlide.prototype.checkDrawingUniNvPr = Slide.prototype.checkDrawingUniNvPr
MasterSlide.prototype.checkSlideColorScheme = function () {
    this.recalcInfo.recalculateSpTree = true;
    this.recalcInfo.recalculateBackground = true;
    for (var i = 0; i < this.cSld.spTree.length; ++i) {
        if (!this.cSld.spTree[i].isPlaceholder()) {
            this.cSld.spTree[i].handleUpdateFill();
            this.cSld.spTree[i].handleUpdateLn();
        }
    }
};
MasterSlide.prototype.needRecalc = function(){
    var recalcInfo = this.recalcInfo;
    return recalcInfo.recalculateBackground ||
        recalcInfo.recalculateSpTree ||
        recalcInfo.recalculateBounds ||
        recalcInfo.recalculateSlideLayouts;
};
MasterSlide.prototype.setSlideSize = function (w, h) {
    History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideMasterSetSize, new AscFormat.CDrawingBaseCoordsWritable(this.Width, this.Height), new AscFormat.CDrawingBaseCoordsWritable(w, h)));
    this.Width = w;
    this.Height = h;
};
MasterSlide.prototype.applyTransition = function(transition) {
    var oldTransition;
    if(this.transition) {
        oldTransition = this.transition.createDuplicate();
    }
    else {
        oldTransition = null;
    }

    var oNewTransition;
    if(transition) {
        if(this.transition) {
            oNewTransition = this.transition.createDuplicate();
        }
        else {
            oNewTransition = new Asc.CAscSlideTransition();
            oNewTransition.setDefaultParams();
        }
        oNewTransition.applyProps(transition);
    }
    else {
        oNewTransition = null;
    }
    this.transition = oNewTransition;
    History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideMasterSetTransition, oldTransition, oNewTransition));
};
MasterSlide.prototype.setTiming = function(oTiming)
{
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideMasterSetTiming, this.timing, oTiming));
    this.timing = oTiming;
    if(this.timing)
    {
        this.timing.setParent(this);
    }
};
MasterSlide.prototype.changeSize = Slide.prototype.changeSize;
MasterSlide.prototype.getAllRasterImages = Slide.prototype.getAllRasterImages;
MasterSlide.prototype.Reassign_ImageUrls = Slide.prototype.Reassign_ImageUrls;
MasterSlide.prototype.setTheme = function (theme) {
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideMasterSetTheme, this.Theme, theme));
    this.Theme = theme;
};
MasterSlide.prototype.shapeAdd = function (pos, item) {
    let pos_ = pos;
    if(!AscFormat.isRealNumber(pos)) {
        pos_ = this.cSld.spTree.length;
    }
    this.checkDrawingUniNvPr(item);
    History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_SlideMasterAddToSpTree, pos_, [item], true));
    this.cSld.spTree.splice(pos_, 0, item);
    item.setParent2(this);
    this.recalcInfo.recalculateSpTree = true;
};
MasterSlide.prototype.addToSpTreeToPos = function(pos, obj)
{
    this.shapeAdd(pos, obj);
};
MasterSlide.prototype.shapeRemove = function (pos, count) {
    History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_SlideMasterRemoveFromSpTree, pos, this.cSld.spTree.slice(pos, pos + count), false));
    this.cSld.spTree.splice(pos, count);
};
MasterSlide.prototype.changeBackground = function (bg) {
    History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideMasterSetBg, this.cSld.Bg, bg));
    this.cSld.Bg = bg;
    this.recalcInfo.recalculateBackground = true;
};
MasterSlide.prototype.setHF = function(pr) {
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideMasterSetHF, this.hf, pr));
    this.hf = pr;
};
MasterSlide.prototype.setTxStyles = function (txStyles) {
    History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_SlideMasterSetTxStyles, this.txStyles, txStyles));
    this.txStyles = txStyles;
};
MasterSlide.prototype.setCSldName = function (name) {
    History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_SlideMasterSetCSldName, this.cSld.name, name));
    this.cSld.name = name;
};
MasterSlide.prototype.setClMapOverride = function (clrMap) {
    History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SlideMasterSetClrMapOverride, this.clrMap, clrMap));
    this.clrMap = clrMap;
};
MasterSlide.prototype.addToSldLayoutLstToPos = function (pos, obj) {
    History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_SlideMasterAddLayout, pos, [obj], true));
    this.sldLayoutLst.splice(pos, 0, obj);
    obj.setMaster(this);
    this.recalcInfo.recalculateSlideLayouts = true;
};
MasterSlide.prototype.removeFromSldLayoutLstByPos = function (pos, count) {
    History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_SlideMasterRemoveLayout, pos, this.sldLayoutLst.slice(pos, pos + count), false));
    this.sldLayoutLst.splice(pos, count);
};
MasterSlide.prototype.removeLayout = function (oLayout) {
    for(let nIdx = 0; nIdx < this.sldLayoutLst.length; ++nIdx) {
        if(oLayout === this.sldLayoutLst[nIdx]) {
            this.removeFromSldLayoutLstByPos(nIdx, 1);
            return;
        }
    }
};
MasterSlide.prototype.moveLayouts  = function (layoutsIndexes, pos) {
    var insert_pos = pos;
    var removed_layouts = [];
    for (var i = layoutsIndexes.length - 1; i > -1; --i) {
        removed_layouts.push(this.sldLayoutLst[layoutsIndexes[i]]);
        this.removeFromSldLayoutLstByPos(layoutsIndexes[i], 1);
    }
    removed_layouts.reverse();
    for (i = 0; i < removed_layouts.length; ++i) {
        this.addToSldLayoutLstToPos(insert_pos + i, removed_layouts[i]);
    }
    this.recalculate();
};
MasterSlide.prototype.getAllImages = function (images) {
    if (this.cSld.Bg && this.cSld.Bg.bgPr && this.cSld.Bg.bgPr.Fill && this.cSld.Bg.bgPr.Fill.fill instanceof AscFormat.CBlipFill && typeof this.cSld.Bg.bgPr.Fill.fill.RasterImageId === "string") {
        images[AscCommon.getFullImageSrc2(this.cSld.Bg.bgPr.Fill.fill.RasterImageId)] = true;
    }
    for (var i = 0; i < this.cSld.spTree.length; ++i) {
        if (typeof this.cSld.spTree[i].getAllImages === "function") {
            this.cSld.spTree[i].getAllImages(images);
        }
    }
};
MasterSlide.prototype.addToRecalculate = function()
{
    History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
};
MasterSlide.prototype.addNewLayout = function()
{
    let oLayout = AscCommonSlide.CreateDefaultLayout(this);
    let oPresentation = Asc.editor.private_GetLogicDocument();
    oLayout.changeSize(oPresentation.GetWidthMM(), oPresentation.GetHeightMM());
    let aSelected = oPresentation.GetSelectedSlideObjects();
    let nPos = this.sldLayoutLst.length;
    for(let nIdx = aSelected.length - 1; nIdx > -1; --nIdx) {
        if(aSelected[nIdx] === this) {
            nPos = 0;
            break;
        }
        if(aSelected[nIdx].Master === this) {
            let oLayout = aSelected[nIdx];
            for(let nLt = 0; nLt < this.sldLayoutLst.length; ++nLt) {
                if(oLayout === this.sldLayoutLst[nLt]) {
                    nPos = nLt + 1;
                    break;
                }
            }
            break;
        }
    }
    this.addToSldLayoutLstToPos(nPos, oLayout);
    return oLayout;
};

MasterSlide.prototype.getName = function () {
    if(this.Theme) {
        return this.Theme.name || "";
    }
    return "";
};

MasterSlide.prototype.setName = function (sName) {
    if(this.Theme) {
        return this.Theme.setName(sName);
    }
    return "";
};

MasterSlide.prototype.Refresh_RecalcData = function (data) {
    if(data)
    {
        switch(data.Type)
        {
            case AscDFH.historyitem_SlideMasterSetBg:
            {
                this.recalcInfo.recalculateBackground = true;
                for (var _slideLayout_index = 0; _slideLayout_index < this.sldLayoutLst.length; _slideLayout_index++) {
                    this.sldLayoutLst[_slideLayout_index].addToRecalculate();
                }
                break;
            }
            case AscDFH.historyitem_SlideMasterAddToSpTree:
                this.recalcInfo.recalculateSpTree = true;
                for (var _slideLayout_index = 0; _slideLayout_index < this.sldLayoutLst.length; _slideLayout_index++) {
                    this.sldLayoutLst[_slideLayout_index].addToRecalculate();
                }
                break;
        }
        this.addToRecalculate();
    }
};
MasterSlide.prototype.getAllFonts = function (fonts) {
    var i;
    if (this.Theme) {
        this.Theme.Document_Get_AllFontNames(fonts);
    }

    if (this.txStyles) {
        this.txStyles.Document_Get_AllFontNames(fonts);
    }

    for (i = 0; i < this.sldLayoutLst.length; ++i) {
        this.sldLayoutLst[i].getAllFonts(fonts);
    }

    for (i = 0; i < this.cSld.spTree.length; ++i) {
        if (typeof  this.cSld.spTree[i].getAllFonts === "function")
            this.cSld.spTree[i].getAllFonts(fonts);
    }
};
MasterSlide.prototype.createFontMap = function (oFontsMap, oCheckedMap, isNoPh) {
    if(oCheckedMap[this.Get_Id()]) {
        return;
    }
    var aSpTree = this.cSld.spTree;
    var nSp, oSp, nSpCount = aSpTree.length;
    for(nSp = 0; nSp < nSpCount; ++nSp) {
        oSp = aSpTree[nSp];
        if(isNoPh)
        {
            if(oSp.isPlaceholder())
            {
                continue;
            }
        }
        oSp.createFontMap(oFontsMap);
    }
    oCheckedMap[this.Get_Id()] = this;
};
MasterSlide.prototype.createDuplicate = function (IdMap) {
    var copy = new MasterSlide(null, null);
    var oIdMap = IdMap || {};
    var oPr = new AscFormat.CCopyObjectProperties();
    oPr.idMap = oIdMap;
    var i;

    if (this.clrMap) {
        copy.setClMapOverride(this.clrMap.createDuplicate());
    }
    if (typeof this.cSld.name === "string" && this.cSld.name.length > 0) {
        copy.setCSldName(this.cSld.name);
    }
    if (this.cSld.Bg) {
        copy.changeBackground(this.cSld.Bg.createFullCopy());
    }
    if(this.hf) {
        copy.setHF(this.hf.createDuplicate());
    }
    for (i = 0; i < this.cSld.spTree.length; ++i) {
        var _copy = this.cSld.spTree[i].copy(oPr);
        oIdMap[this.cSld.spTree[i].Id] = _copy.Id;
        copy.shapeAdd(copy.cSld.spTree.length, _copy);
        copy.cSld.spTree[copy.cSld.spTree.length - 1].setParent2(copy);
    }
    if (this.txStyles) {
        copy.setTxStyles(this.txStyles.createDuplicate());
    }
    if(this.timing) {
        copy.setTiming(this.timing.createDuplicate(oIdMap));
    }
    copy.setSlideSize(this.Width, this.Height);
    for(let nIdx = 0; nIdx < this.sldLayoutLst.length; ++nIdx) {
        copy.addToSldLayoutLstToPos(nIdx, this.sldLayoutLst[nIdx].createDuplicate(IdMap));
    }
    if(this.Theme) {
        copy.setTheme(this.Theme.createDuplicate());
    }
    return copy;
};
MasterSlide.prototype.Clear_ContentChanges = function()
{
};
MasterSlide.prototype.Add_ContentChanges = function(Changes)
{
};
MasterSlide.prototype.Refresh_ContentChanges = function()
{
};
MasterSlide.prototype.scale = function (kw, kh) {
    for(var i = 0; i < this.cSld.spTree.length; ++i)
    {
        this.cSld.spTree[i].changeSize(kw, kh);
    }
};
MasterSlide.prototype.copySelectedObjects = function () {
    AscCommonSlide.Slide.prototype.copySelectedObjects.call(this);
};

MasterSlide.prototype.getPlaceholdersControls = function () {
    return AscCommonSlide.Slide.prototype.getPlaceholdersControls.call(this);
};
MasterSlide.prototype.getDrawingObjects = function() {
    return AscCommonSlide.Slide.prototype.getDrawingObjects.call(this);
};
MasterSlide.prototype.getParentObjects = function () {
    return {
        presentation: Asc.editor.private_GetLogicDocument(),
        master: this,
        layout: null,
        slide: null
    };
};
MasterSlide.prototype.recalculateNotesShape = function () {
};
MasterSlide.prototype.getNotesHeight = function () {
    return 0;
};

MasterSlide.prototype.recalcText = function() {
    return AscCommonSlide.Slide.prototype.recalcText.call(this);
};
MasterSlide.prototype.checkSlideTheme = function() {
    return AscCommonSlide.Slide.prototype.checkSlideTheme.call(this);
};

MasterSlide.prototype.checkSlideColorScheme = function() {
    return AscCommonSlide.Slide.prototype.checkSlideColorScheme.call(this);
};
MasterSlide.prototype.isVisible = function(){
    return true;
};
MasterSlide.prototype.getWorksheet = function(){
    return null;
};
MasterSlide.prototype.removeFromSpTreeById = function(sId) {
    return AscCommonSlide.Slide.prototype.removeFromSpTreeById.call(this, sId);
};
MasterSlide.prototype.removeFromSpTreeByPos = function(pos) {
    return AscCommonSlide.Slide.prototype.removeFromSpTreeByPos.call(this, pos);
};
MasterSlide.prototype.convertPixToMM = function(pix) {
    return editor.WordControl.m_oDrawingDocument.GetMMPerDot(pix);
};
MasterSlide.prototype.getThemeIndex = function() {
   let aMasters = Asc.editor.WordControl.m_oLogicDocument.slideMasters;
   for(let nIdx = 0; nIdx < aMasters.length; ++nIdx) {
       if(aMasters[nIdx] === this) {
           return -nIdx - 1;
       }
   }
   return 0;
};
MasterSlide.prototype.isSlide = function () {
    return false;
};
MasterSlide.prototype.isLayout = function () {
    return false;
};
MasterSlide.prototype.isMaster = function () {
    return true;
};
MasterSlide.prototype.RestartSpellCheck = function() {
    AscCommonSlide.Slide.prototype.RestartSpellCheck.call(this);
};
MasterSlide.prototype.Search = function(Engine, Type) {
    AscCommonSlide.Slide.prototype.Search.call(this, Engine, Type);
};
MasterSlide.prototype.GetSearchElementId = function(isNext, StartPos) {
    return AscCommonSlide.Slide.prototype.GetSearchElementId.call(this, isNext, StartPos);
};
MasterSlide.prototype.Get_ColorMap = function() {
    if(this.clrMap)
    {
        return this.clrMap;
    }
    return AscFormat.GetDefaultColorMap();
};
MasterSlide.prototype.replaceSp = function(oPh, oObject) {
    return Slide.prototype.replaceSp.call(this, oPh, oObject);
};
MasterSlide.prototype.showChartSettings = function() {
    return AscCommonSlide.Slide.prototype.showChartSettings.call(this);
};
MasterSlide.prototype.IsUseInDocument = function() {
    let oPresentation = Asc.editor.private_GetLogicDocument();
    if(!oPresentation) return false;
    for(let nMaster = 0; nMaster < oPresentation.slideMasters.length; ++nMaster) {
        if(oPresentation.slideMasters[nMaster] === this) {
            return true;
        }
    }
    return false;
};
MasterSlide.prototype.drawViewPrMarks = function(oGraphics) {
    if(oGraphics.isSupportTextDraw && !oGraphics.isSupportTextDraw()) return;
    return AscCommonSlide.Slide.prototype.drawViewPrMarks.call(this, oGraphics);
};
function CMasterThumbnailDrawer()
{
    this.CanvasImage    = null;
    this.WidthMM        = 0;
    this.HeightMM       = 0;

    this.WidthPx        = 0;
    this.HeightPx       = 0;

    this.DrawingDocument = null;

    this.GetPlaceholderByTypesFromObject = function(oContainer, aPhTypes)
    {
        var nPhType;
        var oPlaceholder;
        if(oContainer)
        {
            for(nPhType = 0; nPhType < aPhTypes.length; ++nPhType)
            {
                oPlaceholder = oContainer.getMatchingShape(aPhTypes[nPhType], null, undefined, undefined);
                if(oPlaceholder && oPlaceholder.getObjectType() === AscDFH.historyitem_type_Shape)
                {
                    return oPlaceholder;
                }
            }
        }
        return null;
    };

    this.GetPlaceholderByTypes = function(_master, _layout, aPhTypes)
    {
        var oPlaceholder;
        oPlaceholder = this.GetPlaceholderByTypesFromObject(_layout, aPhTypes);
        if(oPlaceholder)
        {
            return oPlaceholder;
        }
        return this.GetPlaceholderByTypesFromObject(_master, aPhTypes);
    };
    this.GetPlaceholderTextProperties = function(_master, _layout, aPhTypes)
    {
        var oPlaceholder = this.GetPlaceholderByTypes(_master, _layout, aPhTypes);
        if(!oPlaceholder)
        {
            return null;
        }
        var oStylesObj = oPlaceholder.Get_Styles(0);
        if(oStylesObj && oStylesObj.styles)
        {
            var oPr = oStylesObj.styles.Get_Pr(oStylesObj.lastId, styletype_Paragraph, null);
            if(oPr)
            {
                return oPr.TextPr;
            }
        }
        return null;
    };
    this.GetTitleTextColor = function (_master, _layout)
    {
        var aPhTypes = [AscFormat.phType_ctrTitle, AscFormat.phType_title];
        var oTextPr = this.GetPlaceholderTextProperties(_master, _layout, aPhTypes);
        return this.GetTextColor(oTextPr, _master);
    };
    this.GetBodyTextColor = function (_master, _layout)
    {
        var aPhTypes = [AscFormat.phType_body, AscFormat.phType_subTitle, AscFormat.phType_obj];
        var oTextPr = this.GetPlaceholderTextProperties(_master, _layout, aPhTypes);
        return this.GetTextColor(oTextPr, _master);
    };
    this.GetTextColor = function(oTextPr, _master)
    {
        var oColor;
        var oFormatColor;
        var _theme = _master.Theme;
        var RGBA = {R:0, G:0, B:0, A:255};
        if(!oTextPr || !oTextPr.Unifill || !oTextPr.Unifill.fill)
        {
            oTextPr = {}
            oTextPr.Unifill = new AscFormat.CUniFill();
            oTextPr.Unifill.fill = new AscFormat.CSolidFill();
            oTextPr.Unifill.fill.color = AscFormat.builder_CreateSchemeColor('tx1');
        }
        oTextPr.Unifill.calculate(_theme, null, null, _master, RGBA, null);
        oFormatColor = oTextPr.Unifill.getRGBAColor();
        oColor = new CDocumentColor(oFormatColor.R, oFormatColor.G, oFormatColor.B);
        return oColor;
    };
    this.Draw2 = function(g, _master, use_background, use_master_shapes, params) {
        var w_px = this.WidthPx;
        var h_px = this.HeightPx;

        var _params = [
            6,  // color_w
            3,  // color_h,
            4,  // color_x
            31, // color_y
            1,  // color_delta,
            8,  // text_x
            11, // text_y (from bottom)
            18  // font_size
        ];

        if (params && params.length)
        {
            // first 2 - width & height
            for (var i = 2, len = params.length; i < len; i++) {
                _params[i - 2] = params[i];
            }
        }

        var dKoefPixToMM = this.HeightMM / h_px;
        var _back_fill = null;
        var RGBA = {R:0, G:0, B:0, A:255};
        var _layout = null;
        for (var i = 0; i < _master.sldLayoutLst.length; i++) {
            if (_master.sldLayoutLst[i].type == AscFormat.nSldLtTTitle) {
                _layout = _master.sldLayoutLst[i];
                break;
            }
        }
        var _theme = _master.Theme;
        if (_layout != null && _layout.cSld.Bg != null) {
            if (null != _layout.cSld.Bg.bgPr) {
                _back_fill = _layout.cSld.Bg.bgPr.Fill;
            } else {
                if (_layout.cSld.Bg.bgRef != null) {
                    _layout.cSld.Bg.bgRef.Color.Calculate(_theme, null, _layout, _master, RGBA);
                    RGBA = _layout.cSld.Bg.bgRef.Color.RGBA;
                    _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(_layout.cSld.Bg.bgRef.idx, _layout.cSld.Bg.bgRef.Color);
                }
            }
        } else {
            if (_master != null) {
                if (_master.cSld.Bg != null) {
                    if (null != _master.cSld.Bg.bgPr) {
                        _back_fill = _master.cSld.Bg.bgPr.Fill;
                    } else {
                        if (_master.cSld.Bg.bgRef != null) {
                            _master.cSld.Bg.bgRef.Color.Calculate(_theme, null, _layout, _master, RGBA);
                            RGBA = _master.cSld.Bg.bgRef.Color.RGBA;
                            _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(_master.cSld.Bg.bgRef.idx, _master.cSld.Bg.bgRef.Color);
                        }
                    }
                } else {
                    _back_fill = new AscFormat.CUniFill;
                    _back_fill.fill = new AscFormat.CSolidFill;
                    _back_fill.fill.color = new AscFormat.CUniColor;
                    _back_fill.fill.color.color = new AscFormat.CRGBColor;
                    _back_fill.fill.color.color.RGBA = {R:255, G:255, B:255, A:255};
                }
            }
        }

        _master.changeSize(this.WidthMM, this.HeightMM);
        _master.recalculate();
        if (_layout)
        {
            _layout.changeSize(this.WidthMM, this.HeightMM);
            _layout.recalculate();
        }

        if (_back_fill != null) {
            _back_fill.calculate(_theme, null, _layout, _master, RGBA);
        }
        if (use_background !== false) {
            DrawBackground(g, _back_fill, this.WidthMM, this.HeightMM);
        }

        if (use_master_shapes !== false)
        {
            if (null == _layout)
            {
                if(_master.needRecalc && _master.needRecalc())
                {
                    _master.recalculate();
                }
                _master.drawNoPlaceholders(g);
            }
            else
            {
                if (_layout.showMasterSp)
                {
                    if(_master.needRecalc && _master.needRecalc())
                    {
                        _master.recalculate();
                    }
                    _master.drawNoPlaceholders(g);
                }
                _layout.recalculate();
                _layout.drawNoPlaceholders(g);
            }
        }
        g.reset();
        var _color_w = _params[0] * dKoefPixToMM;
        var _color_h = _params[1] * dKoefPixToMM;
        var _color_x = _params[2] * dKoefPixToMM;
        var _color_y = _params[3] * dKoefPixToMM;
        var _color_delta = _params[4] * dKoefPixToMM;

        g.p_color(255, 255, 255, 255);
        g.b_color1(255, 255, 255, 255);
        g._s();
        g.rect(_color_x - _color_delta, _color_y - _color_delta, _color_w * 6 + 7 * _color_delta, _color_h + 2 * _color_delta);
        g.df();
        g._s();
        var _color = new AscFormat.CSchemeColor;
        for (var i = 0; i < 6; i++) {
            g._s();
            _color.id = i;
            _color.Calculate(_theme, null, null, _master, RGBA);
            g.b_color1(_color.RGBA.R, _color.RGBA.G, _color.RGBA.B, 255);
            g.rect(_color_x, _color_y, _color_w, _color_h);
            g.df();
            _color_x += _color_w + _color_delta;
        }
        g._s();

        var _api = this.DrawingDocument.m_oWordControl.m_oApi;
        AscFormat.ExecuteNoHistory(function(){
            var _oldTurn = _api.isViewMode;
            _api.isViewMode = true;

            var nFontSize = _params[7];
            var _textPr1 = new CTextPr;
            _textPr1.FontFamily = {Name:_theme.themeElements.fontScheme.majorFont.latin, Index:-1};
            _textPr1.RFonts.Ascii = {Name: _theme.themeElements.fontScheme.majorFont.latin, Index: -1};
            _textPr1.FontSize = nFontSize;
            _textPr1.Color = this.GetTitleTextColor(_master, _layout);

            var _textPr2 = new CTextPr;
            _textPr2.FontFamily = {Name:_theme.themeElements.fontScheme.minorFont.latin, Index:-1};
            _textPr2.RFonts.Ascii = {Name: _theme.themeElements.fontScheme.minorFont.latin, Index: -1};
            _textPr2.FontSize = nFontSize;
            _textPr2.Color = this.GetBodyTextColor(_master, _layout);
            var docContent = new CDocumentContent(editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, 0, 0, 1000, 1000, false, false, true);
            var par = docContent.Content[0];
            par.MoveCursorToStartPos();
            var _paraPr = new CParaPr;
            par.Pr = _paraPr;
            var parRun = new ParaRun(par);
            parRun.Set_Pr(_textPr1);
            parRun.AddText("A");
            par.Add_ToContent(0, parRun);
            parRun = new ParaRun(par);
            parRun.Set_Pr(_textPr2);
            parRun.AddText("a");
            par.Add_ToContent(1, parRun);
            par.Reset(0, 0, 1000, 1000, 0, 0, 1);
            par.Recalculate_Page(0);

            var _text_x = _params[5] * dKoefPixToMM;
            var _text_y = (h_px - _params[6]) * dKoefPixToMM;
            par.Lines[0].Ranges[0].XVisible = _text_x;
            par.Lines[0].Y = _text_y;
            var old_marks = _api.ShowParaMarks;
            _api.ShowParaMarks = false;
            par.Draw(0, g);
            _api.ShowParaMarks = old_marks;

            _api.isViewMode = _oldTurn;
        }, this, []);
    };

    this.Draw = function(g, _master, use_background, use_master_shapes) {

        /*
        var _params = [
            0, 0, // w/h - not used
            6,  // color_w
            3,  // color_h,
            4,  // color_x
            31, // color_y
            1,  // color_delta,
            8,  // text_x
            11, // text_y (from bottom)
            18 // font_size
        ];
        _params[9] *= ((this.HeightMM / this.HeightPx) * (96 / 25.4));


        for (var i = 0; i < _params.length; i++)
        {
            _params[i] = AscCommon.AscBrowser.convertToRetinaValue(_params[i], true);
        }

        return this.Draw2(g, _master, use_background, use_master_shapes, _params);
        */

        var w_px = this.WidthPx;
        var h_px = this.HeightPx;
        var dKoefPixToMM = this.HeightMM / h_px;
        var _back_fill = null;
        var RGBA = {R:0, G:0, B:0, A:255};
        var _layout = null;
        for (var i = 0; i < _master.sldLayoutLst.length; i++) {
          if (_master.sldLayoutLst[i].type == AscFormat.nSldLtTTitle) {
            _layout = _master.sldLayoutLst[i];
            break;
          }
        }
        var _theme = _master.Theme;
        if (_layout != null && _layout.cSld.Bg != null) {
          if (null != _layout.cSld.Bg.bgPr) {
            _back_fill = _layout.cSld.Bg.bgPr.Fill;
          } else {
            if (_layout.cSld.Bg.bgRef != null) {
              _layout.cSld.Bg.bgRef.Color.Calculate(_theme, null, _layout, _master, RGBA);
              RGBA = _layout.cSld.Bg.bgRef.Color.RGBA;
              _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(_layout.cSld.Bg.bgRef.idx, _layout.cSld.Bg.bgRef.Color);
            }
          }
        } else {
          if (_master != null) {
            if (_master.cSld.Bg != null) {
              if (null != _master.cSld.Bg.bgPr) {
                _back_fill = _master.cSld.Bg.bgPr.Fill;
              } else {
                if (_master.cSld.Bg.bgRef != null) {
                  _master.cSld.Bg.bgRef.Color.Calculate(_theme, null, _layout, _master, RGBA);
                  RGBA = _master.cSld.Bg.bgRef.Color.RGBA;
                  _back_fill = _theme.themeElements.fmtScheme.GetFillStyle(_master.cSld.Bg.bgRef.idx, _master.cSld.Bg.bgRef.Color);
                }
              }
            } else {
              _back_fill = new AscFormat.CUniFill;
              _back_fill.fill = new AscFormat.CSolidFill;
              _back_fill.fill.color = new AscFormat.CUniColor;
              _back_fill.fill.color.color = new AscFormat.CRGBColor;
              _back_fill.fill.color.color.RGBA = {R:255, G:255, B:255, A:255};
            }
          }
        }
        if (_back_fill != null) {
          _back_fill.calculate(_theme, null, _layout, _master, RGBA);
        }
        if (use_background !== false) {
          DrawBackground(g, _back_fill, this.WidthMM, this.HeightMM);
        }

        if (use_master_shapes !== false)
        {
            if (null == _layout)
            {
                if(_master.needRecalc && _master.needRecalc())
                {
                    _master.recalculate();
                }
                _master.drawNoPlaceholders(g);
            }
            else
            {
                if (_layout.showMasterSp)
                {
                    if(_master.needRecalc && _master.needRecalc())
                    {
                        _master.recalculate();
                    }
                    _master.drawNoPlaceholders(g);
                }
                _layout.recalculate();
                _layout.drawNoPlaceholders(g);
            }
        }
        g.reset();
        g.SetIntegerGrid(true);
        var _text_x = 8 * dKoefPixToMM;
        var _text_y = (h_px - 10) * dKoefPixToMM;
        var _color_w = 6;
        var _color_h = 3;
        var _color_x = 4;
        var _color_y = 31;
        var _color_delta = 1;
        if (!window["NATIVE_EDITOR_ENJINE"]) {
          _color_w = AscCommon.AscBrowser.convertToRetinaValue(_color_w, true);
          _color_h = AscCommon.AscBrowser.convertToRetinaValue(_color_h, true);
          _color_x = AscCommon.AscBrowser.convertToRetinaValue(_color_x, true);
          _color_y = AscCommon.AscBrowser.convertToRetinaValue(_color_y, true);
          _color_delta = AscCommon.AscBrowser.convertToRetinaValue(_color_delta, true);

          g.p_color(255, 255, 255, 255);
          g.init(g.m_oContext, w_px, h_px, w_px, h_px);
          g.CalculateFullTransform();
          g.m_bIntegerGrid = true;
          g.b_color1(255, 255, 255, 255);
          g._s();
          g.rect(_color_x - _color_delta, _color_y - _color_delta, _color_w * 6 + 7 * _color_delta, _color_h + 2 * _color_delta);
          g.df();
          g._s();
          var _color = new AscFormat.CSchemeColor;
          for (var i = 0; i < 6; i++) {
            g._s();
            _color.id = i;
            _color.Calculate(_theme, null, null, _master, RGBA);
            g.b_color1(_color.RGBA.R, _color.RGBA.G, _color.RGBA.B, 255);
            g.rect(_color_x, _color_y, _color_w, _color_h);
            g.df();
            _color_x += _color_w + _color_delta;
          }
          g._s();
        } else {
          _color_w = this.WidthMM/8.0;
          _color_h = this.HeightMM/10.0;
          _color_x = this.WidthMM/20.0;
          _color_y = this.HeightMM - _color_x*(w_px/this.WidthMM)*(this.HeightMM/h_px) - _color_h;
          _color_delta = 2 * dKoefPixToMM;
            var __color_x = _color_x;
          g.p_color(255, 255, 255, 255);
          g.m_bIntegerGrid = true;
          g.b_color1(255, 255, 255, 255);
          g._s();
          g.rect(_color_x - _color_delta, _color_y - _color_delta, _color_w * 6 + 7 * _color_delta, _color_h + 2 * _color_delta);
          g.df();
          g._s();
          var _color = new AscFormat.CSchemeColor;
          for (var i = 0; i < 6; i++) {
            g._s();
            _color.id = i;
            _color.Calculate(_theme, null, null, _master, RGBA);
            g.b_color1(_color.RGBA.R, _color.RGBA.G, _color.RGBA.B, 255);
            g.rect(_color_x, _color_y, _color_w, _color_h);
            g.df();
            _color_x += _color_w + _color_delta;
          }
          g._s();
          _color_x = __color_x;
        }
        var _api = this.DrawingDocument.m_oWordControl.m_oApi;
        AscFormat.ExecuteNoHistory(function(){
            var _oldTurn = _api.isViewMode;
            _api.isViewMode = true;
            _color.id = 15;
            _color.Calculate(_theme, null, null, _master, RGBA);
            var nFontSize = 18;
            if (window["NATIVE_EDITOR_ENJINE"]) {
                nFontSize = 600;
            }
            var _textPr1 = new CTextPr;
            _textPr1.FontFamily = {Name:_theme.themeElements.fontScheme.majorFont.latin, Index:-1};
            _textPr1.RFonts.Ascii = {Name: _theme.themeElements.fontScheme.majorFont.latin, Index: -1};
            _textPr1.FontSize = nFontSize;
            _textPr1.Color = this.GetTitleTextColor(_master, _layout);
            var _textPr2 = new CTextPr;
            _textPr2.FontFamily = {Name:_theme.themeElements.fontScheme.minorFont.latin, Index:-1};
            _textPr2.RFonts.Ascii = {Name: _theme.themeElements.fontScheme.minorFont.latin, Index: -1};
            _textPr2.FontSize = nFontSize;
            _textPr2.Color = this.GetBodyTextColor(_master, _layout);
            var docContent = new CDocumentContent(editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, 0, 0, 1000, 1000, false, false, true);
            var par = docContent.Content[0];
            par.MoveCursorToStartPos();
            var _paraPr = new CParaPr;
            par.Pr = _paraPr;
            var parRun = new ParaRun(par);
            parRun.Set_Pr(_textPr1);
            parRun.AddText("A");
            par.Add_ToContent(0, parRun);
            parRun = new ParaRun(par);
            parRun.Set_Pr(_textPr2);
            parRun.AddText("a");
            par.Add_ToContent(1, parRun);
            par.Reset(0, 0, 1000, 1000, 0, 0, 1);
            par.Recalculate_Page(0);
            if (!window["NATIVE_EDITOR_ENJINE"]) {
                var koefFont = AscCommon.g_dKoef_pix_to_mm / AscCommon.AscBrowser.retinaPixelRatio;
                g.init(g.m_oContext, w_px, h_px, w_px * koefFont, h_px * koefFont);
                g.CalculateFullTransform();
                _text_x = AscCommon.AscBrowser.retinaPixelRatio * 8 * koefFont;
                _text_y = (h_px - 11 * AscCommon.AscBrowser.retinaPixelRatio) * koefFont;
                par.Lines[0].Ranges[0].XVisible = _text_x;
                par.Lines[0].Y = _text_y;
                var old_marks = _api.ShowParaMarks;
                _api.ShowParaMarks = false;
                par.Draw(0, g);
                _api.ShowParaMarks = old_marks;
            } else {
                _text_x = _color_x;
                _text_y = _color_y - _color_h;
                par.Lines[0].Ranges[0].XVisible = _text_x;
                par.Lines[0].Y = _text_y;
                var old_marks = _api.ShowParaMarks;
                _api.ShowParaMarks = false;
                par.Draw(0, g);
                _api.ShowParaMarks = old_marks;
            }
            _api.isViewMode = _oldTurn;
        }, this, []);
      };

    this.GetThumbnail = function(_master, use_background, use_master_shapes)
    {
        if(window["NATIVE_EDITOR_ENJINE"])
        {
            return "";
        }

        var w_px = AscCommon.AscBrowser.convertToRetinaValue(AscCommon.GlobalSkin.THEMES_THUMBNAIL_WIDTH, true);
        var h_px = AscCommon.AscBrowser.convertToRetinaValue(AscCommon.GlobalSkin.THEMES_THUMBNAIL_HEIGHT, true);

        this.WidthPx  = w_px;
        this.HeightPx = h_px;

        if (this.CanvasImage == null)
            this.CanvasImage = document.createElement('canvas');

        this.CanvasImage.width = w_px;
        this.CanvasImage.height = h_px;

        var _ctx = this.CanvasImage.getContext('2d');

        var g = new AscCommon.CGraphics();
        g.init(_ctx, w_px, h_px, this.WidthMM, this.HeightMM);
        g.m_oFontManager = AscCommon.g_fontManager;

        g.transform(1,0,0,1,0,0);

        this.Draw(g, _master, use_background, use_master_shapes);

        try
        {
            return this.CanvasImage.toDataURL("image/png");
        }
        catch (err)
        {
            this.CanvasImage = null;
            if (undefined === use_background && undefined === use_master_shapes)
                return this.GetThumbnail(_master, true, false);
            else if (use_background && !use_master_shapes)
                return this.GetThumbnail(_master, false, false);
        }
        return "";
    }
}

function fFillFromCSld(oSlideLikeObject, oCSld) {
    for(var i = 0; i < oCSld.spTree.length; ++i)
    {
        oSlideLikeObject.addToSpTreeToPos(i, oCSld.spTree[i]);
    }
    if(oCSld.Bg)
    {
        oSlideLikeObject.changeBackground(oCSld.Bg);
    }
    oSlideLikeObject.setCSldName(oCSld.name);
	oSlideLikeObject.changeBackground(oCSld.Bg);
}

const DEFAULT_SLIDE_W = 338.6666666666667;
const DEFAULT_SLIDE_H = 190.5;
AscCommonSlide.DEFAULT_MASTER_BINARY = "PPTY;v10;8302;FmkgAAD6+wB/DAAA+vsAHAAAAPr7ARUAAAD6AOkDAAD7AAkAAAADBAAAAPoABvsBVwwAAARSDAAAAC0AAAAADAAAAPoAAQAAAAEAAAAA+wECAAAA+vsCEAAAAPr7AQAAAAACBAAAAAAAAAABOwAAAPr7ACoAAAD6AAAAAAABAAAAAAIAAAAAAwAAAAAEAAAAAAUAAAAABgAAAAAHAAAAAPsBAAAAAAIAAAAAAtsLAAAFAAAAAJsBAAABlgEAAPr7AF4AAAAAMgAAAPoAAgAAAAETAAAAVABpAHQAbABlACAAUABsAGEAYwBlAGgAbwBsAGQAZQByACAAMQD7AQQAAAD6BgH7AhkAAAD6+wAEAAAA+gQP+wEAAAAAAgQAAAAAAAAAAUkAAAD6+wAWAAAA+gA4ygwAAUWSBQACkHSgAAP7ORQA+wEdAAAAARgAAAD6AAQAAAByAGUAYwB0APsABAAAAAAAAAACAAAAAAQAAAAAA94AAAAAKAAAAPoBAQOYsgAACDBlAQAKMGUBAAwAD5iyAAARAfsBBwAAAPoAAwAAAPsBAAAAAAKnAAAAAQAAAACeAAAAAR0AAAD6BQAKBQAAAGUAbgAtAFUAUwD7AQAAAAACAAAAAAJ3AAAAAQAAAABuAAAAAWkAAAD6ACAAAABDAGwAaQBjAGsAIAB0AG8AIABlAGQAaQB0ACAATQBhAHMAdABlAHIAIAB0AGkAdABsAGUAIABzAHQAeQBsAGUA+wAdAAAA+goFAAAAZQBuAC0AVQBTAA0A+wEAAAAAAgAAAAAAywMAAAHGAwAA+vsAYwAAAAAwAAAA+gADAAAAARIAAABUAGUAeAB0ACAAUABsAGEAYwBlAGgAbwBsAGQAZQByACAAMgD7AQQAAAD6BgH7AiAAAAD6+wALAAAA+gEBAAAAMQAEAPsBAAAAAAIEAAAAAAAAAAFJAAAA+vsAFgAAAPoAOMoMAAFZ2xsAApB0oAADamVCAPsBHQAAAAEYAAAA+gAEAAAAcgBlAGMAdAD7AAQAAAAAAAAAAgAAAAAEAAAAAAMJAwAAACYAAAD6A5iyAAAIMGUBAAowZQEADAAPmLIAABEB+wEHAAAA+gADAAAA+wEAAAAAAtQCAAAFAAAAAKUAAAAAJAAAAPoHAAAAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAJ3AAAAAQAAAABuAAAAAWkAAAD6ACAAAABDAGwAaQBjAGsAIAB0AG8AIABlAGQAaQB0ACAATQBhAHMAdABlAHIAIAB0AGUAeAB0ACAAcwB0AHkAbABlAHMA+wAdAAAA+goFAAAAZQBuAC0AVQBTAA0A+wEAAAAAAgAAAAAAfQAAAAAkAAAA+gcBAAAA+wMAAAAABAAAAAAFAAAAAAYAAAAABwQAAAAAAAAAAk8AAAABAAAAAEYAAAABQQAAAPoADAAAAFMAZQBjAG8AbgBkACAAbABlAHYAZQBsAPsAHQAAAPoKBQAAAGUAbgAtAFUAUwANAPsBAAAAAAIAAAAAAHsAAAAAJAAAAPoHAgAAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAJNAAAAAQAAAABEAAAAAT8AAAD6AAsAAABUAGgAaQByAGQAIABsAGUAdgBlAGwA+wAdAAAA+goFAAAAZQBuAC0AVQBTAA0A+wEAAAAAAgAAAAAAfQAAAAAkAAAA+gcDAAAA+wMAAAAABAAAAAAFAAAAAAYAAAAABwQAAAAAAAAAAk8AAAABAAAAAEYAAAABQQAAAPoADAAAAEYAbwB1AHIAdABoACAAbABlAHYAZQBsAPsAHQAAAPoKBQAAAGUAbgAtAFUAUwANAPsBAAAAAAIAAAAAAJ0AAAAAJAAAAPoHBAAAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAEdAAAA+gUACgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAACTQAAAAEAAAAARAAAAAE/AAAA+gALAAAARgBpAGYAdABoACAAbABlAHYAZQBsAPsAHQAAAPoKBQAAAGUAbgAtAFUAUwANAPsBAAAAAAIAAAAAAGECAAABXAIAAPr7AGUAAAAAMAAAAPoABAAAAAESAAAARABhAHQAZQAgAFAAbABhAGMAZQBoAG8AbABkAGUAcgAgADMA+wEEAAAA+gYB+wIiAAAA+vsADQAAAPoBAQAAADIAAwEEBfsBAAAAAAIEAAAAAAAAAAFJAAAA+vsAFgAAAPoAOMoMAAF+/WAAAqDbKQADRZIFAPsBHQAAAAEYAAAA+gAEAAAAcgBlAGMAdAD7AAQAAAAAAAAAAgAAAAAEAAAAAAOdAQAAACgAAAD6AQEDmLIAAAgwZQEACjBlAQAMAA+YsgAAEQH7AQcAAAD6AAAAAAD7AXUAAAAAcAAAAPoABPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhKAAAA+hGwBAAA+wE5AAAAAzQAAAAALwAAAAMqAAAA+gAP+wAhAAAAAQAAAAEYAAAA+gAGAAAAYQA6AHQAaQBuAHQAAfgkAQD7AgAAAAAC8QAAAAEAAAAA6AAAAAEbAAAA+goFAAAAZQBuAC0AVQBTAPsBAAAAAAIAAAAAAsMAAAABAAAAALoAAAACtQAAAPoAJgAAAHsAQgBDAEMAMQA4AEYANQAxAC0AMAA5AEUAQwAtADQAMwA1AEMALQBBADMAQgBBAC0ANgA0AEEANwA2ADYARQAwADkAOQBDADAAfQABEQAAAGQAYQB0AGUAdABpAG0AZQBGAGkAZwB1AHIAZQBPAHUAdAACCgAAADMAMAAuADEAMAAuADIAMAAxADMA+wAdAAAA+goFAAAAZQBuAC0AVQBTAA0A+wEAAAAAAgAAAAAApgEAAAGhAQAA+vsAaQAAAAA0AAAA+gAFAAAAARQAAABGAG8AbwB0AGUAcgAgAFAAbABhAGMAZQBoAG8AbABkAGUAcgAgADQA+wEEAAAA+gYB+wIiAAAA+vsADQAAAPoBAQAAADMAAwIEBvsBAAAAAAIEAAAAAAAAAAFJAAAA+vsAFgAAAPoAyJ89AAF+/WAAAnDJPgADRZIFAPsBHQAAAAEYAAAA+gAEAAAAcgBlAGMAdAD7AAQAAAAAAAAAAgAAAAAEAAAAAAPeAAAAACgAAAD6AQEDmLIAAAgwZQEACjBlAQAMAA+YsgAAEQH7AQcAAAD6AAAAAAD7AXUAAAAAcAAAAPoAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhKAAAA+hGwBAAA+wE5AAAAAzQAAAAALwAAAAMqAAAA+gAP+wAhAAAAAQAAAAEYAAAA+gAGAAAAYQA6AHQAaQBuAHQAAfgkAQD7AgAAAAACMgAAAAEAAAAAKQAAAAEbAAAA+goFAAAAZQBuAC0AVQBTAPsBAAAAAAIAAAAAAgQAAAAAAAAAAFECAAABTAIAAPr7AHUAAAAAQAAAAPoABgAAAAEaAAAAUwBsAGkAZABlACAATgB1AG0AYgBlAHIAIABQAGwAYQBjAGUAaABvAGwAZABlAHIAIAA1APsBBAAAAPoGAfsCIgAAAPr7AA0AAAD6AQEAAAA0AAMCBAz7AQAAAAACBAAAAAAAAAABSQAAAPr7ABYAAAD6AChjgwABfv1gAAKg2ykAA0WSBQD7AR0AAAABGAAAAPoABAAAAHIAZQBjAHQA+wAEAAAAAAAAAAIAAAAABAAAAAADfQEAAAAoAAAA+gEBA5iyAAAIMGUBAAowZQEADAAPmLIAABEB+wEHAAAA+gAAAAAA+wF1AAAAAHAAAAD6AAX7AwAAAAAEAAAAAAUAAAAABgAAAAAHBAAAAAAAAAAISgAAAPoRsAQAAPsBOQAAAAM0AAAAAC8AAAADKgAAAPoAD/sAIQAAAAEAAAABGAAAAPoABgAAAGEAOgB0AGkAbgB0AAH4JAEA+wIAAAAAAtEAAAABAAAAAMgAAAABGwAAAPoKBQAAAGUAbgAtAFUAUwD7AQAAAAACAAAAAAKjAAAAAQAAAACaAAAAApUAAAD6ACYAAAB7ADAAOAAzADkANQA1ADgANgAtAEYAMAAzAEEALQA0ADgARAAxAC0AOQA0AEQARgAtADEANgBCADIAMwA5AEQARgA0AEYAQgA1AH0AAQgAAABzAGwAaQBkAGUAbgB1AG0AAgMAAAA5ICMAOiD7AB0AAAD6CgUAAABlAG4ALQBVAFMADQD7AQAAAAACAAAAAAEaAAAA+gYMDwgHDRAJAAABAQICAwMEBAUFCwsKCvsC9wEAAAsAAAAAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA3ADMAAQgAAAByAEkAZAAxAPsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA3ADQAAQgAAAByAEkAZAAyAPsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA3ADUAAQgAAAByAEkAZAAzAPsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA3ADYAAQgAAAByAEkAZAA0APsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA3ADcAAQgAAAByAEkAZAA1APsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA3ADgAAQgAAAByAEkAZAA2APsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA3ADkAAQgAAAByAEkAZAA3APsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA4ADAAAQgAAAByAEkAZAA4APsAKAAAAPoACgAAADIAMQA0ADcANAA4ADMANgA4ADEAAQgAAAByAEkAZAA5APsAKgAAAPoACgAAADIAMQA0ADcANAA4ADMANgA4ADIAAQoAAAByAEkAZAAxADAA+wAqAAAA+gAKAAAAMgAxADQANwA0ADgAMwA2ADgAMwABCgAAAHIASQBkADEAMQD7BsMRAAAAxgAAAADBAAAA+gAEAeDzDQACAQQBBgAKAPsABwAAAPoAkF8BAPsCBwAAAPoAAAAAAPsDAAAAAAQAAAAABQAAAAAGBQAAAAAAAAAABwQAAAAAAAAACHEAAAD6CLAEAAARMBEAAPsBEwAAAAMOAAAAAAkAAAADBAAAAPoAD/sCAAAAAAMTAAAA+gMGAAAAKwBtAGoALQBsAHQA+wQTAAAA+gMGAAAAKwBtAGoALQBlAGEA+wUTAAAA+gMGAAAAKwBtAGoALQBjAHMA+wGMCgAAACcBAAD6AAQB4PMNAAIBBAEFCIP8/wYACPh8AwAKAPsABwAAAPoAkF8BAPsCBwAAAPoB6AMAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABHwCgAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7AScBAAD6AAQB4PMNAAIBBAEFCIP8/wYACOh2CgAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABFgCQAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7AicBAAD6AAQB4PMNAAIBBAEFCIP8/wYACNhwEQAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABHQBwAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7AycBAAD6AAQB4PMNAAIBBAEFCIP8/wYACMhqGAAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABEIBwAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7BCcBAAD6AAQB4PMNAAIBBAEFCIP8/wYACLhkHwAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABEIBwAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7BScBAAD6AAQB4PMNAAIBBAEFCIP8/wYACKheJgAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABEIBwAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7BicBAAD6AAQB4PMNAAIBBAEFCIP8/wYACJhYLQAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABEIBwAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7BycBAAD6AAQB4PMNAAIBBAEFCIP8/wYACIhSNAAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABEIBwAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7CCcBAAD6AAQB4PMNAAIBBAEFCIP8/wYACHhMOwAKAPsABwAAAPoAkF8BAPsCBwAAAPoB9AEAAPsDAAAAAAQAAAAABVMAAAACTgAAAPoAAQAAADAAARQAAAAwADIAMABCADAANgAwADQAMAAyADAAMgAwADIAMAAyADAAMgAwADQAAgIAAAAzADQAAwUAAABBAHIAaQBhAGwA+wYOAAAAAQkAAAD6AAEAAAAiIPsHBAAAAAAAAAAIcQAAAPoIsAQAABEIBwAA+wETAAAAAw4AAAAACQAAAAMEAAAA+gAP+wIAAAAAAxMAAAD6AwYAAAArAG0AbgAtAGwAdAD7BBMAAAD6AwYAAAArAG0AbgAtAGUAYQD7BRMAAAD6AwYAAAArAG0AbgAtAGMAcwD7AmIGAAAAqQAAAPoABAHg8w0AAgEEAQYACAAAAAAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsBqQAAAPoABAHg8w0AAgEEAQYACPD5BgAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsCqQAAAPoABAHg8w0AAgEEAQYACODzDQAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsDqQAAAPoABAHg8w0AAgEEAQYACNDtFAAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsEqQAAAPoABAHg8w0AAgEEAQYACMDnGwAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsFqQAAAPoABAHg8w0AAgEEAQYACLDhIgAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsGqQAAAPoABAHg8w0AAgEEAQYACKDbKQAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsHqQAAAPoABAHg8w0AAgEEAQYACJDVMAAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsIqQAAAPoABAHg8w0AAgEEAQYACIDPNwAKAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAhxAAAA+giwBAAAEQgHAAD7ARMAAAADDgAAAAAJAAAAAwQAAAD6AA/7AgAAAAADEwAAAPoDBgAAACsAbQBuAC0AbAB0APsEEwAAAPoDBgAAACsAbQBuAC0AZQBhAPsFEwAAAPoDBgAAACsAbQBuAC0AYwBzAPsJPwAAAPr7AwAAAAAEAAAAAAUAAAAABgAAAAAHBAAAAAAAAAAIGwAAAPoKBQAAAGUAbgAtAFUAUwD7AQAAAAACAAAAAA==";
AscCommonSlide.DEFAULT_LAYOUTS_BINARY = "PPTY;v10;32844;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";
AscCommonSlide.DEFAULT_THEME_BINARY = "PPTY;v10;4972;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";
AscCommonSlide.PH_BINARIES = [];
AscCommonSlide.PH_BODY_BINARY = "PPTY;v10;881;bQMAAPr7AGsAAAAANgAAAPoACAAAAAEVAAAAQwBvAG4AdABlAG4AdAAgAFAAbABhAGMAZQBoAG8AbABkAGUAcgAgADcA+wEEAAAA+gYB+wIiAAAA+vsADQAAAPoBAgAAADEAMAADAvsBAAAAAAIEAAAAAAAAAAEsAAAA+vsAFgAAAPoADBAFAAHAGQMAArEdMQADPtciAPsBAAAAAAIAAAAABAAAAAADxQIAAAAOAAAA+vsBBwAAAPoAAAAAAPsBAAAAAAKoAgAABQAAAACjAAAAACQAAAD6BwAAAAD7AwAAAAAEAAAAAAUAAAAABgAAAAAHBAAAAAAAAAACdQAAAAEAAAAAbAAAAAFnAAAA+gAgAAAAQwBsAGkAYwBrACAAdABvACAAZQBkAGkAdAAgAE0AYQBzAHQAZQByACAAdABlAHgAdAAgAHMAdAB5AGwAZQBzAPsAGwAAAPoKBQAAAGUAbgAtAFUAUwD7AQAAAAACAAAAAAB7AAAAACQAAAD6BwEAAAD7AwAAAAAEAAAAAAUAAAAABgAAAAAHBAAAAAAAAAACTQAAAAEAAAAARAAAAAE/AAAA+gAMAAAAUwBlAGMAbwBuAGQAIABsAGUAdgBlAGwA+wAbAAAA+goFAAAAZQBuAC0AVQBTAPsBAAAAAAIAAAAAAHkAAAAAJAAAAPoHAgAAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAJLAAAAAQAAAABCAAAAAT0AAAD6AAsAAABUAGgAaQByAGQAIABsAGUAdgBlAGwA+wAbAAAA+goFAAAAZQBuAC0AVQBTAPsBAAAAAAIAAAAAAHsAAAAAJAAAAPoHAwAAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAJNAAAAAQAAAABEAAAAAT8AAAD6AAwAAABGAG8AdQByAHQAaAAgAGwAZQB2AGUAbAD7ABsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAAAeQAAAAAkAAAA+gcEAAAA+wMAAAAABAAAAAAFAAAAAAYAAAAABwQAAAAAAAAAAksAAAABAAAAAEIAAAABPQAAAPoACwAAAEYAaQBmAHQAaAAgAGwAZQB2AGUAbAD7ABsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAA=";
AscCommonSlide.PH_BINARIES[AscFormat.phType_body] = "PPTY;v10;877;aQMAAPr7AGcAAAAAMAAAAPoACgAAAAESAAAAVABlAHgAdAAgAFAAbABhAGMAZQBoAG8AbABkAGUAcgAgADkA+wEEAAAA+gYB+wIkAAAA+vsADwAAAPoBAgAAADEAMQADAgQA+wEAAAAAAgQAAAAAAAAAASwAAAD6+wAWAAAA+gBM4TsAAcNXAwACVzssAAOshCIA+wEAAAAAAgAAAAAEAAAAAAPFAgAAAA4AAAD6+wEHAAAA+gAAAAAA+wEAAAAAAqgCAAAFAAAAAKMAAAAAJAAAAPoHAAAAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAJ1AAAAAQAAAABsAAAAAWcAAAD6ACAAAABDAGwAaQBjAGsAIAB0AG8AIABlAGQAaQB0ACAATQBhAHMAdABlAHIAIAB0AGUAeAB0ACAAcwB0AHkAbABlAHMA+wAbAAAA+goFAAAAZQBuAC0AVQBTAPsBAAAAAAIAAAAAAHsAAAAAJAAAAPoHAQAAAPsDAAAAAAQAAAAABQAAAAAGAAAAAAcEAAAAAAAAAAJNAAAAAQAAAABEAAAAAT8AAAD6AAwAAABTAGUAYwBvAG4AZAAgAGwAZQB2AGUAbAD7ABsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAAAeQAAAAAkAAAA+gcCAAAA+wMAAAAABAAAAAAFAAAAAAYAAAAABwQAAAAAAAAAAksAAAABAAAAAEIAAAABPQAAAPoACwAAAFQAaABpAHIAZAAgAGwAZQB2AGUAbAD7ABsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAAAewAAAAAkAAAA+gcDAAAA+wMAAAAABAAAAAAFAAAAAAYAAAAABwQAAAAAAAAAAk0AAAABAAAAAEQAAAABPwAAAPoADAAAAEYAbwB1AHIAdABoACAAbABlAHYAZQBsAPsAGwAAAPoKBQAAAGUAbgAtAFUAUwD7AQAAAAACAAAAAAB5AAAAACQAAAD6BwQAAAD7AwAAAAAEAAAAAAUAAAAABgAAAAAHBAAAAAAAAAACSwAAAAEAAAAAQgAAAAE9AAAA+gALAAAARgBpAGYAdABoACAAbABlAHYAZQBsAPsAGwAAAPoKBQAAAGUAbgAtAFUAUwD7AQAAAAACAAAAAA==";
AscCommonSlide.PH_BINARIES[AscFormat.phType_pic] = "PPTY;v10;257;/QAAAPr7AG8AAAAAOAAAAPoADAAAAAEWAAAAUABpAGMAdAB1AHIAZQAgAFAAbABhAGMAZQBoAG8AbABkAGUAcgAgADEAMQD7AQQAAAD6BgH7AiQAAAD6+wAPAAAA+gECAAAAMQAyAAMCBAr7AQAAAAACBAAAAAAAAAABLAAAAPr7ABYAAAD6AAwQBQABD4MqAAK0PiMAA3Z/FwD7AQAAAAACAAAAAAQAAAAAA1EAAAAADgAAAPr7AQcAAAD6AAAAAAD7AQAAAAACNAAAAAEAAAAAKwAAAAEdAAAA+gUACgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAACBAAAAAAAAAA=";
AscCommonSlide.PH_BINARIES[AscFormat.phType_chart] = "PPTY;v10;251;9wAAAPr7AGsAAAAANAAAAPoADgAAAAEUAAAAQwBoAGEAcgB0ACAAUABsAGEAYwBlAGgAbwBsAGQAZQByACAAMQAzAPsBBAAAAPoGAfsCJAAAAPr7AA8AAAD6AQIAAAAxADMAAwIEAfsBAAAAAAIEAAAAAAAAAAEsAAAA+vsAFgAAAPoA4xcrAAEPgyoAAmkwHQADI38XAPsBAAAAAAIAAAAABAAAAAADTwAAAAAOAAAA+vsBBwAAAPoAAAAAAPsBAAAAAAIyAAAAAQAAAAApAAAAARsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAACBAAAAAAAAAA=";
AscCommonSlide.PH_BINARIES[AscFormat.phType_tbl] = "PPTY;v10;251;9wAAAPr7AGsAAAAANAAAAPoAEAAAAAEUAAAAVABhAGIAbABlACAAUABsAGEAYwBlAGgAbwBsAGQAZQByACAAMQA1APsBBAAAAPoGAfsCJAAAAPr7AA8AAAD6AQIAAAAxADQAAwIEDvsBAAAAAAIEAAAAAAAAAAEsAAAA+vsAFgAAAPoA3bhLAAGplSoAAn6rGgADYD8YAPsBAAAAAAIAAAAABAAAAAADTwAAAAAOAAAA+vsBBwAAAPoAAAAAAPsBAAAAAAIyAAAAAQAAAAApAAAAARsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAACBAAAAAAAAAA=";
AscCommonSlide.PH_BINARIES[AscFormat.phType_dgm] = "PPTY;v10;257;/QAAAPr7AHEAAAAAOgAAAPoAEgAAAAEXAAAAUwBtAGEAcgB0AEEAcgB0ACAAUABsAGEAYwBlAGgAbwBsAGQAZQByACAAMQA3APsBBAAAAPoGAfsCJAAAAPr7AA8AAAD6AQIAAAAxADUAAwIEBPsBAAAAAAIEAAAAAAAAAAEsAAAA+vsAFgAAAPoAYPhqAAE77ikAAmowHQAD16AZAPsBAAAAAAIAAAAABAAAAAADTwAAAAAOAAAA+vsBBwAAAPoAAAAAAPsBAAAAAAIyAAAAAQAAAAApAAAAARsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAACBAAAAAAAAAA=";
AscCommonSlide.PH_BINARIES[AscFormat.phType_title] = "PPTY;v10;461;yQEAAPr7AF4AAAAAMgAAAPoAAgAAAAETAAAAVABpAHQAbABlACAAUABsAGEAYwBlAGgAbwBsAGQAZQByACAAMQD7AQQAAAD6BgH7AhkAAAD6+wAEAAAA+gQP+wEAAAAAAgQAAAAAAAAAAUsAAAD6AAD7ABYAAAD6ADjKDAABRZIFAAKQdKAAA/s5FAD7AR0AAAABGAAAAPoABAAAAHIAZQBjAHQA+wAEAAAAAAAAAAIAAAAABAAAAAADDwEAAAAoAAAA+gEBA5iyAAAIMGUBAAowZQEADAAPmLIAABEB+wEHAAAA+gADAAAA+wEAAAAAAtgAAAABAAAAAM8AAAAAMAAAAPr7AwAAAAAEAAAAAAUAAAAABgAAAAAHBAAAAAAAAAAIDAAAAPr7AQAAAAACAAAAAAEbAAAA+goFAAAAZQBuAC0AVQBTAPsBAAAAAAIAAAAAAnUAAAABAAAAAGwAAAABZwAAAPoAIAAAAEMAbABpAGMAawAgAHQAbwAgAGUAZABpAHQAIABNAGEAcwB0AGUAcgAgAHQAaQB0AGwAZQAgAHMAdAB5AGwAZQD7ABsAAAD6CgUAAABlAG4ALQBVAFMA+wEAAAAAAgAAAAA=";
function CreateDefaultMaster() {
    let stream = AscFormat.CreateBinaryReader(AscCommonSlide.DEFAULT_MASTER_BINARY, "PPTY;v10;".length, AscCommonSlide.DEFAULT_MASTER_BINARY.length);
    let oBinaryReader = new AscCommon.BinaryPPTYLoader();
    oBinaryReader.stream = new AscCommon.FileStream();
    oBinaryReader.stream.obj = stream.obj;
    oBinaryReader.stream.data = stream.data;
    oBinaryReader.stream.size = stream.size;
    oBinaryReader.stream.pos = stream.pos;
    oBinaryReader.stream.cur = stream.cur;
    let oMaster = oBinaryReader.ReadSlideMaster();

    oMaster.setSlideSize(DEFAULT_SLIDE_W, DEFAULT_SLIDE_H);
    stream = AscFormat.CreateBinaryReader(AscCommonSlide.DEFAULT_LAYOUTS_BINARY, "PPTY;v10;".length, AscCommonSlide.DEFAULT_LAYOUTS_BINARY.length);
    oBinaryReader = new AscCommon.BinaryPPTYLoader();
    oBinaryReader.stream = new AscCommon.FileStream();
    oBinaryReader.stream.obj = stream.obj;
    oBinaryReader.stream.data = stream.data;
    oBinaryReader.stream.size = stream.size;
    oBinaryReader.stream.pos = stream.pos;
    oBinaryReader.stream.cur = stream.cur;

    let _sl_count = oBinaryReader.stream.GetULong();
    let oPresentation = Asc.editor.private_GetLogicDocument();
    for (let i = 0; i < _sl_count; i++) {
        let oLt = oBinaryReader.ReadSlideLayout();
        oLt.setSlideSize(DEFAULT_SLIDE_W, DEFAULT_SLIDE_H);
        oMaster.addToSldLayoutLstToPos(oMaster.sldLayoutLst.length, oLt);
    }

    stream = AscFormat.CreateBinaryReader(AscCommonSlide.DEFAULT_THEME_BINARY, "PPTY;v10;".length, AscCommonSlide.DEFAULT_THEME_BINARY.length);
    oBinaryReader = new AscCommon.BinaryPPTYLoader();
    oBinaryReader.stream = new AscCommon.FileStream();
    oBinaryReader.stream.obj = stream.obj;
    oBinaryReader.stream.data = stream.data;
    oBinaryReader.stream.size = stream.size;
    oBinaryReader.stream.pos = stream.pos;
    oBinaryReader.stream.cur = stream.cur;
    let oTheme = oBinaryReader.ReadTheme();
    oTheme.presentation = oPresentation;
    oMaster.setTheme(oTheme);
    return oMaster;
}

function CreateDefaultLayout(oMaster) {
    let stream = AscFormat.CreateBinaryReader(AscCommonSlide.DEFAULT_LAYOUTS_BINARY, "PPTY;v10;".length, AscCommonSlide.DEFAULT_LAYOUTS_BINARY.length);
    let oBinaryReader = new AscCommon.BinaryPPTYLoader();
    oBinaryReader.stream = new AscCommon.FileStream();
    oBinaryReader.stream.obj = stream.obj;
    oBinaryReader.stream.data = stream.data;
    oBinaryReader.stream.size = stream.size;
    oBinaryReader.stream.pos = stream.pos;
    oBinaryReader.stream.cur = stream.cur;

    let _sl_count = oBinaryReader.stream.GetULong();
    let oPresentation = Asc.editor.private_GetLogicDocument();

    oBinaryReader.stream.Skip2(1);
    let end = oBinaryReader.stream.cur + oBinaryReader.stream.GetULong() + 4;
    oBinaryReader.stream.Seek2(end);
    let oLt = oBinaryReader.ReadSlideLayout();
    oLt.setSlideSize(DEFAULT_SLIDE_W, DEFAULT_SLIDE_H);
    oLt.setMaster(oMaster);
    return oLt;

}

function CreatePlaceholder(nType, bVertical) {
    let sBinary = AscCommonSlide.PH_BINARIES[nType];
    if(!sBinary) {
        sBinary = AscCommonSlide.PH_BODY_BINARY;
    }
    let stream = AscFormat.CreateBinaryReader(sBinary, "PPTY;v10;".length, sBinary.length);
    let oBinaryReader = new AscCommon.BinaryPPTYLoader();
    oBinaryReader.stream = new AscCommon.FileStream();
    oBinaryReader.stream.obj = stream.obj;
    oBinaryReader.stream.data = stream.data;
    oBinaryReader.stream.size = stream.size;
    oBinaryReader.stream.pos = stream.pos;
    oBinaryReader.stream.cur = stream.cur;
    let oSp = oBinaryReader.ReadShape();
    if(bVertical) {
        let oBodyPr = oSp.txBody && oSp.txBody.bodyPr;
        if(oBodyPr) {
            oBodyPr = oBodyPr.createDuplicate();
            oBodyPr.setVert(AscFormat.nVertTTvert);
            oSp.txBody.setBodyPr(oBodyPr);
        }
    }
    let presentation = Asc.editor.private_GetLogicDocument();
    let w = presentation.GetWidthMM();
    let h = presentation.GetHeightMM();
    if(!AscFormat.fApproxEqual(w, DEFAULT_SLIDE_W, 0.1) ||
        !AscFormat.fApproxEqual(h, DEFAULT_SLIDE_H, 0.1)) {
        let scaleW = w / DEFAULT_SLIDE_W;
        let scaleH = h / DEFAULT_SLIDE_H;
        let oXfrm = oSp.spPr.xfrm;
        oXfrm.setOffX(oXfrm.offX * scaleW);
        oXfrm.setOffY(oXfrm.offY * scaleH);
        oXfrm.setExtX(oXfrm.extX * scaleW);
        oXfrm.setExtY(oXfrm.extY * scaleH);
    }
    return oSp;
}
//--------------------------------------------------------export----------------------------------------------------
window['AscCommonSlide'] = window['AscCommonSlide'] || {};
window['AscCommonSlide'].MasterSlide = MasterSlide;
window['AscCommonSlide'].fFillFromCSld = fFillFromCSld;
window['AscCommonSlide'].CreateDefaultMaster = CreateDefaultMaster;
window['AscCommonSlide'].CreateDefaultLayout = CreateDefaultLayout;
window['AscCommonSlide'].CreatePlaceholder = CreatePlaceholder;
