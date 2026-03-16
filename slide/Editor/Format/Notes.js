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

(function(){
    AscDFH.changesFactory[AscDFH.historyitem_NotesSetClrMap]  = AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_NotesSetShowMasterPhAnim]  = AscDFH.CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_NotesSetShowMasterSp]  = AscDFH.CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_NotesAddToSpTree]  = AscDFH.CChangesDrawingsContentPresentation;
    AscDFH.changesFactory[AscDFH.historyitem_NotesRemoveFromTree]  = AscDFH.CChangesDrawingsContentPresentation;
    AscDFH.changesFactory[AscDFH.historyitem_NotesSetBg]  = AscDFH.CChangesDrawingsObjectNoId;
    AscDFH.changesFactory[AscDFH.historyitem_NotesSetName]        = AscDFH.CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_NotesSetSlide]        = AscDFH.CChangesDrawingsObject;
    AscDFH.changesFactory[AscDFH.historyitem_NotesSetNotesMaster]        = AscDFH.CChangesDrawingsObject;


    AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetClrMap]  = function(oClass, value){oClass.clrMap = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetShowMasterPhAnim]  = function(oClass, value){oClass.showMasterPhAnim = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetShowMasterSp]  = function(oClass, value){oClass.showMasterSp = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetName]        = function(oClass, value){oClass.cSld.name = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetSlide]        = function(oClass, value){oClass.slide = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetNotesMaster]        = function(oClass, value){oClass.Master = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_NotesSetBg]  = function(oClass, value, FromLoad){
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

    AscDFH.drawingsConstructorsMap[AscDFH.historyitem_NotesSetBg]         = AscFormat.CBg;

    AscDFH.drawingContentChanges[AscDFH.historyitem_NotesAddToSpTree]    = function(oClass){return oClass.cSld.spTree;};
    AscDFH.drawingContentChanges[AscDFH.historyitem_NotesRemoveFromTree] = function(oClass){return oClass.cSld.spTree;};

    //Temporary function
    function GetNotesWidth(){
        return editor.WordControl.m_oDrawingDocument.Notes_GetWidth();
    }

    function CNotes() {
        AscCommonSlide.SlideBase.call(this);
        this.showMasterPhAnim = null;
        this.showMasterSp     = null;
        this.slide            = null;

        this.Master      = null;
			this.recalcInfo =
				{
					recalculateBackground: true,
					recalculateSpTree: true,
					recalculateSlide: true
				};

        this.kind = AscFormat.TYPE_KIND.NOTES;

				this.setSlideSize(this.presentation.GetNotesWidthMM(), this.presentation.GetNotesHeightMM());

        this.Lock = new AscCommon.CLock();
        this.graphicObjects = new AscFormat.DrawingObjectsController(this);
			this.deleteLock = new PropLocker(this.Id);
			this.backgroundLock = new PropLocker(this.Id);
			this.timingLock = new PropLocker(this.Id);
			this.transitionLock = new PropLocker(this.Id);
			this.layoutLock = new PropLocker(this.Id);
			this.showLock = new PropLocker(this.Id);
    }
    AscFormat.InitClass(CNotes, AscCommonSlide.SlideBase, AscDFH.historyitem_type_Notes);
    CNotes.prototype.setClMapOverride = function(pr){
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesSetClrMap, this.clrMap, pr));
        this.clrMap = pr;
    };

    CNotes.prototype.setShowMasterPhAnim = function(pr){
        History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_NotesSetShowMasterPhAnim, this.showMasterPhAnim, pr));
        this.showMasterPhAnim = pr;
    };

    CNotes.prototype.setShowMasterSp = function (pr) {
        History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_NotesSetShowMasterSp, this.showMasterSp, pr));
        this.showMasterSp = pr;
    };

    CNotes.prototype.removeFromSpTreeById = function(id){
        for(var i = this.cSld.spTree.length - 1; i > -1; --i){
            if(this.cSld.spTree[i].Get_Id() === id){
                this.removeFromSpTreeByPos(i);
            }
        }
    };


    CNotes.prototype.setCSldName = function(pr){
        History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_NotesSetName, this.cSld.name , pr));
        this.cSld.name = pr;
    };
    CNotes.prototype.setSlide = function(pr){
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesSetSlide, this.slide , pr));
        this.slide = pr;
    };

    CNotes.prototype.setNotesMaster = function(pr){
        History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_NotesSetNotesMaster, this.Master , pr));
        this.Master = pr;
    };

    CNotes.prototype.getMatchingShape = Slide.prototype.getMatchingShape;

    CNotes.prototype.getWidth = function(){
        return GetNotesWidth();
    };

	CNotes.prototype.createBodyShape = function () {
		const oSp = new AscFormat.CShape();
		oSp.setBDeleted(false);

		const oNvSpPr = new AscFormat.UniNvPr();

		const oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(3);
		oCNvPr.setName('Notes Placeholder 2');

		const oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_body);
		oPh.setIdx(1 + "");

		oNvSpPr.nvPr.setPh(oPh);
		oSp.setNvSpPr(oNvSpPr);
		oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
		oSp.setSpPr(new AscFormat.CSpPr());
		oSp.spPr.setParent(oSp);
		oSp.createTextBody();

		const oBodyPr = new AscFormat.CBodyPr();
		oSp.txBody.setBodyPr(oBodyPr);

		const oTxLstStyle = new AscFormat.TextListStyle();
		oSp.txBody.setLstStyle(oTxLstStyle);

		oSp.setParent(this);
		this.addToSpTreeToPos(1, oSp);

		return oSp;
	};

    CNotes.prototype.getBodyShape = function(){
        var aSpTree = this.cSld.spTree;
        for(var i = 0; i < aSpTree.length; ++i){
            var sp = aSpTree[i];
            if(sp.isPlaceholder()){
                if(sp.getPlaceholderType() === AscFormat.phType_body){
                    return sp;
                }
            }
        }
        return null;
    };

    CNotes.prototype.Refresh_RecalcData = function(){

    };

    CNotes.prototype.Refresh_RecalcData2 = function(){

    };

    CNotes.prototype.createDuplicate = function(IdMap){

        var oIdMap = IdMap || {};
        var oPr = new AscFormat.CCopyObjectProperties();
        oPr.idMap = oIdMap;
        var copy = new CNotes();
        if(this.clrMap){
            copy.setClMapOverride(this.clrMap.createDuplicate());
        }

        if(typeof this.cSld.name === "string" && this.cSld.name.length > 0)
        {
            copy.setCSldName(this.cSld.name);
        }
        if(this.cSld.Bg)
        {
            copy.changeBackground(this.cSld.Bg.createFullCopy());
        }
        for(var i = 0; i < this.cSld.spTree.length; ++i)
        {
            var _copy = this.cSld.spTree[i].copy(oPr);
            oIdMap[this.cSld.spTree[i].Id] = _copy.Id;
            copy.addToSpTreeToPos(copy.cSld.spTree.length, _copy);
            copy.cSld.spTree[copy.cSld.spTree.length - 1].setParent2(copy);
        }
        if(AscFormat.isRealBool(this.showMasterPhAnim))
        {
            copy.setShowMasterPhAnim(this.showMasterPhAnim);
        }
        if(AscFormat.isRealBool(this.showMasterSp))
        {
            copy.setShowMasterSp(this.showMasterSp);
        }
        copy.setNotesMaster(this.Master);

        return copy;
    };

		//todo think about it
    CNotes.prototype.showDrawingObjects = function(){
        var oPresentation = editor.WordControl.m_oLogicDocument;
				if (oPresentation.IsNotesPageMode()) {
					AscCommonSlide.SlideBase.prototype.showDrawingObjects.call(this);
				}else if(this.slide){
            if(oPresentation.CurPage === this.slide.num){
                editor.WordControl.m_oDrawingDocument.Notes_OnRecalculate(this.slide.num, this.slide.NotesWidth, this.slide.getNotesHeight());
            }
        }
    };
    CNotes.prototype.isViewerMode = function()
    {
        return editor.WordControl.m_oLogicDocument.IsViewMode();
    };
    CNotes.prototype.convertPixToMM = function(pix)
    {
        return editor.WordControl.m_oDrawingDocument.GetMMPerDot(pix);
    };

	CNotes.prototype.drawViewPrMarks = function(oGraphics) {
		if(oGraphics.isSupportTextDraw && !oGraphics.isSupportTextDraw()) return;
		return AscCommonSlide.Slide.prototype.drawViewPrMarks.call(this, oGraphics);
	};
	CNotes.prototype.getDrawingsForController = function() {
		return this.cSld.spTree;
	};
	CNotes.prototype.getTheme = function() {
		return this.Master && this.Master.getTheme() || null;
	};
	CNotes.prototype.getColorMap = function() {
		if(this.Master)
		{
			if(this.Master.clrMap)
			{
				return this.Master.clrMap;
			}
		}
		return AscFormat.GetDefaultColorMap();
	};
	CNotes.prototype.recalculateBackground = function() {
		let RGBA = {R: 0, G: 0, B: 0, A: 255};

		const _master = this.Master;
		const _theme = _master.Theme;
		let ret;
		if (this && this.cSld.Bg) {
			ret = this.cSld.recalculateBackground(_theme, this, null, _master, RGBA);
		} else if (_master && _master.cSld.Bg) {
			ret = _master.cSld.recalculateBackground(_theme, this, null, _master, RGBA);
		} else {
			ret = this.getDefaultBackFill();
		}
		RGBA = ret.RGBA || RGBA;
		const backFill = ret.backFill;
		if (backFill != null)
			backFill.calculate(_theme, this, this, _master, RGBA);

		this.backgroundFill = backFill;
	};
	CNotes.prototype.getAllFonts = function(fonts) {
		var i;
		for(i = 0; i < this.cSld.spTree.length; ++i)
		{
			if(typeof  this.cSld.spTree[i].getAllFonts === "function")
				this.cSld.spTree[i].getAllFonts(fonts);
		}
	};
	CNotes.prototype.needMasterSpDraw = function() {
		return this.showMasterSp;
	};
	CNotes.prototype.createFontMap = function(oFontsMap, oCheckedMap, isNoPh) {
		if (oCheckedMap[this.Get_Id()]) {
			return;
		}
		var aSpTree = this.cSld.spTree;
		var nSp, nSpCount = aSpTree.length;
		for(nSp = 0; nSp < nSpCount; ++nSp) {
			aSpTree[nSp].createFontMap(oFontsMap);
		}
		if(this.needMasterSpDraw()) {
			this.Master.createFontMap(oFontsMap, oCheckedMap, true);
		}
		oCheckedMap[this.Get_Id()] = this;
	};
	CNotes.prototype.recalculate = function() {
		if(!this.Master) {
			return;
		}
		if(this.recalcInfo.recalculateSlide) {
			this.recalcInfo.recalculateSlide = false;
			this.recalculateSlide();
		}
		this.Master.recalculate();
		if(this.recalcInfo.recalculateBackground) {
			this.recalculateBackground();
			this.recalcInfo.recalculateBackground = false;
		}
		if(this.recalcInfo.recalculateSpTree) {
			for(let i = 0; i < this.cSld.spTree.length; ++i)
				this.cSld.spTree[i].recalculate();

			this.recalcInfo.recalculateSpTree = false;
		}
		//todo cachedImage
		this.cachedImage = null;
	};
	CNotes.prototype.recalculateSlide = function() {
		const presentation = this.presentation;
		if (presentation) {
			for (let i = 0; i < presentation.Slides.length; i += 1) {
				const slide = presentation.Slides[i];
				if (slide.notes === this) {
					this.slide = slide;
					slide.recalculate();
					return;
				}
			}
		}
	};
	CNotes.prototype.checkSlideColorScheme = function() {
		this.recalcInfo.recalculateSpTree = true;
		this.recalcInfo.recalculateBackground = true;
		this.cSld.forEachSp(function(oSp) {
			oSp.handleUpdateFill();
			oSp.handleUpdateLn();
		});
	};
	//todo
	CNotes.prototype.getParentObjects = function() {
		return {};
	};
	CNotes.prototype.Get_ColorMap = function() {
		if (this.clrMap) {
			return this.clrMap;
		} else if (this.Master && this.Master.clrMap) {
			return this.Master.clrMap;
		}
		return AscFormat.GetDefaultColorMap();
	};
	CNotes.prototype.IsUseInDocument = function() {
		if(this.slide){
			return this.slide.IsUseInDocument();
		}
		return false;
	};
	CNotes.prototype.draw = function(graphics, slide) {
		let i;
		this.drawBgMaster(graphics);
		for(i = 0; i < this.cSld.spTree.length; ++i) {
			let oSp = this.cSld.spTree[i];
			if(AscCommon.IsHiddenObj(oSp)) {
				continue;
			}
			oSp.draw(graphics);
		}

		if(slide) {
			this.drawViewPrMarks(graphics);
		}
	};
	CNotes.prototype.drawBgMaster = function(graphics) {
		DrawBackground(graphics, this.backgroundFill, this.Width, this.Height);
		if(this.needMasterSpDraw()) {
				this.Master.drawNoPlaceholdersShapesOnly(graphics, this);
		}
	};
	CNotes.prototype.shapeRemove = function(pos, count) {
		if(pos > -1 && pos < this.cSld.spTree.length){
			History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_NotesRemoveFromTree, pos, this.cSld.spTree.splice(pos, 1), false));
		}
	};
	CNotes.prototype.shapeAdd = function(pos, obj) {
		var _pos = Math.max(0, Math.min(pos, this.cSld.spTree.length));
		History.Add(new AscDFH.CChangesDrawingsContentPresentation(this, AscDFH.historyitem_NotesAddToSpTree, _pos, [obj], true));
		this.cSld.spTree.splice(_pos, 0, obj);
		obj.setParent2(this);
	};
	CNotes.prototype.setSlideSize = function(w, h) {
		this.Width = w;
		this.Height = h;
	};
	CNotes.prototype.changeBackground = function(bg) {
		History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_NotesSetBg, this.cSld.Bg , bg));
		this.cSld.Bg = bg;
	};
	CNotes.prototype.isNote = function() {
		return true;
	};


    function CreateNotes(){
        var oN = new CNotes();
        var oSp = new AscFormat.CShape();
        oSp.setBDeleted(false);
        var oNvSpPr = new AscFormat.UniNvPr();
        var oCNvPr = oNvSpPr.cNvPr;
        oCNvPr.setId(2);
        oCNvPr.setName("Slide Image Placeholder 1");
        var oPh = new AscFormat.Ph();
        oPh.setType(AscFormat.phType_sldImg);
        oNvSpPr.nvPr.setPh(oPh);
        oSp.setNvSpPr(oNvSpPr);
        oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
        oSp.setLockValue(AscFormat.LOCKS_MASKS.noRot, true);
        oSp.setLockValue(AscFormat.LOCKS_MASKS.noChangeAspect, true);
        oSp.setSpPr(new AscFormat.CSpPr());
        oSp.spPr.setParent(oSp);
        oSp.setParent(oN);
        oN.addToSpTreeToPos(0, oSp);

        oSp = new AscFormat.CShape();
        oSp.setBDeleted(false);
        oNvSpPr = new AscFormat.UniNvPr();
        oCNvPr = oNvSpPr.cNvPr;
        oCNvPr.setId(3);
        oCNvPr.setName("Notes Placeholder 2");
        oPh = new AscFormat.Ph();
        oPh.setType(AscFormat.phType_body);
        oPh.setIdx(1 + "");
        oNvSpPr.nvPr.setPh(oPh);
        oSp.setNvSpPr(oNvSpPr);
        oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
        oSp.setSpPr(new AscFormat.CSpPr());
        oSp.spPr.setParent(oSp);
        oSp.createTextBody();
        var oBodyPr = new AscFormat.CBodyPr();
        oSp.txBody.setBodyPr(oBodyPr);
        var oTxLstStyle = new AscFormat.TextListStyle();
        oSp.txBody.setLstStyle(oTxLstStyle);
        oSp.setParent(oN);
        oN.addToSpTreeToPos(1, oSp);

        oSp = new AscFormat.CShape();
        oSp.setBDeleted(false);
        oNvSpPr = new AscFormat.UniNvPr();
        oCNvPr = oNvSpPr.cNvPr;
        oCNvPr.setId(4);
        oCNvPr.setName("Slide Number Placeholder 3");
        oPh = new AscFormat.Ph();
        oPh.setType(AscFormat.phType_sldNum);
        oPh.setSz(2);
        oPh.setIdx(10 + "");
        oNvSpPr.nvPr.setPh(oPh);
        oSp.setNvSpPr(oNvSpPr);
        oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
        oSp.setSpPr(new AscFormat.CSpPr());
        oSp.spPr.setParent(oSp);
        oSp.createTextBody();
        oBodyPr = new AscFormat.CBodyPr();
        oSp.txBody.setBodyPr(oBodyPr);
        oTxLstStyle = new AscFormat.TextListStyle();
        oSp.txBody.setLstStyle(oTxLstStyle);
        const oContent = oSp.getDocContent();
        if(oContent) {
            oContent.ClearContent(true);
            const oParagraph = oContent.Content[0];
            const oFld = new AscCommonWord.CPresentationField(oParagraph);
            oFld.SetGuid(AscCommon.CreateGUID());
            oFld.SetFieldType("slidenum");
            oParagraph.Internal_Content_Add(0, oFld);
        }
        oSp.setParent(oN);
        oN.addToSpTreeToPos(2, oSp);
        return oN;
    }

    window['AscCommonSlide'] = window['AscCommonSlide'] || {};
    window['AscCommonSlide'].CNotes = CNotes;
    window['AscCommonSlide'].GetNotesWidth = GetNotesWidth;
    window['AscCommonSlide'].CreateNotes = CreateNotes;
})();
