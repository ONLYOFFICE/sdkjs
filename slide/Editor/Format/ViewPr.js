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

(function (window, undefined) {


    let CBaseObject = AscFormat.CBaseObject;
    let oHistory = AscCommon.History;
    let CChangeBool = AscDFH.CChangesDrawingsBool;
    let CChangeLong = AscDFH.CChangesDrawingsLong;
    let CChangeDouble = AscDFH.CChangesDrawingsDouble;
    let CChangeString = AscDFH.CChangesDrawingsString;
    let CChangeObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
    let CChangeObject = AscDFH.CChangesDrawingsObject;
    let CChangeContent = AscDFH.CChangesDrawingsContent;
    let CChangeDouble2 = AscDFH.CChangesDrawingsDouble2;

    let IC = AscFormat.InitClass;
    let CBFO = AscFormat.CBaseFormatObject;



    AscDFH.changesFactory[AscDFH.historyitem_ViewPrGridSpacing] = CChangeObject;
    AscDFH.changesFactory[AscDFH.historyitem_ViewPrSlideViewerPr] = CChangeObject;
    AscDFH.changesFactory[AscDFH.historyitem_ViewPrLastView] = CChangeLong;
    AscDFH.changesFactory[AscDFH.historyitem_ViewPrShowComments] = CChangeBool;
    AscDFH.changesFactory[AscDFH.historyitem_CommonViewPrCSldViewPr] = CChangeObject;
    AscDFH.changesFactory[AscDFH.historyitem_CSldViewPrCViewPr] = CChangeObject;
    AscDFH.changesFactory[AscDFH.historyitem_CSldViewPrGuideLst] = CChangeContent;
    AscDFH.changesFactory[AscDFH.historyitem_CSldViewPrShowGuides] = CChangeBool;
    AscDFH.changesFactory[AscDFH.historyitem_CSldViewPrSnapToGrid] = CChangeBool;
    AscDFH.changesFactory[AscDFH.historyitem_CSldViewPrSnapToObjects] = CChangeBool;
    AscDFH.changesFactory[AscDFH.historyitem_ViewPrGuidePos] = CChangeLong;
    AscDFH.changesFactory[AscDFH.historyitem_ViewPrGuideOrient] = CChangeLong;
    AscDFH.changesFactory[AscDFH.historyitem_CViewPrOrigin] = CChangeObject;
    AscDFH.changesFactory[AscDFH.historyitem_CViewPrScale] = CChangeObject;
    AscDFH.changesFactory[AscDFH.historyitem_ViewPrScaleSx] = CChangeObject;
    AscDFH.changesFactory[AscDFH.historyitem_ViewPrScaleSy] = CChangeObject;


    let drawingsChangesMap = window['AscDFH'].drawingsChangesMap;

    drawingsChangesMap[AscDFH.historyitem_ViewPrGridSpacing] = function(oClass, value){oClass.gridSpacing = value;};
    drawingsChangesMap[AscDFH.historyitem_ViewPrSlideViewerPr] = function(oClass, value){oClass.slideViewPr = value;};
    drawingsChangesMap[AscDFH.historyitem_ViewPrLastView] = function(oClass, value){oClass.lastView = value;};
    drawingsChangesMap[AscDFH.historyitem_ViewPrShowComments] = function(oClass, value){oClass.showComments = value;};
    drawingsChangesMap[AscDFH.historyitem_CommonViewPrCSldViewPr] = function(oClass, value){oClass.cSldViewPr = value;};
    drawingsChangesMap[AscDFH.historyitem_CSldViewPrCViewPr] = function(oClass, value){oClass.cViewPr = value;};
    drawingsChangesMap[AscDFH.historyitem_CSldViewPrShowGuides] = function(oClass, value){oClass.showGuides = value;};
    drawingsChangesMap[AscDFH.historyitem_CSldViewPrSnapToGrid] = function(oClass, value){oClass.snapToGrid = value;};
    drawingsChangesMap[AscDFH.historyitem_CSldViewPrSnapToObjects] = function(oClass, value){oClass.snapToObjects = value;};
    drawingsChangesMap[AscDFH.historyitem_ViewPrGuidePos] = function(oClass, value){oClass.pos = value;};
    drawingsChangesMap[AscDFH.historyitem_ViewPrGuideOrient] = function(oClass, value){oClass.orient = value;};
    drawingsChangesMap[AscDFH.historyitem_CViewPrOrigin] = function(oClass, value){oClass.origin = value;};
    drawingsChangesMap[AscDFH.historyitem_CViewPrScale] = function(oClass, value){oClass.scale = value;};
    drawingsChangesMap[AscDFH.historyitem_ViewPrScaleSx] = function(oClass, value){oClass.sx = value;};
    drawingsChangesMap[AscDFH.historyitem_ViewPrScaleSy] = function(oClass, value){oClass.sy = value;};

    AscDFH.drawingContentChanges[AscDFH.historyitem_CSldViewPrGuideLst] = function(oClass){return oClass.guideLst;};



    function fReadSlideSize(oStream) {
        let oSlideSize = new AscCommonSlide.CSlideSize();
        let nStart = oStream.cur;
        let nEnd = nStart + oStream.GetULong() + 4;
        let nType = oStream.GetUChar();//start attributes
        oStream.GetUChar();
        oSlideSize.setCX(oStream.GetLong());
        oStream.GetUChar();
        oSlideSize.setCY(oStream.GetLong());
        nType = oStream.GetUChar();//end attributes
        oStream.Seek2(nEnd);
        return oSlideSize;
    }
    function fWriteSlideSize(oSize, pWriter, nType) {
        if(!oSize) {
            return;
        }
        pWriter.StartRecord(nType);
        pWriter.WriteUChar(AscCommon.g_nodeAttributeStart);
        pWriter._WriteInt2(0, oSize.cx);
        pWriter._WriteInt2(1, oSize.cy);
        pWriter.WriteUChar(AscCommon.g_nodeAttributeEnd);
        pWriter.EndRecord();
    }

    const lastView_handoutView = 0;
    const lastView_notesMasterView = 1;
    const lastView_notesView = 2;
    const lastView_outlineView = 3;
    const lastView_sldMasterView = 4;
    const lastView_sldSorterView = 5;
    const lastView_sldThumbnailView = 6;
    const lastView_sldView = 7;
    function CViewPr() {
        CBFO.call(this);
        this.gridSpacing = null;
        this.slideViewPr = null;

        this.normalViewPr = null;
        this.notesTextViewPr = null;
        this.notesViewPr = null;
        this.outlineViewPr = null;
        this.sorterViewPr = null;

        this.lastView = null;
        this.showComments = null;
    }
    IC(CViewPr, CBFO, AscDFH.historyitem_type_ViewPr);
    CViewPr.prototype.setGridSpacing = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ViewPrGridSpacing, this.gridSpacing, oPr));
        this.gridSpacing = oPr;
        this.setParentToChild(oPr);
    };
    CViewPr.prototype.setGridSpacingVal = function(nVal) {
        let oSpacing = new AscCommonSlide.CSlideSize();
        oSpacing.setCX(nVal);
        oSpacing.setCY(nVal);
        this.setGridSpacing(oSpacing);
    };
    CViewPr.prototype.setSlideViewPr = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ViewPrSlideViewerPr, this.slideViewPr, oPr));
        this.slideViewPr = oPr;
        this.setParentToChild(oPr);
    };
    CViewPr.prototype.setLastView = function(oPr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ViewPrLastView, this.lastView, oPr));
        this.lastView = oPr;
    };
    CViewPr.prototype.setShowComments = function(oPr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_ViewPrShowComments, this.showComments, oPr));
        this.showComments = oPr;
    };
    CViewPr.prototype.checkSlideViewPr = function() {
        if(!this.slideViewPr) {
            this.setSlideViewPr(new CCommonViewPr());
        }
        else {
            this.setSlideViewPr(this.slideViewPr.createDuplicate())
        }
        return this.slideViewPr;
    };
    CViewPr.prototype.readAttribute = function(nType, pReader) {
        let oStream = pReader.stream;
        switch (nType) {
            case 0: {
                this.setLastView(oStream.GetUChar());
                break;
            }
            case 1: {
                this.setShowComments(oStream.GetBool());
                break;
            }
        }
    };
    CViewPr.prototype.readChild = function(nType, pReader) {
        let oStream = pReader.stream;
        switch (nType) {
            case 0: {
                let oGridSpacing = fReadSlideSize(oStream);
                this.setGridSpacing(oGridSpacing);
                //oStream.SkipRecord();
                break;
            }
            case 1:
            case 2:
            case 3: {
                oStream.SkipRecord();
                break;
            }
            case 4: {
                oStream.SkipRecord();
                break;
            }
            case 5: {
                let oSlideViewPr = new CCommonViewPr();
                oSlideViewPr.fromPPTY(pReader);
                this.setSlideViewPr(oSlideViewPr);
                break;
            }
            case 6: {
                oStream.SkipRecord();
                break;
            }
            default: {
                oStream.SkipRecord();
                break;
            }
        }
    };


    CViewPr.prototype.privateWriteAttributes = function(pWriter) {
    };
    CViewPr.prototype.writeChildren = function(pWriter) {
        fWriteSlideSize(this.gridSpacing, pWriter, 0);
        this.writeRecord2(pWriter, 5, this.slideViewPr);
    };
    CViewPr.prototype.readChildXml = function (name, reader) {
        let oChild;
        switch(name) {
        }
    };
    CViewPr.prototype.toXml = function (writer, name) {
    };
    CViewPr.prototype.DEFAULT_GRID_SPACING = 360000;
    CViewPr.prototype.MAX_GRID_SPACING = 2*914400;//2 inches
    CViewPr.prototype.getGridSpacing = function () {
        let oSpacing = this.gridSpacing;
        if(oSpacing) {
            if(AscFormat.isRealNumber(oSpacing.cx) && oSpacing.cx > 0) {
                let nGridSpacing = oSpacing.cx;
                while(nGridSpacing > this.MAX_GRID_SPACING) {
                    nGridSpacing /= 1000;
                }
                return nGridSpacing
            }
            return this.DEFAULT_GRID_SPACING;
        }
        return  this.DEFAULT_GRID_SPACING;
    };
    CViewPr.prototype.isSnapToGrid = function () {
        if(this.slideViewPr) {
            return this.slideViewPr.isSnapToGrid();
        }
        return false;
    };
    CViewPr.prototype.setSnapToGrid = function(bVal) {
        this.checkSlideViewPr().checkCSldViewPr().setSnapToGrid(bVal);
    };
    CViewPr.prototype.Refresh_RecalcData = function(Data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(Data);
        }
    };
    CViewPr.prototype.drawGuides = function(oGraphics) {
        if(this.slideViewPr) {
            this.slideViewPr.drawGuides(oGraphics);
        }
    };
    CViewPr.prototype.addHorizontalGuide = function () {
        this.checkSlideViewPr().checkCSldViewPr().addHorizontalGuide();
    };
    CViewPr.prototype.addVerticalGuide = function () {
        this.checkSlideViewPr().checkCSldViewPr().addVerticalGuide();
    };
    CViewPr.prototype.fillObject = function (oCopy, oIdMap) {
        oCopy.setLastView(this.lastView);
        oCopy.setShowComments(this.showComments);
        if(this.slideViewPr) {
            oCopy.setSlideViewPr(this.slideViewPr.createDuplicate());
        }
    };
    CViewPr.prototype.getHorGuidesPos = function() {
        if(this.slideViewPr) {
            return this.slideViewPr.getHorGuidesPos();
        }
        return [];

    };
    CViewPr.prototype.getVertGuidesPos = function() {
        if(this.slideViewPr) {
            return this.slideViewPr.getVertGuidesPos();
        }
        return [];
    };
    CViewPr.prototype.canClearGuides = function() {
        return this.getHorGuidesPos().length > 0 || this.getVertGuidesPos().length > 0;
    };
    CViewPr.prototype.clearGuides = function() {
        if(!this.canClearGuides()) {
            return;
        }
        if(this.slideViewPr) {
            return this.slideViewPr.clearGuides();
        }
    };
    CViewPr.prototype.removeGuideById = function(sId) {
        if(this.slideViewPr) {
            return this.slideViewPr.removeGuideById(sId);
        }
    };
    CViewPr.prototype.hitInGuide = function(x, y) {
        if(this.slideViewPr) {
            return this.slideViewPr.hitInGuide(x, y);
        }
        return null;
    };
    CViewPr.prototype.scaleGuides = function(dCW, dCH) {
        if(this.slideViewPr) {
            this.slideViewPr.scaleGuides(dCW, dCH);
        }
    };
    CViewPr.prototype.Refresh_RecalcData = function(Data) {
        this.Refresh_RecalcData2(Data);
    };
    CViewPr.prototype.Refresh_RecalcData2 = function(Data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(Data);
        }
    };


    function CCommonViewPr() {
        CBFO.call(this);
        this.cSldViewPr = null;
        this.extLst = null;
    }
    IC(CCommonViewPr, CBFO, AscDFH.historyitem_type_CommonViewPr);
    CCommonViewPr.prototype.setCSldViewPr = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CommonViewPrCSldViewPr, this.cSldViewPr, oPr));
        this.cSldViewPr = oPr;
        this.setParentToChild(oPr);
    };
    CCommonViewPr.prototype.checkCSldViewPr = function() {
        if(!this.cSldViewPr) {
            this.setCSldViewPr(new CCSldViewPr());
        }
        return this.cSldViewPr;
    };
    CCommonViewPr.prototype.isSnapToGrid = function() {
        if(this.cSldViewPr) {
            return this.cSldViewPr.isSnapToGrid();
        }
        return false;
    };
    CCommonViewPr.prototype.readAttribute = function(nType, pReader) {
    };
    CCommonViewPr.prototype.readAttributes = function(pReader) {
    };
    CCommonViewPr.prototype.readChild = function(nType, pReader) {
        if(nType === 0) {
            let oCSldViewPr = new CCSldViewPr();
            oCSldViewPr.fromPPTY(pReader);
            this.setCSldViewPr(oCSldViewPr);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CCommonViewPr.prototype.writeAttributes = function (pWriter) {
    };
    CCommonViewPr.prototype.privateWriteAttributes = function(pWriter) {
    };
    CCommonViewPr.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.cSldViewPr);
    };
    CCommonViewPr.prototype.readChildXml = function (name, reader) {
    };
    CCommonViewPr.prototype.toXml = function (writer, name) {
    };
    CCommonViewPr.prototype.drawGuides = function (oGraphics) {
        if(this.cSldViewPr) {
            this.cSldViewPr.drawGuides(oGraphics);
        }
    };
    CCommonViewPr.prototype.fillObject = function (oCopy, oIdMap) {
        if(this.cSldViewPr) {
            oCopy.setCSldViewPr(this.cSldViewPr.createDuplicate());
        }
    };
    CCommonViewPr.prototype.getHorGuidesPos = function() {
        if(this.cSldViewPr) {
            return this.cSldViewPr.getHorGuidesPos();
        }
        return [];

    };
    CCommonViewPr.prototype.getVertGuidesPos = function() {
        if(this.cSldViewPr) {
            return this.cSldViewPr.getVertGuidesPos();
        }
        return [];
    };
    CCommonViewPr.prototype.clearGuides = function() {
        if(this.cSldViewPr) {
            return this.cSldViewPr.clearGuides();
        }
    };
    CCommonViewPr.prototype.removeGuideById = function(sId) {
        if(this.cSldViewPr) {
            return this.cSldViewPr.removeGuideById(sId);
        }
    };
    CCommonViewPr.prototype.hitInGuide = function(x, y) {
        if(this.cSldViewPr) {
            return this.cSldViewPr.hitInGuide(x, y);
        }
        return null;
    };
    CCommonViewPr.prototype.scaleGuides = function(dCW, dCH) {
        if(this.cSldViewPr) {
            this.cSldViewPr.scaleGuides(dCW, dCH);
        }
    };
    CCommonViewPr.prototype.Refresh_RecalcData = function(Data) {
        this.Refresh_RecalcData2(Data);
    };
    CCommonViewPr.prototype.Refresh_RecalcData2 = function(Data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(Data);
        }
    };


    const GUIDE_POS_TO_EMU = 1587.5;

    function GdPosToMm(nVal) {
        return nVal * GUIDE_POS_TO_EMU / 36000;
    }
    function MmToGdPos(dVal) {
        return (dVal / GdPosToMm(1) + 0.5) >> 0;
    }
    function EmuToGdPos(nVal) {
        return (nVal / GUIDE_POS_TO_EMU + 0.5) >> 0;
    }

    function CCSldViewPr() {
        CBFO.call(this);

        this.cViewPr = null;
        this.guideLst = [];

        this.showGuides = null;
        this.snapToGrid = null;
        this.snapToObjects = null;
    }
    IC(CCSldViewPr, CBFO, AscDFH.historyitem_type_CSldViewPr);
    CCSldViewPr.prototype.setCViewPr = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CSldViewPrCViewPr, this.cViewPr, oPr));
        this.cViewPr = oPr;
        this.setParentToChild(oPr);
    };
    CCSldViewPr.prototype.addGuide = function(oPr) {
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_CSldViewPrGuideLst, this.guideLst.length, [oPr], true));
        this.guideLst.push(oPr);
        if(oPr) {
            oPr.setParent(this);
        }
    };
    CCSldViewPr.prototype.setShowGuides = function(oPr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CSldViewPrShowGuides, this.showGuides, oPr));
        this.showGuides = oPr;
    };
    CCSldViewPr.prototype.setSnapToGrid = function(oPr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CSldViewPrSnapToGrid, this.snapToGrid, oPr));
        this.snapToGrid = oPr;
    };
    CCSldViewPr.prototype.setSnapToObjects = function(oPr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CSldViewPrSnapToObjects, this.snapToObjects, oPr));
        this.snapToObjects = oPr;
    };
    CCSldViewPr.prototype.readAttribute = function(nType, pReader) {
        let oStream = pReader.stream;
        switch (nType) {
            case 0: {
                this.setShowGuides(oStream.GetBool());
                break;
            }
            case 1: {
                this.setSnapToGrid(oStream.GetBool());
                break;
            }
            case 2: {
                this.setSnapToObjects(oStream.GetBool());
                break;
            }
        }
    };
    CCSldViewPr.prototype.readChild = function(nType, pReader) {
        let oStream = pReader.stream;
        switch(nType) {
            case 0: {
                let oViewPr = new CCViewPr();
                oViewPr.fromPPTY(pReader);
                this.setCViewPr(oViewPr);
                break;
            }
            case 1: {
                let nStart = oStream.cur;
                let nEnd = nStart + oStream.GetULong() + 4;
                let nGuideCount = oStream.GetULong();
                for(let nGuide = 0; nGuide < nGuideCount; ++nGuide) {
                    let nChildType = oStream.GetUChar();
                    if(nChildType === 2) {
                        let oGuide = new CGuide();
                        oGuide.fromPPTY(pReader);
                        this.addGuide(oGuide);
                    }
                    else {
                        oStream.SkipRecord();
                    }
                }
                oStream.Seek2(nEnd);
                break;
            }
            default: {
                oStream.SkipRecord();
                break;
            }
        }
    };
    CCSldViewPr.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteBool2(0, this.showGuides);
        pWriter._WriteBool2(1, this.snapToGrid);
        pWriter._WriteBool2(2, this.snapToObjects);
    };
    CCSldViewPr.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.cViewPr);

        pWriter.StartRecord(1);
        pWriter.WriteULong(this.guideLst.length);
        for (let nGd = 0; nGd < this.guideLst.length; ++nGd) {
            this.writeRecord2(pWriter, 2, this.guideLst[nGd]);
        }

        pWriter.EndRecord();
    };
    CCSldViewPr.prototype.readChildXml = function (name, reader) {
        let oChild;
        switch(name) {
        }
    };
    CCSldViewPr.prototype.toXml = function (writer, name) {
    };
    CCSldViewPr.prototype.isSnapToGrid = function () {
        return this.snapToGrid === true;
    };
    CCSldViewPr.prototype.drawGuides = function (oGraphics) {
        let aGds = this.guideLst;
	    oGraphics.SaveGrState();
	    oGraphics.SetIntegerGrid(true);
	    oGraphics.transform3(new AscCommon.CMatrix());
        for(let nGd = 0; nGd < aGds.length; ++nGd) {
            aGds[nGd].draw(oGraphics);
        }
	    oGraphics.RestoreGrState();
    };
    CCSldViewPr.prototype.insertGuide = function(bHorizontal) {
        let oLastGuide = null;
        let oGuideToAdd = new CGuide();
        if(bHorizontal) {
            oGuideToAdd.setOrient(orient_horz);
        }
        for(let nGd = 0; nGd < this.guideLst.length; ++nGd) {
            let oGuide = this.guideLst[nGd];
            if(bHorizontal && oGuide.isHorizontal() || !bHorizontal && oGuide.isVertical()) {
                if(!oLastGuide) {
                    oLastGuide = oGuide;
                }
                else {
                    if(oLastGuide.pos < oGuide.pos) {
                        oLastGuide = oGuide;
                    }
                }
            }
        }
        let oPresentation = editor.WordControl.m_oLogicDocument;
        let nWidth = EmuToGdPos(oPresentation.GetWidthEMU());
        let nHeight = EmuToGdPos(oPresentation.GetHeightEMU());
        if(!oLastGuide) {
            if(bHorizontal) {
                oGuideToAdd.setPos(nHeight / 2 >> 0);
            }
            else {
                oGuideToAdd.setPos(nWidth / 2 >> 0);
            }
        }
        else {
            if(bHorizontal) {
                oGuideToAdd.setPos(Math.min(nHeight, oLastGuide.pos + 100));
            }
            else {
                oGuideToAdd.setPos(Math.min(nWidth, oLastGuide.pos + 100));
            }
        }
        this.addGuide(oGuideToAdd);
    };
    CCSldViewPr.prototype.addHorizontalGuide = function () {
        this.insertGuide(true);
    };
    CCSldViewPr.prototype.addVerticalGuide = function () {
        this.insertGuide(false);
    };
    CCSldViewPr.prototype.fillObject = function (oCopy, oIdMap) {
        oCopy.setShowGuides(this.showGuides);
        oCopy.setSnapToGrid(this.snapToGrid);
        oCopy.setSnapToObjects(this.snapToObjects);

        if(this.cViewPr) {
            oCopy.setCViewPr(this.cViewPr.createDuplicate());
        }
        for(let nGd = 0; nGd < this.guideLst.length; ++nGd) {
            oCopy.addGuide(this.guideLst[nGd].createDuplicate());
        }
    };
    CCSldViewPr.prototype.getGuidesPos = function(bHor) {
        let aRet = [];
        for(let nGd = 0; nGd < this.guideLst.length; ++nGd) {
            let oGd = this.guideLst[nGd];
            if(bHor && oGd.isHorizontal() || !bHor && oGd.isVertical()) {
                aRet.push(GdPosToMm(oGd.pos))
            }
        }
        return aRet;

    };
    CCSldViewPr.prototype.getHorGuidesPos = function() {
        return this.getGuidesPos(true);

    };
    CCSldViewPr.prototype.getVertGuidesPos = function() {
        return this.getGuidesPos(false);
    };
    CCSldViewPr.prototype.clearGuides = function() {
        let nLength = this.guideLst.length;
        for(let nGd = nLength - 1; nGd > -1; nGd --) {
            this.removeGuide(nGd);
        }
    };
    CCSldViewPr.prototype.removeGuide = function(nIdx) {
        if(this.guideLst[nIdx]) {
            oHistory.Add(new CChangeContent(this, AscDFH.historyitem_CSldViewPrGuideLst, nIdx, [this.guideLst[nIdx]], false))
            this.guideLst.splice(nIdx, 1);
        }
    };
    CCSldViewPr.prototype.removeGuideById = function(sId) {
        let nLength = this.guideLst.length;
        for(let nGd = nLength - 1; nGd > -1; nGd --) {
            let oGd = this.guideLst[nGd];
            if(oGd.Id === sId) {
                this.removeGuide(nGd);
                return;
            }
        }
    };
    CCSldViewPr.prototype.hitInGuide = function(x, y) {
        let nLength = this.guideLst.length;
        for(let nGd = nLength - 1; nGd > -1; nGd --) {
            let oGd = this.guideLst[nGd];
            if(oGd.hit(x, y)) {
                return oGd.Id;
            }
        }
        return null;
    };

    CCSldViewPr.prototype.scaleGuides = function(dCW, dCH) {
        let nLength = this.guideLst.length;
        for(let nGd = nLength - 1; nGd > -1; nGd --) {
            let oGd = this.guideLst[nGd];
            oGd.scale(dCW, dCH);
        }
    };
    CCSldViewPr.prototype.Refresh_RecalcData = function(Data) {
        this.Refresh_RecalcData2(Data);
    };
    CCSldViewPr.prototype.Refresh_RecalcData2 = function(Data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(Data);
        }
    };



    const orient_horz = 0;
    const orient_vert = 1;

    function CGuide() {
        CBFO.call(this);
        this.pos = null;
        this.orient = null;
    }
    IC(CGuide, CBFO, AscDFH.historyitem_type_ViewPrGuide);
    CGuide.prototype.setPos = function(oPr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ViewPrGuidePos, this.pos, oPr));
        this.pos = oPr;
    };
    CGuide.prototype.setOrient = function(oPr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ViewPrGuideOrient, this.orient, oPr));
        this.orient = oPr;
    };
    CGuide.prototype.readAttribute = function(nType, pReader) {
        let oStream = pReader.stream;
        switch (nType) {
            case 0: {
                this.setPos(oStream.GetLong());
                break;
            }
            case 1: {
                this.setOrient(oStream.GetUChar());
                break;
            }
        }
    };
    CGuide.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteUInt2(0, this.pos);
        pWriter._WriteLimit2(1, this.orient);
    };
    CGuide.prototype.draw = function(oGraphics) {
        let oPresentation = editor.WordControl.m_oLogicDocument;
        let dPos = GdPosToMm(this.pos);
        let bOldVal = editor.isShowTableEmptyLineAttack;
        editor.isShowTableEmptyLineAttack = true;
        if(this.isHorizontal()) {
            oGraphics.DrawEmptyTableLine(0, dPos, oPresentation.GetWidthMM(), dPos);
        }
        else {
            oGraphics.DrawEmptyTableLine(dPos, 0, dPos, oPresentation.GetHeightMM());
        }
        editor.isShowTableEmptyLineAttack = bOldVal;
    };
    CGuide.prototype.isHorizontal = function() {
        return (this.orient === orient_horz);
    };
    CGuide.prototype.isVertical = function() {
        return !this.isHorizontal();
    };
    CGuide.prototype.fillObject = function (oCopy, oIdMap) {
        oCopy.setPos(this.pos);
        oCopy.setOrient(this.orient);
    };

    CGuide.prototype.hit = function (x, y) {
        let dPos = GdPosToMm(this.pos);
        let dCheckPos;
        if(this.isHorizontal()) {
            dCheckPos = y;
        }
        else {
            dCheckPos = x;
        }
        if(Math.abs(dCheckPos - dPos) < AscCommon.TRACK_CIRCLE_RADIUS) {
            return true;
        }
        return false;
    };

    CGuide.prototype.scale = function(dCW, dCH) {
        if(this.isHorizontal()) {
            this.setPos(this.pos * dCH + 0.5 >> 0);
        }
        else {
            this.setPos(this.pos * dCW + 0.5 >> 0);
        }
    };
    CGuide.prototype.Refresh_RecalcData = function (Data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(Data);
        }
    };

    function CCViewPr() {
        CBFO.call(this);
        this.origin = null;
        this.scale = null;

    }
    IC(CCViewPr, CBFO, AscDFH.historyitem_type_CViewPr);
    CCViewPr.prototype.setOrigin = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CViewPrOrigin, this.origin, oPr));
        this.origin = oPr;
        this.setParentToChild(oPr);
    };
    CCViewPr.prototype.setScale = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CViewPrScale, this.scale, oPr));
        this.scale = oPr;
        this.setParentToChild(oPr);
    };
    CCViewPr.prototype.readAttribute = function(nType, pReader) {
        if(nType === 0) {
            this.setScale(pReader.stream.GetBool());
        }
    };
    CCViewPr.prototype.readChild = function(nType, pReader) {
        let oStream = pReader.stream;
        switch (nType) {
            case 0: {
                let oOrigin = fReadSlideSize(oStream);
                this.setOrigin(oOrigin);
                break;
            }
            case 1: {
                let oScale = new CScale();
                oScale.fromPPTY(pReader);
                this.setScale(oScale);
                break;
            }
        }
    };
    CCViewPr.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteBool2(0, this.scale);
    };
    CCViewPr.prototype.writeChildren = function(pWriter) {
        fWriteSlideSize(this.origin, pWriter, 0);
        this.writeRecord2(pWriter, 1, this.scale);
    };
    CCViewPr.prototype.readChildXml = function (name, reader) {
        let oChild;
        switch(name) {

        }
    };
    CCViewPr.prototype.toXml = function (writer, name) {
    };
    CCViewPr.prototype.fillObject = function (oCopy, oIdMap) {
        oCopy.setScale(this.scale);
        if(this.origin) {
            oCopy.setOrigin(this.origin.createDuplicate());
        }
    };
    CCViewPr.prototype.Refresh_RecalcData = function(Data) {
        this.Refresh_RecalcData2(Data);
    };
    CCViewPr.prototype.Refresh_RecalcData2 = function(Data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(Data);
        }
    };

    function CScale() {
        CBFO.call(this);
        this.sx = null;
        this.sy = null;
    }
    IC(CScale, CBFO, AscDFH.historyitem_type_ViewPrScale);
    CScale.prototype.setSx = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ViewPrScaleSx, this.sx, oPr));
        this.sx = oPr;
        this.setParentToChild(oPr);
    };
    CScale.prototype.setSy = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ViewPrScaleSy, this.sy, oPr));
        this.sy = oPr;
        this.setParentToChild(oPr);
    };
    CScale.prototype.fromPPTY = function(pReader) {

        let oStream = pReader.stream;

        let nStart = oStream.cur;
        let nEnd = nStart + oStream.GetULong() + 4;
        let val;
        val = oStream.GetUChar();
        this.setSx(fReadSlideSize(oStream));
        val = oStream.GetUChar();
        this.setSy(fReadSlideSize(oStream));

        oStream.Seek2(nEnd);
    };
    CScale.prototype.toPPTY = function (pWriter) {
        fWriteSlideSize(this.sx, pWriter, 0);
        fWriteSlideSize(this.sy, pWriter, 1);
    };
    CScale.prototype.readChildXml = function (name, reader) {
        let oChild;
        switch(name) {
        }
    };
    CScale.prototype.toXml = function (writer, name) {
    };
    CScale.prototype.fillObject = function (oCopy, oIdMap) {
        oCopy.setSx(this.sx);
        oCopy.setSy(this.sy);
    };
    CScale.prototype.Refresh_RecalcData = function(Data) {
        this.Refresh_RecalcData2(Data);
    };
    CScale.prototype.Refresh_RecalcData2 = function(Data) {
        if(this.parent) {
            this.parent.Refresh_RecalcData2(Data);
        }
    };


    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].CViewPr = CViewPr;
    window['AscFormat'].CCommonViewPr = CCommonViewPr;
    window['AscFormat'].CCSldViewPr = CCSldViewPr;
    window['AscFormat'].CCViewPr = CCViewPr;
    window['AscFormat'].CViewPrScale = CScale;
    window['AscFormat'].CViewPrGuide = CGuide;
    window['AscFormat'].MmToGdPos = MmToGdPos;
    window['AscFormat'].GdPosToMm = GdPosToMm;
}) (window);
