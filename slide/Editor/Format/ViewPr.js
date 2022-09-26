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


    function fReadSlideSize(oStream) {
        let oSlideSize = new AscCommonSlide.CSlideSize();
        let nStart = oStream.cur;
        let nEnd = nStart + oStream.GetULong() + 4;
        let nType = oStream.GetUChar();//start attributes
        oSlideSize.setCX(oStream.GetLong());
        oSlideSize.setCY(oStream.GetLong());
        nType = oStream.GetUChar();//end attributes
        oStream.Seek2(nEnd);
        return oSlideSize;
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
                //let oGridSpacing = fReadSlideSize(oStream);
                //this.setGridSpacing(oGridSpacing);

                oStream.SkipRecord();
                break;
            }
            case 1:
            case 2:
            case 3: {
                oStream.SkipRecord();
                break;
            }
            case 4: {
                let oSlideViewPr = new CCommonViewPr();
                oSlideViewPr.fromPPTY(pReader);
                this.setSlideViewPr(oSlideViewPr);
                break;
            }
            case 5:
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
    CViewPr.prototype.readChildXml = function (name, reader) {
        let oChild;
        switch(name) {
        }
    };
    CViewPr.prototype.toXml = function (writer, name) {
    };
    CViewPr.prototype.DEFAULT_GRID_SPACING = 360000;
    CViewPr.prototype.getGridSpacing = function () {
        let oSpacing = this.gridSpacing;
        if(oSpacing) {
            return oSpacing.cx || this.DEFAULT_GRID_SPACING;
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


    function CCommonViewPr() {
        CBFO.call(this);
        this.cSldViewPr = null;
        this.extLst = null;
    }
    IC(CCommonViewPr, CBFO, AscDFH.historyitem_type_CommonViewPr);
    CCommonViewPr.prototype.setCSldViewPr = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CommonViewPrCSldViewPr, this.cSldViewPr, oPr));
        this.cSldViewPr = oPr;
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
    CCommonViewPr.prototype.fromPPTY = function(pReader) {
        let nType = pReader.stream.GetUChar();
        if(nType === 0) {
            let oCSldViewPr = new CCSldViewPr();
            oCSldViewPr.fromPPTY(pReader);
            this.setCSldViewPr(oCSldViewPr);
        }
        else {
            pReader.stream.SkipRecord();
        }
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

    const GUIDE_POS_TO_EMU = 1587.5;

    function GdPosToMm(nVal) {
        return AscFormat.Emu_To_Mm(nVal * GUIDE_POS_TO_EMU);
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
    };
    CCSldViewPr.prototype.addGuide = function(oPr) {
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_CSldViewPrGuideLst, this.guideLst.length, [oPr], true));
        this.guideLst.push(oPr);
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
        for(let nGd = 0; nGd < aGds.length; ++nGd) {
            aGds[nGd].draw(oGraphics);
        }
    };
    CCSldViewPr.prototype.insertGuide = function(bHorizontal) {
        let oLastGuide = null;
        let oGuideToAdd = new CGuide();
        if(bHorizontal) {
            oGuideToAdd.setOrient(orient_horz);
        }
        for(let nGd = 0; nGd < this.guideLst.length; ++nGd) {
            let oGuide = this.guideLst[nGd];
            if(bHorizontal && oGuide.isHorizontal() || oGuide.isVertical()) {
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



    const orient_horz = 0;
    const orient_vert = 1;

    function CGuide() {
        CBFO.call(this);
        this.pos = null;
        this.orient = null;
    }
    IC(CGuide, CBFO, AscDFH.historyitem_type_Unknown);
    CGuide.prototype.setPos = function(oPr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_Unknown_Unknown, this.pos, oPr));
        this.pos = oPr;
    };
    CGuide.prototype.setOrient = function(oPr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_Unknown_Unknown, this.orient, oPr));
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
    CGuide.prototype.draw = function(oGraphics) {
        let oPresentation = editor.WordControl.m_oLogicDocument;
        let dPos = GdPosToMm(this.pos);
        if(this.isHorizontal()) {
            oGraphics.DrawEmptyTableLine(0, dPos, oPresentation.GetWidthMM(), dPos);
        }
        else {
            oGraphics.DrawEmptyTableLine(dPos, 0, dPos, oPresentation.GetHeightMM());
        }
    };
    CGuide.prototype.isHorizontal = function() {
        return (this.orient === orient_horz);
    };
    CGuide.prototype.isVertical = function() {
        return !this.isVertical();
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
    };
    CCViewPr.prototype.setScale = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CViewPrScale, this.scale, oPr));
        this.scale = oPr;
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
    CCViewPr.prototype.readChildXml = function (name, reader) {
        let oChild;
        switch(name) {

        }
    };
    CCViewPr.prototype.toXml = function (writer, name) {
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
    };
    CScale.prototype.setSy = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ViewPrScaleSy, this.sy, oPr));
        this.sy = oPr;
    };
    CScale.prototype.fromPPTY = function(pReader) {
        let oStream = pReader.stream;
        oStream.GetUChar();
        this.setSx(fReadSlideSize(oStream));
        oStream.GetUChar();
        this.setSy(fReadSlideSize(oStream));
    };
    CScale.prototype.readChildXml = function (name, reader) {
        let oChild;
        switch(name) {
        }
    };
    CScale.prototype.toXml = function (writer, name) {
    };

    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].CViewPr = CViewPr;
}) (window);
