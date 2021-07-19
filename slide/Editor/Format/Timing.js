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

(function(window, undefined) {
    var CBaseObject = AscFormat.CBaseObject;
    var oHistory = AscCommon.History;
    var CChangeBool = AscDFH.CChangesDrawingsBool;
    var CChangeLong = AscDFH.CChangesDrawingsLong;
    var CChangeDouble = AscDFH.CChangesDrawingsDouble;
    var CChangeString = AscDFH.CChangesDrawingsString;
    var CChangeObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
    var CChangeObject = AscDFH.CChangesDrawingsObject;
    var CChangeContent = AscDFH.CChangesDrawingsContent;
    var CChangeDouble2 = AscDFH.CChangesDrawingsDouble2;

    var drawingsChangesMap = AscDFH.drawingsChangesMap;
    var drawingContentChanges = AscDFH.drawingContentChanges;
    var changesFactory = AscDFH.changesFactory;
    var drawingConstructorsMap = window['AscDFH'].drawingsConstructorsMap;



    var InitClass = AscFormat.InitClass;
    var CBaseFormatObject = AscFormat.CBaseFormatObject;


    function CEmptyObject() {
        CBaseFormatObject.call(this);
    }
    InitClass(CEmptyObject, CBaseFormatObject, AscDFH.historyitem_type_EmptyObject);

    changesFactory[AscDFH.historyitem_TimingBldLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_TimingTnLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TimingBldLst] = function(oClass, value) {oClass.bldLst = value;};
    drawingsChangesMap[AscDFH.historyitem_TimingTnLst] = function(oClass, value) {oClass.tnLst = value;};
    function CTiming() {
        CBaseFormatObject.call(this);
        this.bldLst = null;
        this.tnLst  = null;
    }
    InitClass(CTiming, CBaseFormatObject, AscDFH.historyitem_type_Timing);
    CTiming.prototype.setBldLst = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TimingBldLst, this.bldLst, oPr));
        this.bldLst = oPr;
        this.setParentToChild(oPr);
    };
    CTiming.prototype.setTnLst = function(oPr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TimingTnLst, this.tnLst, oPr));
        this.tnLst = oPr;
        this.setParentToChild(oPr);
    };
    CTiming.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.bldLst) {
            oCopy.setBldLst(this.bldLst.createDuplicate(oIdMap));
        }
        if(this.tnLst) {
            oCopy.setTnLst(this.tnLst.createDuplicate(oIdMap));
        }
    };
    CTiming.prototype.privateWriteAttributes = function(pWriter) {
    };
    CTiming.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.tnLst);
        this.writeRecord2(pWriter, 1, this.bldLst);
    };
    CTiming.prototype.readAttribute = function(nType, pReader) {
    };
    CTiming.prototype.readChild = function(nType, pReader) {
        var oChild;
        switch(nType) {
            case 0: {
                oChild = new CTnLst();
                oChild.fromPPTY(pReader);
                this.setTnLst(oChild);
                break;
            }
            case 1: {
                oChild = new CBldLst();
                oChild.fromPPTY(pReader);
                this.setBldLst(oChild);
                break;
            }
        }
    };
    CTiming.prototype.getChildren = function() {
        return [this.bldLst, this.tnLst];
    };
    CTiming.prototype.onRemoveObject = function(sObjectId) {
        this.traverse(function(oNode) {
            oNode.handleRemoveObject(sObjectId);
            return false;
        });
    };
    CTiming.prototype.onRemoveChild = function(oChild) {
        if(oChild === this.tnLst) {
            this.setTnLst(null);
        }
        else if(oChild === this.bldLst) {
            this.setBldLst(null);
        }
    };


    changesFactory[AscDFH.historyitem_CommonTimingListAdd] = CChangeContent;
    changesFactory[AscDFH.historyitem_CommonTimingListRemove] = CChangeContent;
    drawingContentChanges[AscDFH.historyitem_CommonTimingListAdd] = function(oClass) {return oClass.list;};
    drawingContentChanges[AscDFH.historyitem_CommonTimingListRemove] = function(oClass) {return oClass.list;};
    function CCommonTimingList() {
        CBaseFormatObject.call(this);
        this.list = [];
    }
    InitClass(CCommonTimingList, CBaseFormatObject, AscDFH.historyitem_type_CommonTimingList);
    CCommonTimingList.prototype.addToLst = function(nIdx, oPr) {
        var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
        History.Add(new CChangeContent(this, AscDFH.historyitem_CommonTimingListAdd, nInsertIdx, [oPr], true));
        this.list.splice(nInsertIdx, 0, oPr);
        this.setParentToChild(oPr);
    };
    CCommonTimingList.prototype.removeFromLst = function(nIdx) {
        if(nIdx > -1 && nIdx < this.list.length) {
            this.list[nIdx].setParent(null);
            History.Add(new CChangeContent(this, AscDFH.historyitem_CommonTimingListRemove, nIdx, [this.list[nIdx]], false));
            this.list.splice(nIdx, 1);
        }
    };
    CCommonTimingList.prototype.fillObject = function(oCopy, oIdMap) {
        for(var nIdx = 0; nIdx < this.list.length; ++nIdx) {
            oCopy.addToLst(nIdx, this.list[nIdx].createDuplicate(oIdMap));
        }
    };
    CCommonTimingList.prototype.privateWriteAttributes = function(pWriter) {
    };
    CCommonTimingList.prototype.writeChildren = function(pWriter) {
        if(this.list.length > 0) {
            pWriter.StartRecord(0);
            pWriter.WriteULong(this.list.length);
            for(var nIndex = 0; nIndex < this.list.length; ++nIndex) {
                this.writeRecord1(pWriter, 0, this.list[nIndex]);
            }
            pWriter.EndRecord();
        }
    };
    CCommonTimingList.prototype.readAttribute = function(nType, pReader) {
    };
    CCommonTimingList.prototype.readElement = function(pReader) {
        var oStream = pReader.stream;
        oStream.GetUChar();
        oStream.SkipRecord();
        return null;
    };
    CCommonTimingList.prototype.readChild = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            oStream.GetULong();//skip record length
            var nLength = oStream.GetULong();
            for(var nIndex = 0; nIndex < nLength; ++nIndex) {
                var oElement = this.readElement(pReader);
                if(oElement) {
                    this.addToLst(this.list.length, oElement);
                }
            }
        }
    };
    CCommonTimingList.prototype.getChildren = function() {
        return [].concat(this.list);
    };
    CCommonTimingList.prototype.onRemoveChild = function(oChild) {
        if(this.parent) {
            for(var nIdx = this.list.length - 1; nIdx > -1; --nIdx) {
                if(this.list[nIdx] === oChild) {
                    this.removeFromLst(nIdx);
                    if(this.list.length === 0) {
                        this.parent.onRemoveChild(this);
                    }
                    return;
                }
            }
        }
    };

    function CAttrNameLst() {
        CCommonTimingList.call(this);
    }
    InitClass(CAttrNameLst, CCommonTimingList, AscDFH.historyitem_type_AttrNameLst);
    CAttrNameLst.prototype.readElement = function(pReader) {
        var oElement = new CAttrName();
        pReader.stream.GetUChar(); //skip ..
        oElement.fromPPTY(pReader);
        return oElement;
    };
    function CBldLst() {
        CCommonTimingList.call(this);
    }
    InitClass(CBldLst, CCommonTimingList, AscDFH.historyitem_type_BldLst);
    CBldLst.prototype.readElement = function(pReader) {
        var oStream = pReader.stream;
        var nType = oStream.GetUChar();
        var oElement = null;
        switch(nType) {
            case 1: oElement = new CBldDgm(); break;
            case 2: oElement = new CBldOleChart(); break;
            case 3: oElement = new CBldGraphic(); break;
            case 4: oElement = new CBldP(); break;
            default:break;
        }
        if(oElement) {
            oElement.fromPPTY(pReader);
        }
        return oElement;
    };
    CBldLst.prototype.writeChildren = function(pWriter) {
        if(this.list.length > 0) {
            pWriter.StartRecord(0);
            pWriter.WriteULong(this.list.length);
            for(var nIndex = 0; nIndex < this.list.length; ++nIndex) {
                var oElement = this.list[nIndex];
                var nType = null;
                switch (oElement.getObjectType()) {
                    case AscDFH.historyitem_type_BldDgm: nType = 1; break;
                    case AscDFH.historyitem_type_BldOleChart: nType = 2; break;
                    case AscDFH.historyitem_type_BldGraphic: nType = 3; break;
                    case AscDFH.historyitem_type_BldP: nType = 4; break;
                }
                if(nType !== null) {
                    this.writeRecord1(pWriter, nType, oElement);
                }
            }
            pWriter.EndRecord();
        }
    };

    function CCondLst() {
        CCommonTimingList.call(this);
    }
    InitClass(CCondLst, CCommonTimingList, AscDFH.historyitem_type_CondLst);
    CCondLst.prototype.readElement = function(pReader) {
        var oElement = new CCond();
        pReader.stream.GetUChar(); //skip ..
        oElement.fromPPTY(pReader);
        return oElement;
    };

    function CChildTnLst() {
        CCommonTimingList.call(this);
    }
    InitClass(CChildTnLst, CCommonTimingList, AscDFH.historyitem_type_ChildTnLst);
    CChildTnLst.prototype.readElement = function(pReader) {
        var oStream = pReader.stream;
        var nType = oStream.GetUChar();
        var oElement = null;
        switch(nType) {
            case 1: oElement = new CPar(); break;
            case 2: oElement = new CSeq(); break;
            case 3: oElement = new CAudio(); break;
            case 4: oElement = new CVideo(); break;
            case 5: oElement = new CExcl(); break;
            case 6: oElement = new CAnim(); break;
            case 7: oElement = new CAnimClr(); break;
            case 8: oElement = new CAnimEffect(); break;
            case 9: oElement = new CAnimMotion(); break;
            case 10: oElement = new CAnimRot(); break;
            case 11: oElement = new CAnimScale(); break;
            case 12: oElement = new CCmd(); break;
            case 13: oElement = new CSet(); break;
            default:break;
        }
        if(oElement) {
            oElement.fromPPTY(pReader);
        }
        return oElement;
    };
    CChildTnLst.prototype.writeChildren = function(pWriter) {
        if(this.list.length > 0) {
            pWriter.StartRecord(0);
            pWriter.WriteULong(this.list.length);
            for(var nIndex = 0; nIndex < this.list.length; ++nIndex) {
                var oElement = this.list[nIndex];
                var nType = null;
                switch (oElement.getObjectType()) {
                    case AscDFH.historyitem_type_Par: nType = 1; break;
                    case AscDFH.historyitem_type_Seq: nType = 2; break;
                    case AscDFH.historyitem_type_Audio: nType = 3; break;
                    case AscDFH.historyitem_type_Video: nType = 4; break;
                    case AscDFH.historyitem_type_Excl: nType = 5; break;
                    case AscDFH.historyitem_type_Anim: nType = 6; break;
                    case AscDFH.historyitem_type_AnimClr: nType = 7; break;
                    case AscDFH.historyitem_type_AnimEffect: nType = 8; break;
                    case AscDFH.historyitem_type_AnimMotion: nType = 9; break;
                    case AscDFH.historyitem_type_AnimRot: nType = 10; break;
                    case AscDFH.historyitem_type_AnimScale: nType = 11; break;
                    case AscDFH.historyitem_type_Cmd: nType = 12; break;
                    case AscDFH.historyitem_type_Set : nType = 13; break;
                }
                if(nType !== null) {
                    this.writeRecord1(pWriter, nType, oElement);
                }
            }
            pWriter.EndRecord();
        }
    };


    function CTmplLst() {
        CCommonTimingList.call(this);
    }
    InitClass(CTmplLst, CCommonTimingList, AscDFH.historyitem_type_TmplLst);
    CTmplLst.prototype.readElement = function(pReader) {
        var oElement = new CTmpl();
        pReader.stream.GetUChar(); //skip ..
        oElement.fromPPTY(pReader);
        return oElement;
    };

    function CTnLst() {
        CChildTnLst.call(this);
    }
    InitClass(CTnLst, CChildTnLst, AscDFH.historyitem_type_TnLst);


    function CTavLst() {
        CCommonTimingList.call(this);
    }
    InitClass(CTavLst, CCommonTimingList, AscDFH.historyitem_type_TavLst);
    CTavLst.prototype.readElement = function(pReader) {
        var oElement = new CTav();
        pReader.stream.GetUChar(); //skip ..
        oElement.fromPPTY(pReader);
        return oElement;
    };

    changesFactory[AscDFH.historyitem_ObjectTargetSpid] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_ObjectTargetSpid] = function(oClass, value) {oClass.spid = value;};
    function CObjectTarget() {//subsp
        CBaseFormatObject.call(this);
        this.spid = null;
    }
    InitClass(CObjectTarget, CBaseFormatObject, AscDFH.historyitem_type_ObjectTarget);
    CObjectTarget.prototype.setSpid = function(pr, pReader) {
        if(pReader) {
            pReader.AddConnectedObject(this);
        }
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_ObjectTargetSpid, this.spid, pr));
        this.spid = pr;
    };
    CObjectTarget.prototype.assignConnection = function(oObjectsMap) {
        if(AscCommon.isRealObject(oObjectsMap[this.spid])){
            this.setSpid(oObjectsMap[this.spid].Id);
        }
        else {
            if(this.parent) {
                this.parent.onRemoveChild(this);
            }
        }
    };
    CObjectTarget.prototype.fillObject = function(oCopy, oIdMap) {
        var sSpId = this.spid;
        if(oIdMap && oIdMap[this.spid]) {
            sSpId = oIdMap[this.spid];
        }
        oCopy.setSpid(sSpId);
    };
    CObjectTarget.prototype.privateWriteAttributes = function(pWriter) {
    };
    CObjectTarget.prototype.writeChildren = function(pWriter) {
    };
    CObjectTarget.prototype.readAttribute = function(nType, pReader) {
    };
    CObjectTarget.prototype.readChild = function(nType, pReader) {
    };
    CObjectTarget.prototype.handleRemoveObject = function(sObjectId) {
        if(this.spid === sObjectId) {
            if(this.parent) {
                this.parent.onRemoveChild(this);
            }
        }
    };

    changesFactory[AscDFH.historyitem_BldBaseGrpId] = CChangeLong;
    changesFactory[AscDFH.historyitem_BldBaseUIExpand] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_BldBaseGrpId] = function(oClass, value) {oClass.grpId = value;};
    drawingsChangesMap[AscDFH.historyitem_BldBaseUIExpand] = function(oClass, value) {oClass.uiExpand = value;};
    function CBldBase() {
        CObjectTarget.call(this);
        this.grpId = null;
        this.uiExpand = null;
    }
    InitClass(CBldBase, CObjectTarget, AscDFH.historyitem_type_BldBase);
    CBldBase.prototype.setGrpId = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BldBaseGrpId, this.grpId, pr));
        this.grpId = pr;
    };

    CBldBase.prototype.assignConnection = function(oObjectsMap) {
        if(AscCommon.isRealObject(oObjectsMap[this.spid]) &&
            (oObjectsMap[this.spid].getObjectType && oObjectsMap[this.spid].getObjectType() === AscDFH.historyitem_type_ChartSpace)) {
            this.setSpid(oObjectsMap[this.spid].Id);
        }
        else {
            if(this.parent) {
                this.parent.onRemoveChild(this);
            }
        }
    };
    CBldBase.prototype.setUiExpand = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BldBaseUIExpand, this.uiExpand, pr));
        this.uiExpand = pr;
    };
    CBldBase.prototype.fillObject = function(oCopy, oIdMap) {
        CObjectTarget.prototype.fillObject.call(this, oCopy, oIdMap);
        oCopy.setGrpId(this.grpId);
        oCopy.setUiExpand(this.uiExpand);
    };
    CBldBase.prototype.privateWriteAttributes = function(pWriter) {
    };
    CBldBase.prototype.writeChildren = function(pWriter) {
    };
    CBldBase.prototype.readAttribute = function(nType, pReader) {
    };
    CBldBase.prototype.readChild = function(nType, pReader) {
    };

    changesFactory[AscDFH.historyitem_BldDgmBld] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_BldDgmBld] = function(oClass, value) {oClass.bld = value;};
    function CBldDgm() {
        CBldBase.call(this);
        this.bld = null;
    }
    InitClass(CBldDgm, CBldBase, AscDFH.historyitem_type_BldDgm);
    CBldDgm.prototype.setBld = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BldDgmBld, this.bld, pr));
        this.bld = pr;
    };
    CBldDgm.prototype.fillObject = function(oCopy, oIdMap) {
        CBldBase.prototype.fillObject.call(this, oCopy, oIdMap);
        oCopy.setBld(this.bld);
    };
    CBldDgm.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.bld);
        pWriter._WriteBool2(1, this.uiExpand);
        var nSpId = pWriter.GetSpIdxId(this.spid);
        if(nSpId !== null) {
            pWriter._WriteString1(2, nSpId + "");
        }
        pWriter._WriteInt1(3, this.grpId);
    };
    CBldDgm.prototype.writeChildren = function(pWriter) {
    };
    CBldDgm.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setBld(oStream.GetUChar());
        else if (1 === nType) this.setUiExpand(oStream.GetBool());
        else if (2 === nType) this.setSpid(oStream.GetString2(), pReader);
        else if (3 === nType) this.setGrpId(oStream.GetLong());
    };
    CBldDgm.prototype.readChild = function(nType, pReader) {
        pReader.stream.SkipRecord();
    };

    changesFactory[AscDFH.historyitem_BldGraphicBldAsOne] = CChangeObject;
    changesFactory[AscDFH.historyitem_BldGraphicBldSub] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_BldGraphicBldAsOne] = function(oClass, value) {oClass.bldAsOne = value;};
    drawingsChangesMap[AscDFH.historyitem_BldGraphicBldSub] = function(oClass, value) {oClass.bldSub = value;};
    function CBldGraphic() {
        CBldBase.call(this);
        this.bldAsOne = null;
        this.bldSub = null;
    }
    InitClass(CBldGraphic, CBldBase, AscDFH.historyitem_type_BldGraphic);
    CBldGraphic.prototype.setBldAsOne = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_BldGraphicBldAsOne, this.bldAsOne, pr));
        this.bldAsOne = pr;
        this.setParentToChild(pr);
    };
    CBldGraphic.prototype.setBldSub = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_BldGraphicBldSub, this.bldSub, pr));
        this.bldSub = pr;
        this.setParentToChild(pr);
    };
    CBldGraphic.prototype.fillObject = function(oCopy, oIdMap) {
        CBldBase.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.bldAsOne) {
            oCopy.setBldAsOne(this.bldAsOne.createDuplicate(oIdMap));
        }
        if(this.bldSub) {
            oCopy.setBldAsOne(this.bldSub.createDuplicate(oIdMap));
        }
    };
    CBldGraphic.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteBool2(0, this.uiExpand);
        var nSpId = pWriter.GetSpIdxId(this.spid);
        if(nSpId !== null) {
            pWriter._WriteString1(1, nSpId + "");
        }
        pWriter._WriteInt1(2, this.grpId);
    };
    CBldGraphic.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.bldSub);
    };
    CBldGraphic.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 == nType) this.setUiExpand(oStream.GetBool());
        else if (1 == nType) this.setSpid(oStream.GetString2(), pReader);
        else if (2 == nType) this.setGrpId(oStream.GetLong());
    };
    CBldGraphic.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setBldSub(new CBldSub());
            this.bldSub.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CBldGraphic.prototype.getChildren = function() {
        return [this.bldSub];
    };

    changesFactory[AscDFH.historyitem_BldOleChartAnimBg] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_BldOleChartAnimBg] = function(oClass, value) {oClass.animBg = value;};
    function CBldOleChart() {
        CBldDgm.call(this);
        this.animBg = null;
    }
    InitClass(CBldOleChart, CBldDgm, AscDFH.historyitem_type_BldOleChart);
    CBldOleChart.prototype.setAnimBg = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BldOleChartAnimBg, this.animBg, pr));
        this.animBg = pr;
    };
    CBldOleChart.prototype.fillObject = function(oCopy, oIdMap) {
        CBldDgm.prototype.fillObject.call(this, oCopy, oIdMap);
        oCopy.setAnimBg(this.animBg);
    };
    CBldOleChart.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.bld);
        pWriter._WriteBool2(1, this.uiExpand);
        var nSpId = pWriter.GetSpIdxId(this.spid);
        if(nSpId !== null) {
            pWriter._WriteString1(2, nSpId + "");
        }
        pWriter._WriteInt1(3, this.grpId);
        pWriter._WriteBool2(4, this.animBg);
    };
    CBldOleChart.prototype.writeChildren = function(pWriter) {
    };
    CBldOleChart.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setBld(oStream.GetUChar());
        else if (1 === nType) this.setUiExpand(oStream.GetBool());
        else if (2 === nType) this.setSpid(oStream.GetString2(), pReader);
        else if (3 === nType) this.setGrpId(oStream.GetLong());
        else if (4 === nType) this.setAnimBg(oStream.GetBool());
    };
    CBldOleChart.prototype.readChild = function(nType, pReader) {
        pReader.stream.SkipRecord();
    };


    changesFactory[AscDFH.historyitem_BldPTmplLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_BldPAdvAuto] = CChangeLong;
    changesFactory[AscDFH.historyitem_BldPAutoUpdateAnimBg] = CChangeBool;
    changesFactory[AscDFH.historyitem_BldPBldLvl] = CChangeLong;
    changesFactory[AscDFH.historyitem_BldPBuild] = CChangeLong;
    changesFactory[AscDFH.historyitem_BldPRev] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_BldPTmplLst] = function(oClass, value) {oClass.tmplLst = value;};
    drawingsChangesMap[AscDFH.historyitem_BldPAdvAuto] = function(oClass, value) {oClass.advAuto = value;};
    drawingsChangesMap[AscDFH.historyitem_BldPAutoUpdateAnimBg] = function(oClass, value) {oClass.autoUpdateAnimBg = value;};
    drawingsChangesMap[AscDFH.historyitem_BldPBldLvl] = function(oClass, value) {oClass.bldLvl = value;};
    drawingsChangesMap[AscDFH.historyitem_BldPBuild] = function(oClass, value) {oClass.build = value;};
    drawingsChangesMap[AscDFH.historyitem_BldPRev] = function(oClass, value) {oClass.rev = value;};

    function CBldP() {
        CBldOleChart.call(this);
        this.tmplLst = null;
        this.advAuto = null;
        this.autoUpdateAnimBg = null;
        this.bldLvl = null;
        this.build = null;
        this.rev = null;
    }
    InitClass(CBldP, CBldOleChart, AscDFH.historyitem_type_BldP);
    CBldP.prototype.setTmplLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_BldPTmplLst, this.tmplLst, pr));
        this.tmplLst = pr;
        this.setParentToChild(pr);
    };
    CBldP.prototype.setAdvAuto = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BldPAdvAuto, this.advAuto, pr));
        this.advAuto = pr;
    };
    CBldP.prototype.setAutoUpdateAnimBg = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BldPAutoUpdateAnimBg, this.autoUpdateAnimBg, pr));
        this.autoUpdateAnimBg = pr;
    };
    CBldP.prototype.setBldLvl = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BldPBldLvl, this.bldLvl, pr));
        this.bldLvl = pr;
    };
    CBldP.prototype.setBuild = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BldPBuild, this.build, pr));
        this.build = pr;
    };
    CBldP.prototype.setRev = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BldPRev, this.rev, pr));
        this.rev = pr;
    };
    CBldP.prototype.fillObject = function(oCopy, oIdMap) {
        CBldOleChart.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.tmplLst) {
            oCopy.setTmplLst(this.tmplLst.createDuplicate(oIdMap));
        }
        oCopy.setAdvAuto(this.advAuto);
        oCopy.setAutoUpdateAnimBg(this.autoUpdateAnimBg);
        oCopy.setBldLvl(this.bldLvl);
        oCopy.setBuild(this.build);
        oCopy.setRev(this.rev);
    };
    CBldP.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.build);
        pWriter._WriteBool2(1, this.uiExpand);
        var nSpId = pWriter.GetSpIdxId(this.spid);
        if(nSpId !== null) {
            pWriter._WriteString1(2, nSpId + "");
        }
        pWriter._WriteInt1(3, this.grpId);
        pWriter._WriteInt2(4, this.bldLvl);
        pWriter._WriteBool2(5, this.animBg);
        pWriter._WriteBool2(6, this.autoUpdateAnimBg);
        pWriter._WriteBool2(7, this.rev);
        pWriter._WriteString2(8, this.advAuto);
    };
    CBldP.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.tmplLst);
    };
    CBldP.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setBuild(oStream.GetUChar());
        else if (1 === nType) this.setUiExpand(oStream.GetBool());
        else if (2 === nType) this.setSpid(oStream.GetString2(), pReader);
        else if (3 === nType) this.setGrpId(oStream.GetLong());
        else if (4 === nType) this.setBldLvl(oStream.GetLong());
        else if (5 === nType) this.setAnimBg(oStream.GetBool());
        else if (6 === nType) this.setAutoUpdateAnimBg(oStream.GetBool());
        else if (7 === nType) this.setRev(oStream.GetBool());
        else if (8 === nType) this.setAdvAuto(oStream.GetString2());
    };
    CBldP.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setTmplLst(new CTmplLst());
            this.tmplLst.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CBldP.prototype.getChildren = function() {
        return [this.tmplLst];
    };


    changesFactory[AscDFH.historyitem_BldSubChart] = CChangeBool;
    changesFactory[AscDFH.historyitem_BldSubAnimBg] = CChangeBool;
    changesFactory[AscDFH.historyitem_BldSubRev] = CChangeBool;
    changesFactory[AscDFH.historyitem_BldSubBldChart] = CChangeLong;
    changesFactory[AscDFH.historyitem_BldSubBldDgm] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_BldSubBldChart] = function(oClass, value) {oClass.bldChart = value;};
    drawingsChangesMap[AscDFH.historyitem_BldSubBldDgm] = function(oClass, value) {oClass.bldDgm = value;};
    drawingsChangesMap[AscDFH.historyitem_BldSubChart] = function(oClass, value) {oClass.chart = value;};
    drawingsChangesMap[AscDFH.historyitem_BldSubAnimBg] = function(oClass, value) {oClass.animBg = value;};
    drawingsChangesMap[AscDFH.historyitem_BldSubRev] = function(oClass, value) {oClass.rev = value;};
    function CBldSub() {
        CBaseFormatObject.call(this);
        this.chart = null;
        this.animBg = null;
        this.bldChart = null;
        this.bldDgm = null;
        this.rev = null;
    }
    InitClass(CBldSub, CBaseFormatObject, AscDFH.historyitem_type_BldSub);
    CBldSub.prototype.setBldChart = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BldSubBldChart, this.bldChart, pr));
        this.bldChart = pr;
    };
    CBldSub.prototype.setBldDgm = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BldSubBldDgm, this.bldDgm, pr));
        this.bldDgm = pr;
    };
    CBldSub.prototype.setChart = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BldSubChart, this.chart, pr));
        this.chart = pr;
    };
    CBldSub.prototype.setAnimBg = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BldSubAnimBg, this.animBg, pr));
        this.animBg = pr;
    };
    CBldSub.prototype.setRev = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BldSubRev, this.rev, pr));
        this.rev = pr;
    };
    CBldSub.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.chart !== null) {
            oCopy.setChart(this.chart);
        }
        if(this.animBg !== null) {
            oCopy.setAnimBg(this.animBg);
        }

        if(this.bldChart !== null) {
            oCopy.setBldChart(this.bldChart);
        }
        if(this.bldDgm !== null) {
            oCopy.setBldDgm(this.bldDgm);
        }
        if(this.rev !== null) {
            oCopy.setRev(this.rev);
        }
    };
    CBldSub.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteBool2(0, this.chart);
        pWriter._WriteBool2(1, this.animBg);
        pWriter._WriteLimit2(2, this.bldChart);
        pWriter._WriteLimit2(3, this.bldDgm);
        pWriter._WriteBool2(4, this.rev);
    };
    CBldSub.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setChart(oStream.GetBool());
        else if (1 === nType) this.setAnimBg(oStream.GetBool());
        else if (2 === nType) this.setBldChart(oStream.GetUChar());
        else if (3 === nType) this.setBldDgm(oStream.GetUChar());
        else if (4 === nType) this.setRev(oStream.GetBool());
    };

    changesFactory[AscDFH.historyitem_DirTransitionDir] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_DirTransitionDir] = function(oClass, value) {oClass.dir = value;};
    function CDirTransition() {//CBlinds, checker, comb, cover, pull, push, randomBar, strips, wipe, zoom
        CBaseFormatObject.call(this);
        this.dir = null;
    }
    InitClass(CDirTransition, CBaseFormatObject, AscDFH.historyitem_type_DirTransition);
    CDirTransition.prototype.setDir = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_DirTransitionDir, this.dir, pr));
        this.dir = pr;
    };
    CDirTransition.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setDir(this.dir);
    };
    CDirTransition.prototype.privateWriteAttributes = function(pWriter) {
    };
    CDirTransition.prototype.writeChildren = function(pWriter) {
    };
    CDirTransition.prototype.readAttribute = function(nType, pReader) {
    };
    CDirTransition.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_OptBlackTransitionThruBlk] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_OptBlackTransitionThruBlk] = function(oClass, value) {oClass.thruBlk = value;};
    function COptionalBlackTransition() {//cut, fade
        CBaseFormatObject.call(this);
        this.thruBlk = null;
    }
    InitClass(COptionalBlackTransition, CBaseFormatObject, AscDFH.historyitem_type_OptBlackTransition);
    COptionalBlackTransition.prototype.setThruBlk = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_OptBlackTransitionThruBlk, this.thruBlk, pr));
        this.thruBlk = pr;
    };
    COptionalBlackTransition.prototype.fillObject = function(oCopy, oIdMap) {
        oCopy.setThruBlk(this.thruBlk);
    };
    COptionalBlackTransition.prototype.privateWriteAttributes = function(pWriter) {
    };
    COptionalBlackTransition.prototype.writeChildren = function(pWriter) {
    };
    COptionalBlackTransition.prototype.readAttribute = function(nType, pReader) {
    };
    COptionalBlackTransition.prototype.readChild = function(nType, pReader) {
    };

    changesFactory[AscDFH.historyitem_GraphicElDgmId] = CChangeString;
    changesFactory[AscDFH.historyitem_GraphicElDgmBuildStep] = CChangeLong;
    changesFactory[AscDFH.historyitem_GraphicElChartBuildStep] = CChangeLong;
    changesFactory[AscDFH.historyitem_GraphicElSeriesIdx] = CChangeLong;
    changesFactory[AscDFH.historyitem_GraphicElCategoryIdx] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_GraphicElDgmId] = function(oClass, value) {oClass.dgmId = value;};
    drawingsChangesMap[AscDFH.historyitem_GraphicElDgmBuildStep] = function(oClass, value) {oClass.dgmBuildStep = value;};
    drawingsChangesMap[AscDFH.historyitem_GraphicElChartBuildStep] = function(oClass, value) {oClass.chartBuildStep = value;};
    drawingsChangesMap[AscDFH.historyitem_GraphicElSeriesIdx] = function(oClass, value) {oClass.seriesIdx = value;};
    drawingsChangesMap[AscDFH.historyitem_GraphicElCategoryIdx] = function(oClass, value) {oClass.categoryIdx = value;};

    function CGraphicEl() {
        CBaseFormatObject.call(this);
        this.dgmId = null;
        this.dgmBuildStep = null;
        this.chartBuildStep = null;
        this.seriesIdx = null;
        this.categoryIdx = null;
    }
    InitClass(CGraphicEl, CBaseFormatObject, AscDFH.historyitem_type_GraphicEl);
    CGraphicEl.prototype.setDgmId = function(pr, pReader) {
        if(pReader) {
            pReader.AddConnectedObject(this);
        }
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_GraphicElDgmId, this.dgmId, pr));
        this.dgmId = pr;
    };
    CGraphicEl.prototype.assignConnection = function(oObjectsMap) {
        if(AscCommon.isRealObject(oObjectsMap[this.spid])){
            this.setSpid(oObjectsMap[this.dgmId].Id);
        }
        else {
            if(this.parent) {
                this.parent.onRemoveChild(this);
            }
        }
    };
    CGraphicEl.prototype.setDgmBuildStep = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_GraphicElDgmBuildStep, this.dgmBuildStep, pr));
        this.dgmBuildStep = pr;
    };
    CGraphicEl.prototype.setChartBuildStep = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_GraphicElChartBuildStep, this.chartBuildStep, pr));
        this.chartBuildStep = pr;
    };
    CGraphicEl.prototype.setSeriesIdx = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_GraphicElSeriesIdx, this.dgmId, pr));
        this.seriesIdx = pr;
    };
    CGraphicEl.prototype.setCategoryIdx = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_GraphicElCategoryIdx, this.dgmId, pr));
        this.categoryIdx = pr;
    };
    CGraphicEl.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.dgmId !== null) {
            var sDgmId = this.dgmId;
            if(oIdMap && oIdMap[this.dgmId]) {
                sDgmId = oIdMap[this.dgmId];
            }
            oCopy.setDgmId(sDgmId);
        }
        if(this.dgmBuildStep !== null) {
            oCopy.setDgmBuildStep(this.dgmBuildStep);
        }
        if(this.chartBuildStep !== null) {
            oCopy.setChartBuildStep(this.chartBuildStep);
        }
        if(this.seriesIdx !== null) {
            oCopy.setSeriesIdx(this.seriesIdx);
        }
        if(this.categoryIdx !== null) {
            oCopy.setCategoryIdx(this.categoryIdx);
        }
    };
    CGraphicEl.prototype.privateWriteAttributes = function(pWriter) {
        var nSpId = pWriter.GetSpIdxId(this.dgmId);
        if(nSpId !== null) {
            pWriter._WriteString2(0, nSpId + "");
        }
        pWriter._WriteLimit2(1, this.dgmBuildStep);
        pWriter._WriteLimit2(2, this.chartBuildStep);
        pWriter._WriteInt2(3, this.seriesIdx);
        pWriter._WriteInt2(4, this.categoryIdx);
    };
    CGraphicEl.prototype.writeChildren = function(pWriter) {
    };
    CGraphicEl.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setDgmId(oStream.GetString2(), pReader);
        else if (1 === nType) this.setDgmBuildStep(oStream.GetUChar());
        else if (2 === nType) this.setChartBuildStep(oStream.GetUChar());
        else if (3 === nType) this.setSeriesIdx(oStream.GetLong());
        else if (4 === nType) this.setCategoryIdx(oStream.GetLong());
    };
    CGraphicEl.prototype.readChild = function(nType, pReader) {
        pReader.stream.SkipRecord();
    };
    CGraphicEl.prototype.handleRemoveObject = function(sObjectId) {
        if(this.dgmId === sObjectId) {
            if(this.parent) {
                this.parent.onRemoveChild(this);
            }
        }
    };

    changesFactory[AscDFH.historyitem_IndexRgSt] = CChangeLong;
    changesFactory[AscDFH.historyitem_IndexRgEnd] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_IndexRgSt] = function(oClass, value) {oClass.st = value;};
    drawingsChangesMap[AscDFH.historyitem_IndexRgEnd] = function(oClass, value) {oClass.st = value;};
    function CIndexRg() {//charrg, pRg
        CBaseFormatObject.call(this);
        this.st = null;
        this.end = null;
    }
    InitClass(CIndexRg, CBaseFormatObject, AscDFH.historyitem_type_IndexRg);
    CIndexRg.prototype.setSt = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_IndexRgSt, this.st, pr));
        this.st = pr;
    };
    CIndexRg.prototype.setEnd = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_IndexRgEnd, this.end, pr));
        this.end = pr;
    };
    CIndexRg.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.end !== null) {
            oCopy.setEnd(this.end);
        }
        if(this.st !== null) {
            oCopy.setSt(this.st);
        }
    };
    CIndexRg.prototype.privateWriteAttributes = function(pWriter) {
    };
    CIndexRg.prototype.writeChildren = function(pWriter) {
    };
    CIndexRg.prototype.readAttribute = function(nType, pReader) {
    };
    CIndexRg.prototype.readChild = function(nType, pReader) {
    };

    changesFactory[AscDFH.historyitem_TmplLvl] = CChangeLong;
    changesFactory[AscDFH.historyitem_TmplTnLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TmplLvl] = function(oClass, value) {oClass.lvl = value;};
    drawingsChangesMap[AscDFH.historyitem_TmplTnLst] = function(oClass, value) {oClass.tnLst = value;};
    function CTmpl() {
        CBaseFormatObject.call(this);
        this.lvl = null;
        this.tnLst = null
    }
    InitClass(CTmpl, CBaseFormatObject, AscDFH.historyitem_type_Tmpl);
    CTmpl.prototype.setLvl = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_TmplLvl, this.lvl, pr));
        this.lvl = pr;
    };
    CTmpl.prototype.setTnLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TmplTnLst, this.tnLst, pr));
        this.tnLst = pr;
        this.setParentToChild(pr);
    };
    CTmpl.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.lvl !== null) {
            oCopy.setLvl(this.lvl);
        }
        if(this.tnLst !== null) {
            oCopy.setTnLst(this.tnLst.createDuplicate(oIdMap));
        }
    };
    CTmpl.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteInt2(0, this.lvl);
    };
    CTmpl.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.tnLst);
    };
    CTmpl.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
     if(0 === nType) {
            this.setLvl(oStream.GetLong())
        }
    };
    CTmpl.prototype.readChild = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setTnLst(new CTnLst());
            this.tnLst.fromPPTY(pReader);
        }
        else {
            oStream.SkipRecord();
        }
    };
    CTmpl.prototype.getChildren = function() {
        return [this.tnLst];
    };

    changesFactory[AscDFH.historyitem_AnimCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimTavLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimBy] = CChangeString;
    changesFactory[AscDFH.historyitem_AnimCalcmode] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimFrom] = CChangeString;
    changesFactory[AscDFH.historyitem_AnimTo] = CChangeString;
    changesFactory[AscDFH.historyitem_AnimValueType] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_AnimCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimTavLst] = function(oClass, value) {oClass.tavLst = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimBy] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimCalcmode] = function(oClass, value) {oClass.calcmode = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimFrom] = function(oClass, value) {oClass.from = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimTo] = function(oClass, value) {oClass.to = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimValueType] = function(oClass, value) {oClass.valueType = value;};
    function CAnim() {
        CBaseFormatObject.call(this);
        this.cBhvr = null;
        this.tavLst = null;
        this.by = null;
        this.calcmode = null;
        this.from = null;
        this.to = null;
        this.valueType = null;
    }
    InitClass(CAnim, CBaseFormatObject, AscDFH.historyitem_type_Anim);
    CAnim.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CAnim.prototype.setTavLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimTavLst, this.tavLst, pr));
        this.tavLst = pr;
        this.setParentToChild(pr);
    };
    CAnim.prototype.setBy = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimBy, this.by, pr));
        this.by = pr;
    };
    CAnim.prototype.setCalcmode = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimCalcmode, this.calcmode, pr));
        this.calcmode = pr;
    };
    CAnim.prototype.setFrom = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimFrom, this.from, pr));
        this.from = pr;
    };
    CAnim.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimTo, this.to, pr));
        this.to = pr;
    };
    CAnim.prototype.setValueType = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimValueType, this.valueType, pr));
        this.valueType = pr;
    };
    CAnim.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.tavLst !== null) {
            oCopy.setTavLst(this.tavLst.createDuplicate(oIdMap));
        }
        if(this.by !== null) {
            oCopy.setBy(this.by);
        }
        if(this.calcmode !== null) {
            oCopy.setCalcmode(this.calcmode);
        }
        if(this.from !== null) {
            oCopy.setFrom(this.from);
        }
        if(this.to !== null) {
            oCopy.setTo(this.to);
        }
        if(this.valueType !== null) {
            oCopy.setValueType(this.valueType);
        }
    };
    CAnim.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0,  this.calcmode);
        pWriter._WriteString2(1, this.by);
        pWriter._WriteString2(2, this.from);
        pWriter._WriteString2(3, this.to);
        pWriter._WriteLimit2(4,  this.valueType);
    };
    CAnim.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);
        this.writeRecord2(pWriter, 1, this.tavLst);
    };
    CAnim.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType)     this.setCalcmode(oStream.GetUChar());
        else if (1 === nType) this.setBy(oStream.GetString2());
        else if (2 === nType) this.setFrom(oStream.GetString2());
        else if (3 === nType) this.setTo(oStream.GetString2());
        else if (4 === nType) this.setValueType(oStream.GetUChar());
    };
    CAnim.prototype.readChild = function(nType, pReader) {
        var s = this.stream;
        switch (nType) {
            case 0: {
                this.setCBhvr(new CCBhvr());
                this.cBhvr.fromPPTY(pReader);
                break;
            }
            case 1: {
                this.setTavLst(new CTavLst());
                this.tavLst.fromPPTY(pReader);
                break;
            }
            default: {
                s.SkipRecord();
                break;
            }
        }
    };
    CAnim.prototype.getChildren = function() {
        return [this.cBhvr, this.tavLst];
    };

    changesFactory[AscDFH.historyitem_CBhvrAttrNameLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_CBhvrCTn] = CChangeObject;
    changesFactory[AscDFH.historyitem_CBhvrTgtEl] = CChangeObject;
    changesFactory[AscDFH.historyitem_CBhvrAccumulate] = CChangeLong;
    changesFactory[AscDFH.historyitem_CBhvrAdditive] = CChangeLong;
    changesFactory[AscDFH.historyitem_CBhvrBy] = CChangeString;
    changesFactory[AscDFH.historyitem_CBhvrFrom] = CChangeString;
    changesFactory[AscDFH.historyitem_CBhvrOverride] = CChangeLong;
    changesFactory[AscDFH.historyitem_CBhvrRctx] = CChangeString;
    changesFactory[AscDFH.historyitem_CBhvrTo] = CChangeString;
    changesFactory[AscDFH.historyitem_CBhvrXfrmType] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_CBhvrAttrNameLst] = function(oClass, value) {oClass.attrNameLst = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrCTn] = function(oClass, value) {oClass.cTn = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrTgtEl] = function(oClass, value) {oClass.tgtEl = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrAccumulate] = function(oClass, value) {oClass.accumulate = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrAdditive] = function(oClass, value) {oClass.additive = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrBy] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrFrom] = function(oClass, value) {oClass.from = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrOverride] = function(oClass, value) {oClass.override = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrRctx] = function(oClass, value) {oClass.rctx = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrTo] = function(oClass, value) {oClass.to = value;};
    drawingsChangesMap[AscDFH.historyitem_CBhvrXfrmType] = function(oClass, value) {oClass.xfrmType = value;};
    function CCBhvr() {
        CBaseFormatObject.call(this);
        this.attrNameLst = null;
        this.cTn = null;
        this.tgtEl = null;
        this.accumulate = null;
        this.additive = null;
        this.by = null;
        this.from = null;
        this.override = null;
        this.rctx = null;
        this.to = null;
        this.xfrmType = null;
    }
    InitClass(CCBhvr, CBaseFormatObject, AscDFH.historyitem_type_CBhvr);
    CCBhvr.prototype.setAttrNameLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CBhvrAttrNameLst, this.attrNameLst, pr));
        this.attrNameLst = pr;
        this.setParentToChild(pr);
    };
    CCBhvr.prototype.setCTn = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CBhvrCTn, this.cTn, pr));
        this.cTn = pr;
        this.setParentToChild(pr);
    };
    CCBhvr.prototype.setTgtEl = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CBhvrTgtEl, this.tgtEl, pr));
        this.tgtEl = pr;
        this.setParentToChild(pr);
    };
    CCBhvr.prototype.setAccumulate = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CBhvrAccumulate, this.accumulate, pr));
        this.accumulate = pr;
    };
    CCBhvr.prototype.setAdditive = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CBhvrAdditive, this.additive, pr));
        this.additive = pr;
    };
    CCBhvr.prototype.setBy = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CBhvrBy, this.by, pr));
        this.by = pr;
    };
    CCBhvr.prototype.setFrom = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CBhvrFrom, this.from, pr));
        this.from = pr;
    };
    CCBhvr.prototype.setOverride = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CBhvrOverride, this.override, pr));
        this.override = pr;
    };
    CCBhvr.prototype.setRctx = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CBhvrRctx, this.rctx, pr));
        this.rctx = pr;
    };
    CCBhvr.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CBhvrTo, this.to, pr));
        this.to = pr;
    };
    CCBhvr.prototype.setXfrmType = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CBhvrXfrmType, this.xfrmType, pr));
        this.xfrmType = pr;
    };
    CCBhvr.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.attrNameLst !== null) {
            oCopy.setAttrNameLst(this.attrNameLst.createDuplicate(oIdMap));
        }
        if(this.cTn !== null) {
            oCopy.setCTn(this.cTn.createDuplicate(oIdMap));
        }
        if(this.tgtEl !== null) {
            oCopy.setTgtEl(this.tgtEl.createDuplicate(oIdMap));
        }
        if(this.accumulate !== null) {
            oCopy.setAccumulate(this.accumulate);
        }
        if(this.additive !== null) {
            oCopy.setAdditive(this.additive);
        }
        if(this.by !== null) {
            oCopy.setBy(this.by);
        }
        if(this.from !== null) {
            oCopy.setFrom(this.from);
        }
        if(this.override !== null) {
            oCopy.setOverride(this.override);
        }
        if(this.rctx !== null) {
            oCopy.setRctx(this.rctx);
        }
        if(this.to !== null) {
            oCopy.setTo(this.to);
        }
        if(this.xfrmType !== null) {
            oCopy.setXfrmType(this.xfrmType);
        }
    };
    CCBhvr.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.accumulate);
        pWriter._WriteLimit2(1, this.additive);
        pWriter._WriteString2(2, this.by);
        pWriter._WriteString2(3, this.from);
        pWriter._WriteLimit2(4, this.override);
        pWriter._WriteString2(5, this.rctx);
        pWriter._WriteString2(6, this.to);
        pWriter._WriteLimit2(7, this.xfrmType);
    };
    CCBhvr.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cTn);
        this.writeRecord1(pWriter, 1, this.tgtEl);
        this.writeRecord2(pWriter, 2, this.attrNameLst);
    };
    CCBhvr.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setAccumulate(oStream.GetUChar());
        else if (1 === nType) this.setAdditive(oStream.GetUChar());
        else if (2 === nType) this.setBy(oStream.GetString2());
        else if (3 === nType) this.setFrom(oStream.GetString2());
        else if (4 === nType) this.setOverride(oStream.GetUChar());
        else if (5 === nType) this.setRctx(oStream.GetString2());
        else if (6 === nType) this.setTo(oStream.GetString2());
        else if (7 === nType) this.setXfrmType(oStream.GetUChar());
    };
    CCBhvr.prototype.readChild = function(nType, pReader) {
        var oStream = pReader.stream;
        switch (nType) {
            case 0: {
                this.setCTn(new CCTn());
                this.cTn.fromPPTY(pReader);
                break;
            }
            case 1: {
                this.setTgtEl(new CTgtEl());
                this.tgtEl.fromPPTY(pReader);
                break;
            }
            case 2: {
                this.setAttrNameLst(new CAttrNameLst());
                this.attrNameLst.fromPPTY(pReader);
                break;
            }
            default: {
                oStream.SkipRecord();
                break;

            }
        }
    };
    CCBhvr.prototype.getChildren = function() {
        return [this.cTn, this.tgtEl, this.attrNameLst];
    };
    CCBhvr.prototype.onRemoveChild = function(oChild) {
        if(oChild === this.tgtEl) {
            if(this.parent) {
                this.parent.onRemoveChild(this);
            }
        }
    };

    changesFactory[AscDFH.historyitem_CTnChildTnLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_CTnEndCondLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_CTnEndSync] = CChangeObject;
    changesFactory[AscDFH.historyitem_CTnIterate] = CChangeObject;
    changesFactory[AscDFH.historyitem_CTnStCondLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_CTnSubTnLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_CTnAccel] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnAfterEffect] = CChangeBool;
    changesFactory[AscDFH.historyitem_CTnAutoRev] = CChangeBool;
    changesFactory[AscDFH.historyitem_CTnBldLvl] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnDecel] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnDisplay] = CChangeBool;
    changesFactory[AscDFH.historyitem_CTnDur] = CChangeString;
    changesFactory[AscDFH.historyitem_CTnEvtFilter] = CChangeString;
    changesFactory[AscDFH.historyitem_CTnFill] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnGrpId] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnId] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnMasterRel] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnNodePh] = CChangeBool;
    changesFactory[AscDFH.historyitem_CTnNodeType] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnPresetClass] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnPresetID] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnPresetSubtype] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnRepeatCount] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnRepeatDur] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnRestart] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnSpd] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnSyncBehavior] = CChangeLong;
    changesFactory[AscDFH.historyitem_CTnTmFilter] = CChangeString;

    drawingsChangesMap[AscDFH.historyitem_CTnChildTnLst] = function(oClass, value) {oClass.childTnLst = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnEndCondLst] = function(oClass, value) {oClass.endCondLst = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnEndSync] = function(oClass, value) {oClass.endSync = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnIterate] = function(oClass, value) {oClass.iterate = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnStCondLst] = function(oClass, value) {oClass.stCondLst = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnSubTnLst] = function(oClass, value) {oClass.subTnLst = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnAccel] = function(oClass, value) {oClass.accel = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnAfterEffect] = function(oClass, value) {oClass.afterEffect = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnAutoRev] = function(oClass, value) {oClass.autoRev = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnBldLvl] = function(oClass, value) {oClass.bldLvl = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnDecel] = function(oClass, value) {oClass.decel = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnDisplay] = function(oClass, value) {oClass.display = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnDur] = function(oClass, value) {oClass.dur = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnEvtFilter] = function(oClass, value) {oClass.evtFilter = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnFill] = function(oClass, value) {oClass.fill = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnGrpId] = function(oClass, value) {oClass.grpId = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnId] = function(oClass, value) {oClass.id = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnMasterRel] = function(oClass, value) {oClass.masterRel = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnNodePh] = function(oClass, value) {oClass.nodePh = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnNodeType] = function(oClass, value) {oClass.nodeType = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnPresetClass] = function(oClass, value) {oClass.presetClass = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnPresetID] = function(oClass, value) {oClass.presetID = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnPresetSubtype] = function(oClass, value) {oClass.presetSubtype = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnRepeatCount] = function(oClass, value) {oClass.repeatCount = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnRepeatDur] = function(oClass, value) {oClass.repeatDur = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnRestart] = function(oClass, value) {oClass.restart = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnSpd] = function(oClass, value) {oClass.spd = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnSyncBehavior] = function(oClass, value) {oClass.syncBehavior = value;};
    drawingsChangesMap[AscDFH.historyitem_CTnTmFilter] = function(oClass, value) {oClass.tmFilter = value;};
    function CCTn() {
        CBaseFormatObject.call(this);
        this.childTnLst = null;
        this.endCondLst = null;
        this.endSync = null;
        this.iterate = null;
        this.stCondLst = null;
        this.subTnLst = null;
        this.accel = null;
        this.afterEffect = null;
        this.autoRev = null;
        this.bldLvl = null;
        this.decel = null;
        this.display = null;
        this.dur = null;
        this.evtFilter = null;
        this.fill = null;
        this.grpId = null;
        this.id = null;
        this.masterRel = null;
        this.nodePh = null;
        this.nodeType = null;
        this.presetClass = null;
        this.presetID = null;
        this.presetSubtype = null;
        this.repeatCount = null;
        this.repeatDur = null;
        this.restart = null;
        this.spd = null;
        this.syncBehavior = null;
        this.tmFilter = null;
    }
    InitClass(CCTn, CBaseFormatObject, AscDFH.historyitem_type_CTn);
    CCTn.prototype.setChildTnLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CTnChildTnLst, this.childTnLst, pr));
        this.childTnLst = pr;
        this.setParentToChild(pr);
    };
    CCTn.prototype.setEndCondLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CTnEndCondLst, this.endCondLst, pr));
        this.endCondLst = pr;
        this.setParentToChild(pr);
    };
    CCTn.prototype.setEndSync = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CTnEndSync, this.endSync, pr));
        this.endSync = pr;
        this.setParentToChild(pr);
    };
    CCTn.prototype.setIterate = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CTnIterate, this.iterate, pr));
        this.iterate = pr;
        this.setParentToChild(pr);
    };
    CCTn.prototype.setStCondLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CTnStCondLst, this.stCondLst, pr));
        this.stCondLst = pr;
        this.setParentToChild(pr);
    };
    CCTn.prototype.setSubTnLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CTnSubTnLst, this.subTnLst, pr));
        this.subTnLst = pr;
        this.setParentToChild(pr);
    };
    CCTn.prototype.setAccel = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnAccel, this.accel, pr));
        this.accel = pr;
    };
    CCTn.prototype.setAfterEffect = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CTnAfterEffect, this.afterEffect, pr));
        this.afterEffect = pr;
    };
    CCTn.prototype.setAutoRev = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CTnAutoRev, this.autoRev, pr));
        this.autoRev = pr;
    };
    CCTn.prototype.setBldLvl = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnBldLvl, this.bldLvl, pr));
        this.bldLvl = pr;
    };
    CCTn.prototype.setDecel = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnDecel, this.decel, pr));
        this.decel = pr;
    };
    CCTn.prototype.setDisplay = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CTnDisplay, this.display, pr));
        this.display = pr;
    };
    CCTn.prototype.setDur = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CTnDur, this.dur, pr));
        this.dur = pr;
    };
    CCTn.prototype.setEvtFilter = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CTnEvtFilter, this.evtFilter, pr));
        this.evtFilter = pr;
    };
    CCTn.prototype.setFill = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnFill, this.fill, pr));
        this.fill = pr;
    };
    CCTn.prototype.setGrpId = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnGrpId, this.grpId, pr));
        this.grpId = pr;
    };
    CCTn.prototype.setId = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnId, this.id, pr));
        this.id = pr;
    };
    CCTn.prototype.setMasterRel = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnMasterRel, this.masterRel, pr));
        this.masterRel = pr;
    };
    CCTn.prototype.setNodePh = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CTnNodePh, this.nodePh, pr));
        this.nodePh = pr;
    };
    CCTn.prototype.setNodeType = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnNodeType, this.nodeType, pr));
        this.nodeType = pr;
    };
    CCTn.prototype.setPresetClass = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnPresetClass, this.presetClass, pr));
        this.presetClass = pr;
    };
    CCTn.prototype.setPresetID = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnPresetID, this.presetID, pr));
        this.presetID = pr;
    };
    CCTn.prototype.setPresetSubtype = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnPresetSubtype, this.presetSubtype, pr));
        this.presetSubtype = pr;
    };
    CCTn.prototype.setRepeatCount = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnRepeatCount, this.repeatCount, pr));
        this.repeatCount = pr;
    };
    CCTn.prototype.setRepeatDur = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnRepeatDur, this.repeatDur, pr));
        this.repeatDur = pr;
    };
    CCTn.prototype.setRestart = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnRestart, this.restart, pr));
        this.restart = pr;
    };
    CCTn.prototype.setSpd = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnSpd, this.spd, pr));
        this.spd = pr;
    };
    CCTn.prototype.setSyncBehavior = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CTnSyncBehavior, this.syncBehavior, pr));
        this.syncBehavior = pr;
    };
    CCTn.prototype.setTmFilter = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CTnTmFilter, this.tmFilter, pr));
        this.tmFilter = pr;
    };
    CCTn.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.childTnLst !== null) {
            oCopy.setChildTnLst(this.childTnLst.createDuplicate(oIdMap));
        }
        if(this.endCondLst !== null) {
            oCopy.setEndCondLst(this.endCondLst.createDuplicate(oIdMap));
        }
        if(this.endSync !== null) {
            oCopy.setEndSync(this.endSync.createDuplicate(oIdMap));
        }
        if(this.iterate !== null) {
            oCopy.setIterate(this.iterate.createDuplicate(oIdMap));
        }
        if(this.stCondLst !== null) {
            oCopy.setStCondLst(this.stCondLst.createDuplicate(oIdMap));
        }
        if(this.subTnLst !== null) {
            oCopy.setSubTnLst(this.subTnLst.createDuplicate(oIdMap));
        }
        if(this.accel !== null) {
            oCopy.setAccel(this.accel);
        }
        if(this.afterEffect !== null) {
            oCopy.setAfterEffect(this.afterEffect);
        }
        if(this.autoRev !== null) {
            oCopy.setAutoRev(this.autoRev);
        }
        if(this.bldLvl !== null) {
            oCopy.setBldLvl(this.bldLvl);
        }
        if(this.decel !== null) {
            oCopy.setDecel(this.decel);
        }
        if(this.display !== null) {
            oCopy.setDisplay(this.display);
        }
        if(this.dur !== null) {
            oCopy.setDur(this.dur);
        }
        if(this.evtFilter !== null) {
            oCopy.setEvtFilter(this.evtFilter);
        }
        if(this.fill !== null) {
            oCopy.setFill(this.fill);
        }
        if(this.grpId !== null) {
            oCopy.setGrpId(this.grpId);
        }
        if(this.id !== null) {
            oCopy.setId(this.id);
        }
        if(this.masterRel !== null) {
            oCopy.setMasterRel(this.masterRel);
        }
        if(this.nodePh !== null) {
            oCopy.setNodePh(this.nodePh);
        }
        if(this.nodeType !== null) {
            oCopy.setNodeType(this.nodeType);
        }
        if(this.presetClass !== null) {
            oCopy.setPresetClass(this.presetClass);
        }
        if(this.presetID !== null) {
            oCopy.setPresetID(this.presetID);
        }
        if(this.presetSubtype !== null) {
            oCopy.setPresetSubtype(this.presetSubtype);
        }
        if(this.repeatCount !== null) {
            oCopy.setRepeatCount(this.repeatCount);
        }
        if(this.repeatDur !== null) {
            oCopy.setRepeatDur(this.repeatDur);
        }
        if(this.restart !== null) {
            oCopy.setRestart(this.restart);
        }
        if(this.spd !== null) {
            oCopy.setSpd(this.spd);
        }
        if(this.syncBehavior !== null) {
            oCopy.setSyncBehavior(this.syncBehavior);
        }
        if(this.tmFilter !== null) {
            oCopy.setTmFilter(this.tmFilter);
        }
    };
    CCTn.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteInt2(0, this.accel);
        pWriter._WriteBool2(1, this.afterEffect);
        pWriter._WriteBool2(2, this.autoRev);
        pWriter._WriteLimit2(3, this.fill);
        pWriter._WriteLimit2(4, this.masterRel);
        pWriter._WriteLimit2(5, this.nodeType);
        pWriter._WriteLimit2(6, this.presetClass);
        pWriter._WriteLimit2(7, this.restart);
        pWriter._WriteLimit2(8, this.syncBehavior);
        pWriter._WriteBool2(9, this.display);
        pWriter._WriteBool2(10, this.nodePh);
        pWriter._WriteInt2(11, this.bldLvl);
        pWriter._WriteInt2(12, this.decel);
        pWriter._WriteInt2(13, this.bldLvl);
        pWriter._WriteInt2(14, this.grpId);
        pWriter._WriteInt2(15, this.id);
        pWriter._WriteInt2(16, this.presetID);
        pWriter._WriteInt2(17, this.presetSubtype);
        pWriter._WriteInt2(18, this.spd);
        pWriter._WriteString2(19, this.dur);
        pWriter._WriteString2(20, this.evtFilter);
        pWriter._WriteString2(21, this.repeatCount);
        pWriter._WriteString2(22, this.repeatDur);
        pWriter._WriteString2(23, this.tmFilter);
    };
    CCTn.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.stCondLst);
        this.writeRecord2(pWriter, 1, this.endCondLst);
        this.writeRecord2(pWriter, 2, this.endSync);
        this.writeRecord2(pWriter, 3, this.iterate);
        this.writeRecord2(pWriter, 4, this.childTnLst);
        this.writeRecord2(pWriter, 5, this.subTnLst);
    };
    CCTn.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
             if (0 === nType) this.setAccel(oStream.GetLong());
        else if (1 === nType) this.setAfterEffect(oStream.GetBool());
        else if (2 === nType) this.setAutoRev(oStream.GetBool());
        else if (3 === nType) this.setFill(oStream.GetUChar());
        else if (4 === nType) this.setMasterRel(oStream.GetUChar());
        else if (5 === nType) this.setNodeType(oStream.GetUChar());
        else if (6 === nType) this.setPresetClass(oStream.GetUChar());
        else if (7 === nType) this.setRestart(oStream.GetUChar());
        else if (8 === nType) this.setSyncBehavior(oStream.GetUChar());
        else if (9 === nType) this.setDisplay(oStream.GetBool());
        else if (10 === nType) this.setNodePh(oStream.GetBool());
        else if (11 === nType) this.setBldLvl(oStream.GetLong());
        else if (12 === nType) this.setDecel(oStream.GetLong());
        else if (13 === nType) this.setBldLvl(oStream.GetLong());
        else if (14 === nType) this.setGrpId(oStream.GetLong());
        else if (15 === nType) this.setId(oStream.GetLong());
        else if (16 === nType) this.setPresetID(oStream.GetLong());
        else if (17 === nType) this.setPresetSubtype(oStream.GetLong());
        else if (18 === nType) this.setSpd(oStream.GetLong());
        else if (19 === nType) this.setDur(oStream.GetString2());
        else if (20 === nType) this.setEvtFilter(oStream.GetString2());
        else if (21 === nType) this.setRepeatCount(oStream.GetString2());
        else if (22 === nType) this.setRepeatDur(oStream.GetString2());
        else if (23 === nType) this.setTmFilter(oStream.GetString2());
    };
    CCTn.prototype.readChild = function(nType, pReader) {
        switch (nType) {
            case 0: {
                this.setStCondLst(new CCondLst());
                this.stCondLst.fromPPTY(pReader);
                break;
            }
            case 1: {
                this.setEndCondLst(new CCondLst());
                this.endCondLst.fromPPTY(pReader);
                break;
            }
            case 2: {
                this.setEndSync(new CCond());
                this.endSync.fromPPTY(pReader);
                break;
            }
            case 3: {
                this.setIterate(new CIterateData());
                this.iterate.fromPPTY(pReader);
                break;
            }
            case 4: {
                this.setChildTnLst(new CChildTnLst());
                this.childTnLst.fromPPTY(pReader);
                break;
            }
            case 5: {
                this.setSubTnLst(new CTnLst());
                this.subTnLst.fromPPTY(pReader);
                break;
            }
            default: {
                pReader.stream.SkipRecord();
                break;
            }
        }
    };
    CCTn.prototype.getChildren = function() {
        return [this.stCondLst, this.endCondLst, this.endSync, this.iterate, this.childTnLst, this.subTnLst];
    };

    changesFactory[AscDFH.historyitem_CondRtn] = CChangeObject;
    changesFactory[AscDFH.historyitem_CondTgtEl] = CChangeObject;
    changesFactory[AscDFH.historyitem_CondTn] = CChangeObject;
    changesFactory[AscDFH.historyitem_CondDelay] = CChangeString;
    changesFactory[AscDFH.historyitem_CondEvt] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_CondRtn] = function(oClass, value) {oClass.rtn = value;};
    drawingsChangesMap[AscDFH.historyitem_CondTgtEl] = function(oClass, value) {oClass.tgtEl = value;};
    drawingsChangesMap[AscDFH.historyitem_CondTn] = function(oClass, value) {oClass.tn = value;};
    drawingsChangesMap[AscDFH.historyitem_CondDelay] = function(oClass, value) {oClass.delay = value;};
    drawingsChangesMap[AscDFH.historyitem_CondEvt] = function(oClass, value) {oClass.evt = value;};

    function CCond() {
        CBaseFormatObject.call(this);
        this.rtn = null;
        this.tgtEl = null;
        this.tn = null;
        this.delay = null;
        this.evt = null;
    }
    InitClass(CCond, CBaseFormatObject, AscDFH.historyitem_type_Cond);
    CCond.prototype.setRtn = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CondRtn, this.rtn, pr));
        this.rtn = pr;
        this.setParentToChild(pr);
    };
    CCond.prototype.setTgtEl = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CondTgtEl, this.tgtEl, pr));
        this.tgtEl = pr;
        this.setParentToChild(pr);
    };
    CCond.prototype.setTn = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CondTn, this.tn, pr));
        this.tn = pr;
    };
    CCond.prototype.setDelay = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CondDelay, this.delay, pr));
        this.delay = pr;
    };
    CCond.prototype.setEvt = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CondEvt, this.evt, pr));
        this.evt = pr;
    };
    CCond.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.rtn !== null) {
            oCopy.setRtn(this.rtn.createDuplicate(oIdMap));
        }
        if(this.tgtEl !== null) {
            oCopy.setTgtEl(this.tgtEl.createDuplicate(oIdMap));
        }
        if(this.tn !== null) {
            oCopy.setTn(this.tn.createDuplicate(oIdMap));
        }
        if(this.delay !== null) {
            oCopy.setDelay(this.delay);
        }
        if(this.evt !== null) {
            oCopy.setEvt(this.evt);
        }
    };
    CCond.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteInt2(0, this.tn);
        pWriter._WriteLimit2(1, this.rtn);
        pWriter._WriteLimit2(2, this.evt);
        pWriter._WriteString2(3, this.delay);
    };
    CCond.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.tgtEl);
    };
    CCond.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setTn(oStream.GetLong());
        else if (1 === nType) this.setRtn(oStream.GetUChar());
        else if (2 === nType) this.setEvt(oStream.GetUChar());
        else if (3 === nType) this.setDelay(oStream.GetString2());
    };
    CCond.prototype.readChild = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setTgtEl(new CTgtEl());
            this.tgtEl.fromPPTY(pReader);
        }
        else {
            oStream.SkipRecord();
        }
    };
    CCond.prototype.getChildren = function() {
        return [this.tgtEl];
    };

    changesFactory[AscDFH.historyitem_RtnVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_RtnVal] = function(oClass, value) {oClass.val = value;};
    function CRtn() {
        CBaseFormatObject.call(this);
        this.val = null;
    }
    InitClass(CRtn, CBaseFormatObject, AscDFH.historyitem_type_Rtn);
    CRtn.prototype.setVal = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_RtnVal, this.val, pr));
        this.val = pr;
    };
    CRtn.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.val !== null) {
            oCopy.setVal(this.val);
        }
    };
    CRtn.prototype.privateWriteAttributes = function(pWriter) {
    };
    CRtn.prototype.writeChildren = function(pWriter) {
    };
    CRtn.prototype.readAttribute = function(nType, pReader) {
    };
    CRtn.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_TgtElInkTgt] = CChangeObject;
    changesFactory[AscDFH.historyitem_TgtElSldTgt] = CChangeObject;
    changesFactory[AscDFH.historyitem_TgtElSndTgt] = CChangeObject;
    changesFactory[AscDFH.historyitem_TgtElSpTgt] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TgtElInkTgt] = function(oClass, value) {oClass.inkTgt = value;};
    drawingsChangesMap[AscDFH.historyitem_TgtElSldTgt] = function(oClass, value) {oClass.sldTgt = value;};
    drawingsChangesMap[AscDFH.historyitem_TgtElSndTgt] = function(oClass, value) {oClass.sndTgt = value;};
    drawingsChangesMap[AscDFH.historyitem_TgtElSpTgt] = function(oClass, value) {oClass.spTgt = value;};
    function CTgtEl() {
        CBaseFormatObject.call(this);
        this.inkTgt = null;//CObjectTarget
        this.sldTgt = null;
        this.sndTgt = null;
        this.spTgt = null;
    }
    InitClass(CTgtEl, CBaseFormatObject, AscDFH.historyitem_type_TgtEl);
    CTgtEl.prototype.setInkTgt = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TgtElInkTgt, this.inkTgt, pr));
        this.inkTgt = pr;
        this.setParentToChild(pr);
    };
    CTgtEl.prototype.setSldTgt = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TgtElSldTgt, this.sldTgt, pr));
        this.sldTgt = pr;
        this.setParentToChild(pr);
    };
    CTgtEl.prototype.setSndTgt = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TgtElSndTgt, this.sndTgt, pr));
        this.sndTgt = pr;
        this.setParentToChild(pr);
    };
    CTgtEl.prototype.setSpTgt = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TgtElSpTgt, this.spTgt, pr));
        this.spTgt = pr;
        this.setParentToChild(pr);
    };
    CTgtEl.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.inkTgt !== null) {
            oCopy.setInkTgt(this.inkTgt.createDuplicate(oIdMap));
        }
        if(this.sldTgt !== null) {
            oCopy.setSldTgt(this.sldTgt.createDuplicate(oIdMap));
        }
        if(this.sndTgt !== null) {
            oCopy.setSndTgt(this.sndTgt.createDuplicate(oIdMap));
        }
        if(this.spTgt !== null) {
            oCopy.setSpTgt(this.spTgt.createDuplicate(oIdMap));
        }
    };
    CTgtEl.prototype.privateWriteAttributes = function(pWriter) {
        if(this.inkTgt) {
            var nSpId = pWriter.GetSpIdxId(this.inkTgt.spid);
            if(nSpId !== null) {
                pWriter._WriteString2(0, nSpId + "");
            }
        }
        if(this.sndTgt) {
            pWriter._WriteString2(1, this.sndTgt.name);
            pWriter._WriteBool2(2, this.sndTgt.builtIn);
            pWriter._WriteString2(3, this.sndTgt.embed);
        }

    };
    CTgtEl.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.spTgt);
    };
    CTgtEl.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            if(!this.inkTgt) {
                this.setInkTgt(new CObjectTarget());
            }
            this.inkTgt.setSpid(oStream.GetString2(), pReader);
        }
        else if(1 === nType) {
            if(!this.sndTgt) {
                this.setSndTgt(new CSndTgt());
            }
            this.sndTgt.setName(oStream.GetString2());
        }
        else if(2 === nType) {
            if(!this.sndTgt) {
                this.setSndTgt(new CSndTgt());
            }
            this.sndTgt.setBuiltIn(oStream.GetBool());
        }
        else if(3 === nType) {
            if(!this.sndTgt) {
                this.setSndTgt(new CSndTgt());
            }
            this.sndTgt.setEmbed(oStream.GetString2());
        }

    };
    CTgtEl.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setSpTgt(new CSpTgt());
            this.spTgt.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CTgtEl.prototype.getChildren = function() {
        return [this.spTgt];
    };
    CTgtEl.prototype.onRemoveChild = function(oChild) {
        if(this.parent) {
            this.parent.onRemoveChild(this);
        }
    };


    changesFactory[AscDFH.historyitem_SndTgtEmbed] = CChangeLong;
    changesFactory[AscDFH.historyitem_SndTgtName] = CChangeString;
    changesFactory[AscDFH.historyitem_SndTgtBuiltIn] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_SndTgtEmbed] = function(oClass, value) {oClass.embed = value;};
    drawingsChangesMap[AscDFH.historyitem_SndTgtName] =  function(oClass, value) {oClass.name = value;};
    drawingsChangesMap[AscDFH.historyitem_SndTgtBuiltIn] =  function(oClass, value) {oClass.builtIn = value;};
    function CSndTgt() {//snd
        CBaseFormatObject.call(this);
        this.embed = null;
        this.name = null;
        this.builtIn = null;
    }
    InitClass(CSndTgt, CBaseFormatObject, AscDFH.historyitem_type_SndTgt);
    CSndTgt.prototype.setEmbed = function (pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_SndTgtEmbed, this.embed, pr));
        this.embed = pr;
    };
    CSndTgt.prototype.setName = function (pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_SndTgtName, this.name, pr));
        this.name = pr;
    };
    CSndTgt.prototype.setBuiltIn = function (pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_SndTgtBuiltIn, this.builtIn, pr));
        this.builtIn = pr;
    };
    CSndTgt.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.embed !== null) {
            oCopy.setEmbed(this.embed);
        }
        if(this.name !== null) {
            oCopy.setName(this.name);
        }
        if(this.builtIn !== null) {
            oCopy.setBuiltIn(this.builtIn);
        }
    };
    CSndTgt.prototype.privateWriteAttributes = function(pWriter) {
    };
    CSndTgt.prototype.writeChildren = function(pWriter) {
    };
    CSndTgt.prototype.readAttribute = function(nType, pReader) {
    };
    CSndTgt.prototype.readChild = function(nType, pReader) {
    };

    changesFactory[AscDFH.historyitem_SpTgtBg] = CChangeBool;
    changesFactory[AscDFH.historyitem_SpTgtGraphicEl] = CChangeObject;
    changesFactory[AscDFH.historyitem_SpTgtOleChartEl] = CChangeObject;
    changesFactory[AscDFH.historyitem_SpTgtSubSpId] = CChangeString;
    changesFactory[AscDFH.historyitem_SpTgtTxEl] = CChangeObject;

    drawingsChangesMap[AscDFH.historyitem_SpTgtBg] = function(oClass, value) {oClass.bg = value;};
    drawingsChangesMap[AscDFH.historyitem_SpTgtGraphicEl] = function(oClass, value) {oClass.graphicEl = value;};
    drawingsChangesMap[AscDFH.historyitem_SpTgtOleChartEl] = function(oClass, value) {oClass.oleChartEl = value;};
    drawingsChangesMap[AscDFH.historyitem_SpTgtSubSpId] = function(oClass, value) {oClass.subSpId = value;};
    drawingsChangesMap[AscDFH.historyitem_SpTgtTxEl] = function(oClass, value) {oClass.bg = value;};
    function CSpTgt() {
        CObjectTarget.call(this);
        this.bg = null;
        this.graphicEl = null;
        this.oleChartEl = null;
        this.subSpId = null;
        this.txEl = null;
    }
    InitClass(CSpTgt, CObjectTarget, AscDFH.historyitem_type_SpTgt);
    CSpTgt.prototype.setBg = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_SpTgtBg, this.bg, pr));
        this.bg = pr;
    };
    CSpTgt.prototype.setGraphicEl = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SpTgtGraphicEl, this.graphicEl, pr));
        this.graphicEl = pr;
        this.setParentToChild(pr);
    };
    CSpTgt.prototype.setOleChartEl = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SpTgtOleChartEl, this.oleChartEl, pr));
        this.oleChartEl = pr;
        this.setParentToChild(pr);
    };
    CSpTgt.prototype.setSubSpId = function(pr, pReader) {
        if(pReader) {
            pReader.AddConnectedObject(this);
        }
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_SpTgtSubSpId, this.subSpId, pr));
        this.subSpId = pr;
    };
    CSpTgt.prototype.assignConnection = function(oObjectsMap) {
        if(this.spid !== null) {
            if(AscCommon.isRealObject(oObjectsMap[this.spid])){
                this.setSpid(oObjectsMap[this.spid].Id);
            }
            else {
                if(this.parent) {
                    this.parent.onRemoveChild(this);
                }
            }
        }
        if(this.subSpId !== null) {
            if(AscCommon.isRealObject(oObjectsMap[this.subSpId])){
                this.setSubSpId(oObjectsMap[this.subSpId].Id);
            }
            else {
                if(this.parent) {
                    this.parent.onRemoveChild(this);
                }
            }
        }
    };
    CSpTgt.prototype.setTxEl = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SpTgtTxEl, this.txEl, pr));
        this.txEl = pr;
        this.setParentToChild(pr);
    };
    CSpTgt.prototype.fillObject = function(oCopy, oIdMap) {
        CObjectTarget.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.bg !== null) {
            oCopy.setBg(this.bg);
        }
        if(this.graphicEl !== null) {
            oCopy.setGraphicEl(this.graphicEl.createDuplicate(oIdMap));
        }
        if(this.oleChartEl !== null) {
            oCopy.setOleChartEl(this.oleChartEl.createDuplicate(oIdMap));
        }
        if(this.subSpId !== null) {
            var sId = this.subSpId;
            if(oIdMap && oIdMap[this.subSpId]) {
                sId = oIdMap[this.subSpId];
            }
            oCopy.setSubSpId(sId);
        }
        if(this.txEl !== null) {
            oCopy.setTxEl(this.txEl.createDuplicate(oIdMap));
        }
    };
    CSpTgt.prototype.privateWriteAttributes = function(pWriter) {
        var nSpId = pWriter.GetSpIdxId(this.spid);
        if(nSpId !== null) {
            pWriter._WriteString1(0, nSpId + "");
        }
        var spId = pWriter.GetSpIdxId(this.subSpId);
        if(spId !== null) {
            pWriter._WriteString2(1, spId + "");
        }
        pWriter._WriteBool2(2, this.bg);
        if(this.oleChartEl) {
            pWriter._WriteLimit2(3, this.oleChartEl.type);
            pWriter._WriteInt2(4, this.oleChartEl.lvl);
        }
    };
    CSpTgt.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.txEl);
        this.writeRecord2(pWriter, 1, this.graphicEl);
    };
    CSpTgt.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setSpid(oStream.GetString2(), pReader);
        else if (1 === nType) this.setSubSpId(oStream.GetString2());
        else if (2 === nType) this.setBg(oStream.GetBool());
        else if (3 === nType) {
            if(!this.oleChartEl) {
                this.setOleChartEl(new COleChartEl());
            }
            this.oleChartEl.setType(oStream.GetUChar());
        }
        else if (4 === nType) {
            if(!this.oleChartEl) {
                this.setOleChartEl(new COleChartEl());
            }
            this.oleChartEl.setLvl(oStream.GetLong());
        }
    };
    CSpTgt.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setTxEl(new CTxEl());
            this.txEl.fromPPTY(pReader);
        }
        else if(1 === nType) {
            this.setGraphicEl(new CGraphicEl());
            this.graphicEl.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CSpTgt.prototype.getChildren = function() {
        return [this.txEl, this.graphicEl];
    };
    CSpTgt.prototype.handleRemoveObject = function(sObjectId) {
        if(this.spid === sObjectId
        || this.subSpId === sObjectId) {
            if(this.parent) {
                this.parent.onRemoveChild(this);
            }
        }
    };

    changesFactory[AscDFH.historyitem_IterateDataTmAbs] = CChangeString;
    changesFactory[AscDFH.historyitem_IterateDataTmPct] = CChangeLong;
    changesFactory[AscDFH.historyitem_IterateDataBackwards] = CChangeBool;
    changesFactory[AscDFH.historyitem_IterateDataType] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_IterateDataTmAbs] = function(oClass, value) {oClass.tmAbs = value;};
    drawingsChangesMap[AscDFH.historyitem_IterateDataTmPct] = function(oClass, value) {oClass.tmPct = value;};
    drawingsChangesMap[AscDFH.historyitem_IterateDataBackwards] = function(oClass, value) {oClass.backwards = value;};
    drawingsChangesMap[AscDFH.historyitem_IterateDataType] = function(oClass, value) {oClass.type = value;};

    function CIterateData() {//iterate
        CBaseFormatObject.call(this);
        this.tmAbs = null;
        this.tmPct = null;
        this.backwards = null;
        this.type = null;
    }
    InitClass(CIterateData, CBaseFormatObject, AscDFH.historyitem_type_IterateData);
    CIterateData.prototype.setTmAbs = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_IterateDataTmAbs, this.tmAbs, pr));
        this.tmAbs = pr;
    };
    CIterateData.prototype.setTmPct = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_IterateDataTmPct, this.tmPct, pr));
        this.tmPct = pr;
    };
    CIterateData.prototype.setBackwards = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_IterateDataBackwards, this.backwards, pr));
        this.backwards = pr;
    };
    CIterateData.prototype.setType = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_IterateDataType, this.type, pr));
        this.type = pr;
    };
    CIterateData.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.tmAbs !== null) {
            oCopy.setTmAbs(this.tmAbs.createDuplicate(oIdMap));
        }
        if(this.tmPct !== null) {
            oCopy.setTmPct(this.tmPct.createDuplicate(oIdMap));
        }
        if(this.backwards !== null) {
            oCopy.setBackwards(this.backwards);
        }
        if(this.type !== null) {
            oCopy.setType(this.type);
        }
    };
    CIterateData.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.type);
        pWriter._WriteBool2(1, this.backwards);
        pWriter._WriteString2(2, this.tmAbs);
        pWriter._WriteInt2(3, this.tmPct);
    };
    CIterateData.prototype.writeChildren = function(pWriter) {
    };
    CIterateData.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setType(oStream.GetUChar());
        else if (1 === nType) this.setBackwards(oStream.GetBool());
        else if (2 === nType) this.setTmAbs(oStream.GetString2());
        else if (3 === nType) this.setTmPct(oStream.GetLong());
    };
    CIterateData.prototype.readChild = function(nType, pReader) {
        pReader.stream.SkipRecord();
    };


    changesFactory[AscDFH.historyitem_TmVal] = CChangeDouble2;
    drawingsChangesMap[AscDFH.historyitem_TmVal] = function(oClass, value) {oClass.val = value;};
    function CTm() {
        CBaseFormatObject.call(this);
        this.val = null
    }
    InitClass(CTm, CBaseFormatObject, AscDFH.historyitem_type_Tm);
    CTm.prototype.setVal = function(pr) {
        oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_TmVal, this.val, pr));
        this.val = pr;
    };
    CTm.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.val !== null) {
            oCopy.setVal(this.val);
        }
    };
    CTm.prototype.privateWriteAttributes = function(pWriter) {
    };
    CTm.prototype.writeChildren = function(pWriter) {
    };
    CTm.prototype.readAttribute = function(nType, pReader) {
    };
    CTm.prototype.readChild = function(nType, pReader) {
    };

    changesFactory[AscDFH.historyitem_TavVal] = CChangeObject;
    changesFactory[AscDFH.historyitem_TavFmla] = CChangeString;
    changesFactory[AscDFH.historyitem_TavTm] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_TavVal] = function(oClass, value) {oClass.val = value;};
    drawingsChangesMap[AscDFH.historyitem_TavFmla] = function(oClass, value) {oClass.fmla = value;};
    drawingsChangesMap[AscDFH.historyitem_TavTm] = function(oClass, value) {oClass.tm = value;};
    function CTav() {
        CBaseFormatObject.call(this);
        this.val = null;
        this.fmla = null;
        this.tm = null;
    }
    InitClass(CTav, CBaseFormatObject, AscDFH.historyitem_type_Tav);
    CTav.prototype.setVal = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TavVal, this.val, pr));
        this.val = pr;
        this.setParentToChild(pr);
    };
    CTav.prototype.setFmla = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_TavFmla, this.fmla, pr));
        this.fmla = pr;
    };
    CTav.prototype.setTm = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_TavTm, this.tm, pr));
        this.tm = pr;
    };
    CTav.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.val  !== null) {
            oCopy.setVal(this.val.createDuplicate(oIdMap));
        }
        if(this.fmla !== null) {
            oCopy.setFmla(this.fmla);
        }
        if(this.tm !== null) {
            oCopy.setTm(this.tm);
        }
    };
    CTav.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteString2(0, this.tm);
        pWriter._WriteString2(1, this.fmla);
    };
    CTav.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.val);
    };
    CTav.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setTm(oStream.GetString2());
        }
        else if(1 === nType) {
            this.setFmla(oStream.GetString2());
        }
    };
    CTav.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setVal(new CAnimVariant());
            this.val.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CTav.prototype.getChildren = function() {
        return [this.val];
    };


    changesFactory[AscDFH.historyitem_AnimVariantBoolVal] = CChangeBool;
    changesFactory[AscDFH.historyitem_AnimVariantClrVal] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_AnimVariantFltVal] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_AnimVariantIntVal] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimVariantStrVal] = CChangeString;

    drawingConstructorsMap[AscDFH.historyitem_AnimVariantClrVal] = AscFormat.CUniColor;

    drawingsChangesMap[AscDFH.historyitem_AnimVariantBoolVal] = function(oClass, value) {oClass.boolVal = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimVariantClrVal] = function(oClass, value) {oClass.clrVal = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimVariantFltVal] = function(oClass, value) {oClass.fltVal = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimVariantIntVal] = function(oClass, value) {oClass.intVal = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimVariantStrVal] = function(oClass, value) {oClass.strVal = value;};
    function CAnimVariant() {//progress, val
        CBaseFormatObject.call(this);
        this.boolVal = null;
        this.clrVal = null;
        this.fltVal = null;
        this.intVal = null;
        this.strVal = null;
    }
    InitClass(CAnimVariant, CBaseFormatObject, AscDFH.historyitem_type_AnimVariant);
    CAnimVariant.prototype.setBoolVal = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_AnimVariantBoolVal, this.boolVal, pr));
        this.boolVal = pr;
    };
    CAnimVariant.prototype.setClrVal = function(pr) {
        oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_AnimVariantClrVal, this.clrVal, pr));
        this.clrVal = pr;
    };
    CAnimVariant.prototype.setFltVal = function(pr) {
        oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_AnimVariantFltVal, this.fltVal, pr));
        this.fltVal = pr;
    };
    CAnimVariant.prototype.setIntVal = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimVariantIntVal, this.intVal, pr));
        this.intVal = pr;
    };
    CAnimVariant.prototype.setStrVal = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimVariantStrVal, this.strVal, pr));
        this.strVal = pr;
    };
    CAnimVariant.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.boolVal !== null) {
            oCopy.setBoolVal(this.boolVal);
        }
        if(this.clrVal !== null) {
            oCopy.setClrVal(this.clrVal.createDuplicate());
        }
        if(this.fltVal !== null) {
            oCopy.setFltVal(this.fltVal);
        }
        if(this.intVal !== null) {
            oCopy.setIntVal(this.intVal);
        }
        if(this.strVal !== null) {
            oCopy.setStrVal(this.strVal);
        }
    };
    CAnimVariant.prototype.privateWriteAttributes = function(pWriter) {
        if(null !== this.boolVal) {
            pWriter._WriteBool2(0, this.boolVal);
            return;
        }
        if(null !== this.strVal) {
            pWriter._WriteString2(1, this.strVal);
            return;
        }
        if(null !== this.intVal) {
            pWriter._WriteInt2(2, this.intVal);
            return;
        }
        if(null !== this.fltVal) {
            pWriter._WriteInt2(3, this.fltVal * 100000 + 0.5 >> 0);
            return;
        }
    };
    CAnimVariant.prototype.writeChildren = function(pWriter) {
        if(this.clrVal) {
            pWriter.WriteRecord1(0, this.clrVal, pWriter.WriteUniColor);
        }
    };
    CAnimVariant.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setBoolVal(oStream.GetBool());
        else if (1 === nType) this.setStrVal(oStream.GetString2());
        else if (2 === nType) this.setIntVal(oStream.GetLong());
        else if (3 === nType) this.setFltVal(oStream.GetLong() / 100000);
    };
    CAnimVariant.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setClrVal(pReader.ReadUniColor());
        }
    };

    changesFactory[AscDFH.historyitem_AnimClrByRGB] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_AnimClrByHSL] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_AnimClrCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimClrFrom] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimClrTo] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimClrClrSpc] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimClrDir] = CChangeLong;

    drawingConstructorsMap[AscDFH.historyitem_AnimClrByRGB] = CColorPercentage;
    drawingConstructorsMap[AscDFH.historyitem_AnimClrByHSL] = CColorPercentage;

    drawingsChangesMap[AscDFH.historyitem_AnimClrByRGB] = function(oClass, value) {oClass.byRGB = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimClrByHSL] = function(oClass, value) {oClass.byHSL = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimClrCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimClrFrom] = function(oClass, value) {oClass.from = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimClrTo] = function(oClass, value) {oClass.to = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimClrClrSpc] = function(oClass, value) {oClass.clrSpc = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimClrDir] = function(oClass, value) {oClass.dir = value;};

    function CColorPercentage() {
        this.c1 = 10000;
        this.c2 = 10000;
        this.c2 = 10000;
    }
    CColorPercentage.prototype.Write_ToBinary = function(oWriter) {
        oWriter.WriteLong(this.c1);
        oWriter.WriteLong(this.c2);
        oWriter.WriteLong(this.c3);
    };
    CColorPercentage.prototype.Read_FromBinary = function(oReader) {
        this.c1 = oReader.GetLong();
        this.c2 = oReader.GetLong();
        this.c2 = oReader.GetLong();
    };
    CColorPercentage.prototype.copy = function() {
        var oCopy = new CColorPercentage();
        oCopy.c1 = this.c1;
        oCopy.c2 = this.c2;
        oCopy.c3 = this.c3;
        return oCopy;
    };
    CColorPercentage.prototype.createDuplicate = function() {
        return this.copy();
    };
    function CAnimClr() {
        CBaseFormatObject.call(this);
        this.byRGB = null;
        this.byHSL = null;

        this.cBhvr = null;
        this.from = null;
        this.to = null;
        this.clrSpc = null;
        this.dir = null;
    }
    InitClass(CAnimClr, CBaseFormatObject, AscDFH.historyitem_type_AnimClr);
    CAnimClr.prototype.setByRGB = function(pr) {
        oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_AnimClrByRGB, this.byRGB, pr));
        this.byRGB = pr;
    };
    CAnimClr.prototype.setByHSL = function(pr) {
        oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_AnimClrByHSL, this.byHSL, pr));
        this.byHSL = pr;
    };
    CAnimClr.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimClrCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CAnimClr.prototype.setFrom = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimClrFrom, this.from, pr));
        this.from = pr;
        this.setParentToChild(pr);
    };
    CAnimClr.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimClrTo, this.to, pr));
        this.to = pr;
        this.setParentToChild(pr);
    };
    CAnimClr.prototype.setClrSpc = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimClrClrSpc, this.clrSpc, pr));
        this.clrSpc = pr;
    };
    CAnimClr.prototype.setDir = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimClrDir, this.dir, pr));
        this.dir = pr;
    };
    CAnimClr.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.byRGB !== null) {
            oCopy.setByRGB(this.byRGB.createDuplicate(oIdMap));
        }
        if(this.byHSL !== null) {
            oCopy.setByHSL(this.byHSL.createDuplicate(oIdMap));
        }
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.from !== null) {
            oCopy.setFrom(this.from.createDuplicate(oIdMap));
        }
        if(this.to !== null) {
            oCopy.setTo(this.to.createDuplicate(oIdMap));
        }
        if(this.clrSpc !== null) {
            oCopy.setClrSpc(this.clrSpc);
        }
        if(this.dir !== null) {
            oCopy.setDir(this.dir);
        }
    };
    CAnimClr.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.clrSpc);
        pWriter._WriteLimit2(1, this.dir);
        if(this.byRGB) {
            pWriter._WriteInt2(2, this.byRGB.c1);
            pWriter._WriteInt2(3, this.byRGB.c2);
            pWriter._WriteInt2(4, this.byRGB.c3);
        }
        if(this.byHSL) {
            pWriter._WriteInt2(5, this.byHSL.c1);
            pWriter._WriteInt2(6, this.byHSL.c2);
            pWriter._WriteInt2(7, this.byHSL.c3);
        }
    };
    CAnimClr.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);
        pWriter.WriteRecord1(1, this.from, pWriter.WriteUniColor);
        pWriter.WriteRecord1(2, this.to, pWriter.WriteUniColor);
    };
    CAnimClr.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) {
            this.setClrSpc(oStream.GetUChar());
        }
        else if (1 === nType) {
            this.setDir(oStream.GetUChar());
        }
        else if(2 === nType ||  3 === nType ||   4 === nType ||
                5 === nType || 6 === nType ||  7 === nType) {
            var oColor;
            if(2 === nType || 3 === nType || 4 === nType) {
                if(this.byRGB) {
                    oColor = this.byRGB.createDuplicate({});
                }
                else {
                    oColor = new CColorPercentage();
                }
                if (2 === nType) oColor.c1 = oStream.GetLong();
                else if (3 === nType) oColor.c2 = oStream.GetLong();
                else if (4 === nType) oColor.c3 = oStream.GetLong();
                this.setByRGB(oColor);
            }
            else {
                if(this.byHSL) {
                    oColor = this.byHSL.createDuplicate({});
                }
                else {
                    oColor = new CColorPercentage();
                }
                if(5 === nType) oColor.c1 = oStream.GetLong();
                else if (6 === nType) oColor.c2 = oStream.GetLong();
                else if (7 === nType) oColor.c3 = oStream.GetLong();
                this.setByHSL(oColor);
            }

        }
    };
    CAnimClr.prototype.readChild = function(nType, pReader) {
        var oStream = pReader.stream;
        switch (nType) {
            case 0: {
                this.setCBhvr(new CCBhvr());
                this.cBhvr.fromPPTY(pReader);
                break;
            }
            case 1: {
                this.setFrom(pReader.ReadUniColor());
                break;
            }
            case 2: {
                this.setTo(pReader.ReadUniColor());
                break;
            }
            default: {
                oStream.SkipRecord();
                break;
            }
        }
    };
    CAnimClr.prototype.getChildren = function() {
        return [this.cBhvr];
    };

    changesFactory[AscDFH.historyitem_AnimEffectCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimEffectProgress] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimEffectFilter] = CChangeString;
    changesFactory[AscDFH.historyitem_AnimEffectPrLst] = CChangeString;
    changesFactory[AscDFH.historyitem_AnimEffectTransition] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_AnimEffectCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimEffectProgress] = function(oClass, value) {oClass.progress = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimEffectFilter] = function(oClass, value) {oClass.filter = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimEffectPrLst] = function(oClass, value) {oClass.prLst = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimEffectTransition] = function(oClass, value) {oClass.transition = value;};
    function CAnimEffect() {
        CBaseFormatObject.call(this);
        this.cBhvr = null;
        this.progress = null;
        this.filter = null;
        this.prLst = null;
        this.transition = null;
    }
    InitClass(CAnimEffect, CBaseFormatObject, AscDFH.historyitem_type_AnimEffect);
    CAnimEffect.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimEffectCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CAnimEffect.prototype.setProgress = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimEffectProgress, this.progress, pr));
        this.progress = pr;
        this.setParentToChild(pr);
    };
    CAnimEffect.prototype.setFilter = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimEffectFilter, this.filter, pr));
        this.filter = pr;
    };
    CAnimEffect.prototype.setPrLst = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimEffectPrLst, this.prLst, pr));
        this.prLst = pr;
    };
    CAnimEffect.prototype.setTransition = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimEffectTransition, this.transition, pr));
        this.transition = pr;
    };
    CAnimEffect.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.progress !== null) {
            oCopy.setProgress(this.progress.createDuplicate(oIdMap));
        }
        if(this.filter !== null) {
            oCopy.setFilter(this.filter);
        }
        if(this.prLst !== null) {
            oCopy.setPrLst(this.prLst);
        }
        if(this.transition !== null) {
            oCopy.setTransition(this.transition);
        }
    };
    CAnimEffect.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0,  this.transition);
        pWriter._WriteString2(1, this.filter);
        pWriter._WriteString2(2, this.prLst);
    };
    CAnimEffect.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);
        this.writeRecord2(pWriter, 1, this.progress);
    };
    CAnimEffect.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 == nType)       this.setTransition(oStream.GetUChar());
        else if (1 == nType) this.setFilter(oStream.GetString2());
        else if (2 == nType) this.setPrLst(oStream.GetString2());
    };
    CAnimEffect.prototype.readChild = function(nType, pReader) {
        var s = pReader.stream;
        switch (nType) {
            case 0: {
                this.setCBhvr(new CCBhvr());
                this.cBhvr.fromPPTY(pReader);break;
            }
            case 1: {
                this.setProgress(new CAnimVariant());
                this.progress.fromPPTY(pReader);
                break;
            }
            default: {
                s.SkipRecord();
                break;
            }
        }
    };
    CAnimEffect.prototype.getChildren = function() {
        return [this.cBhvr, this.progress];
    };

    changesFactory[AscDFH.historyitem_AnimMotionBy] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimMotionCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimMotionFrom] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimMotionRCtr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimMotionTo] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimMotionOrigin] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimMotionPath] = CChangeString;
    changesFactory[AscDFH.historyitem_AnimMotionPathEditMode] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimMotionPtsTypes] = CChangeString;
    changesFactory[AscDFH.historyitem_AnimMotionRAng] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_AnimMotionBy] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionCBhvr] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionFrom] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionRCtr] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionTo] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionOrigin] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPath] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPathEditMode] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPtsTypes] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionRAng] = function(oClass, value) {oClass.by = value;};

    function CAnimMotion() {
        CBaseFormatObject.call(this);
        this.by = null;
        this.cBhvr = null;
        this.from = null;
        this.rCtr = null;
        this.to = null;
        this.origin = null;
        this.path = null;
        this.pathEditMode = null;
        this.ptsTypes = null;
        this.rAng = null;
    }
    InitClass(CAnimMotion, CBaseFormatObject, AscDFH.historyitem_type_AnimMotion);
    CAnimMotion.prototype.setBy = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimMotionBy, this.by, pr));
        this.by = pr;
        this.setParentToChild(pr);
    };
    CAnimMotion.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimMotionCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CAnimMotion.prototype.setFrom = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimMotionFrom, this.from, pr));
        this.from = pr;
        this.setParentToChild(pr);
    };
    CAnimMotion.prototype.setRCtr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimMotionRCtr, this.rCtr, pr));
        this.rCtr = pr;
        this.setParentToChild(pr);
    };
    CAnimMotion.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimMotionTo, this.to, pr));
        this.to = pr;
        this.setParentToChild(pr);
    };
    CAnimMotion.prototype.setOrigin = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimMotionOrigin, this.origin, pr));
        this.origin = pr;
    };
    CAnimMotion.prototype.setPath = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimMotionPath, this.path, pr));
        this.path = pr;
    };
    CAnimMotion.prototype.setPathEditMode = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimMotionPathEditMode, this.pathEditMode, pr));
        this.pathEditMode = pr;
    };
    CAnimMotion.prototype.setPtsTypes = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AnimMotionPtsTypes, this.ptsTypes, pr));
        this.ptsTypes = pr;
    };
    CAnimMotion.prototype.setRAng = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimMotionRAng, this.rAng, pr));
        this.rAng = pr;
    };
    CAnimMotion.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.by !== null) {
            oCopy.setBy(this.by.createDuplicate(oIdMap));
        }
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.from !== null) {
            oCopy.setFrom(this.from.createDuplicate(oIdMap));
        }
        if(this.rCtr !== null) {
            oCopy.setRCtr(this.rCtr.createDuplicate(oIdMap));
        }
        if(this.to !== null) {
            oCopy.setTo(this.to.createDuplicate(oIdMap));
        }
        if(this.origin !== null) {
            oCopy.setOrigin(this.origin);
        }
        if(this.path !== null) {
            oCopy.setPath(this.path);
        }
        if(this.pathEditMode !== null) {
            oCopy.setPathEditMode(this.pathEditMode);
        }
        if(this.ptsTypes !== null) {
            oCopy.setPtsTypes(this.ptsTypes);
        }
        if(this.rAng !== null) {
            oCopy.setRAng(this.rAng);
        }
    };
    CAnimMotion.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.origin);
        pWriter._WriteLimit2(1, this.pathEditMode);
        pWriter._WriteString2(2, this.path);
        pWriter._WriteString2(3, this.ptsTypes);
        if(this.by) {
            pWriter._WriteInt2(4, this.by.x);
            pWriter._WriteInt2(5, this.by.y);
        }
        if(this.from) {
            pWriter._WriteInt2(6, this.from.x);
            pWriter._WriteInt2(7, this.from.y);
        }
        if(this.to) {
            pWriter._WriteInt2(8, this.to.x);
            pWriter._WriteInt2(9, this.to.y);
        }
        if(this.rCtr) {
            pWriter._WriteInt2(10, this.rCtr.x);
            pWriter._WriteInt2(11, this.rCtr.y);
        }
        pWriter._WriteInt2(12, this.rAng);
    };
    CAnimMotion.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);
    };
    CAnimMotion.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setOrigin(oStream.GetUChar());
        else if (1 === nType) this.setPathEditMode(oStream.GetUChar());
        else if (2 === nType) this.setPath(oStream.GetString2());
        else if (3 === nType) this.setPtsTypes(oStream.GetString2());
        else if (4 === nType) {
            if(!this.by) {
                this.setBy(new CTLPoint());
            }
            this.by.setX(oStream.GetLong());
        }
        else if (5 === nType) {
            if(!this.by) {
                this.setBy(new CTLPoint());
            }
            this.by.setY(oStream.GetLong());
        }
        else if (6 === nType) {
            if(!this.from) {
                this.setFrom(new CTLPoint());
            }
            this.from.setX(oStream.GetLong());
        }
        else if (7 === nType) {
            if(!this.from) {
                this.setFrom(new CTLPoint());
            }
            this.from.setY(oStream.GetLong());
        }
        else if (8 === nType) {
            if(!this.to) {
                this.setTo(new CTLPoint());
            }
            this.to.setX(oStream.GetLong());
        }
        else if (9 === nType) {
            if(!this.to) {
                this.setTo(new CTLPoint());
            }
            this.to.setY(oStream.GetLong());
        }
        else if (10 === nType) {
            if(!this.rCtr) {
                this.setRCtr(new CTLPoint());
            }
            this.rCtr.setX(oStream.GetLong());
        }
        else if (11 === nType){
            if(!this.rCtr) {
                this.setRCtr(new CTLPoint());
            }
            this.rCtr.setY(oStream.GetLong());
        }
        else if (12 === nType) this.setRAng(oStream.GetLong());
    };
    CAnimMotion.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCBhvr(new CCBhvr());
            this.cBhvr.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CAnimMotion.prototype.getChildren = function() {
        return [this.cBhvr];
    };

    changesFactory[AscDFH.historyitem_AnimRotCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimRotBy] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimRotFrom] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimRotTo] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_AnimRotCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimRotBy] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimRotFrom] = function(oClass, value) {oClass.from = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimRotTo] = function(oClass, value) {oClass.to = value;};

    function CAnimRot() {
        CBaseFormatObject.call(this);
        this.cBhvr = null;
        this.by = null;
        this.from = null;
        this.to = null;
    }
    InitClass(CAnimRot, CBaseFormatObject, AscDFH.historyitem_type_AnimRot);
    CAnimRot.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimRotCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CAnimRot.prototype.setBy = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimRotBy, this.by, pr));
        this.by = pr;
    };
    CAnimRot.prototype.setFrom = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimRotFrom, this.from, pr));
        this.from = pr;
    };
    CAnimRot.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimRotTo, this.to, pr));
        this.to = pr;
    };
    CAnimRot.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.by !== null) {
            oCopy.setBy(this.by);
        }
        if(this.from !== null) {
            oCopy.setFrom(this.from);
        }
        if(this.to !== null) {
            oCopy.setTo(this.to);
        }
    };
    CAnimRot.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteInt2(0, this.by);
        pWriter._WriteInt2(1, this.from);
        pWriter._WriteInt2(2, this.to);
    };
    CAnimRot.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);

    };
    CAnimRot.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setBy(oStream.GetLong());
        }
        else if(1 === nType) {
            this.setFrom(oStream.GetLong());
        }
        else if(2 === nType) {
            this.setTo(oStream.GetLong());
        }
    };
    CAnimRot.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCBhvr(new CCBhvr());
            this.cBhvr.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CAnimRot.prototype.getChildren = function() {
        return [this.cBhvr];
    };


    changesFactory[AscDFH.historyitem_AnimScaleCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimScaleBy] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimScaleFrom] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimScaleTo] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimScaleZoomContents] = CChangeBool;

    drawingsChangesMap[AscDFH.historyitem_AnimScaleCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimScaleBy] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimScaleFrom] = function(oClass, value) {oClass.from = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimScaleTo] = function(oClass, value) {oClass.to = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimScaleZoomContents] = function(oClass, value) {oClass.zoomContents = value;};
    function CAnimScale() {
        CBaseFormatObject.call(this);
        this.cBhvr = null;
        this.by = null;
        this.from = null;
        this.to = null;
        this.zoomContents = null;
    }
    InitClass(CAnimScale, CBaseFormatObject, AscDFH.historyitem_type_AnimScale);
    CAnimScale.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimScaleCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CAnimScale.prototype.setBy = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimScaleBy, this.by, pr));
        this.by = pr;
        this.setParentToChild(pr);
    };
    CAnimScale.prototype.setFrom = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimScaleFrom, this.from, pr));
        this.from = pr;
        this.setParentToChild(pr);
    };
    CAnimScale.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AnimScaleTo, this.to, pr));
        this.to = pr;
        this.setParentToChild(pr);
    };
    CAnimScale.prototype.setZoomContents = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_AnimScaleZoomContents, this.zoomContents, pr));
        this.zoomContents = pr;
    };
    CAnimScale.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.by !== null) {
            oCopy.setBy(this.by.createDuplicate(oIdMap));
        }
        if(this.from !== null) {
            oCopy.setFrom(this.from.createDuplicate(oIdMap));
        }
        if(this.to !== null) {
            oCopy.setTo(this.to.createDuplicate(oIdMap));
        }
        if(this.zoomContents !== null) {
            oCopy.setZoomContents(this.zoomContents);
        }
    };
    CAnimScale.prototype.privateWriteAttributes = function(pWriter) {
        if(this.by) {
            pWriter._WriteInt2(0, this.by.x);
            pWriter._WriteInt2(1, this.by.y);
        }
        if(this.from) {
            pWriter._WriteInt2(2, this.from.x);
            pWriter._WriteInt2(3, this.from.y);
        }
        if(this.to) {
            pWriter._WriteInt2(4, this.to.x);
            pWriter._WriteInt2(5, this.to.y);
        }
        pWriter._WriteBool2(6, this.zoomContents);
    };
    CAnimScale.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);
    };
    CAnimScale.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            if(!this.by) {
                this.setBy(new CTLPoint());
            }
            this.by.setX(oStream.GetLong());
        }
        else if(1 === nType) {
            if(!this.by) {
                this.setBy(new CTLPoint());
            }
            this.by.setY(oStream.GetLong());
        }
        else if(2 === nType) {
            if(!this.from) {
                this.setFrom(new CTLPoint());
            }
            this.from.setX(oStream.GetLong());
        }
        else if(3 === nType) {
            if(!this.from) {
                this.setFrom(new CTLPoint());
            }
            this.from.setY(oStream.GetLong());
        }
        else if(4 === nType) {
            if(!this.to) {
                this.setTo(new CTLPoint());
            }
            this.to.setX(oStream.GetLong());
        }
        else if(5 === nType) {
            if(!this.to) {
                this.setTo(new CTLPoint());
            }
            this.to.setY(oStream.GetLong());
        }
        else if(6 === nType) {
            this.setZoomContents(oStream.GetBool());
        }
    };
    CAnimScale.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCBhvr(new CCBhvr());
            this.cBhvr.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CAnimScale.prototype.getChildren = function() {
        return [this.cBhvr];
    };

    changesFactory[AscDFH.historyitem_AudioCMediaNode] = CChangeObject;
    changesFactory[AscDFH.historyitem_AudioIsNarration] = CChangeBool;

    drawingsChangesMap[AscDFH.historyitem_AudioCMediaNode] = function(oClass, value) {oClass.cMediaNode = value;};
    drawingsChangesMap[AscDFH.historyitem_AudioIsNarration] = function(oClass, value) {oClass.isNarration = value;};

    function CAudio() {
        CBaseFormatObject.call(this);
        this.cMediaNode = null;
        this.isNarration = null;
    }
    InitClass(CAudio, CBaseFormatObject, AscDFH.historyitem_type_Audio);

    CAudio.prototype.setCMediaNode = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AudioCMediaNode, this.cMediaNode, pr));
        this.cMediaNode = pr;
        this.setParentToChild(pr);
    };
    CAudio.prototype.setIsNarration = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_AudioIsNarration, this.isNarration, pr));
        this.isNarration = pr;
    };
    CAudio.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cMediaNode !== null) {
            oCopy.setCMediaNode(this.cMediaNode.createDuplicate(oIdMap));
        }
        if(this.isNarration !== null) {
            oCopy.setIsNarration(this.isNarration);
        }
    };
    CAudio.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteBool2(0, this.isNarration);
    };
    CAudio.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cMediaNode);
    };
    CAudio.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setIsNarration(oStream.GetBool());
        }
    };
    CAudio.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCMediaNode(new CCMediaNode());
            this.cMediaNode.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CAudio.prototype.getChildren = function() {
        return [this.cMediaNode];
    };


    changesFactory[AscDFH.historyitem_CMediaNodeCTn] = CChangeObject;
    changesFactory[AscDFH.historyitem_CMediaNodeTgtEl] = CChangeObject;
    changesFactory[AscDFH.historyitem_CMediaNodeMute] = CChangeBool;
    changesFactory[AscDFH.historyitem_CMediaNodeNumSld] = CChangeLong;
    changesFactory[AscDFH.historyitem_CMediaNodeShowWhenStopped] = CChangeBool;
    changesFactory[AscDFH.historyitem_CMediaNodeVol] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_CMediaNodeCTn] = function(oClass, value) {oClass.cTn = value;};
    drawingsChangesMap[AscDFH.historyitem_CMediaNodeTgtEl] = function(oClass, value) {oClass.tgtEl = value;};
    drawingsChangesMap[AscDFH.historyitem_CMediaNodeMute] = function(oClass, value) {oClass.mute = value;};
    drawingsChangesMap[AscDFH.historyitem_CMediaNodeNumSld] = function(oClass, value) {oClass.numSld = value;};
    drawingsChangesMap[AscDFH.historyitem_CMediaNodeShowWhenStopped] = function(oClass, value) {oClass.showWhenStopped = value;};
    drawingsChangesMap[AscDFH.historyitem_CMediaNodeVol] = function(oClass, value) {oClass.vol = value;};

    function CCMediaNode() {
        CBaseFormatObject.call(this);
        this.cTn = null;
        this.tgtEl = null;
        this.mute = null;
        this.numSld = null;
        this.showWhenStopped = null;
        this.vol = null;
    }
    InitClass(CCMediaNode, CBaseFormatObject, AscDFH.historyitem_type_CMediaNode);
    CCMediaNode.prototype.setCTn = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CMediaNodeCTn, this.cTn, pr));
        this.cTn = pr;
        this.setParentToChild(pr);
    };
    CCMediaNode.prototype.setTgtEl = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CMediaNodeTgtEl, this.tgtEl, pr));
        this.tgtEl = pr;
        this.setParentToChild(pr);
    };
    CCMediaNode.prototype.setMute = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CMediaNodeMute, this.mute, pr));
        this.mute = pr;
    };
    CCMediaNode.prototype.setNumSld = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CMediaNodeNumSld, this.numSld, pr));
        this.numSld = pr;
    };
    CCMediaNode.prototype.setShowWhenStopped = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_CMediaNodeShowWhenStopped, this.showWhenStopped, pr));
        this.showWhenStopped = pr;
    };
    CCMediaNode.prototype.setVol = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CMediaNodeVol, this.vol, pr));
        this.vol = pr;
    };
    CCMediaNode.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cTn !== null) {
            oCopy.setCTn(this.cTn.createDuplicate(oIdMap));
        }
        if(this.tgtEl !== null) {
            oCopy.setTgtEl(this.tgtEl.createDuplicate(oIdMap));
        }
        if(this.mute !== null) {
            oCopy.setMute(this.mute);
        }
        if(this.numSld !== null) {
            oCopy.setNumSld(this.numSld);
        }
        if(this.showWhenStopped !== null) {
            oCopy.setShowWhenStopped(this.showWhenStopped);
        }
        if(this.vol !== null) {
            oCopy.setVol(this.vol);
        }
    };
    CCMediaNode.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteInt2(0, this.numSld);
        pWriter._WriteInt2(1, this.vol);
        pWriter._WriteBool2(2, this.mute);
        pWriter._WriteBool2(3, this.showWhenStopped);
    };
    CCMediaNode.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cTn);
        this.writeRecord1(pWriter, 1, this.tgtEl);
    };
    CCMediaNode.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setNumSld(oStream.GetLong());
        else if (1 === nType) this.setVol(oStream.GetLong());
        else if (2 === nType) this.setMute(oStream.GetBool());
        else if (3 === nType) this.setShowWhenStopped(oStream.GetBool());
    };
    CCMediaNode.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCTn(new CCTn());
            this.cTn.fromPPTY(pReader);
        }
        else if(1 === nType) {
            this.setTgtEl(new CTgtEl());
            this.tgtEl.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CCMediaNode.prototype.getChildren = function() {
        return [this.cTn, this.tgtEl];
    };

    changesFactory[AscDFH.historyitem_CmdCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_CmdCmd] = CChangeString;
    changesFactory[AscDFH.historyitem_CmdType] = CChangeLong;

    drawingsChangesMap[AscDFH.historyitem_CmdCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_CmdCmd] = function(oClass, value) {oClass.cmd = value;};
    drawingsChangesMap[AscDFH.historyitem_CmdType] = function(oClass, value) {oClass.type = value;};
    function CCmd() {
        CBaseFormatObject.call(this);
        this.cBhvr = null;
        this.cmd = null;
        this.type = null;
    }
    InitClass(CCmd, CBaseFormatObject, AscDFH.historyitem_type_Cmd);
    CCmd.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CmdCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CCmd.prototype.setCmd = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CmdCmd, this.cmd, pr));
        this.cmd = pr;
    };
    CCmd.prototype.setType = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CmdType, this.type, pr));
        this.type = pr;
    };
    CCmd.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.cmd !== null) {
            oCopy.setCmd(this.cmd);
        }
        if(this.type !== null) {
            oCopy.setType(this.type);
        }
    };
    CCmd.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteLimit2(0, this.type);
        pWriter._WriteString2(1, this.cmd);
    };
    CCmd.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);
    };
    CCmd.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setType(oStream.GetUChar());
        }
        else if(1 === nType) {
            this.setCmd(oStream.GetString2());
        }
    };
    CCmd.prototype.readChild = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setCBhvr(new CCBhvr());
            this.cBhvr.fromPPTY(pReader);
        }
        else {
            oStream.SkipRecord();
        }
    };
    CCmd.prototype.getChildren = function() {
        return [this.cBhvr];
    };

    changesFactory[AscDFH.historyitem_TimeNodeContainerCTn] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TimeNodeContainerCTn] = function(oClass, value) {oClass.cTn = value;};

    function CTimeNodeContainer() {//par, excl
        CBaseFormatObject.call(this);
        this.cTn = null;
    }
    InitClass(CTimeNodeContainer, CBaseFormatObject, AscDFH.historyitem_type_TimeNodeContainer);
    CTimeNodeContainer.prototype.setCTn = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TimeNodeContainerCTn, this.cTn, pr));
        this.cTn = pr;
        this.setParentToChild(pr);
    };
    CTimeNodeContainer.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cTn !== null) {
            oCopy.setCTn(this.cTn.createDuplicate(oIdMap));
        }
    };
    CTimeNodeContainer.prototype.privateWriteAttributes = function(pWriter) {
    };
    CTimeNodeContainer.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cTn);
    };
    CTimeNodeContainer.prototype.readAttribute = function(nType, pReader) {
    };
    CTimeNodeContainer.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCTn(new CCTn());
            this.cTn.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CTimeNodeContainer.prototype.getChildren = function() {
        return [this.cTn];
    };

    function CPar() {//par, excl
        CBaseFormatObject.call(this);
        this.cTn = null;
    }
    InitClass(CPar, CTimeNodeContainer, AscDFH.historyitem_type_Par);

    function CExcl() {//par, excl
        CBaseFormatObject.call(this);
        this.cTn = null;
    }
    InitClass(CExcl, CTimeNodeContainer, AscDFH.historyitem_type_Excl);



    changesFactory[AscDFH.historyitem_SeqNextCondLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_SeqPrevCondLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_SeqConcurrent] = CChangeBool;
    changesFactory[AscDFH.historyitem_SeqNextAc] = CChangeLong;
    changesFactory[AscDFH.historyitem_SeqPrevAc] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_SeqNextCondLst] = function(oClass, value) {oClass.nextCondLst = value;};
    drawingsChangesMap[AscDFH.historyitem_SeqPrevCondLst] = function(oClass, value) {oClass.prevCondLst = value;};
    drawingsChangesMap[AscDFH.historyitem_SeqConcurrent] = function(oClass, value) {oClass.concurrent = value;};
    drawingsChangesMap[AscDFH.historyitem_SeqNextAc] = function(oClass, value) {oClass.nextAc = value;};
    drawingsChangesMap[AscDFH.historyitem_SeqPrevAc] = function(oClass, value) {oClass.prevAc = value;};
    function CSeq() {//
        CTimeNodeContainer.call(this);
        this.nextCondLst = null;
        this.prevCondLst = null;
        this.concurrent = null;
        this.nextAc = null;
        this.prevAc = null;
    }
    InitClass(CSeq, CTimeNodeContainer, AscDFH.historyitem_type_Seq);
    CSeq.prototype.setNextCondLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SeqNextCondLst, this.nextCondLst, pr));
        this.nextCondLst = pr;
        this.setParentToChild(pr);
    };
    CSeq.prototype.setPrevCondLst = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SeqPrevCondLst, this.prevCondLst, pr));
        this.prevCondLst = pr;
        this.setParentToChild(pr);
    };
    CSeq.prototype.setConcurrent = function(pr) {
        oHistory.Add(new CChangeBool(this, AscDFH.historyitem_SeqConcurrent, this.concurrent, pr));
        this.concurrent = pr;
    };
    CSeq.prototype.setNextAc = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_SeqNextAc, this.nextAc, pr));
        this.nextAc = pr;
    };
    CSeq.prototype.setPrevAc = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_SeqPrevAc, this.prevAc, pr));
        this.prevAc = pr;
    };
    CSeq.prototype.fillObject = function(oCopy, oIdMap) {
        CTimeNodeContainer.prototype.fillObject.call(this, oCopy, oIdMap);
        if(this.nextCondLst !== null) {
            oCopy.setNextCondLst(this.nextCondLst.createDuplicate(oIdMap));
        }
        if(this.prevCondLst !== null) {
            oCopy.setPrevCondLst(this.prevCondLst.createDuplicate(oIdMap));
        }
        if(this.concurrent !== null) {
            oCopy.setConcurrent(this.concurrent);
        }
        if(this.nextAc !== null) {
            oCopy.setNextAc(this.nextAc);
        }
        if(this.prevAc !== null) {
            oCopy.setPrevAc(this.prevAc);
        }
    };
    CSeq.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteBool2(0, this.concurrent);
        pWriter._WriteLimit2(1, this.nextAc);
        pWriter._WriteLimit2(2, this.prevAc);
    };
    CSeq.prototype.writeChildren = function(pWriter) {
        this.writeRecord2(pWriter, 0, this.prevCondLst);
        this.writeRecord2(pWriter, 1, this.nextCondLst);
        this.writeRecord1(pWriter, 2, this.cTn);
    };
    CSeq.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if (0 === nType) this.setConcurrent(oStream.GetBool());
        else if (1 === nType) this.setNextAc(oStream.GetUChar());
        else if (2 === nType) this.setPrevAc(oStream.GetUChar());
    };
    CSeq.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setPrevCondLst(new CCondLst());
            this.prevCondLst.fromPPTY(pReader);
        }
        else if(1 === nType) {
            this.setNextCondLst(new CCondLst());
            this.nextCondLst.fromPPTY(pReader);
        }
        else if(2 === nType) {
            this.setCTn(new CCTn());
            this.cTn.fromPPTY(pReader);
        }
    };
    CSeq.prototype.getChildren = function() {
        return [this.prevCondLst, this.nextCondLst, this.cTn];
    };


    changesFactory[AscDFH.historyitem_SetCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_SetTo] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_SetCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_SetTo] = function(oClass, value) {oClass.to = value;};
    function CSet() {//par, excl
        CBaseFormatObject.call(this);
        this.cBhvr = null;
        this.to = null;
    }
    InitClass(CSet, CBaseFormatObject, AscDFH.historyitem_type_Set);
    CSet.prototype.setCBhvr = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SetCBhvr, this.cBhvr, pr));
        this.cBhvr = pr;
        this.setParentToChild(pr);
    };
    CSet.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SetTo, this.cBhvr, pr));
        this.to = pr;
        this.setParentToChild(pr);
    };
    CSet.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cBhvr !== null) {
            oCopy.setCBhvr(this.cBhvr.createDuplicate(oIdMap));
        }
        if(this.to !== null) {
            oCopy.setTo(this.to.createDuplicate(oIdMap));
        }
    };
    CSet.prototype.privateWriteAttributes = function(pWriter) {
    };
    CSet.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cBhvr);
        this.writeRecord2(pWriter, 1, this.to);
    };
    CSet.prototype.readAttribute = function(nType, pReader) {
    };
    CSet.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCBhvr(new CCBhvr());
            this.cBhvr.fromPPTY(pReader);
        }
        else if(1 === nType) {
            this.setTo(new CAnimVariant());
            this.to.fromPPTY(pReader);
        }
    };
    CSet.prototype.getChildren = function() {
        return [this.cBhvr, this.to];
    };

    changesFactory[AscDFH.historyitem_VideoCMediaNode] = CChangeObject;
    changesFactory[AscDFH.historyitem_VideoFullScrn] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_VideoCMediaNode] = function(oClass, value) {oClass.cMediaNode = value;};
    drawingsChangesMap[AscDFH.historyitem_VideoFullScrn] = function(oClass, value) {oClass.fullScrn = value;};
    function CVideo() {//par, excl
        CBaseFormatObject.call(this);
        this.cMediaNode = null;
        this.fullScrn = null;
    }
    InitClass(CVideo, CBaseFormatObject, AscDFH.historyitem_type_Video);
    CVideo.prototype.setCMediaNode = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VideoCMediaNode, this.cMediaNode, pr));
        this.cMediaNode = pr;
        this.setParentToChild(pr);
    };
    CVideo.prototype.setFullScrn = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VideoFullScrn, this.fullScrn, pr));
        this.fullScrn = pr;
        this.setParentToChild(pr);
    };
    CVideo.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cMediaNode !== null) {
            oCopy.setCMediaNode(this.cMediaNode.createDuplicate(oIdMap));
        }
        if(this.fullScrn !== null) {
            oCopy.setFullScrn(this.fullScrn);
        }
    };
    CVideo.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteBool2(0, this.fullScrn);
    };
    CVideo.prototype.writeChildren = function(pWriter) {
        this.writeRecord1(pWriter, 0, this.cMediaNode);
    };
    CVideo.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(0 === nType) {
            this.setFullScrn(oStream.GetBool());
        }
    };
    CVideo.prototype.readChild = function(nType, pReader) {
        if(0 === nType) {
            this.setCMediaNode(new CCMediaNode());
            this.cMediaNode.fromPPTY(pReader);
        }
        else {
            pReader.stream.SkipRecord();
        }
    };
    CVideo.prototype.getChildren = function() {
        return [this.cMediaNode];
    };


    changesFactory[AscDFH.historyitem_OleChartElLvl] = CChangeLong;
    changesFactory[AscDFH.historyitem_OleChartElType] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_OleChartElLvl] = function(oClass, value) {oClass.lvl = value;};
    drawingsChangesMap[AscDFH.historyitem_OleChartElType] = function(oClass, value) {oClass.type = value;};
    function COleChartEl() {
        CBaseFormatObject.call(this);
        this.lvl = null;
        this.type = null;
    }
    InitClass(COleChartEl, CBaseFormatObject, AscDFH.historyitem_type_OleChartEl);
    COleChartEl.prototype.setLvl = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_OleChartElLvl, this.lvl, pr));
        this.lvl = pr;
    };
    COleChartEl.prototype.setType = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_OleChartElType, this.type, pr));
        this.type = pr;
    };
    COleChartEl.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.lvl !== null) {
            oCopy.setLvl(this.lvl);
        }
        if(this.type !== null) {
            oCopy.setType(this.type);
        }
    };
    COleChartEl.prototype.privateWriteAttributes = function(pWriter) {
    };
    COleChartEl.prototype.writeChildren = function(pWriter) {
    };
    COleChartEl.prototype.readAttribute = function(nType, pReader) {
    };
    COleChartEl.prototype.readChild = function(nType, pReader) {
    };

    changesFactory[AscDFH.historyitem_TlPointX] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_TlPointY] = CChangeDouble2;
    drawingsChangesMap[AscDFH.historyitem_TlPointX] = function(oClass, value) {oClass.x = value;};
    drawingsChangesMap[AscDFH.historyitem_TlPointY] = function(oClass, value) {oClass.y = value;};
    function CTLPoint() {//rCtr
        CBaseFormatObject.call(this);
        this.x = null;
        this.y = null;
    }
    InitClass(CTLPoint, CBaseFormatObject, AscDFH.historyitem_type_TlPoint);
    CTLPoint.prototype.setX = function(pr) {
        oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_TlPointX, this.x, pr));
        this.x = pr;
    };
    CTLPoint.prototype.setY = function(pr) {
        oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_TlPointY, this.y, pr));
        this.y = pr;
    };
    CTLPoint.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.x !== null) {
            oCopy.setX(this.x);
        }
        if(this.y !== null) {
            oCopy.setY(this.y);
        }
    };
    CTLPoint.prototype.privateWriteAttributes = function(pWriter) {
    };
    CTLPoint.prototype.writeChildren = function(pWriter) {
    };
    CTLPoint.prototype.readAttribute = function(nType, pReader) {
    };
    CTLPoint.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_SndAcEndSnd] = CChangeObject;
    changesFactory[AscDFH.historyitem_SndAcStSnd] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_SndAcEndSnd] = function(oClass, value) {oClass.endSnd = value;};
    drawingsChangesMap[AscDFH.historyitem_SndAcStSnd] = function(oClass, value) {oClass.stSnd = value;};
    function CSndAc() {//
        CBaseFormatObject.call(this);
        this.endSnd = null;
        this.stSnd = null;
    }
    InitClass(CSndAc, CBaseFormatObject, AscDFH.historyitem_type_SndAc);
    CSndAc.prototype.setEndSnd = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_SndAcEndSnd, this.endSnd, pr));
        this.endSnd = pr;
    };
    CSndAc.prototype.setStSnd = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_SndAcStSnd, this.stSnd, pr));
        this.stSnd = pr;
    };
    CSndAc.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.endSnd !== null) {
            oCopy.setEndSnd(this.endSnd.createDuplicate(oIdMap));
        }
        if(this.stSnd !== null) {
            oCopy.setStSnd(this.stSnd.createDuplicate(oIdMap));
        }
    };
    CSndAc.prototype.privateWriteAttributes = function(pWriter) {
    };
    CSndAc.prototype.writeChildren = function(pWriter) {
    };
    CSndAc.prototype.readAttribute = function(nType, pReader) {
    };
    CSndAc.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_StSndSnd] = CChangeObject;
    changesFactory[AscDFH.historyitem_StSndLoop] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_StSndSnd] = function(oClass, value) {oClass.snd = value;};
    drawingsChangesMap[AscDFH.historyitem_StSndLoop] = function(oClass, value) {oClass.loop = value;};
    function CStSnd() {//
        CBaseFormatObject.call(this);
        this.snd = null;
        this.loop = null;
    }
    InitClass(CStSnd, CBaseFormatObject, AscDFH.historyitem_type_StSnd);
    CStSnd.prototype.setSnd = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_StSndSnd, this.snd, pr));
        this.snd = pr;
    };
    CStSnd.prototype.setLoop = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_StSndLoop, this.loop, pr));
        this.loop = pr;
    };
    CStSnd.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.snd !== null) {
            oCopy.setSnd(this.snd.createDuplicate(oIdMap));
        }
        if(this.loop !== null) {
            oCopy.setLoop(this.loop);
        }
    };
    CStSnd.prototype.privateWriteAttributes = function(pWriter) {
    };
    CStSnd.prototype.writeChildren = function(pWriter) {
    };
    CStSnd.prototype.readAttribute = function(nType, pReader) {
    };
    CStSnd.prototype.readChild = function(nType, pReader) {
    };



    changesFactory[AscDFH.historyitem_TxElCharRg] = CChangeObject;
    changesFactory[AscDFH.historyitem_TxElPRg] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TxElCharRg] = function(oClass, value) {oClass.charRg = value;};
    drawingsChangesMap[AscDFH.historyitem_TxElPRg] = function(oClass, value) {oClass.pRg = value;};
    function CTxEl() {//rCtr
        CBaseFormatObject.call(this);
        this.charRg = null;//CIndexRg
        this.pRg = null;
    }
    InitClass(CTxEl, CBaseFormatObject, AscDFH.historyitem_type_TxEl);
    CTxEl.prototype.setCharRg = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TxElCharRg, this.charRg, pr));
        this.charRg = pr;
        this.setParentToChild(pr);
    };
    CTxEl.prototype.setPRg = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TxElPRg, this.pRg, pr));
        this.pRg = pr;
        this.setParentToChild(pr);
    };
    CTxEl.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.charRg !== null) {
            oCopy.setCharRg(this.charRg.createDuplicate(oIdMap));
        }
        if(this.pRg !== null) {
            oCopy.setPRg(this.pRg.createDuplicate(oIdMap));
        }
    };
    CTxEl.prototype.privateWriteAttributes = function(pWriter) {
        if(this.charRg) {
            pWriter._WriteBool2(0, true);
            pWriter._WriteUInt2(1, this.charRg.st);
            pWriter._WriteUInt2(2, this.charRg.end);
        }
        else if(this.pRg) {
            pWriter._WriteBool2(0, false);
            pWriter._WriteUInt2(1, this.pRg.st);
            pWriter._WriteUInt2(2, this.pRg.end);
        }
    };
    CTxEl.prototype.writeChildren = function(pWriter) {
    };
    CTxEl.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
     if(nType === 0) {
            var bCharRg = oStream.GetBool();
            if(bCharRg) {
                this.setCharRg(new CIndexRg())
            }
            else {
                this.setPRg(new CIndexRg());
            }
        }
        else if(1 === nType) {
            var nSt = oStream.GetULong();
            if(this.charRg) {
                this.charRg.setSt(nSt);
            }
            else if(this.pRg) {
                this.pRg.setSt(nSt);
            }
        }
        else if(2 === nType) {
            var nEnd = oStream.GetULong();
            if(this.charRg) {
                this.charRg.setEnd(nEnd);
            }
            else if(this.pRg) {
                this.pRg.setEnd(nEnd);
            }
        }
    };
    CTxEl.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_WheelSpokes] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_WheelSpokes] = function(oClass, value) {oClass.spokes = value;};
    function CWheel() {
        CBaseFormatObject.call(this);
        this.spokes = null;
    }
    InitClass(CWheel, CBaseFormatObject, AscDFH.historyitem_type_Wheel);
    CWheel.prototype.setSpokes = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_WheelSpokes, this.spokes, pr));
        this.spokes = pr;
    };
    CWheel.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.spokes !== null) {
            oCopy.setSpokes(this.spokes);
        }
    };
    CWheel.prototype.privateWriteAttributes = function(pWriter) {
    };
    CWheel.prototype.writeChildren = function(pWriter) {
    };
    CWheel.prototype.readAttribute = function(nType, pReader) {
    };
    CWheel.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_AttrNameText] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_AttrNameText] = function(oClass, value) {oClass.text = value;};
    function CAttrName() {
        CBaseFormatObject.call(this);
        this.text = null;
    }
    InitClass(CAttrName, CBaseFormatObject, AscDFH.historyitem_type_AttrName);
    CAttrName.prototype.setText = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_AttrNameText, this.spokes, pr));
        this.text = pr;
    };
    CAttrName.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.text !== null) {
            oCopy.setText(this.text);
        }
    };
    CAttrName.prototype.privateWriteAttributes = function(pWriter) {
        pWriter._WriteString1(0, this.text);
    };
    CAttrName.prototype.writeChildren = function(pWriter) {
    };
    CAttrName.prototype.readAttribute = function(nType, pReader) {
        var oStream = pReader.stream;
        if(nType === 0) {
            this.setText(oStream.GetString2());
        }
    };
    CAttrName.prototype.readChild = function(nType, pReader) {
        var oStream = pReader.stream;
        oStream.SkipRecord();
    };

    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].CTiming = CTiming;
    window['AscFormat'].CEmptyObject = CEmptyObject;
    window['AscFormat'].CCommonTimingList = CCommonTimingList;
    window['AscFormat'].CAttrNameLst = CAttrNameLst;
    window['AscFormat'].CBldLst = CBldLst;
    window['AscFormat'].CCondLst = CCondLst;
    window['AscFormat'].CChildTnLst = CChildTnLst;
    window['AscFormat'].CTmplLst = CTmplLst;
    window['AscFormat'].CTnLst = CTnLst;
    window['AscFormat'].CTavLst = CTavLst;
    window['AscFormat'].CObjectTarget = CObjectTarget;
    window['AscFormat'].CBldBase = CBldBase;
    window['AscFormat'].CBldDgm = CBldDgm;
    window['AscFormat'].CBldGraphic = CBldGraphic;
    window['AscFormat'].CBldOleChart = CBldOleChart;
    window['AscFormat'].CBldP = CBldP;
    window['AscFormat'].CBldSub = CBldSub;
    window['AscFormat'].CDirTransition = CDirTransition;
    window['AscFormat'].COptionalBlackTransition = COptionalBlackTransition;
    window['AscFormat'].CGraphicEl = CGraphicEl;
    window['AscFormat'].CIndexRg = CIndexRg;
    window['AscFormat'].CTmpl = CTmpl;
    window['AscFormat'].CAnim = CAnim;
    window['AscFormat'].CCBhvr = CCBhvr;
    window['AscFormat'].CCTn = CCTn;
    window['AscFormat'].CCond = CCond;
    window['AscFormat'].CRtn = CRtn;
    window['AscFormat'].CTgtEl = CTgtEl;
    window['AscFormat'].CSndTgt = CSndTgt;
    window['AscFormat'].CSpTgt = CSpTgt;
    window['AscFormat'].CIterateData = CIterateData;
    window['AscFormat'].CTm = CTm;
    window['AscFormat'].CTav = CTav;
    window['AscFormat'].CAnimVariant = CAnimVariant;
    window['AscFormat'].CAnimClr = CAnimClr;
    window['AscFormat'].CAnimEffect = CAnimEffect;
    window['AscFormat'].CAnimMotion = CAnimMotion;
    window['AscFormat'].CAnimRot = CAnimRot;
    window['AscFormat'].CAnimScale = CAnimScale;
    window['AscFormat'].CAudio = CAudio;
    window['AscFormat'].CCMediaNode = CCMediaNode;
    window['AscFormat'].CCmd = CCmd;
    window['AscFormat'].CTimeNodeContainer = CTimeNodeContainer;
    window['AscFormat'].CPar = CPar;
    window['AscFormat'].CExcl = CExcl;
    window['AscFormat'].CSeq = CSeq;
    window['AscFormat'].CSet = CSet;
    window['AscFormat'].CVideo = CVideo;
    window['AscFormat'].COleChartEl = COleChartEl;
    window['AscFormat'].CTLPoint = CTLPoint;
    window['AscFormat'].CSndAc = CSndAc;
    window['AscFormat'].CStSnd = CStSnd;
    window['AscFormat'].CTxEl = CTxEl;
    window['AscFormat'].CWheel = CWheel;
    window['AscFormat'].CAttrName = CAttrName;
})(window);
