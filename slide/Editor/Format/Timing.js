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

    var oPercentageRegeExp = new RegExp("((100)|([0-9][0-9]?))(\.[0-9][0-9]?)?%", "g");

    var InitClass = AscFormat.InitClass;
    var CBaseFormatObject = AscFormat.CBaseFormatObject;

    var GENERATE_PRESETS_SCRIPT = false;
    var aPresetClasses = [];
    function getAnimPresetsScript() {
        var sScript = "var ANIMATION_PRESET_CLASSES = [];\n";
        sScript += "var PRESET_TYPES;\n";
        sScript += "var PRESET_SUBTYPES;\n";
        for(var nPresetClass = 0; nPresetClass < aPresetClasses.length; ++nPresetClass) {
            if(aPresetClasses[nPresetClass]) {
                sScript += "ANIMATION_PRESET_CLASSES[" + nPresetClass + "] = [];\n";
                sScript += "PRESET_TYPES = ANIMATION_PRESET_CLASSES[" + nPresetClass + "] = [];\n";
                var aPresets = aPresetClasses[nPresetClass];
                if(aPresets) {
                    for(var nPreset = 0; nPreset < aPresets.length; ++nPreset) {
                        var aPresetSubtypes = aPresets[nPreset];
                        if(aPresetSubtypes) {
                            sScript += "PRESET_TYPES[" + nPreset + "] = [];\n";
                            sScript += "PRESET_SUBTYPES = PRESET_TYPES[" + nPreset + "] = [];\n";
                            for(var nSubtype = 0; nSubtype < aPresetSubtypes.length; ++nSubtype) {
                                if(aPresetSubtypes[nSubtype]) {
                                    sScript += "PRESET_SUBTYPES[" + nSubtype + "] = \"" + aPresetSubtypes[nSubtype] + "\";\n";
                                }
                            }
                        }
                    }
                }
            }
        }
        return console.log(sScript);
    }
    AscFormat.getAnimPresetsScript = getAnimPresetsScript;


    


   


    function CBaseAnimObject() {
        CBaseFormatObject.call(this);
    }
    InitClass(CBaseAnimObject, CBaseFormatObject, AscDFH.historyitem_type_Unknown);
    CBaseAnimObject.prototype.Refresh_RecalcData2 = function() {
        if(this.parent && this.parent.Refresh_RecalcData2) {
            this.parent.Refresh_RecalcData2();
        }
    };

    if(GENERATE_PRESETS_SCRIPT) {
        CBaseAnimObject.prototype.fromPPTY = function(pReader) {
            var oStream = pReader.stream;
            var nStart = oStream.cur;
            var nEnd = nStart + oStream.GetULong() + 4;
            this.readAttributes(pReader);
            this.readChildren(nEnd, pReader);
            oStream.Seek2(nEnd);
            if(this.getObjectType() === AscDFH.historyitem_type_Par) {
                var oCTn = this.cTn;
                if(oCTn && oCTn.presetClass != null && oCTn.presetID != null) {
                    console.log("SLIDENUM: " + editor.WordControl.m_oLogicDocument.Slides.length);
                    if(!aPresetClasses[oCTn.presetClass]) {
                        aPresetClasses[oCTn.presetClass] = [];
                    }
                    if(!aPresetClasses[oCTn.presetClass][oCTn.presetID]) {
                        aPresetClasses[oCTn.presetClass][oCTn.presetID] = [];
                    }
                    var nPresetSubtype = oCTn.presetSubtype || 0;
                    var nLength = nEnd - nStart;
                    var aData = oStream.data.slice(nStart, nStart + nLength);
                    var sData = "PPTY;v10;";
                    sData += (nLength + ";");
                    sData += AscCommon.Base64Encode(aData,aData.length, 0);

                    aPresetClasses[oCTn.presetClass][oCTn.presetID][nPresetSubtype] = sData;
                }
            }
        };
    }
    CBaseAnimObject.prototype.isTimeNode = function() {
        return false;
    };
    CBaseAnimObject.prototype.isTimingContainer = function() {
        return false;
    };
    CBaseAnimObject.prototype.parseTime = function(val) {
        return new CAnimationTime(this.parseTimeValue(val));
    };
    CBaseAnimObject.prototype.parseTimeValue = function(val) {
        if(!(typeof val === "string")) {
            return CAnimationTime.prototype.Unspecified;
        }
        if(val === "indefinite") {
            return CAnimationTime.prototype.Indefinite;
        }
        var nVal = parseInt(val);
        if(AscFormat.isRealNumber(nVal)) {
            return nVal;
        }
        return CAnimationTime.prototype.Unspecified;
    };
    CBaseAnimObject.prototype.getNearestParentOrEqualTimeNode = function() {
        var oCurObj = this;
        while(oCurObj && !oCurObj.isTimeNode()) {
            oCurObj = oCurObj.parent;
        }
        return oCurObj;
    };
    CBaseAnimObject.prototype.getRootNode = function() {
        var oNode = this.getNearestParentOrEqualTimeNode();
        if(oNode) {
            return oNode.getRoot();
        }
        return null;
    };
    CBaseAnimObject.prototype.findTimeNodeById = function(id) {
        var oRoot = this.getRootNode();
        if(oRoot) {
            return oRoot.getTimeNodeById(id);
        }
        return null;
    };
    CBaseAnimObject.prototype.parsePercentage = function(sVal) {
        var oResult = oPercentageRegeExp.exec(sVal);
        if(oResult && oResult.index === 0) {
            var sValue = sVal.slice(0, sVal.length - 1);
            var aParts = sValue.split(".");
            var dResult = parseInt(aParts[0]) / 100;
            if(aParts.length > 1) {
                dResult += parseInt(aParts[1]) / 100/ Math.pow(10, aParts[1].length);
            }
            return dResult;
        }
        return 0;
    };
    CBaseAnimObject.prototype.getEffectById = function(sId, oMltEffect) {
        if(!oMltEffect) {
            var oNewMltEffect = new CMultiple();
            return this.getEffectById(sId, oMltEffect);
        }
        this.traverse(function(oChild) {
            oChild.getEffectById(sId, oMltEffect);
            return false;
        });
        return oMltEffect;
    };
    CBaseAnimObject.prototype.createCCTn = function(sDur, nFill, sDelay, nNodeType, nRestart, bCreateChldLst, nAccel) {
        var oCCTn = new CCTn();
        if(sDur) {
            oCCTn.setDur(sDur);
        }
        if(AscFormat.isRealNumber(nFill)) {
            oCCTn.setFill(nFill);
        }
        if(AscFormat.isRealNumber(nNodeType)) {
            oCCTn.setNodeType(nNodeType);
        }
        if(sDelay) {
            oCCTn.createStCondLstWithDelay(sDelay);
        }
        if(AscFormat.isRealNumber(nRestart)) {
            oCCTn.setRestart(nRestart);
        }
        if(AscFormat.isRealNumber(nAccel)) {
            if(nAccel >= 0) {
                oCCTn.setAccel(nAccel);
            }
            else {
                oCCTn.setDecel(-nAccel);
            }
        }
        if(bCreateChldLst) {
            oCCTn.setChildTnLst(new CChildTnLst());
        }
        return oCCTn;
    };
    CBaseAnimObject.prototype.getPresentation = function() {
        return editor.WordControl.m_oLogicDocument;
    };
    CBaseAnimObject.prototype.isAnimObject = true;
    CBaseAnimObject.prototype.getTiming = function() {
        var oCurElement = this;
        while(oCurElement && !(oCurElement instanceof CTiming)) {
            oCurElement = oCurElement.parent;
        }
        if(oCurElement instanceof CTiming) {
            return oCurElement;
        }
        return null;
    };
    var TIME_NODE_STATE_IDLE = 0;
    var TIME_NODE_STATE_ACTIVE = 1;
    var TIME_NODE_STATE_FROZEN = 2;
    var TIME_NODE_STATE_FINISHED = 3;

    var oSTATEDESCRMAP = {};

    oSTATEDESCRMAP[TIME_NODE_STATE_IDLE] = 'IDLE';
    oSTATEDESCRMAP[TIME_NODE_STATE_ACTIVE] = 'ACTIVE';
    oSTATEDESCRMAP[TIME_NODE_STATE_FROZEN] = 'FROZEN';
    oSTATEDESCRMAP[TIME_NODE_STATE_FINISHED] = 'FINISHED';





    var NODE_TYPE_MAP = {};
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_AFTEREFFECT	 ] = "AFTEREFFECT";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_AFTERGROUP	 ] = "AFTERGROUP";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_CLICKEFFECT	 ] = "CLICKEFFECT";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_CLICKPAR		 ] = "CLICKPAR";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_INTERACTIVESEQ ] = "INTERACTIVESEQ";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_MAINSEQ		 ] = "MAINSEQ";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_TMROOT		 ] = "TMROOT";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_WITHEFFECT	 ] = "WITHEFFECT";
    NODE_TYPE_MAP[AscFormat.NODE_TYPE_WITHGROUP		 ] = "WITHGROUP";


    var NODE_FILL_FREEZE = 0;
    var NODE_FILL_HOLD = 1;
    var NODE_FILL_REMOVE = 2;
    var NODE_FILL_TRANSITION = 3;

    var ANIM_TREE_LAVELS_COUNT = 5;


    function CTimeNodeBase() {
        CBaseAnimObject.call(this);
        this.state = TIME_NODE_STATE_IDLE;

        this.simpleDurationIdx = -1;

        this.originalNode = undefined;
    }
    InitClass(CTimeNodeBase, CBaseAnimObject, AscDFH.historyitem_type_Unknown);
    CTimeNodeBase.prototype.isTimingContainer = function() {
        return this.isPar() || this.isSeq() || this.isExcl();
    };
    CTimeNodeBase.prototype.isPar = function() {
        var nType = this.getObjectType();
        return (nType === AscDFH.historyitem_type_Par);
    };
    CTimeNodeBase.prototype.isSeq = function() {
        var nType = this.getObjectType();
        return (nType === AscDFH.historyitem_type_Seq);
    };
    CTimeNodeBase.prototype.isExcl = function() {
        var nType = this.getObjectType();
        return (nType === AscDFH.historyitem_type_Excl);
    };
    CTimeNodeBase.prototype.isTimeNode = function() {
        return true;
    };
    CTimeNodeBase.prototype.getParentTimeNode = function() {
        var oCurParent = this.parent;
        while (oCurParent && !oCurParent.isTimeNode()) {
            if(oCurParent.getObjectType() === AscDFH.historyitem_type_Timing) {
                return null;
            }
            oCurParent = oCurParent.parent;
        }
        return oCurParent;
    };
    CTimeNodeBase.prototype.isPredecessor = function(oNode) {
        var oCurParent = this.parent;
        while (oCurParent && oNode !== oCurParent) {
            oCurParent = oCurParent.parent;
        }
        return AscCommon.isRealObject(oCurParent);
    };
    CTimeNodeBase.prototype.getChildrenTimeNodes = function() {
        if(!this.isTimingContainer()) {
            return [];
        }
        return this.getChildrenTimeNodesInternal();
    };
    CTimeNodeBase.prototype.getChildrenTimeNodesInternal = function() {
        return [];
    };
    CTimeNodeBase.prototype.resetChildrenState = function() {
        if(this.isTimingContainer()) {
            var aChildren = this.getChildrenTimeNodes();
            for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                aChildren[nChild].resetState();
            }
        }
    };
    CTimeNodeBase.prototype.resetState = function() {
        this.setState(TIME_NODE_STATE_IDLE);
        this.simpleDurationIdx = -1;
        this.resetChildrenState();
    };
    CTimeNodeBase.prototype.isRoot = function() {
        var oParentNode = this.getParentTimeNode();
        if(!oParentNode) {
            return true;
        }
    };
    CTimeNodeBase.prototype.getRoot = function() {
        var oCurElem = this;
        var oCurParent;
        while (true) {
            oCurParent = oCurElem.getParentTimeNode();
            if(!oCurParent) {
                break;
            }
            oCurElem = oCurParent;
        }
        return oCurElem;
    };
    CTimeNodeBase.prototype.getDepth = function() {
        var nDepth = 0;
        var oCurNode = this;
        while (oCurNode = oCurNode.getParentTimeNode()) {
            ++nDepth;
        }
        return nDepth;
    };
    CTimeNodeBase.prototype.getPreviousNode = function() {
        var oParentNode = this.getParentTimeNode();
        if(!oParentNode) {
            return null;
        }
        return oParentNode.getChildNode(oParentNode.getChildNodeIdx(this) - 1);
    };
    CTimeNodeBase.prototype.getChildNode = function(nIdx) {
        return this.getChildrenTimeNodes()[nIdx] || null;
    };
    CTimeNodeBase.prototype.scheduleStart = function(oPlayer) {
        oPlayer.scheduleEvent(new CAnimEvent(this.getActivateCallback(oPlayer), this.getStartTrigger(oPlayer), this));
    };
    CTimeNodeBase.prototype.createExternalEventTrigger = function(oPlayer, oTrigger, nType, sSpId) {
        var oThis = this;
        //check slide transition advance after
        var bAdvanceAfter = false;
        var aChildren = this.getChildrenTimeNodes();
        if(nType === COND_EVNT_ON_NEXT && this.isMainSequence()) {
            var oSlide = oPlayer.slide;
            if(oSlide) {
                if(oSlide.isAdvanceAfterTransition()) {
                    bAdvanceAfter = true;
                }
            }
        }
        var fTrigger = function () {
            var oEvent = oPlayer.getExternalEvent();
            if(!oEvent) {
                if(bAdvanceAfter) {
                    var bCanAdvance = false;
                    for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                        var oChild = aChildren[nChild];
                        if(!oChild.isIdle()) {
                            if(!oChild.isAtEnd()) {
                                break;
                            }
                        }
                        else {
                            bCanAdvance = true;
                            break;
                        }
                    }
                    if(bCanAdvance) {
                        oPlayer.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_NEXT, null));
                        return fTrigger();
                    }
                    return false;
                }
                else {
                    return false;
                }
            }
            var aHandledNodes = oEvent.handledNodes;
            var nNode, oNode;
            var oHandledTrigger;
            var sHandledSpId;
            if(oEvent.isEqualType(nType) && (!sSpId || oEvent.target === sSpId)) {
                for(nNode = 0; nNode < aHandledNodes.length; ++nNode) {
                    oNode = aHandledNodes[nNode].node;
                    oHandledTrigger = aHandledNodes[nNode].trigger;
                    sHandledSpId = aHandledNodes[nNode].sp;
                    if(oNode.isSibling(oThis)) {
                        return false;
                    }
                    if(oNode === oThis) {
                        if(oHandledTrigger !== oTrigger) {
                            return false;
                        }
                    }
                    if(sHandledSpId && !sSpId && !oThis.isPredecessor(oNode)) {
                        return false;
                    }
                }
                for(nNode = 0; nNode < aHandledNodes.length; ++nNode) {
                    oNode = aHandledNodes[nNode].node;
                    oHandledTrigger = aHandledNodes[nNode].trigger;
                    if(oNode === oThis && oHandledTrigger === oTrigger) {
                        break;
                    }
                }
                if(nNode === aHandledNodes.length) {
                    oEvent.handledNodes.push({node: oThis, trigger: oTrigger, type: nType, sp: sSpId});
                }
                return true;
            }
            return false;
        };
        return fTrigger;
    };
    CTimeNodeBase.prototype.isSibling = function(oNode) {
        if(this !== oNode && oNode.getParentTimeNode() === this.getParentTimeNode()) {
            return true;
        }
        return false;
    };
    CTimeNodeBase.prototype.createEffectTrigger = function(fExternalTrigger, oPlayer) {
        var oAttributes = this.getAttributesObject();
        var fTrigger = (function() {
            var oAddtionalTrigger;
            return function() {
                if(!oAddtionalTrigger) {
                    if(fExternalTrigger()) {
                        oAddtionalTrigger = oAttributes.stCondLst.createComplexTrigger(oPlayer);
                    }
                }
                if(oAddtionalTrigger) {
                    return oAddtionalTrigger.isFired(oPlayer);
                }
                return false;
            };
        })();
        var oTrigger = this.getDefaultTrigger(oPlayer);
        oTrigger.addTrigger(fTrigger);
        return oTrigger;
    };
    CTimeNodeBase.prototype.isMediaCallEffect = function() {
        var oAttributes = this.getAttributesObject();
        return !!(oAttributes && oAttributes.presetClass === AscFormat.PRESET_CLASS_MEDIACALL);
    };
    CTimeNodeBase.prototype.getStartTrigger = function(oPlayer) {
        var oAttributes = this.getAttributesObject();
        if(!oAttributes || !oAttributes.stCondLst) {
            return this.getDefaultTrigger(oPlayer);
        }
        var oTrigger;
        var nNodeType = this.getNodeType();
        var oPreviousTimeNode;
        switch (nNodeType) {
            case AscFormat.NODE_TYPE_MAINSEQ: {
                oTrigger = this.getDefaultTrigger(oPlayer);
                break;
            }
            case AscFormat.NODE_TYPE_CLICKEFFECT: {
                oTrigger = this.createEffectTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_CLICK, null), oPlayer);
                break;
            }
            case AscFormat.NODE_TYPE_WITHEFFECT: {
                oPreviousTimeNode = this;
                while ((oPreviousTimeNode = oPreviousTimeNode.getPreviousNode()) &&
                (oPreviousTimeNode.getNodeType() === AscFormat.NODE_TYPE_WITHEFFECT ||
                 oPreviousTimeNode.getNodeType() === AscFormat.NODE_TYPE_AFTEREFFECT)) {
                }
                if(oPreviousTimeNode) {
                    oTrigger = this.createEffectTrigger(function() {
                        return oPreviousTimeNode.isActive() || oPreviousTimeNode.isAtEnd();
                    }, oPlayer);
                }
                else {
                    oTrigger = oAttributes.stCondLst.createComplexTrigger(oPlayer);
                }
                break;
            }
            case AscFormat.NODE_TYPE_AFTEREFFECT: {
                oPreviousTimeNode = this.getPreviousNode();
                if(oPreviousTimeNode) {
                    oTrigger = this.createEffectTrigger(function() {
                        return oPreviousTimeNode.isAtEnd();
                    }, oPlayer);
                }
                else {
                    oTrigger = oAttributes.stCondLst.createComplexTrigger(oPlayer);
                }
                break;
            }
            default: {
                oTrigger = oAttributes.stCondLst.createComplexTrigger(oPlayer);
                if(oTrigger.isDefault()) {
                    var oParentNode = this.getParentTimeNode();
                    if(oParentNode) {
                        if(oParentNode.isSeq()) {
                            oTrigger.addNever();
                        }
                    }
                }
                break;
            }
        }
        return oTrigger;
    };
    CTimeNodeBase.prototype.getDur = function() {
        var oAttr = this.getAttributesObject();
        if(oAttr.dur === null) {
            return (new CAnimationTime(0)).setUnspecified();
        }
        return this.parseTime(oAttr.dur);
    };
    CTimeNodeBase.prototype.getRepeatDur = function() {
        return this.parseTime(this.getAttributesObject().repeatDur);
    };
    CTimeNodeBase.prototype.getRepeatCount = function() {
         return this.parseTime(this.getAttributesObject().repeatCount);
    };
    CTimeNodeBase.prototype.getNodeType = function() {
         return this.getAttributesObject().nodeType;
    };
    CTimeNodeBase.prototype.getImplicitDuration = function() {
        return 2000;//TODO:
    };
    CTimeNodeBase.prototype.calculateSimpleDuration = function() {
        var oAttr = this.getAttributesObject();
        var oTime = new CAnimationTime();
        if(oAttr.dur === null && oAttr.endCondLst) {
            oTime.setIndefinite();
            return oTime;
        }
        oTime.setUnresolved();
        var oDurTime = this.getDur();
        if(oDurTime.isDefinite()) {
            if(oAttr.spd !== null) {
                oDurTime.scale(Math.abs(oAttr.spd));
            }
            return oDurTime;
        }
        else if(oDurTime.isIndefinite()) {
            return oDurTime;
        }
        else if(oDurTime.isUnspecified()) {
            return oDurTime;
        }
        else if(oDurTime.isMedia()) {
            oDurTime.setVal(this.getImplicitDuration());
            return oDurTime;
        }
        return oTime;
    };
    CTimeNodeBase.prototype.calculateRepeatCount = function() {
        var oCount = new CAnimationTime();
        var oRepeatDur = this.getRepeatDur();
        if(oRepeatDur.isDefinite()) {
            oCount.assign(oRepeatDur);
            var oSimpleDur = this.calculateSimpleDuration();
            oCount.divideAssign(oSimpleDur);
            oCount.multiplyAssign(1000);
            return oCount;
        }
        var oRepeatCount = this.getRepeatCount();
        if(!oRepeatCount.isUnspecified()) {
            return oRepeatCount;
        }
        oCount.setVal(1000);
        return oCount;
    };
    CTimeNodeBase.prototype.activateCallback = function(oPlayer) {
        if(this.isActive()) {
            return;
        }
        var oParent = this.getParentTimeNode();
        if(oParent) {
            if(!oParent.isActive()) {
                return;
            }
        }
        this.calculateParams(oPlayer);
        this.setState(TIME_NODE_STATE_ACTIVE);
        var oParentNode = this.getParentTimeNode();
        if(oParentNode) {
            oParentNode.onActivated(this, oPlayer);
        }
        this.startSimpleDuration(0, oPlayer);
        this.scheduleEnd(oPlayer);
    };
    CTimeNodeBase.prototype.startSimpleDuration = function(nIdx, oPlayer) {
        this.simpleDurationIdx = nIdx;
        this.resetChildrenState();
        this.activateChildrenCallback(oPlayer);
    };
    CTimeNodeBase.prototype.calculateParams = function(oPlayer) {
        this.startTick = oPlayer.getElapsedTicks();
        this.simpleDuration = this.calculateSimpleDuration();
        this.repeatCount = this.calculateRepeatCount();
        this.privateCalculateParams()
    };
    CTimeNodeBase.prototype.privateCalculateParams = function(oPlayer) {
    };
    CTimeNodeBase.prototype.scheduleEnd = function(oPlayer) {
        var fFinishTrigger = this.getEndTrigger(oPlayer);
        if(fFinishTrigger !== null) {
            oPlayer.scheduleEvent(new CAnimEvent(
                this.getEndCallback(oPlayer),
                fFinishTrigger,
                this
            ));
        }
    };
    CTimeNodeBase.prototype.getEndTrigger = function(oPlayer) {
        if(!this.isTimingContainer() && !this.getTargetObjectId()) {
            return this.getDefaultTrigger(oPlayer);
        }
        if(this.simpleDuration.isDefinite() && this.repeatCount.isDefinite()) {
            return this.getTimeTrigger(
                oPlayer,
                this.startTick + this.simpleDuration.getVal() * this.repeatCount.getVal() / 1000
            );
        }
        else {
            if(this.isTimingContainer()) {
                var oEndSync = this.getAttributesObject().endSync;
                if(!this.repeatCount.isIndefinite() && !this.isRoot() && !this.isMainSequence() || oEndSync) {
                    var oTrigger = new CAnimComplexTrigger();
                    var aChildren = this.getChildrenTimeNodes();
                    var oThis = this;
                    oTrigger.addTrigger(function() {
                        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                            if(!aChildren[nChild].isAtEnd()) {
                                return false;
                            }
                        }
                        if(oThis.checkRepeatCondition(oPlayer)) {
                            return false;
                        }
                        return true;
                    });
                    if(oEndSync) {
                        oEndSync.fillTrigger(oPlayer, oTrigger);
                    }
                    return oTrigger;
                }
            }
        }
        return null;
    };
    CTimeNodeBase.prototype.finishCallback = function(oPlayer) {
        var aChildren = this.getChildrenTimeNodes();
        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
            aChildren[nChild].getEndCallback(oPlayer)();
        }
        this.setFinished(oPlayer);
    };
    CTimeNodeBase.prototype.freezeCallback = function(oPlayer) {
        var aChildren = this.getChildrenTimeNodes();
        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
            aChildren[nChild].getEndCallback(oPlayer)();
        }
        this.setFrozen(oPlayer);
    };
    CTimeNodeBase.prototype.setFrozen = function(oPlayer) {
        if(this.isFrozen()) {
            return;
        }
        if(this.isIdle()) {
            this.calculateParams(oPlayer);
        }
        this.setState(TIME_NODE_STATE_FROZEN);
        oPlayer.cancelCallerEvent(this);
        var oParentNode = this.getParentTimeNode();
        if(oParentNode) {
            oParentNode.onFrozen(this, oPlayer);
        }
        if(this.isRoot()) {
            oPlayer.stop();
        }
    };
    CTimeNodeBase.prototype.setFinished = function(oPlayer) {
        if(this.isFinished()) {
            return;
        }
        if(this.isIdle()) {
            this.calculateParams(oPlayer);
        }
        this.setState(TIME_NODE_STATE_FINISHED);
        oPlayer.cancelCallerEvent(this);
        var oParentNode = this.getParentTimeNode();
        if(oParentNode) {
            oParentNode.onFinished(this, oPlayer);
        }
        //if(this.isRoot()) {
        //    oPlayer.stop();
        //}
    };
    CTimeNodeBase.prototype.getEndCallback = function(oPlayer) {
        var oThis = this;
        var oAttribute = this.getAttributesObject();
        if(oAttribute.fill === NODE_FILL_HOLD || oAttribute.fill === NODE_FILL_FREEZE) {
            return function () {
                oThis.freezeCallback(oPlayer);
                oThis.checkTriggerStartOnEnd(oPlayer);
            }
        }
        return function () {
            oThis.finishCallback(oPlayer);
            oThis.checkTriggerStartOnEnd(oPlayer);
        };
    };
    CTimeNodeBase.prototype.checkTriggerStartOnEnd = function(oPlayer) {
        var oThis = this;
        if(oThis.getSpClickInteractiveSeq()) {
            var nElapsed = oPlayer.getElapsedTicks();
            oPlayer.scheduleEvent(new CAnimEvent(function() {
                    oThis.scheduleStart(oPlayer);
                },
                new CAnimComplexTrigger(function () {
                    return oPlayer.getElapsedTicks() > nElapsed;
                }),
                oThis
            ));
        }
    };
    CTimeNodeBase.prototype.activateChildrenCallback = function(oPlayer) {
    };
    CTimeNodeBase.prototype.getActivateCallback = function(oPlayer) {
        var oThis = this;
        return function() {
            oThis.activateCallback(oPlayer);
        };
    };
    CTimeNodeBase.prototype.getDefaultTrigger = function(oPlayer) {
        return new CAnimComplexTrigger();
    };
    CTimeNodeBase.prototype.getTimeTrigger = function(oPlayer, nTime) {
        return new CAnimComplexTrigger(function() {
            return nTime <= oPlayer.getElapsedTicks();
        });
    };
    CTimeNodeBase.prototype.setState = function(nState) {
        this.state = nState;

        //this.logState("SET STATE:");
    };
    CTimeNodeBase.prototype.logState = function (sPrefix) {
        var oAttr = this.getAttributesObject();
        var sNodeType = NODE_TYPE_MAP[oAttr.nodeType];
        if(sNodeType) 
        console.log(sPrefix + " | ID: " + this.Id + " | TYPE: " + this.constructor.name + " | NODE_TYPE: " + sNodeType + " | STATE: " + oSTATEDESCRMAP[this.state] + " | TIME: " + (new Date()).getTime() + " | FORMAT ID: " + oAttr.id);
    };

    CTimeNodeBase.prototype.printTree = function () {
        var nDepth = this.getDepth();
        var sString = "";
        for(var nIdx = 0; nIdx < nDepth; ++nIdx) {
            sString += "        ";
        }

        var oAttr = this.getAttributesObject();
        var sNodeType = NODE_TYPE_MAP[oAttr.nodeType];
        sString += (nDepth + " TYPE: " + this.constructor.name + " | NODE_TYPE: " + sNodeType + " | FORMAT ID: " + oAttr.id + " | ID: " + this.Id);
        console.log(sString);
        var aChildren = this.getChildrenTimeNodes();
        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
            aChildren[nChild].printTree();
        }
    };
    CTimeNodeBase.prototype.getFormatId = function () {
        return this.getAttributesObject().id;
    };
    CTimeNodeBase.prototype.cancelEventsRecursive = function(oPlayer) {
        oPlayer.cancelCallerEvent(this);
        if(this.isTimingContainer()) {
            this.cancelChildrenEventsRecursive(oPlayer);
        }
    };
    CTimeNodeBase.prototype.cancelChildrenEventsRecursive = function(oPlayer) {
        if(this.isTimingContainer()) {
            var aChildren = this.getChildrenTimeNodes();
            for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                aChildren[nChild].cancelEventsRecursive(oPlayer);
            }
        }
    };
    CTimeNodeBase.prototype.getChildNodeIdx = function(oChildNode) {
        var aChildNodes = this.getChildrenTimeNodes();
        var nRet = -1;
        for(var nIdx = 0; nIdx < aChildNodes.length; ++nIdx) {
            if(aChildNodes[nIdx] === oChildNode) {
                return nIdx;
            }
        }
        return nRet;
    };
    CTimeNodeBase.prototype.onIdle = function(oChild, oPlayer) {
        //TODO
    };
    CTimeNodeBase.prototype.onActivated = function(oChild, oPlayer) {
        //TODO
    };
    CTimeNodeBase.prototype.onFrozen = function(oChild, oPlayer) {
        return this.onFinished(oChild, oPlayer);
    };
    CTimeNodeBase.prototype.checkRepeatCondition = function(oPlayer) {
        if(oPlayer && oPlayer.bDoNotRestart) {
            return false;
        }
        return this.repeatCount.isSpecified() && this.simpleDurationIdx + 1 < this.repeatCount.getVal() / 1000;
    };
    CTimeNodeBase.prototype.onFinished = function(oChild, oPlayer) {
        if(!this.isActive()) {
            return;
        }
        if(!this.isTimingContainer()) {
            return;
        }
        var aChildren = this.getChildrenTimeNodes();
        var nChild;
        if(this.isPar()) {
            for(nChild = 0; nChild < aChildren.length; ++nChild) {
                if(!aChildren[nChild].isAtEnd()) {
                    break;
                }
            }
            if(nChild === aChildren.length) {
                if(this.checkRepeatCondition(oPlayer)) {
                    this.startSimpleDuration(++this.simpleDurationIdx, oPlayer);
                }
            }
        }
        else if(this.isSeq()) {
            var nChildIdx = this.getChildNodeIdx(oChild);
            if(nChildIdx < aChildren.length - 1) {
                aChildren[nChildIdx + 1].scheduleStart(oPlayer);
                // //handle advance after
                // if(this.getNodeType() === AscFormat.NODE_TYPE_MAINSEQ) {
                //     var oSlide = oPlayer.slide;
                //     if(oSlide) {
                //         var oTransition = oSlide.transition;
                //         if(oTransition) {
                //             if(oTransition.SlideAdvanceAfter) {
                //                 oPlayer.onNextSlide();
                //             }
                //         }
                //     }
                // }
            }
            else {
                if(this.checkRepeatCondition(oPlayer)) {
                    this.startSimpleDuration(++this.simpleDurationIdx, oPlayer);
                }
                else {
                    if(this.isMainSequence()) {
                        oPlayer.onMainSeqFinished();
                    }
                }
            }
        }
        if(oChild.isMainSequence()) {
            oPlayer.onMainSeqFinished();
        }
    };
    CTimeNodeBase.prototype.isIdle = function() {
        return this.state === TIME_NODE_STATE_IDLE;
    };
    CTimeNodeBase.prototype.isActive = function() {
        return this.state === TIME_NODE_STATE_ACTIVE;
    };
    CTimeNodeBase.prototype.isFrozen = function() {
        return this.state === TIME_NODE_STATE_FROZEN;
    };
    CTimeNodeBase.prototype.isFinished = function() {
        return this.state === TIME_NODE_STATE_FINISHED;
    };
    CTimeNodeBase.prototype.isDrawable = function() {
        return this.isActive() || this.isFrozen() || (this.isTimingContainer() || this.isFinished());
    };
    CTimeNodeBase.prototype.isAtEnd = function() {
        if(this.isMainSequence()) {
            var aChildren = this.getChildrenTimeNodes();
            if(aChildren.length === 0) {
                return true;
            }
            if(aChildren[aChildren.length - 1].isAtEnd()) {
                return true;
            }
        }
        return this.isFinished() || this.isFrozen();
    };
    CTimeNodeBase.prototype.getAttributesObject = function() {
        if(this.cTn) {
            return this.cTn;
        }
        if(this.cBhvr) {
            return this.cBhvr.cTn;
        }
        if(this.cMediaNode) {
            return this.cMediaNode.getAttributesObject();
        }
        return null;
    };
    CTimeNodeBase.prototype.isMainSequence = function() {
        var oAttributes = this.getAttributesObject();
        if(oAttributes && oAttributes.nodeType === AscFormat.NODE_TYPE_MAINSEQ) {
            return true;
        }
        return false;
    };
    CTimeNodeBase.prototype.isPartOfMainSequence = function() {
        var aHierarchy = this.getHierarchy();
        if(aHierarchy[1]) {
            return aHierarchy[1].isMainSequence();
        }
        return false;
    };
    CTimeNodeBase.prototype.isInteractiveSeq = function(sSpId) {
        return sSpId === this.getSpClickInteractiveSeq();
    };

    CTimeNodeBase.prototype.isPartOfInteractiveSeq = function() {
        var aHierarchy = this.getHierarchy();
        if(aHierarchy[1]) {
            return aHierarchy[1].getSpClickInteractiveSeq();
        }
        return null;
    };
    CTimeNodeBase.prototype.getSpClickInteractiveSeq = function() {
        if(this.getNodeType() === AscFormat.NODE_TYPE_INTERACTIVESEQ) {
            return this.getSpClickAdvance();
        }
        return null;
    };
    CTimeNodeBase.prototype.isClickEffect = function() {
        return this.isAnimEffect() && this.getNodeType() === AscFormat.NODE_TYPE_CLICKEFFECT;
    };
    CTimeNodeBase.prototype.isWithEffect = function() {
        return this.isAnimEffect() && this.getNodeType() === AscFormat.NODE_TYPE_WITHEFFECT;
    };
    CTimeNodeBase.prototype.isAfterEffect = function() {
        return this.isAnimEffect() && this.getNodeType() === AscFormat.NODE_TYPE_AFTEREFFECT;
    };
    CTimeNodeBase.prototype.traverseTimeNodes = function(fCallback) {
        fCallback(this);
        var aChildren = this.getChildrenTimeNodes();
        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
            aChildren[nChild].traverseTimeNodes(fCallback);
        }
    };
    CTimeNodeBase.prototype.traverseDrawable = function(oPlayer) {
        if(!this.isDrawable()) {
            return;
        }
        if(this.isTimingContainer()) {
            var aChildren = this.getChildrenTimeNodes();
            for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                aChildren[nChild].traverseDrawable(oPlayer);
            }
        }
        else {
            var sTargertId = this.getTargetObjectId();
            if(sTargertId) {
                oPlayer.addAnimationToDraw(sTargertId, this);
            }
        }
    };
    CTimeNodeBase.prototype.getTimeNodeById = function(id) {
        if(this.getAttributesObject().id === id) {
            return this;
        }
        if(this.isTimingContainer()) {
            var aChildern = this.getChildrenTimeNodes();
            var oNode = null;
            for(var nChild = 0; nChild < aChildern.length; ++nChild) {
                oNode = aChildern[nChild].getTimeNodeById(id);
                if(oNode) {
                    return oNode;
                }
            }
        }
        return null;
    };
    CTimeNodeBase.prototype.getTargetObjectId = function() {
        if(this.cBhvr) {
           return this.cBhvr.getTargetObjectId();
        }
        return null;
    };
    CTimeNodeBase.prototype.getTargetObject = function() {
        return AscCommon.g_oTableId.Get_ById(this.getTargetObjectId());
    };
    CTimeNodeBase.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
    };
    CTimeNodeBase.prototype.setAttributeValue = function(oAttributes, sName, value) {
        if(AscFormat.isRealNumber(oAttributes[sName])) {
            oAttributes[sName] += value;
        }
        else {
            oAttributes[sName] = value;
        }
    };
    CTimeNodeBase.prototype.getRewind = function() {
        var oParentTimeNode = this.getParentTimeNode();
        if(oParentTimeNode) {
            return oParentTimeNode.getRewind();
        }
        return false;
    };
    CTimeNodeBase.prototype.getRelativeTime = function(nElapsedTime) {
        var oAttr = this.getAttributesObject();
        var oParentTimeNode = this.getParentTimeNode();
        var oParentAttr = null;
        if(oParentTimeNode) {
            oParentAttr = oParentTimeNode.getAttributesObject();
        }
        var bAutoRev = oAttr.autoRev || (oParentAttr && oParentAttr.autoRev);
        var sTmFilter = oAttr.tmFilter;
        var fRelTime = 0.0;
        if(this.isFrozen() || this.isFinished()) {
            if(this.getRewind()) {
                fRelTime = 0.0;
            }
            else {
                if(bAutoRev) {
                    fRelTime = 0.0;
                }
                else {
                    fRelTime = 1.0;
                }
            }
        }
        else {
            var fSimpleDur = this.simpleDuration.getVal();
            fRelTime = (nElapsedTime - this.startTick) / fSimpleDur;
            if(bAutoRev) {
                if(fRelTime <= 0.5) {
                    fRelTime *= 2;
                }
                else {
                    fRelTime = (1 - fRelTime) * 2;
                }
            }
        }
        if(typeof sTmFilter === "string" && sTmFilter.length > 0) {
            var aPairs = sTmFilter.split(";");
            var aNumPairs = [];
            for(var nPair = 0; nPair < aPairs.length; ++nPair) {
                var aPair = aPairs[nPair].split(",");
                if(aPair.length !== 2) {
                    return fRelTime;
                }
                var fNum1 = parseFloat(aPair[0]);
                if(!AscFormat.isRealNumber(fNum1)) {
                    return fRelTime;
                }
                var fNum2 = parseFloat(aPair[1]);
                if(!AscFormat.isRealNumber(fNum2)) {
                    return fRelTime;
                }
                if(AscFormat.fApproxEqual(fRelTime, fNum1)){
                    return fNum2;
                }
                if(fRelTime <= fNum1) {
                    if(aNumPairs.length > 0) {
                        var aPrevPair = aNumPairs[aNumPairs.length - 1];
                        return aPrevPair[1] + (fRelTime - aPrevPair[0])*((fNum2 - aPrevPair[1]) / (fNum1 - aPrevPair[0]));
                    }
                    else {
                        return fRelTime;
                    }
                }
                else {
                    aNumPairs.push([fNum1, fNum2]);
                }
            }
        }
        if(oAttr.spd !== null && oAttr.spd < 0) {
            fRelTime = 1 - fRelTime;
        }
        return fRelTime;
    };
    CTimeNodeBase.prototype.getSlideWidth = function() {
        return this.getPresentation().GetWidthMM();
    };
    CTimeNodeBase.prototype.getSlideHeight = function() {
        return this.getPresentation().GetHeightMM();
    };
    CTimeNodeBase.prototype.getTargetObjectPosX = function() {
        var oObject = this.getTargetObject();
        if(!oObject) {
            return null;
        }
        return oObject.x;
    };
    CTimeNodeBase.prototype.getTargetObjectPosY = function() {
        var oObject = this.getTargetObject();
        if(!oObject) {
            return null;
        }
        return oObject.y;
    };
    CTimeNodeBase.prototype.getTargetObjectExtX = function() {
        var oObject = this.getTargetObject();
        if(!oObject) {
            return null;
        }
        return oObject.extX;
    };
    CTimeNodeBase.prototype.getTargetObjectExtY = function() {
        var oObject = this.getTargetObject();
        if(!oObject) {
            return null;
        }
        return oObject.extY;
    };
    CTimeNodeBase.prototype.getTargetObjectRelPosX = function() {
        var x = this.getTargetObjectPosX();
        if(x !== null) {
            return x/ this.getSlideWidth();
        }
        return null;
    };
    CTimeNodeBase.prototype.getTargetObjectRelPosY = function() {
        var y = this.getTargetObjectPosY();
        if(y !== null) {
            return y/ this.getSlideHeight();
        }
        return null;
    };
    CTimeNodeBase.prototype.getTargetObjectRot = function() {
        var oObject = this.getTargetObject();
        if(!oObject) {
            return null;
        }
        return oObject.rot;
    };
    CTimeNodeBase.prototype.getTargetObjectBrush = function() {
        var oObject = this.getTargetObject();
        if(!oObject) {
            return null;
        }
        if(!oObject.brush || !oObject.brush.isNoFill()) {
            var oBrush = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(255, 255, 255));
            oBrush.fill.color.RGBA.R = 255;
            oBrush.fill.color.RGBA.G = 255;
            oBrush.fill.color.RGBA.B = 255;
            oBrush.fill.color.RGBA.A = 255;
            return oBrush;
        }
        return oObject.brush;
    };
    CTimeNodeBase.prototype.getTargetObjectPen = function() {
        var oObject = this.getTargetObject();
        if(!oObject) {
            return null;
        }
        return oObject.pen;
    };
    CTimeNodeBase.prototype.getAnimatedVal = function(fTime, fStart, fEnd) {
        return fStart*(1 - fTime) + fEnd * fTime;
    };
    CTimeNodeBase.prototype.getAnimatedClr = function(fTime, oStartUniColor, oEndUniColor) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
            return null;
        }
        if(!oStartUniColor || !oEndUniColor) {
            return null;
        }
        var parents = oTargetObject.getParentObjects();
        var RGBA = {R:0, G:0, B:0, A:255};
        oEndUniColor.Calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
        oStartUniColor.Calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
        var oEndColor = oEndUniColor.RGBA;
        var oStartColor = oStartUniColor.RGBA;
        var R = this.getAnimatedVal(fTime, oStartColor.R, oEndColor.R);
        var G = this.getAnimatedVal(fTime, oStartColor.G, oEndColor.G);
        var B = this.getAnimatedVal(fTime, oStartColor.B, oEndColor.B);
        var oResultColor =  AscFormat.CreateUniColorRGB(R, G, B);
        oResultColor.Calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
        return oResultColor;
    };
    CTimeNodeBase.prototype.getAttributes = function() {
        if(!this.cBhvr) {
            return [];
        }
        return this.cBhvr.getAttributes();
    };
    CTimeNodeBase.prototype.setAttributesValue = function(oAttributes, val) {
        var aAttributes = this.getAttributes();
        for(var nAttr = 0; nAttr < aAttributes.length; ++nAttr) {
            var oAttr = aAttributes[nAttr];
            if(oAttr.text !== null) {
                this.setAttributeValue(oAttributes, oAttr.text, val);
            }
        }
    };
    CTimeNodeBase.prototype.getOrigAttrVal = function(sAttrName) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
            return null;
        }
        var oBounds = oTargetObject.getBoundsByDrawing();
        var dCenterX = oBounds.x + oBounds.w / 2;
        var dCenterY = oBounds.y + oBounds.h / 2;
        var dSlideW = this.getSlideWidth();
        var dSlideH = this.getSlideHeight();
        var dSpWidth = oBounds.w;
        var dSpHeight = oBounds.h;

        if(sAttrName === "ppt_x") {
            return dCenterX / dSlideW;
        }
        else if(sAttrName === "ppt_y") {
            return dCenterY / dSlideH;
        }
        if(sAttrName === "ppt_w") {
            return dSpWidth / dSlideW;
        }
        else if(sAttrName === "ppt_h") {
            return dSpHeight / dSlideH;
        }
        return null;
    };
    CTimeNodeBase.prototype.doesHideObject = function() {
        return false;
    };
    CTimeNodeBase.prototype.doesShowObject = function() {
        var oParentNode = this.getParentTimeNode();
        if(oParentNode) {
            var oAttrObject = oParentNode.getAttributesObject();
            if(oAttrObject) {
                if(AscFormat.PRESET_CLASS_ENTR === oAttrObject.presetClass ||
                    AscFormat.PRESET_CLASS_PATH === oAttrObject.presetClass||
                    AscFormat.PRESET_CLASS_EMPH === oAttrObject.presetClass) {
                    return true;
                }
            }
        }
        return false;
    };
    CTimeNodeBase.prototype.isAncestor = function(oNode) {
        var oCurNode = oNode;
        while(oCurNode = oCurNode.getParentTimeNode()) {
            if(oCurNode === this) {
                return true;
            }
        }
        return false;
    };
    CTimeNodeBase.prototype.isDescendant = function(oNode) {
        return oNode.isAncestor(this);
    };
    CTimeNodeBase.prototype.isAnimEffect = function() {
        return false;
    };
    CTimeNodeBase.prototype.isObjectEffect = function(sObjectId) {
        return false;
    };
    CTimeNodeBase.prototype.createSpTgt = function(sObjectId) {
        var oTgt = new CTgtEl();
        var oSpTgt = new CSpTgt();
        oSpTgt.setSpid(sObjectId);
        oTgt.setSpTgt(oSpTgt);
        return oTgt;
    };
    CTimeNodeBase.prototype.getHierarchy = function() {
        var oParentTimeNode = this.getParentTimeNode();
        var aHierarchy;
        if(oParentTimeNode) {
            aHierarchy = oParentTimeNode.getHierarchy();
        }
        else {
            aHierarchy = [];
        }
        aHierarchy.push(this);
        return aHierarchy;
    };
    CTimeNodeBase.prototype.getAllAnimEffects = function (aEffects) {
        var aEffectsInternal = aEffects;
        if(!Array.isArray(aEffectsInternal)) {
            aEffectsInternal = []
        }
        if(this.isAnimEffect()) {
            aEffectsInternal.push(this);
        }
        else {
            var aChildren = this.getChildrenTimeNodes();
            for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                aChildren[nChild].getAllAnimEffects(aEffectsInternal);
            }
        }
        return aEffectsInternal;
    };

    CTimeNodeBase.prototype.getNeighbourEffect = function(bPrev) {
        if(!this.isAnimEffect()) {
            return null;
        }
        if(this.originalNode) {
            return this.originalNode.getNeighbourEffect(bPrev);
        }
        var aHierarchy = this.getHierarchy();
        if(aHierarchy.length !== ANIM_TREE_LAVELS_COUNT) {
            return null;
        }
        var oL1 = aHierarchy[1];
        var aAllEffects = oL1.getAllAnimEffects();
        for(var nEffect = 0; nEffect < aAllEffects.length; ++nEffect) {
            if(aAllEffects[nEffect] === this) {
                if(bPrev) {
                    return aAllEffects[nEffect - 1] || null;
                }
                else {
                    return aAllEffects[nEffect + 1] || null;
                }
            }
        }

    };
    CTimeNodeBase.prototype.getPreviousEffect = function() {
        return this.getNeighbourEffect(true);
    };
    CTimeNodeBase.prototype.getNextEffect = function() {
        return this.getNeighbourEffect(false);
    };
    CTimeNodeBase.prototype.isAdvancedByShapeClick = function(sSpId) {
        return sSpId === this.getSpClickAdvance();
    };
    CTimeNodeBase.prototype.getSpClickAdvance = function() {
        var oAttr = this.getAttributesObject();
        if(!oAttr) {
            return null;
        }
        if(oAttr.stCondLst && this.nextCondLst) {
            var sStSpClick = oAttr.stCondLst.getSpClick();
            if(!sStSpClick) {
                return sStSpClick;
            }
            var sNextSpClick = this.nextCondLst.getSpClick();
            if(sNextSpClick === sStSpClick) {
                return sStSpClick;
            }
        }
        return null;
    };
    CTimeNodeBase.prototype.isRemoveAfterFill = function() {
        var oAttr = this.getAttributesObject();
        if(oAttr) {
            return oAttr.fill === NODE_FILL_REMOVE;
        }
        return false;
    };
    CTimeNodeBase.prototype.checkRemoveAtEnd = function() {
        if(this.isAtEnd() && this.isRemoveAfterFill()) {
            return true;
        }
        return false;
    };

    function CAnimationTime(val) {
        this.val = 0;
        if(typeof val === "number") {
            this.val = val;
        }
        else if(val instanceof CAnimationTime) {
            this.assign(val);
        }
    }
    CAnimationTime.prototype.Indefinite = Number.MAX_SAFE_INTEGER;
    CAnimationTime.prototype.Unresolved = Number.POSITIVE_INFINITY;
    CAnimationTime.prototype.Unspecified = CAnimationTime.prototype.Unresolved;
    CAnimationTime.prototype.Media = Number.NEGATIVE_INFINITY;
    CAnimationTime.prototype.getVal = function() {
        return this.val;
    };
    CAnimationTime.prototype.setVal = function(val) {
        this.val = val;
        return this;
    };
    CAnimationTime.prototype.setIndefinite = function() {
        return this.setVal(this.Indefinite);
    };
    CAnimationTime.prototype.setUnresolved = function() {
        return this.setVal(this.Unresolved);
    };
    CAnimationTime.prototype.setUnspecified = function() {
        return this.setVal(this.Unspecified);
    };
    CAnimationTime.prototype.setMedia = function() {
        return this.setVal(this.Media);
    };
    CAnimationTime.prototype.copy = function() {
        return new CAnimationTime(this.val);
    };
    CAnimationTime.prototype.isIndefinite = function() {
        return this.val === this.Indefinite;
    };
    CAnimationTime.prototype.isUnresolved = function() {
        return this.val === this.Unresolved;
    };
    CAnimationTime.prototype.isUnspecified = function() {
        return this.val === this.Unspecified;
    };
    CAnimationTime.prototype.isMedia = function() {
        return this.val === this.Media;
    };
    CAnimationTime.prototype.isDefinite = function() {
        return !this.isIndefinite() && !this.isUnresolved();
    };
    CAnimationTime.prototype.isResolved = function() {
        return !this.isUnresolved();
    };
    CAnimationTime.prototype.isSpecified = function() {
        return !this.isUnspecified();
    };
    CAnimationTime.prototype.less = function(oTime) {
        return this.val < oTime.val;
    };
    CAnimationTime.prototype.lessOrEquals = function(oTime) {
        return this.val <= oTime.val;
    };
    CAnimationTime.prototype.equals = function(oTime) {
        return this.val === oTime.val;
    };
    CAnimationTime.prototype.greater = function(oTime) {
        return this.val > oTime.val;
    };
    CAnimationTime.prototype.greaterOrEquals = function(oTime) {
        return this.val >= oTime.val;
    };
    CAnimationTime.prototype.notEquals = function(oTime) {
        return !this.equals(oTime);
    };
    CAnimationTime.prototype.plusAssign = function (oTime) {
        if(this.isUnresolved()) {
            return this;
        }
        else if(oTime.isUnresolved()) {
            this.val = this.Unresolved;
        }
        else if(this.isIndefinite()) {
            return this;
        }
        else if(oTime.isIndefinite()) {
            this.val = this.Indefinite;
        }
        else {
            this.val += oTime.getVal();
        }
        return this;
    };
    CAnimationTime.prototype.minusAssign = function (oTime) {
        if(this.isUnresolved()) {
            return this;
        }
        else if(oTime.isUnresolved()) {
            this.val = this.Unresolved;
        }
        else if(this.isIndefinite()) {
            return this;
        }
        else if(oTime.isIndefinite()) {
            this.val = this.Indefinite;
        }
        else {
            this.val -= oTime.getVal();
        }
        return this;
    };
    CAnimationTime.prototype.multiplyAssign = function (oTime) {
        if(!(oTime instanceof CAnimationTime)) {
            var oTimeObject = new CAnimationTime(oTime);
            return this.multiplyAssign(oTimeObject);
        }
        if(this.isUnresolved()) {
            return this;
        }
        else if(oTime.isUnresolved()) {
            this.val = this.Unresolved;
        }
        else if(this.isIndefinite()) {
            if(oTime.getVal() != 0) {
                return this;
            }
            else {
                this.val = 0;
            }
        }
        else if(oTime.isIndefinite()) {
            if(this.val != 0) {
                this.val = this.Indefinite;
            } else {
                return this;
            }
        }
        else {
            this.val *= oTime.getVal();
        }
        return this;
    };
    CAnimationTime.prototype.divideAssign = function (nCount) {
        this.val /= nCount;
        return this;
    };
    CAnimationTime.prototype.scale = function (nPrecentage) {
        this.val = this.val * nPrecentage / 100000;
        return this;
    };
    CAnimationTime.prototype.unaryMinus = function () {
        this.val = -this.val;
        return this;
    };
    CAnimationTime.prototype.plus = function (oTime) {
        return this.copy().plusAssign(oTime);
    };
    CAnimationTime.prototype.minus = function (oTime) {
        return this.copy().minusAssign(oTime);
    };
    CAnimationTime.prototype.multiply = function (oTime) {
        return this.copy().multiplyAssign(oTime);
    };
    CAnimationTime.prototype.divide = function (oTime) {
        return this.copy().divideAssign(oTime.getVal());
    };
    CAnimationTime.prototype.assign = function (oTime) {
        return this.setVal(oTime.getVal());
    };

    function CAnimationTimeInterval(begin, end) {
        this.begin = begin ? begin : new CAnimationTime();
        this.end = end ? end : new CAnimationTime();
    }
    CAnimationTimeInterval.prototype.isValid = function() {
        return this.begin.isDefinite() && this.begin.lessOrEquals(this.end);
    };
    CAnimationTimeInterval.prototype.isValidChild = function(oParent) {
        return this.begin.less(oParent.end) && this.end.greaterOrEquals(oParent.begin);
    };
    CAnimationTimeInterval.prototype.isZeroDuration = function() {
        return this.isValid() && this.begin.equals(this.end);
    };
    CAnimationTimeInterval.prototype.isDefinite = function() {
        return this.isValid() && this.end.isDefinite();
    };
    CAnimationTimeInterval.prototype.translate = function(oTime) {
        this.begin.plusAssign(oTime);
        this.end.plusAssign(oTime);
    };
    CAnimationTimeInterval.prototype.translateToBegin = function() {
        this.end.minusAssign(this.begin);
        this.begin.setVal(0);
    };
    CAnimationTimeInterval.prototype.containsTime = function(oTime) {
        return (this.begin.equals(oTime) || oTime.greater(this.begin) && oTime.less(this.end));
    };
    CAnimationTimeInterval.prototype.containsInterval = function(oInterval) {
        return this.containsTime(oInterval.begin) && (this.containsTime(oInterval.end) || this.end.equals(oInterval.end));
    };
    CAnimationTimeInterval.prototype.contains = function(oObject) {
        if(oObject instanceof CAnimationTime) {
            return this.containsTime(oObject);
        }
        if(oObject instanceof CAnimationTimeInterval) {
            return this.containsInterval(oObject);
        }
        return false;
    };
    CAnimationTimeInterval.prototype.before = function(oTime) {
        return this.end.less(oTime);
    };
    CAnimationTimeInterval.prototype.after = function(oTime) {
        return this.begin.greater(oTime);
    };
    CAnimationTimeInterval.prototype.overlaps = function(oTime1, oTime2) {
        return oTime1.lessOrEquals(this.end) && oTime2.greaterOrEquals(this.begin);
    };
    CAnimationTimeInterval.prototype.equals = function(oInterval) {
        return this.begin.equals(oInterval.begin) && this.end.equals(oInterval.end);
    };
    CAnimationTimeInterval.prototype.notEquals = function(oInterval) {
        return !this.equals(oInterval);
    };
    CAnimationTimeInterval.prototype.less = function(oInterval) {
        return this.begin.lessOrEquals(oInterval.begin) && this.end.less(oInterval.end);
    };
    CAnimationTimeInterval.prototype.greater = function(oInterval) {
        return oInterval.less(this);
    };
    CAnimationTimeInterval.prototype.lessOrEquals = function(oInterval) {
        return !(oInterval.less(this));
    };
    CAnimationTimeInterval.prototype.greaterOrEquals = function(oInterval) {
        return !(this.less(oInterval));
    };

    function CEmptyObject() {
        CBaseAnimObject.call(this);
    }
    InitClass(CEmptyObject, CBaseAnimObject, AscDFH.historyitem_type_EmptyObject);

    changesFactory[AscDFH.historyitem_TimingBldLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_TimingTnLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TimingBldLst] = function(oClass, value) {oClass.bldLst = value;};
    drawingsChangesMap[AscDFH.historyitem_TimingTnLst] = function(oClass, value) {oClass.tnLst = value;};
    function CTiming() {
        CBaseAnimObject.call(this);
        this.bldLst = null;
        this.tnLst  = null;
    }
    InitClass(CTiming, CBaseAnimObject, AscDFH.historyitem_type_Timing);
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
    
    CTiming.prototype.Refresh_RecalcData2 = function() {
        AscCommon.History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
    };
    CTiming.prototype.Refresh_RecalcData = function() {
        this.Refresh_RecalcData2();
    };
    CTiming.prototype.recalculate = function() {
    };
    CTiming.prototype.getSlideIndex = function() {
        if(this.parent && this.parent.getSlideIndex) {
            return this.parent.getSlideIndex();
        }
        return null;
    };
    CTiming.prototype.getChildren = function() {
        return [this.bldLst, this.tnLst];
    };
    CTiming.prototype.onRemoveObject = function(sObjectId) {
        this.removeObjectAnimation(sObjectId);
    };
    CTiming.prototype.removeObjectAnimation = function(sObjectId) {
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
    CTiming.prototype.onAnimPaneChanged = function(oRect) {
        var oSlide = this.parent;
        if(!oSlide) {
            return;
        }
        var oPresentation = this.getPresentation();
        oPresentation.OnAnimPaneChanged(oSlide.num, oRect)
    };
    CTiming.prototype.getTimingRootNode = function() {
        if(this.tnLst) {
            return this.tnLst.getTimeNodeByType(AscFormat.NODE_TYPE_TMROOT);
        }
        return null;
    };
    CTiming.prototype.updateNodesIDs = function() {
        var oReplaceMap = {};
        var nIdCounter = 0;
        
        //remove empty nodes
        this.traverse(function(oChild) {
            if(oChild.getObjectType() === AscDFH.historyitem_type_CTn) {
                if(oChild.parent && oChild.parent.isTimingContainer() && 
                oChild.childTnLst && oChild.childTnLst.isEmpty()) {
                    if(oChild.parent) {
                        oChild.parent.onRemoveChild(oChild);
                    }
                    return;
                }
            }
        });
        this.traverse(function(oChild) {
            if(oChild.getObjectType() === AscDFH.historyitem_type_CTn) {
                var nOldId = oChild.id;
                oChild.setId(++nIdCounter);
                if(AscFormat.isRealNumber(nOldId)) {
                    oReplaceMap[nOldId] = oChild.id;
                }
            }
        });
        this.traverse(function(oChild) {
            if(oChild.getObjectType() === AscDFH.historyitem_type_Cond) {
                if(AscFormat.isRealNumber(oChild.tn) && AscFormat.isRealNumber(oReplaceMap[oChild.tn])) {
                    oChild.setTn(oReplaceMap[oChild.tn]);
                }
            }
        });

        var aAllEffects = this.getAllAnimEffects();
        var oEffect = aAllEffects[0];
        if(oEffect && (oEffect.isAfterEffect() || oEffect.isWithEffect())) {
            var aHierarchy = oEffect.getHierarchy();
            var oSeqNode = aHierarchy[1];
            var oPar2Lvl = aHierarchy[2];
            var oPar3Lvl = aHierarchy[3];
            if(oSeqNode && oPar2Lvl && oPar3Lvl) {
                if(oPar3Lvl.getChildNodeIdx(oEffect) === 0 &&
                oPar2Lvl.getChildNodeIdx(oPar3Lvl) === 0 && 
                oSeqNode.getChildNodeIdx(oPar2Lvl) === 0) {
                    if(oPar2Lvl.cTn.stCondLst.getLength() < 2) {
                        var oCond = new CCond();
                        oCond.setEvt(COND_EVNT_BEGIN);
                        oCond.setDelay("0");
                        oCond.setTn(oSeqNode.cTn.id);
                        oPar2Lvl.cTn.stCondLst.push(oCond);
                    }
                }
            }
        }

    };
    CTiming.prototype.createEffect = function(sObjectId, nPresetClass, nPresetId, nPresetSubtype) {
        var aPresetClass = ANIMATION_PRESET_CLASSES[nPresetClass];
        if(aPresetClass) {
            var aPresetType = aPresetClass[nPresetId];
            if(aPresetType) {
                var sPresetBinary = aPresetType[nPresetSubtype];
                if(!sPresetBinary) {
                    for(var nSubtype = 0; nSubtype < aPresetType.length; ++nSubtype) {
                        if(aPresetType[nSubtype]) {
                            sPresetBinary = aPresetType[nSubtype];
                            break;
                        }
                    }
                }
                if(sPresetBinary) {
                    AscCommon.pptx_content_loader.Clear(true);
                    var stream = AscFormat.CreateBinaryReader(sPresetBinary, "PPTY;v10;".length, sPresetBinary.length);
                    var oBinaryReader = new AscCommon.BinaryPPTYLoader();
                    oBinaryReader.stream = new AscCommon.FileStream();
                    oBinaryReader.stream.obj    = stream.obj;
                    oBinaryReader.stream.data   = stream.data;
                    oBinaryReader.stream.size   = stream.size;
                    oBinaryReader.stream.pos    = stream.pos;
                    oBinaryReader.stream.cur    = stream.cur;
                    var oPar = new CPar();
                    oPar.fromPPTY(oBinaryReader);
                    var oConnetctedObjects = oBinaryReader.oConnectedObjects;
                    for(var sKey in oConnetctedObjects) {
                        var oConnectedObject = oConnetctedObjects[sKey];
                        if(oConnectedObject.spid !== null && oConnectedObject.spid !== undefined) {
                            oConnectedObject.setSpid(sObjectId);
                        }
                    }
                    return oPar;
                }
            }
        }
        return null;
    };
    CTiming.prototype.isMainSequenceAtEnd = function() {
        var oRoot = this.getTimingRootNode();
        if(!oRoot) {
            return true;
        }
        var aRootChildren = oRoot.getChildrenTimeNodes();
        var oMainSeq;
        for(var nChild = 0; nChild < aRootChildren.length; ++nChild) {
            if(aRootChildren[nChild].isMainSequence()) {
                oMainSeq = aRootChildren[nChild];
                break;
            }
        }
        if(!oMainSeq) {
            return true;
        }
        return oMainSeq.isAtEnd();
    };
    CTiming.prototype.isSpClickTrigger = function(oSp) {
        var oRoot = this.getTimingRootNode();
        if(!oRoot) {
            return true;
        }
        var aRootChildren = oRoot.getChildrenTimeNodes();
        var sSpId = oSp.Get_Id();
        for(var nChild = 0; nChild < aRootChildren.length; ++nChild) {
            if(aRootChildren[nChild].isInteractiveSeq(sSpId)) {
                return true;
            }
        }
        return false;
    };
    CTiming.prototype.staticCreateNoneEffect = function() {
        return AscFormat.ExecuteNoHistory(function() {
            return CTiming.prototype.createPar(NODE_FILL_HOLD, "indefinite")
        }, this, []);
    };
    CTiming.prototype.getAllAnimEffects = function() {
        if(!this.tnLst) {
            return [];
        }
        var oTmRoot = this.getTimingRootNode();
        if(!oTmRoot) {
            return [];
        }
        return oTmRoot.getAllAnimEffects();
    };
    CTiming.prototype.createTimingRoot = function() {
        var oTnContainer, oCTn;
        this.setTnLst(new CTnLst());
        oTnContainer = new CPar();
        oCTn = this.createCCTn("indefinite", null, null, AscFormat.NODE_TYPE_TMROOT, RESTART_TYPE_NEVER, true);
        oTnContainer.setCTn(oCTn);
        this.tnLst.addToLst(0, oTnContainer);
    };
    CTiming.prototype.checkTimeRoot = function() {
        var oTmRoot = this.getTimingRootNode();
        if(!this.tnLst || !oTmRoot) {
            this.createTimingRoot();
        }
        return this.getTimingRootNode();
    };
    CTiming.prototype.checkMainSequence = function() {
        var oTnContainer, oCTn;
        var oTmRoot = this.checkTimeRoot();
        var oMainSeq = oTmRoot.getChildTimeNodeByType(AscFormat.NODE_TYPE_MAINSEQ);
        if(!oMainSeq) {
            oTnContainer = new CSeq();
            oTnContainer.setConcurrent(true);
            oTnContainer.setNextAc(NEXT_AC_SEEK);
            oCTn = this.createCCTn("indefinite", null, null, AscFormat.NODE_TYPE_MAINSEQ, null, true);
            oTnContainer.setCTn(oCTn);
            oTmRoot.addToChildTnLst(0, oTnContainer);
            var oPrevCondLst = new CCondLst();
            var oCond = new CCond();
            oCond.setEvt(COND_EVNT_ON_PREV);
            oCond.setDelay("0");
            var oTgt = new CTgtEl();
            oCond.setTgtEl(oTgt);
            oPrevCondLst.addToLst(0, oCond);
            oTnContainer.setPrevCondLst(oPrevCondLst);
            var oNextCondLst = new CCondLst();
            oCond = new CCond();
            oCond.setEvt(COND_EVNT_ON_NEXT);
            oCond.setDelay("0");
            oTgt = new CTgtEl();
            oCond.setTgtEl(oTgt);
            oNextCondLst.addToLst(0, oCond);
            oTnContainer.setNextCondLst(oNextCondLst);
            oMainSeq = oTnContainer;
        }
        return oMainSeq;
    };
    CTiming.prototype.checkInteractiveSequence = function(sObjectId) {
        var oTnContainer, oCTn;
        var oTmRoot = this.checkTimeRoot();
        var aSeq = oTmRoot.getChildrenTimeNodes();
        var oInteractiveSeq;
        var oSeq;
        for(var nSeq = 0; nSeq < aSeq.length; ++nSeq) {
            oSeq = aSeq[nSeq];
            if(oSeq.isInteractiveSeq(sObjectId)) {
                oInteractiveSeq = oSeq;
                break;
            } 
        }
        if(!oInteractiveSeq) {
            oTnContainer = new CSeq();
            oTnContainer.setConcurrent(true);
            oTnContainer.setNextAc(NEXT_AC_SEEK);
            oCTn = this.createCCTn(null, NODE_FILL_HOLD, null, AscFormat.NODE_TYPE_INTERACTIVESEQ, RESTART_TYPE_WHEN_NOT_ACTIVE, true, null);
            oTnContainer.setCTn(oCTn);
            oCTn.setEvtFilter("cancelBubble");
            var oStCondLst = new CCondLst();
            var oCond = new CCond();
            oCond.setEvt(COND_EVNT_ON_CLICK);
            oCond.setDelay("0");
            var oTgt = new CTgtEl();
            var oSpTgt = new CSpTgt();
            oSpTgt.setSpid(sObjectId);
            oTgt.setSpTgt(oSpTgt);
            oCond.setTgtEl(oTgt);
            oStCondLst.push(oCond);
            oCTn.setStCondLst(oStCondLst);
            var oNextCondLst = oStCondLst.createDuplicate();
            oTnContainer.setNextCondLst(oNextCondLst);
            var oEndSync = new CCond();
            oEndSync.setEvt(COND_EVNT_END);
            oEndSync.setDelay("0");
            oEndSync.setRtn(RTN_ALL);
            oCTn.setEndSync(oEndSync);
            oTmRoot.pushToChildTnLst(oTnContainer);
            oInteractiveSeq = oTnContainer;
        }
        return oInteractiveSeq;
    };
    CTiming.prototype.getMainSequence = function() {
        var oTmRoot = this.getTimingRootNode();
        if(!oTmRoot) {
            return null;
        }
        return oTmRoot.getChildTimeNodeByType(AscFormat.NODE_TYPE_MAINSEQ);
    };
    CTiming.prototype.addToMainSequence = function(oEffect) {
        var aSeqs = this.getEffectsSequences();
        var aMainSeq;
        if(!aSeqs[0] || aSeqs[0][0] !== null) {
            aMainSeq = [null];
            aSeqs.splice(0, 0, aMainSeq);
        }
        else {
            aMainSeq = aSeqs[0];
        }
        aMainSeq.push(oEffect);
        return this.buildTree(aSeqs);
    };
    CTiming.prototype.addToInteractiveSequence = function(oEffect, sObjectId) {
        var aSeqs = this.getEffectsSequences();
        var aMainSeq;
        for(var nSeq = 0; nSeq < aSeqs.length; ++nSeq) {
            if(aSeqs[nSeq][0] === sObjectId) {
                aMainSeq = aSeqs[nSeq];
                break;
            }
        }
        if(!aMainSeq) {
            aMainSeq = [sObjectId];
            aSeqs.push(aMainSeq);
        }
        aMainSeq.push(oEffect);
        return this.buildTree(aSeqs);
    };
    CTiming.prototype.addAnimationToSelectedObjects = function(nPresetClass, nPresetId, nPresetSubtype) {
        var aSelectedObjects = this.parent.graphicObjects.selectedObjects;
        var aAddedEffects = [];
        for(var nIdx = 0; nIdx < aSelectedObjects.length; ++nIdx) {
            var sObjectId = aSelectedObjects[nIdx].Get_Id();
            var oEffect = this.addEffectToMainSequence(sObjectId, nPresetClass, nPresetId, nPresetSubtype, false);
            if(oEffect) {
                aAddedEffects.push(oEffect);
            }
        }
        return aAddedEffects;
    };
    CTiming.prototype.addAnimation = function(nPresetClass, nPresetId, nPresetSubtype, bReplace) {
        var aAddedEffects = [];
        if(nPresetId === AscFormat.ANIM_PRESET_NONE) {
            this.removeSelectedEffects();
            return aAddedEffects;
        }
        if(!AscFormat.isRealNumber(nPresetClass)
		|| !AscFormat.isRealNumber(nPresetId))
		{
			return aAddedEffects;
		}
        var aSelectedEffects = this.getSelectedEffects();
        var nIdx;
        var nEffectIdx;
        var oEffect;
        var oNewEffect = null;
        var sObjectId;
        var oDrawingsIdMap = {};
        if(bReplace) {
            if(aSelectedEffects.length === 0) {
                return this.addAnimationToSelectedObjects(nPresetClass, nPresetId, nPresetSubtype);
            }
            else {
                var oMapOfObjects = {};
				var aSelectedObjects = this.parent.graphicObjects.selectedObjects;
                var bNeedRemoveExtra = (aSelectedObjects.length > 0);
                var aSeqs = this.getEffectsSequences();
                var aSeq;
                var bNeedRebuild = false;
                for(var nSeq = 0; nSeq < aSeqs.length; ++nSeq) {
                    aSeq = aSeqs[nSeq];
                    for(nEffectIdx = 1; nEffectIdx < aSeq.length; ++nEffectIdx) {
                        oEffect = aSeq[nEffectIdx];
                        if(oEffect.isSelected()) {
                            sObjectId = oEffect.getObjectId();
                            if(bNeedRemoveExtra) {
                                if(oMapOfObjects[sObjectId]) {
                                    aSeq.splice(nEffectIdx, 1);
                                    nEffectIdx--;
                                    continue;
                                }
                                else {
                                    oMapOfObjects[sObjectId] = true;
                                }
                            }
                            oNewEffect = this.createEffect(sObjectId, nPresetClass, nPresetId, nPresetSubtype);
                            if(oNewEffect) {
                                oNewEffect.cTn.setNodeType(oEffect.cTn.nodeType);
                                oNewEffect.cTn.changeDelay(oEffect.cTn.getDelay(true));
                                oNewEffect.select();
                                aSeq[nEffectIdx] = oNewEffect;
                                bNeedRebuild = true;
                                aAddedEffects.push(oNewEffect);
                            }
                        }
                    }
                }
            }
            if(bNeedRebuild) {
                aAddedEffects = this.buildTree(aSeqs);
            }
        }
        else {
            if(aSelectedEffects.length > 0) {
                for(nIdx = 0; nIdx < aSelectedEffects.length; ++nIdx) {
                    oEffect = aSelectedEffects[nIdx];
                    sObjectId = oEffect.getObjectId();
                    if(sObjectId) {
                        if(!oDrawingsIdMap[sObjectId]) {
                            oNewEffect = this.addEffectToMainSequence(sObjectId, nPresetClass, nPresetId, nPresetSubtype, false);
                            if(oNewEffect) {
                                aAddedEffects.push(oNewEffect);
                            }
                            oDrawingsIdMap[sObjectId] = true;
                        }
                    }
                }
            }
            else {
                aAddedEffects = this.addAnimationToSelectedObjects(nPresetClass, nPresetId, nPresetSubtype);
            }
        }
        return aAddedEffects;
    };
    CTiming.prototype.removeSelectedEffects = function() {
        this.removeEffects(this.getSelectedEffects());
    };
    CTiming.prototype.removeEffects = function(aEffectsToRemove) {
        var aSeqs = this.getEffectsSequences();
        var nSeq, nEffect, aSeq, oEffect;
        var nEffectToRemove;
        for(nSeq = aSeqs.length - 1; nSeq > -1; nSeq--) {
            aSeq = aSeqs[nSeq];
            for(nEffect = aSeq.length - 1; nEffect > 0; --nEffect) {
                oEffect = aSeq[nEffect];
                for(nEffectToRemove = aEffectsToRemove.length - 1; nEffectToRemove > -1; --nEffectToRemove) {
                    if(oEffect === aEffectsToRemove[nEffectToRemove]) {
                        aSeq.splice(nEffect, 1);
                        break;
                    }
                }
            }
        }
        this.buildTree(aSeqs);
    };
    CTiming.prototype.addEffectToMainSequence = function(sObjectId, nPresetClass, nPresetId, nPresetSubtype, bReplace) {
     
        if(bReplace) {
            this.removeObjectAnimation(sObjectId);
        }
        var oEffect = this.createEffect(sObjectId, nPresetClass, nPresetId, nPresetSubtype);
        if(!oEffect) {
            return null;
        }
        oEffect = this.addToMainSequence(oEffect)[0];
        oEffect && oEffect.select();
        return oEffect;
    };
    CTiming.prototype.createPar = function(nFill, sDelay) {
        var oPar = new CPar();
        var oCTn = CTiming.prototype.createCCTn(null, nFill, sDelay, null, null, true);
        oPar.setCTn(oCTn);
        return oPar;
    };
    CTiming.prototype.getSelectionRanges = function(aSeqs) {
        var nSeq, nEffect;
        var aRanges = [];
        var aLastRange = null;
        for(nSeq = 0; nSeq < aSeqs.length; ++nSeq) {
            var aSeq = aSeqs[nSeq];
            for(nEffect = 1; nEffect < aSeq.length; ++nEffect) {
                if(aSeq[nEffect].isSelected()) {
                    if(!Array.isArray(aLastRange)) {
                        aLastRange = [[nSeq, nEffect], [nSeq, nEffect]];
                        aRanges.push(aLastRange);
                    }
                    else {
                        aLastRange[1][0] = nSeq;
                        aLastRange[1][1] = nEffect;
                    }
                }
                else {
                    aLastRange = null;
                }
            }
        }
        return aRanges;
    };
    CTiming.prototype.getSequencesForMove = function(bEarlier, bCheckPossibility) {
        var aSeqs = this.getEffectsSequences();
        if(bEarlier && (!aSeqs[0] || aSeqs[0][0] !== null)) {
            aSeqs.splice(0, 0, [null]);
        }
        var aRanges = this.getSelectionRanges(aSeqs);
        var nSeq, aSeq;
        if(aRanges.length !== 1) {
            return bCheckPossibility ? false : null;
        }
        var aRange = aRanges[0];
        var aStart = aRange[0];
        var aEnd = aRange[1];
        var nEffectStart;
        var nEffectEnd;
        var nCount;
        var aEffectsToInsert = [];
        var nPos;
        if(bEarlier) {
            if(aStart[0] === 0) {
                if(aStart[1] === 1) {
                    return bCheckPossibility ? false : null;
                }
            }
            if(bCheckPossibility) {
                return true;
            }
        }
        else {
            if(aEnd[0] === aSeqs.length - 1) {
                if(aEnd[1] === aSeqs[aSeqs.length - 1].length - 1) {
                    return bCheckPossibility ? false : null;
                }
            }
            if(bCheckPossibility) {
                return true;
            }
        }


        var nPosStartEnd;
        var aSeqToInsert;
        if(bEarlier) {
            if(aStart[1] === 1) {
                aSeqToInsert = aSeqs[aStart[0] - 1];
                nPosStartEnd = aSeqToInsert.length;
            }
            else {
                aSeqToInsert = aSeqs[aStart[0]];
                nPosStartEnd = aStart[1] - 1;
            }
        }
        else {
            if(aEnd[1] === aSeqs[aEnd[0]].length - 1) {
                aSeqToInsert = aSeqs[aEnd[0] + 1];
                nPosStartEnd = aSeqToInsert.length - 1;
            }
            else {
                aSeqToInsert = aSeqs[aEnd[0]];
                nPosStartEnd = aSeqToInsert.length - (aEnd[1] + 2);
            }
        }


        for(nSeq = aStart[0]; nSeq <= aEnd[0]; ++nSeq) {
            aSeq = aSeqs[nSeq];
            if(nSeq === aStart[0]) {
                nEffectStart = aStart[1];
            }
            else {
                nEffectStart = 1;
            }
            if(nSeq === aEnd[0]) {
                nEffectEnd = aEnd[1];
            }
            else {
                nEffectEnd = aSeq.length - 1;
            }
            nCount = nEffectEnd - nEffectStart + 1;
            aEffectsToInsert = aEffectsToInsert.concat(aSeq.splice(nEffectStart, nCount));
        }
        if(bEarlier) {
            nPos = nPosStartEnd;
        }
        else {
            nPos = aSeqToInsert.length - nPosStartEnd;
        }
        aSeqToInsert.splice.apply(aSeqToInsert, [nPos, 0].concat(aEffectsToInsert));
        return aSeqs;
    };
    CTiming.prototype.canMoveAnimation = function(bEarlier) {
        return this.getSequencesForMove(bEarlier, true);
    };
    CTiming.prototype.moveAnimation = function(bEarlier) {
        var aSeqs = this.getSequencesForMove(bEarlier, false);
        if(!Array.isArray(aSeqs)) {
            return;
        }
        this.buildTree(aSeqs);
    };
    CTiming.prototype.drawAnimPane = function(oGraphics) {
        if(!this.animPane) {
            this.animPane = new CAnimPane(this);
        }
        this.animPane.recalculate();
        this.animPane.draw(oGraphics);
    };
    CTiming.prototype.getAnimPane = function() {
        if(!this.animPane) {
            this.animPane = new CAnimPane(this);
        }
        return this.animPane;
    };
    CTiming.prototype.onAnimPaneResize = function() {
        this.getAnimPane().onResize();
    };
    CTiming.prototype.onAnimPaneMouseDown = function(e, x, y) {
        this.getAnimPane().onMouseDown(e, x, y);
    };
    CTiming.prototype.onAnimPaneMouseMove = function(e, x, y) {
        this.getAnimPane().onMouseMove(e, x, y);
    };
    CTiming.prototype.onAnimPaneMouseUp = function(e, x, y) {
        this.getAnimPane().onMouseUp(e, x, y);
    };
    CTiming.prototype.onAnimPaneMouseWheel = function(e, deltaY, X, Y) {
        this.getAnimPane().onMouseWheel(e, deltaY, X, Y);
    };
    CTiming.prototype.getRootSequences = function() {
        var oTmRoot = this.getTimingRootNode();
        if(!oTmRoot) {
            return [];
        }
        var aSeqs = [];
        var aChildren = oTmRoot.getChildrenTimeNodes();
        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
            var oChild = aChildren[nChild];
            if(oChild.getObjectType() === AscDFH.historyitem_type_Seq) {
                aSeqs.push(oChild);
            }
        }

        return aSeqs;
    };
    CTiming.prototype.getEffectsSequences = function() {
        var aSequences = [];
        var aAllEffects = this.getAllAnimEffects();
        var aCurSequence = null;
        var oEffect;
        var sSeqId;
        var nEffect;
        for(nEffect = 0; nEffect < aAllEffects.length; ++nEffect) {
            oEffect = aAllEffects[nEffect];
            if(oEffect.isPartOfMainSequence()) {
                sSeqId = null;
            }
            else {
                sSeqId = oEffect.isPartOfInteractiveSeq();
            }
            if(!Array.isArray(aCurSequence) || aCurSequence[0] !== sSeqId) {
                aCurSequence = [sSeqId];
                aSequences.push(aCurSequence);
            }
            aCurSequence.push(oEffect);
        }
        return aSequences;
    };
    CTiming.prototype.buildTree = function(aSequences, bRestedDelayShift) {
        var aCurSequence;
        var oEffect;
        var sSeqId;
        var nEffect;
        var nSeq;
        var oCont1;//containers by depth
        var aAddedEffects = [];
        if(bRestedDelayShift !== false) {
            //substract delay shift from afterEffect nodes
            for(nSeq = 0; nSeq < aSequences.length; ++nSeq) {
               aCurSequence = aSequences[nSeq];
               for(nEffect = 1; nEffect < aCurSequence.length; ++nEffect) {
                   oEffect = aCurSequence[nEffect];
                   oEffect.resetDelayShift();
               }
           }
        }
        this.createTimingRoot();
        var oTmRoot = this.getTimingRootNode();
        if(oTmRoot) {
            oTmRoot.clearChildTnLst();
            var oCTn = oTmRoot.cTn;
            if(oCTn) {
                oTmRoot.setCTn(oCTn.createDuplicate());
                oCTn.setParent(null);
            }
        }
        for(nSeq = 0; nSeq < aSequences.length; ++nSeq) {
            aCurSequence = aSequences[nSeq];
            if(aCurSequence.length > 1) {
                sSeqId = aCurSequence[0];
                if(sSeqId === null) {
                    oCont1 = this.checkMainSequence();
                }
                else {
                    oCont1 = this.checkInteractiveSequence(sSeqId);
                }
                for(nEffect = 1; nEffect < aCurSequence.length; ++nEffect) {
                    oEffect = aCurSequence[nEffect];
                    var oEffectToAdd;
                    if(oEffect.parent) {
                        oEffectToAdd = oEffect.createDuplicate();
                    }
                    else {
                        oEffectToAdd = oEffect;
                        aAddedEffects.push(oEffect);
                    }
                    if(oEffect.selected) {
                        oEffectToAdd.selected = true;
                    }
                    oCont1.addEffectToTheEndOfSeq(oEffectToAdd);
                }
            }
        }
        this.updateNodesIDs();
        return aAddedEffects;
    };
    CTiming.prototype.executeWithCheckDelay = function(fCallback, aEffects) {
        var aDelays = [];
        for(var nEffect = 0; nEffect < aEffects.length; ++nEffect) {
            aDelays.push(aEffects[nEffect].cTn.getDelay(true));
        }
        fCallback();
        for(nEffect = 0; nEffect < aEffects.length; ++nEffect) {
            if(aEffects[nEffect].isAfterEffect()) {
                aEffects[nEffect].cTn.changeDelay(aDelays[nEffect], true);
            }
        }
    };
    CTiming.prototype.setAnimationProperties = function(oPr) {
        var aEffects = this.getSelectedEffects();
        var oCurPr = this.getAnimProperties();
        var nEffect, oEffect;
        var aAllEffects = this.getAllAnimEffects();
        var aEffectsForCheck;
        var aSeqs, aSeq, nSeq;
        if(aEffects.length < 1) {
            return null;
        }

        if(oPr.asc_getDelay() !== oCurPr.asc_getDelay() && AscFormat.isRealNumber(oPr.asc_getDelay())) {
            for(nEffect = aAllEffects.length - 1; nEffect > -1; --nEffect) {
                if(aAllEffects[nEffect].isSelected()) {
                    break;
                }
            }
            aEffectsForCheck = aAllEffects.slice(nEffect + 1);
            this.executeWithCheckDelay(function () {
                for(nEffect = 0; nEffect < aEffects.length; ++nEffect) {
                    oEffect = aEffects[nEffect];
                    oEffect.cTn.changeDelay(oPr.asc_getDelay());
                }
            }, aEffectsForCheck);
        }
        
        if(oPr.asc_getDuration() !== oCurPr.asc_getDuration() && AscFormat.isRealNumber(oPr.asc_getDuration())) {
            for(nEffect = 0; nEffect < aAllEffects.length; ++nEffect) {
                if(aAllEffects[nEffect].isSelected()) {
                    break;
                }
            }
            aEffectsForCheck = aAllEffects.slice(nEffect + 1);
            this.executeWithCheckDelay(function () {
                for(nEffect = 0; nEffect < aEffects.length; ++nEffect) {
                    oEffect = aEffects[nEffect];
                    oEffect.cTn.changeEffectDuration(oPr.asc_getDuration());
                }
            }, aEffectsForCheck);
        }
        if(oPr.asc_getSubtype() !== oCurPr.asc_getSubtype()) {
            for(nEffect = 0; nEffect < aEffects.length; ++nEffect) {
                oEffect = aEffects[nEffect];
                oEffect.cTn.changeSubtype(oPr.asc_getSubtype());
            }
        }
        if(oPr.asc_getRepeatCount() !== oCurPr.asc_getRepeatCount() && AscFormat.isRealNumber(oPr.asc_getRepeatCount())) {
            for(nEffect = 0; nEffect < aEffects.length; ++nEffect) {
                oEffect = aEffects[nEffect];
                oEffect.cTn.changeRepeatCount(oPr.asc_getRepeatCount());
            }
        }

        if(oPr.asc_getRewind() !== oCurPr.asc_getRewind() && AscFormat.isRealBool(oCurPr.asc_getRewind())) {
            for(nEffect = 0; nEffect < aEffects.length; ++nEffect) {
                oEffect = aEffects[nEffect];
                oEffect.cTn.changeRewind(oPr.asc_getRewind());
            }
        }

        
        if(oPr.asc_getStartType() !== oCurPr.asc_getStartType()) {
            aSeqs = this.getEffectsSequences();
            for(nSeq = 0; nSeq < aSeqs.length; ++nSeq) {
                aSeq = aSeqs[nSeq];
                for(nEffect = 1; nEffect < aSeq.length; ++nEffect) {
                    oEffect = aSeq[nEffect];
                    oEffect.resetDelayShift();
                    if(oEffect.isSelected()) {
                        oEffect.cTn.setNodeType(oPr.asc_getStartType());
                    }
                }
            }
            this.buildTree(aSeqs, false);
        }


        if(oPr.asc_getTriggerClickSequence() !== oCurPr.asc_getTriggerClickSequence()
        || oPr.asc_getTriggerObjectClick() !== oCurPr.asc_getTriggerObjectClick()) {
            aSeqs = this.getEffectsSequences();
            
            var sSeqId;
            if(oPr.asc_getTriggerClickSequence() || !oPr.asc_getTriggerObjectClick()) {
                sSeqId = null;
            }
            else {
                var oTimingParent = this.parent;//might be slide, layout, master
                if(!oTimingParent) {
                    return;
                }
                var oCSld = oTimingParent.cSld;
                if(!oCSld) {
                    return;
                }
                var sObjectId;
                
                var oDrawing = oCSld.getObjectByName(oPr.asc_getTriggerObjectClick());
                if(!oDrawing) {
                    return;
                }
                sObjectId = oDrawing.Get_Id();
                sSeqId = sObjectId;
            }
            var aEffectsToInsert = [];
            var sCurSeqId;
            var aTriggerSeq = null;
            for(nSeq = 0; nSeq < aSeqs.length; ++nSeq) {
                aSeq = aSeqs[nSeq];
                sCurSeqId = aSeq[0];
                if(sCurSeqId === sSeqId) {
                    aTriggerSeq = aSeq;
                }
                for(nEffect = aSeq.length - 1; nEffect > 0; --nEffect) {
                    oEffect = aSeq[nEffect];
                    oEffect.resetDelayShift();
                    if(oEffect.isSelected()) {
                        if(sCurSeqId !== sSeqId) {
                            aEffectsToInsert.splice(0, 0, aSeq.splice(nEffect, 1)[0]);
                        }
                    }
                }
            }
            if(!aTriggerSeq) {
                aTriggerSeq = [sSeqId];
                aSeqs.push(aTriggerSeq);
            }
            for(nEffect = 0; nEffect < aEffectsToInsert.length; ++nEffect) {
                aTriggerSeq.push(aEffectsToInsert[nEffect]);
            }
            this.buildTree(aSeqs, false);
        }
    };
    CTiming.prototype.getObjectEffects = function(sObjectId) {
        var aEffects = [];
        
        if(!sObjectId) {
            return aEffects;
        }
        if(!this.tnLst) {
            return aEffects;
        }
        var oTmRoot = this.getTimingRootNode();
        if(!oTmRoot) {
            return aEffects;
        }
        oTmRoot.traverseTimeNodes(function(oNode) {
            if(oNode.isObjectEffect(sObjectId)) {
                aEffects.push(oNode);
            }
        });
        return aEffects;
    };
    CTiming.prototype.getAnimEffect = function(sObjectId) {
        var aEffects = this.getObjectEffects(sObjectId);
        return this.getPropertiesFromEffects(aEffects);
    };
    CTiming.prototype.getSelectedEffects = function() {
        //todo 
        var aEffects = [];
        var aAllEffects = this.getAllAnimEffects();
        for(var nIdx = 0; nIdx < aAllEffects.length; ++nIdx) {
            var oEffect = aAllEffects[nIdx];
            if(oEffect.isSelected()) {
                aEffects.push(oEffect);
            }
        }
        return aEffects;
    };
    CTiming.prototype.getPropertiesFromEffects = function(aEffects) {
        var oResultEffect = null;
        var oEffect;
        for(var nEffect = 0; nEffect < aEffects.length; ++nEffect) {
            oEffect = aEffects[nEffect];
            if(!oResultEffect) {
                oResultEffect = oEffect.createDuplicate();
            }
            oResultEffect.merge(oEffect);
        }
        return oResultEffect;
    };
    CTiming.prototype.getAnimProperties = function() {
        return AscFormat.ExecuteNoHistory(function(){
            var aEffects = this.getSelectedEffects();
            if(aEffects.length === 0) {
                var oSlide = this.parent;
                if(oSlide) {
                    var oGrObjects = oSlide.graphicObjects;
                    if(oGrObjects) {
                        var aSelectedDrawings = oGrObjects.selectedObjects;
                        if(aSelectedDrawings.length > 0) {
                            aEffects.push(this.staticCreateNoneEffect());
                        }
                    }
                }
            }
            return this.getPropertiesFromEffects(aEffects);
        }, this, []);
    };
    CTiming.prototype.printTree = function() {
        var oRoot = this.getTimingRootNode();
        if(oRoot) {
            oRoot.printTree();
        }
    };
    CTiming.prototype.getEffectsForDemo = function() {
        var aEffectsForDemo, aCurEffects;
        var aSelectedEffects = this.getSelectedEffects();
        if(aSelectedEffects.length > 0) {
            aCurEffects = aSelectedEffects;
        }
        else {
            aCurEffects = this.getAllAnimEffects();
        }

        aEffectsForDemo = [];
        for(var nEffect = 0; nEffect < aCurEffects.length; ++nEffect) {
            var oEffect = aCurEffects[nEffect];
            if(oEffect.isPartOfMainSequence()) {
                aEffectsForDemo.push(oEffect);
            }
        }
        if(aEffectsForDemo.length === 0) {
            return null;
        }
        return aEffectsForDemo;
    };
    CTiming.prototype.canStartDemo = function() {
        return this.getEffectsForDemo() !== null;
    };
    CTiming.prototype.createDemoTiming = function() {
        return AscFormat.ExecuteNoHistory(function() {
            if(!this.canStartDemo()) {
                return null;
            }
            var aEffectsForDemo = this.getEffectsForDemo();
            if(!aEffectsForDemo) {
                return null;
            }
            var aSeqs = [];
            var aSeq = [null];
            var oEffect;
            aSeqs.push(aSeq);
            for(var nIdx = 0; nIdx < aEffectsForDemo.length; ++nIdx) {
                oEffect = aEffectsForDemo[nIdx];
                var oCopyEffect = oEffect.createDuplicate();
                oCopyEffect.originalNode = oEffect;
                oCopyEffect.cTn.resetDelayShift();
                if(oCopyEffect.cTn.nodeType === AscFormat.NODE_TYPE_CLICKEFFECT) {
                    oCopyEffect.cTn.setNodeType(nIdx === 0 ? AscFormat.NODE_TYPE_WITHEFFECT : AscFormat.NODE_TYPE_AFTEREFFECT);
                }
                var nRepeatCount = oCopyEffect.asc_getRepeatCount();
                if(nRepeatCount === AscFormat.untilNextSlide || nRepeatCount === AscFormat.untilNextClick) {
                    oCopyEffect.cTn.changeRepeatCount(1000);
                }
                var nDur = oCopyEffect.asc_getDuration();
                if(nDur === AscFormat.untilNextSlide || nDur === AscFormat.untilNextClick) {
                    oCopyEffect.cTn.changeEffectDuration(1000);
                }
                if(AscFormat.isRealNumber(nDur) && nDur < 50) {
                    oCopyEffect.cTn.changeEffectDuration(750);
                }

                oCopyEffect.originalNode = null;
                aSeq.push(oCopyEffect);
            }
            var oTiming = new CTiming();
            oTiming.setParent(this.parent);
            oTiming.buildTree(aSeqs, false);
            return oTiming;
        }, this, []);
    };
    CTiming.prototype.createDemoPlayer = function() {
        if(!this.canStartDemo()) {
            return null;
        }
        return new CDemoAnimPlayer(this.parent);
    };
    CTiming.prototype.onChangeDrawingsSelection = function() {
        var oSlide = this.parent;
        if(!oSlide) {
            return;
        }
        var aSelectedDrawings = oSlide.graphicObjects.selectedObjects;
        var aEffects = this.getAllAnimEffects();
        for(var nEffect = 0; nEffect < aEffects.length; ++nEffect) {
            var oEffect = aEffects[nEffect];
            for(var nDrawing = 0; nDrawing < aSelectedDrawings.length; ++nDrawing) {
                if(oEffect.isObjectEffect(aSelectedDrawings[nDrawing].Get_Id())) {
                    break;
                }
            }
            if(nDrawing < aSelectedDrawings.length) {
                oEffect.select();
            }
            else {
                oEffect.deselect();
            }
        }
    };
    CTiming.prototype.resetSelection = function() {
        var aSelectedEffects = this.getSelectedEffects();
        var aRet = [];
        for(var nEff = 0; nEff < aSelectedEffects.length; ++nEff) {
            aSelectedEffects[nEff].deselect();
        }
        return aRet;
    };
    CTiming.prototype.getSelectionState = function() {
        var aSelectedEffects = this.getSelectedEffects();
        var aRet = [];
        for(var nEff = 0; nEff < aSelectedEffects.length; ++nEff) {
            aRet.push(aSelectedEffects[nEff].Get_Id());
        }
        return aRet;
    };
    CTiming.prototype.setSelectionState = function(aSelected) {
        if(!Array.isArray(aSelected)) {
            this.resetSelection();
            return;
        }
        var aAllEffects = this.getAllAnimEffects();
        var sCurId;
        for(var nEff = 0; nEff < aAllEffects.length; ++nEff) {
            var oEff = aAllEffects[nEff];
            sCurId = oEff.Get_Id();
            for(var nSel = 0; nSel < aSelected.length; ++nSel) {
                if(aSelected[nSel] === sCurId) {
                    break;
                }
            }
            if(nSel < aSelected.length) {
                oEff.select();
            }
            else {
                oEff.deselect();
            }
        }
    };
    CTiming.prototype.getEffectsForLabelsDraw = function() {
        var aResult = [];
        var aAllEffects = this.getAllAnimEffects();
        var aSelectedEffects = [];
        var nEffect, oEffect;
        //draw selected effects after non-selected
        //draw non-selected effects
        for(nEffect = 0; nEffect < aAllEffects.length; ++nEffect) {
            oEffect = aAllEffects[nEffect];
            if(oEffect.isSelected()) {
                aSelectedEffects.push(oEffect);
            }
            else {
                aResult.push(oEffect);
            }
        }
        aResult = aResult.concat(aSelectedEffects);
        return aResult;
    };
    CTiming.prototype.drawEffectsLabels = function(oGraphics) {
        if(oGraphics.IsThumbnail === true || oGraphics.IsDemonstrationMode === true || AscCommon.IsShapeToImageConverter) {
            return;
        }
        var oApi = editor || Asc.editor;
        if(!oApi.isDrawAnimLabels || !oApi.isDrawAnimLabels()) {
            return;
        }
        var aEffectsForDraw = this.getEffectsForLabelsDraw();

        var oContext = oGraphics.m_oContext;
        var sOldFill;
        if(oContext) {
            var dPR = AscCommon.AscBrowser.retinaPixelRatio;
            oContext.font = Math.round(8 * dPR) + "pt Arial";
            oContext.textAlign = "center"; 
        }
        for(var nEffect = 0; nEffect < aEffectsForDraw.length; ++nEffect) {
            aEffectsForDraw[nEffect].drawEffectLabel(oGraphics);
        }
        if(oContext) {
            oContext.fillStyle = sOldFill;
        }
    };
    CTiming.prototype.onMouseDown = function(e, x, y, bHandle) {
        var oApi = editor || Asc.editor;
        if(!oApi.isDrawAnimLabels || !oApi.isDrawAnimLabels()) {
            return bHandle ? false : null;
        }
        var aEffectsForDraw = this.getEffectsForLabelsDraw();
        for(var nEffect = aEffectsForDraw.length - 1; nEffect > -1; --nEffect) {
            var oEffect = aEffectsForDraw[nEffect];
            if(oEffect.hit(x, y)) {
                if(bHandle) {
                    if(e.CtrlKey) {
                        if(oEffect.isSelected()) {
                            oEffect.deselect();
                        }
                        else {
                            oEffect.select();
                        }
                    }
                    else {
                        this.resetSelection();
                        oEffect.select();
                    }
                    return true;
                }
                else {
                    return {cursorType: "default", objectId: "1"};
                }
            }
        }
        return bHandle ? false : null;
    };
    CTiming.prototype.checkCorrect = function() {
        var oRoot;
        if(this.tnLst) {
            if(!this.tnLst.CheckCorrect()) {
                this.setTnLst(null);
                this.setBldLst(null);
                return;
            }
            var aList = this.tnLst.list;
            oRoot = aList[0];
            if(oRoot) {
                var aToRemove = [];
                oRoot.traverseTimeNodes(function(oTimeNode) {
                    if(oTimeNode.getDepth() === 4) {
                        if(!oTimeNode.isCorrect()) {
                            if(oTimeNode.parent) {
                                aToRemove.push(oTimeNode);
                            }
                        }
                    }
                });
                for(var nEffect = aToRemove.length - 1; nEffect > -1; --nEffect) {
                    var oEffect = aToRemove[nEffect];
                    oEffect.parent.onRemoveChild(oEffect);
                }
            }
        }
    };

    changesFactory[AscDFH.historyitem_CommonTimingListAdd] = CChangeContent;
    changesFactory[AscDFH.historyitem_CommonTimingListRemove] = CChangeContent;
    drawingContentChanges[AscDFH.historyitem_CommonTimingListAdd] = function(oClass) {return oClass.list;};
    drawingContentChanges[AscDFH.historyitem_CommonTimingListRemove] = function(oClass) {return oClass.list;};
    function CCommonTimingList() {
        CBaseAnimObject.call(this);
        this.list = [];
    }
    InitClass(CCommonTimingList, CBaseAnimObject, AscDFH.historyitem_type_CommonTimingList);
    CCommonTimingList.prototype.addToLst = function(nIdx, oPr) {
        var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
        History.Add(new CChangeContent(this, AscDFH.historyitem_CommonTimingListAdd, nInsertIdx, [oPr], true));
        this.list.splice(nInsertIdx, 0, oPr);
        this.setParentToChild(oPr);
    };
    CCommonTimingList.prototype.push = function(oPr) {
        this.addToLst(this.getLength(), oPr);
    };
    CCommonTimingList.prototype.splice = function() {
        var nStart = arguments[0];
        var nDeleteCount;
        if(arguments.length > 1) {
            nDeleteCount = arguments[1];
        }
        else {
            nDeleteCount = this.getLength() - nStart;
        }
        var aDeleted = [];
        for(var nIdx = nStart + nDeleteCount - 1; nIdx >= nStart; --nIdx) {
            aDeleted.push(this.removeFromLst(nIdx));
        }
        aDeleted.reverse();
        for(nIdx = arguments.length - 1; nIdx > 1 ; --nIdx) {
            this.addToLst(nStart, arguments[nIdx]);
        }
        return aDeleted;
    };
    CCommonTimingList.prototype.removeFromLst = function(nIdx) {
        if(nIdx > -1 && nIdx < this.list.length) {
            this.list[nIdx].setParent(null);
            History.Add(new CChangeContent(this, AscDFH.historyitem_CommonTimingListRemove, nIdx, [this.list[nIdx]], false));
            return this.list.splice(nIdx, 1)[0];
        }
        return null;
    };
    CCommonTimingList.prototype.clear = function() {
        for(var nIdx = this.list.length - 1; nIdx > -1; --nIdx) {
            this.removeFromLst(nIdx);
        }
    };
    CCommonTimingList.prototype.isEmpty = function() {
        return this.getLength() === 0;
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
    
    CCommonTimingList.prototype.removeChild = function(oChild) {
        if(this.parent) {
            for(var nIdx = this.list.length - 1; nIdx > -1; --nIdx) {
                if(this.list[nIdx] === oChild) {
                    this.removeFromLst(nIdx);
                    return nIdx;
                }
            }
        }
        return -1;
    };
    CCommonTimingList.prototype.onRemoveChild = function(oChild) {
        this.removeChild(oChild);
        if(this.parent) {
            if(this.list.length === 0) {
                this.parent.onRemoveChild(this);
            }
        }
    };
    CCommonTimingList.prototype.getLength = function() {
        return this.list.length;
    };
    CCommonTimingList.prototype.getTimeNodeByType = function(nType) {
        for(var nNode = 0; nNode < this.list.length; ++nNode) {
            if(this.list[nNode].getNodeType() === nType) {
                return this.list[nNode];
            }
        }
        return null;
    };
    CCommonTimingList.prototype.getLast = function(nType) {
        if(this.list.length > 0) {
            return this.list[this.list.length - 1];
        }
        return null;
    };
    CCommonTimingList.prototype.getChildIdx = function(oChild) {
        for(var nIdx = this.list.length - 1; nIdx > -1; --nIdx) {
            if(this.list[nIdx] === oChild) {
                return nIdx;
            }
        }
        return -1;
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
    CCondLst.prototype.createComplexTrigger = function(oPlayer) {
        var oComplexTrigger = new CAnimComplexTrigger();
        for(var nCond = 0; nCond < this.list.length; ++nCond) {
            this.list[nCond].fillTrigger(oPlayer, oComplexTrigger)
        }
        return oComplexTrigger;
    };
    CCondLst.prototype.getSpClick = function() {
        if(this.list.length === 1) {
            var oCond = this.list[0];
            if(oCond) {
                if(oCond.evt === COND_EVNT_ON_CLICK) {
                    return oCond.getTargetObjectId();
                }
            }
        }
        return null;
    };
    CCondLst.prototype.isSpClick = function(sSpId) {
        return this.getSpClick() === sSpId;
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
    CChildTnLst.prototype.Refresh_RecalcData = function(oData) {
        this.Refresh_RecalcData2();
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
    CTnLst.prototype.CheckCorrect = function() {
        var aList = this.list;
        if(aList.length !== 1) {
            return false;
        }
        else {
            var oRoot = aList[0];
            var oAttr = oRoot.getAttributesObject();
            if(!oAttr || oAttr.nodeType !== AscFormat.NODE_TYPE_TMROOT) {
                return false;
            }
        }
        return true;
    };

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
        CBaseAnimObject.call(this);
        this.spid = null;
    }
    InitClass(CObjectTarget, CBaseAnimObject, AscDFH.historyitem_type_ObjectTarget);
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
            oCopy.setBldSub(this.bldSub.createDuplicate(oIdMap));
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
        CBaseAnimObject.call(this);
        this.chart = null;
        this.animBg = null;
        this.bldChart = null;
        this.bldDgm = null;
        this.rev = null;
    }
    InitClass(CBldSub, CBaseAnimObject, AscDFH.historyitem_type_BldSub);
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
        CBaseAnimObject.call(this);
        this.dir = null;
    }
    InitClass(CDirTransition, CBaseAnimObject, AscDFH.historyitem_type_DirTransition);
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
        CBaseAnimObject.call(this);
        this.thruBlk = null;
    }
    InitClass(COptionalBlackTransition, CBaseAnimObject, AscDFH.historyitem_type_OptBlackTransition);
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
        CBaseAnimObject.call(this);
        this.dgmId = null;
        this.dgmBuildStep = null;
        this.chartBuildStep = null;
        this.seriesIdx = null;
        this.categoryIdx = null;
    }
    InitClass(CGraphicEl, CBaseAnimObject, AscDFH.historyitem_type_GraphicEl);
    CGraphicEl.prototype.setDgmId = function(pr, pReader) {
        if(pReader) {
            pReader.AddConnectedObject(this);
        }
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_GraphicElDgmId, this.dgmId, pr));
        this.dgmId = pr;
    };
    CGraphicEl.prototype.assignConnection = function(oObjectsMap) {
        if(AscCommon.isRealObject(oObjectsMap[this.spid])){
            this.setSpid(oObjectsMap[this.spid].Id);
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
    drawingsChangesMap[AscDFH.historyitem_IndexRgEnd] = function(oClass, value) {oClass.end = value;};
    function CIndexRg() {//charrg, pRg
        CBaseAnimObject.call(this);
        this.st = null;
        this.end = null;
    }
    InitClass(CIndexRg, CBaseAnimObject, AscDFH.historyitem_type_IndexRg);
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
        CBaseAnimObject.call(this);
        this.lvl = null;
        this.tnLst = null
    }
    InitClass(CTmpl, CBaseAnimObject, AscDFH.historyitem_type_Tmpl);
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


    var VALUE_TYPE_NUM = 0;
    var VALUE_TYPE_CLR = 1;
    var VALUE_TYPE_STR = 2;


    var CALCMODE_DISCRETE = 0;
    var CALCMODE_LIN = 1;
    var CALCMODE_FMLA = 2;
    function CAnim() {
        CTimeNodeBase.call(this);
        this.cBhvr = null;
        this.tavLst = null;
        this.by = null;
        this.calcmode = null;
        this.from = null;
        this.to = null;
        this.valueType = null;
    }
    InitClass(CAnim, CTimeNodeBase, AscDFH.historyitem_type_Anim);
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
    CAnim.prototype.getValueType = function() {
        if(this.valueType === null) {
            return VALUE_TYPE_NUM;
        }
        return this.valueType;
    };
    CAnim.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
            return;
        }
        if(this.checkRemoveAtEnd()) {
            return;
        }
        var aAttributes = this.getAttributes();
        if(aAttributes.length < 1) {
            return;
        }
        var oFirstAttribute = aAttributes[0];
        var sAnimAttrName = oFirstAttribute.text;
        if(!(typeof sAnimAttrName === "string") || sAnimAttrName.length === 0) {
            return;
        }
        var val = null;
        var sFmla = null;
        var fRelTime = this.getRelativeTime(nElapsedTime);
        var nValueType = this.getValueType();
        var oVarMap;
        var oFirstTav;
        var oSecondTav;
        var oTav;
        var fTimeInsideInterval;
        if(this.tavLst) {
            var aTav = this.tavLst.list;
            if(aTav.length > 0) {
                var nTav = -1;
                if(fRelTime >= aTav[aTav.length - 1].getTime()) {
                    nTav = aTav.length - 1;
                }
                else if(fRelTime <= aTav[0].getTime()) {
                    nTav = 0;
                    sFmla = aTav[0].fmla;
                }
                if(nTav > -1) {
                    oTav = aTav[nTav];
                    if(aTav[nTav - 1]) {
                        sFmla = aTav[nTav - 1].fmla;
                    }
                    oFirstTav = oTav;
                    oSecondTav = oTav;
                    fTimeInsideInterval = 0;
                    if(nTav === 0) {
                        if(aTav[nTav + 1] && AscFormat.fApproxEqual(aTav[nTav + 1].getTime(), oTav.getTime())) {
                            oSecondTav = aTav[nTav + 1];
                            sFmla = oTav.fmla;
                            fTimeInsideInterval = (fRelTime) / (oSecondTav.getTime());
                        }
                    }
                    val = this.calculateBetweenTwoVals(oFirstTav.val, oSecondTav.val, fTimeInsideInterval, oAttributes);
                }
                else {
                    for(nTav = 1; nTav < aTav.length; ++nTav) {
                        if(fRelTime >= aTav[nTav - 1].getTime() && fRelTime <= aTav[nTav].getTime()) {
                            break;
                        }
                    }
                    if(nTav < aTav.length) {
                        var nCalcMode = this.getCalcMode();
                        if(nCalcMode === CALCMODE_DISCRETE) {
                            if(AscFormat.fApproxEqual(fRelTime, aTav[nTav].getTime())) {
                                oTav = aTav[nTav];
                            }
                            else {
                                oTav = aTav[nTav - 1];
                            }
                            val = this.calculateBetweenTwoVals(oTav.val, oTav.val, 0, oAttributes);
                            if(aTav[nTav - 1]) {
                                sFmla = aTav[nTav - 1].fmla;
                            }
                        }
                        else {
                            if(AscFormat.fApproxEqual(fRelTime, aTav[nTav].getTime())) {
                                oTav = aTav[nTav];
                                val = this.calculateBetweenTwoVals(oTav.val, oTav.val, 0, oAttributes);
                                if(aTav[nTav - 1]) {
                                    sFmla = aTav[nTav - 1].fmla;
                                }
                            }
                            else {
                                oFirstTav = aTav[nTav - 1];
                                oSecondTav = aTav[nTav];
                                sFmla = oFirstTav.fmla;
                                fTimeInsideInterval = (fRelTime - aTav[nTav - 1].getTime()) / (aTav[nTav].getTime() - aTav[nTav - 1].getTime());
                                val = this.calculateBetweenTwoVals(oFirstTav.val, oSecondTav.val, fTimeInsideInterval, oAttributes);
                            }
                        }
                    }
                }
                if(val !== null) {
                    if(sFmla) {
                        oVarMap = this.getVarMapForFmla(oAttributes);
                        oVarMap["$"] = val;
                        var fFmlaResult = this.getFormulaResult(sFmla, oVarMap);
                        if(fFmlaResult !== null) {
                            oAttributes[oFirstAttribute.text] = fFmlaResult;
                        }
                    }
                    else {
                        oAttributes[oFirstAttribute.text] = val;
                    }
                }
            }
        }
        else {
            if(this.from !== null && this.to !== null && this.by === null ||
            this.from !== null && this.to === null && this.by !== null ||
            this.from === null && this.to !== null && this.by === null ||
            this.from === null && this.to === null && this.by !== null) {
                if(nValueType === VALUE_TYPE_NUM) {
                    oVarMap = this.getVarMapForFmla();
                    var fFrom = null, fTo = null, fBy = null;
                    if(this.from !== null) {
                        fFrom = this.getFormulaResult(this.from, oVarMap);
                        if(fFrom === null) {
                            return;
                        }
                    }
                    if(this.to !== null) {
                        fTo = this.getFormulaResult(this.to, oVarMap);
                        if(fTo === null) {
                            return;
                        }
                    }
                    if(this.by !== null) {
                        fBy = this.getFormulaResult(this.by, oVarMap);
                        if(fBy === null) {
                            return;
                        }
                    }
                    if(this.from !== null && this.to !== null && this.by === null) {
                        if(fFrom !== null && fTo !== null && fBy === null) {
                            oAttributes[sAnimAttrName] =  this.getAnimatedVal(fRelTime, fFrom, fTo);
                        }
                    }
                    else if(this.from !== null && this.to === null && this.by !== null) {
                        if(fFrom !== null && fTo === null && fBy !== null) {
                            oAttributes[sAnimAttrName] =  this.getAnimatedVal(fRelTime, fFrom, fFrom + fBy);
                        }
                    }
                    else if(this.from === null && this.to !== null && this.by === null) {
                        if(fFrom === null && fTo !== null && fBy === null) {
                            oAttributes[sAnimAttrName] =  this.getAnimatedVal(fRelTime, 0.0, fTo);
                        }
                    }
                    else if(this.from === null && this.to === null && this.by !== null) {
                        if(fFrom === null && fTo === null && fBy !== null) {
                            var fStartVal = AscFormat.isRealNumber(oVarMap[sAnimAttrName]) ? oVarMap[sAnimAttrName] : 0;
                            oAttributes[sAnimAttrName] =  this.getAnimatedVal(fRelTime, fStartVal, fStartVal + fBy);
                        }
                    }
                }
                else if(nValueType === VALUE_TYPE_CLR) {
                    //TODO: implement
                }
                else if(nValueType === VALUE_TYPE_STR) {
                    //TODO: implement
                }
            }
        }
    };
    CAnim.prototype.getVarMapForFmla = function(oAttributes) {
        return {
            "#ppt_x": this.getOrigAttrVal("ppt_x"),
            "#ppt_y": this.getOrigAttrVal("ppt_y"),
            "#ppt_w": this.getOrigAttrVal("ppt_w"),
            "#ppt_h": this.getOrigAttrVal("ppt_h"),
            "ppt_x": oAttributes && AscFormat.isRealNumber(oAttributes["ppt_x"]) ? oAttributes["ppt_x"] : this.getOrigAttrVal("ppt_x"),
            "ppt_y": oAttributes && AscFormat.isRealNumber(oAttributes["ppt_y"]) ? oAttributes["ppt_y"] : this.getOrigAttrVal("ppt_y"),
            "ppt_w": oAttributes && AscFormat.isRealNumber(oAttributes["ppt_w"]) ? oAttributes["ppt_w"] : this.getOrigAttrVal("ppt_w"),
            "ppt_h": oAttributes && AscFormat.isRealNumber(oAttributes["ppt_h"]) ? oAttributes["ppt_h"] : this.getOrigAttrVal("ppt_h"),
            "ppt_x_no_attr": !(oAttributes && AscFormat.isRealNumber(oAttributes["ppt_x"])),
            "ppt_y_no_attr": !(oAttributes && AscFormat.isRealNumber(oAttributes["ppt_y"])),
            "ppt_w_no_attr": !(oAttributes && AscFormat.isRealNumber(oAttributes["ppt_w"])),
            "ppt_h_no_attr": !(oAttributes && AscFormat.isRealNumber(oAttributes["ppt_h"]))
        }
    };
    CAnim.prototype.getFormulaResult = function(sFormula, oVarMap) {
        return (new CFormulaParser(sFormula, oVarMap)).getResult();
    };
    CAnim.prototype.calculateBetweenTwoVals = function(oVal1, oVal2, fRelTime, oAttributes) {
        if(!oVal1 || !oVal2) {
            return null;
        }
        if(oVal1.isClr() !== oVal2.isClr()) {
            return null;
        }
        if(oVal1.isClr()) {
            return this.getAnimatedClr(fRelTime, oVal1.clrVal, oVal2.clrVal);
        }
        var oVarMap;

        var fVal1 = null;
        var fVal2 = null;

        if(oVal1.isFlt()) {
            fVal1 = oVal1.fltVal;
        }
        if(oVal1.isInt()) {
            fVal1 = oVal1.intVal;
        }
        if(oVal1.isStr()) {
            var sStrVal1 = oVal1.getVal();
            if(sStrVal1 === "hidden" || sStrVal1 === "visible") {
                return sStrVal1;
            }
            oVarMap = this.getVarMapForFmla(oAttributes);
            oVarMap["$"] = fRelTime;
            fVal1 = this.getFormulaResult(sStrVal1, oVarMap);
        }
        if(!AscFormat.isRealNumber(fVal1)) {
            return null;
        }

        if(oVal2.isFlt()) {
            fVal2 = oVal2.fltVal;
        }
        if(oVal2.isInt()) {
            fVal2 = oVal2.intVal;
        }
        if(oVal2.isStr()) {
            oVarMap = this.getVarMapForFmla(oAttributes);
            oVarMap["$"] = fRelTime;
            var sStrVal2 = oVal2.getVal();
            fVal2 = this.getFormulaResult(sStrVal2, oVarMap);
        }
        if(!AscFormat.isRealNumber(fVal2)) {
            return null;
        }
        return this.getAnimatedVal(fRelTime, fVal1, fVal2);
    };
    CAnim.prototype.getCalcMode = function() {
        if(this.calcmode === null) {
            return CALCMODE_LIN;
        }
        return this.calcmode;
    };
    CAnim.prototype.create = function(nCalcMode, nValueType, aAttrNames, aTavs, sObjectId, sDur, sDelay, nAccel) {
        this.setCalcmode(nCalcMode);
        this.setValueType(nValueType);
        var oBhvr = new CCBhvr();
        var oCTn = this.createCCTn(sDur, NODE_FILL_HOLD, sDelay, null, null, null, nAccel);
        oBhvr.setCTn(oCTn);
        oBhvr.createAttrNameLst(aAttrNames);
        this.setTavLst(this.createTavLst(aTavs));
        this.setCBhvr(oBhvr);
    };
    CAnim.prototype.createTav = function(sTm, sFmla, bBoolVal, oClrVal, fFltVal, nIntVal, sStrVal) {
        var oTav = new CTav();
        if(sTm !== null) {
            oTav.setTm(sTm);
        }
        if(sFmla !== null && sFmla !== undefined) {
            oTav.setFmla(sFmla);
        }
        var oAnimVariant = new CAnimVariant();
        if(bBoolVal !== null && bBoolVal !== undefined) {
            oAnimVariant.setBoolVal(bBoolVal);
        }
        else if(oClrVal !== null && oClrVal !== undefined) {
            oAnimVariant.setClrVal(oClrVal);
        }
        else if(fFltVal !== null && fFltVal !== undefined) {
            oAnimVariant.setFltVal(fFltVal);
        }
        else if(nIntVal !== null && nIntVal !== undefined) {
            oAnimVariant.setIntVal(nIntVal);
        }
        else if(sStrVal !== null && sStrVal !== undefined) {
            oAnimVariant.setStrVal(sStrVal);
        }
        oTav.setVal(oAnimVariant);

        return oTav;
    };
    CAnim.prototype.createTavFromObject = function(oTav) {
        return this.createTav(oTav.tm, oTav.fmla, oTav.boolVal, oTav.clrVal, oTav.fltVal, oTav.intVal, oTav.strVal);
    };
    CAnim.prototype.createTavLst = function(aTavs) {
        var oTavLst = new CTavLst();
        for(var nTav = 0; nTav < aTavs.length; ++nTav) {
            oTavLst.push(aTavs[nTav]);
        }
        return oTavLst;
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
        CBaseAnimObject.call(this);
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
    InitClass(CCBhvr, CBaseAnimObject, AscDFH.historyitem_type_CBhvr);
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
    CCBhvr.prototype.getTargetObjectId = function() {
        if(this.tgtEl) {
            return this.tgtEl.getSpId();
        }
        return null;
    };
    CCBhvr.prototype.getAttributes = function() {
        if(!this.attrNameLst) {
            return [];
        }
        return this.attrNameLst.list;
    };
    CCBhvr.prototype.createAttrNameLst = function(aAttrNames) {
        this.setAttrNameLst(new CAttrNameLst());
        for(var nName = 0; nName < aAttrNames.length; ++nName) {
            var oAttrName = new CAttrName();
            oAttrName.setText(aAttrNames[nName]);
            this.attrNameLst.push(oAttrName);
        }
    };


    var RESTART_TYPE_ALWAYS = 0;
    var RESTART_TYPE_NEVER = 1;
    var RESTART_TYPE_WHEN_NOT_ACTIVE = 2;

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
    changesFactory[AscDFH.historyitem_CTnRepeatCount] = CChangeString;
    changesFactory[AscDFH.historyitem_CTnRepeatDur] = CChangeString;
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
        CBaseAnimObject.call(this);
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
    InitClass(CCTn, CBaseAnimObject, AscDFH.historyitem_type_CTn);
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
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CTnRepeatCount, this.repeatCount, pr));
        this.repeatCount = pr;
    };
    CCTn.prototype.setRepeatDur = function(pr) {
        oHistory.Add(new CChangeString(this, AscDFH.historyitem_CTnRepeatDur, this.repeatDur, pr));
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
        if(AscCommon.isRealObject(this.childTnLst)) {
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
    CCTn.prototype.merge = function(oOther) {
        this.childTnLst = undefined;
        this.endCondLst = this.checkEqualChild(this.endCondLst, oOther.endCondLst);
        this.endSync = this.checkEqualChild(this.endSync, oOther.endSync);
        this.iterate = this.checkEqualChild(this.iterate, oOther.iterate);
        this.stCondLst = this.checkEqualChild(this.stCondLst, oOther.stCondLst);
        this.subTnLst = this.checkEqualChild(this.subTnLst, oOther.subTnLst);
        this.accel = this.checkEqualChild(this.accel, oOther.accel);
        this.afterEffect = this.checkEqualChild(this.afterEffect, oOther.afterEffect);
        this.autoRev = this.checkEqualChild(this.autoRev, oOther.autoRev);
        this.bldLvl = this.checkEqualChild(this.bldLvl, oOther.bldLvl);
        this.decel = this.checkEqualChild(this.decel, oOther.decel);
        this.display = this.checkEqualChild(this.display, oOther.display);
        this.dur = this.checkEqualChild(this.dur, oOther.dur);
        this.evtFilter = this.checkEqualChild(this.evtFilter, oOther.evtFilter);
        this.fill = this.checkEqualChild(this.fill, oOther.fill);
        this.grpId = this.checkEqualChild(this.grpId, oOther.grpId);
        this.id = this.checkEqualChild(this.id, oOther.id);
        this.masterRel = this.checkEqualChild(this.masterRel, oOther.masterRel);
        this.nodePh = this.checkEqualChild(this.nodePh, oOther.nodePh);
        this.nodeType = this.checkEqualChild(this.nodeType, oOther.nodeType);
        this.presetClass = this.checkEqualChild(this.presetClass, oOther.presetClass);
        this.presetID = this.checkEqualChild(this.presetID, oOther.presetID);
        this.presetSubtype = this.checkEqualChild(this.presetSubtype, oOther.presetSubtype);
        this.repeatCount = this.checkEqualChild(this.repeatCount, oOther.repeatCount);
        this.repeatDur = this.checkEqualChild(this.repeatDur, oOther.repeatDur);
        this.restart = this.checkEqualChild(this.restart, oOther.restart);
        this.spd = this.checkEqualChild(this.spd, oOther.spd);
        this.syncBehavior = this.checkEqualChild(this.syncBehavior, oOther.syncBehavior);
        this.tmFilter = this.checkEqualChild(this.tmFilter, oOther.tmFilter);
    };
    CCTn.prototype.createStCondLst = function() {
        this.setStCondLst(new CCondLst());
    };
    CCTn.prototype.createStCondLstWithDelay = function(sDelay) {
        this.createStCondLst();
        var oCond = new CCond();
        oCond.setDelay(sDelay);
        this.stCondLst.addToLst(0, oCond);
    };
    CCTn.prototype.isAnimEffect = function() {
        if(this.presetID !== null && this.presetClass !== null) {
            return true;
        }
        return false;
    };
    CCTn.prototype.getTimeNodeByType = function(nType) {
        if(this.childTnLst) {
            return this.childTnLst.getTimeNodeByType(nType);
        }
        return null;
    };
    CCTn.prototype.addToChildTnLst = function(nIdx, oNode) {
        if(this.childTnLst) {
            this.childTnLst.addToLst(nIdx, oNode);
        }
    };
    CCTn.prototype.pushToChildTnLst = function(oNode) {
        if(this.childTnLst) {
            this.childTnLst.push(oNode);
        }
    };
    CCTn.prototype.clearChildTnLst = function() {
        if(this.childTnLst) {
            this.childTnLst.clear();
        }
    };
    CCTn.prototype.createChildTnLst = function() {
        if(!this.childTnLst) {
            this.setChildTnLst(new CChildTnLst());
        }
    };
    CCTn.prototype.getDelayShift = function() {
        if(this.nodeType === AscFormat.NODE_TYPE_AFTEREFFECT || 
            this.nodeType === AscFormat.NODE_TYPE_WITHEFFECT) {
            var oPrev = this.parent.getPreviousEffect();
            if(oPrev && oPrev.cTn) {
                var nShift =  oPrev.cTn.getDelay(false);
                if(this.nodeType === AscFormat.NODE_TYPE_AFTEREFFECT) {
                    nShift += oPrev.cTn.getEffectDuration()
                }
                return nShift;
            }
        }
        return 0;
    };
    CCTn.prototype.getDelay = function(bUseDelayShift) {
        var delay = 0;
        var aConds;
        if(this.stCondLst) {
            aConds = this.stCondLst.list;
            for(var nCond = 0; nCond < aConds.length; ++nCond) {
                if(aConds[nCond].delay !== null) {
                    delay = aConds[nCond].getDelayTime().getVal();
                    break;
                }
            }
        }
        var nDelayShift;
        if(bUseDelayShift === false) {
            nDelayShift = 0;
        }
        else {
            nDelayShift = this.getDelayShift();
        }
        delay = Math.max(0, delay - nDelayShift);
        return delay;
    };
    CCTn.prototype.changeDelay = function(nDelay, bUseDelayShift) {
        var nDelayShift;
        if(bUseDelayShift === false) {
            nDelayShift = 0;
        }
        else {
            nDelayShift = this.getDelayShift();
        }
        var nNewDelay = ((Math.max(0, nDelay + nDelayShift) + 0.5) >> 0);
        var sNewDelay = nNewDelay + "";
        var aConds;
        if(!this.stCondLst) {
            this.setStCondLst(new CCondLst());
        }
        aConds = this.stCondLst.list;
        for(var nCond = 0; nCond < aConds.length; ++nCond) {
            if(aConds[nCond].delay !== null) {
                return aConds[nCond].setDelay(sNewDelay);
            }
        }
        var oCond = new CCond();
        this.stCondLst.push(oCond);
        oCond.setDelay(sNewDelay);
    };
    CCTn.prototype.resetDelayShift = function() {
        var nDelayShift = this.getDelayShift();
        if(nDelayShift > 0) {
            this.changeDelay(this.getDelay(false) - this.getDelayShift(), false);
        }
    };
    CCTn.prototype.setDelayShift = function() {
        var nDelayShift = this.getDelayShift();
        if(nDelayShift > 0) {
            this.changeDelay(this.getDelay(false), true);
        }
    };
    CCTn.prototype.getEffectDuration = function() {
        var nDur = 0;
        var aChildren = this.childTnLst && this.childTnLst.list;
        if(aChildren) {
            for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                var oChild = aChildren[nChild];
                var oDur = oChild.getDur();
                if(oDur.isSpecified()) {
					
					var oAttr = oChild.getAttributesObject();
					var nDelay = oAttr.getDelay(false);
                    nDur = Math.max(nDur, oDur.getVal() + nDelay);
                }
            }
        }
        return nDur;
    };
    CCTn.prototype.changeEffectDuration = function(v) {
        var dOldV = this.getEffectDuration();
        var dCoef = null;
        var v_ = Math.max(10, v);
        if(dOldV > 0) {
            dCoef = v_/dOldV;
            if(AscFormat.fApproxEqual(dCoef, 1.0)) {
                return;
            }
        }
        var bIsIndefinite = (v === AscFormat.untilNextSlide || v === AscFormat.untilNextClick);
        var aChildren = this.childTnLst && this.childTnLst.list;
        if(aChildren) {
            for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                var oChild = aChildren[nChild];
                var oDur = oChild.getDur();
                if(oDur.isSpecified()) {
                    var oAttr = oChild.getAttributesObject();
                    var nDelay = oAttr.getDelay(false);
                    if(bIsIndefinite) {
                        oAttr.setDur("indefinite");
                    }
                    else if(dCoef !== null) {
                        oAttr.setDur((oDur.getVal()*dCoef + 0.5 >> 0) + "");
                        if(AscFormat.isRealNumber(nDelay) && nDelay !== 0) {
                            oAttr.changeDelay(nDelay*dCoef);
                        }
                    }
                    else {
                        oAttr.setDur(v_ + "");
                    }

                }
            }
        }
        if(bIsIndefinite) {
            this.changeRepeatCount(v);
        }
    };
    CCTn.prototype.changeRepeatCount = function(v) {
        var oAttrObject = this;
        if(v !== AscFormat.untilNextSlide && v !== AscFormat.untilNextClick) {
            oAttrObject.setRepeatCount(v + "");
        }
        else {
            oAttrObject.setRepeatCount("indefinite");
        }
        if(v === AscFormat.untilNextClick) {
            if(!oAttrObject.endCondLst) {
                oAttrObject.setEndCondLst(new CCondLst()) ;
            }
            oAttrObject.endCondLst.clear();
            var oCond = new CCond();
            oCond = new CCond();
            oCond.setEvt(COND_EVNT_ON_NEXT);
            oCond.setDelay("0");
            var oTgt = new CTgtEl();
            oCond.setTgtEl(oTgt);
            oAttrObject.endCondLst.push(oCond);
        }
        else {
            if(oAttrObject && oAttrObject.endCondLst) {
                oAttrObject.setEndCondLst(null);
            }
        }
    };
    CCTn.prototype.changeRewind = function(v) {
        this.setFill(v === true ? NODE_FILL_REMOVE : NODE_FILL_HOLD);
    };
    CCTn.prototype.getObjectId = function(v) {
        var sObjectId = null;
        this.traverse(function(oChild) {
            if(oChild.isTimeNode() && (sObjectId = oChild.getTargetObjectId())) {
                return true;
            }
            return false
        });
        return sObjectId;
    };
    CCTn.prototype.changeSubtype = function(v) {
        if(AscFormat.isRealNumber(v)) {
            var sObjectId = this.getObjectId();
            if(sObjectId !== null) {
                var oNewEffect = CTiming.prototype.createEffect(sObjectId, this.presetClass, this.presetID, v);
                if(oNewEffect) {
                    var oNewCTn = oNewEffect.cTn;
                    if(this.getEffectDuration() !== oNewCTn.getEffectDuration()) {
                        oNewCTn.changeEffectDuration(this.getEffectDuration());
                    }
                    this.setPresetSubtype(v);
                    this.childTnLst.clear();
                    oNewCTn.childTnLst.fillObject(this.childTnLst, {});
                }
            }
        }
    };
    CCTn.prototype.changeStartType = function(v) {
        if(this.nodeType === v) {
            return;
        }
        var oEffectNode = this.parent;
        if(!oEffectNode) {
            return;
        }
        var oCurPar2Lvl, oCurPar3Lvl, oCurMainSeq;
        var nIdx2, nIdx3, nMainIdx;
        oCurPar3Lvl = oEffectNode.getParentTimeNode();
        if(!oCurPar3Lvl) {
            return;
        }
        nIdx3 = oCurPar3Lvl.getChildNodeIdx(oEffectNode);
        oCurPar2Lvl = oCurPar3Lvl.getParentTimeNode();
        if(!oCurPar2Lvl) {
            return;
        }
        nIdx2 = oCurPar2Lvl.getChildNodeIdx(oCurPar3Lvl);
        oCurMainSeq = oCurPar2Lvl.getParentTimeNode();
        if(!oCurMainSeq) {
            return;
        }
        nMainIdx = oCurMainSeq.getChildNodeIdx(oCurPar2Lvl);

        var oPar2Lvl, oPar3Lvl;
        var aWithEffects, aAfterEffects;
        if(v === AscFormat.NODE_TYPE_CLICKEFFECT) {
            oEffectNode.cTn.setNodeType(AscFormat.NODE_TYPE_CLICKEFFECT);
            oPar3Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "0");
            aWithEffects = oCurPar3Lvl.splice(nIdx3);
            oPar3Lvl.addEffects(0, aWithEffects);
            oPar2Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "indefinite");
            oPar2Lvl.pushToChildTnLst(oPar3Lvl);
            oPar2Lvl.addEffects(1, oCurPar2Lvl.splice(nIdx2 + 1));
            oCurMainSeq.splice(nMainIdx + 1, 0, oPar2Lvl);
        }
        else if(v === AscFormat.NODE_TYPE_WITHEFFECT) {
            oEffectNode.cTn.setNodeType(AscFormat.NODE_TYPE_WITHEFFECT);
            if(nIdx3 === 0) {
                if(nIdx2 === 0) {
                    if(nMainIdx === 0) {
                        //do nothing: no previous animation
                    }
                    else {
                        oCurPar3Lvl.splice(nIdx3);
                        oPar2Lvl = oCurMainSeq.getChildNode(nMainIdx - 1);
                        oPar3Lvl = oPar2Lvl.getChildNode(oPar2Lvl.getChildrenCount() - 1);
                        if(!oPar3Lvl) {
                            oPar3Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "0");
                            oPar2Lvl.pushToChildTnLst(oPar3Lvl);
                        }
                        oPar3Lvl.pushToChildTnLst(oEffectNode);
                    }
                }
                else {
                    oCurPar3Lvl.splice(nIdx3);
                    oPar3Lvl = oCurPar2Lvl.getChildNode(nIdx2 - 1);
                    oPar3Lvl.pushToChildTnLst(oEffectNode);
                }
            }
        }
        else if(v === AscFormat.NODE_TYPE_AFTEREFFECT) {
            oEffectNode.cTn.setNodeType(AscFormat.NODE_TYPE_AFTEREFFECT);
            if(nIdx3 === 0) {
                if(nIdx2 > 0) {
                    //do nothing
                }
                else {
                    if(nMainIdx > 0) {
                        oPar3Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "0");
                        oPar3Lvl.addEffects(0, oCurPar3Lvl.splice(nIdx3));
                        oPar2Lvl = oCurMainSeq.getChildNode(nMainIdx - 1);
                        oPar2Lvl.pushToChildTnLst(oPar3Lvl);
                    }
                }
            }
            else {
                oPar3Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "0");
                oPar3Lvl.addEffects(0, oCurPar3Lvl.splice(nIdx3));
                oPar2Lvl = oCurPar2Lvl;
                oPar2Lvl.splice(nIdx2 + 1, 0, oPar3Lvl);
            }
            oEffectNode.setDelayShift();
        }
        if(oCurPar3Lvl.getChildrenCount() === 0) {
            oCurPar3Lvl.parent.onRemoveChild(oCurPar3Lvl);
        }
        var oTiming = oEffectNode.getTiming();
        if(oTiming) {
            oTiming.updateNodesIDs();
        }
        oEffectNode.getRoot().printTree();
    };
    CCTn.prototype.changeTrigger = function(v) {
        var oEffect = this.parent;
        if(!oEffect) {
            return;
        }
        var oTiming = oEffect.getTiming();
        if(!oTiming) {
            return;
        }
        var oTimingParent = oTiming.parent;//might be slide, layout, master
        if(!oTimingParent) {
            return;
        }
        var oCSld = oTimingParent.cSld;
        if(!oCSld) {
            return;
        }

        var aHierarchy = oEffect.getHierarchy();
        
        var oCurPar2Lvl, oCurPar3Lvl, oCurTopSeq;
        oCurPar3Lvl = aHierarchy[3];
        oCurPar2Lvl = aHierarchy[2];
        oCurTopSeq = aHierarchy[1];
        
        if(!oCurTopSeq || !oCurPar2Lvl || !oCurPar3Lvl) {
            return;
        }
        var nIdx3 = oCurPar3Lvl.getChildNodeIdx(oEffect);
        if(!v) {
            //move to the main sequence
            if(oCurTopSeq.isMainSequence()) {
                //do nothing
                return;
            }
            else {
                oCurPar3Lvl.splice(nIdx3);
                if(oCurPar3Lvl.getChildrenCount() === 0) {
                    oCurPar3Lvl.parent.onRemoveChild(oCurPar3Lvl);
                }
                oTiming.addToMainSequence(oEffect);
            }
        }
        else {
            //move to interactive seq with spId = v;
            var sObjectId;
            var oDrawing = oCSld.getObjectByName(v);
            if(!oDrawing) {
                return;
            }
            sObjectId = oDrawing.Get_Id();
            if(oCurTopSeq.isInteractiveSeq(sObjectId)) {
                //do nothing
                return;
            }
            else {
                oCurPar3Lvl.splice(nIdx3);
                if(oCurPar3Lvl.getChildrenCount() === 0) {
                    oCurPar3Lvl.parent.onRemoveChild(oCurPar3Lvl);
                }
                oTiming.addToInteractiveSequence(oEffect, sObjectId);
            }
        }
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



    var COND_EVNT_BEGIN = 0;
    var COND_EVNT_END = 1;
    var COND_EVNT_ON_BEGIN = 2;
    var COND_EVNT_ON_CLICK = 3;
    var COND_EVNT_ON_DBLCLICK = 4;
    var COND_EVNT_ON_END = 5;
    var COND_EVNT_ON_MOUSEOUT = 6;
    var COND_EVNT_ON_MOUSEOVER = 7;
    var COND_EVNT_ON_NEXT = 8;
    var COND_EVNT_ON_PREV = 9;
    var COND_EVNT_ON_STOPAUDIO = 10;

    var EVENT_DESCR_MAP = {};
    EVENT_DESCR_MAP[COND_EVNT_BEGIN] = "BEGIN";
    EVENT_DESCR_MAP[COND_EVNT_END] = "END";
    EVENT_DESCR_MAP[COND_EVNT_ON_BEGIN] = "ON_BEGIN";
    EVENT_DESCR_MAP[COND_EVNT_ON_CLICK] = "ON_CLICK";
    EVENT_DESCR_MAP[COND_EVNT_ON_DBLCLICK] = "ON_DBLCLICK";
    EVENT_DESCR_MAP[COND_EVNT_ON_END] = "ON_END";
    EVENT_DESCR_MAP[COND_EVNT_ON_MOUSEOUT] = "ON_MOUSEOUT";
    EVENT_DESCR_MAP[COND_EVNT_ON_MOUSEOVER] = "ON_MOUSEOVER";
    EVENT_DESCR_MAP[COND_EVNT_ON_NEXT] = "ON_NEXT";
    EVENT_DESCR_MAP[COND_EVNT_ON_PREV] = "ON_PREV";
    EVENT_DESCR_MAP[COND_EVNT_ON_STOPAUDIO] = "ON_STOPAUDIO";

    var RTN_ALL = 0;
    var RTN_FIRST = 1;
    var RTN_LAST = 2;

    function CCond() {
        CBaseAnimObject.call(this);
        this.rtn = null;
        this.tgtEl = null;
        this.tn = null;
        this.delay = null;
        this.evt = null;
    }
    InitClass(CCond, CBaseAnimObject, AscDFH.historyitem_type_Cond);
    CCond.prototype.setRtn = function(pr) {
        oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CondRtn, this.rtn, pr));
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
            oCopy.setRtn(this.rtn);
        }
        if(this.tgtEl !== null) {
            oCopy.setTgtEl(this.tgtEl.createDuplicate(oIdMap));
        }
        if(this.tn !== null) {
            oCopy.setTn(this.tn);
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
    CCond.prototype.getDelayTime = function() {
        if(this.delay === null) {
            return new CAnimationTime(0);
        }
        return this.parseTime(this.delay);
    };
    CCond.prototype.createDelaySimpleTrigger = function (oPlayer) {
        var oDelay = this.getDelayTime();
        if(oDelay.isIndefinite()) {
            return null;
        }
        var oStart = oPlayer.getElapsedTime();
        var oEnd;
        oEnd = oStart.plus(oDelay);
        return function () {
            var oElapsedTime = oPlayer.getElapsedTime();
            return oElapsedTime.greaterOrEquals(oEnd);
        };
    };
    CCond.prototype.createExternalEventTrigger = function (oPlayer, oTrigger, nType) {
        var oTimeNode = this.getNearestParentOrEqualTimeNode();
        if(!oTimeNode) {
            return null;
        }
        var sSpId = null;
        if(this.tgtEl) {
            sSpId = this.tgtEl.getSpId();
        }
        return oTimeNode.createExternalEventTrigger(oPlayer, oTrigger, nType, sSpId);
    };
    CCond.prototype.createEventTrigger = function (oPlayer, fEvent) {
        var oThis = this;
        return (function() {
            var oEnd = null;
            var oDelay = oThis.getDelayTime();
            return function () {
                if(oEnd) {
                    var oElapsedTime = oPlayer.getElapsedTime();
                    return oElapsedTime.greaterOrEquals(oEnd);
                }
                if(fEvent()) {
                    if(oDelay.isIndefinite() || oDelay.getVal() === 0) {
                        return true;
                    }
                    else {
                        var oStart = oPlayer.getElapsedTime();
                        oEnd = oStart.plus(oDelay);
                    }
                }
            }
        })();
    };
    CCond.prototype.createTimeNodeTrigger = function(oPlayer, fTimeNodeCheck) {
        var oTimeNode = null;
        if(this.tn !== null) {
            oTimeNode = this.findTimeNodeById(this.tn);
        }
        if(!oTimeNode) {
            var oCurTimeNode = this.getNearestParentOrEqualTimeNode();
            if(oCurTimeNode) {
                oTimeNode = oCurTimeNode.getParentTimeNode();
            }
            if(!oTimeNode) {
                return null;
            }
        }
        return this.createEventTrigger(oPlayer, function() {
            return fTimeNodeCheck(oTimeNode);
        });
    };
    CCond.prototype.createOnBeginTrigger = function(oPlayer) {
        return this.createTimeNodeTrigger(oPlayer, function (oTimeNode) {
            return oTimeNode.isActive();
        });
    };
    CCond.prototype.createOnEndTrigger = function(oPlayer) {
        return this.createTimeNodeTrigger(oPlayer, function (oTimeNode) {
            return oTimeNode.isFinished();
        });
    };
    CCond.prototype.fillTrigger = function(oPlayer, oTrigger) {
        switch (this.evt) {
            case COND_EVNT_BEGIN: {
                oTrigger.addTrigger(this.createDelaySimpleTrigger(oPlayer));
                break;
            }
            case COND_EVNT_END: {
                oTrigger.addTrigger(this.createDelaySimpleTrigger(oPlayer));
                break;
            }
            case COND_EVNT_ON_BEGIN: {
                oTrigger.addTrigger(this.createOnBeginTrigger(oPlayer));
                break;
            }
            case COND_EVNT_ON_CLICK: {
                oTrigger.addTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_CLICK));
                break;
            }
            case COND_EVNT_ON_DBLCLICK: {
                oTrigger.addTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_DBLCLICK));
                break;
            }
            case COND_EVNT_ON_END: {
                oTrigger.addTrigger(this.createOnEndTrigger(oPlayer));
                break;
            }
            case COND_EVNT_ON_MOUSEOUT: {
                oTrigger.addTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_MOUSEOUT));
                break;
            }
            case COND_EVNT_ON_MOUSEOVER: {
                oTrigger.addTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_MOUSEOVER));
                break;
            }
            case COND_EVNT_ON_NEXT: {
                oTrigger.addTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_NEXT));
                break;
            }
            case COND_EVNT_ON_PREV: {
                oTrigger.addTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_PREV));
                break;
            }
            case COND_EVNT_ON_STOPAUDIO: {
                break;
            }
            default: {
                oTrigger.addTrigger(this.createDelaySimpleTrigger(oPlayer));
                break;
            }
        }
    };
    CCond.prototype.getTargetObjectId = function() {
        if(this.tgtEl) {
            return this.tgtEl.getSpId();
        }
        return null;
    };

    changesFactory[AscDFH.historyitem_RtnVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_RtnVal] = function(oClass, value) {oClass.val = value;};
    function CRtn() {
        CBaseAnimObject.call(this);
        this.val = null;
    }
    InitClass(CRtn, CBaseAnimObject, AscDFH.historyitem_type_Rtn);
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
        CBaseAnimObject.call(this);
        this.inkTgt = null;//CObjectTarget
        this.sldTgt = null;
        this.sndTgt = null;
        this.spTgt = null;
    }
    InitClass(CTgtEl, CBaseAnimObject, AscDFH.historyitem_type_TgtEl);
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
    CTgtEl.prototype.getSpId = function() {
        if(this.spTgt) {
            return this.spTgt.spid;
        }
        return null;
    };


    changesFactory[AscDFH.historyitem_SndTgtEmbed] = CChangeLong;
    changesFactory[AscDFH.historyitem_SndTgtName] = CChangeString;
    changesFactory[AscDFH.historyitem_SndTgtBuiltIn] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_SndTgtEmbed] = function(oClass, value) {oClass.embed = value;};
    drawingsChangesMap[AscDFH.historyitem_SndTgtName] =  function(oClass, value) {oClass.name = value;};
    drawingsChangesMap[AscDFH.historyitem_SndTgtBuiltIn] =  function(oClass, value) {oClass.builtIn = value;};
    function CSndTgt() {//snd
        CBaseAnimObject.call(this);
        this.embed = null;
        this.name = null;
        this.builtIn = null;
    }
    InitClass(CSndTgt, CBaseAnimObject, AscDFH.historyitem_type_SndTgt);
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
        CBaseAnimObject.call(this);
        this.tmAbs = null;
        this.tmPct = null;
        this.backwards = null;
        this.type = null;
    }
    InitClass(CIterateData, CBaseAnimObject, AscDFH.historyitem_type_IterateData);
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
            oCopy.setTmAbs(this.tmAbs);
        }
        if(this.tmPct !== null) {
            oCopy.setTmPct(this.tmPct);
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

    changesFactory[AscDFH.historyitem_TavVal] = CChangeObject;
    changesFactory[AscDFH.historyitem_TavFmla] = CChangeString;
    changesFactory[AscDFH.historyitem_TavTm] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_TavVal] = function(oClass, value) {oClass.val = value;};
    drawingsChangesMap[AscDFH.historyitem_TavFmla] = function(oClass, value) {oClass.fmla = value;};
    drawingsChangesMap[AscDFH.historyitem_TavTm] = function(oClass, value) {oClass.tm = value;};
    function CTav() {
        CBaseAnimObject.call(this);
        this.val = null;
        this.fmla = null;
        this.tm = null;
    }
    InitClass(CTav, CBaseAnimObject, AscDFH.historyitem_type_Tav);
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
    CTav.prototype.getTime = function() {
        if(this.tm === null) {
            return 1.0;
        }
        if(this.tm.indexOf("%") === this.tm.length - 1) {
            return this.parsePercentage(this.tm);
        }
        var nTm = parseInt(this.tm);
        if(!isNaN(nTm)) {
            return nTm / 100000;
        }
        return 0;
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
        CBaseAnimObject.call(this);
        this.boolVal = null;
        this.clrVal = null;
        this.fltVal = null;
        this.intVal = null;
        this.strVal = null;
    }
    InitClass(CAnimVariant, CBaseAnimObject, AscDFH.historyitem_type_AnimVariant);
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
    CAnimVariant.prototype.getVal = function() {
        if(this.boolVal !== null) {
            return this.boolVal;
        }
        else if(this.clrVal !== null) {
            if(this.parent && this.parent.getTargetObject) {
                var oTargetObject = this.parent.getTargetObject();
                if(oTargetObject) {
                    var parents = oTargetObject.getParentObjects();
                    var RGBA = {R:0, G:0, B:0, A:255};
                    this.clrVal.Calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
                }
            }
            return this.clrVal;
        }
        else if(this.fltVal !== null) {
            return this.fltVal;
        }
        else if(this.intVal !== null) {
            return this.intVal;
        }
        else if(this.strVal !== null) {
            return this.strVal;
        }
        return null;
    };
    CAnimVariant.prototype.isBool = function() {
        return this.boolVal !== null;
    };
    CAnimVariant.prototype.isClr = function() {
        return this.clrVal !== null;
    };
    CAnimVariant.prototype.isFlt = function() {
        return this.fltVal !== null;
    };
    CAnimVariant.prototype.isInt = function() {
        return this.intVal !== null;
    };
    CAnimVariant.prototype.isStr = function() {
        return this.strVal !== null;
    };
    CAnimVariant.prototype.isSameType = function(oVariant) {
        if(!oVariant) {
            return false;
        }
        if(this.isBool() && oVariant.isBool()) {
            return true;
        }
        if(this.isClr() && oVariant.isClr()) {
            return true;
        }
        if(this.isFlt() && oVariant.isFlt()) {
            return true;
        }
        if(this.isInt() && oVariant.isInt()) {
            return true;
        }
        if(this.isStr() && oVariant.isStr()) {
            return true;
        }
        return false;
    };

    changesFactory[AscDFH.historyitem_AnimClrByRGB] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_AnimClrByHSL] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_AnimClrCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_AnimClrFrom] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_AnimClrTo] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_AnimClrClrSpc] = CChangeLong;
    changesFactory[AscDFH.historyitem_AnimClrDir] = CChangeLong;

    drawingConstructorsMap[AscDFH.historyitem_AnimClrByRGB] = CColorPercentage;
    drawingConstructorsMap[AscDFH.historyitem_AnimClrByHSL] = CColorPercentage;
    drawingConstructorsMap[AscDFH.historyitem_AnimClrFrom] = AscFormat.CUniColor;
    drawingConstructorsMap[AscDFH.historyitem_AnimClrTo] = AscFormat.CUniColor;

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

    var DIR_CCW = 0;
    var DIR_CW = 1;

    function CAnimClr() {
        CTimeNodeBase.call(this);
        this.byRGB = null;
        this.byHSL = null;

        this.cBhvr = null;
        this.from = null;
        this.to = null;
        this.clrSpc = null;
        this.dir = null;
    }
    InitClass(CAnimClr, CTimeNodeBase, AscDFH.historyitem_type_AnimClr);
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
        oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_AnimClrFrom, this.from, pr));
        this.from = pr;
        this.setParentToChild(pr);
    };
    CAnimClr.prototype.setTo = function(pr) {
        oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_AnimClrTo, this.to, pr));
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
    CAnimClr.prototype.isAllowedAttribute = function(sAttrName) {
        return sAttrName === "style.color" || sAttrName === "fillcolor" || sAttrName === "stroke.color";
    };
    CAnimClr.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
            return;
        }
        if(this.checkRemoveAtEnd()) {
            return;
        }
        var aAttributes = this.getAttributes();
        if(aAttributes.length < 1) {
            return;
        }
        var sFirstAttrName = aAttributes[0].text;
        if(!this.isAllowedAttribute(sFirstAttrName)) {
            return;
        }
        var oStartUniColor, oStartRGBColor;
        if(this.from) {
            oStartUniColor = this.from;
        }
        else {
            var oBrush;
            if(sFirstAttrName === "stroke.color") {
                var oPen = this.getTargetObjectPen();
                oBrush = oPen && oPen.Fill || AscFormat.CreateUnfilFromRGB(0, 0, 0);
            }
            else {
                oBrush = this.getTargetObjectBrush();
            }

            if(oBrush) {
                oStartRGBColor = oBrush.getRGBAColor();
                oStartUniColor = AscFormat.CreateUniColorRGB(oStartRGBColor.R, oStartRGBColor.G, oStartRGBColor.B);
            }
            else {
                oStartUniColor = AscFormat.CreateUniColorRGB(255, 255, 255);
            }
        }
        var oEndUniColor = this.to || this.by;
        var fRelTime;
        if(this.to || this.by) {
            oEndUniColor = this.to || this.by;
        }
        else if(this.byRGB || this.byHSL) {
            var parents = oTargetObject.getParentObjects();
            var RGBA = {R:0, G:0, B:0, A:255};
            oStartUniColor.Calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
            oStartRGBColor = oStartUniColor.RGBA;
            var oEndRGBColor = {R: 255, G: 255, B:255, A: 255};
            if(this.byRGB) {
                oEndRGBColor.R = oStartRGBColor.R * (1 + this.byRGB.c1/100000);
                oEndRGBColor.R = Math.min(255, Math.max(0, oStartRGBColor.R));
                oEndRGBColor.G = oStartRGBColor.G * (1 + this.byRGB.c2/100000);
                oEndRGBColor.G = Math.min(255, Math.max(0, oStartRGBColor.G));
                oEndRGBColor.B = oStartRGBColor.B * (1 + this.byRGB.c3/100000);
                oEndRGBColor.B = Math.min(255, Math.max(0, oStartRGBColor.B));
            }
            else if(this.byHSL) {
                fRelTime = this.getRelativeTime(nElapsedTime);
                var oStartHSL = this.toFormatHSLColor(oStartRGBColor);
                var oResultHSL = {};

                var dAlignAngle = 360*60000;
                var dStartAng = this.alignNumber(oStartHSL.H, dAlignAngle);
                var dEndAng = this.alignNumber(oStartHSL.H + this.byHSL.c1, dAlignAngle);

                dEndAng = this.alignNumber(dEndAng - dStartAng, dAlignAngle);
                if(this.dir === null || this.dir === DIR_CW) {
                    oResultHSL.H = this.alignNumber(dEndAng*fRelTime + dStartAng);
                }
                else {
                    oResultHSL.H = this.alignNumber(dAlignAngle - fRelTime * (dAlignAngle - dEndAng) + dStartAng, 360*60000);
                }
                oResultHSL.S = Math.min(100000, Math.max(-100000, oStartHSL.S + fRelTime * this.byHSL.c2));
                oResultHSL.L = Math.min(100000, Math.max(-100000, oStartHSL.L + fRelTime * this.byHSL.c3));
                var oResultRGB = this.toRGBAColor(oResultHSL);
                var oResultUnicolor = AscFormat.CreateUniColorRGB(oResultRGB.R, oResultRGB.G, oResultRGB.B);
                oResultUnicolor.Calculate(parents.theme, parents.slide, parents.layout, parents.master, RGBA);
                oAttributes[sFirstAttrName] = oResultUnicolor;
                return;
            }
            oEndUniColor = AscFormat.CreateUniColorRGB(oEndRGBColor.R, oEndRGBColor.G, oEndRGBColor.B);
        }

        fRelTime = this.getRelativeTime(nElapsedTime);
        oAttributes[sFirstAttrName] = this.getAnimatedClr(fRelTime, oStartUniColor, oEndUniColor);
    };

    CAnimClr.prototype.toFormatHSLColor = function(oRGBA) {
        var oHSL = {};
        var oColorModifiers = new AscFormat.CColorModifiers();
        oColorModifiers.RGB2HSL(oRGBA.R, oRGBA.G, oRGBA.B, oHSL);
        oHSL.H /= 255;
        oHSL.H *= 360*60000;

        oHSL.S /= 255;
        oHSL.S *= 200000;
        oHSL.S -= 100000;

        oHSL.L /= 255;
        oHSL.L *= 200000;
        oHSL.L -= 100000;
        return oHSL;
    };
    CAnimClr.prototype.toRGBAColor = function(oFormatHSL) {
        var oHSL = {};
        oHSL.H = this.alignNumber(255 * oFormatHSL.H /(360*60000), 255);
        oHSL.S = Math.min(255, Math.max(0, 255 * (oFormatHSL.S + 100000) / 200000));
        oHSL.L = Math.min(255, Math.max(0, 255 * (oFormatHSL.L + 100000) / 200000));
        var oRGBColor = {R: 255, G: 255, B:255, A: 255};
        var oColorModifiers = new AscFormat.CColorModifiers();
        oColorModifiers.HSL2RGB(oHSL, oRGBColor);
        return oRGBColor;
    };
    CAnimClr.prototype.alignNumber = function(dVal, dMax) {
        var dValChecked = dVal;
        while(dValChecked < 0) {
            dValChecked += dMax;
        }
        while(dValChecked >= dMax) {
            dValChecked -= dMax;
        }
        return dValChecked;
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

    var TRANSITION_TYPE_IN = 0;
    var TRANSITION_TYPE_OUT = 1;
    var TRANSITION_TYPE_NONE = 2;
    function CAnimEffect() {
        CTimeNodeBase.call(this);
        this.cBhvr = null;
        this.progress = null;
        this.filter = null;
        this.prLst = null;
        this.transition = null;
    }
    InitClass(CAnimEffect, CTimeNodeBase, AscDFH.historyitem_type_AnimEffect);
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
    CAnimEffect.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        if(!this.filter) {
            return;
        }
        if(this.checkRemoveAtEnd()) {
            return;
        }
        var fRelTime = this.getRelativeTime(nElapsedTime);
        // if(this.transition === TRANSITION_TYPE_IN) {
        //     fRelTime = 1 - fRelTime;
        // }
        if(this.progress && this.progress.isFlt()) {
            fRelTime = this.progress.getVal()
        }
        var aFiltersStrings = this.filter.split(";");
        var aFilters = [];
        for(var nFilter = 0; nFilter < aFiltersStrings.length; ++nFilter) {
            var nFilterType = FILTER_MAP[aFiltersStrings[nFilter]];
            if(AscFormat.isRealNumber(nFilterType)) {
                aFilters.push(nFilterType);
            }
        }
        return oAttributes["effect"] = new CEffectData(aFilters, fRelTime, this.prLst, this.transition);
    };

    CAnimEffect.prototype.create = function(nTransition, sFilter, sObjectId, sDur) {
        this.setTransition(nTransition);
        this.setFilter(sFilter);
        var oBhvr = new CCBhvr();
        var oCTn = this.createCCTn(sDur, null, null, null, null, null);
        oBhvr.setCTn(oCTn);
        oBhvr.setTgtEl(this.createSpTgt(sObjectId));
        this.setCBhvr(oBhvr);
    };

    function CEffectData(aFilters, fRelTime, sPrLst, nTransition) {
        this.filters = aFilters;
        this.data = {
            time: fRelTime,
            prLst: sPrLst,
            transition: nTransition
        }
    }
    CEffectData.prototype.isEqual = function(oOther) {
        if(!AscFormat.fApproxEqual(this.data.time, oOther.data.time)){
            return false;
        }
        if(this.filters.length !== oOther.filters.length) {
            return false;
        }
        for(var nFilter = 0; nFilter < this.filters.length; ++nFilter) {
            if(this.filters[nFilter] !== oOther.filters[nFilter]) {
                return false;
            }
        }
        if(this.data.prLst !== oOther.data.prLst) {
            return false;
        }
        if(this.data.transition !== oOther.data.transition) {
            return false;
        }
        return true;
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
    drawingsChangesMap[AscDFH.historyitem_AnimMotionCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionFrom] = function(oClass, value) {oClass.from = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionRCtr] = function(oClass, value) {oClass.rCtr = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionTo] = function(oClass, value) {oClass.to = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionOrigin] = function(oClass, value) {oClass.origin = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPath] = function(oClass, value) {oClass.path = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPathEditMode] = function(oClass, value) {oClass.pathEditMode = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPtsTypes] = function(oClass, value) {oClass.ptsTypes = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionRAng] = function(oClass, value) {oClass.rAng = value;};


    var ORIGIN_PARENT = 0;
    var ORIGIN_LAYOUT = 1;

    function CAnimMotion() {
        CTimeNodeBase.call(this);
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
    InitClass(CAnimMotion, CTimeNodeBase, AscDFH.historyitem_type_AnimMotion);
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
    CAnimMotion.prototype.privateCalculateParams = function() {
        if(this.path) {
            this.parsedPath = new CSVGPath(this.path);
        }
        else {
            this.parsedPath = null;
        }
    };
    CAnimMotion.prototype.getOrigin = function() {
        if(this.origin !== null) {
            return this.origin;
        }
        return ORIGIN_PARENT;
    };
    CAnimMotion.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
            return;
        }
        if(this.checkRemoveAtEnd()) {
            return;
        }
        var nOrigin = this.getOrigin();
        var fRelTime = this.getRelativeTime(nElapsedTime);
        var dPosX = null, dPosY = null;
        var dRelX = null, dRelY = null;
        var oBounds = oTargetObject.getBoundsByDrawing();
        var fSlideW = this.getSlideWidth();
        var fSlideH = this.getSlideHeight();
        var fObjRelX = oBounds.x / fSlideW;
        var fObjRelY = oBounds.y / fSlideH;
        if (this.parsedPath) {
            var oPos = this.parsedPath.getPosition(fRelTime);
            if(oPos) {
                dRelX = oPos.x;
                dRelY = oPos.y;
            }
        }
        else if(this.to && this.from) {
            dRelX = (this.from.x + (this.to.x - this.from.x) * fRelTime) / 100;
            dRelY = (this.from.y + (this.to.y - this.from.y) * fRelTime) / 100;
        }
        else if(this.by && this.from) {
            dRelX = (this.from.x + this.by.x * fRelTime) / 100;
            dRelY = (this.from.y + this.by.y * fRelTime) / 100;
        }
        else if(this.by) {
            dRelX = fObjRelX + (this.by.x / 100)*fRelTime;
            dRelY = fObjRelY + (this.by.y / 100)*fRelTime;
        }
        else if(this.to) {
            dRelX = fObjRelX + ((this.to.x / 100) - fObjRelX)*fRelTime;
            dRelY = fObjRelY + ((this.to.y / 100) - fObjRelY)*fRelTime;
        }
        if(dRelX !== null && dRelY !== null) {
            if(nOrigin === ORIGIN_LAYOUT) {
                dRelX += ((oBounds.x + oBounds.w/2)/fSlideW);
                dRelY += ((oBounds.y + oBounds.h/2)/fSlideH);
            }
            var aAttr = this.getAttributes();
            if(aAttr[0] && this.isAllowedAttribute(aAttr[0].text)) {
                oAttributes[aAttr[0].text] = dRelX;
            }
            else {
                oAttributes["ppt_x"] = dRelX;
            }
            if(aAttr[1] && this.isAllowedAttribute(aAttr[1].text)) {
                oAttributes[aAttr[1].text] = dRelY;
            }
            else {
                oAttributes["ppt_y"] = dRelY;
            }
        }
        else {
            //console.log("Something went wrong");
        }
    };
    CAnimMotion.prototype.isAllowedAttribute = function(sAttrName) {
        return sAttrName === "ppt_x" || "ppt_y" || "ppt_w" ||
            "ppt_h" || "ppt_r" || "style.fontSize" ||
            "xskew" || "yskew" || "xshear" ||
            "yshear" || "scaleX" || "scaleY";
    };

    function CSVGPath(sPath) {
        this.pathString = sPath;
        this.commands = [];
        this.lengths = [];
        this.parsePath(this.pathString);
    }
    CSVGPath.prototype.numberRegExp = /-?[0-9]*\.?[0-9]+(?:e[-+]?\d+)?/ig;
    CSVGPath.prototype.setEmpty = function() {
        this.commands.length = 0;
        this.lengths.length = 0;
    };
    CSVGPath.prototype.parsePath = function(sPath) {
        var aLastCommand = null;
        var aElementsSplit = sPath.split(" ");
        var nCurElement = 0;
        var sCurElement;
        var aCurCommand;
        var dPX1, dPY1, dPX2, dPY2, dPX3, dPY3;
        var aLastPoint = null;
        var fLastLength = 0.0;
        var fL;
        var aElements = [];
        for(nCurElement = 0; nCurElement < aElementsSplit.length; ++nCurElement) {
            if(aElementsSplit[nCurElement].length > 0) {
                aElements.push(aElementsSplit[nCurElement]);
            }
        }
        nCurElement = 0;
        while(nCurElement < aElements.length) {
            sCurElement = aElements[nCurElement];
            if(sCurElement.length === 0) {
                nCurElement++;
                continue;
            }
            aCurCommand = [];
            fL = 0;
            if("M" === sCurElement || "m" === sCurElement
            || "L" === sCurElement || "l" === sCurElement) {
                if("L" === sCurElement || "l" === sCurElement) {
                    if(!aLastPoint) {
                        this.setEmpty();
                        return;
                    }
                }
                aCurCommand.push(sCurElement.toUpperCase());
                dPX1 = this.parseValues(aElements[++nCurElement])[0];
                dPY1 = this.parseValues(aElements[++nCurElement])[0];
                if(sCurElement.toLowerCase() === sCurElement) {
                    if(!aLastPoint) {
                        this.setEmpty();
                        return;
                    }
                    dPX1 += aLastPoint[0];
                    dPY1 += aLastPoint[1];
                }
                aCurCommand.push(dPX1);
                aCurCommand.push(dPY1);
                if("L" === sCurElement || "l" === sCurElement) {
                    fL = this.calculateLineLength(aLastPoint, aCurCommand.slice(1));
                }
            }
            else if("C" === sCurElement || "c" === sCurElement) {
                if(!aLastPoint) {
                    this.setEmpty();
                    return;
                }
                aCurCommand.push(sCurElement.toUpperCase());
                dPX1 = this.parseValues(aElements[++nCurElement])[0];
                dPY1 = this.parseValues(aElements[++nCurElement])[0];
                dPX2 = this.parseValues(aElements[++nCurElement])[0];
                dPY2 = this.parseValues(aElements[++nCurElement])[0];
                dPX3 = this.parseValues(aElements[++nCurElement])[0];
                dPY3 = this.parseValues(aElements[++nCurElement])[0];
                if(sCurElement.toLowerCase() === sCurElement) {
                    if(!aLastPoint) {
                        this.setEmpty();
                        return;
                    }
                    dPX1 += aLastPoint[0];
                    dPY1 += aLastPoint[1];
                    dPX2 += aLastPoint[0];
                    dPY2 += aLastPoint[1];
                    dPX3 += aLastPoint[0];
                    dPY3 += aLastPoint[1];
                }
                aCurCommand.push(dPX1);
                aCurCommand.push(dPY1);
                aCurCommand.push(dPX2);
                aCurCommand.push(dPY2);
                aCurCommand.push(dPX3);
                aCurCommand.push(dPY3);
                fL = this.calculateBezierLength(aLastPoint, [dPX1, dPY1], [dPX2, dPY2], [dPX3, dPY3]);
            }
            else if("Z" === sCurElement || "z" === sCurElement) {
                if(!aLastPoint) {
                    this.setEmpty();
                    return;
                }
                aCurCommand.push("Z");
                var aMoveToCommand = this.findMoveToCommand(this.commands.length - 1);
                if(!aMoveToCommand) {
                    this.setEmpty();
                    return;
                }
                fL = this.calculateLineLength(aLastPoint, aMoveToCommand.slice(1));
            }
            else if("E" === sCurElement || "e" === sCurElement) {
                aCurCommand.push("E");
                this.commands.push(aCurCommand);
                this.lengths.push(fLastLength + fL);
                return;
            }
            else {
                this.setEmpty();
                return;
            }
            if(aCurCommand.length > 0) {
                this.commands.push(aCurCommand);
                this.lengths.push(fLastLength + fL);
                aLastCommand = aCurCommand;
                fLastLength += fL;
                if(aLastCommand.length > 2) {
                    aLastPoint = aLastCommand.slice(aLastCommand.length - 2);
                }
                else {
                    aLastPoint = null;
                }
            }
            nCurElement++;
        }
    };
    CSVGPath.prototype.parseValues = function(args) {
        var numbers = args.match(this.numberRegExp);
        return numbers ? numbers.map(Number) : []
    };
    CSVGPath.prototype.findMoveToCommand = function(nStartIdx) {
        for(var nIdx = nStartIdx; nIdx > -1; nIdx--) {
            var aCommand = this.commands[nIdx];
            if(aCommand[0] === "M") {
                return aCommand;
            }
        }
        return null;
    };
    CSVGPath.prototype.calculateLineLength = function(aP0, aP1) {
        var dx = aP0[0] - aP1[0];
        var dy = aP0[1] - aP1[1];
        return Math.sqrt(dx * dx + dy * dy);
    };
    CSVGPath.prototype.calculateBezierLength = function(aP0, aP1, aP2, aP3) {
        var chord = this.calculateLineLength(aP3, aP0);
        var p0_p1 = this.calculateLineLength(aP0, aP1);
        var p2_p1 = this.calculateLineLength(aP2, aP1);
        var p3_p2 = this.calculateLineLength(aP3, aP2);
        var cont_net = (p0_p1) + (p2_p1) + (p3_p2);
        return (cont_net + chord) / 2;
    };
    CSVGPath.prototype.calculateBezierLength = function(aP0, aP1, aP2, aP3) {
        var chord = this.calculateLineLength(aP3, aP0);
        var p0_p1 = this.calculateLineLength(aP0, aP1);
        var p2_p1 = this.calculateLineLength(aP2, aP1);
        var p3_p2 = this.calculateLineLength(aP3, aP2);
        var cont_net = (p0_p1) + (p2_p1) + (p3_p2);
        return (cont_net + chord) / 2;
    };
    CSVGPath.prototype.getPosition = function(fTime) {
        if(this.lengths.length === 0) {
            return null;
        }
        var fLength = this.lengths[this.lengths.length - 1];
        var fCurLen = fLength * fTime;
        for(var nP = 0; nP < this.lengths.length - 1; ++nP) {
            if(this.lengths[nP] >= fCurLen) {
                break;
            }
        }
        var oCommand = this.commands[nP];
        var fX = 0.0, fY = 0.0;
        var fCurveLength= this.lengths[nP] - (this.lengths[nP - 1] || 0);
        var fLenInCurve = fCurLen - (this.lengths[nP - 1] || 0);
        var t = fLenInCurve / fCurveLength;
        var fPrevX = 0;
        var fPrevY = 0;
        var oPrevCommand = this.commands[nP - 1];
        if(oPrevCommand) {
            fPrevX = oPrevCommand[oPrevCommand.length - 2];
            fPrevY = oPrevCommand[oPrevCommand.length - 1];
        }
        if(oCommand[0] === "M") {
            fX = oCommand[1];
            fY = oCommand[2];
        }
        else if(oCommand[0] === "L") {
            fX = (1 - t)*fPrevX + t*oCommand[1];
            fY = (1 - t)*fPrevY + t*oCommand[2];
        }
        else if(oCommand[0] === "C") {
            var x0 = fPrevX;
            var y0 = fPrevY;
            var x1 = oCommand[1];
            var y1 = oCommand[2];
            var x2 = oCommand[3];
            var y2 = oCommand[4];
            var x3 = oCommand[5];
            var y3 = oCommand[6];
            var dt = (1 - t);
            var dt2 = dt * dt;
            var dt3 = dt2 * dt;
            var t2 = t*t;
            var t3 = t*t2;
            var p1 = 3*dt2*t;
            var p2 = 3*dt*t2;
            fX = dt3*x0 + p1*x1 + p2*x2 + t3*x3;
            fY = dt3*y0 + p1*y1 + p2*y2 + t3*y3;
        }
        else if(oCommand[0] === "Z") {
            var aMoveToCommand = this.findMoveToCommand(nP - 1);
            if(aMoveToCommand) {
                fX = (1 - t)*fPrevX + t*aMoveToCommand[1];
                fY = (1 - t)*fPrevY + t*aMoveToCommand[2];
            }
        }
        return {x: fX, y: fY};
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
        CTimeNodeBase.call(this);
        this.cBhvr = null;
        this.by = null;
        this.from = null;
        this.to = null;
    }
    InitClass(CAnimRot, CTimeNodeBase, AscDFH.historyitem_type_AnimRot);
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
    CAnimRot.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
            return;
        }
        if(this.checkRemoveAtEnd()) {
            return;
        }
        var fRelTime = this.getRelativeTime(nElapsedTime);
        var dR = null;
        if(this.to && this.from) {
            dR = this.from + (this.to - this.from) & fRelTime;
        }
        else if(this.by && this.from) {
            dR = this.from + fRelTime*this.by;
        }
        if(this.by !== null) {
            dR = this.by*fRelTime;
        }
        else if(this.to !== null) {
            dR = this.to*fRelTime;
        }
        if(dR !== null) {
            var aAttr = this.getAttributes();
            if(aAttr[0] && this.isAllowedAttribute(aAttr[0].text)) {
                this.setAttributeValue(oAttributes, aAttr[0].text, dR);
            }
            else {
                this.setAttributeValue(oAttributes, "r", dR);
            }
        }
    };
    CAnimRot.prototype.isAllowedAttribute = function() {
        return "ppt_r" || "r" || "style.rotation";
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
        CTimeNodeBase.call(this);
        this.cBhvr = null;
        this.by = null;
        this.from = null;
        this.to = null;
        this.zoomContents = null;
    }
    InitClass(CAnimScale, CTimeNodeBase, AscDFH.historyitem_type_AnimScale);
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
    CAnimScale.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
            return;
        }
        if(this.checkRemoveAtEnd()) {
            return;
        }
        var fRelTime = this.getRelativeTime(nElapsedTime);
        var dRelX = null, dRelY = null;
        if(this.to && this.from) {
            dRelX = (this.from.x + (this.to.x - this.from.x) * fRelTime) / 100000;
            dRelY = (this.from.y + (this.to.y - this.from.y) * fRelTime) / 100000;
        }
        else if(this.by && this.from) {
            dRelX = (this.from.x + this.by.x * fRelTime) / 100000;
            dRelY = (this.from.y + this.by.y * fRelTime) / 100000;
        }
        else if(this.by) {
            dRelX = 1*(1 - fRelTime) + (this.by.x / 100000)*fRelTime;
            dRelY = 1*(1 - fRelTime) + (this.by.y / 100000)*fRelTime;
        }
        else if(this.to) {
            dRelX = 1*(1 - fRelTime) + (this.to.x / 100000)*fRelTime;
            dRelY = 1*(1 - fRelTime) + (this.to.y / 100000)*fRelTime;
        }
        if(dRelX !== null && dRelY !== null) {
            oAttributes["ScaleX"] = dRelX;
            oAttributes["ScaleY"] = dRelY;
        }
    };

    changesFactory[AscDFH.historyitem_AudioCMediaNode] = CChangeObject;
    changesFactory[AscDFH.historyitem_AudioIsNarration] = CChangeBool;

    drawingsChangesMap[AscDFH.historyitem_AudioCMediaNode] = function(oClass, value) {oClass.cMediaNode = value;};
    drawingsChangesMap[AscDFH.historyitem_AudioIsNarration] = function(oClass, value) {oClass.isNarration = value;};

    function CAudio() {
        CTimeNodeBase.call(this);
        this.cMediaNode = null;
        this.isNarration = null;
    }
    InitClass(CAudio, CTimeNodeBase, AscDFH.historyitem_type_Audio);

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
        CTimeNodeBase.call(this);
        this.cTn = null;
        this.tgtEl = null;
        this.mute = null;
        this.numSld = null;
        this.showWhenStopped = null;
        this.vol = null;
    }
    InitClass(CCMediaNode, CTimeNodeBase, AscDFH.historyitem_type_CMediaNode);
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
        CTimeNodeBase.call(this);
        this.cBhvr = null;
        this.cmd = null;
        this.type = null;
    }
    InitClass(CCmd, CTimeNodeBase, AscDFH.historyitem_type_Cmd);
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

    CCmd.prototype.setState = function(nState) {
        CTimeNodeBase.prototype.setState.call(this, nState);
        if(nState === TIME_NODE_STATE_ACTIVE) {
            var sCmd = this.cmd;
            if(sCmd) {
                if(sCmd.indexOf("play") || sCmd === "resume" || sCmd === "togglePause") {
                    var oSp = this.getTargetObject();
                    if(oSp) {
                        var sMediaName = oSp.getMediaFileName();
                        if(sMediaName) {
                            var oApi = Asc.editor || editor;
                            if(oApi && oApi.showVideoControl) {
                                oApi.showVideoControl(sMediaName, oSp.extX, oSp.extY, oSp.transform);
                            }
                        }
                    }
                }
            }
        }
        //this.logState("SET STATE:");
    };

    changesFactory[AscDFH.historyitem_TimeNodeContainerCTn] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TimeNodeContainerCTn] = function(oClass, value) {oClass.cTn = value;};




    var ANIM_LABEL_WIDTH_PIX = 22;
    var ANIM_LABEL_HEIGHT_PIX = 17;
    var HOR_LABEL_SPACE = 9;
    var VERT_LABEL_SPACE = 4;

    function CTimeNodeContainer() {//par, excl
        CTimeNodeBase.call(this);
        this.cTn = null;

        this.triggerClickSequence = undefined;
        this.triggerObjectClick = undefined;
        this.settingsDelay = undefined;
        this.selected = false;
    }
    InitClass(CTimeNodeContainer, CTimeNodeBase, AscDFH.historyitem_type_TimeNodeContainer);
    CTimeNodeContainer.prototype.setCTn = function(pr) {
        oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TimeNodeContainerCTn, this.cTn, pr));
        this.cTn = pr;
        this.setParentToChild(pr);
    };
    CTimeNodeContainer.prototype.fillObject = function(oCopy, oIdMap) {
        if(this.cTn !== null) {
            oCopy.setCTn(this.cTn.createDuplicate(oIdMap));
        }
        if(this.selected) {
            oCopy.selected = true;
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
    CTimeNodeContainer.prototype.getChildrenTimeNodesInternal = function() {
        if(this.cTn) {
            if(this.cTn.childTnLst) {
                return this.cTn.childTnLst.list;
            }
        }
        return [];
    };
    CTimeNodeContainer.prototype.isAnimEffect = function() {
        return this.cTn.isAnimEffect();
    };
    CTimeNodeContainer.prototype.findSeqWithIdx = function() {
        var oTiming = this.getTiming();
        if(!oTiming) {
            return null;
        }
        var aSeqs = oTiming.getEffectsSequences();
        for(var nSeq = 0; nSeq < aSeqs.length; ++nSeq) {
            var aSeq = aSeqs[nSeq];
            for(var nEffect = aSeq.length - 1; nEffect > 0; --nEffect) {
                if(aSeq[nEffect] === this) {
                    return {seq: aSeq, idx: nEffect};
                }
            }
        }
        return null;
    };
    CTimeNodeContainer.prototype.getIndexInSequence = function() {
        var aHierarchy = this.getHierarchy();
        if(aHierarchy[1] && aHierarchy[2]) {
            return aHierarchy[1].getChildNodeIdx(aHierarchy[2]);
        }
        return -1;
    };
    CTimeNodeContainer.prototype.getLabelFillColor = function() {

    };
    CTimeNodeContainer.prototype.getLabelRect = function() {
        var oTiming = this.getTiming();
        if(!oTiming) {
            return null;
        }
        var sObjectId = this.getObjectId();
        var oObject = AscCommon.g_oTableId.Get_ById(sObjectId);
        if(!oObject) {
            return null;
        }
        var aObjectEffects = oTiming.getObjectEffects(sObjectId);
        var dX, dY;
        var dW = oObject.convertPixToMM(ANIM_LABEL_WIDTH_PIX);
        var dH = oObject.convertPixToMM(ANIM_LABEL_HEIGHT_PIX);
        var oObjectBounds = oObject.bounds;
        dX = oObjectBounds.x - oObject.convertPixToMM(HOR_LABEL_SPACE + ANIM_LABEL_WIDTH_PIX);
        dY = oObjectBounds.y;
        for(var nEffect = 0; nEffect < aObjectEffects.length; ++nEffect) {
            if(aObjectEffects[nEffect] === this) {
                break;
            }
            dY += (dH + oObject.convertPixToMM(VERT_LABEL_SPACE));
        }
        return new AscFormat.CGraphicBounds(dX, dY, dX + dW, dY + dH);
    };
    CTimeNodeContainer.prototype.drawEffectLabel = function(oGraphics) {
        this.internalDrawEffectLabel(oGraphics);
    };
    CTimeNodeContainer.prototype.hit = function(x, y) {
        var oRect = this.getLabelRect();
        if(!oRect) {
            return;
        }
        return oRect.hit(x, y);
    };
    CTimeNodeContainer.prototype.isCorrect = function() {
        if(!this.cTn) {
            return false;
        }
        var sObjectId = this.getObjectId();
        var oObj = AscCommon.g_oTableId.Get_ById(sObjectId);
        if(!oObj) {
            return false;
        }
        if(!oObj.checkCorrect() || !(oObj.Is_UseInDocument && oObj.Is_UseInDocument())) {
            return false;
        }
        return true;
    };
    var ICON_TRIGGER = "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTEiIGhlaWdodD0iMTQiIHZpZXdCb3g9IjAgMCAxMSAxNCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMTEgMEg1TDAgN0g0TDAgMTRMMTEgNUg2TDExIDBaIiBmaWxsPSIjNDQ0NDQ0Ii8+PC9zdmc+";
    
    CTimeNodeContainer.prototype.internalDrawEffectLabel = function(oGraphics) {
        var oRect = this.getLabelRect();
        if(!oRect) {
            return;
        }
        AscFormat.ExecuteNoHistory(function(){
            var dX, dY, dW, dH;
            dX = oRect.l;
            dY = oRect.t;
            dW = oRect.w;
            dH = oRect.h;
            
			if (oGraphics.IsSlideBoundsCheckerType) {
                oGraphics.rect(dX, dY, dW, dH);
                return;
			}
            var oContext = oGraphics.m_oContext;
            if(!oContext) {
                return;
            }

            oGraphics.SaveGrState();
            
            var oMatrix = new AscCommon.CMatrix();
            oMatrix.tx = dX;
            oMatrix.ty = dY;
            oGraphics.transform3(oMatrix);
            //draw rect
            
            oGraphics.SetIntegerGrid(true);
            var nFillColor = this.isSelected() ? 0xCBCBCB : 0xFFFFFF;
            var nLineColor = 0xC0C0C0; 
            oGraphics.b_color1((nFillColor >> 16) & 0xFF, (nFillColor >> 8) & 0xFF, nFillColor & 0xFF, 0xFF);
            oGraphics.p_color((nLineColor >> 16) & 0xFF, (nLineColor >> 8) & 0xFF, nLineColor & 0xFF, 255);
            oGraphics.p_width(0);
            oGraphics._s();

            
            var oFullTr = oGraphics.m_oFullTransform;
            
            var _x1 = (oFullTr.TransformPointX(0, 0)) >> 0;
            var _y1 = (oFullTr.TransformPointY(0, 0)) >> 0;
            var _x2 = (oFullTr.TransformPointX(dW, dH)) >> 0;
            var _y2 = (oFullTr.TransformPointY(dW, dH)) >> 0;
            
            oContext.lineWidth = 1;
            oContext.rect(_x1 + 0.5, _y1 + 0.5, _x2 - _x1, _y2 - _y1);
            //oGraphics.rect(0, 0, dW, dH);
            oGraphics.df();
            oGraphics.ds();
            
            // oGraphics.p_color((nLineColor >> 16) & 0xFF, (nLineColor >> 8) & 0xFF, nLineColor & 0xFF, 255);
            // oGraphics.p_width(0);
            // oGraphics._s();
            // oGraphics.drawVerLine(1, 0, 0, dH, 0);
            // oGraphics.drawVerLine(1, dW, 0, dH, 0);
            // oGraphics.drawHorLine(1, 0, 0, dW, 0);
            // oGraphics.drawHorLine(1, dH, 0, dW, 0);
            // oGraphics.ds();
            //draw internal part
            
            oGraphics.RestoreGrState();

            var sObjectId = this.getObjectId();
            var oObject = AscCommon.g_oTableId.Get_ById(sObjectId);
            var oT = oGraphics.m_oCoordTransform;
            if(this.isPartOfMainSequence()) {
                var nIdx = this.getIndexInSequence();
                if(AscFormat.isRealNumber(nIdx)) {
                        if(!oObject) {
                            return null;
                        }
                        var dTX = dX + dW / 2;
                        var dTY = dY + dH - oObject.convertPixToMM(4);


                        var nX = oT.TransformPointX(dTX, dTY);
                        var nY = oT.TransformPointY(dTX, dTY);
                        var sOldFill = oContext.fillStyle;
                        oContext.fillStyle = "#000000";
                        oContext.fillText((nIdx + 1) + "", nX, nY);    
                        oContext.fillStyle = sOldFill;
                }
            }
            else {
                
                var oApi = editor || Asc.editor;
                if(oApi && oApi.ImageLoader) {
                    var oImage = oApi.ImageLoader.map_image_index[ICON_TRIGGER];
                    if(oImage)  {
                        var oNImage = oImage.Image;
                        var nNativeW = oNImage.width / 2;
                        var nNativeH = oNImage.height / 2;
                        var nWidth = AscCommon.AscBrowser.convertToRetinaValue(nNativeW, true);
                        var nHeight = AscCommon.AscBrowser.convertToRetinaValue(nNativeH, true);
                        var dTX = dX + (dW - oObject.convertPixToMM(nNativeW)) / 2;
                        var dTY = dY + (dH - oObject.convertPixToMM(nNativeH)) / 2;
                        var nX = oT.TransformPointX(dTX, dTY);
                        var nY = oT.TransformPointY(dTX, dTY);
                        oContext.drawImage(oImage.Image, nX, nY, nWidth, nHeight);
                    }
                }
            }
            
            

        }, this, []);
    };
    CTimeNodeContainer.prototype.getObjectId = function() {
        if(this.isAnimEffect()) {
            return this.cTn.getObjectId();
        }
        return null;
    };
    CTimeNodeContainer.prototype.isObjectEffect = function(sObjectId) {
        if(this.isAnimEffect()) {
            return this.getObjectId() === sObjectId;
        }
        return false;
    };
    CTimeNodeContainer.prototype.merge = function(oTNContainer) {
        AscFormat.ExecuteNoHistory(function() {
            if(oTNContainer === this) {
                return;
            }
            this.cTn.merge(oTNContainer.cTn);
            if(!Array.isArray(this.merged)) {
                this.merged = [];
            }
            var oCopy = oTNContainer.createDuplicate();
            oCopy.originalNode = oTNContainer;
            oCopy.setParent(oTNContainer.parent);
            this.merged.push(oCopy);
        }, this, []);
    };
    CTimeNodeContainer.prototype.isMultipleEffect = function() {
        if(Array.isArray(this.merged) && this.merged.length > 1) {
            if(this.cTn.presetClass === undefined || this.cTn.presetID === undefined) {
                return true;
            }
        }
        return false;
    };
    CTimeNodeContainer.prototype.addToChildTnLst = function(nIdx, oNode) {
        this.cTn.addToChildTnLst(nIdx, oNode);
    };
    CTimeNodeContainer.prototype.pushToChildTnLst = function(oNode) {
        this.cTn.pushToChildTnLst(oNode);
    };
    CTimeNodeContainer.prototype.clearChildTnLst = function() {
        this.cTn.clearChildTnLst();
    };
    CTimeNodeContainer.prototype.resetDelayShift = function() {
        if(this.isAnimEffect()) {
            this.cTn.resetDelayShift();
        }
    };
    CTimeNodeContainer.prototype.setDelayShift = function() {
        if(this.isAnimEffect()) {
            this.cTn.setDelayShift();
        }
    };
    CTimeNodeContainer.prototype.splice = function() {
        return this.cTn.childTnLst.splice.apply(this.cTn.childTnLst, arguments);
    };
    CTimeNodeContainer.prototype.getLastChild = function() {
        return this.cTn.childTnLst.getLast();
    };
    CTimeNodeContainer.prototype.addEffects = function(nInsertIdx, aEffects) {
        for(var nIdx = 0; nIdx < aEffects.length; ++nIdx) {
            this.splice(nInsertIdx + nIdx, 0, aEffects[nIdx]);
        }
    };
    CTimeNodeContainer.prototype.addEffectToTheEndOfSeqAsClickEffect = function(oEffect) {
        var oPar2Lvl, oPar3Lvl;
        var sSecondLevelDelay = this.isMainSequence() ? "indefinite" : "0";
        oPar2Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, sSecondLevelDelay);
        oPar3Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "0");
        oPar3Lvl.pushToChildTnLst(oEffect);
        oPar2Lvl.pushToChildTnLst(oPar3Lvl);
        this.pushToChildTnLst(oPar2Lvl);
    };
    CTimeNodeContainer.prototype.addEffectToTheEndOfSeq = function(oEffect) {
        if(!this.isMainSequence() && !this.getSpClickInteractiveSeq()) {
            return;
        }
        if(!oEffect) {
            return;
        }
        var oPar2Lvl, oPar3Lvl;
        if(oEffect.isClickEffect()) {
            this.addEffectToTheEndOfSeqAsClickEffect(oEffect);
        }
        else if(oEffect.isAfterEffect()) {
            oPar2Lvl = this.getLastChild();
            if(!oPar2Lvl) {
                this.addEffectToTheEndOfSeqAsClickEffect(oEffect);
            }
            else {
                oPar3Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "0");
                oPar3Lvl.pushToChildTnLst(oEffect);
                oPar2Lvl.pushToChildTnLst(oPar3Lvl);
            }
        }
        else if(oEffect.isWithEffect()) {
            oPar2Lvl = this.getLastChild();
            if(!oPar2Lvl) {
                this.addEffectToTheEndOfSeqAsClickEffect(oEffect);
            }
            else {
                oPar3Lvl = oPar2Lvl.getLastChild();
                if(oPar3Lvl) {
                    oPar3Lvl.pushToChildTnLst(oEffect);
                }
                else {
                    oPar3Lvl = CTiming.prototype.createPar(NODE_FILL_HOLD, "0");
                    oPar3Lvl.pushToChildTnLst(oEffect);
                    oPar2Lvl.pushToChildTnLst(oPar3Lvl);
                }
            }
        }
        oEffect.setDelayShift();
    };
    CTimeNodeContainer.prototype.getChildrenCount = function() {
        return this.cTn.childTnLst.getLength();
    };
    CTimeNodeContainer.prototype.getChildTimeNodeByType = function(nType) {
        return this.cTn.getTimeNodeByType(nType);
    };
    CTimeNodeContainer.prototype.getChildNodeIdx = function(oChildNode) {
        return this.cTn.childTnLst.getChildIdx(oChildNode);
    };
    CTimeNodeContainer.prototype.getAllEffects = function() {
        var aAllEffects = [];
        this.traverse(function(oChild){
            if(oChild.isTimeNode() && oChild.isAnimEffect()) {
                aAllEffects.push(oChild);
            }
        });
        return aAllEffects;
    };
    CTimeNodeContainer.prototype.getLabel = function() {
        if(this.isMainSequence()) {
            return null;
        }
        var sClickSp = this.getSpClickInteractiveSeq();
        if(sClickSp) {
            var oSp = AscCommon.g_oTableId.Get_ById(sClickSp);
            if(oSp) {
                return AscCommon.translateManager.getValue("Trigger:") + " " + oSp.getObjectName();
            }
        }
        return null;
    };
    CTimeNodeContainer.prototype.getObjectName = function() {
        var sObjectId = this.getObjectId();
        var oObject = AscCommon.g_oTableId.Get_ById(sObjectId);
        if(oObject) {
            return oObject.getObjectName();
        }
        return "";
    };
    CTimeNodeContainer.prototype.asc_getStartType = function() {
        if(this.cTn) {
            return this.cTn.nodeType;
        }
        return null;
    };
    CTimeNodeContainer.prototype["asc_getStartType"] = CTimeNodeContainer.prototype.asc_getStartType;
    CTimeNodeContainer.prototype.asc_putStartType = function(v) {
        if(this.cTn) {
            this.cTn.nodeType = v;
        }
        return null;
    };
    CTimeNodeContainer.prototype["asc_putStartType"] = CTimeNodeContainer.prototype.asc_putStartType;
    CTimeNodeContainer.prototype.asc_getDelay = function() {
        if(AscFormat.isRealNumber(this.settingsDelay)) {
            return this.settingsDelay;
        }
        if(Array.isArray(this.merged) && this.merged.length > 0) {
            var nFirst = this.merged[0].asc_getDelay();
            var nCurDelay;
            for(var nIdx = 1; nIdx < this.merged.length; ++nIdx) {
                nCurDelay = this.merged[nIdx].asc_getDelay();
                if(nFirst !== nCurDelay) {
                    return undefined;
                }
            }
            return nFirst;
        }
        if(this.cTn) {
            return this.cTn.getDelay();
        }
        return 0;
    };
    CTimeNodeContainer.prototype["asc_getDelay"] = CTimeNodeContainer.prototype.asc_getDelay;
    CTimeNodeContainer.prototype.asc_putDelay = function(v) {
        AscFormat.ExecuteNoHistory(function() {
            this.settingsDelay = v;
            if(this.cTn) {
                return this.cTn.changeDelay(v);
            }
        }, this, []);
    };
    CTimeNodeContainer.prototype["asc_putDelay"] = CTimeNodeContainer.prototype.asc_putDelay;
    CTimeNodeContainer.prototype.getUndefiniteDuration = function() {
        if(this.cTn.endCondLst && this.cTn.endCondLst) {
            var aCond = this.cTn.endCondLst.list;
            if(aCond[0] &&  aCond[0].evt === COND_EVNT_ON_NEXT ) {
                return AscFormat.untilNextClick;
            }
        }
        return AscFormat.untilNextSlide;
    };
    CTimeNodeContainer.prototype.asc_getDuration = function() {
        var nDur = 0;
        if(Array.isArray(this.merged) && this.merged.length > 0) {
            nDur = this.merged[0].asc_getDuration();
            for(var nIdx = 1; nIdx < this.merged.length; ++nIdx) {
                if(nDur !== this.merged[nIdx].asc_getDuration()) {
                    return undefined;
                }
            }
        }
        else {
            if(this.cTn) {
                nDur = this.cTn.getEffectDuration();
            }
        }
        var oTime = new CAnimationTime(nDur);
        if(oTime.isIndefinite()) {
            return this.getUndefiniteDuration();
        }
        else {
            return nDur;
        }
    };
    CTimeNodeContainer.prototype["asc_getDuration"] = CTimeNodeContainer.prototype.asc_getDuration;
    CTimeNodeContainer.prototype.asc_getIsAutoDuration = function() {
        return (new CAnimationTime(this.asc_getDuration())).isIndefinite();
    };
    CTimeNodeContainer.prototype["asc_getIsAutoDuration"] = CTimeNodeContainer.prototype.asc_getIsAutoDuration;
    CTimeNodeContainer.prototype.asc_putDuration = function(v) {
        AscFormat.ExecuteNoHistory(function() {
            if(this.cTn) {
                this.cTn.changeEffectDuration(v);
            }
            if(Array.isArray(this.merged) && this.merged.length > 0) {
                for(var nIdx = 0; nIdx < this.merged.length; ++nIdx) {
                    this.merged[nIdx].asc_putDuration(v);
                }
            }
        }, this, []);
    };
    CTimeNodeContainer.prototype["asc_putDuration"] = CTimeNodeContainer.prototype.asc_putDuration;
    CTimeNodeContainer.prototype.asc_getRepeatCount = function() {
        var oRepeatCount = this.getRepeatCount();
        if(oRepeatCount.isDefinite()) {
            return oRepeatCount.val;
        }
        if(oRepeatCount.isIndefinite()) {
            return this.getUndefiniteDuration();
        }
        return 1000;
    };
    CTimeNodeContainer.prototype["asc_getRepeatCount"] = CTimeNodeContainer.prototype.asc_getRepeatCount;
    CTimeNodeContainer.prototype.asc_putRepeatCount = function(v) {
        AscFormat.ExecuteNoHistory(function() {
            if(this.cTn) {
                this.cTn.changeRepeatCount(v);
            }
        }, this, []);
    };
    CTimeNodeContainer.prototype["asc_putRepeatCount"] = CTimeNodeContainer.prototype.asc_putRepeatCount;
    CTimeNodeContainer.prototype.getRewind = function() {
        if(this.isAnimEffect()) {
            return this.getAttributesObject().fill === NODE_FILL_REMOVE;
        }
        return CTimeNodeBase.prototype.getRewind.call(this);
    };
    CTimeNodeContainer.prototype.asc_getRewind = function() {
        return this.getRewind();
    };
    CTimeNodeContainer.prototype["asc_getRewind"] = CTimeNodeContainer.prototype.asc_getRewind;
    CTimeNodeContainer.prototype.asc_putRewind = function(v) {
        return AscFormat.ExecuteNoHistory(function() {
                this.cTn.setFill(v === true ? NODE_FILL_REMOVE : NODE_FILL_HOLD);
            },
            this, []);
    };
    CTimeNodeContainer.prototype["asc_putRewind"] = CTimeNodeContainer.prototype.asc_putRewind;

    CTimeNodeContainer.prototype.asc_getClass = function() {
        if(AscFormat.isRealNumber(this.cTn.presetID) 
        && AscFormat.isRealNumber(this.cTn.presetClass)) {
            return this.cTn.presetClass;
        }
        return undefined;
    };
    CTimeNodeContainer.prototype["asc_getClass"] = CTimeNodeContainer.prototype.asc_getClass;
    CTimeNodeContainer.prototype.asc_getType = function() {
        if(this.isMultipleEffect()) {
            return AscFormat.ANIM_PRESET_MULTIPLE;
        }
        if(typeof this.cTn.presetClass === "undefined" || typeof this.cTn.presetID === "undefined") {
            return AscFormat.ANIM_PRESET_NONE;
        }
        return this.cTn.presetID;
    };
    CTimeNodeContainer.prototype["asc_getType"] = CTimeNodeContainer.prototype.asc_getType;
    CTimeNodeContainer.prototype.asc_getSubtype = function() {
        return this.isMultipleEffect() ? undefined : this.cTn.presetSubtype;
    };
    CTimeNodeContainer.prototype["asc_getSubtype"] = CTimeNodeContainer.prototype.asc_getSubtype;
    CTimeNodeContainer.prototype.asc_putSubtype = function(v) {
        AscFormat.ExecuteNoHistory(function() {
            this.cTn.setPresetSubtype(v);
        }, this, []);
    };
    CTimeNodeContainer.prototype["asc_putSubtype"] = CTimeNodeContainer.prototype.asc_putSubtype;
    CTimeNodeContainer.prototype.asc_getTriggerClickSequence = function() {
        if(this.triggerClickSequence !== undefined) {
            return this.triggerClickSequence;
        }
        if(Array.isArray(this.merged) && this.merged.length > 0) {
            var bIsInMainSeq = this.merged[0].isPartOfMainSequence();
            var bCurVal;
            for(var nEffect = 1; nEffect < this.merged.length; ++nEffect) {
                bCurVal = this.merged[nEffect].isPartOfMainSequence();;
                if(bCurVal !== bIsInMainSeq) {
                    return undefined;
                }
            }
            return bIsInMainSeq;
        }
        return this.isPartOfMainSequence();
    };
    CTimeNodeContainer.prototype["asc_getTriggerClickSequence"] = CTimeNodeContainer.prototype.asc_getTriggerClickSequence;
    CTimeNodeContainer.prototype.asc_putTriggerClickSequence = function(v) {
        this.triggerClickSequence = v;
    };
    CTimeNodeContainer.prototype["asc_putTriggerClickSequence"] = CTimeNodeContainer.prototype.asc_putTriggerClickSequence;

    CTimeNodeContainer.prototype.asc_getTriggerObjectClick = function() {
        if(this.triggerObjectClick !== undefined) {
            return this.triggerObjectClick;
        }
        var sInteractiveSeq;
        if(Array.isArray(this.merged) && this.merged.length > 0) {
            sInteractiveSeq = this.merged[0].isPartOfInteractiveSeq();
            var sCurVal;
            for(var nEffect = 1; nEffect < this.merged.length; ++nEffect) {
                sCurVal = this.merged[nEffect].isPartOfInteractiveSeq();
                if(sCurVal !== sInteractiveSeq) {
                    return undefined;
                }
            }
        }
        if(!sInteractiveSeq) {
            sInteractiveSeq = this.isPartOfInteractiveSeq();
        }
        var oSp = AscCommon.g_oTableId.Get_ById(sInteractiveSeq);
        if(oSp) {
            return oSp.getObjectName();
        }
    };
    CTimeNodeContainer.prototype["asc_getTriggerObjectClick"] = CTimeNodeContainer.prototype.asc_getTriggerObjectClick;
    CTimeNodeContainer.prototype.asc_putTriggerObjectClick = function(v) {
        this.triggerObjectClick = v;
    };
    CTimeNodeContainer.prototype["asc_putTriggerObjectClick"] = CTimeNodeContainer.prototype.asc_putTriggerObjectClick;
    CTimeNodeContainer.prototype.isEqualProperties = function(oPr) {
        if(!oPr) {
            return false;
        }
        if(this.asc_getStartType() !== oPr.asc_getStartType()) {
            return false;
        }
        if(this.asc_getDelay() !== oPr.asc_getDelay()) {
            return false;
        }
        if(this.asc_getDuration() !== oPr.asc_getDuration()) {
            return false;
        }
        if(this.asc_getRepeatCount() !== oPr.asc_getRepeatCount()) {
            return false;
        }
        if(this.asc_getRewind() !== oPr.asc_getRewind()) {
            return false;
        }
        if(this.asc_getClass() !== oPr.asc_getClass()) {
            return false;
        }
        if(this.asc_getType() !== oPr.asc_getType()) {
            return false;
        }
        if(this.asc_getSubtype() !== oPr.asc_getSubtype()) {
            return false;
        }
        if(this.asc_getTriggerClickSequence() !== oPr.asc_getTriggerClickSequence()) {
            return false;
        }
        if(this.asc_getTriggerObjectClick() !== oPr.asc_getTriggerObjectClick()) {
            return false;
        }
        return true;
    };
    CTimeNodeContainer.prototype.isSelected = function() {
        if(this.isAnimEffect()) {
            return this.selected === true;
        }
        return false;
    };
    CTimeNodeContainer.prototype.select = function() {
        this.selected = true;
    };
    CTimeNodeContainer.prototype.deselect = function() {
        this.selected = false;
    };


    function CPar() {
        CTimeNodeContainer.call(this);
        this.cTn = null;
    }
    InitClass(CPar, CTimeNodeContainer, AscDFH.historyitem_type_Par);
    CPar.prototype.activateChildrenCallback = function(oPlayer) {
        var oThis = this;
        var aChildren = oThis.getChildrenTimeNodes();
        var nChild;
        for(nChild = 0; nChild < aChildren.length; ++nChild) {
            aChildren[nChild].scheduleStart(oPlayer);
        }
    };
    CPar.prototype.createEffect = function(sObjectId, nPresetClass, nPresetId, nPresetSubtype) {
    };

    function CExcl() {//par, excl
        CTimeNodeContainer.call(this);
        this.cTn = null;
    }
    InitClass(CExcl, CTimeNodeContainer, AscDFH.historyitem_type_Excl);
    CExcl.prototype.activateChildrenCallback = function(oPlayer) {
        //TODO:
        var oThis = this;
        var aChildren = oThis.getChildrenTimeNodes();
        var nChild;
        for(nChild = 0; nChild < aChildren.length; ++nChild) {
            aChildren[nChild].scheduleStart(oPlayer);
        }
    };

    var NEXT_AC_NONE = 0;
    var NEXT_AC_SEEK = 1;

    var PREV_AC_NONE = 0;
    var PREV_AC_SKIP_TIMED = 1;

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
    CSeq.prototype.activateChildrenCallback = function(oPlayer) {
        var oThis = this;
        var aChildren = oThis.getChildrenTimeNodes();
        if(aChildren.length > 0) {
            var oFistChild = aChildren[0];
            oThis.scheduleNext(oPlayer);
            oThis.schedulePrev(oPlayer);
            var oAnimEvent = new CAnimEvent(function(){
                oFistChild.activateCallback(oPlayer);
            }, oFistChild.getStartTrigger(oPlayer), oFistChild);
            oPlayer.scheduleEvent(oAnimEvent);
        }
    };
    CSeq.prototype.scheduleNext = function(oPlayer) {
        if(this.nextCondLst) {
            var oThis = this;
            var oComplexTrigger = this.nextCondLst.createComplexTrigger(oPlayer);
            var aChildren = oThis.getChildrenTimeNodes();


            oComplexTrigger.addTrigger(function () {
               var oLastChild = aChildren[aChildren.length - 1];
               if(oLastChild) {
                   if(oLastChild.isAtEnd()) {
                       return false;
                   }
                   if(oLastChild.isActive()) {
                       var oSimpleDuration = oLastChild.simpleDuration;
                       if(oSimpleDuration && (oSimpleDuration.isIndefinite() || oSimpleDuration.isUnresolved())) {
                           return false;
                       }
                   }
               }
               return true;
            });
            var oEvent = new CAnimEvent(function() {
                for(var nChild = aChildren.length - 1; nChild > -1; --nChild) {
                    var oChild = aChildren[nChild];
                    if(!oChild.isIdle()) {
                        break;
                    }
                }
                if(nChild > -1) {
                    if(!oChild.isAtEnd()) {
                        if(oThis.concurrent !== true) {
                            oChild.getEndCallback(oPlayer)();
                        }
                        else {
                            if(oThis.nextAc === NEXT_AC_SEEK) {
                                var bFreeze = true;
                                oChild.traverseTimeNodes(function(oNode) {
                                    if(!bFreeze) {
                                        return
                                    }
                                    if(oNode.isAnimEffect()) {
                                        if(oNode.asc_getRepeatCount() === AscFormat.untilNextSlide) {
                                            bFreeze = false;
                                        }
                                    }
                                })
                                if(bFreeze) {
                                    oPlayer.bDoNotRestart = true;
                                    oChild.freezeCallback(oPlayer);
                                    delete oPlayer.bDoNotRestart;
                                }
                            }
                        }
                    }

                }
                if(nChild + 1 < aChildren.length) {
                    aChildren[nChild + 1].activateCallback(oPlayer);
                }
                else {
                    if(!oThis.isMainSequence()) {
                        oThis.freezeCallback(oPlayer);
                    }
                }
                if(oThis.isActive()) {
                    oThis.scheduleNext(oPlayer);
                }
            }, oComplexTrigger, this);
            oPlayer.scheduleEvent(oEvent);
        }
    };
    CSeq.prototype.findLastNoIdleNode = function() {
        var aChildren = this.getChildrenTimeNodes();
        for(var nChild = aChildren.length - 1; nChild > -1; --nChild) {
            var oChild = aChildren[nChild];
            if(!oChild.isIdle()) {
                return nChild;
            }
        }
        return -1;
    };
    CSeq.prototype.schedulePrev = function(oPlayer) {
        if(this.prevCondLst) {
            var oThis = this;
            var oComplexTrigger = this.prevCondLst.createComplexTrigger(oPlayer);
            oComplexTrigger.addTrigger(function() {
                var nChild = oThis.findLastNoIdleNode();
                if(nChild > -1) {
                    return true;
                }
                return false;
            });
            var oEvent = new CAnimEvent(function() {
                var nChild = oThis.findLastNoIdleNode();
                if(nChild > -1) {
                    var oChild = oThis.getChildNode(nChild);
                    if(oChild) {
                        oChild.getEndCallback(oPlayer)();
                        oChild.resetState();
                    }
                }
                oThis.schedulePrev(oPlayer);
            }, oComplexTrigger, this);
            oPlayer.scheduleEvent(oEvent);
        }
    };

    changesFactory[AscDFH.historyitem_SetCBhvr] = CChangeObject;
    changesFactory[AscDFH.historyitem_SetTo] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_SetCBhvr] = function(oClass, value) {oClass.cBhvr = value;};
    drawingsChangesMap[AscDFH.historyitem_SetTo] = function(oClass, value) {oClass.to = value;};

    function CSet() {
        CTimeNodeBase.call(this);
        this.cBhvr = null;
        this.to = null;
    }
    InitClass(CSet, CTimeNodeBase, AscDFH.historyitem_type_Set);
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
    CSet.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        if(!this.to) {
            return;
        }
        if(this.checkRemoveAtEnd()) {
            return;
        }
        this.setAttributesValue(oAttributes, this.to.getVal());
    };
    CSet.prototype.doesHideObject = function() {
        var oAttributes = {};
        this.setAttributesValue(oAttributes, this.to.getVal());
        if(oAttributes["style.visibility"] === "visible") {
            var oCurNode = this;
            var oParentNode;
            while(oParentNode = oCurNode.getParentTimeNode()) {
                var oAttrObject = oParentNode.getAttributesObject();
                if(AscFormat.PRESET_CLASS_ENTR === oAttrObject.presetClass) {
                    return true;
                }
                oCurNode = oParentNode;
            }
        }
        return false;
    };
    CSet.prototype.doesShowObject = function() {
        var oAttributes = {};
        this.setAttributesValue(oAttributes, this.to.getVal());
        if(oAttributes["style.visibility"] === "hidden") {
            var oCurNode = this;
            var oParentNode;
            while(oParentNode = oCurNode.getParentTimeNode()) {
                var oAttrObject = oParentNode.getAttributesObject();
                if(AscFormat.PRESET_CLASS_EXIT === oAttrObject.presetClass) {
                    return true;
                }
                oCurNode = oParentNode;
            }
        }
        return false;
    };
    CSet.prototype.createSetVisibility = function(sObjectId, bVisible) {
        var oBhvr = new CCBhvr();
        this.setCBhvr(oBhvr);
        var oCTn = this.createCCTn("1", NODE_FILL_HOLD, 0, null, null, null);
        oBhvr.setCTn(oCTn);
        var oTgtEl = this.createSpTgt(sObjectId);
        oBhvr.setTgtEl(oTgtEl);
        var oAttrLst = new CAttrNameLst();
        oBhvr.setAttrNameLst(oAttrLst);
        var oAttr = new CAttrName();
        oAttrLst.addToLst(0, oAttr);
        oAttr.setText("style.visibility");
        var oTo = new CAnimVariant();
        this.setTo(oTo);
        oTo.setStrVal(bVisible ? "visible" : "hidden");
        return this;
    };


    changesFactory[AscDFH.historyitem_VideoCMediaNode] = CChangeObject;
    changesFactory[AscDFH.historyitem_VideoFullScrn] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_VideoCMediaNode] = function(oClass, value) {oClass.cMediaNode = value;};
    drawingsChangesMap[AscDFH.historyitem_VideoFullScrn] = function(oClass, value) {oClass.fullScrn = value;};
    function CVideo() {//par, excl
        CTimeNodeBase.call(this);
        this.cMediaNode = null;
        this.fullScrn = null;
    }
    InitClass(CVideo, CTimeNodeBase, AscDFH.historyitem_type_Video);
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
        CBaseAnimObject.call(this);
        this.lvl = null;
        this.type = null;
    }
    InitClass(COleChartEl, CBaseAnimObject, AscDFH.historyitem_type_OleChartEl);
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
        CBaseAnimObject.call(this);
        this.x = null;
        this.y = null;
    }
    InitClass(CTLPoint, CBaseAnimObject, AscDFH.historyitem_type_TlPoint);
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
        CBaseAnimObject.call(this);
        this.endSnd = null;
        this.stSnd = null;
    }
    InitClass(CSndAc, CBaseAnimObject, AscDFH.historyitem_type_SndAc);
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
        CBaseAnimObject.call(this);
        this.snd = null;
        this.loop = null;
    }
    InitClass(CStSnd, CBaseAnimObject, AscDFH.historyitem_type_StSnd);
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
        CBaseAnimObject.call(this);
        this.charRg = null;//CIndexRg
        this.pRg = null;
    }
    InitClass(CTxEl, CBaseAnimObject, AscDFH.historyitem_type_TxEl);
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
        CBaseAnimObject.call(this);
        this.spokes = null;
    }
    InitClass(CWheel, CBaseAnimObject, AscDFH.historyitem_type_Wheel);
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
        CBaseAnimObject.call(this);
        this.text = null;
    }
    InitClass(CAttrName, CBaseAnimObject, AscDFH.historyitem_type_AttrName);
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

    function CExternalEvent(oEventsProcessor, type, target) {
        this.eventsProcessor = oEventsProcessor;
        this.type = type;
        this.target = target;
        this.handledNodes = [];
    }
    CExternalEvent.prototype.isEqual = function(oEvent) {
        if(!oEvent) {
            return false;
        }
        if(this.target !== oEvent.target) {
            return false;
        }
        if(this.isEqualType(oEvent.type)) {
            return true;
        }
        return false;
    };
    CExternalEvent.prototype.isEqualType = function(nType) {
        if(this.type === nType) {
            return true;
        }
        if((nType === COND_EVNT_ON_NEXT || nType === COND_EVNT_ON_CLICK) &&
            (this.type === COND_EVNT_ON_NEXT || this.type === COND_EVNT_ON_CLICK)) {
            return true;
        }
        return false;
    };
    CExternalEvent.prototype.log = function(sPrefix) {
        //console.log(sPrefix + " | EXTERNAL EVENT TYPE: " + EVENT_DESCR_MAP[this.type] + " | TARGET: " + this.target);
    };

    function CEventsProcessor(player) {
        this.player = player;
        this.event = null;
    }
    CEventsProcessor.prototype.addEvent = function(oEvent) {
        this.event = oEvent;
        this.player.onFrame();
        return oEvent.handledNodes.length > 0;
    };
    CEventsProcessor.prototype.clear = function() {
        this.event = null;
    };
    CEventsProcessor.prototype.onFrame = function() {
        this.clear();
    };
    CEventsProcessor.prototype.getExternalEvent = function () {
        return this.event;
    };

    var PLAYER_STATE_IDLE = 0;
    var PLAYER_STATE_PLAYING = 1;
    var PLAYER_STATE_PAUSING = 2;
    var PLAYER_STATE_DONE = 3;

    function CAnimationTimer(player) {
        this.player = player;
        this.elapsed = null;
        this.lastTime = null;

        this.elapsedTime = new CAnimationTime();

        //this.lastFire = null;

        this.frameId = null;
    }
    CAnimationTimer.prototype.start = function () {
        if(this.isStarted()) {
            return;
        }
        if(this.isStopped()) {
            this.elapsed = 0;
        }
        this.lastTime = (new Date()).getTime();
        this.startFrames();
    };
    CAnimationTimer.prototype.stop = function () {
        if(this.isStopped()) {
            return;
        }
        this.elapsed = null;
        this.lastTime = null;
        this.stopFrames();
       // this.lastFire = null;
      //  console.log("Timer is stopped");
    };
    CAnimationTimer.prototype.pause = function () {
        if(!this.isStarted()) {
            return;
        }
        this.lastTime = null;
        //this.lastFire = null;
        this.stopFrames();
    };
    CAnimationTimer.prototype.getElapsedTicks = function () {
        if(this.isStopped()) {
            return 0;
        }
        return this.elapsed;
    };
    CAnimationTimer.prototype.checkTimeObject = function () {
        this.elapsedTime.setVal(this.getElapsedTicks());
    };
    CAnimationTimer.prototype.getElapsedTime = function () {
        this.checkTimeObject();
        return this.elapsedTime;
    };
    CAnimationTimer.prototype.isPaused = function () {
        if(this.elapsed !== null && this.lastTime === null) {
            return true;
        }
        return false;
    };
    CAnimationTimer.prototype.isStopped = function () {
        if(this.elapsed === null && this.lastTime === null) {
            return true;
        }
        return false;
    };
    CAnimationTimer.prototype.isStarted = function () {
        return !this.isPaused() && !this.isStopped();
    };
    CAnimationTimer.prototype.onFrame = function () {
        if(this.isStarted()) {
            var nCurTime = (new Date()).getTime();
            var nDiff = nCurTime - this.lastTime;
            this.elapsed += nDiff;
            this.lastTime = nCurTime;


            //for test
            //if(this.lastFire === null || this.elapsed - this.lastFire >= 5000) {
            //    this.lastFire = this.elapsed;
            //  //  console.log("Timer is ON " + this.elapsed)
            //}
        }
    };
    CAnimationTimer.prototype.startFrames = function () {
        this.stopFrames();
        this.nextFrame();
    };
    CAnimationTimer.prototype.stopFrames = function () {
        if(this.frameId !== null) {
            __cancelFrame(this.frameId);
            this.frameId = null;
        }
    };
    CAnimationTimer.prototype.nextFrame = function () {
        var oThis = this;
        this.frameId = __nextFrame(function () {
            oThis.player.onFrame();
            oThis.nextFrame();
        })
    };

    function CAnimComplexTrigger(param) {
        this.triggers = [];
        this.addDefault();
        if(Array.isArray(param)) {
            this.addTriggers(param);
        }
        else if(typeof param === "function") {
            this.addTrigger(param);
        }
    }
    CAnimComplexTrigger.prototype.areTriggersFired = function(oPlayer) {
        var oExternalEvent = oPlayer.getExternalEvent();
        var nOldHandledNodes = null;
        if(oExternalEvent) {
            nOldHandledNodes = oExternalEvent.handledNodes.length;
        }
        for(var nTrigger = 0; nTrigger < this.triggers.length; ++nTrigger) {
            if(!this.triggers[nTrigger]()) {
                if(oExternalEvent && nOldHandledNodes !== null && oExternalEvent.handledNodes.length !== nOldHandledNodes) {
                    oExternalEvent.handledNodes.length = nOldHandledNodes;
                }
                return false;
            }
        }
        return true;
    };
    CAnimComplexTrigger.prototype.isFired = function(oPlayer) {
        return this.triggers.length > 0 && this.areTriggersFired(oPlayer);
    };
    CAnimComplexTrigger.prototype.addDefault = function() {
        this.addTrigger(DEFAULT_SIMPLE_TRIGGER);
    };
    CAnimComplexTrigger.prototype.addNever = function() {
        this.addTrigger(DEFAULT_NEVER_TRIGGER);
    };
    CAnimComplexTrigger.prototype.addTrigger = function(fTrigger) {
        if(fTrigger !== null) {
            this.triggers.push(fTrigger);
        }
    };
    CAnimComplexTrigger.prototype.addTriggers = function(aTriggers) {
        for(var nTrigger = 0; nTrigger < aTriggers.length; ++nTrigger) {
            this.addTrigger(aTriggers[nTrigger]);
        }
    };
    CAnimComplexTrigger.prototype.isDefault = function() {
        return this.triggers.length === 1 && this.triggers[0] === DEFAULT_SIMPLE_TRIGGER;
    };


    function CAnimEvent(fCallback, oTrigger, oCaller) {
        this.trigger = oTrigger;
        this.callback = fCallback;
        this.fired = false;
        this.caller = oCaller;
        this.scheduler = null;
    }
    CAnimEvent.prototype.setScheduler = function(oScheduler) {
        this.scheduler = oScheduler;
    };
    CAnimEvent.prototype.isScheduled = function() {
        return this.scheduler !== null;
    };
    CAnimEvent.prototype.fire = function() {
        this.callback.call();
        this.fired = true;
        //this.caller.logState("FIRE CALLBACK");
    };
    CAnimEvent.prototype.checkTrigger = function(oPlayer) {
        return this.trigger.isFired(oPlayer);
    };
    CAnimEvent.prototype.checkCaller = function(oCaller) {
        if(this.caller === oCaller) {
            return true;
        }
        return false;
    };

    function CAnimationScheduler(player) {
        this.player = player;
        this.events = [];
    }
    CAnimationScheduler.prototype.onFrame = function() {
        this.handleEvents();
    };
    CAnimationScheduler.prototype.handleEvents = function() {
        var nEvent = 0;
        while(nEvent < this.events.length) {
            var oEvent = this.events[nEvent];
            if(oEvent.checkTrigger(this.player)) {
                this.events.splice(nEvent, 1);
                oEvent.fire();
                nEvent = 0;
            }
            else {
                ++nEvent;
            }
        }
        return false;
    };
    CAnimationScheduler.prototype.addEvent = function(oEvent) {
        this.events.push(oEvent);
        oEvent.setScheduler(this);
    };
    CAnimationScheduler.prototype.removeEvent = function(oEvent) {
        for(var nEvent = 0; nEvent < this.events.length; ++nEvent) {
            if(this.events[nEvent] === oEvent) {
                this.events.splice(nEvent, 1);
                oEvent.setScheduler(null);
                return;
            }
        }
    };
    CAnimationScheduler.prototype.cancelAll = function() {
        this.events.length = 0;
    };
    CAnimationScheduler.prototype.stop = function() {
        this.cancelAll();
    };
    CAnimationScheduler.prototype.cancelCallerEvents = function(oCaller) {
        for(var nCallbacks = this.events.length - 1; nCallbacks > -1; --nCallbacks) {
            if(this.events[nCallbacks].checkCaller(oCaller)) {
                var oEvent = this.events.splice(nCallbacks, 1)[0];
                oEvent.setScheduler(null);
            }
        }
    };
    CAnimationScheduler.prototype.getElapsedTicks = function() {
        return this.player.getElapsedTicks();
    };
    CAnimationScheduler.prototype.hasScheduledEvents = function() {
        return this.events.length > 0;
    };

    function shuffleArray(array) {
        for (var i = array.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = array[i];
            array[i] = array[j];
            array[j] = temp;
        }
    }

    var RANDOM_BARS_ARRAY = [62, 4, 27, 42, 80, 34, 67, 20, 74, 32, 10, 54, 3, 77, 36, 55, 26, 53, 97, 90, 68, 65, 57, 12, 52, 70, 23, 64, 30, 73, 79, 22, 14, 51, 9, 0, 49, 1, 15, 71, 93, 86, 19, 28, 45, 41, 39, 60, 25, 7, 92, 46, 2, 98, 33, 40, 31, 72, 69, 24, 75, 84, 43, 47, 87, 50, 18, 56, 13, 61, 76, 17, 91, 37, 8, 11, 78, 6, 5, 48, 59, 95, 66, 63, 81, 96, 35, 88, 94, 89, 38, 99, 82, 29, 16, 83, 21, 58, 44, 85];
    var STRIPS_COUNT = 16;

    function CBaseAnimTexture(oCanvas, fScale, nX, nY) {
        this.canvas = oCanvas;
        this.scale = fScale;
        this.x = nX;
        this.y = nY;
    }

    function CAnimTexture(oCache, oCanvas, fScale, nX, nY) {
        CBaseAnimTexture.call(this, oCanvas, fScale, nX, nY);
        this.cache = oCache;
        this.effectTexture = null;
    }
    CAnimTexture.prototype.checkScale = function(fScale) {
        if(!AscFormat.fApproxEqual(this.scale, fScale)) {
            return false;
        }
        return true;
    };
    CAnimTexture.prototype.draw = function(oGraphics, oTransform) {
        var bNoTransform = false;
        if(!oTransform) {
            bNoTransform = true;
        }
        else {
            // if(oTransform.IsIdentity2()) {
            //     var fDelta = 2;
            //     if(AscFormat.fApproxEqual(oTransform.tx*this.scale, this.x, fDelta) &&
            //         AscFormat.fApproxEqual(oTransform.ty*this.scale, this.y, fDelta)) {
            //         bNoTransform = true;
            //     }
            // }
        }
        if(bNoTransform) {
            oGraphics.SaveGrState();
            oGraphics.SetIntegerGrid(true);
            var nDx = oGraphics.m_oCoordTransform.tx;
            var nDy = oGraphics.m_oCoordTransform.ty;
            oGraphics.m_oContext.drawImage(this.canvas, (nDx + this.x + 0.5) >> 0, (nDy + this.y + 0.5) >> 0, this.canvas.width, this.canvas.height);
            oGraphics.RestoreGrState();
            oGraphics.FreeFont();
        }
        else {
            oGraphics.SaveGrState();
            oGraphics.SetIntegerGrid(false);
            oGraphics.transform3(oTransform, false);
            oGraphics.drawImage2(this.canvas, 0, 0, this.canvas.width / this.scale, this.canvas.height / this.scale);
            oGraphics.RestoreGrState();
            oGraphics.FreeFont();
        }
    };
    CAnimTexture.prototype.createEffectTexture = function(oEffect) {
        if(!oEffect) {
            return this;
        }
        var aFilters = oEffect.filters;
        var oEffectData = oEffect.data;
        var dTime = oEffectData.time;
        var nTransition = oEffectData.transition;
        for(var nFilter = 0; nFilter < aFilters.length; ++nFilter) {
            var nFilterType = aFilters[nFilter];
            switch (nFilterType) {
                case FILTER_TYPE_BLINDS_HORIZONTAL: {
                    return this.createBlindsHorizontal(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_BLINDS_VERTICAL: {
                    return this.createBlindsVertical(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_BOX_IN: {
                    return this.createBoxIn(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_BOX_OUT: {
                    return this.createBoxOut(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_CHECKERBOARD_ACROSS: {
                    return this.createCheckerBoardAcross(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_CHECKERBOARD_DOWN: {
                    return this.createCheckerBoardDown(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_CIRCLE:
                case FILTER_TYPE_CIRCLE_IN: {
                    return this.createCircleIn(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_CIRCLE_OUT: {
                    return this.createCircleOut(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_DIAMOND:
                case FILTER_TYPE_DIAMOND_IN: {
                    return this.createDiamondIn(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_DIAMOND_OUT: {
                    return this.createDiamondOut(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_DISSOLVE: {
                    return this.createDissolve(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_FADE: {
                    return this.createFade(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_SLIDE_FROM_TOP: {
                    return this.createSlideFromTop(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_SLIDE_FROM_BOTTOM: {
                    return this.createSlideFromBottom(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_SLIDE_FROM_LEFT: {
                    return this.createSlideFromLeft(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_SLIDE_FROM_RIGHT: {
                    return this.createSlideFromRight(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_PLUS_IN: {
                    return this.createPlusIn(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_PLUS_OUT: {
                    return this.createPlusOut(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_BARN_IN_VERTICAL: {
                    return this.createBarnInVertical(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_BARN_IN_HORIZONTAL: {
                    return this.createBarnInHorizontal(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_BARN_OUT_VERTICAL: {
                    return this.createBarnOutVertical(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_BARN_OUT_HORIZONTAL: {
                    return this.createBarnOutHorizontal(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_RANDOM_BARS_HORIZONTAL: {
                    return this.createRandomBarsHorizontal(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_RANDOM_BARS_VERTICAL: {
                    return this.createRandomBarsVertical(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_STRIPS_DOWN_LEFT: {
                    return this.createStripsDownLeft(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_STRIPS_UP_LEFT: {
                    return this.createStripsUpLeft(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_STRIPS_DOWN_RIGHT: {
                    return this.createStripsDownRight(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_STRIPS_UP_RIGHT: {
                    return this.createStripsUpRight(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_SLIDE_WEDGE: {
                    return this.createWedge(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WHEEL_1: {
                    return this.createWheel1(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WHEEL_2: {
                    return this.createWheel2(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WHEEL_3: {
                    return this.createWheel3(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WHEEL_4: {
                    return this.createWheel4(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WHEEL_8: {
                    return this.createWheel8(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WIPE_RIGHT: {
                    return this.createWipeRight(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WIPE_LEFT: {
                    return this.createWipeLeft(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WIPE_DOWN: {
                    return this.createWipeDown(dTime, nTransition);
                    break;
                }
                case FILTER_TYPE_WIPE_UP: {
                    return this.createWipeUp(dTime, nTransition);
                    break;
                }
            }
        }
        return this;
    };
    CAnimTexture.prototype.createTexture = function() {
        if(!this.effectTexture) {
            var oCanvas = document.createElement('canvas');
            oCanvas.width = this.canvas.width;
            oCanvas.height = this.canvas.height;
            this.effectTexture = new CAnimTexture(this.cache, oCanvas, this.scale, this.x, this.y);
        }
        else {
            //this.effectTexture.canvas.width = this.effectTexture.canvas.width;
           var oCtx = this.effectTexture.canvas.getContext('2d');
           oCtx.globalCompositeOperation = 'source-over';
           oCtx.clearRect(0, 0, this.canvas.width, this.canvas.height);
           //oCtx.setTransform(1, 0, 0, 1, 0, 0);
        }
        return this.effectTexture;
    };
    CAnimTexture.prototype.createCopy = function() {
        var oTexture = this.createTexture();
        var oCtx = oTexture.canvas.getContext('2d');
        oCtx.drawImage(this.canvas, 0, 0);
        return oTexture;
    };
    CAnimTexture.prototype.drawRect = function(oCtx, nX, nY, nWidth, nHeight) {
        oCtx.beginPath();
        oCtx.rect(nX, nY, nWidth, nHeight);
        oCtx.closePath();
        oCtx.fill();
    };
    CAnimTexture.prototype.createBlindsHorizontal = function(fTime, nTransition) {
        var fResultTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResultTime = fTime;
        }
        else {
            fResultTime = 1 - fTime;
        }
        var nRows = 6;
        var nVertStride = this.canvas.height / nRows + 0.5 >> 0;
        var nWidth = this.canvas.width;
        var nHeight = nVertStride * fResultTime + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = 'destination-in';
        var nY;
        oCtx.beginPath();
        for(var nRow = 0; nRow < nRows; ++nRow) {
            nY = nVertStride * nRow;
            oCtx.rect(0, nY, nWidth, nHeight);
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createBlindsVertical = function(fTime, nTransition) {
        var fResultTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResultTime = fTime;
        }
        else {
            fResultTime = 1 - fTime;
        }
        var nCols = 6;
        var nHorStride = this.canvas.width / nCols + 0.5 >> 0;
        var nWidth = nHorStride * fResultTime + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var nHeight = this.canvas.height;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = 'destination-in';
        var nY;
        oCtx.beginPath();
        for(var nCol = 0; nCol < nCols; ++nCol) {
            var nX = nHorStride * nCol;
            oCtx.rect(nX, 0, nWidth, nHeight);
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createBoxIn = function(fTime, nTransition) {
        var sOperationType;
        var fEffectTime = (1 - fTime);
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperationType = 'destination-out';
        }
        else {
            sOperationType = 'destination-in';
        }
        var nBoxW = this.canvas.width * fEffectTime + 0.5 >> 0;
        var nBoxH = this.canvas.height * fEffectTime + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = sOperationType;
        var nX = (this.canvas.width - nBoxW) / 2 + 0.5 >> 0;
        var nY = (this.canvas.height - nBoxH) / 2 + 0.5 >> 0;
        this.drawRect(oCtx, nX, nY, nBoxW, nBoxH);
        return oTexture;
    };
    CAnimTexture.prototype.createBoxOut = function(fTime, nTransition) {
        var sOperationType;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperationType = 'destination-in';
        }
        else {
            sOperationType = 'destination-out';
        }
        var nBoxW = this.canvas.width * fTime + 0.5 >> 0;
        var nBoxH = this.canvas.height * fTime + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = sOperationType;
        var nX = (this.canvas.width - nBoxW) / 2 + 0.5 >> 0;
        var nY = (this.canvas.height - nBoxH) / 2 + 0.5 >> 0;
        this.drawRect(oCtx, nX, nY, nBoxW, nBoxH);
        return oTexture;
    };
    CAnimTexture.prototype.createCheckerBoardAcross = function(fTime, nTransition) {
        var nRows = 6;
        var nCols = nRows;
        var nHorStride = this.canvas.width / nCols + 0.5 >> 0;
        var nHalfHorStride = nHorStride / 2 + 0.5 >> 0;
        var nVertStride = this.canvas.height / nRows + 0.5 >> 0;
        var fResultTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResultTime = (1 - fTime);
        }
        else {
            fResultTime = fTime;
        }
        var nWidth = nHorStride * fResultTime + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = 'destination-out';
        oCtx.beginPath()
        var nRow, nCol;
        var nX, nY;
        for(nRow = 0; nRow < nRows; ++nRow) {
            var bOdd = (nRow % 2) === 1;
            for(nCol = 0; nCol < nCols; ++nCol) {
                nX = (nCol + 1) * nHorStride - nWidth;
                if(bOdd) {
                    nX -= nHalfHorStride;
                }
                nY = nRow * nVertStride;
                oCtx.rect(nX, nY, nWidth, nVertStride);
            }
            if(bOdd) {
                nX = (nCol + 1) * nHorStride - nWidth - nHalfHorStride;
                nY = nRow * nVertStride;
                oCtx.rect(nX, nY, nWidth, nVertStride);
            }
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createCheckerBoardDown = function(fTime, nTransition) {
        var nRows = 6;
        var nCols = nRows;
        var nHorStride = this.canvas.width / nCols + 0.5 >> 0;
        var nVertStride = this.canvas.height / nRows + 0.5 >> 0;
        var nHalfVertStride = nVertStride / 2 + 0.5 >> 0;
        var fResultTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResultTime = (1 - fTime);
        }
        else {
            fResultTime = fTime;
        }
        var nHeight = (nVertStride * fResultTime + 0.5) >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = 'destination-out';
        var nRow, nCol;
        var nX, nY;
        oCtx.beginPath();
        for(nCol = 0; nCol < nCols; ++nCol) {
            var bOdd = (nCol % 2) === 1;
            for(nRow = 0; nRow < nRows; ++nRow) {
                nY = (nRow + 1) * nVertStride - nHeight;
                if(bOdd) {
                    nY -= nHalfVertStride;
                }
                nX = nCol * nHorStride;
                oCtx.rect(nX, nY, nHorStride, nHeight);
            }
            if(bOdd) {
                nY = (nRow + 1) * nVertStride - nHeight - nHalfVertStride;
                nX = nCol * nHorStride;
                oCtx.rect(nX, nY, nHorStride, nHeight);
            }
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createCircle = function(fTime, sOperation) {
        var nMaxRadius = this.canvas.width * Math.SQRT1_2;
        var nRadius = nMaxRadius * fTime;
        if(nRadius === 0) {
            return this;
        }
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = sOperation;
        oCtx.beginPath();
        var nX = this.canvas.width / 2 + 0.5 >> 0;
        var nY = this.canvas.height / 2 + 0.5 >> 0;
        var nRadiusX = nRadius;
        var nRadiusY = nRadiusX * (this.canvas.height / this.canvas.width) + 0.5 >> 0;
        var fRotation = 0;
        var fStartAngle = 0;
        var fEndAngle = 2 * Math.PI;
        var bCounterclockwise = false;
        oCtx.ellipse(nX, nY, nRadiusX, nRadiusY, fRotation, fStartAngle, fEndAngle, bCounterclockwise);
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createCircleIn = function(fTime, nTransition) {
        var sOperation;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperation = "destination-out";
        }
        else {
            sOperation = "destination-in";
        }
        return this.createCircle(1 - fTime, sOperation);
    };
    CAnimTexture.prototype.createCircleOut = function(fTime, nTransition) {
        var sOperation;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperation = "destination-in";
        }
        else {
            sOperation = "destination-out";
        }
        return this.createCircle(fTime, sOperation);
    };
    CAnimTexture.prototype.createStripsUpRightDiag = function(fTime, sOperation) {
        var nWidth = this.canvas.width / STRIPS_COUNT;
        var nHeight = this.canvas.height / STRIPS_COUNT;
        var nCount = 2*this.canvas.width * fTime / nWidth + 0.5 >> 0;
        if(nCount === 0 && "destination-out" === sOperation ||
            AscFormat.fApproxEqual(fTime, 1) && "destination-in" === sOperation) {
            return this;
        }
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var nX = this.canvas.width - nWidth * nCount;
        var nY = 0;
        oCtx.globalCompositeOperation = sOperation;
        oCtx.beginPath();
        oCtx.moveTo(this.canvas.width, 0);
        oCtx.lineTo(nX, nY);
        for(var nRect = 0; nRect < nCount; ++nRect) {
            oCtx.lineTo(nX, nY + nHeight);
            oCtx.lineTo(nX + nWidth, nY + nHeight);
            nX += nWidth;
            nY += nHeight;
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createStripsUpRight = function(fTime, nTransition) {
        var fResTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResTime = 1 - fTime;
        }
        else {
            fResTime = fTime;
        }
        return this.createStripsUpRightDiag(fResTime, "destination-out");
    };
    CAnimTexture.prototype.createStripsDownLeft = function(fTime, nTransition) {
        var nWidth = this.canvas.width / STRIPS_COUNT;
        var nHeight = this.canvas.height / STRIPS_COUNT;
        var fResTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResTime = fTime;
        }
        else {
            fResTime = 1 - fTime;
        }
        var nCount = 2*this.canvas.width * fResTime / nWidth + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var nX = this.canvas.width - nWidth * nCount;
        var nY = 0;
        oCtx.globalCompositeOperation = "destination-out";
        oCtx.beginPath();
        oCtx.moveTo(nX, nY);
        for(var nRect = 0; nRect < nCount; ++nRect) {
            oCtx.lineTo(nX, nY + nHeight);
            oCtx.lineTo(nX + nWidth, nY + nHeight);
            nX += nWidth;
            nY += nHeight;
        }
        oCtx.lineTo(this.canvas.width, 2*this.canvas.height);
        oCtx.lineTo(-this.canvas.width, 2*this.canvas.height);
        oCtx.lineTo(-this.canvas.width, 0);
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createStripsDownRight = function(fTime, nTransition) {
        var nWidth = this.canvas.width / STRIPS_COUNT;
        var nHeight = this.canvas.height / STRIPS_COUNT;

        var fResTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResTime = fTime;
        }
        else {
            fResTime = 1 - fTime;
        }
        var nCount = 2*this.canvas.width * fResTime / nWidth + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var nX = nWidth * nCount;
        var nY = 0;
        oCtx.globalCompositeOperation = "destination-out";
        oCtx.beginPath();
        oCtx.moveTo(nX, nY);
        for(var nRect = 0; nRect < nCount; ++nRect) {
            oCtx.lineTo(nX, nY + nHeight);
            oCtx.lineTo(nX - nWidth, nY + nHeight);
            nX -= nWidth;
            nY += nHeight;
        }
        oCtx.lineTo(0, 2*this.canvas.height);
        oCtx.lineTo(2*this.canvas.width, 2*this.canvas.height);
        oCtx.lineTo(this.canvas.width, 0);
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createStripsUpLeft = function(fTime, nTransition) {
        var nWidth = this.canvas.width / STRIPS_COUNT;
        var nHeight = this.canvas.height / STRIPS_COUNT;

        var fResTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResTime = fTime;
        }
        else {
            fResTime = 1 - fTime;
        }
        var nCount = 2*this.canvas.width * fResTime / nWidth + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var nX = this.canvas.width;
        var nY = this.canvas.height - nHeight * nCount;
        oCtx.globalCompositeOperation = "destination-out";
        oCtx.beginPath();
        oCtx.moveTo(nX, nY);
        for(var nRect = 0; nRect < nCount; ++nRect) {
            oCtx.lineTo(nX - nWidth, nY);
            oCtx.lineTo(nX - nWidth, nY + nHeight);
            nX -= nWidth;
            nY += nHeight;
        }
        oCtx.lineTo(-this.canvas.width, this.canvas.height);
        oCtx.lineTo(-this.canvas.width, -this.canvas.height);
        oCtx.lineTo(this.canvas.width, -this.canvas.height);
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createDiamond = function(fTime, sOperation) {
        var nMaxWidth = 2*this.canvas.width;
        var nWidth = nMaxWidth*fTime + 0.5 >> 0;
        if(nWidth === 0) {
            return this;
        }
        var nMaxHeight = 2*this.canvas.height;
        var nHeight = nMaxHeight*fTime + 0.5 >> 0;
        var oTexture = this.createCopy();

        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = sOperation;
        var nHalfCanvasWidth = this.canvas.width / 2 + 0.5 >> 0;
        var nHalfCanvasHeight = this.canvas.height / 2 + 0.5 >> 0;
        var nHalfWidth = nWidth / 2 + 0.5 >> 0;
        var nHalfHeight = nHeight / 2 + 0.5 >> 0;
        oCtx.beginPath();
        oCtx.moveTo(nHalfCanvasWidth, nHalfCanvasHeight - nHalfHeight);
        oCtx.lineTo(nHalfCanvasWidth + nHalfWidth, nHalfCanvasHeight);
        oCtx.lineTo(nHalfCanvasWidth, nHalfCanvasHeight + nHalfHeight);
        oCtx.lineTo(nHalfCanvasWidth - nHalfWidth, nHalfCanvasHeight);
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createDiamondIn = function(fTime, nTransition) {
        var sOperation;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperation = "destination-out";
        }
        else {
            sOperation = "destination-in";
        }
        return this.createDiamond(1 - fTime, sOperation);
    };
    CAnimTexture.prototype.createDiamondOut = function(fTime, nTransition) {
        var sOperationType;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperationType = 'destination-in';
        }
        else {
            sOperationType = 'destination-out';
        }
        return this.createDiamond(fTime, sOperationType);
    };
    CAnimTexture.prototype.getRandomRanges = function(fTime, nTransition) {
        var nFilledBars = RANDOM_BARS_ARRAY.length * fTime + 0.5 >> 0;
        if(nFilledBars === 0) {
            return [];
        }
        var aFilledBars = RANDOM_BARS_ARRAY.slice(0, nFilledBars);
        aFilledBars.sort(function(a, b) {
            return a - b;
        });
        var aFilledRanges = [];
        var aCurRange = [aFilledBars[0], aFilledBars[0]];
        aFilledRanges.push(aCurRange);
        for(var nBar = 1; nBar < aFilledBars.length; ++nBar) {
            if(aFilledBars[nBar] === (aCurRange[1] + 1)) {
                aCurRange[1] = aFilledBars[nBar];
            }
            else {
                aCurRange = [aFilledBars[nBar], aFilledBars[nBar]];
                aFilledRanges.push(aCurRange);
            }
        }
        return aFilledRanges;
    };
    CAnimTexture.prototype.createRandomBarsHorizontal = function(fTime, nTransition) {

        var fResTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResTime = 1 - fTime;
        }
        else {
            fResTime = fTime;
        }
        var aFilledRanges = this.getRandomRanges(fResTime);
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var nX, nY, nWidth, nHeight;

        oCtx.globalCompositeOperation = 'destination-out';
        oCtx.beginPath();
        for(var nRange = 0; nRange < aFilledRanges.length; ++nRange) {
            var aRange = aFilledRanges[nRange];
            nX = 0;
            nY = (aRange[0] / RANDOM_BARS_ARRAY.length) * this.canvas.height + 0.5 >> 0;
            nWidth = this.canvas.width;
            nHeight = (aRange[1] - aRange[0] + 1) / RANDOM_BARS_ARRAY.length * this.canvas.height + 0.5 >> 0;
            oCtx.fillRect(nX, nY, nWidth, nHeight);
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createRandomBarsVertical = function(fTime, nTransition) {
		var fResTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResTime = 1 - fTime;
        }
        else {
            fResTime = fTime;
        }
        var aFilledRanges = this.getRandomRanges(fResTime);
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = 'destination-out';
        var nX, nY, nWidth, nHeight;
        oCtx.beginPath();
        for(var nRange = 0; nRange < aFilledRanges.length; ++nRange) {
            var aRange = aFilledRanges[nRange];
            nX = (aRange[0] / RANDOM_BARS_ARRAY.length) * this.canvas.width + 0.5 >> 0;
            nY = 0;
            nWidth = (aRange[1] - aRange[0] + 1) / RANDOM_BARS_ARRAY.length * this.canvas.width + 0.5 >> 0;
            nHeight = this.canvas.height;
            oCtx.fillRect(nX, nY, nWidth, nHeight);
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createWedge = function(fTime, nTransition) {
        var fHalfAngle = Math.PI * fTime;
        var fAngle = 2 * fHalfAngle;
        var nRadius = Math.sqrt(this.canvas.width* this.canvas.width + this.canvas.height*this.canvas.height) / 2 + 0.5 >> 0;
        var nXCenter = this.canvas.width / 2 + 0.5 >> 0;
        var nYCenter = this.canvas.height / 2 + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var sOperation;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperation = 'destination-in';
        }
        else {
            sOperation = "destination-out";
        }
        oCtx.globalCompositeOperation = sOperation;
        var nX1 = nXCenter + (nRadius * Math.cos(fHalfAngle - Math.PI / 2) + 0.5 >> 0);
        var nY1 = nYCenter + (nRadius * Math.sin(fHalfAngle - Math.PI / 2) + 0.5 >> 0);
        oCtx.beginPath();
        oCtx.moveTo(nXCenter, nYCenter);
        oCtx.lineTo(nX1, nY1);
        oCtx.arc(nXCenter, nYCenter, nRadius, fHalfAngle - Math.PI / 2, -fHalfAngle - Math.PI / 2, true);
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createWheel1 = function (fTime, nTransition) {
        return this.createWheel(fTime, 1, nTransition);
    };
    CAnimTexture.prototype.createWheel2 = function (fTime, nTransition) {
        return this.createWheel(fTime, 2, nTransition);
    };
    CAnimTexture.prototype.createWheel3 = function (fTime, nTransition) {
        return this.createWheel(fTime, 3, nTransition);
    };
    CAnimTexture.prototype.createWheel4 = function (fTime, nTransition) {
        return this.createWheel(fTime, 4, nTransition);
    };
    CAnimTexture.prototype.createWheel8 = function (fTime, nTransition) {
        return this.createWheel(fTime, 8, nTransition);
    };
    CAnimTexture.prototype.createWheel = function(fTime, nCount, nTransition) {

        var sOperation;
        if(nTransition === TRANSITION_TYPE_IN) {
            if(AscFormat.fApproxEqual(fTime, 1.0)) {
                return this;
            }
            sOperation = 'destination-in';

        }
        else {
            if(AscFormat.fApproxEqual(fTime, 0.0)) {
                return this;
            }
            sOperation = "destination-out";
        }
        var fStride = 2 * Math.PI / nCount;
        var fAngle = fStride * fTime + 0.001;
        var nRadius = Math.sqrt(this.canvas.width*this.canvas.width + this.canvas.height*this.canvas.height) / 2 + 0.5 >> 0;
        var nXCenter = this.canvas.width / 2 + 0.5 >> 0;
        var nYCenter = this.canvas.height / 2 + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = sOperation;
        oCtx.beginPath();
        for(var nAngle = 0; nAngle < nCount; ++nAngle) {
            var fStartAngle = fStride * nAngle - Math.PI / 2 + fAngle;
            var nX1 = nXCenter + (nRadius * Math.cos(fStartAngle) + 0.5 >> 0);
            var nY1 = nYCenter + (nRadius * Math.sin(fStartAngle) + 0.5 >> 0);
            oCtx.moveTo(nXCenter, nYCenter);
            oCtx.lineTo(nX1, nY1);
            oCtx.arc(nXCenter, nYCenter, nRadius, fStartAngle, fStartAngle - fAngle, true);
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createSlideFromTop = function(fTime, nTransition) {
        var oTexture = this.createTexture();
        var nX = 0;
        var nY = -(this.canvas.height * (1 - fTime) + 0.5 >> 0);
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.drawImage(this.canvas, nX, nY);
        return oTexture;
    };
    CAnimTexture.prototype.createSlideFromBottom = function(fTime, nTransition) {
        var oTexture = this.createTexture();
        var nX = 0;
        var nY = (this.canvas.height * (1 - fTime) + 0.5) >> 0;
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.drawImage(this.canvas, nX, nY);
        return oTexture;
    };
    CAnimTexture.prototype.createSlideFromLeft = function(fTime, nTransition) {
        var oTexture = this.createTexture();
        var nX = -(this.canvas.width * (1 - fTime) + 0.5 >> 0);
        var nY = 0;
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.drawImage(this.canvas, nX, nY);
        return oTexture;
    };
    CAnimTexture.prototype.createSlideFromRight = function(fTime, nTransition) {
        var oTexture = this.createTexture();
        var nX = (this.canvas.width * (1 - fTime) + 0.5) >> 0;
        var nY = 0;
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.drawImage(this.canvas, nX, nY);
        return oTexture;
    };
    CAnimTexture.prototype.createPlusOut = function(fTime, nTransition) {
        var nRectWidth = this.canvas.width * (1 - fTime) / 2 + 0.5 >> 0;
        var nRectHeight = this.canvas.height * (1 - fTime) / 2 + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var sOperation;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperation = 'destination-in';
        }
        else {
            sOperation = 'destination-out';
        }
        oCtx.globalCompositeOperation = sOperation;
        oCtx.beginPath();
        oCtx.moveTo(nRectWidth, 0);
        oCtx.lineTo(this.canvas.width - nRectWidth, 0);
        oCtx.lineTo(this.canvas.width - nRectWidth, nRectHeight);
        oCtx.lineTo(this.canvas.width, nRectHeight);
        oCtx.lineTo(this.canvas.width, this.canvas.height - nRectHeight);
        oCtx.lineTo(this.canvas.width - nRectWidth, this.canvas.height - nRectHeight);
        oCtx.lineTo(this.canvas.width - nRectWidth, this.canvas.height);
        oCtx.lineTo(nRectWidth, this.canvas.height);
        oCtx.lineTo(nRectWidth, this.canvas.height - nRectHeight);
        oCtx.lineTo(0, this.canvas.height - nRectHeight);
        oCtx.lineTo(0, nRectHeight);
        oCtx.lineTo(nRectWidth, nRectHeight);
        oCtx.closePath();
        oCtx.fill();

        //this.drawRect(oCtx, 0, 0, nRectWidth, nRectHeight);
        //this.drawRect(oCtx, this.canvas.width - nRectWidth, 0, nRectWidth, nRectHeight);
        //this.drawRect(oCtx, 0, this.canvas.height - nRectHeight, nRectWidth, nRectHeight);
        //this.drawRect(oCtx, this.canvas.width - nRectWidth, this.canvas.height - nRectHeight, nRectWidth, nRectHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createPlusIn = function(fTime, nTransition) {
        var nRectWidth = this.canvas.width * (fTime) / 2 + 0.5 >> 0;
        var nRectHeight = this.canvas.height * (fTime) / 2 + 0.5 >> 0;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var sOperation;
        if(nTransition === TRANSITION_TYPE_IN) {
            sOperation = 'destination-out';
        }
        else {
            sOperation = 'destination-in';
        }
        oCtx.globalCompositeOperation = sOperation;
        var nPlusWidth = this.canvas.width - 2*nRectWidth;
        var nPlusHeight = this.canvas.height - 2*nRectHeight;
        oCtx.beginPath();
        oCtx.moveTo(nRectWidth, 0);
        oCtx.lineTo(this.canvas.width - nRectWidth, 0);
        oCtx.lineTo(this.canvas.width - nRectWidth, nRectHeight);
        oCtx.lineTo(this.canvas.width, nRectHeight);
        oCtx.lineTo(this.canvas.width, this.canvas.height - nRectHeight);
        oCtx.lineTo(this.canvas.width- nRectWidth, this.canvas.height - nRectHeight);
        oCtx.lineTo(this.canvas.width - nRectWidth, this.canvas.height);
        oCtx.lineTo(nRectWidth, this.canvas.height);
        oCtx.lineTo(nRectWidth, this.canvas.height - nRectHeight);
        oCtx.lineTo(0, this.canvas.height - nRectHeight);
        oCtx.lineTo(0, nRectHeight);
        oCtx.lineTo(nRectWidth, nRectHeight);
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createWipeLeft = function(fTime, nTransition) {
        var nWidth = this.canvas.width * (fTime) + 0.5 >> 0;
        var nHeight = this.canvas.height;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-in';

        }
        else {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        this.drawRect(oCtx, 0, 0, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createWipeRight = function(fTime, nTransition) {
        var nWidth = this.canvas.width * (fTime) + 0.5 >> 0;
        var nHeight = this.canvas.height;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-in';

        }
        else {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        this.drawRect(oCtx, this.canvas.width - nWidth, 0, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createWipeDown = function(fTime, nTransition) {
        var nHeight = this.canvas.height * (fTime) + 0.5 >> 0;
        var nWidth = this.canvas.width;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-in';

        }
        else {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        this.drawRect(oCtx, 0, this.canvas.height - nHeight, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createWipeUp = function(fTime, nTransition) {
        var nHeight = this.canvas.height * fTime + 0.5 >> 0;
        var nWidth = this.canvas.width;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-in';
        }
        else {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        this.drawRect(oCtx, 0, 0, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createBarnOutVertical = function(fTime, nTransition) {
        var nWidth = (this.canvas.width * (fTime))  + 0.5 >> 0;
        var nHeight = this.canvas.height;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-in';
        }
        else {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        this.drawRect(oCtx, (this.canvas.width - nWidth) / 2 + 0.5 >> 0, 0, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createBarnInVertical = function(fTime, nTransition) {
        var nWidth = (this.canvas.width * (1 - fTime))  + 0.5 >> 0;
        var nHeight = this.canvas.height;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        else {
            oCtx.globalCompositeOperation = 'destination-in';
        }
        this.drawRect(oCtx, (this.canvas.width - nWidth) / 2 + 0.5 >> 0, 0, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createBarnOutHorizontal = function(fTime, nTransition) {
        var nHeight = (this.canvas.height * (fTime))  + 0.5 >> 0;
        var nWidth = this.canvas.width;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-in';
        }
        else {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        this.drawRect(oCtx, 0, (this.canvas.height - nHeight) / 2 + 0.5 >> 0, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createBarnInHorizontal = function(fTime, nTransition) {
        var nHeight = (this.canvas.height * (1 - fTime))  + 0.5 >> 0;
        var nWidth = this.canvas.width;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        if(nTransition === TRANSITION_TYPE_IN) {
            oCtx.globalCompositeOperation = 'destination-out';
        }
        else {
            oCtx.globalCompositeOperation = 'destination-in';
        }
        this.drawRect(oCtx, 0, (this.canvas.height - nHeight) / 2 + 0.5 >> 0, nWidth, nHeight);
        return oTexture;
    };
    CAnimTexture.prototype.createDissolve = function(fTime, nTransition) {
        var fResultTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResultTime = 1 - fTime;
        }
        else {
            fResultTime = fTime;
        }
        var nFilledBars = RANDOM_BARS_ARRAY.length * fResultTime + 0.5 >> 0;
        if(nFilledBars === 0) {
            return this;
        }
        var nWidth = this.canvas.width;
        var oTexture = this.createCopy();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        oCtx.globalCompositeOperation = 'destination-out';
        var aFilledPix = RANDOM_BARS_ARRAY.slice(0, nFilledBars);
        oCtx.beginPath();
        for(var nPix = 0; nPix < aFilledPix.length; ++nPix) {
            var nPixNum = aFilledPix[nPix];
            var nX = nPixNum / 10 >> 0;
            var nY = nPixNum % 10;
            while (nX < this.canvas.width) {
                nY = nPix % 10;
                while (nY < this.canvas.height) {
                    oCtx.fillRect(nX - 1, nY - 1, 2, 2);
                    nY += 10;
                }
                nX += 10;
            }
        }
        oCtx.closePath();
        oCtx.fill();
        return oTexture;
    };
    CAnimTexture.prototype.createFade = function(fTime, nTransition) {
        var oTexture = this.createTexture();
        var oCanvas = oTexture.canvas;
        var oCtx = oCanvas.getContext('2d');
        var fResTime;
        if(nTransition === TRANSITION_TYPE_IN) {
            fResTime = fTime;
        }
        else {
            fResTime = 1 - fTime;
        }
        oCtx.globalAlpha = fResTime;
        oCtx.drawImage(this.canvas, 0, 0);
        oCtx.globalAlpha = 1;
        return oTexture;
    };


    function CTexturesCache(oDrawer) {
        this.drawer = oDrawer;
        this.map = {};
    }
    CTexturesCache.prototype.checkTexture = function(sId, fScale) {
        if(!this.map[sId] || !this.map[sId].checkScale(fScale)) {
            this.map[sId] = this.createDrawingTexture(sId, fScale);
            if(!this.map[sId]) {
                this.removeTexture(sId);
                return undefined;
            }
        }
        return this.map[sId];
    };
    CTexturesCache.prototype.createDrawingTexture = function(sId, fScale) {
        var oDrawing = AscCommon.g_oTableId.Get_ById(sId);
        if(!oDrawing) {
            return undefined;
        }
        var oBaseTexture = oDrawing.getAnimTexture(fScale);
        return new CAnimTexture(this, oBaseTexture.canvas, oBaseTexture.scale, oBaseTexture.x, oBaseTexture.y);
    };
    CTexturesCache.prototype.removeTexture = function(sId) {
        if(this.map[sId]) {
            delete this.map[sId];
        }
    };
    CTexturesCache.prototype.clear = function() {
        this.map = {};
    };

    function CAnimationDrawer(player) {
        this.player = player;
        this.sandwiches = {};//map by drawing id
        this.lastFrameSandwiches = {};
        this.texturesCache = new CTexturesCache(this);
        this.hiddenObjects = {};
        this.showObjects = {};
        this.collectHiddenObjects();
    }
    CAnimationDrawer.prototype.clearSandwiches = function() {
        this.sandwiches = {};
    };
    CAnimationDrawer.prototype.clearLastFrameSandwiches = function() {
        this.lastFrameSandwiches = {};
    };
    CAnimationDrawer.prototype.clearTextureCache = function() {
        this.texturesCache.clear();
    };
    CAnimationDrawer.prototype.stop = function() {
        this.clearSandwiches();
        this.clearLastFrameSandwiches();
        this.clearTextureCache();
    };
    CAnimationDrawer.prototype.addAnimationToDraw = function(sDrawingId, oAnimation) {
        if(!this.sandwiches[sDrawingId]) {
            this.sandwiches[sDrawingId] = new CAnimSandwich(sDrawingId, this.player.getElapsedTicks());
        }
        var oSandwich = this.sandwiches[sDrawingId];
        oSandwich.addAnimation(oAnimation);
    };
    CAnimationDrawer.prototype.onFrame = function() {
        this.lastFrameSandwiches = this.sandwiches;
        this.collectSandwiches();
        if(this.checkNeedRedrawFrame()) {
            this.onRecalculateFrame();
        }
    };
    CAnimationDrawer.prototype.checkNeedRedrawFrame = function() {
        var sDrawingId;
        var oCurSandwich, oOldSandwich;
        for(sDrawingId in this.sandwiches) {
            if(!this.lastFrameSandwiches[sDrawingId]) {
                return true;
            }
            oCurSandwich = this.sandwiches[sDrawingId];
            oOldSandwich = this.lastFrameSandwiches[sDrawingId];
            if(!oCurSandwich.isEqualResultAttributes(oOldSandwich)) {
                return true;
            }
            //compare
        }
        for(sDrawingId in this.lastFrameSandwiches) {
            if(!this.sandwiches[sDrawingId]) {
                return true;
            }
        }
        return false;
    };
    CAnimationDrawer.prototype.drawFrame = function(oCanvas, oRect) {
        if(!oCanvas) {
            return;
        }
        var oSlide = this.getSlide();
        if(!oSlide) {
            return;
        }
        var oGraphics = this.createGraphics(oCanvas, oRect);
        oGraphics.m_oContext.clearRect(0, 0, oRect.width, oRect.height);
        var bClip = false;
        if(oRect.x !== 0 || oRect.y !== 0 ||
            oRect.width !== oCanvas.width || oRect.height !== oCanvas.height) {
            oGraphics.SaveGrState();
            oGraphics.AddClipRect(0, 0, this.getSlideWidth(), this.getSlideHeight());
            bClip = true;
        }
        oGraphics.animationDrawer = this;
        oSlide.draw(oGraphics);
        if(bClip) {
            oGraphics.RestoreGrState();
        }
    };
    CAnimationDrawer.prototype.drawObject = function(oDrawing, oGraphics) {
        var sDrawingId = oDrawing.Get_Id();
        var oSandwich = this.getSandwich(sDrawingId);
        var oAttributes = oSandwich && oSandwich.getAttributesMap()
        var fScale = oGraphics.m_oCoordTransform.sx;
        if(!this.isDrawingHidden(sDrawingId) || (oAttributes && oAttributes["style.visibility"] === "visible")) {
            if(!oSandwich) {
                var oTexture = this.texturesCache.checkTexture(sDrawingId, fScale);
                oTexture.draw(oGraphics);
            }
            else {
                oSandwich.drawObject(oGraphics, oDrawing, this.texturesCache, oAttributes);
            }
        }
    };
    CAnimationDrawer.prototype.createGraphics = function(oCanvas, oRect) {
        var wPix = oRect.w;
        var hPix = oRect.h;
        var wMM = this.getSlideWidth();
        var hMM = this.getSlideHeight();
        var oGraphics = new AscCommon.CGraphics();
        var oCtx = oCanvas.getContext('2d');
        oGraphics.init(oCtx, wPix, hPix, wMM, hMM);
        oGraphics.m_oCoordTransform.tx = oRect.x;
        oGraphics.m_oCoordTransform.ty = oRect.y;
        oGraphics.m_oFontManager = AscCommon.g_fontManager;
        oGraphics.transform(1,0,0,1,0,0);
        oGraphics.IsNoDrawingEmptyPlaceholder = true;
        oGraphics.IsDemonstrationMode = true;
        return oGraphics;
    };
    CAnimationDrawer.prototype.onRecalculateFrame = function() {
        this.player.onRecalculateFrame();
    };
    CAnimationDrawer.prototype.collectSandwiches = function() {
        this.clearSandwiches();
        for(var nTiming = 0; nTiming < this.player.timings.length; ++nTiming) {
            var oRoot = this.player.timings[nTiming].getTimingRootNode();
            if(oRoot) {
                oRoot.traverseDrawable(this.player);
            }
        }
    };
    CAnimationDrawer.prototype.getSlideWidth = function() {
        return this.player.getSlideWidth();
    };
    CAnimationDrawer.prototype.getSlideHeight = function() {
        return this.player.getSlideHeight();
    };
    CAnimationDrawer.prototype.getSlide = function() {
        return this.player.slide;
    };
    CAnimationDrawer.prototype.getSandwich = function(sId) {
        var oSandwich = this.sandwiches[sId];
        if(!oSandwich) {
            return null;
        }
        return oSandwich;
    };
    CAnimationDrawer.prototype.isDrawingHidden = function(sId, oSandwich) {
        var oAnim = this.hiddenObjects[sId];
        if(!oAnim) {
            return false;
        }
        if(!oAnim.isDrawable()) {

            return true;
        }
        return false;
    };
    CAnimationDrawer.prototype.checkShowObject = function(oTimeNode) {
        if(oTimeNode.doesShowObject()) {
            var sId = oTimeNode.getTargetObjectId();
            if(sId !== null) {
                this.showObjects[sId] = oTimeNode;
            }
        }
    };
    CAnimationDrawer.prototype.checkHiddenObject = function(oTimeNode) {
        if(oTimeNode.doesHideObject()) {
            var sId = oTimeNode.getTargetObjectId();
            if(sId !== null && !this.showObjects[sId]) {
                this.hiddenObjects[oTimeNode.getTargetObjectId()] = oTimeNode;
            }
        }
    };
    CAnimationDrawer.prototype.collectHiddenObjects = function() {
        var aTimings = this.player.timings;
        var oThis = this;
        this.showObjects = {};
        this.hiddenObjects = {};
        for(var nTiming = 0; nTiming < aTimings.length; ++nTiming) {
            var oRoot = aTimings[nTiming].getTimingRootNode();
            if(oRoot) {
                oRoot.traverseTimeNodes(function(oTimeNode) {
                    oThis.checkShowObject(oTimeNode);
                    oThis.checkHiddenObject(oTimeNode);
                });
            }
        }
        this.showObjects = {};
    };
    CAnimationDrawer.prototype.clearObjectTexture = function(sId) {
        this.texturesCache.removeTexture(sId);
    };

    function CAnimationPlayer(oSlide, drawer) {
        this.slide = oSlide;
        this.timings = [];
        this.updateTimingList();
        this.eventsProcessor = new CEventsProcessor(this);
        this.animationScheduler = new CAnimationScheduler(this);
        this.animationDrawer = new CAnimationDrawer(this);
        this.timer = new CAnimationTimer(this);
        this.drawer = drawer;
    }
    CAnimationPlayer.prototype.updateTimingList = function() {
        this.timings.length = 0;
        if(this.slide.timing) {
            this.timings.push(this.slide.timing);
        }
        if(this.slide.Layout.timing) {
            this.timings.push(this.slide.Layout.timing);
        }
        if(this.slide.Layout.Master.timing) {
            this.timings.push(this.slide.Layout.Master.timing);
        }
        this.resetNodesState();
    };
    CAnimationPlayer.prototype.getPresentation = function() {
        return editor.WordControl.m_oLogicDocument;
    };
    CAnimationPlayer.prototype.getSlideWidth = function() {
        return this.getPresentation().GetWidthMM();
    };
    CAnimationPlayer.prototype.getSlideHeight = function() {
        return this.getPresentation().GetHeightMM();
    };
    CAnimationPlayer.prototype.start = function() {
        if(this.isStarted()) {
            if(this.isMainSequenceFinished()) {
                this.onMainSeqFinished();
            }
            return;
        }
        var bIsPaused = this.isPaused();
        this.timer.start();
        if(!bIsPaused) {
            this.updateTimingList();
            this.scheduleNodesStart();
            this.animationDrawer.clearTextureCache();
            this.animationDrawer.collectHiddenObjects();
        }
        if(this.isMainSequenceFinished()) {
            this.onMainSeqFinished();
        }
    };
    CAnimationPlayer.prototype.resetNodesState = function() {
        for(var nTiming = 0; nTiming < this.timings.length; ++nTiming) {
            var oRoot = this.timings[nTiming].getTimingRootNode();
            if(oRoot) {
                oRoot.resetState();
            }
        }
    };
    CAnimationPlayer.prototype.scheduleNodesStart = function() {
        for(var nTiming = 0; nTiming < this.timings.length; ++nTiming) {
            var oRoot = this.timings[nTiming].getTimingRootNode();
            if(oRoot) {
                oRoot.scheduleStart(this);
            }
        }
    };
    CAnimationPlayer.prototype.stop = function() {
        this.timer.stop();
        this.animationScheduler.stop();
        this.animationDrawer.stop();
        this.resetNodesState();
    };
    CAnimationPlayer.prototype.onMainSeqFinished = function () {
        if(this.drawer) {
            var nSlideNum = -1;
            if(this.slide) {
                nSlideNum = this.slide.num;
            }
            var oThis = this;
            setTimeout(function() {
                oThis.drawer.OnAnimMainSeqFinished(nSlideNum);
            }, 1);
        }
    };
    CAnimationPlayer.prototype.isMainSequenceFinished = function () {
        var oTiming = this.timings[0];
        if(oTiming) {
            return oTiming.isMainSequenceAtEnd();
        }
        return true;
    };
    CAnimationPlayer.prototype.pause = function() {
        this.timer.pause();
        this.animationDrawer.clearTextureCache();
    };
    CAnimationPlayer.prototype.onFrame = function() {
        if(!this.isStarted()) {
            return;
        }
        //console.log("------------------TICK START----------------------------------------");
        this.timer.onFrame();
        this.animationScheduler.onFrame();
        this.animationDrawer.onFrame();
        this.eventsProcessor.onFrame();
        //console.log("------------------TICK END-------------------------------------------");
    };
    CAnimationPlayer.prototype.getElapsedTicks = function() {
        return this.timer.getElapsedTicks();
    };
    CAnimationPlayer.prototype.getElapsedTime = function() {
        return this.timer.getElapsedTime();
    };
    CAnimationPlayer.prototype.isStarted = function() {
        return this.timer.isStarted();
    };
    CAnimationPlayer.prototype.isPaused = function() {
        return this.timer.isPaused();
    };
    CAnimationPlayer.prototype.isStopped = function() {
        return this.timer.isStopped();
    };
    CAnimationPlayer.prototype.onClicked = function(sId, nTime) {
    };
    CAnimationPlayer.prototype.scheduleEvent = function(oEvent) {
        if(!this.isStarted()) {
            return;
        }
        this.animationScheduler.addEvent(oEvent);
    };
    CAnimationPlayer.prototype.cancelCallerEvent = function(oCaller) {
        this.animationScheduler.cancelCallerEvents(oCaller);
    };
    CAnimationPlayer.prototype.addExternalEvent = function(oExternalEvent) {
        if(!this.isStarted()) {
            return false;
        }
        return this.eventsProcessor.addEvent(oExternalEvent);
    };
    CAnimationPlayer.prototype.onClick = function() {
        var bClick = this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_CLICK, null));
        if(bClick) {
            return true;
        }
        return this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_NEXT, null));
    };
    CAnimationPlayer.prototype.onSpClick = function(oSp) {
        if(!oSp) {
            return false;
        }
        var bClick = this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_CLICK, oSp.Get_Id()));
        if(bClick) {
            return true;
        }
        var sMediaName = oSp.getMediaFileName();
        if(sMediaName) {
            if (window["AscDesktopEditor"])
                return false;
        }
        return this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_NEXT, null));
    };
    CAnimationPlayer.prototype.isSpClickTrigger = function(oSp) {
        for(var nTiming = 0; nTiming < this.timings.length; ++nTiming) {
            if(this.timings[nTiming].isSpClickTrigger(oSp)) {
                return true;
            }
        }
        return false;
    };
    CAnimationPlayer.prototype.onSpDblClick = function(oSp) {
        if(!oSp) {
            return false;
        }
        return this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_DBLCLICK, oSp.Get_Id()));
    };
    CAnimationPlayer.prototype.onSpMouseOver = function(oSp) {
        if(!oSp) {
            return false;
        }
        return this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_MOUSEOVER, oSp.Get_Id()));
    };
    CAnimationPlayer.prototype.onSpMouseOut = function(oSp) {
        if(!oSp) {
            return false;
        }
        return this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_MOUSEOUT, oSp.Get_Id()));
    };
    CAnimationPlayer.prototype.onNextSlide = function() {
        var bNext = this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_NEXT, null));
        if(bNext) {
            return true;
        }
        this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_CLICK, null));
    };
    CAnimationPlayer.prototype.onPrevSlide = function() {
        return this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_PREV, null));
    };
    CAnimationPlayer.prototype.addAnimationToDraw = function(sDrawingId, oAnimation) {
        this.animationDrawer.addAnimationToDraw(sDrawingId, oAnimation);
    };
    CAnimationPlayer.prototype.drawFrame = function(oCanvas, oRect) {
        this.animationDrawer.drawFrame(oCanvas, oRect);
    };
    CAnimationPlayer.prototype.onRecalculateFrame = function() {
        if(this.drawer) {
            this.drawer.OnRecalculateAnimationFrame(this);
        }
    };
    CAnimationPlayer.prototype.getExternalEvent = function() {
      return this.eventsProcessor.getExternalEvent();
    };
    CAnimationPlayer.prototype.clearObjectTexture = function(sId) {
        this.animationDrawer.clearObjectTexture(sId);
    };
    CAnimationPlayer.prototype.isDrawingHidden = function(sId) {
        return this.animationDrawer.isDrawingHidden(sId);
    };


    function CDemoAnimPlayer(oSlide) {
        CAnimationPlayer.call(this, oSlide, null);
    }
    InitClass(CDemoAnimPlayer, CAnimationPlayer, 0);
    CDemoAnimPlayer.prototype.updateTimingList = function() {
        this.timings.length = 0;
        var oTiming = this.slide.timing;
        if(oTiming) {
            var oDemoTiming = oTiming.createDemoTiming();
            if(oDemoTiming) {

            }
            this.timings.push(oDemoTiming);
        }
        var oTr      = editor.WordControl.m_oDrawingDocument.TransitionSlide;
        oTr.CalculateRect();
        var oR = oTr.Rect;
        this.rect = new AscFormat.CGraphicBounds(oR.x, oR.y, oR.x + oR.w, oR.y + oR.h);
        this.overlayCanvas = editor.WordControl.m_oOverlayApi.m_oContext.canvas;
        this.overlay = editor.WordControl.m_oOverlayApi;
    };
    CDemoAnimPlayer.prototype.onMainSeqFinished = function () {
        var oThis = this;
        setTimeout(function () {
            if(!oThis.isStopped()) {
                oThis.stop();
                editor.WordControl.m_oLogicDocument.StopAnimationPreview();
            }
        }, 1000);
    };
    
    CDemoAnimPlayer.prototype.start = function () {
        CAnimationPlayer.prototype.start.call(this);
        this.overlay.CheckRect(this.rect.x, this.rect.y, this.rect.w, this.rect.h);
        this.onRecalculateFrame();
    };
    CDemoAnimPlayer.prototype.stop = function () {
        CAnimationPlayer.prototype.stop.call(this);
        this.overlay.Clear();
        this.overlay.CheckRect(this.rect.x, this.rect.y, this.rect.w, this.rect.h);
        this.slide.showDrawingObjects();
    };
    CDemoAnimPlayer.prototype.onRecalculateFrame = function() {
        this.overlay.Clear();
        this.overlay.CheckRect(this.rect.x, this.rect.y, this.rect.w, this.rect.h);
        this.drawFrame(this.overlayCanvas, this.rect);
    };

    var DEFAULT_SIMPLE_TRIGGER = function() {
        return true;
    };
    var DEFAULT_NEVER_TRIGGER = function() {
        return false;
    };


/* Attributes names
    style.opacity
    style.rotation
    style.visibility
    style.color
    style.fontSize
    style.fontWeight
    style.fontStyle
    style.fontFamily
    style.textEffectEmboss
    style.textShadow
    style.textTransform
    style.textDecorationUnderline
    style.textEffectOutline
    style.textDecorationLineThrough
    style.sRotation
    imageData.cropTop
    imageData.cropBottom
    imageData.cropLeft
    imageData.cropRight
    imageData.cropRight
    imageData.gain
    imageData.blacklevel
    imageData.gamma
    imageData.grayscale
    imageData.chromakey
    fill.on
    fill.type
    fill.color
    fill.opacity
    fill.color2
    fill.method
    fill.opacity2
    fill.angle
    fill.focus
    fill.focusposition.x
    fill.focusposition.y
    fill.focussize.x
    fill.focussize.y
    stroke.on
    stroke.color
    stroke.weight
    stroke.opacity
    stroke.linestyle
    stroke.dashstyle
    stroke.filltype
    stroke.src
    stroke.color2
    stroke.imagesize.x
    stroke.imagesize.y
    stroke.startArrow
    stroke.endArrow
    stroke.startArrowWidth
    stroke.startArrowLength
    stroke.endArrowWidth
    stroke.endArrowLength
    shadow.on
    shadow.type
    shadow.color
    shadow.color2
    shadow.opacity
    shadow.offset.x
    shadow.offset.y
    shadow.offset2.x
    shadow.offset2.y
    shadow.origin.x
    shadow.origin.y
    shadow.matrix.xtox
    shadow.matrix.ytox
    shadow.matrix.xtox
    shadow.matrix.ytoy
    shadow.matrix.perspectiveX
    shadow.matrix.perspectiveY
    skew.on
    skew.offset.x
    skew.offset.y
    skew.origin.x
    skew.origin.y
    skew.matrix.xtox
    skew.matrix.ytox
    skew.matrix.xtox
    skew.matrix.ytoy
    skew.matrix.perspectiveX
    skew.matrix.perspectiveY
    extrusion.on
    extrusion.type
    extrusion.render
    extrusion.viewpointorigin.x
    extrusion.viewpointorigin.y,
    extrusion.viewpoint.x
    extrusion.viewpoint.y
    extrusion.viewpoint.z
    extrusion.plane
    extrusion.skewangle
    extrusion.skewamt
    extrusion.backdepth
    extrusion.foredepth
    extrusion.orientation.x
    extrusion.orientation.y
    extrusion.orientation.z
    extrusion.orientationangle
    extrusion.color
    extrusion.rotationangle.x
    extrusion.rotationangle.y
    extrusion.lockrotationcenter
    extrusion.autorotationcenter
    extrusion.rotationcenter.x
    extrusion.rotationcenter.y
    extrusion.rotationcenter.z
    extrusion.colormode
    ppt_x
    ppt_y
    ppt_w
    ppt_h
    ppt_c
    ppt_r
    xshear
    yshear
    image
    ScaleX
    ScaleY
    r
    fillcolor
    3d.object.rotation.x
    3d.object.rotation.y
    3d.object.rotation.z
    3d.view.rotation.x
    3d.view.rotation.y
    3d.view.rotation.z
    3d.object.scale.x
    3d.object.scale.y
    3d.object.scale.z
    3d.view.scale.x
    3d.view.scale.y
    3d.view.scale.z
    3d.object.translation.x
    3d.object.translation.y
    3d.object.translation.z
    3d.view.translation.x
    3d.view.translation.y
    3d.view.translation.z
    drawProgress
    drawProgressAllAtOnce
    */

    var FILTER_TYPE_BLINDS_HORIZONTAL = 0;
    var FILTER_TYPE_BLINDS_VERTICAL = 1;
    var FILTER_TYPE_BOX_IN = 2;
    var FILTER_TYPE_BOX_OUT = 3;
    var FILTER_TYPE_CHECKERBOARD_ACROSS = 4;
    var FILTER_TYPE_CHECKERBOARD_DOWN = 5;
    var FILTER_TYPE_CIRCLE = 6;
    var FILTER_TYPE_CIRCLE_IN = 7;
    var FILTER_TYPE_CIRCLE_OUT = 8;
    var FILTER_TYPE_DIAMOND = 9;
    var FILTER_TYPE_DIAMOND_IN = 10;
    var FILTER_TYPE_DIAMOND_OUT = 11;
    var FILTER_TYPE_DISSOLVE = 12;
    var FILTER_TYPE_FADE = 13;
    var FILTER_TYPE_SLIDE_FROM_TOP = 14;
    var FILTER_TYPE_SLIDE_FROM_BOTTOM = 15;
    var FILTER_TYPE_SLIDE_FROM_LEFT = 16;
    var FILTER_TYPE_SLIDE_FROM_RIGHT = 17;
    var FILTER_TYPE_PLUS_IN = 18;
    var FILTER_TYPE_PLUS_OUT = 19;
    var FILTER_TYPE_BARN_IN_VERTICAL = 20;
    var FILTER_TYPE_BARN_IN_HORIZONTAL = 21;
    var FILTER_TYPE_BARN_OUT_VERTICAL = 22;
    var FILTER_TYPE_BARN_OUT_HORIZONTAL = 23;
    var FILTER_TYPE_RANDOM_BARS_HORIZONTAL = 24;
    var FILTER_TYPE_RANDOM_BARS_VERTICAL = 25;
    var FILTER_TYPE_STRIPS_DOWN_LEFT = 26;
    var FILTER_TYPE_STRIPS_UP_LEFT = 27;
    var FILTER_TYPE_STRIPS_DOWN_RIGHT = 28;
    var FILTER_TYPE_STRIPS_UP_RIGHT = 29;
    var FILTER_TYPE_SLIDE_WEDGE = 30;
    var FILTER_TYPE_WHEEL_1 = 31;
    var FILTER_TYPE_WHEEL_2 = 32;
    var FILTER_TYPE_WHEEL_3 = 33;
    var FILTER_TYPE_WHEEL_4 = 34;
    var FILTER_TYPE_WHEEL_8 = 35;
    var FILTER_TYPE_WIPE_RIGHT = 36;
    var FILTER_TYPE_WIPE_LEFT = 37;
    var FILTER_TYPE_WIPE_DOWN = 38;
    var FILTER_TYPE_WIPE_UP = 39;

    var FILTER_MAP = {};
    FILTER_MAP["blinds(horizontal)"] = FILTER_TYPE_BLINDS_HORIZONTAL;
    FILTER_MAP["blinds(vertical)"] = FILTER_TYPE_BLINDS_VERTICAL;
    FILTER_MAP["box(in)"] = FILTER_TYPE_BOX_IN;
    FILTER_MAP["box(out)"] = FILTER_TYPE_BOX_OUT;
    FILTER_MAP["checkerboard(across)"] = FILTER_TYPE_CHECKERBOARD_ACROSS;
    FILTER_MAP["checkerboard(down)"] = FILTER_TYPE_CHECKERBOARD_DOWN;
    FILTER_MAP["circle"] = FILTER_TYPE_CIRCLE;
    FILTER_MAP["circle(in)"] = FILTER_TYPE_CIRCLE_IN;
    FILTER_MAP["circle(out)"] = FILTER_TYPE_CIRCLE_OUT;
    FILTER_MAP["diamond"] = FILTER_TYPE_DIAMOND;
    FILTER_MAP["diamond(in)"] = FILTER_TYPE_DIAMOND_IN;
    FILTER_MAP["diamond(out)"] = FILTER_TYPE_DIAMOND_OUT;
    FILTER_MAP["dissolve"] = FILTER_TYPE_DISSOLVE;
    FILTER_MAP["fade"] = FILTER_TYPE_FADE;
    FILTER_MAP["slide(fromTop)"] = FILTER_TYPE_SLIDE_FROM_TOP;
    FILTER_MAP["slide(fromBottom)"] = FILTER_TYPE_SLIDE_FROM_BOTTOM;
    FILTER_MAP["slide(fromLeft)"] = FILTER_TYPE_SLIDE_FROM_LEFT;
    FILTER_MAP["slide(fromRight)"] = FILTER_TYPE_SLIDE_FROM_RIGHT;
    FILTER_MAP["plus(in)"] = FILTER_TYPE_PLUS_IN;
    FILTER_MAP["plus(out)"] = FILTER_TYPE_PLUS_OUT;
    FILTER_MAP["barn(inVertical)"] = FILTER_TYPE_BARN_IN_VERTICAL;
    FILTER_MAP["barn(inHorizontal)"] = FILTER_TYPE_BARN_IN_HORIZONTAL;
    FILTER_MAP["barn(outVertical)"] = FILTER_TYPE_BARN_OUT_VERTICAL;
    FILTER_MAP["barn(outHorizontal)"] = FILTER_TYPE_BARN_OUT_HORIZONTAL;
    FILTER_MAP["randomBars(horizontal)"] = FILTER_TYPE_RANDOM_BARS_HORIZONTAL;
    FILTER_MAP["randombar(horizontal)"] = FILTER_TYPE_RANDOM_BARS_HORIZONTAL;
    FILTER_MAP["randomBars(vertical)"] = FILTER_TYPE_RANDOM_BARS_VERTICAL;
    FILTER_MAP["randombar(vertical)"] = FILTER_TYPE_RANDOM_BARS_VERTICAL;
    FILTER_MAP["strips(downLeft)"] = FILTER_TYPE_STRIPS_DOWN_LEFT;
    FILTER_MAP["strips(upLeft)"] = FILTER_TYPE_STRIPS_UP_LEFT;
    FILTER_MAP["strips(downRight)"] = FILTER_TYPE_STRIPS_DOWN_RIGHT;
    FILTER_MAP["strips(upRight)"] = FILTER_TYPE_STRIPS_UP_RIGHT;
    FILTER_MAP["wedge"] = FILTER_TYPE_SLIDE_WEDGE;
    FILTER_MAP["wheel(1)"] = FILTER_TYPE_WHEEL_1;
    FILTER_MAP["wheel(2)"] = FILTER_TYPE_WHEEL_2;
    FILTER_MAP["wheel(3)"] = FILTER_TYPE_WHEEL_3;
    FILTER_MAP["wheel(4)"] = FILTER_TYPE_WHEEL_4;
    FILTER_MAP["wheel(8)"] = FILTER_TYPE_WHEEL_8;
    FILTER_MAP["wipe(right)"] = FILTER_TYPE_WIPE_RIGHT;
    FILTER_MAP["wipe(left)"] = FILTER_TYPE_WIPE_LEFT;
    FILTER_MAP["wipe(down)"] = FILTER_TYPE_WIPE_DOWN;
    FILTER_MAP["wipe(up)"] = FILTER_TYPE_WIPE_UP;

    function CAnimSandwich(sDrawingId, nElapsedTime) {
        this.drawingId = sDrawingId;
        this.elapsedTime = nElapsedTime;
        this.animations = [];
        this.cachedAttributes = null;
    }
    CAnimSandwich.prototype.getDrawingId = function() {
        return this.drawingId;
    };
    CAnimSandwich.prototype.addAnimation = function(oAnimation) {
        this.animations.push(oAnimation);
        if(this.cachedAttributes) {
            this.cachedAttributes = null;
        }
        this.checkOnAdd();
    };
    CAnimSandwich.prototype.checkOnAdd = function() {
    };
    CAnimSandwich.prototype.getDrawing = function() {
        return AscCommon.g_oTableId.Get_ById(this.drawingId);
    };
    CAnimSandwich.prototype.checkRemoveOldAnim = function() {
        var oEntrEffect = null, oExitEffect = null;
        for(var nAnim = 0; nAnim < this.animations.length; ++nAnim) {
            var oAnim = this.animations[nAnim];
            var oEffect = oAnim.getParentTimeNode();
            if(oEffect.isAnimEffect()) {
                var oAttrObject = oEffect.getAttributesObject();
                if(oAttrObject && AscFormat.PRESET_CLASS_EXIT === oAttrObject.presetClass) {
                    oExitEffect = oEffect;
                }
                if(oAttrObject && AscFormat.PRESET_CLASS_ENTR === oAttrObject.presetClass) {
                    oEntrEffect = oEffect;
                }
                if(oEntrEffect && oExitEffect) {
                    break;
                }
            }
        }
        var oEffectToDelete = null;
        if(oEntrEffect && oExitEffect) {
            if(oEntrEffect.isAtEnd() && !oExitEffect.isAtEnd()) {
                oEffectToDelete = oEntrEffect;
            }
            if(!oEntrEffect.isAtEnd() && oExitEffect.isAtEnd()) {
                oEffectToDelete = oExitEffect;
            }

            if(oEntrEffect.isAtEnd() && oExitEffect.isAtEnd()) {
                if(oEntrEffect.startTick < oExitEffect.startTick) {
                    oEffectToDelete = oEntrEffect;
                }
                else {
                    oEffectToDelete = oExitEffect;
                }
            }
        }
        if(oEffectToDelete) {
            for(var nAnim = this.animations.length - 1; nAnim > -1; --nAnim) {
                var oAnim = this.animations[nAnim];
                var oEffect = oAnim.getParentTimeNode();
                if(oEffect === oEffectToDelete) {
                    this.animations.splice(nAnim, 1);
                }
            }
            return true;
        }
        return false;
    };
    CAnimSandwich.prototype.getAttributesMap = function() {
        if(this.cachedAttributes) {
            return this.cachedAttributes;
        }

        var bCheckRemove = true;

        while(bCheckRemove) {
            bCheckRemove = this.checkRemoveOldAnim();
        }


        this.animations.sort(function(oAnim1, oAnim2){
            if(AscFormat.isRealNumber(oAnim1.startTick) && AscFormat.isRealNumber(oAnim2.startTick)) {
                return oAnim1.startTick - oAnim2.startTick;
            }
            return 0;
        });
        var oAttributes = {};
        for(var nAnim = 0; nAnim < this.animations.length; ++nAnim) {
            this.animations[nAnim].calculateAttributes(this.elapsedTime, oAttributes);
        }
        this.cachedAttributes = oAttributes;
        return oAttributes;
    };
    CAnimSandwich.prototype.print = function() {
        var oAttributes = this.getAttributesMap();
        //console.log(oAttributes);
    };
    CAnimSandwich.prototype.drawObject = function(oGraphics, oDrawing, oTextureCache, oAttributesMap) {
        //this.print();
        //console.log(oAttributesMap);
        var sVisibility = oAttributesMap["style.visibility"];
        if(sVisibility === "hidden") {
            return;
        }
        var oFillColor = oAttributesMap["fillcolor"] || oAttributesMap["style.color"];
        var sFillType = oAttributesMap["fill.type"];
        var bFillOn = oAttributesMap["fill.on"];
        var fOpacity = oAttributesMap["style.opacity"];

        var oStrokeColor = oAttributesMap["stroke.color"];
        var bStrokeOn = oAttributesMap["stroke.on"];

        var fScale = oGraphics.m_oCoordTransform.sx;
        var sId = oDrawing.Get_Id();
        var oTexture = oTextureCache.checkTexture(sId, fScale);
        if(oFillColor || sFillType || bFillOn !== undefined || oStrokeColor || bStrokeOn !== undefined) {
            var oOldBrush = oDrawing.brush;
            var oOldPen = oDrawing.pen;
            if(bFillOn === false) {
                oDrawing.brush = AscFormat.CreateNoFillUniFill();
            }
            else {
                if(oFillColor) {
                    if(oDrawing.brush && oDrawing.brush.fill && oDrawing.brush.fill.type === AscFormat.FILL_TYPE_SOLID || sFillType === "solid") {
                        oDrawing.brush = AscFormat.CreateUniFillByUniColor(oFillColor);
                    }
                }
            }
            if(bStrokeOn === false) {
                oDrawing.pen = AscFormat.CreateNoFillLine();
            }
            else {
                if(oStrokeColor) {
                    var oPen;
                    if(oDrawing.pen) {
                        oPen = oDrawing.pen.createDuplicate();
                        var oMods;
                        if(oPen.Fill &&
                            oPen.Fill.fill &&
                            oPen.Fill.fill.color &&
                            oPen.Fill.fill.color.Mods &&
                            oPen.Fill.fill.color.Mods.Mods.length !== 0) {
                            oMods = oPen.Fill.fill.color.Mods;
                            oMods.Apply(oStrokeColor.RGBA);
                        }
                    }
                    else {
                        oPen = AscFormat.CreateNoFillLine();
                    }
                    oPen.Fill = AscFormat.CreateUniFillByUniColor(oStrokeColor);
                    oDrawing.pen = oPen;
                }
            }
            oTexture = oTextureCache.createDrawingTexture(sId, fScale);
            oDrawing.brush = oOldBrush;
            oDrawing.pen = oOldPen;
        }
        var oBounds = oDrawing.getBoundsByDrawing();
        var oPresSize = oDrawing.getPresentationSize();
        var fCenterX, fCenterY;
        var bTransform = false;
        fCenterX = oBounds.x + oBounds.w / 2;
        fCenterY = oBounds.y + oBounds.h / 2;
        if(AscFormat.isRealNumber(oAttributesMap["ppt_x"])) {
            fCenterX = oAttributesMap["ppt_x"] * oPresSize.w;
            bTransform = true;
        }
        if(AscFormat.isRealNumber(oAttributesMap["ppt_y"])) {
            fCenterY = oAttributesMap["ppt_y"] * oPresSize.h;
            bTransform = true;
        }
        var fScaleX = 1.0, fScaleY = 1.0;
        if(AscFormat.isRealNumber(oAttributesMap["ScaleX"]) && AscFormat.isRealNumber(oAttributesMap["ScaleY"])) {
            fScaleX = oAttributesMap["ScaleX"];
            fScaleY = oAttributesMap["ScaleY"];
            bTransform = true;
        }
        if(AscFormat.isRealNumber(oAttributesMap["ppt_w"])) {
            var fOrigW = oBounds.w / oPresSize.w;
            fScaleX *= oAttributesMap["ppt_w"]/ fOrigW;
            bTransform = true;
        }
        if(AscFormat.isRealNumber(oAttributesMap["ppt_h"])) {
            var fOrigH = oBounds.h / oPresSize.h;
            fScaleY *= oAttributesMap["ppt_h"]/ fOrigH;
            bTransform = true;
        }
        var fR = 0;
        var fAttrRot = oAttributesMap["ppt_r"] || oAttributesMap["r"] || oAttributesMap["style.rotation"];
        if(AscFormat.isRealNumber(fAttrRot)) {
            if(oAttributesMap["ppt_r"] || oAttributesMap["r"]) {
                fR = AscFormat.cToRad * fAttrRot;
            }
            else if(oAttributesMap["style.rotation"]) {
                fR = Math.PI * fAttrRot / 180;
            }
            bTransform = true;
        }

        var oTransform = null;
        if(bTransform) {
            oTransform = new AscCommon.CMatrix();
            var hc = oBounds.w * 0.5;
            var vc = oBounds.h * 0.5;
            AscCommon.global_MatrixTransformer.TranslateAppend(oTransform, -hc, -vc);
            if(fScaleX !== 1 || fScaleY !== 1) {
                AscCommon.global_MatrixTransformer.ScaleAppend(oTransform, fScaleX, fScaleY);
            }
            if(fR !== 0) {
                AscCommon.global_MatrixTransformer.RotateRadAppend(oTransform, -fR);
            }
            AscCommon.global_MatrixTransformer.TranslateAppend(oTransform, fCenterX, fCenterY);
        }

        oTexture = oTexture.createEffectTexture(oAttributesMap["effect"]);

        if(fOpacity !== undefined) {
           oGraphics.put_GlobalAlpha(true,1 - fOpacity);
        }
        oTexture.draw(oGraphics, oTransform);
        if(fOpacity !== undefined) {
            oGraphics.put_GlobalAlpha(false, 1);
        }
    };
    CAnimSandwich.prototype.isEqualResultAttributes = function(oOtherSandwich) {
        var oAttributes = this.getAttributesMap();
        var oOtherAttributes = oOtherSandwich.getAttributesMap();
        var sKey, val, otherVal;
        for(sKey in oAttributes) {
            val = oAttributes[sKey];
            otherVal = oOtherAttributes[sKey];
            if(otherVal === undefined) {
                return false;
            }
            if(val === otherVal) {
                continue;
            }
            if(sKey === "effect") {
                if(!val.isEqual(otherVal)){
                    return false;
                }
            }
            if(AscFormat.isRealNumber(val)) {
                if(!AscFormat.fApproxEqual(val, otherVal)) {
                    return false;
                }
            }
            else if(typeof val === "string") {
                if(val !== otherVal) {
                    return false;
                }
            }
            else if(val instanceof AscFormat.CUniColor) {
                if(!val.IsIdentical(otherVal)) {
                    return false;
                }
            }
        }
        for(sKey in oOtherAttributes) {
            val = oAttributes[sKey];
            if(val === undefined) {
                return false;
            }
        }
        return true;
    };


    //--------------------------------------------
    //Formula parser


    function CParseQueue(oParser) {
        this.queue = [];
        this.pos = -1;
        this.parser = oParser;
    }
    CParseQueue.prototype.add = function(oToken){
        this.queue.push(oToken);
        this.pos = this.queue.length - 1;
    };
    CParseQueue.prototype.last = function(){
        return this.queue[this.queue.length - 1];
    };
    CParseQueue.prototype.getNext = function(){
        if(this.pos > -1){
            return this.queue[--this.pos];
        }
        return null;
    };
    CParseQueue.prototype.calculate = function(oVarMap){
        this.pos = this.queue.length - 1;
        var oLastToken = this.queue[this.pos];
        if(!oLastToken) {
            return null;
        }
        if(this.checkReplaceVar(oVarMap)) {
            this.replaceVar(oVarMap);
        }
        oLastToken.calculate(oVarMap);
        return oLastToken.result;
    };
    CParseQueue.prototype.checkReplaceVar = function(oVarMap) {
        for(var nToken = 0; nToken < this.queue.length; ++nToken) {
            if(this.queue[nToken].checkReplaceVar(oVarMap)) {
                return true;
            }
        }
        return false;
    };
    CParseQueue.prototype.replaceVar = function(oVarMap) {
        for(var nToken = 0; nToken < this.queue.length; ++nToken) {
            this.queue[nToken].replaceVar(oVarMap);
        }
    };


    function CTokenBase(oQueue) {
        this.queue = oQueue;
        this.result = null;
        this.error = null;
    }
    CTokenBase.prototype.argumentsCount = 0;
    CTokenBase.prototype.precedence = 0;
    CTokenBase.prototype.getPrecedence = function() {
        return this.precedence;
    };
    CTokenBase.prototype.calculate = function(oVarMap) {
        this.result = null;
        this.error = null;
        if(!this.queue) {
            this.error = true;
            return false;
        }
        var aArgs = [];
        var oToken;
        var nArgCount = this.getArgumentsCount();
        for(var nArg = 0; nArg < nArgCount; ++nArg) {
            oToken = this.queue.getNext();
            if(!oToken) {
                this.error = true;
                return;
            }
            var bOk = oToken.calculate(oVarMap);
            if(bOk) {
                aArgs.push(oToken.getResult());
            }
            else {
                return false;
            }
        }
        this._calculate(aArgs, oVarMap);
        this.error = !AscFormat.isRealNumber(this.result);
        return !this.error;
    };
    CTokenBase.prototype._calculate = function(aArgs, oVarMap) {
        this.result = null;
    };
    CTokenBase.prototype.getArgumentsCount = function() {
        return this.argumentsCount;
    };
    CTokenBase.prototype.getResult = function() {
        return this.result;
    };
    CTokenBase.prototype.isFunction = function() {
        return false;
    };
    CTokenBase.prototype.isOperator = function() {
        return false;
    };
    CTokenBase.prototype.checkReplaceVar = function(oVarMap) {
        return false;
    };
    CTokenBase.prototype.replaceVar = function(oVarMap) {
    };

    function CConstantToken(oQueue, sValue) {
        CTokenBase.call(this, oQueue);
        this.value = sValue;
    }
    InitClass(CConstantToken, CTokenBase, undefined);
    CConstantToken.prototype.argumentsCount = 0;
    CConstantToken.prototype.precedence = 9;
    CConstantToken.prototype._calculate = function(aArgs, oVarMap) {
        this.result = this.value;
    };

    function CVariableToken(oQueue, sName) {
        CTokenBase.call(this, oQueue);
        this.name = sName;
    }
    InitClass(CVariableToken, CTokenBase, undefined);
    CVariableToken.prototype.argumentsCount = 0;
    CVariableToken.prototype.precedence = 9;
    CVariableToken.prototype._calculate = function(aArgs, oVarMap) {
        this.result = oVarMap[this.name];
    };
    CVariableToken.prototype.setName = function(sName) {
        this.name = sName;
    };
    CVariableToken.prototype.checkReplaceVar = function(oVarMap) {
        if(oVarMap[this.name + "_no_attr"] && AscFormat.isRealNumber(oVarMap["#" + this.name])) {
            return true;
        }
        return false;
    };
    CVariableToken.prototype.replaceVar = function(oVarMap) {
        if(AscFormat.isRealNumber(oVarMap["#" + this.name])) {
            this.name = "#" + this.name;
        }
    };
    function CFunctionToken(oQueue, sName) {
        CTokenBase.call(this, oQueue);
        this.name = sName;
        this.operands = [];
    }
    InitClass(CFunctionToken, CTokenBase, undefined);
    CFunctionToken.prototype.argumentsCount = 0;
    CFunctionToken.prototype.precedence = 9;
    CFunctionToken.prototype._calculate = function(aArgs, oVarMap) {
        var fFunction = this.functions[this.name];
        if(!fFunction) {
            this.result = null;
            return;
        }
        this.result = fFunction.apply(null, aArgs);
    };
    CFunctionToken.prototype.functions = {
        "abs": function(x) {return Math.abs(x);},
        "acos": function(x) {return Math.acos(x);},
        "asin": function (x) {return Math.asin(x);},
        "atan": function (x) {return Math.atan(x);},
        "ceil": function(x){return Math.ceil(x);},
        "cos": function(x){return Math.cos(x);},
        "cosh": function(x){return Math.cosh(x);},
        "deg": function(x){return x * AscFormat.cToDeg;},
        "exp": function(x){return Math.exp(x);},
        "floor": function(x){return Math.floor(x);},
        "ln": function(x){return Math.log(x);},
        "max": function(x, y){return Math.max(x, y);},
        "min": function(x, y){return Math.min(x, y);},
        "rad": function(x){return x * AscFormat.cToRad;},
        "rand": function(x){return Math.random()* x;},
        "sin": function(x){return Math.sin(x);},
        "sinh": function(x){return Math.sinh(x);},
        "sqrt": function(x){return Math.sqrt(x);},
        "tan": function(x){return Math.tan(x);},
        "tanh": function(x){return Math.tanh(x);}
    };
    CFunctionToken.prototype.getArgumentsCount = function() {
        if(this.name === "max" || this.name === "min") {
            return 2;
        }
        var fFunction = this.functions[this.name];
        if(fFunction) {
            return 1;
        }
        return 0;
    };
    CFunctionToken.prototype.addOperand = function(oOperand) {
        return this.operands.push(oOperand);
    };
    CFunctionToken.prototype.getOperandsCount = function() {
        return this.operands.length;
    };
    CFunctionToken.prototype.isFunction = function() {
        return true;
    };

    function CBinaryOperatorToken(oQueue, sName) {
        CTokenBase.call(this, oQueue);
        this.name = sName;
    }
    InitClass(CBinaryOperatorToken, CTokenBase, undefined);
    CBinaryOperatorToken.prototype.argumentsCount = 2;
    CBinaryOperatorToken.prototype.getPrecedence = function() {
        if(this.name === "^") {
            return 7;
        }
        if(this.name === "*" || this.name === "/" || this.name === "%") {
            return 6;
        }
        if(this.name === "+" || this.name === "-") {
            return 5;
        }
        return 0;
    };
    CBinaryOperatorToken.prototype._calculate = function(aArgs, oVarMap) {
        var fFunction = this.operators[this.name];
        if(!fFunction) {
            this.result = null;
            return;
        }
        this.result = fFunction.apply(null, aArgs);
    };
    CBinaryOperatorToken.prototype.operators = {
        "+": function (x, y) {return y + x},
        "-": function (x, y) {return y - x},
        "*": function (x, y) {return y * x},
        "/": function (x, y) {return y / x},
        "%": function (x, y) {return y % x},
        "^": function (x, y) {return Math.pow(y, x)}
    };
    CBinaryOperatorToken.prototype.isOperator = function() {
        return true;
    };

    function CUnaryOperatorToken(oQueue, sName) {
        CTokenBase.call(this, oQueue);
        this.name = sName;
    }
    InitClass(CUnaryOperatorToken, CTokenBase, undefined);
    CUnaryOperatorToken.prototype.argumentsCount = 1;
    CUnaryOperatorToken.prototype.precedence = 8;
    CUnaryOperatorToken.prototype._calculate = function(aArgs, oVarMap) {
        var fFunction = this.operators[this.name];
        if(!fFunction) {
            this.result = null;
            return;
        }
        this.result = fFunction.apply(null, aArgs);
    };
    CUnaryOperatorToken.prototype.operators = {
        "+": function (x) {return +x},
        "-": function (x) {return -x}
    };
    CUnaryOperatorToken.prototype.isOperator = function() {
        return true;
    };
    function CLeftParenToken(oQueue) {
        CTokenBase.call(this, oQueue);
    }
    InitClass(CLeftParenToken, CTokenBase, undefined);
    CLeftParenToken.prototype.argumentsCount = 0;
    CLeftParenToken.prototype.precedence = 1;
    CLeftParenToken.prototype._calculate = function(aArgs, oVarMap) {
    };
    function CRightParenToken(oQueue) {
        CTokenBase.call(this, oQueue);
    }
    InitClass(CRightParenToken, CTokenBase, undefined);
    CRightParenToken.prototype.argumentsCount = 0;
    CRightParenToken.prototype.precedence = 1;
    CRightParenToken.prototype._calculate = function(aArgs, oVarMap) {
    };

    function CArgSeparatorToken(oQueue) {
        CTokenBase.call(this, oQueue);
    }
    InitClass(CRightParenToken, CTokenBase, undefined);
    CArgSeparatorToken.prototype.argumentsCount = 0;
    CArgSeparatorToken.prototype.precedence = 9;
    CArgSeparatorToken.prototype._calculate = function(aArgs, oVarMap) {
    };


    var OPERATORS_MAP = {
        "+": true,
        "-": true,
        "*": true,
        "/": true,
        "%": true,
        "^": true
    };

    var FUNC_REGEXPSTR = "(abs\|acos\|asin\|atan\|ceil\|cos\|cosh\|deg\|exp\|floor\|ln\|max\|min\|rad\|rand\|sin\|sinh\|sqrt\|tan\|tanh)";
    var FUNC_REGEXP = new RegExp(FUNC_REGEXPSTR, "g");

    var CONST_REGEXPSTR = "(pi\|e)";
    var CONST_REGEXP = new RegExp(CONST_REGEXPSTR, "g");

    var NUMBER_REGEXPSTR = "[-+]?[0-9]*\\.?[0-9]+([eE][-+]?[0-9]+)?";
    var NUMBER_REGEXP = new RegExp(NUMBER_REGEXPSTR, "g");


    var PARSER_FLAGS_CONSTVAR = 1;
    var PARSER_FLAGS_FUNCTION = 2;
    var PARSER_FLAGS_BINARYOP = 4;
    var PARSER_FLAGS_UNARYOP  = 8;
    var PARSER_FLAGS_LEFTPAR  = 16;
    var PARSER_FLAGS_RIGHTPAR = 32;
    var PARSER_FLAGS_ARGSEP   = 64;
    function CFormulaParser(sFormula, oVarMap) {
        this.formula = sFormula;
        var aVarNames = [];
        for(var sVarName in oVarMap) {
            if(oVarMap.hasOwnProperty(sVarName)) {
                aVarNames.push(sVarName);
            }
        }
        this.varMap = oVarMap;
        this.variables = aVarNames;
        this.pos = 0;
        this.flags = 0;
        this.queue = new CParseQueue();
    }
    CFormulaParser.prototype.getResult = function() {
        var oParseResult = this.parse();
        if(!oParseResult) {
            return null;
        }
        return oParseResult.calculate(this.varMap);
    };
    CFormulaParser.prototype.setFlag = function(nMask, bVal) {
        if(bVal){
            this.flags |= nMask;
        }
        else{
            this.flags &= (~nMask);
        }
    };
    CFormulaParser.prototype.getFlag = function(nMask) {
        return (this.flags & nMask) === nMask;
    };
    CFormulaParser.prototype.parse = function() {
        this.pos = 0;
        this.setFlag(PARSER_FLAGS_CONSTVAR, true);
        this.setFlag(PARSER_FLAGS_FUNCTION, true);
        this.setFlag(PARSER_FLAGS_BINARYOP, false);
        this.setFlag(PARSER_FLAGS_UNARYOP, true);
        this.setFlag(PARSER_FLAGS_LEFTPAR, true);
        this.setFlag(PARSER_FLAGS_RIGHTPAR, false);
        this.setFlag(PARSER_FLAGS_ARGSEP, false);
        var oCurToken;
        var aStack = [];
        var aFunctionsStack = [];
        var oLastToken = null;
        var oToken;
        var oLastFunction;
        while (oCurToken = this.parseCurrent()){
            if(oCurToken instanceof CConstantToken || oCurToken instanceof CVariableToken) {
                if(!this.getFlag(PARSER_FLAGS_CONSTVAR)) {
                    return null;
                }
                this.queue.add(oCurToken);
                this.setFlag(PARSER_FLAGS_CONSTVAR, false);
                this.setFlag(PARSER_FLAGS_FUNCTION, false);
                this.setFlag(PARSER_FLAGS_BINARYOP, true);
                this.setFlag(PARSER_FLAGS_UNARYOP, false);
                this.setFlag(PARSER_FLAGS_LEFTPAR, false);
                this.setFlag(PARSER_FLAGS_RIGHTPAR, true);
                this.setFlag(PARSER_FLAGS_ARGSEP, aFunctionsStack.length > 0);
            }
            else if(oCurToken instanceof CFunctionToken) {
                if(!this.getFlag(PARSER_FLAGS_FUNCTION)) {
                    return null;
                }
                aStack.push(oCurToken);
                this.setFlag(PARSER_FLAGS_CONSTVAR, false);
                this.setFlag(PARSER_FLAGS_FUNCTION, false);
                this.setFlag(PARSER_FLAGS_BINARYOP, false);
                this.setFlag(PARSER_FLAGS_UNARYOP, false);
                this.setFlag(PARSER_FLAGS_LEFTPAR, true);
                this.setFlag(PARSER_FLAGS_RIGHTPAR, false);
                this.setFlag(PARSER_FLAGS_ARGSEP, false);
            }
            else if(oCurToken instanceof CArgSeparatorToken){
                if(!this.getFlag(PARSER_FLAGS_ARGSEP)) {
                    return null;
                }
                if(aFunctionsStack.length > 0){
                    while(aStack.length > 0 && !(aStack[aStack.length-1] instanceof CLeftParenToken)){
                        oToken = aStack.pop();

                        this.queue.add(oToken);
                    }
                    if(aStack.length === 0){
                        return null;
                    }
                    oLastFunction = aFunctionsStack[aFunctionsStack.length-1];
                    oLastFunction.addOperand(this.queue.last());
                    if(oLastFunction.getOperandsCount() >= oLastFunction.getArgumentsCount()){
                        return null;
                    }
                }
                else{
                    return null;
                }
                this.setFlag(PARSER_FLAGS_CONSTVAR, true);
                this.setFlag(PARSER_FLAGS_FUNCTION, true);
                this.setFlag(PARSER_FLAGS_BINARYOP, false);
                this.setFlag(PARSER_FLAGS_UNARYOP, true);
                this.setFlag(PARSER_FLAGS_LEFTPAR, true);
                this.setFlag(PARSER_FLAGS_RIGHTPAR, false);
                this.setFlag(PARSER_FLAGS_ARGSEP, false);
            }
            else if(oCurToken instanceof CLeftParenToken){
                if(!this.getFlag(PARSER_FLAGS_LEFTPAR)) {
                    return null;
                }
                aStack.push(oCurToken);
                if(oLastToken && oLastToken.isFunction(oLastToken)){
                    aFunctionsStack.push(oLastToken);
                }
                this.setFlag(PARSER_FLAGS_CONSTVAR, true);
                this.setFlag(PARSER_FLAGS_FUNCTION, true);
                this.setFlag(PARSER_FLAGS_BINARYOP, false);
                this.setFlag(PARSER_FLAGS_UNARYOP, true);
                this.setFlag(PARSER_FLAGS_LEFTPAR, true);
                this.setFlag(PARSER_FLAGS_RIGHTPAR, true);
                this.setFlag(PARSER_FLAGS_ARGSEP, false);
            }
            else if(oCurToken instanceof CRightParenToken){
                while(aStack.length > 0 && !(aStack[aStack.length-1] instanceof CLeftParenToken)){
                    oToken = aStack.pop();
                    this.queue.add(oToken);
                }

                if(aStack.length === 0){
                    return null;
                }
                aStack.pop();//remove left paren
                if(aStack[aStack.length-1] && aStack[aStack.length-1].isFunction()){
                    aFunctionsStack.pop();
                    oLastFunction = aStack[aStack.length-1];
                    oLastFunction.addOperand(this.queue.last());
                    if(oLastFunction.getOperandsCount() !== oLastFunction.getArgumentsCount()){
                        return null;
                    }
                    oToken = aStack.pop();
                    this.queue.add(oToken);
                }
                this.setFlag(PARSER_FLAGS_CONSTVAR, false);
                this.setFlag(PARSER_FLAGS_FUNCTION, false);
                this.setFlag(PARSER_FLAGS_BINARYOP, true);
                this.setFlag(PARSER_FLAGS_UNARYOP, false);
                this.setFlag(PARSER_FLAGS_LEFTPAR, false);
                this.setFlag(PARSER_FLAGS_RIGHTPAR, true);
                this.setFlag(PARSER_FLAGS_ARGSEP, aFunctionsStack.length > 0);
            }
            else if(oCurToken.isOperator()){
                if(oCurToken instanceof CUnaryOperatorToken){
                    if(!this.getFlag(PARSER_FLAGS_UNARYOP)){
                        return null;
                    }
                    this.setFlag(PARSER_FLAGS_UNARYOP, false);
                }
                else{
                    if(!this.getFlag(PARSER_FLAGS_BINARYOP)) {
                        return null;
                    }
                    this.setFlag(PARSER_FLAGS_UNARYOP, true);
                }
                while(aStack.length > 0 && (!(aStack[aStack.length-1] instanceof CLeftParenToken) && aStack[aStack.length-1].getPrecedence() >= oCurToken.getPrecedence())){
                    oToken = aStack.pop();
                    this.queue.add(oToken);
                }
                this.setFlag(PARSER_FLAGS_CONSTVAR, true);
                this.setFlag(PARSER_FLAGS_FUNCTION, true);
                this.setFlag(PARSER_FLAGS_BINARYOP, false);
                this.setFlag(PARSER_FLAGS_LEFTPAR, true);
                this.setFlag(PARSER_FLAGS_RIGHTPAR, false);
                this.setFlag(PARSER_FLAGS_ARGSEP, false);
                aStack.push(oCurToken);
            }

            oLastToken = oCurToken;
        }

        if(this.pos < this.formula.length){
            return null;
        }
        while (aStack.length > 0){
            oCurToken = aStack.pop();
            if(oCurToken instanceof CLeftParenToken || oCurToken instanceof CRightParenToken){
                return null;
            }
            this.queue.add(oCurToken);
        }
        return this.queue;
    };
    CFormulaParser.prototype.isOperator = function(sSymbol) {
        return !!OPERATORS_MAP[sSymbol];
    };
    CFormulaParser.prototype.parseCurrent = function() {
        //skip spaces
        while(this.formula[this.pos] == " ") {
            ++this.pos;
        }
        if(this.pos >= this.formula.length) {
            return null;
        }
        var sCurSymbol = this.formula[this.pos];
        if(sCurSymbol === "(") {
            ++this.pos;
            return new CLeftParenToken(this.queue);
        }
        if(sCurSymbol === ")") {
            ++this.pos;
            return new CRightParenToken(this.queue);
        }
        if(sCurSymbol === ",") {
            ++this.pos;
            return new CArgSeparatorToken(this.queue);
        }
        if(this.isOperator(sCurSymbol)) {
            ++this.pos;
            return this.parseOperator(sCurSymbol);
        }
        //check function
        var oRet = this.checkExpression(FUNC_REGEXP, this.parseFunction);
        if(oRet){
            return oRet;
        }
        for(var nVarName = 0; nVarName < this.variables.length; ++nVarName) {
            var sVarName = this.variables[nVarName];
            if(this.formula.indexOf(sVarName, this.pos) === this.pos) {
                this.pos += sVarName.length;
                return new CVariableToken(this.queue, sVarName);
            }
        }
        if(oRet){
            return oRet;
        }
        oRet = this.checkExpression(CONST_REGEXP, this.parseConst);
        if(oRet){
            return oRet;
        }
        oRet = this.checkExpression(NUMBER_REGEXP, this.parseNumber);
        if(oRet){
            return oRet;
        }
        return null;
    };
    CFormulaParser.prototype.parseFunction = function(nStartPos, nEndPos){
        var sFunction = this.formula.slice(nStartPos, nEndPos);
        if(CFunctionToken.prototype.functions[sFunction]) {
            return new CFunctionToken(this.queue, sFunction);
        }
        return null;
    };
    CFormulaParser.prototype.parseConst = function(nStartPos, nEndPos){
        var sConst = this.formula.slice(nStartPos, nEndPos);
        if(sConst === "pi") {
            return new CConstantToken(this.queue, Math.PI);
        }
        else if(sConst === "e") {
            return new CConstantToken(this.queue, Math.E);
        }
        return null;
    };
    CFormulaParser.prototype.parseNumber = function(nStartPos, nEndPos){
        var sNumber = this.formula.slice(nStartPos, nEndPos);
        var fNumer = parseFloat(sNumber);
        if(AscFormat.isRealNumber(fNumer)) {
            return  new CConstantToken(this.queue, fNumer);
        }
        return null;
    };
    CFormulaParser.prototype.checkExpression = function(oRegExp, fCallback){
        oRegExp.lastIndex = this.pos;
        var oRes = oRegExp.exec(this.formula);
        if(oRes && oRes.index === this.pos){
            var ret = fCallback.call(this, this.pos, oRegExp.lastIndex);
            this.pos = oRegExp.lastIndex;
            return ret;
        }
        return null;
    };
    CFormulaParser.prototype.parseOperator = function(sOperator) {
        if(sOperator === "+" || sOperator === "-") {
            if(this.getFlag(PARSER_FLAGS_UNARYOP)) {
                return new CUnaryOperatorToken(this.queue, sOperator);
            }
        }
        return new CBinaryOperatorToken(this.queue, sOperator);
    };
    //--------------------------------------------------------------------------


    var STATE_FLAG_SELECTED = 1;
    var STATE_FLAG_HOVERED = 2;

    var CONTROL_TYPE_UNKNOWN = 0;
    var CONTROL_TYPE_HEADER = 1;
    var CONTROL_TYPE_TOOLBAR = 2;
    var CONTROL_TYPE_SEQ_LIST_CONTAINER = 3;
    var CONTROL_TYPE_SCROLL_VERT = 4;
    var CONTROL_TYPE_SCROLL_HOR = 5;
    var CONTROL_TYPE_SEQ_LIST = 6;
    var CONTROL_TYPE_ANIM_SEQ = 7;
    var CONTROL_TYPE_ANIM_GROUP_LIST = 8;
    var CONTROL_TYPE_ANIM_GROUP = 9;
    var CONTROL_TYPE_ANIM_ITEM = 10;
    var CONTROL_TYPE_LABEL = 11;
    var CONTROL_TYPE_BUTTON = 12;
    var CONTROL_TYPE_IMAGE = 13;
    var CONTROL_TYPE_TIMELINE_CONTAINER = 14;
    var CONTROL_TYPE_TIMELINE = 15;
    var CONTROL_TYPE_EFFECT_BAR = 16;

    var LEFT_TIMELINE_INDENT = 14 * AscCommon.g_dKoef_pix_to_mm;
    var LABEL_TIMELINE_WIDTH = 155 * AscCommon.g_dKoef_pix_to_mm;

    function CControl(oParentControl) {
        AscFormat.ExecuteNoHistory(function() {
            AscFormat.CShape.call(this);
            this.setRecalculateInfo();
            this.setBDeleted(false);
            this.setLayout(0, 0, 0, 0);
        }, this, []);

        this.parent = editor.WordControl.m_oLogicDocument.Slides[0];
        this.parentControl = oParentControl;
        this.state = 0;
        this.hidden = false;
        this.previous = null;
        this.next = null;
    }
    InitClass(CControl, AscFormat.CShape, CONTROL_TYPE_UNKNOWN);
    CControl.prototype.DEFALT_WRAP_OBJECT = {oTxWarpStruct: null, oTxWarpStructParamarks: null, oTxWarpStructNoTransform: null, oTxWarpStructParamarksNoTransform: null};
    CControl.prototype.setHidden = function(bVal) {
        if(this.hidden !== bVal) {
            this.hidden = bVal;
            this.onUpdate();
        }
    };
    CControl.prototype.show = function() {
        this.setHidden(false);
    };
    CControl.prototype.hide = function() {
        this.setHidden(true);
    };
    CControl.prototype.isHidden = function() {
        return this.hidden;
    };
    CControl.prototype.notAllowedWithoutId = function() {
        return false;
    };
    //define shape methods
    CControl.prototype.getBodyPr = function () {
        return this.bodyPr;
    };
    CControl.prototype.getScrollOffsetX = function (oChild) {
        return 0;
    };
    CControl.prototype.getScrollOffsetY = function (oChild) {
        return 0;
    };
    CControl.prototype.getParentScrollOffsetX = function (oChild) {
        if(this.parentControl) {
            return this.parentControl.getScrollOffsetX(oChild);
        }
        return 0;
    };
    CControl.prototype.getParentScrollOffsetY = function (oChild) {
        if(this.parentControl) {
            return this.parentControl.getScrollOffsetY(oChild);
        }
        return 0;
    };
    CControl.prototype.getFullTransformMatrix = function () {
        return this.transform;
    };
    CControl.prototype.getInvFullTransformMatrix = function () {
        return this.invertTransform;
    };
    CControl.prototype.multiplyParentTransforms = function(oLocalTransform) {
        var oMT = AscCommon.global_MatrixTransformer;
        var oTransform = oMT.CreateDublicateM(oLocalTransform);
        var oScrollMatrix = new AscCommon.CMatrix();
        oScrollMatrix.tx = this.getParentScrollOffsetX(this);
        oScrollMatrix.ty = this.getParentScrollOffsetY(this);
        oMT.MultiplyAppend(oTransform, oScrollMatrix);
        var oParentTransform = this.parentControl && this.parentControl.getFullTransformMatrix();
        oParentTransform && oMT.MultiplyAppend(oTransform, oParentTransform);
        return oTransform;
    };
    CControl.prototype.getFullTransform = function() {
        return this.transform;
    };
    CControl.prototype.getFullTextTransform = function() {
        return this.transformText;
    };

    CControl.prototype.recalculate = function() {
        AscFormat.CShape.prototype.recalculate.call(this);
    };
    CControl.prototype.recalculateBrush = function () {
        this.brush = null;
    };
    CControl.prototype.recalculatePen = function () {
        this.pen = null;
    };
    CControl.prototype.recalculateContent = function () {
    };
    CControl.prototype.recalculateGeometry = function() {
        //this.calcGeometry = AscFormat.CreateGeometry("rect");
        //this.calcGeometry.Recalculate(this.extX, this.extY);
    };
    CControl.prototype.recalculateTransform = function() {
        if(!this.transform) {
            this.transform = new AscCommon.CMatrix();
        }
        var tx = this.getLeft();
        var ty = this.getTop();
        this.x = tx;
        this.y = ty;
        this.rot = 0;
        this.extX = this.getWidth();
        this.extY = this.getHeight();
        this.flipH = false;
        this.flipV = false;
        ty += this.getParentScrollOffsetY(this);
        var oCurParent = this.parentControl;

        if(oCurParent) {
            tx += oCurParent.transform.tx;
            ty += oCurParent.transform.ty
        }
        this.transform.tx = tx;
        this.transform.ty = ty;
        if(!this.invertTransform) {
            this.invertTransform = new AscCommon.CMatrix();
        }
        this.invertTransform.tx = -tx;
        this.invertTransform.ty = -ty;
        this.localTransform = this.transform;
    };
    CControl.prototype.recalculateTransformText = function() {
        if(!this.transformText) {
            this.transformText = new AscCommon.CMatrix();
        }
        this.transformText.tx = this.transform.tx;
        this.transformText.ty = this.transform.ty;

        if(!this.invertTransformText) {
            this.invertTransformText = new AscCommon.CMatrix();
        }
        this.invertTransformText.tx = -this.transform.tx;
        this.invertTransformText.ty = -this.transform.ty;
        this.localTransformText = this.transformText;
    };
    CControl.prototype.recalculateBounds = function() {
        var dX = this.transform.tx;
        var dY = this.transform.ty;
        this.bounds.reset(dX, dY, dX + this.getWidth(), dY + this.getHeight())
    };
    CControl.prototype.recalculateSnapArrays = function() {
    };
    CControl.prototype.checkAutofit = function (bIgnoreWordShape) {
        return false;
    };
    CControl.prototype.checkTextWarp = function(oContent, oBodyPr, dWidth, dHeight, bNeedNoTransform, bNeedWarp) {
        return this.DEFALT_WRAP_OBJECT;
    };
    CControl.prototype.addToRecalculate = function() {
    };
    CControl.prototype.canHandleEvents = function() {
        return true;
    };
    CControl.prototype.getPenWidth = function(graphics) {
        var fScale = graphics.m_oCoordTransform.sx;
        var nPenW = AscCommon.AscBrowser.convertToRetinaValue(1, true)/fScale;
        return nPenW;
    };
    CControl.prototype.draw = function (graphics) {
        if(this.isHidden()){
            return false;
        }
        if(!this.checkUpdateRect(graphics.updatedRect)) {
            return false;
        }

        this.recalculateTransform();
        this.recalculateTransformText();

        var sFillColor = this.getFillColor();
        var sOutlineColor = this.getOutlineColor();
        var oColor;
        if(sOutlineColor || sFillColor) {
            graphics.SaveGrState();
            graphics.transform3(this.transform);
            var x = 0;
            var y = 0;
            var extX = this.getWidth();
            var extY = this.getHeight();
            if(sFillColor) {
                oColor = AscCommon.RgbaHexToRGBA(sFillColor);
                graphics.b_color1(oColor.R, oColor.G, oColor.B, 0xFF);
                graphics.rect(x, y, extX, extY);
                graphics.df();
            }
            if(sOutlineColor) {
                oColor = AscCommon.RgbaHexToRGBA(sOutlineColor);
                graphics.SetIntegerGrid(true);

                var nPenW = this.getPenWidth(graphics);
                //graphics.p_width(100);//AscCommon.AscBrowser.convertToRetinaValue(1, true);
                graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
                graphics.drawHorLine(0, y, x, x + extX, nPenW);
                graphics.drawHorLine(0, y + extY, x, x + extX, nPenW);
                graphics.drawVerLine(2, x, y, y + extY, nPenW);
                graphics.drawVerLine(2, x + extX, y, y + extY, nPenW);
                graphics.ds();
            }
            graphics.RestoreGrState();
        }
        AscFormat.CShape.prototype.draw.call(this, graphics);
        return true;

    };
    CControl.prototype.hit = function(x, y) {
        if(this.parentControl && !this.parentControl.hit(x, y)) {
            return false;
        }
        var oInv = this.invertTransform;
        var tx = oInv.TransformPointX(x, y);
        var ty = oInv.TransformPointY(x, y);
        return tx >= 0 && tx <= this.extX && ty >= 0 && ty <= this.extY;
    };
    CControl.prototype.isHovered = function() {
        return this.getStateFlag(STATE_FLAG_HOVERED);
    };
    CControl.prototype.isActive = function() {
        if(this.parentControl) {
            if(!this.eventListener && this.parentControl.isEventListener(this)) {
                return true;
            }
        }
        return false;
    };
    CControl.prototype.setStateFlag = function(nFlag, bValue) {
        var nOldState = this.state;
        if(bValue) {
            this.state |= nFlag;
        }
        else {
            this.state &= (~nFlag);
        }
        if(nOldState !== this.state) {
            this.onUpdate();
        }
    };
    CControl.prototype.setHoverState = function() {
        this.setStateFlag(STATE_FLAG_HOVERED, true);
    };
    CControl.prototype.setNotHoverState = function() {
        this.setStateFlag(STATE_FLAG_HOVERED, false);
    };
    CControl.prototype.getStateFlag = function(nFlag) {
        return (this.state & nFlag) !== 0;
    };
    CControl.prototype.onMouseMove = function (e, x, y) {
        if(e.IsLocked) {
            return false;
        }
        if(!this.canHandleEvents()) {
            return false;
        }
        var bHover = this.hit(x, y);
        var bRet = bHover !== this.isHovered();
        if(bHover) {
            this.setHoverState();
        }
        else {
            this.setNotHoverState();
        }
        return bRet;
    };
    CControl.prototype.onMouseDown = function (e, x, y) {
        if(!this.canHandleEvents()) {
            return false;
        }
        if(this.hit(x, y)) {
            if(this.parentControl) {
                this.parentControl.setEventListener(this);
            }
            return true;
        }
        return false;
    };
    CControl.prototype.onMouseUp = function (e, x, y) {
        if(this.parentControl) {
            this.parentControl.setEventListener(null);
        }
        return false;
    };
    CControl.prototype.onMouseWheel = function (e, deltaY, X, Y) {
        return false;
    };
    CControl.prototype.onUpdate = function() {
        if(this.parentControl) {
            var oBounds = this.getBounds();
            this.parentControl.onChildUpdate(oBounds);
        }
    };
    CControl.prototype.onChildUpdate = function(oBounds) {
        if(this.parentControl) {
            this.parentControl.onChildUpdate(oBounds);
        }
    };
    CControl.prototype.getCursorInfo = function(e, x, y) {
        if(!this.hit(x, y)) {
            return null;
        }
        else {
            return {
                cursorType: "default",
                tooltip: this.getTooltipText()
            }
        }
    };
    CControl.prototype.checkUpdateRect = function(oUpdateRect) {
        var oBounds = this.getBounds();
        if(oUpdateRect && oBounds) {
            if(!oUpdateRect.isIntersectOther(oBounds)) {
                return false;
            }
        }
        return true;
    };
    CControl.prototype.recalculate = function() {
        AscFormat.CShape.prototype.recalculate.call(this);
    };
    CControl.prototype.setLayout = function(dX, dY, dExtX, dExtY) {
        if(!this.spPr) {
            this.spPr = new AscFormat.CSpPr();
        }
        if(!this.spPr.xfrm) {
            this.spPr.xfrm = new AscFormat.CXfrm();
        }

        this.spPr.xfrm.offX = dX;
        this.spPr.xfrm.offY = dY;
        this.spPr.xfrm.extX = dExtX;
        this.spPr.xfrm.extY = dExtY;
        this.handleUpdateExtents();
    };
    CControl.prototype.getLeft = function() {
        return this.spPr.xfrm.offX;
    };
    CControl.prototype.getTop = function() {
        return this.spPr.xfrm.offY;
    };
    CControl.prototype.getRight = function() {
        return this.spPr.xfrm.offX + this.spPr.xfrm.extX;
    };
    CControl.prototype.getBottom = function() {
        return this.spPr.xfrm.offY + this.spPr.xfrm.extY;
    };
    CControl.prototype.getWidth = function() {
        return this.spPr.xfrm.extX;
    };
    CControl.prototype.getHeight = function() {
        return this.spPr.xfrm.extY;
    };
    CControl.prototype.getBounds = function() {
        this.recalculateBounds();
        this.recalculateTransform();
        this.recalculateTransformText();
        return this.bounds;
    };
    CControl.prototype.convertRelToAbs = function(oPos) {
        var oAbsPos = {x: oPos.x, y: oPos.y};
        var oParent = this;
        while(oParent) {
            oAbsPos.x += oParent.getLeft();
            oAbsPos.y += oParent.getTop();
            oParent = oParent.parentControl;
        }
        return oAbsPos;
    };
    CControl.prototype.convertAbsToRel = function(oPos) {
        var oRelPos = {x: oPos.x, y: oPos.y};
        var oParent = this;
        while(oParent) {
            oRelPos.x -= oParent.getLeft();
            oRelPos.y -= oParent.getTop();
            oParent = oParent.parentControl;
        }
        return oRelPos;
    };
    CControl.prototype.getNext = function() {
        return this.next;
    };
    CControl.prototype.getPrevious = function() {
        return this.previous;
    };
    CControl.prototype.setNext = function(v) {
        this.next = v;
    };
    CControl.prototype.setPrevious = function(v) {
        this.previous = v;
    };
    CControl.prototype.setParentControl = function(v) {
        this.parentControl = v;
    };
    CControl.prototype.getTiming = function() {
        var oSlide = this.getSlide();
        if(oSlide) {
            return oSlide.timing;
        }
        return null;
    };
    CControl.prototype.getSlide = function() {
        var oSlide = null;
        if(editor.WordControl && editor.WordControl.m_oLogicDocument) {
            oSlide = editor.WordControl.m_oLogicDocument.GetCurrentSlide();
            return oSlide;
        }
        return null;
    };
    CControl.prototype.getSlideNum = function() {
        var oSlide = this.getSlide();
        if(oSlide) {
            return oSlide.num;
        }
        return -1;
    };
    CControl.prototype.getFillColor = function() {
        var sFillColor;
        var oSkin = AscCommon.GlobalSkin;
        if(this.isActive()) {
            sFillColor = oSkin.ThumbnailsPageOutlineActive;
        }
        else if(this.isHovered()) {
            sFillColor = oSkin.ScrollerHoverColor;
        }
        else {
            sFillColor = oSkin.BackgroundColorThumbnails;
        }
        return sFillColor;
    };
    CControl.prototype.getOutlineColor = function() {
        var sOutlineColor;
        var oSkin = AscCommon.GlobalSkin;
        if(this.isActive()) {
            sOutlineColor = oSkin.ScrollOutlineActiveColor;
        }
        else if(this.isHovered()) {
            sOutlineColor = oSkin.ThumbnailsPageOutlineHover;
        }
        else {
            sOutlineColor = oSkin.ScrollOutlineColor;
        }
        return sOutlineColor;
    };
    CControl.prototype.drawShdw = function() {

    };

    function CControlContainer(oParentControl) {
        CControl.call(this, oParentControl);
        this.children = [];
        this.recalcInfo.recalculateChildrenLayout = true;
        this.recalcInfo.recalculateChildren = true;

        this.eventListener = null;
    }
    InitClass(CControlContainer, CControl, CONTROL_TYPE_UNKNOWN);
    CControlContainer.prototype.isEventListener = function(oChild) {
        return this.eventListener === oChild;
    };
    CControlContainer.prototype.onScroll = function() {
    };
    CControlContainer.prototype.onStartScroll = function() {
    };
    CControlContainer.prototype.onEndScroll = function() {
    };
    CControlContainer.prototype.clear = function() {
        for(var nIdx = this.children.length - 1; nIdx > -1; --nIdx) {
            this.removeControl(this.children[nIdx]);
        }
    };
    CControlContainer.prototype.addControl = function(oChild) {
        var oLast = this.children[this.children.length - 1];
        this.children.push(oChild);
        if(oLast) {
            oLast.setNext(oChild);
            oChild.setPrevious(oLast);
            oChild.setParentControl(this);
        }
        return oChild;
    };
    CControlContainer.prototype.removeControl = function(oChild) {
        var nIdx = this.getChildIdx(oChild);
        this.removeByIdx(nIdx);
    };
    CControlContainer.prototype.removeByIdx = function(nIdx) {
        if(nIdx > -1 && nIdx < this.children.length) {
            var oChild = this.children[nIdx];
            oChild.setNext(null);
            oChild.setPrevious(null);
            oChild.setParentControl(null);
            var oPrev = this.children[nIdx - 1] || null;
            var oNext = this.children[nIdx + 1] || null;
            if(oPrev) {
                oPrev.setNext(oNext);
            }
            if(oNext) {
                oNext.setPrevious(oPrev);
            }
            this.children.splice(nIdx, 1);
        }
    };
    CControlContainer.prototype.getChildIdx = function(oChild) {
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            if(this.children[nChild] === oChild) {
                return nChild;
            }
        }
        return -1;
    };
    CControlContainer.prototype.getChildByType = function(nType) {
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            var oChild = this.children[nChild];
            if(oChild.getObjectType() === nType) {
                return oChild;
            }
        }
        return null;
    };
    CControlContainer.prototype.getChild = function(nIdx) {
        if(nIdx > -1 && nIdx < this.children.length) {
            return this.children[nIdx];
        }
    };
    CControlContainer.prototype.draw = function(graphics) {
        if(!CControl.prototype.draw.call(this, graphics)) {
            return false;
        }
        this.clipStart(graphics);
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            this.children[nChild].draw(graphics);
        }
        this.clipEnd(graphics);
        return true;
    };
    CControlContainer.prototype.clipStart = function(graphics) {
    };
    CControlContainer.prototype.clipEnd = function(graphics) {
    };
    CControlContainer.prototype.recalculateChildrenLayout = function() {
    };
    CControlContainer.prototype.recalculateChildren = function() {
    };
    CControlContainer.prototype.recalculate = function() {
        AscFormat.ExecuteNoHistory(function() {
            CControl.prototype.recalculate.call(this);
            if(this.recalcInfo.recalculateChildren) {
                this.recalculateChildren();
                this.recalcInfo.recalculateChildren = false;
            }
            if(this.recalcInfo.recalculateChildrenLayout) {
                this.recalculateChildrenLayout();
                this.recalcInfo.recalculateChildrenLayout = false;
            }
            for(var nChild = 0; nChild < this.children.length; ++nChild) {
                this.children[nChild].recalculate();
            }
        }, this, []);
    };
    CControlContainer.prototype.setLayout = function(dX, dY, dExtX, dExtY) {
        AscFormat.ExecuteNoHistory(function() {
            CControl.prototype.setLayout.call(this, dX, dY, dExtX, dExtY);
            this.recalcInfo.recalculateChildrenLayout = true;
        }, this, []);
    };
    CControlContainer.prototype.handleUpdateExtents = function() {
        this.recalcInfo.recalculateChildrenLayout = true;
        CControl.prototype.handleUpdateExtents.call(this);
    };
    CControlContainer.prototype.setEventListener = function(oChild) {
        if(oChild) {
            this.eventListener = oChild;
            if(this.parentControl) {
                this.parentControl.setEventListener(this);
            }
        }
        else {
            this.eventListener = null;
            if(this.parentControl) {
                this.parentControl.setEventListener(null);
            }
        }
    };
    CControlContainer.prototype.onMouseDown = function(e, x, y) {
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            if(this.children[nChild].onMouseDown(e, x, y)) {
                return true;
            }
        }
        return CControl.prototype.onMouseDown.call(this, e, x, y);
    };
    CControlContainer.prototype.onMouseMove = function(e, x, y) {
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            if(this.children[nChild].onMouseMove(e, x, y)) {
                return true;
            }
        }
        return CControl.prototype.onMouseMove.call(this, e, x, y);
    };
    CControlContainer.prototype.onMouseUp = function(e, x, y) {
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            if(this.children[nChild].onMouseUp(e, x, y)) {
                return true;
            }
        }
        return CControl.prototype.onMouseUp.call(this, e, x, y);
    };
    CControlContainer.prototype.onMouseWheel = function(e, deltaY, X, Y) {
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            if(this.children[nChild].onMouseWheel(e, deltaY, X, Y)) {
                return true;
            }
        }
        return CControl.prototype.onMouseWheel.call(this, e, deltaY, X, Y);
    };
    CControlContainer.prototype.isScrolling = function() {
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            var oChild = this.children[nChild];
            if(oChild.isOnScroll && oChild.isOnScroll()) {
                return true;
            }
        }
        return false;
    };
    CControlContainer.prototype.canHandleEvents = function() {
        return false;
    };
    CControlContainer.prototype.onResize = function() {
        this.handleUpdateExtents();
        this.recalculate();
    };


    function CTopControl(oDrawer) {
        CControlContainer.call(this, null);
        this.drawer = oDrawer;
    }
    InitClass(CTopControl, CControlContainer, CONTROL_TYPE_UNKNOWN);
    CTopControl.prototype.onUpdateRect = function(oBounds) {
        if(this.drawer) {
            var oSlide = this.getSlide();
            if(oSlide) {
                this.drawer.OnAnimPaneChanged(oSlide.num, oBounds);
            }
        }
    };
    CTopControl.prototype.onUpdate = function() {
        var oBounds = this.getBounds();
        this.onUpdateRect(oBounds);
    };
    CTopControl.prototype.onChildUpdate = function(oBounds) {
        this.onUpdateRect(oBounds);
    };
    CTopControl.prototype.onResize = function() {
        this.setLayout(0, 0, this.drawer.GetWidth(), this.drawer.GetHeight());
        CControlContainer.prototype.onResize.call(this);
        this.onUpdate();
    };

    function CSeqListContainer(oDrawer) {
        CTopControl.call(this, oDrawer);
        this.seqList = this.addControl(new CSeqList(this));
    }
    InitClass(CSeqListContainer, CTopControl, CONTROL_TYPE_SEQ_LIST_CONTAINER);
    CSeqListContainer.prototype.getScrollOffsetY = function(oChild) {
        return 0;
    };
    CSeqListContainer.prototype.recalculateChildrenLayout = function() {
        this.seqList.setLayout(0, 0, this.getWidth(), this.seqList.getHeight());
        this.seqList.recalculate();
        this.setLayout(0, 0, this.seqList.getWidth(), this.seqList.getHeight());
    };
    CSeqListContainer.prototype.clipStart = function(graphics) {

    };
    CSeqListContainer.prototype.clipEnd = function(graphics) {
    };
    CSeqListContainer.prototype.onScroll = function() {
        this.onUpdate();
    };
    CSeqListContainer.prototype.getFillColor = function() {
        return null;
    };
    CSeqListContainer.prototype.getOutlineColor = function() {
        return null;
    };
    CSeqListContainer.prototype.onMouseWheel = function (e, deltaY, X, Y) {
        return false;
    };

    var SCROLL_TIMER_INTERVAL = 200;
    function CScrollBase(oParentControl, oContainer, oChild) {
        CControlContainer.call(this, oParentControl);
        this.addControl(new CButton(this, function(e, x, y) {
            if(this.hit(x, y)) {
                this.parentControl.setEventListener(this);
                this.parentControl.startScroll(-ANIM_ITEM_HEIGHT);
            }
        }, null, function(e, x, y) {
            this.parentControl.setEventListener(null);
            this.parentControl.endScroll();
        }));//left or top button
        this.addControl(new CButton(this, function(e, x, y) {
            if(this.hit(x, y)) {
                this.parentControl.setEventListener(this);
                this.parentControl.startScroll(ANIM_ITEM_HEIGHT);
            }
        }, null, function(e, x, y) {
            this.parentControl.setEventListener(null);
            this.parentControl.endScroll();
        }));//right or bottom button
        this.container = oContainer;
        this.scrolledChild = oChild;
        this.scrollOffset = 0;
        this.tmpScrollOffset = null;
        this.startScrollerPos = null;
        this.startScrollTop = null;
        this.timerId = null;
    }
    InitClass(CScrollBase, CControlContainer, CONTROL_TYPE_UNKNOWN);
    CScrollBase.prototype.getScrollOffset = function() {
        if(this.tmpScrollOffset !== null) {
            return this.tmpScrollOffset;
        }
        this.checkOffset();
        return this.scrollOffset;
    };
    CScrollBase.prototype.checkOffset = function() {
        this.scrollOffset = Math.max(0, Math.min(this.scrollOffset, this.getMaxScrollOffset()));
    };
    CScrollBase.prototype.setTmpScroll = function (val) {
        this.tmpScrollOffset = Math.max(0, Math.min(this.getMaxScrollOffset(), val));
        this.parentControl.onScroll();
        this.onUpdate();
    };
    CScrollBase.prototype.clearTmpScroll = function () {
        if(this.tmpScrollOffset !== null) {
            this.scrollOffset = this.tmpScrollOffset;
            this.tmpScrollOffset = null;
            this.parentControl.onScroll();
            this.onUpdate();
        }
    };
    CScrollBase.prototype.getMaxScrollOffset = function (val) {
        return 0;
    };
    CScrollBase.prototype.getScrollerX = function(dScrollOffset) {
        return 0;
    };
    CScrollBase.prototype.getScrollerY = function(dScrollOffset) {
        return 0;
    };
    CScrollBase.prototype.getScrollerWidth = function(dScrollOffset) {
        return 0;
    };
    CScrollBase.prototype.getScrollerHeight = function(dScrollOffset) {
        return 0;
    };
    CScrollBase.prototype.hitInScroller = function(x, y) {
        if(this.isHidden()) {
            return false;
        }
        var oInv = this.getInvFullTransformMatrix();
        var tx = oInv.TransformPointX(x, y);
        var ty = oInv.TransformPointY(x, y);
        var l = this.getScrollerX();
        var t = this.getScrollerY();
        var r = l + this.getScrollerWidth();
        var b = t + this.getScrollerHeight();
        return tx >= l && tx <= r && ty >= t && ty <= b;
    };
    CScrollBase.prototype.startScroll = function (step) {
        this.endScroll();
        var oScroll = this;
        this.tmpScrollOffset = this.getScrollOffset();
        oScroll.addScroll(step);
        this.timerId = setInterval(function () {
            oScroll.addScroll(step);
        }, SCROLL_TIMER_INTERVAL);
    };
    CScrollBase.prototype.addScroll = function (step) {
        this.setTmpScroll(this.tmpScrollOffset + step);
        this.parentControl.onScroll();
    };
    CScrollBase.prototype.endScroll = function () {
        if(this.timerId !== null) {
            clearInterval(this.timerId);
            this.timerId = null;
        }
        this.clearTmpScroll();
        this.setStateFlag(STATE_FLAG_SELECTED, false);
        this.startScrollerPos = null;
        this.startScrollTop = null;
    };
    CScrollBase.prototype.isOnScroll = function (step) {
        return this.timerId !== null || this.parentControl.isEventListener(this);
    };
    CScrollBase.prototype.getFillColor = function() {
        return null;
    };
    CScrollBase.prototype.getOutlineColor = function() {
        return null;
    };

    function CScrollVert(oParentControl, oContainer, oChild) {
        CScrollBase.call(this, oParentControl, oContainer, oChild);
        this.topButton = this.children[0];
        this.bottomButton = this.children[1];
    }
    InitClass(CScrollVert, CScrollBase, CONTROL_TYPE_SCROLL_VERT);
    CScrollVert.prototype.recalculateChildrenLayout = function() {
        this.topButton.setLayout(0, 0, SCROLL_BUTTON_SIZE, SCROLL_BUTTON_SIZE);
        this.bottomButton.setLayout(0, this.getHeight() - SCROLL_BUTTON_SIZE, SCROLL_BUTTON_SIZE, SCROLL_BUTTON_SIZE);
    };
    CScrollVert.prototype.getRailHeight = function() {
        return this.getHeight() - this.children[0].getHeight() - this.children[1].getHeight();
    };
    CScrollVert.prototype.getRelScrollerPos = function(dScrollOffset) {
        return this.topButton.getBottom() + dScrollOffset * ((this.getRailHeight() - this.getScrollerHeight()) / (this.getMaxScrollOffset()));
    };
    CScrollVert.prototype.getScrollerX = function(dScrollOffset) {
        return 0;
    };
    CScrollVert.prototype.getScrollerY = function() {
        return this.getRelScrollerPos(this.getScrollOffset());
    };
    CScrollVert.prototype.getScrollerWidth = function(dScrollOffset) {
        return this.getWidth();
    };
    CScrollVert.prototype.getScrollerHeight = function() {
        var dRailH = this.getRailHeight();
        var dMinRailH = dRailH / 4;
        return Math.max(dMinRailH, dRailH * (dRailH / this.scrolledChild.getHeight()))
    };
    CScrollVert.prototype.getMaxScrollOffset = function() {
        return Math.max(0, this.scrolledChild.getHeight() - this.container.getHeight());
    };
    CScrollVert.prototype.getMaxRelScrollOffset = function() {
        return Math.max(0, this.getRailHeight() - this.getScrollerHeight());
    };
    CScrollVert.prototype.draw = function(graphics) {
        if(this.isHidden()){
            return false;
        }
        if(!this.checkUpdateRect(graphics.updatedRect)) {
            return false;
        }
        this.children[0].draw(graphics);
        this.children[1].draw(graphics);


        graphics.SaveGrState();
        var oSkin = AscCommon.GlobalSkin;
            //ScrollBackgroundColor     : "#EEEEEE",
            //ScrollOutlineColor        : "#CBCBCB",
            //ScrollOutlineHoverColor   : "#CBCBCB",
            //ScrollOutlineActiveColor  : "#ADADAD",
            //ScrollerColor             : "#F7F7F7",
            //ScrollerHoverColor        : "#C0C0C0",
            //ScrollerActiveColor       : "#ADADAD",
            //ScrollArrowColor          : "#ADADAD",
            //ScrollArrowHoverColor     : "#F7F7F7",
            //ScrollArrowActiveColor    : "#F7F7F7",
            //ScrollerTargetColor       : "#CFCFCF",
            //ScrollerTargetHoverColor  : "#F1F1F1",
            //ScrollerTargetActiveColor : "#F1F1F1",
        var x = this.getScrollerX();
        var y = this.getRelScrollerPos(this.getScrollOffset());
        var extX = this.getScrollerWidth();
        var extY = this.getScrollerHeight();
        graphics.transform3(this.transform);

        var sFillColor;
        var sOutlineColor;
        var oColor;
        if(this.isActive()) {
            sFillColor = oSkin.ScrollerActiveColor;
            sOutlineColor = oSkin.ScrollOutlineActiveColor;
        }
        else if(this.isHovered()) {
            sFillColor = oSkin.ScrollerHoverColor;
            sOutlineColor = oSkin.ScrollOutlineHoverColor;
        }
        else {
            sFillColor = oSkin.ScrollerColor;
            sOutlineColor = oSkin.ScrollOutlineColor;
        }
        oColor = AscCommon.RgbaHexToRGBA(sFillColor);
        graphics.b_color1(oColor.R, oColor.G, oColor.B, 0xFF);
        graphics.rect(x, y, extX, extY);
        graphics.df();
        oColor = AscCommon.RgbaHexToRGBA(sOutlineColor);

        graphics.SetIntegerGrid(true);
        var nPenW = this.getPenWidth(graphics);
        graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
        graphics.drawHorLine(0, y, x, x + extX, nPenW);
        graphics.drawHorLine(0, y + extY, x, x + extX, nPenW);
        graphics.drawVerLine(2, x, y, y + extY, nPenW);
        graphics.drawVerLine(2, x + extX, y, y + extY, nPenW);
        graphics.ds();
        graphics.RestoreGrState();
        return true;
    };
    CScrollVert.prototype.onMouseMove = function (e, x, y) {
        if(this.isHidden()) {
            return false;
        }
        var bRet = false;
        if(this.eventListener) {
            this.eventListener.onMouseMove(e, x, y);
            return true;
        }

        if(this.parentControl.isEventListener(this)){
            if(this.startScrollerPos === null) {
                this.startScrollerPos = y;
            }
            if(this.startScrollTop === null) {
                this.startScrollTop = this.getScrollOffset();
            }
            var dCoeff = this.getMaxScrollOffset()/this.getMaxRelScrollOffset();
            var dy = dCoeff*(y - this.startScrollerPos);
            this.setTmpScroll(dy  + this.startScrollTop);
            return true;
        }
        bRet |= this.children[0].onMouseMove(e, x, y);
        bRet |= this.children[1].onMouseMove(e, x, y);

        var bHit = this.hitInScroller(x, y);
        var nState = this.isHovered();
        if(this.isHovered()) {
            if(!bHit) {
                this.setStateFlag(STATE_FLAG_HOVERED, false);
                bRet = true;
            }
        }
        else {
            if(bHit) {
                this.setStateFlag(STATE_FLAG_HOVERED, true);
                bRet = true;
            }
        }
        //-----------------------------
        return bRet;
    };
    CScrollVert.prototype.onMouseDown = function (e, x, y) {
        var bRet = false;
        if(this.hit(x, y)) {
            bRet |= this.children[0].onMouseDown(e, x, y);
            bRet |= this.children[1].onMouseDown(e, x, y);
            if(!bRet) {
                if(this.hitInScroller(x, y)) {
                    this.startScrollerPos = y;
                    this.startScrollTop = this.getScrollOffset();
                    this.setStateFlag(STATE_FLAG_SELECTED, true);
                    this.parentControl.setEventListener(this);
                    this.parentControl.onScroll();
                    //-----------------------------
                }
                else {
                    this.parentControl.setEventListener(this);
                    var oInv = this.getInvFullTransformMatrix();
                    var ty = oInv.TransformPointY(x, y);
                    if(ty < this.getScrollerY()) {
                        this.startScroll(-ANIM_ITEM_HEIGHT);
                    }
                    else {
                        this.startScroll(ANIM_ITEM_HEIGHT);
                    }
                }
                return true;
            }
        }
        return bRet;
    };
    CScrollVert.prototype.onMouseUp = function (e, x, y) {
        this.endScroll();
        var bRet = false;
        if(this.eventListener) {
            bRet = this.eventListener.onMouseUp(e, x, y);
            this.eventListener = null;
            return bRet;
        }
        bRet |= this.children[0].onMouseUp(e, x, y);
        bRet |= this.children[1].onMouseUp(e, x, y);
        this.setEventListener(null);
        return bRet;
    };

    function CScrollHor(oParentControl, oContainer, oChild) {
        CScrollBase.call(this, oParentControl, oContainer, oChild);
    }
    InitClass(CScrollHor, CScrollBase, CONTROL_TYPE_SCROLL_HOR);

    function CSeqList(oParentControl) {
        CControlContainer.call(this, oParentControl);
        this.sequences = this.children;
    }
    InitClass(CSeqList, CControlContainer, CONTROL_TYPE_SEQ_LIST);
    CSeqList.prototype.getIndexLabelRight = function() {
        return 10;//TODO
    };
    CSeqList.prototype.recalculateChildren = function() {
        this.clear();
        var oTiming = this.getTiming();
        if(oTiming) {
            var aAllSeqs = oTiming.getRootSequences();
            var oLastSeqView = null;
            for(var nSeq = 0; nSeq < aAllSeqs.length; ++nSeq) {
                var oSeqView = new CAnimSequence(this, aAllSeqs[nSeq]);
                this.addControl(oSeqView);
                oLastSeqView = oSeqView;
            }
        }
    };
    CSeqList.prototype.recalculateChildrenLayout = function() {
        var dLastBottom = 0;
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            var oSeq = this.children[nChild];
            oSeq.setLayout(0, dLastBottom, this.getWidth(), 0);
            oSeq.recalculate();
            dLastBottom = oSeq.getBottom();
        }
        this.setLayout(this.getLeft(), this.getTop(), this.getWidth(), dLastBottom);
    };
    CSeqList.prototype.getFillColor = function() {
        return null;
    };
    CSeqList.prototype.getOutlineColor = function() {
        return null;
    };
    // CSeqList.prototype.draw = function(graphics) {
    //     if(!this.checkUpdateRect(graphics.updateRect)) {
    //         return false;
    //     }
    //     if(this.parentControl.isScrolling() && !this.bDrawTexture) {
    //         this.recalculateTransform();
    //         this.checkCachedTexture(graphics).draw(graphics, new AscCommon.CMatrix());
    //         return;
    //     }
    //     this.clearCachedTexture();
    //     return CControlContainer.prototype.draw.call(this, graphics);
    // };


    CSeqList.prototype.checkCachedTexture = function(graphics) {
        var dGraphicsScale = graphics.m_oCoordTransform.sx;
        if(this.cachedCanvas) {
            var dScale = this.cachedCanvas.scale;
            if(AscFormat.fApproxEqual(dScale, dGraphicsScale)) {
                return this.cachedCanvas;
            }
        }
        this.bDrawTexture = true;
        var oBaseTexture = this.getAnimTexture(dGraphicsScale);
        this.cachedCanvas = new CAnimTexture(this, oBaseTexture.canvas, oBaseTexture.scale, oBaseTexture.x, oBaseTexture.y);
        this.bDrawTexture = false;
        return this.cachedCanvas;
    };
    CSeqList.prototype.clearCachedTexture = function() {
        if(this.cachedCanvas) {
            this.cachedCanvas = null;
        }
    };

    function CAnimSequence(oParentControl, oSeq) {//main seq, interactive seq
        CControlContainer.call(this, oParentControl);
        this.seq = oSeq;
        this.label = null; //this.addControl(new CLabel(this, "seq"));
        this.groupList = null;//this.addControl(new CAnimGroupList(this));
    }
    InitClass(CAnimSequence, CControlContainer, CONTROL_TYPE_ANIM_SEQ);
    CAnimSequence.prototype.getIndexLabelRight = function() {
        return this.parentControl.getIndexLabelRight() - this.getLeft();
    };
    CAnimSequence.prototype.recalculateChildren = function() {
        this.clear();
        var sLabel = this.seq.getLabel();
        if(typeof sLabel === "string" && sLabel.length > 0) {
            this.label = this.addControl(new CLabel(this, sLabel, 9, true));
        }
        else {
            this.label = null;
        }
        this.groupList = this.addControl(new CAnimGroupList(this));
    };
    CAnimSequence.prototype.getSeq = function() {
        return this.seq;
    };
    CAnimSequence.prototype.recalculateChildrenLayout = function() {
        var dCurY = 0;
        if(this.label) {
            this.label.setLayout(0, dCurY, this.getWidth(), SEQ_LABEL_HEIGHT);
            this.label.recalculate();
            dCurY += this.label.getHeight();
        }
        if(this.groupList) {
            this.groupList.setLayout(0, dCurY, this.getWidth(), 0);
            this.groupList.recalculate();
            dCurY += this.groupList.getHeight();
        }
        this.setLayout(this.getLeft(), this.getTop(), this.getWidth(), dCurY);
    };
    CAnimSequence.prototype.getFillColor = function() {
        return null;
    };
    CAnimSequence.prototype.getOutlineColor = function() {
        return null;
    };

    function CAnimGroupList(oParentControl) {//main seq, interactive seq
        CControlContainer.call(this, oParentControl);
    }
    InitClass(CAnimGroupList, CControlContainer, CONTROL_TYPE_ANIM_GROUP_LIST);
    CAnimGroupList.prototype.getIndexLabelRight = function() {
        return this.parentControl.getIndexLabelRight() - this.getLeft();
    };
    CAnimGroupList.prototype.getSeq = function() {
        return this.parentControl.getSeq();
    };
    CAnimGroupList.prototype.recalculateChildren = function() {
        this.clear();
        var oSeq = this.getSeq();
        var aAllEffects = oSeq.getAllEffects();

        for(var nCurEffect = 0; nCurEffect < aAllEffects.length; ++nCurEffect) {
            var oItem = new CAnimItem(this, aAllEffects[nCurEffect]);
            this.addControl(oItem);
        }
    };
    CAnimGroupList.prototype.getFillColor = function() {
        return null;
    };
    CAnimGroupList.prototype.getOutlineColor = function() {
        return null;
    };

    //CAnimGroupList.prototype.draw = function() {
    //};
    CAnimGroupList.prototype.recalculateChildrenLayout = function() {
        var dLastBottom = 0;
        for(var nChild = 0; nChild < this.children.length; ++nChild) {
            var oChild = this.children[nChild];
            oChild.setLayout(0, dLastBottom, this.getWidth(), ANIM_ITEM_HEIGHT);
            oChild.recalculate();
            dLastBottom = oChild.getBottom();
        }
        this.setLayout(this.getLeft(), this.getTop(), this.getWidth(), dLastBottom);
    };

    function CAnimGroup(oParentControl, aEffects) {
        CControlContainer.call(this, oParentControl);
    }
    InitClass(CAnimGroup, CControlContainer, CONTROL_TYPE_ANIM_GROUP);
    CAnimGroup.prototype.getIndexLabelRight = function() {
        return this.parentControl.getIndexLabelRight() - this.getLeft();
    };

    function CImageControl(oParentControl) {
        CControl.call(this, oParentControl)
    }
    InitClass(CImageControl, CControl, CONTROL_TYPE_IMAGE);
    CImageControl.prototype.canHandleEvents = function () {
        return false;
    };
    //CImageControl.prototype.draw = function() {
    //};

    function CEffectBar(oParentControl) {
        CControl.call(this, oParentControl)
    }
    InitClass(CEffectBar, CControl, CONTROL_TYPE_EFFECT_BAR);

    function CAnimItem(oParentControl, oEffect) {
        CControlContainer.call(this, oParentControl);
        this.indexLabel = this.addControl(new CLabel(this, "1.", 7.5));
        this.eventTypeImage = this.addControl(new CImageControl(this));
        this.effectTypeImage = this.addControl(new CImageControl(this));
        this.effectLabel = this.addControl(new CLabel(this, oEffect.getObjectName(), 7.5));
        this.effectBar = this.addControl(new CEffectBar(this));
        this.contextMenuButton = this.addControl(new CButton(this));

        this.effect = oEffect;
    }
    InitClass(CAnimItem, CControlContainer, CONTROL_TYPE_ANIM_ITEM);
    CAnimItem.prototype.getIndexLabelRight = function() {
        return this.parentControl.getIndexLabelRight() - this.getLeft();
    };
    CAnimItem.prototype.getEffectLabelRight = function() {
        return LABEL_TIMELINE_WIDTH;
    };
    CAnimItem.prototype.recalculateChildrenLayout = function() {
        var dIndexLabelRight = this.getIndexLabelRight();
        var dYInside = (this.getHeight() - EFFECT_BAR_HEIGHT) / 2;
        this.indexLabel.setLayout(0, dYInside, dIndexLabelRight, EFFECT_BAR_HEIGHT);
        this.eventTypeImage.setLayout(this.indexLabel.getRight(), dYInside, EFFECT_BAR_HEIGHT, EFFECT_BAR_HEIGHT);
        this.effectTypeImage.setLayout(this.eventTypeImage.getRight(), dYInside, EFFECT_BAR_HEIGHT, EFFECT_BAR_HEIGHT);
        var dLabelRight = this.getEffectLabelRight();
        var dEffectLabelLeft = this.effectTypeImage.getRight();
        this.effectLabel.setLayout(dEffectLabelLeft, dYInside, dLabelRight - dEffectLabelLeft, EFFECT_BAR_HEIGHT);
        this.effectBar.setLayout(0, 0, 0, 0);//todo
        var dRightSpace = dYInside;
        this.contextMenuButton.setLayout(this.getRight() - dRightSpace - EFFECT_BAR_HEIGHT, dYInside, EFFECT_BAR_HEIGHT, EFFECT_BAR_HEIGHT);
    };
    CAnimItem.prototype.canHandleEvents = function () {
        return true;
    };
    // CAnimItem.prototype.getFillColor = function() {
    //     return null;
    // };
    CAnimItem.prototype.getOutlineColor = function() {
        return null;
    };

    //CAnimItem.prototype.draw = function() {
    //};

    function CLabel(oParentControl, sString, nFontSize, bBold, nParaAlign) {
        CControl.call(this, oParentControl);
        AscFormat.ExecuteNoHistory(function(){
            this.string = sString;
            this.fontSize = nFontSize;
            this.createTextBody();
            var oTxLstStyle = new AscFormat.TextListStyle();
            oTxLstStyle.levels[0] =  new CParaPr();
            oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
            oTxLstStyle.levels[0].DefaultRunPr.FontSize = nFontSize;
            oTxLstStyle.levels[0].DefaultRunPr.Bold = bBold;
            oTxLstStyle.levels[0].DefaultRunPr.Color = new AscCommonWord.CDocumentColor(0x44, 0x44, 0x44, false);
            oTxLstStyle.levels[0].DefaultRunPr.RFonts.SetAll("Arial", -1);
            if(AscFormat.isRealNumber(nParaAlign)) {
                oTxLstStyle.levels[0].Jc = nParaAlign;
            }
            this.txBody.setLstStyle(oTxLstStyle);
            this.bodyPr = new AscFormat.CBodyPr();
            this.bodyPr.setDefault();
            this.bodyPr.anchor = 1;//vertical align ctr
            this.bodyPr.lIns = 0;
            this.bodyPr.rIns = 0;
            this.bodyPr.tIns = 0;
            this.bodyPr.bIns = 0;
            this.bodyPr.horzOverflow = AscFormat.nOTClip;
            this.bodyPr.vertOverflow = AscFormat.nOTClip;
        }, this, []);
    }
    InitClass(CLabel, CControl, CONTROL_TYPE_LABEL);
    CLabel.prototype.getString = function() {
        return AscCommon.translateManager.getValue(this.string);
    };
    CLabel.prototype.recalculateContent = function () {
        //this.recalculateGeometry();
        this.recalculateTransform();
//        this.txBody.content.Recalc_AllParagraphs_CompiledPr();
        if(!this.txBody.bFit || !AscFormat.isRealNumber(this.txBody.fitWidth) || this.txBody.fitWidth > this.getWidth()) {
            this.txBody.recalculateOneString(this.getString());
        }
    };
    CLabel.prototype.canHandleEvents = function () {
        return false;
    };
    CLabel.prototype.getFillColor = function() {
        return null;
    };
    CLabel.prototype.getOutlineColor = function() {
        return null;
    };
    CLabel.prototype.recalculateTransformText = function() {
        var Y = this.getHeight() / 2 - this.txBody.content.GetSummaryHeight() / 2;
        if(!this.transformText) {
            this.transformText = new AscCommon.CMatrix();
        }
        this.transformText.tx = this.transform.tx;
        this.transformText.ty = this.transform.ty + Y;

        if(!this.invertTransformText) {
            this.invertTransformText = new AscCommon.CMatrix();
        }
        this.invertTransformText.tx = -this.transformText.tx;
        this.invertTransformText.ty = -this.transformText.ty;
        this.localTransformText = this.transformText;
    };
    CLabel.prototype.recalculateTransformText2 = function() {
        return null;
    };

    function CButton(oParentControl, fOnMouseDown, fOnMouseMove, fOnMouseUp) {
        CControlContainer.call(this, oParentControl);
        this.onMouseDownCallback = fOnMouseDown;
        this.onMouseMoveCallback = fOnMouseMove;
        this.onMouseUpCallback = fOnMouseUp;
    }
    InitClass(CButton, CControlContainer, CONTROL_TYPE_BUTTON);
    CButton.prototype.onMouseDown = function(e, x, y) {
        if(this.onMouseDownCallback && this.onMouseDownCallback.call(this, e, x, y)) {
            return true;
        }
        return CControlContainer.prototype.onMouseDown.call(this, e, x, y);
    };
    CButton.prototype.onMouseMove = function(e, x, y) {
        if(this.onMouseMoveCallback && this.onMouseMoveCallback.call(this, e, x, y)) {
            return true;
        }
        return CControlContainer.prototype.onMouseMove.call(this, e, x, y);
    };
    CButton.prototype.onMouseUp = function(e, x, y) {
        if(this.onMouseUpCallback && this.onMouseUpCallback.call(this, e, x, y)) {
            return true;
        }
        return CControlContainer.prototype.onMouseUp.call(this, e, x, y);
    };
    CButton.prototype.canHandleEvents = function () {
        return true;
    };
    CButton.prototype.canHandleEvents = function () {
        return true;
    };
    // CButton.prototype.draw = function(graphics) {
    //     if(this.isHidden()){
    //         return false;
    //     }
    //     if(!this.checkUpdateRect(graphics.updatedRect)) {
    //         return false;
    //     }
    //
    //     graphics.SaveGrState();
    //     var oSkin = AscCommon.GlobalSkin;
    //     //ScrollBackgroundColor     : "#EEEEEE",
    //     //ScrollOutlineColor        : "#CBCBCB",
    //     //ScrollOutlineHoverColor   : "#CBCBCB",
    //     //ScrollOutlineActiveColor  : "#ADADAD",
    //     //ScrollerColor             : "#F7F7F7",
    //     //ScrollerHoverColor        : "#C0C0C0",
    //     //ScrollerActiveColor       : "#ADADAD",
    //     //ScrollArrowColor          : "#ADADAD",
    //     //ScrollArrowHoverColor     : "#F7F7F7",
    //     //ScrollArrowActiveColor    : "#F7F7F7",
    //     //ScrollerTargetColor       : "#CFCFCF",
    //     //ScrollerTargetHoverColor  : "#F1F1F1",
    //     //ScrollerTargetActiveColor : "#F1F1F1",
    //     var x = 0;
    //     var y = 0;
    //     var extX = this.getWidth();
    //     var extY = this.getHeight();
    //     graphics.transform3(this.transform);
    //
    //     var sFillColor;
    //     var sOutlineColor;
    //     var oColor;
    //     if(this.isActive()) {
    //         sFillColor = oSkin.ScrollerActiveColor;
    //         sOutlineColor = oSkin.ScrollOutlineActiveColor;
    //     }
    //     else if(this.isHovered()) {
    //         sFillColor = oSkin.ScrollerHoverColor;
    //         sOutlineColor = oSkin.ScrollOutlineHoverColor;
    //     }
    //     else {
    //         sFillColor = oSkin.ScrollerColor;
    //         sOutlineColor = oSkin.ScrollOutlineColor;
    //     }
    //     oColor = AscCommon.RgbaHexToRGBA(sFillColor);
    //     graphics.b_color1(oColor.R, oColor.G, oColor.B, 0xFF);
    //     graphics.rect(x, y, extX, extY);
    //     graphics.df();
    //     oColor = AscCommon.RgbaHexToRGBA(sOutlineColor);
    //
    //     graphics.SetIntegerGrid(true);
    //     graphics.p_width(0);
    //     graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
    //     graphics.drawHorLine(0, y, x, x + extX, 0);
    //     graphics.drawHorLine(0, y + extY, x, x + extX, 0);
    //     graphics.drawVerLine(2, x, y, y + extY, 0);
    //     graphics.drawVerLine(2, x + extX, y, y + extY, 0);
    //     graphics.ds();
    //     graphics.RestoreGrState();
    //     return true;
    // };

    CButton.prototype.getFillColor = function() {
        // if(this.parentControl instanceof CTimelineContainer) {
        //     return null;
        // }
        var oSkin = AscCommon.GlobalSkin;
        if(this.isActive()) {
            return oSkin.ScrollerActiveColor;
        }
        else if(this.isHovered()) {
            return oSkin.ScrollerHoverColor;
        }
        else {
            return oSkin.ScrollerColor;
        }
    };
    CButton.prototype.getOutlineColor = function() {
        // if(this.parentControl instanceof CTimelineContainer) {
        //     return null;
        // }
        var oSkin = AscCommon.GlobalSkin;
        if(this.isActive()) {
            return oSkin.ScrollOutlineActiveColor;
        }
        else if(this.isHovered()) {
            return oSkin.ScrollOutlineHoverColor;
        }
        else {
            return oSkin.ScrollOutlineColor;
        }
    };

    var PLAY_BUTTON_WIDTH = 82 * AscCommon.g_dKoef_pix_to_mm;
    var PLAY_BUTTON_HEIGHT = 24 * AscCommon.g_dKoef_pix_to_mm;
    var PLAY_BUTTON_LEFT = 145 * AscCommon.g_dKoef_pix_to_mm;
    var PLAY_BUTTON_TOP = 12 * AscCommon.g_dKoef_pix_to_mm;
    function CAnimPaneHeader(oDrawer) {
        CTopControl.call(this, oDrawer);
        this.label = this.addControl(new CLabel(this, "Animation Pane", 10, true));
        this.playButton = this.addControl(new CButton(this));
        this.closeButton = this.addControl(new CButton(this));
    }
    InitClass(CAnimPaneHeader, CTopControl, CONTROL_TYPE_HEADER);
    CAnimPaneHeader.prototype.recalculateChildrenLayout = function() {
        this.closeButton.setLayout(
            this.getWidth() - AscCommon.TIMELINE_LIST_RIGHT_MARGIN - BUTTON_SIZE,
            (this.getHeight() - BUTTON_SIZE) / 2,
            BUTTON_SIZE,
            BUTTON_SIZE
        );
        this.playButton.setLayout(
            PLAY_BUTTON_LEFT,
            PLAY_BUTTON_TOP,
            PLAY_BUTTON_WIDTH,
            PLAY_BUTTON_HEIGHT
        );
        this.label.setLayout(
            AscCommon.TIMELINE_LEFT_MARGIN,
            0,
            this.playButton.getLeft(),
            this.getHeight()
        );

    };
    CAnimPaneHeader.prototype.getFillColor = function() {
        return null;
    };
    CAnimPaneHeader.prototype.getOutlineColor = function() {
        return null;
    };

    function CToolbar(oParentControl) {
        CControlContainer.call(this, oParentControl);
        this.playButton = this.addControl(new CButton(this));
        this.upButton = this.addControl(new CButton(this));
        this.downButton = this.addControl(new CButton(this));
    }
    InitClass(CToolbar, CControlContainer, CONTROL_TYPE_TOOLBAR);
    CToolbar.prototype.recalculateChildrenLayout = function() {
        this.playButton.setLayout(0, 0, this.getWidth(), BUTTON_SIZE);
        this.downButton.setLayout(
            this.getWidth() - BUTTON_SIZE,
            this.getHeight() - BUTTON_SIZE,
            BUTTON_SIZE,
            BUTTON_SIZE
        );
        this.upButton.setLayout(
            this.downButton.getLeft() - BUTTON_SPACE - BUTTON_SIZE,
            this.getHeight() - BUTTON_SIZE,
            BUTTON_SIZE,
            BUTTON_SIZE
        );
    };
    CToolbar.prototype.getFillColor = function() {
        return null;
    };
    CToolbar.prototype.getOutlineColor = function() {
        return null;
    };



    var SECONDS_BUTTON_WIDTH = 76 * AscCommon.g_dKoef_pix_to_mm;
    var SECONDS_BUTTON_HEIGHT = 24 * AscCommon.g_dKoef_pix_to_mm;
    var SECONDS_BUTTON_LEFT = 57 * AscCommon.g_dKoef_pix_to_mm;
    function CTimelineContainer(oDrawer) {
        CTopControl.call(this, oDrawer);
        this.drawer = oDrawer;
        this.secondsButton = this.addControl(new CButton(this));
        this.timeline = this.addControl(new CTimeline(this));
    }
    InitClass(CTimelineContainer, CTopControl, CONTROL_TYPE_TIMELINE_CONTAINER);
    CTimelineContainer.prototype.recalculateChildrenLayout = function() {
        var dPosY = (this.getHeight() - SECONDS_BUTTON_HEIGHT) / 2;
        this.secondsButton.setLayout(SECONDS_BUTTON_LEFT, dPosY, SECONDS_BUTTON_WIDTH, SECONDS_BUTTON_HEIGHT);
        var dLeft = LABEL_TIMELINE_WIDTH + AscCommon.TIMELINE_LEFT_MARGIN - 1.5*SCROLL_THICKNESS;
        var dWidth = this.getWidth() - AscCommon.TIMELINE_LIST_RIGHT_MARGIN - dLeft;
        dPosY = (this.getHeight() - SCROLL_THICKNESS) / 2;
        this.timeline.setLayout(dLeft, dPosY, dWidth, SCROLL_THICKNESS);
    };
    CTimelineContainer.prototype.getFillColor = function() {
        return null;
    };
    CTimelineContainer.prototype.getOutlineColor = function() {
        return null;
    };

    //Time scales in seconds
    var TIME_SCALES = [
        1,
        1,
        2,
        5,
        10,
        20,
        60,
        120,
        300,
        600,
        600
    ];
    //lengths
    var SMALL_TIME_INTERVAL = 15;
    var MIDDLE_1_TIME_INTERVAL = 20;
    var MIDDLE_2_TIME_INTERVAL = 25;
    var LONG_TIME_INTERVAL = 30;
    var TIME_INTERVALS = [
        LONG_TIME_INTERVAL, //1
        SMALL_TIME_INTERVAL, //1
        SMALL_TIME_INTERVAL, //2
        MIDDLE_1_TIME_INTERVAL, //5
        MIDDLE_1_TIME_INTERVAL,//10
        MIDDLE_1_TIME_INTERVAL,//20
        MIDDLE_2_TIME_INTERVAL,//60
        MIDDLE_2_TIME_INTERVAL,//120
        MIDDLE_2_TIME_INTERVAL,//300
        MIDDLE_2_TIME_INTERVAL,//600
        SMALL_TIME_INTERVAL//600
    ];
    var LABEL_WIDTH = 100;
    function CTimeline(oParentControl) {
        CScrollHor.call(this, oParentControl);
        this.startTimePos = 0;
        this.curTimePos = 0;
        this.tmScaleIdx = 5;

        //labels cache
        this.labels = {};
        this.usedLabels = {};

        this.cachedParaPr = null;
    }
    InitClass(CTimeline, CScrollHor, CONTROL_TYPE_TIMELINE);
    CTimeline.prototype.startDrawLabels = function() {
        this.usedLabels = {};
    };
    CTimeline.prototype.endDrawLabels = function() {
        for(var nTime in this.labels) {
            if(!this.usedLabels[nTime]) {
                var oLabel = this.labels[nTime];
                oLabel.parentControl = null;
                oLabel.bDeleted = true;
                delete this.labels[nTime];
            }
        }
    };
    CTimeline.prototype.getLabel = function(nTime, scale) {
      this.usedLabels[nTime] = true;
      if(this.labels[nTime] && AscFormat.fApproxEqual(this.labels[nTime].scale, scale, 0.01)) {
          return this.labels[nTime];
      }
      return this.cacheLabel(nTime, scale);
    };
    CTimeline.prototype.cacheLabel = function(nTime, scale) {
      var oLabel = new CLabel(this, this.getTimeString(nTime), 7.5);
      var oContent = oLabel.txBody.content;
      oLabel.setLayout(0, 0, LABEL_WIDTH, this.getHeight());
      if(this.cachedParaPr) {
          oContent.Content[0].CompiledPr = this.cachedParaPr;
      }
      else {
          oContent.SetApplyToAll(true);
          oContent.SetParagraphAlign(AscCommon.align_Center);
          oContent.SetApplyToAll(false);
      }
      oLabel.recalculate();
      if(!this.cachedParaPr) {
          this.cachedParaPr = oContent.Content[0].CompiledPr;
      }
      var oBaseTexture = oLabel.getAnimTexture(scale);
      this.labels[nTime] = new CAnimTexture(this, oBaseTexture.canvas, oBaseTexture.scale, oBaseTexture.x, oBaseTexture.y);
      return this.labels[nTime];
    };
    CTimeline.prototype.getTimeString = function(nTime) {
      if(nTime < 60) {
          return "" + nTime;
      }
      var nMin, nSec;
      var sMin, sSec;
      nSec = (nTime % 60);
      if(nSec === 0) {
          sSec = "00";
      }
      else {
          sSec = "" + nSec;
      }
      if(nTime < 3600) {
          return (((nTime / 60) >> 0)  + ":") + sSec;
      }

        nMin = ((nTime / 60) >> 0);
        if(nMin === 0) {
            sMin = "00";
        }
        else {
            sMin = "" + nMin;
        }
      return (((nTime / 3600) >> 0)  + ":") + (sMin  + ":") + sSec;
    };
    CTimeline.prototype.drawLabel = function(graphics, dPos, nTime) {
        var oLabelTexture = this.getLabel(nTime, graphics.m_oCoordTransform.sx);
        var oMatrix = new AscCommon.CMatrix();
        var dWidth = oLabelTexture.canvas.width / oLabelTexture.scale;
        var dHeight = oLabelTexture.canvas.height / oLabelTexture.scale ;
        graphics.drawImage2(oLabelTexture.canvas,
            dPos - dWidth / 2, this.getHeight() / 2 - dHeight / 2,
            dWidth,
            dHeight);
        // var oContent = oLabel.txBody.content;
        // oContent.ShiftView(dPos - LABEL_WIDTH / 2, this.getHeight() / 2 - oContent.GetSummaryHeight() / 2);
        // oContent.Draw(0, graphics);
        // oContent.ResetShiftView();
    };
    CTimeline.prototype.getPaneLeft = function() {
        return SCROLL_BUTTON_SIZE;
    };
    CTimeline.prototype.getFillColor = function() {
        return null;
    };
    CTimeline.prototype.getOutlineColor = function() {
        return null;
    };
    CTimeline.prototype.recalculateChildrenLayout = function() {
        this.children[0].setLayout(0, 0, SCROLL_BUTTON_SIZE, SCROLL_BUTTON_SIZE);
        this.children[1].setLayout(this.getWidth() - SCROLL_BUTTON_SIZE, 0, SCROLL_BUTTON_SIZE, SCROLL_BUTTON_SIZE);
    };
    CTimeline.prototype.canHandleEvents = function() {
        return true;
    };
    CTimeline.prototype.drawMark = function(graphics, dPos) {
        var dHeight = this.getHeight() / 3;
        var nPenW = this.getPenWidth(graphics);
        graphics.drawVerLine(1, dPos, dHeight, dHeight + dHeight, nPenW);
    };
    CTimeline.prototype.start = function(graphics, dPos) {
        var dHeight = this.getHeight() / 3;
        var nPenW = this.getPenWidth(graphics);
        graphics.drawVerLine(1, dPos, dHeight, dHeight + dHeight, nPenW);
    };
    CTimeline.prototype.draw = function(graphics) {
        if(!CScrollHor.prototype.draw.call(this, graphics)) {
            return false;
        }
        graphics.SaveGrState();
        // var dPenW = this.getPenWidth(graphics);
        // graphics.SetIntegerGrid(true);
        // graphics.p_width(dPenW);
        // var sColor = this.children[0].getOutlineColor();
        // var oColor = AscCommon.RgbaHexToRGBA(sColor);
        // graphics.p_color(oColor.R, oColor.G, oColor.B, 255);
        // var dPaneLeft = this.children[0].getRight();
        // var dPaneWidth = this.getWidth() - (this.children[0].getWidth() + this.children[1].getWidth());
        // graphics.rect(dPaneLeft, 0, dPaneWidth, this.getHeight());
        // graphics.ds();
        // graphics.RestoreGrState();
        var sColor = this.children[0].getOutlineColor();
        var oColor = AscCommon.RgbaHexToRGBA(sColor);
        var dPaneLeft = this.getRulerStart();
        var dPaneWidth = this.getRulerEnd() - dPaneLeft;
        var x = dPaneLeft;
        var y = 0;
        var extX = dPaneWidth;
        var extY = this.getHeight();
        graphics.transform3(this.transform);
        graphics.SetIntegerGrid(true);
        var nPenW = this.getPenWidth(graphics);
        graphics.p_color(oColor.R, oColor.G, oColor.B, 0xFF);
        graphics.drawHorLine(0, y, x, x + extX, nPenW);
        graphics.drawHorLine(0, y + extY, x, x + extX, nPenW);
        graphics.drawVerLine(2, x, y, y + extY, nPenW);
        graphics.drawVerLine(2, x + extX, y, y + extY, nPenW);
        graphics.ds();

        //draw marks
        //find first visible
        var fStartTime = this.posToTime(this.getRulerStart());
        var fTimeInterval = TIME_SCALES[this.tmScaleIdx];
        var nMarksCount = TIME_INTERVALS[this.tmScaleIdx] === LONG_TIME_INTERVAL ? 10 : 2;

        var dTimeOfSmallInterval = fTimeInterval / nMarksCount;
        var nStartIntervalIdx = this.startTimePos/dTimeOfSmallInterval >> 0;
        var nEndIntervalIdx = this.posToTime(this.getRulerEnd()) / dTimeOfSmallInterval + 0.5 >> 0;
        this.startDrawLabels();
        var nInterval;
        graphics.AddClipRect(x, y, extX, extY);
        for(nInterval = nStartIntervalIdx; nInterval <= nEndIntervalIdx; ++nInterval) {
            var dTime = nInterval*dTimeOfSmallInterval;
            var dPos = this.timeToPos(dTime);
            if(nInterval % nMarksCount !== 0) {
                this.drawMark(graphics, dPos);
            }
            else {
                this.drawLabel(graphics, dPos, dTime);
            }
        }
        graphics.ds();
        // for(nInterval = nFirstInterval; nInterval <= nLastInterval; ++nInterval) {
        //     var dTime = nInterval*dSmallInterval;
        //     var dPos = this.timeToPos(dTime);
        //     if(nInterval % nMarksCount === 0) {
        //         this.drawLabel(graphics, dPos, dTime);
        //     }
        // }
        this.endDrawLabels();
        //

        graphics.RestoreGrState();
    };
    CTimeline.prototype.getRulerStart = function() {
        return this.children[0].getRight();
    };
    CTimeline.prototype.getRulerEnd = function() {
        return this.getWidth() - this.children[1].getWidth();
    };
    CTimeline.prototype.getCursorSize = function() {
        return BUTTON_SIZE;
    };
    CTimeline.prototype.getZeroShift = function() {
        return this.getRulerStart() + this.getCursorSize() / 2;
    };
    CTimeline.prototype.timeToPos = function(fTime) {
        //linear relationship x = a*t + b
        var oCoefs = this.getLinearCoeffs();
        return oCoefs.a*fTime + oCoefs.b;
    };
    CTimeline.prototype.getLinearCoeffs = function() {
        //linear relationship x = a*t + b
        var a = TIME_INTERVALS[this.tmScaleIdx]/TIME_SCALES[this.tmScaleIdx];
        var b = this.getZeroShift() - a*this.startTimePos;
        return {a: a, b: b};
    };
    CTimeline.prototype.posToTime = function(fPos) {
        //linear relationship x = a*t + b
        var oCoefs = this.getLinearCoeffs();
        return (fPos - oCoefs.b) / oCoefs.a;
    };

    var HEADER_HEIGHT = 7.5;
    var BUTTON_SIZE = HEADER_HEIGHT;
    var TOOLBAR_HEIGHT = HEADER_HEIGHT;
    var PADDING_LEFT = 3;
    var PADDING_TOP = PADDING_LEFT;
    var PADDING_RIGHT = PADDING_LEFT;
    var PADDING_BOTTOM = PADDING_LEFT;
    var VERTICAL_SPACE = PADDING_LEFT;
    var HORIZONTAL_SPACE = PADDING_LEFT;
    var SCROLL_THICKNESS = 15 * AscCommon.g_dKoef_pix_to_mm;
    var SCROLL_BUTTON_SIZE = SCROLL_THICKNESS;
    var TIMELINE_HEIGHT = SCROLL_THICKNESS + 1;
    var BUTTON_SPACE = HORIZONTAL_SPACE / 2;
    var TOOLBAR_WIDTH = 25;
    var ANIM_LABEL_WIDTH = 40;
    var ANIM_ITEM_HEIGHT = TIMELINE_HEIGHT;
    var EFFECT_BAR_HEIGHT = 2*ANIM_ITEM_HEIGHT/3;
    var SEQ_LABEL_HEIGHT = EFFECT_BAR_HEIGHT;


    function CAnimPane(oTiming) {
        CControlContainer.call(this, null);
        this.timing = oTiming;
        this.header = this.addControl(new CAnimPaneHeader(this));
        this.toolbar = this.addControl(new CToolbar(this));
        this.seqListContainer = this.addControl(new CSeqListContainer(this));
        this.timelineContainer = this.addControl(new CTimelineContainer(this));

        this.recalcInfo.recalculateHeader = true;
        this.recalcInfo.recalculateToolbar = true;
        this.recalcInfo.recalculateSeqListContainer = true;
        this.recalcInfo.recalculateTimelineContainer = true;
    }
    InitClass(CAnimPane, CControlContainer, CONTROL_TYPE_UNKNOWN);
    CAnimPane.prototype.getHeader = function() {
        return this.getChildByType(CONTROL_TYPE_HEADER);
    };
    CAnimPane.prototype.getToolbar = function() {
        return this.getChildByType(CONTROL_TYPE_TOOLBAR);
    };
    CAnimPane.prototype.getSeqListContainer = function() {
        return this.getChildByType(CONTROL_TYPE_SEQ_LIST_CONTAINER);
    };
    CAnimPane.prototype.getTimelineContainer = function() {
        return this.getChildByType(CONTROL_TYPE_TIMELINE_CONTAINER);
    };
    CAnimPane.prototype.onChanged = function(oRect) {
        this.timing.onAnimPaneChanged(oRect);
    };
    CAnimPane.prototype.onResize = function() {
        return;
        this.setLayout(
            0,
            0,
            this.getExternalControlWidth(),
            this.getExternalControlHeight()
        );
        this.recalculate();
        this.onUpdate();
    };
    CAnimPane.prototype.getExternalControl = function() {
        return editor.WordControl.m_oAnimPaneApi;
    };
    CAnimPane.prototype.getExternalControlWidth = function() {
        return this.getExternalControl().GetWidth();
    };
    CAnimPane.prototype.getExternalControlHeight = function() {
        return this.getExternalControl().GetHeight();
    };
    CAnimPane.prototype.onChildUpdate = function(oBounds) {
        this.getExternalControl().OnAnimPaneChanged(this.getSlideNum(), oBounds);
    };
    CAnimPane.prototype.onUpdate = function() {
        this.getExternalControl().OnAnimPaneChanged(this.getSlideNum(), this.getBounds());
    };
    CAnimPane.prototype.getSlideNum = function(oBounds) {
        return this.timing.parent.num;
    };
    CAnimPane.prototype.recalculateChildrenLayout = function() {
        var dControlWidth = Math.max(0, this.getWidth() - PADDING_LEFT - PADDING_RIGHT);
        this.header.setLayout(
            PADDING_LEFT,
            PADDING_TOP,
            this.getWidth() - PADDING_LEFT - PADDING_RIGHT,
            HEADER_HEIGHT
        );
        var dBottomPartY = this.header.getBottom() + VERTICAL_SPACE;
        this.toolbar.setLayout(
            PADDING_LEFT,
            dBottomPartY,
            TOOLBAR_WIDTH,
            this.getHeight() - dBottomPartY - PADDING_BOTTOM
        );

        var dRightPartX = PADDING_LEFT + this.toolbar.getWidth() + VERTICAL_SPACE;
        var dRightPartWidth = Math.max(0, this.getWidth() - dRightPartX - PADDING_RIGHT);
        this.timelineContainer.setLayout(
            dRightPartX,
            this.getHeight() - PADDING_BOTTOM - TIMELINE_HEIGHT,
            dRightPartWidth,
            TIMELINE_HEIGHT
        );
        var dListTop = dBottomPartY;
        var dListBottom = this.timelineContainer.getTop();
        var dListLeft = dRightPartX;
        this.seqListContainer.setLayout(
            dListLeft,
            dListTop,
            dRightPartWidth,
            Math.max(0, dListBottom - dListTop)
        );
    };
    CAnimPane.prototype.recalculateHeader = function() {
        this.header.recalculate();
    };
    CAnimPane.prototype.recalculateToolbar = function() {
        this.toolbar.recalculate();
    };
    CAnimPane.prototype.recalculateSeqListContainer = function() {
        this.seqListContainer.recalculate();
    };
    CAnimPane.prototype.recalculateTimelineContainer = function() {
        this.timelineContainer.recalculate();
    };
    
    CAnimPane.prototype.recalculate = function() {
        return;
    };
    //CAnimPane.prototype.draw = function(oGraphics) {
    //    oGraphics.b_color1(255, 0, 0, 255);
    //    oGraphics.rect(0, 0, 100, 100);
    //    oGraphics.df();
    //};


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
    window['AscFormat'].CAnimationTimer = CAnimationTimer;
    window['AscFormat'].CAnimationPlayer = CAnimationPlayer;
    window['AscFormat'].CBaseAnimObject = CBaseAnimObject;
    window['AscFormat'].CAnimFormulaParser = CFormulaParser;
    window['AscFormat'].CBaseAnimTexture = CBaseAnimTexture;
    window['AscFormat'].CDemoAnimPlayer = CDemoAnimPlayer;
    window['AscFormat'].ICON_TRIGGER = ICON_TRIGGER;
    


    function generate_preset_data() {
        var aPresets = "emph,emphasis_blink,35,0;emph,emphasis_bold_flash,10,0;emph,emphasis_bold_reveal,15,0;emph,emphasis_brush_color,16,0;emph,emphasis_color_pulse,27,0;emph,emphasis_complementary_color,21,0;emph,emphasis_complementary_color_2,22,0;emph,emphasis_contrasting_color,23,0;emph,emphasis_contrasting_darken,24,0;emph,emphasis_desaturate,25,0;emph,emphasis_fill_color,1,2;emph,emphasis_font_color,3,2;emph,emphasis_grow_shrink,6,0;emph,emphasis_grow_with_color,28,0;emph,emphasis_lighten,30,0;emph,emphasis_line_color,7,2;emph,emphasis_object_color,19,0;emph,emphasis_pulse,26,0;emph,emphasis_shimmer,36,0;emph,emphasis_spin,8,0;emph,emphasis_teeter,32,0;emph,emphasis_transparency,9,0;emph,emphasis_underline,18,0;emph,emphasis_wave,34,0;entr,entrance_appear,1,0;entr,entrance_basic_swivel_horizontal,19,10;entr,entrance_basic_swivel_vertical,19,5;entr,entrance_basic_zoom_in,23,16;entr,entrance_basic_zoom_in_from_screen_center,23,528;entr,entrance_basic_zoom_in_slightly,23,272;entr,entrance_basic_zoom_out,23,32;entr,entrance_basic_zoom_out_from_screen_bottom,23,36;entr,entrance_basic_zoom_out_slightly,23,288;entr,entrance_blinds_horizontal,3,10;entr,entrance_blinds_vertical,3,5;entr,entrance_boomerang,25,0;entr,entrance_bounce,26,0;entr,entrance_box_in,4,16;entr,entrance_box_out,4,32;entr,entrance_center_compress,50,0;entr,entrance_center_revolve,43,0;entr,entrance_checkerboard_across,5,10;entr,entrance_checkerboard_down,5,5;entr,entrance_circle_in,6,16;entr,entrance_circle_out,6,32;entr,entrance_credits,28,0;entr,entrance_curve_up,52,0;entr,entrance_diamond_in,8,16;entr,entrance_diamond_out,8,32;entr,entrance_dissolve_in,9,0;entr,entrance_drop,38,0;entr,entrance_expand,55,0;entr,entrance_fade,10,0;entr,entrance_flip,56,0;entr,entrance_float,30,0;entr,entrance_float_down,47,0;entr,entrance_float_up,42,0;entr,entrance_fly_in_from_bottom,2,4;entr,entrance_fly_in_from_bottom_left,2,12;entr,entrance_fly_in_from_bottom_right,2,6;entr,entrance_fly_in_from_left,2,8;entr,entrance_fly_in_from_right,2,2;entr,entrance_fly_in_from_top,2,1;entr,entrance_fly_in_from_top_left,2,9;entr,entrance_fly_in_from_top_right,2,3;entr,entrance_grow_and_turn,31,0;entr,entrance_peek_in_from_bottom,12,4;entr,entrance_peek_in_from_left,12,8;entr,entrance_peek_in_from_right,12,2;entr,entrance_peek_in_from_top,12,1;entr,entrance_pinwheel,35,0;entr,entrance_plus_in,13,16;entr,entrance_plus_out,13,32;entr,entrance_random_bars_horizontal,14,10;entr,entrance_random_bars_vertical,14,5;entr,entrance_rise_up,37,0;entr,entrance_spinner,49,0;entr,entrance_spiral_in,15,0;entr,entrance_split_horizontal_in,16,26;entr,entrance_split_horizontal_out,16,42;entr,entrance_split_vertical_in,16,21;entr,entrance_split_vertical_out,16,37;entr,entrance_stretch_across,17,10;entr,entrance_stretch_from_bottom,17,4;entr,entrance_stretch_from_left,17,8;entr,entrance_stretch_from_right,17,2;entr,entrance_stretch_from_top,17,1;entr,entrance_strips_left_down,18,12;entr,entrance_strips_left_up,18,9;entr,entrance_strips_right_down,18,6;entr,entrance_strips_right_up,18,3;entr,entrance_swivel,45,0;entr,entrance_wedge,20,0;entr,entrance_wheel_1_spoke,21,1;entr,entrance_wheel_2_spokes,21,2;entr,entrance_wheel_3_spokes,21,3;entr,entrance_wheel_4_spokes,21,4;entr,entrance_wheel_8_spokes,21,8;entr,entrance_whip,41,0;entr,entrance_whipe_from_bottom,22,4;entr,entrance_whipe_from_left,22,8;entr,entrance_whipe_from_right,22,2;entr,entrance_whipe_from_top,22,1;entr,entrance_zoom_object_center,53,16;entr,entrance_zoom_slide_center,53,528;exit,exit_basic_swivel_horizontal,19,10;exit,exit_basic_swivel_vertical,19,5;exit,exit_basic_zoom_in,23,16;exit,exit_basic_zoom_in_slightly,23,272;exit,exit_basic_zoom_in_to_screen_bottom,23,20;exit,exit_basic_zoom_out,23,32;exit,exit_basic_zoom_out_slightly,23,288;exit,exit_basic_zoom_out_to_screen_center,23,544;exit,exit_blinds_horizontal,3,10;exit,exit_blinds_vertical,3,5;exit,exit_boomerang,25,0;exit,exit_bounce,26,0;exit,exit_box_in,4,16;exit,exit_box_out,4,32;exit,exit_center_revolve,43,0;exit,exit_checkerboard_across,5,10;exit,exit_checkerboard_up,5,5;exit,exit_circle_in,6,16;exit,exit_circle_out,6,32;exit,exit_collapse_across,17,10;exit,exit_collapse_to_bottom,17,4;exit,exit_collapse_to_left,17,8;exit,exit_collapse_to_right,17,2;exit,exit_collapse_to_top,17,1;exit,exit_contract,55,0;exit,exit_credits,28,0;exit,exit_curve_down,52,0;exit,exit_diamond_in,8,16;exit,exit_diamond_out,8,32;exit,exit_disappear,1,0;exit,exit_dissolve_out,9,0;exit,exit_drop,38,0;exit,exit_fade,10,0;exit,exit_flip,56,0;exit,exit_float,30,0;exit,exit_float_down,42,0;exit,exit_float_up,47,0;exit,exit_fly_out_to_bottom,2,4;exit,exit_fly_out_to_bottom_left,2,12;exit,exit_fly_out_to_bottom_right,2,6;exit,exit_fly_out_to_left,2,8;exit,exit_fly_out_to_right,2,2;exit,exit_fly_out_to_top,2,1;exit,exit_fly_out_to_top_left,2,9;exit,exit_fly_out_to_top_right,2,3;exit,exit_peek_out_to_bottom,12,4;exit,exit_peek_out_to_left,12,8;exit,exit_peek_out_to_right,12,2;exit,exit_peek_out_to_top,12,1;exit,exit_pinwheel,35,0;exit,exit_plus_in,13,16;exit,exit_plus_out,13,32;exit,exit_random_bars_horizontal,14,10;exit,exit_random_bars_vertical,14,5;exit,exit_shrink_and_turn,31,0;exit,exit_shrink_down,37,0;exit,exit_spinner,49,0;exit,exit_spiral_out,15,0;exit,exit_split_horizontal_in,16,26;exit,exit_split_horizontal_out,16,42;exit,exit_split_vertical_in,16,21;exit,exit_split_vertical_out,16,37;exit,exit_stretchy,50,0;exit,exit_strips_left_down,18,12;exit,exit_strips_left_up,18,9;exit,exit_strips_right_down,18,6;exit,exit_strips_right_up,18,3;exit,exit_swivel,45,0;exit,exit_wedge,20,0;exit,exit_wheel_1_spoke,21,1;exit,exit_wheel_2_spokes,21,2;exit,exit_wheel_3_spokes,21,3;exit,exit_wheel_4_spokes,21,4;exit,exit_wheel_8_spokes,21,8;exit,exit_whip,41,0;exit,exit_whipe_from_bottom,22,4;exit,exit_whipe_from_left,22,8;exit,exit_whipe_from_right,22,2;exit,exit_whipe_from_top,22,1;exit,exit_zoom_object_center,53,32;exit,exit_zoom_slide_center,53,544;path,motion_arc_down,37,0;path,motion_arc_left,51,0;path,motion_arc_right,58,0;path,motion_arc_up,44,0;path,motion_bean,31,0;path,motion_bounce_left,41,0;path,motion_bounce_right,54,0;path,motion_circle,1,0;path,motion_crescent_moon,6,0;path,motion_curved_square,20,0;path,motion_curved_x,21,0;path,motion_curvy_left,48,0;path,motion_curvy_right,61,0;path,motion_curvy_star,23,0;path,motion_custom_path,0,0;path,motion_decaying_wave,60,0;path,motion_diagonal_down_right,49,0;path,motion_diagonal_up_right,56,0;path,motion_diamond,3,0;path,motion_down,42,0;path,motion_equal_triangle,13,0;path,motion_figure_8_four,28,0;path,motion_football,12,0;path,motion_funnel,52,0;path,motion_heart,9,0;path,motion_heartbeat,45,0;path,motion_hexagon,4,0;path,motion_horizontal_figure_8_four,26,0;path,motion_inverted_square,34,0;path,motion_inverted_triangle,33,0;path,motion_left,35,0;path,motion_loop_de_loop,24,0;path,motion_neutron,29,0;path,motion_octagon,10,0;path,motion_parallelogram,14,0;path,motion_path_4_point_star,16,0;path,motion_path_5_point_star,5,0;path,motion_path_6_point_star,11,0;path,motion_path_8_point_star,17,0;path,motion_peanut,27,0;path,motion_pentagon,15,0;path,motion_plus,32,0;path,motion_pointy_star,19,0;path,motion_right,63,0;path,motion_right_triangle,2,0;path,motion_sine_spiral_left,55,0;path,motion_sine_spiral_right,46,0;path,motion_sine_wave,40,0;path,motion_spring,53,0;path,motion_square,7,0;path,motion_stairs_down,62,0;path,motion_swoosh,30,0;path,motion_s_curve_1,59,0;path,motion_s_curve_2,39,0;path,motion_teardrop,18,0;path,motion_trapezoid,8,0;path,motion_turn_down,50,0;path,motion_turn_down_right,36,0;path,motion_turn_up,43,0;path,motion_turn_up_right,57,0;path,motion_up,64,0;path,motion_vertical_figure_8,22,0;path,motion_wave,47,0;path,motion_zigzag,38,0".split(";");

        function getPresetData(sPreset) {
            var aPreset = sPreset.split(",");
            return {
                presetClass: aPreset[0],
                presetID: aPreset[2],
                presetSubtype: aPreset[3],
                effectName: aPreset[1]
            };
        }
        var oClassesNameMap = {"emph": 0, "entr": 1, "exit": 2, "path": 4};
        var oPresetClassMap = {};
        var oPresetIDMap = {};

        var sKey;
        var sConstName;
        var sClassesConstScript = "";
        for(sKey in oClassesNameMap) {
            sConstName = "PRESET_CLASS_" + sKey.toUpperCase();
            sClassesConstScript += ("AscFormat." + sConstName + " = AscFormat[\"" + sConstName + "\"] = " + oClassesNameMap[sKey]);

        }

        var sResultScript = "";
        var sClasses;
        var oEffectsByClass = {};
        while(aPresets.length > 0) {
            var sClassPreset = aPresets.pop();
            var oClassPreset = getPresetData(sClassPreset);
            var oCurPreset;
            //find all presets with this presetClass and presetID
            var aSameClassAndTypeEffects = [];
            aSameClassAndTypeEffects.push(oClassPreset);
            var nPreset = aPresets.length - 1;
            var nMinEffectNameLenght = oClassPreset.effectName.length;
            while(nPreset > -1) {
                oCurPreset = getPresetData(aPresets[nPreset]);
                if(oCurPreset.presetClass === oClassPreset.presetClass && oCurPreset.presetID === oClassPreset.presetID) {
                    aPresets.splice(nPreset, 1);
                    aSameClassAndTypeEffects.push(oCurPreset);
                    nMinEffectNameLenght = Math.min(nMinEffectNameLenght, oClassPreset.effectName.length);
                }
                --nPreset;
            }
            //find preset name
            var nPresetNameLength = nMinEffectNameLenght;
            while(nPresetNameLength > 0) {
                var sCheckName = aSameClassAndTypeEffects[0].effectName.slice(0, nPresetNameLength);
                for(nPreset = 1; nPreset < aSameClassAndTypeEffects.length; ++nPreset) {
                    if(aSameClassAndTypeEffects[nPreset].effectName.indexOf(sCheckName) !== 0) {
                        break;
                    }
                }
                if(nPreset === aSameClassAndTypeEffects.length) {
                    break;
                }
                nPresetNameLength--;
            }
            var nPresetNameForSliceLength = nPresetNameLength;
            if(oClassPreset.effectName.charAt(nPresetNameLength - 1) === "_") {
                nPresetNameForSliceLength = nPresetNameLength - 1;
            }
            var sEffectName = oClassPreset.effectName.slice(0, nPresetNameForSliceLength).toUpperCase();
            if(!oEffectsByClass[oClassPreset.presetClass]) {
                oEffectsByClass[oClassPreset.presetClass] = {"type": oClassesNameMap[oClassPreset.presetClass], "subtypes": {}};
            }
            var oEffectClass = oEffectsByClass[oClassPreset.presetClass]["subtypes"];
            oEffectClass[sEffectName] = {
                "type": parseInt(oClassPreset.presetID),
                "subtypes": {

                }
            };
            if(aSameClassAndTypeEffects.length > 1) {
                for(nPreset = 0; nPreset < aSameClassAndTypeEffects.length; ++nPreset) {
                    oCurPreset = aSameClassAndTypeEffects[nPreset];
                    var sPresetSubtypeName = oCurPreset.effectName.slice(nPresetNameLength);
                    oEffectClass[sEffectName]["subtypes"][sPresetSubtypeName] = oCurPreset.presetSubtype;
                }
            }
        }
        var aEffectClassesStrings = [];
        var aArraysOfEffects = [];
        var aArrayOfEffectsSubtypes = [];
        var sConstString, oConstStringStruct;

        function getStr(sObj, sKey, sVal) {
            return sConstString = sObj + "." + sKey + " = " + sObj + "[\"" + sKey + "\"] = " + sVal + ";";
        }
        for(sKey in oEffectsByClass) {
            var oClassOfEffects = oEffectsByClass[sKey];
            var sEffectClassName = "PRESET_CLASS_" + sKey.toUpperCase();
            sConstString = getStr("AscFormat", sEffectClassName, oClassOfEffects["type"]);
            oConstStringStruct = {idx: parseInt(oClassOfEffects["type"]), str: sConstString};
            aEffectClassesStrings.push(oConstStringStruct);
            var oTypesOfEffects = oClassOfEffects["subtypes"];
            var aCurEffects = [];
            aArraysOfEffects.push(aCurEffects);
            for(var sKey2 in oTypesOfEffects) {
                var oEffect = oTypesOfEffects[sKey2];
                sConstString = getStr("AscFormat", sKey2, oEffect["type"]);
                oConstStringStruct = {idx: parseInt(oEffect["type"]), str: sConstString};
                aCurEffects.push(oConstStringStruct);
                if(Object.keys(oEffect["subtypes"]).length !== 0) {
                    var oCurEffectSubtypes = [];
                    aArrayOfEffectsSubtypes.push(oCurEffectSubtypes);
                    var oEffectSubtypes = oEffect["subtypes"];
                    for(var sKey3 in oEffectSubtypes) {
                        var sSubtypeConstName = sKey2 + "_" + sKey3.toUpperCase();
                        sConstString = getStr("AscFormat", sSubtypeConstName, oEffectSubtypes[sKey3]);
                        oCurEffectSubtypes.push({idx: parseInt(oEffectSubtypes[sKey3]), str: sConstString});
                    }
                }
            }
        }
        var sResultScript = "";
        function writeConstString(aStrObjects) {
            var sResult = "";
            if(Array.isArray(aStrObjects[0])) {
                for(var nIdx = 0; nIdx < aStrObjects.length; ++nIdx) {
                    sResult += writeConstString(aStrObjects[nIdx]);
                }
                return sResult;
            }
            sResult = "\n";
            aStrObjects.sort(function(a, b){return a.idx - b.idx});
            for(var nIdx = 0; nIdx < aStrObjects.length; ++nIdx) {
                sResult += (aStrObjects[nIdx].str + "\n");
            }
            return sResult;
        }
        sResultScript += writeConstString(aEffectClassesStrings);
        sResultScript += writeConstString(aArraysOfEffects);
        sResultScript += writeConstString(aArrayOfEffectsSubtypes);
        console.log(sResultScript);
    }
    AscFormat.generate_preset_data = generate_preset_data;


    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].CAnimPaneHeader = CAnimPaneHeader;
    window['AscCommon'].CSeqListContainer = CSeqListContainer;
    window['AscCommon'].CTimelineContainer = CTimelineContainer;

    var ANIMATION_PRESET_CLASSES = [];
    var PRESET_TYPES;
    var PRESET_SUBTYPES;
    ANIMATION_PRESET_CLASSES[0] = [];
    PRESET_TYPES = ANIMATION_PRESET_CLASSES[0] = [];
    PRESET_TYPES[1] = [];
    PRESET_SUBTYPES = PRESET_TYPES[1] = [];
    PRESET_SUBTYPES[2] = "PPTY;v10;486;4gEAAPr7ANsBAAD6AwEFAgYADgAAAAAPBQAAABABAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BJwBAAD6+wCVAQAAAwAAAAeAAAAA+gAAAQH7AGIAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgkAAAADBAAAAPoAAfsNhAAAAPr7AGIAAAD6+wAWAAAA+gMBDwcAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAA1+AAAA+vsAXgAAAPr7ABYAAAD6AwEPCAAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCJQAAAPr7AB4AAAABAAAAABUAAAD6AAcAAABmAGkAbABsAC4AbwBuAPsBFAAAAPoBBAAAAHQAcgB1AGUA+wAAAAAA";
    PRESET_TYPES[3] = [];
    PRESET_SUBTYPES = PRESET_TYPES[3] = [];
    PRESET_SUBTYPES[2] = "PPTY;v10;224;3AAAAPr7ANUAAAD6AwEFAgYADgAAAAAPBQAAABADAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BJYAAAD6+wCPAAAAAQAAAAeGAAAA+gAAAQH7AGgAAAD6BAD7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLQAAAPr7ACYAAAABAAAAAB0AAAD6AAsAAABzAHQAeQBsAGUALgBjAG8AbABvAHIA+wEAAAAAAgkAAAADBAAAAPoAAfs=";
    PRESET_TYPES[6] = [];
    PRESET_SUBTYPES = PRESET_TYPES[6] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;159;mwAAAPr7AJQAAAD6AwEFAgYADgAAAAAPBQAAABAGAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BFUAAAD6+wBOAAAAAQAAAAtFAAAA+gDwSQIAAfBJAgD7ADQAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[7] = [];
    PRESET_SUBTYPES = PRESET_TYPES[7] = [];
    PRESET_SUBTYPES[2] = "PPTY;v10;359;YwEAAPr7AFwBAAD6AwEFAgYADgAAAAAPBQAAABAHAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB0BAAD6+wAWAQAAAgAAAAeGAAAA+gAAAQH7AGgAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai8AAAD6+wAoAAAAAQAAAAAfAAAA+gAMAAAAcwB0AHIAbwBrAGUALgBjAG8AbABvAHIA+wEAAAAAAgkAAAADBAAAAPoAAfsNggAAAPr7AGIAAAD6+wAWAAAA+gMBDwcAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAcwB0AHIAbwBrAGUALgBvAG4A+wEUAAAA+gEEAAAAdAByAHUAZQD7AAAAAAA=";
    PRESET_TYPES[8] = [];
    PRESET_SUBTYPES = PRESET_TYPES[8] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;184;tAAAAPr7AK0AAAD6AwEFAgYADgAAAAAPBQAAABAIAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BG4AAAD6+wBnAAAAAQAAAApeAAAA+gAAl0kB+wBSAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIZAAAA+vsAEgAAAAEAAAAACQAAAPoAAQAAAHIA+w==";
    PRESET_TYPES[9] = [];
    PRESET_SUBTYPES = PRESET_TYPES[9] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;361;ZQEAAPr7AF4BAAD6BQIGAA4AAAAADwUAAAAQCQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQhAQAA+vsAGgEAAAIAAAANkgAAAPr7AHQAAAD6+wAgAAAA+g8GAAAAEwoAAABpAG4AZABlAGYAaQBuAGkAdABlAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIxAAAA+vsAKgAAAAEAAAAAIQAAAPoADQAAAHMAdAB5AGwAZQAuAG8AcABhAGMAaQB0AHkA+wESAAAA+gEDAAAAMAAuADUA+wAAAAAACHoAAAD6AQUAAABpAG0AYQBnAGUAAgwAAABvAHAAYQBjAGkAdAB5ADoAIAAwAC4ANQD7AEcAAAD6BQIAAABJAEUA+wAgAAAA+g8HAAAAEwoAAABpAG4AZABlAGYAaQBuAGkAdABlAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[10] = [];
    PRESET_SUBTYPES = PRESET_TYPES[10] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;425;pQEAAPr7AJ4BAAD6AwEFAgYADgAAAAAPBQAAABAKAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BF8BAAD6+wBYAQAAAQAAAAZPAQAA+gAABAL7AHIAAAD6BAD7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgBmAG8AbgB0AFcAZQBpAGcAaAB0APsBzQAAAPr7AMYAAAAEAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAAG4AbwByAG0AYQBsAPsAAAAAAAAqAAAA+gAFAAAANQAwADAAMAAwAPsAFAAAAPoBBAAAAGIAbwBsAGQA+wAAAAAAAC4AAAD6AAUAAAA2ADAAMAAwADAA+wAYAAAA+gEGAAAAbgBvAHIAbQBhAGwA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAABuAG8AcgBtAGEAbAD7AAAAAAA=";
    PRESET_TYPES[15] = [];
    PRESET_SUBTYPES = PRESET_TYPES[15] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;262;AgEAAPr7APsAAAD6BQIGAA4AAAAADwUAAAAQDwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wMNAAAA+gABAgIAAAAyADUA+wSsAAAA+vsApQAAAAEAAAANnAAAAPr7AHwAAAD6BAD7ACAAAAD6DwYAAAATCgAAAGkAbgBkAGUAZgBpAG4AaQB0AGUA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AZgBvAG4AdABXAGUAaQBnAGgAdAD7ARQAAAD6AQQAAABiAG8AbABkAPsAAAAAAA==";
    PRESET_TYPES[16] = [];
    PRESET_SUBTYPES = PRESET_TYPES[16] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;498;7gEAAPr7AOcBAAD6AwEFAgYADgAAAAAPBQAAABAQAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDoA8AAPsEmgEAAPr7AJMBAAADAAAADYIAAAD6+wBmAAAA+gQA+wAUAAAA+gMBDwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wItAAAA+vsAJgAAAAEAAAAAHQAAAPoACwAAAHMAdAB5AGwAZQAuAGMAbwBsAG8AcgD7ARAAAAD6+wAJAAAAAwQAAAD6AAH7DXwAAAD6+wBgAAAA+vsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCKQAAAPr7ACIAAAABAAAAABkAAAD6AAkAAABmAGkAbABsAGMAbwBsAG8AcgD7ARAAAAD6+wAJAAAAAwQAAAD6AAH7DYIAAAD6+wBgAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCKQAAAPr7ACIAAAABAAAAABkAAAD6AAkAAABmAGkAbABsAC4AdAB5AHAAZQD7ARYAAAD6AQUAAABzAG8AbABpAGQA+wAAAAAA";
    PRESET_TYPES[18] = [];
    PRESET_SUBTYPES = PRESET_TYPES[18] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;274;DgEAAPr7AAcBAAD6AwEFAgYADgAAAAAPBQAAABASAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDoA8AAPsEugAAAPr7ALMAAAABAAAADaoAAAD6+wCKAAAA+gQA+wAUAAAA+gMBDwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wJRAAAA+vsASgAAAAEAAAAAQQAAAPoAHQAAAHMAdAB5AGwAZQAuAHQAZQB4AHQARABlAGMAbwByAGEAdABpAG8AbgBVAG4AZABlAHIAbABpAG4AZQD7ARQAAAD6AQQAAAB0AHIAdQBlAPsAAAAAAA==";
    PRESET_TYPES[19] = [];
    PRESET_SUBTYPES = PRESET_TYPES[19] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;617;ZQIAAPr7AF4CAAD6AwEFAgYADgAAAAAPBQAAABATAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB8CAAD6+wAYAgAABAAAAAeEAAAA+gAAAQH7AGYAAAD6BAD7ABQAAAD6AwEPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai0AAAD6+wAmAAAAAQAAAAAdAAAA+gALAAAAcwB0AHkAbABlAC4AYwBvAGwAbwByAPsBAAAAAAIJAAAAAwQAAAD6AAH7B34AAAD6AAABAfsAYAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgkAAAADBAAAAPoAAfsNggAAAPr7AGAAAAD6+wAUAAAA+gMBDwgAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIpAAAA+vsAIgAAAAEAAAAAGQAAAPoACQAAAGYAaQBsAGwALgB0AHkAcABlAPsBFgAAAPoBBQAAAHMAbwBsAGkAZAD7AAAAAAANfAAAAPr7AFwAAAD6+wAUAAAA+gMBDwkAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIlAAAA+vsAHgAAAAEAAAAAFQAAAPoABwAAAGYAaQBsAGwALgBvAG4A+wEUAAAA+gEEAAAAdAByAHUAZQD7AAAAAAA=";
    PRESET_TYPES[21] = [];
    PRESET_SUBTYPES = PRESET_TYPES[21] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;643;fwIAAPr7AHgCAAD6AwEFAgYADgAAAAAPBQAAABAVAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BDkCAAD6+wAyAgAABAAAAAeKAAAA+gABAQEFAN1tAAYAAAAABwAAAAD7AGYAAAD6BAD7ABQAAAD6AwEPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai0AAAD6+wAmAAAAAQAAAAAdAAAA+gALAAAAcwB0AHkAbABlAC4AYwBvAGwAbwByAPsBAAAAAAIAAAAAB4QAAAD6AAEBAQUA3W0ABgAAAAAHAAAAAPsAYAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgAAAAAHigAAAPoAAQEBBQDdbQAGAAAAAAcAAAAA+wBmAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLwAAAPr7ACgAAAABAAAAAB8AAAD6AAwAAABzAHQAcgBvAGsAZQAuAGMAbwBsAG8AcgD7AQAAAAACAAAAAA2CAAAA+vsAYAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAA==";
    PRESET_TYPES[22] = [];
    PRESET_SUBTYPES = PRESET_TYPES[22] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;643;fwIAAPr7AHgCAAD6AwEFAgYADgAAAAAPBQAAABAWAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BDkCAAD6+wAyAgAABAAAAAeKAAAA+gABAQEFACOS/wYAAAAABwAAAAD7AGYAAAD6BAD7ABQAAAD6AwEPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai0AAAD6+wAmAAAAAQAAAAAdAAAA+gALAAAAcwB0AHkAbABlAC4AYwBvAGwAbwByAPsBAAAAAAIAAAAAB4QAAAD6AAEBAQUAI5L/BgAAAAAHAAAAAPsAYAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgAAAAAHigAAAPoAAQEBBQAjkv8GAAAAAAcAAAAA+wBmAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLwAAAPr7ACgAAAABAAAAAB8AAAD6AAwAAABzAHQAcgBvAGsAZQAuAGMAbwBsAG8AcgD7AQAAAAACAAAAAA2CAAAA+vsAYAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAA==";
    PRESET_TYPES[23] = [];
    PRESET_SUBTYPES = PRESET_TYPES[23] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;643;fwIAAPr7AHgCAAD6AwEFAgYADgAAAAAPBQAAABAXAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BDkCAAD6+wAyAgAABAAAAAeKAAAA+gABAQEF8XClAAYAAAAABwAAAAD7AGYAAAD6BAD7ABQAAAD6AwEPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai0AAAD6+wAmAAAAAQAAAAAdAAAA+gALAAAAcwB0AHkAbABlAC4AYwBvAGwAbwByAPsBAAAAAAIAAAAAB4QAAAD6AAEBAQXxcKUABgAAAAAHAAAAAPsAYAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgAAAAAHigAAAPoAAQEBBfFwpQAGAAAAAAcAAAAA+wBmAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLwAAAPr7ACgAAAABAAAAAB8AAAD6AAwAAABzAHQAcgBvAGsAZQAuAGMAbwBsAG8AcgD7AQAAAAACAAAAAA2CAAAA+vsAYAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAA==";
    PRESET_TYPES[24] = [];
    PRESET_SUBTYPES = PRESET_TYPES[24] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;643;fwIAAPr7AHgCAAD6AwEFAgYADgAAAAAPBQAAABAYAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BDkCAAD6+wAyAgAABAAAAAeKAAAA+gABAQEFAAAAAAb7zv//B/ad///7AGYAAAD6BAD7ABQAAAD6AwEPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai0AAAD6+wAmAAAAAQAAAAAdAAAA+gALAAAAcwB0AHkAbABlAC4AYwBvAGwAbwByAPsBAAAAAAIAAAAAB4QAAAD6AAEBAQUAAAAABvvO//8H9p3///sAYAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgAAAAAHigAAAPoAAQEBBQAAAAAG+87//wf2nf//+wBmAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLwAAAPr7ACgAAAABAAAAAB8AAAD6AAwAAABzAHQAcgBvAGsAZQAuAGMAbwBsAG8AcgD7AQAAAAACAAAAAA2CAAAA+vsAYAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAA==";
    PRESET_TYPES[25] = [];
    PRESET_SUBTYPES = PRESET_TYPES[25] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;643;fwIAAPr7AHgCAAD6AwEFAgYADgAAAAAPBQAAABAZAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BDkCAAD6+wAyAgAABAAAAAeKAAAA+gABAQEFAAAAAAZE7P7/BwAAAAD7AGYAAAD6BAD7ABQAAAD6AwEPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai0AAAD6+wAmAAAAAQAAAAAdAAAA+gALAAAAcwB0AHkAbABlAC4AYwBvAGwAbwByAPsBAAAAAAIAAAAAB4QAAAD6AAEBAQUAAAAABkTs/v8HAAAAAPsAYAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgAAAAAHigAAAPoAAQEBBQAAAAAGROz+/wcAAAAA+wBmAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLwAAAPr7ACgAAAABAAAAAB8AAAD6AAwAAABzAHQAcgBvAGsAZQAuAGMAbwBsAG8AcgD7AQAAAAACAAAAAA2CAAAA+vsAYAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAA==";
    PRESET_TYPES[26] = [];
    PRESET_SUBTYPES = PRESET_TYPES[26] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;291;HwEAAPr7ABgBAAD6AwEFAgYADgAAAAAPBQAAABAaAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BNkAAAD6+wDSAAAAAgAAAAh/AAAA+gABAQQAAABmAGEAZABlAPsAaQAAAPr7AEsAAAD6DwYAAAATAwAAADUAMAAwABcaAAAAMAAsACAAMAA7ACAALgAyACwAIAAuADUAOwAgAC4AOAAsACAALgA1ADsAIAAxACwAIAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wtFAAAA+gAomgEAASiaAQD7ADQAAAD6+wAWAAAA+gIBAwEPBwAAABMDAAAAMgA1ADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[27] = [];
    PRESET_SUBTYPES = PRESET_TYPES[27] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;625;bQIAAPr7AGYCAAD6AwIFAgYADgAAAAAPBQAAABAbAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCcCAAD6+wAgAgAABAAAAAeGAAAA+gAAAQH7AGgAAAD6BAD7ABYAAAD6AgEDAg8GAAAAEwMAAAAyADUAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLQAAAPr7ACYAAAABAAAAAB0AAAD6AAsAAABzAHQAeQBsAGUALgBjAG8AbABvAHIA+wEAAAAAAgkAAAADBAAAAPoABvsHgAAAAPoAAAEB+wBiAAAA+vsAFgAAAPoCAQMCDwcAAAATAwAAADIANQAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIpAAAA+vsAIgAAAAEAAAAAGQAAAPoACQAAAGYAaQBsAGwAYwBvAGwAbwByAPsBAAAAAAIJAAAAAwQAAAD6AAb7DYQAAAD6+wBiAAAA+vsAFgAAAPoCAQMCDwgAAAATAwAAADIANQAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIpAAAA+vsAIgAAAAEAAAAAGQAAAPoACQAAAGYAaQBsAGwALgB0AHkAcABlAPsBFgAAAPoBBQAAAHMAbwBsAGkAZAD7AAAAAAANfgAAAPr7AF4AAAD6+wAWAAAA+gIBAwIPCQAAABMDAAAAMgA1ADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiUAAAD6+wAeAAAAAQAAAAAVAAAA+gAHAAAAZgBpAGwAbAAuAG8AbgD7ARQAAAD6AQQAAAB0AHIAdQBlAPsAAAAAAA==";
    PRESET_TYPES[28] = [];
    PRESET_SUBTYPES = PRESET_TYPES[28] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;637;eQIAAPr7AHICAAD6AwEFAgYADgAAAAAPBQAAABAcAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDECcAAPsEJQIAAPr7AB4CAAAEAAAAB4QAAAD6AAABAfsAZgAAAPoEAPsAFAAAAPoDAQ8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLQAAAPr7ACYAAAABAAAAAB0AAAD6AAsAAABzAHQAeQBsAGUALgBjAG8AbABvAHIA+wEAAAAAAgkAAAADBAAAAPoAAfsHfgAAAPoAAAEB+wBgAAAA+vsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCKQAAAPr7ACIAAAABAAAAABkAAAD6AAkAAABmAGkAbABsAGMAbwBsAG8AcgD7AQAAAAACCQAAAAMEAAAA+gAB+w2CAAAA+vsAYAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAAaCAAAA+gABAwMAAAAxAC4ANQAEAPsAbAAAAPoEAPsAFAAAAPoDAQ8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCMwAAAPr7ACwAAAABAAAAACMAAAD6AA4AAABzAHQAeQBsAGUALgBmAG8AbgB0AFMAaQB6AGUA+w==";
    PRESET_TYPES[30] = [];
    PRESET_SUBTYPES = PRESET_TYPES[30] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;643;fwIAAPr7AHgCAAD6AwEFAgYADgAAAAAPBQAAABAeAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BDkCAAD6+wAyAgAABAAAAAeKAAAA+gABAQEFAAAAAAYFMQAABwpiAAD7AGYAAAD6BAD7ABQAAAD6AwEPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Ai0AAAD6+wAmAAAAAQAAAAAdAAAA+gALAAAAcwB0AHkAbABlAC4AYwBvAGwAbwByAPsBAAAAAAIAAAAAB4QAAAD6AAEBAQUAAAAABgUxAAAHCmIAAPsAYAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbABjAG8AbABvAHIA+wEAAAAAAgAAAAAHigAAAPoAAQEBBQAAAAAGBTEAAAcKYgAA+wBmAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCLwAAAPr7ACgAAAABAAAAAB8AAAD6AAwAAABzAHQAcgBvAGsAZQAuAGMAbwBsAG8AcgD7AQAAAAACAAAAAA2CAAAA+vsAYAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AikAAAD6+wAiAAAAAQAAAAAZAAAA+gAJAAAAZgBpAGwAbAAuAHQAeQBwAGUA+wEWAAAA+gEFAAAAcwBvAGwAaQBkAPsAAAAAAA==";
    PRESET_TYPES[32] = [];
    PRESET_SUBTYPES = PRESET_TYPES[32] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;736;3AIAAPr7ANUCAAD6AwEFAgYADgAAAAAPBQAAABAgAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BJYCAAD6+wCPAgAABQAAAAp6AAAA+gDA1AEA+wBuAAAA+vsAMgAAAPoDAQ8GAAAAEwMAAAAxADAAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCGQAAAPr7ABIAAAABAAAAAAkAAAD6AAEAAAByAPsKfgAAAPoAgFb8//sAcgAAAPr7ADYAAAD6AwEPBwAAABMDAAAAMgAwADAA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADIAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIZAAAA+vsAEgAAAAEAAAAACQAAAPoAAQAAAHIA+wp+AAAA+gCAqQMA+wByAAAA+vsANgAAAPoDAQ8IAAAAEwMAAAAyADAAMAD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AhkAAAD6+wASAAAAAQAAAAAJAAAA+gABAAAAcgD7Cn4AAAD6AIBW/P/7AHIAAAD6+wA2AAAA+gMBDwkAAAATAwAAADIAMAAwAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA2ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCGQAAAPr7ABIAAAABAAAAAAkAAAD6AAEAAAByAPsKfgAAAPoAwNQBAPsAcgAAAPr7ADYAAAD6AwEPCgAAABMDAAAAMgAwADAA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADgAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIZAAAA+vsAEgAAAAEAAAAACQAAAPoAAQAAAHIA+w==";
    PRESET_TYPES[34] = [];
    PRESET_SUBTYPES = PRESET_TYPES[34] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;845;SQMAAPr7AEIDAAD6AwEFAgYADgAAAAAPBQAAABAiAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDECcAAPsE9QIAAPr7AO4CAAAFAAAACd0AAAD6AAEBAQIYAAAATQAgADAALgAwACAAMAAuADAAIABMACAAMAAuADAAIAAtADAALgAwADcAMgAxADMAAwAAAAD7AJgAAAD6+wA+AAAA+gBQwwAAAgEDAQxQwwAADwYAAAATAwAAADIANQAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+wp6AAAA+gBg4xYA+wBuAAAA+vsAMgAAAPoDAQ8HAAAAEwMAAAAxADIANQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCGQAAAPr7ABIAAAABAAAAAAkAAAD6AAEAAAByAPsKfgAAAPoAoBzp//sAcgAAAPr7ADYAAAD6AwEPCAAAABMDAAAAMQAyADUA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADEAMgA1APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIZAAAA+vsAEgAAAAEAAAAACQAAAPoAAQAAAHIA+wp+AAAA+gCgHOn/+wByAAAA+vsANgAAAPoDAQ8JAAAAEwMAAAAxADIANQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAAMgA1ADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AhkAAAD6+wASAAAAAQAAAAAJAAAA+gABAAAAcgD7Cn4AAAD6AGDjFgD7AHIAAAD6+wA2AAAA+gMBDwoAAAATAwAAADEAMgA1APsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAAzADcANQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCGQAAAPr7ABIAAAABAAAAAAkAAAD6AAEAAAByAPs=";
    PRESET_TYPES[35] = [];
    PRESET_SUBTYPES = PRESET_TYPES[35] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;325;QQEAAPr7ADoBAAD6AwEFAgYADgAAAAAPBQAAABAjAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BPsAAAD6+wD0AAAAAQAAAAbrAAAA+gAABAL7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7AWsAAAD6+wBkAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAAAMAAAAPoABQAAADUAMAAwADAAMAD7ABoAAAD6AQcAAAB2AGkAcwBpAGIAbABlAPsAAAAAAA==";
    PRESET_TYPES[36] = [];
    PRESET_SUBTYPES = PRESET_TYPES[36] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;668;mAIAAPr7AJECAAD6AwEFAgYADgAAAAAPBQAAABAkAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDECcAAPsERAIAAPr7AD0CAAAEAAAAC2MAAAD6BIA4AQAFoIYBAPsAUgAAAPr7ADQAAAD6AgEDAQ8GAAAAEwMAAAAyADUAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsGogAAAPoAAQENAAAAKAAjAHAAcAB0AF8AdwAqADAALgAxADAAKQAEAPsAeAAAAPr7ADQAAAD6AgEDAQ8HAAAAEwMAAAAyADUAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wakAAAA+gABAQ4AAAAoAC0AIwBwAHAAdABfAHcAKgAwAC4AMQAwACkABAD7AHgAAAD6+wA0AAAA+gIBAwEPCAAAABMDAAAAMgA1ADAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsKfAAAAPoAAK34//sAcAAAAPr7ADQAAAD6AgEDAQ8JAAAAEwMAAAAyADUAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCGQAAAPr7ABIAAAABAAAAAAkAAAD6AAEAAAByAPs=";
    ANIMATION_PRESET_CLASSES[1] = [];
    PRESET_TYPES = ANIMATION_PRESET_CLASSES[1] = [];
    PRESET_TYPES[1] = [];
    PRESET_SUBTYPES = PRESET_TYPES[1] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;264;BAEAAPr7AP0AAAD6AwEFAgYBDgAAAAAPBQAAABABAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BL4AAAD6+wC3AAAAAQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAA";
    PRESET_TYPES[2] = [];
    PRESET_SUBTYPES = PRESET_TYPES[2] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;708;wAIAAPr7ALkCAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHoCAAD6+wBzAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtUAAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFrAAAA+vsAZAAAAAIAAAAAJgAAAPoAAQAAADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAG3QAAAPoAAQQA+wBaAAAA+gEA+wAUAAAA+gMBDwgAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeQD7AXMAAAD6+wBsAAAAAgAAAAAuAAAA+gABAAAAMAD7ACAAAAD6AQoAAAAwAC0AIwBwAHAAdABfAGgALwAyAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_SUBTYPES[2] = "PPTY;v10;708;wAIAAPr7ALkCAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHoCAAD6+wBzAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABt0AAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAAMQArACMAcABwAHQAXwB3AC8AMgD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbVAAAA+gABBAD7AFoAAAD6AQD7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB5APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_SUBTYPES[3] = "PPTY;v10;716;yAIAAPr7AMECAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQMAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BIICAAD6+wB7AgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABt0AAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAAMQArACMAcABwAHQAXwB3AC8AMgD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbdAAAA+gABBAD7AFoAAAD6AQD7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBcwAAAPr7AGwAAAACAAAAAC4AAAD6AAEAAAAwAPsAIAAAAPoBCgAAADAALQAjAHAAcAB0AF8AaAAvADIA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAA=";
    PRESET_SUBTYPES[4] = "PPTY;v10;708;wAIAAPr7ALkCAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHoCAAD6+wBzAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtUAAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFrAAAA+vsAZAAAAAIAAAAAJgAAAPoAAQAAADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAG3QAAAPoAAQQA+wBaAAAA+gEA+wAUAAAA+gMBDwgAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeQD7AXMAAAD6+wBsAAAAAgAAAAAuAAAA+gABAAAAMAD7ACAAAAD6AQoAAAAxACsAIwBwAHAAdABfAGgALwAyAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_SUBTYPES[6] = "PPTY;v10;716;yAIAAPr7AMECAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQYAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BIICAAD6+wB7AgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABt0AAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAAMQArACMAcABwAHQAXwB3AC8AMgD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbdAAAA+gABBAD7AFoAAAD6AQD7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBcwAAAPr7AGwAAAACAAAAAC4AAAD6AAEAAAAwAPsAIAAAAPoBCgAAADEAKwAjAHAAcAB0AF8AaAAvADIA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAA=";
    PRESET_SUBTYPES[8] = "PPTY;v10;708;wAIAAPr7ALkCAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHoCAAD6+wBzAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABt0AAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAAMAAtACMAcABwAHQAXwB3AC8AMgD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbVAAAA+gABBAD7AFoAAAD6AQD7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB5APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_SUBTYPES[9] = "PPTY;v10;716;yAIAAPr7AMECAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQkAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BIICAAD6+wB7AgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABt0AAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAAMAAtACMAcABwAHQAXwB3AC8AMgD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbdAAAA+gABBAD7AFoAAAD6AQD7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBcwAAAPr7AGwAAAACAAAAAC4AAAD6AAEAAAAwAPsAIAAAAPoBCgAAADAALQAjAHAAcAB0AF8AaAAvADIA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAA=";
    PRESET_SUBTYPES[12] = "PPTY;v10;716;yAIAAPr7AMECAAD6AwEFAgYBDgAAAAAPBQAAABACAAAAEQwAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BIICAAD6+wB7AgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABt0AAAD6AAEEAPsAWgAAAPoBAPsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAAMAAtACMAcABwAHQAXwB3AC8AMgD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbdAAAA+gABBAD7AFoAAAD6AQD7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBcwAAAPr7AGwAAAACAAAAAC4AAAD6AAEAAAAwAPsAIAAAAPoBCgAAADEAKwAjAHAAcAB0AF8AaAAvADIA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAA=";
    PRESET_TYPES[3] = [];
    PRESET_SUBTYPES = PRESET_TYPES[3] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;363;ZwEAAPr7AGABAAD6AwEFAgYBDgAAAAAPBQAAABADAAAAEQUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCEBAAD6+wAaAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACF4AAAD6AAABEAAAAGIAbABpAG4AZABzACgAdgBlAHIAdABpAGMAYQBsACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_SUBTYPES[10] = "PPTY;v10;367;awEAAPr7AGQBAAD6AwEFAgYBDgAAAAAPBQAAABADAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCUBAAD6+wAeAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGIAAAD6AAABEgAAAGIAbABpAG4AZABzACgAaABvAHIAaQB6AG8AbgB0AGEAbAApAPsAMAAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[4] = [];
    PRESET_SUBTYPES = PRESET_TYPES[4] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;347;VwEAAPr7AFABAAD6AwEFAgYBDgAAAAAPBQAAABAEAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBEBAAD6+wAKAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACE4AAAD6AAABBwAAAGIAbwB4ACgAaQBuACkA+wAyAAAA+vsAFAAAAPoPBwAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[32] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYBDgAAAAAPBQAAABAEAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFAAAAD6AAABCAAAAGIAbwB4ACgAbwB1AHQAKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[5] = [];
    PRESET_SUBTYPES = PRESET_TYPES[5] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;367;awEAAPr7AGQBAAD6AwEFAgYBDgAAAAAPBQAAABAFAAAAEQUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCUBAAD6+wAeAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGIAAAD6AAABEgAAAGMAaABlAGMAawBlAHIAYgBvAGEAcgBkACgAZABvAHcAbgApAPsAMAAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[10] = "PPTY;v10;371;bwEAAPr7AGgBAAD6AwEFAgYBDgAAAAAPBQAAABAFAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCkBAAD6+wAiAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGYAAAD6AAABFAAAAGMAaABlAGMAawBlAHIAYgBvAGEAcgBkACgAYQBjAHIAbwBzAHMAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_TYPES[6] = [];
    PRESET_SUBTYPES = PRESET_TYPES[6] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYBDgAAAAAPBQAAABAGAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFQAAAD6AAABCgAAAGMAaQByAGMAbABlACgAaQBuACkA+wAyAAAA+vsAFAAAAPoPBwAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[32] = "PPTY;v10;355;XwEAAPr7AFgBAAD6AwEFAgYBDgAAAAAPBQAAABAGAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBkBAAD6+wASAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFYAAAD6AAABCwAAAGMAaQByAGMAbABlACgAbwB1AHQAKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[8] = [];
    PRESET_SUBTYPES = PRESET_TYPES[8] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;355;XwEAAPr7AFgBAAD6AwEFAgYBDgAAAAAPBQAAABAIAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBkBAAD6+wASAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFYAAAD6AAABCwAAAGQAaQBhAG0AbwBuAGQAKABpAG4AKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[32] = "PPTY;v10;357;YQEAAPr7AFoBAAD6AwEFAgYBDgAAAAAPBQAAABAIAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBsBAAD6+wAUAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFgAAAD6AAABDAAAAGQAaQBhAG0AbwBuAGQAKABvAHUAdAApAPsAMgAAAPr7ABQAAAD6DwcAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[9] = [];
    PRESET_SUBTYPES = PRESET_TYPES[9] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;347;VwEAAPr7AFABAAD6AwEFAgYBDgAAAAAPBQAAABAJAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBEBAAD6+wAKAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACE4AAAD6AAABCAAAAGQAaQBzAHMAbwBsAHYAZQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_TYPES[10] = [];
    PRESET_SUBTYPES = PRESET_TYPES[10] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;339;TwEAAPr7AEgBAAD6AwEFAgYBDgAAAAAPBQAAABAKAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BAkBAAD6+wACAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACEYAAAD6AAABBAAAAGYAYQBkAGUA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[12] = [];
    PRESET_SUBTYPES = PRESET_TYPES[12] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;599;UwIAAPr7AEwCAAD6AwEFAgYBDgAAAAAPBQAAABAMAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA0CAAD6+wAGAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABvMAAAD6AAEEAPsAWAAAAPoBAPsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBiwAAAPr7AIQAAAACAAAAAEYAAAD6AAEAAAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB5AC0AIwBwAHAAdABfAGgAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAAIUgAAAPoAAAEKAAAAdwBpAHAAZQAoAGQAbwB3AG4AKQD7ADAAAAD6+wASAAAA+g8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[2] = "PPTY;v10;599;UwIAAPr7AEwCAAD6AwEFAgYBDgAAAAAPBQAAABAMAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA0CAAD6+wAGAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABvMAAAD6AAEEAPsAWAAAAPoBAPsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBiwAAAPr7AIQAAAACAAAAAEYAAAD6AAEAAAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB4ACsAIwBwAHAAdABfAHcAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAIUgAAAPoAAAEKAAAAdwBpAHAAZQAoAGwAZQBmAHQAKQD7ADAAAAD6+wASAAAA+g8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[4] = "PPTY;v10;595;TwIAAPr7AEgCAAD6AwEFAgYBDgAAAAAPBQAAABAMAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BAkCAAD6+wACAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABvMAAAD6AAEEAPsAWAAAAPoBAPsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBiwAAAPr7AIQAAAACAAAAAEYAAAD6AAEAAAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB5ACsAIwBwAHAAdABfAGgAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAAITgAAAPoAAAEIAAAAdwBpAHAAZQAoAHUAcAApAPsAMAAAAPr7ABIAAAD6DwgAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[8] = "PPTY;v10;601;VQIAAPr7AE4CAAD6AwEFAgYBDgAAAAAPBQAAABAMAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA8CAAD6+wAIAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABvMAAAD6AAEEAPsAWAAAAPoBAPsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBiwAAAPr7AIQAAAACAAAAAEYAAAD6AAEAAAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB4AC0AIwBwAHAAdABfAHcAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAIVAAAAPoAAAELAAAAdwBpAHAAZQAoAHIAaQBnAGgAdAApAPsAMAAAAPr7ABIAAAD6DwgAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[13] = [];
    PRESET_SUBTYPES = PRESET_TYPES[13] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYBDgAAAAAPBQAAABANAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFAAAAD6AAABCAAAAHAAbAB1AHMAKABpAG4AKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[32] = "PPTY;v10;351;WwEAAPr7AFQBAAD6AwEFAgYBDgAAAAAPBQAAABANAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBUBAAD6+wAOAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFIAAAD6AAABCQAAAHAAbAB1AHMAKABvAHUAdAApAPsAMgAAAPr7ABQAAAD6DwcAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[14] = [];
    PRESET_SUBTYPES = PRESET_TYPES[14] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;369;bQEAAPr7AGYBAAD6AwEFAgYBDgAAAAAPBQAAABAOAAAAEQUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCcBAAD6+wAgAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGQAAAD6AAABEwAAAHIAYQBuAGQAbwBtAGIAYQByACgAdgBlAHIAdABpAGMAYQBsACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_SUBTYPES[10] = "PPTY;v10;373;cQEAAPr7AGoBAAD6AwEFAgYBDgAAAAAPBQAAABAOAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCsBAAD6+wAkAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGgAAAD6AAABFQAAAHIAYQBuAGQAbwBtAGIAYQByACgAaABvAHIAaQB6AG8AbgB0AGEAbAApAPsAMAAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[15] = [];
    PRESET_SUBTYPES = PRESET_TYPES[15] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1342;OgUAAPr7ADMFAAD6AwEFAgYBDgAAAAAPBQAAABAPAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BPQEAAD6+wDtBAAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABskAAAD6AAEEAPsAWgAAAPr7ABYAAAD6AwEPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFfAAAA+vsAWAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAAAAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AdwD7AAAAAAAGyQAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AV8AAAD6+wBYAAAAAgAAAAAaAAAA+gABAAAAMAD7AAwAAAD6AwAAAAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAZIAQAA+gABBAD7AFoAAAD6+wAWAAAA+gMBDwkAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsB3gAAAPr7ANcAAAACAAAAAKUAAAD6AAEAAAAwAAFDAAAAIwBwAHAAdABfAHgAKwAoAGMAbwBzACgALQAyACoAcABpACoAKAAxAC0AJAApACkAKgAtACMAcABwAHQAXwB4AC0AcwBpAG4AKAAtADIAKgBwAGkAKgAoADEALQAkACkAKQAqACgAMQAtACMAcABwAHQAXwB5ACkAKQAqACgAMQAtACQAKQD7AAwAAAD6AwAAAAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDoIYBAPsAAAAAAAZIAQAA+gABBAD7AFoAAAD6+wAWAAAA+gMBDwoAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsB3gAAAPr7ANcAAAACAAAAAKUAAAD6AAEAAAAwAAFDAAAAIwBwAHAAdABfAHkAKwAoAHMAaQBuACgALQAyACoAcABpACoAKAAxAC0AJAApACkAKgAtACMAcABwAHQAXwB4ACsAYwBvAHMAKAAtADIAKgBwAGkAKgAoADEALQAkACkAKQAqACgAMQAtACMAcABwAHQAXwB5ACkAKQAqACgAMQAtACQAKQD7AAwAAAD6AwAAAAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDoIYBAPsAAAAAAA==";
    PRESET_TYPES[16] = [];
    PRESET_SUBTYPES = PRESET_TYPES[16] = [];
    PRESET_SUBTYPES[21] = "PPTY;v10;363;ZwEAAPr7AGABAAD6AwEFAgYBDgAAAAAPBQAAABAQAAAAERUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCEBAAD6+wAaAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACF4AAAD6AAABEAAAAGIAYQByAG4AKABpAG4AVgBlAHIAdABpAGMAYQBsACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_SUBTYPES[26] = "PPTY;v10;367;awEAAPr7AGQBAAD6AwEFAgYBDgAAAAAPBQAAABAQAAAAERoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCUBAAD6+wAeAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGIAAAD6AAABEgAAAGIAYQByAG4AKABpAG4ASABvAHIAaQB6AG8AbgB0AGEAbAApAPsAMAAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[37] = "PPTY;v10;365;aQEAAPr7AGIBAAD6AwEFAgYBDgAAAAAPBQAAABAQAAAAESUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCMBAAD6+wAcAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGAAAAD6AAABEQAAAGIAYQByAG4AKABvAHUAdABWAGUAcgB0AGkAYwBhAGwAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[42] = "PPTY;v10;369;bQEAAPr7AGYBAAD6AwEFAgYBDgAAAAAPBQAAABAQAAAAESoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCcBAAD6+wAgAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGQAAAD6AAABEwAAAGIAYQByAG4AKABvAHUAdABIAG8AcgBpAHoAbwBuAHQAYQBsACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[17] = [];
    PRESET_SUBTYPES = PRESET_TYPES[17] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;1134;agQAAPr7AGMEAAD6AwEFAgYBDgAAAAAPBQAAABARAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCQEAAD6+wAdBAAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAABuUAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBfQAAAPr7AHYAAAACAAAAADgAAAD6AAEAAAAwAPsAKgAAAPoBDwAAACMAcABwAHQAXwB5AC0AIwBwAHAAdABfAGgALwAyAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB3APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAA";
    PRESET_SUBTYPES[2] = "PPTY;v10;1134;agQAAPr7AGMEAAD6AwEFAgYBDgAAAAAPBQAAABARAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCQEAAD6+wAdBAAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABuUAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBfQAAAPr7AHYAAAACAAAAADgAAAD6AAEAAAAwAPsAKgAAAPoBDwAAACMAcABwAHQAXwB4ACsAIwBwAHAAdABfAHcALwAyAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB5APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAA";
    PRESET_SUBTYPES[4] = "PPTY;v10;1134;agQAAPr7AGMEAAD6AwEFAgYBDgAAAAAPBQAAABARAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCQEAAD6+wAdBAAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAABuUAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBfQAAAPr7AHYAAAACAAAAADgAAAD6AAEAAAAwAPsAKgAAAPoBDwAAACMAcABwAHQAXwB5ACsAIwBwAHAAdABfAGgALwAyAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB3APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAA";
    PRESET_SUBTYPES[8] = "PPTY;v10;1134;agQAAPr7AGMEAAD6AwEFAgYBDgAAAAAPBQAAABARAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCQEAAD6+wAdBAAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABuUAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBfQAAAPr7AHYAAAACAAAAADgAAAD6AAEAAAAwAPsAKgAAAPoBDwAAACMAcABwAHQAXwB4AC0AIwBwAHAAdABfAHcALwAyAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB5APsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAA";
    PRESET_SUBTYPES[10] = "PPTY;v10;684;qAIAAPr7AKECAAD6AwEFAgYBDgAAAAAPBQAAABARAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BGICAAD6+wBbAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABtMAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAA";
    PRESET_TYPES[18] = [];
    PRESET_SUBTYPES = PRESET_TYPES[18] = [];
    PRESET_SUBTYPES[3] = "PPTY;v10;361;ZQEAAPr7AF4BAAD6AwEFAgYBDgAAAAAPBQAAABASAAAAEQMAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB8BAAD6+wAYAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFwAAAD6AAABDwAAAHMAdAByAGkAcABzACgAdQBwAFIAaQBnAGgAdAApAPsAMAAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[6] = "PPTY;v10;365;aQEAAPr7AGIBAAD6AwEFAgYBDgAAAAAPBQAAABASAAAAEQYAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCMBAAD6+wAcAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACGAAAAD6AAABEQAAAHMAdAByAGkAcABzACgAZABvAHcAbgBSAGkAZwBoAHQAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[9] = "PPTY;v10;359;YwEAAPr7AFwBAAD6AwEFAgYBDgAAAAAPBQAAABASAAAAEQkAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB0BAAD6+wAWAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFoAAAD6AAABDgAAAHMAdAByAGkAcABzACgAdQBwAEwAZQBmAHQAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[12] = "PPTY;v10;363;ZwEAAPr7AGABAAD6AwEFAgYBDgAAAAAPBQAAABASAAAAEQwAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCEBAAD6+wAaAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACF4AAAD6AAABEAAAAHMAdAByAGkAcABzACgAZABvAHcAbgBMAGUAZgB0ACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[19] = [];
    PRESET_SUBTYPES = PRESET_TYPES[19] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;721;zQIAAPr7AMYCAAD6AwEFAgYBDgAAAAAPBQAAABATAAAAEQUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BIcCAAD6+wCAAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtUAAAD6AAEEAPsAWgAAAPr7ABYAAAD6AwEPBwAAABMEAAAANQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFrAAAA+vsAZAAAAAIAAAAAJgAAAPoAAQAAADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AdwD7AAAAAAAG6gAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAA1ADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AYAAAAD6+wB5AAAAAgAAAABHAAAA+gABAAAAMAABFAAAACMAcABwAHQAXwBoACoAcwBpAG4AKAAyAC4ANQAqAHAAaQAqACQAKQD7AAwAAAD6AwAAAAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDoIYBAPsAAAAAAA==";
    PRESET_SUBTYPES[10] = "PPTY;v10;721;zQIAAPr7AMYCAAD6AwEFAgYBDgAAAAAPBQAAABATAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BIcCAAD6+wCAAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABuoAAAD6AAEEAPsAWgAAAPr7ABYAAAD6AwEPBwAAABMEAAAANQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wGAAAAA+vsAeQAAAAIAAAAARwAAAPoAAQAAADAAARQAAAAjAHAAcAB0AF8AdwAqAHMAaQBuACgAMgAuADUAKgBwAGkAKgAkACkA+wAMAAAA+gMAAAAA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6A6CGAQD7AAAAAAAG1QAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAA1ADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AWsAAAD6+wBkAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AaAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAA==";
    PRESET_TYPES[20] = [];
    PRESET_SUBTYPES = PRESET_TYPES[20] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;343;UwEAAPr7AEwBAAD6AwEFAgYBDgAAAAAPBQAAABAUAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA0BAAD6+wAGAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACEoAAAD6AAABBQAAAHcAZQBkAGcAZQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[21] = [];
    PRESET_SUBTYPES = PRESET_TYPES[21] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYBDgAAAAAPBQAAABAVAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFAAAAD6AAABCAAAAHcAaABlAGUAbAAoADEAKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[2] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYBDgAAAAAPBQAAABAVAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFAAAAD6AAABCAAAAHcAaABlAGUAbAAoADIAKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[3] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYBDgAAAAAPBQAAABAVAAAAEQMAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFAAAAD6AAABCAAAAHcAaABlAGUAbAAoADMAKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[4] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYBDgAAAAAPBQAAABAVAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFAAAAD6AAABCAAAAHcAaABlAGUAbAAoADQAKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_SUBTYPES[8] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYBDgAAAAAPBQAAABAVAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFAAAAD6AAABCAAAAHcAaABlAGUAbAAoADgAKQD7ADIAAAD6+wAUAAAA+g8HAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[22] = [];
    PRESET_SUBTYPES = PRESET_TYPES[22] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;347;VwEAAPr7AFABAAD6AwEFAgYBDgAAAAAPBQAAABAWAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBEBAAD6+wAKAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACE4AAAD6AAABCAAAAHcAaQBwAGUAKAB1AHAAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[2] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYBDgAAAAAPBQAAABAWAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFQAAAD6AAABCwAAAHcAaQBwAGUAKAByAGkAZwBoAHQAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_SUBTYPES[4] = "PPTY;v10;351;WwEAAPr7AFQBAAD6AwEFAgYBDgAAAAAPBQAAABAWAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBUBAAD6+wAOAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFIAAAD6AAABCgAAAHcAaQBwAGUAKABkAG8AdwBuACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_SUBTYPES[8] = "PPTY;v10;351;WwEAAPr7AFQBAAD6AwEFAgYBDgAAAAAPBQAAABAWAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBUBAAD6+wAOAQAAAgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACFIAAAD6AAABCgAAAHcAaQBwAGUAKABsAGUAZgB0ACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[23] = [];
    PRESET_SUBTYPES = PRESET_TYPES[23] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;672;nAIAAPr7AJUCAAD6AwEFAgYBDgAAAAAPBQAAABAXAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BFYCAAD6+wBPAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAA";
    PRESET_SUBTYPES[32] = "PPTY;v10;704;vAIAAPr7ALUCAAD6AwEFAgYBDgAAAAAPBQAAABAXAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHYCAAD6+wBvAgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtcAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBbwAAAPr7AGgAAAACAAAAACoAAAD6AAEAAAAwAPsAHAAAAPoBCAAAADQAKgAjAHAAcAB0AF8AdwD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB3APsAAAAAAAbXAAAA+gABBAD7AFgAAAD6+wAUAAAA+gMBDwgAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AW8AAAD6+wBoAAAAAgAAAAAqAAAA+gABAAAAMAD7ABwAAAD6AQgAAAA0ACoAIwBwAHAAdABfAGgA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AaAD7AAAAAAA=";
    PRESET_SUBTYPES[36] = "PPTY;v10;1370;VgUAAPr7AE8FAAD6AwEFAgYBDgAAAAAPBQAAABAXAAAAESQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBAFAAD6+wAJBQAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABiUBAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBvQAAAPr7ALYAAAACAAAAAHgAAAD6AAEAAAAwAPsAagAAAPoBLwAAACgANgAqAG0AaQBuACgAbQBhAHgAKAAjAHAAcAB0AF8AdwAqACMAcABwAHQAXwBoACwALgAzACkALAAxACkALQA3AC4ANAApAC8ALQAuADcAKgAjAHAAcAB0AF8AdwD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB3APsAAAAAAAYlAQAA+gABBAD7AFgAAAD6+wAUAAAA+gMBDwgAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7Ab0AAAD6+wC2AAAAAgAAAAB4AAAA+gABAAAAMAD7AGoAAAD6AS8AAAAoADYAKgBtAGkAbgAoAG0AYQB4ACgAIwBwAHAAdABfAHcAKgAjAHAAcAB0AF8AaAAsAC4AMwApACwAMQApAC0ANwAuADQAKQAvAC0ALgA3ACoAIwBwAHAAdABfAGgA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AaAD7AAAAAAAGxwAAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFfAAAA+vsAWAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gNQwwAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAGLQEAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8KAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wHFAAAA+vsAvgAAAAIAAAAAgAAAAPoAAQAAADAA+wByAAAA+gEzAAAAMQArACgANgAqAG0AaQBuACgAbQBhAHgAKAAjAHAAcAB0AF8AdwAqACMAcABwAHQAXwBoACwALgAzACkALAAxACkALQA3AC4ANAApAC8ALQAuADcAKgAjAHAAcAB0AF8AaAAvADIA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAA=";
    PRESET_SUBTYPES[272] = "PPTY;v10;712;xAIAAPr7AL0CAAD6AwEFAgYBDgAAAAAPBQAAABAXAAAAERABAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BH4CAAD6+wB3AgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtsAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBcwAAAPr7AGwAAAACAAAAAC4AAAD6AAEAAAAwAPsAIAAAAPoBCgAAADIALwAzACoAIwBwAHAAdABfAHcA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AdwD7AAAAAAAG2wAAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAAMgAvADMAKgAjAHAAcAB0AF8AaAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAA==";
    PRESET_SUBTYPES[288] = "PPTY;v10;712;xAIAAPr7AL0CAAD6AwEFAgYBDgAAAAAPBQAAABAXAAAAESABAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BH4CAAD6+wB3AgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtsAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBcwAAAPr7AGwAAAACAAAAAC4AAAD6AAEAAAAwAPsAIAAAAPoBCgAAADQALwAzACoAIwBwAHAAdABfAHcA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AdwD7AAAAAAAG2wAAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFzAAAA+vsAbAAAAAIAAAAALgAAAPoAAQAAADAA+wAgAAAA+gEKAAAANAAvADMAKgAjAHAAcAB0AF8AaAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAA==";
    PRESET_SUBTYPES[528] = "PPTY;v10;1080;NAQAAPr7AC0EAAD6AwEFAgYBDgAAAAAPBQAAABAXAAAAERACAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BO4DAAD6+wDnAwAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDUMMAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDUMMAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_TYPES[25] = [];
    PRESET_SUBTYPES = PRESET_TYPES[25] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;2142;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";
    PRESET_TYPES[26] = [];
    PRESET_SUBTYPES = PRESET_TYPES[26] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;3328;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";
    PRESET_TYPES[28] = [];
    PRESET_SUBTYPES = PRESET_TYPES[28] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;712;xAIAAPr7AL0CAAD6AwEFAgYBDgAAAAAPBQAAABAcAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BH4CAAD6+wB3AgAAAwAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABtcAAAD6AAEEAPsAXAAAAPr7ABgAAAD6AwEPBwAAABMFAAAAMQA1ADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AWsAAAD6+wBkAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbfAAAA+gABBAD7AFwAAAD6+wAYAAAA+gMBDwgAAAATBQAAADEANQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFzAAAA+vsAbAAAAAIAAAAAKgAAAPoAAQAAADAA+wAcAAAA+gEIAAAAIwBwAHAAdABfAHkAKwAxAPsAAAAAAAA0AAAA+gAGAAAAMQAwADAAMAAwADAA+wAcAAAA+gEIAAAAIwBwAHAAdABfAHkALQAxAPsAAAAAAA==";
    PRESET_TYPES[30] = [];
    PRESET_SUBTYPES = PRESET_TYPES[30] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1563;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";
    PRESET_TYPES[31] = [];
    PRESET_SUBTYPES = PRESET_TYPES[31] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;965;wQMAAPr7ALoDAAD6AwEFAgYBDgAAAAAPBQAAABAfAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHsDAAD6+wB0AwAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABskAAAD6AAEEAPsAWgAAAPr7ABYAAAD6AwEPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFfAAAA+vsAWAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAAAAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AdwD7AAAAAAAGyQAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AV8AAAD6+wBYAAAAAgAAAAAaAAAA+gABAAAAMAD7AAwAAAD6AwAAAAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAbPAAAA+gABBAD7AGwAAAD6+wAWAAAA+gMBDwkAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjMAAAD6+wAsAAAAAQAAAAAjAAAA+gAOAAAAcwB0AHkAbABlAC4AcgBvAHQAYQB0AGkAbwBuAPsBUwAAAPr7AEwAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDQFSJAPsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAAAAA+wAAAAAACEgAAAD6AAABBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPCgAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_TYPES[35] = [];
    PRESET_SUBTYPES = PRESET_TYPES[35] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;965;wQMAAPr7ALoDAAD6AwEFAgYBDgAAAAAPBQAAABAjAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHsDAAD6+wB0AwAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACEgAAAD6AAABBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPBwAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsGzwAAAPoAAQQA+wBsAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIzAAAA+vsALAAAAAEAAAAAIwAAAPoADgAAAHMAdAB5AGwAZQAuAHIAbwB0AGEAdABpAG8AbgD7AVMAAAD6+wBMAAAAAgAAAAAaAAAA+gABAAAAMAD7AAwAAAD6AwCiSgT7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDAAAAAPsAAAAAAAbJAAAA+gABBAD7AFoAAAD6+wAWAAAA+gMBDwkAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAABskAAAD6AAEEAPsAWgAAAPr7ABYAAAD6AwEPCgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFfAAAA+vsAWAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAAAAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AdwD7AAAAAAA=";
    PRESET_TYPES[37] = [];
    PRESET_SUBTYPES = PRESET_TYPES[37] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1055;GwQAAPr7ABQEAAD6AwEFAgYBDgAAAAAPBQAAABAlAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BNUDAAD6+wDOAwAABQAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACEgAAAD6AAABBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsG1QAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AWsAAAD6+wBkAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbkAAAA+gABBAD7AF0AAAD6+wAZAAAA+gMBDKCGAQAPCQAAABMDAAAAOQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBdwAAAPr7AHAAAAACAAAAACoAAAD6AAEAAAAwAPsAHAAAAPoBCAAAACMAcABwAHQAXwB5ACsAMQD7AAAAAAAAOAAAAPoABgAAADEAMAAwADAAMAAwAPsAIAAAAPoBCgAAACMAcABwAHQAXwB5AC0ALgAwADMA+wAAAAAABgIBAAD6AAEEAPsAfwAAAPr7ADsAAAD6AKCGAQADAQ8KAAAAEwMAAAAxADAAMAD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAAOQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBcwAAAPr7AGwAAAACAAAAAC4AAAD6AAEAAAAwAPsAIAAAAPoBCgAAACMAcABwAHQAXwB5AC0ALgAwADMA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAA=";
    PRESET_TYPES[38] = [];
    PRESET_SUBTYPES = PRESET_TYPES[38] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1740;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";
    PRESET_TYPES[41] = [];
    PRESET_SUBTYPES = PRESET_TYPES[41] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1441;nQUAAPr7AJYFAAD6AwEFAgYBDgAAAAAPBQAAABApAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDECcAAPsESQUAAPr7AEIFAAAGAAAADa4AAAD6+wCIAAAA+vsALgAAAPoDAQ8GAAAAEwEAAAAxAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEaAAAA+gEHAAAAdgBpAHMAaQBiAGwAZQD7AAAAAAAGDAEAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wGkAAAA+vsAnQAAAAMAAAAAJgAAAPoAAQAAADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAAADQAAAD6AAUAAAA1ADAAMAAwADAA+wAeAAAA+gEJAAAAIwBwAHAAdABfAHgAKwAuADEA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAG0wAAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFrAAAA+vsAZAAAAAIAAAAAJgAAAPoAAQAAADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAAGFAEAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wGsAAAA+vsApQAAAAMAAAAALAAAAPoAAQAAADAA+wAeAAAA+gEJAAAAIwBwAHAAdABfAGgALwAxADAA+wAAAAAAADYAAAD6AAUAAAA1ADAAMAAwADAA+wAgAAAA+gEKAAAAIwBwAHAAdABfAGgAKwAuADAAMQD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAYUAQAA+gABBAD7AFgAAAD6+wAUAAAA+gMBDwoAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7AawAAAD6+wClAAAAAwAAAAAsAAAA+gABAAAAMAD7AB4AAAD6AQkAAAAjAHAAcAB0AF8AdwAvADEAMAD7AAAAAAAANgAAAPoABQAAADUAMAAwADAAMAD7ACAAAAD6AQoAAAAjAHAAcAB0AF8AdwArAC4AMAAxAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAACGsAAAD6AAABBAAAAGYAYQBkAGUA+wBVAAAA+vsANwAAAPoPCwAAABMDAAAANQAwADAAFxAAAAAwACwAMAA7ACAALgA1ACwAIAAxADsAIAAxACwAIAAxAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[42] = [];
    PRESET_SUBTYPES = PRESET_TYPES[42] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;783;CwMAAPr7AAQDAAD6AwEFAgYBDgAAAAAPBQAAABAqAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BMUCAAD6+wC+AgAABAAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACEgAAAD6AAABBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsG1QAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AWsAAAD6+wBkAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbbAAAA+gABBAD7AFoAAAD6+wAWAAAA+gMBDwkAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBcQAAAPr7AGoAAAACAAAAACwAAAD6AAEAAAAwAPsAHgAAAPoBCQAAACMAcABwAHQAXwB5ACsALgAxAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_TYPES[43] = [];
    PRESET_SUBTYPES = PRESET_TYPES[43] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;3773;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";
    PRESET_TYPES[45] = [];
    PRESET_SUBTYPES = PRESET_TYPES[45] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;798;GgMAAPr7ABMDAAD6AwEFAgYBDgAAAAAPBQAAABAtAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BNQCAAD6+wDNAgAABAAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACEgAAAD6AAABBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPBwAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsG6gAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7AYAAAAD6+wB5AAAAAgAAAABHAAAA+gABAAAAMAABFAAAACMAcABwAHQAXwB3ACoAcwBpAG4AKAAyAC4ANQAqAHAAaQAqACQAKQD7AAwAAAD6AwAAAAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDoIYBAPsAAAAAAAbVAAAA+gABBAD7AFoAAAD6+wAWAAAA+gMBDwkAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAA";
    PRESET_TYPES[47] = [];
    PRESET_SUBTYPES = PRESET_TYPES[47] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;783;CwMAAPr7AAQDAAD6AwEFAgYBDgAAAAAPBQAAABAvAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BMUCAAD6+wC+AgAABAAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAACEgAAAD6AAABBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsG1QAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8IAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AWsAAAD6+wBkAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB4APsAAAAAAAbbAAAA+gABBAD7AFoAAAD6+wAWAAAA+gMBDwkAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBcQAAAPr7AGoAAAACAAAAACwAAAD6AAEAAAAwAPsAHgAAAPoBCQAAACMAcABwAHQAXwB5AC0ALgAxAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_TYPES[49] = [];
    PRESET_SUBTYPES = PRESET_TYPES[49] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;962;vgMAAPr7ALcDAAD6AwEFAgYBDKCGAQAOAAAAAA8FAAAAEDEAAAARAAAAAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsEcwMAAPr7AGwDAAAFAAAADa4AAAD6+wCIAAAA+vsALgAAAPoDAQ8GAAAAEwEAAAAxAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEaAAAA+gEHAAAAdgBpAHMAaQBiAGwAZQD7AAAAAAAGxwAAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFfAAAA+vsAWAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAAAAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AdwD7AAAAAAAGxwAAAPoAAQQA+wBYAAAA+vsAFAAAAPoDAQ8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFfAAAA+vsAWAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAAAAA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AaAD7AAAAAAAGzQAAAPoAAQQA+wBqAAAA+vsAFAAAAPoDAQ8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCMwAAAPr7ACwAAAABAAAAACMAAAD6AA4AAABzAHQAeQBsAGUALgByAG8AdABhAHQAaQBvAG4A+wFTAAAA+vsATAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAUSUC+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAAIRgAAAPoAAAEEAAAAZgBhAGQAZQD7ADAAAAD6+wASAAAA+g8KAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_TYPES[50] = [];
    PRESET_SUBTYPES = PRESET_TYPES[50] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;788;EAMAAPr7AAkDAAD6AwEFAgYBDKCGAQAOAAAAAA8FAAAAEDIAAAARAAAAAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsExQIAAPr7AL4CAAAEAAAADa4AAAD6+wCIAAAA+vsALgAAAPoDAQ8GAAAAEwEAAAAxAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEaAAAA+gEHAAAAdgBpAHMAaQBiAGwAZQD7AAAAAAAG2wAAAPoAAQQA+wBaAAAA+vsAFgAAAPoDAQ8HAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7AXEAAAD6+wBqAAAAAgAAAAAsAAAA+gABAAAAMAD7AB4AAAD6AQkAAAAjAHAAcAB0AF8AdwArAC4AMwD7AAAAAAAAMAAAAPoABgAAADEAMAAwADAAMAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwB3APsAAAAAAAbVAAAA+gABBAD7AFoAAAD6+wAWAAAA+gMBDwgAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBawAAAPr7AGQAAAACAAAAACYAAAD6AAEAAAAwAPsAGAAAAPoBBgAAACMAcABwAHQAXwBoAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAACEgAAAD6AAABBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPCQAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPs=";
    PRESET_TYPES[52] = [];
    PRESET_SUBTYPES = PRESET_TYPES[52] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1077;MQQAAPr7ACoEAAD6AwEFAgYBDgAAAAAPBQAAABA0AAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BOsDAAD6+wDkAwAABAAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAAC3IAAAD6ApDQAwADkNADAASghgEABaCGAQD7AFcAAAD6+wA5AAAA+gMBDFDDAAAPBwAAABMEAAAAMQAwADAAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsJZAIAAPoAAQEBAt4AAABNACAALQAwAC4ANAA2ADcAMwA2ACAAMAAuADkAMgA4ADgANwAgACAAQwAgAC0AMAAuADMANwA1ADEANwAgADAALgA4ADgANQAwADgAIAAgAC0AMAAuADAAMgA1ADUAMgAgADAALgA3ADUAMgA3ADkAIAAgADAALgAwADkAMAA4ACAAMAAuADYANgA2ADEAMwAgACAAQwAgACAAMAAuADIAMAA3ADQANwAgADAALgA1ADcAOQA0ADgAIAAgADAALgAyADEANgA0ADkAIAAwAC4ANQAwADMAOQA0ACAAIAAwAC4AMgAzADEANwA3ACAAMAAuADQAMAA4ADIANQAgACAAQwAgADAALgAyADQANwAwADUAIAAwAC4AMwAxADIANQA2ACAAIAAwAC4AMgAyADEAMQA4ACAAMAAuADEANQA5ADYANAAgACAAIAAwAC4AMQA4ADIANgA0ACAAMAAuADAAOQAxADUAMgAgACAAQwAgADAALgAxADQANAAxACAAMAAuADAAMgAzADQAMQAgACAAMAAuADAAMwA4ADAAMgAgADAALgAwACAAIAAwAC4AMAAgADAALgAwACAAIAADAAAAAPsAkwAAAPr7ADkAAAD6AwEMUMMAAA8IAAAAEwQAAAAxADAAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+whIAAAA+gAAAQQAAABmAGEAZABlAPsAMgAAAPr7ABQAAAD6DwkAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_TYPES[53] = [];
    PRESET_SUBTYPES = PRESET_TYPES[53] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;747;5wIAAPr7AOACAAD6AwEFAgYBDgAAAAAPBQAAABA1AAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BKECAAD6+wCaAgAABAAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAACEYAAAD6AAABBAAAAGYAYQBkAGUA+wAwAAAA+vsAEgAAAPoPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7";
    PRESET_SUBTYPES[528] = "PPTY;v10;1155;fwQAAPr7AHgEAAD6AwEFAgYBDgAAAAAPBQAAABA1AAAAERACAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BDkEAAD6+wAyBAAABgAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAACEYAAAD6AAABBAAAAGYAYQBkAGUA+wAwAAAA+vsAEgAAAPoPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7BscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDUMMAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHgA+wAAAAAABscAAAD6AAEEAPsAWAAAAPr7ABQAAAD6AwEPCwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBXwAAAPr7AFgAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDUMMAAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHkA+wAAAAAA";
    PRESET_TYPES[55] = [];
    PRESET_SUBTYPES = PRESET_TYPES[55] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;787;DwMAAPr7AAgDAAD6AwEFAgYBDgAAAAAPBQAAABA3AAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BMkCAAD6+wDCAgAABAAAAA2uAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGgAAAPoBBwAAAHYAaQBzAGkAYgBsAGUA+wAAAAAABt8AAAD6AAEEAPsAWgAAAPr7ABYAAAD6AwEPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wF1AAAA+vsAbgAAAAIAAAAAMAAAAPoAAQAAADAA+wAiAAAA+gELAAAAIwBwAHAAdABfAHcAKgAwAC4ANwAwAPsAAAAAAAAwAAAA+gAGAAAAMQAwADAAMAAwADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAHcA+wAAAAAABtUAAAD6AAEEAPsAWgAAAPr7ABYAAAD6AwEPCAAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFrAAAA+vsAZAAAAAIAAAAAJgAAAPoAAQAAADAA+wAYAAAA+gEGAAAAIwBwAHAAdABfAGgA+wAAAAAAADAAAAD6AAYAAAAxADAAMAAwADAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AaAD7AAAAAAAISAAAAPoAAAEEAAAAZgBhAGQAZQD7ADIAAAD6+wAUAAAA+g8JAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w==";
    PRESET_TYPES[56] = [];
    PRESET_SUBTYPES = PRESET_TYPES[56] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;937;pQMAAPr7AJ4DAAD6AwEFAgYBDgAAAAAPBQAAABA4AAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDECcAAPsEUQMAAPr7AEoDAAAFAAAADa4AAAD6+wCIAAAA+vsALgAAAPoDAQ8GAAAAEwEAAAAxAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEaAAAA+gEHAAAAdgBpAHMAaQBiAGwAZQD7AAAAAAAGqQAAAPoAAQELAAAAKAAtACMAcABwAHQAXwB3ACoAMgApAAQA+wCDAAAA+gUDAAAAUABQAFQA+wA0AAAA+gIBAwEPBwAAABMDAAAANQAwADAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsGpwAAAPoAAQENAAAAKAAjAHAAcAB0AF8AdwAqADAALgA1ADAAKQAEAPsAfQAAAPr7ADkAAAD6AgEDAQxQwwAADwgAAAATAwAAADUAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7BrMAAAD6AAECCwAAACgALQAjAHAAcAB0AF8AaAAvADIAKQADCAAAACgAIwBwAHAAdABfAHkAKQAEAPsAeAAAAPr7ADQAAAD6AwEPCQAAABMEAAAAMQAwADAAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wp8AAAA+gAAl0kB+wBwAAAA+vsANAAAAPoDAQ8KAAAAEwQAAAAxADAAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIZAAAA+vsAEgAAAAEAAAAACQAAAPoAAQAAAHIA+w==";
    ANIMATION_PRESET_CLASSES[2] = [];
    PRESET_TYPES = ANIMATION_PRESET_CLASSES[2] = [];
    PRESET_TYPES[1] = [];
    PRESET_SUBTYPES = PRESET_TYPES[1] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;262;AgEAAPr7APsAAAD6AwEFAgYCDgAAAAAPBQAAABABAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BLwAAAD6+wC1AAAAAQAAAA2sAAAA+vsAiAAAAPr7AC4AAAD6AwEPBgAAABMBAAAAMQD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[2] = [];
    PRESET_SUBTYPES = PRESET_TYPES[2] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;698;tgIAAPr7AK8CAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHACAAD6+wBpAgAAAwAAAAbPAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AWcAAAD6+wBgAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAAC4AAAD6AAYAAAAxADAAMAAwADAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAABtcAAAD6AAEEAPsAWAAAAPoBAPsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBbwAAAPr7AGgAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeQD7AAAAAAAANgAAAPoABgAAADEAMAAwADAAMAAwAPsAHgAAAPoBCQAAADAALQBwAHAAdABfAGgALwAyAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCAAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[2] = "PPTY;v10;698;tgIAAPr7AK8CAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHACAAD6+wBpAgAAAwAAAAbXAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAAxACsAcABwAHQAXwB3AC8AMgD7AAAAAAAGzwAAAPoAAQQA+wBYAAAA+gEA+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCAAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[3] = "PPTY;v10;706;vgIAAPr7ALcCAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQMAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHgCAAD6+wBxAgAAAwAAAAbXAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAAxACsAcABwAHQAXwB3AC8AMgD7AAAAAAAG1wAAAPoAAQQA+wBYAAAA+gEA+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFvAAAA+vsAaAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA2AAAA+gAGAAAAMQAwADAAMAAwADAA+wAeAAAA+gEJAAAAMAAtAHAAcAB0AF8AaAAvADIA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[4] = "PPTY;v10;698;tgIAAPr7AK8CAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHACAAD6+wBpAgAAAwAAAAbPAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AWcAAAD6+wBgAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAAC4AAAD6AAYAAAAxADAAMAAwADAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAABtcAAAD6AAEEAPsAWAAAAPoBAPsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBbwAAAPr7AGgAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeQD7AAAAAAAANgAAAPoABgAAADEAMAAwADAAMAAwAPsAHgAAAPoBCQAAADEAKwBwAHAAdABfAGgALwAyAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCAAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[6] = "PPTY;v10;706;vgIAAPr7ALcCAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQYAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHgCAAD6+wBxAgAAAwAAAAbXAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAAxACsAcABwAHQAXwB3AC8AMgD7AAAAAAAG1wAAAPoAAQQA+wBYAAAA+gEA+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFvAAAA+vsAaAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA2AAAA+gAGAAAAMQAwADAAMAAwADAA+wAeAAAA+gEJAAAAMQArAHAAcAB0AF8AaAAvADIA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[8] = "PPTY;v10;698;tgIAAPr7AK8CAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHACAAD6+wBpAgAAAwAAAAbXAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAAwAC0AcABwAHQAXwB3AC8AMgD7AAAAAAAGzwAAAPoAAQQA+wBYAAAA+gEA+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCAAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[9] = "PPTY;v10;706;vgIAAPr7ALcCAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQkAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHgCAAD6+wBxAgAAAwAAAAbXAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAAwAC0AcABwAHQAXwB3AC8AMgD7AAAAAAAG1wAAAPoAAQQA+wBYAAAA+gEA+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFvAAAA+vsAaAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA2AAAA+gAGAAAAMQAwADAAMAAwADAA+wAeAAAA+gEJAAAAMAAtAHAAcAB0AF8AaAAvADIA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[12] = "PPTY;v10;706;vgIAAPr7ALcCAAD6AwEFAgYCDgAAAAAPBQAAABACAAAAEQwAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHgCAAD6+wBxAgAAAwAAAAbXAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAAwAC0AcABwAHQAXwB3AC8AMgD7AAAAAAAG1wAAAPoAAQQA+wBYAAAA+gEA+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFvAAAA+vsAaAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA2AAAA+gAGAAAAMQAwADAAMAAwADAA+wAeAAAA+gEJAAAAMQArAHAAcAB0AF8AaAAvADIA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[3] = [];
    PRESET_SUBTYPES = PRESET_TYPES[3] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;365;aQEAAPr7AGIBAAD6AwEFAgYCDgAAAAAPBQAAABADAAAAEQUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCMBAAD6+wAcAQAAAgAAAAheAAAA+gABARAAAABiAGwAaQBuAGQAcwAoAHYAZQByAHQAaQBjAGEAbAApAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[10] = "PPTY;v10;369;bQEAAPr7AGYBAAD6AwEFAgYCDgAAAAAPBQAAABADAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCcBAAD6+wAgAQAAAgAAAAhiAAAA+gABARIAAABiAGwAaQBuAGQAcwAoAGgAbwByAGkAegBvAG4AdABhAGwAKQD7ADAAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwcAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[4] = [];
    PRESET_SUBTYPES = PRESET_TYPES[4] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;351;WwEAAPr7AFQBAAD6AwEFAgYCDgAAAAAPBQAAABAEAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBUBAAD6+wAOAQAAAgAAAAhOAAAA+gABAQcAAABiAG8AeAAoAGkAbgApAPsAMgAAAPr7ABQAAAD6DwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbIAAAD6+wCOAAAA+vsANAAAAPoDAQ8HAAAAEwEAAAAxAPsAHwAAAPr7ABgAAAABAAAAAA8AAAD6AwQAAAAxADkAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[32] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAEAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhQAAAA+gABAQgAAABiAG8AeAAoAG8AdQB0ACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[5] = [];
    PRESET_SUBTYPES = PRESET_TYPES[5] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;369;bQEAAPr7AGYBAAD6AwEFAgYCDgAAAAAPBQAAABAFAAAAEQUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCcBAAD6+wAgAQAAAgAAAAhiAAAA+gABARIAAABjAGgAZQBjAGsAZQByAGIAbwBhAHIAZAAoAGQAbwB3AG4AKQD7ADAAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwcAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[10] = "PPTY;v10;373;cQEAAPr7AGoBAAD6AwEFAgYCDgAAAAAPBQAAABAFAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCsBAAD6+wAkAQAAAgAAAAhmAAAA+gABARQAAABjAGgAZQBjAGsAZQByAGIAbwBhAHIAZAAoAGEAYwByAG8AcwBzACkA+wAwAAAA+vsAEgAAAPoPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8HAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[6] = [];
    PRESET_SUBTYPES = PRESET_TYPES[6] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;357;YQEAAPr7AFoBAAD6AwEFAgYCDgAAAAAPBQAAABAGAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBsBAAD6+wAUAQAAAgAAAAhUAAAA+gABAQoAAABjAGkAcgBjAGwAZQAoAGkAbgApAPsAMgAAAPr7ABQAAAD6DwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbIAAAD6+wCOAAAA+vsANAAAAPoDAQ8HAAAAEwEAAAAxAPsAHwAAAPr7ABgAAAABAAAAAA8AAAD6AwQAAAAxADkAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[32] = "PPTY;v10;359;YwEAAPr7AFwBAAD6AwEFAgYCDgAAAAAPBQAAABAGAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB0BAAD6+wAWAQAAAgAAAAhWAAAA+gABAQsAAABjAGkAcgBjAGwAZQAoAG8AdQB0ACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[8] = [];
    PRESET_SUBTYPES = PRESET_TYPES[8] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;359;YwEAAPr7AFwBAAD6AwEFAgYCDgAAAAAPBQAAABAIAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB0BAAD6+wAWAQAAAgAAAAhWAAAA+gABAQsAAABkAGkAYQBtAG8AbgBkACgAaQBuACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[32] = "PPTY;v10;361;ZQEAAPr7AF4BAAD6AwEFAgYCDgAAAAAPBQAAABAIAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB8BAAD6+wAYAQAAAgAAAAhYAAAA+gABAQwAAABkAGkAYQBtAG8AbgBkACgAbwB1AHQAKQD7ADIAAAD6+wAUAAAA+g8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2yAAAA+vsAjgAAAPr7ADQAAAD6AwEPBwAAABMBAAAAMQD7AB8AAAD6+wAYAAAAAQAAAAAPAAAA+gMEAAAAMQA5ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[9] = [];
    PRESET_SUBTYPES = PRESET_TYPES[9] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYCDgAAAAAPBQAAABAJAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAAhOAAAA+gABAQgAAABkAGkAcwBzAG8AbAB2AGUA+wAwAAAA+vsAEgAAAPoPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8HAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[10] = [];
    PRESET_SUBTYPES = PRESET_TYPES[10] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;341;UQEAAPr7AEoBAAD6AwEFAgYCDgAAAAAPBQAAABAKAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BAsBAAD6+wAEAQAAAgAAAAhGAAAA+gABAQQAAABmAGEAZABlAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[12] = [];
    PRESET_SUBTYPES = PRESET_TYPES[12] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;597;UQIAAPr7AEoCAAD6AwEFAgYCDgAAAAAPBQAAABAMAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BAsCAAD6+wAEAgAAAwAAAAbzAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeQD7AYsAAAD6+wCEAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAAAUAAAAPoABgAAADEAMAAwADAAMAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB5AC0AIwBwAHAAdABfAGgAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAACE4AAAD6AAEBCAAAAHcAaQBwAGUAKAB1AHAAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwgAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[2] = "PPTY;v10;603;VwIAAPr7AFACAAD6AwEFAgYCDgAAAAAPBQAAABAMAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBECAAD6+wAKAgAAAwAAAAbzAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AYsAAAD6+wCEAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAAUAAAAPoABgAAADEAMAAwADAAMAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB4ACsAIwBwAHAAdABfAHcAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAACFQAAAD6AAEBCwAAAHcAaQBwAGUAKAByAGkAZwBoAHQAKQD7ADAAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwgAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[4] = "PPTY;v10;601;VQIAAPr7AE4CAAD6AwEFAgYCDgAAAAAPBQAAABAMAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA8CAAD6+wAIAgAAAwAAAAbzAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeQD7AYsAAAD6+wCEAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeQD7AAAAAAAAUAAAAPoABgAAADEAMAAwADAAMAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB5ACsAIwBwAHAAdABfAGgAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAACFIAAAD6AAEBCgAAAHcAaQBwAGUAKABkAG8AdwBuACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[8] = "PPTY;v10;601;VQIAAPr7AE4CAAD6AwEFAgYCDgAAAAAPBQAAABAMAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA8CAAD6+wAIAgAAAwAAAAbzAAAA+gABBAD7AFgAAAD6AQD7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AYsAAAD6+wCEAAAAAgAAAAAmAAAA+gABAAAAMAD7ABgAAAD6AQYAAAAjAHAAcAB0AF8AeAD7AAAAAAAAUAAAAPoABgAAADEAMAAwADAAMAAwAPsAOAAAAPoBFgAAACMAcABwAHQAXwB4AC0AIwBwAHAAdABfAHcAKgAxAC4AMQAyADUAMAAwADAA+wAAAAAACFIAAAD6AAEBCgAAAHcAaQBwAGUAKABsAGUAZgB0ACkA+wAwAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[13] = [];
    PRESET_SUBTYPES = PRESET_TYPES[13] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABANAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhQAAAA+gABAQgAAABwAGwAdQBzACgAaQBuACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[32] = "PPTY;v10;355;XwEAAPr7AFgBAAD6AwEFAgYCDgAAAAAPBQAAABANAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBkBAAD6+wASAQAAAgAAAAhSAAAA+gABAQkAAABwAGwAdQBzACgAbwB1AHQAKQD7ADIAAAD6+wAUAAAA+g8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2yAAAA+vsAjgAAAPr7ADQAAAD6AwEPBwAAABMBAAAAMQD7AB8AAAD6+wAYAAAAAQAAAAAPAAAA+gMEAAAAMQA5ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[14] = [];
    PRESET_SUBTYPES = PRESET_TYPES[14] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;371;bwEAAPr7AGgBAAD6AwEFAgYCDgAAAAAPBQAAABAOAAAAEQUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCkBAAD6+wAiAQAAAgAAAAhkAAAA+gABARMAAAByAGEAbgBkAG8AbQBiAGEAcgAoAHYAZQByAHQAaQBjAGEAbAApAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[10] = "PPTY;v10;375;cwEAAPr7AGwBAAD6AwEFAgYCDgAAAAAPBQAAABAOAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BC0BAAD6+wAmAQAAAgAAAAhoAAAA+gABARUAAAByAGEAbgBkAG8AbQBiAGEAcgAoAGgAbwByAGkAegBvAG4AdABhAGwAKQD7ADAAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwcAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[15] = [];
    PRESET_SUBTYPES = PRESET_TYPES[15] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;6232;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";
    PRESET_TYPES[16] = [];
    PRESET_SUBTYPES = PRESET_TYPES[16] = [];
    PRESET_SUBTYPES[21] = "PPTY;v10;365;aQEAAPr7AGIBAAD6AwEFAgYCDgAAAAAPBQAAABAQAAAAERUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCMBAAD6+wAcAQAAAgAAAAheAAAA+gABARAAAABiAGEAcgBuACgAaQBuAFYAZQByAHQAaQBjAGEAbAApAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[26] = "PPTY;v10;369;bQEAAPr7AGYBAAD6AwEFAgYCDgAAAAAPBQAAABAQAAAAERoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCcBAAD6+wAgAQAAAgAAAAhiAAAA+gABARIAAABiAGEAcgBuACgAaQBuAEgAbwByAGkAegBvAG4AdABhAGwAKQD7ADAAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwcAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[37] = "PPTY;v10;367;awEAAPr7AGQBAAD6AwEFAgYCDgAAAAAPBQAAABAQAAAAESUAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCUBAAD6+wAeAQAAAgAAAAhgAAAA+gABAREAAABiAGEAcgBuACgAbwB1AHQAVgBlAHIAdABpAGMAYQBsACkA+wAwAAAA+vsAEgAAAPoPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8HAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[42] = "PPTY;v10;371;bwEAAPr7AGgBAAD6AwEFAgYCDgAAAAAPBQAAABAQAAAAESoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCkBAAD6+wAiAQAAAgAAAAhkAAAA+gABARMAAABiAGEAcgBuACgAbwB1AHQASABvAHIAaQB6AG8AbgB0AGEAbAApAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[17] = [];
    PRESET_SUBTYPES = PRESET_TYPES[17] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;1112;VAQAAPr7AE0EAAD6AwEFAgYCDgAAAAAPBQAAABARAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA4EAAD6+wAHBAAABQAAAAbNAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB4APsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwB4APsAAAAAAAbdAAAA+gABBAD7AFYAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wF3AAAA+vsAcAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA+AAAA+gAGAAAAMQAwADAAMAAwADAA+wAmAAAA+gENAAAAcABwAHQAXwB5AC0AcABwAHQAXwBoAC8AMgD7AAAAAAAGzQAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AdwD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AdwD7AAAAAAAGwwAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AaAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDAAAAAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCgAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[2] = "PPTY;v10;1112;VAQAAPr7AE0EAAD6AwEFAgYCDgAAAAAPBQAAABARAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA4EAAD6+wAHBAAABQAAAAbdAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wF3AAAA+vsAcAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB4APsAAAAAAAA+AAAA+gAGAAAAMQAwADAAMAAwADAA+wAmAAAA+gENAAAAcABwAHQAXwB4ACsAcABwAHQAXwB3AC8AMgD7AAAAAAAGzQAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeQD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeQD7AAAAAAAGwwAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AdwD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDAAAAAPsAAAAAAAbNAAAA+gABBAD7AFYAAAD6+wASAAAA+g8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCgAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[4] = "PPTY;v10;1112;VAQAAPr7AE0EAAD6AwEFAgYCDgAAAAAPBQAAABARAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA4EAAD6+wAHBAAABQAAAAbNAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB4APsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwB4APsAAAAAAAbdAAAA+gABBAD7AFYAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wF3AAAA+vsAcAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA+AAAA+gAGAAAAMQAwADAAMAAwADAA+wAmAAAA+gENAAAAcABwAHQAXwB5ACsAcABwAHQAXwBoAC8AMgD7AAAAAAAGzQAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AdwD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AdwD7AAAAAAAGwwAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AaAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDAAAAAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCgAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[8] = "PPTY;v10;1112;VAQAAPr7AE0EAAD6AwEFAgYCDgAAAAAPBQAAABARAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BA4EAAD6+wAHBAAABQAAAAbdAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wF3AAAA+vsAcAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB4APsAAAAAAAA+AAAA+gAGAAAAMQAwADAAMAAwADAA+wAmAAAA+gENAAAAcABwAHQAXwB4AC0AcABwAHQAXwB3AC8AMgD7AAAAAAAGzQAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeQD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeQD7AAAAAAAGwwAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AdwD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDAAAAAPsAAAAAAAbNAAAA+gABBAD7AFYAAAD6+wASAAAA+g8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCgAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[10] = "PPTY;v10;676;oAIAAPr7AJkCAAD6AwEFAgYCDgAAAAAPBQAAABARAAAAEQoAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BFoCAAD6+wBTAgAAAwAAAAbDAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAAAAA+wAAAAAABs0AAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AWcAAAD6+wBgAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAAC4AAAD6AAYAAAAxADAAMAAwADAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[18] = [];
    PRESET_SUBTYPES = PRESET_TYPES[18] = [];
    PRESET_SUBTYPES[3] = "PPTY;v10;363;ZwEAAPr7AGABAAD6AwEFAgYCDgAAAAAPBQAAABASAAAAEQMAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCEBAAD6+wAaAQAAAgAAAAhcAAAA+gABAQ8AAABzAHQAcgBpAHAAcwAoAHUAcABSAGkAZwBoAHQAKQD7ADAAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwcAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[6] = "PPTY;v10;367;awEAAPr7AGQBAAD6AwEFAgYCDgAAAAAPBQAAABASAAAAEQYAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCUBAAD6+wAeAQAAAgAAAAhgAAAA+gABAREAAABzAHQAcgBpAHAAcwAoAGQAbwB3AG4AUgBpAGcAaAB0ACkA+wAwAAAA+vsAEgAAAPoPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8HAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[9] = "PPTY;v10;361;ZQEAAPr7AF4BAAD6AwEFAgYCDgAAAAAPBQAAABASAAAAEQkAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BB8BAAD6+wAYAQAAAgAAAAhaAAAA+gABAQ4AAABzAHQAcgBpAHAAcwAoAHUAcABMAGUAZgB0ACkA+wAwAAAA+vsAEgAAAPoPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8HAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[12] = "PPTY;v10;365;aQEAAPr7AGIBAAD6AwEFAgYCDgAAAAAPBQAAABASAAAAEQwAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCMBAAD6+wAcAQAAAgAAAAheAAAA+gABARAAAABzAHQAcgBpAHAAcwAoAGQAbwB3AG4ATABlAGYAdAApAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[19] = [];
    PRESET_SUBTYPES = PRESET_TYPES[19] = [];
    PRESET_SUBTYPES[5] = "PPTY;v10;1755;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";
    PRESET_SUBTYPES[10] = "PPTY;v10;1755;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";
    PRESET_TYPES[20] = [];
    PRESET_SUBTYPES = PRESET_TYPES[20] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;347;VwEAAPr7AFABAAD6AwEFAgYCDgAAAAAPBQAAABAUAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBEBAAD6+wAKAQAAAgAAAAhKAAAA+gABAQUAAAB3AGUAZABnAGUA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[21] = [];
    PRESET_SUBTYPES = PRESET_TYPES[21] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAVAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhQAAAA+gABAQgAAAB3AGgAZQBlAGwAKAAxACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[2] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAVAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhQAAAA+gABAQgAAAB3AGgAZQBlAGwAKAAyACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[3] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAVAAAAEQMAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhQAAAA+gABAQgAAAB3AGgAZQBlAGwAKAAzACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[4] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAVAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhQAAAA+gABAQgAAAB3AGgAZQBlAGwAKAA0ACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[8] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAVAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhQAAAA+gABAQgAAAB3AGgAZQBlAGwAKAA4ACkA+wAyAAAA+vsAFAAAAPoPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwcAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[22] = [];
    PRESET_SUBTYPES = PRESET_TYPES[22] = [];
    PRESET_SUBTYPES[1] = "PPTY;v10;349;WQEAAPr7AFIBAAD6AwEFAgYCDgAAAAAPBQAAABAWAAAAEQEAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBMBAAD6+wAMAQAAAgAAAAhOAAAA+gABAQgAAAB3AGkAcABlACgAdQBwACkA+wAwAAAA+vsAEgAAAPoPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8HAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[2] = "PPTY;v10;355;XwEAAPr7AFgBAAD6AwEFAgYCDgAAAAAPBQAAABAWAAAAEQIAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBkBAAD6+wASAQAAAgAAAAhUAAAA+gABAQsAAAB3AGkAcABlACgAcgBpAGcAaAB0ACkA+wAwAAAA+vsAEgAAAPoPBgAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8HAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[4] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAWAAAAEQQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhSAAAA+gABAQoAAAB3AGkAcABlACgAZABvAHcAbgApAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_SUBTYPES[8] = "PPTY;v10;353;XQEAAPr7AFYBAAD6AwEFAgYCDgAAAAAPBQAAABAWAAAAEQgAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BBcBAAD6+wAQAQAAAgAAAAhSAAAA+gABAQoAAAB3AGkAcABlACgAbABlAGYAdAApAPsAMAAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPBwAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAANAA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[23] = [];
    PRESET_SUBTYPES = PRESET_TYPES[23] = [];
    PRESET_SUBTYPES[16] = "PPTY;v10;694;sgIAAPr7AKsCAAD6AwEFAgYCDgAAAAAPBQAAABAXAAAAERAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BGwCAAD6+wBlAgAAAwAAAAbRAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFrAAAA+vsAZAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAAyAAAA+gAGAAAAMQAwADAAMAAwADAA+wAaAAAA+gEHAAAANAAqAHAAcAB0AF8AdwD7AAAAAAAG0QAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBawAAAPr7AGQAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AaAD7AAAAAAAAMgAAAPoABgAAADEAMAAwADAAMAAwAPsAGgAAAPoBBwAAADQAKgBwAHAAdABfAGgA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8IAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[20] = "PPTY;v10;1138;bgQAAPr7AGcEAAD6AwEFAgYCDgAAAAAPBQAAABAXAAAAERQAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCgEAAD6+wAhBAAABAAAAAYbAQAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wG1AAAA+vsArgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAB8AAAA+gAGAAAAMQAwADAAMAAwADAA+wBkAAAA+gEsAAAAKAA2ACoAbQBpAG4AKABtAGEAeAAoAHAAcAB0AF8AdwAqAHAAcAB0AF8AaAAsAC4AMwApACwAMQApAC0ANwAuADQAKQAvAC0ALgA3ACoAcABwAHQAXwB3APsAAAAAAAYbAQAA+gABBAD7AFYAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wG1AAAA+vsArgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAAB8AAAA+gAGAAAAMQAwADAAMAAwADAA+wBkAAAA+gEsAAAAKAA2ACoAbQBpAG4AKABtAGEAeAAoAHAAcAB0AF8AdwAqAHAAcAB0AF8AaAAsAC4AMwApACwAMQApAC0ANwAuADQAKQAvAC0ALgA3ACoAcABwAHQAXwBoAPsAAAAAAAYjAQAA+gABBAD7AFYAAAD6+wASAAAA+g8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wG9AAAA+vsAtgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAACEAAAA+gAGAAAAMQAwADAAMAAwADAA+wBsAAAA+gEwAAAAMQArACgANgAqAG0AaQBuACgAbQBhAHgAKABwAHAAdABfAHcAKgBwAHAAdABfAGgALAAuADMAKQAsADEAKQAtADcALgA0ACkALwAtAC4ANwAqAHAAcAB0AF8AaAAvADIA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8JAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_SUBTYPES[32] = "PPTY;v10;666;lgIAAPr7AI8CAAD6AwEFAgYCDgAAAAAPBQAAABAXAAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BFACAAD6+wBJAgAAAwAAAAbDAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAAAAA+wAAAAAABsMAAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AV0AAAD6+wBWAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAANsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwgAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[272] = "PPTY;v10;702;ugIAAPr7ALMCAAD6AwEFAgYCDgAAAAAPBQAAABAXAAAAERABAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHQCAAD6+wBtAgAAAwAAAAbVAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFvAAAA+vsAaAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAA2AAAA+gAGAAAAMQAwADAAMAAwADAA+wAeAAAA+gEJAAAANAAvADMAKgBwAHAAdABfAHcA+wAAAAAABtUAAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAA0AC8AMwAqAHAAcAB0AF8AaAD7AAAAAAANsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwgAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[288] = "PPTY;v10;702;ugIAAPr7ALMCAAD6AwEFAgYCDgAAAAAPBQAAABAXAAAAESABAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHQCAAD6+wBtAgAAAwAAAAbVAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFvAAAA+vsAaAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAA2AAAA+gAGAAAAMQAwADAAMAAwADAA+wAeAAAA+gEJAAAAMgAvADMAKgBwAHAAdABfAHcA+wAAAAAABtUAAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AW8AAAD6+wBoAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAADYAAAD6AAYAAAAxADAAMAAwADAAMAD7AB4AAAD6AQkAAAAyAC8AMwAqAHAAcAB0AF8AaAD7AAAAAAANsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwgAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[544] = "PPTY;v10;1066;JgQAAPr7AB8EAAD6AwEFAgYCDgAAAAAPBQAAABAXAAAAESACAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BOADAAD6+wDZAwAABQAAAAbDAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAAAAA+wAAAAAABsMAAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AV0AAAD6+wBWAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAAGwwAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCAAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDUMMAAPsAAAAAAAbDAAAA+gABBAD7AFYAAAD6+wASAAAA+g8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gNQwwAA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8KAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[25] = [];
    PRESET_SUBTYPES = PRESET_TYPES[25] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;2108;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";
    PRESET_TYPES[26] = [];
    PRESET_SUBTYPES = PRESET_TYPES[26] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;6084;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";
    PRESET_TYPES[28] = [];
    PRESET_SUBTYPES = PRESET_TYPES[28] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;706;vgIAAPr7ALcCAAD6AwEFAgYCDgAAAAAPBQAAABAcAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHgCAAD6+wBxAgAAAwAAAAbRAAAA+gABBAD7AFoAAAD6+wAWAAAA+g8GAAAAEwUAAAAxADUAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAG2QAAAPoAAQQA+wBaAAAA+vsAFgAAAPoPBwAAABMFAAAAMQA1ADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeQD7AW8AAAD6+wBoAAAAAgAAAAAoAAAA+gABAAAAMAD7ABoAAAD6AQcAAABwAHAAdABfAHkALQAxAPsAAAAAAAAyAAAA+gAGAAAAMQAwADAAMAAwADAA+wAaAAAA+gEHAAAAcABwAHQAXwB5ACsAMQD7AAAAAAANtAAAAPr7AJAAAAD6+wA2AAAA+gMBDwgAAAATAQAAADEA+wAhAAAA+vsAGgAAAAEAAAAAEQAAAPoDBQAAADEANAA5ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[30] = [];
    PRESET_SUBTYPES = PRESET_TYPES[30] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1607;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";
    PRESET_TYPES[31] = [];
    PRESET_SUBTYPES = PRESET_TYPES[31] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;957;uQMAAPr7ALIDAAD6AwEFAgYCDgAAAAAPBQAAABAfAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHMDAAD6+wBsAwAABQAAAAbFAAAA+gABBAD7AFgAAAD6+wAUAAAA+g8GAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7AV0AAAD6+wBWAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHcA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAAGxQAAAPoAAQQA+wBYAAAA+vsAFAAAAPoPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAAAAA+wAAAAAABs0AAAD6AAEEAPsAagAAAPr7ABQAAAD6DwgAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjMAAAD6+wAsAAAAAQAAAAAjAAAA+gAOAAAAcwB0AHkAbABlAC4AcgBvAHQAYQB0AGkAbwBuAPsBUwAAAPr7AEwAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gNAVIkA+wAAAAAACEgAAAD6AAEBBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPCQAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwoAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADkAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[35] = [];
    PRESET_SUBTYPES = PRESET_TYPES[35] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;959;uwMAAPr7ALQDAAD6AwEFAgYCDgAAAAAPBQAAABAjAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BHUDAAD6+wBuAwAABQAAAAhIAAAA+gABAQQAAABmAGEAZABlAPsAMgAAAPr7ABQAAAD6DwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Bs0AAAD6AAEEAPsAagAAAPr7ABQAAAD6DwcAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjMAAAD6+wAsAAAAAQAAAAAjAAAA+gAOAAAAcwB0AHkAbABlAC4AcgBvAHQAYQB0AGkAbwBuAPsBUwAAAPr7AEwAAAACAAAAABoAAAD6AAEAAAAwAPsADAAAAPoDAAAAAPsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAokoE+wAAAAAABsUAAAD6AAEEAPsAWAAAAPr7ABQAAAD6DwgAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AaAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDAAAAAPsAAAAAAAbFAAAA+gABBAD7AFgAAAD6+wAUAAAA+g8JAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7AV0AAAD6+wBWAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHcA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAANsgAAAPr7AI4AAAD6+wA0AAAA+gMBDwoAAAATAQAAADEA+wAfAAAA+vsAGAAAAAEAAAAADwAAAPoDBAAAADEAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[37] = [];
    PRESET_SUBTYPES = PRESET_TYPES[37] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1031;AwQAAPr7APwDAAD6AwEFAgYCDgAAAAAPBQAAABAlAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BL0DAAD6+wC2AwAABQAAAAhIAAAA+gABAQQAAABmAGEAZABlAPsAMgAAAPr7ABQAAAD6DwYAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Bs8AAAD6AAEEAPsAWAAAAPr7ABQAAAD6DwcAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAG2gAAAPoAAQQA+wBbAAAA+vsAFwAAAPoMoIYBAA8IAAAAEwMAAAAxADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFvAAAA+vsAaAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA2AAAA+gAGAAAAMQAwADAAMAAwADAA+wAeAAAA+gEJAAAAcABwAHQAXwB5AC0ALgAwADMA+wAAAAAABvgAAAD6AAEEAPsAfQAAAPr7ADkAAAD6AKCGAQAPCQAAABMDAAAAOQAwADAA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADEAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeQD7AWsAAAD6+wBkAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHkA+wAAAAAAADIAAAD6AAYAAAAxADAAMAAwADAAMAD7ABoAAAD6AQcAAABwAHAAdABfAHkAKwAxAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCgAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[38] = [];
    PRESET_SUBTYPES = PRESET_TYPES[38] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;771;/wIAAPr7APgCAAD6AFDDAAADAQUCBgIOAAAAAA8FAAAAECYAAAARAAAAAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsDCQAAAPoAAQNQwwAA+wSmAgAA+vsAnwIAAAMAAAAG6wAAAPoAAQQA+wCIAAAA+vsAMgAAAPoPBgAAABMEAAAAMQAwADAAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCMwAAAPr7ACwAAAABAAAAACMAAAD6AA4AAABzAHQAeQBsAGUALgByAG8AdABhAHQAaQBvAG4A+wFTAAAA+vsATAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAAAAA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AyCqRAD7AAAAAAAG8QAAAPoAAQQA+wB2AAAA+vsAMgAAAPoPBwAAABMEAAAAMQAwADAAMAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFrAAAA+vsAZAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAAyAAAA+gAGAAAAMQAwADAAMAAwADAA+wAaAAAA+gEHAAAAcABwAHQAXwB5ACsAMQD7AAAAAAANsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwgAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADkAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[41] = [];
    PRESET_SUBTYPES = PRESET_TYPES[41] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1413;gQUAAPr7AHoFAAD6AwEFAgYCDgAAAAAPBQAAABApAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDECcAAPsELQUAAPr7ACYFAAAGAAAABgQBAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AZ4AAAD6+wCXAAAAAwAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADIAAAD6AAUAAAA1ADAAMAAwADAA+wAcAAAA+gEIAAAAcABwAHQAXwB4ACsALgAxAPsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwB4APsAAAAAAAbNAAAA+gABBAD7AFYAAAD6+wASAAAA+g8HAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAYMAQAA+gABBAD7AFYAAAD6+wASAAAA+g8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wGmAAAA+vsAnwAAAAMAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAAA0AAAA+gAFAAAANQAwADAAMAAwAPsAHgAAAPoBCQAAAHAAcAB0AF8AaAArAC4AMAAxAPsAAAAAAAA0AAAA+gAGAAAAMQAwADAAMAAwADAA+wAcAAAA+gEIAAAAcABwAHQAXwBoAC8AMQAwAPsAAAAAAAYMAQAA+gABBAD7AFYAAAD6+wASAAAA+g8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wGmAAAA+vsAnwAAAAMAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAA0AAAA+gAFAAAANQAwADAAMAAwAPsAHgAAAPoBCQAAAHAAcAB0AF8AdwArAC4AMAAxAPsAAAAAAAA0AAAA+gAGAAAAMQAwADAAMAAwADAA+wAcAAAA+gEIAAAAcABwAHQAXwB3AC8AMQAwAPsAAAAAAAhrAAAA+gABAQQAAABmAGEAZABlAPsAVQAAAPr7ADcAAAD6DwoAAAATAwAAADUAMAAwABcQAAAAMAAsADAAOwAgAC4ANQAsACAAMAA7ACAAMQAsACAAMQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwsAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[42] = [];
    PRESET_SUBTYPES = PRESET_TYPES[42] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;773;AQMAAPr7APoCAAD6AwEFAgYCDgAAAAAPBQAAABAqAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BLsCAAD6+wC0AgAABAAAAAhIAAAA+gABAQQAAABmAGEAZABlAPsAMgAAAPr7ABQAAAD6DwYAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Bs8AAAD6AAEEAPsAWAAAAPr7ABQAAAD6DwcAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAG1QAAAPoAAQQA+wBYAAAA+vsAFAAAAPoPCAAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFtAAAA+vsAZgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA0AAAA+gAGAAAAMQAwADAAMAAwADAA+wAcAAAA+gEIAAAAcABwAHQAXwB5ACsALgAxAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCQAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[43] = [];
    PRESET_SUBTYPES = PRESET_TYPES[43] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;3747;nw4AAPr7AJgOAAD6AwEFAgYCDgAAAAAPBQAAABArAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BFkOAAD6+wBSDgAABgAAAAaZBQAA+gABBAD7AHkAAAD6+wA1AAAA+gxQwwAADwYAAAATAwAAADYAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ARAFAAD6+wAJBQAAFQAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAADgAAAD6AAQAAAA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAwADIANAAyAPsAAAAAAAA6AAAA+gAFAAAAMQAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAwADQANwA5APsAAAAAAAA6AAAA+gAFAAAAMQA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAwADcAMAA0APsAAAAAAAA6AAAA+gAFAAAAMgAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAwADkAMQAxAPsAAAAAAAA6AAAA+gAFAAAAMgA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAxADAAOQA2APsAAAAAAAA6AAAA+gAFAAAAMwAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAxADIANQA0APsAAAAAAAA6AAAA+gAFAAAAMwA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAxADMAOAAxAPsAAAAAAAA6AAAA+gAFAAAANAAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAxADQANwA0APsAAAAAAAA6AAAA+gAFAAAANAA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeAArADAALgAxADUAMwAxAPsAAAAAAAA4AAAA+gAFAAAANQAwADAAMAAwAPsAIgAAAPoBCwAAAHAAcAB0AF8AeAArADAALgAxADUANQD7AAAAAAAAOgAAAPoABQAAADUANQAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMQA1ADMAMQD7AAAAAAAAOgAAAPoABQAAADYAMAAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMQA0ADcANAD7AAAAAAAAOgAAAPoABQAAADYANQAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMQAzADgAMQD7AAAAAAAAOgAAAPoABQAAADcAMAAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMQAyADUANAD7AAAAAAAAOgAAAPoABQAAADcANQAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMQAwADkANgD7AAAAAAAAOgAAAPoABQAAADgAMAAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMAA5ADEAMQD7AAAAAAAAOgAAAPoABQAAADgANQAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMAA3ADAANAD7AAAAAAAAOgAAAPoABQAAADkAMAAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMAA0ADcAOQD7AAAAAAAAOgAAAPoABQAAADkANQAwADAAMAD7ACQAAAD6AQwAAABwAHAAdABfAHgAKwAwAC4AMAAyADQAMgD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAG7wAAAPoAAQQA+wB4AAAA+vsANAAAAPoPBwAAABMDAAAANAAwADAA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADYAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7AWcAAAD6+wBgAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAAAC4AAAD6AAYAAAAxADAAMAAwADAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHgA+wAAAAAABqEFAAD6AAEEAPsAeQAAAPr7ADUAAAD6DFDDAAAPCAAAABMDAAAANgAwADAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB5APsBGAUAAPr7ABEFAAAVAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeQD7AAAAAAAAOAAAAPoABAAAADUAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADAAMAAxADkA+wAAAAAAADoAAAD6AAUAAAAxADAAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADAAMAA3ADYA+wAAAAAAADoAAAD6AAUAAAAxADUAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADAAMQA2ADkA+wAAAAAAADoAAAD6AAUAAAAyADAAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADAAMgA5ADYA+wAAAAAAADoAAAD6AAUAAAAyADUAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADAANAA1ADQA+wAAAAAAADoAAAD6AAUAAAAzADAAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADAANgAzADkA+wAAAAAAADoAAAD6AAUAAAAzADUAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADAAOAA0ADYA+wAAAAAAADoAAAD6AAUAAAA0ADAAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADEAMAA3ADEA+wAAAAAAADoAAAD6AAUAAAA0ADUAMAAwADAA+wAkAAAA+gEMAAAAcABwAHQAXwB5ACsAMAAuADEAMwAwADcA+wAAAAAAADgAAAD6AAUAAAA1ADAAMAAwADAA+wAiAAAA+gELAAAAcABwAHQAXwB5ACsAMAAuADEANQA1APsAAAAAAAA6AAAA+gAFAAAANQA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAxADcAOQAyAPsAAAAAAAA6AAAA+gAFAAAANgAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAyADAAMgA5APsAAAAAAAA6AAAA+gAFAAAANgA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAyADIANQAzAPsAAAAAAAA6AAAA+gAFAAAANwAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAyADQANgAxAPsAAAAAAAA6AAAA+gAFAAAANwA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAyADYANAA2APsAAAAAAAA6AAAA+gAFAAAAOAAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAyADgAMAA0APsAAAAAAAA6AAAA+gAFAAAAOAA1ADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAyADkAMwAxAPsAAAAAAAA6AAAA+gAFAAAAOQAwADAAMAAwAPsAJAAAAPoBDAAAAHAAcAB0AF8AeQArADAALgAzADAAMgA0APsAAAAAAAA4AAAA+gAFAAAAOQA1ADAAMAAwAPsAIgAAAPoBCwAAAHAAcAB0AF8AeQArADAALgAzADAAOAD7AAAAAAAAOAAAAPoABgAAADEAMAAwADAAMAAwAPsAIAAAAPoBCgAAAHAAcAB0AF8AeQArADAALgAzADEA+wAAAAAABu8AAAD6AAEEAPsAeAAAAPr7ADQAAAD6DwkAAAATAwAAADQAMAAwAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA2ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAhoAAAA+gABAQQAAABmAGEAZABlAPsAUgAAAPr7ADQAAAD6DwoAAAATAwAAADEAMAAwAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA5ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwsAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADkAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[45] = [];
    PRESET_SUBTYPES = PRESET_TYPES[45] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1832;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";
    PRESET_TYPES[47] = [];
    PRESET_SUBTYPES = PRESET_TYPES[47] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;773;AQMAAPr7APoCAAD6AwEFAgYCDgAAAAAPBQAAABAvAAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BLsCAAD6+wC0AgAABAAAAAhIAAAA+gABAQQAAABmAGEAZABlAPsAMgAAAPr7ABQAAAD6DwYAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7Bs8AAAD6AAEEAPsAWAAAAPr7ABQAAAD6DwcAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBZwAAAPr7AGAAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAALgAAAPoABgAAADEAMAAwADAAMAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAG1QAAAPoAAQQA+wBYAAAA+vsAFAAAAPoPCAAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFtAAAA+vsAZgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAA0AAAA+gAGAAAAMQAwADAAMAAwADAA+wAcAAAA+gEIAAAAcABwAHQAXwB5AC0ALgAxAPsAAAAAAA2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCQAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    PRESET_TYPES[49] = [];
    PRESET_SUBTYPES = PRESET_TYPES[49] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;954;tgMAAPr7AK8DAAD6AKCGAQADAQUCBgIOAAAAAA8FAAAAEDEAAAARAAAAAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsEawMAAPr7AGQDAAAFAAAABsMAAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwYAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7AV0AAAD6+wBWAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHcA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAAGwwAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPBwAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwBoAPsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AaAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDAAAAAPsAAAAAAAbLAAAA+gABBAD7AGgAAAD6+wASAAAA+g8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCMwAAAPr7ACwAAAABAAAAACMAAAD6AA4AAABzAHQAeQBsAGUALgByAG8AdABhAHQAaQBvAG4A+wFTAAAA+vsATAAAAAIAAAAAGgAAAPoAAQAAADAA+wAMAAAA+gMAAAAA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwBRJQL7AAAAAAAIRgAAAPoAAQEEAAAAZgBhAGQAZQD7ADAAAAD6+wASAAAA+g8JAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwoAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[50] = [];
    PRESET_SUBTYPES = PRESET_TYPES[50] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;778;BgMAAPr7AP8CAAD6AKCGAQADAQUCBgIOAAAAAA8FAAAAEDIAAAARAAAAAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsEuwIAAPr7ALQCAAAEAAAABtUAAAD6AAEEAPsAWAAAAPr7ABQAAAD6DwYAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB3APsBbQAAAPr7AGYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AdwD7AAAAAAAANAAAAPoABgAAADEAMAAwADAAMAAwAPsAHAAAAPoBCAAAAHAAcAB0AF8AdwArAC4AMwD7AAAAAAAGzwAAAPoAAQQA+wBYAAAA+vsAFAAAAPoPBwAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAGgA+wFnAAAA+vsAYAAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAAAuAAAA+gAGAAAAMQAwADAAMAAwADAA+wAWAAAA+gEFAAAAcABwAHQAXwBoAPsAAAAAAAhIAAAA+gABAQQAAABmAGEAZABlAPsAMgAAAPr7ABQAAAD6DwgAAAATBAAAADEAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7DbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8JAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA5ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[52] = [];
    PRESET_SUBTYPES = PRESET_TYPES[52] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1041;DQQAAPr7AAYEAAD6AwEFAgYCDgAAAAAPBQAAABA0AAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BMcDAAD6+wDAAwAABAAAAAtwAAAA+gKghgEAA6CGAQAEkNADAAWQ0AMA+wBVAAAA+vsANwAAAPoAUMMAAA8GAAAAEwQAAAAxADAAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wlAAgAA+gABAQECzQAAAE0AIAAwAC4AMAAwADAAMAAgADAALgAwADAAMAAwACAAQwAgADAALgAwADMAOAAwADIAIAAwAC4AMAAgADAALgAxADQANAAxACAAMAAuADAAMgAzADQAMQAgADAALgAxADgAMgA2ACAAMAAuADAAOQAxADUAIABDACAAMAAuADIAMgAxADEAOAAgADAALgAxADUAOQA2ADQAIAAwAC4AMgA0ADcAMAA1ACAAMAAuADMAMQAyADUANgAgADAALgAyADMAMQA4ACAAMAAuADQAMAA4ADMAIABDACAAMAAuADIAMQA2ADQAOQAgADAALgA1ADAAMwA5ADQAIAAwAC4AMgAwADcANAA3ACAAMAAuADUANwA5ADQAOAAgADAALgAwADkAMAA4ACAAMAAuADYANgA2ADEAIABDACAALQAwAC4AMAAyADUANQAyACAAMAAuADcANQAyADcAOQAgAC0AMAAuADMANwA1ADEANwAgADAALgA4ADgANQAwADgAIAAtADAALgA0ADYANwA0ACAAMAAuADkAMgA4ADkAAwAAAAD7AJEAAAD6+wA3AAAA+gBQwwAADwcAAAATBAAAADEAMAAwADAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7CEgAAAD6AAEBBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPCAAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwkAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADkAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[53] = [];
    PRESET_SUBTYPES = PRESET_TYPES[53] = [];
    PRESET_SUBTYPES[32] = "PPTY;v10;741;4QIAAPr7ANoCAAD6AwEFAgYCDgAAAAAPBQAAABA1AAAAESAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BJsCAAD6+wCUAgAABAAAAAbDAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAAAAA+wAAAAAABsMAAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AV0AAAD6+wBWAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAAIRgAAAPoAAQEEAAAAZgBhAGQAZQD7ADAAAAD6+wASAAAA+g8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwkAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADQAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_SUBTYPES[544] = "PPTY;v10;1141;cQQAAPr7AGoEAAD6AwEFAgYCDgAAAAAPBQAAABA1AAAAESACAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BCsEAAD6+wAkBAAABgAAAAbDAAAA+gABBAD7AFYAAAD6+wASAAAA+g8GAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHcA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB3APsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gMAAAAA+wAAAAAABsMAAAD6AAEEAPsAVgAAAPr7ABIAAAD6DwcAAAATAwAAADUAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AV0AAAD6+wBWAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAACQAAAD6AAYAAAAxADAAMAAwADAAMAD7AAwAAAD6AwAAAAD7AAAAAAAIRgAAAPoAAQEEAAAAZgBhAGQAZQD7ADAAAAD6+wASAAAA+g8IAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsGwwAAAPoAAQQA+wBWAAAA+vsAEgAAAPoPCQAAABMDAAAANQAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AiEAAAD6+wAaAAAAAQAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsBXQAAAPr7AFYAAAACAAAAACQAAAD6AAEAAAAwAPsAFgAAAPoBBQAAAHAAcAB0AF8AeAD7AAAAAAAAJAAAAPoABgAAADEAMAAwADAAMAAwAPsADAAAAPoDUMMAAPsAAAAAAAbDAAAA+gABBAD7AFYAAAD6+wASAAAA+g8KAAAAEwMAAAA1ADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCIQAAAPr7ABoAAAABAAAAABEAAAD6AAUAAABwAHAAdABfAHkA+wFdAAAA+vsAVgAAAAIAAAAAJAAAAPoAAQAAADAA+wAWAAAA+gEFAAAAcABwAHQAXwB5APsAAAAAAAAkAAAA+gAGAAAAMQAwADAAMAAwADAA+wAMAAAA+gNQwwAA+wAAAAAADbAAAAD6+wCMAAAA+vsAMgAAAPoDAQ8LAAAAEwEAAAAxAPsAHQAAAPr7ABYAAAABAAAAAA0AAAD6AwMAAAA0ADkAOQD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAABAAAAACcAAAD6ABAAAABzAHQAeQBsAGUALgB2AGkAcwBpAGIAaQBsAGkAdAB5APsBGAAAAPoBBgAAAGgAaQBkAGQAZQBuAPsAAAAAAA==";
    PRESET_TYPES[55] = [];
    PRESET_SUBTYPES = PRESET_TYPES[55] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;777;BQMAAPr7AP4CAAD6AwEFAgYCDgAAAAAPBQAAABA3AAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7BL8CAAD6+wC4AgAABAAAAAbZAAAA+gABBAD7AFgAAAD6+wAUAAAA+g8GAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7AXEAAAD6+wBqAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAHcA+wAAAAAAADgAAAD6AAYAAAAxADAAMAAwADAAMAD7ACAAAAD6AQoAAABwAHAAdABfAHcAKgAwAC4ANwAwAPsAAAAAAAbPAAAA+gABBAD7AFgAAAD6+wAUAAAA+g8HAAAAEwQAAAAxADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AaAD7AWcAAAD6+wBgAAAAAgAAAAAkAAAA+gABAAAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAAAC4AAAD6AAYAAAAxADAAMAAwADAAMAD7ABYAAAD6AQUAAABwAHAAdABfAGgA+wAAAAAACEgAAAD6AAEBBAAAAGYAYQBkAGUA+wAyAAAA+vsAFAAAAPoPCAAAABMEAAAAMQAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsNsAAAAPr7AIwAAAD6+wAyAAAA+gMBDwkAAAATAQAAADEA+wAdAAAA+vsAFgAAAAEAAAAADQAAAPoDAwAAADkAOQA5APsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAEAAAAAJwAAAPoAEAAAAHMAdAB5AGwAZQAuAHYAaQBzAGkAYgBpAGwAaQB0AHkA+wEYAAAA+gEGAAAAaABpAGQAZABlAG4A+wAAAAAA";
    PRESET_TYPES[56] = [];
    PRESET_SUBTYPES = PRESET_TYPES[56] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;944;rAMAAPr7AKUDAAD6AwEFAgYCDgAAAAAPBQAAABA4AAAAEQAAAAD7ABkAAAD6+wASAAAAAQAAAAAJAAAA+gMBAAAAMAD7AwkAAAD6AAEDECcAAPsEWAMAAPr7AFEDAAAFAAAABrgAAAD6AAECBwAAACgAcABwAHQAXwB3ACkAAwoAAAAoAC0AcABwAHQAXwB3ACoAMgApAAQA+wCBAAAA+gUDAAAAUABQAFQA+wAyAAAA+gIBDwYAAAATAwAAADUAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AdwD7BqMAAAD6AAEBDAAAACgAcABwAHQAXwB3ACoAMAAuADUAMAApAAQA+wB7AAAA+vsANwAAAPoCAQxQwwAADwcAAAATAwAAADUAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7Bq8AAAD6AAECBwAAACgAcABwAHQAXwB5ACkAAwsAAAAoADEAKwBwAHAAdABfAGgALwAyACkABAD7AHYAAAD6+wAyAAAA+g8IAAAAEwQAAAAxADAAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIhAAAA+vsAGgAAAAEAAAAAEQAAAPoABQAAAHAAcAB0AF8AeQD7CnoAAAD6AACXSQH7AG4AAAD6+wAyAAAA+g8JAAAAEwQAAAAxADAAMAAwAPsAGQAAAPr7ABIAAAABAAAAAAkAAAD6AwEAAAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wIZAAAA+vsAEgAAAAEAAAAACQAAAPoAAQAAAHIA+w2wAAAA+vsAjAAAAPr7ADIAAAD6AwEPCgAAABMBAAAAMQD7AB0AAAD6+wAWAAAAAQAAAAANAAAA+gMDAAAAOQA5ADkA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAQAAAAAnAAAA+gAQAAAAcwB0AHkAbABlAC4AdgBpAHMAaQBiAGkAbABpAHQAeQD7ARgAAAD6AQYAAABoAGkAZABkAGUAbgD7AAAAAAA=";
    ANIMATION_PRESET_CLASSES[4] = [];
    PRESET_TYPES = ANIMATION_PRESET_CLASSES[4] = [];
    PRESET_TYPES[0] = [];
    PRESET_SUBTYPES = PRESET_TYPES[0] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;3275;xwwAAPr7AMAMAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQAAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wR3DAAA+vsAcAwAAAEAAAAJZwwAAPoAAQEBAs8FAABNACAAMAAuADAAOQA4ADcAIAAwAC4AMAA3ADcANwA4ACAATAAgADAALgAwADkAOAA3ACAAMAAuADAANwA3ADcAOAAgAEMAIAAwAC4AMQAwADIANgAxACAAMAAuADAANwA4ADcAMQAgADAALgAxADAANgA2ADQAIAAwAC4AMAA3ADkAMQA3ACAAMAAuADEAMQAwADUANQAgADAALgAwADgAMAA1ADYAIABDACAAMAAuADEAMgAwADMAMgAgADAALgAwADgAMwA4ACAAMAAuADEAMgA1ACAAMAAuADAAOAA2ADMANQAgADAALgAxADMAMwA0ADcAIAAwAC4AMAA5ADAANwA0ACAATAAgADAALgAxADcAMwA1ADcAIAAwAC4AMAA4ADcAOQA3ACAAQwAgADAALgAxADkAMAAzADcAIAAwAC4AMAA4ADYANQA4ACAAMAAuADEANgA3ADUAOAAgADAALgAwADgANwAwADQAIAAwAC4AMQA4ADgAMQA1ACAAMAAuADAAOAA3ADAANAAgAEwAIAAwAC4AMgAxADYAMgA4ACAAMAAuADEAMgAyADIAMwAgAEMAIAAwAC4AMgAxADYANAAxACAAMAAuADEAMgAyADIAMwAgADAALgAyADYAMQAwADcAIAAwAC4AMQA3ADcANwA4ACAAMAAuADIANgA2ADIAOAAgADAALgAxADgAMwAzADQAIABDACAAMAAuADMAMgAzADEAOAAgADAALgAyADQAMAA5ADgAIAAwAC4AMgA2ADMAMQA1ACAAMAAuADEAOAAxADcAMgAgADAALgAzADAAMQAxADgAIAAwAC4AMgAxADUANwA0ACAAQwAgADAALgAzADAAOQA1ADEAIAAwAC4AMgAyADIAOQAyACAAMAAuADMAMQA3ADMAMgAgADAALgAyADMAMQA5ADUAIAAwAC4AMwAyADUANgA1ACAAMAAuADIAMwA4ADgAOQAgAEMAIAAwAC4AMwA0ADQAMQA0ACAAMAAuADIANQAzADcAMQAgADAALgAzADcAOAAgADAALgAyADcANgAxADYAIAAwAC4AMwA5ADcAMAAxACAAMAAuADIAOAA1ADEAOQAgAEMAIAAwAC4ANAAwADgANwAzACAAMAAuADIAOQAwADUAMQAgADAALgA0ADIAMAA3ADEAIAAwAC4AMgA5ADMAMAA2ACAAMAAuADQAMwAyADQAMwAgADAALgAyADkANwAyADMAIABDACAAMAAuADQANQAyADYAMQAgADAALgAyADkAMgA2ACAAMAAuADQANwAzADEAOAAgADAALgAyADkAMgAxADMAIAAwAC4ANAA5ADIAOAA0ACAAMAAuADIAOAAzADMANAAgAEMAIAAwAC4ANQAwADYAMQAyACAAMAAuADIANwA3ADAAOQAgADAALgA1ADMAMwA0ADcAIAAwAC4AMgA0ADAANwA0ACAAMAAuADUANAAyADMAMgAgADAALgAyADIAMwAxADUAIABDACAAMAAuADUANwAxADcANQAgADAALgAxADYANAAxADIAIAAwAC4ANQA2ADYAOAAgADAALgAxADMANQA2ADUAIAAwAC4ANQA5ADUAOQA3ACAAMAAuADAANgAyADAANAAgAEMAIAAwAC4ANQA5ADgAOQA2ACAAMAAuADAANQA0ADQAIAAwAC4ANgAwADgAMAA4ACAAMAAuADAAMgA5ADgANwAgADAALgA2ADEAMwA2ADgAIAAwAC4AMAAxADkANAA1ACAAQwAgADAALgA2ADEANgAwADIAIAAwAC4AMAAxADUAMAA1ACAAMAAuADYAMQA4ADQAOQAgADAALgAwADEAMQAxADIAIAAwAC4ANgAyADAAOQA3ACAAMAAuADAAMAA3ADQAMQAgAEMAIAAwAC4ANgAyADIAMAAxACAAMAAuADAAMAA1ADcAOQAgADAALgA2ADIAMwAwADUAIAAwAC4AMAAwADQAOAA3ACAAMAAuADYAMgA0ADAAOQAgADAALgAwADAAMwA3ADEAIABDACAAMAAuADYAMgA2ADAANQAgADAALgAwADEANwA2ACAAMAAuADYAMgA1ADIANgAgADAALgAwADEAMAA0ADIAIAAwAC4ANgAyADQAMAA5ACAAMAAuADAANAAwADcANAAgAEMAIAAwAC4ANgAyADMAOAAzACAAMAAuADAANAA3ADYAOQAgADAALgA2ADIAMwAwADUAIAAwAC4AMAA1ADQAOAA3ACAAMAAuADYAMgAyADUAMwAgADAALgAwADYAMgAwADQAIABDACAAMAAuADYAMgAxADYAMgAgADAALgAwADQANgA1ADMAIAAwAC4ANgAyADIANwA5ACAAMAAuADAANQA5ADkANgAgADAALgA2ADEAOQA0ACAAMAAuADAAMwA5ADgAMgAgAEMAIAAwAC4ANgAxADkAMAAxACAAMAAuADAAMwA3ADAANAAgADAALgA2ADEAOAAzADYAIAAwAC4AMAAzADEANAA5ACAAMAAuADYAMQA4ADMANgAgADAALgAwADMAMQA0ADkAIABMACAAMAAuADUAOQAzADMANgAgAC0AMAAuADAAMQA1ADcANAAgAEMAIAAwAC4ANQA5ADEAOAAgAC0AMAAuADAAMQA2ADQAMwAgADAALgA1ADkAMAAyADQAIAAtADAALgAwADEANwAxADMAIAAwAC4ANQA4ADgANgA4ACAALQAwAC4AMAAxADcANQA5ACAAQwAgADAALgA1ADgANwA1ACAALQAwAC4AMAAxADgAMgA4ACAAMAAuADUAOAA2ADMAMwAgAC0AMAAuADAAMQA5ADIAMQAgADAALgA1ADgANQAwADMAIAAtADAALgAwADEAOQA0ADQAIABDACAAMAAuADUAOAAzADcAMwAgAC0AMAAuADAAMgAwADEAMwAgADAALgA1ADgAMgAzACAALQAwAC4AMAAyADAAMQAzACAAMAAuADUAOAAwADgANgAgAC0AMAAuADAAMgAwADMANwAgAEMAIAAwAC4ANQA3ADcAMAA5ACAALQAwAC4AMAAyADEANQAyACAAMAAuADUANwAzADMAMQAgAC0AMAAuADAAMgAzADEANAAgADAALgA1ADYAOQA0ACAALQAwAC4AMAAyADQAMAA3ACAAQwAgADAALgA1ADYANQAxADEAIAAtADAALgAwADIANQAyADMAIAAwAC4ANQA1ADYAMwA4ACAALQAwAC4AMAAyADYAOAA1ACAAMAAuADUANQA2ADMAOAAgAC0AMAAuADAAMgA2ADgANQAgAEwAIAAwAC4ANAA1ADYAOQAgAC0AMAAuADAAMwAyADQAIABDACAAMAAuADQANQAzADEAMwAgAC0AMAAuADAAMwAzADcAOQAgADAALgA0ADQAOQAzADUAIAAtADAALgAwADMANQAxADgAIAAwAC4ANAA0ADUANAA1ACAALQAwAC4AMAAzADYAMQAxACAAQwAgADAALgA0ADMANwAxADEAIAAtADAALgAwADMAOAA0ADIAIAAwAC4ANAA0ADAANgAzACAALQAwAC4AMAAzADcAMgA2ACAAMAAuADQAMwA1ADAAMwAgAC0AMAAuADAAMwA4ADgAOAAgAEMAIAAwAC4ANAAzADQAMgA1ACAALQAwAC4AMAAzADkANQA4ACAAMAAuADQAMwAzADMANAAgAC0AMAAuADAANAAwADIANwAgADAALgA0ADMAMgA0ADMAIAAtADAALgAwADQAMAA3ADQAIABDACAAMAAuADQAMgA5ADkANQAgAC0AMAAuADAANAAyADUAOQAgADAALgA0ADMAMQA5ACAALQAwAC4AMAA0ADAAOQA3ACAAMAAuADQAMgA5ADgAMgAgAC0AMAAuADAANAAyADUAOQAgAEwAIAAwAC4ANAAyADkAOAAyACAALQAwAC4AMAA0ADIANQA5ACAAAyIAAABBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAEEAQQBBAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[1] = [];
    PRESET_SUBTYPES = PRESET_TYPES[1] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;515;/wEAAPr7APgBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQAQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wSvAQAA+vsAqAEAAAEAAAAJnwEAAPoAAQEBAo0AAABNACAAMAAgADAAIABDACAAMAAuADAANgA5ACAAMAAgADAALgAxADIANQAgADAALgAwADUANgAgADAALgAxADIANQAgADAALgAxADIANQAgAEMAIAAwAC4AMQAyADUAIAAwAC4AMQA5ADQAIAAwAC4AMAA2ADkAIAAwAC4AMgA1ACAAMAAgADAALgAyADUAIABDACAALQAwAC4AMAA2ADkAIAAwAC4AMgA1ACAALQAwAC4AMQAyADUAIAAwAC4AMQA5ADQAIAAtADAALgAxADIANQAgADAALgAxADIANQAgAEMAIAAtADAALgAxADIANQAgADAALgAwADUANgAgAC0AMAAuADAANgA5ACAAMAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[2] = [];
    PRESET_SUBTYPES = PRESET_TYPES[2] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;299;JwEAAPr7ACABAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQAgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTXAAAA+vsA0AAAAAEAAAAJxwAAAPoAAQEBAiEAAABNACAAMAAgADAAIABMACAAMAAgAC0AMAAuADEANAA3ACAATAAgADAALgAyADUAIAAwACAATAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[3] = [];
    PRESET_SUBTYPES = PRESET_TYPES[3] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;335;SwEAAPr7AEQBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQAwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wT7AAAA+vsA9AAAAAEAAAAJ6wAAAPoAAQEBAjMAAABNACAAMAAgADAAIABMACAAMAAuADEAMgA1ACAALQAwAC4AMAA4ADQAIABMACAAMAAuADIANQAgADAAIABMACAAMAAuADEAMgA1ACAAMAAuADAAOAA0ACAATAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[4] = [];
    PRESET_SUBTYPES = PRESET_TYPES[4] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;385;fQEAAPr7AHYBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQBAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQtAQAA+vsAJgEAAAEAAAAJHQEAAPoAAQEBAkwAAABNACAAMAAgADAAIABMACAAMAAuADEAMgA1ACAAMAAgAEwAIAAwAC4AMQA4ADgAIAAwAC4AMQAwADkAIABMACAAMAAuADEAMgA1ACAAMAAuADIAMQA3ACAATAAgADAAIAAwAC4AMgAxADcAIABMACAALQAwAC4AMAA2ADMAIAAwAC4AMQAwADkAIABMACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[5] = [];
    PRESET_SUBTYPES = PRESET_TYPES[5] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;511;+wEAAPr7APQBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQBQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wSrAQAA+vsApAEAAAEAAAAJmwEAAPoAAQEBAosAAABNACAAMAAgADAAIABMACAAMAAuADAAMgA5ACAAMAAuADAAOQAxACAATAAgADAALgAxADIANQAgADAALgAwADkAMQAgAEwAIAAwAC4AMAA0ADgAIAAwAC4AMQA0ADcAIABMACAAMAAuADAANwA3ACAAMAAuADIAMwA4ACAATAAgADAAIAAwAC4AMQA4ADIAIABMACAALQAwAC4AMAA3ADcAIAAwAC4AMgAzADgAIABMACAALQAwAC4AMAA0ADgAIAAwAC4AMQA0ADcAIABMACAALQAwAC4AMQAyADUAIAAwAC4AMAA5ADEAIABMACAALQAwAC4AMAAyADkAIAAwAC4AMAA5ADEAIABMACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[6] = [];
    PRESET_SUBTYPES = PRESET_TYPES[6] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;711;wwIAAPr7ALwCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQBgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRzAgAA+vsAbAIAAAEAAAAJYwIAAPoAAQEBAu8AAABNACAAMAAgADAAIABDACAALQAwAC4AMAAxADQAIAAtADAALgAwADAANQAgAC0AMAAuADAAMgA5ACAALQAwAC4AMAAwADkAIAAtADAALgAwADQANAAgAC0AMAAuADAAMAA5ACAAQwAgAC0AMAAuADEAMQA0ACAALQAwAC4AMAAwADkAIAAtADAALgAxADYAOQAgADAALgAwADQAOAAgAC0AMAAuADEANgA5ACAAMAAuADEAMQA3ACAAQwAgAC0AMAAuADEANgA5ACAAMAAuADEAOAA1ACAALQAwAC4AMQAxADQAIAAwAC4AMgA0ADEAIAAtADAALgAwADQANAAgADAALgAyADQAMQAgAEMAIAAtADAALgAwADIAOQAgADAALgAyADQAMQAgAC0AMAAuADAAMQA0ACAAMAAuADIAMwA4ACAAMAAgADAALgAyADMAMwAgAEMAIAAtADAALgAwADQANwAgADAALgAyADEANQAgAC0AMAAuADAAOAAgADAALgAxADcAIAAtADAALgAwADgAIAAwAC4AMQAxADcAIABDACAALQAwAC4AMAA4ACAAMAAuADAANgAzACAALQAwAC4AMAA0ADcAIAAwAC4AMAAxADgAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[7] = [];
    PRESET_SUBTYPES = PRESET_TYPES[7] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;319;OwEAAPr7ADQBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQBwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTrAAAA+vsA5AAAAAEAAAAJ2wAAAPoAAQEBAisAAABNACAAMAAgADAAIABMACAAMAAuADIANQAgADAAIABMACAAMAAuADIANQAgADAALgAyADUAIABMACAAMAAgADAALgAyADUAIABMACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[8] = [];
    PRESET_SUBTYPES = PRESET_TYPES[8] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;333;SQEAAPr7AEIBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQCAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wT5AAAA+vsA8gAAAAEAAAAJ6QAAAPoAAQEBAjIAAABNACAAMAAgADAAIABMACAAMAAuADEANgA3ACAAMAAgAEwAIAAwAC4AMgAxACAAMAAuADEANgA3ACAATAAgAC0AMAAuADAANAAgADAALgAxADYANwAgAEwAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[9] = [];
    PRESET_SUBTYPES = PRESET_TYPES[9] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;855;UwMAAPr7AEwDAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQCQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQDAwAA+vsA/AIAAAEAAAAJ8wIAAPoAAQEBAjcBAABNACAAMAAgADAAIABDACAAMAAuADAAMQAyACAALQAwAC4AMAAxADgAIAAwAC4AMAAzADMAIAAtADAALgAwADQANAAgADAALgAwADUAOAAgAC0AMAAuADAANAA0ACAAQwAgADAALgAwADkANQAgAC0AMAAuADAANAA0ACAAMAAuADEAMgA1ACAALQAwAC4AMAAxADcAIAAwAC4AMQAyADUAIAAwAC4AMAAxADcAIABDACAAMAAuADEAMgA1ACAAMAAuADAAMgA4ACAAMAAuADEAMgAyACAAMAAuADAAMwA4ACAAMAAuADEAMQA2ACAAMAAuADAANAA3ACAAQwAgADAALgAxADEANwAgADAALgAwADQANwAgADAAIAAwAC4AMQA4ADIAIAAwACAAMAAuADEAOAAzACAAQwAgADAAIAAwAC4AMQA4ADIAIAAtADAALgAxADEANwAgADAALgAwADQANwAgAC0AMAAuADEAMQA2ACAAMAAuADAANAA3ACAAQwAgAC0AMAAuADEAMgAyACAAMAAuADAAMwA4ACAALQAwAC4AMQAyADUAIAAwAC4AMAAyADgAIAAtADAALgAxADIANQAgADAALgAwADEANwAgAEMAIAAtADAALgAxADIANQAgAC0AMAAuADAAMQA3ACAALQAwAC4AMAA5ADUAIAAtADAALgAwADQANAAgAC0AMAAuADAANQA3ACAALQAwAC4AMAA0ADQAIABDACAALQAwAC4AMAAzADMAIAAtADAALgAwADQANAAgAC0AMAAuADAAMQAyACAALQAwAC4AMAAxADgAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[10] = [];
    PRESET_SUBTYPES = PRESET_TYPES[10] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;439;swEAAPr7AKwBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQCgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRjAQAA+vsAXAEAAAEAAAAJUwEAAPoAAQEBAmcAAABNACAAMAAgADAAIABMACAAMAAuADAANwAzACAALQAwAC4AMAA3ADMAIABMACAAMAAuADEANwA3ACAALQAwAC4AMAA3ADMAIABMACAAMAAuADIANQAgADAAIABMACAAMAAuADIANQAgADAALgAxADAANAAgAEwAIAAwAC4AMQA3ADcAIAAwAC4AMQA3ADcAIABMACAAMAAuADAANwAzACAAMAAuADEANwA3ACAATAAgADAAIAAwAC4AMQAwADQAIABMACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[11] = [];
    PRESET_SUBTYPES = PRESET_TYPES[11] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;567;MwIAAPr7ACwCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQCwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTjAQAA+vsA3AEAAAEAAAAJ0wEAAPoAAQEBAqcAAABNACAAMAAgADAAIABMACAAMAAuADAAMwA2ACAAMAAuADAANgAyACAATAAgADAALgAxADAAOAAgADAALgAwADYAMgAgAEwAIAAwAC4AMAA3ADIAIAAwAC4AMQAyADUAIABMACAAMAAuADEAMAA4ACAAMAAuADEAOAA3ACAATAAgADAALgAwADMANgAgADAALgAxADgANwAgAEwAIAAwACAAMAAuADIANQAgAEwAIAAtADAALgAwADMANgAgADAALgAxADgANwAgAEwAIAAtADAALgAxADAAOAAgADAALgAxADgANwAgAEwAIAAtADAALgAwADcAMgAgADAALgAxADIANQAgAEwAIAAtADAALgAxADAAOAAgADAALgAwADYAMgAgAEwAIAAtADAALgAwADMANgAgADAALgAwADYAMgAgAEwAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[12] = [];
    PRESET_SUBTYPES = PRESET_TYPES[12] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;527;CwIAAPr7AAQCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQDAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS7AQAA+vsAtAEAAAEAAAAJqwEAAPoAAQEBApMAAABNACAAMAAgADAAIABDACAAMAAuADAAMwAgAC0AMAAuADAAMwA4ACAAMAAuADAANwA1ACAALQAwAC4AMAA2ADIAIAAwAC4AMQAyADUAIAAtADAALgAwADYAMgAgAEMAIAAwAC4AMQA3ADUAIAAtADAALgAwADYAMgAgADAALgAyADIAIAAtADAALgAwADMAOAAgADAALgAyADUAIAAwACAAQwAgADAALgAyADIAIAAwAC4AMAAzADgAIAAwAC4AMQA3ADUAIAAwAC4AMAA2ADIAIAAwAC4AMQAyADUAIAAwAC4AMAA2ADIAIABDACAAMAAuADAANwA1ACAAMAAuADAANgAyACAAMAAuADAAMwAgADAALgAwADMAOAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[13] = [];
    PRESET_SUBTYPES = PRESET_TYPES[13] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;317;OQEAAPr7ADIBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQDQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTpAAAA+vsA4gAAAAEAAAAJ2QAAAPoAAQEBAioAAABNACAAMAAgADAAIABMACAAMAAuADEAMgA1ACAAMAAuADIAMQA2ACAATAAgAC0AMAAuADEAMgA1ACAAMAAuADIAMQA2ACAATAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[14] = [];
    PRESET_SUBTYPES = PRESET_TYPES[14] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;333;SQEAAPr7AEIBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQDgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wT5AAAA+vsA8gAAAAEAAAAJ6QAAAPoAAQEBAjIAAABNACAAMAAgADAAIABMACAAMAAuADEANwA4ACAAMAAgAEwAIAAwAC4AMgA1ACAAMAAuADEAMgAxACAATAAgADAALgAwADcAMgAgADAALgAxADIAMQAgAEwAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[15] = [];
    PRESET_SUBTYPES = PRESET_TYPES[15] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;375;cwEAAPr7AGwBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQDwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQjAQAA+vsAHAEAAAEAAAAJEwEAAPoAAQEBAkcAAABNACAAMAAgADAAIABMACAAMAAuADEAMgA1ACAAMAAuADAAOQAxACAATAAgADAALgAwADcANwAgADAALgAyADMAOAAgAEwAIAAtADAALgAwADcANwAgADAALgAyADMAOAAgAEwAIAAtADAALgAxADIANQAgADAALgAwADkAMQAgAEwAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[16] = [];
    PRESET_SUBTYPES = PRESET_TYPES[16] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;453;wQEAAPr7ALoBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQEAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRxAQAA+vsAagEAAAEAAAAJYQEAAPoAAQEBAm4AAABNACAAMAAgADAAIABMACAAMAAuADAAOQAxACAALQAwAC4AMAAzADQAIABMACAAMAAuADEAMgA1ACAALQAwAC4AMQAyADUAIABMACAAMAAuADEANQA4ACAALQAwAC4AMAAzADQAIABMACAAMAAuADIANAA5ACAAMAAgAEwAIAAwAC4AMQA1ADgAIAAwAC4AMAAzADQAIABMACAAMAAuADEAMgA1ACAAMAAuADEAMgA1ACAATAAgADAALgAwADkAMQAgADAALgAwADMANAAgAEwAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[17] = [];
    PRESET_SUBTYPES = PRESET_TYPES[17] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;635;dwIAAPr7AHACAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQEQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQnAgAA+vsAIAIAAAEAAAAJFwIAAPoAAQEBAskAAABNACAAMAAgADAAIABMACAAMAAuADAANQAyACAAMAAgAEwAIAAwAC4AMAA4ADkAIAAtADAALgAwADMANwAgAEwAIAAwAC4AMQAyADUAIAAwACAATAAgADAALgAxADcANwAgADAAIABMACAAMAAuADEANwA3ACAAMAAuADAANQAyACAATAAgADAALgAyADEAMwAgADAALgAwADgAOQAgAEwAIAAwAC4AMQA3ADcAIAAwAC4AMQAyADUAIABMACAAMAAuADEANwA3ACAAMAAuADEANwA3ACAATAAgADAALgAxADIANQAgADAALgAxADcANwAgAEwAIAAwAC4AMAA4ADkAIAAwAC4AMgAxADMAIABMACAAMAAuADAANQAyACAAMAAuADEANwA3ACAATAAgADAAIAAwAC4AMQA3ADcAIABMACAAMAAgADAALgAxADIANQAgAEwAIAAtADAALgAwADMANwAgADAALgAwADgAOQAgAEwAIAAwACAAMAAuADAANQAyACAATAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[18] = [];
    PRESET_SUBTYPES = PRESET_TYPES[18] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;843;RwMAAPr7AEADAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQEgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wT3AgAA+vsA8AIAAAEAAAAJ5wIAAPoAAQEBAjEBAABNACAAMAAgADAAIABDACAAMAAuADAAMAAxACAAMAAuADAAMwA0ACAAMAAuADAAMQAxACAAMAAuADAANgA1ACAAMAAuADAAMgA4ACAAMAAuADAAOAA1ACAAQwAgADAALgAwADIAOAAgADAALgAwADgANgAgADAALgAwADUANQAgADAALgAxADEAMwAgADAALgAwADUANQAgADAALgAxADEAMgAgAEMAIAAwAC4AMAA3ACAAMAAuADEAMgA3ACAAMAAuADAANwA5ACAAMAAuADEANAA4ACAAMAAuADAANwA5ACAAMAAuADEANwAgAEMAIAAwAC4AMAA3ADkAIAAwAC4AMgAxADQAIAAwAC4AMAA0ADQAIAAwAC4AMgA0ADkAIAAwACAAMAAuADIANQAgAEMAIAAtADAALgAwADQANAAgADAALgAyADQAOQAgAC0AMAAuADAANwA5ACAAMAAuADIAMQA0ACAALQAwAC4AMAA3ADkAIAAwAC4AMQA3ACAAQwAgAC0AMAAuADAANwA5ACAAMAAuADEANAA4ACAALQAwAC4AMAA3ACAAMAAuADEAMgA3ACAALQAwAC4AMAA1ADUAIAAwAC4AMQAxADIAIABDACAALQAwAC4AMAA1ADUAIAAwAC4AMQAxADMAIAAtADAALgAwADIAOAAgADAALgAwADgANgAgAC0AMAAuADAAMgA4ACAAMAAuADAAOAA1ACAAQwAgAC0AMAAuADAAMQAxACAAMAAuADAANgA1ACAALQAwAC4AMAAwADEAIAAwAC4AMAAzADQAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[19] = [];
    PRESET_SUBTYPES = PRESET_TYPES[19] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;531;DwIAAPr7AAgCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQEwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS/AQAA+vsAuAEAAAEAAAAJrwEAAPoAAQEBApUAAABNACAAMAAgADAAIABDACAAMAAuADAANgA5ACAAMAAgADAALgAxADIANAAgAC0AMAAuADAANQA2ACAAMAAuADEAMgA0ACAALQAwAC4AMQAyADUAIABDACAAMAAuADEAMgA0ACAALQAwAC4AMAA1ADYAIAAwAC4AMQA3ADkAIAAtADAALgAwADAAMQAgADAALgAyADQAOAAgAC0AMAAuADAAMAAxACAAQwAgADAALgAxADcAOQAgAC0AMAAuADAAMAAxACAAMAAuADEAMgA1ACAAMAAuADAANQA2ACAAMAAuADEAMgA1ACAAMAAuADEAMgA1ACAAQwAgADAALgAxADIANQAgADAALgAwADUANgAgADAALgAwADYAOQAgADAAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[20] = [];
    PRESET_SUBTYPES = PRESET_TYPES[20] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;607;WwIAAPr7AFQCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQFAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQLAgAA+vsABAIAAAEAAAAJ+wEAAPoAAQEBArsAAABNACAAMAAgADAAIABDACAAMAAgAC0AMAAuADAAMwAyACAAMAAuADAAMgA2ACAALQAwAC4AMAA1ADgAIAAwAC4AMAA1ADgAIAAtADAALgAwADUAOAAgAEwAIAAwAC4AMQA5ADIAIAAtADAALgAwADUAOAAgAEMAIAAwAC4AMgAyADQAIAAtADAALgAwADUAOAAgADAALgAyADUAIAAtADAALgAwADMAMgAgADAALgAyADUAIAAwACAATAAgADAALgAyADUAIAAwAC4AMQAzADIAIABDACAAMAAuADIANQAgADAALgAxADYANAAgADAALgAyADIANAAgADAALgAxADkAMQAgADAALgAxADkAMgAgADAALgAxADkAMQAgAEwAIAAwAC4AMAA1ADgAIAAwAC4AMQA5ADEAIABDACAAMAAuADAAMgA2ACAAMAAuADEAOQAxACAAMAAgADAALgAxADYANAAgADAAIAAwAC4AMQAzADIAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[21] = [];
    PRESET_SUBTYPES = PRESET_TYPES[21] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1429;kQUAAPr7AIoFAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQFQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRBBQAA+vsAOgUAAAEAAAAJMQUAAPoAAQEBAlYCAABNACAAMAAgADAAIABDACAAMAAuADAAMAA2ACAAMAAuADAAMAA2ACAAMAAuADAAMQAxACAAMAAuADAAMQAxACAAMAAuADAAMQA1ACAAMAAuADAAMQA3ACAAQwAgADAALgAwADIAIAAwAC4AMAAxADEAIAAwAC4AMAAyADQAIAAwAC4AMAAwADYAIAAwAC4AMAAzACAAMAAgAEMAIAAwAC4AMAA2ADUAIAAtADAALgAwADMANQAgADAALgAxADAANwAgAC0AMAAuADAANQAgADAALgAxADIANAAgAC0AMAAuADAAMwA0ACAAQwAgADAALgAxADQAIAAtADAALgAwADEANwAgADAALgAxADIANQAgADAALgAwADIANQAgADAALgAwADkAIAAwAC4AMAA2ACAAQwAgADAALgAwADgANAAgADAALgAwADYANQAgADAALgAwADcAOQAgADAALgAwADcAIAAwAC4AMAA3ADMAIAAwAC4AMAA3ADUAIABDACAAMAAuADAANwA5ACAAMAAuADAANwA5ACAAMAAuADAAOAA0ACAAMAAuADAAOAA0ACAAMAAuADAAOQAgADAALgAwADkAIABDACAAMAAuADEAMgA1ACAAMAAuADEAMgA1ACAAMAAuADEANAAgADAALgAxADYANwAgADAALgAxADIANAAgADAALgAxADgAMwAgAEMAIAAwAC4AMQAwADcAIAAwAC4AMgAgADAALgAwADYANQAgADAALgAxADgANQAgADAALgAwADMAIAAwAC4AMQA1ACAAQwAgADAALgAwADIANAAgADAALgAxADQANAAgADAALgAwADIAIAAwAC4AMQAzADkAIAAwAC4AMAAxADUAIAAwAC4AMQAzADMAIABDACAAMAAuADAAMQAxACAAMAAuADEAMwA5ACAAMAAuADAAMAA2ACAAMAAuADEANAA0ACAAMAAgADAALgAxADUAIABDACAALQAwAC4AMAAzADUAIAAwAC4AMQA4ADUAIAAtADAALgAwADcANwAgADAALgAyACAALQAwAC4AMAA5ADQAIAAwAC4AMQA4ADMAIABDACAALQAwAC4AMQAxACAAMAAuADEANgA3ACAALQAwAC4AMAA5ADUAIAAwAC4AMQAyADUAIAAtADAALgAwADYAIAAwAC4AMAA5ACAAQwAgAC0AMAAuADAANQA0ACAAMAAuADAAOAA0ACAALQAwAC4AMAA0ADkAIAAwAC4AMAA3ADkAIAAtADAALgAwADQAMwAgADAALgAwADcANQAgAEMAIAAtADAALgAwADQAOQAgADAALgAwADcAIAAtADAALgAwADUANAAgADAALgAwADYANQAgAC0AMAAuADAANgAgADAALgAwADYAIABDACAALQAwAC4AMAA5ADUAIAAwAC4AMAAyADUAIAAtADAALgAxADEAIAAtADAALgAwADEANwAgAC0AMAAuADAAOQA0ACAALQAwAC4AMAAzADQAIABDACAALQAwAC4AMAA3ADcAIAAtADAALgAwADUAIAAtADAALgAwADMANQAgAC0AMAAuADAAMwA1ACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[22] = [];
    PRESET_SUBTYPES = PRESET_TYPES[22] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;849;TQMAAPr7AEYDAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQFgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wT9AgAA+vsA9gIAAAEAAAAJ7QIAAPoAAQEBAjQBAABNACAAMAAgADAAIABDACAAMAAuADAAMwAzACAAMAAgADAALgAwADYAIAAwAC4AMAAyADcAIAAwAC4AMAA2ACAAMAAuADAANgAgAEMAIAAwAC4AMAA2ACAAMAAuADAAOQA5ACAAMAAuADAAMwAgADAALgAxADEAMwAgADAALgAwADEAMgAgADAALgAxADEAOQAgAEwAIAAtADAALgAwADEAMgAgADAALgAxADIANQAgAEMAIAAtADAALgAwADMAIAAwAC4AMQAzADEAIAAtADAALgAwADYAIAAwAC4AMQA0ADYAIAAtADAALgAwADYAIAAwAC4AMQA5ACAAQwAgAC0AMAAuADAANgAgADAALgAyADEAOAAgAC0AMAAuADAAMwAzACAAMAAuADIANQAgADAAIAAwAC4AMgA1ACAAQwAgADAALgAwADMAMwAgADAALgAyADUAIAAwAC4AMAA2ACAAMAAuADIAMQA4ACAAMAAuADAANgAgADAALgAxADkAIABDACAAMAAuADAANgAgADAALgAxADQANgAgADAALgAwADMAIAAwAC4AMQAzADEAIAAwAC4AMAAxADIAIAAwAC4AMQAyADUAIABMACAALQAwAC4AMAAxADIAIAAwAC4AMQAxADkAIABDACAALQAwAC4AMAAzACAAMAAuADEAMQAzACAALQAwAC4AMAA2ACAAMAAuADAAOQA5ACAALQAwAC4AMAA2ACAAMAAuADAANgAgAEMAIAAtADAALgAwADYAIAAwAC4AMAAyADcAIAAtADAALgAwADMAMwAgADAAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[23] = [];
    PRESET_SUBTYPES = PRESET_TYPES[23] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;619;ZwIAAPr7AGACAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQFwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQXAgAA+vsAEAIAAAEAAAAJBwIAAPoAAQEBAsEAAABNACAAMAAgADAAIABDACAAMAAuADAANwAyACAAMAAuADAANQA4ACAAMAAuADEAIAAwAC4AMQA1ADIAIAAwAC4AMAA3ADcAIAAwAC4AMgAzADgAIABDACAALQAwAC4AMAAxADUAIAAwAC4AMgAzADMAIAAtADAALgAwADkAMwAgADAALgAxADcAMwAgAC0AMAAuADEAMgA1ACAAMAAuADAAOQAxACAAQwAgAC0AMAAuADAANAA3ACAAMAAuADAANAAgADAALgAwADUAMQAgADAALgAwADQAMwAgADAALgAxADIANQAgADAALgAwADkAMQAgAEMAIAAwAC4AMAA5ADIAIAAwAC4AMQA3ADgAIAAwAC4AMAAxADEAIAAwAC4AMgAzADMAIAAtADAALgAwADcANwAgADAALgAyADMAOAAgAEMAIAAtADAALgAxADAAMQAgADAALgAxADQAOAAgAC0AMAAuADAANgA4ACAAMAAuADAANQA2ACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[24] = [];
    PRESET_SUBTYPES = PRESET_TYPES[24] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1273;9QQAAPr7AO4EAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQGAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wSlBAAA+vsAngQAAAEAAAAJlQQAAPoAAQEBAggCAABNACAAMAAgADAAIABDACAAMAAuADAAMgAzACAAMAAuADAAMAAxACAAMAAuADAANAAyACAAMAAuADAAMAA5ACAAMAAuADAANQAyACAAMAAuADAAMgAxACAATAAgADAALgAwADcANQAgADAALgAwADQAOQAgAEMAIAAwAC4AMAA4ACAAMAAuADAANQA1ACAAMAAuADAAOAA4ACAAMAAuADAANQA4ACAAMAAuADAAOQA4ACAAMAAuADAANQA4ACAAQwAgADAALgAxADEAMgAgADAALgAwADUAOAAgADAALgAxADIANAAgADAALgAwADUAIAAwAC4AMQAyADUAIAAwAC4AMAAzADgAIABDACAAMAAuADEAMgA0ACAAMAAuADAAMgA4ACAAMAAuADEAMQAyACAAMAAuADAAMQA5ACAAMAAuADAAOQA4ACAAMAAuADAAMQA5ACAAQwAgADAALgAwADgAOAAgADAALgAwADEAOQAgADAALgAwADgAIAAwAC4AMAAyADMAIAAwAC4AMAA3ADUAIAAwAC4AMAAyADgAIABMACAAMAAuADAANQAyACAAMAAuADAANQA2ACAAQwAgADAALgAwADQAMgAgADAALgAwADYAOAAgADAALgAwADIAMwAgADAALgAwADcANgAgADAAIAAwAC4AMAA3ADcAIABDACAALQAwAC4AMAAyADMAIAAwAC4AMAA3ADYAIAAtADAALgAwADQAMgAgADAALgAwADYAOAAgAC0AMAAuADAANQAyACAAMAAuADAANQA2ACAATAAgAC0AMAAuADAANwA1ACAAMAAuADAAMgA4ACAAQwAgAC0AMAAuADAAOAAgADAALgAwADIAMwAgAC0AMAAuADAAOAA4ACAAMAAuADAAMQA5ACAALQAwAC4AMAA5ADgAIAAwAC4AMAAxADkAIABDACAALQAwAC4AMQAxADIAIAAwAC4AMAAxADkAIAAtADAALgAxADIANAAgADAALgAwADIAOAAgAC0AMAAuADEAMgA1ACAAMAAuADAAMwA4ACAAQwAgAC0AMAAuADEAMgA0ACAAMAAuADAANQAgAC0AMAAuADEAMQAyACAAMAAuADAANQA4ACAALQAwAC4AMAA5ADgAIAAwAC4AMAA1ADgAIABDACAALQAwAC4AMAA4ADgAIAAwAC4AMAA1ADgAIAAtADAALgAwADgAIAAwAC4AMAA1ADUAIAAtADAALgAwADcANQAgADAALgAwADQAOQAgAEwAIAAtADAALgAwADUAMgAgADAALgAwADIAMQAgAEMAIAAtADAALgAwADQAMgAgADAALgAwADAAOQAgAC0AMAAuADAAMgAzACAAMAAuADAAMAAxACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[26] = [];
    PRESET_SUBTYPES = PRESET_TYPES[26] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;849;TQMAAPr7AEYDAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQGgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wT9AgAA+vsA9gIAAAEAAAAJ7QIAAPoAAQEBAjQBAABNACAAMAAgADAAIABDACAAMAAgADAALgAwADMAMwAgADAALgAwADIANwAgADAALgAwADYAIAAwAC4AMAA2ACAAMAAuADAANgAgAEMAIAAwAC4AMAA5ADkAIAAwAC4AMAA2ACAAMAAuADEAMQAzACAAMAAuADAAMwAgADAALgAxADEAOQAgADAALgAwADEAMgAgAEwAIAAwAC4AMQAyADUAIAAtADAALgAwADEAMgAgAEMAIAAwAC4AMQAzADEAIAAtADAALgAwADMAIAAwAC4AMQA0ADYAIAAtADAALgAwADYAIAAwAC4AMQA5ACAALQAwAC4AMAA2ACAAQwAgADAALgAyADEAOAAgAC0AMAAuADAANgAgADAALgAyADUAIAAtADAALgAwADMAMwAgADAALgAyADUAIAAwACAAQwAgADAALgAyADUAIAAwAC4AMAAzADMAIAAwAC4AMgAxADgAIAAwAC4AMAA2ACAAMAAuADEAOQAgADAALgAwADYAIABDACAAMAAuADEANAA2ACAAMAAuADAANgAgADAALgAxADMAMQAgADAALgAwADMAIAAwAC4AMQAyADUAIAAwAC4AMAAxADIAIABMACAAMAAuADEAMQA5ACAALQAwAC4AMAAxADIAIABDACAAMAAuADEAMQAzACAALQAwAC4AMAAzACAAMAAuADAAOQA5ACAALQAwAC4AMAA2ACAAMAAuADAANgAgAC0AMAAuADAANgAgAEMAIAAwAC4AMAAyADcAIAAtADAALgAwADYAIAAwACAALQAwAC4AMAAzADMAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[27] = [];
    PRESET_SUBTYPES = PRESET_TYPES[27] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1325;KQUAAPr7ACIFAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQGwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTZBAAA+vsA0gQAAAEAAAAJyQQAAPoAAQEBAiICAABNACAAMAAgADAAIABDACAAMAAuADAAMwA4ACAAMAAgADAALgAwADYAOQAgADAALgAwADMAMQAgADAALgAwADYAOQAgADAALgAwADYAOQAgAEMAIAAwAC4AMAA2ADkAIAAwAC4AMAA5ADQAIAAwAC4AMAA1ADYAIAAwAC4AMQAxADYAIAAwAC4AMAAzADcAIAAwAC4AMQAyADkAIABDACAAMAAuADAAMwA3ACAAMAAuADEAMgA5ACAAMAAuADAAMwA2ACAAMAAuADEAMgA5ACAAMAAuADAAMwA2ACAAMAAuADEAMgA5ACAAQwAgADAALgAwADIAOQAgADAALgAxADMANAAgADAALgAwADIANQAgADAALgAxADQAMgAgADAALgAwADIANQAgADAALgAxADUAMQAgAEMAIAAwAC4AMAAyADUAIAAwAC4AMQA1ADkAIAAwAC4AMAAyADkAIAAwAC4AMQA2ADYAIAAwAC4AMAAzADQAIAAwAC4AMQA3ADEAIABDACAAMAAuADAANAAyACAAMAAuADEANwA5ACAAMAAuADAANAA3ACAAMAAuADEAOQAxACAAMAAuADAANAA3ACAAMAAuADIAMAAzACAAQwAgADAALgAwADQANwAgADAALgAyADIAOQAgADAALgAwADIANgAgADAALgAyADUAIAAwACAAMAAuADIANQAgAEMAIAAtADAALgAwADIANgAgADAALgAyADUAIAAtADAALgAwADQANwAgADAALgAyADIAOQAgAC0AMAAuADAANAA3ACAAMAAuADIAMAAzACAAQwAgAC0AMAAuADAANAA3ACAAMAAuADEAOQAxACAALQAwAC4AMAA0ADIAIAAwAC4AMQA3ADkAIAAtADAALgAwADMANAAgADAALgAxADcAMQAgAEMAIAAtADAALgAwADIAOQAgADAALgAxADYANgAgAC0AMAAuADAAMgA2ACAAMAAuADEANQA5ACAALQAwAC4AMAAyADYAIAAwAC4AMQA1ADEAIABDACAALQAwAC4AMAAyADYAIAAwAC4AMQA0ADIAIAAtADAALgAwADMAIAAwAC4AMQAzADQAIAAtADAALgAwADMANgAgADAALgAxADIAOQAgAEMAIAAtADAALgAwADMANgAgADAALgAxADIAOQAgAC0AMAAuADAAMwA3ACAAMAAuADEAMgA5ACAALQAwAC4AMAAzADcAIAAwAC4AMQAyADkAIABDACAALQAwAC4AMAA1ADcAIAAwAC4AMQAxADYAIAAtADAALgAwADcAIAAwAC4AMAA5ADQAIAAtADAALgAwADcAIAAwAC4AMAA2ADkAIABDACAALQAwAC4AMAA3ACAAMAAuADAAMwAxACAALQAwAC4AMAAzADkAIAAwACAAMAAgADAAIABDACAAMAAgADAAIAAwACAAMAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[28] = [];
    PRESET_SUBTYPES = PRESET_TYPES[28] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1409;fQUAAPr7AHYFAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQHAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQtBQAA+vsAJgUAAAEAAAAJHQUAAPoAAQEBAkwCAABNACAAMAAgADAAIABDACAAMAAuADAAMQA3ACAAMAAgADAALgAwADMAMQAgADAALgAwADEANAAgADAALgAwADMAMQAgADAALgAwADMAMQAgAEMAIAAwAC4AMAAzADEAIAAwAC4AMAA0ADkAIAAwAC4AMAAxADcAIAAwAC4AMAA2ADMAIAAwACAAMAAuADAANgAzACAAQwAgAC0AMAAuADAAMQA3ACAAMAAuADAANgAzACAALQAwAC4AMAAzADEAIAAwAC4AMAA3ADcAIAAtADAALgAwADMAMQAgADAALgAwADkANAAgAEMAIAAtADAALgAwADMAMQAgADAALgAxADEAMQAgAC0AMAAuADAAMQA3ACAAMAAuADEAMgA1ACAAMAAgADAALgAxADIANQAgAEMAIAAwAC4AMAAxADcAIAAwAC4AMQAyADUAIAAwAC4AMAAzADEAIAAwAC4AMQAzADkAIAAwAC4AMAAzADEAIAAwAC4AMQA1ADYAIABDACAAMAAuADAAMwAxACAAMAAuADEANwAzACAAMAAuADAAMQA3ACAAMAAuADEAOAA3ACAAMAAgADAALgAxADgANwAgAEMAIAAtADAALgAwADEANwAgADAALgAxADgANwAgAC0AMAAuADAAMwAxACAAMAAuADIAMAAxACAALQAwAC4AMAAzADEAIAAwAC4AMgAxADkAIABDACAALQAwAC4AMAAzADEAIAAwAC4AMgAzADYAIAAtADAALgAwADEANwAgADAALgAyADUAIAAwACAAMAAuADIANQAgAEMAIAAwAC4AMAAxADcAIAAwAC4AMgA1ACAAMAAuADAAMwAxACAAMAAuADIAMwA2ACAAMAAuADAAMwAxACAAMAAuADIAMQA5ACAAQwAgADAALgAwADMAMQAgADAALgAyADAAMQAgADAALgAwADEANwAgADAALgAxADgANwAgADAAIAAwAC4AMQA4ADcAIABDACAALQAwAC4AMAAxADcAIAAwAC4AMQA4ADcAIAAtADAALgAwADMAMQAgADAALgAxADcAMwAgAC0AMAAuADAAMwAxACAAMAAuADEANQA2ACAAQwAgAC0AMAAuADAAMwAxACAAMAAuADEAMwA5ACAALQAwAC4AMAAxADcAIAAwAC4AMQAyADUAIAAwACAAMAAuADEAMgA1ACAAQwAgADAALgAwADEANwAgADAALgAxADIANQAgADAALgAwADMAMQAgADAALgAxADEAMQAgADAALgAwADMAMQAgADAALgAwADkANAAgAEMAIAAwAC4AMAAzADEAIAAwAC4AMAA3ADcAIAAwAC4AMAAxADcAIAAwAC4AMAA2ADMAIAAwACAAMAAuADAANgAzACAAQwAgAC0AMAAuADAAMQA3ACAAMAAuADAANgAzACAALQAwAC4AMAAzADEAIAAwAC4AMAA0ADkAIAAtADAALgAwADMAMQAgADAALgAwADMAMQAgAEMAIAAtADAALgAwADMAMQAgADAALgAwADEANAAgAC0AMAAuADAAMQA3ACAAMAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[29] = [];
    PRESET_SUBTYPES = PRESET_TYPES[29] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;2505;xQkAAPr7AL4JAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQHQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wR1CQAA+vsAbgkAAAEAAAAJZQkAAPoAAQEBAnAEAABNACAAMAAgADAAIABDACAAMAAuADAAMAA3ACAALQAwAC4AMAAxACAAMAAuADAAMQA0ACAALQAwAC4AMAAyADEAIAAwAC4AMAAyADEAIAAtADAALgAwADMANQAgAEMAIAAwAC4AMAA0ACAALQAwAC4AMAA3ADUAIAAwAC4AMAA0ADUAIAAtADAALgAxADEANAAgADAALgAwADMAMQAgAC0AMAAuADEAMgAgAEMAIAAwAC4AMAAxADcAIAAtADAALgAxADIANwAgAC0AMAAuADAAMQAgAC0AMAAuADAAOQA5ACAALQAwAC4AMAAyADkAIAAtADAALgAwADUAOQAgAEMAIAAtADAALgAwADMAOQAgAC0AMAAuADAAMwA4ACAALQAwAC4AMAA0ADUAIAAtADAALgAwADEAOAAgAC0AMAAuADAANAA3ACAALQAwAC4AMAAwADMAIABDACAALQAwAC4AMAA1ACAAMAAuADAAMAA5ACAALQAwAC4AMAA1ADEAIAAwAC4AMAAyADEAIAAtADAALgAwADUAMQAgADAALgAwADMANQAgAEMAIAAtADAALgAwADUAMQAgADAALgAwADgAIAAtADAALgAwADMAOAAgADAALgAxADEANwAgAC0AMAAuADAAMgAzACAAMAAuADEAMQA3ACAAQwAgAC0AMAAuADAAMAA4ACAAMAAuADEAMQA3ACAAMAAuADAAMAA1ACAAMAAuADAAOAAgADAALgAwADAANQAgADAALgAwADMANQAgAEMAIAAwAC4AMAAwADUAIAAwAC4AMAAxADQAIAAwAC4AMAAwADIAIAAtADAALgAwADAANgAgAC0AMAAuADAAMAAzACAALQAwAC4AMAAyACAAQwAgAC0AMAAuADAAMAA1ACAALQAwAC4AMAAzADIAIAAtADAALgAwADEAIAAtADAALgAwADQANQAgAC0AMAAuADAAMQA2ACAALQAwAC4AMAA1ADgAIABDACAALQAwAC4AMAAzADYAIAAtADAALgAwADkAOQAgAC0AMAAuADAANgAzACAALQAwAC4AMQAyADcAIAAtADAALgAwADcANwAgAC0AMAAuADEAMgAgAEMAIAAtADAALgAwADkAMQAgAC0AMAAuADEAMQAzACAALQAwAC4AMAA4ADYAIAAtADAALgAwADcANQAgAC0AMAAuADAANgA2ACAALQAwAC4AMAAzADQAIABDACAALQAwAC4AMAA1ADgAIAAtADAALgAwADEANQAgAC0AMAAuADAANAA3ACAAMAAuADAAMAAxACAALQAwAC4AMAAzADYAIAAwAC4AMAAxADIAIABDACAALQAwAC4AMAAyADgAIAAwAC4AMAAyADIAIAAtADAALgAwADEAOQAgADAALgAwADMAMQAgAC0AMAAuADAAMAA3ACAAMAAuADAANAAgAEMAIAAwAC4AMAAyADkAIAAwAC4AMAA2ADkAIAAwAC4AMAA2ADUAIAAwAC4AMAA4ADIAIAAwAC4AMAA3ADUAIAAwAC4AMAA3ACAAQwAgADAALgAwADgANAAgADAALgAwADUAOAAgADAALgAwADYANAAgADAALgAwADIANQAgADAALgAwADIAOAAgAC0AMAAuADAAMAAzACAAQwAgADAALgAwADEAMwAgAC0AMAAuADAAMQA1ACAALQAwAC4AMAAwADMAIAAtADAALgAwADIANAAgAC0AMAAuADAAMQA2ACAALQAwAC4AMAAzACAAQwAgAC0AMAAuADAAMgA4ACAALQAwAC4AMAAzADYAIAAtADAALgAwADQAMwAgAC0AMAAuADAANAAxACAALQAwAC4AMAA1ADkAIAAtADAALgAwADQANAAgAEMAIAAtADAALgAxADAAMwAgAC0AMAAuADAANQA0ACAALQAwAC4AMQA0ADEAIAAtADAALgAwADUAMQAgAC0AMAAuADEANAA0ACAALQAwAC4AMAAzADUAIABDACAALQAwAC4AMQA0ADgAIAAtADAALgAwADIAIAAtADAALgAxADEANQAgADAAIAAtADAALgAwADcAMQAgADAALgAwADEAIABDACAALQAwAC4AMAA1ADEAIAAwAC4AMAAxADQAIAAtADAALgAwADMAMgAgADAALgAwADEANgAgAC0AMAAuADAAMQA3ACAAMAAuADAAMQA1ACAAQwAgAC0AMAAuADAAMAA0ACAAMAAuADAAMQA1ACAAMAAuADAAMQAgADAALgAwADEAMwAgADAALgAwADIANQAgADAALgAwADEAIABDACAAMAAuADAANgA5ACAAMAAgADAALgAxADAAMgAgAC0AMAAuADAAMgAxACAAMAAuADAAOQA4ACAALQAwAC4AMAAzADYAIABDACAAMAAuADAAOQA1ACAALQAwAC4AMAA1ADEAIAAwAC4AMAA1ADcAIAAtADAALgAwADUANQAgADAALgAwADEAMwAgAC0AMAAuADAANAA1ACAAQwAgAC0AMAAuADAAMAA4ACAALQAwAC4AMAA0ACAALQAwAC4AMAAyADcAIAAtADAALgAwADMAMwAgAC0AMAAuADAANAAgAC0AMAAuADAAMgA1ACAAQwAgAC0AMAAuADAANQAxACAALQAwAC4AMAAxADkAIAAtADAALgAwADYAMgAgAC0AMAAuADAAMQAyACAALQAwAC4AMAA3ADQAIAAtADAALgAwADAAMwAgAEMAIAAtADAALgAxADAAOQAgADAALgAwADIANgAgAC0AMAAuADEAMwAgADAALgAwADUAOAAgAC0AMAAuADEAMgAgADAALgAwADcAIABDACAALQAwAC4AMQAxADEAIAAwAC4AMAA4ADIAIAAtADAALgAwADcANAAgADAALgAwADYAOQAgAC0AMAAuADAAMwA5ACAAMAAuADAANAAxACAAQwAgAC0AMAAuADAAMgAyACAAMAAuADAAMgA3ACAALQAwAC4AMAAwADgAIAAwAC4AMAAxADMAIAAwACAAMAAgAFoAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[30] = [];
    PRESET_SUBTYPES = PRESET_TYPES[30] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1153;fQQAAPr7AHYEAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQHgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQtBAAA+vsAJgQAAAEAAAAJHQQAAPoAAQEBAswBAABNACAAMAAgADAAIABDACAAMAAgADAAIAAwAC4AMAAxADcAIAAtADAALgAwADYANQAgADAALgAwADEANwAgAC0AMAAuADAANgA1ACAAQwAgADAALgAwADMANAAgAC0AMAAuADEAMQA4ACAAMAAuADAANgAxACAALQAwAC4AMQAzADkAIAAwAC4AMQAgAC0AMAAuADEAMwA5ACAAQwAgADAALgAxADIAIAAtADAALgAxADMAOQAgADAALgAxADMAOAAgAC0AMAAuADEAMwAxACAAMAAuADEANQAyACAALQAwAC4AMQAxADgAIABDACAAMAAuADEANgAyACAALQAwAC4AMQAwADkAIAAwAC4AMQA3ADQAIAAtADAALgAxADAANAAgADAALgAxADgANwAgAC0AMAAuADEAMAA0ACAAQwAgADAALgAyADEAMgAgAC0AMAAuADEAMAA0ACAAMAAuADIAMwAzACAALQAwAC4AMQAyADIAIAAwAC4AMgA0ADEAIAAtADAALgAxADQAOAAgAEMAIAAwAC4AMgA0ADEAIAAtADAALgAxADQAOAAgADAALgAyADUAIAAtADAALgAxADcAOQAgADAALgAyADUAIAAtADAALgAxADcAOQAgAEMAIAAwAC4AMgA1ACAALQAwAC4AMQA3ADkAIAAwAC4AMgAzADIAIAAtADAALgAxADEAMwAgADAALgAyADMAMgAgAC0AMAAuADEAMQAzACAAQwAgADAALgAyADEANQAgAC0AMAAuADAANgAxACAAMAAuADEAOAA4ACAALQAwAC4AMAA0ACAAMAAuADEANQAgAC0AMAAuADAANAAgAEMAIAAwAC4AMQAzACAALQAwAC4AMAA0ACAAMAAuADEAMQAxACAALQAwAC4AMAA0ADgAIAAwAC4AMAA5ADYAIAAtADAALgAwADYAMgAgAEMAIAAwAC4AMAA4ADcAIAAtADAALgAwADcAIAAwAC4AMAA3ADUAIAAtADAALgAwADcANQAgADAALgAwADYAMwAgAC0AMAAuADAANwA1ACAAQwAgADAALgAwADMAOAAgAC0AMAAuADAANwA1ACAAMAAuADAAMQA3ACAALQAwAC4AMAA1ADcAIAAwAC4AMAAwADkAIAAtADAALgAwADMAMQAgAEMAIAAwAC4AMAAwADkAIAAtADAALgAwADMAMQAgADAAIAAwACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[31] = [];
    PRESET_SUBTYPES = PRESET_TYPES[31] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;709;wQIAAPr7ALoCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQHwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRxAgAA+vsAagIAAAEAAAAJYQIAAPoAAQEBAu4AAABNACAAMAAgADAAIABDACAAMAAuADAAMAAyACAALQAwAC4AMAAwADMAIAAwAC4AMAAxADIAIAAtADAALgAwADMANAAgADAALgAwADMANwAgAC0AMAAuADAAMwAyACAAQwAgADAALgAwADcANQAgAC0AMAAuADAAMgA5ACAAMAAuADAAOQAgAC0AMAAuADAAMAA3ACAAMAAuADEAMgA1ACAALQAwAC4AMAAyADkAIABDACAAMAAuADEANAA3ACAALQAwAC4AMAA0ADIAIAAwAC4AMQA3ADMAIAAtADAALgAwADcANQAgADAALgAxADkAMgAgAC0AMAAuADAANwA0ACAAQwAgADAALgAyADMANQAgAC0AMAAuADAANwAzACAAMAAuADIANAA0ACAALQAwAC4AMAAzADkAIAAwAC4AMgA0ADQAIAAtADAALgAwADAAOAAgAEMAIAAwAC4AMgA0ADUAIAAwAC4AMAAzADYAIAAwAC4AMQA4ADkAIAAwAC4AMAA3ADMAIAAwAC4AMQAyADEAIAAwAC4AMAA3ADcAIABDACAAMAAuADAANQAyACAAMAAuADAAOAAgAC0AMAAuADAAMAA1ACAAMAAuADAAMwAzACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[32] = [];
    PRESET_SUBTYPES = PRESET_TYPES[32] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;535;EwIAAPr7AAwCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQIAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTDAQAA+vsAvAEAAAEAAAAJswEAAPoAAQEBApcAAABNACAAMAAgADAAIABDACAALQAwAC4AMQAxADgAIAAtADAALgAxADEAOAAgADAALgAxADMAMgAgAC0AMAAuADEAMQA4ACAAMAAuADAAMQAxACAAMAAgAEMAIAAwAC4AMQAzADIAIAAtADAALgAxADEAOAAgADAALgAxADMAMgAgADAALgAxADMAMgAgADAALgAwADEAMQAgADAALgAwADEAMQAgAEMAIAAwAC4AMQAzADIAIAAwAC4AMQAzADIAIAAtADAALgAxADEAOAAgADAALgAxADMAMgAgADAAIAAwAC4AMAAxADEAIABDACAALQAwAC4AMQAxADgAIAAwAC4AMQAzADIAIAAtADAALgAxADEAOAAgAC0AMAAuADEAMQA4ACAAMAAgADAAIABaAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[33] = [];
    PRESET_SUBTYPES = PRESET_TYPES[33] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1193;pQQAAPr7AJ4EAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQIQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRVBAAA+vsATgQAAAEAAAAJRQQAAPoAAQEBAuABAABNACAAMAAgADAAIABDACAAMAAuADAAMQA1ACAAMAAuADAAMgA0ACAAMAAuADAAMwA3ACAAMAAuADAANAA5ACAAMAAuADAANQA1ACAAMAAuADAANQA5ACAAQwAgADAALgAwADgAMgAgADAALgAwADcANQAgADAALgAxADAAOAAgADAALgAwADgAMQAgADAALgAxADEAMwAgADAALgAwADcAMwAgAEMAIAAwAC4AMQAxADcAIAAwAC4AMAA2ADUAIAAwAC4AMAA5ADkAIAAwAC4AMAA0ADUAIAAwAC4AMAA3ADIAIAAwAC4AMAAyADkAIABDACAAMAAuADAANQA0ACAAMAAuADAAMQA5ACAAMAAuADAAMgAxACAAMAAuADAAMQAyACAALQAwAC4AMAAwADgAIAAwAC4AMAAxADEAIABDACAALQAwAC4AMAAzADYAIAAwAC4AMAAxADIAIAAtADAALgAwADcAIAAwAC4AMAAxADkAIAAtADAALgAwADgAOAAgADAALgAwADIAOQAgAEMAIAAtADAALgAxADEANQAgADAALgAwADQANQAgAC0AMAAuADEAMwAzACAAMAAuADAANgA1ACAALQAwAC4AMQAyADgAIAAwAC4AMAA3ADMAIABDACAALQAwAC4AMQAyADMAIAAwAC4AMAA4ADEAIAAtADAALgAwADkANwAgADAALgAwADcANQAgAC0AMAAuADAANwAxACAAMAAuADAANQA5ACAAQwAgAC0AMAAuADAANQAzACAAMAAuADAANAA5ACAALQAwAC4AMAAzACAAMAAuADAAMgA0ACAALQAwAC4AMAAxADYAIAAwACAAQwAgAC0AMAAuADAAMAAxACAALQAwAC4AMAAyADUAIAAwAC4AMAAwADkAIAAtADAALgAwADUAOAAgADAALgAwADAAOQAgAC0AMAAuADAANwA5ACAAQwAgADAALgAwADAAOQAgAC0AMAAuADEAMQAxACAAMAAuADAAMAAyACAALQAwAC4AMQAzADYAIAAtADAALgAwADAAOAAgAC0AMAAuADEAMwA2ACAAQwAgAC0AMAAuADAAMQA3ACAALQAwAC4AMQAzADYAIAAtADAALgAwADIANQAgAC0AMAAuADEAMQAxACAALQAwAC4AMAAyADUAIAAtADAALgAwADcAOQAgAEMAIAAtADAALgAwADIANQAgAC0AMAAuADAANQA4ACAALQAwAC4AMAAxADQAIAAtADAALgAwADIANQAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[34] = [];
    PRESET_SUBTYPES = PRESET_TYPES[34] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;2645;UQoAAPr7AEoKAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQIgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQBCgAA+vsA+gkAAAEAAAAJ8QkAAPoAAQEBArYEAABNACAAMAAgADAAIABDACAAMAAuADAAMAA0ACAALQAwAC4AMAAwADQAIAAwAC4AMAAxACAALQAwAC4AMAAwADYAIAAwAC4AMAAxADUAIAAtADAALgAwADAANgAgAEMAIAAwAC4AMAAyADIAIAAtADAALgAwADAANgAgADAALgAwADIAOQAgAC0AMAAuADAAMAAzACAAMAAuADAAMwAzACAAMAAuADAAMAAyACAAQwAgADAALgAwADUAIAAwAC4AMAAyADIAIAAwAC4AMAA2ADMAIAAwAC4AMAA2ADYAIAAwAC4AMAA2ADMAIAAwAC4AMQAxADgAIABDACAAMAAuADAANgAzACAAMAAuADEAMQA4ACAAMAAuADAANgAzACAAMAAuADEAMQA5ACAAMAAuADAANgAzACAAMAAuADEAMQA5ACAAQwAgADAALgAwADYAMwAgADAALgAxADEAOQAgADAALgAwADYAMwAgADAALgAxADIAIAAwAC4AMAA2ADMAIAAwAC4AMQAyACAAQwAgADAALgAwADYAMwAgADAALgAxADcAMgAgADAALgAwADUAIAAwAC4AMgAxADcAIAAwAC4AMAAzADMAIAAwAC4AMgAzADcAIABDACAAMAAuADAAMgA5ACAAMAAuADIANAAxACAAMAAuADAAMgAyACAAMAAuADIANAA0ACAAMAAuADAAMQA1ACAAMAAuADIANAA0ACAAQwAgADAALgAwADEAIAAwAC4AMgA0ADQAIAAwAC4AMAAwADQAIAAwAC4AMgA0ADIAIAAwACAAMAAuADIAMwA4ACAAQwAgAC0AMAAuADAAMAA0ACAAMAAuADIAMwA0ACAALQAwAC4AMAAwADYAIAAwAC4AMgAyADkAIAAtADAALgAwADAANgAgADAALgAyADIAMwAgAEMAIAAtADAALgAwADAANgAgADAALgAyADEANgAgAC0AMAAuADAAMAAzACAAMAAuADIAMQAgADAALgAwADAAMgAgADAALgAyADAANgAgAEMAIAAwAC4AMAAyADIAIAAwAC4AMQA4ADgAIAAwAC4AMAA2ADYAIAAwAC4AMQA3ADUAIAAwAC4AMQAxADgAIAAwAC4AMQA3ADUAIABDACAAMAAuADEAMQA4ACAAMAAuADEANwA1ACAAMAAuADEAMQA5ACAAMAAuADEANwA1ACAAMAAuADEAMQA5ACAAMAAuADEANwA1ACAAQwAgADAALgAxADEAOQAgADAALgAxADcANQAgADAALgAxADIAIAAwAC4AMQA3ADUAIAAwAC4AMQAyACAAMAAuADEANwA1ACAAQwAgADAALgAxADcAMgAgADAALgAxADcANQAgADAALgAyADEANwAgADAALgAxADgAOAAgADAALgAyADMANwAgADAALgAyADAANgAgAEMAIAAwAC4AMgA0ADEAIAAwAC4AMgAxACAAMAAuADIANAA0ACAAMAAuADIAMQA2ACAAMAAuADIANAA0ACAAMAAuADIAMgAzACAAQwAgADAALgAyADQANAAgADAALgAyADIAOQAgADAALgAyADQAMgAgADAALgAyADMANAAgADAALgAyADMAOAAgADAALgAyADMAOAAgAEMAIAAwAC4AMgAzADQAIAAwAC4AMgA0ADIAIAAwAC4AMgAyADkAIAAwAC4AMgA0ADQAIAAwAC4AMgAyADMAIAAwAC4AMgA0ADQAIABDACAAMAAuADIAMQA2ACAAMAAuADIANAA0ACAAMAAuADIAMQAgADAALgAyADQAMQAgADAALgAyADAANgAgADAALgAyADMANwAgAEMAIAAwAC4AMQA4ADgAIAAwAC4AMgAxADcAIAAwAC4AMQA3ADUAIAAwAC4AMQA3ADIAIAAwAC4AMQA3ADUAIAAwAC4AMQAyACAAQwAgADAALgAxADcANQAgADAALgAxADIAIAAwAC4AMQA3ADUAIAAwAC4AMQAxADkAIAAwAC4AMQA3ADUAIAAwAC4AMQAxADkAIABDACAAMAAuADEANwA1ACAAMAAuADEAMQA5ACAAMAAuADEANwA1ACAAMAAuADEAMQA4ACAAMAAuADEANwA1ACAAMAAuADEAMQA4ACAAQwAgADAALgAxADcANQAgADAALgAwADYANgAgADAALgAxADgAOAAgADAALgAwADIAMgAgADAALgAyADAANgAgADAALgAwADAAMQAgAEMAIAAwAC4AMgAxACAALQAwAC4AMAAwADMAIAAwAC4AMgAxADYAIAAtADAALgAwADAANgAgADAALgAyADIAMwAgAC0AMAAuADAAMAA2ACAAQwAgADAALgAyADIAOQAgAC0AMAAuADAAMAA2ACAAMAAuADIAMwA0ACAALQAwAC4AMAAwADQAIAAwAC4AMgAzADgAIAAwACAAQwAgADAALgAyADQAMgAgADAALgAwADAANAAgADAALgAyADQANAAgADAALgAwADEAIAAwAC4AMgA0ADQAIAAwAC4AMAAxADUAIABDACAAMAAuADIANAA0ACAAMAAuADAAMgAyACAAMAAuADIANAAxACAAMAAuADAAMgA4ACAAMAAuADIAMwA3ACAAMAAuADAAMwAzACAAQwAgADAALgAyADEANwAgADAALgAwADUAIAAwAC4AMQA3ADIAIAAwAC4AMAA2ADMAIAAwAC4AMQAyACAAMAAuADAANgAzACAAQwAgADAALgAxADIAIAAwAC4AMAA2ADMAIAAwAC4AMQAyACAAMAAuADAANgAzACAAMAAuADEAMQA5ACAAMAAuADAANgAzACAAQwAgADAALgAxADEAOQAgADAALgAwADYAMwAgADAALgAxADEAOAAgADAALgAwADYAMwAgADAALgAxADEAOAAgADAALgAwADYAMwAgAEMAIAAwAC4AMAA2ADYAIAAwAC4AMAA2ADMAIAAwAC4AMAAyADIAIAAwAC4AMAA1ACAAMAAuADAAMAAyACAAMAAuADAAMwAzACAAQwAgAC0AMAAuADAAMAAzACAAMAAuADAAMgA4ACAALQAwAC4AMAAwADYAIAAwAC4AMAAyADIAIAAtADAALgAwADAANgAgADAALgAwADEANQAgAEMAIAAtADAALgAwADAANgAgADAALgAwADEAIAAtADAALgAwADAANAAgADAALgAwADAANAAgADAAIAAwACAAWgADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[35] = [];
    PRESET_SUBTYPES = PRESET_TYPES[35] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;267;BwEAAPr7AAABAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQIwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS3AAAA+vsAsAAAAAEAAAAJpwAAAPoAAQEBAhEAAABNACAAMAAgADAAIABMACAALQAwAC4AMgA1ACAAMAAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[36] = [];
    PRESET_SUBTYPES = PRESET_TYPES[36] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;355;XwEAAPr7AFgBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQJAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQPAQAA+vsACAEAAAEAAAAJ/wAAAPoAAQEBAj0AAABNACAAMAAgADAAIABMACAAMAAgADAALgAxADIANQAgAEMAIAAwACAAMAAuADEAOAAxACAAMAAuADAANgA5ACAAMAAuADIANQAgADAALgAxADIANQAgADAALgAyADUAIABMACAAMAAuADIANQAgADAALgAyADUAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[37] = [];
    PRESET_SUBTYPES = PRESET_TYPES[37] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;441;tQEAAPr7AK4BAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQJQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRlAQAA+vsAXgEAAAEAAAAJVQEAAPoAAQEBAmgAAABNACAAMAAgADAAIABMACAAMAAuADAANgA3ACAAMAAuADAANAAgAEMAIAAwAC4AMAA4ADEAIAAwAC4AMAA0ADkAIAAwAC4AMQAwADIAIAAwAC4AMAA1ADQAIAAwAC4AMQAyADQAIAAwAC4AMAA1ADQAIABDACAAMAAuADEANAA5ACAAMAAuADAANQA0ACAAMAAuADEANgA5ACAAMAAuADAANAA5ACAAMAAuADEAOAAzACAAMAAuADAANAAgAEwAIAAwAC4AMgA1ACAAMAAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[38] = [];
    PRESET_SUBTYPES = PRESET_TYPES[38] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;629;cQIAAPr7AGoCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQJgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQhAgAA+vsAGgIAAAEAAAAJEQIAAPoAAQEBAsYAAABNACAAMAAgADAAIABMACAAMAAuADAAMQA2ACAAMAAuADAAOQA5ACAATAAgADAALgAwADMAMQAgADAAIABMACAAMAAuADAANAA3ACAAMAAuADAAOQA5ACAATAAgADAALgAwADYAMwAgADAAIABMACAAMAAuADAANwA4ACAAMAAuADAAOQA5ACAATAAgADAALgAwADkANAAgADAAIABMACAAMAAuADEAMAA5ACAAMAAuADAAOQA5ACAATAAgADAALgAxADIANQAgADAAIABMACAAMAAuADEANAAxACAAMAAuADAAOQA5ACAATAAgADAALgAxADUANgAgADAAIABMACAAMAAuADEANwAyACAAMAAuADAAOQA5ACAATAAgADAALgAxADgANwAgADAAIABMACAAMAAuADIAMAAzACAAMAAuADAAOQA5ACAATAAgADAALgAyADEAOQAgADAAIABMACAAMAAuADIAMwA0ACAAMAAuADAAOQA5ACAATAAgADAALgAyADUAIAAwACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[39] = [];
    PRESET_SUBTYPES = PRESET_TYPES[39] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;533;EQIAAPr7AAoCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQJwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTBAQAA+vsAugEAAAEAAAAJsQEAAPoAAQEBApYAAABNACAAMAAgADAAIABDACAAMAAgADAALgAwADMANQAgADAALgAwADIAOAAgADAALgAwADYAMgAgADAALgAwADYAMgAgADAALgAwADYAMgAgAEMAIAAwAC4AMAA5ADcAIAAwAC4AMAA2ADIAIAAwAC4AMQAyADUAIAAwAC4AMAAzADUAIAAwAC4AMQAyADUAIAAwACAAQwAgADAALgAxADIANQAgAC0AMAAuADAAMwA1ACAAMAAuADEANQAzACAALQAwAC4AMAA2ADIAIAAwAC4AMQA4ADgAIAAtADAALgAwADYAMgAgAEMAIAAwAC4AMgAyADIAIAAtADAALgAwADYAMgAgADAALgAyADUAIAAtADAALgAwADMANQAgADAALgAyADUAIAAwACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[40] = [];
    PRESET_SUBTYPES = PRESET_TYPES[40] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1419;hwUAAPr7AIAFAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQKAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQ3BQAA+vsAMAUAAAEAAAAJJwUAAPoAAQEBAlECAABNACAAMAAgADAAIABDACAAMAAuADAAMAAzACAALQAwAC4AMAAxADkAIAAwAC4AMAAwADcAIAAtADAALgAwADMANwAgADAALgAwADEANQAgAC0AMAAuADAAMwA3ACAAQwAgADAALgAwADIANAAgAC0AMAAuADAAMwA3ACAAMAAuADAAMgA3ACAALQAwAC4AMAAxADkAIAAwAC4AMAAzACAAMAAgAEMAIAAwAC4AMAAzADQAIAAwAC4AMAAyADEAIAAwAC4AMAAzADcAIAAwAC4AMAA0ADIAIAAwAC4AMAA0ADcAIAAwAC4AMAA0ADIAIABDACAAMAAuADAANQA2ACAAMAAuADAANAAyACAAMAAuADAANQA5ACAAMAAuADAAMgAxACAAMAAuADAANgAzACAAMAAgAEMAIAAwAC4AMAA2ADUAIAAtADAALgAwADEAOQAgADAALgAwADYAOQAgAC0AMAAuADAAMwA3ACAAMAAuADAANwA4ACAALQAwAC4AMAAzADcAIABDACAAMAAuADAAOAA2ACAALQAwAC4AMAAzADcAIAAwAC4AMAA5ACAALQAwAC4AMAAxADkAIAAwAC4AMAA5ADMAIAAwACAAQwAgADAALgAwADkANgAgADAALgAwADIAMQAgADAALgAxACAAMAAuADAANAAyACAAMAAuADEAMAA5ACAAMAAuADAANAAyACAAQwAgADAALgAxADEAOAAgADAALgAwADQAMgAgADAALgAxADIANQAgADAAIAAwAC4AMQAyADUAIAAwACAAQwAgADAALgAxADIAOAAgAC0AMAAuADAAMQA5ACAAMAAuADEAMwAxACAALQAwAC4AMAAzADcAIAAwAC4AMQA0ACAALQAwAC4AMAAzADcAIABDACAAMAAuADEANAA5ACAALQAwAC4AMAAzADcAIAAwAC4AMQA1ADIAIAAtADAALgAwADEAOQAgADAALgAxADUANQAgADAAIABDACAAMAAuADEANQA5ACAAMAAuADAAMgAxACAAMAAuADEANgAyACAAMAAuADAANAAyACAAMAAuADEANwAyACAAMAAuADAANAAyACAAQwAgADAALgAxADgAMQAgADAALgAwADQAMgAgADAALgAxADgANAAgADAALgAwADIAMQAgADAALgAxADgANwAgADAAIABDACAAMAAuADEAOQAxACAALQAwAC4AMAAxADkAIAAwAC4AMQA5ADQAIAAtADAALgAwADMANwAgADAALgAyADAAMwAgAC0AMAAuADAAMwA3ACAAQwAgADAALgAyADEAMQAgAC0AMAAuADAAMwA3ACAAMAAuADIAMQA1ACAALQAwAC4AMAAxADkAIAAwAC4AMgAxADgAIAAwACAAQwAgADAALgAyADIAMQAgADAALgAwADIAMQAgADAALgAyADIANQAgADAALgAwADQAMgAgADAALgAyADMANAAgADAALgAwADQAMgAgAEMAIAAwAC4AMgA0ADMAIAAwAC4AMAA0ADIAIAAwAC4AMgA0ADYAIAAwAC4AMAAyADEAIAAwAC4AMgA1ACAAMAAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[41] = [];
    PRESET_SUBTYPES = PRESET_TYPES[41] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1591;MwYAAPr7ACwGAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQKQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTjBQAA+vsA3AUAAAEAAAAJ0wUAAPoAAQEBAqcCAABNACAAMAAgADAAIABjACAALQAwAC4AMAAwADQAIAAtADAALgAwADAAOAAgAC0AMAAuADAAMQA4ACAALQAwAC4AMAAxADYAIAAtADAALgAwADIAMwAgAC0AMAAuADAAMQA2ACAAYwAgAC0AMAAuADAAMwAxACAAMAAgAC0AMAAuADAANgAzACAAMAAuADEAMgA1ACAALQAwAC4AMAA2ADMAIAAwAC4AMgA1ACAAYwAgADAAIAAtADAALgAwADYAMwAgAC0AMAAuADAAMQA2ACAALQAwAC4AMQAyADUAIAAtADAALgAwADMAMQAgAC0AMAAuADEAMgA1ACAAYwAgAC0AMAAuADAAMQA2ACAAMAAgAC0AMAAuADAAMwAxACAAMAAuADAANgAzACAALQAwAC4AMAAzADEAIAAwAC4AMQAyADUAIABjACAAMAAgAC0AMAAuADAAMwAxACAALQAwAC4AMAAwADgAIAAtADAALgAwADYAMwAgAC0AMAAuADAAMQA2ACAALQAwAC4AMAA2ADMAIABjACAALQAwAC4AMAAwADgAIAAwACAALQAwAC4AMAAxADYAIAAwAC4AMAAzADEAIAAtADAALgAwADEANgAgADAALgAwADYAMwAgAGMAIAAwACAALQAwAC4AMAAxADYAIAAtADAALgAwADAANAAgAC0AMAAuADAAMwAxACAALQAwAC4AMAAwADgAIAAtADAALgAwADMAMQAgAGMAIAAtADAALgAwADAANAAgADAAIAAtADAALgAwADAAOAAgADAALgAwADEANgAgAC0AMAAuADAAMAA4ACAAMAAuADAAMwAxACAAYwAgADAAIAAtADAALgAwADAAOAAgAC0AMAAuADAAMAAyACAALQAwAC4AMAAxADYAIAAtADAALgAwADAANAAgAC0AMAAuADAAMQA2ACAAYwAgAC0AMAAuADAAMAAxACAAMAAgAC0AMAAuADAAMAA0ACAAMAAuADAAMAA4ACAALQAwAC4AMAAwADQAIAAwAC4AMAAxADYAIABjACAAMAAgAC0AMAAuADAAMAA0ACAALQAwAC4AMAAwADEAIAAtADAALgAwADAAOAAgAC0AMAAuADAAMAAyACAALQAwAC4AMAAwADgAIABjACAAMAAgAC0AMAAuADAAMAAxACAALQAwAC4AMAAwADIAIAAwAC4AMAAwADQAIAAtADAALgAwADAAMgAgADAALgAwADAAOAAgAGMAIAAwACAALQAwAC4AMAAwADIAIAAwACAALQAwAC4AMAAwADQAIAAtADAALgAwADAAMQAgAC0AMAAuADAAMAA0ACAAYwAgADAAIAAwAC4AMAAwADEAIAAtADAALgAwADAAMQAgADAALgAwADAAMgAgAC0AMAAuADAAMAAxACAAMAAuADAAMAA0ACAAYwAgADAAIAAtADAALgAwADAAMQAgADAAIAAtADAALgAwADAAMgAgADAAIAAtADAALgAwADAAMwAgAGMAIAAtADAALgAwADAAMQAgADAAIAAtADAALgAwADAAMQAgADAALgAwADAAMQAgAC0AMAAuADAAMAAxACAAMAAuADAAMAAyACAAYwAgAC0AMAAuADAAMAAxACAAMAAgAC0AMAAuADAAMAAxACAALQAwAC4AMAAwADEAIAAtADAALgAwADAAMQAgAC0AMAAuADAAMAAyACAAYwAgAC0AMAAuADAAMAAxACAAMAAgAC0AMAAuADAAMAAxACAAMAAuADAAMAAxACAALQAwAC4AMAAwADEAIAAwAC4AMAAwADIAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[42] = [];
    PRESET_SUBTYPES = PRESET_TYPES[42] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;265;BQEAAPr7AP4AAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQKgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS1AAAA+vsArgAAAAEAAAAJpQAAAPoAAQEBAhAAAABNACAAMAAgADAAIABMACAAMAAgADAALgAyADUAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[43] = [];
    PRESET_SUBTYPES = PRESET_TYPES[43] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;361;ZQEAAPr7AF4BAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQKwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQVAQAA+vsADgEAAAEAAAAJBQEAAPoAAQEBAkAAAABNACAAMAAgADAAIABMACAAMAAuADEAMgA1ACAAMAAgAEMAIAAwAC4AMQA4ADEAIAAwACAAMAAuADIANQAgAC0AMAAuADAANgA5ACAAMAAuADIANQAgAC0AMAAuADEAMgA1ACAATAAgADAALgAyADUAIAAtADAALgAyADUAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[44] = [];
    PRESET_SUBTYPES = PRESET_TYPES[44] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;455;wwEAAPr7ALwBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQLAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRzAQAA+vsAbAEAAAEAAAAJYwEAAPoAAQEBAm8AAABNACAAMAAgADAAIABMACAAMAAuADAANgA3ACAALQAwAC4AMAA0ACAAQwAgADAALgAwADgAMQAgAC0AMAAuADAANAA5ACAAMAAuADEAMAAyACAALQAwAC4AMAA1ADQAIAAwAC4AMQAyADQAIAAtADAALgAwADUANAAgAEMAIAAwAC4AMQA0ADkAIAAtADAALgAwADUANAAgADAALgAxADYAOQAgAC0AMAAuADAANAA5ACAAMAAuADEAOAAzACAALQAwAC4AMAA0ACAATAAgADAALgAyADUAIAAwACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[45] = [];
    PRESET_SUBTYPES = PRESET_TYPES[45] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;995;3wMAAPr7ANgDAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQLQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wSPAwAA+vsAiAMAAAEAAAAJfwMAAPoAAQEBAn0BAABNACAAMAAgADAAIABMACAAMAAuADAAMQA3ACAAMAAgAEMAIAAwAC4AMAAyADUAIAAwACAAMAAuADAAMwA0ACAALQAwAC4AMAAxADQAIAAwAC4AMAA0ADIAIAAtADAALgAwADEANgAgAEMAIAAwAC4AMAA0ADgAIAAtADAALgAwADEANgAgADAALgAwADUAOQAgAC0AMAAuADAAMAAzACAAMAAuADAANgA0ACAALQAwAC4AMAAwADMAIABDACAAMAAuADAANwAxACAALQAwAC4AMAAwADMAIAAwAC4AMAA3ADgAIAAtADAALgAwADAANwAgADAALgAwADkAMQAgAC0AMAAuADAAMAA3ACAATAAgADAALgAxACAALQAwAC4AMQA2ADIAIABMACAAMAAuADEAMQAgADAALgAwADIANQAgAEwAIAAwAC4AMQAyADIAIAAwACAATAAgADAALgAxADMAMgAgAC0AMAAuADAAMAA3ACAATAAgADAALgAxADUANgAgAC0AMAAuADAAMAAxACAAQwAgADAALgAxADYANwAgAC0AMAAuADAAMAA0ACAAMAAuADEANwA2ACAALQAwAC4AMAAxADcAIAAwAC4AMQA4ADcAIAAtADAALgAwADIAMgAgAEMAIAAwAC4AMQA5ADEAIAAtADAALgAwADIAMwAgADAALgAyACAALQAwAC4AMAAyADQAIAAwAC4AMgAwADYAIAAtADAALgAwADIAMgAgAEMAIAAwAC4AMgAxADIAIAAtADAALgAwADIAIAAwAC4AMgAxADcAIAAtADAALgAwADAANgAgADAALgAyADEAOQAgAC0AMAAuADAAMAA1ACAAQwAgADAALgAyADIAMgAgAC0AMAAuADAAMAAxACAAMAAuADIAMgA5ACAALQAwAC4AMAAwADUAIAAwAC4AMgAzADMAIAAtADAALgAwADAAMwAgAEwAIAAwAC4AMgAzADkAIAAwACAATAAgADAALgAyADUAIAAwACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[46] = [];
    PRESET_SUBTYPES = PRESET_TYPES[46] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1489;zQUAAPr7AMYFAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQLgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wR9BQAA+vsAdgUAAAEAAAAJbQUAAPoAAQEBAnQCAABNACAAMAAgADAAIABDACAALQAwAC4AMAAwADQAIAAtADAALgAwADYANwAgADAALgAwADQANgAgAC0AMAAuADEAMgA1ACAAMAAuADEAMQAzACAALQAwAC4AMQAyADkAIABDACAAMAAuADEANwA3ACAALQAwAC4AMQAzADQAIAAwAC4AMgAzADcAIAAtADAALgAwADgAOQAgADAALgAyADQAMQAgAC0AMAAuADAAMgA0ACAAQwAgADAALgAyADQANgAgADAALgAwADMANgAgADAALgAyADAANAAgADAALgAwADkAMgAgADAALgAxADQANAAgADAALgAwADkANgAgAEMAIAAwAC4AMAA4ADkAIAAwAC4AMAA5ADkAIAAwAC4AMAAzADcAIAAwAC4AMAA2ADIAIAAwAC4AMAAzADMAIAAwAC4AMAAwADYAIABDACAAMAAuADAAMgA5ACAALQAwAC4AMAA0ADUAIAAwAC4AMAA2ADQAIAAtADAALgAwADkAMwAgADAALgAxADEANQAgAC0AMAAuADAAOQA3ACAAQwAgADAALgAxADYAMgAgAC0AMAAuADEAIAAwAC4AMgAwADYAIAAtADAALgAwADYAOQAgADAALgAyADAAOQAgAC0AMAAuADAAMgAyACAAQwAgADAALgAyADEAMgAgADAALgAwADIAIAAwAC4AMQA4ADQAIAAwAC4AMAA2ADEAIAAwAC4AMQA0ADIAIAAwAC4AMAA2ADMAIABDACAAMAAuADEAMAA0ACAAMAAuADAANgA2ACAAMAAuADAANgA4ACAAMAAuADAANAAyACAAMAAuADAANgA1ACAAMAAuADAAMAA0ACAAQwAgADAALgAwADYAMwAgAC0AMAAuADAAMwAgADAALgAwADgANAAgAC0AMAAuADAANgAzACAAMAAuADEAMQA3ACAALQAwAC4AMAA2ADUAIABDACAAMAAuADEANAA2ACAALQAwAC4AMAA2ADcAIAAwAC4AMQA3ADUAIAAtADAALgAwADQAOQAgADAALgAxADcANwAgAC0AMAAuADAAMgAgAEMAIAAwAC4AMQA3ADkAIAAwAC4AMAAwADUAIAAwAC4AMQA2ADQAIAAwAC4AMAAyADkAIAAwAC4AMQA0ACAAMAAuADAAMwAxACAAQwAgADAALgAxADIAIAAwAC4AMAAzADMAIAAwAC4AMAA5ADkAIAAwAC4AMAAyADIAIAAwAC4AMAA5ADgAIAAwAC4AMAAwADIAIABDACAAMAAuADAAOQA2ACAALQAwAC4AMAAxADQAIAAwAC4AMQAwADQAIAAtADAALgAwADMAMQAgADAALgAxADEAOQAgAC0AMAAuADAAMwAzACAAQwAgADAALgAxADMAMQAgAC0AMAAuADAAMwAzACAAMAAuADEANAAzACAALQAwAC4AMAAyADkAIAAwAC4AMQA0ADUAIAAtADAALgAwADEAOAAgAEMAIAAwAC4AMQA0ADYAIAAtADAALgAwADEAMQAgADAALgAxADQANAAgAC0AMAAuADAAMAA0ACAAMAAuADEAMwA4ACAALQAwAC4AMAAwADEAIABDACAAMAAuADEAMwA1ACAAMAAgADAALgAxADMAMwAgADAAIAAwAC4AMQAzACAALQAwAC4AMAAwADEAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[47] = [];
    PRESET_SUBTYPES = PRESET_TYPES[47] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1387;ZwUAAPr7AGAFAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQLwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQXBQAA+vsAEAUAAAEAAAAJBwUAAPoAAQEBAkECAABNACAAMAAgADAAIABDACAAMAAuADAAMAAyACAAMAAuADAANgAzACAAMAAuADAAMAA5ACAAMAAuADEAMAA4ACAAMAAuADAAMQA2ACAAMAAuADEAMAA4ACAAQwAgADAALgAwADIAMwAgADAALgAxADAAOAAgADAALgAwADIAOQAgADAALgAwADYAMwAgADAALgAwADMAMQAgADAAIABDACAAMAAuADAAMwA0ACAAMAAuADAANgAzACAAMAAuADAANAAgADAALgAxADAAOAAgADAALgAwADQANwAgADAALgAxADAAOAAgAEMAIAAwAC4AMAA1ADQAIAAwAC4AMQAwADgAIAAwAC4AMAA2ACAAMAAuADAANgAzACAAMAAuADAANgAyACAAMAAgAEMAIAAwAC4AMAA2ADUAIAAwAC4AMAA2ADMAIAAwAC4AMAA3ADEAIAAwAC4AMQAwADgAIAAwAC4AMAA3ADgAIAAwAC4AMQAwADgAIABDACAAMAAuADAAOAA1ACAAMAAuADEAMAA4ACAAMAAuADAAOQAyACAAMAAuADAANgAzACAAMAAuADAAOQA0ACAAMAAgAEMAIAAwAC4AMAA5ADYAIAAwAC4AMAA2ADMAIAAwAC4AMQAwADIAIAAwAC4AMQAwADgAIAAwAC4AMQAxACAAMAAuADEAMAA4ACAAQwAgADAALgAxADEANgAgADAALgAxADAAOAAgADAALgAxADIAMwAgADAALgAwADYAMwAgADAALgAxADIANQAgADAAIABDACAAMAAuADEAMgA3ACAAMAAuADAANgAzACAAMAAuADEAMwA0ACAAMAAuADEAMAA4ACAAMAAuADEANAAxACAAMAAuADEAMAA4ACAAQwAgADAALgAxADQAOAAgADAALgAxADAAOAAgADAALgAxADUANAAgADAALgAwADYAMwAgADAALgAxADUANgAgADAAIABDACAAMAAuADEANQA5ACAAMAAuADAANgAzACAAMAAuADEANgA1ACAAMAAuADEAMAA4ACAAMAAuADEANwAyACAAMAAuADEAMAA4ACAAQwAgADAALgAxADcAOQAgADAALgAxADAAOAAgADAALgAxADgANQAgADAALgAwADYAMwAgADAALgAxADgAOAAgADAAIABDACAAMAAuADEAOQAgADAALgAwADYAMwAgADAALgAxADkANgAgADAALgAxADAAOAAgADAALgAyADAAMwAgADAALgAxADAAOAAgAEMAIAAwAC4AMgAxACAAMAAuADEAMAA4ACAAMAAuADIAMQA3ACAAMAAuADAANgAzACAAMAAuADIAMQA5ACAAMAAgAEMAIAAwAC4AMgAyADEAIAAwAC4AMAA2ADMAIAAwAC4AMgAyADcAIAAwAC4AMQAwADgAIAAwAC4AMgAzADUAIAAwAC4AMQAwADgAIABDACAAMAAuADIANAAyACAAMAAuADEAMAA4ACAAMAAuADIANAA4ACAAMAAuADAANgAzACAAMAAuADIANQAgADAAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[48] = [];
    PRESET_SUBTYPES = PRESET_TYPES[48] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1997;yQcAAPr7AMIHAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQMAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wR5BwAA+vsAcgcAAAEAAAAJaQcAAPoAAQEBAnIDAABNACAAMAAgADAAIABDACAAMAAuADAAMAA4ACAAMAAuADAAMAA4ACAAMAAuADAAMQA3ACAAMAAuADAAMQA2ACAAMAAuADAAMgAxACAAMAAuADAAMgA2ACAAQwAgADAALgAwADIANQAgADAALgAwADMANwAgADAALgAwADIANwAgADAALgAwADUAIAAwAC4AMAAyADkAIAAwAC4AMAA2ADMAIABDACAAMAAuADAAMwAxACAAMAAuADAANwA2ACAAMAAuADAAMgA5ACAAMAAuADAAOAA3ACAAMAAuADAAMgA3ACAAMAAuADAAOQA5ACAAQwAgADAALgAwADIANQAgADAALgAxADEAIAAwAC4AMAAyADIAIAAwAC4AMQAyADIAIAAwAC4AMAAxADUAIAAwAC4AMQAzADIAIABDACAAMAAuADAAMAA5ACAAMAAuADEANAAyACAALQAwAC4AMAAwADEAIAAwAC4AMQA1ACAALQAwAC4AMAAxADIAIAAwAC4AMQA1ADYAIABDACAALQAwAC4AMAAyADIAIAAwAC4AMQA2ADIAIAAtADAALgAwADMANAAgADAALgAxADYANgAgAC0AMAAuADAANAA2ACAAMAAuADEANgA4ACAAQwAgAC0AMAAuADAANQA4ACAAMAAuADEANwAgAC0AMAAuADAANwAgADAALgAxADcAIAAtADAALgAwADgAMQAgADAALgAxADYAOAAgAEMAIAAtADAALgAwADkAMwAgADAALgAxADYANgAgAC0AMAAuADEAMAA0ACAAMAAuADEANgAxACAALQAwAC4AMQAxADMAIAAwAC4AMQA1ADMAIABDACAALQAwAC4AMQAyADIAIAAwAC4AMQA0ADYAIAAtADAALgAxADMAIAAwAC4AMQAzADcAIAAtADAALgAxADMANAAgADAALgAxADIANgAgAEMAIAAtADAALgAxADMAOQAgADAALgAxADEANgAgAC0AMAAuADEANAAxACAAMAAuADEAMAAyACAALQAwAC4AMQA0ADEAIAAwAC4AMAA5ADEAIABDACAALQAwAC4AMQA0ADIAIAAwAC4AMAA4ACAALQAwAC4AMQA0ADEAIAAwAC4AMAA2ADcAIAAtADAALgAxADMANgAgADAALgAwADUANgAgAEMAIAAtADAALgAxADMAMQAgADAALgAwADQANgAgAC0AMAAuADEAMgAyACAAMAAuADAAMwA4ACAALQAwAC4AMQAxACAAMAAuADAAMwA0ACAAQwAgAC0AMAAuADAAOQA4ACAAMAAuADAAMwAxACAALQAwAC4AMAA4ADYAIAAwAC4AMAAzADUAIAAtADAALgAwADcAOAAgADAALgAwADQAMgAgAEMAIAAtADAALgAwADcAMQAgADAALgAwADQAOQAgAC0AMAAuADAANgA2ACAAMAAuADAANgAgAC0AMAAuADAANgA1ACAAMAAuADAANwAzACAAQwAgAC0AMAAuADAANgA1ACAAMAAuADAAOAA2ACAALQAwAC4AMAA2ADYAIAAwAC4AMAA5ADgAIAAtADAALgAwADcAMQAgADAALgAxADAAOAAgAEMAIAAtADAALgAwADcANgAgADAALgAxADEAOAAgAC0AMAAuADAANwA1ACAAMAAuADEAMgAgAC0AMAAuADAAOQA1ACAAMAAuADEAMwAzACAAQwAgAC0AMAAuADEAMQAzACAAMAAuADEANAA3ACAALQAwAC4AMQAzADEAIAAwAC4AMQA0ADMAIAAtADAALgAxADQAMgAgADAALgAxADQANAAgAEMAIAAtADAALgAxADUAMwAgADAALgAxADQANAAgAC0AMAAuADEANgAyACAAMAAuADEANAAgAC0AMAAuADEANwAzACAAMAAuADEAMwA2ACAAQwAgAC0AMAAuADEAOAA1ACAAMAAuADEAMwAxACAALQAwAC4AMQA5ADUAIAAwAC4AMQAyADIAIAAtADAALgAyADAAMgAgADAALgAxADEANAAgAEMAIAAtADAALgAyADAAOQAgADAALgAxADAANgAgAC0AMAAuADIAMQAyACAAMAAuADAAOQA2ACAALQAwAC4AMgAxADYAIAAwAC4AMAA4ACAAQwAgAC0AMAAuADIAMQA5ACAAMAAuADAANgA0ACAALQAwAC4AMgAxADkAIAAwAC4AMAA1ADYAIAAtADAALgAyADEAOQAgADAALgAwADQANAAgAEMAIAAtADAALgAyADEAOQAgADAALgAwADMAMgAgAC0AMAAuADIAMQA5ACAAMAAuADAAMgAgAC0AMAAuADIAMQA5ACAAMAAuADAAMAA4ACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[49] = [];
    PRESET_SUBTYPES = PRESET_TYPES[49] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;271;CwEAAPr7AAQBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQMQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS7AAAA+vsAtAAAAAEAAAAJqwAAAPoAAQEBAhMAAABNACAAMAAgADAAIABMACAAMAAuADIANQAgADAALgAyADUAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[50] = [];
    PRESET_SUBTYPES = PRESET_TYPES[50] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;355;XwEAAPr7AFgBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQMgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQPAQAA+vsACAEAAAEAAAAJ/wAAAPoAAQEBAj0AAABNACAAMAAgADAAIABMACAAMAAuADEAMgA1ACAAMAAgAEMAIAAwAC4AMQA4ADEAIAAwACAAMAAuADIANQAgADAALgAwADYAOQAgADAALgAyADUAIAAwAC4AMQAyADUAIABMACAAMAAuADIANQAgADAALgAyADUAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[51] = [];
    PRESET_SUBTYPES = PRESET_TYPES[51] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;455;wwEAAPr7ALwBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQMwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRzAQAA+vsAbAEAAAEAAAAJYwEAAPoAAQEBAm8AAABNACAAMAAgADAAIABMACAALQAwAC4AMAA0ACAAMAAuADAANgA3ACAAQwAgAC0AMAAuADAANAA5ACAAMAAuADAAOAAxACAALQAwAC4AMAA1ADQAIAAwAC4AMQAwADIAIAAtADAALgAwADUANAAgADAALgAxADIANAAgAEMAIAAtADAALgAwADUANAAgADAALgAxADQAOQAgAC0AMAAuADAANAA5ACAAMAAuADEANgA5ACAALQAwAC4AMAA0ACAAMAAuADEAOAAzACAATAAgADAAIAAwAC4AMgA1ACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[52] = [];
    PRESET_SUBTYPES = PRESET_TYPES[52] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;5197;SRQAAPr7AEIUAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQNAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wT5EwAA+vsA8hMAAAEAAAAJ6RMAAPoAAQEBArIJAABNACAAMAAgADAAIABDACAALQAwAC4AMAAwADEAIAAwAC4AMAAyADUAIAAwAC4AMAA2ACAAMAAuADAANAA3ACAAMAAuADEAMwA3ACAAMAAuADAANAA4ACAAQwAgADAALgAxADkAOAAgADAALgAwADUAIAAwAC4AMgA0ADgAIAAwAC4AMAAzADgAIAAwAC4AMgA0ADkAIAAwAC4AMAAyADMAIABDACAAMAAuADIANAA5ACAAMAAuADAAMAA4ACAAMAAuADIAIAAtADAALgAwADAANgAgADAALgAxADMAOAAgAC0AMAAuADAAMAA3ACAAQwAgADAALgAxADAANwAgAC0AMAAuADAAMAA3ACAAMAAuADAANwA5ACAALQAwAC4AMAAwADUAIAAwAC4AMAA1ADkAIAAwACAAQwAgADAALgAwADMAIAAwAC4AMAAwADcAIAAwAC4AMAAxADMAIAAwAC4AMAAxADgAIAAwAC4AMAAxADMAIAAwAC4AMAAzADEAIABDACAAMAAuADAAMQAzACAAMAAuADAAMwA4ACAAMAAuADAAMQA4ACAAMAAuADAANAA1ACAAMAAuADAAMgA3ACAAMAAuADAANQAxACAAQwAgADAALgAwADQAOAAgADAALgAwADYANAAgADAALgAwADgAOQAgADAALgAwADcAMwAgADAALgAxADMANgAgADAALgAwADcANAAgAEMAIAAwAC4AMQA5ADEAIAAwAC4AMAA3ADYAIAAwAC4AMgAzADYAIAAwAC4AMAA2ADUAIAAwAC4AMgAzADYAIAAwAC4AMAA1ADIAIABDACAAMAAuADIAMwA3ACAAMAAuADAAMwA4ACAAMAAuADEAOQAyACAAMAAuADAAMgA2ACAAMAAuADEAMwA3ACAAMAAuADAAMgA0ACAAQwAgADAALgAxADAAOQAgADAALgAwADIANAAgADAALgAwADgANAAgADAALgAwADIANgAgADAALgAwADYANQAgADAALgAwADMAIABDACAAMAAuADAANAAgADAALgAwADMANwAgADAALgAwADIANAAgADAALgAwADQAOAAgADAALgAwADIANAAgADAALgAwADUAOQAgAEMAIAAwAC4AMAAyADQAIAAwAC4AMAA2ADUAIAAwAC4AMAAyADkAIAAwAC4AMAA3ADEAIAAwAC4AMAAzADcAIAAwAC4AMAA3ADcAIABDACAAMAAuADAANQA2ACAAMAAuADAAOAA4ACAAMAAuADAAOQAyACAAMAAuADAAOQA3ACAAMAAuADEAMwA1ACAAMAAuADAAOQA4ACAAQwAgADAALgAxADgANQAgADAALgAwADkAOQAgADAALgAyADIANQAgADAALgAwADgAOQAgADAALgAyADIANQAgADAALgAwADcANwAgAEMAIAAwAC4AMgAyADYAIAAwAC4AMAA2ADUAIAAwAC4AMQA4ADYAIAAwAC4AMAA1ADQAIAAwAC4AMQAzADYAIAAwAC4AMAA1ADMAIABDACAAMAAuADEAMQAxACAAMAAuADAANQAyACAAMAAuADAAOAA4ACAAMAAuADAANQA0ACAAMAAuADAANwAxACAAMAAuADAANQA4ACAAQwAgADAALgAwADQAOAAgADAALgAwADYANAAgADAALgAwADMANQAgADAALgAwADcAMwAgADAALgAwADMANQAgADAALgAwADgANAAgAEMAIAAwAC4AMAAzADUAIAAwAC4AMAA4ADkAIAAwAC4AMAAzADkAIAAwAC4AMAA5ADUAIAAwAC4AMAA0ADYAIAAwAC4AMQAgAEMAIAAwAC4AMAA2ADMAIAAwAC4AMQAxACAAMAAuADAAOQA2ACAAMAAuADEAMQA4ACAAMAAuADEAMwA0ACAAMAAuADEAMQA5ACAAQwAgADAALgAxADcAOQAgADAALgAxADEAOQAgADAALgAyADEANQAgADAALgAxADEAMQAgADAALgAyADEANQAgADAALgAxACAAQwAgADAALgAyADEANQAgADAALgAwADgAOQAgADAALgAxADgAIAAwAC4AMAA3ADkAIAAwAC4AMQAzADUAIAAwAC4AMAA3ADgAIABDACAAMAAuADEAMQAzACAAMAAuADAANwA4ACAAMAAuADAAOQAyACAAMAAuADAAOAAgADAALgAwADcANwAgADAALgAwADgAMwAgAEMAIAAwAC4AMAA1ADYAIAAwAC4AMAA4ADgAIAAwAC4AMAA0ADQAIAAwAC4AMAA5ADcAIAAwAC4AMAA0ADMAIAAwAC4AMQAwADYAIABDACAAMAAuADAANAAzACAAMAAuADEAMQAxACAAMAAuADAANAA4ACAAMAAuADEAMQA2ACAAMAAuADAANQA0ACAAMAAuADEAMgAgAEMAIAAwAC4AMAA2ADkAIAAwAC4AMQAzACAAMAAuADAAOQA5ACAAMAAuADEAMwA3ACAAMAAuADEAMwAzACAAMAAuADEAMwA3ACAAQwAgADAALgAxADcAMwAgADAALgAxADMAOAAgADAALgAyADAANgAgADAALgAxADMAMQAgADAALgAyADAANgAgADAALgAxADIAMQAgAEMAIAAwAC4AMgAwADcAIAAwAC4AMQAxADEAIAAwAC4AMQA3ADQAIAAwAC4AMQAwADIAIAAwAC4AMQAzADQAIAAwAC4AMQAwADEAIABDACAAMAAuADEAMQA0ACAAMAAuADEAMAAxACAAMAAuADAAOQA1ACAAMAAuADEAMAAyACAAMAAuADAAOAAyACAAMAAuADEAMAA2ACAAQwAgADAALgAwADYAMwAgADAALgAxADEAIAAwAC4AMAA1ADIAIAAwAC4AMQAxADgAIAAwAC4AMAA1ADIAIAAwAC4AMQAyADYAIABDACAAMAAuADAANQAyACAAMAAuADEAMwAxACAAMAAuADAANQA1ACAAMAAuADEAMwA1ACAAMAAuADAANgAxACAAMAAuADEAMwA5ACAAQwAgADAALgAwADcANQAgADAALgAxADQAOAAgADAALgAxADAAMQAgADAALgAxADUANAAgADAALgAxADMAMgAgADAALgAxADUANQAgAEMAIAAwAC4AMQA2ADkAIAAwAC4AMQA1ADUAIAAwAC4AMQA5ADgAIAAwAC4AMQA0ADkAIAAwAC4AMQA5ADgAIAAwAC4AMQA0ACAAQwAgADAALgAxADkAOQAgADAALgAxADMAMQAgADAALgAxADcAIAAwAC4AMQAyADMAIAAwAC4AMQAzADMAIAAwAC4AMQAyADIAIABDACAAMAAuADEAMQA1ACAAMAAuADEAMgAyACAAMAAuADAAOQA5ACAAMAAuADEAMgAzACAAMAAuADAAOAA3ACAAMAAuADEAMgA2ACAAQwAgADAALgAwADcAIAAwAC4AMQAzACAAMAAuADAANgAgADAALgAxADMANwAgADAALgAwADYAIAAwAC4AMQA0ADUAIABDACAAMAAuADAANgAgADAALgAxADQAOQAgADAALgAwADYAMwAgADAALgAxADUAMgAgADAALgAwADYAOAAgADAALgAxADUANgAgAEMAIAAwAC4AMAA4ACAAMAAuADEANgA0ACAAMAAuADEAMAA0ACAAMAAuADEANgA5ACAAMAAuADEAMwAyACAAMAAuADEANwAgAEMAIAAwAC4AMQA2ADUAIAAwAC4AMQA3ADEAIAAwAC4AMQA5ADEAIAAwAC4AMQA2ADUAIAAwAC4AMQA5ADEAIAAwAC4AMQA1ADYAIABDACAAMAAuADEAOQAxACAAMAAuADEANAA5ACAAMAAuADEANgA2ACAAMAAuADEANAAxACAAMAAuADEAMwAzACAAMAAuADEANAAxACAAQwAgADAALgAxADEANgAgADAALgAxADQAIAAwAC4AMQAwADEAIAAwAC4AMQA0ADIAIAAwAC4AMAA5ACAAMAAuADEANAA0ACAAQwAgADAALgAwADcANQAgADAALgAxADQAOAAgADAALgAwADYANgAgADAALgAxADUANAAgADAALgAwADYANgAgADAALgAxADYAMQAgAEMAIAAwAC4AMAA2ADYAIAAwAC4AMQA2ADUAIAAwAC4AMAA2ADkAIAAwAC4AMQA2ADgAIAAwAC4AMAA3ADQAIAAwAC4AMQA3ADEAIABDACAAMAAuADAAOAA1ACAAMAAuADEANwA4ACAAMAAuADEAMAA3ACAAMAAuADEAOAAzACAAMAAuADEAMwAxACAAMAAuADEAOAA0ACAAQwAgADAALgAxADYAMQAgADAALgAxADgANQAgADAALgAxADgANQAgADAALgAxADcAOQAgADAALgAxADgANQAgADAALgAxADcAMgAgAEMAIAAwAC4AMQA4ADUAIAAwAC4AMQA2ADQAIAAwAC4AMQA2ADEAIAAwAC4AMQA1ADgAIAAwAC4AMQAzADIAIAAwAC4AMQA1ADcAIABDACAAMAAuADEAMQA4ACAAMAAuADEANQA3ACAAMAAuADEAMAA0ACAAMAAuADEANQA4ACAAMAAuADAAOQA0ACAAMAAuADEANgAxACAAQwAgADAALgAwADgAIAAwAC4AMQA2ADQAIAAwAC4AMAA3ADIAIAAwAC4AMQA2ADkAIAAwAC4AMAA3ADIAIAAwAC4AMQA3ADYAIABDACAAMAAuADAANwAyACAAMAAuADEANwA5ACAAMAAuADAANwA1ACAAMAAuADEAOAAyACAAMAAuADAANwA5ACAAMAAuADEAOAA1ACAAQwAgADAALgAwADgAOQAgADAALgAxADkAMQAgADAALgAxADAAOAAgADAALgAxADkANgAgADAALgAxADMAMQAgADAALgAxADkANgAgAEMAIAAwAC4AMQA1ADcAIAAwAC4AMQA5ADcAIAAwAC4AMQA3ADkAIAAwAC4AMQA5ADIAIAAwAC4AMQA3ADkAIAAwAC4AMQA4ADUAIABDACAAMAAuADEANwA5ACAAMAAuADEANwA5ACAAMAAuADEANQA4ACAAMAAuADEANwAzACAAMAAuADEAMwAxACAAMAAuADEANwAzACAAQwAgADAALgAxADEAOQAgADAALgAxADcAMgAgADAALgAxADAANgAgADAALgAxADcAMwAgADAALgAwADkANwAgADAALgAxADcANQAgAEMAIAAwAC4AMAA4ADUAIAAwAC4AMQA3ADkAIAAwAC4AMAA3ADgAIAAwAC4AMQA4ADQAIAAwAC4AMAA3ADgAIAAwAC4AMQA4ADkAIABDACAAMAAuADAANwA4ACAAMAAuADEAOQAyACAAMAAuADAAOAAgADAALgAxADkANQAgADAALgAwADgANAAgADAALgAxADkANwAgAEMAIAAwAC4AMAA5ADMAIAAwAC4AMgAwADMAIAAwAC4AMQAxACAAMAAuADIAMAA3ACAAMAAuADEAMwAxACAAMAAuADIAMAA4ACAAQwAgADAALgAxADUANQAgADAALgAyADAAOAAgADAALgAxADcANAAgADAALgAyADAAMwAgADAALgAxADcANAAgADAALgAxADkAOAAgAEMAIAAwAC4AMQA3ADQAIAAwAC4AMQA5ADIAIAAwAC4AMQA1ADUAIAAwAC4AMQA4ADYAIAAwAC4AMQAzADEAIAAwAC4AMQA4ADYAIABDACAAMAAuADEAMQA5ACAAMAAuADEAOAA2ACAAMAAuADEAMAA4ACAAMAAuADEAOAA3ACAAMAAuADEAMAAxACAAMAAuADEAOAA5ACAAQwAgADAALgAwADgAOQAgADAALgAxADkAMQAgADAALgAwADgAMwAgADAALgAxADkANgAgADAALgAwADgAMwAgADAALgAyADAAMQAgAEMAIAAwAC4AMAA4ADMAIAAwAC4AMgAwADMAIAAwAC4AMAA4ADUAIAAwAC4AMgAwADYAIAAwAC4AMAA4ADgAIAAwAC4AMgAwADgAIABDACAAMAAuADAAOQA2ACAAMAAuADIAMQA0ACAAMAAuADEAMQAyACAAMAAuADIAMQA3ACAAMAAuADEAMwAgADAALgAyADEAOAAgAEMAIAAwAC4AMQA1ADIAIAAwAC4AMgAxADgAIAAwAC4AMQA2ADkAIAAwAC4AMgAxADQAIAAwAC4AMQA2ADkAIAAwAC4AMgAwADkAIABDACAAMAAuADEANgA5ACAAMAAuADIAMAAzACAAMAAuADEANQAyACAAMAAuADEAOQA5ACAAMAAuADEAMwAxACAAMAAuADEAOQA4ACAAQwAgADAALgAxADIAIAAwAC4AMQA5ADgAIAAwAC4AMQAxACAAMAAuADEAOQA5ACAAMAAuADEAMAAzACAAMAAuADIAMAAxACAAQwAgADAALgAwADkAMwAgADAALgAyADAAMwAgADAALgAwADgANwAgADAALgAyADAANwAgADAALgAwADgANwAgADAALgAyADEAMgAgAEMAIAAwAC4AMAA4ADcAIAAwAC4AMgAxADQAIAAwAC4AMAA4ADkAIAAwAC4AMgAxADYAIAAwAC4AMAA5ADIAIAAwAC4AMgAxADgAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[53] = [];
    PRESET_SUBTYPES = PRESET_TYPES[53] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;2731;pwoAAPr7AKAKAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQNQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRXCgAA+vsAUAoAAAEAAAAJRwoAAPoAAQEBAuEEAABNACAAMAAgADAAIABDACAALQAwAC4AMAA2ADYAIAAwAC4AMAAwADYAIAAtADAALgAxADEANQAgADAALgAwADIAMQAgAC0AMAAuADEAMQA1ACAAMAAuADAAMwAzACAAQwAgAC0AMAAuADEAMQA1ACAAMAAuADAANAA0ACAALQAwAC4AMAA2ADcAIAAwAC4AMAA1ADIAIAAtADAALgAwADAAMwAgADAALgAwADUAMgAgAEMAIAAwAC4AMAA2ADEAIAAwAC4AMAA1ADIAIAAwAC4AMQAxADUAIAAwAC4AMAA0ADQAIAAwAC4AMQAxADUAIAAwAC4AMAAzADMAIABDACAAMAAuADEAMQA1ACAAMAAuADAAMgAxACAAMAAuADAANQA5ACAAMAAuADAAMQA4ACAALQAwAC4AMAAwADUAIAAwAC4AMAAyADYAIABDACAALQAwAC4AMAA2ADgAIAAwAC4AMAAzADUAIAAtADAALgAxADEANQAgADAALgAwADUAIAAtADAALgAxADEANQAgADAALgAwADYAMQAgAEMAIAAtADAALgAxADEANQAgADAALgAwADcAMgAgAC0AMAAuADAANgA2ACAAMAAuADAAOAAxACAALQAwAC4AMAAwADMAIAAwAC4AMAA4ADEAIABDACAAMAAuADAANgAxACAAMAAuADAAOAAxACAAMAAuADEAMQA1ACAAMAAuADAANwAyACAAMAAuADEAMQA1ACAAMAAuADAANgAxACAAQwAgADAALgAxADEANQAgADAALgAwADUAIAAwAC4AMAA1ADkAIAAwAC4AMAA0ADcAIAAtADAALgAwADAANAAgADAALgAwADUANQAgAEMAIAAtADAALgAwADYAOAAgADAALgAwADYAMwAgAC0AMAAuADEAMQA1ACAAMAAuADAANwA4ACAALQAwAC4AMQAxADUAIAAwAC4AMAA4ADkAIABDACAALQAwAC4AMQAxADUAIAAwAC4AMQAwADEAIAAtADAALgAwADYANgAgADAALgAxADEAIAAtADAALgAwADAAMgAgADAALgAxADEAIABDACAAMAAuADAANgAxACAAMAAuADEAMQAgADAALgAxADEANQAgADAALgAxADAAMQAgADAALgAxADEANQAgADAALgAwADgAOQAgAEMAIAAwAC4AMQAxADUAIAAwAC4AMAA3ADkAIAAwAC4AMAA1ADkAIAAwAC4AMAA3ADYAIAAtADAALgAwADAANAAgADAALgAwADgAMwAgAEMAIAAtADAALgAwADYANwAgADAALgAwADkAMQAgAC0AMAAuADEAMQA1ACAAMAAuADEAMAA3ACAALQAwAC4AMQAxADUAIAAwAC4AMQAxADgAIABDACAALQAwAC4AMQAxADUAIAAwAC4AMQAyADkAIAAtADAALgAwADYANQAgADAALgAxADMAOAAgAC0AMAAuADAAMAAyACAAMAAuADEAMwA4ACAAQwAgADAALgAwADYAMwAgADAALgAxADMAOAAgADAALgAxADEANQAgADAALgAxADIAOQAgADAALgAxADEANQAgADAALgAxADEAOAAgAEMAIAAwAC4AMQAxADUAIAAwAC4AMQAwADcAIAAwAC4AMAA2ACAAMAAuADEAMAA0ACAALQAwAC4AMAAwADMAIAAwAC4AMQAxADIAIABDACAALQAwAC4AMAA2ADYAIAAwAC4AMQAyACAALQAwAC4AMQAxADUAIAAwAC4AMQAzADUAIAAtADAALgAxADEANQAgADAALgAxADQANgAgAEMAIAAtADAALgAxADEANQAgADAALgAxADUAOAAgAC0AMAAuADAANgA1ACAAMAAuADEANgA2ACAALQAwAC4AMAAwADEAIAAwAC4AMQA2ADYAIABDACAAMAAuADAANgAzACAAMAAuADEANgA2ACAAMAAuADEAMQA1ACAAMAAuADEANQA3ACAAMAAuADEAMQA1ACAAMAAuADEANAA2ACAAQwAgADAALgAxADEANQAgADAALgAxADMANQAgADAALgAwADYAIAAwAC4AMQAzADIAIAAtADAALgAwADAAMwAgADAALgAxADQAIABDACAALQAwAC4AMAA2ADYAIAAwAC4AMQA0ADgAIAAtADAALgAxADEANQAgADAALgAxADYANAAgAC0AMAAuADEAMQA1ACAAMAAuADEANwA0ACAAQwAgAC0AMAAuADEAMQA1ACAAMAAuADEAOAA1ACAALQAwAC4AMAA2ADQAIAAwAC4AMQA5ADQAIAAtADAALgAwADAAMQAgADAALgAxADkANAAgAEMAIAAwAC4AMAA2ADMAIAAwAC4AMQA5ADQAIAAwAC4AMQAxADUAIAAwAC4AMQA4ADUAIAAwAC4AMQAxADUAIAAwAC4AMQA3ADQAIABDACAAMAAuADEAMQA1ACAAMAAuADEANgA0ACAAMAAuADAANgAxACAAMAAuADEANgAxACAALQAwAC4AMAAwADMAIAAwAC4AMQA2ADgAIABDACAALQAwAC4AMAA2ADYAIAAwAC4AMQA3ADYAIAAtADAALgAxADEANQAgADAALgAxADkAMgAgAC0AMAAuADEAMQA1ACAAMAAuADIAMAAzACAAQwAgAC0AMAAuADEAMQA1ACAAMAAuADIAMQAzACAALQAwAC4AMAA2ADQAIAAwAC4AMgAyADMAIAAwACAAMAAuADIAMgAzACAAQwAgADAALgAwADYANAAgADAALgAyADIAMwAgADAALgAxADEANQAgADAALgAyADEANAAgADAALgAxADEANQAgADAALgAyADAAMwAgAEMAIAAwAC4AMQAxADUAIAAwAC4AMQA5ADIAIAAwAC4AMAA2ADEAIAAwAC4AMQA4ADkAIAAtADAALgAwADAAMgAgADAALgAxADkANwAgAEMAIAAtADAALgAwADYANQAgADAALgAyADAANQAgAC0AMAAuADEAMQA2ACAAMAAuADIAMgAgAC0AMAAuADEAMQA1ACAAMAAuADIAMwAxACAAQwAgAC0AMAAuADEAMQA0ACAAMAAuADIANAAyACAALQAwAC4AMAA2ADQAIAAwAC4AMgA1ACAAMAAgADAALgAyADUAIABDACAAMAAuADAANgA0ACAAMAAuADIANQAgADAALgAxADEANQAgADAALgAyADQAMQAgADAALgAxADEANQAgADAALgAyADMAIABDACAAMAAuADEAMQA1ACAAMAAuADIAMgAgADAALgAwADYAMwAgADAALgAyADEANwAgADAAIAAwAC4AMgAyADYAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[54] = [];
    PRESET_SUBTYPES = PRESET_TYPES[54] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1505;3QUAAPr7ANYFAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQNgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wSNBQAA+vsAhgUAAAEAAAAJfQUAAPoAAQEBAnwCAABNACAAMAAgADAAIABjACAAMAAuADAAMAA0ACAALQAwAC4AMAAwADgAIAAwAC4AMAAxADgAIAAtADAALgAwADEANgAgADAALgAwADIAMwAgAC0AMAAuADAAMQA2ACAAYwAgADAALgAwADMAMQAgADAAIAAwAC4AMAA2ADMAIAAwAC4AMQAyADUAIAAwAC4AMAA2ADMAIAAwAC4AMgA1ACAAYwAgADAAIAAtADAALgAwADYAMwAgADAALgAwADEANgAgAC0AMAAuADEAMgA1ACAAMAAuADAAMwAxACAALQAwAC4AMQAyADUAIABjACAAMAAuADAAMQA2ACAAMAAgADAALgAwADMAMQAgADAALgAwADYAMwAgADAALgAwADMAMQAgADAALgAxADIANQAgAGMAIAAwACAALQAwAC4AMAAzADEAIAAwAC4AMAAwADgAIAAtADAALgAwADYAMwAgADAALgAwADEANgAgAC0AMAAuADAANgAzACAAYwAgADAALgAwADAAOAAgADAAIAAwAC4AMAAxADYAIAAwAC4AMAAzADEAIAAwAC4AMAAxADYAIAAwAC4AMAA2ADMAIABjACAAMAAgAC0AMAAuADAAMQA2ACAAMAAuADAAMAA0ACAALQAwAC4AMAAzADEAIAAwAC4AMAAwADgAIAAtADAALgAwADMAMQAgAGMAIAAwAC4AMAAwADQAIAAwACAAMAAuADAAMAA4ACAAMAAuADAAMQA2ACAAMAAuADAAMAA4ACAAMAAuADAAMwAxACAAYwAgADAAIAAtADAALgAwADAAOAAgADAALgAwADAAMgAgAC0AMAAuADAAMQA2ACAAMAAuADAAMAA0ACAALQAwAC4AMAAxADYAIABjACAAMAAuADAAMAAxACAAMAAgADAALgAwADAANAAgADAALgAwADAAOAAgADAALgAwADAANAAgADAALgAwADEANgAgAGMAIAAwACAALQAwAC4AMAAwADQAIAAwAC4AMAAwADEAIAAtADAALgAwADAAOAAgADAALgAwADAAMgAgAC0AMAAuADAAMAA4ACAAYwAgADAAIAAwAC4AMAAwADEAIAAwAC4AMAAwADIAIAAwAC4AMAAwADQAIAAwAC4AMAAwADIAIAAwAC4AMAAwADgAIABjACAAMAAgAC0AMAAuADAAMAAyACAAMAAgAC0AMAAuADAAMAA0ACAAMAAuADAAMAAxACAALQAwAC4AMAAwADQAIABjACAAMAAgADAALgAwADAAMQAgADAALgAwADAAMQAgADAALgAwADAAMgAgADAALgAwADAAMQAgADAALgAwADAANAAgAGMAIAAwACAALQAwAC4AMAAwADEAIAAwACAALQAwAC4AMAAwADIAIAAwACAALQAwAC4AMAAwADMAIABjACAAMAAuADAAMAAxACAAMAAgADAALgAwADAAMQAgADAALgAwADAAMQAgADAALgAwADAAMQAgADAALgAwADAAMgAgAGMAIAAwAC4AMAAwADEAIAAwACAAMAAuADAAMAAxACAALQAwAC4AMAAwADEAIAAwAC4AMAAwADEAIAAtADAALgAwADAAMgAgAGMAIAAwAC4AMAAwADEAIAAwACAAMAAuADAAMAAxACAAMAAuADAAMAAxACAAMAAuADAAMAAxACAAMAAuADAAMAAyACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[55] = [];
    PRESET_SUBTYPES = PRESET_TYPES[55] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1581;KQYAAPr7ACIGAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQNwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTZBQAA+vsA0gUAAAEAAAAJyQUAAPoAAQEBAqICAABNACAAMAAgADAAIABDACAAMAAuADAAMAA0ACAALQAwAC4AMAA2ADcAIAAtADAALgAwADQANgAgAC0AMAAuADEAMgA1ACAALQAwAC4AMQAxADMAIAAtADAALgAxADIAOQAgAEMAIAAtADAALgAxADcANwAgAC0AMAAuADEAMwA0ACAALQAwAC4AMgAzADcAIAAtADAALgAwADgAOQAgAC0AMAAuADIANAAxACAALQAwAC4AMAAyADQAIABDACAALQAwAC4AMgA0ADYAIAAwAC4AMAAzADYAIAAtADAALgAyADAANAAgADAALgAwADkAMgAgAC0AMAAuADEANAA0ACAAMAAuADAAOQA2ACAAQwAgAC0AMAAuADAAOAA5ACAAMAAuADAAOQA5ACAALQAwAC4AMAAzADcAIAAwAC4AMAA2ADIAIAAtADAALgAwADMAMwAgADAALgAwADAANgAgAEMAIAAtADAALgAwADIAOQAgAC0AMAAuADAANAA1ACAALQAwAC4AMAA2ADQAIAAtADAALgAwADkAMwAgAC0AMAAuADEAMQA1ACAALQAwAC4AMAA5ADcAIABDACAALQAwAC4AMQA2ADIAIAAtADAALgAxACAALQAwAC4AMgAwADYAIAAtADAALgAwADYAOQAgAC0AMAAuADIAMAA5ACAALQAwAC4AMAAyADIAIABDACAALQAwAC4AMgAxADIAIAAwAC4AMAAyACAALQAwAC4AMQA4ADQAIAAwAC4AMAA2ADEAIAAtADAALgAxADQAMgAgADAALgAwADYAMwAgAEMAIAAtADAALgAxADAANAAgADAALgAwADYANgAgAC0AMAAuADAANgA4ACAAMAAuADAANAAyACAALQAwAC4AMAA2ADUAIAAwAC4AMAAwADQAIABDACAALQAwAC4AMAA2ADMAIAAtADAALgAwADMAIAAtADAALgAwADgANAAgAC0AMAAuADAANgAzACAALQAwAC4AMQAxADcAIAAtADAALgAwADYANQAgAEMAIAAtADAALgAxADQANgAgAC0AMAAuADAANgA3ACAALQAwAC4AMQA3ADUAIAAtADAALgAwADQAOQAgAC0AMAAuADEANwA3ACAALQAwAC4AMAAyACAAQwAgAC0AMAAuADEANwA5ACAAMAAuADAAMAA1ACAALQAwAC4AMQA2ADQAIAAwAC4AMAAyADkAIAAtADAALgAxADQAIAAwAC4AMAAzADEAIABDACAALQAwAC4AMQAyACAAMAAuADAAMwAzACAALQAwAC4AMAA5ADkAIAAwAC4AMAAyADIAIAAtADAALgAwADkAOAAgADAALgAwADAAMgAgAEMAIAAtADAALgAwADkANgAgAC0AMAAuADAAMQA0ACAALQAwAC4AMQAwADQAIAAtADAALgAwADMAMQAgAC0AMAAuADEAMQA5ACAALQAwAC4AMAAzADMAIABDACAALQAwAC4AMQAzADEAIAAtADAALgAwADMAMwAgAC0AMAAuADEANAAzACAALQAwAC4AMAAyADkAIAAtADAALgAxADQANQAgAC0AMAAuADAAMQA4ACAAQwAgAC0AMAAuADEANAA2ACAALQAwAC4AMAAxADEAIAAtADAALgAxADQANAAgAC0AMAAuADAAMAA0ACAALQAwAC4AMQAzADgAIAAtADAALgAwADAAMQAgAEMAIAAtADAALgAxADMANQAgADAAIAAtADAALgAxADMAMwAgADAAIAAtADAALgAxADMAIAAtADAALgAwADAAMQAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[56] = [];
    PRESET_SUBTYPES = PRESET_TYPES[56] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;273;DQEAAPr7AAYBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQOAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS9AAAA+vsAtgAAAAEAAAAJrQAAAPoAAQEBAhQAAABNACAAMAAgADAAIABMACAAMAAuADIANQAgAC0AMAAuADIANQAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[57] = [];
    PRESET_SUBTYPES = PRESET_TYPES[57] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;365;aQEAAPr7AGIBAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQOQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQZAQAA+vsAEgEAAAEAAAAJCQEAAPoAAQEBAkIAAABNACAAMAAgADAAIABMACAAMAAgAC0AMAAuADEAMgA1ACAAQwAgADAAIAAtADAALgAxADgAMQAgADAALgAwADYAOQAgAC0AMAAuADIANQAgADAALgAxADIANQAgAC0AMAAuADIANQAgAEwAIAAwAC4AMgA1ACAALQAwAC4AMgA1ACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[58] = [];
    PRESET_SUBTYPES = PRESET_TYPES[58] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;441;tQEAAPr7AK4BAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQOgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wRlAQAA+vsAXgEAAAEAAAAJVQEAAPoAAQEBAmgAAABNACAAMAAgADAAIABMACAAMAAuADAANAAgADAALgAwADYANwAgAEMAIAAwAC4AMAA0ADkAIAAwAC4AMAA4ADEAIAAwAC4AMAA1ADQAIAAwAC4AMQAwADIAIAAwAC4AMAA1ADQAIAAwAC4AMQAyADQAIABDACAAMAAuADAANQA0ACAAMAAuADEANAA5ACAAMAAuADAANAA5ACAAMAAuADEANgA5ACAAMAAuADAANAAgADAALgAxADgAMwAgAEwAIAAwACAAMAAuADIANQAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[59] = [];
    PRESET_SUBTYPES = PRESET_TYPES[59] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;533;EQIAAPr7AAoCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQOwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTBAQAA+vsAugEAAAEAAAAJsQEAAPoAAQEBApYAAABNACAAMAAgADAAIABDACAAMAAgAC0AMAAuADAAMwA1ACAAMAAuADAAMgA4ACAALQAwAC4AMAA2ADIAIAAwAC4AMAA2ADIAIAAtADAALgAwADYAMgAgAEMAIAAwAC4AMAA5ADcAIAAtADAALgAwADYAMgAgADAALgAxADIANQAgAC0AMAAuADAAMwA1ACAAMAAuADEAMgA1ACAAMAAgAEMAIAAwAC4AMQAyADUAIAAwAC4AMAAzADUAIAAwAC4AMQA1ADMAIAAwAC4AMAA2ADIAIAAwAC4AMQA4ADgAIAAwAC4AMAA2ADIAIABDACAAMAAuADIAMgAyACAAMAAuADAANgAyACAAMAAuADIANQAgADAALgAwADMANQAgADAALgAyADUAIAAwACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[60] = [];
    PRESET_SUBTYPES = PRESET_TYPES[60] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;791;EwMAAPr7AAwDAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQPAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wTDAgAA+vsAvAIAAAEAAAAJswIAAPoAAQEBAhcBAABNACAAMAAgADAAIABDACAAMAAuADAAMAAyACAAMAAuADAANQAzACAAMAAuADAAMAA3ACAAMAAuADEAMgA3ACAAMAAuADAAMgA1ACAAMAAuADEAMgA2ACAAQwAgADAALgAwADUAMQAgADAALgAxADIANgAgADAALgAwADUAMwAgAC0AMAAuADEAMgAyACAAMAAuADAAOAA0ACAALQAwAC4AMQAyADMAIABDACAAMAAuADEAMQAyACAALQAwAC4AMQAyADMAIAAwAC4AMAA5ADcAIAAwAC4AMAA5ADQAIAAwAC4AMQAyADQAIAAwAC4AMAA5ADMAIABDACAAMAAuADEANQAyACAAMAAuADAAOQAzACAAMAAuADEAMwA3ACAALQAwAC4AMAA2ADQAIAAwAC4AMQA2ADcAIAAtADAALgAwADYANAAgAEMAIAAwAC4AMQA5ADQAIAAtADAALgAwADYANAAgADAALgAxADcAOQAgADAALgAwADQAMgAgADAALgAyADAAMwAgADAALgAwADQAMgAgAEMAIAAwAC4AMgAyADYAIAAwAC4AMAA0ADIAIAAwAC4AMgAxADQAIAAtADAALgAwADMAOQAgADAALgAyADMANQAgAC0AMAAuADAAMwA5ACAAQwAgADAALgAyADQANwAgAC0AMAAuADAAMwA5ACAAMAAuADIANAA4ACAALQAwAC4AMAAxADcAIAAwAC4AMgA0ADkAIAAwACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[61] = [];
    PRESET_SUBTYPES = PRESET_TYPES[61] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;1917;eQcAAPr7AHIHAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQPQAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wQpBwAA+vsAIgcAAAEAAAAJGQcAAPoAAQEBAkoDAABNACAAMAAgADAAIABDACAALQAwAC4AMAAwADgAIAAwAC4AMAAwADgAIAAtADAALgAwADEANwAgADAALgAwADEANgAgAC0AMAAuADAAMgAxACAAMAAuADAAMgA2ACAAQwAgAC0AMAAuADAAMgA1ACAAMAAuADAAMwA3ACAALQAwAC4AMAAyADcAIAAwAC4AMAA1ACAALQAwAC4AMAAyADkAIAAwAC4AMAA2ADMAIABDACAALQAwAC4AMAAzADEAIAAwAC4AMAA3ADYAIAAtADAALgAwADIAOQAgADAALgAwADgANwAgAC0AMAAuADAAMgA3ACAAMAAuADAAOQA5ACAAQwAgAC0AMAAuADAAMgA1ACAAMAAuADEAMQAgAC0AMAAuADAAMgAyACAAMAAuADEAMgAyACAALQAwAC4AMAAxADUAIAAwAC4AMQAzADIAIABDACAALQAwAC4AMAAwADkAIAAwAC4AMQA0ADIAIAAwAC4AMAAwADEAIAAwAC4AMQA1ACAAMAAuADAAMQAyACAAMAAuADEANQA2ACAAQwAgADAALgAwADIAMgAgADAALgAxADYAMgAgADAALgAwADMANAAgADAALgAxADYANgAgADAALgAwADQANgAgADAALgAxADYAOAAgAEMAIAAwAC4AMAA1ADgAIAAwAC4AMQA3ACAAMAAuADAANwAgADAALgAxADcAIAAwAC4AMAA4ADEAIAAwAC4AMQA2ADgAIABDACAAMAAuADAAOQAzACAAMAAuADEANgA2ACAAMAAuADEAMAA0ACAAMAAuADEANgAxACAAMAAuADEAMQAzACAAMAAuADEANQAzACAAQwAgADAALgAxADIAMgAgADAALgAxADQANgAgADAALgAxADMAIAAwAC4AMQAzADcAIAAwAC4AMQAzADQAIAAwAC4AMQAyADYAIABDACAAMAAuADEAMwA5ACAAMAAuADEAMQA2ACAAMAAuADEANAAxACAAMAAuADEAMAAyACAAMAAuADEANAAxACAAMAAuADAAOQAxACAAQwAgADAALgAxADQAMgAgADAALgAwADgAIAAwAC4AMQA0ADEAIAAwAC4AMAA2ADcAIAAwAC4AMQAzADYAIAAwAC4AMAA1ADYAIABDACAAMAAuADEAMwAxACAAMAAuADAANAA2ACAAMAAuADEAMgAyACAAMAAuADAAMwA4ACAAMAAuADEAMQAgADAALgAwADMANAAgAEMAIAAwAC4AMAA5ADgAIAAwAC4AMAAzADEAIAAwAC4AMAA4ADYAIAAwAC4AMAAzADUAIAAwAC4AMAA3ADgAIAAwAC4AMAA0ADIAIABDACAAMAAuADAANwAxACAAMAAuADAANAA5ACAAMAAuADAANgA2ACAAMAAuADAANgAgADAALgAwADYANQAgADAALgAwADcAMwAgAEMAIAAwAC4AMAA2ADUAIAAwAC4AMAA4ADYAIAAwAC4AMAA2ADYAIAAwAC4AMAA5ADgAIAAwAC4AMAA3ADEAIAAwAC4AMQAwADgAIABDACAAMAAuADAANwA2ACAAMAAuADEAMQA4ACAAMAAuADAANwA1ACAAMAAuADEAMgAgADAALgAwADkANQAgADAALgAxADMAMwAgAEMAIAAwAC4AMQAxADMAIAAwAC4AMQA0ADcAIAAwAC4AMQAzADEAIAAwAC4AMQA0ADMAIAAwAC4AMQA0ADIAIAAwAC4AMQA0ADQAIABDACAAMAAuADEANQAzACAAMAAuADEANAA0ACAAMAAuADEANgAyACAAMAAuADEANAAgADAALgAxADcAMwAgADAALgAxADMANgAgAEMAIAAwAC4AMQA4ADUAIAAwAC4AMQAzADEAIAAwAC4AMQA5ADUAIAAwAC4AMQAyADIAIAAwAC4AMgAwADIAIAAwAC4AMQAxADQAIABDACAAMAAuADIAMAA5ACAAMAAuADEAMAA2ACAAMAAuADIAMQAyACAAMAAuADAAOQA2ACAAMAAuADIAMQA2ACAAMAAuADAAOAAgAEMAIAAwAC4AMgAxADkAIAAwAC4AMAA2ADQAIAAwAC4AMgAxADkAIAAwAC4AMAA1ADYAIAAwAC4AMgAxADkAIAAwAC4AMAA0ADQAIABDACAAMAAuADIAMQA5ACAAMAAuADAAMwAyACAAMAAuADIAMQA5ACAAMAAuADAAMgAgADAALgAyADEAOQAgADAALgAwADAAOAAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";
    PRESET_TYPES[62] = [];
    PRESET_SUBTYPES = PRESET_TYPES[62] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;527;CwIAAPr7AAQCAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQPgAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS7AQAA+vsAtAEAAAEAAAAJqwEAAPoAAQEBApMAAABNACAAMAAgADAAIABsACAAMAAuADAAMwA2ACAAMAAgAGwAIAAwACAAMAAuADAAMwA2ACAAbAAgADAALgAwADMANgAgADAAIABsACAAMAAgADAALgAwADMANgAgAGwAIAAwAC4AMAAzADYAIAAwACAAbAAgADAAIAAwAC4AMAAzADYAIABsACAAMAAuADAAMwA2ACAAMAAgAGwAIAAwACAAMAAuADAAMwA2ACAAbAAgADAALgAwADMANgAgADAAIABsACAAMAAgADAALgAwADMANgAgAGwAIAAwAC4AMAAzADYAIAAwACAAbAAgADAAIAAwAC4AMAAzADYAIABsACAAMAAuADAAMwA2ACAAMAAgAGwAIAAwACAAMAAuADAAMwA2ACAARQADAAAAAPsAcAAAAPr7ABYAAAD6AwEPBgAAABMEAAAAMgAwADAAMAD7ARIAAAD6+wALAAAA+gABAAAANAACAPsCNwAAAPr7ADAAAAACAAAAABEAAAD6AAUAAABwAHAAdABfAHgA+wARAAAA+gAFAAAAcABwAHQAXwB5APs=";
    PRESET_TYPES[63] = [];
    PRESET_SUBTYPES = PRESET_TYPES[63] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;265;BQEAAPr7AP4AAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQPwAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS1AAAA+vsArgAAAAEAAAAJpQAAAPoAAQEBAhAAAABNACAAMAAgADAAIABMACAAMAAuADIANQAgADAAIABFAAMAAAAA+wBwAAAA+vsAFgAAAPoDAQ8GAAAAEwQAAAAyADAAMAAwAPsBEgAAAPr7AAsAAAD6AAEAAAA0AAIA+wI3AAAA+vsAMAAAAAIAAAAAEQAAAPoABQAAAHAAcAB0AF8AeAD7ABEAAAD6AAUAAABwAHAAdABfAHkA+w==";
    PRESET_TYPES[64] = [];
    PRESET_SUBTYPES = PRESET_TYPES[64] = [];
    PRESET_SUBTYPES[0] = "PPTY;v10;267;BwEAAPr7AAABAAD6AFDDAAADAQUCBgQMUMMAAA4AAAAADwUAAAAQQAAAABEAAAAA+wAZAAAA+vsAEgAAAAEAAAAACQAAAPoDAQAAADAA+wS3AAAA+vsAsAAAAAEAAAAJpwAAAPoAAQEBAhEAAABNACAAMAAgADAAIABMACAAMAAgAC0AMAAuADIANQAgAEUAAwAAAAD7AHAAAAD6+wAWAAAA+gMBDwYAAAATBAAAADIAMAAwADAA+wESAAAA+vsACwAAAPoAAQAAADQAAgD7AjcAAAD6+wAwAAAAAgAAAAARAAAA+gAFAAAAcABwAHQAXwB4APsAEQAAAPoABQAAAHAAcAB0AF8AeQD7";


})(window);
