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

    function CBaseAnimObject() {
        CBaseFormatObject.call(this);
    }
    InitClass(CBaseAnimObject, CBaseFormatObject, AscDFH.historyitem_type_Unknown);
    CBaseAnimObject.prototype.isTimeNode = function() {
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

    var TIME_NODE_STATE_IDLE = 0;
    var TIME_NODE_STATE_ACTIVE = 1;
    var TIME_NODE_STATE_FROZEN = 2;
    var TIME_NODE_STATE_FINISHED = 3;

    var oSTATEDESCRMAP = {};

    oSTATEDESCRMAP[TIME_NODE_STATE_IDLE] = 'IDLE';
    oSTATEDESCRMAP[TIME_NODE_STATE_ACTIVE] = 'ACTIVE';
    oSTATEDESCRMAP[TIME_NODE_STATE_FROZEN] = 'FROZEN';
    oSTATEDESCRMAP[TIME_NODE_STATE_FINISHED] = 'FINISHED';



    var NODE_TYPE_AFTEREFFECT	 = 0;
    var NODE_TYPE_AFTERGROUP	 = 1;
    var NODE_TYPE_CLICKEFFECT	 = 2;
    var NODE_TYPE_CLICKPAR		 = 3;
    var NODE_TYPE_INTERACTIVESEQ = 4;
    var NODE_TYPE_MAINSEQ		 = 5;
    var NODE_TYPE_TMROOT		 = 6;
    var NODE_TYPE_WITHEFFECT	 = 7;
    var NODE_TYPE_WITHGROUP		 = 8;

    var NODE_TYPE_MAP = {};
    NODE_TYPE_MAP[NODE_TYPE_AFTEREFFECT	 ] = "AFTEREFFECT";
    NODE_TYPE_MAP[NODE_TYPE_AFTERGROUP	 ] = "AFTERGROUP";
    NODE_TYPE_MAP[NODE_TYPE_CLICKEFFECT	 ] = "CLICKEFFECT";
    NODE_TYPE_MAP[NODE_TYPE_CLICKPAR		 ] = "CLICKPAR";
    NODE_TYPE_MAP[NODE_TYPE_INTERACTIVESEQ ] = "INTERACTIVESEQ";
    NODE_TYPE_MAP[NODE_TYPE_MAINSEQ		 ] = "MAINSEQ";
    NODE_TYPE_MAP[NODE_TYPE_TMROOT		 ] = "TMROOT";
    NODE_TYPE_MAP[NODE_TYPE_WITHEFFECT	 ] = "WITHEFFECT";
    NODE_TYPE_MAP[NODE_TYPE_WITHGROUP		 ] = "WITHGROUP";


    var NODE_FILL_FREEZE = 0;
    var NODE_FILL_HOLD = 1;
    var NODE_FILL_REMOVE = 2;
    var NODE_FILL_TRANSITION = 3;

    function CTimeNodeBase() {
        CBaseAnimObject.call(this);
        this.state = TIME_NODE_STATE_IDLE;

        this.simpleDurationIdx = -1;
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
    CTimeNodeBase.prototype.resetState = function() {
        this.state = TIME_NODE_STATE_IDLE;
        this.simpleDurationIdx = -1;
        if(this.isTimingContainer()) {
            var aChildren = this.getChildrenTimeNodes();
            for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                aChildren[nChild].resetState();
            }
        }
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
    CTimeNodeBase.prototype.getStartTrigger = function(oPlayer) {
        var oAttributes = this.getAttributesObject();
        if(!oAttributes || !oAttributes.stCondLst) {
            return this.getDefaultTrigger(oPlayer);
        }
        var oTrigger;
        var nNodeType = this.getNodeType();
        var oPreviousTimeNode;
        switch (nNodeType) {
            case NODE_TYPE_MAINSEQ: {
                oTrigger = this.getDefaultTrigger(oPlayer);
                break;
            }
            case NODE_TYPE_CLICKEFFECT: {
                oTrigger = this.createEffectTrigger(this.createExternalEventTrigger(oPlayer, oTrigger, COND_EVNT_ON_CLICK, null), oPlayer);
                break;
            }
            case NODE_TYPE_WITHEFFECT: {
                oPreviousTimeNode = this;
                while ((oPreviousTimeNode = oPreviousTimeNode.getPreviousNode()) &&
                (oPreviousTimeNode.getNodeType() === NODE_TYPE_WITHEFFECT)) {
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
            case NODE_TYPE_AFTEREFFECT: {
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
                    oTrigger.addTrigger(function() {
                        for(var nChild = 0; nChild < aChildren.length; ++nChild) {
                            if(!aChildren[nChild].isAtEnd()) {
                                return false;
                            }
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
            }
        }
        return function () {
            oThis.finishCallback(oPlayer);
        };
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
        console.log(sPrefix + " | ID: " + this.Id + " | TYPE: " + this.constructor.name + " | NODE_TYPE: " + sNodeType + " | STATE: " + oSTATEDESCRMAP[this.state] + " | TIME: " + (new Date()).getTime() + " | FORMAT ID: " + oAttr.id);
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
                if(this.repeatCount.isSpecified() && this.simpleDurationIdx + 1 < this.repeatCount.getVal() / 1000) {
                    this.startSimpleDuration(++this.simpleDurationIdx, oPlayer);
                }
            }
        }
        else if(this.isSeq()) {
            var nChildIdx = this.getChildNodeIdx(oChild);
            if(nChildIdx < aChildren.length - 1) {
                aChildren[nChildIdx + 1].scheduleStart(oPlayer);
                // //handle advance after
                // if(this.getNodeType() === NODE_TYPE_MAINSEQ) {
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
                if(this.repeatCount.isSpecified() && this.simpleDurationIdx + 1 < this.repeatCount.getVal() / 1000) {
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
        if(oAttributes && oAttributes.nodeType === NODE_TYPE_MAINSEQ) {
            return true;
        }
        return false;
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
            if(bAutoRev) {
                fRelTime = 0.0;
            }
            else {
                fRelTime = 1.0;
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
    CTimeNodeBase.prototype.getPresentation = function() {
        return editor.WordControl.m_oLogicDocument;
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
    CTiming.prototype.getTimingRootNode = function() {
        if(this.tnLst) {
            if(this.tnLst) {
                if(this.tnLst.list.length > 0) {
                    return this.tnLst.list[0];
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
    CCondLst.prototype.createComplexTrigger = function(oPlayer) {
        var oComplexTrigger = new CAnimComplexTrigger();
        for(var nCond = 0; nCond < this.list.length; ++nCond) {
            this.list[nCond].fillTrigger(oPlayer, oComplexTrigger)
        }
        return oComplexTrigger;
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
    drawingsChangesMap[AscDFH.historyitem_IndexRgEnd] = function(oClass, value) {oClass.st = value;};
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
            "ppt_h": oAttributes && AscFormat.isRealNumber(oAttributes["ppt_h"]) ? oAttributes["ppt_h"] : this.getOrigAttrVal("ppt_h")
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

    var PRESTET_CLASS_EMPH = 0;
    var PRESTET_CLASS_ENTR = 1;
    var PRESTET_CLASS_EXIT = 2;
    var PRESTET_CLASS_MEDIACALL = 3;
    var PRESTET_CLASS_PATH = 4;
    var PRESTET_CLASS_VERB = 5;

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
    CAnimClr.prototype.isAllowedAttribute = function(sAttrName) {
        return sAttrName === "style.color" || sAttrName === "fillcolor" || sAttrName === "stroke.color";
    };
    CAnimClr.prototype.calculateAttributes = function(nElapsedTime, oAttributes) {
        var oTargetObject = this.getTargetObject();
        if(!oTargetObject) {
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
                oBrush = oPen && oPen.Fill;
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
    drawingsChangesMap[AscDFH.historyitem_AnimMotionCBhvr] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionFrom] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionRCtr] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionTo] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionOrigin] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPath] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPathEditMode] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionPtsTypes] = function(oClass, value) {oClass.by = value;};
    drawingsChangesMap[AscDFH.historyitem_AnimMotionRAng] = function(oClass, value) {oClass.by = value;};


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

    changesFactory[AscDFH.historyitem_TimeNodeContainerCTn] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TimeNodeContainerCTn] = function(oClass, value) {oClass.cTn = value;};

    function CTimeNodeContainer() {//par, excl
        CTimeNodeBase.call(this);
        this.cTn = null;
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

    function CPar() {//par, seq, excl
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
                                oChild.freezeCallback(oPlayer);
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
                if(PRESTET_CLASS_ENTR === oAttrObject.presetClass) {
                    return true;
                }
                oCurNode = oParentNode;
            }
        }
        return false;
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
            if(oTransform.IsIdentity2()) {
                var fDelta = 2;
                if(AscFormat.fApproxEqual(oTransform.tx*this.scale, this.x, fDelta) &&
                    AscFormat.fApproxEqual(oTransform.ty*this.scale, this.y, fDelta)) {
                    bNoTransform = true;
                }
            }
        }
        if(bNoTransform) {
            oGraphics.SaveGrState();
            oGraphics.SetIntegerGrid(true);
            var nDx = oGraphics.m_oCoordTransform.tx;
            var nDy = oGraphics.m_oCoordTransform.ty;
            oGraphics.m_oContext.drawImage(this.canvas, nDx + this.x, nDy + this.y, this.canvas.width, this.canvas.height);
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
        return this.createDiamond(1 - fTime, "destination-in");
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
        var aFilledRanges = this.getRandomRanges(fTime);
        if(aFilledRanges.length === 0) {
            return this;
        }
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
        if(editor.WordControl.DemonstrationManager.Mode) {
            oGraphics.IsDemonstrationMode = true;
        }
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
    CAnimationDrawer.prototype.checkHiddenObject = function(oTimeNode) {
        if(oTimeNode.doesHideObject()) {
            var sId = oTimeNode.getTargetObjectId();
            if(sId !== null) {
                this.hiddenObjects[oTimeNode.getTargetObjectId()] = oTimeNode;
            }
        }
    };
    CAnimationDrawer.prototype.collectHiddenObjects = function() {
        var aTimings = this.player.timings;
        var oThis = this;
        for(var nTiming = 0; nTiming < aTimings.length; ++nTiming) {
            var oRoot = aTimings[nTiming].getTimingRootNode();
            if(oRoot) {
                oRoot.traverseTimeNodes(function(oTimeNode) {
                    oThis.checkHiddenObject(oTimeNode);
                });
            }
        }
    };

    function CAnimationPlayer(oSlide, drawer) {
        this.slide = oSlide;
        this.timings = [];
        if(oSlide.timing) {
            this.timings.push(oSlide.timing);
        }
        if(oSlide.Layout.timing) {
            this.timings.push(oSlide.Layout.timing);
        }
        if(oSlide.Layout.Master.timing) {
            this.timings.push(oSlide.Layout.Master.timing);
        }
        for(var nTiming = 0; nTiming < this.timings.length; ++nTiming) {
            var oRoot = this.timings[nTiming].getTimingRootNode();
            if(oRoot) {
                oRoot.resetState();
            }
        }
        this.eventsProcessor = new CEventsProcessor(this);
        this.animationScheduler = new CAnimationScheduler(this);
        this.animationDrawer = new CAnimationDrawer(this);
        this.timer = new CAnimationTimer(this);
        this.drawer = drawer;
    }
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
            this.resetNodesState();
            this.scheduleNodesStart();
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
        var oTiming = this.slide && this.slide.timing;
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
        return this.addExternalEvent(new CExternalEvent(this.eventsProcessor, COND_EVNT_ON_NEXT, null));
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
        //TODO: sort
    };
    CAnimSandwich.prototype.getDrawing = function() {
        return AscCommon.g_oTableId.Get_ById(this.drawingId);
    };
    CAnimSandwich.prototype.getAttributesMap = function() {
        if(this.cachedAttributes) {
            return this.cachedAttributes;
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
                    if(oDrawing.pen) {
                        var oPen = oDrawing.pen.createDuplicate();
                        var oMods;
                        if(oPen.Fill &&
                            oPen.Fill.fill &&
                            oPen.Fill.fill.color &&
                            oPen.Fill.fill.color.Mods &&
                            oPen.Fill.fill.color.Mods.Mods.length !== 0) {
                            oMods = oPen.Fill.fill.color.Mods;
                            oMods.Apply(oStrokeColor.RGBA);
                        }
                        oPen.Fill = AscFormat.CreateUniFillByUniColor(oStrokeColor);
                        oDrawing.pen = oPen;
                    }
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
        oLastToken.calculate(oVarMap);
        return oLastToken.result;
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

    var NUMBER_REGEXPSTR = "[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?";
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
                    if(oLastFunction.addOperand(this.queue.last()) >= oLastFunction.getArgumentsCount()){
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


    var GLOBAL_PLAYER = null;


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
})(window);
