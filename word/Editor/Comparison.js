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

(function (undefined) {

    var MAX_COMPARES = 1200000;
    var MIN_JACCARD_RUDE = 0.5;
    var MIN_JACCARD = 0.34;
    var MIN_DIFF = 0.7;
    var EXCLUDED_PUNCTUATION = {};
    EXCLUDED_PUNCTUATION[46] = true;
    //EXCLUDED_PUNCTUATION[95] = true;
    EXCLUDED_PUNCTUATION[160] = true;
    function CNode(oElement, oParent)
    {
        this.element = oElement;
        this.par = null;
        this.children = [];
        this.depth   = 0;
        this.changes = [];
        this.partner = null;
        this.childidx = null;

        this.hashWords = null;
        if(oParent)
        {
            oParent.addChildNode(this);
        }
    }
    CNode.prototype.getElement = function()
    {
        return this.element;
    };

    CNode.prototype.getNeighbors  = function()
    {
        if(!this.par)
        {
            return [undefined, undefined];
        }
        return [this.par.children[this.childidx - 1],  this.par.children[this.childidx + 1]];
    };

    CNode.prototype.print = function()
    {
        if(this.element.print)
        {
            this.element.print();
        }
    };
    CNode.prototype.getDepth = function()
    {
        return this.depth;
    };
    CNode.prototype.addChange = function(oOperation)
    {
        return this.changes.push(oOperation);
    };
    CNode.prototype.equals = function(oNode)
    {
        if(this.depth === oNode.depth)
        {
            var oParent1, oParent2;
            oParent1 = this.par;
            oParent2 = oNode.par;
            if(oParent1 && !oParent2 || !oParent1 && oParent2)
            {
                return false;
            }
            if(oParent1)
            {
                if(!oParent1.equals(oParent2))
                {
                    return false;
                }
            }
            return this.privateCompareElements(oNode, true);
        }
        return false;
    };

    CNode.prototype.privateCompareElements = function(oNode, bCheckNeighbors)
    {
        var oElement1 = this.element;
        var oElement2 = oNode.element;
        if(oElement1.constructor === oElement2.constructor)
        {
            if(typeof oElement1.Value === "number")
            {
                return oElement1.Value === oElement2.Value;
            }
            if(oElement1 instanceof CTextElement)
            {
                if(bCheckNeighbors && oElement1.isSpaceText() && oElement2.isSpaceText())
                {
                    var aNeighbors1 = this.getNeighbors();
                    var aNeighbors2 = oNode.getNeighbors();
                    if(!aNeighbors1[0] && !aNeighbors2[0] || !aNeighbors1[1] && !aNeighbors2[1])
                    {
                        return true;
                    }
                    if(aNeighbors1[0] && aNeighbors2[0])
                    {
                        if(aNeighbors1[0].privateCompareElements(aNeighbors2[0], false))
                        {
                            return true;
                        }
                    }
                    if(aNeighbors1[1] && aNeighbors2[1])
                    {
                        if(aNeighbors1[1].privateCompareElements(aNeighbors2[1], false))
                        {
                            return true;
                        }
                    }
                    return false;
                }
                else
                {
                    return oElement1.equals(oElement2);
                }
            }
            if(oElement1 instanceof AscCommonWord.CTable)
            {
                if(oElement1.TableGrid.length !== oElement2.TableGrid.length)
                {
                    return false;
                }
            }
            if(oElement1 instanceof AscCommonWord.CTableRow)
            {
                if(oElement1.Content.length !== oElement2.Content.length)
                {
                    return false;
                }
            }
            if(oElement1 instanceof AscCommonWord.CDocumentContent && oElement1.Parent instanceof AscCommonWord.CTableCell)
            {
                if(!oElement2.Parent)
                {
                    return false;
                }
                if(oElement1.Parent.Index !== oElement2.Parent.Index)
                {
                    return false;
                }
            }
            if(oElement1 instanceof AscCommonWord.ParaMath)
            {
                return false;
            }
            return true;
        }
        return false;
    };
    CNode.prototype.isLeaf = function()
    {
        return this.children.length === 0;
    };

    CNode.prototype.addChildNode = function(oNode)
    {
        oNode.childidx = this.children.length;
        this.children.push(oNode);
        oNode.depth = this.depth + 1;
        oNode.par = this;
    };

    CNode.prototype.isStructure = function()
    {
        return !this.isLeaf();
    };


    CNode.prototype.forEachDescendant = function(callback, T) {
        this.children.forEach(function(node) {
            node.forEach(callback, T);
        });
    };
    CNode.prototype.forEach = function(callback, T) {
        callback.call(T, this);
        this.children.forEach(function(node) {
            node.forEach(callback, T);
        });
    };

    CNode.prototype.setPartner = function (oNode) {
        this.partner = oNode;
        oNode.partner = this;
        return null;
    };

    CNode.prototype.isComparable = function (oNode) {
        if(this.element && oNode.element && this.element.constructor === oNode.element.constructor)
        {
            if(this.element instanceof CTable)
            {
                if(this.element.TableGrid.length !== oNode.element.TableGrid.length)
                {
                    return false;
                }
            }
            if(this.element instanceof AscCommonWord.Paragraph)
            {
                if(!this.element.SectPr && oNode.element.SectPr)
                {
                    return false;
                }
                if(!oNode.element.SectPr && this.element.SectPr)
                {
                    return false;
                }
            }
            return true;
        }
        return false;
    };

    function CTextElement()
    {
        this.elements = [];
        this.firstRun = null;
        this.lastRun = null;
    }

    CTextElement.prototype.equals = function (other)
    {
        if(this.elements.length !== other.elements.length)
        {
            return false;
        }
        var oElement, oOtherElement;
        for(var i = 0; i < this.elements.length; ++i)
        {
            oElement = this.elements[i];
            oOtherElement = other.elements[i];
            if(oElement.constructor !== oOtherElement.constructor)
            {
                return false;
            }
            if(typeof oElement.Value === "number")
            {
                if(oElement.Value !== oOtherElement.Value)
                {
                    return false;
                }
            }
            if(oElement instanceof ParaDrawing)
            {
                return oElement.IsComparable(oOtherElement)
            }
        }
        return true;
    };

    CTextElement.prototype.updateHash = function(oHash){
        var aCheckArray = [];
        var bVal = false;
        for(var i = 0; i < this.elements.length; ++i)
        {
            var oElement = this.elements[i];
            if(AscFormat.isRealNumber(oElement.Value))
            {
                aCheckArray.push(oElement.Value);
                bVal = true;
            }
            else
            {
                if(oElement instanceof ParaNewLine)
                {
                    if(oElement.BreakType === AscCommonWord.break_Line)
                    {
                        aCheckArray.push(0x000A);
                    }
                    else
                    {
                        aCheckArray.push(0x21A1);
                    }
                }
                else if(oElement instanceof ParaTab)
                {
                    aCheckArray.push(0x0009);
                }
                else if(oElement instanceof ParaSpace)
                {
                    aCheckArray.push(0x20);
                }
            }
        }
        if(aCheckArray.length > 0)
        {
            oHash.update(aCheckArray);
            if(bVal)
            {
                oHash.countLetters++;
            }
        }
    };

    CTextElement.prototype.print = function ()
    {
        var sResultString = "";
        for(var i = 0; i < this.elements.length; ++i)
        {
            if(this.elements[i] instanceof ParaText)
            {
                sResultString += String.fromCharCode(this.elements[i].Value);
            }
            else if(this.elements[i] instanceof ParaSpace)
            {
                sResultString += " ";
            }
        }
        //console.log(sResultString);
    };
    CTextElement.prototype.setFirstRun = function (oRun)
    {
        this.firstRun = oRun;
    };
    CTextElement.prototype.setLastRun = function (oRun)
    {
        this.lastRun = oRun;
    };

    CTextElement.prototype.isSpaceText = function ()
    {
        if(this.elements.length === 1)
        {
            return (this.elements[0].Type === para_Space);
        }
        return false;
    };

    CTextElement.prototype.isParaEnd = function ()
    {
        if(this.elements.length === 1)
        {
            return (this.elements[0].Type === para_End);
        }
        return false;
    };

    CTextElement.prototype.compareFootnotes = function (oTextElement)
    {
        if(this.elements.length === 1 && oTextElement.elements.length === 1
        && (this.elements[0].Type === para_FootnoteReference && oTextElement.elements[0].Type === para_FootnoteReference
		 || this.elements[0].Type === para_EndnoteReference && oTextElement.elements[0].Type === para_EndnoteReference))
        {
            var oBaseContent = this.elements[0].Footnote;
            var oCompareContent = oTextElement.elements[0].Footnote;
            if(oBaseContent && oCompareContent)
            {
                if(!AscCommon.g_oTableId.Get_ById(oBaseContent.Id))
                {
                    var t = oBaseContent;
                    oBaseContent = oCompareContent;
                    oCompareContent = t;
                }
                return [oBaseContent, oCompareContent];
            }
        }
        return null;
    };

    CTextElement.prototype.compareDrawings = function(oTextElement)
    {
        if(this.elements.length === 1 && oTextElement.elements.length === 1)
        {
            var oElement = this.elements[0];
            var oOtherElement = oTextElement.elements[0];
            if(oElement.Type === para_Drawing && oOtherElement.Type === para_Drawing)
            {
                if(oElement.IsComparable(oOtherElement))
                {
                    if(AscCommon.g_oTableId.Get_ById(oElement.Id))
                    {
                        return [oElement, oOtherElement];
                    }
                    else
                    {
                        return [oOtherElement, oElement];
                    }
                }
            }
        }
        return null;
    };

    function CMatching()
    {
        this.Footnotes = {};
        this.Drawings = {};
    }
    CMatching.prototype.get = function(oNode)
    {
        return oNode.partner;
    };

    CMatching.prototype.put = function(oNode1, oNode2)
    {
        oNode1.setPartner(oNode2);

        var aFootnotes, aDrawings;
        if(oNode1.element instanceof CTextElement)
        {
            aFootnotes = oNode1.element.compareFootnotes(oNode2.element);
            if(aFootnotes)
            {
                this.Footnotes[aFootnotes[0].Id] = aFootnotes[1];
            }
            else
            {
                aDrawings = oNode1.element.compareDrawings(oNode2.element);
                if(aDrawings)
                {
                    this.Drawings[aDrawings[0].Id] = aDrawings[1];
                }
            }
        }
    };

    function ComparisonOptions()
    {
        this.insertionsAndDeletions = null;
        this.moves = null;
        this.comments = null;
        this.formatting = null;
        this.caseChanges = null;
        this.whiteSpace = null;
        this.tables = null;
        this.headersAndFooters = null;
        this.footNotes  = null;
        this.textBoxes = null;
        this.fields = null;
        this.words = null;
    }
    ComparisonOptions.prototype["getInsertionsAndDeletions"] = ComparisonOptions.prototype.getInsertionsAndDeletions = function(){return this.insertionsAndDeletions !== false;};
    ComparisonOptions.prototype["getMoves"] = ComparisonOptions.prototype.getMoves = function(){return this.moves !== false;};
    ComparisonOptions.prototype["getComments"] = ComparisonOptions.prototype.getComments = function(){return this.comments !== false;};
    ComparisonOptions.prototype["getFormatting"] = ComparisonOptions.prototype.getFormatting = function(){return this.formatting !== false;};
    ComparisonOptions.prototype["getCaseChanges"] = ComparisonOptions.prototype.getCaseChanges = function(){return this.caseChanges !== false;};
    ComparisonOptions.prototype["getWhiteSpace"] = ComparisonOptions.prototype.getWhiteSpace = function(){return this.whiteSpace !== false;};
    ComparisonOptions.prototype["getTables"] = ComparisonOptions.prototype.getTables = function(){return true;/*this.tables !== false;*/};
    ComparisonOptions.prototype["getHeadersAndFooters"] = ComparisonOptions.prototype.getHeadersAndFooters = function(){return this.headersAndFooters !== false;};
    ComparisonOptions.prototype["getFootNotes"] = ComparisonOptions.prototype.getFootNotes = function(){return this.footNotes !== false;};
    ComparisonOptions.prototype["getTextBoxes"] = ComparisonOptions.prototype.getTextBoxes = function(){return this.textBoxes !== false;};
    ComparisonOptions.prototype["getFields"] = ComparisonOptions.prototype.getFields = function(){return this.fields !== false;};
    ComparisonOptions.prototype["getWords"] = ComparisonOptions.prototype.getWords = function(){return true;/* this.words !== false;*/};


    ComparisonOptions.prototype["putInsertionsAndDeletions"] = ComparisonOptions.prototype.putInsertionsAndDeletions = function(v){this.insertionsAndDeletions = v;};
    ComparisonOptions.prototype["putMoves"] = ComparisonOptions.prototype.putMoves = function(v){this.moves = v;};
    ComparisonOptions.prototype["putComments"] = ComparisonOptions.prototype.putComments = function(v){this.comments = v;};
    ComparisonOptions.prototype["putFormatting"] = ComparisonOptions.prototype.putFormatting = function(v){this.formatting = v;};
    ComparisonOptions.prototype["putCaseChanges"] = ComparisonOptions.prototype.putCaseChanges = function(v){this.caseChanges = v;};
    ComparisonOptions.prototype["putWhiteSpace"] = ComparisonOptions.prototype.putWhiteSpace = function(v){this.whiteSpace = v;};
    ComparisonOptions.prototype["putTables"] = ComparisonOptions.prototype.putTables = function(v){this.tables = v;};
    ComparisonOptions.prototype["putHeadersAndFooters"] = ComparisonOptions.prototype.putHeadersAndFooters = function(v){this.headersAndFooters = v;};
    ComparisonOptions.prototype["putFootNotes"] = ComparisonOptions.prototype.putFootNotes = function(v){this.footNotes = v;};
    ComparisonOptions.prototype["putTextBoxes"] = ComparisonOptions.prototype.putTextBoxes = function(v){this.textBoxes = v;};
    ComparisonOptions.prototype["putFields"] = ComparisonOptions.prototype.putFields = function(v){this.fields = v;};
    ComparisonOptions.prototype["putWords"] = ComparisonOptions.prototype.putWords = function(v){this.words = v;};


    function CDocumentComparison(oOriginalDocument, oRevisedDocument, oOptions)
    {
        this.originalDocument = oOriginalDocument;
        this.revisedDocument = oRevisedDocument;
        this.options = oOptions;
        this.api = oOriginalDocument.GetApi();
        this.StylesMap = {};
        this.CommentsMap = {};
        this.matchedNums = {};
        this.checkedNums = {};

        this.copyPr = {
            CopyReviewPr: false,
            Comparison: this
        };
    }
    CDocumentComparison.prototype.getUserName = function()
    {
        var oCore = this.revisedDocument.Core;
        if(oCore && typeof oCore.lastModifiedBy === "string" && oCore.lastModifiedBy.length > 0)
        {
            return  oCore.lastModifiedBy.split(";")[0];
        }
        else
        {
            return AscCommon.translateManager.getValue("Author");
        }
    };
    CDocumentComparison.prototype.compareElementsArray = function(aBase, aCompare, bOrig, bUseMinDiff)
    {
        var oMapEquals = {};
        var aBase2 = [];
        var aCompare2 = [];
        var oCompareMap = {};
        var bMatchNoEmpty = false, i, j, key;
        var oLCS;
        var oThis = this;
        var fLCSCallback = function(x, y) {
            var oOrigNode = oLCS.a[x];
            var oReviseNode = oLCS.b[y];
            var oDiff  = new AscCommon.Diff(oOrigNode, oReviseNode);
            oDiff.equals = function(a, b)
            {
                return a.equals(b);
            };
            var oMatching = new CMatching();
            oDiff.matchTrees(oMatching);
            var oDeltaCollector = new AscCommon.DeltaCollector(oMatching, oOrigNode, oReviseNode);
            oDeltaCollector.forEachChange(function(oOperation){
                oOperation.anchor.base.addChange(oOperation);
            });
            oThis.applyChangesToChildNode(oOrigNode);
            oThis.compareNotes(oMatching);
            oThis.compareDrawingObjects(oMatching);
        };
        var fEquals;

        fEquals = function(a, b)
        {

            var bEquals = oMapEquals[a.element.Id] || oMapEquals[b.element.Id];
            if(oEqualMap[a.element.Id])
            {
                if(bEquals && !AscFormat.fApproxEqual(oEqualMap[a.element.Id].jaccard, 1.0, 0.01))
                {
                    return false;
                }
                if(oEqualMap[a.element.Id].map[b.element.Id])
                {
                    return true;
                }
            }
            else
            {
                if(bEquals && !AscFormat.fApproxEqual(oEqualMap[b.element.Id].jaccard, 1.0, 0.01))
                {
                    return false;
                }
                if(oEqualMap[b.element.Id])
                {
                    if(oEqualMap[b.element.Id].map[a.element.Id])
                    {
                        return true;
                    }
                }
            }
            return false;
        };

        var oEqualMap = {};
        for(i = 0; i < aBase.length; ++i)
        {
            var oCurNode =  aBase[i];
            if(oCurNode.hashWords)
            {
                var oCurInfo = {

                    jaccard: 0,
                    map: {},
                    minDiff: 0,
                    intersection: 0
                };
                oEqualMap[oCurNode.element.Id] = oCurInfo;
                for(j = 0; j < aCompare.length; ++j)
                {
                    var oCompareNode = aCompare[j];
                    if(oCompareNode.hashWords && oCurNode.isComparable(oCompareNode))
                    {
                        var dJaccard = oCurNode.hashWords.jaccard(oCompareNode.hashWords);
                        if(oCurNode.element instanceof CTable)
                        {
                            dJaccard += MIN_JACCARD;
                        }
                        var dIntersection = dJaccard*(oCurNode.hashWords.count + oCompareNode.hashWords.count)/(1+dJaccard);
                        var diffA = 0, diffB = 0, dMinDiff = 0;
                        if(dJaccard > 0)
                        {
                            if(oCurNode.hashWords.count > 0)
                            {
                                diffA = dIntersection/oCurNode.hashWords.count;
                            }
                            if(oCompareNode.hashWords.count > 0)
                            {
                                diffB = dIntersection/oCompareNode.hashWords.count;
                            }
                            dMinDiff = Math.max(diffA, diffB);

                            if(oCurInfo.jaccard <= dJaccard && dJaccard >= MIN_JACCARD || (oCurInfo.jaccard < MIN_JACCARD && dMinDiff > MIN_DIFF && oCurInfo.minDiff <= dMinDiff))
                            {
                                if(oCurInfo.jaccard < dJaccard && dJaccard >= MIN_JACCARD)
                                {
                                    oCurInfo.map = {};
                                    oCurInfo.minDiff = 0;
                                }
                                oCurInfo.map[oCompareNode.element.Id] = oCompareNode;
                                oCurInfo.jaccard = dJaccard;
                                oCurInfo.intersection = dIntersection;
                                oCurInfo.minDiff = dMinDiff;
                                if(AscFormat.fApproxEqual(dJaccard, 1.0, 0.01))
                                {
                                    oMapEquals[oCompareNode.element.Id] = true;
                                }
                            }
                        }

                    }
                }
                if(oCurInfo.jaccard >= MIN_JACCARD || (bUseMinDiff && oCurInfo.minDiff > MIN_DIFF && oCurNode.hashWords.countLetters > 0 ))
                {
                    aBase2.push(oCurNode);
                    for(key in oCurInfo.map)
                    {
                        if(oCurInfo.map.hasOwnProperty(key))
                        {
                            oCompareMap[key] = true;
                            if(oCurNode.hashWords.countLetters > 0 && oCurInfo.map[key].hashWords.countLetters > 0)
                            {
                                bMatchNoEmpty = true;
                            }
                        }
                    }
                }
            }
        }
        for(j = 0; j < aCompare.length; ++j)
        {
            oCompareNode = aCompare[j];
            if(oCompareMap[oCompareNode.element.Id])
            {
                aCompare2.push(oCompareNode);
            }
        }
        if(!bMatchNoEmpty)
        {
            if(bOrig)
            {
                for(i = 0; i < aBase2.length; ++i)
                {
                    if(i !== aBase2[i].childidx)
                    {
                        aBase2.splice(i, aBase2[i].length - i);
                        break;
                    }
                }
                for(i = aCompare2.length - 1; i > -1; i--)
                {
                    if(i !== aCompare2[i].childidx)
                    {
                        aCompare2.splice(0, i + 1);
                        break;
                    }
                }
            }
            else
            {

                for(i = 0; i < aCompare2.length; ++i)
                {
                    if(i !== aCompare2[i].childidx)
                    {
                        aCompare2.splice(i, aCompare2[i].length - i);
                        break;
                    }
                }
                for(i = aBase2.length - 1; i > -1; i--)
                {
                    if(i !== aBase2[i].childidx)
                    {
                        aBase2.splice(0, i + 1);
                        break;
                    }
                }
            }

        }
        if(aBase2.length > 0 && aCompare2.length > 0)
        {
            if(bOrig)
            {
                oLCS = new AscCommon.LCS(aBase2, aCompare2);
            }
            else
            {
                oLCS = new AscCommon.LCS(aCompare2, aBase2);
            }
            oLCS.equals = fEquals;
            oLCS.forEachCommonSymbol(fLCSCallback);
        }
        oEqualMap.bMatchNoEmpty = bMatchNoEmpty;
        return oEqualMap;
    };
    CDocumentComparison.prototype.compareNotes = function(oMatching)
    {
        for(var key in oMatching.Footnotes)
        {
            if(oMatching.Footnotes.hasOwnProperty(key))
            {
                var oBaseFootnotes = AscCommon.g_oTableId.Get_ById(key);
                var oCompareFootnotes = oMatching.Footnotes[key];
                if(oBaseFootnotes && oCompareFootnotes)
                {
                    this.compareRoots(oBaseFootnotes, oCompareFootnotes);
                }
            }
        }
    };
    CDocumentComparison.prototype.compareShapes = function(oBaseShape, oCompareShape)
    {
        if(oBaseShape.textBoxContent && oCompareShape.textBoxContent)
        {
            this.compareRoots(oBaseShape.textBoxContent, oCompareShape.textBoxContent);
        }
        else if(oBaseShape.textBoxContent && !oCompareShape.textBoxContent)
        {
            this.setReviewInfoRecursive(oBaseShape.textBoxContent, reviewtype_Remove);
        }
        else if(!oBaseShape.textBoxContent && oCompareShape.textBoxContent)
        {
            oBaseShape.setTextBoxContent(oCompareShape.textBoxContent.Copy(oBaseShape, editor.WordControl.m_oDrawingDocument, this.copyPr))
        }
    };
    CDocumentComparison.prototype.compareGroups = function(oBaseGroup, oCompareGroup)
    {
        var oLCS = new AscCommon.LCS(oBaseGroup.spTree, oCompareGroup.spTree);
        oLCS.equals = function(a, b) {
            return a.isComparable(b);
        };

        var oBaseNode = new CNode(oBaseGroup, null);
        var oChildNode;
        for(var nSp = 0; nSp < oBaseGroup.spTree.length; ++nSp)
        {
            oChildNode = new CNode(oBaseGroup.spTree[nSp], oBaseNode);
        }
        var oCompareNode = new CNode(oCompareGroup, null);
        for(nSp = 0; nSp < oCompareGroup.spTree.length; ++nSp)
        {
            oChildNode = new CNode(oCompareGroup.spTree[nSp], oCompareNode);
        }
        var oDiff  = new AscCommon.Diff(oBaseNode, oCompareNode);
        oDiff.equals = function(a, b)
        {
            return a.isComparable(b);
        };
        var oMatching = new CMatching();
        oDiff.matchTrees(oMatching);
        var oDeltaCollector = new AscCommon.DeltaCollector(oMatching, oBaseNode, oCompareNode);
        oDeltaCollector.forEachChange(function(oOperation){
            oOperation.anchor.base.addChange(oOperation);
        });
        oBaseNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(var nChild = 0; nChild < oBaseNode.children.length; ++nChild)
        {
            var oChild = oBaseNode.children[nChild];
            if(oChild.partner)
            {
                this.compareGraphicObject(oChild.element, oChild.partner.element);
            }
        }
        for(var nChange = 0; nChange < oBaseNode.changes.length; ++nChange)
        {
            var oChange = oBaseNode.changes[nChange];
            for(var nRemove = oChange.remove.length - 1; nRemove > -1;  --nRemove)
            {
                var oRemoveSp = oChange.remove[nRemove].element;
                this.setGraphicObjectReviewInfo(oRemoveSp, reviewtype_Remove);
            }
            for(var nInsert = oChange.insert.length - 1; nInsert > -1;  --nInsert)
            {
                var oInsertSp = oChange.insert[nInsert].element.copy({contentCopyPr: this.copyPr});
                oBaseGroup.addToSpTree(oChange.anchor.index, oInsertSp);
                oInsertSp.setGroup(oBaseGroup);
            }
        }
    };
    CDocumentComparison.prototype.setGraphicObjectReviewInfo = function(oGrObj, nType)
    {
        switch (oGrObj.getObjectType())
        {
            case AscDFH.historyitem_type_Shape:
            {
                if(oGrObj.textBoxContent)
                {
                    this.setReviewInfoRecursive(oGrObj.textBoxContent, nType);
                }
                break;
            }
            case AscDFH.historyitem_type_GroupShape:
            {
                for(var nSp = 0; nSp < oGrObj.spTree.length; ++nSp)
                {
                    this.setGraphicObjectReviewInfo(oGrObj.spTree[nSp], nType);
                }
                break;
            }
        }
    };
    CDocumentComparison.prototype.compareDrawingObjects = function(oMatching)
    {
        for(var key in oMatching.Drawings)
        {
            if(oMatching.Drawings.hasOwnProperty(key))
            {
                var oBaseDrawing = AscCommon.g_oTableId.Get_ById(key);
                var oCompareDrawing = oMatching.Drawings[key];
                if(oBaseDrawing && oCompareDrawing)
                {
                    var oBaseGrObject = oBaseDrawing.GraphicObj;
                    var oCompareGrObject = oCompareDrawing.GraphicObj;
                    this.compareGraphicObject(oBaseGrObject, oCompareGrObject);
                }
            }
        }
    };
    CDocumentComparison.prototype.compareGraphicObject = function(oBaseGrObject, oCompareGrObject)
    {
        if(!oBaseGrObject || !oCompareGrObject)
        {
            return;
        }
        var nObjectType = oBaseGrObject.getObjectType();
        if(nObjectType !== oCompareGrObject.getObjectType())
        {
            return;
        }
        switch (nObjectType)
        {
            case AscDFH.historyitem_type_Shape:
            {
                this.compareShapes(oBaseGrObject, oCompareGrObject);
                break;
            }
            case AscDFH.historyitem_type_GroupShape:
            {
                this.compareGroups(oBaseGrObject, oCompareGrObject);
                break;
            }
        }
    };
    CDocumentComparison.prototype.compareRoots = function(oRoot1, oRoot2)
    {

        var oOrigRoot = this.createNodeFromDocContent(oRoot1, null, null);
        var oRevisedRoot =  this.createNodeFromDocContent(oRoot2, null, null);
        var i, j;
        var oEqualMap;
        var aBase, aCompare, bOrig = true;
        if(oOrigRoot.children.length <= oRevisedRoot.children.length)
        {
            aBase = oOrigRoot.children;
            aCompare = oRevisedRoot.children;
        }
        else
        {
            bOrig = false;
            aBase = oRevisedRoot.children;
            aCompare = oOrigRoot.children;
        }

        var aBase2 = [];
        var aCompare2 = [];
        var oCompareMap = {};
        var bMatchNoEmpty;

        var oCurNode, oCompareNode;
        oEqualMap = this.compareElementsArray(aBase, aCompare, bOrig, false);
        bMatchNoEmpty = oEqualMap.bMatchNoEmpty;

        //included paragraphs
        if(bMatchNoEmpty)
        {
            i = 0;
            j = 0;
            oCompareMap = {};

            while(i < aBase.length && j < aCompare.length)
            {
                oCurNode = aBase[i];
                oCompareNode = aCompare[j];
                if(oCurNode.partner && oCompareNode.partner)
                {
                    ++i;
                    ++j;
                }
                else
                {
                    var nStartI = i;
                    var nStartJ = j;
                    var nStartComparIndex = j - 1;
                    var nEndCompareIndex = nStartComparIndex;
                    aCompare2.length = 0;
                    while(j < aCompare.length && !aCompare[j].partner)
                    {
                        aCompare2.push(aCompare[j]);
                        ++j;
                    }
                    nEndCompareIndex = j;
                    if((nEndCompareIndex - nStartComparIndex) > 1)
                    {
                        oCompareMap = {};
                        aBase2.length = 0;
                        while (i < aBase.length && !aBase[i].partner)
                        {
                             oCurNode = aBase[i];
                            aBase2.push(oCurNode);
                            ++i;
                        }

                        if(aBase2.length > 0 && aCompare2.length > 0)
                        {
                            this.compareElementsArray(aBase2, aCompare2, bOrig, true);
                        }
                    }
                    i = nStartI;
                    j = nStartJ;
                    while(j < aCompare.length && !aCompare[j].partner)
                    {
                        ++j;
                    }
                    while(i < aBase.length && !aBase[i].partner)
                    {
                        ++i;
                    }
                }
            }
        }

        j = oRevisedRoot.children.length - 1;
        var aInserContent = [];
        var nRemoveCount = 0;
        for(i = oOrigRoot.children.length - 1; i > -1 ; --i)
        {
            if(!oOrigRoot.children[i].partner)
            {
                this.setReviewInfoRecursive(oOrigRoot.children[i].element, reviewtype_Remove);
                ++nRemoveCount;
            }
            else
            {
                aInserContent.length = 0;
                for(j = oOrigRoot.children[i].partner.childidx + 1;
                    j < oRevisedRoot.children.length && !oRevisedRoot.children[j].partner; ++j)
                {
                    aInserContent.push(oRevisedRoot.children[j]);
                }
                if(aInserContent.length > 0)
                {
                    this.insertNodesToDocContent(oOrigRoot.element, i + 1 + nRemoveCount, aInserContent);
                }
                nRemoveCount = 0;
            }
        }
        aInserContent.length = 0;
        for(j = 0; j < oRevisedRoot.children.length && !oRevisedRoot.children[j].partner; ++j)
        {
            aInserContent.push(oRevisedRoot.children[j]);
        }
        if(aInserContent.length > 0)
        {
            this.insertNodesToDocContent(oOrigRoot.element, nRemoveCount, aInserContent);
        }
    };
    CDocumentComparison.prototype.compare = function()
    {
        var oOriginalDocument = this.originalDocument;
        var oRevisedDocument = this.revisedDocument;
        if(!oOriginalDocument || !oRevisedDocument)
        {
            return;
        }
        var oThis = this;
        var aImages = AscCommon.pptx_content_loader.End_UseFullUrl();
        var oObjectsForDownload = AscCommon.GetObjectsForImageDownload(aImages);
        var oApi = oOriginalDocument.GetApi(), i;
        if(!oApi)
        {
            return;
        }
        var fCallback = function (data) {
            var oImageMap = {};
			AscFormat.ExecuteNoHistory(function () {
				AscCommon.ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
			}, oThis, []);
            oOriginalDocument.StopRecalculate();
            oOriginalDocument.StartAction(AscDFH.historydescription_Document_CompareDocuments);
            oOriginalDocument.Start_SilentMode();
            var oldTrackRevisions = oOriginalDocument.IsTrackRevisions();
            oOriginalDocument.SetTrackRevisions(false);
            var LogicDocuments = oOriginalDocument.TrackRevisionsManager.Get_AllChangesLogicDocuments();
            for (var LogicDocId in LogicDocuments)
            {
                var LogicDoc = AscCommon.g_oTableId.Get_ById(LogicDocId);
                if (LogicDoc)
                {
                    LogicDoc.AcceptRevisionChanges(undefined, true);
                }
            }
            var NewNumbering = oRevisedDocument.Numbering.CopyAllNums(oOriginalDocument.Numbering);
            oRevisedDocument.CopyNumberingMap = NewNumbering.NumMap;
            oOriginalDocument.Numbering.AppendAbstractNums(NewNumbering.AbstractNum);
            oOriginalDocument.Numbering.AppendNums(NewNumbering.Num);
            var key;
            for(key in NewNumbering.NumMap)
            {
                if (NewNumbering.NumMap.hasOwnProperty(key))
                {
                    oThis.checkedNums[NewNumbering.NumMap[key]] = true;
                }
            }
            oThis.compareRoots(oOriginalDocument, oRevisedDocument);
            oThis.compareSectPr(oOriginalDocument, oRevisedDocument);

            var oFonts = oOriginalDocument.Document_Get_AllFontNames();
            var aFonts = [];
            for (i in oFonts)
            {
                if(oFonts.hasOwnProperty(i))
                {
                    aFonts[aFonts.length] = new AscFonts.CFont(i, 0, "", 0, null);
                }
            }
            oApi.pre_Paste(aFonts, oImageMap, function()
            {
                oOriginalDocument.SetTrackRevisions(oldTrackRevisions);
                oOriginalDocument.End_SilentMode(false);
                oOriginalDocument.Recalculate();
                oOriginalDocument.UpdateInterface();
                oOriginalDocument.FinalizeAction();
                oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
            });

        };
        AscCommon.sendImgUrls(oApi, oObjectsForDownload.aUrls, fCallback, null, true);
        return null;
    };
    CDocumentComparison.prototype.getNewParaPrWithDiff = function(oElementPr, oPartnerPr)
    {
        var oOldParaPr = oElementPr.Copy(undefined, undefined);
        var oNewParaPr = oPartnerPr.Copy(undefined, this.copyPr);
        if(oOldParaPr.Is_Equal(oNewParaPr))
        {
            return null;
        }
        oNewParaPr.PrChange = oOldParaPr;
        oNewParaPr.ReviewInfo = new CReviewInfo();
        this.setReviewInfo(oNewParaPr.ReviewInfo);
        return oNewParaPr;
    };
    CDocumentComparison.prototype.isElementForAdd = function (oElement)
    {
        if(oElement.IsParaEndRun && oElement.IsParaEndRun())
        {
            return false;
        }
        return true;
    };
    CDocumentComparison.prototype.applyChangesToParagraph = function(oNode)
    {
        var oElement = oNode.element, oChange, i, j, k, t, oChildElement,
            oChildNode, oLastText, oFirstText, oCurRun, oNewRun, oFirstRun;
        var oParentParagraph, aContentToInsert;
        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(i = 0; i < oNode.changes.length; ++i)
        {
            oChange = oNode.changes[i];
            oLastText = null;
            oFirstText = null;
            aContentToInsert = [];
            if(oChange.insert.length > 0)
            {
                oFirstText = oChange.insert[0].element;
                oLastText = oChange.insert[oChange.insert.length - 1].element;

                var oLastRemoveText = null;
                var oFirstRemoveText = null;
                if(oChange.remove.length > 0)
                {
                    if(oChange.remove[oChange.remove.length - 1].element instanceof CTextElement)
                    {
                        oLastRemoveText = oChange.remove[oChange.remove.length - 1].element;
                    }
                    if(oChange.remove[0].element instanceof CTextElement)
                    {
                        oFirstRemoveText = oChange.remove[0].element;
                    }
                }
                oCurRun = oLastText.lastRun ? oLastText.lastRun : oLastText;
                oFirstRun = oFirstText.firstRun ? oFirstText.firstRun : oFirstText;
                oParentParagraph =  (oNode.partner && oNode.partner.element) || oCurRun.Paragraph;
                var oParentParagraph2 =  oCurRun.Paragraph;
                for(k = oParentParagraph.Content.length - 1; k > -1; --k)
                {
                    if(oCurRun === oParentParagraph.Content[k])
                    {
                        if(oCurRun instanceof ParaRun)
                        {
                            for(t = oCurRun.Content.length - 1; t > -1; --t)
                            {
                                oCurRun.Paragraph = oElement.Paragraph || oElement;
                                oNewRun = oCurRun.Copy2(this.copyPr);
                                oCurRun.Paragraph = oParentParagraph2;
                                if(oLastText.elements[oLastText.elements.length - 1] === oCurRun.Content[t])
                                {
                                    if(t < oCurRun.Content.length - 1)
                                    {
                                        oNewRun.Remove_FromContent(t + 1, oNewRun.Content.length - (t + 1), false);
                                    }
                                    aContentToInsert.splice(0, 0, oNewRun);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            aContentToInsert.splice(0, 0, oCurRun.Copy(false, this.copyPr));
                        }
                        break;
                    }
                    else if(oLastText === oParentParagraph.Content[k])
                    {
                        aContentToInsert.splice(0, 0,  oParentParagraph.Content[k].Copy(false, this.copyPr));
                        break;
                    }
                }
                if( (oLastText.lastRun && oFirstText.firstRun) && oLastText.lastRun === oFirstText.firstRun || (!oLastText.lastRun && !oFirstText.firstRun) && oLastText === oFirstText)
                {
                    if(aContentToInsert.length > 0)
                    {
                        oNewRun = aContentToInsert[0];
                        if(oNewRun instanceof  ParaRun)
                        {
                            for(t = 0; t < oFirstText.firstRun.Content.length; ++t)
                            {
                                if(oFirstText.elements[0] === oFirstText.firstRun.Content[t])
                                {
                                    oNewRun.Remove_FromContent(0, t, false);
                                    break;
                                }
                            }
                        }
                    }
                }
                else
                {
                    for(k -= 1; k > -1; --k)
                    {
                        oCurRun = oParentParagraph.Content[k];
                        if(oCurRun !== oFirstRun && oCurRun !== oFirstText)
                        {
                            aContentToInsert.splice(0, 0, oCurRun.Copy(false, this.copyPr));
                        }
                        else
                        {
                            if(oCurRun === oFirstText)
                            {
                                aContentToInsert.splice(0, 0,  oCurRun.Copy(false, this.copyPr));
                            }
                            else
                            {

                                for(t = 0; t < oCurRun.Content.length; ++t)
                                {
                                    if(oFirstText.elements[0] === oCurRun.Content[t])
                                    {
                                        if(oLastText.lastRun === oFirstText.firstRun)
                                        {
                                            oNewRun = aContentToInsert[0];
                                        }
                                        else
                                        {
                                            oCurRun.Paragraph = oElement.Paragraph || oElement;
                                            oNewRun = oCurRun.Copy2(this.copyPr);
                                            oCurRun.Paragraph = oParentParagraph2;
                                            aContentToInsert.splice(0, 0, oNewRun);
                                        }
                                        oNewRun.Remove_FromContent(0, t, false);
                                    }
                                }
                            }
                            break;
                        }
                    }
                }

                if(oChange.remove.length === 0)
                {
                    if(aContentToInsert.length > 0)
                    {
                        var index = oChange.anchor.index;
                        oChildNode = oNode.children[index];
                        if(oChildNode)
                        {
                            oFirstText = oChildNode.element;
                            for(j = 0; j < oElement.Content.length; ++j)
                            {
                                if(Array.isArray(oElement.Content))
                                {
                                    oCurRun = oElement.Content[j];
                                    if(oFirstText === oCurRun)
                                    {
                                        for(t = aContentToInsert.length - 1; t > - 1; --t)
                                        {
                                            if(this.isElementForAdd(aContentToInsert[t]))
                                            {
                                                oElement.AddToContent(j + 1, aContentToInsert[t]);
                                            }
                                        }
                                        break;
                                    }
                                    else if(Array.isArray(oCurRun.Content) && Array.isArray(oFirstText.elements))
                                    {
                                        for(k = 0; k < oCurRun.Content.length; ++k)
                                        {
                                            if(oFirstText.elements[0] === oCurRun.Content[k])
                                            {
                                                break;
                                            }
                                        }
                                        var bFind = false;
                                        if(k === oCurRun.Content.length)
                                        {
                                            if(oFirstText.firstRun === oCurRun)
                                            {
                                                k = 0;
                                                bFind = true;
                                            }
                                        }
                                        else
                                        {
                                            bFind = true;
                                        }
                                        if(k <= oCurRun.Content.length && bFind)
                                        {
                                            oCurRun.Split2(k, oElement, j);
                                            for(t = aContentToInsert.length - 1; t > - 1; --t)
                                            {
                                                if(this.isElementForAdd(aContentToInsert[t]))
                                                {
                                                    oElement.AddToContent(j + 1, aContentToInsert[t]);
                                                }
                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            //handle removed elements
            if(oChange.remove.length > 0)
            {
                oLastText = oChange.remove[oChange.remove.length - 1].element;
                oFirstText = oChange.remove[0].element;
                if(oLastText.lastRun)
                {
                    oCurRun = oLastText.lastRun;
                }
                else
                {
                    oCurRun = oLastText;
                }

                var nInsertPosition = -1;
                for(k = oElement.Content.length - 1; k > -1; --k)
                {
                    if(oElement.Content[k] === oCurRun)
                    {
                        if(oLastText instanceof CTextElement)
                        {
                            for(t = oCurRun.Content.length - 1; t > -1; t--)
                            {
                                if(oCurRun.Content[t] === oLastText.elements[oLastText.elements.length - 1])
                                {
                                    break;
                                }
                            }
                            if(t > -1)
                            {
                                //  if(t !== oCurRun.Content.length - 1)
                                {
                                    nInsertPosition = k + 1;
                                    oNewRun = oCurRun.Split2(t + 1, oElement, k);
                                    oNewRun.SetReviewTypeWithInfo(reviewtype_Common, oCurRun.ReviewInfo.Copy());
                                }
                                // else
                                // {}
                            }
                        }
                        else
                        {
                            nInsertPosition = k + 1;
                        }
                        break;
                    }
                }
                for(; k > -1; --k)
                {
                    oChildElement = oElement.Content[k];
                    if(oChildElement !== oFirstText.firstRun && oChildElement !== oFirstText)
                    {
                        if(!(oChildElement.IsParaEndRun && oChildElement.IsParaEndRun()))
                        {
                            this.setReviewInfoRecursive(oChildElement, reviewtype_Remove);
                        }
                    }
                    else
                    {
                        if(oChildElement instanceof ParaRun)
                        {
                            for(t = 0; t < oChildElement.Content.length; t++)
                            {
                                if(oChildElement.Content[t] === oFirstText.elements[0])
                                {
                                    break;
                                }
                            }
                            t = Math.min(Math.max(t, 0), oChildElement.Content.length - 1);
                            if(t > 0)
                            {
                                oNewRun = oChildElement.Split2(t, oElement, k);
                                if(!(oNewRun.IsParaEndRun && oNewRun.IsParaEndRun()))
                                {
                                    this.setReviewInfoRecursive(oNewRun, reviewtype_Remove);
                                }
                                nInsertPosition++;
                            }
                            else
                            {

                                if(!(oChildElement.IsParaEndRun && oChildElement.IsParaEndRun()))
                                {
                                    this.setReviewInfoRecursive(oChildElement, reviewtype_Remove);
                                }
                            }
                        }
                        else
                        {
                            this.setReviewInfoRecursive(oChildElement, reviewtype_Remove);
                        }
                        break;
                    }
                }
                if(nInsertPosition > -1)
                {
                    for(t = aContentToInsert.length - 1; t > - 1; --t)
                    {
                        if(this.isElementForAdd(aContentToInsert[t]))
                        {
                            oElement.AddToContent(nInsertPosition, aContentToInsert[t]);
                        }
                    }
                }
            }
        }
        for(i = 0; i < oNode.children.length; ++i)
        {
            oChildNode = oNode.children[i];
            if(Array.isArray(oChildNode.element.Content))
            {
                this.applyChangesToParagraph(oChildNode);
            }
            else
            {
                for(j = 0; j < oChildNode.children.length; ++j)
                {
                    if(oChildNode.children[j].element instanceof CDocumentContent)
                    {
                        this.applyChangesToDocContent(oChildNode.children[j]);
                    }
                }
            }
        }

        if(oNode.partner)
        {
            var oPartnerNode = oNode.partner;
            var oPartnerElement = oPartnerNode.element;
            if(oPartnerElement instanceof Paragraph)
            {
                var oNewParaPr = this.getNewParaPrWithDiff(oElement.Pr, oPartnerElement.Pr);
                if(oNewParaPr)
                {
                    oElement.Set_Pr(oNewParaPr);
                }
                this.compareSectPr(oElement, oPartnerElement);
            }
        }
    };
    CDocumentComparison.prototype.compareSectPr = function(oElement, oPartnerElement)
    {
        var oOrigSectPr = oElement.SectPr;
        var oReviseSectPr = oPartnerElement.SectPr;
        var oOrigContent, oReviseContent;
        if(!oOrigSectPr && oReviseSectPr)
        {
            var oLogicDocument = this.originalDocument;
            var bCopyHdrFtr = true;
            var SectPr = new CSectionPr(oLogicDocument);
            SectPr.Copy(oReviseSectPr, bCopyHdrFtr, this.copyPr);
            if(oElement.Set_SectionPr)
            {
                oElement.Set_SectionPr(SectPr);
            }
        }
        if(oOrigSectPr)
        {
            oOrigContent = oOrigSectPr.HeaderFirst && oOrigSectPr.HeaderFirst.Content;
            oReviseContent = oReviseSectPr && oReviseSectPr.HeaderFirst && oReviseSectPr.HeaderFirst.Content;
            if(oOrigContent && !oReviseContent)
            {
                this.setReviewInfoRecursive(oOrigContent, reviewtype_Remove);
            }
            else if(oOrigContent && oReviseContent)
            {
                this.compareRoots(oOrigContent, oReviseContent);
            }
            else if(!oOrigContent && oReviseContent && oReviseSectPr.HeaderFirst)
            {
                oOrigSectPr.Set_Header_First(oReviseSectPr.HeaderFirst.Copy(this.originalDocument, this.copyPr));
            }


            oOrigContent = oOrigSectPr.HeaderEven && oOrigSectPr.HeaderEven.Content;
            oReviseContent = oReviseSectPr && oReviseSectPr.HeaderEven && oReviseSectPr.HeaderEven.Content;
            if(oOrigContent && !oReviseContent)
            {
                this.setReviewInfoRecursive(oOrigContent, reviewtype_Remove);
            }
            else if(oOrigContent && oReviseContent)
            {
                this.compareRoots(oOrigContent, oReviseContent);
            }
            else if(!oOrigContent && oReviseContent && oReviseSectPr.HeaderEven)
            {
                oOrigSectPr.Set_Header_Even(oReviseSectPr.HeaderEven.Copy(this.originalDocument, this.copyPr));
            }


            oOrigContent = oOrigSectPr.HeaderDefault && oOrigSectPr.HeaderDefault.Content;
            oReviseContent = oReviseSectPr && oReviseSectPr.HeaderDefault && oReviseSectPr.HeaderDefault.Content;
            if(oOrigContent && !oReviseContent)
            {
                this.setReviewInfoRecursive(oOrigContent, reviewtype_Remove);
            }
            else if(oOrigContent && oReviseContent)
            {
                this.compareRoots(oOrigContent, oReviseContent);
            }
            else if(!oOrigContent && oReviseContent && oReviseSectPr.HeaderDefault)
            {
                oOrigSectPr.Set_Header_Default(oReviseSectPr.HeaderDefault.Copy(this.originalDocument, this.copyPr));
            }


            oOrigContent = oOrigSectPr.FooterFirst && oOrigSectPr.FooterFirst.Content;
            oReviseContent = oReviseSectPr && oReviseSectPr.FooterFirst && oReviseSectPr.FooterFirst.Content;
            if(oOrigContent && !oReviseContent)
            {
                this.setReviewInfoRecursive(oOrigContent, reviewtype_Remove);
            }
            else if(oOrigContent && oReviseContent)
            {
                this.compareRoots(oOrigContent, oReviseContent);
            }
            else if(!oOrigContent && oReviseContent && oReviseSectPr.FooterFirst)
            {
                oOrigSectPr.Set_Footer_First(oReviseSectPr.FooterFirst.Copy(this.originalDocument, this.copyPr));
            }


            oOrigContent = oOrigSectPr.FooterEven && oOrigSectPr.FooterEven.Content;
            oReviseContent = oReviseSectPr && oReviseSectPr.FooterEven && oReviseSectPr.FooterEven.Content;
            if(oOrigContent && !oReviseContent)
            {
                this.setReviewInfoRecursive(oOrigContent, reviewtype_Remove);
            }
            else if(oOrigContent && oReviseContent)
            {
                this.compareRoots(oOrigContent, oReviseContent);
            }
            else if(!oOrigContent && oReviseContent && oReviseSectPr.FooterEven)
            {
                oOrigSectPr.Set_Footer_Even(oReviseSectPr.FooterEven.Copy(this.originalDocument, this.copyPr));
            }


            oOrigContent = oOrigSectPr.FooterDefault && oOrigSectPr.FooterDefault.Content;
            oReviseContent = oReviseSectPr && oReviseSectPr.FooterDefault && oReviseSectPr.FooterDefault.Content;
            if(oOrigContent && !oReviseContent)
            {
                this.setReviewInfoRecursive(oOrigContent, reviewtype_Remove);
            }
            else if(oOrigContent && oReviseContent)
            {
                this.compareRoots(oOrigContent, oReviseContent);
            }
            else if(!oOrigContent && oReviseContent && oReviseSectPr.FooterDefault)
            {
                oOrigSectPr.Set_Footer_Default(oReviseSectPr.FooterDefault.Copy(this.originalDocument, this.copyPr));
            }


            if(oReviseSectPr)
            {
                var oReviseHeaderFirst = oReviseSectPr.HeaderFirst;
                var oReviseHeaderEven = oReviseSectPr.HeaderEven;
                var oReviseHeaderDefault = oReviseSectPr.HeaderDefault;
                var oReviseFooterFirst = oReviseSectPr.FooterFirst;
                var oReviseFooterEven = oReviseSectPr.FooterEven;
                var oReviseFooterDefault = oReviseSectPr.FooterDefault;

                oReviseSectPr.HeaderFirst = oOrigSectPr.HeaderFirst;
                oReviseSectPr.HeaderEven = oOrigSectPr.HeaderEven;
                oReviseSectPr.HeaderDefault = oOrigSectPr.HeaderDefault;
                oReviseSectPr.FooterFirst = oOrigSectPr.FooterFirst;
                oReviseSectPr.FooterEven = oOrigSectPr.FooterEven;
                oReviseSectPr.FooterDefault = oOrigSectPr.FooterDefault;

                oOrigSectPr.Copy(oReviseSectPr, false);

                oReviseSectPr.HeaderFirst = oReviseHeaderFirst;
                oReviseSectPr.HeaderEven = oReviseHeaderEven;
                oReviseSectPr.HeaderDefault = oReviseHeaderDefault;
                oReviseSectPr.FooterFirst = oReviseFooterFirst;
                oReviseSectPr.FooterEven = oReviseFooterEven;
                oReviseSectPr.FooterDefault = oReviseFooterDefault;
            }
        }
    };
    CDocumentComparison.prototype.applyChangesToTable = function(oNode)
    {
        var oElement = oNode.element, oChange, i, j, oRow;
        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(i = 0; i < oNode.changes.length; ++i)
        {
            oChange = oNode.changes[i];
            for(j = oChange.remove.length - 1; j > -1;  --j)
            {
                oRow = oChange.remove[j].element;
                this.setReviewInfoRecursive(oRow, reviewtype_Remove);
            }
            for(j = oChange.insert.length - 1; j > -1;  --j)
            {
                oElement.Content.splice(oChange.anchor.index, 0, oChange.insert[j].element.Copy(oElement, this.copyPr));
                AscCommon.History.Add(new CChangesTableAddRow(oElement, oChange.anchor.index, [oElement.Content[oChange.anchor.index]]));
            }
            oElement.Internal_ReIndexing(0);
            if (oElement.Content.length > 0 && oElement.Content[0].Get_CellsCount() > 0)
                oElement.CurCell = oElement.Content[0].Get_Cell(0);
        }
        for(i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToTableRow(oNode.children[i]);
        }
    };
    CDocumentComparison.prototype.applyChangesToTableRow = function(oNode)
    {
        //TODO: handle cell inserts and removes

        for(var i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToDocContent(oNode.children[i]);
        }
    };
    CDocumentComparison.prototype.getCopyNumId = function(sNumId)
    {
        var NewId = undefined;
        if(this.matchedNums[sNumId])
        {
            NewId = this.matchedNums[sNumId];
        }
        else
        {
            if(this.revisedDocument.CopyNumberingMap[sNumId])
            {
                NewId = this.revisedDocument.CopyNumberingMap[sNumId];
                var oCopyNum = AscCommon.g_oTableId.Get_ById(NewId);
                var oOrigNumbering = this.originalDocument.Numbering.Num;
                if(oCopyNum && oOrigNumbering)
                {
                    for(var keyOrig in oOrigNumbering)
                    {
                        if(oOrigNumbering.hasOwnProperty(keyOrig))
                        {
                            if(!this.checkedNums[keyOrig])
                            {
                                var oOrigNum = AscCommon.g_oTableId.Get_ById(keyOrig);
                                if(oOrigNum && oOrigNum.IsSimilar(oCopyNum))
                                {
                                    this.matchedNums[sNumId] = keyOrig;
                                    this.checkedNums[keyOrig] = true;
                                    NewId = keyOrig;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
        return NewId;
    };
    CDocumentComparison.prototype.copyStyleById = function(sId)
    {
        return this.copyStyle(this.revisedDocument.Styles.Get(sId));

    };
    CDocumentComparison.prototype.copyStyle = function(oStyle)
    {
        if(!oStyle)
        {
            return null;
        }
        if(this.StylesMap[oStyle.Id])
        {
            return this.StylesMap[oStyle.Id];
        }
        var oStyleCopy;
        var sStyleId = this.originalDocument.Styles.GetStyleIdByName(oStyle.Name, false);
        if(oStyleCopy = this.originalDocument.Styles.Get(sStyleId))
        {
            this.StylesMap[oStyle.Id] = sStyleId;
            var oNewParaPr = this.getNewParaPrWithDiff(oStyleCopy.ParaPr, oStyle.ParaPr);
            if(oNewParaPr)
            {
                oStyleCopy.Set_ParaPr(oNewParaPr);
            }
            return sStyleId;
        }
        oStyleCopy = oStyle.Copy();
        oStyleCopy.Set_Name(oStyle.Name);
        oStyleCopy.Set_Next(oStyle.Next);
        oStyleCopy.Set_Type(oStyle.Type);
        oStyleCopy.Set_QFormat(oStyle.qFormat);
        oStyleCopy.Set_UiPriority(oStyle.uiPriority);
        oStyleCopy.Set_Hidden(oStyle.hidden);
        oStyleCopy.Set_SemiHidden(oStyle.semiHidden);
        oStyleCopy.Set_UnhideWhenUsed(oStyle.unhideWhenUsed);
        oStyleCopy.Set_TextPr(oStyle.TextPr.Copy(undefined, this.copyPr));
        oStyleCopy.Set_ParaPr( oStyle.ParaPr.Copy(undefined, this.copyPr));
        oStyleCopy.Set_TablePr(oStyle.TablePr.Copy());
        oStyleCopy.Set_TableRowPr(oStyle.TableRowPr.Copy());
        oStyleCopy.Set_TableCellPr(oStyle.TableCellPr.Copy());
        if (undefined !== oStyle.TableBand1Horz)
        {
            oStyleCopy.Set_TableBand1Horz(oStyle.TableBand1Horz.Copy(this.copyPr));
            oStyleCopy.Set_TableBand1Vert(oStyle.TableBand1Vert.Copy(this.copyPr));
            oStyleCopy.Set_TableBand2Horz(oStyle.TableBand2Horz.Copy(this.copyPr));
            oStyleCopy.Set_TableBand2Vert(oStyle.TableBand2Vert.Copy(this.copyPr));
            oStyleCopy.Set_TableFirstCol(oStyle.TableFirstCol.Copy(this.copyPr));
            oStyleCopy.Set_TableFirstRow(oStyle.TableFirstRow.Copy(this.copyPr));
            oStyleCopy.Set_TableLastCol(oStyle.TableLastCol.Copy(this.copyPr));
            oStyleCopy.Set_TableLastRow(oStyle.TableLastRow.Copy(this.copyPr));
            oStyleCopy.Set_TableTLCell(oStyle.TableTLCell.Copy(this.copyPr));
            oStyleCopy.Set_TableTRCell(oStyle.TableTRCell.Copy(this.copyPr));
            oStyleCopy.Set_TableBLCell(oStyle.TableBLCell.Copy(this.copyPr));
            oStyleCopy.Set_TableBRCell(oStyle.TableBRCell.Copy(this.copyPr));
            oStyleCopy.Set_TableWholeTable(oStyle.TableWholeTable.Copy(this.copyPr));
        }
        if(oStyle.BasedOn)
        {
            oStyleCopy.Set_BasedOn(this.copyStyle(this.revisedDocument.Styles.Get(oStyle.BasedOn)));
        }
        this.originalDocument.Styles.Add(oStyleCopy);
        this.StylesMap[oStyle.Id] = oStyleCopy.Id;
        return oStyleCopy.Id;
    };
    CDocumentComparison.prototype.copyComment = function(sId)
    {
        var oCopyComment;
        if(this.CommentsMap[sId])
        {
            oCopyComment = this.CommentsMap[sId];
        }
        else 
        {
            var oComment = this.revisedDocument.Comments.Get_ById(sId);
            if(oComment)
            {
                var oOrigComments = this.originalDocument.Comments;
                if(oOrigComments)
                {
                    var oOldParent = oComment.Parent;
                    oComment.Parent = oOrigComments;
                    var oCopyComment = oComment.Copy();
                    this.CommentsMap[sId] = oCopyComment;
                    this.originalDocument.Comments.Add(oCopyComment);
                    this.api.sync_AddComment(oCopyComment.Id, oCopyComment.Data);
                    oComment.Parent = oOldParent;
                }
            }
        }
        oCopyComment = this.CommentsMap[sId] || null;
        if(oCopyComment) 
        {
            return oCopyComment.Get_Id();
        }
        return null;
    };
    CDocumentComparison.prototype.getRevisedStyle = function(sStyleId)
    {
        if(this.revisedDocument)
        {
            return this.revisedDocument.Styles.Get(sStyleId);
        }
        return null;
    };
    CDocumentComparison.prototype.insertNodesToDocContent = function(oElement, nIndex, aInsert)
    {

        var k = 0;
        for(var j = 0; j < aInsert.length; ++j)
        {
            var oChildElement = null;
            if(aInsert[j].element.Get_Type)
            {
                oChildElement = aInsert[j].element.Copy(oElement, oElement.DrawingDocument, this.copyPr);
            }
            else
            {
                if(aInsert[j].element.Parent && aInsert[j].element.Parent.Get_Type)
                {
                    oChildElement = aInsert[j].element.Parent.Copy(oElement, oElement.DrawingDocument, this.copyPr);
                }
            }
            if(oChildElement)
            {
                oElement.Internal_Content_Add(nIndex + k, oChildElement, false);
                ++k;
            }
        }
    };
    CDocumentComparison.prototype.applyChangesToChildNode = function(oChildNode)
    {
        var oChildElement = oChildNode.element;
        if(oChildElement instanceof Paragraph)
        {
            this.applyChangesToParagraph(oChildNode);
        }
        else if(oChildElement instanceof CDocumentContent)
        {
            this.applyChangesToDocContent(oChildNode);
        }
        else if(oChildElement instanceof CTable)
        {
            this.applyChangesToTable(oChildNode);
        }
    };
    CDocumentComparison.prototype.applyChangesToDocContent = function(oNode)
    {

        if(oNode.partner)
        {
            this.compareRoots(oNode.element, oNode.partner.element);
            return;
        }
        var oElement = oNode.element, oChange, i, j, oChildElement, oChildNode;

        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(i = 0; i < oNode.changes.length; ++i)
        {
            oChange = oNode.changes[i];
            for(j = oChange.remove.length - 1; j > -1; --j)
            {
                oChildNode = oChange.remove[j];
                oChildElement = oChildNode.element;
                this.setReviewInfoRecursive(oChildElement, reviewtype_Remove);
            }
            this.insertNodesToDocContent(oElement, oChange.anchor.index + oChange.remove.length, oChange.insert);
        }
        for(i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToChildNode(oNode.children[i]);
        }
    };
    CDocumentComparison.prototype.setReviewInfo = function(oReviewIno)
    {
        oReviewIno.Editor   = this.api;
        oReviewIno.UserId   = "";
        oReviewIno.MoveType = Asc.c_oAscRevisionsMove.NoMove;
        oReviewIno.PrevType = -1;
        oReviewIno.PrevInfo = null;
        oReviewIno.UserName = this.getUserName();
        var oCore = this.revisedDocument.Core;
        if(oCore)
        {
            if(oCore.modified instanceof Date)
            {
                oReviewIno.DateTime = oCore.modified.getTime();
            }
        }
        else
        {
            oReviewIno.DateTime = "Unknown";
        }
    };
    CDocumentComparison.prototype.setReviewInfoRecursive = function(oObject, nType)
    {
        if(!oObject)
        {
            return;
        }
        if(oObject.ReviewInfo && oObject.SetReviewTypeWithInfo)
        {
            var oReviewIno = oObject.ReviewInfo.Copy();
            this.setReviewInfo(oReviewIno);
            oObject.SetReviewTypeWithInfo(nType, oReviewIno, false);
        }
        var i;
        if(Array.isArray(oObject.Content))
        {
            for(i = 0; i < oObject.Content.length; ++i)
            {
                this.setReviewInfoRecursive(oObject.Content[i], nType)
            }
        }
        if(AscCommon.isRealObject(oObject.Content))
        {
            this.setReviewInfoRecursive(oObject.Content, nType);
        }
        if(oObject.Type === para_FootnoteReference || oObject.Type === para_EndnoteReference)
        {
            this.setReviewInfoRecursive(oObject.Footnote, nType);
        }
        if(oObject.GetAllDocContents)
        {
            var aContents = oObject.GetAllDocContents();
            for(i = 0; i < aContents.length; ++i)
            {
                this.setReviewInfoRecursive(aContents[i], nType);
            }
        }
        if(oObject.Root)
        {
            this.setReviewInfoRecursive(oObject.Root, nType);
        }
        if(AscCommon.isRealObject(oObject.SectPr) && (oObject instanceof Paragraph))
        {
            var oOrigSectPr = oObject.SectPr, oOrigContent;
            if(oOrigSectPr)
            {
                oOrigContent = oOrigSectPr.HeaderFirst && oOrigSectPr.HeaderFirst.Content;
                if(oOrigContent)
                {
                    this.setReviewInfoRecursive(oOrigContent, nType);
                }
                oOrigContent = oOrigSectPr.HeaderEven && oOrigSectPr.HeaderEven.Content;
                if(oOrigContent)
                {
                    this.setReviewInfoRecursive(oOrigContent, nType);
                }
                oOrigContent = oOrigSectPr.HeaderDefault && oOrigSectPr.HeaderDefault.Content;
                if(oOrigContent)
                {
                    this.setReviewInfoRecursive(oOrigContent, nType);
                }
                oOrigContent = oOrigSectPr.FooterFirst && oOrigSectPr.FooterFirst.Content;
                if(oOrigContent)
                {
                    this.setReviewInfoRecursive(oOrigContent, nType);
                }

                oOrigContent = oOrigSectPr.FooterEven && oOrigSectPr.FooterEven.Content;
                if(oOrigContent)
                {
                    this.setReviewInfoRecursive(oOrigContent, nType);
                }
                oOrigContent = oOrigSectPr.FooterDefault && oOrigSectPr.FooterDefault.Content;
                if(oOrigContent)
                {
                    this.setReviewInfoRecursive(oOrigContent, nType);
                }
            }
        }

    };
    CDocumentComparison.prototype.updateReviewInfo = function(oObject, nType, bParaEnd)
    {
        if(oObject.ReviewInfo && oObject.SetReviewTypeWithInfo)
        {
            var oReviewIno = oObject.ReviewInfo.Copy();
            this.setReviewInfo(oReviewIno);
            oObject.SetReviewTypeWithInfo(nType, oReviewIno, false);

        }
    };
    CDocumentComparison.prototype.createNodeFromDocContent = function(oElement, oParentNode, oHashWords)
    {
        var oRet = new CNode(oElement, oParentNode);
        var bRoot = (oParentNode === null);
        for(var i = 0; i < oElement.Content.length; ++i)
        {
            var oChElement = oElement.Content[i];
            if(oChElement instanceof Paragraph)
            {
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                }
                var oParagraphNode = this.createNodeFromRunContentElement(oChElement, oRet, oHashWords);
                if(bRoot)
                {
                    oParagraphNode.hashWords = oHashWords;
                }
            }
            else if(oChElement instanceof CBlockLevelSdt)
            {
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                }
                var oBlockNode = this.createNodeFromDocContent(oChElement.Content, oRet, oHashWords);
                if(bRoot)
                {
                    oBlockNode.hashWords = oHashWords;
                }
            }
            else if(oChElement instanceof CTable)
            {
                if(this.options.getTables())
                {
                    if(bRoot)
                    {
                        oHashWords = new Minhash({});
                    }
                    var oTableNode = new CNode(oChElement, oRet);
                    if(bRoot)
                    {
                        oHashWords = new Minhash({});
                        oTableNode.hashWords = oHashWords;
                    }
                    for(var j = 0; j < oChElement.Content.length; ++j)
                    {
                        var oRowNode = new CNode(oChElement.Content[j], oTableNode);
                        for(var k = 0; k < oChElement.Content[j].Content.length; ++k)
                        {
                            this.createNodeFromDocContent(oChElement.Content[j].Content[k].Content, oRowNode, oHashWords);
                        }
                    }
                }
            }
            else
            {
                var oNode = new CNode(oChElement, oRet);
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                    oNode.hashWords = oHashWords;
                }
            }

        }
        return oRet;
    };
    CDocumentComparison.prototype.createNodeFromRunContentElement = function(oElement, oParentNode, oHashWords)
    {
        var oRet = new CNode(oElement, oParentNode);
        var oLastText = null, oRun, oRunElement, i, j;
        var aLastWord = [];
        for(i = 0; i < oElement.Content.length; ++i)
        {
            oRun = oElement.Content[i];
            if(oRun instanceof ParaRun)
            {
                if(oRun.Content.length > 0)
                {
                   // if(this.options.getWords())
                    {
                        if(!oLastText)
                        {
                            oLastText = new CTextElement();
                            oLastText.setFirstRun(oRun);
                        }
                        if(oLastText.elements.length === 0)
                        {
                            oLastText.setFirstRun(oRun);
                            oLastText.setLastRun(oRun);
                        }
                        for(j = 0; j < oRun.Content.length; ++j)
                        {
                            oRunElement = oRun.Content[j];
                            var bPunctuation = para_Text === oRunElement.Type && (AscCommon.g_aPunctuation[oRunElement.Value] && !EXCLUDED_PUNCTUATION[oRunElement.Value]);
                            if(oRunElement.Type === para_Space || oRunElement.Type === para_Tab
                                || oRunElement.Type === para_Separator || oRunElement.Type === para_NewLine
                                || oRunElement.Type === para_FootnoteReference
                                || oRunElement.Type === para_EndnoteReference
                                || bPunctuation)
                            {
                                if(oLastText.elements.length > 0)
                                {
                                    new CNode(oLastText, oRet);
                                    oLastText.updateHash(oHashWords);
                                    oLastText = new CTextElement();
                                    oLastText.setFirstRun(oRun);
                                }

                                oLastText.setLastRun(oRun);
                                oLastText.elements.push(oRunElement);
                                new CNode(oLastText, oRet);
                                oLastText.updateHash(oHashWords);

                                oLastText = new CTextElement();
                                oLastText.setFirstRun(oRun);
                                oLastText.setLastRun(oRun);
                            }
                            else if(oRunElement.Type === para_Drawing)
                            {
                                if(oLastText.elements.length > 0)
                                {
                                    oLastText.updateHash(oHashWords);
                                    new CNode(oLastText, oRet);
                                    oLastText = new CTextElement();
                                    oLastText.setFirstRun(oRun);
                                    oLastText.setLastRun(oRun);
                                }
                                oLastText.elements.push(oRun.Content[j]);
                                new CNode(oLastText, oRet);
                                oLastText = new CTextElement();
                                oLastText.setFirstRun(oRun);
                                oLastText.setLastRun(oRun);
                            }
                            else if(oRunElement.Type === para_End)
                            {
                                if(oLastText.elements.length > 0)
                                {
                                    oLastText.updateHash(oHashWords);
                                    new CNode(oLastText, oRet);
                                    oLastText = new CTextElement();
                                    oLastText.setFirstRun(oRun);
                                    oLastText.setLastRun(oRun);
                                }
                                oLastText.setFirstRun(oRun);
                                oLastText.setLastRun(oRun);
                                oLastText.elements.push(oRun.Content[j]);
                                new CNode(oLastText, oRet);
                                oLastText.updateHash(oHashWords);
                                oLastText = new CTextElement();
                                oLastText.setFirstRun(oRun);
                                oLastText.setLastRun(oRun);
                            }
                            else
                            {
                                if(oLastText.elements.length === 0)
                                {
                                    oLastText.setFirstRun(oRun);
                                }
                                oLastText.setLastRun(oRun);
                                oLastText.elements.push(oRun.Content[j]);
                            }
                        }
                    }
                    // else
                    // {
                    //     if(oLastText && oLastText.elements.length > 0)
                    //     {
                    //         new CNode(oLastText, oRet);
                    //     }
                    //     for(j = 0; j < oRun.Content.length; ++j)
                    //     {
                    //         oRunElement = oRun.Content[j];
                    //         if(AscFormat.isRealNumber(oRunElement.Value))
                    //         {
                    //             aLastWord.push(oRunElement.Value);
                    //         }
                    //         else
                    //         {
                    //             if(aLastWord.length > 0)
                    //             {
                    //                 oHashWords.update(aLastWord);
                    //                 aLastWord.length = 0;
                    //             }
                    //         }
                    //         oLastText = new CTextElement();
                    //         oLastText.setFirstRun(oRun);
                    //         oLastText.setLastRun(oRun);
                    //         oLastText.elements.push(oRunElement);
                    //         new CNode(oLastText, oRet);
                    //     }
                    //     oLastText = new CTextElement();
                    //     oLastText.setFirstRun(oRun);
                    //     oLastText.setLastRun(oRun);
                    // }
                }
            }
            else
            {
                if(!(oRun instanceof CParagraphBookmark))
                {
                    if(oLastText && oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oHashWords);
                        new CNode(oLastText, oRet);
                    }
                    if(aLastWord.length > 0)
                    {
                        oHashWords.update(aLastWord);
                        aLastWord.length = 0;
                    }
                    oLastText = null;
                    if(Array.isArray(oRun.Content))
                    {
                        this.createNodeFromRunContentElement(oRun, oRet, oHashWords);
                    }
                    else
                    {
                        new CNode(oRun, oRet);
                    }
                }
            }
        }
        if(oLastText && oLastText.elements.length > 0)
        {
            oLastText.updateHash(oHashWords);
            new CNode(oLastText, oRet);
        }
        return oRet;
    };
    CDocumentComparison.prototype.createFootNote = function()
    {
        return this.originalDocument.Footnotes.CreateFootnote();
    };
    CDocumentComparison.prototype.createEndNote = function()
    {
        return this.originalDocument.Endnotes.CreateEndnote();
    };
    CDocumentComparison.prototype.getComment = function(sId)
    {
        var oComment = this.originalDocument.Comments.Get_ById(sId);
        if(oComment)
        {
            return oComment;
        }
        oComment = this.revisedDocument.Comments.Get_ById(sId);
        return oComment || null;
    };


    window['AscCommonWord'] = window['AscCommonWord'] || {};
    window['AscCommonWord'].CDocumentComparison = CDocumentComparison;
    window['AscCommonWord'].ComparisonOptions = window['AscCommonWord']["ComparisonOptions"] = ComparisonOptions;

    function CompareBinary(oApi, sBinary2, oOptions, bForceApplyChanges)
    {

        var oDoc1 = oApi.WordControl.m_oLogicDocument;
        if(!window['NATIVE_EDITOR_ENJINE'])
        {
            var oCollaborativeEditing = oDoc1.CollaborativeEditing;
            if(oCollaborativeEditing && !oCollaborativeEditing.Is_SingleUser())
            {
                oApi.sendEvent("asc_onError", Asc.c_oAscError.ID.CannotCompareInCoEditing, c_oAscError.Level.NoCritical);
                return;
            }
        }
        oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
        var bHaveRevisons2 = false;
        var oDoc2 = AscFormat.ExecuteNoHistory(function(){

            var oBinaryFileReader, openParams        = {checkFileSize : /*this.isMobileVersion*/false, charCount : 0, parCount : 0, disableRevisions: true, noSendComments: true};
            var oDoc2 = new CDocument(oApi.WordControl.m_oDrawingDocument, true);
            oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oDoc2;
            oApi.WordControl.m_oLogicDocument = oDoc2;
            oBinaryFileReader = new AscCommonWord.BinaryFileReader(oDoc2, openParams);
            AscCommon.pptx_content_loader.Start_UseFullUrl(oApi.insertDocumentUrlsData);
            if (!oBinaryFileReader.Read(sBinary2))
            {
                oDoc2 = null;
            }
            if(oDoc2)
            {
                bHaveRevisons2 = oBinaryFileReader.oReadResult && oBinaryFileReader.oReadResult.hasRevisions;
            }
            oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oDoc1;
            oApi.WordControl.m_oLogicDocument = oDoc1;
			if (oDoc1.History)
				oDoc1.History.Set_LogicDocument(oDoc1);
			if (oDoc1.CollaborativeEditing)
				oDoc1.CollaborativeEditing.m_oLogicDocument = oDoc1;
            return oDoc2;
        }, this, []);
        oDoc1.History.Document = oDoc1;
        if(oDoc2)
        {
            var fCallback = function()
            {
                var oComp = new AscCommonWord.CDocumentComparison(oDoc1, oDoc2, oOptions ? oOptions : new ComparisonOptions());
                oComp.compare();
            };

            if(window['NATIVE_EDITOR_ENJINE'] || bForceApplyChanges)
            {
                fCallback();
            }
            else
            {

                oDoc1.TrackRevisionsManager.ContinueTrackRevisions();
                if(oDoc1.TrackRevisionsManager.Have_Changes() || bHaveRevisons2)
                {

                    oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
                    oApi.sendEvent("asc_onAcceptChangesBeforeCompare", function (bAccept) {
                        if(bAccept){
                            oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
                            fCallback();
                        }
                        else
                        {
                        }
                    })
                }
                else
                {
                    fCallback();
                }
            }
        }
        else
        {
            AscCommon.pptx_content_loader.End_UseFullUrl();
        }
    }


    function CompareDocuments(oApi, oTmpDocument)
    {
        oApi.insertDocumentUrlsData = {
            imageMap: oTmpDocument["GetImageMap"](), documents: [], convertCallback: function (_api, url) {
            }, endCallback: function (_api) {
            }
        };
        CompareBinary(oApi, oTmpDocument["GetBinary"](), null, true);
        oApi.insertDocumentUrlsData = null;
    }

    window['AscCommonWord']["CompareBinary"] =  window['AscCommonWord'].CompareBinary = CompareBinary;
    window['AscCommonWord']["ComparisonOptions"] = window['AscCommonWord'].ComparisonOptions = ComparisonOptions;
    window['AscCommonWord']['CompareDocuments'] = CompareDocuments;
})();
