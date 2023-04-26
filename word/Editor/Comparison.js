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

(function (undefined) {

    const MIN_JACCARD = 0.34;
    const MIN_DIFF = 0.7;
    const EXCLUDED_PUNCTUATION = {};
    //EXCLUDED_PUNCTUATION[46] = true;
    //EXCLUDED_PUNCTUATION[95] = true;
    EXCLUDED_PUNCTUATION[160] = true;
    //EXCLUDED_PUNCTUATION[63] = true;

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
		// debug method
	CNode.prototype.getText = function()
	{
		return this.element.getText();
	};

    CNode.prototype.cleanEndOfInsert = function (aContentToInsert, idxOfChange, comparison) {
        const oChange = this.changes[idxOfChange];
        const oLastText = oChange.insert[oChange.insert.length - 1].element;
        const oEndOfInsertRun = oLastText.lastRun ? oLastText.lastRun : oLastText;
        const oParentParagraph =  (this.partner && this.partner.element) || oEndOfInsertRun.Paragraph;
        const applyingParagraph = this.getApplyParagraph(comparison);

        let k = oParentParagraph.Content.length - 1;
        let lastCheckRun;
        for(k; k > -1; --k)
        {
            // если мы встретили последний ран, где встречается слово
            const oCurrentRun = oParentParagraph.Content[k];
            if(oEndOfInsertRun === oCurrentRun)
            {
                if(oEndOfInsertRun instanceof ParaRun)
                {
                    for(let t = oEndOfInsertRun.Content.length - 1; t > -1; --t)
                    {
                        const oNewRun = this.copyRunWithMockParagraph(oEndOfInsertRun, applyingParagraph.Paragraph || applyingParagraph, comparison);
                        //очищаем конец слова, которое нужно вставить
                        if(oLastText.elements[oLastText.elements.length - 1] === oEndOfInsertRun.Content[t])
                        {
                            if(t < oEndOfInsertRun.Content.length - 1)
                            {
                                lastCheckRun = oNewRun.Split2(t + 1);
                                comparison.checkOriginalAndSplitRun(oNewRun, lastCheckRun);
                                this.setCommonReviewTypeWithInfo(lastCheckRun, oEndOfInsertRun.ReviewInfo.Copy());
                            }
                            this.pushToArrInsertContent(aContentToInsert, oNewRun, comparison);
                            break;
                        }
                    }
                }
                else
                {
                //целиком вставим то, что встретили
                    this.pushToArrInsertContentWithCopy(aContentToInsert, oEndOfInsertRun, comparison);
                }
                break;
            }
            else if(oLastText === oCurrentRun)
            {
                //целиком вставим то, что встретили
                this.pushToArrInsertContentWithCopy(aContentToInsert, oCurrentRun, comparison);
                break;
            }
        }
        return k;
    }
    // comparison need for extends
    CNode.prototype.pushToArrInsertContent = function (aContentToInsert, elem, comparison) {
        aContentToInsert.push(elem);
    }

    CNode.prototype.pushToArrInsertContentWithCopy = function (aContentToInsert, elem, comparison) {
        const elemCopy = elem.Copy(false, comparison.copyPr);
        this.pushToArrInsertContent(aContentToInsert, elemCopy, comparison);
    }

    CNode.prototype.cleanStartOfInsertSameRun = function (oNewRun, idxOfChange) {
        const oChange = this.changes[idxOfChange];
        const oFirstText = oChange.insert[0].element;
        if(oNewRun)
        {
            if(oNewRun instanceof ParaRun)
            {
                for(let t = 0; t < oFirstText.firstRun.Content.length; ++t)
                {
                    // удаляем начало рана до изменения в слове
                    if(oFirstText.elements[0] === oFirstText.firstRun.Content[t])
                    {
                        oNewRun.Remove_FromContent(0, t, false);
                        break;
                    }
                }
            }
        }
    }

    CNode.prototype.copyRunWithMockParagraph = function (oRun, mockParagraph, comparison) {
        const oTempParagraph = oRun.Paragraph;
        oRun.Paragraph = mockParagraph;
        const oNewRun = oRun.Copy2(comparison.copyPr);
        oRun.Paragraph = oTempParagraph;

        return oNewRun;
    }

    CNode.prototype.cleanStartOfInsertDifferentRun = function (aContentToInsert, posOfLastInsertRun, idxOfChange, comparison) {
        const oChange = this.changes[idxOfChange];
        const oFirstText = oChange.insert[0].element;
        const oFirstRun = oFirstText.firstRun ? oFirstText.firstRun : oFirstText;
        const oLastText = oChange.insert[oChange.insert.length - 1].element;
        const applyingParagraph = this.getApplyParagraph(comparison);

        let oCurRun = oLastText.lastRun ? oLastText.lastRun : oLastText;
        const oParentParagraph =  (this.partner && this.partner.element) || oCurRun.Paragraph;
        let k = posOfLastInsertRun;
        let lastCheckRun;
        for(k -= 1; k > -1; --k)
        {
            oCurRun = oParentParagraph.Content[k];
            // Пока не дошли до первого рана слова, закидываем его на добавление
            if(!(oCurRun === oFirstRun || oCurRun === oFirstText))
            {
                this.pushToArrInsertContentWithCopy(aContentToInsert, oCurRun, comparison);
            }
            else
            {
                if(oCurRun === oFirstText)
                {
                    this.pushToArrInsertContentWithCopy(aContentToInsert, oCurRun, comparison);
                }
                else
                {

                    for(let t = 0; t < oCurRun.Content.length; ++t)
                    {
                        if(oFirstText.elements[0] === oCurRun.Content[t])
                        {
                            let oNewRun;
                            if(oLastText.lastRun === oFirstText.firstRun)
                            {
                                lastCheckRun = aContentToInsert[0];
                                oNewRun = lastCheckRun.Split2(t);
                                comparison.checkOriginalAndSplitRun(lastCheckRun, oNewRun);
                            }
                            else
                            {
                                lastCheckRun = this.copyRunWithMockParagraph(oCurRun, applyingParagraph.Paragraph || applyingParagraph, comparison);
                                oNewRun = lastCheckRun.Split2(t);
                                comparison.checkOriginalAndSplitRun(lastCheckRun, oNewRun);
                                this.pushToArrInsertContent(aContentToInsert, oNewRun, comparison);
                                this.setCommonReviewTypeWithInfo(lastCheckRun, oCurRun.ReviewInfo.Copy());
                            }
                        }
                    }
                }
                break;
            }
        }
    }

    CNode.prototype.getArrOfInsertsFromChanges = function (idxOfChange, comparison) {
        const oChange = this.changes[idxOfChange];
        const aContentToInsert = [];

        if(oChange.insert.length > 0)
        {
            const oFirstText = oChange.insert[0].element;
            const oLastText = oChange.insert[oChange.insert.length - 1].element;

            const posLastRunOfInsert = this.cleanEndOfInsert(aContentToInsert, idxOfChange, comparison);

            // изменения находятся внутри одного рана или это один и тот же элемент
            if( (oLastText.lastRun && oFirstText.firstRun) && oLastText.lastRun === oFirstText.firstRun || (!oLastText.lastRun && !oFirstText.firstRun) && oLastText === oFirstText)
            {
                this.cleanStartOfInsertSameRun(aContentToInsert[0], idxOfChange);
            }
            else
            {
                this.cleanStartOfInsertDifferentRun(aContentToInsert, posLastRunOfInsert, idxOfChange, comparison);
            }
        }
        return aContentToInsert;
    };


    CNode.prototype.applyInsertsToParagraph = function (comparison, aContentToInsert, idxOfChange) {
        const oChange = this.changes[idxOfChange];
        if (oChange.remove.length > 0) {
            this.applyInsertsToParagraphsWithRemove(comparison, aContentToInsert, idxOfChange);
        } else {
            this.applyInsertsToParagraphsWithoutRemove(comparison, aContentToInsert, idxOfChange);
        }
    }

    CNode.prototype.getStartPosition = function (comparison) {
        return 0;
    }

    CNode.prototype.applyInsertsToParagraphsWithoutRemove = function (comparison, aContentToInsert, idxOfChange) {
        const oChange = this.changes[idxOfChange];
        const applyingParagraph = this.getApplyParagraph(comparison);

        if(aContentToInsert.length > 0)
        {
            const index = oChange.anchor.index;
            const oChildNode = this.children[index];
            if(oChildNode)
            {
                const oFirstText = oChildNode.element;
                for(let j = 0; j < applyingParagraph.Content.length; ++j)
                {
                    if(Array.isArray(applyingParagraph.Content))
                    {
                        const oCurRun = applyingParagraph.Content[j];
                        // если совпали ран, после которого нужно вставлять и ран из цикла
                        if(oFirstText === oCurRun)
                        {
                            this.applyInsert(aContentToInsert, [], j + 1, comparison);
                            break;
                        }
                        // иначе надо посмотреть, возможно стоит вставлять элементы не после рана, а после конкретного элемента и текущий ран из цикла нужно засплитить
                        else if(Array.isArray(oCurRun.Content) && Array.isArray(oFirstText.elements))
                        {
                            let k = 0;
                            for(k; k < oCurRun.Content.length; ++k)
                            {
                                // если элементы совпали, значит, мы нашли место вставки
                                if(oFirstText.elements[0] === oCurRun.Content[k])
                                {
                                    break;
                                }
                            }
                            let bFind = false;
                            // проверим, не дошли ли мы просто до конца массива, ничего не встретив
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
                                const oNewRun = oCurRun.Split2(k, applyingParagraph, j);
                                comparison.checkOriginalAndSplitRun(oCurRun, oNewRun)
                                //TODO: think about it
                                this.applyInsert(aContentToInsert, [], j + 1, comparison);
                                return true;
                            }
                        }
                    }
                }
            }
        }
        return false;
    };

    CNode.prototype.getApplyParagraph = function (comparison) {
        return this.element;
    }

    CNode.prototype.setCommonReviewTypeWithInfo = function (element, info) {
        element.SetReviewTypeWithInfo(reviewtype_Common, info);
    }

    CNode.prototype.prepareEndOfRemoveChange = function (idxOfChange, comparison, arrSetRemove) {
        const oChange = this.changes[idxOfChange];
        const oApplyParagraph = this.getApplyParagraph(comparison);
        const oLastText = oChange.remove[oChange.remove.length - 1].element;
        const oEndOfRemoveRun = oLastText.lastRun || oLastText;

        let k = oApplyParagraph.Content.length - 1;
        let nInsertPosition = -1;
        
        for(k; k > -1; --k)
        {
            const oCurRun = oApplyParagraph.Content[k];
            if(oCurRun === oEndOfRemoveRun)
            {
                if(oLastText instanceof CTextElement)
                {
                    let t = oEndOfRemoveRun.Content.length - 1;
                    for(t; t > -1; t--)
                    {
                        if(oEndOfRemoveRun.Content[t] === oLastText.elements[oLastText.elements.length - 1])
                        {
                            break;
                        }
                    }
                    if(t > -1)
                    {
                        nInsertPosition = k + 1;
                        const oNewRun = oEndOfRemoveRun.Split2(t + 1, oApplyParagraph, k);
                        comparison.checkOriginalAndSplitRun(oEndOfRemoveRun, oNewRun);
                        this.setCommonReviewTypeWithInfo(oNewRun, oEndOfRemoveRun.ReviewInfo.Copy());
                    }
                }
                else
                {
                    nInsertPosition = k + 1;
                }
                break;
            }
        }
        return {posLastRunInContent: k, nInsertPosition: nInsertPosition };
    }

    CNode.prototype.setReviewTypeForRemoveChanges = function (comparison, idxOfChange, posLastRunInContent, nInsertPosition, arrSetRemoveReviewType) {
        const oApplyParagraph = this.getApplyParagraph(comparison);
        const oChange = this.changes[idxOfChange];
        const oFirstText = oChange.remove[0].element;

        let k = posLastRunInContent;

        let bBreak = false;
        for(k; k > -1; --k)
        {
            const oChildElement = oApplyParagraph.Content[k];
            if(!(oChildElement === oFirstText.firstRun || oChildElement === oFirstText))
            {
                arrSetRemoveReviewType.push(oChildElement);
            }
            else
            {
                if(oChildElement instanceof ParaRun)
                {
                    let t = 0;
                    for(t; t < oChildElement.Content.length; t++)
                    {
                        if(oChildElement.Content[t] === oFirstText.elements[0])
                        {
                            break;
                        }
                    }
                    t = Math.min(Math.max(t, 0), oChildElement.Content.length - 1);
                    if(t > 0)
                    {
                        const oNewRun = oChildElement.Split2(t, oApplyParagraph, k);
                        comparison.checkOriginalAndSplitRun(oChildElement, oNewRun);
                        arrSetRemoveReviewType.push(oNewRun);
                        nInsertPosition++;
                    }
                    else
                    {
                        arrSetRemoveReviewType.push(oChildElement);
                    }
                }
                else
                {
                    arrSetRemoveReviewType.push(oChildElement);
                }
                break;
            }
        }
        return nInsertPosition;
    }

    CNode.prototype.applyInsert = function (arrToInsert, arrToRemove, nInsertPosition, comparison, options) {
        for (let i = 0; i < arrToRemove.length; i += 1) {
            comparison.setRemoveReviewType(arrToRemove[i]);
        }
        this.insertContentAfterRemoveChanges(arrToInsert, nInsertPosition, comparison);
    }

    CNode.prototype.applyInsertsToParagraphsWithRemove = function (comparison, aContentToInsert, idxOfChange) {
        const arrSetRemoveReviewType = [];
        const infoAboutEndOfRemoveChange = this.prepareEndOfRemoveChange(idxOfChange, comparison, arrSetRemoveReviewType);
        const posLastRunInContent = infoAboutEndOfRemoveChange.posLastRunInContent;

        let nInsertPosition = infoAboutEndOfRemoveChange.nInsertPosition;
        nInsertPosition = this.setReviewTypeForRemoveChanges(comparison, idxOfChange, posLastRunInContent, nInsertPosition, arrSetRemoveReviewType);

        this.applyInsert(aContentToInsert, arrSetRemoveReviewType, nInsertPosition, comparison, {needReverse: true});
    };

    CNode.prototype.insertContentAfterRemoveChanges = function (aContentToInsert, nInsertPosition, comparison) {
        const oElement = this.getApplyParagraph(comparison);
        if(nInsertPosition > -1)
        {
            for (let t = 0; t < aContentToInsert.length; t += 1) {
                if(this.isElementForAdd(aContentToInsert[t]))
                {
                    oElement.AddToContent(nInsertPosition, aContentToInsert[t]);
                }
            }
        }
    }

    CNode.prototype.getNeighbors  = function()
    {
        return [this.getLeftNeighbor(), this.getRightNeighbor()];
    };
    CNode.prototype.getLeftNeighbor = function () {
        if(!this.par) {
            return;
        }
        for (let i = this.childidx - 1; i >= 0; i -= 1) {
            if (this.par.children[i].element instanceof CTextElement) {
                return this.par.children[i];
            }
        }
    };
    CNode.prototype.getRightNeighbor = function () {
        if(!this.par) {
            return;
        }
        for (let i = this.childidx + 1; i < this.par.children.length; i += 1) {
            if (this.par.children[i].element instanceof CTextElement) {
                return this.par.children[i];
            }
        }
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
            const oParent1 = this.par;
            const oParent2 = oNode.par;
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
        const oElement1 = this.element;
        const oElement2 = oNode.element;
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
									if (!oElement1.compareReviewElements(oElement2))
									{
										return false;
									}
                    const aNeighbors1 = this.getNeighbors();
                    const aNeighbors2 = oNode.getNeighbors();
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
                    return oElement1.equals(oElement2, bCheckNeighbors);
                }
            }
            if (oElement1 instanceof AscCommon.CParaRevisionMove)
            {
                return false;
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
            return !(oElement1 instanceof AscCommonWord.ParaMath);
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
    
    CNode.prototype.forEachRight = function (callback, T) {
        const arrNodes = [];
        this.forEach(function (oNode) {
            arrNodes.push(oNode);
        });
        for (let i = arrNodes.length - 1; i > -1; i -= 1) {
            callback.call(T, arrNodes[i]);
        }
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
	CTextElement.prototype.compareReviewElements = function (oAnotherElement)
	{
		return true;
	}
    CTextElement.prototype.getPosOfStart = function () {
        const startElement = this.elements[0];
        return this.firstRun.GetElementPosition(startElement);
    }

    CTextElement.prototype.addToElements = function (element, options) {
        this.elements.push(element);
    }

    CTextElement.prototype.getPosOfEnd = function () {
        const endElement = this.elements[this.elements.length - 1];
        return this.lastRun.GetElementPosition(endElement);
    }

    CTextElement.prototype.getElement = function (idx) {
        return this.elements[idx];
    }

    CTextElement.prototype.equals = function (other, bNeedCheckReview)
    {
        if(this.elements.length !== other.elements.length)
        {
            return false;
        }
        for(let i = 0; i < this.elements.length; ++i)
        {
            const oElement = this.getElement(i);
            const oOtherElement = other.getElement(i);
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
        const aCheckArray = [];

        let bVal = false;
        for(let i = 0; i < this.elements.length; ++i)
        {
            const oElement = this.elements[i];
            if(AscFormat.isRealNumber(oElement.Value))
            {
                aCheckArray.push(oElement.Value);
                bVal = true;
            }
            else
            {
				if (oElement instanceof AscWord.CRunBreak)
				{
					if (oElement.IsLineBreak())
					{
						aCheckArray.push(0x000A);
					}
					else
					{
						aCheckArray.push(0x21A1);
					}
				}
                else if(oElement instanceof AscWord.CRunTab)
                {
                    aCheckArray.push(0x0009);
                }
                else if(oElement instanceof AscWord.CRunSpace)
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
        let sResultString = "";
        for(let i = 0; i < this.elements.length; ++i)
        {
            if(this.elements[i] instanceof AscWord.CRunText)
            {
                sResultString += String.fromCharCode(this.elements[i].GetCodePoint());
            }
            else if(this.elements[i] instanceof AscWord.CRunSpace)
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

		//debug method
	CTextElement.prototype.getText = function ()
	{
		return {text:this.elements.map(function (e) {return String.fromCharCode(e.Value)}).join(''), isReviewWord: this.isReviewWord, reviewElements:this.reviewElementTypes};
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
            let oBaseContent = this.elements[0].Footnote;
            let oCompareContent = oTextElement.elements[0].Footnote;
            if(oBaseContent && oCompareContent)
            {
                if(!AscCommon.g_oTableId.Get_ById(oBaseContent.Id))
                {
                    const t = oBaseContent;
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
            const oElement = this.elements[0];
            const oOtherElement = oTextElement.elements[0];
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
        if(oNode1.element instanceof CTextElement)
        {
            const aFootnotes = oNode1.element.compareFootnotes(oNode2.element);
            if(aFootnotes)
            {
                this.Footnotes[aFootnotes[0].Id] = aFootnotes[1];
            }
            else
            {
                const aDrawings = oNode1.element.compareDrawings(oNode2.element);
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
        this.bSaveCustomReviewType = false;
        this.copyPr = {
            CopyReviewPr: false,
            Comparison: this
        };
        this.nInsertChangesType = reviewtype_Add;
        this.nRemoveChangesType = reviewtype_Remove;
        this.oComparisonMoveMarkManager = new CMoveMarkComparisonManager();
    }
    CDocumentComparison.prototype.checkOriginalAndSplitRun = function (oOriginalRun, oSplitRun) {

    }
    CDocumentComparison.prototype.checkCopyParaRun = function (oNewRun, oOldRun) {
        const sMoveName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOldRun);
        this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveName, oNewRun);
        const nMoveReviewType = oOldRun.GetReviewMoveType();
        if (AscFormat.isRealNumber(nMoveReviewType) && nMoveReviewType !== Asc.c_oAscRevisionsMove.NoMove)
        {
            this.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sMoveName, oNewRun);
        }
        if (this.copyPr.SkipUpdateInfo)
        {
            this.saveReviewInfo(oNewRun, oOldRun);
        } else if (this.copyPr.bSaveCustomReviewType)
        {
            this.saveCustomReviewInfo(oNewRun, oOldRun, this.nInsertChangesType);
        } else
        {
            this.updateReviewInfo(oNewRun, this.nInsertChangesType);
        }
    }
    CDocumentComparison.prototype.setRemoveReviewType = function (element) {
        if(!(element.IsParaEndRun && element.IsParaEndRun()))
        {
            this.setReviewInfoRecursive(element, this.nRemoveChangesType);
        }
    };
    CDocumentComparison.prototype.getUserName = function()
    {
        const oCore = this.revisedDocument.Core;
        if(oCore && typeof oCore.lastModifiedBy === "string" && oCore.lastModifiedBy.length > 0)
        {
            return  oCore.lastModifiedBy.split(";")[0];
        }
        else
        {
            return AscCommon.translateManager.getValue("Author");
        }
    };

    CDocumentComparison.prototype.getMinJaccardCoefficient = function () {
        return MIN_JACCARD;
    }

    CDocumentComparison.prototype.getMinDiffCoefficient = function () {
        return MIN_DIFF;
    }
    CDocumentComparison.prototype.getLCSCallback = function (oLCS, bOrig) {
        const oThis = this;
        return function(x, y) {
            const oOrigNode = oLCS.a[x];
            const oReviseNode = oLCS.b[y];
            const oDiff  = new AscCommon.Diff(oOrigNode, oReviseNode);
            oDiff.equals = function(a, b)
            {
                return a.equals(b);
            };
            const oMatching = new CMatching();
            oDiff.matchTrees(oMatching);
            const oDeltaCollector = new AscCommon.DeltaCollector(oMatching, oOrigNode, oReviseNode);
            oDeltaCollector.forEachChange(oThis.forEachChangeCallback);
            oThis.compareDrawingObjectsFromMatching(oMatching, bOrig);
            oThis.applyChangesToChildNode(oOrigNode);
            oThis.compareNotes(oMatching);
        };
    }

    CDocumentComparison.prototype.forEachChangeCallback = function(oOperation) {
        oOperation.anchor.base.addChange(oOperation);
    };
    
    CDocumentComparison.prototype.getLCSEqualsMethod = function (oEqualMap, oMapEquals) {
        return function(a, b) {
            const bEquals = oMapEquals[a.element.Id] || oMapEquals[b.element.Id];
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
    }
    
    CDocumentComparison.prototype.compareElementsArray = function(aBase, aCompare, bOrig, bUseMinDiff)
    {
        const oMapEquals = {};
        const aBase2 = [];
        const aCompare2 = [];
        const oCompareMap = {};
        const MIN_JACCARD_COEFFICIENT = this.getMinJaccardCoefficient();
        const MIN_DIFF_COEFFICIENT = this.getMinDiffCoefficient();

        let bMatchNoEmpty = false;
        let oEqualMap = {};
        for(let i = 0; i < aBase.length; ++i)
        {
            const oCurNode =  aBase[i];
            if(oCurNode.hashWords)
            {
                const oCurInfo = {
                    jaccard: 0,
                    map: {},
                    minDiff: 0,
                    intersection: 0
                };
                oEqualMap[oCurNode.element.Id] = oCurInfo;
                for(let j = 0; j < aCompare.length; ++j)
                {
                    const oCompareNode = aCompare[j];
                    if(oCompareNode.hashWords && oCurNode.isComparable(oCompareNode))
                    {
                        let dJaccard = oCurNode.hashWords.jaccard(oCompareNode.hashWords);
                        if(oCurNode.element instanceof CTable)
                        {
                            dJaccard += MIN_JACCARD_COEFFICIENT;
                        }
                        const dIntersection = dJaccard*(oCurNode.hashWords.count + oCompareNode.hashWords.count)/(1+dJaccard);

                        if(dJaccard > 0)
                        {
                            let diffA = 0, diffB = 0, dMinDiff = 0;
                            if(oCurNode.hashWords.count > 0)
                            {
                                diffA = dIntersection/oCurNode.hashWords.count;
                            }
                            if(oCompareNode.hashWords.count > 0)
                            {
                                diffB = dIntersection/oCompareNode.hashWords.count;
                            }
                            dMinDiff = Math.max(diffA, diffB);

                            if(oCurInfo.jaccard <= dJaccard && dJaccard >= MIN_JACCARD_COEFFICIENT || (oCurInfo.jaccard < MIN_JACCARD_COEFFICIENT && dMinDiff > MIN_DIFF_COEFFICIENT && oCurInfo.minDiff <= dMinDiff))
                            {
                                if(oCurInfo.jaccard < dJaccard && dJaccard >= MIN_JACCARD_COEFFICIENT)
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
                if(oCurInfo.jaccard >= MIN_JACCARD_COEFFICIENT || (bUseMinDiff && oCurInfo.minDiff > MIN_DIFF_COEFFICIENT && oCurNode.hashWords.countLetters > 0 ))
                {
                    aBase2.push(oCurNode);
                    for(let key in oCurInfo.map)
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
        for(let j = 0; j < aCompare.length; ++j)
        {
            const oCompareNode = aCompare[j];
            if(oCompareMap[oCompareNode.element.Id])
            {
                aCompare2.push(oCompareNode);
            }
        }
        if(!bMatchNoEmpty)
        {
            if(bOrig)
            {
                for(let i = 0; i < aBase2.length; ++i)
                {
                    if(i !== aBase2[i].childidx)
                    {
                        aBase2.splice(i, aBase2[i].length - i);
                        break;
                    }
                }
                for(let i = aCompare2.length - 1; i > -1; i--)
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

                for(let i = 0; i < aCompare2.length; ++i)
                {
                    if(i !== aCompare2[i].childidx)
                    {
                        aCompare2.splice(i, aCompare2[i].length - i);
                        break;
                    }
                }
                for(let i = aBase2.length - 1; i > -1; i--)
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
            let oLCS;
            if(bOrig)
            {
                oLCS = new AscCommon.LCS(aBase2, aCompare2);
            }
            else
            {
                oLCS = new AscCommon.LCS(aCompare2, aBase2);
            }
            const fLCSCallback = this.getLCSCallback(oLCS, bOrig);
            oLCS.equals = this.getLCSEqualsMethod(oEqualMap, oMapEquals);
            oLCS.forEachCommonSymbol(fLCSCallback);
        }
        oEqualMap.bMatchNoEmpty = bMatchNoEmpty;
        return oEqualMap;
    };
    CDocumentComparison.prototype.compareNotes = function(oMatching)
    {
        for(let key in oMatching.Footnotes)
        {
            if(oMatching.Footnotes.hasOwnProperty(key))
            {
                const oBaseFootnotes = AscCommon.g_oTableId.Get_ById(key);
                const oCompareFootnotes = oMatching.Footnotes[key];
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
            this.setRemoveReviewType(oBaseShape.textBoxContent);
        }
        else if(!oBaseShape.textBoxContent && oCompareShape.textBoxContent)
        {
            oBaseShape.setTextBoxContent(oCompareShape.textBoxContent.Copy(oBaseShape, editor.WordControl.m_oDrawingDocument, this.copyPr))
        }
    };
    CDocumentComparison.prototype.compareGroups = function(oBaseGroup, oCompareGroup)
    {
        const NodeConstructor = this.getNodeConstructor();
        const oLCS = new AscCommon.LCS(oBaseGroup.spTree, oCompareGroup.spTree);
        oLCS.equals = function(a, b) {
            return a.isComparable(b);
        };

        const oBaseNode = new NodeConstructor(oBaseGroup, null);
        for(let nSp = 0; nSp < oBaseGroup.spTree.length; ++nSp)
        {
            new NodeConstructor(oBaseGroup.spTree[nSp], oBaseNode);
        }
        const oCompareNode = new NodeConstructor(oCompareGroup, null);
        for(let nSp = 0; nSp < oCompareGroup.spTree.length; ++nSp)
        {
            new NodeConstructor(oCompareGroup.spTree[nSp], oCompareNode);
        }
        const oDiff  = new AscCommon.Diff(oBaseNode, oCompareNode);
        oDiff.equals = function(a, b)
        {
            return a.isComparable(b);
        };
        const oMatching = new CMatching();
        oDiff.matchTrees(oMatching);
        const oDeltaCollector = new AscCommon.DeltaCollector(oMatching, oBaseNode, oCompareNode);
        oDeltaCollector.forEachChange(function(oOperation){
            oOperation.anchor.base.addChange(oOperation);
        });
        oBaseNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(let nChild = 0; nChild < oBaseNode.children.length; ++nChild)
        {
            const oChild = oBaseNode.children[nChild];
            if(oChild.partner)
            {
                this.compareGraphicObject(oChild.element, oChild.partner.element);
            }
        }
        for(let nChange = 0; nChange < oBaseNode.changes.length; ++nChange)
        {
            const oChange = oBaseNode.changes[nChange];
            for(let nRemove = oChange.remove.length - 1; nRemove > -1;  --nRemove)
            {
                const oRemoveSp = oChange.remove[nRemove].element;
                this.setGraphicObjectReviewInfo(oRemoveSp, this.nRemoveChangesType);
            }
            for(let nInsert = oChange.insert.length - 1; nInsert > -1;  --nInsert)
            {
                const oInsertSp = oChange.insert[nInsert].element.copy({contentCopyPr: this.copyPr});
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
                for(let nSp = 0; nSp < oGrObj.spTree.length; ++nSp)
                {
                    this.setGraphicObjectReviewInfo(oGrObj.spTree[nSp], nType);
                }
                break;
            }
        }
    };

    CDocumentComparison.prototype.compareDrawingObjects = function (oBaseDrawing, oCompareDrawing, bOrig) {
        if(oBaseDrawing && oCompareDrawing)
        {
            const oBaseGrObject = oBaseDrawing.GraphicObj;
            const oCompareGrObject = oCompareDrawing.GraphicObj;
            this.compareGraphicObject(oBaseGrObject, oCompareGrObject);
        }
    }

    CDocumentComparison.prototype.compareDrawingObjectsFromMatching = function(oMatching, bOrig)
    {
        for(let key in oMatching.Drawings)
        {
            if(oMatching.Drawings.hasOwnProperty(key))
            {
                const oBaseDrawing = AscCommon.g_oTableId.Get_ById(key);
                const oCompareDrawing = oMatching.Drawings[key];
                this.compareDrawingObjects(oBaseDrawing, oCompareDrawing, bOrig);
            }
        }
    };
    CDocumentComparison.prototype.compareGraphicObject = function(oBaseGrObject, oCompareGrObject)
    {
        if(!oBaseGrObject || !oCompareGrObject)
        {
            return;
        }
        const nObjectType = oBaseGrObject.getObjectType();
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

    CDocumentComparison.prototype.applyParagraphComparison = function (oOrigRoot, oRevisedRoot) {
        const aInsertContent = [];
        let nRemoveCount = 0;
        for(let i = oOrigRoot.children.length - 1; i > -1 ; --i)
        {
            if(!oOrigRoot.children[i].partner)
            {
                this.setReviewInfoRecursive(oOrigRoot.children[i].element, this.nRemoveChangesType);
                ++nRemoveCount;
            }
            else
            {
                aInsertContent.length = 0;
                for(let j = oOrigRoot.children[i].partner.childidx + 1;
                    j < oRevisedRoot.children.length && !oRevisedRoot.children[j].partner; ++j)
                {
                    aInsertContent.push(oRevisedRoot.children[j]);
                }
                if(aInsertContent.length > 0)
                {
                    this.insertNodesToDocContent(oOrigRoot.element, i + 1 + nRemoveCount, aInsertContent);
                }
                nRemoveCount = 0;
            }
        }
        aInsertContent.length = 0;
        for(let j = 0; j < oRevisedRoot.children.length && !oRevisedRoot.children[j].partner; ++j)
        {
            aInsertContent.push(oRevisedRoot.children[j]);
        }
        if(aInsertContent.length > 0)
        {
            this.insertNodesToDocContent(oOrigRoot.element, nRemoveCount, aInsertContent);
        }
    };

    CDocumentComparison.prototype.compareRoots = function(oRoot1, oRoot2)
    {
        const oOrigRoot = this.createNodeFromDocContent(oRoot1, null, null, true);
        const oRevisedRoot =  this.createNodeFromDocContent(oRoot2, null, null, false);
        let aBase, aCompare, bOrig = true;
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

        const aBase2 = [];
        const aCompare2 = [];
        const oEqualMap = this.compareElementsArray(aBase, aCompare, bOrig, false);
        const bMatchNoEmpty = oEqualMap.bMatchNoEmpty;

        //included paragraphs
        if(bMatchNoEmpty)
        {
            let i = 0;
            let j = 0;
            let oCompareMap = {};

            while(i < aBase.length && j < aCompare.length)
            {
                let oCurNode = aBase[i];
                let oCompareNode = aCompare[j];
                if(oCurNode.partner && oCompareNode.partner)
                {
                    ++i;
                    ++j;
                }
                else
                {
                    const nStartI = i;
                    const nStartJ = j;
                    const nStartCompareIndex = j - 1;
                    let nEndCompareIndex = nStartCompareIndex;
                    aCompare2.length = 0;
                    while(j < aCompare.length && !aCompare[j].partner)
                    {
                        aCompare2.push(aCompare[j]);
                        ++j;
                    }
                    nEndCompareIndex = j;
                    if((nEndCompareIndex - nStartCompareIndex) > 1)
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

        this.applyParagraphComparison(oOrigRoot, oRevisedRoot);
    };
    CDocumentComparison.prototype.compare = function(callback)
    {
        const oOriginalDocument = this.originalDocument;
        const oRevisedDocument = this.revisedDocument;
        if(!oOriginalDocument || !oRevisedDocument)
        {
            return;
        }
        const oThis = this;
        const aImages = AscCommon.pptx_content_loader.End_UseFullUrl();
        const oObjectsForDownload = AscCommon.GetObjectsForImageDownload(aImages);
        const oApi = oOriginalDocument.GetApi();
        if(!oApi)
        {
            return;
        }
        const fCallback = function (data) {
            const oImageMap = {};
			AscFormat.ExecuteNoHistory(function () {
				AscCommon.ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
			}, oThis, []);
            oOriginalDocument.StopRecalculate();
            oOriginalDocument.StartAction(AscDFH.historydescription_Document_CompareDocuments);
            oOriginalDocument.Start_SilentMode();
            const oldTrackRevisions = oOriginalDocument.IsTrackRevisions();
            oOriginalDocument.SetTrackRevisions(false);
            const LogicDocuments = oOriginalDocument.TrackRevisionsManager.Get_AllChangesLogicDocuments();
            for (let LogicDocId in LogicDocuments)
            {
                const LogicDoc = AscCommon.g_oTableId.Get_ById(LogicDocId);
                if (LogicDoc)
                {
                    LogicDoc.AcceptRevisionChanges(undefined, true);
                }
            }
            const NewNumbering = oRevisedDocument.Numbering.CopyAllNums(oOriginalDocument.Numbering);
            oRevisedDocument.CopyNumberingMap = NewNumbering.NumMap;
            oOriginalDocument.Numbering.AppendAbstractNums(NewNumbering.AbstractNum);
            oOriginalDocument.Numbering.AppendNums(NewNumbering.Num);
            for(let key in NewNumbering.NumMap)
            {
                if (NewNumbering.NumMap.hasOwnProperty(key))
                {
                    oThis.checkedNums[NewNumbering.NumMap[key]] = true;
                }
            }
            oThis.compareRoots(oOriginalDocument, oRevisedDocument);
            oThis.compareSectPr(oOriginalDocument, oRevisedDocument);

            const oFonts = oOriginalDocument.Document_Get_AllFontNames();
            const aFonts = [];
            for (let i in oFonts)
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
                callback && callback();
            });
        };
        AscCommon.sendImgUrls(oApi, oObjectsForDownload.aUrls, fCallback, true);
        return null;
    };
    CDocumentComparison.prototype.getNewParaPrWithDiff = function(oElementPr, oPartnerPr)
    {
        const oOldParaPr = oElementPr.Copy(undefined, undefined);
        const oNewParaPr = oPartnerPr.Copy(undefined, this.copyPr);
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
        return !(oElement.IsParaEndRun && oElement.IsParaEndRun());

    };
    CNode.prototype.isElementForAdd = CDocumentComparison.prototype.isElementForAdd;
    CDocumentComparison.prototype.executeWithCheckInsertAndRemove = function (callback, oChange) {
        callback();
    };
    CDocumentComparison.prototype.applyChangesToParagraph = function(oNode)
    {
        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(let i = 0; i < oNode.changes.length; ++i)
        {
            this.executeWithCheckInsertAndRemove(function () {
                const aContentToInsert = oNode.getArrOfInsertsFromChanges(i, this);
                //handle removed elements
                oNode.applyInsertsToParagraph(this, aContentToInsert, i);
            }.bind(this), oNode.changes[i]);

        }
        this.applyChangesToChildrenOfParagraphNode(oNode);
        this.applyChangesToSectPr(oNode);
    };

    CDocumentComparison.prototype.applyChangesToChildrenOfParagraphNode = function (oNode) {
        for(let i = 0; i < oNode.children.length; ++i)
        {
            const oChildNode = oNode.children[i];
            if(Array.isArray(oChildNode.element.Content))
            {
                this.applyChangesToParagraph(oChildNode);
            }
            else
            {
                for(let j = 0; j < oChildNode.children.length; ++j)
                {
                    if(oChildNode.children[j].element instanceof CDocumentContent)
                    {
                        this.applyChangesToDocContent(oChildNode.children[j]);
                    }
                }
            }
        }
    }
    
    CDocumentComparison.prototype.applyChangesToSectPr = function (oNode) {
        const oElement = oNode.element;
        if(oNode.partner)
        {
            const oPartnerNode = oNode.partner;
            const oPartnerElement = oPartnerNode.element;
            if(oPartnerElement instanceof Paragraph)
            {
                const oNewParaPr = this.getNewParaPrWithDiff(oElement.Pr, oPartnerElement.Pr);
                if(oNewParaPr)
                {
                    oElement.Set_Pr(oNewParaPr);
                }
                this.compareSectPr(oElement, oPartnerElement);
            }
        }
    }

    CDocumentComparison.prototype.compareSectPr = function(oElement, oPartnerElement)
    {
        const oOrigSectPr = oElement.SectPr;
        const oReviseSectPr = oPartnerElement.SectPr;
        let oOrigContent, oReviseContent;
        if(!oOrigSectPr && oReviseSectPr)
        {
            const oLogicDocument = this.originalDocument;
            const bCopyHdrFtr = true;
            const SectPr = new CSectionPr(oLogicDocument);
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
                this.setReviewInfoRecursive(oOrigContent, this.nRemoveChangesType);
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
                this.setReviewInfoRecursive(oOrigContent, this.nRemoveChangesType);
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
                this.setReviewInfoRecursive(oOrigContent, this.nRemoveChangesType);
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
                this.setReviewInfoRecursive(oOrigContent, this.nRemoveChangesType);
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
                this.setReviewInfoRecursive(oOrigContent, this.nRemoveChangesType);
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
                this.setReviewInfoRecursive(oOrigContent, this.nRemoveChangesType);
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
                const oReviseHeaderFirst = oReviseSectPr.HeaderFirst;
                const oReviseHeaderEven = oReviseSectPr.HeaderEven;
                const oReviseHeaderDefault = oReviseSectPr.HeaderDefault;
                const oReviseFooterFirst = oReviseSectPr.FooterFirst;
                const oReviseFooterEven = oReviseSectPr.FooterEven;
                const oReviseFooterDefault = oReviseSectPr.FooterDefault;

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
        this.applyChangesToTableSize(oNode);
        this.applyChangesToTableRows(oNode);
    };
    CDocumentComparison.prototype.applyChangesToTableSize = function(oNode)
    {
        const oElement = oNode.element;
        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(let i = 0; i < oNode.changes.length; ++i)
        {
            const oChange = oNode.changes[i];
            for(let j = oChange.remove.length - 1; j > -1;  --j)
            {
                const oRow = oChange.remove[j].element;
                this.setRemoveReviewType(oRow);
            }
            for(let j = oChange.insert.length - 1; j > -1;  --j)
            {
                oElement.Content.splice(oChange.anchor.index, 0, oChange.insert[j].element.Copy(oElement, this.copyPr));
                AscCommon.History.Add(new CChangesTableAddRow(oElement, oChange.anchor.index, [oElement.Content[oChange.anchor.index]]));
            }
            oElement.Internal_ReIndexing(0);
            if (oElement.Content.length > 0 && oElement.Content[0].Get_CellsCount() > 0)
                oElement.CurCell = oElement.Content[0].Get_Cell(0);
        }
    };

    CDocumentComparison.prototype.applyChangesToTableRows = function(oNode)
    {
        for(let i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToTableRow(oNode.children[i]);
        }
    };
    CDocumentComparison.prototype.checkRowReview = function(oRowNode) {};
    CDocumentComparison.prototype.applyChangesToTableRow = function(oNode)
    {
        this.checkRowReview(oNode);
        //TODO: handle cell inserts and removes

        for(let i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToDocContent(oNode.children[i]);
        }
    };
    CDocumentComparison.prototype.getCopyNumId = function(sNumId)
    {
        let NewId;
        if(this.matchedNums[sNumId])
        {
            NewId = this.matchedNums[sNumId];
        }
        else
        {
            if(this.revisedDocument.CopyNumberingMap[sNumId])
            {
                NewId = this.revisedDocument.CopyNumberingMap[sNumId];
                const oCopyNum = AscCommon.g_oTableId.Get_ById(NewId);
                const oOrigNumbering = this.originalDocument.Numbering.Num;
                if(oCopyNum && oOrigNumbering)
                {
                    for(let keyOrig in oOrigNumbering)
                    {
                        if(oOrigNumbering.hasOwnProperty(keyOrig))
                        {
                            if(!this.checkedNums[keyOrig])
                            {
                                const oOrigNum = AscCommon.g_oTableId.Get_ById(keyOrig);
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

        const sStyleId = this.originalDocument.Styles.GetStyleIdByName(oStyle.Name, false);
        let oStyleCopy = this.originalDocument.Styles.Get(sStyleId);
        if(oStyleCopy)
        {
            this.StylesMap[oStyle.Id] = sStyleId;
            const oNewParaPr = this.getNewParaPrWithDiff(oStyleCopy.ParaPr, oStyle.ParaPr);
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
        let oCopyComment;
        if(this.CommentsMap[sId])
        {
            oCopyComment = this.CommentsMap[sId];
        }
        else
        {
            const oComment = this.revisedDocument.Comments.Get_ById(sId);
            if(oComment)
            {
                const oOrigComments = this.originalDocument.Comments;
                if(oOrigComments)
                {
                    const oOldParent = oComment.Parent;
                    oComment.Parent = oOrigComments;
                    oCopyComment = oComment.Copy();
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
        let k = 0;
        for(let j = 0; j < aInsert.length; ++j)
        {
            let oChildElement;
            let oOriginalElement;
            if(aInsert[j].element.Get_Type)
            {
                oOriginalElement = aInsert[j].element;
                oChildElement = oOriginalElement.Copy(oElement, oElement.DrawingDocument, this.copyPr);
            }
            else
            {
                if(aInsert[j].element.Parent && aInsert[j].element.Parent.Get_Type)
                {
                    oOriginalElement = aInsert[j].element.Parent;
                    oChildElement = oOriginalElement.Copy(oElement, oElement.DrawingDocument, this.copyPr);
                }
            }
            if(oChildElement)
            {

                if (oOriginalElement instanceof Paragraph) {
                    const oParaEnd = oOriginalElement.GetParaEndRun();
                    let oRunMoveMark = oParaEnd.GetLastTrackMoveMark();
                    if (oRunMoveMark) {
                        oRunMoveMark = oRunMoveMark.Copy();
                        const sMoveMarkName = this.oComparisonMoveMarkManager.getChangedMoveMarkName(oRunMoveMark);
                        oRunMoveMark.Name = sMoveMarkName;
                        const oCopyParaEnd = oChildElement.GetParaEndRun();
                        this.oComparisonMoveMarkManager.addRevisedMoveMarkToInserts(oRunMoveMark, oCopyParaEnd, oChildElement, true);
                    }
                }
                oElement.Internal_Content_Add(nIndex + k, oChildElement, false);
                ++k;
            }
        }
    };
    CDocumentComparison.prototype.applyChangesToChildNode = function(oChildNode)
    {
        const oChildElement = oChildNode.element;
        if(oChildElement instanceof Paragraph || oChildElement instanceof AscCommonWord.CMockParagraph)
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
        const oElement = oNode.element;
        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(let i = 0; i < oNode.changes.length; ++i)
        {
            const oChange = oNode.changes[i];
            for(let j = oChange.remove.length - 1; j > -1; --j)
            {
                const oChildNode = oChange.remove[j];
                const oChildElement = oChildNode.element;
                this.setReviewInfoRecursive(oChildElement, this.nRemoveChangesType);
            }
            this.insertNodesToDocContent(oElement, oChange.anchor.index + oChange.remove.length, oChange.insert);
        }
        for(let i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToChildNode(oNode.children[i]);
        }
    };
    CDocumentComparison.prototype.setReviewInfo = function(oReviewInfo)
    {
        oReviewInfo.Editor   = this.api;
        oReviewInfo.UserId   = "";
        oReviewInfo.MoveType = Asc.c_oAscRevisionsMove.NoMove;
        oReviewInfo.PrevType = -1;
        oReviewInfo.PrevInfo = null;
        oReviewInfo.UserName = this.getUserName();
        const oCore = this.revisedDocument.Core;
        if(oCore)
        {
            if(oCore.modified instanceof Date)
            {
                oReviewInfo.DateTime = oCore.modified.getTime();
            }
        }
        else
        {
            oReviewInfo.DateTime = "Unknown";
        }
    };
    CDocumentComparison.prototype.getElementsForSetReviewType = function (oObject) {
        if (!oObject) {
            return [];
        }
        const arrReturnObjects = [];
        const arrCheckObjects = [oObject];

        while (arrCheckObjects.length) {
            const oCheckObject = arrCheckObjects.pop();

            if(oCheckObject.ReviewInfo && oCheckObject.SetReviewTypeWithInfo)
            {
               arrReturnObjects.push(oCheckObject);
            }
            if(Array.isArray(oCheckObject.Content))
            {
                for(let i = 0; i < oCheckObject.Content.length; ++i)
                {
                    arrCheckObjects.push(oCheckObject.Content[i]);
                }
            }
            if(AscCommon.isRealObject(oCheckObject.Content))
            {
                arrCheckObjects.push(oCheckObject.Content);
            }
            if(oCheckObject.Type === para_FootnoteReference || oCheckObject.Type === para_EndnoteReference)
            {
                arrCheckObjects.push(oCheckObject.Footnote);
            }
            if(oCheckObject.GetAllDocContents)
            {
                const aContents = oCheckObject.GetAllDocContents();
                for(let i = 0; i < aContents.length; ++i)
                {
                    arrCheckObjects.push(aContents[i]);
                }
            }
            if(oCheckObject.Root)
            {
                arrCheckObjects.push(oCheckObject.Root);
            }
            if(AscCommon.isRealObject(oCheckObject.SectPr) && (oCheckObject instanceof Paragraph))
            {
                const oOrigSectPr = oCheckObject.SectPr;
                let oOrigContent;
                if(oOrigSectPr)
                {
                    oOrigContent = oOrigSectPr.HeaderFirst && oOrigSectPr.HeaderFirst.Content;
                    if(oOrigContent)
                    {
                        arrCheckObjects.push(oOrigContent);
                    }
                    oOrigContent = oOrigSectPr.HeaderEven && oOrigSectPr.HeaderEven.Content;
                    if(oOrigContent)
                    {
                        arrCheckObjects.push(oOrigContent);
                    }
                    oOrigContent = oOrigSectPr.HeaderDefault && oOrigSectPr.HeaderDefault.Content;
                    if(oOrigContent)
                    {
                        arrCheckObjects.push(oOrigContent);
                    }
                    oOrigContent = oOrigSectPr.FooterFirst && oOrigSectPr.FooterFirst.Content;
                    if(oOrigContent)
                    {
                        arrCheckObjects.push(oOrigContent);
                    }

                    oOrigContent = oOrigSectPr.FooterEven && oOrigSectPr.FooterEven.Content;
                    if(oOrigContent)
                    {
                        arrCheckObjects.push(oOrigContent);
                    }
                    oOrigContent = oOrigSectPr.FooterDefault && oOrigSectPr.FooterDefault.Content;
                    if(oOrigContent)
                    {
                        arrCheckObjects.push(oOrigContent);
                    }
                }
            }
        }

        return arrReturnObjects;
    };

    CDocumentComparison.prototype.setReviewInfoForArray = function (arrNeedReviewObjects, nType, sCustomReviewUserName, nCustomReviewDate) {
        for (let i = 0; i < arrNeedReviewObjects.length; i += 1) {
            const oNeedReviewObject = arrNeedReviewObjects[i];
            if (oNeedReviewObject.SetReviewTypeWithInfo) {
                let oReviewInfo = oNeedReviewObject.ReviewInfo.Copy();
                this.setReviewInfo(oReviewInfo, sCustomReviewUserName, nCustomReviewDate);
                let reviewType;
                let moveReviewType;
                if (this.bSaveCustomReviewType) {
                    reviewType = oNeedReviewObject.GetReviewType && oNeedReviewObject.GetReviewType();
                    moveReviewType = oNeedReviewObject.GetReviewMoveType && oNeedReviewObject.GetReviewMoveType();
                }
                const newReviewType = reviewType !== undefined && reviewType !== reviewtype_Common ? reviewType : null;
                const newMoveReviewType = moveReviewType !== undefined && moveReviewType !== Asc.c_oAscRevisionsMove.NoMove ? moveReviewType : null;
                if (!(AscFormat.isRealNumber(newReviewType) || AscFormat.isRealNumber(newMoveReviewType))) {
                    oNeedReviewObject.SetReviewTypeWithInfo(nType, oReviewInfo, false);
                }
            }
        }
    }

    CDocumentComparison.prototype.setReviewInfoRecursive = function(oObject, nType)
    {
        const arrNeedReviewObjects = this.getElementsForSetReviewType(oObject);
        this.setReviewInfoForArray(arrNeedReviewObjects, nType);
    };
    CDocumentComparison.prototype.saveReviewInfo = function(oNewObject, oOldObject)
    {
        if (oOldObject.GetReviewInfo && oNewObject.SetReviewTypeWithInfo) {
            const reviewType = oOldObject.GetReviewType && oOldObject.GetReviewType();
            const reviewInfo = oOldObject.GetReviewInfo().Copy();
            oNewObject.SetReviewTypeWithInfo && oNewObject.SetReviewTypeWithInfo(reviewType, reviewInfo, false);
        }
    };
    CDocumentComparison.prototype.saveCustomReviewInfo = function(oObject, oOldObject, nType)
    {
        if (oOldObject.GetReviewInfo && oObject.SetReviewTypeWithInfo) {
            const oldReviewType = oOldObject.GetReviewType && oOldObject.GetReviewType();
            if (oldReviewType === reviewtype_Add || oldReviewType === reviewtype_Remove) {
                const reviewInfo = oOldObject.GetReviewInfo().Copy();
                oObject.SetReviewTypeWithInfo && oObject.SetReviewTypeWithInfo(oldReviewType, reviewInfo, false);
            } else {
                this.updateReviewInfo(oObject, nType);
            }
        }
    };
    CDocumentComparison.prototype.updateReviewInfo = function(oObject, nType)
    {
        if(oObject.ReviewInfo && oObject.SetReviewTypeWithInfo)
        {
            const oReviewInfo = oObject.ReviewInfo.Copy();
            this.setReviewInfo(oReviewInfo);
            oObject.SetReviewTypeWithInfo(nType, oReviewInfo, false);

        }
    };
    CDocumentComparison.prototype.getNodeConstructor = function () {
        return CNode;
    };

    CDocumentComparison.prototype.createNodeFromDocContent = function(oElement, oParentNode, oHashWords, isOriginalDocument)
    {
        const NodeConstructor = this.getNodeConstructor();
        const oRet = new NodeConstructor(oElement, oParentNode);
        const bRoot = (oParentNode === null);
        for(let i = 0; i < oElement.Content.length; ++i)
        {
            const oChElement = oElement.Content[i];
            if(oChElement instanceof Paragraph)
            {
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                }
                const oParagraphNode = this.createNodeFromRunContentElement(oChElement, oRet, oHashWords, isOriginalDocument);
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
                const oBlockNode = this.createNodeFromDocContent(oChElement.Content, oRet, oHashWords, isOriginalDocument);
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
                    const oTableNode = new NodeConstructor(oChElement, oRet);
                    if(bRoot)
                    {
                        oHashWords = new Minhash({});
                        oTableNode.hashWords = oHashWords;
                    }
                    for(let j = 0; j < oChElement.Content.length; ++j)
                    {
                        const oRowNode = new NodeConstructor(oChElement.Content[j], oTableNode);
                        for(let k = 0; k < oChElement.Content[j].Content.length; ++k)
                        {
                            this.createNodeFromDocContent(oChElement.Content[j].Content[k].Content, oRowNode, oHashWords, isOriginalDocument);
                        }
                    }
                }
            }
            else if (oChElement instanceof AscCommonWord.CMockParagraph) {
                if(bRoot)
                {
                    oHashWords = new AscCommonWord.CMockMinHash();
                }
                const oParagraphNode = this.createNodeFromRunContentElement(oChElement, oRet, oHashWords, isOriginalDocument);
                if(bRoot)
                {
                    oParagraphNode.hashWords = oHashWords;
                }
            }
            else
            {
                const oNode = new NodeConstructor(oChElement, oRet);
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                    oNode.hashWords = oHashWords;
                }
            }

        }
        return oRet;
    };

    CDocumentComparison.prototype.getTextElementConstructor = function () {
        return CTextElement;
    };

    function isBreakWordElement(oRunElement) {
        const bPunctuation = para_Text === oRunElement.Type && (AscCommon.g_aPunctuation[oRunElement.Value] && !EXCLUDED_PUNCTUATION[oRunElement.Value]);
        return (oRunElement.Type === para_Space || oRunElement.Type === para_Tab
            || oRunElement.Type === para_Separator || oRunElement.Type === para_NewLine
            || oRunElement.Type === para_FootnoteReference
            || oRunElement.Type === para_EndnoteReference
            || bPunctuation);
    }

    function createNodeFromRun(oRun, oLastText, oHashWords, oRet, TextElementConstructor, NodeConstructor, oReviewInfo) {
        if(oRun.Content.length > 0)
        {
            if(!oLastText)
            {
                oLastText = new TextElementConstructor();
                oLastText.setFirstRun(oRun);
            }
            if(oLastText.elements.length === 0)
            {
                oLastText.setFirstRun(oRun);
                oLastText.setLastRun(oRun);
            }
            for(let j = 0; j < oRun.Content.length; ++j)
            {
                const oRunElement = oRun.Content[j];
                if(isBreakWordElement(oRunElement))
                {
                    if(oLastText.elements.length > 0)
                    {
                        new NodeConstructor(oLastText, oRet);
                        oLastText.updateHash(oHashWords);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                    }

                    oLastText.setLastRun(oRun);
                    oLastText.addToElements(oRunElement, oReviewInfo);
                    new NodeConstructor(oLastText, oRet);
                    oLastText.updateHash(oHashWords);

                    oLastText = new TextElementConstructor();
                    oLastText.setFirstRun(oRun);
                    oLastText.setLastRun(oRun);
                }
                else if(oRunElement.Type === para_Drawing)
                {
                    if(oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oHashWords);
                        new NodeConstructor(oLastText, oRet);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                        oLastText.setLastRun(oRun);
                    }
                    oLastText.addToElements(oRun.Content[j], oReviewInfo);
                    new NodeConstructor(oLastText, oRet);
                    oLastText = new TextElementConstructor();
                    oLastText.setFirstRun(oRun);
                    oLastText.setLastRun(oRun);
                }
                else if (oRunElement.Type === para_RevisionMove) {
                    if(oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oHashWords);
                        new NodeConstructor(oLastText, oRet);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                        oLastText.setLastRun(oRun);
                    }
                }
                else if(oRunElement.Type === para_End)
                {
                    if(oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oHashWords);
                        new NodeConstructor(oLastText, oRet);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                        oLastText.setLastRun(oRun);
                    }
                    oLastText.setFirstRun(oRun);
                    oLastText.setLastRun(oRun);
                    // мы будем сравнивать ревью paraEnd отдельно, поскольку это единственный общий элемент в параграфе, до которого мы можем вставить любой различающийся контент
                    oLastText.addToElements(oRun.Content[j], {reviewType: reviewtype_Common, moveReviewType: Asc.c_oAscRevisionsMove.NoMove});
                    new NodeConstructor(oLastText, oRet);
                    oLastText.updateHash(oHashWords);
                    oLastText = new TextElementConstructor();
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
                    oLastText.addToElements(oRun.Content[j], oReviewInfo);
                }
            }
        }
        return oLastText;
    }

    CDocumentComparison.prototype.getCompareReviewInfo = function (oRun) {

    }

    CDocumentComparison.prototype.createNodeFromRun = function (oRun, oLastText, oHashWords, oRet) {
        const TextElementConstructor = this.getTextElementConstructor();
        const NodeConstructor = this.getNodeConstructor();
        const oReviewInfo = this.getCompareReviewInfo(oRun);
        return createNodeFromRun(oRun, oLastText, oHashWords, oRet, TextElementConstructor, NodeConstructor, oReviewInfo);
    }

    CDocumentComparison.prototype.createNodeFromRunContentElement = function(oElement, oParentNode, oHashWords, isOriginalDocument)
    {
        const NodeConstructor = this.getNodeConstructor();
        const TextElementConstructor = this.getTextElementConstructor();
        const oRet = new NodeConstructor(oElement, oParentNode);
        const aLastWord = [];
        let oLastText = null;
        for(let i = 0; i < oElement.Content.length; ++i)
        {
            const oRun = oElement.Content[i];
            if(oRun instanceof ParaRun)
            {
                oLastText = this.createNodeFromRun(oRun, oLastText, oHashWords, oRet);
            }
            else if (oRun instanceof AscCommon.CParaRevisionMove)
            {
                if(oLastText && oLastText.elements.length > 0)
                {
                    oLastText.updateHash(oHashWords);
                    new NodeConstructor(oLastText, oRet);
                }
                if(aLastWord.length > 0)
                {
                    oHashWords.update(aLastWord);
                    aLastWord.length = 0;
                }
                oLastText = null;
                new NodeConstructor(oRun, oRet);
                if (!isOriginalDocument && !this.bSkipChangeMoveType) {

                    this.oComparisonMoveMarkManager.changeRevisedMoveMarkId(oRun, this);
                    this.oComparisonMoveMarkManager.updateInsertMoveMarkId(oRun);
                }
                this.oComparisonMoveMarkManager.updateMoveMarksStack(oRun);

                oParentNode.bHaveMoveMarks = true;
            }
            else
            {
                if(!(oRun instanceof CParagraphBookmark))
                {
                    if(oLastText && oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oHashWords);
                        new NodeConstructor(oLastText, oRet);
                    }
                    if(aLastWord.length > 0)
                    {
                        oHashWords.update(aLastWord);
                        aLastWord.length = 0;
                    }
                    oLastText = null;
                    if(Array.isArray(oRun.Content))
                    {
                        this.createNodeFromRunContentElement(oRun, oRet, oHashWords, isOriginalDocument);
                    }
                    else
                    {
                        new NodeConstructor(oRun, oRet);
                    }
                }
            }
        }
        if(oLastText && oLastText.elements.length > 0)
        {
            oLastText.updateHash(oHashWords);
            new NodeConstructor(oLastText, oRet);
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
        let oComment = this.originalDocument.Comments.Get_ById(sId);
        if(oComment)
        {
            return oComment;
        }
        oComment = this.revisedDocument.Comments.Get_ById(sId);
        return oComment || null;
    };

    CDocumentComparison.prototype.checkCopyParagraphElement = function (oOldItem, oNewItem, arrMoveMarks) {
        return false;
    };


    window['AscCommonWord'] = window['AscCommonWord'] || {};
    window['AscCommonWord'].CDocumentComparison = CDocumentComparison;
    window['AscCommonWord'].ComparisonOptions = window['AscCommonWord']["ComparisonOptions"] = ComparisonOptions;

    function CompareBinary(oApi, sBinary2, oOptions, bForceApplyChanges)
    {
        const oDoc1 = oApi.WordControl.m_oLogicDocument;
        if(!window['NATIVE_EDITOR_ENJINE'])
        {
            const oCollaborativeEditing = oDoc1.CollaborativeEditing;
            if(oCollaborativeEditing && !oCollaborativeEditing.Is_SingleUser())
            {
                oApi.sendEvent("asc_onError", Asc.c_oAscError.ID.CannotCompareInCoEditing, c_oAscError.Level.NoCritical);
                return;
            }
        }
        oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
        let bHaveRevisons2 = false;
        const oDoc2 = AscFormat.ExecuteNoHistory(function(){
            const openParams = {disableRevisions: true, noSendComments: true};
            let oDoc2 = new CDocument(oApi.WordControl.m_oDrawingDocument, true);
            oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oDoc2;
            oApi.WordControl.m_oLogicDocument = oDoc2;
            const oBinaryFileReader = new AscCommonWord.BinaryFileReader(oDoc2, openParams);
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
            const fCallback = function()
            {
                const oComp = new AscCommonWord.CDocumentComparison(oDoc1, oDoc2, oOptions ? oOptions : new ComparisonOptions());
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

    function CMoveMarkComparisonManager() {
        this.oChangeRevisedMarkId = {};
        this.oChangeRevisedMarkIdRevert = {};
        this.oRevisedMoveMarks = {};
        this.oInsertMoveMarkId = {};
        this.oRevisedMoveMarksInserts = {};
        this.checkMoveMarks = [];
        this.oRunIdMoveNameRelation = {};
        this.oRevertMoveTypeByName = {};
        this.moveMarksStack = [];
        this.isResolveConflictMode = false;
    }

    CMoveMarkComparisonManager.prototype.changeRevisedMoveMarkId = function (oMoveMark, oComparison) {
        if (!this.oChangeRevisedMarkId[oMoveMark.Name]) {
            this.oChangeRevisedMarkId[oMoveMark.Name] = oComparison.originalDocument.TrackRevisionsManager.GetNewMoveId();
            this.oChangeRevisedMarkIdRevert[this.oChangeRevisedMarkId[oMoveMark.Name]] = oMoveMark.Name;
        }
        oMoveMark.Name = this.oChangeRevisedMarkId[oMoveMark.Name];
    };

    CMoveMarkComparisonManager.prototype.updateInsertMoveMarkId = function (oMoveMark) {
        this.oInsertMoveMarkId[oMoveMark.Name] = true;

        if (!this.oRevisedMoveMarks[oMoveMark.Name]) {
            this.oRevisedMoveMarks[oMoveMark.Name] = [];
        }
        this.oRevisedMoveMarks[oMoveMark.Name].push(oMoveMark);
    };

    CMoveMarkComparisonManager.prototype.updateMoveMarksStack = function (oMoveMark) {
        if (this.isResolveConflictMode) {
            return;
        }
        if (oMoveMark.Start) {
            this.moveMarksStack.push(oMoveMark);
        } else {
            this.moveMarksStack.pop();
        }
    };

    CMoveMarkComparisonManager.prototype.resetMoveMarkStack = function () {
        if (this.isResolveConflictMode) {
            return;
        }
        this.moveMarksStack = [];
    };

    CMoveMarkComparisonManager.prototype.addMoveMarkNameRunRelation = function (sMoveMarkName, oRun) {
        if (sMoveMarkName) {
            if (!this.oRevertMoveTypeByName[sMoveMarkName]) {
                this.oRevertMoveTypeByName[sMoveMarkName] = [];
            }
            this.oRevertMoveTypeByName[sMoveMarkName].push(oRun);
        }
    };

    CMoveMarkComparisonManager.prototype.addRunMoveMarkNameRelation = function (sMoveMarkName, oRun) {
        if (sMoveMarkName) {
            this.oRunIdMoveNameRelation[oRun.Id] = sMoveMarkName;
        }
    };

    CMoveMarkComparisonManager.prototype.getChangedMoveMarkName = function (oMoveMark) {
            return this.oChangeRevisedMarkId[oMoveMark.Name];
    };

    CMoveMarkComparisonManager.prototype.getMoveMarkNameByRun = function (oRun) {
        return this.oRunIdMoveNameRelation[oRun.Id]
    };

    CMoveMarkComparisonManager.prototype.checkMoveMarksContentNode = function (oNode) {
        if (oNode.bHaveMoveMarks) {
            this.checkMoveMarks.push(oNode.element);
        }
    };

    CMoveMarkComparisonManager.prototype.getCheckMoveMarkElements = function () {
        return this.checkMoveMarks;
    };

    CMoveMarkComparisonManager.prototype.addRevisedMoveMarkToInserts = function (oMoveMark, oFrontElement, oParentElement, bIsParaEnd) {
        if (!this.oRevisedMoveMarksInserts[oFrontElement.Id]) {
            this.oRevisedMoveMarksInserts[oFrontElement.Id] = [];
        }
        this.oRevisedMoveMarksInserts[oFrontElement.Id].push({isParaEnd: bIsParaEnd, frontElement: oFrontElement, parentElement: oParentElement, moveMark: oMoveMark});
    };

    CMoveMarkComparisonManager.prototype.getRevisedMoveMarkToInserts = function () {
        return this.oRevisedMoveMarksInserts;
    };

    CMoveMarkComparisonManager.prototype.executeResolveConflictMode = function (callback) {
        const bOldModeValue = this.isResolveConflictMode;
        this.isResolveConflictMode = true;
        callback();
        this.isResolveConflictMode = bOldModeValue;
    };

    window['AscCommonWord']["CompareBinary"] =  window['AscCommonWord'].CompareBinary = CompareBinary;
    window['AscCommonWord']["ComparisonOptions"] = window['AscCommonWord'].ComparisonOptions = ComparisonOptions;
    window['AscCommonWord']['CompareDocuments'] = CompareDocuments;
    window['AscCommonWord'].CDocumentComparison = CDocumentComparison;
    window['AscCommonWord'].CNode = CNode;
    window['AscCommonWord'].CTextElement = CTextElement;
    window['AscCommonWord'].isBreakWordElement = isBreakWordElement;
})();
