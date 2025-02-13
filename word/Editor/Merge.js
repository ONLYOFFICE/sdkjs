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

(function (undefined) {
    const CDocumentComparison = AscCommonWord.CDocumentComparison;
    const CNode = AscCommonWord.CNode;
    const CTextElement = AscCommonWord.CTextElement;

    function getPriorityReviewType(arrOfTypes) {
        const bRemove = arrOfTypes.some(function (reviewType) {
            return reviewType === reviewtype_Remove;
        });
        if (bRemove) return reviewtype_Remove;

        const bAdd = arrOfTypes.some(function (reviewType) {
            return reviewType === reviewtype_Add;
        });
        if (bAdd) return reviewtype_Add;
        return reviewtype_Common;
    }

    function CMergeComparisonNode(oElement, oParent) {
        CNode.call(this, oElement, oParent);
        this.bHaveMoveMarks = false;
    }

    CMergeComparisonNode.prototype = Object.create(CNode.prototype);
    CMergeComparisonNode.prototype.constructor = CMergeComparisonNode;
	CMergeComparisonNode.prototype.applyInsertsToParagraphsWithoutRemove = function (comparison, aContentToInsert, oChange) {
		if (this.checkLastNodeOnInsertWithResolveConflict(comparison, aContentToInsert, oChange)) {
			CNode.prototype.applyInsertsToParagraphsWithoutRemove.call(this, comparison, aContentToInsert, oChange);
		}
	};
		CMergeComparisonNode.prototype.applyInsertsToParagraphsWithRemove = function(comparison, aContentToInsert, oChange) {
			const arrSetRemoveReviewType = [];
			const infoAboutEndOfRemoveChange = this.prepareEndOfRemoveChange(oChange, comparison, arrSetRemoveReviewType);
			const posLastRunInContent = infoAboutEndOfRemoveChange.posLastRunInContent;

			let nInsertPosition = infoAboutEndOfRemoveChange.nInsertPosition;
			nInsertPosition = this.setReviewTypeForRemoveChanges(comparison, oChange, posLastRunInContent, nInsertPosition, arrSetRemoveReviewType);

			this.applyInsert(aContentToInsert, arrSetRemoveReviewType, nInsertPosition, comparison, {needReverse: true});
		};
    CMergeComparisonNode.prototype.privateCompareElements = function (oNode, bCheckNeighbors, oComparison) {
        const oElement1 = this.element;
        const oElement2 = oNode.element;
        if (oElement1.isReviewWord !== oElement2.isReviewWord) {
            return false;
        }
        return CNode.prototype.privateCompareElements.call(this, oNode, bCheckNeighbors, oComparison);
    }

    CMergeComparisonNode.prototype.setCommonReviewTypeWithInfo = function (element, info) {
        element.SetReviewTypeWithInfo((element.GetReviewType && element.GetReviewType()) || reviewtype_Common, info);
    };

    CMergeComparisonNode.prototype.applyInsert = function (arrToInsert, arrToRemove, nInsertPosition, comparison, opts) {
        opts = opts || {};
        if (arrToInsert.length === 0) {
            for (let i = 0; i < arrToRemove.length; i += 1) {
                comparison.setRemoveReviewType(arrToRemove[i]);
            }
        } else if (arrToRemove.length === 0) {
            this.insertContentAfterRemoveChanges(arrToInsert, nInsertPosition, comparison);
        } else {
            arrToInsert = arrToInsert.reverse();
            if (opts.needReverse) {
                arrToRemove = arrToRemove.reverse();
            }
            comparison.resolveConflicts(arrToInsert, arrToRemove, this.getApplyParagraph(comparison));
        }
    }

    CMergeComparisonNode.prototype.insertContentAfterRemoveChanges = function (aContentToInsert, nInsertPosition, comparison, options) {
			const oElement = this.getApplyParagraph(comparison);
	    let t = 0;
			const arrSkippedComments = [];
			if (options && AscFormat.isRealNumber(options.nCommentInsertIndex))
			{
				for (t;t < aContentToInsert.length; t += 1)
				{
					if (aContentToInsert[t] instanceof AscCommon.ParaComment)
					{
						const sCommentId = aContentToInsert[t].GetCommentId();
						if (comparison.oCommentManager.mapLink[sCommentId])
						{
							oElement.AddToContent(options.nCommentInsertIndex, aContentToInsert[t]);
						}
						else
						{
							arrSkippedComments.push(aContentToInsert[t]);
						}
					}
					else
					{
						break;
					}
				}
			}
        if(nInsertPosition > -1)
        {
	        for (let i = 0; i < arrSkippedComments.length; i++)
	        {
		        oElement.AddToContent(nInsertPosition, arrSkippedComments[i]);
	        }
            for (t; t < aContentToInsert.length; t += 1) {
                if(comparison.isElementForAdd(aContentToInsert[t]))
                {
                    if (aContentToInsert[t] instanceof AscCommon.CParaRevisionMove) {
                        const oInsertBefore = oElement.Content[nInsertPosition];
                        if (oInsertBefore) {
                            comparison.oComparisonMoveMarkManager.addRevisedMoveMarkToInserts(aContentToInsert[t], oInsertBefore, oElement, false);
                        }
                    } else {
                        const sMoveName = comparison.oComparisonMoveMarkManager.getMoveMarkNameByRun(aContentToInsert[t]);
                        comparison.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sMoveName, aContentToInsert[t]);
                        oElement.AddToContent(nInsertPosition, aContentToInsert[t]);
                    }

                }
            }
        }
    };

    function CMergeComparisonTextElement() {
        CTextElement.call(this);
        this.isReviewWord = false;
        this.reviewElementTypes = [];
    }

    CMergeComparisonTextElement.prototype = Object.create(CTextElement.prototype);
    CMergeComparisonTextElement.prototype.constructor = CMergeComparisonTextElement;

    CMergeComparisonTextElement.prototype.addToElements = function (element, reviewType) {
        CTextElement.prototype.addToElements.call(this, element);
        this.reviewElementTypes.push(reviewType);
        if (reviewType.reviewType !== reviewtype_Common || reviewType.moveReviewType !== Asc.c_oAscRevisionsMove.NoMove) {
            this.isReviewWord = true;
        }
    };

	function compareReviewElements(reviewElement1, reviewElement2)
	{
		return (
			reviewElement1.reviewType === reviewtype_Common && reviewElement1.reviewType === reviewElement2.reviewType &&
			reviewElement1.moveReviewType === Asc.c_oAscRevisionsMove.NoMove && reviewElement1.moveReviewType === reviewElement2.moveReviewType ||
			reviewElement1.moveReviewType === reviewElement2.moveReviewType &&
			reviewElement1.reviewType === reviewElement2.reviewType &&
			reviewElement1.reviewInfo && reviewElement2.reviewInfo && reviewElement1.reviewInfo.IsEqual(reviewElement2.reviewInfo, true)
		);
	}
	CMergeComparisonTextElement.prototype.compareReviewElements = function (oAnotherElement)
	{
		if (this.reviewElementTypes.length === oAnotherElement.reviewElementTypes.length) {
			for (let i = 0; i < this.reviewElementTypes.length; i += 1) {
				const bNotEqualsReviewTypes = !compareReviewElements(this.reviewElementTypes[i], oAnotherElement.reviewElementTypes[i]);
				if (bNotEqualsReviewTypes) {
					return false;
				}
			}
		}
		else
		{
			return false;
		}
		return true;
	};
    CMergeComparisonTextElement.prototype.equals = function (oOtherElement, bNeedCheckTypes, oComparison) {
        const bEquals = CTextElement.prototype._equals.call(this, oOtherElement, bNeedCheckTypes, oComparison);
        if (!bEquals) {
            return false;
        }
				if (bNeedCheckTypes)
				{
					return this.compareReviewElements(oOtherElement);
				}

        return true;
    };

    function CResolveConflictTextElement() {
        CTextElement.call(this);
        this.reviewElementTypes = [];
    }

    CResolveConflictTextElement.prototype = Object.create(CTextElement.prototype);
    CResolveConflictTextElement.prototype.constructor = CResolveConflictTextElement;
    CResolveConflictTextElement.prototype.addToElements = CMergeComparisonTextElement.prototype.addToElements;

    CResolveConflictTextElement.prototype.checkRemoveReviewType = function (nIndex)
    {
        const oReviewInformation = this.reviewElementTypes[nIndex];
        if (oReviewInformation)
        {
            return oReviewInformation.reviewType === reviewtype_Remove;
        }
        return false;
    };

    function CDocumentResolveConflictComparison(oOriginalDocument, oRevisedDocument, oOptions, bIsWordsByOneSymbol) {
        CDocumentComparison.call(this, oOriginalDocument, oRevisedDocument, oOptions, bIsWordsByOneSymbol);
	    this.needCopyForResolveEqualWords = false;
        this.parentParagraph = null;
        this.bSkipChangeMoveType = true;
	      this.needCheckReview = true;
        this.copyPr = {
            CopyReviewPr: false,
            Comparison: this,
	        SkipFootnoteReference: !oOptions.footNotes
        };
        this.bSaveCustomReviewType = true;
    }
    CDocumentResolveConflictComparison.prototype = Object.create(CDocumentComparison.prototype);
    CDocumentResolveConflictComparison.prototype.constructor = CDocumentResolveConflictComparison;
	CDocumentResolveConflictComparison.prototype.executeWithCheckInsertAndRemove = function (callback, oChange) {
		callback();
	};
	CDocumentResolveConflictComparison.prototype.removeCommentsFromMap = function ()
	{

	};
    CDocumentResolveConflictComparison.prototype.getNodeConstructor = function () {
        return CConflictResolveNode;
    }
    CDocumentResolveConflictComparison.prototype.checkOriginalAndSplitRun = function (oOriginalRun, oSplitRun) {
        const sName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOriginalRun);
        this.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sName, oSplitRun);
        this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sName, oSplitRun);
    };

    CDocumentResolveConflictComparison.prototype.setReviewInfoForArray = function (arrNeedReviewObjects, nType) {
        for (let i = 0; i < arrNeedReviewObjects.length; i += 1) {
            const oNeedReviewObject = arrNeedReviewObjects[i];
            if (oNeedReviewObject.SetReviewTypeWithInfo) {
                let oReviewInfo = this.getReviewInfo();
                if (this.bSaveCustomReviewType) {
                    const reviewType = oNeedReviewObject.GetReviewType && oNeedReviewObject.GetReviewType();
                    if (reviewType === reviewtype_Add || reviewType === reviewtype_Remove) {
                        if (nType === reviewtype_Add && reviewType === reviewtype_Remove) {
                            oReviewInfo = oNeedReviewObject.ReviewInfo.Copy();
                            oReviewInfo.SavePrev(reviewtype_Add);
                            nType = reviewtype_Remove;
                        } else if (reviewType === reviewtype_Add && nType === reviewtype_Remove) {
                            oReviewInfo = oNeedReviewObject.ReviewInfo.Copy();
                            oReviewInfo.SavePrev(reviewtype_Add);
                        }
                    }
                }
                oNeedReviewObject.SetReviewTypeWithInfo(nType, oReviewInfo, false);
            }
        }
    };

    CDocumentResolveConflictComparison.prototype.getTextElementConstructor = function () {
        return CResolveConflictTextElement;
    };

    CDocumentResolveConflictComparison.prototype.getCompareReviewInfo = function (oRun) {
        let oReviewInfo = oRun.GetReviewInfo && oRun.GetReviewInfo();
        const prevAdded = oReviewInfo && oReviewInfo.GetPrevAdded();
        const reviewType = oRun.GetReviewType && oRun.GetReviewType();
        const moveReviewType = oRun.GetReviewMoveType && oRun.GetReviewMoveType();
        const bNotRunIdInit = !this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oRun);
        if (moveReviewType !== Asc.c_oAscRevisionsMove.NoMove && bNotRunIdInit && this.oComparisonMoveMarkManager.moveMarksStack.length) {
            const oCurrentMoveMark = this.oComparisonMoveMarkManager.moveMarksStack[this.oComparisonMoveMarkManager.moveMarksStack.length - 1];
            this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(oCurrentMoveMark.Name, oRun);
        }
        return {
            reviewType: reviewType,
            moveReviewType: moveReviewType,
            moveName: this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oRun),
            reviewInfo: oReviewInfo,
            prevAdded: prevAdded
        };
    };

    CDocumentResolveConflictComparison.prototype.applyChangesToParagraph = function (oNode) {
        oNode.changes.sort(function (c1, c2) {
            return c2.anchor.index - c1.anchor.index;
        });
        let currentChangeId = 0;
        for (let i = oNode.children.length - 1; i >= 0; i -= 1) {
            const oChildNode = oNode.children[i];
            if (i !== oNode.children.length - 1) {
							if (oChildNode.partner && oChildNode.element instanceof CTextElement) {
									oChildNode.updateEqualNode(this, i);
							}
            }
						const oChange = oNode.changes[currentChangeId];
            if (oChange && oChange.anchor.index === i) {
							this.copyPr.SkipUpdateInfo = !this.options.words && !this.isWordsByOneSymbol && oChange.insert.length === 1 && oChange.remove.length === 1;
                const aContentToInsert = this.isSkipWhitespaces(oChange.insert) ? [] : oNode.getArrOfInsertsFromChanges(oChange, this); // todo: check skip on symbol comparing
	            this.copyPr.SkipUpdateInfo = false;
                //handle removed elements
                oNode.applyInsertsToParagraph(this, aContentToInsert, oChange);
                currentChangeId += 1
            }
            if (oChange && oChange.anchor.index > i) {
                currentChangeId += 1;
            }
        }

        this.applyChangesToChildrenOfParagraphNode(oNode);
        this.applyChangesToSectPr(oNode);
    };

    CDocumentResolveConflictComparison.prototype.setRemoveReviewType = function (element) {
        if (!(element.IsParaEndRun && element.IsParaEndRun())) {
            if (!element.GetReviewType || element.GetReviewType && element.GetReviewType() === reviewtype_Common) {
                this.setReviewInfoRecursive(element, this.nRemoveChangesType);
            }
        }
    };

    CDocumentResolveConflictComparison.prototype.resolveCustomReviewTypesBetweenElements = function (oMainElement, nRevisedReviewType, oRevisedReviewInfo) {
        const nMainReviewType = oMainElement.GetReviewType();
        if (nRevisedReviewType !== reviewtype_Common && nRevisedReviewType !== nMainReviewType) {
            let oMainReviewInfo = oMainElement.GetReviewInfo();
			if (oMainReviewInfo)
				oMainReviewInfo = oMainReviewInfo.Copy();
			
			if (oRevisedReviewInfo)
				oRevisedReviewInfo = oRevisedReviewInfo.Copy();
			
            if (nMainReviewType === reviewtype_Common) {
                oMainElement.SetReviewTypeWithInfo(nRevisedReviewType, oRevisedReviewInfo);
            } else if (nMainReviewType === reviewtype_Add) {
                oRevisedReviewInfo.SetPrevReviewTypeWithInfoRecursively(reviewtype_Add, oMainReviewInfo);
                oMainElement.SetReviewTypeWithInfo(reviewtype_Remove, oRevisedReviewInfo);
            } else if (nMainReviewType === reviewtype_Remove) {
                oMainReviewInfo.SetPrevReviewTypeWithInfoRecursively(reviewtype_Add, oRevisedReviewInfo);
                oMainElement.SetReviewTypeWithInfo(reviewtype_Remove, oMainReviewInfo);
            }
        }
    };
	CDocumentResolveConflictComparison.prototype.applyResolveTypes = function (oNeedReviewWithUser)
	{
		for (let sReviewDate in oNeedReviewWithUser)
		{
			for (let sUserName in oNeedReviewWithUser[sReviewDate])
			{
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Add].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Add][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					this.resolveCustomReviewTypesBetweenElements(element, reviewtype_Add, reviewInfo);

				}
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Remove].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].reviewTypes[reviewtype_Remove][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					this.resolveCustomReviewTypesBetweenElements(element, reviewtype_Remove, reviewInfo);
				}
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveFrom].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveFrom][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					const nOldReviewType = element.GetReviewType();
					if (nOldReviewType !== reviewtype_Common)
					{
						const oOldReviewInfo = element.GetReviewInfo().Copy();
						reviewInfo.SetPrevReviewTypeWithInfoRecursively(nOldReviewType, oOldReviewInfo);
					}
					element.SetReviewTypeWithInfo(reviewtype_Remove, reviewInfo);
				}
				for (let i = 0; i < oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveTo].length; i += 1)
				{
					const info = oNeedReviewWithUser[sReviewDate][sUserName].moveReviewTypes[Asc.c_oAscRevisionsMove.MoveTo][i];
					const element = info.element;
					const reviewInfo = info.reviewInfo;
					const nOldReviewType = element.GetReviewType();
					if (nOldReviewType !== reviewtype_Common)
					{
						const oOldReviewInfo = element.GetReviewInfo().Copy();
						reviewInfo.SetPrevReviewTypeWithInfoRecursively(nOldReviewType, oOldReviewInfo);
					}
					element.SetReviewTypeWithInfo(reviewtype_Add, reviewInfo);
				}
			}
		}
	};

    function CConflictResolveNode(oElement, oParent) {
        CNode.call(this, oElement, oParent);
    }

    CConflictResolveNode.prototype = Object.create(CNode.prototype);
    CConflictResolveNode.prototype.constructor = CConflictResolveNode;

    CConflictResolveNode.prototype.applyInsertsToParagraphsWithRemove = function (comparison, aContentToInsert, oChange) {
        const arrSetRemoveReviewType = [];
        const infoAboutEndOfRemoveChange = this.prepareEndOfRemoveChange(oChange, comparison, arrSetRemoveReviewType);
        const posLastRunInContent = infoAboutEndOfRemoveChange.posLastRunInContent;

        const nInsertPosition = infoAboutEndOfRemoveChange.nInsertPosition;
        this.setReviewTypeForRemoveChanges(comparison, oChange, posLastRunInContent, nInsertPosition, arrSetRemoveReviewType);

        const nInsertPosition2 = arrSetRemoveReviewType[arrSetRemoveReviewType.length - 1].GetPosInParent();
				if (!comparison.options.words && !comparison.isWordsByOneSymbol && oChange.insert.length === 1 && oChange.remove.length === 1) {
					aContentToInsert.reverse();
					arrSetRemoveReviewType.reverse();
					comparison.resolveConflicts(aContentToInsert, arrSetRemoveReviewType, this.getApplyParagraph(comparison), true);
				} else {
					this.applyInsert(aContentToInsert, arrSetRemoveReviewType, nInsertPosition2, comparison, {needReverse: true, nCommentInsertIndex: nInsertPosition});
				}
    };

    CConflictResolveNode.prototype.applyInsertsToParagraphsWithoutRemove = function (comparison, aContentToInsert, oChange) {
			if (this.checkLastNodeOnInsertWithResolveConflict(comparison, aContentToInsert, oChange)) {
				CNode.prototype.applyInsertsToParagraphsWithoutRemove.call(this, comparison, aContentToInsert, oChange);
			}
    };
    CConflictResolveNode.prototype.insertContentAfterRemoveChanges = CMergeComparisonNode.prototype.insertContentAfterRemoveChanges;

    CConflictResolveNode.prototype.copyRunWithMockParagraph = function (oRun, mockParagraph, comparison) {
        comparison.copyPr.bSaveCustomReviewType = true;
        const oRet = CNode.prototype.copyRunWithMockParagraph.call(this, oRun, mockParagraph, comparison);
        delete comparison.copyPr.bSaveCustomReviewType;
        return oRet;
    };

    CConflictResolveNode.prototype.pushToArrInsertContentWithCopy = function (aContentToInsert, elem, comparison) {
            comparison.copyPr.bSaveCustomReviewType = true;
            CNode.prototype.pushToArrInsertContentWithCopy.call(this, aContentToInsert, elem, comparison);
            delete comparison.copyPr.bSaveCustomReviewType;
    };

    CConflictResolveNode.prototype.setCommonReviewTypeWithInfo = function (element, info) {
        element.SetReviewTypeWithInfo((element.GetReviewType && element.GetReviewType()) || reviewtype_Common, info);
    };

    function CMockDocument() {
        this.Content = [];
    }

    function CMockParagraph() {
        this.Content = [];
    }
	CMockParagraph.prototype.Add_ToContent = function (position, element) {
		this.Content.splice(position, 0, element);
	};

    function CDocumentMergeComparison(oOriginalDocument, oRevisedDocument, oOptions, bIsWordsByOneSymbol) {
        CDocumentComparison.call(this, oOriginalDocument, oRevisedDocument, oOptions, bIsWordsByOneSymbol);
        this.bSaveCustomReviewType = true;
        this.copyPr = {
            CopyReviewPr: false,
            Comparison: this,
            SkipUpdateInfo: true,
            CheckComparisonMoveMarks: true,
	        SkipFootnoteReference: !oOptions.footNotes
        };
    }

    CDocumentMergeComparison.prototype = Object.create(CDocumentComparison.prototype);
    CDocumentMergeComparison.prototype.constructor = CDocumentMergeComparison;

	CDocumentMergeComparison.prototype.insertCopyTextBoxContent = function (oBaseShape, oTextBoxContent) {
		this.executeDisableSkipUpdateInfo(function () {
			CDocumentComparison.prototype.insertCopyTextBoxContent.call(this, oBaseShape, oTextBoxContent);
		}.bind(this));
	};
	CDocumentMergeComparison.prototype.executeDisableSkipUpdateInfo = function (callback) {
		const bOldSkipUpdateInfo = this.copyPr.SkipUpdateInfo;
		const bSaveCustomReviewType = this.copyPr.bSaveCustomReviewType;
		this.copyPr.SkipUpdateInfo = false;
		this.copyPr.bSaveCustomReviewType = true;
		callback();
		this.copyPr.SkipUpdateInfo = bOldSkipUpdateInfo;
		this.copyPr.bSaveCustomReviewType = bSaveCustomReviewType;
	};

    CDocumentMergeComparison.prototype.executeWithCheckInsertAndRemove = function (callback, oChange) {
        if (!oChange.remove.length || !oChange.insert.length) {
					this.executeDisableSkipUpdateInfo(callback);
        } else {
					callback();
        }
    };

    CDocumentMergeComparison.prototype.checkOriginalAndSplitRun = CDocumentResolveConflictComparison.prototype.checkOriginalAndSplitRun;

    CDocumentMergeComparison.prototype.createNodeFromDocContent = function (oElement, oParentNode, oHashWords, isOriginalDocument) {
        this.oComparisonMoveMarkManager.resetMoveMarkStack();
        const oRet = CDocumentComparison.prototype.createNodeFromDocContent.call(this, oElement, oParentNode, oHashWords, isOriginalDocument);
        this.oComparisonMoveMarkManager.checkMoveMarksContentNode(oRet);
        return oRet;
    };
    CDocumentMergeComparison.prototype.checkCopyParagraphElement = function (oOldItem, oNewItem, arrMoveMarks) {
        if (para_RevisionMove === oOldItem.Type) {
            arrMoveMarks.unshift({moveMark: oNewItem, parentElement: this});
            return true;
        } else if (Array.isArray(oNewItem)) {
            const sMoveName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOldItem[0]);
            this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveName, oNewItem[0]);
            this.oComparisonMoveMarkManager.oRevisedMoveMarksInserts[oNewItem[0].Id] = arrMoveMarks;
        } else {
            const sMoveName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOldItem);
            this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveName, oNewItem);
            this.oComparisonMoveMarkManager.oRevisedMoveMarksInserts[oNewItem.Id] = arrMoveMarks;
        }
        return false;
    };
    CDocumentMergeComparison.prototype.correctMoveMarks = function () {
        const oRemoveMoveTypeNames = {};
        const oTrackRevisionManager = this.api.WordControl.m_oLogicDocument.TrackRevisionsManager;
        const oComparisonMoveMarkManager = this.oComparisonMoveMarkManager;
        const oMoveMarks = oTrackRevisionManager.MoveMarks;
        const oRevisedMoveMarksInserts = oComparisonMoveMarkManager.getRevisedMoveMarkToInserts();
        const oInsertMoveMarkId = oComparisonMoveMarkManager.oInsertMoveMarkId;
        const arrCheckMoveMarksElements = oComparisonMoveMarkManager.getCheckMoveMarkElements();

        for (let i = 0; i < arrCheckMoveMarksElements.length; i += 1) {
            const oRootElement = arrCheckMoveMarksElements[i];
            let nStartEndOriginalCounter = 0;
            let oStartMove = {};
            function checkOriginalMoveMark(oElement) {
                if (oMoveMarks[oElement.Name]) {
                    const oMoveMarkInfo = oMoveMarks[oElement.Name];
                    if (oMoveMarkInfo.From.Start === oElement || oMoveMarkInfo.To.Start === oElement) {
                        nStartEndOriginalCounter += 1;
                    } else if (oMoveMarkInfo.From.End === oElement || oMoveMarkInfo.To.End === oElement) {
                        nStartEndOriginalCounter -= 1;
                    }
                    for (let sName in oStartMove) {
                        oInsertMoveMarkId[sName] = false;
                        oRemoveMoveTypeNames[sName] = true;
                    }
                    oStartMove = {};
                }
            }

            function checkRevisedMoveMark(oElement) {
                if (oRevisedMoveMarksInserts[oElement.Id]) {
                    const arrMoveMarks = oRevisedMoveMarksInserts[oElement.Id];
                    for (let t = 0; t < arrMoveMarks.length; t += 1) {
                        const oRevisedInsertInfo = arrMoveMarks[t];
                        if (oRevisedInsertInfo.moveMark.Start) {
                            oStartMove[oRevisedInsertInfo.moveMark.Name] = true;
                        } else {
                            if (!nStartEndOriginalCounter) {
                                delete oStartMove[oRevisedInsertInfo.moveMark.Name];
                            }
                        }
                    }
                }
            }

            for (let j = 0; j < oRootElement.Content.length; j += 1) {
                const oChildElement = oRootElement.Content[j];
                if (oChildElement instanceof Paragraph) {
                    const arrContent = oChildElement.Content;
                    for (let k = 0; k < arrContent.length - 1; k += 1) {
                        const oElement = arrContent[k];
                        if (oElement instanceof AscCommon.CParaRevisionMove) {
                            checkOriginalMoveMark(oElement);
                        } else {
                            checkRevisedMoveMark(oElement);
                        }
                    }
                    const oParaEnd = oChildElement.GetParaEndRun();
                    checkRevisedMoveMark(oParaEnd);
                    const oLastMoveMark = oParaEnd.GetLastTrackMoveMark();
                    if (oLastMoveMark) {
                        checkOriginalMoveMark(oLastMoveMark);
                    }
                }
            }
        }
        for (let sName in oRemoveMoveTypeNames) {
            const oRevertMoveTypeByName = oComparisonMoveMarkManager.oRevertMoveTypeByName;
            if (oRevertMoveTypeByName[sName]) {
                for (let i = 0; i < oRevertMoveTypeByName[sName].length; i += 1) {
                    oRevertMoveTypeByName[sName][i].RemoveReviewMoveType();
                }
            }
        }
        for (let sId in oRevisedMoveMarksInserts) {
            const arrInsert = oRevisedMoveMarksInserts[sId];
            for (let i = arrInsert.length - 1; i >= 0; i -= 1) {
                const oInsertInfo = arrInsert[i];
                const oInsertParaMove = oInsertInfo.moveMark;
                if (oInsertMoveMarkId[oInsertParaMove.Name]) {
                    const oRun = AscCommon.g_oTableId.Get_ById(sId);
                    if (oInsertInfo.isParaEnd) {
                        oRun.AddAfterParaEnd(oInsertParaMove);
                        } else {
                        const oParent = oRun.GetParent();
                        const nPosition = oRun.GetPosInParent(oParent);
                        oParent.AddToContent(nPosition, oInsertParaMove);
                    }
                }
            }
        }
    };

    CDocumentMergeComparison.prototype.setRemoveReviewType = function (element) {
        if (!(element.IsParaEndRun && element.IsParaEndRun())) {
            if (!element.GetReviewType || element.GetReviewType && element.GetReviewType() === reviewtype_Common) {
                this.setReviewInfoRecursive(element, this.nRemoveChangesType);
            }
        }
    };
    CDocumentMergeComparison.prototype.resolveCustomReviewTypesBetweenElements = CDocumentResolveConflictComparison.prototype.resolveCustomReviewTypesBetweenElements;

    CDocumentMergeComparison.prototype.checkParaEndReview = function (oNode) {
        if (oNode && oNode.element.GetType && oNode.element.GetType() === type_Paragraph && oNode.partner) {
            const oMainParaEnd = oNode.element.GetParaEndRun();
            const oRevisedParaEnd = oNode.partner.element.GetParaEndRun();
            const nRevisedReviewType = oRevisedParaEnd.GetReviewType();
            const oRevisedReviewInfo = oRevisedParaEnd.GetReviewInfo();
            const nOldMainMoveReviewType = oMainParaEnd.GetReviewMoveType();
            this.resolveCustomReviewTypesBetweenElements(oMainParaEnd, nRevisedReviewType, oRevisedReviewInfo);
            if (nOldMainMoveReviewType === Asc.c_oAscRevisionsMove.NoMove && oMainParaEnd.GetReviewMoveType() !== Asc.c_oAscRevisionsMove.NoMove) {
                const oRunMoveMark = oRevisedParaEnd.GetLastTrackMoveMark();
                if (oRunMoveMark) {
                    const oCopyMoveMark = oRunMoveMark.Copy(this.copyPr);
                    const sChangedMoveMarkName = this.oComparisonMoveMarkManager.getChangedMoveMarkName(oCopyMoveMark);
                    oCopyMoveMark.Name = sChangedMoveMarkName;
                    this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sChangedMoveMarkName, oMainParaEnd);
                    this.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sChangedMoveMarkName, oMainParaEnd);
                    this.oComparisonMoveMarkManager.addRevisedMoveMarkToInserts(oCopyMoveMark, oMainParaEnd, oNode.element, true);

                }
            }
        }
    };

    CDocumentMergeComparison.prototype.applyChangesToTableSize = function(oNode) {
	    this.executeDisableSkipUpdateInfo(function () {
		    CDocumentComparison.prototype.applyChangesToTableSize.call(this, oNode);
	    }.bind(this));
    };

    CDocumentMergeComparison.prototype.checkRowReview = function(oRowNode) {
        const oPartnerNode = oRowNode.partner;
        if (oPartnerNode) {
            const oMainRow = oRowNode.element;
            const oPartnerRow = oPartnerNode.element;
            const nRevisedReviewType = oPartnerRow.GetReviewType();
            const oRevisedReviewInfo = oPartnerRow.GetReviewInfo();
            this.resolveCustomReviewTypesBetweenElements(oMainRow, nRevisedReviewType, oRevisedReviewInfo);
        }
    };
    CDocumentMergeComparison.prototype.getCompareReviewInfo = CDocumentResolveConflictComparison.prototype.getCompareReviewInfo;

    CDocumentMergeComparison.prototype.applyParagraphComparison = function (oOrigRoot, oRevisedRoot) {
				this.executeDisableSkipUpdateInfo(function () {
					CDocumentComparison.prototype.applyParagraphComparison.call(this, oOrigRoot, oRevisedRoot);
					for (let i = oOrigRoot.children.length - 1; i >= 0; i -= 1) {
						this.checkParaEndReview(oOrigRoot.children[i]);
					}
				}.bind(this));
    };

    CDocumentMergeComparison.prototype.getNodeConstructor = function () {
        return CMergeComparisonNode;
    };


    CDocumentMergeComparison.prototype.getTextElementConstructor = function () {
        return CMergeComparisonTextElement;
    };

    CDocumentMergeComparison.prototype.GetReviewTypeFromParaDrawing = function (oParaDrawing) {
        const oRun = oParaDrawing.GetRun();
        if (oRun) {
            return oRun.GetReviewType();
        }
        return reviewtype_Common;
    };

    CDocumentMergeComparison.prototype.compareDrawingObjects = function (oBaseDrawing, oCompareDrawing, bOrig) {
        if (oBaseDrawing && oCompareDrawing) {
            const baseReviewType = this.GetReviewTypeFromParaDrawing(oBaseDrawing);
            const compareReviewType = this.GetReviewTypeFromParaDrawing(oCompareDrawing);
            const arrOfReviewTypes = [];

            if (baseReviewType) arrOfReviewTypes.push(baseReviewType);
            if (compareReviewType) arrOfReviewTypes.push(compareReviewType);

            const priorityReviewType = getPriorityReviewType(arrOfReviewTypes);

            const oBaseRun = bOrig ? oBaseDrawing.GetRun() : oCompareDrawing.GetRun();
            this.setReviewInfoForArray([oBaseRun], priorityReviewType);
        }
        CDocumentComparison.prototype.compareDrawingObjects.call(this, oBaseDrawing, oCompareDrawing);
    };



    CDocumentMergeComparison.prototype.compare = function (callback) {
        const oOriginalDocument = this.originalDocument;
        const oRevisedDocument = this.revisedDocument;
        if (!oOriginalDocument || !oRevisedDocument) {
            return;
        }
	    this.oBookmarkManager.init(oOriginalDocument, oRevisedDocument);
        const oThis = this;
        const aImages = AscCommon.pptx_content_loader.End_UseFullUrl();
        const oObjectsForDownload = AscCommon.GetObjectsForImageDownload(aImages);
        const oApi = oOriginalDocument.GetApi();
        if (!oApi) {
            return;
        }
        const fCallback = function (data) {
            const oImageMap = {};
            AscFormat.ExecuteNoHistory(function () {
                AscCommon.ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
            }, oThis, []);

            const NewNumbering = oRevisedDocument.Numbering.CopyAllNums(oOriginalDocument.Numbering);
            oRevisedDocument.CopyNumberingMap = NewNumbering.NumMap;
            oOriginalDocument.Numbering.AppendAbstractNums(NewNumbering.AbstractNum);
            oOriginalDocument.Numbering.AppendNums(NewNumbering.Num);
            for (let key in NewNumbering.NumMap) {
                if (NewNumbering.NumMap.hasOwnProperty(key)) {
                    oThis.checkedNums[NewNumbering.NumMap[key]] = true;
                }
            }
            oThis.compareRoots(oOriginalDocument, oRevisedDocument);
            oThis.compareSectPr(oOriginalDocument, oRevisedDocument, !oThis.options.headersAndFooters);

            const oFonts = oOriginalDocument.Document_Get_AllFontNames();
            const aFonts = [];
            for (let i in oFonts) {
                if (oFonts.hasOwnProperty(i)) {
                    aFonts[aFonts.length] = new AscFonts.CFont(i);
                }
            }
            oApi.pre_Paste(aFonts, oImageMap, function () {
                callback && callback();
            });
        };
        AscCommon.sendImgUrls(oApi, oObjectsForDownload.aUrls, fCallback, true);
        return null;
    };

    function CDocumentMerge(oOriginalDocument, oRevisedDocument, oOptions) {
        this.originalDocument = oOriginalDocument;
        this.revisedDocument = oRevisedDocument;
        this.options = oOptions;
        this.api = oOriginalDocument.GetApi();
        this.comparison = new CDocumentMergeComparison(oOriginalDocument, oRevisedDocument, oOptions ? oOptions : new AscCommonWord.ComparisonOptions());
        this.oldTrackRevisions = false;
    }
    CDocumentMerge.prototype.applyLastMergeCallback = function () {
        const oOriginalDocument = this.originalDocument;
        const oApi = this.api;
        if (!(oApi && oOriginalDocument)) {
            return;
        }
				this.comparison.removeCommentsFromMap();
        this.comparison.correctMoveMarks();
        oOriginalDocument.SetTrackRevisions(this.oldTrackRevisions);
        const oTrackRevisionManager = oOriginalDocument.TrackRevisionsManager;
        oTrackRevisionManager.SkipPreDeleteMoveMarks = this.oldSkipPreDeleteMoveMarks;
        oOriginalDocument.End_SilentMode(false);
	    if (this.comparison.oBookmarkManager.needUpdateBookmarks)
	    {
		    oOriginalDocument.UpdateBookmarks();
	    }
	      this.comparison.updateCommentsQuoteText();
        oOriginalDocument.Recalculate();
        oOriginalDocument.UpdateInterface();
        oOriginalDocument.FinalizeAction();
        oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
    };

    CDocumentMerge.prototype.merge = function () {
        const oOriginalDocument = this.originalDocument;
        const oRevisedDocument = this.revisedDocument;
        if (!oOriginalDocument || !oRevisedDocument) {
            return;
        }
        oOriginalDocument.StopRecalculate();
        oOriginalDocument.StartAction(AscDFH.historydescription_Document_MergeDocuments);
        oOriginalDocument.Start_SilentMode();
        this.oldTrackRevisions = oOriginalDocument.GetLocalTrackRevisions();
        oOriginalDocument.SetTrackRevisions(false);
        const oTrackRevisionManager = oOriginalDocument.TrackRevisionsManager;
        this.oldSkipPreDeleteMoveMarks = oTrackRevisionManager.SkipPreDeleteMoveMarks;
        oTrackRevisionManager.SkipPreDeleteMoveMarks = true;
        this.comparison.compare(this.applyLastMergeCallback.bind(this));
    };


    function mergeBinary(oApi, sBinary2, oOptions) {
        const oDoc1 = oApi.WordControl.m_oLogicDocument;
        if (!window['NATIVE_EDITOR_ENJINE']) {
            const oCollaborativeEditing = oDoc1.CollaborativeEditing;
            if (oCollaborativeEditing && !oCollaborativeEditing.Is_SingleUser()) {
                oApi.sendEvent("asc_onError", Asc.c_oAscError.ID.CannotCompareInCoEditing, c_oAscError.Level.NoCritical);
                return;
            }
        }
        oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);

	    const oldTrackRevisions = oDoc1.GetLocalTrackRevisions();
	    oDoc1.SetTrackRevisions(false);
        const oDoc2 = AscFormat.ExecuteNoHistory(function () {
            const openParams = {noSendComments: true, noGenerateSmartArts: true};
            const oTempDocument = new CDocument(oApi.WordControl.m_oDrawingDocument, false);
            const oBinaryFileReader = new AscCommonWord.BinaryFileReader(oTempDocument, openParams);
            AscCommon.pptx_content_loader.Start_UseFullUrl(oApi.insertDocumentUrlsData);
            if (!oBinaryFileReader.Read(sBinary2)) {
                return null;
            }
            return oTempDocument;
        }, this, []);
	    oDoc1.SetTrackRevisions(oldTrackRevisions);
        if (oDoc2) {
            const oMerge = new AscCommonWord.CDocumentMerge(oDoc1, oDoc2, oOptions ? oOptions : new AscCommonWord.ComparisonOptions());
            oMerge.merge();
        } else {
            AscCommon.pptx_content_loader.End_UseFullUrl();
        }

    }
    
    function mergeDocuments(oApi, oTmpDocument) {
        oApi.insertDocumentUrlsData = {
            imageMap: oTmpDocument["GetImageMap"](), documents: [], convertCallback: function (_api, url) {
            }, endCallback: function (_api) {
            }
        };
        mergeBinary(oApi, oTmpDocument["GetBinary"](), null, true);
        oApi.insertDocumentUrlsData = null;
    }

    window['AscCommonWord'].CDocumentMerge = CDocumentMerge;
    window['AscCommonWord'].mergeBinary = mergeBinary;
    window['AscCommonWord'].CMockDocument = CMockDocument;
    window['AscCommonWord'].CMockParagraph = CMockParagraph;
    window['AscCommonWord'].CDocumentResolveConflictComparison = CDocumentResolveConflictComparison;
    window['AscCommonWord'].CDocumentMergeComparison = CDocumentMergeComparison;
    window['AscCommonWord']["mergeDocuments"] = window['AscCommonWord'].mergeDocuments = mergeDocuments;

})();
