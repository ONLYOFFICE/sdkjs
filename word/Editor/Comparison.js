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

(function (undefined) {

    const MIN_JACCARD = 0.34;
    const MIN_DIFF = 0.7;
    const EXCLUDED_PUNCTUATION = {};
    //EXCLUDED_PUNCTUATION[46] = true;
    //EXCLUDED_PUNCTUATION[95] = true;
    EXCLUDED_PUNCTUATION[160] = true;
    //EXCLUDED_PUNCTUATION[63] = true;
	function getMockDocumentFromRuns(arrRuns) {
		const arrSlice = arrRuns.slice();
		arrSlice.push(new AscCommonWord.ParaRun());
		arrSlice[arrSlice.length - 1].Content.push(new AscWord.CRunParagraphMark());

		const mockDocument = new AscCommonWord.CMockDocument();
		const mockParagraph = new AscCommonWord.CMockParagraph();
		mockDocument.Content.push(mockParagraph);
		mockParagraph.Content = arrSlice;

		return mockDocument;
	}
	function getChanges(arrOriginalTextElements, arrRevisedTextElements, comparison, oMainParent, oRevisedParent) {
		let arrTextPrChanges = [];
		if (comparison.options.formatting) {
			const oTextPrChangeCollector = new CTextPrChangeCollector(arrOriginalTextElements, arrRevisedTextElements, oMainParent, oRevisedParent, comparison.copyPr);
			arrTextPrChanges = oTextPrChangeCollector.getTextPrChanges();
		}

		let arrReviewChanges = [];
		if (comparison.needCheckReview) {
			const oReviewChangeCollector = new CReviewChangeCollector(arrOriginalTextElements, arrRevisedTextElements, comparison.oComparisonMoveMarkManager, oMainParent, oRevisedParent, comparison.copyPr);
			arrReviewChanges = oReviewChangeCollector.getReviewChanges();
		}

		let oCopyPr;
		if (comparison.needCopyForResolveEqualWords) {
			oCopyPr = comparison.copyPr;
		}
		const oCommentChangeCollector = new CCommentChangesCollector(arrOriginalTextElements, arrRevisedTextElements, oCopyPr, comparison.oCommentManager);
		const arrCommentChanges = oCommentChangeCollector.getCommentChanges();

		const oBookmarkChangeCollector = new CBookmarkChangesCollector(arrRevisedTextElements, oCopyPr);
		const arrBookmarkChanges = oBookmarkChangeCollector.getBookmarkChanges();

		const arrChanges = arrTextPrChanges.concat(arrReviewChanges, arrBookmarkChanges, arrCommentChanges);
		return arrChanges.sort(function (a, b) {
			if (a.elementIndex === b.elementIndex) {
				if (b.innerElementIndex === a.innerElementIndex) {
					return !!b.isStart - !!a.isStart;
				}
				return b.innerElementIndex - a.innerElementIndex;
			}
			return b.elementIndex - a.elementIndex;
		});
	}
	function resolveTypesWithPartner(arrNodes, comparison) {
		const arrMainElements = [];
		const arrRevisedElements = [];
		for (let i = 0; i < arrNodes.length; i += 1) {
			const oNode = arrNodes[i];
			const oMainElement = oNode.element;
			const oRevisedElement = oNode.partner.element;

			arrMainElements.push(oMainElement);
			arrRevisedElements.push(oRevisedElement);
		}
		if (!arrMainElements.length) {
			return;
		}

		const oFirstNode = arrNodes[0];
		const oRevisedParent = oFirstNode.partner.par.element;
		const oMainParent = oFirstNode.par.element;


		const oNeedReviewWithUser = {};
		const arrChanges = getChanges(arrMainElements, arrRevisedElements, comparison, oMainParent, oRevisedParent);

		const oTextIterator = new CRunCollector(arrMainElements, comparison.parentParagraph);

		for (let i = 0; i < arrChanges.length; i += 1) {
			const oChange = arrChanges[i];
			oChange.apply(oTextIterator, comparison, oNeedReviewWithUser);
		}
		comparison.applyResolveTypes(oNeedReviewWithUser);
	}
	function WordCounter() {

		this.mapWords = {};
		this.count = 0;
	}
	WordCounter.prototype.update = function(aCodes, oComparison) {
		++this.count;
		let text = String.fromCharCode.apply(String, aCodes);
		if (!oComparison.options.caseChanges) {
			text = text.toLowerCase();
		}
		if (!this.mapWords[text]) {
			this.mapWords[text] = 0;
		}
		this.mapWords[text] += 1;
	};
	WordCounter.prototype.diff = function(other) {
		let shared = 0;
		for (let i in this.mapWords) {
			if (other.mapWords[i]) {
				shared += Math.min(other.mapWords[i], this.mapWords[i]);
			}
		}
		return other.count + this.count - 2 * shared;
	};

	function changeFirstTextElement(oTextIterator, oRun)
	{
		const oElement = oTextIterator.getCurrentElement();
		oElement.elements[0] = oRun.Content[0];
		oElement.firstRun = oRun;
	}

	function CReviewChange(oReviewInfo)
	{
		this.elementIndex = -1;
		this.innerElementIndex = -1;

		this.reviewInfo = oReviewInfo;
		this.reviewType = null;
		this.moveReviewType = null
		this.moveReviewMarkName = null;

		this.collectRuns = [];
		this.isStart = false;
		this.partner = null;
		this.mapRuns = {};
	}
	CReviewChange.prototype.apply = function (oTextIterator, comparison, oNeedReviewWithUser) {
		if (this.isStart) {
			this.applyStart(oTextIterator, comparison, oNeedReviewWithUser);
		} else {
			this.applyEnd(oTextIterator);
		}
	};
	CReviewChange.prototype.applyStart = function (oTextIterator, comparison, oNeedReviewWithUser) {
		const bIsLast = oTextIterator.elements[this.elementIndex].elements.length === this.innerElementIndex;
		oTextIterator.skipTo(this.elementIndex, bIsLast ? this.innerElementIndex - 1 : this.innerElementIndex);
		const oRun = oTextIterator.splitCurrentRun(bIsLast);
		const oPartner = this.partner;
		oPartner.dropLastRun();
		oTextIterator.addToCollectBack(oRun);
		if (this.elementIndex === 0 && this.innerElementIndex === 0)
		{
			changeFirstTextElement(oTextIterator, oRun);
		}


		oTextIterator.setCollectReviewRuns(null);
		const arrRuns = oPartner.collectRuns;

		const nPriorityReviewType = oPartner.reviewType;
		const nPriorityMoveReviewType = oPartner.moveReviewType;
		const sMoveReviewMarkName = oPartner.moveReviewMarkName;
		const oReviewInfo = oPartner.reviewInfo;
		const sReviewUserName = oReviewInfo.GetUserName();
		const sReviewDate = oReviewInfo.GetDateTime();


		if (!oNeedReviewWithUser[sReviewDate]) {
			oNeedReviewWithUser[sReviewDate] = {};
		}

		if (!oNeedReviewWithUser[sReviewDate][sReviewUserName]) {
			const oNeedReview = {reviewTypes: {}, moveReviewTypes: {}};
			oNeedReview.reviewTypes[reviewtype_Add] = [];
			oNeedReview.reviewTypes[reviewtype_Remove] = [];
			oNeedReview.moveReviewTypes[Asc.c_oAscRevisionsMove.MoveTo] = [];
			oNeedReview.moveReviewTypes[Asc.c_oAscRevisionsMove.MoveFrom] = [];
			oNeedReviewWithUser[sReviewDate][sReviewUserName] = oNeedReview;
		}

		if (AscFormat.isRealNumber(nPriorityReviewType)) {
			for (let i = 0; i < arrRuns.length; i += 1)
			{
				oNeedReviewWithUser[sReviewDate][sReviewUserName].reviewTypes[nPriorityReviewType].push({element: arrRuns[i], reviewInfo: oReviewInfo});
			}
		}
		else if (AscFormat.isRealNumber(nPriorityMoveReviewType))
		{
			for (let i = 0; i < arrRuns.length; i += 1)
			{
				comparison.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sMoveReviewMarkName, arrRuns[i]);
				comparison.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveReviewMarkName, arrRuns[i]);
				oNeedReviewWithUser[sReviewDate][sReviewUserName].moveReviewTypes[nPriorityMoveReviewType].push({element: arrRuns[i], reviewInfo: oReviewInfo});												}
		}
	};
	CReviewChange.prototype.applyEnd = function (oTextIterator) {
		const bIsLast = oTextIterator.elements[this.elementIndex].elements.length === this.innerElementIndex;
		oTextIterator.skipTo(this.elementIndex, bIsLast ? this.innerElementIndex - 1 : this.innerElementIndex);
		oTextIterator.setCollectReviewRuns(this);
		oTextIterator.addToCollectCurrentRun();
		const oSplitRun = oTextIterator.splitCurrentRun(bIsLast);
		oTextIterator.addToCollectTextPr(oSplitRun);
	};
	CReviewChange.prototype.addRun = function (oRun) {
		if (!this.mapRuns[oRun.Id]) {
			this.mapRuns[oRun.Id] = true;
			this.collectRuns.push(oRun);
		}
	};

	CReviewChange.prototype.addRunBack = function (oRun) {
		if (!this.mapRuns[oRun.Id]) {
			this.mapRuns[oRun.Id] = true;
			this.collectRuns.unshift(oRun);
		}
	};
	CReviewChange.prototype.dropLastRun = function () {
		const oRun = this.collectRuns.pop();
		if (oRun) {
			delete this.mapRuns[oRun.Id];
		}
	};
	CReviewChange.prototype.setPosition = function (nElementIndex, nInnerElementIndex) {
		this.elementIndex = nElementIndex;
		this.innerElementIndex = nInnerElementIndex;
	};
	CReviewChange.prototype.setMoveReviewType = function (nMoveReviewType, sMoveReviewMarkName)
	{
		this.moveReviewType = nMoveReviewType;
		this.moveReviewMarkName = sMoveReviewMarkName;
	}
	CReviewChange.prototype.isSimilar = function (oOtherChange) {
		if (this.reviewType !== oOtherChange.reviewType) {
			return false;
		}

		if (this.moveReviewType !== oOtherChange.moveReviewType) {
			return false;
		}

		if (this.moveReviewMarkName !== oOtherChange.moveReviewMarkName) {
			return false;
		}

		if (this.reviewInfo && !this.reviewInfo.IsEqual(oOtherChange.reviewInfo) || !this.reviewInfo && oOtherChange.reviewInfo) {
			return false;
		}

		return true;
	};
	CReviewChange.prototype.setPartner = function (oPartner) {
		this.partner = oPartner;
		oPartner.partner = this;
	};
	CReviewChange.prototype.setReviewType = function (nReviewType)
	{
		this.reviewType = nReviewType;
	}

	function CLabelChange(arrLabels, nElementIndex, nInnerElementIndex)
	{
		this.labels = arrLabels;
		this.elementIndex = nElementIndex;
		this.innerElementIndex = nInnerElementIndex;
	}
	CLabelChange.prototype.apply = function (oTextIterator, comparison, oNeedReviewWithUser) {
		const bIsLast = oTextIterator.elements[this.elementIndex].elements.length === this.innerElementIndex;
		oTextIterator.skipTo(this.elementIndex, bIsLast ? this.innerElementIndex - 1 : this.innerElementIndex);
		const oRun = oTextIterator.splitCurrentRun(bIsLast);
		oTextIterator.addToCollectBack(oRun);
		for (let i = this.labels.length - 1; i >= 0; i -= 1)
		{
			const oLabel = this.labels[i];
			oTextIterator.parent.AddToContent(oTextIterator.runIndex + 1, oLabel);
		}
		if (this.elementIndex === 0 && this.innerElementIndex === 0)
		{
			const oElement = oTextIterator.getCurrentElement();
			const oFirstLabel = this.labels[0];
			if (oFirstLabel)
			{
				oElement.lastSwitchElement = oFirstLabel;
			}
			else
			{
				changeFirstTextElement(oTextIterator, oRun);
			}
		}
	};

	function CBookmarkChangesCollector(arrElements, copyPr) {
		this.elements = arrElements;
		this.copyPr = copyPr;
	}
	CBookmarkChangesCollector.prototype.getBookmarkChanges = function () {
		const arrResult = [];
		for (let i = this.elements.length - 1; i >= 0; i -= 1) {
			const oElement = this.elements[i];
			const arrInsertIndexes = oElement.getBookmarkInsertIndexes();
			for (let j = 0; j < arrInsertIndexes.length; j += 1) {
				const nInsertIndex = arrInsertIndexes[j];
				const arrElements = this.getBookmarkElements(oElement, nInsertIndex);
				arrResult.push(new CLabelChange(arrElements, i, nInsertIndex));
			}
		}
		return arrResult;
	}
	CBookmarkChangesCollector.prototype.getBookmarkElements = function (oTextElement, nInsertIndex) {
		const arrElements = oTextElement.bookmarks[nInsertIndex].slice();
		if (this.copyPr) {
			for (let i = 0; i < arrElements.length; i += 1) {
				arrElements[i] = arrElements[i].Copy(false, this.copyPr);
			}
		}
		return arrElements;
	};

	function CCommentChangesCollector(arrMainElements, arrRevisedElements, copyPr, oCommentManager) {
		this.mainElements = arrMainElements;
		this.revisedElements = arrRevisedElements;
		this.commentManager = oCommentManager;
		this.copyPr = copyPr;
	}
	CCommentChangesCollector.prototype.getCommentChanges = function () {
		const arrResult = [];
		for (let i = this.mainElements.length - 1; i >= 0; i -= 1) {
			const oMainElement = this.mainElements[i];
			const oRevisedElement = this.revisedElements[i];
			const arrInsertIndexes = oRevisedElement.getCommentInsertIndexes();
			for (let j = 0; j < arrInsertIndexes.length; j += 1) {
				const nInsertIndex = arrInsertIndexes[j];
				const arrCommentElements = this.getCommentElements(oMainElement, oRevisedElement, nInsertIndex);
				arrResult.push(new CLabelChange(arrCommentElements, i, nInsertIndex));
			}
		}
		return arrResult;
	}
	CCommentChangesCollector.prototype.getCommentElements = function (oMainElement, oRevisedElement, nInsertIndex) {
		this.checkComments(oMainElement, oRevisedElement, nInsertIndex);
		const arrElements = [];
		for (let i = 0; i < oRevisedElement.comments[nInsertIndex].length; i += 1)
		{
			const oParaComment = oRevisedElement.comments[nInsertIndex][i].element;
			arrElements.push(oParaComment);
			if (oParaComment.IsCommentStart())
			{
				const sCommentId = oParaComment.GetCommentId();
				const arrLaterMerge = this.commentManager.getMergeLater(sCommentId);
				if (arrLaterMerge)
				{
					arrElements.push.apply(arrElements, arrLaterMerge);
				}
			}
		}

		if (this.copyPr) {
			for (let i = 0; i < arrElements.length; i += 1) {
				arrElements[i] = arrElements[i].Copy(false, this.copyPr);
			}
		}
		return arrElements;
	};

	CCommentChangesCollector.prototype.checkComments = function (oMainElement, oRevisedElement, nInsertIndex) {
		const arrRevisedComments = oRevisedElement.comments[nInsertIndex];
		const arrMainComments = oMainElement.comments[nInsertIndex];
		this.commentManager.checkComments(arrMainComments, arrRevisedComments, arrRevisedComments);
	};

	function CReviewChangeCollector(arrMainElements, arrRevisedElements, oMoveMarkManager, oMainParagraph, oRevisedParagraph, oCopyPr) {
		this.mainElements = arrMainElements;
		this.revisedElements = arrRevisedElements;
		this.moveMarkManager = oMoveMarkManager;
		this.mainParagraph = oMainParagraph;
		this.revisedParagraph = oRevisedParagraph;
		this.copyPr = oCopyPr;
	}
	CReviewChangeCollector.prototype.isSkippedElement = function (oRun) {
		if (!(oRun instanceof AscCommonWord.ParaRun)) {
			return true;
		}
		if (this.copyPr.SkipFootnoteReference) {
			return oRun.IsFootEndnoteReferenceRun();
		}
		return false;
	};
	CReviewChangeCollector.prototype.getPriorityChange = function (oRevisedRun, oMainRun) {
		const oRevisedReviewInfo = oRevisedRun.GetReviewInfo();
		const nRevisedReviewType = oRevisedRun.GetReviewType();
		const nRevisedMoveReviewType = oRevisedRun.GetReviewMoveType();
		const sRevisedNameMoveMark = this.moveMarkManager.getMoveMarkNameByRun(oRevisedRun);
		const oRevisedPrevAdded = oRevisedReviewInfo && oRevisedReviewInfo.GetPrevAdded();

		const oMainReviewInfo = oMainRun.GetReviewInfo();
		const nMainReviewType = oMainRun.GetReviewType();
		const oMainPrevAdded = oMainReviewInfo && oMainReviewInfo.GetPrevAdded();
		const nMainMoveReviewType = oMainRun.GetReviewMoveType();

		let nPriorityReviewType;
		let oPriorityReviewInfo;
		let nPriorityMoveReviewType;
		let bIsMove = false;
		if (nRevisedReviewType !== reviewtype_Common && !(oMainPrevAdded && nMainReviewType === reviewtype_Remove)) {
			if (nMainReviewType !== nRevisedReviewType) {
				nPriorityReviewType = nRevisedReviewType;
				oPriorityReviewInfo = oRevisedReviewInfo;
			} else if (nMainReviewType === reviewtype_Remove && oRevisedPrevAdded) {
				nPriorityReviewType = reviewtype_Add;
				oPriorityReviewInfo = oRevisedPrevAdded;
			}
		}
		if (nMainMoveReviewType === Asc.c_oAscRevisionsMove.NoMove && nRevisedMoveReviewType !== Asc.c_oAscRevisionsMove.NoMove) {
			nPriorityMoveReviewType = nRevisedMoveReviewType;
			oPriorityReviewInfo = oRevisedReviewInfo;
			bIsMove = true;
		}

		if ((AscFormat.isRealNumber(nPriorityReviewType) || AscFormat.isRealNumber(nPriorityMoveReviewType)) && oPriorityReviewInfo)
		{
			const oChange = new CReviewChange(oPriorityReviewInfo);
			if (bIsMove)
			{
				oChange.setMoveReviewType(nPriorityMoveReviewType, sRevisedNameMoveMark);
			}
			else
			{
				oChange.setReviewType(nPriorityReviewType);
			}
			return oChange;
		}
	};
	CReviewChangeCollector.prototype.getRevisedLastRunIndex = function () {
		const oLastElement = this.revisedElements[this.revisedElements.length - 1];
		const oLastRun = oLastElement.lastRun;
		return this.getLastRunIndex(oLastRun, oLastElement.elements[oLastElement.elements.length - 1], this.revisedParagraph);
	};
	CReviewChangeCollector.prototype.getMainLastRunIndex = function () {
		const oLastElement = this.mainElements[this.mainElements.length - 1];
		const oLastRun = oLastElement.lastRun;
		return this.getLastRunIndex(oLastRun, oLastElement.elements[oLastElement.elements.length - 1], this.mainParagraph);
	};
	CReviewChangeCollector.prototype.getLastRunIndex = function (oRun, oElement, oRunParagraph) {
		const oContent = oRunParagraph.Content;
		for (let i = oContent.length - 1; i >= 0; i -= 1) {
			if (oContent[i] === oRun) {
				for (let j = oRun.Content.length - 1; j >= 0; j -= 1) {
					const oRunElement = oRun.Content[j];
					if (oRunElement === oElement) {
						return {runIndex: i, elementIndex: j};
					}
				}
			}
		}
		return null;
	};
	CReviewChangeCollector.prototype.getReviewChanges = function () {
		const arrResult = [];

		let nElementIndex = this.revisedElements.length - 1;
		let nInnerElementIndex = this.revisedElements[nElementIndex].elements.length - 1;

		const oRevisedContent = this.revisedParagraph.Content;
		const oMainContent = this.mainParagraph.Content;

		const oLastRevisedRunInfo = this.getRevisedLastRunIndex();
		const oLastMainRunInfo = this.getMainLastRunIndex();
		if (!oLastMainRunInfo || !oLastRevisedRunInfo) {
			return arrResult;
		}

		let nRevisedRunIndex = oLastRevisedRunInfo.runIndex;
		let nMainRunIndex = oLastMainRunInfo.runIndex;

		let bSaveOldRevisedCount = true;
		let bSaveOldMainCount = true;
		let nRevisedElementsCount = oLastRevisedRunInfo.elementIndex + 1;
		let nMainElementsCount = oLastMainRunInfo.elementIndex + 1;
		while (nRevisedRunIndex >= 0 && nMainRunIndex >= 0) {
			const oRevisedRun = oRevisedContent[nRevisedRunIndex];
			const oMainRun = oMainContent[nMainRunIndex];

			if (!bSaveOldRevisedCount) {
				nRevisedElementsCount = oRevisedRun.Content.length;
				bSaveOldRevisedCount = true;
			}
			if (!bSaveOldMainCount) {
				nMainElementsCount = oMainRun.Content.length;
				bSaveOldMainCount = true;
			}
			if (nMainElementsCount !== 0 && nRevisedElementsCount !== 0) {
				const oLastPrChange = arrResult[arrResult.length - 1];
				const oPriorityChange = this.getPriorityChange(oRevisedRun, oMainRun);
				if (!oPriorityChange) {
					if (oLastPrChange && !oLastPrChange.isStart) {
						const oStartReviewChange = new CReviewChange(null);
						oStartReviewChange.isStart = true;
						oStartReviewChange.setPosition(nElementIndex, nInnerElementIndex + 1);
						oStartReviewChange.setPartner(oLastPrChange);
						arrResult.push(oStartReviewChange);
					}
				} else if (oLastPrChange) {
					if (oLastPrChange.isStart) {
						oPriorityChange.setPosition(nElementIndex, nInnerElementIndex + 1);
						arrResult.push(oPriorityChange);
					} else if (!oLastPrChange.isSimilar(oPriorityChange)) {
						const oStartReviewChange = new CReviewChange(null);
						oStartReviewChange.isStart = true;
						oStartReviewChange.setPosition(nElementIndex, nInnerElementIndex + 1);
						oStartReviewChange.setPartner(oLastPrChange);
						arrResult.push(oStartReviewChange);

						oPriorityChange.setPosition(nElementIndex, nInnerElementIndex + 1);
						arrResult.push(oPriorityChange);
					}
				} else {
					oPriorityChange.setPosition(nElementIndex, nInnerElementIndex + 1);
					arrResult.push(oPriorityChange);
				}
			}

			while (nRevisedElementsCount !== 0 && nMainElementsCount !== 0) {
				const nMinCount = Math.min(nRevisedElementsCount, nMainElementsCount);
				if (nInnerElementIndex + 1 > nMinCount) {
					nInnerElementIndex -= nMinCount;
					nRevisedElementsCount -= nMinCount;
					nMainElementsCount -= nMinCount;
				} else {
					nRevisedElementsCount -= nInnerElementIndex + 1;
					nMainElementsCount -= nInnerElementIndex + 1;
					nElementIndex -= 1;
					if (this.revisedElements[nElementIndex] && this.mainElements[nElementIndex]) {
						nInnerElementIndex = this.revisedElements[nElementIndex].revisedElements.length - 1;
					}
				}
				if (nElementIndex < 0) {
					const oLastPrChange = arrResult[arrResult.length - 1];
					if (oLastPrChange && !oLastPrChange.isStart) {
						const oStartReviewChange = new CReviewChange(null);
						oStartReviewChange.isStart = true;
						oStartReviewChange.setPosition(0, 0);
						oStartReviewChange.setPartner(oLastPrChange);
						arrResult.push(oStartReviewChange);
					}
					return arrResult;
				}
			}
			if (nRevisedElementsCount === 0) {
				do {
					nRevisedRunIndex -= 1;
					bSaveOldRevisedCount = false;
				} while (this.isSkippedElement(oRevisedContent[nRevisedRunIndex]) && nRevisedRunIndex >= 0)
			}
			if (nMainElementsCount === 0) {
				do {
					nMainRunIndex -= 1;
					bSaveOldMainCount = false;
				} while (this.isSkippedElement(oMainContent[nMainRunIndex]) && nMainRunIndex >= 0)
			}
		}
		return arrResult;
	};

	function CTextPrChange(isStart, elementIndex, innerElementIndex, textPr) {
		this.isStart = !!isStart;
		this.elementIndex = elementIndex;
		this.innerElementIndex = innerElementIndex;
		this.textPr = textPr;
		this.partner = null;
		this.collectRuns = [];
		this.mapRuns = {};
	}

	CTextPrChange.prototype.apply = function (oTextIterator, comparison, oNeedReviewWithUser) {
		if (this.isStart) {
			this.applyStart(oTextIterator, comparison);
		} else {
			this.applyEnd(oTextIterator);
		}
	};
	CTextPrChange.prototype.applyStart = function (oTextIterator, oComparison) {
		const bIsLast = oTextIterator.elements[this.elementIndex].elements.length === this.innerElementIndex;
		oTextIterator.skipTo(this.elementIndex, bIsLast ? this.innerElementIndex - 1 : this.innerElementIndex);
		const oPartner = this.partner;
		const oSplitRun = oTextIterator.splitCurrentRun(bIsLast);
		oPartner.dropLastRun();
		oTextIterator.addToCollectBack(oSplitRun);
		if (this.elementIndex === 0 && this.innerElementIndex === 0)
		{
			changeFirstTextElement(oTextIterator, oSplitRun);
		}
		oTextIterator.setCollectTextPrRuns(null);
		const arrRuns = oPartner.collectRuns;
		const oNewTextPr = oPartner.textPr;

		for (let i = 0; i < arrRuns.length; i += 1) {
			const oRun = arrRuns[i];
			const oCopyTextPr = oNewTextPr.Copy();
			const oOldTextPr = oRun.GetTextPr();
			const oReviewInfo = oComparison.getReviewInfo();
			oCopyTextPr.SetPrChange(oOldTextPr, oReviewInfo);
			oRun.SetPr(oCopyTextPr);
		}
	};
	CTextPrChange.prototype.applyEnd = function (oTextIterator) {
		const bIsLast = oTextIterator.elements[this.elementIndex].elements.length === this.innerElementIndex;
		oTextIterator.skipTo(this.elementIndex, bIsLast ? this.innerElementIndex - 1 : this.innerElementIndex);

		oTextIterator.setCollectTextPrRuns(this);
		oTextIterator.addToCollectCurrentRun();
		const oSplitRun = oTextIterator.splitCurrentRun(bIsLast);
		oTextIterator.addToCollectReviewChange(oSplitRun);
	};
	CTextPrChange.prototype.setPartner = function (partner) {
		this.partner = partner;
	};

	CTextPrChange.prototype.addRun = function (oRun) {
		if (!this.mapRuns[oRun.Id]) {
			this.mapRuns[oRun.Id] = true;
			this.collectRuns.push(oRun);
		}
	};

	CTextPrChange.prototype.addRunBack = function (oRun) {
		if (!this.mapRuns[oRun.Id]) {
			this.mapRuns[oRun.Id] = true;
			this.collectRuns.unshift(oRun);
		}
	};
	CTextPrChange.prototype.dropLastRun = function () {
		const oRun = this.collectRuns.pop();
		if (oRun) {
			delete this.mapRuns[oRun.Id];
		}
	};

	function CTextPrChangeCollector(arrMainElements, arrRevisedElements, oMainParagraph, oRevisedParagraph, oCopyPr) {
		this.elements = arrRevisedElements;
		this.comparedElements = arrMainElements;
		this.mainParagraph = oMainParagraph;
		this.revisedParagraph = oRevisedParagraph;
		this.copyPr = oCopyPr;
	}
	CTextPrChangeCollector.prototype.isSkippedElement = CReviewChangeCollector.prototype.isSkippedElement;
	CTextPrChangeCollector.prototype.getRevisedLastRunIndex = function () {
		const oLastElement = this.elements[this.elements.length - 1];
		const oLastRun = oLastElement.lastRun;
		return this.getLastRunIndex(oLastRun, oLastElement.elements[oLastElement.elements.length - 1], this.revisedParagraph);
	};
	CTextPrChangeCollector.prototype.getMainLastRunIndex = function () {
		const oLastElement = this.comparedElements[this.comparedElements.length - 1];
		const oLastRun = oLastElement.lastRun;
		return this.getLastRunIndex(oLastRun, oLastElement.elements[oLastElement.elements.length - 1], this.mainParagraph);
	};
	CTextPrChangeCollector.prototype.getLastRunIndex = function (oRun, oElement, oRunParagraph) {
		const oContent = oRunParagraph.Content;
		for (let i = oContent.length - 1; i >= 0; i -= 1) {
			if (oContent[i] === oRun) {
				for (let j = oRun.Content.length - 1; j >= 0; j -= 1) {
					const oRunElement = oRun.Content[j];
					if (oRunElement === oElement) {
						return {runIndex: i, elementIndex: j};
					}
				}
			}
		}
		return null;
	};
	CTextPrChangeCollector.prototype.getTextPrChanges = function () {
		const arrResult = [];

		let nElementIndex = this.elements.length - 1;
		let nInnerElementIndex = this.elements[nElementIndex].elements.length - 1;

		const oRevisedContent = this.revisedParagraph.Content;
		const oMainContent = this.mainParagraph.Content;

		const oLastRevisedRunInfo = this.getRevisedLastRunIndex();
		const oLastMainRunInfo = this.getMainLastRunIndex();
		if (!oLastMainRunInfo || !oLastRevisedRunInfo) {
			return arrResult;
		}

		let nRevisedRunIndex = oLastRevisedRunInfo.runIndex;
		let nMainRunIndex = oLastMainRunInfo.runIndex;

		let bSaveOldRevisedCount = true;
		let bSaveOldMainCount = true;
		let nRevisedElementsCount = oLastRevisedRunInfo.elementIndex + 1;
		let nMainElementsCount = oLastMainRunInfo.elementIndex + 1;
		while (nRevisedRunIndex >= 0 && nMainRunIndex >= 0) {
			const oRevisedRun = oRevisedContent[nRevisedRunIndex];
			const oMainRun = oMainContent[nMainRunIndex];
			const oMainTextPr = oMainRun.Pr;
			const oRevisedTextPr = oRevisedRun.Pr.Copy(undefined, this.copyPr);

			if (!bSaveOldRevisedCount) {
				nRevisedElementsCount = oRevisedRun.Content.length;
				bSaveOldRevisedCount = true;
			}
			if (!bSaveOldMainCount) {
				nMainElementsCount = oMainRun.Content.length;
				bSaveOldMainCount = true;
			}
			if (nMainElementsCount !== 0 && nRevisedElementsCount !== 0) {
				const oLastPrChange = arrResult[arrResult.length - 1];
				if (oMainTextPr.Is_Equal(oRevisedTextPr)) {
					if (oLastPrChange && !oLastPrChange.isStart) {
						const oStartPrChange = new CTextPrChange(true, nElementIndex, nInnerElementIndex + 1);
						oStartPrChange.setPartner(oLastPrChange);
						arrResult.push(oStartPrChange);
					}
				} else if (oLastPrChange) {
						if (oLastPrChange.isStart) {
							arrResult.push(new CTextPrChange(false, nElementIndex, nInnerElementIndex + 1, oRevisedTextPr));
						} else if (!oLastPrChange.textPr.Is_Equal(oRevisedTextPr)) {
							const oStartPrChange = new CTextPrChange(true, nElementIndex, nInnerElementIndex + 1);
							oStartPrChange.setPartner(oLastPrChange);
							arrResult.push(oStartPrChange);
							arrResult.push(new CTextPrChange(false, nElementIndex, nInnerElementIndex + 1, oRevisedTextPr));
						}
					} else {
						arrResult.push(new CTextPrChange(false, nElementIndex, nInnerElementIndex + 1, oRevisedTextPr));
					}
				
			}

			while (nRevisedElementsCount !== 0 && nMainElementsCount !== 0) {
				const nMinCount = Math.min(nRevisedElementsCount, nMainElementsCount);
				if (nInnerElementIndex + 1 > nMinCount) {
					nInnerElementIndex -= nMinCount;
					nRevisedElementsCount -= nMinCount;
					nMainElementsCount -= nMinCount;
				} else {
					nRevisedElementsCount -= nInnerElementIndex + 1;
					nMainElementsCount -= nInnerElementIndex + 1;
					nElementIndex -= 1;
					if (this.elements[nElementIndex] && this.comparedElements[nElementIndex]) {
						nInnerElementIndex = this.elements[nElementIndex].elements.length - 1;
					}
				}
				if (nElementIndex < 0) {
					const oLastPrChange = arrResult[arrResult.length - 1];
					if (oLastPrChange && !oLastPrChange.isStart) {
						const oStartPrChange = new CTextPrChange(true, 0, 0);
						oStartPrChange.setPartner(oLastPrChange);
						arrResult.push(oStartPrChange);
					}
					return arrResult;
				}
			}
			if (nRevisedElementsCount === 0) {
				do {
					nRevisedRunIndex -= 1;
					bSaveOldRevisedCount = false;
				} while (this.isSkippedElement(oRevisedContent[nRevisedRunIndex]) && nRevisedRunIndex >= 0)
			}
			if (nMainElementsCount === 0) {
				do {
					nMainRunIndex -= 1;
					bSaveOldMainCount = false;
				} while (this.isSkippedElement(oMainContent[nMainRunIndex]) && nMainRunIndex >= 0)
			}
		}
		return arrResult;
	};

    function CNode(oElement, oParent)
    {
        this.element = oElement;
        this.par = null;
        this.children = [];
        this.depth   = 0;
        this.changes = [];
        this.partner = null;
        this.childidx = null;

        this.wordCounter = null;
        if(oParent)
        {
            oParent.addChildNode(this);
        }
    }
	CNode.prototype.checkLastNodeOnInsertWithResolveConflict = function(comparison, aContentToInsert, oChange) {
			if (!comparison.parentParagraph) {
				return true;
			}
		const applyingParagraph = comparison.parentParagraph;
		const index = oChange.anchor.index;
		if (index === this.children.length - 1) {

			const oLastConflictElement = this.children[this.children.length - 2].element;
			const nInsertIndex = oLastConflictElement.lastRun.GetPosInParent(applyingParagraph);
			const nLastSymbolPosition = oLastConflictElement.lastRun.GetElementPosition(oLastConflictElement.elements[oLastConflictElement.elements.length - 1]);
			if (nLastSymbolPosition !== -1) {
				const oNewRun = oLastConflictElement.lastRun.Split2(nLastSymbolPosition + 1, applyingParagraph, nInsertIndex);
				comparison.checkOriginalAndSplitRun(oNewRun, oLastConflictElement.lastRun);
				this.applyInsert(aContentToInsert, [], nInsertIndex + 1, comparison);
				return false;
			}
		}
		return true;
	};
		CNode.prototype.getDiff = function (oAnotherNode, oComparison) {
			if (this.equals(oAnotherNode, oComparison)) {
				let nWordDiff = this.wordCounter.diff(oAnotherNode.wordCounter);
				if (this.element instanceof AscWord.Paragraph) {
					let startType;
					let endType;
					if (this.element.SectPr) {
						startType = this.element.SectPr.Type;
					}
					if (oAnotherNode.element.SectPr) {
						endType = oAnotherNode.element.SectPr.Type;
					}
					nWordDiff += endType !== startType ? 1 : 0;
				}
				return nWordDiff;
			}
			return this.wordCounter.count + oAnotherNode.wordCounter.count + 2;
		};
		CNode.prototype.isSpaceText = function() {
			return this.element instanceof CTextElement && this.element.isSpaceText();
		};
	CNode.prototype.resolveTypesWithPartner = function (comparison) {
			if (this.partner && this.element instanceof CTextElement) {
				resolveTypesWithPartner([this], comparison);
			}
	};
    CNode.prototype.getElement = function()
    {
        return this.element;
    };
	CNode.prototype.addBookmark = function (oBookmark, nInsertIndex)
	{
		if (this.element instanceof CTextElement)
		{
			this.element.addBookmark(oBookmark, nInsertIndex);
		}
	};
	CNode.prototype.getLastBookmarks = function ()
	{
		if (this.element instanceof CTextElement)
		{
			return this.element.getLastBookmarks();
		}
	};
	CNode.prototype.getFirstBookmarks = function ()
	{
		if (this.element instanceof CTextElement)
		{
			return this.element.getFirstBookmarks();
		}
	};
	CNode.prototype.addComment = function (oBookmark, nInsertIndex)
	{
		if (this.element instanceof CTextElement)
		{
			this.element.addComment(oBookmark, nInsertIndex);
		}
	};
	CNode.prototype.getLastComments = function ()
	{
		if (this.element instanceof CTextElement)
		{
			return this.element.getLastComments();
		}
	};
	CNode.prototype.getFirstComments = function ()
	{
		if (this.element instanceof CTextElement)
		{
			return this.element.getFirstComments();
		}
	};
		// debug method
	CNode.prototype.getText = function()
	{
		return this.element.getText();
	};

	CNode.prototype.getContentElement = function (isStart)
	{
		const element = this.element;
		if (isStart)
		{
			if (element.firstRun)
			{
				return element.firstRun;
			}
		}
		else
		{
			if (element.lastRun)
			{
				return element.lastRun;
			}
		}
		if (element instanceof CCommentElement)
		{
			return element.element;
		}
		return element;
	};
    CNode.prototype.cleanEndOfInsert = function (aContentToInsert, oChange, comparison) {
				const oNode = oChange.insert[oChange.insert.length - 1];
        const oLastText = oNode.element;
				const oEndOfInsertRun = oNode.getContentElement();
        const oParentParagraph =  (this.partner && this.partner.element) || oEndOfInsertRun.GetParent();
        const applyingParagraph = this.getApplyParagraph(comparison);
	    const arrEndBookmarks = oNode.getLastBookmarks();
	    if (arrEndBookmarks)
	    {
				for (let i = arrEndBookmarks.length - 1; i >= 0; i -= 1)
				{
					this.pushToArrInsertContentWithCopy(aContentToInsert, arrEndBookmarks[i], comparison);
				}
	    }
	    const arrLastComments = oNode.getLastComments();
	    if (arrLastComments)
	    {
		    for (let i = arrLastComments.length - 1; i >= 0; i -= 1)
		    {
			    this.pushToArrInsertContentWithCopy(aContentToInsert, arrLastComments[i].element, comparison);
		    }
	    }
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
								let reviewInfo = oEndOfInsertRun.GetReviewInfo();
                                this.setCommonReviewTypeWithInfo(lastCheckRun, reviewInfo ? reviewInfo.Copy() : undefined);
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
    };
    // comparison need for extends
    CNode.prototype.pushToArrInsertContent = function (aContentToInsert, elem, comparison) {
        aContentToInsert.push(elem);
    };

    CNode.prototype.pushToArrInsertContentWithCopy = function (aContentToInsert, elem, comparison) {
			if (elem instanceof AscCommon.ParaComment && !comparison.options.comments) {
				return;
			}
        const elemCopy = elem.Copy(false, comparison.copyPr);
        this.pushToArrInsertContent(aContentToInsert, elemCopy, comparison);
				if (elem instanceof AscCommon.ParaComment)
				{
					comparison.oCommentManager.pushToArrInsertContentFromMerge(elem, aContentToInsert);
				}
    };

    CNode.prototype.cleanStartOfInsertSameRun = function (oNewRun, oChange) {
				const oNode = oChange.insert[0];
        const oFirstText = oNode.element;
        if(oNewRun)
        {
            if((oNewRun instanceof ParaRun) && oFirstText.firstRun)
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
    };

    CNode.prototype.copyRunWithMockParagraph = function (oRun, mockParagraph, comparison) {
        const oTempParagraph = oRun.Paragraph;
        oRun.Paragraph = mockParagraph;
        const oNewRun = oRun.Copy2(comparison.copyPr);
        oRun.Paragraph = oTempParagraph;

        return oNewRun;
    };

    CNode.prototype.cleanStartOfInsertDifferentRun = function (aContentToInsert, posOfLastInsertRun, oChange, comparison) {
				const oFirstNode = oChange.insert[0];
				const oLastNode = oChange.insert[oChange.insert.length - 1];
	    const oFirstText = oFirstNode.element;
			const oFirstElement = oFirstNode.getContentElement(true);
	    const oLastText = oLastNode.element;
			const oLastElement = oLastNode.getContentElement();
        const applyingParagraph = this.getApplyParagraph(comparison);

        const oParentParagraph =  (this.partner && this.partner.element) || oLastElement.GetParent();
        let k = posOfLastInsertRun;
        let lastCheckRun;
        for(k -= 1; k > -1; --k)
        {
            const oCurRun = oParentParagraph.Content[k];
            // Пока не дошли до первого рана слова, закидываем его на добавление
            if(oCurRun !== oFirstElement)
            {
                this.pushToArrInsertContentWithCopy(aContentToInsert, oCurRun, comparison);
            }
            else
            {
								if (oCurRun instanceof AscCommonWord.ParaRun)
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
								let reviewInfo = oCurRun.GetReviewInfo();
                                this.setCommonReviewTypeWithInfo(lastCheckRun, reviewInfo ? reviewInfo.Copy() : undefined);
                            }
                        }
                    }
                }
								else
								{
									this.pushToArrInsertContentWithCopy(aContentToInsert, oCurRun, comparison);
								}
                break;
            }
        }
    };
	CNode.prototype.checkCommentsFromInsert = function (oChange, comparison)
	{
		const mapRevisedComments = {};
		const arrRevisedComments = [];
		for (let i = 0; i < oChange.insert.length; i += 1)
		{
			const oTextElement = oChange.insert[i].element;
			if (oTextElement.comments)
			{
				for (let sInsertIndex in oTextElement.comments)
				{
					const arrComments = oTextElement.comments[sInsertIndex];
					for (let j = 0; j < arrComments.length; j += 1)
					{
						const oParaComment = arrComments[j].element;
						if (oParaComment.IsCommentStart())
						{
							delete mapRevisedComments[oParaComment.GetCommentId()];
						}
						else
						{
							mapRevisedComments[oParaComment.GetCommentId()] = arrComments[j];
						}
					}
				}
			}
		}
		for (let sCommentId in mapRevisedComments)
		{
			arrRevisedComments.push(mapRevisedComments[sCommentId]);
		}
		let oCheckCommentNode;
		if (oChange.remove.length)
		{
			oCheckCommentNode = oChange.remove[oChange.remove.length - 1];

		}
		else
		{
			oCheckCommentNode = this.children[oChange.anchor.index - 1];
		}
		if (oCheckCommentNode && oCheckCommentNode.element.comments)
		{
			const nLastIndex = oCheckCommentNode.element.elements.length;
			const arrMainComments = oCheckCommentNode.element.comments[nLastIndex];
			comparison.oCommentManager.checkComments(arrMainComments, arrRevisedComments, arrMainComments);
		}
	};
    CNode.prototype.getArrOfInsertsFromChanges = function (oChange, comparison) {
        const aContentToInsert = [];
				this.checkCommentsFromInsert(oChange, comparison);

        if(oChange.insert.length > 0)
        {
	        const oFirstNode = oChange.insert[0];
					const oLastNode = oChange.insert[oChange.insert.length - 1];
            const oFirstText = oFirstNode.getContentElement(true);
            const oLastText = oLastNode.getContentElement();

            const posLastRunOfInsert = this.cleanEndOfInsert(aContentToInsert, oChange, comparison);

            // изменения находятся внутри одного рана или это один и тот же элемент
            if(oFirstText === oLastText)
            {
                this.cleanStartOfInsertSameRun(aContentToInsert[aContentToInsert.length - 1], oChange);
            }
            else
            {
                this.cleanStartOfInsertDifferentRun(aContentToInsert, posLastRunOfInsert, oChange, comparison);
            }

	        const arrStartBookmarks = oFirstNode.getFirstBookmarks();
	        if (arrStartBookmarks)
	        {
						for (let i = arrStartBookmarks.length - 1; i >= 0; i -= 1)
							this.pushToArrInsertContentWithCopy(aContentToInsert, arrStartBookmarks[i], comparison);
	        }
	        const arrStartComments = oFirstNode.getFirstComments();
	        if (arrStartComments)
	        {
		        for (let i = arrStartComments.length - 1; i >= 0; i -= 1)
			        this.pushToArrInsertContentWithCopy(aContentToInsert, arrStartComments[i].element, comparison);
	        }
        }
        return aContentToInsert;
    };


    CNode.prototype.applyInsertsToParagraph = function (comparison, aContentToInsert, oChange) {
        if (oChange.remove.length > 0 && !comparison.isSkipWhitespaces(oChange.remove)) {
            this.applyInsertsToParagraphsWithRemove(comparison, aContentToInsert, oChange);
        } else {
            this.applyInsertsToParagraphsWithoutRemove(comparison, aContentToInsert, oChange);
        }
    };

    CNode.prototype.applyInsertsToParagraphsWithoutRemove = function (comparison, aContentToInsert, oChange) {
        const applyingParagraph = this.getApplyParagraph(comparison);
        if(aContentToInsert.length > 0)
        {
            const index = oChange.anchor.index;
            const oChildNode = this.children[index];
            if(oChildNode)
            {
	            let oFirstText = oChildNode.element;
	            if (oFirstText.lastSwitchElement)
	            {
		            oFirstText = oFirstText.lastSwitchElement;
	            }
							else if (oFirstText instanceof CCommentElement)
							{
								oFirstText = oFirstText.element;
							}

                for(let j = 0; j < applyingParagraph.Content.length; ++j)
                {
                    if(Array.isArray(applyingParagraph.Content))
                    {
                        const oCurRun = applyingParagraph.Content[j];
                        // если совпали ран, после которого нужно вставлять и ран из цикла
                        if(oFirstText === oCurRun)
                        {
                            this.applyInsert(aContentToInsert, [], j, comparison);
                            return true;
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
        return comparison.parentParagraph || this.element;
    };

    CNode.prototype.setCommonReviewTypeWithInfo = function (element, info) {
        element.SetReviewTypeWithInfo(reviewtype_Common, info);
    };

    CNode.prototype.prepareEndOfRemoveChange = function (oChange, comparison, arrSetRemove) {
        const oApplyParagraph = this.getApplyParagraph(comparison);
				const oNode = oChange.remove[oChange.remove.length - 1];
        const oLastText = oNode.element;
	    const oEndOfRemoveRun = oNode.getContentElement();
        let k = oApplyParagraph.Content.length - 1;
        let nInsertPosition = -1;

        for(k; k > -1; --k)
        {
            const oCurRun = oApplyParagraph.Content[k];
            if(oCurRun === oEndOfRemoveRun)
            {
                if(oLastText instanceof CTextElement)
                {
									const arrComments = oLastText.comments[oLastText.elements.length];
									arrSetRemove.push.apply(arrSetRemove, arrComments)

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
						let reviewInfo = oEndOfRemoveRun.GetReviewInfo();
                        this.setCommonReviewTypeWithInfo(oNewRun, reviewInfo ? reviewInfo.Copy() : undefined);
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
    };

    CNode.prototype.setReviewTypeForRemoveChanges = function (comparison, oChange, posLastRunInContent, nInsertPosition, arrSetRemoveReviewType) {
        const oApplyParagraph = this.getApplyParagraph(comparison);
				const oNode = oChange.remove[0];
        const oFirstText = oNode.element;
				const oStartOfRemoveRun = oNode.getContentElement(true);
        let k = posLastRunInContent;
        for(k; k > -1; --k)
        {
            const oChildElement = oApplyParagraph.Content[k];
            if(!(oChildElement === oStartOfRemoveRun))
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
    };

    CNode.prototype.applyInsert = function (arrToInsert, arrToRemove, nInsertPosition, comparison, options) {
			for (let i = 0; i < arrToRemove.length; i += 1) {
            comparison.setRemoveReviewType(arrToRemove[i]);
        }
        this.insertContentAfterRemoveChanges(arrToInsert, nInsertPosition, comparison, options);
    };

    CNode.prototype.applyInsertsToParagraphsWithRemove = function (comparison, aContentToInsert, oChange) {
        const arrSetRemoveReviewType = [];
        const infoAboutEndOfRemoveChange = this.prepareEndOfRemoveChange(oChange, comparison, arrSetRemoveReviewType);
        const posLastRunInContent = infoAboutEndOfRemoveChange.posLastRunInContent;

        let nInsertPosition = infoAboutEndOfRemoveChange.nInsertPosition;
        nInsertPosition = this.setReviewTypeForRemoveChanges(comparison, oChange, posLastRunInContent, nInsertPosition, arrSetRemoveReviewType);
				if (!comparison.options.words && !comparison.isWordsByOneSymbol && oChange.insert.length === 1 && oChange.remove.length === 1) {
					aContentToInsert.reverse();
					arrSetRemoveReviewType.reverse();
					comparison.resolveConflicts(aContentToInsert, arrSetRemoveReviewType, this.getApplyParagraph(comparison), true);
				} else {
					this.applyInsert(aContentToInsert, arrSetRemoveReviewType, nInsertPosition, comparison, {needReverse: true});
				}
    };

    CNode.prototype.insertContentAfterRemoveChanges = function (aContentToInsert, nInsertPosition, comparison) {
        const oElement = this.getApplyParagraph(comparison);
        if(nInsertPosition > -1)
        {
            for (let t = 0; t < aContentToInsert.length; t += 1) {
                if(comparison.isElementForAdd(aContentToInsert[t]))
                {
                    oElement.AddToContent(nInsertPosition, aContentToInsert[t]);
                }
            }
        }
    };

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
    CNode.prototype.equals = function(oNode, oComparison)
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
            return this.privateCompareElements(oNode, true, oComparison);
        }
        return false;
    };

    CNode.prototype.privateCompareElements = function(oNode, bCheckNeighbors, oComparison)
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
                        if(aNeighbors1[0].privateCompareElements(aNeighbors2[0], false, oComparison))
                        {
                            return true;
                        }
                    }
                    if(aNeighbors1[1] && aNeighbors2[1])
                    {
                        if(aNeighbors1[1].privateCompareElements(aNeighbors2[1], false, oComparison))
                        {
                            return true;
                        }
                    }
                    return false;
                }
                else
                {
                    return oElement1.equals(oElement2, bCheckNeighbors, oComparison);
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
            if(this.element instanceof AscWord.Paragraph)
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


	function CCommentElement(oComment, oParaComment)
	{
		this.element = oParaComment;
		if (oComment)
		{
			this.data = oComment.Data;
		}
	}
	CCommentElement.prototype.getQuoteText = function ()
	{
		if (this.data)
		{
			return this.data.Get_QuoteText();
		}
		return '';
	};
	CCommentElement.prototype.getCommentId = function ()
	{
		return this.element.GetCommentId();
	};
	CCommentElement.prototype.getFullMergeComments = function (oAnotherElement)
	{
		const arrMainAnswers = this.getAnswers();
		const arrAnotherAnswers = oAnotherElement.getAnswers();
		const length = Math.min(arrMainAnswers.length, arrAnotherAnswers.length);
		for (let i = 0; i < length; i++)
		{
			const sMainAnswer = arrMainAnswers[i].Get_Text();
			const sAnotherAnswer = arrAnotherAnswers[i].Get_Text();
			if (sMainAnswer !== sAnotherAnswer)
			{
				return null;
			}
		}
		return Math.abs(arrMainAnswers.length - arrAnotherAnswers.length);
	};
	CCommentElement.prototype.getAnswerMap = function ()
	{
		const oAnswerMap = {};
		const arrAnswers = this.getAnswers();
		for (let i = 0; i < arrAnswers.length; i += 1)
		{
			const oAnswer = arrAnswers[i];
			const sText = oAnswer.Get_Text();
			if (!oAnswerMap[sText])
			{
				oAnswerMap[sText] = [];
			}
			oAnswerMap[sText].push(oAnswer);
		}
		return oAnswerMap;
	};
	CCommentElement.prototype.getPartMergeComments = function (oAnotherElement)
	{
		let nDifference = 0;
		const oMainAnswerMap = this.getAnswerMap();
		const oRevisedAnswerMap = oAnotherElement.getAnswerMap();
		let arrData = [];
		for (let sCommentText in oMainAnswerMap)
		{
			const arrMainAnswers = oMainAnswerMap[sCommentText];
			if (oRevisedAnswerMap[sCommentText])
			{
				const nDeleteCount = oRevisedAnswerMap[sCommentText].length;
				nDifference += Math.abs(arrMainAnswers.length - nDeleteCount);
				arrMainAnswers.splice(0, nDeleteCount);
				arrData = arrData.concat(arrMainAnswers);
			}
			else
			{
				arrData = arrData.concat(arrMainAnswers);
				nDifference += arrMainAnswers.length;
			}
		}
		for (let sCommentText in oRevisedAnswerMap)
		{
			if (!oMainAnswerMap[sCommentText])
			{
				nDifference += oRevisedAnswerMap[sCommentText].length;
			}
		}
		arrData.sort(function (a, b)
		{
			return parseInt(a.m_sOOTime, 10) - parseInt(b.m_sOOTime, 10);
		});
		return {arrData: arrData, difference: nDifference};
	};

	CCommentElement.prototype.getDifference = function (oAnotherElement)
	{
		const sText = this.getText();
		const sAnotherText = oAnotherElement.getText();
		if (sText !== sAnotherText)
		{
			return null;
		}
		const oRes = {mainElement: this, anotherElement: oAnotherElement, quoteDifference: this.getQuoteDifference(oAnotherElement)};
		const nCommentDifference = this.getFullMergeComments(oAnotherElement);
		if (nCommentDifference !== null)
		{
			oRes.commentDifference = nCommentDifference;
			return oRes;
		}
		const oPartDifference = this.getPartMergeComments(oAnotherElement);
		oRes.commentDifference = oPartDifference.difference;
		oRes.arrData = oPartDifference.arrData;
		return oRes;
	};
	CCommentElement.prototype.getQuoteDifference = function (oAnotherElement)
	{
		const sMainQuoteText = this.getQuoteText();
		const sAnotherQuoteText = oAnotherElement.getQuoteText();

		return Math.abs(sMainQuoteText.length - sAnotherQuoteText.length);
	};
	CCommentElement.prototype.getText = function ()
	{
		if (this.data)
		{
			return this.data.Get_Text();
		}
		return '';
	};
	CCommentElement.prototype.getAnswers = function ()
	{
		if (this.data)
		{
			return this.data.m_aReplies;
		}
		return [];
	};
    function CTextElement()
    {
        this.elements = [];
        this.firstRun = null;
        this.lastRun = null;
				this.bookmarks = {};
				this.comments = {};
	    this.lastSwitchElement = null;
    }
	CTextElement.prototype.isWordBeginWith = function (oOther)
	{
		if (this.elements.length < oOther.elements.length) {
			return false;
		}

		for (let i = 0; i < oOther.elements.length; i += 1)
		{
			const oMainElement = this.elements[i];
			const oSecondaryElement = oOther.elements[i];
			if (typeof oMainElement.Value !== 'number' || typeof oSecondaryElement.Value !== 'number')
			{
				return false;
			} else if (oMainElement.Value !== oSecondaryElement.Value) {
				return false;
			}
		}
		return true;
	};
	CTextElement.prototype.isWordEndWith = function (oOther)
	{
		if (this.elements.length < oOther.elements.length) {
			return false;
		}

		for (let i = 0; i < oOther.elements.length; i += 1)
		{
			const oMainElement = this.elements[this.elements.length - 1 - i];
			const oSecondaryElement = oOther.elements[oOther.elements.length - 1 - i];
			if (typeof oMainElement.Value !== 'number' || typeof oSecondaryElement.Value !== 'number')
			{
				return false;
			} else if (oMainElement.Value !== oSecondaryElement.Value) {
				return false;
			}
		}
		return true;
	};
	CTextElement.prototype.getSortedInsertIndexesFromMap = function (arrLabels)
	{
		return (Object.keys(arrLabels).map(function (e) {return parseInt(e, 10);})
			.sort(function (a, b) {return b - a;}));
	};
	CTextElement.prototype.compareReviewElements = function (oAnotherElement)
	{
		return true;
	};
	CTextElement.prototype.getBookmarkInsertIndexes = function ()
	{
		return this.getSortedInsertIndexesFromMap(this.bookmarks);
	};
	CTextElement.prototype.getCommentInsertIndexes = function ()
	{
		return this.getSortedInsertIndexesFromMap(this.comments);
	};
	CTextElement.prototype.getLastBookmarks = function ()
	{
		return this.bookmarks[this.elements.length];
	};
	CTextElement.prototype.getFirstBookmarks = function ()
	{
		return this.bookmarks[0];
	};
	CTextElement.prototype.getLastComments = function ()
	{
		return this.comments[this.elements.length];
	};
	CTextElement.prototype.getFirstComments = function ()
	{
		return this.comments[0];
	};
	CTextElement.prototype.addBookmark = function (oBookmark, nInsertIndex)
	{
		if (!this.bookmarks[nInsertIndex])
		{
			this.bookmarks[nInsertIndex] = [];
		}
		this.bookmarks[nInsertIndex].push(oBookmark);
	};
	CTextElement.prototype.addComment = function (oComment, nInsertIndex)
	{
		if (!this.comments[nInsertIndex])
		{
			this.comments[nInsertIndex] = [];
		}
		this.comments[nInsertIndex].push(oComment);
	};
    CTextElement.prototype.getPosOfStart = function () {
        const startElement = this.elements[0];
        return this.firstRun.GetElementPosition(startElement);
    };

    CTextElement.prototype.addToElements = function (element, options) {
        this.elements.push(element);
    };

    CTextElement.prototype.getPosOfEnd = function () {
        const endElement = this.elements[this.elements.length - 1];
        return this.lastRun.GetElementPosition(endElement);
    };

    CTextElement.prototype.getElement = function (idx) {
        return this.elements[idx];
    };
	CTextElement.prototype.checkRemoveReviewType = function () {
		return false;
	};
	CTextElement.prototype.checkUncompletedWords = function (other) {
		if (this.elements.length === other.elements.length) {
			return false;
		}
		let oMainTextElement;
		let oSecondaryTextElement;
		if (this.elements.length > other.elements.length) {
			oMainTextElement = this;
			oSecondaryTextElement = other;
		}
		else {
			oMainTextElement = other;
			oSecondaryTextElement = this;
		}
		const nEndIndex = oSecondaryTextElement.elements.length - 1;
		const oEndElement = oMainTextElement.elements[nEndIndex];
		if (oMainTextElement.isWordBeginWith(oSecondaryTextElement)) {
			return oMainTextElement.checkRemoveReviewType(nEndIndex) || AscCommon.isEastAsianScript(oEndElement.Value);
		}

		const nStartIndex = oMainTextElement.elements.length - oSecondaryTextElement.elements.length;
		const oStartElement = oMainTextElement.elements[nStartIndex];
		if (oMainTextElement.isWordEndWith(oSecondaryTextElement)) {
			return oMainTextElement.checkRemoveReviewType(nStartIndex) || AscCommon.isEastAsianScript(oStartElement.Value);
		}
		return false;
	};
    CTextElement.prototype._equals = function (other, bNeedCheckReview, oComparison)
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
							if (!oComparison.options.caseChanges) {
								if (String.fromCharCode(oElement.Value).toLowerCase() !== String.fromCharCode(oOtherElement.Value).toLowerCase()) {
									return false;
								}
							}
							else if(oElement.Value !== oOtherElement.Value)
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
	CTextElement.prototype.equals = function (other, bNeedCheckReview, oComparison) {
		return this._equals(other, bNeedCheckReview, oComparison) || this.checkUncompletedWords(other);
	};

    CTextElement.prototype.updateHash = function(oHash, oComparison){
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
            oHash.update(aCheckArray, oComparison);
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
	    this.textBoxes = true;
	    this.tables = true;
	    this.words = true;
	    this.headersAndFooters = true;
	    this.footNotes  = true;
	    this.caseChanges = true;
			this.formatting = true;
			this.whiteSpace = true;
	    this.comments = true;
			
	    this.insertionsAndDeletions = true;
	    // this.moves = true;
	    // this.fields = true;
    }
    ComparisonOptions.prototype["getInsertionsAndDeletions"] = ComparisonOptions.prototype.getInsertionsAndDeletions = function(){return this.insertionsAndDeletions;};
    // ComparisonOptions.prototype["getMoves"] = ComparisonOptions.prototype.getMoves = function(){return this.moves;};
    ComparisonOptions.prototype["getComments"] = ComparisonOptions.prototype.getComments = function(){return this.comments;};
    ComparisonOptions.prototype["getFormatting"] = ComparisonOptions.prototype.getFormatting = function(){return this.formatting;};
    ComparisonOptions.prototype["getCaseChanges"] = ComparisonOptions.prototype.getCaseChanges = function(){return this.caseChanges;};
    ComparisonOptions.prototype["getWhiteSpace"] = ComparisonOptions.prototype.getWhiteSpace = function(){return this.whiteSpace;};
    ComparisonOptions.prototype["getTables"] = ComparisonOptions.prototype.getTables = function(){return this.tables;};
    ComparisonOptions.prototype["getHeadersAndFooters"] = ComparisonOptions.prototype.getHeadersAndFooters = function(){return this.headersAndFooters;};
    ComparisonOptions.prototype["getFootNotes"] = ComparisonOptions.prototype.getFootNotes = function(){return this.footNotes;};
    ComparisonOptions.prototype["getTextBoxes"] = ComparisonOptions.prototype.getTextBoxes = function(){return this.textBoxes;};
    // ComparisonOptions.prototype["getFields"] = ComparisonOptions.prototype.getFields = function(){return this.fields;};
    ComparisonOptions.prototype["getWords"] = ComparisonOptions.prototype.getWords = function(){return this.words;};


    ComparisonOptions.prototype["putInsertionsAndDeletions"] = ComparisonOptions.prototype.putInsertionsAndDeletions = function(v){this.insertionsAndDeletions = v;};
    // ComparisonOptions.prototype["putMoves"] = ComparisonOptions.prototype.putMoves = function(v){this.moves = v;};
    ComparisonOptions.prototype["putComments"] = ComparisonOptions.prototype.putComments = function(v){this.comments = v;};
    ComparisonOptions.prototype["putFormatting"] = ComparisonOptions.prototype.putFormatting = function(v){this.formatting = v;};
    ComparisonOptions.prototype["putCaseChanges"] = ComparisonOptions.prototype.putCaseChanges = function(v){this.caseChanges = v;};
    ComparisonOptions.prototype["putWhiteSpace"] = ComparisonOptions.prototype.putWhiteSpace = function(v){this.whiteSpace = v;};
    ComparisonOptions.prototype["putTables"] = ComparisonOptions.prototype.putTables = function(v){this.tables = v;};
    ComparisonOptions.prototype["putHeadersAndFooters"] = ComparisonOptions.prototype.putHeadersAndFooters = function(v){this.headersAndFooters = v;};
    ComparisonOptions.prototype["putFootNotes"] = ComparisonOptions.prototype.putFootNotes = function(v){this.footNotes = v;};
    ComparisonOptions.prototype["putTextBoxes"] = ComparisonOptions.prototype.putTextBoxes = function(v){this.textBoxes = v;};
    // ComparisonOptions.prototype["putFields"] = ComparisonOptions.prototype.putFields = function(v){this.fields = v;};
    ComparisonOptions.prototype["putWords"] = ComparisonOptions.prototype.putWords = function(v){this.words = v;};


    function CDocumentComparison(oOriginalDocument, oRevisedDocument, oOptions, bIsWordsByOneSymbol)
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
            Comparison: this,
	        SkipFootnoteReference: !oOptions.footNotes
				};
	    this.firstCheckNumId = null;

        this.oComparisonMoveMarkManager = new CMoveMarkComparisonManager();
		this.oBookmarkManager = new CComparisonBookmarkManager(oOriginalDocument, oRevisedDocument);
		this.oCommentManager = new CComparisonCommentManager(this);

		this.isWordsByOneSymbol = bIsWordsByOneSymbol;

	    this.needCheckReview = false;
	    this.needCopyForResolveEqualWords = true;
			this.skipCopiedElements = false;

	    this.nInsertChangesType = reviewtype_Add;
	    this.nRemoveChangesType = reviewtype_Remove;
    }
		CDocumentComparison.prototype.skipCommentElementOnCopyParagraph = function (oParaComment) {
			if (this.options.comments) {
				return false;
			}
			return !this.oCommentManager.savedParaComments[oParaComment.Id];
		}
	CDocumentComparison.prototype.executeResolveConflictMode = function (oResolveConflictComparison, fCallback) {
		const oOldCommentsMeeting = this.oCommentManager.mapCommentMeeting;
		this.oCommentManager.mapCommentMeeting = {};
		oResolveConflictComparison.oCommentManager = this.oCommentManager;

		const oOldBookmarkMeeting = this.oBookmarkManager.mapBookmarkMeeting;
		this.oBookmarkManager.mapBookmarkMeeting = {};
		oResolveConflictComparison.oBookmarkManager = this.oBookmarkManager;
		oResolveConflictComparison.oComparisonMoveMarkManager = this.oComparisonMoveMarkManager;
		oResolveConflictComparison.CommentsMap = this.CommentsMap;
		oResolveConflictComparison.StylesMap = this.StylesMap;
		fCallback();
		this.oBookmarkManager.mapBookmarkMeeting = oOldBookmarkMeeting;
		this.oCommentManager.mapCommentMeeting = oOldCommentsMeeting;
	};
		CDocumentComparison.prototype.resolveConflicts = function(arrToInserts, arrToRemove, applyParagraph, bIsWordsByOneSymbol) {
			if (arrToInserts.length === 0 || arrToRemove.length === 0) return;
			bIsWordsByOneSymbol = bIsWordsByOneSymbol || this.isWordsByOneSymbol;
			let comparison;
			if (this instanceof AscCommonWord.CDocumentResolveConflictComparison) {
				comparison = new AscCommonWord.CDocumentMergeComparison(this.originalDocument, this.revisedDocument, this.options, bIsWordsByOneSymbol);
			} else {

				comparison = new AscCommonWord.CDocumentResolveConflictComparison(this.originalDocument, this.revisedDocument, this.options, bIsWordsByOneSymbol);
			}

			this.executeResolveConflictMode(comparison, function () {
				const originalDocument = getMockDocumentFromRuns(arrToRemove);
				const revisedDocument = getMockDocumentFromRuns(arrToInserts);
				comparison.parentParagraph = applyParagraph;
				comparison.oComparisonMoveMarkManager.executeResolveConflictMode(function () {
					comparison.compareRoots(originalDocument, revisedDocument);
				});
			});
		};
    CDocumentComparison.prototype.checkOriginalAndSplitRun = function (oOriginalRun, oSplitRun) {

    };
	CDocumentComparison.prototype.removeCommentsFromMap = function ()
	{
		for (let sCommentId in this.oCommentManager.mapDelete)
		{
			this.originalDocument.RemoveComment(sCommentId, true);
		}
		this.oCommentManager.mapDelete = {};
	};
	CDocumentComparison.prototype.executeWithSkipCopiedElements = function (fCallback) {
		const bOldSkipCheckCopyRuns = this.skipCopiedElements;
		this.skipCopiedElements = true;
		fCallback();
		this.skipCopiedElements = bOldSkipCheckCopyRuns;
	};
	CDocumentComparison.prototype.checkCopyParaRunMoveMarks = function (oNewRun, oOldRun) {
		const sMoveName = this.oComparisonMoveMarkManager.getMoveMarkNameByRun(oOldRun);
		this.oComparisonMoveMarkManager.addRunMoveMarkNameRelation(sMoveName, oNewRun);
		const nMoveReviewType = oOldRun.GetReviewMoveType();
		if (AscFormat.isRealNumber(nMoveReviewType) && nMoveReviewType !== Asc.c_oAscRevisionsMove.NoMove)
		{
			this.oComparisonMoveMarkManager.addMoveMarkNameRunRelation(sMoveName, oNewRun);
		}
	}
	CDocumentComparison.prototype.checkCopyParaRun = function (oNewRun, oOldRun) {
		this.checkCopyParaRunMoveMarks(oNewRun, oOldRun);
		this.checkReviewInfoOfCopiedElements(oNewRun, oOldRun);
	}
    CDocumentComparison.prototype.checkReviewInfoOfCopiedElements = function (oNewRun, oOldRun) {
			if (this.skipCopiedElements) {
				return;
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
    };
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
    CDocumentComparison.prototype.getLCSCallback = function (bOrig) {
        const oThis = this;
        return function(x, y) {
            const oOrigNode = x;
            const oReviseNode = y;
            const oDiff  = new AscCommon.Diff(oOrigNode, oReviseNode);
            oDiff.equals = function(a, b)
            {
                return a.equals(b, oThis);
            };
            const oMatching = new CMatching();
            oDiff.matchTrees(oMatching);
            const oDeltaCollector = new AscCommon.DeltaCollector(oMatching, oOrigNode, oReviseNode);
            oDeltaCollector.forEachChange(oThis.forEachChangeCallback);
            oThis.compareDrawingObjectsFromMatching(oMatching, bOrig);
            oThis.applyChangesToChildNode(oOrigNode);
            oThis.compareNotes(oMatching);
        };
    };

    CDocumentComparison.prototype.forEachChangeCallback = function(oOperation) {
        oOperation.anchor.base.addChange(oOperation);
    };
	CDocumentComparison.prototype.compareElementsArray = function(aBase, aCompare, bOrig) {
		if (!aBase.length || !aCompare.length) {
			return;
		}
		const arrComparedSums = [0];
		for (let i = 0; i < aCompare.length; i++) {
			arrComparedSums.push(arrComparedSums[i] + aCompare[i].wordCounter.count + 1);
		}

		const arrBaseSums = [0];
		for (let i = 0; i < aBase.length; i++) {
			arrBaseSums.push(arrBaseSums[i] + aBase[i].wordCounter.count + 1);
		}

		const oResult = {};
		const oFirstRow = {};
		oResult[0] = oFirstRow;
		oFirstRow[0] = {diff: aCompare[0].getDiff(aBase[0], this), baseIndex: 0, compareIndex: 0};
		for (let i = 1; i < aCompare.length; i += 1) {
			const nCompareSum = arrComparedSums[i];
			const nTotalDiff = nCompareSum + aCompare[i].getDiff(aBase[0], this);
			if (oFirstRow[i - 1].diff >= nTotalDiff) {
				oFirstRow[i] = {diff: nTotalDiff, baseIndex: 0, compareIndex: i};
			} else {
				oFirstRow[i] = oFirstRow[i - 1];
			}
		}
		for (let i = 1; i < aBase.length; i += 1) {
			const oPreviousRow = oResult[i - 1];
			const oCurrentRow = {};
			oResult[i] = oCurrentRow;
			const nCurTotalDiff = aBase[i].getDiff(aCompare[0], this) + arrBaseSums[i];

			const oPreviousCell = oPreviousRow[0];
			const nPreviousTotalDiff = oPreviousCell.diff + arrBaseSums[i + 1] - arrBaseSums[oPreviousCell.baseIndex + 1];

			if (nPreviousTotalDiff >= nCurTotalDiff) {
				oCurrentRow[0] = {diff: nCurTotalDiff, compareIndex: 0, baseIndex: i};
			} else {
				oCurrentRow[0] = oPreviousCell;
			}
		}
		for (let i = 1; i < aBase.length; i += 1) {
			const oBaseElem = aBase[i];
			const oCurrentBestResult = oResult[i];
			const oPreviousBestResult = oResult[i - 1];
			for (let j = 1; j < aCompare.length; j++) {
				const oCompareElem = aCompare[j];

				const oPrevCurrentBest = oCurrentBestResult[j - 1];
				const baseSum = arrBaseSums[i + 1] - arrBaseSums[oPrevCurrentBest.baseIndex + 1];
				const compareSum = arrComparedSums[j + 1] - arrComparedSums[oPrevCurrentBest.compareIndex + 1];
				const nPrevColumnCurDiff = baseSum + compareSum + oPrevCurrentBest.diff;

				const oCurPreviousBest = oPreviousBestResult[j];
				const nCurBaseSum = arrBaseSums[i + 1] - arrBaseSums[oCurPreviousBest.baseIndex + 1];
				const nCurCompareSum = arrComparedSums[j + 1] - arrComparedSums[oCurPreviousBest.compareIndex + 1];
				const nPrevColumnPrevDiff = nCurBaseSum + nCurCompareSum + oCurPreviousBest.diff;

				const oPrevPreviousBest = oPreviousBestResult[j - 1];
				const nPrevBaseSum = arrBaseSums[i] - arrBaseSums[oPrevPreviousBest.baseIndex + 1];
				const nPrevCompareSum = arrComparedSums[j] - arrComparedSums[oPrevPreviousBest.compareIndex + 1];
				const nDiff = oBaseElem.getDiff(oCompareElem, this);
				const nCurColumnPrevDiff = nPrevBaseSum + nPrevCompareSum + oPrevPreviousBest.diff + nDiff;

				if (nCurColumnPrevDiff <= nPrevColumnPrevDiff && nCurColumnPrevDiff <= nPrevColumnCurDiff) {
					oCurrentBestResult[j] = {diff: nCurColumnPrevDiff, baseIndex: i, compareIndex: j};
				} else if (nPrevColumnCurDiff <= nCurColumnPrevDiff && nPrevColumnCurDiff <= nPrevColumnPrevDiff) {
					oCurrentBestResult[j] = oPrevCurrentBest;
				} else {
					oCurrentBestResult[j] = oCurPreviousBest;
				}
			}
		}

		const fLCSCallback = this.getLCSCallback(bOrig);
		let curBaseIndex = aBase.length - 1;
		let curCompareIndex = aCompare.length - 1;
		while (curBaseIndex >= 0 && curCompareIndex >= 0) {
			const curResultObject = oResult[curBaseIndex][curCompareIndex];
			const oBaseElement = aBase[curResultObject.baseIndex];
			const oCompareElement = aCompare[curResultObject.compareIndex];
			if (oBaseElement.equals(oCompareElement)) {
				oBaseElement.setPartner(oCompareElement);
				if (bOrig) {
					fLCSCallback(oBaseElement, oCompareElement);
				} else {
					fLCSCallback(oCompareElement, oBaseElement);
				}
			}
			curBaseIndex = curResultObject.baseIndex - 1;
			curCompareIndex = curResultObject.compareIndex - 1;
		}
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
	CDocumentComparison.prototype.insertCopyTextBoxContent = function (oBaseShape, oTextBoxContent) {
		oBaseShape.setTextBoxContent(oTextBoxContent.Copy(oBaseShape, editor.WordControl.m_oDrawingDocument, this.copyPr));
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
					this.insertCopyTextBoxContent(oBaseShape, oCompareShape.textBoxContent);
        }
    };
	CDocumentComparison.prototype.updateBookmarksStack = function (oNode)
	{
		this.oBookmarkManager.updateBookmarksStack(oNode);
	};
	CDocumentComparison.prototype.updateCommentsStack = function (oNode)
	{
		this.oCommentManager.updateCommentsStack(oNode);
	};
	CDocumentComparison.prototype.createNode = function (oElement, oParent)
	{
		const CNodeConstructor = this.getNodeConstructor();
		const oNode = new CNodeConstructor(oElement, oParent);
		if (oElement instanceof CTextElement)
		{
			this.updateBookmarksStack(oNode);
			this.updateCommentsStack(oNode);
			this.oBookmarkManager.previousNode = oNode;
			this.oCommentManager.previousNode = oNode;
		}
		return oNode;
	};
    CDocumentComparison.prototype.compareGroups = function(oBaseGroup, oCompareGroup)
    {
        const NodeConstructor = this.getNodeConstructor();
        const oLCS = new AscCommon.LCS(oBaseGroup.spTree, oCompareGroup.spTree);
        oLCS.equals = function(a, b) {
            return a.isComparable(b);
        };

        const oBaseNode = this.createNode(oBaseGroup, null);
        for(let nSp = 0; nSp < oBaseGroup.spTree.length; ++nSp)
        {
            this.createNode(oBaseGroup.spTree[nSp], oBaseNode);
        }
        const oCompareNode = this.createNode(oCompareGroup, null);
        for(let nSp = 0; nSp < oCompareGroup.spTree.length; ++nSp)
        {
            this.createNode(oCompareGroup.spTree[nSp], oCompareNode);
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
    };

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
        if(!this.options.textBoxes || !oBaseGrObject || !oCompareGrObject)
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
	    this.oCommentManager.generateSavedParaComments();
			const bOldSkipFootnoteReference = this.copyPr.SkipFootnoteReference;
	    this.copyPr.SkipFootnoteReference = false;
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
	    this.copyPr.SkipFootnoteReference = bOldSkipFootnoteReference;
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

        this.compareElementsArray(aBase, aCompare, bOrig);
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
				this.oBookmarkManager.init(oOriginalDocument, oRevisedDocument);
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
            const oldTrackRevisions = oOriginalDocument.GetLocalTrackRevisions();
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
            oThis.compareSectPr(oOriginalDocument, oRevisedDocument, !oThis.options.headersAndFooters);

            const oFonts = oOriginalDocument.Document_Get_AllFontNames();
            const aFonts = [];
            for (let i in oFonts)
            {
                if(oFonts.hasOwnProperty(i))
                {
                    aFonts[aFonts.length] = new AscFonts.CFont(i);
                }
            }
            oApi.pre_Paste(aFonts, oImageMap, function()
            {
							oThis.removeCommentsFromMap();
                oOriginalDocument.SetTrackRevisions(oldTrackRevisions);
                oOriginalDocument.End_SilentMode(false);
								if (oThis.oBookmarkManager.needUpdateBookmarks)
								{
									oOriginalDocument.UpdateBookmarks();
								}
								oThis.updateCommentsQuoteText();
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
	CDocumentComparison.prototype.getNewParaPrWithDiff = function (oElementPr, oPartnerPr)
	{
		const oOldParaPr = oElementPr.Copy(undefined, undefined);

		this.firstCheckNumId = oOldParaPr.NumPr && oOldParaPr.NumPr.NumId;
		const oNewParaPr = oPartnerPr.Copy(undefined, this.copyPr);
		this.firstCheckNumId = null;

		if (oOldParaPr.Is_Equal(oNewParaPr))
		{
			return null;
		}
		oNewParaPr.PrChange = oOldParaPr;
		oNewParaPr.ReviewInfo = this.getReviewInfo();
		return oNewParaPr;
	};
    CDocumentComparison.prototype.isElementForAdd = function (oElement)
    {
        return !(oElement.IsParaEndRun && oElement.IsParaEndRun() || this.oBookmarkManager.isSkip(oElement));
    };
    CDocumentComparison.prototype.executeWithCheckInsertAndRemove = function (callback, oChange) {
			this.copyPr.SkipUpdateInfo = !this.options.words && oChange.remove.length === 1 && oChange.insert.length === 1;
			callback();
			this.copyPr.SkipUpdateInfo = false;
    };

	CDocumentComparison.prototype.applyResolveTypes = function () {};

	CDocumentComparison.prototype.isSkipWhitespaces = function(arrNodes) {
		if (this.options.whiteSpace) {
			return false;
		}

		for (let i = 0; i < arrNodes.length; i++) {
			const oNode = arrNodes[i];
			if (!oNode.isSpaceText()) {
				return false;

			}
		}
		return true;
	};
	CNode.prototype.updateEqualNode = function (comparison, index) {
		const oPartnerNode = this.partner;
		const oOriginalTextElement = this.element;
		const oPartnerTextElement = oPartnerNode.element;
		if (oPartnerTextElement.elements.length === oOriginalTextElement.elements.length) {
			this.resolveTypesWithPartner(comparison);
		} else {
			const oParentNode = this.par;
			const oChange = new AscCommon.AttachedOperation(
				new AscCommon.Anchor(oParentNode, oParentNode, index),
				AscCommon.UPDATE_FOREST_TYPE, [index],
				[this], [oPartnerNode]);
			const bOldComparisonWordOption = comparison.options.words;
			comparison.options.words = false;
			comparison.copyPr.SkipUpdateInfo = true;
			const aContentToInsert = comparison.isSkipWhitespaces(oChange.insert) ? [] : oParentNode.getArrOfInsertsFromChanges(oChange, comparison); // todo: check skip on symbol comparing
			oParentNode.applyInsertsToParagraph(comparison, aContentToInsert, oChange);
			comparison.options.words = bOldComparisonWordOption;
			comparison.copyPr.SkipUpdateInfo = false;
		}
	};
	CDocumentComparison.prototype.applyChangesToParagraph = function(oNode)
	{
		oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
		let nLastIndex = oNode.children.length - 1;
		for(let i = 0; i < oNode.changes.length; ++i)
		{
			const oChange = oNode.changes[i];
			let arrResult = [];
			for (let j = nLastIndex; j >= oChange.anchor.index; j -= 1)
			{
				const oChild = oNode.children[j];
				if (oChild.partner && oChild.element instanceof CTextElement) {
					if (oChild.element.checkUncompletedWords(oChild.partner.element)) {
						resolveTypesWithPartner(arrResult, this);
						arrResult = [];
						oChild.updateEqualNode(this, j);
					} else {
						arrResult.unshift(oChild);
					}
				} else if (arrResult.length) {
					resolveTypesWithPartner(arrResult, this);
					arrResult = [];
				}
			}
			if (arrResult.length) {
				resolveTypesWithPartner(arrResult, this);
			}
			this.executeWithCheckInsertAndRemove(function () {
				const aContentToInsert = this.isSkipWhitespaces(oChange.insert) ? [] : oNode.getArrOfInsertsFromChanges(oChange, this);
				//handle removed elements
				oNode.applyInsertsToParagraph(this, aContentToInsert, oChange);
			}.bind(this), oChange);

			nLastIndex = oChange.anchor.index - 1;
		}
		let arrResult = [];
		for (let j = nLastIndex; j >= 0; j -= 1)
		{
			const oChild = oNode.children[j];
			if (oChild.partner && oChild.element instanceof CTextElement) {
				if (oChild.element.checkUncompletedWords(oChild.partner.element)) {
					resolveTypesWithPartner(arrResult, this);
					arrResult = [];
					oChild.updateEqualNode(this, j);
				} else {
					arrResult.unshift(oChild);
				}
			} else if (arrResult.length) {
				resolveTypesWithPartner(arrResult, this);
				arrResult = [];
			}
		}
		if (arrResult.length) {
			resolveTypesWithPartner(arrResult, this);
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
    };

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
    };

    CDocumentComparison.prototype.compareSectPr = function(oElement, oPartnerElement, bSkipCheckHeadersFooters)
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
					if (!bSkipCheckHeadersFooters) {
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
	CDocumentComparison.prototype.applyChangesToTablePr = function (oNode) {
		if (oNode.partner) {
			const oMainTable = oNode.element;
			const oRevisedTable = oNode.partner.element;

			const oMainPr = oMainTable.Pr;
			const oRevisedPr = oRevisedTable.Pr;

			const sMainStyleId = oMainTable.TableStyle;
			const sRevisedStyleId  = this.copyStyleById(oRevisedTable.TableStyle);

			const oMainTableLook = oMainTable.Get_TableLook();
			const oRevisedTableLook = oMainTable.Get_TableLook();
			if (oMainPr.Is_Equal(oRevisedPr) && sMainStyleId === sRevisedStyleId && (oMainTableLook && oMainTableLook.IsEqual(oRevisedTableLook) || oMainTableLook === oRevisedTableLook)) {
				return;
			}
			if (sMainStyleId !== sRevisedStyleId) {
				oMainTable.Set_TableStyle(sRevisedStyleId, true);
			}
			if (oRevisedTableLook) {
				oMainTable.Set_TableLook(oRevisedTableLook.Copy());
			} else {
				oMainTable.Set_TableLook(oRevisedTableLook);
			}
			const oNewPr = oRevisedPr.Copy();
			oNewPr.PrChange = oMainPr.Copy();
			oNewPr.ReviewInfo = this.getReviewInfo();
			oMainTable.Set_Pr(oNewPr);
		}
	}
    CDocumentComparison.prototype.applyChangesToTable = function(oNode)
    {
			this.applyChangesToTablePr(oNode);
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
	CDocumentComparison.prototype.addTablePrChange = function (oTablePr) {
		oTablePr.PrChange = oTablePr.Copy(false);
		oTablePr.ReviewInfo = this.getReviewInfo();
	};
		CDocumentComparison.prototype.applyChangesToTableCellPr = function (oNode) {
			if (oNode.partner) {
				const oMainCell = oNode.element.GetParent();
				const oRevisedCell = oNode.partner.element.GetParent();
				this.setTableElementChange(oMainCell, oRevisedCell);
			}
		};
		CDocumentComparison.prototype.applyChangesToTableRowPr = function (oNode) {
			if (oNode.partner) {
				const oMainRow = oNode.element;
				const oRevisedRow = oNode.partner.element;
				this.setTableElementChange(oMainRow, oRevisedRow);
			}
		};
	CDocumentComparison.prototype.setTableElementChange = function (oMainElement, oRevisedElement) {
		const oMainPr = oMainElement.Pr;
		const oRevisedPr = oRevisedElement.Pr;

		if (oMainPr.Is_Equal(oRevisedPr)) {
			return;
		}
		const oNewPr = oRevisedPr.Copy();
		oNewPr.PrChange = oMainPr.Copy();
		oNewPr.ReviewInfo = this.getReviewInfo();
		oMainElement.Set_Pr(oNewPr);
		const oTable = oMainElement.GetTable();
		oTable.AddPrChange(this.copyPr);
	};
    CDocumentComparison.prototype.applyChangesToTableRow = function(oNode)
    {
        this.checkRowReview(oNode);
				this.applyChangesToTableRowPr(oNode);
        //TODO: handle cell inserts and removes

        for(let i = 0; i < oNode.children.length; ++i)
        {
					const oChild = oNode.children[i];
					this.applyChangesToTableCellPr(oChild);
          this.applyChangesToDocContent(oChild);
        }
    };
	CDocumentComparison.prototype.getCopyNumId = function (sNumId)
	{
		if (this.matchedNums[sNumId])
		{
			return this.matchedNums[sNumId];
		}
		const sCopyNumId = this.revisedDocument.CopyNumberingMap[sNumId];
		if (typeof this.firstCheckNumId === 'string' && !this.checkedNums[this.firstCheckNumId] && sCopyNumId)
		{
			const oCopyNum = AscCommon.g_oTableId.Get_ById(sCopyNumId);
			const oOrigNumbering = this.originalDocument.Numbering.Num;
			if (oCopyNum && oOrigNumbering)
			{
				const oOrigNum = AscCommon.g_oTableId.Get_ById(this.firstCheckNumId);
				if (oOrigNum && oOrigNum.IsEqual(oCopyNum))
				{
					this.matchedNums[sNumId] = this.firstCheckNumId;
					this.checkedNums[this.firstCheckNumId] = true;
					return this.firstCheckNumId;
				}
			}
		}
		return sCopyNumId === undefined ? sNumId : sCopyNumId;
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
        if(!this.CommentsMap[sId])
        {
            const oComment = this.revisedDocument.Comments.Get_ById(sId);
            if(oComment)
            {
                const oOrigComments = this.originalDocument.Comments;
                if(oOrigComments)
                {
                    const oOldParent = oComment.Parent;
                    oComment.Parent = oOrigComments;
                    const oCopyComment = oComment.Copy();
                    this.CommentsMap[sId] = oCopyComment;
                    this.CommentsMap[oCopyComment.GetId()] = oCopyComment;
										if (this.oCommentManager.mapLink[sId])
										{
											this.oCommentManager.addToLink(oCopyComment.GetId(), sId);
										}
                    this.originalDocument.Comments.Add(oCopyComment);
                    this.api.sync_AddComment(oCopyComment.Id, oCopyComment.Data);
                    oComment.Parent = oOldParent;
                }
            }
        }
        return this.CommentsMap[sId];
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
    CDocumentComparison.prototype.getReviewInfo = function()
    {
		let oReviewInfo = new AscWord.ReviewInfo();
        oReviewInfo.Editor   = this.api;
        oReviewInfo.UserId   = "";
        oReviewInfo.MoveType = Asc.c_oAscRevisionsMove.NoMove;
        oReviewInfo.PrevType = -1;
        oReviewInfo.PrevInfo = null;
        oReviewInfo.UserName = this.getUserName();
        const oCore = this.revisedDocument.Core;
        if(oCore && (oCore.modified instanceof Date))
        {
            oReviewInfo.DateTime = oCore.modified.getTime();
        }
        else
        {
            oReviewInfo.DateTime = (new Date()).getTime();
        }
		return oReviewInfo;
    };
    CDocumentComparison.prototype.getElementsForSetReviewType = function (oObject) {
        if (!oObject) {
            return [];
        }
        const arrReturnObjects = [];
        const arrCheckObjects = [oObject];

        while (arrCheckObjects.length) {
            const oCheckObject = arrCheckObjects.pop();
						if (oCheckObject.Type === para_Drawing && !this.options.textBoxes) {
							continue;
						}
            if(oCheckObject.GetReviewInfo && oCheckObject.SetReviewTypeWithInfo)
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
            else if(AscCommon.isRealObject(oCheckObject.Content))
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
                let oReviewInfo = this.getReviewInfo(sCustomReviewUserName, nCustomReviewDate);
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
    };

    CDocumentComparison.prototype.setReviewInfoRecursive = function(oObject, nType)
    {
        const arrNeedReviewObjects = this.getElementsForSetReviewType(oObject);
        this.setReviewInfoForArray(arrNeedReviewObjects, nType);
    };
    CDocumentComparison.prototype.saveReviewInfo = function(oNewObject, oOldObject)
    {
        if (oOldObject.GetReviewInfo && oNewObject.SetReviewTypeWithInfo) {
            const reviewType = oOldObject.GetReviewType && oOldObject.GetReviewType();
            const reviewInfo = oOldObject.GetReviewInfo();
            oNewObject.SetReviewTypeWithInfo && oNewObject.SetReviewTypeWithInfo(reviewType, reviewInfo ? reviewInfo.Copy() : undefined, false);
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
        if(oObject.GetReviewInfo && oObject.SetReviewTypeWithInfo)
        {
            const oReviewInfo = this.getReviewInfo();
            oObject.SetReviewTypeWithInfo(nType, oReviewInfo, false);

        }
    };
    CDocumentComparison.prototype.getNodeConstructor = function () {
        return CNode;
    };

    CDocumentComparison.prototype.createNodeFromDocContent = function(oElement, oParentNode, oWordCounter, isOriginalDocument)
    {
        const oRet = this.createNode(oElement, oParentNode);
        const bRoot = (oParentNode === null);
        for(let i = 0; i < oElement.Content.length; ++i)
        {
            const oChElement = oElement.Content[i];
            if(oChElement instanceof Paragraph)
            {
                if(bRoot)
                {
                    oWordCounter = new WordCounter();
                }
                const oParagraphNode = this.createNodeFromRunContentElement(oChElement, oRet, oWordCounter, isOriginalDocument);
                if(bRoot)
                {
                    oParagraphNode.wordCounter = oWordCounter;
                }
            }
            else if(oChElement instanceof CBlockLevelSdt)
            {
                if(bRoot)
                {
                    oWordCounter = new WordCounter();
                }
                const oBlockNode = this.createNodeFromDocContent(oChElement.Content, oRet, oWordCounter, isOriginalDocument);
                if(bRoot)
                {
                    oBlockNode.wordCounter = oWordCounter;
                }
            }
            else if(oChElement instanceof CTable)
            {
                if(this.options.getTables())
                {
                    if(bRoot)
                    {
                        oWordCounter = new WordCounter();
                    }
                    const oTableNode = this.createNode(oChElement, oRet);
                    if(bRoot)
                    {
                        oWordCounter = new WordCounter();
                        oTableNode.wordCounter = oWordCounter;
                    }
                    for(let j = 0; j < oChElement.Content.length; ++j)
                    {
                        const oRowNode = this.createNode(oChElement.Content[j], oTableNode);
                        for(let k = 0; k < oChElement.Content[j].Content.length; ++k)
                        {
                            this.createNodeFromDocContent(oChElement.Content[j].Content[k].Content, oRowNode, oWordCounter, isOriginalDocument);
                        }
                    }
                }
            }
            else if (oChElement instanceof AscCommonWord.CMockParagraph) {
                if(bRoot)
                {
                    oWordCounter = new WordCounter();
                }
                const oParagraphNode = this.createNodeFromRunContentElement(oChElement, oRet, oWordCounter, isOriginalDocument);
                if(bRoot)
                {
                    oParagraphNode.wordCounter = oWordCounter;
                }
            }
            else
            {
                const oNode = this.createNode(oChElement, oRet);
                if(bRoot)
                {
                    oWordCounter = new WordCounter();
                    oNode.wordCounter = oWordCounter;
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

	CDocumentComparison.prototype._createNodeFromRun = function(oRun, oLastText, oWordCounter, oRet, TextElementConstructor, NodeConstructor, oReviewInfo) {
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
                if(isBreakWordElement(oRunElement) || this.isWordsByOneSymbol)
                {
                    if(oLastText.elements.length > 0)
                    {
                        this.createNode(oLastText, oRet);
                        oLastText.updateHash(oWordCounter, this);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                    }
										if (this.options.footNotes || !(oRunElement.Type === para_FootnoteReference || oRunElement.Type === para_EndnoteReference)) {
											oLastText.setLastRun(oRun);
											oLastText.addToElements(oRunElement, oReviewInfo);
											this.createNode(oLastText, oRet);
											oLastText.updateHash(oWordCounter, this);
										}


										if (oLastText.elements.length > 0) {
											oLastText = new TextElementConstructor();
											oLastText.setFirstRun(oRun);
										}
                    oLastText.setLastRun(oRun);
                }
                else if(oRunElement.Type === para_Drawing)
                {
                    if(oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oWordCounter, this);
                        this.createNode(oLastText, oRet);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                        oLastText.setLastRun(oRun);
                    }
                    oLastText.addToElements(oRun.Content[j], oReviewInfo);
                    this.createNode(oLastText, oRet);
                    oLastText = new TextElementConstructor();
                    oLastText.setFirstRun(oRun);
                    oLastText.setLastRun(oRun);
                }
                else if (oRunElement.Type === para_RevisionMove) {
                    if(oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oWordCounter, this);
                        this.createNode(oLastText, oRet);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                        oLastText.setLastRun(oRun);
                    }
                }
                else if(oRunElement.Type === para_End)
                {
                    if(oLastText.elements.length > 0)
                    {
                        oLastText.updateHash(oWordCounter, this);
                        this.createNode(oLastText, oRet);
                        oLastText = new TextElementConstructor();
                        oLastText.setFirstRun(oRun);
                        oLastText.setLastRun(oRun);
                    }
                    oLastText.setFirstRun(oRun);
                    oLastText.setLastRun(oRun);
                    // мы будем сравнивать ревью paraEnd отдельно, поскольку это единственный общий элемент в параграфе, до которого мы можем вставить любой различающийся контент
                    oLastText.addToElements(oRun.Content[j], {reviewType: reviewtype_Common, moveReviewType: Asc.c_oAscRevisionsMove.NoMove});
                    this.createNode(oLastText, oRet);
                    oLastText.updateHash(oWordCounter, this);
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
    };

    CDocumentComparison.prototype.getCompareReviewInfo = function (oRun) {

    };

    CDocumentComparison.prototype.createNodeFromRun = function (oRun, oLastText, oWordCounter, oRet) {
        const TextElementConstructor = this.getTextElementConstructor();
        const NodeConstructor = this.getNodeConstructor();
        const oReviewInfo = this.getCompareReviewInfo(oRun);
        return this._createNodeFromRun(oRun, oLastText, oWordCounter, oRet, TextElementConstructor, NodeConstructor, oReviewInfo);
    };
    CDocumentComparison.prototype.createNodeFromRunContentElement = function(oElement, oParentNode, oWordCounter, isOriginalDocument)
    {
	    this.oBookmarkManager.previousNode = null;
	    this.oCommentManager.previousNode = null;
        const oRet = this.createNode(oElement, oParentNode);
        let oLastText = null;
        for(let i = 0; i < oElement.Content.length; ++i)
        {
            const oRun = oElement.Content[i];
            if(oRun instanceof ParaRun)
            {
                oLastText = this.createNodeFromRun(oRun, oLastText, oWordCounter, oRet);
            }
            else if (oRun instanceof AscCommon.CParaRevisionMove)
            {
                if(oLastText && oLastText.elements.length > 0)
                {
                    oLastText.updateHash(oWordCounter, this);
                    this.createNode(oLastText, oRet);
                }
                oLastText = null;
                this.createNode(oRun, oRet);
                if (!isOriginalDocument && !this.bSkipChangeMoveType) {

                    this.oComparisonMoveMarkManager.changeRevisedMoveMarkId(oRun, this);
                    this.oComparisonMoveMarkManager.updateInsertMoveMarkId(oRun);
                }
                this.oComparisonMoveMarkManager.updateMoveMarksStack(oRun);

                oParentNode.bHaveMoveMarks = true;
            }
						else if (oRun instanceof AscCommonWord.CParagraphBookmark)
            {
							if (oLastText)
							{
								this.oBookmarkManager.addToStack(oRun, oLastText.elements.length);
							}
							else
							{
								this.oBookmarkManager.addToStack(oRun, 0);
							}
            }
						else if (oRun instanceof AscCommon.ParaComment)
            {
							if (this.options.comments) {
								const oComment = this.getComment(oRun.GetCommentId());
								const oElement = new CCommentElement(oComment, oRun);
								if (oLastText)
								{
									this.oCommentManager.addToStack(oElement, oLastText.elements.length);
								}
								else
								{
									this.oCommentManager.addToStack(oElement, 0);
								}
							} else {
								this.oCommentManager.addToCheckSaveParaComment(oRun);
							}
            }
            else
            {
	            if (oLastText && oLastText.elements.length > 0)
	            {
		            oLastText.updateHash(oWordCounter, this);
		            this.createNode(oLastText, oRet);
	            }
	            oLastText = null;
	            if (Array.isArray(oRun.Content))
	            {
		            this.createNodeFromRunContentElement(oRun, oRet, oWordCounter, isOriginalDocument);
	            }
	            else
	            {
		            this.createNode(oRun, oRet);
	            }
            }
        }
        if(oLastText && oLastText.elements.length > 0)
        {
            oLastText.updateHash(oWordCounter, this);
            this.createNode(oLastText, oRet);
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

	CDocumentComparison.prototype.updateCommentsQuoteText = function ()
	{
		const oLogicDocument = this.originalDocument;
		const oSelectionState = oLogicDocument.SaveDocumentState(false);
		const oComments = oLogicDocument.GetCommentsManager();
		for (let sId in oComments.m_arrCommentsById)
		{
			const oComment = oComments.m_arrCommentsById[sId];
			if (oComment.SelectCommentText())
			{
				const oCommentData = oComment.GetData();
				const sSelectedText = oLogicDocument.GetSelectedText(false);
				if (sSelectedText !== oCommentData.GetQuoteText())
				{
					const oCopyData = oCommentData.Copy();
					oCopyData.Set_QuoteText(sSelectedText === "" ? null : sSelectedText);
					oComment.SetData(oCopyData);
					this.api.sync_ChangeCommentData(oComment.Id, oCopyData);
				}
			}
		}
		oLogicDocument.LoadDocumentState(oSelectionState);
	};

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
				const oldTrackRevisions = oDoc1.GetLocalTrackRevisions();
	    oDoc1.SetTrackRevisions(false);
	    let bHaveRevisons2 = false;
	    const oDoc2 = AscFormat.ExecuteNoHistory(function(){
            const openParams = {disableRevisions: true, noSendComments: true, noGenerateSmartArts: true};
            const oTempDocument = new CDocument(oApi.WordControl.m_oDrawingDocument, false);
            const oBinaryFileReader = new AscCommonWord.BinaryFileReader(oTempDocument, openParams);
            AscCommon.pptx_content_loader.Start_UseFullUrl(oApi.insertDocumentUrlsData);
            if (!oBinaryFileReader.Read(sBinary2))
            {
                return null;
            }
            if(oTempDocument)
            {
                bHaveRevisons2 = oBinaryFileReader.oReadResult && oBinaryFileReader.oReadResult.hasRevisions;
            }
            return oTempDocument;
        }, this, []);
	    oDoc1.SetTrackRevisions(oldTrackRevisions);
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

	function CComparisonBookmarkManager()
	{
		this.mapNamesForMerge = {};
		this.mapIdToBookmarkId = {};
		this.needUpdateBookmarks = false;
		this.bookmarkStack = [];
		this.previousNode = null;
	}

	CComparisonBookmarkManager.prototype.init = function (oMainDocument, oRevisedDocument)
	{
		oRevisedDocument.UpdateBookmarks();
		oMainDocument.UpdateBookmarks();
		const arrRevisedBookmarks = oRevisedDocument.BookmarksManager.Bookmarks;
		const arrMainBookmarks = oMainDocument.BookmarksManager.Bookmarks;
		for (let i = 0; i < arrRevisedBookmarks.length; i += 1)
		{
			const oBookmark = arrRevisedBookmarks[i][0];
			this.mapNamesForMerge[oBookmark.GetBookmarkName()] = arrRevisedBookmarks[i];
		}
		for (let i = 0; i < arrMainBookmarks.length; i += 1)
		{
			const oBookmark = arrMainBookmarks[i][0];
			delete this.mapNamesForMerge[oBookmark.GetBookmarkName()];
		}
		for (let sName in this.mapNamesForMerge)
		{
			this.needUpdateBookmarks = true;
			const oStartBookmark = this.mapNamesForMerge[sName][0];
			const oEndBookmark = this.mapNamesForMerge[sName][1];
			const nNewId = oMainDocument.BookmarksManager.GetNewBookmarkId();
			this.addToLink(oStartBookmark.GetId(), nNewId);
			this.addToLink(oEndBookmark.GetId(), nNewId);
		}
	};
	CComparisonBookmarkManager.prototype.addToLink = function (sId, nBookmarkId) {
		this.mapIdToBookmarkId[sId] = nBookmarkId;
	};
	CComparisonBookmarkManager.prototype.isSkip = function (oElement)
	{
		if (oElement instanceof AscCommonWord.CParagraphBookmark)
		{
			return !this.mapIdToBookmarkId[oElement.GetId()];
		}
		return false;
	};
	CComparisonBookmarkManager.prototype.getId = function (oBookmark)
	{
		const sId = this.mapIdToBookmarkId[oBookmark.GetId()];
		if (sId) {
			return parseInt(sId, 10);
		}
	};
	CComparisonBookmarkManager.prototype.addToStack = function (oBookmark, nInsertIndex)
	{
		if (!this.isSkip(oBookmark))
		{
			const oLabel = new CBookmarkLabel(oBookmark, nInsertIndex);
			this.bookmarkStack.push(oLabel);
		}
	};

	CComparisonBookmarkManager.prototype.updateBookmarksStack = function (oNode)
	{
		const arrNextStack = [];
		const mapCheckAdd = {};
		const oPreviousElement = this.previousNode && this.previousNode.element;
		for (let i = 0; i < this.bookmarkStack.length; i += 1)
		{
			const oLabel = this.bookmarkStack[i];
			const oBookmark = oLabel.bookmark;
			if (oLabel.insertIndex === oNode.element.elements.length)
			{
				oLabel.insertIndex = 0;
				arrNextStack.push(oLabel);
			}
			else if (!mapCheckAdd[oBookmark.GetBookmarkId()] && oLabel.insertIndex === 0 && (oPreviousElement instanceof CTextElement) && !oBookmark.IsStart())
			{
					oPreviousElement.addBookmark(oBookmark, oPreviousElement.elements.length);
			}
			else
			{
				oNode.addBookmark(oBookmark, oLabel.insertIndex);
				mapCheckAdd[oBookmark.GetBookmarkId()] = true;
			}
		}
		this.bookmarkStack = arrNextStack;
	};

	function CBookmarkLabel(oBookmark, nInsertIndex)
	{
		this.bookmark = oBookmark;
		this.insertIndex = nInsertIndex;
	}

	function CComparisonCommentManager(comparison)
	{
		this.comparison = comparison;
		this.commentsStack = [];
		this.previousNode = null;
		this.mapDelete = {};
		this.mapChecked = {};
		this.mapLink = {};
		this.mapMergeLater = {};
		this.checkSkipParaComment = {};
		this.savedParaComments = {};
	}

	CComparisonCommentManager.prototype.addToCheckSaveParaComment = function (oParaComment) {
		const sCommentId = oParaComment.GetCommentId();
		if (!this.checkSkipParaComment[sCommentId]) {
			this.checkSkipParaComment[sCommentId] = Array(2);
		}
		if (oParaComment.IsCommentStart()) {
			this.checkSkipParaComment[sCommentId][0] = oParaComment;
		} else {
			this.checkSkipParaComment[sCommentId][1] = oParaComment;
		}
	};
	CComparisonCommentManager.prototype.generateSavedParaComments = function () {
		for (let sCommentId in this.checkSkipParaComment) {
			const arrParaComments = this.checkSkipParaComment[sCommentId];
			const oStartParaComment = arrParaComments[0];
			const oEndParaComment = arrParaComments[1];
			if (oStartParaComment && oEndParaComment) {
				if (oStartParaComment.Paragraph === oEndParaComment.Paragraph && oStartParaComment.Parent === oEndParaComment.Parent) {
					this.savedParaComments[oStartParaComment.Id] = true;
					this.savedParaComments[oEndParaComment.Id] = true;
				}
			}
		}
		this.checkSkipParaComment = {};
	}
	CComparisonCommentManager.prototype.addToLink = function (sKey, sValue)
	{
		this.mapLink[sKey] = sValue;
	};
	CComparisonCommentManager.prototype.getAddedComment = function (oData)
	{
		const oComments = this.comparison.originalDocument.Comments;
		const oComment = new AscCommon.CComment(oComments, oData);
		oComments.Add(oComment);
		this.comparison.api.sync_AddComment(oComment.Id, oData);
		return oComment;
	};
	CComparisonCommentManager.prototype.getMergeLater = function (sCommentId)
	{
		return this.mapMergeLater[sCommentId];
	};
	CComparisonCommentManager.prototype.addToMergeLater = function (oParaComment, sCommentId)
	{
		if (!this.mapMergeLater[sCommentId])
		{
			this.mapMergeLater[sCommentId] = [];
		}
		this.mapMergeLater[sCommentId].push(oParaComment);
	};
	CComparisonCommentManager.prototype.pushToArrInsertContentFromMerge = function (oParaComment, arrContentToInsert)
	{
		if (!oParaComment.IsCommentStart())
			return;
		const sCommentId = oParaComment.GetCommentId();
		const arrMerge = this.getMergeLater(sCommentId);
		if (arrMerge)
		{
			arrContentToInsert.push.apply(arrContentToInsert, arrMerge);
		}
	};
	CComparisonCommentManager.prototype.applyCommentFromData = function (oData, arrCommentElements, sRevisedQuoteText, sRevisedCommentId)
	{
		const oCopyData = oData.Copy();
		oCopyData.Set_QuoteText(sRevisedQuoteText);
		const oComment = this.getAddedComment(oCopyData);
		const oStartParaComment = new AscCommon.ParaComment(true, oComment.GetId());
		const oEndParaComment = new AscCommon.ParaComment(false, oComment.GetId());
		oComment.SetRangeStart(oStartParaComment.GetId());
		oComment.SetRangeEnd(oEndParaComment.GetId());
		const oEndCommentElement = new CCommentElement(oComment, oEndParaComment);
		arrCommentElements.push(oEndCommentElement);
		this.addToMergeLater(oStartParaComment, sRevisedCommentId);
	};
	CComparisonCommentManager.prototype.getSortedFunction = function ()
	{
		return function (a, b)
		{
			const nQuoteCommentDifference = a.quoteDifference - b.quoteDifference;
			if (nQuoteCommentDifference !== 0)
			{
				return nQuoteCommentDifference;
			}
			return a.commentDifference - b.commentDifference;
		}
	};

	CComparisonCommentManager.prototype.checkComments = function (arrMainComments, arrRevisedComments, arrCommentElements)
	{
		if (!arrRevisedComments)
			return;
		if (!arrMainComments)
		{
			arrMainComments = [];
		}

		const arrComparing = [];
		const fSortedFunction = this.getSortedFunction();
		for (let i = 0; i < arrRevisedComments.length; i++)
		{
			const oRevisedElement = arrRevisedComments[i];
			const sRevisedCommentId = oRevisedElement.element.GetCommentId();
			const bIsStart = oRevisedElement.element.IsCommentStart();
			if (!this.isChecked(sRevisedCommentId) && !bIsStart)
			{
				this.check(sRevisedCommentId);
				for (let j = 0; j < arrMainComments.length; j++)
				{
					const oMainElement = arrMainComments[j];
					const sMainCommentId = oMainElement.element.GetCommentId();
					if (!this.isForDelete(sMainCommentId))
					{
						const oDifference = oMainElement.getDifference(oRevisedElement);

						if (oDifference)
						{
							arrComparing.push(oDifference);
						}
					}
				}
			}
		}
		arrComparing.sort(fSortedFunction);
		for (let i = 0; i < arrComparing.length; i++)
		{
			const oMainElement = arrComparing[i].mainElement;
			const oRevisedElement = arrComparing[i].anotherElement;
			const sRevisedCommentId = oRevisedElement.getCommentId();
			const sMainCommentId = oMainElement.getCommentId();
			if (!this.mapLink[sRevisedCommentId] && !this.isForDelete(sMainCommentId))
			{
				if (arrComparing[i].arrData)
				{
					for (let j = 0; j < arrComparing[i].arrData.length; j += 1)
					{
						this.applyCommentFromData(arrComparing[i].arrData[j], arrCommentElements, oRevisedElement.getQuoteText(), sRevisedCommentId);
					}
				}
				else
				{
					const arrRevisedAnswers = oRevisedElement.getAnswers();
					const arrMainAnswers = oMainElement.getAnswers();
					const oComment = this.getComment(sRevisedCommentId);
					if (oComment)
					{
						const oCopyData = oComment.Data.Copy();
						for (let j = arrRevisedAnswers.length; j < arrMainAnswers.length; j += 1)
						{
							const oCopyAnswer = arrMainAnswers[j].Copy();
							oCopyData.Add_Reply(oCopyAnswer);
						}
						oComment.SetData(oCopyData);
						this.comparison.api.sync_ChangeCommentData(oComment.Id, oCopyData);
					}
				}
				this.addToLink(sRevisedCommentId, sMainCommentId);
				this.addToDelete(sMainCommentId);
			}
		}
	};

	CComparisonCommentManager.prototype.check = function (sCommentId)
	{
		this.mapChecked[sCommentId] = true;
	};
	CComparisonCommentManager.prototype.isChecked = function (sCommentId)
	{
		return this.mapChecked[sCommentId];
	};
	CComparisonCommentManager.prototype.isForDelete = function (sCommentId)
	{
		return this.mapDelete[sCommentId];
	};
	CComparisonCommentManager.prototype.addToDelete = function (sCommentId)
	{
		this.mapDelete[sCommentId] = true;
	};

	CComparisonCommentManager.prototype.getComment = function (sCommentId)
	{
		if (this.comparison.CommentsMap[sCommentId])
			return this.comparison.CommentsMap[sCommentId];

		return this.comparison.getComment(sCommentId);
	};
	CComparisonCommentManager.prototype.addToStack = function (oComment, nInsertIndex)
	{
			const oLabel = new CCommentLabel(oComment, nInsertIndex);
			this.commentsStack.push(oLabel);
	};

	CComparisonCommentManager.prototype.updateCommentsStack = function (oNode)
	{
		const arrNextStack = [];
		const mapCheckAdd = {};
		const oPreviousElement = this.previousNode && this.previousNode.element;
		for (let i = 0; i < this.commentsStack.length; i += 1)
		{
			const oLabel = this.commentsStack[i];
			const oParaComment = oLabel.comment.element;
			if (oLabel.insertIndex === oNode.element.elements.length)
			{
				oLabel.insertIndex = 0;
				arrNextStack.push(oLabel);
			}
			else if (!mapCheckAdd[oParaComment.GetCommentId()] && oLabel.insertIndex === 0 && (oPreviousElement instanceof CTextElement) && !oParaComment.IsCommentStart())
			{
				oPreviousElement.addComment(oLabel.comment, oPreviousElement.elements.length);
			}
			else
			{
				oNode.addComment(oLabel.comment, oLabel.insertIndex);
				mapCheckAdd[oParaComment.GetCommentId()] = true;
			}
		}
		this.commentsStack = arrNextStack;
	};

	function CCommentLabel(oParaComment, nInsertIndex)
	{
		this.comment = oParaComment;
		this.insertIndex = nInsertIndex;
	}
	
	function CRunCollector(arrElements, parentParagraph) {
		this.elements = arrElements;
		this.runIndex = null;
		this.runElementIndex = null;
		this.elementIndex = null;
		this.innerElementIndex = null;
		this.parent = null;
		this.collectTextPrChange = null;
		this.collectReviewChange = null;

		this.init(parentParagraph);
	}

	CRunCollector.prototype.init = function (parentParagraph) {
		this.elementIndex = this.elements.length - 1;
		const oLastElement = this.elements[this.elementIndex];
		this.innerElementIndex = oLastElement.elements.length - 1;
		const oLastRun = oLastElement.lastRun;
		const oParent = parentParagraph || oLastRun.GetParent();
		this.parent = oParent;
		const oContent = oParent.Content;
		for (let i = oContent.length - 1; i >= 0; i -= 1) {
			const oRun = oContent[i];
			if (oRun === oLastRun) {
				this.runIndex = i;
				break;
			}
		}
		 const oLastRunElement = oLastElement.elements[this.innerElementIndex];
		for (let i = oLastRun.Content.length - 1; i >= 0; i -= 1) {
			const oElement = oLastRun.Content[i];
			if (oElement === oLastRunElement) {
				this.runElementIndex = i;
				break;
			}
		}
	};
	
	CRunCollector.prototype.getCurrentElement = function () {
		return this.elements[this.elementIndex];
	};

	CRunCollector.prototype.skipTo = function (elementIndex, innerElementIndex) {
		while (this.elementIndex > elementIndex) {
			let countElements = this.innerElementIndex + 1;
			while (countElements) {
				if (this.runElementIndex + 1 > countElements) {
					this.runElementIndex -= countElements;
					countElements = 0;
				} else {
					countElements -= this.runElementIndex + 1;
					this.nextRun();
				}
			}
			this.elementIndex -= 1;
			const oElement = this.elements[this.elementIndex];
			if (oElement) {
				this.innerElementIndex = oElement.elements.length - 1;
			}
		}
		if (this.innerElementIndex > innerElementIndex) {
			let countElements = this.innerElementIndex - innerElementIndex;
			while (countElements) {
				if (this.runElementIndex + 1 > countElements) {
					this.runElementIndex -= countElements;
					countElements = 0;
				} else {
					countElements -= this.runElementIndex + 1;
					this.nextRun();
				}
			}
			this.innerElementIndex = innerElementIndex;
		}
	};
	CRunCollector.prototype.dropLastCollect = function ()
	{
		if (this.collectTextPrChange) {
			this.collectTextPrChange.dropLastRun();
		}
		if (this.collectReviewChange) {
			this.collectReviewChange.dropLastRun();
		}
	};
	CRunCollector.prototype.setCollectTextPrRuns = function (oChange) {
		this.collectTextPrChange = oChange;
	};
	CRunCollector.prototype.setCollectReviewRuns = function (oChange) {
		this.collectReviewChange = oChange;
	}
	CRunCollector.prototype.addToCollectTextPr = function (oRun) {
		if (this.collectTextPrChange) {
			this.collectTextPrChange.addRun(oRun);
		}
	};
	CRunCollector.prototype.addToCollectReviewChange = function (oRun) {
		if (this.collectReviewChange) {
			this.collectReviewChange.addRun(oRun);
		}
	};
	CRunCollector.prototype.addToCollectBackTextPr = function (oRun) 
	{
		if (this.collectTextPrChange) {
			this.collectTextPrChange.addRunBack(oRun);
		}
	};
	CRunCollector.prototype.addToCollectBackReviewChange = function (oRun) 
	{
		if (this.collectReviewChange) {
			this.collectReviewChange.addRunBack(oRun);
		}
	};
	CRunCollector.prototype.addToCollect = function (oRun)
	{
		this.addToCollectReviewChange(oRun);
		this.addToCollectTextPr(oRun);
	};
	CRunCollector.prototype.addToCollectBack = function (oRun)
	{
		this.addToCollectBackTextPr(oRun);
		this.addToCollectBackReviewChange(oRun);
	};
	CRunCollector.prototype.addToCollectCurrentRun = function ()
	{
		this.addToCollect(this.getRun());
	};
	CRunCollector.prototype.nextRun = function () {
		do {
			this.runIndex -= 1;
		} while (!(this.getRun() instanceof AscCommonWord.ParaRun) && this.runIndex >= 0 || this.getRun().Content.length === 0)
		const oCurRun = this.getRun();
		if (oCurRun) {
			this.addToCollectCurrentRun();
			this.runElementIndex = oCurRun.Content.length - 1;
		}
	};
	CRunCollector.prototype.getRun = function () {
		const oContent = this.parent.Content;
		return oContent[this.runIndex];
	};
	CRunCollector.prototype.splitCurrentRun = function (bAfter)
	{
		const oRun = this.getRun();
		return oRun.Split2(bAfter ? this.runElementIndex + 1 : this.runElementIndex, this.parent, this.runIndex);
	};

    window['AscCommonWord']["CompareBinary"] =  window['AscCommonWord'].CompareBinary = CompareBinary;
    window['AscCommonWord']["ComparisonOptions"] = window['AscCommonWord'].ComparisonOptions = ComparisonOptions;
    window['AscCommonWord']['CompareDocuments'] = CompareDocuments;
    window['AscCommonWord'].CDocumentComparison = CDocumentComparison;
    window['AscCommonWord'].CNode = CNode;
    window['AscCommonWord'].CTextElement = CTextElement;
    window['AscCommonWord'].isBreakWordElement = isBreakWordElement;
})();
