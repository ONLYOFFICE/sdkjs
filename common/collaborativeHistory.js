/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function (window)
{
	/**
	 * Class for working with a common history in collaborative editing
	 * @param {AscCommon.CCollaborativeEditingBase} coEditing
	 * @constructor
	 */
	function CCollaborativeHistory(coEditing)
	{
		this.CoEditing = coEditing;

		this.Changes   = []; // List of all changes
		this.ChangesSplitByPoints = [] // List of changes split by points
		this.OwnRanges = []; // Ranges of own changes

		this.SyncIndex      = -1; // Position in the array of changes that are synced with server
		this.curChangeIndex = -1; // Current position in the array of changes split by points
		//this.StepTextPoint = undefined; // Position of previous state
		
		this.textRecovery = null;
	}
	/**
	 * Split changes by points
	 */
	CCollaborativeHistory.prototype.SplitChangesByPoints = function ()
	{
		if (!this.Changes || !this.Changes || this.ChangesSplitByPoints.length !== 0)
			return;

		let arrCurrent = [0];
		for (let i = 1; i < this.Changes.length; i++)
		{
			let oCurrentChange = this.Changes[i];
			let oPrevChange = (i === 0)
				? undefined
				: this.Changes[i - 1];

			if (!(oPrevChange && (oPrevChange instanceof AscCommon.CChangesTableIdDescription || oPrevChange.IsDescriptionChange() === oCurrentChange.IsDescriptionChange())))
			{
				arrCurrent.push(i);
			}
		}

		arrCurrent.push(this.Changes.length);
		this.ChangesSplitByPoints = arrCurrent;
		this.curChangeIndex = arrCurrent.length - 1;
	};
	CCollaborativeHistory.prototype.clear = function()
	{
		this.Changes   = [];
		this.OwnRanges = [];
		this.ChangesSplitByPoints = [];
		
		this.SyncIndex      = -1;
		this.curChangeIndex = -1;
		
		this.textRecovery = null;
	};
	/**
	 * Navigate to a specific point in revision history
	 * @param {number} pointIndex - Position to navigate to
	 * @constructor
	 */
	CCollaborativeHistory.prototype.NavigationRevisionHistoryByStep = function(pointIndex)
	{
		let logicDocument = this.CoEditing.GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return false;

		this.UndoDeletedTextRecovery();
		this.SplitChangesByPoints();
		if (this.curChangeIndex < 0)
			return false;
		
		pointIndex = Math.max(0, Math.min(pointIndex, this.ChangesSplitByPoints.length - 1));
		if (pointIndex === this.curChangeIndex)
			return false;
		
		let changes;
		if (this.curChangeIndex < pointIndex)
			changes = this._RedoChanges(this.ChangesSplitByPoints[this.curChangeIndex], this.ChangesSplitByPoints[pointIndex]);
		else
			changes = this._UndoChanges(this.ChangesSplitByPoints[this.curChangeIndex], this.ChangesSplitByPoints[pointIndex]);
		
		this.curChangeIndex	= pointIndex;
		logicDocument.RecalculateByChanges(changes);
		return true;
	};
	CCollaborativeHistory.prototype.AddChange = function(change)
	{
		this.Changes.push(change);
	};
	CCollaborativeHistory.prototype.AddOwnChanges = function(ownChanges, deleteIndex)
	{
		// TODO: When deleting changes, OwnRanges that could reference these changes are not deleted
		//       Need to check if this is correct

		if (null !== deleteIndex)
			this.Changes.length = this.SyncIndex + deleteIndex;
		else
			this.SyncIndex = this.Changes.length;

		// TODO: For now we do this as a single point to undo. Need to iterate through the array and split it
		//       by individual actions. In fact, this scheme works in fast collaborative editing,
		//       so two points shouldn't fall into one save.
		if (ownChanges.length > 0)
		{
			this.OwnRanges.push(new COwnRange(this.Changes.length, ownChanges.length));
			this.Changes = this.Changes.concat(ownChanges);
		}
	};
	CCollaborativeHistory.prototype.GetAllChanges = function()
	{
		return this.Changes;
	};
	CCollaborativeHistory.prototype.GetChangeCount = function()
	{
		return this.Changes.length;
	};
	CCollaborativeHistory.prototype.CanUndoGlobal = function()
	{
		return (this.Changes.length > 0);
	};
	CCollaborativeHistory.prototype.CanUndoOwn = function()
	{
		return (this.OwnRanges.length > 0)
	};
	CCollaborativeHistory.prototype._RedoChanges = function(startPos, endPos)
	{
		let changes = [];
		for (let i = startPos; i < endPos; ++i)
		{
			let change = this.Changes[i];
			if (change.IsContentChange())
			{
				let simpleChanges = change.ConvertToSimpleChanges();
				for (let simpleIndex = 0; simpleIndex < simpleChanges.length; ++simpleIndex)
				{
					simpleChanges[simpleIndex].Redo();
					changes.push(simpleChanges[simpleIndex]);
				}
			}
			else
			{
				change.Redo();
				changes.push(change);
			}
		}
		return changes;
	};
	CCollaborativeHistory.prototype._UndoChanges = function(startPos, endPos)
	{
		let changes = [];
		for (let i = startPos - 1; i >= endPos; --i)
		{
			let change = this.Changes[i];
			if (change.IsContentChange())
			{
				let simpleChanges = change.ConvertToSimpleChanges();
				for (let simpleIndex = simpleChanges.length - 1; simpleIndex >= 0; --simpleIndex)
				{
					simpleChanges[simpleIndex].Undo();
					changes.push(simpleChanges[simpleIndex]);
				}
			}
			else
			{
				change.Undo();
				changes.push(change);
			}
		}
		return changes;
	};
	/**
	 * Undo a specified number of actions
	 * @param {number} count
	 * @returns {[]} returns array of reverted changes
	 */
	CCollaborativeHistory.prototype.UndoGlobalChanges = function(count)
	{
		this.UndoDeletedTextRecovery();

		count = Math.min(count, this.Changes.length);

		if (!count)
			return [];

		let index = this.Changes.length - 1;
		let changeArray = [];
		while (index >= this.Changes.length - count)
		{
			let change = this.Changes[index--];
			if (!change)
				continue;

			if (change.IsContentChange())
			{
				let simpleChanges = change.ConvertToSimpleChanges();
				for (let simpleIndex = simpleChanges.length - 1; simpleIndex >= 0; --simpleIndex)
				{
					simpleChanges[simpleIndex].Undo();
					changeArray.push(simpleChanges[simpleIndex]);
				}
			}
			else
			{
				change.Undo();
				changeArray.push(change);
			}
		}

		this.Changes.length = this.Changes.length - count;
		return changeArray;
	};
	/**
	 * Undo all actions that fell into the last history point
	 * @returns {[]} returns array of reverted changes
	 */
	CCollaborativeHistory.prototype.UndoGlobalPoint = function()
	{
		this.UndoDeletedTextRecovery();

		let count = 0;
		for (let index = this.Changes.length - 1; index > 0; --index, ++count)
		{
			let change = this.Changes[index];
			if (!change)
				continue;

			if (change.IsDescriptionChange())
			{
				count++;
				break;
			}
		}

		return count ? this.UndoGlobalChanges(count) : [];
	};
	/**
	 * Get the number of history positions in the current revision
	 * @return {number}
	 */
	CCollaborativeHistory.prototype.GetGlobalPointCount = function()
	{
		this.SplitChangesByPoints();
		return this.ChangesSplitByPoints.length;
	};
	/**
	 * Get the current position in revision history
	 * @return {number}
	 */
	CCollaborativeHistory.prototype.GetGlobalPointIndex = function()
	{
		this.SplitChangesByPoints();
		return this.curChangeIndex;
	};
	/**
	 * Navigate to the required point in revision history
	 * @param nPos - position in history
	 * @return {boolean} - whether navigation to this position was performed
	 */
	CCollaborativeHistory.prototype.MoveToPoint = function(nPos)
	{
		return this.NavigationRevisionHistoryByStep(nPos);
	};
	CCollaborativeHistory.prototype.InitTextRecover = function ()
	{
		if (this.textRecovery)
			return;
		
		let logicDocument = this.CoEditing.GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return;
		
		this.textRecovery = new AscCommon.DeletedTextRecovery(logicDocument);
	};
	/**
	 * Display deleted text for the given point in revision history
	 * @return {boolean} - whether the deleted text was displayed
	 */
	CCollaborativeHistory.prototype.RecoverDeletedText = function()
	{
		this.InitTextRecover();
		return this.textRecovery.RecoverDeletedText();
	};
	/**
	 * Cancel display of deleted text at this point in revision history
	 * @return {boolean}
	 */
	CCollaborativeHistory.prototype.UndoDeletedTextRecovery = function()
	{
		if (this.textRecovery)
			return this.textRecovery.UndoRecoveredText();

		return false;
	};
	CCollaborativeHistory.prototype.HaveDeletedTextRecovery = function()
	{
		return !!(this.textRecovery && this.textRecovery.HaveRecoveredText());
	};
	CCollaborativeHistory.prototype.GetCollaborativeMarks = function ()
	{
		return this.CoEditing.Get_CollaborativeMarks();
	}
	/**
	 * Undo own last actions by rolling them through others' changes
	 * @returns {[]} returns array of new changes
	 */
	CCollaborativeHistory.prototype.UndoOwnPoint = function()
	{
		// Form a new batch of actions that will roll back the actions we need
		let reverseChanges = this.GetReverseOwnChanges();
		if (reverseChanges.length <= 0)
		{
			// To prevent infinite saving (reconsider to work without saveChanges)
			this.saveChanges([]);
			return [];
		}

		for (let index = 0, count = reverseChanges.length; index < count; ++index)
		{
			reverseChanges[index].Load();
			reverseChanges[index].CheckNeedRecalculate();
		}

		// Create a point in history. Execute actions through regular functions (with disabled recalculation) that write to
		// history. Save the list of changes in a new point, delete this point.

		let historyPoint = this.CreateLocalHistoryPointByReverseChanges(reverseChanges);
		let changesToSend = [], changesToRecalc = [];
		for (let index = 0, count = historyPoint.Items.length; index < count; ++index)
		{
			let historyItem   = historyPoint.Items[index];
			let historyChange;
			let historyClass  = historyItem.Class;

			//todo: Refactor Class and Data in spreadsheet to follow other editors design
			if (AscCommon.History.Item_ToSerializable) {
				historyChange = AscCommon.History.Item_ToSerializable(historyItem);
			} else {
				historyChange = historyItem.Data;
			}

			if (!historyClass || !(historyClass.Get_Id || historyClass.Class && historyClass.Class.Get_Id))
				continue;

			if (historyItem.Binary.Len) //spreadsheet local changes
			{
				let data = AscCommon.CCollaborativeChanges.ToBase64(historyItem.Binary.Pos, historyItem.Binary.Len);
				changesToSend.push(data);
			}

			changesToRecalc.push(historyChange);
			this.AddChange(historyChange);
		}
		AscCommon.History.Remove_LastPoint();
		this.CoEditing.Clear_DCChanges();

		this.saveChanges(changesToSend);
		return changesToRecalc;
	};
	CCollaborativeHistory.prototype.saveChanges = function(changesToSend)
	{
		//separate function to override in excel
		(Asc.editor || editor).CoAuthoringApi.saveChanges(changesToSend, null, null, false, this.CoEditing.getCollaborativeEditing());
	};
	CCollaborativeHistory.prototype.GetEmptyContentChanges = function()
	{
		let changes = [];
		for (let index = this.Changes.length - 1; index >= 0; --index)
		{
			let tempChange = this.Changes[index];

			if (tempChange.IsContentChange() && tempChange.GetItemsCount() <= 0)
				changes.push(tempChange);
		}
		return changes;
	};
	/**
	 * Check that the last change in the document is the input of the specified character
	 * @param {AscWord.CRun} run
	 * @param {number} inRunPos
	 * @param {?number} codePoint
	 * @returns {boolean}
	 */
	CCollaborativeHistory.prototype.checkAsYouTypeEnterText = function(run, inRunPos, codePoint)
	{
		if (!this.Changes.length || !this.OwnRanges.length)
			return false;
		
		let lastOwnRange = this.OwnRanges[this.OwnRanges.length - 1];
		if (this.Changes.length !== lastOwnRange.Position + lastOwnRange.Length)
			return false;
		
		let lastChange = this.Changes[this.Changes.length - 1];
		return (AscDFH.historyitem_ParaRun_AddItem === lastChange.Type
			&& lastChange.Class === run
			&& lastChange.Pos === inRunPos - 1
			&& lastChange.Items.length
			&& (undefined === codePoint || lastChange.Items[0].GetCodePoint() === codePoint));
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CCollaborativeHistory.prototype.GetReverseOwnChanges = function()
	{
		// On the first step we commute the given batch of changes with the latest changes. We look at what set
		// of changes we get.
		// Our object model is simple: a class that possibly has a content as an array of elements (also classes),
		// and also this class may have a set of properties. Therefore we have exactly 2 types of changes:
		// content changes or property changes.
		// Changes of these two types commute with each other, changes of different classes also commute.

		if (this.OwnRanges.length <= 0)
			return [];

		let range     = this.OwnRanges[this.OwnRanges.length - 1];
		let nPosition = range.Position;
		let nCount    = range.Length;

		let arrReverseChanges = [];
		for (let nIndex = nCount - 1; nIndex >= 0; --nIndex)
		{
			let oChange = this.Changes[nPosition + nIndex];
			if (!oChange)
				continue;

			let oClass = oChange.GetClass();
			if (oChange.IsContentChange())
			{
				let _oChange = oChange.Copy();
				let simpleChanges = _oChange.ConvertToSimpleChanges();
				for (let simpleIndex = simpleChanges.length - 1; simpleIndex >= 0; --simpleIndex)
				{
					let simpleChange = simpleChanges[simpleIndex];
					if (this.CommuteContentChange(simpleChange, nPosition + nCount))
					{
						let oReverseChange = simpleChange.CreateReverseChange();
						if (oReverseChange)
						{
							arrReverseChanges.push(oReverseChange);
							oReverseChange.SetReverted(true);
						}
					}
				}
				oChange.SetReverted(true);
			}
			else if (oChange.IsSpreadsheetChange())
			{
				let _oChange = oChange.Copy();
				// Create the reverse change first
				let oReverseChange = _oChange.CreateReverseChange();
				if (oReverseChange) {
					if (this.CommuteRelated(oClass, oReverseChange, nPosition + nCount))
					{
						oReverseChange.SetReverted(true);
						arrReverseChanges.push(oReverseChange);
					}
					else
					{
						// TODO: for autoshapes we don't need to hide the entire point
						//       in tables the entire point is not accepted
						//       for example when inserting a column, the fill of the adjacent column is copied
						arrReverseChanges = [];
						for (let i = nCount - 1; i > nIndex; --i)
						{
							this.Changes[nPosition + i].SetReverted(false);
						}
						break;
					}
				}
				else if(null !== oReverseChange)
				{
					//do nothing if there are changes that are not ready
					arrReverseChanges = [];
					for (let i = nCount - 1; i > nIndex; --i)
					{
						this.Changes[nPosition + i].SetReverted(false);
					}
					break;
				}
				oChange.SetReverted(true);
			}
			else
			{
				let _oChange = oChange; // TODO: Should make a copy here

				if (this.CommutePropertyChange(oClass, _oChange, nPosition + nCount))
				{
					let oReverseChange = _oChange.CreateReverseChange();
					if (oReverseChange)
					{
						arrReverseChanges.push(oReverseChange);
						oReverseChange.SetReverted(true);
					}
				}
			}
		}

		this.OwnRanges.length = this.OwnRanges.length - 1;

		return arrReverseChanges;
	};
	CCollaborativeHistory.prototype.CommuteContentChange = function(oChange, nStartPosition, arrChanges)
	{
		var arrChangesForProceed = this.Changes;
		if (arrChanges && arrChanges.length > 0)
			arrChangesForProceed = arrChanges;

		var arrActions          = oChange.ConvertToSimpleActions();
		var arrCommutateActions = [];

		for (var nActionIndex = arrActions.length - 1; nActionIndex >= 0; --nActionIndex)
		{
			var oAction = arrActions[nActionIndex];
			var oResult = oAction;

			for (var nIndex = nStartPosition, nOverallCount = arrChangesForProceed.length; nIndex < nOverallCount; ++nIndex)
			{
				var oTempChange = arrChangesForProceed[nIndex];
				if (!oTempChange)
					continue;

				if (oChange.IsRelated(oTempChange) &&  true !== oTempChange.IsReverted())
				{
					var arrOtherActions = oTempChange.ConvertToSimpleActions();
					for (var nIndex2 = 0, nOtherActionsCount2 = arrOtherActions.length; nIndex2 < nOtherActionsCount2; ++nIndex2)
					{
						var oOtherAction = arrOtherActions[nIndex2];

						if (false === this.CommuteContentChangeActions(oAction, oOtherAction))
						{
							arrOtherActions.splice(nIndex2, 1);
							oResult = null;
							break;
						}
					}

					oTempChange.ConvertFromSimpleActions(arrOtherActions);
				}

				if (!oResult)
					break;

			}

			if (null !== oResult)
				arrCommutateActions.push(oResult);
		}
		
		if (arrCommutateActions.length <= 0)
			return false;

		oChange.ConvertFromSimpleActions(arrCommutateActions, true);
		return true;
	};
	CCollaborativeHistory.prototype.CommuteContentChangeActions = function(oActionL, oActionR)
	{
		if (oActionL.Add)
		{
			if (oActionR.Add)
			{
				if (oActionL.Pos >= oActionR.Pos)
					oActionL.Pos++;
				else
					oActionR.Pos--;
			}
			else
			{
				if (oActionL.Pos > oActionR.Pos)
					oActionL.Pos--;
				else if (oActionL.Pos === oActionR.Pos)
					return false;
				else
					oActionR.Pos--;
			}
		}
		else
		{
			if (oActionR.Add)
			{
				if (oActionL.Pos >= oActionR.Pos)
					oActionL.Pos++;
				else
					oActionR.Pos++;
			}
			else
			{
				if (oActionL.Pos > oActionR.Pos)
					oActionL.Pos--;
				else
					oActionR.Pos++;
			}
		}

		return true;
	};
	CCollaborativeHistory.prototype.CommutePropertyChange = function(oClass, oChange, nStartPosition)
	{
		// In GoogleDocs if 2 users modify the same property of the same class, then Undo works
		// for both. For example, the first user sets paragraph align to center (initially left), the second one after
		// that sets its to right alignment. Then on Undo the first user returns paragraph to left alignment, and the
		// second to center, regardless of the sequence in which they call Undo.
		// We do the same: always roll back property changes, even if this property was changed in subsequent changes.

		// Alternative here: don't roll back the property if it was changed in one of the subsequent actions. (for this
		// option to work, you need to implement the IsRelated function for all changes).

		// // This means it's a property change. We iterate through all subsequent changes and check if this
		// // property was changed for this class, if so, then this change cannot be commuted.
		// for (var nIndex = nStartPosition, nOverallCount = this.Changes.length; nIndex < nOverallCount; ++nIndex)
		// {
		// 	var oTempChange = this.Changes[nIndex];
		// 	if (!oTempChange || !oTempChange.IsChangesClass || !oTempChangeIsChangesClass())
		// 		continue;
		//
		// 	if (oChange.IsRelated(oTempChange))
		// 		return false;
		// }


		return true;
	};

	CCollaborativeHistory.prototype.CommuteRelated = function(oClass, oChange, nStartPosition)
	{
		return true;
	}
	CCollaborativeHistory.prototype.CreateLocalHistoryPointByReverseChanges = function(reverseChanges)
	{
		let localHistory = AscCommon.History;

		let pointIndex = localHistory.CreateNewPointToCollectChanges(AscDFH.historydescription_Collaborative_Undo);
		
		// Before calling this method, we ran all reverseChanges through Load, where
		// position changes due to these changes were already marked
		this.CoEditing.StopTrackingPositions();
		for (let index = 0, count = reverseChanges.length; index < count; ++index)
		{
			let change = reverseChanges[index];
			localHistory.Add(change);
		}
		this.CoEditing.StartTrackingPositions();

		this.CorrectReveredChanges(reverseChanges);

		localHistory.Update_PointInfoItem(pointIndex, pointIndex, pointIndex, 0, null);
		localHistory.ConvertPointItemsToSimpleChanges(pointIndex);

		return localHistory.Points[pointIndex];
	};
	CCollaborativeHistory.prototype.CorrectReveredChanges = function(arrReverseChanges)
	{
		let oLogicDocument = this.CoEditing.GetLogicDocument();

		// It may happen that in some DocumentContent classes all elements were deleted, or
		// in the Paragraph class the end-of-paragraph mark was deleted. We need to check all classes and if
		// necessary, make additional changes to normalize them

		var mapDrawings         = {};
		var mapDocumentContents = {};
		var mapParagraphs       = {};
		var mapRuns             = {};
		var mapTables           = {};
		var mapGrObjects        = {};
		var mapSlides           = {};
		var mapLayouts          = {};
		var mapTimings          = {};
		var bChangedLayout      = false;
		var bAddSlides          = false;
		var mapAddedSlides      = {};
		var mapCommentsToDelete = {};
		const mapSmartArtShapes = {};

		for (let nIndex = 0, nCount = arrReverseChanges.length; nIndex < nCount; ++nIndex)
		{
			var oChange = arrReverseChanges[nIndex];
			var oClass  = oChange.GetClass();

			if (oClass instanceof AscCommonWord.CDocument || oClass instanceof AscCommonWord.CDocumentContent)
			{
				mapDocumentContents[oClass.Get_Id()] = oClass;
			}
			else if (oClass instanceof AscWord.Paragraph)
			{
				mapParagraphs[oClass.Get_Id()] = oClass;
			}
			else if (oClass.IsParagraphContentElement && true === oClass.IsParagraphContentElement() && true === oChange.IsContentChange() && oClass.GetParagraph())
			{
				const oParagraph = oClass.GetParagraph();
				mapParagraphs[oParagraph.Get_Id()] = oParagraph;
				if (oClass instanceof AscCommonWord.ParaRun)
					mapRuns[oClass.Get_Id()] = oClass;
				const oSmartArtShape = oParagraph.IsInsideSmartArtShape(true);
				if (oSmartArtShape)
				{
					mapSmartArtShapes[oSmartArtShape.Get_Id()] = oSmartArtShape;
				}
			}
			else if (oClass && oClass.parent && oClass.parent instanceof AscCommonWord.ParaDrawing)
			{
				mapDrawings[oClass.parent.Get_Id()] = oClass.parent;
			}
			else if (oClass instanceof AscCommonWord.ParaDrawing)
			{
				mapDrawings[oClass.Get_Id()] = oClass;
			}
			else if (oClass instanceof AscCommonWord.ParaRun)
			{
				mapRuns[oClass.Get_Id()] = oClass;
				const oSmartArtShape = oClass.IsInsideSmartArtShape(true);
				if (oSmartArtShape)
				{
					mapSmartArtShapes[oSmartArtShape.Get_Id()] = oSmartArtShape;
				}
			}
			else if (oClass instanceof AscCommonWord.CTable)
			{
				mapTables[oClass.Get_Id()] = oClass;
			}
			else if (oClass instanceof AscFormat.CShape
				|| oClass instanceof AscFormat.CImageShape
				|| oClass instanceof AscFormat.CChartSpace
				|| oClass instanceof AscFormat.CGroupShape
				|| oClass instanceof AscFormat.CGraphicFrame)
			{
				mapGrObjects[oClass.Get_Id()] = oClass;
				let oParent                   = oClass.parent;
				if (oParent && oParent.timing)
				{
					mapTimings[oParent.timing.Get_Id()] = oParent.timing;
				}
			}
			else if (typeof AscCommonSlide !== "undefined" && AscCommonSlide.Slide && oClass instanceof AscCommonSlide.Slide)
			{
				mapSlides[oClass.Get_Id()] = oClass;
				if (oClass.timing)
				{
					mapTimings[oClass.timing.Get_Id()] = oClass.timing;
				}
			}
			else if (typeof AscCommonSlide !== "undefined" && AscCommonSlide.SlideLayout && oClass instanceof AscCommonSlide.SlideLayout)
			{
				mapLayouts[oClass.Get_Id()] = oClass;
				if (oClass.timing)
				{
					mapTimings[oClass.timing.Get_Id()] = oClass.timing;
				}
				bChangedLayout = true;
			}
			else if (typeof AscCommonSlide !== "undefined" && AscCommonSlide.CPresentation && oClass instanceof AscCommonSlide.CPresentation)
			{
				if (oChange.Type === AscDFH.historyitem_Presentation_RemoveSlide || oChange.Type === AscDFH.historyitem_Presentation_AddSlide)
				{
					bAddSlides = true;
					for (var i = 0; i < oChange.Items.length; ++i)
					{
						mapAddedSlides[oChange.Items[i].Get_Id()] = oChange.Items[i];
					}
				}
			}
			else if (AscDFH.historyitem_ParaComment_CommentId === oChange.Type)
			{
				mapCommentsToDelete[oChange.New] = oClass;
			}
			else if (oClass.isAnimObject)
			{
				let oTiming = oClass.getTiming();
				if (oTiming)
				{
					mapTimings[oTiming.Get_Id()] = oTiming;
				}
			}
		}


		if (bAddSlides)
		{
			for (var i = oLogicDocument.Slides.length - 1; i > -1; --i)
			{
				if (mapAddedSlides[oLogicDocument.Slides[i].Get_Id()] && !oLogicDocument.Slides[i].Layout)
				{
					oLogicDocument.removeSlide(i);
				}
			}
		}

		for (var sId in mapSlides)
		{
			if (mapSlides.hasOwnProperty(sId))
			{
				mapSlides[sId].correctContent();
			}
		}

		if (bChangedLayout)
		{
			for (var i = oLogicDocument.Slides.length - 1; i > -1; --i)
			{
				var Layout = oLogicDocument.Slides[i].Layout;
				if (!Layout || mapLayouts[Layout.Get_Id()])
				{
					if (!oLogicDocument.Slides[i].CheckLayout())
					{
						oLogicDocument.removeSlide(i);
					}
				}
			}
		}


		for (var sId in mapGrObjects)
		{
			var oShape = mapGrObjects[sId];
			if (!oShape.checkCorrect())
			{
				oShape.setBDeleted(true);
				if (oShape.group)
				{
					oShape.group.removeFromSpTree(oShape.Get_Id());
				}
				else if (AscFormat.Slide && (oShape.parent instanceof AscFormat.Slide))
				{
					oShape.parent.removeFromSpTreeById(oShape.Get_Id());
				}
				else if (AscCommonWord.ParaDrawing)
				{
					if(oShape.parent instanceof AscCommonWord.ParaDrawing)
					{
						mapDrawings[oShape.parent.Get_Id()] = oShape.parent;
					}
					if(oShape.oldParent && oShape.oldParent instanceof AscCommonWord.ParaDrawing)
					{
						mapDrawings[oShape.oldParent.Get_Id()] = oShape.oldParent;
						oShape.oldParent = undefined;
					}
				}
			}
			else
			{
				if (oShape.resetGroups)
				{
					oShape.resetGroups();
				}
			}
		}
		var oDrawing;
		for (var sId in mapDrawings)
		{
			if (mapDrawings.hasOwnProperty(sId))
			{
				oDrawing = mapDrawings[sId];
				if (!oDrawing.CheckCorrect())
				{
					var oParentParagraph = oDrawing.Get_ParentParagraph();
					let oParentRun = oDrawing.Get_Run();
					if(oParentRun)
					{
						mapRuns[oParentRun.Id] = oParentRun;
					}
					oDrawing.PreDelete();
					oDrawing.Remove_FromDocument(false);
					if (oParentParagraph)
					{
						mapParagraphs[oParentParagraph.Get_Id()] = oParentParagraph;
					}
				}
			}
		}

		for (var sId in mapRuns)
		{
			if (mapRuns.hasOwnProperty(sId))
			{
				var oRun = mapRuns[sId];
				for (var nIndex = oRun.Content.length - 1; nIndex > -1; --nIndex)
				{
					if (oRun.Content[nIndex] instanceof AscCommonWord.ParaDrawing)
					{
						if (!oRun.Content[nIndex].CheckCorrect())
						{
							oRun.Remove_FromContent(nIndex, 1, false);
							if (oRun.Paragraph)
							{
								mapParagraphs[oRun.Paragraph.Get_Id()] = oRun.Paragraph;
							}
						}
					}
				}
			}
		}
		for (let sId in mapSmartArtShapes)
		{
			const oSmartArtShape = mapSmartArtShapes[sId];
			oSmartArtShape.correctSmartArtUndo();
		}
		for (var sId in mapTables)
		{
			var oTable = mapTables[sId];
			for (var nCurRow = oTable.Content.length - 1; nCurRow >= 0; --nCurRow)
			{
				var oRow = oTable.Get_Row(nCurRow);
				if (oRow.Get_CellsCount() <= 0)
					oTable.private_RemoveRow(nCurRow);
			}

			if (oTable.Parent instanceof AscCommonWord.CDocument || oTable.Parent instanceof AscCommonWord.CDocumentContent)
				mapDocumentContents[oTable.Parent.Get_Id()] = oTable.Parent;
		}

		for (var sId in mapDocumentContents)
		{
			var oDocumentContent = mapDocumentContents[sId];
			var nContentLen      = oDocumentContent.Content.length;
			for (var nIndex = nContentLen - 1; nIndex >= 0; --nIndex)
			{
				var oElement = oDocumentContent.Content[nIndex];
				if ((AscCommonWord.type_Paragraph === oElement.GetType() || AscCommonWord.type_Table === oElement.GetType()) && oElement.Content.length <= 0)
				{
					oDocumentContent.Remove_FromContent(nIndex, 1);
				}
			}

			nContentLen = oDocumentContent.Content.length;
			if (nContentLen <= 0 || AscCommonWord.type_Paragraph !== oDocumentContent.Content[nContentLen - 1].GetType())
			{
				var oNewParagraph = new AscWord.Paragraph(oDocumentContent, 0, 0, 0, 0, 0, false);
				oDocumentContent.Add_ToContent(nContentLen, oNewParagraph);
			}
		}

		for (var sId in mapParagraphs)
		{
			var oParagraph = mapParagraphs[sId];
			oParagraph.CheckParaEnd();
			oParagraph.Correct_Content(null, null, true);
		}

		for (var sId in mapTimings)
		{
			if (mapTimings.hasOwnProperty(sId))
			{
				let oTiming = mapTimings[sId];
				oTiming.checkCorrect();
			}
		}
		if (oLogicDocument && oLogicDocument.IsDocumentEditor())
		{
			for (var sCommentId in mapCommentsToDelete)
			{
				oLogicDocument.RemoveComment(sCommentId, false, false);
			}
		}
	};
	//------------------------------------------------------------------------------------------------------------------
	/**
	 * Own changes segment in the common array
	 * @param position
	 * @param length
	 * @constructor
	 */
	function COwnRange(position, length)
	{
		this.Position = position;
		this.Length   = length;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CCollaborativeHistory = CCollaborativeHistory;
	
	CCollaborativeHistory.prototype["GetGlobalPointCount"]     = CCollaborativeHistory.prototype.GetGlobalPointCount;
	CCollaborativeHistory.prototype["getGlobalPointIndex"]     = CCollaborativeHistory.prototype.GetGlobalPointIndex;
	CCollaborativeHistory.prototype["moveToPoint"]             = CCollaborativeHistory.prototype.MoveToPoint;
	CCollaborativeHistory.prototype["recoverDeletedText"]      = CCollaborativeHistory.prototype.RecoverDeletedText;
	CCollaborativeHistory.prototype["undoDeletedTextRecovery"] = CCollaborativeHistory.prototype.UndoDeletedTextRecovery;
	
})(window);
