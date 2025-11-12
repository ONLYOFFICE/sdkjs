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

(function ()
{
	/**
	 * Класс, восстанавливающий удаленные части документа
	 * @param {AscWord.Document} logicDocument
	 * @constructor
	 */
	function DeletedTextRecovery(logicDocument)
	{
		this.document = logicDocument;
		
		/**
		 * Список всех изменений связанных с удалением текста
		 * @type {*[]}
		 */
		this.m_RewiewDelPoints		= [];
		this.arrColor				= [];

		this.userId					= undefined;
		this.userName				= undefined;
		this.userTime				= undefined;
	}
	/**
	 * Восстанавливаем удаленный текст в текущей точке истории версий
	 * @return {boolean}
	 */
	DeletedTextRecovery.prototype.RecoverDeletedText = function()
	{
		this.UndoRecoveredText();
		return this.ShowDelText();
	};
	/**
	 * Отменяем восстановление удаленного текста, если оно было
	 */
	DeletedTextRecovery.prototype.UndoRecoveredText = function ()
	{
		if (!this.HaveRecoveredText())
			return;
		
		let localHistory = AscCommon.History;
		let changes = localHistory.Undo();
		this.document.UpdateAfterUndoRedo(changes);
		localHistory.ClearRedo();
	};
	/**
	 * Запрашивем,есть ли восстановленный удаленный текст
	 * @return {boolean}
	 */
	DeletedTextRecovery.prototype.HaveRecoveredText = function()
	{
		let localHistory = AscCommon.History;
		let lastPoint    = localHistory.Points[localHistory.Points.length - 1];
		return !!(lastPoint && lastPoint.Description === AscDFH.historydescription_Collaborative_DeletedTextRecovery);
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * Инициализация и создание промежуточных данных для отображения удаленного текста в текущей ревизии
	 */
	DeletedTextRecovery.prototype.HandleChanges = function()
	{
		AscCommon.CollaborativeEditing.CoHistory.SplitChangesByPoints();
		
		let oCoHistory			= AscCommon.CollaborativeEditing.CoHistory;
		let arrChangesList		= oCoHistory.Changes;
		let arrPointsList		= AscCommon.CollaborativeEditing.CoHistory.ChangesSplitByPoints;
		let nIndex				= AscCommon.CollaborativeEditing.CoHistory.curChangeIndex;
		let arrChanges			= arrChangesList.slice(0, arrPointsList[nIndex]);
		
		if (!arrChanges || !arrChanges.length)
			return;
		
		this.m_RewiewDelPoints = arrChanges;
	};
	/**
	 * Получаем подготовленные данные, разбитые по точкам
	 * @return {*[]}
	 */
	DeletedTextRecovery.prototype.GetChanges = function()
	{
		let arr = this.m_RewiewDelPoints.slice();

		let arrOutput = []
		for (let i = 0; i < arr.length; i++)
		{
			let oChange = arr[i];
			if (oChange.ConvertToSimpleChanges)
			{
				let arrSplitChange = oChange.ConvertToSimpleChanges();
				if (arrSplitChange && arrSplitChange.length)
				{
					for (let j = 0; j < arrSplitChange.length; j++)
					{
						arrOutput.push(arrSplitChange[j])
					}
				} else
				{
					arrOutput.push(oChange);
				}
			}
			else
				arrOutput.push(oChange);

		}
		return arrOutput;
	};
	DeletedTextRecovery.prototype.GetRemoveTextChanges = function (arrInputChanges, oRemoveText)
	{
		let oAddText		= new AddTextPositions();

		// отбираем удаленный текст связанный с текущей ревизией
		for (let i = 0; i < arrInputChanges.length; i++)
		{
			let oCurChange		= arrInputChanges[i];

			// пропускаем все изменения связанные со сплитом
			if (oCurChange instanceof CChangesRunOnStartSplit)
			{
				while (oCurChange && !(oCurChange instanceof CChangesRunOnEndSplit))
				{
					i++;
					oCurChange		= arrInputChanges[i];
				}
			}

			if (oCurChange === undefined)
				break;

			if (oCurChange instanceof AscCommon.CChangesTableIdDescription || !oCurChange.Copy)
				continue;

			let oNewCurChange	= oCurChange.Copy();

			oRemoveText.ProceedChange(oNewCurChange);

			if (oNewCurChange instanceof CChangesRunAddItem || oNewCurChange instanceof CChangesParagraphAddItem || oNewCurChange instanceof CChangesDocumentAddItem)
				oAddText.Add(oCurChange.Class, oCurChange, oCurChange.UseArray ? oCurChange.PosArray[0] : oCurChange.Pos, i);
			else if (oNewCurChange instanceof CChangesRunRemoveItem || oNewCurChange instanceof CChangesParagraphRemoveItem || oNewCurChange instanceof CChangesDocumentRemoveItem)
				oRemoveText.Check(oAddText, oCurChange, i);
		}
		oRemoveText.DelDuplicate();
	}
	DeletedTextRecovery.prototype.CommuteChanges = function (arrInputChanges, arrSaveData, oRemoveText)
	{
		// коммутируем изменения
		let arrRevInput = arrInputChanges;

		let arrDelChangesForCommute = oRemoveText.GetArrayChanges();

		for (let j = 0; j < arrDelChangesForCommute.length; j++)
		{
			let oCurItem	= arrDelChangesForCommute[j];
			let nPos		= oCurItem.nIndex;
			let oChange		= oCurItem.item;

			if (oChange.IsContentChange())
			{
				let _oChange = oChange.Copy();

				if (AscCommon.CollaborativeEditing.CoHistory.CommuteContentChange(_oChange, nPos, arrRevInput))
					arrSaveData.push(_oChange);
			}
			else
			{
				arrSaveData.push(oChange);
			}
		}

		oRemoveText.ResetData();
	}
	/**
	 * Отменяем заданные изменения
	 * @param arrInputChanges
	 * @return {*[]}
	 */
	DeletedTextRecovery.prototype.RedoUndoChanges = function (arrInputChanges)
	{
		let oRemoveText		= new RemoveTextPositions();
		let arrChanges		= [];
		let arrDelChanges	= [];

		this.GetRemoveTextChanges(arrInputChanges, oRemoveText);
		this.CommuteChanges(arrInputChanges, arrDelChanges, oRemoveText);

		for (let i = 0; i < arrDelChanges.length; i++)
		{
			this.RedoUndoChange(arrDelChanges[i], false, arrChanges);
			let nPos = arrDelChanges[i].UseArray
				? arrDelChanges[i].PosArray[0]
				: arrDelChanges[i].Pos;

			oRemoveText.Add(
				arrDelChanges[i].Class,
				arrDelChanges[i],
				nPos,
				i
			);
		}

		let result = {
			data: oRemoveText.ProceedPositions(this),
			classes: oRemoveText.oClasses
		};

		return [arrChanges, result];
	};
	DeletedTextRecovery.prototype.ShowDelText = function ()
	{
		let versionHistory = this.document.GetApi().getVersionHistory();
		if (!versionHistory)
			return false;
		
		this.HandleChanges();
		let arrInput = this.GetChanges();
		if (arrInput.length === 0)
			return false;
		
		let localHistory = AscCommon.History;
		localHistory.Create_NewPoint(AscDFH.historydescription_Collaborative_DeletedTextRecovery);
		
		this.userId   = versionHistory.userId;
		this.userName = versionHistory.userName;
		this.userTime = new Date(versionHistory.dateOfRevision).getTime();

		// отменяем изменения до нужного места (необходимо для перемещения по истории)
		let arrCurrentPoint	= this.RedoUndoChanges(arrInput);
		let delChanges = arrCurrentPoint[0];
		let arrResult  = arrCurrentPoint[1];

		for (let i = 0; i < delChanges.length; i++)
		{
			this.document.History.Add(delChanges[i]);
		}
		this.Split(arrResult);
		this.document.RecalculateByChanges(delChanges);

		this.m_RewiewDelPoints = [];
		return true;
	};
	DeletedTextRecovery.prototype.Split = function (arrInput)
	{
		let data	= arrInput.data;
		let classes	= arrInput.classes;
		let arrKeys	= Object.keys(data);

		for (let nKey = 0; nKey < arrKeys.length; nKey++)
		{
			let strCurrentKey		= arrKeys[nKey];
			let arrCurrentRunData	= data[strCurrentKey];
			let oCurrentRun			= classes[strCurrentKey];

			arrCurrentRunData.sort(function (a, b) { return a.nStart - b.nStart });

			for (let j = arrCurrentRunData.length - 1; j >= 0; j--)
			{
				let oCurrentRule = arrCurrentRunData[j];
				let nStart = oCurrentRule.nStart;
				let nEnd = oCurrentRule.nEnd;

				if (oCurrentRun instanceof CDocument)
				{
					let arrContent	= oCurrentRun.Content;

					for (let j = nStart; j <= nEnd; j++)
					{
						let oCurrentParagraph = arrContent[j];
						if (oCurrentParagraph)
							this.SetReviewInfo(oCurrentParagraph);
					}
				}
				else if (oCurrentRun instanceof Paragraph)
				{
					let arrContent	= oCurrentRun.Content;

					if (nStart === 0 && arrContent.length === nEnd)
					{
						this.SetReviewInfo(oCurrentRun);
					}
					else
					{
						for (let i = nStart; i <= nEnd; i++)
						{
							this.SetReviewInfo(arrContent[i]);
						}
					}
				}
				else if (oCurrentRun instanceof ParaRun)
				{
					let newCollab = [];

					if (oCurrentRun.Content.length === 0 || (nEnd + 1 - nStart) === oCurrentRun.Content.length)
					{
						this.SetReviewInfo(oCurrentRun);
						continue;
					}

					for (let i = 0; i < oCurrentRun.CollaborativeMarks.Ranges.length; i++)
					{
						let oCollab = oCurrentRun.CollaborativeMarks.Ranges[i];
						newCollab.push({PosS: oCollab.PosS, PosE: oCollab.PosE, Color: oCollab.Color, oCurrentRun: oCurrentRun});
					}

					let oParent		= oCurrentRun.GetParent();
					let RunPos		= this.FindPosInParent(oCurrentRun);
					let RightRun	= oCurrentRun.SplitForSpreadCollaborativeMark(nStart);

					oParent.Add_ToContent(RunPos + 1, RightRun);
					let oNewer = RightRun.SplitForSpreadCollaborativeMark(nEnd - nStart + 1);

					oParent.Add_ToContent(RunPos + 2, oNewer);
					this.SetReviewInfo(RightRun);

					for (let i = 0; i < newCollab.length; i++)
					{
						let oCurCollaborativeMark	= newCollab[i];
						this.arrColor.push(oCurCollaborativeMark);
					}
				}
			}
		}
	}
	DeletedTextRecovery.prototype.SetReviewInfo = function (oReviewInfoParent)
	{
		if (!oReviewInfoParent === undefined)
			return;

		// TODO: Проверку по каким классам нужно проходить лучше переделать
		if (!oReviewInfoParent || !oReviewInfoParent.GetReviewInfo)
		{
			if (oReviewInfoParent instanceof ParaMath)
			{
				let oRootContent = oReviewInfoParent.Root.Content;
				for (let i = 0; i < oRootContent.length; i++)
				{
					let oCurrentContent = oRootContent[i];
					this.SetReviewInfo(oCurrentContent);
				}
			}
			else if (oReviewInfoParent.Content.length > 0)
			{
				for (let i = 0; i < oReviewInfoParent.Content.length; i++)
				{
					let oCurrentContent = oReviewInfoParent.Content[i];
					this.SetReviewInfo(oCurrentContent);
				}
			}
			return;
		}

		if (oReviewInfoParent.GetReviewType() !== reviewtype_Remove)
		{
			let reviewInfo = oReviewInfoParent.GetReviewInfo();
			if (!reviewInfo)
				reviewInfo = new AscWord.ReviewInfo();
			else
				reviewInfo = reviewInfo.Copy();
			
			reviewInfo.UserId		= this.userId;
			reviewInfo.UserName		= this.userName;
			reviewInfo.DateTime		= this.userTime;

			oReviewInfoParent.SetReviewTypeWithInfo(reviewtype_Remove, reviewInfo, false);
		}
	};
	DeletedTextRecovery.prototype.FindPosInParent = function(oClass)
	{
		let oParent				= oClass.GetParent();
		let arrParentContent	= oParent.Content;

		for (let i = 0; i < arrParentContent.length; i++)
		{
			if (arrParentContent[i] === oClass)
				return i;
		}
	};
	DeletedTextRecovery.prototype.RedoUndoChange = function (oChange, isRedo, arrToSave)
	{
		if (!oChange)
			return;

		if (oChange.IsContentChange())
		{
			let arrSimpleChanges = oChange.ConvertToSimpleChanges();

			for (let simpleIndex = arrSimpleChanges.length - 1; simpleIndex >= 0; simpleIndex--)
			{
				if (isRedo)
					arrSimpleChanges[simpleIndex].Redo();
				else
					arrSimpleChanges[simpleIndex].Undo();


				let oRevChange = arrSimpleChanges[simpleIndex];

				if (!isRedo)
					oRevChange = oRevChange.CreateReverseChange();

				if (oRevChange)
					arrToSave.push(oRevChange);
			}
		}
		else
		{
			if (isRedo)
				oChange.Redo();
			else
				oChange.Undo();

			let oRevChange = oChange;

			if (!isRedo)
				oRevChange = oRevChange.CreateReverseChange();

			if (oRevChange)
				arrToSave.push(oRevChange);
		}
	};

	function AddTextPositions()
	{
		this.data		= {};

		this.Add = function (oClass, oItem, Pos, nIndex)
		{
			if (!this.data[oClass.Id])
				this.data[oClass.Id] = [];

			this.data[oClass.Id].push({item: oItem, pos: Pos, nIndex: nIndex});
		}
	}
	function RemoveTextPositions()
	{
		this.data		= {};
		this.oClasses	= {};
		this.arrClasses	= [];

		this.Check = function (oAddText, oCurChange, i)
		{
			if (oCurChange.Class)
			{
				let strCurrentId	= oCurChange.Class.Id;
				let arrRemData		= oAddText.data[strCurrentId];

				if (!arrRemData)
				{
					this.Add(oCurChange.Class, oCurChange, oCurChange.UseArray ? oCurChange.PosArray[0] : oCurChange.Pos, i);
					return;
				}

				let addItem			= oCurChange.Items[0];

				for (let i = 0; i < arrRemData.length; i++)
				{
					let oCurrentRemItem = arrRemData[i];

					if (oCurChange.UseArray && oCurChange.PosArray[0] === oCurrentRemItem.pos && addItem.Value === oCurrentRemItem.item.Items[0].Value)
					{
						return false;
					}
					else if (!oCurChange.UseArray && oCurChange.Pos === oCurrentRemItem.pos && addItem.Value === oCurrentRemItem.item.Items[0].Value)
					{
						return false;
					}
				}
			}

			this.Add(oCurChange.Class, oCurChange, oCurChange.UseArray ? oCurChange.PosArray[0] : oCurChange.Pos, i);
		}
		this.DelDuplicate = function ()
		{
			let arrKeys = Object.keys(this.data);

			for (let nKey = 0; nKey < arrKeys.length; nKey++)
			{
				let strCurrentKey = arrKeys[nKey];
				let arrCurrentRunData = this.data[strCurrentKey];

				for (let i = 0; i < arrCurrentRunData.length; i++)
				{
					for (let j = i + 1; j < arrCurrentRunData.length; j++)
					{
						if (arrCurrentRunData[i].item === arrCurrentRunData[j].item)
						{
							arrCurrentRunData.splice(j, 1);
							break
						}
					}
				}
			}
		}
		this.Add = function (oClass, oItem, Pos, nIndex)
		{
			if (!this.data[oClass.Id])
				this.data[oClass.Id] = [];

			this.data[oClass.Id].push({item: oItem, pos: Pos, nIndex: nIndex});
		}
		this.GetArrayChanges = function ()
		{
			let arrOutput = []
			let arrKeys = Object.keys(this.data);

			for (let nKey = 0; nKey < arrKeys.length; nKey++)
			{
				let strCurrentKey = arrKeys[nKey];
				let arrCurrentRunData = this.data[strCurrentKey];

				for (let i = 0; i < arrCurrentRunData.length; i++)
				{
					arrOutput.push(arrCurrentRunData[i]);
				}
			}
			return arrOutput;
		}
		this.ProceedChange = function (oChange)
		{
			if (oChange.Class && !this.oClasses[oChange.Class.Id])
			{
				this.oClasses[oChange.Class.Id] = oChange.Class;
				if (-1 === this.arrClasses.indexOf(oChange.Class.Id))
				{
					this.arrClasses.push(oChange.Class.Id);
				}
			}
		}
		this.FindInParent = function (oClass, oItem)
		{
			let arrParentContent	= oClass.Content;

			for (let nPos = 0; nPos < arrParentContent.length; nPos++)
			{
				if (arrParentContent[nPos] === oItem)
					return nPos;
			}
		}
		this.ProceedPositions = function ()
		{
			for (let nKey = 0; nKey < this.arrClasses.length; nKey++)
			{
				let strCurrentKey			= this.arrClasses[nKey];
				let arrCurrentRunData		= this.data[strCurrentKey];
				if (!arrCurrentRunData)
					continue;
				let oClass					= this.oClasses[strCurrentKey];
				let newArrCurrentRunData	= [];

				for (let i = 0; i < arrCurrentRunData.length; i++)
				{
					let oItem = arrCurrentRunData[i];

					let nPos = this.FindInParent(oClass, oItem.item.Items[0]);
					newArrCurrentRunData.push(nPos);
				}

				let nCurrentPos;
				let nPrevPos;
				let nTempPrevPos;

				for (let nPos = 1; nPos < newArrCurrentRunData.length; nPos++)
				{
					nPrevPos = newArrCurrentRunData[nPos - 1];
					nCurrentPos = newArrCurrentRunData[nPos];

					if (nTempPrevPos === nCurrentPos || nCurrentPos === nPrevPos)
					{
						if (nTempPrevPos)
						{
							nTempPrevPos = undefined;
							newArrCurrentRunData[nPos] = nPrevPos + 1;
							nTempPrevPos = nCurrentPos;
						}
						else
						{
							newArrCurrentRunData[nPos] = nCurrentPos + 1;
							nTempPrevPos = nCurrentPos;
						}
					}
				}
				this.data[strCurrentKey] = newArrCurrentRunData;
			}
			let transformedObject = CollapsePositions(this.data);
			return transformedObject
		}
		this.ResetData = function ()
		{
			this.data = {};
		}
	}
	function CollapsePositions (oInput)
	{
		let transformedObject = {};
		for (let key in oInput)
		{
			if (oInput.hasOwnProperty(key))
			{
				let values = oInput[key];
				let pairs = [];
				let nStart = null;
				let nEnd = null;
				let decreasingSequence = false;

				for (let i = 0; i < values.length; i++)
				{
					let value = values[i];

					if (nStart === null)
					{
						nStart = value;
						nEnd = value;
					}
					else if (value === nEnd + 1)
					{
						nEnd = value;
						decreasingSequence = false;
					}
					else if (value === nEnd - 1)
					{
						nStart = value;
						decreasingSequence = true;
					}
					else
					{
						pairs.push({ nStart : nStart, nEnd: nEnd });
						nStart = value;
						nEnd = value;
						decreasingSequence = false;
					}
				}

				if (nStart !== null && nEnd !== null)
					pairs.push({ nStart : nStart, nEnd: nEnd });

				transformedObject[key] = pairs;
			}
		}
		return transformedObject;
	}
	//--------------------------------------------------------export----------------------------------------------------
	AscCommon.DeletedTextRecovery = DeletedTextRecovery;
	
})();
