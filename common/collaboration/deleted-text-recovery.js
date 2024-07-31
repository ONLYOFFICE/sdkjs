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

(function (window)
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
	 * Инициализация и создание промежуточных данных для отображения удаленного текста в текущей ревизии
	 */
	DeletedTextRecovery.prototype.HandleChanges = function()
	{
		AscCommon.CollaborativeEditing.CoHistory.SplitChangesByPoints();
		
		let oCoHistory			= AscCommon.CollaborativeEditing.CoHistory;
		let arrChangesList		= oCoHistory.Changes;
		let arrPointsList		= AscCommon.CollaborativeEditing.CoHistory.ChangesSplitByPoints;
		let nIndex				= AscCommon.CollaborativeEditing.CoHistory.curChangeIndex + 1;
		let arrChanges			= arrChangesList.slice(0, arrPointsList[nIndex]);

		if (!arrChanges || !arrChanges.length)
			return;

		let arrDelChanges	= [];

		// разделяем по типу изменений
		for (let i = 0; i < arrChanges.length; i++)
		{
			let oCurrentChange = arrChanges[i];

			if (oCurrentChange instanceof AscCommon.CChangesTableIdDescription)
			{
				if (arrDelChanges.length > 0 && arrDelChanges[arrDelChanges.length - 1].length === 1)
					arrDelChanges.length = arrDelChanges.length - 1;

				arrDelChanges.push([oCurrentChange]);
			}

			if (oCurrentChange instanceof CChangesRunRemoveItem || oCurrentChange instanceof CChangesParagraphRemoveItem || oCurrentChange instanceof CChangesDocumentRemoveItem)
			{
				if (arrDelChanges.length > 0)
				{
					let oLast = arrDelChanges[arrDelChanges.length - 1];
					oLast.push(oCurrentChange);
				} else
				{
					arrDelChanges.push([oCurrentChange]);
				}
			} else if (oCurrentChange instanceof CChangesRunAddItem || oCurrentChange instanceof CChangesParagraphAddItem || oCurrentChange instanceof CChangesDocumentAddItem)
			{
				if (arrDelChanges.length > 0)
				{
					let oLast = arrDelChanges[arrDelChanges.length - 1];
					oLast.push(oCurrentChange);
				}
				else
				{
					arrDelChanges.push([oCurrentChange]);
				}
			}
		}

		// если блок изменений состоит только из CChangesTableIdDescription, убираем этот блок
		for (let i = 0; i < arrDelChanges.length; i++)
		{
			if (arrDelChanges[i].length === 1 && arrDelChanges[i][0] instanceof AscCommon.CChangesTableIdDescription)
				arrDelChanges[i].length = 0;
		}

		this.m_RewiewDelPoints = arrDelChanges;
	};
	/**
	 * Отображаем удаленный текст в текущей точки истории ревизии
	 * @return {boolean}
	 */
	DeletedTextRecovery.prototype.RecoverDeletedText = function()
	{
		this.UndoShowDelText();
		return this.ShowDelText();
	};
	/**
	 * Получаем подготовленные данные, разбитые по точкам
	 * @return {*[]}
	 */
	DeletedTextRecovery.prototype.GetChanges = function()
	{
		let arrInput = [];

		for (let nCounter = 0; nCounter < this.m_RewiewDelPoints.length; nCounter++)
		{
			let arrCurr = this.m_RewiewDelPoints[nCounter];
			if (nCounter < AscCommon.CollaborativeEditing.CoHistory.curChangeIndex)
			{
				let arrTemp = [];
				for (let j = 0; j < arrCurr.length; j++)
				{
					arrTemp.push(arrCurr[j]);
				}

				if (arrTemp.length > 0)
					arrInput.push(arrTemp);
			}
		}

		return arrInput;
	};
	/**
	 * Отменяем заданные изменения
	 * @param arrInputChanges
	 * @return {*[]}
	 */
	DeletedTextRecovery.prototype.RedoUndoChanges = function (arrInputChanges)
	{
		let oAddText	= new AddTextPositions();
		let oRemoveText	= new RemoveTextPositions();
		arrInputChanges	= arrInputChanges.reverse();

		let delChanges	= [];

		for (let i = 0; i < arrInputChanges.length; i++)
		{
			let arrChanges = arrInputChanges[i].reverse();

			for (let y = 0; y < arrChanges.length; y++)
			{
				let oCurChange = arrChanges[y];

				if (oCurChange instanceof AscCommon.CChangesTableIdDescription)
					continue;

				oCurChange = oCurChange.Copy();

				if (oCurChange.ConvertToSimpleChanges && oCurChange.Items.length > 1)
					oCurChange = oCurChange.ConvertToSimpleChanges();

				if (!Array.isArray(oCurChange))
					oCurChange = [oCurChange];

				for (let k = 0; k < oCurChange.length; k++)
				{
					let oCur = oCurChange[k];
					if (arrChanges[y].Items.length > 1)
						oCur.Pos += k;

					oRemoveText.ProceedChange(oCur)

					if (oCur instanceof CChangesRunAddItem || oCur instanceof CChangesParagraphAddItem || oCur instanceof CChangesDocumentAddItem)
					{
						if (oAddText.Check(oCur, oRemoveText))
							oAddText.AddChange(oCur)
					}
					else if (oCur instanceof CChangesRunRemoveItem || oCur instanceof CChangesParagraphRemoveItem || oCur instanceof CChangesDocumentRemoveItem)
					{
						if (oCur.UseArray)
						{
							oRemoveText.AddToClass(oCur.Class, oCur, oCur.PosArray[0]);
						}
						else
						{
							oRemoveText.AddToClass(oCur.Class, oCur, oCur.Pos);
						}

						let oCurrentRun = oCur.Class;
						if (oCurrentRun.CollaborativeMarks)
						{
							let arrMarks = oCurrentRun.CollaborativeMarks.Ranges;
							for (let p = 0; p < arrMarks.length; p++)
							{
								let oCurrentMark = arrMarks[p];
								if (oCurrentMark.PosS >= oCur.PosArray[0])
								{
									oCurrentMark.PosS++;
									oCurrentMark.PosE++;
								}
							}
						}
					}
				}
			}
		}

		let arr = oRemoveText.GetLinear();
		for (let i = 0; i < arr.length; i++)
		{
			this.RedoUndoChange(arr[i], false, delChanges);
		}

		let result = {
			data: oRemoveText.ProceedPositions(this),
			classes: oRemoveText.oClasses
		};

		return [delChanges, result];
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

		if (!oReviewInfoParent || !oReviewInfoParent.ReviewInfo)
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

		if (oReviewInfoParent.ReviewType !== 1)
		{
			let oCurrentReviewType			= oReviewInfoParent.GetReviewInfo().Copy();
			oCurrentReviewType.UserId		= this.userId;
			oCurrentReviewType.UserName		= this.userName;
			oCurrentReviewType.DateTime		= this.userTime;

			oReviewInfoParent.SetReviewTypeWithInfo(1, oCurrentReviewType, false);
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
	DeletedTextRecovery.prototype.ColorText = function ()
	{
		for (let i = 0; i < this.arrColor.length; i++)
		{
			let oCurCollaborativeMark = this.arrColor[i];

			let nStartPos				= oCurCollaborativeMark.PosS;
			let nEndPos					= oCurCollaborativeMark.PosE;
			let oColor					= oCurCollaborativeMark.Color;
			let oRun					= oCurCollaborativeMark.oCurrentRun;

			oRun.CollaborativeMarks.Add(nStartPos, nEndPos, oColor);
		}

		this.arrColor = [];
	}
	DeletedTextRecovery.prototype.UndoShowDelText = function (isNavigation)
	{
		let localHistory	= AscCommon.History;
		let oLastPoint		= localHistory.Points[localHistory.Points.length - 1];

		if ((oLastPoint && oLastPoint.Description === AscDFH.historydescription_Collaborative_DeletedTextRecovery) || (isNavigation && oLastPoint && oLastPoint.Description === AscDFH.historydescription_Collaborative_MoveByHistory))
		{
			let arrChanges = localHistory.UndoLastPoint();
			this.document.RecalculateByChanges(arrChanges);
			localHistory.Remove_LastPoint();

			this.ColorText();
			return true
		}
	};

	function AddTextPositions()
	{
		this.data = {};
		this.wait = {};

		this.Check = function (oCurChange, oRemove)
		{
			if (oCurChange.Class)
			{
				let strCurrentId	= oCurChange.Class.Id;
				let arrRemData		= oRemove.data[strCurrentId];
				if (!arrRemData)
					return true;

				let addItem			= oCurChange.Items[0];

				for (let i = 0; i < arrRemData.length; i++)
				{
					let oCurrentRemItem = arrRemData[i];

					if (oCurChange.UseArray && oCurChange.PosArray[0] === oCurrentRemItem.pos && addItem.Value === oCurrentRemItem.item.Items[0].Value)
					{
						if (oRemove.data[strCurrentId])
						{
							let currentArr = oRemove.data[strCurrentId];
							for (let i = 0; i < currentArr.length; i++)
							{
								if (oCurChange.UseArray)
								{
									if (oCurChange.PosArray[0] < currentArr[i].pos)
									{
										if (oCurChange.UseArray)
											oCurChange.PosArray[0]--
										else
											oCurChange.Pos--;
									}
								}
							}
						}

						let addArr = this.data[strCurrentId];
						if (addArr && addArr.length > 0)
						{
							for (let j = 0; j < addArr.length; j++)
							{
								let oCur = addArr[j];

								if (oCurChange.PosArray[0] < oCur.pos)
								{
									if (oCurChange.UseArray)
										oCur.change.PosArray[0]--
									else
										oCur.change.Pos--;

									oCur.pos--;
								}
							}
						}

						arrRemData.splice(i, 1);
						return false;
					}
					else if (!oCurChange.UseArray && oCurChange.Pos === oCurrentRemItem.pos && addItem.Value === oCurrentRemItem.item.Items[0].Value)
					{
						if (oRemove.data[strCurrentId])
						{
							let currentArr = oRemove.data[strCurrentId];
							for (let i = 0; i < currentArr.length; i++)
							{
								if (oCurChange.UseArray)
								{
									if (oCurChange.PosArray[0] < currentArr[i].pos)
									{
										if (oCurChange.UseArray)
											oCurChange.PosArray[0]--
										else
											oCurChange.Pos--;
									}
								}
							}
						}

						arrRemData.splice(i, 1);
						return false;
					}
				}
			}

			return true;
		}
		this.AddChange = function (oChange)
		{
			let strCurrentId = oChange.Class.Id;

			if (strCurrentId)
			{
				if (!this.data[strCurrentId])
					this.data[strCurrentId] = []

				let oWait;

				if (this.wait[strCurrentId])
					oWait = this.wait[strCurrentId].filter( function (oItem) { return oItem.pos === oChange.pos});

				let nChange = oChange.PosArray ? oChange.PosArray[0] : oChange.Pos;

				// repeat for PosArray array

				if (oWait)
					this.data[strCurrentId][oWait.nCount] = {change: oChange, pos: nChange}
				else
					this.data[strCurrentId].push({change: oChange, pos: nChange});
			}
		}
	}
	function RemoveTextPositions()
	{
		this.data		= {};
		this.oClasses	= {};
		this.arrClasses	= [];

		this.AddToClass = function (oClass, oItem, Pos)
		{
			if (!this.data[oClass.Id])
				this.data[oClass.Id] = [];

			this.data[oClass.Id].push({item: oItem, pos: Pos});
		}
		this.GetLinear = function ()
		{
			let arrOutput = []
			let arrKeys = Object.keys(this.data);

			for (let nKey = 0; nKey < arrKeys.length; nKey++)
			{
				let strCurrentKey = arrKeys[nKey];
				let arrCurrentRunData = this.data[strCurrentKey];


				for (let i = 0; i < arrCurrentRunData.length; i++)
				{
					arrOutput.push(arrCurrentRunData[i].item);
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
			const transformedObject = CollapsePositions(this.data);
			return transformedObject
		}
	}
	function CollapsePositions (oInput)
	{
		const transformedObject = {};
		for (const key in oInput)
		{
			if (oInput.hasOwnProperty(key))
			{
				const values = oInput[key];
				const pairs = [];
				let nStart = null;
				let nEnd = null;
				let decreasingSequence = false;

				for (let i = 0; i < values.length; i++)
				{
					const value = values[i];

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
						pairs.push({ nStart, nEnd });
						nStart = value;
						nEnd = value;
						decreasingSequence = false;
					}
				}

				if (nStart !== null && nEnd !== null)
					pairs.push({ nStart, nEnd });

				transformedObject[key] = pairs;
			}
		}
		return transformedObject;
	}
	//--------------------------------------------------------export----------------------------------------------------
	AscCommon.DeletedTextRecovery = DeletedTextRecovery;

})(window);
