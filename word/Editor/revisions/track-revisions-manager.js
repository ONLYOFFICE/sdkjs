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

(function(window)
{
	let elementToCheck = {};
	function checkElementInRevision(element)
	{
		if (!element)
			return;
		
		let logicDocument = element.GetLogicDocument();
		if (!element.GetLogicDocument())
		{
			elementToCheck[element.GetId()] = element;
			return;
		}
		
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return;
		
		logicDocument.GetTrackRevisionsManager().CheckElement(element);
	}
	function haveElementToCheck()
	{
		for (let id in elementToCheck)
		{
			return true;
		}
		
		return false;
	}
	function flushElementToCheck()
	{
		for (let id in elementToCheck)
		{
			let element = elementToCheck[id];
			let logicDocument = element.GetLogicDocument();
			if (!logicDocument || !logicDocument.IsDocumentEditor())
				continue;
			
			logicDocument.GetTrackRevisionsManager().CheckElement(elementToCheck[id]);
		}
		
		elementToCheck = {};
	}
	
	// ВАЖНО: CheckArray - специальный массив-дублер для мапа CheckMap, для более быстрого выполнения функции
	//        ContinueTrackRevisions. Заметим, что на функции CompleteTrackChangesForElements мы не выкидываем
	//        элемент из CheckArray, но выкидываем из CheckMap. Это не страшно, т.к. при последующей проверке по
	//        массиву CheckArray элемент в CheckMap просто не найдется, а из CheckArray мы его выкинем, и, тем самым,
	//        CheckArray и CheckMap синхронизуются.

	/**
	 * Класс для отслеживания изменений в режиме рецензирования
	 * @param {AscWord.CDocument} oLogicDocument
	 * @constructor
	 */
	function CTrackRevisionsManager(oLogicDocument)
	{
		this.LogicDocument = oLogicDocument;

		this.CheckMap   = {};   // Элементы, которые нужно проверить
		this.CheckArray = [];   // Дублирующий массив элементов, которые мы проверяем в ContinueTrackRevisions

		this.Changes        = {};   // Объект с ключом - Id параграфа, в котором лежит массив изменений
		this.ChangesOutline = [];   // Упорядоченный массив с объектами, в которых есть изменения в рецензировании
		this.CurChange      = null; // Текущее изменение
		this.CurElement     = null; // Элемент с текущим изменением

		this.SelectedChanges     = [];    // Список изменений, попавших в выделение
		this.PrevSelectedChanges = [];
		this.PrevShowChanges     = true;

		this.MoveId      = 1;
		this.MoveMarks   = {};
		this.ProcessMove = null;

		this.SkipPreDeleteMoveMarks = false;
	}

	/**
	 * Отправляем элемент на проверку на наличие рецензирования
	 * @param oElement {AscWord.Paragraph | AscWord.CTable}
	 */
	CTrackRevisionsManager.prototype.CheckElement = function(oElement)
	{
		if (!(oElement instanceof AscWord.Paragraph) && !(oElement instanceof AscWord.CTable))
			return;

		let Id = oElement.GetId();
		if (!this.CheckMap[Id])
		{
			this.CheckMap[Id] = oElement;
			this.CheckArray.push(oElement);
		}
	};
	/**
	 * Добавляем изменение в рецензировании по Id элемента
	 * @param sId {string}
	 * @param oChange {CRevisionsChange}
	 */
	CTrackRevisionsManager.prototype.AddChange = function(sId, oChange)
	{
		if (this.private_CheckChangeObject(sId))
			this.Changes[sId].push(oChange);
	};
	/**
	 * Получаем массив изменений заданного элемента
	 * @param sElementId
	 * @returns {CRevisionsChange[]}
	 */
	CTrackRevisionsManager.prototype.GetElementChanges = function(sElementId)
	{
		if (this.Changes[sElementId])
			return this.Changes[sElementId];

		return [];
	};
	CTrackRevisionsManager.prototype.ContinueTrackRevisions = function(isComplete)
	{
		if (this.IsAllChecked())
			return;
		
		flushElementToCheck();

		var nStartTime = performance.now();

		// За раз обрабатываем не больше 500 параграфов либо не больше 4мс по времени,
		// чтобы не подвешивать клиент на открытии файлов
		const nMaxCounter = 500;
		let nCounter      = 0;

		var bNeedUpdate = false;

		let nIndex = this.CheckArray.length - 1;
		for (; nIndex >= 0; --nIndex)
		{
			if (this.private_TrackChangesForSingleElement(this.CheckArray[nIndex].GetId()))
				bNeedUpdate = true;

			if (true !== isComplete)
			{
				++nCounter;
				if (nCounter >= nMaxCounter || (performance.now() - nStartTime) > 4)
					break;
			}
		}

		this.CheckArray.length = nIndex < 0 ? 0 : nIndex;

		// TODO: Сделать нормальную проверку, что текущий параграф или текущее выделение попала в проверенные элементы
		//if (bNeedUpdate)
			this.LogicDocument.UpdateInterface();
	};
	/**
	 * Ищем следующее изменение
	 * @returns {?CRevisionsChange}
	 */
	CTrackRevisionsManager.prototype.GetNextChange = function()
	{
		if (this.CurChange && this.CurChange.IsComplexChange())
		{
			var arrChanges = this.CurChange.GetSimpleChanges();

			this.CurChange  = null;
			this.CurElement = null;

			if (arrChanges.length > 0)
			{
				this.CurChange  = arrChanges[arrChanges.length - 1];
				this.CurElement = this.CurChange.GetElement();

				if (!this.CurElement || !this.Changes[this.CurElement.GetId()])
				{
					this.CurChange  = null;
					this.CurElement = null;
				}
			}
		}

		var oChange = this.private_GetNextChange();
		if (oChange && oChange.IsMove() && !oChange.IsComplexChange())
		{
			oChange        = this.CollectMoveChange(oChange);
			this.CurChange = oChange;
		}

		return oChange;
	};
	CTrackRevisionsManager.prototype.private_GetNextChange = function()
	{
		var oInitialCurChange  = this.CurChange;
		var oInitialCurElement = this.CurElement;

		var oNextElement = null;
		if (null !== this.CurChange && null !== this.CurElement && this.Changes[this.CurElement.GetId()])
		{
			var arrChangesArray = this.Changes[this.CurElement.GetId()];

			var nChangeIndex = -1;
			for (var nIndex = 0, nCount = arrChangesArray.length; nIndex < nCount; ++nIndex)
			{
				if (this.CurChange === arrChangesArray[nIndex])
				{
					nChangeIndex = nIndex;
					break;
				}
			}

			if (-1 !== nChangeIndex && nChangeIndex < arrChangesArray.length - 1)
			{
				this.CurChange = arrChangesArray[nChangeIndex + 1];
				return this.CurChange;
			}

			oNextElement = this.LogicDocument.GetRevisionsChangeElement(1, this.CurElement);
		}
		else
		{
			var oSearchEngine = this.LogicDocument.private_GetRevisionsChangeElement(1, null);
			oNextElement      = oSearchEngine.GetFoundedElement();
			if (null !== oNextElement && oNextElement === oSearchEngine.GetCurrentElement())
			{
				var arrNextChangesArray = this.Changes[oNextElement.GetId()];
				if (arrNextChangesArray && arrNextChangesArray.length > 0)
				{
					if (oNextElement instanceof Paragraph)
					{
						var ParaContentPos = oNextElement.Get_ParaContentPos(oNextElement.IsSelectionUse(), true);
						for (var nChangeIndex = 0, nCount = arrNextChangesArray.length; nChangeIndex < nCount; ++nChangeIndex)
						{
							var ChangeEndPos = arrNextChangesArray[nChangeIndex].get_EndPos();
							if (ParaContentPos.Compare(ChangeEndPos) <= 0)
							{
								this.CurChange  = arrNextChangesArray[nChangeIndex];
								this.CurElement = oNextElement;
								return this.CurChange;
							}
						}
					}
					else if (oNextElement instanceof CTable && oNextElement.IsCellSelection())
					{
						var arrSelectedCells = oNextElement.GetSelectionArray();
						if (arrSelectedCells.length > 0)
						{
							var nTableRow = arrSelectedCells[0].Row;
							for (var nChangeIndex = 0, nCount = arrNextChangesArray.length; nChangeIndex < nCount; ++nChangeIndex)
							{
								var nStartRow = arrNextChangesArray[nChangeIndex].get_StartPos();
								if (nTableRow <= nStartRow)
								{
									this.CurChange  = arrNextChangesArray[nChangeIndex];
									this.CurElement = oNextElement;
									return this.CurChange;
								}
							}
						}
					}

					oNextElement = this.LogicDocument.GetRevisionsChangeElement(1, oNextElement);
				}
			}
		}

		if (null !== oNextElement)
		{
			var arrNextChangesArray = this.Changes[oNextElement.GetId()];
			if (arrNextChangesArray && arrNextChangesArray.length > 0)
			{
				this.CurChange  = arrNextChangesArray[0];
				this.CurElement = oNextElement;
				return this.CurChange;
			}
		}

		if (null !== oInitialCurChange && null !== oInitialCurElement)
		{
			this.CurChange  = oInitialCurChange;
			this.CurElement = oInitialCurElement;
			return oInitialCurChange;
		}

		this.CurChange  = null;
		this.CurElement = null;
		return null;
	};
	/**
	 * Ищем следующее изменение
	 * @returns {?CRevisionsChange}
	 */
	CTrackRevisionsManager.prototype.GetPrevChange = function()
	{
		if (this.CurChange && this.CurChange.IsComplexChange())
		{
			var arrChanges = this.CurChange.GetSimpleChanges();

			this.CurChange  = null;
			this.CurElement = null;

			if (arrChanges.length > 0)
			{
				this.CurChange  = arrChanges[0];
				this.CurElement = this.CurChange.GetElement();

				if (!this.CurElement || !this.Changes[this.CurElement.GetId()])
				{
					this.CurChange  = null;
					this.CurElement = null;
				}
			}
		}

		var oChange = this.private_GetPrevChange();
		if (oChange && oChange.IsMove() && !oChange.IsComplexChange())
		{
			oChange        = this.CollectMoveChange(oChange);
			this.CurChange = oChange;
		}

		return oChange;
	};
	CTrackRevisionsManager.prototype.private_GetPrevChange = function()
	{
		var oInitialCurChange  = this.CurChange;
		var oInitialCurElement = this.CurElement;

		var oPrevElement = null;
		if (null !== this.CurChange && null !== this.CurElement)
		{
			var arrChangesArray = this.Changes[this.CurElement.GetId()];
			var nChangeIndex    = -1;
			for (var nIndex = 0, nCount = arrChangesArray.length; nIndex < nCount; ++nIndex)
			{
				if (this.CurChange === arrChangesArray[nIndex])
				{
					nChangeIndex = nIndex;
					break;
				}
			}

			if (-1 !== nChangeIndex && nChangeIndex > 0)
			{
				this.CurChange = arrChangesArray[nChangeIndex - 1];
				return this.CurChange;
			}

			oPrevElement = this.LogicDocument.GetRevisionsChangeElement(-1, this.CurElement);
		}
		else
		{
			var SearchEngine = this.LogicDocument.private_GetRevisionsChangeElement(-1, null);
			oPrevElement     = SearchEngine.GetFoundedElement();
			if (null !== oPrevElement && oPrevElement === SearchEngine.GetCurrentElement())
			{
				var arrPrevChangesArray = this.Changes[oPrevElement.GetId()];
				if (undefined !== arrPrevChangesArray && arrPrevChangesArray.length > 0)
				{
					if (oPrevElement instanceof Paragraph)
					{
						var ParaContentPos = oPrevElement.Get_ParaContentPos(oPrevElement.IsSelectionUse(), true);
						for (var ChangeIndex = arrPrevChangesArray.length - 1; ChangeIndex >= 0; ChangeIndex--)
						{
							var ChangeStartPos = arrPrevChangesArray[ChangeIndex].get_StartPos();
							if (ParaContentPos.Compare(ChangeStartPos) >= 0)
							{
								this.CurChange  = arrPrevChangesArray[ChangeIndex];
								this.CurElement = oPrevElement;
								return this.CurChange;
							}
						}
					}
					else if (oPrevElement instanceof CTable && oPrevElement.IsCellSelection())
					{
						var arrSelectedCells = oPrevElement.GetSelectionArray();
						if (arrSelectedCells.length > 0)
						{
							var nTableRow = arrSelectedCells[0].Row;
							for (var nChangeIndex = arrPrevChangesArray.length - 1; nChangeIndex >= 0; --nChangeIndex)
							{
								var nStartRow = arrPrevChangesArray[nChangeIndex].get_StartPos();
								if (nTableRow >= nStartRow)
								{
									this.CurChange  = arrPrevChangesArray[nChangeIndex];
									this.CurElement = oPrevElement;
									return this.CurChange;
								}
							}
						}
					}

					oPrevElement = this.LogicDocument.GetRevisionsChangeElement(-1, oPrevElement);
				}
			}
		}

		if (null !== oPrevElement)
		{
			var arrPrevChangesArray = this.Changes[oPrevElement.GetId()];
			if (undefined !== arrPrevChangesArray && arrPrevChangesArray.length > 0)
			{
				this.CurChange  = arrPrevChangesArray[arrPrevChangesArray.length - 1];
				this.CurElement = oPrevElement;
				return this.CurChange;
			}
		}

		if (null !== oInitialCurChange && null !== oInitialCurElement)
		{
			this.CurChange  = oInitialCurChange;
			this.CurElement = oInitialCurElement;
			return oInitialCurChange;
		}

		this.CurChange  = null;
		this.CurElement = null;
		return null;
	};
	/**
	 * Проверяем есть ли непримененные изменения в документе
	 * @returns {boolean}
	 */
	CTrackRevisionsManager.prototype.Have_Changes = function()
	{
		var oTableId = this.LogicDocument ? this.LogicDocument.GetTableId() : null;

		for (var sElementId in this.Changes)
		{
			var oElement = oTableId ? oTableId.Get_ById(sElementId) : null;

			if (!oElement || !oElement.IsUseInDocument || !oElement.IsUseInDocument())
				continue;

			if (this.Changes[sElementId].length > 0)
				return true;
		}

		return false;
	};
	/**
	 * Проверяем есть ли изменения, сделанные другими пользователями
	 * @returns {boolean}
	 */
	CTrackRevisionsManager.prototype.HaveOtherUsersChanges = function()
	{
		var sUserId = this.LogicDocument.GetUserId(false);
		for (var sParaId in this.Changes)
		{
			var oParagraph = AscCommon.g_oTableId.Get_ById(sParaId);
			if (!oParagraph || !oParagraph.IsUseInDocument())
				continue;

			for (var nIndex = 0, nCount = this.Changes[sParaId].length; nIndex < nCount; ++nIndex)
			{
				var oChange = this.Changes[sParaId][nIndex];
				if (oChange.get_UserId() !== sUserId)
					return true;
			}
		}

		return false;
	};
	CTrackRevisionsManager.prototype.ClearCurrentChange = function()
	{
		this.CurChange  = null;
		this.CurElement = null;
	};
	CTrackRevisionsManager.prototype.SetCurrentChange = function(oCurChange)
	{
		if (oCurChange)
		{
			this.CurChange  = oCurChange;
			this.CurElement = oCurChange.GetElement();
		}
	};
	CTrackRevisionsManager.prototype.GetCurrentChangeElement = function()
	{
		return this.CurElement;
	};
	CTrackRevisionsManager.prototype.GetCurrentChange = function()
	{
		return this.CurChange;
	};
	CTrackRevisionsManager.prototype.InitSelectedChanges = function()
	{
		var oEditorApi = this.LogicDocument.GetApi();
		if (!oEditorApi)
			return;

		oEditorApi.sync_BeginCatchRevisionsChanges();
		oEditorApi.sync_EndCatchRevisionsChanges();
	};
	CTrackRevisionsManager.prototype.ClearSelectedChanges = function()
	{
		if (this.SelectedChanges.length > 0)
		{
			var oEditorApi = this.LogicDocument.GetApi();
			if (!oEditorApi)
				return;

			oEditorApi.sync_BeginCatchRevisionsChanges();
			oEditorApi.sync_EndCatchRevisionsChanges();
		}

		this.SelectedChanges = [];
	};
	/**
	 * Добавляем изменение, видимое в текущей позиции
	 * @param oChange
	 */
	CTrackRevisionsManager.prototype.AddSelectedChange = function(oChange)
	{
		if (this.CurChange)
			return;

		if (oChange && c_oAscRevisionsChangeType.MoveMark === oChange.get_Type())
			return;

		if (oChange.IsMove() && !oChange.IsComplexChange())
			oChange = this.CollectMoveChange(oChange);

		for (var nIndex = 0, nCount = this.SelectedChanges.length; nIndex < nCount; ++nIndex)
		{
			var oVisChange = this.SelectedChanges[nIndex];
			if (oVisChange.IsComplexChange() && !oChange.IsComplexChange())
			{
				var arrSimpleChanges = oVisChange.GetSimpleChanges();
				for (var nSimpleIndex = 0, nSimpleCount = arrSimpleChanges.length; nSimpleIndex < nSimpleCount; ++nSimpleIndex)
				{
					if (arrSimpleChanges[nSimpleIndex] === oChange)
						return;
				}
			}
			else if (!oVisChange.IsComplexChange() && oChange.IsComplexChange())
			{
				var arrSimpleChanges = oChange.GetSimpleChanges();
				for (var nSimpleIndex = 0, nSimpleCount = arrSimpleChanges.length; nSimpleIndex < nSimpleCount; ++nSimpleIndex)
				{
					if (arrSimpleChanges[nSimpleIndex] === oVisChange)
					{
						this.SelectedChanges.splice(nIndex, 1);
						nCount--;
						nIndex--;
						break;
					}
				}
			}
			else if (oVisChange.IsComplexChange() && oChange.IsComplexChange())
			{
				var arrVisSC    = oVisChange.GetSimpleChanges();
				var arrChangeSC = oChange.GetSimpleChanges();

				var isEqual = false;
				if (arrVisSC.length === arrChangeSC.length)
				{
					isEqual = true;
					for (var nSimpleIndex = 0, nSimplesCount = arrVisSC.length; nSimpleIndex < nSimplesCount; ++nSimpleIndex)
					{
						if (arrVisSC[nSimpleIndex] !== arrChangeSC[nSimpleIndex])
						{
							isEqual = false;
							break;
						}
					}
				}

				if (isEqual)
					return;
			}
			else if (oChange === oVisChange)
			{
				return;
			}
		}

		this.SelectedChanges.push(oChange);
	};
	CTrackRevisionsManager.prototype.GetSelectedChanges = function()
	{
		return this.SelectedChanges;
	};
	CTrackRevisionsManager.prototype.BeginCollectChanges = function(bSaveCurrentChange)
	{
		if (!this.IsAllChecked())
			return;

		this.PrevSelectedChanges = this.SelectedChanges;
		this.SelectedChanges     = [];

		if (true !== bSaveCurrentChange)
		{
			this.ClearCurrentChange();
		}
		else if (this.CurElement && this.CurChange)
		{
			var oSelectionBounds = this.CurElement.GetSelectionBounds();

			var oBounds = oSelectionBounds.Direction > 0 ? oSelectionBounds.Start : oSelectionBounds.End;

			if (oBounds)
			{
				var X = this.LogicDocument.Get_PageLimits(oBounds.Page).XLimit;
				this.CurChange.put_InternalPos(X, oBounds.Y, oBounds.Page);
				this.SelectedChanges.push(this.CurChange);
			}
		}
	};
	CTrackRevisionsManager.prototype.EndCollectChanges = function()
	{
		if (!this.IsAllChecked())
			return;

		var oEditor = this.LogicDocument.GetApi();
		if (this.CurChange)
			this.SelectedChanges = [this.CurChange];

		var isPositionChanged = false;
		var isArrayChanged    = false;
		var isShowChanges     = this.CurChange || !this.LogicDocument.IsTextSelectionUse();

		var nChangesCount = this.SelectedChanges.length;
		if (this.PrevSelectedChanges.length !== nChangesCount || this.PrevShowChanges !== isShowChanges)
		{
			isArrayChanged = true;
		}
		else if (0 !== nChangesCount)
		{
			for (var nChangeIndex = 0; nChangeIndex < nChangesCount; ++nChangeIndex)
			{
				if (this.SelectedChanges[nChangeIndex] !== this.PrevSelectedChanges[nChangeIndex])
				{
					isArrayChanged = true;
					break;
				}
				else if (this.SelectedChanges[nChangeIndex].IsPositionChanged())
				{
					isPositionChanged = true;
				}
			}
		}

		if (isArrayChanged)
		{
			oEditor.sync_BeginCatchRevisionsChanges();

			if (nChangesCount > 0)
			{
				var oPos = this.private_GetSelectedChangesXY();
				for (var ChangeIndex = 0; ChangeIndex < nChangesCount; ChangeIndex++)
				{
					var Change = this.SelectedChanges[ChangeIndex];
					Change.put_XY(oPos.X, oPos.Y);
					oEditor.sync_AddRevisionsChange(Change);
				}
			}
			oEditor.sync_EndCatchRevisionsChanges(isShowChanges);
		}
		else if (isPositionChanged)
		{
			this.UpdateSelectedChangesPosition(oEditor);
		}

		this.PrevShowChanges = isShowChanges;
	};
	CTrackRevisionsManager.prototype.UpdateSelectedChangesPosition = function(oEditor)
	{
		if (this.SelectedChanges.length > 0)
		{
			var oPos = this.private_GetSelectedChangesXY();
			oEditor.sync_UpdateRevisionsChangesPosition(oPos.X, oPos.Y);
		}
	};
	CTrackRevisionsManager.prototype.private_GetSelectedChangesXY = function()
	{
		if (!this.SelectedChanges.length)
			return {X : 0, Y : 0};
		
		let change = this.SelectedChanges[0];
		
		let element = change.GetElement();
		if (!element)
			return {X : 0, Y : 0};
		
		let x    = change.GetInternalPosX();
		let y    = change.GetInternalPosY();
		let page = change.GetInternalPosPageNum();

		let transform = element.Get_ParentTextTransform();
		if (transform)
			y = transform.TransformPointY(x, y);
		
		let drawingCoords = this.LogicDocument.getDrawingDocument().ConvertCoordsToCursorWR(x, y, page);
		return {X : drawingCoords.X, Y : drawingCoords.Y};
	};
	CTrackRevisionsManager.prototype.Get_AllChangesLogicDocuments = function()
	{
		this.CompleteTrackChanges();
		var LogicDocuments = {};

		for (var ParaId in this.Changes)
		{
			var Para = AscCommon.g_oTableId.Get_ById(ParaId);
			if (Para && Para.Get_Parent())
			{
				LogicDocuments[Para.Get_Parent().Get_Id()] = true;
			}
		}

		return LogicDocuments;
	};
	CTrackRevisionsManager.prototype.GetChangeRelatedParagraphs = function(oChange, bAccept)
	{
		var oRelatedParas = {};

		if (oChange.IsComplexChange())
		{
			var arrSimpleChanges = oChange.GetSimpleChanges();
			for (var nIndex = 0, nCount = arrSimpleChanges.length; nIndex < nCount; ++nIndex)
			{
				this.private_GetChangeRelatedParagraphs(arrSimpleChanges[nIndex], bAccept, oRelatedParas);
			}
		}
		else
		{
			this.private_GetChangeRelatedParagraphs(oChange, bAccept, oRelatedParas);
		}

		return this.private_ConvertParagraphsObjectToArray(oRelatedParas);
	};
	CTrackRevisionsManager.prototype.private_GetChangeRelatedParagraphs = function(oChange, bAccept, oRelatedParas)
	{
		if (oChange)
		{
			var nType    = oChange.GetType();
			var oElement = oChange.GetElement();
			if (oElement && oElement.IsUseInDocument())
			{
				oRelatedParas[oElement.GetId()] = true;
				if ((c_oAscRevisionsChangeType.ParaAdd === nType && true !== bAccept) || (c_oAscRevisionsChangeType.ParaRem === nType && true === bAccept))
				{
					var oLogicDocument = oElement.GetParent();
					var nParaIndex     = oElement.GetIndex();

					if (oLogicDocument && -1 !== nParaIndex)
					{
						if (nParaIndex < oLogicDocument.GetElementsCount() - 1)
						{
							var oNextElement = oLogicDocument.GetElement(nParaIndex + 1);
							if (oNextElement && oNextElement.IsParagraph())
								oRelatedParas[oNextElement.GetId()] = true;
						}
					}
				}
			}
		}
	};
	CTrackRevisionsManager.prototype.private_ConvertParagraphsObjectToArray = function(ParagraphsObject)
	{
		var ParagraphsArray = [];
		for (var ParaId in ParagraphsObject)
		{
			var Para = AscCommon.g_oTableId.Get_ById(ParaId);
			if (null !== Para)
			{
				ParagraphsArray.push(Para);
			}
		}
		return ParagraphsArray;
	};
	CTrackRevisionsManager.prototype.Get_AllChangesRelatedParagraphs = function(bAccept)
	{
		var RelatedParas = {};
		for (var ParaId in this.Changes)
		{
			for (var ChangeIndex = 0, ChangesCount = this.Changes[ParaId].length; ChangeIndex < ChangesCount; ++ChangeIndex)
			{
				var Change = this.Changes[ParaId][ChangeIndex];
				this.private_GetChangeRelatedParagraphs(Change, bAccept, RelatedParas);
			}
		}
		return this.private_ConvertParagraphsObjectToArray(RelatedParas);
	};
	CTrackRevisionsManager.prototype.Get_AllChangesRelatedParagraphsBySelectedParagraphs = function(SelectedParagraphs, bAccept)
	{
		var RelatedParas = {};
		for (var ParaIndex = 0, ParasCount = SelectedParagraphs.length; ParaIndex < ParasCount; ++ParaIndex)
		{
			var Para = SelectedParagraphs[ParaIndex];
			var ParaId = Para.Get_Id();
			if (this.Changes[ParaId] && this.Changes[ParaId].length > 0)
			{
				RelatedParas[ParaId] = true;
				if (true === Para.Selection_CheckParaEnd())
				{
					var CheckNext = false;
					for (var ChangeIndex = 0, ChangesCount = this.Changes[ParaId].length; ChangeIndex < ChangesCount; ++ChangeIndex)
					{
						var ChangeType = this.Changes[ParaId][ChangeIndex].get_Type();
						if ((c_oAscRevisionsChangeType.ParaAdd === ChangeType && true !== bAccept) || (c_oAscRevisionsChangeType.ParaRem === ChangeType && true === bAccept))
						{
							CheckNext = true;
							break;
						}
					}

					if (true === CheckNext)
					{
						var NextElement = Para.Get_DocumentNext();
						if (null !== NextElement && type_Paragraph === NextElement.Get_Type())
						{
							RelatedParas[NextElement.Get_Id()] = true;
						}
					}
				}
			}
		}
		return this.private_ConvertParagraphsObjectToArray(RelatedParas);
	};
	CTrackRevisionsManager.prototype.Get_AllChanges = function()
	{
		this.CompleteTrackChanges();
		return this.Changes;
	};
	CTrackRevisionsManager.prototype.IsAllChecked = function()
	{
		return (!this.CheckArray.length && !haveElementToCheck());
	};
	/**
	 * Завершаем проверку всех элементов на наличие рецензирования
	 */
	CTrackRevisionsManager.prototype.CompleteTrackChanges = function()
	{
		while (!this.IsAllChecked())
			this.ContinueTrackRevisions();
	};
	/**
	 * Завершаем проверку рецензирования для заданных элементов
	 * @param arrElements
	 * @returns {boolean}
	 */
	CTrackRevisionsManager.prototype.CompleteTrackChangesForElements = function(arrElements)
	{
		var isChecked = false;
		for (var nIndex = 0, nCount = arrElements.length; nIndex < nCount; ++nIndex)
		{
			if (this.private_TrackChangesForSingleElement(arrElements[nIndex].GetId()))
				isChecked = true;
		}

		return isChecked;
	};
	CTrackRevisionsManager.prototype.private_TrackChangesForSingleElement = function(Id)
	{
		if (!this.CheckMap[Id])
			return false;

		let oElement = this.CheckMap[Id];

		delete this.CheckMap[Id];

		if (!oElement.IsUseInDocument())
			return false;

		let isHaveChanges = !!this.Changes[Id];

		this.private_RemoveChangeObject(Id);
		oElement.CheckRevisionsChanges(this);

		return !(!isHaveChanges && !this.Changes[Id]);
	};
	/**
	 * При чтении файла обновляем Id перетаскиваний в рецензировании
	 * @param sMoveId
	 */
	CTrackRevisionsManager.prototype.UpdateMoveId = function(sMoveId)
	{
		if (0 === sMoveId.indexOf("move"))
		{
			var nId = parseInt(sMoveId.substring(4));
			if (!isNaN(nId))
				this.MoveId = Math.max(this.MoveId, nId);
		}
	};
	/**
	 * Возвращаем новый идентификатор перемещений
	 * @returns {string}
	 */
	CTrackRevisionsManager.prototype.GetNewMoveId = function()
	{
		this.MoveId++;
		return "move" + this.MoveId;
	};
	CTrackRevisionsManager.prototype.RegisterMoveMark = function(oMark)
	{
		if (this.LogicDocument && this.LogicDocument.PrintSelection)
			return;

		if (!oMark)
			return;

		var sMarkId = oMark.GetMarkId();
		var isFrom  = oMark.IsFrom();
		var isStart = oMark.IsStart();

		this.UpdateMoveId(sMarkId);

		if (!this.MoveMarks[sMarkId])
		{
			this.MoveMarks[sMarkId] = {

				From : {
					Start : null,
					End   : null
				},

				To : {
					Start : null,
					End   : null
				}
			};
		}

		if (isFrom)
		{
			if (isStart)
				this.MoveMarks[sMarkId].From.Start = oMark;
			else
				this.MoveMarks[sMarkId].From.End = oMark;
		}
		else
		{
			if (isStart)
				this.MoveMarks[sMarkId].To.Start = oMark;
			else
				this.MoveMarks[sMarkId].To.End = oMark;
		}
	};
	CTrackRevisionsManager.prototype.UnregisterMoveMark = function(oMark)
	{
		if (this.LogicDocument && this.LogicDocument.PrintSelection)
			return;

		if (!oMark)
			return;

		var sMarkId = oMark.GetMarkId();
		delete this.MoveMarks[sMarkId];

		// TODO: Возможно тут нужно проделать дополнительные действия
	};
	CTrackRevisionsManager.prototype.private_CheckChangeObject = function(sId)
	{
		var oElement = AscCommon.g_oTableId.Get_ById(sId);
		if (!oElement)
			return false;

		if (!this.Changes[sId])
			this.Changes[sId] = [];

		var nDeletePosition = -1;
		for (var nIndex = 0, nCount = this.ChangesOutline.length; nIndex < nCount; ++nIndex)
		{
			if (this.ChangesOutline[nIndex].GetId() === sId)
			{
				nDeletePosition = nIndex;
				break;
			}
		}

		var oDocPos = oElement.GetDocumentPositionFromObject();
		if (!oDocPos)
			return;

		var nAddPosition = -1;
		for (var nIndex = 0, nCount = this.ChangesOutline.length; nIndex < nCount; ++nIndex)
		{
			var oTempDocPos = this.ChangesOutline[nIndex].GetDocumentPositionFromObject();

			if (this.private_CompareDocumentPositions(oDocPos, oTempDocPos) < 0)
			{
				nAddPosition = nIndex;
				break;
			}
		}

		if (-1 === nAddPosition)
			nAddPosition = this.ChangesOutline.length;

		if (nAddPosition === nDeletePosition || (-1 !== nAddPosition && -1 !== nDeletePosition && nDeletePosition === nAddPosition - 1))
			return true;

		if (-1 !== nDeletePosition)
		{
			this.ChangesOutline.splice(nDeletePosition, 1);

			if (nAddPosition > nDeletePosition)
				nAddPosition--;
		}

		this.ChangesOutline.splice(nAddPosition, 0, oElement);

		return true;
	};
	CTrackRevisionsManager.prototype.private_CompareDocumentPositions = function(oDocPos1, oDocPos2)
	{
		if (oDocPos1.Class !== oDocPos2.Class)
		{
			// TODO: Здесь нужно доработать сравнение позиций, когда они из разных частей документа
			if (oDocPos1.Class instanceof CDocument)
				return -1;
			else if (oDocPos1.Class instanceof CDocument)
				return 1;
			else
				return 1;
		}

		for (var nIndex = 0, nCount = oDocPos1.length; nIndex < nCount; ++nIndex)
		{
			if (oDocPos2.length <= nIndex)
				return 1;

			if (oDocPos1[nIndex].Position < oDocPos2[nIndex].Position)
				return -1;
			else if (oDocPos1[nIndex].Position > oDocPos2[nIndex].Position)
				return 1;
		}

		return 0;
	};
	CTrackRevisionsManager.prototype.private_RemoveChangeObject = function(sId)
	{
		if (this.Changes[sId])
			delete this.Changes[sId];

		for (var nIndex = 0, nCount = this.ChangesOutline.length; nIndex < nCount; ++nIndex)
		{
			if (this.ChangesOutline[nIndex].GetId() === sId)
			{
				this.ChangesOutline.splice(nIndex, 1);
				return;
			}
		}
	};
	/**
	 * Собираем изменение связанное с переносом
	 * @param {CRevisionsChange} oChange
	 * @returns {CRevisionsChange}
	 */
	CTrackRevisionsManager.prototype.CollectMoveChange = function(oChange)
	{
		var isFrom = c_oAscRevisionsChangeType.TextRem === oChange.GetType() || c_oAscRevisionsChangeType.ParaRem === oChange.GetType() || (c_oAscRevisionsChangeType.MoveMark === oChange.GetType() && oChange.GetValue().IsFrom());

		var nStartIndex  = -1;
		var oStartChange = null;

		var oElement = oChange.GetElement();
		if (!oElement)
			return oChange;

		var nDeep = 0;
		var nSearchIndex = -1;
		for (var nIndex = 0, nCount = this.ChangesOutline.length; nIndex < nCount; ++nIndex)
		{
			if (this.ChangesOutline[nIndex] === oElement)
			{
				nSearchIndex = nIndex;
				break;
			}
		}

		if (-1 === nSearchIndex)
			return oChange;

		var isStart = false;

		for (var nIndex = nSearchIndex; nIndex >= 0; --nIndex)
		{
			var arrCurChanges = this.Changes[this.ChangesOutline[nIndex].GetId()];

			if (!arrCurChanges)
			{
				isStart = true;
				continue;
			}

			for (var nChangeIndex = arrCurChanges.length - 1; nChangeIndex >= 0; --nChangeIndex)
			{
				var oCurChange = arrCurChanges[nChangeIndex];
				if (!isStart)
				{
					if (oCurChange === oChange)
						isStart = true;
				}

				if (isStart)
				{
					var nCurChangeType = oCurChange.GetType();
					if (nCurChangeType === c_oAscRevisionsChangeType.MoveMark)
					{
						var oMoveMark = oCurChange.GetValue();
						if ((isFrom && oMoveMark.IsFrom()) || (!isFrom && !oMoveMark.IsFrom()))
						{
							if (oMoveMark.IsStart())
							{
								if (nDeep > 0)
								{
									nDeep--;
								}
								else if (nDeep === 0)
								{
									nStartIndex  = nIndex;
									oStartChange = oCurChange;
									break;
								}
							}
							else if (oCurChange !== oChange)
							{
								nDeep++;
							}
						}
					}
				}
			}

			if (oStartChange)
				break;

			isStart = true;
		}

		if (!oStartChange || -1 === nStartIndex)
			return oChange;

		var sValue     = "";
		var arrChanges = [oStartChange];

		isStart = false;
		nDeep   = 0;
		var isEnd = false;
		for (var nIndex = nStartIndex, nCount = this.ChangesOutline.length; nIndex < nCount; ++nIndex)
		{
			var arrCurChanges = this.Changes[this.ChangesOutline[nIndex].GetId()];
			for (var nChangeIndex = 0, nChangesCount = arrCurChanges.length; nChangeIndex < nChangesCount; ++nChangeIndex)
			{
				var oCurChange = arrCurChanges[nChangeIndex];
				if (!isStart)
				{
					if (oCurChange === oStartChange)
						isStart = true;
				}
				else
				{
					var nCurChangeType = oCurChange.GetType();
					if (isFrom)
					{
						if (c_oAscRevisionsChangeType.TextRem === nCurChangeType || c_oAscRevisionsChangeType.ParaRem === nCurChangeType)
						{
							if (0 === nDeep)
							{
								sValue += c_oAscRevisionsChangeType.TextRem === nCurChangeType ? oCurChange.GetValue() : "\n";
								arrChanges.push(oCurChange);
							}
						}
						else if (c_oAscRevisionsChangeType.MoveMark === nCurChangeType && oCurChange.GetValue().IsFrom())
						{
							if (oCurChange.GetValue().IsStart())
							{
								nDeep++;
							}
							else if (nDeep > 0)
							{
								nDeep--;
							}
							else
							{
								arrChanges.push(oCurChange);
								isEnd = true;
								break;
							}
						}
					}
					else
					{
						if (c_oAscRevisionsChangeType.TextAdd === nCurChangeType || c_oAscRevisionsChangeType.ParaAdd === nCurChangeType)
						{
							if (0 === nDeep)
							{
								sValue += c_oAscRevisionsChangeType.TextAdd === nCurChangeType ? oCurChange.GetValue() : "\n";
								arrChanges.push(oCurChange);
							}
						}
						else if (c_oAscRevisionsChangeType.MoveMark === nCurChangeType && !oCurChange.GetValue().IsFrom())
						{
							if (oCurChange.GetValue().IsStart())
							{
								nDeep++;
							}
							else if (nDeep > 0)
							{
								nDeep--;
							}
							else
							{
								arrChanges.push(oCurChange);
								isEnd = true;
								break;
							}
						}
					}
				}
			}

			if (!isStart)
				return oChange;

			if (isEnd)
				break;
		}

		var sMoveId = oStartChange.GetValue().GetMarkId();
		var isDown  = null;

		for (var nIndex = 0, nCount = this.ChangesOutline.length; nIndex < nCount; ++nIndex)
		{
			var arrCurChanges = this.Changes[this.ChangesOutline[nIndex].GetId()];
			if (!arrCurChanges)
				continue;

			for (var nChangeIndex = 0, nChangesCount = arrCurChanges.length; nChangeIndex < nChangesCount; ++nChangeIndex)
			{
				var oCurChange = arrCurChanges[nChangeIndex];
				if (c_oAscRevisionsChangeType.MoveMark === oCurChange.GetType())
				{
					var oMark = oCurChange.GetValue();
					if (sMoveId === oMark.GetMarkId())
					{
						isDown = !!oMark.IsFrom();
						break;
					}
				}
			}

			if (null !== isDown)
				break;
		}

		if (!isEnd || null === isDown)
			return oChange;

		var oMoveChange = new CRevisionsChange();
		oMoveChange.SetType(isFrom ? c_oAscRevisionsChangeType.TextRem : c_oAscRevisionsChangeType.TextAdd);
		oMoveChange.SetValue(sValue);
		oMoveChange.SetElement(oStartChange.GetElement());
		oMoveChange.SetUserId(oStartChange.GetUserId());
		oMoveChange.SetUserName(oStartChange.GetUserName());
		oMoveChange.SetDateTime(oStartChange.GetDateTime());
		oMoveChange.SetMoveType(isFrom ? Asc.c_oAscRevisionsMove.MoveFrom : Asc.c_oAscRevisionsMove.MoveTo);
		oMoveChange.SetSimpleChanges(arrChanges);
		oMoveChange.SetMoveId(sMoveId);
		oMoveChange.SetMovedDown(isDown);
		oMoveChange.SetXY(oChange.GetX(), oChange.GetY());
		oMoveChange.SetInternalPos(oChange.GetInternalPosX(), oChange.GetInternalPosY(), oChange.GetInternalPosPageNum());
		return oMoveChange;
	};
	/**
	 * Получаем массив всех изменений связанных с заданным переносом
	 * @param {string} sMoveId
	 * @returns {CRevisionsChange[]}
	 */
	CTrackRevisionsManager.prototype.GetAllMoveChanges = function(sMoveId)
	{
		var oStartFromChange = null;
		var oStartToChange   = null;

		for (var sElementId in this.Changes)
		{
			var arrElementChanges = this.Changes[sElementId];
			for (var nChangeIndex = 0, nChangesCount = arrElementChanges.length; nChangeIndex < nChangesCount; ++nChangeIndex)
			{
				var oCurChange = arrElementChanges[nChangeIndex];
				if (c_oAscRevisionsChangeType.MoveMark === oCurChange.GetType() && sMoveId === oCurChange.GetValue().GetMarkId() && oCurChange.GetValue().IsStart())
				{
					if (oCurChange.GetValue().IsFrom())
						oStartFromChange = oCurChange;
					else
						oStartToChange = oCurChange;
				}
			}

			if (oStartFromChange && oStartToChange)
				break;
		}

		if (!oStartFromChange || !oStartToChange)
			return {From : [], To : []};

		return {
			From : this.CollectMoveChange(oStartFromChange).GetSimpleChanges(),
			To   : this.CollectMoveChange(oStartToChange).GetSimpleChanges()
		};
	};
	/**
	 * Начинаем процесс обработки(принятия или отклонения) перетаскивания текста
	 * @param sMoveId {string} идентификатор перетаскивания
	 * @param sUserId {string} идентификатор пользователя
	 * @returns {CTrackRevisionsMoveProcessEngine}
	 */
	CTrackRevisionsManager.prototype.StartProcessReviewMove = function(sMoveId, sUserId)
	{
		return (this.ProcessMove = new CTrackRevisionsMoveProcessEngine(sMoveId, sUserId));
	};
	/**
	 * Завершаем процесс обработки перетаскивания текста
	 */
	CTrackRevisionsManager.prototype.EndProcessReviewMove = function()
	{
		// TODO: Здесь нужно сделать обработку MovesToDelete

		this.ProcessMove = null;
	};
	/**
	 * Проверям, запущен ли процесс обрабокти перетаскивания текста
	 * @returns {?CTrackRevisionsMoveProcessEngine}
	 */
	CTrackRevisionsManager.prototype.GetProcessTrackMove = function()
	{
		return this.ProcessMove;
	};
	/**
	 * Получаем метки переноса
	 * @param sMarkId
	 */
	CTrackRevisionsManager.prototype.GetMoveMarks = function(sMarkId)
	{
		return this.MoveMarks[sMarkId];
	};
	/**
	 * Получаем элементарное изменение связанное с заданным переносом, относящееся к метке переноса
	 * @param {string} sMoveId
	 * @param {boolean} isFrom
	 * @param {boolean} isStart
	 */
	CTrackRevisionsManager.prototype.GetMoveMarkChange = function(sMoveId, isFrom, isStart)
	{
		this.CompleteTrackChanges();

		var oMoveChanges = this.GetAllMoveChanges(sMoveId);
		var arrChanges   = isFrom ? oMoveChanges.From : oMoveChanges.To;

		for (var nIndex = 0, nCount = arrChanges.length; nIndex < nCount; ++nIndex)
		{
			var oChange = arrChanges[nIndex];
			if (Asc.c_oAscRevisionsChangeType.MoveMark === oChange.GetType())
			{
				var oMark = oChange.GetValue();
				if (oMark.IsFrom() === isFrom && oMark.IsStart() === isStart)
				{
					return oChange;
				}
			}
		}

		return null;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CTrackRevisionsManager = CTrackRevisionsManager;
	window['AscWord'].checkElementInRevision = checkElementInRevision;

})(window);
