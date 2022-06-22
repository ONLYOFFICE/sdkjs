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
var c_oAscSectionBreakType    = Asc.c_oAscSectionBreakType;
CTable.prototype.Recalculate_Page = function(PageIndex)
{
	this.SetIsRecalculated(true);

	if (0 === PageIndex)
	{
		// TODO: Внутри функции private_RecalculateBorders происходит персчет метрик каждой ячейки, это надо бы
		//       вынести в отдельную функцию. Из-за этого функцию private_RecalculateHeader приходится запускать позже.

		this.private_RecalculateGrid();
		this.private_RecalculateBorders();
		this.private_RecalculateHeader();
	}

	this.private_RecalculatePageXY(PageIndex);

	if (true !== this.private_RecalculateCheckPageColumnBreak(PageIndex))
		return recalcresult_NextPage | recalcresultflags_Column;

	this.private_RecalculatePositionX(PageIndex);

	var Result = this.private_RecalculatePage(PageIndex);
	if (Result & recalcresult_CurPage)
		return Result;

	this.private_RecalculatePositionY(PageIndex);

	if (Result & recalcresult_NextElement)
		this.RecalcInfo.Reset(false);

	if (Result & recalcresult_NextElement && window['AscCommon'].g_specialPasteHelper && window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph === this.GetId())
		window['AscCommon'].g_specialPasteHelper.SpecialPasteButtonById_Show();

	return Result;
};
CTable.prototype.Recalculate_SkipPage = function(PageIndex)
{
	if (0 === PageIndex)
	{
		this.StartFromNewPage();
	}
	else
	{
		var PrevPage = this.Pages[PageIndex - 1];

		var LastRow      = Math.max(PrevPage.FirstRow, PrevPage.LastRow); // На случай, если предыдущая страница тоже пустая
		var NewPage      = new CTablePage(PrevPage.X, PrevPage.Y, PrevPage.XLimit, PrevPage.YLimit, LastRow, PrevPage.MaxTopBorder);
		NewPage.FirstRow = LastRow;
		NewPage.LastRow  = LastRow - 1;

		this.Pages[PageIndex] = NewPage;
	}
};
CTable.prototype.Recalculate_Grid = function()
{
	this.private_RecalculateGrid();
};
CTable.prototype.SaveRecalculateObject = function()
{
	var RecalcObj = new CTableRecalculateObject();
	RecalcObj.Save(this);
	return RecalcObj;
};
CTable.prototype.LoadRecalculateObject = function(RecalcObj)
{
	RecalcObj.Load(this);
};
CTable.prototype.PrepareRecalculateObject = function()
{
	this.TableSumGrid  = [];
	this.TableGridCalc = [];

	this.TableRowsBottom = [];
	this.RowsInfo        = [];

	this.HeaderInfo = {
		Count     : 0,
		H         : 0,
		PageIndex : 0,
		Pages     : []
	};

	this.Pages = [];

	this.MaxTopBorder = [];
	this.MaxBotBorder = [];
	this.MaxBotMargin = [];

	this.RecalcInfo.Reset(true);

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Index].PrepareRecalculateObject();
	}
};
CTable.prototype.StartFromNewPage = function()
{
	this.Pages.length     = 1;
	this.Pages[0]         = new CTablePage(0, 0, 0, 0, 0, 0);
	this.Pages[0].LastRow = -1;

	this.HeaderInfo.Pages[0]      = {};
	this.HeaderInfo.Pages[0].Draw = false;

	this.RowsInfo[0] = new CTableRowsInfo();
	this.RowsInfo[0].Init();

	// Обнуляем таблицу суммарных высот ячеек
	for (var Index = -1; Index < this.Content.length; Index++)
	{
		this.TableRowsBottom[Index]    = [];
		this.TableRowsBottom[Index][0] = 0;
	}

	this.Pages[0].MaxBotBorder = 0;
	this.Pages[0].BotBorders   = [];

	if (this.Content.length > 0)
	{
		var CellsCount = this.Content[0].Get_CellsCount();
		for (var CurCell = 0; CurCell < CellsCount; CurCell++)
		{
			var Cell = this.Content[0].Get_Cell(CurCell);
			Cell.Content.StartFromNewPage();
			Cell.PagesCount = 2;
		}
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Приватные функции связанные с расчетом таблицы
//----------------------------------------------------------------------------------------------------------------------
CTable.prototype.private_RecalculateCheckPageColumnBreak = function(CurPage)
{
    if (true !== this.Is_Inline()) // Случай Flow разбирается в Document.js
        return true;

    if (!this.LogicDocument || this.Parent !== this.LogicDocument)
        return true;

    var isPageBreakOnPrevLine   = false;
    var isColumnBreakOnPrevLine = false;

    var PrevElement = this.Get_DocumentPrev();

	while (PrevElement && (PrevElement instanceof CBlockLevelSdt))
		PrevElement = PrevElement.GetLastElement();

    if (null !== PrevElement && type_Paragraph === PrevElement.Get_Type() && true === PrevElement.Is_Empty() && undefined !== PrevElement.Get_SectionPr())
	{
		var PrevSectPr = PrevElement.Get_SectionPr();
		var CurSectPr  = this.LogicDocument.SectionsInfo.Get_SectPr(this.Index).SectPr;
		if (c_oAscSectionBreakType.Continuous === CurSectPr.Get_Type() && true === CurSectPr.Compare_PageSize(PrevSectPr))
			PrevElement = PrevElement.Get_DocumentPrev();
	}

    if ((0 === CurPage || true === this.Check_EmptyPages(CurPage - 1)) && null !== PrevElement && type_Paragraph === PrevElement.Get_Type())
    {
        var bNeedPageBreak = true;
        if (undefined !== PrevElement.Get_SectionPr())
        {
            var PrevSectPr = PrevElement.Get_SectionPr();
            var CurSectPr  = this.LogicDocument.SectionsInfo.Get_SectPr(this.Index).SectPr;
            if (c_oAscSectionBreakType.Continuous !== CurSectPr.Get_Type() || true !== CurSectPr.Compare_PageSize(PrevSectPr))
                bNeedPageBreak = false;
        }

        if (true === bNeedPageBreak)
        {
            var EndLine = PrevElement.Pages[PrevElement.Pages.length - 1].EndLine;
            if (-1 !== EndLine && PrevElement.Lines[EndLine].Info & paralineinfo_BreakRealPage)
                isPageBreakOnPrevLine = true;
        }
    }

    // ColumnBreak для случая CurPage > 0 не разбираем здесь, т.к. он срабатывает автоматически
    if (0 === CurPage && null !== PrevElement && type_Paragraph === PrevElement.Get_Type())
    {
        var EndLine = PrevElement.Pages[PrevElement.Pages.length - 1].EndLine;
        if (-1 !== EndLine && !(PrevElement.Lines[EndLine].Info & paralineinfo_BreakRealPage) && PrevElement.Lines[EndLine].Info & paralineinfo_BreakPage)
            isColumnBreakOnPrevLine = true;
    }

    if ((true === isPageBreakOnPrevLine && (0 !== this.private_GetColumnIndex(CurPage) || (0 === CurPage && null !== PrevElement)))
        || (true === isColumnBreakOnPrevLine && 0 === CurPage))
    {
        this.private_RecalculateSkipPage(CurPage);
        return false;
    }

    return true;
};
CTable.prototype.private_RecalculateGrid = function()
{
	if (this.GetRowsCount() <= 0)
		return;

	//---------------------------------------------------------------------------
	// 1 часть пересчета ширины таблицы : Рассчитываем фиксированную ширину
	//---------------------------------------------------------------------------
	var TablePr = this.Get_CompiledPr(false).TablePr;

	var PctWidth = this.CalculatedPctWidth;
	var MinWidth = this.CalculatedMinWidth;
	var nTableW  = this.CalculatedTableW;

	if (null === this.CalculatedX || null === this.CalculatedXLimit || Math.abs(this.X - this.CalculatedX) > 0.001 || Math.abs(this.XLimit - this.CalculatedXLimit))
	{
		this.CalculatedX      = this.X;
		this.CalculatedXLimit = this.XLimit;

		this.RecalcInfo.TableGrid    = true;
		this.RecalcInfo.TableBorders = true;

		// Запускаем пересчет границ, потому что там пересчитываются метрики ячеек, которые зависят от X, XLimit
	}

	if (this.RecalcInfo.TableGrid)
	{
		var arrSumGrid = [];
		var nTempSum   = 0;
		arrSumGrid[-1] = 0;
		for (var nIndex = 0, nCount = this.TableGrid.length; nIndex < nCount; ++nIndex)
		{
			nTempSum += this.TableGrid[nIndex];
			arrSumGrid[nIndex] = nTempSum;
		}

		PctWidth = this.private_RecalculatePercentWidth();
		MinWidth = this.private_GetTableMinWidth();

		nTableW = 0;
		if (tblwidth_Auto === TablePr.TableW.Type)
		{
			nTableW = 0;
		}
		else if (tblwidth_Nil === TablePr.TableW.Type)
		{
			nTableW = MinWidth;
		}
		else
		{
			if (tblwidth_Pct === TablePr.TableW.Type)
				nTableW = PctWidth * TablePr.TableW.W / 100;
			else
				nTableW = TablePr.TableW.W;

			if (0.001 > nTableW)
				nTableW = 0;
			else if (nTableW < MinWidth)
				nTableW = MinWidth;
		}

		var nCurGridCol = 0;
		for (var nCurRow = 0, nRowsCount = this.GetRowsCount(); nCurRow < nRowsCount; ++nCurRow)
		{
			var oRow = this.GetRow(nCurRow);
			oRow.SetIndex(nCurRow);

			// Смотрим на ширину пропущенных колонок сетки в начале строки
			var oBeforeInfo = oRow.GetBefore();
			nCurGridCol     = oBeforeInfo.Grid;

			if (nCurGridCol > 0 && arrSumGrid[nCurGridCol - 1] < oBeforeInfo.W)
			{
				var nTempDiff = oBeforeInfo.W - arrSumGrid[nCurGridCol - 1];
				for (var nTempIndex = nCurGridCol - 1; nTempIndex < arrSumGrid.length; ++nTempIndex)
				{
					arrSumGrid[nTempIndex] += nTempDiff;
				}
			}

			var nCellSpacing = oRow.GetCellSpacing();
			for (var nCurCell = 0, nCellsCount = oRow.GetCellsCount(); nCurCell < nCellsCount; ++nCurCell)
			{
				var oCell = oRow.GetCell(nCurCell);
				oCell.SetIndex(nCurCell);

				var oCellW    = oCell.GetW();
				var nGridSpan = oCell.GetGridSpan();

				if (vmerge_Restart !== oCell.GetVMerge())
				{
					nCurGridCol += nGridSpan;
					continue;
				}

				if (nCurGridCol + nGridSpan - 1 > arrSumGrid.length)
				{
					for (var nAddIndex = arrSumGrid.length; nAddIndex <= nCurGridCol + nGridSpan - 1; ++nAddIndex)
					{
						// Добавляем столбик шириной в 2 см
						arrSumGrid[nAddIndex] = arrSumGrid[nAddIndex - 1] + 20;
					}
				}

				if (tblwidth_Auto !== oCellW.Type && tblwidth_Nil !== oCellW.Type)
				{
					var nCellWidth = 0;
					if (tblwidth_Pct === oCellW.Type)
						nCellWidth = PctWidth * oCellW.W / 100;
					else
						nCellWidth = oCellW.W;

					if (null !== nCellSpacing)
					{
						if (0 === nCurCell)
							nCellWidth += nCellSpacing / 2;

						nCellWidth += nCellSpacing;

						if (nCellsCount - 1 === nCurCell)
							nCellWidth += nCellSpacing / 2;
					}

					if (nCellWidth + arrSumGrid[nCurGridCol - 1] > arrSumGrid[nCurGridCol + nGridSpan - 1])
					{
						var nTempDiff = nCellWidth + arrSumGrid[nCurGridCol - 1] - arrSumGrid[nCurGridCol + nGridSpan - 1];
						for (var nTempIndex = nCurGridCol + nGridSpan - 1; nTempIndex < arrSumGrid.length; ++nTempIndex)
						{
							arrSumGrid[nTempIndex] += nTempDiff;
						}
					}
				}

				nCurGridCol += nGridSpan;
			}

			// Смотрим на ширину пропущенных колонок сетки в конце строки
			var oAfterInfo = oRow.GetAfter();
			if (nCurGridCol + oAfterInfo.Grid - 1 > arrSumGrid.length)
			{
				for (var nAddIndex = arrSumGrid.length; nAddIndex <= nCurGridCol + oAfterInfo.Grid - 1; ++nAddIndex)
				{
					// Добавляем столбик шириной в 2 см
					arrSumGrid[nAddIndex] = arrSumGrid[nAddIndex - 1] + 20;
				}
			}

			if (arrSumGrid[nCurGridCol + oAfterInfo.Grid - 1] < oAfterInfo.W + arrSumGrid[nCurGridCol - 1])
			{
				var nTempDiff = oAfterInfo.W + arrSumGrid[nCurGridCol - 1] - arrSumGrid[nCurGridCol + oAfterInfo.Grid - 1];
				for (var nTempIndex = nCurGridCol + oAfterInfo.Grid - 1; nTempIndex < arrSumGrid.length; ++nTempIndex)
				{
					arrSumGrid[nTempIndex] += nTempDiff;
				}
			}
		}

		// TODO: разобраться с минимальной шириной таблицы и ячеек

		// Задана общая ширина таблицы и последняя ячейка вышла за пределы
		// данной ширины. Уменьшаем все столбцы сетки пропорционально, чтобы
		// суммарная ширина стала равной заданной ширине таблицы.
		if (nTableW > 0 && Math.abs(arrSumGrid[arrSumGrid.length - 1] - nTableW) > 0.01)
			arrSumGrid = this.Internal_ScaleTableWidth(arrSumGrid, nTableW);
		else if (MinWidth > arrSumGrid[arrSumGrid.length - 1])
			arrSumGrid = this.Internal_ScaleTableWidth(arrSumGrid, arrSumGrid[arrSumGrid.length - 1]);

		// По массиву SumGrid восстанавливаем ширины самих колонок
		this.TableGridCalc    = [];
		this.TableGridCalc[0] = arrSumGrid[0];
		for (var nIndex = 1, nCount = arrSumGrid.length; nIndex < nCount; ++nIndex)
		{
			this.TableGridCalc[nIndex] = arrSumGrid[nIndex] - arrSumGrid[nIndex - 1];
		}

		this.TableSumGrid = arrSumGrid;

		this.CalculatedPctWidth = PctWidth;
		this.CalculatedMinWidth = MinWidth;
		this.CalculatedTableW   = nTableW;

		this.RecalcInfo.TableGrid = false;
	}

	var TopTable = this.Parent ? this.Parent.IsInTable(true) : null;
	if ((null === TopTable && tbllayout_AutoFit === TablePr.TableLayout) || (null != TopTable && tbllayout_AutoFit === TopTable.Get_CompiledPr(false).TablePr.TableLayout))
	{
		//---------------------------------------------------------------------------
		// 2 часть пересчета ширины таблицы : Рассчитываем ширину по содержимому
		//---------------------------------------------------------------------------

		// 1. Расчитаем минимальные и максимальные значения ширины для всех колонок, а также минимальное значение
		//    отступов и предпочитаемую ширину колонок
		var arrMinMargin  = [],
			arrMinContent = [],
			arrMaxContent = [],
			arrPreferred  = [], // 0 - ориентируемся на содержимое ячеек, > 0 - ориентируемся только на ширину ячеек записанную в свойствах
			arrMinNoPref  = []; // минимальное значение контента, только без учета предпочитаемы ширин

		var nColsCount = this.TableGridCalc.length;
		for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
		{
			arrMinMargin[nCurCol]  = 0;
			arrMinContent[nCurCol] = 0;
			arrMaxContent[nCurCol] = 0;
			arrPreferred[nCurCol]  = 0;
			arrMinNoPref[nCurCol]  = 0;
		}
		this.private_RecalculateGridMinContent(PctWidth, arrMinMargin, arrMinContent, arrMaxContent, arrPreferred, arrMinNoPref);

		// 2. Проследим, чтобы значения arrMinContent и arrMaxContent не превосходили значение 55,87см(так работает Word)
		for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
		{
			if (arrMaxContent[nCurCol] < arrMinContent[nCurCol])
				arrMaxContent[nCurCol] = arrMinContent[nCurCol];

			if (arrMinNoPref[nCurCol] > 558.7)
				arrMinNoPref[nCurCol] = 558.7;

			if (arrMinContent[nCurCol] > 558.7)
				arrMinContent[nCurCol] = 558.7;

			if (arrMaxContent[nCurCol] > 558.7)
				arrMaxContent[nCurCol] = 558.7;
		}

		// 3. Рассчитаем максимально допустимую ширину под всю таблицу

		var oPageFields;

		if (!this.Parent)
		{
			oPageFields  = {X : 0, Y : 0, XLimit : 0, YLimit : 0};
		}
		else
		{
			// Случай, когда таблица лежит внутри CBlockLevelSdt
			if (this.Parent instanceof CDocumentContent && this.LogicDocument && this.Parent.IsBlockLevelSdtContent() && this.Parent.GetTopDocumentContent() === this.LogicDocument && !this.Parent.IsTableCellContent())
			{
				var nTopIndex = -1;
				var arrPos    = this.GetDocumentPositionFromObject();
				if (arrPos.length > 0)
					nTopIndex = arrPos[0].Position;

				if (-1 !== nTopIndex)
					oPageFields = this.LogicDocument.Get_ColumnFields(nTopIndex, this.Get_AbsoluteColumn(this.PageNum), this.GetAbsolutePage(this.PageNum));
			}

			if (!oPageFields)
				oPageFields = this.Parent.Get_ColumnFields ? this.Parent.Get_ColumnFields(this.Get_Index(), this.Get_AbsoluteColumn(this.PageNum), this.GetAbsolutePage(this.PageNum)) : this.Parent.Get_PageFields(this.private_GetRelativePageIndex(this.PageNum), this.Parent.IsHdrFtr());
		}

		var oFramePr = this.GetFramePr();
		if (oFramePr && undefined !== oFramePr.GetW())
		{
			oPageFields.X      = 0;
			oPageFields.XLimit = oFramePr.GetW();
		}
		else if (this.LogicDocument && this.LogicDocument.IsDocumentEditor() && this.IsInline())
		{
			var _X      = oPageFields.X;
			var _XLimit = oPageFields.XLimit;

			var arrRanges = this.Parent.CheckRange(_X, this.Y, _XLimit, this.Y + 0.001, this.Y, this.Y + 0.001, _X, _XLimit, this.private_GetRelativePageIndex(0));
			if (arrRanges.length > 0)
			{
				for (var nRangeIndex = 0, nRangesCount = arrRanges.length; nRangeIndex < nRangesCount; ++nRangeIndex)
				{
					if (arrRanges[nRangeIndex].X0 < oPageFields.X + 3.2 && arrRanges[nRangeIndex].X1 > _X)
						_X = arrRanges[nRangeIndex].X1 + 0.001;

					if (arrRanges[nRangeIndex].X1 > oPageFields.XLimit - 3.2 && arrRanges[nRangeIndex].X0 < _XLimit)
						_XLimit = arrRanges[nRangeIndex].X0 - 0.001;
				}
			}

			oPageFields.X      = _X;
			oPageFields.XLimit = _XLimit;
		}

		var nMaxTableW = oPageFields.XLimit - oPageFields.X - TablePr.TableInd - this.GetTableOffsetCorrection() + this.GetRightTableOffsetCorrection();

		// 4. Рассчитаем желаемую ширину таблицы таблицы
		var arrMaxOverMin  = [],
			nSumMaxOverMin = 0,
			nSumMin        = 0,
			nSumMinMargin  = 0,
			nSumMax        = 0;

		for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
		{
			arrMaxOverMin[nCurCol] = Math.max(0, arrMaxContent[nCurCol] - arrMinContent[nCurCol]);

			nSumMin += arrMinContent[nCurCol];
			nSumMaxOverMin += arrMaxOverMin[nCurCol];
			nSumMinMargin += arrMinMargin[nCurCol];
			nSumMax += arrMinContent[nCurCol] + arrMaxOverMin[nCurCol];
		}

		var nMaxTableWOrigin = nMaxTableW;
		if ((TablePr.TableW.IsMM() || TablePr.TableW.IsPercent()) && nMaxTableW < nTableW)
			nMaxTableW = nTableW;

		if (nSumMin < nMaxTableW)
		{
			if (nSumMax <= nMaxTableW || nSumMaxOverMin < 0.001)
			{
				for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
				{
					this.TableGridCalc[nCurCol] = Math.max(arrMinContent[nCurCol], arrMaxContent[nCurCol]);
				}
			}
			else
			{
				for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
				{
					this.TableGridCalc[nCurCol] = arrMinContent[nCurCol] + (nMaxTableW - nSumMin) * arrMaxOverMin[nCurCol] / nSumMaxOverMin;
				}
			}

			// Если у таблицы задана ширина, тогда ориентируемся по ширине, а если нет, тогда ориентируемся по
			// максимальным значениям.
			if (TablePr.TableW.IsMM() || TablePr.TableW.IsPercent())
			{
				if (nSumMin < 0.001 && nSumMax < 0.001)
				{
					for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
					{
						this.TableGridCalc[nCurCol] = nTableW / nColsCount;
					}
				}
				else if (nSumMin >= nTableW)
				{
					var nSumMinContent         = 0,
						nSumPrefOverMinContent = 0,
						arrPrefOverMinContent  = [];

					for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
					{
						arrPrefOverMinContent[nCurCol] = Math.max(0, arrMinContent[nCurCol] - arrMinNoPref[nCurCol]);
						nSumPrefOverMinContent += arrPrefOverMinContent[nCurCol];
						nSumMinContent += arrMinNoPref[nCurCol];
					}

					if (nSumMinContent >= nTableW)
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							this.TableGridCalc[nCurCol] = arrMinNoPref[nCurCol];
						}
					}
					else
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							this.TableGridCalc[nCurCol] = arrMinNoPref[nCurCol] + arrPrefOverMinContent[nCurCol] * (nTableW - nSumMinContent) / nSumPrefOverMinContent;
						}
					}
				}
				else
				{
					// nSumMin < nMaxTableW, значит у нас есть свободное пространство для распределения
					// Колонки, у которых задана предпочитаемая ширина не трогаем, с одним исключением, когда
					// такие колонки все, в этом случае растягиваем их пропрорционально

					var nSumMaxOverMin      = 0,
						isAllPreferred      = true,
						arrNoPrefMaxOverMin = [],
						nSumMaxOverMin      = 0,
						nSumNonPrefMax      = 0,
						nSumPrefMin         = 0,
						nSumNoPrefMin       = 0;

					for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
					{
						if (arrPreferred[nCurCol] < 0.001)
						{
							isAllPreferred = false;

							var nMinW = arrMinContent[nCurCol];
							var nMaxW = Math.max(arrMaxContent[nCurCol], nMinW);

							arrNoPrefMaxOverMin[nCurCol] = nMaxW - nMinW;
							nSumMaxOverMin += nMaxW - nMinW;
							nSumNonPrefMax += nMaxW;
							nSumNoPrefMin += nMinW;
						}
						else
						{
							nSumPrefMin += arrMinContent[nCurCol];
							arrNoPrefMaxOverMin[nCurCol] = 0;
						}
					}


					// Если данное условие выполняется, значит у нас все ячейки с предпочитаемыми значениями, тогда
					// мы растягиваем все ячейки равномерно. Если не выполняется, значит есть ячейки, в которых
					// предпочитаемое значение не задано, и тогда растягиваем только такие ячейки.
					if (isAllPreferred)
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							this.TableGridCalc[nCurCol] = arrMinContent[nCurCol] * nTableW / nSumMin;
						}
					}
					else if (nSumNonPrefMax < nMaxTableW - nSumMin && nSumNonPrefMax > 0.001)
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							if (arrPreferred[nCurCol] < 0.001)
								this.TableGridCalc[nCurCol] = arrMaxContent[nCurCol] * (nTableW - nSumPrefMin) / nSumNonPrefMax;
							else
								this.TableGridCalc[nCurCol] = arrMinContent[nCurCol];
						}
					}
					else if (nSumMaxOverMin > 0.001)
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							this.TableGridCalc[nCurCol] = arrMinContent[nCurCol] + (nTableW - nSumMin) * arrNoPrefMaxOverMin[nCurCol] / nSumMaxOverMin;
						}
					}
					else if (nSumNoPrefMin > 0.001)
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							if (arrPreferred[nCurCol] < 0.001)
								this.TableGridCalc[nCurCol] = arrMinContent[nCurCol] * (nTableW - nSumPrefMin) / nSumNoPrefMin;
							else
								this.TableGridCalc[nCurCol] = arrMinContent[nCurCol];
						}
					}
					else
					{
						// Такого быть не должно
					}
				}
			}
		}
		else
		{
			// Если в таблице сделать все ячейки нулевой ширины (для контента), и все равно она получается шире
			// максимальной допустимой ширины, тогда выставляем ширины всех колоно по минимальному значению
			// маргинов и оставляем так как есть
			if (nMaxTableW - nSumMinMargin < 0.001)
			{
				for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
				{
					this.TableGridCalc[nCurCol] = arrMinMargin[nCurCol];
				}
			}
			else
			{
				var nSumMinContent         = 0,
					nSumPrefOverMinContent = 0,
					arrPrefOverMinContent  = [];

				for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
				{
					arrPrefOverMinContent[nCurCol] = Math.max(0, arrMinContent[nCurCol] - arrMinNoPref[nCurCol]);
					nSumPrefOverMinContent += arrPrefOverMinContent[nCurCol];
					nSumMinContent += arrMinNoPref[nCurCol];
				}


				if (nSumMinContent > nMaxTableW)
				{
					if (nTableW > nMaxTableWOrigin)
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							this.TableGridCalc[nCurCol] = arrMinNoPref[nCurCol];
						}
					}
					else
					{
						var nSumMinNoPrefOverMinMargin = 0,
							arrMinNoPrefOverMinMargin  = [];

						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							arrMinNoPrefOverMinMargin[nCurCol] = Math.max(arrMinNoPref[nCurCol] - arrMinMargin[nCurCol], 0);
							nSumMinNoPrefOverMinMargin += arrMinNoPrefOverMinMargin[nCurCol];
						}

						if (nSumMinNoPrefOverMinMargin > 0.001)
						{
							for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
							{
								this.TableGridCalc[nCurCol] = arrMinMargin[nCurCol] + arrMinNoPrefOverMinMargin[nCurCol] * (nMaxTableW - nSumMinMargin) / nSumMinNoPrefOverMinMargin;
							}
						}
						else
						{
							// Такого быть не должно, т.к. мы в ветке nMaxTableW > nSumMinMargin + 0.001 && nSumMinContent > nMaxTableW
						}
					}
				}
				else
				{
					// Если попали в эту ветку, то это может означать только одно, что у нас есть ячейки с заданной
					// шириной, превышающей минимальное значения ширины контента

					if ((TablePr.TableW.IsMM() || TablePr.TableW.IsPercent()) && nTableW < nSumMinContent)
					{
						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							this.TableGridCalc[nCurCol] = arrMinNoPref[nCurCol];
						}
					}
					else
					{
						var _nTableW = (TablePr.TableW.IsMM() || TablePr.TableW.IsPercent() ? nTableW : nMaxTableWOrigin);

						var nSumPrefOverMinContent = 0,
							arrPrefOverMinContent  = [];

						for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
						{
							arrPrefOverMinContent[nCurCol] = Math.max(arrMinContent[nCurCol] - arrMinNoPref[nCurCol], 0);
							nSumPrefOverMinContent += arrPrefOverMinContent[nCurCol];
						}

						if (nSumPrefOverMinContent > 0.001)
						{
							for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
							{
								this.TableGridCalc[nCurCol] = arrMinNoPref[nCurCol] + arrPrefOverMinContent[nCurCol] * (_nTableW - nSumMinContent) / nSumPrefOverMinContent;
							}
						}
						else
						{
							// Такого быть не должно
							for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
							{
								this.TableGridCalc[nCurCol] = arrMinNoPref[nCurCol];
							}
						}
					}
				}
			}
		}

		this.TableSumGrid[-1] = 0;
		for (var nCurCol = 0; nCurCol < nColsCount; ++nCurCol)
			this.TableSumGrid[nCurCol] = this.TableSumGrid[nCurCol - 1] + this.TableGridCalc[nCurCol];
	}
};
CTable.prototype.private_RecalculateGridMinContent = function(nPctWidth, arrMinMargins, arrMinContent, arrMaxContent, arrPreferred, arrMinNoPreferred)
{
	// Сначала мы высчитываем минимальный и максимальный контент для всех ячеек с GridSpan=1, ячейки
	// у которых GridSpan > 1 заносим в массив arrMergedColumns.
	// После обработки всех ячеек с GridSpan=1, мы пробегаемся по массиву arrMergedPreferred, это ячейки, у которых
	// GridSpan > 1 и у которых задана предпочтительная ширина. Далее на основе этих ячееки высчитываем предпочтительную
	// ширину колонок. Причем делаем это начиная с ячеек с наименьшим GridSpan.
	// Далее пробегаемся по arrMergedCells и если какой-нибудь отрезок из колонок (StartGrid, StartGrid+GridSpan)
	// по ширине занимает место больше, чем эти колонки по отдельности, тогда мы добавочное место распределяем среди
	// колонок  с сохранением соотношения их ширины, как частный случай, если у нас встретилась колонка нулевой ширины,
	// например, у которой вообще не было целой ячейки внутри колонки, тогда она так и останется нулевой ширины
	// после распределения. Причем тут важно, что мы пробегаемся по массиву подряд, а не на основе GridSpan, как
	// у arrMergedPreferred, поэтому результат для одной и той же таблицы, с переставленными двумя строками может
	// быть разным.

	var arrMergedColumns   = [];
	var arrMergedPreferred = [];
	var nMergedPrefCount   = 0;
	for (var nCurRow = 0, nRowsCount = this.GetRowsCount(); nCurRow < nRowsCount; ++nCurRow)
	{
		var oRow = this.GetRow(nCurRow);

		var nSpacing  = oRow.GetCellSpacing();
		var nSpacingW = (null !== nSpacing ? nSpacing : 0);

		var nCurGridCol = 0;

		var oBeforeInfo = oRow.GetBefore();
		var nGridBefore = oBeforeInfo.Grid;
		var nBeforeW    = oBeforeInfo.W.GetCalculatedValue(nPctWidth);
		if (nBeforeW > 0 && nGridBefore > 0)
		{
			if (1 === nGridBefore)
			{
				if (arrMinContent[nCurGridCol] < nBeforeW)
					arrMinContent[nCurGridCol] = nBeforeW;

				if (0 === arrPreferred[nCurGridCol] || arrPreferred[nCurGridCol] < nBeforeW)
					arrPreferred[nCurGridCol] = nBeforeW;

				if (arrPreferred[nCurGridCol])
					arrMaxContent[nCurGridCol] = arrPreferred[nCurGridCol];
				else if (arrMaxContent[nCurGridCol] < nBeforeW)
					arrMaxContent[nCurGridCol] = nBeforeW;
			}
			else
			{
				arrMergedColumns.push({
					Start     : nCurGridCol,
					Min       : nBeforeW,
					Max       : nBeforeW,
					Preferred : nBeforeW,
					Margins   : 0,
					Grid      : nGridBefore,
					MinNoPref : 0
				});

				if (!arrMergedPreferred[nGridBefore])
					arrMergedPreferred[nGridBefore] = [];

				arrMergedPreferred[nGridBefore].push({
					Start : nCurGridCol,
					W     : nBeforeW
				});

				nMergedPrefCount++;
			}
		}

		nCurGridCol = nGridBefore;

		for (var nCurCell = 0, nCellsCount = oRow.GetCellsCount(); nCurCell < nCellsCount; ++nCurCell)
		{
			var oCell       = oRow.GetCell(nCurCell);
			var oCellMinMax = oCell.RecalculateMinMaxContentWidth(false, nPctWidth);
			var nGridSpan   = oCell.GetGridSpan();
			var nCellMin    = oCellMinMax.Min;
			var nCellMax    = oCellMinMax.Max;
			var oCellW      = oCell.GetW();
			var oMargins    = oCell.GetMargins();
			var nPreferred  = oCellW.GetCalculatedValue(nPctWidth);

			var nSpacingAdd = (0 === nCurCell || nCellsCount - 1 === nCurCell) ? 3 / 2 * nSpacingW : nSpacingW;
			var nMarginsMin = nSpacingAdd + oMargins.Left.W + oMargins.Right.W;

			nCellMin += nSpacingAdd;
			nCellMax += nSpacingAdd;

			if (1 === nGridSpan)
			{
				if (arrMinContent[nCurGridCol] < nCellMin)
					arrMinContent[nCurGridCol] = nCellMin;

				if (nPreferred && arrPreferred[nCurGridCol] < nPreferred)
					arrPreferred[nCurGridCol] = nPreferred;

				if (arrPreferred[nCurGridCol])
					arrMaxContent[nCurGridCol] = arrPreferred[nCurGridCol];
				else if (arrMaxContent[nCurGridCol] < nCellMax)
					arrMaxContent[nCurGridCol] = nCellMax;

				if (arrMinMargins[nCurGridCol] < nMarginsMin)
					arrMinMargins[nCurGridCol] = nMarginsMin;

				if (arrMinNoPreferred[nCurGridCol] < oCellMinMax.ContentMin + nSpacingAdd)
					arrMinNoPreferred[nCurGridCol] = oCellMinMax.ContentMin + nSpacingAdd;
			}
			else if (nGridSpan > 1)
			{
				arrMergedColumns.push({
					Start     : nCurGridCol,
					Min       : nCellMin,
					Max       : nCellMax,
					Preferred : nPreferred,
					Margins   : nMarginsMin,
					Grid      : nGridSpan,
					MinNoPref : oCellMinMax.ContentMin + nSpacingAdd
				});

				if (nPreferred > 0.001)
				{
					if (!arrMergedPreferred[nGridSpan])
						arrMergedPreferred[nGridSpan] = [];

					arrMergedPreferred[nGridSpan].push({
						Start : nCurGridCol,
						W     : nPreferred
					});

					nMergedPrefCount++;
				}
			}

			nCurGridCol += nGridSpan;
		}

		var oAfterInfo = oRow.GetAfter();
		var nGridAfter = oAfterInfo.Grid;
		var nAfterW    = oAfterInfo.W.GetCalculatedValue(nPctWidth);
		if (nAfterW > 0 && nGridAfter > 0)
		{
			if (1 === nGridAfter)
			{
				if (arrMinContent[nCurGridCol] < nAfterW)
					arrMinContent[nCurGridCol] = nAfterW;

				if (0 === arrPreferred[nCurGridCol] || arrPreferred[nCurGridCol] < nAfterW)
					arrPreferred[nCurGridCol] = nAfterW;

				if (arrPreferred[nCurGridCol])
					arrMaxContent[nCurGridCol] = arrPreferred[nCurGridCol];
				else if (arrMaxContent[nCurGridCol] < nAfterW)
					arrMaxContent[nCurGridCol] = nAfterW;
			}
			else
			{
				arrMergedColumns.push({
					Start     : nCurGridCol,
					Min       : nAfterW,
					Max       : nAfterW,
					Preferred : nAfterW,
					Margins   : 0,
					Grid      : nGridAfter,
					MinNoPref : 0
				});

				if (!arrMergedPreferred[nGridAfter])
					arrMergedPreferred[nGridAfter] = [];

				arrMergedPreferred[nGridAfter].push({
					Start : nCurGridCol,
					W     : nAfterW
				});

				nMergedPrefCount++;
			}
		}
	}
	
	while (nMergedPrefCount > 0)
	{
		var nPrevCount = nMergedPrefCount;

		for (var nGridSpan = 2, nMaxGridSpan = arrMergedPreferred.length; nGridSpan < nMaxGridSpan; ++nGridSpan)
		{
			if (!arrMergedPreferred[nGridSpan])
				continue;

			for (var nIndex = arrMergedPreferred[nGridSpan].length - 1; nIndex >= 0; --nIndex)
			{
				var nStart     = arrMergedPreferred[nGridSpan][nIndex].Start;
				var nPreferred = arrMergedPreferred[nGridSpan][nIndex].W;

				var nPreferredSum = 0;

				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan - 1; ++nCurSpan)
				{
					if (arrPreferred[nCurSpan] > 0 && -1 !== nPreferredSum)
						nPreferredSum += arrPreferred[nCurSpan];
					else
						nPreferredSum = -1;
				}

				if (-1 !== nPreferredSum && nPreferred > nPreferredSum && arrPreferred[nStart + nGridSpan - 1] < nPreferred - nPreferredSum)
				{
					arrPreferred[nStart + nGridSpan - 1] = nPreferred - nPreferredSum;

					if (arrMinContent[nStart + nGridSpan - 1] < arrPreferred[nStart + nGridSpan - 1])
						arrMinContent[nStart + nGridSpan - 1] = arrPreferred[nStart + nGridSpan - 1];

					arrMergedPreferred[nGridSpan].splice(nIndex, 1);
					nMergedPrefCount--;
				}
			}
		}

		if (nPrevCount <= nMergedPrefCount)
			break;
	}

	for (var nIndex = 0, nCount = arrMergedColumns.length; nIndex < nCount; ++nIndex)
	{
		var nStart     = arrMergedColumns[nIndex].Start;
		var nMin       = arrMergedColumns[nIndex].Min;
		var nMax       = arrMergedColumns[nIndex].Max;
		var nMargins   = arrMergedColumns[nIndex].Margins;
		var nPreferred = arrMergedColumns[nIndex].Preferred;
		var nGridSpan  = arrMergedColumns[nIndex].Grid;
		var nMinNoPref = arrMergedColumns[nIndex].MinNoPref;

		var nSumSpanMin       = 0;
		var nSumSpanMax       = 0;
		var nSumMargin        = 0;
		var nSumSpanMinNoPref = 0;

		for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
		{
			nSumSpanMin += arrMinContent[nCurSpan];
			nSumSpanMax += arrMaxContent[nCurSpan];
			nSumMargin += arrMinMargins[nCurSpan];
			nSumSpanMinNoPref += arrMinNoPreferred[nCurSpan];
		}

		if (nMargins > nSumMargin)
		{
			if (nSumMargin < 0.001)
			{
				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
				{
					arrMinMargins[nCurSpan] = nMargins / nGridSpan;
				}
			}
			else
			{
				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
				{
					arrMinMargins[nCurSpan] = arrMinMargins[nCurSpan] * nMargins / nSumMargin;
				}
			}
		}

		if (nMin > nSumSpanMin)
		{
			if (nSumSpanMin < 0.001)
			{
				// Такого не должно быть, но на всякий случай в данной ситуации делаем равномерное распределение
				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
				{
					arrMinContent[nCurSpan] = nMin / nGridSpan;
				}
			}
			else
			{
				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
				{
					arrMinContent[nCurSpan] = arrMinContent[nCurSpan] * nMin / nSumSpanMin;
				}
			}
		}

		if (nMinNoPref > nSumSpanMinNoPref)
		{
			if (nSumSpanMin < 0.001)
			{
				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
				{
					arrMinNoPreferred[nCurSpan] = nMinNoPref / nGridSpan;
				}
			}
			else
			{
				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
				{
					arrMinNoPreferred[nCurSpan] = arrMinContent[nCurSpan] * nMinNoPref / nSumSpanMin;
				}
			}
		}

		if (nMax > nSumSpanMax)
		{
			// Если у нас в объединении несколько колонок, тогда явно записанная ширина ячейки не
			// перекрывает ширину ни одной из колонок, она всего лишь учавствует в определении
			// максимальной ширины.

			// Высчитываем сумму максимумов по колонкам, у которых стоит флаг false
			var nSumSpanMax2 = 0;
			for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
			{
				if (arrPreferred[nCurSpan] < 0.001)
					nSumSpanMax2 += arrMaxContent[nCurSpan];
			}

			// Если nSumSpanMax2=0, тогда во всех колонках выставлен флаг=true и мы должны учитывать
			// уже имеющиеся там значения, а текущее значение будет проигнорировано
			if (nSumSpanMax2 > 0.001)
			{
				for (var nCurSpan = nStart; nCurSpan < nStart + nGridSpan; ++nCurSpan)
				{
					if (arrPreferred[nCurSpan] < 0.001)
					{
						arrMaxContent[nCurSpan] = arrMaxContent[nCurSpan] * (nSumSpanMax - nMax + nSumSpanMax2) / nSumSpanMax2;

						if (nPreferred > 0.001)
							arrPreferred[nCurSpan] = arrMaxContent[nCurSpan];
					}
				}
			}
		}
	}
};
CTable.prototype.private_RecalculateBorders = function()
{
    if ( true != this.RecalcInfo.TableBorders )
        return;

    // Обнуляем таблицу суммарных высот ячеек
    for ( var Index = -1; Index < this.Content.length; Index++ )
    {
        this.TableRowsBottom[Index] = [];
        this.TableRowsBottom[Index][0] = 0;
    }

    // Изначально найдем верхние границы и (если нужно) нижние границы
    // для каждой ячейки.
    var MaxTopBorder = [];
    var MaxBotBorder = [];
    var MaxBotMargin = [];

    for ( var Index = 0; Index < this.Content.length; Index++ )
    {
        MaxBotBorder[Index] = 0;
        MaxTopBorder[Index] = 0;
        MaxBotMargin[Index] = 0;
    }

    var TablePr = this.Get_CompiledPr(false).TablePr;
    var TableBorders = this.Get_Borders();

    var nRowsCountInHeader = this.GetRowsCountInHeader();
    var oHeaderLastRow     = nRowsCountInHeader ? this.GetRow(nRowsCountInHeader - 1) : null;

    for ( var CurRow = 0; CurRow < this.Content.length; CurRow++ )
    {
        var Row         = this.Content[CurRow];
        var CellsCount  = Row.Get_CellsCount();
        var CellSpacing = Row.Get_CellSpacing();

        var BeforeInfo = Row.Get_Before();
        var AfterInfo  = Row.Get_After();
        var CurGridCol = BeforeInfo.GridBefore;

        // Нам нужно пробежаться по текущей строке и выяснить максимальное значение ширины верхней
        // границы и ширины нижней границы, заодно рассчитаем вид границы у каждой ячейки,
        // также надо рассчитать максимальное значение нижнего отступа всей строки.

        var bSpacing_Top = false;
        var bSpacing_Bot = false;

        if ( null != CellSpacing )
        {
            bSpacing_Bot = true;
            bSpacing_Top = true;
        }
        else
        {
            if ( 0 != CurRow )
            {
                var PrevCellSpacing = this.Content[CurRow - 1].Get_CellSpacing();
                if ( null != PrevCellSpacing )
                    bSpacing_Top = true;
            }

            if ( this.Content.length - 1 != CurRow )
            {
                var NextCellSpacing = this.Content[CurRow + 1].Get_CellSpacing();
                if ( null != NextCellSpacing )
                    bSpacing_Bot = true;
            }
        }

        Row.Set_SpacingInfo( bSpacing_Top, bSpacing_Bot );

        for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
        {
            var Cell     = Row.Get_Cell( CurCell );
            var GridSpan = Cell.Get_GridSpan();
            var Vmerge   = Cell.GetVMerge();

            Row.Set_CellInfo( CurCell, CurGridCol, 0, 0, 0, 0, 0, 0 );

            // Bug 32418 ячейки, участвующие в вертикальном объединении, все равно участвуют в определении границы
            // строки, поэтому здесь мы их не пропускаем.

            var VMergeCount = this.Internal_GetVertMergeCount( CurRow, CurGridCol, GridSpan );

            var CellMargins = Cell.GetMargins();
            if ( CellMargins.Bottom.W > MaxBotMargin[CurRow + VMergeCount - 1] )
                MaxBotMargin[CurRow + VMergeCount - 1] = CellMargins.Bottom.W;

            var CellBorders = Cell.Get_Borders();
            if ( true === bSpacing_Top )
            {
                if ( border_Single === CellBorders.Top.Value && MaxTopBorder[CurRow] < CellBorders.Top.Size )
                    MaxTopBorder[CurRow] = CellBorders.Top.Size;

				Cell.SetBorderInfoTop([CellBorders.Top]);
				Cell.SetBorderInfoTopHeader([CellBorders.Top]);
            }
            else
            {
                if ( 0 === CurRow )
                {
                    // Сравним границы
                    var Result_Border = this.private_ResolveBordersConflict( TableBorders.Top, CellBorders.Top, true, false );
                    if ( border_Single === Result_Border.Value && MaxTopBorder[CurRow] < Result_Border.Size )
                        MaxTopBorder[CurRow] = Result_Border.Size;

                    var BorderInfo_Top = [];
                    for ( var TempIndex = 0; TempIndex < GridSpan; TempIndex++ )
                        BorderInfo_Top.push( Result_Border );

					Cell.SetBorderInfoTop(BorderInfo_Top);
					Cell.SetBorderInfoTopHeader(BorderInfo_Top);
                }
                else
                {
					var oCellTopInfo       = this.private_RecalculateCellTopBorder(this.GetRow(CurRow - 1), CurRow, CurGridCol, GridSpan, TableBorders, CellBorders);
					var oCellTopHeaderInfo = oHeaderLastRow ? this.private_RecalculateCellTopBorder(oHeaderLastRow, CurRow, CurGridCol, GridSpan, TableBorders, CellBorders) : oCellTopInfo;

					if (MaxTopBorder[CurRow] < oCellTopInfo.Max)
						MaxTopBorder[CurRow] = oCellTopInfo.Max;

                    Cell.SetBorderInfoTop(oCellTopInfo.Info);
                    Cell.SetBorderInfoTopHeader(oCellTopHeaderInfo.Info);
                }
            }

            var CellBordersBottom = CellBorders.Bottom;
            if (VMergeCount > 1)
            {
                // Берем нижнюю границу нижней ячейки вертикального объединения.
                var BottomCell = this.Internal_Get_EndMergedCell(CurRow, CurGridCol, GridSpan);
                if (null !== BottomCell)
                    CellBordersBottom = BottomCell.Get_Borders().Bottom;
            }

            if ( true === bSpacing_Bot )
            {
                Cell.Set_BorderInfo_Bottom( [CellBordersBottom], -1, -1 );

                if ( border_Single === CellBordersBottom.Value && CellBordersBottom.Size > MaxBotBorder[CurRow + VMergeCount - 1] )
                    MaxBotBorder[CurRow + VMergeCount - 1] = CellBordersBottom.Size;
            }
            else
            {
                if ( this.Content.length - 1 === CurRow + VMergeCount - 1 )
                {
                    // Сравним границы
                    var Result_Border = this.private_ResolveBordersConflict( TableBorders.Bottom, CellBordersBottom, true, false );

                    if ( border_Single === Result_Border.Value && Result_Border.Size > MaxBotBorder[CurRow + VMergeCount - 1] )
                        MaxBotBorder[CurRow + VMergeCount - 1] = Result_Border.Size;

                    if ( GridSpan > 0 )
                    {
                        for ( var TempIndex = 0; TempIndex < GridSpan; TempIndex++ )
                            Cell.Set_BorderInfo_Bottom( [ Result_Border ], -1, -1 );
                    }
                    else
                        Cell.Set_BorderInfo_Bottom( [ ], -1, -1 );
                }
                else
                {
                    // Мы должны проверить нижнюю границу ячейки, на предмет того, что со следующей строкой
                    // она может пересекаться по GridBefore и/или GridAfter. Везде в таких местах мы должны
                    // нарисовать нижнюю границу. Пересечение с ячейками нам неинтересено, потому что этот
                    // случай будет учтен при обсчете следующей строки (там будет случай bSpacing_Top = false
                    // и 0 != CurRow )

                    var Next_Row = this.Content[CurRow + VMergeCount];
                    var Next_CellsCount = Next_Row.Get_CellsCount();
                    var Next_BeforeInfo = Next_Row.Get_Before();
                    var Next_AfterInfo  = Next_Row.Get_After();

                    var Border_Bottom_Info = [];

                    // Сначала посмотрим пересечение с GridBefore предыдущей строки
                    var BeforeCount = 0;
                    if ( CurGridCol <= Next_BeforeInfo.GridBefore - 1 )
                    {
                        var Result_Border = this.private_ResolveBordersConflict( TableBorders.Left, CellBordersBottom, true, false );
                        BeforeCount = Math.min( Next_BeforeInfo.GridBefore - CurGridCol, GridSpan );

                        for ( var TempIndex = 0; TempIndex < BeforeCount; TempIndex++ )
                            Border_Bottom_Info.push( Result_Border );
                    }

                    var Next_GridCol = Next_BeforeInfo.GridBefore;
                    for ( var NextCell = 0; NextCell < Next_CellsCount; NextCell++ )
                    {
                        var Next_Cell     = Next_Row.Get_Cell( NextCell );
                        var Next_GridSpan = Next_Cell.Get_GridSpan();
                        Next_GridCol += Next_GridSpan;
                    }

                    // Посмотрим пересечение с GridAfter предыдущей строки
                    var AfterCount = 0;
                    if ( Next_AfterInfo.GridAfter > 0 )
                    {
                        var StartAfterGrid = Next_GridCol;

                        if ( CurGridCol + GridSpan - 1 >= StartAfterGrid )
                        {
                            var Result_Border = this.private_ResolveBordersConflict( TableBorders.Right, CellBordersBottom, true, false );
                            AfterCount = Math.min( CurGridCol + GridSpan - StartAfterGrid, GridSpan );
                            for ( var TempIndex = 0; TempIndex < AfterCount; TempIndex++ )
                                Border_Bottom_Info.push( Result_Border );
                        }
                    }

                    Cell.Set_BorderInfo_Bottom( Border_Bottom_Info, BeforeCount, AfterCount );
                }
            }

            CurGridCol += GridSpan;
        }
    }

    this.MaxTopBorder = MaxTopBorder;
    this.MaxBotBorder = MaxBotBorder;
    this.MaxBotMargin = MaxBotMargin;

    // Также для каждой ячейки обсчитаем ее метрики и левую и правую границы
    for ( var CurRow = 0; CurRow < this.Content.length; CurRow++  )
    {
        var Row         = this.Content[CurRow];
        var CellsCount  = Row.Get_CellsCount();
        var CellSpacing = Row.Get_CellSpacing();

        var BeforeInfo  = Row.Get_Before();
        var AfterInfo   = Row.Get_After();
        var CurGridCol  = BeforeInfo.GridBefore;

        var Row_x_max = 0;
        var Row_x_min = 0;

        for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
        {
            var Cell     = Row.Get_Cell( CurCell );
            var GridSpan = Cell.Get_GridSpan();
            var Vmerge   = Cell.GetVMerge();

            // начальная и конечная точки данного GridSpan'a
            var X_grid_start = this.TableSumGrid[CurGridCol - 1];
            var X_grid_end   = this.TableSumGrid[CurGridCol + GridSpan - 1];

            // границы самой ячейки
            var X_cell_start = X_grid_start;
            var X_cell_end   = X_grid_end;

            if ( null != CellSpacing )
            {

                if ( 0 === CurCell )
                {
                    if ( 0 === BeforeInfo.GridBefore )
                    {
                        if ( border_None === TableBorders.Left.Value || CellSpacing > TableBorders.Left.Size / 2 )
                            X_cell_start += CellSpacing;
                        else
                            X_cell_start += TableBorders.Left.Size / 2;
                    }
                    else
                    {
                        if ( border_None === TableBorders.Left.Value || CellSpacing > TableBorders.Left.Size ) // CellSpacing / 2 > TableBorders.Left.Size / 2
                            X_cell_start += CellSpacing / 2;
                        else
                            X_cell_start += TableBorders.Left.Size / 2;
                    }
                }
                else
                    X_cell_start += CellSpacing / 2;

                if ( CellsCount - 1 === CurCell )
                {
                    if ( 0 === AfterInfo.GridAfter )
                    {
                        if ( border_None === TableBorders.Right.Value || CellSpacing > TableBorders.Right.Size / 2 )
                            X_cell_end -= CellSpacing;
                        else
                            X_cell_end -= TableBorders.Right.Size / 2;
                    }
                    else
                    {
                        if ( border_None === TableBorders.Right.Value || CellSpacing > TableBorders.Right.Size ) // CellSpacing / 2 > TableBorders.Right.Size / 2
                            X_cell_end -= CellSpacing / 2;
                        else
                            X_cell_end -= TableBorders.Right.Size / 2;
                    }
                }
                else
                    X_cell_end -= CellSpacing / 2;
            }

            var CellMar = Cell.GetMargins();

            var VMergeCount = this.Internal_GetVertMergeCount( CurRow, CurGridCol, GridSpan );

            // начальная и конечная точка для содержимого данной ячейки
            var X_content_start = X_cell_start;
            var X_content_end   = X_cell_end;

            // Левая и правая границы ячейки рисуются вовнутрь ячейки, если Spacing != null.
            var CellBorders = Cell.Get_Borders();
            if ( null != CellSpacing )
            {
                X_content_start += CellMar.Left.W;
                X_content_end   -= CellMar.Right.W;

                if ( border_Single === CellBorders.Left.Value )
                    X_content_start += CellBorders.Left.Size;

                if ( border_Single === CellBorders.Right.Value )
                    X_content_end -= CellBorders.Right.Size;
            }
            else
            {
                if ( vmerge_Continue === Vmerge )
                {
                    X_content_start += CellMar.Left.W;
                    X_content_end   -= CellMar.Right.W;
                }
                else
                {
                    // Линии правой и левой границы рисуются ровно по сетке
                    // (середина линии(всмысле толщины линии) совпадает с линией сетки).
                    // Мы должны найти максимальную толщину линии, участвущую в правой/левой
                    // границах. Если данная толщина меньше соответствующего отступа, тогда
                    // она не влияет на расположение содержимого ячейки, в противном случае,
                    // максимальная толщина линии и задает отступ для содержимого.

                    // Поэтому первым шагом определим максимальную толщину правой и левой границ.

                    var Max_r_w = 0;
                    var Max_l_w = 0;

					var Borders_Info = {
						Right : [],
						Left  : [],

						Right_Max : 0,
						Left_Max  : 0
					};

					for (var nTempCurRow = 0; nTempCurRow < VMergeCount; ++nTempCurRow)
					{
						var oTempRow = this.GetRow(CurRow + nTempCurRow);

						var nTempCurCell = this.private_GetCellIndexByStartGridCol(CurRow + nTempCurRow, CurGridCol);
						if (nTempCurCell < 0)
							continue;

						var oTempCell        = oTempRow.GetCell(nTempCurCell);
						var oTempCellBorders = oTempCell.GetBorders();

						// Обработка левой границы
						if (0 === nTempCurCell)
						{
							var oLeftBorder = this.private_ResolveBordersConflict(TableBorders.Left, oTempCellBorders.Left, true, false);
							if (border_Single === oLeftBorder.Value && oLeftBorder.Size > Max_l_w)
								Max_l_w = oLeftBorder.Size;

							Borders_Info.Left.push(oLeftBorder);
						}
						else
						{
							var oLeftBorder = this.private_ResolveBordersConflict(oTempRow.GetCell(nTempCurCell - 1).GetBorders().Right, oTempCellBorders.Left, false, false);
							if (border_Single === oLeftBorder.Value && oLeftBorder.Size > Max_l_w)
								Max_l_w = oLeftBorder.Size;

							Borders_Info.Left.push(oLeftBorder);
						}

						// Обработка правой границы
						if (oTempRow.GetCellsCount() - 1 === nTempCurCell)
						{
							var oRightBorder = this.private_ResolveBordersConflict(TableBorders.Right, oTempCellBorders.Right, true, false);
							if (border_Single === oRightBorder.Value && oRightBorder.Size > Max_r_w)
								Max_r_w = oRightBorder.Size;

							Borders_Info.Right.push(oRightBorder);
						}
						else
						{
							var oRightBorder = this.private_ResolveBordersConflict(oTempRow.GetCell(nTempCurCell + 1).GetBorders().Left, oTempCellBorders.Right, false, false);
							if (border_Single === oRightBorder.Value && oRightBorder.Size > Max_r_w)
								Max_r_w = oRightBorder.Size;

							Borders_Info.Right.push(oRightBorder);
						}
					}

                    Borders_Info.Right_Max = Max_r_w;
                    Borders_Info.Left_Max  = Max_l_w;

                    if ( Max_l_w / 2 > CellMar.Left.W )
                        X_content_start += Max_l_w / 2;
                    else
                        X_content_start += CellMar.Left.W;

                    if ( Max_r_w / 2 > CellMar.Right.W )
                        X_content_end -= Max_r_w / 2;
                    else
                        X_content_end -= CellMar.Right.W;

                    Cell.Set_BorderInfo_Left ( Borders_Info.Left,  Max_l_w );
                    Cell.Set_BorderInfo_Right( Borders_Info.Right, Max_r_w );
                }
            }

            if ( 0 === CurCell )
            {
                if ( null != CellSpacing )
                {
                    Row_x_min = X_grid_start;
                    if ( border_Single === TableBorders.Left.Value )
                        Row_x_min -= TableBorders.Left.Size / 2;
                }
                else
                {
                    var BorderInfo = Cell.GetBorderInfo();
                    Row_x_min = X_grid_start - BorderInfo.MaxLeft / 2;
                }
            }

            if ( CellsCount - 1 === CurCell )
            {
                if ( null != CellSpacing )
                {
                    Row_x_max = X_grid_end;
                    if ( border_Single === TableBorders.Right.Value )
                        Row_x_max += TableBorders.Right.Size / 2;
                }
                else
                {
                    var BorderInfo = Cell.GetBorderInfo();
                    Row_x_max = X_grid_end + BorderInfo.MaxRight / 2;
                }
            }

            Cell.Set_Metrics( CurGridCol, X_grid_start, X_grid_end, X_cell_start, X_cell_end, X_content_start, X_content_end );

            CurGridCol += GridSpan;
        }

        Row.Set_Metrics_X( Row_x_min, Row_x_max );
    }

    this.RecalcInfo.TableBorders = false;
};
CTable.prototype.private_RecalculateCellTopBorder = function(oPrevRow, nCurRow, nCurGridCol, nGridSpan, oTableBorders, oCellBorders)
{
	// Ищем в предыдущей строке первую ячейку, пересекающуюся с [nCurGridCol, nCurGridCol + nGridSpan]

	var nPrevCellsCount = oPrevRow.GetCellsCount();
	var oPrevBefore     = oPrevRow.GetBefore();
	var oPrevAfter      = oPrevRow.GetAfter();

	var nPrevPos = -1;

	var nPrevGridCol = oPrevBefore.Grid;
	for (var nPrevCell = 0; nPrevCell < nPrevCellsCount; ++nPrevCell)
	{
		var oPrevCell     = oPrevRow.GetCell(nPrevCell);
		var nPrevGridSpan = oPrevCell.GetGridSpan();

		if (nPrevGridCol <= nCurGridCol + nGridSpan - 1 && nPrevGridCol + nPrevGridSpan - 1 >= nCurGridCol)
		{
			nPrevPos = nPrevCell;
			break;
		}

		nPrevGridCol += nPrevGridSpan;
	}

	var arrBorderTopInfo = [];
	var nMaxTopBorder    = 0;

	// Сначала посмотрим пересечение с Before.Grid предыдущей строки
	if (nCurGridCol <= oPrevBefore.Grid - 1)
	{
		var oBorder  = this.private_ResolveBordersConflict(oTableBorders.Left, oCellBorders.Top, true, false);
		var nBorderW = oBorder.GetWidth();
		if (nMaxTopBorder < nBorderW)
			nMaxTopBorder = nBorderW;

		for (var nCurGrid = 0, nGridCount = Math.min(oPrevBefore.Grid - nCurGridCol, nGridSpan); nCurGrid < nGridCount; ++nCurGrid)
			arrBorderTopInfo.push(oBorder);
	}

	if (-1 !== nPrevPos)
	{
		while (nPrevGridCol <= nCurGridCol + nGridSpan - 1 && nPrevPos < nPrevCellsCount)
		{
			var oPrevCell     = oPrevRow.GetCell(nPrevPos);
			var nPrevGridSpan = oPrevCell.GetGridSpan();

			// Если данная ячейка участвует в вертикальном объединении, тогда нам нужно использовать нижнюю ячейку
			// в этом объединении, но в не в случае, когда мы расчитываем границу соприкасающуюся с заголовком таблицы
			// TODO: Надо проверить, зачем вообще это тут добавлено, т.к. ячейка предыдущей строки по логике
			//       не должны быть не последней в своем вертикальном объединении
			if (vmerge_Continue === oPrevCell.GetVMerge() && oPrevRow === this.GetRow(nCurRow - 1))
				oPrevCell = this.Internal_Get_EndMergedCell(nCurRow - 1, nPrevGridCol, nPrevGridSpan);

			var oPrevBottom = oPrevCell.GetBorders().Bottom;

			var oBorder  = this.private_ResolveBordersConflict(oPrevBottom, oCellBorders.Top, false, false);
			var nBorderW = oBorder.GetWidth();
			if (nMaxTopBorder < nBorderW)
				nMaxTopBorder = nBorderW;

			// Надо добавить столько раз, сколько колонок находится в пересечении этих двух ячееки
			var nGridCount = 0;
			if (nPrevGridCol >= nCurGridCol)
			{
				if (nPrevGridCol + nPrevGridSpan - 1 > nCurGridCol + nGridSpan - 1)
					nGridCount = nCurGridCol + nGridSpan - nPrevGridCol;
				else
					nGridCount = nPrevGridSpan;
			}
			else if (nPrevGridCol + nPrevGridSpan - 1 > nCurGridCol + nGridSpan - 1)
			{
				nGridCount = nGridSpan;
			}
			else
			{
				nGridCount = nPrevGridCol + nPrevGridSpan - nCurGridCol;
			}

			for (var nCurGrid = 0; nCurGrid < nGridCount; ++nCurGrid)
				arrBorderTopInfo.push(oBorder);

			nPrevPos++;
			nPrevGridCol += nPrevGridSpan;
		}
	}

	// Посмотрим пересечение с GridAfter предыдущей строки
	if (oPrevAfter.Grid > 0)
	{
		var nStartAfterGrid = oPrevRow.GetCellInfo(nPrevCellsCount - 1).StartGridCol + oPrevRow.GetCell(nPrevCellsCount - 1).GetGridSpan();
		if (nCurGridCol + nGridSpan - 1 >= nStartAfterGrid)
		{
			var oBorder  = this.private_ResolveBordersConflict(oTableBorders.Right, oCellBorders.Top, true, false);
			var nBorderW = oBorder.GetWidth();
			if (nMaxTopBorder < nBorderW)
				nMaxTopBorder = nBorderW;

			for (var nCurGrid = 0, nGridCount = Math.min(nCurGridCol + nGridSpan - nStartAfterGrid, nGridSpan); nCurGrid < nGridCount; ++nCurGrid)
				arrBorderTopInfo.push(oBorder);
		}
	}

	return {
		Info : arrBorderTopInfo,
		Max  : nMaxTopBorder
	};
};
CTable.prototype.private_RecalculateHeader = function()
{
    // Если у нас таблица внутри таблицы, тогда в ней заголовочных строк не должно быть,
    // потому что так делает Word.
    if (!this.Parent || true === this.Parent.IsTableCellContent())
    {
        this.HeaderInfo.Count = 0;
        return;
    }

    // Здесь мы подготавливаем информацию для пересчета заголовка таблицы
    var Header_RowsCount = 0;
    var Rows_Count = this.Content.length;
    for ( var Index = 0; Index < Rows_Count; Index++ )
    {
        var Row = this.Content[Index];
        if ( true != Row.IsHeader() )
            break;

        Header_RowsCount++;
    }

    // Избавимся от строк, в которых есть вертикально объединенные ячейки, которые одновременно есть в заголовке
    // и не в заголовке
    for ( var CurRow = Header_RowsCount - 1; CurRow >= 0; CurRow-- )
    {
        var Row = this.Content[CurRow];
        var Cells_Count = Row.Get_CellsCount();

        var bContinue = false;
        for ( var CurCell = 0; CurCell < Cells_Count; CurCell++ )
        {
            var Cell        = Row.Get_Cell( CurCell );
            var GridSpan    = Cell.Get_GridSpan();
            var CurGridCol  = Cell.Metrics.StartGridCol;
            var VMergeCount = this.Internal_GetVertMergeCount( CurRow, CurGridCol, GridSpan );

            // В данной строке нашли вертикально объединенную ячейку с ячейкой не из заголовка
            // Поэтому выкидываем данную строку и проверяем предыдущую
            if ( VMergeCount > 1 )
            {
                Header_RowsCount--;
                bContinue = true;
                break;
            }
        }

        if ( true != bContinue )
        {
            // Если дошли до этого места, значит данная строка, а, следовательно, и все строки выше
            // нормальные в плане объединенных вертикально ячеек.
            break;
        }
    }

    this.HeaderInfo.Count = Header_RowsCount;
};
CTable.prototype.private_RecalculatePageXY = function(CurPage)
{
    var FirstRow = 0;
    if (0 !== CurPage)
    {
        if (true === this.IsEmptyPage(CurPage - 1))
            FirstRow = this.Pages[CurPage - 1].FirstRow;
        else if (true === this.Pages[CurPage - 1].LastRowSplit)
            FirstRow = this.Pages[CurPage - 1].LastRow;
        else
            FirstRow = Math.min(this.Pages[CurPage - 1].LastRow + 1, this.Content.length - 1);
    }

    var TempMaxTopBorder = this.Get_MaxTopBorder(FirstRow);
    this.Pages.length = Math.max(CurPage, 0);
    if (0 === CurPage)
    {
    	this.Pages.length = CurPage + 1;
        this.Pages[CurPage] = new CTablePage(this.X, this.Y, this.XLimit, this.YLimit, FirstRow, TempMaxTopBorder);
    }
    else
    {
        var StartPos = this.Parent.Get_PageContentStartPos2(this.PageNum, this.ColumnNum, CurPage, this.Index);

		this.Pages.length = CurPage + 1;
        this.Pages[CurPage] = new CTablePage(StartPos.X, StartPos.Y, StartPos.XLimit, StartPos.YLimit, FirstRow, TempMaxTopBorder);
    }
};
CTable.prototype.private_RecalculatePositionX = function(CurPage)
{
	var isHdtFtr = this.Parent.IsHdrFtr();

    var TablePr = this.Get_CompiledPr(false).TablePr;
    var PageLimits = this.Parent.Get_PageLimits(this.PageNum);
    var PageFields = this.Parent.Get_PageFields(this.PageNum, isHdtFtr);

	var LD_PageLimits = this.LogicDocument.Get_PageLimits(this.Get_StartPage_Absolute());
	var LD_PageFields = this.LogicDocument.Get_PageFields(this.Get_StartPage_Absolute(), isHdtFtr);

    if ( true === this.Is_Inline() )
    {
        var Page = this.Pages[CurPage];
        if (0 === CurPage)
        {
            this.AnchorPosition.CalcX = this.X_origin + TablePr.TableInd;
            this.AnchorPosition.Set_X(this.TableSumGrid[this.TableSumGrid.length - 1], this.X_origin, LD_PageFields.X, LD_PageFields.XLimit, LD_PageLimits.XLimit, PageLimits.X, PageLimits.XLimit);
        }

        switch (TablePr.Jc)
        {
            case AscCommon.align_Left :
            {
                Page.X = Page.X_origin + this.GetTableOffsetCorrection() + TablePr.TableInd;
                break;
            }
            case AscCommon.align_Right :
            {
                var TableWidth = this.TableSumGrid[this.TableSumGrid.length - 1];

                if (false === this.Parent.IsTableCellContent())
                    Page.X = Page.XLimit - TableWidth + this.GetRightTableOffsetCorrection();
                else
                    Page.X = Page.XLimit - TableWidth;
                break;
            }
            case AscCommon.align_Center :
            {
                var TableWidth = this.TableSumGrid[this.TableSumGrid.length - 1];
                var RangeWidth = Page.XLimit - Page.X_origin;

                Page.X = Page.X_origin + ( RangeWidth - TableWidth ) / 2;
                break;
            }
        }
    }
    else
    {
        if (0 === CurPage)
        {
        	var oSectPr = this.Get_SectPr();
        	if (oSectPr)
			{
				var oFrame = oSectPr.GetContentFrame(this.GetAbsolutePage(CurPage));

				PageFields.Y      = oFrame.Top;
				PageFields.YLimit = oFrame.Bottom;
				PageFields.X      = oFrame.Left;
				PageFields.XLimit = oFrame.Right;
			}

            var OffsetCorrection_Left  = this.GetTableOffsetCorrection();
            var OffsetCorrection_Right = this.GetRightTableOffsetCorrection();

            this.X = this.X_origin + OffsetCorrection_Left;
            this.AnchorPosition.Set_X(this.TableSumGrid[this.TableSumGrid.length - 1], this.X_origin, PageFields.X + OffsetCorrection_Left, PageFields.XLimit + OffsetCorrection_Right, LD_PageLimits.XLimit, PageLimits.X + OffsetCorrection_Left, PageLimits.XLimit + OffsetCorrection_Right);

            // Непонятно по какой причине, но Word для плавающих таблиц добаляется значение TableInd
			this.AnchorPosition.Calculate_X(this.PositionH.RelativeFrom, this.PositionH.Align, this.PositionH.Value);
			this.AnchorPosition.CalcX += TablePr.TableInd;

            this.X        = this.AnchorPosition.CalcX;
            this.X_origin = this.X - OffsetCorrection_Left;

			if (undefined != this.PositionH_Old)
            {
                // Восстанови старые значения, чтобы в историю изменений все нормально записалось
                this.PositionH.RelativeFrom = this.PositionH_Old.RelativeFrom;
                this.PositionH.Align        = this.PositionH_Old.Align;
                this.PositionH.Value        = this.PositionH_Old.Value;

                // Рассчитаем сдвиг с учетом старой привязки
                var Value = this.AnchorPosition.Calculate_X_Value(this.PositionH_Old.RelativeFrom);
                this.Set_PositionH(this.PositionH_Old.RelativeFrom, false, Value);

                // На всякий случай пересчитаем заново координату
				this.X        = this.AnchorPosition.Calculate_X(this.PositionH.RelativeFrom, this.PositionH.Align, this.PositionH.Value);
                this.X_origin = this.X - OffsetCorrection_Left;

                this.PositionH_Old = undefined;
            }
        }

        this.Pages[CurPage].X        = this.X;
        this.Pages[CurPage].XLimit   = this.XLimit;
        this.Pages[CurPage].X_origin = this.X_origin;
    }
};
CTable.prototype.private_RecalculatePage = function(CurPage)
{
    if ( true === this.TurnOffRecalc )
        return;

	var isInnerTable       = this.Parent.IsTableCellContent();
	var oTopDocument       = this.Parent.Is_TopDocument(true);
	var isTopLogicDocument = (oTopDocument instanceof CDocument ? true : false);
	var oFootnotes         = (isTopLogicDocument && !isInnerTable ? oTopDocument.Footnotes : null);
	var nPageAbs           = this.Get_AbsolutePage(CurPage);
	var nColumnAbs         = this.Get_AbsoluteColumn(CurPage);

    this.TurnOffRecalc = true;

    var FirstRow = 0;
    var LastRow  = 0;

    var ResetStartElement = false;
    if ( 0 === CurPage )
    {
        // Обнуляем таблицу суммарных высот ячеек
        for ( var Index = -1; Index < this.Content.length; Index++ )
        {
            this.TableRowsBottom[Index] = [];
            this.TableRowsBottom[Index][0] = 0;
        }
    }
    else
    {
        if (true === this.IsEmptyPage(CurPage - 1))
        {
            ResetStartElement = false;
            FirstRow = this.Pages[CurPage - 1].FirstRow;
        }
        else if (true === this.Pages[CurPage - 1].LastRowSplit)
        {
            ResetStartElement = false;
            FirstRow = this.Pages[CurPage - 1].LastRow;
        }
        else
        {
            ResetStartElement = true;
            FirstRow = Math.min(this.Pages[CurPage - 1].LastRow + 1, this.Content.length - 1);
        }

        LastRow = FirstRow;
    }

    var MaxTopBorder     = this.MaxTopBorder;
    var MaxBotBorder     = this.MaxBotBorder;
    var MaxBotMargin     = this.MaxBotMargin;

    var StartPos = this.Pages[CurPage];
    if (true === this.Check_EmptyPages(CurPage - 1))
        this.HeaderInfo.PageIndex = -1;

    var Page = this.Pages[CurPage];
    var TempMaxTopBorder = Page.MaxTopBorder;

    var Y = StartPos.Y;
    var TableHeight = 0;

    var oLogicDocument = this.GetLogicDocument();
    if (oLogicDocument && oLogicDocument.IsDocumentEditor() && this.IsInline())
	{
		var nTableX_min = -1;
		var nTableX_max = -1;

		for (var nCurRow = 0, nRowsCount = this.GetRowsCount(); nCurRow < nRowsCount; ++nCurRow)
		{
			var oRow = this.GetRow(nCurRow);
			var nCellsCount = oRow.GetCellsCount();

			if (!nCellsCount)
				continue;

			var nRowX_min = oRow.GetCell(0).Metrics.X_content_start;
			var nRowX_max = oRow.GetCell(nCellsCount - 1).Metrics.X_content_end;

			if (-1 === nTableX_min || nRowX_min < nTableX_min)
				nTableX_min = nRowX_min;

			if (-1 === nTableX_max || nRowX_max > nTableX_max)
				nTableX_max = nRowX_max;
		}

		nTableX_min += Page.X;
		nTableX_max += Page.X;

		var arrRanges = this.Parent.CheckRange(nTableX_min, Page.Y, nTableX_max, Page.Y + 0.001, Page.Y, Page.Y + 0.001, nTableX_min, nTableX_max, this.private_GetRelativePageIndex(CurPage));
		if (arrRanges.length > 0)
		{
			for (var nRangeIndex = 0, nRangesCount = arrRanges.length; nRangeIndex < nRangesCount; ++nRangeIndex)
			{
				if (Y < arrRanges[nRangeIndex].Y1)
				{
					var nShiftY = arrRanges[nRangeIndex].Y1 - Y;

					Y      = arrRanges[nRangeIndex].Y1 + 0.001;
					Page.Y = Y;

					Page.Bounds.Top += nShiftY;
				}
			}
		}
	}

    var TableBorders = this.Get_Borders();

    var nHeaderMaxTopBorder = -1;

    var X_max = -1;
    var X_min = -1;
	if (this.HeaderInfo.Count > 0 && this.HeaderInfo.PageIndex != -1 && CurPage > this.HeaderInfo.PageIndex && this.IsInline())
    {
    	this.HeaderInfo.HeaderRecalculate = true;
        this.HeaderInfo.Pages[CurPage] = {};
        this.HeaderInfo.Pages[CurPage].RowsInfo = [];
        var HeaderPage = this.HeaderInfo.Pages[CurPage];

        // Рисуем ли заголовок на данной странице
        HeaderPage.Draw = true;

        // Скопируем целиком строки
        HeaderPage.Rows = [];

        // Временно отключаем регистрацию новых классов
        AscCommon.g_oTableId.m_bTurnOff = true;
        AscCommon.History.TurnOff();

		this.LogicDocument.RecalcTableHeader = true;

		var aContentDrawings = [];
        for ( var Index = 0; Index < this.HeaderInfo.Count; Index++ )
        {
            HeaderPage.Rows[Index] = this.Content[Index].Copy(this);
            HeaderPage.Rows[Index].Index = Index;
            for(var CellIndex = 0; CellIndex < HeaderPage.Rows[Index].Content.length; ++CellIndex)
            {
                HeaderPage.Rows[Index].Content[CellIndex].Content.GetAllDrawingObjects(aContentDrawings);
            }
        }
        for(var DrawingIndex = 0; DrawingIndex < aContentDrawings.length; ++DrawingIndex)
        {
            if(aContentDrawings[DrawingIndex] && aContentDrawings[DrawingIndex].GraphicObj)
            {
                aContentDrawings[DrawingIndex].GraphicObj.recalculate();
                if(aContentDrawings[DrawingIndex].GraphicObj.recalculateText)
                {
                    aContentDrawings[DrawingIndex].GraphicObj.recalculateText();
                }
            }
        }

        AscCommon.g_oTableId.m_bTurnOff = false;
        AscCommon.History.TurnOn();

        var bHeaderNextPage = false;
        for ( var CurRow = 0; CurRow < this.HeaderInfo.Count; CurRow++  )
        {
            HeaderPage.RowsInfo[CurRow] = {};
            HeaderPage.RowsInfo[CurRow].Y               = 0;
            HeaderPage.RowsInfo[CurRow].H               = 0;
            HeaderPage.RowsInfo[CurRow].TopDy           = 0;
            HeaderPage.RowsInfo[CurRow].MaxTopBorder    = 0;
            HeaderPage.RowsInfo[CurRow].TableRowsBottom = 0;

            var Row         = HeaderPage.Rows[CurRow];
            var CellsCount  = Row.Get_CellsCount();
            var CellSpacing = Row.Get_CellSpacing();

            var BeforeInfo  = Row.Get_Before();
            var CurGridCol  = BeforeInfo.GridBefore;

            // Добавляем ширину верхней границы у текущей строки (используем MaxTopBorder самой таблицы)
            Y           += MaxTopBorder[CurRow];
            TableHeight += MaxTopBorder[CurRow];

            // Если таблица с расстоянием между ячейками, тогда добавляем его
            if ( 0 === CurRow )
            {
                if ( null != CellSpacing )
                {
                    var TableBorder_Top = this.Get_Borders().Top;
                    if ( border_Single === TableBorder_Top.Value )
                    {
                        Y           += TableBorder_Top.Size;
                        TableHeight += TableBorder_Top.Size;
                    }

                    Y           += CellSpacing;
                    TableHeight += CellSpacing;
                }
            }
            else
            {
                var PrevCellSpacing = HeaderPage.Rows[CurRow - 1].Get_CellSpacing();

                if ( null != CellSpacing && null != PrevCellSpacing )
                {
                    Y           += (PrevCellSpacing + CellSpacing) / 2;
                    TableHeight += (PrevCellSpacing + CellSpacing) / 2;
                }
                else if ( null != CellSpacing )
                {
                    Y           += CellSpacing / 2;
                    TableHeight += CellSpacing / 2;
                }
                else if ( null != PrevCellSpacing )
                {
                    Y           += PrevCellSpacing / 2;
                    TableHeight += PrevCellSpacing / 2;
                }
            }

            var Row_x_max = Row.Metrics.X_max;
            var Row_x_min = Row.Metrics.X_min;

            if ( -1 === X_min || Row_x_min < X_min )
                X_min = Row_x_min;

            if ( -1 === X_max || Row_x_max > X_max )
                X_max = Row_x_max;

            // Дополнительный параметр для случая, если данная строка начнется с новой страницы.
            // Мы запоминаем максимальное значение нижней границы(первой страницы (текущей)) у ячеек,
            // объединенных вертикально так, чтобы это объединение заканчивалось на данной строке.
            // И если данная строка начнется сразу с новой страницы (Pages > 0, FirstPage = false), тогда
            // мы должны данный параметр сравнить со значением нижней границы предыдущей строки.
            var MaxBotValue_vmerge = -1;

            var RowH = Row.Get_Height();
            var VerticallCells = [];

            for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
            {
                var Cell     = Row.Get_Cell( CurCell );
                var GridSpan = Cell.Get_GridSpan();
                var Vmerge   = Cell.GetVMerge();
                var CellMar  = Cell.GetMargins();

                Row.Update_CellInfo(CurCell);

                var CellMetrics = Row.Get_CellInfo( CurCell );

                var X_content_start = Page.X + CellMetrics.X_content_start;
                var X_content_end   = Page.X + CellMetrics.X_content_end;

                var Y_content_start = Y + CellMar.Top.W;
                var Y_content_end   = this.Pages[CurPage].YLimit;

                // TODO: При расчете YLimit для ячейки сделать учет толщины нижних
                //       границ ячейки и таблицы
                if ( null != CellSpacing )
                {
                    if ( this.Content.length - 1 === CurRow )
                        Y_content_end -= CellSpacing;
                    else
                        Y_content_end -= CellSpacing / 2;
                }

                var VMergeCount = this.Internal_GetVertMergeCount( CurRow, CurGridCol, GridSpan );
                var BottomMargin = this.MaxBotMargin[CurRow + VMergeCount - 1];
                Y_content_end -= BottomMargin;

                // Такие ячейки мы обсчитываем, если либо сейчас происходит переход на новую страницу, либо
                // это последняя ячейка в объединении.
                // Обсчет такик ячеек произошел ранее

                Cell.Temp.Y = Y_content_start;

                if ( VMergeCount > 1 )
                {
                    CurGridCol += GridSpan;
                    continue;
                }
                else
                {
                    // Возьмем верхнюю ячейку теккущего объединения
                    if ( vmerge_Restart != Vmerge )
                    {
                        // Найдем ячейку в самой таблице, а дальше по индексам ячейки и строки получим ее в скопированном заголовке
                        Cell    = this.Internal_Get_StartMergedCell( CurRow, CurGridCol, GridSpan );
                        var cIndex = Cell.Index;
                        var rIndex = Cell.Row.Index;

                        Cell = HeaderPage.Rows[rIndex].Get_Cell( cIndex );

                        CellMar = Cell.GetMargins();

                        Y_content_start = Cell.Temp.Y + CellMar.Top.W;
                    }
                }

                if (true === Cell.IsVerticalText())
                {
                    VerticallCells.push(Cell);
                    CurGridCol += GridSpan;
                    continue;
                }

                Cell.Content.Set_StartPage( CurPage );
                Cell.Content.Reset( X_content_start, Y_content_start, X_content_end, Y_content_end );
                Cell.Content.Set_ClipInfo(0, Page.X + CellMetrics.X_cell_start, Page.X + CellMetrics.X_cell_end);
                Cell.Content.RecalculateEndInfo();

                if ( recalcresult2_NextPage === Cell.Content.Recalculate_Page( 0, true ) )
                {
                    bHeaderNextPage = true;
                    break;
                }

                var CellContentBounds = Cell.Content.Get_PageBounds( 0, undefined, true );
                var CellContentBounds_Bottom = CellContentBounds.Bottom + BottomMargin;

                if ( undefined === HeaderPage.RowsInfo[CurRow].TableRowsBottom || HeaderPage.RowsInfo[CurRow].TableRowsBottom < CellContentBounds_Bottom )
                    HeaderPage.RowsInfo[CurRow].TableRowsBottom = CellContentBounds_Bottom;

                if ( vmerge_Continue === Vmerge )
                {
                    if ( -1 === MaxBotValue_vmerge || MaxBotValue_vmerge < CellContentBounds_Bottom )
                        MaxBotValue_vmerge = CellContentBounds_Bottom;
                }

                CurGridCol += GridSpan;
            }

            // Если заголовок целиком на странице не убирается, тогда мы его попросту не рисуем на данной странице
            if ( true === bHeaderNextPage )
            {
                Y = StartPos.Y;
                TableHeight = 0;
                HeaderPage.Draw = false;
                break;
            }


            // Здесь мы выставляем только начальную координату строки (для каждой страницы)
            // высоту строки(для каждой страницы) мы должны обсчитать после общего цикла, т.к.
            // в одной из следйющих строк может оказаться ячейка с вертикальным объединением,
            // захватывающим данную строку. Значит, ее содержимое может изменить высоту нашей строки.
            var TempY            = Y;
            var TempMaxTopBorder = MaxTopBorder[CurRow];

            if ( null != CellSpacing )
            {
                HeaderPage.RowsInfo[CurRow].Y            = TempY;
                HeaderPage.RowsInfo[CurRow].TopDy        = 0;
                HeaderPage.RowsInfo[CurRow].X0           = Row_x_min;
                HeaderPage.RowsInfo[CurRow].X1           = Row_x_max;
                HeaderPage.RowsInfo[CurRow].MaxTopBorder = TempMaxTopBorder;
                HeaderPage.RowsInfo[CurRow].MaxBotBorder = MaxBotBorder[CurRow];
            }
            else
            {
                HeaderPage.RowsInfo[CurRow].Y            = TempY - TempMaxTopBorder;
                HeaderPage.RowsInfo[CurRow].TopDy        = TempMaxTopBorder;
                HeaderPage.RowsInfo[CurRow].X0           = Row_x_min;
                HeaderPage.RowsInfo[CurRow].X1           = Row_x_max;
                HeaderPage.RowsInfo[CurRow].MaxTopBorder = TempMaxTopBorder;
                HeaderPage.RowsInfo[CurRow].MaxBotBorder = MaxBotBorder[CurRow];
            }

            var CellHeight = HeaderPage.RowsInfo[CurRow].TableRowsBottom - Y;

            // TODO: улучшить проверку на высоту строки (для строк разбитых на страницы)
            if (false === bHeaderNextPage && (Asc.linerule_AtLeast === RowH.HRule || Asc.linerule_Exact === RowH.HRule) && CellHeight < RowH.Value - MaxTopBorder[CurRow])
            {
                CellHeight = RowH.Value - MaxTopBorder[CurRow];
                HeaderPage.RowsInfo[CurRow].TableRowsBottom = Y + CellHeight;
            }

            // Рассчитываем ячейки с вертикальным текстом
            var CellsCount2 = VerticallCells.length;
            for (var TempCellIndex = 0; TempCellIndex < CellsCount2; TempCellIndex++)
            {
                var Cell       = VerticallCells[TempCellIndex];
                var CurCell    = Cell.Index;
                var GridSpan   = Cell.Get_GridSpan();
                var CurGridCol = Cell.Metrics.StartGridCol;

                // Возьмем верхнюю ячейку текущего объединения
                Cell = this.Internal_Get_StartMergedCell(CurRow, CurGridCol, GridSpan);
                var cIndex = Cell.Index;
                var rIndex = Cell.Row.Index;
                Cell = HeaderPage.Rows[rIndex].Get_Cell( cIndex );

                var CellMar     = Cell.GetMargins();
                var CellMetrics = Cell.Row.Get_CellInfo(CurCell);

                var X_content_start = Page.X + CellMetrics.X_content_start;
                var X_content_end   = Page.X + CellMetrics.X_content_end;
                var Y_content_start = Cell.Temp.Y;
                var Y_content_end   = HeaderPage.RowsInfo[CurRow].TableRowsBottom;

                // TODO: При расчете YLimit для ячейки сделать учет толщины нижних
                //       границ ячейки и таблицы
                if (null != CellSpacing)
                {
                    if (this.Content.length - 1 === CurRow)
                        Y_content_end -= CellSpacing;
                    else
                        Y_content_end -= CellSpacing / 2;
                }

                var VMergeCount = this.Internal_GetVertMergeCount(CurRow, CurGridCol, GridSpan);
                var BottomMargin = this.MaxBotMargin[CurRow + VMergeCount - 1];
                Y_content_end -= BottomMargin;

                Cell.PagesCount = 1;
                Cell.Content.Set_StartPage(CurPage);
                Cell.Content.Reset(0, 0, Y_content_end - Y_content_start, 10000);
                Cell.Temp.X_start = X_content_start;
                Cell.Temp.Y_start = Y_content_start;
                Cell.Temp.X_end   = X_content_end;
                Cell.Temp.Y_end   = Y_content_end;

                Cell.Temp.X_cell_start = Page.X + CellMetrics.X_cell_start;
                Cell.Temp.X_cell_end   = Page.X + CellMetrics.X_cell_end;
                Cell.Temp.Y_cell_start = Y_content_start - CellMar.Top.W;
                Cell.Temp.Y_cell_end   = Y_content_end + BottomMargin;


                // Какие-то ячейки в строке могут быть не разбиты на строки, а какие то разбиты.
                // Здесь контролируем этот момент, чтобы у тех, которые не разбиты не вызывать
                // Recalculate_Page от несуществующих страниц.
                var CellPageIndex = CurPage - Cell.Content.Get_StartPage_Relative();
                if (0 === CellPageIndex)
                {
                    Cell.Content.Recalculate_Page(CellPageIndex, true);
                }
            }

            if ( null != CellSpacing )
                HeaderPage.RowsInfo[CurRow].H = CellHeight;
            else
                HeaderPage.RowsInfo[CurRow].H = CellHeight + TempMaxTopBorder;

            Y           += CellHeight;
            TableHeight += CellHeight;

            Row.Height   = CellHeight;

            Y           += MaxBotBorder[CurRow];
            TableHeight += MaxBotBorder[CurRow];

            // Сделаем вертикальное выравнивание ячеек в таблице. Делаем как Word, если ячейка разбилась на несколько
            // страниц, тогда вертикальное выравнивание применяем только к первой странице.
        }

        if ( false === bHeaderNextPage )
        {
            // Сделаем вертикальное выравнивание ячеек в таблице. Делаем как Word, если ячейка разбилась на несколько
            // страниц, тогда вертикальное выравнивание применяем только к первой странице.
            // Делаем это не в общем цикле, потому что объединенные вертикально ячейки могут вносить поправки в значения
            // this.TableRowsBottom, в последней строке.
            for ( var CurRow = 0; CurRow < this.HeaderInfo.Count; CurRow++ )
            {
                var Row = HeaderPage.Rows[CurRow];
                var CellsCount = Row.Get_CellsCount();
                for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
                {
                    var Cell = Row.Get_Cell( CurCell );
                    var VMergeCount = this.Internal_GetVertMergeCount( CurRow, Cell.Metrics.StartGridCol, Cell.Get_GridSpan() );

                    if ( VMergeCount > 1 )
                        continue;
                    else
                    {
                        var Vmerge = Cell.GetVMerge();
                        // Возьмем верхнюю ячейку теккущего объединения
                        if ( vmerge_Restart != Vmerge )
                        {
                            Cell = this.Internal_Get_StartMergedCell( CurRow, Cell.Metrics.StartGridCol, Cell.Get_GridSpan() );
                            var cIndex = Cell.Index;
                            var rIndex = Cell.Row.Index;

                            Cell = HeaderPage.Rows[rIndex].Get_Cell( cIndex );
                        }
                    }

                    var CellMar       = Cell.GetMargins();
                    var VAlign        = Cell.Get_VAlign();
                    var CellPageIndex = CurPage - Cell.Content.Get_StartPage_Relative();

                    if ( CellPageIndex >= Cell.PagesCount )
                        continue;

                    // Для прилегания к верху или для второй страницы ничего не делаем (так изначально рассчитывалось)
                    if ( vertalignjc_Top === VAlign || CellPageIndex > 1 )
                    {
                        Cell.Temp.Y_VAlign_offset[CellPageIndex] = 0;
                        continue;
                    }

                    // Рассчитаем имеющуюся в распоряжении высоту ячейки
                    var TempCurRow = Cell.Row.Index;
                    var TempCellSpacing = HeaderPage.Rows[TempCurRow].Get_CellSpacing();
                    var Y_0 = HeaderPage.RowsInfo[TempCurRow].Y;

                    if ( null === TempCellSpacing )
                        Y_0 += MaxTopBorder[TempCurRow];

                    Y_0 += CellMar.Top.W;

                    var Y_1 = HeaderPage.RowsInfo[CurRow].TableRowsBottom - CellMar.Bottom.W;
                    var CellHeight = Y_1 - Y_0;

                    var CellContentBounds = Cell.Content.Get_PageBounds( CellPageIndex, CellHeight, true );
                    var ContentHeight = CellContentBounds.Bottom - CellContentBounds.Top;

                    var Dy = 0;

                    if (true === Cell.IsVerticalText())
                    {
                        var CellMetrics = Cell.Row.Get_CellInfo(Cell.Index);
                        CellHeight = CellMetrics.X_cell_end - CellMetrics.X_cell_start - CellMar.Left.W - CellMar.Right.W;
                    }

                    if ( CellHeight - ContentHeight > 0.001 )
                    {
                        if ( vertalignjc_Bottom === VAlign )
                            Dy = CellHeight - ContentHeight;
                        else if ( vertalignjc_Center === VAlign )
                            Dy = (CellHeight - ContentHeight) / 2;

                        Cell.ShiftCellContent(CellPageIndex, 0, Dy);
                    }

                    Cell.Temp.Y_VAlign_offset[CellPageIndex] = Dy;
                }
            }
        }

        nHeaderMaxTopBorder = this.private_GetMaxTopBorderWidth(FirstRow, true);

		this.LogicDocument.RecalcTableHeader = false;
    }
    else
    {
        this.HeaderInfo.Pages[CurPage] = {};
        this.HeaderInfo.Pages[CurPage].Draw = false;
    }
    this.HeaderInfo.HeaderRecalculate = false;

    var bNextPage = false;

    // Блок переменных для учета сносок
	var nFootnotesHeight     = 0;
	var arrSavedY            = [];
	var arrSavedTableHeight  = [];
	var arrFootnotesObject   = [];
	var nResetFootnotesIndex = -1;

    for (var CurRow = FirstRow; CurRow < this.Content.length; ++CurRow)
    {
		if (oFootnotes && (-1 === nResetFootnotesIndex || CurRow > nResetFootnotesIndex))
		{
			nFootnotesHeight            = oFootnotes.GetHeight(nPageAbs, nColumnAbs);
			arrFootnotesObject[CurRow]  = oFootnotes.SaveRecalculateObject(nPageAbs, nColumnAbs);
			arrSavedY[CurRow]           = Y;
			arrSavedTableHeight[CurRow] = TableHeight;

			this.Pages[CurPage].FootnotesH = nFootnotesHeight;
		}

		if ((0 === CurRow && true === this.Check_EmptyPages(CurPage - 1)) || CurRow != FirstRow || (CurRow === FirstRow && true === ResetStartElement))
        {
            this.RowsInfo[CurRow] = new CTableRowsInfo();
            this.RowsInfo[CurRow].StartPage = CurPage;
            this.TableRowsBottom[CurRow]    = [];
        }
        else
        {
            this.RowsInfo[CurRow].Pages = CurPage - this.RowsInfo[CurRow].StartPage + 1;
        }

        this.TableRowsBottom[CurRow][CurPage] = Y;

        var Row         = this.Content[CurRow];
        var CellsCount  = Row.Get_CellsCount();
        var CellSpacing = Row.Get_CellSpacing();

        var BeforeInfo  = Row.Get_Before();
        var AfterInfo   = Row.Get_After();
        var CurGridCol  = BeforeInfo.GridBefore;

        var nMaxTopBorder = MaxTopBorder[CurRow];
        if (CurRow === FirstRow && nHeaderMaxTopBorder > 0)
        	nMaxTopBorder = nHeaderMaxTopBorder;

        // Добавляем ширину верхней границы у текущей строки
        Y           += nMaxTopBorder;
        TableHeight += nMaxTopBorder;

        // Если таблица с расстоянием между ячейками, тогда добавляем его
        if (FirstRow === CurRow)
        {
            if (null != CellSpacing)
            {
                var TableBorder_Top = this.Get_Borders().Top;
                if (border_Single === TableBorder_Top.Value)
                {
                    Y           += TableBorder_Top.Size;
                    TableHeight += TableBorder_Top.Size;
                }

                if (true === this.HeaderInfo.Pages[CurPage].Draw || (0 === CurRow && (0 === CurPage || (1 === CurPage && false === this.RowsInfo[0].FirstPage))))
                {
                    Y           += CellSpacing;
                    TableHeight += CellSpacing;
                }
                else
                {
                    Y           += CellSpacing / 2;
                    TableHeight += CellSpacing / 2;
                }
            }
        }
        else
        {
            var PrevCellSpacing = this.Content[CurRow - 1].Get_CellSpacing();

            if (null != CellSpacing && null != PrevCellSpacing)
            {
                Y           += (PrevCellSpacing + CellSpacing) / 2;
                TableHeight += (PrevCellSpacing + CellSpacing) / 2;
            }
            else if (null != CellSpacing)
            {
                Y           += CellSpacing / 2;
                TableHeight += CellSpacing / 2;
            }
            else if (null != PrevCellSpacing)
            {
                Y           += PrevCellSpacing / 2;
                TableHeight += PrevCellSpacing / 2;
            }
        }

        var Row_x_max = Row.Metrics.X_max;
        var Row_x_min = Row.Metrics.X_min;

        if (-1 === X_min || Row_x_min < X_min)
            X_min = Row_x_min;

        if (-1 === X_max || Row_x_max > X_max)
            X_max = Row_x_max;

		var MaxTopMargin = 0;

		for (var CurCell = 0; CurCell < CellsCount; CurCell++)
		{
			var Cell    = Row.Get_Cell(CurCell);
			var Vmerge  = Cell.GetVMerge();
			var CellMar = Cell.GetMargins();
			if (vmerge_Restart === Vmerge && CellMar.Top.W > MaxTopMargin)
				MaxTopMargin = CellMar.Top.W;
		}

		var RowH = Row.Get_Height();
		// В данном значении не учитываются маргины
		var RowHValue = RowH.Value + this.MaxBotMargin[CurRow] + MaxTopMargin;

		// В таблице с отступами размер отступа входит в значение высоты строки
		if (null !== CellSpacing)
			RowHValue -= CellSpacing;

		// Для строк с точной высотой строк значение высоты считается вместе с шириной верхней границы
		if (Asc.linerule_Exact === RowH.HRule)
			RowHValue -= nMaxTopBorder;

		if (oFootnotes && (Asc.linerule_AtLeast === RowH.HRule || Asc.linerule_Exact == RowH.HRule))
		{
			oFootnotes.PushCellLimit(Y + RowHValue);
		}

        // Дополнительный параметр для случая, если данная строка начнется с новой страницы.
        // Мы запоминаем максимальное значение нижней границы(первой страницы (текущей)) у ячеек,
        // объединенных вертикально так, чтобы это объединение заканчивалось на данной строке.
        // И если данная строка начнется сразу с новой страницы (Pages > 0, FirstPage = false), тогда
        // мы должны данный параметр сравнить со значением нижней границы предыдущей строки.
        var MaxBotValue_vmerge = -1;

        var Merged_Cell  = [];
        var VerticallCells = [];
        var bAllCellsVertical = true;
        var bFootnoteBreak = false;
        for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
        {
            var Cell     = Row.Get_Cell( CurCell );
            var GridSpan = Cell.Get_GridSpan();
            var Vmerge   = Cell.GetVMerge();
            var CellMar  = Cell.GetMargins();

            Row.Update_CellInfo(CurCell);

            var CellMetrics   = Row.Get_CellInfo( CurCell );

            var X_content_start = Page.X + CellMetrics.X_content_start;
            var X_content_end   = Page.X + CellMetrics.X_content_end;

            var Y_content_start = Y + CellMar.Top.W;
            var Y_content_end   = this.Pages[CurPage].YLimit - nFootnotesHeight;

            // TODO: При расчете YLimit для ячейки сделать учет толщины нижних
            //       границ ячейки и таблицы
            if ( null != CellSpacing )
            {
                if ( this.Content.length - 1 === CurRow )
                    Y_content_end -= CellSpacing;
                else
                    Y_content_end -= CellSpacing / 2;
            }

            var VMergeCount = this.Internal_GetVertMergeCount( CurRow, CurGridCol, GridSpan );
            var BottomMargin = this.MaxBotMargin[CurRow + VMergeCount - 1];
            Y_content_end -= BottomMargin;

            // Такие ячейки мы обсчитываем, если либо сейчас происходит переход на новую страницу, либо
            // это последняя ячейка в объединении.
            // Обсчет такик ячеек произошел ранее

            Cell.Temp.Y = Y_content_start;

            // Сохраняем ссылку на исходную ячейку
			var oOriginCell = Cell;
            if ( VMergeCount > 1 )
            {
                CurGridCol += GridSpan;
                Merged_Cell.push( Cell );

				// Приходится здесь обновлять EndInfo, т.к. мы можем начать пересчет следующей ячейки, до
				// окончания полного пересчета предыдущей ячейки в строке
				Cell.Content.RecalculateEndInfo();

                continue;
            }
            else
            {
                // Возьмем верхнюю ячейку текущего объединения
                if ( vmerge_Restart != Vmerge )
                {
                    Cell = this.Internal_Get_StartMergedCell( CurRow, CurGridCol, GridSpan );
                    CellMar = Cell.GetMargins();

                    var oTempRow         = Cell.GetRow();
					var oTempCellMetrics = oTempRow.GetCellInfo(Cell.GetIndex());

					X_content_start = Page.X + oTempCellMetrics.X_content_start;
					X_content_end   = Page.X + oTempCellMetrics.X_content_end;

                    Y_content_start = Cell.Temp.Y + CellMar.Top.W;
                }
            }

            if (true === Cell.IsVerticalText())
            {
                VerticallCells.push(Cell);
                CurGridCol += GridSpan;
                continue;
            }

            bAllCellsVertical = false;

            var bCanShift = false;
            var ShiftDy   = 0;
            var ShiftDx   = 0;

            if ((0 === Cell.Row.Index && true === this.Check_EmptyPages(CurPage - 1)) || Cell.Row.Index > FirstRow || (Cell.Row.Index === FirstRow && true === ResetStartElement))
            {
                Cell.Content.Set_StartPage( CurPage );

                if (  true === this.Is_Inline() && 1 === Cell.PagesCount && 1 === Cell.Content.Pages.length && true != this.RecalcInfo.Check_Cell( Cell ) )
                {
                    var X_content_start_old  = Cell.Content.Pages[0].X;
                    var X_content_end_old    = Cell.Content.Pages[0].XLimit;
                    var Y_content_height_old = Cell.Content.Pages[0].Bounds.Bottom - Cell.Content.Pages[0].Bounds.Top;

					// Проверим по X, Y
					if (Math.abs(X_content_start - X_content_start_old) < 0.001 && Math.abs(X_content_end_old - X_content_end) < 0.001 && Y_content_start + Y_content_height_old < Y_content_end)
                    {
                        bCanShift = true;
                        ShiftDy   = -Cell.Content.Pages[0].Y + Y_content_start;

						// Если в ячейке есть ссылки на сноски, тогда такую ячейку нужно пересчитывать
						var arrFootnotes = Cell.Content.GetFootnotesList(null, null, false);
						var arrEndnotes  = Cell.Content.GetFootnotesList(null, null, true);
						if ((arrFootnotes && arrFootnotes.length > 0)
							|| (arrEndnotes && arrEndnotes.length > 0))
						{
							bCanShift = false;
						}
                    }
                }

                Cell.PagesCount = 1;
                Cell.Content.Reset(X_content_start, Y_content_start, X_content_end, Y_content_end);
            }

            // Какие-то ячейки в строке могут быть не разбиты на строки, а какие то разбиты.
            // Здесь контролируем этот момент, чтобы у тех, которые не разбиты не вызывать
            // Recalculate_Page от несуществующих страниц.
            var CellPageIndex = CurPage - Cell.Content.Get_StartPage_Relative();
            Cell.Content.Set_ClipInfo(CellPageIndex, Page.X + CellMetrics.X_cell_start, Page.X + CellMetrics.X_cell_end);
            if ( CellPageIndex < Cell.PagesCount )
            {
            	// Приходится здесь обновлять EndInfo, т.к. мы можем начать пересчет следующей ячейки, до
				// окончания полного пересчета предыдущей ячейки в строке
            	Cell.Content.RecalculateEndInfo();

                if ( true === bCanShift )
                {
					Cell.ShiftCell(0, ShiftDx, ShiftDy);
                    Cell.Content.UpdateEndInfo();
                }
                else
                {
					var RecalcResult = Cell.Content.Recalculate_Page(CellPageIndex, true);
					if (recalcresult2_CurPage & RecalcResult)
					{
						var _RecalcResult = recalcresult_CurPage;

						if (RecalcResult & recalcresultflags_Column)
							_RecalcResult |= recalcresultflags_Column;

						if (RecalcResult & recalcresultflags_Footnotes)
							_RecalcResult |= recalcresultflags_Footnotes;

						this.TurnOffRecalc = false;
						return _RecalcResult;
					}
					else if (recalcresult2_NextPage & RecalcResult)
					{
						Cell.PagesCount = Cell.Content.Pages.length + 1;
						bNextPage       = true;
					}
					else if (recalcresult2_End & RecalcResult)
					{
						// Ничего не делаем
					}
                }

                var CellContentBounds = Cell.Content.Get_PageBounds( CellPageIndex, undefined, true );
                var CellContentBounds_Bottom = CellContentBounds.Bottom + BottomMargin;

                if ( undefined === this.TableRowsBottom[CurRow][CurPage] || this.TableRowsBottom[CurRow][CurPage] < CellContentBounds_Bottom )
                    this.TableRowsBottom[CurRow][CurPage] = CellContentBounds_Bottom;

                if ( vmerge_Continue === Vmerge )
                {
                    if ( -1 === MaxBotValue_vmerge || MaxBotValue_vmerge < CellContentBounds_Bottom )
                        MaxBotValue_vmerge = CellContentBounds_Bottom;
                }
            }

			var nCurFootnotesHeight = oFootnotes ? oFootnotes.GetHeight(nPageAbs, nColumnAbs) : 0;
			if (oFootnotes && nCurFootnotesHeight > nFootnotesHeight + 0.001)
			{
				this.Pages[CurPage].FootnotesH = nCurFootnotesHeight;

				nFootnotesHeight     = nCurFootnotesHeight;
				nResetFootnotesIndex = CurRow;
				Y                    = arrSavedY[oOriginCell.Row.Index];
				TableHeight          = arrSavedTableHeight[oOriginCell.Row.Index];
				oFootnotes.LoadRecalculateObject(nPageAbs, nColumnAbs, arrFootnotesObject[oOriginCell.Row.Index]);

				CurRow = oOriginCell.Row.Index - 1;

				bFootnoteBreak = true;
				break;
			}

            CurGridCol += GridSpan;
        }

		if (oFootnotes && (Asc.linerule_AtLeast === RowH.HRule || Asc.linerule_Exact == RowH.HRule))
		{
			oFootnotes.PopCellLimit();
		}

		if (bFootnoteBreak)
			continue;

        if (undefined === this.TableRowsBottom[CurRow][CurPage])
            this.TableRowsBottom[CurRow][CurPage] = Y;

        // Если в строке все ячейки с вертикальным выравниванием
        if (true === bAllCellsVertical && Asc.linerule_Auto === RowH.HRule)
            this.TableRowsBottom[CurRow][CurPage] = Y + 4.5 + this.MaxBotMargin[CurRow] + MaxTopMargin;

        if ((Asc.linerule_AtLeast === RowH.HRule || Asc.linerule_Exact == RowH.HRule) && Y + RowHValue > Y_content_end && ((0 === CurRow && 0 === CurPage && null !== this.Get_DocumentPrev() && !this.Parent.IsFirstElementOnPage(this.private_GetRelativePageIndex(CurPage), this.GetIndex())) || CurRow != FirstRow))
        {
            bNextPage = true;

            for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
            {
                var Cell   = Row.Get_Cell( CurCell );
                var Vmerge = Cell.GetVMerge();

                var VMergeCount = this.Internal_GetVertMergeCount( CurRow, Cell.Metrics.StartGridCol, Cell.Get_GridSpan() );

                // Проверяем только начальные ячейки вертикального объединения..
                if ( vmerge_Continue === Vmerge || VMergeCount > 1 )
                    continue;

                Cell.Content.StartFromNewPage();
                Cell.PagesCount = 2;
            }
        }

        // Данная строка разбилась на несколько страниц. Нам нужно сделать несколько дополнительных действий:
        // 1. Проверяем есть ли хоть какой-либо контент данной строки на первой странице, т.е. реально данная
        //    строка начинается со 2-ой страницы.
        // 2. Пересчитать все смерженные вертикально ячейки, которые также разбиваются на несколько страниц,
        //    но у которых вертикальное объединение не заканчивается на данной странице.
        if ( true === bNextPage )
        {
            var bContentOnFirstPage   = false;
            var bNoContentOnFirstPage = false;
            for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
            {
                var Cell   = Row.Get_Cell( CurCell );
                var Vmerge = Cell.GetVMerge();

                var VMergeCount = this.Internal_GetVertMergeCount( CurRow, Cell.Metrics.StartGridCol, Cell.Get_GridSpan() );

                // Проверяем только начальные ячейки вертикального объединения..
                if ( vmerge_Continue === Vmerge || VMergeCount > 1 )
                    continue;

                if (true === Cell.IsVerticalText() || true === Cell.Content_Is_ContentOnFirstPage())
                {
                    bContentOnFirstPage = true;
                }
                else
                {
                    bNoContentOnFirstPage = true;
                }
            }

            if ( true === bContentOnFirstPage && true === bNoContentOnFirstPage )
            {
                for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
                {
                    var Cell   = Row.Get_Cell( CurCell );
                    var Vmerge = Cell.GetVMerge();

                    var VMergeCount = this.Internal_GetVertMergeCount( CurRow, Cell.Metrics.StartGridCol, Cell.Get_GridSpan() );

                    // Проверяем только начальные ячейки вертикального объединения..
                    if ( vmerge_Continue === Vmerge || VMergeCount > 1 )
                        continue;

                    Cell.Content.StartFromNewPage();
                    Cell.PagesCount = 2;
                }

                bContentOnFirstPage = false;
            }

            this.RowsInfo[CurRow].FirstPage = bContentOnFirstPage;

            // Не сраниваем MaxBotValue_vmerge с -1, т.к. значения в this.TableRowsBottom в любом случае неотрицательные
            if ( 0 != CurRow && false === this.RowsInfo[CurRow].FirstPage )
            {
                if ( this.TableRowsBottom[CurRow - 1][CurPage] < MaxBotValue_vmerge )
                {
                    // Поскольку мы правим настройку не текущей строки, надо подправить и
                    // запись о рассчитанной высоте строки
                    var Diff = MaxBotValue_vmerge - this.TableRowsBottom[CurRow - 1][CurPage];
                    this.TableRowsBottom[CurRow - 1][CurPage] = MaxBotValue_vmerge;
                    this.RowsInfo[CurRow - 1].H[CurPage] += Diff;
                }
            }

            // Здесь мы должны рассчитать ячейки, которые попали в вертикальное объединение и из-за этого не были рассчитаны
            var CellsCount2 = Merged_Cell.length;
            var bFootnoteBreak = false;
            for (var TempCellIndex = 0; TempCellIndex < CellsCount2; TempCellIndex++)
            {
                var Cell = Merged_Cell[TempCellIndex];
                var CurCell = Cell.Index;
                var GridSpan = Cell.Get_GridSpan();
                var CurGridCol = Cell.Metrics.StartGridCol;

                // Возьмем верхнюю ячейку теккущего объединения
                Cell = this.Internal_Get_StartMergedCell(CurRow, CurGridCol, GridSpan);

                if (true === Cell.IsVerticalText())
                {
                    VerticallCells.push(Cell);
                    CurGridCol += GridSpan;
                    continue;
                }

                var CellMar = Cell.GetMargins();
                var CellMetrics = Row.Get_CellInfo(CurCell);

                var X_content_start = Page.X + CellMetrics.X_content_start;
                var X_content_end   = Page.X + CellMetrics.X_content_end;

                // Если в текущей строке на данной странице не убралось ничего из других ячеек, тогда
                // рассчитываем вертикально объединенные ячейки до начала данной строки.
                var Y_content_start = Cell.Temp.Y;
                var Y_content_end   = false === bContentOnFirstPage ? Y : this.Pages[CurPage].YLimit - nFootnotesHeight;

                // TODO: При расчете YLimit для ячейки сделать учет толщины нижних
                //       границ ячейки и таблицы
                if (null != CellSpacing)
                {
                    if (this.Content.length - 1 === CurRow)
                        Y_content_end -= CellSpacing;
                    else
                        Y_content_end -= CellSpacing / 2;
                }

                var VMergeCount = this.Internal_GetVertMergeCount(CurRow, CurGridCol, GridSpan);
                var BottomMargin = this.MaxBotMargin[CurRow + VMergeCount - 1];
                Y_content_end -= BottomMargin;

                if ((0 === Cell.Row.Index && 0 === CurPage) || Cell.Row.Index > FirstRow)
                {
                    Cell.PagesCount = 1;
                    Cell.Content.Set_StartPage(CurPage);
                    Cell.Content.Reset(X_content_start, Y_content_start, X_content_end, Y_content_end);
                }

                // Какие-то ячейки в строке могут быть не разбиты на строки, а какие то разбиты.
                // Здесь контролируем этот момент, чтобы у тех, которые не разбиты не вызывать
                // Recalculate_Page от несуществующих страниц.
                var CellPageIndex = CurPage - Cell.Content.Get_StartPage_Relative();
                if (CellPageIndex < Cell.PagesCount)
                {
                    if (recalcresult2_NextPage === Cell.Content.Recalculate_Page(CellPageIndex, true))
                    {
                        Cell.PagesCount = Cell.Content.Pages.length + 1;
                        bNextPage = true;
                    }

                    var CellContentBounds = Cell.Content.Get_PageBounds(CellPageIndex, undefined, true);
                    var CellContentBounds_Bottom = CellContentBounds.Bottom + BottomMargin;

                    if (0 != CurRow && false === this.RowsInfo[CurRow].FirstPage)
                    {
                        if (this.TableRowsBottom[CurRow - 1][CurPage] < CellContentBounds_Bottom)
                        {
                            // Поскольку мы правим настройку не текущей строки, надо подправить и
                            // запись о рассчитанной высоте строки
                            var Diff = CellContentBounds_Bottom - this.TableRowsBottom[CurRow - 1][CurPage];
                            this.TableRowsBottom[CurRow - 1][CurPage] = CellContentBounds_Bottom;
                            this.RowsInfo[CurRow - 1].H[CurPage] += Diff;
                        }
                    }
                    else
                    {
                        if (undefined === this.TableRowsBottom[CurRow][CurPage] || this.TableRowsBottom[CurRow][CurPage] < CellContentBounds_Bottom)
                            this.TableRowsBottom[CurRow][CurPage] = CellContentBounds_Bottom;
                    }
                }

				// Проверяем наличие сносок, т.к. они могли появится в смерженных ячейках
				var nCurFootnotesHeight = oFootnotes ? oFootnotes.GetHeight(nPageAbs, nColumnAbs) : 0;
				if (oFootnotes && nCurFootnotesHeight > nFootnotesHeight + 0.001 && Cell.Row.Index >= FirstRow)
				{
					this.Pages[CurPage].FootnotesH = nCurFootnotesHeight;

					nFootnotesHeight     = nCurFootnotesHeight;
					nResetFootnotesIndex = CurRow;
					Y                    = arrSavedY[Cell.Row.Index];
					TableHeight          = arrSavedTableHeight[Cell.Row.Index];
					oFootnotes.LoadRecalculateObject(nPageAbs, nColumnAbs, arrFootnotesObject[Cell.Row.Index]);

					CurRow = Cell.Row.Index - 1;

					bFootnoteBreak = true;
					break;
				}

                CurGridCol += GridSpan;
            }

            if (bFootnoteBreak)
				continue;


            // Еще раз обновляем параметр, есть ли текст на первой странице
            bContentOnFirstPage = false;
            for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
            {
                var Cell   = Row.Get_Cell( CurCell );
                var Vmerge = Cell.GetVMerge();

                // Проверяем только начальные ячейки вертикального объединения..
                if ( vmerge_Continue === Vmerge )
                    continue;

                if (true === Cell.IsVerticalText())
                    continue;

                if ( true === Cell.Content_Is_ContentOnFirstPage() )
                {
                    bContentOnFirstPage = true;
                    break;
                }
            }

            this.RowsInfo[CurRow].FirstPage = bContentOnFirstPage;
        }

        // Выставляем так, чтобы высота была равна 0
        if (true !== this.RowsInfo[CurRow].FirstPage && CurPage === this.RowsInfo[CurRow].StartPage)
            this.TableRowsBottom[CurRow][CurPage] = Y;

        // Здесь мы выставляем только начальную координату строки (для каждой страницы)
        // высоту строки(для каждой страницы) мы должны обсчитать после общего цикла, т.к.
        // в одной из следйющих строк может оказаться ячейка с вертикальным объединением,
        // захватывающим данную строку. Значит, ее содержимое может изменить высоту нашей строки.
        var TempY            = Y;
        var TempMaxTopBorder = nMaxTopBorder;

        if ( null != CellSpacing )
        {
            this.RowsInfo[CurRow].Y[CurPage]            = TempY;
            this.RowsInfo[CurRow].TopDy[CurPage]        = 0;
            this.RowsInfo[CurRow].X0                    = Row_x_min;
            this.RowsInfo[CurRow].X1                    = Row_x_max;
            this.RowsInfo[CurRow].MaxTopBorder[CurPage] = TempMaxTopBorder;
            this.RowsInfo[CurRow].MaxBotBorder          = MaxBotBorder[CurRow];
        }
        else
        {
            this.RowsInfo[CurRow].Y[CurPage]            = TempY - TempMaxTopBorder;
            this.RowsInfo[CurRow].TopDy[CurPage]        = TempMaxTopBorder;
            this.RowsInfo[CurRow].X0                    = Row_x_min;
            this.RowsInfo[CurRow].X1                    = Row_x_max;
            this.RowsInfo[CurRow].MaxTopBorder[CurPage] = TempMaxTopBorder;
            this.RowsInfo[CurRow].MaxBotBorder          = MaxBotBorder[CurRow];
        }

        var CellHeight = this.TableRowsBottom[CurRow][CurPage] - Y;

		// TODO: улучшить проверку на высоту строки (для строк разбитых на страницы)
		// Условие Y + RowHValue < Y_content_end добавлено из-за сносок.
		if (false === bNextPage && (Asc.linerule_AtLeast === RowH.HRule || Asc.linerule_Exact === RowH.HRule) && CellHeight < RowHValue && (nFootnotesHeight < 0.001 || Y + RowHValue < Y_content_end))
		{
			CellHeight                            = RowHValue;
			this.TableRowsBottom[CurRow][CurPage] = Y + CellHeight;
		}

        // Рассчитываем ячейки с вертикальным текстом
        var CellsCount2 = VerticallCells.length;
        for (var TempCellIndex = 0; TempCellIndex < CellsCount2; TempCellIndex++)
        {
            var Cell       = VerticallCells[TempCellIndex];
            var CurCell    = Cell.Index;
            var GridSpan   = Cell.Get_GridSpan();
            var CurGridCol = Cell.Metrics.StartGridCol;

            // Возьмем верхнюю ячейку текущего объединения
            Cell = this.Internal_Get_StartMergedCell(CurRow, CurGridCol, GridSpan);

            var CellMar     = Cell.GetMargins();
            var CellMetrics = Cell.Row.Get_CellInfo(Cell.Index);

            var X_content_start = Page.X + CellMetrics.X_content_start;
            var X_content_end   = Page.X + CellMetrics.X_content_end;
            var Y_content_start = Cell.Temp.Y;
            var Y_content_end   = this.TableRowsBottom[CurRow][CurPage];

            // TODO: При расчете YLimit для ячейки сделать учет толщины нижних
            //       границ ячейки и таблицы
            if (null != CellSpacing)
            {
                if (this.Content.length - 1 === CurRow)
                    Y_content_end -= CellSpacing;
                else
                    Y_content_end -= CellSpacing / 2;
            }

            var VMergeCount = this.Internal_GetVertMergeCount(CurRow, CurGridCol, GridSpan);
            var BottomMargin = this.MaxBotMargin[CurRow + VMergeCount - 1];
            Y_content_end -= BottomMargin;

            if ((0 === Cell.Row.Index && true === this.Check_EmptyPages(CurPage - 1)) || Cell.Row.Index > FirstRow || (Cell.Row.Index === FirstRow && true === ResetStartElement))
            {
                // TODO: Здесь надо сделать, чтобы ячейка не билась на страницы
                Cell.PagesCount = 1;
                Cell.Content.Set_StartPage(CurPage);
                Cell.Content.Reset(0, 0, Y_content_end - Y_content_start, 10000);
                Cell.Temp.X_start = X_content_start;
                Cell.Temp.Y_start = Y_content_start;
                Cell.Temp.X_end   = X_content_end;
                Cell.Temp.Y_end   = Y_content_end;

                Cell.Temp.X_cell_start = Page.X + CellMetrics.X_cell_start;
                Cell.Temp.X_cell_end   = Page.X + CellMetrics.X_cell_end;
                Cell.Temp.Y_cell_start = Y_content_start - CellMar.Top.W;
                Cell.Temp.Y_cell_end   = Y_content_end + BottomMargin;
            }

            // Какие-то ячейки в строке могут быть не разбиты на строки, а какие то разбиты.
            // Здесь контролируем этот момент, чтобы у тех, которые не разбиты не вызывать
            // Recalculate_Page от несуществующих страниц.
            var CellPageIndex = CurPage - Cell.Content.Get_StartPage_Relative();
            if (0 === CellPageIndex)
            {
                Cell.Content.Recalculate_Page(CellPageIndex, true);
            }
        }

		// Еще раз проверим, были ли сноски
		var nCurFootnotesHeight = oFootnotes ? oFootnotes.GetHeight(nPageAbs, nColumnAbs) : 0;
		if (oFootnotes && nCurFootnotesHeight > nFootnotesHeight + 0.001)
		{
			this.Pages[CurPage].FootnotesH = nCurFootnotesHeight;

			nFootnotesHeight     = nCurFootnotesHeight;
			nResetFootnotesIndex = CurRow;
			Y                    = arrSavedY[CurRow];
			TableHeight          = arrSavedTableHeight[CurRow];
			oFootnotes.LoadRecalculateObject(nPageAbs, nColumnAbs, arrFootnotesObject[CurRow]);

			CurRow--;
			continue;
		}

        if ( null != CellSpacing )
            this.RowsInfo[CurRow].H[CurPage] = CellHeight;
        else
            this.RowsInfo[CurRow].H[CurPage] = CellHeight + TempMaxTopBorder;

        Y           += CellHeight;
        TableHeight += CellHeight;

        Row.Height   = CellHeight;

        Y           += MaxBotBorder[CurRow];
        TableHeight += MaxBotBorder[CurRow];

        if ( this.Content.length - 1 === CurRow )
        {
            if ( null != CellSpacing )
            {
                TableHeight += CellSpacing;

                var TableBorder_Bottom = this.Get_Borders().Bottom;
                if ( border_Single === TableBorder_Bottom.Value )
                    TableHeight += TableBorder_Bottom.Size;
            }
        }

        if ( true === bNextPage )
        {
            LastRow = CurRow;
            this.Pages[CurPage].LastRow = CurRow;

            if  ( -1 === this.HeaderInfo.PageIndex && this.HeaderInfo.Count > 0 && CurRow >= this.HeaderInfo.Count )
                this.HeaderInfo.PageIndex = CurPage;

            break;
        }
        else if ( this.Content.length - 1 === CurRow )
        {
            LastRow = this.Content.length - 1;
            this.Pages[CurPage].LastRow = this.Content.length - 1;
        }
    }

    var nCompatibilityMode = oLogicDocument && oLogicDocument.GetCompatibilityMode ? oLogicDocument.GetCompatibilityMode() : AscCommon.document_compatibility_mode_Current;
    // Сделаем вертикальное выравнивание ячеек в таблице. Делаем как Word, если ячейка разбилась на несколько
    // страниц, тогда вертикальное выравнивание применяем только к первой странице.
    // Делаем это не в общем цикле, потому что объединенные вертикально ячейки могут вносить поправки в значения
    // this.TableRowsBottom, в последней строке.
    for ( var CurRow = FirstRow; CurRow <= LastRow; CurRow++ )
    {
        var Row = this.Content[CurRow];
        var CellsCount = Row.Get_CellsCount();
        for ( var CurCell = 0; CurCell < CellsCount; CurCell++ )
        {
            var Cell = Row.Get_Cell( CurCell );
            var VMergeCount = this.Internal_GetVertMergeCount( CurRow, Cell.Metrics.StartGridCol, Cell.Get_GridSpan() );

            if ( VMergeCount > 1 && CurRow != LastRow )
                continue;
            else
            {
                var Vmerge = Cell.GetVMerge();
                // Возьмем верхнюю ячейку текущего объединения
                if ( vmerge_Restart != Vmerge )
                {
                    Cell = this.Internal_Get_StartMergedCell( CurRow, Cell.Metrics.StartGridCol, Cell.Get_GridSpan() );
                }
            }

            var CellMar       = Cell.GetMargins();
            var VAlign        = Cell.Get_VAlign();
            var CellPageIndex = CurPage - Cell.Content.Get_StartPage_Relative();

            if ( CellPageIndex >= Cell.PagesCount )
                continue;

            // Рассчитаем имеющуюся в распоряжении высоту ячейки
            var TempCurRow = Cell.Row.Index;

            // Для прилегания к верху или для второй страницы ничего не делаем (так изначально рассчитывалось)
            if ( vertalignjc_Top === VAlign || CellPageIndex > 1 || (1 === CellPageIndex && true === this.RowsInfo[TempCurRow].FirstPage ) )
            {
                Cell.Temp.Y_VAlign_offset[CellPageIndex] = 0;
                continue;
            }

            var TempCellSpacing = this.Content[TempCurRow].Get_CellSpacing();
            var Y_0 = this.RowsInfo[TempCurRow].Y[CurPage];

            if ( null === TempCellSpacing )
                Y_0 += MaxTopBorder[TempCurRow];

            Y_0 += CellMar.Top.W;

            var Y_1 = this.TableRowsBottom[CurRow][CurPage] - CellMar.Bottom.W;
            var CellHeight = Y_1 - Y_0;

            var CellContentBounds = Cell.Content.Get_PageBounds( CellPageIndex, CellHeight, true );
            var ContentHeight = CellContentBounds.Bottom - CellContentBounds.Top;

            var Dy = 0;

            if (true === Cell.IsVerticalText())
            {
                var CellMetrics = Row.Get_CellInfo(CurCell);
                CellHeight = CellMetrics.X_cell_end - CellMetrics.X_cell_start - CellMar.Left.W - CellMar.Right.W;
            }

            if (CellHeight - ContentHeight > 0.001 || nCompatibilityMode <= AscCommon.document_compatibility_mode_Word14)
            {
                if (vertalignjc_Bottom === VAlign)
                    Dy = CellHeight - ContentHeight;
                else if (vertalignjc_Center === VAlign)
                    Dy = (CellHeight - ContentHeight) / 2;

                Cell.ShiftCellContent(CellPageIndex, 0, Dy);
            }

            Cell.Temp.Y_VAlign_offset[CellPageIndex] = Dy;
        }
    }


    // Просчитаем нижнюю границу таблицы на данной странице
    var CurRow = LastRow;
    if ( 0 === CurRow && false === this.RowsInfo[CurRow].FirstPage && 0 === CurPage )
    {
        // Таблица сразу переносится на следующую страницу
        this.Pages[0].MaxBotBorder = 0;
        this.Pages[0].BotBorders   = [];
    }
    else
    {
        // Если последняя строка на данной странице не имеет контента, тогда рассчитываем
        // границу у предыдущей строки.
        if ( false === this.RowsInfo[CurRow].FirstPage && CurPage === this.RowsInfo[CurRow].StartPage )
            CurRow--;

        var MaxBotBorder = 0;
        var BotBorders   = [];

        if (CurRow >= this.Pages[CurPage].FirstRow)
        {
            // Для ряда CurRow вычисляем нижнюю границу
            if (this.Content.length - 1 === CurRow)
            {
                // Для последнего ряда уже есть готовые нижние границы
                var Row        = this.Content[CurRow];
                var CellsCount = Row.Get_CellsCount();
                for (var CurCell = 0; CurCell < CellsCount; CurCell++)
                {
                    var Cell = Row.Get_Cell(CurCell);
                    if (vmerge_Continue === Cell.GetVMerge())
                        Cell = this.Internal_Get_StartMergedCell(CurRow, Row.Get_CellInfo(CurCell).StartGridCol, Cell.Get_GridSpan());

                    var Border_Info = Cell.GetBorderInfo().Bottom;

                    for (var BorderId = 0; BorderId < Border_Info.length; BorderId++)
                    {
                        var Border = Border_Info[BorderId];
                        if (border_Single === Border.Value && MaxBotBorder < Border.Size)
                            MaxBotBorder = Border.Size;

                        BotBorders.push(Border);
                    }
                }
            }
            else
            {
                var Row         = this.Content[CurRow];
                var CellSpacing = Row.Get_CellSpacing();
                var CellsCount  = Row.Get_CellsCount();

                if (null != CellSpacing)
                {
                    // BotBorders можно не заполнять, т.к. у каждой ячейки своя граница,
                    // нам надо только посчитать максимальную толщину.
                    for (var CurCell = 0; CurCell < CellsCount; CurCell++)
                    {
                        var Cell   = Row.Get_Cell(CurCell);
                        var Border = Cell.Get_Borders().Bottom;

                        if (border_Single === Border.Value && MaxBotBorder < Border.Size)
                            MaxBotBorder = Border.Size;
                    }
                }
                else
                {
                    // Сравниваем нижнюю границу ячейки и нижнюю границу таблицы
                    for (var CurCell = 0; CurCell < CellsCount; CurCell++)
                    {
                        var Cell = Row.Get_Cell(CurCell);

                        if (vmerge_Continue === Cell.GetVMerge())
                        {
                            Cell = this.Internal_Get_StartMergedCell(CurRow, Row.Get_CellInfo(CurCell).StartGridCol, Cell.Get_GridSpan());
                            if (null === Cell)
                            {
                                BotBorders.push(TableBorders.Bottom);
                                continue;
                            }
                        }

                        var Border = Cell.Get_Borders().Bottom;

                        // Сравним границы
                        var Result_Border = this.private_ResolveBordersConflict(Border, TableBorders.Bottom, false, true);
                        if (border_Single === Result_Border.Value && MaxBotBorder < Result_Border.Size)
                            MaxBotBorder = Result_Border.Size;

                        BotBorders.push(Result_Border);
                    }
                }
            }
        }

        this.Pages[CurPage].MaxBotBorder = MaxBotBorder;
        this.Pages[CurPage].BotBorders   = BotBorders;
    }

    this.Pages[CurPage].Bounds.Bottom = this.Pages[CurPage].Bounds.Top + TableHeight;
    this.Pages[CurPage].Bounds.Left   = X_min + this.Pages[CurPage].X;
    this.Pages[CurPage].Bounds.Right  = X_max + this.Pages[CurPage].X;
    this.Pages[CurPage].Height        = TableHeight;

    if (true === bNextPage)
    {
        var LastRow = this.Pages[CurPage].LastRow;
        if (false === this.RowsInfo[LastRow].FirstPage)
            this.Pages[CurPage].LastRow = LastRow - 1;
        else
            this.Pages[CurPage].LastRowSplit = true;
    }

    this.TurnOffRecalc = false;

    this.Bounds = this.Pages[this.Pages.length - 1].Bounds;

    if ( true == bNextPage )
        return recalcresult_NextPage;
    else
        return recalcresult_NextElement;
};
CTable.prototype.private_RecalculatePositionY = function(CurPage)
{
	var isHdrFtr = this.Parent.IsHdrFtr();
	var nPageRel = this.GetRelativePage(CurPage);
	var nPageAbs = this.GetAbsolutePage(CurPage);

	var PageLimits = this.Parent.Get_PageLimits(nPageRel);
	var PageFields = this.Parent.Get_PageFields(nPageRel, isHdrFtr);

	var LD_PageFields = this.LogicDocument.Get_PageFields(nPageAbs, isHdrFtr);
	var LD_PageLimits = this.LogicDocument.Get_PageLimits(nPageAbs);

    if ( true === this.Is_Inline() && 0 === CurPage )
    {
        this.AnchorPosition.CalcY = this.Y;
        this.AnchorPosition.Set_Y(this.Pages[CurPage].Height, this.Y, LD_PageFields.Y, LD_PageFields.YLimit, LD_PageLimits.YLimit, PageLimits.Y, PageLimits.YLimit, PageLimits.Y, PageLimits.YLimit);
    }
    else if ( true != this.Is_Inline() && ( 0 === CurPage || ( 1 === CurPage && false === this.RowsInfo[0].FirstPage ) ) )
    {
        this.AnchorPosition.Set_Y(this.Pages[CurPage].Height, this.Pages[CurPage].Y, PageFields.Y, PageFields.YLimit, LD_PageLimits.YLimit, PageLimits.Y, PageLimits.YLimit, PageLimits.Y, PageLimits.YLimit);

        var OtherFlowTables = !this.bPresentation ? editor.WordControl.m_oLogicDocument.DrawingObjects.getAllFloatTablesOnPage( this.Get_StartPage_Absolute() ) : [];
        this.AnchorPosition.Calculate_Y(this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value);
        this.AnchorPosition.Correct_Values( PageLimits.X, PageLimits.Y, PageLimits.XLimit, PageLimits.YLimit, this.AllowOverlap, OtherFlowTables, this );

        if ( undefined != this.PositionV_Old )
        {
            // Восстанови старые значения, чтобы в историю изменений все нормально записалось
            this.PositionV.RelativeFrom = this.PositionV_Old.RelativeFrom;
            this.PositionV.Align        = this.PositionV_Old.Align;
            this.PositionV.Value        = this.PositionV_Old.Value;

            // Рассчитаем сдвиг с учетом старой привязки
            var Value = this.AnchorPosition.Calculate_Y_Value(this.PositionV_Old.RelativeFrom);
            this.Set_PositionV( this.PositionV_Old.RelativeFrom, false, Value );
            // На всякий случай пересчитаем заново координату
            this.AnchorPosition.Calculate_Y(this.PositionV.RelativeFrom, this.PositionV.Align, this.PositionV.Value);

            this.PositionV_Old = undefined;
        }

        var NewX = this.AnchorPosition.CalcX;
        var NewY = this.AnchorPosition.CalcY;

        this.Shift( CurPage, NewX - this.Pages[CurPage].X, NewY - this.Pages[CurPage].Y );
    }
};
CTable.prototype.private_RecalculateSkipPage = function(CurPage)
{
    this.HeaderInfo.Pages[CurPage] = {};
    this.HeaderInfo.Pages[CurPage].Draw = false;

    for ( var Index = -1; Index < this.Content.length; Index++ )
    {
        if (!this.TableRowsBottom[Index])
            this.TableRowsBottom[Index] = [];

        this.TableRowsBottom[Index][CurPage] = 0;
    }

    this.Pages[CurPage].MaxBotBorder = 0;
    this.Pages[CurPage].BotBorders   = [];

    if (0 === CurPage)
    {
        this.Pages[CurPage].FirstRow = 0;
        this.Pages[CurPage].LastRow  = -1;
    }
    else
    {
        var FirstRow;
        if (true === this.IsEmptyPage(CurPage - 1))
            FirstRow = this.Pages[CurPage - 1].FirstRow;
        else
            FirstRow = this.Pages[CurPage - 1].LastRow;

        this.Pages[CurPage].FirstRow = FirstRow;
        this.Pages[CurPage].LastRow  = FirstRow -1;
    }
};
CTable.prototype.private_RecalculatePercentWidth = function()
{
    return this.TableWidthRange - this.GetTableOffsetCorrection() + this.GetRightTableOffsetCorrection();
};
CTable.prototype.private_RecalculateGridCols = function()
{
	for (var nCurRow = 0, nRowsCount = this.Content.length; nCurRow < nRowsCount; ++nCurRow)
	{
		var oRow        = this.Content[nCurRow];
		var oBeforeInfo = oRow.Get_Before();
		var nCurGridCol = oBeforeInfo.GridBefore;

		for (var nCurCell = 0, nCellsCount = oRow.GetCellsCount(); nCurCell < nCellsCount; ++nCurCell)
		{
			var oCell = oRow.GetCell(nCurCell);
			oRow.Set_CellInfo(nCurCell, nCurGridCol, 0, 0, 0, 0, 0, 0);
			oCell.Set_Metrics(nCurGridCol, 0, 0, 0, 0, 0, 0);
			nCurGridCol += oCell.Get_GridSpan();
		}
	}
};
CTable.prototype.private_GetMaxTopBorderWidth = function(nCurRow, isHeader)
{
	var nMax = 0;

	var oRow = this.GetRow(nCurRow);
	for (var nCurCell = 0, nCellsCount = oRow.GetCellsCount(); nCurCell < nCellsCount; ++nCurCell)
	{
		var oCell = oRow.GetCell(nCurCell);
		var arrBorderInfo = isHeader ? oCell.GetBorderInfo().TopHeader : oCell.GetBorderInfo().Top;

		for (var nInfoIndex = 0, nInfosCount = arrBorderInfo.length; nInfoIndex < nInfosCount; ++nInfoIndex)
		{
			var nBorderW = arrBorderInfo[nInfoIndex].GetWidth();
			if (nMax < nBorderW)
				nMax = nBorderW;
		}
	}

	return nMax;
};
//----------------------------------------------------------------------------------------------------------------------
// Класс CTablePage
//----------------------------------------------------------------------------------------------------------------------
function CTablePage(X, Y, XLimit, YLimit, FirstRow, MaxTopBorder)
{
    this.X_origin     = X;
    this.X            = X;
    this.Y            = Y;
    this.XLimit       = XLimit;
    this.YLimit       = YLimit;
    this.Bounds       = new CDocumentBounds(X, Y, XLimit, Y);
    this.MaxTopBorder = MaxTopBorder;
    this.FirstRow     = FirstRow;
    this.LastRow      = FirstRow;
    this.Height       = 0;
    this.LastRowSplit = false;
	this.FootnotesH   = 0;
}
CTablePage.prototype.Shift = function(Dx, Dy)
{
    this.X += Dx;
    this.Y += Dy;
    this.XLimit += Dx;
    this.YLimit += Dy;
    this.Bounds.Shift(Dx, Dy);
};
//----------------------------------------------------------------------------------------------------------------------
// Класс CTableRecalcInfo
//----------------------------------------------------------------------------------------------------------------------
function CTableRecalcInfo()
{
    this.TableGrid     = true;
    this.TableBorders  = true;

    this.CellsToRecalc = {};
    this.CellsAll      = true;
}
CTableRecalcInfo.prototype.RecalcBorders = function()
{
    this.TableBorders = true;
};
CTableRecalcInfo.prototype.Add_Cell = function(Cell)
{
    this.CellsToRecalc[Cell.Get_Id()] = Cell;
};
CTableRecalcInfo.prototype.Check_Cell = function(Cell)
{
    if ( true === this.CellsAll || undefined != this.CellsToRecalc[Cell.Get_Id()] )
        return true;

    return false;
};
CTableRecalcInfo.prototype.Recalc_AllCells = function()
{
    this.CellsAll = true;
};
CTableRecalcInfo.prototype.Reset = function(isNeedRecalculate)
{
    this.TableGrid     = isNeedRecalculate;
    this.TableBorders  = isNeedRecalculate;
    this.CellsAll      = isNeedRecalculate;
    this.CellsToRecalc = {};
};
//----------------------------------------------------------------------------------------------------------------------
// Класс CTableRecalculateObject
//----------------------------------------------------------------------------------------------------------------------
function CTableRecalculateObject()
{
    this.TableSumGrid    = [];
    this.TableGridCalc   = [];

    this.TableRowsBottom = [];
    this.HeaderInfo      = {};
    this.RowsInfo        = [];

    this.X_origin = 0;
    this.X        = 0;
    this.Y        = 0;
    this.XLimit   = 0;
    this.YLimit   = 0;

    this.Pages    = [];

    this.MaxTopBorder = [];
    this.MaxBotBorder = [];
    this.MaxBotMargin = [];

    this.Content = [];
}
CTableRecalculateObject.prototype.Save = function(Table)
{
    this.TableSumGrid    = Table.TableSumGrid;
    this.TableGridCalc   = Table.TableGridCalc;

    this.TableRowsBottom = Table.TableRowsBottom;
    this.HeaderInfo      = Table.HeaderInfo;
    this.RowsInfo        = Table.RowsInfo;

    this.X_origin        = Table.X_origin;
    this.X               = Table.X;
    this.Y               = Table.Y;
    this.XLimit          = Table.XLimit;
    this.YLimit          = Table.YLimit;

    this.Pages           = Table.Pages;

    this.MaxTopBorder    = Table.MaxTopBorder;
    this.MaxBotBorder    = Table.MaxBotBorder;
    this.MaxBotMargin    = Table.MaxBotBorder;

    var Count = Table.Content.length;
    for ( var Index = 0; Index < Count; Index++ )
    {
        this.Content[Index] = Table.Content[Index].SaveRecalculateObject();
    }
};
CTableRecalculateObject.prototype.Load = function(Table)
{
    Table.TableSumGrid    = this.TableSumGrid;
    Table.TableGridCalc   = this.TableGridCalc;

    Table.TableRowsBottom = this.TableRowsBottom;
    Table.HeaderInfo      = this.HeaderInfo;
    Table.RowsInfo        = this.RowsInfo;

    Table.X_origin        = this.X_origin;
    Table.X               = this.X;
    Table.Y               = this.Y;
    Table.XLimit          = this.XLimit;
    Table.YLimit          = this.YLimit;

    Table.Pages           = this.Pages;

    Table.MaxTopBorder    = this.MaxTopBorder;
    Table.MaxBotBorder    = this.MaxBotBorder;
    Table.MaxBotMargin    = this.MaxBotBorder;

    var Count = this.Content.length;
    for ( var Index = 0; Index < Count; Index++ )
    {
        Table.Content[Index].LoadRecalculateObject( this.Content[Index] );
    }
};
CTableRecalculateObject.prototype.Get_DrawingFlowPos = function(FlowPos)
{
    var Count = this.Content.length;
    for (var Index = 0; Index < Count; Index++)
    {
        this.Content[Index].Get_DrawingFlowPos(FlowPos);
    }
};
