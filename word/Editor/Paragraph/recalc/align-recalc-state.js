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

(function(window)
{
	/**
	 * @param {AscWord.Paragraph.WrapRecalcState} wrapState
	 * @constructor
	 */
	function AlignRecalcState(wrapState)
	{
		this.wrapState     = wrapState;
		this.X             = 0; // Текущая позиция по горизонтали
		this.Y             = 0; // Текущая позиция по вертикали
		this.XEnd          = 0; // Предельная позиция по горизонтали
		this.JustifyWord   = 0; // Добавочная ширина символов
		this.JustifySpace  = 0; // Добавочная ширина пробелов
		this.SpacesCounter = 0; // Счетчик пробелов с добавочной шириной (чтобы пробелы в конце строки не трогать)
		this.SpacesSkip    = 0; // Количество пробелов, которые мы пропускаем в начале строки
		this.LettersSkip   = 0; // Количество букв, которые мы пропускаем (из-за таба)
		this.LastW         = 0; // Ширина последнего элемента (необходимо для позиционирования картинки)
		this.Paragraph     = undefined;
		this.RTL           = false;
		this.RecalcResult  = 0x00;//recalcresult_NextElement;
		
		this.Range         = null; // ParaRange
			
		this.LeftSpace     = 0;
		
		this.Y0            = 0; // Верхняя граница строки
		this.Y1            = 0; // Нижняя граница строки

		this.CurPage       = 0;
		this.PageY         = 0;
		this.PageX         = 0;

		this.RecalcFast    = false; // Если пересчет быстрый, тогда все "плавающие" объекты мы не трогаем
		this.RecalcFast2   = false; // Второй вариант быстрого пересчета

		this.ComplexFields = new AscWord.ParagraphComplexFieldStack();
		
		this.bidiFlow = new AscWord.BidiFlow(this);
	}
	AlignRecalcState.prototype.beginPage = function(paragraph, pageNum, isFast)
	{
		this.Paragraph = paragraph;
		this.RTL = paragraph.Get_CompiledPr2(false).ParaPr.Bidi;
		
		this.LastW        = 0;
		this.RecalcFast   = isFast;
		this.RecalcResult = recalcresult_NextElement;
		this.PageY        = paragraph.Pages[pageNum].Bounds.Top;
		this.PageX        = paragraph.Pages[pageNum].Bounds.Left;
		this.CurPage      = pageNum;
	};
	AlignRecalcState.prototype.beginRange = function(range, rangeNum, lineNum, x, counterState, justifyWord, justifySpace)
	{
		this.Range   = range;
		
		this.CurRange = rangeNum;
		this.CurLine  = lineNum;
		
		let paragraph = this.Paragraph;
		
		let lineMetrics = paragraph.Lines[this.CurLine].Metrics;
		let y = paragraph.Pages[this.CurPage].Y + paragraph.Lines[this.CurLine].Y;
		let y0 = y - lineMetrics.Ascent;
		let y1 = y + lineMetrics.Descent;
		if (lineMetrics.LineGap < 0)
			y1 += lineMetrics.LineGap;
		
		this.X = x;
		this.Y = y;
		this.XEnd = range.XEnd;
		
		this.Y0 = y0;
		this.Y1 = y1;
		
		this.LeftSpace = x - range.X;
		
		this.JustifyWord   = justifyWord;
		this.JustifySpace  = justifySpace;
		
		this.SpacesCounter = counterState.Spaces;
		this.SpacesSkip    = counterState.SpacesSkip;
		this.LettersSkip   = counterState.LettersSkip;

		this.RecalcResult  = recalcresult_NextElement;
		
		range.XVisible = x;
		
		this.bidiFlow.begin(this.RTL);
	};
	AlignRecalcState.prototype.endRange = function()
	{
		this.bidiFlow.end();
		
		this.Range.XEndVisible = this.X;
		
		if (this.RTL)
		{
			this.Range.XVisible -= this.Range.WBreak + this.Range.WEnd;
			this.Range.XEndVisible -= this.Range.WBreak + this.Range.WEnd;
		}
	};
	AlignRecalcState.prototype.handleRunElement = function(element, run)
	{
		let type = element.Type;

		let isHiddenCFPart = this.ComplexFields.isHiddenComplexFieldPart();
		if (this.ComplexFields.isHiddenFieldContent() && para_End !== type && para_FieldChar !== type)
		{
			// Чтобы правильно позиционировался курсор и селект
			element.WidthVisible = 0;
			return;
		}

		if (isHiddenCFPart && para_End !== type && para_FieldChar !== type)
		{
			element.WidthVisible = 0;
			return;
		}
	
		if (para_FieldChar === type)
			this.ComplexFields.processFieldChar(element);
		
		this.bidiFlow.add([element, run], element.getBidiType());
	};
	AlignRecalcState.prototype.handleParaMath = function(paraMath)
	{
		if (!paraMath || paraMath.Root.IsEmptyRange(this.CurLine, this.CurRange))
			return;
		
		this.bidiFlow.add([paraMath], paraMath.getBidiType());
	};
	AlignRecalcState.prototype.handleBidiFlow = function(data, direction)
	{
		let element = data[0];
		if (element instanceof AscWord.ParaMath)
			return this.handleBidiFlowParaMath(element);
		
		let run     = data[1];
		let type = element.Type;
		switch (type)
		{
			case para_Sym:
			case para_Text:
			case para_FootnoteReference:
			case para_FootnoteRef:
			case para_EndnoteReference:
			case para_EndnoteRef:
			case para_Separator:
			case para_ContinuationSeparator:
			{
				let WidthVisible = 0;

				if (0 !== this.LettersSkip)
				{
					WidthVisible = element.GetWidth();
					this.LettersSkip--;
				}
				else
					WidthVisible = element.GetWidth() + this.JustifyWord;

				element.SetWidthVisible(WidthVisible, run.Get_CompiledPr(false));

				if (para_FootnoteReference === type || para_EndnoteReference === type)
				{
					var oFootnote = element.GetFootnote();
					oFootnote.UpdatePositionInfo(this.Paragraph, run, this.CurLine, this.CurRange, this.X, WidthVisible);
				}

				this.X += WidthVisible;
				this.LastW = WidthVisible;

				break;
			}
			case para_Math_Text:
			case para_Math_Placeholder:
			case para_Math_BreakOperator:
			case para_Math_Ampersand:
			{
				var WidthVisible = element.GetWidth() / AscWord.TEXTWIDTH_DIVIDER; // GetWidth рассчитываем ширину с учетом состояний Gaps
				element.WidthVisible = (WidthVisible * AscWord.TEXTWIDTH_DIVIDER) | 0;//element.SetWidthVisible(WidthVisible);

				this.X += WidthVisible;
				this.LastW = WidthVisible;

				break;
			}
			case para_Space:
			{
				var WidthVisible = element.GetWidth();

				if (0 !== this.SpacesSkip)
				{
					this.SpacesSkip--;
				}
				else if (0 !== this.SpacesCounter)
				{
					WidthVisible += this.JustifySpace;
					this.SpacesCounter--;
				}

				element.SetWidthVisible(WidthVisible);

				this.X += WidthVisible;
				this.LastW = WidthVisible;

				break;
			}
			case para_Drawing:
			{
				this.handleDrawing(element);
				break;
			}
			case para_PageNum:
			case para_PageCount:
			{
				this.X += element.WidthVisible;
				this.LastW = element.WidthVisible;

				break;
			}
			case para_Tab:
			{
				this.X += element.GetWidthVisible();

				break;
			}
			case para_End:
			{
				element.CheckMark(this.Paragraph, this.RTL ? this.LeftSpace : this.XEnd - this.X);
				let paraEndW = element.GetWidthVisible();
				this.Range.WEnd = paraEndW;
				this.X += paraEndW;

				break;
			}
			case para_NewLine:
			{
				if (element.IsPageBreak() || element.IsColumnBreak())
					element.Update_String(this.RTL ? this.LeftSpace : this.XEnd - this.X);

				let breakW = element.GetWidthVisible();
				this.Range.WBreak = breakW;
				this.X += breakW;

				break;
			}
			case para_FieldChar:
			{
				if (element.IsVisual())
				{
					this.X += element.GetWidthVisible();
					this.LastW = element.GetWidthVisible();
				}

				break;
			}

		}
	};
	AlignRecalcState.prototype.handleBidiFlowParaMath = function(paraMath)
	{
		// до пересчета Bounds для текущей строки ранее должны быть вызваны Recalculate_Range_Width (для ширины), Recalculate_LineMetrics(для высоты и аскента)

		// для инлайновой формулы не вызывается ф-ия setPosition, поэтому необходимо вызвать здесь
		// для неилайновой setPosition вызывается на Get_AlignToLine
		var PosInfo = new CMathPosInfo();

		PosInfo.CurLine  = this.CurLine;
		PosInfo.CurRange = this.CurRange;

		paraMath.Root.setPosition(new CMathPosition(), PosInfo);

		// страиницу для смещния параграфа относительно документа добавим на Get_Bounds, т.к. если формула находится в автофигуре, то для нее не прийдет Recalculate_Range_Spaces при перемещении автофигуры а другую страницу
		paraMath.Root.UpdateBoundsPosInfo(this, this.CurLine, this.CurRange, this.CurPage);
		paraMath.Root.Recalculate_Range_Spaces(this, this.CurLine, this.CurRange, this.CurPage);
	};
	AlignRecalcState.prototype.handleDrawing = function(element)
	{
		let CurPage = this.CurPage;
		let _CurLine = this.CurLine;
		let _CurRange = this.CurRange;
		
		let CurLine = _CurLine;
		
		element.SetForceNoWrap(false);

		var Para = this.Paragraph;
		var isInHdrFtr = Para.Parent.IsHdrFtr();

		var PageAbs = Para.GetAbsolutePage(CurPage);
		var PageRel = Para.GetRelativePage(CurPage);
		var ColumnAbs = Para.GetAbsoluteColumn(CurPage);

		var LogicDocument = this.getLogicDocument();
		var LD_PageLimits = LogicDocument.Get_PageLimits(PageAbs);
		var LD_PageFields = LogicDocument.Get_PageFields(PageAbs, isInHdrFtr);

		var Page_Width = LD_PageLimits.XLimit;
		var Page_Height = LD_PageLimits.YLimit;

		var DrawingObjects = Para.Parent.DrawingObjects;
		var PageLimits = Para.Parent.Get_PageLimits(PageRel);
		var PageFields = Para.Parent.Get_PageFields(PageRel, isInHdrFtr);

		var X_Left_Field = PageFields.X;
		var Y_Top_Field = PageFields.Y;
		var X_Right_Field = PageFields.XLimit;
		var Y_Bottom_Field = PageFields.YLimit;

		var X_Left_Margin = PageFields.X - PageLimits.X;
		var Y_Top_Margin = PageFields.Y - PageLimits.Y;
		var X_Right_Margin = PageLimits.XLimit - PageFields.XLimit;
		var Y_Bottom_Margin = PageLimits.YLimit - PageFields.YLimit;

		var isTableCellContent = Para.IsTableCellContent();
		var isUseWrap = element.Use_TextWrap();
		var isLayoutInCell = element.IsLayoutInCell();

		// TODO: Надо здесь почистить все, а то названия переменных путаются, и некоторые имеют неправильное значение

		if (isTableCellContent && !isLayoutInCell)
		{
			X_Left_Field = LD_PageFields.X;
			Y_Top_Field = LD_PageFields.Y;
			X_Right_Field = LD_PageFields.XLimit;
			Y_Bottom_Field = LD_PageFields.YLimit;

			X_Left_Margin = X_Left_Field;
			X_Right_Margin = Page_Width - X_Right_Field;
			Y_Bottom_Margin = Page_Height - Y_Bottom_Field;
			Y_Top_Margin = Y_Top_Field;
		}

		var _CurPage = 0;
		if (0 !== PageAbs && CurPage > ColumnAbs)
			_CurPage = CurPage - ColumnAbs;

		var ColumnStartX, ColumnEndX;
		if (0 === CurPage)
		{
			// Нужно обновлять, т.к. картинка могла быть внутри данного параграфа и она могла изменить
			// позицию привязки, а при этом в функцию Para.Reset мы не зашли (баг #44739)
			if (Para.Parent.RecalcInfo.Can_RecalcObject() && 0 === CurLine)
				Para.private_RecalculateColumnLimits();

			ColumnStartX = Para.X_ColumnStart;
			ColumnEndX = Para.X_ColumnEnd;
		}
		else
		{
			ColumnStartX = Para.Pages[_CurPage].X;
			ColumnEndX = Para.Pages[_CurPage].XLimit;
		}

		var Top_Margin = Y_Top_Margin;
		var Bottom_Margin = Y_Bottom_Margin;
		var Page_H = Page_Height;

		if (isTableCellContent && isUseWrap)
		{
			Top_Margin = 0;
			Bottom_Margin = 0;
			Page_H = 0;
		}

		var PageLimitsOrigin = Para.Parent.Get_PageLimits(PageRel);
		if (isTableCellContent && !isLayoutInCell)
		{
			PageLimitsOrigin = LogicDocument.Get_PageLimits(PageAbs);
			var PageFieldsOrigin = LogicDocument.Get_PageFields(PageAbs, isInHdrFtr);
			ColumnStartX = PageFieldsOrigin.X;
			ColumnEndX = PageFieldsOrigin.XLimit;
		}

		let isInTable = isTableCellContent && isLayoutInCell;

		if (!isUseWrap)
		{
			PageFields.X = X_Left_Field;
			PageFields.Y = Y_Top_Field;
			PageFields.XLimit = X_Right_Field;
			PageFields.YLimit = Y_Bottom_Field;

			if (!isTableCellContent || !isLayoutInCell)
			{
				PageLimits.X = 0;
				PageLimits.Y = 0;
				PageLimits.XLimit = Page_Width;
				PageLimits.YLimit = Page_Height;
			}
		}

		if (true === element.Is_Inline() || true === Para.Parent.Is_DrawingShape())
		{
			if (linerule_Exact === Para.Get_CompiledPr2(false).ParaPr.Spacing.LineRule)
				element.SetVerticalClip(this.getLineTop(), this.getLineBottom());
			else
				element.SetVerticalClip(null, null);

			element.Update_Position(this.Paragraph, new CParagraphLayout(this.X, this.Y, PageAbs, this.LastW, ColumnStartX, ColumnEndX, X_Left_Margin, X_Right_Margin, Page_Width, Top_Margin, Bottom_Margin, Page_H, PageFields.X, PageFields.Y, Para.Pages[CurPage].Y + Para.Lines[CurLine].Y - Para.Lines[CurLine].Metrics.Ascent, Para.Pages[CurPage].Y), PageLimits, PageLimitsOrigin, _CurLine, isInTable);
			element.Reset_SavedPosition();

			this.X += element.WidthVisible;
			this.LastW = element.WidthVisible;
		}
		else if (!element.IsSkipOnRecalculate())
		{
			Para.Pages[CurPage].Add_Drawing(element);

			if (true === this.RecalcFast)
			{
				// Если у нас быстрый пересчет, тогда мы не трогаем плавающие картинки
				// TODO: Если здесь привязка к символу, тогда быстрый пересчет надо отменить
				return;
			}

			if (true === this.RecalcFast2)
			{
				// Тут мы должны сравнить положение картинок
				var oRecalcObj = element.SaveRecalculateObject();
				element.Update_Position(this.Paragraph, new CParagraphLayout(this.X, this.Y, PageAbs, this.LastW, ColumnStartX, ColumnEndX, X_Left_Margin, X_Right_Margin, Page_Width, Top_Margin, Bottom_Margin, Page_H, PageFields.X, PageFields.Y, Para.Pages[CurPage].Y + Para.Lines[CurLine].Y - Para.Lines[CurLine].Metrics.Ascent, Para.Pages[_CurPage].Y), PageLimits, PageLimitsOrigin, _CurLine, isInTable);

				if (Math.abs(element.X - oRecalcObj.X) > 0.001 || Math.abs(element.Y - oRecalcObj.Y) > 0.001 || element.PageNum !== oRecalcObj.PageNum)
				{
					// Положение картинок не совпало, отправляем пересчет текущей страницы.
					this.RecalcResult = recalcresult_CurPage | recalcresultflags_Page;
					return;
				}

				return;
			}

			// У нас Flow-объект. Если он с обтеканием, тогда мы останавливаем пересчет и
			// запоминаем текущий объект. В функции Internal_Recalculate_2 пересчитываем
			// его позицию и сообщаем ее внешнему классу.

			// Не учитываем обтекание, если у нас на странице больше 100 объектов с обтеканием (баг 73462)
			if (isUseWrap
				&& DrawingObjects && DrawingObjects.graphicPages
				&& DrawingObjects.graphicPages[PageAbs]
				&& DrawingObjects.graphicPages[PageAbs].beforeTextObjects.length >= 100)
			{
				isUseWrap = false;
				let LDRecalcInfo = Para.Parent.RecalcInfo;
				if (LDRecalcInfo.FlowObject)
					LDRecalcInfo.Reset();
			}

			if (isUseWrap)
			{
				var LogicDocument = Para.Parent;
				var LDRecalcInfo = Para.Parent.RecalcInfo;
				if (true === LDRecalcInfo.Can_RecalcObject())
				{
					// Обновляем позицию объекта
					element.Update_Position(this.Paragraph, new CParagraphLayout(this.X, this.Y, PageAbs, this.LastW, ColumnStartX, ColumnEndX, X_Left_Margin, X_Right_Margin, Page_Width, Top_Margin, Bottom_Margin, Page_H, PageFields.X, PageFields.Y, Para.Pages[CurPage].Y + Para.Lines[_CurLine].Y - Para.Lines[_CurLine].Metrics.Ascent, Para.Pages[_CurPage].Y), PageLimits, PageLimitsOrigin, _CurLine, isInTable);

					// For headers we do not check for exceeding lower bound in this case
					if (!isInHdrFtr
						&& this.getCompatibilityMode() >= AscCommon.document_compatibility_mode_Word15
						&& 0 === CurPage
						&& !this.isParagraphStartFromNewPage()
						&& element.Get_Bounds().Bottom >= Y_Bottom_Field
						&& element.IsMoveWithTextVertically())
					{
						// TODO: По хорошему надо пересчитать заново всю текущую страницу с условием, что заданный параграф начинается с новой страницы
						Para.StartFromNewPage();
						this.RecalcResult = recalcresult_NextPage;
					}
					else
					{
						LDRecalcInfo.Set_FlowObject(element, 0, recalcresult_NextElement, -1);

						// TODO: Добавить проверку на не попадание в предыдущие колонки
						if (0 === this.CurPage && element.wrappingPolygon.top > this.PageY + 0.001 && element.wrappingPolygon.left > this.PageX + 0.001)
							this.RecalcResult = recalcresult_CurPagePara;
						else
							this.RecalcResult = recalcresult_CurPage | recalcresultflags_Page;
					}

					return;
				}
				else if (true === LDRecalcInfo.Check_FlowObject(element))
				{
					// Если мы находимся с таблице, тогда делаем как Word, не пересчитываем предыдущую страницу,
					// даже если это необходимо. Такое поведение нужно для точного определения рассчиталась ли
					// данная страница окончательно или нет. Если у нас будет ветка с переходом на предыдущую страницу,
					// тогда не рассчитав следующую страницу мы о конечном рассчете текущей страницы не узнаем.

					// Если данный объект нашли, значит он уже был рассчитан и нам надо проверить номер страницы.
					// Заметим, что даже если картинка привязана к колонке, и после пересчета место привязки картинки
					// сдвигается в следующую колонку, мы проверяем все равно только реальную страницу (без
					// учета колонок, так делает и Word).
					if (element.PageNum === PageAbs)
					{
						if (LDRecalcInfo.IsForceNoWrap())
						{
							element.SetForceNoWrap(true);
							element.Update_Position(this.Paragraph, new CParagraphLayout(this.X, this.Y, PageAbs, this.LastW, ColumnStartX, ColumnEndX, X_Left_Margin, X_Right_Margin, Page_Width, Top_Margin, Bottom_Margin, Page_H, PageFields.X, PageFields.Y, Para.Pages[CurPage].Y + Para.Lines[_CurLine].Y - Para.Lines[_CurLine].Metrics.Ascent, Para.Pages[_CurPage].Y), PageLimits, PageLimitsOrigin, _CurLine, isInTable);
						}

						// Все нормально, можно продолжить пересчет
						LDRecalcInfo.Reset();
						element.Reset_SavedPosition();
					}
					else if (isTableCellContent)
					{
						// Картинка не на нужной странице, но так как это таблица
						// мы пересчитываем заново текущую страницу, а не предыдущую

						// Обновляем позицию объекта
						element.Update_Position(this.Paragraph, new CParagraphLayout(this.X, this.Y, PageAbs, this.LastW, ColumnStartX, ColumnEndX, X_Left_Margin, X_Right_Margin, Page_Width, Top_Margin, Bottom_Margin, Page_H, PageFields.X, PageFields.Y, Para.Pages[CurPage].Y + Para.Lines[CurLine].Y - Para.Lines[CurLine].Metrics.Ascent, Para.Pages[_CurPage].Y), PageLimits, PageLimitsOrigin, _CurLine, isInTable);

						LDRecalcInfo.Set_FlowObject(element, 0, recalcresult_NextElement, -1);
						LDRecalcInfo.Set_PageBreakBefore(false);
						this.RecalcResult = recalcresult_CurPage | recalcresultflags_Page;
						return;
					}
					else
					{
						LDRecalcInfo.Set_PageBreakBefore(true);
						DrawingObjects.removeById(element.PageNum, element.Get_Id());
						this.RecalcResult = recalcresult_PrevPage | recalcresultflags_Page;
						return;
					}
				}
				else
				{
					// Либо данный элемент уже обработан, либо будет обработан в будущем
				}

				return;
			}
			else
			{
				let nParaTop = Para.Pages[_CurPage].Y;
				let nLineTop = Para.Pages[CurPage].Y + Para.Lines[CurLine].Y - Para.Lines[CurLine].Metrics.Ascent;

				// ColumnStartX и ParaTop считаются по тексту, смотри баг #50253
				let compatibilityMode = LogicDocument && LogicDocument.IsDocumentEditor() ? LogicDocument.GetCompatibilityMode() : AscCommon.document_compatibility_mode_Current;
				if (compatibilityMode <= AscCommon.document_compatibility_mode_Word14)
				{
					let nPageStartLine = Para.Pages[_CurPage].StartLine;
					nParaTop = Para.Pages[_CurPage].Y + Para.Lines[nPageStartLine].Top;
					if (Para.Lines[nPageStartLine].Ranges.length > 1)
					{
						var arrTempRanges = Para.Lines[nPageStartLine].Ranges;
						for (var nTempCurRange = 0, nTempRangesCount = arrTempRanges.length; nTempCurRange < nTempRangesCount; ++nTempCurRange)
						{
							if (arrTempRanges[nTempCurRange].W > 0.001 || arrTempRanges[nTempCurRange].WEnd > 0.001)
							{
								ColumnStartX = arrTempRanges[nTempCurRange].X;
								break;
							}
						}
					}
				}

				// Картинка ложится на или под текст, в данном случае пересчет можно спокойно продолжать
				element.Update_Position(this.Paragraph, new CParagraphLayout(this.X, this.Y, PageAbs, this.LastW, ColumnStartX, ColumnEndX, X_Left_Margin, X_Right_Margin, Page_Width, Top_Margin, Bottom_Margin, Page_H, PageFields.X, PageFields.Y, nLineTop, nParaTop), PageLimits, PageLimitsOrigin, _CurLine, isInTable);
				element.Reset_SavedPosition();
			}
		}
	};
	AlignRecalcState.prototype.IsFastRangeRecalc = function()
	{
		return this.RecalcFast;
	};
	AlignRecalcState.prototype.getLogicDocument = function()
	{
		return this.wrapState.Paragraph.GetLogicDocument();
	};
	AlignRecalcState.prototype.getDocumentSettings = function()
	{
		let logicDocument = this.Paragraph.GetLogicDocument();
		if (logicDocument && logicDocument.IsDocumentEditor())
			return logicDocument.getDocumentSettings();

		return AscWord.DEFAULT_DOCUMENT_SETTINGS;
	};
	AlignRecalcState.prototype.getCompatibilityMode = function()
	{
		return this.getDocumentSettings().getCompatibilityMode();
	};
	AlignRecalcState.prototype.getLineTop = function()
	{
		let p = this.Paragraph;
		return p.Pages[this.wrapState.Page].Y + p.Lines[this.wrapState.Line].Top;
	};
	AlignRecalcState.prototype.getLineBottom = function()
	{
		let p = this.Paragraph;
		return p.Pages[this.wrapState.Page].Y + p.Lines[this.wrapState.Line].Bottom;
	};
	AlignRecalcState.prototype.isParagraphStartFromNewPage = function()
	{
		return this.Paragraph.IsStartFromNewPage();
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.Paragraph.AlignRecalcState = AlignRecalcState;

})(window);
