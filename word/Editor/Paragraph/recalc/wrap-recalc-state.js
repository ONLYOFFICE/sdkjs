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
	 * @constructor
	 */
	function WrapRecalcState()
	{
		AscWord.ParagraphRecalculateStateBase.call(this);

	    // Общие параметры, которые заполняются 1 раз на пересчет всей страницы
	    this.Paragraph       = null;
	    this.Parent          = null;
	    this.TopDocument     = null;
	    this.TopIndex        = -1;   // Номер элемента контейнера (содержащего данный параграф), либо номер данного параграфа в самом верхнем документе
	    this.PageAbs         = 0;
	    this.ColumnAbs       = 0;
		this.InTable         = false;
	    this.SectPr          = null; // настройки секции, к которой относится данный параграф
		this.CondensedSpaces = false;
		this.BalanceSBDB     = false; // BalanceSingleByteDoubleByteWidth
		this.autoHyphenation = false;

		this.Fast            = false; // Быстрый ли пересчет

		this.alignState   = new AscWord.Paragraph.AlignRecalcState(this);
		this.counterState = new AscWord.Paragraph.CounterRecalcState(this);

	    //
	    this.Page            = 0;
	    this.Line            = 0;
	    this.Range           = 0;

	    this.Ranges          = [];
	    this.RangesCount     = 0;

		this.LineY = [];

	    this.FirstItemOnLine = true;
		this.PrevItemFirst   = false;
	    this.EmptyLine       = true;
	    this.StartWord       = false;
	    this.Word            = false;
	    this.AddNumbering    = true;
	    this.TextOnLine      = false;
	    this.RangeSpaces     = [];

	    this.BreakPageLine      = false; // Разрыв страницы (параграфа) в данной строке
	    this.UseFirstLine       = false;
	    this.BreakPageLineEmpty = false;
	    this.BreakRealPageLine  = false; // Разрыв страницы документа (не только параграфа) в данной строке
	    this.BadLeftTab         = false; // Левый таб правее правой границы
		this.BreakLine          = false; // Строка закончилась принудительным разрывом
		this.LongWord           = false;

		this.ComplexFields = new AscWord.ParagraphComplexFieldStack();

		this.WordLen         = 0;
	    this.SpaceLen        = 0;
	    this.SpacesCount     = 0;
	    this.LastTab         = new CParagraphRecalculateTabInfo();

	    this.LineTextAscent  = 0;
	    this.LineTextDescent = 0;
	    this.LineTextAscent2 = 0;
	    this.LineAscent      = 0;
	    this.LineDescent     = 0;

	    this.LineTop        = 0;
	    this.LineBottom     = 0;
	    this.LineTop2       = 0;
	    this.LineBottom2    = 0;
	    this.LinePrevBottom = 0;

	    this.XRange = 0; // Начальное положение по горизонтали для данного отрезка
	    this.X      = 0; // Текущее положение по горизонтали
	    this.XEnd   = 0; // Предельное значение по горизонтали для текущего отрезка

	    this.Y      = 0; // Текущее положение по вертикали

	    this.XStart = 0; // Начальное значение для X на данной страницы
	    this.YStart = 0; // Начальное значение для Y на данной страницы
	    this.XLimit = 0; // Предельное значение для X на данной страницы
	    this.YLimit = 0; // Предельное значение для Y на данной страницы

	    this.NewPage  = false; // Переходим на новую страницу
	    this.NewRange = false; // Переходим к новому отрезку
	    this.End      = false;
	    this.RangeY   = false; // Текущая строка переносится по Y из-за обтекания

	    this.CurPos       = new AscWord.CParagraphContentPos();

	    this.NumberingPos = new AscWord.CParagraphContentPos(); // Позиция элемента вместе с которым идет нумерация

		this.MoveToLBP      = false; // Делаем ли разрыв в позиции this.LineBreakPos
		this.UpdateLBP      = true;  // Флаг для первичного обновления позиции переноса в отрезке
		this.LineBreakFirst = true;  // Последняя позиция для переноса - это первый элемент в отрезке

		// Последняя позиция в которой можно будет добавить разрыв отрезка или строки, если что-то не умещается (например,
		// если у нас не убирается слово, то разрыв ставим перед ним)
		this.LineBreakPos   = new AscWord.CParagraphContentPos();

		this.LastItem        = null; // Последний непробельный элемент
		this.LastItemRun     = null; // Run, в котором лежит последний элемент LastItem
		this.LastHyphenItem  = null;
		this.lastAutoHyphen  = null; // Последний элемент с переносом, который убирался в отрезке вместо с дефисом
		this.autoHyphenLimit = 0;
		this.hyphenationZone = 0;

	    this.RunRecalcInfoLast  = null; // RecalcInfo последнего рана
	    this.RunRecalcInfoBreak = null; // RecalcInfo рана, на котором произошел разрыв отрезка/строки

	    this.BaseLineOffset = 0;

	    this.RecalcResult = 0x00;//recalcresult_NextElement;

		// Управляющий объект для пересчета неинлайновой формулы
		this.MathRecalcInfo = {
			Line : 0,    // Номер строки, с которой начинается формула на текущей странице
			Math : null  // Сам объект формулы
		};

	    this.Footnotes                  = [];
		this.FootnotesRecalculateObject = null;

		this.Endnotes = [];

	    // for ParaMath
	    this.bMath_OneLine       = false;
	    this.bMathWordLarge      = false;
	    this.bEndRunToContent    = false;
	    this.PosEndRun           = new AscWord.CParagraphContentPos();

	    // параметры, необходимые для расчета разбиения по операторам
	    // у "крайних" в строке операторов/мат объектов сооответствующий Gap равен нулю
	    this.OperGapRight        = 0;
	    this.OperGapLeft         = 0;
	    this.bPriorityOper       = true;  // есть ли в контенте операторы с высоким приоритетом разбиения
	    this.WrapIndent          = 0;     // WrapIndent нужен для сравнения с длиной слова (когда слово разбивается по Compare Oper): ширина первой строки формулы не должна быть меньше WrapIndent
	    this.bContainCompareOper = true;  // содержаться ли в текущем контенте операторы с высоким приоритетом
	    this.MathFirstItem       = true;  // параметр необходим для принудительного переноса
	    this.bFirstLine          = false;

	    this.bNoOneBreakOperator = true;  // прежде чем обновлять позицию в контент Run, учтем были ли до этого break-операторы (проверки на Word == false не достаточно, т.к. формула мб инлайновая и тогда не нужно обновлять позицию)
	    this.bForcedBreak        = false;
	    this.bInsideOper         = false; // учитываем есть ли разбивка внутри мат объекта, чтобы случайно не вставить в конец пред оператора (при Brk_Before == false)
	    this.bOnlyForcedBreak    = false; // учитывается, если возможна разбивка только по операторам выше уровням => в этом случае можно сделать принудительный разрыв во внутреннем контенте
	    this.bBreakBox           = false;

	    //-----------------------------//
	    this.bFastRecalculate    = false;
	    this.bBreakPosInLWord    = true; // обновляем LineBreakPos (Set_LineBreakPos) для WordLarge. Не обновляем для инлайновой формулы, перед формулой есть еще текст, чтобы не перебить LineBreakPos и выставить по тем меткам, которые были до формулы разбиение
	    this.bContinueRecalc     = false;
	    this.bMathRangeY         = false; // используется для переноса формулы под картинку
	    this.MathNotInline       = null;
	}
	WrapRecalcState.prototype = Object.create(AscWord.ParagraphRecalculateStateBase.prototype);
	WrapRecalcState.prototype.constructor = WrapRecalcState;

	WrapRecalcState.prototype.getAlignState = function()
	{
		return this.alignState;
	};
	WrapRecalcState.prototype.getCounterState = function()
	{
		return this.counterState;
	};
	WrapRecalcState.prototype.Reset_Page = function(Paragraph, CurPage)
	{
		this.Paragraph   = Paragraph;
		this.Parent      = Paragraph.Parent;
		this.TopDocument = Paragraph.Parent.GetTopDocumentContent();
		this.PageAbs     = Paragraph.GetAbsolutePage(CurPage);
		this.ColumnAbs   = Paragraph.GetAbsoluteColumn(CurPage);
		this.InTable     = Paragraph.IsTableCellContent();
		this.SectPr      = null;
		this.TopIndex    = -1;

		this.CondensedSpaces = Paragraph.IsCondensedSpaces();
		this.BalanceSBDB     = Paragraph.IsBalanceSingleByteDoubleByteWidth();

		let settings = this.getDocumentSettings();
		this.autoHyphenation = settings.isAutoHyphenation();
		this.autoHyphenLimit = settings.getConsecutiveHyphenLimit();
		this.hyphenationZone = AscCommon.TwipsToMM(settings.getHyphenationZone());

		if (settings.getCompatibilityMode() >= AscCommon.document_compatibility_mode_Word15)
			this.hyphenationZone = AscCommon.TwipsToMM(AscWord.DEFAULT_HYPHENATION_ZONE);

		this.Page               = CurPage;
		this.RunRecalcInfoLast  = (0 === CurPage ? null : Paragraph.Pages[CurPage - 1].EndInfo.RunRecalcInfo);
		this.RunRecalcInfoBreak = this.RunRecalcInfoLast;

		this.ComplexFields.resetPage(Paragraph, CurPage);
		this.alignState.ComplexFields.resetPage(Paragraph, CurPage);
		this.counterState.ComplexFields.resetPage(Paragraph, CurPage);
	};
	WrapRecalcState.prototype.Reset_Line = function()
	{
		this.RecalcResult = recalcresult_NextLine;

		this.EmptyLine         = true;
		this.BreakPageLine     = false;
		this.BreakLine         = false;
		this.LongWord          = false;
		this.End               = false;
		this.UseFirstLine      = false;
		this.BreakRealPageLine = false;
		this.BadLeftTab        = false
		this.TextOnLine        = false;

		this.LineTextAscent  = 0;
		this.LineTextAscent2 = 0;
		this.LineTextDescent = 0;
		this.LineAscent      = 0;
		this.LineDescent     = 0;

		this.NewPage      = false;
		this.ForceNewPage = false;
		this.ForceNewLine = false;
		this.ForceNewPageAfter = false;

		this.bMath_OneLine    = false;
		this.bMathWordLarge   = false;
		this.bEndRunToContent = false;
		this.PosEndRun        = new AscWord.CParagraphContentPos();
		this.Footnotes        = [];
		this.Endnotes         = [];

		this.OperGapRight        = 0;
		this.OperGapLeft         = 0;
		this.WrapIndent          = 0;
		this.MathFirstItem       = true;
		this.bContainCompareOper = true;
		this.bInsideOper         = false;
		this.bOnlyForcedBreak    = false;
		this.bBreakBox           = false;
		this.bNoOneBreakOperator = true;
		this.bFastRecalculate    = false;
		this.bForcedBreak        = false;
		this.bBreakPosInLWord    = true;

		this.MathNotInline = null;

		this.LineY.length = this.Line + 1;
		if (this.Line >= 0)
			this.LineY[this.Line] = this.Y;
	};
	WrapRecalcState.prototype.resetRange = function(range)
	{
		this.LastTab.Reset();

		this.BreakLine       = false;
		this.SpaceLen        = 0;
		this.WordLen         = 0;
		this.SpacesCount     = 0;
		this.Word            = false;
		this.FirstItemOnLine = true;
		this.StartWord       = false;
		this.NewRange        = false;
		this.X               = range.X;
		this.XEnd            = range.XEnd;
		this.XRange          = range.X;
		this.RangeSpaces     = [];

		this.MoveToLBP      = false;
		this.LineBreakPos   = new AscWord.CParagraphContentPos();
		this.LineBreakFirst = true;
		this.LastItem       = null;
		this.LastItemRun    = null;
		this.UpdateLBP      = true;
		this.LastHyphenItem = null;
		this.lastAutoHyphen = null;

		// for ParaMath
		this.bMath_OneLine    = false;
		this.bMathWordLarge   = false;
		this.bEndRunToContent = false;
		this.PosEndRun        = new AscWord.CParagraphContentPos();

		this.OperGapRight        = 0;
		this.OperGapLeft         = 0;
		this.WrapIndent          = 0;
		this.bContainCompareOper = true;
		this.bInsideOper         = false;
		this.bOnlyForcedBreak    = false;
		this.bBreakBox           = false;
		this.bNoOneBreakOperator = true;
		this.bForcedBreak        = false;
		this.bFastRecalculate    = false;
		this.bBreakPosInLWord    = true;
	};
	WrapRecalcState.prototype.Set_LineBreakPos = function(PosObj, isFirstItemOnLine)
	{
		this.LineBreakPos.Set(this.CurPos);
		this.LineBreakPos.Add(PosObj);
		this.LineBreakFirst = isFirstItemOnLine;
		this.ResetLastAutoHyphen();
	};
	WrapRecalcState.prototype.getDocumentSettings = function()
	{
		let logicDocument = this.Paragraph.GetLogicDocument();
		if (logicDocument && logicDocument.IsDocumentEditor())
			return logicDocument.getDocumentSettings();

		return AscWord.DEFAULT_DOCUMENT_SETTINGS;
	};
	WrapRecalcState.prototype.isDocumentEditor = function()
	{
		let logicDocument = this.Paragraph.GetLogicDocument();
		return (logicDocument && logicDocument.IsDocumentEditor());
	};
	WrapRecalcState.prototype.getCompatibilityMode = function()
	{
		return this.getDocumentSettings().getCompatibilityMode();
	};
	WrapRecalcState.prototype.getXLimit = function()
	{
		// TODO: Когда перенесем весь расчет в данный класс (из Run.Recalculate_Range), то
		//       при изменении XEnd сразу расчитывать это значение и заменить вызов на простой this.XEnd
		return this.Paragraph.IsUseXLimit() ? this.XEnd : MEASUREMENT_MAX_MM_VALUE * 10;
	};
	WrapRecalcState.prototype.ResetLastAutoHyphen = function()
	{
		if (!this.LastHyphenItem)
			return;

		this.LastHyphenItem.SetTemporaryHyphenAfter(false);
		this.LastHyphenItem = null;
	};
	WrapRecalcState.prototype.checkLastAutoHyphen = function()
	{
		if (!this.isAutoHyphenation())
			return;

		this.ResetLastAutoHyphen();
		let lastItem = this.LastItem;
		if (!lastItem || lastItem !== this.lastAutoHyphen)
			return;

		if (this.isExceedConsecutiveAutoHyphenLimit())
			return;

		this.LastHyphenItem = lastItem;
		lastItem.SetTemporaryHyphenAfter(true);
	};
	WrapRecalcState.prototype.Set_NumberingPos = function(PosObj, Item)
	{
		this.NumberingPos.Set(this.CurPos);
		this.NumberingPos.Add(PosObj);

		this.Paragraph.Numbering.Pos  = this.NumberingPos;
		this.Paragraph.Numbering.Item = Item;
	};
	WrapRecalcState.prototype.Update_CurPos = function(PosObj, Depth)
	{
		this.CurPos.Update(PosObj, Depth);
	};
	WrapRecalcState.prototype.Reset_Ranges = function()
	{
		this.Ranges      = [];
		this.RangesCount = 0;
	};
	WrapRecalcState.prototype.Reset_RunRecalcInfo = function()
	{
		this.RunRecalcInfoBreak = this.RunRecalcInfoLast;
	};
	WrapRecalcState.prototype.Reset_MathRecalcInfo = function()
	{
		this.bContinueRecalc = false;
	};
	WrapRecalcState.prototype.Restore_RunRecalcInfo = function()
	{
		this.RunRecalcInfoLast = this.RunRecalcInfoBreak;
	};
	WrapRecalcState.prototype.Recalculate_Numbering = function(Item, Run, ParaPr, _X)
	{
		var CurPage = this.Page, CurLine = this.Line, CurRange = this.Range;
		var Para    = this.Paragraph;
		var X       = _X, LineAscent = this.LineAscent;

		// Если нужно добавить нумерацию и на текущем элементе ее можно добавить, тогда добавляем её
		var NumberingItem = Para.Numbering;
		var NumberingType = Para.Numbering.Type;

		if (para_Numbering === NumberingType)
		{
			var oReviewInfo = this.Paragraph.GetReviewInfo();
			var nReviewType = this.Paragraph.GetReviewType();

			var isHavePrChange = this.Paragraph.HavePrChange();
			var oPrevNumPr     = this.Paragraph.GetPrChangeNumPr();

			var NumPr = ParaPr.NumPr;

			if (!NumPr || !NumPr.IsValid())
				NumPr = undefined;

			if (!oPrevNumPr || !oPrevNumPr.IsValid())
			{
				oPrevNumPr = undefined;
			}
			else
			{
				oPrevNumPr = oPrevNumPr.Copy();
				if (undefined === oPrevNumPr.Lvl)
					oPrevNumPr.Lvl = 0;
			}

			var isHaveNumbering = false;
			if ((undefined === Para.Get_SectionPr() || true !== Para.IsEmpty()) && (NumPr || oPrevNumPr))
			{
				isHaveNumbering = true;
			}

			if (!isHaveNumbering || (!NumPr && !oPrevNumPr) || (!NumPr && reviewtype_Add === nReviewType))
			{
				// Так мы обнуляем все рассчитанные ширины данного элемента
				NumberingItem.Measure(g_oTextMeasurer, undefined);
			}
			else
			{
				var oSavedNumberingValues = this.Paragraph.GetSavedNumberingValues();
				var arrSavedNumInfo       = oSavedNumberingValues ? oSavedNumberingValues.NumInfo : null;
				var arrSavedPrevNumInfo   = oSavedNumberingValues ? oSavedNumberingValues.PrevNumInfo : null;

				var oNumbering = Para.Parent.GetNumbering();

				var oNumLvl  = null;
				let nNumSuff = Asc.c_oAscNumberingSuff.None;

				if (NumPr)
				{
					oNumLvl  = oNumbering.GetNum(NumPr.NumId).GetLvl(NumPr.Lvl);
					nNumSuff = oNumLvl.GetSuff();
				}
				else if (oPrevNumPr)
				{
					// MSWord uses tab instead of suff from PrevNum (74525)
					oNumLvl  = oNumbering.GetNum(oPrevNumPr.NumId).GetLvl(oPrevNumPr.Lvl);
					nNumSuff = Asc.c_oAscNumberingSuff.Tab;
				}

				var oNumTextPr = Para.GetNumberingTextPr();
				var nNumJc     = oNumLvl.GetJc();

				// Здесь измеряется только ширина символов нумерации, без суффикса
				if ((!isHavePrChange && NumPr) || (oPrevNumPr && NumPr && oPrevNumPr.NumId === NumPr.NumId && oPrevNumPr.Lvl === NumPr.Lvl))
				{
					var arrNumInfo = arrSavedNumInfo ? arrSavedNumInfo : Para.Parent.CalculateNumberingValues(Para, NumPr, true);
					var nLvl       = NumPr.Lvl;

					var arrRelatedLvls = oNumLvl.GetRelatedLvlList();
					var isEqual        = true;
					for (var nLvlIndex = 0, nLvlsCount = arrRelatedLvls.length; nLvlIndex < nLvlsCount; ++nLvlIndex)
					{
						var nTempLvl = arrRelatedLvls[nLvlIndex];
						if (arrNumInfo[0][nTempLvl] !== arrNumInfo[1][nTempLvl])
						{
							isEqual = false;
							break;
						}
					}

					if (!isEqual)
					{
						if (reviewtype_Common === nReviewType)
						{
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, arrNumInfo[1], NumPr);
						}
						else
						{
							if (reviewtype_Remove === nReviewType && oReviewInfo.GetPrevAdded())
							{
								NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, undefined, undefined);
							}
							else if (reviewtype_Remove === nReviewType)
							{
								NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo[1], NumPr);
							}
							else
							{
								NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, undefined, undefined);
							}
						}
					}
					else
					{
						if (reviewtype_Remove === nReviewType)
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo[1], NumPr);
						else
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr);
					}
				}
				else if (oPrevNumPr && !NumPr)
				{
					var arrNumInfo2 = arrSavedPrevNumInfo ? arrSavedPrevNumInfo : Para.Parent.CalculateNumberingValues(Para, oPrevNumPr, true);
					NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo2[1], oPrevNumPr);
				}
				else if (isHavePrChange && !oPrevNumPr && NumPr)
				{
					if (reviewtype_Remove === nReviewType)
					{
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, undefined, undefined);
					}
					else
					{
						var arrNumInfo = arrSavedNumInfo ? arrSavedNumInfo : Para.Parent.CalculateNumberingValues(Para, NumPr, true);
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, undefined, undefined);
					}
				}
				else if (oPrevNumPr && NumPr)
				{
					var arrNumInfo  = arrSavedNumInfo ? arrSavedNumInfo : Para.Parent.CalculateNumberingValues(Para, NumPr, true);
					var arrNumInfo2 = arrSavedPrevNumInfo ? arrSavedPrevNumInfo : Para.Parent.CalculateNumberingValues(Para, oPrevNumPr, true);

					var isEqual = false;
					if (arrNumInfo[0][NumPr.Lvl] === arrNumInfo[1][oPrevNumPr.Lvl])
					{
						var oSourceNumLvl = oNumbering.GetNum(oPrevNumPr.NumId).GetLvl(oPrevNumPr.Lvl);
						var oFinalNumLvl  = oNumbering.GetNum(NumPr.NumId).GetLvl(NumPr.Lvl);

						isEqual = oSourceNumLvl.IsSimilar(oFinalNumLvl);
						if (isEqual)
						{
							var arrRelatedLvls = oSourceNumLvl.GetRelatedLvlList();
							for (var nLvlIndex = 0, nLvlsCount = arrRelatedLvls.length; nLvlIndex < nLvlsCount; ++nLvlIndex)
							{
								var nTempLvl = arrRelatedLvls[nLvlIndex];
								if (arrNumInfo[0][nTempLvl] !== arrNumInfo[1][nTempLvl])
								{
									isEqual = false;
									break;
								}
							}
						}
					}

					if (isEqual)
					{
						NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr);
					}
					else
					{
						if (reviewtype_Remove === nReviewType)
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), undefined, undefined, arrNumInfo2[1], oPrevNumPr);
						else if (reviewtype_Add === nReviewType)
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, undefined, undefined);
						else
							NumberingItem.Measure(g_oTextMeasurer, oNumbering, oNumTextPr, Para.Get_Theme(), arrNumInfo[0], NumPr, arrNumInfo2[1], oPrevNumPr);
					}
				}
				else
				{
					// Такого быть не должно
				}

				// При рассчете высоты строки, если у нас параграф со списком, то размер символа
				// в списке влияет только на высоту строки над Baseline, но не влияет на высоту строки
				// ниже baseline.
				if (LineAscent < NumberingItem.Height)
					LineAscent = NumberingItem.Height;

				switch (nNumJc)
				{
					case AscCommon.align_Right:
					{
						NumberingItem.WidthVisible = 0;
						break;
					}
					case AscCommon.align_Center:
					{
						NumberingItem.WidthVisible = NumberingItem.WidthNum / 2;
						break;
					}
					case AscCommon.align_Left:
					default:
					{
						NumberingItem.WidthVisible = NumberingItem.WidthNum;
						break;
					}
				}

				X += NumberingItem.WidthVisible;

				if (oNumLvl.IsLegacy())
				{
					var nLegacySpace  = AscCommon.TwipsToMM(oNumLvl.GetLegacySpace());
					var nLegacyIndent = AscCommon.TwipsToMM(oNumLvl.GetLegacyIndent());
					var nNumWidth     = NumberingItem.WidthNum;

					NumberingItem.WidthSuff = Math.max(nNumWidth, nLegacyIndent, nNumWidth + nLegacySpace) - nNumWidth;
				}
				else
				{
					switch (nNumSuff)
					{
						case Asc.c_oAscNumberingSuff.None:
						{
							// Ничего не делаем
							break;
						}
						case Asc.c_oAscNumberingSuff.Space:
						{
							var OldTextPr = g_oTextMeasurer.GetTextPr();

							var Theme = Para.Get_Theme();
							g_oTextMeasurer.SetTextPr(oNumTextPr, Theme);
							g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII);
							NumberingItem.WidthSuff = g_oTextMeasurer.Measure(" ").Width;
							g_oTextMeasurer.SetTextPr(OldTextPr, Theme);
							break;
						}
						case Asc.c_oAscNumberingSuff.Tab:
						{
							NumberingItem.WidthSuff = Para.private_RecalculateGetTabPos(this, X, ParaPr, CurPage, true).TabWidth;
							break;
						}
					}
				}

				NumberingItem.Width = NumberingItem.WidthNum;
				NumberingItem.WidthVisible += NumberingItem.WidthSuff;

				X += NumberingItem.WidthSuff;
			}
		}
		else if (para_PresentationNumbering === NumberingType)
		{
			var Level  = Para.PresentationPr.Level;
			var Bullet = Para.PresentationPr.Bullet;

			var BulletNum = Para.GetBulletNum();
			if (BulletNum === null)
			{
				BulletNum = 1;
			}
			// Найдем настройки для первого текстового элемента
			var FirstTextPr = Para.Get_FirstTextPr2();

			if (Bullet.IsAlpha())
			{
				if (BulletNum > 780)
				{
					BulletNum = (BulletNum % 780);
				}
			}
			if (BulletNum > 32767)
			{
				BulletNum = (BulletNum % 32767);
			}

			NumberingItem.Bullet    = Bullet;
			NumberingItem.BulletNum = BulletNum;
			NumberingItem.Measure(g_oTextMeasurer, FirstTextPr, Para.Get_Theme(), Para.Get_ColorMap());

			if (!Bullet.IsNone())
			{
				if (ParaPr.Ind.FirstLine < 0)
					NumberingItem.WidthVisible = Math.max(NumberingItem.Width, Para.Pages[CurPage].X + ParaPr.Ind.Left + ParaPr.Ind.FirstLine - X, Para.Pages[CurPage].X + ParaPr.Ind.Left - X);
				else
					NumberingItem.WidthVisible = Math.max(Para.Pages[CurPage].X + ParaPr.Ind.Left + NumberingItem.Width - X, Para.Pages[CurPage].X + ParaPr.Ind.Left + ParaPr.Ind.FirstLine - X, Para.Pages[CurPage].X + ParaPr.Ind.Left - X);
			}

			X += NumberingItem.WidthVisible;
		}

		// Заполним обратные данные в элементе нумерации
		NumberingItem.Item       = Item;
		NumberingItem.Run        = Run;
		NumberingItem.Line       = CurLine;
		NumberingItem.Range      = CurRange;
		NumberingItem.LineAscent = LineAscent;
		NumberingItem.Page       = CurPage;

		return X;
	};
	WrapRecalcState.prototype.IsFast = function()
	{
		return this.Fast;
	};
	WrapRecalcState.prototype.AddFootnoteReference = function(oFootnoteReference, oPos)
	{
		// Ссылки могут добавляться несколько раз, если строка разбита на несколько отрезков
		for (var nIndex = 0, nCount = this.Footnotes.length; nIndex < nCount; ++nIndex)
		{
			if (this.Footnotes[nIndex].FootnoteReference === oFootnoteReference)
				return;
		}

		this.Footnotes.push({FootnoteReference : oFootnoteReference, Pos : oPos});
	};
	WrapRecalcState.prototype.GetFootnoteReferencesCount = function(oFootnoteReference, isAllowCustom)
	{
		var _isAllowCustom = (true === isAllowCustom ? true : false);

		// Если данную ссылку мы добавляли уже в строке, тогда ищем сколько было элементов до нее, если не добавляли,
		// тогда возвращаем просто количество ссылок. Ссылки с флагом CustomMarkFollows не учитываются

		var nRefsCount = 0;
		for (var nIndex = 0, nCount = this.Footnotes.length; nIndex < nCount; ++nIndex)
		{
			if (this.Footnotes[nIndex].FootnoteReference === oFootnoteReference)
				return nRefsCount;

			if (true === _isAllowCustom || true !== this.Footnotes[nIndex].FootnoteReference.IsCustomMarkFollows())
				nRefsCount++;
		}

		return nRefsCount;
	};
	WrapRecalcState.prototype.AddEndnoteReference = function(oEndnoteReference, oPos)
	{
		for (var nIndex = 0, nCount = this.Endnotes.length; nIndex < nCount; ++nIndex)
		{
			if (this.Endnotes[nIndex].EndnoteReference === oEndnoteReference)
				return;
		}

		this.Endnotes.push({EndnoteReference : oEndnoteReference, Pos : oPos});
	};
	WrapRecalcState.prototype.GetEndnoteReferenceNumber = function(oEndnoteReference)
	{
		if (this.Endnotes.length <= 0 || this.Endnotes[0].EndnoteReference === oEndnoteReference)
			return -1;

		var nRefsCount = 0;
		for (var nIndex = 0, nCount = this.Endnotes.length; nIndex < nCount; ++nIndex)
		{
			if (this.Endnotes[nIndex].EndnoteReference === oEndnoteReference)
				return (this.Endnotes[0].EndnoteReference.Number + nRefsCount);

			if (true !== this.Endnotes[nIndex].EndnoteReference.IsCustomMarkFollows())
				nRefsCount++;
		}

		return (this.Endnotes[0].EndnoteReference.Number + nRefsCount);
	};
	WrapRecalcState.prototype.GetEndnoteReferenceCount = function()
	{
		return this.Endnotes.length;
	};
	WrapRecalcState.prototype.SetFast = function(bValue)
	{
		this.Fast = bValue;
	};
	WrapRecalcState.prototype.IsFastRecalculate = function()
	{
		return this.Fast;
	};
	WrapRecalcState.prototype.isFastRecalculation = function()
	{
		return this.Fast;
	};
	WrapRecalcState.prototype.GetPageAbs = function()
	{
		return this.PageAbs;
	};
	WrapRecalcState.prototype.GetColumnAbs = function()
	{
		return this.ColumnAbs;
	};
	WrapRecalcState.prototype.GetCurrentContentPos = function(nPos)
	{
		var oContentPos = this.CurPos.Copy();
		oContentPos.Set(this.CurPos);
		oContentPos.Add(nPos);
		return oContentPos;
	};
	WrapRecalcState.prototype.SaveFootnotesInfo = function()
	{
		var oTopDocument = this.TopDocument;
		if (oTopDocument instanceof CDocument)
			this.FootnotesRecalculateObject = oTopDocument.Footnotes.SaveRecalculateObject(this.PageAbs, this.ColumnAbs);
	};
	WrapRecalcState.prototype.LoadFootnotesInfo = function()
	{
		var oTopDocument = this.TopDocument;
		if (oTopDocument instanceof CDocument && this.FootnotesRecalculateObject)
			oTopDocument.Footnotes.LoadRecalculateObject(this.PageAbs, this.ColumnAbs, this.FootnotesRecalculateObject);
	};
	WrapRecalcState.prototype.IsInTable = function()
	{
		return this.InTable;
	};
	WrapRecalcState.prototype.GetSectPr = function()
	{
		if (null === this.SectPr && this.Paragraph)
			this.SectPr = this.Paragraph.Get_SectPr();

		return this.SectPr;
	};
	WrapRecalcState.prototype.GetTopDocument = function()
	{
		return this.TopDocument;
	};
	WrapRecalcState.prototype.GetTopIndex = function()
	{
		if (-1 === this.TopIndex)
		{
			var arrPos = this.Paragraph.GetDocumentPositionFromObject();
			if (arrPos.length > 0)
				this.TopIndex = arrPos[0].Position;
		}

		return this.TopIndex;
	};
	WrapRecalcState.prototype.ResetMathRecalcInfo = function()
	{
		this.MathRecalcInfo.Line = 0;
		this.MathRecalcInfo.Math = null;
	};
	WrapRecalcState.prototype.SetMathRecalcInfo = function(line, math)
	{
		this.MathRecalcInfo.Line = line;
		this.MathRecalcInfo.Math = math;
	};
	WrapRecalcState.prototype.resetToMathFirstLine = function()
	{
		this.Line = this.MathRecalcInfo.Line;
		this.Y    = this.LineY[this.Line];
		return this.Line;
	};
	WrapRecalcState.prototype.GetMathRecalcInfoObject = function()
	{
		return this.MathRecalcInfo.Math;
	};
	WrapRecalcState.prototype.SetMathRecalcInfoObject = function(oMath)
	{
		this.MathRecalcInfo.Math = oMath;
	};
	WrapRecalcState.prototype.IsCondensedSpaces = function()
	{
		return this.CondensedSpaces;
	};
	WrapRecalcState.prototype.IsBalanceSingleByteDoubleByteWidth = function(oRun, nPos)
	{
		if (this.BalanceSBDB)
		{
			let oParaPos = this.Paragraph.GetPosByElement(oRun);
			if (!oParaPos)
				return true;

			oParaPos.Add(nPos);

			let oRunElements = new CParagraphRunElements(oParaPos, 1, null);
			this.Paragraph.GetPrevRunElements(oRunElements);
			let arrElements = oRunElements.GetElements();
			if (arrElements.length <= 0)
				return true;

			let oItem = arrElements[0];
			if (!oItem || para_Text !== oItem.Type || AscCommon.isEastAsianScript(oItem.Value))
				return true;

			oParaPos.Update(nPos + 1, oParaPos.GetDepth());

			oRunElements = new CParagraphRunElements(oParaPos, 1, null);
			this.Paragraph.GetNextRunElements(oRunElements);
			arrElements = oRunElements.GetElements();
			if (arrElements.length <= 0)
				return true;

			oItem = arrElements[0];
			return (!oItem || para_Text !== oItem.Type || AscCommon.isEastAsianScript(oItem.Value));
		}

		return false;
	};
	WrapRecalcState.prototype.AddCondensedSpaceToRange = function(oSpace)
	{
		this.RangeSpaces.push(oSpace);
		oSpace.ResetCondensedWidth();
	};
	/**
	 * Проверяем убирается ли в заданном отрезке заданная ширина содержимого
	 * @param x {number} - текущая позиция
	 * @param width {number} - ширина проверяемого промежутка
	 * @returns {boolean}
	 */
	WrapRecalcState.prototype.isFitOnLine = function(x, width)
	{
		let xLimit = this.getXLimit();
		if (x + width <= xLimit)
			return true;

		return this.tryCondenseSpaces(x, xLimit, width);
	};
	/**
	 * Пытаемся ужать пробелы по
	 * @param x {number} - текущая позиция
	 * @param xLimit {number} - предельная позиция
	 * @param width {number} - ширина проверяемого промежутка
	 * @returns {boolean}
	 */
	WrapRecalcState.prototype.tryCondenseSpaces = function(x, xLimit, width)
	{
		if (!this.CondensedSpaces)
			return false;

		var nKoef = 1 - 0.25 * (Math.min(12.5, width) / 12.5);

		var nSumSpaces = 0;
		for (var nIndex = 0, nCount = this.RangeSpaces.length; nIndex < nCount; ++nIndex)
		{
			nSumSpaces += this.RangeSpaces[nIndex].WidthOrigin / AscWord.TEXTWIDTH_DIVIDER;
		}

		var nSpace = nSumSpaces * (1 - nKoef);
		if (x - nSpace + width < xLimit)
		{
			for (var nIndex = 0, nCount = this.RangeSpaces.length; nIndex < nCount; ++nIndex)
			{
				this.RangeSpaces[nIndex].SetCondensedWidth(nKoef);
			}

			return true;
		}
		else
		{
			for (var nIndex = 0, nCount = this.RangeSpaces.length; nIndex < nCount; ++nIndex)
			{
				this.RangeSpaces[nIndex].ResetCondensedWidth();
			}
		}

		return false;
	};
	WrapRecalcState.prototype.CheckUpdateLBP = function(nInRunPos)
	{
		 if (this.UpdateLBP)
		 {
			 this.UpdateLBP = false;
			 this.LineBreakPos.Set(this.CurPos);
			 this.LineBreakPos.Add(nInRunPos);
		 }
	};
	WrapRecalcState.prototype.IsNeedShapeFirstWord = function(nCurLine)
	{
		let arrLines = this.Paragraph.Lines;

		return (0 !== nCurLine
			&& arrLines.length > nCurLine
			&& arrLines[nCurLine - 1].Info & paralineinfo_LongWord);
	};
	WrapRecalcState.prototype.IsLastElementInWord = function(oRun, nPos)
	{
		let oItem = oRun.GetElement(nPos)
		if (!oItem)
			return false;

		if (oItem.IsSpaceAfter())
			return true;

		let oParent      = oRun.GetParent();
		let nInParentPos = oRun.GetPosInParent(oParent);
		if (!oParent || -1 === nInParentPos)
			return false;

		let oNextItem  = oRun.GetElement(nPos + 1);
		let nParentLen = oParent.GetElementsCount();
		while (!oNextItem && nInParentPos < nParentLen - 1)
		{
			oRun = oParent.GetElement(++nInParentPos);
			if (!oRun || !(oRun instanceof ParaRun))
				return true;

			oNextItem = oRun.GetElement(0);
		}

		return (!oNextItem || !oNextItem.IsText());
	};
	WrapRecalcState.prototype.isAutoHyphenation = function()
	{
		return this.autoHyphenation;
	};
	WrapRecalcState.prototype.getAutoHyphenLimit = function()
	{
		return this.autoHyphenLimit;
	};
	WrapRecalcState.prototype.getHyphenationZone = function()
	{
		return this.hyphenationZone;
	};
	WrapRecalcState.prototype.onEndRecalculateLineRange = function()
	{
		// Сюда заходим, если закончили пересчиытывать отрезок насильно, а не из-за того, что какой-то элемент не убрался
		// (перенос строк или конец параграфа)
		this.ResetLastAutoHyphen();
	};
	/**
	 * Получам ширину дефиса, если на данном элементе можно разбить слово
	 * @returns {number}
	 */
	WrapRecalcState.prototype.getAutoHyphenWidth = function(item, run)
	{
		if (!this.isAutoHyphenation() || !item || !item.IsText() || !item.isHyphenAfter())
			return 0;

		let textPr = run.Get_CompiledPr(false);
		let fontInfo = textPr.GetFontInfo(AscWord.fontslot_ASCII);
		return AscFonts.GetGraphemeWidth(AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(0x002D, fontInfo.Name, fontInfo.Style)) * textPr.FontSize;
	};
	/**
	 * Проверяем нужно ли сделать обязательный перенос строки после расчета одного диапазона
	 * @returns {boolean}
	 */
	WrapRecalcState.prototype.isForceLineBreak = function()
	{
		return (this.ForceNewPage
			|| this.ForceNewPageAfter
			|| this.NewPage
			|| this.ForceNewLine
			|| this.LastHyphenItem);
	};
	WrapRecalcState.prototype.isExceedConsecutiveAutoHyphenLimit = function()
	{
		let limit = this.getAutoHyphenLimit();
		if (!limit)
			return false;

		let lines   = this.Paragraph.Lines;
		let curLine = this.Line - 1;

		while (curLine >= 0 && lines[curLine].Info & paralineinfo_AutoHyphen)
			--curLine;

		++curLine;

		return this.Line - curLine >= limit;
	};
	WrapRecalcState.prototype.canPlaceAutoHyphenAfter = function(runItem)
	{
		return (this.isAutoHyphenation()
			&& !this.isExceedConsecutiveAutoHyphenLimit()
			&& runItem.isHyphenAfter());
	};
	WrapRecalcState.prototype.checkHyphenationZone = function(x)
	{
		// Делаем как в MSWord (проверено в 2019 версии):
		// отмеряем сколько уже занято на текущей строке от начала строки, добавляем это значение к левому полю документа
		// и вычитаем из позиции правого поля параграфа. Если полученное значение больше hyphenationZone, значит можно
		// делать перенос.
		// Схема немного странная, т.к. мы считаем расстояние от левой границы параграфа, а добавляем его к левому полю,
		// поэтому при смещении параграфа целиком влево или вправо (одинаковом изменении левого и правого отступов)
		// разбиение может происходить по-разному, хотя ширина параграфа не меняется

		let paraPr = this.Paragraph.Get_CompiledPr2(false).ParaPr;

		let shift = paraPr.Ind.Left;
		if (this.UseFirstLine)
			shift += paraPr.Ind.FirstLine;

		return x - shift < this.XLimit - this.getHyphenationZone();
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.Paragraph.WrapRecalcState = WrapRecalcState;

})(window);
